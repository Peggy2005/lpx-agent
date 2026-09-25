#!/usr/bin/env python3
"""
國語辭典查詢工具
用法: python 國語辭典查詢.py
需要: Python 3.8+（套件會自動安裝）
"""

# ── 自動安裝缺少的套件（打包成 .exe 後略過，套件已內建）──────────────
import subprocess, sys, importlib.util

_DEPS = {"flask": "flask", "requests": "requests",
         "bs4": "beautifulsoup4", "openpyxl": "openpyxl",
         "docx": "python-docx", "pypdf": "pypdf",
         "playwright": "playwright", "imageio_ffmpeg": "imageio-ffmpeg"}
_missing = [] if getattr(sys, "frozen", False) else [
    pkg for mod, pkg in _DEPS.items() if importlib.util.find_spec(mod) is None]
if _missing:
    print(f"首次執行，安裝套件：{', '.join(_missing)} ...")
    _base_cmd = [sys.executable, "-m", "pip", "install", "--user", *_missing, "-q"]
    try:
        subprocess.check_call(_base_cmd, stdout=subprocess.DEVNULL)
    except subprocess.CalledProcessError:
        # 較新版 pip 在「externally managed」環境（如部分 macOS/Linux）需要這個旗標；
        # 舊版 pip（如 Windows 內建）不認得此旗標，故先不帶，失敗才補上重試
        _base_cmd.insert(5, "--break-system-packages")
        subprocess.check_call(_base_cmd, stdout=subprocess.DEVNULL)
    print("安裝完成\n")
# ────────────────────────────────────────────────────────────────────

import re, json, time, threading, webbrowser, io, os
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from flask import Flask, request, Response, render_template_string, send_from_directory, send_file
import requests as req
import urllib3
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
import moe_video

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── 設定 ─────────────────────────────────────────────────────────────
PORT = 5000
APP_VERSION = "1.1.0"

BASE = "https://dict.concised.moe.edu.tw"
MOEDICT_URL = "https://www.moedict.tw/a/{word}.json"
# 教育部《國字標準字體筆順學習網》：筆順動畫 CC BY-NC-ND 3.0 TW，只能用官方嵌入碼原樣嵌入、不得改作
STROKE_BASE = "https://stroke-order.learningweb.moe.edu.tw"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept-Language": "zh-TW,zh;q=0.9"}
VOCAB_XLSX = Path(__file__).parent / "vocab_14452.xlsx"
SYNONYMS_JSON = Path(__file__).parent / "synonyms.json"
MASCOT_DIR = Path(__file__).parent / "mascot"
BRANDING_DIR = Path(__file__).parent / "branding"

# 國教院《教材編輯輔助系統》斷詞（語料庫固定遠流語料，關聯詞數量固定 10）
COCT_URL = "https://coct.naer.edu.tw/edit.jsp"
COCT_CHUNK_SIZE = 1500  # 只當「整段完全沒有句號」時的保底切法，正常都用句號切

# 斷詞校正用的辭典查詢併發數（萌典/簡編本都是正常現代系統，開多一點沒問題）
DICT_CHECK_WORKERS = 8

# 斷詞校正第一輪先跳過的字：動貌助詞、數詞、常見量詞。
# 這些字大多是文法功能字或計量單位，跟前後字湊出來的「詞」多半不會是辭典收錄的
# 詞條（如「一碗」「了一」），先跳過讓其他實詞優先配對，配不出新詞了才放回來一起試。
_STOP_CHARS = set(
    "了著過"                    # 動貌助詞
    "的地得"                    # 同音結構助詞（的字短語／狀語／補語標記）
    "一二三四五六七八九十"       # 數詞（小寫）
    "壹貳參肆伍陸柒捌玖拾"       # 數詞（大寫／正式）
    "兩"                        # 數詞（口語，同「二」）
    "個次"                      # 使用者指定的量詞
    "位隻張條件支枝把顆粒塊片本冊篇首幅幢棟間層座棵株朵束串"
    "副對雙打組套輛艘架台部所盒箱包袋瓶罐杯碗盤碟勺匙口頭尾"
    "匹群堆疊排行列圈段節截陣場回遍趟番度聲筆樁宗樣種類款項"
    "名章幕局巡輪批通頓曲折齣卷軸面扇枚尊帖服劑味道餐拳"
    "縷絲滴撮捆綑摞領頂盆缸桶擔籃簍綹"    # 補充量詞
    "也都還又更最很太較再才就便只僅皆均俱悉全共總統獨光純"
    "已曾將常屢頻輒素向本原終始猶尚恰適卻倒竟畢徒枉白殆幸彌尤益愈"
    "不未莫勿毋別"              # 常見單字副詞
)

def _load_vocab_levels():
    mapping = {}
    try:
        wb = load_workbook(str(VOCAB_XLSX), read_only=True, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=2, values_only=True):
            word_cell, deng, ji = row[1], row[2], row[3]
            if word_cell and deng and ji:
                level_str = f"{deng}{ji}"
                for w in str(word_cell).split("/"):
                    w = w.strip()
                    if w:
                        mapping[w] = level_str
        wb.close()
    except Exception:
        pass
    return mapping

VOCAB_LEVEL = _load_vocab_levels()

def _load_synonyms():
    """讀《重編國語辭典修訂本》相似詞索引表（build_synonyms.py 產生）。
    索引表只有「詞目→相似詞」單向，這裡補上反向（A 列 B 為相似詞，查 B 也看得到 A）。"""
    table = {}
    try:
        raw = json.loads(SYNONYMS_JSON.read_text(encoding="utf-8"))
    except Exception:
        return table
    for head, syns in raw.items():
        table.setdefault(head, [])
        for w in syns:
            if w not in table[head]:
                table[head].append(w)
            back = table.setdefault(w, [])
            if head not in back:
                back.append(head)
    return table

SYNONYMS = _load_synonyms()

DE_POS_MAP = [
    (r"\(V\)", "動詞"), (r"\(Adj\)", "形容詞"), (r"\(Adv\)", "副詞"),
    (r"\(S[,\)]", "名詞"), (r"\(N[,\)]", "名詞"), (r"\(Pron\)", "代詞"),
    (r"\(Num\)", "數詞"), (r"\(P\)", "介詞"), (r"\(Conj\)", "連詞"),
]
EN_VERB = re.compile(r"^to\s", re.IGNORECASE)
_POS_ABBR = {
    '形': '形容詞', '動': '動詞', '名': '名詞', '副': '副詞',
    '介': '介詞', '連': '連詞', '代': '代詞', '助': '助詞',
    '嘆': '嘆詞', '量': '量詞', '數': '數詞', '擬': '擬聲詞',
}

def _strip_moedict(s):
    return re.sub(r'`([^~]*)~', r'\1', s or '').strip()

# ── 爬蟲邏輯 ─────────────────────────────────────────────────────────
def make_session():
    s = req.Session()
    s.headers.update(HEADERS)
    s.verify = False
    s.get(f"{BASE}/search.jsp?la=0&powerMode=0", timeout=10)
    return s

_BREAK_CHARS = "。！？；\n.!?;，,、 "

_NON_CHINESE_RE = re.compile(r"[^一-鿿]")  # 詞裡只要含阿拉伯數字/英文字母/其他非中文字元就整詞排除

def _clean_source_text(text):
    """丟進 COCT 斷詞前先清掉考卷/講義類 PDF 常見的非本文雜訊：
    【國文科高一寒假作業】這類標題括號、單獨一行的頁碼裝飾（- 1 -）。
    作者名這種無法通用規則判斷的，交給使用者自己填「排除詞」處理。"""
    text = re.sub(r"【[^】]*】", "", text)
    text = re.sub(r"(?m)^\s*[-－]+\s*\d+\s*[-－]+\s*$", "", text)
    return text

def _split_chunks(text, size=COCT_CHUNK_SIZE):
    """把長文字切成不超過 size 字的區塊。優先在標點/空白處斷開，
    但絕對不允許單一區塊超過 size。只當保底用（見 _split_sentences）：
    整段完全沒有句號時（如整段英文引用/參考文獻）用這個切，避免單次請求過大逾時。"""
    text = text.strip()
    if not text:
        return []
    chunks = []
    i, n = 0, len(text)
    while i < n:
        end = min(i + size, n)
        if end < n:
            cut = max((text.rfind(ch, i, end) for ch in _BREAK_CHARS), default=-1)
            if cut > i:
                end = cut + 1
        piece = text[i:end].strip()
        if piece:
            chunks.append(piece)
        i = end
    return chunks

def _split_sentences(text):
    """一句一句丟進 COCT 斷詞：以句號「。」為斷點，一句（含句號）就是一個區塊，
    不再用固定字數切。某段落中間完全沒有句號（如整段英文引用）才退回
    _split_chunks 的固定字數切法保底，避免單一區塊過大逾時。"""
    text = text.strip()
    if not text:
        return []
    sentences, buf = [], ""
    for ch in text:
        buf += ch
        if ch == "。":
            sentences.append(buf)
            buf = ""
    if buf.strip():
        sentences.append(buf)
    chunks = []
    for s in sentences:
        s = s.strip()
        if not s:
            continue
        if len(s) > COCT_CHUNK_SIZE:
            chunks.extend(_split_chunks(s))
        else:
            chunks.append(s)
    return chunks

def _segment_chunk(text, session, retries=2):
    """呼叫 COCT，回傳這個區塊依序斷出的詞（已濾掉標點）。單一區塊失敗會重試，
    重試仍失敗就放棄這一小段、繼續處理其他區塊，不讓整篇長文因一段網路異常而全滅。"""
    data = {"no": "- 存檔 -", "lv": "0", "db": "1", "num": "10",
            "size": "M", "WORD": text, "tp": "0"}
    last_err = None
    for attempt in range(retries + 1):
        try:
            r = session.post(COCT_URL, data=data,
                             headers={"Content-Type": "application/x-www-form-urlencoded"},
                             timeout=45, verify=False)
            r.raise_for_status()
            r.encoding = "utf-8"
            soup = BeautifulSoup(r.text, "html.parser")
            out1 = soup.find(id="out1")
            if not out1:
                return []
            words = []
            for idiv in out1.find_all("idiv"):
                code = idiv.find("code")
                pos = code.get_text(strip=True) if code else ""
                if not pos:
                    continue  # 沒有詞性標記 = 標點符號，跳過
                word = (idiv.get("title") or idiv.get_text(strip=True)).strip()
                if word:
                    words.append(word)
            return words
        except Exception as e:
            last_err = e
            time.sleep(1)
    print(f"[segment] 區塊斷詞失敗，跳過（{len(text)} 字）：{last_err}")
    return []

def _align_tokens(text, words):
    """把 COCT 斷出的詞（已濾掉標點）依序對回原文，詞與詞之間的空隙
    （標點、空白、COCT 漏掉的字）逐字補成 token，確保 token 串起來 == 原文。
    某個詞在原文找不到（COCT 偶爾正規化字形）就略過，它的字會落在下一個空隙裡。"""
    tokens, cur = [], 0
    for w in words:
        idx = text.find(w, cur)
        if idx < 0:
            continue
        tokens.extend(text[cur:idx])
        tokens.append(w)
        cur = idx + len(w)
    tokens.extend(text[cur:])
    return tokens

def synonyms_of(word):
    """查相似詞索引表，查不到再試「台→臺」異體字。每個相似詞附上詞彙等級。"""
    syns = SYNONYMS.get(word) or SYNONYMS.get(word.translate(_VARIANT_TABLE)) or []
    return [{"word": w, "level": VOCAB_LEVEL.get(w, "—")} for w in syns]

def segment_text(text, session):
    """把一整段文字送去 COCT 斷詞，回傳去重複、去標點、經過斷詞校正的詞語清單。
    校正邏輯見 _resolve_words_stream。
    （網頁版走 /segment 路由自己迭代 _split_sentences 以回報進度，這支給其他呼叫端用。）"""
    all_words = []
    for chunk in _split_sentences(text):
        all_words.extend(_segment_chunk(chunk, session))
    words = list(dict.fromkeys(all_words))
    for event in _resolve_words_stream(words, session):
        if "__final__" in event:
            return event["__final__"]
    return words

def _char_has_definition(word, session):
    """查 moedict 這個字本身有沒有實際定義。比完整的 lookup() 輕量（只打
    moedict 一次，不 fallback 查 dict.concised），單純用來判斷「這個字
    是不是真的有意思」，不是要在這裡就把最終釋義生出來。"""
    try:
        r = session.get(MOEDICT_URL.format(word=word), timeout=8)
        if r.status_code != 200:
            return False
        data = r.json()
        return bool(data.get("h"))
    except Exception:
        return False

_VARIANT_TABLE = str.maketrans({"台": "臺"})  # 辭典正式詞條多用「臺」，但「台灣」「台北」等一般慣用「台」

def _has_dict_entry(word, session):
    """判斷 word 在萌典或教育部簡編本任一查得到詞條，用來決定斷詞校正時
    一個多字詞/候選詞是否成立——只看「辭典收不收」，不管語料庫怎麼標。
    查不到再試「台→臺」異體字，否則「台灣」「台北」這種常見詞會被誤判成
    辭典查無、進而在後面的配對邏輯裡被拆回單字。"""
    if get_moedict_pos_groups(word, session):
        return True
    entries = lookup_concised(word, session)
    if bool(entries) and entries[0].get("pinyin") != "錯誤":
        return True
    variant = word.translate(_VARIANT_TABLE)
    if variant == word:
        return False
    if get_moedict_pos_groups(variant, session):
        return True
    entries = lookup_concised(variant, session)
    return bool(entries) and entries[0].get("pinyin") != "錯誤"

def _resolve_words_stream(words, session, ordered=False):
    """斷詞校正（generator）：
    ①掃過所有 2 字以上的詞，辭典（萌典＋簡編本，任一查到就算）查得到就保留，
      查不到判定是斷詞切錯，拆成單字放回原本位置。
    ②剩下的單字反覆做前後字配對：第一輪先跳過數詞/量詞/動貌助詞等停用字
      （_STOP_CHARS），由左到右貪婪認領查得到的候選詞，避免同一個字被兩個
      詞重複用掉；配不出新詞了才把停用字放回來一起再試一輪；直到某輪完全
      配不出新詞為止，剩下的單字維持單字（交由呼叫端逐字查辭典）。
    每次批次查詢後 yield 一個進度 dict；結束時 yield {"__final__": [...]}。
    ordered=True（近義詞替換用）：words 是一句話依序的 token（含標點），
    配對成功的詞留在原位置、不去重複、不丟掉查無定義的單字，
    __final__ 串起來就是原句。非中文 token（標點等）不參與配對，當隔板。"""
    multi = [w for w in dict.fromkeys(words) if len(w) > 1 and not _NON_CHINESE_RE.search(w)]
    hit = set()
    if multi:
        total = len(multi)
        yield {"stage": "word_check", "done": 0, "total": total}
        with ThreadPoolExecutor(max_workers=DICT_CHECK_WORKERS) as ex:
            futures = {ex.submit(_has_dict_entry, w, session): w for w in multi}
            done_n = 0
            for fut in as_completed(futures):
                w = futures[fut]
                try:
                    if fut.result():
                        hit.add(w)
                except Exception:
                    pass
                done_n += 1
                yield {"stage": "word_check", "done": done_n, "total": total}

    expanded = []
    for w in words:
        if len(w) > 1 and w not in hit and not _NON_CHINESE_RE.search(w):
            expanded.extend(list(w))
        else:
            expanded.append(w)

    # pool 保留完整序列（含已保留的多字詞）——多字詞是無法跨越的隔板，
    # 兩側原本不相鄰的單字不能因為中間的多字詞而被誤判成相鄰去配對。
    # 只有單字才會被拿掉／替換；多字詞的位置永遠不變。
    pool = list(expanded)

    def _segments(stoplist_parked):
        """回傳這一輪可互相配對的連續單字片段（各為 pool 的 index 清單）。
        多字詞切斷片段；停用字（僅第一輪）只是跳過不參與配對，不切斷片段——
        片段內兩側的字還是可以隔著被跳過的停用字互相配對。"""
        segs, cur = [], []
        for i, tok in enumerate(pool):
            if len(tok) > 1 or _NON_CHINESE_RE.search(tok):
                if cur:
                    segs.append(cur)
                cur = []
                continue
            if stoplist_parked and tok in _STOP_CHARS:
                if ordered and cur:
                    # 依序模式配成的詞要放回原句，兩字必須真的相鄰，停用字改成切斷片段
                    segs.append(cur)
                    cur = []
                continue
            cur.append(i)
        if cur:
            segs.append(cur)
        return segs

    merged_words = []
    stoplist_parked = True
    round_no = 0
    while round_no < 20:
        segments = _segments(stoplist_parked)
        pair_slots = [(seg[k], seg[k + 1]) for seg in segments for k in range(len(seg) - 1)]
        if not pair_slots:
            if stoplist_parked:
                stoplist_parked = False
                continue
            break

        round_no += 1
        candidates = sorted({pool[a] + pool[b] for a, b in pair_slots})
        total = len(candidates)
        yield {"stage": "pair_check", "round": round_no, "done": 0, "total": total}
        hit_pairs = set()
        with ThreadPoolExecutor(max_workers=DICT_CHECK_WORKERS) as ex:
            futures = {ex.submit(_has_dict_entry, c, session): c for c in candidates}
            done_n = 0
            for fut in as_completed(futures):
                c = futures[fut]
                try:
                    if fut.result():
                        hit_pairs.add(c)
                except Exception:
                    pass
                done_n += 1
                yield {"stage": "pair_check", "round": round_no, "done": done_n, "total": total}

        # 由左到右貪婪認領：同一片段內，候選詞查得到、且左右兩字都還沒被
        # 這輪其他詞用掉才算數；片段之間互不影響。
        claimed = set()
        claimed_pairs = []
        found_any = False
        for seg in segments:
            k = 0
            while k < len(seg) - 1:
                a, b = seg[k], seg[k + 1]
                pair = pool[a] + pool[b]
                if pair in hit_pairs and a not in claimed and b not in claimed:
                    merged_words.append(pair)
                    claimed.add(a); claimed.add(b)
                    claimed_pairs.append((a, b))
                    found_any = True
                    k += 2
                else:
                    k += 1

        if not found_any:
            if stoplist_parked:
                stoplist_parked = False
                continue
            break

        if ordered:
            # 配成的詞放回左字的位置、拿掉右字；變成多字詞後自然成為下一輪的隔板
            merged_at = {a: pool[a] + pool[b] for a, b in claimed_pairs}
            pool = [merged_at.get(i, tok) for i, tok in enumerate(pool)
                    if i in merged_at or i not in claimed]
        else:
            pool = [tok for i, tok in enumerate(pool) if i not in claimed]

    if ordered:
        yield {"__final__": pool}
        return

    others = [w for w in pool if len(w) != 1]
    pool = [w for w in pool if len(w) == 1]

    # 完全配不出新詞的單字：不代表沒意義（像「超」「給」這種字本身在字典裡
    # 就查得到定義，只是前後湊不出辭典收錄的詞），查 moedict 有沒有實際定義，
    # 有就保留，真的完全查無資料（罕見，通常是切錯的雜訊殘片）才整個丟掉。
    if pool:
        yield {"stage": "def_check", "done": 0, "total": len(pool)}
        has_def = set()
        with ThreadPoolExecutor(max_workers=DICT_CHECK_WORKERS) as ex:
            futures = {ex.submit(_char_has_definition, w, session): w for w in pool}
            done_n = 0
            for fut in as_completed(futures):
                w = futures[fut]
                try:
                    if fut.result():
                        has_def.add(w)
                except Exception:
                    pass
                done_n += 1
                yield {"stage": "def_check", "done": done_n, "total": len(pool)}
        pool = [w for w in pool if w in has_def]

    yield {"__final__": list(dict.fromkeys(others + merged_words + pool))}

def extract_text_from_file(filename, raw):
    """依副檔名從上傳的檔案內容抽出純文字。支援 .txt(.md/.csv) / .docx / .pdf。"""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext in ("txt", "md", "csv"):
        for enc in ("utf-8", "big5", "gbk"):
            try:
                return raw.decode(enc)
            except UnicodeDecodeError:
                continue
        return raw.decode("utf-8", errors="ignore")
    if ext == "docx":
        import io, docx
        return "\n".join(p.text for p in docx.Document(io.BytesIO(raw)).paragraphs)
    if ext == "pdf":
        import io
        from pypdf import PdfReader
        text = "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(raw)).pages)
        # 有些 PDF 的字型把「手、大、一」等字對應成外觀相同的康熙部首／部首補充字元
        # （U+2E80–2FDF），不轉回來斷詞、查辭典都會查不到；只轉這一段，全形標點不動
        import unicodedata
        return "".join(unicodedata.normalize("NFKC", c) if "\u2e80" <= c <= "\u2fdf" else c for c in text)
    hint = "；舊版 Word 的 .doc 請用 Word「另存新檔」成 .docx" if ext == "doc" else ""
    raise ValueError(f"不支援的檔案格式：.{ext}（僅支援 txt / docx / pdf{hint}）")

def get_pos(word, session):
    try:
        r = session.get(MOEDICT_URL.format(word=word), timeout=8)
        if r.status_code != 200:
            return "—"
        data = r.json()
        de = data.get("Deutsch", "")
        for pat, label in DE_POS_MAP:
            if re.search(pat, de):
                return label
        if EN_VERB.match(data.get("English", "")):
            return "動詞"
    except Exception:
        pass
    return "—"

def get_moedict_pos_groups(word, session):
    """
    Returns list of {pinyin, pos, definition} grouped by POS per pronunciation.
    Uses moedict.tw `type` field. Returns None when no type info available.
    """
    try:
        r = session.get(MOEDICT_URL.format(word=word), timeout=8)
        if r.status_code != 200:
            return None
        data = r.json()
        results = []
        fallback_pos = None
        for h in data.get('h', []):
            pinyin = h.get('p', '').replace(' ', '')
            defs = h.get('d', [])
            has_type = any(d.get('type') for d in defs)
            groups, order = {}, []
            for d in defs:
                abbr = _strip_moedict(d.get('type') or '')
                pos = _POS_ABBR.get(abbr, abbr or '—')
                f = _strip_moedict(d.get('f', ''))
                if not f:
                    continue
                if pos not in groups:
                    groups[pos] = []
                    order.append(pos)
                groups[pos].append(f)
            if not has_type and groups:
                # 這個讀音完全沒有詞性標記，改用德文欄位猜詞性，釋義合併成一組
                if fallback_pos is None:
                    fallback_pos = get_pos(word, session)
                merged = [d for lst in groups.values() for d in lst]
                groups, order = {fallback_pos: merged}, [fallback_pos]
            for pos in order:
                dl = groups[pos]
                def_str = dl[0] if len(dl) == 1 else \
                          '\n'.join(f'{i+1}. {d}' for i, d in enumerate(dl))
                results.append({'pinyin': pinyin, 'pos': pos, 'definition': def_str})
        return results or None
    except Exception:
        return None

def _section_text(h4_tag):
    parts = []
    for sib in h4_tag.next_siblings:
        if sib.name == "h4":
            break
        if hasattr(sib, "get_text"):
            t = sib.get_text(" ", strip=True)
            if t:
                parts.append(t)
        elif isinstance(sib, str) and sib.strip():
            parts.append(sib.strip())
    return " ".join(parts).strip()

def _li_def(li):
    """Extract definition text from <li>, skipping <idiv> example blocks."""
    from bs4 import NavigableString, Tag
    parts = []
    for child in li.children:
        if isinstance(child, Tag) and child.name == "idiv":
            continue
        elif isinstance(child, Tag):
            t = child.get_text("", strip=True)
            if t: parts.append(t)
        elif isinstance(child, NavigableString):
            t = str(child).strip()
            if t: parts.append(t)
    return "".join(parts).strip()

def _extract_defs(h4_tag):
    """Return list of definition strings from the <ol>/<li> structure under 釋義."""
    from bs4 import Tag
    for sib in h4_tag.next_siblings:
        if isinstance(sib, Tag) and sib.name == "h4":
            break
        if isinstance(sib, Tag) and sib.name == "div":
            ol = sib.find("ol")
            if ol:
                defs = [_li_def(li) for li in ol.find_all("li", recursive=False)]
                defs = [d for d in defs if d]
                if defs:
                    return defs
            # No <ol> — inline definition (each word in <a> tags), skip <idiv> examples
            from bs4 import NavigableString, Comment
            parts = []
            for child in sib.children:
                if isinstance(child, Comment):
                    continue
                if isinstance(child, Tag) and child.name == "idiv":
                    continue
                elif isinstance(child, Tag):
                    t = child.get_text("", strip=True)
                    if t: parts.append(t)
                elif isinstance(child, NavigableString):
                    t = str(child).strip()
                    if t: parts.append(t)
            text = "".join(parts).strip()
            if text:
                return [text]
    return []

def parse_entry_page(soup):
    entries = []
    cur_pinyin = cur_def = ""
    for h4 in soup.find_all("h4"):
        label = h4.get_text(strip=True).replace("　", "").replace(" ", "")
        if label == "漢語拼音":
            if cur_pinyin or cur_def:
                entries.append({"pinyin": cur_pinyin, "definition": cur_def})
            cur_pinyin, cur_def = _section_text(h4).replace(" ", ""), ""
        elif label == "釋義":
            defs = _extract_defs(h4)
            if len(defs) > 1:
                cur_def = "\n".join(f"{i+1}. {d}" for i, d in enumerate(defs))
            elif defs:
                cur_def = defs[0]
    if cur_pinyin or cur_def:
        entries.append({"pinyin": cur_pinyin, "definition": cur_def})
    return entries

def lookup_concised(word, session):
    try:
        r = session.get(f"{BASE}/search.jsp",
                        params={"la": "0", "powerMode": "0", "word": word},
                        timeout=10)
        r.encoding = "utf-8"
        soup = BeautifulSoup(r.text, "html.parser")
        entries = parse_entry_page(soup)
        if entries:
            return entries
        # Disambiguation / search-results page — follow ALL matching links
        # so multi-pronunciation characters (多音字) are fully captured.
        all_entries, seen = [], set()
        for a in soup.find_all("a", href=True):
            href = a["href"]
            if a.get_text(strip=True) == word and "dictView.jsp" in href and href not in seen:
                seen.add(href)
                r2 = session.get(f"{BASE}/{href.lstrip('/')}",
                                 headers={"Referer": r.url}, timeout=10)
                r2.encoding = "utf-8"
                all_entries.extend(parse_entry_page(BeautifulSoup(r2.text, "html.parser")))
        return all_entries
    except Exception as e:
        return [{"pinyin": "錯誤", "definition": str(e)}]

def lookup(word, session):
    level = VOCAB_LEVEL.get(word, "—")
    # Try moedict.tw per-POS grouping first
    pos_groups = get_moedict_pos_groups(word, session)
    if pos_groups:
        return [{"word": word, "pinyin": e["pinyin"] or "—",
                 "pos": e["pos"], "definition": e["definition"], "level": level}
                for e in pos_groups]
    # Fallback: simplified dict + single POS from German translation
    entries = lookup_concised(word, session)
    pos = get_pos(word, session)
    if entries:
        results = []
        for e in entries:
            if e["pinyin"] == "錯誤":
                results.append({"word": word, "pinyin": "錯誤", "pos": "—", "definition": e["definition"], "level": level})
            else:
                results.append({"word": word, "pinyin": e["pinyin"] or "—",
                                "pos": pos, "definition": e["definition"] or "—", "level": level})
        return results
    # 兩個辭典都查無完整詞條：先試「台→臺」異體字（辭典正式詞條多用臺，
    # 「台灣」「台北」這種詞查不到，「臺灣」「臺北」才查得到），顯示仍用
    # 原字形——這一步要放在「只要有詞性/等級就不判查無資料」的保底之前，
    # 否則「台灣」這種本身就在等級表裡的詞，會在還沒試過異體字前就被那條
    # 保底規則接走，只給「—」佔位而查不到真正的釋義。
    variant = word.translate(_VARIANT_TABLE)
    if variant != word:
        variant_pos_groups = get_moedict_pos_groups(variant, session)
        if variant_pos_groups:
            return [{"word": word, "pinyin": e["pinyin"] or "—",
                     "pos": e["pos"], "definition": e["definition"], "level": level}
                    for e in variant_pos_groups]
        variant_entries = lookup_concised(variant, session)
        if variant_entries:
            variant_pos = get_pos(variant, session)
            results = []
            for e in variant_entries:
                if e["pinyin"] == "錯誤":
                    results.append({"word": word, "pinyin": "錯誤", "pos": "—", "definition": e["definition"], "level": level})
                else:
                    results.append({"word": word, "pinyin": e["pinyin"] or "—",
                                    "pos": variant_pos, "definition": e["definition"] or "—", "level": level})
            return results
    # 異體字也查無完整詞條：只要還撈得到詞性或等級，就把能給的給出去，
    # 不要整列判「查無資料」
    if pos != "—" or level != "—":
        return [{"word": word, "pinyin": "—", "pos": pos, "definition": "—", "level": level}]
    # 整詞兩個辭典都查無、詞性也抓不到：這種通常是「簽好」「的話」「拿給」
    # 這類斷詞沒切錯、但辭典本身不收錄的文法組合詞（動詞＋補語、語助詞短語等），
    # 拆成單字個別查，比整詞掛「查無資料」有用。
    if len(word) > 1:
        results = []
        for ch in word:
            results.extend(lookup(ch, session))
        return results
    return [{"word": word, "pinyin": "查無資料", "pos": "—", "definition": "—", "level": level}]

def build_excel(results, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Results"
    hf = PatternFill("solid", fgColor="4472C4")
    hfont = Font(bold=True, color="FFFFFF", size=12)
    for col, h in enumerate(["漢字詞彙", "音標", "詞類", "意思", "詞彙等級"], 1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = hfont; c.fill = hf
        c.alignment = Alignment(horizontal="center", vertical="center")
    for ri, entry in enumerate(results, 2):
        rf = PatternFill("solid", fgColor="D9E2F3" if ri % 2 == 0 else "FFFFFF")
        for ci, val in enumerate([entry["word"], entry["pinyin"],
                                   entry["pos"], entry["definition"],
                                   entry.get("level", "—")], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill = rf
            c.alignment = Alignment(
                horizontal="center" if ci != 4 else "left",
                vertical="center", wrap_text=True)
    for col, w in enumerate([12, 16, 10, 60, 12], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.row_dimensions[1].height = 22
    for row in range(2, len(results) + 2):
        ws.row_dimensions[row].height = 30
    wb.save(output_path)

# ── Flask 應用────────────────────────────────────────────────────────
app = Flask(__name__)
_session = None
_session_lock = threading.Lock()

def get_session():
    global _session
    with _session_lock:
        if _session is None:
            _session = make_session()
    return _session

HTML = r"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>國語辭典查詢</title>
  <style>
    *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      position: relative;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "PingFang TC", "Noto Sans TC", sans-serif;
      background: #0f1117; color: #e2e8f0;
      min-height: 100vh; display: flex; flex-direction: column;
      align-items: center; padding: 48px 16px 80px;
    }
    h1 { font-size: 1.6rem; font-weight: 700; letter-spacing: .04em; color: #f8fafc; margin-bottom: 6px; }
    .subtitle { font-size: .85rem; color: #64748b; margin-bottom: 36px; }
    .card {
      background: #1e2330; border: 1px solid #2d3548;
      border-radius: 14px; padding: 28px 32px; width: 100%; max-width: 780px;
    }
    label { display: block; font-size: .8rem; color: #94a3b8; margin-bottom: 8px;
            letter-spacing: .06em; text-transform: uppercase; }
    .input-row { display: flex; gap: 10px; }
    input[type="text"] {
      flex: 1; background: #0f1117; border: 1.5px solid #2d3548;
      border-radius: 8px; padding: 12px 16px; font-size: 1rem;
      color: #e2e8f0; outline: none; transition: border-color .2s;
    }
    input[type="text"]:focus { border-color: #4f80ff; }
    input[type="text"]::placeholder { color: #475569; }
    button {
      cursor: pointer; border: none; border-radius: 8px;
      font-size: .95rem; font-weight: 600; padding: 12px 22px;
      transition: opacity .15s, transform .1s; white-space: nowrap;
    }
    button:active { transform: scale(.97); }
    button:disabled, button.busy { opacity: .45; cursor: not-allowed; }
    button.busy:active { transform: none; }
    #btn-search { background: #4f80ff; color: #fff; }
    #btn-search:hover:not(:disabled) { background: #3b6ef0; }
    #btn-export { background: #1a3a2a; color: #4ade80; border: 1px solid #166534; display: none; }
    #btn-export:hover:not(:disabled) { background: #14532d; }
    #status { margin-top: 18px; font-size: .82rem; color: #64748b;
              min-height: 20px; display: flex; align-items: center; gap: 8px; }
    .spinner { width: 14px; height: 14px; border: 2px solid #2d3548;
               border-top-color: #4f80ff; border-radius: 50%;
               animation: spin .7s linear infinite; display: none; }
    @keyframes spin { to { transform: rotate(360deg); } }
    .rw-tokens { margin-top: 14px; line-height: 2.3; font-size: 1.05rem; }
    .rw-tok { display: inline; color: #cbd5e1; }
    .rw-tok.pick { background: #232b40; border-radius: 5px; padding: 3px 6px; margin: 0 2px; cursor: pointer;
                    border: 1px solid #2d3548; }
    .rw-tok.pick.has-syn { border-color: #2563eb; color: #f1f5f9; }
    .rw-tok.pick:hover { background: #1e3a5f; }
    .rw-tok.pick.active { background: #1e3a5f; border-color: #7dd3fc; }
    .rw-tok.pick.changed { background: #3a2a10; border-color: #f59e0b; color: #fbbf24; }
    .rw-hint { font-size: .75rem; color: #64748b; margin-top: 8px; }
    .rw-panel { display: none; margin-top: 14px; background: #0f1117; border: 1px solid #2d3548;
                 border-radius: 8px; padding: 12px 14px; }
    .rw-panel-title { font-size: .85rem; color: #94a3b8; margin-bottom: 10px; }
    .rw-panel-title b { color: #f1f5f9; font-size: 1rem; }
    .rw-syns { display: flex; flex-wrap: wrap; gap: 8px; }
    .rw-syn { display: inline-flex; align-items: center; gap: 6px; background: #1e2330; color: #e2e8f0;
               border: 1px solid #2d3548; border-radius: 7px; padding: 6px 10px; font-size: .9rem; }
    .rw-syn:hover:not(:disabled) { border-color: #7dd3fc; }
    .rw-syn.chosen { border-color: #f59e0b; background: #3a2a10; }
    .drop-overlay { position: fixed; inset: 0; z-index: 900; display: none; align-items: center; justify-content: center;
                     background: rgba(15,17,23,.82); border: 3px dashed #7dd3fc; pointer-events: none;
                     color: #7dd3fc; font-size: 1.4rem; font-weight: 700; text-align: center; line-height: 1.8; }
    .drop-overlay.show { display: flex; }
    .drop-hint { flex: 1; color: #64748b; font-size: .85rem; border: 1.5px dashed #2d3548; border-radius: 8px;
                  padding: 10px 14px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
    .drop-hint.loaded { color: #7dd3fc; border-color: #2563eb; border-style: solid; }
    .drop-overlay small { font-size: .85rem; color: #94a3b8; font-weight: 400; }
    .tabs { display: flex; gap: 8px; width: 100%; max-width: 780px; margin-bottom: 18px; }
    .tab { background: #1e2330; color: #94a3b8; border: 1px solid #2d3548; }
    .tab.active { background: #1e3a5f; color: #7dd3fc; border-color: #2563eb; }
    .page { display: none; width: 100%; flex-direction: column; align-items: center; }
    .page.active { display: flex; }
    .progress-slot { width: 100%; max-width: 780px; display: flex; justify-content: center; }
    .rw-work { display: none; margin-top: 20px; }
    .rw-actions { display: flex; gap: 10px; justify-content: flex-end; margin-top: 14px; flex-wrap: wrap; }
    .btn-go { background: #1a3a2a !important; color: #4ade80 !important; border: 1px solid #166534 !important; }
    .rw-history { width: 100%; max-width: 780px; }
    .hw-grid { width: 100%; max-width: 780px; margin-top: 20px; display: flex; flex-wrap: wrap; gap: 14px;
                justify-content: center; }
    .hw-cell { background: #f8fafc; border-radius: 10px; overflow: hidden; width: 300px; position: relative; }
    .hw-cell iframe { display: block; border: 0; width: 300px; height: 520px; position: relative; }
    /* 教育部網站約要 5 秒才畫出動畫，iframe 底下先墊一行載入中 */
    .hw-cell::before { content: "教育部動畫載入中…"; position: absolute; top: 240px; left: 0; right: 0;
                        text-align: center; color: #94a3b8; font-size: .85rem; }
    .hw-cell .hw-cap { display: flex; justify-content: space-between; padding: 6px 10px; font-size: .78rem;
                        background: #1e2330; color: #94a3b8; }
    .hw-cell .hw-cap a { color: #7dd3fc; }
    .btn-open { display: inline-block; margin: 6px 0; padding: 10px 16px; border-radius: 8px; background: #1a3a2a;
                 color: #4ade80; border: 1px solid #166534; font-weight: 700; text-decoration: none; }
    .hw-zipbar { display: none; width: 100%; max-width: 780px; margin-top: 16px; justify-content: flex-end; }
    .hw-dl { width: 100%; border-radius: 0; background: #1a3a2a; color: #4ade80; border: 0;
              border-top: 1px solid #166534; padding: 10px; font-size: .85rem; }
    .hw-dl:hover:not(:disabled) { background: #14532d; }
    .hw-miss { background: #1e2330; border: 1px dashed #2d3548; border-radius: 10px; width: 300px; padding: 18px;
                color: #f87171; font-size: .9rem; text-align: center; }
    .hw-miss b { display: block; font-size: 2.4rem; color: #e2e8f0; margin-bottom: 6px; }
    .rw-result { margin-top: 20px; }
    .rw-result-head { font-size: .85rem; color: #7dd3fc; font-weight: 600; margin-bottom: 12px; }
    .rw-copy-row { display: flex; gap: 10px; align-items: flex-start; margin-bottom: 10px; }
    .rw-copy-row .rw-box { flex: 1; }
    .rw-copy-row button { white-space: nowrap; min-width: 130px; }
    .rw-ask { display: flex; gap: 10px; align-items: center; justify-content: flex-end; margin-top: 6px;
               font-size: .9rem; color: #e2e8f0; flex-wrap: wrap; }
    .rw-result .table-wrap { display: block; margin-top: 14px; }
    .rw-compare { display: none; width: 100%; max-width: 780px; margin-top: 14px;
                   grid-template-columns: 1fr 1fr; gap: 14px; }
    .rw-box { background: #1e2330; border: 1px solid #2d3548; border-radius: 10px; padding: 14px 16px;
               line-height: 1.9; white-space: pre-wrap; }
    .rw-box h3 { font-size: .8rem; color: #7dd3fc; margin-bottom: 8px; letter-spacing: .06em; }
    .rw-box mark { background: #3a2a10; color: #fbbf24; border-radius: 4px; padding: 0 3px; }
    .rw-box del { color: #f87171; text-decoration: none; background: #3a1a1a; border-radius: 4px; padding: 0 3px; }
    @media (max-width: 640px) { .rw-compare { grid-template-columns: 1fr; } .rw-copy-row { flex-direction: column; } }
    .table-wrap { margin-top: 28px; overflow-x: auto; border-radius: 10px;
                  border: 1px solid #2d3548; display: none; }
    table { width: 100%; border-collapse: collapse; font-size: .92rem; }
    thead th {
      background: #1a2236; padding: 12px 16px; text-align: left;
      font-size: .75rem; letter-spacing: .08em; text-transform: uppercase;
      color: #94a3b8; border-bottom: 1px solid #2d3548;
    }
    thead th:first-child { border-radius: 10px 0 0 0; }
    thead th:last-child  { border-radius: 0 10px 0 0; }
    tbody tr { border-bottom: 1px solid #1e2330; transition: background .15s;
               animation: fadeIn .3s ease; }
    tbody tr:last-child { border-bottom: none; }
    tbody tr:hover { background: #232b40; }
    @keyframes fadeIn { from { opacity: 0; transform: translateY(4px); } to { opacity: 1; transform: none; } }
    tbody tr:nth-child(even) { background: #1a2030; }
    tbody tr:nth-child(even):hover { background: #232b40; }
    td { padding: 13px 16px; vertical-align: top; }
    td.word { font-weight: 700; color: #f1f5f9; font-size: 1rem; }
    td.pin  { color: #7dd3fc; font-style: italic; white-space: nowrap; }
    td.pos  { color: #a78bfa; white-space: nowrap; }
    td.def  { color: #cbd5e1; line-height: 1.6; white-space: pre-wrap; }
    td.error { color: #f87171; font-size: .85rem; }
    td.lv { color: #94a3b8; font-size: .82rem; white-space: nowrap; }
    .badge { display: inline-block; background: #2d1f5e; color: #a78bfa;
             border: 1px solid #4c2d9c; border-radius: 5px;
             padding: 1px 7px; font-size: .78rem; font-weight: 600; }
    .lv-tag { display: inline-block; border-radius: 5px; padding: 2px 8px;
              font-size: .76rem; font-weight: 600; white-space: nowrap; }
    .lv-基礎 { background: #1a3a2a; color: #4ade80; border: 1px solid #166534; }
    .lv-進階 { background: #1e2e4a; color: #60a5fa; border: 1px solid #1d4ed8; }
    .lv-精熟 { background: #3a1a2a; color: #f472b6; border: 1px solid #9d174d; }
    #btn-bpmf { background: #1a2a3a; color: #64748b; border: 1px solid #2d3548; font-size:.85rem; }
    #btn-bpmf.active { background: #1e3a5f; color: #7dd3fc; border-color: #2563eb; }
    #bpmf-hint { font-size:.82rem; color:#7dd3fc; margin-top:8px; min-height:18px; letter-spacing:.04em; }
    .progress-wrap { display: none; width: 100%; max-width: 780px; margin-top: 18px;
                      flex-direction: column; align-items: center; gap: 6px; }
    .progress-body { width: 100%; }
    .loading-pet { width: 160px; height: 160px; object-fit: contain;
                    filter: drop-shadow(0 4px 10px rgba(0,0,0,.55));
                    animation: petBob 1s ease-in-out infinite; }
    @keyframes petBob {
      0%,100% { transform: translateY(0) rotate(0deg); }
      50% { transform: translateY(-4px) rotate(-2deg); }
    }
    .progress-label { display: flex; justify-content: space-between; align-items: baseline;
                       font-size: .8rem; color: #94a3b8; margin-bottom: 7px; letter-spacing: .04em; }
    .progress-label .progress-pct { color: #fbbf24; font-weight: 700; font-size: .95rem; }
    .progress-track { width: 100%; height: 16px; background: #0b0d13; border: 1.5px solid #2d3548;
                       border-radius: 999px; overflow: hidden; box-shadow: inset 0 1px 4px rgba(0,0,0,.5); }
    .progress-fill { height: 100%; width: 0%; background: linear-gradient(90deg,#f59e0b,#fbbf24);
                      border-radius: 999px; transition: width .25s ease;
                      box-shadow: 0 0 10px rgba(251,191,36,.6);
                      animation: progressPulse 1.4s ease-in-out infinite; }
    @keyframes progressPulse {
      0%,100% { box-shadow: 0 0 6px rgba(251,191,36,.45); }
      50% { box-shadow: 0 0 18px rgba(251,191,36,.9); }
    }
    .joke-text { margin-top: 10px; font-size: .82rem; color: #64748b; font-style: italic;
                  min-height: 16px; text-align: center; }
    .segment-line { display: none; width: 100%; max-width: 780px; margin-top: 16px;
                     background: #1a2236; border: 1px solid #2d3548; border-radius: 8px;
                     padding: 10px 14px; font-size: .85rem; overflow-x: auto; white-space: nowrap; }
    .segment-line .seg-label { color: #7dd3fc; font-weight: 600; margin-right: 8px; white-space: nowrap; }
    .segment-line .seg-word { display: inline-block; background: #232b40; color: #cbd5e1;
                               border-radius: 5px; padding: 2px 8px; margin-right: 6px; font-size: .82rem;
                               cursor: pointer; transition: background .15s; }
    .segment-line .seg-word:hover { background: #2d3a5a; }
    .segment-line .seg-word.on { background: #2563eb; color: #fff; }
    .rw-result .segment-line { display: block; max-width: none; }
    /* 顯示用表格上方的篩選列：全部顯示／各詞彙等級，點斷詞時多一個「只顯示某詞」標籤 */
    .result-filter { display: flex; flex-wrap: wrap; gap: 6px; align-items: center;
                      padding: 10px 12px; background: #161b27; border-bottom: 1px solid #2d3548; }
    .result-filter button { font-size: .78rem; font-weight: 600; padding: 4px 11px; border-radius: 999px;
                             background: #1a2030; color: #cbd5e1; border: 1px solid #2d3548; }
    .result-filter button:hover { background: #232b40; }
    .result-filter button.on { background: #2563eb; color: #fff; border-color: #2563eb; }
    .result-filter .rf-word { margin-left: auto; font-size: .8rem; color: #7dd3fc; }
    .result-filter .rf-word button { margin-left: 6px; padding: 2px 9px; }
    td .more-btn { display: block; margin-top: 6px; font-size: .75rem; font-weight: 600; padding: 3px 10px;
                    background: #1e3a5f; color: #7dd3fc; border: 1px solid #2563eb; border-radius: 6px; }
    td .more-btn:hover { background: #24497a; }
    td.more-sum { color: #64748b; font-size: .85rem; }
    tbody tr.more-row, tbody tr.more-row:nth-child(even) { background: #141926; }
    tbody tr.more-row td.word { color: #94a3b8; font-weight: 600; padding-left: 28px; }
    td.rf-empty { text-align: center; color: #64748b; padding: 22px; }
    /* 第一次打開的完整教學：右上角一鍵跳過 */
    .tour-skip { position: fixed; top: 16px; right: 18px; z-index: 850; display: none;
                  background: #1e2330; color: #fbbf24; border: 1px solid #fbbf24; padding: 8px 16px;
                  font-size: .85rem; box-shadow: 0 6px 18px rgba(0,0,0,.45); }
    .tour-skip.show { display: block; }
    .tour-arrow { position: fixed; z-index: 850; display: none; font-size: 2.6rem; line-height: 1;
                   color: #fbbf24; text-shadow: 0 0 14px rgba(251,191,36,.6); pointer-events: none;
                   animation: arrowNudge .9s ease-in-out infinite; }
    .tour-arrow.show { display: block; }
    @keyframes arrowNudge { 0%,100% { transform: translateY(0); } 50% { transform: translateY(10px); } }
    .tour-skip:hover { background: #2a2410; }
    .fullscreen-overlay img.skip-img { width: min(80vw, 520px); max-height: none; height: auto;
                                        image-rendering: auto; border-radius: 10px; }
    .header-row { display: flex; align-items: flex-start; gap: 18px; margin-bottom: 22px; }
    .logo-menu { position: relative; }
    .site-logo { height: 88px; width: auto; display: block; cursor: pointer;
                  filter: drop-shadow(0 3px 8px rgba(0,0,0,.5));
                  transition: transform .15s ease; }
    .site-logo:hover { transform: scale(1.06); }
    .logo-dropdown { position: absolute; top: 100%; left: 0; margin-top: 10px;
                      background: #1e2330; border: 1px solid #2d3548; border-radius: 10px;
                      min-width: 250px; padding: 6px; box-shadow: 0 14px 34px rgba(0,0,0,.55);
                      display: none; flex-direction: column; z-index: 600; }
    .logo-dropdown.show { display: flex; }
    .dd-item { display: block; padding: 10px 12px; border-radius: 7px; cursor: pointer;
                color: #e2e8f0; text-decoration: none; }
    .dd-item:hover { background: #232b40; }
    .dd-title { font-size: .92rem; font-weight: 600; }
    .dd-sub { display: block; font-size: .76rem; color: #7dd3fc; margin-top: 3px; text-decoration: none; }
    .fullscreen-overlay { position: fixed; inset: 0; background: rgba(4,5,9,.94);
                           display: none; flex-direction: column; align-items: center;
                           justify-content: center; gap: 20px; z-index: 900; cursor: pointer; }
    .fullscreen-overlay.show { display: flex; }
    .fullscreen-overlay img { max-width: min(80vw, 480px); max-height: 58vh; object-fit: contain; }
    .fullscreen-overlay .roast-caption { font-size: 1.3rem; font-weight: 700; color: #fbbf24;
                                          text-align: center; padding: 0 24px; }
    /* 桌子和螢幕是黑的，墊一圈淡光暈才看得出輪廓 */
    .fullscreen-overlay img.okfine-img { width: min(70vw, 360px); height: auto;
      background: radial-gradient(circle, rgba(203,213,225,.55) 0%, rgba(148,163,184,.25) 45%, transparent 70%);
      /* 桌子、螢幕貼著原圖的下緣和左緣，整張圖往四周淡出，避免被切出一條硬邊 */
      -webkit-mask-image: radial-gradient(circle at 55% 45%, #000 50%, transparent 71%);
      mask-image: radial-gradient(circle at 55% 45%, #000 50%, transparent 71%); }
    .fullscreen-overlay img.impatient-img { max-width: min(90vw, 860px); max-height: 72vh; border-radius: 10px;
                                             box-shadow: 0 12px 36px rgba(0,0,0,.6); }
    .fullscreen-overlay img.thanks-img { width: min(60vw, 300px); height: auto; }
    /* 人物貼著原圖下緣，下方淡出，避免被切出一條硬邊 */
    .fullscreen-overlay img.why-img { -webkit-mask-image: linear-gradient(to top, transparent, #000 22%);
                                       mask-image: linear-gradient(to top, transparent, #000 22%); }
    .tour-pick { position: fixed; inset: 0; z-index: 850; display: none; align-items: center; justify-content: center;
                  background: rgba(4,5,9,.72); }
    .tour-pick.show { display: flex; }
    .tour-pick-box { background: #1e2330; border: 1px solid #2d3548; border-radius: 14px; padding: 24px 28px;
                      display: flex; flex-direction: column; gap: 10px; min-width: 260px; }
    .tour-pick-title { font-size: 1.15rem; font-weight: 700; color: #f8fafc; margin-bottom: 6px; text-align: center; }
    .tour-pick-box .tour-cancel { background: transparent; color: #64748b; border: 1px solid #2d3548; }
    .tour-bar { position: fixed; left: 50%; bottom: 18px; transform: translateX(-50%); z-index: 800;
                 width: min(92vw, 720px); display: none; background: #172033; border: 2px solid #fbbf24;
                 border-radius: 12px; padding: 14px 18px; box-shadow: 0 12px 34px rgba(0,0,0,.6); cursor: pointer; }
    .tour-bar.show { display: block; }
    .tour-step { font-size: .75rem; color: #fbbf24; font-weight: 700; margin-bottom: 4px; }
    .tour-text { font-size: .95rem; color: #f1f5f9; line-height: 1.7; }
    .tour-hint { font-size: .75rem; color: #94a3b8; margin-top: 6px; text-align: right; }
    /* 教學中目前在講的那個元件 */
    .tour-focus { outline: 3px solid #fbbf24 !important; outline-offset: 4px;
                   box-shadow: 0 0 0 8px rgba(251,191,36,.18), 0 0 26px rgba(251,191,36,.45) !important;
                   border-radius: 8px; transition: outline-color .2s;
                   scroll-margin-top: 24px; scroll-margin-bottom: 190px; }  /* 捲動時停在說明列上方 */
    body.touring { padding-bottom: 220px; }  /* 最底下的元件也能捲到說明列上方 */
    .fullscreen-overlay .roast-hint { font-size: .8rem; color: #64748b; }
    .toast-overlay { position: fixed; inset: 0; background: rgba(4,5,9,.6);
                      display: none; align-items: center; justify-content: center;
                      z-index: 900; cursor: pointer; }
    .toast-overlay.show { display: flex; }
    .toast-box { background: #1e2330; border: 1px solid #2d3548; border-radius: 12px;
                  padding: 22px 30px; font-size: 1.05rem; color: #f8fafc; text-align: center; }
    .credits-box { min-width: 260px; text-align: left; padding: 24px 32px; }
    .credits-title { font-size: 1.15rem; font-weight: 700; color: #f8fafc;
                      text-align: center; margin-bottom: 16px; letter-spacing: .04em; }
    .credits-row { display: flex; flex-direction: column; gap: 2px;
                    padding: 10px 0; border-top: 1px solid #2d3548; }
    .credits-row:first-of-type { border-top: none; }
    .credits-name { font-size: 1rem; font-weight: 600; color: #e2e8f0; }
    .credits-role { font-size: .78rem; color: #7dd3fc; }
    .credits-note { margin-top: 14px; padding-top: 14px; border-top: 1px solid #2d3548;
                     font-size: .82rem; line-height: 1.6; color: #94a3b8; }
    .credits-note a { color: #7dd3fc; }
    .credits-coffee { margin-top: 10px; font-size: .85rem; color: #cbd5e1; }
    .credits-coffee a { color: #fbbf24; font-weight: 700; }
    .fullscreen-overlay video { max-width: min(80vw, 420px); max-height: 66vh; border-radius: 12px;
                                 box-shadow: 0 10px 30px rgba(0,0,0,.6); }
    .credits-hint { margin-top: 16px; font-size: .76rem; color: #64748b; text-align: center; }
  </style>
</head>
<body>
  <div class="header-row">
    <div class="logo-menu" id="logo-menu">
      <img id="site-logo" class="site-logo" src="/branding/logo.png" alt="選單">
      <div class="logo-dropdown" id="logo-dropdown">
        <a class="dd-item" id="dd-join" href="https://discord.gg/sPUY4kYJ92" target="_blank" rel="noopener"><div class="dd-title">加入我們</div></a>
        <div class="dd-item" id="dd-tutorial"><div class="dd-title">新手教學</div></div>
        <a class="dd-item" id="dd-relax" href="https://youtu.be/dQw4w9WgXcQ?si=3Lh9VQOuqhgY5Wfi" target="_blank" rel="noopener"><div class="dd-title">放鬆一下</div></a>
        <div class="dd-item" id="dd-roast"><div class="dd-title">網站做太差？</div></div>
        <div class="dd-item" id="dd-credits"><div class="dd-title">作者頁</div></div>
      </div>
    </div>
    <div style="margin-top:22px;">
      <h1>國語辭典查詢</h1>
      <p class="subtitle">教育部《國語辭典簡編本》· 萌典 · 國教院「教材編輯輔助系統」斷詞 · 自動輸出 Excel</p>
    </div>
  </div>
  <div class="fullscreen-overlay" id="dog-overlay">
    <video id="dog-video" src="/branding/good-person-dog.mp4" loop playsinline preload="none"></video>
    <div class="roast-caption">你是好人，給你看狗狗</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="impatient-overlay">
    <img id="impatient-img" class="impatient-img" src="/branding/impatient.jpg" alt="你的性子也太急了">
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="meteor-overlay">
    <img id="meteor-img" class="thanks-img" src="/mascot/meteor_00.png" alt="">
    <div class="roast-caption">忘記切換鍵盤了齁</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="why-overlay">
    <img id="why-img" class="thanks-img why-img" src="/mascot/why_00.png" alt="">
    <div class="roast-caption">那你點我幹嘛</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="thanks-overlay">
    <img id="thanks-img" class="thanks-img" src="/mascot/thanks_00.png" alt="">
    <div class="roast-caption">說謝謝</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="tour-pick" id="tour-pick">
    <div class="tour-pick-box">
      <div class="tour-pick-title">你想要知道關於</div>
      <button data-tour="all">全部功能</button>
      <button data-tour="lookup">查詢表格</button>
      <button data-tour="rewrite">近義詞替換</button>
      <button data-tour="handwriting">筆順動畫</button>
      <button class="tour-cancel" data-tour="">取消</button>
    </div>
  </div>
  <div class="tour-bar" id="tour-bar">
    <div class="tour-step" id="tour-step"></div>
    <div class="tour-text" id="tour-text"></div>
    <div class="tour-hint" id="tour-hint"></div>
  </div>
  <button class="tour-skip" id="tour-skip">一鍵跳過教學 ⏭</button>
  <div class="tour-arrow" id="tour-arrow">⬆</div>
  <div class="fullscreen-overlay" id="skip-overlay">
    <img id="skip-img" class="skip-img" src="/branding/skip-okay.gif" alt="">
    <div class="roast-caption">好吧</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="okfine-overlay">
    <img id="okfine-img" class="okfine-img" src="/mascot/okfine_00.png" alt="">
    <div class="roast-caption">喔好吧</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="fullscreen-overlay" id="roast-overlay">
    <img src="/branding/roast-dog.png" alt="">
    <div class="roast-caption">罵了他就不能罵我了喔</div>
    <div class="roast-hint">按任意鍵關閉</div>
  </div>
  <div class="toast-overlay" id="credits-overlay">
    <div class="toast-box credits-box">
      <div class="credits-title">作者頁</div>
      <div class="credits-row">
        <span class="credits-name">林錦崧</span>
        <span class="credits-role">銘傳大學華語文教學學系 · 分詞語法邏輯設計（載入中的文案是他寫的）</span>
      </div>
      <div class="credits-row">
        <span class="credits-name">林佩萱</span>
        <span class="credits-role">國立中興大學電機工程學系 · 網頁開發</span>
      </div>
      <p class="credits-note">
        這是我們的專題題目，做得倉促的地方應該不少，如果你發現什麼奇怪的、覺得可以改進的部分，
        真的很歡迎跟我們說一聲——<a href="mailto:linjefferson0518@gmail.com">linjefferson0518@gmail.com</a>。
      </p>
      <p class="credits-coffee">歡迎請作者喝杯<a href="#" id="coffee-link">咖啡</a></p>
      <div class="credits-hint">點擊任意處關閉</div>
    </div>
  </div>
  <div class="drop-overlay" id="drop-overlay"><div>放開滑鼠，載入檔案文字<br><small>支援 txt／docx／pdf</small></div></div>
  <div class="tabs" id="tabs">
    <button class="tab active" data-page="lookup">查詞表格</button>
    <button class="tab" data-page="rewrite">近義詞替換</button>
    <button class="tab" data-page="handwriting">筆順動畫</button>
  </div>
  <div class="page active" id="page-lookup">
  <div class="card">
    <label for="words-input">輸入詞彙</label>
    <div class="input-row">
      <input type="text" id="words-input" placeholder="例：給予 讚美、進食；快樂,熱血"
             autocomplete="off" spellcheck="false">
      <button id="btn-bpmf" title="大千式注音輸入模式">⌨ 注音</button>
      <button id="btn-search">查詢</button>
      <button id="btn-export">下載 Excel</button>
    </div>
    <div id="bpmf-hint"></div>
    <div id="status"><div class="spinner" id="spinner"></div><span id="status-text"></span></div>
  </div>
  <div class="card" style="margin-top:20px;">
    <label for="paste-text">整段文字（貼上，或把檔案拖進視窗）</label>
    <textarea id="paste-text" rows="5" placeholder="貼上一整段文字，或把檔案拖進視窗，會先斷詞再查詢"
      style="width:100%;resize:vertical;background:#0f1117;border:1.5px solid #2d3548;border-radius:8px;
             padding:12px 16px;font-size:.95rem;color:#e2e8f0;outline:none;font-family:inherit;"></textarea>
    <div class="input-row" style="margin-top:12px;align-items:center;">
      <div class="drop-hint" id="lookup-file">把 txt／docx／pdf 檔直接拖進視窗即可載入</div>
      <button id="btn-segment">斷詞並查詢</button>
    </div>
    <div style="margin-top:12px;">
      <label for="exclude-input" style="margin-bottom:5px;">排除詞（可留空，逗號分隔，例：張愛玲，附件二）</label>
      <input type="text" id="exclude-input" placeholder="標題括號【】和頁碼已自動去除；作者名等自訂詞填這裡"
             autocomplete="off" spellcheck="false"
             style="width:100%;background:#0f1117;border:1.5px solid #2d3548;border-radius:8px;
                    padding:10px 14px;font-size:.85rem;color:#e2e8f0;outline:none;">
    </div>
    <p style="margin-top:8px;font-size:.75rem;color:#64748b;">
      支援把 txt／docx／pdf 檔拖進視窗載入。用國教院「教材編輯輔助系統」（遠流語料，關聯詞數量 10）斷詞，
      自動去除標點、重複詞、【】標題括號、頁碼裝飾（如「- 1 -」）、
      以及含阿拉伯數字或非中文字元的詞後查詢。
    </p>
  </div>
  <div class="progress-slot" id="lookup-progress-slot">
  <div class="progress-wrap" id="progress-wrap">
    <div class="progress-body">
      <div class="progress-label"><span>查詢進度</span><span class="progress-pct" id="progress-pct">0%</span></div>
      <div class="progress-track"><div class="progress-fill" id="progress-fill"></div></div>
      <div class="joke-text" id="joke-text"></div>
    </div>
    <img id="loading-pet" class="loading-pet" src="/mascot/bunny_typing_00.png" alt="查詢中">
  </div>
  </div>
  <div class="segment-line" id="segment-line"></div>
  <div class="table-wrap" id="table-wrap">
    <div class="result-filter" id="result-filter"></div>
    <table>
      <thead><tr><th>漢字詞彙</th><th>音標</th><th>詞類</th><th>意思</th><th>詞彙等級</th></tr></thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
  </div>
  <div class="page" id="page-rewrite">
    <div class="card">
      <label for="rw-text">文本（用滑鼠選取一段，再按「近義詞替換」）</label>
      <textarea id="rw-text" rows="12" placeholder="貼上文本，或把檔案拖進視窗；切到這頁時若是空的，會自動帶入「查詞表格」頁貼的整段文字"
        style="width:100%;resize:vertical;background:#0f1117;border:1.5px solid #2d3548;border-radius:8px;
               padding:12px 16px;font-size:.95rem;line-height:1.8;color:#e2e8f0;outline:none;font-family:inherit;"></textarea>
      <div class="input-row" style="margin-top:12px;align-items:center;">
        <div class="drop-hint" id="rw-file">把 txt／docx／pdf 檔直接拖進視窗即可載入</div>
        <button id="btn-rw-seg">近義詞替換</button>
      </div>
      <div id="rw-status" class="rw-hint"></div>
    </div>
    <div class="progress-slot" id="rewrite-progress-slot"></div>
    <div class="card rw-work" id="rw-work">
      <label>選取段落（點詞選近義詞，藍框＝有近義詞）</label>
      <div class="rw-tokens" id="rw-tokens"></div>
      <div class="rw-panel" id="rw-panel">
        <div class="rw-panel-title" id="rw-panel-title"></div>
        <div class="rw-syns" id="rw-syns"></div>
      </div>
      <div class="rw-compare" id="rw-compare">
        <div class="rw-box"><h3>原句</h3><div id="rw-orig"></div></div>
        <div class="rw-box"><h3>改寫句</h3><div id="rw-new"></div></div>
      </div>
      <div class="rw-actions">
        <button id="btn-rw-cancel">取消</button>
        <button id="btn-rw-reset">全部還原</button>
        <button id="btn-rw-confirm" class="btn-go">確認替換</button>
      </div>
    </div>
    <div class="rw-history" id="rw-history"></div>
  </div>
  <div class="page" id="page-handwriting">
    <div class="card">
      <label for="hw-input">輸入一個字或一段話（最多 20 字，標點與空白會略過）</label>
      <div class="input-row">
        <input type="text" id="hw-input" placeholder="例：永、學習" autocomplete="off" spellcheck="false">
        <button id="btn-hw">查筆順</button>
      </div>
      <div id="hw-status" class="rw-hint"></div>
      <p style="margin-top:8px;font-size:.75rem;color:#64748b;">
        筆順動畫來源：中華民國教育部《國字標準字體筆順學習網》
        <a href="https://stroke-order.learningweb.moe.edu.tw/" target="_blank" rel="noopener" style="color:#7dd3fc;">stroke-order.learningweb.moe.edu.tw</a>，
        以官方嵌入碼原樣呈現（創用CC 姓名標示－非商業性－禁止改作 3.0 臺灣版）。收錄教育部標準字體（繁體），簡體字查不到。
        下載的是「白底 MP4」：灰底換成白色、未寫的字框淡灰、筆畫黑色，供自編教材使用，放進教材時請註明「中華民國教育部《國字標準字體筆順學習網》」。
        以官方「快」速錄製，每字約 5–20 秒（兩個字同時錄）；需要電腦上有 Chrome 或 Edge。
      </p>
    </div>
    <div class="progress-slot" id="handwriting-progress-slot"></div>
    <div class="hw-zipbar" id="hw-zipbar"><button id="btn-hw-zip" class="btn-go"></button></div>
    <div class="hw-grid" id="hw-grid"></div>
  </div>
  <script>
    // ── 大千式注音輸入支援 ──────────────────────────────────────────────
    const BPMF_KEYS={'1':'ㄅ','q':'ㄆ','a':'ㄇ','z':'ㄈ','2':'ㄉ','w':'ㄊ','s':'ㄋ','x':'ㄌ',
      'e':'ㄍ','d':'ㄎ','c':'ㄏ','r':'ㄐ','f':'ㄑ','v':'ㄒ','5':'ㄓ','t':'ㄔ','g':'ㄕ','b':'ㄖ',
      'y':'ㄗ','h':'ㄘ','n':'ㄙ','u':'ㄧ','j':'ㄨ','m':'ㄩ','8':'ㄚ','i':'ㄛ','k':'ㄜ',
      '9':'ㄞ','o':'ㄟ','l':'ㄠ','.':'ㄡ','0':'ㄢ','p':'ㄣ',';':'ㄤ','/':'ㄥ','-':'ㄦ',',':'ㄝ',
      '6':'ˊ','3':'ˇ','4':'ˋ','7':'˙'};
    const BPMF_CHARS={
      'ㄧ':'一','ㄧˊ':'移','ㄧˇ':'以','ㄧˋ':'意',
      'ㄨ':'烏','ㄨˊ':'無','ㄨˇ':'五','ㄨˋ':'物',
      'ㄨㄛ':'窩','ㄨㄛˊ':'沃','ㄨㄛˇ':'我','ㄨㄛˋ':'臥',
      'ㄨㄞ':'歪','ㄨㄞˋ':'外',
      'ㄨㄟ':'威','ㄨㄟˊ':'為','ㄨㄟˇ':'尾','ㄨㄟˋ':'位',
      'ㄨㄢ':'彎','ㄨㄢˊ':'完','ㄨㄢˇ':'晚','ㄨㄢˋ':'萬',
      'ㄨㄣ':'溫','ㄨㄣˊ':'文','ㄨㄣˇ':'穩','ㄨㄣˋ':'問',
      'ㄨㄤ':'汪','ㄨㄤˊ':'王','ㄨㄤˇ':'網','ㄨㄤˋ':'望',
      'ㄨㄥ':'翁','ㄨㄥˊ':'雄',
      'ㄩ':'迂','ㄩˊ':'魚','ㄩˇ':'語','ㄩˋ':'育',
      'ㄩㄝ':'約','ㄩㄝˊ':'月','ㄩㄝˇ':'樂','ㄩㄝˋ':'越',
      'ㄩㄢ':'冤','ㄩㄢˊ':'原','ㄩㄢˇ':'遠','ㄩㄢˋ':'願',
      'ㄩㄣ':'暈','ㄩㄣˊ':'雲','ㄩㄣˇ':'允','ㄩㄣˋ':'運',
      'ㄩㄥ':'擁','ㄩㄥˊ':'用','ㄩㄥˇ':'永',
      'ㄅㄚ':'八','ㄅㄚˊ':'拔','ㄅㄚˇ':'把','ㄅㄚˋ':'爸','ㄅㄚ˙':'吧',
      'ㄅㄛ':'波','ㄅㄛˊ':'伯','ㄅㄛˋ':'博',
      'ㄅㄞ':'掰','ㄅㄞˊ':'白','ㄅㄞˇ':'百','ㄅㄞˋ':'拜',
      'ㄅㄟ':'杯','ㄅㄟˊ':'陪','ㄅㄟˇ':'北','ㄅㄟˋ':'背',
      'ㄅㄠ':'包','ㄅㄠˊ':'薄','ㄅㄠˇ':'保','ㄅㄠˋ':'報',
      'ㄅㄢ':'班','ㄅㄢˊ':'辦','ㄅㄢˇ':'板','ㄅㄢˋ':'半',
      'ㄅㄣ':'奔','ㄅㄣˇ':'本','ㄅㄣˋ':'笨',
      'ㄅㄤ':'幫','ㄅㄤˊ':'旁','ㄅㄤˇ':'榜','ㄅㄤˋ':'棒',
      'ㄅㄥ':'崩','ㄅㄥˊ':'朋','ㄅㄥˋ':'迸',
      'ㄅㄧ':'逼','ㄅㄧˊ':'鼻','ㄅㄧˇ':'比','ㄅㄧˋ':'必',
      'ㄅㄧㄝ':'憋','ㄅㄧㄝˊ':'別','ㄅㄧㄝˋ':'鱉',
      'ㄅㄧㄠ':'標','ㄅㄧㄠˇ':'表',
      'ㄅㄧㄢ':'邊','ㄅㄧㄢˊ':'便','ㄅㄧㄢˇ':'扁','ㄅㄧㄢˋ':'變',
      'ㄅㄧㄣ':'賓','ㄅㄧㄣˋ':'鬢',
      'ㄅㄧㄥ':'冰','ㄅㄧㄥˇ':'丙','ㄅㄧㄥˋ':'病',
      'ㄅㄨ':'不','ㄅㄨˊ':'步','ㄅㄨˇ':'補','ㄅㄨˋ':'部',
      'ㄆㄚ':'趴','ㄆㄚˊ':'爬','ㄆㄚˋ':'怕',
      'ㄆㄛ':'坡','ㄆㄛˊ':'婆','ㄆㄛˇ':'頗','ㄆㄛˋ':'破',
      'ㄆㄞ':'拍','ㄆㄞˊ':'牌','ㄆㄞˋ':'派',
      'ㄆㄟ':'呸','ㄆㄟˊ':'賠','ㄆㄟˋ':'配',
      'ㄆㄠ':'拋','ㄆㄠˊ':'袍','ㄆㄠˇ':'跑','ㄆㄠˋ':'泡',
      'ㄆㄡ':'剖',
      'ㄆㄢ':'攀','ㄆㄢˊ':'盤','ㄆㄢˋ':'判',
      'ㄆㄣ':'噴','ㄆㄣˊ':'盆',
      'ㄆㄤ':'乓','ㄆㄤˊ':'旁','ㄆㄤˋ':'胖',
      'ㄆㄥ':'烹','ㄆㄥˊ':'朋','ㄆㄥˇ':'捧','ㄆㄥˋ':'碰',
      'ㄆㄧ':'批','ㄆㄧˊ':'皮','ㄆㄧˇ':'匹','ㄆㄧˋ':'屁',
      'ㄆㄧㄢ':'篇','ㄆㄧㄢˊ':'便','ㄆㄧㄢˇ':'片','ㄆㄧㄢˋ':'騙',
      'ㄆㄧㄠ':'飄','ㄆㄧㄠˊ':'瓢','ㄆㄧㄠˇ':'漂','ㄆㄧㄠˋ':'票',
      'ㄆㄧㄣ':'拼','ㄆㄧㄣˊ':'頻','ㄆㄧㄣˋ':'聘',
      'ㄆㄧㄥ':'乒','ㄆㄧㄥˊ':'平','ㄆㄧㄥˇ':'瓶','ㄆㄧㄥˋ':'評',
      'ㄆㄨ':'鋪','ㄆㄨˊ':'葡','ㄆㄨˇ':'普','ㄆㄨˋ':'舖',
      'ㄇㄚ':'媽','ㄇㄚˊ':'麻','ㄇㄚˇ':'馬','ㄇㄚˋ':'罵','ㄇㄚ˙':'嗎',
      'ㄇㄛ':'摸','ㄇㄛˊ':'磨','ㄇㄛˇ':'抹','ㄇㄛˋ':'末',
      'ㄇㄜ˙':'麼',
      'ㄇㄞ':'埋','ㄇㄞˇ':'買','ㄇㄞˋ':'賣',
      'ㄇㄟ':'眉','ㄇㄟˊ':'沒','ㄇㄟˇ':'美','ㄇㄟˋ':'妹',
      'ㄇㄠ':'貓','ㄇㄠˊ':'毛','ㄇㄠˇ':'卯','ㄇㄠˋ':'帽',
      'ㄇㄡ':'謀','ㄇㄡˊ':'某','ㄇㄡˇ':'某',
      'ㄇㄢ':'蠻','ㄇㄢˊ':'饅','ㄇㄢˇ':'滿','ㄇㄢˋ':'慢',
      'ㄇㄣ':'悶','ㄇㄣˊ':'門','ㄇㄣ˙':'們',
      'ㄇㄤ':'忙','ㄇㄤˊ':'忙','ㄇㄤˇ':'莽',
      'ㄇㄥ':'蒙','ㄇㄥˊ':'蒙','ㄇㄥˇ':'猛','ㄇㄥˋ':'夢',
      'ㄇㄧ':'彌','ㄇㄧˊ':'迷','ㄇㄧˇ':'米','ㄇㄧˋ':'密',
      'ㄇㄧㄝ':'滅',
      'ㄇㄧㄠ':'苗','ㄇㄧㄠˊ':'苗','ㄇㄧㄠˇ':'秒','ㄇㄧㄠˋ':'廟',
      'ㄇㄧㄢ':'眠','ㄇㄧㄢˊ':'棉','ㄇㄧㄢˇ':'免','ㄇㄧㄢˋ':'面',
      'ㄇㄧㄣ':'民','ㄇㄧㄣˊ':'民','ㄇㄧㄣˇ':'敏',
      'ㄇㄧㄥ':'明','ㄇㄧㄥˊ':'名','ㄇㄧㄥˇ':'明','ㄇㄧㄥˋ':'命',
      'ㄇㄨ':'木','ㄇㄨˊ':'模','ㄇㄨˇ':'母','ㄇㄨˋ':'目',
      'ㄈㄚ':'發','ㄈㄚˊ':'罰','ㄈㄚˇ':'法','ㄈㄚˋ':'髮',
      'ㄈㄛ':'佛','ㄈㄛˊ':'佛',
      'ㄈㄟ':'飛','ㄈㄟˊ':'肥','ㄈㄟˇ':'斐','ㄈㄟˋ':'費',
      'ㄈㄡ':'否','ㄈㄡˊ':'浮','ㄈㄡˇ':'否',
      'ㄈㄢ':'翻','ㄈㄢˊ':'凡','ㄈㄢˇ':'反','ㄈㄢˋ':'飯',
      'ㄈㄣ':'分','ㄈㄣˊ':'憤','ㄈㄣˇ':'粉','ㄈㄣˋ':'份',
      'ㄈㄤ':'方','ㄈㄤˊ':'房','ㄈㄤˇ':'訪','ㄈㄤˋ':'放',
      'ㄈㄥ':'風','ㄈㄥˊ':'逢','ㄈㄥˇ':'諷','ㄈㄥˋ':'奉',
      'ㄈㄨ':'夫','ㄈㄨˊ':'服','ㄈㄨˇ':'府','ㄈㄨˋ':'父',
      'ㄉㄚ':'搭','ㄉㄚˊ':'答','ㄉㄚˇ':'打','ㄉㄚˋ':'大',
      'ㄉㄜ':'得','ㄉㄜˊ':'得','ㄉㄜ˙':'的',
      'ㄉㄞ':'呆','ㄉㄞˇ':'逮','ㄉㄞˋ':'帶',
      'ㄉㄟˇ':'得',
      'ㄉㄠ':'刀','ㄉㄠˊ':'道','ㄉㄠˇ':'倒','ㄉㄠˋ':'到',
      'ㄉㄡ':'兜','ㄉㄡ˙':'都',
      'ㄉㄢ':'單','ㄉㄢˇ':'膽','ㄉㄢˋ':'但',
      'ㄉㄤ':'當','ㄉㄤˇ':'黨','ㄉㄤˋ':'當',
      'ㄉㄥ':'燈','ㄉㄥˇ':'等','ㄉㄥˋ':'瞪',
      'ㄉㄧ':'低','ㄉㄧˊ':'滴','ㄉㄧˇ':'底','ㄉㄧˋ':'地',
      'ㄉㄧㄝ':'爹','ㄉㄧㄝˊ':'碟','ㄉㄧㄝˋ':'跌',
      'ㄉㄧㄠ':'雕','ㄉㄧㄠˋ':'掉',
      'ㄉㄧㄢ':'點','ㄉㄧㄢˇ':'典','ㄉㄧㄢˋ':'電',
      'ㄉㄧㄥ':'丁','ㄉㄧㄥˇ':'頂','ㄉㄧㄥˋ':'定',
      'ㄉㄨ':'都','ㄉㄨˊ':'讀','ㄉㄨˇ':'賭','ㄉㄨˋ':'度',
      'ㄉㄨㄛ':'多','ㄉㄨㄛˊ':'奪','ㄉㄨㄛˇ':'躲','ㄉㄨㄛˋ':'墮',
      'ㄉㄨㄟ':'堆','ㄉㄨㄟˋ':'對',
      'ㄉㄨㄢ':'端','ㄉㄨㄢˇ':'短','ㄉㄨㄢˋ':'斷',
      'ㄉㄨㄥ':'東','ㄉㄨㄥˇ':'懂','ㄉㄨㄥˋ':'動',
      'ㄊㄚ':'他','ㄊㄚˊ':'她','ㄊㄚˇ':'塔','ㄊㄚˋ':'踏',
      'ㄊㄞ':'胎','ㄊㄞˊ':'臺','ㄊㄞˇ':'太','ㄊㄞˋ':'太',
      'ㄊㄠ':'掏','ㄊㄠˊ':'逃','ㄊㄠˇ':'討','ㄊㄠˋ':'套',
      'ㄊㄡ':'偷','ㄊㄡˊ':'頭',
      'ㄊㄢ':'貪','ㄊㄢˊ':'談','ㄊㄢˇ':'毯','ㄊㄢˋ':'嘆',
      'ㄊㄤ':'湯','ㄊㄤˊ':'糖','ㄊㄤˇ':'躺','ㄊㄤˋ':'燙',
      'ㄊㄥ':'疼','ㄊㄥˊ':'騰',
      'ㄊㄧ':'梯','ㄊㄧˊ':'題','ㄊㄧˇ':'體','ㄊㄧˋ':'替',
      'ㄊㄧㄝ':'貼','ㄊㄧㄝˊ':'鐵',
      'ㄊㄧㄠ':'挑','ㄊㄧㄠˊ':'條','ㄊㄧㄠˋ':'跳',
      'ㄊㄧㄢ':'天','ㄊㄧㄢˊ':'甜','ㄊㄧㄢˇ':'舔',
      'ㄊㄧㄥ':'聽','ㄊㄧㄥˊ':'庭','ㄊㄧㄥˇ':'挺',
      'ㄊㄨ':'禿','ㄊㄨˊ':'圖','ㄊㄨˇ':'土','ㄊㄨˋ':'兔',
      'ㄊㄨㄛ':'拖','ㄊㄨㄛˊ':'脫','ㄊㄨㄛˇ':'妥','ㄊㄨㄛˋ':'拓',
      'ㄊㄨㄟ':'推','ㄊㄨㄟˋ':'退',
      'ㄊㄨㄣ':'吞','ㄊㄨㄣˊ':'屯',
      'ㄊㄨㄥ':'通','ㄊㄨㄥˊ':'同','ㄊㄨㄥˇ':'桶','ㄊㄨㄥˋ':'痛',
      'ㄋㄚ':'拿','ㄋㄚˊ':'哪','ㄋㄚˇ':'哪','ㄋㄚˋ':'那','ㄋㄚ˙':'呢',
      'ㄋㄞ':'奶','ㄋㄞˇ':'乃','ㄋㄞˋ':'奈',
      'ㄋㄠ':'撓','ㄋㄠˇ':'腦','ㄋㄠˋ':'鬧',
      'ㄋㄢ':'難','ㄋㄢˊ':'南','ㄋㄢˋ':'難',
      'ㄋㄤ':'囊',
      'ㄋㄥ':'能','ㄋㄥˊ':'能',
      'ㄋㄧ':'呢','ㄋㄧˇ':'你','ㄋㄧˋ':'膩',
      'ㄋㄧㄢ':'念','ㄋㄧㄢˊ':'年','ㄋㄧㄢˋ':'念',
      'ㄋㄧㄠ':'鳥','ㄋㄧㄠˇ':'鳥',
      'ㄋㄧㄝ':'捏',
      'ㄋㄧㄣ':'您','ㄋㄧㄣˊ':'您',
      'ㄋㄧㄥ':'寧','ㄋㄧㄥˊ':'凝','ㄋㄧㄥˇ':'擰',
      'ㄋㄨ':'奴','ㄋㄨˊ':'怒','ㄋㄨˇ':'努','ㄋㄨˋ':'怒',
      'ㄋㄨㄛ':'挪','ㄋㄨㄛˊ':'諾','ㄋㄨㄛˋ':'糯',
      'ㄋㄨㄢ':'暖','ㄋㄨㄢˇ':'暖',
      'ㄋㄩ':'女','ㄋㄩˇ':'女','ㄋㄩˋ':'怒',
      'ㄋㄩㄝ':'虐','ㄋㄩㄝˋ':'虐',
      'ㄌㄚ':'拉','ㄌㄚˊ':'辣','ㄌㄚˋ':'辣','ㄌㄚ˙':'啦',
      'ㄌㄜ':'了','ㄌㄜˊ':'勒','ㄌㄜ˙':'了',
      'ㄌㄞ':'來','ㄌㄞˊ':'來','ㄌㄞˋ':'賴',
      'ㄌㄟ':'雷','ㄌㄟˊ':'累','ㄌㄟˇ':'壘','ㄌㄟˋ':'累',
      'ㄌㄠ':'撈','ㄌㄠˊ':'勞','ㄌㄠˇ':'老','ㄌㄠˋ':'澇',
      'ㄌㄡ':'樓','ㄌㄡˊ':'婁','ㄌㄡˋ':'漏',
      'ㄌㄢ':'蘭','ㄌㄢˊ':'蘭','ㄌㄢˇ':'覽','ㄌㄢˋ':'爛',
      'ㄌㄤ':'狼','ㄌㄤˊ':'涼','ㄌㄤˇ':'朗','ㄌㄤˋ':'浪',
      'ㄌㄥ':'楞','ㄌㄥˊ':'冷','ㄌㄥˇ':'冷',
      'ㄌㄧ':'里','ㄌㄧˊ':'離','ㄌㄧˇ':'裡','ㄌㄧˋ':'力',
      'ㄌㄧㄝ':'裂','ㄌㄧㄝˋ':'列',
      'ㄌㄧㄠ':'撩','ㄌㄧㄠˇ':'了','ㄌㄧㄠˋ':'料',
      'ㄌㄧㄢ':'蓮','ㄌㄧㄢˊ':'臉','ㄌㄧㄢˇ':'臉','ㄌㄧㄢˋ':'練',
      'ㄌㄧㄣ':'林','ㄌㄧㄣˊ':'臨','ㄌㄧㄣˋ':'淋',
      'ㄌㄧㄤ':'涼','ㄌㄧㄤˊ':'良','ㄌㄧㄤˇ':'兩','ㄌㄧㄤˋ':'量',
      'ㄌㄧㄥ':'靈','ㄌㄧㄥˊ':'零','ㄌㄧㄥˇ':'領','ㄌㄧㄥˋ':'令',
      'ㄌㄨ':'盧','ㄌㄨˊ':'路','ㄌㄨˇ':'旅','ㄌㄨˋ':'路',
      'ㄌㄨㄛ':'羅','ㄌㄨㄛˊ':'落','ㄌㄨㄛˋ':'落',
      'ㄌㄨㄢ':'亂','ㄌㄨㄢˋ':'亂',
      'ㄌㄨㄥ':'龍','ㄌㄨㄥˊ':'隆','ㄌㄨㄥˋ':'弄',
      'ㄌㄩ':'旅','ㄌㄩˊ':'呂','ㄌㄩˇ':'旅','ㄌㄩˋ':'率',
      'ㄌㄩㄝ':'略','ㄌㄩㄝˋ':'略',
      'ㄍㄚ':'嘎',
      'ㄍㄜ':'歌','ㄍㄜˊ':'格','ㄍㄜˇ':'各','ㄍㄜˋ':'個','ㄍㄜ˙':'個',
      'ㄍㄞ':'該','ㄍㄞˇ':'改','ㄍㄞˋ':'蓋',
      'ㄍㄠ':'高','ㄍㄠˇ':'稿','ㄍㄠˋ':'告',
      'ㄍㄡ':'溝','ㄍㄡˇ':'狗','ㄍㄡˋ':'夠',
      'ㄍㄢ':'乾','ㄍㄢˇ':'敢','ㄍㄢˋ':'幹',
      'ㄍㄣ':'根','ㄍㄣˊ':'跟',
      'ㄍㄤ':'剛','ㄍㄤˇ':'港',
      'ㄍㄥ':'更','ㄍㄥˇ':'梗','ㄍㄥˋ':'更',
      'ㄍㄨ':'姑','ㄍㄨˊ':'骨','ㄍㄨˇ':'古','ㄍㄨˋ':'故',
      'ㄍㄨㄚ':'瓜','ㄍㄨㄚˇ':'寡','ㄍㄨㄚˋ':'掛',
      'ㄍㄨㄞ':'乖','ㄍㄨㄞˇ':'拐','ㄍㄨㄞˋ':'怪',
      'ㄍㄨㄟ':'規','ㄍㄨㄟˇ':'鬼','ㄍㄨㄟˋ':'貴',
      'ㄍㄨㄢ':'關','ㄍㄨㄢˇ':'管','ㄍㄨㄢˋ':'慣',
      'ㄍㄨㄣ':'滾','ㄍㄨㄣˇ':'滾',
      'ㄍㄨㄤ':'光','ㄍㄨㄤˇ':'廣','ㄍㄨㄤˋ':'逛',
      'ㄍㄨㄥ':'工','ㄍㄨㄥˇ':'拱','ㄍㄨㄥˋ':'共',
      'ㄎㄚ':'卡','ㄎㄚˇ':'卡',
      'ㄎㄜ':'科','ㄎㄜˊ':'可','ㄎㄜˇ':'可','ㄎㄜˋ':'客',
      'ㄎㄞ':'開','ㄎㄞˇ':'凱',
      'ㄎㄠ':'烤','ㄎㄠˇ':'考','ㄎㄠˋ':'靠',
      'ㄎㄡ':'摳','ㄎㄡˇ':'口','ㄎㄡˋ':'扣',
      'ㄎㄢ':'看','ㄎㄢˇ':'砍','ㄎㄢˋ':'看',
      'ㄎㄤ':'扛','ㄎㄤˊ':'抗','ㄎㄤˋ':'抗',
      'ㄎㄥ':'坑',
      'ㄎㄨ':'哭','ㄎㄨˊ':'苦','ㄎㄨˇ':'苦','ㄎㄨˋ':'庫',
      'ㄎㄨㄚ':'誇','ㄎㄨㄚˋ':'跨',
      'ㄎㄨㄞ':'快','ㄎㄨㄞˋ':'快',
      'ㄎㄨㄟ':'虧','ㄎㄨㄟˊ':'葵','ㄎㄨㄟˋ':'愧',
      'ㄎㄨㄢ':'寬','ㄎㄨㄢˇ':'款',
      'ㄎㄨㄤ':'狂','ㄎㄨㄤˊ':'狂','ㄎㄨㄤˋ':'況',
      'ㄎㄨㄥ':'空','ㄎㄨㄥˊ':'孔','ㄎㄨㄥˇ':'恐','ㄎㄨㄥˋ':'控',
      'ㄏㄚ':'哈',
      'ㄏㄜ':'喝','ㄏㄜˊ':'和','ㄏㄜˋ':'喝',
      'ㄏㄞ':'海','ㄏㄞˊ':'還','ㄏㄞˇ':'海','ㄏㄞˋ':'害',
      'ㄏㄠ':'好','ㄏㄠˇ':'好','ㄏㄠˋ':'好',
      'ㄏㄡ':'喉','ㄏㄡˊ':'候','ㄏㄡˋ':'後',
      'ㄏㄢ':'喊','ㄏㄢˊ':'函','ㄏㄢˇ':'喊','ㄏㄢˋ':'漢',
      'ㄏㄣ':'很','ㄏㄣˊ':'痕','ㄏㄣˇ':'很','ㄏㄣˋ':'恨',
      'ㄏㄤ':'航','ㄏㄤˊ':'行',
      'ㄏㄥ':'哼','ㄏㄥˊ':'橫',
      'ㄏㄨ':'呼','ㄏㄨˊ':'胡','ㄏㄨˇ':'虎','ㄏㄨˋ':'護',
      'ㄏㄨㄚ':'花','ㄏㄨㄚˊ':'滑','ㄏㄨㄚˋ':'話',
      'ㄏㄨㄞ':'懷','ㄏㄨㄞˋ':'壞',
      'ㄏㄨㄟ':'回','ㄏㄨㄟˊ':'回','ㄏㄨㄟˇ':'毀','ㄏㄨㄟˋ':'會',
      'ㄏㄨㄢ':'歡','ㄏㄨㄢˊ':'還','ㄏㄨㄢˇ':'緩','ㄏㄨㄢˋ':'換',
      'ㄏㄨㄣ':'婚','ㄏㄨㄣˊ':'魂','ㄏㄨㄣˋ':'混',
      'ㄏㄨㄤ':'荒','ㄏㄨㄤˊ':'黃','ㄏㄨㄤˇ':'謊','ㄏㄨㄤˋ':'晃',
      'ㄏㄨㄥ':'烘','ㄏㄨㄥˊ':'紅','ㄏㄨㄥˋ':'鬨',
      'ㄐㄧ':'機','ㄐㄧˊ':'及','ㄐㄧˇ':'幾','ㄐㄧˋ':'記',
      'ㄐㄧㄚ':'家','ㄐㄧㄚˊ':'夾','ㄐㄧㄚˇ':'假','ㄐㄧㄚˋ':'價',
      'ㄐㄧㄝ':'街','ㄐㄧㄝˊ':'節','ㄐㄧㄝˇ':'解','ㄐㄧㄝˋ':'借',
      'ㄐㄧㄠ':'交','ㄐㄧㄠˊ':'腳','ㄐㄧㄠˇ':'角','ㄐㄧㄠˋ':'叫',
      'ㄐㄧㄡ':'九','ㄐㄧㄡˊ':'就','ㄐㄧㄡˇ':'酒','ㄐㄧㄡˋ':'就',
      'ㄐㄧㄢ':'間','ㄐㄧㄢˊ':'件','ㄐㄧㄢˇ':'揀','ㄐㄧㄢˋ':'建',
      'ㄐㄧㄣ':'今','ㄐㄧㄣˊ':'緊','ㄐㄧㄣˇ':'緊','ㄐㄧㄣˋ':'進',
      'ㄐㄧㄤ':'江','ㄐㄧㄤˊ':'薑','ㄐㄧㄤˇ':'講','ㄐㄧㄤˋ':'將',
      'ㄐㄧㄥ':'京','ㄐㄧㄥˊ':'晶','ㄐㄧㄥˇ':'井','ㄐㄧㄥˋ':'靜',
      'ㄐㄩ':'居','ㄐㄩˊ':'局','ㄐㄩˇ':'舉','ㄐㄩˋ':'句',
      'ㄐㄩㄝ':'決','ㄐㄩㄝˊ':'覺','ㄐㄩㄝˋ':'絕',
      'ㄐㄩㄢ':'卷','ㄐㄩㄢˇ':'捲','ㄐㄩㄢˋ':'眷',
      'ㄐㄩㄣ':'君','ㄐㄩㄣˋ':'俊',
      'ㄑㄧ':'七','ㄑㄧˊ':'其','ㄑㄧˇ':'起','ㄑㄧˋ':'氣',
      'ㄑㄧㄚ':'掐',
      'ㄑㄧㄝ':'切','ㄑㄧㄝˊ':'茄','ㄑㄧㄝˋ':'竊',
      'ㄑㄧㄠ':'橋','ㄑㄧㄠˊ':'喬','ㄑㄧㄠˇ':'巧','ㄑㄧㄠˋ':'翹',
      'ㄑㄧㄡ':'秋','ㄑㄧㄡˊ':'球',
      'ㄑㄧㄢ':'千','ㄑㄧㄢˊ':'前','ㄑㄧㄢˇ':'淺','ㄑㄧㄢˋ':'欠',
      'ㄑㄧㄣ':'親','ㄑㄧㄣˊ':'勤','ㄑㄧㄣˇ':'寢',
      'ㄑㄧㄤ':'搶','ㄑㄧㄤˊ':'強','ㄑㄧㄤˋ':'嗆',
      'ㄑㄧㄥ':'清','ㄑㄧㄥˊ':'晴','ㄑㄧㄥˇ':'請','ㄑㄧㄥˋ':'慶',
      'ㄑㄩ':'區','ㄑㄩˊ':'曲','ㄑㄩˇ':'取','ㄑㄩˋ':'去',
      'ㄑㄩㄝ':'缺','ㄑㄩㄝˊ':'卻','ㄑㄩㄝˋ':'卻',
      'ㄑㄩㄢ':'全','ㄑㄩㄢˊ':'泉','ㄑㄩㄢˇ':'犬','ㄑㄩㄢˋ':'勸',
      'ㄑㄩㄣ':'群','ㄑㄩㄣˊ':'群',
      'ㄒㄧ':'西','ㄒㄧˊ':'習','ㄒㄧˇ':'洗','ㄒㄧˋ':'系',
      'ㄒㄧㄚ':'蝦','ㄒㄧㄚˊ':'霞','ㄒㄧㄚˋ':'下',
      'ㄒㄧㄝ':'些','ㄒㄧㄝˊ':'鞋','ㄒㄧㄝˇ':'寫','ㄒㄧㄝˋ':'謝',
      'ㄒㄧㄠ':'消','ㄒㄧㄠˊ':'小','ㄒㄧㄠˇ':'小','ㄒㄧㄠˋ':'笑',
      'ㄒㄧㄡ':'修','ㄒㄧㄡˊ':'袖','ㄒㄧㄡˋ':'秀',
      'ㄒㄧㄢ':'先','ㄒㄧㄢˊ':'咸','ㄒㄧㄢˇ':'顯','ㄒㄧㄢˋ':'現',
      'ㄒㄧㄣ':'心','ㄒㄧㄣˊ':'新','ㄒㄧㄣˋ':'信',
      'ㄒㄧㄤ':'香','ㄒㄧㄤˊ':'詳','ㄒㄧㄤˇ':'想','ㄒㄧㄤˋ':'像',
      'ㄒㄧㄥ':'星','ㄒㄧㄥˊ':'行','ㄒㄧㄥˇ':'醒','ㄒㄧㄥˋ':'性',
      'ㄒㄩ':'虛','ㄒㄩˊ':'徐','ㄒㄩˇ':'許','ㄒㄩˋ':'需',
      'ㄒㄩㄝ':'雪','ㄒㄩㄝˊ':'學','ㄒㄩㄝˋ':'血',
      'ㄒㄩㄢ':'宣','ㄒㄩㄢˊ':'旋','ㄒㄩㄢˇ':'選','ㄒㄩㄢˋ':'炫',
      'ㄒㄩㄣ':'薰','ㄒㄩㄣˊ':'尋','ㄒㄩㄣˋ':'訓',
      'ㄓ':'知','ㄓˊ':'直','ㄓˇ':'只','ㄓˋ':'志',
      'ㄓㄚ':'扎','ㄓㄚˊ':'炸','ㄓㄚˋ':'炸',
      'ㄓㄜ':'這','ㄓㄜˊ':'折','ㄓㄜˇ':'者','ㄓㄜˋ':'這','ㄓㄜ˙':'著',
      'ㄓㄞ':'摘','ㄓㄞˋ':'債',
      'ㄓㄠ':'找','ㄓㄠˊ':'著','ㄓㄠˇ':'找','ㄓㄠˋ':'照',
      'ㄓㄡ':'周','ㄓㄡˇ':'肘','ㄓㄡˋ':'晝',
      'ㄓㄢ':'展','ㄓㄢˇ':'展','ㄓㄢˋ':'戰',
      'ㄓㄣ':'真','ㄓㄣˇ':'枕','ㄓㄣˋ':'振',
      'ㄓㄤ':'張','ㄓㄤˇ':'掌','ㄓㄤˋ':'漲',
      'ㄓㄥ':'蒸','ㄓㄥˊ':'政','ㄓㄥˇ':'整','ㄓㄥˋ':'正',
      'ㄓㄨ':'主','ㄓㄨˊ':'竹','ㄓㄨˇ':'煮','ㄓㄨˋ':'住',
      'ㄓㄨㄚ':'抓',
      'ㄓㄨㄟ':'追','ㄓㄨㄟˋ':'醉',
      'ㄓㄨㄢ':'專','ㄓㄨㄢˊ':'轉','ㄓㄨㄢˇ':'轉','ㄓㄨㄢˋ':'賺',
      'ㄓㄨㄤ':'裝','ㄓㄨㄤˊ':'莊','ㄓㄨㄤˋ':'撞',
      'ㄓㄨㄥ':'中','ㄓㄨㄥˊ':'重','ㄓㄨㄥˇ':'腫','ㄓㄨㄥˋ':'種',
      'ㄔ':'吃','ㄔˊ':'遲','ㄔˇ':'齒','ㄔˋ':'赤',
      'ㄔㄚ':'插','ㄔㄚˊ':'查','ㄔㄚˋ':'差',
      'ㄔㄜ':'車','ㄔㄜˊ':'扯',
      'ㄔㄞ':'拆','ㄔㄞˊ':'柴',
      'ㄔㄠ':'抄','ㄔㄠˊ':'朝','ㄔㄠˇ':'炒',
      'ㄔㄡ':'抽','ㄔㄡˊ':'愁','ㄔㄡˇ':'醜','ㄔㄡˋ':'臭',
      'ㄔㄢ':'摻','ㄔㄢˊ':'纏','ㄔㄢˇ':'產','ㄔㄢˋ':'顫',
      'ㄔㄣ':'沈','ㄔㄣˊ':'陳',
      'ㄔㄤ':'昌','ㄔㄤˊ':'長','ㄔㄤˇ':'場','ㄔㄤˋ':'唱',
      'ㄔㄥ':'稱','ㄔㄥˊ':'程','ㄔㄥˇ':'逞',
      'ㄔㄨ':'出','ㄔㄨˊ':'除','ㄔㄨˇ':'楚','ㄔㄨˋ':'處',
      'ㄔㄨㄟ':'吹','ㄔㄨㄟˊ':'垂',
      'ㄔㄨㄢ':'穿','ㄔㄨㄢˊ':'傳','ㄔㄨㄢˇ':'喘','ㄔㄨㄢˋ':'串',
      'ㄔㄨㄤ':'窗','ㄔㄨㄤˊ':'床','ㄔㄨㄤˇ':'闖','ㄔㄨㄤˋ':'創',
      'ㄔㄨㄥ':'衝','ㄔㄨㄥˊ':'蟲','ㄔㄨㄥˇ':'寵',
      'ㄕ':'師','ㄕˊ':'時','ㄕˇ':'使','ㄕˋ':'是',
      'ㄕㄚ':'沙','ㄕㄚˇ':'傻','ㄕㄚˋ':'殺',
      'ㄕㄜ':'奢','ㄕㄜˊ':'蛇','ㄕㄜˋ':'社',
      'ㄕㄞ':'篩','ㄕㄞˋ':'曬',
      'ㄕㄠ':'燒','ㄕㄠˇ':'少','ㄕㄠˋ':'哨',
      'ㄕㄡ':'收','ㄕㄡˊ':'熟','ㄕㄡˇ':'手','ㄕㄡˋ':'受',
      'ㄕㄢ':'山','ㄕㄢˇ':'閃','ㄕㄢˋ':'善',
      'ㄕㄣ':'深','ㄕㄣˊ':'神','ㄕㄣˇ':'審','ㄕㄣˋ':'慎',
      'ㄕㄤ':'商','ㄕㄤˊ':'上','ㄕㄤˇ':'賞','ㄕㄤˋ':'上',
      'ㄕㄥ':'聲','ㄕㄥˊ':'繩','ㄕㄥˇ':'省','ㄕㄥˋ':'勝',
      'ㄕㄨ':'書','ㄕㄨˊ':'熟','ㄕㄨˇ':'鼠','ㄕㄨˋ':'樹',
      'ㄕㄨㄞ':'摔','ㄕㄨㄞˊ':'帥',
      'ㄕㄨㄟ':'水','ㄕㄨㄟˋ':'睡',
      'ㄕㄨㄢ':'拴','ㄕㄨㄢˋ':'涮',
      'ㄕㄨㄤ':'雙','ㄕㄨㄤˇ':'爽',
      'ㄖ':'日','ㄖˊ':'熱',
      'ㄖㄜˋ':'熱','ㄖㄜˇ':'惹',
      'ㄖㄠ':'繞','ㄖㄠˊ':'饒','ㄖㄠˋ':'繞',
      'ㄖㄡ':'揉','ㄖㄡˊ':'柔','ㄖㄡˋ':'肉',
      'ㄖㄢ':'燃','ㄖㄢˊ':'然','ㄖㄢˇ':'染',
      'ㄖㄣ':'人','ㄖㄣˊ':'仁','ㄖㄣˇ':'忍','ㄖㄣˋ':'任',
      'ㄖㄤ':'讓','ㄖㄤˋ':'讓',
      'ㄖㄥ':'扔','ㄖㄥˊ':'仍',
      'ㄖㄨ':'入','ㄖㄨˊ':'如','ㄖㄨˇ':'乳','ㄖㄨˋ':'入',
      'ㄖㄨㄢ':'軟','ㄖㄨㄢˊ':'軟',
      'ㄖㄨㄥ':'容','ㄖㄨㄥˊ':'融','ㄖㄨㄥˋ':'榮',
      'ㄖㄜ':'熱',
      'ㄗ':'資','ㄗˊ':'字','ㄗˇ':'字','ㄗˋ':'自',
      'ㄗㄚ':'雜','ㄗㄚˊ':'雜','ㄗㄚˋ':'砸',
      'ㄗㄜ':'則','ㄗㄜˊ':'澤','ㄗㄜˋ':'責',
      'ㄗㄞ':'在','ㄗㄞˊ':'才','ㄗㄞˇ':'宰','ㄗㄞˋ':'在','ㄗㄞ˙':'哉',
      'ㄗㄠ':'早','ㄗㄠˊ':'糟','ㄗㄠˇ':'早','ㄗㄠˋ':'造',
      'ㄗㄡ':'走','ㄗㄡˇ':'走','ㄗㄡˋ':'奏',
      'ㄗㄢ':'咱','ㄗㄢˊ':'暫','ㄗㄢˋ':'贊',
      'ㄗㄣ':'怎','ㄗㄣˇ':'怎',
      'ㄗㄤ':'髒','ㄗㄤˊ':'藏','ㄗㄤˋ':'葬',
      'ㄗㄥ':'增','ㄗㄥˊ':'曾',
      'ㄗㄨ':'租','ㄗㄨˊ':'族','ㄗㄨˇ':'祖','ㄗㄨˋ':'阻',
      'ㄗㄨㄛ':'作','ㄗㄨㄛˊ':'坐','ㄗㄨㄛˇ':'左','ㄗㄨㄛˋ':'做',
      'ㄗㄨㄟ':'最','ㄗㄨㄟˋ':'最',
      'ㄗㄨㄥ':'總','ㄗㄨㄥˊ':'從','ㄗㄨㄥˇ':'總',
      'ㄘ':'次','ㄘˊ':'詞','ㄘˇ':'此','ㄘˋ':'次',
      'ㄘㄚ':'擦',
      'ㄘㄜ':'測','ㄘㄜˋ':'冊',
      'ㄘㄠ':'操','ㄘㄠˊ':'曹','ㄘㄠˇ':'草',
      'ㄘㄡ':'湊',
      'ㄘㄢ':'參','ㄘㄢˊ':'殘','ㄘㄢˇ':'慘','ㄘㄢˋ':'燦',
      'ㄘㄤ':'倉','ㄘㄤˊ':'藏',
      'ㄘㄥ':'層','ㄘㄥˊ':'曾',
      'ㄘㄨ':'粗','ㄘㄨˋ':'促',
      'ㄘㄨㄛ':'搓','ㄘㄨㄛˋ':'錯',
      'ㄘㄨㄟ':'催','ㄘㄨㄟˋ':'翠',
      'ㄘㄨㄣ':'村','ㄘㄨㄣˊ':'存','ㄘㄨㄣˋ':'寸',
      'ㄘㄨㄥ':'聰','ㄘㄨㄥˊ':'從',
      'ㄙ':'絲','ㄙˊ':'寺','ㄙˇ':'死','ㄙˋ':'四',
      'ㄙㄚ':'灑',
      'ㄙㄜ':'色','ㄙㄜˋ':'澀',
      'ㄙㄞ':'塞','ㄙㄞˋ':'賽',
      'ㄙㄠ':'搔','ㄙㄠˋ':'掃',
      'ㄙㄡ':'搜',
      'ㄙㄢ':'三','ㄙㄢˋ':'散',
      'ㄙㄤ':'桑','ㄙㄤˊ':'喪',
      'ㄙㄥ':'僧',
      'ㄙㄨ':'蘇','ㄙㄨˊ':'俗','ㄙㄨˋ':'素',
      'ㄙㄨㄛ':'所','ㄙㄨㄛˋ':'索',
      'ㄙㄨㄟ':'雖','ㄙㄨㄟˊ':'隨','ㄙㄨㄟˋ':'歲',
      'ㄙㄨㄣ':'孫','ㄙㄨㄣˊ':'損',
      'ㄙㄨㄥ':'鬆','ㄙㄨㄥˋ':'送'
    };
    function parseBpmfSyllables(zhuyinStr){
      const INIT=new Set('ㄅㄆㄇㄈㄉㄊㄋㄌㄍㄎㄏㄐㄑㄒㄓㄔㄕㄖㄗㄘㄙ');
      const MED =new Set('ㄧㄨㄩ');
      const FIN =new Set('ㄚㄛㄜㄝㄞㄟㄠㄡㄢㄣㄤㄥㄦ');
      const TONE=new Set('ˊˇˋ˙');
      const syls=[]; let i='',m='',f='',t='';
      const flush=()=>{const s=i+m+f+t;if(s)syls.push(s);i=m=f=t='';};
      for(const sym of zhuyinStr){
        if(TONE.has(sym)){t=sym;flush();}
        else if(INIT.has(sym)){if(i||m||f)flush();i=sym;}
        else if(MED.has(sym)){if(m||f)flush();m=sym;}
        else if(FIN.has(sym)){if(f)flush();f=sym;}
        else{flush();}  // space/unknown → end current syllable (implicit 1st tone)
      }
      flush();
      return syls;
    }
    function bpmfConvert(raw){
      // ASCII comma=ㄝ in 大千式; use full-width ，、 to separate search terms
      return raw.split(/([，、]+)/).map(part=>{
        if(/[，、]/.test(part)) return part;
        const zhuyin=[...part].map(c=>BPMF_KEYS[c.toLowerCase()]||c).join('');
        return parseBpmfSyllables(zhuyin).map(s=>BPMF_CHARS[s]||s).join('');
      }).join('');
    }
    let bpmfMode=false;
    // ────────────────────────────────────────────────────────────────────
    const input=document.getElementById("words-input"),btnSearch=document.getElementById("btn-search"),
          btnExport=document.getElementById("btn-export"),spinner=document.getElementById("spinner"),
          statusTxt=document.getElementById("status-text"),tableWrap=document.getElementById("table-wrap"),
          tbody=document.getElementById("tbody");
    const btnBpmf=document.getElementById("btn-bpmf"),bpmfHint=document.getElementById("bpmf-hint");
    const progressWrap=document.getElementById("progress-wrap"),progressFill=document.getElementById("progress-fill"),
          progressPct=document.getElementById("progress-pct"),
          jokeText=document.getElementById("joke-text"),segmentLine=document.getElementById("segment-line"),
          loadingPet=document.getElementById("loading-pet");
    let allResults=[],es=null,searchGen=0,currentPage="lookup";
    // ── 查詢中的短短幹話（隨機輪播）─────────────────────────────────
    // 查詢表格頁、近義詞替換頁、筆順動畫頁各 100 句，互不共用
    const JOKES={
      lookup:[
        "別急，CPU也是要喘口氣的。","查字典比查感情史快多了。","進度條走得比感情還慢。",
        "資料庫說它已經很努力了。","不要catch我，我在try。",
        "系統裝忙中。","再等一下，馬上就好，大概吧。","別看我，我只是個進度條。",
        "這個字很難，讓我們思考一下人生。","查詢中，禁止吃瓜。","程式碼跑得比你早起的速度還快，放心。",
        "萌典正在偷偷打字中。","開發人員隱藏了一些小驚喜，到處點點吧！",
        "窗外有人","這個工具的起源只是想要加快寫作業的速度。",
        "九綱、偏光、烏與聲明、表裏之間","系統正在努力假裝有進度。","你的查詢已上路，目前塞車中。",
        "如果想要加入更多小驚喜，歡迎聯絡開發人員。","字海茫茫，撈一下就回來。","伺服器目前已讀不回，正在努力說服它運作......",
        "先別關掉，奇蹟正在載入。","誠摯感謝你的使用","詞庫很大，我先找一下路。",
        "搜尋結果正在從索引裡爬出來。","API 已出發，還在等紅綠燈。","這個詞有點害羞，不肯出現在結果裡。",
        "正在翻遍字典的每一個角落。","系統在跑了","關鍵字已收到，答案還在趕路。",
        "資料庫正在回想它把資料放哪了。","搜尋範圍太大，正在縮小人生。",
        "結果快到了，先不要重新整理。","位相、黃昏、智慧之瞳","正在從千萬個字裡認出你要的那個。",
        "索引翻得太快，不小心翻過頭了。","你的關鍵字正在接受身分驗證。","系統正在確認這不是錯別字。",
        "伺服器正在用盡全力保持淡定。","正在聯絡詞彙學教授......",
        "資料已經找到，現在找回來的路。","正在把句子拆開來逐一盤問。","查詢送出去了，希望它記得回來。",
        "系統正在和資料庫進行眼神交流。","每個詞都說自己是關鍵字。","正在排除那些看起來很像的答案。",
        "字典太厚，伺服器翻得有點喘。","結果正在解壓縮，請勿施壓。","查詢引擎正在熱身，等等就起跑。",
        "正在確認這個詞有沒有雙重身分。","資料庫正在努力回憶，很久以前的事。",
        "你的查詢已插隊失敗，乖乖排隊中。","文字分析中，標點符號請先退場。",
        "正在替每個字安排適合的位置。","找詞像找鑰匙，通常就在最明顯的地方。","系統讀懂了一半，另一半正在裝懂。",
        "查詢結果正在穿鞋，馬上出門。","語料庫很深，正在放繩子下去撈。","這個詞跑進冷門區了，我去追。",
        "正在比對詞義，請避免突然改口。","搜尋雷達已啟動，目前偵測到文字。","答案正在排版，不能衣衫不整地見你。",
        "分詞模型正在舉手表決。",
        "系統正在將耐心轉換成搜尋結果。","斷詞器卡在一段複雜的關係裡。","字典正在開會，討論該派誰回答。",
        "搜尋進度正常，正常地有一點慢。","你的關鍵字已成功引起系統注意。","正在從文字堆裡挖掘知識化石。","資料庫翻了個身，繼續努力。","正在建立索引，也建立彼此的信任。",
        "查詢快完成了，這句不是安慰。","詞義正在對焦，請保持畫面穩定。","系統已經知道答案，只差想起來。",
        "找到了幾個結果，正在確認沒有冒牌貨。","字詞分析進入深水區，請勿跳船。","最後一個資料表正在慢慢走來。",
        "查詢完成前，先假裝一切都很順利。"
      ],
      rewrite:[
        "這個詞躲得很好，再找一下。","查詢跑很快，只是終點有點遠。","系統沒有卡，是時間變慢了。",
        "正在把資料從宇宙另一端搬來。","斷詞進行中，請勿打斷。","字典翻到一半，突然忘記要找什麼。",
        "別催，再催就顯示「查無資料」。","正在召喚失蹤的詞彙。","資料有來，只是走得比較優雅。",
        "系統正在思考這個詞值不值得查。","查詢很順利，除了還沒有結果。","進度條已經盡力表演了。",
        "稍候，資料正在排隊進場。","正在努力理解你想表達什麼。",
        "有些詞一轉身，就是一輩子的載入。","系統忙著找詞，沒空解釋。",
        "快了快了，這次可能是真的。","你的詞正在跟伺服器玩躲貓貓。","正在整理文字，順便整理心情。",
        "資料不是不來，只是不想面對。","找不到答案時，先怪網路。","每一次載入，都是對耐心的測驗。",
        "請保持冷靜，我們的辭典比你更慌。","正在幫這個詞換一套說法。","原句很好，只是想換個髮型。",
        "近義詞正在後台排隊試鏡。","正在找一個意思相近、脾氣更好的詞。","換句話說之前，先想想要怎麼說。",
        "語氣正在微調，請勿大聲喧嘩。","這個詞沒有錯，只是有點用膩了。","正在替句子物色新搭檔。",
        "同一個意思，正在尋找不同的人生。","文字改造中，原意保留、包裝更新。","正在避免把近義詞改成陌生人。",
        "替換候選太多，正在舉辦海選。","句子正在試穿另一種語氣。","這個詞想休假，正在找代班。",
        "正在找一個不搶戲但很有用的詞。","原意正在現場監督，放心。","文字換位中，意思請留在原地。",
        "正在替平凡的句子加一點戲。","近義詞很多，合得來的沒幾個。",
        "正在確認替代詞沒有偷偷改變立場。","句子需要新鮮感，但不需要新劇情。","語意不能跑，文字可以換。",
        "正在把重複用字請出會議室。","這個詞太常加班，換別人上場。",
        "正在調整措辭，不調整你的本意。","句子正在重新穿搭，風格稍候。","想法不變，只幫它換一種表情。",
        "正在把生硬的詞按摩得自然一點。","這句話有點直，正在幫它轉彎。","改寫不是變心，只是換個說法。",
        "正在確認新詞能不能融入這個家庭。",
        "句子已經不錯，現在追求更不錯。","替換詞正在做最後的語意體檢。",
        "正在避免優雅地表達錯誤意思。","文字看似相同，個性其實差很多。","正在挑一個讀起來不會跌倒的版本。","這個詞可以換，但感情不能淡。","改寫引擎正在避免畫蛇添足。",
        "正在刪除多餘的字，以及多餘的焦慮。","句型正在重組，意思沒有受傷。","正在找一句「就是這個感覺」的版本。",
        "候選詞都說自己最適合，正在查證。","這個詞語氣不對，先請它坐回去。","正在把你的話說得像你真的想過。",
        "文字精修中，每個字都要交代理由。","正在替句子增加一點呼吸空間。","新版本快好了，原句先不要吃醋。",
        "近義詞不是複製人，也有自己的脾氣。","正在處理「意思對，但感覺怪」的問題。",
        "正在讓句子自然到像沒改過。","正在將普通用字升級成稍微不普通。",
        "語意已鎖定，文字正在自由發揮。","最佳替代詞即將從候選名單中勝出。",
        "改寫完成前，請先相信文字的可能性。"
      ],
      handwriting:[
        "一筆一劃，急不來的。","筆順錯了，老師會哭。","正在練字，請勿搖桌子。",
        "先橫後豎，先別催我。","毛筆沾墨中，稍候。","這一撇，寫了三輩子。",
        "寫字比打字有誠意多了。","字如其人，所以寫慢一點。","正在描紅，不要偷看。",
        "每一筆都是手工限定款。","點橫豎撇捺，一個都不能少。","寫到一半，墨水說它累了。",
        "這個字筆畫有點多，給它點時間。","書法家正在找靈感。","寫錯不能按 Ctrl+Z 的年代。",
        "筆順正確，心情才會正確。","正在把楷書寫得很楷。","字寫得好看，要慢慢來。",
        "由上而下、由左而右、由你等待。","筆還在紙上，魂已經飛了。","這筆捺拉得有點長，請稍候。",
        "先外後內再封口，好了再叫你。","練字十年，只為這一刻。","寫字中，禁止打翻墨汁。",
        "起筆很重要，所以先讓它深呼吸。","收筆也很重要，不能下班得太隨便。",
        "正在從第一筆開始，沒有偷吃步。","橫要平，心可以不用。","豎要直，人生不一定。",
        "撇出去容易，捺回來很難。","這一點雖小，少了整個字都不對。","筆尖已就位，紙張正在做心理準備。",
        "正在確認這一筆該彎還是該放下。","墨還沒乾，請不要急著翻頁。","字正在長大，一筆就是一歲。",
        "寫快叫簽名，寫慢才叫筆順。","每一筆都有方向，除了人生。",
        "正在幫筆畫排好出場順序。","偏旁先站好，部首等等要點名。",
        "筆尖轉彎中，請繫好安全帶。","正在把方塊字蓋得四四方方。","字還沒寫完，先別猜它是誰。",
        "一筆落下，撤回鍵正式失效。","正在示範什麼叫做筆筆有交代。","這個鉤有點難鉤，等等再收網。",
        "筆畫正在集合，尚有一撇未到。","寫字不能瞬間移動，只能慢慢走。","正在計算這一折到底要折幾度。",
        "橫折彎鉤已進入高難度路段。","字帖說慢一點，手說它盡力了。","正在把每一筆送到正確的位置。",
        "筆鋒正在轉向，方向燈已經打了。","這一豎很有原則，說直就直。","正在處理一筆看似簡單的長橫。",
        "寫歪沒關係，動畫會假裝沒看見。","這一撇出去，可能就不回來了。","正在收拾剛剛那個豪邁的捺。",
        "筆尖走過的路，都是標準答案。",
        "字的骨架已完成，細節正在入住。","筆順動畫正在逐格證明自己沒寫錯。",
        "這個部首戲份很多，請給它一點時間。","正在處理藏在角落裡的最後一點。","筆畫太多，這個字可能有片尾名單。",
        "墨跡正在伸展，稍後形成一個字。","先撇後捺，左右兩邊都要照顧。","這個框還沒封口，字正在通風。",
        "正在從外面寫進去，再把門關好。","筆尖今天很忙，一刻都沒離開紙面。","這一筆需要氣勢，正在醞釀。",
        "寫字是門藝術，載入也是。","正在讓點、橫、豎各就各位。","收筆收得漂亮，才算體面地結束。","正在演示筆尖如何優雅地急轉彎。",
        "筆畫正在紙上進行接力賽。",
        "正在對齊重心，免得這個字站不穩。","字形施工中，請戴好安全帽。","第一筆已落下，現在沒有回頭路了。",
        "筆尖正在按照祖傳規則前進。","正在完成那一筆畫龍點睛的點。",
        "手腕已轉彎，動畫很快就跟上。","字已完成九成，最後一筆正在耍帥。",
        "正在讓筆畫首尾呼應......","紙上路線已規劃，筆尖準備導航。","最後一筆即將落下，請屏住呼吸。",
        "字寫完了嗎？等墨乾了才算。"
      ]
    };
    // 洗牌輪播：一輪 100 句全部講完才會重複，也不會連續兩次同一句
    let jokeTimer=null,jokeBag=[],jokeLast="";
    const nextJoke=()=>{
      if(!jokeBag.length){
        jokeBag=[...JOKES[currentPage]].sort(()=>Math.random()-.5);
        if(jokeBag[jokeBag.length-1]===jokeLast) jokeBag.unshift(jokeBag.pop());
      }
      jokeLast=jokeBag.pop();jokeText.textContent=jokeLast;
    };
    const startJokes=()=>{
      jokeBag=[];nextJoke();
      clearInterval(jokeTimer);
      jokeTimer=setInterval(nextJoke,2200);
    };
    const stopJokes=()=>{clearInterval(jokeTimer);jokeTimer=null;jokeText.textContent="";};
    // ── 吉祥物：只在查詢/斷詞真正跑的時候出現，跟進度條同開同關 ──
    // 查詞頁是粉色兔子，近義詞替換頁是豹紋粉髮妹（IMG_2519.GIF 去背後的 23 格）
    const MASCOTS={
      lookup:Array.from({length:8},(_,i)=>`/mascot/bunny_typing_0${i}.png`),
      rewrite:Array.from({length:23},(_,i)=>`/mascot/cheetah_${String(i).padStart(2,"0")}.png`)
    };
    MASCOTS.handwriting=Array.from({length:8},(_,i)=>`/mascot/writer_${String(i).padStart(2,"0")}.png`);
    const MASCOT_MS={lookup:90,rewrite:90,handwriting:40};  // 各自照原 GIF 的節奏
    let MASCOT_FRAMES=MASCOTS.lookup;
    let mascotTimer=null,mascotIdx=0;
    const startMascot=()=>{
      MASCOT_FRAMES=MASCOTS[currentPage];
      mascotIdx=0;loadingPet.src=MASCOT_FRAMES[0];
      clearInterval(mascotTimer);
      mascotTimer=setInterval(()=>{
        mascotIdx=(mascotIdx+1)%MASCOT_FRAMES.length;
        loadingPet.src=MASCOT_FRAMES[mascotIdx];
      },MASCOT_MS[currentPage]);
    };
    const stopMascot=()=>{clearInterval(mascotTimer);mascotTimer=null;};
    // 累積式進度：一整個操作（斷詞→查核→配對→定義確認→查詢）只往前走，
    // 不會因為換到下一個階段就掉回 0%。stagePct 把「這個階段自己的
    // done/total」換算成整條進度條裡佔的那一段（offset~offset+span）；
    // bumpProgress 用 Math.max 確保畫面上的百分比只增不減。
    let progressMax=0;
    const resetProgress=()=>{progressMax=0;progressFill.style.width="0%";progressPct.textContent="0%";};
    const bumpProgress=pct=>{
      progressMax=Math.max(progressMax,Math.min(100,Math.max(0,pct)));
      progressFill.style.width=progressMax+"%";
      progressPct.textContent=Math.round(progressMax)+"%";
    };
    const stagePct=(done,total,offset,span)=>offset+(total?Math.min(1,done/total):0)*span;
    // 查詢類按鈕送出中：只變灰、不設 disabled——disabled 的按鈕瀏覽器直接吞掉點擊，
    // 就數不到「連按三次」；重複送出改由 querySubmit 檢查 busy 擋掉
    const setBusy=(btn,on)=>{btn.classList.toggle("busy",on);btn.setAttribute("aria-disabled",on?"true":"false");};
    const isBusy=btn=>btn.classList.contains("busy");
    // 查詢類按鈕／輸入框 Enter 共用：送出中再按不重送（防呆）；1.5 秒內連按 3 次跳「你的性子也太急了」。
    // Enter 只算真正按下的那一次：按住不放的自動重複、注音輸入法選字中的 Enter 都不算、也不送出。
    // 另外送出後 1.5 秒內同一個按鈕不再送：查詢很快（例如查過的詞有快取）時，
    // 第一次早就跑完、按鈕已不在忙碌中，光看 busy 擋不住連按造成的重複送出。
    const RAPID_MS=1500,rapidHits={},lastSent={};
    function querySubmit(key,btn,run){
      return e=>{
        if(e&&e.type==="keydown"&&(e.key!=="Enter"||e.repeat||e.isComposing||e.keyCode===229))return;
        const now=Date.now(),hits=(rapidHits[key]||[]).filter(t=>now-t<RAPID_MS);
        hits.push(now);rapidHits[key]=hits;
        if(hits.length>=3){rapidHits[key]=[];showImpatient();}
        if(isBusy(btn)||now-(lastSent[key]||0)<RAPID_MS)return;
        lastSent[key]=now;
        run();
      };
    }
    const setLoading=on=>{
      spinner.style.display=on?"block":"none";setBusy(btnSearch,on);
      excludeInput.disabled=on;
      progressWrap.style.display=on?"flex":"none";
      if(on){startJokes();startMascot();} else {stopJokes();stopMascot();}
    };
    const setStatus=msg=>{statusTxt.textContent=msg;};
    // 由低到高排等級順序，"—"（沒有等級資料）永遠排最後
    const LEVEL_ORDER=["基礎第1級","基礎第1*級","基礎第2級","基礎第2*級","基礎第3級","基礎第3*級",
                        "進階第4級","進階第4*級","進階第5級","精熟第6級","精熟第7級"];
    const LEVEL_RANK=new Map(LEVEL_ORDER.map((lv,i)=>[lv,i]));
    const levelRank=lv=>LEVEL_RANK.has(lv)?LEVEL_RANK.get(lv):-1;
    function sortByLevelDesc(){
      // 查無資料的排最上面（優先讓使用者看到哪些詞完全沒查到），
      // 其餘照單字級數降冪排序，沒等級資料的（"—"）排最後。
      allResults.sort(byLevelDesc);
      lookupView.set(allResults);
    }
    function byLevelDesc(a,b){
      const aNone=a.pinyin==="查無資料"?1:0,bNone=b.pinyin==="查無資料"?1:0;
      if(aNone!==bNone) return bNone-aNone;
      return levelRank(b.level)-levelRank(a.level);
    }
    function lvTag(lv){
      if(!lv||lv==="—") return "—";
      const cls=lv.startsWith("基礎")?"基礎":lv.startsWith("進階")?"進階":lv.startsWith("精熟")?"精熟":"";
      return cls?`<span class="lv-tag lv-${cls}">${lv}</span>`:lv;
    }
    function rowHtml(entry){
      if(entry.pinyin==="查無資料"||entry.pinyin==="錯誤")
        return `<td class="word">${entry.word}</td><td class="error" colspan="3">${entry.pinyin==="錯誤"?"錯誤："+entry.definition:"查無資料"}</td><td class="lv">${lvTag(entry.level)}</td>`;
      const pos=entry.pos!=="—"?`<span class="badge">${entry.pos}</span>`:"—";
      return `<td class="word">${entry.word}</td><td class="pin">${entry.pinyin}</td><td class="pos">${pos}</td><td class="def">${entry.definition}</td><td class="lv">${lvTag(entry.level)}</td>`;
    }
    // ── 顯示用表格（只影響網頁上看到的，下載的 Excel 照樣是全部結果）──
    // 同一個詞有多筆結果時收成一列「查看更多」；上方篩選列可只看某個詞彙等級；
    // chips（斷詞結果那一排）點哪個詞就只顯示那個詞，再點一次取消。
    function makeResultView(tbody,bar,chips){
      const v={results:[],word:null,level:null,open:new Set()};
      const lvKey=lv=>lv&&lv!=="—"?lv:"—";
      // 照排序後第一次出現的順序，把同一個詞的結果收在一起
      const groups=()=>{
        const m=new Map();
        v.results.forEach(e=>{if(!m.has(e.word))m.set(e.word,[]);m.get(e.word).push(e);});
        return [...m.values()];
      };
      function renderBar(all){
        const counts=new Map();
        all.forEach(g=>{const k=lvKey(g[0].level);counts.set(k,(counts.get(k)||0)+1);});
        // 等級由低到高，沒有等級資料的放最後
        const keys=[...counts.keys()].sort((a,b)=>(a==="—")-(b==="—")||levelRank(a)-levelRank(b));
        bar.innerHTML=`<button data-lv="" class="${v.level?"":"on"}">全部顯示（${all.length}）</button>`+
          keys.map(k=>`<button data-lv="${esc(k)}" class="${v.level===k?"on":""}">${k==="—"?"無等級":esc(k)}（${counts.get(k)}）</button>`).join("")+
          (v.word?`<span class="rf-word">只顯示「${esc(v.word)}」<button data-clear-word>✕ 取消</button></span>`:"");
      }
      function render(){
        const all=groups();
        renderBar(all);
        const shown=all.filter(g=>(!v.level||lvKey(g[0].level)===v.level)&&
                                  (!v.word||g.some(e=>e.word===v.word||e.query===v.word)));
        tbody.innerHTML=shown.map(g=>{
          if(g.length===1)return `<tr>${rowHtml(g[0])}</tr>`;
          const key=esc(g[0].word),open=v.open.has(g[0].word);
          const head=`<tr><td class="word">${g[0].word}<button class="more-btn" data-more="${key}">${open?"收起 ▴":`查看更多（${g.length} 筆）▾`}</button></td>`+
                     `<td class="more-sum" colspan="3">${open?"":`這個詞有 ${g.length} 筆結果，點「查看更多」全部列出`}</td><td class="lv">${lvTag(g[0].level)}</td></tr>`;
          return open?head+g.map(e=>`<tr class="more-row">${rowHtml(e)}</tr>`).join(""):head;
        }).join("")||`<tr><td class="rf-empty" colspan="5">沒有符合的結果</td></tr>`;
        if(chips)chips.querySelectorAll(".seg-word").forEach(c=>c.classList.toggle("on",c.dataset.w===v.word));
      }
      bar.addEventListener("click",e=>{
        const b=e.target.closest("button");if(!b)return;
        if(b.hasAttribute("data-clear-word"))v.word=null;else v.level=b.dataset.lv||null;
        render();
      });
      tbody.addEventListener("click",e=>{
        const b=e.target.closest("[data-more]");if(!b)return;
        const w=b.dataset.more;v.open.has(w)?v.open.delete(w):v.open.add(w);
        render();
      });
      if(chips)chips.addEventListener("click",e=>{
        const c=e.target.closest(".seg-word");if(!c)return;
        v.word=v.word===c.dataset.w?null:c.dataset.w;
        render();
      });
      return {set(results){v.results=results;v.word=null;v.level=null;v.open.clear();render();}};
    }
    const lookupView=makeResultView(tbody,document.getElementById("result-filter"),segmentLine);
    btnBpmf.addEventListener("click",()=>{
      bpmfMode=!bpmfMode;
      btnBpmf.classList.toggle("active",bpmfMode);
      if(bpmfMode){
        showMeteor();
        input.placeholder="注音鍵盤：u ek7 = 一個，jau3 = 找";
        bpmfHint.textContent=bpmfConvert(input.value)||"";
      } else {
        input.placeholder="例：給予 讚美、進食；快樂,熱血";
        bpmfHint.textContent="";
      }
    });
    input.addEventListener("input",()=>{
      if(bpmfMode) bpmfHint.textContent=bpmfConvert(input.value)||"";
    });
    async function runWordSearch(raw,range){
      if(!raw)return;
      searchGen++;
      const {offset,span}=range||{offset:0,span:100};
      if(es){es.abort();es=null;}
      tbody.innerHTML="";allResults=[];tableWrap.style.display="none";
      btnExport.style.display="none";setLoading(true);setStatus("查詢中…");
      let totalWords=0;
      // 詞彙清單改用 POST body 傳（不是 EventSource 的 GET query string），
      // 因為長文斷詞出來動輒幾千個詞，塞進 URL 會被 414 拒絕。
      const ctrl=new AbortController();es=ctrl;
      try{
        const res=await fetch("/search",{method:"POST",
          headers:{"Content-Type":"application/x-www-form-urlencoded"},
          body:"words="+encodeURIComponent(raw),signal:ctrl.signal});
        if(!res.ok||!res.body){setStatus("連線錯誤："+res.status);setLoading(false);return;}
        const reader=res.body.getReader(),decoder=new TextDecoder();
        let buf="";
        while(true){
          const {value,done}=await reader.read();
          if(done)break;
          buf+=decoder.decode(value,{stream:true});
          let idx;
          while((idx=buf.indexOf("\n\n"))>=0){
            const line=buf.slice(0,idx);buf=buf.slice(idx+2);
            if(!line.startsWith("data: "))continue;
            const data=JSON.parse(line.slice(6));
            if(data.done){
              // 查詢跑完才把表格整個顯示出來，而不是邊查邊冒出來
              sortByLevelDesc();
              bumpProgress(offset+span);
              tableWrap.style.display="block";
              setLoading(false);setStatus(`完成，共 ${allResults.length} 筆（依單字級數降冪排序）`);
              if(allResults.length)btnExport.style.display="inline-block";
              continue;
            }
            if(data.error){setStatus("錯誤："+data.error);setLoading(false);continue;}
            if(data.total){totalWords=data.total;setStatus(`查詢中… 0 / ${totalWords} 個詞`);continue;}
            allResults.push(data);
            bumpProgress(stagePct(allResults.length,totalWords,offset,span));
            setStatus(totalWords?`查詢中… ${allResults.length} 筆（約 ${totalWords} 個詞）`:`查詢中… ${allResults.length} 筆`);
          }
        }
      }catch(err){
        if(err.name!=="AbortError"){setStatus("連線錯誤，請重試："+err.message);}
        setLoading(false);
      }finally{
        es=null;
      }
    }
    async function doSearch(){
      let raw=input.value.trim();if(!raw){input.focus();return;}
      if(bpmfMode) raw=bpmfConvert(raw);
      segmentLine.style.display="none";segmentLine.innerHTML="";
      resetProgress();
      await runWordSearch(raw,{offset:0,span:100});
    }
    btnSearch.addEventListener("click",querySubmit("search",btnSearch,doSearch));
    input.addEventListener("keydown",querySubmit("search",btnSearch,doSearch));
    btnExport.addEventListener("click",async()=>{
      btnExport.disabled=true;
      try{
        const res=await fetch("/export",{method:"POST",headers:{"Content-Type":"application/json"},
          body:JSON.stringify({results:allResults})});
        if(!res.ok){
          const err=await res.json().catch(()=>({}));
          setStatus("下載失敗："+(err.error||res.status));
          return;
        }
        const blob=await res.blob();
        const url=URL.createObjectURL(blob);
        const a=document.createElement("a");
        a.href=url;a.download="查詢結果.xlsx";
        document.body.appendChild(a);a.click();a.remove();
        URL.revokeObjectURL(url);
        setStatus("已下載 查詢結果.xlsx");
      }catch(err){
        setStatus("下載失敗："+err.message);
      }finally{
        btnExport.disabled=false;
      }
    });

    const pasteText=document.getElementById("paste-text"),
          lookupFile=document.getElementById("lookup-file"),
          excludeInput=document.getElementById("exclude-input"),
          btnSegment=document.getElementById("btn-segment");
    // 選完（或拖進）檔案立刻抽出文字放進文本框，讓使用者看得到、能選取／修改。
    async function loadFileInto(file,textarea,report,fileLabel){
      if(!file)return;
      report(`讀取 ${file.name} 中…`);
      try{
        const fd=new FormData();fd.append("file",file);
        const data=await (await fetch("/extract-text",{method:"POST",body:fd})).json();
        if(data.error){report("檔案讀取失敗："+data.error);return;}
        textarea.value=data.text;
        fileLabel.textContent=`目前檔案：${file.name}（${data.text.length} 字）`;
        fileLabel.title=file.name;fileLabel.classList.add("loaded");
        report(`已載入 ${file.name}（${data.text.length} 字）`);
      }catch(err){
        report("檔案讀取失敗："+err.message);
      }
    }
    const segmentRun=async()=>{
      const text=pasteText.value.trim();
      if(!text){pasteText.focus();setStatus("請貼上文字，或把檔案拖進視窗");return;}
      setBusy(btnSegment,true);setLoading(true);setStatus("斷詞中…（長文會分段處理，請稍候）");
      segmentLine.style.display="none";segmentLine.innerHTML="";
      tbody.innerHTML="";allResults=[];tableWrap.style.display="none";btnExport.style.display="none";
      resetProgress();
      try{
        const fd=new FormData();
        fd.append("text",text);
        fd.append("exclude",excludeInput.value.trim());
        const res=await fetch("/segment",{method:"POST",body:fd});
        if(!res.ok||!res.body){setStatus("斷詞失敗："+res.status);setLoading(false);return;}

        const reader=res.body.getReader(),decoder=new TextDecoder();
        let buf="",words=null,errMsg=null;
        while(true){
          const {value,done}=await reader.read();
          if(done)break;
          buf+=decoder.decode(value,{stream:true});
          let idx;
          while((idx=buf.indexOf("\n\n"))>=0){
            const line=buf.slice(0,idx);buf=buf.slice(idx+2);
            if(!line.startsWith("data: "))continue;
            const data=JSON.parse(line.slice(6));
            if(data.error){errMsg=data.error;}
            else if(data.done){words=data.words;}
            else if(data.progress!==undefined){bumpProgress(stagePct(data.progress,data.total,0,25));setStatus(`斷詞中… 第 ${data.progress}/${data.total} 段，已找到 ${data.words_so_far} 個詞`);}
            else if(data.stage==="word_check"){bumpProgress(stagePct(data.done,data.total,25,10));setStatus(`查核多字詞是否為辭典詞條中… ${data.done}/${data.total}`);}
            else if(data.stage==="pair_check"){bumpProgress(stagePct(data.done,data.total,35,25));setStatus(`第 ${data.round} 輪前後字配對中… ${data.done}/${data.total}`);}
            else if(data.stage==="def_check"){bumpProgress(stagePct(data.done,data.total,60,10));setStatus(`確認剩餘單字是否有定義中… ${data.done}/${data.total}`);}
          }
        }

        if(errMsg){setStatus("斷詞失敗："+errMsg);setLoading(false);return;}
        if(!words){setStatus("斷詞失敗：連線中斷");setLoading(false);return;}
        // 分詞結果獨立顯示一條在表格上面，不塞進上面單字查詢的輸入框
        segmentLine.innerHTML=`<span class="seg-label">斷詞結果（${words.length} 個詞，點詞只看那個詞）</span>`+
          words.map(w=>`<span class="seg-word" data-w="${esc(w)}" title="點一下只顯示這個詞">${esc(w)}</span>`).join("");
        segmentLine.style.display="block";
        setStatus(`斷詞完成，共 ${words.length} 個詞，查詢中…`);
        bumpProgress(70);
        await runWordSearch(words.join("，"),{offset:70,span:30});
      }catch(err){
        setStatus("斷詞失敗："+err.message);setLoading(false);
      }finally{
        setBusy(btnSegment,false);
      }
    };
    btnSegment.addEventListener("click",querySubmit("segment",btnSegment,segmentRun));

    // ── 分頁：查詞表格／近義詞替換 ──────────────────────────────────────
    // 兩頁共用同一條進度條＋吉祥物，切頁時把它搬到該頁的 progress-slot。
    const pages={lookup:document.getElementById("page-lookup"),rewrite:document.getElementById("page-rewrite"),
                 handwriting:document.getElementById("page-handwriting")};
    const progressSlots={lookup:document.getElementById("lookup-progress-slot"),rewrite:document.getElementById("rewrite-progress-slot"),
                         handwriting:document.getElementById("handwriting-progress-slot")};
    function showPage(name){
      currentPage=name;
      Object.entries(pages).forEach(([k,el])=>el.classList.toggle("active",k===name));
      document.querySelectorAll(".tab").forEach(t=>t.classList.toggle("active",t.dataset.page===name));
      progressSlots[name].appendChild(progressWrap);
      if(name==="rewrite"&&!rwText.value.trim()&&pasteText.value.trim()) rwText.value=pasteText.value;
    }
    document.querySelectorAll(".tab").forEach(t=>t.addEventListener("click",()=>showPage(t.dataset.page)));

    // ── 近義詞替換：文本選一段 → 斷詞（保留原順序）→ 點詞換近義詞 → 確認 ──
    // 近義詞來源：教育部《重編國語辭典修訂本》相似詞索引表，全部列出並標等級。
    const rwText=document.getElementById("rw-text"),rwFile=document.getElementById("rw-file"),
          btnRwSeg=document.getElementById("btn-rw-seg"),rwStatus=document.getElementById("rw-status"),
          rwWork=document.getElementById("rw-work"),rwTokensEl=document.getElementById("rw-tokens"),
          rwPanel=document.getElementById("rw-panel"),rwPanelTitle=document.getElementById("rw-panel-title"),
          rwSyns=document.getElementById("rw-syns"),rwCompare=document.getElementById("rw-compare"),
          rwOrig=document.getElementById("rw-orig"),rwNew=document.getElementById("rw-new"),
          rwHistory=document.getElementById("rw-history");
    const esc=t=>String(t).replace(/[&<>"]/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;"}[c]));
    let rwTokens=[],rwActive=-1,rwSynCache={},rwRange=null,rwCount=0;
    const rwChanged=()=>rwTokens.map((t,i)=>({t,i})).filter(({t})=>t.cur!==t.text);
    const rwOrigHtml=()=>rwTokens.map(t=>t.cur!==t.text?`<del>${esc(t.text)}</del>`:esc(t.text)).join("");
    const rwNewHtml=()=>rwTokens.map(t=>t.cur!==t.text?`<mark>${esc(t.cur)}</mark>`:esc(t.cur)).join("");
    function rwRender(){
      rwTokensEl.innerHTML=rwTokens.map((t,i)=>{
        if(!t.word) return `<span class="rw-tok">${esc(t.text)}</span>`;
        const cls=["rw-tok","pick",t.has_syn?"has-syn":"",t.cur!==t.text?"changed":"",i===rwActive?"active":""].join(" ");
        return `<span class="${cls}" data-i="${i}" title="${esc(t.text)}｜${esc(t.level)}">${esc(t.cur)}</span>`;
      }).join("");
      rwOrig.innerHTML=rwOrigHtml();rwNew.innerHTML=rwNewHtml();
      rwCompare.style.display="grid";
    }
    async function rwOpen(i){
      rwActive=i;rwRender();
      const t=rwTokens[i];
      rwPanel.style.display="block";
      rwPanelTitle.innerHTML=`<b>${esc(t.text)}</b>　${lvTag(t.level)}　的近義詞：查詢中…`;
      rwSyns.innerHTML="";
      try{
        const data=rwSynCache[t.text]||(rwSynCache[t.text]=await (await fetch("/synonyms?word="+encodeURIComponent(t.text))).json());
        if(rwActive!==i)return;
        // 全部列出，依詞彙等級由低到高排，沒等級資料的排最後
        const syns=[...data.synonyms].sort((a,b)=>{
          const ra=levelRank(a.level),rb=levelRank(b.level);
          return (ra<0?999:ra)-(rb<0?999:rb);
        });
        rwPanelTitle.innerHTML=`<b>${esc(t.text)}</b>　${lvTag(t.level)}　的近義詞（${syns.length} 個，依等級排序）`;
        const opts=[{word:t.text,level:t.level,orig:true},...syns];
        rwSyns.innerHTML=syns.length?opts.map((o,k)=>
          `<button class="rw-syn${o.word===t.cur?" chosen":""}" data-k="${k}">${esc(o.word)}${o.orig?"（原詞）":""} ${lvTag(o.level)}</button>`).join("")
          :`<span class="rw-hint">教育部相似詞索引表查無此詞的近義詞</span>`;
        rwSyns.querySelectorAll(".rw-syn").forEach(btn=>btn.addEventListener("click",()=>{
          const o=opts[+btn.dataset.k];
          t.cur=o.word;t.curLevel=o.level;
          rwOpen(i);
        }));
      }catch(err){
        rwPanelTitle.textContent="近義詞查詢失敗："+err.message;
      }
    }
    rwTokensEl.addEventListener("click",e=>{
      const el=e.target.closest(".rw-tok.pick");
      if(el) rwOpen(+el.dataset.i);
    });
    const rwReport=msg=>{rwStatus.textContent=msg.startsWith("已載入")?msg+"，用滑鼠選取一段後按「近義詞替換」":msg;};
    // 檔案直接拖到視窗任何地方：放進目前這一頁的文本框
    const dropOverlay=document.getElementById("drop-overlay");
    let dragDepth=0;
    const isFileDrag=e=>[...(e.dataTransfer?.types||[])].includes("Files");
    window.addEventListener("dragenter",e=>{if(!isFileDrag(e))return;e.preventDefault();dragDepth++;dropOverlay.classList.add("show");});
    window.addEventListener("dragover",e=>{if(isFileDrag(e))e.preventDefault();});
    window.addEventListener("dragleave",e=>{if(!isFileDrag(e))return;if(--dragDepth<=0){dragDepth=0;dropOverlay.classList.remove("show");}});
    window.addEventListener("drop",e=>{
      if(!isFileDrag(e))return;
      e.preventDefault();dragDepth=0;dropOverlay.classList.remove("show");
      const file=e.dataTransfer.files[0];
      if(currentPage==="handwriting"){hwStatus.textContent="這一頁不用檔案，直接在上面輸入字就好";return;}
      if(currentPage==="rewrite") loadFileInto(file,rwText,rwReport,rwFile);
      else loadFileInto(file,pasteText,setStatus,lookupFile);
    });
    async function rwSegment(){
      const start=rwText.selectionStart,end=rwText.selectionEnd;
      const text=rwText.value.slice(start,end);
      if(!text.trim()){rwStatus.textContent="請先在文本中用滑鼠選取一段文字";rwText.focus();return;}
      rwRange={start,end,text};
      setBusy(btnRwSeg,true);rwTokens=[];rwActive=-1;
      rwWork.style.display="none";rwPanel.style.display="none";rwCompare.style.display="none";
      resetProgress();setLoading(true);rwStatus.textContent=`斷詞中…（選取 ${text.length} 字）`;
      try{
        const fd=new FormData();fd.append("text",text);
        const res=await fetch("/sentence-segment",{method:"POST",body:fd});
        if(!res.ok||!res.body){rwStatus.textContent="斷詞失敗："+res.status;return;}
        const reader=res.body.getReader(),decoder=new TextDecoder();
        let buf="",tokens=null,errMsg=null;
        while(true){
          const {value,done}=await reader.read();
          if(done)break;
          buf+=decoder.decode(value,{stream:true});
          let idx;
          while((idx=buf.indexOf("\n\n"))>=0){
            const line=buf.slice(0,idx);buf=buf.slice(idx+2);
            if(!line.startsWith("data: "))continue;
            const data=JSON.parse(line.slice(6));
            if(data.error){errMsg=data.error;}
            else if(data.done){tokens=data.tokens;}
            else if(data.progress!==undefined){bumpProgress(stagePct(data.progress,data.total,0,30));}
            else if(data.stage==="word_check"){bumpProgress(stagePct(data.done,data.total,30,30));rwStatus.textContent=`查核多字詞是否為辭典詞條中… ${data.done}/${data.total}`;}
            else if(data.stage==="pair_check"){bumpProgress(stagePct(data.done,data.total,60,40));rwStatus.textContent=`第 ${data.round} 輪前後字配對中… ${data.done}/${data.total}`;}
          }
        }
        if(errMsg){rwStatus.textContent="斷詞失敗："+errMsg;return;}
        if(!tokens){rwStatus.textContent="斷詞失敗：連線中斷";return;}
        bumpProgress(100);
        rwTokens=tokens.map(t=>({...t,cur:t.text,curLevel:t.level}));
        const n=rwTokens.filter(t=>t.word).length,ns=rwTokens.filter(t=>t.has_syn).length;
        rwStatus.textContent=`斷詞完成：${n} 個詞，其中 ${ns} 個有近義詞。點詞選擇替換，改完按「確認替換」。`;
        rwWork.style.display="block";rwRender();
        rwWork.scrollIntoView({behavior:"smooth",block:"start"});
      }catch(err){
        rwStatus.textContent="斷詞失敗："+err.message;
      }finally{
        setLoading(false);setBusy(btnRwSeg,false);
      }
    }
    btnRwSeg.addEventListener("click",querySubmit("rewrite",btnRwSeg,rwSegment));
    document.getElementById("btn-rw-reset").addEventListener("click",()=>{
      rwTokens.forEach(t=>{t.cur=t.text;t.curLevel=t.level;});
      if(rwActive>=0) rwOpen(rwActive); else rwRender();
    });
    document.getElementById("btn-rw-cancel").addEventListener("click",()=>{
      rwWork.style.display="none";rwTokens=[];rwStatus.textContent="已取消";
    });
    async function copyText(text,btn){
      try{
        await navigator.clipboard.writeText(text);
      }catch(_){
        // 非 https / 舊瀏覽器沒有 clipboard API 時的退路
        const ta=document.createElement("textarea");ta.value=text;document.body.appendChild(ta);
        ta.select();document.execCommand("copy");ta.remove();
      }
      const old=btn.textContent;btn.textContent="已複製 ✓";setTimeout(()=>{btn.textContent=old;},1500);
    }
    // 這句（改寫句）所有的詞：走第一頁同一支 /search 查詢，回傳格式、排序、Excel 都跟第一頁一樣
    async function rwLookupWords(words,onProgress){
      const res=await fetch("/search",{method:"POST",
        headers:{"Content-Type":"application/x-www-form-urlencoded"},
        body:"words="+encodeURIComponent(words.join("，"))});
      if(!res.ok||!res.body) throw new Error("連線錯誤："+res.status);
      const reader=res.body.getReader(),decoder=new TextDecoder();
      let buf="";const results=[];
      while(true){
        const {value,done}=await reader.read();
        if(done)break;
        buf+=decoder.decode(value,{stream:true});
        let idx;
        while((idx=buf.indexOf("\n\n"))>=0){
          const line=buf.slice(0,idx);buf=buf.slice(idx+2);
          if(!line.startsWith("data: "))continue;
          const data=JSON.parse(line.slice(6));
          if(data.error) throw new Error(data.error);
          if(data.done||data.total)continue;
          results.push(data);onProgress(results.length);
        }
      }
      return results.sort(byLevelDesc);
    }
    async function rwDownload(rec,btn){
      btn.disabled=true;
      try{
        const res=await fetch("/export",{method:"POST",headers:{"Content-Type":"application/json"},
          body:JSON.stringify({results:rec.results})});
        if(!res.ok){btn.textContent="下載失敗："+res.status;return;}
        const url=URL.createObjectURL(await res.blob());
        const a=document.createElement("a");a.href=url;a.download=`查詢結果_第${rec.no}段.xlsx`;
        document.body.appendChild(a);a.click();a.remove();URL.revokeObjectURL(url);
      }finally{
        btn.disabled=false;
      }
    }
    document.getElementById("btn-rw-confirm").addEventListener("click",()=>{
      const changed=rwChanged();
      const rec={no:++rwCount,
        original:rwTokens.map(t=>t.text).join(""),rewritten:rwTokens.map(t=>t.cur).join(""),
        origHtml:rwOrigHtml(),newHtml:rwNewHtml(),
        words:[...new Set(rwTokens.filter(t=>t.word).map(t=>t.cur))]};
      // 改寫結果寫回文本（選取範圍的內容沒被動過才寫，避免蓋掉使用者期間手動改的字）
      let writeBack="";
      if(changed.length&&rwRange&&rwText.value.slice(rwRange.start,rwRange.end)===rwRange.text){
        const lead=rwRange.text.length-rwRange.text.trimStart().length;
        const s=rwRange.start+lead,e=s+rec.original.length;
        if(rwText.value.slice(s,e)===rec.original){
          rwText.value=rwText.value.slice(0,s)+rec.rewritten+rwText.value.slice(e);
          writeBack="，已寫回文本";
        }
      }
      const card=document.createElement("div");
      card.className="card rw-result";
      card.innerHTML=`<div class="rw-result-head">第 ${rec.no} 段：替換 ${changed.length} 處${writeBack}</div>
        <div class="rw-copy-row"><div class="rw-box"><h3>原句</h3><div>${rec.origHtml}</div></div><button data-copy="original">複製原句</button></div>
        <div class="rw-copy-row"><div class="rw-box"><h3>改寫句</h3><div>${rec.newHtml}</div></div><button data-copy="rewritten">複製改寫句</button></div>
        ${rec.words.length?`<div class="rw-ask">是否輸出這句的詞表？<button class="btn-go" data-ask="yes">輸出詞表</button><button data-ask="no">不用</button></div>`:""}
        <div class="rw-table-slot"></div>`;
      card.querySelectorAll("[data-copy]").forEach(b=>b.addEventListener("click",()=>copyText(rec[b.dataset.copy],b)));
      const ask=card.querySelector(".rw-ask");
      if(ask){
        ask.querySelector('[data-ask="no"]').addEventListener("click",()=>{ask.remove();showOkFine();});
        ask.querySelector('[data-ask="yes"]').addEventListener("click",async()=>{
          ask.remove();
          const slot=card.querySelector(".rw-table-slot");
          slot.innerHTML=`<div class="rw-hint">查詢這句的 ${rec.words.length} 個詞中…</div>`;
          resetProgress();setLoading(true);
          try{
            rec.results=await rwLookupWords(rec.words,n=>{
              bumpProgress(stagePct(n,rec.words.length,0,100));
              slot.firstElementChild.textContent=`查詢這句的詞中… ${n} 筆（${rec.words.length} 個詞）`;
            });
          }catch(err){
            slot.innerHTML=`<div class="rw-hint">詞表查詢失敗：${esc(err.message)}</div>`;return;
          }finally{
            setLoading(false);
          }
          slot.innerHTML=`<div class="rw-hint">共 ${rec.results.length} 筆（依單字級數降冪排序）</div>
            <div class="segment-line"><span class="seg-label">這句的詞（${rec.words.length} 個，點詞只看那個詞）</span>${
              rec.words.map(w=>`<span class="seg-word" data-w="${esc(w)}">${esc(w)}</span>`).join("")}</div>
            <div class="table-wrap"><div class="result-filter"></div><table>
            <thead><tr><th>漢字詞彙</th><th>音標</th><th>詞類</th><th>意思</th><th>詞彙等級</th></tr></thead>
            <tbody></tbody></table></div>
            <div class="rw-actions"><button class="btn-go">下載這句詞表 Excel</button></div>`;
          makeResultView(slot.querySelector("tbody"),slot.querySelector(".result-filter"),
                         slot.querySelector(".segment-line")).set(rec.results);
          const dl=slot.querySelector(".btn-go");
          dl.addEventListener("click",()=>rwDownload(rec,dl));
        });
      }
      rwHistory.prepend(card);
      rwWork.style.display="none";rwTokens=[];
      rwStatus.textContent=`第 ${rec.no} 段完成${writeBack}。可以繼續選下一段。`;
      card.scrollIntoView({behavior:"smooth",block:"start"});
    });

    // ── 筆順動畫：每個字查教育部筆順網 ID，用官方嵌入碼（iframe）原樣呈現 ──
    const hwInput=document.getElementById("hw-input"),btnHw=document.getElementById("btn-hw"),
          hwStatus=document.getElementById("hw-status"),hwGrid=document.getElementById("hw-grid");
    async function hwLookup(){
      const text=hwInput.value.trim();
      if(!text){hwInput.focus();return;}
      setBusy(btnHw,true);hwGrid.innerHTML="";hwZipbar.style.display="none";
      resetProgress();setLoading(true);hwStatus.textContent="查詢教育部筆順資料中…";
      try{
        const data=await (await fetch("/moe-stroke?text="+encodeURIComponent(text))).json();
        if(!data.chars.length){hwStatus.textContent="請輸入中文字";return;}
        bumpProgress(100);
        hwGrid.innerHTML=data.chars.map(c=>c.id
          ?`<div class="hw-cell"><iframe src="${c.frame}" loading="lazy" allow="fullscreen" title="${esc(c.char)} 筆順動畫"></iframe>`+
            `<div class="hw-cap"><span>「${esc(c.char)}」教育部筆順</span><a href="${c.page}" target="_blank" rel="noopener">開啟原網頁</a></div>`+
            `<button class="hw-dl" data-id="${c.id}" data-char="${esc(c.char)}">下載白底 MP4</button></div>`
          :`<div class="hw-miss"><b>${esc(c.char)}</b>教育部筆順網查無此字<br><small>（簡體字或罕用字不在收錄範圍）</small></div>`).join("");
        const miss=data.chars.filter(c=>!c.id).length;
        hwFound=data.chars.filter(c=>c.id);
        btnHwZip.textContent=`全部下載 .zip（${hwFound.length} 個字）`;
        hwZipbar.style.display=hwFound.length>1?"flex":"none";
        hwStatus.textContent=`共 ${data.chars.length} 個字`+(miss?`，其中 ${miss} 個查無筆順`:"")+"。點動畫下方按鈕可重播、練習。";
      }catch(err){
        hwStatus.textContent="查詢失敗："+err.message;
      }finally{
        setLoading(false);setBusy(btnHw,false);
      }
    }
    const IS_IOS=/iPad|iPhone|iPod/.test(navigator.userAgent)||(navigator.platform==="MacIntel"&&navigator.maxTouchPoints>1);
    const hwZipbar=document.getElementById("hw-zipbar"),btnHwZip=document.getElementById("btn-hw-zip");
    let hwFound=[];
    // 錄好的檔案交給使用者：電腦直接下載；iOS 下載的檔案常打不開，給一顆讓使用者自己點的開啟按鈕
    // （錄完時已不在使用者點擊的當下，自動開新分頁會被 iOS 擋）
    function hwDeliver(data,name,label){
      const url=`/moe-video/${data.job}.${data.fmt}?name=${encodeURIComponent(name)}`;
      if(IS_IOS){
        const how=data.fmt==="zip"?"存到「檔案」後點一下 zip 就會解壓縮":"影片開始播放後按「分享」→「儲存影片」";
        hwStatus.innerHTML=`<a class="btn-open" href="${url}" target="_blank" rel="noopener">點我開啟 ${esc(label)}</a><br>開啟後：${how}`;
      }else{
        const a=document.createElement("a");a.href=url+"&dl=1";
        document.body.appendChild(a);a.click();a.remove();
        hwStatus.textContent=`${label} 已下載，在瀏覽器的「下載項目」資料夾（Mac：~/Downloads）。`;
      }
    }
    // 單字下載：伺服器用看不見的瀏覽器把官方動畫從頭播到完錄下來，要等實際播放時間
    hwGrid.addEventListener("click",async e=>{
      const btn=e.target.closest(".hw-dl");
      if(!btn||btn.disabled)return;
      btn.disabled=true;
      const label=btn.textContent;
      const t0=Date.now(),tick=setInterval(()=>{btn.textContent=`錄製中… ${Math.round((Date.now()-t0)/1000)} 秒`;},500);
      btn.textContent="錄製中…";
      try{
        const fd=new FormData();fd.append("id",btn.dataset.id);
        const data=await (await fetch("/moe-video",{method:"POST",body:fd})).json();
        if(data.error){btn.textContent=label;hwStatus.textContent="錄製失敗："+data.error;return;}
        btn.textContent=label+" ✓";
        hwDeliver(data,btn.dataset.char,`「${btn.dataset.char}」白底 MP4`);
      }catch(err){
        btn.textContent=label;hwStatus.textContent="錄製失敗："+err.message;
      }finally{
        clearInterval(tick);btn.disabled=false;
      }
    });
    // 全部下載：兩個字同時錄，SSE 回報錄完幾個字，最後打包成一個 zip
    btnHwZip.addEventListener("click",async()=>{
      if(!hwFound.length)return;
      const label=btnHwZip.textContent,chars=hwFound.map(c=>c.char).join("");
      btnHwZip.disabled=true;setBusy(btnHw,true);
      resetProgress();setLoading(true);hwStatus.textContent=`錄製中… 0 / ${hwFound.length} 個字`;
      try{
        const fd=new FormData();fd.append("items",JSON.stringify(hwFound.map(c=>({id:c.id,char:c.char}))));
        const res=await fetch("/moe-video-zip",{method:"POST",body:fd});
        if(!res.ok||!res.body){hwStatus.textContent="打包失敗："+res.status;return;}
        const reader=res.body.getReader(),decoder=new TextDecoder();
        let buf="",result=null,errMsg=null;
        while(true){
          const {value,done}=await reader.read();
          if(done)break;
          buf+=decoder.decode(value,{stream:true});
          let idx;
          while((idx=buf.indexOf("\n\n"))>=0){
            const line=buf.slice(0,idx);buf=buf.slice(idx+2);
            if(!line.startsWith("data: "))continue;
            const d=JSON.parse(line.slice(6));
            if(d.error)errMsg=d.error;
            else if(d.finished)result=d;
            else if(d.total){
              bumpProgress(stagePct(d.done,d.total,3,94));
              hwStatus.textContent=`錄製中… ${d.done} / ${d.total} 個字`+(d.char?`（「${d.char}」完成）`:"")+"，兩個字同時錄";
            }
          }
        }
        if(errMsg){hwStatus.textContent="打包失敗："+errMsg;return;}
        if(!result){hwStatus.textContent="打包失敗：連線中斷";return;}
        bumpProgress(100);
        hwDeliver(result,chars,`「${chars}」${hwFound.length} 個字的 zip`);
        showThanks();
      }catch(err){
        hwStatus.textContent="打包失敗："+err.message;
      }finally{
        setLoading(false);btnHwZip.disabled=false;setBusy(btnHw,false);btnHwZip.textContent=label;
      }
    });
    btnHw.addEventListener("click",querySubmit("stroke",btnHw,hwLookup));
    hwInput.addEventListener("keydown",querySubmit("stroke",btnHw,hwLookup));

    // ── 主 logo 選單 ──────────────────────────────────────────────────
    const logoMenu=document.getElementById("logo-menu"),logoDropdown=document.getElementById("logo-dropdown"),
          siteLogo=document.getElementById("site-logo");
    const openDropdown=()=>logoDropdown.classList.add("show");
    const closeDropdown=()=>logoDropdown.classList.remove("show");
    let ddHoverTimer=null;
    logoMenu.addEventListener("mouseenter",()=>{clearTimeout(ddHoverTimer);openDropdown();});
    logoMenu.addEventListener("mouseleave",()=>{ddHoverTimer=setTimeout(closeDropdown,200);});
    // 只負責「打開」，不做 toggle：滑鼠靠近時 mouseenter 已經先開了，
    // 緊接著的 click 若做 toggle 反而會把剛打開的選單關掉，變成要點兩次。
    siteLogo.addEventListener("click",openDropdown);
    document.addEventListener("click",e=>{if(!logoMenu.contains(e.target))closeDropdown();});
    document.getElementById("dd-relax").addEventListener("click",closeDropdown);
    document.getElementById("dd-join").addEventListener("click",closeDropdown);

    // ── 新手教學：先問想了解哪一頁，再到那一頁逐步實際操作示範，每步配說明 ──
    // 按任意鍵（或點說明列）下一步、Esc 結束；要等候的示範（查詢、斷詞…）跑完才能往下。
    const tourPick=document.getElementById("tour-pick"),tourBar=document.getElementById("tour-bar"),
          tourStepEl=document.getElementById("tour-step"),tourTextEl=document.getElementById("tour-text"),
          tourHintEl=document.getElementById("tour-hint");
    const sleep=ms=>new Promise(r=>setTimeout(r,ms));
    const waitUntil=async(cond,ms=180000)=>{const t=Date.now();while(!cond()&&Date.now()-t<ms)await sleep(200);};
    async function typeInto(el,text){el.value="";for(const ch of text){el.value+=ch;await sleep(90);}}
    // 模擬把檔案拖進視窗：先亮拖放遮罩，再用示範 PDF 走跟真的拖放同一條路
    async function dropDemoPdf(textarea,report,label){
      dropOverlay.classList.add("show");await sleep(1100);dropOverlay.classList.remove("show");
      const blob=await (await fetch("/branding/tutorial-demo.pdf")).blob();
      await loadFileInto(new File([blob],"教學示範講義.pdf",{type:"application/pdf"}),textarea,report,label);
    }
    const TOURS={
      lookup:[
        {text:"這是「查詢表格」頁：輸入詞彙或貼上整篇文章，系統將會查出每個詞的漢拼、詞類、意思和詞彙等級（國教院華語文語料庫），另外也可以做成Excel檔！",
         run:async()=>showPage("lookup")},
        {el:"#words-input",text:"① 在這裡輸入要查的詞，詞與詞之間可以用空白、頓號、逗號隔開𖦹' ‐ '𖦹",
         run:()=>typeInto(input,"快樂，朋友，學校")},
        {el:"#btn-bpmf",text:"② 不方便打中文？按「⌨ 注音」可以用大千式注音鍵盤輸入，例如打 jau3 會變成「找」。"},
        {el:"#btn-search",text:"③ 按「查詢」（或 Enter）開始查。",run:()=>doSearch(),after:"#table-wrap"},
        {el:"#table-wrap",text:"④ 結果表格：漢字詞彙、音標、詞類、意思、詞彙等級。依詞彙等級由高至低降冪排序，查無資料的會在最上面- ̗̀( ˶^ᵕ'˶)b 上方的按鈕可以只看某個詞彙等級；同一個詞有好幾筆結果時會收起來，按「查看更多」才全部列出。"},
        {el:"#btn-export",text:"⑤ 按「下載 Excel」，整張表會存成「查詢結果.xlsx」。"},
        {el:"#paste-text",text:"⑥ 整篇文章也可以查噢！把文字貼進框框就好ʕ•ﻌ•ʔฅ",
         run:async()=>{pasteText.value="";await typeInto(pasteText,"今天天氣很好。");}},
        {el:"#paste-text",text:"⑦ 或是把 txt／docx／pdf 檔直接拖進視窗，文字會自動放進框裡。現在示範拖入一個 PDF：",
         run:()=>dropDemoPdf(pasteText,setStatus,lookupFile),after:"#lookup-file"},
        {el:"#exclude-input",text:"⑧ 不想查的詞（例如作者名、標題）填在「排除詞」欄位，用逗號隔開。"},
        {el:"#btn-segment",text:"⑨ 按「斷詞並查詢」後，系統會先根據國教院語料庫的資料庫斷詞，再自動查詢，文章越長查詢的時間越久，請耐心等待(🍁•᎑•🍁)",
         run:async()=>{btnSegment.click();await sleep(300);await waitUntil(()=>!isBusy(btnSegment));},after:"#segment-line"},
        {el:"#segment-line",text:"⑩ 這一排是斷出來的詞，點選其中一個詞，下面表格只會顯示那個詞（再點一次取消）。教學結束！"},
      ],
      rewrite:[
        {text:"這是「近義詞替換」頁：選一段文字，把裡面的詞換成近義詞，改完後會輸出替換後的句子及詞表。",
         run:async()=>showPage("rewrite")},
        {el:"#rw-text",text:"① 把文章貼進文本框，或把 txt／docx／pdf 檔直接拖進視窗。現在示範拖入一個 PDF：",
         run:async()=>{rwText.value="";await dropDemoPdf(rwText,rwReport,rwFile);},after:"#rw-file"},
        {el:"#rw-text",text:"② 用滑鼠選取想改寫的那一段。示範選取第二句：",
         run:async()=>{const t="小明覺得很快樂，也很感謝朋友的陪伴。",i=rwText.value.indexOf(t);
                       rwText.focus();if(i>=0)rwText.setSelectionRange(i,i+t.length);}},
        {el:"#btn-rw-seg",text:"③ 按「近義詞替換」，選到的這段會照原本的順序斷詞。",run:()=>rwSegment(),after:"#rw-tokens"},
        {el:"#rw-tokens",text:"④ 藍框的詞代表有近義詞可以替換。點一下會列出相關近義詞和詞彙等級（等級由低到高排序）。示範點「快樂」：",
         run:async()=>{let i=rwTokens.findIndex(t=>t.text==="快樂"&&t.has_syn);if(i<0)i=rwTokens.findIndex(t=>t.has_syn);
                       if(i>=0)await rwOpen(i);},after:"#rw-panel"},
        {el:"#rw-syns",text:"⑤ 點一個近義詞就換上去，換過的詞會變黃框。示範換成「高興」：",
         run:async()=>{const bs=[...rwSyns.querySelectorAll(".rw-syn")];
                       const b=bs.find(x=>x.textContent.startsWith("高興"))||bs[1];if(b)b.click();await sleep(500);},
         after:"#rw-compare"},
        {el:"#rw-compare",text:"⑥ 這裡將會對照原句和改寫句：紅色是被替換掉的、黃色是替換過後的。不滿意可以按「全部還原」。"},
        {el:"#btn-rw-confirm",text:"⑦ 改好按「確認替換」：改寫句會覆寫回文本。",
         run:async()=>{document.getElementById("btn-rw-confirm").click();await sleep(400);},after:".rw-result .rw-copy-row"},
        {el:".rw-result .rw-copy-row",text:"⑧ 原句、改寫句旁邊都有「複製」按鈕，按一下即可複製。"},
        {el:".rw-result .rw-ask",text:"⑨ 接著會問「是否輸出這句的詞表」。示範按「輸出詞表」：",
         run:async()=>{const card=document.querySelector(".rw-result");card.querySelector('[data-ask="yes"]').click();
                       await sleep(300);await waitUntil(()=>card.querySelector(".rw-table-slot table")||/失敗/.test(card.textContent));},
         after:".rw-result .rw-table-slot"},
        {el:".rw-result .rw-table-slot",text:"⑩ 這句所有的詞都查好了，跟「查詢表格」頁一樣可以點詞、依等級篩選、查看更多，也能下載 Excel。教學結束！"},
      ],
      handwriting:[
        {text:"這是「筆順動畫」頁：會顯示教育部官方的標準筆順動畫，還能把筆順錄成影片下載。",
         run:async()=>showPage("handwriting")},
        {el:"#hw-input",text:"① 輸入一個字或一段話（最多 20 字，標點會自動略過）。",run:()=>typeInto(hwInput,"永學")},
        {el:"#btn-hw",text:"② 按「查筆順」，每個字都會出現教育部官方的筆順動畫。",
         run:()=>hwLookup(),after:"#hw-grid"},
        {el:".hw-cell",text:"③ 動畫下方的按鈕可以暫停、逐筆播放、切換快中慢；上方「筆順練習」可以自己描寫看看。"},
        {el:".hw-cell .hw-dl",text:"④ 點選「下載白底 MP4」會把這個字的筆順錄成影片。"},
        {el:"#btn-hw-zip",text:"⑤ 當下載兩個及以上動畫時，按「全部下載 .zip」會一次打包所有的動畫（下載完有驚喜）。教學結束！"},
      ],
    };
    // 「全部功能」：三頁教學接成一串，前兩頁結尾的「教學結束！」改成接下一頁
    const tourChain=(name,next="接著看下一頁 →")=>TOURS[name].map(s=>({...s,text:s.text.replace("教學結束！",next)}));
    // 三頁示範完，最後用箭頭指向左上角 logo：其他功能（重看教學、加入我們…）都在 logo 選單裡
    TOURS.all=[...tourChain("lookup"),...tourChain("rewrite"),...tourChain("handwriting","最後一步 →"),
      {el:"#site-logo",arrow:"#site-logo",text:"其他問題請點這裡查看更多",run:async()=>{window.scrollTo(0,0);}}];
    const tourSkip=document.getElementById("tour-skip"),tourArrow=document.getElementById("tour-arrow");
    // 箭頭放在目標元件正下方、水平置中往上指（右邊是標題，放旁邊會壓到字），捲動或縮放視窗時跟著移
    function placeArrow(){
      const el=tour&&tour.arrow&&document.querySelector(tour.arrow);
      tourArrow.classList.toggle("show",!!el);
      if(!el)return;
      const r=el.getBoundingClientRect();
      tourArrow.style.left=`${r.left+r.width/2-tourArrow.offsetWidth/2}px`;tourArrow.style.top=`${r.bottom+10}px`;
    }
    window.addEventListener("scroll",placeArrow,{passive:true});window.addEventListener("resize",placeArrow);
    let tour=null;
    function tourFocus(sel){
      document.querySelectorAll(".tour-focus").forEach(e=>e.classList.remove("tour-focus"));
      const el=sel&&document.querySelector(sel);
      if(el){el.classList.add("tour-focus");el.scrollIntoView({behavior:"smooth",block:"nearest"});}
    }
    async function tourShow(){
      const st=tour.steps[tour.i],last=tour.i===tour.steps.length-1;
      tourStepEl.textContent=`新手教學 ${tour.i+1} / ${tour.steps.length}`;
      tour.arrow=null;placeArrow();
      tourTextEl.textContent=st.text;
      tourFocus(st.el);
      if(st.run){
        tour.busy=true;tourHintEl.textContent="示範中…請稍候";
        try{await st.run();}catch(err){tourTextEl.textContent=st.text+`（示範失敗：${err.message}）`;}
        if(!tour)return;
        tour.busy=false;
        if(st.after)tourFocus(st.after);
      }
      tourHintEl.textContent=last?"按任意鍵結束教學":"按任意鍵（或點這裡）下一步 · Esc 結束";
      tour.arrow=st.arrow||null;placeArrow();
    }
    function tourNext(){
      if(!tour||tour.busy)return;
      if(tour.i>=tour.steps.length-1){tourEnd();return;}
      tour.i++;tourShow();
    }
    function tourEnd(){
      tour=null;tourBar.classList.remove("show");tourFocus(null);document.body.classList.remove("touring");
      tourSkip.classList.remove("show");tourArrow.classList.remove("show");
      document.removeEventListener("keydown",tourKey,true);
    }
    // capture 階段攔下按鍵：教學中按鍵只用來換步驟，不會打進輸入框或觸發其他快捷鍵
    function tourKey(e){
      e.preventDefault();e.stopPropagation();
      if(e.key==="Escape")tourEnd();else tourNext();
    }
    function tourStart(name,firstRun=false){
      tourPick.classList.remove("show");
      if(!TOURS[name])return;
      tour={steps:TOURS[name],i:0,busy:false,firstRun};
      tourSkip.classList.toggle("show",firstRun);
      // 一開始播就記下來：只有這台電腦第一次執行會自動播，中途關掉程式再開也不會重播
      if(firstRun)markTutorialDone();
      tourBar.classList.add("show");document.body.classList.add("touring");
      document.addEventListener("keydown",tourKey,true);
      tourShow();
    }
    tourBar.addEventListener("click",tourNext);
    tourPick.querySelectorAll("[data-tour]").forEach(b=>b.addEventListener("click",()=>{
      if(b.dataset.tour)tourStart(b.dataset.tour);else{tourPick.classList.remove("show");showWhy();}
    }));
    document.getElementById("dd-tutorial").addEventListener("click",()=>{
      closeDropdown();
      if(tour)tourEnd();
      tourPick.classList.add("show");
    });
    // 這台電腦第一次打開：自動播「全部功能」教學。桌面版記在電腦的使用者資料夾（伺服器回 true/false），
    // 雲端版伺服器回 null，改記在這個瀏覽器
    const LS_TOUR="cndict-tutorial-done";
    function markTutorialDone(){
      fetch("/first-run/done",{method:"POST"}).catch(()=>{});
      try{localStorage.setItem(LS_TOUR,"1");}catch(_){}
    }
    tourSkip.addEventListener("click",e=>{
      e.stopPropagation();
      tourEnd();
      showSkipOkay();
    });

    // 近義詞替換問「是否輸出詞表」按「不用」：全螢幕「喔好吧」動畫，按任意鍵（或點擊）關閉
    // ms：每格固定毫秒數，或每一格各自的毫秒數陣列（照原 GIF 的節奏）
    function makeAnimOverlay(overlay,img,frames,ms){
      let timer=null;
      const close=()=>{
        overlay.classList.remove("show");clearTimeout(timer);timer=null;
        document.removeEventListener("keydown",close);
      };
      overlay.addEventListener("click",close);
      frames.forEach(src=>{new Image().src=src;});  // 預載，第一次跳出來就不會閃
      const delay=k=>Array.isArray(ms)?ms[k]:ms;
      return ()=>{
        let k=0;img.src=frames[0];
        clearTimeout(timer);
        const tick=()=>{k=(k+1)%frames.length;img.src=frames[k];timer=setTimeout(tick,delay(k));};
        timer=setTimeout(tick,delay(0));
        overlay.classList.add("show");
        // 延到下一輪再掛，避免觸發它的同一次點擊/按鍵馬上把它關掉
        setTimeout(()=>document.addEventListener("keydown",close),0);
      };
    }
    const framesOf=(name,n)=>Array.from({length:n},(_,i)=>`/mascot/${name}_${String(i).padStart(2,"0")}.png`);
    const showOkFine=makeAnimOverlay(document.getElementById("okfine-overlay"),document.getElementById("okfine-img"),
                                     framesOf("okfine",12),50);
    // 查詢類按鈕（或 Enter）連按三次：「你的性子也太急了」
    const showImpatient=makeAnimOverlay(document.getElementById("impatient-overlay"),
                                        document.getElementById("impatient-img"),["/branding/impatient.jpg"],60000);
    // 查詢表格頁打開「⌨ 注音」模式：「忘記切換鍵盤了齁」
    const showMeteor=makeAnimOverlay(document.getElementById("meteor-overlay"),document.getElementById("meteor-img"),
                                     framesOf("meteor",15),80);
    // 新手教學選單按「取消」：「那你點我幹嘛」
    const showWhy=makeAnimOverlay(document.getElementById("why-overlay"),document.getElementById("why-img"),
                                  framesOf("why",12),[100,100,800,100,100,800,100,100,800,100,100,900]);
    // 筆順一鍵下載 zip 成功後：「說謝謝」
    const showThanks=makeAnimOverlay(document.getElementById("thanks-overlay"),document.getElementById("thanks-img"),
                                     framesOf("thanks",9),80);

    // 網站做太差？：全螢幕吐槽狗照片，按任意鍵（或點擊）關閉
    const roastOverlay=document.getElementById("roast-overlay");
    const closeRoast=()=>{roastOverlay.classList.remove("show");document.removeEventListener("keydown",closeRoast);};
    document.getElementById("dd-roast").addEventListener("click",()=>{
      closeDropdown();
      roastOverlay.classList.add("show");
      document.addEventListener("keydown",closeRoast);
    });
    roastOverlay.addEventListener("click",closeRoast);

    // 網站成員
    const creditsOverlay=document.getElementById("credits-overlay");
    const closeCredits=()=>{creditsOverlay.classList.remove("show");document.removeEventListener("keydown",closeCredits);};
    document.getElementById("dd-credits").addEventListener("click",()=>{
      closeDropdown();
      creditsOverlay.classList.add("show");
      document.addEventListener("keydown",closeCredits);
    });
    creditsOverlay.addEventListener("click",closeCredits);

    // 作者頁「請作者喝杯咖啡」：全螢幕播狗狗影片，按任意鍵（或點擊）關閉
    const dogOverlay=document.getElementById("dog-overlay"),dogVideo=document.getElementById("dog-video");
    const closeDog=()=>{
      dogOverlay.classList.remove("show");dogVideo.pause();
      document.removeEventListener("keydown",closeDog);
    };
    document.getElementById("coffee-link").addEventListener("click",e=>{
      e.preventDefault();
      closeCredits();
      dogOverlay.classList.add("show");
      dogVideo.currentTime=0;
      dogVideo.play().catch(()=>{dogVideo.muted=true;dogVideo.play();});  // 瀏覽器擋有聲自動播放時改靜音播
      setTimeout(()=>document.addEventListener("keydown",closeDog),0);
    });
    dogOverlay.addEventListener("click",closeDog);

    // 第一次打開的教學按「一鍵跳過」：「好吧」GIF，按任意鍵（或點擊）關閉
    const showSkipOkay=makeAnimOverlay(document.getElementById("skip-overlay"),document.getElementById("skip-img"),
                                       ["/branding/skip-okay.gif"],600000);
    (async()=>{
      let first=null;
      try{first=(await (await fetch("/first-run")).json()).first;}catch(_){}
      if(first===null){try{first=!localStorage.getItem(LS_TOUR);}catch(_){first=false;}}
      if(first)tourStart("all",true);
    })();
  </script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML)

@app.route("/mascot/<name>")
def mascot(name):
    """查詢中在進度條下面放的粉兔子吉祥物（bunny_typing_00~07.png 逐格切換）。"""
    return send_from_directory(MASCOT_DIR, name)

@app.route("/branding/<name>")
def branding(name):
    """主 logo、「網站做太差？」全螢幕彩蛋用的狗照片等站內視覺素材。"""
    return send_from_directory(BRANDING_DIR, name)

SEARCH_WORKERS = 10  # 併發查詢數：詞典查詢是網路 I/O，開多執行緒平行打才不會單詞逐一排隊

@app.route("/search", methods=["GET", "POST"])
def search():
    # 詞彙清單用 POST body（form）傳，query string 塞幾千個詞會超過
    # URL 長度上限被 414 拒絕（長文斷詞出來的詞表就是這樣爆的）。
    # GET 保留給少量詞彙的手動測試用。
    raw = request.values.get("words", "")
    words = [w.strip() for w in re.split(r"[,，、;；\s]+", raw) if w.strip()]
    if not words:
        return Response('data: {"error":"請輸入詞彙"}\n\n', mimetype="text/event-stream")
    session = get_session()
    def generate():
        yield f'data: {json.dumps({"total": len(words)})}\n\n'
        # 原本是「查一個詞、睡 0.2 秒、再查下一個」的序列式查詢，
        # 幾千個詞就要幾十分鐘，網頁看起來像卡死。改用執行緒池平行查，
        # 用 ex.map 保留輸入順序，一個詞失敗不影響其他詞繼續跑。
        with ThreadPoolExecutor(max_workers=SEARCH_WORKERS) as ex:
            def safe_lookup(w):
                try:
                    return lookup(w, session)
                except Exception as e:
                    return [{"word": w, "pinyin": "錯誤", "pos": "—", "definition": str(e), "level": "—"}]
            for w, entries in zip(words, ex.map(safe_lookup, words)):
                for entry in entries:
                    # 查無整詞時會拆成單字查，entry["word"] 就變成單字；記下原本查的詞，
                    # 網頁點斷詞結果只看某個詞時才找得到它拆出來的那幾列
                    entry["query"] = w
                    yield f"data: {json.dumps(entry, ensure_ascii=False)}\n\n"
        yield 'data: {"done":true}\n\n'
    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/segment", methods=["POST"])
def segment():
    """接收整段文字或上傳檔案(txt/docx/pdf)，丟給國教院 COCT 斷詞，
    以 SSE 回報進度（長文會切很多段，逐段回報避免前端像卡死），
    最後回傳去標點、去重複的詞語清單。"""
    upload = request.files.get("file")
    if upload and upload.filename:
        try:
            text = extract_text_from_file(upload.filename, upload.read())
        except Exception as e:
            return Response(f'data: {json.dumps({"error": f"檔案讀取失敗：{e}"}, ensure_ascii=False)}\n\n',
                            mimetype="text/event-stream")
    else:
        text = (request.form.get("text") or "").strip()

    if not text.strip():
        return Response('data: {"error":"請輸入文字或上傳檔案"}\n\n', mimetype="text/event-stream")

    text = _clean_source_text(text)
    exclude = {w.strip() for w in re.split(r"[,，、;；\s]+", request.form.get("exclude", ""))
               if w.strip()}
    session = get_session()

    def generate():
        chunks = _split_sentences(text)  # 一句（以句號為界）就是一個區塊，不再固定字數切
        all_words = []
        total = len(chunks)
        for i, chunk in enumerate(chunks, 1):
            all_words.extend(_segment_chunk(chunk, session))
            yield f'data: {json.dumps({"progress": i, "total": total, "words_so_far": len(set(all_words))}, ensure_ascii=False)}\n\n'

        words = [w for w in dict.fromkeys(all_words)
                  if w not in exclude and not _NON_CHINESE_RE.search(w)]
        if not words:
            yield f'data: {json.dumps({"error": "斷詞結果為空"}, ensure_ascii=False)}\n\n'
            return

        # 斷詞校正（多字詞辭典查核 → 單字前後字配對，見 _resolve_words_stream）
        final_words = words
        for event in _resolve_words_stream(words, session):
            if "__final__" in event:
                final_words = event["__final__"]
            else:
                yield f'data: {json.dumps(event, ensure_ascii=False)}\n\n'

        yield f'data: {json.dumps({"done": True, "words": final_words}, ensure_ascii=False)}\n\n'

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/export", methods=["POST"])
def export():
    data = request.get_json()
    results = data.get("results", [])
    if not results:
        return {"error": "no data"}, 400
    buf = io.BytesIO()
    build_excel(results, buf)
    buf.seek(0)
    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      as_attachment=True, download_name="查詢結果.xlsx")

@app.route("/sentence-segment", methods=["POST"])
def sentence_segment():
    """近義詞替換用的斷詞：跟 /segment 同一套 COCT 斷詞＋辭典校正，
    但保留原句順序與標點（不去重複），回傳 tokens，串起來就是原句。"""
    text = (request.form.get("text") or "").strip()
    if not text:
        return Response('data: {"error":"請輸入句子"}\n\n', mimetype="text/event-stream")
    session = get_session()

    def generate():
        chunks = _split_sentences(text)
        words = []
        for i, chunk in enumerate(chunks, 1):
            words.extend(_segment_chunk(chunk, session))
            yield f'data: {json.dumps({"progress": i, "total": len(chunks)})}\n\n'
        tokens = _align_tokens(text, words)
        for event in _resolve_words_stream(tokens, session, ordered=True):
            if "__final__" in event:
                tokens = event["__final__"]
            else:
                yield f'data: {json.dumps(event, ensure_ascii=False)}\n\n'
        out = [{"text": t, "word": not _NON_CHINESE_RE.search(t),
                "level": VOCAB_LEVEL.get(t, "—"), "has_syn": bool(synonyms_of(t))}
               for t in tokens]
        yield f'data: {json.dumps({"done": True, "tokens": out}, ensure_ascii=False)}\n\n'

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

def moe_stroke_id(char, session):
    """查教育部筆順網這個字的 ID（嵌入碼 dictFrame.jsp?ID= 用）。
    查詢頁對查不到的字有時會導到別的字，所以要比對結果頁上的字（w='X'）是不是本人。"""
    try:
        r = session.get(f"{STROKE_BASE}/searchW.jsp", params={"ID2": "1", "WORD": char}, timeout=15)
        r.encoding = "utf-8"
        m = re.search(r"dictView\.jsp\?ID=(\d+)", r.url + r.text)
        if not m:
            return None
        if "dictView.jsp" not in r.url:
            r = session.get(f"{STROKE_BASE}/dictView.jsp", params={"ID": m.group(1)}, timeout=15)
            r.encoding = "utf-8"
        return m.group(1) if f"w='{char}'" in r.text else None
    except Exception:
        return None

@app.route("/moe-stroke")
def moe_stroke():
    """筆順動畫頁：每個字查教育部筆順網的嵌入 ID，查不到的字回傳 id=null。"""
    chars = [c for c in (request.args.get("text") or "") if not _NON_CHINESE_RE.match(c)][:20]
    session = get_session()
    with ThreadPoolExecutor(max_workers=6) as ex:
        ids = list(ex.map(lambda c: moe_stroke_id(c, session), chars))
    return {"chars": [{"char": c, "id": i,
                       "frame": f"{STROKE_BASE}/dictFrame.jsp?ID={i}" if i else None,
                       "page": f"{STROKE_BASE}/dictView.jsp?ID={i}" if i else None}
                      for c, i in zip(chars, ids)]}

@app.route("/moe-video", methods=["POST"])
def moe_video_record():
    """錄一個字的教育部筆順動畫成白底 MP4；第一次要等動畫實際播完（官方快速），重複下載用暫存。"""
    try:
        path = moe_video.record(request.form.get("id", ""))
    except Exception as e:
        return {"error": str(e)}, 500
    return {"job": path.stem, "fmt": "mp4"}

@app.route("/moe-video-zip", methods=["POST"])
def moe_video_zip():
    """多個字一鍵下載：兩個字同時錄，SSE 回報進度，全部錄完打包成一個 zip。"""
    try:
        items = [(str(it["id"]), str(it["char"])) for it in json.loads(request.form.get("items", "[]"))][:20]
    except Exception:
        items = []
    if not items:
        return Response('data: {"error":"沒有可下載的字"}\n\n', mimetype="text/event-stream")

    def generate():
        import queue
        q = queue.Queue()
        box = {}

        def work():
            try:
                box["zip"] = moe_video.zip_videos(items, "", lambda d, t, c: q.put({"done": d, "total": t, "char": c}))
            except Exception as e:
                box["error"] = str(e)
            q.put(None)

        threading.Thread(target=work, daemon=True).start()
        yield f'data: {json.dumps({"done": 0, "total": len(items)})}\n\n'
        while (ev := q.get()) is not None:
            yield f"data: {json.dumps(ev, ensure_ascii=False)}\n\n"
        if "error" in box:
            yield f'data: {json.dumps({"error": box["error"]}, ensure_ascii=False)}\n\n'
        else:
            yield f'data: {json.dumps({"finished": True, "job": box["zip"].stem, "fmt": "zip"})}\n\n'

    return Response(generate(), mimetype="text/event-stream",
                    headers={"Cache-Control": "no-cache", "X-Accel-Buffering": "no"})

@app.route("/moe-video/<job>.<fmt>")
def moe_video_file(job, fmt):
    path = moe_video.output_path(job, fmt)
    if not path:
        return {"error": "影片已過期，請重新錄製"}, 404
    name = re.sub(r"[^一-鿿]", "", request.args.get("name", ""))[:20] or "筆順"
    mime = "video/mp4" if fmt == "mp4" else "application/zip"
    # dl=1 → 下載（電腦）；否則直接在瀏覽器開啟（iPhone／iPad 下載的檔案常打不開，改成開啟後用「分享」存）
    return send_file(path, mimetype=mime, as_attachment=request.args.get("dl") == "1" or fmt == "zip",
                     download_name=f"筆順_{name}.{fmt}")

@app.route("/extract-text", methods=["POST"])
def extract_text():
    """近義詞替換頁上傳檔案：只抽出純文字放進文本框，不斷詞。"""
    upload = request.files.get("file")
    if not upload or not upload.filename:
        return {"error": "沒有檔案"}, 400
    try:
        return {"text": _clean_source_text(extract_text_from_file(upload.filename, upload.read()))}
    except Exception as e:
        return {"error": str(e)}, 400

@app.route("/synonyms")
def synonyms():
    word = (request.args.get("word") or "").strip()
    return {"word": word, "level": VOCAB_LEVEL.get(word, "—"), "synonyms": synonyms_of(word)}

# ── 本機狀態：這台電腦看過新手教學沒、目前最新版本號、上次跑的 exe 在哪 ──
# 存在使用者資料夾而不是瀏覽器 localStorage：每次啟動 port 可能不同（5000 被佔就換），
# localStorage 跟著 port 走會被當成新網站，教學就會一直重播。
_STATE_DIR = Path(os.environ.get("APPDATA") or Path.home()) / ("CNdict-tool" if os.name == "nt" else ".cndict-tool")
_STATE_FILE = _STATE_DIR / "state.json"
_state_lock = threading.Lock()

def _load_state():
    try:
        return json.loads(_STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _update_state(**kv):
    with _state_lock:
        state = _load_state()
        state.update(kv)
        try:
            _STATE_DIR.mkdir(parents=True, exist_ok=True)
            _STATE_FILE.write_text(json.dumps(state, ensure_ascii=False), encoding="utf-8")
        except OSError:
            pass

def _ver(v):
    try:
        return tuple(int(x) for x in str(v).split("."))
    except ValueError:
        return (0,)

# 打包後的 exe 檔名：GitHub Releases 上叫 CNdict-tool.exe，一鍵打包到桌面的叫 國文辭典查詢工具.exe，
# 使用者也可能自己改成「國語辭典查詢工具.exe」之類；瀏覽器重複下載會再加上「(1)」「 (2)」。
# 所以認 CNdict-tool 開頭，或檔名裡有「辭典查詢」的 exe。
def _is_tool_exe(name):
    n = name.lower()
    return n.endswith(".exe") and (n.startswith("cndict-tool") or "辭典查詢" in n)

_DUP_SUFFIX = re.compile(r"\s*\(\d+\)$")  # 「國語辭典查詢工具(1)」「CNdict-tool (2)」

def _cleanup_log(msg):
    try:
        _STATE_DIR.mkdir(parents=True, exist_ok=True)
        with open(_STATE_DIR / "cleanup.log", "a", encoding="utf-8") as fp:
            fp.write(f"{time.strftime('%Y-%m-%d %H:%M:%S')} v{APP_VERSION} {msg}\n")
    except OSError:
        pass

def _remove_old_versions():
    """新版 exe 啟動時：關掉背景還在跑的舊版（使用者常常關了瀏覽器但 exe 還在背景），再把舊版 exe 刪掉，
    最後把自己的「(1)」去掉改回原檔名。下載當下什麼都做不了，一定要打開新版才會清。
    只在 Windows 打包版執行；如果這台電腦已經跑過更新的版本（使用者誤開舊版），什麼都不動，免得舊版刪到新版。
    每一步都寫進 %APPDATA%\\CNdict-tool\\cleanup.log，沒清乾淨時可以看原因。"""
    if not (getattr(sys, "frozen", False) and os.name == "nt"):
        return
    state = _load_state()
    if _ver(state.get("latest_version", "0")) > _ver(APP_VERSION):
        _cleanup_log(f"略過：這台電腦跑過更新的 v{state.get('latest_version')}")
        return
    me = Path(sys.executable).resolve()
    _cleanup_log(f"啟動 {me}")
    no_window = 0x08000000  # CREATE_NO_WINDOW：--windowed 的 exe 叫 PowerShell 時不要閃黑視窗
    # onefile exe 會有兩個行程（外層解壓的 bootloader + 裡面真正跑的 Python），兩個都是自己，不能殺
    mine = {os.getpid(), os.getppid()}
    old_paths = set()
    try:
        out = subprocess.run(
            ["powershell", "-NoProfile", "-Command",
             "[Console]::OutputEncoding=[Text.Encoding]::UTF8;"
             "Get-CimInstance Win32_Process | Select-Object ProcessId,Name,ExecutablePath | ConvertTo-Json -Compress"],
            capture_output=True, timeout=30, creationflags=no_window).stdout.decode("utf-8", "ignore")
        procs = json.loads(out or "[]")
        for p in procs if isinstance(procs, list) else [procs]:
            if p.get("ProcessId") in mine or not _is_tool_exe(p.get("Name") or ""):
                continue
            r = subprocess.run(["taskkill", "/F", "/T", "/PID", str(p["ProcessId"])],
                               capture_output=True, timeout=15, creationflags=no_window)
            _cleanup_log(f"關閉背景舊版 PID {p['ProcessId']} {p.get('ExecutablePath')}（taskkill={r.returncode}）")
            if p.get("ExecutablePath"):
                old_paths.add(Path(p["ExecutablePath"]))
    except Exception as e:
        _cleanup_log(f"列出行程失敗：{e}")
    # 背景沒在跑的舊版：上次記下的 exe 位置，加上下載／桌面（含 OneDrive 桌面）／同資料夾裡比自己舊的同類 exe
    if state.get("exe_path"):
        old_paths.add(Path(state["exe_path"]))
    my_mtime = me.stat().st_mtime
    home = Path(os.environ.get("USERPROFILE") or Path.home())
    folders = {home / "Downloads", home / "Desktop", me.parent}
    if os.environ.get("OneDrive"):
        folders.add(Path(os.environ["OneDrive"]) / "Desktop")
    for folder in folders:
        try:
            for f in folder.iterdir():
                if _is_tool_exe(f.name) and f.is_file() and f.stat().st_mtime < my_mtime:
                    old_paths.add(f)
        except OSError:
            pass
    for f in old_paths:
        try:
            if f.resolve() == me or not _is_tool_exe(f.name) or not f.exists():
                continue
        except OSError:
            continue
        for _ in range(10):  # 剛被 taskkill 的行程要一下子才會放開檔案
            try:
                f.unlink()
                _cleanup_log(f"刪除舊版 {f}")
                break
            except OSError as e:
                err = e
                time.sleep(0.5)
        else:
            _cleanup_log(f"刪不掉 {f}：{err}")
    # 舊版刪掉後，原本的檔名空出來了：把自己的「(1)」拿掉（Windows 允許改名執行中的 exe，只是不能刪）
    clean = me.with_name(_DUP_SUFFIX.sub("", me.stem) + me.suffix)
    if clean != me and not clean.exists():
        try:
            me.rename(clean)
            _cleanup_log(f"改名 {me.name} → {clean.name}")
            me = clean
        except OSError as e:
            _cleanup_log(f"改名失敗 {me.name}：{e}")
    _update_state(latest_version=APP_VERSION, exe_path=str(me))

@app.route("/first-run")
def first_run():
    # 雲端模式大家共用同一台伺服器，改由瀏覽器自己記（回 null）
    if os.environ.get("PORT"):
        return {"first": None}
    return {"first": not _load_state().get("tutorial_done", False)}

@app.route("/first-run/done", methods=["POST"])
def first_run_done():
    if not os.environ.get("PORT"):
        _update_state(tutorial_done=True)
    return {"ok": True}

if __name__ == "__main__":
    # 雲端平台（如 Zeabur）會注入 PORT 環境變數；本機/打包的 .exe 沒有這個變數，
    # 維持原本「跑在 127.0.0.1 並自動開瀏覽器」的桌面體驗。
    cloud_port = os.environ.get("PORT")
    if cloud_port:
        print(f"啟動中（雲端模式）… 0.0.0.0:{cloud_port}")
        app.run(debug=False, host="0.0.0.0", port=int(cloud_port), threaded=True)
    else:
        # 5000 被別的程式佔走（macOS 的 AirPlay、Windows 上其他服務）就改用系統給的空閒 port
        import socket
        port = PORT
        with socket.socket() as probe:
            try:
                probe.bind(("127.0.0.1", PORT))
            except OSError:
                probe.bind(("127.0.0.1", 0))
                port = probe.getsockname()[1]
        url = f"http://127.0.0.1:{port}"
        print(f"啟動中… {url}")
        threading.Thread(target=_remove_old_versions, daemon=True).start()
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
        app.run(debug=False, host="127.0.0.1", port=port, threaded=True)
