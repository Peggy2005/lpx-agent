#!/usr/bin/env python3
"""
國語辭典查詢工具
用法: python 國語辭典查詢.py
需要: Python 3.8+（套件會自動安裝）
"""

# ── 自動安裝缺少的套件 ───────────────────────────────────────────────
import subprocess, sys, importlib.util

_DEPS = {"flask": "flask", "requests": "requests",
         "bs4": "beautifulsoup4", "openpyxl": "openpyxl",
         "docx": "python-docx", "pypdf": "pypdf"}
_missing = [pkg for mod, pkg in _DEPS.items()
            if importlib.util.find_spec(mod) is None]
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

import re, json, time, threading, webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from flask import Flask, request, Response, render_template_string, send_from_directory
import requests as req
import urllib3
from bs4 import BeautifulSoup
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── 設定 ─────────────────────────────────────────────────────────────
PORT = 5000
DEFAULT_OUTPUT = Path.home() / "Desktop" / "dict_results"

BASE = "https://dict.concised.moe.edu.tw"
MOEDICT_URL = "https://www.moedict.tw/a/{word}.json"
HEADERS = {"User-Agent": "Mozilla/5.0", "Accept-Language": "zh-TW,zh;q=0.9"}
VOCAB_XLSX = Path(__file__).parent / "vocab_14452.xlsx"
EASTER_EGG_DIR = Path(__file__).parent / "easter_eggs"
EASTER_EGG_COUNT = 10

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

def _has_dict_entry(word, session):
    """判斷 word 在萌典或教育部簡編本任一查得到詞條，用來決定斷詞校正時
    一個多字詞/候選詞是否成立——只看「辭典收不收」，不管語料庫怎麼標。"""
    if get_moedict_pos_groups(word, session):
        return True
    entries = lookup_concised(word, session)
    return bool(entries) and entries[0].get("pinyin") != "錯誤"

def _resolve_words_stream(words, session):
    """斷詞校正（generator）：
    ①掃過所有 2 字以上的詞，辭典（萌典＋簡編本，任一查到就算）查得到就保留，
      查不到判定是斷詞切錯，拆成單字放回原本位置。
    ②剩下的單字反覆做前後字配對：第一輪先跳過數詞/量詞/動貌助詞等停用字
      （_STOP_CHARS），由左到右貪婪認領查得到的候選詞，避免同一個字被兩個
      詞重複用掉；配不出新詞了才把停用字放回來一起再試一輪；直到某輪完全
      配不出新詞為止，剩下的單字維持單字（交由呼叫端逐字查辭典）。
    每次批次查詢後 yield 一個進度 dict；結束時 yield {"__final__": [...]}。"""
    multi = [w for w in dict.fromkeys(words) if len(w) > 1]
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
        if len(w) > 1 and w not in hit:
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
            if len(tok) > 1:
                if cur:
                    segs.append(cur)
                cur = []
                continue
            if stoplist_parked and tok in _STOP_CHARS:
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
        found_any = False
        for seg in segments:
            k = 0
            while k < len(seg) - 1:
                a, b = seg[k], seg[k + 1]
                pair = pool[a] + pool[b]
                if pair in hit_pairs and a not in claimed and b not in claimed:
                    merged_words.append(pair)
                    claimed.add(a); claimed.add(b)
                    found_any = True
                    k += 2
                else:
                    k += 1

        if not found_any:
            if stoplist_parked:
                stoplist_parked = False
                continue
            break

        pool = [tok for i, tok in enumerate(pool) if i not in claimed]

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
    """依副檔名從上傳的檔案內容抽出純文字。支援 .txt / .docx / .pdf。"""
    ext = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if ext == "txt":
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
        return "\n".join(page.extract_text() or "" for page in PdfReader(io.BytesIO(raw)).pages)
    raise ValueError(f"不支援的檔案格式：.{ext}（僅支援 txt / docx / pdf）")

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
    # 兩個辭典都沒有完整詞條 — 只要還撈得到詞性或等級，就把能給的給出去，
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

# ── Flask 應用 ────────────────────────────────────────────────────────
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
    button:disabled { opacity: .45; cursor: not-allowed; }
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
    #egg-popup {
      position: fixed; right: 20px; bottom: 20px; width: 180px;
      background: #1e2330; border: 1px solid #2d3548; border-radius: 12px;
      overflow: hidden; box-shadow: 0 8px 28px rgba(0,0,0,.45);
      opacity: 0; transform: translateY(12px) scale(.96); pointer-events: none;
      transition: opacity .35s ease, transform .35s ease; z-index: 999; cursor: pointer;
    }
    #egg-popup.show { opacity: 1; transform: none; pointer-events: auto; }
    #egg-popup img { display: block; width: 100%; height: 160px; object-fit: cover; }
    #egg-popup .egg-caption { padding: 8px 10px; font-size: .78rem; color: #cbd5e1; text-align: center; }
  </style>
</head>
<body>
  <h1>國語辭典查詢</h1>
  <p class="subtitle">教育部《國語辭典簡編本》· 自動輸出 Excel</p>
  <div class="card">
    <label for="words-input">輸入詞彙</label>
    <div class="input-row">
      <input type="text" id="words-input" placeholder="例：給予 讚美、進食；快樂,熱血"
             autocomplete="off" spellcheck="false">
      <button id="btn-bpmf" title="大千式注音輸入模式">⌨ 注音</button>
      <button id="btn-search">查詢</button>
      <button id="btn-export">儲存 Excel</button>
    </div>
    <div id="bpmf-hint"></div>
    <div style="margin-top:14px;">
      <label for="output-path" style="margin-bottom:5px;">儲存路徑</label>
      <input type="text" id="output-path" value="__DEFAULT_OUTPUT__"
             placeholder="留空使用預設路徑" autocomplete="off" spellcheck="false"
             style="width:100%;font-size:.88rem;color:#94a3b8;">
    </div>
    <div id="status"><div class="spinner" id="spinner"></div><span id="status-text"></span></div>
  </div>
  <div class="card" style="margin-top:20px;">
    <label for="paste-text">整段文字（或改用下方上傳檔案）</label>
    <textarea id="paste-text" rows="5" placeholder="貼上一整段文字，會先斷詞再查詢"
      style="width:100%;resize:vertical;background:#0f1117;border:1.5px solid #2d3548;border-radius:8px;
             padding:12px 16px;font-size:.95rem;color:#e2e8f0;outline:none;font-family:inherit;"></textarea>
    <div class="input-row" style="margin-top:12px;align-items:center;">
      <input type="file" id="file-input" accept=".txt,.docx,.pdf"
             style="flex:1;color:#94a3b8;font-size:.85rem;">
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
      支援 txt／docx／pdf。用國教院「教材編輯輔助系統」（遠流語料，關聯詞數量 10）斷詞，
      自動去除標點、重複詞、【】標題括號、頁碼裝飾（如「- 1 -」）、
      以及含阿拉伯數字或非中文字元的詞後查詢。
    </p>
  </div>
  <div class="table-wrap" id="table-wrap">
    <table>
      <thead><tr><th>漢字詞彙</th><th>音標</th><th>詞類</th><th>意思</th><th>詞彙等級</th></tr></thead>
      <tbody id="tbody"></tbody>
    </table>
  </div>
  <div id="egg-popup"><img id="egg-img" alt=""><div class="egg-caption" id="egg-caption"></div></div>
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
    // ── 查詢完成小彩蛋：低機率跳出隨機圖 ─────────────────────────────
    const EGG_CAPTIONS=["查到啦～","辛苦了，休息一下","小獎勵，笑一個","咦，發現你了","繼續加油查詢"];
    let eggTimer=null;
    function maybeShowEasterEgg(){
      if(Math.random()>0.08)return;  // 8% 機率
      const popup=document.getElementById("egg-popup"),
            img=document.getElementById("egg-img"),
            caption=document.getElementById("egg-caption");
      const n=1+Math.floor(Math.random()*10);
      img.src=`/egg/${n}`;
      caption.textContent=EGG_CAPTIONS[Math.floor(Math.random()*EGG_CAPTIONS.length)];
      popup.classList.add("show");
      clearTimeout(eggTimer);
      eggTimer=setTimeout(()=>popup.classList.remove("show"),4000);
    }
    document.getElementById("egg-popup").addEventListener("click",()=>{
      clearTimeout(eggTimer);
      document.getElementById("egg-popup").classList.remove("show");
    });
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
    let allResults=[],es=null;
    const setLoading=on=>{spinner.style.display=on?"block":"none";btnSearch.disabled=on;};
    const setStatus=msg=>{statusTxt.textContent=msg;};
    // 由低到高排等級順序，"—"（沒有等級資料）永遠排最後
    const LEVEL_ORDER=["基礎第1級","基礎第1*級","基礎第2級","基礎第2*級","基礎第3級","基礎第3*級",
                        "進階第4級","進階第4*級","進階第5級","精熟第6級","精熟第7級"];
    const LEVEL_RANK=new Map(LEVEL_ORDER.map((lv,i)=>[lv,i]));
    const levelRank=lv=>LEVEL_RANK.has(lv)?LEVEL_RANK.get(lv):-1;
    function sortByLevelDesc(){
      // 查無資料的排最上面（優先讓使用者看到哪些詞完全沒查到），
      // 其餘照單字級數降冪排序，沒等級資料的（"—"）排最後。
      allResults.sort((a,b)=>{
        const aNone=a.pinyin==="查無資料"?1:0,bNone=b.pinyin==="查無資料"?1:0;
        if(aNone!==bNone) return bNone-aNone;
        return levelRank(b.level)-levelRank(a.level);
      });
      tbody.innerHTML="";
      allResults.forEach(addRow);
    }
    function lvTag(lv){
      if(!lv||lv==="—") return "—";
      const cls=lv.startsWith("基礎")?"基礎":lv.startsWith("進階")?"進階":lv.startsWith("精熟")?"精熟":"";
      return cls?`<span class="lv-tag lv-${cls}">${lv}</span>`:lv;
    }
    function addRow(entry){
      const tr=document.createElement("tr");
      if(entry.pinyin==="查無資料"||entry.pinyin==="錯誤"){
        tr.innerHTML=`<td class="word">${entry.word}</td><td class="error" colspan="3">${entry.pinyin==="錯誤"?"錯誤："+entry.definition:"查無資料"}</td><td class="lv">${lvTag(entry.level)}</td>`;
      } else {
        const pos=entry.pos!=="—"?`<span class="badge">${entry.pos}</span>`:"—";
        tr.innerHTML=`<td class="word">${entry.word}</td><td class="pin">${entry.pinyin}</td><td class="pos">${pos}</td><td class="def">${entry.definition}</td><td class="lv">${lvTag(entry.level)}</td>`;
      }
      tbody.appendChild(tr);
    }
    btnBpmf.addEventListener("click",()=>{
      bpmfMode=!bpmfMode;
      btnBpmf.classList.toggle("active",bpmfMode);
      if(bpmfMode){
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
    async function doSearch(){
      let raw=input.value.trim();if(!raw){input.focus();return;}
      if(bpmfMode) raw=bpmfConvert(raw);
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
              sortByLevelDesc();
              setLoading(false);setStatus(`完成，共 ${allResults.length} 筆（依單字級數降冪排序）`);
              if(allResults.length)btnExport.style.display="inline-block";
              maybeShowEasterEgg();
              continue;
            }
            if(data.error){setStatus("錯誤："+data.error);setLoading(false);continue;}
            if(data.total){totalWords=data.total;setStatus(`查詢中… 0 / ${totalWords} 個詞`);continue;}
            allResults.push(data);tableWrap.style.display="block";addRow(data);
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
    btnSearch.addEventListener("click",doSearch);
    input.addEventListener("keydown",e=>{if(e.key==="Enter")doSearch();});
    btnExport.addEventListener("click",async()=>{
      btnExport.disabled=true;
      const outputDir=document.getElementById("output-path").value.trim();
      const res=await fetch("/export",{method:"POST",headers:{"Content-Type":"application/json"},
        body:JSON.stringify({results:allResults,output_dir:outputDir})});
      const data=await res.json();
      setStatus(data.ok?`已儲存：${data.path}`:"儲存失敗："+(data.error||""));
      btnExport.disabled=false;
    });

    const pasteText=document.getElementById("paste-text"),
          fileInput=document.getElementById("file-input"),
          excludeInput=document.getElementById("exclude-input"),
          btnSegment=document.getElementById("btn-segment");
    btnSegment.addEventListener("click",async()=>{
      const text=pasteText.value.trim(),file=fileInput.files[0];
      if(!text&&!file){pasteText.focus();return;}
      btnSegment.disabled=true;setLoading(true);setStatus("斷詞中…（長文會分段處理，請稍候）");
      try{
        const fd=new FormData();
        if(file) fd.append("file",file); else fd.append("text",text);
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
            else if(data.progress!==undefined){setStatus(`斷詞中… 第 ${data.progress}/${data.total} 段，已找到 ${data.words_so_far} 個詞`);}
            else if(data.stage==="word_check"){setStatus(`查核多字詞是否為辭典詞條中… ${data.done}/${data.total}`);}
            else if(data.stage==="pair_check"){setStatus(`第 ${data.round} 輪前後字配對中… ${data.done}/${data.total}`);}
            else if(data.stage==="def_check"){setStatus(`確認剩餘單字是否有定義中… ${data.done}/${data.total}`);}
          }
        }

        if(errMsg){setStatus("斷詞失敗："+errMsg);setLoading(false);return;}
        if(!words){setStatus("斷詞失敗：連線中斷");setLoading(false);return;}
        setStatus(`斷詞完成，共 ${words.length} 個詞，查詢中…`);
        input.value=words.join("，");
        doSearch();
      }catch(err){
        setStatus("斷詞失敗："+err.message);setLoading(false);
      }finally{
        btnSegment.disabled=false;
      }
    });
  </script>
</body>
</html>"""

@app.route("/")
def index():
    return render_template_string(HTML.replace("__DEFAULT_OUTPUT__", str(DEFAULT_OUTPUT)))

@app.route("/egg/<int:n>")
def easter_egg(n):
    """小彩蛋圖片，查詢完成時前端低機率隨機跳出一張。"""
    return send_from_directory(EASTER_EGG_DIR, f"{n}.jpg")

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
            for entries in ex.map(safe_lookup, words):
                for entry in entries:
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
    out_dir = Path(data.get("output_dir", "").strip() or DEFAULT_OUTPUT)
    try:
        out_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        return {"error": f"路徑無效：{e}"}, 400
    out_path = out_dir / "results.xlsx"
    build_excel(results, str(out_path))
    return {"ok": True, "path": str(out_path)}

if __name__ == "__main__":
    url = f"http://127.0.0.1:{PORT}"
    print(f"啟動中… {url}")
    print(f"預設儲存路徑：{DEFAULT_OUTPUT}")
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(debug=False, host="127.0.0.1", port=PORT, threaded=True)
