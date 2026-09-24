"""
教育部《國字標準字體筆順學習網》筆順動畫 → MP4 影片下載。

授權：筆順動畫為 CC 姓名標示－非商業性－禁止改作 3.0 臺灣版。
用看不見的瀏覽器打開官方嵌入頁、以官方「快」速從第一筆播到最後一筆並錄影，
輸出「白底 MP4」：灰底換純白、還沒寫的白色字框改淡灰、黑色筆畫不變，給使用者自己編教材用
（放在白色投影片／講義上看起來就像去背；放進教材時請註明出處）。
多個字可打包成一個 .zip。同一個字錄一次後原始錄影暫存 1 小時，重複下載不用重錄。

瀏覽器用使用者電腦上的 Chrome／Edge（Windows 一定有 Edge），不另外打包；
都找不到才嘗試下載 Playwright 自帶的 Chromium。影片編碼用 imageio-ffmpeg 內建的 ffmpeg。
"""
import json, re, shutil, subprocess, sys, tempfile, threading, time, uuid
from pathlib import Path

STROKE_BASE = "https://stroke-order.learningweb.moe.edu.tw"
OUT_DIR = Path(tempfile.gettempdir()) / "dict-tool-moe-video"
VIEW = {"width": 560, "height": 900}   # dictFrame 的方格 = min(高, 寬-32, 512) → 512px
LEAD, TAIL = 0.4, 1.2                  # 開頭留白、寫完後多停幾秒（秒）
OUT_SIZE = 1024                        # 輸出方格邊長（原本 512，放大 2 倍）
# 灰底(168)→白、白色字框(255)→淡灰(225)、黑(0)不變；中間的抗鋸齒過渡跟著線性換算
WHITE_BG_LUT = "if(lte(val,168),val*255/168,255-(val-168)*30/87)"
INSET = 6                              # 往內裁幾 px（原始 512 尺寸），切掉黑色圓角外框與外面的粉紅底
_WIN_NO_CONSOLE = {"creationflags": 0x08000000} if sys.platform == "win32" else {}
_slots = threading.Semaphore(2)        # 同時最多錄 2 個字，避免把電腦拖慢


def ffmpeg_exe():
    try:
        import imageio_ffmpeg
        return imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        return shutil.which("ffmpeg")


def _pw_install(component):
    """用 Playwright 內建的 driver 下載元件（chromium／ffmpeg）。打包成 exe 時 Playwright 會把
    PLAYWRIGHT_BROWSERS_PATH 設成 0（放在套件資料夾裡），這裡要跟它一致才找得到。"""
    from playwright._impl._driver import compute_driver_executable, get_driver_env
    env = get_driver_env()
    if getattr(sys, "frozen", False):
        env.setdefault("PLAYWRIGHT_BROWSERS_PATH", "0")
    subprocess.run([*compute_driver_executable(), "install", component], env=env, check=True,
                   capture_output=True, **_WIN_NO_CONSOLE)


def _launch(pw):
    """依序試：系統 Chrome → 系統 Edge → Playwright 自帶 Chromium（沒有就先下載）。"""
    for channel in ("chrome", "msedge"):
        try:
            return pw.chromium.launch(channel=channel, headless=True)
        except Exception:
            continue
    try:
        return pw.chromium.launch(headless=True)
    except Exception:
        _pw_install("chromium")
        return pw.chromium.launch(headless=True)


def _cleanup_old(max_age=3600):
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    now = time.time()
    for f in OUT_DIR.iterdir():
        try:
            if now - f.stat().st_mtime > max_age:
                shutil.rmtree(f, ignore_errors=True) if f.is_dir() else f.unlink()
        except OSError:
            pass


def _record_raw(stroke_id):
    """錄官方動畫原始影片；Playwright 錄影需要它自己的 ffmpeg 元件，exe 已內建，
    萬一缺少（例如直接用 python 執行、沒裝過）就自動下載一次再重錄。"""
    try:
        return _record_raw_once(stroke_id)
    except Exception as e:
        if "ffmpeg" not in str(e):
            raise
        _pw_install("ffmpeg")
        return _record_raw_once(stroke_id)


def _record_raw_once(stroke_id):
    """錄官方動畫原始影片，回傳 (raw.webm, meta)；同一個字 1 小時內重複要求直接用快取。"""
    from playwright.sync_api import sync_playwright
    work = OUT_DIR / f"raw_{stroke_id}"
    meta_file = work / "meta.json"
    if meta_file.exists():
        meta = json.loads(meta_file.read_text(encoding="utf-8"))
        raw = work / meta["raw"]
        if raw.exists():
            return raw, meta
    shutil.rmtree(work, ignore_errors=True)
    work.mkdir(parents=True, exist_ok=True)
    with sync_playwright() as pw:
        browser = _launch(pw)
        try:
            ctx = browser.new_context(viewport=VIEW, record_video_dir=str(work), record_video_size=VIEW)
            page = ctx.new_page()
            t0 = time.monotonic()                     # 錄影大約從開新分頁這一刻開始
            page.goto(f"{STROKE_BASE}/dictFrame.jsp?ID={stroke_id}", timeout=30000)
            page.wait_for_function("typeof demo !== 'undefined' && demo && demo.status !== undefined",
                                   timeout=30000)
            box = page.locator("#svg").bounding_box()
            # 官方頁面載入後會自己開始播；停下來、切到官方「快」速，再從第一筆完整播一次
            page.evaluate("demo.stop(); demo.setSpeed(FAST)")
            page.wait_for_timeout(int(LEAD * 1000))
            start = time.monotonic() - t0 - LEAD
            page.evaluate("demo.start()")
            page.wait_for_function("demo.status === 'STOP'", timeout=180000, polling=100)
            page.wait_for_timeout(int(TAIL * 1000))
            end = time.monotonic() - t0
            video = page.video
            ctx.close()                               # 關掉 context 影片才會寫完
            raw = Path(video.path())
        finally:
            browser.close()
    if not box:
        raise RuntimeError("找不到筆順方格")
    meta = {"raw": raw.name, "start": start, "end": end,
            "box": [int(round(box[k])) for k in ("x", "y", "width", "height")]}
    meta_file.write_text(json.dumps(meta), encoding="utf-8")
    return raw, meta


def record(stroke_id):
    """錄一個字的官方筆順動畫並輸出白底 MP4，回傳檔案路徑；失敗丟 RuntimeError（訊息給使用者看）。"""
    if not re.fullmatch(r"\d{1,6}", str(stroke_id)):
        raise RuntimeError("字的編號不正確")
    exe = ffmpeg_exe()
    if not exe:
        raise RuntimeError("找不到 ffmpeg（請安裝 imageio-ffmpeg）")
    try:
        import playwright.sync_api  # noqa: F401
    except ImportError:
        raise RuntimeError("缺少 playwright 套件，無法錄影")

    _cleanup_old()
    with _slots:
        raw, meta = _record_raw(stroke_id)
    out = OUT_DIR / f"{uuid.uuid4().hex[:12]}.mp4"
    x, y, w, h = meta["box"]
    vf = (f"trim=start={meta['start']:.2f}:end={meta['end']:.2f},setpts=PTS-STARTPTS,"
          f"crop={w - 2 * INSET}:{h - 2 * INSET}:{x + INSET}:{y + INSET},format=rgb24,"
          f"lutrgb=r='{WHITE_BG_LUT}':g='{WHITE_BG_LUT}':b='{WHITE_BG_LUT}',"
          f"scale={OUT_SIZE}:{OUT_SIZE}:flags=lanczos,fps=30,format=yuv420p")
    cmd = [exe, "-y", "-v", "error", "-i", str(raw), "-vf", vf, "-c:v", "libx264", "-preset", "medium",
           "-crf", "20", "-color_primaries", "bt709", "-color_trc", "bt709", "-colorspace", "bt709",
           "-movflags", "+faststart", str(out)]
    res = subprocess.run(cmd, capture_output=True, **_WIN_NO_CONSOLE)
    if res.returncode:
        msg = res.stderr.decode("utf-8", "ignore").strip().splitlines()
        raise RuntimeError("影片轉檔失敗：" + (msg[-1] if msg else "未知錯誤"))
    return out


def zip_videos(items, name, on_progress=None):
    """items = [(stroke_id, 字)]；兩個字同時錄，全部錄完打包成一個 zip，回傳 zip 路徑。
    on_progress(已完成數, 總數, 字) 每錄完一個字呼叫一次。某個字失敗就整包失敗（訊息含是哪個字）。"""
    from concurrent.futures import ThreadPoolExecutor
    import zipfile
    done = 0
    lock = threading.Lock()

    def one(item):
        nonlocal done
        sid, ch = item
        try:
            path = record(sid)
        except RuntimeError as e:
            raise RuntimeError(f"「{ch}」{e}")
        with lock:
            done += 1
            if on_progress:
                on_progress(done, len(items), ch)
        return ch, path

    with ThreadPoolExecutor(max_workers=2) as ex:
        results = list(ex.map(one, items))
    out = OUT_DIR / f"{uuid.uuid4().hex[:12]}.zip"
    seen = {}
    with zipfile.ZipFile(out, "w", zipfile.ZIP_STORED) as z:   # mp4 已壓縮過，不用再壓
        for n, (ch, path) in enumerate(results, 1):
            seen[ch] = seen.get(ch, 0) + 1
            dup = f"_{seen[ch]}" if seen[ch] > 1 else ""
            z.write(path, f"{n:02d}_筆順_{ch}{dup}.mp4")
    return out


def output_path(job, fmt):
    if not re.fullmatch(r"[0-9a-f]{12}", job or "") or fmt not in ("mp4", "zip"):
        return None
    p = OUT_DIR / f"{job}.{fmt}"
    return p if p.exists() else None
