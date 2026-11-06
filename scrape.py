import os, re, sys, asyncio, time, subprocess, threading
from urllib.parse import urlparse, parse_qs, unquote
import requests
from openpyxl import Workbook, load_workbook
from playwright.async_api import async_playwright

HOW_MANY      = int(os.environ.get("HOW_MANY", "500"))       # max items per main URL
CONCURRENCY   = int(os.environ.get("CONCURRENCY", "6"))      # reels handled at the same time
CHUNK         = int(os.environ.get("CHUNK", "30"))           # download + upload + save in groups of this size
MEGA_DIR      = os.environ.get("MEGA_DIR", "reelsnow")       # folder on Mega that receives the video files
REMOTE        = os.environ.get("MEGA_REMOTE", "mega:")       # rclone remote (root folder of Mega)
SEEN_FILE     = "already_downloaded_urls_reels.txt"
XLSX_FILE     = "results.xlsx"
OUT, DL_DIR, TMP_DIR = "output", "downloads", "tmp_dl"
for _d in (OUT, DL_DIR, TMP_DIR):
    os.makedirs(_d, exist_ok=True)

HEADERS    = ["filename", "title", "final_url"]
VIEWPORT   = {"width": 430, "height": 900}
USER_AGENT = ("Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
              "(KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36")

AGE_SELECTORS = [
    "text=/I\\s*am\\s*18\\+?/i", "text=/I\\s*am\\s*over\\s*18/i", "text=/18\\+/",
    "text=/\\b(enter|agree|accept|continue|yes)\\b/i",
    "button:has-text('18')", "a:has-text('18')",
]
# most specific first: the first one that matches on the page is used
SELECTORS = [".imgprev[onclick*='jumptoid'] img.lazyload",   # reel tiles only (jumptoid / jumptoid2)
             ".imgprev img.lazyload", ".profileside img.lazyload", "img.lazyload"]

# One entry per thumbnail: the reel link (<a>, data-*, or the id inside onclick="jumptoid(203433,..)")
LIST_JS = """els => els.map(e => {
  const abs = u => { try { return new URL(u, location.href).href; } catch (_) { return null; } };
  let href = null, node = e;
  for (let k = 0; k < 6 && node && node !== document.body; k++, node = node.parentElement) {
    if (node.tagName === 'A' && node.getAttribute('href')) { href = abs(node.getAttribute('href')); break; }
    const d = node.dataset || {};
    const key = ['href', 'url', 'link', 'path', 'permalink', 'target'].find(x => d[x]);
    if (key) { href = abs(d[key]); break; }
    const oc = node.getAttribute('onclick');
    if (oc) {
      const pid = oc.match(/jumptoid2?[(][ ]*([0-9]+)/);
      if (pid) { href = location.origin + '/post/' + pid[1]; break; }
      const m = oc.match(/['"]([^'"]+)['"]/); if (m) { href = abs(m[1]); break; }
    }
    if (node.querySelectorAll('img.lazyload').length === 1) {
      const a = node.querySelector('a[href]');
      if (a) { href = abs(a.getAttribute('href')); break; }
    }
  }
  return { alt: e.alt, href: href };
})"""

# Scrolls the last thumbnail into view AND every scrollable container around it, and fires scroll events
SCROLL_JS = """sel => {
  const imgs = document.querySelectorAll(sel);
  const last = imgs[imgs.length - 1];
  if (last) last.scrollIntoView({block: 'end'});
  let p = last ? last.parentElement : null;
  while (p) {
    const oy = getComputedStyle(p).overflowY;
    if ((oy === 'auto' || oy === 'scroll') && p.scrollHeight > p.clientHeight) {
      p.scrollTop = p.scrollHeight; p.dispatchEvent(new Event('scroll'));
    }
    p = p.parentElement;
  }
  for (const s of ['#contentblock', '.rightside', '.rightsideinner', '.profileside']) {
    const el = document.querySelector(s);
    if (el && el.scrollHeight > el.clientHeight) { el.scrollTop = el.scrollHeight; el.dispatchEvent(new Event('scroll')); }
  }
  window.scrollTo(0, document.documentElement.scrollHeight);
  window.dispatchEvent(new Event('scroll'));
}"""

EXTRACT_JS = """() => ({
  final_url: location.href,
  title: (document.title || (document.querySelector('h1') || {}).innerText || '').trim()
})"""

# Every place on a reel page where the video file address can hide
VIDEO_JS = """() => {
  const out = [];
  const add = (u, kind) => {
    if (!u || u.startsWith('blob:') || u.startsWith('data:') || u.startsWith('javascript:')) return;
    try { out.push({url: new URL(u, location.href).href, kind: kind}); } catch (_) {}
  };
  document.querySelectorAll('a[download]').forEach(a => add(a.getAttribute('href'), 'dlattr'));
  document.querySelectorAll('a[href]').forEach(a => {
    const t = (a.innerText || '').trim();
    const h = a.getAttribute('href') || '';
    if (/^download/i.test(t) || /download/i.test(h)) add(h, 'dltext');
  });
  document.querySelectorAll('video').forEach(v => {
    add(v.currentSrc, 'video'); add(v.getAttribute('src'), 'video');
    v.querySelectorAll('source').forEach(s => add(s.getAttribute('src'), 'video'));
  });
  for (const n of ['og:video', 'og:video:url', 'og:video:secure_url', 'twitter:player:stream']) {
    const m = document.querySelector('meta[property="' + n + '"],meta[name="' + n + '"]');
    if (m) add(m.content, 'meta');
  }
  return out;
}"""

DEBUG_VIDEO_JS = """() => {
  const v = document.querySelector('video');
  const links = Array.from(document.querySelectorAll('a')).map(a =>
    (a.innerText || '').trim().slice(0, 25) + ' | ' + (a.getAttribute('href') || '').slice(0, 70))
    .filter(s => /download|mp4|video/i.test(s)).slice(0, 6);
  return { video_tag: v ? v.outerHTML.slice(0, 500) : null, links: links };
}"""

BLOCK = False            # True while opening reel pages: skips images / media / fonts for speed
USED_NAMES = set()       # filenames already used (sheet + this run), so nothing is overwritten
NAME_LOCK = threading.Lock()
DEBUG_SHOWN = 0


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------ small helpers
def norm(u):
    return u.split("#")[0].rstrip("/")


def is_reel_url(u):
    """A real reel page looks like https://www.shorts.xxx/post/208305 (creator pages are /user/..., never saved)."""
    return re.fullmatch(r"/post/\d+/?", urlparse(u).path) is not None


def post_id(u):
    m = re.search(r"/post/(\d+)", u)
    return m.group(1) if m else "x"


def tab_name(url):
    """Tab name from the main URL: ?search=mama+feed -> mamafeed, /user/teressa/ -> teressa"""
    u = urlparse(url)
    q = parse_qs(u.query)
    name = None
    for key in ("search", "q", "s", "tag", "hashtag", "query", "k"):
        if q.get(key):
            name = q[key][0]
            break
    if not name:
        parts = [x for x in u.path.split("/") if x]
        if parts:
            name = parts[-1]
        elif q:
            name = next(iter(q.values()))[0]
        else:
            name = u.netloc
    name = re.sub(r"[^\w]+", "", unquote(name))
    return (name or "sheet")[:31]


def clean_cell(s):
    return re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f]", "", s or "")


def get_main_urls():
    raw = os.environ.get("MAIN_URLS", "").strip()
    if not raw and os.path.exists("main_urls.txt"):
        with open("main_urls.txt", encoding="utf-8") as f:
            raw = "\n".join(l for l in f if not l.strip().startswith("#"))
    urls = []
    for u in re.split(r"[\s,]+", raw):
        if u.startswith("http") and u not in urls:
            urls.append(u)
    return urls


def chunks(seq, n):
    for i in range(0, len(seq), n):
        yield seq[i:i + n]


# ------------------------------------------------------------------ mega (rclone)
def rclone(*args):
    r = subprocess.run(["rclone", *args], capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"rclone {args[0]} failed ({r.returncode}): {r.stderr.strip()[:400]}")
    return r.stdout


def mega_pull():
    # If listing fails we STOP, so a network/login problem can never overwrite the saved files.
    remote_files = set(rclone("lsf", REMOTE, "--files-only").splitlines())
    for name in (SEEN_FILE, XLSX_FILE):
        if name in remote_files:
            rclone("copyto", REMOTE + name, f"{OUT}/{name}")
            log(f"Downloaded {name} from Mega")
        else:
            log(f"{name} is not on Mega yet - a new one will be created")


def upload_downloads():
    """Moves every finished video in downloads/ to Mega folder 'reelsnow'. Files that fail stay in downloads/."""
    files = os.listdir(DL_DIR)
    if not files:
        return
    log(f"Uploading {len(files)} video file(s) to Mega folder '{MEGA_DIR}' ...")
    r = subprocess.run(
        ["rclone", "move", DL_DIR, f"{REMOTE}{MEGA_DIR}", "--transfers", "3",
         "--retries", "3", "--low-level-retries", "8"],
        capture_output=True, text=True)
    left = os.listdir(DL_DIR)
    if r.returncode != 0 or left:
        log(f"  upload problem (rc={r.returncode}): {len(left)} file(s) NOT uploaded. {r.stderr.strip()[-300:]}")
    else:
        log("  upload OK")


def uploaded_name(res):
    """The filename, but only if that file really left downloads/ (= it reached Mega)."""
    fn = (res or {}).get("filename")
    if fn and not os.path.exists(os.path.join(DL_DIR, fn)):
        return fn
    return ""


# ------------------------------------------------------------------ workbook: filename | title | final_url
def style_ws(ws):
    ws.column_dimensions["A"].width = 50
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 60
    ws.freeze_panes = "A2"


def new_ws(wb, name, index=None):
    ws = wb.create_sheet(name, index) if index is not None else wb.create_sheet(name)
    ws.append(HEADERS)
    style_ws(ws)
    return ws


def get_ws(wb, name):
    return wb[name] if name in wb.sheetnames else new_ws(wb, name)


def load_wb():
    path = f"{OUT}/{XLSX_FILE}"
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def migrate_wb(wb):
    """Old tabs had  final_url | title.  Convert them to  filename | title | final_url (filename empty = not downloaded yet)."""
    n = 0
    for ws in list(wb.worksheets):
        head = [c.value for c in ws[1]][:2] if ws.max_row >= 1 else []
        if head == ["final_url", "title"]:
            old = [(r[0], r[1]) for r in ws.iter_rows(min_row=2, max_col=2, values_only=True) if r and r[0]]
            title, idx = ws.title, wb.sheetnames.index(ws.title)
            wb.remove(ws)
            new = new_ws(wb, title, idx)
            for fu, t in old:
                new.append(["", clean_cell(str(t or "")), str(fu).strip()])
            n += 1
    return n


def purge_bad_rows(wb):
    """Deletes rows whose URL is not a reel page (creator links, home page, ads) - saved there by mistake earlier."""
    removed = 0
    for ws in wb.worksheets:
        for r in range(ws.max_row, 1, -1):
            v = ws.cell(row=r, column=3).value
            if isinstance(v, str) and v.strip().startswith("http") and not is_reel_url(v.strip()):
                ws.delete_rows(r)
                removed += 1
    return removed


def scan(wb):
    """saved: every reel url in the sheet.  pending: rows that have no downloaded file yet."""
    saved, pending = {}, []
    for ws in wb.worksheets:
        for r in range(2, ws.max_row + 1):
            url = ws.cell(row=r, column=3).value
            if not (isinstance(url, str) and is_reel_url(url.strip())):
                continue
            url = url.strip()
            k = norm(url)
            if k in saved:
                continue
            saved[k] = url
            fn = ws.cell(row=r, column=1).value
            if not (isinstance(fn, str) and fn.strip()):
                pending.append((ws, r, url))
    return saved, pending


def sheet_filenames(wb):
    out = set()
    for ws in wb.worksheets:
        for r in range(2, ws.max_row + 1):
            fn = ws.cell(row=r, column=1).value
            if isinstance(fn, str) and fn.strip():
                out.add(fn.strip())
    return out


def downloaded_urls(wb):
    out, seen = [], set()
    for ws in wb.worksheets:
        for r in range(2, ws.max_row + 1):
            fn = ws.cell(row=r, column=1).value
            url = ws.cell(row=r, column=3).value
            if isinstance(fn, str) and fn.strip() and isinstance(url, str) and is_reel_url(url.strip()):
                k = norm(url.strip())
                if k not in seen:
                    seen.add(k)
                    out.append(url.strip())
    return out


def persist(wb):
    if not wb.sheetnames:
        return
    wb.save(f"{OUT}/{XLSX_FILE}")
    urls = downloaded_urls(wb)
    with open(f"{OUT}/{SEEN_FILE}", "w", encoding="utf-8") as f:
        f.write("\n".join(urls) + ("\n" if urls else ""))
    # sheet first, then the url list: if something breaks in between we re-check a reel, never lose a row
    rclone("copyto", f"{OUT}/{XLSX_FILE}", REMOTE + XLSX_FILE)
    rclone("copyto", f"{OUT}/{SEEN_FILE}", REMOTE + SEEN_FILE)
    log(f"Saved to Mega: {XLSX_FILE} + {SEEN_FILE} ({len(urls)} downloaded reels)")


# ------------------------------------------------------------------ downloading
def default_filename(headers, url, content_type):
    """The name the browser would give the file: Content-Disposition first, else the last part of the URL."""
    cd = headers.get("Content-Disposition", "") or ""
    name = ""
    m = re.search(r"filename\*\s*=\s*(?:UTF-8'')?([^;]+)", cd, re.I)
    if m:
        name = unquote(m.group(1).strip().strip('"'))
    else:
        m = re.search(r'filename\s*=\s*"?([^";]+)"?', cd, re.I)
        if m:
            name = m.group(1).strip()
    if not name:
        name = unquote(os.path.basename(urlparse(url).path))
    name = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name).strip(" .")
    if not name:
        name = "video"
    if not os.path.splitext(name)[1]:
        ct = (content_type or "").split(";")[0].strip().lower()
        name += {"video/webm": ".webm", "video/quicktime": ".mov"}.get(ct, ".mp4")
    return name


def unique_name(name, pid):
    """Keeps the default name. Only if that exact name is already taken, the reel id is added before the extension."""
    with NAME_LOCK:
        if name in USED_NAMES:
            stem, ext = os.path.splitext(name)
            name = f"{stem}_{pid}{ext}"
        USED_NAMES.add(name)
    return name


def http_download(url, referer):
    ref_host = (urlparse(referer).hostname or "").replace("www.", "")
    headers = {"User-Agent": USER_AGENT, "Referer": referer, "Accept": "*/*"}
    with requests.get(url, headers=headers, stream=True, timeout=(15, 90)) as r:
        r.raise_for_status()
        ct = r.headers.get("Content-Type", "")
        if ct.lower().startswith(("text/", "application/json")):
            raise RuntimeError(f"not a video file ({ct})")
        name = default_filename(r.headers, url, ct)
        tmp = os.path.join(TMP_DIR, f"{threading.get_ident()}_{name}.part")
        size = 0
        with open(tmp, "wb") as f:
            for chunk in r.iter_content(1 << 20):
                f.write(chunk)
                size += len(chunk)
    if size == 0:
        raise RuntimeError("empty file")
    return tmp, name


def ffmpeg_download(url, referer):
    """For .m3u8 playlists: ffmpeg joins the pieces into one mp4 (ffmpeg is preinstalled on GitHub runners)."""
    base = unquote(os.path.basename(urlparse(url).path)) or "video"
    name = os.path.splitext(base)[0] + ".mp4"
    tmp = os.path.join(TMP_DIR, f"{threading.get_ident()}_{name}.part.mp4")
    subprocess.run(
        ["ffmpeg", "-y", "-loglevel", "error",
         "-headers", f"Referer: {referer}\r\nUser-Agent: {USER_AGENT}\r\n",
         "-i", url, "-c", "copy", tmp],
        check=True, timeout=900, capture_output=True)
    return tmp, name


def download_video(url, referer, pid):
    if urlparse(url).path.lower().endswith(".m3u8"):
        tmp, name = ffmpeg_download(url, referer)
    else:
        tmp, name = http_download(url, referer)
    name = unique_name(name, pid)
    os.replace(tmp, os.path.join(DL_DIR, name))
    return name


def pick_video(cands):
    """Chooses the best video address: download button > <video> tag > meta tag > network request."""
    best = None
    for c in cands:
        path = urlparse(c["url"]).path.lower()
        if re.search(r"\.(jpe?g|png|gif|webp|svg|js|css|json|html?|ico|vtt|srt)$", path):
            continue
        vid_ext = re.search(r"\.(mp4|webm|mov|m4v|mkv)$", path) is not None
        m3u8 = path.endswith(".m3u8")
        kind = c["kind"]
        if kind == "dltext" and not (vid_ext or m3u8):
            score = 9                                    # a link that only says 'download' - last resort
        else:
            score = {"dlattr": 0, "video": 1, "meta": 2, "net": 3}.get(kind, 4)
            if m3u8:
                score += 5                               # prefer a plain file over a playlist
        key = (score, 0 if vid_ext else 1)
        if best is None or key < best[0]:
            best = (key, c)
    return best[1] if best else None


# ------------------------------------------------------------------ browser helpers
async def route_handler(route):
    if BLOCK and route.request.resource_type in ("image", "media", "font"):
        await route.abort()
    else:
        await route.continue_()


async def make_ctx(browser):
    ctx = await browser.new_context(viewport=VIEWPORT, user_agent=USER_AGENT)
    await ctx.route("**/*", route_handler)
    return ctx


async def click_age_gate(page, quick=False):
    for sel in AGE_SELECTORS:
        try:
            loc = page.locator(sel).first
            if await loc.count() and await loc.is_visible():
                await loc.click(timeout=2000)
                await page.wait_for_timeout(500 if quick else 1000)
                log("Age gate clicked:", sel)
                return True
        except Exception:
            pass
    return False


async def click_load_more(page):
    for pat in ("load more", "show more", "view more", "see more"):
        try:
            loc = page.locator(f"text=/{pat}/i").first
            if await loc.count() and await loc.is_visible():
                await loc.click(timeout=1500)
                return True
        except Exception:
            pass
    return False


async def scroll_all(page, sel):
    """Keep scrolling (and pressing 'load more') until no new thumbnails appear for 6 rounds in a row."""
    last, stable = 0, 0
    for _ in range(80):
        n = await page.locator(sel).count()
        if n >= HOW_MANY:
            break
        if n == last:
            stable += 1
            if stable >= 6:
                break
        else:
            stable = 0
        last = n
        await click_load_more(page)
        try:
            await page.evaluate(SCROLL_JS, sel)
        except Exception:
            pass
        await page.mouse.move(215, 450)
        await page.mouse.wheel(0, 6000)
        await page.wait_for_timeout(1000)


async def pick_selector(page):
    best, counts = None, {}
    for s in SELECTORS:
        n = await page.locator(s).count()
        counts[s] = n
        if best is None and n > 0:
            best = s
    return best, counts


async def load_list(page, main_url):
    await page.goto(main_url, wait_until="domcontentloaded", timeout=45000)
    await page.wait_for_timeout(1500)
    await click_age_gate(page)
    sel, counts = None, {}
    for _ in range(30):
        sel, counts = await pick_selector(page)
        if sel:
            break
        await page.wait_for_timeout(500)
    if not sel:
        raise RuntimeError("no thumbnails found on the page")
    await page.wait_for_timeout(1500)            # let the main feed load before choosing
    sel, counts = await pick_selector(page)
    await scroll_all(page, sel)
    sel, counts = await pick_selector(page)
    log(f"Thumbnails per selector: {counts}")
    log(f"Scrolled the page: {counts[sel]} thumbnails loaded -> using '{sel}'")
    return sel


async def list_reels(ctx, main_url, seen):
    """Opens one main URL and returns the NEW reel links on it (ads, creator cards and saved reels left out)."""
    global BLOCK
    BLOCK = False
    page = await ctx.new_page()
    try:
        sel = await load_list(page, main_url)
        info = await page.eval_on_selector_all(sel, LIST_JS)
    except Exception as e:
        log("ERROR on main page:", e)
        try:
            log("Page title:", await page.title())
            log("URL:", page.url)
            log("Body text:", (await page.inner_text("body"))[:500])
        except Exception:
            pass
        await page.close()
        raise
    await page.close()

    info = info[:HOW_MANY]
    hrefs, ads, others, nolink = [], 0, 0, 0
    for x in info:
        h = x["href"]
        if not h:
            nolink += 1
        elif "/vast_" in h:
            ads += 1
        elif not is_reel_url(h):
            others += 1
        else:
            hrefs.append(h)
    uniq, keys = [], set()
    for h in hrefs:
        if norm(h) not in keys:
            keys.add(norm(h))
            uniq.append(h)
    new = [h for h in uniq if norm(h) not in seen]
    log(f"{len(info)} tiles | {len(uniq)} reels: {len(uniq) - len(new)} already saved, {len(new)} new | "
        f"ignored: {ads} ads, {others} non-reel links, {nolink} without link")
    return new


async def open_reel(ctx, url):
    """Opens a reel page and collects title, final url and every possible video address."""
    pg = await ctx.new_page()
    sniffed = []

    def on_request(req):
        u = req.url
        if req.resource_type == "media" or re.search(r"\.(mp4|webm|m3u8|mov|m4v)(\?|$)", u, re.I):
            sniffed.append(u)

    pg.on("request", on_request)
    try:
        await pg.goto(url, wait_until="domcontentloaded", timeout=25000)
        data = await pg.evaluate(EXTRACT_JS)
        if not data["title"] and await click_age_gate(pg, quick=True):
            data = await pg.evaluate(EXTRACT_JS)
        cands = []
        for _ in range(8):                                   # wait up to ~4s for the player to set its source
            cands = await pg.evaluate(VIDEO_JS)
            if cands or sniffed:
                break
            await pg.wait_for_timeout(500)
        try:
            await pg.evaluate("() => { const v = document.querySelector('video'); if (v && v.play) v.play().catch(() => {}); }")
        except Exception:
            pass
        await pg.wait_for_timeout(800)
        cands = await pg.evaluate(VIDEO_JS)
        cands += [{"url": u, "kind": "net"} for u in dict.fromkeys(sniffed)]
        debug = None
        if not pick_video(cands):
            try:
                debug = await pg.evaluate(DEBUG_VIDEO_JS)
            except Exception:
                debug = None
        return {**data, "cands": cands, "debug": debug}
    finally:
        await pg.close()


async def handle_reel(ctx, sem, url, seen=None):
    """Opens one reel, downloads its video into downloads/. Returns a dict, or None when it is a duplicate."""
    global DEBUG_SHOWN
    async with sem:
        try:
            info = await open_reel(ctx, url)
        except Exception as e:
            log(f"  fail (open) {url}: {str(e)[:70]}")
            return {"error": "open"}
        fu = info["final_url"]
        title = clean_cell(info["title"])
        if not is_reel_url(fu) or not title:               # age gate / redirect: try again next run
            log(f"  fail (not a reel page) {url} -> {fu}")
            return {"error": "not a reel page"}
        if seen is not None:
            if norm(fu) in seen:
                return None
            seen.add(norm(fu))                              # no await between check and add, so this is safe
        res = {"final_url": fu, "title": title, "filename": ""}
        cand = pick_video(info["cands"])
        if not cand:
            log(f"  no video address found: {fu}")
            if DEBUG_SHOWN < 3 and info.get("debug"):
                DEBUG_SHOWN += 1
                log(f"     debug: {info['debug']}")
                log(f"     candidates: {[c['url'][:90] for c in info['cands']][:5]}")
            return res
        try:
            res["filename"] = await asyncio.to_thread(download_video, cand["url"], fu, post_id(fu))
            log(f"  ok  {res['filename']}   <-  {fu}")
        except Exception as e:
            log(f"  download failed {fu}: {str(e)[:100]}")
        return res


async def main():
    global BLOCK
    t0 = time.time()
    urls = get_main_urls()
    if not urls:
        log("No main URLs. Add them to main_urls.txt or the workflow input.")
        return 1
    log(f"{len(urls)} main URL(s)")

    mega_pull()
    wb = load_wb()
    migrated = migrate_wb(wb)
    removed = purge_bad_rows(wb)
    saved, pending = scan(wb)
    seen = set(saved)
    USED_NAMES.update(sheet_filenames(wb))
    log(f"Reels in the sheet: {len(saved)} | without a downloaded file yet: {len(pending)}")
    for ws in wb.worksheets:
        log(f"  tab '{ws.title}': {max(ws.max_row - 1, 0)} rows")
    if migrated:
        log(f"Converted {migrated} old tab(s) to the layout: filename | title | final_url")
    if removed:
        log(f"Removed {removed} rows that were not reel pages (creator links / home page)")
    dirty = bool(migrated or removed)

    ok_count, total_new, total_files = 0, 0, 0
    async with async_playwright() as p:
        # uses the Google Chrome already installed on GitHub runners (no browser download needed)
        browser = await p.chromium.launch(
            channel="chrome", headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = await make_ctx(browser)

        # ---- 1) reels that are already in the sheet but were never downloaded
        if pending:
            log(f"\n=== Downloading {len(pending)} reel(s) that are in the sheet but have no file yet ===")
            sem = asyncio.Semaphore(CONCURRENCY)
            for chunk in chunks(pending, CHUNK):
                BLOCK = True
                results = await asyncio.gather(*(handle_reel(ctx, sem, url) for _, _, url in chunk))
                BLOCK = False
                upload_downloads()
                got = 0
                for (ws, r, url), res in zip(chunk, results):
                    fn = uploaded_name(res)
                    if fn:
                        ws.cell(row=r, column=1).value = fn
                        if res.get("title") and not ws.cell(row=r, column=2).value:
                            ws.cell(row=r, column=2).value = res["title"]
                        got += 1
                total_files += got
                log(f"  {got}/{len(chunk)} reels now on Mega")
                persist(wb)
                dirty = False

        # ---- 2) new reels from every main URL: one main URL at a time (scrape -> download -> upload -> save)
        for main_url in urls:
            tab = tab_name(main_url)
            log(f"\n=== {main_url}  ->  tab '{tab}' ===")
            try:
                new_hrefs = await list_reels(ctx, main_url, seen)
            except Exception as e:
                log(f"Skipped this URL: {str(e)[:120]}")
                continue
            ok_count += 1
            created = tab not in wb.sheetnames
            ws = get_ws(wb, tab)
            sem = asyncio.Semaphore(CONCURRENCY)
            added = 0
            for chunk in chunks(new_hrefs, CHUNK):
                BLOCK = True
                results = await asyncio.gather(*(handle_reel(ctx, sem, h, seen=seen) for h in chunk))
                BLOCK = False
                upload_downloads()
                for res in results:
                    if not res or res.get("error"):
                        continue
                    fn = uploaded_name(res)
                    ws.append([fn, res["title"], res["final_url"]])
                    added += 1
                    total_files += 1 if fn else 0
                persist(wb)
                dirty = False
            total_new += added
            log(f"{added} new reels -> tab '{tab}' (tab now has {max(ws.max_row - 1, 0)} rows)")
            if added == 0 and (created or dirty):
                persist(wb)
                dirty = False
        await browser.close()

    if dirty:
        persist(wb)
    log(f"\nDone in {time.time() - t0:.0f}s: {total_new} new reels saved to the sheet, "
        f"{total_files} video file(s) uploaded to Mega folder '{MEGA_DIR}', {ok_count}/{len(urls)} main URLs worked")
    return 0 if ok_count else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
