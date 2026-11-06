import os, re, sys, asyncio, time, subprocess
from urllib.parse import urlparse, parse_qs, unquote
from openpyxl import Workbook, load_workbook
from playwright.async_api import async_playwright

HOW_MANY      = int(os.environ.get("HOW_MANY", "75"))        # items per main URL
CONCURRENCY   = int(os.environ.get("CONCURRENCY", "8"))      # parallel tabs for pages that have links
CLICK_WORKERS = int(os.environ.get("CLICK_WORKERS", "4"))    # parallel tabs for click-by-index
REMOTE        = os.environ.get("MEGA_REMOTE", "mega:")       # rclone remote (root folder of Mega)
SEEN_FILE     = "already_downloaded_urls_reels.txt"
XLSX_FILE     = "results.xlsx"
OUT           = "output"
os.makedirs(OUT, exist_ok=True)

AGE_SELECTORS = [
    "text=/I\\s*am\\s*18\\+?/i", "text=/I\\s*am\\s*over\\s*18/i", "text=/18\\+/",
    "text=/\\b(enter|agree|accept|continue|yes)\\b/i",
    "button:has-text('18')", "a:has-text('18')",
]
# tried in order; the first one that matches on the page is used
SELECTORS = [".profileside img.lazyload", "img.lazyload"]

LIST_JS = """els => els.map(e => {
  const abs = u => { try { return new URL(u, location.href).href; } catch (_) { return null; } };
  let href = null, node = e;
  for (let k = 0; k < 6 && node && node !== document.body; k++, node = node.parentElement) {
    if (node.tagName === 'A' && node.getAttribute('href')) { href = abs(node.getAttribute('href')); break; }
    const d = node.dataset || {};
    const key = ['href', 'url', 'link', 'path', 'permalink', 'target'].find(x => d[x]);
    if (key) { href = abs(d[key]); break; }
    const oc = node.getAttribute('onclick');
    if (oc) { const m = oc.match(/['"]([^'"]+)['"]/); if (m) { href = abs(m[1]); break; } }
    if (node.querySelectorAll('img.lazyload').length === 1) {
      const a = node.querySelector('a[href]');
      if (a) { href = abs(a.getAttribute('href')); break; }
    }
  }
  return { alt: e.alt, href: href };
})"""

DEBUG_JS = """sel => {
  const e = document.querySelector(sel);
  if (!e) return null;
  const out = []; let n = e;
  for (let k = 0; k < 3 && n && n !== document.body; k++, n = n.parentElement) out.push(n.outerHTML.slice(0, 400));
  return out;
}"""

CLICK_JS = """([sel, i]) => {
  const img = document.querySelectorAll(sel)[i];
  if (!img) return false;
  img.scrollIntoView({block: 'center'});
  const a = img.closest('a');
  (a || img).click();
  return true;
}"""

EXTRACT_JS = """() => ({
  final_url: location.href,
  title: (document.title || (document.querySelector('h1') || {}).innerText || '').trim()
})"""

BLOCK = False   # True while scraping detail pages: skips images / media / fonts for speed


def log(*a):
    print(*a, flush=True)


# ------------------------------------------------------------------ helpers
def norm(u):
    return u.split("#")[0].rstrip("/")


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


# ------------------------------------------------------------------ mega (rclone)
def rclone(*args):
    r = subprocess.run(["rclone", *args], capture_output=True, text=True)
    if r.returncode != 0:
        raise RuntimeError(f"rclone {args[0]} failed ({r.returncode}): {r.stderr.strip()[:400]}")
    return r.stdout


def mega_pull():
    # If listing fails we STOP, so a network/login problem can never overwrite the saved list.
    remote_files = set(rclone("lsf", REMOTE, "--files-only").splitlines())
    for name in (SEEN_FILE, XLSX_FILE):
        if name in remote_files:
            rclone("copyto", REMOTE + name, f"{OUT}/{name}")
            log(f"Downloaded {name} from Mega")
        else:
            log(f"{name} is not on Mega yet - a new one will be created")


def persist(wb, seen_list):
    wb.save(f"{OUT}/{XLSX_FILE}")
    with open(f"{OUT}/{SEEN_FILE}", "w", encoding="utf-8") as f:
        f.write("\n".join(seen_list) + "\n")
    # sheet first, then the url list: if something breaks in between we get a duplicate, never a lost row
    rclone("copyto", f"{OUT}/{XLSX_FILE}", REMOTE + XLSX_FILE)
    rclone("copyto", f"{OUT}/{SEEN_FILE}", REMOTE + SEEN_FILE)
    log("Uploaded results.xlsx and already_downloaded_urls_reels.txt to Mega")


def load_seen():
    path = f"{OUT}/{SEEN_FILE}"
    seen_list = []
    if os.path.exists(path):
        with open(path, encoding="utf-8") as f:
            seen_list = [l.strip() for l in f if l.strip()]
    return seen_list, {norm(u) for u in seen_list}


def load_wb():
    path = f"{OUT}/{XLSX_FILE}"
    if os.path.exists(path):
        return load_workbook(path)
    wb = Workbook()
    wb.remove(wb.active)
    return wb


def get_ws(wb, name):
    if name in wb.sheetnames:
        return wb[name]
    ws = wb.create_sheet(name)
    ws.append(["final_url", "title"])
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 70
    ws.freeze_panes = "A2"
    return ws


# ------------------------------------------------------------------ browser helpers
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


async def route_handler(route):
    if BLOCK and route.request.resource_type in ("image", "media", "font"):
        await route.abort()
    else:
        await route.continue_()


async def scroll_all(page, sel):
    last = 0
    for _ in range(15):
        n = await page.locator(sel).count()
        if n >= HOW_MANY or n == last:
            break
        last = n
        await page.mouse.wheel(0, 5000)
        await page.wait_for_timeout(400)


async def load_list(page, main_url, first=False):
    await page.goto(main_url, wait_until="domcontentloaded", timeout=45000 if first else 30000)
    await page.wait_for_timeout(1500 if first else 500)
    await click_age_gate(page, quick=not first)
    sel = None
    for _ in range(30):
        for s in SELECTORS:
            if await page.locator(s).count():
                sel = s
                break
        if sel:
            break
        await page.wait_for_timeout(500)
    if not sel:
        raise RuntimeError("no thumbnails found on the page")
    await scroll_all(page, sel)
    return sel


def good_link(h, main_url):
    if not h or not h.startswith("http"):
        return None
    if norm(h) == norm(main_url):
        return None
    return h


def add_result(job, n, data):
    """Keep a result only if it is a real new reel page. Returns True when added."""
    fu = (data.get("final_url") or "").strip()
    title = clean_cell(data.get("title"))
    if not fu.startswith("http") or not title:
        return False
    if norm(fu) == norm(job["main"]) or norm(fu) in job["seen"]:
        return False
    job["seen"].add(norm(fu))                 # no await between check and add, so this is safe
    job["bucket"].append((n, fu, title))
    return True


async def scrape_link(ctx, sem, row, job):
    async with sem:
        pg = await ctx.new_page()
        try:
            await pg.goto(row["href"], wait_until="domcontentloaded", timeout=20000)
            data = await pg.evaluate(EXTRACT_JS)
            if not data["title"] and await click_age_gate(pg, quick=True):
                data = await pg.evaluate(EXTRACT_JS)
            ok = add_result(job, row["n"], data)
            log(f"[{row['n']}] {'new ' if ok else 'skip'} {row['href']}")
        except Exception as e:
            log(f"[{row['n']}] fail {row['href']}: {str(e)[:70]}")
        finally:
            await pg.close()


async def click_worker(ctx, wid, job, rows):
    """Thumbnails with no link: click by index, read the page that opens, come back."""
    page = await ctx.new_page()
    try:
        sel = await load_list(page, job["main"])
    except Exception as e:
        log(f"[worker {wid}] cannot load list: {str(e)[:80]}")
        await page.close()
        return
    for row in rows:
        n = row["n"]
        try:
            before_pages = set(ctx.pages)
            before_url = page.url
            ok = await page.evaluate(CLICK_JS, [sel, n - 1])
            if not ok:
                sel = await load_list(page, job["main"])
                ok = await page.evaluate(CLICK_JS, [sel, n - 1])
            target = None
            for _ in range(30):                      # wait up to ~6s for a new tab or URL change
                await page.wait_for_timeout(200)
                new = [p for p in ctx.pages if p not in before_pages]
                if new:
                    target = new[-1]
                    break
                if page.url != before_url:
                    target = page
                    break
            if target is None:
                log(f"[{n}] click fail: nothing opened")
                continue
            try:
                await target.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            data = await target.evaluate(EXTRACT_JS)
            added = add_result(job, n, data)
            log(f"[{n}] {'new ' if added else 'skip'} (click) {target.url}")
            if target is not page:
                await target.close()
            else:
                sel = await load_list(page, job["main"])   # back to the list
        except Exception as e:
            log(f"[{n}] click fail: {str(e)[:70]}")
            try:
                sel = await load_list(page, job["main"])
            except Exception:
                pass
    await page.close()


async def scrape_main(ctx, main_url, seen):
    """Returns the list of NEW (n, final_url, title) found on one main URL."""
    global BLOCK
    BLOCK = False
    bucket = []
    job = {"main": main_url, "seen": seen, "bucket": bucket}
    page = await ctx.new_page()
    try:
        sel = await load_list(page, main_url, first=True)
        info = await page.eval_on_selector_all(sel, LIST_JS)
        debug = await page.evaluate(DEBUG_JS, sel)
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
    rows = [{"n": i + 1, "href": good_link(x["href"], main_url)} for i, x in enumerate(info)]
    linked   = [r for r in rows if r["href"]]
    unlinked = [r for r in rows if not r["href"]]
    skipped  = [r for r in linked if norm(r["href"]) in seen]
    linked   = [r for r in linked if norm(r["href"]) not in seen]
    log(f"Selector '{sel}': {len(rows)} thumbnails | {len(linked)} new links, "
        f"{len(skipped)} already saved, {len(unlinked)} need click")
    if unlinked and debug:
        log("Thumbnail HTML (child -> parents) for debugging:")
        for d in debug:
            log("  ", d.replace("\n", " "))

    BLOCK = True
    try:
        tasks = []
        if linked:
            sem = asyncio.Semaphore(CONCURRENCY)
            tasks += [scrape_link(ctx, sem, r, job) for r in linked]
        if unlinked:
            k = max(1, min(CLICK_WORKERS, len(unlinked)))
            for w in range(k):
                tasks.append(click_worker(ctx, w + 1, job, unlinked[w::k]))
        await asyncio.gather(*tasks)
    finally:
        BLOCK = False
    bucket.sort(key=lambda t: t[0])
    return bucket


async def main():
    t0 = time.time()
    urls = get_main_urls()
    if not urls:
        log("No main URLs. Add them to main_urls.txt or the workflow input.")
        return 1
    log(f"{len(urls)} main URL(s)")

    mega_pull()
    seen_list, seen = load_seen()
    wb = load_wb()
    log(f"Already saved URLs: {len(seen_list)}")

    ok_count, total_new = 0, 0
    async with async_playwright() as p:
        # uses the Google Chrome already installed on GitHub runners (no browser download needed)
        browser = await p.chromium.launch(
            channel="chrome", headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = await browser.new_context(
            viewport={"width": 430, "height": 900},
            user_agent=("Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36"))
        await ctx.route("**/*", route_handler)

        for main_url in urls:
            tab = tab_name(main_url)
            log(f"\n=== {main_url}  ->  tab '{tab}' ===")
            try:
                new_rows = await scrape_main(ctx, main_url, seen)
            except Exception as e:
                log(f"Skipped this URL: {str(e)[:120]}")
                continue
            ok_count += 1
            ws = get_ws(wb, tab)
            for _, fu, title in new_rows:
                ws.append([fu, title])
                seen_list.append(fu)
            total_new += len(new_rows)
            log(f"{len(new_rows)} new rows -> tab '{tab}'")
            if new_rows:
                persist(wb, seen_list)       # upload after every main URL, so progress is never lost
        await browser.close()

    log(f"\nDone in {time.time()-t0:.0f}s: {total_new} new reels, {ok_count}/{len(urls)} main URLs worked")
    return 0 if ok_count else 1


if __name__ == "__main__":
    sys.exit(asyncio.run(main()))
