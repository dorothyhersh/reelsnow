import os, json, asyncio, csv, time
from playwright.async_api import async_playwright

MAIN_URL       = os.environ.get("MAIN_URL", "https://www.shorts.xxx/user/teressa/")  # put the real address here
HOW_MANY       = int(os.environ.get("HOW_MANY", "75"))
CONCURRENCY    = int(os.environ.get("CONCURRENCY", "8"))       # parallel tabs for pages that have links
CLICK_WORKERS  = int(os.environ.get("CLICK_WORKERS", "4"))     # parallel tabs for click-by-index
SKIP_DETAILS   = os.environ.get("SKIP_DETAILS", "0") == "1"    # 1 = only collect the list (fastest)
OUT = "output"
os.makedirs(OUT, exist_ok=True)

AGE_SELECTORS = [
    "text=/I\\s*am\\s*18\\+?/i", "text=/I\\s*am\\s*over\\s*18/i", "text=/18\\+/",
    "text=/\\b(enter|agree|accept|continue|yes)\\b/i",
    "button:has-text('18')", "a:has-text('18')",
]
THUMBS = ".profileside img.lazyload"

# Finds the link for each thumbnail: <a>, data-* attributes, onclick, or a single <a> inside the card
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
  return {
    alt: e.alt,
    src: e.getAttribute('data-src') || e.getAttribute('data-original') || e.src || null,
    href: href
  };
})"""

DEBUG_JS = """() => {
  const e = document.querySelector('.profileside img.lazyload');
  if (!e) return null;
  const out = []; let n = e;
  for (let k = 0; k < 3 && n && n !== document.body; k++, n = n.parentElement) out.push(n.outerHTML.slice(0, 400));
  return out;
}"""

CLICK_JS = """(i) => {
  const img = document.querySelectorAll('.profileside img.lazyload')[i];
  if (!img) return false;
  img.scrollIntoView({block: 'center'});
  const a = img.closest('a');
  (a || img).click();
  return true;
}"""

EXTRACT_JS = """() => {
  const m = n => (document.querySelector(`meta[property="${n}"],meta[name="${n}"]`)||{}).content || null;
  const v = document.querySelector('video');
  return {
    final_url: location.href,
    title: document.title,
    h1: (document.querySelector('h1')||{}).innerText || null,
    description: m('description') || m('og:description'),
    og_title: m('og:title'),
    og_image: m('og:image'),
    video_src: v ? (v.currentSrc || v.src || (v.querySelector('source')||{}).src || null) : (m('og:video') || null),
  };
}"""


def log(*a):
    print(*a, flush=True)


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


async def route_light(route):
    # after phase 1: skip images / media / fonts so pages load much faster
    if route.request.resource_type in ("image", "media", "font"):
        await route.abort()
    else:
        await route.continue_()


def save(rows, name):
    if not rows:
        return
    with open(f"{OUT}/{name}.json", "w", encoding="utf-8") as f:
        json.dump(rows, f, indent=2, ensure_ascii=False)
    keys = []
    for r in rows:
        for k in r:
            if k not in keys:
                keys.append(k)
    with open(f"{OUT}/{name}.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerows(rows)


async def scroll_all(page):
    last = 0
    for _ in range(15):
        n = await page.locator(THUMBS).count()
        if n >= HOW_MANY or n == last:
            break
        last = n
        await page.mouse.wheel(0, 5000)
        await page.wait_for_timeout(400)


async def load_list(page, first=False):
    await page.goto(MAIN_URL, wait_until="domcontentloaded", timeout=45000 if first else 30000)
    await page.wait_for_timeout(1500 if first else 500)
    await click_age_gate(page, quick=not first)
    await page.wait_for_selector(THUMBS, timeout=15000, state="attached")
    await scroll_all(page)


def good_link(h):
    if not h or not h.startswith("http"):
        return None
    if h.split("#")[0].rstrip("/") == MAIN_URL.rstrip("/"):
        return None
    return h


async def scrape_link(ctx, sem, row, results):
    async with sem:
        pg = await ctx.new_page()
        try:
            await pg.goto(row["href"], wait_until="domcontentloaded", timeout=20000)
            data = await pg.evaluate(EXTRACT_JS)
            if not (data.get("og_title") or data.get("h1")):
                if await click_age_gate(pg, quick=True):
                    data = await pg.evaluate(EXTRACT_JS)
            results.append({**row, **data})
            log(f"[{row['n']}] ok   {row['href']}")
        except Exception as e:
            log(f"[{row['n']}] fail {row['href']}: {str(e)[:70]}")
        finally:
            await pg.close()


async def click_worker(ctx, wid, rows, results):
    """For thumbnails without a link: click by index, read the page that opens, come back."""
    page = await ctx.new_page()
    try:
        await load_list(page)
    except Exception as e:
        log(f"[worker {wid}] cannot load list: {str(e)[:80]}")
        await page.close()
        return
    for row in rows:
        n = row["n"]
        try:
            before_pages = set(ctx.pages)
            before_url = page.url
            ok = await page.evaluate(CLICK_JS, n - 1)
            if not ok:
                await load_list(page)
                ok = await page.evaluate(CLICK_JS, n - 1)
            target = None
            for _ in range(30):                       # wait up to ~6s for a new tab or a URL change
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
            results.append({**row, **data, "href": target.url})
            log(f"[{n}] ok (click) {target.url}")
            if target is not page:
                await target.close()
            else:
                await load_list(page)                 # back to the list for the next one
        except Exception as e:
            log(f"[{n}] click fail: {str(e)[:70]}")
            try:
                await load_list(page)
            except Exception:
                pass
    await page.close()


async def main():
    t0 = time.time()
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = await browser.new_context(
            viewport={"width": 430, "height": 900},
            user_agent=("Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36"))
        page = await ctx.new_page()

        # ---------- PHASE 1: collect the list (saved immediately) ----------
        try:
            await load_list(page, first=True)
            info = await page.eval_on_selector_all(THUMBS, LIST_JS)
            debug = await page.evaluate(DEBUG_JS)
        except Exception as e:
            log("ERROR on main page:", e)
            log("Page title:", await page.title())
            log("URL:", page.url)
            try:
                log("Body text:", (await page.inner_text("body"))[:500])
            except Exception:
                pass
            await browser.close()
            return []

        info = info[:HOW_MANY]
        rows = []
        for i, x in enumerate(info):
            x["href"] = good_link(x["href"])
            rows.append({"n": i + 1, **x})
        save(rows, "list")
        linked   = [r for r in rows if r["href"]]
        unlinked = [r for r in rows if not r["href"]]
        log(f"PHASE 1 done in {time.time()-t0:.0f}s: {len(rows)} thumbnails, "
            f"{len(linked)} with links, {len(unlinked)} without -> output/list.csv")
        if rows:
            log("First 3 rows:")
            for r in rows[:3]:
                log("  ", r["n"], (r["alt"] or "")[:50], "|", r["href"], "|", (r["src"] or "")[:80])
        if unlinked and debug:
            log("Thumbnail HTML (child -> parents) for debugging:")
            for d in debug:
                log("  ", d.replace("\n", " "))
        await page.close()

        # ---------- PHASE 2: details ----------
        results = []
        if not SKIP_DETAILS:
            await ctx.route("**/*", route_light)
            tasks = []
            if linked:
                sem = asyncio.Semaphore(CONCURRENCY)
                tasks += [scrape_link(ctx, sem, r, results) for r in linked]
            if unlinked:
                k = max(1, min(CLICK_WORKERS, len(unlinked)))
                for w in range(k):
                    tasks.append(click_worker(ctx, w + 1, unlinked[w::k], results))
            await asyncio.gather(*tasks)
            results.sort(key=lambda r: r["n"])
            save(results, "results")
            log(f"PHASE 2 done in {time.time()-t0:.0f}s: {len(results)}/{len(rows)} pages scraped -> output/results.csv")

        await browser.close()
    return rows


if __name__ == "__main__":
    rows = asyncio.run(main())
    if not rows:
        log("No results. Check the diagnostics printed above.")
        raise SystemExit(1)
