import os, json, asyncio, csv, time
from playwright.async_api import async_playwright

MAIN_URL     = os.environ.get("MAIN_URL", "https://www.shorts.xxx/user/teressa/")  # put the real address here
HOW_MANY     = int(os.environ.get("HOW_MANY", "75"))
CONCURRENCY  = int(os.environ.get("CONCURRENCY", "8"))
SKIP_DETAILS = os.environ.get("SKIP_DETAILS", "0") == "1"   # 1 = only collect the list (fastest)
OUT = "output"
os.makedirs(OUT, exist_ok=True)

AGE_SELECTORS = [
    "text=/I\\s*am\\s*18\\+?/i", "text=/I\\s*am\\s*over\\s*18/i", "text=/18\\+/",
    "text=/\\b(enter|agree|accept|continue|yes)\\b/i",
    "button:has-text('18')", "a:has-text('18')",
]
THUMBS = ".profileside img.lazyload"

LIST_JS = """els => els.map(e => ({
  alt: e.alt,
  src: e.getAttribute('data-src') || e.getAttribute('data-original') || e.src || null,
  href: e.closest('a') ? e.closest('a').href : null
}))"""

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
    # detail pages only: skip images / media / fonts to load much faster
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


async def scrape_link(ctx, sem, row, results):
    async with sem:
        pg = await ctx.new_page()
        try:
            await pg.route("**/*", route_light)
            await pg.goto(row["href"], wait_until="domcontentloaded", timeout=20000)
            data = await pg.evaluate(EXTRACT_JS)
            # age gate showing again? click once and re-read
            if not (data.get("og_title") or data.get("h1")):
                if await click_age_gate(pg, quick=True):
                    data = await pg.evaluate(EXTRACT_JS)
            results.append({**row, **data})
            log(f"[{row['n']}] ok   {row['href']}")
        except Exception as e:
            log(f"[{row['n']}] fail {row['href']}: {str(e)[:70]}")
        finally:
            await pg.close()


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
            await page.goto(MAIN_URL, wait_until="domcontentloaded", timeout=45000)
            await page.wait_for_timeout(1500)
            await click_age_gate(page)
            await page.wait_for_selector(THUMBS, timeout=15000, state="attached")

            last = 0
            for _ in range(15):
                n = await page.locator(THUMBS).count()
                if n >= HOW_MANY or n == last:
                    break
                last = n
                await page.mouse.wheel(0, 5000)
                await page.wait_for_timeout(400)

            info = await page.eval_on_selector_all(THUMBS, LIST_JS)
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
        rows = [{"n": i + 1, **x} for i, x in enumerate(info)]
        save(rows, "list")
        linked = [r for r in rows if r["href"]]
        log(f"PHASE 1 done in {time.time()-t0:.0f}s: {len(rows)} thumbnails, {len(linked)} with links -> output/list.csv")
        await page.close()

        # ---------- PHASE 2: details in parallel ----------
        results = []
        if not SKIP_DETAILS and linked:
            sem = asyncio.Semaphore(CONCURRENCY)
            await asyncio.gather(*(scrape_link(ctx, sem, r, results) for r in linked))
            results.sort(key=lambda r: r["n"])
            save(results, "results")
            log(f"PHASE 2 done in {time.time()-t0:.0f}s: {len(results)}/{len(linked)} pages scraped")

        await browser.close()
    return rows


if __name__ == "__main__":
    rows = asyncio.run(main())
    if not rows:
        log("No results. Check the diagnostics printed above.")
        raise SystemExit(1)
