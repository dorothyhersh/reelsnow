import os, json, asyncio, csv
from playwright.async_api import async_playwright

MAIN_URL    = os.environ.get("MAIN_URL", "https://www.shorts.xxx/user/teressa/")  # put the real address here
HOW_MANY    = int(os.environ.get("HOW_MANY", "75"))
CONCURRENCY = int(os.environ.get("CONCURRENCY", "3"))
OUT = "output"
os.makedirs(OUT, exist_ok=True)

AGE_SELECTORS = [
    "text=/I\\s*am\\s*18\\+?/i", "text=/I\\s*am\\s*over\\s*18/i", "text=/18\\+/",
    "text=/\\b(enter|agree|accept|continue|yes)\\b/i",
    "button:has-text('18')", "a:has-text('18')",
]
THUMBS = ".profileside img.lazyload"

EXTRACT_JS = """() => {
  const m = n => (document.querySelector(`meta[property="${n}"],meta[name="${n}"]`)||{}).content || null;
  return {
    final_url: location.href,
    title: document.title,
    h1: (document.querySelector('h1')||{}).innerText || null,
    description: m('description') || m('og:description'),
    og_title: m('og:title'),
    og_image: m('og:image'),
  };
}"""


async def click_age_gate(page):
    for sel in AGE_SELECTORS:
        try:
            loc = page.locator(sel).first
            if await loc.count() and await loc.is_visible():
                await loc.click(timeout=3000)
                await page.wait_for_timeout(1500)
                print("Age gate clicked:", sel)
                return
        except Exception:
            pass


async def open_main(page):
    await page.goto(MAIN_URL, wait_until="domcontentloaded", timeout=60000)
    await page.wait_for_timeout(2500)
    await click_age_gate(page)
    await page.wait_for_selector(THUMBS, timeout=15000, state="attached")


async def load_all_thumbs(page):
    last = 0
    for _ in range(15):
        n = await page.locator(THUMBS).count()
        if n >= HOW_MANY or n == last:
            break
        last = n
        await page.mouse.wheel(0, 4000)
        await page.wait_for_timeout(800)


async def scrape_link(ctx, sem, idx, alt, url, results):
    async with sem:
        pg = await ctx.new_page()
        try:
            await pg.goto(url, wait_until="domcontentloaded", timeout=45000)
            await pg.wait_for_timeout(1500)
            await click_age_gate(pg)
            data = await pg.evaluate(EXTRACT_JS)
            results.append({"n": idx, "alt": alt, "href": url, **data})
            print(f"[{idx}] ok   {url}")
        except Exception as e:
            print(f"[{idx}] fail {url}: {str(e)[:80]}")
        finally:
            await pg.close()


async def scrape_by_click(ctx, page, idx, alt, results):
    """Fallback for thumbnails with no <a> link: click by index."""
    try:
        await open_main(page)
        await load_all_thumbs(page)
        before = len(ctx.pages)
        await page.evaluate(
            """(i) => { const img = document.querySelectorAll('.profileside img.lazyload')[i];
                        const a = img.closest('a'); (a || img).click(); }""", idx - 1)
        await page.wait_for_timeout(3000)
        target = ctx.pages[-1] if len(ctx.pages) > before else page
        data = await target.evaluate(EXTRACT_JS)
        results.append({"n": idx, "alt": alt, "href": None, **data})
        print(f"[{idx}] ok (click) {target.url}")
        if target is not page:
            await target.close()
    except Exception as e:
        print(f"[{idx}] click fail: {str(e)[:80]}")


async def main():
    results = []
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=True, args=["--no-sandbox", "--disable-dev-shm-usage"])
        ctx = await browser.new_context(
            viewport={"width": 430, "height": 900},
            user_agent=("Mozilla/5.0 (Linux; Android 13; Pixel 7) AppleWebKit/537.36 "
                        "(KHTML, like Gecko) Chrome/124.0 Mobile Safari/537.36"))
        page = await ctx.new_page()
        try:
            await open_main(page)
            await load_all_thumbs(page)
            info = await page.eval_on_selector_all(
                THUMBS,
                "els => els.map(e => ({alt: e.alt, href: e.closest('a') ? e.closest('a').href : null}))")
        except Exception as e:
            print("ERROR on main page:", e)
            print("Page title:", await page.title())
            print("URL:", page.url)
            try:
                print("Body text:", (await page.inner_text("body"))[:500])
            except Exception:
                pass
            await browser.close()
            return results

        info = info[:HOW_MANY]
        print("Thumbnails found:", len(info))
        linked   = [(i + 1, x["alt"], x["href"]) for i, x in enumerate(info) if x["href"]]
        unlinked = [(i + 1, x["alt"]) for i, x in enumerate(info) if not x["href"]]
        print("With link:", len(linked), "| without link:", len(unlinked))

        sem = asyncio.Semaphore(CONCURRENCY)
        await asyncio.gather(*(scrape_link(ctx, sem, i, a, u, results) for i, a, u in linked))
        for i, a in unlinked:
            await scrape_by_click(ctx, page, i, a, results)

        await browser.close()
    return sorted(results, key=lambda r: r["n"])


if __name__ == "__main__":
    results = asyncio.run(main())
    if results:
        with open(f"{OUT}/results.json", "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        with open(f"{OUT}/results.csv", "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(results[0].keys()))
            w.writeheader()
            w.writerows(results)
        print(f"Saved {len(results)} rows to {OUT}/")
    else:
        print("No results. Check the diagnostics printed above.")
        raise SystemExit(1)
