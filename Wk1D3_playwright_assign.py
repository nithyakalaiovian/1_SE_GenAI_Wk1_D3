from playwright.sync_api import sync_playwright

with sync_playwright() as p:
    browser = p.chromium.launch(channel="msedge", headless=False)
    page = browser.new_page()

    page.goto(
        "https://www.cricbuzz.com/cricket-match/live-scores",
        wait_until="domcontentloaded",
    )
    page.locator("body").wait_for()

    page.screenshot(path="cricbuzz_screenshot.png", full_page=True)

    page_text = page.locator("body").inner_text()
    with open("cricbuzz_scorecard.txt", "w", encoding="utf-8") as file:
        file.write(page_text)

    print("Page text saved to cricbuzz_scorecard.txt")
    browser.close()