from __future__ import annotations

import re
import shutil
from pathlib import Path

from openpyxl import Workbook
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError
from playwright.sync_api import sync_playwright


TARGET_URL = "https://www.tripadvisor.cn/Attractions-g294211-Activities-China.html"
MAX_PAGES = 10
OUTPUT_XLSX = Path("attractionsForAgent.xlsx")
OUTPUT_VIDEO = Path("tripadvisor-attractions-capture.webm")
VIDEO_DIR = Path("video_tmp")

RANKED_NAME_PATTERN = re.compile(r"^(\d+)\.(.+)$")


def extract_ranked_attractions(page) -> list[tuple[str, str]]:
    """Extract (rank, name) from attraction cards on the current list page."""
    items: list[tuple[str, str]] = []
    cards = page.locator("article")
    for idx in range(cards.count()):
        lines = [line.strip() for line in cards.nth(idx).inner_text().splitlines() if line.strip()]
        for line in lines:
            match = RANKED_NAME_PATTERN.match(line)
            if match:
                rank = match.group(1)
                name = match.group(2).strip()
                items.append((rank, name))
                break
    return items


def save_to_excel(rows: list[tuple[int, str, str]], output_path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Attractions"
    sheet.append(["页码", "编号", "景点名称"])
    for row in rows:
        sheet.append(list(row))
    workbook.save(output_path)


def main() -> None:
    VIDEO_DIR.mkdir(parents=True, exist_ok=True)
    if OUTPUT_VIDEO.exists():
        OUTPUT_VIDEO.unlink()

    all_rows: list[tuple[int, str, str]] = []

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=True)
        context = browser.new_context(
            viewport={"width": 1366, "height": 768},
            locale="zh-CN",
            record_video_dir=str(VIDEO_DIR),
            record_video_size={"width": 1366, "height": 768},
        )
        page = context.new_page()
        page.goto(TARGET_URL, wait_until="domcontentloaded", timeout=120_000)
        page.wait_for_timeout(4_000)

        captured_pages = 0
        video_path = page.video.path()

        while captured_pages < MAX_PAGES:
            current_page_index = captured_pages + 1
            page_items = extract_ranked_attractions(page)
            for rank, name in page_items:
                all_rows.append((current_page_index, rank, name))

            captured_pages += 1
            if captured_pages >= MAX_PAGES:
                break

            next_button = page.locator("a[aria-label='Next page']").first
            if next_button.count() == 0:
                break

            if next_button.get_attribute("aria-disabled") == "true":
                break

            first_item_text = page_items[0][0] + "." + page_items[0][1] if page_items else ""
            next_button.click()

            if first_item_text:
                try:
                    page.wait_for_function(
                        """(needle) => {
                            const firstArticle = document.querySelector("article");
                            if (!firstArticle) return false;
                            return !firstArticle.innerText.includes(needle);
                        }""",
                        arg=first_item_text,
                        timeout=30_000,
                    )
                except PlaywrightTimeoutError:
                    page.wait_for_timeout(3_000)
            else:
                page.wait_for_timeout(3_000)

            page.wait_for_load_state("domcontentloaded")
            page.wait_for_timeout(2_000)

        save_to_excel(all_rows, OUTPUT_XLSX)

        context.close()
        browser.close()

    video_file = Path(video_path)
    if video_file.exists():
        shutil.move(str(video_file), OUTPUT_VIDEO)

    shutil.rmtree(VIDEO_DIR, ignore_errors=True)
    print(f"Captured pages: {captured_pages}")
    print(f"Rows written: {len(all_rows)}")
    print(f"Excel file: {OUTPUT_XLSX.resolve()}")
    print(f"Video file: {OUTPUT_VIDEO.resolve()}")


if __name__ == "__main__":
    main()
