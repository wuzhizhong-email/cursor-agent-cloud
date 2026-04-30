/**
 * Opens TripAdvisor China attractions listing in a browser, walks 10 pages
 * via the bottom "Next page" control, extracts attraction id + name from
 * each list card, writes attractionsForAgent.xlsx, and saves a screen recording.
 */
import { chromium } from "playwright";
import ExcelJS from "exceljs";
import { existsSync } from "node:fs";
import { mkdir, readdir } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { execFileSync } from "node:child_process";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const REPO_ROOT = path.resolve(__dirname, "..");
const START_URL =
  "https://www.tripadvisor.cn/Attractions-g294211-Activities-China.html";
const MAX_PAGES = 10;
const ARTIFACTS_DIR = path.join(REPO_ROOT, "artifacts");
const XLSX_PATH = path.join(REPO_ROOT, "attractionsForAgent.xlsx");
const MP4_PATH = path.join(
  REPO_ROOT,
  "tripadvisor-attractions-scrape-recording.mp4",
);

function delay(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

function extractRowsFromPage() {
  return Array.from(document.querySelectorAll("article"))
    .map((article) => {
      const link = article.querySelector('a[href*="Attraction_Review"]');
      const href = link?.getAttribute("href") || "";
      const idMatch = href.match(/-d(\d+)-/);
      const id = idMatch ? idMatch[1] : "";
      const name =
        article.querySelector("h2")?.textContent?.trim() ||
        article.querySelector("img[alt]")?.getAttribute("alt")?.trim() ||
        "";
      return id && name ? { id, name } : null;
    })
    .filter((row) => row != null);
}

async function slowScrollFullPage(page) {
  const height = await page.evaluate(() => document.body.scrollHeight);
  for (let y = 0; y <= height; y += 450) {
    await page.evaluate((yy) => window.scrollTo(0, yy), y);
    await delay(120);
  }
  await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
}

async function extractBatchWithRetry(page) {
  await waitForListingCards(page);
  let batch = await page.evaluate(extractRowsFromPage);
  if (batch.length < 30) {
    await slowScrollFullPage(page);
    await delay(1500);
    batch = await page.evaluate(extractRowsFromPage);
  }
  if (batch.length === 29) {
    await page.reload({ waitUntil: "domcontentloaded", timeout: 90000 });
    await waitForListingCards(page);
    await slowScrollFullPage(page);
    await delay(1500);
    batch = await page.evaluate(extractRowsFromPage);
  }
  return batch;
}

async function waitForListingCards(page) {
  await page.waitForFunction(
    () =>
      document.querySelectorAll('article a[href*="Attraction_Review"]')
        .length >= 10,
    { timeout: 45000 },
  );
  await page
    .locator("article")
    .first()
    .waitFor({ state: "visible", timeout: 15000 });
}

async function main() {
  await mkdir(ARTIFACTS_DIR, { recursive: true });

  const browser = await chromium.launch({
    headless: true,
    args: ["--no-sandbox", "--disable-dev-shm-usage"],
  });

  const context = await browser.newContext({
    viewport: { width: 1280, height: 800 },
    locale: "zh-CN",
    recordVideo: { dir: ARTIFACTS_DIR, size: { width: 1280, height: 800 } },
  });

  const page = await context.newPage();
  await page.goto(START_URL, { waitUntil: "domcontentloaded", timeout: 90000 });
  await waitForListingCards(page);

  const allRows = [];

  for (let pageIndex = 0; pageIndex < MAX_PAGES; pageIndex++) {
    await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight));
    await delay(800);
    const batch = await extractBatchWithRetry(page);

    for (const row of batch) {
      allRows.push(row);
    }

    if (pageIndex < MAX_PAGES - 1) {
      const next = page.locator('a[aria-label="Next page"]').first();
      await next.scrollIntoViewIfNeeded();
      await next.click();
      await page.waitForLoadState("domcontentloaded");
      await delay(1500);
    }
  }

  await context.close();
  await browser.close();

  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("景点", {
    views: [{ state: "frozen", ySplit: 1 }],
  });
  sheet.columns = [
    { header: "编号", key: "id", width: 14 },
    { header: "景点名称", key: "name", width: 48 },
  ];
  sheet.getRow(1).font = { bold: true };

  for (const row of allRows) {
    sheet.addRow({ id: row.id, name: row.name });
  }

  await workbook.xlsx.writeFile(XLSX_PATH);

  const videoFiles = (await readdir(ARTIFACTS_DIR)).filter((f) =>
    f.endsWith(".webm"),
  );
  const webmPath =
    videoFiles.length === 1
      ? path.join(ARTIFACTS_DIR, videoFiles[0])
      : path.join(ARTIFACTS_DIR, videoFiles.sort().at(-1) || "recording.webm");

  try {
    execFileSync(
      "ffmpeg",
      [
        "-y",
        "-i",
        webmPath,
        "-c:v",
        "libx264",
        "-preset",
        "veryfast",
        "-crf",
        "28",
        "-pix_fmt",
        "yuv420p",
        "-movflags",
        "+faststart",
        MP4_PATH,
      ],
      { stdio: "inherit" },
    );
  } catch {
    console.warn("ffmpeg failed; leaving Playwright .webm in artifacts/");
  }

  console.log(
    JSON.stringify(
      {
        rows: allRows.length,
        xlsx: XLSX_PATH,
        videoWebm: webmPath,
        videoMp4: existsSync(MP4_PATH) ? MP4_PATH : null,
      },
      null,
      2,
    ),
  );
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});
