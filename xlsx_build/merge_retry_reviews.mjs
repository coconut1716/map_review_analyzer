import fs from "node:fs/promises";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, "..");

const mainPath = path.join(root, "outputs", "kakao_place_text", "송도과학로_통합리뷰_자동스크롤.xlsx");
const retryPath = path.join(root, "outputs", "kakao_place_text", "재수집_실패3개.xlsx");
const outPath = path.join(root, "outputs", "kakao_place_text", "송도과학로_통합리뷰_자동스크롤_재수집3개반영.xlsx");

const retryUrls = new Set([
  "https://place.map.kakao.com/1684444798#review",
  "https://place.map.kakao.com/300216955#review",
  "https://place.map.kakao.com/1733825326#review",
]);

function normalizeUrl(value) {
  const raw = String(value || "").trim();
  const m = raw.match(/https?:\/\/place\.map\.kakao\.com\/\d+/i);
  return m ? `${m[0].replace(/^http:/i, "https:")}#review` : raw;
}

async function readSheetRows(xlsxPath, sheetName) {
  const input = await FileBlob.load(xlsxPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItem(sheetName);
  const values = sheet.getUsedRange(true).values || [];
  if (!values.length) return { headers: [], rows: [] };
  return { headers: values[0], rows: values.slice(1) };
}

function indexOf(headers, name) {
  const idx = headers.findIndex((h) => String(h || "").trim() === name);
  if (idx < 0) throw new Error(`Missing column: ${name}`);
  return idx;
}

function rowMap(headers, row) {
  const out = {};
  headers.forEach((h, i) => out[String(h || "").trim()] = row[i] ?? "");
  return out;
}

function valueFrom(map, name) {
  return map[name] ?? "";
}

function buildWorkbook(summaryHeaders, summaryRows, reviewHeaders, reviewRows) {
  const workbook = Workbook.create();
  const summarySheet = workbook.worksheets.add("식당요약");
  const reviewSheet = workbook.worksheets.add("리뷰전체");
  summarySheet.showGridLines = false;
  reviewSheet.showGridLines = false;

  summarySheet.getRangeByIndexes(0, 0, summaryRows.length + 1, summaryHeaders.length).values = [summaryHeaders, ...summaryRows];
  summarySheet.getRange("A1:N1").format = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF" } };
  if (summaryRows.length) summarySheet.tables.add(`A1:N${summaryRows.length + 1}`, true, "PlaceSummaryTable");
  summarySheet.freezePanes.freezeRows(1);
  summarySheet.getRange("A:A").format.columnWidthPx = 70;
  summarySheet.getRange("B:C").format.columnWidthPx = 150;
  summarySheet.getRange("D:K").format.columnWidthPx = 110;
  summarySheet.getRange("L:L").format.columnWidthPx = 100;
  summarySheet.getRange("M:N").format.columnWidthPx = 360;
  summarySheet.getRange("D:D").format.numberFormat = "0.0";
  summarySheet.getRange("E:F").format.numberFormat = "#,##0";
  summarySheet.getRange("G:G").format.numberFormat = "0.0%";
  summarySheet.getRange("J:K").format.numberFormat = "#,##0";

  reviewSheet.getRangeByIndexes(0, 0, reviewRows.length + 1, reviewHeaders.length).values = [reviewHeaders, ...reviewRows];
  reviewSheet.getRange("A1:L1").format = { fill: "#244062", font: { bold: true, color: "#FFFFFF" } };
  if (reviewRows.length) {
    reviewSheet.tables.add(`A1:L${reviewRows.length + 1}`, true, "AllReviewsTable");
    reviewSheet.getRange(`C2:C${reviewRows.length + 1}`).format.numberFormat = "0.0";
    reviewSheet.getRange(`D2:D${reviewRows.length + 1}`).format.numberFormat = "#,##0";
    reviewSheet.getRange(`G2:H${reviewRows.length + 1}`).format.numberFormat = "0.0";
    reviewSheet.getRange(`I2:J${reviewRows.length + 1}`).format.numberFormat = "#,##0";
    reviewSheet.getRange(`K2:K${reviewRows.length + 1}`).format.wrapText = true;
  }
  reviewSheet.freezePanes.freezeRows(1);
  reviewSheet.getRange("A:A").format.columnWidthPx = 70;
  reviewSheet.getRange("B:B").format.columnWidthPx = 150;
  reviewSheet.getRange("C:J").format.columnWidthPx = 105;
  reviewSheet.getRange("K:K").format.columnWidthPx = 420;
  reviewSheet.getRange("L:L").format.columnWidthPx = 360;
  return workbook;
}

const mainSummary = await readSheetRows(mainPath, "식당요약");
const mainReviews = await readSheetRows(mainPath, "리뷰전체");
const retrySummary = await readSheetRows(retryPath, "식당요약");
const retryReviews = await readSheetRows(retryPath, "리뷰전체");

const mainSummaryUrlCol = indexOf(mainSummary.headers, "원본 URL");
const retrySummaryUrlCol = indexOf(retrySummary.headers, "원본 URL");
const mainReviewUrlCol = indexOf(mainReviews.headers, "원본 URL");
const retryReviewUrlCol = indexOf(retryReviews.headers, "원본 URL");
const summaryIdCol = indexOf(mainSummary.headers, "식당번호");
const reviewIdCol = indexOf(mainReviews.headers, "식당번호");

const retrySummaryByUrl = new Map();
for (const row of retrySummary.rows) {
  const url = normalizeUrl(row[retrySummaryUrlCol]);
  if (retryUrls.has(url)) retrySummaryByUrl.set(url, row);
}

const retryReviewsByUrl = new Map();
for (const row of retryReviews.rows) {
  const url = normalizeUrl(row[retryReviewUrlCol]);
  if (!retryUrls.has(url)) continue;
  if (!retryReviewsByUrl.has(url)) retryReviewsByUrl.set(url, []);
  retryReviewsByUrl.get(url).push(row);
}

const replaced = [];
const finalSummaryRows = mainSummary.rows.map((row) => {
  const url = normalizeUrl(row[mainSummaryUrlCol]);
  if (!retryUrls.has(url)) return row;
  const retryRow = retrySummaryByUrl.get(url);
  if (!retryRow) return row;
  const next = [...row];
  const originalId = row[summaryIdCol];
  for (let i = 0; i < mainSummary.headers.length; i++) {
    const header = String(mainSummary.headers[i] || "").trim();
    const retryIdx = retrySummary.headers.findIndex((h) => String(h || "").trim() === header);
    if (retryIdx >= 0) next[i] = retryRow[retryIdx] ?? "";
  }
  next[summaryIdCol] = originalId;
  replaced.push({ url, id: originalId, summaryReviews: next[indexOf(mainSummary.headers, "캡처된 리뷰수")] });
  return next;
});

const idByUrl = new Map();
for (const row of finalSummaryRows) idByUrl.set(normalizeUrl(row[mainSummaryUrlCol]), row[summaryIdCol]);

const finalReviewRows = [];
for (const row of mainReviews.rows) {
  const url = normalizeUrl(row[mainReviewUrlCol]);
  if (!retryUrls.has(url)) finalReviewRows.push(row);
}
for (const url of retryUrls) {
  const rows = retryReviewsByUrl.get(url) || [];
  const originalId = idByUrl.get(url) || "";
  for (const row of rows) {
    const next = new Array(mainReviews.headers.length).fill("");
    for (let i = 0; i < mainReviews.headers.length; i++) {
      const header = String(mainReviews.headers[i] || "").trim();
      const retryIdx = retryReviews.headers.findIndex((h) => String(h || "").trim() === header);
      if (retryIdx >= 0) next[i] = row[retryIdx] ?? "";
    }
    next[reviewIdCol] = originalId;
    finalReviewRows.push(next);
  }
}

finalReviewRows.sort((a, b) => {
  const ai = Number(a[reviewIdCol]) || 0;
  const bi = Number(b[reviewIdCol]) || 0;
  if (ai !== bi) return ai - bi;
  return String(a[indexOf(mainReviews.headers, "리뷰 날짜")] || "").localeCompare(String(b[indexOf(mainReviews.headers, "리뷰 날짜")] || ""));
});

const workbook = buildWorkbook(mainSummary.headers, finalSummaryRows, mainReviews.headers, finalReviewRows);
await fs.mkdir(path.dirname(outPath), { recursive: true });
const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save(outPath);

console.log(JSON.stringify({
  outPath,
  replaced,
  mainSummaryRows: mainSummary.rows.length,
  finalSummaryRows: finalSummaryRows.length,
  mainReviewRows: mainReviews.rows.length,
  retryReviewRows: [...retryReviewsByUrl.values()].reduce((s, rows) => s + rows.length, 0),
  finalReviewRows: finalReviewRows.length,
}, null, 2));
