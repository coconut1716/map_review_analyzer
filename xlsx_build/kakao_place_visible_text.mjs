import fs from "node:fs/promises";
import { accessSync } from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { FileBlob, SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const toolDir = __dirname;
const workspaceDir = path.resolve(__dirname, "..");
const defaultOutDir = path.join(toolDir, "outputs", "kakao_place_text");

function usage() {
  console.log(`Usage:
  node xlsx_build/kakao_place_visible_text.mjs <url...> [options]

Options:
  --url-file <path>    URL 목록 파일. 한 줄에 URL 하나씩.
  --xlsx-file <path>   XLSX의 링크 열을 읽어서 전체 실행.
  --url-column <name>  XLSX에서 읽을 열 이름. 기본값: review_url
  --combined-xlsx <p>  모든 장소 리뷰를 하나의 XLSX로 모아서 저장.
  --txt-file <path>    이미 저장된 TXT를 파싱해서 XLSX만 생성.
  --out-dir <path>     저장 폴더. 기본값: outputs/kakao_place_text
  --scrolls <n>        자동 스크롤 전 최소 스크롤 횟수. 기본값: 3
  --no-auto-scroll     총 리뷰수 기반 자동 추가 스크롤을 끔.
  --fast-threshold <n> 총 리뷰수가 n개 이하면 빠른 수집 프로파일 사용. 기본값: 20
  --fast-max-scrolls <n> 빠른 프로파일 최대 스크롤. 기본값: 12
  --fast-stall-wait-ms <n> 빠른 프로파일 추가 대기(ms). 기본값: 600
  --max-scrolls <n> 느린 프로파일 자동 스크롤 최대 횟수. 기본값: 200
  --target-coverage <r> 총 리뷰수 대비 목표 수집률. 기본값: 0.95
  --target-review-cap <n> 목표 수집 리뷰 수 상한. 0이면 사용 안 함.
  --resource-mode <normal|lite> Resource loading mode. normal blocks nothing; lite blocks image/media/font. Default: normal
  --resource-stats    Print requested/blocked resource counts to the console. Also enabled by --verbose.
  --stall-limit <n>    느린 프로파일에서 리뷰 수가 늘지 않을 때 멈추는 반복 횟수. 기본값: 8
  --stall-wait-ms <n> 느린 프로파일에서 리뷰 수가 안 늘 때 추가 대기 시간(ms). 기본값: 2000
  --scroll-delay-ms <n> 느린 프로파일 스크롤 사이 대기 시간(ms). 기본값: 300
  --wait-ms <n>        페이지 진입 후 대기 시간(ms). 기본값: 1000
  --limit <n>          앞 n개 URL만 실행. 테스트용.
  --screenshot         전체 페이지 PNG 캡처도 저장.
  --preview            XLSX 미리보기 PNG도 저장.
  --headed             실제 브라우저 창을 띄움.

Example:
  node xlsx_build/kakao_place_visible_text.mjs "https://place.map.kakao.com/450422239#review" --scrolls 10
  node xlsx_build/kakao_place_visible_text.mjs --url-file place_urls.txt --scrolls 8
  node xlsx_build/kakao_place_visible_text.mjs --xlsx-file outputs/kakao_restaurants/restaurants.xlsx --scrolls 10
  node xlsx_build/kakao_place_visible_text.mjs --txt-file outputs/kakao_place_text/450422239_캡틴루이___카카오맵.txt
`);
}

function parseArgs(argv) {
  const args = {
    urls: [],
    urlFile: "",
    xlsxFiles: [],
    urlColumn: "review_url",
    combinedXlsx: "",
    limit: 0,
    txtFiles: [],
    outDir: defaultOutDir,
    scrolls: 3,
    autoScroll: true,
    fastThreshold: 20,
    fastMaxScrolls: 12,
    fastStallWaitMs: 600,
    maxScrolls: 200,
    targetCoverage: 0.95,
    targetReviewCap: 0,
    resourceMode: "normal",
    resourceStats: false,
    stallLimit: 8,
    stallWaitMs: 2000,
    scrollDelayMs: 300,
    waitMs: 1000,
    headed: false,
    screenshot: false,
    preview: false,
  };

  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if (arg === "-h" || arg === "--help") {
      usage();
      process.exit(0);
    }
    if (arg === "--url-file") args.urlFile = argv[++i] || "";
    else if (arg === "--xlsx-file") args.xlsxFiles.push(path.resolve(argv[++i] || ""));
    else if (arg === "--url-column") args.urlColumn = argv[++i] || "review_url";
    else if (arg === "--combined-xlsx") args.combinedXlsx = path.resolve(argv[++i] || "");
    else if (arg === "--limit") args.limit = Number(argv[++i] || 0);
    else if (arg === "--txt-file") args.txtFiles.push(path.resolve(argv[++i] || ""));
    else if (arg === "--out-dir") args.outDir = path.resolve(argv[++i] || defaultOutDir);
    else if (arg === "--scrolls") args.scrolls = Number(argv[++i] || 0);
    else if (arg === "--no-auto-scroll") args.autoScroll = false;
    else if (arg === "--fast-threshold") args.fastThreshold = Number(argv[++i] || 0);
    else if (arg === "--fast-max-scrolls") args.fastMaxScrolls = Number(argv[++i] || 0);
    else if (arg === "--fast-stall-wait-ms") args.fastStallWaitMs = Number(argv[++i] || 0);
    else if (arg === "--max-scrolls") args.maxScrolls = Number(argv[++i] || 0);
    else if (arg === "--target-coverage") args.targetCoverage = Number(argv[++i] || 0);
    else if (arg === "--target-review-cap") args.targetReviewCap = Number(argv[++i] || 0);
    else if (arg === "--resource-mode") args.resourceMode = argv[++i] || "normal";
    else if (arg === "--resource-stats" || arg === "--verbose") args.resourceStats = true;
    else if (arg === "--stall-limit") args.stallLimit = Number(argv[++i] || 0);
    else if (arg === "--stall-wait-ms") args.stallWaitMs = Number(argv[++i] || 0);
    else if (arg === "--scroll-delay-ms") args.scrollDelayMs = Number(argv[++i] || 0);
    else if (arg === "--wait-ms") args.waitMs = Number(argv[++i] || 0);
    else if (arg === "--screenshot") args.screenshot = true;
    else if (arg === "--preview") args.preview = true;
    else if (arg === "--headed") args.headed = true;
    else args.urls.push(arg);
  }
  if (!["normal", "lite"].includes(args.resourceMode)) {
    throw new Error(`--resource-mode must be normal or lite: ${args.resourceMode}`);
  }
  return args;
}

const liteBlockedResourceTypes = new Set(["image", "media", "font"]);

function createResourceStats() {
  return { requested: {}, blocked: {} };
}

function incrementResourceCount(bucket, type) {
  const key = String(type || "unknown");
  bucket[key] = (bucket[key] || 0) + 1;
}

function cloneResourceCounts(counts) {
  return Object.fromEntries(Object.entries(counts || {}).map(([key, value]) => [key, Number(value) || 0]));
}

function diffResourceCounts(after, before) {
  const diff = {};
  for (const key of new Set([...Object.keys(after || {}), ...Object.keys(before || {})])) {
    const value = (after?.[key] || 0) - (before?.[key] || 0);
    if (value) diff[key] = value;
  }
  return diff;
}

function snapshotResourceStats(stats) {
  return {
    requested: cloneResourceCounts(stats.requested),
    blocked: cloneResourceCounts(stats.blocked),
  };
}

function diffResourceStats(after, before) {
  return {
    requested: diffResourceCounts(after.requested, before.requested),
    blocked: diffResourceCounts(after.blocked, before.blocked),
  };
}

function mergeResourceCounts(total, counts) {
  for (const [type, count] of Object.entries(counts || {})) {
    total[type] = (total[type] || 0) + count;
  }
}

function formatResourceCounts(counts) {
  const entries = Object.entries(counts || {}).filter(([, count]) => count);
  if (!entries.length) return "none";
  return entries
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([type, count]) => `${type}:${count}`)
    .join(", ");
}

async function configureResourceMode(context, args, stats) {
  context.on("request", (request) => {
    incrementResourceCount(stats.requested, request.resourceType());
  });

  if (args.resourceMode !== "lite") return;

  await context.route("**/*", async (route) => {
    const type = route.request().resourceType();
    if (liteBlockedResourceTypes.has(type)) {
      incrementResourceCount(stats.blocked, type);
      await route.abort();
      return;
    }
    await route.continue();
  });
}
async function loadPlaywright() {
  const require = createRequire(import.meta.url);
  const candidates = [
    process.env.PLAYWRIGHT_MODULE,
    path.join(toolDir, "node_modules", "playwright-core", "index.js"),
    path.join(toolDir, "node_modules", "playwright", "index.js"),
    path.join(workspaceDir, "node_modules", "playwright-core", "index.js"),
    path.join(workspaceDir, "node_modules", "playwright", "index.js"),
    path.join(workspaceDir, "node_modules", ".pnpm", "node_modules", "playwright-core", "index.js"),
    path.join(workspaceDir, "node_modules", ".pnpm", "node_modules", "playwright", "index.js"),
    path.join(process.env.USERPROFILE || "", ".cache", "codex-runtimes", "codex-primary-runtime", "dependencies", "node", "node_modules", "playwright", "index.js"),
  ].filter(Boolean);

  for (const candidate of candidates) {
    try {
      await fs.access(candidate);
      return require(candidate);
    } catch {
      // Try the next known location.
    }
  }

  throw new Error("Playwright를 찾지 못했습니다. npm install playwright-core 또는 npm install playwright가 필요합니다.");
}

function sanitizeFilePart(value) {
  return String(value || "page")
    .replace(/^https?:\/\//i, "")
    .replace(/[\\/:*?"<>|#&=%]+/g, "_")
    .replace(/\s+/g, "_")
    .slice(0, 120)
    .replace(/^_+|_+$/g, "") || "page";
}

function placeIdFromUrl(url) {
  return String(url).match(/place\.map\.kakao\.com\/(\d+)/)?.[1] || "";
}

function normalizePlaceUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const match = raw.match(/https?:\/\/place\.map\.kakao\.com\/\d+(?:#review)?/i);
  if (!match) return "";
  const base = match[0].replace(/^http:/i, "https:").replace(/#.*$/, "");
  return `${base}#review`;
}

async function readUrlsFromXlsx(xlsxPath, columnName = "review_url") {
  const input = await FileBlob.load(xlsxPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const sheet = workbook.worksheets.getItemAt(0);
  const values = sheet.getUsedRange(true).values || [];
  if (!values.length) return [];

  const headers = values[0].map((value) => String(value || "").trim());
  const preferred = headers.findIndex((header) => header.toLowerCase() === columnName.toLowerCase());
  const fallback = headers.findIndex((header) => /review.*url|place.*url|url/i.test(header));
  const colIndex = preferred >= 0 ? preferred : fallback;

  const urls = [];
  if (colIndex >= 0) {
    for (const row of values.slice(1)) {
      const url = normalizePlaceUrl(row[colIndex]);
      if (url) urls.push(url);
    }
  } else {
    for (const row of values) {
      for (const cell of row) {
        const url = normalizePlaceUrl(cell);
        if (url) urls.push(url);
      }
    }
  }
  return urls;
}

async function readUrls(args) {
  const urls = args.urls.map(normalizePlaceUrl).filter(Boolean);
  if (args.urlFile) {
    const text = await fs.readFile(args.urlFile, "utf8");
    urls.push(...text.split(/\r?\n/).map((line) => normalizePlaceUrl(line)).filter(Boolean));
  }
  for (const xlsxFile of args.xlsxFiles) {
    const xlsxUrls = await readUrlsFromXlsx(xlsxFile, args.urlColumn);
    console.log(`XLSX URL ${xlsxUrls.length}개 읽음: ${xlsxFile}`);
    urls.push(...xlsxUrls);
  }
  const unique = [...new Set(urls)];
  return args.limit > 0 ? unique.slice(0, args.limit) : unique;
}

async function readVisiblePage(page) {
  return page.evaluate(() => {
    const title = document.title || "";
    const text = document.body?.innerText || "";
    const links = [...document.querySelectorAll("a[href]")]
      .map((a) => ({ text: (a.innerText || "").trim(), href: a.href }))
      .filter((a) => a.text || a.href);
    return { title, text, links };
  });
}

function summaryFromVisibleText(url, data) {
  const txt = [`URL: ${url}`, `Title: ${data.title || ""}`, "", data.text || ""].join("\n");
  return parsePlaceText(txt);
}

function captureProgress(summary, args, settings = null) {
  const total = Number(summary.totalReviews) || 0;
  const parsed = Number(summary.parsedReviewCount) || 0;
  const targetCoverage = settings ? settings.targetCoverage : Math.max(0, Math.min(1, Number(args.targetCoverage) || 0));
  const targetReviewCap = Math.max(0, Number(settings ? settings.targetReviewCap : args.targetReviewCap) || 0);
  const unavailable = Boolean(summary.reviewUnavailable);
  const rawTarget = total ? Math.ceil(total * targetCoverage) : 0;
  const target = rawTarget && targetReviewCap ? Math.min(rawTarget, targetReviewCap) : rawTarget;
  const coverage = total ? parsed / total : "";
  return { total, parsed, target, coverage, unavailable };
}


function scrollSettingsFor(progress, args) {
  const total = Number(progress.total) || 0;
  const threshold = Math.max(0, Number(args.fastThreshold) || 0);
  const useFast = args.autoScroll && total > 0 && total <= threshold;
  const minScrolls = Math.max(0, Number(args.scrolls) || 0);
  if (useFast) {
    return {
      profile: "fast",
      minScrolls,
      maxScrolls: Math.max(minScrolls, Number(args.fastMaxScrolls) || minScrolls),
      stallLimit: 2,
      stallWait: Math.max(0, Number(args.fastStallWaitMs) || 0),
      delay: 300,
      targetCoverage: 1.0,
      targetReviewCap: Math.max(0, Number(args.targetReviewCap) || 0),
    };
  }
  return {
    profile: args.autoScroll ? "slow" : "fixed",
    minScrolls,
    maxScrolls: args.autoScroll ? Math.max(minScrolls, Number(args.maxScrolls) || minScrolls) : minScrolls,
    stallLimit: Math.max(1, Number(args.stallLimit) || 1),
    stallWait: Math.max(0, Number(args.stallWaitMs) || 0),
    delay: Math.max(100, Number(args.scrollDelayMs) || 300),
    targetCoverage: Math.max(0, Math.min(1, Number(args.targetCoverage) || 0)),
    targetReviewCap: Math.max(0, Number(args.targetReviewCap) || 0),
  };
}
async function capturePage(page, url, args, resourceStats = null) {
  const startedAt = Date.now();
  const resourceBefore = resourceStats ? snapshotResourceStats(resourceStats) : createResourceStats();
  await page.goto(url, { waitUntil: "domcontentloaded", timeout: 45000 });
  await page.waitForTimeout(args.waitMs);

  let data = await readVisiblePage(page);
  let summary = summaryFromVisibleText(url, data);
  let progress = captureProgress(summary, args);
  const settings = scrollSettingsFor(progress, args);
  progress = captureProgress(summary, args, settings);
  let previousParsed = progress.parsed;
  let stalled = 0;
  let scrollCount = 0;

  for (let i = 0; i < settings.maxScrolls; i++) {
    const enough = progress.target > 0 && progress.parsed >= progress.target;
    const canStop = progress.unavailable || (i >= settings.minScrolls && (enough || stalled >= settings.stallLimit));
    if (canStop) break;

    await page.mouse.wheel(0, 1400);
    scrollCount += 1;
    await page.waitForTimeout(settings.delay);

    data = await readVisiblePage(page);
    summary = summaryFromVisibleText(url, data);
    progress = captureProgress(summary, args, settings);
    if (progress.parsed <= previousParsed && settings.stallWait > 0 && !progress.unavailable) {
      await page.waitForTimeout(settings.stallWait);
      data = await readVisiblePage(page);
      summary = summaryFromVisibleText(url, data);
      progress = captureProgress(summary, args, settings);
    }
    if (progress.parsed <= previousParsed) stalled += 1;
    else stalled = 0;
    previousParsed = progress.parsed;
  }

  const status = progress.unavailable
    ? "후기미제공"
    : progress.total
      ? progress.parsed >= progress.target
        ? "target_reached"
        : scrollCount >= settings.maxScrolls
          ? "max_scrolls_reached"
          : "stalled"
      : "unknown_total";
  const coverageText = progress.total ? `${progress.parsed}/${progress.total} (${(progress.coverage * 100).toFixed(1)}%)` : `${progress.parsed}/?`;
  const durationMs = Date.now() - startedAt;
  const resources = resourceStats ? diffResourceStats(snapshotResourceStats(resourceStats), resourceBefore) : createResourceStats();
  console.log(`  reviews: ${coverageText}, scrolls ${scrollCount}, status ${status}, profile ${settings.profile}, duration ${durationMs}ms`);
  if (args.resourceStats) {
    console.log(`  resources requested: ${formatResourceCounts(resources.requested)} / blocked: ${formatResourceCounts(resources.blocked)}`);
  }

  return {
    url,
    capturedAt: new Date().toISOString(),
    title: data.title,
    text: data.text.replace(/\n{3,}/g, "\n\n").trim(),
    links: data.links,
    captureStats: {
      totalReviews: progress.total,
      parsedReviews: progress.parsed,
      targetReviews: progress.target,
      coverage: progress.coverage,
      scrollCount,
      maxScrolls: settings.maxScrolls,
      profile: settings.profile,
      status,
      durationMs,
      resourceMode: args.resourceMode,
      resources,
    },
  };
}
function findBrowserExecutable() {
  const candidates = [
    process.env.CHROME_PATH,
    "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe",
    "C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe",
    "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe",
  ].filter(Boolean);
  for (const candidate of candidates) {
    try {
      accessSync(candidate);
      return candidate;
    } catch {
      // Try next browser path.
    }
  }
  return undefined;
}

function cleanLines(text) {
  return String(text || "")
    .split(/\r?\n/)
    .map((line) => line.replace(/\u00a0/g, " ").trim())
    .filter(Boolean);
}

function numericAfter(lines, label, fromIndex = 0) {
  const index = lines.findIndex((line, i) => i >= fromIndex && line === label);
  if (index < 0) return "";
  const value = lines[index + 1] || "";
  return /^[0-9]+(?:\.[0-9]+)?$/.test(value) ? Number(value) : value;
}

function firstAfter(lines, label, fromIndex = 0) {
  const index = lines.findIndex((line, i) => i >= fromIndex && line === label);
  return index >= 0 ? (lines[index + 1] || "") : "";
}

function firstIndexOfAny(lines, labels, fromIndex = 0) {
  const indexes = labels
    .map((label) => lines.findIndex((line, i) => i >= fromIndex && line === label))
    .filter((index) => index >= 0);
  return indexes.length ? Math.min(...indexes) : -1;
}

function lineHasReviewUnavailable(line) {
  return /후기\s*미제공|후기미제공|후기가 제공되지 않는 장소/.test(String(line || ""));
}

function parseTopReviewTotal(lines, categoryIndex) {
  const start = Math.max(0, categoryIndex);
  const reviewListIndex = lines.findIndex((line, i) => i >= start && line === "후기 목록");
  const boundary = firstIndexOfAny(lines, ["장소 기본 정보", "업로드 된 사진 수", "이 지역 길찾기 랭킹", "블로그명", "서비스 이용정보"], start);
  const endCandidates = [reviewListIndex, boundary].filter((index) => index >= 0);
  const end = endCandidates.length ? Math.min(...endCandidates) : Math.min(lines.length, start + 80);
  const top = lines.slice(start, end);
  if (top.some(lineHasReviewUnavailable)) return { totalReviews: "후기미제공", reviewUnavailable: true };

  const reviewLabelIndex = lines.findIndex((line, i) => i >= start && i < end && line === "후기");
  if (reviewLabelIndex >= 0 && /^[0-9]+$/.test(lines[reviewLabelIndex + 1] || "")) {
    return { totalReviews: Number(lines[reviewLabelIndex + 1]), reviewUnavailable: false };
  }
  const combinedReviewLine = top.find((line) => /^후기\s+\d+/.test(line));
  if (combinedReviewLine) {
    return { totalReviews: Number(combinedReviewLine.match(/^후기\s+(\d+)/)?.[1] || 0), reviewUnavailable: false };
  }
  return { totalReviews: "", reviewUnavailable: false };
}

function parseReviewText(segment, dateIndex) {
  const stopLabels = new Set([
    "사진 목록",
    "근처 사진",
    "위치기반",
    "좋아요 개수,",
    "좋아요를 누른 사용자",
    "메뉴 더보기",
  ]);
  const tagLabels = new Set(["맛", "친절", "분위기", "주차", "가성비"]);
  const parts = [];
  for (let i = dateIndex + 1; i < segment.length; i++) {
    const line = segment[i];
    if (stopLabels.has(line)) break;
    if (tagLabels.has(line)) continue;
    if (/^\+\d+$/.test(line)) continue;
    parts.push(line);
  }
  return parts.join("\n").trim();
}

function parsePlaceText(fullText, sourcePath = "") {
  const lines = cleanLines(fullText);
  const url = firstAfter(lines, "URL:") || (lines.find((line) => line.startsWith("URL: ")) || "").replace(/^URL:\s*/, "");
  const title = (lines.find((line) => line.startsWith("Title: ")) || "").replace(/^Title:\s*/, "");
  const placeName = firstAfter(lines, "장소명") || title.replace(/\s*\|\s*카카오맵$/, "");
  const category = firstAfter(lines, "장소 카테고리");
  const categoryIndex = lines.findIndex((line) => line === "장소 카테고리");
  const storeRating = numericAfter(lines, "별점", Math.max(0, categoryIndex));
  const topReview = parseTopReviewTotal(lines, categoryIndex);
  const totalReviews = topReview.totalReviews;
  const reviewUnavailable = topReview.reviewUnavailable;
  const blogLine = lines.find((line) => /^블로그\s*\d+/.test(line));
  const blogReviews = blogLine ? Number(blogLine.match(/^블로그\s*(\d+)/)?.[1] || 0) : "";

  const startIndex = Math.max(0, lines.findIndex((line) => line === "후기 목록"));
  const reviewerIndexes = [];
  for (let i = startIndex; i < lines.length; i++) {
    if (lines[i] === "리뷰어 이름,") reviewerIndexes.push(i);
  }

  const reviews = [];
  for (let r = 0; r < reviewerIndexes.length; r++) {
    const start = reviewerIndexes[r];
    const end = reviewerIndexes[r + 1] || lines.length;
    const segment = lines.slice(start, end);
    const statsLine = segment.find((line) => /후기\s*\d+별점평균\s*[0-9.]+팔로워\s*\d+/.test(line)) || "";
    const stats = statsLine.match(/후기\s*(\d+)별점평균\s*([0-9.]+)팔로워\s*(\d+)/);
    const ratingLabelIndex = segment.findIndex((line, i) => i > 0 && line === "별점");
    const reviewRating = ratingLabelIndex >= 0 ? Number(segment[ratingLabelIndex + 1] || "") : "";
    const dateIndex = segment.findIndex((line) => /^\d{4}\.\d{2}\.\d{2}\.$/.test(line));
    const reviewDate = dateIndex >= 0 ? segment[dateIndex].replace(/\.$/, "").replace(/\./g, "-") : "";
    if (!stats && reviewRating === "" && !reviewDate) continue;
    reviews.push({
      anonymousReviewer: `익명${reviews.length + 1}`,
      reviewerReviewCount: stats ? Number(stats[1]) : "",
      reviewerAverageRating: stats ? Number(stats[2]) : "",
      reviewerFollowerCount: stats ? Number(stats[3]) : "",
      reviewRating,
      reviewDate,
      reviewText: dateIndex >= 0 ? parseReviewText(segment, dateIndex) : "",
    });
  }

  return {
    sourcePath,
    url,
    title,
    placeName,
    category,
    storeRating,
    totalReviews,
    reviewUnavailable,
    blogReviews,
    parsedReviewCount: reviews.length,
    reviews,
  };
}

async function saveCombinedWorkbook(summaries, xlsxPath) {
  const workbook = Workbook.create();
  const summarySheet = workbook.worksheets.add("식당요약");
  const reviewSheet = workbook.worksheets.add("리뷰전체");
  summarySheet.showGridLines = false;
  reviewSheet.showGridLines = false;

  const summaryHeaders = ["식당번호", "장소명", "카테고리", "식당 평균 별점", "식당 총 리뷰수", "캡처된 리뷰수", "수집률", "수집상태", "수집프로파일", "스크롤횟수", "목표리뷰수", "블로그 수", "원본 URL", "원본 TXT"];
  const summaryRows = summaries.map((summary, index) => [
    index + 1,
    summary.placeName || "",
    summary.category || "",
    summary.storeRating || "",
    summary.totalReviews || "",
    summary.parsedReviewCount || 0,
    summary.captureStats?.coverage === "" || summary.captureStats?.coverage == null ? "" : summary.captureStats.coverage,
    summary.captureStats?.status || (summary.reviewUnavailable ? "후기미제공" : ""),
    summary.captureStats?.profile || "",
    summary.captureStats?.scrollCount ?? "",
    summary.captureStats?.targetReviews || "",
    summary.blogReviews || "",
    summary.url || "",
    summary.sourcePath || "",
  ]);
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

  const reviewHeaders = ["식당번호", "장소명", "식당 평균 별점", "식당 총 리뷰수", "익명 리뷰어", "리뷰 날짜", "리뷰어 평균 별점", "이 식당에 준 별점", "리뷰어 총 리뷰수", "팔로워 수", "리뷰 내용", "원본 URL"];
  const reviewRows = [];
  for (let s = 0; s < summaries.length; s++) {
    const summary = summaries[s];
    for (const review of summary.reviews) {
      reviewRows.push([
        s + 1,
        summary.placeName || "",
        summary.storeRating || "",
        summary.totalReviews || "",
        review.anonymousReviewer,
        review.reviewDate,
        review.reviewerAverageRating,
        review.reviewRating,
        review.reviewerReviewCount,
        review.reviewerFollowerCount,
        review.reviewText,
        summary.url || "",
      ]);
    }
  }
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
  reviewSheet.getRange("K:K").format.columnWidthPx = 430;
  reviewSheet.getRange("L:L").format.columnWidthPx = 360;
  reviewSheet.getRange("A:L").format = { verticalAlignment: "top" };

  await fs.mkdir(path.dirname(xlsxPath), { recursive: true });
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(xlsxPath);
}

async function saveReviewWorkbook(summary, xlsxPath, options = {}) {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("리뷰요약");
  sheet.showGridLines = false;

  sheet.getRange("A1:I1").merge();
  sheet.getRange("A1").values = [[`${summary.placeName || "카카오맵 장소"} 리뷰 요약`]];
  sheet.getRange("A1").format = { fill: "#1F4E78", font: { bold: true, color: "#FFFFFF", size: 16 } };

  const summaryRows = [
    ["장소명", summary.placeName || "", "카테고리", summary.category || "", "원본 URL", summary.url || ""],
    ["식당 평균 별점", summary.storeRating || "", "식당 총 리뷰수", summary.totalReviews || "", "캡처된 리뷰수", summary.parsedReviewCount],
    ["블로그 수", summary.blogReviews || "", "원본 TXT", summary.sourcePath || "", "", ""],
  ];
  sheet.getRangeByIndexes(2, 0, summaryRows.length, 6).values = summaryRows;
  sheet.getRange("A3:A5").format = { fill: "#D9EAF7", font: { bold: true } };
  sheet.getRange("C3:C5").format = { fill: "#D9EAF7", font: { bold: true } };
  sheet.getRange("E3:E5").format = { fill: "#D9EAF7", font: { bold: true } };
  sheet.getRange("B4:D4").format.numberFormat = "0.0";

  const headers = ["익명 리뷰어", "리뷰 날짜", "리뷰어 평균 별점", "이 식당에 준 별점", "리뷰어 총 리뷰수", "팔로워 수", "리뷰 내용", "장소명", "원본 URL"];
  const rows = summary.reviews.map((review) => [
    review.anonymousReviewer,
    review.reviewDate,
    review.reviewerAverageRating,
    review.reviewRating,
    review.reviewerReviewCount,
    review.reviewerFollowerCount,
    review.reviewText,
    summary.placeName || "",
    summary.url || "",
  ]);
  const table = [headers, ...rows];
  const startRow = 7;
  sheet.getRangeByIndexes(startRow - 1, 0, table.length, headers.length).values = table;
  sheet.getRangeByIndexes(startRow - 1, 0, 1, headers.length).format = { fill: "#244062", font: { bold: true, color: "#FFFFFF" } };
  if (rows.length) {
    sheet.tables.add(`A${startRow}:I${startRow + rows.length}`, true, "ReviewTable");
    sheet.getRange(`C${startRow + 1}:D${startRow + rows.length}`).format.numberFormat = "0.0";
    sheet.getRange(`E${startRow + 1}:F${startRow + rows.length}`).format.numberFormat = "#,##0";
    sheet.getRange(`G${startRow + 1}:G${startRow + rows.length}`).format.wrapText = true;
  }
  sheet.freezePanes.freezeRows(startRow);
  sheet.getRange("A:A").format.columnWidthPx = 95;
  sheet.getRange("B:B").format.columnWidthPx = 95;
  sheet.getRange("C:D").format.columnWidthPx = 120;
  sheet.getRange("E:F").format.columnWidthPx = 115;
  sheet.getRange("G:G").format.columnWidthPx = 420;
  sheet.getRange("H:H").format.columnWidthPx = 150;
  sheet.getRange("I:I").format.columnWidthPx = 360;
  sheet.getRange("A:I").format = { verticalAlignment: "top" };

  if (options.preview) {
    const preview = await workbook.render({ sheetName: "리뷰요약", autoCrop: "all", scale: 1, format: "png" });
    await fs.writeFile(xlsxPath.replace(/\.xlsx$/i, ".preview.png"), new Uint8Array(await preview.arrayBuffer()));
  }
  const output = await SpreadsheetFile.exportXlsx(workbook);
  await output.save(xlsxPath);
}

async function saveCapturedArtifacts(captured, args, page = null) {
  const id = placeIdFromUrl(captured.url);
  const base = sanitizeFilePart(id ? `${id}_${captured.title}` : captured.title || captured.url);
  const txtPath = path.join(args.outDir, `${base}.txt`);
  const jsonPath = path.join(args.outDir, `${base}.json`);
  const xlsxPath = path.join(args.outDir, `${base}_reviews.xlsx`);
  const screenshotPath = path.join(args.outDir, `${base}.png`);

  const txt = [
    `URL: ${captured.url}`,
    `Captured At: ${captured.capturedAt}`,
    `Title: ${captured.title}`,
    "",
    captured.text,
  ].join("\n");

  await fs.writeFile(txtPath, txt, "utf8");
  await fs.writeFile(jsonPath, JSON.stringify(captured, null, 2), "utf8");
  if (args.screenshot && page) await page.screenshot({ path: screenshotPath, fullPage: true });
  const summary = parsePlaceText(txt, txtPath);
  summary.captureStats = captured.captureStats || {};
  if (!args.combinedXlsx) await saveReviewWorkbook(summary, xlsxPath, args);
  return {
    url: captured.url,
    title: captured.title,
    txtPath,
    jsonPath,
    xlsxPath: args.combinedXlsx ? "" : xlsxPath,
    screenshotPath: args.screenshot ? screenshotPath : "",
    textLength: captured.text.length,
    parsedReviewCount: summary.parsedReviewCount,
    totalReviews: summary.totalReviews || "",
    coverage: summary.captureStats.coverage ?? "",
    captureStatus: summary.captureStats.status || "",
    scrollCount: summary.captureStats.scrollCount ?? "",
    durationMs: summary.captureStats.durationMs ?? "",
    resourceMode: summary.captureStats.resourceMode || args.resourceMode,
    resourceRequests: summary.captureStats.resources?.requested || {},
    resourceBlocked: summary.captureStats.resources?.blocked || {},
    summary,
  };
}

async function saveFromTxt(txtFile, args) {
  const txt = await fs.readFile(txtFile, "utf8");
  const summary = parsePlaceText(txt, txtFile);
  const base = sanitizeFilePart(path.basename(txtFile, path.extname(txtFile)));
  const xlsxPath = path.join(args.outDir, `${base}_reviews.xlsx`);
  if (!args.combinedXlsx) await saveReviewWorkbook(summary, xlsxPath, args);
  return { url: summary.url, title: summary.title, txtPath: txtFile, jsonPath: "", xlsxPath: args.combinedXlsx ? "" : xlsxPath, screenshotPath: "", textLength: txt.length, parsedReviewCount: summary.parsedReviewCount, summary };
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const urls = await readUrls(args);
  if (!urls.length && !args.txtFiles.length) {
    usage();
    process.exitCode = 2;
    return;
  }

  await fs.mkdir(args.outDir, { recursive: true });
  const indexRows = [];
  const combinedSummaries = [];

  for (let i = 0; i < args.txtFiles.length; i++) {
    const txtFile = args.txtFiles[i];
    console.log(`TXT ${i + 1}/${args.txtFiles.length} ${txtFile}`);
    const row = await saveFromTxt(txtFile, args);
    indexRows.push(row);
    if (row.summary) combinedSummaries.push(row.summary);
  }

  if (urls.length) {
    if (args.screenshot && args.resourceMode === "lite") {
      console.warn("Warning: --screenshot with --resource-mode lite may produce incomplete screenshots because images are blocked.");
    }
    const { chromium } = await loadPlaywright();
    const browser = await chromium.launch({ headless: !args.headed, executablePath: findBrowserExecutable() });
    const contextOptions = { viewport: { width: 1280, height: 1600 } };
    if (args.resourceMode === "lite") contextOptions.serviceWorkers = "block";
    const context = await browser.newContext(contextOptions);
    const resourceStats = createResourceStats();
    const totalResources = createResourceStats();
    await configureResourceMode(context, args, resourceStats);
    const page = await context.newPage();
    try {
      for (let i = 0; i < urls.length; i++) {
        const url = urls[i];
        console.log(`${i + 1}/${urls.length} ${url}`);
        const captured = await capturePage(page, url, args, resourceStats);
        mergeResourceCounts(totalResources.requested, captured.captureStats?.resources?.requested);
        mergeResourceCounts(totalResources.blocked, captured.captureStats?.resources?.blocked);
        const row = await saveCapturedArtifacts(captured, args, page);
        indexRows.push(row);
        if (row.summary) combinedSummaries.push(row.summary);
      }
      if (args.resourceStats) {
        console.log(`resource totals (${args.resourceMode}): requested ${formatResourceCounts(totalResources.requested)} / blocked ${formatResourceCounts(totalResources.blocked)}`);
      }
    } finally {
      await context.close();
      await browser.close();
    }
  }

  if (args.combinedXlsx) {
    await saveCombinedWorkbook(combinedSummaries, args.combinedXlsx);
    console.log(`통합 XLSX: ${args.combinedXlsx} (${combinedSummaries.length} places)`);
  }

  for (const row of indexRows) delete row.summary;
  const indexPath = path.join(args.outDir, `index_${new Date().toISOString().replace(/[:.]/g, "-")}.json`);
  await fs.writeFile(indexPath, JSON.stringify(indexRows, null, 2), "utf8");
  console.log(`저장 완료: ${args.outDir}`);
  console.log(`인덱스: ${indexPath}`);
  for (const row of indexRows) {
    if (row.xlsxPath) console.log(`XLSX: ${row.xlsxPath} (${row.parsedReviewCount} reviews)`);
    else console.log(`처리: ${row.url} (${row.parsedReviewCount} reviews)`);
  }
}

main().catch((err) => {
  console.error(err);
  process.exit(1);
});

















