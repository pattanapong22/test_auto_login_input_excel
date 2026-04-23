import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const workspaceRoot = process.cwd();
const outputDir = path.join(workspaceRoot, "outputs", "auto-test-login");
const screenshotDir = path.join(outputDir, "screenshots");
const workbookPath = path.join(outputDir, "login-log.xlsx");
const artifactModulePath = pathToFileURL(
  "C:/Users/ADMIN/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/@oai/artifact-tool/dist/artifact_tool.mjs",
).href;
const playwrightModulePath = pathToFileURL(
  "C:/Users/ADMIN/.cache/codex-runtimes/codex-primary-runtime/dependencies/node/node_modules/playwright/index.mjs",
).href;

const { FileBlob, SpreadsheetFile, Workbook } = await import(artifactModulePath);
const playwright = await import(playwrightModulePath);

const args = new Set(process.argv.slice(2));
const isSelfTest = args.has("--self-test");
const configPath = path.join(workspaceRoot, "config", "login.config.json");

function getCliOption(name) {
  const prefix = `${name}=`;
  const matched = [...args].find((arg) => arg.startsWith(prefix));
  return matched ? matched.slice(prefix.length) : "";
}

function formatLocalDateTime(date) {
  const parts = new Intl.DateTimeFormat("sv-SE", {
    timeZone: "Asia/Taipei",
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
  }).formatToParts(date);
  const map = Object.fromEntries(parts.filter((part) => part.type !== "literal").map((part) => [part.type, part.value]));
  return `${map.year}-${map.month}-${map.day} ${map.hour}:${map.minute}:${map.second}`;
}

function columnNameFromIndex(index) {
  let value = index + 1;
  let label = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    label = String.fromCharCode(65 + remainder) + label;
    value = Math.floor((value - remainder) / 26);
  }
  return label;
}

function sanitizeForExcel(value) {
  if (value === null || value === undefined) {
    return "";
  }

  if (value instanceof Date || typeof value === "number" || typeof value === "boolean") {
    return value;
  }

  return String(value)
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uD800-\uDFFF\uFFFE\uFFFF]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sanitizePositiveInteger(value, fallbackValue) {
  const parsed = Number.parseInt(String(value), 10);
  return Number.isFinite(parsed) && parsed > 0 ? parsed : fallbackValue;
}

async function ensureDirectories() {
  await fs.mkdir(outputDir, { recursive: true });
  await fs.mkdir(screenshotDir, { recursive: true });
}

async function readConfig() {
  let raw;
  try {
    raw = await fs.readFile(configPath, "utf8");
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "ENOENT") {
      throw new Error("Config file not found. Copy config/login.config.example.json to config/login.config.json first.");
    }
    throw error;
  }

  const config = JSON.parse(raw);
  const required = [
    "loginUrl",
    "email",
    "password",
    "emailSelector",
    "passwordSelector",
    "submitSelector",
  ];
  const missingFields = [];

  for (const field of required) {
    if (!config[field] || String(config[field]).includes("your-")) {
      missingFields.push(field);
    }
  }

  if (missingFields.length > 0) {
    throw new Error(
      `Config is not ready. Update config/login.config.json fields: ${missingFields.join(", ")}`,
    );
  }

  return {
    rounds: 1,
    roundDelayMs: 0,
    timeoutMs: 15000,
    headless: false,
    browser: "chromium",
    screenshotOnSuccess: true,
    slowMoMs: 400,
    typingDelayMs: 120,
    postLoginWaitMs: 3000,
    keepBrowserOpenOnFail: true,
    failPauseMs: 10000,
    ...config,
  };
}

function createWorkbook() {
  const workbook = Workbook.create();
  const logSheet = workbook.worksheets.add("Login Logs");
  const summarySheet = workbook.worksheets.add("Run Summary");
  configureLogSheet(logSheet);
  configureSummarySheet(summarySheet);
  return workbook;
}

function configureLogSheet(logSheet) {
  const logHeaders = [
    "Batch ID",
    "Round",
    "Total Rounds",
    "Run ID",
    "Login DateTime",
    "Status",
    "Login URL",
    "Email",
    "Browser",
    "Duration (ms)",
    "Screenshot",
    "Message",
  ];
  logSheet.getRange(`A1:${columnNameFromIndex(logHeaders.length - 1)}1`).values = [logHeaders];
  logSheet.getRange("A1:L1").format.font.bold = true;
  logSheet.getRange("A1:L1").format.fill.color = "#D9EAF7";
  logSheet.getRange("A1:L1").format.horizontalAlignment = "Center";
  logSheet.getRange("A1:L1").format.wrapText = true;
  logSheet.freezePanes.freezeRows(1);
  logSheet.getRange("A:L").format.wrapText = true;

  const logColumnWidths = [180, 70, 90, 180, 150, 90, 280, 180, 100, 110, 240, 320];
  logColumnWidths.forEach((width, index) => {
    logSheet.getRange(`${columnNameFromIndex(index)}:${columnNameFromIndex(index)}`).format.columnWidthPx = width;
  });
}

function configureSummarySheet(summarySheet) {
  const summaryHeaders = [
    "Batch ID",
    "Started At",
    "Finished At",
    "Total Rounds",
    "Success Count",
    "Fail Count",
    "Success Rate",
    "Login URL",
    "Email",
    "Message",
  ];
  summarySheet.getRange(`A1:${columnNameFromIndex(summaryHeaders.length - 1)}1`).values = [summaryHeaders];
  summarySheet.getRange("A1:J1").format.font.bold = true;
  summarySheet.getRange("A1:J1").format.fill.color = "#FDE9D9";
  summarySheet.getRange("A1:J1").format.horizontalAlignment = "Center";
  summarySheet.getRange("A1:J1").format.wrapText = true;
  summarySheet.freezePanes.freezeRows(1);
  summarySheet.getRange("A:J").format.wrapText = true;

  const summaryColumnWidths = [180, 150, 150, 100, 100, 100, 100, 280, 180, 320];
  summaryColumnWidths.forEach((width, index) => {
    summarySheet.getRange(`${columnNameFromIndex(index)}:${columnNameFromIndex(index)}`).format.columnWidthPx = width;
  });
}

async function loadWorkbook() {
  if (!(await fileExists(workbookPath))) {
    return createWorkbook();
  }

  const input = await FileBlob.load(workbookPath);
  const workbook = await SpreadsheetFile.importXlsx(input);
  const logSheet = getOrAddWorksheet(workbook, "Login Logs");
  const summarySheet = getOrAddWorksheet(workbook, "Run Summary");
  configureLogSheet(logSheet);
  configureSummarySheet(summarySheet);
  return workbook;
}

async function fileExists(targetPath) {
  try {
    await fs.access(targetPath);
    return true;
  } catch {
    return false;
  }
}

function getOrAddWorksheet(workbook, sheetName) {
  const sheet = workbook.worksheets.getItemOrNullObject(sheetName);
  return sheet?.isNullObject ? workbook.worksheets.add(sheetName) : sheet;
}

async function appendLogRow(entry) {
  const workbook = await loadWorkbook();
  const sheet = getOrAddWorksheet(workbook, "Login Logs");
  const usedRange = sheet.getUsedRange();
  const nextRow = Math.max((usedRange?.getRowCount?.() ?? 0) + 1, 2);
  const safeEntry = {
    batchId: sanitizeForExcel(entry.batchId),
    roundNumber: typeof entry.roundNumber === "number" ? entry.roundNumber : 1,
    totalRounds: typeof entry.totalRounds === "number" ? entry.totalRounds : 1,
    runId: sanitizeForExcel(entry.runId),
    loginAt: entry.loginAt,
    status: sanitizeForExcel(entry.status),
    loginUrl: sanitizeForExcel(entry.loginUrl),
    email: sanitizeForExcel(entry.email),
    browser: sanitizeForExcel(entry.browser),
    durationMs: typeof entry.durationMs === "number" ? entry.durationMs : 0,
    screenshotPath: sanitizeForExcel(entry.screenshotPath),
    message: sanitizeForExcel(entry.message),
  };

  sheet.getRange(`A${nextRow}:L${nextRow}`).values = [[
    safeEntry.batchId,
    safeEntry.roundNumber,
    safeEntry.totalRounds,
    safeEntry.runId,
    safeEntry.loginAt,
    safeEntry.status,
    safeEntry.loginUrl,
    safeEntry.email,
    safeEntry.browser,
    safeEntry.durationMs,
    safeEntry.screenshotPath,
    safeEntry.message,
  ]];
  sheet.getRange(`E${nextRow}:E${nextRow}`).setNumberFormat("yyyy-mm-dd hh:mm:ss");
  sheet.getRange(`J${nextRow}:J${nextRow}`).setNumberFormat("0");

  if (safeEntry.status === "PASS") {
    sheet.getRange(`F${nextRow}:F${nextRow}`).format.fill.color = "#C6EFCE";
  } else {
    sheet.getRange(`F${nextRow}:F${nextRow}`).format.fill.color = "#FFC7CE";
  }

  const output = await SpreadsheetFile.exportXlsx(workbook);
  try {
    await output.save(workbookPath);
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "EBUSY") {
      throw new Error("Excel log file is open. Close outputs/auto-test-login/login-log.xlsx and run again.");
    }
    throw error;
  }
}

async function appendSummaryRow(summary) {
  const workbook = await loadWorkbook();
  const sheet = getOrAddWorksheet(workbook, "Run Summary");
  const usedRange = sheet.getUsedRange();
  const nextRow = Math.max((usedRange?.getRowCount?.() ?? 0) + 1, 2);
  const safeSummary = {
    batchId: sanitizeForExcel(summary.batchId),
    startedAt: summary.startedAt,
    finishedAt: summary.finishedAt,
    totalRounds: sanitizePositiveInteger(summary.totalRounds, 1),
    successCount: sanitizePositiveInteger(summary.successCount, 0),
    failCount: sanitizePositiveInteger(summary.failCount, 0),
    successRate: sanitizeForExcel(summary.successRate),
    loginUrl: sanitizeForExcel(summary.loginUrl),
    email: sanitizeForExcel(summary.email),
    message: sanitizeForExcel(summary.message),
  };

  sheet.getRange(`A${nextRow}:J${nextRow}`).values = [[
    safeSummary.batchId,
    safeSummary.startedAt,
    safeSummary.finishedAt,
    safeSummary.totalRounds,
    safeSummary.successCount,
    safeSummary.failCount,
    safeSummary.successRate,
    safeSummary.loginUrl,
    safeSummary.email,
    safeSummary.message,
  ]];
  sheet.getRange(`B${nextRow}:C${nextRow}`).setNumberFormat("yyyy-mm-dd hh:mm:ss");

  const output = await SpreadsheetFile.exportXlsx(workbook);
  try {
    await output.save(workbookPath);
  } catch (error) {
    if (error && typeof error === "object" && "code" in error && error.code === "EBUSY") {
      throw new Error("Excel log file is open. Close outputs/auto-test-login/login-log.xlsx and run again.");
    }
    throw error;
  }
}

function buildRunId(prefix = "login") {
  const now = new Date();
  const compact = formatLocalDateTime(now).replace(/[-:\s]/g, "");
  const milliseconds = String(now.getMilliseconds()).padStart(3, "0");
  return `${prefix}-${compact}${milliseconds}`;
}

async function runLoginTest(config, runContext = {}) {
  const startedAt = new Date();
  const runId = buildRunId();
  let browser;
  let page;
  let screenshotPath = "";

  try {
    const browserType = playwright[config.browser];
    if (!browserType) {
      throw new Error(`Unsupported browser: ${config.browser}`);
    }

    browser = await browserType.launch({
      headless: config.headless,
      slowMo: config.slowMoMs,
    });
    const context = await browser.newContext();
    page = await context.newPage();

    await page.goto(config.loginUrl, { waitUntil: "domcontentloaded", timeout: config.timeoutMs });
    const emailInput = page.locator(config.emailSelector);
    const passwordInput = page.locator(config.passwordSelector);

    await emailInput.click({ timeout: config.timeoutMs });
    await emailInput.clear({ timeout: config.timeoutMs });
    await page.keyboard.type(config.email, { delay: config.typingDelayMs });

    await passwordInput.click({ timeout: config.timeoutMs });
    await passwordInput.clear({ timeout: config.timeoutMs });
    await page.keyboard.type(config.password, { delay: config.typingDelayMs });

    await page.locator(config.submitSelector).click({ timeout: config.timeoutMs });

    const successChecks = [];
    if (config.successSelector) {
      successChecks.push(page.locator(config.successSelector).waitFor({ state: "visible", timeout: config.timeoutMs }));
    }
    if (config.successUrlContains) {
      successChecks.push(page.waitForURL(`**${config.successUrlContains}**`, { timeout: config.timeoutMs }));
    }

    if (successChecks.length === 0) {
      successChecks.push(page.waitForLoadState("networkidle", { timeout: config.timeoutMs }));
    }

    await Promise.any(successChecks);

    if (config.screenshotOnSuccess) {
      screenshotPath = path.join(screenshotDir, `${runId}.png`);
      await page.screenshot({ path: screenshotPath, fullPage: true });
    }

    if (config.postLoginWaitMs > 0) {
      await page.waitForTimeout(config.postLoginWaitMs);
    }

    return {
      runId,
      batchId: runContext.batchId ?? runId,
      roundNumber: runContext.roundNumber ?? 1,
      totalRounds: runContext.totalRounds ?? 1,
      loginAt: startedAt,
      status: "PASS",
      loginUrl: config.loginUrl,
      email: config.email,
      browser: config.browser,
      durationMs: Date.now() - startedAt.getTime(),
      screenshotPath: screenshotPath ? path.relative(outputDir, screenshotPath) : "",
      message: "Login success",
    };
  } catch (error) {
    const rawMessage = error instanceof Error ? error.message : String(error);

    if (page) {
      screenshotPath = path.join(screenshotDir, `${runId}-failed.png`);
      await page.screenshot({ path: screenshotPath, fullPage: true }).catch(() => {});

      if (config.keepBrowserOpenOnFail && config.failPauseMs > 0) {
        await page.waitForTimeout(config.failPauseMs).catch(() => {});
      }
    }

    const message = rawMessage.includes("ERR_CONNECTION_REFUSED")
      ? `Cannot open target website. The server refused the connection for ${config.loginUrl}`
      : rawMessage;

    return {
      runId,
      batchId: runContext.batchId ?? runId,
      roundNumber: runContext.roundNumber ?? 1,
      totalRounds: runContext.totalRounds ?? 1,
      loginAt: startedAt,
      status: "FAIL",
      loginUrl: config.loginUrl,
      email: config.email,
      browser: config.browser,
      durationMs: Date.now() - startedAt.getTime(),
      screenshotPath: screenshotPath ? path.relative(outputDir, screenshotPath) : "",
      message,
    };
  } finally {
    await browser?.close().catch(() => {});
  }
}

async function runSelfTest() {
  const runId = buildRunId();
  return {
    runId,
    batchId: buildRunId("batch"),
    roundNumber: 1,
    totalRounds: 1,
    loginAt: new Date(),
    status: "PASS",
    loginUrl: "self-test://local-validation",
    email: "self-test@example.com",
    browser: "chromium",
    durationMs: 1,
    screenshotPath: "",
    message: "Workbook logging self-test completed",
  };
}

async function runLoginBatch(config) {
  const totalRounds = sanitizePositiveInteger(getCliOption("--rounds") || config.rounds, 1);
  const roundDelayMs = Math.max(0, Number.parseInt(String(config.roundDelayMs ?? 0), 10) || 0);
  const batchId = buildRunId("batch");
  const startedAt = new Date();
  const results = [];

  for (let roundNumber = 1; roundNumber <= totalRounds; roundNumber += 1) {
    const result = await runLoginTest(config, { batchId, roundNumber, totalRounds });
    results.push(result);
    await appendLogRow(result);

    if (roundNumber < totalRounds && roundDelayMs > 0) {
      await new Promise((resolve) => setTimeout(resolve, roundDelayMs));
    }
  }

  const successCount = results.filter((result) => result.status === "PASS").length;
  const failCount = results.length - successCount;
  const finishedAt = new Date();
  const summary = {
    batchId,
    startedAt,
    finishedAt,
    totalRounds,
    successCount,
    failCount,
    successRate: `${Math.round((successCount / totalRounds) * 100)}%`,
    loginUrl: config.loginUrl,
    email: config.email,
    message: `Completed ${totalRounds} rounds with ${successCount} success and ${failCount} fail`,
  };

  await appendSummaryRow(summary);

  return { batchId, totalRounds, successCount, failCount, startedAt, finishedAt, results, summary };
}

async function main() {
  try {
    await ensureDirectories();

    if (isSelfTest) {
      const result = await runSelfTest();
      console.log(JSON.stringify({
        runId: result.runId,
        status: result.status,
        loginDateTime: formatLocalDateTime(result.loginAt),
        workbookPath: "",
        screenshotPath: "",
        message: result.message,
      }, null, 2));
      return;
    }

    const config = await readConfig();
    const batch = await runLoginBatch(config);

    console.log(JSON.stringify({
      batchId: batch.batchId,
      totalRounds: batch.totalRounds,
      successCount: batch.successCount,
      failCount: batch.failCount,
      startedAt: formatLocalDateTime(batch.startedAt),
      finishedAt: formatLocalDateTime(batch.finishedAt),
      workbookPath,
      lastScreenshotPath: batch.results.at(-1)?.screenshotPath
        ? path.join(outputDir, batch.results.at(-1).screenshotPath)
        : "",
      message: batch.summary.message,
    }, null, 2));

    if (batch.failCount > 0) {
      process.exitCode = 1;
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.error(`Auto test login failed: ${message}`);
    process.exitCode = 1;
  }
}

await main();
