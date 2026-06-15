#!/usr/bin/env node

/**
 * 教师详情补采脚本骨架
 *
 * 输入：八爪鱼导出的 CSV
 * 输出：符合模版字段的 CSV
 *
 * 设计目标：
 * 1. 零第三方依赖，拿到就能在本机跑
 * 2. 先跑通通用网站，再逐步加站点规则
 * 3. 保留日志，方便人工复查
 */

const fs = require("node:fs");
const path = require("node:path");

const TEMPLATE_HEADERS = [
  "序号",
  "姓名",
  "职称",
  "邮箱",
  "学校",
  "学院",
  "采集人",
  "采集日期",
  "简介",
  "主页",
];

const DEFAULT_HEADERS = {
  name: ["name", "姓名", "教师姓名", "字段1", "姓名名称"],
  profileUrl: ["profile_url", "主页", "教师主页", "详情页", "链接", "url", "链接地址", "网址"],
  school: ["school", "学校", "院校"],
  college: ["college", "学院", "院系", "学部"],
  title: ["title", "职称"],
  email: ["email", "邮箱", "电子邮箱"],
  intro: ["intro", "简介", "个人简介"],
};

const KNOWN_TITLES = [
  "助理教授",
  "副教授",
  "教授",
  "助理研究员",
  "副研究员",
  "研究员",
  "高级工程师",
  "助理工程师",
  "工程师",
  "实验师",
  "讲师",
  "教师",
];
const PRIORITY_TITLES = [
  "副教授",
  "教授",
  "讲师",
];
const TITLE_ALIASES = {
  "助理教授": ["助理教授", "assistant professor"],
  "副教授": ["副教授", "associate professor"],
  "教授": ["教授", "professor"],
  "助理研究员": ["助理研究员", "assistant researcher"],
  "副研究员": ["副研究员", "associate researcher"],
  "研究员": ["研究员", "researcher"],
  "高级工程师": ["高级工程师", "senior engineer"],
  "助理工程师": ["助理工程师", "assistant engineer"],
  "工程师": ["工程师", "engineer"],
  "实验师": ["实验师"],
  "讲师": ["讲师", "lecturer"],
  "教师": ["教师", "teacher"],
};
const TITLE_LABEL_CAPTURE_PATTERNS = [
  /职\s*称\s*[:：]\s*([^\n]{0,30})/,
  /基本情况[^\n]{0,120}?职\s*称\s*[:：]\s*([^\n]{0,30})/,
];
const PROFESSOR_LEVEL_PATTERN = /([一二三四1234])级?教授/;
const ASSOCIATE_PROFESSOR_LEVEL_PATTERN = /([五六七567])级?(?:副教授|教授)/;

async function main() {
  const args = parseArgs(process.argv.slice(2));
  if (args.help || !args.input) {
    printHelp();
    process.exit(args.help ? 0 : 1);
  }

  const inputPath = path.resolve(args.input);
  const outputPath = path.resolve(args.output || createDefaultOutputPath(inputPath));
  const logPath = path.resolve(args.log || createDefaultLogPath(outputPath));

  ensureFileExists(inputPath);

  const rows = parseCsv(fs.readFileSync(inputPath, "utf8"));
  if (rows.length === 0) {
    throw new Error("输入文件为空，无法处理。");
  }

  const [headerRow, ...dataRows] = rows;
  const headerMap = buildHeaderMap(headerRow);
  const collector = args.collector || "未填写";
  const collectedAt = args.date || today();
  const concurrency = clampInt(args.concurrency || "5", 1, 20);
  const limit = args.limit ? clampInt(args.limit, 1, Number.MAX_SAFE_INTEGER) : dataRows.length;

  const selectedRows = dataRows.slice(0, limit).filter((row) => row.some(Boolean));
  const records = selectedRows.map((row, index) => normalizeSeedRow(row, headerMap, index));

  const logEntries = [];
  const results = await mapWithConcurrency(records, concurrency, async (record, index) => {
    const detail = await enrichRecord(record, { timeoutMs: 15000, logEntries });
    return toTemplateRow(detail, {
      index: index + 1,
      collector,
      collectedAt,
    });
  });

  const csvText = stringifyCsv([TEMPLATE_HEADERS, ...results]);
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, csvText, "utf8");

  const logPayload = {
    createdAt: new Date().toISOString(),
    inputPath,
    outputPath,
    totalInputRows: dataRows.length,
    processedRows: selectedRows.length,
    collector,
    collectedAt,
    logs: logEntries,
  };
  fs.mkdirSync(path.dirname(logPath), { recursive: true });
  fs.writeFileSync(logPath, JSON.stringify(logPayload, null, 2), "utf8");

  console.log(`处理完成：${selectedRows.length} 条`);
  console.log(`输出文件：${outputPath}`);
  console.log(`日志文件：${logPath}`);
}

function parseArgs(args) {
  const parsed = {};
  for (let i = 0; i < args.length; i += 1) {
    const arg = args[i];
    if (arg === "--help" || arg === "-h") {
      parsed.help = true;
      continue;
    }
    if (!arg.startsWith("--")) {
      continue;
    }
    const [key, inlineValue] = arg.slice(2).split("=", 2);
    if (inlineValue !== undefined) {
      parsed[key] = inlineValue;
      continue;
    }
    const next = args[i + 1];
    if (!next || next.startsWith("--")) {
      parsed[key] = "true";
      continue;
    }
    parsed[key] = next;
    i += 1;
  }
  return parsed;
}

function printHelp() {
  console.log(`
用法：
  node scripts/teacher_profile_pipeline.js --input <列表CSV> [--output <结果CSV>]

常用参数：
  --collector <采集人>
  --date <采集日期，例 2026-06-01>
  --concurrency <并发数，默认 5>
  --limit <只处理前 N 条，便于调试>
  --log <日志 JSON 路径>

示例：
  node scripts/teacher_profile_pipeline.js \\
    --input "./原始数据/八爪鱼导出/天津商业大学_公共管理学院_列表.csv" \\
    --output "./输出结果/天津商业大学_公共管理学院_导入.csv" \\
    --collector "你的名字" \\
    --date "2026-06-01"
`.trim());
}

function ensureFileExists(filePath) {
  if (!fs.existsSync(filePath)) {
    throw new Error(`文件不存在：${filePath}`);
  }
}

function createDefaultOutputPath(inputPath) {
  const dir = path.resolve("输出结果");
  const base = path.basename(inputPath, path.extname(inputPath));
  return path.join(dir, `${buildOutputStem(base)}.csv`);
}

function createDefaultLogPath(outputPath) {
  const dir = path.resolve("日志");
  const base = path.basename(outputPath, path.extname(outputPath));
  return path.join(dir, `${base}_日志.json`);
}

function buildOutputStem(inputStem) {
  let stem = inputStem.trim();
  for (const suffix of ["_导入_Python", "_导入", "_导出", "导入_Python", "导入", "导出", "输出"]) {
    if (stem.endsWith(suffix)) {
      stem = stem.slice(0, -suffix.length).replace(/[_\\-\\s]+$/u, "");
      break;
    }
  }
  return `${stem}输出`;
}

function buildHeaderMap(headerRow) {
  const normalizedHeaders = headerRow.map((value) => normalizeHeader(value));
  const map = {};
  for (const [field, aliases] of Object.entries(DEFAULT_HEADERS)) {
    const aliasSet = new Set(aliases.map(normalizeHeader));
    map[field] = normalizedHeaders.findIndex((header) => aliasSet.has(header));
  }
  if (map.name < 0 && headerRow.length >= 1) {
    map.name = 0;
  }
  if (map.profileUrl < 0 && headerRow.length >= 2) {
    map.profileUrl = 1;
  }
  return map;
}

function normalizeSeedRow(row, headerMap, index) {
  return {
    rowNumber: index + 2,
    name: getField(row, headerMap.name),
    profileUrl: getField(row, headerMap.profileUrl),
    school: getField(row, headerMap.school),
    college: getField(row, headerMap.college),
    title: getField(row, headerMap.title),
    email: getField(row, headerMap.email),
    intro: getField(row, headerMap.intro),
  };
}

async function enrichRecord(record, { timeoutMs, logEntries }) {
  const result = { ...record };

  if (!record.profileUrl) {
    logEntries.push({
      level: "warn",
      rowNumber: record.rowNumber,
      name: record.name,
      message: "缺少主页链接，已跳过详情抓取",
    });
    return result;
  }

  try {
    const response = await fetchWithTimeout(record.profileUrl, timeoutMs);
    const html = await response.text();
    const cleanedHtml = normalizeWhitespace(html);
    const text = htmlToText(html);
    const context = buildTeacherContext(text, record.name);

    if (!result.email) {
      result.email = extractEmail(cleanedHtml, text, context);
    }
    if (!result.title) {
      result.title = extractTitle(text, context);
    }
    if (!result.intro) {
      result.intro = extractIntro(html, text);
    }

    if (!result.email) {
      logEntries.push({
        level: "info",
        rowNumber: record.rowNumber,
        name: record.name,
        url: record.profileUrl,
        message: "未提取到邮箱",
      });
    }
    if (!result.title) {
      logEntries.push({
        level: "info",
        rowNumber: record.rowNumber,
        name: record.name,
        url: record.profileUrl,
        message: "未提取到职称",
      });
    }
  } catch (error) {
    logEntries.push({
      level: "error",
      rowNumber: record.rowNumber,
      name: record.name,
      url: record.profileUrl,
      message: error.message,
    });
  }

  return result;
}

async function fetchWithTimeout(url, timeoutMs) {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, {
      signal: controller.signal,
      headers: {
        "user-agent": "Mozilla/5.0 Codex Teacher Pipeline/1.0",
      },
    });
    if (!response.ok) {
      throw new Error(`请求失败：HTTP ${response.status}`);
    }
    return response;
  } finally {
    clearTimeout(timer);
  }
}

function extractEmail(html, text, context = "") {
  const sources = [context, html, text].filter(Boolean);
  for (const source of sources) {
    const normalized = normalizeEmailLikeText(source);
    const emails = collectEmails(normalized);
    const preferred = pickBestEmail(emails);
    if (preferred) {
      return preferred;
    }
  }
  return "";
}

function extractTitle(text, context = "") {
  const sources = [context, text].filter(Boolean).map(normalizeWhitespace);
  for (const source of sources) {
    const labeledTitle = extractLabeledTitle(source);
    if (labeledTitle) {
      return labeledTitle;
    }
    const title = findTitleInText(source);
    if (title) {
      return title;
    }
  }

  return "";
}

function extractIntro(html, text) {
  const metaDescription = matchMetaDescription(html);
  if (metaDescription) {
    return truncateText(metaDescription, 200);
  }

  const lines = text
    .split("\n")
    .map((line) => normalizeWhitespace(line))
    .filter(Boolean)
    .filter((line) => line.length >= 20)
    .filter((line) => !looksLikeNavigation(line));

  return truncateText(lines[0] || "", 200);
}

function matchMetaDescription(html) {
  const match = html.match(/<meta[^>]+name=["']description["'][^>]+content=["']([^"']+)["']/i);
  return match ? normalizeWhitespace(decodeHtml(match[1])) : "";
}

function looksLikeNavigation(line) {
  const badTokens = [
    "首页",
    "当前位置",
    "联系我们",
    "版权所有",
    "学院概况",
    "人才培养",
    "科学研究",
    "招生就业",
    "师资队伍",
    "关闭",
    "打印",
  ];
  return badTokens.some((token) => line.includes(token));
}

function toTemplateRow(record, { index, collector, collectedAt }) {
  return [
    String(index),
    cleanName(record.name),
    normalizeTitle(record.title),
    record.email || "",
    record.school || "",
    record.college || "",
    collector,
    collectedAt,
    truncateText(normalizeWhitespace(record.intro || ""), 200),
    record.profileUrl || "",
  ];
}

function cleanName(name) {
  return normalizePersonName(
    normalizeWhitespace(name)
    .replace(/(教授|副教授|讲师|博士|硕导|博导|老师)$/g, "")
    .trim()
  );
}

function normalizeTitle(title) {
  const normalized = normalizeWhitespace(title);
  const aliasTitle = findTitleInText(normalized);
  if (aliasTitle) {
    return aliasTitle;
  }
  const priorityTitle = pickPriorityTitle(normalized);
  if (priorityTitle) {
    return priorityTitle;
  }
  return KNOWN_TITLES.find((item) => normalized.includes(item)) || normalized;
}

function getField(row, index) {
  if (index < 0 || index >= row.length) {
    return "";
  }
  return normalizeWhitespace(row[index] || "");
}

function normalizeHeader(value) {
  return normalizeWhitespace(String(value || "")).toLowerCase();
}

function normalizeWhitespace(value) {
  return String(value || "")
    .replace(/\u00a0/g, " ")
    .replace(/\u3000/g, " ")
    .replace(/\r/g, "\n")
    .replace(/\t/g, " ")
    .replace(/[ ]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function normalizePersonName(value) {
  return normalizeWhitespace(value).replace(/(?<=[\u4e00-\u9fa5])\s+(?=[\u4e00-\u9fa5])/g, "");
}

function buildTeacherContext(text, name) {
  const normalizedText = normalizeWhitespace(text);
  const cleanTeacherName = cleanName(name);
  if (!normalizedText || !cleanTeacherName) {
    return "";
  }

  const index = normalizedText.indexOf(cleanTeacherName);
  if (index < 0) {
    return "";
  }

  const start = Math.max(0, index - 120);
  const end = Math.min(normalizedText.length, index + cleanTeacherName.length + 240);
  return normalizedText.slice(start, end);
}

function normalizeEmailLikeText(value) {
  return String(value || "")
    .replace(/\[at\]|\(at\)|（at）/gi, "@")
    .replace(/\[dot\]|\(dot\)|（dot）/gi, ".")
    .replace(/&#64;|＠/g, "@")
    .replace(/。/g, ".")
    .replace(/\s*@\s*/g, "@")
    .replace(/\s*\.\s*/g, ".")
    .replace(/mailto:/gi, "");
}

function collectEmails(value) {
  const emailRegex = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi;
  return [...new Set((value.match(emailRegex) || []).map((item) => item.trim()))];
}

function pickBestEmail(candidates) {
  if (candidates.length === 0) {
    return "";
  }

  const filtered = candidates.filter((email) => {
    const normalized = email.toLowerCase();
    return ![
      "example.com",
      "test.com",
      "163.com.cn",
    ].some((badDomain) => normalized.endsWith(`@${badDomain}`));
  });

  return filtered[0] || candidates[0] || "";
}

function extractLabeledTitle(source) {
  for (const pattern of TITLE_LABEL_CAPTURE_PATTERNS) {
    const match = source.match(pattern);
    if (!match) {
      continue;
    }
    const value = normalizeWhitespace(match[1]);
    const normalized = mapTitleByLevel(value);
    if (normalized) {
      return normalized;
    }
    const aliasTitle = findTitleInText(value);
    if (aliasTitle) {
      return aliasTitle;
    }
    const priorityTitle = pickPriorityTitle(value);
    if (priorityTitle) {
      return priorityTitle;
    }
    const title = KNOWN_TITLES.find((item) => value.includes(item));
    if (title) {
      return title;
    }
  }
  return "";
}

function mapTitleByLevel(value) {
  if (PROFESSOR_LEVEL_PATTERN.test(value)) {
    return "教授";
  }
  if (ASSOCIATE_PROFESSOR_LEVEL_PATTERN.test(value)) {
    return "副教授";
  }
  return "";
}

function pickPriorityTitle(value) {
  return PRIORITY_TITLES.find((item) => titleMatchesAlias(value, item) || value.includes(item)) || "";
}

function findTitleInText(source) {
  const matches = findTitleMatches(source);
  return matches.length > 0 ? matches[0].canonical : "";
}

function findTitleMatches(source) {
  const matches = [];
  for (const [canonical, aliases] of Object.entries(TITLE_ALIASES)) {
    for (const alias of aliases) {
      for (const match of iterTitleAliasMatches(source, alias)) {
        matches.push({
          start: match.index,
          end: match.index + match[0].length,
          canonical,
        });
      }
    }
  }
  matches.sort((a, b) => (
    a.start - b.start
    || titlePriority(a.canonical) - titlePriority(b.canonical)
    || (b.end - b.start) - (a.end - a.start)
  ));

  const deduped = [];
  const seen = new Set();
  for (const item of matches) {
    const key = `${item.start}:${item.end}:${item.canonical}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    deduped.push(item);
  }
  return deduped;
}

function iterTitleAliasMatches(source, alias) {
  const regex = /[A-Za-z]/.test(alias)
    ? new RegExp(`\\b${escapeRegex(alias)}\\b`, "gi")
    : new RegExp(escapeRegex(alias), "g");
  return source.matchAll(regex);
}

function titleMatchesAlias(value, canonical) {
  const aliases = TITLE_ALIASES[canonical] || [canonical];
  return aliases.some((alias) => iterTitleAliasMatches(value, alias).next().value);
}

function titlePriority(title) {
  const index = KNOWN_TITLES.indexOf(title);
  return index >= 0 ? index : KNOWN_TITLES.length;
}

function escapeRegex(value) {
  return String(value || "").replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function truncateText(value, maxLength) {
  const normalized = normalizeWhitespace(value);
  return normalized.length > maxLength ? `${normalized.slice(0, maxLength)}...` : normalized;
}

function htmlToText(html) {
  return decodeHtml(
    String(html || "")
      .replace(/<script[\s\S]*?<\/script>/gi, " ")
      .replace(/<style[\s\S]*?<\/style>/gi, " ")
      .replace(/<br\s*\/?>/gi, "\n")
      .replace(/<\/p>/gi, "\n")
      .replace(/<[^>]+>/g, " ")
  )
    .replace(/[ ]+/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function decodeHtml(text) {
  return String(text || "")
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");
}

async function mapWithConcurrency(items, concurrency, worker) {
  const results = new Array(items.length);
  let cursor = 0;

  async function runWorker() {
    while (true) {
      const current = cursor;
      cursor += 1;
      if (current >= items.length) {
        return;
      }
      results[current] = await worker(items[current], current);
    }
  }

  const workers = Array.from({ length: Math.min(concurrency, items.length) }, () => runWorker());
  await Promise.all(workers);
  return results;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        cell += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (char === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((char === "\n" || char === "\r") && !inQuotes) {
      if (char === "\r" && next === "\n") {
        i += 1;
      }
      row.push(cell);
      rows.push(row);
      row = [];
      cell = "";
      continue;
    }

    cell += char;
  }

  if (cell.length > 0 || row.length > 0) {
    row.push(cell);
    rows.push(row);
  }

  return rows;
}

function stringifyCsv(rows) {
  return rows
    .map((row) =>
      row
        .map((cell) => {
          const value = String(cell ?? "");
          if (/[",\n]/.test(value)) {
            return `"${value.replace(/"/g, '""')}"`;
          }
          return value;
        })
        .join(",")
    )
    .join("\n");
}

function today() {
  const now = new Date();
  const year = now.getFullYear();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const day = String(now.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function clampInt(value, min, max) {
  const parsed = Number.parseInt(String(value), 10);
  if (Number.isNaN(parsed)) {
    return min;
  }
  return Math.min(Math.max(parsed, min), max);
}

main().catch((error) => {
  console.error(`处理失败：${error.message}`);
  process.exit(1);
});
