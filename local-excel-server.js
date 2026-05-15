const http = require("http");
const fs = require("fs");
const path = require("path");
const { URL } = require("url");

const PORT = 8787;
const ROOT = __dirname;
const EXCEL_PATH = path.join(ROOT, "工作資料.xls");

const LOG_HEADERS = ["時間戳記", "日期", "工作類型", "工作時數", "今天完成了什麼", "遇到的問題", "解法", "明天要做什麼"];
const ISSUE_HEADERS = ["同步時間", "內部ID", "Issue編號", "標題", "類型", "狀態", "優先度", "Epic", "到期日", "描述", "驗收條件", "建立時間", "更新時間"];

function xmlEscape(value) {
  return String(value == null ? "" : value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function xmlUnescape(value) {
  return String(value || "")
    .replaceAll("&apos;", "'")
    .replaceAll("&quot;", '"')
    .replaceAll("&gt;", ">")
    .replaceAll("&lt;", "<")
    .replaceAll("&amp;", "&");
}

function normalizeWorkType(type) {
  return type === "AI Agent" || type === "AIAgent" || type === "Development" ? "Feature" : type;
}

function normalizeIssueStatus(status) {
  const validStatuses = ["Backlog", "To Do", "In Progress", "Done"];
  if (status === "Review") return "To Do";
  return validStatuses.includes(status) ? status : "Backlog";
}

function worksheetXml(name, headers, rows) {
  const rowXml = [headers].concat(rows).map((row) => {
    return "<Row>" + row.map((cell) => {
      return '<Cell><Data ss:Type="String">' + xmlEscape(cell) + "</Data></Cell>";
    }).join("") + "</Row>";
  }).join("");

  return '<Worksheet ss:Name="' + xmlEscape(name) + '"><Table>' + rowXml + "</Table></Worksheet>";
}

function writeWorkbook(logs, issues) {
  const logRows = logs.map((log) => [
    log.submitted_at || "",
    log.date || "",
    normalizeWorkType(log.type || ""),
    log.hours || "",
    log.work_done || "",
    log.problem || "",
    log.solution || "",
    log.tomorrow || ""
  ]);

  const issueRows = issues.map((issue) => [
    issue.updatedAt || issue.createdAt || "",
    issue.id || "",
    issue.key || "",
    issue.title || "",
    normalizeWorkType(issue.type || ""),
    normalizeIssueStatus(issue.status || "Backlog"),
    issue.priority || "",
    issue.epic || "",
    issue.dueDate || "",
    issue.description || "",
    issue.acceptance || "",
    issue.createdAt || "",
    issue.updatedAt || ""
  ]);

  const workbook = [
    '<?xml version="1.0" encoding="UTF-8"?>',
    '<?mso-application progid="Excel.Sheet"?>',
    '<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">',
    worksheetXml("總記錄", LOG_HEADERS, logRows),
    worksheetXml("工作排程", ISSUE_HEADERS, issueRows),
    "</Workbook>"
  ].join("");

  fs.writeFileSync(EXCEL_PATH, "\uFEFF" + workbook, "utf8");
}

function readSheetRows(xml, sheetName) {
  const sheetPattern = new RegExp('<Worksheet[^>]*(?:ss:Name|Name)="' + sheetName + '"[\\s\\S]*?<Table>([\\s\\S]*?)<\\/Table>[\\s\\S]*?<\\/Worksheet>');
  const sheetMatch = xml.match(sheetPattern);
  if (!sheetMatch) return [];

  return [...sheetMatch[1].matchAll(/<Row>([\s\S]*?)<\/Row>/g)]
    .slice(1)
    .map((rowMatch) => {
      return [...rowMatch[1].matchAll(/<Data[^>]*>([\s\S]*?)<\/Data>/g)]
        .map((cellMatch) => xmlUnescape(cellMatch[1]));
    });
}

function readWorkbook() {
  if (!fs.existsSync(EXCEL_PATH)) {
    writeWorkbook([], []);
  }

  const xml = fs.readFileSync(EXCEL_PATH, "utf8");
  const logs = readSheetRows(xml, "總記錄").map((row) => ({
    source: "work_log",
    submitted_at: row[0] || "",
    date: row[1] || "",
    type: normalizeWorkType(row[2] || ""),
    hours: row[3] || "",
    work_done: row[4] || "",
    problem: row[5] || "",
    solution: row[6] || "",
    tomorrow: row[7] || ""
  })).filter((log) => log.date || log.type || log.hours || log.work_done || log.problem || log.solution || log.tomorrow);

  const issues = readSheetRows(xml, "工作排程").map((row) => ({
    id: row[1] || "",
    key: row[2] || "",
    title: row[3] || "",
    type: normalizeWorkType(row[4] || "Other"),
    status: normalizeIssueStatus(row[5] || "Backlog"),
    priority: row[6] || "Medium",
    epic: row.length >= 14 ? row[8] || "" : row[7] || "",
    dueDate: row.length >= 14 ? row[9] || "" : row[8] || "",
    description: row.length >= 14 ? row[10] || "" : row[9] || "",
    acceptance: row.length >= 14 ? row[11] || "" : row[10] || "",
    createdAt: row.length >= 14 ? row[12] || "" : row[11] || "",
    updatedAt: row.length >= 14 ? row[13] || "" : row[12] || ""
  })).filter((issue) => issue.id || issue.key || issue.title);

  return { logs, issues };
}

function readBody(req) {
  return new Promise((resolve, reject) => {
    let body = "";
    req.on("data", (chunk) => body += chunk);
    req.on("end", () => {
      try {
        resolve(body ? JSON.parse(body) : {});
      } catch (error) {
        reject(error);
      }
    });
  });
}

function sendJson(res, data, status = 200) {
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Access-Control-Allow-Origin": "*"
  });
  res.end(JSON.stringify(data));
}

function sendFile(res, filePath) {
  const ext = path.extname(filePath).toLowerCase();
  const contentType = ext === ".html" ? "text/html; charset=utf-8" :
    ext === ".xls" ? "application/vnd.ms-excel; charset=utf-8" :
    "text/plain; charset=utf-8";

  res.writeHead(200, { "Content-Type": contentType });
  fs.createReadStream(filePath).pipe(res);
}

const server = http.createServer(async (req, res) => {
  try {
    if (req.method === "OPTIONS") {
      res.writeHead(204, {
        "Access-Control-Allow-Origin": "*",
        "Access-Control-Allow-Methods": "GET,POST,OPTIONS",
        "Access-Control-Allow-Headers": "Content-Type"
      });
      res.end();
      return;
    }

    const parsedUrl = new URL(req.url, "http://localhost:" + PORT);

    if (req.method === "GET" && parsedUrl.pathname === "/api/data") {
      sendJson(res, Object.assign({ excelPath: EXCEL_PATH }, readWorkbook()));
      return;
    }

    if (req.method === "POST" && parsedUrl.pathname === "/api/work-log") {
      const payload = await readBody(req);
      const data = readWorkbook();
      data.logs.push(payload);
      writeWorkbook(data.logs, data.issues);
      sendJson(res, { success: true, count: data.logs.length, excelPath: EXCEL_PATH });
      return;
    }

    if (req.method === "POST" && parsedUrl.pathname === "/api/work-logs") {
      const payload = await readBody(req);
      const data = readWorkbook();
      writeWorkbook(Array.isArray(payload.logs) ? payload.logs : [], data.issues);
      sendJson(res, { success: true, count: (payload.logs || []).length, excelPath: EXCEL_PATH });
      return;
    }

    if (req.method === "POST" && parsedUrl.pathname === "/api/issues") {
      const payload = await readBody(req);
      const data = readWorkbook();
      writeWorkbook(data.logs, Array.isArray(payload.issues) ? payload.issues : []);
      sendJson(res, { success: true, count: (payload.issues || []).length, excelPath: EXCEL_PATH });
      return;
    }

    if (req.method === "GET" && parsedUrl.pathname === "/api/excel") {
      readWorkbook();
      sendFile(res, EXCEL_PATH);
      return;
    }

    const requestPath = decodeURIComponent(parsedUrl.pathname === "/" ? "/工作紀錄.html" : parsedUrl.pathname);
    let filePath = path.normalize(path.join(ROOT, requestPath));

    if (!filePath.startsWith(ROOT) || !fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
      if (req.method === "GET" && path.extname(requestPath).toLowerCase() === ".html") {
        filePath = path.join(ROOT, "工作紀錄.html");
      } else {
        res.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
        res.end("Not found");
        return;
      }
    }

    sendFile(res, filePath);
  } catch (error) {
    sendJson(res, { success: false, error: error.message }, 500);
  }
});

server.listen(PORT, () => {
  console.log("Local Excel server running:");
  console.log("  http://localhost:" + PORT + "/工作排程.html");
  console.log("  http://localhost:" + PORT + "/工作紀錄.html");
  console.log("Excel file:");
  console.log("  " + EXCEL_PATH);
});
