import express from "express";
import ExcelJS from "exceljs";
import { createWorker } from "tesseract.js";

const app = express();
const maxRows = 100;
const maxColumns = 16;
const maxImages = 320;
let worker;

app.use(express.json({ limit: "2mb" }));
app.use((_, response, next) => {
  response.setHeader("Access-Control-Allow-Origin", "*");
  response.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  response.setHeader("Access-Control-Allow-Headers", "Content-Type");
  next();
});
app.options("*", (_, response) => response.sendStatus(204));

function isImageSource(value) {
  if (typeof value !== "string") return false;
  if (/^data:image\/(png|jpe?g|webp|gif);base64,/i.test(value)) return true;
  try { const url = new URL(value); return url.protocol === "http:" || url.protocol === "https:"; } catch { return false; }
}
function clean(text) { return text.replace(/：/g, ":").replace(/[\r\n]+/g, "\n").replace(/[ \t]+/g, " "); }
function find(text, expression) { return clean(text).match(expression)?.[1]?.trim() || ""; }
function fields(text, role) {
  if (role === "uid_did") return {
    uid: find(text, /(?:User\s*[IiLl1][dD]|UID|用户\s*ID)\s*[:：]?\s*([A-Za-z0-9_-]{5,})/i),
    did: find(text, /(?:Device\s*[IiLl1][dD]|DID|设备\s*ID)\s*[:：]?\s*([A-Za-z0-9_-]{5,})/i),
  };
  if (role === "xhs") return { xhs: find(text, /(?:小\s*红\s*书\s*(?:号|ID)|RED\s*ID|XHS)\s*[:：]?\s*([A-Za-z0-9_.-]{3,})/i) };
  if (role === "phone") {
    const labelled = find(text, /(?:型号\s*名称|手机\s*型号|设备\s*型号|机型|Model)\s*[:：]?\s*([A-Za-z][A-Za-z0-9 .+_/-]{2,60})/i);
    const branded = find(text, /\b((?:iPhone|vivo|OPPO|HUAWEI|HONOR|Xiaomi|Redmi|Samsung)[ -]?[A-Za-z0-9][A-Za-z0-9 .+_/-]{0,40})/i);
    return { phone: labelled || branded };
  }
  return {};
}
async function image(url) {
  const inline = /^data:(image\/(?:png|jpe?g|webp|gif));base64,([A-Za-z0-9+/=\r\n]+)$/i.exec(url);
  if (inline) {
    const data = Buffer.from(inline[2], "base64");
    if (data.length > 15 * 1024 * 1024) throw new Error("图片超过 15 MB");
    return { data, type: inline[1].toLowerCase() };
  }
  const response = await fetch(url, { signal: AbortSignal.timeout(20000), redirect: "follow" });
  if (!response.ok) throw new Error(`图片下载失败 (${response.status})`);
  const type = response.headers.get("content-type")?.split(";")[0] || "image/jpeg";
  if (!type.startsWith("image/")) throw new Error("链接不是图片");
  const data = Buffer.from(await response.arrayBuffer());
  if (data.length > 15 * 1024 * 1024) throw new Error("图片超过 15 MB");
  return { data, type };
}
async function ocr() { if (!worker) worker = await createWorker("chi_sim+eng", 1, { logger: () => undefined }); return worker; }

app.post("/recognize", async (request, response) => {
  try {
    const { matrix, columnRoles = [] } = request.body;
    if (!Array.isArray(matrix)) throw new Error("数据格式无效");
    const result = {};
    const engine = await ocr();
    for (let r = 0; r < Math.min(matrix.length, maxRows); r += 1) for (let c = 0; c < Math.min(matrix[r].length, maxColumns); c += 1) {
      const role = columnRoles[c] || "image";
      if (role === "image" || !isImageSource(matrix[r][c])) continue;
      try { const source = await image(matrix[r][c]); const read = await engine.recognize(source.data); result[`${r}-${c}`] = fields(read.data.text, role); } catch { result[`${r}-${c}`] = {}; }
    }
    response.json(result);
  } catch (error) { response.status(400).send(error.message || "识别失败"); }
});

app.post("/export-xlsx", async (request, response) => {
  try {
    const { matrix, columnRoles = [], results = {} } = request.body;
    if (!Array.isArray(matrix) || !matrix.every(Array.isArray)) throw new Error("数据格式无效");
    const rows = matrix.slice(0, maxRows);
    const columns = Math.min(maxColumns, Math.max(0, ...rows.map((row) => row.length)));
    if (!rows.length || !columns || rows.length * columns > maxImages) throw new Error("导出数据超出限制");
    const plan = [];
    for (let c = 0; c < columns; c += 1) {
      const role = columnRoles[c] || "image";
      plan.push({ source: c, kind: "image", title: `图片列 ${c + 1}` });
      if (role === "uid_did") plan.push({ source: c, kind: "uid", title: "UID" }, { source: c, kind: "did", title: "DID" });
      if (role === "xhs") plan.push({ source: c, kind: "xhs", title: "小红书号" });
      if (role === "phone") plan.push({ source: c, kind: "phone", title: "手机型号" });
    }
    const workbook = new ExcelJS.Workbook(); const sheet = workbook.addWorksheet("图片"); sheet.views = [{ showGridLines: false }];
    sheet.addRow(plan.map((item) => item.title)); sheet.getRow(1).font = { bold: true, color: { argb: "FFFFFFFF" } }; sheet.getRow(1).fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF217A52" } }; sheet.getRow(1).height = 24;
    plan.forEach((item, c) => { sheet.getColumn(c + 1).width = item.kind === "image" ? 25 : 22; });
    for (let r = 0; r < rows.length; r += 1) {
      const row = sheet.getRow(r + 2); row.height = 100;
      for (let c = 0; c < plan.length; c += 1) {
        const item = plan[c]; const value = results[`${r}-${item.source}`] || {};
        if (item.kind !== "image") row.getCell(c + 1).value = value[item.kind] || "";
        else if (isImageSource(rows[r][item.source])) {
          try { const source = await image(rows[r][item.source]); const imageId = workbook.addImage({ buffer: source.data, extension: source.type.includes("png") ? "png" : "jpeg" }); sheet.addImage(imageId, { tl: { col: c + 0.1, row: r + 1.1 }, ext: { width: 150, height: 90 } }); } catch { row.getCell(c + 1).value = "图片下载失败"; }
        }
      }
    }
    const content = await workbook.xlsx.writeBuffer();
    response.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); response.setHeader("Content-Disposition", "attachment; filename*=UTF-8''%E9%93%BE%E6%8E%A5%E8%BD%AC%E5%9B%BE%E7%89%87.xlsx"); response.send(Buffer.from(content));
  } catch (error) { response.status(400).send(error.message || "导出失败"); }
});

app.get("/health", (_, response) => response.json({ ok: true }));
app.listen(process.env.PORT || 8080, () => console.log("Picture Workspace API ready"));
