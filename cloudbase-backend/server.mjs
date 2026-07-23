import express from "express";
import ExcelJS from "exceljs";
import { createWorker } from "tesseract.js";

const app = express();
const maxRows = 100, maxColumns = 16, maxImages = 320;
let worker;
app.use(express.json({ limit: "20mb" }));
app.use((_, res, next) => { res.setHeader("Access-Control-Allow-Origin", "*"); res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS"); res.setHeader("Access-Control-Allow-Headers", "Content-Type"); next(); });
app.options("*", (_, res) => res.sendStatus(204));

function isImageSource(value) {
  if (typeof value !== "string") return false;
  if (/^data:image\/(png|jpe?g|webp|gif);base64,/i.test(value)) return true;
  try { const url = new URL(value); return url.protocol === "http:" || url.protocol === "https:"; } catch { return false; }
}
function clean(text) { return String(text).replace(/[\r\n]+/g, "\n").replace(/[ \t]+/g, " "); }
function find(text, expression) { return clean(text).match(expression)?.[1]?.trim() || ""; }
function phoneCandidates(text) {
  const patterns = [
    /\b(iPhone\s+(?:[1-9]|1\d|2\d)(?:\s+(?:Pro(?:\s+Max)?|Plus|mini|Air))?)\b/gi,
    /\b((?:Xiaomi|Redmi)\s*\d{1,3}(?:\s*(?:Pro|Ultra|Max|T|S|SE))?)\b/gi,
    /\b(nova\s*\d{1,2}(?:\s*(?:Pro|Ultra|SE|i))?)\b/gi,
    /\b((?:HUAWEI|HONOR|vivo|OPPO|Samsung)\s*[A-Za-z]*\d+[A-Za-z0-9 .+-]{0,18})\b/gi,
  ];
  const values = [];
  for (const expression of patterns) for (const match of text.matchAll(expression)) {
    const value = match[1].replace(/\s+/g, " ").trim().replace(/\s+(?:HarmonyOS|HyperOS|OriginOS|ColorOS).*$/i, "");
    if (!values.some((item) => item.value.toLowerCase() === value.toLowerCase())) values.push({ value, index: match.index ?? 0 });
  }
  return values;
}
function pickPhone(text) {
  const content = clean(text), candidates = phoneCandidates(content);
  if (!candidates.length) return "";
  const modelLabel = /(?:\u578b\s*[\u53f7\u865f]\s*\u540d\s*[\u79f0\u7a31]|\u673a\s*\u578b\s*\u540d\s*[\u79f0\u7a31]|\u6a5f\s*\u578b\s*\u540d\s*[\u79f0\u7a31]|\u624b\s*[\u673a\u6a5f]\s*\u578b\s*[\u53f7\u865f]|Model\s*Name)/i.exec(content);
  if (modelLabel) { const after = (modelLabel.index ?? 0) + modelLabel[0].length; const hit = candidates.find((x) => x.index >= after && x.index - after < 150); if (hit) return hit.value; }
  const deviceLabel = /(?:\u8bbe\s*\u5907\s*\u540d\s*\u79f0|\u8a2d\s*\u5099\s*\u540d\s*\u7a31|\u8bbe\s*\u5907\s*\u578b\s*[\u53f7\u865f]|\u8a2d\s*\u5099\s*\u578b\s*[\u53f7\u865f]|Device\s*Name)/i.exec(content);
  if (deviceLabel) { const after = (deviceLabel.index ?? 0) + deviceLabel[0].length; const hit = candidates.find((x) => x.index >= after && x.index - after < 150); if (hit) return hit.value; }
  return candidates[0].value;
}
function pickXhs(text) {
  const content = clean(text);
  const label = /(?:\u5c0f\s*[\u7ea2\u7d05]\s*[\u4e66\u66f8]\s*[\u53f7\u865f]|RED\s*ID|XHS)\s*[:\uff1a]?\s*([A-Za-z0-9_.-]{3,})/i;
  const direct = content.match(label)?.[1]?.trim(); if (direct) return direct;
  const labelledLine = content.split("\n").find((line) => /\u5c0f\s*[\u7ea2\u7d05]\s*[\u4e66\u66f8]/i.test(line));
  const fromLine = labelledLine?.match(/\b(\d{6,15})\b/)?.[1]; if (fromLine) return fromLine;
  // Some stylized Traditional-Chinese screenshots lose label glyphs in OCR; accept one long ID near the profile header only.
  const candidates = [...content.slice(0, 700).matchAll(/\b(\d{7,12})\b/g)].map((item) => item[1]);
  return candidates.length === 1 ? candidates[0] : "";
}
function fields(text, role) {
  if (role === "uid_did") return {
    uid: find(text, /(?:User\s*[IiLl1][dD]|UID|\u7528\u6237\s*ID)\s*[:\uff1a]?\s*([A-Za-z0-9_-]{5,})/i),
    did: find(text, /(?:Device\s*[IiLl1][dD]|DID|\u8bbe\u5907\s*ID)\s*[:\uff1a]?\s*([A-Za-z0-9_-]{5,})/i),
  };
  if (role === "xhs") return { xhs: pickXhs(text) };
  if (role === "phone") return { phone: pickPhone(text) };
  return {};
}
async function image(source) {
  const inline = /^data:(image\/(?:png|jpe?g|webp|gif));base64,([A-Za-z0-9+/=\r\n]+)$/i.exec(source);
  if (inline) { const data = Buffer.from(inline[2], "base64"); if (data.length > 15 * 1024 * 1024) throw new Error("Image exceeds 15 MB"); return { data, type: inline[1].toLowerCase() }; }
  const response = await fetch(source, { signal: AbortSignal.timeout(20000), redirect: "follow" });
  if (!response.ok) throw new Error(`Image download failed (${response.status})`);
  const type = response.headers.get("content-type")?.split(";")[0] || "image/jpeg";
  if (!type.startsWith("image/")) throw new Error("Source is not an image");
  const data = Buffer.from(await response.arrayBuffer()); if (data.length > 15 * 1024 * 1024) throw new Error("Image exceeds 15 MB"); return { data, type };
}
async function ocr() { if (!worker) worker = await createWorker("chi_sim+chi_tra+eng", 1, { logger: () => undefined }); return worker; }

app.post("/recognize", async (req, res) => {
  try {
    const { matrix, columnRoles = [] } = req.body; if (!Array.isArray(matrix)) throw new Error("Invalid data");
    const result = {}, engine = await ocr();
    for (let r = 0; r < Math.min(matrix.length, maxRows); r += 1) for (let c = 0; c < Math.min(matrix[r].length, maxColumns); c += 1) {
      const role = columnRoles[c] || "image"; if (role === "image" || !isImageSource(matrix[r][c])) continue;
      try { const source = await image(matrix[r][c]); const read = await engine.recognize(source.data); result[`${r}-${c}`] = fields(read.data.text, role); } catch { result[`${r}-${c}`] = {}; }
    }
    res.json(result);
  } catch (error) { res.status(400).send(error.message || "Recognition failed"); }
});

app.post("/export-xlsx", async (req, res) => {
  try {
    const { matrix, columnRoles = [], results = {} } = req.body; if (!Array.isArray(matrix) || !matrix.every(Array.isArray)) throw new Error("Invalid data");
    const rows = matrix.slice(0, maxRows), columns = Math.min(maxColumns, Math.max(0, ...rows.map((row) => row.length)));
    if (!rows.length || !columns || rows.length * columns > maxImages) throw new Error("Export limit exceeded");
    const plan = [];
    for (let c = 0; c < columns; c += 1) { const role = columnRoles[c] || "image"; plan.push({ source:c, kind:"image", title:`Image ${c + 1}` }); if (role === "uid_did") plan.push({source:c,kind:"uid",title:"UID"},{source:c,kind:"did",title:"DID"}); if (role === "xhs") plan.push({source:c,kind:"xhs",title:"\u5c0f\u7ea2\u4e66\u53f7"}); if (role === "phone") plan.push({source:c,kind:"phone",title:"\u624b\u673a\u578b\u53f7"}); }
    const workbook = new ExcelJS.Workbook(), sheet = workbook.addWorksheet("\u56fe\u7247"); sheet.views=[{showGridLines:false}]; sheet.addRow(plan.map((x)=>x.title)); sheet.getRow(1).font={bold:true,color:{argb:"FFFFFFFF"}}; sheet.getRow(1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF217A52"}}; sheet.getRow(1).height=24;
    plan.forEach((x,c)=>sheet.getColumn(c+1).width=x.kind==="image"?25:22);
    for (let r=0;r<rows.length;r+=1) { const row=sheet.getRow(r+2); row.height=100; for (let c=0;c<plan.length;c+=1) { const item=plan[c], value=results[`${r}-${item.source}`]||{}; if(item.kind!=="image") { const cell=row.getCell(c+1); cell.value=String(value[item.kind]||""); cell.numFmt="@"; } else if(isImageSource(rows[r][item.source])) { try { const src=await image(rows[r][item.source]), id=workbook.addImage({buffer:src.data,extension:src.type.includes("png")?"png":"jpeg"}); sheet.addImage(id,{tl:{col:c+.1,row:r+1.1},ext:{width:150,height:90}}); } catch { row.getCell(c+1).value="Image download failed"; } } } }
    const content=await workbook.xlsx.writeBuffer(); res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); res.setHeader("Content-Disposition","attachment; filename=link-images.xlsx"); res.send(Buffer.from(content));
  } catch (error) { res.status(400).send(error.message || "Export failed"); }
});
app.post("/export-fields-xlsx", async (req, res) => {
  try {
    const { matrix, columnRoles = [], results = {} } = req.body;
    if (!Array.isArray(matrix) || !matrix.every(Array.isArray)) throw new Error("Invalid data");
    const workbook = new ExcelJS.Workbook(), sheet = workbook.addWorksheet("Recognition results");
    const headers = ["UID", "DID", "\u5c0f\u7ea2\u4e66\u53f7", "\u624b\u673a\u578b\u53f7"];
    sheet.addRow(headers); sheet.getRow(1).font={bold:true,color:{argb:"FFFFFFFF"}}; sheet.getRow(1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF217A52"}};
    headers.forEach((_, index) => { sheet.getColumn(index + 1).width = 24; });
    matrix.slice(0,maxRows).forEach((row, rowIndex) => {
      const values={uid:"",did:"",xhs:"",phone:""};
      row.forEach((_, columnIndex) => { const role=columnRoles[columnIndex]||"image", item=results[`${rowIndex}-${columnIndex}`]||{}; if(role==="uid_did"){values.uid=item.uid||"";values.did=item.did||""} if(role==="xhs")values.xhs=item.xhs||""; if(role==="phone")values.phone=item.phone||""; });
      const excelRow=sheet.addRow([values.uid,values.did,values.xhs,values.phone]);
      excelRow.eachCell((cell) => { cell.value=String(cell.value||""); cell.numFmt="@"; });
    });
    const content=await workbook.xlsx.writeBuffer(); res.setHeader("Content-Type","application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"); res.setHeader("Content-Disposition","attachment; filename=recognition-results.xlsx"); res.send(Buffer.from(content));
  } catch (error) { res.status(400).send(error.message || "Export failed"); }
});
app.get("/health", (_, res) => res.json({ ok:true }));
app.listen(process.env.PORT || 8080, () => console.log("Picture Workspace API ready"));
