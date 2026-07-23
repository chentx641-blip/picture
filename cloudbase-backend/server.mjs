import express from "express";
import ExcelJS from "exceljs";
import { createWorker } from "tesseract.js";

const app = express();
const maxRows = 100, maxColumns = 16, maxImages = 320;
let worker;
const feishuSessions = new Map();
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
function feishuConfigured() { return Boolean(process.env.FEISHU_APP_ID && process.env.FEISHU_APP_SECRET && process.env.FEISHU_REDIRECT_URI); }
function cleanFeishuSessions() { const now = Date.now(); for (const [key, value] of feishuSessions) if (value.expiresAt < now) feishuSessions.delete(key); }
function makePlan(columns, columnRoles) {
  const plan = [];
  for (let c = 0; c < columns; c += 1) { const role = columnRoles[c] || "image"; plan.push({ source:c, kind:"image", title:`Image ${c + 1}` }); if (role === "uid_did") plan.push({source:c,kind:"uid",title:"UID"},{source:c,kind:"did",title:"DID"}); if (role === "xhs") plan.push({source:c,kind:"xhs",title:"\u5c0f\u7ea2\u4e66\u53f7"}); if (role === "phone") plan.push({source:c,kind:"phone",title:"\u624b\u673a\u578b\u53f7"}); }
  return plan;
}
async function buildEmbeddedWorkbook(matrix, columnRoles = [], results = {}) {
  if (!Array.isArray(matrix) || !matrix.every(Array.isArray)) throw new Error("Invalid data");
  const rows = matrix.slice(0, maxRows), columns = Math.min(maxColumns, Math.max(0, ...rows.map((row) => row.length)));
  if (!rows.length || !columns || rows.length * columns > maxImages) throw new Error("Export limit exceeded");
  const plan = makePlan(columns, columnRoles);
  const workbook = new ExcelJS.Workbook(), sheet = workbook.addWorksheet("\u56fe\u7247"); sheet.views=[{showGridLines:false}]; sheet.addRow(plan.map((x)=>x.title)); sheet.getRow(1).font={bold:true,color:{argb:"FFFFFFFF"}}; sheet.getRow(1).fill={type:"pattern",pattern:"solid",fgColor:{argb:"FF217A52"}}; sheet.getRow(1).height=24;
  plan.forEach((x,c)=>sheet.getColumn(c+1).width=x.kind==="image"?25:22);
  for (let r=0;r<rows.length;r+=1) { const row=sheet.getRow(r+2); row.height=100; for (let c=0;c<plan.length;c+=1) { const item=plan[c], value=results[`${r}-${item.source}`]||{}; if(item.kind!=="image") { const cell=row.getCell(c+1); cell.value=String(value[item.kind]||""); cell.numFmt="@"; } else if(isImageSource(rows[r][item.source])) { try { const src=await image(rows[r][item.source]), id=workbook.addImage({buffer:src.data,extension:src.type.includes("png")?"png":"jpeg"}); sheet.addImage(id,{tl:{col:c+.1,row:r+1.1},ext:{width:150,height:90}}); } catch { row.getCell(c+1).value="Image download failed"; } } } }
  return Buffer.from(await workbook.xlsx.writeBuffer());
}

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
app.get("/feishu/status", (_, res) => res.json({ configured: feishuConfigured() }));
app.get("/feishu/authorize", (req, res) => {
  cleanFeishuSessions();
  if (!feishuConfigured()) return res.status(503).send("Feishu integration is not configured");
  const session = String(req.query.session || "");
  if (!/^[A-Za-z0-9_-]{24,120}$/.test(session)) return res.status(400).send("Invalid session");
  const state = crypto.randomUUID().replaceAll("-", "");
  feishuSessions.set(state, { session, expiresAt: Date.now() + 10 * 60 * 1000 });
  const params = new URLSearchParams({ app_id: process.env.FEISHU_APP_ID, redirect_uri: process.env.FEISHU_REDIRECT_URI, state });
  res.redirect(`https://open.feishu.cn/open-apis/authen/v1/authorize?${params}`);
});
app.get("/feishu/callback", async (req, res) => {
  const state = String(req.query.state || ""), code = String(req.query.code || ""), pending = feishuSessions.get(state);
  const origin = process.env.FRONTEND_ORIGIN || "https://chentx641-blip.github.io";
  const page = (message) => `<!doctype html><meta charset=\"utf-8\"><script>if(window.opener)window.opener.postMessage(${JSON.stringify({ type:"feishu-auth", message, session:pending?.session || "" })},${JSON.stringify(origin)});window.close()</script><p>${message}</p>`;
  if (!pending || pending.expiresAt < Date.now() || !code) return res.status(400).send(page("Feishu authorization expired or was cancelled. You can close this page."));
  try {
    const tokenResponse = await fetch("https://open.feishu.cn/open-apis/authen/v1/access_token", { method:"POST", headers:{"Content-Type":"application/json; charset=utf-8"}, body:JSON.stringify({ grant_type:"authorization_code", code, client_id:process.env.FEISHU_APP_ID, client_secret:process.env.FEISHU_APP_SECRET, redirect_uri:process.env.FEISHU_REDIRECT_URI }) });
    const tokenJson = await tokenResponse.json();
    if (!tokenResponse.ok || tokenJson.code) throw new Error(tokenJson.msg || "Token exchange failed");
    feishuSessions.delete(state);
    feishuSessions.set(pending.session, { userAccessToken: tokenJson.data?.access_token, expiresAt: Date.now() + Math.max(60, Number(tokenJson.data?.expires_in || 7200) - 60) * 1000 });
    res.send(page("Feishu authorization completed. You can close this page and return to Picture Workspace."));
  } catch (error) { res.status(400).send(page(`Feishu authorization failed: ${error.message}`)); }
});
app.post("/upload-feishu-xlsx", async (req, res) => {
  try {
    cleanFeishuSessions();
    if (!feishuConfigured()) return res.status(503).send("Feishu integration is not configured");
    const session = feishuSessions.get(String(req.body?.session || ""));
    if (!session?.userAccessToken) return res.status(401).send("Please connect Feishu first");
    const content = await buildEmbeddedWorkbook(req.body.matrix, req.body.columnRoles, req.body.results);
    const root = await fetch("https://open.feishu.cn/open-apis/drive/explorer/v2/root_folder/meta", { headers:{ Authorization:`Bearer ${session.userAccessToken}` } });
    const rootJson = await root.json(); if (!root.ok || rootJson.code) throw new Error(rootJson.msg || "Unable to access Feishu Drive root folder");
    const folderToken = rootJson.data?.meta?.token;
    if (!folderToken) throw new Error("Feishu Drive root folder token was not returned");
    const filename = `\u56fe\u94fe\u5de5\u574a-${new Date().toISOString().slice(0,10)}.xlsx`;
    const form = new FormData(); form.set("file_name", filename); form.set("parent_type", "explorer"); form.set("parent_node", folderToken); form.set("size", String(content.length)); form.set("file", new Blob([content], { type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), filename);
    const upload = await fetch("https://open.feishu.cn/open-apis/drive/v1/files/upload_all", { method:"POST", headers:{ Authorization:`Bearer ${session.userAccessToken}` }, body:form });
    const uploadJson = await upload.json(); if (!upload.ok || uploadJson.code) throw new Error(uploadJson.msg || "Feishu upload failed");
    const fileToken = uploadJson.data?.file_token;
    if (!fileToken) throw new Error("Feishu upload completed but file token was not returned");
    res.json({ name: filename, fileToken, url:`https://feishu.cn/file/${fileToken}` });
  } catch (error) { res.status(400).send(error.message || "Feishu upload failed"); }
});
app.get("/health", (_, res) => res.json({ ok:true }));
app.listen(process.env.PORT || 8080, () => console.log("Picture Workspace API ready"));
