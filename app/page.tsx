"use client";

import { ChangeEvent, useMemo, useState } from "react";

type Result = {
  url: string;
  status: "ready" | "error";
  uid: string;
  did: string;
  xhs: string;
  phone: string;
};

type ColumnRole = "image" | "uid_did" | "xhs" | "phone";
const roleLabels: Record<ColumnRole, string> = { image: "普通图片", uid_did: "UID + DID 图片", xhs: "小红书号图片", phone: "手机型号图片" };

const example =
  "https://images.unsplash.com/photo-1497366811353-6870744d04b2?auto=format&fit=crop&w=900&q=80\thttps://images.unsplash.com/photo-1516321318423-f06f85e504b3?auto=format&fit=crop&w=900&q=80\nhttps://images.unsplash.com/photo-1551836022-d5d88e9218df?auto=format&fit=crop&w=900&q=80\thttps://images.unsplash.com/photo-1556761175-b413da4baf72?auto=format&fit=crop&w=900&q=80";

function parseMatrix(text: string) {
  return text
    .trim()
    .split(/\r?\n/)
    .filter(Boolean)
    .map((row) => row.split("\t").map((value) => value.trim()));
}

function normaliseUrl(value: string) {
  try {
    const url = new URL(value);
    return /^https?:$/.test(url.protocol) ? url.toString() : "";
  } catch {
    return "";
  }
}

export default function Home() {
  const [source, setSource] = useState("");
  const [matrix, setMatrix] = useState<string[][]>([]);
  const [results, setResults] = useState<Record<string, Result>>({});
  const [notice, setNotice] = useState("");
  const [exporting, setExporting] = useState(false);
  const [recognising, setRecognising] = useState(false);
  const [columnRoles, setColumnRoles] = useState<ColumnRole[]>([]);

  const columnCount = useMemo(
    () => Math.max(0, ...matrix.map((row) => row.length)),
    [matrix],
  );

  const convert = () => {
    const parsed = parseMatrix(source);
    const next: Record<string, Result> = {};
    parsed.forEach((row, rowIndex) =>
      row.forEach((value, index) => {
        const url = normaliseUrl(value);
        const key = `${rowIndex}-${index}`;
        next[key] = { url, status: url ? "ready" : "error", uid: "", did: "", xhs: "", phone: "" };
      }),
    );
    setMatrix(parsed);
    setColumnRoles(Array.from({ length: Math.max(...parsed.map((row) => row.length)) }, () => "image"));
    setResults(next);
    setNotice(parsed.length ? `已生成 ${parsed.length} 行、${Math.max(...parsed.map((r) => r.length))} 列图片卡片。` : "请先粘贴链接。");
  };

  const recognize = async () => {
    if (!matrix.length || !columnRoles.some((role) => role !== "image")) {
      setNotice("请先在下方为至少一列选择需要识别的字段类型。");
      return;
    }
    setRecognising(true);
    setNotice("正在识别已选字段图片，请稍候…");
    try {
      const response = await fetch("http://localhost:3101/recognize", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ matrix, columnRoles }) });
      if (!response.ok) throw new Error(await response.text());
      const data = await response.json();
      setResults((current) => {
        const next = { ...current };
        Object.entries(data).forEach(([key, value]) => {
          const fields = value as { uid?: string; did?: string; xhs?: string; phone?: string };
          next[key] = { ...next[key], uid: fields.uid || next[key].uid, did: fields.did || next[key].did, xhs: fields.xhs || next[key].xhs, phone: fields.phone || next[key].phone };
        });
        return next;
      });
      setNotice("识别完成：请检查图片下方的结果，必要时直接修改。");
    } catch {
      setNotice("识别服务暂不可用，请确认本地工具正在运行后重试。");
    } finally { setRecognising(false); }
  };

  const updateResult = (key: string, field: "uid" | "did" | "xhs" | "phone", value: string) => {
    setResults((current) => ({
      ...current,
      [key]: { ...current[key], [field]: value },
    }));
  };

  const copyText = async () => {
    const rows = matrix.map((row, rowIndex) =>
      row
        .flatMap((url, columnIndex) => {
          const result = results[`${rowIndex}-${columnIndex}`];
          const role = columnRoles[columnIndex] || "image";
          return role === "uid_did" ? [url, result?.uid ?? "", result?.did ?? ""] : role === "xhs" ? [url, result?.xhs ?? ""] : role === "phone" ? [url, result?.phone ?? ""] : [url];
        })
        .join("\t"),
    );
    await navigator.clipboard.writeText(rows.join("\n"));
    setNotice("已复制链接、UID、DID 文本，可直接粘贴到飞书电子表格。");
  };

  const loadExample = () => {
    setSource(example);
    setNotice("已填入演示链接，点击“生成图片”查看多列效果。");
  };

  const exportExcel = async () => {
    if (!matrix.length) {
      setNotice("请先生成图片，再下载 Excel。");
      return;
    }
    setExporting(true);
    setNotice("正在下载图片并生成 Excel，请稍候…");
    try {
      const response = await fetch("http://localhost:3101/export-xlsx", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ matrix, columnRoles, results }),
      });
      if (!response.ok) throw new Error(await response.text());
      const blob = await response.blob();
      const anchor = document.createElement("a");
      anchor.href = URL.createObjectURL(blob);
      anchor.download = "链接转图片_嵌入图片.xlsx";
      anchor.click();
      URL.revokeObjectURL(anchor.href);
      setNotice("Excel 已下载：图片已经嵌入单元格，可直接上传到飞书。");
    } catch {
      setNotice("Excel 导出服务未启动。请保持本地预览服务运行后重试。");
    } finally {
      setExporting(false);
    }
  };

  return (
    <main>
      <section className="hero">
        <div className="eyebrow">LINK TO IMAGE WORKSPACE</div>
        <h1><span className="headline-nowrap">把链接，整齐地变成图片。</span></h1>
        <p>按原有 Excel 行列粘贴图片链接；生成预览、补充识别字段，再把文本结果一键带回飞书表格。</p>
      </section>

      <section className="workspace" aria-label="链接转图片工具">
        <div className="panel input-panel">
          <div className="panel-heading">
            <div><span className="step">01</span><h2>粘贴链接矩阵</h2></div>
            <button className="text-button" onClick={loadExample}>载入示例</button>
          </div>
          <p className="hint">从 Excel 直接复制多行、多列链接。换行代表新行，Tab 代表新列。</p>
          <textarea
            value={source}
            onChange={(event: ChangeEvent<HTMLTextAreaElement>) => setSource(event.target.value)}
            placeholder={"https://example.com/a.jpg\thttps://example.com/b.jpg\nhttps://example.com/c.jpg\thttps://example.com/d.jpg"}
            aria-label="图片链接"
          />
          <div className="actions"><button className="primary" onClick={convert}>生成图片</button><span>{notice}</span></div>
        </div>

        <div className="panel guide-panel">
          <span className="step">02</span><h2>识别字段</h2>
          <p>当前页面已完成链接转图片与表格保序。UID、DID 自动识别将在下一阶段接入 OCR 服务；现在可先在每张图下手动校对填写。</p>
          <span className="privacy">图片仅从原链接加载，不在此演示版保存。</span>
        </div>
      </section>

      {matrix.length > 0 && (
        <section className="results">
          <div className="results-heading">
            <div><div className="eyebrow">RESULT MATRIX</div><h2>{matrix.length} 行 × {columnCount} 列</h2></div>
            <div className="result-actions"><button className="secondary" onClick={recognize} disabled={recognising}>{recognising ? "正在识别…" : "识别已选字段"}</button><button className="secondary" onClick={copyText}>复制结果文本</button><button className="primary" onClick={exportExcel} disabled={exporting}>{exporting ? "正在生成 Excel…" : "下载嵌入图片的 Excel"}</button></div>
          </div>
          <div className="column-settings" role="group" aria-label="列识别字段设置">{columnRoles.map((role, index) => <label key={index}>图片列 {index + 1}<select value={role} onChange={(event) => setColumnRoles((current) => current.map((item, itemIndex) => itemIndex === index ? event.target.value as ColumnRole : item))}>{Object.entries(roleLabels).map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></label>)}</div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>行</th>{Array.from({ length: columnCount }, (_, index) => <th key={index}>图片列 {index + 1}<small>{roleLabels[columnRoles[index] || "image"]}</small></th>)}</tr></thead>
              <tbody>{matrix.map((row, rowIndex) => <tr key={rowIndex}><th className="row-number">{rowIndex + 1}</th>{Array.from({ length: columnCount }, (_, columnIndex) => {
                const key = `${rowIndex}-${columnIndex}`;
                const result = results[key];
                const role = columnRoles[columnIndex] || "image";
                return <td key={key}>{result?.url ? <article className="image-card"><img src={result.url} alt={`第 ${rowIndex + 1} 行第 ${columnIndex + 1} 列图片`} onError={() => setResults((current) => ({ ...current, [key]: { ...current[key], status: "error" } }))} />{result.status === "error" && <div className="image-error">图片无法加载：请检查链接是否过期或禁止跨域访问。</div>}{role !== "image" && <div className="fields">{role === "uid_did" ? <><label>UID<input value={result.uid} onChange={(event) => updateResult(key, "uid", event.target.value)} placeholder="待 OCR 识别" /></label><label>DID<input value={result.did} onChange={(event) => updateResult(key, "did", event.target.value)} placeholder="待 OCR 识别" /></label></> : <label className="wide-field">{role === "xhs" ? "小红书号" : "手机型号"}<input value={role === "xhs" ? result.xhs : result.phone} onChange={(event) => updateResult(key, role, event.target.value)} placeholder="待 OCR 识别" /></label>}</div>}</article> : <div className="invalid">无效链接</div>}</td>;
              })}</tr>)}</tbody>
            </table>
          </div>
        </section>
      )}
    </main>
  );
}
