const elements = {
  dropZone: document.querySelector("#dropZone"),
  chooseFileButton: document.querySelector("#chooseFileButton"),
  replaceFileButton: document.querySelector("#replaceFileButton"),
  clearFileButton: document.querySelector("#clearFileButton"),
  fileInput: document.querySelector("#fileInput"),
  fileSummary: document.querySelector("#fileSummary"),
  fileName: document.querySelector("#fileName"),
  fileMeta: document.querySelector("#fileMeta"),
  messageCount: document.querySelector("#messageCount"),
  previewList: document.querySelector("#previewList"),
  includeIndex: document.querySelector("#includeIndex"),
  includeTime: document.querySelector("#includeTime"),
  includeAssets: document.querySelector("#includeAssets"),
  documentTitle: document.querySelector("#documentTitle"),
  status: document.querySelector("#status"),
  downloadButton: document.querySelector("#downloadButton")
};

const state = {
  file: null,
  title: "豆包聊天记录",
  messages: [],
  sourceFormat: ""
};

bindEvents();

function bindEvents() {
  elements.chooseFileButton.addEventListener("click", openFilePicker);
  elements.replaceFileButton.addEventListener("click", openFilePicker);
  elements.clearFileButton.addEventListener("click", clearFile);
  elements.fileInput.addEventListener("change", () => loadFile(elements.fileInput.files?.[0]));
  elements.downloadButton.addEventListener("click", downloadConvertedDocument);
  elements.dropZone.addEventListener("dragover", (event) => {
    event.preventDefault();
    elements.dropZone.classList.add("dragging");
  });
  elements.dropZone.addEventListener("dragleave", () => elements.dropZone.classList.remove("dragging"));
  elements.dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    elements.dropZone.classList.remove("dragging");
    loadFile(event.dataTransfer?.files?.[0]);
  });
}

function openFilePicker() {
  elements.fileInput.value = "";
  elements.documentTitle.value = "豆包聊天记录";
  elements.fileInput.click();
}

async function loadFile(file) {
  if (!file) return;
  setStatus(`正在读取 ${file.name}...`);
  elements.downloadButton.disabled = true;
  try {
    const parsed = await parseChatFile(file, (count) => {
      setStatus(`正在解析，已读取 ${count.toLocaleString("zh-CN")} 条消息...`);
    });
    if (!parsed.messages.length) throw new Error("文件中没有找到聊天消息。请确认选择的是 messages.ndjson 或聊天 JSON 导出包。");
    state.file = file;
    state.title = parsed.title || titleFromFileName(file.name);
    state.messages = parsed.messages;
    state.sourceFormat = parsed.sourceFormat;
    renderLoadedFile();
    setStatus(`解析完成：${parsed.messages.length.toLocaleString("zh-CN")} 条消息。`, "success");
  } catch (error) {
    clearFile({ keepStatus: true });
    setStatus(`无法转换：${error.message}`, "error");
  }
}

function clearFile(options = {}) {
  state.file = null;
  state.title = "豆包聊天记录";
  state.messages = [];
  state.sourceFormat = "";
  elements.fileInput.value = "";
  elements.dropZone.hidden = false;
  elements.fileSummary.classList.remove("visible");
  elements.downloadButton.disabled = true;
  elements.messageCount.textContent = "尚未选择文件";
  elements.previewList.innerHTML = '<div class="preview-empty">选择文件后显示前 20 条消息</div>';
  if (!options.keepStatus) setStatus("");
}

function renderLoadedFile() {
  elements.dropZone.hidden = true;
  elements.fileSummary.classList.add("visible");
  elements.fileName.textContent = state.file.name;
  elements.fileMeta.textContent = `${formatBytes(state.file.size)} · ${state.sourceFormat} · ${state.title}`;
  elements.documentTitle.value = state.title;
  elements.messageCount.textContent = `共 ${state.messages.length.toLocaleString("zh-CN")} 条，预览前 20 条`;
  elements.previewList.innerHTML = "";
  for (const [index, message] of state.messages.slice(0, 20).entries()) {
    const row = document.createElement("div");
    row.className = "preview-message";
    const speaker = document.createElement("div");
    speaker.className = `speaker${message.role === "system" ? " system" : ""}`;
    speaker.textContent = speakerLabel(message.role);
    const text = document.createElement("p");
    text.className = "message-text";
    text.textContent = message.content || `[${assetSummary(message.assets) || "空消息"}]`;
    row.append(speaker, text);
    elements.previewList.appendChild(row);
  }
  elements.downloadButton.disabled = false;
}

async function downloadConvertedDocument() {
  if (!state.messages.length) return;
  const format = document.querySelector('input[name="outputFormat"]:checked')?.value || "txt";
  const options = {
    includeIndex: elements.includeIndex.checked,
    includeTime: elements.includeTime.checked,
    includeAssets: elements.includeAssets.checked
  };
  elements.downloadButton.disabled = true;
  setStatus("正在生成文档...");
  try {
    await yieldToBrowser();
    const documentTitle = elements.documentTitle.value.trim() || state.title;
    const safeTitle = safeFileName(documentTitle);
    if (format === "docx") {
      downloadBlob(buildDocx(documentTitle, state.messages, options), `${safeTitle}.docx`);
    } else if (format === "md") {
      downloadBlob(withUtf8Bom(buildMarkdown(documentTitle, state.messages, options)), `${safeTitle}.md`, "text/markdown;charset=utf-8");
    } else {
      downloadBlob(withUtf8Bom(buildText(documentTitle, state.messages, options)), `${safeTitle}.txt`, "text/plain;charset=utf-8");
    }
    setStatus(`已生成 ${safeTitle}.${format}。`, "success");
  } catch (error) {
    setStatus(`生成失败：${error.message}`, "error");
  } finally {
    elements.downloadButton.disabled = false;
  }
}

async function parseChatFile(file, onProgress = () => {}) {
  const name = String(file?.name || "").toLowerCase();
  if (name.endsWith(".ndjson") || name.endsWith(".jsonl")) {
    return parseNdjsonFile(file, onProgress);
  }
  const text = await file.text();
  try {
    return normalizeJsonExport(JSON.parse(text), file.name);
  } catch (jsonError) {
    if (text.includes("\n")) return parseNdjsonText(text, file.name, onProgress);
    throw new Error(`JSON 格式不完整：${jsonError.message}`);
  }
}

async function parseNdjsonFile(file, onProgress) {
  const messages = [];
  let manifest = null;
  let lineNumber = 0;
  for await (const line of iterateFileLines(file)) {
    lineNumber += 1;
    const trimmed = line.trim();
    if (!trimmed) continue;
    let row;
    try {
      row = JSON.parse(trimmed);
    } catch (error) {
      throw new Error(`第 ${lineNumber} 行不是有效 JSON：${error.message}`);
    }
    if (isManifestRow(row) && !manifest) {
      manifest = row;
      continue;
    }
    const message = normalizeMessage(row, messages.length);
    if (message) messages.push(message);
    if (messages.length && messages.length % 500 === 0) {
      onProgress(messages.length);
      await yieldToBrowser();
    }
  }
  return {
    title: manifest?.title || titleFromFileName(file.name),
    messages,
    sourceFormat: "NDJSON"
  };
}

function parseNdjsonText(text, fileName, onProgress) {
  const messages = [];
  let manifest = null;
  for (const [index, rawLine] of text.split(/\r?\n/).entries()) {
    const line = rawLine.trim();
    if (!line) continue;
    let row;
    try {
      row = JSON.parse(line);
    } catch (error) {
      throw new Error(`第 ${index + 1} 行不是有效 JSON：${error.message}`);
    }
    if (isManifestRow(row) && !manifest) manifest = row;
    else {
      const message = normalizeMessage(row, messages.length);
      if (message) messages.push(message);
    }
    if (messages.length && messages.length % 500 === 0) onProgress(messages.length);
  }
  return { title: manifest?.title || titleFromFileName(fileName), messages, sourceFormat: "NDJSON" };
}

function normalizeJsonExport(value, fileName = "chat.json") {
  let payload = value;
  if (payload?.payload && typeof payload.payload === "object") payload = payload.payload;
  else if (payload?.data?.payload && typeof payload.data.payload === "object") payload = payload.data.payload;
  const sourceMessages = Array.isArray(payload)
    ? payload
    : Array.isArray(payload?.messages)
      ? payload.messages
      : Array.isArray(payload?.data?.messages)
        ? payload.data.messages
        : [];
  const messages = sourceMessages.map(normalizeMessage).filter(Boolean);
  const title = payload?.title
    || value?.importAgent?.title
    || value?.import?.title
    || value?.title
    || titleFromFileName(fileName);
  return { title: String(title || "豆包聊天记录"), messages, sourceFormat: "JSON" };
}

function normalizeMessage(row = {}, index = 0) {
  if (!row || typeof row !== "object" || isManifestRow(row)) return null;
  const content = normalizeContent(row.content ?? row.text ?? row.message ?? row.html ?? "");
  const assets = normalizeAssets(row.assets || row.metadata?.importedAssets || row.attachments || []);
  if (!content && !assets.length) return null;
  return {
    index: Number.isFinite(Number(row.index)) ? Number(row.index) : index,
    role: normalizeRole(row.role || row.author || row.sender || row.type),
    content,
    createdAt: normalizeDate(row.createdAt || row.timestamp || row.time || row.created_at),
    assets
  };
}

function normalizeRole(value) {
  const role = String(value || "").toLowerCase();
  if (["user", "human", "me", "用户", "我"].includes(role)) return "user";
  if (["system", "系统"].includes(role)) return "system";
  return "assistant";
}

function normalizeContent(value) {
  if (typeof value === "object" && value !== null) {
    if (typeof value.text === "string") value = value.text;
    else value = JSON.stringify(value);
  }
  const text = String(value || "");
  if (!/<[a-z][\s\S]*>/i.test(text)) return text.trim();
  const documentNode = new DOMParser().parseFromString(text, "text/html");
  return String(documentNode.body?.textContent || "").replace(/\u00a0/g, " ").trim();
}

function normalizeAssets(value) {
  return (Array.isArray(value) ? value : []).map((asset) => {
    if (typeof asset === "string") return { type: "附件", url: asset, name: "" };
    return {
      type: String(asset?.type || "附件"),
      url: String(asset?.url || asset?.src || asset?.cloudAssetKey || ""),
      name: String(asset?.alt || asset?.name || asset?.localPath || "")
    };
  }).filter((asset) => asset.url || asset.name);
}

function isManifestRow(row) {
  return row?.type === "manifest" || (row?.format && !row?.role && !row?.content && !row?.text);
}

async function* iterateFileLines(file) {
  if (file?.stream && typeof TextDecoderStream !== "undefined") {
    const reader = file.stream().pipeThrough(new TextDecoderStream()).getReader();
    let pending = "";
    try {
      while (true) {
        const { value, done } = await reader.read();
        if (done) break;
        pending += value;
        let newline = pending.indexOf("\n");
        while (newline >= 0) {
          yield pending.slice(0, newline).replace(/\r$/, "");
          pending = pending.slice(newline + 1);
          newline = pending.indexOf("\n");
        }
      }
      if (pending) yield pending;
    } finally {
      reader.releaseLock();
    }
    return;
  }
  for (const line of String(await file.text()).split(/\r?\n/)) yield line;
}

function buildText(title, messages, options = {}) {
  const lines = [title, `消息数量：${messages.length}`, ""];
  for (const [index, message] of messages.entries()) {
    lines.push(messageHeading(message, index, options));
    if (message.content) lines.push(message.content);
    if (options.includeAssets) lines.push(...assetLines(message.assets));
    lines.push("");
  }
  return lines.join("\r\n");
}

function buildMarkdown(title, messages, options = {}) {
  const lines = [`# ${escapeMarkdown(title)}`, "", `消息数量：${messages.length}`, ""];
  for (const [index, message] of messages.entries()) {
    lines.push(`## ${escapeMarkdown(messageHeading(message, index, options))}`, "");
    if (message.content) lines.push(message.content, "");
    if (options.includeAssets) {
      for (const asset of message.assets) {
        const label = asset.name || asset.type || "附件";
        lines.push(asset.url ? `- [${escapeMarkdown(label)}](${asset.url})` : `- ${escapeMarkdown(label)}`);
      }
      if (message.assets.length) lines.push("");
    }
  }
  return lines.join("\n");
}

function buildDocx(title, messages, options = {}) {
  const body = [];
  body.push(docxParagraph(title, "Title"));
  body.push(docxParagraph(`消息数量：${messages.length}`, "Subtitle"));
  for (const [index, message] of messages.entries()) {
    body.push(docxParagraph(messageHeading(message, index, options), "Heading2"));
    if (message.content) body.push(docxParagraph(message.content, "Normal"));
    if (options.includeAssets) {
      for (const line of assetLines(message.assets)) body.push(docxParagraph(line, "Asset"));
    }
  }
  body.push('<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>');
  const documentXml = xmlHeader()
    + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'
    + body.join("")
    + "</w:body></w:document>";
  const files = [
    ["[Content_Types].xml", contentTypesXml()],
    ["_rels/.rels", rootRelationshipsXml()],
    ["word/document.xml", documentXml],
    ["word/_rels/document.xml.rels", documentRelationshipsXml()],
    ["word/styles.xml", stylesXml()]
  ];
  return createStoredZip(files, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
}

function messageHeading(message, index, options) {
  const parts = [];
  if (options.includeIndex) parts.push(`#${index + 1}`);
  parts.push(speakerLabel(message.role));
  if (options.includeTime && message.createdAt) parts.push(formatDateTime(message.createdAt));
  return parts.join(" · ");
}

function speakerLabel(role) {
  if (role === "user") return "用户";
  if (role === "system") return "系统";
  return "豆包";
}

function assetLines(assets = []) {
  return assets.map((asset) => {
    const label = asset.name || asset.type || "附件";
    return asset.url ? `附件：${label} ${asset.url}` : `附件：${label}`;
  });
}

function assetSummary(assets = []) {
  return assets.map((asset) => asset.name || asset.type).filter(Boolean).join("、");
}

function docxParagraph(text, style) {
  const runs = String(text || "").split(/\r?\n/).map((line, index) => {
    const breakTag = index ? "<w:br/>" : "";
    return `<w:r>${breakTag}<w:t xml:space="preserve">${escapeXml(line)}</w:t></w:r>`;
  }).join("");
  return `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr>${runs || "<w:r><w:t></w:t></w:r>"}</w:p>`;
}

function contentTypesXml() {
  return xmlHeader() + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
    + '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
    + "</Types>";
}

function rootRelationshipsXml() {
  return xmlHeader() + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
    + "</Relationships>";
}

function documentRelationshipsXml() {
  return xmlHeader() + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    + "</Relationships>";
}

function stylesXml() {
  return xmlHeader() + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
    + '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Microsoft YaHei" w:eastAsia="Microsoft YaHei"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>'
    + '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:spacing w:after="180" w:line="360" w:lineRule="auto"/></w:pPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Title"><w:name w:val="Title"/><w:pPr><w:spacing w:after="240"/></w:pPr><w:rPr><w:b/><w:sz w:val="36"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Subtitle"><w:name w:val="Subtitle"/><w:pPr><w:spacing w:after="360"/></w:pPr><w:rPr><w:color w:val="647174"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="Heading 2"/><w:pPr><w:spacing w:before="260" w:after="100"/></w:pPr><w:rPr><w:b/><w:color w:val="087F80"/><w:sz w:val="24"/></w:rPr></w:style>'
    + '<w:style w:type="paragraph" w:styleId="Asset"><w:name w:val="Asset"/><w:pPr><w:spacing w:after="80"/></w:pPr><w:rPr><w:color w:val="647174"/><w:sz w:val="18"/></w:rPr></w:style>'
    + "</w:styles>";
}

function xmlHeader() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
}

function createStoredZip(files, mimeType) {
  const encoder = new TextEncoder();
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  const { time, date } = dosDateTime(new Date());
  for (const [name, content] of files) {
    const nameBytes = encoder.encode(name);
    const data = typeof content === "string" ? encoder.encode(content) : content;
    const checksum = crc32(data);
    const local = new Uint8Array(30 + nameBytes.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(4, 20, true);
    localView.setUint16(6, 0x0800, true);
    localView.setUint16(8, 0, true);
    localView.setUint16(10, time, true);
    localView.setUint16(12, date, true);
    localView.setUint32(14, checksum, true);
    localView.setUint32(18, data.length, true);
    localView.setUint32(22, data.length, true);
    localView.setUint16(26, nameBytes.length, true);
    local.set(nameBytes, 30);
    localParts.push(local, data);

    const central = new Uint8Array(46 + nameBytes.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint16(4, 20, true);
    centralView.setUint16(6, 20, true);
    centralView.setUint16(8, 0x0800, true);
    centralView.setUint16(10, 0, true);
    centralView.setUint16(12, time, true);
    centralView.setUint16(14, date, true);
    centralView.setUint32(16, checksum, true);
    centralView.setUint32(20, data.length, true);
    centralView.setUint32(24, data.length, true);
    centralView.setUint16(28, nameBytes.length, true);
    centralView.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length + data.length;
  }
  const centralSize = centralParts.reduce((sum, part) => sum + part.length, 0);
  const end = new Uint8Array(22);
  const endView = new DataView(end.buffer);
  endView.setUint32(0, 0x06054b50, true);
  endView.setUint16(8, files.length, true);
  endView.setUint16(10, files.length, true);
  endView.setUint32(12, centralSize, true);
  endView.setUint32(16, offset, true);
  return new Blob([...localParts, ...centralParts, end], { type: mimeType });
}

const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let value = 0; value < 256; value += 1) {
    let crc = value;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    table[value] = crc >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = (crc >>> 8) ^ crcTable[(crc ^ byte) & 0xff];
  return (crc ^ 0xffffffff) >>> 0;
}

function dosDateTime(value) {
  const year = Math.max(1980, value.getFullYear());
  return {
    time: (value.getHours() << 11) | (value.getMinutes() << 5) | Math.floor(value.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((value.getMonth() + 1) << 5) | value.getDate()
  };
}

function downloadBlob(value, fileName, type = "application/octet-stream") {
  const blob = value instanceof Blob ? value : new Blob([value], { type });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 30_000);
}

function withUtf8Bom(text) {
  return `\ufeff${text}`;
}

function setStatus(text, mode = "") {
  elements.status.textContent = text;
  elements.status.className = `status${mode ? ` ${mode}` : ""}`;
}

function normalizeDate(value) {
  if (!value && value !== 0) return "";
  const date = typeof value === "number" ? new Date(value > 10_000_000_000 ? value : value * 1000) : new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toISOString();
}

function formatDateTime(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toLocaleString("zh-CN", { hour12: false });
}

function formatBytes(value) {
  const bytes = Number(value || 0);
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 ** 2) return `${(bytes / 1024).toFixed(1)} KB`;
  if (bytes < 1024 ** 3) return `${(bytes / 1024 ** 2).toFixed(1)} MB`;
  return `${(bytes / 1024 ** 3).toFixed(1)} GB`;
}

function titleFromFileName(name) {
  return String(name || "豆包聊天记录")
    .replace(/\.2link-doubao(?:-text)?\.ndjson$/i, "")
    .replace(/\.(?:json|jsonl|ndjson)$/i, "")
    .replace(/^messages$/i, "豆包聊天记录")
    || "豆包聊天记录";
}

function safeFileName(value) {
  return String(value || "豆包聊天记录").replace(/[\\/:*?"<>|\u0000-\u001f]+/g, "_").trim().slice(0, 80) || "豆包聊天记录";
}

function escapeXml(value) {
  return String(value || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function escapeMarkdown(value) {
  return String(value || "").replace(/([\\`*_{}\[\]()#+.!|-])/g, "\\$1");
}

function yieldToBrowser() {
  return new Promise((resolve) => window.setTimeout(resolve, 0));
}

window.DoubaoChatConverter = Object.freeze({
  parseChatFile,
  normalizeJsonExport,
  buildText,
  buildMarkdown,
  buildDocx
});
