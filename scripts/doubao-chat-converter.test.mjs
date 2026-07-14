import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import vm from "node:vm";

const source = await fs.readFile(new URL("../assets/doubao-chat-converter.js", import.meta.url), "utf8");
assert.doesNotMatch(source, /\bfetch\s*\(|XMLHttpRequest|sendBeacon|WebSocket|\/api\//);

function fakeElement() {
  return {
    value: "",
    textContent: "",
    innerHTML: "",
    hidden: false,
    disabled: false,
    checked: false,
    files: [],
    className: "",
    classList: { add() {}, remove() {} },
    addEventListener() {},
    append() {},
    click() {}
  };
}

const elements = new Map();
const document = {
  querySelector(selector) {
    if (selector === 'input[name="outputFormat"]:checked') return { value: "txt" };
    if (!elements.has(selector)) elements.set(selector, fakeElement());
    return elements.get(selector);
  },
  createElement() {
    return fakeElement();
  }
};
const window = {
  setTimeout,
  addEventListener() {}
};
const context = vm.createContext({
  Blob,
  DataView,
  Date,
  DOMParser: class DOMParser {
    parseFromString(value) {
      return { body: { textContent: String(value).replace(/<[^>]+>/g, "") } };
    }
  },
  JSON,
  Map,
  Math,
  Number,
  Object,
  Promise,
  RegExp,
  Set,
  String,
  TextDecoderStream: undefined,
  TextEncoder,
  Uint8Array,
  Uint32Array,
  URL: { createObjectURL() { return "blob:test"; }, revokeObjectURL() {} },
  console,
  document,
  setTimeout,
  window
});
vm.runInContext(source, context, { filename: "doubao-chat-converter.js" });
const api = window.DoubaoChatConverter;
assert.ok(api);

const manifest = {
  type: "manifest",
  format: "2link-doubao-text-backup",
  version: 1,
  title: "顾泽川"
};
const rows = [
  { type: "message", role: "assistant", content: "第一条回复", index: 0 },
  { type: "message", role: "user", content: "第二条提问", index: 1, assets: [{ type: "image", url: "https://example.com/a.jpg", alt: "图片" }] }
];
const ndjson = [manifest, ...rows].map((row) => JSON.stringify(row)).join("\n");
const parsedNdjson = await api.parseChatFile({ name: "messages.ndjson", text: async () => ndjson });
assert.equal(parsedNdjson.title, "顾泽川");
assert.equal(parsedNdjson.messages.length, 2);
assert.deepEqual(Array.from(parsedNdjson.messages, (message) => String(message.role)), ["assistant", "user"]);

const parsedJson = api.normalizeJsonExport({ payload: { title: "测试角色", messages: rows } }, "chat.json");
assert.equal(parsedJson.title, "测试角色");
assert.equal(parsedJson.messages.length, 2);

const largeRows = Array.from({ length: 10_000 }, (_, index) => JSON.stringify({
  type: "message",
  role: index % 2 ? "user" : "assistant",
  content: `压力消息 ${index}`,
  index
}));
const parsedLarge = await api.parseChatFile({ name: "large.ndjson", text: async () => largeRows.join("\n") });
assert.equal(parsedLarge.messages.length, 10_000);
assert.equal(parsedLarge.messages[0].content, "压力消息 0");
assert.equal(parsedLarge.messages.at(-1).content, "压力消息 9999");

const options = { includeIndex: true, includeTime: false, includeAssets: true };
const text = api.buildText(parsedJson.title, parsedJson.messages, options);
assert.match(text, /#1 · 豆包/);
assert.match(text, /#2 · 用户/);
assert.match(text, /https:\/\/example\.com\/a\.jpg/);

const markdown = api.buildMarkdown(parsedJson.title, parsedJson.messages, options);
assert.match(markdown, /^# 测试角色/m);
assert.match(markdown, /\[图片\]\(https:\/\/example\.com\/a\.jpg\)/);

const docx = api.buildDocx(parsedJson.title, parsedJson.messages, options);
const docxBytes = Buffer.from(await docx.arrayBuffer());
assert.deepEqual([...docxBytes.subarray(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
for (const entry of ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/_rels/document.xml.rels", "word/styles.xml"]) {
  assert.ok(docxBytes.includes(Buffer.from(entry)), `missing DOCX entry: ${entry}`);
}
const outputPath = path.join(os.tmpdir(), "2link-doubao-converter-test.docx");
await fs.writeFile(outputPath, docxBytes);
console.log("doubao chat converter ok", {
  messages: parsedNdjson.messages.length,
  stressMessages: parsedLarge.messages.length,
  docxBytes: docxBytes.length,
  outputPath
});

if (process.env.CONVERTER_REAL_FILE) {
  const realPath = path.resolve(process.env.CONVERTER_REAL_FILE);
  const realText = await fs.readFile(realPath, "utf8");
  const realParsed = await api.parseChatFile({
    name: path.basename(realPath),
    text: async () => realText
  });
  assert.ok(realParsed.messages.length > 0, "real export contains no messages");
  const realDocx = api.buildDocx(realParsed.title, realParsed.messages, options);
  const realOutputPath = path.join(os.tmpdir(), "2link-doubao-converter-real-export.docx");
  await fs.writeFile(realOutputPath, Buffer.from(await realDocx.arrayBuffer()));
  console.log("real export conversion ok", {
    source: realPath,
    messages: realParsed.messages.length,
    first: realParsed.messages[0].content.slice(0, 60),
    last: realParsed.messages.at(-1).content.slice(0, 60),
    outputPath: realOutputPath
  });
}
