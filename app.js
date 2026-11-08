import Anthropic from "https://cdn.jsdelivr.net/npm/@anthropic-ai/sdk@0.133.0/+esm";
import * as L from "./logic.js";
import { jpegToPdf } from "./pdf.js";
import { MS_TENANT_ID, MS_CLIENT_ID, SETTINGS_PATH, CLAUDE_MODEL } from "./config.js";

const $ = (id) => document.getElementById(id);
const SCOPES = ["Files.ReadWrite", "User.Read"];
const GRAPH = "https://graph.microsoft.com/v1.0";

// ---------- small storage helpers (private windows may throw) ----------
const store = {
  get(k, d) { try { const v = localStorage.getItem(k); return v == null ? d : JSON.parse(v); } catch { return d; } },
  set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch {} },
};

// ---------- Microsoft sign-in ----------
const msalApp = new msal.PublicClientApplication({
  auth: {
    clientId: MS_CLIENT_ID,
    authority: `https://login.microsoftonline.com/${MS_TENANT_ID}`,
    redirectUri: location.origin + location.pathname,
  },
  cache: { cacheLocation: "localStorage" },
});

function account() {
  return msalApp.getActiveAccount() || msalApp.getAllAccounts()[0] || null;
}

async function token() {
  const acc = account();
  if (!acc) { await msalApp.loginRedirect({ scopes: SCOPES }); return new Promise(() => {}); }
  try {
    return (await msalApp.acquireTokenSilent({ scopes: SCOPES, account: acc })).accessToken;
  } catch {
    await msalApp.acquireTokenRedirect({ scopes: SCOPES, account: acc });
    return new Promise(() => {});
  }
}

// ---------- OneDrive (Microsoft Graph) ----------
const enc = (segs) => segs.join("/").split("/").map(encodeURIComponent).join("/");
const itemUrl = (segs) => `${GRAPH}/me/drive/root:/${enc(segs)}:`;

async function g(url, opts = {}) {
  const t = await token();
  return fetch(url, { ...opts, headers: { Authorization: `Bearer ${t}`, ...(opts.headers || {}) } });
}

async function gError(res, what) {
  let detail = "";
  try { detail = (await res.json()).error?.message || ""; } catch {}
  return new Error(`${what} 失败 (${res.status}) ${detail}`);
}

const childCache = new Map();
async function listChildren(segs) {
  const key = segs.join("/");
  if (childCache.has(key)) return childCache.get(key);
  const res = await g(`${itemUrl(segs)}/children?$select=name,folder&$top=999`);
  if (res.status === 404) return [];
  if (!res.ok) throw await gError(res, "读取文件夹");
  const names = (await res.json()).value.filter((x) => x.folder).map((x) => x.name);
  childCache.set(key, names);
  return names;
}

async function ensureFolder(segs) {
  const parts = segs.join("/").split("/");
  const name = parts.pop();
  const res = await g(`${itemUrl(parts)}/children`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ name, folder: {}, "@microsoft.graph.conflictBehavior": "fail" }),
  });
  if (!res.ok && res.status !== 409) throw await gError(res, `新建文件夹 ${name}`);
  childCache.clear();
}

async function readText(segs) {
  const res = await g(itemUrl(segs));
  if (res.status === 404) return null;
  if (!res.ok) throw await gError(res, "读取文件");
  const url = (await res.json())["@microsoft.graph.downloadUrl"];
  const file = await fetch(url);
  return (await file.text()).replace(/^﻿/, "");
}

async function upload(segs, blob, conflict) {
  if (blob.size < 4 * 1024 * 1024) {
    return g(`${itemUrl(segs)}/content?@microsoft.graph.conflictBehavior=${conflict}`, { method: "PUT", body: blob });
  }
  const s = await g(`${itemUrl(segs)}/createUploadSession`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": conflict } }),
  });
  if (!s.ok) return s;
  const { uploadUrl } = await s.json();
  const CHUNK = 320 * 1024 * 16; // 5 MiB, a multiple of 320 KiB as Graph requires
  let res;
  for (let start = 0; start < blob.size; start += CHUNK) {
    const end = Math.min(start + CHUNK, blob.size);
    res = await fetch(uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": `bytes ${start}-${end - 1}/${blob.size}` },
      body: blob.slice(start, end),
    });
    if (!res.ok && res.status !== 202) return res;
  }
  return res;
}

async function uploadUnique(folderSegs, name, blob) {
  for (let n = 1; n <= 30; n++) {
    const candidate = n === 1 ? name : L.withSuffix(name, n);
    const res = await upload([...folderSegs, candidate], blob, "fail");
    if (res.ok) return candidate;
    if (res.status !== 409) throw await gError(res, "上传收据");
  }
  throw new Error("同名文件太多");
}

async function appendRegister(regSegs, row) {
  const text = (await readText(regSegs)) ?? L.REGISTER_HEADER.map((h) => `"${h}"`).join(",");
  const body = new Blob(["﻿" + text.replace(/\s*$/, "") + "\r\n" + row + "\r\n"], { type: "text/csv" });
  const res = await upload(regSegs, body, "replace");
  if (!res.ok) throw await gError(res, "写入登记表");
}

// Rows that could not be written (e.g. CSV open in Excel) are retried next time.
async function flushPendingRows() {
  const pending = store.get("pendingRows", []);
  const left = [];
  for (const p of pending) {
    try { await appendRegister(p.reg, p.row); } catch { left.push(p); }
  }
  store.set("pendingRows", left);
  return left.length;
}

// ---------- settings ----------
let cfg = store.get("cfg", null);
async function loadSettings() {
  const text = await readText(SETTINGS_PATH.split("/"));
  if (!text) throw new Error(`找不到设置文件 ${SETTINGS_PATH}`);
  cfg = JSON.parse(text);
  store.set("cfg", cfg);
}

// ---------- Claude ----------
const SCHEMA = {
  type: "object",
  additionalProperties: false,
  properties: {
    is_receipt: { type: "boolean", description: "true if this is a receipt or invoice for a payment" },
    multiple_receipts: { type: "boolean", description: "true if the image contains more than one separate receipt" },
    date: { type: "string", description: "transaction date as YYYY-MM-DD, empty if unreadable" },
    merchant: { type: "string", description: "short trading name, e.g. Screwfix, Dunelm, Tesco" },
    purpose: { type: "string", description: "2-6 English words: what was bought or why" },
    property: { type: "string", description: "property short code if identifiable, else empty" },
    currency: { type: "string", description: "ISO currency code, normally GBP" },
    total: { type: "number", description: "gross amount paid, 0 if unreadable" },
    vat_amount: { type: "number", description: "VAT amount shown, 0 if none" },
    supplier_vat_number: { type: "string", description: "supplier VAT registration number, digits/GB only, empty if none" },
    card_last4: { type: "string", description: "last 4 digits of the paying card, empty if not shown or ambiguous" },
    payment_method: { type: "string", description: "card, cash, online, or unknown" },
    notes: { type: "string", description: "in Simplified Chinese: anything uncertain Leon should check; empty if all clear" },
  },
  required: [
    "is_receipt", "multiple_receipts", "date", "merchant", "purpose", "property", "currency", "total",
    "vat_amount", "supplier_vat_number", "card_last4", "payment_method", "notes",
  ],
};

function systemPrompt() {
  return `You read UK receipts and invoices for a property management company and extract fields for bookkeeping.

Rules:
- Dates on UK receipts are DD/MM/YYYY or DD/MM/YY. Use the transaction date, not a print or due date.
- card_last4: the last 4 digits of the masked card number (e.g. "****4019", "XXXXXXXXXXXX4019", "ending 4019", "Card ************5695"). Ignore auth codes, terminal IDs (TID), merchant IDs (MID), AIDs, order and invoice numbers. If several masked numbers appear and none is clearly the card, leave it empty and say so in notes.
- supplier_vat_number: the supplier's VAT registration number ("VAT No", "VAT Reg", "GB123456789"). Not company registration numbers.
- vat_amount: the VAT actually shown (a VAT line, or a VAT breakdown). If the receipt only says prices include VAT without an amount, but shows a rate, compute it from the gross; otherwise 0.
- property: identify the property only from the receipt itself (delivery address, item text, handwritten note). Code rule: ${cfg?.propertyCodeRule || "house number + street initials"}. Known properties: ${(cfg?.properties || []).join("; ")}.
- purpose: short and specific, e.g. "repair light", "paint", "parking", "mattress protector".
- Food and drink (restaurants, cafes, coffee, McDonald's, takeaway, snacks, supermarket meal deals) always get exactly one of these purposes: "Working Lunch" if the transaction time is before 17:00 or not shown, "Working Dinner" if 17:00 or later. Never describe the food itself.
- notes: write in Simplified Chinese, only for things Leon should double-check.`;
}

async function extract(blob, mediaType) {
  const apiKey = store.get("claudeKey", "");
  if (!apiKey) throw new Error("请先在设置里填 Claude API key");
  const client = new Anthropic({ apiKey, dangerouslyAllowBrowser: true });
  const data = await blobToBase64(blob);
  const source = { type: "base64", media_type: mediaType, data };
  const content = [
    mediaType === "application/pdf" ? { type: "document", source } : { type: "image", source },
    { type: "text", text: "Extract this receipt." },
  ];
  const params = {
    model: CLAUDE_MODEL,
    max_tokens: 16000,
    system: systemPrompt(),
    output_config: { effort: "medium", format: { type: "json_schema", schema: SCHEMA } },
    messages: [{ role: "user", content }],
  };
  let msg;
  try {
    // Server-side fallback: if a safety classifier declines, the API retries on a fallback model.
    msg = await client.beta.messages.create({ ...params, betas: ["server-side-fallback-2026-07-01"], fallbacks: "default" });
  } catch (e) {
    if (e?.status !== 400) throw e;
    msg = await client.messages.create(params);
  }
  if (msg.stop_reason === "refusal") throw new Error("Claude 拒绝处理这张图片");
  if (msg.stop_reason === "max_tokens") throw new Error("识别输出被截断，请重试");
  const text = msg.content.filter((b) => b.type === "text").map((b) => b.text).join("");
  return JSON.parse(text);
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(String(r.result).split(",")[1]);
    r.onerror = reject;
    r.readAsDataURL(blob);
  });
}

// Shrink phone photos: 1800px on the long edge is still sharp for receipts and
// above what Claude reads (it downsizes to ~1568px anyway).
async function prepareImage(file) {
  const bmp = await createImageBitmap(file, { imageOrientation: "from-image" });
  const scale = Math.min(1, 1800 / Math.max(bmp.width, bmp.height));
  const c = document.createElement("canvas");
  c.width = Math.round(bmp.width * scale);
  c.height = Math.round(bmp.height * scale);
  c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
  const jpeg = await new Promise((res) => c.toBlob(res, "image/jpeg", 0.75));
  return { jpeg, width: c.width, height: c.height };
}

// ---------- UI ----------
// One item per receipt. status: reading | ready | check | filing | filed | error
let queue = [];
let current = null; // the item open in the review form
let batchMode = false;
let nextId = 1;

function show(view) {
  for (const v of ["home", "busy", "review", "done", "batch", "settings"]) $(v).hidden = v !== view;
}

function setBusy(text) { $("busy-text").textContent = text; show("busy"); }

function toast(text, isError) {
  const t = $("toast");
  t.textContent = text;
  t.className = isError ? "toast error" : "toast";
  t.hidden = false;
  clearTimeout(toast.timer);
  toast.timer = setTimeout(() => (t.hidden = true), isError ? 8000 : 3500);
}

const isPdfFile = (f) => f.type === "application/pdf" || /\.pdf$/i.test(f.name);

// Read one file with Claude and turn it into a queue item.
async function analyse(item) {
  const { file } = item;
  const isPdf = isPdfFile(file);
  // Photos: Claude reads the JPEG; OneDrive gets the same JPEG wrapped in a PDF.
  const img = isPdf ? null : await prepareImage(file);
  item.mediaType = isPdf ? "application/pdf" : "image/jpeg";
  item.previewBlob = img?.jpeg;
  const x = await extract(isPdf ? file : img.jpeg, item.mediaType);
  item.blob = isPdf ? file : await jpegToPdf(img.jpeg, img.width, img.height);
  item.ext = "pdf";
  item.sourceName = file.name || "camera.jpg";
  item.data = {
    date: x.date, merchant: x.merchant, purpose: x.purpose, property: x.property,
    currency: x.currency || "GBP", total: x.total, vat: x.vat_amount, vatNumber: x.supplier_vat_number,
    cardLast4: x.card_last4, isVat: Boolean(x.supplier_vat_number) && x.vat_amount > 0,
  };
  item.notes = [
    !x.is_receipt && "这看起来不是收据/发票。",
    x.multiple_receipts && "图里好像有多张收据，建议分开拍。",
    x.notes,
  ].filter(Boolean).join(" ");
  item.manual = null;
}

// Work out company, card, file name, target folder and duplicate flags for an item.
async function planItem(item) {
  const d = item.data;
  const { company, card, known } = L.resolveCard(cfg.cards, d.cardLast4, item.manual);
  Object.assign(item, { company, card, known, plan: null, name: null, dup: false });
  if (!company || !L.parseISO(d.date) || !(d.total > 0)) return;
  d.company = company;
  item.name = L.buildFileName(d, item.ext);
  item.plan = await L.planTarget(company, cfg, d, card, listChildren);
  const reg = await readText(item.plan.register).catch(() => null);
  const sameInBatch = queue.some((o) => o !== item && o.data && o.status !== "error" &&
    o.data.date === d.date && Number(o.data.total) === Number(d.total) && o.company === company);
  item.dup = Boolean((reg && L.findDuplicate(reg, d)) || sameInBatch);
}

// "ready" = nothing for Leon to check: known company card, date + amount read, no notes, no duplicate.
function judge(item) {
  item.status = item.confirmed || (item.plan && item.known && !item.manual && !item.notes && !item.dup) ? "ready" : "check";
}

async function fileItem(item) {
  const { plan, name, data: d } = item;
  for (const segs of plan.create) await ensureFolder(segs);
  const saved = await uploadUnique(plan.folder, name, item.blob);
  const relPath = [...plan.folder.slice(1), saved].join("/");
  const row = L.registerRow(d, item.card, relPath, item.sourceName);
  let regNote = "";
  try {
    await flushPendingRows();
    await appendRegister(plan.register, row);
  } catch {
    store.set("pendingRows", [...store.get("pendingRows", []), { reg: plan.register, row }]);
    regNote = "（登记表暂时写不进去，下次自动补写）";
  }
  const hist = store.get("history", []);
  hist.unshift({ name: saved, folder: plan.folder.join(" / "), at: new Date().toISOString() });
  store.set("history", hist.slice(0, 30));
  Object.assign(item, { status: "filed", saved, regNote });
  childCache.clear();
}

// ----- single photo -----
async function onSingle(file) {
  if (!file) return;
  batchMode = false;
  const item = { id: nextId++, file, status: "reading" };
  queue = [item];
  try {
    setBusy("准备图片…");
    // Re-read the OneDrive settings every time, so folder/card changes apply without restarting the app.
    await loadSettings().catch(() => {});
    setBusy("Claude 正在识别…");
    await analyse(item);
    openReview(item);
  } catch (e) {
    console.error(e);
    show("home");
    toast(e.message || String(e), true);
  }
}

// ----- batch -----
async function onMany(files) {
  if (!files.length) return;
  if (files.length === 1 && !batchMode) return onSingle(files[0]);
  if (!batchMode) queue = [];
  batchMode = true;
  await loadSettings().catch(() => {});
  const items = files.map((file) => ({ id: nextId++, file, status: "reading" }));
  queue.push(...items);
  renderBatch();
  show("batch");
  // Three at a time: quick, and gentle on API rate limits.
  let i = 0;
  const worker = async () => {
    while (i < items.length) {
      const item = items[i++];
      try {
        await analyse(item);
        await planItem(item);
        judge(item);
      } catch (e) {
        console.error(e);
        Object.assign(item, { status: "error", error: e.message || String(e) });
      }
      renderBatch();
    }
  };
  await Promise.all([worker(), worker(), worker()]);
}

const STATUS = {
  reading: "识别中…", ready: "✓ 可归档", check: "⚠ 需要确认", filing: "归档中…", filed: "已归档", error: "出错",
};

function issueText(item) {
  if (item.status === "error") return item.error;
  if (item.status !== "check") return "";
  const d = item.data;
  if (!item.company) return `卡号 ${d.cardLast4 ? "****" + d.cardLast4 : "没看到"}，请选公司`;
  if (!item.plan) return "日期或金额没读到";
  if (item.dup) return "可能重复";
  if (!item.known) return "非公司卡 → OTHER CARD";
  return item.notes || "请看一下";
}

// The list is keyed by item id and updated in place, so the edit form that is open
// inside one row keeps its focus while other receipts finish reading.
const rows = new Map();
let openId = null;

function rowFor(item) {
  let r = rows.get(item.id);
  if (r) return r;
  const li = document.createElement("li");
  const head = document.createElement("div");
  head.className = "bhead";
  const thumb = document.createElement("div");
  thumb.className = "thumb";
  const text = document.createElement("div");
  text.className = "btext";
  const title = document.createElement("div");
  title.className = "bname";
  const sub = document.createElement("div");
  sub.className = "bsub";
  text.append(title, sub);
  const x = document.createElement("button");
  x.className = "ghost bx";
  x.textContent = "✕";
  x.setAttribute("aria-label", "移除");
  x.addEventListener("click", (e) => {
    e.stopPropagation();
    if (openId === item.id) collapse();
    queue = queue.filter((o) => o !== item);
    renderBatch();
  });
  head.append(thumb, text, x);
  head.addEventListener("click", () => {
    if (!["ready", "check"].includes(item.status)) return;
    if (openId === item.id) collapse(); else expand(item);
  });
  li.appendChild(head);
  r = { li, thumb, title, sub, x };
  rows.set(item.id, r);
  return r;
}

function renderBatch() {
  const ul = $("batch-list");
  for (const [id, r] of rows) {
    if (!queue.some((i) => i.id === id)) { r.li.remove(); rows.delete(id); }
  }
  queue.forEach((item, idx) => {
    const r = rowFor(item);
    r.li.className = `bitem s-${item.status}${openId === item.id ? " open" : ""}`;
    if (item.previewBlob && !r.thumb.firstChild?.src) {
      r.thumb.textContent = "";
      const img = document.createElement("img");
      img.src = item.thumbUrl ||= URL.createObjectURL(item.previewBlob);
      r.thumb.appendChild(img);
    } else if (!item.previewBlob) r.thumb.textContent = item.status === "reading" ? "…" : "📄";
    const d = item.data;
    r.title.textContent = item.saved || item.name ||
      (d ? [d.date, d.merchant, d.total ? `GBP ${Number(d.total).toFixed(2)}` : ""].filter(Boolean).join(" · ") : "") ||
      item.file.name || "收据";
    r.sub.textContent = [STATUS[item.status], issueText(item)].filter(Boolean).join(" · ");
    r.x.hidden = !["ready", "check", "error"].includes(item.status);
    if (ul.children[idx] !== r.li) ul.insertBefore(r.li, ul.children[idx] || null);
  });

  const count = (s) => queue.filter((i) => i.status === s).length;
  const ready = count("ready"), check = count("check"), reading = count("reading");
  $("batch-title").textContent =
    `共 ${queue.length} 张 · 可归档 ${ready} · 待确认 ${check} · 已归档 ${count("filed")}` +
    (reading ? ` · 识别中 ${reading}` : "");
  const btn = $("batch-file");
  btn.disabled = !ready || reading > 0 || check > 0;
  btn.textContent = reading ? `识别中，请稍候…`
    : check ? `还有 ${check} 张待确认`
    : ready ? `一键归档 ${ready} 张` : "没有要归档的";
}

// Open the edit form inside the row (accordion: one at a time).
function expand(item) {
  if (openId !== null && openId !== item.id) collapse();
  openId = item.id;
  const r = rowFor(item);
  r.li.appendChild($("review-card"));
  fillForm(item);
  $("cancel").textContent = "收起";
  $("confirm").textContent = "确认";
  renderBatch();
  r.li.scrollIntoView({ behavior: "smooth", block: "start" });
}

function collapse() {
  $("review").appendChild($("review-card"));
  openId = null;
  renderBatch();
}

function expandNextCheck() {
  const next = queue.find((i) => i.status === "check");
  if (next) expand(next); else collapse();
}

async function fileAllReady() {
  if (openId !== null) collapse();
  for (const item of queue.filter((i) => i.status === "ready")) {
    item.status = "filing";
    renderBatch();
    try {
      await fileItem(item);
    } catch (e) {
      Object.assign(item, { status: "check", confirmed: false, notes: `上传失败：${e.message}` });
    }
    renderBatch();
  }
  renderHistory();
  const left = queue.filter((i) => i.status === "check").length;
  toast(left ? `有 ${left} 张没归档成功，请再确认` : "全部归档完成");
}

// ----- review form (single photo page, or inline in a batch row) -----
const FIELDS = [
  ["date", "f-date"], ["merchant", "f-merchant"], ["purpose", "f-purpose"], ["property", "f-property"],
  ["total", "f-total"], ["vat", "f-vat"], ["vatNumber", "f-vatno"], ["cardLast4", "f-card"],
];

function fillForm(item) {
  current = item;
  const d = item.data;
  for (const [k, id] of FIELDS) $(id).value = d[k] ?? "";
  $("f-isvat").checked = d.isVat;
  $("notes").textContent = item.notes;
  $("notes").hidden = !item.notes;
  const prev = $("preview");
  prev.innerHTML = "";
  if (item.previewBlob) {
    const img = document.createElement("img");
    img.src = item.thumbUrl ||= URL.createObjectURL(item.previewBlob);
    prev.appendChild(img);
  } else {
    prev.textContent = "📄 PDF";
  }
  updatePlan();
}

function openReview(item) {
  $("review").appendChild($("review-card"));
  $("cancel").textContent = "取消";
  $("confirm").textContent = "确认归档";
  fillForm(item);
  show("review");
}

function readForm() {
  const d = current.data;
  for (const [k, id] of FIELDS) d[k] = $(id).value.trim();
  d.total = Number(d.total) || 0;
  d.vat = Number(d.vat) || 0;
  d.isVat = $("f-isvat").checked;
}

let planSeq = 0;
async function updatePlan() {
  readForm();
  const item = current;
  const d = item.data;
  const seq = ++planSeq;
  const { company, known } = L.resolveCard(cfg.cards, d.cardLast4, item.manual);
  for (const b of document.querySelectorAll("[data-company]")) b.classList.toggle("on", b.dataset.company === company);
  const warn = $("card-warn");
  warn.hidden = known;
  warn.textContent = !company
    ? `卡号 ${d.cardLast4 ? "****" + d.cardLast4 : "（没看到）"} 不是已登记的卡 —— 请选择公司`
    : `非公司卡（${d.cardLast4 ? "****" + d.cardLast4 : "没有卡号"}）→ 放进 OTHER CARD，标记需报销`;
  $("confirm").disabled = true;
  $("dup-warn").hidden = true;
  $("plan").textContent = company ? "…" : "选择公司后显示归档位置";
  try {
    await planItem(item);
    if (seq !== planSeq) return;
    if (!item.plan) { $("plan").textContent = company ? "请补全日期和金额" : "选择公司后显示归档位置"; return; }
    $("plan").textContent = `${item.name}\n→ ${item.plan.folder.join(" / ")}${item.plan.create.length ? "（新建）" : ""}`;
    $("dup-warn").hidden = !item.dup;
    $("confirm").disabled = false;
  } catch (e) {
    if (seq === planSeq) $("plan").textContent = e.message;
  }
}

async function confirmFile() {
  readForm();
  await updatePlan();
  const item = current;
  if (!item.plan) return;
  if (batchMode) {
    // Batch: confirming only marks it; everything is filed together at the end.
    item.confirmed = true;
    item.status = "ready";
    expandNextCheck();
    return;
  }
  try {
    setBusy("存入 OneDrive…");
    await fileItem(item);
    renderHistory();
    $("done-name").textContent = item.saved;
    $("done-folder").textContent = item.plan.folder.join(" / ") + item.regNote;
    show("done");
  } catch (e) {
    console.error(e);
    show("review");
    toast(e.message, true);
  }
}

function leaveReview() {
  if (batchMode) collapse(); else show("home");
}

function renderHistory() {
  const ul = $("history");
  ul.innerHTML = "";
  for (const h of store.get("history", []).slice(0, 10)) {
    const li = document.createElement("li");
    li.textContent = h.name;
    li.title = h.folder;
    ul.appendChild(li);
  }
  $("history-wrap").hidden = !ul.children.length;
}

function renderSettings() {
  $("s-key").value = store.get("claudeKey", "");
  const acc = account();
  $("s-account").textContent = acc ? acc.username : "未登录";
  $("s-cards").textContent = cfg
    ? cfg.cards.map((c) => `****${c.last4} ${c.company} ${c.card} ${c.holder || ""}`).join("\n") +
      "\n\n" + Object.entries(cfg.companies).map(([k, c]) => `${k} → ${c.root}`).join("\n")
    : "未加载";
}

function closeBatch() {
  const open = queue.filter((i) => ["reading", "ready", "check", "filing"].includes(i.status)).length;
  if (open && !confirm(`还有 ${open} 张没归档，确定离开？（没归档的会被丢弃）`)) return;
  if (openId !== null) collapse();
  queue = [];
  renderBatch();
  batchMode = false;
  show("home");
}

async function main() {
  $("cam").addEventListener("change", (e) => { onSingle(e.target.files[0]); e.target.value = ""; });
  $("pick").addEventListener("change", (e) => { onMany([...e.target.files]); e.target.value = ""; });
  $("batch-add").addEventListener("change", (e) => { onMany([...e.target.files]); e.target.value = ""; });
  $("batch-file").addEventListener("click", fileAllReady);
  $("batch-close").addEventListener("click", closeBatch);
  for (const [, id] of FIELDS) $(id).addEventListener("input", () => updatePlan());
  $("f-isvat").addEventListener("change", () => updatePlan());
  for (const b of document.querySelectorAll("[data-company]")) {
    b.addEventListener("click", () => { current.manual = b.dataset.company; updatePlan(); });
  }
  for (const b of document.querySelectorAll("[data-purpose]")) {
    b.addEventListener("click", () => { $("f-purpose").value = b.dataset.purpose; updatePlan(); });
  }
  $("confirm").addEventListener("click", confirmFile);
  $("cancel").addEventListener("click", leaveReview);
  $("again").addEventListener("click", () => show("home"));
  $("open-settings").addEventListener("click", () => { renderSettings(); show("settings"); });
  $("close-settings").addEventListener("click", () => show(batchMode ? "batch" : "home"));
  $("s-save").addEventListener("click", () => { store.set("claudeKey", $("s-key").value.trim()); toast("已保存"); show("home"); });
  $("s-reload").addEventListener("click", async () => {
    try { await loadSettings(); renderSettings(); toast("设置已更新"); } catch (e) { toast(e.message, true); }
  });
  $("s-signout").addEventListener("click", () => msalApp.logoutRedirect({ account: account() }));
  $("signin").addEventListener("click", () => msalApp.loginRedirect({ scopes: SCOPES }));

  renderHistory();
  show("home");
  if (MS_CLIENT_ID.startsWith("PASTE")) { toast("还没填 Microsoft 应用 ID（config.js）", true); return; }

  await msalApp.initialize();
  const r = await msalApp.handleRedirectPromise().catch((e) => { toast(e.message, true); return null; });
  if (r?.account) msalApp.setActiveAccount(r.account);
  const acc = account();
  $("signin").hidden = Boolean(acc);
  $("who").textContent = acc ? acc.username : "";
  if (!acc) return;
  try {
    await loadSettings();
    const left = await flushPendingRows();
    if (left) toast(`还有 ${left} 条登记未写入`, true);
  } catch (e) {
    if (!cfg) toast(e.message, true);
  }
  if (!store.get("claudeKey", "")) { renderSettings(); show("settings"); }
}

main();
