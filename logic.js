// Pure filing rules — no network. Shared by app.js and tests.html.

const pad = (n) => String(n).padStart(2, "0");

// "2026-05-19" -> {y, m, d}
export function parseISO(iso) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || "");
  if (!m) return null;
  return { y: +m[1], m: +m[2], d: +m[3] };
}

// JML financial year runs 1 Feb – 31 Jan.
export function fyStartYear({ y, m }) {
  return m >= 2 ? y : y - 1;
}

// Year folders look like "8 Reciept 01 02 2026 _31 01 2027" (sequence number = start year - 2018).
export function canonicalYearFolder(date) {
  const s = fyStartYear(date);
  return `${s - 2018} Reciept 01 02 ${s} _31 01 ${s + 1}`;
}

export function matchYearFolder(names, date) {
  const start = `01 02 ${fyStartYear(date)}`;
  return names.find((n) => /RECIEPT|RECEIPT/.test(norm(n)) && norm(n).includes(start)) || null;
}

// VAT quarters: Feb–Apr, May–Jul, Aug–Oct, Nov–Jan.
export function vatQuarter(date) {
  const s = fyStartYear(date);
  const qs = [
    { from: [1, 2, s], to: [30, 4, s] },
    { from: [1, 5, s], to: [31, 7, s] },
    { from: [1, 8, s], to: [31, 10, s] },
    { from: [1, 11, s], to: [31, 1, s + 1] },
  ];
  const idx = date.m === 1 ? 3 : Math.floor((date.m - 2) / 3);
  return qs[idx];
}

const fmt = ([d, m, y]) => `${pad(d)} ${pad(m)} ${y}`;

// Canonical folder name, used only when no existing folder matches.
export function canonicalFolder(date, isVat, card) {
  const s = fyStartYear(date);
  const range = isVat ? vatQuarter(date) : { from: [1, 2, s], to: [31, 1, s + 1] };
  return `${isVat ? "VAT " : ""}${card} CARD ${fmt(range.from)}_${fmt(range.to)}`;
}

const norm = (name) =>
  name.toUpperCase().replace(/CREDT/g, "CREDIT").replace(/[\s_]+/g, " ").trim();

// Find an existing folder among `names` that matches; else null.
// Existing names are inconsistent ("VAT CREDT CARD 01 05 2026_31 07 2026", "DEBIT CARD 01 02 2026_ 30 01 2027").
export function matchFolder(names, date, isVat, card) {
  const s = fyStartYear(date);
  const start = isVat ? fmt(vatQuarter(date).from) : fmt([1, 2, s]);
  return (
    names.find((n) => {
      const x = norm(n);
      return x.startsWith("VAT") === isVat && x.includes(`${card} CARD`) && x.includes(start);
    }) || null
  );
}

const clean = (s) =>
  String(s || "")
    .replace(/&/g, "and")
    .replace(/[\\/:*?"<>|#%]/g, " ")
    .replace(/\s+/g, " ")
    .replace(/[ ,.]+$/, "")
    .trim();

export function buildFileName(r, ext) {
  const d = parseISO(r.date);
  const parts = [r.merchant, r.purpose, r.property].map(clean).filter(Boolean);
  const hasVat = r.isVat && Number(r.vat) > 0;
  const amount = `${r.currency || "GBP"} ${Number(r.total).toFixed(2)}${hasVat ? `, VAT ${Number(r.vat).toFixed(2)}` : ""}`;
  return `${d.y} ${pad(d.m)} ${pad(d.d)} ${r.company}, ${parts.join(", ")}, ${amount}.${ext}`;
}

export function withSuffix(name, n) {
  const i = name.lastIndexOf(".");
  return `${name.slice(0, i)} (${n})${name.slice(i)}`;
}

// Decide company + card label from the card table and an optional manual override.
// Returns {company, card, known} — company null means "ask Leon".
export function resolveCard(cards, last4, manualCompany) {
  const hit = cards.find((c) => c.last4 === (last4 || "").trim());
  if (manualCompany) {
    if (hit && hit.company === manualCompany) return { company: manualCompany, card: hit.card, known: true };
    return { company: manualCompany, card: "OTHER", known: false };
  }
  if (hit) return { company: hit.company, card: hit.card, known: true };
  return { company: null, card: "OTHER", known: false };
}

// Folder path segments under the company root, plus the register path.
// `listChildren(pathSegments)` must return the child folder names (or [] if missing).
export async function planTarget(company, cfg, r, card, listChildren) {
  const co = cfg.companies[company];
  const date = parseISO(r.date);
  if (co.layout === "flat") {
    return { folder: [co.root], create: [], register: [co.root, co.register || "Receipt Register.csv"] };
  }
  if (co.layout === "pu") return planPU(co, date, card, listChildren);
  if (co.layout === "inv") return planINV(co, date, card, listChildren);
  if (co.layout === "client") return planClient(co, r, cfg.projects);
  const yearFound = matchYearFolder(await listChildren([co.root]), date);
  const year = yearFound || canonicalYearFolder(date);
  const existing = yearFound ? await listChildren([co.root, year]) : [];
  const found = matchFolder(existing, date, r.isVat, card);
  const sub = found || canonicalFolder(date, r.isVat, card);
  const s = fyStartYear(date);
  return {
    folder: [co.root, year, sub],
    create: [...(yearFound ? [] : [[co.root, year]]), ...(found ? [] : [[co.root, year, sub]])],
    register: [co.root, year, `Receipt Register ${s}_${s + 1}.csv`],
  };
}

// ----- PropUnited: one folder per company year, no VAT/card split, personal/cash in a subfolder,
// and every receipt logged in the Excel ledgers instead of a CSV register.

// Company year starts on `fyStart` ("MM-DD", e.g. "10-15" = 15 Oct).
export function fyStartYearOn({ y, m, d }, fyStart) {
  const [sm, sd] = fyStart.split("-").map(Number);
  return m > sm || (m === sm && d >= sd) ? y : y - 1;
}

// "01 reciept 2025 10 15_2026 10 14" (number = years since the first year).
export function puYearFolder(date, co) {
  const s = fyStartYearOn(date, co.fyStart);
  const [sm, sd] = co.fyStart.split("-").map(Number);
  const end = new Date(Date.UTC(s + 1, sm - 1, sd - 1));
  const n = s - co.firstYear + 1;
  return `${pad(n)} reciept ${s} ${pad(sm)} ${pad(sd)}_${end.getUTCFullYear()} ${pad(end.getUTCMonth() + 1)} ${pad(end.getUTCDate())}`;
}

export function matchPuYearFolder(names, date, co) {
  const s = fyStartYearOn(date, co.fyStart);
  const [sm, sd] = co.fyStart.split("-").map(Number);
  const start = `${s} ${pad(sm)} ${pad(sd)}`;
  return names.find((n) => /RECIEPT|RECEIPT/.test(norm(n)) && norm(n).includes(start)) || null;
}

async function planPU(co, date, card, listChildren) {
  const yearFound = matchPuYearFolder(await listChildren([co.root]), date, co);
  const year = yearFound || puYearFolder(date, co);
  const personal = card === "OTHER";
  const folder = personal ? [co.root, year, co.personalFolder] : [co.root, year];
  const subExists = personal && yearFound && (await listChildren([co.root, year])).includes(co.personalFolder);
  return {
    folder,
    create: [...(yearFound ? [] : [[co.root, year]]), ...(personal && !subExists ? [folder] : [])],
    register: null,
    ledger: personal ? { kind: "expense", ...co.expenses } : { kind: "bank", ...co.ledger },
  };
}

// ----- JML Investment: year runs 1 Dec – 30 Nov, receipts sit straight in the year's tax folder
// ("07 0112 2025-30 11 2026 报税"), personal/cash in a subfolder, no company ledger.
export function invYearFolder(date) {
  const s = fyStartYearOn(date, "12-01");
  return `${pad(s - 2018)} 0112 ${s}-30 11 ${s + 1} 报税`;
}

export function matchInvYear(names, date) {
  const s = fyStartYearOn(date, "12-01");
  return names.find((n) => norm(n).includes(`0112 ${s}-30 11 ${s + 1}`)) || null;
}

async function planINV(co, date, card, listChildren) {
  const yearFound = matchInvYear(await listChildren([co.root]), date);
  const year = yearFound || invYearFolder(date);
  const personal = card === "OTHER";
  const folder = personal ? [co.root, year, co.personalFolder] : [co.root, year];
  const subExists = personal && yearFound && (await listChildren([co.root, year])).includes(co.personalFolder);
  return {
    folder,
    create: [...(yearFound ? [] : [[co.root, year]]), ...(personal && !subExists ? [folder] : [])],
    register: null,
    ledger: null,
  };
}

// ----- Projects (e.g. 174CA): any company's receipt can also be logged in the project's own sheet.
// The note says who paid, because a JML card paying for 174 keeps the cost (and VAT) in JML's books.
export function projectNote(company, card, payment) {
  // A 代开发 bank transfer has no card; who paid isn't on the screenshot, so leave the note to Leon.
  if (company === "CLIENT" && payment === "bank_transfer") return "";
  if (card === "OTHER") return "个人垫付";
  // CLIENT = the client's own card / transfer for a 代开发 project: nothing to note.
  return { INV: "investment", JML: "Letting 垫付", PU: "PropUnited 垫付", CLIENT: "" }[company] ?? company;
}

const compact = (s) => String(s || "").toUpperCase().replace(/[^A-Z0-9]/g, "");

// Which project a receipt belongs to before Leon picks one: the company's default, else a project
// whose code appears in the property Claude read (e.g. "174 CA" → 174CA).
export function guessProject(projects, company, cfgCompany, property) {
  const p = compact(property);
  const names = (pr) => [pr.code, ...(pr.aliases || [])].map(compact);
  const hit = (projects || []).find((pr) => p && names(pr).some((n) => p.includes(n) || n.includes(p)));
  return hit?.code || cfgCompany?.defaultProject || null;
}

const PAYMENT = { online: "online", card: "shop", cash: "cash", bank_transfer: "bank transfer" };

// ----- 代开发 (client projects): no company books. The receipt goes to the project's own receipt
// folder, named with the sheet's 编号 in front ("010 2026 11 21 39SA, Selco, …").
function planClient(co, r, projects) {
  const project = (projects || []).find((p) => p.code === r.property && p.company === "CLIENT");
  if (!project) return null; // a client project must be chosen
  return { folder: [project.folder], create: [], register: null, ledger: null };
}

// Excel stores dates as days since 1899-12-30.
export function excelSerial(iso) {
  const { y, m, d } = parseISO(iso);
  return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000;
}

export const ukDate = (iso) => { const { y, m, d } = parseISO(iso); return `${pad(d)}/${pad(m)}/${y}`; };

// Values for one ledger row. Bank sheet: columns D..P; Expense Claims sheet: C..H (+ J).
export function ledgerCells(kind, r) {
  const total = Number(Number(r.total).toFixed(2));
  const explanation = [r.purpose, r.property].filter(Boolean).join(" · ");
  if (kind === "bank") {
    //       D Description, E Paid Out, F In, G Bal, H, I Explanation, J, K Receipt, L Filing, M, N Net, O VAT, P Gross
    return { first: "D", last: "P", values: [r.merchant, total, "", "", "", explanation, "", "Y", "", "", "", "", total] };
  }
  if (kind === "project") {
    //       D 退货, E 支出, F 内容, G 商店, H 方式, I 备注   (C 编号 is pre-filled / handled separately)
    return { first: "D", last: "I", values: ["", total, r.purpose, r.merchant, PAYMENT[r.payment] || "", r.note || ""] };
  }
  //       C Supplier, D Description, E Net, F VAT, G Total, H Receipt
  return { first: "C", last: "H", values: [r.merchant, explanation, total, "", total, "Y"], reimbursed: "N" };
}

// Ledger duplicate: a row with the same date and the same amount.
// `values` is usedRange values (2D, starting at A1). Real dates come back as Excel serial numbers
// (Excel Online's *text* is US-formatted, so don't compare text); dates typed as text are dd/mm/yyyy.
// Amount column E (bank) or G (expense).
// Row number (1-based) of the same receipt in a ledger/project sheet, or 0.
// Merchant column: bank D, expense C, project G.
// Merchant column: bank D, expense C, project G. Description: bank I, expense D, project F.
export function findLedgerRow(values, kind, r) {
  const amountCol = kind === "expense" ? 6 : 4;
  const merchantCol = { bank: 3, expense: 2, project: 6 }[kind];
  const purposeCol = { bank: 8, expense: 3, project: 5 }[kind];
  const i = values.findIndex((row, idx) => {
    if (idx < 3 || !row) return false;
    const date = cellDate(row[1]);
    if (!date) return false;
    const total = String(row[amountCol] ?? "").replace(/[£,]/g, "");
    if (total === "" || isNaN(Number(total))) return false;
    return samePurchase({ date, merchant: row[merchantCol], purpose: row[purposeCol], total }, r);
  });
  return i + 1;
}

// A sheet date: an Excel serial number, or text typed as dd/mm/yyyy.
function cellDate(v) {
  if (typeof v === "number" && v > 30000) {
    const t = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
    return `${t.getUTCFullYear()}-${pad(t.getUTCMonth() + 1)}-${pad(t.getUTCDate())}`;
  }
  const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(v ?? "").trim());
  return m ? `${m[3]}-${pad(m[2])}-${pad(m[1])}` : null;
}

export const ledgerHasDup = (values, kind, r) => findLedgerRow(values, kind, r) > 0;

// Next free row: bank sheet = after the last row with anything in B..P (data starts at row 15);
// Expense Claims = first row 17..51 with no date, supplier or total (EX numbers are pre-filled in I).
export function nextLedgerRow(text, kind) {
  if (kind === "project") {
    // Project sheet ("材料购买1"): row after the last one with a description or shop (data from row 4;
    // numbers 027, 028… are pre-filled in C; the SUM row has an amount but no description, so ignore amounts).
    let last = 3;
    text.forEach((row, i) => { if (i >= 3 && [row[5], row[6]].some((c) => String(c ?? "").trim())) last = i + 1; });
    return last + 1;
  }
  if (kind === "bank") {
    let last = 14;
    text.forEach((row, i) => { if (i >= 14 && row.slice(1, 16).some((c) => String(c).trim())) last = i + 1; });
    return last + 1;
  }
  for (let r = 17; r <= 51; r++) {
    const row = text[r - 1] || [];
    if (![row[1], row[2], row[6]].some((c) => String(c ?? "").trim())) return r;
  }
  return null;
}

// Project number for a row: keep a pre-filled one, else one more than the highest so far ("027").
export function projectNumber(text, row) {
  const existing = String(text[row - 1]?.[2] ?? "").trim();
  if (existing) return null;
  const max = Math.max(0, ...text.slice(3, row - 1).map((r) => parseInt(String(r?.[2] ?? ""), 10)).filter((n) => !isNaN(n)));
  return pad3(max + 1);
}
const pad3 = (n) => String(n).padStart(3, "0");

export const REGISTER_HEADER = [
  "Date", "Company", "Merchant", "Purpose", "Property", "Total", "VAT", "Net", "VAT receipt",
  "Supplier VAT no", "Card", "Card last4", "Reclaim from company", "File", "Filed on", "Source name",
];

const q = (v) => `"${String(v ?? "").replace(/"/g, '""')}"`;

export function registerRow(r, card, filePath, sourceName) {
  const net = (Number(r.total) - Number(r.vat || 0)).toFixed(2);
  return [
    r.date, r.company, r.merchant, r.purpose, r.property, Number(r.total).toFixed(2),
    Number(r.vat || 0).toFixed(2), net, r.isVat ? "Yes" : "No", r.vatNumber, card, r.cardLast4,
    card === "OTHER" ? "Yes" : "No", filePath, new Date().toISOString().slice(0, 10), sourceName,
  ].map(q).join(",");
}

// Minimal CSV parse — enough to spot duplicates (date + total).
// ----- Duplicates: same date + same amount + same merchant = the same receipt, and the new one
// replaces the old one (file, register line or sheet row). Merchant is matched loosely
// ("Selco" ~ "Selco Builders Warehouse"); a blank merchant on either side doesn't block a match.
export function sameMerchant(a, b) {
  const x = compact(a), y = compact(b);
  if (!x || !y) return true;
  return x.includes(y) || y.includes(x) || x.slice(0, 5) === y.slice(0, 5);
}

const csvCells = (line) => line.match(/"((?:[^"]|"")*)"/g)?.map((c) => c.slice(1, -1).replace(/""/g, '"')) || [];

// Two copies of one purchase: (A) same date, or (B) the old one is an order confirmation —
// purpose ends "(order)" — up to 60 days apart (the formal invoice comes after delivery).
// Amount and merchant must match in both cases.
export const ORDER_TAG = "(order)";
const ORDER_WINDOW_DAYS = 60;
const isOrder = (text) => String(text || "").toLowerCase().includes(ORDER_TAG);
const daysApart = (isoA, isoB) => Math.abs(excelSerial(isoA) - excelSerial(isoB));
export function samePurchase(old, r) {
  if (Number(old.total).toFixed(2) !== Number(r.total).toFixed(2) || !sameMerchant(old.merchant, r.merchant)) return false;
  if (old.date === r.date) return true;
  return isOrder(old.purpose) && parseISO(old.date) && daysApart(old.date, r.date) <= ORDER_WINDOW_DAYS;
}

// Index (in csvText.split(/\r?\n/)) of the register line for the same purchase, or -1.
export function findDuplicateLine(csvText, r) {
  return csvText.split(/\r?\n/).findIndex((line, i) => {
    if (i === 0) return false;
    const c = csvCells(line);
    return c[1] === r.company && samePurchase({ date: c[0], merchant: c[2], purpose: c[3], total: c[5] }, r);
  });
}

export function findDuplicate(csvText, r) {
  const i = findDuplicateLine(csvText, r);
  return i < 0 ? null : csvText.split(/\r?\n/)[i];
}

// File path (relative to the company root) recorded on a register line.
export const registerLineFile = (line) => csvCells(line)[13] || null;

// Old file in a folder for the same receipt, by its name: "[NNN ]YYYY MM DD CO, Merchant, …, GBP 12.34[, VAT …].ext".
export function findOldFile(names, r) {
  const amount = `GBP ${Number(r.total).toFixed(2)}`;
  return names.find((n) => {
    const body = n.replace(/^\d{3}(?:[ -]\d{2})? (?=\d{4} \d{2} \d{2} )/, ""); // drop a "010 " number prefix
    const m = /^(\d{4}) (\d{2}) (\d{2}) [^,]*, ([^,]*),(.*)$/.exec(body);
    if (!m || !(n.includes(`${amount}.`) || n.includes(`${amount},`))) return false;
    return samePurchase({ date: `${m[1]}-${m[2]}-${m[3]}`, merchant: m[4], purpose: m[5], total: r.total }, r);
  }) || null;
}
