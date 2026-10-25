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
  const amount = `${r.currency || "GBP"} ${Number(r.total).toFixed(2)}`;
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
export function findDuplicate(csvText, r) {
  const total = Number(r.total).toFixed(2);
  return csvText.split(/\r?\n/).slice(1).find((line) => {
    const cells = line.match(/"((?:[^"]|"")*)"/g)?.map((c) => c.slice(1, -1)) || [];
    return cells[0] === r.date && cells[5] === total && cells[1] === r.company;
  }) || null;
}
