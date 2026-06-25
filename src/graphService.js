// src/graphService.js — CSV file version
// Reads/writes "Portfolio Strategy Data.csv" in SharePoint Documents
// Avoids SharePoint list write permission issues

const GRAPH = "https://graph.microsoft.com/v1.0";
const SP_HOST = "adobe.sharepoint.com";
const SP_PATH = "/sites/Monitor_SP";
const CSV_FILENAME = "Portfolio Strategy Data.csv";

let _siteId = null;

async function gFetch(url, token, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(opts.headers || {}),
    },
  });
  if (res.status === 204) return null;
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body?.error?.message || `HTTP ${res.status}`);
  }
  return res;
}

async function getSiteId(token) {
  if (_siteId) return _siteId;
  const res = await gFetch(`${GRAPH}/sites/${SP_HOST}:${SP_PATH}`, token);
  const d = await res.json();
  _siteId = d.id;
  return _siteId;
}

function parseCSVRow(line) {
  const result = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const c = line[i];
    if (c === '"') {
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (c === "," && !inQuotes) {
      result.push(current.trim());
      current = "";
    } else {
      current += c;
    }
  }
  result.push(current.trim());
  return result;
}

function escapeCSV(val) {
  if (val === null || val === undefined) return "";
  const s = String(val);
  if (s.includes(",") || s.includes('"') || s.includes("\n")) {
    return `"${s.replace(/"/g, '""')}"`;
  }
  return s;
}

function findHeader(headers, ...candidates) {
  for (const c of candidates) {
    const m = headers.find(
      (h) => h.toLowerCase() === c.toLowerCase() ||
             h.toLowerCase().replace(/\s/g, "") === c.toLowerCase().replace(/\s/g, "")
    );
    if (m) return m;
  }
  return candidates[0];
}

function toDateStr(v) {
  if (!v) return null;
  if (/^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
  if (/^\d{1,2}\/\d{1,2}\/\d{4}/.test(v)) {
    const [m, d, y] = v.split("/");
    return `${y}-${m.padStart(2, "0")}-${d.padStart(2, "0")}`;
  }
  return v || null;
}

function decodeRow(row, headers, index) {
  const H = {
    id:   findHeader(headers, "id", "ID"),
    site: findHeader(headers, "title", "site", "Site"),
    cat:  findHeader(headers, "category", "Category"),
    led:  findHeader(headers, "led", "LED"),
    brk:  findHeader(headers, "break", "Break"),
    cap:  findHeader(headers, "capital project", "CapitalProject"),
    own:  findHeader(headers, "owned", "Owned"),
    note: findHeader(headers, "notes", "Notes"),
  };
  const owned = (row[H.own] || "").toLowerCase();
  return {
    id: row[H.id] || String(index),
    Site: row[H.site] || "",
    Category: row[H.cat] || "",
    LED: toDateStr(row[H.led]),
    Break: toDateStr(row[H.brk]),
    CapitalProject: row[H.cap] || null,
    Owned: owned === "yes" || owned === "true" || owned === "1",
    Notes: row[H.note] || "",
  };
}

function applyForm(row, headers, form) {
  const H = {
    site: findHeader(headers, "title", "site", "Site"),
    cat:  findHeader(headers, "category", "Category"),
    led:  findHeader(headers, "led", "LED"),
    brk:  findHeader(headers, "break", "Break"),
    cap:  findHeader(headers, "capital project", "CapitalProject"),
    own:  findHeader(headers, "owned", "Owned"),
    note: findHeader(headers, "notes", "Notes"),
  };
  return {
    ...row,
    [H.site]: form.Site || "",
    [H.cat]:  form.Category || "",
    [H.led]:  form.LED || "",
    [H.brk]:  form.Break || "",
    [H.cap]:  form.CapitalProject || "",
    [H.own]:  form.Owned ? "Yes" : "No",
    [H.note]: form.Notes || "",
  };
}

async function readCSVText(token) {
  const siteId = await getSiteId(token);
  const res = await gFetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${encodeURIComponent(CSV_FILENAME)}:/content`,
    token
  );
  return res.text();
}

async function writeCSVText(token, text) {
  const siteId = await getSiteId(token);
  await gFetch(
    `${GRAPH}/sites/${siteId}/drive/root:/${encodeURIComponent(CSV_FILENAME)}:/content`,
    token,
    {
      method: "PUT",
      headers: { "Content-Type": "text/csv; charset=utf-8" },
      body: text,
    }
  );
}

function parseCSV(text) {
  const lines = text.replace(/\r/g, "").split("\n").filter((l) => l.trim());
  if (!lines.length) return { headers: [], rows: [], items: [] };
  const headers = parseCSVRow(lines[0]);
  const rows = [];
  const items = [];
  for (let i = 1; i < lines.length; i++) {
    const values = parseCSVRow(lines[i]);
    const row = {};
    headers.forEach((h, idx) => { row[h] = values[idx] ?? ""; });
    rows.push(row);
    items.push(decodeRow(row, headers, i));
  }
  return { headers, rows, items };
}

function serializeCSV(headers, rows) {
  const lines = [headers.join(",")];
  rows.forEach((row) => {
    lines.push(headers.map((h) => escapeCSV(row[h])).join(","));
  });
  return lines.join("\n");
}

export async function fetchAllItems(token) {
  const text = await readCSVText(token);
  const { items } = parseCSV(text);
  return items;
}

export async function updateItem(token, itemId, form) {
  const text = await readCSVText(token);
  const { headers, rows, items } = parseCSV(text);
  const idx = items.findIndex((it) => String(it.id) === String(itemId));
  if (idx === -1) throw new Error(`Item ${itemId} not found in CSV`);
  rows[idx] = applyForm(rows[idx], headers, form);
  await writeCSVText(token, serializeCSV(headers, rows));
}

export async function createItem(token, form) {
  const text = await readCSVText(token);
  const { headers, rows, items } = parseCSV(text);
  const idHeader = findHeader(headers, "id", "ID");
  const maxId = items.reduce((m, it) => Math.max(m, parseInt(it.id) || 0), 0);
  const newId = String(maxId + 1);
  const newRow = {};
  headers.forEach((h) => { newRow[h] = ""; });
  const filled = applyForm(newRow, headers, form);
  filled[idHeader] = newId;
  rows.push(filled);
  await writeCSVText(token, serializeCSV(headers, rows));
  return newId;
}

export async function deleteItem(token, itemId) {
  const text = await readCSVText(token);
  const { headers, rows, items } = parseCSV(text);
  const idx = items.findIndex((it) => String(it.id) === String(itemId));
  if (idx === -1) throw new Error(`Item ${itemId} not found`);
  rows.splice(idx, 1);
  await writeCSVText(token, serializeCSV(headers, rows));
}
