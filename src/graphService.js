// src/graphService.js
const GRAPH = "https://graph.microsoft.com/v1.0";
const SP_HOST = "adobe.sharepoint.com";
const SP_PATH = "/sites/Monitor_SP";
const LIST_NAME = "portfolio strategy data temp";

let _siteId = null;
let _listId = null;
let _fieldMap = null;

async function gFetch(url, token, opts = {}) {
  const res = await fetch(url, {
    ...opts,
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
      ...(opts.headers || {}),
    },
  });
  if (res.status === 204) return null;
  const body = await res.json().catch(() => ({}));
  if (!res.ok) {
    const msg = body?.error?.message || body?.error?.code || `HTTP ${res.status}`;
    throw new Error(`Graph API error: ${msg}`);
  }
  return body;
}

async function getSiteId(token) {
  if (_siteId) return _siteId;
  const d = await gFetch(`${GRAPH}/sites/${SP_HOST}:${SP_PATH}`, token);
  _siteId = d.id;
  return _siteId;
}

// Expose for diagnostics
export async function getAllLists(token) {
  const siteId = await getSiteId(token);
  const d = await gFetch(`${GRAPH}/sites/${siteId}/lists`, token);
  return d.value || [];
}

async function getListId(token) {
  if (_listId) return _listId;
  const siteId = await getSiteId(token);

  // Try fetching all lists and find by name (case-insensitive)
  const allLists = await getAllLists(token);
  const match = allLists.find(
    (l) => l.displayName?.toLowerCase() === LIST_NAME.toLowerCase() ||
           l.name?.toLowerCase() === LIST_NAME.toLowerCase()
  );

  if (!match) {
    const names = allLists.map((l) => `"${l.displayName}"`).join(", ");
    throw new Error(
      `List "${LIST_NAME}" not found on this site.\n\nAvailable lists: ${names}`
    );
  }

  _listId = match.id;
  return _listId;
}

async function getFieldMap(token) {
  if (_fieldMap) return _fieldMap;
  const siteId = await getSiteId(token);
  const listId = await getListId(token);
  const d = await gFetch(`${GRAPH}/sites/${siteId}/lists/${listId}/columns`, token);
  const map = {};
  (d.value || []).forEach((col) => {
    if (col.displayName) map[col.displayName.toLowerCase()] = col.name;
  });
  _fieldMap = map;
  return _fieldMap;
}

function decodeItem(rawFields, itemId, fieldMap) {
  const f = rawFields || {};
  const get = (...names) => {
    for (const n of names) {
      const internalName = fieldMap[n.toLowerCase()] || n;
      if (f[internalName] !== undefined && f[internalName] !== null) return f[internalName];
      if (f[n] !== undefined && f[n] !== null) return f[n];
    }
    return null;
  };
  return {
    id: itemId,
    Site: get("title", "Title", "Site"),
    Category: get("category", "Category"),
    LED: get("led", "LED"),
    Break: get("break", "Break"),
    CapitalProject: get("capital project", "Capital Project", "Capital_x0020_Project", "CapitalProject"),
    Owned: Boolean(get("owned", "Owned")),
    Notes: get("notes", "Notes") || "",
  };
}

function encodeFields(form, fieldMap) {
  const toISO = (d) => (d ? `${d}T00:00:00Z` : null);
  const resolve = (...names) => {
    for (const n of names) {
      if (fieldMap[n.toLowerCase()]) return fieldMap[n.toLowerCase()];
    }
    return names[0];
  };
  return {
    [resolve("title", "Title")]: form.Site,
    [resolve("category", "Category")]: form.Category || null,
    [resolve("led", "LED")]: toISO(form.LED),
    [resolve("break", "Break")]: toISO(form.Break),
    [resolve("capital project", "Capital Project", "Capital_x0020_Project")]: form.CapitalProject || null,
    [resolve("owned", "Owned")]: form.Owned,
    [resolve("notes", "Notes")]: form.Notes || null,
  };
}

export async function fetchAllItems(token) {
  const siteId = await getSiteId(token);
  const listId = await getListId(token);
  const fieldMap = await getFieldMap(token);
  const items = [];
  let url = `${GRAPH}/sites/${siteId}/lists/${listId}/items?expand=fields&$top=999`;
  while (url) {
    const d = await gFetch(url, token);
    (d.value || []).forEach((raw) => {
      items.push(decodeItem(raw.fields, raw.id, fieldMap));
    });
    url = d["@odata.nextLink"] || null;
  }
  return items;
}

export async function updateItem(token, itemId, form) {
  const siteId = await getSiteId(token);
  const listId = await getListId(token);
  const fieldMap = await getFieldMap(token);
  await gFetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}/fields`, token,
    { method: "PATCH", body: JSON.stringify(encodeFields(form, fieldMap)) });
}

export async function createItem(token, form) {
  const siteId = await getSiteId(token);
  const listId = await getListId(token);
  const fieldMap = await getFieldMap(token);
  const d = await gFetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items`, token,
    { method: "POST", body: JSON.stringify({ fields: encodeFields(form, fieldMap) }) });
  return d.id;
}

export async function deleteItem(token, itemId) {
  const siteId = await getSiteId(token);
  const listId = await getListId(token);
  await gFetch(`${GRAPH}/sites/${siteId}/lists/${listId}/items/${itemId}`, token,
    { method: "DELETE" });
}

export async function getDebugFieldMap(token) {
  return getFieldMap(token);
}
