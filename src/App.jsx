// src/App.jsx
// Portfolio Strategy Overview
// Auth: MSAL (Microsoft login) → Graph API → SharePoint list (real-time read/write)

import { useState, useEffect, useMemo, useCallback } from "react";
import {
  useMsal,
  useIsAuthenticated,
  AuthenticatedTemplate,
  UnauthenticatedTemplate,
} from "@azure/msal-react";
import { InteractionRequiredAuthError } from "@azure/msal-browser";
import { graphScopes, EDITOR_EMAILS } from "./authConfig";
import { fetchAllItems, updateItem, createItem, deleteItem } from "./graphService";

// ── Color helpers ──────────────────────────────────────────────────────────────
const YEAR_COLORS = { 2026: "#a8d08d", 2027: "#548235", 2028: "#4472c4", 2029: "#203864" };
const yearColor = (d) => {
  if (!d) return "#aaa";
  const y = new Date(d).getFullYear();
  return YEAR_COLORS[y] || (y <= 2026 ? "#a8d08d" : "#7030a0");
};
const cardBg = (item) => item.Owned ? "#c00000" : yearColor(item.Break || item.LED);
const isLight = (hex) => {
  if (!hex || hex[0] !== "#") return true;
  const r = parseInt(hex.slice(1, 3), 16) * 299;
  const g = parseInt(hex.slice(3, 5), 16) * 587;
  const b = parseInt(hex.slice(5, 7), 16) * 114;
  return (r + g + b) / 1000 > 140;
};
const fmtDate = (d) => {
  if (!d) return "";
  const s = typeof d === "string" ? d.slice(0, 10) : "";
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
    const [y, m, day] = s.split("-").map(Number);
    return `${m}/${day}/${y}`;
  }
  const dt = new Date(d);
  return isNaN(dt) ? String(d) : `${dt.getMonth() + 1}/${dt.getDate()}/${dt.getFullYear()}`;
};
const toDateInput = (d) => {
  if (!d) return "";
  const s = typeof d === "string" ? d : "";
  return s.slice(0, 10); // "YYYY-MM-DD"
};

// ── Category → column key ──────────────────────────────────────────────────────
const groupKey = (cat = "") => {
  const c = (cat || "").toLowerCase();
  if (c.includes("office closur")) return "closure";
  if (c.includes("optim")) return "optim";
  if (c.includes("potential clos")) return "potclos";
  if (c.includes("active") || c.includes("actively")) return "active";
  if (c.includes("reloc") || c.includes("expan")) return "reloc";
  if (c.includes("potential service") || c.includes("potential serviced")) return "potservice";
  if (c.includes("service") || c.includes("serviced")) return "service";
  return "active";
};

const SECTION_TO_CAT = {
  "Office Closures": "Office Closures",
  "Optimization Opps": "Optimization Opps",
  "Potential Closures": "Potential Closures",
  "Active Monitoring": "Actively Monitoring",
  "Relocation / Expansion": "Relocation/Expansion",
  "Service Office": "Serviced Office",
  "Potential Service Office": "Potential Service Office",
};

const ALL_CATS = [
  "Actively Monitoring", "Office Closures", "Optimization Opps",
  "Potential Closures", "Relocation/Expansion", "Serviced Office",
  "Potential Service Office",
];

const CARD_W = 200; // fixed card column width, shared across all sections

const sortByNextBreak = (arr) => [...arr].sort((a, b) => {
  const aOwned = !!a.Owned, bOwned = !!b.Owned;
  if (aOwned !== bOwned) return aOwned ? 1 : -1;
  const dateA = a.Break || a.LED;
  const dateB = b.Break || b.LED;
  const da = dateA ? new Date(dateA).getTime() : Infinity;
  const db = dateB ? new Date(dateB).getTime() : Infinity;
  return da - db;
});

// ── Shared style tokens ────────────────────────────────────────────────────────
const HDR = {
  background: "#1a1a2e", color: "#fff", textAlign: "center",
  padding: "7px 6px", fontWeight: 700, fontSize: "14px",
  marginBottom: "2px", letterSpacing: "0.02em",
};
const ADD_BTN = {
  display: "block", width: "100%", marginTop: "3px", padding: "2px 4px",
  fontSize: "9.5px", background: "transparent", border: "1px dashed #bbb",
  cursor: "pointer", color: "#888", borderRadius: "2px",
};

// ── Card ───────────────────────────────────────────────────────────────────────
function Card({ item, editMode, onClick }) {
  const [hovered, setHovered] = useState(false);
  const bg = cardBg(item);
  const col = isLight(bg) ? "#111" : "#fff";
  return (
    <div
      onClick={() => editMode && onClick(item)}
      onMouseEnter={() => setHovered(true)}
      onMouseLeave={() => setHovered(false)}
      style={{
        background: bg, color: col, padding: "3px 6px", marginBottom: "2px",
        fontSize: "14px", lineHeight: 1.35,
        cursor: editMode ? "pointer" : "default",
        opacity: editMode && hovered ? 0.82 : 1,
        outline: editMode ? "1px dashed rgba(0,0,0,0.18)" : "none",
        transition: "opacity 0.1s", position: "relative",
        height: "49px", overflow: "hidden", boxSizing: "border-box",
      }}
    >
      <div style={{ fontWeight: 700, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", paddingRight: editMode ? 14 : 0 }}>
        {item.Site}
      </div>
      <div style={{ fontSize: "12px", opacity: 0.88, marginTop: "3px" }}>
        {item.Break ? `Break:${fmtDate(item.Break)} ` : ""}
        {item.LED ? `LED:${fmtDate(item.LED)}` : ""}
      </div>
      {editMode && hovered && (
        <span style={{ position: "absolute", top: 3, right: 4, fontSize: "9px", opacity: 0.7 }}>✎</span>
      )}
      {item.CapitalProject === "Active" && (
        <div style={{ position: "absolute", top: 0, right: 0, width: 0, height: 0, borderStyle: "solid", borderWidth: "0 18px 18px 0", borderColor: "transparent #ffc000 transparent transparent" }} />
      )}
      {item.CapitalProject === "Potential/Future" && (
        <div style={{ position: "absolute", top: 0, right: 0, width: 0, height: 0, borderStyle: "solid", borderWidth: "0 18px 18px 0", borderColor: "transparent #ff92d0 transparent transparent" }} />
      )}
    </div>
  );
}

// ── Multi-column section ───────────────────────────────────────────────────────
function ColSection({ label, items, cols, editMode, onEdit, onAdd }) {
  const chunks = useMemo(() => {
    const sz = Math.max(1, Math.ceil(items.length / cols));
    return Array.from({ length: cols }, (_, i) => items.slice(i * sz, (i + 1) * sz));
  }, [items, cols]);

  return (
    <div style={{ flexShrink: 0, width: cols * CARD_W + (cols - 1) * 2 }}>
      <div style={HDR}>{label}</div>
      <div style={{ display: "flex", gap: "2px" }}>
        {chunks.map((chunk, ci) => (
          <div key={ci} style={{ width: CARD_W, flexShrink: 0 }}>
            {chunk.map((i) => <Card key={i.id} item={i} editMode={editMode} onClick={onEdit} />)}
          </div>
        ))}
      </div>
      {editMode && <button onClick={() => onAdd(label)} style={ADD_BTN}>+ Add to {label}</button>}
    </div>
  );
}

// ── Sidebar section ────────────────────────────────────────────────────────────
function SideSection({ label, items, editMode, onEdit, onAdd }) {
  return (
    <div style={{ marginBottom: "8px" }}>
      <div style={HDR}>{label}</div>
      {items.map((i) => <Card key={i.id} item={i} editMode={editMode} onClick={onEdit} />)}
      {editMode && <button onClick={() => onAdd(label)} style={ADD_BTN}>+ Add</button>}
    </div>
  );
}

// ── Edit / Add Modal ───────────────────────────────────────────────────────────
function EditModal({ item, defaultCat, saving, onSave, onClose, onDelete }) {
  const [f, setF] = useState({
    Site: item?.Site || "",
    Category: item?.Category || defaultCat || "",
    LED: toDateInput(item?.LED),
    Break: toDateInput(item?.Break),
    CapitalProject: item?.CapitalProject || "",
    Owned: item?.Owned || false,
    Notes: item?.Notes || "",
  });
  const upd = (k) => (e) =>
    setF((v) => ({ ...v, [k]: e.target.type === "checkbox" ? e.target.checked : e.target.value }));
  const preview = { ...f, id: item?.id, LED: f.LED || null, Break: f.Break || null };

  return (
    <div style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.52)", zIndex: 1000, display: "flex", alignItems: "center", justifyContent: "center" }}>
      <div style={{ background: "#fff", borderRadius: "8px", padding: "20px 22px", width: "360px", maxHeight: "90vh", overflowY: "auto", boxShadow: "0 16px 56px rgba(0,0,0,0.32)" }}>
        <h3 style={{ margin: "0 0 14px", color: "#c00000", fontSize: "14px", fontWeight: 700 }}>
          {item?.id ? `Edit: ${item.Site}` : "Add New Site"}
        </h3>

        {[["Site Name","Site","text"],["LED Date","LED","date"],["Break Date","Break","date"],["Notes","Notes","text"]].map(([lbl, k, t]) => (
          <div key={k} style={{ marginBottom: "9px" }}>
            <label style={{ display: "block", fontSize: "10px", fontWeight: 700, color: "#555", marginBottom: "2px" }}>{lbl}</label>
            <input type={t} value={f[k]} onChange={upd(k)}
              style={{ width: "100%", padding: "5px 7px", border: "1px solid #ccc", borderRadius: "3px", fontSize: "11px", boxSizing: "border-box" }}
            />
          </div>
        ))}

        <div style={{ marginBottom: "9px" }}>
          <label style={{ display: "block", fontSize: "10px", fontWeight: 700, color: "#555", marginBottom: "2px" }}>Category</label>
          <select value={f.Category} onChange={upd("Category")}
            style={{ width: "100%", padding: "5px 7px", border: "1px solid #ccc", borderRadius: "3px", fontSize: "11px" }}>
            <option value="">Select…</option>
            {ALL_CATS.map((c) => <option key={c} value={c}>{c}</option>)}
          </select>
        </div>

        <div style={{ marginBottom: "9px" }}>
          <label style={{ display: "block", fontSize: "10px", fontWeight: 700, color: "#555", marginBottom: "2px" }}>Capital Project</label>
          <select value={f.CapitalProject} onChange={upd("CapitalProject")}
            style={{ width: "100%", padding: "5px 7px", border: "1px solid #ccc", borderRadius: "3px", fontSize: "11px" }}>
            <option value="">None</option>
            <option value="Active">Active</option>
            <option value="Potential/Future">Potential/Future</option>
          </select>
        </div>

        <label style={{ display: "flex", alignItems: "center", gap: "7px", fontSize: "11px", cursor: "pointer", marginBottom: "14px" }}>
          <input type="checkbox" checked={f.Owned} onChange={upd("Owned")} />
          Owned Property
        </label>

        <div style={{ marginBottom: "14px" }}>
          <div style={{ fontSize: "10px", fontWeight: 700, color: "#555", marginBottom: "4px" }}>Card Preview</div>
          <Card item={preview} editMode={false} onClick={() => {}} />
        </div>

        <div style={{ display: "flex", gap: "6px", justifyContent: "flex-end" }}>
          {item?.id && (
            <button onClick={() => onDelete(item.id)} disabled={saving}
              style={{ padding: "6px 12px", background: "#dc3545", color: "#fff", border: "none", borderRadius: "4px", fontSize: "11px", cursor: "pointer" }}>
              Delete
            </button>
          )}
          <button onClick={onClose} disabled={saving}
            style={{ padding: "6px 12px", border: "1px solid #ccc", borderRadius: "4px", fontSize: "11px", cursor: "pointer", background: "#fff" }}>
            Cancel
          </button>
          <button onClick={() => onSave(f)} disabled={saving}
            style={{ padding: "6px 12px", background: "#c00000", color: "#fff", border: "none", borderRadius: "4px", fontSize: "11px", cursor: "pointer", fontWeight: 700 }}>
            {saving ? "Saving…" : "Save to SharePoint"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ── Legend ─────────────────────────────────────────────────────────────────────
function Legend() {
  const solidSwatch = (c) => (
    <span style={{ display: "inline-block", width: 17, height: 17, background: c, marginRight: 6, verticalAlign: "middle", border: "1px solid rgba(0,0,0,0.12)", flexShrink: 0 }} />
  );
  const triangleSwatch = (c) => (
    <span style={{ display: "inline-block", width: 0, height: 0, borderStyle: "solid", borderWidth: "0 17px 17px 0", borderColor: `transparent ${c} transparent transparent`, marginRight: 6, verticalAlign: "middle", flexShrink: 0 }} />
  );
  const YEARS = [["2026","#a8d08d"],["2027","#548235"],["2028","#4472c4"],["2029","#203864"],["2030+","#7030a0"]];
  const rowStyle = { display: "flex", alignItems: "center", marginBottom: "4px", fontSize: "14px" };
  const labelStyle = { fontWeight: 700, fontSize: "16px", marginBottom: "6px" };

  return (
    <div style={{ display: "flex", justifyContent: "flex-end", marginTop: "10px" }}>
      <div style={{ border: "1px solid #ccc", padding: "10px 17px", display: "inline-grid", gridTemplateColumns: "auto auto", gap: "0 24px" }}>
        <div>
          <div style={labelStyle}>Next Break:</div>
          {YEARS.map(([y, c]) => (
            <div key={y} style={rowStyle}>{solidSwatch(c)}{y}</div>
          ))}
        </div>
        <div>
          <div style={labelStyle}>Owned:</div>
          <div style={{ ...rowStyle, marginBottom: "12px" }}>{solidSwatch("#c00000")}Owned</div>
          <div style={labelStyle}>Capital Project:</div>
          <div style={rowStyle}>{triangleSwatch("#ffc000")}Active</div>
          <div style={rowStyle}>{triangleSwatch("#ff92d0")}Potential/Future</div>
        </div>
      </div>
    </div>
  );
}

// ── Sign-in screen ─────────────────────────────────────────────────────────────
function SignInScreen({ onSignIn }) {
  return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", background: "#f5f5f5" }}>
      <div style={{ background: "#fff", borderRadius: "12px", padding: "40px 48px", boxShadow: "0 8px 32px rgba(0,0,0,0.12)", textAlign: "center", maxWidth: 400 }}>
        <img src="/portfolio-strategy-status/Adobe_icon_RGB_red.png" alt="Adobe" style={{ height: "32px", marginBottom: "8px" }} />
        <h1 style={{ color: "#c00000", fontSize: "20px", margin: "0 0 6px" }}>Portfolio Strategy Overview</h1>
        <p style={{ color: "#666", fontSize: "12px", marginBottom: "28px" }}>Sign in with your Adobe Microsoft account to access the dashboard.</p>
        <button onClick={onSignIn}
          style={{ padding: "12px 28px", background: "#0078d4", color: "#fff", border: "none", borderRadius: "6px", fontSize: "14px", fontWeight: 700, cursor: "pointer", display: "flex", alignItems: "center", gap: "10px", margin: "0 auto" }}>
          <svg width="18" height="18" viewBox="0 0 21 21" fill="none"><rect width="10" height="10" fill="#f25022"/><rect x="11" width="10" height="10" fill="#7fba00"/><rect y="11" width="10" height="10" fill="#00a4ef"/><rect x="11" y="11" width="10" height="10" fill="#ffb900"/></svg>
          Sign in with Microsoft
        </button>
      </div>
    </div>
  );
}

// ── Main report (authenticated) ────────────────────────────────────────────────
function PortfolioReport() {
  const { instance, accounts } = useMsal();
  const account = accounts[0];

  const [items, setItems] = useState([]);
  const [loading, setLoading] = useState(true);
  const [fetchErr, setFetchErr] = useState(null);
  const [editMode, setEditMode] = useState(false);
  const [editing, setEditing] = useState(null);
  const [defCat, setDefCat] = useState("");
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState({ msg: "", ok: true });
  const [zoom, setZoom] = useState(1);
  const zoomIn  = () => setZoom(z => Math.min(2.0, +(z + 0.25).toFixed(2)));
  const zoomOut = () => setZoom(z => Math.max(0.75, +(z - 0.25).toFixed(2)));
  const zoomReset = () => setZoom(1);

  const userEmail = account?.username?.toLowerCase() || "";
  const canEdit = EDITOR_EMAILS.length === 0 || EDITOR_EMAILS.includes(userEmail);

  const flash = (msg, ok = true) => {
    setStatus({ msg, ok });
    setTimeout(() => setStatus({ msg: "", ok: true }), 4500);
  };

  // ── Get access token (silent → popup fallback) ─────────────────────────────
  const getToken = useCallback(async () => {
    try {
      const result = await instance.acquireTokenSilent({ ...graphScopes, account });
      return result.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        const result = await instance.acquireTokenPopup({ ...graphScopes, account });
        return result.accessToken;
      }
      throw e;
    }
  }, [instance, account]);

  // ── Fetch all sites from SharePoint ────────────────────────────────────────
  const loadData = useCallback(async () => {
    setLoading(true);
    setFetchErr(null);
    try {
      const token = await getToken();
      const data = await fetchAllItems(token);
      setItems(data);
      flash(`✅ Loaded ${data.length} sites`);
    } catch (e) {
      setFetchErr(e.message);
    } finally {
      setLoading(false);
    }
  }, [getToken]);

  useEffect(() => { loadData(); }, [loadData]);

  // ── Save (create or update) ────────────────────────────────────────────────
  const handleSave = useCallback(async (form) => {
    setSaving(true);
    try {
      const token = await getToken();
      const isNew = !editing?.id;

      if (isNew) {
        const newId = await createItem(token, form);
        setItems((prev) => [...prev, { ...form, id: newId }]);
        flash("✅ Created in SharePoint");
      } else {
        await updateItem(token, editing.id, form);
        setItems((prev) => prev.map((i) => (i.id === editing.id ? { ...i, ...form } : i)));
        flash("✅ Updated in SharePoint");
      }
      setEditing(null);
    } catch (e) {
      flash(`❌ ${e.message}`, false);
    } finally {
      setSaving(false);
    }
  }, [getToken, editing]);

  // ── Delete ─────────────────────────────────────────────────────────────────
  const handleDelete = useCallback(async (id) => {
    if (!window.confirm("Delete this site from SharePoint? This cannot be undone.")) return;
    setSaving(true);
    try {
      const token = await getToken();
      await deleteItem(token, id);
      setItems((prev) => prev.filter((i) => i.id !== id));
      flash("✅ Deleted from SharePoint");
      setEditing(null);
    } catch (e) {
      flash(`❌ ${e.message}`, false);
    } finally {
      setSaving(false);
    }
  }, [getToken]);

  // ── Group items ────────────────────────────────────────────────────────────
  const g = useMemo(() => {
    const groups = { closure:[], optim:[], potclos:[], active:[], reloc:[], service:[], potservice:[] };
    items.forEach((i) => { const k = groupKey(i.Category); if (groups[k]) groups[k].push(i); });
    Object.keys(groups).forEach((k) => { groups[k] = sortByNextBreak(groups[k]); });
    return groups;
  }, [items]);

  const openEdit = useCallback((item) => { setDefCat(item.Category); setEditing(item); }, []);
  const openAdd = useCallback((label) => { setDefCat(SECTION_TO_CAT[label] || ""); setEditing({}); }, []);

  if (loading) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh" }}>
      <style>{`@keyframes spin{from{transform:rotate(0deg)}to{transform:rotate(360deg)}}`}</style>
      <div style={{ fontSize: "36px", color: "#c00000", animation: "spin 1s linear infinite", marginBottom: "12px" }}>⟳</div>
      <div style={{ fontWeight: 700, color: "#c00000" }}>Loading portfolio data…</div>
      <div style={{ color: "#888", fontSize: "11px", marginTop: "6px" }}>Connecting to SharePoint via Microsoft Graph</div>
    </div>
  );

  if (fetchErr) return (
    <div style={{ padding: "28px", maxWidth: 600 }}>
      <h2 style={{ color: "#c00000", fontSize: "16px", margin: "0 0 10px" }}>⚠ SharePoint Error</h2>
      <p style={{ color: "#444", fontSize: "12px", lineHeight: 1.6, whiteSpace: "pre-wrap", marginBottom: 12 }}>{fetchErr}</p>
      <p style={{ color: "#888", fontSize: "11px", marginBottom: 12 }}>
        This usually means the Azure AD app doesn't have <strong>Sites.ReadWrite.All</strong> permission, or admin consent hasn't been granted yet. Check the SETUP.md guide.
      </p>
      <button onClick={loadData} style={{ padding: "8px 20px", background: "#c00000", color: "#fff", border: "none", borderRadius: "4px", cursor: "pointer", fontSize: "12px", fontWeight: 700 }}>
        Retry
      </button>
    </div>
  );

  return (
    <div style={{ overflowX: "auto", minHeight: "100vh" }}>
      <div style={{ minWidth: "960px", maxWidth: "1800px", margin: "0 auto", padding: "8px 10px 20px", zoom: zoom }}>

        {/* Header */}
        <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", marginBottom: "16px" }}>
          <div>
            <h1 style={{ margin: 0, fontSize: "30px", color: "#c00000", display: "flex", alignItems: "center", gap: "7px" }}>
              <img src="/portfolio-strategy-status/Adobe_icon_RGB_red.png" alt="Adobe" style={{ height: "24px", verticalAlign: "middle" }} /> Portfolio Strategy Overview
            </h1>
            <div style={{ fontSize: "15px", fontWeight: 700, color: "#666", marginTop: "1px" }}>
              By Type and Year&nbsp;|&nbsp;{items.length} sites&nbsp;|&nbsp;Signed in as {account?.name || userEmail}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: "8px", flexShrink: 0, paddingTop: 2 }}>
            {status.msg && <span style={{ fontSize: "10.5px", fontWeight: 500, color: status.ok ? "#28a745" : "#c00000" }}>{status.msg}</span>}
            {canEdit && (
              <button onClick={() => setEditMode((m) => !m)}
                style={{ padding: "5px 12px", border: "none", borderRadius: "4px", cursor: "pointer", fontSize: "11px", fontWeight: 700, background: editMode ? "#c00000" : "#444", color: "#fff" }}>
                {editMode ? "✎ Edit Mode ON" : "✎ Edit Mode"}
              </button>
            )}
            <button onClick={loadData}
              style={{ padding: "5px 14px", background: "#7030a0", color: "#fff", border: "none", borderRadius: "4px", cursor: "pointer", fontSize: "11px", fontWeight: 700 }}>
              ▶ Refresh
            </button>
            <div style={{ display: "flex", alignItems: "center", gap: "4px" }}>
              <button onClick={zoomOut} disabled={zoom <= 0.75} style={{ background: "none", border: "none", cursor: "pointer", fontSize: "14px", color: "#555", lineHeight: 1, padding: "0 2px" }}>−</button>
              <input type="range" min={0.75} max={2.0} step={0.25} value={zoom}
                onChange={e => setZoom(parseFloat(e.target.value))}
                style={{ width: "80px", cursor: "pointer", accentColor: "#555" }}
              />
              <button onClick={zoomIn} disabled={zoom >= 2.0} style={{ background: "none", border: "none", cursor: "pointer", fontSize: "14px", color: "#555", lineHeight: 1, padding: "0 2px" }}>+</button>
              <span onDoubleClick={zoomReset} title="Double-click to reset" style={{ fontSize: "10px", color: "#555", minWidth: "30px", cursor: "default", userSelect: "none" }}>{Math.round(zoom * 100)}%</span>
            </div>
            <button onClick={() => instance.logoutPopup()}
              style={{ padding: "5px 10px", background: "transparent", color: "#888", border: "1px solid #ddd", borderRadius: "4px", cursor: "pointer", fontSize: "10px" }}>
              Sign out
            </button>
          </div>
        </div>

        {editMode && (
          <div style={{ background: "#fff3cd", border: "1px solid #ffc107", padding: "4px 10px", marginBottom: "6px", borderRadius: "3px", fontSize: "10px", color: "#664d03" }}>
            ✎ <strong>Edit Mode</strong> — Click any card to edit it, or use "+ Add" to create new entries. All changes write to SharePoint immediately.
          </div>
        )}

        {/* Main grid */}
        <div style={{ display: "flex", gap: "10px", alignItems: "flex-start" }}>
          <div style={{ width: CARD_W, flexShrink: 0 }}>
            <SideSection label="Office Closures" items={g.closure} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
            <SideSection label="Optimization Opps" items={g.optim} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
            <SideSection label="Potential Closures" items={g.potclos} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
          </div>
          <ColSection label="Active Monitoring" cols={3} items={g.active} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
          <ColSection label="Relocation / Expansion" cols={1} items={g.reloc} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
          <ColSection label="Service Office" cols={2} items={g.service} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
          <ColSection label="Potential Service Office" cols={1} items={g.potservice} editMode={editMode} onEdit={openEdit} onAdd={openAdd} />
        </div>

        <Legend />
      </div>

      {editing !== null && (
        <EditModal
          item={editing?.id ? editing : null}
          defaultCat={defCat}
          saving={saving}
          onSave={handleSave}
          onClose={() => setEditing(null)}
          onDelete={handleDelete}
        />
      )}
    </div>
  );
}

// ── Root — handles auth gate ───────────────────────────────────────────────────
export default function App() {
  const { instance } = useMsal();
  const isAuthenticated = useIsAuthenticated();

  const handleSignIn = async () => {
    try {
      await instance.loginPopup(graphScopes);
    } catch (e) {
      console.error("Login failed:", e);
    }
  };

  return (
    <>
      <AuthenticatedTemplate>
        <PortfolioReport />
      </AuthenticatedTemplate>
      <UnauthenticatedTemplate>
        <SignInScreen onSignIn={handleSignIn} />
      </UnauthenticatedTemplate>
    </>
  );
}
