import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Target, PhoneCall, CheckSquare, Layers, Settings as SettingsIcon,
  Plus, X, Trash2, TrendingUp, AlertCircle, DollarSign, UploadCloud, Check,
  ArrowRight, RefreshCw, FileSpreadsheet, CheckCircle2, ArrowUpRight, ArrowDownRight, Minus, Building2
} from "lucide-react";

const STORAGE_KEY = "scc-data-v1";

const STAGES = ["Prospecting", "Pipeline", "Upside", "Strong Upside", "Commit", "Closed Won", "Closed Lost"];
const OPEN_STAGES = STAGES.filter(s => s !== "Closed Won" && s !== "Closed Lost");
const STAGE_COLOR = {
  "Prospecting": "#94A3B8",
  "Pipeline": "#3B82F6",
  "Upside": "#F59E0B",
  "Strong Upside": "#EA580C",
  "Commit": "#16A34A",
  "Closed Won": "#0D9488",
  "Closed Lost": "#DC2626",
};

const CATEGORIES = [
  { key: "productMargin", label: "Product margin", color: "#3B82F6" },
  { key: "servicesMargin", label: "Services margin", color: "#334155" },
  { key: "managedServices", label: "Managed services", color: "#14B8A6" },
];

const ACTIVITY_TYPES = ["Call", "Email", "Meeting", "Demo", "Other"];
const PRIORITIES = ["Low", "Moderate", "High", "Critical"];
const PRIORITY_RANK = { Critical: 0, High: 1, Moderate: 2, Low: 3 };
const PRIORITY_COLOR = { Critical: "#B91C1C", High: "#EA580C", Moderate: "#D97706", Low: "#64748B" };
const PRIORITY_BG = { Critical: "#FEE2E2", High: "#FFEDD5", Moderate: "#FEF3C7", Low: "#F1F5F9" };
const TASK_STATUSES = ["Not Started", "Doing", "Follow Up", "Done"];
const TASK_STATUS_COLOR = { "Not Started": "#64748B", "Doing": "#3B82F6", "Follow Up": "#D97706", "Done": "#16A34A" };
const TASK_STATUS_BG = { "Not Started": "#F1F5F9", "Doing": "#DBEAFE", "Follow Up": "#FEF3C7", "Done": "#DCFCE7" };
const FORECAST_CATS = ["Pipeline", "Stretch", "Commit"];
const FORECAST_COLOR = { Pipeline: "#64748B", Stretch: "#D97706", Commit: "#16A34A" };

const IMPORT_FIELDS = [
  { key: "sfId", label: "Opportunity ID", required: true },
  { key: "name", label: "Opportunity name", required: true },
  { key: "account", label: "Account name", required: true },
  { key: "stage", label: "Stage", required: true },
  { key: "closeDate", label: "Close date", required: true },
  { key: "owner", label: "Owner / rep", required: false },
  { key: "productMargin", label: "Product margin $", required: false },
  { key: "servicesMargin", label: "Services margin $", required: false },
  { key: "managedServices", label: "Managed services $", required: false },
];

const ACCOUNT_IMPORT_FIELDS = [
  { key: "accountId", label: "Account ID", required: false },
  { key: "name", label: "Account name", required: true },
  { key: "currentCustomer", label: "Current Ahead customer (Y/N)", required: false },
  { key: "mdPriority", label: "MD priority (Y/N)", required: false },
  { key: "strategic", label: "Ahead strategic (Y/N)", required: false },
  { key: "cd", label: "CD", required: false },
  { key: "owner", label: "Owner / rep", required: false },
];

const TASK_IMPORT_FIELDS = [
  { key: "name", label: "Name", required: true },
  { key: "account", label: "Account", required: true },
  { key: "cd", label: "CD", required: false },
  { key: "owner", label: "Owner", required: false },
  { key: "priority", label: "Priority", required: false },
  { key: "dueDate", label: "Due date", required: false },
  { key: "status", label: "Status", required: false },
  { key: "sseSme", label: "SSE / SME", required: false },
  { key: "isOpp", label: "Opp (marked if tied to an opportunity)", required: false },
  { key: "notes", label: "Notes / Next Steps", required: false },
];

// ---------- Color tokens (light theme) ----------
const C = {
  page: "#EEF2F7",
  card: "#FFFFFF",
  cardDark: "#0F172A",
  border: "#E2E8F0",
  borderStrong: "#CBD5E1",
  text: "#0F172A",
  textSoft: "#475569",
  textMute: "#94A3B8",
  onDark: "#F1F5F9",
  onDarkMute: "#94A3B8",
  red: "#DC2626",
  redBg: "#FEF2F2",
  amber: "#D97706",
  amberBg: "#FFFBEB",
  green: "#16A34A",
  greenBg: "#F0FDF4",
  blue: "#3B82F6",
  purple: "#7C3AED",
  sidebarActiveBg: "#EEF2FF",
  sidebarActiveText: "#4338CA",
};

const uid = () => Math.random().toString(36).slice(2, 10);
const todayStr = () => new Date().toISOString().slice(0, 10);

const fmtMoney = (n) => {
  const v = Number(n) || 0;
  return v.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
};
const fmtMoneyShort = (n) => {
  const v = Number(n) || 0;
  const sign = v < 0 ? "-" : "";
  const abs = Math.abs(v);
  if (abs >= 1000000) return sign + "$" + (abs / 1000000).toFixed(2).replace(/\.00$/, "") + "M";
  if (abs >= 1000) return sign + "$" + (abs / 1000).toFixed(1).replace(/\.0$/, "") + "k";
  return sign + "$" + abs.toFixed(0);
};
const fmtDate = (d) => {
  if (!d) return "—";
  const dt = new Date(d + "T00:00:00");
  if (isNaN(dt)) return "—";
  return dt.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" });
};
const daysUntil = (d) => {
  if (!d) return null;
  const dt = new Date(d + "T00:00:00");
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return Math.round((dt - today) / 86400000);
};
const daysSince = (d) => (d ? -daysUntil(d) : null);
const oppTotal = (o) => (Number(o.productMargin) || 0) + (Number(o.servicesMargin) || 0) + (Number(o.managedServices) || 0);
const getQuarter = (dateStr) => {
  if (!dateStr) return null;
  const dt = new Date(dateStr + "T00:00:00");
  if (isNaN(dt)) return null;
  const q = Math.floor(dt.getMonth() / 3) + 1;
  return `${dt.getFullYear()}-Q${q}`;
};
const currentQuarter = () => {
  const dt = new Date();
  const q = Math.floor(dt.getMonth() / 3) + 1;
  return `${dt.getFullYear()}-Q${q}`;
};
const quarterOptions = (data) => {
  const set = new Set([currentQuarter()]);
  data.opportunities.forEach(o => { const q = getQuarter(o.closeDate); if (q) set.add(q); });
  data.goals.forEach(g => set.add(g.period));
  return Array.from(set).sort();
};
const suggestForecast = (stage) => {
  if (stage === "Commit") return "Commit";
  if (stage === "Upside" || stage === "Strong Upside") return "Stretch";
  return "Pipeline";
};

const normalizeDate = (raw) => {
  if (raw === null || raw === undefined || raw === "") return "";
  if (typeof raw === "number") {
    const d = XLSX.SSF ? XLSX.SSF.parse_date_code(raw) : null;
    if (d) return `${d.y}-${String(d.m).padStart(2, "0")}-${String(d.d).padStart(2, "0")}`;
  }
  const s = String(raw).trim();
  const mdy = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdy) return `${mdy[3]}-${mdy[1].padStart(2, "0")}-${mdy[2].padStart(2, "0")}`;
  const iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (iso) return `${iso[1]}-${iso[2].padStart(2, "0")}-${iso[3].padStart(2, "0")}`;
  const parsed = new Date(s);
  if (!isNaN(parsed)) return parsed.toISOString().slice(0, 10);
  return "";
};
const numOrZero = (v) => {
  if (v === null || v === undefined || v === "") return 0;
  const n = Number(String(v).replace(/[^0-9.-]/g, ""));
  return isNaN(n) ? 0 : n;
};
const toBool = (v) => {
  if (typeof v === "boolean") return v;
  if (v === null || v === undefined) return false;
  const s = String(v).trim().toLowerCase();
  return ["yes", "y", "true", "1", "x", "✓", "active"].includes(s);
};
const normName = (s) => (s || "").toString().trim().toLowerCase();
const guessField = (headers, keywords, exclude = []) => {
  const lower = headers.map(h => h.toLowerCase());
  for (const kw of keywords) {
    const idx = lower.findIndex(h => h.includes(kw) && !exclude.some(ex => h.includes(ex)));
    if (idx >= 0) return headers[idx];
  }
  return "";
};

const DEFAULT_DATA = {
  orgName: "Sales Organization",
  regions: [{ id: "central", name: "Central" }],
  teams: [{ id: "mi-in", name: "MI / IN Team", regionId: "central", reps: ["Rep 1", "Rep 2"] }],
  opportunities: [],
  activities: [],
  tasks: [],
  goals: [],
  accounts: [],
  settings: { coverageTarget: 3, staleDays: 30 },
  importConfig: { colMap: {}, stageMap: {}, lastImportDate: null, lastImportFileName: null },
  accountImportConfig: { colMap: {}, lastImportDate: null, lastImportFileName: null },
  taskImportConfig: { colMap: {}, lastImportDate: null, lastImportFileName: null },
};

// ---------- Region / team / rep helpers ----------
const newId = () => uid();
const orphanRepNames = (data) => {
  const known = new Set(data.teams.flatMap(t => t.reps));
  const found = new Set();
  (data.opportunities || []).forEach(o => { if (o.rep && !known.has(o.rep)) found.add(o.rep); });
  (data.tasks || []).forEach(t => { if (t.rep && !known.has(t.rep)) found.add(t.rep); });
  (data.activities || []).forEach(a => { if (a.rep && !known.has(a.rep)) found.add(a.rep); });
  (data.goals || []).forEach(g => { if (g.rep && !known.has(g.rep)) found.add(g.rep); });
  (data.accounts || []).forEach(a => { if (a.rep && !known.has(a.rep)) found.add(a.rep); });
  return Array.from(found);
};
const allRepNames = (data) => {
  const list = [...data.teams.flatMap(t => t.reps), ...orphanRepNames(data)];
  const deduped = Array.from(new Set(list));
  return deduped.length ? deduped : ["Rep 1", "Rep 2"];
};
const ensureUnassignedTeam = (regions, teams) => {
  let nextRegions = regions.some(r => r.id === "unassigned") ? regions : [...regions, { id: "unassigned", name: "Unassigned" }];
  let team = teams.find(t => t.id === "unassigned");
  let nextTeams = teams;
  if (!team) {
    team = { id: "unassigned", name: "Unassigned", regionId: "unassigned", reps: [] };
    nextTeams = [...teams, team];
  }
  return { regions: nextRegions, teams: nextTeams, team };
};

const mergeLoadedData = (parsed) => {
  let merged = {
    ...DEFAULT_DATA, ...parsed,
    settings: { ...DEFAULT_DATA.settings, ...(parsed.settings || {}) },
    importConfig: { ...DEFAULT_DATA.importConfig, ...(parsed.importConfig || {}) },
    accountImportConfig: { ...DEFAULT_DATA.accountImportConfig, ...(parsed.accountImportConfig || {}) },
    taskImportConfig: { ...DEFAULT_DATA.taskImportConfig, ...(parsed.taskImportConfig || {}) },
    accounts: parsed.accounts || [],
  };
  // migrate legacy flat teamName/reps shape into regions/teams
  if (!parsed.teams && (parsed.teamName || parsed.reps)) {
    merged.orgName = parsed.teamName || DEFAULT_DATA.orgName;
    merged.regions = [{ id: "central", name: "Central" }];
    merged.teams = [{ id: "mi-in", name: parsed.teamName || "Team 1", regionId: "central", reps: parsed.reps || ["Rep 1", "Rep 2"] }];
  }
  if (!merged.regions || merged.regions.length === 0) merged.regions = DEFAULT_DATA.regions;
  if (!merged.teams || merged.teams.length === 0) merged.teams = DEFAULT_DATA.teams;
  return merged;
};

const SYNC_POLL_MS = 6000;

function useStore() {
  const [data, setDataState] = useState(DEFAULT_DATA);
  const [loaded, setLoaded] = useState(false);
  const [saving, setSaving] = useState(false);
  const [lastSyncedAt, setLastSyncedAt] = useState(null);
  const [loadError, setLoadError] = useState(false);
  const lastRawRef = useRef(null);

  const attemptLoad = useCallback(async () => {
    setLoadError(false);
    for (let attempt = 0; attempt < 3; attempt++) {
      try {
        const res = await window.storage.get(STORAGE_KEY, true);
        if (res && res.value) {
          lastRawRef.current = res.value;
          setDataState(mergeLoadedData(JSON.parse(res.value)));
        }
        setLoaded(true);
        setLastSyncedAt(Date.now());
        return;
      } catch (e) {
        if (attempt < 2) {
          await new Promise(r => setTimeout(r, 1200));
          continue;
        }
        // All retries failed. Do NOT silently show a blank slate as if it were
        // genuinely empty — surface this so nobody mistakes a load failure for lost data.
        setLoadError(true);
        setLoaded(true);
      }
    }
  }, []);

  useEffect(() => { attemptLoad(); }, [attemptLoad]);

  const save = useCallback(async (next) => {
    setDataState(next);
    setSaving(true);
    try {
      const json = JSON.stringify(next);
      lastRawRef.current = json;
      await window.storage.set(STORAGE_KEY, json, true);
      setLastSyncedAt(Date.now());
    } catch (e) {
      console.error("Storage save failed", e);
    } finally {
      setSaving(false);
    }
  }, []);

  // Poll for changes made by other team members and pull them in automatically
  useEffect(() => {
    if (!loaded || loadError) return;
    const interval = setInterval(async () => {
      try {
        const res = await window.storage.get(STORAGE_KEY, true);
        if (res && res.value && res.value !== lastRawRef.current) {
          lastRawRef.current = res.value;
          setDataState(mergeLoadedData(JSON.parse(res.value)));
          setLastSyncedAt(Date.now());
        }
      } catch (e) {
        // ignore transient poll failures
      }
    }, SYNC_POLL_MS);
    return () => clearInterval(interval);
  }, [loaded, loadError]);

  return { data, setData: save, loaded, saving, lastSyncedAt, loadError, retryLoad: attemptLoad };
}

function Modal({ title, onClose, children, wide }) {
  return (
    <div style={{
      position: "fixed", inset: 0, background: "rgba(15,23,42,0.45)", zIndex: 100,
      display: "flex", alignItems: "flex-start", justifyContent: "center",
      padding: "40px 16px", overflowY: "auto"
    }} onClick={onClose}>
      <div
        onClick={e => e.stopPropagation()}
        style={{
          background: C.card, border: `1px solid ${C.border}`, borderRadius: 14,
          width: "100%", maxWidth: wide ? 640 : 460, padding: 24, color: C.text,
          boxShadow: "0 20px 50px rgba(15,23,42,0.18)"
        }}
      >
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 18 }}>
          <h3 style={{ margin: 0, fontSize: 16, fontWeight: 700, letterSpacing: 0.1 }}>{title}</h3>
          <button onClick={onClose} style={iconBtnStyle}><X size={16} /></button>
        </div>
        {children}
      </div>
    </div>
  );
}

const iconBtnStyle = {
  background: "transparent", border: `1px solid ${C.border}`, borderRadius: 8,
  color: C.textSoft, width: 30, height: 30, display: "flex", alignItems: "center",
  justifyContent: "center", cursor: "pointer"
};
const labelStyle = { fontSize: 12, color: C.textSoft, marginBottom: 6, display: "block", fontWeight: 600, letterSpacing: 0.2 };
const inputStyle = {
  width: "100%", background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8,
  color: C.text, padding: "9px 11px", fontSize: 13.5, marginBottom: 14, fontFamily: "inherit",
  boxSizing: "border-box"
};
const rowStyle = { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 };
const primaryBtn = {
  background: C.text, color: "#FFFFFF", border: "none", borderRadius: 8,
  padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer"
};
const ghostBtn = {
  background: "transparent", color: C.textSoft, border: `1px solid ${C.border}`, borderRadius: 8,
  padding: "9px 16px", fontSize: 13.5, fontWeight: 600, cursor: "pointer"
};
const dangerText = { color: C.red };

function Field({ label, children }) {
  return <div><label style={labelStyle}>{label}</label>{children}</div>;
}

// ---------- Opportunity form ----------
function OpportunityForm({ initial, reps, tasks, onSave, onCancel, onDelete, onOpenTask, onAddTask }) {
  const [f, setF] = useState(initial || {
    name: "", account: "", stage: "Prospecting", closeDate: "",
    productMargin: "", servicesMargin: "", managedServices: "",
    rep: reps[0] || "", nextStep: "", notes: "", forecastCategory: "Pipeline"
  });
  const [fcTouched, setFcTouched] = useState(!!initial);
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const setStage = (stage) => {
    setF(prev => ({ ...prev, stage, forecastCategory: fcTouched ? prev.forecastCategory : suggestForecast(stage) }));
  };
  const submit = () => {
    if (!(f.name && f.account)) return;
    const stageChanged = !initial || initial.stage !== f.stage;
    onSave({
      ...f,
      id: initial?.id || uid(),
      lastStageChange: stageChanged ? todayStr() : (initial?.lastStageChange || todayStr()),
      createdAt: initial?.createdAt || todayStr(),
    });
  };
  return (
    <div>
      {f.sfId && (
        <div style={{
          display: "flex", alignItems: "center", gap: 6, fontSize: 11.5, color: C.textMute,
          marginBottom: 14, background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8, padding: "7px 10px"
        }}>
          <FileSpreadsheet size={13} />
          Synced from Salesforce (ID {f.sfId}){f.lastSynced ? ` · last synced ${fmtDate(f.lastSynced)}` : ""}
          {f.sfStatus === "removed" && <span style={{ color: C.amber, marginLeft: 4 }}>· no longer in latest export</span>}
        </div>
      )}
      <Field label="Opportunity name">
        <input style={inputStyle} value={f.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Acme Corp – Network Refresh" />
      </Field>
      <div style={rowStyle}>
        <Field label="Account"><input style={inputStyle} value={f.account} onChange={e => set("account", e.target.value)} placeholder="Company name" /></Field>
        <Field label="Owner (rep)">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Stage">
          <select style={inputStyle} value={f.stage} onChange={e => setStage(e.target.value)}>
            {STAGES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Expected / actual close date">
          <input type="date" style={inputStyle} value={f.closeDate} onChange={e => set("closeDate", e.target.value)} />
        </Field>
      </div>
      <Field label="Forecast category">
        <select style={inputStyle} value={f.forecastCategory || "Pipeline"} onChange={e => { setFcTouched(true); set("forecastCategory", e.target.value); }}>
          {FORECAST_CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
      </Field>
      <label style={{ ...labelStyle, marginTop: 2 }}>Revenue by category ($)</label>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
        {CATEGORIES.map(c => (
          <div key={c.key}>
            <label style={{ fontSize: 11, color: c.color, marginBottom: 4, display: "block", fontWeight: 600 }}>{c.label}</label>
            <input type="number" style={{ ...inputStyle, marginBottom: 14 }} value={f[c.key]}
              onChange={e => set(c.key, e.target.value)} placeholder="0" />
          </div>
        ))}
      </div>
      <Field label="Next step">
        <input style={inputStyle} value={f.nextStep} onChange={e => set("nextStep", e.target.value)} placeholder="What happens next?" />
      </Field>
      <Field label="Notes">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>
      {initial && (
        <div style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <label style={{ ...labelStyle, marginBottom: 0 }}>Tasks for this account ({(tasks || []).length})</label>
            {onAddTask && (
              <button type="button" style={{ ...ghostBtn, padding: "4px 9px", fontSize: 11.5 }} onClick={onAddTask}>
                <Plus size={12} style={{ verticalAlign: -2, marginRight: 4 }} />Add task
              </button>
            )}
          </div>
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
            {(tasks || []).length === 0 && (
              <div style={{ padding: "10px 12px", fontSize: 12.5, color: C.textMute, background: "#F8FAFC" }}>No tasks linked to this account yet.</div>
            )}
            {(tasks || []).map((t, i) => {
              const overdue = t.status !== "Done" && t.dueDate && daysUntil(t.dueDate) < 0;
              return (
                <div key={t.id} onClick={() => onOpenTask && onOpenTask(t)} style={{
                  display: "flex", alignItems: "center", gap: 8, padding: "9px 12px",
                  borderBottom: i < tasks.length - 1 ? `1px solid ${C.border}` : "none",
                  background: "#F8FAFC", cursor: onOpenTask ? "pointer" : "default"
                }}>
                  <span style={{
                    fontSize: 13, color: t.status === "Done" ? C.textMute : C.text,
                    textDecoration: t.status === "Done" ? "line-through" : "none", flex: 1
                  }}>{t.title}</span>
                  <Pill color={PRIORITY_COLOR[t.priority] || C.textMute} bg={PRIORITY_BG[t.priority] || "#F1F5F9"}>{t.priority}</Pill>
                  <span style={{ fontSize: 11.5, minWidth: 66, textAlign: "right", color: overdue ? C.red : C.textMute }}>{fmtDate(t.dueDate)}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <div>
          {initial && onDelete && (
            <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => onDelete(initial.id)}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Delete
            </button>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={ghostBtn} onClick={onCancel}>Cancel</button>
          <button style={primaryBtn} onClick={submit}>{initial ? "Save changes" : "Add opportunity"}</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Activity form ----------
function ActivityForm({ initial, reps, opportunities, onSave, onCancel, onDelete }) {
  const today = todayStr();
  const [f, setF] = useState(initial || {
    type: "Call", date: today, rep: reps[0] || "", account: "", contact: "", oppId: "", notes: ""
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  return (
    <div>
      <div style={rowStyle}>
        <Field label="Activity type">
          <select style={inputStyle} value={f.type} onChange={e => set("type", e.target.value)}>
            {ACTIVITY_TYPES.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
        <Field label="Date">
          <input type="date" style={inputStyle} value={f.date} onChange={e => set("date", e.target.value)} />
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Rep">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="Linked opportunity (optional)">
          <select style={inputStyle} value={f.oppId} onChange={e => set("oppId", e.target.value)}>
            <option value="">None</option>
            {opportunities.map(o => <option key={o.id} value={o.id}>{o.name}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Account"><input style={inputStyle} value={f.account} onChange={e => set("account", e.target.value)} /></Field>
        <Field label="Contact"><input style={inputStyle} value={f.contact} onChange={e => set("contact", e.target.value)} /></Field>
      </div>
      <Field label="Notes / outcome">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <div>
          {initial && onDelete && (
            <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => onDelete(initial.id)}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Delete
            </button>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={ghostBtn} onClick={onCancel}>Cancel</button>
          <button style={primaryBtn} onClick={() => f.account && onSave({ ...f, id: initial?.id || uid() })}>
            {initial ? "Save changes" : "Log activity"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Account search/autocomplete (free text) ----------
function AccountAutocomplete({ accounts, value, onChange }) {
  const [open, setOpen] = useState(false);
  const q = (value || "").trim().toLowerCase();
  const matches = (q
    ? accounts.filter(a => a.name.toLowerCase().includes(q))
    : accounts
  ).slice(0, 8);
  return (
    <div style={{ position: "relative", marginBottom: 14 }}>
      <input
        style={{ ...inputStyle, marginBottom: 0 }}
        placeholder="Account name"
        value={value}
        onFocus={() => setOpen(true)}
        onChange={e => { onChange(e.target.value); setOpen(true); }}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
      />
      {open && matches.length > 0 && (
        <div style={{
          position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 30,
          background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
          boxShadow: "0 8px 24px rgba(15,23,42,0.14)", maxHeight: 200, overflowY: "auto"
        }}>
          {matches.map(a => (
            <div key={a.id}
              onMouseDown={e => { e.preventDefault(); onChange(a.name); setOpen(false); }}
              style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, color: C.text }}
            >
              {a.name}{a.cd && <span style={{ color: C.textMute, fontSize: 11.5 }}> · CD: {a.cd}</span>}
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

// ---------- Task form ----------
function TaskForm({ initial, reps, accounts, onSave, onCancel, onDelete }) {
  const isEdit = !!(initial && initial.id);
  const [f, setF] = useState({
    title: "", account: "", cd: "", sseSme: "", rep: reps[0] || "", dueDate: "",
    priority: "Moderate", status: "Not Started", isOpp: false, notes: "",
    ...(initial || {}),
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  return (
    <div>
      <Field label="Name">
        <input style={inputStyle} value={f.title} onChange={e => set("title", e.target.value)} placeholder="What needs to happen?" />
      </Field>
      <div style={rowStyle}>
        <Field label="Account">
          <AccountAutocomplete accounts={accounts} value={f.account} onChange={v => set("account", v)} />
        </Field>
        <Field label="CD">
          <input style={inputStyle} value={f.cd} onChange={e => set("cd", e.target.value)} placeholder="Client / coverage director" />
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Owner">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
        <Field label="SSE / SME">
          <input style={inputStyle} value={f.sseSme} onChange={e => set("sseSme", e.target.value)} placeholder="Technical resource" />
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Priority">
          <select style={inputStyle} value={f.priority} onChange={e => set("priority", e.target.value)}>
            {PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </Field>
        <Field label="Due date">
          <input type="date" style={inputStyle} value={f.dueDate} onChange={e => set("dueDate", e.target.value)} />
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Status">
          <select style={inputStyle} value={f.status} onChange={e => set("status", e.target.value)}>
            {TASK_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label=" ">
          <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: C.text, marginTop: 9, cursor: "pointer" }}>
            <input type="checkbox" checked={!!f.isOpp} onChange={e => set("isOpp", e.target.checked)}
              style={{ width: 15, height: 15, accentColor: C.green, cursor: "pointer" }} />
            Tied to an opportunity
          </label>
        </Field>
      </div>
      <Field label="Notes / Next steps">
        <textarea style={{ ...inputStyle, minHeight: 50, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <div>
          {isEdit && onDelete && (
            <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => onDelete(initial.id)}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Delete
            </button>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={ghostBtn} onClick={onCancel}>Cancel</button>
          <button style={primaryBtn} onClick={() => f.title && onSave({ ...f, id: initial?.id || uid() })}>
            {isEdit ? "Save changes" : "Add task"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Goal form ----------
function GoalForm({ initial, reps, period, onSave, onCancel, onDelete }) {
  const [f, setF] = useState(initial || {
    period: period, rep: reps[0] || "", productMarginTarget: "", servicesMarginTarget: "", managedServicesTarget: ""
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  return (
    <div>
      <div style={rowStyle}>
        <Field label="Period (e.g. 2026-Q3)">
          <input style={inputStyle} value={f.period} onChange={e => set("period", e.target.value)} />
        </Field>
        <Field label="Rep">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
      </div>
      <label style={labelStyle}>Quarterly target by category ($)</label>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
        {CATEGORIES.map(c => (
          <div key={c.key}>
            <label style={{ fontSize: 11, color: c.color, marginBottom: 4, display: "block", fontWeight: 600 }}>{c.label}</label>
            <input type="number" style={{ ...inputStyle, marginBottom: 14 }}
              value={f[c.key + "Target"]} onChange={e => set(c.key + "Target", e.target.value)} placeholder="0" />
          </div>
        ))}
      </div>
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <div>
          {initial && onDelete && (
            <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => onDelete(initial.id)}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Delete
            </button>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={ghostBtn} onClick={onCancel}>Cancel</button>
          <button style={primaryBtn} onClick={() => f.period && onSave({ ...f, id: initial?.id || uid() })}>
            {initial ? "Save changes" : "Set goal"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Account form ----------
function AccountForm({ initial, reps, computed, tasks, onSave, onCancel, onDelete, onOpenTask, onAddTask }) {
  const [f, setF] = useState(initial || {
    name: "", cd: "", rep: reps[0] || "", currentCustomer: false, mdPriority: false, strategic: false, notes: ""
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const checkboxRow = (key, label) => (
    <label style={{ display: "flex", alignItems: "center", gap: 8, fontSize: 13, color: C.text, marginBottom: 10, cursor: "pointer" }}>
      <input type="checkbox" checked={!!f[key]} onChange={e => set(key, e.target.checked)}
        style={{ width: 15, height: 15, accentColor: C.green, cursor: "pointer" }} />
      {label}
    </label>
  );
  return (
    <div>
      {computed && (
        <div style={{
          display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap"
        }}>
          <Pill color={computed.hasActiveOpp ? C.green : C.textMute} bg={computed.hasActiveOpp ? "#DCFCE7" : "#F1F5F9"}>
            {computed.hasActiveOpp ? "Active opportunity" : "No active opportunity"}
          </Pill>
          <Pill color={computed.hasMet ? C.green : C.textMute} bg={computed.hasMet ? "#DCFCE7" : "#F1F5F9"}>
            {computed.hasMet ? "Met with them" : "Not met yet"}
          </Pill>
        </div>
      )}
      <Field label="Account name">
        <input style={inputStyle} value={f.name} onChange={e => set("name", e.target.value)} placeholder="Company name" />
      </Field>
      <div style={rowStyle}>
        <Field label="CD"><input style={inputStyle} value={f.cd} onChange={e => set("cd", e.target.value)} placeholder="Client / coverage director" /></Field>
        <Field label="Rep">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
      </div>
      <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, marginBottom: 14 }}>
        {checkboxRow("currentCustomer", "Current Ahead customer")}
        {checkboxRow("mdPriority", "MD priority")}
        {checkboxRow("strategic", "Ahead strategic")}
      </div>
      <Field label="Notes">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>
      {initial && (
        <div style={{ marginBottom: 16 }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
            <label style={{ ...labelStyle, marginBottom: 0 }}>Tasks ({(tasks || []).length})</label>
            {onAddTask && (
              <button type="button" style={{ ...ghostBtn, padding: "4px 9px", fontSize: 11.5 }} onClick={onAddTask}>
                <Plus size={12} style={{ verticalAlign: -2, marginRight: 4 }} />Add task
              </button>
            )}
          </div>
          <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
            {(tasks || []).length === 0 && (
              <div style={{ padding: "10px 12px", fontSize: 12.5, color: C.textMute, background: "#F8FAFC" }}>No tasks linked to this account yet.</div>
            )}
            {(tasks || []).map((t, i) => {
              const overdue = t.status !== "Done" && t.dueDate && daysUntil(t.dueDate) < 0;
              return (
                <div key={t.id} onClick={() => onOpenTask && onOpenTask(t)} style={{
                  display: "flex", alignItems: "center", gap: 8, padding: "9px 12px",
                  borderBottom: i < tasks.length - 1 ? `1px solid ${C.border}` : "none",
                  background: "#F8FAFC", cursor: onOpenTask ? "pointer" : "default"
                }}>
                  <span style={{
                    fontSize: 13, color: t.status === "Done" ? C.textMute : C.text,
                    textDecoration: t.status === "Done" ? "line-through" : "none", flex: 1
                  }}>{t.title}</span>
                  <Pill color={PRIORITY_COLOR[t.priority] || C.textMute} bg={PRIORITY_BG[t.priority] || "#F1F5F9"}>{t.priority}</Pill>
                  <span style={{ fontSize: 11.5, minWidth: 66, textAlign: "right", color: overdue ? C.red : C.textMute }}>{fmtDate(t.dueDate)}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}
      <div style={{ display: "flex", justifyContent: "space-between", marginTop: 6 }}>
        <div>
          {initial && onDelete && (
            <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => onDelete(initial.id)}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Delete
            </button>
          )}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={ghostBtn} onClick={onCancel}>Cancel</button>
          <button style={primaryBtn} onClick={() => f.name && onSave({ ...f, id: initial?.id || uid() })}>
            {initial ? "Save changes" : "Add account"}
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Import wizard ----------
function ImportWizard({ data, onClose, onApply }) {
  const [step, setStep] = useState(1);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [colMap, setColMap] = useState(data.importConfig.colMap || {});
  const [stageMap, setStageMap] = useState(data.importConfig.stageMap || {});
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const isCsv = /\.csv$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        let parsedRows = [];
        if (isCsv) {
          const result = Papa.parse(evt.target.result, { header: true, skipEmptyLines: true });
          parsedRows = result.data;
        } else {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          parsedRows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "" });
        }
        parsedRows = parsedRows.filter(r => Object.values(r).some(v => String(v).trim() !== ""));
        if (parsedRows.length === 0) { setError("No rows found in that file."); return; }
        const hdrs = Object.keys(parsedRows[0]);
        setHeaders(hdrs);
        setRows(parsedRows);

        const savedMap = data.importConfig.colMap || {};
        const nextMap = {};
        IMPORT_FIELDS.forEach(f => {
          if (savedMap[f.key] && hdrs.includes(savedMap[f.key])) { nextMap[f.key] = savedMap[f.key]; return; }
          if (f.key === "sfId") nextMap[f.key] = guessField(hdrs, ["opportunity id", "record id", "id"]);
          else if (f.key === "name") nextMap[f.key] = guessField(hdrs, ["opportunity name", "name"], ["account", "id"]);
          else if (f.key === "account") nextMap[f.key] = guessField(hdrs, ["account name", "account"]);
          else if (f.key === "stage") nextMap[f.key] = guessField(hdrs, ["stage"]);
          else if (f.key === "closeDate") nextMap[f.key] = guessField(hdrs, ["close date", "closedate"]);
          else if (f.key === "owner") nextMap[f.key] = guessField(hdrs, ["owner", "rep"]);
          else if (f.key === "productMargin") nextMap[f.key] = guessField(hdrs, ["product margin", "product"]);
          else if (f.key === "servicesMargin") nextMap[f.key] = guessField(hdrs, ["services margin", "service"]);
          else if (f.key === "managedServices") nextMap[f.key] = guessField(hdrs, ["managed services", "managed"]);
        });
        setColMap(nextMap);
        setStep(2);
      } catch (err) {
        setError("Couldn't read that file. Make sure it's a .csv or .xlsx export.");
      }
    };
    if (isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
  };

  const uniqueStageValues = useMemo(() => {
    if (!colMap.stage) return [];
    const set = new Set();
    rows.forEach(r => { const v = (r[colMap.stage] || "").toString().trim(); if (v) set.add(v); });
    return Array.from(set);
  }, [rows, colMap.stage]);

  useEffect(() => {
    if (uniqueStageValues.length === 0) return;
    setStageMap(prev => {
      const next = { ...prev };
      uniqueStageValues.forEach(v => {
        if (next[v]) return;
        const lower = v.toLowerCase();
        const match = STAGES.find(s => s.toLowerCase() === lower) ||
          (lower.includes("won") ? "Closed Won" : lower.includes("lost") ? "Closed Lost" :
           lower.includes("commit") ? "Commit" :
           lower.includes("strong") && lower.includes("upside") ? "Strong Upside" :
           lower.includes("upside") ? "Upside" :
           lower.includes("negotiat") || lower.includes("propos") || lower.includes("quote") ? "Strong Upside" :
           lower.includes("qualif") ? "Pipeline" :
           lower.includes("pipeline") ? "Pipeline" : "Prospecting");
        next[v] = match;
      });
      return next;
    });
  }, [uniqueStageValues]);

  const requiredMapped = IMPORT_FIELDS.filter(f => f.required).every(f => colMap[f.key]);

  const diff = useMemo(() => {
    if (step < 4) return null;
    const existingBySfId = new Map(data.opportunities.filter(o => o.sfId).map(o => [o.sfId, o]));
    const seen = new Set();
    let added = 0, updated = 0;
    const newRepNames = new Set();
    const knownReps = allRepNames(data);
    rows.forEach(r => {
      const sfId = (r[colMap.sfId] || "").toString().trim();
      if (!sfId) return;
      seen.add(sfId);
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.includes(owner)) newRepNames.add(owner);
      if (existingBySfId.has(sfId)) updated++; else added++;
    });
    const removed = data.opportunities.filter(o => o.sfId && o.sfStatus !== "removed" && !seen.has(o.sfId));
    return { added, updated, removed, newRepNames: Array.from(newRepNames), totalRows: rows.length };
  }, [step, rows, colMap, data]);

  const commit = () => {
    const existingBySfId = new Map(data.opportunities.filter(o => o.sfId).map(o => [o.sfId, o]));
    const seen = new Set();
    let opps = data.opportunities.map(o => ({ ...o }));
    let regions = data.regions.map(r => ({ ...r }));
    let teams = data.teams.map(t => ({ ...t, reps: [...t.reps] }));
    const knownReps = new Set(allRepNames(data));

    const addUnassignedRep = (name) => {
      const ensured = ensureUnassignedTeam(regions, teams);
      regions = ensured.regions;
      teams = ensured.teams;
      const ut = teams.find(t => t.id === "unassigned");
      if (!ut.reps.includes(name)) ut.reps.push(name);
      knownReps.add(name);
    };

    rows.forEach(r => {
      const sfId = (r[colMap.sfId] || "").toString().trim();
      if (!sfId) return;
      seen.add(sfId);
      const stageRaw = (r[colMap.stage] || "").toString().trim();
      const stage = stageMap[stageRaw] || "Prospecting";
      const closeDate = normalizeDate(r[colMap.closeDate]);
      const name = (r[colMap.name] || "").toString().trim();
      const account = (r[colMap.account] || "").toString().trim();
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);
      const productMargin = colMap.productMargin ? numOrZero(r[colMap.productMargin]) : undefined;
      const servicesMargin = colMap.servicesMargin ? numOrZero(r[colMap.servicesMargin]) : undefined;
      const managedServices = colMap.managedServices ? numOrZero(r[colMap.managedServices]) : undefined;

      const existing = existingBySfId.get(sfId);
      if (existing) {
        const idx = opps.findIndex(o => o.id === existing.id);
        const stageChanged = opps[idx].stage !== stage;
        opps[idx] = {
          ...opps[idx],
          name: name || opps[idx].name,
          account: account || opps[idx].account,
          stage, closeDate,
          rep: owner || opps[idx].rep,
          productMargin: productMargin !== undefined ? productMargin : opps[idx].productMargin,
          servicesMargin: servicesMargin !== undefined ? servicesMargin : opps[idx].servicesMargin,
          managedServices: managedServices !== undefined ? managedServices : opps[idx].managedServices,
          forecastCategory: stageChanged ? suggestForecast(stage) : (opps[idx].forecastCategory || suggestForecast(stage)),
          lastStageChange: stageChanged ? todayStr() : (opps[idx].lastStageChange || todayStr()),
          sfStatus: "active", lastSynced: todayStr(),
        };
      } else {
        opps.push({
          id: uid(), sfId, name, account, stage, closeDate,
          rep: owner || "Unassigned",
          productMargin: productMargin || 0, servicesMargin: servicesMargin || 0, managedServices: managedServices || 0,
          nextStep: "", notes: "", sfStatus: "active", lastSynced: todayStr(),
          forecastCategory: suggestForecast(stage), lastStageChange: todayStr(), createdAt: todayStr(),
        });
      }
    });

    opps = opps.map(o => (o.sfId && o.sfStatus !== "removed" && !seen.has(o.sfId)) ? { ...o, sfStatus: "removed" } : o);

    onApply({
      ...data,
      regions, teams,
      opportunities: opps,
      importConfig: { colMap, stageMap, lastImportDate: todayStr(), lastImportFileName: fileName },
    });
  };

  return (
    <Modal title="Sync opportunities from Salesforce" wide onClose={onClose}>
      <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
        {["Upload", "Map columns", "Map stages", "Review"].map((s, i) => (
          <div key={s} style={{
            flex: 1, textAlign: "center", fontSize: 11, padding: "6px 4px", borderRadius: 6,
            background: step === i + 1 ? C.sidebarActiveBg : "transparent",
            color: step >= i + 1 ? C.sidebarActiveText : C.textMute, fontWeight: 700
          }}>{i + 1}. {s}</div>
        ))}
      </div>

      {step === 1 && (
        <div>
          <p style={{ fontSize: 13, color: C.textSoft, marginBottom: 16 }}>
            Export your opportunity report from Salesforce as .csv or .xlsx, then upload it here. Nothing leaves this browser session.
          </p>
          <div
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `1.5px dashed ${C.borderStrong}`, borderRadius: 10, padding: "32px 20px", textAlign: "center",
              cursor: "pointer", color: C.textSoft
            }}>
            <UploadCloud size={22} style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>Click to choose a file</div>
            <div style={{ fontSize: 11.5, marginTop: 4 }}>.csv or .xlsx</div>
          </div>
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          {error && <div style={{ color: C.red, fontSize: 12.5, marginTop: 10 }}>{error}</div>}
          {data.importConfig.lastImportDate && (
            <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 14 }}>
              Last synced {fmtDate(data.importConfig.lastImportDate)} from {data.importConfig.lastImportFileName}.
            </div>
          )}
        </div>
      )}

      {step === 2 && (
        <div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
            Match each field to a column from <strong style={{ color: C.text }}>{fileName}</strong>. This is remembered for next time.
          </p>
          {IMPORT_FIELDS.map(f => (
            <div key={f.key} style={{ display: "grid", gridTemplateColumns: "160px 1fr", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <label style={{ fontSize: 12.5, color: C.textSoft, fontWeight: 600 }}>{f.label}{f.required && <span style={{ color: C.red }}> *</span>}</label>
              <select style={{ ...inputStyle, marginBottom: 0 }} value={colMap[f.key] || ""} onChange={e => setColMap({ ...colMap, [f.key]: e.target.value })}>
                <option value="">— Not in file —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          ))}
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16 }}>
            <button style={ghostBtn} onClick={() => setStep(1)}>Back</button>
            <button style={{ ...primaryBtn, opacity: requiredMapped ? 1 : 0.5 }} disabled={!requiredMapped} onClick={() => setStep(3)}>
              Continue <ArrowRight size={13} style={{ verticalAlign: -2, marginLeft: 4 }} />
            </button>
          </div>
        </div>
      )}

      {step === 3 && (
        <div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
            Map each Salesforce stage value to a stage in the command center.
          </p>
          {uniqueStageValues.map(v => (
            <div key={v} style={{ display: "grid", gridTemplateColumns: "1fr 1fr", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <span style={{ fontSize: 12.5, color: C.text }}>{v}</span>
              <select style={{ ...inputStyle, marginBottom: 0 }} value={stageMap[v] || ""} onChange={e => setStageMap({ ...stageMap, [v]: e.target.value })}>
                {STAGES.map(s => <option key={s} value={s}>{s}</option>)}
              </select>
            </div>
          ))}
          {uniqueStageValues.length === 0 && <div style={{ fontSize: 12.5, color: C.textMute }}>No stage values found in that column.</div>}
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16 }}>
            <button style={ghostBtn} onClick={() => setStep(2)}>Back</button>
            <button style={primaryBtn} onClick={() => setStep(4)}>Continue <ArrowRight size={13} style={{ verticalAlign: -2, marginLeft: 4 }} /></button>
          </div>
        </div>
      )}

      {step === 4 && diff && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 16 }}>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.green }}>{diff.added}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>New opportunities</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.blue }}>{diff.updated}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Updated</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.amber }}>{diff.removed.length}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>No longer in export</div>
            </div>
          </div>
          {diff.removed.length > 0 && (
            <div style={{ marginBottom: 14 }}>
              <div style={{ fontSize: 12, color: C.textSoft, marginBottom: 6 }}>These will be marked "removed from CRM" (kept, not deleted, notes preserved):</div>
              <div style={{ maxHeight: 100, overflowY: "auto", fontSize: 12, color: C.text }}>
                {diff.removed.map(o => <div key={o.id} style={{ padding: "3px 0" }}>{o.name} — {o.account}</div>)}
              </div>
            </div>
          )}
          {diff.newRepNames.length > 0 && (
            <div style={{ fontSize: 12, color: C.textSoft, marginBottom: 14 }}>
              New rep name{diff.newRepNames.length > 1 ? "s" : ""} found: {diff.newRepNames.join(", ")}. They'll be placed on an "Unassigned" team — reassign them to a real team and region under Settings.
            </div>
          )}
          <div style={{ fontSize: 11.5, color: C.textMute, marginBottom: 14 }}>
            Manually-entered next steps and notes are never overwritten by a sync.
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button style={ghostBtn} onClick={() => setStep(3)}>Back</button>
            <button style={primaryBtn} onClick={commit}>
              <Check size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Apply sync
            </button>
          </div>
        </div>
      )}
    </Modal>
  );
}

// ---------- Account import wizard ----------
function AccountImportWizard({ data, onClose, onApply }) {
  const [step, setStep] = useState(1);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [colMap, setColMap] = useState(data.accountImportConfig.colMap || {});
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const isCsv = /\.csv$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        let parsedRows = [];
        if (isCsv) {
          const result = Papa.parse(evt.target.result, { header: true, skipEmptyLines: true });
          parsedRows = result.data;
        } else {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          parsedRows = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "" });
        }
        parsedRows = parsedRows.filter(r => Object.values(r).some(v => String(v).trim() !== ""));
        if (parsedRows.length === 0) { setError("No rows found in that file."); return; }
        const hdrs = Object.keys(parsedRows[0]);
        setHeaders(hdrs);
        setRows(parsedRows);

        const savedMap = data.accountImportConfig.colMap || {};
        const nextMap = {};
        ACCOUNT_IMPORT_FIELDS.forEach(f => {
          if (savedMap[f.key] && hdrs.includes(savedMap[f.key])) { nextMap[f.key] = savedMap[f.key]; return; }
          if (f.key === "accountId") nextMap[f.key] = guessField(hdrs, ["account id", "record id", "id"]);
          else if (f.key === "name") nextMap[f.key] = guessField(hdrs, ["account name", "name"], ["id"]);
          else if (f.key === "currentCustomer") nextMap[f.key] = guessField(hdrs, ["current customer", "customer"]);
          else if (f.key === "mdPriority") nextMap[f.key] = guessField(hdrs, ["md priority", "priority"]);
          else if (f.key === "strategic") nextMap[f.key] = guessField(hdrs, ["strategic"]);
          else if (f.key === "cd") nextMap[f.key] = guessField(hdrs, ["cd"]);
          else if (f.key === "owner") nextMap[f.key] = guessField(hdrs, ["owner", "rep"]);
        });
        setColMap(nextMap);
        setStep(2);
      } catch (err) {
        setError("Couldn't read that file. Make sure it's a .csv or .xlsx export.");
      }
    };
    if (isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
  };

  const requiredMapped = ACCOUNT_IMPORT_FIELDS.filter(f => f.required).every(f => colMap[f.key]);

  const diff = useMemo(() => {
    if (step < 3) return null;
    const byId = new Map(data.accounts.filter(a => a.accountId).map(a => [a.accountId, a]));
    const byName = new Map(data.accounts.map(a => [normName(a.name), a]));
    let added = 0, updated = 0;
    const newRepNames = new Set();
    const knownReps = allRepNames(data);
    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      if (!name) return;
      const accountId = colMap.accountId ? (r[colMap.accountId] || "").toString().trim() : "";
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.includes(owner)) newRepNames.add(owner);
      const existing = (accountId && byId.get(accountId)) || byName.get(normName(name));
      if (existing) updated++; else added++;
    });
    return { added, updated, newRepNames: Array.from(newRepNames), totalRows: rows.length };
  }, [step, rows, colMap, data]);

  const commit = () => {
    const byId = new Map(data.accounts.filter(a => a.accountId).map(a => [a.accountId, a]));
    const byName = new Map(data.accounts.map(a => [normName(a.name), a]));
    let accounts = data.accounts.map(a => ({ ...a }));
    let regions = data.regions.map(r => ({ ...r }));
    let teams = data.teams.map(t => ({ ...t, reps: [...t.reps] }));
    const knownReps = new Set(allRepNames(data));

    const addUnassignedRep = (name) => {
      const ensured = ensureUnassignedTeam(regions, teams);
      regions = ensured.regions; teams = ensured.teams;
      const ut = teams.find(t => t.id === "unassigned");
      if (!ut.reps.includes(name)) ut.reps.push(name);
      knownReps.add(name);
    };

    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      if (!name) return;
      const accountId = colMap.accountId ? (r[colMap.accountId] || "").toString().trim() : "";
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);

      const patch = {
        name,
        accountId: accountId || undefined,
        currentCustomer: colMap.currentCustomer ? toBool(r[colMap.currentCustomer]) : undefined,
        mdPriority: colMap.mdPriority ? toBool(r[colMap.mdPriority]) : undefined,
        strategic: colMap.strategic ? toBool(r[colMap.strategic]) : undefined,
        cd: colMap.cd ? (r[colMap.cd] || "").toString().trim() : undefined,
        rep: owner || undefined,
      };

      const existing = (accountId && byId.get(accountId)) || byName.get(normName(name));
      if (existing) {
        const idx = accounts.findIndex(a => a.id === existing.id);
        accounts[idx] = {
          ...accounts[idx],
          name: patch.name,
          accountId: patch.accountId !== undefined ? patch.accountId : accounts[idx].accountId,
          currentCustomer: patch.currentCustomer !== undefined ? patch.currentCustomer : accounts[idx].currentCustomer,
          mdPriority: patch.mdPriority !== undefined ? patch.mdPriority : accounts[idx].mdPriority,
          strategic: patch.strategic !== undefined ? patch.strategic : accounts[idx].strategic,
          cd: patch.cd !== undefined ? patch.cd : accounts[idx].cd,
          rep: patch.rep !== undefined ? patch.rep : accounts[idx].rep,
        };
      } else {
        accounts.push({
          id: uid(), name: patch.name, accountId: patch.accountId || "",
          currentCustomer: patch.currentCustomer || false, mdPriority: patch.mdPriority || false, strategic: patch.strategic || false,
          cd: patch.cd || "", rep: patch.rep || "", notes: "", createdAt: todayStr(),
        });
      }
    });

    onApply({
      ...data,
      regions, teams, accounts,
      accountImportConfig: { colMap, lastImportDate: todayStr(), lastImportFileName: fileName },
    });
  };

  return (
    <Modal title="Import accounts from Excel" wide onClose={onClose}>
      <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
        {["Upload", "Map columns", "Review"].map((s, i) => (
          <div key={s} style={{
            flex: 1, textAlign: "center", fontSize: 11, padding: "6px 4px", borderRadius: 6,
            background: step === i + 1 ? C.sidebarActiveBg : "transparent",
            color: step >= i + 1 ? C.sidebarActiveText : C.textMute, fontWeight: 700
          }}>{i + 1}. {s}</div>
        ))}
      </div>

      {step === 1 && (
        <div>
          <p style={{ fontSize: 13, color: C.textSoft, marginBottom: 16 }}>
            Export your account list from Excel as .csv or .xlsx, then upload it here. Nothing leaves this browser session.
          </p>
          <div
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `1.5px dashed ${C.borderStrong}`, borderRadius: 10, padding: "32px 20px", textAlign: "center",
              cursor: "pointer", color: C.textSoft
            }}>
            <UploadCloud size={22} style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>Click to choose a file</div>
            <div style={{ fontSize: 11.5, marginTop: 4 }}>.csv or .xlsx</div>
          </div>
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          {error && <div style={{ color: C.red, fontSize: 12.5, marginTop: 10 }}>{error}</div>}
          {data.accountImportConfig.lastImportDate && (
            <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 14 }}>
              Last imported {fmtDate(data.accountImportConfig.lastImportDate)} from {data.accountImportConfig.lastImportFileName}.
            </div>
          )}
        </div>
      )}

      {step === 2 && (
        <div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
            Match each field to a column from <strong style={{ color: C.text }}>{fileName}</strong>. This is remembered for next time.
            For Yes/No columns, values like "Yes", "Y", "True", "1", or "X" are read as checked.
          </p>
          {ACCOUNT_IMPORT_FIELDS.map(f => (
            <div key={f.key} style={{ display: "grid", gridTemplateColumns: "200px 1fr", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <label style={{ fontSize: 12.5, color: C.textSoft, fontWeight: 600 }}>{f.label}{f.required && <span style={{ color: C.red }}> *</span>}</label>
              <select style={{ ...inputStyle, marginBottom: 0 }} value={colMap[f.key] || ""} onChange={e => setColMap({ ...colMap, [f.key]: e.target.value })}>
                <option value="">— Not in file —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          ))}
          <p style={{ fontSize: 11.5, color: C.textMute, marginTop: 4 }}>
            Matching uses Account ID when available, otherwise account name. "Active opportunity" and "met with them" aren't imported — the command center tracks those automatically from your pipeline and activity log.
          </p>
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16 }}>
            <button style={ghostBtn} onClick={() => setStep(1)}>Back</button>
            <button style={{ ...primaryBtn, opacity: requiredMapped ? 1 : 0.5 }} disabled={!requiredMapped} onClick={() => setStep(3)}>
              Continue <ArrowRight size={13} style={{ verticalAlign: -2, marginLeft: 4 }} />
            </button>
          </div>
        </div>
      )}

      {step === 3 && diff && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, marginBottom: 16 }}>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.green }}>{diff.added}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>New accounts</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.blue }}>{diff.updated}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Updated</div>
            </div>
          </div>
          {diff.newRepNames.length > 0 && (
            <div style={{ fontSize: 12, color: C.textSoft, marginBottom: 14 }}>
              New rep name{diff.newRepNames.length > 1 ? "s" : ""} found: {diff.newRepNames.join(", ")}. They'll be placed on an "Unassigned" team — reassign them under Settings.
            </div>
          )}
          <div style={{ fontSize: 11.5, color: C.textMute, marginBottom: 14 }}>
            Manually-entered notes on existing accounts are never overwritten by an import.
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button style={ghostBtn} onClick={() => setStep(2)}>Back</button>
            <button style={primaryBtn} onClick={commit}>
              <Check size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Apply import
            </button>
          </div>
        </div>
      )}
    </Modal>
  );
}

// ---------- Task import wizard ----------
function TaskImportWizard({ data, reps, onClose, onApply }) {
  const [step, setStep] = useState(1);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [colMap, setColMap] = useState(data.taskImportConfig.colMap || {});
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const handleFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setError("");
    setFileName(file.name);
    const isCsv = /\.csv$/i.test(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        let parsedRows = [];
        if (isCsv) {
          const result = Papa.parse(evt.target.result, { header: true, skipEmptyLines: true });
          parsedRows = result.data;
        } else {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes("gsd")) || wb.SheetNames[0];
          const sheet = wb.Sheets[sheetName];
          // GSD-style trackers have a title row above the real header row; find the row that looks like headers
          const raw = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", header: 1 });
          let headerRowIdx = 0;
          for (let i = 0; i < Math.min(raw.length, 5); i++) {
            if (raw[i].some(c => String(c).trim().toLowerCase() === "name")) { headerRowIdx = i; break; }
          }
          const hdrs = raw[headerRowIdx].map(h => String(h).trim());
          parsedRows = raw.slice(headerRowIdx + 1).map(r => {
            const obj = {};
            hdrs.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ""; });
            return obj;
          });
        }
        parsedRows = parsedRows.filter(r => Object.values(r).some(v => String(v).trim() !== ""));
        if (parsedRows.length === 0) { setError("No rows found in that file."); return; }
        const hdrs = Object.keys(parsedRows[0]);
        setHeaders(hdrs);
        setRows(parsedRows);

        const savedMap = data.taskImportConfig.colMap || {};
        const nextMap = {};
        TASK_IMPORT_FIELDS.forEach(f => {
          if (savedMap[f.key] && hdrs.includes(savedMap[f.key])) { nextMap[f.key] = savedMap[f.key]; return; }
          if (f.key === "name") nextMap[f.key] = guessField(hdrs, ["name"], ["account"]);
          else if (f.key === "account") nextMap[f.key] = guessField(hdrs, ["account"]);
          else if (f.key === "cd") nextMap[f.key] = guessField(hdrs, ["cd"]);
          else if (f.key === "owner") nextMap[f.key] = guessField(hdrs, ["owner"]);
          else if (f.key === "priority") nextMap[f.key] = guessField(hdrs, ["priority"]);
          else if (f.key === "dueDate") nextMap[f.key] = guessField(hdrs, ["due date", "duedate"]);
          else if (f.key === "status") nextMap[f.key] = guessField(hdrs, ["status"]);
          else if (f.key === "sseSme") nextMap[f.key] = guessField(hdrs, ["sse", "sme"]);
          else if (f.key === "isOpp") nextMap[f.key] = guessField(hdrs, ["opp"]);
          else if (f.key === "notes") nextMap[f.key] = guessField(hdrs, ["notes", "next steps"]);
        });
        setColMap(nextMap);
        setStep(2);
      } catch (err) {
        setError("Couldn't read that file. Make sure it's a .csv or .xlsx export.");
      }
    };
    if (isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
  };

  const requiredMapped = TASK_IMPORT_FIELDS.filter(f => f.required).every(f => colMap[f.key]);
  const normPriority = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    if (s === "critical") return "Critical";
    if (s === "high") return "High";
    if (s === "moderate" || s === "medium") return "Moderate";
    if (s === "low") return "Low";
    return "Moderate";
  };
  const normStatus = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    if (s === "done" || s === "complete" || s === "completed") return "Done";
    if (s === "doing" || s === "in progress") return "Doing";
    if (s === "follow up" || s === "followup") return "Follow Up";
    return "Not Started";
  };

  const diff = useMemo(() => {
    if (step < 3) return null;
    const existingByKey = new Map(data.tasks.map(t => [normName(t.title) + "|" + normName(t.account || ""), t]));
    let added = 0, updated = 0, skipped = 0;
    const newRepNames = new Set();
    const knownReps = allRepNames(data);
    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      const account = colMap.account ? (r[colMap.account] || "").toString().trim() : "";
      if (!name) { skipped++; return; }
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.includes(owner)) newRepNames.add(owner);
      const key = normName(name) + "|" + normName(account);
      if (existingByKey.has(key)) updated++; else added++;
    });
    return { added, updated, skipped, newRepNames: Array.from(newRepNames), totalRows: rows.length };
  }, [step, rows, colMap, data]);

  const commit = () => {
    const existingByKey = new Map(data.tasks.map(t => [normName(t.title) + "|" + normName(t.account || ""), t]));
    let tasks = data.tasks.map(t => ({ ...t }));
    let regions = data.regions.map(r => ({ ...r }));
    let teams = data.teams.map(t => ({ ...t, reps: [...t.reps] }));
    const knownReps = new Set(allRepNames(data));

    const addUnassignedRep = (name) => {
      const ensured = ensureUnassignedTeam(regions, teams);
      regions = ensured.regions; teams = ensured.teams;
      const ut = teams.find(t => t.id === "unassigned");
      if (!ut.reps.includes(name)) ut.reps.push(name);
      knownReps.add(name);
    };

    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      if (!name) return;
      const account = colMap.account ? (r[colMap.account] || "").toString().trim() : "";
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);

      const patch = {
        title: name,
        account,
        cd: colMap.cd ? (r[colMap.cd] || "").toString().trim() : "",
        rep: owner || (teams[0]?.reps[0] || "Unassigned"),
        priority: colMap.priority ? normPriority(r[colMap.priority]) : "Moderate",
        dueDate: colMap.dueDate ? normalizeDate(r[colMap.dueDate]) : "",
        status: colMap.status ? normStatus(r[colMap.status]) : "Not Started",
        sseSme: colMap.sseSme ? (r[colMap.sseSme] || "").toString().trim() : "",
        isOpp: colMap.isOpp ? !!(r[colMap.isOpp] && r[colMap.isOpp].toString().trim()) : false,
        notes: colMap.notes ? (r[colMap.notes] || "").toString().trim() : "",
      };

      const key = normName(name) + "|" + normName(account);
      const existing = existingByKey.get(key);
      if (existing) {
        const idx = tasks.findIndex(t => t.id === existing.id);
        tasks[idx] = { ...tasks[idx], ...patch, id: tasks[idx].id };
      } else {
        tasks.push({ id: uid(), ...patch });
      }
    });

    onApply({
      ...data,
      regions, teams, tasks,
      taskImportConfig: { colMap, lastImportDate: todayStr(), lastImportFileName: fileName },
    });
  };

  return (
    <Modal title="Import tasks from Excel" wide onClose={onClose}>
      <div style={{ display: "flex", gap: 6, marginBottom: 20 }}>
        {["Upload", "Map columns", "Review"].map((s, i) => (
          <div key={s} style={{
            flex: 1, textAlign: "center", fontSize: 11, padding: "6px 4px", borderRadius: 6,
            background: step === i + 1 ? C.sidebarActiveBg : "transparent",
            color: step >= i + 1 ? C.sidebarActiveText : C.textMute, fontWeight: 700
          }}>{i + 1}. {s}</div>
        ))}
      </div>

      {step === 1 && (
        <div>
          <p style={{ fontSize: 13, color: C.textSoft, marginBottom: 16 }}>
            Upload your task tracker as .csv or .xlsx. For trackers like GSD with a title row above the headers, the real header row (the one with "Name") is detected automatically. Nothing leaves this browser session.
          </p>
          <div
            onClick={() => fileInputRef.current?.click()}
            style={{
              border: `1.5px dashed ${C.borderStrong}`, borderRadius: 10, padding: "32px 20px", textAlign: "center",
              cursor: "pointer", color: C.textSoft
            }}>
            <UploadCloud size={22} style={{ marginBottom: 8 }} />
            <div style={{ fontSize: 13, color: C.text, fontWeight: 600 }}>Click to choose a file</div>
            <div style={{ fontSize: 11.5, marginTop: 4 }}>.csv or .xlsx</div>
          </div>
          <input ref={fileInputRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} style={{ display: "none" }} />
          {error && <div style={{ color: C.red, fontSize: 12.5, marginTop: 10 }}>{error}</div>}
          {data.taskImportConfig.lastImportDate && (
            <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 14 }}>
              Last imported {fmtDate(data.taskImportConfig.lastImportDate)} from {data.taskImportConfig.lastImportFileName}.
            </div>
          )}
        </div>
      )}

      {step === 2 && (
        <div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
            Match each field to a column from <strong style={{ color: C.text }}>{fileName}</strong>. This is remembered for next time.
          </p>
          {TASK_IMPORT_FIELDS.map(f => (
            <div key={f.key} style={{ display: "grid", gridTemplateColumns: "220px 1fr", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <label style={{ fontSize: 12.5, color: C.textSoft, fontWeight: 600 }}>{f.label}{f.required && <span style={{ color: C.red }}> *</span>}</label>
              <select style={{ ...inputStyle, marginBottom: 0 }} value={colMap[f.key] || ""} onChange={e => setColMap({ ...colMap, [f.key]: e.target.value })}>
                <option value="">— Not in file —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          ))}
          <p style={{ fontSize: 11.5, color: C.textMute, marginTop: 4 }}>
            Priority values map to Low/Moderate/High/Critical and Status values map to Not Started/Doing/Follow Up/Done — anything unrecognized defaults to Moderate priority and Not Started.
          </p>
          <div style={{ display: "flex", justifyContent: "space-between", marginTop: 16 }}>
            <button style={ghostBtn} onClick={() => setStep(1)}>Back</button>
            <button style={{ ...primaryBtn, opacity: requiredMapped ? 1 : 0.5 }} disabled={!requiredMapped} onClick={() => setStep(3)}>
              Continue <ArrowRight size={13} style={{ verticalAlign: -2, marginLeft: 4 }} />
            </button>
          </div>
        </div>
      )}

      {step === 3 && diff && (
        <div>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 16 }}>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.green }}>{diff.added}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>New tasks</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.blue }}>{diff.updated}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Updated</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.amber }}>{diff.skipped}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Skipped (no name)</div>
            </div>
          </div>
          {diff.newRepNames.length > 0 && (
            <div style={{ fontSize: 12, color: C.textSoft, marginBottom: 14 }}>
              New owner name{diff.newRepNames.length > 1 ? "s" : ""} found: {diff.newRepNames.join(", ")}. They'll be placed on an "Unassigned" team — reassign them under Settings.
            </div>
          )}
          <div style={{ fontSize: 11.5, color: C.textMute, marginBottom: 14 }}>
            Matching uses task name + account together, so re-importing the same tracker later updates existing tasks instead of duplicating them.
          </div>
          <div style={{ display: "flex", justifyContent: "space-between" }}>
            <button style={ghostBtn} onClick={() => setStep(2)}>Back</button>
            <button style={primaryBtn} onClick={commit}>
              <Check size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Apply import
            </button>
          </div>
        </div>
      )}
    </Modal>
  );
}

// ---------- Shared visual bits ----------
function Card({ children, style }) {
  return <div style={{ background: C.card, border: `1px solid ${C.border}`, borderRadius: 14, padding: 18, ...style }}>{children}</div>;
}

function AlertCard({ label, icon, children, tone }) {
  const toneColor = tone === "red" ? C.red : tone === "amber" ? C.amber : C.green;
  return (
    <div style={{
      flex: 1, minWidth: 200, background: tone === "red" ? "#FFF5F5" : C.card,
      border: `1.5px solid ${tone === "red" ? "#FCA5A5" : C.border}`, borderRadius: 14, padding: "16px 18px"
    }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ fontSize: 11, letterSpacing: 0.5, textTransform: "uppercase", fontWeight: 700, color: C.textMute }}>{label}</span>
        {icon}
      </div>
      {children}
    </div>
  );
}

function DarkCard({ label, value, valueColor, sub, icon }) {
  return (
    <div style={{ flex: 1, minWidth: 200, background: C.cardDark, borderRadius: 14, padding: "16px 18px" }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
        <span style={{ fontSize: 11, letterSpacing: 0.5, textTransform: "uppercase", fontWeight: 700, color: C.onDarkMute }}>{label}</span>
        {icon}
      </div>
      <div style={{ fontSize: 28, fontWeight: 700, color: valueColor || C.onDark, fontVariantNumeric: "tabular-nums" }}>{value}</div>
      {sub && <div style={{ fontSize: 12, color: C.onDarkMute, marginTop: 6 }}>{sub}</div>}
    </div>
  );
}

function Sparkline({ values, color }) {
  const w = 120, h = 30, pad = 3;
  const max = Math.max(1, ...values);
  const pts = values.map((v, i) => {
    const x = pad + (i / (values.length - 1)) * (w - pad * 2);
    const y = h - pad - (v / max) * (h - pad * 2);
    return `${x},${y}`;
  }).join(" ");
  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`}>
      <polyline points={pts} fill="none" stroke={color} strokeWidth="2" strokeLinecap="round" strokeLinejoin="round" />
    </svg>
  );
}

function TrendBadge({ current, previous }) {
  if (!previous) return <span style={{ fontSize: 11, color: C.textMute }}>–</span>;
  const pct = Math.round(((current - previous) / previous) * 100);
  if (pct === 0) return <span style={{ fontSize: 11, color: C.textMute, display: "flex", alignItems: "center", gap: 2 }}><Minus size={11} />0%</span>;
  const up = pct > 0;
  return (
    <span style={{ fontSize: 11, color: up ? C.green : C.red, display: "flex", alignItems: "center", gap: 2, fontWeight: 600 }}>
      {up ? <ArrowUpRight size={11} /> : <ArrowDownRight size={11} />}{Math.abs(pct)}%
    </span>
  );
}

function Pill({ children, color, bg }) {
  return (
    <span style={{
      fontSize: 11, fontWeight: 700, padding: "3px 9px", borderRadius: 999,
      color: color, background: bg, whiteSpace: "nowrap"
    }}>{children}</span>
  );
}

const NAV = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "pipeline", label: "Pipeline", icon: Layers },
  { key: "forecast", label: "Forecast", icon: TrendingUp },
  { key: "accounts", label: "Accounts", icon: Building2 },
  { key: "goals", label: "Goals", icon: Target },
  { key: "prospecting", label: "Prospecting", icon: PhoneCall },
  { key: "tasks", label: "Tasks", icon: CheckSquare },
  { key: "settings", label: "Settings", icon: SettingsIcon },
];

// build refresh marker
export default function SalesCommandCenter() {
  const { data, setData, loaded, saving, lastSyncedAt, loadError, retryLoad } = useStore();
  const [tab, setTab] = useState("dashboard");
  const [period, setPeriod] = useState(currentQuarter());
  const [regionFilter, setRegionFilter] = useState("All");
  const [teamFilter, setTeamFilter] = useState("All");
  const [repFilter, setRepFilter] = useState("All");
  const [includeStretchQ, setIncludeStretchQ] = useState(false);
  const [includeStretchFY, setIncludeStretchFY] = useState(false);
  const [forecastPeriodType, setForecastPeriodType] = useState("quarter");
  const [forecastPeriodValue, setForecastPeriodValue] = useState(currentQuarter());

  const [oppModal, setOppModal] = useState(null);
  const [actModal, setActModal] = useState(null);
  const [taskModal, setTaskModal] = useState(null);
  const [goalModal, setGoalModal] = useState(null);
  const [accModal, setAccModal] = useState(null);
  const [importOpen, setImportOpen] = useState(false);
  const [accountImportOpen, setAccountImportOpen] = useState(false);
  const [taskImportOpen, setTaskImportOpen] = useState(false);
  const [taskFilter, setTaskFilter] = useState("Active");
  const [taskFieldFilters, setTaskFieldFilters] = useState({
    priority: "All", account: "All", cd: "All", owner: "All", sseSme: "All", isOpp: "All", due: "All",
  });
  const setTaskFieldFilter = (key, val) => setTaskFieldFilters(prev => ({ ...prev, [key]: val }));
  const clearTaskFieldFilters = () => { setTaskFilter("Active"); setTaskFieldFilters({ priority: "All", account: "All", cd: "All", owner: "All", sseSme: "All", isOpp: "All", due: "All" }); };
  const [accountFilter, setAccountFilter] = useState("All");
  const [pendingRestore, setPendingRestore] = useState(null);
  const [restoreError, setRestoreError] = useState("");
  const backupFileRef = useRef(null);

  useEffect(() => { setPeriod(currentQuarter()); }, [loaded]);

  const reps = useMemo(() => allRepNames(data), [data]);

  const teamsInRegionFilter = useMemo(() =>
    regionFilter === "All" ? data.teams : data.teams.filter(t => t.regionId === regionFilter),
    [data.teams, regionFilter]);

  const teamsInScope = useMemo(() =>
    teamFilter === "All" ? teamsInRegionFilter : teamsInRegionFilter.filter(t => t.id === teamFilter),
    [teamsInRegionFilter, teamFilter]);

  const repsInScope = useMemo(() => {
    if (regionFilter === "All" && teamFilter === "All") return reps; // includes orphans at the full org level
    return teamsInScope.flatMap(t => t.reps);
  }, [regionFilter, teamFilter, teamsInScope, reps]);

  const repScope = repFilter === "All" ? repsInScope : [repFilter];

  const scopeLabel = useMemo(() => {
    if (repFilter !== "All") return repFilter;
    if (teamFilter !== "All") return data.teams.find(t => t.id === teamFilter)?.name || "Team";
    if (regionFilter !== "All") return (data.regions.find(r => r.id === regionFilter)?.name || "Region") + " region";
    return "All regions (USA)";
  }, [regionFilter, teamFilter, repFilter, data.regions, data.teams]);

  const onRegionChange = (val) => { setRegionFilter(val); setTeamFilter("All"); setRepFilter("All"); };
  const onTeamChange = (val) => { setTeamFilter(val); setRepFilter("All"); };

  const filteredOpps = useMemo(() =>
    data.opportunities.filter(o => repScope.includes(o.rep)), [data.opportunities, repScope]);

  const fullOrgScope = regionFilter === "All" && teamFilter === "All" && repFilter === "All";
  const filteredAccounts = useMemo(() =>
    data.accounts.filter(a => fullOrgScope || repScope.includes(a.rep)), [data.accounts, repScope, fullOrgScope]);

  const accountComputed = useCallback((acc) => {
    const key = normName(acc.name);
    const hasActiveOpp = data.opportunities.some(o => normName(o.account) === key && OPEN_STAGES.includes(o.stage) && o.sfStatus !== "removed");
    const hasMet = data.activities.some(a => normName(a.account) === key && (a.type === "Meeting" || a.type === "Demo"));
    return { hasActiveOpp, hasMet };
  }, [data.opportunities, data.activities]);

  const periodOpts = useMemo(() => quarterOptions(data), [data]);

  const upsert = (listKey, item) => {
    const list = data[listKey];
    const idx = list.findIndex(x => x.id === item.id);
    const next = idx >= 0 ? list.map(x => x.id === item.id ? item : x) : [...list, item];
    setData({ ...data, [listKey]: next });
  };
  const remove = (listKey, id) => {
    setData({ ...data, [listKey]: data[listKey].filter(x => x.id !== id) });
  };

  const exportBackup = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `${(data.orgName || "sales-command-center").replace(/\s+/g, "-").toLowerCase()}-backup-${todayStr()}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleBackupFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRestoreError("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const parsed = JSON.parse(evt.target.result);
        if (!parsed || typeof parsed !== "object") throw new Error("bad shape");
        setPendingRestore(parsed);
      } catch (err) {
        setRestoreError("That file doesn't look like a valid backup from this tool.");
      }
    };
    reader.readAsText(file);
    e.target.value = "";
  };

  const confirmRestore = () => {
    if (!pendingRestore) return;
    setData(mergeLoadedData(pendingRestore));
    setPendingRestore(null);
  };

  const activeOpps = filteredOpps.filter(o => o.sfStatus !== "removed");
  const openOpps = activeOpps.filter(o => OPEN_STAGES.includes(o.stage));
  const openPipelineValue = openOpps.reduce((s, o) => s + oppTotal(o), 0);
  const removedCount = filteredOpps.filter(o => o.sfStatus === "removed").length;

  // ---- YTD / goal-pace computations ----
  const now = new Date();
  const ytdYear = now.getFullYear();
  const annualGoal = data.goals
    .filter(g => g.period.startsWith(ytdYear + "-Q") && repScope.includes(g.rep))
    .reduce((s, g) => s + CATEGORIES.reduce((cs, c) => cs + (Number(g[c.key + "Target"]) || 0), 0), 0);
  const ytdWon = filteredOpps.filter(o => o.stage === "Closed Won" && o.closeDate && o.closeDate.slice(0, 4) === String(ytdYear));
  const ytdActual = ytdWon.reduce((s, o) => s + oppTotal(o), 0);
  const ytdByCategory = CATEGORIES.reduce((acc, c) => {
    acc[c.key] = ytdWon.reduce((s, o) => s + (Number(o[c.key]) || 0), 0);
    return acc;
  }, {});
  const goalRemaining = Math.max(annualGoal - ytdActual, 0);
  const pctToGoal = annualGoal > 0 ? (ytdActual / annualGoal) * 100 : 0;
  const startOfYear = new Date(ytdYear, 0, 1), endOfYear = new Date(ytdYear, 11, 31);
  const fracYear = ((now - startOfYear) / (endOfYear - startOfYear)) * 100;
  const paceStatus = annualGoal === 0 ? "none" : pctToGoal < fracYear * 0.75 ? "red" : pctToGoal < fracYear ? "amber" : "green";
  const paceLabel = paceStatus === "red" ? "CODE RED" : paceStatus === "amber" ? "BEHIND PACE" : paceStatus === "green" ? "ON PACE" : "NO GOAL SET";

  const coverageTarget = data.settings.coverageTarget || 3;
  const coverage = goalRemaining > 0 ? openPipelineValue / goalRemaining : (openPipelineValue > 0 ? coverageTarget : 0);
  const coverageGap = Math.max(coverageTarget * goalRemaining - openPipelineValue, 0);
  const coverageTone = coverage < coverageTarget * 0.5 ? "red" : coverage < coverageTarget ? "amber" : "green";

  const staleDays = data.settings.staleDays || 30;
  const staleOpps = openOpps.filter(o => o.lastStageChange && daysSince(o.lastStageChange) >= staleDays);
  const staleValue = staleOpps.reduce((s, o) => s + oppTotal(o), 0);

  // ---- Forecast ----
  const qLabel = currentQuarter();
  const qOpen = openOpps.filter(o => getQuarter(o.closeDate) === qLabel);
  const qCommit = qOpen.filter(o => (o.forecastCategory || "Pipeline") === "Commit");
  const qStretch = qOpen.filter(o => (o.forecastCategory || "Pipeline") === "Stretch");
  const qForecastValue = qCommit.reduce((s, o) => s + oppTotal(o), 0) + (includeStretchQ ? qStretch.reduce((s, o) => s + oppTotal(o), 0) : 0);
  const qForecastCount = qCommit.length + (includeStretchQ ? qStretch.length : 0);

  const fyOpen = openOpps.filter(o => o.closeDate && o.closeDate.slice(0, 4) === String(ytdYear));
  const fyCommit = fyOpen.filter(o => (o.forecastCategory || "Pipeline") === "Commit");
  const fyStretch = fyOpen.filter(o => (o.forecastCategory || "Pipeline") === "Stretch");
  const fyForecastValue = fyCommit.reduce((s, o) => s + oppTotal(o), 0) + (includeStretchFY ? fyStretch.reduce((s, o) => s + oppTotal(o), 0) : 0);
  const fyForecastCount = fyCommit.length + (includeStretchFY ? fyStretch.length : 0);

  // ---- Monthly forecast (next 4 months) ----
  const months = useMemo(() => [0, 1, 2, 3].map(i => {
    const d = new Date(now.getFullYear(), now.getMonth() + i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    return { key, label: d.toLocaleDateString("en-US", { month: "short" }) + " " + d.getFullYear(), isCurrent: i === 0 };
  }), [now.getMonth(), now.getFullYear()]);

  const monthlyData = months.map(m => {
    const inMonth = openOpps.filter(o => o.closeDate && o.closeDate.slice(0, 7) === m.key);
    const commit = inMonth.filter(o => (o.forecastCategory || "Pipeline") === "Commit");
    const stretch = inMonth.filter(o => (o.forecastCategory || "Pipeline") === "Stretch");
    return {
      ...m,
      commitSum: commit.reduce((s, o) => s + oppTotal(o), 0), commitCount: commit.length,
      stretchSum: stretch.reduce((s, o) => s + oppTotal(o), 0), stretchCount: stretch.length,
      allSum: inMonth.reduce((s, o) => s + oppTotal(o), 0), allCount: inMonth.length,
    };
  });
  const fourMoCommit = monthlyData.reduce((s, m) => s + m.commitSum, 0);
  const fourMoStretch = monthlyData.reduce((s, m) => s + m.stretchSum, 0);
  const fourMoDelta = fourMoStretch - fourMoCommit;

  // ---- Prospecting pulse (last 4 weeks) ----
  const weekBuckets = useMemo(() => {
    const arr = [];
    for (let i = 3; i >= 0; i--) {
      const end = new Date(); end.setHours(23, 59, 59, 999); end.setDate(end.getDate() - i * 7);
      const start = new Date(end); start.setDate(start.getDate() - 6); start.setHours(0, 0, 0, 0);
      arr.push({ start, end });
    }
    return arr;
  }, [now.getDate()]);

  const scopedActivities = data.activities.filter(a => repScope.includes(a.rep));
  const scopedOppsForPulse = data.opportunities.filter(o => repScope.includes(o.rep));
  const inBucket = (dateStr, b) => {
    if (!dateStr) return false;
    const d = new Date(dateStr + "T12:00:00");
    return d >= b.start && d <= b.end;
  };
  const pulse = weekBuckets.map(b => {
    const acts = scopedActivities.filter(a => inBucket(a.date, b));
    const prospectsTouched = new Set(acts.map(a => a.account)).size;
    const emailsSent = acts.filter(a => a.type === "Email").length;
    const pipelineGenerated = scopedOppsForPulse.filter(o => inBucket(o.createdAt, b)).reduce((s, o) => s + oppTotal(o), 0);
    const meetingsDemos = acts.filter(a => a.type === "Meeting" || a.type === "Demo").length;
    return { prospectsTouched, emailsSent, pipelineGenerated, meetingsDemos };
  });
  const pulseHasData = pulse.some(p => p.prospectsTouched || p.emailsSent || p.pipelineGenerated || p.meetingsDemos);

  const next30 = openOpps
    .filter(o => o.closeDate && daysUntil(o.closeDate) >= 0 && daysUntil(o.closeDate) <= 30)
    .sort((a, b) => new Date(a.closeDate) - new Date(b.closeDate));

  const openTasks = data.tasks.filter(t => t.status !== "Done" && repScope.includes(t.rep));
  const upcomingTasks = openTasks.slice().sort((a, b) => (a.dueDate || "9999").localeCompare(b.dueDate || "9999")).slice(0, 5);

  if (!loaded) {
    return <div style={{ color: C.textMute, padding: 40, textAlign: "center" }}>Loading command center…</div>;
  }

  const stageGroups = STAGES
    .map(stage => {
      const opps = filteredOpps.filter(o => o.stage === stage && o.sfStatus !== "removed").sort((a, b) => (a.closeDate || "9999").localeCompare(b.closeDate || "9999"));
      return { key: stage, label: stage, color: STAGE_COLOR[stage], opps, sum: opps.reduce((s, o) => s + oppTotal(o), 0) };
    })
    .filter(g => g.opps.length > 0);
  const allStageSummary = STAGES.map(stage => {
    const opps = filteredOpps.filter(o => o.stage === stage && o.sfStatus !== "removed");
    return { key: stage, label: stage, color: STAGE_COLOR[stage], count: opps.length, sum: opps.reduce((s, o) => s + oppTotal(o), 0) };
  });
  const removedOpps = filteredOpps.filter(o => o.sfStatus === "removed").sort((a, b) => (a.closeDate || "9999").localeCompare(b.closeDate || "9999"));
  if (removedOpps.length > 0) {
    stageGroups.push({ key: "removed", label: "Removed from CRM", color: C.amber, opps: removedOpps, sum: removedOpps.reduce((s, o) => s + oppTotal(o), 0) });
  }

  const currentMonthKey = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}`;
  const monthLabel = (key) => {
    const [y, m] = key.split("-");
    return new Date(Number(y), Number(m) - 1, 1).toLocaleDateString("en-US", { month: "short", year: "numeric" });
  };
  const forecastMonthOpts = Array.from(new Set([
    currentMonthKey, ...openOpps.map(o => o.closeDate ? o.closeDate.slice(0, 7) : null).filter(Boolean),
  ])).sort();
  const forecastQuarterOpts = Array.from(new Set([
    currentQuarter(), ...openOpps.map(o => getQuarter(o.closeDate)).filter(Boolean),
  ])).sort();
  const forecastYearOpts = Array.from(new Set([
    String(ytdYear), ...openOpps.map(o => o.closeDate ? o.closeDate.slice(0, 4) : null).filter(Boolean),
  ])).sort();
  const forecastPeriodOpts = forecastPeriodType === "month" ? forecastMonthOpts : forecastPeriodType === "quarter" ? forecastQuarterOpts : forecastYearOpts;
  const forecastPeriodLabel = forecastPeriodType === "month" ? monthLabel(forecastPeriodValue)
    : forecastPeriodType === "quarter" ? forecastPeriodValue : `FY ${forecastPeriodValue}`;
  const changeForecastPeriodType = (type) => {
    setForecastPeriodType(type);
    setForecastPeriodValue(type === "month" ? currentMonthKey : type === "quarter" ? currentQuarter() : String(ytdYear));
  };
  const inForecastPeriod = (o) => {
    if (!o.closeDate) return false;
    if (forecastPeriodType === "month") return o.closeDate.slice(0, 7) === forecastPeriodValue;
    if (forecastPeriodType === "quarter") return getQuarter(o.closeDate) === forecastPeriodValue;
    return o.closeDate.slice(0, 4) === forecastPeriodValue;
  };
  const openOppsInForecastPeriod = openOpps.filter(inForecastPeriod);

  const forecastGroups = FORECAST_CATS.map(cat => {
    const opps = openOppsInForecastPeriod.filter(o => (o.forecastCategory || "Pipeline") === cat).sort((a, b) => (a.closeDate || "9999").localeCompare(b.closeDate || "9999"));
    return { key: cat, label: cat, color: FORECAST_COLOR[cat], opps, sum: opps.reduce((s, o) => s + oppTotal(o), 0) };
  });

  const byPriority = (list) => list.slice().sort((a, b) =>
    (PRIORITY_RANK[a.priority] ?? 4) - (PRIORITY_RANK[b.priority] ?? 4) || (a.dueDate || "9999").localeCompare(b.dueDate || "9999"));
  const earliestDue = (list) => list.reduce((min, t) => (!t.dueDate ? min : (min === null || t.dueDate < min ? t.dueDate : min)), null);

  const repScopedTasks = data.tasks.filter(t => repScope.includes(t.rep));
  const distinctTaskVals = (key) => Array.from(new Set(repScopedTasks.map(t => (t[key] || "").toString().trim()).filter(Boolean))).sort();
  const taskAccountOpts = distinctTaskVals("account");
  const taskCdOpts = distinctTaskVals("cd");
  const taskSseOpts = distinctTaskVals("sseSme");

  const scopedTasks = repScopedTasks
    .filter(t => taskFilter === "All" ? true : taskFilter === "Active" ? t.status !== "Done" : t.status === taskFilter)
    .filter(t => taskFieldFilters.priority === "All" || t.priority === taskFieldFilters.priority)
    .filter(t => taskFieldFilters.account === "All" || normName(t.account) === normName(taskFieldFilters.account))
    .filter(t => taskFieldFilters.cd === "All" || (t.cd || "") === taskFieldFilters.cd)
    .filter(t => taskFieldFilters.owner === "All" || t.rep === taskFieldFilters.owner)
    .filter(t => taskFieldFilters.sseSme === "All" || (t.sseSme || "") === taskFieldFilters.sseSme)
    .filter(t => taskFieldFilters.isOpp === "All" ? true : taskFieldFilters.isOpp === "Yes" ? !!t.isOpp : !t.isOpp)
    .filter(t => {
      if (taskFieldFilters.due === "All") return true;
      if (taskFieldFilters.due === "No due date") return !t.dueDate;
      if (!t.dueDate) return false;
      const d = daysUntil(t.dueDate);
      if (taskFieldFilters.due === "Overdue") return d < 0;
      if (taskFieldFilters.due === "Next 7 days") return d >= 0 && d <= 7;
      if (taskFieldFilters.due === "Next 30 days") return d >= 0 && d <= 30;
      return true;
    });
  const unassignedTaskGroup = (() => {
    const list = byPriority(scopedTasks.filter(t => !t.account || !t.account.trim()));
    return list.length > 0 ? { key: "unassigned", title: "Unassigned tasks", subtitle: "No account linked", tasks: list } : null;
  })();
  const acctTaskMap = new Map();
  scopedTasks.filter(t => t.account && t.account.trim()).forEach(t => {
    const key = normName(t.account);
    if (!acctTaskMap.has(key)) {
      const accRec = data.accounts.find(a => normName(a.name) === key);
      acctTaskMap.set(key, { key, title: t.account.trim(), subtitle: accRec?.cd ? `CD: ${accRec.cd}` : "Account", tasks: [] });
    }
    acctTaskMap.get(key).tasks.push(t);
  });
  const acctTaskGroups = Array.from(acctTaskMap.values())
    .map(g => ({ ...g, tasks: byPriority(g.tasks) }))
    .sort((a, b) => {
      const ap = PRIORITY_RANK[a.tasks[0]?.priority] ?? 4;
      const bp = PRIORITY_RANK[b.tasks[0]?.priority] ?? 4;
      if (ap !== bp) return ap - bp;
      const ad = earliestDue(a.tasks), bd = earliestDue(b.tasks);
      if (ad && bd) return ad.localeCompare(bd);
      if (ad) return -1;
      if (bd) return 1;
      return a.title.localeCompare(b.title);
    });
  const taskGroups = unassignedTaskGroup ? [unassignedTaskGroup, ...acctTaskGroups] : acctTaskGroups;

  return (
    <div style={{
      fontFamily: "'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      background: C.page, color: C.text, borderRadius: 14, overflow: "hidden",
      border: `1px solid ${C.border}`, minHeight: 600
    }}>
      {loadError && (
        <div style={{
          background: "#FEF2F2", borderBottom: `2px solid ${C.red}`, padding: "12px 22px",
          display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap"
        }}>
          <AlertCircle size={18} color={C.red} style={{ flexShrink: 0 }} />
          <div style={{ flex: 1, minWidth: 240 }}>
            <div style={{ fontSize: 13, fontWeight: 700, color: C.red }}>Couldn't confirm your saved data loaded</div>
            <div style={{ fontSize: 12, color: "#7F1D1D", marginTop: 2 }}>
              This may be a connection issue, not lost data — your team's real data could still be safe in storage.
              <strong> Don't add or edit anything until this is resolved</strong>, or you risk overwriting it. Try refreshing the page, or click Retry below.
            </div>
          </div>
          <button style={{ ...primaryBtn, background: C.red, flexShrink: 0 }} onClick={retryLoad}>
            <RefreshCw size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Retry now
          </button>
        </div>
      )}
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        padding: "16px 22px", borderBottom: `1px solid ${C.border}`, background: C.card, flexWrap: "wrap", gap: 12
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{
            width: 30, height: 30, borderRadius: 8, background: C.cardDark,
            display: "flex", alignItems: "center", justifyContent: "center", fontWeight: 800, fontSize: 13, color: "#FFFFFF"
          }}>{data.orgName.slice(0, 1)}</div>
          <div>
            <div style={{ fontSize: 14, fontWeight: 700, letterSpacing: 0.1, display: "flex", alignItems: "center", gap: 8 }}>
              {data.orgName} command center
              <Pill color={C.sidebarActiveText} bg={C.sidebarActiveBg}>{scopeLabel}</Pill>
            </div>
            <div style={{ fontSize: 11, color: C.textMute, display: "flex", alignItems: "center", gap: 8, marginTop: 2 }}>
              <span style={{ display: "flex", alignItems: "center", gap: 5 }}>
                <span style={{ width: 6, height: 6, borderRadius: "50%", background: saving ? C.amber : C.green, display: "inline-block", flexShrink: 0 }} />
                {saving ? "Saving…" : lastSyncedAt
                  ? `Live · synced ${(() => { const s = Math.max(0, Math.round((Date.now() - lastSyncedAt) / 1000)); return s < 60 ? s + "s" : Math.round(s / 60) + "m"; })()} ago`
                  : "Connecting…"}
              </span>
              {data.importConfig.lastImportDate && <span>· Last Salesforce sync {fmtDate(data.importConfig.lastImportDate)}</span>}
            </div>
          </div>
        </div>
        <div style={{ display: "flex", gap: 10, alignItems: "center", flexWrap: "wrap" }}>
          <button onClick={() => setImportOpen(true)} style={{ ...ghostBtn, display: "flex", alignItems: "center", gap: 6, padding: "7px 12px" }}>
            <RefreshCw size={13} />Sync from Salesforce
          </button>
          <select value={regionFilter} onChange={e => onRegionChange(e.target.value)} style={{
            background: "#F8FAFC", border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", fontSize: 12.5
          }}>
            <option value="All">All regions</option>
            {data.regions.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
          </select>
          <select value={teamFilter} onChange={e => onTeamChange(e.target.value)} style={{
            background: "#F8FAFC", border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", fontSize: 12.5
          }}>
            <option value="All">All teams</option>
            {teamsInRegionFilter.map(t => <option key={t.id} value={t.id}>{t.name}</option>)}
          </select>
          <select value={repFilter} onChange={e => setRepFilter(e.target.value)} style={{
            background: "#F8FAFC", border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", fontSize: 12.5
          }}>
            <option value="All">All reps</option>
            {repsInScope.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
          <select value={period} onChange={e => setPeriod(e.target.value)} style={{
            background: "#F8FAFC", border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", fontSize: 12.5
          }}>
            {periodOpts.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </div>
      </div>

      <div style={{ display: "flex" }}>
        <div style={{ width: 176, borderRight: `1px solid ${C.border}`, padding: "16px 10px", flexShrink: 0, background: C.card }}>
          {NAV.map(n => {
            const Icon = n.icon;
            const active = tab === n.key;
            return (
              <div key={n.key} onClick={() => setTab(n.key)} style={{
                display: "flex", alignItems: "center", gap: 10, padding: "9px 12px", borderRadius: 8,
                cursor: "pointer", marginBottom: 3, fontSize: 13, fontWeight: active ? 700 : 500,
                background: active ? C.sidebarActiveBg : "transparent", color: active ? C.sidebarActiveText : C.textSoft
              }}>
                <Icon size={16} />{n.label}
              </div>
            );
          })}
        </div>

        <div style={{ flex: 1, padding: 22, minWidth: 0, background: C.page }}>

          {tab === "dashboard" && (
            <div>
              {/* Alert row */}
              <div style={{ display: "flex", gap: 14, marginBottom: 16, flexWrap: "wrap" }}>
                <AlertCard label="YTD percentage to goal" tone={paceStatus === "red" ? "red" : paceStatus === "amber" ? "amber" : "green"}
                  icon={<AlertCircle size={16} color={paceStatus === "red" ? C.red : C.textMute} />}>
                  <div style={{ fontSize: 32, fontWeight: 800, color: paceStatus === "red" ? C.red : C.text }}>{pctToGoal.toFixed(1)}%</div>
                  <div style={{ height: 6, background: "#F1F5F9", borderRadius: 4, overflow: "hidden", margin: "8px 0" }}>
                    <div style={{ height: "100%", width: `${Math.min(pctToGoal, 100)}%`, background: paceStatus === "red" ? C.red : paceStatus === "amber" ? C.amber : C.green, borderRadius: 4 }} />
                  </div>
                  <span style={{
                    fontSize: 11, fontWeight: 800, letterSpacing: 0.4,
                    color: paceStatus === "red" ? C.red : paceStatus === "amber" ? C.amber : C.green
                  }}>{paceLabel}</span>
                </AlertCard>

                <DarkCard label="YTD total margin" value={fmtMoneyShort(ytdActual)}
                  sub={`Goal remaining · ${fmtMoneyShort(goalRemaining)}`} icon={<DollarSign size={15} color={C.onDarkMute} />} />

                <DarkCard label="Stale opportunities" value={staleOpps.length}
                  valueColor={staleOpps.length === 0 ? C.green : C.red}
                  sub={`No stage change in ${staleDays}+ days · ${fmtMoneyShort(staleValue)} at risk`}
                  icon={staleOpps.length === 0 ? <CheckCircle2 size={16} color={C.green} /> : <AlertCircle size={16} color={C.red} />} />
              </div>

              {/* YTD margin breakdown */}
              <Card style={{ marginBottom: 16 }}>
                <h4 style={{ margin: "0 0 12px", fontSize: 12.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textSoft }}>YTD margin breakdown</h4>
                <div style={{ display: "flex", height: 40, borderRadius: 8, overflow: "hidden", gap: 2, marginBottom: 8 }}>
                  {CATEGORIES.map(c => {
                    const val = ytdByCategory[c.key];
                    const pct = ytdActual > 0 ? (val / ytdActual) * 100 : 100 / CATEGORIES.length;
                    return (
                      <div key={c.key} style={{ width: `${pct}%`, background: c.color, display: "flex", flexDirection: "column", justifyContent: "center", alignItems: "center", color: "#fff", fontSize: 10, minWidth: 60 }}>
                        <span style={{ fontWeight: 700, fontSize: 10, textTransform: "uppercase", opacity: 0.85 }}>{c.label}</span>
                        <span style={{ fontWeight: 700 }}>{fmtMoneyShort(val)}</span>
                      </div>
                    );
                  })}
                </div>
                <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11, color: C.textMute, marginBottom: 6 }}>
                  <span>0%</span><span>50%</span><span>100%</span>
                </div>
                <div style={{ display: "flex", gap: 16, fontSize: 12, color: C.textSoft }}>
                  {CATEGORIES.map(c => (
                    <span key={c.key} style={{ display: "flex", alignItems: "center", gap: 5 }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: c.color, display: "inline-block" }} />
                      {c.label} {fmtMoneyShort(ytdByCategory[c.key])} ({ytdActual > 0 ? ((ytdByCategory[c.key] / ytdActual) * 100).toFixed(1) : "0.0"}%)
                    </span>
                  ))}
                </div>
              </Card>

              {/* Pipeline margin + coverage */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
                <Card>
                  <h4 style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>YTD total pipeline margin</h4>
                  <div style={{ fontSize: 30, fontWeight: 800, color: C.text }}>{fmtMoneyShort(openPipelineValue)}</div>
                  <div style={{ fontSize: 12, color: C.textMute, marginTop: 6 }}>Unweighted pipeline margin · {openOpps.length} deals</div>
                </Card>
                <Card>
                  <h4 style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>
                    Pipeline coverage (goal to {coverageTarget}x)
                  </h4>
                  <div style={{ fontSize: 30, fontWeight: 800, color: coverageTone === "red" ? C.red : coverageTone === "amber" ? C.amber : C.green }}>
                    {coverage.toFixed(1)}x <span style={{ fontSize: 13, fontWeight: 600, color: C.textMute }}>Target: {coverageTarget}.0x</span>
                  </div>
                  <div style={{ height: 6, background: "#F1F5F9", borderRadius: 4, overflow: "hidden", margin: "10px 0 6px" }}>
                    <div style={{ height: "100%", width: `${Math.min((coverage / coverageTarget) * 100, 100)}%`, background: coverageTone === "red" ? C.red : coverageTone === "amber" ? C.amber : C.green, borderRadius: 4 }} />
                  </div>
                  {coverageGap > 0 && <div style={{ fontSize: 12, color: C.textMute }}>Gap: {fmtMoneyShort(coverageGap)} to {coverageTarget}x</div>}
                </Card>
              </div>

              {/* Forecast cards */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16, marginBottom: 16 }}>
                <Card>
                  <h4 style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>{qLabel} forecast</h4>
                  <div style={{ fontSize: 28, fontWeight: 800, color: C.purple }}>{fmtMoneyShort(qForecastValue)}</div>
                  <div style={{ fontSize: 12, color: C.textMute, margin: "6px 0 10px" }}>{qForecastCount} deals · {includeStretchQ ? "Commit + stretch" : "Commit"}</div>
                  <button style={{ ...ghostBtn, padding: "5px 10px", fontSize: 12 }} onClick={() => setIncludeStretchQ(v => !v)}>
                    {includeStretchQ ? "− Exclude stretch" : "+ Include stretch"}
                  </button>
                </Card>
                <Card>
                  <h4 style={{ margin: "0 0 10px", fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>FY{String(ytdYear).slice(2)} forecast</h4>
                  <div style={{ fontSize: 28, fontWeight: 800, color: C.purple }}>{fmtMoneyShort(fyForecastValue)}</div>
                  <div style={{ fontSize: 12, color: C.textMute, margin: "6px 0 10px" }}>{fyForecastCount} deals · Pipeline through {includeStretchFY ? "stretch" : "commit"}</div>
                  <button style={{ ...ghostBtn, padding: "5px 10px", fontSize: 12 }} onClick={() => setIncludeStretchFY(v => !v)}>
                    {includeStretchFY ? "− Exclude stretch" : "+ Include stretch"}
                  </button>
                </Card>
              </div>

              {/* Prospecting pulse */}
              <Card style={{ marginBottom: 16 }}>
                <div style={{ display: "flex", alignItems: "baseline", justifyContent: "space-between", marginBottom: 4 }}>
                  <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Prospecting pulse</h4>
                  <span style={{ fontSize: 11.5, color: C.textMute }}>Activity trends · last 4 weeks</span>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginTop: 14 }}>
                  {[
                    { key: "prospectsTouched", label: "Prospects touched", color: C.blue },
                    { key: "emailsSent", label: "Emails sent", color: C.purple },
                    { key: "pipelineGenerated", label: "Pipeline generated", color: C.green, money: true },
                    { key: "meetingsDemos", label: "Meetings & demos", color: C.amber },
                  ].map(m => {
                    const values = pulse.map(p => p[m.key]);
                    const current = values[3], prev = values[2];
                    return (
                      <div key={m.key} style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 12 }}>
                        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 4 }}>
                          <span style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>{m.label}</span>
                          <TrendBadge current={current} previous={prev} />
                        </div>
                        <div style={{ fontSize: 20, fontWeight: 800, color: C.text, marginBottom: 6 }}>{m.money ? fmtMoneyShort(current) : current}</div>
                        <Sparkline values={values} color={m.color} />
                        <div style={{ display: "flex", justifyContent: "space-between", fontSize: 9.5, color: C.textMute, marginTop: 2 }}>
                          <span>Wk 1</span><span>Wk 2</span><span>Wk 3</span><span>Wk 4</span>
                        </div>
                      </div>
                    );
                  })}
                </div>
                {!pulseHasData && <div style={{ fontSize: 12, color: C.textMute, marginTop: 10, textAlign: "center" }}>No pulse data yet — this fills in automatically as you log prospecting activity.</div>}
              </Card>

              {/* Monthly forecast */}
              <Card style={{ marginBottom: 16, padding: 0, overflow: "hidden" }}>
                <div style={{ padding: "16px 18px 4px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
                  <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Monthly forecast</h4>
                  <div style={{ display: "flex", gap: 6 }}>
                    <Pill color="#166534" bg="#DCFCE7">Commit</Pill>
                    <Pill color="#92400E" bg="#FEF3C7">Stretch</Pill>
                  </div>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: `repeat(${monthlyData.length}, 1fr)`, gap: 0, padding: "12px 18px" }}>
                  {monthlyData.map(m => (
                    <div key={m.key} style={{
                      borderRight: `1px solid ${C.border}`, padding: "10px 14px", position: "relative"
                    }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 10 }}>
                        <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{m.label}</span>
                        {m.isCurrent && <Pill color={C.sidebarActiveText} bg={C.sidebarActiveBg}>Current</Pill>}
                      </div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, color: C.green, textTransform: "uppercase", letterSpacing: 0.3 }}>Commit</div>
                      <div style={{ fontSize: 17, fontWeight: 800, color: C.text }}>{fmtMoneyShort(m.commitSum)}</div>
                      <div style={{ fontSize: 11, color: C.textMute, marginBottom: 8 }}>{m.commitCount} deals</div>
                      <div style={{ fontSize: 10.5, fontWeight: 700, color: C.amber, textTransform: "uppercase", letterSpacing: 0.3 }}>Stretch</div>
                      <div style={{ fontSize: 17, fontWeight: 800, color: C.text }}>{fmtMoneyShort(m.stretchSum)}</div>
                      <div style={{ fontSize: 11, color: C.textMute, marginBottom: 8 }}>{m.stretchCount} deals</div>
                      <div style={{ fontSize: 10.5, color: C.textMute, borderTop: `1px solid ${C.border}`, paddingTop: 6 }}>
                        All pipeline <strong style={{ color: C.textSoft }}>{fmtMoneyShort(m.allSum)}</strong> · {m.allCount} deals
                      </div>
                    </div>
                  ))}
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", background: "#F8FAFC", borderTop: `1px solid ${C.border}`, padding: "14px 18px" }}>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>4-month commit</div>
                    <div style={{ fontSize: 17, fontWeight: 800, color: C.green }}>{fmtMoneyShort(fourMoCommit)}</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>4-month stretch</div>
                    <div style={{ fontSize: 17, fontWeight: 800, color: C.amber }}>{fmtMoneyShort(fourMoStretch)}</div>
                  </div>
                  <div style={{ textAlign: "center" }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>Delta</div>
                    <div style={{ fontSize: 17, fontWeight: 800, color: C.purple }}>{fmtMoneyShort(fourMoDelta)}</div>
                  </div>
                </div>
              </Card>

              {/* Closing soon + tasks */}
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
                <Card>
                  <h4 style={{ margin: "0 0 12px", fontSize: 13, fontWeight: 700, color: C.text }}>Closing in 30 days</h4>
                  {next30.length === 0 && <div style={{ fontSize: 12, color: C.textMute }}>Nothing on the horizon.</div>}
                  {next30.map(o => (
                    <div key={o.id} onClick={() => setOppModal(o)} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${C.border}`, cursor: "pointer", fontSize: 12.5 }}>
                      <span style={{ color: C.text }}>{o.name}</span>
                      <span style={{ color: C.textMute }}>{fmtDate(o.closeDate)}</span>
                    </div>
                  ))}
                </Card>
                <Card>
                  <h4 style={{ margin: "0 0 12px", fontSize: 13, fontWeight: 700, color: C.text }}>Upcoming tasks</h4>
                  {upcomingTasks.length === 0 && <div style={{ fontSize: 12, color: C.textMute }}>All caught up.</div>}
                  {upcomingTasks.map(t => (
                    <div key={t.id} onClick={() => setTaskModal(t)} style={{ display: "flex", justifyContent: "space-between", padding: "7px 0", borderBottom: `1px solid ${C.border}`, cursor: "pointer", fontSize: 12.5 }}>
                      <span style={{ color: C.text }}>{t.title}</span>
                      <span style={{ color: t.dueDate && daysUntil(t.dueDate) < 0 ? C.red : C.textMute }}>{fmtDate(t.dueDate)}</span>
                    </div>
                  ))}
                </Card>
              </div>
            </div>
          )}

          {tab === "pipeline" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Revenue by stage</h4>
                <button style={primaryBtn} onClick={() => setOppModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add opportunity</button>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 14, marginBottom: 22 }}>
                {allStageSummary.map(g => (
                  <Card key={g.key} style={{ padding: "14px 18px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 6 }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                      <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>{g.label}</span>
                    </div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(g.sum)}</div>
                    <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 2 }}>{g.count} {g.count === 1 ? "deal" : "deals"}</div>
                  </Card>
                ))}
              </div>
              {stageGroups.length === 0 && (
                <Card style={{ padding: "20px", textAlign: "center", color: C.textMute, fontSize: 13, marginBottom: 22 }}>No opportunities yet. Add one or sync from Salesforce.</Card>
              )}

              {stageGroups.map(g => (
                <div key={g.key} style={{ marginBottom: 22 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                    <h4 style={{ margin: 0, fontSize: 13, fontWeight: 700, color: C.text }}>{g.label}</h4>
                    <span style={{ fontSize: 11.5, color: C.textMute }}>{g.opps.length} {g.opps.length === 1 ? "deal" : "deals"} · {fmtMoneyShort(g.sum)}</span>
                  </div>
                  <Card style={{ padding: 0, overflow: "hidden" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.8fr 0.8fr 0.8fr 0.8fr 0.7fr", padding: "9px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                      <span>Opportunity</span><span>Account</span><span>Product</span><span>Services</span><span>Managed</span><span>Close date</span><span>Forecast</span>
                    </div>
                    {g.opps.map(o => (
                      <div key={o.id} onClick={() => setOppModal(o)} style={{
                        display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.8fr 0.8fr 0.8fr 0.8fr 0.7fr", padding: "11px 16px",
                        fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center",
                        opacity: o.sfStatus === "removed" ? 0.55 : 1
                      }}>
                        <span style={{ color: C.text, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                          {o.name}{o.sfId && <FileSpreadsheet size={11} color={C.textMute} />}
                        </span>
                        <span style={{ color: C.textSoft }}>{o.account}</span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.productMargin)}</span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.servicesMargin)}</span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.managedServices)}</span>
                        <span style={{ color: o.closeDate && daysUntil(o.closeDate) < 0 && OPEN_STAGES.includes(o.stage) ? C.red : C.textSoft }}>{fmtDate(o.closeDate)}</span>
                        <span>{OPEN_STAGES.includes(o.stage) && <Pill color={FORECAST_COLOR[o.forecastCategory || "Pipeline"]} bg={FORECAST_COLOR[o.forecastCategory || "Pipeline"] + "1A"}>{o.forecastCategory || "Pipeline"}</Pill>}</span>
                      </div>
                    ))}
                  </Card>
                </div>
              ))}
            </div>
          )}

          {tab === "forecast" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, flexWrap: "wrap", gap: 10 }}>
                <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Forecast totals — {forecastPeriodLabel}</h4>
                <button style={primaryBtn} onClick={() => setOppModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add opportunity</button>
              </div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", gap: 8 }}>
                  {[["month", "Month"], ["quarter", "Quarter"], ["year", "Fiscal year"]].map(([type, label]) => (
                    <div key={type} onClick={() => changeForecastPeriodType(type)} style={{
                      fontSize: 12, padding: "6px 12px", borderRadius: 999, cursor: "pointer", fontWeight: 600,
                      background: forecastPeriodType === type ? C.sidebarActiveBg : C.card,
                      color: forecastPeriodType === type ? C.sidebarActiveText : C.textSoft, border: `1px solid ${forecastPeriodType === type ? "#C7D2FE" : C.border}`
                    }}>{label}</div>
                  ))}
                </div>
                <select value={forecastPeriodValue} onChange={e => setForecastPeriodValue(e.target.value)} style={{
                  background: C.card, border: `1px solid ${C.border}`, color: C.text, borderRadius: 8, padding: "7px 10px", fontSize: 12.5
                }}>
                  {forecastPeriodOpts.map(p => (
                    <option key={p} value={p}>{forecastPeriodType === "month" ? monthLabel(p) : forecastPeriodType === "year" ? `FY ${p}` : p}</option>
                  ))}
                </select>
              </div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, marginBottom: 22 }}>
                {forecastGroups.map(g => (
                  <Card key={g.key} style={{ padding: "14px 18px" }}>
                    <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 6 }}>
                      <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                      <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>{g.label}</span>
                    </div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(g.sum)}</div>
                    <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 2 }}>{g.opps.length} open {g.opps.length === 1 ? "deal" : "deals"}</div>
                  </Card>
                ))}
              </div>

              {forecastGroups.map(g => (
                <div key={g.key} style={{ marginBottom: 22 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                    <h4 style={{ margin: 0, fontSize: 13, fontWeight: 700, color: C.text }}>{g.label}</h4>
                    <span style={{ fontSize: 11.5, color: C.textMute }}>{g.opps.length} {g.opps.length === 1 ? "deal" : "deals"} · {fmtMoneyShort(g.sum)}</span>
                  </div>
                  <Card style={{ padding: 0, overflow: "hidden" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.9fr 0.8fr 0.8fr 0.8fr 0.8fr", padding: "9px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                      <span>Opportunity</span><span>Account</span><span>Stage</span><span>Product</span><span>Services</span><span>Managed</span><span>Close date</span>
                    </div>
                    {g.opps.map(o => (
                      <div key={o.id} onClick={() => setOppModal(o)} style={{
                        display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.9fr 0.8fr 0.8fr 0.8fr 0.8fr", padding: "11px 16px",
                        fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center"
                      }}>
                        <span style={{ color: C.text, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                          {o.name}{o.sfId && <FileSpreadsheet size={11} color={C.textMute} />}
                        </span>
                        <span style={{ color: C.textSoft }}>{o.account}</span>
                        <span><Pill color={STAGE_COLOR[o.stage]} bg={STAGE_COLOR[o.stage] + "1A"}>{o.stage}</Pill></span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.productMargin)}</span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.servicesMargin)}</span>
                        <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(o.managedServices)}</span>
                        <span style={{ color: o.closeDate && daysUntil(o.closeDate) < 0 ? C.red : C.textSoft }}>{fmtDate(o.closeDate)}</span>
                      </div>
                    ))}
                    {g.opps.length === 0 && (
                      <div style={{ padding: 20, textAlign: "center", color: C.textMute, fontSize: 12.5 }}>No open deals in this category.</div>
                    )}
                  </Card>
                </div>
              ))}
            </div>
          )}

          {tab === "accounts" && (
            <div>
              <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12, marginBottom: 16 }}>
                {[
                  { label: "Total accounts", value: filteredAccounts.length, color: C.text },
                  { label: "Current customers", value: filteredAccounts.filter(a => a.currentCustomer).length, color: C.blue },
                  { label: "Ahead strategic", value: filteredAccounts.filter(a => a.strategic).length, color: C.purple },
                  { label: "MD priority", value: filteredAccounts.filter(a => a.mdPriority).length, color: C.amber },
                  { label: "Active opportunity", value: filteredAccounts.filter(a => accountComputed(a).hasActiveOpp).length, color: C.green },
                ].map(s => (
                  <Card key={s.label} style={{ padding: "12px 14px" }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>{s.label}</div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: s.color }}>{s.value}</div>
                  </Card>
                ))}
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16, flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {["All", "Active opportunity", "Current customer", "Strategic", "MD priority", "Never met"].map(s => (
                    <div key={s} onClick={() => setAccountFilter(s)} style={{
                      fontSize: 12, padding: "6px 12px", borderRadius: 999, cursor: "pointer", fontWeight: 600,
                      background: accountFilter === s ? C.sidebarActiveBg : C.card,
                      color: accountFilter === s ? C.sidebarActiveText : C.textSoft, border: `1px solid ${accountFilter === s ? "#C7D2FE" : C.border}`
                    }}>{s}</div>
                  ))}
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={ghostBtn} onClick={() => setAccountImportOpen(true)}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Import from Excel</button>
                  <button style={primaryBtn} onClick={() => setAccModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add account</button>
                </div>
              </div>

              <Card style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr 1fr 1fr 1fr 1fr", padding: "10px 16px", fontSize: 11, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                  <span>Account</span><span>CD</span><span>Active opp</span><span>Customer</span><span>Met</span><span>MD priority</span><span>Strategic</span>
                </div>
                {filteredAccounts
                  .filter(a => {
                    if (accountFilter === "All") return true;
                    const c = accountComputed(a);
                    if (accountFilter === "Active opportunity") return c.hasActiveOpp;
                    if (accountFilter === "Current customer") return a.currentCustomer;
                    if (accountFilter === "Strategic") return a.strategic;
                    if (accountFilter === "MD priority") return a.mdPriority;
                    if (accountFilter === "Never met") return !c.hasMet;
                    return true;
                  })
                  .sort((a, b) => a.name.localeCompare(b.name))
                  .map(a => {
                    const c = accountComputed(a);
                    const yn = (v) => v
                      ? <Pill color={C.green} bg="#DCFCE7"><Check size={10} style={{ verticalAlign: -1 }} /></Pill>
                      : <Pill color={C.textMute} bg="#F1F5F9">—</Pill>;
                    return (
                      <div key={a.id} onClick={() => setAccModal(a)} style={{
                        display: "grid", gridTemplateColumns: "1.6fr 1fr 1fr 1fr 1fr 1fr 1fr", padding: "12px 16px",
                        fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center"
                      }}>
                        <span style={{ color: C.text, fontWeight: 600 }}>{a.name}</span>
                        <span style={{ color: C.textSoft }}>{a.cd || "—"}</span>
                        <span>{yn(c.hasActiveOpp)}</span>
                        <span>{yn(a.currentCustomer)}</span>
                        <span>{yn(c.hasMet)}</span>
                        <span>{yn(a.mdPriority)}</span>
                        <span>{yn(a.strategic)}</span>
                      </div>
                    );
                  })}
                {filteredAccounts.length === 0 && (
                  <div style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No accounts yet. Import your account list or add one manually.</div>
                )}
              </Card>
            </div>
          )}

          {tab === "goals" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Goals by rep and period</h4>
                <button style={primaryBtn} onClick={() => setGoalModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Set a goal</button>
              </div>
              <Card style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", padding: "10px 16px", fontSize: 11, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                  <span>Period</span><span>Rep</span><span>Product target</span><span>Services target</span><span>Managed target</span>
                </div>
                {data.goals.slice().sort((a, b) => b.period.localeCompare(a.period)).map(g => (
                  <div key={g.id} onClick={() => setGoalModal(g)} style={{
                    display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", padding: "12px 16px",
                    fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", color: C.textSoft
                  }}>
                    <span style={{ color: C.text, fontWeight: 600 }}>{g.period}</span>
                    <span>{g.rep}</span>
                    <span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtMoney(g.productMarginTarget)}</span>
                    <span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtMoney(g.servicesMarginTarget)}</span>
                    <span style={{ fontVariantNumeric: "tabular-nums" }}>{fmtMoney(g.managedServicesTarget)}</span>
                  </div>
                ))}
                {data.goals.length === 0 && (
                  <div style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No goals set yet. Set targets for each rep by quarter.</div>
                )}
              </Card>
            </div>
          )}

          {tab === "prospecting" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Prospecting activity log</h4>
                <button style={primaryBtn} onClick={() => setActModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Log activity</button>
              </div>
              <Card style={{ padding: 0, overflow: "hidden" }}>
                <div style={{ display: "grid", gridTemplateColumns: "0.8fr 0.8fr 1.2fr 1fr 1fr 1.6fr", padding: "10px 16px", fontSize: 11, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                  <span>Date</span><span>Type</span><span>Account</span><span>Contact</span><span>Rep</span><span>Notes</span>
                </div>
                {data.activities.filter(a => repScope.includes(a.rep)).slice().sort((a, b) => b.date.localeCompare(a.date)).map(a => (
                  <div key={a.id} onClick={() => setActModal(a)} style={{
                    display: "grid", gridTemplateColumns: "0.8fr 0.8fr 1.2fr 1fr 1fr 1.6fr", padding: "12px 16px",
                    fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", color: C.textSoft, alignItems: "center"
                  }}>
                    <span>{fmtDate(a.date)}</span>
                    <span><Pill color={C.blue} bg="#DBEAFE">{a.type}</Pill></span>
                    <span style={{ color: C.text }}>{a.account}</span>
                    <span>{a.contact || "—"}</span>
                    <span>{a.rep}</span>
                    <span style={{ color: C.textMute, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.notes}</span>
                  </div>
                ))}
                {data.activities.length === 0 && (
                  <div style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No activity logged yet.</div>
                )}
              </Card>
            </div>
          )}

          {tab === "tasks" && (
            <div>
              <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 12 }}>
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={ghostBtn} onClick={() => setTaskImportOpen(true)}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Import from Excel</button>
                  <button style={primaryBtn} onClick={() => setTaskModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add task</button>
                </div>
              </div>

              <Card style={{ padding: "14px 16px", marginBottom: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 10 }}>
                  <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>Filter by field</span>
                  <button style={{ ...ghostBtn, padding: "4px 9px", fontSize: 11.5 }} onClick={clearTaskFieldFilters}>Clear all filters</button>
                </div>
                <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fill, minmax(150px, 1fr))", gap: 10 }}>
                  {[
                    { key: "status", label: "Status", value: taskFilter, onChange: setTaskFilter, options: ["Active", ...TASK_STATUSES, "All"] },
                    { key: "priority", label: "Priority", options: ["All", "Critical", "High", "Moderate", "Low"] },
                    { key: "account", label: "Account", options: ["All", ...taskAccountOpts] },
                    { key: "cd", label: "CD", options: ["All", ...taskCdOpts] },
                    { key: "owner", label: "Owner", options: ["All", ...reps] },
                    { key: "sseSme", label: "SSE / SME", options: ["All", ...taskSseOpts] },
                    { key: "isOpp", label: "Opp", options: ["All", "Yes", "No"] },
                    { key: "due", label: "Due date", options: ["All", "Overdue", "Next 7 days", "Next 30 days", "No due date"] },
                  ].map(f => (
                    <div key={f.key}>
                      <label style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, display: "block", marginBottom: 4 }}>{f.label}</label>
                      <select
                        value={f.value !== undefined ? f.value : taskFieldFilters[f.key]}
                        onChange={e => f.onChange ? f.onChange(e.target.value) : setTaskFieldFilter(f.key, e.target.value)}
                        style={{ ...inputStyle, marginBottom: 0, fontSize: 12.5, padding: "7px 9px" }}
                      >
                        {f.options.map(o => <option key={o} value={o}>{o}</option>)}
                      </select>
                    </div>
                  ))}
                </div>
              </Card>

              {taskGroups.map(g => {
                const topPriorityColor = g.tasks[0] ? (PRIORITY_COLOR[g.tasks[0].priority] || C.textMute) : C.textMute;
                return (
                  <Card key={g.key} style={{ padding: 0, overflow: "hidden", marginBottom: 12 }}>
                    <div style={{
                      display: "flex", alignItems: "center", gap: 10, padding: "13px 16px",
                      borderLeft: `3px solid ${topPriorityColor}`
                    }}>
                      <div style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13.5, fontWeight: 700, color: C.text, display: "flex", alignItems: "center", gap: 6 }}>
                          {g.key === "unassigned" ? <AlertCircle size={13} color={C.textMute} /> : <Building2 size={13} color={C.textMute} />}
                          {g.title}
                        </div>
                        <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 1 }}>{g.subtitle}</div>
                      </div>
                      <span style={{ fontSize: 11.5, color: C.textMute, fontWeight: 600 }}>{g.tasks.length} {g.tasks.length === 1 ? "task" : "tasks"}</span>
                    </div>
                    <div style={{ borderTop: `1px solid ${C.border}`, overflowX: "auto" }}>
                      <div style={{ minWidth: 980 }}>
                        <div style={{
                          display: "grid", gridTemplateColumns: "1.6fr 0.8fr 0.8fr 0.9fr 0.8fr 0.8fr 1fr 0.6fr 1.8fr",
                          padding: "8px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4,
                          borderBottom: `1px solid ${C.border}`, fontWeight: 700, background: "#F8FAFC"
                        }}>
                          <span>Name</span><span>CD</span><span>Owner</span><span>SSE / SME</span><span>Priority</span><span>Due date</span><span>Status</span><span>Opp</span><span>Notes / Next steps</span>
                        </div>
                        {g.tasks.map(t => {
                          const overdue = t.status !== "Done" && t.dueDate && daysUntil(t.dueDate) < 0;
                          return (
                            <div key={t.id} style={{
                              display: "grid", gridTemplateColumns: "1.6fr 0.8fr 0.8fr 0.9fr 0.8fr 0.8fr 1fr 0.6fr 1.8fr",
                              padding: "11px 16px", borderBottom: `1px solid ${C.border}`, background: "#FAFBFC", alignItems: "center", gap: 6
                            }}>
                              <span onClick={() => setTaskModal(t)} style={{
                                fontSize: 13, fontWeight: 500, cursor: "pointer",
                                color: t.status === "Done" ? C.textMute : C.text, textDecoration: t.status === "Done" ? "line-through" : "none"
                              }}>{t.title}</span>
                              <span style={{ fontSize: 12.5, color: C.textSoft }}>{t.cd || "—"}</span>
                              <span style={{ fontSize: 12.5, color: C.textSoft }}>{t.rep || "—"}</span>
                              <span style={{ fontSize: 12.5, color: C.textSoft }}>{t.sseSme || "—"}</span>
                              <span><Pill color={PRIORITY_COLOR[t.priority] || C.textMute} bg={PRIORITY_BG[t.priority] || "#F1F5F9"}>{t.priority}</Pill></span>
                              <span style={{ fontSize: 12.5, color: overdue ? C.red : C.textSoft, fontWeight: overdue ? 700 : 400 }}>{fmtDate(t.dueDate)}</span>
                              <span>
                                <select
                                  value={t.status}
                                  onChange={e => upsert("tasks", { ...t, status: e.target.value })}
                                  style={{
                                    fontSize: 11, fontWeight: 700, padding: "4px 8px", borderRadius: 999, border: "none", cursor: "pointer",
                                    color: TASK_STATUS_COLOR[t.status], background: TASK_STATUS_BG[t.status]
                                  }}
                                >
                                  {TASK_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
                                </select>
                              </span>
                              <span>{t.isOpp ? <Pill color={C.blue} bg="#DBEAFE">Opp</Pill> : <span style={{ color: C.textMute, fontSize: 12.5 }}>—</span>}</span>
                              <span style={{ fontSize: 12, color: C.textMute, whiteSpace: "normal", wordBreak: "break-word" }}>{t.notes || "—"}</span>
                            </div>
                          );
                        })}
                      </div>
                    </div>
                  </Card>
                );
              })}
              {taskGroups.length === 0 && (
                <Card style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No tasks yet.</Card>
              )}
            </div>
          )}

          {tab === "settings" && (
            <div style={{ maxWidth: 620 }}>
              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 4 }}>Organization settings</h4>
              <p style={{ fontSize: 12, color: C.textMute, marginBottom: 18 }}>
                This data is shared with anyone who opens this command center. Regions roll up teams, and teams roll up reps — use the filters in the header to view any level, from a single rep up to the whole organization.
              </p>
              <Field label="Organization name">
                <input style={inputStyle} value={data.orgName} onChange={e => setData({ ...data, orgName: e.target.value })} />
              </Field>

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 4, marginTop: 20 }}>Regions</h4>
              <p style={{ fontSize: 12, color: C.textMute, marginBottom: 12 }}>e.g. Central, East, West.</p>
              {data.regions.map((r, i) => {
                const teamCount = data.teams.filter(t => t.regionId === r.id).length;
                return (
                  <div key={r.id} style={{ display: "flex", gap: 8, marginBottom: 8, alignItems: "center" }}>
                    <input style={{ ...inputStyle, marginBottom: 0 }} value={r.name} onChange={e => {
                      const next = data.regions.map(x => x.id === r.id ? { ...x, name: e.target.value } : x);
                      setData({ ...data, regions: next });
                    }} />
                    <span style={{ fontSize: 11, color: C.textMute, minWidth: 70 }}>{teamCount} team{teamCount === 1 ? "" : "s"}</span>
                    <button
                      style={{ ...iconBtnStyle, opacity: teamCount > 0 ? 0.4 : 1, cursor: teamCount > 0 ? "not-allowed" : "pointer" }}
                      title={teamCount > 0 ? "Reassign this region's teams before deleting it" : "Delete region"}
                      onClick={() => { if (teamCount === 0) setData({ ...data, regions: data.regions.filter(x => x.id !== r.id) }); }}>
                      <Trash2 size={14} />
                    </button>
                  </div>
                );
              })}
              <button style={{ ...ghostBtn, marginTop: 4, marginBottom: 24 }}
                onClick={() => setData({ ...data, regions: [...data.regions, { id: newId(), name: "New region" }] })}>
                <Plus size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Add region
              </button>

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 4 }}>Teams</h4>
              <p style={{ fontSize: 12, color: C.textMute, marginBottom: 12 }}>Each team belongs to one region and can have any number of reps. Add as many teams per region as you need.</p>
              {data.teams.map(t => (
                <div key={t.id} style={{ border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, marginBottom: 12, background: "#F8FAFC" }}>
                  <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr auto", gap: 8, marginBottom: 10, alignItems: "center" }}>
                    <input style={{ ...inputStyle, marginBottom: 0 }} value={t.name} onChange={e => {
                      const next = data.teams.map(x => x.id === t.id ? { ...x, name: e.target.value } : x);
                      setData({ ...data, teams: next });
                    }} />
                    <select style={{ ...inputStyle, marginBottom: 0 }} value={t.regionId} onChange={e => {
                      const next = data.teams.map(x => x.id === t.id ? { ...x, regionId: e.target.value } : x);
                      setData({ ...data, teams: next });
                    }}>
                      {data.regions.map(r => <option key={r.id} value={r.id}>{r.name}</option>)}
                    </select>
                    <button
                      style={{ ...iconBtnStyle, opacity: data.teams.length > 1 ? 1 : 0.4, cursor: data.teams.length > 1 ? "pointer" : "not-allowed" }}
                      title={data.teams.length > 1 ? "Delete team" : "At least one team is required"}
                      onClick={() => { if (data.teams.length > 1) setData({ ...data, teams: data.teams.filter(x => x.id !== t.id) }); }}>
                      <Trash2 size={14} />
                    </button>
                  </div>
                  <label style={{ ...labelStyle, marginBottom: 6 }}>Reps</label>
                  {t.reps.map((r, ri) => (
                    <div key={ri} style={{ display: "flex", gap: 8, marginBottom: 6 }}>
                      <input style={{ ...inputStyle, marginBottom: 0 }} value={r} onChange={e => {
                        const nextReps = t.reps.map((x, idx) => idx === ri ? e.target.value : x);
                        setData({ ...data, teams: data.teams.map(x => x.id === t.id ? { ...x, reps: nextReps } : x) });
                      }} />
                      <button style={iconBtnStyle} onClick={() => {
                        const nextReps = t.reps.filter((_, idx) => idx !== ri);
                        setData({ ...data, teams: data.teams.map(x => x.id === t.id ? { ...x, reps: nextReps } : x) });
                      }}><Trash2 size={14} /></button>
                    </div>
                  ))}
                  <button style={{ ...ghostBtn, padding: "5px 10px", fontSize: 12 }} onClick={() => {
                    setData({ ...data, teams: data.teams.map(x => x.id === t.id ? { ...x, reps: [...x.reps, "New rep"] } : x) });
                  }}><Plus size={12} style={{ verticalAlign: -2, marginRight: 4 }} />Add rep</button>
                </div>
              ))}
              <button style={{ ...ghostBtn, marginBottom: 24 }} onClick={() => {
                setData({ ...data, teams: [...data.teams, { id: newId(), name: "New team", regionId: data.regions[0]?.id || "central", reps: [] }] });
              }}><Plus size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Add team</button>

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 10 }}>Dashboard thresholds</h4>
              <div style={rowStyle}>
                <Field label="Pipeline coverage target (×)">
                  <input type="number" style={inputStyle} value={data.settings.coverageTarget}
                    onChange={e => setData({ ...data, settings: { ...data.settings, coverageTarget: Number(e.target.value) || 0 } })} />
                </Field>
                <Field label="Stale threshold (days)">
                  <input type="number" style={inputStyle} value={data.settings.staleDays}
                    onChange={e => setData({ ...data, settings: { ...data.settings, staleDays: Number(e.target.value) || 0 } })} />
                </Field>
              </div>

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 4, marginTop: 10 }}>Backup & restore</h4>
              <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 12 }}>
                Download everything in this command center as a file you keep yourself — a safety net independent of Claude's storage.
              </p>
              <div style={{ display: "flex", gap: 8, marginBottom: 12 }}>
                <button style={ghostBtn} onClick={exportBackup}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6, transform: "rotate(180deg)" }} />Download backup</button>
                <button style={ghostBtn} onClick={() => backupFileRef.current?.click()}><RefreshCw size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Restore from backup</button>
                <input ref={backupFileRef} type="file" accept=".json" onChange={handleBackupFile} style={{ display: "none" }} />
              </div>
              {restoreError && <div style={{ fontSize: 12.5, color: C.red, marginBottom: 12 }}>{restoreError}</div>}
              {pendingRestore && (
                <div style={{ background: "#FFFBEB", border: `1px solid #FDE68A`, borderRadius: 10, padding: 14, marginBottom: 16 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#92400E", marginBottom: 6 }}>Confirm restore</div>
                  <div style={{ fontSize: 12.5, color: "#78350F", marginBottom: 10 }}>
                    This file contains {(pendingRestore.opportunities || []).length} opportunities, {(pendingRestore.tasks || []).length} tasks,
                    {" "}{(pendingRestore.accounts || []).length} accounts, and {(pendingRestore.goals || []).length} goals.
                    Restoring will <strong>replace everything currently in this command center</strong> — this can't be undone.
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button style={ghostBtn} onClick={() => setPendingRestore(null)}>Cancel</button>
                    <button style={{ ...primaryBtn, background: C.amber }} onClick={confirmRestore}>Replace current data</button>
                  </div>
                </div>
              )}

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.text, marginBottom: 4, marginTop: 10 }}>Salesforce sync</h4>
              {data.importConfig.lastImportDate ? (
                <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
                  Last synced <strong style={{ color: C.text }}>{fmtDate(data.importConfig.lastImportDate)}</strong> from {data.importConfig.lastImportFileName}.
                  Column and stage mapping is remembered for next time.
                </p>
              ) : (
                <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>No sync has been run yet.</p>
              )}
              <button style={ghostBtn} onClick={() => setImportOpen(true)}><RefreshCw size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Sync now</button>
              {Object.keys(data.importConfig.colMap || {}).length > 0 && (
                <button style={{ ...ghostBtn, marginLeft: 8, ...dangerText, borderColor: "#FECACA" }}
                  onClick={() => setData({ ...data, importConfig: { colMap: {}, stageMap: {}, lastImportDate: data.importConfig.lastImportDate, lastImportFileName: data.importConfig.lastImportFileName } })}>
                  Reset column mapping
                </button>
              )}
            </div>
          )}
        </div>
      </div>

      {oppModal && (
        <Modal title={oppModal.id ? "Edit opportunity" : "New opportunity"} wide onClose={() => setOppModal(null)}>
          <OpportunityForm
            initial={oppModal.id ? oppModal : null}
            reps={reps}
            tasks={oppModal.id ? data.tasks.filter(t => normName(t.account) === normName(oppModal.account)) : []}
            onCancel={() => setOppModal(null)}
            onSave={(item) => { upsert("opportunities", item); setOppModal(null); }}
            onDelete={(id) => { remove("opportunities", id); setOppModal(null); }}
            onOpenTask={(t) => { setOppModal(null); setTaskModal(t); }}
            onAddTask={() => { const account = oppModal.account; setOppModal(null); setTaskModal({ account }); }}
          />
        </Modal>
      )}
      {actModal && (
        <Modal title={actModal.id ? "Edit activity" : "Log activity"} onClose={() => setActModal(null)}>
          <ActivityForm
            initial={actModal.id ? actModal : null}
            reps={reps}
            opportunities={data.opportunities}
            onCancel={() => setActModal(null)}
            onSave={(item) => { upsert("activities", item); setActModal(null); }}
            onDelete={(id) => { remove("activities", id); setActModal(null); }}
          />
        </Modal>
      )}
      {taskModal && (
        <Modal title={taskModal.id ? "Edit task" : "New task"} onClose={() => setTaskModal(null)}>
          <TaskForm
            initial={taskModal}
            reps={reps}
            accounts={data.accounts}
            onCancel={() => setTaskModal(null)}
            onSave={(item) => { upsert("tasks", item); setTaskModal(null); }}
            onDelete={(id) => { remove("tasks", id); setTaskModal(null); }}
          />
        </Modal>
      )}
      {goalModal && (
        <Modal title={goalModal.id ? "Edit goal" : "Set goal"} onClose={() => setGoalModal(null)}>
          <GoalForm
            initial={goalModal.id ? goalModal : null}
            reps={reps}
            period={period}
            onCancel={() => setGoalModal(null)}
            onSave={(item) => { upsert("goals", item); setGoalModal(null); }}
            onDelete={(id) => { remove("goals", id); setGoalModal(null); }}
          />
        </Modal>
      )}
      {importOpen && (
        <ImportWizard
          data={data}
          onClose={() => setImportOpen(false)}
          onApply={(next) => { setData(next); setImportOpen(false); }}
        />
      )}
      {accModal && (
        <Modal title={accModal.id ? "Edit account" : "New account"} onClose={() => setAccModal(null)}>
          <AccountForm
            initial={accModal.id ? accModal : null}
            reps={reps}
            computed={accModal.id ? accountComputed(accModal) : null}
            tasks={accModal.id ? data.tasks.filter(t => normName(t.account) === normName(accModal.name)) : []}
            onCancel={() => setAccModal(null)}
            onSave={(item) => { upsert("accounts", item); setAccModal(null); }}
            onDelete={(id) => { remove("accounts", id); setAccModal(null); }}
            onOpenTask={(t) => { setAccModal(null); setTaskModal(t); }}
            onAddTask={() => { const account = accModal.name; setAccModal(null); setTaskModal({ account }); }}
          />
        </Modal>
      )}
      {accountImportOpen && (
        <AccountImportWizard
          data={data}
          onClose={() => setAccountImportOpen(false)}
          onApply={(next) => { setData(next); setAccountImportOpen(false); }}
        />
      )}
      {taskImportOpen && (
        <TaskImportWizard
          data={data}
          reps={reps}
          onClose={() => setTaskImportOpen(false)}
          onApply={(next) => { setData(next); setTaskImportOpen(false); }}
        />
      )}
    </div>
  );
}
