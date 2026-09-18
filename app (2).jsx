import React, { useState, useEffect, useCallback, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, Target, PhoneCall, CheckSquare, Layers, Settings as SettingsIcon,
  Plus, X, Trash2, TrendingUp, AlertCircle, DollarSign, UploadCloud, Check,
  ArrowRight, RefreshCw, FileSpreadsheet, CheckCircle2, ArrowUpRight, ArrowDownRight, Minus, Building2, ChevronRight, Pencil, Search
} from "lucide-react";

const STORAGE_KEY = "scc-data-v1";

const STAGES = ["Prospecting", "Forecasting", "Pipeline", "Upside", "Strong Upside", "Commit", "Closed Won", "Closed Lost"];
const OPEN_STAGES = STAGES.filter(s => s !== "Closed Won" && s !== "Closed Lost");
const STAGE_COLOR = {
  "Prospecting": "#94A3B8",
  "Forecasting": "#8B5CF6",
  "Pipeline": "#3B82F6",
  "Upside": "#F59E0B",
  "Strong Upside": "#EA580C",
  "Commit": "#16A34A",
  "Closed Won": "#0D9488",
  "Closed Lost": "#DC2626",
};
const STAGE_PROBABILITY = {
  "Prospecting": 0.1, "Forecasting": 0.2, "Pipeline": 0.3, "Upside": 0.5, "Strong Upside": 0.7, "Commit": 0.9,
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
const PARTNER_ENGAGED_OPTIONS = ["Yes", "No", "NA"];
const TASK_STATUS_COLOR = { "Not Started": "#64748B", "Doing": "#3B82F6", "Follow Up": "#D97706", "Done": "#16A34A" };
const TASK_STATUS_BG = { "Not Started": "#F1F5F9", "Doing": "#DBEAFE", "Follow Up": "#FEF3C7", "Done": "#DCFCE7" };
const FORECAST_CATS = ["Commit", "Gut", "Stretch", "Not Forecasted"];
const FORECAST_COLOR = { Commit: "#16A34A", Gut: "#D97706", Stretch: "#3B82F6", "Not Forecasted": "#94A3B8" };

const IMPORT_FIELDS = [
  { key: "sfId", label: "Opportunity ID", required: true },
  { key: "name", label: "Opportunity name", required: true },
  { key: "account", label: "Account name", required: true },
  { key: "stage", label: "Stage", required: true },
  { key: "closeDate", label: "Close date", required: true },
  { key: "owner", label: "Owner / rep", required: false },
  { key: "commissionableMargin", label: "Commissionable margin $", required: false },
];

const ACCOUNT_IMPORT_FIELDS = [
  { key: "name", label: "Account", required: true },
  { key: "sfAccountId", label: "Salesforce Account ID (optional, preferred for matching)", required: false },
  { key: "territory", label: "Territory", required: false },
  { key: "tier", label: "Tier", required: false },
  { key: "priority", label: "Priority", required: false },
  { key: "status", label: "Status", required: false },
  { key: "cd", label: "AHEAD CD", required: false },
  { key: "notes", label: "Notes", required: false },
  { key: "flagServiceNow", label: "ServiceNow (X)", required: false },
  { key: "flagDynatrace", label: "Dynatrace (X)", required: false },
  { key: "flagDtDedicatedProg", label: "DT - Dedicated Prog (X)", required: false },
  { key: "flagTaniumRevList", label: "Tanium Rev List (X)", required: false },
  { key: "flagTaniumPtp", label: "Tanium PTP (X)", required: false },
  { key: "flagTaniumTarget", label: "Tanium Target (X)", required: false },
  { key: "flagGrafana", label: "Grafana (X)", required: false },
  { key: "flagObservability", label: "Observability (X)", required: false },
  { key: "flagNeuBird", label: "NeuBird AI (X)", required: false },
  { key: "snRep", label: "SN Rep", required: false },
  { key: "snStatus", label: "SN Status", required: false },
  { key: "snPriority", label: "SN Priority", required: false },
  { key: "snContacts", label: "SN Contacts", required: false },
  { key: "snNotes", label: "SN Notes", required: false },
  { key: "dtRep", label: "DT Rep", required: false },
  { key: "dtPsm", label: "PSM", required: false },
  { key: "dtStatus", label: "DT Status", required: false },
  { key: "dtPriority", label: "DT Priority", required: false },
  { key: "dtContacts", label: "DT Contacts", required: false },
  { key: "dtNotes", label: "DT Notes", required: false },
  { key: "taniumRep", label: "Tanium Rep", required: false },
  { key: "taniumStatus", label: "Tanium Status", required: false },
  { key: "taniumPriority", label: "Tanium Priority", required: false },
  { key: "taniumContacts", label: "Tanium Contacts", required: false },
  { key: "taniumNotes", label: "Tanium Notes", required: false },
  { key: "neubirdRep", label: "NeuBird Rep", required: false },
  { key: "neubirdStatus", label: "NeuBird Status", required: false },
  { key: "neubirdPriority", label: "NeuBird Priority", required: false },
  { key: "neubirdContacts", label: "NeuBird Contacts", required: false },
  { key: "neubirdNotes", label: "NeuBird Notes", required: false },
  { key: "owner", label: "Owner / rep (for team scoping)", required: false },
];

const ACCOUNT_TIERS = ["Enterprise", "Select", "Mid-Market"];
const ACCOUNT_PRIORITIES = ["High", "Moderate", "Low"];
const ACCOUNT_STATUSES = ["Active", "Not Active"];
const VENDOR_STATUSES = ["Customer", "Prospect", "Active Conversations"];
const VENDOR_STATUS_COLOR = { Customer: "#16A34A", Prospect: "#3B82F6", "Active Conversations": "#D97706" };
const VENDOR_STATUS_BG = { Customer: "#DCFCE7", Prospect: "#DBEAFE", "Active Conversations": "#FEF3C7" };
const BURNT_ORANGE = "#BF5700";
const ACCOUNT_GROUP_FIELD_LABEL = { cd: "CD", status: "Status", territory: "Territory", priority: "Priority", tier: "Tier" };
const ACCOUNT_GROUP_ORDER = { priority: ACCOUNT_PRIORITIES, status: ACCOUNT_STATUSES, tier: ACCOUNT_TIERS };

const CAMPAIGN_STATUSES = ["Planning", "Active", "Paused", "Completed"];
const CAMPAIGN_STATUS_COLOR = { Planning: "#64748B", Active: "#16A34A", Paused: "#D97706", Completed: "#3B82F6" };
const CAMPAIGN_STATUS_BG = { Planning: "#F1F5F9", Active: "#DCFCE7", Paused: "#FEF3C7", Completed: "#DBEAFE" };

const SOURCE_OPTIONS = ["Install Base", "Prospecting", "Partners"];
const SOURCE_COLOR = { "Install Base": "#3B82F6", "Prospecting": "#16A34A", "Partners": "#7C3AED", "Unspecified": "#94A3B8" };
const SOURCE_BG = { "Install Base": "#DBEAFE", "Prospecting": "#DCFCE7", "Partners": "#EDE9FE", "Unspecified": "#F1F5F9" };

const TASK_IMPORT_FIELDS = [
  { key: "name", label: "Name", required: true },
  { key: "account", label: "Account", required: true },
  { key: "cd", label: "CD", required: false },
  { key: "owner", label: "Owner", required: false },
  { key: "priority", label: "Priority", required: false },
  { key: "dueDate", label: "Due date", required: false },
  { key: "status", label: "Status", required: false },
  { key: "sseSme", label: "SSE / SME", required: false },
  { key: "partnerContact", label: "Partner Contact", required: false },
  { key: "partnerEngaged", label: "Partner Engaged (Yes/No/NA)", required: false },
  { key: "isOpp", label: "Opp (marked if tied to an opportunity)", required: false },
  { key: "dealReg", label: "Deal Reg?", required: false },
  { key: "notes", label: "Notes / Next Steps", required: false },
];

const OPP_IMPORT_FIELDS = [
  { key: "name", label: "Opportunity name", required: true },
  { key: "account", label: "Account", required: true },
  { key: "sfOppId", label: "Salesforce Opportunity ID (optional, preferred for matching)", required: false },
  { key: "sfAccountId", label: "Salesforce Account ID (optional, preferred for matching)", required: false },
  { key: "stage", label: "Stage", required: false },
  { key: "closeDate", label: "Close date", required: false },
  { key: "owner", label: "Owner / rep", required: false },
  { key: "commissionableMargin", label: "Commissionable margin $", required: false },
  { key: "source", label: "Source (Install Base / Prospecting / Partners)", required: false },
  { key: "forecastCategory", label: "Forecast category", required: false },
  { key: "nextStep", label: "Next step", required: false },
  { key: "notes", label: "Notes", required: false },
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
const oppTotal = (o) => {
  if (o.commissionableMargin !== undefined && o.commissionableMargin !== null && o.commissionableMargin !== "") {
    return Number(o.commissionableMargin) || 0;
  }
  return (Number(o.productMargin) || 0) + (Number(o.servicesMargin) || 0) + (Number(o.managedServices) || 0);
};
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
  if (stage === "Strong Upside") return "Gut";
  if (stage === "Upside") return "Stretch";
  return "Not Forecasted";
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
  campaigns: [],
  settings: { coverageTarget: 3, staleDays: 30 },
  importConfig: { colMap: {}, stageMap: {}, lastImportDate: null, lastImportFileName: null },
  accountImportConfig: { colMap: {}, lastImportDate: null, lastImportFileName: null },
  taskImportConfig: { colMap: {}, lastImportDate: null, lastImportFileName: null },
  oppImportConfig: { colMap: {}, lastImportDate: null, lastImportFileName: null },
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

// Resolve an account by name to its id, creating a bare-bones Account record
// if nothing matches yet. Returns { accounts, id, name } — accounts is the
// (possibly updated) list to save back, id/name are the resolved account.
const findOrCreateAccount = (accounts, name) => {
  const trimmed = (name || "").toString().trim();
  if (!trimmed) return { accounts, id: null, name: "" };
  const existing = accounts.find(a => normName(a.name) === normName(trimmed));
  if (existing) return { accounts, id: existing.id, name: existing.name };
  const created = {
    id: uid(), name: trimmed, territory: "", tier: "", priority: "Moderate", status: "Active", cd: "", rep: "", notes: "",
    mdPriority: false, strategic: false, metCustomer: false,
    flagServiceNow: false, flagDynatrace: false, flagDtDedicatedProg: false, flagTaniumRevList: false,
    flagTaniumPtp: false, flagTaniumTarget: false, flagGrafana: false, flagObservability: false, flagNeuBird: false,
    snRep: "", snStatus: "", snPriority: "", snContacts: "", snNotes: "",
    dtRep: "", dtPsm: "", dtStatus: "", dtPriority: "", dtContacts: "", dtNotes: "",
    taniumRep: "", taniumStatus: "", taniumPriority: "", taniumContacts: "", taniumNotes: "",
    neubirdRep: "", neubirdStatus: "", neubirdPriority: "", neubirdContacts: "", neubirdNotes: "",
    createdAt: todayStr(),
  };
  return { accounts: [...accounts, created], id: created.id, name: created.name };
};

const mergeLoadedData = (parsed) => {
  let merged = {
    ...DEFAULT_DATA, ...parsed,
    settings: { ...DEFAULT_DATA.settings, ...(parsed.settings || {}) },
    importConfig: { ...DEFAULT_DATA.importConfig, ...(parsed.importConfig || {}) },
    accountImportConfig: { ...DEFAULT_DATA.accountImportConfig, ...(parsed.accountImportConfig || {}) },
    taskImportConfig: { ...DEFAULT_DATA.taskImportConfig, ...(parsed.taskImportConfig || {}) },
    oppImportConfig: { ...DEFAULT_DATA.oppImportConfig, ...(parsed.oppImportConfig || {}) },
    accounts: parsed.accounts || [],
    campaigns: parsed.campaigns || [],
  };
  // migrate legacy flat teamName/reps shape into regions/teams
  if (!parsed.teams && (parsed.teamName || parsed.reps)) {
    merged.orgName = parsed.teamName || DEFAULT_DATA.orgName;
    merged.regions = [{ id: "central", name: "Central" }];
    merged.teams = [{ id: "mi-in", name: parsed.teamName || "Team 1", regionId: "central", reps: parsed.reps || ["Rep 1", "Rep 2"] }];
  }
  if (!merged.regions || merged.regions.length === 0) merged.regions = DEFAULT_DATA.regions;
  if (!merged.teams || merged.teams.length === 0) merged.teams = DEFAULT_DATA.teams;

  // Backfill accountId on opportunities/tasks that only have the legacy
  // free-text account name, so the new relational links exist for data
  // entered before accounts/opportunities/tasks were tied together by id.
  let backfillAccounts = merged.accounts;
  merged.opportunities = (merged.opportunities || []).map(o => {
    let next = o;
    if (!next.accountId && next.account && next.account.trim()) {
      const res = findOrCreateAccount(backfillAccounts, next.account);
      backfillAccounts = res.accounts;
      next = { ...next, accountId: res.id };
    }
    if (next.commissionableMargin === undefined || next.commissionableMargin === null || next.commissionableMargin === "") {
      const legacySum = (Number(next.productMargin) || 0) + (Number(next.servicesMargin) || 0) + (Number(next.managedServices) || 0);
      if (legacySum > 0) next = { ...next, commissionableMargin: legacySum };
    }
    return next;
  });
  merged.tasks = (merged.tasks || []).map(t => {
    if (t.accountId || !t.account || !t.account.trim()) return t;
    const res = findOrCreateAccount(backfillAccounts, t.account);
    backfillAccounts = res.accounts;
    return { ...t, accountId: res.id };
  });
  merged.accounts = backfillAccounts;

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
        // A "not found" style error just means nothing has ever been saved
        // under this key yet — that's a normal, expected first-time state,
        // NOT a failure. Only genuine errors (network, auth, etc.) should
        // trigger retries and the visible warning banner.
        const msg = (e && e.message ? e.message : String(e)).toLowerCase();
        const isMissingKey = msg.includes("not found") || msg.includes("404") || msg.includes("no rows");
        if (isMissingKey) {
          setLoaded(true);
          setLastSyncedAt(Date.now());
          return;
        }
        if (attempt < 2) {
          await new Promise(r => setTimeout(r, 1200));
          continue;
        }
        // All retries failed on a genuine error. Do NOT silently show a blank
        // slate as if it were empty — surface this so nobody mistakes a real
        // load failure for lost data.
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
function OpportunityForm({ initial, reps, accounts, campaigns, tasks, onSave, onCancel, onDelete, onOpenTask, onAddTask }) {
  const isEdit = !!(initial && initial.id);
  const [f, setF] = useState({
    name: "", account: "", accountId: null, campaignId: "", source: "", stage: "Prospecting", closeDate: "",
    commissionableMargin: "",
    rep: reps[0] || "", nextStep: "", notes: "", forecastCategory: "Pipeline",
    ...(initial || {}),
  });
  const [fcTouched, setFcTouched] = useState(isEdit);
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const setStage = (stage) => {
    setF(prev => ({ ...prev, stage, forecastCategory: fcTouched ? prev.forecastCategory : suggestForecast(stage) }));
  };
  const submit = () => {
    if (!(f.name && f.account)) return;
    const stageChanged = !isEdit || initial.stage !== f.stage;
    onSave({
      ...f,
      id: initial?.id || uid(),
      lastStageChange: stageChanged ? todayStr() : (initial?.lastStageChange || todayStr()),
      createdAt: initial?.createdAt || todayStr(),
    });
  };
  const campaignsForAccount = f.accountId ? campaigns.filter(c => c.accountIds.includes(f.accountId)) : [];
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
        <Field label="Account">
          <AccountAutocomplete accounts={accounts} name={f.account} accountId={f.accountId} onChange={(name, accountId) => setF(prev => ({ ...prev, account: name, accountId, campaignId: prev.accountId === accountId ? prev.campaignId : "" }))} />
        </Field>
        <Field label="Owner (rep)">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Source">
          <select style={inputStyle} value={f.source || ""} onChange={e => set("source", e.target.value)}>
            <option value="">— Unspecified —</option>
            {SOURCE_OPTIONS.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="Sourced from campaign">
          <CampaignPicker campaigns={campaignsForAccount} value={f.campaignId} onChange={v => set("campaignId", v)} />
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
        <select style={inputStyle} value={f.forecastCategory || "Not Forecasted"} onChange={e => { setFcTouched(true); set("forecastCategory", e.target.value); }}>
          {FORECAST_CATS.map(c => <option key={c} value={c}>{c}</option>)}
        </select>
      </Field>
      <Field label="Commissionable margin ($)">
        <input type="number" style={inputStyle} value={f.commissionableMargin} onChange={e => set("commissionableMargin", e.target.value)} placeholder="0" />
      </Field>
      <Field label="Next step">
        <input style={inputStyle} value={f.nextStep} onChange={e => set("nextStep", e.target.value)} placeholder="What happens next?" />
      </Field>
      <Field label="Notes">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>
      {isEdit && (
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
              <div style={{ padding: "10px 12px", fontSize: 12.5, color: C.textMute, background: "#F8FAFC" }}>No tasks linked to this opportunity yet.</div>
            )}
            {(tasks || []).map((t, i) => {
              const overdue = t.status !== "Done" && t.dueDate && daysUntil(t.dueDate) < 0;
              const hasNotes = !!(t.notes && t.notes.trim());
              return (
                <div key={t.id} onClick={() => onOpenTask && onOpenTask(t)} style={{
                  padding: "9px 12px",
                  borderBottom: i < tasks.length - 1 ? `1px solid ${C.border}` : "none",
                  background: "#F8FAFC", cursor: onOpenTask ? "pointer" : "default"
                }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                    <span style={{
                      fontSize: 13, fontWeight: 600, color: t.status === "Done" ? C.textMute : C.text,
                      textDecoration: t.status === "Done" ? "line-through" : "none", flex: 1
                    }}>{t.title}</span>
                    <Pill color={PRIORITY_COLOR[t.priority] || C.textMute} bg={PRIORITY_BG[t.priority] || "#F1F5F9"}>{t.priority}</Pill>
                    <span style={{ fontSize: 11.5, minWidth: 66, textAlign: "right", color: overdue ? C.red : C.textMute }}>{fmtDate(t.dueDate)}</span>
                  </div>
                  {hasNotes && (
                    <div style={{ fontSize: 12.5, color: C.textSoft, whiteSpace: "pre-wrap", marginTop: 6 }}>{t.notes}</div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
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
          <button style={primaryBtn} onClick={submit}>{isEdit ? "Save changes" : "Add opportunity"}</button>
        </div>
      </div>
    </div>
  );
}

// ---------- Account search/autocomplete (free text) ----------
function AccountAutocomplete({ accounts, name, accountId, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const linked = accountId ? accounts.find(a => a.id === accountId) : null;
  const q = query.trim().toLowerCase();
  const matches = (q
    ? accounts.filter(a => a.name.toLowerCase().includes(q))
    : accounts
  ).slice(0, 8);

  if (linked) {
    return (
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        background: "#F0FDF4", border: "1px solid #BBF7D0", borderRadius: 8, padding: "8px 11px", marginBottom: 14
      }}>
        <div style={{ minWidth: 0, display: "flex", alignItems: "center", gap: 6 }}>
          <Check size={13} color={C.green} style={{ flexShrink: 0 }} />
          <span style={{ fontSize: 13, color: C.text, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{linked.name}</span>
        </div>
        <button type="button" onClick={() => { onChange("", null); setQuery(""); }} style={{ ...iconBtnStyle, width: 26, height: 26, flexShrink: 0 }} title="Change account">
          <X size={13} />
        </button>
      </div>
    );
  }

  return (
    <div style={{ position: "relative", marginBottom: 14 }}>
      <input
        style={{ ...inputStyle, marginBottom: 0 }}
        placeholder="Search accounts, or type a new name…"
        value={query || name || ""}
        onFocus={() => setOpen(true)}
        onChange={e => { setQuery(e.target.value); onChange(e.target.value, null); setOpen(true); }}
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
              onMouseDown={e => { e.preventDefault(); onChange(a.name, a.id); setQuery(""); setOpen(false); }}
              style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, color: C.text }}
            >
              {a.name}{a.cd && <span style={{ color: C.textMute, fontSize: 11.5 }}> · CD: {a.cd}</span>}
            </div>
          ))}
        </div>
      )}
      {!q && name && (
        <div style={{ fontSize: 11, color: C.amber, marginTop: 4 }}>No existing account matches "{name}" — a new account record will be created when you save.</div>
      )}
    </div>
  );
}

// ---------- Opportunity search/select (for linking tasks to a specific deal) ----------
function OpportunityPicker({ opportunities, value, onChange }) {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);
  const selected = opportunities.find(o => o.id === value);
  const q = query.trim().toLowerCase();
  const matches = (q
    ? opportunities.filter(o => o.name.toLowerCase().includes(q) || (o.account || "").toLowerCase().includes(q))
    : opportunities
  ).slice(0, 8);

  if (selected) {
    return (
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 11px", marginBottom: 14
      }}>
        <div style={{ minWidth: 0 }}>
          <div style={{ fontSize: 13, color: C.text, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{selected.name}</div>
          <div style={{ fontSize: 11.5, color: C.textMute }}>{selected.account}</div>
        </div>
        <button type="button" onClick={() => onChange("")} style={{ ...iconBtnStyle, width: 26, height: 26, flexShrink: 0 }} title="Remove link">
          <X size={13} />
        </button>
      </div>
    );
  }

  return (
    <div style={{ position: "relative", marginBottom: 14 }}>
      <input
        style={{ ...inputStyle, marginBottom: 0 }}
        placeholder={opportunities.length ? "Search opportunities…" : "No opportunities for this account yet"}
        value={query}
        onFocus={() => setOpen(true)}
        onChange={e => { setQuery(e.target.value); setOpen(true); }}
        onBlur={() => setTimeout(() => setOpen(false), 150)}
      />
      {open && (
        <div style={{
          position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 30,
          background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
          boxShadow: "0 8px 24px rgba(15,23,42,0.14)", maxHeight: 230, overflowY: "auto"
        }}>
          <div
            onMouseDown={e => { e.preventDefault(); onChange(""); setQuery(""); setOpen(false); }}
            style={{ padding: "9px 12px", fontSize: 12.5, color: C.textMute, fontStyle: "italic", cursor: "pointer", borderBottom: `1px solid ${C.border}` }}
          >No opportunity</div>
          {matches.map(o => (
            <div key={o.id}
              onMouseDown={e => { e.preventDefault(); onChange(o.id); setQuery(""); setOpen(false); }}
              style={{ padding: "9px 12px", cursor: "pointer" }}
            >
              <div style={{ fontSize: 13, color: C.text, fontWeight: 500 }}>{o.name}</div>
              <div style={{ fontSize: 11.5, color: C.textMute }}>{o.account}</div>
            </div>
          ))}
          {matches.length === 0 && <div style={{ padding: "9px 12px", fontSize: 12.5, color: C.textMute }}>No matches</div>}
        </div>
      )}
    </div>
  );
}

// ---------- Campaign picker (for attributing an opportunity back to a campaign) ----------
function CampaignPicker({ campaigns, value, onChange }) {
  const [open, setOpen] = useState(false);
  const selected = campaigns.find(c => c.id === value);
  if (selected) {
    return (
      <div style={{
        display: "flex", alignItems: "center", justifyContent: "space-between",
        background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8, padding: "8px 11px", marginBottom: 14
      }}>
        <div style={{ minWidth: 0, display: "flex", alignItems: "center", gap: 8 }}>
          <span style={{ fontSize: 13, color: C.text, fontWeight: 600, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{selected.name}</span>
          <Pill color={CAMPAIGN_STATUS_COLOR[selected.status]} bg={CAMPAIGN_STATUS_BG[selected.status]}>{selected.status}</Pill>
        </div>
        <button type="button" onClick={() => onChange("")} style={{ ...iconBtnStyle, width: 26, height: 26, flexShrink: 0 }} title="Remove link">
          <X size={13} />
        </button>
      </div>
    );
  }
  return (
    <div style={{ position: "relative", marginBottom: 14 }}>
      <select
        style={inputStyle}
        value=""
        onChange={e => e.target.value && onChange(e.target.value)}
      >
        <option value="">{campaigns.length ? "No campaign — select to attribute" : "No campaigns target this account"}</option>
        {campaigns.map(c => <option key={c.id} value={c.id}>{c.name}</option>)}
      </select>
    </div>
  );
}

// ---------- Task form ----------
function TaskForm({ initial, reps, accounts, opportunities, onSave, onCancel, onDelete }) {
  const isEdit = !!(initial && initial.id);
  const [f, setF] = useState({
    title: "", account: "", accountId: null, oppId: "", cd: "", sseSme: "", partnerContact: "", partnerEngaged: "", dealReg: "", rep: reps[0] || "", dueDate: "",
    priority: "Moderate", status: "Not Started", isOpp: false, notes: "",
    ...(initial || {}),
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const setAccount = (name, accountId) => {
    setF(prev => {
      // if switching to a different account, drop any opportunity link that no longer applies
      const oppStillValid = prev.oppId && opportunities.find(o => o.id === prev.oppId)?.accountId === accountId;
      return { ...prev, account: name, accountId, oppId: oppStillValid ? prev.oppId : "" };
    });
  };
  const isAhead = normName(f.account) === "ahead";
  const oppsForAccount = f.accountId ? opportunities.filter(o => o.accountId === f.accountId) : [];
  const submit = () => {
    if (!f.title) return;
    onSave({ ...f, id: initial?.id || uid(), isOpp: f.oppId ? true : !!f.isOpp });
  };
  return (
    <div>
      <Field label="Name">
        <input style={inputStyle} value={f.title} onChange={e => set("title", e.target.value)} placeholder="What needs to happen?" />
      </Field>
      <div style={rowStyle}>
        <Field label="Account">
          <AccountAutocomplete accounts={accounts} name={f.account} accountId={f.accountId} onChange={setAccount} />
        </Field>
        <Field label="CD">
          <input style={inputStyle} value={f.cd} onChange={e => set("cd", e.target.value)} placeholder="Client / coverage director" />
        </Field>
      </div>
      <Field label={`Opportunity${isAhead ? " (not required for AHEAD)" : ""}`}>
        <OpportunityPicker opportunities={oppsForAccount} value={f.oppId} onChange={v => set("oppId", v)} />
      </Field>
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
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 12 }}>
        <Field label="Partner Contact">
          <input style={inputStyle} value={f.partnerContact} onChange={e => set("partnerContact", e.target.value)} placeholder="Partner-side contact" />
        </Field>
        <Field label="Partner Engaged">
          <select style={inputStyle} value={f.partnerEngaged || ""} onChange={e => set("partnerEngaged", e.target.value)}>
            <option value="">—</option>
            {PARTNER_ENGAGED_OPTIONS.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </Field>
        <Field label="Deal Reg?">
          <input style={inputStyle} value={f.dealReg} onChange={e => set("dealReg", e.target.value)} placeholder="e.g. TBD, Pending, Approved" />
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
      <Field label="Status">
        <select style={inputStyle} value={f.status} onChange={e => set("status", e.target.value)}>
          {TASK_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
        </select>
      </Field>
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
          <button style={primaryBtn} onClick={submit}>
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
    name: "", territory: "", tier: "", priority: "Moderate", status: "Active", cd: "", rep: reps[0] || "", notes: "",
    mdPriority: false, strategic: false, metCustomer: false,
    flagServiceNow: false, flagDynatrace: false, flagDtDedicatedProg: false, flagTaniumRevList: false,
    flagTaniumPtp: false, flagTaniumTarget: false, flagGrafana: false, flagObservability: false, flagNeuBird: false,
    snRep: "", snStatus: "", snPriority: "", snContacts: "", snNotes: "",
    dtRep: "", dtPsm: "", dtStatus: "", dtPriority: "", dtContacts: "", dtNotes: "",
    taniumRep: "", taniumStatus: "", taniumPriority: "", taniumContacts: "", taniumNotes: "",
    neubirdRep: "", neubirdStatus: "", neubirdPriority: "", neubirdContacts: "", neubirdNotes: "",
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));
  const flagLabel = (key, label) => (
    <label style={{ display: "flex", alignItems: "center", gap: 7, fontSize: 12.5, color: C.text, cursor: "pointer" }}>
      <input type="checkbox" checked={!!f[key]} onChange={e => set(key, e.target.checked)}
        style={{ width: 14, height: 14, accentColor: C.green, cursor: "pointer" }} />
      {label}
    </label>
  );
  const sectionHeader = (label, color) => (
    <div style={{ fontSize: 12.5, fontWeight: 700, color: color || C.text, marginBottom: 10, marginTop: 18, paddingBottom: 6, borderBottom: `1px solid ${C.border}` }}>{label}</div>
  );

  return (
    <div>
      {computed && (
        <div style={{ display: "flex", gap: 8, marginBottom: 14, flexWrap: "wrap" }}>
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
        <Field label="Territory"><input style={inputStyle} value={f.territory} onChange={e => set("territory", e.target.value)} placeholder="e.g. Michigan, Indiana" /></Field>
        <Field label="Tier">
          <select style={inputStyle} value={f.tier} onChange={e => set("tier", e.target.value)}>
            <option value="">—</option>
            {ACCOUNT_TIERS.map(t => <option key={t} value={t}>{t}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Priority">
          <select style={inputStyle} value={f.priority} onChange={e => set("priority", e.target.value)}>
            {ACCOUNT_PRIORITIES.map(p => <option key={p} value={p}>{p}</option>)}
          </select>
        </Field>
        <Field label="Status">
          <select style={inputStyle} value={f.status} onChange={e => set("status", e.target.value)}>
            {ACCOUNT_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="AHEAD CD"><input style={inputStyle} value={f.cd} onChange={e => set("cd", e.target.value)} placeholder="Client / coverage director" /></Field>
        <Field label="Rep (team scoping)">
          <select style={inputStyle} value={f.rep} onChange={e => set("rep", e.target.value)}>
            {reps.map(r => <option key={r} value={r}>{r}</option>)}
          </select>
        </Field>
      </div>
      <Field label="Notes">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>

      {sectionHeader("Practice flags")}
      <div style={{
        background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, marginBottom: 6,
        display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10
      }}>
        {flagLabel("flagServiceNow", "ServiceNow")}
        {flagLabel("flagDynatrace", "Dynatrace")}
        {flagLabel("flagDtDedicatedProg", "DT - Dedicated Prog")}
        {flagLabel("flagTaniumRevList", "Tanium Rev List")}
        {flagLabel("flagTaniumPtp", "Tanium PTP")}
        {flagLabel("flagTaniumTarget", "Tanium Target")}
        {flagLabel("flagGrafana", "Grafana")}
        {flagLabel("flagObservability", "Observability")}
        {flagLabel("flagNeuBird", "NeuBird AI")}
      </div>

      {sectionHeader("ServiceNow", "#3B82F6")}
      <div style={rowStyle}>
        <Field label="SN Rep"><input style={inputStyle} value={f.snRep} onChange={e => set("snRep", e.target.value)} /></Field>
        <Field label="SN Status">
          <select style={inputStyle} value={f.snStatus} onChange={e => set("snStatus", e.target.value)}>
            <option value="">—</option>
            {VENDOR_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      </div>
      <Field label="SN Priority"><input style={inputStyle} value={f.snPriority} onChange={e => set("snPriority", e.target.value)} placeholder="e.g. High, Medium" /></Field>
      <Field label="SN Contacts"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.snContacts} onChange={e => set("snContacts", e.target.value)} /></Field>
      <Field label="SN Notes"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.snNotes} onChange={e => set("snNotes", e.target.value)} /></Field>

      {sectionHeader("Dynatrace", "#7C3AED")}
      <div style={rowStyle}>
        <Field label="DT Rep"><input style={inputStyle} value={f.dtRep} onChange={e => set("dtRep", e.target.value)} /></Field>
        <Field label="PSM"><input style={inputStyle} value={f.dtPsm} onChange={e => set("dtPsm", e.target.value)} /></Field>
      </div>
      <div style={rowStyle}>
        <Field label="DT Status">
          <select style={inputStyle} value={f.dtStatus} onChange={e => set("dtStatus", e.target.value)}>
            <option value="">—</option>
            {VENDOR_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
        <Field label="DT Priority"><input style={inputStyle} value={f.dtPriority} onChange={e => set("dtPriority", e.target.value)} placeholder="e.g. High, Moderate, Low" /></Field>
      </div>
      <Field label="DT Contacts"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.dtContacts} onChange={e => set("dtContacts", e.target.value)} /></Field>
      <Field label="DT Notes"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.dtNotes} onChange={e => set("dtNotes", e.target.value)} /></Field>

      {sectionHeader("Tanium", "#EA580C")}
      <div style={rowStyle}>
        <Field label="Tanium Rep"><input style={inputStyle} value={f.taniumRep} onChange={e => set("taniumRep", e.target.value)} /></Field>
        <Field label="Tanium Status">
          <select style={inputStyle} value={f.taniumStatus} onChange={e => set("taniumStatus", e.target.value)}>
            <option value="">—</option>
            {VENDOR_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      </div>
      <Field label="Tanium Priority"><input style={inputStyle} value={f.taniumPriority} onChange={e => set("taniumPriority", e.target.value)} placeholder="e.g. High, Moderate, Low, A, B" /></Field>
      <Field label="Tanium Contacts"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.taniumContacts} onChange={e => set("taniumContacts", e.target.value)} /></Field>
      <Field label="Tanium Notes"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.taniumNotes} onChange={e => set("taniumNotes", e.target.value)} /></Field>

      {sectionHeader("NeuBird AI", BURNT_ORANGE)}
      <div style={rowStyle}>
        <Field label="NeuBird Rep"><input style={inputStyle} value={f.neubirdRep} onChange={e => set("neubirdRep", e.target.value)} /></Field>
        <Field label="NeuBird Status">
          <select style={inputStyle} value={f.neubirdStatus} onChange={e => set("neubirdStatus", e.target.value)}>
            <option value="">—</option>
            {VENDOR_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      </div>
      <Field label="NeuBird Priority"><input style={inputStyle} value={f.neubirdPriority} onChange={e => set("neubirdPriority", e.target.value)} placeholder="e.g. High, Moderate, Low" /></Field>
      <Field label="NeuBird Contacts"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.neubirdContacts} onChange={e => set("neubirdContacts", e.target.value)} /></Field>
      <Field label="NeuBird Notes"><textarea style={{ ...inputStyle, minHeight: 44, resize: "vertical" }} value={f.neubirdNotes} onChange={e => set("neubirdNotes", e.target.value)} /></Field>

      {initial && (
        <div style={{ marginBottom: 16, marginTop: 18 }}>
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

// ---------- Multi-account picker (for assigning accounts to a campaign) ----------
function AccountMultiPicker({ accounts, selectedIds, onChange }) {
  const [open, setOpen] = useState(false);
  const [query, setQuery] = useState("");
  const selected = selectedIds.map(id => accounts.find(a => a.id === id)).filter(Boolean);
  const q = query.trim().toLowerCase();
  const matches = accounts
    .filter(a => !selectedIds.includes(a.id))
    .filter(a => !q || a.name.toLowerCase().includes(q))
    .slice(0, 8);
  return (
    <div style={{ marginBottom: 14 }}>
      {selected.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
          {selected.map(a => (
            <span key={a.id} style={{
              display: "flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 600, color: C.text,
              background: "#F1F5F9", border: `1px solid ${C.border}`, borderRadius: 999, padding: "4px 6px 4px 10px"
            }}>
              {a.name}
              <span onClick={() => onChange(selectedIds.filter(id => id !== a.id))} style={{ cursor: "pointer", color: C.textMute, display: "flex" }}>
                <X size={12} />
              </span>
            </span>
          ))}
        </div>
      )}
      <div style={{ position: "relative" }}>
        <input
          style={{ ...inputStyle, marginBottom: 0 }}
          placeholder="Search accounts to add…"
          value={query}
          onFocus={() => setOpen(true)}
          onChange={e => { setQuery(e.target.value); setOpen(true); }}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
        />
        {open && (
          <div style={{
            position: "absolute", top: "calc(100% + 4px)", left: 0, right: 0, zIndex: 30,
            background: C.card, border: `1px solid ${C.border}`, borderRadius: 8,
            boxShadow: "0 8px 24px rgba(15,23,42,0.14)", maxHeight: 200, overflowY: "auto"
          }}>
            {matches.map(a => (
              <div key={a.id}
                onMouseDown={e => { e.preventDefault(); onChange([...selectedIds, a.id]); setQuery(""); }}
                style={{ padding: "8px 12px", cursor: "pointer", fontSize: 13, color: C.text }}
              >
                {a.name}{a.cd && <span style={{ color: C.textMute, fontSize: 11.5 }}> · CD: {a.cd}</span>}
              </div>
            ))}
            {matches.length === 0 && <div style={{ padding: "8px 12px", fontSize: 12.5, color: C.textMute }}>{accounts.length === 0 ? "No accounts yet — add some in the Accounts tab" : "No matches"}</div>}
          </div>
        )}
      </div>
    </div>
  );
}

// ---------- Campaign form ----------
function CampaignForm({ initial, accounts, opportunities, onSave, onCancel, onDelete }) {
  const [f, setF] = useState(initial || {
    name: "", bdr: "", status: "Planning", startDate: "", endDate: "", notes: "", accountIds: [],
  });
  const set = (k, v) => setF(prev => ({ ...prev, [k]: v }));

  const attributed = initial ? opportunities.filter(o => o.campaignId === initial.id) : [];
  const won = attributed.filter(o => o.stage === "Closed Won");
  const totalValue = attributed.reduce((s, o) => s + oppTotal(o), 0);
  const wonValue = won.reduce((s, o) => s + oppTotal(o), 0);

  return (
    <div>
      <Field label="Campaign name">
        <input style={inputStyle} value={f.name} onChange={e => set("name", e.target.value)} placeholder="e.g. Q3 ServiceNow Push — Michigan" />
      </Field>
      <div style={rowStyle}>
        <Field label="BDR"><input style={inputStyle} value={f.bdr} onChange={e => set("bdr", e.target.value)} placeholder="Who's running this campaign" /></Field>
        <Field label="Status">
          <select style={inputStyle} value={f.status} onChange={e => set("status", e.target.value)}>
            {CAMPAIGN_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
          </select>
        </Field>
      </div>
      <div style={rowStyle}>
        <Field label="Start date"><input type="date" style={inputStyle} value={f.startDate} onChange={e => set("startDate", e.target.value)} /></Field>
        <Field label="End date"><input type="date" style={inputStyle} value={f.endDate} onChange={e => set("endDate", e.target.value)} /></Field>
      </div>
      <Field label={`Accounts (${f.accountIds.length})`}>
        <AccountMultiPicker accounts={accounts} selectedIds={f.accountIds} onChange={ids => set("accountIds", ids)} />
      </Field>
      <Field label="Notes">
        <textarea style={{ ...inputStyle, minHeight: 60, resize: "vertical" }} value={f.notes} onChange={e => set("notes", e.target.value)} />
      </Field>

      {initial && (
        <div style={{ marginBottom: 16 }}>
          <label style={{ ...labelStyle, marginBottom: 8 }}>Attribution</label>
          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8, padding: 12, textAlign: "center" }}>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.text }}>{attributed.length}</div>
              <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Opportunities</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 8, padding: 12, textAlign: "center" }}>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.text }}>{fmtMoneyShort(totalValue)}</div>
              <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Pipeline value</div>
            </div>
            <div style={{ background: "#F0FDF4", border: "1px solid #BBF7D0", borderRadius: 8, padding: 12, textAlign: "center" }}>
              <div style={{ fontSize: 18, fontWeight: 800, color: C.green }}>{won.length} · {fmtMoneyShort(wonValue)}</div>
              <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Closed won</div>
            </div>
          </div>
          {attributed.length > 0 && (
            <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
              {attributed.map((o, i) => (
                <div key={o.id} style={{
                  display: "flex", justifyContent: "space-between", padding: "8px 12px",
                  borderBottom: i < attributed.length - 1 ? `1px solid ${C.border}` : "none", fontSize: 12.5, background: "#F8FAFC"
                }}>
                  <span style={{ color: C.text }}>{o.name} <span style={{ color: C.textMute }}>— {o.account}</span></span>
                  <span style={{ color: C.textSoft }}>{o.stage} · {fmtMoneyShort(oppTotal(o))}</span>
                </div>
              ))}
            </div>
          )}
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
          <button style={primaryBtn} onClick={() => f.name && onSave({ ...f, id: initial?.id || uid(), createdAt: initial?.createdAt || todayStr() })}>
            {initial ? "Save changes" : "Add campaign"}
          </button>
        </div>
      </div>
    </div>
  );
}
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
          else if (f.key === "commissionableMargin") nextMap[f.key] = guessField(hdrs, ["commissionable margin", "commission", "margin", "gross profit", "gp"]);
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
    let accounts = data.accounts.map(a => ({ ...a }));
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
      let accountId = null;
      if (account) {
        const res = findOrCreateAccount(accounts, account);
        accounts = res.accounts;
        accountId = res.id;
      }
      const commissionableMargin = colMap.commissionableMargin ? numOrZero(r[colMap.commissionableMargin]) : undefined;

      const existing = existingBySfId.get(sfId);
      if (existing) {
        const idx = opps.findIndex(o => o.id === existing.id);
        const stageChanged = opps[idx].stage !== stage;
        opps[idx] = {
          ...opps[idx],
          name: name || opps[idx].name,
          account: account || opps[idx].account,
          accountId: accountId || opps[idx].accountId || null,
          stage, closeDate,
          rep: owner || opps[idx].rep,
          commissionableMargin: commissionableMargin !== undefined ? commissionableMargin : opps[idx].commissionableMargin,
          forecastCategory: stageChanged ? suggestForecast(stage) : (opps[idx].forecastCategory || suggestForecast(stage)),
          lastStageChange: stageChanged ? todayStr() : (opps[idx].lastStageChange || todayStr()),
          sfStatus: "active", lastSynced: todayStr(),
        };
      } else {
        opps.push({
          id: uid(), sfId, name, account, accountId, stage, closeDate,
          rep: owner || "Unassigned",
          commissionableMargin: commissionableMargin || 0,
          nextStep: "", notes: "", sfStatus: "active", lastSynced: todayStr(),
          forecastCategory: suggestForecast(stage), lastStageChange: todayStr(), createdAt: todayStr(),
        });
      }
    });

    opps = opps.map(o => (o.sfId && o.sfStatus !== "removed" && !seen.has(o.sfId)) ? { ...o, sfStatus: "removed" } : o);

    onApply({
      ...data,
      regions, teams, accounts,
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

  const findHeaderRowIdx = (raw) => {
    for (let i = 0; i < Math.min(raw.length, 6); i++) {
      if ((raw[i] || []).some(c => String(c).trim().toLowerCase() === "account")) return i;
    }
    return 0;
  };

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
          const raw = Papa.parse(evt.target.result, { skipEmptyLines: false }).data;
          const headerRowIdx = findHeaderRowIdx(raw);
          const hdrs = raw[headerRowIdx].map(h => String(h).trim());
          parsedRows = raw.slice(headerRowIdx + 1).map(r => {
            const obj = {};
            hdrs.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ""; });
            return obj;
          });
        } else {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const raw = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", header: 1 });
          const headerRowIdx = findHeaderRowIdx(raw);
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

        const savedMap = data.accountImportConfig.colMap || {};
        const nextMap = {};
        ACCOUNT_IMPORT_FIELDS.forEach(f => {
          if (savedMap[f.key] && hdrs.includes(savedMap[f.key])) { nextMap[f.key] = savedMap[f.key]; return; }
          if (f.key === "name") nextMap[f.key] = guessField(hdrs, ["account"], ["id"]);
          else if (f.key === "sfAccountId") nextMap[f.key] = guessField(hdrs, ["account id", "account  id", "sfdc id", "record id"]);
          else if (f.key === "territory") nextMap[f.key] = guessField(hdrs, ["territory"]);
          else if (f.key === "tier") nextMap[f.key] = guessField(hdrs, ["tier"]);
          else if (f.key === "priority") nextMap[f.key] = guessField(hdrs, ["priority"], ["sn ", "dt ", "tanium"]);
          else if (f.key === "status") nextMap[f.key] = guessField(hdrs, ["status"], ["sn ", "dt ", "tanium"]);
          else if (f.key === "cd") nextMap[f.key] = guessField(hdrs, ["ahead cd", "cd"]);
          else if (f.key === "notes") nextMap[f.key] = guessField(hdrs, ["notes"], ["sn ", "dt ", "tanium"]);
          else if (f.key === "flagServiceNow") nextMap[f.key] = guessField(hdrs, ["servicenow"], ["rep", "status", "priority", "contact", "note"]);
          else if (f.key === "flagDynatrace") nextMap[f.key] = guessField(hdrs, ["dynatrace"], ["rep", "status", "priority", "contact", "note", "dedicated"]);
          else if (f.key === "flagDtDedicatedProg") nextMap[f.key] = guessField(hdrs, ["dedicated prog"]);
          else if (f.key === "flagTaniumRevList") nextMap[f.key] = guessField(hdrs, ["tanium rev list"]);
          else if (f.key === "flagTaniumPtp") nextMap[f.key] = guessField(hdrs, ["tanium ptp"]);
          else if (f.key === "flagTaniumTarget") nextMap[f.key] = guessField(hdrs, ["tanium target"]);
          else if (f.key === "flagGrafana") nextMap[f.key] = guessField(hdrs, ["grafana"]);
          else if (f.key === "flagObservability") nextMap[f.key] = guessField(hdrs, ["observability"]);
          else if (f.key === "snRep") nextMap[f.key] = guessField(hdrs, ["sn rep"]);
          else if (f.key === "snStatus") nextMap[f.key] = guessField(hdrs, ["sn status"]);
          else if (f.key === "snPriority") nextMap[f.key] = guessField(hdrs, ["sn priority"]);
          else if (f.key === "snContacts") nextMap[f.key] = guessField(hdrs, ["sn contact"]);
          else if (f.key === "snNotes") nextMap[f.key] = guessField(hdrs, ["sn note"]);
          else if (f.key === "dtRep") nextMap[f.key] = guessField(hdrs, ["dt rep"]);
          else if (f.key === "dtPsm") nextMap[f.key] = guessField(hdrs, ["psm"]);
          else if (f.key === "dtStatus") nextMap[f.key] = guessField(hdrs, ["dt status"]);
          else if (f.key === "dtPriority") nextMap[f.key] = guessField(hdrs, ["dt priority"]);
          else if (f.key === "dtContacts") nextMap[f.key] = guessField(hdrs, ["dt contact"]);
          else if (f.key === "dtNotes") nextMap[f.key] = guessField(hdrs, ["dt note"]);
          else if (f.key === "taniumRep") nextMap[f.key] = guessField(hdrs, ["tanium rep"]);
          else if (f.key === "taniumStatus") nextMap[f.key] = guessField(hdrs, ["tanium status"]);
          else if (f.key === "taniumPriority") nextMap[f.key] = guessField(hdrs, ["tanium priority"]);
          else if (f.key === "taniumContacts") nextMap[f.key] = guessField(hdrs, ["tanium contact"]);
          else if (f.key === "taniumNotes") nextMap[f.key] = guessField(hdrs, ["tanium note"]);
          else if (f.key === "neubirdRep") nextMap[f.key] = guessField(hdrs, ["neubird rep"]);
          else if (f.key === "neubirdStatus") nextMap[f.key] = guessField(hdrs, ["neubird status"]);
          else if (f.key === "neubirdPriority") nextMap[f.key] = guessField(hdrs, ["neubird priority"]);
          else if (f.key === "neubirdContacts") nextMap[f.key] = guessField(hdrs, ["neubird contact"]);
          else if (f.key === "neubirdNotes") nextMap[f.key] = guessField(hdrs, ["neubird note"]);
          else if (f.key === "flagNeuBird") nextMap[f.key] = guessField(hdrs, ["neubird"]);
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
  const TEXT_FIELDS = ["territory", "tier", "priority", "status", "cd", "notes",
    "snRep", "snStatus", "snPriority", "snContacts", "snNotes",
    "dtRep", "dtPsm", "dtStatus", "dtPriority", "dtContacts", "dtNotes",
    "taniumRep", "taniumStatus", "taniumPriority", "taniumContacts", "taniumNotes",
    "neubirdRep", "neubirdStatus", "neubirdPriority", "neubirdContacts", "neubirdNotes"];
  const FLAG_FIELDS = ["flagServiceNow", "flagDynatrace", "flagDtDedicatedProg", "flagTaniumRevList", "flagTaniumPtp", "flagTaniumTarget", "flagGrafana", "flagObservability", "flagNeuBird"];

  const buildPatch = (r) => {
    const patch = { name: (r[colMap.name] || "").toString().trim() };
    TEXT_FIELDS.forEach(k => { if (colMap[k]) patch[k] = (r[colMap[k]] || "").toString().trim(); });
    FLAG_FIELDS.forEach(k => { if (colMap[k]) patch[k] = toBool(r[colMap[k]]); });
    return patch;
  };

  const diff = useMemo(() => {
    if (step < 3) return null;
    const byName = new Map(data.accounts.map(a => [normName(a.name), a]));
    const bySfId = new Map(data.accounts.filter(a => a.sfAccountId).map(a => [a.sfAccountId, a]));
    let added = 0, updated = 0, matchedById = 0;
    const newRepNames = new Set();
    const knownReps = allRepNames(data);
    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      if (!name) return;
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.includes(owner)) newRepNames.add(owner);
      const sfAccountId = colMap.sfAccountId ? (r[colMap.sfAccountId] || "").toString().trim() : "";
      if (sfAccountId && bySfId.has(sfAccountId)) { updated++; matchedById++; }
      else if (byName.has(normName(name))) updated++;
      else added++;
    });
    return { added, updated, matchedById, newRepNames: Array.from(newRepNames), totalRows: rows.length };
  }, [step, rows, colMap, data]);

  const commit = () => {
    const byName = new Map(data.accounts.map(a => [normName(a.name), a]));
    const bySfId = new Map(data.accounts.filter(a => a.sfAccountId).map(a => [a.sfAccountId, a]));
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
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);
      const patch = buildPatch(r);
      if (owner) patch.rep = owner;
      const sfAccountId = colMap.sfAccountId ? (r[colMap.sfAccountId] || "").toString().trim() : "";
      if (sfAccountId) patch.sfAccountId = sfAccountId;

      const existing = (sfAccountId && bySfId.get(sfAccountId)) || byName.get(normName(name));
      if (existing) {
        const idx = accounts.findIndex(a => a.id === existing.id);
        accounts[idx] = { ...accounts[idx], ...patch };
      } else {
        accounts.push({
          id: uid(), name, territory: "", tier: "", priority: "Moderate", status: "Active", cd: "", rep: "", notes: "",
          mdPriority: false, strategic: false, metCustomer: false,
          flagServiceNow: false, flagDynatrace: false, flagDtDedicatedProg: false, flagTaniumRevList: false,
          flagTaniumPtp: false, flagTaniumTarget: false, flagGrafana: false, flagObservability: false, flagNeuBird: false,
          snRep: "", snStatus: "", snPriority: "", snContacts: "", snNotes: "",
          dtRep: "", dtPsm: "", dtStatus: "", dtPriority: "", dtContacts: "", dtNotes: "",
          taniumRep: "", taniumStatus: "", taniumPriority: "", taniumContacts: "", taniumNotes: "",
    neubirdRep: "", neubirdStatus: "", neubirdPriority: "", neubirdContacts: "", neubirdNotes: "",
          createdAt: todayStr(), ...patch,
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
            Export your account list from Excel as .csv or .xlsx, then upload it here. If the sheet has title rows above the real header row (like the IO Account Tracker), the header row (the one with "Account") is detected automatically. Nothing leaves this browser session.
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
            Matching uses account name, so re-importing the same tracker later updates existing accounts instead of duplicating them. "Active opportunity" and "met with them" aren't imported — the command center tracks those automatically from your pipeline and activity log.
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
            {diff.matchedById > 0
              ? `${diff.matchedById} of those matched by Salesforce Account ID; everything else matched by name. `
              : "Matching uses Salesforce Account ID when you've mapped that column, falling back to account name. "}
            Only the fields you mapped in step 2 get updated on matching accounts — anything left unmapped (like Notes, if you skip it) stays untouched.
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
          else if (f.key === "partnerContact") nextMap[f.key] = guessField(hdrs, ["partner contact", "partner"], ["engaged"]);
          else if (f.key === "partnerEngaged") nextMap[f.key] = guessField(hdrs, ["partner engaged", "engaged"]);
          else if (f.key === "isOpp") nextMap[f.key] = guessField(hdrs, ["opp"], ["deal reg"]);
          else if (f.key === "dealReg") nextMap[f.key] = guessField(hdrs, ["deal reg"]);
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
  const normPartnerEngaged = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    if (["yes", "y", "true", "1"].includes(s)) return "Yes";
    if (["no", "n", "false", "0"].includes(s)) return "No";
    if (["na", "n/a", "not applicable"].includes(s)) return "NA";
    return "";
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
      const account = colMap.account ? (r[colMap.account] || "").toString().trim() : "";
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);
      let accountId = null;
      if (account) {
        const res = findOrCreateAccount(accounts, account);
        accounts = res.accounts;
        accountId = res.id;
      }

      const patch = {
        title: name,
        account, accountId,
        cd: colMap.cd ? (r[colMap.cd] || "").toString().trim() : "",
        rep: owner || (teams[0]?.reps[0] || "Unassigned"),
        priority: colMap.priority ? normPriority(r[colMap.priority]) : "Moderate",
        dueDate: colMap.dueDate ? normalizeDate(r[colMap.dueDate]) : "",
        status: colMap.status ? normStatus(r[colMap.status]) : "Not Started",
        sseSme: colMap.sseSme ? (r[colMap.sseSme] || "").toString().trim() : "",
        partnerContact: colMap.partnerContact ? (r[colMap.partnerContact] || "").toString().trim() : "",
        partnerEngaged: colMap.partnerEngaged ? normPartnerEngaged(r[colMap.partnerEngaged]) : "",
        isOpp: colMap.isOpp ? !!(r[colMap.isOpp] && r[colMap.isOpp].toString().trim()) : false,
        dealReg: colMap.dealReg ? (r[colMap.dealReg] || "").toString().trim() : "",
        notes: colMap.notes ? (r[colMap.notes] || "").toString().trim() : "",
      };

      const key = normName(name) + "|" + normName(account);
      const existing = existingByKey.get(key);
      if (existing) {
        const idx = tasks.findIndex(t => t.id === existing.id);
        tasks[idx] = { ...tasks[idx], ...patch, id: tasks[idx].id };
      } else {
        tasks.push({ id: uid(), oppId: "", ...patch });
      }
    });

    onApply({
      ...data,
      regions, teams, accounts, tasks,
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

// ---------- Opportunity import wizard (general Excel/CSV, not Salesforce) ----------
function OpportunityImportWizard({ data, reps, onClose, onApply }) {
  const [step, setStep] = useState(1);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState([]);
  const [rows, setRows] = useState([]);
  const [colMap, setColMap] = useState(data.oppImportConfig?.colMap || {});
  const [error, setError] = useState("");
  const fileInputRef = useRef(null);

  const findHeaderRowIdx = (raw) => {
    for (let i = 0; i < Math.min(raw.length, 6); i++) {
      if ((raw[i] || []).some(c => String(c).trim().toLowerCase() === "name" || String(c).trim().toLowerCase() === "opportunity name")) return i;
    }
    return 0;
  };

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
          const raw = Papa.parse(evt.target.result, { skipEmptyLines: false }).data;
          const headerRowIdx = findHeaderRowIdx(raw);
          const hdrs = raw[headerRowIdx].map(h => String(h).trim());
          parsedRows = raw.slice(headerRowIdx + 1).map(r => {
            const obj = {};
            hdrs.forEach((h, i) => { obj[h] = r[i] !== undefined ? r[i] : ""; });
            return obj;
          });
        } else {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const raw = XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "", header: 1 });
          const headerRowIdx = findHeaderRowIdx(raw);
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

        const savedMap = data.oppImportConfig?.colMap || {};
        const nextMap = {};
        OPP_IMPORT_FIELDS.forEach(f => {
          if (savedMap[f.key] && hdrs.includes(savedMap[f.key])) { nextMap[f.key] = savedMap[f.key]; return; }
          if (f.key === "name") nextMap[f.key] = guessField(hdrs, ["opportunity name", "name"], ["account"]);
          else if (f.key === "account") nextMap[f.key] = guessField(hdrs, ["account"], ["id"]);
          else if (f.key === "sfOppId") nextMap[f.key] = guessField(hdrs, ["opportunity id", "opportunity  id", "sfdc id", "record id"], ["account"]);
          else if (f.key === "sfAccountId") nextMap[f.key] = guessField(hdrs, ["account id", "account  id"]);
          else if (f.key === "stage") nextMap[f.key] = guessField(hdrs, ["stage"]);
          else if (f.key === "closeDate") nextMap[f.key] = guessField(hdrs, ["close date", "closedate"]);
          else if (f.key === "owner") nextMap[f.key] = guessField(hdrs, ["owner", "rep"]);
          else if (f.key === "commissionableMargin") nextMap[f.key] = guessField(hdrs, ["commissionable margin", "commission", "margin", "gross profit", "gp"]);
          else if (f.key === "source") nextMap[f.key] = guessField(hdrs, ["source"]);
          else if (f.key === "forecastCategory") nextMap[f.key] = guessField(hdrs, ["forecast"]);
          else if (f.key === "nextStep") nextMap[f.key] = guessField(hdrs, ["next step"]);
          else if (f.key === "notes") nextMap[f.key] = guessField(hdrs, ["notes"]);
        });
        setColMap(nextMap);
        setStep(2);
      } catch (err) {
        setError("Couldn't read that file. Make sure it's a .csv or .xlsx export.");
      }
    };
    if (isCsv) reader.readAsText(file); else reader.readAsArrayBuffer(file);
  };

  const requiredMapped = OPP_IMPORT_FIELDS.filter(f => f.required).every(f => colMap[f.key]);
  const normStage = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    const found = STAGES.find(st => st.toLowerCase() === s);
    return found || "Prospecting";
  };
  const normSource = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    const found = SOURCE_OPTIONS.find(so => so.toLowerCase() === s);
    return found || "";
  };
  const normForecast = (v) => {
    const s = (v || "").toString().trim().toLowerCase();
    const found = FORECAST_CATS.find(fc => fc.toLowerCase() === s);
    return found || "";
  };

  const diff = useMemo(() => {
    if (step < 3) return null;
    const existingByKey = new Map(data.opportunities.map(o => [normName(o.name) + "|" + normName(o.account || ""), o]));
    const existingBySfId = new Map(data.opportunities.filter(o => o.sfId).map(o => [o.sfId, o]));
    let added = 0, updated = 0, skipped = 0, matchedById = 0;
    const newRepNames = new Set();
    const knownReps = allRepNames(data);
    rows.forEach(r => {
      const name = (r[colMap.name] || "").toString().trim();
      const account = (r[colMap.account] || "").toString().trim();
      if (!name || !account) { skipped++; return; }
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.includes(owner)) newRepNames.add(owner);
      const sfOppId = colMap.sfOppId ? (r[colMap.sfOppId] || "").toString().trim() : "";
      const existsById = sfOppId && existingBySfId.has(sfOppId);
      const key = normName(name) + "|" + normName(account);
      if (existsById) { updated++; matchedById++; }
      else if (existingByKey.has(key)) updated++;
      else added++;
    });
    return { added, updated, skipped, matchedById, newRepNames: Array.from(newRepNames), totalRows: rows.length };
  }, [step, rows, colMap, data]);

  const commit = () => {
    const existingByKey = new Map(data.opportunities.map(o => [normName(o.name) + "|" + normName(o.account || ""), o]));
    const existingBySfId = new Map(data.opportunities.filter(o => o.sfId).map(o => [o.sfId, o]));
    const accountsBySfId = new Map(data.accounts.filter(a => a.sfAccountId).map(a => [a.sfAccountId, a]));
    let opps = data.opportunities.map(o => ({ ...o }));
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
      const account = (r[colMap.account] || "").toString().trim();
      if (!name || !account) return;
      const owner = colMap.owner ? (r[colMap.owner] || "").toString().trim() : "";
      if (owner && !knownReps.has(owner)) addUnassignedRep(owner);

      // Resolve the account — prefer a Salesforce Account ID match if we have
      // one, falling back to name matching (and stamping the id on afterward
      // so future imports by id "stick" even if this row matched by name).
      const sfAccountId = colMap.sfAccountId ? (r[colMap.sfAccountId] || "").toString().trim() : "";
      let accountId;
      if (sfAccountId && accountsBySfId.has(sfAccountId)) {
        accountId = accountsBySfId.get(sfAccountId).id;
      } else {
        const res = findOrCreateAccount(accounts, account);
        accounts = res.accounts;
        accountId = res.id;
      }
      if (sfAccountId) {
        const idx = accounts.findIndex(a => a.id === accountId);
        if (idx >= 0 && !accounts[idx].sfAccountId) {
          accounts[idx] = { ...accounts[idx], sfAccountId };
          accountsBySfId.set(sfAccountId, accounts[idx]);
        }
      }

      const stage = colMap.stage ? normStage(r[colMap.stage]) : "Prospecting";
      const sfOppId = colMap.sfOppId ? (r[colMap.sfOppId] || "").toString().trim() : "";

      const patch = {
        name, account, accountId, stage,
        closeDate: colMap.closeDate ? normalizeDate(r[colMap.closeDate]) : "",
        rep: owner || (teams[0]?.reps[0] || "Unassigned"),
        commissionableMargin: colMap.commissionableMargin ? numOrZero(r[colMap.commissionableMargin]) : 0,
        source: colMap.source ? normSource(r[colMap.source]) : "",
        nextStep: colMap.nextStep ? (r[colMap.nextStep] || "").toString().trim() : "",
        notes: colMap.notes ? (r[colMap.notes] || "").toString().trim() : "",
      };
      if (sfOppId) patch.sfId = sfOppId;
      const forecastCategory = colMap.forecastCategory ? normForecast(r[colMap.forecastCategory]) : "";

      // Prefer matching by Salesforce Opportunity ID; fall back to name + account.
      const existing = (sfOppId && existingBySfId.get(sfOppId)) || existingByKey.get(normName(name) + "|" + normName(account));
      if (existing) {
        const idx = opps.findIndex(o => o.id === existing.id);
        const stageChanged = opps[idx].stage !== stage;
        opps[idx] = {
          ...opps[idx], ...patch,
          forecastCategory: forecastCategory || (stageChanged ? suggestForecast(stage) : opps[idx].forecastCategory) || suggestForecast(stage),
          lastStageChange: stageChanged ? todayStr() : (opps[idx].lastStageChange || todayStr()),
          id: opps[idx].id,
        };
      } else {
        opps.push({
          id: uid(), campaignId: "", ...patch,
          forecastCategory: forecastCategory || suggestForecast(stage),
          lastStageChange: todayStr(), createdAt: todayStr(),
        });
      }
    });

    onApply({
      ...data,
      regions, teams, accounts, opportunities: opps,
      oppImportConfig: { colMap, lastImportDate: todayStr(), lastImportFileName: fileName },
    });
  };

  return (
    <Modal title="Import opportunities from Excel" wide onClose={onClose}>
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
            Upload a spreadsheet of opportunities as .csv or .xlsx. If there's a title row above the real header row, the header row (the one with "Name") is detected automatically. This is separate from the Salesforce sync — use this for a one-off list or a source that isn't Salesforce. Nothing leaves this browser session.
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
          {data.oppImportConfig?.lastImportDate && (
            <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 14 }}>
              Last imported {fmtDate(data.oppImportConfig.lastImportDate)} from {data.oppImportConfig.lastImportFileName}.
            </div>
          )}
        </div>
      )}

      {step === 2 && (
        <div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 14 }}>
            Match each field to a column from <strong style={{ color: C.text }}>{fileName}</strong>. This is remembered for next time.
          </p>
          {OPP_IMPORT_FIELDS.map(f => (
            <div key={f.key} style={{ display: "grid", gridTemplateColumns: "220px 1fr", alignItems: "center", gap: 12, marginBottom: 10 }}>
              <label style={{ fontSize: 12.5, color: C.textSoft, fontWeight: 600 }}>{f.label}{f.required && <span style={{ color: C.red }}> *</span>}</label>
              <select style={{ ...inputStyle, marginBottom: 0 }} value={colMap[f.key] || ""} onChange={e => setColMap({ ...colMap, [f.key]: e.target.value })}>
                <option value="">— Not in file —</option>
                {headers.map(h => <option key={h} value={h}>{h}</option>)}
              </select>
            </div>
          ))}
          <p style={{ fontSize: 11.5, color: C.textMute, marginTop: 4 }}>
            Stage values match against {STAGES.join(", ")} — anything unrecognized defaults to Prospecting. Source matches Install Base/Prospecting/Partners; Forecast category matches Commit/Gut/Stretch/Not Forecasted — either left blank if unrecognized (forecast category then auto-suggests from stage, same as adding a deal by hand). The two Salesforce ID fields are optional — map them if your export has them and matching will use the ID instead of name + account, which is more reliable across re-imports and Salesforce syncs.
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
              <div style={{ fontSize: 11, color: C.textMute }}>New opportunities</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.blue }}>{diff.updated}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Updated</div>
            </div>
            <div style={{ background: "#F8FAFC", border: `1px solid ${C.border}`, borderRadius: 10, padding: 14, textAlign: "center" }}>
              <div style={{ fontSize: 20, fontWeight: 700, color: C.amber }}>{diff.skipped}</div>
              <div style={{ fontSize: 11, color: C.textMute }}>Skipped (missing name/account)</div>
            </div>
          </div>
          {diff.newRepNames.length > 0 && (
            <div style={{ fontSize: 12, color: C.textSoft, marginBottom: 14 }}>
              New owner name{diff.newRepNames.length > 1 ? "s" : " "} found: {diff.newRepNames.join(", ")}. They'll be placed on an "Unassigned" team — reassign them under Settings.
            </div>
          )}
          <div style={{ fontSize: 11.5, color: C.textMute, marginBottom: 14 }}>
            {diff.matchedById > 0
              ? `${diff.matchedById} of those matched by Salesforce Opportunity ID; everything else matched by opportunity name + account. `
              : "Matching uses Salesforce Opportunity ID when you've mapped that column, falling back to opportunity name + account. "}
            Accounts match by Salesforce Account ID when mapped, otherwise by name — either way, accounts that don't already exist are created automatically.
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
  { key: "gtro", label: "GTRO", icon: FileSpreadsheet },
  { key: "tasks", label: "Tasks", icon: CheckSquare },
  { key: "goals", label: "Goals", icon: Target },
  { key: "prospecting", label: "Prospecting", icon: PhoneCall },
  { key: "settings", label: "Settings", icon: SettingsIcon },
];

// build refresh marker 2
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
  const [taskModal, setTaskModal] = useState(null);
  const [goalModal, setGoalModal] = useState(null);
  const [campModal, setCampModal] = useState(null);
  const [accModal, setAccModal] = useState(null);
  const [importOpen, setImportOpen] = useState(false);
  const [accountImportOpen, setAccountImportOpen] = useState(false);
  const [taskImportOpen, setTaskImportOpen] = useState(false);
  const [oppImportOpen, setOppImportOpen] = useState(false);
  const [taskFilter, setTaskFilter] = useState("Active");
  const [taskGroupBy, setTaskGroupBy] = useState("account");
  const [taskSearch, setTaskSearch] = useState("");
  const [taskFieldFilters, setTaskFieldFilters] = useState({
    priority: "All", account: "All", cd: "All", owner: "All", sseSme: "All", partnerContact: "All", partnerEngaged: "All", isOpp: "All", dealReg: "All", due: "All",
  });
  const setTaskFieldFilter = (key, val) => setTaskFieldFilters(prev => ({ ...prev, [key]: val }));
  const clearTaskFieldFilters = () => { setTaskFilter("Active"); setTaskSearch(""); setTaskFieldFilters({ priority: "All", account: "All", cd: "All", owner: "All", sseSme: "All", partnerContact: "All", partnerEngaged: "All", isOpp: "All", dealReg: "All", due: "All" }); };
  const [accountFilter, setAccountFilter] = useState("All");
  const [accountGroupBy, setAccountGroupBy] = useState("none");
  const [expandedAccounts, setExpandedAccounts] = useState(new Set());
  const toggleAccountExpand = (id) => setExpandedAccounts(prev => prev.has(id) ? new Set() : new Set([id]));
  const [expandedTasks, setExpandedTasks] = useState(new Set());
  const toggleTaskExpand = (id) => setExpandedTasks(prev => prev.has(id) ? new Set() : new Set([id]));
  const [expandedOpps, setExpandedOpps] = useState(new Set());
  const toggleOppExpand = (id) => setExpandedOpps(prev => prev.has(id) ? new Set() : new Set([id]));
  const [pendingRestore, setPendingRestore] = useState(null);
  const [restoreError, setRestoreError] = useState("");
  const backupFileRef = useRef(null);
  const restoreExcelFileRef = useRef(null);
  const [resetConfirmOpen, setResetConfirmOpen] = useState(false);

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
    const hasActiveOpp = data.opportunities.some(o =>
      (o.accountId ? o.accountId === acc.id : normName(o.account) === normName(acc.name)) && OPEN_STAGES.includes(o.stage) && o.sfStatus !== "removed"
    );
    const hasMet = !!acc.metCustomer;
    return { hasActiveOpp, hasMet };
  }, [data.opportunities]);

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

  // Resolve an item's account link before saving: if it points at a real
  // account, keep the display name in sync (self-heals renames); if it's a
  // freshly-typed name with no id yet, create the account record now.
  const resolveAccountLink = (item, accountsList) => {
    let accounts = accountsList;
    let accountId = item.accountId || null;
    let accountName = (item.account || "").trim();
    if (accountId) {
      const existing = accounts.find(a => a.id === accountId);
      if (existing) accountName = existing.name;
      else accountId = null; // linked account was deleted; fall through to re-resolve by name
    }
    if (!accountId && accountName) {
      const res = findOrCreateAccount(accounts, accountName);
      accounts = res.accounts;
      accountId = res.id;
      accountName = res.name;
    }
    return { accounts, item: { ...item, account: accountName, accountId } };
  };

  const saveOpportunity = (item) => {
    const { accounts, item: resolved } = resolveAccountLink(item, data.accounts);
    const opportunities = data.opportunities.some(o => o.id === resolved.id)
      ? data.opportunities.map(o => o.id === resolved.id ? resolved : o)
      : [...data.opportunities, resolved];
    setData({ ...data, accounts, opportunities });
  };

  const saveTask = (item) => {
    const { accounts, item: resolved } = resolveAccountLink(item, data.accounts);
    const tasks = data.tasks.some(t => t.id === resolved.id)
      ? data.tasks.map(t => t.id === resolved.id ? resolved : t)
      : [...data.tasks, resolved];
    setData({ ...data, accounts, tasks });
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

  const exportExcelBackup = () => {
    const wb = XLSX.utils.book_new();
    const addSheet = (name, rows) => {
      const sheet = XLSX.utils.json_to_sheet(rows && rows.length ? rows : [{}]);
      XLSX.utils.book_append_sheet(wb, sheet, name.slice(0, 31));
    };

    addSheet("Accounts", data.accounts.map(a => ({
      Account: a.name, Territory: a.territory, Tier: a.tier, Priority: a.priority, Status: a.status, CD: a.cd, Rep: a.rep,
      "MD Priority": a.mdPriority ? "Yes" : "", "Ahead Strategic": a.strategic ? "Yes" : "", "Met Customer": a.metCustomer ? "Yes" : "No",
      ServiceNow: a.flagServiceNow ? "X" : "", "SN Rep": a.snRep, "SN Status": a.snStatus, "SN Priority": a.snPriority, "SN Contacts": a.snContacts, "SN Notes": a.snNotes,
      Dynatrace: a.flagDynatrace ? "X" : "", "DT Rep": a.dtRep, PSM: a.dtPsm, "DT Status": a.dtStatus, "DT Priority": a.dtPriority, "DT Contacts": a.dtContacts, "DT Notes": a.dtNotes,
      "Tanium Rev List": a.flagTaniumRevList ? "X" : "", "Tanium PTP": a.flagTaniumPtp ? "X" : "", "Tanium Target": a.flagTaniumTarget ? "X" : "",
      "Tanium Rep": a.taniumRep, "Tanium Status": a.taniumStatus, "Tanium Priority": a.taniumPriority, "Tanium Contacts": a.taniumContacts, "Tanium Notes": a.taniumNotes,
      "NeuBird AI": a.flagNeuBird ? "X" : "", "NeuBird Rep": a.neubirdRep, "NeuBird Status": a.neubirdStatus, "NeuBird Priority": a.neubirdPriority, "NeuBird Contacts": a.neubirdContacts, "NeuBird Notes": a.neubirdNotes,
      Grafana: a.flagGrafana ? "X" : "", Observability: a.flagObservability ? "X" : "", "DT Dedicated Prog": a.flagDtDedicatedProg ? "X" : "",
      Notes: a.notes,
    })));

    addSheet("Opportunities", data.opportunities.map(o => ({
      Name: o.name, Account: o.account, Stage: o.stage, "Close Date": o.closeDate, Owner: o.rep,
      "Commissionable Margin": o.commissionableMargin, Source: o.source, "Forecast Category": o.forecastCategory,
      Campaign: o.campaignId ? (data.campaigns.find(c => c.id === o.campaignId)?.name || "") : "",
      "Next Step": o.nextStep, Notes: o.notes, "Salesforce ID": o.sfId || "", "Record ID": o.id,
    })));

    addSheet("Tasks", data.tasks.map(t => ({
      Name: t.title, Account: t.account, Opportunity: t.oppId ? (data.opportunities.find(o => o.id === t.oppId)?.name || "") : "",
      CD: t.cd, Owner: t.rep, "SSE / SME": t.sseSme,
      "Partner Contact": t.partnerContact, "Partner Engaged": t.partnerEngaged, Priority: t.priority,
      "Due Date": t.dueDate, Status: t.status, Opp: t.isOpp ? "Yes" : "", "Deal Reg": t.dealReg, Notes: t.notes,
    })));

    addSheet("Campaigns", data.campaigns.map(c => ({
      Name: c.name, BDR: c.bdr, Status: c.status, "Start Date": c.startDate, "End Date": c.endDate,
      Accounts: c.accountIds.map(id => data.accounts.find(a => a.id === id)?.name).filter(Boolean).join(", "),
      Notes: c.notes,
    })));

    addSheet("Goals", data.goals.map(g => ({
      Rep: g.rep, Period: g.period,
      ...Object.fromEntries(CATEGORIES.map(c => [c.label + " Target", g[c.key + "Target"]])),
    })));

    XLSX.writeFile(wb, `${(data.orgName || "sales-command-center").replace(/\s+/g, "-").toLowerCase()}-backup-${todayStr()}.xlsx`);
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

  const handleExcelBackupFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setRestoreError("");
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
        const sheetRows = (name) => {
          const sheet = wb.Sheets[name];
          return sheet ? XLSX.utils.sheet_to_json(sheet, { raw: false, defval: "" }) : [];
        };
        const str = (v) => (v === undefined || v === null ? "" : v.toString().trim());
        const bool = (v) => ["yes", "x", "true", "1"].includes(str(v).toLowerCase());

        // Accounts first — everything else links to these by name.
        let accounts = [];
        const accountIdByName = new Map();
        sheetRows("Accounts").forEach(r => {
          const name = str(r["Account"]);
          if (!name) return;
          const acc = {
            id: uid(), name, territory: str(r["Territory"]), tier: str(r["Tier"]), priority: str(r["Priority"]) || "Moderate",
            status: str(r["Status"]) || "Active", cd: str(r["CD"]), rep: str(r["Rep"]), notes: str(r["Notes"]),
            mdPriority: bool(r["MD Priority"]), strategic: bool(r["Ahead Strategic"]), metCustomer: bool(r["Met Customer"]),
            flagServiceNow: bool(r["ServiceNow"]), snRep: str(r["SN Rep"]), snStatus: str(r["SN Status"]), snPriority: str(r["SN Priority"]), snContacts: str(r["SN Contacts"]), snNotes: str(r["SN Notes"]),
            flagDynatrace: bool(r["Dynatrace"]), dtRep: str(r["DT Rep"]), dtPsm: str(r["PSM"]), dtStatus: str(r["DT Status"]), dtPriority: str(r["DT Priority"]), dtContacts: str(r["DT Contacts"]), dtNotes: str(r["DT Notes"]),
            flagTaniumRevList: bool(r["Tanium Rev List"]), flagTaniumPtp: bool(r["Tanium PTP"]), flagTaniumTarget: bool(r["Tanium Target"]),
            taniumRep: str(r["Tanium Rep"]), taniumStatus: str(r["Tanium Status"]), taniumPriority: str(r["Tanium Priority"]), taniumContacts: str(r["Tanium Contacts"]), taniumNotes: str(r["Tanium Notes"]),
            flagNeuBird: bool(r["NeuBird AI"]), neubirdRep: str(r["NeuBird Rep"]), neubirdStatus: str(r["NeuBird Status"]), neubirdPriority: str(r["NeuBird Priority"]), neubirdContacts: str(r["NeuBird Contacts"]), neubirdNotes: str(r["NeuBird Notes"]),
            flagGrafana: bool(r["Grafana"]), flagObservability: bool(r["Observability"]), flagDtDedicatedProg: bool(r["DT Dedicated Prog"]),
            createdAt: todayStr(),
          };
          accounts.push(acc);
          accountIdByName.set(normName(name), acc.id);
        });
        const ensureAccountId = (name) => {
          const key = normName(name);
          if (!key) return null;
          if (accountIdByName.has(key)) return accountIdByName.get(key);
          const acc = { id: uid(), name: str(name), territory: "", tier: "", priority: "Moderate", status: "Active", cd: "", rep: "", notes: "", createdAt: todayStr() };
          accounts.push(acc);
          accountIdByName.set(key, acc.id);
          return acc.id;
        };

        // Campaigns next — only need account names resolved.
        const campaigns = [];
        const campaignIdByName = new Map();
        sheetRows("Campaigns").forEach(r => {
          const name = str(r["Name"]);
          if (!name) return;
          const accountNames = str(r["Accounts"]).split(",").map(s => s.trim()).filter(Boolean);
          const camp = {
            id: uid(), name, bdr: str(r["BDR"]), status: CAMPAIGN_STATUSES.includes(str(r["Status"])) ? str(r["Status"]) : "Planning",
            startDate: normalizeDate(r["Start Date"]), endDate: normalizeDate(r["End Date"]), notes: str(r["Notes"]),
            accountIds: accountNames.map(n => ensureAccountId(n)).filter(Boolean),
            createdAt: todayStr(),
          };
          campaigns.push(camp);
          campaignIdByName.set(normName(name), camp.id);
        });

        // Opportunities — resolve account + campaign by name.
        const opportunities = [];
        const oppIdByNameAccount = new Map();
        sheetRows("Opportunities").forEach(r => {
          const name = str(r["Name"]);
          const accountName = str(r["Account"]);
          if (!name || !accountName) return;
          const accountId = ensureAccountId(accountName);
          const stage = STAGES.find(s => s.toLowerCase() === str(r["Stage"]).toLowerCase()) || "Prospecting";
          const campaignName = str(r["Campaign"]);
          const opp = {
            id: uid(), name, account: accountName, accountId, stage,
            closeDate: normalizeDate(r["Close Date"]), rep: str(r["Owner"]) || "Unassigned",
            commissionableMargin: numOrZero(r["Commissionable Margin"]),
            source: SOURCE_OPTIONS.find(s => s.toLowerCase() === str(r["Source"]).toLowerCase()) || "",
            forecastCategory: FORECAST_CATS.find(f => f.toLowerCase() === str(r["Forecast Category"]).toLowerCase()) || suggestForecast(stage),
            campaignId: campaignName ? (campaignIdByName.get(normName(campaignName)) || "") : "",
            nextStep: str(r["Next Step"]), notes: str(r["Notes"]), sfId: str(r["Salesforce ID"]) || undefined,
            lastStageChange: todayStr(), createdAt: todayStr(),
          };
          opportunities.push(opp);
          oppIdByNameAccount.set(normName(name) + "|" + normName(accountName), opp.id);
        });

        // Tasks — resolve account + specific opportunity by name.
        const tasks = [];
        sheetRows("Tasks").forEach(r => {
          const title = str(r["Name"]);
          if (!title) return;
          const accountName = str(r["Account"]);
          const accountId = accountName ? ensureAccountId(accountName) : null;
          const oppName = str(r["Opportunity"]);
          const oppId = oppName && accountName ? (oppIdByNameAccount.get(normName(oppName) + "|" + normName(accountName)) || "") : "";
          tasks.push({
            id: uid(), title, account: accountName, accountId, oppId,
            cd: str(r["CD"]), rep: str(r["Owner"]), sseSme: str(r["SSE / SME"]),
            partnerContact: str(r["Partner Contact"]), partnerEngaged: str(r["Partner Engaged"]),
            priority: PRIORITIES.includes(str(r["Priority"])) ? str(r["Priority"]) : "Moderate",
            dueDate: normalizeDate(r["Due Date"]), status: TASK_STATUSES.includes(str(r["Status"])) ? str(r["Status"]) : "Not Started",
            isOpp: oppId ? true : bool(r["Opp"]), dealReg: str(r["Deal Reg"]), notes: str(r["Notes"]),
          });
        });

        // Goals — independent of the rest.
        const goals = [];
        sheetRows("Goals").forEach(r => {
          const rep = str(r["Rep"]);
          const period = str(r["Period"]);
          if (!rep || !period) return;
          const goal = { id: uid(), rep, period };
          CATEGORIES.forEach(c => { goal[c.key + "Target"] = numOrZero(r[c.label + " Target"]); });
          goals.push(goal);
        });

        if (accounts.length === 0 && opportunities.length === 0 && tasks.length === 0 && campaigns.length === 0 && goals.length === 0) {
          setRestoreError("Couldn't find any recognizable data in that workbook — make sure it's the Excel backup this tool generated.");
          return;
        }

        setPendingRestore({ ...data, accounts, opportunities, tasks, campaigns, goals });
      } catch (err) {
        setRestoreError("Couldn't read that file. Make sure it's the .xlsx backup downloaded from this tool.");
      }
    };
    reader.readAsArrayBuffer(file);
    e.target.value = "";
  };

  const performReset = () => {
    setData(JSON.parse(JSON.stringify(DEFAULT_DATA)));
    setResetConfirmOpen(false);
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

  // ---- Executive summary metrics (Pipeline / Forecast top sections) ----
  const weightedPipeline = openOpps.reduce((s, o) => s + oppTotal(o) * (STAGE_PROBABILITY[o.stage] ?? 0.1), 0);
  const closedWonAll = activeOpps.filter(o => o.stage === "Closed Won");
  const closedLostAll = activeOpps.filter(o => o.stage === "Closed Lost");
  const winRateDenom = closedWonAll.length + closedLostAll.length;
  const winRate = winRateDenom > 0 ? (closedWonAll.length / winRateDenom) * 100 : null;
  const avgDealSize = openOpps.length > 0 ? openPipelineValue / openOpps.length : 0;
  const topOpenOpps = openOpps.slice().sort((a, b) => oppTotal(b) - oppTotal(a)).slice(0, 5);

  const staleDays = data.settings.staleDays || 30;
  const staleOpps = openOpps.filter(o => o.lastStageChange && daysSince(o.lastStageChange) >= staleDays);
  const staleValue = staleOpps.reduce((s, o) => s + oppTotal(o), 0);

  // ---- Forecast ----
  const qLabel = currentQuarter();
  const qOpen = openOpps.filter(o => getQuarter(o.closeDate) === qLabel);
  const qCommit = qOpen.filter(o => (o.forecastCategory || "Not Forecasted") === "Commit");
  const qStretch = qOpen.filter(o => (o.forecastCategory || "Not Forecasted") === "Stretch");
  const qForecastValue = qCommit.reduce((s, o) => s + oppTotal(o), 0) + (includeStretchQ ? qStretch.reduce((s, o) => s + oppTotal(o), 0) : 0);
  const qForecastCount = qCommit.length + (includeStretchQ ? qStretch.length : 0);

  const fyOpen = openOpps.filter(o => o.closeDate && o.closeDate.slice(0, 4) === String(ytdYear));
  const fyCommit = fyOpen.filter(o => (o.forecastCategory || "Not Forecasted") === "Commit");
  const fyStretch = fyOpen.filter(o => (o.forecastCategory || "Not Forecasted") === "Stretch");
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
    const commit = inMonth.filter(o => (o.forecastCategory || "Not Forecasted") === "Commit");
    const stretch = inMonth.filter(o => (o.forecastCategory || "Not Forecasted") === "Stretch");
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

  const chipFilteredAccounts = filteredAccounts
    .filter(a => {
      if (accountFilter === "All") return true;
      const c = accountComputed(a);
      if (accountFilter === "Active") return a.status === "Active";
      if (accountFilter === "High priority") return a.priority === "High";
      if (accountFilter === "ServiceNow") return !!a.flagServiceNow || !!a.snStatus;
      if (accountFilter === "Dynatrace") return !!a.flagDynatrace || !!a.dtStatus;
      if (accountFilter === "Tanium") return !!a.flagTaniumRevList || !!a.flagTaniumPtp || !!a.flagTaniumTarget || !!a.taniumStatus;
      if (accountFilter === "NeuBird AI") return !!a.flagNeuBird || !!a.neubirdStatus;
      if (accountFilter === "Active opportunity") return c.hasActiveOpp;
      if (accountFilter === "Never met") return !c.hasMet;
      return true;
    })
    .sort((a, b) => a.name.localeCompare(b.name));

  const accountGroups = accountGroupBy === "none" ? null : (() => {
    const map = new Map();
    chipFilteredAccounts.forEach(a => {
      const raw = (a[accountGroupBy] || "").toString().trim();
      const key = raw ? normName(raw) : "unassigned";
      const label = raw || `No ${ACCOUNT_GROUP_FIELD_LABEL[accountGroupBy].toLowerCase()}`;
      if (!map.has(key)) map.set(key, { key, label, accounts: [] });
      map.get(key).accounts.push(a);
    });
    const order = ACCOUNT_GROUP_ORDER[accountGroupBy];
    return Array.from(map.values()).sort((x, y) => {
      if (x.key === "unassigned") return 1;
      if (y.key === "unassigned") return -1;
      if (order) {
        const xi = order.indexOf(x.label), yi = order.indexOf(y.label);
        const xr = xi === -1 ? 999 : xi, yr = yi === -1 ? 999 : yi;
        if (xr !== yr) return xr - yr;
      }
      return x.label.localeCompare(y.label);
    });
  })();

  const accountTableHeader = (
    <div style={{ display: "grid", gridTemplateColumns: "20px 1.3fr 0.8fr 0.8fr 0.8fr 0.8fr 0.9fr 1.6fr", padding: "10px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
      <span></span><span>Account</span><span>Territory</span><span>Tier</span><span>Priority</span><span>Status</span><span>CD</span><span>Notes</span>
    </div>
  );

  const renderAccountRow = (a) => {
    const c = accountComputed(a);
    const expanded = expandedAccounts.has(a.id);
    const vendorPill = (status) => status
      ? <Pill color={VENDOR_STATUS_COLOR[status] || C.textMute} bg={VENDOR_STATUS_BG[status] || "#F1F5F9"}>{status}</Pill>
      : <span style={{ color: C.textMute, fontSize: 12.5 }}>—</span>;
    const yn = (v) => v
      ? <Pill color={C.green} bg="#DCFCE7"><Check size={10} style={{ verticalAlign: -1 }} /></Pill>
      : <Pill color={C.textMute} bg="#F1F5F9">—</Pill>;
    const flagList = [
      ["flagServiceNow", "ServiceNow"], ["flagDynatrace", "Dynatrace"], ["flagDtDedicatedProg", "DT - Dedicated Prog"],
      ["flagTaniumRevList", "Tanium Rev List"], ["flagTaniumPtp", "Tanium PTP"], ["flagTaniumTarget", "Tanium Target"],
      ["flagGrafana", "Grafana"], ["flagObservability", "Observability"],
    ].filter(([key]) => a[key]);
    return (
      <div key={a.id}>
        <div onClick={() => toggleAccountExpand(a.id)} style={{
          display: "grid", gridTemplateColumns: "20px 1.3fr 0.8fr 0.8fr 0.8fr 0.8fr 0.9fr 1.6fr", padding: "12px 16px",
          fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center"
        }}>
          <ChevronRight size={14} color={C.textMute} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s ease" }} />
          <span style={{ color: C.text, fontWeight: 600 }}>{a.name}</span>
          <span style={{ color: C.textSoft }}>{a.territory || "—"}</span>
          <span style={{ color: C.textSoft }}>{a.tier || "—"}</span>
          <span>{a.priority ? <Pill color={PRIORITY_COLOR[a.priority] || C.textMute} bg={PRIORITY_BG[a.priority] || "#F1F5F9"}>{a.priority}</Pill> : "—"}</span>
          <span>{a.status ? <Pill color={a.status === "Active" ? C.green : C.textMute} bg={a.status === "Active" ? "#DCFCE7" : "#F1F5F9"}>{a.status}</Pill> : "—"}</span>
          <span style={{ color: C.textSoft }}>{a.cd || "—"}</span>
          <span style={{ color: C.textMute, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{a.notes || "—"}</span>
        </div>
        {expanded && (
          <div style={{ padding: "16px 16px 20px 46px", borderBottom: `1px solid ${C.border}`, background: "#FAFBFC" }}>
            {(() => {
              const acctOpps = data.opportunities.filter(o => o.accountId ? o.accountId === a.id : normName(o.account) === normName(a.name));
              return (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>Opportunities ({acctOpps.length})</div>
                    <button style={{ ...ghostBtn, padding: "4px 9px", fontSize: 11.5 }} onClick={(e) => { e.stopPropagation(); setOppModal({ accountId: a.id, account: a.name }); }}>
                      <Plus size={12} style={{ verticalAlign: -2, marginRight: 4 }} />Add opportunity
                    </button>
                  </div>
                  <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
                    {acctOpps.length === 0 ? (
                      <div style={{ padding: "10px 12px", fontSize: 12.5, color: C.textMute, background: C.card }}>No opportunities linked to this account yet.</div>
                    ) : (
                      <div style={{ overflowX: "auto" }}>
                        <div style={{ minWidth: 860 }}>
                          {oppTableHeader}
                          {acctOpps
                            .sort((x, y) => (x.closeDate || "9999").localeCompare(y.closeDate || "9999"))
                            .map(o => renderOpportunityRow(o))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}

            {(() => {
              const acctTasks = data.tasks.filter(t => t.accountId ? t.accountId === a.id : normName(t.account) === normName(a.name));
              return (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3 }}>Tasks ({acctTasks.length})</div>
                    <button style={{ ...ghostBtn, padding: "4px 9px", fontSize: 11.5 }} onClick={(e) => { e.stopPropagation(); setTaskModal({ accountId: a.id, account: a.name }); }}>
                      <Plus size={12} style={{ verticalAlign: -2, marginRight: 4 }} />Add task
                    </button>
                  </div>
                  <div style={{ border: `1px solid ${C.border}`, borderRadius: 8, overflow: "hidden" }}>
                    {acctTasks.length === 0 ? (
                      <div style={{ padding: "10px 12px", fontSize: 12.5, color: C.textMute, background: C.card }}>No tasks linked to this account yet.</div>
                    ) : (
                      <div style={{ overflowX: "auto" }}>
                        <div style={{ minWidth: 980 }}>
                          {taskTableHeader}
                          {acctTasks.map(t => renderTaskRow(t))}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              );
            })()}

            {(() => {
              const acctCampaigns = data.campaigns.filter(c => c.accountIds.includes(a.id));
              if (acctCampaigns.length === 0) return null;
              return (
                <div style={{ marginBottom: 16 }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>Campaigns ({acctCampaigns.length})</div>
                  <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                    {acctCampaigns.map(c => (
                      <span key={c.id} onClick={(e) => { e.stopPropagation(); setCampModal(c); }} style={{ cursor: "pointer", display: "inline-flex" }}>
                        <Pill color={CAMPAIGN_STATUS_COLOR[c.status]} bg={CAMPAIGN_STATUS_BG[c.status]}>{c.name}</Pill>
                      </span>
                    ))}
                  </div>
                </div>
              );
            })()}

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 14 }}>
              <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                {yn(c.hasActiveOpp)}<span style={{ fontSize: 11.5, color: C.textMute, marginRight: 10 }}>Active opportunity</span>
                {yn(c.hasMet)}<span style={{ fontSize: 11.5, color: C.textMute }}>Met with them</span>
              </div>
              <button style={{ ...ghostBtn, padding: "5px 10px", fontSize: 11.5 }} onClick={(e) => { e.stopPropagation(); setAccModal(a); }}>
                <Pencil size={12} style={{ verticalAlign: -2, marginRight: 5 }} />Edit full details
              </button>
            </div>

            {flagList.length > 0 && (
              <div style={{ marginBottom: 14 }}>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 6 }}>Practice flags</div>
                <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
                  {flagList.map(([key, label]) => <Pill key={key} color={C.text} bg="#F1F5F9">{label}</Pill>)}
                </div>
              </div>
            )}

            <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
              {[
                { label: "ServiceNow", color: "#16A34A", bg: "#F0FDF4", rep: a.snRep, status: a.snStatus, priority: a.snPriority, contacts: a.snContacts, notes: a.snNotes },
                { label: "Dynatrace", color: "#3B82F6", bg: "#EFF6FF", rep: a.dtRep, psm: a.dtPsm, status: a.dtStatus, priority: a.dtPriority, contacts: a.dtContacts, notes: a.dtNotes },
                { label: "Tanium", color: "#DC2626", bg: "#FEF2F2", rep: a.taniumRep, status: a.taniumStatus, priority: a.taniumPriority, contacts: a.taniumContacts, notes: a.taniumNotes },
                { label: "NeuBird AI", color: BURNT_ORANGE, bg: "#FFF3EA", rep: a.neubirdRep, status: a.neubirdStatus, priority: a.neubirdPriority, contacts: a.neubirdContacts, notes: a.neubirdNotes },
              ].map(v => (
                <div key={v.label} style={{ background: v.bg, border: `1px solid ${v.color}33`, borderRadius: 8, padding: 14 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: v.color, marginBottom: 10 }}>{v.label}</div>
                  <div style={{ display: "grid", gridTemplateColumns: v.psm !== undefined ? "1fr 1fr 1fr 1fr" : "1fr 1fr 1fr", gap: 10, marginBottom: 10 }}>
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>Rep</div>
                      <div style={{ fontSize: 12.5, color: C.text }}>{v.rep || "—"}</div>
                    </div>
                    {v.psm !== undefined && (
                      <div>
                        <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>PSM</div>
                        <div style={{ fontSize: 12.5, color: C.text }}>{v.psm || "—"}</div>
                      </div>
                    )}
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>Status</div>
                      <div>{vendorPill(v.status)}</div>
                    </div>
                    <div>
                      <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>Priority</div>
                      <div style={{ fontSize: 12.5, color: C.text }}>{v.priority || "—"}</div>
                    </div>
                  </div>
                  <div style={{ marginBottom: 8 }}>
                    <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>Contacts</div>
                    <div style={{ fontSize: 12.5, color: C.text, whiteSpace: "pre-wrap" }}>{v.contacts || "—"}</div>
                  </div>
                  <div>
                    <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, marginBottom: 2 }}>Notes</div>
                    <div style={{ fontSize: 12.5, color: C.text, whiteSpace: "pre-wrap" }}>{v.notes || "—"}</div>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}
      </div>
    );
  };

  const taskTableHeader = (
    <div style={{
      display: "grid", gridTemplateColumns: "20px 1.3fr 1fr 0.7fr 0.8fr 0.8fr 0.9fr 1fr",
      padding: "11px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.7,
      borderBottom: `1px solid ${C.border}`, fontWeight: 700, background: "#F8FAFC"
    }}>
      <span></span><span>Name</span><span>Account</span><span>CD</span><span>Owner</span><span>Priority</span><span>Due date</span><span>Status</span>
    </div>
  );

  const renderTaskRow = (t) => {
    const expanded = expandedTasks.has(t.id);
    const overdue = t.status !== "Done" && t.dueDate && daysUntil(t.dueDate) < 0;
    const rowAccent = PRIORITY_COLOR[t.priority] || C.textMute;
    const hasNotes = !!(t.notes && t.notes.trim());
    return (
      <div key={t.id}>
        <div onClick={() => toggleTaskExpand(t.id)} style={{
          display: "grid", gridTemplateColumns: "20px 1.3fr 1fr 0.7fr 0.8fr 0.8fr 0.9fr 1fr",
          padding: "14px 16px", borderBottom: hasNotes ? "none" : `1px solid ${C.border}`, borderLeft: `3px solid ${rowAccent}`,
          cursor: "pointer", alignItems: "center", gap: 6, background: C.card, transition: "background 0.1s ease"
        }}>
          <ChevronRight size={14} color={C.textMute} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s ease" }} />
          <span style={{
            fontSize: 14, fontWeight: 600, color: t.status === "Done" ? C.textMute : C.text, letterSpacing: -0.1, lineHeight: 1.3,
            textDecoration: t.status === "Done" ? "line-through" : "none"
          }}>{t.title}</span>
          <span style={{ fontSize: 12.5, fontWeight: 500, color: C.text, lineHeight: 1.3 }}>{t.account || "—"}</span>
          <span style={{ fontSize: 12.5, fontWeight: 500, color: C.text, lineHeight: 1.3 }}>{t.cd || "—"}</span>
          <span style={{ fontSize: 12.5, fontWeight: 500, color: C.text, lineHeight: 1.3 }}>{t.rep || "—"}</span>
          <span><Pill color={PRIORITY_COLOR[t.priority] || C.textMute} bg={PRIORITY_BG[t.priority] || "#F1F5F9"}>{t.priority}</Pill></span>
          <span style={{ fontSize: 12.5, color: overdue ? C.red : C.text, fontWeight: overdue ? 700 : 500, lineHeight: 1.3 }}>{fmtDate(t.dueDate)}</span>
          <span onClick={e => e.stopPropagation()}>
            <select
              value={t.status}
              onChange={e => upsert("tasks", { ...t, status: e.target.value })}
              style={{
                fontSize: 11, fontWeight: 700, padding: "5px 9px", borderRadius: 999, border: "none", cursor: "pointer",
                color: TASK_STATUS_COLOR[t.status], background: TASK_STATUS_BG[t.status], letterSpacing: 0.2
              }}
            >
              {TASK_STATUSES.map(s => <option key={s} value={s}>{s}</option>)}
            </select>
          </span>
        </div>
        {hasNotes && (
          <div onClick={() => toggleTaskExpand(t.id)} style={{
            padding: "0 16px 13px 42px", borderBottom: `1px solid ${C.border}`, borderLeft: `3px solid ${rowAccent}`,
            background: C.card, cursor: "pointer"
          }}>
            <span style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginRight: 7 }}>Notes</span>
            <span style={{ fontSize: 12.5, fontWeight: 400, color: C.textSoft, lineHeight: 1.5, whiteSpace: "pre-wrap" }}>{t.notes}</span>
          </div>
        )}
        {expanded && (
          <div style={{ padding: "17px 16px 19px 40px", borderBottom: `1px solid ${C.border}`, borderLeft: `3px solid ${rowAccent}`, background: "#F8FAFC" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 14 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr 1fr", gap: 16, flex: 1 }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>SSE / SME</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text, lineHeight: 1.35 }}>{t.sseSme || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>Partner Contact</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text, lineHeight: 1.35 }}>{t.partnerContact || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>Partner Engaged</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text, lineHeight: 1.35 }}>{t.partnerEngaged || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>Opportunity</div>
                  <div>
                    {(() => {
                      const linkedOpp = t.oppId ? data.opportunities.find(o => o.id === t.oppId) : null;
                      if (linkedOpp) {
                        return (
                          <span
                            onClick={(e) => { e.stopPropagation(); setOppModal(linkedOpp); }}
                            style={{ color: C.blue, fontWeight: 600, fontSize: 13, cursor: "pointer", textDecoration: "underline" }}
                            title="Open this opportunity"
                          >
                            {linkedOpp.name}
                          </span>
                        );
                      }
                      if (t.isOpp) return <Pill color={C.blue} bg="#DBEAFE">Opp</Pill>;
                      return <span style={{ color: C.textMute, fontSize: 13, fontWeight: 500 }}>—</span>;
                    })()}
                  </div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.5, marginBottom: 4 }}>Deal Reg?</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text, lineHeight: 1.35 }}>{t.dealReg || "—"}</div>
                </div>
              </div>
              <button style={{ ...ghostBtn, padding: "6px 12px", fontSize: 12, fontWeight: 600, flexShrink: 0 }} onClick={(e) => { e.stopPropagation(); setTaskModal(t); }}>
                <Pencil size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Edit full details
              </button>
            </div>
          </div>
        )}
      </div>
    );
  };

  const oppTableHeader = (
    <div style={{
      display: "grid", gridTemplateColumns: "20px 1.7fr 1.1fr 0.9fr 1fr 0.9fr 0.8fr",
      padding: "10px 16px", fontSize: 10.5, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.5,
      borderBottom: `1px solid ${C.border}`, fontWeight: 700, background: "#F8FAFC"
    }}>
      <span></span><span>Name</span><span>Account</span><span>Stage</span><span>Commissionable Margin</span><span>Close date</span><span>Forecast</span>
    </div>
  );

  const renderOpportunityRow = (o) => {
    const expanded = expandedOpps.has(o.id);
    const overdue = o.closeDate && daysUntil(o.closeDate) < 0 && OPEN_STAGES.includes(o.stage);
    const rowAccent = STAGE_COLOR[o.stage] || C.textMute;
    const linkedCampaign = o.campaignId ? data.campaigns.find(c => c.id === o.campaignId) : null;
    return (
      <div key={o.id}>
        <div onClick={() => toggleOppExpand(o.id)} style={{
          display: "grid", gridTemplateColumns: "20px 1.7fr 1.1fr 0.9fr 1fr 0.9fr 0.8fr",
          padding: "13px 16px", borderBottom: `1px solid ${C.border}`, borderLeft: `3px solid ${rowAccent}`,
          cursor: "pointer", alignItems: "center", gap: 6, background: C.card, opacity: o.sfStatus === "removed" ? 0.55 : 1
        }}>
          <ChevronRight size={14} color={C.textMute} style={{ transform: expanded ? "rotate(90deg)" : "none", transition: "transform 0.15s ease" }} />
          <span style={{ fontSize: 13.5, fontWeight: 600, color: C.text, display: "flex", alignItems: "center", gap: 6 }}>
            {o.name}{o.sfId && <FileSpreadsheet size={11} color={C.textMute} />}
          </span>
          <span style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{o.account}</span>
          <span>
            {o.sfStatus === "removed"
              ? <Pill color={C.amber} bg={C.amberBg}>Removed</Pill>
              : <Pill color={STAGE_COLOR[o.stage]} bg={STAGE_COLOR[o.stage] + "1A"}>{o.stage}</Pill>}
          </span>
          <span style={{ fontSize: 13, fontWeight: 700, color: C.text }}>{fmtMoneyShort(oppTotal(o))}</span>
          <span style={{ fontSize: 13, color: overdue ? C.red : C.text, fontWeight: overdue ? 700 : 500 }}>{fmtDate(o.closeDate)}</span>
          <span>{OPEN_STAGES.includes(o.stage) && <Pill color={FORECAST_COLOR[o.forecastCategory || "Not Forecasted"]} bg={FORECAST_COLOR[o.forecastCategory || "Not Forecasted"] + "1A"}>{o.forecastCategory || "Not Forecasted"}</Pill>}</span>
        </div>
        {expanded && (
          <div style={{ padding: "16px 16px 18px 40px", borderBottom: `1px solid ${C.border}`, borderLeft: `3px solid ${rowAccent}`, background: "#F8FAFC" }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 12, marginBottom: 14 }}>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr", gap: 14, flex: 1 }}>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 3 }}>Owner</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{o.rep || "—"}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 3 }}>Total value</div>
                  <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{fmtMoneyShort(oppTotal(o))}</div>
                </div>
                <div>
                  <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 3 }}>Campaign</div>
                  <div>
                    {linkedCampaign
                      ? <span onClick={(e) => { e.stopPropagation(); setCampModal(linkedCampaign); }} style={{ color: C.blue, fontWeight: 600, fontSize: 13, cursor: "pointer", textDecoration: "underline" }}>{linkedCampaign.name}</span>
                      : <span style={{ color: C.textMute, fontSize: 13, fontWeight: 500 }}>—</span>}
                  </div>
                </div>
              </div>
              <button style={{ ...ghostBtn, padding: "6px 12px", fontSize: 12, flexShrink: 0 }} onClick={(e) => { e.stopPropagation(); setOppModal(o); }}>
                <Pencil size={13} style={{ verticalAlign: -2, marginRight: 5 }} />Edit full details
              </button>
            </div>
            <div style={{ marginBottom: 10 }}>
              <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 3 }}>Next step</div>
              <div style={{ fontSize: 13, fontWeight: 500, color: C.text }}>{o.nextStep || "—"}</div>
            </div>
            <div>
              <div style={{ fontSize: 10, fontWeight: 700, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, marginBottom: 3 }}>Notes</div>
              <div style={{ fontSize: 13, fontWeight: 400, color: C.text, whiteSpace: "pre-wrap" }}>{o.notes || "—"}</div>
            </div>
          </div>
        )}
      </div>
    );
  };

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
    const opps = openOppsInForecastPeriod.filter(o => (o.forecastCategory || "Not Forecasted") === cat).sort((a, b) => (a.closeDate || "9999").localeCompare(b.closeDate || "9999"));
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
  const taskPartnerOpts = distinctTaskVals("partnerContact");
  const taskDealRegOpts = distinctTaskVals("dealReg");

  const scopedTasks = repScopedTasks
    .filter(t => taskFilter === "All" ? true : taskFilter === "Active" ? t.status !== "Done" : t.status === taskFilter)
    .filter(t => taskFieldFilters.priority === "All" || t.priority === taskFieldFilters.priority)
    .filter(t => taskFieldFilters.account === "All" || normName(t.account) === normName(taskFieldFilters.account))
    .filter(t => taskFieldFilters.cd === "All" || (t.cd || "") === taskFieldFilters.cd)
    .filter(t => taskFieldFilters.owner === "All" || t.rep === taskFieldFilters.owner)
    .filter(t => taskFieldFilters.sseSme === "All" || (t.sseSme || "") === taskFieldFilters.sseSme)
    .filter(t => taskFieldFilters.partnerContact === "All" || (t.partnerContact || "") === taskFieldFilters.partnerContact)
    .filter(t => taskFieldFilters.partnerEngaged === "All" || (t.partnerEngaged || "") === taskFieldFilters.partnerEngaged)
    .filter(t => taskFieldFilters.dealReg === "All" || (t.dealReg || "") === taskFieldFilters.dealReg)
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
    })
    .filter(t => {
      if (!taskSearch.trim()) return true;
      const q = taskSearch.trim().toLowerCase();
      return (t.title || "").toLowerCase().includes(q) || (t.account || "").toLowerCase().includes(q);
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
  const groupedByAccount = unassignedTaskGroup ? [unassignedTaskGroup, ...acctTaskGroups] : acctTaskGroups;

  const groupedByPriority = ["Critical", "High", "Moderate", "Low"]
    .map(p => {
      const tasks = scopedTasks
        .filter(t => (t.priority || "Moderate") === p)
        .slice()
        .sort((a, b) => (a.dueDate || "9999").localeCompare(b.dueDate || "9999") || (a.account || "").localeCompare(b.account || ""));
      const acctCount = new Set(tasks.map(t => normName(t.account || "unassigned"))).size;
      return {
        key: p, title: p, subtitle: `${acctCount} account${acctCount === 1 ? "" : "s"}`,
        tasks, color: PRIORITY_COLOR[p],
      };
    })
    .filter(g => g.tasks.length > 0);

  const noCdTaskGroup = (() => {
    const list = byPriority(scopedTasks.filter(t => !t.cd || !t.cd.trim()));
    return list.length > 0 ? { key: "no-cd", title: "No CD assigned", subtitle: "", tasks: list } : null;
  })();
  const cdTaskMap = new Map();
  scopedTasks.filter(t => t.cd && t.cd.trim()).forEach(t => {
    const key = normName(t.cd);
    if (!cdTaskMap.has(key)) {
      const acctSet = new Set();
      cdTaskMap.set(key, { key, title: t.cd.trim(), acctSet, tasks: [] });
    }
    cdTaskMap.get(key).acctSet.add(normName(t.account || ""));
    cdTaskMap.get(key).tasks.push(t);
  });
  const cdTaskGroups = Array.from(cdTaskMap.values())
    .map(g => {
      const acctCount = Array.from(g.acctSet).filter(Boolean).length;
      return { key: g.key, title: g.title, subtitle: `${acctCount} account${acctCount === 1 ? "" : "s"}`, tasks: byPriority(g.tasks) };
    })
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
  const groupedByCd = noCdTaskGroup ? [noCdTaskGroup, ...cdTaskGroups] : cdTaskGroups;

  const groupedByNone = byPriority(scopedTasks).length > 0 ? [{ key: "all", title: "All tasks", subtitle: "", tasks: byPriority(scopedTasks) }] : [];

  const taskGroups = taskGroupBy === "priority" ? groupedByPriority : taskGroupBy === "cd" ? groupedByCd : taskGroupBy === "none" ? groupedByNone : groupedByAccount;


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
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
                <div>
                  <h4 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: C.text }}>Pipeline</h4>
                  <p style={{ margin: "2px 0 0", fontSize: 12, color: C.textMute }}>Current state of the open book of business.</p>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={ghostBtn} onClick={() => setOppImportOpen(true)}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Import from Excel</button>
                  <button style={primaryBtn} onClick={() => setOppModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add opportunity</button>
                </div>
              </div>

              {/* Executive KPI strip */}
              <div style={{ display: "grid", gridTemplateColumns: "repeat(5, 1fr)", gap: 12, marginBottom: 16 }}>
                <Card style={{ padding: "14px 16px" }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Open pipeline</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(openPipelineValue)}</div>
                  <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>{openOpps.length} open {openOpps.length === 1 ? "deal" : "deals"}</div>
                </Card>
                <Card style={{ padding: "14px 16px" }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Weighted pipeline</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(weightedPipeline)}</div>
                  <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>Probability-adjusted by stage</div>
                </Card>
                <Card style={{ padding: "14px 16px" }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Coverage</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: coverageTone === "green" ? C.green : coverageTone === "amber" ? C.amber : C.red }}>{coverage.toFixed(1)}x</div>
                  <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>Target {coverageTarget.toFixed(1)}x of goal remaining</div>
                </Card>
                <Card style={{ padding: "14px 16px" }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Avg deal size</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(avgDealSize)}</div>
                  <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>Across open deals</div>
                </Card>
                <Card style={{ padding: "14px 16px" }}>
                  <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Win rate</div>
                  <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{winRate === null ? "—" : `${winRate.toFixed(0)}%`}</div>
                  <div style={{ fontSize: 11, color: C.textMute, marginTop: 2 }}>{winRateDenom > 0 ? `${closedWonAll.length} won / ${closedLostAll.length} lost` : "No closed deals yet"}</div>
                </Card>
              </div>

              {/* Stage funnel */}
              <Card style={{ marginBottom: 16 }}>
                <h5 style={{ margin: "0 0 14px", fontSize: 12.5, fontWeight: 700, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>Funnel by stage</h5>
                {(() => {
                  const funnelStages = OPEN_STAGES;
                  const maxSum = Math.max(1, ...allStageSummary.filter(g => funnelStages.includes(g.key)).map(g => g.sum));
                  return (
                    <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                      {allStageSummary.filter(g => funnelStages.includes(g.key)).map(g => (
                        <div key={g.key} style={{ display: "flex", alignItems: "center", gap: 10 }}>
                          <div style={{ width: 110, fontSize: 12, fontWeight: 600, color: C.text, flexShrink: 0 }}>{g.label}</div>
                          <div style={{ flex: 1, background: "#F1F5F9", borderRadius: 6, height: 26, position: "relative", overflow: "hidden" }}>
                            <div style={{
                              width: `${Math.max((g.sum / maxSum) * 100, g.sum > 0 ? 4 : 0)}%`, height: "100%", background: g.color,
                              borderRadius: 6, display: "flex", alignItems: "center", justifyContent: "flex-end", paddingRight: 8, minWidth: g.sum > 0 ? 46 : 0, transition: "width 0.2s ease"
                            }}>
                              {g.sum > 0 && <span style={{ fontSize: 11, fontWeight: 700, color: "#fff", whiteSpace: "nowrap" }}>{fmtMoneyShort(g.sum)}</span>}
                            </div>
                          </div>
                          <div style={{ width: 70, fontSize: 11.5, color: C.textMute, textAlign: "right", flexShrink: 0 }}>{g.count} {g.count === 1 ? "deal" : "deals"}</div>
                        </div>
                      ))}
                    </div>
                  );
                })()}
              </Card>

              {/* Top open deals */}
              {topOpenOpps.length > 0 && (
                <Card style={{ marginBottom: 22, padding: 0, overflow: "hidden" }}>
                  <h5 style={{ margin: 0, padding: "14px 16px 10px", fontSize: 12.5, fontWeight: 700, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>Top open deals</h5>
                  {topOpenOpps.map((o, i) => (
                    <div key={o.id} onClick={() => setOppModal(o)} style={{
                      display: "flex", alignItems: "center", gap: 12, padding: "10px 16px",
                      borderTop: `1px solid ${C.border}`, cursor: "pointer"
                    }}>
                      <span style={{ fontSize: 12, fontWeight: 700, color: C.textMute, width: 16 }}>{i + 1}</span>
                      <span style={{ flex: 1, minWidth: 0 }}>
                        <div style={{ fontSize: 13, fontWeight: 600, color: C.text, overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" }}>{o.name}</div>
                        <div style={{ fontSize: 11.5, color: C.textMute }}>{o.account}</div>
                      </span>
                      <Pill color={STAGE_COLOR[o.stage]} bg={STAGE_COLOR[o.stage] + "1A"}>{o.stage}</Pill>
                      <span style={{ fontSize: 13, fontWeight: 700, color: C.text, minWidth: 80, textAlign: "right" }}>{fmtMoneyShort(oppTotal(o))}</span>
                    </div>
                  ))}
                </Card>
              )}

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
                    {oppTableHeader}
                    {g.opps.map(o => renderOpportunityRow(o))}
                  </Card>
                </div>
              ))}
            </div>
          )}

          {tab === "forecast" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 6, flexWrap: "wrap", gap: 10 }}>
                <div>
                  <h4 style={{ margin: 0, fontSize: 16, fontWeight: 800, color: C.text }}>Forecast</h4>
                  <p style={{ margin: "2px 0 0", fontSize: 12, color: C.textMute }}>{forecastPeriodLabel} — what's committed, what's stretch, and what's still uncalled.</p>
                </div>
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

              {(() => {
                const committedSum = forecastGroups.find(g => g.key === "Commit")?.sum || 0;
                const calledSum = forecastGroups.filter(g => g.key !== "Not Forecasted").reduce((s, g) => s + g.sum, 0);
                const uncalledSum = forecastGroups.find(g => g.key === "Not Forecasted")?.sum || 0;
                const periodGoal = forecastPeriodType === "quarter"
                  ? data.goals.filter(g => g.period === forecastPeriodValue && repScope.includes(g.rep))
                      .reduce((s, g) => s + CATEGORIES.reduce((cs, c) => cs + (Number(g[c.key + "Target"]) || 0), 0), 0)
                  : null;
                const hasPeriodGoal = periodGoal !== null && periodGoal > 0;
                const commitPctOfGoal = hasPeriodGoal ? (committedSum / periodGoal) * 100 : null;
                const totalMix = Math.max(1, forecastGroups.reduce((s, g) => s + g.sum, 0));

                return (
                  <>
                    <div style={{ display: "grid", gridTemplateColumns: hasPeriodGoal ? "repeat(4, 1fr)" : "repeat(3, 1fr)", gap: 12, marginBottom: 16 }}>
                      <Card style={{ padding: "14px 16px" }}>
                        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Called (Commit + Gut + Stretch)</div>
                        <div style={{ fontSize: 22, fontWeight: 800, color: C.text }}>{fmtMoneyShort(calledSum)}</div>
                      </Card>
                      <Card style={{ padding: "14px 16px" }}>
                        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Commit only</div>
                        <div style={{ fontSize: 22, fontWeight: 800, color: FORECAST_COLOR.Commit }}>{fmtMoneyShort(committedSum)}</div>
                      </Card>
                      <Card style={{ padding: "14px 16px" }}>
                        <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Not yet forecasted</div>
                        <div style={{ fontSize: 22, fontWeight: 800, color: C.textMute }}>{fmtMoneyShort(uncalledSum)}</div>
                      </Card>
                      {hasPeriodGoal && (
                        <Card style={{ padding: "14px 16px" }}>
                          <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>Commit vs. quarter goal</div>
                          <div style={{ fontSize: 22, fontWeight: 800, color: commitPctOfGoal >= 100 ? C.green : commitPctOfGoal >= 75 ? C.amber : C.red }}>{commitPctOfGoal.toFixed(0)}%</div>
                        </Card>
                      )}
                    </div>

                    <Card style={{ marginBottom: 16 }}>
                      <h5 style={{ margin: "0 0 12px", fontSize: 12.5, fontWeight: 700, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>Confidence mix</h5>
                      <div style={{ display: "flex", height: 36, borderRadius: 8, overflow: "hidden", gap: 2, marginBottom: 10 }}>
                        {forecastGroups.filter(g => g.sum > 0).map(g => (
                          <div key={g.key} style={{
                            width: `${(g.sum / totalMix) * 100}%`, background: g.color,
                            display: "flex", alignItems: "center", justifyContent: "center", minWidth: 44
                          }}>
                            <span style={{ fontSize: 10.5, fontWeight: 700, color: "#fff" }}>{fmtMoneyShort(g.sum)}</span>
                          </div>
                        ))}
                      </div>
                      <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                        {forecastGroups.map(g => (
                          <span key={g.key} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: C.textSoft }}>
                            <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                            {g.label} {fmtMoneyShort(g.sum)} ({totalMix > 1 ? ((g.sum / totalMix) * 100).toFixed(0) : 0}%)
                          </span>
                        ))}
                      </div>
                    </Card>

                    {hasPeriodGoal && (
                      <Card style={{ marginBottom: 16 }}>
                        <h5 style={{ margin: "0 0 12px", fontSize: 12.5, fontWeight: 700, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>Goal attainment breakdown</h5>
                        {[["Goal", periodGoal, C.textMute], ["Commit", committedSum, FORECAST_COLOR.Commit], ["Called (Commit + Gut + Stretch)", calledSum, C.blue]].map(([label, val, color]) => (
                          <div key={label} style={{ marginBottom: 10 }}>
                            <div style={{ display: "flex", justifyContent: "space-between", fontSize: 11.5, color: C.textSoft, marginBottom: 3 }}>
                              <span>{label}</span><span style={{ fontWeight: 700, color: C.text }}>{fmtMoneyShort(val)}</span>
                            </div>
                            <div style={{ height: 10, background: "#F1F5F9", borderRadius: 999, overflow: "hidden" }}>
                              <div style={{ width: `${Math.min((val / periodGoal) * 100, 100)}%`, height: "100%", background: color, borderRadius: 999 }} />
                            </div>
                          </div>
                        ))}
                      </Card>
                    )}
                  </>
                );
              })()}

              {forecastGroups.map(g => (
                <div key={g.key} style={{ marginBottom: 22 }}>
                  <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                    <span style={{ width: 8, height: 8, borderRadius: 2, background: g.color, display: "inline-block" }} />
                    <h4 style={{ margin: 0, fontSize: 13, fontWeight: 700, color: C.text }}>{g.label}</h4>
                    <span style={{ fontSize: 11.5, color: C.textMute }}>{g.opps.length} {g.opps.length === 1 ? "deal" : "deals"} · {fmtMoneyShort(g.sum)}</span>
                  </div>
                  <Card style={{ padding: 0, overflow: "hidden" }}>
                    <div style={{ display: "grid", gridTemplateColumns: "2fr 1.3fr 0.9fr 1fr 0.8fr", padding: "9px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700 }}>
                      <span>Opportunity</span><span>Account</span><span>Stage</span><span>Commissionable Margin</span><span>Close date</span>
                    </div>
                    {g.opps.map(o => (
                      <div key={o.id} onClick={() => setOppModal(o)} style={{
                        display: "grid", gridTemplateColumns: "2fr 1.3fr 0.9fr 1fr 0.8fr", padding: "11px 16px",
                        fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center"
                      }}>
                        <span style={{ color: C.text, fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                          {o.name}{o.sfId && <FileSpreadsheet size={11} color={C.textMute} />}
                        </span>
                        <span style={{ color: C.textSoft }}>{o.account}</span>
                        <span><Pill color={STAGE_COLOR[o.stage]} bg={STAGE_COLOR[o.stage] + "1A"}>{o.stage}</Pill></span>
                        <span style={{ color: C.text, fontWeight: 700, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(oppTotal(o))}</span>
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
              <div style={{ display: "grid", gridTemplateColumns: "repeat(7, 1fr)", gap: 12, marginBottom: 16 }}>
                {[
                  { label: "Total accounts", value: filteredAccounts.length, color: C.text },
                  { label: "Active", value: filteredAccounts.filter(a => a.status === "Active").length, color: C.green },
                  { label: "High priority", value: filteredAccounts.filter(a => a.priority === "High").length, color: C.red },
                  { label: "ServiceNow customers", value: filteredAccounts.filter(a => a.snStatus === "Customer").length, color: C.blue },
                  { label: "Dynatrace customers", value: filteredAccounts.filter(a => a.dtStatus === "Customer").length, color: C.purple },
                  { label: "Tanium customers", value: filteredAccounts.filter(a => a.taniumStatus === "Customer").length, color: C.amber },
                  { label: "NeuBird customers", value: filteredAccounts.filter(a => a.neubirdStatus === "Customer").length, color: BURNT_ORANGE },
                ].map(s => (
                  <Card key={s.label} style={{ padding: "12px 14px" }}>
                    <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute, marginBottom: 6 }}>{s.label}</div>
                    <div style={{ fontSize: 22, fontWeight: 800, color: s.color }}>{s.value}</div>
                  </Card>
                ))}
              </div>

              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
                  {["All", "Active", "High priority", "ServiceNow", "Dynatrace", "Tanium", "NeuBird AI", "Active opportunity", "Never met"].map(s => (
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

              <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16 }}>
                <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>Group by</span>
                <div style={{ display: "flex", gap: 6 }}>
                  {[["none", "None"], ["cd", "CD"], ["status", "Status"], ["territory", "Territory"], ["priority", "Priority"], ["tier", "Tier"]].map(([val, label]) => (
                    <div key={val} onClick={() => setAccountGroupBy(val)} style={{
                      fontSize: 12, padding: "6px 12px", borderRadius: 999, cursor: "pointer", fontWeight: 600,
                      background: accountGroupBy === val ? C.sidebarActiveBg : C.card,
                      color: accountGroupBy === val ? C.sidebarActiveText : C.textSoft, border: `1px solid ${accountGroupBy === val ? "#C7D2FE" : C.border}`
                    }}>{label}</div>
                  ))}
                </div>
              </div>

              {accountGroupBy === "none" ? (
                <Card style={{ padding: 0, overflow: "hidden" }}>
                  {accountTableHeader}
                  {chipFilteredAccounts.map(a => renderAccountRow(a))}
                  {chipFilteredAccounts.length === 0 && (
                    <div style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No accounts yet. Import your account list or add one manually.</div>
                  )}
                </Card>
              ) : (
                <>
                  {accountGroups.map(g => (
                    <div key={g.key} style={{ marginBottom: 16 }}>
                      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                        <h4 style={{ margin: 0, fontSize: 13, fontWeight: 700, color: C.text }}>{g.label}</h4>
                        <span style={{ fontSize: 11.5, color: C.textMute }}>{g.accounts.length} account{g.accounts.length === 1 ? "" : "s"}</span>
                      </div>
                      <Card style={{ padding: 0, overflow: "hidden" }}>
                        {accountTableHeader}
                        {g.accounts.map(a => renderAccountRow(a))}
                      </Card>
                    </div>
                  ))}
                  {accountGroups.length === 0 && (
                    <Card style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No accounts yet. Import your account list or add one manually.</Card>
                  )}
                </>
              )}
            </div>
          )}

          {tab === "gtro" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>GTRO</h4>
                  <p style={{ margin: "2px 0 0", fontSize: 12, color: C.textMute }}>Every account, spreadsheet-style. Click MD Priority, Ahead Strategic, or Met the Customer to toggle. CD stays in sync with the Accounts tab.</p>
                </div>
              </div>
              <Card style={{ padding: 0, overflow: "hidden" }}>
                <div style={{
                  display: "grid", gridTemplateColumns: "1.8fr 1fr 1fr 1fr 1fr 1fr",
                  fontSize: 11, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, fontWeight: 700
                }}>
                  {["Account", "MD Priority", "Ahead Strategic", "CD", "Met the Customer", "Active IO Opp"].map((h, i) => (
                    <div key={h} style={{ padding: "10px 14px", borderBottom: `1px solid ${C.border}`, borderRight: i < 5 ? `1px solid ${C.border}` : "none", background: "#F8FAFC" }}>{h}</div>
                  ))}
                </div>
                {filteredAccounts.slice().sort((a, b) => a.name.localeCompare(b.name)).map((a, idx) => {
                  const met = !!a.metCustomer;
                  const activeOpp = a.status === "Active";
                  const cellBase = { padding: "10px 14px", fontSize: 12.5, borderRight: `1px solid ${C.border}`, display: "flex", alignItems: "center" };
                  const rowBg = idx % 2 === 1 ? "#FAFBFC" : C.card;
                  const REDWINGS_RED = "#CE1126";
                  return (
                    <div key={a.id} style={{ display: "grid", gridTemplateColumns: "1.8fr 1fr 1fr 1fr 1fr 1fr", background: rowBg }}>
                      <div style={{ ...cellBase, fontWeight: 600, color: C.text, borderBottom: `1px solid ${C.border}` }}>{a.name}</div>
                      <div
                        onClick={() => upsert("accounts", { ...a, mdPriority: !a.mdPriority })}
                        style={{ ...cellBase, borderBottom: `1px solid ${C.border}`, cursor: "pointer", color: C.text }}
                      >
                        {a.mdPriority ? "Yes" : ""}
                      </div>
                      <div
                        onClick={() => upsert("accounts", { ...a, strategic: !a.strategic })}
                        style={{ ...cellBase, borderBottom: `1px solid ${C.border}`, cursor: "pointer", color: C.text }}
                      >
                        {a.strategic ? "Yes" : ""}
                      </div>
                      <div style={{ ...cellBase, color: C.textSoft, borderBottom: `1px solid ${C.border}` }}>{a.cd || "—"}</div>
                      <div
                        onClick={() => upsert("accounts", { ...a, metCustomer: !a.metCustomer })}
                        style={{
                          ...cellBase, borderBottom: `1px solid ${C.border}`, cursor: "pointer",
                          background: met ? rowBg : REDWINGS_RED, color: met ? C.text : "#FFFFFF", fontWeight: met ? 400 : 700
                        }}>
                        {met ? "Yes" : "No"}
                      </div>
                      <div style={{
                        padding: "10px 14px", fontSize: 12.5, display: "flex", alignItems: "center", borderBottom: `1px solid ${C.border}`,
                        background: activeOpp ? rowBg : REDWINGS_RED, color: activeOpp ? C.green : "#FFFFFF", fontWeight: 700
                      }}>
                        {activeOpp ? "Active" : ""}
                      </div>
                    </div>
                  );
                })}
                {filteredAccounts.length === 0 && (
                  <div style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13 }}>No accounts yet — add some in the Accounts tab first.</div>
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
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 14 }}>
                <div>
                  <h4 style={{ margin: 0, fontSize: 14, fontWeight: 700, color: C.text }}>Campaigns</h4>
                  <p style={{ margin: "2px 0 0", fontSize: 12, color: C.textMute }}>What your BDR is running, at which accounts, and what it's produced.</p>
                </div>
                <button style={primaryBtn} onClick={() => setCampModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />New campaign</button>
              </div>

              {data.campaigns.length === 0 && (
                <Card style={{ padding: 30, textAlign: "center", color: C.textMute, fontSize: 13, marginBottom: 24 }}>
                  No campaigns yet. Create one to start tracking what's being run where — and which accounts turn into real pipeline.
                </Card>
              )}

              {data.campaigns.slice().sort((a, b) => (b.startDate || "").localeCompare(a.startDate || "")).map(c => {
                const attributed = data.opportunities.filter(o => o.campaignId === c.id);
                const won = attributed.filter(o => o.stage === "Closed Won");
                const totalValue = attributed.reduce((s, o) => s + oppTotal(o), 0);
                const accountsHit = new Set(attributed.map(o => o.accountId)).size;
                return (
                  <Card key={c.id} onClick={() => setCampModal(c)} style={{ padding: "16px 18px", marginBottom: 12, cursor: "pointer" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 10 }}>
                      <div>
                        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                          <span style={{ fontSize: 14, fontWeight: 700, color: C.text }}>{c.name}</span>
                          <Pill color={CAMPAIGN_STATUS_COLOR[c.status]} bg={CAMPAIGN_STATUS_BG[c.status]}>{c.status}</Pill>
                        </div>
                        <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 3 }}>
                          {c.bdr ? `BDR: ${c.bdr}` : "No BDR assigned"}
                          {(c.startDate || c.endDate) && ` · ${fmtDate(c.startDate)} – ${fmtDate(c.endDate)}`}
                        </div>
                      </div>
                      <button style={{ ...ghostBtn, padding: "5px 10px", fontSize: 11.5, flexShrink: 0 }} onClick={(e) => { e.stopPropagation(); setCampModal(c); }}>
                        <Pencil size={12} style={{ verticalAlign: -2, marginRight: 5 }} />Edit
                      </button>
                    </div>
                    <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 10 }}>
                      <div>
                        <div style={{ fontSize: 17, fontWeight: 800, color: C.text }}>{c.accountIds.length}</div>
                        <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Accounts targeted</div>
                      </div>
                      <div>
                        <div style={{ fontSize: 17, fontWeight: 800, color: C.text }}>{attributed.length} <span style={{ fontSize: 12, fontWeight: 500, color: C.textMute }}>at {accountsHit}</span></div>
                        <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Opportunities generated</div>
                      </div>
                      <div>
                        <div style={{ fontSize: 17, fontWeight: 800, color: C.text }}>{fmtMoneyShort(totalValue)}</div>
                        <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Pipeline value</div>
                      </div>
                      <div>
                        <div style={{ fontSize: 17, fontWeight: 800, color: C.green }}>{won.length}</div>
                        <div style={{ fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.3, fontWeight: 700 }}>Closed won</div>
                      </div>
                    </div>
                  </Card>
                );
              })}

              <div style={{ marginTop: 32 }}>
                <h4 style={{ margin: "0 0 4px", fontSize: 14, fontWeight: 700, color: C.text }}>Sourcing</h4>
                <p style={{ margin: "0 0 14px", fontSize: 12, color: C.textMute }}>How every opportunity in your pipeline actually got here.</p>

                {(() => {
                  const scoped = filteredOpps.filter(o => o.sfStatus !== "removed");
                  const buckets = [...SOURCE_OPTIONS, "Unspecified"];
                  const bySource = buckets.map(s => {
                    const opps = scoped.filter(o => (o.source || "Unspecified") === s);
                    const open = opps.filter(o => OPEN_STAGES.includes(o.stage));
                    const won = opps.filter(o => o.stage === "Closed Won");
                    return {
                      source: s,
                      opps: opps.sort((a, b) => (a.closeDate || "9999").localeCompare(b.closeDate || "9999")),
                      openValue: open.reduce((sum, o) => sum + oppTotal(o), 0),
                      openCount: open.length,
                      wonValue: won.reduce((sum, o) => sum + oppTotal(o), 0),
                      wonCount: won.length,
                    };
                  });
                  const totalOpenValue = Math.max(1, bySource.reduce((s, b) => s + b.openValue, 0));

                  return (
                    <>
                      <div style={{ display: "grid", gridTemplateColumns: "repeat(4, 1fr)", gap: 12, marginBottom: 16 }}>
                        {bySource.map(b => (
                          <Card key={b.source} style={{ padding: "14px 16px" }}>
                            <div style={{ display: "flex", alignItems: "center", gap: 7, marginBottom: 6 }}>
                              <span style={{ width: 8, height: 8, borderRadius: 2, background: SOURCE_COLOR[b.source], display: "inline-block" }} />
                              <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>{b.source}</span>
                            </div>
                            <div style={{ fontSize: 20, fontWeight: 800, color: C.text }}>{fmtMoneyShort(b.openValue)}</div>
                            <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 2 }}>{b.openCount} open · {b.wonCount} won ({fmtMoneyShort(b.wonValue)})</div>
                          </Card>
                        ))}
                      </div>

                      <Card style={{ marginBottom: 20 }}>
                        <h5 style={{ margin: "0 0 12px", fontSize: 12.5, fontWeight: 700, color: C.textSoft, textTransform: "uppercase", letterSpacing: 0.3 }}>Open pipeline mix</h5>
                        <div style={{ display: "flex", height: 36, borderRadius: 8, overflow: "hidden", gap: 2, marginBottom: 10 }}>
                          {bySource.filter(b => b.openValue > 0).map(b => (
                            <div key={b.source} style={{
                              width: `${(b.openValue / totalOpenValue) * 100}%`, background: SOURCE_COLOR[b.source],
                              display: "flex", alignItems: "center", justifyContent: "center", minWidth: 40
                            }}>
                              <span style={{ fontSize: 10.5, fontWeight: 700, color: "#fff" }}>{fmtMoneyShort(b.openValue)}</span>
                            </div>
                          ))}
                        </div>
                        <div style={{ display: "flex", gap: 16, flexWrap: "wrap" }}>
                          {bySource.map(b => (
                            <span key={b.source} style={{ display: "flex", alignItems: "center", gap: 5, fontSize: 12, color: C.textSoft }}>
                              <span style={{ width: 8, height: 8, borderRadius: 2, background: SOURCE_COLOR[b.source], display: "inline-block" }} />
                              {b.source} {fmtMoneyShort(b.openValue)} ({totalOpenValue > 1 ? ((b.openValue / totalOpenValue) * 100).toFixed(0) : 0}%)
                            </span>
                          ))}
                        </div>
                      </Card>

                      {bySource.filter(b => b.opps.length > 0).map(b => (
                        <div key={b.source} style={{ marginBottom: 20 }}>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 8 }}>
                            <span style={{ width: 8, height: 8, borderRadius: 2, background: SOURCE_COLOR[b.source], display: "inline-block" }} />
                            <h5 style={{ margin: 0, fontSize: 13, fontWeight: 700, color: C.text }}>{b.source}</h5>
                            <span style={{ fontSize: 11.5, color: C.textMute }}>{b.opps.length} {b.opps.length === 1 ? "opportunity" : "opportunities"}</span>
                          </div>
                          <Card style={{ padding: 0, overflow: "hidden" }}>
                            <div style={{ display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.9fr 0.9fr 0.9fr", padding: "9px 16px", fontSize: 10.5, color: C.textMute, textTransform: "uppercase", letterSpacing: 0.4, borderBottom: `1px solid ${C.border}`, fontWeight: 700, background: "#F8FAFC" }}>
                              <span>Opportunity</span><span>Account</span><span>Stage</span><span>Value</span><span>Close date</span>
                            </div>
                            {b.opps.map(o => (
                              <div key={o.id} onClick={() => setOppModal(o)} style={{
                                display: "grid", gridTemplateColumns: "1.8fr 1.2fr 0.9fr 0.9fr 0.9fr", padding: "11px 16px",
                                fontSize: 12.5, borderBottom: `1px solid ${C.border}`, cursor: "pointer", alignItems: "center"
                              }}>
                                <span style={{ color: C.text, fontWeight: 600 }}>{o.name}</span>
                                <span style={{ color: C.textSoft }}>{o.account}</span>
                                <span><Pill color={STAGE_COLOR[o.stage]} bg={STAGE_COLOR[o.stage] + "1A"}>{o.stage}</Pill></span>
                                <span style={{ color: C.textSoft, fontVariantNumeric: "tabular-nums" }}>{fmtMoneyShort(oppTotal(o))}</span>
                                <span style={{ color: C.textSoft }}>{fmtDate(o.closeDate)}</span>
                              </div>
                            ))}
                          </Card>
                        </div>
                      ))}
                    </>
                  );
                })()}
              </div>
            </div>
          )}

          {tab === "tasks" && (
            <div>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <span style={{ fontSize: 11, fontWeight: 700, letterSpacing: 0.3, textTransform: "uppercase", color: C.textMute }}>Group by</span>
                  <div style={{ display: "flex", gap: 6 }}>
                    {[["none", "None"], ["account", "Account"], ["priority", "Priority"], ["cd", "CD"]].map(([val, label]) => (
                      <div key={val} onClick={() => setTaskGroupBy(val)} style={{
                        fontSize: 12, padding: "6px 12px", borderRadius: 999, cursor: "pointer", fontWeight: 600,
                        background: taskGroupBy === val ? C.sidebarActiveBg : C.card,
                        color: taskGroupBy === val ? C.sidebarActiveText : C.textSoft, border: `1px solid ${taskGroupBy === val ? "#C7D2FE" : C.border}`
                      }}>{label}</div>
                    ))}
                  </div>
                </div>
                <div style={{ display: "flex", gap: 8 }}>
                  <button style={ghostBtn} onClick={() => setTaskImportOpen(true)}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Import from Excel</button>
                  <button style={primaryBtn} onClick={() => setTaskModal({})}><Plus size={14} style={{ verticalAlign: -2, marginRight: 4 }} />Add task</button>
                </div>
              </div>

              <div style={{ position: "relative", marginBottom: 16, maxWidth: 340 }}>
                <Search size={14} color={C.textMute} style={{ position: "absolute", left: 11, top: "50%", transform: "translateY(-50%)" }} />
                <input
                  value={taskSearch}
                  onChange={e => setTaskSearch(e.target.value)}
                  placeholder="Search by task name or account…"
                  style={{ ...inputStyle, marginBottom: 0, paddingLeft: 32 }}
                />
                {taskSearch && (
                  <button onClick={() => setTaskSearch("")} style={{
                    position: "absolute", right: 8, top: "50%", transform: "translateY(-50%)",
                    background: "none", border: "none", cursor: "pointer", color: C.textMute, padding: 2
                  }}>
                    <X size={14} />
                  </button>
                )}
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
                    { key: "partnerContact", label: "Partner Contact", options: ["All", ...taskPartnerOpts] },
                    { key: "partnerEngaged", label: "Partner Engaged", options: ["All", ...PARTNER_ENGAGED_OPTIONS] },
                    { key: "isOpp", label: "Opp", options: ["All", "Yes", "No"] },
                    { key: "dealReg", label: "Deal Reg?", options: ["All", ...taskDealRegOpts] },
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
                    {taskGroupBy !== "none" && (
                      <div style={{
                        display: "flex", alignItems: "center", gap: 10, padding: "13px 16px",
                        borderLeft: `3px solid ${topPriorityColor}`
                      }}>
                        <div style={{ flex: 1, minWidth: 0 }}>
                          <div style={{ fontSize: 13.5, fontWeight: 700, color: C.text, display: "flex", alignItems: "center", gap: 6 }}>
                            {taskGroupBy === "account"
                              ? (g.key === "unassigned" ? <AlertCircle size={13} color={C.textMute} /> : <Building2 size={13} color={C.textMute} />)
                              : taskGroupBy === "cd"
                              ? (g.key === "no-cd" ? <AlertCircle size={13} color={C.textMute} /> : <span style={{ width: 9, height: 9, borderRadius: 2, background: topPriorityColor, display: "inline-block" }} />)
                              : <span style={{ width: 9, height: 9, borderRadius: 2, background: g.color || topPriorityColor, display: "inline-block" }} />}
                            {g.title}
                          </div>
                          <div style={{ fontSize: 11.5, color: C.textMute, marginTop: 1 }}>{g.subtitle}</div>
                        </div>
                        <span style={{ fontSize: 11.5, color: C.textMute, fontWeight: 600 }}>{g.tasks.length} {g.tasks.length === 1 ? "task" : "tasks"}</span>
                      </div>
                    )}
                    <div style={{ borderTop: taskGroupBy !== "none" ? `1px solid ${C.border}` : "none", overflowX: "auto" }}>
                      <div style={{ minWidth: 980 }}>
                        {taskTableHeader}
                        {g.tasks.map(t => renderTaskRow(t))}
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
                Both JSON and Excel backups can be restored; restoring replaces your accounts, opportunities, tasks, campaigns, and goals with what's in the file (org structure and settings are left as-is).
              </p>
              <div style={{ display: "flex", gap: 8, marginBottom: 12, flexWrap: "wrap" }}>
                <button style={ghostBtn} onClick={exportBackup}><UploadCloud size={13} style={{ verticalAlign: -2, marginRight: 6, transform: "rotate(180deg)" }} />Download backup (JSON)</button>
                <button style={ghostBtn} onClick={exportExcelBackup}><FileSpreadsheet size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Download as Excel</button>
                <button style={ghostBtn} onClick={() => backupFileRef.current?.click()}><RefreshCw size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Restore from JSON</button>
                <button style={ghostBtn} onClick={() => restoreExcelFileRef.current?.click()}><RefreshCw size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Restore from Excel</button>
                <input ref={backupFileRef} type="file" accept=".json" onChange={handleBackupFile} style={{ display: "none" }} />
                <input ref={restoreExcelFileRef} type="file" accept=".xlsx,.xls" onChange={handleExcelBackupFile} style={{ display: "none" }} />
              </div>
              {restoreError && <div style={{ fontSize: 12.5, color: C.red, marginBottom: 12 }}>{restoreError}</div>}
              {pendingRestore && (
                <div style={{ background: "#FFFBEB", border: `1px solid #FDE68A`, borderRadius: 10, padding: 14, marginBottom: 16 }}>
                  <div style={{ fontSize: 13, fontWeight: 700, color: "#92400E", marginBottom: 6 }}>Confirm restore</div>
                  <div style={{ fontSize: 12.5, color: "#78350F", marginBottom: 10 }}>
                    This file contains {(pendingRestore.opportunities || []).length} opportunities, {(pendingRestore.tasks || []).length} tasks,
                    {" "}{(pendingRestore.accounts || []).length} accounts, {(pendingRestore.campaigns || []).length} campaigns, and {(pendingRestore.goals || []).length} goals.
                    Restoring will <strong>replace everything currently in this command center</strong> — this can't be undone.
                  </div>
                  <div style={{ display: "flex", gap: 8 }}>
                    <button style={ghostBtn} onClick={() => setPendingRestore(null)}>Cancel</button>
                    <button style={{ ...primaryBtn, background: C.amber }} onClick={confirmRestore}>Replace current data</button>
                  </div>
                </div>
              )}

              <h4 style={{ fontSize: 14, fontWeight: 700, color: C.red, marginBottom: 4, marginTop: 22 }}>Danger zone</h4>
              <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 12 }}>
                Permanently erase every opportunity, task, account, goal, region, and team, and return this command center to a blank slate. This affects everyone who uses this shared link.
              </p>
              <button style={{ ...ghostBtn, ...dangerText, borderColor: "#FECACA" }} onClick={() => setResetConfirmOpen(true)}>
                <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Reset data
              </button>

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
            initial={oppModal}
            reps={reps}
            accounts={data.accounts}
            campaigns={data.campaigns}
            tasks={oppModal.id ? data.tasks.filter(t => t.oppId === oppModal.id) : []}
            onCancel={() => setOppModal(null)}
            onSave={(item) => { saveOpportunity(item); setOppModal(null); }}
            onDelete={(id) => { remove("opportunities", id); setOppModal(null); }}
            onOpenTask={(t) => { setOppModal(null); setTaskModal(t); }}
            onAddTask={() => { const { id: oppId, accountId, account } = oppModal; setOppModal(null); setTaskModal({ oppId, accountId, account }); }}
          />
        </Modal>
      )}
      {taskModal && (
        <Modal title={taskModal.id ? "Edit task" : "New task"} onClose={() => setTaskModal(null)}>
          <TaskForm
            initial={taskModal}
            reps={reps}
            accounts={data.accounts}
            opportunities={data.opportunities}
            onCancel={() => setTaskModal(null)}
            onSave={(item) => { saveTask(item); setTaskModal(null); }}
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
      {campModal && (
        <Modal title={campModal.id ? "Edit campaign" : "New campaign"} wide onClose={() => setCampModal(null)}>
          <CampaignForm
            initial={campModal.id ? campModal : null}
            accounts={data.accounts}
            opportunities={data.opportunities}
            onCancel={() => setCampModal(null)}
            onSave={(item) => { upsert("campaigns", item); setCampModal(null); }}
            onDelete={(id) => { remove("campaigns", id); setCampModal(null); }}
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
        <Modal title={accModal.id ? "Edit account" : "New account"} wide onClose={() => setAccModal(null)}>
          <AccountForm
            initial={accModal.id ? accModal : null}
            reps={reps}
            computed={accModal.id ? accountComputed(accModal) : null}
            tasks={accModal.id ? data.tasks.filter(t => t.accountId ? t.accountId === accModal.id : normName(t.account) === normName(accModal.name)) : []}
            onCancel={() => setAccModal(null)}
            onSave={(item) => { upsert("accounts", item); setAccModal(null); }}
            onDelete={(id) => { remove("accounts", id); setAccModal(null); }}
            onOpenTask={(t) => { setAccModal(null); setTaskModal(t); }}
            onAddTask={() => { const { id: accountId, name: account } = accModal; setAccModal(null); setTaskModal({ accountId, account }); }}
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
      {oppImportOpen && (
        <OpportunityImportWizard
          data={data}
          reps={reps}
          onClose={() => setOppImportOpen(false)}
          onApply={(next) => { setData(next); setOppImportOpen(false); }}
        />
      )}
      {resetConfirmOpen && (
        <Modal title="Reset all data?" onClose={() => setResetConfirmOpen(false)}>
          <div style={{ display: "flex", gap: 10, marginBottom: 16, background: "#FEF2F2", border: "1px solid #FECACA", borderRadius: 10, padding: 14 }}>
            <AlertCircle size={18} color={C.red} style={{ flexShrink: 0, marginTop: 1 }} />
            <div style={{ fontSize: 13, color: "#7F1D1D" }}>
              This permanently deletes every opportunity, task, account, goal, region, and team in this command center and returns it to a blank slate.
              This cannot be undone, and it affects everyone who uses this shared link — not just you.
            </div>
          </div>
          <p style={{ fontSize: 12.5, color: C.textSoft, marginBottom: 16 }}>
            If you haven't already, consider downloading a backup first (Settings → Backup & restore) so you have something to restore from if this was a mistake.
          </p>
          <div style={{ display: "flex", justifyContent: "flex-end", gap: 8 }}>
            <button style={ghostBtn} onClick={() => setResetConfirmOpen(false)}>Cancel</button>
            <button style={{ ...primaryBtn, background: C.red }} onClick={performReset}>
              <Trash2 size={13} style={{ verticalAlign: -2, marginRight: 6 }} />Yes, reset everything
            </button>
          </div>
        </Modal>
      )}
    </div>
  );
}
