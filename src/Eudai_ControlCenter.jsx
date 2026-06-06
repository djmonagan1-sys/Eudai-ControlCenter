import React, { useState, useMemo, useCallback, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  LayoutDashboard, FlaskConical, Factory, TrendingUp, Users, ClipboardList,
  Shield, Sparkles, Plus, ArrowRight, AlertCircle, Clock, CheckCircle2,
  Circle, DollarSign, Flame, CalendarDays, ChevronRight, X, Send,
  FileText, Search, MoreHorizontal, GitBranch, Cloud, CloudOff,
  Lightbulb, Palette, Megaphone, Trash2, MessagesSquare, RefreshCw, Target,
  Settings, UserRound, Download
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  ControlCenter — the operating system for Eudai.                  */
/*  Light editorial theme drawn from the can: cream, charcoal,         */
/*  burnt-orange. Persistent shared storage across founders.           */
/* ------------------------------------------------------------------ */

const C = {
  bg: "#FAF7F2",          // warm cream
  panel: "#FFFFFF",
  panel2: "#F4EFE7",      // deeper cream
  border: "#E7E0D4",
  borderSoft: "#EFE9DE",
  text: "#1F1B16",        // warm charcoal
  dim: "#6E655A",
  faint: "#A39A8C",
  accent: "#C8732E",      // burnt orange from the can's crown
  accentSoft: "#F3E2D2",
  ink: "#2A251F",
};

const STAGE_COLORS = {
  "Idea": "#A39A8C", "Research": "#5B7A9D", "Validation": "#8A6FA8",
  "In Progress": "#C8932E", "Complete": "#5E7D54",
  "Contacted": "#5B7A9D", "Sample Requested": "#8A6FA8", "Quote Received": "#C8932E",
  "Negotiating": "#C8732E", "Approved": "#5E7D54", "Active Vendor": "#4E7046",
  "Identified": "#A39A8C", "Outreach": "#5B7A9D", "Meeting Scheduled": "#8A6FA8",
  "Follow-Up": "#C8932E", "Due Diligence": "#C8732E", "Passed": "#B4574B", "Invested": "#4E7046",
  "Not Started": "#A39A8C", "Submitted": "#C8932E", "Passed Testing": "#5E7D54",
  "Todo": "#A39A8C", "Doing": "#C8932E", "Done": "#5E7D54",
};

const PRIORITY = { Critical: "#B4574B", High: "#C8732E", Medium: "#C8932E", Low: "#A39A8C" };

/* ----------------------------- seed data ----------------------------- */
/* Sourced ONLY from: Eudai_Todo_By_6_12_26.docx, Eudai_Meeting_.docx,   */
/* Getting_on_the_same_page-Eudai_.docx. No invented records.            */
const seed = {
  company: { cash: 0, burn: 0, runwayMonths: 6, name: "Eudai" },
  tasks: [
    /* ---- From "Eudai Todo By 6_12_26" — due 6/12/26 ---- */
    { id: 1, title: "1 Mock Pitch Deck", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Fundraising" },
    { id: 2, title: "1 Mock Website", owner: "GP", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Marketing" },
    { id: 3, title: "Increase IG followers by 200 (flagged — maybe scratch)", owner: "—", due: "2026-06-12", priority: "Low", status: "Todo", link: "Marketing" },
    { id: 4, title: "Cost projections with dates cash is needed by", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "General" },
    { id: 5, title: "Trademark", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 6, title: "LLC", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 7, title: "Barcode", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 8, title: "Product liability insurance", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 9, title: "Obtain seller's permit", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 10, title: "Workflow system", owner: "—", due: "2026-06-12", priority: "Medium", status: "Doing", link: "General" },
    { id: 11, title: "Create mood board & branding vision for initial mockup", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Marketing" },
    /* ---- From "Eudai Meeting" action items ---- */
    { id: 12, title: "Finish Beverage Feedback Form", owner: "—", due: "2026-06-12", priority: "Medium", status: "Doing", link: "Product" },
    { id: 13, title: "Figure out how we are going to push forward on branding", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Marketing" },
    { id: 14, title: "Bullet points of what can get done while waiting on FDA", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "General" },
    { id: 15, title: "Operating Agreement", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Regulatory" },
    { id: 16, title: "Website URL", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Marketing" },
    { id: 17, title: "Social Media", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "Marketing" },
    { id: 18, title: "Plan non-negotiables for meetings going forward", owner: "—", due: "2026-06-12", priority: "Medium", status: "Todo", link: "General" },
  ],
  product: {
    name: "Eudai",
    flavor: "Orange-Peach Cream",
    size: "12 fl oz (355 ml)",
    caffeine: 120,
    /* Ingredient list from the meeting doc — theobromine removed, lemon balm at the decided 100mg.
       Per-can cost starts at 0; click to fill in real costs as quotes come in. */
    ingredients: [
      { id: 1, name: "Filtered Water", amount: "—", cost: 0 },
      { id: 2, name: "Orange Juice Concentrate", amount: "9g sugar total", cost: 0 },
      { id: 3, name: "Peach Juice Concentrate", amount: "9g sugar total", cost: 0 },
      { id: 4, name: "Malic Acid", amount: "—", cost: 0 },
      { id: 5, name: "Orange Cream Flavor", amount: "—", cost: 0 },
      { id: 6, name: "Peach Flavor", amount: "—", cost: 0 },
      { id: 7, name: "Sodium Citrate", amount: "200mg Na total", cost: 0 },
      { id: 8, name: "L-Tyrosine", amount: "500mg", cost: 0 },
      { id: 9, name: "L-Theanine", amount: "250mg", cost: 0 },
      { id: 10, name: "Lemon Balm", amount: "100mg", cost: 0 },
      { id: 11, name: "Panax Ginseng", amount: "200mg", cost: 0 },
      { id: 12, name: "Rhodiola Rosea", amount: "200mg", cost: 0 },
      { id: 13, name: "Cognizin", amount: "500mg", cost: 0 },
      { id: 14, name: "Sea Salt", amount: "200mg Na total", cost: 0 },
      { id: 15, name: "Monkfruit MV50", amount: "—", cost: 0 },
      { id: 16, name: "Caffeine from Green Tea 98%", amount: "120mg", cost: 0 },
      { id: 17, name: "Magnesium Glycinate", amount: "12mg", cost: 0 },
      { id: 18, name: "Reb M 95%", amount: "—", cost: 0 },
      { id: 19, name: "Vitamin B2 Riboflavin", amount: "100% DV", cost: 0 },
      { id: 20, name: "Vitamin B3 Flush Free Niacin", amount: "100% DV", cost: 0 },
      { id: 21, name: "Vitamin B5 Calcium Pantothenate", amount: "100% DV", cost: 0 },
      { id: 22, name: "Vitamin B6 Pyridoxine HCL", amount: "100% DV", cost: 0 },
      { id: 23, name: "Vitamin B12 Methylcobalamin", amount: "100% DV", cost: 0 },
    ],
    stack: [
      { layer: "Energy", items: "120mg green tea caffeine" },
      { layer: "Focus", items: "250mg L-theanine · Cognizin · 500mg tyrosine" },
      { layer: "Performance Adaptation", items: "Rhodiola · Panax ginseng · electrolytes" },
      { layer: "Calm", items: "Lemon balm (100mg) · magnesium glycinate" },
    ],
  },
  vendors: [],
  investors: [],
  financials: {
    /* From "Getting on the same page" doc */
    pricing: { can: 3.50, twelvePack: 30 },
    /* Market sizing — public ballpark estimates, every number editable. Verify before pitching. */
    market: {
      tamB: 90,
      tamNote: "Global energy drink market. Public estimates cluster around $90B mid-decade (Statista / Grand View ranges roughly $86–100B). This is the ceiling if Eudai competed in every geography and channel — useful for narrative scale, not for planning.",
      samB: 50,
      samNote: "US functional & better-for-you beverage space — healthy energy, nootropic and wellness drinks — commonly estimated near $50B. This is the slice reachable with a US, health-positioned canned drink through our channels: gyms, campus, grocery, convenience, supplement stores.",
      somStudents: 47000, somPct: 40, somCansWeek: 2,
      somNote: "Bottom-up from our beachhead. USC enrollment (~47,000) × share of college students who drink energy drinks (surveys commonly find 30–50%; we assume 40%) × a moderate 2 cans/week × 52 weeks × our $3.50 can price. Deliberately conservative: excludes faculty/staff, nearby gyms, and online sales.",
    },
    /* The financial plan: % of SOM captured per year. */
    projections: [
      { id: 1, label: "Year 1", capture: 3 },
      { id: 2, label: "Year 2", capture: 8 },
      { id: 3, label: "Year 3", capture: 15 },
    ],
    statements: { cans: 10000 },
    revenueToDate: 0,
    /* From "Drink Business — Cost Sheet" */
    setupCosts: [
      { item: "LLC formation", cost: 70, time: "" },
      { item: "Statement of Information (SOI)", cost: 25, time: "" },
      { item: "Certificate of Status", cost: 5, time: "" },
      { item: "End of Year Tax", cost: 800, time: "" },
      { item: "Trademark", cost: 350, time: "" },
      { item: "Manufacturing overhead (outsourced)", cost: 7500, time: "" },
      { item: "R&D", cost: 10000, time: "", spent: true },
      { item: "Permit to sell", cost: 0, time: "Same Day" },
      { item: "Business License", cost: 100, time: "1-7 Days" },
      { item: "Barcode", cost: 30, time: "" },
      { item: "Insurance", cost: 600, time: "" },
      { item: "Website URL", cost: 150, time: "" },
    ],
    productionRun: {
      cans: 10000,
      steps: [
        { step: "Empty aluminum cans", low: 1700, likely: 1950, high: 2200 },
        { step: "Shrink sleeves (printed)", low: 1200, likely: 1500, high: 1800 },
        { step: "Sleeve application", low: 600, likely: 900, high: 1200 },
        { step: "Manufacturing / filling", low: 3200, likely: 4350, high: 5500 },
        { step: "Ingredients", low: 1200, likely: 1800, high: 2400 },
        { step: "Packaging (trays, cartons, wrap)", low: 500, likely: 750, high: 1000 },
      ],
    },
  },
  creative: {
    ideas: [],
    /* From "Eudai Todo By 6_12_26" — the mock deliverables in flight */
    mockups: [
      { id: 1, name: "Mock Pitch Deck", owner: "—", status: "In Progress", notes: "" },
      { id: 2, name: "Mock Website", owner: "GP", status: "In Progress", notes: "" },
      { id: 3, name: "Mood board & branding vision", owner: "—", status: "Not Started", notes: "To send for initial mockup" },
    ],
    /* From "Getting on the same page" — marketing direction already decided */
    marketing: [
      { id: 1, text: "Hero ingredient on marketing materials: Cognizin" },
      { id: 2, text: "Pairing marketing: 120mg caffeine + no sucralose" },
      { id: 3, text: "Collaborations: fitness brands, productivity creators, athletes, universities, tech/startup culture" },
      { id: 4, text: "Transparent dosing of every ingredient as a differentiator" },
    ],
  },
  settings: {
    companyName: "Eudai",
    accent: "#C26A35",
    background: "cream",
    defaultPage: "dashboard",
    founders: ["Grayson", "Phoenix", "Dylan", "Marcellus"],
  },
  gtm: {
    /* All from "Getting on the same page — Eudai" */
    valueProp: "Provides smooth, calmer energy without the crash and overstimulation through a clinically backed nootropic blend including Cognizin, lemon balm, a moderate 120mg of caffeine and more. Sweetened through a unique blend of monk fruit, stevia, and natural sugar from fruit juice with no artificial sweeteners or preservatives.",
    category: "Nootropic — health & wellness energy supplement",
    mainBenefit: "Calm focus — focus without feeling guilty",
    brandWords: ["Presence", "Focus", "Clarity", "Guilt-free"],
    targetCustomer: "An ambitious college student who is active, studious, and mindful of what they consume — people living an intentional lifestyle who want to perform at a high level and care about their health.",
    ageRange: "18–32",
    occasions: ["Gym", "Studying", "Work", "Gaming", "Travel", "Content creation", "Late nights"],
    channels: ["Gyms", "Erewhon / Whole Foods", "College campuses", "Convenience stores", "Coffee shops", "Supplement stores"],
    competitors: ["Neutonic", "Celsius", "Gorilla Mind", "Jocko Fuel", "C4", "Yerba Madre", "Bloom", "Uptime", "Zoa", "Zevia", "FocusAid"],
    competitorGaps: ["Sweetener blend", "Creating a brand identity", "Nootropic focus", "Displaying differentiators"],
    switchingFrom: ["Coffee", "Celsius", "Preworkout", "Red Bull", "Soda", "Nootropics"],
    collaborations: ["Fitness brands", "Productivity creators", "Athletes", "Universities", "Tech/startup culture"],
    threeYear: "$1M in revenue, retail across the country, not necessarily profitable, loyal community",
    approach: "Mass-market with broad appeal",
    brand: {
      coreIdentity: "Being present by living with full effort, commitment, and passion in everything that you do.",
      nameMeaning: "Eudai, short for eudaimonia, coined by Aristotle: the lifelong pursuit of living a meaningful life, filled with purpose, rational decision making, and strong moral values. 'Eu' means good — daimon: spirit, a good-natured inner force.",
      colorsUse: ["White", "Black", "Silver", "Calm / muted tones"],
      colorsNever: ["Pink", "Neon", "Loud colors", "Careful with beige (coffee)"],
      feel: ["Minimal", "Premium", "Calm", "Clean"],
    },
  },
  scheduled: [
    {
      id: 1, title: "Next founders meeting", date: "", time: "", parts: "All founders",
      /* Agenda from the meeting doc's "Next Meeting" list */
      agenda: [],
    },
  ],
  meetings: [
    {
      id: 1, date: "—", parts: "Founders",
      notes: "Ready to sign and understand we are out of samples — want a few small changes before documents are prepared. Stack by function: Energy (120mg green tea caffeine) · Focus (250mg L-theanine, 250mg Cognizin, 500–750mg tyrosine) · Performance Adaptation (rhodiola, panax ginseng, electrolytes) · Calm (smaller lemon balm dose, magnesium glycinate).",
      decisions: [
        "Add 250–500mg Cognizin",
        "Decrease lemon balm to 100mg",
        "Confirm L-tyrosine at 500mg",
      ],
      actions: [
        "Ask formulator: will Cognizin change taste? Dosage suggestions?",
        "Ask: where is potassium coming from? Wanted 40mg K (got 200mg on nutrition facts)",
        "Ask: can printed label list B vitamins by name, not 'pantothenic acid'?",
        "Next meeting: beverage feedback form, branding discussion, plan meeting cadence",
      ],
    },
  ],
  regulatory: [
    { id: 1, item: "FDA — waiting on FDA", status: "Submitted", due: "—" },
    { id: 2, item: "Trademark", status: "Not Started", due: "2026-06-12" },
    { id: 3, item: "Establish new LLC", status: "Not Started", due: "2026-06-12" },
    { id: 4, item: "Operating Agreement", status: "Not Started", due: "2026-06-12" },
    { id: 5, item: "Barcode", status: "Not Started", due: "2026-06-12" },
    { id: 6, item: "Product liability insurance", status: "Not Started", due: "2026-06-12" },
    { id: 7, item: "Seller's permit", status: "Not Started", due: "2026-06-12" },
  ],
  docs: [
    { id: 1, name: "Eudai Meeting .docx", tag: "Product", date: "2026-06-05" },
    { id: 2, name: "Eudai Todo By 6_12_26.docx", tag: "Operations", date: "2026-06-05" },
    { id: 3, name: "Getting on the same page — Eudai.docx", tag: "Brand", date: "2026-06-05" },
  ],
};

const TODAY = new Date("2026-06-05");
const fmt$ = (n) => "$" + n.toLocaleString();
const fmtBig = (n) => n >= 1e9 ? `$${(n / 1e9).toFixed(n % 1e9 === 0 ? 0 : 1)}B` : n >= 1e6 ? `$${(n / 1e6).toFixed(1)}M` : fmt$(Math.round(n));
const daysUntil = (d) => Math.ceil((new Date(d) - TODAY) / 86400000);
const STORE_KEY = "ccos:eudai:data:v13";

/* Shared financial math — single source of truth for totals.
   If the product's ingredients carry per-can costs, the production run's
   "Ingredients" line is derived from them (cost × cans) instead of the sheet. */
const ingPerCan = (product) => product ? product.ingredients.reduce((a, i) => a + (i.cost || 0), 0) : 0;

const finTotals = (fin, product) => {
  if (!fin) return null;
  const setup = fin.setupCosts.reduce((a, r) => a + r.cost, 0);
  const spent = fin.setupCosts.filter((r) => r.spent).reduce((a, r) => a + r.cost, 0);
  const ingCan = ingPerCan(product);
  const ingDerived = ingCan > 0 ? Math.round(ingCan * fin.productionRun.cans) : null;
  const prod = { low: 0, likely: 0, high: 0 };
  fin.productionRun.steps.forEach((s) => {
    const isIng = s.step === "Ingredients" && ingDerived != null;
    prod.low += isIng ? ingDerived : s.low;
    prod.likely += isIng ? ingDerived : s.likely;
    prod.high += isIng ? ingDerived : s.high;
  });
  const needed = {
    low: setup + prod.low - spent,
    likely: setup + prod.likely - spent,
    high: setup + prod.high - spent,
  };
  return { setup, spent, prod, needed, ingCan, ingDerived };
};

/* ----------------------------- primitives ----------------------------- */
/* Small-caps micro label, used across Financials & Creative */
const Micro = ({ children, style }) => (
  <span style={{
    fontSize: 10, fontWeight: 600, letterSpacing: 1.1, textTransform: "uppercase",
    color: C.faint, ...style,
  }}>{children}</span>
);

const Badge = ({ label, color }) => (
  <span style={{
    display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 500,
    color: C.ink, whiteSpace: "nowrap",
  }}>
    <span style={{ width: 7, height: 7, borderRadius: 99, background: color, flexShrink: 0 }} /> {label}
  </span>
);

const Card = ({ children, style, onClick, hover }) => {
  const [h, setH] = useState(false);
  return (
    <div onClick={onClick}
      onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        background: h && hover ? "#FFFEFB" : C.panel, border: `1px solid ${h && hover ? "#D8CFBF" : C.borderSoft}`,
        borderRadius: 8, padding: 16, transition: "border-color .12s ease, background .12s ease",
        cursor: onClick ? "pointer" : "default",
        ...style,
      }}>{children}</div>
  );
};

const Btn = ({ children, onClick, primary, small, icon: Icon }) => {
  const [h, setH] = useState(false);
  return (
    <button onClick={onClick} onMouseEnter={() => setH(true)} onMouseLeave={() => setH(false)}
      style={{
        display: "inline-flex", alignItems: "center", gap: 6, fontWeight: 500,
        fontSize: small ? 12 : 13, padding: small ? "5px 10px" : "6px 12px",
        borderRadius: 6, border: `1px solid ${primary ? C.ink : C.border}`,
        background: primary ? (h ? "#3A332A" : C.ink) : (h ? C.panel2 : C.panel),
        color: primary ? "#FAF7F2" : C.ink, cursor: "pointer",
        transition: "all .12s ease", fontFamily: "'Geist', sans-serif",
      }}>
      {Icon && <Icon size={small ? 13 : 14} strokeWidth={2} />} {children}
    </button>
  );
};

/* ----------------------------- mountain motif ----------------------------- */
const Mountains = ({ opacity = 1, height = 54 }) => (
  <svg viewBox="0 0 240 60" style={{ width: "100%", height, opacity }} preserveAspectRatio="none">
    <g fill="none" stroke="#B9AE9C" strokeWidth="0.8">
      <path d="M0 52 L28 24 L40 36 L62 10 L84 40 L100 26 L122 48" />
      <path d="M110 50 L138 18 L154 34 L176 8 L198 38 L214 28 L240 52" />
      <path d="M55 14 L62 10 L70 16" strokeWidth="0.6" />
      <path d="M168 12 L176 8 L184 14" strokeWidth="0.6" />
      <path d="M0 56 C40 50 80 54 120 52 C160 50 200 55 240 53" strokeWidth="0.5" opacity="0.6" />
    </g>
  </svg>
);

/* ----------------------------- modal ----------------------------- */
const Modal = ({ open, onClose, title, children }) => {
  if (!open) return null;
  return (
    <div onClick={onClose} style={{
      position: "fixed", inset: 0, background: "rgba(42,37,31,.35)", backdropFilter: "blur(5px)",
      zIndex: 100, display: "flex", alignItems: "center", justifyContent: "center", padding: 20,
    }}>
      <div onClick={(e) => e.stopPropagation()} style={{
        background: C.panel, border: `1px solid ${C.border}`, borderRadius: 10,
        width: "100%", maxWidth: 460, padding: 20, animation: "pop .15s ease",
        boxShadow: "0 20px 60px rgba(60,45,25,.20)",
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <h3 style={{ fontSize: 15, fontWeight: 600, fontFamily: "'STIX Two Text', serif", margin: 0 }}>{title}</h3>
          <X size={16} color={C.dim} style={{ cursor: "pointer" }} onClick={onClose} />
        </div>
        {children}
      </div>
    </div>
  );
};

const inputStyle = {
  width: "100%", background: "#FFFFFF", border: `1px solid ${C.border}`, borderRadius: 6,
  padding: "6px 10px", color: C.ink, fontSize: 13, fontFamily: "'Geist', sans-serif",
  outline: "none", boxSizing: "border-box",
};

const Field = ({ label, ...p }) => (
  <div style={{ marginBottom: 13 }}>
    <label style={{ fontSize: 11.5, color: C.dim, fontWeight: 700, display: "block", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.6 }}>{label}</label>
    <input {...p} style={inputStyle} />
  </div>
);

const Select = ({ label, options, ...p }) => (
  <div style={{ marginBottom: 13 }}>
    <label style={{ fontSize: 11.5, color: C.dim, fontWeight: 700, display: "block", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.6 }}>{label}</label>
    <select {...p} style={inputStyle}>{options.map((o) => <option key={o} value={o}>{o}</option>)}</select>
  </div>
);

/* =================================================================== */
/*  MAIN APP                                                           */
/* =================================================================== */
export default function Eudai_ControlCenter({ session, onSignOut }) {
  const [view, setView] = useState("dashboard");
  const [data, setData] = useState(seed);
  const [modal, setModal] = useState(null);
  const [form, setForm] = useState({});
  const [copilotOpen, setCopilotOpen] = useState(false);
  const [jump, setJump] = useState("");
  const [me, setMe] = useState(session?.user?.user_metadata?.full_name || session?.user?.email?.split("@")[0] || "");  // from real auth
  const [idOpen, setIdOpen] = useState(false);
  const [syncState, setSyncState] = useState("loading"); // loading | synced | local
  const loaded = useRef(false);

  /* ---------- workspace settings, applied live ---------- */
  const settings = data.settings || { companyName: "Eudai", accent: "#C26A35", background: "cream", defaultPage: "dashboard", founders: [] };
  C.accent = settings.accent;
  C.accentSoft = settings.accent + "1f";
  C.bg = settings.background === "white" ? "#FFFFFF" : "#FAF7F2";
  const updSettings = (patch) => setData((d) => ({ ...d, settings: { ...(d.settings || settings), ...patch } }));
  const addFounder = (name) => {
    const n = name.trim();
    if (!n) return;
    if (!settings.founders.includes(n)) updSettings({ founders: [...settings.founders, n] });
  };

  /* ---------- persistent shared storage: all founders see one dataset ---------- */
  useEffect(() => {
    (async () => {
      try {
        const res = await window.storage.get(STORE_KEY, true);
        if (res?.value) {
          const parsed = JSON.parse(res.value);
          setData(parsed);
          const valid = ["dashboard", "tasks", "product", "financials", "gtm", "settings"];
          if (parsed.settings?.defaultPage && valid.includes(parsed.settings.defaultPage)) setView(parsed.settings.defaultPage);
        }
        setSyncState("synced");
      } catch {
        setSyncState("synced"); // key doesn't exist yet — first run, seed stays
      }
      loaded.current = true;
    })();
  }, []);

  useEffect(() => {
    if (!loaded.current) return;
    const t = setTimeout(async () => {
      try {
        await window.storage.set(STORE_KEY, JSON.stringify(data), true);
        setSyncState("synced");
      } catch { setSyncState("local"); }
    }, 600);
    return () => clearTimeout(t);
  }, [data]);

  const runway = data.company.runwayMonths ?? (data.company.burn > 0 ? (data.company.cash / data.company.burn).toFixed(1) : "—");
  const openTasks = data.tasks.filter((t) => t.status !== "Done");

  const priorities = useMemo(() => {
    return [...data.tasks]
      .filter((t) => t.status !== "Done")
      .map((t) => ({ ...t, d: daysUntil(t.due) }))
      .sort((a, b) => {
        const pr = { Critical: 0, High: 1, Medium: 2, Low: 3 };
        if (a.d < 0 && b.d >= 0) return -1;
        if (b.d < 0 && a.d >= 0) return 1;
        return pr[a.priority] - pr[b.priority] || a.d - b.d;
      });
  }, [data.tasks]);

  const autoTask = (title, link, priority) =>
    setData((d) => ({ ...d, tasks: [{ id: Date.now() + Math.random(), title, owner: "Auto", due: "2026-06-15", priority, status: "Todo", link, auto: true }, ...d.tasks] }));

  const addRecord = useCallback((type) => {
    setData((d) => {
      const nd = { ...d };
      const id = Date.now();
      if (type === "task") nd.tasks = [{ id, status: "Todo", link: form.link || "General", priority: form.priority || "Medium", ...form }, ...d.tasks];
      if (type === "vendor") nd.vendors = [{ id, status: "Research", ...form }, ...d.vendors];
      if (type === "investor") nd.investors = [{ id, status: "Identified", last: "—", ...form }, ...d.investors];
      if (type === "schedule") nd.scheduled = [{
        id, title: form.title || "Founders meeting", date: form.date || "", time: form.time || "",
        parts: form.parts || "All founders",
        agenda: (form.agenda || "").split("\n").map((s) => s.trim()).filter(Boolean),
      }, ...(d.scheduled || [])];
      if (type === "idea") nd.creative = {
        ...(d.creative || { ideas: [], mockups: [], marketing: [] }),
        ideas: [{ id, text: form.text || "", tag: form.tag || "General" }, ...((d.creative && d.creative.ideas) || [])],
      };
      return nd;
    });
    if (type === "vendor") autoTask(`Review details for ${form.name || "new vendor"}`, "Manufacturing", "Medium");
    setModal(null); setForm({});
  }, [form]);

  const advanceStage = (entity, id, stages) =>
    setData((d) => ({
      ...d, [entity]: d[entity].map((r) => {
        if (r.id !== id) return r;
        const i = stages.indexOf(r.status);
        return { ...r, status: stages[Math.min(i + 1, stages.length - 1)] };
      })
    }));

  const toggleTask = (id) => setData((d) => ({
    ...d, tasks: d.tasks.map((t) => t.id === id ? { ...t, status: t.status === "Done" ? "Todo" : "Done" } : t)
  }));

  const cyclePriority = (id) => setData((d) => ({
    ...d, tasks: d.tasks.map((t) => {
      if (t.id !== id) return t;
      const order = ["Low", "Medium", "High", "Critical"];
      return { ...t, priority: order[(order.indexOf(t.priority) + 1) % order.length] };
    })
  }));

  const deleteTask = (id) => setData((d) => ({ ...d, tasks: d.tasks.filter((t) => t.id !== id) }));

  const patchTask = (id, patch) => setData((d) => ({ ...d, tasks: d.tasks.map((t) => t.id === id ? { ...t, ...patch } : t) }));

  const quickAddTask = (title, link) => setData((d) => ({
    ...d, tasks: [...d.tasks, { id: Date.now(), title, owner: me || "—", due: "2026-06-12", priority: "Medium", status: "Todo", link }],
  }));

  const nav = [
    { id: "dashboard", label: "Dashboard", icon: LayoutDashboard },
    { id: "tasks", label: "Tasks", icon: ClipboardList },
    { id: "product", label: "Product", icon: FlaskConical },
    { id: "financials", label: "Financials", icon: TrendingUp },
    { id: "gtm", label: "Strategy", icon: Target },
    { id: "settings", label: "Settings", icon: Settings },
  ];

  return (
    <div style={{
      display: "flex", height: "100vh", background: C.bg, color: C.ink,
      fontFamily: "'Geist', -apple-system, sans-serif", overflow: "hidden",
    }}>
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=STIX+Two+Text:ital,wght@0,400..700;1,400..700&family=Geist:wght@400;500;600;700&display=swap');
        * { box-sizing: border-box; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: #DDD4C5; border-radius: 8px; }
        @keyframes pop { from { transform: scale(.96); opacity: 0 } to { transform: scale(1); opacity: 1 } }
        @keyframes slidein { from { transform: translateX(20px); opacity: 0 } to { transform: translateX(0); opacity: 1 } }
        @keyframes fade { from { opacity: 0; transform: translateY(8px) } to { opacity: 1; transform: translateY(0) } }
        .fade { animation: fade .3s ease both; }
        select option { background: #fff; }
      `}</style>

      {/* ---------------- Sidebar ---------------- */}
      <aside style={{
        width: 212, background: C.panel, borderRight: `1px solid ${C.borderSoft}`,
        display: "flex", flexDirection: "column", padding: "14px 10px 0", flexShrink: 0,
      }}>
        <div style={{ padding: "2px 8px 12px", display: "flex", alignItems: "center", gap: 10 }}>
          <div style={{
            width: 30, height: 30, borderRadius: "50%", background: "#FFFFFF",
            border: `1px solid ${C.border}`,
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
          }}>
            <span style={{
              fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontWeight: 500,
              fontSize: 9.5, color: C.ink,
            }}>Eu.</span>
          </div>
          <div>
            <div style={{
              fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontWeight: 500,
              fontSize: 16, letterSpacing: -0.2, color: C.ink, lineHeight: 1,
            }}>Eudai.</div>
            <div style={{
              fontSize: 8.5, color: C.faint, fontWeight: 600, letterSpacing: 1.6,
              textTransform: "uppercase", marginTop: 3,
            }}>ControlCenter</div>
          </div>
        </div>

        <div style={{ height: 1, background: C.borderSoft, margin: "0 8px 10px" }} />

        {nav.map((n) => {
          const active = view === n.id;
          return (
            <button key={n.id} onClick={() => setView(n.id)} style={{
              display: "flex", alignItems: "center", gap: 9, padding: "5px 9px", marginBottom: 1,
              borderRadius: 6, border: "none", cursor: "pointer", width: "100%", textAlign: "left",
              background: active ? "#F1EBE0" : "transparent",
              color: active ? C.ink : C.dim, fontFamily: "'Geist', sans-serif",
              fontSize: 13, fontWeight: 500, transition: "background .1s",
            }}
              onMouseEnter={(e) => { if (!active) e.currentTarget.style.background = "#F6F1E8"; }}
              onMouseLeave={(e) => { if (!active) e.currentTarget.style.background = "transparent"; }}>
              <n.icon size={15} strokeWidth={1.8} color={active ? C.accent : C.faint} />
              {n.label}
            </button>
          );
        })}

        <div style={{ marginTop: "auto", paddingBottom: 10 }}>
          <button onClick={() => setCopilotOpen(true)} style={{
            display: "flex", alignItems: "center", gap: 9, padding: "5px 9px", width: "100%",
            borderRadius: 6, border: "none", cursor: "pointer",
            background: "transparent", color: C.dim, fontFamily: "'Geist', sans-serif",
            fontSize: 13, fontWeight: 500,
          }}
            onMouseEnter={(e) => e.currentTarget.style.background = "#F6F1E8"}
            onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
            <Sparkles size={15} strokeWidth={1.8} color={C.accent} /> AI Copilot
          </button>
          <div style={{ fontSize: 10, color: C.faint, padding: "10px 9px 2px", fontStyle: "italic", fontFamily: "'STIX Two Text', serif" }}>
            A life well lived.
          </div>
        </div>
      </aside>

      {/* ---------------- Right column: topbar + main ---------------- */}
      <div style={{ flex: 1, display: "flex", flexDirection: "column", minWidth: 0 }}>
        <header style={{
          height: 44, flexShrink: 0, borderBottom: `1px solid ${C.borderSoft}`, background: C.panel,
          display: "flex", alignItems: "center", gap: 14, padding: "0 18px",
        }}>
          <div style={{ fontSize: 12.5, color: C.dim, display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{ fontWeight: 600, color: C.ink }}>{settings.companyName}</span>
            <ChevronRight size={12} color={C.faint} />
            <span style={{ fontWeight: 500 }}>{(nav.find((n) => n.id === view) || {}).label || ""}</span>
          </div>
          <div style={{ flex: 1 }} />
          <div style={{ position: "relative" }}>
            <Search size={13} color={C.faint} style={{ position: "absolute", left: 9, top: 7.5 }} />
            <input value={jump} placeholder="Go to…" onChange={(e) => setJump(e.target.value)}
              onKeyDown={(e) => {
                if (e.key === "Enter") {
                  const hit = nav.find((n) => n.label.toLowerCase().includes(jump.toLowerCase()));
                  if (hit) { setView(hit.id); setJump(""); }
                }
                if (e.key === "Escape") setJump("");
              }}
              style={{
                width: 190, background: C.bg, border: `1px solid ${C.borderSoft}`, borderRadius: 6,
                padding: "5px 10px 5px 28px", fontSize: 12.5, fontFamily: "inherit", color: C.ink, outline: "none",
              }} />
            {jump && (
              <div style={{
                position: "absolute", top: 32, right: 0, width: 190, background: C.panel,
                border: `1px solid ${C.border}`, borderRadius: 8, zIndex: 50, overflow: "hidden",
                boxShadow: "0 8px 24px rgba(60,45,25,.12)",
              }}>
                {nav.filter((n) => n.label.toLowerCase().includes(jump.toLowerCase())).slice(0, 5).map((n) => (
                  <button key={n.id} onClick={() => { setView(n.id); setJump(""); }} style={{
                    display: "flex", alignItems: "center", gap: 8, width: "100%", padding: "7px 11px",
                    border: "none", background: "transparent", cursor: "pointer", fontFamily: "inherit",
                    fontSize: 12.5, color: C.ink, textAlign: "left",
                  }}
                    onMouseEnter={(e) => e.currentTarget.style.background = C.panel2}
                    onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
                    <n.icon size={13} color={C.faint} /> {n.label}
                  </button>
                ))}
              </div>
            )}
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <span style={{
              width: 7, height: 7, borderRadius: 99,
              background: syncState === "synced" ? "#5E7D54" : syncState === "loading" ? "#C8932E" : "#B4574B",
            }} />
            <span style={{ fontSize: 11.5, color: C.faint, fontWeight: 500 }}>
              {syncState === "loading" ? "Syncing" : syncState === "synced" ? "Synced" : "Local"}
            </span>
          </div>
          <span style={{ fontSize: 11.5, color: C.faint }}>
            {TODAY.toLocaleDateString("en-US", { month: "short", day: "numeric" })}
          </span>
          {/* identity — pick who you are; personalizes greeting, tasks, discussion */}
          <div style={{ position: "relative" }}>
            <button onClick={() => setIdOpen(!idOpen)} style={{
              display: "flex", alignItems: "center", gap: 7, border: `1px solid ${C.borderSoft}`,
              background: me ? C.panel2 : C.panel, borderRadius: 6, padding: "4px 9px 4px 5px",
              cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 500, color: C.ink,
            }}>
              <span style={{
                width: 20, height: 20, borderRadius: 99, background: me ? C.accent : C.panel2,
                border: me ? "none" : `1px solid ${C.border}`,
                display: "inline-flex", alignItems: "center", justifyContent: "center",
                fontSize: 9.5, fontWeight: 700, color: me ? "#fff" : C.faint,
              }}>{me ? me.split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase() : <UserRound size={11} />}</span>
              {me || "Sign in"}
            </button>
            {idOpen && (
              <div style={{
                position: "absolute", top: 32, right: 0, width: 200, background: C.panel,
                border: `1px solid ${C.border}`, borderRadius: 8, zIndex: 60, padding: 8,
                boxShadow: "0 8px 24px rgba(60,45,25,.12)",
              }}>
                <Micro style={{ display: "block", padding: "2px 4px 7px" }}>Sign in as</Micro>
                {settings.founders.map((f) => (
                  <button key={f} onClick={() => { setMe(f); setIdOpen(false); }} style={{
                    display: "flex", alignItems: "center", gap: 8, width: "100%", padding: "6px 8px",
                    border: "none", background: me === f ? C.panel2 : "transparent", borderRadius: 6,
                    cursor: "pointer", fontFamily: "inherit", fontSize: 12.5, color: C.ink, textAlign: "left",
                  }}>
                    <span style={{ width: 18, height: 18, borderRadius: 99, background: C.accent, color: "#fff", fontSize: 9, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                      {f.split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase()}
                    </span>
                    {f}
                  </button>
                ))}
                <IdentityAdd onAdd={(n) => { addFounder(n); setMe(n); setIdOpen(false); }} />
                {me && (
                  <button onClick={() => { setMe(""); setIdOpen(false); onSignOut && onSignOut(); }} style={{
                    width: "100%", padding: "6px 8px", marginTop: 4, border: "none", background: "transparent",
                    borderRadius: 6, cursor: "pointer", fontFamily: "inherit", fontSize: 12, color: C.dim, textAlign: "left",
                  }}>Sign out</button>
                )}
              </div>
            )}
          </div>
        </header>

        {/* ---------------- Main ---------------- */}
        <main style={{ flex: 1, overflowY: "auto", padding: "22px 28px 50px" }}>
        {view === "dashboard" && <Dashboard {...{ data, setData, runway, openTasks, priorities, setModal, toggleTask, cyclePriority, deleteTask, setView, me }} />}
        {view === "tasks" && <TasksHub {...{ data, setData, toggleTask, cyclePriority, deleteTask, patchTask, quickAddTask, setModal, me, founders: settings.founders }} />}
        {view === "product" && <ProductHub {...{ data, setData, setModal, founders: settings.founders }} />}
        {view === "financials" && <Financials {...{ data, setData, setModal, founders: settings.founders }} />}
        {view === "gtm" && <Strategy {...{ data, setData, setModal, me, setMe, founders: settings.founders, addFounder }} />}
        {view === "settings" && <SettingsPage {...{ settings, updSettings, nav, data, setData }} />}
        </main>
      </div>

      {copilotOpen && <Copilot data={data} onClose={() => setCopilotOpen(false)} />}

      {/* ---------------- Modals ---------------- */}
      <Modal open={modal === "task"} onClose={() => setModal(null)} title="Add Task">
        <Field label="Title" placeholder="What needs doing?" onChange={(e) => setForm({ ...form, title: e.target.value })} />
        <Field label="Owner" placeholder="Who owns it?" onChange={(e) => setForm({ ...form, owner: e.target.value })} />
        <Field label="Due date" type="date" onChange={(e) => setForm({ ...form, due: e.target.value })} />
        <Select label="Priority" options={["Critical", "High", "Medium", "Low"]} onChange={(e) => setForm({ ...form, priority: e.target.value })} />
        <Select label="Module" options={["General", "Product", "Manufacturing", "Fundraising", "Regulatory", "Marketing"]} onChange={(e) => setForm({ ...form, link: e.target.value })} />
        <Btn primary onClick={() => addRecord("task")}>Create Task</Btn>
      </Modal>

      <Modal open={modal === "vendor"} onClose={() => setModal(null)} title="Log Supplier">
        <Field label="Vendor name" onChange={(e) => setForm({ ...form, name: e.target.value })} />
        <Select label="Type" options={["Co-packer", "Ingredient supplier", "Packaging vendor", "Sleeve printer", "Freight provider"]} onChange={(e) => setForm({ ...form, type: e.target.value })} />
        <Field label="Contact" onChange={(e) => setForm({ ...form, contact: e.target.value })} />
        <Field label="Pricing" onChange={(e) => setForm({ ...form, price: e.target.value })} />
        <Field label="Lead time" onChange={(e) => setForm({ ...form, lead: e.target.value })} />
        <Field label="MOQ" onChange={(e) => setForm({ ...form, moq: e.target.value })} />
        <Btn primary onClick={() => addRecord("vendor")}>Add Vendor</Btn>
      </Modal>

      <Modal open={modal === "investor"} onClose={() => setModal(null)} title="Add Investor">
        <Field label="Name" onChange={(e) => setForm({ ...form, name: e.target.value })} />
        <Field label="Firm" onChange={(e) => setForm({ ...form, firm: e.target.value })} />
        <Field label="Check size" onChange={(e) => setForm({ ...form, check: e.target.value })} />
        <Field label="Focus area" onChange={(e) => setForm({ ...form, focus: e.target.value })} />
        <Field label="Contact" onChange={(e) => setForm({ ...form, contact: e.target.value })} />
        <Select label="Interest" options={["High", "Medium", "Low"]} onChange={(e) => setForm({ ...form, interest: e.target.value })} />
        <Btn primary onClick={() => addRecord("investor")}>Add Investor</Btn>
      </Modal>

      <Modal open={modal === "schedule"} onClose={() => setModal(null)} title="Schedule a Meeting">
        <Field label="Title" placeholder="Founders meeting" onChange={(e) => setForm({ ...form, title: e.target.value })} />
        <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
          <Field label="Date" type="date" onChange={(e) => setForm({ ...form, date: e.target.value })} />
          <Field label="Time" type="time" onChange={(e) => setForm({ ...form, time: e.target.value })} />
        </div>
        <Field label="Participants" placeholder="All founders" onChange={(e) => setForm({ ...form, parts: e.target.value })} />
        <div style={{ marginBottom: 13 }}>
          <label style={{ fontSize: 11.5, color: C.dim, fontWeight: 700, display: "block", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.6 }}>Agenda (one item per line)</label>
          <textarea rows={4} placeholder={"Beverage feedback form\nBranding discussion"} onChange={(e) => setForm({ ...form, agenda: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} />
        </div>
        <Btn primary onClick={() => addRecord("schedule")}>Schedule Meeting</Btn>
      </Modal>

      <Modal open={modal === "idea"} onClose={() => setModal(null)} title="Log Idea">
        <div style={{ marginBottom: 13 }}>
          <label style={{ fontSize: 11.5, color: C.dim, fontWeight: 700, display: "block", marginBottom: 6, textTransform: "uppercase", letterSpacing: 0.6 }}>Idea</label>
          <textarea rows={3} placeholder="What's the idea?" onChange={(e) => setForm({ ...form, text: e.target.value })} style={{ ...inputStyle, resize: "vertical" }} />
        </div>
        <Select label="Category" options={["General", "Product", "Marketing", "Brand", "Packaging", "Flavor"]} onChange={(e) => setForm({ ...form, tag: e.target.value })} />
        <Btn primary onClick={() => addRecord("idea")}>Add to Idea Board</Btn>
      </Modal>
    </div>
  );
}

/* ----------------------------- Page header ----------------------------- */
const PageHead = ({ title, sub, action }) => (
  <div className="fade" style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
    <div>
      <h1 style={{ fontSize: 19, fontWeight: 600, fontFamily: "'STIX Two Text', serif", letterSpacing: -0.2, margin: 0, color: C.ink }}>{title}</h1>
      {sub && <p style={{ color: C.dim, fontSize: 12.5, margin: "4px 0 0" }}>{sub}</p>}
    </div>
    {action}
  </div>
);

/* =================================================================== */
/*  DASHBOARD                                                          */
/* =================================================================== */
function Dashboard({ data, setData, runway, openTasks, priorities, setModal, toggleTask, cyclePriority, deleteTask, setView, me }) {
  const ft = finTotals(data.financials, data.product);
  const setCash = (n) => setData((d) => ({ ...d, company: { ...d.company, cash: n } }));
  const setRunway = (n) => setData((d) => ({ ...d, company: { ...d.company, runwayMonths: n } }));
  const metrics = [
    { label: "Cash balance", node: <EditNum value={data.company.cash} onCommit={setCash} prefix="$" size={18} bold />, icon: DollarSign, color: "#5E7D54" },
    { label: "Expenses to date", node: ft ? fmt$(ft.spent) : "—", icon: Flame, color: C.accent },
    { label: "Capital to launch", node: ft ? fmt$(ft.needed.likely) : "—", icon: TrendingUp, color: "#B4574B" },
    { label: "Runway", node: <EditNum value={Number(runway)} onCommit={setRunway} suffix=" mo" size={18} bold />, icon: Clock, color: runway < 9 ? "#B4574B" : "#5B7A9D" },
    { label: "Open tasks", node: openTasks.length, icon: ClipboardList, color: "#8A6FA8" },
  ];
  const quick = [
    { l: "Add Task", t: "task", d: "Create a to-do with owner & due date" },
    { l: "Log Idea", t: "idea", d: "Capture an idea on the Creative board" },
    { l: "Add Investor", t: "investor", d: "Start tracking an investor conversation" },
    { l: "Schedule Meeting", t: "schedule", d: "Put a meeting on the calendar with an agenda" },
  ];

  return (
    <div className="fade">
      <PageHead title={me ? `Welcome back, ${me.split(" ")[0]}` : "Dashboard"} sub={TODAY.toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric" })} />

      {/* stat strip — one panel, divided columns */}
      <Card style={{ padding: 0, overflow: "hidden", marginBottom: 14 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(5,1fr)" }}>
          {metrics.map((m, i) => (
            <div key={m.label} style={{
              padding: "13px 16px", borderLeft: i > 0 ? `1px solid ${C.borderSoft}` : "none",
            }}>
              <div style={{ display: "flex", alignItems: "center", gap: 6, marginBottom: 7 }}>
                <m.icon size={12} color={m.color} strokeWidth={2} />
                <span style={{ fontSize: 10, color: C.faint, fontWeight: 600, textTransform: "uppercase", letterSpacing: 0.8 }}>{m.label}</span>
              </div>
              <div style={{ fontSize: 18, fontWeight: 600, fontFamily: "'Geist', sans-serif", letterSpacing: -0.3 }}>{m.node}</div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1.4fr 1fr", gap: 12 }}>
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "12px 16px 9px" }}>
            <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
              <AlertCircle size={16} color={C.accent} />
              <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Today's Priorities</span>
            </div>
            <span onClick={() => setView("tasks")} style={{ fontSize: 12.5, color: C.dim, cursor: "pointer", display: "flex", alignItems: "center", gap: 3, fontWeight: 600 }}>All tasks <ChevronRight size={14} /></span>
          </div>
          <div>
            {priorities.slice(0, 6).map((t) => {
              const overdue = t.d < 0;
              return (
                <div key={t.id} style={{ display: "flex", alignItems: "center", gap: 12, padding: "9px 16px", borderTop: `1px solid ${C.borderSoft}` }}>
                  <Circle size={17} color={C.faint} style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => toggleTask(t.id)} />
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontSize: 13.5, fontWeight: 600, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>{t.title}</div>
                    <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}><span style={{ color: me && t.owner === me ? C.accent : C.faint, fontWeight: me && t.owner === me ? 600 : 400 }}>{t.owner}</span> · {t.link}</div>
                  </div>
                  <span onClick={() => cyclePriority(t.id)} style={{ cursor: "pointer" }} title="Click to change priority">
                    <Badge label={t.priority} color={PRIORITY[t.priority]} />
                  </span>
                  <span style={{ fontSize: 11.5, fontWeight: 700, color: overdue ? "#B4574B" : C.dim, width: 70, textAlign: "right", flexShrink: 0 }}>
                    {overdue ? `${-t.d}d overdue` : t.d === 0 ? "Today" : `${t.d}d`}
                  </span>
                  <Trash2 size={14} color={C.faint} style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => deleteTask(t.id)} />
                </div>
              );
            })}
          </div>
        </Card>

        <Card style={{ padding: 0, overflow: "hidden", alignSelf: "start" }}>
          <div style={{ padding: "12px 16px 9px" }}>
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Quick Actions</span>
          </div>
          <div>
            {quick.map((q) => (
              <button key={q.t} onClick={() => setModal(q.t)} style={{
                display: "flex", alignItems: "center", gap: 13, width: "100%", textAlign: "left",
                padding: "10px 16px", border: "none", borderTop: `1px solid ${C.borderSoft}`,
                background: "transparent", cursor: "pointer", fontFamily: "inherit",
                transition: "background .12s",
              }}
                onMouseEnter={(e) => e.currentTarget.style.background = C.panel2}
                onMouseLeave={(e) => e.currentTarget.style.background = "transparent"}>
                <span style={{
                  width: 28, height: 28, borderRadius: 6, background: C.panel2, border: `1px solid ${C.borderSoft}`,
                  display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
                }}><Plus size={14} color={C.accent} strokeWidth={2.6} /></span>
                <span style={{ flex: 1 }}>
                  <span style={{ display: "block", fontSize: 13.5, fontWeight: 600, color: C.ink }}>{q.l}</span>
                  <span style={{ display: "block", fontSize: 11.5, color: C.faint, marginTop: 2 }}>{q.d}</span>
                </span>
                <ChevronRight size={15} color={C.faint} />
              </button>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  TASKS                                                              */
/* =================================================================== */
/* underline section tabs shared by the hub pages */
function SectionTabs({ tabs, tab, setTab }) {
  return (
    <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.borderSoft}`, marginBottom: 18 }}>
      {tabs.map((t) => (
        <button key={t.id} onClick={() => setTab(t.id)} style={{
          padding: "7px 13px", border: "none", cursor: "pointer", background: "transparent",
          color: tab === t.id ? C.ink : C.dim, fontFamily: "inherit", fontSize: 13,
          fontWeight: tab === t.id ? 600 : 500,
          borderBottom: tab === t.id ? `2px solid ${C.accent}` : "2px solid transparent", marginBottom: -1,
        }}>{t.label}</button>
      ))}
    </div>
  );
}

/* Tasks hub: Tasks · Regulatory · Documents */
function TasksHub(props) {
  const [tab, setTab] = useState("tasks");
  return (
    <div>
      <SectionTabs tab={tab} setTab={setTab} tabs={[
        { id: "tasks", label: "Tasks" }, { id: "regulatory", label: "Regulatory" }, { id: "docs", label: "Documents" },
      ]} />
      {tab === "tasks" && <Tasks {...props} />}
      {tab === "regulatory" && <Regulatory data={props.data} setData={props.setData} founders={props.founders} />}
      {tab === "docs" && <Docs data={props.data} />}
    </div>
  );
}

/* Product hub: The Drink · Manufacturing */
function ProductHub(props) {
  const [tab, setTab] = useState("drink");
  return (
    <div>
      <SectionTabs tab={tab} setTab={setTab} tabs={[
        { id: "drink", label: "The Drink" }, { id: "manufacturing", label: "Manufacturing" },
      ]} />
      {tab === "drink" && <Product data={props.data} setData={props.setData} />}
      {tab === "manufacturing" && <Manufacturing data={props.data} setData={props.setData} setModal={props.setModal} founders={props.founders} />}
    </div>
  );
}

function Tasks({ data, toggleTask, cyclePriority, deleteTask, patchTask, quickAddTask, setModal, me, founders = [] }) {
  const [viewMode, setViewMode] = useState("table"); // table | board
  const [mine, setMine] = useState(false);
  const [q, setQ] = useState("");
  const [person, setPerson] = useState("All");
  const [sortDue, setSortDue] = useState(false);
  const [collapsed, setCollapsed] = useState({});

  const GROUPS = [
    ["Product", "#5B7A9D"], ["Manufacturing", "#C26A35"], ["Regulatory", "#C8932E"],
    ["Fundraising", "#5E7D54"], ["Marketing", "#8A6FA8"], ["General", "#A89C86"],
  ];
  const STATUS_CYCLE = ["Todo", "Doing", "Done"];
  const owners = ["All", ...new Set(data.tasks.map((t) => t.owner).filter((o) => o && o !== "—"))];

  const visible = (t) =>
    (!mine || !me || t.owner === me) &&
    (person === "All" || t.owner === person) &&
    (!q.trim() || t.title.toLowerCase().includes(q.toLowerCase()));

  const sorted = (arr) => sortDue ? [...arr].sort((a, b) => new Date(a.due) - new Date(b.due)) : arr;

  const cellL = { padding: "7px 12px", fontSize: 13, borderTop: `1px solid ${C.hairline}`, background: C.panel };

  return (
    <div className="fade">
      <PageHead title="Tasks" sub="Grouped by module — click any status, priority, owner or due date to change it"
        action={<Btn primary icon={Plus} onClick={() => setModal("task")}>New task</Btn>} />

      {/* view tabs */}
      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.borderSoft}`, marginBottom: 12 }}>
        {[["table", "Main table"], ["board", "Board"]].map(([id, label]) => (
          <button key={id} onClick={() => setViewMode(id)} style={{
            padding: "7px 13px", border: "none", cursor: "pointer", background: "transparent",
            color: viewMode === id ? C.ink : C.dim, fontFamily: "inherit", fontSize: 13,
            fontWeight: viewMode === id ? 600 : 500,
            borderBottom: viewMode === id ? `2px solid ${C.accent}` : "2px solid transparent", marginBottom: -1,
          }}>{label}</button>
        ))}
      </div>

      {/* toolbar */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 16, flexWrap: "wrap" }}>
        <div style={{ position: "relative" }}>
          <Search size={13} color={C.faint} style={{ position: "absolute", left: 9, top: 7.5 }} />
          <input value={q} placeholder="Search tasks…" onChange={(e) => setQ(e.target.value)}
            style={{ ...inputStyle, width: 180, padding: "5px 10px 5px 28px", fontSize: 12.5 }} />
        </div>
        <select value={person} onChange={(e) => setPerson(e.target.value)} style={{ ...inputStyle, width: 130, padding: "5px 10px", fontSize: 12.5 }}>
          {owners.map((o) => <option key={o}>{o}</option>)}
        </select>
        <Btn small onClick={() => setSortDue(!sortDue)}>{sortDue ? "✓ Sorted by due" : "Sort by due"}</Btn>
        {me && <Btn small onClick={() => setMine(!mine)}>{mine ? `✓ ${me.split(" ")[0]}'s tasks` : "My tasks"}</Btn>}
      </div>

      {viewMode === "table" && GROUPS.map(([g, color]) => {
        const items = sorted(data.tasks.filter((t) => t.link === g && visible(t)));
        const open = !collapsed[g];
        return (
          <div key={g} style={{ marginBottom: 18 }}>
            <button onClick={() => setCollapsed({ ...collapsed, [g]: open })} style={{
              display: "flex", alignItems: "center", gap: 7, border: "none", background: "transparent",
              cursor: "pointer", fontFamily: "inherit", padding: "0 0 7px",
            }}>
              <ChevronRight size={14} color={color} style={{ transform: open ? "rotate(90deg)" : "none", transition: "transform .12s" }} />
              <span style={{ fontSize: 14.5, fontWeight: 700, color }}>{g}</span>
              <span style={{ fontSize: 11.5, color: C.faint, fontWeight: 500 }}>{items.length} task{items.length === 1 ? "" : "s"}</span>
            </button>
            {open && (
              <div style={{ display: "flex" }}>
                <div style={{ width: 4, background: color, borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
                <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
                  <thead><tr style={{ background: C.panel2 }}>
                    <th style={{ width: 34 }}></th>
                    <th style={{ padding: "7px 12px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, textAlign: "left" }}>Task</th>
                    <th style={{ padding: "7px 12px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, textAlign: "left", width: 130 }}>Owner</th>
                    <th style={{ padding: "7px 12px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, textAlign: "left", width: 104 }}>Leads</th>
                    <th style={{ padding: "7px 8px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, width: 92 }}>Status</th>
                    <th style={{ padding: "7px 8px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, width: 92 }}>Priority</th>
                    <th style={{ padding: "7px 12px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: C.dim, textAlign: "left", width: 122 }}>Due</th>
                    <th style={{ width: 32 }}></th>
                  </tr></thead>
                  <tbody>
                    {items.map((t) => {
                      const done = t.status === "Done";
                      return (
                        <tr key={t.id}>
                          <td style={{ ...cellL, textAlign: "center", padding: "7px 6px" }}>
                            {done
                              ? <CheckCircle2 size={15} color="#5E7D54" style={{ cursor: "pointer" }} onClick={() => toggleTask(t.id)} />
                              : <Circle size={15} color={C.faint} style={{ cursor: "pointer" }} onClick={() => toggleTask(t.id)} />}
                          </td>
                          <td style={{ ...cellL, fontWeight: 500, textDecoration: done ? "line-through" : "none", color: done ? C.faint : C.ink }}>
                            <InlineText value={t.title} onCommit={(v) => patchTask(t.id, { title: v })} width={300} />
                            {t.auto && <span style={{ fontSize: 10, color: C.accent, fontWeight: 700, marginLeft: 6 }}>⚡</span>}
                          </td>
                          <td style={{ ...cellL }}>
                            <span style={{ display: "inline-flex", alignItems: "center", gap: 6 }}>
                              <span style={{
                                width: 18, height: 18, borderRadius: 99, flexShrink: 0,
                                background: me && t.owner === me ? C.accent : C.panel2,
                                border: me && t.owner === me ? "none" : `1px solid ${C.border}`,
                                color: me && t.owner === me ? "#fff" : C.faint,
                                fontSize: 8.5, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center",
                              }}>{(t.owner || "—").split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase()}</span>
                              <InlineText value={t.owner || "—"} onCommit={(v) => patchTask(t.id, { owner: v })} width={90} style={{ fontSize: 12.5 }} />
                            </span>
                          </td>
                          <td style={{ ...cellL }}>
                            <Spearheads value={t.spear || []} onChange={(v) => patchTask(t.id, { spear: v })} founders={founders} />
                          </td>
                          <ColorCell label={t.status} color={STAGE_COLORS[t.status]} width={92}
                            onClick={() => patchTask(t.id, { status: STATUS_CYCLE[(STATUS_CYCLE.indexOf(t.status) + 1) % 3] })} />
                          <ColorCell label={t.priority} color={PRIORITY[t.priority]} width={92}
                            onClick={() => cyclePriority(t.id)} />
                          <td style={{ ...cellL }}>
                            <input type="date" value={t.due || ""} onChange={(e) => patchTask(t.id, { due: e.target.value })}
                              style={{ border: "none", background: "transparent", fontFamily: "inherit", fontSize: 12, color: C.dim, outline: "none", cursor: "pointer", width: 110 }} />
                          </td>
                          <td style={{ ...cellL, textAlign: "center", padding: "7px 6px" }}>
                            <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => deleteTask(t.id)} />
                          </td>
                        </tr>
                      );
                    })}
                    <tr>
                      <td style={{ ...cellL }}></td>
                      <td colSpan={7} style={{ ...cellL, padding: "5px 12px" }}>
                        <GroupAdd onAdd={(v) => quickAddTask(v, g)} placeholder="Add task" />
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}

      {viewMode === "board" && (
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 14 }}>
          {["Todo", "Doing", "Done"].map((c) => {
            const items = data.tasks.filter((t) => t.status === c && visible(t));
            return (
              <div key={c}>
                <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 12, padding: "0 4px" }}>
                  <span style={{ width: 7, height: 7, borderRadius: 99, background: STAGE_COLORS[c] }} />
                  <span style={{ fontWeight: 700, fontSize: 12, textTransform: "uppercase", letterSpacing: 1 }}>{c}</span>
                  <span style={{ fontSize: 11.5, color: C.faint, background: C.panel2, padding: "1px 8px", borderRadius: 5, fontWeight: 700 }}>{items.length}</span>
                </div>
                <div style={{ display: "flex", flexDirection: "column", gap: 9 }}>
                  {items.map((t) => (
                    <Card key={t.id} hover style={{ padding: 12 }}>
                      <div style={{ display: "flex", gap: 9 }}>
                        {t.status === "Done" ? <CheckCircle2 size={16} color="#5E7D54" style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => toggleTask(t.id)} />
                          : <Circle size={16} color={C.faint} style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => toggleTask(t.id)} />}
                        <div style={{ flex: 1 }}>
                          <div style={{ fontSize: 13, fontWeight: 500, textDecoration: t.status === "Done" ? "line-through" : "none", color: t.status === "Done" ? C.faint : C.ink }}>{t.title}</div>
                          <div style={{ display: "flex", alignItems: "center", gap: 8, marginTop: 8, flexWrap: "wrap" }}>
                            <span onClick={() => cyclePriority(t.id)} style={{ cursor: "pointer" }}>
                              <Badge label={t.priority} color={PRIORITY[t.priority]} />
                            </span>
                            <span style={{ fontSize: 11, color: C.faint, fontWeight: 600 }}>{t.link} · {t.owner}</span>
                            <Spearheads value={t.spear || []} onChange={(v) => patchTask(t.id, { spear: v })} founders={founders} size={16} />
                          </div>
                        </div>
                        <Trash2 size={13} color={C.faint} style={{ cursor: "pointer", flexShrink: 0, marginTop: 2 }} onClick={() => deleteTask(t.id)} />
                      </div>
                    </Card>
                  ))}
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

/* "+ Add task" inline row, Monday-style */
function GroupAdd({ onAdd, placeholder = "Add item" }) {
  const [v, setV] = useState("");
  const go = () => { if (!v.trim()) return; onAdd(v.trim()); setV(""); };
  return (
    <div style={{ display: "flex", alignItems: "center", gap: 7 }}>
      <Plus size={13} color={C.faint} />
      <input value={v} placeholder={placeholder} onChange={(e) => setV(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && go()}
        style={{ border: "none", background: "transparent", fontFamily: "inherit", fontSize: 12.5, color: C.ink, outline: "none", flex: 1 }} />
    </div>
  );
}

/* =================================================================== */
/*  PRODUCT — one drink, the whole page                                */
/* =================================================================== */
function Product({ data, setData }) {
  const p = data.product;
  const fin = data.financials;
  const [newIng, setNewIng] = useState("");
  const [newAmt, setNewAmt] = useState("");

  const updProduct = (patch) => setData((d) => ({ ...d, product: { ...d.product, ...patch } }));
  const updIng = (id, patch) => updProduct({ ingredients: p.ingredients.map((i) => i.id === id ? { ...i, ...patch } : i) });
  const delIng = (id) => updProduct({ ingredients: p.ingredients.filter((i) => i.id !== id) });
  const addIng = () => {
    if (!newIng.trim()) return;
    updProduct({ ingredients: [...p.ingredients, { id: Date.now(), name: newIng.trim(), amount: newAmt.trim() || "—", cost: 0 }] });
    setNewIng(""); setNewAmt("");
  };

  const ingCan = ingPerCan(p);
  const cans = fin?.productionRun?.cans || 10000;
  const runCost = Math.round(ingCan * cans);

  const updStack = (layer, items) => updProduct({ stack: p.stack.map((s) => s.layer === layer ? { ...s, items } : s) });

  const cell = { padding: "8px 14px", fontSize: 13, borderTop: `1px solid ${C.hairline}` };
  const head = { padding: "9px 14px", fontSize: 10.5, fontWeight: 700, letterSpacing: 1.2, textTransform: "uppercase", color: C.dim, textAlign: "left" };

  return (
    <div className="fade">
      <PageHead title="Product" sub="The drink — formula, ingredients & cost" />

      {/* hero: the one drink */}
      <Card style={{ marginBottom: 14, background: C.panel }}>
        <div style={{ display: "flex", alignItems: "center", gap: 22 }}>
          {/* can */}
          <div style={{
            width: 64, height: 104, borderRadius: 14, flexShrink: 0,
            background: `linear-gradient(180deg, #E9A876 0%, #F4E4CE 30%, #FCFAF5 55%)`,
            border: `1px solid ${C.border}`, boxShadow: "0 6px 18px rgba(60,45,25,.12)",
            display: "flex", alignItems: "center", justifyContent: "center",
          }}>
            <span style={{ fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontSize: 13, color: C.ink, transform: "rotate(-90deg)", whiteSpace: "nowrap" }}>Eudai.</span>
          </div>
          <div style={{ flex: 1 }}>
            <div style={{ fontSize: 10.5, fontWeight: 700, letterSpacing: 2.4, textTransform: "uppercase", color: C.accent }}>Orange · Peach · Cream</div>
            <h2 style={{ fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontSize: 22, fontWeight: 500, margin: "5px 0 3px", letterSpacing: -0.3 }}>
              {p.name}<span style={{ color: C.accent }}>.</span> <span style={{ fontSize: 18, color: C.dim, fontStyle: "normal" }}>— <InlineText value={p.flavor} onCommit={(v) => updProduct({ flavor: v })} width={190} /></span>
            </h2>
            <div style={{ fontSize: 13, color: C.dim }}>
              <InlineText value={p.size} onCommit={(v) => updProduct({ size: v })} width={130} /> · nootropic energy drink
            </div>
          </div>
          <div style={{ display: "flex", gap: 30 }}>
            <div style={{ textAlign: "right" }}>
              <Micro>Caffeine</Micro>
              <div style={{ fontSize: 21, fontWeight: 600, fontFamily: "'Geist', sans-serif", marginTop: 4 }}>
                <EditNum value={p.caffeine} onCommit={(n) => updProduct({ caffeine: n })} suffix="mg" size={17} bold />
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <Micro>Ingredient cost / can</Micro>
              <div style={{ fontSize: 21, fontWeight: 600, fontFamily: "'Geist', sans-serif", marginTop: 4, color: ingCan > 0 ? C.ink : C.faint }}>
                {ingCan > 0 ? `$${ingCan.toFixed(2)}` : "—"}
              </div>
            </div>
            <div style={{ textAlign: "right" }}>
              <Micro>Per {cans.toLocaleString()}-can run</Micro>
              <div style={{ fontSize: 21, fontWeight: 600, fontFamily: "'Geist', sans-serif", marginTop: 4, color: runCost > 0 ? C.ink : C.faint }}>
                {runCost > 0 ? fmt$(runCost) : "—"}
              </div>
            </div>
          </div>
        </div>
      </Card>

      <div style={{ display: "grid", gridTemplateColumns: "1.5fr 1fr", gap: 14, alignItems: "start" }}>
        {/* ---------- Ingredients (drives Financials) ---------- */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px 5px" }}>
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Ingredients</span>
            <div style={{ fontSize: 11.5, color: C.faint, marginTop: 4 }}>
              Click any name, amount, or cost to edit. Per-can costs roll up to the Financials production run — adding or removing an ingredient changes the estimated cost everywhere.
            </div>
          </div>
          <div style={{ padding: "10px 16px 14px", display: "flex", gap: 7 }}>
            <input value={newIng} placeholder="Add ingredient…" onChange={(e) => setNewIng(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addIng()} style={{ ...inputStyle, flex: 1, width: "auto" }} />
            <input value={newAmt} placeholder="Amount" onChange={(e) => setNewAmt(e.target.value)} onKeyDown={(e) => e.key === "Enter" && addIng()} style={{ ...inputStyle, width: 92 }} />
            <button onClick={addIng} style={{ width: 34, borderRadius: 6, border: "none", background: C.ink, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}><Plus size={14} color="#FAF7F2" /></button>
          </div>
          <table style={{ width: "100%", borderCollapse: "collapse" }}>
            <thead><tr style={{ background: C.panel2 }}>
              <th style={head}>Ingredient</th><th style={head}>Amount</th><th style={{ ...head, textAlign: "right" }}>Cost / can ($)</th><th style={{ ...head, width: 30 }}></th>
            </tr></thead>
            <tbody>
              {p.ingredients.map((i) => (
                <tr key={i.id}>
                  <td style={{ ...cell, fontWeight: 500 }}><InlineText value={i.name} onCommit={(v) => updIng(i.id, { name: v })} width={190} /></td>
                  <td style={{ ...cell, color: C.dim }}><InlineText value={i.amount} onCommit={(v) => updIng(i.id, { amount: v })} width={110} /></td>
                  <td style={{ ...cell, textAlign: "right" }}><EditNum value={i.cost || 0} onCommit={(n) => updIng(i.id, { cost: n })} decimals={3} bold /></td>
                  <td style={{ ...cell, textAlign: "center" }}><Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delIng(i.id)} /></td>
                </tr>
              ))}
              <tr style={{ background: C.panel2 }}>
                <td style={{ ...cell, fontWeight: 700 }} colSpan={2}>Ingredient cost per can</td>
                <td style={{ ...cell, textAlign: "right", fontWeight: 700, fontFamily: "'Geist', sans-serif", fontSize: 15 }}>${ingCan.toFixed(2)}</td>
                <td style={cell}></td>
              </tr>
            </tbody>
          </table>
        </Card>

        {/* ---------- Stack by function (editable) ---------- */}
        <Card>
          <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Stack by Function</span>
          <div style={{ fontSize: 11.5, color: C.faint, marginTop: 4 }}>Formula locked — click any line to edit wording.</div>
          <div style={{ marginTop: 14, display: "flex", flexDirection: "column", gap: 12 }}>
            {p.stack.map((s) => (
              <div key={s.layer}>
                <Micro style={{ color: C.accent }}>{s.layer}</Micro>
                <div style={{ fontSize: 13, fontWeight: 500, marginTop: 3, lineHeight: 1.5 }}>
                  <InlineText value={s.items} onCommit={(v) => updStack(s.layer, v)} width={250} />
                </div>
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  MANUFACTURING                                                     */
/* =================================================================== */
function Manufacturing({ data, setData, setModal, founders = [] }) {
  const STAGES = ["Research", "Contacted", "Sample Requested", "Quote Received", "Negotiating", "Approved", "Active Vendor"];
  const GROUPS = [
    ["Co-packer", "#C26A35"], ["Ingredient supplier", "#5E7D54"], ["Packaging vendor", "#5B7A9D"],
    ["Sleeve printer", "#8A6FA8"], ["Freight provider", "#C8932E"],
  ];
  const [collapsed, setCollapsed] = useState({});

  const patchVendor = (id, patch) => setData((d) => ({ ...d, vendors: d.vendors.map((v) => v.id === id ? { ...v, ...patch } : v) }));
  const delVendor = (id) => setData((d) => ({ ...d, vendors: d.vendors.filter((v) => v.id !== id) }));
  const quickAddVendor = (name, type) => setData((d) => ({
    ...d, vendors: [...d.vendors, { id: Date.now(), name, type, contact: "—", price: "—", lead: "—", moq: "—", status: "Research" }],
  }));
  const cycleStatus = (v) => patchVendor(v.id, { status: STAGES[(STAGES.indexOf(v.status) + 1) % STAGES.length] });

  return (
    <div className="fade">
      <PageHead title="Manufacturing" sub="Co-packers, suppliers, packaging & freight — click any cell to edit, click status to advance"
        action={<Btn primary icon={Plus} onClick={() => setModal("vendor")}>Log Supplier</Btn>} />

      {GROUPS.map(([g, color]) => {
        const items = data.vendors.filter((v) => v.type === g);
        const open = !collapsed[g];
        return (
          <div key={g} style={{ marginBottom: 18 }}>
            <button onClick={() => setCollapsed({ ...collapsed, [g]: open })} style={{
              display: "flex", alignItems: "center", gap: 7, border: "none", background: "transparent",
              cursor: "pointer", fontFamily: "inherit", padding: "0 0 7px",
            }}>
              <ChevronRight size={14} color={color} style={{ transform: open ? "rotate(90deg)" : "none", transition: "transform .12s" }} />
              <span style={{ fontSize: 14.5, fontWeight: 700, color }}>{g}s</span>
              <span style={{ fontSize: 11.5, color: C.faint, fontWeight: 500 }}>{items.length}</span>
            </button>
            {open && (
              <div style={{ display: "flex" }}>
                <div style={{ width: 4, background: color, borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
                <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
                  <thead><tr style={{ background: C.panel2 }}>
                    <th style={TH}>Vendor</th>
                    <th style={{ ...TH, width: 150 }}>Contact</th>
                    <th style={{ ...TH, width: 110 }}>Pricing</th>
                    <th style={{ ...TH, width: 95 }}>Lead time</th>
                    <th style={{ ...TH, width: 95 }}>MOQ</th>
                    <th style={{ ...TH, textAlign: "center", width: 128 }}>Status</th>
                    <th style={{ ...TH, width: 100 }}>Leads</th>
                    <th style={{ width: 32 }}></th>
                  </tr></thead>
                  <tbody>
                    {items.map((v) => (
                      <tr key={v.id}>
                        <td style={{ ...TD, fontWeight: 500 }}><InlineText value={v.name} onCommit={(x) => patchVendor(v.id, { name: x })} width={170} /></td>
                        <td style={TD}><InlineText value={v.contact || "—"} onCommit={(x) => patchVendor(v.id, { contact: x })} width={130} style={{ fontSize: 12.5 }} /></td>
                        <td style={TD}><InlineText value={v.price || "—"} onCommit={(x) => patchVendor(v.id, { price: x })} width={90} style={{ fontSize: 12.5 }} /></td>
                        <td style={TD}><InlineText value={v.lead || "—"} onCommit={(x) => patchVendor(v.id, { lead: x })} width={75} style={{ fontSize: 12.5 }} /></td>
                        <td style={TD}><InlineText value={v.moq || "—"} onCommit={(x) => patchVendor(v.id, { moq: x })} width={75} style={{ fontSize: 12.5 }} /></td>
                        <ColorCell label={v.status} color={STAGE_COLORS[v.status]} width={128} onClick={() => cycleStatus(v)} />
                        <td style={TD}><Spearheads value={v.spear || []} onChange={(x) => patchVendor(v.id, { spear: x })} founders={founders} size={16} /></td>
                        <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}>
                          <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delVendor(v.id)} />
                        </td>
                      </tr>
                    ))}
                    <tr>
                      <td colSpan={8} style={{ ...TD, padding: "5px 12px" }}>
                        <GroupAdd onAdd={(x) => quickAddVendor(x, g)} placeholder="Add vendor" />
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            )}
          </div>
        );
      })}
    </div>
  );
}

/* =================================================================== */
/*  FINANCIALS — Revenue · Estimated Costs · Fundraising               */
/* =================================================================== */
function Financials({ data, setData, setModal, founders = [] }) {
  const [tab, setTab] = useState("revenue");
  const fin = data.financials || { pricing: { can: 3.5, twelvePack: 30 }, revenueToDate: 0, setupCosts: [], productionRun: { cans: 10000, steps: [] } };
  if (!fin.market) fin.market = { tamB: 90, tamNote: "", samB: 50, samNote: "", somStudents: 47000, somPct: 40, somCansWeek: 2, somNote: "" };
  if (!fin.projections) fin.projections = [];
  if (!fin.statements) fin.statements = { cans: 10000 };
  const tabs = [
    { id: "revenue", label: "Revenue" },
    { id: "costs", label: "Estimated Costs" },
    { id: "projections", label: "Projections" },
    { id: "market", label: "Market Size" },
    { id: "statements", label: "Statements" },
    { id: "fundraising", label: "Fundraising" },
  ];

  return (
    <div className="fade">
      <PageHead title="Financials" sub="Revenue, cost estimates & investor pipeline" />

      {/* underline tabs */}
      <div style={{
        display: "flex", gap: 2, borderBottom: `1px solid ${C.borderSoft}`, marginBottom: 18,
      }}>
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "7px 13px", border: "none", cursor: "pointer", background: "transparent",
            color: tab === t.id ? C.ink : C.dim, fontFamily: "inherit",
            fontSize: 13, fontWeight: tab === t.id ? 600 : 500,
            borderBottom: tab === t.id ? `2px solid ${C.accent}` : "2px solid transparent",
            marginBottom: -1, transition: "color .12s",
          }}>{t.label}</button>
        ))}
      </div>

      {tab === "revenue" && <RevenueTab fin={fin} product={data.product} setData={setData} />}
      {tab === "costs" && <CostsTab fin={fin} product={data.product} setData={setData} />}
      {tab === "projections" && <ProjectionsTab fin={fin} product={data.product} setData={setData} />}
      {tab === "market" && <MarketTab fin={fin} setData={setData} />}
      {tab === "statements" && <StatementsTab fin={fin} product={data.product} setData={setData} />}
      {tab === "fundraising" && <FundraisingTab {...{ data, setData, setModal, founders }} />}
    </div>
  );
}

/* ---------------- Revenue ---------------- */
function RevenueTab({ fin, product, setData }) {
  const [cans, setCans] = useState(10000);
  const [scenario, setScenario] = useState("likely");
  const ft = finTotals(fin, product);
  const run = fin.productionRun;
  const perCan = { low: ft.prod.low / run.cans, likely: ft.prod.likely / run.cans, high: ft.prod.high / run.cans };
  const gross = cans * fin.pricing.can;
  const cost = cans * perCan[scenario];
  const margin = gross - cost;
  const setPrice = (k, n) => setData((d) => ({ ...d, financials: { ...d.financials, pricing: { ...d.financials.pricing, [k]: n } } }));

  return (
    <div className="fade">
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12, marginBottom: 12 }}>
        <Card>
          <Micro>Revenue to date</Micro>
          <div style={{ fontSize: 21, fontWeight: 600, marginTop: 10, fontFamily: "'Geist', sans-serif" }}>{fmt$(fin.revenueToDate)}</div>
          <div style={{ fontSize: 12, color: C.faint, marginTop: 4 }}>Pre-revenue</div>
        </Card>
        <Card>
          <Micro>Single can</Micro>
          <div style={{ marginTop: 12 }}><EditNum value={fin.pricing.can} onCommit={(n) => setPrice("can", n)} prefix="$" decimals={2} size={21} bold /></div>
          <div style={{ fontSize: 12, color: C.faint, marginTop: 4 }}>Target retail price</div>
        </Card>
        <Card>
          <Micro>12-pack</Micro>
          <div style={{ marginTop: 12 }}><EditNum value={fin.pricing.twelvePack} onCommit={(n) => setPrice("twelvePack", n)} prefix="$" size={21} bold /></div>
          <div style={{ fontSize: 12, color: C.faint, marginTop: 4 }}>${(fin.pricing.twelvePack / 12).toFixed(2)} per can</div>
        </Card>
      </div>

      {/* what-if calculator — computed from your price points and cost sheet */}
      <Card>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 16 }}>
          <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Revenue Scenario</span>
          <div style={{ display: "flex", gap: 6 }}>
            {["low", "likely", "high"].map((s) => (
              <button key={s} onClick={() => setScenario(s)} style={{
                fontSize: 11, fontWeight: 700, textTransform: "capitalize", padding: "5px 12px",
                borderRadius: 6, cursor: "pointer", fontFamily: "inherit",
                border: `1px solid ${scenario === s ? C.accent : C.border}`,
                background: scenario === s ? C.accentSoft : "transparent",
                color: scenario === s ? C.accent : C.dim,
              }}>{s} cost</button>
            ))}
          </div>
        </div>
        <Micro style={{ display: "block", marginBottom: 8 }}>Cans sold @ ${fin.pricing.can.toFixed(2)}</Micro>
        <input type="range" min="1000" max="100000" step="1000" value={cans}
          onChange={(e) => setCans(+e.target.value)}
          style={{ width: "100%", accentColor: C.accent, marginBottom: 16 }} />
        <div style={{ display: "grid", gridTemplateColumns: "repeat(4,1fr)", gap: 14 }}>
          {[
            ["Cans sold", cans.toLocaleString()],
            ["Gross revenue", fmt$(Math.round(gross))],
            [`Production cost ($${perCan[scenario].toFixed(2)}/can)`, fmt$(Math.round(cost))],
            ["Gross margin", fmt$(Math.round(margin))],
          ].map(([l, v], i) => (
            <div key={l}>
              <Micro>{l}</Micro>
              <div style={{
                fontSize: 21, fontWeight: 600, marginTop: 6, fontFamily: "'Geist', sans-serif",
                color: i === 3 ? "#5E7D54" : C.ink,
              }}>{v}</div>
            </div>
          ))}
        </div>
        <div style={{ fontSize: 11.5, color: C.faint, marginTop: 14 }}>
          Computed from your cost sheet's per-can production cost and your ${fin.pricing.can.toFixed(2)} price point. Excludes setup costs, distribution, and retail margin.
        </div>
      </Card>
    </div>
  );
}

/* ---------------- Estimated Costs ---------------- */
/* Inline-editable number — click, type, Enter/blur to commit.
   Used everywhere a number shows: dashboard metrics, cost tables, prices, ingredients. */
/* Inline-editable text — the words version of EditNum. Click, type, Enter/blur commits. */
function InlineText({ value, onCommit, width = 150, style }) {
  const [editing, setEditing] = useState(false);
  const [v, setV] = useState(value);
  if (editing) return (
    <input autoFocus value={v} onChange={(e) => setV(e.target.value)}
      onBlur={() => { setEditing(false); onCommit(v); }}
      onKeyDown={(e) => { if (e.key === "Enter") { setEditing(false); onCommit(v); } if (e.key === "Escape") setEditing(false); }}
      style={{ width, fontSize: 12.5, fontFamily: "inherit", border: `1px solid ${C.accent}`, borderRadius: 6, padding: "2px 6px", background: "#fff", outline: "none", color: C.ink }} />
  );
  return <span onClick={() => { setV(value); setEditing(true); }} title="Click to edit"
    style={{ cursor: "text", borderBottom: `1px dashed ${C.border}`, ...style }}>{value}</span>;
}

function EditNum({ value, onCommit, bold, prefix = "", suffix = "", size, serif, decimals }) {
  const [editing, setEditing] = useState(false);
  const [v, setV] = useState(value);
  const commit = () => { setEditing(false); const n = parseFloat(String(v).replace(/[^0-9.]/g, "")); if (!isNaN(n)) onCommit(n); else setV(value); };
  const fmtV = decimals != null ? Number(value).toFixed(decimals) : (value === 0 ? "0" : Number(value).toLocaleString());
  if (editing) return (
    <input autoFocus value={v}
      onChange={(e) => setV(e.target.value)}
      onBlur={commit}
      onKeyDown={(e) => { if (e.key === "Enter") commit(); if (e.key === "Escape") { setEditing(false); setV(value); } }}
      style={{
        width: size ? size * 4.2 : 76, textAlign: "right", fontSize: size || 13, fontFamily: serif ? "'Geist', sans-serif" : "inherit",
        border: `1px solid ${C.accent}`, borderRadius: 6, padding: "2px 6px",
        background: "#fff", color: C.ink, outline: "none",
      }} />
  );
  return (
    <span onClick={() => { setV(value); setEditing(true); }} title="Click to edit"
      style={{
        cursor: "text", fontWeight: bold ? 600 : undefined, borderBottom: `1px dashed ${C.border}`,
        paddingBottom: 1, fontSize: size, fontFamily: serif ? "'Geist', sans-serif" : undefined,
        letterSpacing: size ? -0.5 : undefined,
      }}>{prefix}{fmtV}{suffix}</span>
  );
}

function CostsTab({ fin, product, setData }) {
  const ft = finTotals(fin, product);
  const run = fin.productionRun;

  const updateSetup = (item, patch) => setData((d) => ({
    ...d, financials: { ...d.financials, setupCosts: d.financials.setupCosts.map((r) => r.item === item ? { ...r, ...patch } : r) }
  }));
  const updateStep = (step, key, val) => setData((d) => ({
    ...d, financials: { ...d.financials, productionRun: { ...d.financials.productionRun, steps: d.financials.productionRun.steps.map((s) => s.step === step ? { ...s, [key]: val } : s) } }
  }));
  const addSetup = (item) => setData((d) => ({
    ...d, financials: { ...d.financials, setupCosts: [...d.financials.setupCosts, { item, cost: 0, spent: false }] }
  }));
  const delSetup = (item) => setData((d) => ({
    ...d, financials: { ...d.financials, setupCosts: d.financials.setupCosts.filter((r) => r.item !== item) }
  }));

  const cell = { padding: "9px 14px", fontSize: 13, borderTop: `1px solid ${C.hairline}` };
  const head = { padding: "9px 14px", fontSize: 10.5, fontWeight: 700, letterSpacing: 1.2, textTransform: "uppercase", color: C.dim, textAlign: "left" };

  return (
    <div className="fade" style={{ display: "flex", flexDirection: "column", gap: 14 }}>
      <div style={{ fontSize: 11.5, color: C.faint }}>Click any number to edit — totals recalculate automatically. Toggle "Paid" when money actually goes out; Expenses and Capital Needed update everywhere.</div>

      {/* Setup & fixed costs */}
      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px 9px" }}>
          <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Setup & Fixed Costs</span>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.panel2 }}>
            <th style={head}>Item</th><th style={{ ...head, textAlign: "right" }}>Cost ($)</th><th style={{ ...head, textAlign: "center" }}>Paid</th><th style={{ width: 32 }}></th>
          </tr></thead>
          <tbody>
            {fin.setupCosts.map((r) => (
              <tr key={r.item} style={{ opacity: r.spent ? 0.75 : 1 }}>
                <td style={cell}><InlineText value={r.item} onCommit={(v) => updateSetup(r.item, { item: v })} width={210} /></td>
                <td style={{ ...cell, textAlign: "right" }}><EditNum bold value={r.cost} onCommit={(n) => updateSetup(r.item, { cost: n })} /></td>
                <td style={{ ...cell, textAlign: "center" }}>
                  <span onClick={() => updateSetup(r.item, { spent: !r.spent })} style={{ cursor: "pointer", display: "inline-flex" }}>
                    {r.spent
                      ? <Badge label="Paid" color="#5E7D54" />
                      : <span style={{ width: 15, height: 15, borderRadius: 4, border: `1.5px solid ${C.border}`, display: "inline-block" }} />}
                  </span>
                </td>
                <td style={{ ...cell, textAlign: "center", padding: "8px 6px" }}>
                  <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delSetup(r.item)} />
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={4} style={{ ...cell, padding: "5px 14px" }}>
                <GroupAdd onAdd={addSetup} placeholder="Add cost item" />
              </td>
            </tr>
            <tr style={{ background: C.panel2 }}>
              <td style={{ ...cell, fontWeight: 700 }}>Total setup</td>
              <td style={{ ...cell, textAlign: "right", fontWeight: 700, fontFamily: "'Geist', sans-serif", fontSize: 15 }}>{ft.setup.toLocaleString()}</td>
              <td style={{ ...cell, textAlign: "center", fontSize: 11.5, color: C.dim, fontWeight: 600 }}>{fmt$(ft.spent)} paid</td><td style={cell}></td>
            </tr>
          </tbody>
        </table>
      </Card>

      {/* Production run */}
      <Card style={{ padding: 0, overflow: "hidden" }}>
        <div style={{ padding: "12px 16px 9px" }}>
          <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>
            Production Run — <EditNum value={run.cans} onCommit={(n) => setData((d) => ({ ...d, financials: { ...d.financials, productionRun: { ...d.financials.productionRun, cans: Math.max(1, Math.round(n)) } } }))} size={14} bold /> Sleeved & Filled Cans
          </span>
        </div>
        <table style={{ width: "100%", borderCollapse: "collapse" }}>
          <thead><tr style={{ background: C.panel2 }}>
            <th style={head}>Step</th>
            <th style={{ ...head, textAlign: "right" }}>Low ($)</th>
            <th style={{ ...head, textAlign: "right" }}>Likely ($)</th>
            <th style={{ ...head, textAlign: "right" }}>High ($)</th>
          </tr></thead>
          <tbody>
            {run.steps.map((s) => {
              const derived = s.step === "Ingredients" && ft.ingDerived != null;
              if (derived) return (
                <tr key={s.step} style={{ background: C.accentSoft + "44" }}>
                  <td style={cell}>
                    {s.step}{" "}
                    <span style={{ fontSize: 10, fontWeight: 700, color: C.accent, letterSpacing: 0.5, textTransform: "uppercase" }}>· from Product (${ft.ingCan.toFixed(2)}/can)</span>
                  </td>
                  {["low", "likely", "high"].map((k) => (
                    <td key={k} style={{ ...cell, textAlign: "right", fontWeight: k === "likely" ? 600 : 400 }}>{ft.ingDerived.toLocaleString()}</td>
                  ))}
                </tr>
              );
              return (
                <tr key={s.step}>
                  <td style={cell}><InlineText value={s.step} onCommit={(v) => updateStep(s.step, "step", v)} width={210} /></td>
                  <td style={{ ...cell, textAlign: "right" }}><EditNum value={s.low} onCommit={(n) => updateStep(s.step, "low", n)} /></td>
                  <td style={{ ...cell, textAlign: "right" }}><EditNum bold value={s.likely} onCommit={(n) => updateStep(s.step, "likely", n)} /></td>
                  <td style={{ ...cell, textAlign: "right" }}><EditNum value={s.high} onCommit={(n) => updateStep(s.step, "high", n)} /></td>
                </tr>
              );
            })}
            <tr style={{ background: C.panel2 }}>
              <td style={{ ...cell, fontWeight: 700 }}>Total production run</td>
              {["low", "likely", "high"].map((k) => (
                <td key={k} style={{ ...cell, textAlign: "right", fontWeight: 700 }}>{ft.prod[k].toLocaleString()}</td>
              ))}
            </tr>
            <tr>
              <td style={cell}>Per-can cost (÷ {run.cans.toLocaleString()})</td>
              {["low", "likely", "high"].map((k) => (
                <td key={k} style={{ ...cell, textAlign: "right" }}>${(ft.prod[k] / run.cans).toFixed(2)}</td>
              ))}
            </tr>
          </tbody>
        </table>
      </Card>

      {/* Capital still needed */}
      <Card style={{ borderLeft: `2px solid ${C.accent}` }}>
        <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
          <div>
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Capital to Launch</span>
            <div style={{ fontSize: 12, color: C.dim, marginTop: 4 }}>
              Setup + production, minus {fmt$(ft.spent)} already paid
            </div>
          </div>
          <div style={{ display: "flex", gap: 30 }}>
            {["low", "likely", "high"].map((k) => (
              <div key={k} style={{ textAlign: "right" }}>
                <Micro>{k}</Micro>
                <div style={{ fontSize: 21, fontWeight: 600, fontFamily: "'Geist', sans-serif", marginTop: 3, color: k === "likely" ? "#B4574B" : C.ink }}>
                  {fmt$(ft.needed[k])}
                </div>
                <div style={{ fontSize: 10.5, color: C.faint, marginTop: 2 }}>of {fmt$(ft.setup + ft.prod[k])} total</div>
              </div>
            ))}
          </div>
        </div>
      </Card>
    </div>
  );
}

/* ---------------- Fundraising ---------------- */
/* ------------------------------------------------------------------ */
/*  Market math shared by Market Size, Projections & export            */
/* ------------------------------------------------------------------ */
function marketMath(fin) {
  const m = fin.market || { tamB: 90, samB: 50, somStudents: 47000, somPct: 40, somCansWeek: 2 };
  const price = fin.pricing?.can || 3.5;
  const drinkers = Math.round(m.somStudents * (m.somPct / 100));
  const somCans = Math.round(drinkers * m.somCansWeek * 52);
  const somDollars = Math.round(somCans * price);
  return { m, price, drinkers, somCans, somDollars };
}

/* small editable note paragraph */
function EditNote({ value, onCommit }) {
  const [editing, setEditing] = useState(false);
  const [v, setV] = useState(value || "");
  if (editing) return (
    <textarea autoFocus rows={3} value={v} onChange={(e) => setV(e.target.value)}
      onBlur={() => { setEditing(false); onCommit(v); }}
      style={{ ...inputStyle, resize: "vertical", fontSize: 12.5 }} />
  );
  return <p onClick={() => { setV(value || ""); setEditing(true); }} title="Click to edit reasoning"
    style={{ fontSize: 12.5, lineHeight: 1.6, color: C.dim, margin: 0, cursor: "text" }}>{value}</p>;
}

/* ---------------- Market Size: TAM / SAM / SOM ---------------- */
function MarketTab({ fin, setData }) {
  const { m, price, drinkers, somCans, somDollars } = marketMath(fin);
  const upd = (patch) => setData((d) => ({ ...d, financials: { ...d.financials, market: { ...(d.financials.market || m), ...patch } } }));

  const Ring = ({ label, color, value, sub, note, onNote, children }) => (
    <Card style={{ borderLeft: `3px solid ${color}`, marginBottom: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 20, flexWrap: "wrap" }}>
        <div style={{ minWidth: 220 }}>
          <Micro style={{ color }}>{label}</Micro>
          <div style={{ fontSize: 26, fontWeight: 600, fontFamily: "'Geist', sans-serif", letterSpacing: -0.5, marginTop: 6 }}>{value}</div>
          <div style={{ fontSize: 12, color: C.faint, marginTop: 3 }}>{sub}</div>
        </div>
        <div style={{ flex: 1, minWidth: 280 }}>
          <Micro style={{ display: "block", marginBottom: 6 }}>How this number was reasoned</Micro>
          <EditNote value={note} onCommit={onNote} />
          {children}
        </div>
      </div>
    </Card>
  );

  return (
    <div className="fade">
      <div style={{ fontSize: 12.5, color: C.dim, marginBottom: 14 }}>
        Concentric market view for the pitch: TAM ⊃ SAM ⊃ SOM. Every number and explanation is editable — the TAM and SAM figures are public ballpark estimates, so verify them before they go in a deck. SOM is computed live from the assumptions below.
      </div>

      <Ring label="TAM — Total Addressable Market" color="#5B7A9D"
        value={<EditNum value={m.tamB} onCommit={(n) => upd({ tamB: n })} prefix="$" suffix="B" size={26} bold />}
        sub="Global energy drink market"
        note={m.tamNote} onNote={(v) => upd({ tamNote: v })} />

      <Ring label="SAM — Serviceable Addressable Market" color="#5E7D54"
        value={<EditNum value={m.samB} onCommit={(n) => upd({ samB: n })} prefix="$" suffix="B" size={26} bold />}
        sub="US functional / healthy beverage space"
        note={m.samNote} onNote={(v) => upd({ samNote: v })} />

      <Ring label="SOM — Serviceable Obtainable Market" color={C.accent}
        value={fmtBig(somDollars)}
        sub={`USC beachhead · ${somCans.toLocaleString()} cans/yr · computed live`}
        note={m.somNote} onNote={(v) => upd({ somNote: v })}>
        <div style={{ display: "flex", gap: 22, marginTop: 12, flexWrap: "wrap" }}>
          <div><Micro>USC students</Micro><div style={{ fontSize: 15, fontWeight: 600, marginTop: 3 }}><EditNum value={m.somStudents} onCommit={(n) => upd({ somStudents: n })} bold /></div></div>
          <div><Micro>% energy drinkers</Micro><div style={{ fontSize: 15, fontWeight: 600, marginTop: 3 }}><EditNum value={m.somPct} onCommit={(n) => upd({ somPct: n })} suffix="%" bold /></div></div>
          <div><Micro>Cans / week</Micro><div style={{ fontSize: 15, fontWeight: 600, marginTop: 3 }}><EditNum value={m.somCansWeek} onCommit={(n) => upd({ somCansWeek: n })} bold /></div></div>
          <div><Micro>Price / can</Micro><div style={{ fontSize: 15, fontWeight: 600, marginTop: 3 }}>${price.toFixed(2)}</div></div>
          <div><Micro>= Drinkers</Micro><div style={{ fontSize: 15, fontWeight: 600, marginTop: 3, color: C.dim }}>{drinkers.toLocaleString()}</div></div>
        </div>
      </Ring>
    </div>
  );
}

/* ---------------- Projections: the financial plan ---------------- */
function ProjectionsTab({ fin, product, setData }) {
  const ft = finTotals(fin, product);
  const { somCans, somDollars, price } = marketMath(fin);
  const perCan = ft.prod.likely / ((fin.productionRun && fin.productionRun.cans) || 1);
  const rows = fin.projections || [];

  const patchRow = (id, patch) => setData((d) => ({ ...d, financials: { ...d.financials, projections: (d.financials.projections || []).map((r) => r.id === id ? { ...r, ...patch } : r) } }));
  const delRow = (id) => setData((d) => ({ ...d, financials: { ...d.financials, projections: (d.financials.projections || []).filter((r) => r.id !== id) } }));
  const addRow = (label) => setData((d) => ({ ...d, financials: { ...d.financials, projections: [...(d.financials.projections || []), { id: Date.now(), label, capture: 5 }] } }));

  const calc = (r) => {
    const cans = Math.round(somCans * (r.capture / 100));
    const revenue = Math.round(cans * price);
    const cogs = Math.round(cans * perCan);
    const gross = revenue - cogs;
    const runs = Math.ceil(cans / (fin.productionRun.cans || 10000));
    return { cans, revenue, cogs, gross, margin: revenue > 0 ? Math.round((gross / revenue) * 100) : 0, runs };
  };

  const numTd = { ...TD, textAlign: "right", fontFamily: "'Geist', sans-serif", fontVariantNumeric: "tabular-nums" };

  return (
    <div className="fade">
      <Card style={{ marginBottom: 12, borderLeft: `2px solid ${C.accent}` }}>
        <Micro style={{ color: C.accent }}>The plan</Micro>
        <div style={{ fontSize: 13.5, lineHeight: 1.65, marginTop: 6 }}>
          Each year captures an editable % of the USC SOM (<b>{somCans.toLocaleString()} cans / {fmtBig(somDollars)} per year</b> — set on the Market Size tab).
          Unit cost uses the likely production case (<b>${perCan.toFixed(2)}/can</b>) and your <b>${price.toFixed(2)}</b> price.
        </div>
      </Card>

      <div style={{ display: "flex" }}>
        <div style={{ width: 4, background: C.accent, borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
        <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
          <thead><tr style={{ background: C.panel2 }}>
            <th style={TH}>Period</th>
            <th style={{ ...TH, textAlign: "right", width: 110 }}>Capture of SOM</th>
            <th style={{ ...TH, textAlign: "right", width: 110 }}>Cans sold</th>
            <th style={{ ...TH, textAlign: "right", width: 110 }}>Revenue</th>
            <th style={{ ...TH, textAlign: "right", width: 110 }}>COGS</th>
            <th style={{ ...TH, textAlign: "right", width: 110 }}>Gross profit</th>
            <th style={{ ...TH, textAlign: "right", width: 80 }}>Margin</th>
            <th style={{ ...TH, textAlign: "right", width: 96 }}>Prod. runs</th>
            <th style={{ width: 32 }}></th>
          </tr></thead>
          <tbody>
            {rows.map((r) => {
              const c = calc(r);
              return (
                <tr key={r.id}>
                  <td style={{ ...TD, fontWeight: 600 }}><InlineText value={r.label} onCommit={(v) => patchRow(r.id, { label: v })} width={90} /></td>
                  <td style={numTd}><EditNum value={r.capture} onCommit={(n) => patchRow(r.id, { capture: n })} suffix="%" bold /></td>
                  <td style={numTd}>{c.cans.toLocaleString()}</td>
                  <td style={{ ...numTd, fontWeight: 600 }}>{fmt$(c.revenue)}</td>
                  <td style={numTd}>{fmt$(c.cogs)}</td>
                  <td style={{ ...numTd, fontWeight: 600, color: c.gross >= 0 ? "#5E7D54" : "#B4574B" }}>{fmt$(c.gross)}</td>
                  <td style={numTd}>{c.margin}%</td>
                  <td style={numTd}>{c.runs}×</td>
                  <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}>
                    <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delRow(r.id)} />
                  </td>
                </tr>
              );
            })}
            <tr>
              <td colSpan={9} style={{ ...TD, padding: "5px 12px" }}>
                <GroupAdd onAdd={addRow} placeholder="Add period (e.g. Year 4)" />
              </td>
            </tr>
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 11.5, color: C.faint, marginTop: 10 }}>
        Gross only — setup/fixed costs live on the Statements tab. "Prod. runs" = how many {fin.productionRun.cans.toLocaleString()}-can runs that year requires.
      </div>
    </div>
  );
}

/* ---------------- Statements: managerial P&L, Excel-style, exportable ---------------- */
function StatementsTab({ fin, product, setData }) {
  const ft = finTotals(fin, product);
  const st = fin.statements || { cans: 10000 };
  const perCan = ft.prod.likely / ((fin.productionRun && fin.productionRun.cans) || 1);
  const price = fin.pricing?.can || 3.5;

  const updSt = (patch) => setData((d) => ({ ...d, financials: { ...d.financials, statements: { ...(d.financials.statements || st), ...patch } } }));
  const updPrice = (n) => setData((d) => ({ ...d, financials: { ...d.financials, pricing: { ...d.financials.pricing, can: n } } }));
  const updateSetup = (item, patch) => setData((d) => ({ ...d, financials: { ...d.financials, setupCosts: d.financials.setupCosts.map((r) => r.item === item ? { ...r, ...patch } : r) } }));
  const addSetup = (item) => setData((d) => ({ ...d, financials: { ...d.financials, setupCosts: [...d.financials.setupCosts, { item, cost: 0, spent: false }] } }));
  const delSetup = (item) => setData((d) => ({ ...d, financials: { ...d.financials, setupCosts: d.financials.setupCosts.filter((r) => r.item !== item) } }));

  const revenue = Math.round(st.cans * price);
  const cogs = Math.round(st.cans * perCan);
  const gross = revenue - cogs;
  const opex = ft.setup;
  const net = gross - opex;

  const { m, somCans, somDollars } = marketMath(fin);
  const projRows = (fin.projections || []).map((r) => {
    const cans = Math.round(somCans * (r.capture / 100));
    const rev = Math.round(cans * price);
    const cg = Math.round(cans * perCan);
    return [r.label, r.capture / 100, cans, rev, cg, rev - cg];
  });

  const exportXlsx = () => {
    const wb = XLSX.utils.book_new();
    const pl = [
      ["Eudai — Managerial P&L (single period)"], [],
      ["Assumptions"], ["Cans sold", st.cans], ["Price / can", price], ["Unit cost / can (likely)", Number(perCan.toFixed(2))], [],
      ["Income statement"],
      ["Revenue", revenue],
      ["Cost of goods sold", -cogs],
      ["Gross profit", gross],
      ["Gross margin", revenue > 0 ? gross / revenue : 0], [],
      ["Operating / setup expenses"],
      ...fin.setupCosts.map((r) => [r.item + (r.spent ? " (paid)" : ""), -r.cost]),
      ["Total operating expenses", -opex], [],
      ["Net income", net],
    ];
    const proj = [
      ["Eudai — Projections (capture of USC SOM)"], [],
      ["SOM cans / yr", somCans], ["SOM $ / yr", somDollars], [],
      ["Period", "Capture %", "Cans", "Revenue", "COGS", "Gross profit"],
      ...projRows,
    ];
    const market = [
      ["Eudai — Market sizing"], [],
      ["TAM — Global energy drinks ($B)", m.tamB], [m.tamNote], [],
      ["SAM — US functional beverage ($B)", m.samB], [m.samNote], [],
      ["SOM — USC beachhead ($)", somDollars],
      ["USC students", m.somStudents], ["% energy drinkers", m.somPct / 100], ["Cans / week", m.somCansWeek], [m.somNote],
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(pl), "P&L");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(proj), "Projections");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(market), "Market");
    XLSX.writeFile(wb, "Eudai-Financial-Model.xlsx");
  };

  const num = { ...TD, textAlign: "right", fontFamily: "'Geist', sans-serif", fontVariantNumeric: "tabular-nums", width: 150 };
  const sect = { ...TD, fontWeight: 700, fontSize: 11, textTransform: "uppercase", letterSpacing: 1, color: C.dim, background: C.panel2 };
  const totalRow = { ...TD, fontWeight: 700, background: "#FBF8F2", borderTop: `2px solid ${C.border}` };

  return (
    <div className="fade">
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 12, flexWrap: "wrap", gap: 10 }}>
        <div style={{ fontSize: 12.5, color: C.dim }}>
          Managerial income statement for one period — every cell editable, spreadsheet-style. Expense lines are the same data as the Costs tab.
        </div>
        <Btn primary icon={Download} onClick={exportXlsx}>Export to Excel</Btn>
      </div>

      {/* assumptions strip */}
      <Card style={{ padding: 0, overflow: "hidden", marginBottom: 12 }}>
        <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)" }}>
          {[
            ["Cans sold (period)", <EditNum value={st.cans} onCommit={(n) => updSt({ cans: n })} size={17} bold />],
            ["Price / can", <EditNum value={price} onCommit={updPrice} prefix="$" decimals={2} size={17} bold />],
            ["Unit cost / can (likely)", <span style={{ fontSize: 17, fontWeight: 600 }}>${perCan.toFixed(2)}</span>],
          ].map(([l, node], i) => (
            <div key={l} style={{ padding: "12px 16px", borderLeft: i > 0 ? `1px solid ${C.borderSoft}` : "none" }}>
              <Micro>{l}</Micro>
              <div style={{ marginTop: 6, fontFamily: "'Geist', sans-serif" }}>{node}</div>
            </div>
          ))}
        </div>
      </Card>

      <div style={{ display: "flex" }}>
        <div style={{ width: 4, background: "#5B7A9D", borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
        <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
          <tbody>
            <tr><td style={sect} colSpan={2}>Income</td><td style={{ ...sect, width: 32 }}></td></tr>
            <tr><td style={TD}>Revenue ({st.cans.toLocaleString()} × ${price.toFixed(2)})</td><td style={{ ...num, fontWeight: 600 }}>{fmt$(revenue)}</td><td style={TD}></td></tr>
            <tr><td style={TD}>Cost of goods sold ({st.cans.toLocaleString()} × ${perCan.toFixed(2)})</td><td style={{ ...num, color: "#B4574B" }}>({fmt$(cogs)})</td><td style={TD}></td></tr>
            <tr><td style={totalRow}>Gross profit</td><td style={{ ...num, ...totalRow, color: gross >= 0 ? "#5E7D54" : "#B4574B" }}>{fmt$(gross)} <span style={{ fontSize: 11, color: C.faint, fontWeight: 500 }}>({revenue > 0 ? Math.round((gross / revenue) * 100) : 0}%)</span></td><td style={totalRow}></td></tr>

            <tr><td style={sect} colSpan={2}>Operating / setup expenses</td><td style={sect}></td></tr>
            {fin.setupCosts.map((r) => (
              <tr key={r.item}>
                <td style={TD}>
                  <InlineText value={r.item} onCommit={(v) => updateSetup(r.item, { item: v })} width={230} />
                  {r.spent && <span style={{ fontSize: 10, color: "#5E7D54", fontWeight: 700, marginLeft: 7 }}>PAID</span>}
                </td>
                <td style={{ ...num, color: "#B4574B" }}>(<EditNum value={r.cost} onCommit={(n) => updateSetup(r.item, { cost: n })} prefix="$" bold />)</td>
                <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}>
                  <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delSetup(r.item)} />
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={3} style={{ ...TD, padding: "5px 12px" }}><GroupAdd onAdd={addSetup} placeholder="Add expense line" /></td>
            </tr>
            <tr><td style={totalRow}>Total operating expenses</td><td style={{ ...num, ...totalRow, color: "#B4574B" }}>({fmt$(opex)})</td><td style={totalRow}></td></tr>

            <tr>
              <td style={{ ...totalRow, fontSize: 14.5 }}>Net income</td>
              <td style={{ ...num, ...totalRow, fontSize: 14.5, color: net >= 0 ? "#5E7D54" : "#B4574B" }}>{net < 0 ? `(${fmt$(-net)})` : fmt$(net)}</td>
              <td style={totalRow}></td>
            </tr>
          </tbody>
        </table>
      </div>

      <div style={{ fontSize: 11.5, color: C.faint, marginTop: 10 }}>
        Managerial view, not GAAP: setup costs are expensed in-period, no taxes/depreciation. The Excel export includes this P&L plus the Projections and Market sheets.
      </div>
    </div>
  );
}


function FundraisingTab({ data, setData, setModal, founders = [] }) {
  const STAGES = ["Identified", "Outreach", "Meeting Scheduled", "Follow-Up", "Due Diligence", "Invested", "Passed"];
  const INTEREST = ["High", "Medium", "Low"];
  const INTEREST_COLOR = { High: "#5E7D54", Medium: "#C8932E", Low: "#A89C86" };
  const funnel = STAGES.slice(0, 6).map((s) => ({ s, n: data.investors.filter((i) => i.status === s).length }));

  const patchInv = (id, patch) => setData((d) => ({ ...d, investors: d.investors.map((i) => i.id === id ? { ...i, ...patch } : i) }));
  const delInv = (id) => setData((d) => ({ ...d, investors: d.investors.filter((i) => i.id !== id) }));
  const addInv = (name) => setData((d) => ({
    ...d, investors: [...d.investors, { id: Date.now(), name, firm: "—", focus: "—", check: "—", last: "", interest: "Medium", status: "Identified" }],
  }));

  return (
    <div className="fade">
      <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: 14 }}>
        <Btn primary icon={Plus} onClick={() => setModal("investor")}>Add Investor</Btn>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "repeat(6,1fr)", gap: 9, marginBottom: 18 }}>
        {funnel.map((f) => (
          <Card key={f.s} style={{ padding: 14, textAlign: "center" }}>
            <div style={{ fontSize: 19, fontWeight: 600, fontFamily: "'Geist', sans-serif", color: f.n ? C.ink : C.faint }}>{f.n}</div>
            <div style={{ fontSize: 10, color: C.dim, marginTop: 4, fontWeight: 700, textTransform: "uppercase", letterSpacing: 0.5 }}>{f.s}</div>
          </Card>
        ))}
      </div>

      <div style={{ display: "flex" }}>
        <div style={{ width: 4, background: "#5E7D54", borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
        <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
          <thead><tr style={{ background: C.panel2 }}>
            <th style={TH}>Investor</th>
            <th style={{ ...TH, width: 130 }}>Firm</th>
            <th style={{ ...TH, width: 95 }}>Check</th>
            <th style={{ ...TH, width: 120 }}>Focus</th>
            <th style={{ ...TH, textAlign: "center", width: 84 }}>Interest</th>
            <th style={{ ...TH, textAlign: "center", width: 140 }}>Status</th>
            <th style={{ ...TH, width: 132 }}>Last contact</th>
            <th style={{ ...TH, width: 100 }}>Leads</th>
            <th style={{ width: 32 }}></th>
          </tr></thead>
          <tbody>
            {data.investors.map((i) => (
              <tr key={i.id}>
                <td style={{ ...TD, fontWeight: 500 }}><InlineText value={i.name} onCommit={(v) => patchInv(i.id, { name: v })} width={150} /></td>
                <td style={TD}><InlineText value={i.firm || "—"} onCommit={(v) => patchInv(i.id, { firm: v })} width={110} style={{ fontSize: 12.5 }} /></td>
                <td style={TD}><InlineText value={i.check || "—"} onCommit={(v) => patchInv(i.id, { check: v })} width={75} style={{ fontSize: 12.5 }} /></td>
                <td style={TD}><InlineText value={i.focus || "—"} onCommit={(v) => patchInv(i.id, { focus: v })} width={100} style={{ fontSize: 12.5 }} /></td>
                <ColorCell label={i.interest} color={INTEREST_COLOR[i.interest]} width={84}
                  onClick={() => patchInv(i.id, { interest: INTEREST[(INTEREST.indexOf(i.interest) + 1) % 3] })} />
                <ColorCell label={i.status} color={STAGE_COLORS[i.status]} width={140}
                  onClick={() => patchInv(i.id, { status: STAGES[(STAGES.indexOf(i.status) + 1) % STAGES.length] })} />
                <td style={TD}>
                  <input type="date" value={i.last || ""} onChange={(e) => patchInv(i.id, { last: e.target.value })}
                    style={{ border: "none", background: "transparent", fontFamily: "inherit", fontSize: 12, color: C.dim, outline: "none", cursor: "pointer", width: 112 }} />
                </td>
                <td style={TD}><Spearheads value={i.spear || []} onChange={(v) => patchInv(i.id, { spear: v })} founders={founders} size={16} /></td>
                <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}>
                  <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delInv(i.id)} />
                </td>
              </tr>
            ))}
            <tr>
              <td colSpan={9} style={{ ...TD, padding: "5px 12px" }}>
                <GroupAdd onAdd={addInv} placeholder="Add investor" />
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  STRATEGY — Go-to-market · Brand                                    */
/* =================================================================== */
function Strategy({ data, setData, setModal, me, setMe, founders = [], addFounder }) {
  const [tab, setTab] = useState("gtm");
  const g = data.gtm || {};
  const fin = data.financials;
  const upd = (patch) => setData((d) => ({ ...d, gtm: { ...(d.gtm || {}), ...patch } }));
  const updBrand = (patch) => setData((d) => ({ ...d, gtm: { ...(d.gtm || {}), brand: { ...((d.gtm || {}).brand || {}), ...patch } } }));
  const brand = g.brand || {};

  /* editable chip list — pass get/set so it works for gtm root or brand */
  const ChipList = ({ items = [], onChange, color }) => {
    const [draft, setDraft] = useState("");
    const add = () => { if (!draft.trim()) return; onChange([...items, draft.trim()]); setDraft(""); };
    return (
      <div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 8 }}>
          {items.map((item, i) => (
            <span key={i} style={{
              display: "inline-flex", alignItems: "center", gap: 6, fontSize: 12, fontWeight: 500,
              padding: "3px 9px", borderRadius: 6, color: C.ink, background: C.panel2, border: `1px solid ${C.borderSoft}`,
            }}>
              <span style={{ width: 6, height: 6, borderRadius: 99, background: color }} />
              {item}
              <X size={11} style={{ cursor: "pointer", opacity: 0.5 }} onClick={() => onChange(items.filter((_, j) => j !== i))} />
            </span>
          ))}
        </div>
        <div style={{ display: "flex", gap: 6 }}>
          <input value={draft} placeholder="Add…" onChange={(e) => setDraft(e.target.value)} onKeyDown={(e) => e.key === "Enter" && add()}
            style={{ ...inputStyle, flex: 1, width: "auto", padding: "4px 9px", fontSize: 12.5 }} />
          <button onClick={add} style={{ width: 28, borderRadius: 6, border: `1px solid ${C.border}`, background: C.panel, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}><Plus size={13} color={C.dim} /></button>
        </div>
      </div>
    );
  };

  /* editable paragraph */
  const EditPara = ({ value, onCommit, italic }) => {
    const [editing, setEditing] = useState(false);
    const [v, setV] = useState(value || "");
    if (editing) return (
      <textarea autoFocus rows={3} value={v} onChange={(e) => setV(e.target.value)}
        onBlur={() => { setEditing(false); onCommit(v); }}
        style={{ ...inputStyle, resize: "vertical", fontSize: 13 }} />
    );
    return <p onClick={() => { setV(value || ""); setEditing(true); }} title="Click to edit"
      style={{ fontSize: 13.5, lineHeight: 1.6, color: C.ink, margin: 0, cursor: "text", fontStyle: italic ? "italic" : "normal" }}>{value}</p>;
  };

  const ST = ({ children }) => <span style={{ fontWeight: 600, fontSize: 14, fontFamily: "'Geist', sans-serif" }}>{children}</span>;
  const tabs = [
    { id: "gtm", label: "Go-to-market" }, { id: "brand", label: "Brand" },
    { id: "creative", label: "Creative" }, { id: "discussion", label: "Discussion" }, { id: "meetings", label: "Meetings" },
  ];

  return (
    <div className="fade">
      <PageHead title="Strategy" sub="Go-to-market & brand — click any text to edit" />

      <div style={{ display: "flex", gap: 2, borderBottom: `1px solid ${C.borderSoft}`, marginBottom: 18 }}>
        {tabs.map((t) => (
          <button key={t.id} onClick={() => setTab(t.id)} style={{
            padding: "7px 13px", border: "none", cursor: "pointer", background: "transparent",
            color: tab === t.id ? C.ink : C.dim, fontFamily: "inherit",
            fontSize: 13, fontWeight: tab === t.id ? 600 : 500,
            borderBottom: tab === t.id ? `2px solid ${C.accent}` : "2px solid transparent",
            marginBottom: -1,
          }}>{t.label}</button>
        ))}
      </div>

      {tab === "gtm" && (
        <div className="fade" style={{ display: "flex", flexDirection: "column", gap: 12 }}>
          {/* Positioning */}
          <Card>
            <Micro style={{ color: C.accent }}>Elevator pitch</Micro>
            <div style={{ marginTop: 8, fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontSize: 15.5, lineHeight: 1.6 }}>
              <EditPara value={g.valueProp} onCommit={(v) => upd({ valueProp: v })} />
            </div>
            <div style={{ display: "flex", gap: 28, marginTop: 14, flexWrap: "wrap" }}>
              <div><Micro>Category</Micro><div style={{ fontSize: 13, fontWeight: 500, marginTop: 3 }}><InlineText value={g.category} onCommit={(v) => upd({ category: v })} width={280} /></div></div>
              <div><Micro>Main benefit</Micro><div style={{ fontSize: 13, fontWeight: 500, marginTop: 3 }}><InlineText value={g.mainBenefit} onCommit={(v) => upd({ mainBenefit: v })} width={240} /></div></div>
            </div>
          </Card>

          <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, alignItems: "start" }}>
            {/* Customer */}
            <Card>
              <ST>Customer</ST>
              <div style={{ marginTop: 10 }}><EditPara value={g.targetCustomer} onCommit={(v) => upd({ targetCustomer: v })} /></div>
              <div style={{ display: "flex", gap: 24, marginTop: 12 }}>
                <div><Micro>Age</Micro><div style={{ fontSize: 14, fontWeight: 600, marginTop: 2 }}><InlineText value={g.ageRange} onCommit={(v) => upd({ ageRange: v })} width={70} /></div></div>
                <div><Micro>Pricing</Micro><div style={{ fontSize: 14, fontWeight: 600, marginTop: 2 }}>${fin?.pricing?.can?.toFixed(2)} can · ${fin?.pricing?.twelvePack} 12-pack</div></div>
              </div>
              <div style={{ marginTop: 14 }}>
                <Micro style={{ display: "block", marginBottom: 7 }}>Occasions</Micro>
                <ChipList items={g.occasions} onChange={(v) => upd({ occasions: v })} color="#5B7A9D" />
              </div>
              <div style={{ marginTop: 14 }}>
                <Micro style={{ display: "block", marginBottom: 7 }}>Channels</Micro>
                <ChipList items={g.channels} onChange={(v) => upd({ channels: v })} color="#5E7D54" />
              </div>
            </Card>

            {/* Competition */}
            <Card>
              <ST>Competition</ST>
              <div style={{ marginTop: 12 }}>
                <Micro style={{ display: "block", marginBottom: 7 }}>Competitors</Micro>
                <ChipList items={g.competitors} onChange={(v) => upd({ competitors: v })} color="#B4574B" />
              </div>
              <div style={{ marginTop: 14 }}>
                <Micro style={{ display: "block", marginBottom: 7 }}>Their gaps — our openings</Micro>
                <ChipList items={g.competitorGaps} onChange={(v) => upd({ competitorGaps: v })} color="#C8932E" />
              </div>
              <div style={{ marginTop: 14 }}>
                <Micro style={{ display: "block", marginBottom: 7 }}>Customers switch from</Micro>
                <ChipList items={g.switchingFrom} onChange={(v) => upd({ switchingFrom: v })} color="#C26A35" />
              </div>
            </Card>
          </div>
        </div>
      )}

      {tab === "brand" && (
        <div className="fade" style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12, alignItems: "start" }}>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Card>
              <ST>Core Identity</ST>
              <div style={{ marginTop: 10, fontFamily: "'STIX Two Text', serif", fontStyle: "italic", fontSize: 15, lineHeight: 1.6 }}>
                <EditPara value={brand.coreIdentity} onCommit={(v) => updBrand({ coreIdentity: v })} />
              </div>
            </Card>
            <Card>
              <ST>What the Name Means</ST>
              <div style={{ marginTop: 10 }}><EditPara value={brand.nameMeaning} onCommit={(v) => updBrand({ nameMeaning: v })} /></div>
            </Card>
            <Card>
              <ST>Brand in Four Words</ST>
              <div style={{ marginTop: 12 }}><ChipList items={g.brandWords} onChange={(v) => upd({ brandWords: v })} color={C.accent} /></div>
            </Card>
          </div>
          <div style={{ display: "flex", flexDirection: "column", gap: 12 }}>
            <Card>
              <ST>Look & Feel</ST>
              <div style={{ marginTop: 12 }}><ChipList items={brand.feel} onChange={(v) => updBrand({ feel: v })} color="#8A6FA8" /></div>
            </Card>
            <Card>
              <ST>Colors to Use</ST>
              <div style={{ marginTop: 12 }}><ChipList items={brand.colorsUse} onChange={(v) => updBrand({ colorsUse: v })} color="#5E7D54" /></div>
            </Card>
            <Card>
              <ST>Colors to Never Use</ST>
              <div style={{ marginTop: 12 }}><ChipList items={brand.colorsNever} onChange={(v) => updBrand({ colorsNever: v })} color="#B4574B" /></div>
            </Card>
          </div>
        </div>
      )}

      {tab === "creative" && <Creative {...{ data, setData, founders }} />}
      {tab === "discussion" && <Discussion {...{ me, setMe, founders, addFounder }} />}
      {tab === "meetings" && <Meetings {...{ data, setData, setModal, founders }} />}
    </div>
  );
}

/* ------------------------------------------------------------------ */
/* ------------------------------------------------------------------ */
/*  Shared Monday-style table pieces                                    */
/* ------------------------------------------------------------------ */
const TH = { padding: "7px 12px", fontSize: 10, fontWeight: 600, letterSpacing: 1, textTransform: "uppercase", color: "#A89C86", textAlign: "left" };
const TD = { padding: "7px 12px", fontSize: 13, borderTop: "1px solid #F0EAE0", background: "#FFFFFF" };

function ColorCell({ label, color, onClick, width }) {
  return (
    <td onClick={onClick} style={{
      background: color, color: "#fff", fontSize: 12, fontWeight: 600, textAlign: "center",
      padding: "7px 8px", cursor: onClick ? "pointer" : "default", borderTop: "1px solid #FFFFFF",
      width, userSelect: "none", whiteSpace: "nowrap",
    }} title={onClick ? "Click to change" : undefined}>{label}</td>
  );
}

/* ------------------------------------------------------------------ */
/*  LEADS — founders driving this. Stacked colored initials + picker.  */
/*  Reusable on any record (tasks, mockups, meetings, regulatory…).     */
/*  Data shape is just an array of names → maps to user IDs later.      */
/* ------------------------------------------------------------------ */
const SPEAR_COLORS = ["#C26A35", "#5B7A9D", "#5E7D54", "#8A6FA8", "#C8932E", "#B4574B"];
const spearColor = (name, founders = []) => {
  const i = founders.indexOf(name);
  if (i >= 0) return SPEAR_COLORS[i % SPEAR_COLORS.length];
  let h = 0; for (const ch of name) h = (h * 31 + ch.charCodeAt(0)) % 997;
  return SPEAR_COLORS[h % SPEAR_COLORS.length];
};
const initials = (n) => n.split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase();

function Spearheads({ value = [], onChange, founders = [], size = 18 }) {
  const [open, setOpen] = useState(false);
  const toggle = (f) => onChange(value.includes(f) ? value.filter((x) => x !== f) : [...value, f]);
  return (
    <span style={{ position: "relative", display: "inline-flex", alignItems: "center", gap: 5 }}>
      <span style={{ display: "inline-flex" }}>
        {value.map((f, i) => (
          <span key={f} title={f} style={{
            width: size, height: size, borderRadius: 99, background: spearColor(f, founders),
            color: "#fff", fontSize: size * 0.44, fontWeight: 700,
            display: "inline-flex", alignItems: "center", justifyContent: "center",
            border: "1.5px solid #FFFFFF", marginLeft: i === 0 ? 0 : -6, zIndex: i,
          }}>{initials(f)}</span>
        ))}
      </span>
      <button onClick={(e) => { e.stopPropagation(); setOpen(!open); }} title="Founders leading this" style={{
        width: size, height: size, borderRadius: 99, border: `1px dashed ${C.faint}`,
        background: "transparent", color: C.faint, cursor: "pointer", flexShrink: 0,
        display: "inline-flex", alignItems: "center", justifyContent: "center", padding: 0,
      }}><Plus size={size * 0.55} /></button>
      {open && (
        <>
          <span onClick={() => setOpen(false)} style={{ position: "fixed", inset: 0, zIndex: 70 }} />
          <div style={{
            position: "absolute", top: size + 6, left: 0, width: 168, background: C.panel,
            border: `1px solid ${C.border}`, borderRadius: 8, zIndex: 71, padding: 6,
            boxShadow: "0 8px 24px rgba(60,45,25,.14)",
          }}>
            <Micro style={{ display: "block", padding: "2px 6px 6px" }}>Leads</Micro>
            {founders.map((f) => {
              const on = value.includes(f);
              return (
                <button key={f} onClick={() => toggle(f)} style={{
                  display: "flex", alignItems: "center", gap: 8, width: "100%", padding: "5px 7px",
                  border: "none", background: on ? C.panel2 : "transparent", borderRadius: 6,
                  cursor: "pointer", fontFamily: "inherit", fontSize: 12.5, color: C.ink, textAlign: "left",
                }}>
                  <span style={{
                    width: 17, height: 17, borderRadius: 99, background: spearColor(f, founders),
                    color: "#fff", fontSize: 8, fontWeight: 700,
                    display: "inline-flex", alignItems: "center", justifyContent: "center",
                  }}>{initials(f)}</span>
                  <span style={{ flex: 1 }}>{f}</span>
                  {on && <CheckCircle2 size={13} color="#5E7D54" />}
                </button>
              );
            })}
          </div>
        </>
      )}
    </span>
  );
}

/* add-a-founder row in the identity popover */
function IdentityAdd({ onAdd }) {
  const [v, setV] = useState("");
  const go = () => { if (!v.trim()) return; onAdd(v.trim()); setV(""); };
  return (
    <div style={{ display: "flex", gap: 5, marginTop: 6, padding: "0 2px" }}>
      <input value={v} placeholder="Add founder…" onChange={(e) => setV(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && go()}
        style={{ ...inputStyle, flex: 1, width: "auto", padding: "4px 8px", fontSize: 12 }} />
      <button onClick={go} style={{ width: 26, borderRadius: 6, border: `1px solid ${C.border}`, background: C.panel, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <Plus size={12} color={C.dim} />
      </button>
    </div>
  );
}

/* small inline add-row used by scheduled meeting agendas */
function AgendaAdd({ onAdd }) {
  const [v, setV] = useState("");
  const go = () => { if (!v.trim()) return; onAdd(v.trim()); setV(""); };
  return (
    <div style={{ display: "flex", gap: 6, marginTop: 8 }}>
      <input value={v} placeholder="Add agenda item…" onChange={(e) => setV(e.target.value)}
        onKeyDown={(e) => e.key === "Enter" && go()}
        style={{ ...inputStyle, flex: 1, width: "auto", padding: "4px 9px", fontSize: 12.5 }} />
      <button onClick={go} style={{ width: 28, borderRadius: 6, border: `1px solid ${C.border}`, background: C.panel, cursor: "pointer", display: "flex", alignItems: "center", justifyContent: "center" }}>
        <Plus size={13} color={C.dim} />
      </button>
    </div>
  );
}

/* =================================================================== */
/*  MEETINGS — Schedule · AI Notes · Review                            */
/* =================================================================== */
function Meetings({ data, setData, setModal, founders = [] }) {
  const scheduled = data.scheduled || [];
  const [transcript, setTranscript] = useState("");
  const [aiDate, setAiDate] = useState("");
  const [aiParts, setAiParts] = useState("");
  const [generating, setGenerating] = useState(false);
  const [aiError, setAiError] = useState("");

  const updScheduled = (id, patch) => setData((d) => ({ ...d, scheduled: (d.scheduled || []).map((m) => m.id === id ? { ...m, ...patch } : m) }));
  const delScheduled = (id) => setData((d) => ({ ...d, scheduled: (d.scheduled || []).filter((m) => m.id !== id) }));
  const delMeeting = (id) => setData((d) => ({ ...d, meetings: d.meetings.filter((m) => m.id !== id) }));

  const onFile = (e) => {
    const f = e.target.files?.[0];
    if (!f) return;
    if (!/\.(txt|md)$/i.test(f.name)) { setAiError("Upload a .txt or .md file — for Word docs, paste the text instead."); return; }
    const r = new FileReader();
    r.onload = () => { setTranscript(String(r.result || "")); setAiError(""); };
    r.readAsText(f);
  };

  const generateNotes = async () => {
    if (!transcript.trim() || generating) return;
    setGenerating(true); setAiError("");
    const saveRecord = (rec) => setData((d) => ({
      ...d,
      meetings: [rec, ...d.meetings],
      tasks: [
        ...rec.actions.map((a, i) => ({
          id: Date.now() + i + 1, title: a, owner: "—", due: "2026-06-12",
          priority: "Medium", status: "Todo", link: "General", auto: true,
        })),
        ...d.tasks,
      ],
    }));
    try {
      const response = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(window.getAuthToken ? { Authorization: `Bearer ${await window.getAuthToken()}` } : {}) },
        body: JSON.stringify({
          model: "claude-sonnet-4-20250514",
          max_tokens: 1000,
          messages: [{
            role: "user",
            content: `You are the meeting-notes assistant for Eudai, a nootropic energy drink startup (Orange-Peach Cream, 120mg green tea caffeine, Cognizin hero ingredient, pre-revenue). From the meeting transcript below, respond with ONLY valid JSON and nothing else — no markdown fences, no preamble: {"summary": "2-4 sentence summary", "decisions": ["decision", ...], "action_items": ["concrete next step", ...]}\n\nTranscript:\n${transcript}`,
          }],
        }),
      });
      const d = await response.json();
      const text = (d.content || []).filter((i) => i.type === "text").map((i) => i.text).join("\n");
      const parsed = JSON.parse(text.replace(/```json|```/g, "").trim());
      saveRecord({
        id: Date.now(), date: aiDate || "—", parts: aiParts || "Founders",
        notes: parsed.summary || "", decisions: parsed.decisions || [], actions: parsed.action_items || [],
        ai: true,
      });
      setTranscript(""); setAiDate(""); setAiParts("");
    } catch (e) {
      setAiError("AI notes failed — saved the raw transcript as a meeting record instead.");
      saveRecord({ id: Date.now(), date: aiDate || "—", parts: aiParts || "Founders", notes: transcript.trim(), decisions: [], actions: [] });
      setTranscript("");
    }
    setGenerating(false);
  };

  return (
    <div className="fade">
      <PageHead title="Meetings" sub="Schedule, generate AI notes & review past meetings"
        action={<Btn primary icon={Plus} onClick={() => setModal("schedule")}>Schedule Meeting</Btn>} />

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 14, alignItems: "start", marginBottom: 14 }}>
        {/* ---------- Upcoming / scheduled ---------- */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px 9px", display: "flex", alignItems: "center", gap: 8 }}>
            <CalendarDays size={16} color={C.accent} />
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Upcoming</span>
          </div>
          {scheduled.length === 0 && (
            <div style={{ padding: "24px 18px", borderTop: `1px solid ${C.borderSoft}`, fontSize: 13, color: C.faint, textAlign: "center" }}>
              Nothing scheduled — use Schedule Meeting above.
            </div>
          )}
          {scheduled.map((m) => (
            <div key={m.id} style={{ padding: "11px 16px", borderTop: `1px solid ${C.borderSoft}` }}>
              <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
                <div style={{ fontWeight: 600, fontSize: 14.5, fontFamily: "'Geist', sans-serif" }}>{m.title}</div>
                <Trash2 size={14} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delScheduled(m.id)} />
              </div>
              <div style={{ display: "flex", gap: 8, alignItems: "center", margin: "8px 0" }}>
                <input type="date" value={m.date} onChange={(e) => updScheduled(m.id, { date: e.target.value })}
                  style={{ ...inputStyle, width: 140, padding: "5px 8px", fontSize: 12.5 }} />
                <input type="time" value={m.time} onChange={(e) => updScheduled(m.id, { time: e.target.value })}
                  style={{ ...inputStyle, width: 104, padding: "5px 8px", fontSize: 12.5 }} />
                <span style={{ fontSize: 12, color: C.dim }}>{m.parts}</span>
                <Spearheads value={m.spear || []} onChange={(v) => updScheduled(m.id, { spear: v })} founders={founders} size={16} />
                {!m.date && <Badge label="Date TBD" color="#C8932E" />}
              </div>
              <div style={{ marginTop: 6 }}>
                <Micro>Agenda</Micro>
                {(m.agenda || []).map((a, i) => (
                  <div key={i} style={{ fontSize: 13, fontWeight: 500, display: "flex", alignItems: "flex-start", gap: 7, marginTop: 5 }}>
                    <ChevronRight size={13} color={C.accent} style={{ flexShrink: 0, marginTop: 3 }} />
                    <span style={{ flex: 1 }}>
                      <InlineText value={a} onCommit={(v) => updScheduled(m.id, { agenda: m.agenda.map((x, j) => j === i ? v : x) })} width={230} />
                    </span>
                    <X size={12} color={C.faint} style={{ cursor: "pointer", marginTop: 3 }}
                      onClick={() => updScheduled(m.id, { agenda: m.agenda.filter((_, j) => j !== i) })} />
                  </div>
                ))}
                <AgendaAdd onAdd={(v) => updScheduled(m.id, { agenda: [...(m.agenda || []), v] })} />
              </div>
              <div style={{ marginTop: 10 }}>
                <Btn small onClick={() => { setAiDate(m.date || ""); setAiParts(m.parts || ""); delScheduled(m.id); }}>
                  Met → write notes
                </Btn>
              </div>
            </div>
          ))}
        </Card>

        {/* ---------- AI notes ---------- */}
        <Card style={{ padding: 0, overflow: "hidden", background: C.panel }}>
          <div style={{ padding: "12px 16px 9px", display: "flex", alignItems: "center", gap: 8 }}>
            <Sparkles size={16} color={C.accent} />
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>AI Meeting Notes</span>
          </div>
          <div style={{ padding: "0 16px 12px" }}>
            <div style={{ fontSize: 12, color: C.dim, marginBottom: 12 }}>
              Paste a transcript or rough notes — or upload a .txt/.md file — and AI turns it into a summary, decisions, and action items. Action items become tasks automatically.
            </div>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 8, marginBottom: 8 }}>
              <input type="date" value={aiDate} onChange={(e) => setAiDate(e.target.value)} style={{ ...inputStyle }} />
              <input value={aiParts} placeholder="Participants" onChange={(e) => setAiParts(e.target.value)} style={{ ...inputStyle }} />
            </div>
            <textarea rows={7} value={transcript} placeholder="Paste meeting transcript or notes here…"
              onChange={(e) => setTranscript(e.target.value)}
              style={{ ...inputStyle, resize: "vertical", marginBottom: 8 }} />
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 }}>
              <label style={{ fontSize: 12, color: C.dim, cursor: "pointer", fontWeight: 600, display: "flex", alignItems: "center", gap: 6 }}>
                <FileText size={14} /> Upload .txt / .md
                <input type="file" accept=".txt,.md" onChange={onFile} style={{ display: "none" }} />
              </label>
              <Btn primary icon={Sparkles} onClick={generateNotes}>
                {generating ? "Generating…" : "Generate Notes"}
              </Btn>
            </div>
            {aiError && <div style={{ fontSize: 12, color: "#B4574B", marginTop: 10, fontWeight: 600 }}>{aiError}</div>}
          </div>
        </Card>
      </div>

      {/* ---------- Past meetings ---------- */}
      <div style={{ display: "flex", alignItems: "center", gap: 8, margin: "20px 0 12px" }}>
        <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Past Meetings</span>
        <span style={{ fontSize: 12, color: C.faint }}>{data.meetings.length} record{data.meetings.length === 1 ? "" : "s"}</span>
      </div>
      <div style={{ display: "flex", flexDirection: "column", gap: 13 }}>
        {data.meetings.map((m) => (
          <Card key={m.id} style={{ padding: 16 }}>
            <div style={{ display: "flex", alignItems: "center", gap: 9, marginBottom: 12 }}>
              <CalendarDays size={16} color={C.accent} />
              <span style={{ fontWeight: 600, fontSize: 15, fontFamily: "'Geist', sans-serif" }}>{m.date}</span>
              <span style={{ fontSize: 12.5, color: C.dim }}>· {m.parts}</span>
              {m.ai && <Badge label="AI notes" color={C.accent} />}
              <Trash2 size={14} color={C.faint} style={{ cursor: "pointer", marginLeft: "auto" }} onClick={() => delMeeting(m.id)} />
            </div>
            <p style={{ fontSize: 13.5, color: C.dim, lineHeight: 1.65, margin: "0 0 16px" }}>{m.notes}</p>
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 16 }}>
              <div>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: C.faint, textTransform: "uppercase", letterSpacing: 1, marginBottom: 9 }}>Decisions</div>
                {m.decisions.map((d, idx) => <div key={idx} style={{ fontSize: 13, marginBottom: 6, display: "flex", gap: 7, fontWeight: 500 }}><CheckCircle2 size={14} color="#5E7D54" style={{ flexShrink: 0, marginTop: 2 }} />{d}</div>)}
              </div>
              <div>
                <div style={{ fontSize: 10.5, fontWeight: 700, color: C.faint, textTransform: "uppercase", letterSpacing: 1, marginBottom: 9 }}>Action Items</div>
                {m.actions.map((a, idx) => <div key={idx} style={{ fontSize: 13, marginBottom: 6, display: "flex", gap: 7, fontWeight: 500 }}><ArrowRight size={14} color={C.accent} style={{ flexShrink: 0, marginTop: 2 }} />{a}</div>)}
              </div>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* =================================================================== */
/*  CREATIVE — Idea Board · Mockups · Marketing Ideas                  */
/* =================================================================== */
function Creative({ data, setData, founders = [] }) {
  const cr = data.creative || { ideas: [], mockups: [], marketing: [] };
  const [ideaText, setIdeaText] = useState("");
  const [ideaTag, setIdeaTag] = useState("General");
  const [mkText, setMkText] = useState("");
  const [mockName, setMockName] = useState("");
  const [mockOwner, setMockOwner] = useState("");

  const upd = (patch) => setData((d) => ({ ...d, creative: { ...(d.creative || { ideas: [], mockups: [], marketing: [] }), ...patch } }));

  const addIdea = () => { if (!ideaText.trim()) return; upd({ ideas: [{ id: Date.now(), text: ideaText.trim(), tag: ideaTag }, ...cr.ideas] }); setIdeaText(""); };
  const delIdea = (id) => upd({ ideas: cr.ideas.filter((i) => i.id !== id) });
  const addMk = () => { if (!mkText.trim()) return; upd({ marketing: [{ id: Date.now(), text: mkText.trim() }, ...cr.marketing] }); setMkText(""); };
  const delMk = (id) => upd({ marketing: cr.marketing.filter((m) => m.id !== id) });
  const addMock = () => { if (!mockName.trim()) return; upd({ mockups: [{ id: Date.now(), name: mockName.trim(), owner: mockOwner.trim() || "—", status: "Not Started", notes: "" }, ...cr.mockups] }); setMockName(""); setMockOwner(""); };
  const delMock = (id) => upd({ mockups: cr.mockups.filter((m) => m.id !== id) });
  const cycleMock = (id) => {
    const order = ["Not Started", "In Progress", "Review", "Complete"];
    upd({ mockups: cr.mockups.map((m) => m.id === id ? { ...m, status: order[(order.indexOf(m.status) + 1) % order.length] } : m) });
  };
  const mockColor = { "Not Started": "#A89C86", "In Progress": "#C8932E", "Review": "#5B7A9D", "Complete": "#5E7D54" };
  const tagColor = { General: "#A89C86", Product: "#5B7A9D", Marketing: "#C26A35", Brand: "#8A6FA8", Packaging: "#C8932E", Flavor: "#5E7D54" };

  return (
    <div className="fade">
      <PageHead title="Creative" sub="Ideas, mockups in flight & marketing direction" />

      <div style={{ display: "grid", gridTemplateColumns: "1.1fr 1fr 1fr", gap: 14, alignItems: "start" }}>
        {/* ---------- Idea Board ---------- */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px 9px", display: "flex", alignItems: "center", gap: 8 }}>
            <Lightbulb size={16} color={C.accent} />
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Idea Board</span>
          </div>
          <div style={{ padding: "0 14px 11px", display: "flex", gap: 7 }}>
            <input value={ideaText} placeholder="Drop an idea…"
              onChange={(e) => setIdeaText(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addIdea()}
              style={{ ...inputStyle, flex: 1, width: "auto" }} />
            <select value={ideaTag} onChange={(e) => setIdeaTag(e.target.value)} style={{ ...inputStyle, width: 102 }}>
              {Object.keys(tagColor).map((t) => <option key={t}>{t}</option>)}
            </select>
            <button onClick={addIdea} style={{
              width: 38, borderRadius: 10, border: "none", background: C.accent, cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            }}><Plus size={16} color="#FFF8F0" /></button>
          </div>
          <div>
            {cr.ideas.length === 0 && (
              <div style={{ padding: "22px 18px", borderTop: `1px solid ${C.borderSoft}`, fontSize: 13, color: C.faint, textAlign: "center" }}>
                Nothing here yet — log the first idea above, or use Log Idea on the dashboard.
              </div>
            )}
            {cr.ideas.map((i) => (
              <div key={i.id} className="idea-row" style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "9px 16px", borderTop: `1px solid ${C.borderSoft}` }}>
                <div style={{ flex: 1 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 500, lineHeight: 1.5 }}>{i.text}</div>
                  <div style={{ marginTop: 6 }}><Badge label={i.tag} color={tagColor[i.tag] || C.faint} /></div>
                </div>
                <Trash2 size={14} color={C.faint} style={{ cursor: "pointer", flexShrink: 0, marginTop: 3 }} onClick={() => delIdea(i.id)} />
              </div>
            ))}
          </div>
        </Card>

        {/* ---------- Mockups ---------- */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px 9px", display: "flex", alignItems: "center", gap: 8 }}>
            <Palette size={16} color={C.accent} />
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Mockups</span>
          </div>
          <div style={{ padding: "0 14px 11px", display: "flex", gap: 7 }}>
            <input value={mockName} placeholder="Mockup name…"
              onChange={(e) => setMockName(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addMock()}
              style={{ ...inputStyle, flex: 1, width: "auto" }} />
            <input value={mockOwner} placeholder="Owner"
              onChange={(e) => setMockOwner(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addMock()}
              style={{ ...inputStyle, width: 76 }} />
            <button onClick={addMock} style={{
              width: 38, borderRadius: 10, border: "none", background: C.accent, cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            }}><Plus size={16} color="#FFF8F0" /></button>
          </div>
          <div>
            {cr.mockups.map((m) => (
              <div key={m.id} style={{ display: "flex", alignItems: "center", gap: 10, padding: "9px 16px", borderTop: `1px solid ${C.borderSoft}` }}>
                <div style={{ flex: 1, minWidth: 0 }}>
                  <div style={{ fontSize: 13.5, fontWeight: 600 }}>{m.name}</div>
                  <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}>{m.owner}{m.notes ? ` · ${m.notes}` : ""}</div>
                </div>
                <Spearheads value={m.spear || []} onChange={(v) => upd({ mockups: cr.mockups.map((x) => x.id === m.id ? { ...x, spear: v } : x) })} founders={founders} size={16} />
                <span onClick={() => cycleMock(m.id)} style={{ cursor: "pointer" }} title="Click to advance status">
                  <Badge label={m.status} color={mockColor[m.status]} />
                </span>
                <Trash2 size={14} color={C.faint} style={{ cursor: "pointer", flexShrink: 0 }} onClick={() => delMock(m.id)} />
              </div>
            ))}
          </div>
        </Card>

        {/* ---------- Marketing Ideas ---------- */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "12px 16px 9px", display: "flex", alignItems: "center", gap: 8 }}>
            <Megaphone size={16} color={C.accent} />
            <span style={{ fontWeight: 600, fontSize: 16, fontFamily: "'Geist', sans-serif" }}>Marketing Ideas</span>
          </div>
          <div style={{ padding: "0 14px 11px", display: "flex", gap: 7 }}>
            <input value={mkText} placeholder="Add a marketing idea…"
              onChange={(e) => setMkText(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && addMk()}
              style={{ ...inputStyle, flex: 1, width: "auto" }} />
            <button onClick={addMk} style={{
              width: 38, borderRadius: 10, border: "none", background: C.accent, cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            }}><Plus size={16} color="#FFF8F0" /></button>
          </div>
          <div>
            {cr.marketing.map((m) => (
              <div key={m.id} style={{ display: "flex", alignItems: "flex-start", gap: 10, padding: "9px 16px", borderTop: `1px solid ${C.borderSoft}` }}>
                <Megaphone size={13} color={C.faint} style={{ flexShrink: 0, marginTop: 3 }} />
                <div style={{ flex: 1, fontSize: 13.5, fontWeight: 500, lineHeight: 1.5 }}>{m.text}</div>
                <Trash2 size={14} color={C.faint} style={{ cursor: "pointer", flexShrink: 0, marginTop: 3 }} onClick={() => delMk(m.id)} />
              </div>
            ))}
          </div>
        </Card>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  DISCUSSION — founders board, DMs & group chats                     */
/*  Lives on its own storage key, polled every 7s, merged by id so     */
/*  founders writing at the same time don't overwrite each other.      */
/* =================================================================== */
/* =================================================================== */
/*  SETTINGS — workspace, appearance, founders, danger zone            */
/* =================================================================== */
function SettingsPage({ settings, updSettings, nav, data, setData }) {
  const swatches = ["#C26A35", "#5E7D54", "#5B7A9D", "#8A6FA8", "#B4574B", "#2A251F"];
  const pages = nav.filter((n) => n.id !== "settings");

  const Row = ({ label, hint, children }) => (
    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 20, padding: "13px 0", borderBottom: `1px solid ${C.hairline}` }}>
      <div>
        <div style={{ fontSize: 13.5, fontWeight: 500 }}>{label}</div>
        {hint && <div style={{ fontSize: 11.5, color: C.faint, marginTop: 2 }}>{hint}</div>}
      </div>
      <div style={{ flexShrink: 0 }}>{children}</div>
    </div>
  );

  return (
    <div className="fade" style={{ maxWidth: 640 }}>
      <PageHead title="Settings" sub="Workspace preferences — changes apply for all founders and save automatically" />

      <Card style={{ padding: "4px 18px", marginBottom: 12 }}>
        <Row label="Company name" hint="Shown in the top bar breadcrumb">
          <InlineText value={settings.companyName} onCommit={(v) => updSettings({ companyName: v || "Eudai" })} width={140} style={{ fontSize: 13.5, fontWeight: 600 }} />
        </Row>
        <Row label="Accent color" hint="Used for active states, highlights & the brand period">
          <div style={{ display: "flex", gap: 7 }}>
            {swatches.map((c) => (
              <button key={c} onClick={() => updSettings({ accent: c })} style={{
                width: 22, height: 22, borderRadius: 99, background: c, cursor: "pointer",
                border: settings.accent === c ? `2px solid ${C.ink}` : "2px solid transparent",
                outline: settings.accent === c ? `1px solid ${C.panel}` : "none", outlineOffset: -3,
              }} title={c} />
            ))}
          </div>
        </Row>
        <Row label="Background" hint="Cream is the Eudai house canvas">
          <div style={{ display: "flex", gap: 6 }}>
            {[["cream", "Cream"], ["white", "White"]].map(([id, label]) => (
              <button key={id} onClick={() => updSettings({ background: id })} style={{
                padding: "5px 12px", borderRadius: 6, cursor: "pointer", fontFamily: "inherit",
                fontSize: 12, fontWeight: 500,
                border: `1px solid ${settings.background === id ? C.ink : C.border}`,
                background: settings.background === id ? C.ink : C.panel,
                color: settings.background === id ? "#FAF7F2" : C.ink,
              }}>{label}</button>
            ))}
          </div>
        </Row>
        <Row label="Default page" hint="Where the workspace opens">
          <select value={settings.defaultPage} onChange={(e) => updSettings({ defaultPage: e.target.value })}
            style={{ ...inputStyle, width: 160 }}>
            {pages.map((p) => <option key={p.id} value={p.id}>{p.label}</option>)}
          </select>
        </Row>
      </Card>

      <Card style={{ padding: "12px 18px 16px", marginBottom: 12 }}>
        <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 4 }}>Founders</div>
        <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 12 }}>The sign-in roster. Each founder picks their name from the top bar to personalize the workspace.</div>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginBottom: 10 }}>
          {settings.founders.map((f) => (
            <span key={f} style={{
              display: "inline-flex", alignItems: "center", gap: 7, fontSize: 12.5, fontWeight: 500,
              padding: "4px 10px 4px 5px", borderRadius: 6, background: C.panel2, border: `1px solid ${C.borderSoft}`,
            }}>
              <span style={{ width: 18, height: 18, borderRadius: 99, background: C.accent, color: "#fff", fontSize: 9, fontWeight: 700, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
                {f.split(" ").map((x) => x[0]).join("").slice(0, 2).toUpperCase()}
              </span>
              {f}
              <X size={11} style={{ cursor: "pointer", opacity: 0.5 }}
                onClick={() => updSettings({ founders: settings.founders.filter((x) => x !== f) })} />
            </span>
          ))}
        </div>
        <IdentityAdd onAdd={(n) => { if (!settings.founders.includes(n)) updSettings({ founders: [...settings.founders, n] }); }} />
      </Card>

      <Card style={{ padding: "12px 18px 16px" }}>
        <div style={{ fontSize: 13.5, fontWeight: 600, marginBottom: 4 }}>Backup & Migration</div>
        <div style={{ fontSize: 11.5, color: C.faint, marginBottom: 12 }}>
          Export downloads the entire workspace (tasks, financials, vendors, everything) as a JSON file — your undo button and your bridge to the production app. Import replaces the current workspace with a backup file.
        </div>
        <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
          <Btn small icon={Download} onClick={() => {
            const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `eudai-backup-${new Date().toISOString().slice(0, 10)}.json`;
            a.click();
            URL.revokeObjectURL(url);
          }}>Export workspace</Btn>
          <label style={{ display: "inline-flex" }}>
            <span style={{
              display: "inline-flex", alignItems: "center", gap: 6, fontWeight: 500, fontSize: 12,
              padding: "5px 10px", borderRadius: 6, border: `1px solid ${C.border}`, background: C.panel,
              color: C.ink, cursor: "pointer", fontFamily: "'Geist', sans-serif",
            }}>Import backup…</span>
            <input type="file" accept=".json,application/json" style={{ display: "none" }}
              onChange={(e) => {
                const f = e.target.files?.[0];
                if (!f) return;
                const r = new FileReader();
                r.onload = () => {
                  try {
                    const parsed = JSON.parse(String(r.result));
                    if (!parsed.tasks || !parsed.financials) { alert("That file doesn't look like an Eudai workspace backup."); return; }
                    if (confirm("Replace the current workspace with this backup? This affects all founders.")) setData(parsed);
                  } catch { alert("Couldn't read that file — is it a valid backup JSON?"); }
                };
                r.readAsText(f);
                e.target.value = "";
              }} />
          </label>
        </div>
      </Card>
    </div>
  );
}

/* =================================================================== */
/*  SETTINGS END                                                       */
/* =================================================================== */

const CHAT_KEY = "ccos:eudai:chat:v1";
const emptyChat = () => ({ members: [], threads: [{ id: "board", name: "Founders Board", members: [], messages: [] }] });

function mergeChat(a, b) {
  const out = { members: [...new Set([...(a.members || []), ...(b.members || [])])], threads: [] };
  const ids = [...new Set([...(a.threads || []).map(t => t.id), ...(b.threads || []).map(t => t.id)])];
  ids.forEach((id) => {
    const ta = (a.threads || []).find(t => t.id === id);
    const tb = (b.threads || []).find(t => t.id === id);
    const base = ta || tb;
    const msgs = {};
    [...(ta?.messages || []), ...(tb?.messages || [])].forEach((m) => { msgs[m.id] = m; });
    out.threads.push({ ...base, messages: Object.values(msgs).sort((x, y) => x.ts - y.ts) });
  });
  return out;
}

function Discussion({ me, setMe, founders = [], addFounder }) {
  const [chat, setChat] = useState(emptyChat());
  const [meDraft, setMeDraft] = useState("");
  const [active, setActive] = useState("board");
  const [msg, setMsg] = useState("");
  const [newChatOpen, setNewChatOpen] = useState(false);
  const [ncName, setNcName] = useState("");
  const [ncMembers, setNcMembers] = useState([]);
  const [syncing, setSyncing] = useState(false);
  const chatRef = useRef(chat);
  chatRef.current = chat;

  const load = async () => {
    try {
      const res = await window.storage.get(CHAT_KEY, true);
      if (res?.value) setChat((cur) => mergeChat(cur, JSON.parse(res.value)));
    } catch { /* first run — nothing stored yet */ }
  };

  const save = async (next) => {
    setChat(next);
    setSyncing(true);
    try {
      let remote = emptyChat();
      try { const r = await window.storage.get(CHAT_KEY, true); if (r?.value) remote = JSON.parse(r.value); } catch {}
      const merged = mergeChat(next, remote);
      await window.storage.set(CHAT_KEY, JSON.stringify(merged), true);
      setChat(merged);
    } catch {}
    setSyncing(false);
  };

  useEffect(() => {
    load();
    const t = setInterval(load, 7000);
    return () => clearInterval(t);
  }, []);

  const joinAs = () => {
    const name = meDraft.trim();
    if (!name) return;
    addFounder(name);
    setMe(name);
  };

  const visibleThreads = chat.threads.filter((t) => t.members.length === 0 || !me || t.members.includes(me));
  const thread = chat.threads.find((t) => t.id === active) || chat.threads[0];

  const sendMsg = () => {
    if (!msg.trim() || !me) return;
    const m = { id: Date.now() + "_" + Math.random().toString(36).slice(2, 7), author: me, text: msg.trim(), ts: Date.now() };
    save({ ...chat, threads: chat.threads.map((t) => t.id === thread.id ? { ...t, messages: [...t.messages, m] } : t) });
    setMsg("");
  };

  const createChat = () => {
    const members = [...new Set([me, ...ncMembers])].filter(Boolean);
    if (members.length < 2) return;
    const name = ncName.trim() || members.join(", ");
    const id = "chat_" + Date.now();
    save({ ...chat, threads: [...chat.threads, { id, name, members, messages: [] }] });
    setActive(id); setNewChatOpen(false); setNcName(""); setNcMembers([]);
  };

  const fmtTime = (ts) => new Date(ts).toLocaleString("en-US", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });

  /* identity gate */
  if (!me) return (
    <div className="fade">
      <PageHead title="Discussion" sub="Founders board, direct messages & group chats" />
      <Card style={{ maxWidth: 440, padding: 28 }}>
        <div style={{ fontFamily: "'Geist', sans-serif", fontSize: 19, fontWeight: 600 }}>Who's posting?</div>
        <div style={{ fontSize: 13, color: C.dim, margin: "6px 0 16px" }}>Sign in (top right) or pick your name here — it personalizes the whole workspace and attributes your messages.</div>
        {founders.length > 0 && (
          <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginBottom: 14 }}>
            {founders.map((m) => (
              <button key={m} onClick={() => { setMe(m); }} style={{
                padding: "7px 14px", borderRadius: 6, border: `1px solid ${C.border}`, background: C.panel2,
                cursor: "pointer", fontFamily: "inherit", fontSize: 13, fontWeight: 600, color: C.ink,
              }}>{m}</button>
            ))}
          </div>
        )}
        <div style={{ display: "flex", gap: 8 }}>
          <input value={meDraft} placeholder="Your name…" onChange={(e) => setMeDraft(e.target.value)}
            onKeyDown={(e) => e.key === "Enter" && joinAs()} style={{ ...inputStyle, flex: 1, width: "auto" }} />
          <Btn primary onClick={joinAs}>Join</Btn>
        </div>
      </Card>
    </div>
  );

  return (
    <div className="fade">
      <PageHead title="Discussion" sub={`Posting as ${me} — messages sync between founders every few seconds`}
        action={<Btn icon={RefreshCw} small onClick={load}>{syncing ? "Syncing…" : "Refresh"}</Btn>} />

      <div style={{ display: "grid", gridTemplateColumns: "250px 1fr", gap: 14, alignItems: "start" }}>
        {/* thread list */}
        <Card style={{ padding: 0, overflow: "hidden" }}>
          <div style={{ padding: "14px 16px 10px", display: "flex", justifyContent: "space-between", alignItems: "center" }}>
            <Micro>Conversations</Micro>
            <Plus size={15} color={C.accent} style={{ cursor: "pointer" }} onClick={() => setNewChatOpen(!newChatOpen)} />
          </div>
          {newChatOpen && (
            <div style={{ padding: "4px 14px 14px", borderBottom: `1px solid ${C.borderSoft}` }}>
              <input value={ncName} placeholder="Chat name (optional)" onChange={(e) => setNcName(e.target.value)} style={{ ...inputStyle, marginBottom: 8 }} />
              <Micro style={{ display: "block", marginBottom: 6 }}>With</Micro>
              <div style={{ display: "flex", flexWrap: "wrap", gap: 6, marginBottom: 10 }}>
                {founders.filter((m) => m !== me).map((m) => {
                  const on = ncMembers.includes(m);
                  return (
                    <button key={m} onClick={() => setNcMembers(on ? ncMembers.filter(x => x !== m) : [...ncMembers, m])} style={{
                      padding: "5px 11px", borderRadius: 6, cursor: "pointer", fontFamily: "inherit", fontSize: 12, fontWeight: 600,
                      border: `1px solid ${on ? C.accent : C.border}`, background: on ? C.accentSoft : "transparent", color: on ? C.accent : C.dim,
                    }}>{m}</button>
                  );
                })}
                {founders.filter((m) => m !== me).length === 0 && (
                  <span style={{ fontSize: 11.5, color: C.faint }}>No other founders yet — add them in Settings or the sign-in menu.</span>
                )}
              </div>
              <Btn small primary onClick={createChat}>Create {ncMembers.length === 1 ? "DM" : "group chat"}</Btn>
            </div>
          )}
          {visibleThreads.map((t) => (
            <button key={t.id} onClick={() => setActive(t.id)} style={{
              display: "block", width: "100%", textAlign: "left", padding: "12px 16px",
              border: "none", borderTop: `1px solid ${C.borderSoft}`, cursor: "pointer", fontFamily: "inherit",
              background: active === t.id ? C.panel2 : "transparent",
            }}>
              <div style={{ fontSize: 13.5, fontWeight: active === t.id ? 700 : 600, color: C.ink }}>
                {t.id === "board" ? "📌 " : ""}{t.name}
              </div>
              <div style={{ fontSize: 11, color: C.faint, marginTop: 2 }}>
                {t.members.length === 0 ? "All founders" : t.members.join(", ")} · {t.messages.length} message{t.messages.length === 1 ? "" : "s"}
              </div>
            </button>
          ))}
        </Card>

        {/* messages */}
        <Card style={{ padding: 0, overflow: "hidden", display: "flex", flexDirection: "column", minHeight: 460 }}>
          <div style={{ padding: "11px 16px", borderBottom: `1px solid ${C.borderSoft}` }}>
            <span style={{ fontWeight: 600, fontSize: 15, fontFamily: "'Geist', sans-serif" }}>{thread.name}</span>
            <span style={{ fontSize: 11.5, color: C.faint, marginLeft: 8 }}>{thread.members.length === 0 ? "visible to all founders" : `private · ${thread.members.join(", ")}`}</span>
          </div>
          <div style={{ flex: 1, overflowY: "auto", padding: 18, display: "flex", flexDirection: "column", gap: 12 }}>
            {thread.messages.length === 0 && (
              <div style={{ textAlign: "center", color: C.faint, fontSize: 13, marginTop: 40 }}>No messages yet — start the conversation.</div>
            )}
            {thread.messages.map((m) => {
              const mine = m.author === me;
              return (
                <div key={m.id} style={{ alignSelf: mine ? "flex-end" : "flex-start", maxWidth: "75%" }}>
                  <div style={{ fontSize: 10.5, color: C.faint, fontWeight: 700, marginBottom: 3, textAlign: mine ? "right" : "left" }}>
                    {m.author} · {fmtTime(m.ts)}
                  </div>
                  <div style={{
                    padding: "10px 14px", borderRadius: 14, fontSize: 13.5, lineHeight: 1.55,
                    background: mine ? C.ink : C.panel2, color: mine ? "#FAF7F2" : C.ink,
                    border: mine ? "none" : `1px solid ${C.border}`,
                  }}>{m.text}</div>
                </div>
              );
            })}
          </div>
          <div style={{ padding: 14, borderTop: `1px solid ${C.borderSoft}`, display: "flex", gap: 8 }}>
            <input value={msg} placeholder={`Message ${thread.id === "board" ? "the board" : thread.name}…`}
              onChange={(e) => setMsg(e.target.value)} onKeyDown={(e) => e.key === "Enter" && sendMsg()}
              style={{ ...inputStyle, flex: 1, width: "auto" }} />
            <button onClick={sendMsg} style={{
              width: 42, borderRadius: 6, border: "none", background: C.accent, cursor: "pointer",
              display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            }}><Send size={15} color="#FFF8F0" /></button>
          </div>
        </Card>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  REGULATORY                                                        */
/* =================================================================== */
function Regulatory({ data, setData, founders = [] }) {
  const ORDER = ["Not Started", "In Progress", "Submitted", "Approved"];
  const patchReg = (id, patch) => setData((d) => ({ ...d, regulatory: d.regulatory.map((r) => r.id === id ? { ...r, ...patch } : r) }));
  const delReg = (id) => setData((d) => ({ ...d, regulatory: d.regulatory.filter((r) => r.id !== id) }));
  const addReg = (item) => setData((d) => ({ ...d, regulatory: [...d.regulatory, { id: Date.now(), item, status: "Not Started", due: "" }] }));

  return (
    <div className="fade">
      <PageHead title="Regulatory & Compliance" sub="FDA, trademark, insurance, labeling & registrations — click any cell to edit" />
      <div style={{ display: "flex" }}>
        <div style={{ width: 4, background: "#C8932E", borderRadius: "4px 0 0 4px", flexShrink: 0 }} />
        <table style={{ width: "100%", borderCollapse: "collapse", border: `1px solid ${C.borderSoft}`, borderLeft: "none" }}>
          <thead><tr style={{ background: C.panel2 }}>
            <th style={{ width: 38 }}></th>
            <th style={TH}>Item</th>
            <th style={{ ...TH, width: 104 }}>Leads</th>
            <th style={{ ...TH, width: 138 }}>Due</th>
            <th style={{ ...TH, width: 96 }}>Countdown</th>
            <th style={{ ...TH, textAlign: "center", width: 116 }}>Status</th>
            <th style={{ width: 32 }}></th>
          </tr></thead>
          <tbody>
            {data.regulatory.map((r) => {
              const hasDue = r.due && r.due !== "—";
              const d = hasDue ? daysUntil(r.due) : null;
              return (
                <tr key={r.id}>
                  <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}><Shield size={14} color={STAGE_COLORS[r.status]} /></td>
                  <td style={{ ...TD, fontWeight: 500 }}><InlineText value={r.item} onCommit={(v) => patchReg(r.id, { item: v })} width={260} /></td>
                  <td style={TD}><Spearheads value={r.spear || []} onChange={(v) => patchReg(r.id, { spear: v })} founders={founders} size={16} /></td>
                  <td style={TD}>
                    <input type="date" value={hasDue ? r.due : ""} onChange={(e) => patchReg(r.id, { due: e.target.value })}
                      style={{ border: "none", background: "transparent", fontFamily: "inherit", fontSize: 12, color: C.dim, outline: "none", cursor: "pointer", width: 118 }} />
                  </td>
                  <td style={{ ...TD, fontSize: 12, fontWeight: 600, color: d !== null && d < 7 ? "#B4574B" : C.faint }}>
                    {d === null ? "—" : d < 0 ? `${-d}d overdue` : `${d}d left`}
                  </td>
                  <ColorCell label={r.status} color={STAGE_COLORS[r.status]} width={116}
                    onClick={() => patchReg(r.id, { status: ORDER[(ORDER.indexOf(r.status) + 1) % ORDER.length] })} />
                  <td style={{ ...TD, textAlign: "center", padding: "7px 6px" }}>
                    <Trash2 size={13} color={C.faint} style={{ cursor: "pointer" }} onClick={() => delReg(r.id)} />
                  </td>
                </tr>
              );
            })}
            <tr>
              <td style={TD}></td>
              <td colSpan={6} style={{ ...TD, padding: "5px 12px" }}>
                <GroupAdd onAdd={addReg} placeholder="Add compliance item" />
              </td>
            </tr>
          </tbody>
        </table>
      </div>
    </div>
  );
}

/* =================================================================== */
/*  DOCS                                                              */
/* =================================================================== */
function Docs({ data }) {
  const [q, setQ] = useState("");
  const filtered = data.docs.filter((d) => d.name.toLowerCase().includes(q.toLowerCase()) || d.tag.toLowerCase().includes(q.toLowerCase()));
  const tagColor = { Fundraising: "#5E7D54", Product: "#5B7A9D", Manufacturing: C.accent, Finance: "#8A6FA8" };
  return (
    <div className="fade">
      <PageHead title="Document Repository" sub="Decks, formulations, contracts & models" />
      <div style={{ position: "relative", marginBottom: 18, maxWidth: 380 }}>
        <Search size={15} color={C.faint} style={{ position: "absolute", left: 13, top: 12 }} />
        <input placeholder="Search everything…" value={q} onChange={(e) => setQ(e.target.value)} style={{
          ...inputStyle, background: C.panel, padding: "10px 12px 10px 38px", borderRadius: 6,
        }} />
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "repeat(3,1fr)", gap: 12 }}>
        {filtered.map((d) => (
          <Card key={d.id} hover style={{ padding: 17 }}>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start" }}>
              <FileText size={20} color={C.dim} />
              <MoreHorizontal size={16} color={C.faint} />
            </div>
            <div style={{ fontSize: 14, fontWeight: 600, marginTop: 14, fontFamily: "'Geist', sans-serif" }}>{d.name}</div>
            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginTop: 12 }}>
              <Badge label={d.tag} color={tagColor[d.tag] || C.faint} />
              <span style={{ fontSize: 11.5, color: C.faint, fontWeight: 600 }}>{d.date}</span>
            </div>
          </Card>
        ))}
      </div>
    </div>
  );
}

/* =================================================================== */
/*  AI COPILOT                                                        */
/* =================================================================== */
function Copilot({ data, onClose }) {
  const [msgs, setMsgs] = useState([
    { role: "ai", text: "I'm your Eudai copilot — ask me anything. I can see your live tasks, financials, product, pipeline, and strategy, and I can draft emails, plans, and copy." },
  ]);
  const [input, setInput] = useState("");
  const [thinking, setThinking] = useState(false);
  const prompts = ["Analyze my runway", "What should I do next?", "Draft an investor intro email", "Critique our GTM"];

  /* live workspace snapshot injected into every conversation */
  const buildContext = () => {
    const ft = finTotals(data.financials, data.product);
    const p = data.product;
    const openTasks = data.tasks.filter((t) => t.status !== "Done");
    const g = data.gtm || {};
    return [
      `Company: Eudai — pre-revenue nootropic energy drink (${p.flavor}, ${p.size}, ${p.caffeine}mg green tea caffeine, Cognizin hero ingredient). Formula is locked.`,
      `Financials: cash ${fmt$(data.company.cash)}, runway ${data.company.runwayMonths} mo, spent to date ${fmt$(ft.spent)}, capital to launch ${fmt$(ft.needed.likely)} likely (${fmt$(ft.needed.low)}–${fmt$(ft.needed.high)}). Setup ${fmt$(ft.setup)}; production run ${data.financials.productionRun.cans.toLocaleString()} cans ≈ ${fmt$(ft.prod.likely)} likely. Pricing $${data.financials.pricing.can}/can, $${data.financials.pricing.twelvePack}/12-pack. Ingredient cost/can ${ft.ingCan > 0 ? "$" + ft.ingCan.toFixed(2) : "not entered yet"}.`,
      `Open tasks (${openTasks.length}): ${openTasks.slice(0, 10).map((t) => `${t.title} [${t.priority}, due ${t.due}, owner ${t.owner}]`).join("; ")}.`,
      `Upcoming meetings: ${(data.scheduled || []).map((m) => `${m.title} (${m.date || "date TBD"})`).join("; ") || "none"}.`,
      `Regulatory: ${data.regulatory.map((r) => `${r.item}: ${r.status}`).join("; ")}.`,
      `Investors: ${data.investors.length === 0 ? "none tracked yet" : data.investors.map((i) => `${i.name}/${i.firm} (${i.status})`).join("; ")}. Vendors: ${data.vendors.length === 0 ? "none yet" : data.vendors.map((v) => `${v.name} (${v.status})`).join(", ")}.`,
      `Strategy: category "${g.category}"; benefit "${g.mainBenefit}"; target ${g.ageRange}; competitors ${(g.competitors || []).slice(0, 6).join(", ")}; channels ${(g.channels || []).join(", ")}.`,
    ].join("\n");
  };

  /* rule-based fallback if the API is unreachable */
  const fallback = (q) => {
    const ql = q.toLowerCase();
    const ft = finTotals(data.financials, data.product);
    if (ql.includes("runway") || ql.includes("cash") || ql.includes("burn"))
      return `Cash is ${fmt$(data.company.cash)}, ${fmt$(ft.spent)} spent to date, capital to launch ${fmt$(ft.needed.likely)} (likely), runway ${data.company.runwayMonths} months. (Offline answer — AI is unreachable right now.)`;
    if (ql.includes("drink") || ql.includes("product") || ql.includes("formula") || ql.includes("ingredient"))
      return `${data.product.name} — ${data.product.flavor}, ${data.product.size}, ${data.product.caffeine}mg green tea caffeine, Cognizin hero. Formula locked, ${data.product.ingredients.length} ingredients. (Offline answer — AI is unreachable right now.)`;
    const open = data.tasks.filter((t) => t.status !== "Done");
    return `${open.length} open tasks, mostly due 6/12. Formula is locked, so focus is setup: LLC, operating agreement, trademark, barcode, insurance, seller's permit, plus the mock deck and website. (Offline answer — AI is unreachable right now.)`;
  };

  const send = async (q) => {
    if (!q.trim() || thinking) return;
    setInput("");
    const history = [...msgs, { role: "user", text: q }];
    setMsgs(history);
    setThinking(true);
    try {
      const apiMessages = history.slice(-9).map((m) => ({
        role: m.role === "ai" ? "assistant" : "user",
        content: m.text,
      }));
      apiMessages[apiMessages.length - 1] = {
        role: "user",
        content: `You are the Eudai copilot inside ControlCenter, the founders' operating system for their pre-revenue nootropic energy drink company. Act like a sharp co-founder: direct, specific, concise (under 150 words unless asked for a draft or document). Ground every answer in the live workspace data below; never invent numbers. You may draft emails, plans, and copy on request.\n\n=== LIVE WORKSPACE DATA ===\n${buildContext()}\n=== END DATA ===\n\nFounder: ${q}`,
      };
      const r = await fetch("/api/claude", {
        method: "POST",
        headers: { "Content-Type": "application/json", ...(window.getAuthToken ? { Authorization: `Bearer ${await window.getAuthToken()}` } : {}) },
        body: JSON.stringify({ model: "claude-sonnet-4-20250514", max_tokens: 1000, messages: apiMessages }),
      });
      const d = await r.json();
      const text = (d.content || []).filter((i) => i.type === "text").map((i) => i.text).join("\n").trim();
      setMsgs((m) => [...m, { role: "ai", text: text || fallback(q) }]);
    } catch {
      setMsgs((m) => [...m, { role: "ai", text: fallback(q) }]);
    }
    setThinking(false);
  };

  return (
    <div style={{
      position: "fixed", right: 0, top: 0, bottom: 0, width: 400, background: C.panel,
      borderLeft: `1px solid ${C.border}`, zIndex: 90, display: "flex", flexDirection: "column",
      animation: "slidein .22s ease", boxShadow: "-20px 0 60px rgba(60,45,25,.12)",
    }}>
      <div style={{ padding: "20px 22px", borderBottom: `1px solid ${C.borderSoft}`, display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 9 }}>
          <Sparkles size={17} color={C.accent} />
          <span style={{ fontWeight: 600, fontSize: 17, fontFamily: "'Geist', sans-serif" }}>AI Copilot</span>
        </div>
        <X size={18} color={C.dim} style={{ cursor: "pointer" }} onClick={onClose} />
      </div>

      <div style={{ flex: 1, overflowY: "auto", padding: 18, display: "flex", flexDirection: "column", gap: 14, background: C.bg }}>
        {msgs.map((m, i) => (
          <div key={i} style={{ alignSelf: m.role === "user" ? "flex-end" : "flex-start", maxWidth: "85%" }}>
            <div style={{
              padding: "12px 15px", borderRadius: 14, fontSize: 13, lineHeight: 1.6,
              background: m.role === "user" ? C.accent : C.panel,
              color: m.role === "user" ? "#FFF8F0" : C.ink,
              border: m.role === "user" ? "none" : `1px solid ${C.borderSoft}`,
              fontWeight: m.role === "user" ? 600 : 400,
              boxShadow: "0 1px 3px rgba(60,45,25,.06)",
            }}>{m.text}</div>
          </div>
        ))}
        {thinking && (
          <div style={{ alignSelf: "flex-start", display: "flex", alignItems: "center", gap: 8, fontSize: 12.5, color: C.dim, padding: "8px 12px" }}>
            <Sparkles size={13} color={C.accent} /> Thinking…
          </div>
        )}
      </div>

      <div style={{ padding: 16, borderTop: `1px solid ${C.borderSoft}` }}>
        <div style={{ display: "flex", flexWrap: "wrap", gap: 7, marginBottom: 12 }}>
          {prompts.map((p) => (
            <button key={p} onClick={() => send(p)} style={{
              fontSize: 11.5, padding: "6px 12px", borderRadius: 6, cursor: "pointer",
              background: C.panel2, border: `1px solid ${C.border}`, color: C.dim,
              fontFamily: "'Geist', sans-serif", fontWeight: 600,
            }}>{p}</button>
          ))}
        </div>
        <div style={{ display: "flex", gap: 8 }}>
          <input value={input} onChange={(e) => setInput(e.target.value)} onKeyDown={(e) => e.key === "Enter" && send(input)}
            placeholder="Ask anything…" style={{ ...inputStyle, borderRadius: 6, flex: 1 }} />
          <button onClick={() => send(input)} style={{
            width: 42, height: 40, borderRadius: 7, border: "none", background: C.accent, cursor: "pointer",
            display: "flex", alignItems: "center", justifyContent: "center", flexShrink: 0,
            boxShadow: "0 2px 8px rgba(200,115,46,.3)",
          }}><Send size={16} color="#FFF8F0" /></button>
        </div>
      </div>
    </div>
  );
}
