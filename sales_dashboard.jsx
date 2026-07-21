import React, { useState, useMemo, useCallback, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ResponsiveContainer, AreaChart, Area, LineChart, Line, BarChart, Bar,
  PieChart, Pie, Cell, XAxis, YAxis, CartesianGrid, Tooltip, Legend,
} from "recharts";
import {
  TrendingUp, TrendingDown, Upload, RotateCcw, ChevronUp, ChevronDown,
  ChevronsUpDown, X, Award, Receipt, Boxes, Wallet, AlertCircle,
} from "lucide-react";

/* ------------------------------------------------------------------ */
/*  Design tokens — "Ledger" identity                                  */
/* ------------------------------------------------------------------ */
const T = {
  ink: "#182A33",
  inkSoft: "#33474F",
  paper: "#FBF7EE",
  paperDim: "#F1EADA",
  card: "#FFFDF7",
  line: "#DCD2B8",
  gold: "#B8842E",
  goldDeep: "#8E6421",
  teal: "#3C6E67",
  brick: "#A6453A",
  slate: "#6B7A80",
  ok: "#3C6E67",
  bad: "#A6453A",
};
const CATEGORY_COLORS = ["#B8842E", "#3C6E67", "#A6453A", "#5B7A9C", "#8A6BA8"];

/* ------------------------------------------------------------------ */
/*  Synthetic dataset generator (deterministic PRNG)                   */
/* ------------------------------------------------------------------ */
function mulberry32(seed) {
  return function () {
    seed |= 0; seed = (seed + 0x6d2b79f5) | 0;
    let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
    t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

const PRODUCTS = {
  Electronics: [
    { name: "Aurora Soundbar", price: 189 },
    { name: "Nimbus Tablet", price: 349 },
    { name: "Pulse Earbuds", price: 79 },
    { name: "Halo Smartwatch", price: 229 },
  ],
  Apparel: [
    { name: "Drift Jacket", price: 118 },
    { name: "Ember Hoodie", price: 64 },
    { name: "Solstice Tee", price: 28 },
    { name: "Meridian Denim", price: 86 },
  ],
  "Home & Garden": [
    { name: "Terra Planter", price: 42 },
    { name: "Loom Throw Blanket", price: 58 },
    { name: "Hearth Candle Set", price: 34 },
    { name: "Glade Diffuser", price: 46 },
  ],
  Sports: [
    { name: "Apex Yoga Mat", price: 39 },
    { name: "Ridge Trail Backpack", price: 96 },
    { name: "Current Water Bottle", price: 24 },
    { name: "Summit Resistance Bands", price: 22 },
  ],
  Beauty: [
    { name: "Velvet Lip Set", price: 32 },
    { name: "Dawn Serum", price: 54 },
    { name: "Bloom Body Wash", price: 19 },
    { name: "Mist Facial Toner", price: 27 },
  ],
};
const REGIONS = ["North", "South", "East", "West", "Central"];
const MONTH_NAMES = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function genSampleData() {
  const rng = mulberry32(20260721);
  const rows = [];
  let id = 1;
  // 18 months: Jan 2025 -> Jun 2026
  for (let m = 0; m < 18; m++) {
    const year = 2025 + Math.floor(m / 12);
    const month = m % 12;
    const seasonal = 1 + 0.22 * Math.sin((month / 12) * Math.PI * 2 - 1.4); // seasonal wave
    const growth = 1 + m * 0.018; // gentle upward trend
    const holidayBoost = month === 11 ? 1.45 : month === 10 ? 1.15 : 1;
    const txCount = Math.round((26 + rng() * 14) * seasonal * growth * holidayBoost);
    for (let t = 0; t < txCount; t++) {
      const cats = Object.keys(PRODUCTS);
      const cat = cats[Math.floor(rng() * cats.length)];
      const list = PRODUCTS[cat];
      const prod = list[Math.floor(rng() * list.length)];
      const region = REGIONS[Math.floor(rng() * REGIONS.length)];
      const units = 1 + Math.floor(rng() * 5);
      const priceJitter = 0.9 + rng() * 0.25;
      const revenue = Math.round(prod.price * priceJitter * units * 100) / 100;
      const day = 1 + Math.floor(rng() * 28);
      const date = `${year}-${String(month + 1).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
      rows.push({
        id: id++,
        date,
        product: prod.name,
        category: cat,
        region,
        units,
        revenue,
        customer: `Customer-${1000 + Math.floor(rng() * 4200)}`,
      });
    }
  }
  return rows;
}

/* ------------------------------------------------------------------ */
/*  Formatting helpers                                                 */
/* ------------------------------------------------------------------ */
const fmtCurrency = (n) =>
  "$" + n.toLocaleString("en-US", { maximumFractionDigits: 0 });
const fmtCurrencyPrecise = (n) =>
  "$" + n.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtNumber = (n) => n.toLocaleString("en-US");
const monthKey = (dateStr) => dateStr.slice(0, 7);
const monthLabel = (key) => {
  const [y, m] = key.split("-");
  return `${MONTH_NAMES[parseInt(m, 10) - 1]} '${y.slice(2)}`;
};

/* ------------------------------------------------------------------ */
/*  Small UI atoms                                                     */
/* ------------------------------------------------------------------ */
function LedgerCard({ children, style, className = "" }) {
  return (
    <div
      className={`relative ${className}`}
      style={{
        background: T.card,
        border: `1px solid ${T.line}`,
        borderRadius: 3,
        boxShadow: "0 1px 0 rgba(24,42,51,0.04)",
        ...style,
      }}
    >
      {/* perforated top edge */}
      <div
        style={{
          position: "absolute",
          top: 0, left: 10, right: 10, height: 1,
          backgroundImage: `radial-gradient(circle, ${T.paperDim} 1.4px, transparent 1.4px)`,
          backgroundSize: "8px 1px",
          transform: "translateY(-0.5px)",
        }}
      />
      {children}
    </div>
  );
}

function KpiCard({ icon: Icon, label, value, sub, trend, accent }) {
  return (
    <LedgerCard className="px-5 py-4 flex flex-col gap-2">
      <div className="flex items-center justify-between">
        <span
          className="text-[10px] tracking-[0.14em] uppercase font-semibold"
          style={{ color: T.slate, fontFamily: "'Inter',sans-serif" }}
        >
          {label}
        </span>
        <Icon size={15} style={{ color: accent || T.gold }} strokeWidth={2} />
      </div>
      <div
        style={{
          fontFamily: "'IBM Plex Mono',monospace",
          fontSize: 26,
          fontWeight: 600,
          color: T.ink,
          letterSpacing: "-0.01em",
        }}
      >
        {value}
      </div>
      {sub && (
        <div className="flex items-center gap-1" style={{ fontFamily: "'Inter',sans-serif" }}>
          {trend != null && (
            trend >= 0 ? (
              <TrendingUp size={12} style={{ color: T.ok }} />
            ) : (
              <TrendingDown size={12} style={{ color: T.bad }} />
            )
          )}
          <span
            className="text-[11px]"
            style={{ color: trend == null ? T.slate : trend >= 0 ? T.ok : T.bad }}
          >
            {sub}
          </span>
        </div>
      )}
    </LedgerCard>
  );
}

function CheckboxRow({ label, checked, onChange, color }) {
  return (
    <label
      className="flex items-center gap-2 py-1 cursor-pointer select-none"
      style={{ fontFamily: "'Inter',sans-serif", fontSize: 13, color: T.inkSoft }}
    >
      <span
        onClick={onChange}
        style={{
          width: 14, height: 14, borderRadius: 3,
          border: `1.4px solid ${checked ? (color || T.gold) : T.line}`,
          background: checked ? (color || T.gold) : "transparent",
          display: "inline-flex", alignItems: "center", justifyContent: "center",
          flexShrink: 0,
        }}
      >
        {checked && (
          <svg width="8" height="8" viewBox="0 0 8 8">
            <path d="M1 4l2 2 4-4" stroke="#FFFDF7" strokeWidth="1.5" fill="none" strokeLinecap="round" strokeLinejoin="round" />
          </svg>
        )}
      </span>
      {label}
    </label>
  );
}

/* ------------------------------------------------------------------ */
/*  Main component                                                     */
/* ------------------------------------------------------------------ */
export default function SalesDashboard() {
  const [data, setData] = useState(() => genSampleData());
  const [source, setSource] = useState("sample");
  const [importError, setImportError] = useState(null);

  const allMonths = useMemo(() => {
    const set = new Set(data.map((r) => monthKey(r.date)));
    return Array.from(set).sort();
  }, [data]);

  const [monthFrom, setMonthFrom] = useState(0);
  const [monthTo, setMonthTo] = useState(() => Math.max(0, allMonths.length - 1));
  const [selCategories, setSelCategories] = useState(new Set());
  const [selRegions, setSelRegions] = useState(new Set());
  const [productQuery, setProductQuery] = useState("");
  const [sortConfig, setSortConfig] = useState({ key: "date", dir: "desc" });
  const [page, setPage] = useState(1);
  const fileRef = useRef(null);
  const PAGE_SIZE = 8;

  const categories = useMemo(() => Array.from(new Set(data.map((r) => r.category))).sort(), [data]);
  const regions = useMemo(() => Array.from(new Set(data.map((r) => r.region))).sort(), [data]);

  // keep month range valid whenever dataset changes
  React.useEffect(() => {
    setMonthFrom(0);
    setMonthTo(Math.max(0, allMonths.length - 1));
    setSelCategories(new Set());
    setSelRegions(new Set());
    setPage(1);
  }, [source]); // eslint-disable-line

  const filtered = useMemo(() => {
    const lo = allMonths[monthFrom];
    const hi = allMonths[monthTo];
    const q = productQuery.trim().toLowerCase();
    return data.filter((r) => {
      const mk = monthKey(r.date);
      if (lo && mk < lo) return false;
      if (hi && mk > hi) return false;
      if (selCategories.size && !selCategories.has(r.category)) return false;
      if (selRegions.size && !selRegions.has(r.region)) return false;
      if (q && !r.product.toLowerCase().includes(q)) return false;
      return true;
    });
  }, [data, allMonths, monthFrom, monthTo, selCategories, selRegions, productQuery]);

  /* ---------------- KPIs ---------------- */
  const kpis = useMemo(() => {
    const totalRevenue = filtered.reduce((s, r) => s + r.revenue, 0);
    const totalUnits = filtered.reduce((s, r) => s + r.units, 0);
    const orders = filtered.length;
    const aov = orders ? totalRevenue / orders : 0;

    // month over month growth using the two most recent months present in filtered range
    const byMonth = {};
    filtered.forEach((r) => {
      const mk = monthKey(r.date);
      byMonth[mk] = (byMonth[mk] || 0) + r.revenue;
    });
    const monthsPresent = Object.keys(byMonth).sort();
    let momGrowth = null;
    if (monthsPresent.length >= 2) {
      const last = byMonth[monthsPresent[monthsPresent.length - 1]];
      const prev = byMonth[monthsPresent[monthsPresent.length - 2]];
      if (prev > 0) momGrowth = ((last - prev) / prev) * 100;
    }

    const byProduct = {};
    filtered.forEach((r) => {
      byProduct[r.product] = (byProduct[r.product] || 0) + r.revenue;
    });
    const topProduct = Object.entries(byProduct).sort((a, b) => b[1] - a[1])[0];

    return { totalRevenue, totalUnits, orders, aov, momGrowth, topProduct };
  }, [filtered]);

  /* ---------------- Chart data ---------------- */
  const trendData = useMemo(() => {
    const byMonth = {};
    filtered.forEach((r) => {
      const mk = monthKey(r.date);
      byMonth[mk] = (byMonth[mk] || 0) + r.revenue;
    });
    return Object.keys(byMonth).sort().map((mk) => ({
      month: monthLabel(mk),
      revenue: Math.round(byMonth[mk]),
    }));
  }, [filtered]);

  const topProductsData = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      map[r.product] = (map[r.product] || 0) + r.revenue;
    });
    return Object.entries(map)
      .sort((a, b) => b[1] - a[1])
      .slice(0, 8)
      .map(([name, revenue]) => ({ name, revenue: Math.round(revenue) }));
  }, [filtered]);

  const categoryData = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      map[r.category] = (map[r.category] || 0) + r.revenue;
    });
    return Object.entries(map)
      .sort((a, b) => b[1] - a[1])
      .map(([name, value]) => ({ name, value: Math.round(value) }));
  }, [filtered]);

  const regionData = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      map[r.region] = (map[r.region] || 0) + r.revenue;
    });
    return REGIONS.filter((r) => map[r] !== undefined).map((name) => ({
      name, revenue: Math.round(map[name] || 0),
    }));
  }, [filtered]);

  /* ---------------- Table ---------------- */
  const sorted = useMemo(() => {
    const arr = [...filtered];
    const { key, dir } = sortConfig;
    arr.sort((a, b) => {
      let av = a[key], bv = b[key];
      if (typeof av === "string") { av = av.toLowerCase(); bv = bv.toLowerCase(); }
      if (av < bv) return dir === "asc" ? -1 : 1;
      if (av > bv) return dir === "asc" ? 1 : -1;
      return 0;
    });
    return arr;
  }, [filtered, sortConfig]);

  const pageCount = Math.max(1, Math.ceil(sorted.length / PAGE_SIZE));
  const pageRows = sorted.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE);

  const toggleSort = (key) => {
    setSortConfig((prev) =>
      prev.key === key ? { key, dir: prev.dir === "asc" ? "desc" : "asc" } : { key, dir: "desc" }
    );
    setPage(1);
  };

  const toggleSet = (setter) => (value) => {
    setter((prev) => {
      const next = new Set(prev);
      next.has(value) ? next.delete(value) : next.add(value);
      return next;
    });
    setPage(1);
  };

  /* ---------------- CSV / Excel import ---------------- */
  const normalizeRows = useCallback((rows, filename) => {
    const required = ["date", "product", "category", "region", "revenue", "units"];
    const cols = rows.length ? Object.keys(rows[0]).map((c) => String(c).trim().toLowerCase()) : [];
    const missing = required.filter((c) => !cols.includes(c));
    if (!rows.length || missing.length) {
      setImportError(
        `Couldn't import "${filename}" — missing column${missing.length > 1 ? "s" : ""}: ${missing.join(", ") || "no rows found"}. Expected: date, product, category, region, revenue, units, customer (optional).`
      );
      return;
    }
    // Excel serial-date handling: SheetJS may hand back numbers for date cells
    const excelSerialToISO = (n) => {
      const utcDays = Math.floor(n) - 25569;
      const utcMs = utcDays * 86400 * 1000;
      const d = new Date(utcMs);
      return d.toISOString().slice(0, 10);
    };
    const norm = rows
      .map((r, i) => {
        const get = (k) => {
          const found = Object.keys(r).find((rk) => String(rk).trim().toLowerCase() === k);
          return found ? r[found] : undefined;
        };
        const revenue = parseFloat(get("revenue"));
        const units = parseInt(get("units"), 10);
        let rawDate = get("date");
        let date = "";
        if (typeof rawDate === "number") date = excelSerialToISO(rawDate);
        else date = String(rawDate || "").trim().slice(0, 10);
        if (!date || isNaN(revenue) || isNaN(units)) return null;
        return {
          id: i + 1,
          date,
          product: String(get("product") || "Unknown").trim(),
          category: String(get("category") || "Uncategorized").trim(),
          region: String(get("region") || "Unspecified").trim(),
          revenue,
          units,
          customer: String(get("customer") || "").trim(),
        };
      })
      .filter(Boolean);
    if (!norm.length) {
      setImportError("No valid rows found after parsing — check date/revenue/units values.");
      return;
    }
    setData(norm);
    setSource(filename);
  }, []);

  const handleFile = useCallback((e) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setImportError(null);
    const ext = file.name.split(".").pop().toLowerCase();

    if (ext === "csv") {
      Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: (res) => normalizeRows(res.data, file.name),
        error: () => setImportError("Failed to read the file. Please upload a valid CSV."),
      });
    } else if (ext === "xlsx" || ext === "xls") {
      const reader = new FileReader();
      reader.onload = (evt) => {
        try {
          const wb = XLSX.read(evt.target.result, { type: "array", cellDates: false });
          const sheet = wb.Sheets[wb.SheetNames[0]];
          const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
          normalizeRows(rows, file.name);
        } catch (err) {
          setImportError("Failed to read the spreadsheet. Please upload a valid .xlsx or .xls file.");
        }
      };
      reader.onerror = () => setImportError("Failed to read the file.");
      reader.readAsArrayBuffer(file);
    } else {
      setImportError(`Unsupported file type ".${ext}". Please upload a .csv, .xlsx, or .xls file.`);
    }
    e.target.value = "";
  }, [normalizeRows]);

  const resetSample = () => {
    setData(genSampleData());
    setSource("sample");
    setImportError(null);
  };

  const dateRangeLabel =
    allMonths.length ? `${monthLabel(allMonths[monthFrom])} – ${monthLabel(allMonths[monthTo])}` : "—";

  /* ------------------------------------------------------------------ */
  return (
    <div
      style={{
        background: T.paper,
        minHeight: "100%",
        fontFamily: "'Inter',sans-serif",
        color: T.ink,
      }}
    >
      <style>{`
        @import url('https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap');
        input[type=range]::-webkit-slider-thumb {
          -webkit-appearance: none; width: 13px; height: 13px; border-radius: 50%;
          background: ${T.gold}; cursor: pointer; margin-top: -5px;
          box-shadow: 0 0 0 3px ${T.paper};
        }
        input[type=range]::-webkit-slider-runnable-track { height: 3px; background: ${T.line}; border-radius: 2px; }
        input[type=range] { -webkit-appearance: none; background: transparent; }
        ::-webkit-scrollbar { width: 8px; height: 8px; }
        ::-webkit-scrollbar-thumb { background: ${T.line}; border-radius: 4px; }
      `}</style>

      {/* Header */}
      <div
        className="px-6 md:px-8 py-5 flex flex-wrap items-center justify-between gap-4"
        style={{ borderBottom: `1px solid ${T.line}`, background: T.ink }}
      >
        <div>
          <div className="flex items-center gap-2">
            <span
              style={{
                width: 8, height: 8, borderRadius: "50%", background: T.gold, display: "inline-block",
              }}
            />
            <span className="text-[10px] tracking-[0.2em] uppercase" style={{ color: "#C9BE9E" }}>
              Sales &amp; Revenue Ledger
            </span>
          </div>
          <h1
            style={{
              fontFamily: "'Fraunces',serif",
              fontWeight: 600,
              fontSize: 28,
              color: T.paper,
              marginTop: 2,
            }}
          >
            Revenue Analysis Dashboard
          </h1>
        </div>
        <div className="flex items-center gap-3">
          <div className="text-right hidden sm:block">
            <div className="text-[10px] uppercase tracking-wider" style={{ color: "#9FB0AC" }}>Data source</div>
            <div className="text-[12px]" style={{ color: "#E9E2CC", fontFamily: "'IBM Plex Mono',monospace" }}>
              {source === "sample" ? "sample_transactions.csv" : source}
            </div>
          </div>
          <button
            onClick={() => fileRef.current?.click()}
            className="flex items-center gap-2 px-3 py-2 text-[12px] font-medium"
            style={{ background: T.gold, color: T.ink, borderRadius: 3 }}
          >
            <Upload size={14} /> Import Data
          </button>
          <input ref={fileRef} type="file" accept=".csv,.xlsx,.xls" onChange={handleFile} className="hidden" />
          {source !== "sample" && (
            <button
              onClick={resetSample}
              className="flex items-center gap-2 px-3 py-2 text-[12px]"
              style={{ border: `1px solid ${T.inkSoft}`, color: "#E9E2CC", borderRadius: 3 }}
              title="Return to sample data"
            >
              <RotateCcw size={13} />
            </button>
          )}
        </div>
      </div>

      {importError && (
        <div
          className="mx-6 md:mx-8 mt-4 px-4 py-3 flex items-start gap-2 text-[12.5px]"
          style={{ background: "#F6E4E0", border: `1px solid ${T.brick}`, color: T.brick, borderRadius: 3 }}
        >
          <AlertCircle size={15} className="mt-0.5 flex-shrink-0" />
          <span>{importError}</span>
          <button onClick={() => setImportError(null)} className="ml-auto">
            <X size={14} />
          </button>
        </div>
      )}

      <div className="px-6 md:px-8 py-6 flex flex-col lg:flex-row gap-6">
        {/* Sidebar filters */}
        <aside className="lg:w-60 flex-shrink-0">
          <LedgerCard className="p-4 sticky top-4">
            <div
              className="text-[10px] tracking-[0.14em] uppercase font-semibold mb-3 pb-2"
              style={{ color: T.slate, borderBottom: `1px dashed ${T.line}` }}
            >
              Filters &amp; Slicers
            </div>

            <div className="mb-4">
              <div className="text-[11px] font-medium mb-1.5" style={{ color: T.inkSoft }}>
                Period · {dateRangeLabel}
              </div>
              <div className="flex flex-col gap-2">
                <input
                  type="range" min={0} max={allMonths.length - 1} value={monthFrom}
                  onChange={(e) => {
                    const v = Math.min(Number(e.target.value), monthTo);
                    setMonthFrom(v); setPage(1);
                  }}
                />
                <input
                  type="range" min={0} max={allMonths.length - 1} value={monthTo}
                  onChange={(e) => {
                    const v = Math.max(Number(e.target.value), monthFrom);
                    setMonthTo(v); setPage(1);
                  }}
                />
              </div>
            </div>

            <div className="mb-4">
              <div className="text-[11px] font-medium mb-1.5" style={{ color: T.inkSoft }}>Product search</div>
              <input
                type="text"
                value={productQuery}
                onChange={(e) => { setProductQuery(e.target.value); setPage(1); }}
                placeholder="e.g. Serum, Backpack…"
                className="w-full px-2.5 py-1.5 text-[12.5px] outline-none"
                style={{ border: `1px solid ${T.line}`, borderRadius: 3, background: T.paper, color: T.ink }}
              />
            </div>

            <div className="mb-4">
              <div className="text-[11px] font-medium mb-1.5" style={{ color: T.inkSoft }}>Category</div>
              {categories.map((c, i) => (
                <CheckboxRow
                  key={c} label={c}
                  checked={selCategories.has(c)}
                  onChange={() => toggleSet(setSelCategories)(c)}
                  color={CATEGORY_COLORS[i % CATEGORY_COLORS.length]}
                />
              ))}
            </div>

            <div className="mb-1">
              <div className="text-[11px] font-medium mb-1.5" style={{ color: T.inkSoft }}>Region</div>
              {regions.map((r) => (
                <CheckboxRow
                  key={r} label={r}
                  checked={selRegions.has(r)}
                  onChange={() => toggleSet(setSelRegions)(r)}
                />
              ))}
            </div>

            {(selCategories.size > 0 || selRegions.size > 0 || productQuery) && (
              <button
                onClick={() => { setSelCategories(new Set()); setSelRegions(new Set()); setProductQuery(""); setPage(1); }}
                className="mt-3 text-[11px] underline"
                style={{ color: T.gold }}
              >
                Clear all filters
              </button>
            )}
          </LedgerCard>
        </aside>

        {/* Main content */}
        <main className="flex-1 min-w-0 flex flex-col gap-6">
          {/* KPI strip */}
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4">
            <KpiCard
              icon={Wallet} label="Total Revenue"
              value={fmtCurrency(kpis.totalRevenue)}
              sub={kpis.momGrowth == null ? "insufficient months for trend" : `${kpis.momGrowth >= 0 ? "+" : ""}${kpis.momGrowth.toFixed(1)}% vs prior month`}
              trend={kpis.momGrowth}
            />
            <KpiCard icon={Boxes} label="Units Sold" value={fmtNumber(kpis.totalUnits)} sub={`${fmtNumber(kpis.orders)} orders`} accent={T.teal} />
            <KpiCard icon={Receipt} label="Avg Order Value" value={fmtCurrencyPrecise(kpis.aov)} sub="revenue ÷ orders" accent={T.brick} />
            <KpiCard
              icon={Award} label="Top Product"
              value={kpis.topProduct ? kpis.topProduct[0] : "—"}
              sub={kpis.topProduct ? `${fmtCurrency(kpis.topProduct[1])} in revenue` : "no data in range"}
              accent={T.gold}
            />
          </div>

          {filtered.length === 0 ? (
            <LedgerCard className="p-10 text-center" style={{ color: T.slate }}>
              No transactions match the current filters. Try widening the date range or clearing a filter.
            </LedgerCard>
          ) : (
            <>
              {/* Trend + Category row */}
              <div className="grid grid-cols-1 xl:grid-cols-5 gap-6">
                <LedgerCard className="p-5 xl:col-span-3">
                  <ChartTitle title="Revenue Trend" note={`${trendData.length} month${trendData.length !== 1 ? "s" : ""} shown`} />
                  <ResponsiveContainer width="100%" height={240}>
                    <AreaChart data={trendData} margin={{ top: 6, right: 8, left: -14, bottom: 0 }}>
                      <defs>
                        <linearGradient id="revFill" x1="0" y1="0" x2="0" y2="1">
                          <stop offset="0%" stopColor={T.gold} stopOpacity={0.35} />
                          <stop offset="100%" stopColor={T.gold} stopOpacity={0.02} />
                        </linearGradient>
                      </defs>
                      <CartesianGrid stroke={T.line} strokeDasharray="3 4" vertical={false} />
                      <XAxis dataKey="month" tick={{ fontSize: 11, fill: T.slate }} axisLine={{ stroke: T.line }} tickLine={false} />
                      <YAxis tick={{ fontSize: 11, fill: T.slate }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v >= 1000 ? (v / 1000).toFixed(0) + "k" : v}`} />
                      <Tooltip content={<LedgerTooltip currency />} />
                      <Area type="monotone" dataKey="revenue" stroke={T.goldDeep} strokeWidth={2} fill="url(#revFill)" name="Revenue" />
                    </AreaChart>
                  </ResponsiveContainer>
                </LedgerCard>

                <LedgerCard className="p-5 xl:col-span-2">
                  <ChartTitle title="Revenue by Category" note="share of total" />
                  <ResponsiveContainer width="100%" height={240}>
                    <PieChart>
                      <Pie
                        data={categoryData} dataKey="value" nameKey="name"
                        innerRadius={54} outerRadius={82} paddingAngle={2}
                      >
                        {categoryData.map((_, i) => (
                          <Cell key={i} fill={CATEGORY_COLORS[i % CATEGORY_COLORS.length]} stroke={T.card} strokeWidth={2} />
                        ))}
                      </Pie>
                      <Tooltip content={<LedgerTooltip currency />} />
                      <Legend
                        iconType="circle" iconSize={7}
                        formatter={(v) => <span style={{ color: T.inkSoft, fontSize: 11.5 }}>{v}</span>}
                      />
                    </PieChart>
                  </ResponsiveContainer>
                </LedgerCard>
              </div>

              {/* Top products + Region row */}
              <div className="grid grid-cols-1 xl:grid-cols-5 gap-6">
                <LedgerCard className="p-5 xl:col-span-3">
                  <ChartTitle title="Top Performing Products" note="by revenue" />
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={topProductsData} layout="vertical" margin={{ top: 4, right: 24, left: 8, bottom: 0 }}>
                      <CartesianGrid stroke={T.line} strokeDasharray="3 4" horizontal={false} />
                      <XAxis type="number" tick={{ fontSize: 11, fill: T.slate }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v >= 1000 ? (v / 1000).toFixed(0) + "k" : v}`} />
                      <YAxis type="category" dataKey="name" width={128} tick={{ fontSize: 11, fill: T.inkSoft }} axisLine={false} tickLine={false} />
                      <Tooltip content={<LedgerTooltip currency />} />
                      <Bar dataKey="revenue" fill={T.teal} radius={[0, 3, 3, 0]} barSize={14} />
                    </BarChart>
                  </ResponsiveContainer>
                </LedgerCard>

                <LedgerCard className="p-5 xl:col-span-2">
                  <ChartTitle title="Revenue by Region" note="geographic split" />
                  <ResponsiveContainer width="100%" height={260}>
                    <BarChart data={regionData} margin={{ top: 4, right: 8, left: -14, bottom: 0 }}>
                      <CartesianGrid stroke={T.line} strokeDasharray="3 4" vertical={false} />
                      <XAxis dataKey="name" tick={{ fontSize: 11, fill: T.slate }} axisLine={{ stroke: T.line }} tickLine={false} />
                      <YAxis tick={{ fontSize: 11, fill: T.slate }} axisLine={false} tickLine={false} tickFormatter={(v) => `$${v >= 1000 ? (v / 1000).toFixed(0) + "k" : v}`} />
                      <Tooltip content={<LedgerTooltip currency />} />
                      <Bar dataKey="revenue" fill={T.brick} radius={[3, 3, 0, 0]} barSize={26} />
                    </BarChart>
                  </ResponsiveContainer>
                </LedgerCard>
              </div>

              {/* Transactions table */}
              <LedgerCard className="p-5">
                <ChartTitle title="Transaction Ledger" note={`${fmtNumber(sorted.length)} records`} />
                <div className="overflow-x-auto -mx-1">
                  <table className="w-full text-left" style={{ fontSize: 12.5, borderCollapse: "collapse" }}>
                    <thead>
                      <tr style={{ borderBottom: `1.5px solid ${T.ink}` }}>
                        {[
                          ["date", "Date"], ["product", "Product"], ["category", "Category"],
                          ["region", "Region"], ["units", "Units"], ["revenue", "Revenue"],
                        ].map(([key, label]) => (
                          <Th key={key} label={label} active={sortConfig.key === key} dir={sortConfig.dir} onClick={() => toggleSort(key)} numeric={key === "units" || key === "revenue"} />
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {pageRows.map((r, i) => (
                        <tr
                          key={r.id}
                          style={{
                            borderBottom: `1px solid ${T.line}`,
                            background: i % 2 === 0 ? "transparent" : T.paperDim + "80",
                          }}
                        >
                          <td className="py-2 px-2" style={{ fontFamily: "'IBM Plex Mono',monospace", color: T.slate }}>{r.date}</td>
                          <td className="py-2 px-2" style={{ color: T.ink }}>{r.product}</td>
                          <td className="py-2 px-2" style={{ color: T.inkSoft }}>{r.category}</td>
                          <td className="py-2 px-2" style={{ color: T.inkSoft }}>{r.region}</td>
                          <td className="py-2 px-2 text-right" style={{ fontFamily: "'IBM Plex Mono',monospace" }}>{r.units}</td>
                          <td className="py-2 px-2 text-right font-medium" style={{ fontFamily: "'IBM Plex Mono',monospace", color: T.ink }}>{fmtCurrencyPrecise(r.revenue)}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
                <div className="flex items-center justify-between mt-4 pt-3" style={{ borderTop: `1px dashed ${T.line}` }}>
                  <span className="text-[11px]" style={{ color: T.slate }}>
                    Page {page} of {pageCount}
                  </span>
                  <div className="flex gap-2">
                    <PageBtn disabled={page === 1} onClick={() => setPage((p) => Math.max(1, p - 1))}>Prev</PageBtn>
                    <PageBtn disabled={page === pageCount} onClick={() => setPage((p) => Math.min(pageCount, p + 1))}>Next</PageBtn>
                  </div>
                </div>
              </LedgerCard>
            </>
          )}

          <div className="text-center text-[11px] pb-2" style={{ color: T.slate }}>
            Import your own CSV or Excel file (columns: date, product, category, region, revenue, units, customer) to replace the sample dataset.
          </div>
        </main>
      </div>
    </div>
  );
}

/* ------------------------------------------------------------------ */
/*  Sub-components                                                     */
/* ------------------------------------------------------------------ */
function ChartTitle({ title, note }) {
  return (
    <div className="flex items-baseline justify-between mb-3">
      <h3 style={{ fontFamily: "'Fraunces',serif", fontWeight: 600, fontSize: 15.5, color: T.ink }}>{title}</h3>
      {note && <span className="text-[10.5px]" style={{ color: T.slate }}>{note}</span>}
    </div>
  );
}

function Th({ label, active, dir, onClick, numeric }) {
  const Icon = active ? (dir === "asc" ? ChevronUp : ChevronDown) : ChevronsUpDown;
  return (
    <th
      onClick={onClick}
      className={`py-2 px-2 cursor-pointer select-none font-semibold ${numeric ? "text-right" : "text-left"}`}
      style={{ color: active ? T.ink : T.slate, fontFamily: "'Inter',sans-serif", fontSize: 11, textTransform: "uppercase", letterSpacing: "0.04em" }}
    >
      <span className={`inline-flex items-center gap-1 ${numeric ? "flex-row-reverse" : ""}`}>
        {label} <Icon size={11} />
      </span>
    </th>
  );
}

function PageBtn({ children, disabled, onClick }) {
  return (
    <button
      disabled={disabled}
      onClick={onClick}
      className="px-2.5 py-1 text-[11px]"
      style={{
        border: `1px solid ${T.line}`, borderRadius: 3,
        color: disabled ? T.line : T.inkSoft,
        cursor: disabled ? "default" : "pointer",
      }}
    >
      {children}
    </button>
  );
}

function LedgerTooltip({ active, payload, label, currency }) {
  if (!active || !payload || !payload.length) return null;
  return (
    <div
      style={{
        background: T.ink, color: T.paper, padding: "8px 11px", borderRadius: 3,
        fontSize: 12, fontFamily: "'Inter',sans-serif", boxShadow: "0 4px 14px rgba(0,0,0,0.18)",
      }}
    >
      {label && <div className="mb-1 opacity-70 text-[10.5px] uppercase tracking-wide">{label}</div>}
      {payload.map((p, i) => (
        <div key={i} style={{ fontFamily: "'IBM Plex Mono',monospace" }}>
          {p.name}: {currency ? fmtCurrency(p.value) : fmtNumber(p.value)}
        </div>
      ))}
    </div>
  );
}
