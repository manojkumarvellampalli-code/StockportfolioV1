import React, { useState, useEffect, useMemo, useRef } from "react";
import { TrendingUp, TrendingDown, Minus, Plus, Upload, Download, Trash2, Calculator, ChevronDown, ChevronUp, AlertTriangle, CheckCircle2, Circle, X, Search, Save, FileSpreadsheet, Settings, Wifi, WifiOff, RefreshCw, Loader2, Zap } from "lucide-react";
import * as XLSX from "xlsx";

// ---------- Constants ----------
const STORAGE_KEY = "nb_portfolio_v1";
const API_KEY_STORAGE = "nb_api_key_v1";
const SELL_DROP_PCT = 15;
const BUY_MORE_GAIN_PCT = 20;
const API_BASE = "https://stock.indianapi.in";

const emptyFundamentals = () => ({
  epsQoQ: "", epsYoY: "", pmQoQ: "", pmYoY: "", patQoQ: "", patYoY: "",
});

const emptyStock = () => ({
  id: crypto.randomUUID(),
  code: "",
  exchange: "",
  name: "",
  qty: "",
  avgBuyPrice: "",
  currentPrice: "",
  highSinceBuy: "",
  fundamentals: emptyFundamentals(),
  notes: "",
});

// ---------- NSE/BSE detection ----------
// Heuristic: BSE codes are 6-digit numeric (500325, 532540...).
// NSE codes are alphabetic symbols (RELIANCE, TCS, INFY...).
function detectExchange(code) {
  if (!code) return "";
  const c = code.trim().toUpperCase();
  if (/^\d{6}$/.test(c)) return "BSE";
  if (/^[A-Z&\-]+$/.test(c)) return "NSE";
  return "UNKNOWN";
}

// ---------- Decision engine ----------
function fundamentalFlags(f) {
  const num = (v) => (v === "" || v === null || v === undefined ? null : parseFloat(v));
  const epsQoQ = num(f.epsQoQ);
  const epsYoY = num(f.epsYoY);
  const pmQoQ = num(f.pmQoQ);
  const pmYoY = num(f.pmYoY);
  const patQoQ = num(f.patQoQ);
  const patYoY = num(f.patYoY);

  // Combine QoQ/YoY pairs: positive if either available figure is > 0 (net positive momentum)
  const epsPositive = [epsQoQ, epsYoY].some((v) => v !== null) && [epsQoQ, epsYoY].every((v) => v === null || v > 0);
  const pmPositive = [pmQoQ, pmYoY].some((v) => v !== null) && [pmQoQ, pmYoY].every((v) => v === null || v > 0);
  const patPositive = [patQoQ, patYoY].some((v) => v !== null) && [patQoQ, patYoY].every((v) => v === null || v > 0);

  const epsNegative = [epsQoQ, epsYoY].some((v) => v !== null && v < 0);
  const pmNegative = [pmQoQ, pmYoY].some((v) => v !== null && v < 0);
  const patNegative = [patQoQ, patYoY].some((v) => v !== null && v < 0);

  return {
    eps: { positive: epsPositive, negative: epsNegative, hasData: epsQoQ !== null || epsYoY !== null },
    pm: { positive: pmPositive, negative: pmNegative, hasData: pmQoQ !== null || pmYoY !== null },
    pat: { positive: patPositive, negative: patNegative, hasData: patQoQ !== null || patYoY !== null },
  };
}

function evaluateStock(stock) {
  const price = parseFloat(stock.currentPrice);
  const avg = parseFloat(stock.avgBuyPrice);
  const high = parseFloat(stock.highSinceBuy) || price;

  const flags = fundamentalFlags(stock.fundamentals);
  const declineCount = [flags.eps.negative, flags.pm.negative, flags.pat.negative].filter(Boolean).length;
  const allPositive = flags.eps.positive && flags.pm.positive && flags.pat.positive;

  const hasPriceData = !isNaN(price) && !isNaN(avg) && avg > 0;
  const dropFromHigh = hasPriceData && !isNaN(high) && high > 0 ? ((high - price) / high) * 100 : null;
  const gainFromAvg = hasPriceData ? ((price - avg) / avg) * 100 : null;

  const sellByPrice = dropFromHigh !== null && dropFromHigh >= SELL_DROP_PCT;
  const sellByFundamentals = declineCount >= 2;

  const buyMoreByPrice = gainFromAvg !== null && gainFromAvg >= BUY_MORE_GAIN_PCT;
  const buyMore = buyMoreByPrice && allPositive;

  let decision = "HOLD";
  let reasons = [];

  if (sellByPrice || sellByFundamentals) {
    decision = "SELL";
    if (sellByPrice) reasons.push(`Price down ${dropFromHigh.toFixed(1)}% from high (≥${SELL_DROP_PCT}% trigger)`);
    if (sellByFundamentals) reasons.push(`${declineCount} of 3 fundamental signals declining (min 2 trigger)`);
  } else if (buyMore) {
    decision = "BUY MORE";
    reasons.push(`Price up ${gainFromAvg.toFixed(1)}% from avg buy (≥${BUY_MORE_GAIN_PCT}% trigger)`);
    reasons.push("All fundamentals positive (EPS, Profit Margin, PAT)");
  } else {
    if (gainFromAvg !== null) reasons.push(`Gain from avg: ${gainFromAvg.toFixed(1)}%`);
    if (dropFromHigh !== null) reasons.push(`Drop from high: ${dropFromHigh.toFixed(1)}%`);
    if (declineCount > 0) reasons.push(`${declineCount} fundamental signal(s) declining`);
    if (reasons.length === 0) reasons.push("Insufficient data — enter price & fundamentals");
  }

  return { decision, reasons, dropFromHigh, gainFromAvg, flags, declineCount, allPositive, hasPriceData };
}

// ---------- Persistence ----------
function loadPortfolio() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];
    return JSON.parse(raw);
  } catch {
    return [];
  }
}
function savePortfolio(stocks) {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(stocks));
  } catch {}
}

// ---------- Row normalization (shared by CSV + Excel import) ----------
// Takes a 2D array of cells (first row = headers) and maps recognized
// columns onto our stock shape, regardless of column order.
function normalizeRows(grid) {
  if (!grid || grid.length < 2) return [];
  const headers = grid[0].map((h) => String(h ?? "").trim().toLowerCase());
  const idx = (name) => headers.findIndex((h) => h.includes(name));

  const iCode = idx("code");
  const iQty = idx("qty") >= 0 ? idx("qty") : idx("quantity");
  const iAvg = idx("avg");
  const iCurrent = idx("current") >= 0 ? idx("current") : idx("price");
  const iHigh = idx("high");
  const iName = idx("name");
  const iEpsQoQ = headers.findIndex((h) => h.includes("eps") && h.includes("qoq"));
  const iEpsYoY = headers.findIndex((h) => h.includes("eps") && h.includes("yoy"));
  const iPmQoQ = headers.findIndex((h) => (h.includes("pm") || h.includes("margin")) && h.includes("qoq"));
  const iPmYoY = headers.findIndex((h) => (h.includes("pm") || h.includes("margin")) && h.includes("yoy"));
  const iPatQoQ = headers.findIndex((h) => h.includes("pat") && h.includes("qoq"));
  const iPatYoY = headers.findIndex((h) => h.includes("pat") && h.includes("yoy"));

  const rows = [];
  for (let i = 1; i < grid.length; i++) {
    const cells = grid[i].map((c) => String(c ?? "").trim());
    if (!cells[iCode]) continue;
    const code = cells[iCode];
    const stock = emptyStock();
    stock.code = code.toUpperCase();
    stock.exchange = detectExchange(code);
    if (iName >= 0) stock.name = cells[iName] || "";
    if (iQty >= 0) stock.qty = cells[iQty] || "";
    if (iAvg >= 0) stock.avgBuyPrice = cells[iAvg] || "";
    if (iCurrent >= 0) stock.currentPrice = cells[iCurrent] || "";
    if (iHigh >= 0) stock.highSinceBuy = cells[iHigh] || "";
    if (iEpsQoQ >= 0) stock.fundamentals.epsQoQ = cells[iEpsQoQ] || "";
    if (iEpsYoY >= 0) stock.fundamentals.epsYoY = cells[iEpsYoY] || "";
    if (iPmQoQ >= 0) stock.fundamentals.pmQoQ = cells[iPmQoQ] || "";
    if (iPmYoY >= 0) stock.fundamentals.pmYoY = cells[iPmYoY] || "";
    if (iPatQoQ >= 0) stock.fundamentals.patQoQ = cells[iPatQoQ] || "";
    if (iPatYoY >= 0) stock.fundamentals.patYoY = cells[iPatYoY] || "";
    rows.push(stock);
  }
  return rows;
}

// ---------- CSV parsing ----------
function parseCSV(text) {
  const lines = text.trim().split(/\r?\n/).filter((l) => l.trim().length > 0);
  const grid = lines.map((l) => l.split(","));
  return normalizeRows(grid);
}

// ---------- Excel (.xlsx / .xls) parsing ----------
function parseExcel(arrayBuffer) {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const grid = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  return normalizeRows(grid);
}

// ---------- API key persistence ----------
function loadApiKey() {
  try {
    return localStorage.getItem(API_KEY_STORAGE) || "";
  } catch {
    return "";
  }
}
function saveApiKey(key) {
  try {
    localStorage.setItem(API_KEY_STORAGE, key);
  } catch {}
}

// ---------- Live data fetch (Indian API - stock.indianapi.in) ----------
// Docs: https://indianapi.in/documentation/indian-stock-market
// Auth header: x-api-key
async function fetchStockData(code, apiKey) {
  if (!apiKey) throw new Error("No API key set. Add one in Settings.");
  const url = `${API_BASE}/stock?name=${encodeURIComponent(code)}`;
  const res = await fetch(url, {
    headers: { "x-api-key": apiKey },
  });
  if (!res.ok) {
    if (res.status === 401 || res.status === 403) throw new Error("Invalid API key.");
    if (res.status === 429) throw new Error("Rate limit / quota exceeded on your API plan.");
    if (res.status === 404) throw new Error("Stock not found for this code.");
    throw new Error(`API error (${res.status}).`);
  }
  const data = await res.json();
  return parseApiResponse(data);
}

// Normalizes the Indian API response shape into our stock fields.
// The API's exact field names can vary by endpoint version, so this
// checks several likely paths defensively rather than assuming one shape.
function parseApiResponse(data) {
  const dig = (obj, paths) => {
    for (const path of paths) {
      const val = path.split(".").reduce((o, k) => (o && o[k] !== undefined ? o[k] : undefined), obj);
      if (val !== undefined && val !== null && val !== "") return val;
    }
    return null;
  };

  const currentPrice = dig(data, [
    "currentPrice.NSE", "currentPrice.BSE", "currentPrice", "price", "lastPrice", "ltp",
  ]);
  const high52 = dig(data, [
    "yearHigh", "52WeekHigh", "high52Week", "stockTechnicalData.yearHigh", "keyMetrics.yearHigh",
  ]);
  const name = dig(data, ["companyName", "name", "companyProfile.companyName"]);

  // Fundamentals are often nested under financials/quarterly results arrays;
  // shape varies, so we try to pull the most recent quarter's EPS/PAT/margin
  // trend if present, but this is best-effort — confirm before relying on it.
  const eps = dig(data, ["keyMetrics.eps", "financials.eps", "eps"]);
  const pat = dig(data, ["keyMetrics.pat", "financials.pat", "netProfit"]);
  const margin = dig(data, ["keyMetrics.netProfitMargin", "financials.profitMargin", "profitMargin"]);

  return {
    currentPrice: currentPrice !== null ? String(currentPrice) : "",
    high52: high52 !== null ? String(high52) : "",
    name: name || "",
    epsRaw: eps,
    patRaw: pat,
    marginRaw: margin,
    raw: data,
  };
}

// ---------- UI Bits ----------
const decisionStyle = {
  SELL: { bg: "bg-rose-50", border: "border-rose-300", text: "text-rose-700", dot: "bg-rose-500", icon: TrendingDown },
  "BUY MORE": { bg: "bg-emerald-50", border: "border-emerald-300", text: "text-emerald-700", dot: "bg-emerald-500", icon: TrendingUp },
  HOLD: { bg: "bg-amber-50", border: "border-amber-300", text: "text-amber-700", dot: "bg-amber-500", icon: Minus },
};

function DecisionBadge({ decision }) {
  const s = decisionStyle[decision] || decisionStyle.HOLD;
  const Icon = s.icon;
  return (
    <div className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full border ${s.bg} ${s.border} ${s.text} text-xs font-semibold tracking-wide`}>
      <Icon size={13} strokeWidth={2.5} />
      {decision}
    </div>
  );
}

function FundamentalPill({ label, hasData, positive, negative }) {
  let cls = "bg-stone-100 text-stone-400 border-stone-200";
  let Icon = Circle;
  if (hasData) {
    if (positive) { cls = "bg-emerald-50 text-emerald-700 border-emerald-200"; Icon = CheckCircle2; }
    else if (negative) { cls = "bg-rose-50 text-rose-700 border-rose-200"; Icon = TrendingDown; }
    else { cls = "bg-stone-100 text-stone-500 border-stone-200"; Icon = Minus; }
  }
  return (
    <span className={`inline-flex items-center gap-1 px-2 py-0.5 rounded border text-[10px] font-medium ${cls}`}>
      <Icon size={10} />
      {label}
    </span>
  );
}

function NumInput({ value, onChange, placeholder, prefix, className = "" }) {
  return (
    <div className="relative">
      {prefix && <span className="absolute left-2.5 top-1/2 -translate-y-1/2 text-stone-400 text-xs">{prefix}</span>}
      <input
        type="number"
        inputMode="decimal"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className={`w-full border border-stone-200 rounded-lg py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400 bg-white ${prefix ? "pl-6" : "pl-2.5"} pr-2.5 ${className}`}
      />
    </div>
  );
}

// ---------- Average Calculator ----------
function AverageCalculator({ onClose, onApply }) {
  const [lots, setLots] = useState([{ id: 1, qty: "", price: "" }]);

  const addLot = () => setLots((l) => [...l, { id: Date.now(), qty: "", price: "" }]);
  const removeLot = (id) => setLots((l) => l.filter((x) => x.id !== id));
  const updateLot = (id, field, val) => setLots((l) => l.map((x) => (x.id === id ? { ...x, [field]: val } : x)));

  const result = useMemo(() => {
    let totalQty = 0, totalCost = 0;
    for (const lot of lots) {
      const q = parseFloat(lot.qty);
      const p = parseFloat(lot.price);
      if (!isNaN(q) && !isNaN(p)) {
        totalQty += q;
        totalCost += q * p;
      }
    }
    return { totalQty, totalCost, avg: totalQty > 0 ? totalCost / totalQty : 0 };
  }, [lots]);

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-md max-h-[85vh] overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white border-b border-stone-100 px-5 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Calculator size={18} className="text-teal-600" />
            <h3 className="font-semibold text-stone-800">Average Price Calculator</h3>
          </div>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600"><X size={20} /></button>
        </div>
        <div className="p-5 space-y-3">
          {lots.map((lot, i) => (
            <div key={lot.id} className="flex gap-2 items-center">
              <span className="text-xs text-stone-400 w-4">{i + 1}</span>
              <NumInput value={lot.qty} onChange={(v) => updateLot(lot.id, "qty", v)} placeholder="Qty" />
              <NumInput value={lot.price} onChange={(v) => updateLot(lot.id, "price", v)} placeholder="Price" prefix="₹" />
              {lots.length > 1 && (
                <button onClick={() => removeLot(lot.id)} className="text-stone-300 hover:text-rose-500 shrink-0">
                  <Trash2 size={16} />
                </button>
              )}
            </div>
          ))}
          <button onClick={addLot} className="text-xs font-medium text-teal-600 flex items-center gap-1 hover:text-teal-700">
            <Plus size={14} /> Add purchase lot
          </button>

          <div className="mt-4 pt-4 border-t border-stone-100 bg-teal-50/50 -mx-5 px-5 py-4">
            <div className="grid grid-cols-2 gap-3 text-sm">
              <div>
                <div className="text-stone-400 text-xs">Total Qty</div>
                <div className="font-semibold text-stone-800">{result.totalQty || "—"}</div>
              </div>
              <div>
                <div className="text-stone-400 text-xs">Total Cost</div>
                <div className="font-semibold text-stone-800">₹{result.totalCost ? result.totalCost.toFixed(2) : "—"}</div>
              </div>
            </div>
            <div className="mt-3">
              <div className="text-stone-400 text-xs">New Average Price</div>
              <div className="font-bold text-teal-700 text-2xl">₹{result.avg ? result.avg.toFixed(2) : "0.00"}</div>
            </div>
          </div>

          <div className="flex gap-2 pt-2">
            <button onClick={onClose} className="flex-1 py-2.5 rounded-lg border border-stone-200 text-stone-600 text-sm font-medium">Close</button>
            <button
              onClick={() => onApply(result)}
              disabled={!result.avg}
              className="flex-1 py-2.5 rounded-lg bg-teal-600 text-white text-sm font-medium disabled:opacity-40"
            >
              Apply to Stock
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Settings Panel ----------
function SettingsPanel({ apiKey, onSave, onClose }) {
  const [key, setKey] = useState(apiKey);
  const [testing, setTesting] = useState(false);
  const [testResult, setTestResult] = useState(null);

  const testKey = async () => {
    setTesting(true);
    setTestResult(null);
    try {
      await fetchStockData("RELIANCE", key);
      setTestResult({ ok: true, msg: "Connected successfully." });
    } catch (e) {
      setTestResult({ ok: false, msg: e.message });
    }
    setTesting(false);
  };

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-md max-h-[85vh] overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white border-b border-stone-100 px-5 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <Settings size={18} className="text-teal-600" />
            <h3 className="font-semibold text-stone-800">Live Data Settings</h3>
          </div>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600"><X size={20} /></button>
        </div>
        <div className="p-5 space-y-4">
          <div>
            <label className="text-xs font-medium text-stone-500 mb-1 block">Indian API Key</label>
            <input
              value={key}
              onChange={(e) => { setKey(e.target.value); setTestResult(null); }}
              placeholder="Paste your API key"
              className="w-full border border-stone-200 rounded-lg px-2.5 py-2 text-sm font-mono focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400"
            />
            <p className="text-[11px] text-stone-400 mt-1.5">
              Get a free key at <span className="font-medium">indianapi.in</span> → sign in → Indian Stock Market API → Pricing → Free plan → Dashboard → API Keys.
            </p>
          </div>

          <div className="rounded-lg bg-amber-50 border border-amber-200 p-3">
            <p className="text-[11px] text-amber-800 leading-relaxed">
              This key is stored only on your device (localStorage) and sent directly from your browser to indianapi.in. It is never sent to us. Because it's client-side, avoid sharing this device/app publicly with your key saved.
            </p>
          </div>

          <button
            onClick={testKey}
            disabled={!key || testing}
            className="w-full py-2.5 rounded-lg border border-stone-200 text-stone-600 text-sm font-medium flex items-center justify-center gap-2 disabled:opacity-40"
          >
            {testing ? <Loader2 size={14} className="animate-spin" /> : <Wifi size={14} />}
            Test Connection
          </button>
          {testResult && (
            <div className={`text-xs px-3 py-2 rounded-lg ${testResult.ok ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"}`}>
              {testResult.msg}
            </div>
          )}

          <div className="flex gap-2 pt-2">
            <button onClick={onClose} className="flex-1 py-2.5 rounded-lg border border-stone-200 text-stone-600 text-sm font-medium">Cancel</button>
            <button
              onClick={() => { onSave(key); onClose(); }}
              className="flex-1 py-2.5 rounded-lg bg-teal-600 text-white text-sm font-semibold"
            >
              Save Key
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- Stock Editor ----------
function StockEditor({ stock, onChange, onDelete, onClose, apiKey, onNeedApiKey }) {
  const [local, setLocal] = useState(stock);
  const [showCalc, setShowCalc] = useState(false);
  const [fetching, setFetching] = useState(false);
  const [fetchMsg, setFetchMsg] = useState(null);

  useEffect(() => setLocal(stock), [stock.id]);

  const update = (field, value) => {
    let next = { ...local, [field]: value };
    if (field === "code") next.exchange = detectExchange(value);
    setLocal(next);
  };
  const updateFund = (field, value) => {
    setLocal((l) => ({ ...l, fundamentals: { ...l.fundamentals, [field]: value } }));
  };

  const save = () => {
    onChange(local);
    onClose();
  };

  const fetchLive = async () => {
    if (!local.code) {
      setFetchMsg({ ok: false, msg: "Enter a stock code first." });
      return;
    }
    if (!apiKey) {
      onNeedApiKey();
      return;
    }
    setFetching(true);
    setFetchMsg(null);
    try {
      const result = await fetchStockData(local.code, apiKey);
      setLocal((l) => ({
        ...l,
        currentPrice: result.currentPrice || l.currentPrice,
        highSinceBuy: result.high52 || l.highSinceBuy,
        name: result.name || l.name,
      }));
      const gotFundamentals = result.epsRaw !== null || result.patRaw !== null || result.marginRaw !== null;
      setFetchMsg({
        ok: true,
        msg: gotFundamentals
          ? "Price & fundamentals updated. Please verify fundamentals — the API's quarterly breakdown may not map directly to QoQ/YoY fields."
          : "Price updated. This API doesn't reliably return quarterly EPS/PAT/margin trends — enter those manually below.",
      });
    } catch (e) {
      setFetchMsg({ ok: false, msg: e.message });
    }
    setFetching(false);
  };

  const evalResult = useMemo(() => evaluateStock(local), [local]);

  return (
    <div className="fixed inset-0 z-40 flex items-end sm:items-center justify-center bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-lg max-h-[92vh] overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white border-b border-stone-100 px-5 py-4 flex items-center justify-between z-10">
          <h3 className="font-semibold text-stone-800">{stock.code ? "Edit Holding" : "Add Holding"}</h3>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600"><X size={20} /></button>
        </div>

        <div className="p-5 space-y-5">
          {/* Identity */}
          <div className="grid grid-cols-3 gap-3">
            <div className="col-span-2">
              <label className="text-xs font-medium text-stone-500 mb-1 block">Stock Code</label>
              <input
                value={local.code}
                onChange={(e) => update("code", e.target.value)}
                onBlur={() => {
                  // Auto-fetch live data the moment a code is entered for a
                  // brand-new holding, so CMP/high/name fill in without an
                  // extra manual tap. Skipped if already fetched or no key.
                  if (local.code && !local.currentPrice && apiKey && !fetching) fetchLive();
                }}
                placeholder="e.g. RELIANCE or 500325"
                className="w-full border border-stone-200 rounded-lg px-2.5 py-2 text-sm uppercase focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400"
              />
            </div>
            <div>
              <label className="text-xs font-medium text-stone-500 mb-1 block">Exchange</label>
              <div className={`px-2.5 py-2 rounded-lg text-sm font-semibold text-center border ${
                local.exchange === "NSE" ? "bg-indigo-50 text-indigo-700 border-indigo-200" :
                local.exchange === "BSE" ? "bg-orange-50 text-orange-700 border-orange-200" :
                "bg-stone-50 text-stone-400 border-stone-200"
              }`}>
                {local.exchange || "—"}
              </div>
            </div>
          </div>
          <div>
            <label className="text-xs font-medium text-stone-500 mb-1 block">Company Name (optional)</label>
            <input
              value={local.name}
              onChange={(e) => update("name", e.target.value)}
              placeholder="e.g. Reliance Industries"
              className="w-full border border-stone-200 rounded-lg px-2.5 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400"
            />
          </div>

          {/* Price data */}
          <div>
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-semibold text-stone-600 uppercase tracking-wide">Position</span>
              <div className="flex items-center gap-3">
                <button
                  onClick={fetchLive}
                  disabled={fetching}
                  className="text-[11px] font-medium text-teal-600 flex items-center gap-1 disabled:opacity-50"
                >
                  {fetching ? <Loader2 size={12} className="animate-spin" /> : <RefreshCw size={12} />}
                  Fetch live data
                </button>
                <button onClick={() => setShowCalc(true)} className="text-[11px] font-medium text-teal-600 flex items-center gap-1">
                  <Calculator size={12} /> Avg calculator
                </button>
              </div>
            </div>
            {fetchMsg && (
              <div className={`text-[11px] px-2.5 py-2 rounded-lg mb-2 ${fetchMsg.ok ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-700"}`}>
                {fetchMsg.msg}
              </div>
            )}
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className="text-xs text-stone-500 mb-1 block">Quantity</label>
                <NumInput value={local.qty} onChange={(v) => update("qty", v)} placeholder="0" />
              </div>
              <div>
                <label className="text-xs text-stone-500 mb-1 block">Avg Buy Price</label>
                <NumInput value={local.avgBuyPrice} onChange={(v) => update("avgBuyPrice", v)} placeholder="0.00" prefix="₹" />
              </div>
              <div>
                <label className="text-xs text-stone-500 mb-1 block">Current Price</label>
                <NumInput value={local.currentPrice} onChange={(v) => update("currentPrice", v)} placeholder="0.00" prefix="₹" />
              </div>
              <div>
                <label className="text-xs text-stone-500 mb-1 block">High Since Buy</label>
                <NumInput value={local.highSinceBuy} onChange={(v) => update("highSinceBuy", v)} placeholder="0.00" prefix="₹" />
              </div>
            </div>
          </div>

          {/* Fundamentals */}
          <div>
            <span className="text-xs font-semibold text-stone-600 uppercase tracking-wide block mb-2">Fundamentals — last 4 quarters trend (%)</span>
            <div className="space-y-3">
              <div>
                <div className="text-xs font-medium text-stone-500 mb-1">EPS Growth</div>
                <div className="grid grid-cols-2 gap-3">
                  <NumInput value={local.fundamentals.epsQoQ} onChange={(v) => updateFund("epsQoQ", v)} placeholder="QoQ %" />
                  <NumInput value={local.fundamentals.epsYoY} onChange={(v) => updateFund("epsYoY", v)} placeholder="YoY %" />
                </div>
              </div>
              <div>
                <div className="text-xs font-medium text-stone-500 mb-1">Profit Margin Growth</div>
                <div className="grid grid-cols-2 gap-3">
                  <NumInput value={local.fundamentals.pmQoQ} onChange={(v) => updateFund("pmQoQ", v)} placeholder="QoQ %" />
                  <NumInput value={local.fundamentals.pmYoY} onChange={(v) => updateFund("pmYoY", v)} placeholder="YoY %" />
                </div>
              </div>
              <div>
                <div className="text-xs font-medium text-stone-500 mb-1">PAT Growth</div>
                <div className="grid grid-cols-2 gap-3">
                  <NumInput value={local.fundamentals.patQoQ} onChange={(v) => updateFund("patQoQ", v)} placeholder="QoQ %" />
                  <NumInput value={local.fundamentals.patYoY} onChange={(v) => updateFund("patYoY", v)} placeholder="YoY %" />
                </div>
              </div>
            </div>
            <p className="text-[11px] text-stone-400 mt-2">Enter % change figures from the company's last 4 quarterly reports (e.g. Screener.in, NSE/BSE filings). Positive = growth, negative = decline.</p>
          </div>

          {/* Live preview */}
          <div className="rounded-xl border border-stone-200 p-3.5 bg-stone-50">
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-semibold text-stone-500 uppercase tracking-wide">Live Signal</span>
              <DecisionBadge decision={evalResult.decision} />
            </div>
            <ul className="text-xs text-stone-500 space-y-0.5">
              {evalResult.reasons.map((r, i) => <li key={i}>• {r}</li>)}
            </ul>
          </div>

          <div>
            <label className="text-xs font-medium text-stone-500 mb-1 block">Notes (optional)</label>
            <textarea
              value={local.notes}
              onChange={(e) => update("notes", e.target.value)}
              rows={2}
              className="w-full border border-stone-200 rounded-lg px-2.5 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400 resize-none"
            />
          </div>

          <div className="flex gap-2 pt-1">
            {stock.code && (
              <button
                onClick={() => { onDelete(stock.id); onClose(); }}
                className="px-4 py-2.5 rounded-lg border border-rose-200 text-rose-600 text-sm font-medium flex items-center gap-1.5"
              >
                <Trash2 size={15} /> Delete
              </button>
            )}
            <button onClick={save} className="flex-1 py-2.5 rounded-lg bg-teal-600 text-white text-sm font-semibold">
              Save Holding
            </button>
          </div>
        </div>
      </div>

      {showCalc && (
        <AverageCalculator
          onClose={() => setShowCalc(false)}
          onApply={(res) => {
            setLocal((l) => ({ ...l, avgBuyPrice: res.avg.toFixed(2), qty: res.totalQty.toString() }));
            setShowCalc(false);
          }}
        />
      )}
    </div>
  );
}

// ---------- Import Panel ----------
function ImportPanel({ onClose, onImport }) {
  const [tab, setTab] = useState("paste");
  const [text, setText] = useState("");
  const [excelRows, setExcelRows] = useState(null);
  const [excelFileName, setExcelFileName] = useState("");
  const [excelError, setExcelError] = useState("");
  const fileRef = useRef(null);
  const excelRef = useRef(null);

  const handleCsvFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => setText(ev.target.result);
    reader.readAsText(file);
  };

  const handleExcelFile = (e) => {
    const file = e.target.files[0];
    if (!file) return;
    setExcelError("");
    setExcelFileName(file.name);
    const reader = new FileReader();
    reader.onload = (ev) => {
      try {
        const rows = parseExcel(ev.target.result);
        if (rows.length === 0) {
          setExcelError("No recognizable rows found. Make sure the first row has headers like code, qty, avg, current.");
          setExcelRows(null);
        } else {
          setExcelRows(rows);
        }
      } catch (err) {
        setExcelError("Couldn't read this file. Make sure it's a valid .xlsx or .xls file.");
        setExcelRows(null);
      }
    };
    reader.readAsArrayBuffer(file);
  };

  const doImport = () => {
    if (tab === "excel") {
      if (excelRows) onImport(excelRows);
    } else {
      const rows = parseCSV(text);
      onImport(rows);
    }
    onClose();
  };

  const canImport = tab === "excel" ? !!excelRows && excelRows.length > 0 : !!text.trim();

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-lg max-h-[85vh] overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white border-b border-stone-100 px-5 py-4 flex items-center justify-between">
          <div className="flex items-center gap-2">
            <FileSpreadsheet size={18} className="text-teal-600" />
            <h3 className="font-semibold text-stone-800">Import Holdings</h3>
          </div>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600"><X size={20} /></button>
        </div>
        <div className="p-5">
          <div className="flex gap-1 bg-stone-100 rounded-lg p-1 mb-4">
            <button onClick={() => setTab("paste")} className={`flex-1 py-1.5 rounded-md text-xs font-medium ${tab === "paste" ? "bg-white shadow-sm text-stone-800" : "text-stone-500"}`}>Paste</button>
            <button onClick={() => setTab("file")} className={`flex-1 py-1.5 rounded-md text-xs font-medium ${tab === "file" ? "bg-white shadow-sm text-stone-800" : "text-stone-500"}`}>Upload CSV</button>
            <button onClick={() => setTab("excel")} className={`flex-1 py-1.5 rounded-md text-xs font-medium ${tab === "excel" ? "bg-white shadow-sm text-stone-800" : "text-stone-500"}`}>Upload Excel</button>
          </div>

          {tab === "file" && (
            <div className="mb-3">
              <input ref={fileRef} type="file" accept=".csv,.txt" onChange={handleCsvFile} className="text-xs" />
            </div>
          )}

          {tab === "excel" ? (
            <div>
              <input ref={excelRef} type="file" accept=".xlsx,.xls" onChange={handleExcelFile} className="text-xs mb-3" />
              {excelError && (
                <div className="text-[11px] px-3 py-2 rounded-lg bg-rose-50 text-rose-700 mb-2">{excelError}</div>
              )}
              {excelRows && (
                <div className="border border-stone-200 rounded-lg p-3 max-h-48 overflow-y-auto">
                  <p className="text-xs font-medium text-stone-600 mb-2">{excelFileName} — {excelRows.length} row(s) found</p>
                  <div className="space-y-1">
                    {excelRows.slice(0, 8).map((r, i) => (
                      <div key={i} className="text-[11px] text-stone-500 flex gap-2">
                        <span className="font-semibold text-stone-700 w-20 shrink-0">{r.code}</span>
                        <span>{r.exchange}</span>
                        <span>qty {r.qty || "—"}</span>
                        <span>avg ₹{r.avgBuyPrice || "—"}</span>
                      </div>
                    ))}
                    {excelRows.length > 8 && <p className="text-[11px] text-stone-400">+ {excelRows.length - 8} more</p>}
                  </div>
                </div>
              )}
              <p className="text-[11px] text-stone-400 mt-2">
                First row should be headers. Recognized columns: code, name, qty, avg, current, high, epsQoQ, epsYoY, pmQoQ, pmYoY, patQoQ, patYoY.
              </p>
            </div>
          ) : (
            <>
              <textarea
                value={text}
                onChange={(e) => setText(e.target.value)}
                rows={8}
                placeholder={"code,name,qty,avg,current,high,epsqoq,epsyoy,pmqoq,pmyoy,patqoq,patyoy\nRELIANCE,Reliance Industries,10,2400,2650,2700,5,8,2,3,6,9"}
                className="w-full border border-stone-200 rounded-lg px-3 py-2 text-xs font-mono focus:outline-none focus:ring-2 focus:ring-teal-500/40 focus:border-teal-400 resize-none"
              />
              <p className="text-[11px] text-stone-400 mt-2">
                Header row required. Recognized columns: code, name, qty, avg, current, high, epsQoQ, epsYoY, pmQoQ, pmYoY, patQoQ, patYoY. Missing columns are fine — fill in later.
              </p>
            </>
          )}

          <div className="flex gap-2 mt-4">
            <button onClick={onClose} className="flex-1 py-2.5 rounded-lg border border-stone-200 text-stone-600 text-sm font-medium">Cancel</button>
            <button onClick={doImport} disabled={!canImport} className="flex-1 py-2.5 rounded-lg bg-teal-600 text-white text-sm font-medium disabled:opacity-40">
              Import
            </button>
          </div>
        </div>
      </div>
    </div>
  );
}

// ---------- New Stock Recommendations (screener) ----------
function RecommendationsPanel({ onClose, onAddStock }) {
  const [candidates, setCandidates] = useState([
    { id: 1, code: "", name: "", epsQoQ: "", epsYoY: "", pmQoQ: "", pmYoY: "", patQoQ: "", patYoY: "", currentPrice: "" },
  ]);

  const addCandidate = () => setCandidates((c) => [...c, { id: Date.now(), code: "", name: "", epsQoQ: "", epsYoY: "", pmQoQ: "", pmYoY: "", patQoQ: "", patYoY: "", currentPrice: "" }]);
  const removeCandidate = (id) => setCandidates((c) => c.filter((x) => x.id !== id));
  const updateCandidate = (id, field, val) => setCandidates((c) => c.map((x) => (x.id === id ? { ...x, [field]: val } : x)));

  const scored = useMemo(() => {
    return candidates.map((c) => {
      const flags = fundamentalFlags({ epsQoQ: c.epsQoQ, epsYoY: c.epsYoY, pmQoQ: c.pmQoQ, pmYoY: c.pmYoY, patQoQ: c.patQoQ, patYoY: c.patYoY });
      const allPositive = flags.eps.positive && flags.pm.positive && flags.pat.positive;
      const positiveCount = [flags.eps.positive, flags.pm.positive, flags.pat.positive].filter(Boolean).length;
      return { ...c, flags, allPositive, positiveCount };
    }).sort((a, b) => b.positiveCount - a.positiveCount);
  }, [candidates]);

  return (
    <div className="fixed inset-0 z-50 flex items-end sm:items-center justify-center bg-stone-900/40 backdrop-blur-sm">
      <div className="bg-white rounded-t-2xl sm:rounded-2xl w-full sm:max-w-lg max-h-[92vh] overflow-y-auto shadow-2xl">
        <div className="sticky top-0 bg-white border-b border-stone-100 px-5 py-4 flex items-center justify-between">
          <h3 className="font-semibold text-stone-800">Screen New Stocks</h3>
          <button onClick={onClose} className="text-stone-400 hover:text-stone-600"><X size={20} /></button>
        </div>
        <div className="p-5 space-y-4">
          <p className="text-xs text-stone-500">Enter fundamentals for stocks you're watching. Ranked by how many of the 3 fundamental signals are positive — the same test used for BUY MORE.</p>

          {scored.map((c) => (
            <div key={c.id} className="border border-stone-200 rounded-xl p-3.5 space-y-2.5">
              <div className="flex items-center justify-between gap-2">
                <input
                  value={c.code}
                  onChange={(e) => updateCandidate(c.id, "code", e.target.value.toUpperCase())}
                  placeholder="Code"
                  className="flex-1 border border-stone-200 rounded-lg px-2.5 py-1.5 text-sm font-medium uppercase focus:outline-none focus:ring-2 focus:ring-teal-500/40"
                />
                <span className={`text-[10px] font-semibold px-1.5 py-0.5 rounded ${
                  detectExchange(c.code) === "NSE" ? "bg-indigo-50 text-indigo-700" :
                  detectExchange(c.code) === "BSE" ? "bg-orange-50 text-orange-700" : "bg-stone-100 text-stone-400"
                }`}>{detectExchange(c.code) || "—"}</span>
                {candidates.length > 1 && (
                  <button onClick={() => removeCandidate(c.id)} className="text-stone-300 hover:text-rose-500"><Trash2 size={14} /></button>
                )}
              </div>
              <div className="grid grid-cols-3 gap-2">
                <NumInput value={c.epsQoQ} onChange={(v) => updateCandidate(c.id, "epsQoQ", v)} placeholder="EPS QoQ%" />
                <NumInput value={c.pmQoQ} onChange={(v) => updateCandidate(c.id, "pmQoQ", v)} placeholder="PM QoQ%" />
                <NumInput value={c.patQoQ} onChange={(v) => updateCandidate(c.id, "patQoQ", v)} placeholder="PAT QoQ%" />
                <NumInput value={c.epsYoY} onChange={(v) => updateCandidate(c.id, "epsYoY", v)} placeholder="EPS YoY%" />
                <NumInput value={c.pmYoY} onChange={(v) => updateCandidate(c.id, "pmYoY", v)} placeholder="PM YoY%" />
                <NumInput value={c.patYoY} onChange={(v) => updateCandidate(c.id, "patYoY", v)} placeholder="PAT YoY%" />
              </div>
              <div className="flex items-center justify-between pt-1">
                <div className="flex gap-1">
                  <FundamentalPill label="EPS" hasData={c.flags.eps.hasData} positive={c.flags.eps.positive} negative={c.flags.eps.negative} />
                  <FundamentalPill label="PM" hasData={c.flags.pm.hasData} positive={c.flags.pm.positive} negative={c.flags.pm.negative} />
                  <FundamentalPill label="PAT" hasData={c.flags.pat.hasData} positive={c.flags.pat.positive} negative={c.flags.pat.negative} />
                </div>
                {c.allPositive && c.code && (
                  <button
                    onClick={() => { onAddStock(c.code); onClose(); }}
                    className="text-[11px] font-semibold text-teal-600 flex items-center gap-1"
                  >
                    <Plus size={12} /> Add to portfolio
                  </button>
                )}
              </div>
            </div>
          ))}

          <button onClick={addCandidate} className="w-full py-2.5 rounded-lg border border-dashed border-stone-300 text-stone-500 text-xs font-medium flex items-center justify-center gap-1.5">
            <Plus size={14} /> Add candidate
          </button>
        </div>
      </div>
    </div>
  );
}

// ---------- Stock Card ----------
function StockCard({ stock, onClick }) {
  const result = useMemo(() => evaluateStock(stock), [stock]);
  const price = parseFloat(stock.currentPrice);
  const avg = parseFloat(stock.avgBuyPrice);
  const qty = parseFloat(stock.qty);
  const gainPct = result.gainFromAvg;
  const pnl = !isNaN(price) && !isNaN(avg) && !isNaN(qty) ? (price - avg) * qty : null;

  return (
    <button onClick={onClick} className="w-full text-left bg-white rounded-xl border border-stone-200 p-4 hover:border-stone-300 hover:shadow-sm transition-all active:scale-[0.99]">
      <div className="flex items-start justify-between mb-2">
        <div>
          <div className="flex items-center gap-2">
            <span className="font-semibold text-stone-800 text-sm">{stock.code || "—"}</span>
            {stock.exchange && (
              <span className={`text-[9px] font-bold px-1.5 py-0.5 rounded ${
                stock.exchange === "NSE" ? "bg-indigo-50 text-indigo-700" : "bg-orange-50 text-orange-700"
              }`}>{stock.exchange}</span>
            )}
          </div>
          {stock.name && <div className="text-[11px] text-stone-400 mt-0.5">{stock.name}</div>}
        </div>
        <DecisionBadge decision={result.decision} />
      </div>

      <div className="grid grid-cols-3 gap-2 text-xs mb-2.5">
        <div>
          <div className="text-stone-400 text-[10px]">Qty</div>
          <div className="font-medium text-stone-700">{stock.qty || "—"}</div>
        </div>
        <div>
          <div className="text-stone-400 text-[10px]">Avg / LTP</div>
          <div className="font-medium text-stone-700">₹{stock.avgBuyPrice || "—"} / ₹{stock.currentPrice || "—"}</div>
        </div>
        <div>
          <div className="text-stone-400 text-[10px]">P&L</div>
          <div className={`font-semibold ${pnl > 0 ? "text-emerald-600" : pnl < 0 ? "text-rose-600" : "text-stone-700"}`}>
            {pnl !== null ? `₹${pnl.toFixed(0)}` : "—"} {gainPct !== null && <span className="font-normal text-[10px]">({gainPct >= 0 ? "+" : ""}{gainPct.toFixed(1)}%)</span>}
          </div>
        </div>
      </div>

      <div className="flex gap-1">
        <FundamentalPill label="EPS" hasData={result.flags.eps.hasData} positive={result.flags.eps.positive} negative={result.flags.eps.negative} />
        <FundamentalPill label="PM" hasData={result.flags.pm.hasData} positive={result.flags.pm.positive} negative={result.flags.pm.negative} />
        <FundamentalPill label="PAT" hasData={result.flags.pat.hasData} positive={result.flags.pat.positive} negative={result.flags.pat.negative} />
      </div>
    </button>
  );
}

// ---------- Main App ----------
export default function PortfolioApp() {
  const [stocks, setStocks] = useState([]);
  const [editingStock, setEditingStock] = useState(null);
  const [showImport, setShowImport] = useState(false);
  const [showRecs, setShowRecs] = useState(false);
  const [showSettings, setShowSettings] = useState(false);
  const [filter, setFilter] = useState("ALL");
  const [search, setSearch] = useState("");
  const [loaded, setLoaded] = useState(false);
  const [apiKey, setApiKey] = useState("");
  const [refreshing, setRefreshing] = useState(false);
  const [refreshProgress, setRefreshProgress] = useState({ done: 0, total: 0 });
  const [refreshSummary, setRefreshSummary] = useState(null);

  useEffect(() => {
    setStocks(loadPortfolio());
    setApiKey(loadApiKey());
    setLoaded(true);
  }, []);

  useEffect(() => {
    if (loaded) savePortfolio(stocks);
  }, [stocks, loaded]);

  const upsertStock = (stock) => {
    setStocks((prev) => {
      const exists = prev.find((s) => s.id === stock.id);
      if (exists) return prev.map((s) => (s.id === stock.id ? stock : s));
      return [...prev, stock];
    });
  };
  const deleteStock = (id) => setStocks((prev) => prev.filter((s) => s.id !== id));

  const importStocks = (rows) => setStocks((prev) => [...prev, ...rows]);

  // Refreshes CMP / high / fundamentals for every holding, one request at a
  // time (sequential, with a short pause) to stay within free-tier rate limits.
  const refreshAll = async () => {
    if (!apiKey) {
      setShowSettings(true);
      return;
    }
    if (stocks.length === 0) return;
    setRefreshing(true);
    setRefreshSummary(null);
    setRefreshProgress({ done: 0, total: stocks.length });
    let okCount = 0;
    let failCount = 0;
    const updated = [...stocks];
    for (let i = 0; i < updated.length; i++) {
      const s = updated[i];
      if (s.code) {
        try {
          const result = await fetchStockData(s.code, apiKey);
          updated[i] = {
            ...s,
            currentPrice: result.currentPrice || s.currentPrice,
            highSinceBuy: result.high52 || s.highSinceBuy,
            name: result.name || s.name,
          };
          okCount++;
        } catch {
          failCount++;
        }
      }
      setRefreshProgress({ done: i + 1, total: updated.length });
      setStocks([...updated]);
      // small pause between calls to be gentle on free-tier rate limits
      await new Promise((r) => setTimeout(r, 350));
    }
    setRefreshSummary({ okCount, failCount, total: stocks.length });
    setRefreshing(false);
  };

  const addFromRecommendation = (code) => {
    const s = emptyStock();
    s.code = code.toUpperCase();
    s.exchange = detectExchange(code);
    setEditingStock(s);
  };

  const summary = useMemo(() => {
    const results = stocks.map((s) => ({ stock: s, result: evaluateStock(s) }));
    const counts = { SELL: 0, "BUY MORE": 0, HOLD: 0 };
    results.forEach((r) => counts[r.result.decision]++);
    const totalValue = stocks.reduce((sum, s) => {
      const p = parseFloat(s.currentPrice), q = parseFloat(s.qty);
      return sum + (!isNaN(p) && !isNaN(q) ? p * q : 0);
    }, 0);
    const totalCost = stocks.reduce((sum, s) => {
      const a = parseFloat(s.avgBuyPrice), q = parseFloat(s.qty);
      return sum + (!isNaN(a) && !isNaN(q) ? a * q : 0);
    }, 0);
    return { counts, totalValue, totalCost, totalPnl: totalValue - totalCost, results };
  }, [stocks]);

  const filtered = useMemo(() => {
    return summary.results
      .filter((r) => filter === "ALL" || r.result.decision === filter)
      .filter((r) => !search || r.stock.code.toLowerCase().includes(search.toLowerCase()) || r.stock.name.toLowerCase().includes(search.toLowerCase()))
      .map((r) => r.stock);
  }, [summary, filter, search]);

  const exportCSV = () => {
    const headers = ["code", "exchange", "name", "qty", "avg", "current", "high", "epsQoQ", "epsYoY", "pmQoQ", "pmYoY", "patQoQ", "patYoY"];
    const rows = stocks.map((s) => [
      s.code, s.exchange, s.name, s.qty, s.avgBuyPrice, s.currentPrice, s.highSinceBuy,
      s.fundamentals.epsQoQ, s.fundamentals.epsYoY, s.fundamentals.pmQoQ, s.fundamentals.pmYoY, s.fundamentals.patQoQ, s.fundamentals.patYoY,
    ]);
    const csv = [headers.join(","), ...rows.map((r) => r.join(","))].join("\n");
    const blob = new Blob([csv], { type: "text/csv" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `portfolio-${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  };

  return (
    <div className="min-h-screen bg-stone-50 font-sans" style={{ fontFamily: "'Inter', -apple-system, sans-serif" }}>
      {/* Header */}
      <div className="bg-gradient-to-br from-teal-700 to-teal-900 text-white px-5 pt-6 pb-8 rounded-b-3xl shadow-lg">
        <div className="flex items-center justify-between mb-1">
          <h1 className="text-xl font-bold tracking-tight">Portfolio Signals</h1>
          <div className="flex items-center gap-1.5">
            <button
              onClick={refreshAll}
              disabled={refreshing || stocks.length === 0}
              className="p-2 bg-white/10 rounded-lg hover:bg-white/20 disabled:opacity-40"
              title="Refresh all live data"
            >
              {refreshing ? <Loader2 size={16} className="animate-spin" /> : <RefreshCw size={16} />}
            </button>
            <button onClick={() => setShowSettings(true)} className="p-2 bg-white/10 rounded-lg hover:bg-white/20 relative">
              {apiKey ? <Wifi size={16} /> : <WifiOff size={16} />}
            </button>
            <button onClick={() => setShowRecs(true)} className="p-2 bg-white/10 rounded-lg hover:bg-white/20">
              <Search size={16} />
            </button>
          </div>
        </div>
        <p className="text-teal-100 text-xs">
          {refreshing
            ? `Refreshing ${refreshProgress.done}/${refreshProgress.total}...`
            : refreshSummary
            ? `Refreshed: ${refreshSummary.okCount}/${refreshSummary.total} updated${refreshSummary.failCount ? `, ${refreshSummary.failCount} failed` : ""}`
            : <>NSE · BSE — auto-detected from code {apiKey ? "· Live data connected" : "· Tap the wifi icon to connect live data"}</>}
        </p>

        <div className="grid grid-cols-3 gap-2 mt-5">
          <div className="bg-white/10 backdrop-blur rounded-xl p-3">
            <div className="text-teal-200 text-[10px] uppercase tracking-wide">Value</div>
            <div className="font-bold text-lg">₹{summary.totalValue ? summary.totalValue.toLocaleString("en-IN", { maximumFractionDigits: 0 }) : "0"}</div>
          </div>
          <div className="bg-white/10 backdrop-blur rounded-xl p-3">
            <div className="text-teal-200 text-[10px] uppercase tracking-wide">P&L</div>
            <div className={`font-bold text-lg ${summary.totalPnl >= 0 ? "text-emerald-300" : "text-rose-300"}`}>
              {summary.totalPnl >= 0 ? "+" : ""}₹{Math.abs(summary.totalPnl).toLocaleString("en-IN", { maximumFractionDigits: 0 })}
            </div>
          </div>
          <div className="bg-white/10 backdrop-blur rounded-xl p-3">
            <div className="text-teal-200 text-[10px] uppercase tracking-wide">Holdings</div>
            <div className="font-bold text-lg">{stocks.length}</div>
          </div>
        </div>
      </div>

      {/* Filter pills */}
      <div className="px-5 -mt-4 mb-3">
        <div className="bg-white rounded-2xl shadow-sm border border-stone-100 p-1.5 flex gap-1">
          {["ALL", "BUY MORE", "HOLD", "SELL"].map((f) => (
            <button
              key={f}
              onClick={() => setFilter(f)}
              className={`flex-1 py-2 rounded-xl text-[11px] font-semibold transition-colors ${
                filter === f ? "bg-stone-800 text-white" : "text-stone-500 hover:bg-stone-50"
              }`}
            >
              {f}{f !== "ALL" && ` (${summary.counts[f]})`}
            </button>
          ))}
        </div>
      </div>

      {/* Search + actions */}
      <div className="px-5 mb-3 flex gap-2">
        <div className="relative flex-1">
          <Search size={14} className="absolute left-3 top-1/2 -translate-y-1/2 text-stone-300" />
          <input
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search holdings..."
            className="w-full bg-white border border-stone-200 rounded-xl pl-8 pr-3 py-2.5 text-xs focus:outline-none focus:ring-2 focus:ring-teal-500/30"
          />
        </div>
        <button onClick={() => setShowImport(true)} className="bg-white border border-stone-200 rounded-xl px-3 py-2.5 text-stone-600">
          <Upload size={15} />
        </button>
        <button onClick={exportCSV} className="bg-white border border-stone-200 rounded-xl px-3 py-2.5 text-stone-600">
          <Download size={15} />
        </button>
      </div>

      {/* Stock list */}
      <div className="px-5 pb-28 space-y-2.5">
        {filtered.length === 0 && (
          <div className="text-center py-16">
            <div className="text-stone-300 mb-2"><FileSpreadsheet size={32} className="mx-auto" /></div>
            <p className="text-stone-400 text-sm font-medium">{stocks.length === 0 ? "No holdings yet" : "No matches"}</p>
            <p className="text-stone-300 text-xs mt-1">{stocks.length === 0 ? "Add a stock or import your portfolio" : "Try a different filter or search"}</p>
          </div>
        )}
        {filtered.map((stock) => (
          <StockCard key={stock.id} stock={stock} onClick={() => setEditingStock(stock)} />
        ))}
      </div>

      {/* FAB */}
      <button
        onClick={() => setEditingStock(emptyStock())}
        className="fixed bottom-6 right-5 bg-teal-600 hover:bg-teal-700 text-white rounded-full w-14 h-14 shadow-xl flex items-center justify-center z-30 active:scale-95 transition-transform"
      >
        <Plus size={24} />
      </button>

      {/* Rules footer note */}
      <div className="fixed bottom-0 left-0 right-0 bg-white/95 backdrop-blur border-t border-stone-100 px-5 py-2 text-center z-20">
        <p className="text-[10px] text-stone-400">SELL: −{SELL_DROP_PCT}% from high or 2+ fundamentals declining · BUY MORE: +{BUY_MORE_GAIN_PCT}% & all fundamentals positive</p>
      </div>

      {editingStock && (
        <StockEditor
          stock={editingStock}
          onChange={upsertStock}
          onDelete={deleteStock}
          onClose={() => setEditingStock(null)}
          apiKey={apiKey}
          onNeedApiKey={() => setShowSettings(true)}
        />
      )}
      {showImport && <ImportPanel onClose={() => setShowImport(false)} onImport={importStocks} />}
      {showRecs && <RecommendationsPanel onClose={() => setShowRecs(false)} onAddStock={addFromRecommendation} />}
      {showSettings && (
        <SettingsPanel
          apiKey={apiKey}
          onSave={(key) => { setApiKey(key); saveApiKey(key); }}
          onClose={() => setShowSettings(false)}
        />
      )}
    </div>
  );
}
