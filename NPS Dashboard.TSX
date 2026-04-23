import { useState, useMemo, useCallback, useRef, useEffect } from "react";
import * as XLSX from "xlsx";

// ─── NPS helpers ──────────────────────────────────────────────────────────────
const npsCategory = (score) => {
  const n = Number(score);
  if (isNaN(n)) return null;
  if (n >= 9) return "Promoter";
  if (n >= 7) return "Passive";
  return "Detractor";
};
const calcNPS = (rows) => {
  const valid = rows.filter((r) => r.npsCategory !== null);
  if (!valid.length) return null;
  const p = valid.filter((r) => r.npsCategory === "Promoter").length;
  const d = valid.filter((r) => r.npsCategory === "Detractor").length;
  return Math.round(((p - d) / valid.length) * 100);
};
const npsColor = (score) => {
  if (score === null || score === undefined) return "#64748b";
  if (score >= 50) return "#10b981";
  if (score >= 0) return "#f59e0b";
  return "#ef4444";
};
const npsLabel = (score) => {
  if (score === null || score === undefined) return "—";
  if (score >= 50) return "Excellent";
  if (score >= 0) return "Needs Work";
  return "Critical";
};

// ─── Parsers ──────────────────────────────────────────────────────────────────
function parseExpressReport(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  let headerIdx = raw.findIndex((row) =>
    row.some((cell) => String(cell || "").trim() === "Survey Type")
  );
  if (headerIdx === -1) headerIdx = 2;
  const headers = raw[headerIdx].map((h) => String(h || "").trim());
  const getCol = (name) => headers.findIndex((h) => h === name);
  const idxSurveyType = getCol("Survey Type");
  const idxLTR = getCol("Likelihood to Recommend");
  const idxDate = getCol("Response Date (Local)");
  const idxCountry = getCol("Country/Location");
  const idxTxnId = getCol("CS Transaction ID");
  const idxEmpId = getCol("CS Contact Center Employee ID");
  const idxImpact = getCol("CS Chat Impact Factor");
  const idxResolution = getCol("Resolution (Yes/No)");
  const idxConsultation = getCol("Consultation Category");
  const idxDivision = getCol("Division");
  const idxReasonScore = getCol("Reason for Score");
  const seenTxn = new Set();
  const rows = [];
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row || row.every((c) => c === null)) continue;
    const surveyType = String(row[idxSurveyType] || "").trim();
    if (surveyType !== "Chat") continue;
    const txnId = String(row[idxTxnId] || "").trim();
    if (!txnId || seenTxn.has(txnId)) continue;
    seenTxn.add(txnId);
    const dateRaw = row[idxDate];
    let dateObj = null;
    if (dateRaw instanceof Date) dateObj = dateRaw;
    else if (typeof dateRaw === "string" || typeof dateRaw === "number")
      dateObj = new Date(dateRaw);
    const monthKey = dateObj
      ? `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(
          2,
          "0"
        )}`
      : "Unknown";
    const monthLabel = dateObj
      ? dateObj.toLocaleDateString("en-GB", { month: "short", year: "numeric" })
      : "Unknown";
    const dayKey = dateObj
      ? `${dateObj.getFullYear()}-${String(dateObj.getMonth() + 1).padStart(
          2,
          "0"
        )}-${String(dateObj.getDate()).padStart(2, "0")}`
      : "Unknown";
    const dayLabel = dateObj
      ? dateObj.toLocaleDateString("en-GB", {
          day: "2-digit",
          month: "short",
          year: "numeric",
        })
      : "Unknown";
    const ltr = Number(row[idxLTR]);
    const empId = String(row[idxEmpId] || "")
      .trim()
      .replace(/\.0$/, "");
    rows.push({
      txnId,
      ltr: isNaN(ltr) ? null : ltr,
      npsCategory: npsCategory(ltr),
      dateObj,
      monthKey,
      monthLabel,
      dayKey,
      dayLabel,
      country: String(row[idxCountry] || "").trim() || "Unknown",
      empId,
      impact: String(row[idxImpact] || "").trim() || null,
      resolution: String(row[idxResolution] || "").trim() || null,
      consultation: String(row[idxConsultation] || "").trim() || null,
      division: String(row[idxDivision] || "").trim() || null,
      reasonScore: String(row[idxReasonScore] || "").trim() || null,
    });
  }
  return rows;
}

function parseAgentsFile(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array", cellDates: true });
  const sheetName =
    wb.SheetNames.find((n) => n.toLowerCase().includes("agents")) ||
    wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const raw = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  let headerIdx = raw.findIndex((row) =>
    row.some((cell) => String(cell || "").trim() === "BP Number")
  );
  if (headerIdx === -1) headerIdx = 0;
  const headers = raw[headerIdx].map((h) => String(h || "").trim());
  const idxBP = headers.findIndex((h) => h === "BP Number");
  const idxName = headers.findIndex((h) => h === "Name");
  const idxLOB = headers.findIndex((h) => h === "LOB");
  const idxWave = headers.findIndex((h) => h === "Wave");
  const idxTL = headers.findIndex((h) => h.includes("1st Line Manager"));
  const map = {};
  for (let i = headerIdx + 1; i < raw.length; i++) {
    const row = raw[i];
    if (!row) continue;
    const bp = String(row[idxBP] || "")
      .trim()
      .replace(/\.0$/, "");
    if (!bp || bp === "-" || bp === "null") continue;
    map[bp] = {
      name: String(row[idxName] || "").trim() || null,
      lob: String(row[idxLOB] || "").trim() || null,
      wave: String(row[idxWave] || "").trim() || null,
      tl: String(row[idxTL] || "").trim() || null,
    };
  }
  return map;
}

// ─── Theme definitions ────────────────────────────────────────────────────────
const THEMES = {
  dark: {
    bg: "#060e1c",
    surface: "#0a1628",
    surface2: "#0f1e35",
    border: "#1e293b",
    borderSoft: "#0f172a",
    text: "#e2e8f0",
    textMuted: "#475569",
    textFaint: "#1e293b",
    textSub: "#64748b",
    rowOdd: "#0a1628",
    rowEven: "#060e1c",
    headerBg: "#060e1c",
    inputBg: "#060e1c",
    accent: "#38bdf8",
    overallBg: "#0a1f38",
    stickyOdd: "#0a1628",
    stickyEven: "#060e1c",
    uploadBorder: "#1e293b",
    uploadBg: "#0a1628",
    pillBg: "#0ea5e920",
    toggleBg: "#1e293b",
  },
  light: {
    bg: "#f1f5f9",
    surface: "#ffffff",
    surface2: "#f8fafc",
    border: "#e2e8f0",
    borderSoft: "#f1f5f9",
    text: "#0f172a",
    textMuted: "#64748b",
    textFaint: "#cbd5e1",
    textSub: "#94a3b8",
    rowOdd: "#ffffff",
    rowEven: "#f8fafc",
    headerBg: "#f8fafc",
    inputBg: "#ffffff",
    accent: "#0ea5e9",
    overallBg: "#eff6ff",
    stickyOdd: "#ffffff",
    stickyEven: "#f8fafc",
    uploadBorder: "#e2e8f0",
    uploadBg: "#f8fafc",
    pillBg: "#0ea5e910",
    toggleBg: "#e2e8f0",
  },
};

// ─── Shared UI components ─────────────────────────────────────────────────────
const NPSPill = ({ score }) => {
  if (score === null || score === undefined)
    return <span style={{ color: "#64748b", fontSize: 13 }}>—</span>;
  const c = npsColor(score);
  return (
    <span
      style={{
        background: c + "20",
        color: c,
        border: `1.5px solid ${c}55`,
        borderRadius: 20,
        padding: "3px 13px",
        fontWeight: 800,
        fontSize: 13,
        display: "inline-block",
        minWidth: 56,
        textAlign: "center",
        fontFamily: "'DM Mono', monospace",
      }}
    >
      {score > 0 ? `+${score}` : score}
    </span>
  );
};

const MiniBar = ({ promoters, passives, detractors, total }) => {
  if (!total) return null;
  return (
    <div
      style={{
        display: "flex",
        height: 7,
        borderRadius: 4,
        overflow: "hidden",
        width: "100%",
        gap: 1,
      }}
    >
      <div
        style={{
          width: `${(promoters / total) * 100}%`,
          background: "#10b981",
        }}
      />
      <div
        style={{ width: `${(passives / total) * 100}%`, background: "#f59e0b" }}
      />
      <div
        style={{
          width: `${(detractors / total) * 100}%`,
          background: "#ef4444",
        }}
      />
    </div>
  );
};

const Badge = ({ children, color = "#334155" }) => (
  <span
    style={{
      background: color + "20",
      color,
      border: `1px solid ${color}40`,
      borderRadius: 6,
      padding: "2px 9px",
      fontSize: 11,
      fontWeight: 700,
      letterSpacing: "0.03em",
      whiteSpace: "nowrap",
    }}
  >
    {children}
  </span>
);

// ─── Agent Tooltip ────────────────────────────────────────────────────────────
function AgentTooltip({ agent, onClose, theme }) {
  const T = THEMES[theme];
  const [copiedDsat, setCopiedDsat] = useState(false);
  const [copiedPassive, setCopiedPassive] = useState(false);
  const ref = useRef();

  useEffect(() => {
    const handler = (e) => {
      if (ref.current && !ref.current.contains(e.target)) onClose();
    };
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [onClose]);

  const dsatRows = agent.rows.filter((r) => r.npsCategory === "Detractor");
  const passiveRows = agent.rows.filter((r) => r.npsCategory === "Passive");

  const formatForCopy = (rows, label) => {
    if (!rows.length) return `No ${label}s`;
    const header = `Agent: ${agent.name || agent.empId} | ${label}s (${
      rows.length
    })\n${"─".repeat(50)}`;
    const lines = rows.map(
      (r, i) =>
        `${i + 1}. TXN: ${r.txnId} | Score: ${r.ltr} | Date: ${
          r.dayLabel
        } | Country: ${r.country}${
          r.reasonScore ? ` | Reason: ${r.reasonScore}` : ""
        }`
    );
    return [header, ...lines].join("\n");
  };

  const copy = (rows, label, setCopied) => {
    const text = formatForCopy(rows, label);
    navigator.clipboard
      .writeText(text)
      .then(() => {
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      })
      .catch(() => {
        // fallback
        const ta = document.createElement("textarea");
        ta.value = text;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand("copy");
        document.body.removeChild(ta);
        setCopied(true);
        setTimeout(() => setCopied(false), 2000);
      });
  };

  const nps = calcNPS(agent.rows);

  return (
    <div
      ref={ref}
      style={{
        position: "absolute",
        zIndex: 9999,
        top: "110%",
        left: 0,
        width: 360,
        borderRadius: 14,
        background: T.surface,
        border: `1.5px solid ${T.border}`,
        boxShadow:
          theme === "dark"
            ? "0 20px 60px #000000cc, 0 0 0 1px #1e293b"
            : "0 20px 60px #00000025, 0 0 0 1px #e2e8f0",
        overflow: "hidden",
      }}
    >
      {/* Header */}
      <div
        style={{
          padding: "14px 16px 10px",
          borderBottom: `1px solid ${T.border}`,
          background: theme === "dark" ? "#0a1f38" : "#eff6ff",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
          }}
        >
          <div>
            <div style={{ fontWeight: 800, fontSize: 14, color: T.text }}>
              {agent.name || `⚠ ${agent.empId}`}
            </div>
            <div style={{ fontSize: 11, color: T.textMuted, marginTop: 1 }}>
              {agent.lob && (
                <span style={{ marginRight: 8 }}>🏢 {agent.lob}</span>
              )}
              {agent.wave && (
                <span style={{ marginRight: 8 }}>🌊 Wave {agent.wave}</span>
              )}
              {agent.tl && <span>👥 {agent.tl}</span>}
            </div>
          </div>
          <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
            <NPSPill score={nps} />
            <button
              onClick={onClose}
              style={{
                background: "none",
                border: "none",
                cursor: "pointer",
                color: T.textMuted,
                fontSize: 16,
                lineHeight: 1,
                padding: 2,
              }}
            >
              ✕
            </button>
          </div>
        </div>
        {/* Mini stats */}
        <div style={{ display: "flex", gap: 14, marginTop: 10 }}>
          {[
            { label: "Total", val: agent.total, color: T.textSub },
            { label: "Promoters", val: agent.promoters, color: "#10b981" },
            { label: "Passives", val: agent.passives, color: "#f59e0b" },
            { label: "Detractors", val: agent.detractors, color: "#ef4444" },
          ].map(({ label, val, color }) => (
            <div key={label} style={{ textAlign: "center" }}>
              <div
                style={{
                  fontFamily: "'DM Mono', monospace",
                  fontWeight: 800,
                  fontSize: 16,
                  color,
                }}
              >
                {val}
              </div>
              <div
                style={{
                  fontSize: 9,
                  color: T.textMuted,
                  textTransform: "uppercase",
                  letterSpacing: "0.05em",
                }}
              >
                {label}
              </div>
            </div>
          ))}
          <div style={{ textAlign: "center", marginLeft: "auto" }}>
            <div
              style={{
                fontFamily: "'DM Mono', monospace",
                fontWeight: 800,
                fontSize: 16,
                color: T.accent,
              }}
            >
              {agent.avgLTR || "—"}
            </div>
            <div
              style={{
                fontSize: 9,
                color: T.textMuted,
                textTransform: "uppercase",
                letterSpacing: "0.05em",
              }}
            >
              Avg LTR
            </div>
          </div>
        </div>
        <div style={{ marginTop: 8 }}>
          <MiniBar
            promoters={agent.promoters}
            passives={agent.passives}
            detractors={agent.detractors}
            total={agent.total}
          />
        </div>
      </div>

      {/* Copy buttons */}
      <div
        style={{
          padding: "12px 16px",
          display: "flex",
          flexDirection: "column",
          gap: 8,
        }}
      >
        <div
          style={{
            fontSize: 10,
            fontWeight: 800,
            color: T.textMuted,
            textTransform: "uppercase",
            letterSpacing: "0.07em",
            marginBottom: 2,
          }}
        >
          Copy to Clipboard
        </div>

        {/* DSAT copy */}
        <button
          onClick={() => copy(dsatRows, "Detractor", setCopiedDsat)}
          disabled={!dsatRows.length}
          style={{
            width: "100%",
            padding: "10px 14px",
            borderRadius: 9,
            border: `1.5px solid ${dsatRows.length ? "#ef444450" : T.border}`,
            background: copiedDsat
              ? "#ef444420"
              : dsatRows.length
              ? "#ef444412"
              : "transparent",
            color: dsatRows.length
              ? copiedDsat
                ? "#ef4444"
                : "#f87171"
              : T.textFaint,
            cursor: dsatRows.length ? "pointer" : "not-allowed",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            fontFamily: "inherit",
            fontWeight: 700,
            fontSize: 13,
            transition: "all 0.15s",
          }}
        >
          <span>🔴 Copy DSATs ({dsatRows.length})</span>
          <span style={{ fontSize: 11, fontWeight: 600, opacity: 0.8 }}>
            {copiedDsat ? "✓ Copied!" : "TXN IDs + Scores + Dates"}
          </span>
        </button>

        {/* Passive copy */}
        <button
          onClick={() => copy(passiveRows, "Passive", setCopiedPassive)}
          disabled={!passiveRows.length}
          style={{
            width: "100%",
            padding: "10px 14px",
            borderRadius: 9,
            border: `1.5px solid ${
              passiveRows.length ? "#f59e0b50" : T.border
            }`,
            background: copiedPassive
              ? "#f59e0b20"
              : passiveRows.length
              ? "#f59e0b12"
              : "transparent",
            color: passiveRows.length
              ? copiedPassive
                ? "#f59e0b"
                : "#fbbf24"
              : T.textFaint,
            cursor: passiveRows.length ? "pointer" : "not-allowed",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            fontFamily: "inherit",
            fontWeight: 700,
            fontSize: 13,
            transition: "all 0.15s",
          }}
        >
          <span>🟡 Copy Passives ({passiveRows.length})</span>
          <span style={{ fontSize: 11, fontWeight: 600, opacity: 0.8 }}>
            {copiedPassive ? "✓ Copied!" : "TXN IDs + Scores + Dates"}
          </span>
        </button>

        {/* Preview of most recent DSAT */}
        {dsatRows.length > 0 && (
          <div
            style={{
              marginTop: 2,
              padding: "8px 10px",
              borderRadius: 8,
              background: T.surface2,
              border: `1px solid ${T.border}`,
            }}
          >
            <div
              style={{
                fontSize: 10,
                color: T.textMuted,
                marginBottom: 4,
                fontWeight: 700,
                textTransform: "uppercase",
                letterSpacing: "0.05em",
              }}
            >
              Latest DSAT
            </div>
            {dsatRows.slice(0, 2).map((r) => (
              <div
                key={r.txnId}
                style={{
                  fontSize: 11,
                  color: T.textSub,
                  marginBottom: 2,
                  display: "flex",
                  gap: 8,
                  alignItems: "center",
                }}
              >
                <span style={{ color: "#ef4444", fontWeight: 700 }}>
                  #{r.ltr}
                </span>
                <span style={{ color: T.textMuted }}>{r.dayLabel}</span>
                <span
                  style={{
                    color: T.textFaint,
                    fontFamily: "'DM Mono', monospace",
                    fontSize: 10,
                  }}
                >
                  {r.txnId}
                </span>
                {r.reasonScore && (
                  <span
                    style={{
                      color: T.textMuted,
                      fontStyle: "italic",
                      fontSize: 10,
                      overflow: "hidden",
                      textOverflow: "ellipsis",
                      whiteSpace: "nowrap",
                      maxWidth: 120,
                    }}
                  >
                    "{r.reasonScore}"
                  </span>
                )}
              </div>
            ))}
            {dsatRows.length > 2 && (
              <div style={{ fontSize: 10, color: T.textFaint }}>
                +{dsatRows.length - 2} more...
              </div>
            )}
          </div>
        )}
      </div>
    </div>
  );
}

// ─── Upload Zone ──────────────────────────────────────────────────────────────
function UploadZone({ label, icon, onFile, fileName, color, hint, theme }) {
  const T = THEMES[theme];
  const ref = useRef();
  const [drag, setDrag] = useState(false);
  const handle = (file) => {
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (e) => onFile(e.target.result, file.name);
    reader.readAsArrayBuffer(file);
  };
  return (
    <div
      onClick={() => ref.current.click()}
      onDragOver={(e) => {
        e.preventDefault();
        setDrag(true);
      }}
      onDragLeave={() => setDrag(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDrag(false);
        handle(e.dataTransfer.files[0]);
      }}
      style={{
        border: `2px dashed ${
          drag ? color : fileName ? color : T.uploadBorder
        }`,
        borderRadius: 12,
        padding: "18px 20px",
        cursor: "pointer",
        background: drag ? color + "12" : fileName ? color + "0a" : T.uploadBg,
        transition: "all 0.2s",
        textAlign: "center",
        flex: 1,
        minWidth: 220,
      }}
    >
      <input
        ref={ref}
        type="file"
        accept=".xlsx,.xls"
        style={{ display: "none" }}
        onChange={(e) => handle(e.target.files[0])}
      />
      <div style={{ fontSize: 24, marginBottom: 5 }}>{icon}</div>
      <div
        style={{
          fontWeight: 700,
          color: fileName ? color : T.textMuted,
          fontSize: 13,
        }}
      >
        {fileName || label}
      </div>
      {hint && !fileName && (
        <div style={{ fontSize: 11, color: T.textFaint, marginTop: 3 }}>
          {hint}
        </div>
      )}
      {fileName && (
        <div style={{ fontSize: 11, color: color, marginTop: 3 }}>
          ✓ Loaded — click to replace
        </div>
      )}
    </div>
  );
}

// ─── KPI Card ─────────────────────────────────────────────────────────────────
const KpiCard = ({ label, value, sub, color, icon, theme }) => {
  const T = THEMES[theme];
  return (
    <div
      style={{
        background: T.surface,
        borderRadius: 14,
        padding: "16px 18px",
        border: `1px solid ${color}25`,
        flex: 1,
        minWidth: 120,
        boxShadow:
          theme === "dark" ? `0 0 20px ${color}10` : `0 2px 8px ${color}15`,
      }}
    >
      <div style={{ fontSize: 18, marginBottom: 4 }}>{icon}</div>
      <div
        style={{
          fontFamily: "'DM Mono', monospace",
          fontWeight: 800,
          fontSize: 26,
          color,
          lineHeight: 1,
        }}
      >
        {value}
      </div>
      <div
        style={{
          fontSize: 10,
          color: T.textMuted,
          fontWeight: 700,
          marginTop: 3,
          textTransform: "uppercase",
          letterSpacing: "0.06em",
        }}
      >
        {label}
      </div>
      {sub && (
        <div style={{ fontSize: 11, color: T.textSub, marginTop: 2 }}>
          {sub}
        </div>
      )}
    </div>
  );
};

// ─── Main App ─────────────────────────────────────────────────────────────────
export default function App() {
  const [theme, setTheme] = useState("dark");
  const T = THEMES[theme];

  const [rawRows, setRawRows] = useState([]);
  const [agentMap, setAgentMap] = useState({});
  const [expressName, setExpressName] = useState("");
  const [agentFileName, setAgentFileName] = useState("");
  const [activeMonth, setActiveMonth] = useState("ALL");
  const [activeDay, setActiveDay] = useState("ALL");
  const [activeCountry, setActiveCountry] = useState("ALL");
  const [activeLOB, setActiveLOB] = useState("ALL");
  const [activeTeam, setActiveTeam] = useState("ALL");
  const [activeTab, setActiveTab] = useState("agents");
  const [sortKey, setSortKey] = useState("nps");
  const [sortDir, setSortDir] = useState(-1);
  const [search, setSearch] = useState("");
  const [openTooltip, setOpenTooltip] = useState(null); // empId of open tooltip

  const handleExpress = useCallback((buf, name) => {
    setRawRows(parseExpressReport(buf));
    setExpressName(name);
    setActiveMonth("ALL");
    setActiveDay("ALL");
  }, []);
  const handleAgents = useCallback((buf, name) => {
    setAgentMap(parseAgentsFile(buf));
    setAgentFileName(name);
  }, []);

  const enriched = useMemo(
    () =>
      rawRows.map((r) => {
        const agent = agentMap[r.empId] || null;
        return {
          ...r,
          agentName: agent?.name || null,
          lob: agent?.lob || null,
          wave: agent?.wave || null,
          tl: agent?.tl || null,
          matched: !!agent,
        };
      }),
    [rawRows, agentMap]
  );

  const months = useMemo(
    () =>
      [...new Set(enriched.map((r) => r.monthKey))].sort().map((k) => ({
        key: k,
        label: enriched.find((r) => r.monthKey === k)?.monthLabel || k,
      })),
    [enriched]
  );
  const days = useMemo(() => {
    const base =
      activeMonth === "ALL"
        ? enriched
        : enriched.filter((r) => r.monthKey === activeMonth);
    return [...new Set(base.map((r) => r.dayKey))].sort().map((k) => ({
      key: k,
      label: base.find((r) => r.dayKey === k)?.dayLabel || k,
    }));
  }, [enriched, activeMonth]);
  const countries = useMemo(
    () => [...new Set(enriched.map((r) => r.country))].sort(),
    [enriched]
  );
  const lobs = useMemo(
    () => [...new Set(enriched.map((r) => r.lob).filter(Boolean))].sort(),
    [enriched]
  );
  const teams = useMemo(
    () => [...new Set(enriched.map((r) => r.tl).filter(Boolean))].sort(),
    [enriched]
  );

  const filtered = useMemo(
    () =>
      enriched.filter((r) => {
        if (activeMonth !== "ALL" && r.monthKey !== activeMonth) return false;
        if (activeDay !== "ALL" && r.dayKey !== activeDay) return false;
        if (activeCountry !== "ALL" && r.country !== activeCountry)
          return false;
        if (activeLOB !== "ALL" && r.lob !== activeLOB) return false;
        if (activeTeam !== "ALL" && r.tl !== activeTeam) return false;
        return true;
      }),
    [enriched, activeMonth, activeDay, activeCountry, activeLOB, activeTeam]
  );

  const totalRows = filtered.length;
  const promoters = filtered.filter((r) => r.npsCategory === "Promoter").length;
  const passives = filtered.filter((r) => r.npsCategory === "Passive").length;
  const detractors = filtered.filter(
    (r) => r.npsCategory === "Detractor"
  ).length;
  const overallNPS = calcNPS(filtered);
  const avgLTR = totalRows
    ? (filtered.reduce((s, r) => s + (r.ltr || 0), 0) / totalRows).toFixed(2)
    : "—";
  const unmatchedCount = filtered.filter((r) => !r.matched && r.empId).length;

  const agentTable = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      const key = r.empId || "unknown";
      if (!map[key])
        map[key] = {
          empId: r.empId,
          name: r.agentName,
          lob: r.lob,
          wave: r.wave,
          tl: r.tl,
          matched: r.matched,
          rows: [],
        };
      map[key].rows.push(r);
    });
    return Object.values(map).map((a) => {
      const p = a.rows.filter((r) => r.npsCategory === "Promoter").length;
      const s = a.rows.filter((r) => r.npsCategory === "Passive").length;
      const d = a.rows.filter((r) => r.npsCategory === "Detractor").length;
      const t = a.rows.length;
      return {
        ...a,
        promoters: p,
        passives: s,
        detractors: d,
        total: t,
        nps: calcNPS(a.rows),
        avgLTR: t
          ? (a.rows.reduce((acc, r) => acc + (r.ltr || 0), 0) / t).toFixed(1)
          : null,
      };
    });
  }, [filtered]);

  const agentMonthTable = useMemo(() => {
    const map = {};
    enriched.forEach((r) => {
      if (activeTeam !== "ALL" && r.tl !== activeTeam) return;
      if (activeLOB !== "ALL" && r.lob !== activeLOB) return;
      const key = r.empId || "unknown";
      if (!map[key])
        map[key] = {
          empId: r.empId,
          name: r.agentName,
          lob: r.lob,
          wave: r.wave,
          matched: r.matched,
          months: {},
        };
      if (!map[key].months[r.monthKey]) map[key].months[r.monthKey] = [];
      map[key].months[r.monthKey].push(r);
    });
    return map;
  }, [enriched, activeTeam, activeLOB]);

  const sortedAgents = useMemo(() => {
    const s = [...agentTable].filter((a) => {
      if (!search) return true;
      const q = search.toLowerCase();
      return (
        (a.name || "").toLowerCase().includes(q) ||
        (a.empId || "").toLowerCase().includes(q) ||
        (a.lob || "").toLowerCase().includes(q) ||
        (a.tl || "").toLowerCase().includes(q)
      );
    });
    s.sort((a, b) => {
      let av = a[sortKey] ?? -9999,
        bv = b[sortKey] ?? -9999;
      return (av > bv ? 1 : -1) * sortDir;
    });
    return s;
  }, [agentTable, sortKey, sortDir, search]);

  const countryTable = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      if (!map[r.country]) map[r.country] = [];
      map[r.country].push(r);
    });
    return Object.entries(map)
      .map(([country, rows]) => ({
        country,
        total: rows.length,
        promoters: rows.filter((r) => r.npsCategory === "Promoter").length,
        passives: rows.filter((r) => r.npsCategory === "Passive").length,
        detractors: rows.filter((r) => r.npsCategory === "Detractor").length,
        nps: calcNPS(rows),
      }))
      .sort((a, b) => b.total - a.total);
  }, [filtered]);

  const lobTable = useMemo(() => {
    const map = {};
    filtered.forEach((r) => {
      const k = r.lob || "Unassigned";
      if (!map[k]) map[k] = [];
      map[k].push(r);
    });
    return Object.entries(map)
      .map(([lob, rows]) => ({
        lob,
        total: rows.length,
        promoters: rows.filter((r) => r.npsCategory === "Promoter").length,
        passives: rows.filter((r) => r.npsCategory === "Passive").length,
        detractors: rows.filter((r) => r.npsCategory === "Detractor").length,
        nps: calcNPS(rows),
      }))
      .sort((a, b) => b.total - a.total);
  }, [filtered]);

  const monthTrend = useMemo(() => {
    const base = enriched.filter((r) => {
      if (activeCountry !== "ALL" && r.country !== activeCountry) return false;
      if (activeLOB !== "ALL" && r.lob !== activeLOB) return false;
      if (activeTeam !== "ALL" && r.tl !== activeTeam) return false;
      return true;
    });
    const map = {};
    base.forEach((r) => {
      if (!map[r.monthKey])
        map[r.monthKey] = { key: r.monthKey, label: r.monthLabel, rows: [] };
      map[r.monthKey].rows.push(r);
    });
    return Object.values(map)
      .sort((a, b) => a.key.localeCompare(b.key))
      .map((m) => ({
        ...m,
        nps: calcNPS(m.rows),
        total: m.rows.length,
        promoters: m.rows.filter((r) => r.npsCategory === "Promoter").length,
        passives: m.rows.filter((r) => r.npsCategory === "Passive").length,
        detractors: m.rows.filter((r) => r.npsCategory === "Detractor").length,
      }));
  }, [enriched, activeCountry, activeLOB, activeTeam]);

  const toggleSort = (key) => {
    if (sortKey === key) setSortDir((d) => d * -1);
    else {
      setSortKey(key);
      setSortDir(-1);
    }
  };

  const thBase = {
    padding: "10px 14px",
    textAlign: "left",
    fontSize: 11,
    fontWeight: 700,
    background: T.headerBg,
    borderBottom: `1px solid ${T.border}`,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    whiteSpace: "nowrap",
    userSelect: "none",
  };
  const SortHd = ({ k, children }) => (
    <th
      onClick={() => toggleSort(k)}
      style={{
        ...thBase,
        color: sortKey === k ? T.accent : T.textMuted,
        cursor: "pointer",
      }}
    >
      {children}
      {sortKey === k ? (sortDir === -1 ? " ↓" : " ↑") : ""}
    </th>
  );
  const Th = ({ children, center }) => (
    <th
      style={{
        ...thBase,
        color: T.textMuted,
        textAlign: center ? "center" : "left",
      }}
    >
      {children}
    </th>
  );

  const selStyle = {
    padding: "7px 12px",
    borderRadius: 8,
    border: `1px solid ${T.border}`,
    fontFamily: "inherit",
    fontSize: 13,
    color: T.text,
    background: T.inputBg,
    cursor: "pointer",
    outline: "none",
  };

  const tabBtn = (id) => ({
    padding: "7px 16px",
    borderRadius: 7,
    border: "none",
    cursor: "pointer",
    fontFamily: "inherit",
    fontWeight: 700,
    fontSize: 12,
    background: activeTab === id ? T.accent : "transparent",
    color: activeTab === id ? "#fff" : T.textMuted,
    transition: "all 0.15s",
  });

  const tableBox = {
    background: T.surface,
    borderRadius: 12,
    border: `1px solid ${T.border}`,
    overflow: "hidden",
    boxShadow: theme === "light" ? "0 1px 6px #00000008" : "none",
  };
  const tableHdr = {
    padding: "12px 16px",
    borderBottom: `1px solid ${T.border}`,
    display: "flex",
    gap: 10,
    alignItems: "center",
  };

  const loaded = rawRows.length > 0;
  const activeFiltersCount = [
    activeMonth,
    activeDay,
    activeCountry,
    activeLOB,
    activeTeam,
  ].filter((f) => f !== "ALL").length;

  const headerGrad =
    theme === "dark"
      ? "linear-gradient(135deg, #060e1c 0%, #0b1f3a 60%, #060e1c 100%)"
      : "linear-gradient(135deg, #1B3A6B 0%, #1e4d8c 60%, #0e7490 100%)";

  return (
    <div
      style={{
        fontFamily: "'DM Sans', 'Segoe UI', sans-serif",
        background: T.bg,
        minHeight: "100vh",
        color: T.text,
      }}
    >
      {/* ── Header ── */}
      <div
        style={{
          background: headerGrad,
          borderBottom: `1px solid ${theme === "dark" ? "#1e293b" : "#1e4d8c"}`,
          padding: "0 36px",
        }}
      >
        <div
          style={{
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            paddingTop: 22,
            paddingBottom: 18,
          }}
        >
          <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
            {/* Samsung Logo */}
            <div style={{ display: "flex", flexDirection: "column" }}>
              <div
                style={{
                  fontFamily: "'Arial Black','Arial Bold',Impact,sans-serif",
                  fontSize: 30,
                  fontWeight: 900,
                  color: "#fff",
                  letterSpacing: "-0.5px",
                  lineHeight: 1,
                }}
              >
                SAMSUNG
              </div>
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  marginTop: 3,
                }}
              >
                <div
                  style={{
                    width: 28,
                    height: 2,
                    background: "linear-gradient(90deg,#38bdf8,transparent)",
                  }}
                />
                <span
                  style={{
                    fontSize: 9,
                    color: "#38bdf8",
                    letterSpacing: "0.22em",
                    fontWeight: 700,
                    textTransform: "uppercase",
                  }}
                >
                  Cairo · Contact Center
                </span>
              </div>
            </div>
            <div
              style={{
                width: 1,
                height: 44,
                background: "rgba(255,255,255,0.15)",
              }}
            />
            <div>
              <div
                style={{
                  fontSize: 19,
                  fontWeight: 800,
                  color: "#f8fafc",
                  letterSpacing: "-0.02em",
                }}
              >
                NPS Intelligence Dashboard
              </div>
              <div
                style={{
                  fontSize: 11,
                  color: "rgba(255,255,255,0.45)",
                  marginTop: 2,
                }}
              >
                Chat Surveys · Dynamic · Auto-deduplicated by Transaction ID
              </div>
            </div>
          </div>

          <div style={{ display: "flex", alignItems: "center", gap: 12 }}>
            {loaded && (
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <div
                  style={{
                    width: 7,
                    height: 7,
                    borderRadius: "50%",
                    background: "#10b981",
                    boxShadow: "0 0 8px #10b981",
                  }}
                />
                <span style={{ fontSize: 12, color: "rgba(255,255,255,0.4)" }}>
                  {rawRows.length.toLocaleString()} surveys · {months.length}{" "}
                  months
                </span>
              </div>
            )}
            {/* Dark / Light toggle */}
            <button
              onClick={() => setTheme(theme === "dark" ? "light" : "dark")}
              title={`Switch to ${theme === "dark" ? "light" : "dark"} mode`}
              style={{
                display: "flex",
                alignItems: "center",
                gap: 7,
                padding: "7px 14px",
                borderRadius: 20,
                border: "1.5px solid rgba(255,255,255,0.2)",
                background: "rgba(255,255,255,0.08)",
                color: "#fff",
                cursor: "pointer",
                fontFamily: "inherit",
                fontWeight: 700,
                fontSize: 12,
                transition: "all 0.2s",
              }}
            >
              {theme === "dark" ? "☀ Light" : "🌙 Dark"}
            </button>
          </div>
        </div>

        {/* Upload strip */}
        <div
          style={{
            display: "flex",
            gap: 12,
            paddingBottom: 22,
            flexWrap: "wrap",
          }}
        >
          <UploadZone
            label="Upload Express Report (.xlsx)"
            icon="📋"
            color="#38bdf8"
            hint="Col B: LTR · Col U: Transaction ID · Col AC: Employee ID"
            onFile={handleExpress}
            fileName={expressName}
            theme={theme}
          />
          <UploadZone
            label="Upload Samsung Cairo All-In-One (.xlsx)"
            icon="👥"
            color="#a78bfa"
            hint="Agents Distribution · Col A: BP Number · Col D: Name · Col H: LOB"
            onFile={handleAgents}
            fileName={agentFileName}
            theme={theme}
          />
        </div>
      </div>

      <div style={{ padding: "20px 36px" }}>
        {!loaded && (
          <div
            style={{
              background: T.surface,
              borderRadius: 16,
              padding: "70px 40px",
              textAlign: "center",
              border: `1px solid ${T.border}`,
            }}
          >
            <div style={{ fontSize: 48, marginBottom: 12 }}>📂</div>
            <div
              style={{
                fontSize: 18,
                fontWeight: 700,
                color: T.text,
                marginBottom: 6,
              }}
            >
              Upload both files above to get started
            </div>
            <div style={{ fontSize: 13, color: T.textMuted }}>
              Auto-processes, deduplicates by Transaction ID, matches agents,
              calculates NPS instantly.
            </div>
            <div style={{ fontSize: 12, marginTop: 6, color: T.textFaint }}>
              Each new month — upload the new Express Report. Everything updates
              in seconds.
            </div>
          </div>
        )}

        {loaded && (
          <>
            {/* ── Filter Bar ── */}
            <div
              style={{
                background: T.surface,
                borderRadius: 12,
                padding: "12px 16px",
                display: "flex",
                gap: 10,
                alignItems: "center",
                flexWrap: "wrap",
                marginBottom: 18,
                border: `1px solid ${T.border}`,
              }}
            >
              <span
                style={{
                  fontSize: 10,
                  fontWeight: 800,
                  color: T.textFaint,
                  textTransform: "uppercase",
                  letterSpacing: "0.08em",
                }}
              >
                FILTERS
              </span>
              {activeFiltersCount > 0 && (
                <span
                  style={{
                    background: T.accent + "20",
                    color: T.accent,
                    border: `1px solid ${T.accent}30`,
                    borderRadius: 20,
                    padding: "1px 9px",
                    fontSize: 11,
                    fontWeight: 700,
                  }}
                >
                  {activeFiltersCount} active
                </span>
              )}
              <select
                value={activeMonth}
                onChange={(e) => {
                  setActiveMonth(e.target.value);
                  setActiveDay("ALL");
                }}
                style={selStyle}
              >
                <option value="ALL">📅 All Months</option>
                {months.map((m) => (
                  <option key={m.key} value={m.key}>
                    {m.label}
                  </option>
                ))}
              </select>
              <select
                value={activeDay}
                onChange={(e) => setActiveDay(e.target.value)}
                style={{ ...selStyle, opacity: days.length === 0 ? 0.4 : 1 }}
              >
                <option value="ALL">
                  📆 All Days{activeMonth !== "ALL" ? ` (${days.length})` : ""}
                </option>
                {days.map((d) => (
                  <option key={d.key} value={d.key}>
                    {d.label}
                  </option>
                ))}
              </select>
              <select
                value={activeCountry}
                onChange={(e) => setActiveCountry(e.target.value)}
                style={selStyle}
              >
                <option value="ALL">🌍 All Countries</option>
                {countries.map((c) => (
                  <option key={c} value={c}>
                    {c}
                  </option>
                ))}
              </select>
              <select
                value={activeLOB}
                onChange={(e) => setActiveLOB(e.target.value)}
                style={selStyle}
              >
                <option value="ALL">🏢 All LOBs</option>
                {lobs.map((l) => (
                  <option key={l} value={l}>
                    {l}
                  </option>
                ))}
              </select>
              <select
                value={activeTeam}
                onChange={(e) => setActiveTeam(e.target.value)}
                style={selStyle}
              >
                <option value="ALL">👥 All Teams</option>
                {teams.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
              <button
                onClick={() => {
                  setActiveMonth("ALL");
                  setActiveDay("ALL");
                  setActiveCountry("ALL");
                  setActiveLOB("ALL");
                  setActiveTeam("ALL");
                }}
                style={{
                  marginLeft: "auto",
                  padding: "7px 14px",
                  borderRadius: 8,
                  border: `1px solid ${T.border}`,
                  background: T.bg,
                  color: T.textMuted,
                  cursor: "pointer",
                  fontSize: 12,
                  fontWeight: 700,
                }}
              >
                ✕ Reset
              </button>
              <span style={{ fontSize: 11, color: T.textSub }}>
                {totalRows.toLocaleString()} results
              </span>
            </div>

            {/* ── KPI Row ── */}
            <div
              style={{
                display: "flex",
                gap: 10,
                marginBottom: 18,
                flexWrap: "wrap",
              }}
            >
              <KpiCard
                icon="🎯"
                label="Overall NPS"
                value={
                  overallNPS !== null
                    ? overallNPS > 0
                      ? `+${overallNPS}`
                      : overallNPS
                    : "—"
                }
                color={npsColor(overallNPS)}
                sub={npsLabel(overallNPS)}
                theme={theme}
              />
              <KpiCard
                icon="📊"
                label="Total Surveys"
                value={totalRows.toLocaleString()}
                color={T.accent}
                sub={`${months.length} months`}
                theme={theme}
              />
              <KpiCard
                icon="🟢"
                label="Promoters"
                value={promoters.toLocaleString()}
                color="#10b981"
                sub={
                  totalRows
                    ? `${((promoters / totalRows) * 100).toFixed(1)}%`
                    : ""
                }
                theme={theme}
              />
              <KpiCard
                icon="🟡"
                label="Passives"
                value={passives.toLocaleString()}
                color="#f59e0b"
                sub={
                  totalRows
                    ? `${((passives / totalRows) * 100).toFixed(1)}%`
                    : ""
                }
                theme={theme}
              />
              <KpiCard
                icon="🔴"
                label="Detractors"
                value={detractors.toLocaleString()}
                color="#ef4444"
                sub={
                  totalRows
                    ? `${((detractors / totalRows) * 100).toFixed(1)}%`
                    : ""
                }
                theme={theme}
              />
              <KpiCard
                icon="⭐"
                label="Avg LTR"
                value={avgLTR}
                color="#a78bfa"
                sub="out of 10"
                theme={theme}
              />
              {agentFileName && (
                <KpiCard
                  icon="⚠️"
                  label="Unmatched IDs"
                  value={unmatchedCount.toLocaleString()}
                  color="#f97316"
                  sub="review needed"
                  theme={theme}
                />
              )}
            </div>

            {/* ── Monthly Trend ── */}
            <div
              style={{
                background: T.surface,
                borderRadius: 12,
                padding: "16px 20px",
                marginBottom: 18,
                border: `1px solid ${T.border}`,
              }}
            >
              <div
                style={{
                  fontSize: 11,
                  fontWeight: 800,
                  color: T.textMuted,
                  textTransform: "uppercase",
                  letterSpacing: "0.07em",
                  marginBottom: 14,
                }}
              >
                📅 Monthly NPS Trend — click bar to filter
              </div>
              <div
                style={{
                  display: "flex",
                  gap: 6,
                  alignItems: "flex-end",
                  overflowX: "auto",
                  paddingBottom: 4,
                }}
              >
                {monthTrend.map((m) => {
                  const isActive = activeMonth === m.key;
                  const maxH = 72;
                  const barH =
                    m.total > 0
                      ? Math.max(
                          18,
                          (m.total /
                            Math.max(...monthTrend.map((x) => x.total))) *
                            maxH
                        )
                      : 10;
                  const c = npsColor(m.nps);
                  return (
                    <div
                      key={m.key}
                      onClick={() => {
                        setActiveMonth(isActive ? "ALL" : m.key);
                        setActiveDay("ALL");
                      }}
                      style={{
                        display: "flex",
                        flexDirection: "column",
                        alignItems: "center",
                        gap: 4,
                        cursor: "pointer",
                        minWidth: 62,
                        opacity: activeMonth !== "ALL" && !isActive ? 0.25 : 1,
                        transition: "opacity 0.2s",
                      }}
                    >
                      <div
                        style={{
                          fontSize: 11,
                          fontFamily: "'DM Mono',monospace",
                          fontWeight: 800,
                          color: c,
                        }}
                      >
                        {m.nps !== null
                          ? m.nps > 0
                            ? `+${m.nps}`
                            : m.nps
                          : "—"}
                      </div>
                      <div
                        style={{
                          width: 44,
                          height: barH,
                          borderRadius: "5px 5px 0 0",
                          background: isActive ? c : c + "45",
                          border: `2px solid ${isActive ? c : "transparent"}`,
                          transition: "all 0.2s",
                        }}
                      />
                      <div
                        style={{
                          fontSize: 10,
                          color: T.textMuted,
                          fontWeight: 600,
                          textAlign: "center",
                        }}
                      >
                        {m.label}
                      </div>
                      <div style={{ fontSize: 10, color: T.textFaint }}>
                        {m.total.toLocaleString()}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>

            {/* ── Tabs ── */}
            <div
              style={{
                display: "flex",
                gap: 4,
                marginBottom: 14,
                background: T.surface,
                borderRadius: 10,
                padding: 4,
                width: "fit-content",
                border: `1px solid ${T.border}`,
              }}
            >
              {[
                { id: "agents", label: "👤 By Agent" },
                { id: "agentmonth", label: "📅 Agent × Month" },
                { id: "country", label: "🌍 By Country" },
                { id: "lob", label: "🏢 By LOB" },
              ].map((t) => (
                <button
                  key={t.id}
                  onClick={() => setActiveTab(t.id)}
                  style={tabBtn(t.id)}
                >
                  {t.label}
                </button>
              ))}
            </div>

            {/* ── Agent Table ── */}
            {activeTab === "agents" && (
              <div style={tableBox}>
                <div style={tableHdr}>
                  <span
                    style={{ fontSize: 13, fontWeight: 700, color: T.text }}
                  >
                    Agent NPS Performance
                  </span>
                  <span style={{ fontSize: 11, color: T.textSub }}>
                    {sortedAgents.length} agents
                  </span>
                  <div
                    style={{ fontSize: 11, color: T.textMuted, marginLeft: 4 }}
                  >
                    — click a name to copy DSATs / Passives
                  </div>
                  <input
                    value={search}
                    onChange={(e) => setSearch(e.target.value)}
                    placeholder="Search agent, ID, LOB, team…"
                    style={{
                      marginLeft: "auto",
                      padding: "6px 12px",
                      borderRadius: 8,
                      border: `1px solid ${T.border}`,
                      background: T.inputBg,
                      color: T.text,
                      fontFamily: "inherit",
                      fontSize: 13,
                      outline: "none",
                      width: 220,
                    }}
                  />
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <SortHd k="name">Agent Name</SortHd>
                        <Th>LOB</Th>
                        <Th>Wave</Th>
                        <Th>Team Leader</Th>
                        <SortHd k="total">Surveys</SortHd>
                        <SortHd k="promoters">🟢 P</SortHd>
                        <SortHd k="passives">🟡 Pa</SortHd>
                        <SortHd k="detractors">🔴 D</SortHd>
                        <SortHd k="nps">NPS Score</SortHd>
                        <SortHd k="avgLTR">Avg LTR</SortHd>
                        <Th>Distribution</Th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedAgents.map((a, i) => {
                        const isOpen = openTooltip === (a.empId || "unknown");
                        return (
                          <tr
                            key={a.empId}
                            style={{
                              borderBottom: `1px solid ${T.borderSoft}`,
                              background: i % 2 === 0 ? T.rowOdd : T.rowEven,
                            }}
                          >
                            <td
                              style={{
                                padding: "10px 14px",
                                position: "relative",
                              }}
                            >
                              <span
                                onClick={() =>
                                  setOpenTooltip(
                                    isOpen ? null : a.empId || "unknown"
                                  )
                                }
                                style={{
                                  fontSize: 13,
                                  fontWeight: 600,
                                  color: a.name ? T.accent : "#f97316",
                                  cursor: "pointer",
                                  textDecoration: "underline",
                                  textDecorationStyle: "dotted",
                                  textDecorationColor: a.name
                                    ? T.accent + "80"
                                    : "#f9731680",
                                  textUnderlineOffset: 3,
                                  userSelect: "none",
                                }}
                              >
                                {a.name || `⚠ ID: ${a.empId || "—"}`}
                              </span>
                              {isOpen && (
                                <AgentTooltip
                                  agent={a}
                                  onClose={() => setOpenTooltip(null)}
                                  theme={theme}
                                />
                              )}
                            </td>
                            <td style={{ padding: "10px 14px" }}>
                              {a.lob ? (
                                <Badge color="#a78bfa">{a.lob}</Badge>
                              ) : (
                                <span style={{ color: T.textFaint }}>—</span>
                              )}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 12,
                                color: T.textMuted,
                              }}
                            >
                              {a.wave || "—"}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 12,
                                color: T.textSub,
                              }}
                            >
                              {a.tl || "—"}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 13,
                                fontWeight: 700,
                                color: T.textSub,
                                textAlign: "right",
                              }}
                            >
                              {a.total}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 13,
                                color: "#10b981",
                                fontWeight: 700,
                                textAlign: "right",
                              }}
                            >
                              {a.promoters}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 13,
                                color: "#f59e0b",
                                fontWeight: 700,
                                textAlign: "right",
                              }}
                            >
                              {a.passives}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 13,
                                color: "#ef4444",
                                fontWeight: 700,
                                textAlign: "right",
                              }}
                            >
                              {a.detractors}
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                textAlign: "center",
                              }}
                            >
                              <NPSPill score={a.nps} />
                            </td>
                            <td
                              style={{
                                padding: "10px 14px",
                                fontSize: 12,
                                fontFamily: "'DM Mono',monospace",
                                color: T.textSub,
                                textAlign: "right",
                              }}
                            >
                              {a.avgLTR || "—"}
                            </td>
                            <td
                              style={{
                                padding: "10px 20px 10px 14px",
                                minWidth: 100,
                              }}
                            >
                              <MiniBar
                                promoters={a.promoters}
                                passives={a.passives}
                                detractors={a.detractors}
                                total={a.total}
                              />
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ── Agent × Month ── */}
            {activeTab === "agentmonth" && (
              <div style={tableBox}>
                <div style={tableHdr}>
                  <span
                    style={{ fontSize: 13, fontWeight: 700, color: T.text }}
                  >
                    Agent NPS — Month by Month
                  </span>
                  <span
                    style={{ fontSize: 11, color: T.textSub, marginLeft: 4 }}
                  >
                    filtered by Team & LOB · click name to copy DSATs
                  </span>
                </div>
                <div style={{ overflowX: "auto" }}>
                  <table style={{ width: "100%", borderCollapse: "collapse" }}>
                    <thead>
                      <tr>
                        <th
                          style={{
                            ...thBase,
                            color: T.textMuted,
                            position: "sticky",
                            left: 0,
                            zIndex: 2,
                          }}
                        >
                          Agent
                        </th>
                        <th style={{ ...thBase, color: T.textMuted }}>LOB</th>
                        {months.map((m) => (
                          <th
                            key={m.key}
                            style={{
                              ...thBase,
                              color:
                                activeMonth === m.key ? T.accent : T.textMuted,
                              textAlign: "center",
                              minWidth: 80,
                            }}
                          >
                            {m.label}
                          </th>
                        ))}
                        <th
                          style={{
                            ...thBase,
                            color: T.accent,
                            background:
                              theme === "dark" ? "#0a1f38" : "#eff6ff",
                            textAlign: "center",
                          }}
                        >
                          Overall
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      {Object.values(agentMonthTable)
                        .sort(
                          (a, b) =>
                            (calcNPS(Object.values(b.months).flat()) ?? -999) -
                            (calcNPS(Object.values(a.months).flat()) ?? -999)
                        )
                        .map((a, i) => {
                          const allRows = Object.values(a.months).flat();
                          const isOpen =
                            openTooltip === (a.empId || "unknown") + "-month";
                          const aFull = {
                            ...a,
                            rows: allRows,
                            promoters: allRows.filter(
                              (r) => r.npsCategory === "Promoter"
                            ).length,
                            passives: allRows.filter(
                              (r) => r.npsCategory === "Passive"
                            ).length,
                            detractors: allRows.filter(
                              (r) => r.npsCategory === "Detractor"
                            ).length,
                            total: allRows.length,
                            avgLTR: allRows.length
                              ? (
                                  allRows.reduce(
                                    (s, r) => s + (r.ltr || 0),
                                    0
                                  ) / allRows.length
                                ).toFixed(1)
                              : null,
                          };
                          return (
                            <tr
                              key={a.empId}
                              style={{
                                borderBottom: `1px solid ${T.borderSoft}`,
                                background: i % 2 === 0 ? T.rowOdd : T.rowEven,
                              }}
                            >
                              <td
                                style={{
                                  padding: "10px 14px",
                                  position: "sticky",
                                  left: 0,
                                  background:
                                    i % 2 === 0 ? T.stickyOdd : T.stickyEven,
                                  zIndex: 1,
                                }}
                              >
                                <span
                                  onClick={() =>
                                    setOpenTooltip(
                                      isOpen
                                        ? null
                                        : (a.empId || "unknown") + "-month"
                                    )
                                  }
                                  style={{
                                    fontSize: 13,
                                    fontWeight: 600,
                                    color: a.name ? T.accent : "#f97316",
                                    cursor: "pointer",
                                    textDecoration: "underline",
                                    textDecorationStyle: "dotted",
                                    textUnderlineOffset: 3,
                                    userSelect: "none",
                                  }}
                                >
                                  {a.name || (
                                    <span
                                      style={{ color: "#f97316", fontSize: 12 }}
                                    >
                                      ⚠ {a.empId}
                                    </span>
                                  )}
                                </span>
                                {isOpen && (
                                  <AgentTooltip
                                    agent={aFull}
                                    onClose={() => setOpenTooltip(null)}
                                    theme={theme}
                                  />
                                )}
                              </td>
                              <td style={{ padding: "10px 14px" }}>
                                {a.lob ? (
                                  <Badge color="#a78bfa">{a.lob}</Badge>
                                ) : (
                                  <span style={{ color: T.textFaint }}>—</span>
                                )}
                              </td>
                              {months.map((m) => {
                                const mRows = a.months[m.key] || [];
                                return (
                                  <td
                                    key={m.key}
                                    style={{
                                      padding: "8px",
                                      textAlign: "center",
                                    }}
                                  >
                                    {mRows.length > 0 ? (
                                      <div
                                        style={{
                                          display: "flex",
                                          flexDirection: "column",
                                          alignItems: "center",
                                          gap: 1,
                                        }}
                                      >
                                        <NPSPill score={calcNPS(mRows)} />
                                        <span
                                          style={{
                                            fontSize: 9,
                                            color: T.textSub,
                                          }}
                                        >
                                          {mRows.length}s
                                        </span>
                                      </div>
                                    ) : (
                                      <span
                                        style={{
                                          color: T.textFaint,
                                          fontSize: 11,
                                        }}
                                      >
                                        —
                                      </span>
                                    )}
                                  </td>
                                );
                              })}
                              <td
                                style={{
                                  padding: "10px 14px",
                                  textAlign: "center",
                                  background:
                                    theme === "dark"
                                      ? i % 2 === 0
                                        ? "#0a1f38"
                                        : "#060f1a"
                                      : i % 2 === 0
                                      ? "#eff6ff"
                                      : "#f0f9ff",
                                }}
                              >
                                <div
                                  style={{
                                    display: "flex",
                                    flexDirection: "column",
                                    alignItems: "center",
                                    gap: 1,
                                  }}
                                >
                                  <NPSPill score={calcNPS(allRows)} />
                                  <span
                                    style={{ fontSize: 9, color: T.textSub }}
                                  >
                                    {allRows.length}t
                                  </span>
                                </div>
                              </td>
                            </tr>
                          );
                        })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* ── Country Table ── */}
            {activeTab === "country" && (
              <div style={tableBox}>
                <div style={tableHdr}>
                  <span
                    style={{ fontSize: 13, fontWeight: 700, color: T.text }}
                  >
                    NPS by Country
                  </span>
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      {[
                        "Country",
                        "Surveys",
                        "🟢 Promoters",
                        "🟡 Passives",
                        "🔴 Detractors",
                        "NPS Score",
                        "Distribution",
                      ].map((h) => (
                        <th
                          key={h}
                          style={{
                            ...thBase,
                            color: T.textMuted,
                            textAlign:
                              h === "Country" || h === "Distribution"
                                ? "left"
                                : "center",
                          }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {countryTable.map((c, i) => (
                      <tr
                        key={c.country}
                        style={{
                          borderBottom: `1px solid ${T.borderSoft}`,
                          background: i % 2 === 0 ? T.rowOdd : T.rowEven,
                        }}
                      >
                        <td
                          style={{
                            padding: "12px 14px",
                            fontWeight: 700,
                            fontSize: 14,
                            color: T.text,
                          }}
                        >
                          {c.country}
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            fontFamily: "'DM Mono',monospace",
                            fontWeight: 700,
                            color: T.textSub,
                          }}
                        >
                          {c.total.toLocaleString()}
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#10b981",
                            fontWeight: 700,
                          }}
                        >
                          {c.promoters.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#10b98150" }}>
                            ({((c.promoters / c.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#f59e0b",
                            fontWeight: 700,
                          }}
                        >
                          {c.passives.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#f59e0b50" }}>
                            ({((c.passives / c.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#ef4444",
                            fontWeight: 700,
                          }}
                        >
                          {c.detractors.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#ef444450" }}>
                            ({((c.detractors / c.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{ padding: "12px 14px", textAlign: "center" }}
                        >
                          <NPSPill score={c.nps} />
                        </td>
                        <td
                          style={{
                            padding: "12px 20px 12px 14px",
                            minWidth: 140,
                          }}
                        >
                          <MiniBar
                            promoters={c.promoters}
                            passives={c.passives}
                            detractors={c.detractors}
                            total={c.total}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}

            {/* ── LOB Table ── */}
            {activeTab === "lob" && (
              <div style={tableBox}>
                <div style={tableHdr}>
                  <span
                    style={{ fontSize: 13, fontWeight: 700, color: T.text }}
                  >
                    NPS by Line of Business
                  </span>
                </div>
                <table style={{ width: "100%", borderCollapse: "collapse" }}>
                  <thead>
                    <tr>
                      {[
                        "LOB",
                        "Surveys",
                        "🟢 Promoters",
                        "🟡 Passives",
                        "🔴 Detractors",
                        "NPS Score",
                        "Distribution",
                      ].map((h) => (
                        <th
                          key={h}
                          style={{
                            ...thBase,
                            color: T.textMuted,
                            textAlign:
                              h === "LOB" || h === "Distribution"
                                ? "left"
                                : "center",
                          }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {lobTable.map((l, i) => (
                      <tr
                        key={l.lob}
                        style={{
                          borderBottom: `1px solid ${T.borderSoft}`,
                          background: i % 2 === 0 ? T.rowOdd : T.rowEven,
                        }}
                      >
                        <td style={{ padding: "12px 14px" }}>
                          <Badge color="#a78bfa">{l.lob}</Badge>
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            fontFamily: "'DM Mono',monospace",
                            fontWeight: 700,
                            color: T.textSub,
                          }}
                        >
                          {l.total.toLocaleString()}
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#10b981",
                            fontWeight: 700,
                          }}
                        >
                          {l.promoters.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#10b98150" }}>
                            ({((l.promoters / l.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#f59e0b",
                            fontWeight: 700,
                          }}
                        >
                          {l.passives.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#f59e0b50" }}>
                            ({((l.passives / l.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{
                            padding: "12px 14px",
                            textAlign: "center",
                            color: "#ef4444",
                            fontWeight: 700,
                          }}
                        >
                          {l.detractors.toLocaleString()}{" "}
                          <span style={{ fontSize: 11, color: "#ef444450" }}>
                            ({((l.detractors / l.total) * 100).toFixed(1)}%)
                          </span>
                        </td>
                        <td
                          style={{ padding: "12px 14px", textAlign: "center" }}
                        >
                          <NPSPill score={l.nps} />
                        </td>
                        <td
                          style={{
                            padding: "12px 20px 12px 14px",
                            minWidth: 140,
                          }}
                        >
                          <MiniBar
                            promoters={l.promoters}
                            passives={l.passives}
                            detractors={l.detractors}
                            total={l.total}
                          />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </>
        )}
      </div>
    </div>
  );
}
