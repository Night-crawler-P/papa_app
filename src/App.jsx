import { useEffect, useState, useRef, useCallback } from "react";
import { supabase } from "./supabase";
import * as XLSX from "xlsx";

// ─── Data ────────────────────────────────────────────────────────────────────

const CLUSTERS = [
  { cluster: "TUP", circles: ["UP West", "UP East"] },
  { cluster: "KOB", circles: ["Assam", "NE", "Orissa", "Kolkata", "ROB", "Bihar"] },
  { cluster: "KTN", circles: ["TN", "Kerala"] },
  { cluster: "GUJ", circles: ["Gujarat"] },
  { cluster: "RAD", circles: ["Rajasthan", "Delhi"] },
  { cluster: "MAH", circles: ["M&G"] },
  { cluster: "KAP", circles: ["Karnataka", "AP"] },
  { cluster: "PUH", circles: ["Haryana", "Punjab", "HP", "J&K"] },
  { cluster: "MUM", circles: ["Mumbai"] },
  { cluster: "MP",  circles: ["MP"] },
  { cluster: "Corp",circles: ["Corp"] },
  { cluster: "NLD", circles: ["NLD"] },
  { cluster: "ILD", circles: ["ILD"] },
  { cluster: "ISP", circles: ["ISP"] },
];

const DEPTS = [
  "NW AMC", "NW R&M", "IRU", "Leaseline", "Own OFC",
  "Commercial", "Admin", "HR", "Marketing", "Retail",
  "Finance", "Legal", "Local IT", "CS (non-central)",
  "DKYC/SIMEX/AI", "Fixed Deposit",
];

const STATUS_CFG = {
  received:      { label: "Received",      bg: "#22c55e", text: "#fff",     hasDate: true  },
  not_received:  { label: "Not Received",  bg: "#ef4444", text: "#fff",     hasDate: false },
  partial:       { label: "Partial",       bg: "#f59e0b", text: "#1a1200",  hasDate: false },
  na:            { label: "NA",            bg: "#f3f4f6", text: "#9ca3af",  hasDate: false, border: true },
  late_received: { label: "Late Received", bg: "#f97316", text: "#fff",     hasDate: true  },
};

const STATUS_KEYS = ["received", "not_received", "partial", "na", "late_received"];

function fmtDate(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}-${m}-${y.slice(2)}`;
}

function calcPendency(circleData) {
  let p = 0;
  DEPTS.forEach((dept) => {
    const c = circleData[dept];
    if (!c) return;
    if (c.status === "not_received" || c.status === "partial") p++;
  });
  return p;
}

// ─── Dropdown ────────────────────────────────────────────────────────────────

function Dropdown({ circle, dept, current, onCommit, onClose, anchorRect }) {
  const ref = useRef(null);
  const [pending, setPending] = useState(null);
  const [naBuffer, setNaBuffer] = useState("");
  const naTimer = useRef(null);

  useEffect(() => {
    function handler(e) {
      if (ref.current && !ref.current.contains(e.target)) onClose();
    }
    document.addEventListener("mousedown", handler);
    return () => document.removeEventListener("mousedown", handler);
  }, [onClose]);

  useEffect(() => {
    function handler(e) {
      const k = e.key.toLowerCase();
      if (k === "escape") { onClose(); return; }
      if (e.target && e.target.type === "date") return;

      if (k === "n") {
        setNaBuffer("n");
        if (naTimer.current) clearTimeout(naTimer.current);
        naTimer.current = setTimeout(() => {
          setNaBuffer("");
          handleSelect("not_received");
        }, 400);
        return;
      }
      if (k === "a" && naBuffer === "n") {
        if (naTimer.current) { clearTimeout(naTimer.current); naTimer.current = null; }
        setNaBuffer("");
        handleSelect("na");
        return;
      }
      setNaBuffer("");
      if (k === "r") handleSelect("received");
      else if (k === "p") handleSelect("partial");
      else if (k === "l") handleSelect("late_received");
    }
    document.addEventListener("keydown", handler);
    return () => document.removeEventListener("keydown", handler);
  }, [naBuffer]);

  function handleSelect(status) {
    if (STATUS_CFG[status].hasDate) {
      setPending(status);
      setTimeout(() => {
        const inp = document.getElementById("dd-date-inp");
        if (inp) inp.focus();
      }, 50);
    } else {
      onCommit(status, null);
    }
  }

  const activeStatus = pending || current?.status;

  const style = {
    position: "fixed",
    zIndex: 1000,
    background: "#fff",
    border: "1px solid #e5e7eb",
    borderRadius: "10px",
    boxShadow: "0 8px 24px rgba(0,0,0,0.12)",
    minWidth: "200px",
    padding: "6px",
    top: anchorRect ? anchorRect.bottom + 4 : 100,
    left: anchorRect ? Math.min(anchorRect.left, window.innerWidth - 220) : 100,
  };

  return (
    <div ref={ref} style={style}>
      {STATUS_KEYS.map((key) => {
        const cfg = STATUS_CFG[key];
        const isActive = activeStatus === key;
        return (
          <div
            key={key}
            onClick={() => handleSelect(key)}
            style={{
              display: "flex", alignItems: "center", gap: "8px",
              padding: "7px 10px", borderRadius: "7px", cursor: "pointer",
              background: isActive ? "#f3f4f6" : "transparent",
              fontSize: "12px", color: "#111",
            }}
            onMouseEnter={e => e.currentTarget.style.background = "#f3f4f6"}
            onMouseLeave={e => e.currentTarget.style.background = isActive ? "#f3f4f6" : "transparent"}
          >
            <div style={{
              width: 10, height: 10, borderRadius: "50%",
              background: cfg.bg, flexShrink: 0,
              border: cfg.border ? "1px solid #d1d5db" : "none",
            }} />
            <span style={{ flex: 1 }}>{cfg.label}</span>
            <span style={{
              fontSize: "10px", color: "#9ca3af",
              background: "#f9fafb", border: "1px solid #e5e7eb",
              borderRadius: "4px", padding: "1px 5px",
            }}>
              {key === "received" ? "R" : key === "not_received" ? "N" : key === "partial" ? "P" : key === "na" ? "NA" : "L"}
            </span>
          </div>
        );
      })}

      {(pending === "received" || pending === "late_received" ||
        (!pending && (current?.status === "received" || current?.status === "late_received"))) && (
        <div style={{ borderTop: "1px solid #e5e7eb", marginTop: "4px", padding: "8px 10px" }}>
          <input
            id="dd-date-inp"
            type="date"
            defaultValue={current?.received_date || ""}
            onChange={e => onCommit(pending || current?.status, e.target.value)}
            style={{
              width: "100%", fontSize: "12px", padding: "5px 8px",
              border: "1px solid #d1d5db", borderRadius: "6px",
              background: "#f9fafb", color: "#111",
            }}
          />
        </div>
      )}
    </div>
  );
}

// ─── Main App ─────────────────────────────────────────────────────────────────

export default function App() {
  const [cellData, setCellData] = useState({});
  const [remarks, setRemarks] = useState({});
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [dropdown, setDropdown] = useState(null);
  const [lastSaved, setLastSaved] = useState(null);

  // Load all data from Supabase
  useEffect(() => {
    loadAll();

    // Realtime subscription
    const channel = supabase
      .channel("status_entries_changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "status_entries" },
        (payload) => {
          if (payload.eventType === "INSERT" || payload.eventType === "UPDATE") {
            const row = payload.new;
            setCellData(prev => ({
              ...prev,
              [row.circle]: {
                ...(prev[row.circle] || {}),
                [row.department]: { status: row.status, received_date: row.received_date },
              },
            }));
            if (row.department === "__remark__") {
              setRemarks(prev => ({ ...prev, [row.circle]: row.status }));
            }
          }
        }
      )
      .subscribe();

    return () => supabase.removeChannel(channel);
  }, []);

  async function loadAll() {
    setLoading(true);
    const { data, error } = await supabase.from("status_entries").select("*");
    if (error) { console.error(error); setLoading(false); return; }

    const cells = {};
    const rem = {};
    data.forEach(row => {
      if (row.department === "__remark__") {
        rem[row.circle] = row.status;
        return;
      }
      if (!cells[row.circle]) cells[row.circle] = {};
      cells[row.circle][row.department] = {
        status: row.status,
        received_date: row.received_date,
      };
    });
    setCellData(cells);
    setRemarks(rem);
    setLoading(false);
  }

  const commitCell = useCallback(async (circle, dept, status, date) => {
    // Optimistic update
    setCellData(prev => ({
      ...prev,
      [circle]: {
        ...(prev[circle] || {}),
        [dept]: { status, received_date: date || null },
      },
    }));
    setDropdown(null);
    setSaving(true);

    const { error } = await supabase.from("status_entries").upsert(
      { cluster: CLUSTERS.find(c => c.circles.includes(circle))?.cluster || "", circle, department: dept, status, received_date: date || null },
      { onConflict: "circle,department" }
    );
    if (error) console.error(error);
    setSaving(false);
    setLastSaved(new Date());
  }, []);

  const commitRemark = useCallback(async (circle, value) => {
    setRemarks(prev => ({ ...prev, [circle]: value }));
    const { error } = await supabase.from("status_entries").upsert(
      { cluster: CLUSTERS.find(c => c.circles.includes(circle))?.cluster || "", circle, department: "__remark__", status: value, received_date: null },
      { onConflict: "circle,department" }
    );
    if (error) console.error(error);
  }, []);

  function openDropdown(e, circle, dept) {
    e.stopPropagation();
    const rect = e.currentTarget.getBoundingClientRect();
    setDropdown({ circle, dept, rect });
  }

  // Stats
  let totalPending = 0, totalReceived = 0;
  CLUSTERS.forEach(({ circles }) => circles.forEach(circle => {
    DEPTS.forEach(dept => {
      const c = (cellData[circle] || {})[dept];
      if (!c || c.status === "na") return;
      if (c.status === "not_received" || c.status === "partial") totalPending++;
      if (c.status === "received" || c.status === "late_received") totalReceived++;
    });
  }));

  function exportExcel() {
    const headers = ["Cluster", "Circle", ...DEPTS, "Remark", "Pendency"];
    const rows = [headers];
    CLUSTERS.forEach(({ cluster, circles }) => {
      circles.forEach((circle, ci) => {
        const row = [ci === 0 ? cluster : "", circle];
        DEPTS.forEach(dept => {
          const c = (cellData[circle] || {})[dept];
          if (!c) { row.push(""); return; }
          const cfg = STATUS_CFG[c.status] || STATUS_CFG.na;
          row.push(cfg.hasDate && c.received_date ? `${cfg.label} ${fmtDate(c.received_date)}` : cfg.label);
        });
        row.push(remarks[circle] || "");
        row.push(calcPendency(cellData[circle] || {}));
        rows.push(row);
      });
    });
    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "MIS");
    XLSX.writeFile(wb, "Input_Tracker.xlsx");
  }

  if (loading) return (
    <div style={{ display: "flex", flexDirection: "column", alignItems: "center", justifyContent: "center", height: "100vh", fontFamily: "Inter, sans-serif", background: "#0f172a", gap: "12px" }}>
      <div style={{ width: 32, height: 32, border: "3px solid #334155", borderTop: "3px solid #60a5fa", borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
      <span style={{ color: "#64748b", fontSize: "13px" }}>Loading tracker…</span>
      <style>{`@keyframes spin { to { transform: rotate(360deg); } }`}</style>
    </div>
  );

  return (
    <div style={{ fontFamily: "Inter, system-ui, sans-serif", minHeight: "100vh", background: "#f1f5f9" }}>
      {/* Header */}
      <div style={{
        position: "sticky", top: 0, zIndex: 50,
        background: "linear-gradient(135deg, #0f172a 0%, #1e3a5f 100%)",
        borderBottom: "1px solid #1e3a5f",
        padding: "0 20px",
        display: "flex", alignItems: "center", justifyContent: "space-between",
        height: "52px",
        boxShadow: "0 2px 12px rgba(0,0,0,0.3)",
      }}>
        <div style={{ display: "flex", alignItems: "center", gap: "14px" }}>
          <div style={{
            width: 28, height: 28, borderRadius: "7px",
            background: "linear-gradient(135deg, #3b82f6, #1d4ed8)",
            display: "flex", alignItems: "center", justifyContent: "center",
            fontSize: "13px", fontWeight: 700, color: "#fff",
            boxShadow: "0 2px 8px rgba(59,130,246,0.4)",
          }}>IT</div>
          <div>
            <div style={{ fontWeight: 600, fontSize: "14px", color: "#f8fafc", letterSpacing: "0.01em" }}>Input Tracker</div>
            <div style={{ fontSize: "10px", color: "#64748b", marginTop: "1px" }}>MIS — Monthly Cost Submissions</div>
          </div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
          {saving && (
            <span style={{ fontSize: "11px", color: "#60a5fa", display: "flex", alignItems: "center", gap: "5px" }}>
              <div style={{ width: 8, height: 8, borderRadius: "50%", background: "#60a5fa", animation: "pulse 1s infinite" }} />
              Saving…
            </span>
          )}
          {lastSaved && !saving && (
            <span style={{ fontSize: "10px", color: "#475569" }}>
              Saved {lastSaved.toLocaleTimeString([], { hour: "2-digit", minute: "2-digit" })}
            </span>
          )}
          <div style={{
            display: "flex", alignItems: "center", gap: "4px",
            background: "rgba(239,68,68,0.15)", border: "1px solid rgba(239,68,68,0.3)",
            borderRadius: "20px", padding: "3px 10px",
          }}>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#ef4444" }} />
            <span style={{ fontSize: "11px", color: "#fca5a5", fontWeight: 500 }}>{totalPending} pending</span>
          </div>
          <div style={{
            display: "flex", alignItems: "center", gap: "4px",
            background: "rgba(34,197,94,0.15)", border: "1px solid rgba(34,197,94,0.3)",
            borderRadius: "20px", padding: "3px 10px",
          }}>
            <div style={{ width: 6, height: 6, borderRadius: "50%", background: "#22c55e" }} />
            <span style={{ fontSize: "11px", color: "#86efac", fontWeight: 500 }}>{totalReceived} received</span>
          </div>
          <button
            onClick={exportExcel}
            style={{
              fontSize: "11px", padding: "5px 12px", borderRadius: "7px",
              border: "1px solid rgba(255,255,255,0.15)",
              background: "rgba(255,255,255,0.08)",
              color: "#e2e8f0",
              cursor: "pointer", fontWeight: 500,
              display: "flex", alignItems: "center", gap: "5px",
              transition: "background 0.15s",
            }}
            onMouseEnter={e => e.currentTarget.style.background = "rgba(255,255,255,0.15)"}
            onMouseLeave={e => e.currentTarget.style.background = "rgba(255,255,255,0.08)"}
          >
            ↓ Export Excel
          </button>
        </div>
      </div>

      {/* Table */}
      <div style={{ padding: "16px 20px" }}>
        <div style={{ overflowX: "auto", borderRadius: "12px", border: "1px solid #e2e8f0", background: "#fff", boxShadow: "0 1px 4px rgba(0,0,0,0.06)" }}>
          <table style={{ borderCollapse: "collapse", fontSize: "11px", width: "max-content", minWidth: "100%" }}>
            <thead>
              <tr>
                <th style={thStyle("#1e3a5f", 0, 60, "#93c5fd")}>Cluster</th>
                <th style={thStyle("#1e3a5f", 60, 110, "#93c5fd")}>Circle</th>
                {DEPTS.map(d => (
                  <th key={d} style={{ ...thBase, background: "#1e3a5f", color: "#93c5fd", minWidth: "100px", maxWidth: "110px" }}>{d}</th>
                ))}
                <th style={{ ...thBase, background: "#1e3a5f", color: "#93c5fd", minWidth: "120px" }}>Remark</th>
                <th style={{ ...thBase, background: "#1e3a5f", color: "#93c5fd", minWidth: "60px" }}>Pend.</th>
              </tr>
            </thead>
            <tbody>
              {CLUSTERS.map(({ cluster, circles }) =>
                circles.map((circle, ci) => {
                  const circleData = cellData[circle] || {};
                  const pendency = calcPendency(circleData);
                  return (
                    <tr key={circle} style={{ borderBottom: "1px solid #f1f5f9" }}>
                      {ci === 0 && (
                        <td rowSpan={circles.length} style={{
                          ...tdStickyBase, left: 0, width: 60, minWidth: 60,
                          textAlign: "center", fontWeight: 700, fontSize: "10px",
                          color: "#93c5fd", background: "#0f172a",
                          borderRight: "1px solid #1e3a5f",
                          verticalAlign: "middle",
                          letterSpacing: "0.05em",
                        }}>
                          {cluster}
                        </td>
                      )}
                      <td style={{
                        ...tdStickyBase, left: 60, minWidth: 110,
                        fontWeight: 500, fontSize: "11px", color: "#1e293b",
                        background: "#f8fafc", borderRight: "1px solid #e2e8f0",
                        padding: "0 10px", whiteSpace: "nowrap",
                      }}>
                        {circle}
                      </td>
                      {DEPTS.map(dept => {
                        const c = circleData[dept] || { status: "na", received_date: null };
                        const cfg = STATUS_CFG[c.status] || STATUS_CFG.na;
                        const label = cfg.hasDate && c.received_date
                          ? `${cfg.label} ${fmtDate(c.received_date)}`
                          : cfg.label;
                        return (
                          <td key={dept} style={{ padding: 0, border: "none", borderBottom: "1px solid #f3f4f6" }}>
                            <button
                              onClick={e => openDropdown(e, circle, dept)}
                              style={{
                                width: "100%", minHeight: "32px", padding: "3px 7px",
                                border: "none", borderRight: "1px solid rgba(0,0,0,0.05)",
                                background: cfg.bg, color: cfg.text,
                                cursor: "pointer", fontSize: "10px",
                                textAlign: "left", whiteSpace: "nowrap",
                                outline: cfg.border ? "1px solid #d1d5db" : "none",
                                display: "flex", alignItems: "center",
                              }}
                            >
                              {label}
                            </button>
                          </td>
                        );
                      })}
                      <td style={{ padding: 0, borderBottom: "1px solid #f3f4f6", borderRight: "1px solid #f3f4f6" }}>
                        <input
                          defaultValue={remarks[circle] || ""}
                          onBlur={e => commitRemark(circle, e.target.value)}
                          placeholder="—"
                          style={{
                            width: "100%", border: "none", background: "transparent",
                            fontSize: "10px", padding: "4px 8px", color: "#374151",
                            minWidth: "110px",
                          }}
                        />
                      </td>
                      <td style={{
                        textAlign: "center", fontWeight: 600, fontSize: "11px",
                        padding: "4px 8px", borderBottom: "1px solid #f3f4f6",
                        color: pendency > 0 ? "#ef4444" : "#22c55e",
                      }}>
                        {pendency}
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>

        {/* Legend */}
        <div style={{ display: "flex", gap: "10px", flexWrap: "wrap", marginTop: "12px", alignItems: "center", padding: "8px 12px", background: "#fff", borderRadius: "8px", border: "1px solid #e2e8f0" }}>
          <span style={{ fontSize: "10px", color: "#94a3b8", fontWeight: 600, textTransform: "uppercase", letterSpacing: "0.05em" }}>Legend</span>
          <div style={{ width: 1, height: 12, background: "#e2e8f0" }} />
          {Object.entries(STATUS_CFG).map(([key, cfg]) => (
            <div key={key} style={{ display: "flex", alignItems: "center", gap: "5px", fontSize: "11px", color: "#475569" }}>
              <div style={{
                width: 9, height: 9, borderRadius: "3px", background: cfg.bg,
                border: cfg.border ? "1px solid #cbd5e1" : "none", flexShrink: 0,
              }} />
              {cfg.label}
            </div>
          ))}
          <div style={{ width: 1, height: 12, background: "#e2e8f0", marginLeft: "auto" }} />
          <span style={{ fontSize: "10px", color: "#cbd5e1" }}>Click cell to edit · R N P NA L</span>
        </div>
      </div>

      {/* Dropdown */}
      {dropdown && (
        <Dropdown
          circle={dropdown.circle}
          dept={dropdown.dept}
          current={(cellData[dropdown.circle] || {})[dropdown.dept]}
          anchorRect={dropdown.rect}
          onCommit={(status, date) => commitCell(dropdown.circle, dropdown.dept, status, date)}
          onClose={() => setDropdown(null)}
        />
      )}
    </div>
  );
}

// ─── Style helpers ────────────────────────────────────────────────────────────

const thBase = {
  padding: "8px 10px",
  border: "none",
  borderBottom: "1px solid #1e3a5f",
  borderRight: "1px solid #1e3a5f",
  fontWeight: 600,
  fontSize: "10px",
  color: "#93c5fd",
  whiteSpace: "nowrap",
  textAlign: "left",
  position: "sticky",
  top: 0,
  zIndex: 2,
  letterSpacing: "0.03em",
  textTransform: "uppercase",
};

function thStyle(bg, left, minWidth, color = "#93c5fd") {
  return {
    ...thBase,
    background: bg,
    color,
    position: "sticky",
    left,
    top: 0,
    zIndex: 4,
    minWidth,
    borderRight: "1px solid #1e3a5f",
  };
}

const tdStickyBase = {
  position: "sticky",
  zIndex: 1,
  borderBottom: "1px solid #f3f4f6",
};
