import { useState, useRef, useMemo } from "react";

// ─── Brand Constants from Brand Bible ───
const BRAND = {
  indigo: "#34256B",
  chilli: "#DD3C27",
  flamingo: "#F59BB8",
  canary: "#F5CE00",
  black: "#000000",
  white: "#FFFFFF",
  indigoLight: "#4a3a8a",
  indigoDark: "#251a4f",
  chilliLight: "#e85a45",
  canaryLight: "#f7d633",
  flamingoLight: "#f8b8cc",
};

// ─── Utility: Parse CSV/TSV ───
function parseCSV(text) {
  const sep = text.includes("\t") ? "\t" : ",";
  const lines = text.trim().split("\n");
  const headers = lines[0].split(sep).map((h) => h.trim().replace(/^"|"$/g, ""));
  return lines.slice(1).map((line) => {
    const vals = line.split(sep).map((v) => v.trim().replace(/^"|"$/g, ""));
    const obj = {};
    headers.forEach((h, i) => (obj[h] = vals[i] || ""));
    return obj;
  });
}

// ─── Utility: Parse Excel using SheetJS ───
async function parseExcel(file) {
  const XLSX = await import("xlsx");
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array", cellDates: true });
        const ws = wb.Sheets[wb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { defval: "" });
        resolve(data);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── Utility: Export to Excel ───
async function exportToExcel(data, filename) {
  const XLSX = await import("xlsx");
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Audit Report");
  // Auto-width columns
  const colWidths = Object.keys(data[0] || {}).map((key) => ({
    wch: Math.max(key.length, ...data.map((r) => String(r[key] || "").length)) + 2,
  }));
  ws["!cols"] = colWidths;
  XLSX.writeFile(wb, filename);
}

// ─── Utility: Normalize strings for comparison ───
function normalize(str) {
  return String(str || "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

// ─── Utility: Normalize dates for comparison ───
function normalizeDate(str) {
  if (!str) return "";
  const s = String(str).trim();
  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return d.toISOString().split("T")[0];
  }
  const parts = s.split(/[/\-.]/);
  if (parts.length === 3) {
    if (parts[0].length === 4) return `${parts[0]}-${parts[1].padStart(2, "0")}-${parts[2].padStart(2, "0")}`;
    if (parts[2].length === 4) return `${parts[2]}-${parts[1].padStart(2, "0")}-${parts[0].padStart(2, "0")}`;
  }
  return s;
}

// ─── Utility: Check if IP is private/local ───
function isPrivateIP(ip) {
  return (
    ip.startsWith("10.") ||
    ip.startsWith("192.168.") ||
    /^172\.(1[6-9]|2\d|3[01])\./.test(ip) ||
    ip === "127.0.0.1" ||
    ip === "0.0.0.0" ||
    ip.startsWith("169.254.")
  );
}

// ─── Utility: Resolve IP addresses to geo locations via ipapi.co ───
async function resolveIpAddresses(ips, onProgress, onComplete) {
  const results = {};
  const errors = [];
  const DELAY_MS = 400;

  for (let i = 0; i < ips.length; i++) {
    const ip = ips[i];

    if (!ip || ip === "—" || ip === "-" || ip.trim() === "") {
      onProgress(i + 1, ips.length);
      continue;
    }

    if (isPrivateIP(ip.trim())) {
      results[ip] = { city: "Private Network", region: "", country: "", org: "Private/Local IP", isPrivate: true };
      onProgress(i + 1, ips.length);
      continue;
    }

    try {
      let response = await fetch(`https://ipapi.co/${encodeURIComponent(ip.trim())}/json/`);

      if (response.status === 429) {
        await new Promise((r) => setTimeout(r, 2000));
        response = await fetch(`https://ipapi.co/${encodeURIComponent(ip.trim())}/json/`);
      }

      if (!response.ok) {
        errors.push({ ip, error: `HTTP ${response.status}` });
      } else {
        const data = await response.json();
        if (data.error) {
          errors.push({ ip, error: data.reason || "API error" });
        } else {
          results[ip] = {
            city: data.city || "",
            region: data.region || "",
            country: data.country_name || "",
            org: data.org || "",
          };
        }
      }
    } catch (err) {
      errors.push({ ip, error: err.message || "Network error" });
    }

    onProgress(i + 1, ips.length);

    if (i < ips.length - 1) {
      await new Promise((r) => setTimeout(r, DELAY_MS));
    }
  }

  onComplete(results, errors);
}

// ─── Utility: Parse Sprout LogTime ("Jan  1 2026  3:45AM") ───
function parseLogTime(str) {
  if (!str) return null;
  let d;
  if (str instanceof Date) {
    d = str;
  } else {
    const s = String(str).trim();
    const num = Number(s);
    if (!isNaN(num) && num > 10000 && num < 100000) {
      d = new Date(Math.round((num - 25569) * 86400 * 1000));
    } else {
      const normalized = s.replace(/(\d)(AM|PM)/i, '$1 $2');
      d = new Date(normalized);
    }
  }
  if (isNaN(d.getTime())) return null;
  const days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
  const dateISO = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  let hours = d.getHours();
  const mins = String(d.getMinutes()).padStart(2, "0");
  const ampm = hours >= 12 ? "PM" : "AM";
  hours = hours % 12 || 12;
  const time12 = `${String(hours).padStart(2, "0")}:${mins} ${ampm}`;
  return { date: dateISO, time: time12, dayOfWeek: days[d.getDay()] };
}

// ─── Utility: Extract IP from Notes field ───
function extractIPFromNotes(notes) {
  if (!notes) return "";
  const match = String(notes).match(/at\s+([\d.]+)\s*$/);
  return match ? match[1] : "";
}

// ─── Utility: Parse Work Schedule Tracker Excel (multi-sheet monthly format) ───
async function parseScheduleExcel(file) {
  const XLSX = await import("xlsx");
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const wb = XLSX.read(new Uint8Array(e.target.result), { type: "array" });
        const result = {};
        const monthsParsed = [];

        const monthNames = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];

        for (const sheetName of wb.SheetNames) {
          const match = sheetName.match(/^(JANUARY|FEBRUARY|MARCH|APRIL|MAY|JUNE|JULY|AUGUST|SEPTEMBER|OCTOBER|NOVEMBER|DECEMBER)\s+(\d{4})$/i);
          if (!match) continue;

          const monthIndex = monthNames.indexOf(match[1].toUpperCase());
          const year = parseInt(match[2]);
          monthsParsed.push(sheetName);

          const ws = wb.Sheets[sheetName];
          const allRows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: "" });

          for (let r = 10; r < allRows.length; r++) {
            const row = allRows[r];
            const employeeName = String(row[8] || "").trim(); // Column I
            if (!employeeName) continue;

            const workSetup = String(row[0] || "").trim();
            const client = String(row[7] || "").trim();
            const tbrId = String(row[9] || "").trim();
            const normKey = employeeName.toLowerCase().replace(/[^a-z\s]/g, "").replace(/\s+/g, " ").trim();

            if (!result[normKey]) {
              result[normKey] = { fullName: employeeName, workSetup, client, tbrId, days: {} };
            }

            const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();
            for (let dayCol = 11; dayCol <= 41; dayCol++) {
              const dayNum = dayCol - 10;
              if (dayNum > daysInMonth) break;

              const cellValue = String(row[dayCol] || "").trim().toUpperCase();
              if (cellValue) {
                const dateISO = `${year}-${String(monthIndex + 1).padStart(2, "0")}-${String(dayNum).padStart(2, "0")}`;
                result[normKey].days[dateISO] = cellValue;
              }
            }
          }
        }

        result._meta = { monthsParsed, employeeCount: Object.keys(result).filter(k => k !== "_meta").length };
        resolve(result);
      } catch (err) { reject(err); }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}

// ─── Utility: Fuzzy name match ───
function fuzzyNameMatch(nameA, nameB) {
  if (!nameA || !nameB) return false;
  const clean = (s) =>
    String(s)
      .toLowerCase()
      .replace(/[^a-z\s]/g, "")
      .split(/\s+/)
      .filter((p) => p.length > 1);
  const partsA = clean(nameA);
  const partsB = clean(nameB);
  if (partsA.length === 0 || partsB.length === 0) return false;
  const [shorter, longer] = partsA.length <= partsB.length ? [partsA, partsB] : [partsB, partsA];
  const matched = shorter.filter((p) => longer.includes(p));
  return matched.length >= Math.min(shorter.length, 2);
}

// ─── Utility: Pre-process Sprout data (pair In/Out per employee per day) ───
function preprocessSproutData(rawData, columnMap) {
  const groups = {};

  for (const row of rawData) {
    const name = (row[columnMap.employeeName] || "").trim();
    const logTimeStr = row[columnMap.logTime] || "";
    const inOutMode = (row[columnMap.inOutMode] || "").trim().toLowerCase();
    const notes = row[columnMap.notes] || "";

    if (!name || !logTimeStr) continue;

    const parsed = parseLogTime(logTimeStr);
    if (!parsed) continue;

    const ip = extractIPFromNotes(notes);
    const key = `${name}|||${parsed.date}`;

    if (!groups[key]) {
      groups[key] = { employeeName: name, date: parsed.date, dayOfWeek: parsed.dayOfWeek, ins: [], outs: [] };
    }

    if (inOutMode.includes("in")) {
      groups[key].ins.push({ time: parsed.time, ip, raw: logTimeStr });
    } else if (inOutMode.includes("out")) {
      groups[key].outs.push({ time: parsed.time, ip, raw: logTimeStr });
    }
  }

  const results = [];
  for (const key of Object.keys(groups)) {
    const g = groups[key];
    // Sort by raw timestamp to get earliest in and latest out
    const sortByRaw = (a, b) => new Date(a.raw) - new Date(b.raw);
    g.ins.sort(sortByRaw);
    g.outs.sort(sortByRaw);

    const firstIn = g.ins[0] || null;
    const lastOut = g.outs[g.outs.length - 1] || null;

    results.push({
      employeeName: g.employeeName,
      date: g.date,
      dayOfWeek: g.dayOfWeek,
      clockIn: firstIn ? firstIn.time : "",
      clockOut: lastOut ? lastOut.time : "",
      ipIn: firstIn ? firstIn.ip : "",
      ipOut: lastOut ? lastOut.ip : "",
    });
  }

  // Sort by employee name then date
  results.sort((a, b) => a.employeeName.localeCompare(b.employeeName) || a.date.localeCompare(b.date));
  return results;
}

// ─── Status code labels for display ───
const STATUS_CODE_LABELS = {
  O: "Onsite",
  WFH: "Work From Home",
  OB: "Official Business",
  OW: "Official Work",
  V: "Vacation",
  H: "Holiday",
  LP: "Leave w/ Pay",
  SE: "Sick/Emergency Leave",
};

// ─── Utility: Check if IP falls within an office IP range ───
function ipInRange(ip, ipFrom, ipTo) {
  if (!ip || !ipFrom || !ipTo) return false;
  const parts = ip.trim().split(".").map(Number);
  const fromParts = ipFrom.trim().split(".").map(Number);
  const toParts = ipTo.trim().split(".").map(Number);
  if (parts.length !== 4 || fromParts.length !== 4 || toParts.length !== 4) return false;
  // Check first 3 octets match
  for (let i = 0; i < 3; i++) {
    if (parts[i] !== fromParts[i]) return false;
  }
  return parts[3] >= fromParts[3] && parts[3] <= toParts[3];
}

function isOfficeIP(ip, officeIPs) {
  if (!ip) return false;
  return officeIPs.some((o) => ipInRange(ip, o.ipFrom, o.ipTo));
}

// ─── Utility: Get office IP label ───
function getOfficeIPLabel(ip, officeIPs) {
  if (!ip) return null;
  const match = officeIPs.find((o) => ipInRange(ip, o.ipFrom, o.ipTo));
  return match ? match.label : null;
}

// ─── Plus Icon (Brand Icon) ───
function PlusIcon({ size = 24, color = BRAND.black }) {
  return (
    <svg width={size} height={size} viewBox="0 0 24 24" fill="none">
      <rect x="9" y="2" width="6" height="20" rx="0.5" fill={color} />
      <rect x="2" y="9" width="20" height="6" rx="0.5" fill={color} />
    </svg>
  );
}

// ─── Brand Logo (uses company PNG) ───
function BrandLogo({ inverted = false, height = 48 }) {
  return (
    <img
      src="/tbr-logo.png"
      alt="The Back Room"
      style={{
        height,
        width: "auto",
        filter: inverted ? "none" : "brightness(0)",
      }}
    />
  );
}

// ─── Hardcoded Users ───
const USERS = [
  { email: "admin@thebackroomop.com", password: "password123", name: "Admin" },
  { email: "payroll@thebackroomop.com", password: "password123", name: "Payroll Team" },
  { email: "alan@thebackroomop.com", password: "password123", name: "Alan Herrera" },
];

// ─── Login Page ───
function LoginPage({ onLogin }) {
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);

  const handleSubmit = (e) => {
    e.preventDefault();
    setLoading(true);
    setError("");

    setTimeout(() => {
      const user = USERS.find(
        (u) => u.email.toLowerCase() === email.toLowerCase() && u.password === password
      );

      if (user) {
        onLogin({ email: user.email, name: user.name });
      } else {
        setError("Invalid email or password.");
      }
      setLoading(false);
    }, 600);
  };

  return (
    <div
      style={{
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        background: `linear-gradient(135deg, ${BRAND.indigo} 0%, ${BRAND.indigoDark} 60%, ${BRAND.black} 100%)`,
        fontFamily: "'Montserrat', 'Helvetica Neue', sans-serif",
        position: "relative",
        overflow: "hidden",
      }}
    >
      {/* Decorative plus icons */}
      {[...Array(8)].map((_, i) => (
        <div
          key={i}
          style={{
            position: "absolute",
            top: `${10 + (i * 37) % 80}%`,
            left: `${5 + (i * 23) % 90}%`,
            opacity: 0.04 + (i % 3) * 0.02,
            transform: `rotate(${i * 45}deg) scale(${1 + (i % 3) * 0.5})`,
            animation: `float${i % 3} ${6 + i}s ease-in-out infinite`,
          }}
        >
          <PlusIcon size={40 + i * 10} color={BRAND.white} />
        </div>
      ))}

      {/* Branding above form */}
      <div
        style={{
          display: "flex",
          flexDirection: "column",
          alignItems: "center",
          textAlign: "center",
          marginBottom: 40,
          position: "relative",
          zIndex: 1,
        }}
      >
        <BrandLogo inverted />
        <h1
          style={{
            fontFamily: "'Montserrat', sans-serif",
            fontWeight: 900,
            fontSize: 36,
            color: BRAND.white,
            margin: "32px 0 12px",
            lineHeight: 1.1,
            letterSpacing: -1,
          }}
        >
          TIMESHEET <span style={{ color: BRAND.canary }}>AUDIT</span> SYSTEM
        </h1>
        <p
          style={{
            fontFamily: "'Montserrat', sans-serif",
            fontWeight: 300,
            fontSize: 14,
            color: "rgba(255,255,255,0.6)",
            maxWidth: 400,
            lineHeight: 1.6,
          }}
        >
          Automated comparison of Sprout payroll timesheets against employee work schedules.
        </p>
        <div style={{ display: "flex", gap: 8, marginTop: 20 }}>
          {[BRAND.canary, BRAND.chilli, BRAND.flamingo, BRAND.indigo].map((c, i) => (
            <div
              key={i}
              style={{
                width: 8,
                height: 8,
                borderRadius: "50%",
                background: c,
                opacity: 0.8,
              }}
            />
          ))}
        </div>
      </div>

      {/* Login form */}
      <div
        style={{
          position: "relative",
          zIndex: 1,
        }}
      >
        <div
          style={{
            background: BRAND.white,
            borderRadius: 2,
            padding: "48px 44px",
            width: 400,
            boxShadow: "0 32px 64px rgba(0,0,0,0.3)",
          }}
        >
          <div style={{ marginBottom: 32 }}>
            <h2
              style={{
                fontFamily: "'Montserrat', sans-serif",
                fontWeight: 900,
                fontSize: 22,
                color: BRAND.black,
                margin: 0,
                textTransform: "uppercase",
                letterSpacing: 1,
              }}
            >
              SIGN IN
            </h2>
            <div style={{ width: 40, height: 3, background: BRAND.canary, marginTop: 8 }} />
          </div>

          <form onSubmit={handleSubmit}>
            <div style={{ marginBottom: 20 }}>
              <label
                style={{
                  display: "block",
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 600,
                  fontSize: 11,
                  color: BRAND.black,
                  textTransform: "uppercase",
                  letterSpacing: 1.5,
                  marginBottom: 8,
                }}
              >
                Email
              </label>
              <input
                type="email"
                value={email}
                onChange={(e) => setEmail(e.target.value)}
                placeholder="you@thebackroomop.com"
                style={{
                  width: "100%",
                  padding: "14px 16px",
                  border: "2px solid #e0e0e0",
                  borderRadius: 2,
                  fontSize: 14,
                  fontFamily: "'Montserrat', sans-serif",
                  outline: "none",
                  transition: "border-color 0.2s",
                  boxSizing: "border-box",
                }}
                onFocus={(e) => (e.target.style.borderColor = BRAND.indigo)}
                onBlur={(e) => (e.target.style.borderColor = "#e0e0e0")}
              />
            </div>

            <div style={{ marginBottom: 24 }}>
              <label
                style={{
                  display: "block",
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 600,
                  fontSize: 11,
                  color: BRAND.black,
                  textTransform: "uppercase",
                  letterSpacing: 1.5,
                  marginBottom: 8,
                }}
              >
                Password
              </label>
              <input
                type="password"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Enter your password"
                style={{
                  width: "100%",
                  padding: "14px 16px",
                  border: "2px solid #e0e0e0",
                  borderRadius: 2,
                  fontSize: 14,
                  fontFamily: "'Montserrat', sans-serif",
                  outline: "none",
                  transition: "border-color 0.2s",
                  boxSizing: "border-box",
                }}
                onFocus={(e) => (e.target.style.borderColor = BRAND.indigo)}
                onBlur={(e) => (e.target.style.borderColor = "#e0e0e0")}
              />
            </div>

            {error && (
              <div
                style={{
                  background: "#FEF2F2",
                  border: `1px solid ${BRAND.chilli}`,
                  borderRadius: 2,
                  padding: "10px 14px",
                  marginBottom: 16,
                  fontSize: 13,
                  color: BRAND.chilli,
                  fontFamily: "'Montserrat', sans-serif",
                }}
              >
                {error}
              </div>
            )}

            <button
              type="submit"
              disabled={loading}
              style={{
                width: "100%",
                padding: "14px",
                background: BRAND.indigo,
                color: BRAND.white,
                border: "none",
                borderRadius: 2,
                fontSize: 13,
                fontFamily: "'Montserrat', sans-serif",
                fontWeight: 700,
                textTransform: "uppercase",
                letterSpacing: 2,
                cursor: loading ? "wait" : "pointer",
                transition: "all 0.2s",
                opacity: loading ? 0.7 : 1,
              }}
              onMouseEnter={(e) => !loading && (e.target.style.background = BRAND.indigoLight)}
              onMouseLeave={(e) => !loading && (e.target.style.background = BRAND.indigo)}
            >
              {loading ? "Signing in..." : "Sign In"}
            </button>
          </form>

          <p
            style={{
              textAlign: "center",
              marginTop: 24,
              fontSize: 12,
              color: "#999",
              fontFamily: "'Montserrat', sans-serif",
            }}
          >
            The Back Room Payroll Audit System
          </p>
        </div>
      </div>

      <style>{`
        @keyframes float0 { 0%,100% { transform: translateY(0) rotate(0deg); } 50% { transform: translateY(-20px) rotate(5deg); } }
        @keyframes float1 { 0%,100% { transform: translateY(0) rotate(45deg); } 50% { transform: translateY(-15px) rotate(50deg); } }
        @keyframes float2 { 0%,100% { transform: translateY(0) rotate(90deg); } 50% { transform: translateY(-25px) rotate(95deg); } }
      `}</style>
    </div>
  );
}

// ─── File Upload Zone ───
function FileUploadZone({ label, description, onFileLoad, accepted, fileLoaded }) {
  const [dragOver, setDragOver] = useState(false);
  const inputRef = useRef(null);

  const handleFile = async (file) => {
    if (!file) return;
    const ext = file.name.split(".").pop().toLowerCase();
    if (["xlsx", "xls"].includes(ext)) {
      const data = await parseExcel(file);
      onFileLoad(data, file.name, file);
    } else if (["csv", "tsv"].includes(ext)) {
      const text = await file.text();
      const data = parseCSV(text);
      onFileLoad(data, file.name, file);
    }
  };

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={(e) => {
        e.preventDefault();
        setDragOver(false);
        handleFile(e.dataTransfer.files[0]);
      }}
      onClick={() => inputRef.current?.click()}
      style={{
        border: `2px dashed ${dragOver ? BRAND.canary : fileLoaded ? BRAND.indigo : "#d0d0d0"}`,
        borderRadius: 2,
        padding: "40px 24px",
        textAlign: "center",
        cursor: "pointer",
        transition: "all 0.3s",
        background: dragOver ? "rgba(245,206,0,0.05)" : fileLoaded ? "rgba(52,37,107,0.03)" : BRAND.white,
        position: "relative",
      }}
    >
      <input
        ref={inputRef}
        type="file"
        accept={accepted}
        style={{ display: "none" }}
        onChange={(e) => handleFile(e.target.files[0])}
      />
      <div style={{ marginBottom: 12 }}>
        {fileLoaded ? (
          <div style={{ width: 48, height: 48, borderRadius: "50%", background: BRAND.indigo, display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke={BRAND.white} strokeWidth="2.5" strokeLinecap="round">
              <path d="M5 12l5 5L20 7" />
            </svg>
          </div>
        ) : (
          <div style={{ width: 48, height: 48, borderRadius: "50%", background: "#f5f5f5", display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            <PlusIcon size={20} color={BRAND.indigo} />
          </div>
        )}
      </div>
      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 14, color: BRAND.black, textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>
        {label}
      </div>
      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 12, color: "#888" }}>
        {fileLoaded || description}
      </div>
    </div>
  );
}

// ─── Column Mapper ───
function ColumnMapper({ title, columns, mapping, onMap, requiredFields }) {
  return (
    <div style={{ background: "#fafafa", borderRadius: 2, padding: 20, marginTop: 16 }}>
      <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12, color: BRAND.indigo, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 16 }}>
        {title} — Map Columns
      </div>
      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 }}>
        {requiredFields.map((field) => (
          <div key={field.key}>
            <label style={{ display: "block", fontFamily: "'Montserrat', sans-serif", fontWeight: 600, fontSize: 10, color: "#666", textTransform: "uppercase", letterSpacing: 1, marginBottom: 4 }}>
              {field.label} {field.required && <span style={{ color: BRAND.chilli }}>*</span>}
            </label>
            <select
              value={mapping[field.key] || ""}
              onChange={(e) => onMap({ ...mapping, [field.key]: e.target.value })}
              style={{
                width: "100%",
                padding: "8px 10px",
                border: "2px solid #e0e0e0",
                borderRadius: 2,
                fontSize: 12,
                fontFamily: "'Montserrat', sans-serif",
                background: BRAND.white,
                boxSizing: "border-box",
              }}
            >
              <option value="">— Select —</option>
              {columns.map((col) => (
                <option key={col} value={col}>{col}</option>
              ))}
            </select>
          </div>
        ))}
      </div>
    </div>
  );
}

// ─── Comparison Engine ───
function runComparison(processedSprout, scheduleLookup, officeIPs, ipGeoCache = {}) {
  const results = [];
  const schedKeys = Object.keys(scheduleLookup);

  for (const record of processedSprout) {
    const { employeeName, date, dayOfWeek, clockIn, clockOut, ipIn, ipOut } = record;

    // Fuzzy-match employee to schedule
    let schedEntry = null;
    for (const key of schedKeys) {
      if (fuzzyNameMatch(employeeName, scheduleLookup[key].fullName)) {
        schedEntry = scheduleLookup[key];
        break;
      }
    }

    // Determine IP location label
    let ipLocation = "—";
    if (ipIn) {
      const officeLabel = getOfficeIPLabel(ipIn, officeIPs);
      ipLocation = officeLabel ? `Office (${officeLabel})` : "Remote";
    }

    // Resolved geo from cache
    const resolvedIn = ipIn ? ipGeoCache[ipIn] : null;
    const resolvedLocation = resolvedIn && !resolvedIn.isPrivate
      ? [resolvedIn.city, resolvedIn.region, resolvedIn.country].filter(Boolean).join(", ")
      : "";

    let workSetup = "—";
    let scheduledStatus = "—";
    const discrepancies = [];
    let status = "Match";

    if (!schedEntry) {
      discrepancies.push("No schedule found for this employee");
      status = "Mismatch";
    } else {
      workSetup = schedEntry.workSetup || "—";
      const code = (schedEntry.days[date] || "").toUpperCase();

      if (!code) {
        // No schedule entry = rest day
        scheduledStatus = "Rest Day";
        if (clockIn || clockOut) {
          discrepancies.push("Clocked in on a Rest Day");
        }
      } else if (["V", "H", "LP", "SE"].includes(code)) {
        // Leave / Holiday — flag if clocked in
        scheduledStatus = STATUS_CODE_LABELS[code] || code;
        if (clockIn || clockOut) {
          discrepancies.push(`Clocked in while on ${STATUS_CODE_LABELS[code] || code}`);
        }
      } else if (["OB", "OW"].includes(code)) {
        // Official business — treat as working, no IP check needed
        scheduledStatus = STATUS_CODE_LABELS[code] || code;
      } else if (code === "O") {
        // Onsite — IP must match office
        scheduledStatus = "Onsite";
        if (ipIn && !isOfficeIP(ipIn, officeIPs)) {
          discrepancies.push(`Scheduled Onsite but IP (${ipIn}) is not a known office IP`);
        }
      } else if (code === "WFH") {
        // Work from home — IP should NOT be office
        scheduledStatus = "WFH";
        if (ipIn && isOfficeIP(ipIn, officeIPs)) {
          discrepancies.push(`Scheduled WFH but IP (${ipIn}) is a known office IP`);
        }
      } else {
        // Unknown code — just display it
        scheduledStatus = code;
      }

      // Clock-out IP differs from clock-in IP (different location types)
      if (ipIn && ipOut && ipIn !== ipOut) {
        const outIsOffice = isOfficeIP(ipOut, officeIPs);
        const inIsOffice = isOfficeIP(ipIn, officeIPs);
        if (inIsOffice !== outIsOffice) {
          discrepancies.push(`Clock-in and clock-out from different location types (In: ${ipIn}, Out: ${ipOut})`);
        }
      }

      if (discrepancies.length > 0) status = "Mismatch";
    }

    results.push({
      employeeName,
      date,
      dayOfWeek,
      clockIn: clockIn || "—",
      clockOut: clockOut || "—",
      ipIn: ipIn || "—",
      ipOut: ipOut || "—",
      ipLocation,
      resolvedLocation: resolvedLocation || "—",
      workSetup,
      scheduledStatus,
      auditStatus: status,
      discrepancies: discrepancies.join("; ") || "—",
    });
  }

  return results;
}

// ─── Status Badge ───
function StatusBadge({ status }) {
  const isMatch = status === "Match";
  return (
    <span
      style={{
        display: "inline-flex",
        alignItems: "center",
        gap: 4,
        padding: "4px 10px",
        borderRadius: 2,
        fontSize: 11,
        fontFamily: "'Montserrat', sans-serif",
        fontWeight: 700,
        textTransform: "uppercase",
        letterSpacing: 0.5,
        background: isMatch ? "rgba(52,37,107,0.08)" : "rgba(221,60,39,0.08)",
        color: isMatch ? BRAND.indigo : BRAND.chilli,
      }}
    >
      {isMatch ? "✓ Match" : "✗ Mismatch"}
    </span>
  );
}

// ─── Main Dashboard ───
function Dashboard({ user, onLogout }) {
  const [sproutData, setSproutData] = useState(null);
  const [sproutFile, setSproutFile] = useState("");
  const [scheduleFile, setScheduleFile] = useState("");
  const [sproutCols, setSproutCols] = useState([]);
  const [sproutMap, setSproutMap] = useState({});
  const [results, setResults] = useState(null);
  const [filter, setFilter] = useState("all");
  const [searchTerm, setSearchTerm] = useState("");
  const [dateFilter, setDateFilter] = useState("");
  const [step, setStep] = useState(1);
  const [processing, setProcessing] = useState(false);
  const [ipGeoCache, setIpGeoCache] = useState({});
  const [ipResolutionStatus, setIpResolutionStatus] = useState("idle");
  const [ipResolutionProgress, setIpResolutionProgress] = useState({ current: 0, total: 0 });
  const [ipResolutionErrors, setIpResolutionErrors] = useState([]);
  const [processedSprout, setProcessedSprout] = useState(null);
  const [scheduleLookup, setScheduleLookup] = useState(null);
  const [officeIPs, setOfficeIPs] = useState([
    { label: "EST", ipFrom: "116.50.227.176", ipTo: "116.50.227.181" },
    { label: "ComClark", ipFrom: "161.49.192.112", ipTo: "161.49.192.116" },
    { label: "ComClark 7F", ipFrom: "136.239.243.241", ipTo: "136.239.243.246" },
  ]);
  const [preprocessSummary, setPreprocessSummary] = useState(null);
  const [expandedEmployees, setExpandedEmployees] = useState(new Set());
  const [allExpanded, setAllExpanded] = useState(false);
  const [pageSize, setPageSize] = useState(25);
  const [currentPage, setCurrentPage] = useState(1);
  const [innerSortConfig, setInnerSortConfig] = useState({ key: "date", direction: "asc" });

  const sproutFields = [
    { key: "employeeName", label: "Employee Name", required: true },
    { key: "logTime", label: "Log Time", required: true },
    { key: "inOutMode", label: "In/Out Mode", required: true },
    { key: "notes", label: "Notes / IP Source", required: true },
  ];

  const autoMap = (columns, fields) => {
    const map = {};
    fields.forEach((f) => {
      const nc = (s) => normalize(s);
      const match = columns.find((c) => {
        const col = nc(c);
        if (f.key === "employeeName") return col.includes("fullname") || col === "name" || col.includes("employeename") || (col.includes("full") && col.includes("name"));
        if (f.key === "logTime") return col.includes("logtime") || col.includes("logtim");
        if (f.key === "inOutMode") return col.includes("inout") || col.includes("inoutmode");
        if (f.key === "notes") return col.includes("notes") || col.includes("note");
        return col.includes(nc(f.key)) || col.includes(nc(f.label.split("(")[0].trim()));
      });
      if (match) map[f.key] = match;
    });
    return map;
  };

  const handleSproutLoad = (data, name) => {
    setSproutData(data);
    setSproutFile(name);
    const cols = Object.keys(data[0] || {});
    setSproutCols(cols);
    setSproutMap(autoMap(cols, sproutFields));
  };

  const handleScheduleFile = async (file) => {
    try {
      const lookup = await parseScheduleExcel(file);
      const meta = lookup._meta || { monthsParsed: [], employeeCount: 0 };
      delete lookup._meta;
      setScheduleLookup(lookup);
      setScheduleFile(file.name);
      setPreprocessSummary((prev) => prev ? { ...prev, scheduleEmployees: meta.employeeCount, scheduleMonths: meta.monthsParsed } : null);
    } catch (err) {
      console.error("Failed to parse schedule file:", err);
    }
  };

  const handleResolveIPs = () => {
    if (!processedSprout || processedSprout.length === 0) return;

    const allIPs = processedSprout
      .flatMap((r) => [r.ipIn, r.ipOut])
      .filter((ip) => ip && ip !== "—" && ip !== "-");
    const uniqueIPs = [...new Set(allIPs)];

    if (uniqueIPs.length === 0) {
      setIpResolutionStatus("done");
      setIpResolutionErrors([{ ip: "N/A", error: "No valid IP addresses found in the data" }]);
      return;
    }

    const unresolved = uniqueIPs.filter((ip) => !ipGeoCache[ip]);
    if (unresolved.length === 0) {
      setIpResolutionStatus("done");
      return;
    }

    setIpResolutionStatus("resolving");
    setIpResolutionProgress({ current: 0, total: unresolved.length });
    setIpResolutionErrors([]);

    resolveIpAddresses(
      unresolved,
      (current, total) => setIpResolutionProgress({ current, total }),
      (results, errors) => {
        setIpGeoCache((prev) => ({ ...prev, ...results }));
        setIpResolutionErrors(errors);
        setIpResolutionStatus("done");
      }
    );
  };

  const canProceedToMap = sproutData && scheduleLookup;
  const canRunComparison = processedSprout && processedSprout.length > 0 && scheduleLookup;

  const handlePreprocess = () => {
    if (!sproutData || !scheduleLookup) return;
    if (!sproutMap.employeeName || !sproutMap.logTime || !sproutMap.inOutMode || !sproutMap.notes) return;

    const processed = preprocessSproutData(sproutData, sproutMap);
    setProcessedSprout(processed);

    const uniqueEmployees = new Set(processed.map((r) => r.employeeName));
    const uniqueDates = new Set(processed.map((r) => r.date));
    setPreprocessSummary({
      records: processed.length,
      employees: uniqueEmployees.size,
      dates: uniqueDates.size,
      scheduleEmployees: Object.keys(scheduleLookup).length,
    });
  };

  const handleRunComparison = () => {
    setProcessing(true);
    setTimeout(() => {
      const res = runComparison(processedSprout, scheduleLookup, officeIPs, ipGeoCache);
      setResults(res);
      setStep(3);
      setProcessing(false);
      // Auto-expand employees with mismatches
      const mismatchEmployees = new Set(
        res.filter(r => r.auditStatus === "Mismatch").map(r => r.employeeName)
      );
      setExpandedEmployees(mismatchEmployees);
      setAllExpanded(false);
      setCurrentPage(1);
    }, 1200);
  };

  const filteredResults = results
    ? results.filter((r) => {
        const matchFilter = filter === "all" || (filter === "match" && r.auditStatus === "Match") || (filter === "mismatch" && r.auditStatus === "Mismatch");
        const matchSearch = !searchTerm || r.employeeName.toLowerCase().includes(searchTerm.toLowerCase());
        const matchDate = !dateFilter || r.date === dateFilter;
        return matchFilter && matchSearch && matchDate;
      })
    : [];

  const filtersActive = filter !== "all" || searchTerm !== "" || dateFilter !== "";

  const totalStats = results
    ? {
        total: results.length,
        matches: results.filter((r) => r.auditStatus === "Match").length,
        mismatches: results.filter((r) => r.auditStatus === "Mismatch").length,
        officeIPs: results.filter((r) => r.ipLocation && r.ipLocation.startsWith("Office")).length,
        remoteIPs: results.filter((r) => r.ipLocation === "Remote").length,
      }
    : null;

  const stats = useMemo(() => ({
    total: filteredResults.length,
    matches: filteredResults.filter((r) => r.auditStatus === "Match").length,
    mismatches: filteredResults.filter((r) => r.auditStatus === "Mismatch").length,
    officeIPs: filteredResults.filter((r) => r.ipLocation && r.ipLocation.startsWith("Office")).length,
    remoteIPs: filteredResults.filter((r) => r.ipLocation === "Remote").length,
  }), [filteredResults]);

  const groupedResults = useMemo(() => {
    const groups = {};
    for (const row of filteredResults) {
      if (!groups[row.employeeName]) groups[row.employeeName] = [];
      groups[row.employeeName].push(row);
    }
    return Object.entries(groups)
      .map(([name, rows]) => ({
        employeeName: name,
        rows: rows.sort((a, b) => a.date.localeCompare(b.date)),
        totalDays: rows.length,
        mismatches: rows.filter(r => r.auditStatus === "Mismatch").length,
        workSetup: rows[0]?.workSetup || "—",
        hasDiscrepancies: rows.some(r => r.auditStatus === "Mismatch"),
      }))
      .sort((a, b) => a.employeeName.localeCompare(b.employeeName));
  }, [filteredResults]);

  const totalGroupPages = pageSize === "all" ? 1 : Math.ceil(groupedResults.length / pageSize);
  const paginatedGroups = pageSize === "all"
    ? groupedResults
    : groupedResults.slice((currentPage - 1) * pageSize, currentPage * pageSize);

  const toggleEmployee = (name) => {
    setExpandedEmployees(prev => {
      const next = new Set(prev);
      if (next.has(name)) next.delete(name);
      else next.add(name);
      return next;
    });
  };

  const toggleAllExpanded = () => {
    if (allExpanded) {
      setExpandedEmployees(new Set());
      setAllExpanded(false);
    } else {
      setExpandedEmployees(new Set(paginatedGroups.map(g => g.employeeName)));
      setAllExpanded(true);
    }
  };

  const SUB_TABLE_COLUMNS = [
    { key: "date", label: "Date" },
    { key: "dayOfWeek", label: "Day" },
    { key: "clockIn", label: "Clock In" },
    { key: "clockOut", label: "Clock Out" },
    { key: "ipIn", label: "IP (In)" },
    { key: "ipLocation", label: "IP Location" },
    { key: "resolvedLocation", label: "Resolved Loc." },
    { key: "workSetup", label: "Work Setup" },
    { key: "scheduledStatus", label: "Sched. Status" },
    { key: "auditStatus", label: "Status" },
    { key: "discrepancies", label: "Discrepancies" },
  ];

  const handleExport = () => {
    if (filteredResults.length > 0) {
      const exportData = filteredResults.map((r) => ({
        "Employee": r.employeeName,
        "Date": r.date,
        "Day": r.dayOfWeek,
        "Clock In": r.clockIn,
        "Clock Out": r.clockOut,
        "IP (In)": r.ipIn,
        "IP (Out)": r.ipOut,
        "IP Location": r.ipLocation,
        "Resolved Location": r.resolvedLocation,
        "Work Setup": r.workSetup,
        "Scheduled Status": r.scheduledStatus,
        "Audit Status": r.auditStatus,
        "Discrepancies": r.discrepancies,
      }));
      exportToExcel(exportData, `TBR_Timesheet_Audit_${new Date().toISOString().split("T")[0]}.xlsx`);
    }
  };

  const handlePrint = () => {
    window.print();
  };

  return (
    <div style={{ minHeight: "100vh", background: "#f7f7f8", fontFamily: "'Montserrat', 'Helvetica Neue', sans-serif" }}>
      {/* Top Nav */}
      <nav
        style={{
          background: BRAND.white,
          borderBottom: `3px solid ${BRAND.canary}`,
          padding: "0 32px",
          height: 64,
          display: "flex",
          alignItems: "center",
          justifyContent: "space-between",
          position: "sticky",
          top: 0,
          zIndex: 100,
          boxShadow: "0 2px 12px rgba(0,0,0,0.04)",
        }}
      >
        <BrandLogo />
        <div style={{ display: "flex", alignItems: "center", gap: 20 }}>
          <span style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 500, fontSize: 13, color: "#666" }}>
            {user.name}
          </span>
          <button
            onClick={onLogout}
            style={{
              padding: "8px 16px",
              background: "transparent",
              border: "1px solid #ddd",
              borderRadius: 2,
              fontSize: 11,
              fontFamily: "'Montserrat', sans-serif",
              fontWeight: 600,
              textTransform: "uppercase",
              letterSpacing: 1,
              cursor: "pointer",
              color: "#666",
              transition: "all 0.2s",
            }}
            onMouseEnter={(e) => {
              e.target.style.borderColor = BRAND.chilli;
              e.target.style.color = BRAND.chilli;
            }}
            onMouseLeave={(e) => {
              e.target.style.borderColor = "#ddd";
              e.target.style.color = "#666";
            }}
          >
            Sign Out
          </button>
        </div>
      </nav>

      {/* Step Progress */}
      <div style={{ background: BRAND.white, padding: "20px 32px", borderBottom: "1px solid #eee" }}>
        <div style={{ display: "flex", gap: 32, maxWidth: 600 }}>
          {[
            { n: 1, label: "Upload Files" },
            { n: 2, label: "Map Columns" },
            { n: 3, label: "View Results" },
          ].map((s) => (
            <div
              key={s.n}
              style={{ display: "flex", alignItems: "center", gap: 10, opacity: step >= s.n ? 1 : 0.3, cursor: step > s.n ? "pointer" : "default" }}
              onClick={() => step > s.n && setStep(s.n)}
            >
              <div
                style={{
                  width: 28,
                  height: 28,
                  borderRadius: "50%",
                  background: step >= s.n ? BRAND.indigo : "#e0e0e0",
                  color: BRAND.white,
                  display: "flex",
                  alignItems: "center",
                  justifyContent: "center",
                  fontSize: 12,
                  fontWeight: 700,
                  fontFamily: "'Montserrat', sans-serif",
                }}
              >
                {step > s.n ? "✓" : s.n}
              </div>
              <span
                style={{
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: step === s.n ? 700 : 400,
                  fontSize: 12,
                  color: step >= s.n ? BRAND.black : "#999",
                  textTransform: "uppercase",
                  letterSpacing: 1,
                }}
              >
                {s.label}
              </span>
              {s.n < 3 && (
                <div style={{ width: 40, height: 2, background: step > s.n ? BRAND.indigo : "#e0e0e0", marginLeft: 8 }} />
              )}
            </div>
          ))}
        </div>
      </div>

      <div style={{ padding: "32px", maxWidth: 1400, margin: "0 auto" }}>
        {/* STEP 1: Upload */}
        {step === 1 && (
          <div>
            <h2 style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 28, color: BRAND.black, margin: "0 0 8px", textTransform: "uppercase", letterSpacing: 1 }}>
              Upload Files
            </h2>
            <p style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 14, color: "#888", margin: "0 0 32px" }}>
              Upload the Sprout payroll export and the employee work schedule file.
            </p>

            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 24 }}>
              <div>
                <FileUploadZone
                  label="Sprout Payroll Export"
                  description="Drop .xlsx, .xls, or .csv file here"
                  accepted=".xlsx,.xls,.csv,.tsv"
                  onFileLoad={handleSproutLoad}
                  fileLoaded={sproutFile}
                />
                {sproutData && (
                  <div style={{ marginTop: 12, padding: "12px 16px", background: "#f0f0f0", borderRadius: 2 }}>
                    <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: "#666" }}>
                      <strong>{sproutData.length}</strong> records loaded from <strong>{sproutFile}</strong>
                    </span>
                  </div>
                )}
              </div>
              <div>
                <FileUploadZone
                  label="Work Schedule Tracker"
                  description="Drop the monthly schedule .xlsx file here"
                  accepted=".xlsx,.xls"
                  onFileLoad={(_, name, file) => handleScheduleFile(file)}
                  fileLoaded={scheduleFile}
                />
                {scheduleLookup && (
                  <div style={{ marginTop: 12, padding: "12px 16px", background: "#f0f0f0", borderRadius: 2 }}>
                    <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: "#666" }}>
                      <strong>{Object.keys(scheduleLookup).length}</strong> employees loaded from <strong>{scheduleFile}</strong>
                    </span>
                  </div>
                )}
              </div>
            </div>

            {canProceedToMap && (
              <div style={{ marginTop: 32, textAlign: "right" }}>
                <button
                  onClick={() => setStep(2)}
                  style={{
                    padding: "14px 32px",
                    background: BRAND.indigo,
                    color: BRAND.white,
                    border: "none",
                    borderRadius: 2,
                    fontSize: 13,
                    fontFamily: "'Montserrat', sans-serif",
                    fontWeight: 700,
                    textTransform: "uppercase",
                    letterSpacing: 2,
                    cursor: "pointer",
                    transition: "background 0.2s",
                  }}
                  onMouseEnter={(e) => (e.target.style.background = BRAND.indigoLight)}
                  onMouseLeave={(e) => (e.target.style.background = BRAND.indigo)}
                >
                  Next: Map Columns →
                </button>
              </div>
            )}
          </div>
        )}

        {/* STEP 2: Map Columns */}
        {step === 2 && (
          <div>
            <h2 style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 28, color: BRAND.black, margin: "0 0 8px", textTransform: "uppercase", letterSpacing: 1 }}>
              Map Columns
            </h2>
            <p style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 14, color: "#888", margin: "0 0 32px" }}>
              Match the Sprout payroll columns to the required fields. Schedule data is auto-parsed from the tracker file.
            </p>

            <ColumnMapper
              title="Sprout Payroll"
              columns={sproutCols}
              mapping={sproutMap}
              onMap={setSproutMap}
              requiredFields={sproutFields}
            />

            {/* Schedule Summary */}
            {scheduleLookup && (
              <div style={{ background: "rgba(52,37,107,0.04)", borderRadius: 2, padding: "16px 20px", marginTop: 16, border: `1px solid ${BRAND.indigo}` }}>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12, color: BRAND.indigo, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 8 }}>
                  Schedule Data (Auto-Parsed)
                </div>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 13, color: BRAND.indigo }}>
                  <strong>{Object.keys(scheduleLookup).length}</strong> employees loaded from <strong>{scheduleFile}</strong>
                </div>
              </div>
            )}

            {/* Office IP Configuration */}
            <div
              style={{
                background: "#fafafa",
                borderRadius: 2,
                padding: 20,
                marginTop: 24,
                border: "1px solid #e0e0e0",
              }}
            >
              <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12, color: BRAND.indigo, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 12 }}>
                Known Office IPs
              </div>
              <p style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 12, color: "#666", margin: "0 0 16px", lineHeight: 1.5 }}>
                Configure known office IP addresses. Employees scheduled at the office will be flagged if their clock-in IP doesn't match.
              </p>
              <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
                {officeIPs.map((entry, idx) => (
                  <div key={idx} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                    <input
                      type="text"
                      value={entry.label}
                      onChange={(e) => {
                        const updated = [...officeIPs];
                        updated[idx] = { ...updated[idx], label: e.target.value };
                        setOfficeIPs(updated);
                      }}
                      placeholder="Label (e.g. EST)"
                      style={{ width: 120, padding: "8px 10px", border: "2px solid #e0e0e0", borderRadius: 2, fontSize: 12, fontFamily: "'Montserrat', sans-serif", boxSizing: "border-box" }}
                    />
                    <input
                      type="text"
                      value={entry.ipFrom}
                      onChange={(e) => {
                        const updated = [...officeIPs];
                        updated[idx] = { ...updated[idx], ipFrom: e.target.value };
                        setOfficeIPs(updated);
                      }}
                      placeholder="IP From"
                      style={{ width: 160, padding: "8px 10px", border: "2px solid #e0e0e0", borderRadius: 2, fontSize: 12, fontFamily: "'Montserrat', monospace", boxSizing: "border-box" }}
                    />
                    <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: "#999" }}>to</span>
                    <input
                      type="text"
                      value={entry.ipTo}
                      onChange={(e) => {
                        const updated = [...officeIPs];
                        updated[idx] = { ...updated[idx], ipTo: e.target.value };
                        setOfficeIPs(updated);
                      }}
                      placeholder="IP To"
                      style={{ width: 160, padding: "8px 10px", border: "2px solid #e0e0e0", borderRadius: 2, fontSize: 12, fontFamily: "'Montserrat', monospace", boxSizing: "border-box" }}
                    />
                    <button
                      onClick={() => setOfficeIPs(officeIPs.filter((_, i) => i !== idx))}
                      style={{ padding: "6px 12px", background: "transparent", border: "1px solid #ddd", borderRadius: 2, fontSize: 11, cursor: "pointer", color: BRAND.chilli, fontWeight: 600 }}
                    >
                      Remove
                    </button>
                  </div>
                ))}
                <button
                  onClick={() => setOfficeIPs([...officeIPs, { label: "", ipFrom: "", ipTo: "" }])}
                  style={{ alignSelf: "flex-start", padding: "6px 16px", background: "transparent", border: `1px dashed ${BRAND.indigo}`, borderRadius: 2, fontSize: 11, fontFamily: "'Montserrat', sans-serif", fontWeight: 600, color: BRAND.indigo, cursor: "pointer", textTransform: "uppercase", letterSpacing: 1 }}
                >
                  + Add IP
                </button>
              </div>
            </div>

            {/* Pre-process Data Button */}
            {sproutMap.employeeName && sproutMap.logTime && sproutMap.inOutMode && sproutMap.notes && scheduleLookup && (
              <div style={{ marginTop: 24 }}>
                <button
                  onClick={handlePreprocess}
                  style={{
                    padding: "12px 28px",
                    background: BRAND.canary,
                    color: BRAND.black,
                    border: "none",
                    borderRadius: 2,
                    fontSize: 12,
                    fontFamily: "'Montserrat', sans-serif",
                    fontWeight: 700,
                    textTransform: "uppercase",
                    letterSpacing: 1.5,
                    cursor: "pointer",
                    transition: "background 0.2s",
                  }}
                  onMouseEnter={(e) => (e.target.style.background = BRAND.canaryLight)}
                  onMouseLeave={(e) => (e.target.style.background = BRAND.canary)}
                >
                  Pre-process Data
                </button>
              </div>
            )}

            {/* Pre-processing Summary */}
            {preprocessSummary && (
              <div
                style={{
                  marginTop: 16,
                  padding: "16px 20px",
                  background: "rgba(52,37,107,0.04)",
                  borderRadius: 2,
                  border: `1px solid ${BRAND.indigo}`,
                  display: "flex",
                  alignItems: "center",
                  gap: 12,
                }}
              >
                <span style={{ width: 24, height: 24, borderRadius: "50%", background: BRAND.indigo, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                  <svg width="12" height="12" viewBox="0 0 24 24" fill="none" stroke={BRAND.white} strokeWidth="3" strokeLinecap="round"><path d="M5 12l5 5L20 7" /></svg>
                </span>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 13, color: BRAND.indigo }}>
                  <strong>{preprocessSummary.records}</strong> attendance records for <strong>{preprocessSummary.employees}</strong> employees across <strong>{preprocessSummary.dates}</strong> dates.
                  Schedule loaded for <strong>{preprocessSummary.scheduleEmployees}</strong> employees.
                </div>
              </div>
            )}

            {/* IP Geolocation Resolution (uses preprocessed IPs) */}
            {processedSprout && processedSprout.length > 0 && (
              <div
                style={{
                  background: "#fafafa",
                  borderRadius: 2,
                  padding: 20,
                  marginTop: 24,
                  border: `1px solid ${ipResolutionStatus === "done" ? BRAND.indigo : "#e0e0e0"}`,
                }}
              >
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 12, color: BRAND.indigo, textTransform: "uppercase", letterSpacing: 1.5, marginBottom: 12 }}>
                  IP Geolocation Lookup (Optional)
                </div>
                <p style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 12, color: "#666", margin: "0 0 16px", lineHeight: 1.5 }}>
                  Resolve IP addresses extracted from clock-in/out events to physical locations. This adds a "Resolved Location" column to results.
                </p>

                {ipResolutionStatus === "idle" && (
                  <div>
                    <button
                      onClick={handleResolveIPs}
                      style={{
                        padding: "10px 24px",
                        background: BRAND.canary,
                        color: BRAND.black,
                        border: "none",
                        borderRadius: 2,
                        fontSize: 12,
                        fontFamily: "'Montserrat', sans-serif",
                        fontWeight: 700,
                        textTransform: "uppercase",
                        letterSpacing: 1.5,
                        cursor: "pointer",
                        transition: "background 0.2s",
                      }}
                      onMouseEnter={(e) => (e.target.style.background = BRAND.canaryLight)}
                      onMouseLeave={(e) => (e.target.style.background = BRAND.canary)}
                    >
                      Resolve IP Locations
                    </button>
                    <div style={{ fontSize: 10, color: "#999", marginTop: 8, fontFamily: "'Montserrat', sans-serif" }}>
                      Free tier: 1,000 lookups/day via ipapi.co
                    </div>
                  </div>
                )}

                {ipResolutionStatus === "resolving" && (
                  <div>
                    <div style={{ display: "flex", justifyContent: "space-between", fontFamily: "'Montserrat', sans-serif", fontSize: 11, color: "#666", marginBottom: 8 }}>
                      <span>Resolving {ipResolutionProgress.current} of {ipResolutionProgress.total} unique IPs...</span>
                      <span>{ipResolutionProgress.total > 0 ? Math.round((ipResolutionProgress.current / ipResolutionProgress.total) * 100) : 0}%</span>
                    </div>
                    <div style={{ width: "100%", height: 8, background: "#e0e0e0", borderRadius: 4, overflow: "hidden" }}>
                      <div
                        style={{
                          width: `${ipResolutionProgress.total > 0 ? (ipResolutionProgress.current / ipResolutionProgress.total) * 100 : 0}%`,
                          height: "100%",
                          background: BRAND.canary,
                          borderRadius: 4,
                          transition: "width 0.3s ease",
                        }}
                      />
                    </div>
                    <div style={{ marginTop: 8, fontFamily: "'Montserrat', sans-serif", fontSize: 10, color: "#999" }}>
                      Estimated time remaining: ~{Math.ceil((ipResolutionProgress.total - ipResolutionProgress.current) * 0.4)}s
                    </div>
                  </div>
                )}

                {ipResolutionStatus === "done" && (
                  <div>
                    <div style={{ display: "flex", alignItems: "center", gap: 8, padding: "10px 14px", background: "rgba(52,37,107,0.05)", borderRadius: 2, marginBottom: ipResolutionErrors.length > 0 ? 12 : 0 }}>
                      <span style={{ width: 20, height: 20, borderRadius: "50%", background: BRAND.indigo, display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 }}>
                        <svg width="10" height="10" viewBox="0 0 24 24" fill="none" stroke={BRAND.white} strokeWidth="3" strokeLinecap="round"><path d="M5 12l5 5L20 7" /></svg>
                      </span>
                      <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: BRAND.indigo, fontWeight: 600 }}>
                        Resolved {Object.keys(ipGeoCache).length} unique IP addresses
                      </span>
                      <button
                        onClick={() => { setIpResolutionStatus("idle"); }}
                        style={{ marginLeft: "auto", padding: "4px 12px", background: "transparent", border: `1px solid ${BRAND.indigo}`, borderRadius: 2, fontSize: 10, fontFamily: "'Montserrat', sans-serif", fontWeight: 600, color: BRAND.indigo, cursor: "pointer", textTransform: "uppercase", letterSpacing: 1 }}
                      >
                        Re-run
                      </button>
                    </div>
                    {ipResolutionErrors.length > 0 && (
                      <div style={{ padding: "10px 14px", background: "rgba(221,60,39,0.05)", borderRadius: 2, border: "1px solid rgba(221,60,39,0.15)" }}>
                        <div style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, fontWeight: 600, color: BRAND.chilli, marginBottom: 6 }}>
                          {ipResolutionErrors.length} IP(s) could not be resolved:
                        </div>
                        <div style={{ maxHeight: 80, overflowY: "auto", fontSize: 10, fontFamily: "monospace", color: "#666" }}>
                          {ipResolutionErrors.map((e, i) => (
                            <div key={i}>{e.ip}: {e.error}</div>
                          ))}
                        </div>
                      </div>
                    )}
                  </div>
                )}
              </div>
            )}

            <div style={{ marginTop: 32, display: "flex", justifyContent: "space-between" }}>
              <button
                onClick={() => setStep(1)}
                style={{
                  padding: "14px 32px",
                  background: "transparent",
                  color: BRAND.indigo,
                  border: `2px solid ${BRAND.indigo}`,
                  borderRadius: 2,
                  fontSize: 13,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: 2,
                  cursor: "pointer",
                }}
              >
                ← Back
              </button>
              <button
                onClick={handleRunComparison}
                disabled={!canRunComparison || processing}
                style={{
                  padding: "14px 32px",
                  background: canRunComparison ? BRAND.chilli : "#ccc",
                  color: BRAND.white,
                  border: "none",
                  borderRadius: 2,
                  fontSize: 13,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: 2,
                  cursor: canRunComparison ? "pointer" : "not-allowed",
                  transition: "background 0.2s",
                  display: "flex",
                  alignItems: "center",
                  gap: 8,
                }}
              >
                {processing ? (
                  <>
                    <span style={{ display: "inline-block", width: 16, height: 16, border: "2px solid rgba(255,255,255,0.3)", borderTopColor: BRAND.white, borderRadius: "50%", animation: "spin 0.8s linear infinite" }} />
                    Processing...
                  </>
                ) : (
                  <>Run Comparison</>
                )}
              </button>
            </div>
          </div>
        )}

        {/* STEP 3: Results */}
        {step === 3 && results && (
          <div>
            {/* Stats Cards */}
            <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr 1fr 1fr", gap: 20, marginBottom: 28 }}>
              <div style={{ background: BRAND.white, borderRadius: 2, padding: 24, borderLeft: `4px solid ${BRAND.indigo}`, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 11, color: "#999", textTransform: "uppercase", letterSpacing: 1.5 }}>Total Records</div>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 36, color: BRAND.indigo, marginTop: 4 }}>{stats.total}</div>
                {filtersActive && totalStats && <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 11, color: "#999", marginTop: 2 }}>of {totalStats.total} total</div>}
              </div>
              <div style={{ background: BRAND.white, borderRadius: 2, padding: 24, borderLeft: `4px solid ${BRAND.canary}`, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 11, color: "#999", textTransform: "uppercase", letterSpacing: 1.5 }}>Matches</div>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 36, color: BRAND.indigo, marginTop: 4 }}>
                  {stats.matches}
                  <span style={{ fontSize: 14, fontWeight: 400, color: "#999", marginLeft: 8 }}>({stats.total > 0 ? ((stats.matches / stats.total) * 100).toFixed(1) : 0}%)</span>
                </div>
                {filtersActive && totalStats && <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 11, color: "#999", marginTop: 2 }}>of {totalStats.matches} total</div>}
              </div>
              <div style={{ background: BRAND.white, borderRadius: 2, padding: 24, borderLeft: `4px solid ${BRAND.chilli}`, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 11, color: "#999", textTransform: "uppercase", letterSpacing: 1.5 }}>Mismatches</div>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 36, color: BRAND.chilli, marginTop: 4 }}>
                  {stats.mismatches}
                  <span style={{ fontSize: 14, fontWeight: 400, color: "#999", marginLeft: 8 }}>({stats.total > 0 ? ((stats.mismatches / stats.total) * 100).toFixed(1) : 0}%)</span>
                </div>
                {filtersActive && totalStats && <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 400, fontSize: 11, color: "#999", marginTop: 2 }}>of {totalStats.mismatches} total</div>}
              </div>
              <div style={{ background: BRAND.white, borderRadius: 2, padding: 24, borderLeft: `4px solid ${BRAND.flamingo}`, boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 11, color: "#999", textTransform: "uppercase", letterSpacing: 1.5 }}>IP Location</div>
                <div style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 900, fontSize: 24, color: BRAND.indigo, marginTop: 4 }}>
                  {stats.officeIPs} <span style={{ fontSize: 11, fontWeight: 400, color: "#999" }}>Office</span>
                  {" / "}
                  {stats.remoteIPs} <span style={{ fontSize: 11, fontWeight: 400, color: "#999" }}>Remote</span>
                </div>
              </div>
            </div>

            {/* Filters & Actions */}
            <div
              style={{
                background: BRAND.white,
                borderRadius: 2,
                padding: "16px 20px",
                marginBottom: 20,
                display: "flex",
                alignItems: "center",
                gap: 16,
                flexWrap: "wrap",
                boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
              }}
            >
              <div style={{ display: "flex", gap: 4 }}>
                {[
                  { key: "all", label: "All" },
                  { key: "match", label: "Matches" },
                  { key: "mismatch", label: "Mismatches" },
                ].map((f) => (
                  <button
                    key={f.key}
                    onClick={() => { setFilter(f.key); setCurrentPage(1); }}
                    style={{
                      padding: "8px 16px",
                      background: filter === f.key ? BRAND.indigo : "transparent",
                      color: filter === f.key ? BRAND.white : "#666",
                      border: `1px solid ${filter === f.key ? BRAND.indigo : "#ddd"}`,
                      borderRadius: 2,
                      fontSize: 11,
                      fontFamily: "'Montserrat', sans-serif",
                      fontWeight: 600,
                      textTransform: "uppercase",
                      letterSpacing: 1,
                      cursor: "pointer",
                      transition: "all 0.2s",
                    }}
                  >
                    {f.label}
                  </button>
                ))}
              </div>

              <button
                onClick={toggleAllExpanded}
                style={{
                  padding: "8px 16px",
                  background: "transparent",
                  color: BRAND.indigo,
                  border: `1px solid ${BRAND.indigo}`,
                  borderRadius: 2,
                  fontSize: 11,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 600,
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  cursor: "pointer",
                  transition: "all 0.2s",
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                {allExpanded ? "Collapse All" : "Expand All"}
              </button>

              <input
                type="text"
                placeholder="Search employee..."
                value={searchTerm}
                onChange={(e) => { setSearchTerm(e.target.value); setCurrentPage(1); }}
                style={{
                  padding: "8px 14px",
                  border: "2px solid #e0e0e0",
                  borderRadius: 2,
                  fontSize: 12,
                  fontFamily: "'Montserrat', sans-serif",
                  width: 200,
                  outline: "none",
                }}
                onFocus={(e) => (e.target.style.borderColor = BRAND.indigo)}
                onBlur={(e) => (e.target.style.borderColor = "#e0e0e0")}
              />

              <input
                type="date"
                value={dateFilter}
                onChange={(e) => { setDateFilter(e.target.value); setCurrentPage(1); }}
                style={{
                  padding: "8px 14px",
                  border: "2px solid #e0e0e0",
                  borderRadius: 2,
                  fontSize: 12,
                  fontFamily: "'Montserrat', sans-serif",
                  outline: "none",
                }}
              />

              <div style={{ flex: 1 }} />

              <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: "#999" }}>
                {groupedResults.length} employees · {filteredResults.length} records
              </span>

              <button
                onClick={handleExport}
                style={{
                  padding: "8px 20px",
                  background: BRAND.canary,
                  color: BRAND.black,
                  border: "none",
                  borderRadius: 2,
                  fontSize: 11,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                  transition: "background 0.2s",
                }}
                onMouseEnter={(e) => (e.target.style.background = BRAND.canaryLight)}
                onMouseLeave={(e) => (e.target.style.background = BRAND.canary)}
              >
                Export Excel
              </button>

              <button
                onClick={handlePrint}
                style={{
                  padding: "8px 20px",
                  background: "transparent",
                  color: BRAND.indigo,
                  border: `1px solid ${BRAND.indigo}`,
                  borderRadius: 2,
                  fontSize: 11,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: 1,
                  cursor: "pointer",
                  display: "flex",
                  alignItems: "center",
                  gap: 6,
                }}
              >
                Print
              </button>
            </div>

            {/* Employee Accordion */}
            <div style={{ background: BRAND.white, borderRadius: 2, overflow: "hidden", boxShadow: "0 2px 8px rgba(0,0,0,0.04)" }}>
              {paginatedGroups.length === 0 ? (
                <div style={{ padding: 40, textAlign: "center", color: "#999", fontFamily: "'Montserrat', sans-serif" }}>
                  No records match your filters.
                </div>
              ) : (
                paginatedGroups.map((group) => {
                  const isExpanded = expandedEmployees.has(group.employeeName);
                  const sortedRows = [...group.rows].sort((a, b) => {
                    const dir = innerSortConfig.direction === "asc" ? 1 : -1;
                    const valA = a[innerSortConfig.key] || "";
                    const valB = b[innerSortConfig.key] || "";
                    return dir * String(valA).localeCompare(String(valB));
                  });

                  return (
                    <div key={group.employeeName}>
                      {/* Employee Header Row */}
                      <div
                        onClick={() => toggleEmployee(group.employeeName)}
                        style={{
                          display: "flex",
                          alignItems: "center",
                          gap: 16,
                          padding: "14px 20px",
                          background: isExpanded ? "rgba(52,37,107,0.04)" : BRAND.white,
                          borderBottom: "1px solid #f0f0f0",
                          cursor: "pointer",
                          transition: "background 0.15s",
                          userSelect: "none",
                        }}
                        onMouseEnter={(e) => { if (!isExpanded) e.currentTarget.style.background = "rgba(245,206,0,0.06)"; }}
                        onMouseLeave={(e) => { e.currentTarget.style.background = isExpanded ? "rgba(52,37,107,0.04)" : BRAND.white; }}
                      >
                        {/* Chevron */}
                        <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke={BRAND.indigo} strokeWidth="2.5"
                          style={{ transition: "transform 0.2s", transform: isExpanded ? "rotate(90deg)" : "rotate(0deg)", flexShrink: 0 }}>
                          <path d="M9 18l6-6-6-6" strokeLinecap="round" strokeLinejoin="round"/>
                        </svg>

                        {/* Employee Name */}
                        <span style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 700, fontSize: 13, color: BRAND.black, minWidth: 200 }}>
                          {group.employeeName}
                        </span>

                        {/* Days count */}
                        <span style={{ padding: "3px 10px", borderRadius: 2, background: "rgba(52,37,107,0.08)", color: BRAND.indigo, fontSize: 11, fontWeight: 600, fontFamily: "'Montserrat', sans-serif", whiteSpace: "nowrap" }}>
                          {group.totalDays} {group.totalDays === 1 ? "day" : "days"}
                        </span>

                        {/* Mismatches badge */}
                        {group.mismatches > 0 && (
                          <span style={{ padding: "3px 10px", borderRadius: 2, background: "rgba(221,60,39,0.08)", color: BRAND.chilli, fontSize: 11, fontWeight: 700, fontFamily: "'Montserrat', sans-serif", whiteSpace: "nowrap" }}>
                            {group.mismatches} mismatch{group.mismatches !== 1 ? "es" : ""}
                          </span>
                        )}

                        {/* Work Setup */}
                        <span style={{ fontSize: 11, color: "#888", fontFamily: "'Montserrat', sans-serif", whiteSpace: "nowrap" }}>
                          {group.workSetup}
                        </span>

                        <div style={{ flex: 1 }} />

                        {/* Status indicator */}
                        {group.hasDiscrepancies ? (
                          <span style={{ color: BRAND.chilli, fontSize: 16, fontWeight: 700 }}>!</span>
                        ) : (
                          <span style={{ color: "#16a34a", fontSize: 14 }}>&#x2713;</span>
                        )}
                      </div>

                      {/* Expanded Sub-Table */}
                      <div
                        data-accordion-body=""
                        style={{
                          maxHeight: isExpanded ? `${sortedRows.length * 42 + 44}px` : "0px",
                          overflow: "hidden",
                          transition: "max-height 0.25s ease-in-out",
                        }}
                      >
                        <div style={{ overflowX: "auto", borderBottom: isExpanded ? `2px solid ${BRAND.indigo}` : "none" }}>
                          <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 12, fontFamily: "'Montserrat', sans-serif" }}>
                            <thead>
                              <tr style={{ background: "rgba(52,37,107,0.85)" }}>
                                {SUB_TABLE_COLUMNS.map((col) => (
                                  <th
                                    key={col.key}
                                    onClick={() => {
                                      setInnerSortConfig(prev =>
                                        prev.key === col.key
                                          ? { key: col.key, direction: prev.direction === "asc" ? "desc" : "asc" }
                                          : { key: col.key, direction: "asc" }
                                      );
                                    }}
                                    style={{
                                      padding: "10px 14px",
                                      textAlign: "left",
                                      fontWeight: 700,
                                      fontSize: 10,
                                      color: BRAND.white,
                                      textTransform: "uppercase",
                                      letterSpacing: 1.5,
                                      whiteSpace: "nowrap",
                                      borderBottom: `2px solid ${BRAND.canary}`,
                                      cursor: "pointer",
                                      userSelect: "none",
                                    }}
                                  >
                                    {col.label}
                                    {innerSortConfig.key === col.key && (
                                      <span style={{ marginLeft: 4, fontSize: 9 }}>
                                        {innerSortConfig.direction === "asc" ? "▲" : "▼"}
                                      </span>
                                    )}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody>
                              {sortedRows.map((r, i) => (
                                <tr
                                  key={i}
                                  style={{
                                    background: i % 2 === 0 ? BRAND.white : "#fafafa",
                                    borderBottom: "1px solid #f0f0f0",
                                    transition: "background 0.15s",
                                  }}
                                  onMouseEnter={(e) => (e.currentTarget.style.background = "rgba(245,206,0,0.06)")}
                                  onMouseLeave={(e) => (e.currentTarget.style.background = i % 2 === 0 ? BRAND.white : "#fafafa")}
                                >
                                  <td style={{ padding: "10px 14px", whiteSpace: "nowrap" }}>{r.date}</td>
                                  <td style={{ padding: "10px 14px", whiteSpace: "nowrap", fontSize: 11 }}>{r.dayOfWeek ? r.dayOfWeek.slice(0, 3) : "—"}</td>
                                  <td style={{ padding: "10px 14px", fontFamily: "monospace", fontSize: 11 }}>{r.clockIn}</td>
                                  <td style={{ padding: "10px 14px", fontFamily: "monospace", fontSize: 11 }}>{r.clockOut}</td>
                                  <td style={{ padding: "10px 14px", fontFamily: "monospace", fontSize: 11 }}>{r.ipIn}</td>
                                  <td style={{
                                    padding: "10px 14px",
                                    fontSize: 11,
                                    fontWeight: 600,
                                    color: r.ipLocation.startsWith("Office") ? "#16a34a" : r.ipLocation === "Remote" ? "#2563eb" : "#999",
                                  }}>
                                    {r.ipLocation}
                                  </td>
                                  <td style={{ padding: "10px 14px", fontSize: 11, color: r.resolvedLocation !== "—" ? BRAND.indigo : "#ccc" }}>{r.resolvedLocation}</td>
                                  <td style={{ padding: "10px 14px", fontSize: 11 }}>{r.workSetup}</td>
                                  <td style={{ padding: "10px 14px", fontSize: 11, fontWeight: 600 }}>{r.scheduledStatus}</td>
                                  <td style={{ padding: "10px 14px" }}>
                                    <StatusBadge status={r.auditStatus} />
                                  </td>
                                  <td style={{ padding: "10px 14px", fontSize: 11, color: r.auditStatus === "Mismatch" ? BRAND.chilli : "#999", maxWidth: 300, lineHeight: 1.4 }}>
                                    {r.discrepancies}
                                  </td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    </div>
                  );
                })
              )}
            </div>

            {/* Pagination */}
            <div
              data-pagination=""
              style={{
                display: "flex",
                alignItems: "center",
                justifyContent: "space-between",
                padding: "16px 20px",
                background: BRAND.white,
                borderTop: "1px solid #f0f0f0",
                borderRadius: "0 0 2px 2px",
                marginBottom: 20,
                boxShadow: "0 2px 8px rgba(0,0,0,0.04)",
              }}
            >
              <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, color: "#999", textTransform: "uppercase", letterSpacing: 1 }}>
                  Show
                </span>
                {[25, 50, 100, "all"].map(size => (
                  <button
                    key={size}
                    onClick={() => { setPageSize(size); setCurrentPage(1); }}
                    style={{
                      padding: "4px 10px",
                      background: pageSize === size ? BRAND.indigo : "transparent",
                      color: pageSize === size ? BRAND.white : "#666",
                      border: `1px solid ${pageSize === size ? BRAND.indigo : "#ddd"}`,
                      borderRadius: 2,
                      fontSize: 11,
                      fontFamily: "'Montserrat', sans-serif",
                      fontWeight: 600,
                      cursor: "pointer",
                    }}
                  >
                    {size === "all" ? "All" : size}
                  </button>
                ))}
                <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 11, color: "#999" }}>
                  employees per page
                </span>
              </div>

              {pageSize !== "all" && totalGroupPages > 1 && (
                <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
                  <button
                    disabled={currentPage === 1}
                    onClick={() => setCurrentPage(p => Math.max(1, p - 1))}
                    style={{
                      padding: "6px 12px", background: "transparent", border: "1px solid #ddd",
                      borderRadius: 2, fontSize: 11, cursor: currentPage === 1 ? "not-allowed" : "pointer",
                      opacity: currentPage === 1 ? 0.4 : 1, fontFamily: "'Montserrat', sans-serif", fontWeight: 600,
                    }}
                  >
                    Prev
                  </button>
                  <span style={{ fontFamily: "'Montserrat', sans-serif", fontSize: 12, color: BRAND.indigo, fontWeight: 600 }}>
                    {currentPage} / {totalGroupPages}
                  </span>
                  <button
                    disabled={currentPage === totalGroupPages}
                    onClick={() => setCurrentPage(p => Math.min(totalGroupPages, p + 1))}
                    style={{
                      padding: "6px 12px", background: "transparent", border: "1px solid #ddd",
                      borderRadius: 2, fontSize: 11, cursor: currentPage === totalGroupPages ? "not-allowed" : "pointer",
                      opacity: currentPage === totalGroupPages ? 0.4 : 1, fontFamily: "'Montserrat', sans-serif", fontWeight: 600,
                    }}
                  >
                    Next
                  </button>
                </div>
              )}
            </div>

            {/* Back / New Scan */}
            <div style={{ marginTop: 24, display: "flex", gap: 12 }}>
              <button
                onClick={() => {
                  setResults(null);
                  setStep(1);
                  setSproutData(null);
                  setSproutFile("");
                  setScheduleFile("");
                  setSproutMap({});
                  setFilter("all");
                  setSearchTerm("");
                  setDateFilter("");
                  setIpGeoCache({});
                  setIpResolutionStatus("idle");
                  setIpResolutionProgress({ current: 0, total: 0 });
                  setIpResolutionErrors([]);
                  setProcessedSprout(null);
                  setScheduleLookup(null);
                  setPreprocessSummary(null);
                  setExpandedEmployees(new Set());
                  setAllExpanded(false);
                  setPageSize(25);
                  setCurrentPage(1);
                  setInnerSortConfig({ key: "date", direction: "asc" });
                  setOfficeIPs([
                    { label: "EST", ipFrom: "116.50.227.176", ipTo: "116.50.227.181" },
                    { label: "ComClark", ipFrom: "161.49.192.112", ipTo: "161.49.192.116" },
                    { label: "ComClark 7F", ipFrom: "136.239.243.241", ipTo: "136.239.243.246" },
                  ]);
                }}
                style={{
                  padding: "12px 28px",
                  background: "transparent",
                  color: BRAND.indigo,
                  border: `2px solid ${BRAND.indigo}`,
                  borderRadius: 2,
                  fontSize: 12,
                  fontFamily: "'Montserrat', sans-serif",
                  fontWeight: 700,
                  textTransform: "uppercase",
                  letterSpacing: 1.5,
                  cursor: "pointer",
                }}
              >
                + New Scan
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Footer */}
      <footer
        style={{
          background: BRAND.indigo,
          padding: "20px 32px",
          marginTop: 60,
          display: "flex",
          justifyContent: "space-between",
          alignItems: "center",
        }}
      >
        <BrandLogo inverted />
        <span style={{ fontFamily: "'Montserrat', sans-serif", fontWeight: 300, fontSize: 11, color: "rgba(255,255,255,0.4)" }}>
          © {new Date().getFullYear()} The Back Room. Timesheet Audit System.
        </span>
      </footer>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @media print {
          nav, footer, button { display: none !important; }
          div[style*="boxShadow"] { box-shadow: none !important; }
          [data-accordion-body] { max-height: none !important; overflow: visible !important; }
          [data-pagination] { display: none !important; }
        }
        table { page-break-inside: auto; }
        tr { page-break-inside: avoid; }
      `}</style>
    </div>
  );
}

// ─── App Root ───
export default function App() {
  const [user, setUser] = useState(null);

  if (!user) {
    return <LoginPage onLogin={setUser} />;
  }

  return <Dashboard user={user} onLogout={() => setUser(null)} />;
}
