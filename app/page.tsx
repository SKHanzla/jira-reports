"use client";

import { useState, useEffect } from "react";
import * as XLSX from "xlsx";

interface ReportRow {
  epic_key: string;
  epic_order: number;
  epic_name: string;
  story_name: string;
  project_key: string;
  labels: string;
  project_name: string;
  original_estimate: number;
  total_hours_logged: number;
  hrs_left: number;
}

interface ApiResponse {
  rows?: ReportRow[];
  epicCount?: number;
  storyCount?: number;
  error?: string;
}

interface UserReportRow {
  user: string;
  epic_name: string;
  story_name: string;
  project_key: string;
  label: string;
  project_name: string;
  status: string;
  hours: number;
}

interface UserApiResponse {
  rows?: UserReportRow[];
  issueCount?: number;
  error?: string;
}

interface ProjectReportRow {
  project_name: string;
  project_key: string;
  epic_name: string;
  story_name: string;
  label: string;
  status: string;
  original_estimate: number;
  hours_logged: number;
  hours_left: number;
}

interface ProjectApiResponse {
  rows?: ProjectReportRow[];
  issueCount?: number;
  error?: string;
}

type Status = "idle" | "loading" | "done" | "error";

function isoDate(d: Date): string {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
}

function parseLocalDate(s: string): Date {
  const [y, m, d] = s.split("-").map(Number);
  return new Date(y, (m || 1) - 1, d || 1);
}

function getDefaultRange(): { from: string; to: string } {
  const now = new Date();
  const first = new Date(now.getFullYear(), now.getMonth(), 1);
  const last = new Date(now.getFullYear(), now.getMonth() + 1, 0);
  return { from: isoDate(first), to: isoDate(last) };
}

function getReportPeriod(fromISO: string, toISO: string) {
  const from = parseLocalDate(fromISO);
  const to = parseLocalDate(toISO);
  const fmt = (d: Date) =>
    d.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    });
  const month = from.toLocaleString("default", { month: "long" });
  const year = from.getFullYear();
  return { period: `${fmt(from)} to ${fmt(to)}`, month, year };
}

function LoginPage({ onLogin }: { onLogin: () => void }) {
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPass, setShowPass] = useState(false);

  const purple = "#7c3aab";
  const purpleLight = "#f5eefa";
  const purpleMid = "#a855b5";

  async function handleSubmit(e: React.FormEvent) {
    e.preventDefault();
    setError("");
    const res = await fetch("/api/auth", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ username, password }),
    });
    if (res.ok) {
      sessionStorage.setItem("auth", "1");
      onLogin();
    } else {
      setError("Invalid username or password.");
    }
  }

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#faf9fc",
        fontFamily: "ui-sans-serif, system-ui, sans-serif",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
      }}
    >
      <div
        style={{
          width: "100%",
          maxWidth: 400,
          background: "#fff",
          borderRadius: 16,
          border: "1px solid #e9d5f7",
          boxShadow: "0 4px 32px 0 rgba(124,58,171,0.08)",
          padding: "40px 36px",
        }}
      >
        {/* Logo */}
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            marginBottom: 28,
          }}
        >
          <svg
            width="64"
            height="40"
            viewBox="0 0 72 44"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
            style={{ marginBottom: 8 }}
          >
            <path
              d="M36 38 C30 30, 14 26, 4 14 C10 18, 18 20, 24 24 C18 16, 10 10, 6 4 C14 10, 22 16, 28 22 C24 14, 20 8, 18 2 C24 10, 28 18, 32 26 C30 18, 30 10, 30 4 C32 12, 34 20, 36 28"
              fill="#a855b5"
            />
            <path
              d="M36 38 C42 30, 58 26, 68 14 C62 18, 54 20, 48 24 C54 16, 62 10, 66 4 C58 10, 50 16, 44 22 C48 14, 52 8, 54 2 C48 10, 44 18, 40 26 C42 18, 42 10, 42 4 C40 12, 38 20, 36 28"
              fill="#7c3aab"
            />
            <line
              x1="36"
              y1="28"
              x2="36"
              y2="42"
              stroke="#a855b5"
              strokeWidth="2.5"
              strokeLinecap="round"
            />
          </svg>
          <div
            style={{
              display: "flex",
              alignItems: "baseline",
              letterSpacing: "0.18em",
              fontSize: 15,
              textTransform: "uppercase",
              marginBottom: 4,
            }}
          >
            <span style={{ fontWeight: 400, color: "#6b7280" }}>EXIT</span>
            <span style={{ fontWeight: 700, color: "#1c1917" }}>BLISS</span>
          </div>
          <div style={{ width: 40, height: 2, background: purpleMid }} />
        </div>

        <h2
          style={{
            fontSize: 20,
            fontWeight: 700,
            color: "#1c1917",
            marginBottom: 4,
            textAlign: "center",
          }}
        >
          Sign in
        </h2>
        <p
          style={{
            fontSize: 13,
            color: "#78716c",
            textAlign: "center",
            marginBottom: 24,
          }}
        >
          Exitbliss Report Portal
        </p>

        <form
          onSubmit={handleSubmit}
          style={{ display: "flex", flexDirection: "column", gap: 16 }}
        >
          <div>
            <label
              style={{
                display: "block",
                fontSize: 12,
                fontWeight: 600,
                color: "#374151",
                marginBottom: 6,
                letterSpacing: "0.04em",
              }}
            >
              USERNAME
            </label>
            <input
              type="text"
              value={username}
              onChange={(e) => {
                setUsername(e.target.value);
                setError("");
              }}
              autoComplete="username"
              placeholder="Enter Your Username"
              style={{
                width: "100%",
                boxSizing: "border-box",
                border: `1.5px solid ${error ? "#fca5a5" : "#e5e7eb"}`,
                borderRadius: 8,
                padding: "10px 14px",
                fontSize: 14,
                color: "#1c1917",
                outline: "none",
                background: "#fafafa",
                transition: "border 0.15s",
              }}
            />
          </div>

          <div>
            <label
              style={{
                display: "block",
                fontSize: 12,
                fontWeight: 600,
                color: "#374151",
                marginBottom: 6,
                letterSpacing: "0.04em",
              }}
            >
              PASSWORD
            </label>
            <div style={{ position: "relative" }}>
              <input
                type={showPass ? "text" : "password"}
                value={password}
                onChange={(e) => {
                  setPassword(e.target.value);
                  setError("");
                }}
                autoComplete="current-password"
                placeholder="Enter Your Password"
                style={{
                  width: "100%",
                  boxSizing: "border-box",
                  border: `1.5px solid ${error ? "#fca5a5" : "#e5e7eb"}`,
                  borderRadius: 8,
                  padding: "10px 40px 10px 14px",
                  fontSize: 14,
                  color: "#1c1917",
                  outline: "none",
                  background: "#fafafa",
                  transition: "border 0.15s",
                }}
              />
              <button
                type="button"
                onClick={() => setShowPass((v) => !v)}
                style={{
                  position: "absolute",
                  right: 12,
                  top: "50%",
                  transform: "translateY(-50%)",
                  background: "none",
                  border: "none",
                  cursor: "pointer",
                  color: "#9ca3af",
                  padding: 0,
                }}
                tabIndex={-1}
              >
                {showPass ? (
                  <svg
                    width="16"
                    height="16"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858.908a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21"
                    />
                  </svg>
                ) : (
                  <svg
                    width="16"
                    height="16"
                    fill="none"
                    stroke="currentColor"
                    strokeWidth="2"
                    viewBox="0 0 24 24"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M15 12a3 3 0 11-6 0 3 3 0 016 0z"
                    />
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z"
                    />
                  </svg>
                )}
              </button>
            </div>
          </div>

          {error && (
            <div
              style={{
                background: "#fef2f2",
                border: "1px solid #fca5a5",
                borderRadius: 8,
                padding: "10px 14px",
                fontSize: 13,
                color: "#b91c1c",
                fontWeight: 500,
              }}
            >
              {error}
            </div>
          )}

          <button
            type="submit"
            style={{
              background: purple,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "11px 0",
              fontSize: 14,
              fontWeight: 600,
              cursor: "pointer",
              marginTop: 4,
              letterSpacing: "0.02em",
              transition: "background 0.15s",
            }}
          >
            Sign In
          </button>
        </form>

        <p
          style={{
            fontSize: 11,
            color: "#d1d5db",
            textAlign: "center",
            marginTop: 24,
          }}
        >
          Session expires on page reload
        </p>
      </div>
    </main>
  );
}

function getUserReportPeriod() {
  const now = new Date();
  const start = new Date(now);
  start.setDate(now.getDate() - 30);
  const fmt = (d: Date) =>
    d.toLocaleDateString("en-US", {
      month: "long",
      day: "numeric",
      year: "numeric",
    });
  return fmt(start) + " to " + fmt(now);
}

export default function Page() {
  const [authed, setAuthed] = useState<boolean | null>(null);
  const [status, setStatus] = useState<Status>("idle");
  const [log, setLog] = useState<string[]>([]);
  const [result, setResult] = useState<ApiResponse | null>(null);
  const [fromDate, setFromDate] = useState(() => getDefaultRange().from);
  const [toDate, setToDate] = useState(() => getDefaultRange().to);
  const { period, month, year } = getReportPeriod(fromDate, toDate);

  const [userStatus, setUserStatus] = useState<Status>("idle");
  const [userLog, setUserLog] = useState<string[]>([]);
  const [userResult, setUserResult] = useState<UserApiResponse | null>(null);
  const userPeriod = getUserReportPeriod();

  const [projStatus, setProjStatus] = useState<Status>("idle");
  const [projLog, setProjLog] = useState<string[]>([]);
  const [projResult, setProjResult] = useState<ProjectApiResponse | null>(null);

  useEffect(() => {
    setAuthed(sessionStorage.getItem("auth") === "1");
  }, []);

  if (authed === null) return null; // avoid flash before sessionStorage is read
  if (!authed) return <LoginPage onLogin={() => setAuthed(true)} />;

  function handleLogout() {
    sessionStorage.removeItem("auth");
    setAuthed(false);
  }

  function addLog(msg: string) {
    setLog((prev) => [...prev, msg]);
  }

  // Changing the range invalidates a previously generated report so the
  // invoice / XLSX can never be downloaded for a stale date range.
  function handleRangeChange(which: "from" | "to", value: string) {
    if (which === "from") setFromDate(value);
    else setToDate(value);
    if (status !== "idle") {
      setStatus("idle");
      setResult(null);
      setLog([]);
    }
  }

  async function handleGenerate() {
    if (fromDate > toDate) {
      setStatus("error");
      setLog(["Error: start date must be on or before the end date."]);
      return;
    }
    setStatus("loading");
    setLog([]);
    setResult(null);
    addLog("Connecting to Jira...");
    addLog(`Reporting period: ${period}`);
    addLog("Fetching CDEF Epics...");
    try {
      const res = await fetch(
        `/api/jira?from=${encodeURIComponent(fromDate)}&to=${encodeURIComponent(toDate)}`,
      );
      addLog("Processing stories and fetching worklogs...");
      const data: ApiResponse = await res.json();
      if (!res.ok || data.error) {
        addLog(`Error: ${data.error}`);
        setStatus("error");
        setResult(data);
        return;
      }
      addLog(`Found ${data.epicCount} epics and ${data.storyCount} stories.`);
      addLog("Report ready.");
      setResult(data);
      setStatus("done");
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Network error";
      addLog(`Error: ${msg}`);
      setStatus("error");
    }
  }

  function handleDownload() {
    if (!result?.rows?.length) return;

    const wb = XLSX.utils.book_new();
    const ws: XLSX.WorkSheet = {};

    // Row 1: Report Name
    ws["A1"] = { v: `Report Name: Active CDEF Monthly Report`, t: "s" };
    // Row 2: Report Period
    ws["A2"] = { v: `Report Period: ${period}`, t: "s" };

    // Row 7: Headers
    const headers = [
      "Epic Name",
      "Story Name",
      "Project key",
      "Labels",
      "Project name",
      "Original estimate",
      "Total Hours logged",
      "hrs Left",
    ];
    const cols = ["A", "B", "C", "D", "E", "F", "G", "H"];
    headers.forEach((h, i) => {
      ws[`${cols[i]}7`] = { v: h, t: "s" };
    });

    // Group rows by project_key to insert blank rows between projects
    const epicNum = (key: string) => parseInt(key?.split("-")[1] ?? "0", 10);
    const sorted = [...result.rows].sort(
      (a, b) =>
        a.project_key.localeCompare(b.project_key) ||
        epicNum(a.epic_key) - epicNum(b.epic_key) ||
        a.story_name.localeCompare(b.story_name),
    );

    let currentRow = 8;
    let currentProject = "";

    sorted.forEach((r) => {
      if (currentProject && currentProject !== r.project_key) {
        currentRow++; // blank separator row between projects
      }
      currentProject = r.project_key;

      ws[`A${currentRow}`] = { v: r.epic_name, t: "s" };
      ws[`B${currentRow}`] = { v: r.story_name, t: "s" };
      ws[`C${currentRow}`] = { v: r.project_key, t: "s" };
      ws[`D${currentRow}`] = { v: r.labels, t: "s" };
      ws[`E${currentRow}`] = { v: r.project_name, t: "s" };
      ws[`F${currentRow}`] = { v: r.original_estimate, t: "n" };
      ws[`G${currentRow}`] = { v: r.total_hours_logged, t: "n" };
      ws[`H${currentRow}`] = { v: r.hrs_left, t: "n" };
      currentRow++;
    });

    // Set sheet ref range
    ws["!ref"] = `A1:H${currentRow - 1}`;

    // Column widths matching original
    ws["!cols"] = [
      { wch: 50 }, // A - Epic Name
      { wch: 50 }, // B - Story Name
      { wch: 14 }, // C - Project key
      { wch: 9 }, // D - Labels
      { wch: 31 }, // E - Project name
      { wch: 20 }, // F - Original estimate
      { wch: 21 }, // G - Total Hours logged
      { wch: 11 }, // H - hrs Left
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Time Report");
    XLSX.writeFile(wb, `Active_CDEF_Monthly_Report_${month}_${year}.xlsx`);
  }

  function buildInvoiceCSV(): { csv: string; filename: string } {
    const PRICE_LIST: Record<string, number> = {
      "Standard Operating Procedures": 275,
      "Business Operations and Project Management": 250,
      "Administrative, Finance and Compliance": 250,
      "Administrative Charges": 250,
      "Governance and Policy": 250,
      "Productivity System Set-Up": 350,
      "Organizational Assessment & Audit": 300,
      "Documentation Creation": 250,
      "Vendor Management": 250,
      "Documentation Management": 275,
      "Leadership and Executive Coaching": 250,
      "Organizational Change Management": 275,
      "Succession Planning": 275,
      "Strategic Planning": 250,
      "Business Continuity": 250,
      "Process Improvement/Computer Training": 275,
      "Marketing and Strategic Communications": 250,
    };

    const now = new Date();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const mmdd = `${mm}${dd}`;
    const invoiceDate = `${now.getMonth() + 1}/${now.getDate()}/${now.getFullYear()}`;
    const due = new Date(now);
    due.setDate(due.getDate() + 60);
    const dueStr = `${due.getMonth() + 1}/${due.getDate()}/${due.getFullYear()}`;

    // Date of service is driven by the selected report range.
    const svcStart = parseLocalDate(fromDate);
    const svcEnd = parseLocalDate(toDate);
    const svcMonthName = svcStart.toLocaleString("en-US", { month: "long" });
    const svcYear = svcStart.getFullYear();
    const sameMonth =
      svcStart.getFullYear() === svcEnd.getFullYear() &&
      svcStart.getMonth() === svcEnd.getMonth();
    const fullFmt = (d: Date) =>
      d.toLocaleDateString("en-US", {
        month: "long",
        day: "numeric",
        year: "numeric",
      });
    const dateOfService = sameMonth
      ? `${svcMonthName} ${svcStart.getDate()} - ${svcEnd.getDate()}, ${svcYear}`
      : `${fullFmt(svcStart)} - ${fullFmt(svcEnd)}`;

    const byProject: Record<string, ReportRow[]> = {};
    for (const row of result!.rows!) {
      if (!byProject[row.project_key]) byProject[row.project_key] = [];
      byProject[row.project_key].push(row);
    }

    const esc = (v: string) =>
      v.includes(",") || v.includes('"') || v.includes("\n")
        ? `"${v.replace(/"/g, '""')}"`
        : v;

    const lines = [
      "Invoice No,Customer,Invoice Date,Due Date,Terms,Item (Product/Service),Item Description,Item Quantity,Item Rate,Item Amount,Memo,Note",
    ];
    const epicNum = (k: string) => parseInt(k?.split("-")[1] ?? "0", 10);
    let seq = 1;

    for (const projKey of Object.keys(byProject).sort()) {
      const rows = byProject[projKey];
      const projName = rows[0].project_name;
      const invoiceNo = `CDEF-${projKey}-${mmdd}-${String(seq).padStart(2, "0")}`;
      seq++;
      const memo = `This invoice covers;\nCONTRACT #: DCM-SVCSGEN-17106-2025\nCLIENT NAME: ${projName}\nDATE OF SERVICE: ${dateOfService}\n\nThank you for your business`;
      const sorted = [...rows].sort(
        (a, b) =>
          epicNum(a.epic_key) - epicNum(b.epic_key) ||
          a.story_name.localeCompare(b.story_name),
      );

      let first = true;
      for (const r of sorted) {
        const rate = PRICE_LIST[r.epic_name] ?? 250;
        const qty = r.total_hours_logged;
        const amount = Math.round(qty * rate * 100) / 100;
        if (first) {
          lines.push(
            [
              invoiceNo,
              "Construction Diversity & Equity Fund - Multnomah County",
              invoiceDate,
              dueStr,
              "Net 60",
              esc(r.epic_name),
              esc(r.story_name),
              String(qty),
              String(rate),
              String(amount),
              esc(memo),
              esc(memo),
            ].join(","),
          );
          first = false;
        } else {
          lines.push(
            [
              invoiceNo,
              "",
              "",
              "",
              "",
              esc(r.epic_name),
              esc(r.story_name),
              String(qty),
              String(rate),
              String(amount),
              "",
              "",
            ].join(","),
          );
        }
      }
    }

    return {
      csv: lines.join("\n"),
      filename: `${svcMonthName}_${svcYear}_${now.getMonth() + 1}-${now.getDate()}-${now.getFullYear()}.csv`,
    };
  }

  async function handleGenerateInvoice() {
    if (!result?.rows?.length) return;
    const { csv, filename } = buildInvoiceCSV();
    const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
    try {
      if ("showSaveFilePicker" in window) {
        const handle = await (
          window as Window & {
            showSaveFilePicker: (o: unknown) => Promise<FileSystemFileHandle>;
          }
        ).showSaveFilePicker({
          suggestedName: filename,
          types: [
            { description: "CSV file", accept: { "text/csv": [".csv"] } },
          ],
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
      } else {
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = filename;
        a.click();
        URL.revokeObjectURL(url);
      }
    } catch (e) {
      if ((e as Error).name !== "AbortError") console.error(e);
    }
  }

  async function handleGenerateUser() {
    setUserStatus("loading");
    setUserLog([]);
    setUserResult(null);
    setUserLog((p) => [...p, "Connecting to Jira..."]);
    setUserLog((p) => [
      ...p,
      "Fetching issues with worklogs (last 30 days)...",
    ]);
    try {
      const res = await fetch("/api/jira-user");
      setUserLog((p) => [...p, "Processing worklogs by user..."]);
      const data: UserApiResponse = await res.json();
      if (!res.ok || data.error) {
        setUserLog((p) => [...p, `Error: ${data.error}`]);
        setUserStatus("error");
        setUserResult(data);
        return;
      }
      setUserLog((p) => [
        ...p,
        `Found ${data.issueCount} issues · ${data.rows?.length ?? 0} worklog entries.`,
      ]);
      setUserLog((p) => [...p, "Report ready."]);
      setUserResult(data);
      setUserStatus("done");
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Network error";
      setUserLog((p) => [...p, `Error: ${msg}`]);
      setUserStatus("error");
    }
  }

  function handleDownloadUser() {
    if (!userResult?.rows?.length) return;

    const wb = XLSX.utils.book_new();
    const ws: XLSX.WorkSheet = {};

    ws["A1"] = { v: "Report Name: Time Tracking Report By User", t: "s" };
    ws["A2"] = { v: `Report Period: ${userPeriod}`, t: "s" };

    const headers = [
      "Epic Name",
      "Story Name",
      "Project key",
      "Labels",
      "Project name",
      "Status",
      "Total Hours logged",
    ];
    const cols = ["A", "B", "C", "D", "E", "F", "G"];
    headers.forEach((h, i) => {
      ws[`${cols[i]}7`] = { v: h, t: "s" };
    });

    const byUser: Record<string, UserReportRow[]> = {};
    for (const row of userResult.rows) {
      if (!byUser[row.user]) byUser[row.user] = [];
      byUser[row.user].push(row);
    }

    let cur = 8;
    for (const user of Object.keys(byUser).sort()) {
      ws[`A${cur}`] = { v: `User: ${user}`, t: "s" };
      cur++;
      const sorted = byUser[user].sort(
        (a, b) =>
          a.epic_name.localeCompare(b.epic_name) ||
          a.story_name.localeCompare(b.story_name),
      );
      for (const r of sorted) {
        ws[`A${cur}`] = { v: r.epic_name, t: "s" };
        ws[`B${cur}`] = { v: r.story_name, t: "s" };
        ws[`C${cur}`] = { v: r.project_key, t: "s" };
        ws[`D${cur}`] = { v: r.label, t: "s" };
        ws[`E${cur}`] = { v: r.project_name, t: "s" };
        ws[`F${cur}`] = { v: r.status, t: "s" };
        ws[`G${cur}`] = { v: r.hours, t: "n" };
        cur++;
      }
      cur++;
    }

    ws["!ref"] = `A1:G${cur - 1}`;
    ws["!cols"] = [
      { wch: 40 },
      { wch: 70 },
      { wch: 14 },
      { wch: 20 },
      { wch: 30 },
      { wch: 18 },
      { wch: 20 },
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Time Report");
    const now = new Date();
    XLSX.writeFile(
      wb,
      `User_Time_Report_Last30Days_${now.getFullYear()}_${String(now.getMonth() + 1).padStart(2, "0")}.xlsx`,
    );
  }

  async function handleGenerateProject() {
    setProjStatus("loading");
    setProjLog([]);
    setProjResult(null);
    setProjLog((p) => [...p, "Connecting to Jira..."]);
    setProjLog((p) => [
      ...p,
      "Fetching issues with worklogs (last 30 days)...",
    ]);
    try {
      const res = await fetch("/api/jira-project");
      setProjLog((p) => [...p, "Processing worklogs by project..."]);
      const data: ProjectApiResponse = await res.json();
      if (!res.ok || data.error) {
        setProjLog((p) => [...p, `Error: ${data.error}`]);
        setProjStatus("error");
        setProjResult(data);
        return;
      }
      setProjLog((p) => [
        ...p,
        `Found ${data.issueCount} issues · ${data.rows?.length ?? 0} task entries.`,
      ]);
      setProjLog((p) => [...p, "Report ready."]);
      setProjResult(data);
      setProjStatus("done");
    } catch (err) {
      const msg = err instanceof Error ? err.message : "Network error";
      setProjLog((p) => [...p, `Error: ${msg}`]);
      setProjStatus("error");
    }
  }

  function handleDownloadProject() {
    if (!projResult?.rows?.length) return;

    const wb = XLSX.utils.book_new();
    const ws: XLSX.WorkSheet = {};

    ws["A1"] = { v: "Report Name: Time Tracking Report By Project", t: "s" };
    ws["A2"] = { v: `Report Period: ${userPeriod}`, t: "s" };

    const headers = [
      "Epic Name",
      "Story Name",
      "Project key",
      "Labels",
      "Project name",
      "Status",
      "Original estimate",
      "Total Hours logged",
      "hrs Left",
    ];
    ["A", "B", "C", "D", "E", "F", "G", "H", "I"].forEach((col, i) => {
      ws[`${col}7`] = { v: headers[i], t: "s" };
    });

    const byProject: Record<string, ProjectReportRow[]> = {};
    for (const row of projResult.rows) {
      if (!byProject[row.project_name]) byProject[row.project_name] = [];
      byProject[row.project_name].push(row);
    }

    let cur = 8;
    for (const proj of Object.keys(byProject).sort()) {
      ws[`A${cur}`] = { v: `Project: ${proj}`, t: "s" };
      cur++;
      const sorted = byProject[proj].sort(
        (a, b) =>
          a.epic_name.localeCompare(b.epic_name) ||
          a.story_name.localeCompare(b.story_name),
      );
      for (const r of sorted) {
        ws[`A${cur}`] = { v: r.epic_name, t: "s" };
        ws[`B${cur}`] = { v: r.story_name, t: "s" };
        ws[`C${cur}`] = { v: r.project_key, t: "s" };
        ws[`D${cur}`] = { v: r.label, t: "s" };
        ws[`E${cur}`] = { v: r.project_name, t: "s" };
        ws[`F${cur}`] = { v: r.status, t: "s" };
        ws[`G${cur}`] = { v: r.original_estimate, t: "n" };
        ws[`H${cur}`] = { v: r.hours_logged, t: "n" };
        ws[`I${cur}`] = { v: r.hours_left, t: "n" };
        cur++;
      }
      cur++;
    }

    ws["!ref"] = `A1:I${cur - 1}`;
    ws["!cols"] = [
      { wch: 51 },
      { wch: 51 },
      { wch: 15 },
      { wch: 10 },
      { wch: 32 },
      { wch: 18 },
      { wch: 21 },
      { wch: 22 },
      { wch: 12 },
    ];

    XLSX.utils.book_append_sheet(wb, ws, "Time Report");
    const now = new Date();
    XLSX.writeFile(
      wb,
      `Project_Time_Report_Last30Days_${now.getFullYear()}_${String(now.getMonth() + 1).padStart(2, "0")}.xlsx`,
    );
  }

  const totalEst =
    result?.rows?.reduce((s, r) => s + r.original_estimate, 0) ?? 0;
  const totalLogged =
    result?.rows?.reduce((s, r) => s + r.total_hours_logged, 0) ?? 0;
  const totalLeft = result?.rows?.reduce((s, r) => s + r.hrs_left, 0) ?? 0;

  // Exit Bliss colors
  const purple = "#7c3aab";
  const purpleLight = "#f5eefa";
  const purpleMid = "#a855b5";

  return (
    <main
      style={{
        minHeight: "100vh",
        background: "#faf9fc",
        fontFamily: "ui-sans-serif, system-ui, sans-serif",
      }}
    >
      <div style={{ maxWidth: 820, margin: "0 auto", padding: "40px 24px" }}>
        {/* Exit Bliss Logo Header */}
        <div
          style={{
            display: "flex",
            flexDirection: "column",
            alignItems: "center",
            marginBottom: 36,
            position: "relative",
          }}
        >
          <svg
            width="80"
            height="50"
            viewBox="0 0 72 44"
            fill="none"
            xmlns="http://www.w3.org/2000/svg"
            style={{ marginBottom: 8 }}
          >
            <path
              d="M36 38 C30 30, 14 26, 4 14 C10 18, 18 20, 24 24 C18 16, 10 10, 6 4 C14 10, 22 16, 28 22 C24 14, 20 8, 18 2 C24 10, 28 18, 32 26 C30 18, 30 10, 30 4 C32 12, 34 20, 36 28"
              fill="#a855b5"
            />
            <path
              d="M36 38 C42 30, 58 26, 68 14 C62 18, 54 20, 48 24 C54 16, 62 10, 66 4 C58 10, 50 16, 44 22 C48 14, 52 8, 54 2 C48 10, 44 18, 40 26 C42 18, 42 10, 42 4 C40 12, 38 20, 36 28"
              fill="#7c3aab"
            />
            <line
              x1="36"
              y1="28"
              x2="36"
              y2="42"
              stroke="#a855b5"
              strokeWidth="2.5"
              strokeLinecap="round"
            />
          </svg>
          <div
            style={{
              display: "flex",
              alignItems: "baseline",
              letterSpacing: "0.18em",
              fontSize: 15,
              textTransform: "uppercase",
            }}
          >
            <span style={{ fontWeight: 400, color: "#6b7280" }}>EXIT</span>
            <span style={{ fontWeight: 700, color: "#1c1917" }}>BLISS</span>
          </div>
          <div
            style={{
              width: 48,
              height: 2,
              background: purpleMid,
              marginTop: 4,
            }}
          />

          {/* Logout button */}
          <button
            onClick={handleLogout}
            style={{
              position: "absolute",
              top: 0,
              right: 0,
              display: "flex",
              alignItems: "center",
              gap: 6,
              background: "none",
              border: "1.5px solid #e9d5f7",
              borderRadius: 8,
              padding: "7px 14px",
              fontSize: 12,
              fontWeight: 600,
              color: "#7c3aab",
              cursor: "pointer",
              transition: "background 0.15s",
            }}
            onMouseEnter={(e) => (e.currentTarget.style.background = "#f5eefa")}
            onMouseLeave={(e) => (e.currentTarget.style.background = "none")}
          >
            <svg
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1"
              />
            </svg>
            Logout
          </button>
        </div>

        {/* Page header */}
        <div style={{ marginBottom: 28 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              background: purpleLight,
              color: purple,
              fontSize: 11,
              fontWeight: 500,
              letterSpacing: "0.1em",
              textTransform: "uppercase",
              padding: "4px 12px",
              borderRadius: 99,
              marginBottom: 12,
            }}
          >
            <div
              style={{
                width: 6,
                height: 6,
                borderRadius: "50%",
                background: purpleMid,
              }}
            />
            poshplug.atlassian.net
          </div>
          <h1
            style={{
              fontSize: 26,
              fontWeight: 700,
              color: "#1c1917",
              marginBottom: 4,
              letterSpacing: "-0.01em",
            }}
          >
            Active CDEF Monthly Report
          </h1>
          <p style={{ fontSize: 13, color: "#78716c", marginBottom: 14 }}>
            {period}
          </p>

          {/* Date range selector */}
          <div
            style={{
              display: "flex",
              alignItems: "flex-end",
              gap: 16,
              flexWrap: "wrap",
            }}
          >
            <label
              style={{
                display: "flex",
                flexDirection: "column",
                gap: 4,
                fontSize: 11,
                fontWeight: 600,
                letterSpacing: "0.04em",
                textTransform: "uppercase",
                color: "#78716c",
              }}
            >
              From
              <input
                type="date"
                value={fromDate}
                max={toDate}
                onChange={(e) => handleRangeChange("from", e.target.value)}
                style={{
                  fontSize: 13,
                  padding: "8px 10px",
                  borderRadius: 8,
                  border: "1.5px solid #e7e5e4",
                  color: "#1c1917",
                  background: "#fff",
                  fontFamily: "inherit",
                }}
              />
            </label>
            <label
              style={{
                display: "flex",
                flexDirection: "column",
                gap: 4,
                fontSize: 11,
                fontWeight: 600,
                letterSpacing: "0.04em",
                textTransform: "uppercase",
                color: "#78716c",
              }}
            >
              To
              <input
                type="date"
                value={toDate}
                min={fromDate}
                onChange={(e) => handleRangeChange("to", e.target.value)}
                style={{
                  fontSize: 13,
                  padding: "8px 10px",
                  borderRadius: 8,
                  border: "1.5px solid #e7e5e4",
                  color: "#1c1917",
                  background: "#fff",
                  fontFamily: "inherit",
                }}
              />
            </label>
          </div>
        </div>

        {/* Divider */}
        <div style={{ height: 1, background: "#ede9fe", marginBottom: 24 }} />

        {/* Action row */}
        <div
          style={{
            display: "flex",
            gap: 12,
            marginBottom: 24,
            alignItems: "center",
          }}
        >
          <button
            onClick={handleGenerate}
            disabled={status === "loading"}
            style={{
              background: status === "loading" ? "#c4b5d4" : purple,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: status === "loading" ? "not-allowed" : "pointer",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "background 0.15s",
            }}
          >
            {status === "loading" && (
              <span
                style={{
                  width: 12,
                  height: 12,
                  borderRadius: "50%",
                  border: "2px solid rgba(255,255,255,0.35)",
                  borderTopColor: "#fff",
                  display: "inline-block",
                  animation: "spin 0.7s linear infinite",
                }}
              />
            )}
            {status === "loading" ? "Generating..." : "Generate Report"}
          </button>

          <button
            onClick={handleGenerateInvoice}
            disabled={status !== "done"}
            style={{
              background: status === "done" ? "#059669" : "#f3f4f6",
              color: status === "done" ? "#fff" : "#9ca3af",
              border: "none",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: status === "done" ? "pointer" : "not-allowed",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "background 0.15s",
            }}
          >
            <svg
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
              />
            </svg>
            Download For Quickbooks
          </button>

          <button
            onClick={handleDownload}
            disabled={status !== "done"}
            style={{
              background: status === "done" ? purpleLight : "#f3f4f6",
              color: status === "done" ? purple : "#9ca3af",
              border:
                status === "done"
                  ? `1.5px solid #c084d4`
                  : "1.5px solid #e5e7eb",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: status === "done" ? "pointer" : "not-allowed",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "all 0.15s",
            }}
          >
            <svg
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
              />
            </svg>
            Download XLSX
          </button>

          {/* Status indicator */}
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div
              style={{
                width: 7,
                height: 7,
                borderRadius: "50%",
                background:
                  status === "done"
                    ? "#639922"
                    : status === "error"
                      ? "#E24B4A"
                      : status === "loading"
                        ? purpleMid
                        : "#d1d5db",
                animation: status === "loading" ? "pulse 1s infinite" : "none",
              }}
            />
            <span style={{ fontSize: 12, color: "#78716c" }}>
              {status === "idle"
                ? "Ready"
                : status === "loading"
                  ? "Fetching from Jira..."
                  : status === "done"
                    ? "Report ready"
                    : "Error"}
            </span>
          </div>
        </div>

        {/* Terminal log */}
        {log.length > 0 && (
          <div
            style={{
              background: "#1a0d24",
              borderRadius: 10,
              padding: "16px 20px",
              marginBottom: 24,
              fontFamily: "ui-monospace, monospace",
              fontSize: 12,
              lineHeight: 1.8,
            }}
          >
            {log.map((line, i) => (
              <div
                key={i}
                style={{
                  color: line.startsWith("Error") ? "#f87171" : "#d8b4fe",
                  display: "flex",
                  gap: 10,
                }}
              >
                <span style={{ color: "#6b3a8a", flexShrink: 0 }}>$</span>
                {line}
              </div>
            ))}
            {status === "loading" && (
              <div style={{ color: purpleMid, display: "flex", gap: 10 }}>
                <span style={{ color: "#6b3a8a" }}>$</span>
                <span style={{ animation: "blink 1s step-end infinite" }}>
                  ▌
                </span>
              </div>
            )}
          </div>
        )}

        {/* Error */}
        {status === "error" && (
          <div
            style={{
              background: "#fef2f2",
              border: "1px solid #fca5a5",
              borderRadius: 8,
              padding: "14px 18px",
              marginBottom: 24,
            }}
          >
            <p
              style={{
                fontSize: 13,
                color: "#b91c1c",
                fontWeight: 600,
                marginBottom: 4,
              }}
            >
              Failed to generate report
            </p>
            <p style={{ fontSize: 13, color: "#ef4444" }}>
              {result?.error ?? "Unknown error"}
            </p>
          </div>
        )}

        {/* Summary cards */}
        {status === "done" && result?.rows && (
          <>
            <div
              style={{
                display: "grid",
                gridTemplateColumns: "repeat(4, minmax(0,1fr))",
                gap: 12,
                marginBottom: 24,
              }}
            >
              {[
                {
                  label: "Stories",
                  value: String(result.storyCount ?? 0),
                  color: "#1c1917",
                  accent: false,
                },
                {
                  label: "Estimate",
                  value: `${totalEst.toFixed(1)}h`,
                  color: "#1c1917",
                  accent: true,
                },
                {
                  label: "Logged",
                  value: `${totalLogged.toFixed(1)}h`,
                  color: "#15803d",
                  accent: false,
                },
                {
                  label: "Remaining",
                  value: `${totalLeft.toFixed(1)}h`,
                  color: "#92400e",
                  accent: false,
                },
              ].map((s) => (
                <div
                  key={s.label}
                  style={{
                    background: "#fff",
                    border: `1px solid ${s.accent ? "#c084d4" : "#e9d5f7"}`,
                    borderTop: s.accent
                      ? `3px solid ${purpleMid}`
                      : `3px solid #e9d5f7`,
                    borderRadius: 10,
                    padding: "14px 16px",
                  }}
                >
                  <div
                    style={{
                      fontSize: 11,
                      textTransform: "uppercase",
                      letterSpacing: "0.1em",
                      color: "#a8a29e",
                      marginBottom: 6,
                    }}
                  >
                    {s.label}
                  </div>
                  <div
                    style={{ fontSize: 22, fontWeight: 700, color: s.color }}
                  >
                    {s.value}
                  </div>
                </div>
              ))}
            </div>

            {/* Preview table */}
            <div
              style={{
                background: "#fff",
                border: "1px solid #e9d5f7",
                borderRadius: 10,
                overflow: "hidden",
                marginBottom: 16,
              }}
            >
              <div
                style={{
                  padding: "13px 18px",
                  borderBottom: "1px solid #f3e8ff",
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  background: purpleLight,
                }}
              >
                <span
                  style={{
                    fontSize: 12,
                    fontWeight: 600,
                    textTransform: "uppercase",
                    letterSpacing: "0.1em",
                    color: purple,
                  }}
                >
                  Preview — first 20 rows
                </span>
                <span style={{ fontSize: 12, color: "#9333ea" }}>
                  {result.rows.length} total rows in XLSX
                </span>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table
                  style={{
                    width: "100%",
                    borderCollapse: "collapse",
                    fontSize: 12,
                  }}
                >
                  <thead>
                    <tr style={{ background: "#faf5ff" }}>
                      {[
                        "Epic Name",
                        "Story Name",
                        "Project",
                        "Est (h)",
                        "Logged (h)",
                        "Left (h)",
                      ].map((h) => (
                        <th
                          key={h}
                          style={{
                            padding: "10px 14px",
                            textAlign: "left",
                            fontSize: 11,
                            textTransform: "uppercase",
                            letterSpacing: "0.08em",
                            color: "#9333ea",
                            fontWeight: 600,
                            borderBottom: "1px solid #f3e8ff",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {result.rows.slice(0, 20).map((row, i) => (
                      <tr key={i} style={{ borderBottom: "1px solid #faf5ff" }}>
                        <td
                          style={{
                            padding: "9px 14px",
                            color: "#78716c",
                            maxWidth: 140,
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {row.epic_name}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            color: "#1c1917",
                            maxWidth: 200,
                            overflow: "hidden",
                            textOverflow: "ellipsis",
                            whiteSpace: "nowrap",
                          }}
                        >
                          {row.story_name}
                        </td>
                        <td style={{ padding: "9px 14px" }}>
                          <span
                            style={{
                              background: purpleLight,
                              color: purple,
                              padding: "2px 8px",
                              borderRadius: 4,
                              fontSize: 11,
                              fontWeight: 600,
                            }}
                          >
                            {row.project_key}
                          </span>
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            color: "#1c1917",
                            textAlign: "right",
                          }}
                        >
                          {row.original_estimate}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            color: "#15803d",
                            textAlign: "right",
                            fontWeight: 600,
                          }}
                        >
                          {row.total_hours_logged}
                        </td>
                        <td
                          style={{
                            padding: "9px 14px",
                            color: row.hrs_left === 0 ? "#a8a29e" : "#92400e",
                            textAlign: "right",
                            fontWeight: 600,
                          }}
                        >
                          {row.hrs_left}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            <p style={{ fontSize: 12, color: "#a8a29e", textAlign: "center" }}>
              XLSX includes all {result.rows.length} rows ·
              Active_CDEF_Monthly_Report_{month}_{year}.xlsx
            </p>
          </>
        )}

        {/* ── User Time Report section ─────────────────────────────────── */}
        <div
          style={{ height: 1, background: "#ede9fe", margin: "40px 0 28px" }}
        />

        <div style={{ marginBottom: 20 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              background: purpleLight,
              color: purple,
              fontSize: 11,
              fontWeight: 500,
              letterSpacing: "0.1em",
              textTransform: "uppercase",
              padding: "4px 12px",
              borderRadius: 99,
              marginBottom: 12,
            }}
          >
            <div
              style={{
                width: 6,
                height: 6,
                borderRadius: "50%",
                background: purpleMid,
              }}
            />
            poshplug.atlassian.net
          </div>
          <h2
            style={{
              fontSize: 22,
              fontWeight: 700,
              color: "#1c1917",
              marginBottom: 4,
              letterSpacing: "-0.01em",
            }}
          >
            User Time Tracking Report(Last 30 days)
          </h2>
          <p style={{ fontSize: 13, color: "#78716c" }}>{userPeriod}</p>
        </div>

        {/* Action row */}
        <div
          style={{
            display: "flex",
            gap: 12,
            marginBottom: 24,
            alignItems: "center",
          }}
        >
          <button
            onClick={handleGenerateUser}
            disabled={userStatus === "loading"}
            style={{
              background: userStatus === "loading" ? "#c4b5d4" : purple,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: userStatus === "loading" ? "not-allowed" : "pointer",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "background 0.15s",
            }}
          >
            {userStatus === "loading" && (
              <span
                style={{
                  width: 12,
                  height: 12,
                  borderRadius: "50%",
                  border: "2px solid rgba(255,255,255,0.35)",
                  borderTopColor: "#fff",
                  display: "inline-block",
                  animation: "spin 0.7s linear infinite",
                }}
              />
            )}
            {userStatus === "loading"
              ? "Generating..."
              : "Generate User Report"}
          </button>

          <button
            onClick={handleDownloadUser}
            disabled={userStatus !== "done"}
            style={{
              background: userStatus === "done" ? purpleLight : "#f3f4f6",
              color: userStatus === "done" ? purple : "#9ca3af",
              border:
                userStatus === "done"
                  ? "1.5px solid #c084d4"
                  : "1.5px solid #e5e7eb",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: userStatus === "done" ? "pointer" : "not-allowed",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "all 0.15s",
            }}
          >
            <svg
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
              />
            </svg>
            Download XLSX
          </button>

          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div
              style={{
                width: 7,
                height: 7,
                borderRadius: "50%",
                background:
                  userStatus === "done"
                    ? "#639922"
                    : userStatus === "error"
                      ? "#E24B4A"
                      : userStatus === "loading"
                        ? purpleMid
                        : "#d1d5db",
                animation:
                  userStatus === "loading" ? "pulse 1s infinite" : "none",
              }}
            />
            <span style={{ fontSize: 12, color: "#78716c" }}>
              {userStatus === "idle"
                ? "Ready"
                : userStatus === "loading"
                  ? "Fetching from Jira..."
                  : userStatus === "done"
                    ? "Report ready"
                    : "Error"}
            </span>
          </div>
        </div>

        {/* Terminal log */}
        {userLog.length > 0 && (
          <div
            style={{
              background: "#1a0d24",
              borderRadius: 10,
              padding: "16px 20px",
              marginBottom: 24,
              fontFamily: "ui-monospace, monospace",
              fontSize: 12,
              lineHeight: 1.8,
            }}
          >
            {userLog.map((line, i) => (
              <div
                key={i}
                style={{
                  color: line.startsWith("Error") ? "#f87171" : "#d8b4fe",
                  display: "flex",
                  gap: 10,
                }}
              >
                <span style={{ color: "#6b3a8a", flexShrink: 0 }}>$</span>
                {line}
              </div>
            ))}
            {userStatus === "loading" && (
              <div style={{ color: purpleMid, display: "flex", gap: 10 }}>
                <span style={{ color: "#6b3a8a" }}>$</span>
                <span style={{ animation: "blink 1s step-end infinite" }}>
                  ▌
                </span>
              </div>
            )}
          </div>
        )}

        {/* Error */}
        {userStatus === "error" && (
          <div
            style={{
              background: "#fef2f2",
              border: "1px solid #fca5a5",
              borderRadius: 8,
              padding: "14px 18px",
              marginBottom: 24,
            }}
          >
            <p
              style={{
                fontSize: 13,
                color: "#b91c1c",
                fontWeight: 600,
                marginBottom: 4,
              }}
            >
              Failed to generate report
            </p>
            <p style={{ fontSize: 13, color: "#ef4444" }}>
              {userResult?.error ?? "Unknown error"}
            </p>
          </div>
        )}

        {/* Preview table */}
        {userStatus === "done" &&
          userResult?.rows &&
          (() => {
            const byUser: Record<string, UserReportRow[]> = {};
            for (const r of userResult.rows) {
              if (!byUser[r.user]) byUser[r.user] = [];
              byUser[r.user].push(r);
            }
            const totalHours = userResult.rows.reduce((s, r) => s + r.hours, 0);
            return (
              <>
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(3, minmax(0,1fr))",
                    gap: 12,
                    marginBottom: 24,
                  }}
                >
                  {[
                    {
                      label: "Users",
                      value: String(Object.keys(byUser).length),
                    },
                    { label: "Tasks", value: String(userResult.rows.length) },
                    {
                      label: "Total Hours",
                      value: `${totalHours.toFixed(1)}h`,
                    },
                  ].map((s, i) => (
                    <div
                      key={s.label}
                      style={{
                        background: "#fff",
                        border: `1px solid ${i === 2 ? "#c084d4" : "#e9d5f7"}`,
                        borderTop:
                          i === 2
                            ? `3px solid ${purpleMid}`
                            : "3px solid #e9d5f7",
                        borderRadius: 10,
                        padding: "14px 16px",
                      }}
                    >
                      <div
                        style={{
                          fontSize: 11,
                          textTransform: "uppercase",
                          letterSpacing: "0.1em",
                          color: "#a8a29e",
                          marginBottom: 6,
                        }}
                      >
                        {s.label}
                      </div>
                      <div
                        style={{
                          fontSize: 22,
                          fontWeight: 700,
                          color: "#1c1917",
                        }}
                      >
                        {s.value}
                      </div>
                    </div>
                  ))}
                </div>

                <div
                  style={{
                    background: "#fff",
                    border: "1px solid #e9d5f7",
                    borderRadius: 10,
                    overflow: "hidden",
                    marginBottom: 16,
                  }}
                >
                  <div
                    style={{
                      padding: "13px 18px",
                      borderBottom: "1px solid #f3e8ff",
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      background: purpleLight,
                    }}
                  >
                    <span
                      style={{
                        fontSize: 12,
                        fontWeight: 600,
                        textTransform: "uppercase",
                        letterSpacing: "0.1em",
                        color: purple,
                      }}
                    >
                      Preview — by user
                    </span>
                    <span style={{ fontSize: 12, color: "#9333ea" }}>
                      {userResult.rows.length} total rows in XLSX
                    </span>
                  </div>
                  <div style={{ overflowX: "auto" }}>
                    <table
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        fontSize: 12,
                      }}
                    >
                      <thead>
                        <tr style={{ background: "#faf5ff" }}>
                          {[
                            "User",
                            "Epic Name",
                            "Story Name",
                            "Project",
                            "Status",
                            "Hours",
                          ].map((h) => (
                            <th
                              key={h}
                              style={{
                                padding: "10px 14px",
                                textAlign: "left",
                                fontSize: 11,
                                textTransform: "uppercase",
                                letterSpacing: "0.08em",
                                color: "#9333ea",
                                fontWeight: 600,
                                borderBottom: "1px solid #f3e8ff",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {userResult.rows.slice(0, 20).map((row, i) => (
                          <tr
                            key={i}
                            style={{ borderBottom: "1px solid #faf5ff" }}
                          >
                            <td
                              style={{
                                padding: "9px 14px",
                                fontWeight: 600,
                                color: purple,
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.user}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#78716c",
                                maxWidth: 120,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.epic_name}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#1c1917",
                                maxWidth: 160,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.story_name}
                            </td>
                            <td style={{ padding: "9px 14px" }}>
                              <span
                                style={{
                                  background: purpleLight,
                                  color: purple,
                                  padding: "2px 8px",
                                  borderRadius: 4,
                                  fontSize: 11,
                                  fontWeight: 600,
                                }}
                              >
                                {row.project_key}
                              </span>
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#78716c",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.status}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#15803d",
                                textAlign: "right",
                                fontWeight: 600,
                              }}
                            >
                              {row.hours}h
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <p
                  style={{
                    fontSize: 12,
                    color: "#a8a29e",
                    textAlign: "center",
                  }}
                >
                  XLSX includes all {userResult.rows.length} rows · grouped by
                  user
                </p>
              </>
            );
          })()}

        {/* ── Project Time Report section ──────────────────────────────────── */}
        <div
          style={{ height: 1, background: "#ede9fe", margin: "40px 0 28px" }}
        />

        <div style={{ marginBottom: 20 }}>
          <div
            style={{
              display: "inline-flex",
              alignItems: "center",
              gap: 6,
              background: purpleLight,
              color: purple,
              fontSize: 11,
              fontWeight: 500,
              letterSpacing: "0.1em",
              textTransform: "uppercase",
              padding: "4px 12px",
              borderRadius: 99,
              marginBottom: 12,
            }}
          >
            <div
              style={{
                width: 6,
                height: 6,
                borderRadius: "50%",
                background: purpleMid,
              }}
            />
            Last 30 Days
          </div>
          <h2
            style={{
              fontSize: 22,
              fontWeight: 700,
              color: "#1c1917",
              marginBottom: 4,
              letterSpacing: "-0.01em",
            }}
          >
            Project Time Tracking Report
          </h2>
          <p style={{ fontSize: 13, color: "#78716c" }}>{userPeriod}</p>
        </div>

        {/* Action row */}
        <div
          style={{
            display: "flex",
            gap: 12,
            marginBottom: 24,
            alignItems: "center",
          }}
        >
          <button
            onClick={handleGenerateProject}
            disabled={projStatus === "loading"}
            style={{
              background: projStatus === "loading" ? "#c4b5d4" : purple,
              color: "#fff",
              border: "none",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: projStatus === "loading" ? "not-allowed" : "pointer",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "background 0.15s",
            }}
          >
            {projStatus === "loading" && (
              <span
                style={{
                  width: 12,
                  height: 12,
                  borderRadius: "50%",
                  border: "2px solid rgba(255,255,255,0.35)",
                  borderTopColor: "#fff",
                  display: "inline-block",
                  animation: "spin 0.7s linear infinite",
                }}
              />
            )}
            {projStatus === "loading"
              ? "Generating..."
              : "Generate Project Report"}
          </button>

          <button
            onClick={handleDownloadProject}
            disabled={projStatus !== "done"}
            style={{
              background: projStatus === "done" ? purpleLight : "#f3f4f6",
              color: projStatus === "done" ? purple : "#9ca3af",
              border:
                projStatus === "done"
                  ? "1.5px solid #c084d4"
                  : "1.5px solid #e5e7eb",
              borderRadius: 8,
              padding: "10px 22px",
              fontSize: 13,
              fontWeight: 600,
              cursor: projStatus === "done" ? "pointer" : "not-allowed",
              display: "flex",
              alignItems: "center",
              gap: 8,
              transition: "all 0.15s",
            }}
          >
            <svg
              width="14"
              height="14"
              fill="none"
              stroke="currentColor"
              strokeWidth="2.5"
              viewBox="0 0 24 24"
            >
              <path
                strokeLinecap="round"
                strokeLinejoin="round"
                d="M4 16v1a3 3 0 003 3h10a3 3 0 003-3v-1m-4-4l-4 4m0 0l-4-4m4 4V4"
              />
            </svg>
            Download XLSX
          </button>

          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <div
              style={{
                width: 7,
                height: 7,
                borderRadius: "50%",
                background:
                  projStatus === "done"
                    ? "#639922"
                    : projStatus === "error"
                      ? "#E24B4A"
                      : projStatus === "loading"
                        ? purpleMid
                        : "#d1d5db",
                animation:
                  projStatus === "loading" ? "pulse 1s infinite" : "none",
              }}
            />
            <span style={{ fontSize: 12, color: "#78716c" }}>
              {projStatus === "idle"
                ? "Ready"
                : projStatus === "loading"
                  ? "Fetching from Jira..."
                  : projStatus === "done"
                    ? "Report ready"
                    : "Error"}
            </span>
          </div>
        </div>

        {/* Terminal log */}
        {projLog.length > 0 && (
          <div
            style={{
              background: "#1a0d24",
              borderRadius: 10,
              padding: "16px 20px",
              marginBottom: 24,
              fontFamily: "ui-monospace, monospace",
              fontSize: 12,
              lineHeight: 1.8,
            }}
          >
            {projLog.map((line, i) => (
              <div
                key={i}
                style={{
                  color: line.startsWith("Error") ? "#f87171" : "#d8b4fe",
                  display: "flex",
                  gap: 10,
                }}
              >
                <span style={{ color: "#6b3a8a", flexShrink: 0 }}>$</span>
                {line}
              </div>
            ))}
            {projStatus === "loading" && (
              <div style={{ color: purpleMid, display: "flex", gap: 10 }}>
                <span style={{ color: "#6b3a8a" }}>$</span>
                <span style={{ animation: "blink 1s step-end infinite" }}>
                  ▌
                </span>
              </div>
            )}
          </div>
        )}

        {/* Error */}
        {projStatus === "error" && (
          <div
            style={{
              background: "#fef2f2",
              border: "1px solid #fca5a5",
              borderRadius: 8,
              padding: "14px 18px",
              marginBottom: 24,
            }}
          >
            <p
              style={{
                fontSize: 13,
                color: "#b91c1c",
                fontWeight: 600,
                marginBottom: 4,
              }}
            >
              Failed to generate report
            </p>
            <p style={{ fontSize: 13, color: "#ef4444" }}>
              {projResult?.error ?? "Unknown error"}
            </p>
          </div>
        )}

        {/* Preview table */}
        {projStatus === "done" &&
          projResult?.rows &&
          (() => {
            const byProject: Record<string, ProjectReportRow[]> = {};
            for (const r of projResult.rows) {
              if (!byProject[r.project_name]) byProject[r.project_name] = [];
              byProject[r.project_name].push(r);
            }
            const totalLogged = projResult.rows.reduce(
              (s, r) => s + r.hours_logged,
              0,
            );
            const totalEst = projResult.rows.reduce(
              (s, r) => s + r.original_estimate,
              0,
            );
            return (
              <>
                <div
                  style={{
                    display: "grid",
                    gridTemplateColumns: "repeat(4, minmax(0,1fr))",
                    gap: 12,
                    marginBottom: 24,
                  }}
                >
                  {[
                    {
                      label: "Projects",
                      value: String(Object.keys(byProject).length),
                    },
                    { label: "Tasks", value: String(projResult.rows.length) },
                    { label: "Estimate", value: `${totalEst.toFixed(1)}h` },
                    { label: "Logged", value: `${totalLogged.toFixed(1)}h` },
                  ].map((s, i) => (
                    <div
                      key={s.label}
                      style={{
                        background: "#fff",
                        border: `1px solid ${i === 3 ? "#c084d4" : "#e9d5f7"}`,
                        borderTop:
                          i === 3
                            ? `3px solid ${purpleMid}`
                            : "3px solid #e9d5f7",
                        borderRadius: 10,
                        padding: "14px 16px",
                      }}
                    >
                      <div
                        style={{
                          fontSize: 11,
                          textTransform: "uppercase",
                          letterSpacing: "0.1em",
                          color: "#a8a29e",
                          marginBottom: 6,
                        }}
                      >
                        {s.label}
                      </div>
                      <div
                        style={{
                          fontSize: 22,
                          fontWeight: 700,
                          color: "#1c1917",
                        }}
                      >
                        {s.value}
                      </div>
                    </div>
                  ))}
                </div>

                <div
                  style={{
                    background: "#fff",
                    border: "1px solid #e9d5f7",
                    borderRadius: 10,
                    overflow: "hidden",
                    marginBottom: 16,
                  }}
                >
                  <div
                    style={{
                      padding: "13px 18px",
                      borderBottom: "1px solid #f3e8ff",
                      display: "flex",
                      justifyContent: "space-between",
                      alignItems: "center",
                      background: purpleLight,
                    }}
                  >
                    <span
                      style={{
                        fontSize: 12,
                        fontWeight: 600,
                        textTransform: "uppercase",
                        letterSpacing: "0.1em",
                        color: purple,
                      }}
                    >
                      Preview — by project
                    </span>
                    <span style={{ fontSize: 12, color: "#9333ea" }}>
                      {projResult.rows.length} total rows in XLSX
                    </span>
                  </div>
                  <div style={{ overflowX: "auto" }}>
                    <table
                      style={{
                        width: "100%",
                        borderCollapse: "collapse",
                        fontSize: 12,
                      }}
                    >
                      <thead>
                        <tr style={{ background: "#faf5ff" }}>
                          {[
                            "Project",
                            "Epic Name",
                            "Story Name",
                            "Status",
                            "Est (h)",
                            "Logged (h)",
                            "Left (h)",
                          ].map((h) => (
                            <th
                              key={h}
                              style={{
                                padding: "10px 14px",
                                textAlign: "left",
                                fontSize: 11,
                                textTransform: "uppercase",
                                letterSpacing: "0.08em",
                                color: "#9333ea",
                                fontWeight: 600,
                                borderBottom: "1px solid #f3e8ff",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {h}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {projResult.rows.slice(0, 20).map((row, i) => (
                          <tr
                            key={i}
                            style={{ borderBottom: "1px solid #faf5ff" }}
                          >
                            <td style={{ padding: "9px 14px" }}>
                              <span
                                style={{
                                  background: purpleLight,
                                  color: purple,
                                  padding: "2px 8px",
                                  borderRadius: 4,
                                  fontSize: 11,
                                  fontWeight: 600,
                                }}
                              >
                                {row.project_key}
                              </span>
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#78716c",
                                maxWidth: 120,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.epic_name}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#1c1917",
                                maxWidth: 160,
                                overflow: "hidden",
                                textOverflow: "ellipsis",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.story_name}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#78716c",
                                whiteSpace: "nowrap",
                              }}
                            >
                              {row.status}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#1c1917",
                                textAlign: "right",
                              }}
                            >
                              {row.original_estimate}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color: "#15803d",
                                textAlign: "right",
                                fontWeight: 600,
                              }}
                            >
                              {row.hours_logged}
                            </td>
                            <td
                              style={{
                                padding: "9px 14px",
                                color:
                                  row.hours_left === 0 ? "#a8a29e" : "#92400e",
                                textAlign: "right",
                                fontWeight: 600,
                              }}
                            >
                              {row.hours_left}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>

                <p
                  style={{
                    fontSize: 12,
                    color: "#a8a29e",
                    textAlign: "center",
                  }}
                >
                  XLSX includes all {projResult.rows.length} rows · grouped by
                  project
                </p>
              </>
            );
          })()}
      </div>

      <style>{`
        @keyframes spin { to { transform: rotate(360deg); } }
        @keyframes blink { 0%, 100% { opacity: 1; } 50% { opacity: 0; } }
        @keyframes pulse { 0%, 100% { opacity: 1; } 50% { opacity: 0.4; } }
      `}</style>
    </main>
  );
}
