import { useEffect, useMemo, useRef, useState } from "react";
import { MarketplaceChat, ChatRequestButton, useMarketplaceChat } from "./MarketplaceChat";
import loginBackground from "./assets/background.png";
import myFleetLogo from "./assets/My_Fleet_512x512.png";
import * as pdfjsLib from "pdfjs-dist";
import { createWorker } from "tesseract.js";
import * as XLSX from "xlsx";
import {
  callerOutcomes,
  clients,
  formatMoney,
  getClient,
  getVehicle,
  initialAppData,
  initialNotifications,
  liability,
  applyVerifiedCycleToVehicle,
  buildNextCycleDueTask,
  initialWhatsAppTemplates,
  permissionRows,
  reportGroups,
  users
} from "./kuberFinanceCore";

// ─── Login is verified against the backend database (/api/login) ─────────────
// No hardcoded demo credentials. Admin creates customer accounts in-app.
function getApiBase() {
  const envBase = import.meta.env.VITE_API_BASE_URL;
  if (envBase) return envBase.replace(/\/$/, "");
  return "https://erp.aakashfinance.com";
}

const API_BASE = getApiBase();

const LOGIN_KEY = "kuber-admin-session";
// v2 prevents an old/incompatible browser sync queue from blocking the current database data.
const PENDING_SYNC_KEY = "kuber-finance-pending-sync-v2";
// Keep OCR language data on the same deployment as the app. This avoids
// browser/CDN restrictions that can make PDF imports fail only in production.
const TESSERACT_LANG_PATH = `${import.meta.env.BASE_URL}eng`;

function createFinanceOcrWorker() {
  return createWorker("eng", 1, {
    langPath: TESSERACT_LANG_PATH,
    gzip: false
  });
}

function loadSession() {
  try {
    const raw = localStorage.getItem(LOGIN_KEY) || sessionStorage.getItem(LOGIN_KEY);
    if (raw && !localStorage.getItem(LOGIN_KEY)) {
      localStorage.setItem(LOGIN_KEY, raw);
    }
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function saveSession(user) {
  localStorage.setItem(LOGIN_KEY, JSON.stringify(user));
}

function clearSession() {
  localStorage.removeItem(LOGIN_KEY);
  sessionStorage.removeItem(LOGIN_KEY);
}

function authHeaders(token, headers = {}) {
  return token ? { ...headers, Authorization: `Bearer ${token}` } : headers;
}

// ─── Login Page (role selection: Admin or Customer) ──────────────────────────
function LoginPage({ onLogin }) {
  const [role, setRole] = useState("admin");
  const [username, setUsername] = useState("");
  const [password, setPassword] = useState("");
  const [error, setError] = useState("");
  const [showPassword, setShowPassword] = useState(false);
  const [loading, setLoading] = useState(false);

  const isAdmin = role === "admin";
  const isCaller = role === "caller";

  const handleSubmit = async (event) => {
    event.preventDefault();
    setError("");
    setLoading(true);
    try {
      const response = await fetch(`${API_BASE}/api/login`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ email: username.trim(), password })
      });
      const data = await response.json();
      if (!response.ok) {
        throw new Error(data.error || "Invalid email or password.");
      }
      const accountRole = String(data.role || "Customer");
      const accountKey = accountRole.toLowerCase() === "admin"
        ? "admin"
        : accountRole.toLowerCase() === "caller"
          ? "caller"
          : "customer";
      // The authenticated backend role is authoritative. Keep the selector in
      // sync so a customer/caller account cannot remain highlighted as Admin.
      if (accountKey !== role) setRole(accountKey);
      const kind = accountKey === "admin" ? "admin" : accountKey === "caller" ? "caller" : "customer";
      onLogin({ ...data, kind });
    } catch (err) {
      setError(err.message || "Login failed. Please try again.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="login-shell">
      <aside
        className="login-hero"
        style={{ backgroundImage: `url(${loginBackground})` }}
      >
        <div className="login-hero-glow" aria-hidden="true" />
        <div className="login-hero-brand">
          <img className="login-logo" src={myFleetLogo} alt="My Fleet" />
          <div>
            <strong>My Fleet</strong>
            <small>Fleet Finance CRM</small>
          </div>
        </div>

        <div className="login-hero-body">
          <div className="login-hero-kicker">
            <span>Finance</span><i /> <span>Manage</span><i /> <span>Grow</span>
          </div>
          <h1 className="login-hero-title">
            Drive your fleet
            <br />
            finance <em>forward.</em>
          </h1>
          <p className="login-hero-sub">
            Track renewals, follow up on dues and resell vehicles — all from one
            clean dashboard.
          </p>

          <div className="login-hero-visual" aria-hidden="true">
            <div className="login-visual-card">
              <div className="login-visual-head">
                <span className="login-visual-dots"><i /><i /><i /></span>
                <span>Fleet Overview</span>
                <span className="login-visual-live"><i /> Live</span>
              </div>
              <div className="login-visual-stats">
                <div className="login-visual-stat">
                  <small>Active vehicles</small>
                  <strong>128</strong>
                  <span className="login-visual-bar" style={{ "--w": "82%" }} />
                </div>
                <div className="login-visual-stat">
                  <small>Renewals due</small>
                  <strong>14</strong>
                  <span className="login-visual-bar" style={{ "--w": "38%" }} />
                </div>
                <div className="login-visual-stat">
                  <small>Collected this month</small>
                  <strong>₹4.2L</strong>
                  <span className="login-visual-bar" style={{ "--w": "64%" }} />
                </div>
              </div>
            </div>
            <div className="login-visual-chip login-visual-chip-a">✓ Renewal paid</div>
            <div className="login-visual-chip login-visual-chip-b">🚚 3 new listings</div>
          </div>

          <ul className="login-hero-features">
            <li>
              <span className="login-hero-feature-icon" aria-hidden="true">⚡</span>
              <span>
                <strong>Smart renewal alerts</strong>
                <small>Never miss a policy or loan renewal.</small>
              </span>
            </li>
            <li>
              <span className="login-hero-feature-icon" aria-hidden="true">🚚</span>
              <span>
                <strong>Fleet &amp; dues at a glance</strong>
                <small>Every vehicle, every payment, one view.</small>
              </span>
            </li>
            <li>
              <span className="login-hero-feature-icon" aria-hidden="true">🏷️</span>
              <span>
                <strong>Resale marketplace</strong>
                <small>List and discover vehicles instantly.</small>
              </span>
            </li>
          </ul>
        </div>

        <div className="login-hero-footer">
          <span>© 2026 My Fleet · All rights reserved</span>
        </div>
      </aside>

      <main className="login-panel">
        <div className="login-card">
          <div className="login-card-accent" aria-hidden="true" />
          <div className="login-card-head">
            <span className="login-card-eyebrow">Welcome back</span>
            <h1 className="login-title">Sign in to continue</h1>
            <p className="login-subtitle">
              {isAdmin
                ? "Enter your admin credentials to access the console."
                : isCaller
                  ? "Sign in with your caller account to open the assigned queue."
                  : "Sign in with your customer account to view your fleet and dues."}
            </p>
          </div>

          <div className="login-role" role="radiogroup" aria-label="Login role">
            <button
              type="button"
              className={role === "admin" ? "active" : ""}
              onClick={() => { setRole("admin"); setError(""); }}
              aria-pressed={role === "admin"}
            >
              <span className="login-role-check" aria-hidden="true">✓</span>
              <span className="login-role-icon"><Icon name="settings" /></span>
              <span className="login-role-text">
                <strong>Admin</strong>
                <small>Full control console</small>
              </span>
            </button>
            <button
              type="button"
              className={role === "caller" ? "active" : ""}
              onClick={() => { setRole("caller"); setError(""); }}
              aria-pressed={role === "caller"}
            >
              <span className="login-role-check" aria-hidden="true">&#10003;</span>
              <span className="login-role-icon"><Icon name="phone" /></span>
              <span className="login-role-text">
                <strong>Caller</strong>
                <small>Assigned follow-ups</small>
              </span>
            </button>
            <button
              type="button"
              className={role === "customer" ? "active" : ""}
              onClick={() => { setRole("customer"); setError(""); }}
              aria-pressed={role === "customer"}
            >
              <span className="login-role-check" aria-hidden="true">✓</span>
              <span className="login-role-icon"><Icon name="people" /></span>
              <span className="login-role-text">
                <strong>Customer</strong>
                <small>Fleet, dues & marketplace</small>
              </span>
            </button>
          </div>

          <form className="login-form" onSubmit={handleSubmit} noValidate>
            <label>
              Email
              <input
                type="email"
                autoComplete="username"
                value={username}
                onChange={(e) => setUsername(e.target.value)}
                placeholder={isAdmin ? "admin@kuber.local" : "owner@company.com"}
                required
                aria-label="Email"
              />
            </label>
            <label>
              Password
              <div className="login-password-wrap">
                <input
                  type={showPassword ? "text" : "password"}
                  autoComplete="current-password"
                  value={password}
                  onChange={(e) => setPassword(e.target.value)}
                  placeholder="••••••••"
                  required
                  aria-label="Password"
                />
                <button
                  type="button"
                  className="login-toggle-pw"
                  onClick={() => setShowPassword((v) => !v)}
                  aria-label={showPassword ? "Hide password" : "Show password"}
                >
                  {showPassword ? "Hide" : "Show"}
                </button>
              </div>
            </label>
            {error && <p className="login-error" role="alert">{error}</p>}
            <button className="login-submit" type="submit" disabled={loading}>
              {loading ? "Signing in…" : "Sign in"}
            </button>
          </form>

          <p className="login-hint">
            Accounts are created by the admin. Contact your administrator for credentials.
          </p>
          <p className="login-privacy-link">
            <a href="/privacy-policy" target="_blank" rel="noreferrer">Privacy Policy</a>
          </p>
        </div>
      </main>
    </div>
  );
}

const navItems = [
  ["dashboard", "Dashboard", "home"],
  ["clients", "Clients", "people"],
  ["fleet", "Fleet", "truck"],
  ["dues", "Smart Alert", "calendar"],
  ["verification", "Verification", "check"],
  ["caller", "Caller", "phone"],
  ["marketplace", "Marketplace", "store"],
  ["chat-requests", "Chat requests", "people"],
  ["reports", "Reports", "chart"],
  ["settings", "Settings", "settings"]
];

const rolePermissionIndex = { Admin: 1, Owner: 2, Caller: 3, Buyer: 4, Customer: 4 };
const sectionPermissionMap = {
  clients: "View all clients",
  "client-profile": "View all clients",
  fleet: "View own fleet",
  dues: "View own fleet",
  documents: "View own fleet",
  verification: "Verify payment",
  caller: "Call / WhatsApp",
  marketplace: "Create sale listing",
  reports: "Owner chat",
  chats: "Owner chat",
  import: "Import Excel"
};

const adminUser = users.find((user) => user.role === "Admin") ?? users[0];
const BANK_RELEASE_RATE_PERCENT = 9.41003284260348;
pdfjsLib.GlobalWorkerOptions.workerSrc = new URL("pdfjs-dist/build/pdf.worker.mjs", import.meta.url).toString();

function emptyAppData() {
  // Admin-created records only — no demo/sample data
  return {
    ...initialAppData,
    clients: [],
    vehicles: [],
    dueTasks: [],
    listings: [],
    callerActivities: [],
    auditLogs: [],
    notifications: [],
    importRows: [],
    clientImports: [],
    verificationItems: [],
    documents: [],
    marketplaceThreads: [],
    whatsappTemplates: initialWhatsAppTemplates,
    whatsappLogs: [],
    users: [],
    rolePermissions: permissionRows,
    reminderSettings: { enabled: true, intervalHours: 24, windowDays: 30 }
  };
}

function savePendingSync(data) {
  try {
    localStorage.setItem(PENDING_SYNC_KEY, JSON.stringify(data));
  } catch (error) {
    console.warn("Unable to cache pending database sync", error);
  }
}

function readPendingSync() {
  try {
    const raw = localStorage.getItem(PENDING_SYNC_KEY);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    return parsed && typeof parsed === "object" ? { ...emptyAppData(), ...parsed } : null;
  } catch (error) {
    return null;
  }
}

function clearPendingSync() {
  try {
    localStorage.removeItem(PENDING_SYNC_KEY);
  } catch (error) {
    // Storage may be unavailable in private browsing; the database remains the source of truth.
  }
}

function parseJsonArray(value) {
  if (Array.isArray(value)) return value;
  if (typeof value !== "string" || !value.trim()) return [];
  try {
    const parsed = JSON.parse(value);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    return [];
  }
}

function loadData() {
  return emptyAppData();
}

async function fetchBackendData(token) {
  const fetchJson = async (path, required = true) => {
    try {
      const response = await fetch(`${API_BASE}/api${path}`, { headers: authHeaders(token) });
      if (!response.ok) {
        const result = await response.json().catch(() => ({}));
        if (required) {
          const backendError = new Error(formatBackendError(result, response.status));
          backendError.status = response.status;
          throw backendError;
        }
        return [];
      }
      return response.json();
    } catch (error) {
      if (required) throw error;
      return [];
    }
  };

  const [rawClients, rawVehicles, rawDues, rawListings, rawCaller, rawAudit, rawImports, rawClientImports, rawDocuments, rawVerificationItems, rawSaleClosings, rawMarketplaceThreads, rawUsers, rawSettings, rawWhatsAppTemplates, rawWhatsAppLogs] = await Promise.all([
    fetchJson("/clients"),
    fetchJson("/vehicles"),
    fetchJson("/dues"),
    fetchJson("/listings"),
    fetchJson("/caller-activities", false),
    fetchJson("/audit-logs", false),
    fetchJson("/imports", false),
    fetchJson("/client-imports", false),
    fetchJson("/documents", false),
    fetchJson("/verification-items", false),
    fetchJson("/sale-closings", false),
    fetchJson("/marketplace-threads", false),
    fetchJson("/users", false),
    fetchJson("/settings", false),
    fetchJson("/whatsapp-templates", false),
    fetchJson("/whatsapp-logs", false)
  ]);

  return deduplicateFinanceRecords({
    ...emptyAppData(),
    clients: Array.isArray(rawClients) ? rawClients.map((row) => ({
      id: row.id,
      name: row.name,
      email: row.email ?? "",
      city: row.city ?? "",
      phone: row.phone ?? "",
      callerId: row.caller_id ?? row.callerId ?? ""
    })) : [],
    vehicles: Array.isArray(rawVehicles) ? rawVehicles.map((row) => ({
      id: row.id,
      clientId: row.client_id ?? row.clientId ?? "",
      type: row.type ?? "Truck",
      regNo: row.reg_no ?? row.regNo ?? "",
      make: row.make ?? "",
      model: row.model ?? "",
      year: Number(row.year ?? 0),
      km: Number(row.km ?? 0),
      principal: Number(row.closing_principal ?? row.closingPrincipal ?? row.principal ?? 0),
      overdue: Number(row.overdue ?? 0),
      penalty: Number(row.penalty ?? 0),
      foreclosure: Number(row.foreclosure ?? 0),
      loanId: row.loan_id ?? row.loanId ?? "",
      loanAccount: row.loan_account ?? row.loanAccount ?? "",
      financier: row.financier ?? "",
      loanAmount: Number(row.loan_amount ?? row.loanAmount ?? 0),
      emiAmount: Number(row.emi_amount ?? row.emiAmount ?? 0),
      interestRate: Number(row.interest_rate ?? row.interestRate ?? 0),
      tenure: Number(row.tenure ?? 0),
      paidEmi: Number(row.paid_emi ?? row.paidEmi ?? 0),
      emiStart: row.emi_start ?? row.emiStart ?? "",
      emiEnd: row.emi_end ?? row.emiEnd ?? "",
      emiSchedule: parseJsonArray(row.emi_schedule_json ?? row.emiSchedule),
      emiHistory: parseJsonArray(row.emi_history_json ?? row.emiHistory),
      insuranceCompany: row.insurance_company ?? row.insuranceCompany ?? "",
      insurancePolicyNo: row.insurance_policy_no ?? row.insurancePolicyNo ?? "",
      insuranceStart: row.insurance_start ?? row.insuranceStart ?? "",
      insuranceHistory: parseJsonArray(row.insurance_history_json ?? row.insuranceHistory),
      permitNo: row.permit_no ?? row.permitNo ?? "",
      permitIssue: row.permit_issue ?? row.permitIssue ?? "",
      permitType: row.permit_type ?? row.permitType ?? "",
      nationalPermitExpiry: row.national_permit_expiry ?? row.nationalPermitExpiry ?? "",
      pucNo: row.puc_no ?? row.pucNo ?? "",
      pucExpiry: row.puc_expiry ?? row.pucExpiry ?? "",
      fitnessExpiry: row.fitness_expiry ?? row.fitnessExpiry ?? "",
      complianceHistory: parseJsonArray(row.compliance_history_json ?? row.complianceHistory),
      combinationId: row.combination_id ?? row.combinationId ?? "",
      insuranceExpiry: row.insurance_expiry ?? row.insuranceExpiry ?? "",
      permitExpiry: row.permit_expiry ?? row.permitExpiry ?? "",
      status: row.status ?? "Active",
      soldDate: row.sold_date ?? row.soldDate ?? ""
    })) : [],
    dueTasks: Array.isArray(rawDues) ? rawDues.map((row) => ({
      id: row.id,
      clientId: row.client_id ?? row.clientId ?? "",
      vehicleId: row.vehicle_id ?? row.vehicleId ?? "",
      type: row.type ?? "EMI",
      amount: Number(row.amount ?? 0),
      dueDate: row.due_date ?? row.dueDate ?? "",
      status: row.status ?? "Due",
      callerId: row.caller_id ?? row.callerId ?? "",
      priority: row.priority ?? "Medium"
    })) : [],
    listings: Array.isArray(rawListings) ? rawListings.map((row) => ({
      id: row.id,
      vehicleId: row.vehicle_id ?? row.vehicleId ?? "",
      title: row.title ?? "",
      price: Number(row.price ?? 0),
      location: row.location ?? "",
      status: row.status ?? "Active",
      condition: row.condition_note ?? row.condition ?? "Good",
      photos: Array.isArray(row.photos) ? row.photos : [],
      createdAt: row.created_at ?? row.createdAt ?? ""
    })) : [],
    callerActivities: Array.isArray(rawCaller) ? rawCaller.map((row) => ({
      id: row.id,
      taskId: row.task_id ?? row.taskId ?? "",
      callerId: row.caller_id ?? row.callerId ?? "",
      outcome: row.outcome ?? "",
      notes: row.notes ?? "",
      expectedAmount: row.expected_amount ?? row.expectedAmount ?? "",
      nextFollowUp: row.next_follow_up ?? row.nextFollowUp ?? "",
      channel: row.channel ?? "Call",
      at: row.occurred_at ?? row.at ?? ""
    })) : [],
    auditLogs: Array.isArray(rawAudit) ? rawAudit.map((row) => ({
      id: row.id,
      module: row.module ?? "",
      action: row.action ?? "",
      record: row.record ?? "",
      oldValue: row.old_value ?? row.oldValue ?? "",
      newValue: row.new_value ?? row.newValue ?? "",
      remark: row.remark ?? "",
      at: row.event_at ?? row.at ?? ""
    })) : [],
    importRows: Array.isArray(rawImports) ? rawImports.map((row) => ({
      row: row.row_no ?? row.row ?? "",
      regNo: row.reg_no ?? row.regNo ?? "",
      assetType: row.asset_type ?? row.assetType ?? "",
      clientName: row.client_name ?? row.clientName ?? "",
      loanAccount: row.loan_account ?? row.loanAccount ?? "",
      lender: row.lender ?? "",
      status: row.status ?? "",
      issue: row.issue ?? ""
    })) : [],
    clientImports: Array.isArray(rawClientImports) ? rawClientImports.map((row) => ({
      id: row.id,
      clientId: row.clientId ?? row.client_id ?? "",
      fileName: row.fileName ?? row.file_name ?? "Customer import",
      importedAt: row.importedAt ?? row.imported_at ?? "",
      rows: Array.isArray(row.rows) ? row.rows : []
    })) : [],
    documents: Array.isArray(rawDocuments) ? rawDocuments.map((row) => ({
      id: row.id,
      clientId: row.clientId ?? row.client_id ?? "",
      vehicleId: row.vehicleId ?? row.vehicle_id ?? "",
      taskId: row.taskId ?? row.task_id ?? "",
      type: row.type ?? "Other",
      fileName: row.fileName ?? row.file_name ?? "document",
      mimeType: row.mimeType ?? row.mime_type ?? "application/octet-stream",
      size: Number(row.size ?? row.size_bytes ?? 0),
      dataUrl: row.dataUrl ?? row.data_url ?? "",
      uploadedBy: row.uploadedBy ?? row.uploaded_by ?? "",
      uploadedAt: row.uploadedAt ?? row.uploaded_at ?? "",
      note: row.note ?? ""
    })) : [],
    verificationItems: Array.isArray(rawVerificationItems) ? rawVerificationItems.map((row) => ({
      id: row.id,
      taskId: row.taskId ?? row.task_id ?? "",
      submittedBy: row.submittedBy ?? row.submitted_by ?? "",
      submittedAt: row.submittedAt ?? row.submitted_at ?? "",
      proofType: row.proofType ?? row.proof_type ?? "Proof",
      details: Array.isArray(row.details) ? row.details : [],
      audit: Array.isArray(row.audit) ? row.audit : []
    })) : [],
    saleClosings: Array.isArray(rawSaleClosings) ? rawSaleClosings.map((row) => ({
      id: row.id,
      listingId: row.listingId ?? row.listing_id ?? "",
      vehicleId: row.vehicleId ?? row.vehicle_id ?? "",
      clientId: row.clientId ?? row.client_id ?? "",
      estimatedAmount: Number(row.estimatedAmount ?? row.estimated_amount ?? 0),
      bankConfirmedAmount: Number(row.bankConfirmedAmount ?? row.bank_confirmed_amount ?? 0),
      foreclosureStatement: row.foreclosureStatement ?? row.foreclosure_statement ?? null,
      bankNoc: row.bankNoc ?? row.bank_noc ?? null,
      ownershipTransfer: row.ownershipTransfer ?? row.ownership_transfer ?? null,
      soldDate: row.soldDate ?? row.sold_date ?? "",
      status: row.status ?? "Estimated",
      updatedAt: row.updatedAt ?? row.updated_at ?? ""
    })) : [],
    marketplaceThreads: Array.isArray(rawMarketplaceThreads) ? rawMarketplaceThreads.map((row) => ({
      id: row.id,
      listingId: row.listingId ?? row.listing_id ?? "",
      buyerClientId: row.buyerClientId ?? row.buyer_client_id ?? "",
      sellerClientId: row.sellerClientId ?? row.seller_client_id ?? "",
      status: row.status ?? "Interested",
      messages: Array.isArray(row.messages) ? row.messages : [],
      reported: Boolean(row.reported),
      blocked: Boolean(row.blocked),
      updatedAt: row.updatedAt ?? row.updated_at ?? ""
    })) : [],
    whatsappTemplates: Array.isArray(rawWhatsAppTemplates) && rawWhatsAppTemplates.length
      ? rawWhatsAppTemplates.map((row) => ({ id: row.id, name: row.name, body: row.body, active: row.active !== false }))
      : initialWhatsAppTemplates,
    whatsappLogs: Array.isArray(rawWhatsAppLogs) ? rawWhatsAppLogs.map((row) => ({
      id: row.id,
      taskId: row.taskId ?? row.task_id ?? "",
      clientId: row.clientId ?? row.client_id ?? "",
      callerId: row.callerId ?? row.caller_id ?? "",
      templateId: row.templateId ?? row.template_id ?? "",
      phone: row.phone ?? "",
      body: row.body ?? "",
      status: row.status ?? "Prepared",
      createdAt: row.createdAt ?? row.created_at ?? "",
      updatedAt: row.updatedAt ?? row.updated_at ?? ""
    })) : [],
    users: Array.isArray(rawUsers) ? rawUsers.map((row) => ({
      id: row.id,
      name: row.name,
      role: row.role,
      email: row.email ?? ""
    })) : [],
    rolePermissions: Array.isArray(rawSettings?.rolePermissions) ? rawSettings.rolePermissions : permissionRows,
    reminderSettings: {
      enabled: rawSettings?.reminderSettings?.enabled !== false,
      intervalHours: Number(rawSettings?.reminderSettings?.intervalHours ?? 24),
      windowDays: Number(rawSettings?.reminderSettings?.windowDays ?? 30)
    }
  });
}

function hasStoredValue(value) {
  return value !== null && value !== undefined && String(value).trim() !== "" && String(value).trim() !== "-";
}

function mergeDuplicateRecord(primary, secondary) {
  const merged = { ...secondary, ...primary };
  Object.keys(secondary ?? {}).forEach((key) => {
    if (!hasStoredValue(primary?.[key]) && hasStoredValue(secondary?.[key])) merged[key] = secondary[key];
  });
  if (Array.isArray(secondary?.emiSchedule) && secondary.emiSchedule.length > (primary?.emiSchedule?.length ?? 0)) {
    merged.emiSchedule = secondary.emiSchedule;
  }
  if (Array.isArray(secondary?.emiHistory) && secondary.emiHistory.length > (primary?.emiHistory?.length ?? 0)) {
    merged.emiHistory = secondary.emiHistory;
  }
  return merged;
}

function financeRecordKeys(clientId, record) {
  const clientKey = String(clientId ?? record?.clientId ?? record?.client_id ?? "").trim();
  const loanAccount = normalizeAgreement(record?.loanAccount ?? record?.loan_account ?? "");
  const registration = normalizeRegNo(baseRegNo(record?.regNo ?? record?.reg_no ?? ""));
  const bodyKey = isBodyRegNo(record?.regNo ?? record?.reg_no) ? ":body" : "";
  const keys = [];
  if (isValidAgreementValue(loanAccount)) keys.push(`${clientKey}:loan:${loanAccount}${bodyKey}`);
  if (registration) keys.push(`${clientKey}:reg:${registration}${bodyKey}`);
  return keys.length ? keys : [`${clientKey}:id:${record?.id ?? Math.random()}`];
}

function dueTaskDateKey(value) {
  const text = String(value ?? "").trim();
  if (!text) return "";
  const dateOnly = text.match(/^(\d{4}-\d{2}-\d{2})$/);
  if (dateOnly) return dateOnly[1];
  const parsed = new Date(text);
  if (Number.isNaN(parsed.getTime())) return text.slice(0, 10);
  const parts = new Intl.DateTimeFormat("en-CA", {
    timeZone: "Asia/Kolkata",
    year: "numeric",
    month: "2-digit",
    day: "2-digit"
  }).formatToParts(parsed).reduce((result, part) => {
    result[part.type] = part.value;
    return result;
  }, {});
  return `${parts.year}-${parts.month}-${parts.day}`;
}

function latestDueTasks(tasks) {
  const latestByVehicle = new Map();
  (tasks ?? []).filter((task) => task.status !== "Closed").forEach((task) => {
    const key = `${task.vehicleId ?? ""}:${task.type ?? "EMI"}`;
    const current = latestByVehicle.get(key);
    if (!current || dueTaskDateKey(task.dueDate) > dueTaskDateKey(current.dueDate)) {
      latestByVehicle.set(key, task);
    }
  });
  return [...latestByVehicle.values()].sort((first, second) =>
    dueTaskDateKey(second.dueDate).localeCompare(dueTaskDateKey(first.dueDate))
  );
}

function deduplicateFinanceRecords(source) {
  const data = source ?? {};
  const vehicleKeyToId = new Map();
  const vehicleIdMap = new Map();
  const vehicles = [];
  (data.vehicles ?? []).forEach((vehicle) => {
    const keys = financeRecordKeys(vehicle.clientId, vehicle);
    const existingId = keys.map((key) => vehicleKeyToId.get(key)).find(Boolean);
    const existingIndex = existingId ? vehicles.findIndex((item) => item.id === existingId) : -1;
    if (existingIndex < 0) {
      vehicles.push(vehicle);
      keys.forEach((key) => vehicleKeyToId.set(key, vehicle.id));
      vehicleIdMap.set(vehicle.id, vehicle.id);
      return;
    }
    const existing = vehicles[existingIndex];
    vehicleIdMap.set(vehicle.id, existing.id);
    vehicles[existingIndex] = mergeDuplicateRecord(existing, vehicle);
    keys.forEach((key) => vehicleKeyToId.set(key, existing.id));
  });

  const dueKeyToId = new Map();
  const dueIdMap = new Map();
  const dueTasks = [];
  (data.dueTasks ?? []).forEach((task) => {
    const vehicleId = vehicleIdMap.get(task.vehicleId) ?? task.vehicleId;
    const key = `${vehicleId}:${task.type ?? "EMI"}:${dueTaskDateKey(task.dueDate) || task.id}`;
    const existingIndex = dueKeyToId.has(key) ? dueTasks.findIndex((item) => item.id === dueKeyToId.get(key)) : -1;
    if (existingIndex < 0) {
      const normalizedTask = { ...task, vehicleId };
      dueTasks.push(normalizedTask);
      dueKeyToId.set(key, normalizedTask.id);
      dueIdMap.set(task.id, normalizedTask.id);
      return;
    }
    const existing = dueTasks[existingIndex];
    dueIdMap.set(task.id, existing.id);
    dueTasks[existingIndex] = mergeDuplicateRecord(existing, { ...task, vehicleId });
  });

  const importRowOwners = new Map();
  const clientImports = [];
  (data.clientImports ?? []).forEach((item) => {
    const nextItem = { ...item, rows: [] };
    (item.rows ?? []).forEach((row, index) => {
      const keys = financeRecordKeys(item.clientId, row);
      const owner = keys.map((key) => importRowOwners.get(key)).find(Boolean);
      if (!owner) {
        const savedRow = { ...row };
        nextItem.rows.push(savedRow);
        keys.forEach((key) => importRowOwners.set(key, { item: nextItem, index: nextItem.rows.length - 1 }));
        return;
      }
      owner.item.rows[owner.index] = mergeDuplicateRecord(owner.item.rows[owner.index], row);
      financeRecordKeys(item.clientId, owner.item.rows[owner.index]).forEach((key) => importRowOwners.set(key, owner));
    });
    if (nextItem.rows.length > 0) clientImports.push(nextItem);
  });

  return {
    ...data,
    vehicles,
    dueTasks,
    clientImports,
    callerActivities: (data.callerActivities ?? []).map((activity) => ({
      ...activity,
      taskId: dueIdMap.get(activity.taskId) ?? activity.taskId
    })),
    verificationItems: (data.verificationItems ?? []).map((item) => ({
      ...item,
      taskId: dueIdMap.get(item.taskId) ?? item.taskId
    })),
    listings: (data.listings ?? []).map((listing) => ({
      ...listing,
      vehicleId: vehicleIdMap.get(listing.vehicleId) ?? listing.vehicleId
    }))
  };
}

function formatBackendError(result, status) {
  if (result?.error === "Database connection is not available.") {
    const host = result.config?.host || "not set";
    return `Database is not connected. Local backend is using DB_HOST=${host}. Check backend/.env and MySQL network access.`;
  }
  return result?.message || result?.error || `Backend request failed with status ${status}.`;
}

async function syncDataToBackend(data, token) {
  const response = await fetch(`${API_BASE}/api/sync`, {
    method: "POST",
    headers: authHeaders(token, { "Content-Type": "application/json" }),
    body: JSON.stringify({
      clients: data.clients ?? [],
      vehicles: data.vehicles ?? [],
      dueTasks: data.dueTasks ?? [],
      listings: data.listings ?? [],
      callerActivities: data.callerActivities ?? [],
      auditLogs: data.auditLogs ?? [],
      importRows: data.importRows ?? [],
      clientImports: data.clientImports ?? [],
      documents: data.documents ?? [],
      verificationItems: data.verificationItems ?? [],
      saleClosings: data.saleClosings ?? [],
      marketplaceThreads: data.marketplaceThreads ?? []
    })
  });
  const result = await response.json().catch(() => ({}));
  if (!response.ok) {
    const backendError = new Error(formatBackendError(result, response.status) || "Backend sync failed.");
    backendError.status = response.status;
    throw backendError;
  }
  const syncMode = String(result?.mode || "").toLowerCase();
  if (syncMode && !syncMode.startsWith("mysql")) {
    throw new Error("MySQL database is unavailable. Data was not saved permanently; it will be retried when the database reconnects.");
  }
  return result;
}

export default function App() {
  const [session, setSession] = useState(loadSession);
  const publicPath = window.location.pathname.replace(/\/+$/, "") || "/";
  if (publicPath === "/privacy-policy") return <PrivacyPolicyPage />;
  const handleLogin = (nextSession) => {
    saveSession(nextSession);
    setSession(nextSession);
  };

  if (!session) {
    return <LoginPage onLogin={handleLogin} />;
  }

  if (session.kind === "customer") {
    return (
      <CustomerPortal session={session} onLogout={() => { clearSession(); setSession(null); }} />
    );
  }

  return <AdminApp session={session} onLogout={() => { clearSession(); setSession(null); }} />;
}

function PrivacyPolicyPage() {
  return (
    <main className="public-policy-page">
      <article className="public-policy-card">
        <p className="public-policy-kicker">MY FLEET</p>
        <h1>Privacy Policy</h1>
        <p className="public-policy-updated"><strong>Effective date:</strong> September 28, 2026</p>

        <p>
          My Fleet helps authorized finance and fleet teams manage customers, vehicles,
          loan records, renewals, dues, documents and marketplace activity. This Privacy Policy
          explains what information the app handles and how it is used.
        </p>

        <h2>Information we collect</h2>
        <ul>
          <li><strong>Account information:</strong> name, email address, login identifier and authentication information.</li>
          <li><strong>Contact information:</strong> phone number and other contact details entered by an authorized administrator or customer.</li>
          <li><strong>Finance information:</strong> loan account numbers, lender details, loan amounts, EMI details, interest rates, repayment schedules, vehicle finance records and related documents.</li>
          <li><strong>Files and photos:</strong> bank PDFs, identity or ownership documents, vehicle photos and other files uploaded through the app.</li>
          <li><strong>Messages and user content:</strong> marketplace listings, chat messages, notes, attachments and support or follow-up activity.</li>
          <li><strong>App activity:</strong> records needed to operate the dashboard, assign tasks, process listings and maintain audit history.</li>
        </ul>

        <h2>How we use information</h2>
        <ul>
          <li>Authenticate users and provide role-based access.</li>
          <li>Manage fleet, customer, loan, EMI, renewal and collection workflows.</li>
          <li>Read bank PDFs and extract finance fields when a user chooses the PDF-reading feature.</li>
          <li>Provide marketplace listings, chat, notifications, reports and customer support.</li>
          <li>Protect the service, prevent misuse and maintain audit records.</li>
        </ul>

        <h2>Third-party services and sharing</h2>
        <p>
          We do not sell personal information. Information is shared only as needed to operate
          the service with our authorized backend/database hosting providers and service providers.
          Bank PDFs are read locally in the browser using OCR. Users may also choose to open
          WhatsApp or another external service from the app; that service then handles information
          under its own privacy policy.
        </p>

        <h2>Security</h2>
        <p>
          Data is transmitted through HTTPS in the production app. Access is protected with
          authenticated sessions and role-based permissions. No security method can guarantee
          absolute security, so users should keep their credentials confidential.
        </p>

        <h2>Retention and deletion</h2>
        <p>
          We retain information while the account or business relationship is active and for as
          long as needed for legitimate operational, legal or accounting purposes. To request
          deletion or correction of personal information, email the privacy contact below from
          the account email and include the name of the organization and the account identifier.
          We may need to verify the request before taking action.
        </p>

        <h2>Children</h2>
        <p>
          My Fleet is a business finance and fleet-management service and is not directed
          to children under 13.
        </p>

        <h2>Changes to this policy</h2>
        <p>
          We may update this policy when the app, services or legal requirements change. The
          updated version will be published on this page with a new effective date.
        </p>

        <h2>Contact</h2>
        <p>
          Privacy requests: <a href="mailto:privacy@aakashfinance.com">privacy@aakashfinance.com</a>
        </p>
        <p className="public-policy-footer">My Fleet · Aakash Finance</p>
      </article>
    </main>
  );
}

function AdminApp({ session, onLogout }) {
  const marketplaceChat = useMarketplaceChat(API_BASE, session.token);
  const [data, setData] = useState(loadData);
  const [section, setSection] = useState(session.kind === "caller" ? "caller" : "dashboard");
  const [lastSavedAt, setLastSavedAt] = useState(null);
  const [saveStatus, setSaveStatus] = useState("Ready");
  const [toast, setToast] = useState("");
  const [activeReport, setActiveReport] = useState(reportGroups[0][0]);
  const [selectedClientId, setSelectedClientId] = useState(null);
  const mutationVersionRef = useRef(0);
  const pendingMutationCountRef = useRef(0);
  const mutationQueueRef = useRef(Promise.resolve());
  const enqueueMutation = (operation) => {
    const request = mutationQueueRef.current.then(operation);
    mutationQueueRef.current = request.catch(() => {});
    return request;
  };
  const visibleNavItems = useMemo(
    () => session.role === "Caller"
      ? navItems.filter(([key]) => key === "caller")
      : navItems,
    [session.role]
  );

  const totals = useMemo(() => {
    const activeVehicles = data.vehicles.filter((vehicle) => vehicle.status !== "Sold");
    const totalLiability = activeVehicles.reduce((sum, vehicle) => sum + Number(vehicle.principal || 0), 0);
    const openDues = data.dueTasks.filter((task) => task.status !== "Closed");
    const overdue = data.dueTasks.filter((task) => ["Overdue", "Escalated"].includes(task.status));
    const proofPending = data.dueTasks.filter((task) => task.status === "Proof Pending");
    const pendingListings = data.listings.filter((listing) => listing.status === "Submitted");
    const trucks = activeVehicles.filter((vehicle) => vehicle.type === "Truck");
    const trailers = activeVehicles.filter((vehicle) => vehicle.type === "Trailer");
    const financedAssets = activeVehicles.filter((vehicle) => Number(vehicle.principal || 0) > 0);
    const emiDue = openDues.filter((task) => task.type === "EMI");
    const emiOverdue = data.dueTasks.filter((task) => task.type === "EMI" && ["Overdue", "Escalated"].includes(task.status));
    const today = new Date();
    const inThirtyDays = new Date(today);
    inThirtyDays.setDate(today.getDate() + 30);
    const insuranceExpiring = data.vehicles.filter((vehicle) => {
      if (!vehicle.insuranceExpiry) return false;
      const expiry = new Date(vehicle.insuranceExpiry);
      return !Number.isNaN(expiry.getTime()) && expiry >= today && expiry <= inThirtyDays;
    });
    const activeCallers = (data.users ?? []).filter((user) => user.role === "Caller");
    const callsDueToday = openDues.filter((task) => task.callerId && ["Due", "Overdue", "Escalated", "Promise-to-Pay"].includes(task.status));
    const activeListings = data.listings.filter((listing) => ["Submitted", "Active", "Reserved", "Changes Required"].includes(listing.status));
    const soldVehicles = data.vehicles.filter((vehicle) => vehicle.status === "Sold");
    return { totalLiability, openDues, overdue, proofPending, pendingListings, trucks, trailers, financedAssets, emiDue, emiOverdue, insuranceExpiring, activeCallers, callsDueToday, activeListings, soldVehicles };
  }, [data]);

  useEffect(() => {
    setSaveStatus("Loading");
    const pendingData = readPendingSync();
    const load = (async () => {
      if (!pendingData) return { data: await fetchBackendData(session.token), pendingSyncError: null };
      try {
        await syncDataToBackend(pendingData, session.token);
        clearPendingSync();
        return { data: await fetchBackendData(session.token), pendingSyncError: null };
      } catch (pendingSyncError) {
        const message = String(pendingSyncError?.message || pendingSyncError || "").toLowerCase();
        // A failed structural sync must not hide valid database data on reload.
        if (message.includes("foreign key") || message.includes("client_imports")) clearPendingSync();
        return { data: await fetchBackendData(session.token), pendingSyncError };
      }
    })();
    load
      .then(({ data: backendData, pendingSyncError }) => {
        setData(backendData);
        setSaveStatus("Database");
        setToast(pendingSyncError ? "Database data loaded; pending local changes could not be synced." : (pendingData ? "Pending changes synced to database" : "Database data loaded"));
      })
      .catch((err) => {
        setSaveStatus("Error");
        if (err?.status === 401) {
          clearSession();
          onLogout();
          return;
        }
        setToast(err.message || "Database load failed");
      });
  }, []);

  // Keep an already-open Admin console in sync with customer actions (for
  // example, a vehicle being marked Sold from the mobile app).
  useEffect(() => {
    let active = true;
    const refreshFromDatabase = async () => {
      if (document.visibilityState === "hidden") return;
      if (pendingMutationCountRef.current > 0 || readPendingSync()) return;
      const refreshVersion = mutationVersionRef.current;
      try {
        const latest = await fetchBackendData(session.token);
        if (active && refreshVersion === mutationVersionRef.current && pendingMutationCountRef.current === 0 && !readPendingSync()) {
          setData(latest);
          setSaveStatus("Database");
        }
      } catch (error) {
        // Keep the current screen usable when the network is temporarily unavailable.
      }
    };
    const intervalId = window.setInterval(refreshFromDatabase, 10000);
    window.addEventListener("focus", refreshFromDatabase);
    return () => {
      active = false;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", refreshFromDatabase);
    };
  }, [session.token]);

  const persist = (nextData, message = "Saved") => {
    const cleanData = deduplicateFinanceRecords(nextData);
    const mutationVersion = ++mutationVersionRef.current;
    pendingMutationCountRef.current += 1;
    setData(cleanData);
    savePendingSync(cleanData);
    setLastSavedAt(new Date().toLocaleTimeString("en-IN", { hour: "2-digit", minute: "2-digit" }));
    setSaveStatus("Saving");
    setToast(message);
    enqueueMutation(() => syncDataToBackend(cleanData, session.token))
      .then((result) => {
        if (result?.ok && mutationVersion === mutationVersionRef.current) {
          clearPendingSync();
          setSaveStatus("Database");
          setToast("Saved to database");
        }
      })
      .catch((err) => {
        if (mutationVersion === mutationVersionRef.current) setSaveStatus("Error");
        if (mutationVersion === mutationVersionRef.current) {
        setToast(err.message || "Database sync failed");
        }
      })
      .finally(() => {
        pendingMutationCountRef.current = Math.max(0, pendingMutationCountRef.current - 1);
      });
  };

  const withAudit = (current, entry) => ({
    ...current,
    auditLogs: [
      {
        ...entry,
        id: `a-${Date.now()}`,
        at: new Date().toLocaleString("en-IN")
      },
      ...current.auditLogs
    ]
  });

  const notify = (current, title, detail, target) => ({
    ...current,
    notifications: [{ id: `n-${Date.now()}`, title, detail, target, unread: true }, ...current.notifications]
  });

  const addClient = async (event) => {
    event.preventDefault();
    const formElement = event.currentTarget;
    const form = new FormData(formElement);
    const name = form.get("name")?.toString().trim() ?? "";
    const city = form.get("city")?.toString().trim() ?? "";
    const phone = form.get("phone")?.toString().trim() ?? "";
    const email = form.get("email")?.toString().trim() ?? "";

    if (!name || !city || !phone || !email) {
      setToast("Name, city, phone, and email are required");
      return;
    }

    setSaveStatus("Saving");
    setToast("Creating customer account...");
    try {
      const response = await fetch(`${API_BASE}/api/users`, {
        method: "POST",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ name, email })
      });
      const result = await response.json();
      if (!response.ok) {
        throw new Error(result.error || "Failed to create customer account.");
      }

      const nextClient = {
        id: result.clientId,
        name,
        email,
        city,
        phone,
        callerId: ""
      };
      const next = notify({ ...data, clients: [nextClient, ...data.clients] }, "Client added", nextClient.name, "clients");
      persist(withAudit(next, {
        module: "Clients",
        action: "Created",
        record: nextClient.name,
        oldValue: "-",
        newValue: "Active",
        remark: `Customer account created in database (${email})`
      }), `${nextClient.name} added`);
      formElement.reset();
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Failed to create customer account.");
    }
  };

  const updateTaskStatus = (taskId, status, remark = "Admin web action saved") => {
    const task = data.dueTasks.find((item) => item.id === taskId);
    const vehicle = task ? data.vehicles.find((item) => item.id === task.vehicleId) : null;
    const isVerificationApproval = status === "Closed" && ["Proof Pending", "Verification Pending"].includes(task?.status);
    const nextCycleTask = isVerificationApproval && vehicle ? buildNextCycleDueTask(task, vehicle) : null;
    const alreadyScheduled = nextCycleTask && data.dueTasks.some((item) => item.id === nextCycleTask.id || (
      item.vehicleId === nextCycleTask.vehicleId && item.type === nextCycleTask.type && item.dueDate === nextCycleTask.dueDate
    ));
    const verifiedAt = new Date().toISOString().slice(0, 10);
    const updated = {
      ...data,
      dueTasks: [
        ...(nextCycleTask && !alreadyScheduled ? [nextCycleTask] : []),
        ...data.dueTasks.map((item) => (item.id === taskId ? { ...item, status } : item))
      ],
      vehicles: isVerificationApproval && vehicle
        ? data.vehicles.map((item) => item.id === vehicle.id ? applyVerifiedCycleToVehicle(item, task, verifiedAt, nextCycleTask) : item)
        : data.vehicles,
      verificationItems: ["Closed", "Due"].includes(status)
        ? data.verificationItems.filter((item) => item.taskId !== taskId)
        : data.verificationItems
    };
    const next = notify(updated, `Verification ${status}`, nextCycleTask ? `${remark} Next ${task.type} cycle scheduled for ${nextCycleTask.dueDate}.` : remark, "verification");
    persist(withAudit(next, {
      module: "Verification",
      action: status === "Closed" ? "Approved" : status === "Due" ? "Rejected" : "More Info",
      record: task?.id ?? taskId,
      oldValue: task?.status ?? "-",
      newValue: status,
      remark
    }), `Task ${status}`);
  };

  const assignCallerToTask = async (taskId, callerId) => {
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/caller-assignment/tasks/${encodeURIComponent(taskId)}`, {
        method: "POST",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ callerId })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Caller assignment failed");
      const backendData = await fetchBackendData(session.token);
      setData(backendData);
      setSaveStatus("Database");
      setToast(callerId ? `Due assigned to ${backendData.users.find((user) => user.id === callerId)?.name ?? "caller"}` : "Caller assignment removed");
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Caller assignment failed");
    }
  };

  const saveCallerOutcome = (event, task) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const outcome = form.get("outcome")?.toString() || "No Answer / Busy";
    const notes = form.get("notes")?.toString().trim() || "Follow-up recorded.";
    const expectedAmount = form.get("expectedAmount")?.toString().trim() || "";
    const nextFollowUp = form.get("nextFollowUp")?.toString().trim() || "Next scheduled follow-up";
    const channel = form.get("channel")?.toString() || "Call";
    const nextStatus =
      outcome === "Connected - Will Pay" ? "Promise-to-Pay" :
      outcome === "Already Paid" || outcome === "Insurance Interested" ? "Waiting Documents" :
      outcome === "Wrong Number" || outcome === "Dispute" ? "Escalated" :
      task.status === "Proof Pending" ? "Verification Pending" :
      task.status;
    const activity = {
      id: `ca-${Date.now()}`,
      taskId: task.id,
      callerId: task.callerId,
      outcome,
      notes,
      expectedAmount,
      nextFollowUp: nextStatus === "Verification Pending" ? "Paused" : nextFollowUp,
      channel,
      at: new Date().toLocaleString("en-IN")
    };
    const updated = {
      ...data,
      dueTasks: data.dueTasks.map((item) => (item.id === task.id ? { ...item, status: nextStatus } : item)),
      callerActivities: [activity, ...data.callerActivities]
    };
    const next = notify(updated, `Caller outcome: ${outcome}`, notes, "caller");
    persist(withAudit(next, {
      module: "Caller",
      action: outcome,
      record: task.id,
      oldValue: task.status,
      newValue: nextStatus,
      remark: notes
    }), `Caller outcome saved`);
    event.currentTarget.reset();
  };

  const saveWhatsAppTemplates = async (templates) => {
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/whatsapp-templates`, {
        method: "PUT",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ templates })
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || "WhatsApp templates could not be saved.");
      setData((current) => ({ ...current, whatsappTemplates: result.templates }));
      setSaveStatus("Database");
      setToast("WhatsApp templates saved");
    } catch (error) {
      setSaveStatus("Error");
      setToast(error.message || "WhatsApp templates could not be saved.");
    }
  };

  const openWhatsApp = async (task, templateId) => {
    const client = getDataClient(data, task.clientId);
    const vehicle = getDataVehicle(data, task.vehicleId);
    const template = (data.whatsappTemplates ?? initialWhatsAppTemplates).find((item) => item.id === templateId && item.active)
      ?? (data.whatsappTemplates ?? initialWhatsAppTemplates).find((item) => item.active);
    const phone = String(client?.phone || "").replace(/\D/g, "");
    if (!client || !phone) {
      setToast("Customer phone number is missing.");
      return;
    }
    if (!template) {
      setToast("Create an active WhatsApp template first.");
      return;
    }
    const body = template.body
      .replaceAll("{{customer}}", client.name || "Customer")
      .replaceAll("{{vehicle}}", vehicle?.regNo || "vehicle")
      .replaceAll("{{type}}", task.type || "payment")
      .replaceAll("{{dueDate}}", formatDisplayDate(task.dueDate))
      .replaceAll("{{amount}}", formatMoney(task.amount));
    const status = window.open(`https://wa.me/${phone}?text=${encodeURIComponent(body)}`, "_blank", "noopener,noreferrer") ? "Opened" : "Prepared";
    const now = new Date().toISOString();
    const payload = {
      id: `wa-${Date.now()}`,
      taskId: task.id,
      clientId: task.clientId,
      callerId: task.callerId || session.id || "",
      templateId: template.id,
      phone,
      body,
      status,
      createdAt: now,
      updatedAt: now
    };
    try {
      const response = await fetch(`${API_BASE}/api/whatsapp-logs`, {
        method: "POST",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify(payload)
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || "WhatsApp log could not be saved.");
      setData((current) => ({ ...current, whatsappLogs: [result, ...(current.whatsappLogs ?? [])] }));
      setSaveStatus("Database");
      setToast(`WhatsApp ${status.toLowerCase()} and logged`);
    } catch (error) {
      setSaveStatus("Error");
      setToast(error.message || "WhatsApp log could not be saved.");
    }
  };

  const updateWhatsAppStatus = async (id, status) => {
    try {
      const response = await fetch(`${API_BASE}/api/whatsapp-logs/${encodeURIComponent(id)}`, {
        method: "PATCH",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ status })
      });
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || "WhatsApp status could not be updated.");
      setData((current) => ({ ...current, whatsappLogs: (current.whatsappLogs ?? []).map((item) => item.id === id ? result : item) }));
      setToast(`WhatsApp status updated to ${status}`);
    } catch (error) {
      setToast(error.message || "WhatsApp status could not be updated.");
    }
  };

  const updateListingStatus = (listingId, status) => {
    const listing = data.listings.find((item) => item.id === listingId);
    const updated = {
      ...data,
      listings: data.listings.map((item) => (item.id === listingId ? { ...item, status } : item))
    };
    const next = notify(updated, `Listing ${status}`, listing?.title ?? listingId, "marketplace");
    persist(withAudit(next, {
      module: "Marketplace",
      action: "Status Updated",
      record: listing?.title ?? listingId,
      oldValue: listing?.status ?? "-",
      newValue: status,
      remark: "Admin marketplace decision saved"
    }), `Listing ${status}`);
  };

  const saveSaleClosing = async (event, listingId) => {
    event.preventDefault();
    const listing = data.listings.find((item) => item.id === listingId);
    const vehicle = data.vehicles.find((item) => item.id === listing?.vehicleId);
    if (!listing || !vehicle) return;
    const form = event.currentTarget;
    const formData = new FormData(form);
    const bankConfirmedAmount = Number(formData.get("bankConfirmedAmount") || 0);
    const soldDate = String(formData.get("soldDate") || "");
    const existing = (data.saleClosings ?? []).find((item) => item.listingId === listingId);
    const readClosingFile = async (name, previous) => {
      const file = formData.get(name);
      if (!(file instanceof File) || !file.size) return previous ?? null;
      if (file.size > 8 * 1024 * 1024) throw new Error("Each closing document must be smaller than 8 MB.");
      return { fileName: file.name, mimeType: file.type || "application/octet-stream", size: file.size, dataUrl: await readFileAsDataUrl(file), uploadedAt: new Date().toLocaleString("en-IN") };
    };
    try {
      const combinedVehicles = vehicle.combinationId
        ? data.vehicles.filter((item) => item.combinationId === vehicle.combinationId)
        : [vehicle];
      const record = {
        id: existing?.id ?? `sc-${Date.now()}`,
        listingId,
        vehicleId: vehicle.id,
        clientId: vehicle.clientId,
        estimatedAmount: combinedVehicles.reduce((sum, item) => sum + Number(item.principal || 0), 0),
        bankConfirmedAmount,
        foreclosureStatement: await readClosingFile("foreclosureStatement", existing?.foreclosureStatement),
        bankNoc: await readClosingFile("bankNoc", existing?.bankNoc),
        ownershipTransfer: await readClosingFile("ownershipTransfer", existing?.ownershipTransfer),
        soldDate,
        status: soldDate ? "Sold" : bankConfirmedAmount > 0 ? "Bank Confirmed" : "Estimated",
        updatedAt: new Date().toISOString()
      };
      if (record.status === "Sold" && (!record.bankConfirmedAmount || !record.foreclosureStatement || !record.bankNoc || !record.ownershipTransfer)) {
        setToast("Bank amount, foreclosure statement, NOC and ownership transfer proof are required before marking Sold.");
        return;
      }
      const updated = {
        ...data,
        listings: data.listings.map((item) => item.id === listingId ? { ...item, status: record.status === "Sold" ? "Sold" : item.status } : item),
        vehicles: record.status === "Sold"
          ? data.vehicles.map((item) => item.id === vehicle.id ? { ...item, status: "Sold", soldDate } : item)
          : data.vehicles,
        dueTasks: record.status === "Sold"
          ? data.dueTasks.map((task) => task.vehicleId === vehicle.id ? { ...task, status: "Closed" } : task)
          : data.dueTasks,
        saleClosings: [record, ...(data.saleClosings ?? []).filter((item) => item.id !== record.id)]
      };
      persist(withAudit(updated, {
        module: "Sale Closing",
        action: record.status === "Sold" ? "Sale Completed" : "Bank Confirmation Updated",
        record: listing.title,
        oldValue: existing?.status ?? "Estimated",
        newValue: record.status,
        remark: `Estimated ${formatMoney(record.estimatedAmount)} | Bank confirmed ${formatMoney(record.bankConfirmedAmount)}`
      }), record.status === "Sold" ? "Sale completed and records saved" : "Bank closing details saved");
      form.reset();
    } catch (error) {
      setToast(error.message || "Sale closing details could not be saved.");
    }
  };

  const updateVehicleFinance = (event, vehicleId) => {
    event.preventDefault();
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return;
    const form = new FormData(event.currentTarget);
    const numberFrom = (name, fallback) => Number(form.get(name)?.toString().replace(/\D/g, "")) || fallback;
    const decimalFrom = (name, fallback) => {
      const value = Number(form.get(name)?.toString().replace(/[^\d.-]/g, ""));
      return Number.isFinite(value) && value >= 0 ? value : fallback;
    };
    const loanFields = {
      loanId: form.get("loanId")?.toString().trim() || vehicle.loanId || "",
      loanAccount: form.get("loanAccount")?.toString().trim() || vehicle.loanAccount || "",
      financier: form.get("financier")?.toString().trim() || vehicle.financier || "",
      loanAmount: decimalFrom("loanAmount", vehicle.loanAmount || 0),
      emiAmount: decimalFrom("emiAmount", vehicle.emiAmount || 0),
      interestRate: decimalFrom("interestRate", vehicle.interestRate || 0),
      tenure: numberFrom("tenure", vehicle.tenure || 0),
      paidEmi: numberFrom("paidEmi", vehicle.paidEmi || 0),
      emiStart: form.get("emiStart")?.toString().trim() || vehicle.emiStart || "",
      emiEnd: form.get("emiEnd")?.toString().trim() || vehicle.emiEnd || ""
    };
    const loanPlan = buildLoanPlan(loanFields);
    const nextVehicle = {
      ...vehicle,
      principal: numberFrom("principal", vehicle.principal),
      overdue: numberFrom("overdue", vehicle.overdue),
      penalty: numberFrom("penalty", vehicle.penalty),
      foreclosure: numberFrom("foreclosure", vehicle.foreclosure),
      ...loanFields,
      emiSchedule: loanPlan.schedule.length ? loanPlan.schedule : (vehicle.emiSchedule ?? []),
      emiHistory: loanPlan.history.length ? loanPlan.history : (vehicle.emiHistory ?? []),
      status: form.get("status")?.toString() || vehicle.status
    };
    const updated = {
      ...data,
      vehicles: data.vehicles.map((item) => (item.id === vehicleId ? nextVehicle : item)),
      listings: nextVehicle.status === "Sold"
        ? data.listings.map((item) => item.vehicleId === vehicleId ? { ...item, status: "Sold" } : item)
        : data.listings,
      dueTasks: nextVehicle.status === "Sold"
        ? data.dueTasks.map((task) => task.vehicleId === vehicleId ? { ...task, status: "Closed" } : task)
        : data.dueTasks
    };
    const next = notify(updated, "Vehicle finance updated", vehicle.regNo, "fleet");
    persist(withAudit(next, {
      module: "Fleet",
      action: "Finance Updated",
      record: vehicle.regNo,
      oldValue: formatMoney(liability(vehicle)),
      newValue: formatMoney(liability(nextVehicle)),
      remark: "Admin edited closing principal fields from web"
    }), `${vehicle.regNo} updated`);
  };

  const updateVehicleCompliance = (event, vehicleId) => {
    event.preventDefault();
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return;
    const form = new FormData(event.currentTarget);
    const textFrom = (name, fallback = "") => form.get(name)?.toString().trim() || fallback;
    const nextFields = {
      insuranceCompany: textFrom("insuranceCompany", vehicle.insuranceCompany || ""),
      insurancePolicyNo: textFrom("insurancePolicyNo", vehicle.insurancePolicyNo || ""),
      insuranceStart: textFrom("insuranceStart", vehicle.insuranceStart || ""),
      insuranceExpiry: textFrom("insuranceExpiry", vehicle.insuranceExpiry || ""),
      insuranceStatus: textFrom("insuranceStatus", "Active"),
      permitNo: textFrom("permitNo", vehicle.permitNo || ""),
      permitIssue: textFrom("permitIssue", vehicle.permitIssue || ""),
      permitExpiry: textFrom("permitExpiry", vehicle.permitExpiry || ""),
      permitType: textFrom("permitType", vehicle.permitType || ""),
      nationalPermitExpiry: textFrom("nationalPermitExpiry", vehicle.nationalPermitExpiry || ""),
      pucNo: textFrom("pucNo", vehicle.pucNo || ""),
      pucExpiry: textFrom("pucExpiry", vehicle.pucExpiry || ""),
      fitnessExpiry: textFrom("fitnessExpiry", vehicle.fitnessExpiry || ""),
      complianceStatus: textFrom("complianceStatus", "Active")
    };
    const insuranceRecord = buildInsuranceHistoryEntry(nextFields);
    const complianceRecords = buildComplianceHistoryEntries(nextFields);
    const updated = {
      ...data,
      vehicles: data.vehicles.map((item) => item.id === vehicleId ? {
        ...item,
        insuranceCompany: nextFields.insuranceCompany,
        insurancePolicyNo: nextFields.insurancePolicyNo,
        insuranceStart: nextFields.insuranceStart,
        insuranceExpiry: nextFields.insuranceExpiry,
        insuranceHistory: insuranceRecord ? upsertHistory(item.insuranceHistory, insuranceRecord, (entry) => `${entry.policyNo}:${entry.expiryDate}`) : (item.insuranceHistory ?? []),
        permitNo: nextFields.permitNo,
        permitIssue: nextFields.permitIssue,
        permitExpiry: nextFields.permitExpiry,
        permitType: nextFields.permitType,
        nationalPermitExpiry: nextFields.nationalPermitExpiry,
        pucNo: nextFields.pucNo,
        pucExpiry: nextFields.pucExpiry,
        fitnessExpiry: nextFields.fitnessExpiry,
        complianceHistory: mergeComplianceHistory(item.complianceHistory, complianceRecords)
      } : item)
    };
    const next = notify(updated, "Insurance and compliance updated", vehicle.regNo, "fleet");
    persist(withAudit(next, {
      module: "Fleet",
      action: "Insurance and Compliance Updated",
      record: vehicle.regNo,
      oldValue: `${vehicle.insurancePolicyNo || "No policy"} | ${vehicle.permitExpiry || "No permit"}`,
      newValue: `${nextFields.insurancePolicyNo || "No policy"} | ${nextFields.permitExpiry || "No permit"}`,
      remark: "Admin updated policy and compliance records"
    }), `${vehicle.regNo} compliance saved`);
  };

  const updateVehicleCombination = (event, vehicleId) => {
    event.preventDefault();
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return;
    const partnerId = new FormData(event.currentTarget).get("partnerVehicleId")?.toString() || "";
    const partner = data.vehicles.find((item) => item.id === partnerId);
    if (partner && (partner.clientId !== vehicle.clientId || partner.type === vehicle.type)) {
      setToast("A combination must link one Truck and one Trailer from the same client.");
      return;
    }
    const oldCombinationIds = new Set([vehicle.combinationId, partner?.combinationId].filter(Boolean));
    const oldGroup = data.vehicles
      .filter((item) => item.combinationId && oldCombinationIds.has(item.combinationId))
      .map((item) => item.id);
    const combinationId = partner ? (vehicle.combinationId || partner.combinationId || `combo-${Date.now()}`) : "";
    const linkedIds = new Set([vehicle.id, ...(partner ? [partner.id] : []), ...oldGroup]);
    const updated = {
      ...data,
      vehicles: data.vehicles.map((item) => {
        if (partner && (item.id === vehicle.id || item.id === partner.id)) return { ...item, combinationId };
        if (!partner && linkedIds.has(item.id)) return { ...item, combinationId: "" };
        return item;
      })
    };
    const detail = partner ? `${vehicle.regNo} + ${partner.regNo}` : `${vehicle.regNo} unlinked`;
    const next = notify(updated, partner ? "Truck and Trailer linked" : "Combination unlinked", detail, "fleet");
    persist(withAudit(next, {
      module: "Fleet",
      action: partner ? "Combination Linked" : "Combination Unlinked",
      record: vehicle.regNo,
      oldValue: vehicle.combinationId || "No combination",
      newValue: combinationId || "No combination",
      remark: detail
    }), partner ? `Combination linked: ${detail}` : `${vehicle.regNo} unlinked`);
  };

  const createListing = async (event, vehicleId) => {
    event.preventDefault();
    const vehicle = data.vehicles.find((item) => item.id === vehicleId);
    if (!vehicle) return;
    const form = new FormData(event.currentTarget);
    const photos = await readListingPhotos(form.getAll("photos"));
    const price = Number(form.get("price")?.toString().replace(/\D/g, "")) || liability(vehicle);
    const listing = {
      id: `m-${Date.now()}`,
      vehicleId,
      title: `${vehicle.year} ${vehicle.make} ${vehicle.model}`,
      price,
      location: form.get("location")?.toString().trim() || getDataClient(data, vehicle.clientId)?.city || "Owner location",
      status: "Active",
      condition: form.get("condition")?.toString() || "Good",
      photos
    };
    const updated = {
      ...data,
      listings: [listing, ...data.listings],
      vehicles: data.vehicles.map((item) => (item.id === vehicleId ? { ...item, status: "Listed" } : item))
    };
    const next = notify(updated, "Listing created", listing.title, "marketplace");
    persist(withAudit(next, {
      module: "Marketplace",
      action: "Admin Listing Created",
      record: vehicle.regNo,
      oldValue: vehicle.status,
      newValue: "Listed",
      remark: "Admin created resale listing from fleet"
    }), `Listing created for ${vehicle.regNo}`);
    event.currentTarget.reset();
  };

  const importValidRows = () => {
    const checkedRows = validateStagedImportRows(data.importRows);
    const invalidRows = checkedRows.filter((row) => row.status === "Invalid");
    if (invalidRows.length > 0) {
      const updated = { ...data, importRows: checkedRows };
      const next = notify(updated, "Import validation failed", `${invalidRows.length} rows need correction before final import.`, "import");
      persist(withAudit(next, {
        module: "Excel Import",
        action: "Validation Failed",
        record: "IMP-2026-08-03-001",
        oldValue: "Staged",
        newValue: `${invalidRows.length} invalid rows`,
        remark: "Final import blocked by row-wise validation"
      }), "Fix invalid rows before final import");
      return;
    }
    const validRows = checkedRows.filter((row) => row.status === "Valid" && row.issue !== "Imported and saved");
    if (validRows.length === 0) return;
    const newClient = {
      id: `c-import-${Date.now()}`,
      name: "Patel Freight",
      city: "Ahmedabad",
      phone: "+919900001122",
      callerId: ""
    };
    const newVehicles = validRows.map((row, index) => ({
      id: `v-import-${Date.now()}-${index}`,
      clientId: newClient.id,
      type: row.assetType,
      regNo: row.regNo,
      make: row.assetType === "Truck" ? "Tata" : "DICV",
      model: row.assetType === "Truck" ? "Prima" : "Flatbed",
      year: 2021,
      km: 65000 + index * 12000,
      principal: row.assetType === "Truck" ? 384884 : 981012,
      overdue: 0,
      penalty: 0,
      foreclosure: 15000,
      insuranceExpiry: "2026-09-15",
      permitExpiry: "2026-10-20",
      status: "Active"
    }));
    const updated = {
      ...data,
      clients: data.clients.some((client) => client.name === newClient.name) ? data.clients : [newClient, ...data.clients],
      vehicles: [...newVehicles, ...data.vehicles],
      importRows: data.importRows.map((row) => (row.status === "Valid" ? { ...row, issue: "Imported and saved" } : row))
    };
    const next = notify(updated, "Excel import completed", `${newVehicles.length} assets created.`, "fleet");
    persist(withAudit(next, {
      module: "Excel Import",
      action: "Final Import",
      record: "IMP-2026-08-03-001",
      oldValue: "Staged",
      newValue: `${newVehicles.length} assets created`,
      remark: "Valid rows imported into database data"
    }), `${newVehicles.length} assets imported`);
  };

  const importClientExcel = async (event, clientId) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      const parsedRows = await parseClientExcelFile(file);
      const rows = validateClientExcelRows(parsedRows, data, clientId);
      if (rows.length === 0) {
        setSaveStatus("Error");
        setToast("No valid rows were found in the Excel file.");
        return;
      }
      const client = getDataClient(data, clientId);
      const imported = {
        id: `ci-${Date.now()}`,
        clientId,
        fileName: file.name,
        importedAt: new Date().toLocaleString("en-IN"),
        rows
      };
      const importableRows = rows.filter((row) => row.validationStatus === "Valid" && row.regNo && !isBodyRow(row));
      const importStamp = Date.now();
      const newVehicles = importableRows.map((row, index) => excelRowToVehicle(row, clientId, `v-excel-${importStamp}-${index}`));
      const newDueTasks = importableRows
        .map((row, index) => ({ row, index }))
        .filter(({ row }) => toNumber(row.emiAmount) > 0)
        .map(({ row, index }) => excelRowToDueTask(row, clientId, `v-excel-${importStamp}-${index}`, `d-excel-${importStamp}-${index}`));
      const updated = {
        ...data,
        vehicles: [...newVehicles, ...data.vehicles],
        dueTasks: [...newDueTasks, ...data.dueTasks],
        clientImports: [imported, ...(data.clientImports ?? [])]
      };
      const invalidCount = rows.filter((row) => row.validationStatus === "Invalid").length;
      const next = notify(updated, "Client Excel validated", `${newVehicles.length} vehicles saved, ${invalidCount} rows rejected for ${client?.name ?? "client"}.`, "clients");
      persist(withAudit(next, {
        module: "Client Excel",
        action: "Validated and Imported",
        record: client?.name ?? clientId,
        oldValue: "-",
        newValue: `${newVehicles.length} valid rows, ${invalidCount} invalid rows`,
        remark: `${file.name} | Row-wise validation completed`
      }), "");
      if (invalidCount > 0) {
        setSaveStatus("Error");
        setToast(`${newVehicles.length} valid rows saved. ${invalidCount} rows rejected; see validation details below.`);
      }
    } catch (error) {
      setSaveStatus("Error");
      setToast(error.message || "Excel import failed");
    } finally {
      event.target.value = "";
    }
  };

  const importClientPdf = async (event, clientId) => {
    const file = event.target.files?.[0];
    if (!file) return;
    try {
      setSaveStatus("Reading");
      setToast("Reading bank document...");
      const pdfText = await extractPdfTextWithOcr(file);
      if (!pdfText.trim()) throw new Error("The PDF did not contain readable text or finance fields.");
      const client = getDataClient(data, clientId);
      const importedAssets = (data.clientImports ?? [])
        .filter((item) => item.clientId === clientId)
        .flatMap((item) => item.rows ?? []);
      const pdfRow = normalizePdfFinanceRow(parseBankPdfText(pdfText, file.name));
      const profileAgreement = findProfileAgreementInText(pdfText, importedAssets) ||
        inferProfileAgreementFromPdf(pdfRow, importedAssets);
      if (profileAgreement) pdfRow.loanAccount = profileAgreement;

      const isPdfRowMatch = (row) => (
        agreementMatches(row.loanAccount, pdfRow.loanAccount) ||
        (pdfRow.regNo && normalizeRegNo(baseRegNo(row.regNo)) === normalizeRegNo(baseRegNo(pdfRow.regNo)))
      );
      const matchedRow = importedAssets.find(isPdfRowMatch) || {};

      const matchingVehicle = data.vehicles.find((vehicle) => (
        vehicle.clientId === clientId && (
          agreementMatches(vehicle.loanAccount, pdfRow.loanAccount) ||
          (pdfRow.regNo && normalizeRegNo(baseRegNo(vehicle.regNo)) === normalizeRegNo(baseRegNo(pdfRow.regNo)))
        )
      ));

      if (!isValidAgreementValue(pdfRow.loanAccount)) {
        pdfRow.loanAccount = profileAgreement || matchedRow?.loanAccount || matchingVehicle?.loanAccount || pdfRow.regNo || `AGMT-${Date.now().toString().slice(-6)}`;
      }
      const pdfMergeRow = sanitizePdfRowForMerge(pdfRow);
      // Registration entered by the user takes priority over the bank's text.
      const savedRegNo = [matchingVehicle?.regNo, matchedRow?.regNo]
        .find((value) => String(value ?? "").trim() && String(value).trim() !== "-");
      const defaultRegNo = savedRegNo || pdfRow.regNo || pdfRow.loanAccount || `VEH-${Date.now().toString().slice(-6)}`;

      const mergedRow = {
        ...mergePdfFinanceAsset(matchedRow, pdfMergeRow),
        regNo: defaultRegNo,
        pdfImportedAt: new Date().toLocaleString("en-IN")
      };

      const vehicleId = matchingVehicle?.id || `v-pdf-${Date.now()}`;
      const nextVehicle = matchingVehicle
        ? mergePdfIntoVehicle(matchingVehicle, { ...mergedRow, regNo: defaultRegNo })
        : excelRowToVehicle({ ...mergedRow, regNo: defaultRegNo }, clientId, vehicleId);

      const pdfDueTask = nextVehicle
        ? buildPdfEmiDueTask(mergedRow, clientId, nextVehicle.id, nextVehicle.callerId || client?.callerId || "")
        : null;
      const existingEmiDue = nextVehicle && pdfDueTask
        ? data.dueTasks.find((task) => task.vehicleId === nextVehicle.id && task.type === "EMI" && task.status !== "Closed" && dueTaskDateKey(task.dueDate) === dueTaskDateKey(pdfDueTask.dueDate))
          || data.dueTasks.find((task) => task.vehicleId === nextVehicle.id && task.type === "EMI" && task.status !== "Closed" && String(task.id).startsWith("d-pdf-"))
        : null;
      const updatedDueTasks = pdfDueTask
        ? existingEmiDue
          ? data.dueTasks.map((task) => task.id === existingEmiDue.id ? { ...task, ...pdfDueTask, id: task.id, status: task.status === "Due" || task.status === "Overdue" ? pdfDueTask.status : task.status } : task)
          : [pdfDueTask, ...data.dueTasks]
        : data.dueTasks;

      let rowUpdated = false;
      const updatedClientImports = (data.clientImports ?? []).map((item) => {
        if (item.clientId !== clientId) return item;
        return {
          ...item,
          rows: (item.rows ?? []).map((row) => {
            const isMatch = agreementMatches(row.loanAccount, pdfRow.loanAccount) ||
              (pdfRow.regNo && normalizeRegNo(baseRegNo(row.regNo)) === normalizeRegNo(baseRegNo(pdfRow.regNo)));
            if (!isMatch) return row;
            rowUpdated = true;
            return {
              ...mergePdfFinanceAsset(row, pdfMergeRow),
              pdfImportedAt: mergedRow.pdfImportedAt
            };
          })
        };
      });

      if (!rowUpdated) {
        updatedClientImports.unshift({
          id: `ci-pdf-${Date.now()}`,
          clientId,
          fileName: file.name,
          importedAt: mergedRow.pdfImportedAt,
          rows: [{ ...mergedRow, regNo: defaultRegNo }]
        });
      }

      const updated = {
        ...data,
        clientImports: updatedClientImports,
        vehicles: matchingVehicle
          ? data.vehicles.map((vehicle) => vehicle.id === nextVehicle.id ? nextVehicle : vehicle)
          : [nextVehicle, ...data.vehicles],
        dueTasks: updatedDueTasks
      };
      const scheduleMessage = nextVehicle?.emiSchedule?.length
        ? ` ${nextVehicle.emiSchedule.length} EMI schedule rows and the next due date were calculated.`
        : " EMI schedule was not found; add EMI start and amount if the PDF has no readable schedule.";
      const next = notify(updated, "Bank PDF imported", `${mergedRow.loanAccount} data saved for ${client?.name ?? "client"}.${scheduleMessage}`, "clients");
      persist(withAudit(next, {
        module: "Bank PDF",
        action: "Imported",
        record: mergedRow.loanAccount,
        oldValue: "-",
        newValue: mergedRow.regNo || "Agreement matched",
        remark: file.name
      }), rowUpdated ? `${mergedRow.loanAccount} row updated` : `${mergedRow.loanAccount} PDF data updated`);
    } catch (error) {
      setSaveStatus("Error");
      console.error("PDF import failed", error);
      const detail = String(error?.message || error || "").trim();
      setToast(detail ? `PDF import failed: ${detail.slice(0, 180)}` : "PDF import failed. Please try a readable PDF or image.");
    } finally {
      event.target.value = "";
    }
  };

  const addManualClientVehicle = (event, clientId) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const row = {
      ...mapManualVehicleForm(form),
      srNo: String(data.vehicles.filter((vehicle) => vehicle.clientId === clientId).length + 1)
    };
    if (!row.regNo) {
      setSaveStatus("Error");
      setToast("Regt. No. required hai");
      return;
    }
    const manualRow = {
      ...row,
      // Manual Entry must keep the amount entered by the user.
      bankClosingPrincipal: row.closingPrincipal || ""
    };
    const bodyRow = row.bodyFinanceStatus || row.bodyFinancier || row.bodyLoanAccount || row.bodyLoanAmount || row.bodyEmiAmount
      ? {
          ...row,
          regNo: `${row.regNo} BODY`,
          financeStatus: row.bodyFinanceStatus,
          financier: row.bodyFinancier,
          loanAccount: row.bodyLoanAccount,
          loanAmount: row.bodyLoanAmount,
          emiAmount: row.bodyEmiAmount,
          interestRate: row.bodyInterestRate,
          tenure: row.bodyTenure,
          paidEmi: row.bodyPaidEmi,
          emiStart: row.bodyEmiStart,
          emiEnd: row.bodyEmiEnd,
          closingPrincipal: row.bodyClosingPrincipal || "",
          bankClosingPrincipal: row.bodyClosingPrincipal || ""
        }
      : null;
    const client = getDataClient(data, clientId);
    const importStamp = Date.now();
    const vehicle = excelRowToVehicle(manualRow, clientId, `v-manual-${importStamp}`);
    if (row.closingPrincipal) vehicle.principal = toNumber(row.closingPrincipal);
    const dueTask = toNumber(manualRow.emiAmount) > 0
      ? excelRowToDueTask(manualRow, clientId, vehicle.id, `d-manual-${importStamp}`)
      : null;
    const imported = {
      id: `ci-manual-${importStamp}`,
      clientId,
      fileName: "Manual Entry",
      importedAt: new Date().toLocaleString("en-IN"),
      rows: bodyRow ? [manualRow, bodyRow] : [manualRow]
    };
    const updated = {
      ...data,
      vehicles: [vehicle, ...data.vehicles],
      dueTasks: dueTask ? [dueTask, ...data.dueTasks] : data.dueTasks,
      clientImports: [imported, ...(data.clientImports ?? [])]
    };
    const next = notify(updated, "Manual vehicle added", `${manualRow.regNo} saved for ${client?.name ?? "client"}.`, "clients");
    persist(withAudit(next, {
      module: "Client Manual Entry",
      action: "Created",
      record: manualRow.regNo,
      oldValue: "-",
      newValue: client?.name ?? clientId,
      remark: "Manual vehicle finance details added"
    }), "");
    event.currentTarget.reset();
  };

  const deleteClientVehicle = (row, clientId) => {
    const label = row.regNo || row.loanAccount || "this vehicle";
    if (!window.confirm(`Delete ${label}?`)) return;

    const targetVehicleId = String(row.vehicleId || row.id || "").trim();
    const targetReg = normalizeRegNo(baseRegNo(row.regNo));
    const targetLoan = normalizeAgreement(row.loanAccount);
    const deletedVehicleIds = new Set(
      data.vehicles
        .filter((vehicle) => vehicle.clientId === clientId && (
          (targetVehicleId && vehicle.id === targetVehicleId) ||
          (targetReg && normalizeRegNo(baseRegNo(vehicle.regNo)) === targetReg)
        ))
        .map((vehicle) => vehicle.id)
    );
    const updatedVehicles = data.vehicles.filter((vehicle) => {
      const sameClient = vehicle.clientId === clientId;
      const sameVehicle = targetVehicleId && vehicle.id === targetVehicleId;
      const sameReg = targetReg && normalizeRegNo(baseRegNo(vehicle.regNo)) === targetReg;
      return !(sameClient && (sameVehicle || sameReg));
    });
    const updatedImports = (data.clientImports ?? [])
      .map((item) => {
        if (item.clientId !== clientId) return item;
        return {
          ...item,
          rows: (item.rows ?? []).filter((importRow) => {
            const sameReg = targetReg && normalizeRegNo(baseRegNo(importRow.regNo)) === targetReg;
            const sameLoan = targetLoan && normalizeAgreement(importRow.loanAccount) === targetLoan;
            return !(sameReg || sameLoan);
          })
        };
      })
      .filter((item) => item.clientId !== clientId || (item.rows ?? []).length > 0);
    const updatedDues = data.dueTasks.filter((task) => !deletedVehicleIds.has(task.vehicleId));
    const updated = {
      ...data,
      vehicles: updatedVehicles,
      dueTasks: updatedDues,
      clientImports: updatedImports
    };
    const next = notify(updated, "Vehicle deleted", `${label} removed from profile.`, "clients");
    persist(withAudit(next, {
      module: "Client Vehicles",
      action: "Deleted",
      record: label,
      oldValue: "Active",
      newValue: "Removed",
      remark: "Vehicle row deleted from client profile"
    }), `${label} deleted`);
  };

  const clearNotifications = () => {
    persist({
      ...data,
      notifications: data.notifications.map((item) => ({ ...item, unread: false }))
    }, "Notifications marked read");
  };

  const resetDemoData = () => {
    persist(emptyAppData(), "Workspace data cleared");
  };

  const deleteClientAndAccount = async (client) => {
    if (!window.confirm(`Delete ${client.name} and their account?`)) return;
    setSaveStatus("Deleting");
    setToast(`Deleting ${client.name}...`);
    mutationVersionRef.current += 1;
    pendingMutationCountRef.current += 1;
    try {
      const response = await enqueueMutation(() => fetch(`${API_BASE}/api/clients/${encodeURIComponent(client.id)}`, { method: "DELETE", headers: authHeaders(session.token) }));
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Delete failed");
      const clientVehicles = data.vehicles.filter((v) => v.clientId === client.id);
      const removedVehicleIds = new Set(clientVehicles.map((vehicle) => vehicle.id));
      const removedDueIds = new Set(data.dueTasks.filter((task) => task.clientId === client.id).map((task) => task.id));
      const updated = {
        ...data,
        clients: data.clients.filter((item) => item.id !== client.id),
        vehicles: data.vehicles.filter((v) => v.clientId !== client.id),
        dueTasks: data.dueTasks.filter((t) => t.clientId !== client.id),
        listings: data.listings.filter((l) => !removedVehicleIds.has(l.vehicleId)),
        clientImports: (data.clientImports ?? []).filter((item) => item.clientId !== client.id),
        documents: (data.documents ?? []).filter((item) => item.clientId !== client.id),
        verificationItems: (data.verificationItems ?? []).filter((item) => !removedDueIds.has(item.taskId)),
        saleClosings: (data.saleClosings ?? []).filter((item) => item.clientId !== client.id),
        marketplaceThreads: (data.marketplaceThreads ?? []).filter((item) => item.buyerClientId !== client.id && item.sellerClientId !== client.id),
        callerActivities: (data.callerActivities ?? []).filter((item) => !removedDueIds.has(item.taskId))
      };
      persist(withAudit(updated, {
        module: "Clients",
        action: "Deleted",
        record: client.name,
        oldValue: "Active",
        newValue: "Removed",
        remark: "Customer account and client data deleted from database"
      }), `${client.name} deleted`);
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Delete failed");
    } finally {
      pendingMutationCountRef.current = Math.max(0, pendingMutationCountRef.current - 1);
    }
  };

  const saveCommonPassword = async (event) => {
    event.preventDefault();
    const password = new FormData(event.currentTarget).get("commonPassword")?.toString().trim() ?? "";
    if (!password) {
      setToast("Password required");
      return;
    }
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/common-password`, {
        method: "PUT",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ password })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Failed to update common password");
      setSaveStatus("Saved");
      setToast("Common customer password updated");
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Failed to update common password");
    }
  };

  const saveReminderSettings = async (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const reminderSettings = {
      enabled: form.get("reminderEnabled") === "on",
      intervalHours: Number(form.get("reminderIntervalHours")) || 24,
      windowDays: Number(form.get("reminderWindowDays")) || 30
    };
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/reminder-settings`, {
        method: "PUT",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify(reminderSettings)
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Failed to save reminder settings");
      setData((current) => ({ ...current, reminderSettings: result.reminderSettings ?? reminderSettings }));
      setSaveStatus("Database");
      setToast("Reminder settings updated");
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Failed to save reminder settings");
    }
  };

  const saveRolePermissions = async (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const roles = ["Admin", "Caller", "Customer"];
    const currentRows = Array.isArray(data.rolePermissions) ? data.rolePermissions : permissionRows;
    const rolePermissions = currentRows.map((row) => {
      const values = roles.map((role) => {
        const columnIndex = role === "Admin" ? 1 : role === "Caller" ? 3 : 4;
        return form.get(`permission-${role}-${row[0]}`)?.toString() || row[columnIndex] || "No";
      });
      return [row[0], values[0], row[2] || "No", values[1], values[2]];
    });

    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/permissions`, {
        method: "PUT",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ rolePermissions })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Failed to save permissions");
      setData((current) => ({ ...current, rolePermissions: result.rolePermissions ?? rolePermissions }));
      setSaveStatus("Database");
      setToast("Permissions updated");
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Failed to save permissions");
    }
  };

  const createCallerAccount = async (event) => {
    event.preventDefault();
    const form = new FormData(event.currentTarget);
    const name = form.get("newCallerName")?.toString().trim() ?? "";
    const email = form.get("newCallerEmail")?.toString().trim() ?? "";
    const password = form.get("newCallerPassword")?.toString().trim() ?? "";
    if (!name || !email || !password) {
      setToast("Caller name, email and password are required");
      return;
    }
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/users`, {
        method: "POST",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ name, email, password, role: "Caller" })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Failed to create caller");
      const backendData = await fetchBackendData(session.token);
      setData(backendData);
      setSaveStatus("Saved");
      setToast(`Caller created: ${result.email}`);
      event.currentTarget.reset();
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Failed to create caller");
    }
  };

  const deleteCallerAccount = async (caller) => {
    if (!window.confirm(`Delete caller ${caller.name}? Their assigned dues will become unassigned.`)) return;
    setSaveStatus("Deleting");
    mutationVersionRef.current += 1;
    pendingMutationCountRef.current += 1;
    try {
      const response = await enqueueMutation(() => fetch(`${API_BASE}/api/users/${encodeURIComponent(caller.id)}`, { method: "DELETE", headers: authHeaders(session.token) }));
      const result = await response.json().catch(() => ({}));
      if (!response.ok) throw new Error(result.error || "Caller could not be deleted.");
      const updated = {
        ...data,
        users: (data.users ?? []).filter((item) => item.id !== caller.id),
        dueTasks: data.dueTasks.map((task) => task.callerId === caller.id ? { ...task, callerId: "" } : task),
        clients: data.clients.map((client) => client.callerId === caller.id ? { ...client, callerId: "" } : client)
      };
      setData(updated);
      setSaveStatus("Database");
      setToast(`${caller.name} deleted`);
    } catch (error) {
      setSaveStatus("Error");
      setToast(error.message || "Caller could not be deleted.");
    } finally {
      pendingMutationCountRef.current = Math.max(0, pendingMutationCountRef.current - 1);
    }
  };

  const runCallerAssignment = async (event) => {
    event.preventDefault();
    const mode = new FormData(event.currentTarget).get("assignmentMode")?.toString() || "permanent-client";
    setSaveStatus("Saving");
    try {
      const response = await fetch(`${API_BASE}/api/caller-assignment/run`, {
        method: "POST",
        headers: authHeaders(session.token, { "Content-Type": "application/json" }),
        body: JSON.stringify({ mode })
      });
      const result = await response.json();
      if (!response.ok) throw new Error(result.error || "Caller assignment failed");
      const backendData = await fetchBackendData(session.token);
      setData(backendData);
      setSaveStatus("Database");
      setToast(`Assigned ${result.assigned} due task(s), skipped ${result.skipped}`);
    } catch (err) {
      setSaveStatus("Error");
      setToast(err.message || "Caller assignment failed");
    }
  };

  const exportData = () => {
    const blob = new Blob([JSON.stringify(data, null, 2)], { type: "application/json" });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "kuber-finance-admin-data.json";
    link.click();
    URL.revokeObjectURL(url);
    setToast("Data export downloaded");
  };

  const openSection = (key) => {
    if (key !== "client-profile") {
      setSelectedClientId(null);
    }
    setSection(key);
  };

  const openClientProfile = (clientId) => {
    setSelectedClientId(clientId);
    setSection("client-profile");
  };

  return (
    <main className="app-shell customer-shell">
      <aside className="sidebar customer-sidebar">
        <div className="brand">
          <span className="brand-mark">
            <KuberBrandMark />
          </span>
          <div>
            <strong>My Fleet</strong>
            <small>Admin web console</small>
          </div>
        </div>
        <nav className="sidebar-nav" aria-label="Admin modules">
          {visibleNavItems.map(([key, label, icon]) => (
            <button className={(section === key || (section === "client-profile" && key === "clients")) ? "active" : ""} key={key} onClick={() => openSection(key)}>
              <Icon name={icon} />
              {label}
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <div className="sidebar-footer-inner">
            <div className="sidebar-footer-info">
              <strong>{session.name}</strong>
              <span>{session.role}</span>
            </div>
            <span className="sidebar-footer-avatar" aria-hidden="true">
              {String(session.name || "A").slice(0, 1)}
            </span>
          </div>
          <button className="logout-button" type="button" onClick={onLogout} title="Sign out" aria-label="Sign out">
            <Icon name="logout" />
            <span>Sign out</span>
          </button>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p>Admin access</p>
            <h1>{titleFor(section)}</h1>
          </div>
          <div className={`save-pill ${saveStatus.toLowerCase()}`}>
            <Icon name="cloud" />
            {saveStatus}{lastSavedAt ? ` at ${lastSavedAt}` : ""}
          </div>
        </header>

        {section === "dashboard" && <Dashboard data={data} totals={totals} setSection={openSection} />}
        {section === "clients" && <Clients data={data} addClient={addClient} openClientProfile={openClientProfile} deleteClientAndAccount={deleteClientAndAccount} />}
        {section === "client-profile" && <ClientProfile data={data} clientId={selectedClientId} backToClients={() => openSection("clients")} importClientExcel={importClientExcel} importClientPdf={importClientPdf} addManualClientVehicle={addManualClientVehicle} deleteClientVehicle={deleteClientVehicle} deleteClientAndAccount={deleteClientAndAccount} />}
        {toast && <button className="toast" onClick={() => setToast("")}>{toast}</button>}

        {section === "fleet" && <Fleet data={data} updateVehicleFinance={updateVehicleFinance} updateVehicleCompliance={updateVehicleCompliance} updateVehicleCombination={updateVehicleCombination} createListing={createListing} />}
        {section === "dues" && <Dues data={data} openClientProfile={openClientProfile} />}
        {section === "verification" && <Verification data={data} updateTaskStatus={updateTaskStatus} />}
        {section === "caller" && <CallerQueue data={data} saveCallerOutcome={saveCallerOutcome} openWhatsApp={openWhatsApp} updateWhatsAppStatus={updateWhatsAppStatus} />}
        {section === "marketplace" && <Marketplace data={data} updateListingStatus={updateListingStatus} saveSaleClosing={saveSaleClosing} openFleet={() => openSection("fleet")} />}
        {section === "chat-requests" && <MarketplaceChat chat={marketplaceChat} admin />}
        {section === "reports" && <Reports data={data} activeReport={activeReport} setActiveReport={setActiveReport} />}
        {section === "import" && <ImportRows data={data} importValidRows={importValidRows} lastSavedAt={lastSavedAt} />}
        {section === "settings" && <Settings data={data} lastSavedAt={lastSavedAt} saveStatus={saveStatus} clearNotifications={clearNotifications} resetDemoData={resetDemoData} exportData={exportData} saveCommonPassword={saveCommonPassword} saveReminderSettings={saveReminderSettings} saveRolePermissions={saveRolePermissions} createCallerAccount={createCallerAccount} deleteCallerAccount={deleteCallerAccount} runCallerAssignment={runCallerAssignment} saveWhatsAppTemplates={saveWhatsAppTemplates} setSection={openSection} />}
      </section>
    </main>
  );
}

function Dashboard({ data, totals, setSection }) {
  return (
    <section className="dashboard-page">
      <section className="metrics dashboard-metrics">
        <Metric label="Clients" value={data.clients.length} icon="people" onClick={() => setSection("clients")} />
        <Metric label="Trucks" value={totals.trucks.length} icon="truck" onClick={() => setSection("fleet")} />
        <Metric label="Trailers" value={totals.trailers.length} icon="truck" onClick={() => setSection("fleet")} />
        <Metric label="Financed Assets" value={totals.financedAssets.length} icon="chart" onClick={() => setSection("fleet")} />
        <Metric label="Closing Principal" value={formatMoney(totals.totalLiability)} icon="chart" onClick={() => setSection("fleet")} />
        <Metric label="EMI Due" value={totals.emiDue.length} icon="calendar" onClick={() => setSection("dues")} />
        <Metric label="EMI Overdue" value={totals.emiOverdue.length} icon="calendar" onClick={() => setSection("dues")} />
        <Metric label="Insurance Expiring" value={totals.insuranceExpiring.length} icon="calendar" onClick={() => setSection("dues")} />
        <Metric label="Active Callers" value={totals.activeCallers.length} icon="phone" onClick={() => setSection("settings")} />
        <Metric label="Calls Due Today" value={totals.callsDueToday.length} icon="phone" onClick={() => setSection("caller")} />
        <Metric label="Sale Listings" value={totals.activeListings.length} icon="store" onClick={() => setSection("marketplace")} />
        <Metric label="Vehicles Sold" value={totals.soldVehicles.length} icon="store" onClick={() => setSection("marketplace")} />
      </section>
      <section className="quick-strip dashboard-actions">
        {[
          ["Add client", "clients", "people"],
          ["Caller queue", "caller", "phone"],
          ["Marketplace", "marketplace", "store"]
        ].map(([label, target, icon]) => (
          <button key={target} onClick={() => setSection(target)}>
            <Icon name={icon} />
            {label}
          </button>
        ))}
      </section>
      <section className="split dashboard-split">
        <Panel title="Priority Work">
          {totals.overdue.map((task) => <DueCard key={task.id} task={task} data={data} compact />)}
          {totals.overdue.length === 0 && <Empty text="No overdue or escalated work." />}
        </Panel>
        <Panel title="Notifications">
          {data.notifications.map((item) => (
            <article className="notice" key={item.id}>
              <strong>{item.title}</strong>
              <span>{item.detail}</span>
            </article>
          ))}
        </Panel>
      </section>
    </section>
  );
}

function Clients({ data, addClient, openClientProfile, deleteClientAndAccount }) {
  const [search, setSearch] = useState("");
  const normalizedSearch = search.trim().toLowerCase();
  const filteredClients = data.clients.filter((client) => {
    if (!normalizedSearch) return true;
    return [client.name, client.email, client.city, client.phone, client.id]
      .filter(Boolean)
      .some((value) => String(value).toLowerCase().includes(normalizedSearch));
  });

  return (
    <section className="clients-layout">
      <section className="client-create-card">
        <div className="client-card-head">
          <span className="client-head-icon"><Icon name="plus" /></span>
          <div>
            <h2>Create Client</h2>
            <span>New record</span>
          </div>
        </div>
        <form className="client-form" onSubmit={addClient}>
          <label>Name<input name="name" placeholder="New Client" required /></label>
          <label>City<input name="city" placeholder="Jaipur" required /></label>
          <label>Phone<input name="phone" placeholder="+919812345678" required /></label>
          <label>Email (login)<input name="email" type="email" placeholder="owner@company.com" required /></label>
          <button type="submit"><Icon name="plus" />Add client & create account</button>
        </form>
      </section>

      <section className="clients-table-card">
        <div className="clients-table-head">
          <div>
            <h2>Client Directory</h2>
            <span>{filteredClients.length} of {data.clients.length} records</span>
          </div>
          <span className="client-total-pill">{data.clients.length} Clients</span>
        </div>
        <div className="clients-directory-toolbar">
          <label className="clients-search-field">
            <span><Icon name="search" />Search client</span>
            <input
              value={search}
              onChange={(event) => setSearch(event.target.value)}
              placeholder="Name, email, city, phone or client ID"
              aria-label="Search clients"
            />
          </label>
          <span className="clients-search-count">{filteredClients.length} shown</span>
        </div>
        <div className="clients-table-scroll">
        <table className="clients-table">
          <thead><tr><th>Client</th><th>City</th><th>Phone</th><th>Vehicles</th><th>Caller</th><th>Actions</th></tr></thead>
          <tbody>
            {filteredClients.map((client) => {
              const vehicleCount = data.vehicles.filter((vehicle) => vehicle.clientId === client.id && !isBodyRegNo(vehicle.regNo)).length;
              const callerName = users.find((user) => user.id === client.callerId)?.name;
              return (
                <tr key={client.id}>
                  <td>
                    <div className="client-name-cell">
                      <span className="client-avatar-small">{String(client.name || "?").slice(0, 1)}</span>
                      <strong>{client.name}</strong>
                    </div>
                  </td>
                  <td><span className="client-city-pill">{client.city}</span></td>
                  <td className="client-phone-cell">{client.phone}</td>
                  <td><span className="client-count-pill">{vehicleCount}</span></td>
                  <td>{callerName}</td>
                  <td>
                    <div className="client-row-actions">
                      <button className="client-profile-button" type="button" onClick={() => openClientProfile(client.id)}>
                        <Icon name="people" />Profile
                      </button>
                      <button className="client-delete-button" type="button" onClick={() => deleteClientAndAccount(client)} title={`Delete ${client.name}`} aria-label={`Delete ${client.name}`}>
                        <Icon name="trash" />Delete
                      </button>
                    </div>
                  </td>
                </tr>
              );
            })}
            {filteredClients.length === 0 && (
              <tr>
                <td className="clients-empty-cell" colSpan="6">No clients match your search.</td>
              </tr>
            )}
          </tbody>
        </table>
        </div>
      </section>
    </section>
  );
}

function ManualVehicleForm({ client, onSubmit }) {
  const [hasBodyFinance, setHasBodyFinance] = useState(false);

  return (
    <form className="manual-entry-form" onSubmit={onSubmit}>
      <section>
        <h3>Registration</h3>
        <label>Regt. No.<input name="regNo" placeholder="GJ 16 AY 7703" required /></label>
        <label>Regt. Owner<input name="owner" defaultValue={client.name.toUpperCase()} /></label>
        <label>Manufacturer<input name="manufacturer" placeholder="AL" /></label>
        <label>Model<input name="model" placeholder="4825" /></label>
        <label>Yr of mfg<input name="yearOfMfg" placeholder="11-2025" /></label>
        <label>Reg. Date<input name="regDate" placeholder="24-02-2026" /></label>
      </section>
      <section>
        <h3>Finance</h3>
        <label>Loan ID<input name="loanId" placeholder="LOAN-GJ01AB1234" /></label>
        <label>Free/Fin Fin<input name="financeStatus" placeholder="FIN" /></label>
        <label>Financier's Name<input name="financier" placeholder="AXIS" /></label>
        <label>Loan Acc. No<input name="loanAccount" placeholder="CVR001313712427" /></label>
        <label>Loan Am.<input name="loanAmount" placeholder="4300000" /></label>
        <label>EMI Am.<input name="emiAmount" placeholder="109736" /></label>
        <label>Interest Rate (%)<input name="interestRate" placeholder="9.41" /></label>
        <label>Tenure<input name="tenure" placeholder="48" /></label>
        <label>Paid Emi<input name="paidEmi" placeholder="7" /></label>
        <label>EMI Start<input name="emiStart" placeholder="01-05-2026" /></label>
        <label>EMI End<input name="emiEnd" placeholder="05-12-2029" /></label>
        <label>Closing Principal<input name="closingPrincipal" placeholder="Enter closing principal" /></label>
      </section>
      <section className="manual-body-finance-section">
        <div className="manual-body-finance-heading">
          <h3>Body Loan Details</h3>
          <div className="body-finance-toggle" role="group" aria-label="Body finance available">
            <span>Available?</span>
            <button type="button" className={!hasBodyFinance ? "active" : ""} onClick={() => setHasBodyFinance(false)} aria-pressed={!hasBodyFinance}>No</button>
            <button type="button" className={hasBodyFinance ? "active" : ""} onClick={() => setHasBodyFinance(true)} aria-pressed={hasBodyFinance}>Yes</button>
          </div>
        </div>
        {hasBodyFinance && (
          <div className="body-finance-fields">
            <label>Regt. No.<input name="bodyRegNo" defaultValue="BODY" /></label>
            <label>Free/Fin Fin<input name="bodyFinanceStatus" placeholder="FIN" /></label>
            <label>Financier's Name<input name="bodyFinancier" placeholder="AXIS" /></label>
            <label>Loan Acc. No<input name="bodyLoanAccount" placeholder="BODY-LOAN-001" /></label>
            <label>Loan Am.<input name="bodyLoanAmount" placeholder="500000" /></label>
            <label>EMI Am.<input name="bodyEmiAmount" placeholder="15000" /></label>
            <label>Interest Rate<input name="bodyInterestRate" placeholder="9.41" /></label>
            <label>Tenure<input name="bodyTenure" placeholder="36" /></label>
            <label>Paid Emi<input name="bodyPaidEmi" placeholder="0" /></label>
            <label>EMI Start<input name="bodyEmiStart" placeholder="01-05-2026" /></label>
            <label>EMI End<input name="bodyEmiEnd" placeholder="01-04-2029" /></label>
            <label>Closing Principal<input name="bodyClosingPrincipal" placeholder="Enter body closing principal" /></label>
          </div>
        )}
      </section>
      <section>
        <h3>Policy</h3>
        <label>Policy Company<input name="policyCompany" placeholder="ICICI LOMBARD" /></label>
        <label>Policy No.<input name="policyNo" placeholder="3003/427928269/00/000" /></label>
        <label>Policy Start Date<input name="policyStart" placeholder="02-06-2026" /></label>
        <label>Policy End Date<input name="policyEnd" placeholder="05-02-2027" /></label>
        <label>PUC No.<input name="pucNo" placeholder="Newv4" /></label>
        <label>PUC Expired<input name="pucExpired" placeholder="23-02-2027" /></label>
        <label>Fitness Expired<input name="fitnessExpired" placeholder="18-02-2028" /></label>
      </section>
      <section>
        <h3>Permit</h3>
        <label>Permit No.<input name="permitNo" placeholder="GJ2026-GP-6236B" /></label>
        <label>Permit Issue<input name="permitIssue" placeholder="25-02-2026" /></label>
        <label>Permit Expired<input name="permitExpired" placeholder="24-02-2031" /></label>
        <label>Permit Type<input name="permitType" placeholder="Goods Permit [HGV]" /></label>
        <label>National Permit Expired<input name="nationalPermitExpired" placeholder="20-05-2027" /></label>
      </section>
      <button type="submit"><Icon name="check" />Save Manual Details</button>
    </form>
  );
}

function ClientProfile({ data, clientId, backToClients, importClientExcel, importClientPdf, addManualClientVehicle, deleteClientVehicle, deleteClientAndAccount }) {
  const [entryMode, setEntryMode] = useState("excel");
  const client = getDataClient(data, clientId);

  if (!client) {
    return (
      <section className="stack">
        <button className="ghost-action" type="button" onClick={backToClients}>Back to clients</button>
        <Empty text="Client profile not found." />
      </section>
    );
  }

  const clientVehicles = data.vehicles.filter((vehicle) => vehicle.clientId === client.id);
  // Manual entries are real vehicles even when their registration text contains "body".
  const visibleVehicles = clientVehicles;
  const clientDues = data.dueTasks.filter((task) => task.clientId === client.id);
  // Keep old EMI rows in the database for history, but show only the latest
  // open due for each vehicle/type in the client profile.
  const visibleDues = latestDueTasks(clientDues);
  const clientImports = (data.clientImports ?? []).filter((item) => item.clientId === client.id);
  const importedAssets = clientImports.flatMap((item) => item.rows.map((row) => ({ ...row, importFile: item.fileName, importedAt: item.importedAt })));
  const validationRows = importedAssets.filter((row) => Array.isArray(row.validationIssues));
  const invalidImportRows = validationRows.filter((row) => row.validationStatus === "Invalid");
  const caller = users.find((user) => user.id === client.callerId);
  const excelClosingTotal = importedAssets.reduce((sum, row) => sum + autoClosingPrincipal(row), 0);
  const totalClosing = excelClosingTotal > 0
    ? excelClosingTotal
    : clientVehicles.reduce((sum, vehicle) => sum + liability(vehicle), 0);
  const openAmount = visibleDues.filter((task) => task.status !== "Closed").reduce((sum, task) => sum + task.amount, 0);
  const excelLoanTotal = importedAssets.reduce((sum, row) => sum + toNumber(row.loanAmount), 0);
  const recentActivity = data.callerActivities.filter((activity) =>
    clientDues.some((task) => task.id === activity.taskId)
  );
  const saveManualVehicle = (event) => {
    addManualClientVehicle(event, client.id);
    setEntryMode("excel");
  };

  return (
    <section className="client-profile stack">
      <section className="profile-hero">
        <div className="profile-hero-actions">
          <button className="ghost-action" type="button" onClick={backToClients}>Back to clients</button>
          {deleteClientAndAccount && (
            <button className="ghost-action danger" type="button" onClick={() => deleteClientAndAccount(client)}>
              <Icon name="trash" />Delete client
            </button>
          )}
        </div>
        <div className="profile-identity">
          <div className="profile-avatar">{String(client.name || "?").slice(0, 1)}</div>
          <div className="profile-copy">
            <span>Client profile</span>
            <h2>{client.name}</h2>
            <p>{client.city} | {client.phone} | {client.email || "No email"}</p>
          </div>
        </div>
        <div className="profile-total">
          <span>Closing total</span>
          <strong>{formatMoney(totalClosing)}</strong>
        </div>
      </section>

      <section className="profile-stats">
        <article>
          <Icon name="truck" />
          <span>Vehicles</span>
          <strong>{visibleVehicles.length}</strong>
          <small>Total vehicles</small>
        </article>
        <article>
          <Icon name="calendar" />
          <span>Open dues</span>
          <strong>{visibleDues.filter((task) => task.status !== "Closed").length}</strong>
          <small>Pending records</small>
        </article>
        <article>
          <Icon name="money" />
          <span>Open amount</span>
          <strong>{formatMoney(openAmount)}</strong>
          <small>Total outstanding</small>
        </article>
        <article>
          <Icon name="phone" />
          <span>Caller</span>
          <strong>{caller?.name ?? "Unassigned"}</strong>
          <small>Assigned caller</small>
        </article>
        <article>
          <Icon name="upload" />
          <span>Excel assets</span>
          <strong>{importedAssets.length}</strong>
          <small>Excel files</small>
        </article>
      </section>

      {clientImports.length > 0 && (
        <section className="import-strip">
          <Icon name="upload" />
          <div>
            <strong>{clientImports[0].fileName}</strong>
            <span>{importedAssets.length} Excel rows imported at {clientImports[0].importedAt}</span>
          </div>
          <Badge label="Imported" />
        </section>
      )}

      {validationRows.length > 0 && (
        <section className="import-validation-panel">
          <div className="section-heading">
            <span>Row-wise validation</span>
            <h2>{invalidImportRows.length ? `${invalidImportRows.length} rows need correction` : "All imported rows are valid"}</h2>
          </div>
          <div className="table-wrap">
            <table>
              <thead><tr><th>Reg No</th><th>Status</th><th>Validation result</th></tr></thead>
              <tbody>
                {validationRows.map((row, index) => (
                  <tr key={`${row.regNo}-${row.loanAccount}-${index}`}>
                    <td>{row.regNo || "-"}</td>
                    <td><Badge label={row.validationStatus || "Valid"} /></td>
                    <td>{row.validationIssues?.length ? row.validationIssues.join("; ") : "Ready for import"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </section>
      )}

      <section className="entry-panel">
        <div className="entry-tabs" role="tablist" aria-label="Add client details">
          <button className={entryMode === "excel" ? "active" : ""} type="button" onClick={() => setEntryMode("excel")}>
            <Icon name="upload" />
            Excel Import
          </button>
          <button className={entryMode === "manual" ? "active" : ""} type="button" onClick={() => setEntryMode("manual")}>
            <Icon name="plus" />
            Manual Entry
          </button>
          <button className={entryMode === "pdf" ? "active" : ""} type="button" onClick={() => setEntryMode("pdf")}>
            <Icon name="upload" />
            PDF Upload
          </button>
        </div>
        {entryMode === "excel" ? (
          <div className="excel-import-card">
            <Icon name="upload" />
            <div>
              <strong>Upload vehicle finance Excel</strong>
              <span>Excel columns se registration, finance, policy, PUC, fitness and permit details save honge.</span>
            </div>
            <label className="file-picker">
              <input type="file" accept=".xlsx,.xls" onChange={(event) => importClientExcel(event, client.id)} />
              Choose Excel
            </label>
          </div>
        ) : entryMode === "pdf" ? (
          <div className="excel-import-card pdf-import-card">
            <Icon name="upload" />
            <div>
              <strong>Upload bank PDF or image</strong>
              <span>PDF ya image se EMI schedule read hoga aur Agreement Number se matching vehicle finance data update hoga.</span>
            </div>
            <label className="file-picker">
              <input type="file" accept=".pdf,application/pdf,image/png,image/jpeg,image/webp" onChange={(event) => importClientPdf(event, client.id)} />
              Choose PDF / Image
            </label>
          </div>
        ) : (
          <ManualVehicleForm client={client} onSubmit={saveManualVehicle} />
        )}
      </section>

      <section className="profile-main">
        <Panel title="Client Details">
          <div className="detail-grid">
            <Pair label="Client ID" value={client.id} />
            <Pair label="Name" value={client.name} />
            <Pair label="Email / Login ID" value={client.email || "Not available"} />
            <Pair label="City" value={client.city} />
            <Pair label="Phone" value={client.phone} />
            <Pair label="Assigned caller" value={caller?.name ?? "Unassigned"} />
            <Pair label="Closing total" value={formatMoney(totalClosing)} />
            <Pair label="Excel loan total" value={formatMoney(excelLoanTotal)} />
            <Pair label="Last Excel import" value={clientImports[0]?.importedAt ?? "Not imported"} />
          </div>
        </Panel>

        <Panel title="Dues">
          <div className="profile-dues">
            {visibleDues.map((task) => (
              <DueCard key={task.id} task={task} data={data} compact />
            ))}
            {visibleDues.length === 0 && <Empty text="No dues found for this client." />}
          </div>
        </Panel>
      </section>

      <section className="vehicle-section">
        <div className="section-heading">
          <span>Fleet</span>
          <h2>Client Vehicles</h2>
        </div>
        {visibleVehicles.length > 0 ? (
          <section className="vehicle-browser">
            <VehicleFinanceTable vehicles={visibleVehicles} importedAssets={importedAssets} clientName={client.name} onDeleteVehicle={(row) => deleteClientVehicle(row, client.id)} />
          </section>
        ) : (
          <Empty text="No vehicles found for this client." />
        )}
      </section>

      <Panel title="Recent Caller Activity">
        <div className="profile-timeline">
          {recentActivity.map((activity) => (
            <article key={activity.id}>
              <div />
              <p>
                <strong>{activity.outcome}</strong>
                <span>{activity.notes}</span>
                <small>{activity.channel} | {activity.at}</small>
              </p>
            </article>
          ))}
          {recentActivity.length === 0 && <Empty text="No caller activity recorded yet." />}
        </div>
      </Panel>
    </section>
  );
}

function makeVehicleDetail(vehicle, excelAsset, clientName) {
  const fallback = {
    srNo: "-",
    regNo: vehicle.regNo,
    owner: clientName,
    manufacturer: vehicle.make,
    model: vehicle.model,
    yearOfMfg: String(vehicle.year),
    regDate: "-",
    financeStatus: "-",
    financier: "-",
    loanAccount: "-",
    loanAmount: vehicle.loanAmount ?? "",
    emiAmount: vehicle.emiAmount ?? "",
    tenure: vehicle.tenure ?? "",
    paidEmi: vehicle.paidEmi ?? "",
    emiStart: "-",
    emiEnd: "-",
    closingPrincipal: String(vehicle.principal),
    policyCompany: vehicle.insuranceCompany || "-",
    policyNo: vehicle.insurancePolicyNo || "-",
    policyStart: vehicle.insuranceStart || "-",
    policyEnd: vehicle.insuranceExpiry,
    pucNo: vehicle.pucNo || "-",
    pucExpired: vehicle.pucExpiry || "-",
    fitnessExpired: vehicle.fitnessExpiry || "-",
    permitNo: vehicle.permitNo || "-",
    permitIssue: vehicle.permitIssue || "-",
    permitExpired: vehicle.permitExpiry,
    permitType: vehicle.permitType || "-",
    nationalPermitExpired: vehicle.nationalPermitExpiry || "-",
    remarks: "-"
  };
  const detail = excelAsset ?? fallback;
  const financeNumber = (key) => {
    const value = detail[key];
    return value !== undefined && value !== null && String(value).trim() !== "" && value !== "-"
      ? value : vehicle[key] ?? 0;
  };
  const loanFields = {
    loanId: detail.loanId || vehicle.loanId || `LOAN-${String(vehicle.regNo || vehicle.id).replace(/[^A-Z0-9]/gi, "").toUpperCase()}`,
    loanAccount: detail.loanAccount && detail.loanAccount !== "-" ? detail.loanAccount : vehicle.loanAccount || "",
    financier: detail.financier && detail.financier !== "-" ? detail.financier : vehicle.financier || "",
    loanAmount: financeNumber("loanAmount"),
    emiAmount: financeNumber("emiAmount"),
    interestRate: financeNumber("interestRate"),
    tenure: financeNumber("tenure"),
    paidEmi: financeNumber("paidEmi"),
    emiStart: detail.emiStart && detail.emiStart !== "-" ? detail.emiStart : vehicle.emiStart || "",
    emiEnd: detail.emiEnd && detail.emiEnd !== "-" ? detail.emiEnd : vehicle.emiEnd || ""
  };
  const plan = {
    schedule: vehicle.emiSchedule?.length ? vehicle.emiSchedule : buildLoanPlan(loanFields).schedule,
    history: vehicle.emiHistory?.length ? vehicle.emiHistory : buildLoanPlan(loanFields).history
  };
  return {
    ...detail,
    ...loanFields,
    ...plan,
    policyCompany: detail.policyCompany && detail.policyCompany !== "-" ? detail.policyCompany : vehicle.insuranceCompany || "",
    policyNo: detail.policyNo && detail.policyNo !== "-" ? detail.policyNo : vehicle.insurancePolicyNo || "",
    policyStart: detail.policyStart && detail.policyStart !== "-" ? detail.policyStart : vehicle.insuranceStart || "",
    policyEnd: detail.policyEnd && detail.policyEnd !== "-" ? detail.policyEnd : vehicle.insuranceExpiry || "",
    pucNo: detail.pucNo && detail.pucNo !== "-" ? detail.pucNo : vehicle.pucNo || "",
    pucExpired: detail.pucExpired && detail.pucExpired !== "-" ? detail.pucExpired : vehicle.pucExpiry || "",
    fitnessExpired: detail.fitnessExpired && detail.fitnessExpired !== "-" ? detail.fitnessExpired : vehicle.fitnessExpiry || "",
    permitNo: detail.permitNo && detail.permitNo !== "-" ? detail.permitNo : vehicle.permitNo || "",
    permitIssue: detail.permitIssue && detail.permitIssue !== "-" ? detail.permitIssue : vehicle.permitIssue || "",
    permitExpired: detail.permitExpired && detail.permitExpired !== "-" ? detail.permitExpired : vehicle.permitExpiry || "",
    permitType: detail.permitType && detail.permitType !== "-" ? detail.permitType : vehicle.permitType || "",
    nationalPermitExpired: detail.nationalPermitExpired && detail.nationalPermitExpired !== "-" ? detail.nationalPermitExpired : vehicle.nationalPermitExpiry || "",
    insuranceHistory: vehicle.insuranceHistory ?? [],
    complianceHistory: vehicle.complianceHistory ?? []
  };
}

function Fleet({ data, updateVehicleFinance, updateVehicleCompliance, updateVehicleCombination, createListing }) {
  return (
    <section className="grid-list">
      {data.vehicles.map((vehicle) => (
        <article className="asset" key={vehicle.id}>
          <div className="card-head">
            <div>
              <strong>{vehicle.regNo}</strong>
              <span>{vehicle.year} {vehicle.make} {vehicle.model}</span>
            </div>
            <Badge label={vehicle.status} />
          </div>
          <dl>
            <div><dt>Owner</dt><dd>{getDataClient(data, vehicle.clientId)?.name}</dd></div>
            <div><dt>Type</dt><dd>{vehicle.type}</dd></div>
            <div><dt>KM</dt><dd>{vehicle.km.toLocaleString("en-IN")}</dd></div>
            <div><dt>Closing Principal</dt><dd>{formatMoney(vehicle.principal)}</dd></div>
            <div><dt>Insurance</dt><dd>{vehicle.insurancePolicyNo || formatDisplayDate(vehicle.insuranceExpiry)}</dd></div>
            <div><dt>Permit</dt><dd>{vehicle.permitNo || formatDisplayDate(vehicle.permitExpiry)}</dd></div>
            <div><dt>Combination</dt><dd>{vehicle.combinationId || "Unlinked"}</dd></div>
            {vehicle.status === "Sold" && <div><dt>Sold date</dt><dd>{formatDisplayDate(vehicle.soldDate)}</dd></div>}
          </dl>
          <details>
            <summary>Edit finance</summary>
            <form className="form-grid compact-form" onSubmit={(event) => updateVehicleFinance(event, vehicle.id)}>
              <label>Loan ID<input name="loanId" defaultValue={vehicle.loanId || `LOAN-${String(vehicle.regNo || vehicle.id).replace(/[^A-Z0-9]/gi, "").toUpperCase()}`} /></label>
              <label>Loan Account<input name="loanAccount" defaultValue={vehicle.loanAccount || ""} /></label>
              <label>Financier<input name="financier" defaultValue={vehicle.financier || ""} /></label>
              <label>Loan Amount<input name="loanAmount" type="number" min="0" step="0.01" defaultValue={vehicle.loanAmount || ""} /></label>
              <label>EMI Amount<input name="emiAmount" type="number" min="0" step="0.01" defaultValue={vehicle.emiAmount || ""} /></label>
              <label>Interest %<input name="interestRate" type="number" min="0" step="0.01" defaultValue={vehicle.interestRate || ""} /></label>
              <label>Tenure (months)<input name="tenure" type="number" min="0" step="1" defaultValue={vehicle.tenure || ""} /></label>
              <label>Paid EMI<input name="paidEmi" type="number" min="0" step="1" defaultValue={vehicle.paidEmi || ""} /></label>
              <label>EMI Start<input name="emiStart" type="date" defaultValue={vehicle.emiStart || ""} /></label>
              <label>EMI End<input name="emiEnd" type="date" defaultValue={vehicle.emiEnd || ""} /></label>
              <label>Principal<input name="principal" defaultValue={vehicle.principal} /></label>
              <label>Overdue<input name="overdue" defaultValue={vehicle.overdue} /></label>
              <label>Penalty<input name="penalty" defaultValue={vehicle.penalty} /></label>
              <label>Foreclosure<input name="foreclosure" defaultValue={vehicle.foreclosure} /></label>
              <label>Status<select name="status" defaultValue={vehicle.status}><option>Active</option><option>Listed</option><option>Sold</option></select></label>
              <button type="submit"><Icon name="check" />Save finance</button>
            </form>
          </details>
          <details>
            <summary>Link Truck + Trailer</summary>
            <form className="form-grid compact-form" onSubmit={(event) => updateVehicleCombination(event, vehicle.id)}>
              <label className="span-2">Partner asset
                <select name="partnerVehicleId" defaultValue={data.vehicles.find((item) => item.combinationId && item.combinationId === vehicle.combinationId && item.id !== vehicle.id)?.id || ""}>
                  <option value="">No linked partner</option>
                  {data.vehicles.filter((item) => item.clientId === vehicle.clientId && item.type !== vehicle.type).map((item) => (
                    <option key={item.id} value={item.id}>{item.regNo} - {item.type}{item.combinationId ? ` (${item.combinationId})` : ""}</option>
                  ))}
                </select>
              </label>
              <button type="submit"><Icon name="check" />Save combination</button>
            </form>
          </details>
          <details>
            <summary>Edit insurance & compliance</summary>
            <form className="form-grid compact-form" onSubmit={(event) => updateVehicleCompliance(event, vehicle.id)}>
              <label>Insurance Company<input name="insuranceCompany" defaultValue={vehicle.insuranceCompany || ""} /></label>
              <label>Policy No.<input name="insurancePolicyNo" defaultValue={vehicle.insurancePolicyNo || ""} /></label>
              <label>Policy Start<input name="insuranceStart" type="date" defaultValue={vehicle.insuranceStart || ""} /></label>
              <label>Policy Expiry<input name="insuranceExpiry" type="date" defaultValue={vehicle.insuranceExpiry || ""} /></label>
              <label>Policy Status<select name="insuranceStatus" defaultValue="Active"><option>Active</option><option>Expired</option><option>Pending Verification</option></select></label>
              <label>Permit No.<input name="permitNo" defaultValue={vehicle.permitNo || ""} /></label>
              <label>Permit Issue<input name="permitIssue" type="date" defaultValue={vehicle.permitIssue || ""} /></label>
              <label>Permit Expiry<input name="permitExpiry" type="date" defaultValue={vehicle.permitExpiry || ""} /></label>
              <label>Permit Type<input name="permitType" defaultValue={vehicle.permitType || ""} /></label>
              <label>National Permit Expiry<input name="nationalPermitExpiry" type="date" defaultValue={vehicle.nationalPermitExpiry || ""} /></label>
              <label>PUC No.<input name="pucNo" defaultValue={vehicle.pucNo || ""} /></label>
              <label>PUC Expiry<input name="pucExpiry" type="date" defaultValue={vehicle.pucExpiry || ""} /></label>
              <label>Fitness Expiry<input name="fitnessExpiry" type="date" defaultValue={vehicle.fitnessExpiry || ""} /></label>
              <label>Compliance Status<select name="complianceStatus" defaultValue="Active"><option>Active</option><option>Expired</option><option>Pending Verification</option></select></label>
              <button type="submit"><Icon name="check" />Save policy & compliance</button>
            </form>
          </details>
          <details>
            <summary>Create sale listing</summary>
            <form className="form-grid compact-form" onSubmit={(event) => createListing(event, vehicle.id)}>
              <label>Price<input name="price" placeholder={formatMoney(liability(vehicle))} /></label>
              <label>Location<input name="location" placeholder={getDataClient(data, vehicle.clientId)?.city ?? "City"} /></label>
              <label>Condition<select name="condition" defaultValue="Good"><option>Excellent</option><option>Good</option><option>Average</option></select></label>
              <label className="span-2">Photos<input name="photos" type="file" accept="image/*" multiple /></label>
              <button type="submit"><Icon name="store" />Create listing</button>
            </form>
          </details>
        </article>
      ))}
    </section>
  );
}

function Dues({ data, openClientProfile }) {
  const [expiryFilter, setExpiryFilter] = useState("PUC");
  const smartAlerts = buildSmartAlerts(data);
  const allExpiryAlerts = buildSmartAlerts(data, true);
  const expiredAlerts = smartAlerts.filter((alert) => alert.status === "Expired");
  const soonAlerts = smartAlerts.filter((alert) => ["Today", "Soon"].includes(alert.status));
  const matchesExpiryFilter = (alert) => {
    if (!expiryFilter) return true;
    if (expiryFilter === "Expired") return alert.status === "Expired";
    if (expiryFilter === "Soon") return ["Today", "Soon"].includes(alert.status);
    return alert.type === expiryFilter;
  };
  const visibleSmartAlerts = smartAlerts.filter(matchesExpiryFilter);
  const visibleExpiryAlerts = allExpiryAlerts.filter(matchesExpiryFilter);

  return (
    <section className="dues-board">
      <div className="dues-summary">
        <article className="dues-summary-card summary-alerts">
          <div className="dues-summary-icon"><AlertSvg name="calendar" /></div>
          <div><span>Vehicle documents</span><strong>{allExpiryAlerts.length}</strong><small>Expiry records tracked</small></div>
        </article>
        <article className="dues-summary-card summary-expired">
          <div className="dues-summary-icon"><AlertSvg name="bell" /></div>
          <div><span>Expired</span><strong>{expiredAlerts.length}</strong><small>Need attention now</small></div>
        </article>
        <article className="dues-summary-card summary-soon">
          <div className="dues-summary-icon"><AlertSvg name="calendar" /></div>
          <div><span>Next 10 days</span><strong>{soonAlerts.length}</strong><small>Upcoming renewals</small></div>
        </article>
      </div>
      <section className="smart-alert-panel">
        <div className="smart-alert-head">
          <div className="smart-alert-title">
            <div className="smart-alert-head-icon"><AlertSvg name="calendar" /></div>
            <div><span className="smart-alert-kicker">Priority compliance watch</span><h2>Next 10 days expiry alerts</h2></div>
          </div>
          <span className="smart-alert-count">{visibleSmartAlerts.length} alerts</span>
        </div>
        <div className="smart-alert-list">
          {visibleSmartAlerts.map((alert) => (
            <article className="smart-alert-card" key={alert.id}>
              <div className="smart-alert-icon"><AlertSvg name="calendar" /></div>
              <div className="smart-alert-copy">
                <strong>{alert.vehicle}</strong>
                <span>{alert.client}</span>
              </div>
              <div className="smart-alert-type">{alert.type}</div>
              <div className="smart-alert-date">
                <strong>{formatDisplayDate(alert.date)}</strong>
                <span>{alert.detail}</span>
              </div>
              <Badge label={alert.status} />
            </article>
          ))}
          {visibleSmartAlerts.length === 0 && <div className="smart-alert-empty"><AlertSvg name="check" /><div><strong>No {expiryFilter || "priority"} alerts</strong><span>No matching vehicle expiry alerts found.</span></div></div>}
        </div>
      </section>
      <section className="expiry-nav-panel" aria-label="Expiry categories">
        <div className="expiry-nav-heading">
          <div>
            <span className="smart-alert-kicker">Expiry navigator</span>
            <h2>Expired &amp; upcoming documents</h2>
          </div>
          <span className="smart-alert-count">{allExpiryAlerts.length} records</span>
        </div>
        <nav className="expiry-nav-bar" aria-label="Filter expiry records">
          {[
            ["PUC", "PUC"],
            ["Fitness", "Fitness"],
            ["Permit", "Permit"],
            ["National Permit", "National Permit"],
            ["Policy", "Policy"]
          ].map(([label, value]) => {
            const count = value === "All"
              ? allExpiryAlerts.length
              : value === "Expired"
                ? allExpiryAlerts.filter((alert) => alert.status === "Expired").length
                : value === "Soon"
                  ? allExpiryAlerts.filter((alert) => ["Today", "Soon"].includes(alert.status)).length
                  : allExpiryAlerts.filter((alert) => alert.type === value).length;
            return (
              <button
                type="button"
                key={value}
                className={expiryFilter === value ? "active" : ""}
                aria-pressed={expiryFilter === value}
                onClick={(event) => {
                  event.stopPropagation();
                  setExpiryFilter(value);
                }}
              >
                <span>{label}</span><b>{count}</b>
              </button>
            );
          })}
        </nav>
        <div className="expiry-nav-list">
          {visibleExpiryAlerts.map((alert) => (
              <article
                className="expiry-nav-item clickable-expiry-item"
                key={alert.id}
                role="button"
                tabIndex={0}
                onClick={() => alert.clientId && openClientProfile(alert.clientId)}
                onKeyDown={(event) => {
                  if ((event.key === "Enter" || event.key === " ") && alert.clientId) {
                    event.preventDefault();
                    openClientProfile(alert.clientId);
                  }
                }}
                title="Open customer profile"
              >
                <div><strong>{alert.type} {alert.status === "Expired" ? "Expired" : "Expiry"}</strong><span>{alert.vehicle} · {alert.client}</span></div>
                <div><strong>{formatDisplayDate(alert.date)}</strong><span>{alert.detail}</span></div>
                <Badge label={alert.status} />
              </article>
            ))}
          {visibleExpiryAlerts.length === 0 && <div className="expiry-nav-empty">No expiry records found.</div>}
        </div>
      </section>
    </section>
  );
}

function DueCard({ task, data, callers = [], updateTaskStatus, assignCallerToTask, compact = false }) {
  const client = getDataClient(data, task.clientId);
  const vehicle = getDataVehicle(data, task.vehicleId);
  return (
    <article className={compact ? "due compact" : "due"}>
      {compact && (
        <div className="due-icon">
          <Icon name="calendar" />
        </div>
      )}
      <div className="due-copy">
        <strong>{task.type} - {vehicle?.regNo}</strong>
        <span>{client?.name} | Due {formatDisplayDate(task.dueDate)}</span>
        {!compact && (
          <div className="due-assignment">
            <label htmlFor={`caller-${task.id}`}>Caller</label>
            <select id={`caller-${task.id}`} value={task.callerId || ""} disabled={!callers.length || task.status === "Closed"} onChange={(event) => assignCallerToTask(task.id, event.target.value)}>
              <option value="">Unassigned</option>
              {callers.map((caller) => <option value={caller.id} key={caller.id}>{caller.name}</option>)}
            </select>
          </div>
        )}
      </div>
      <div className="due-status"><Badge label={task.status} /></div>
      <b className="due-amount">{formatMoney(task.amount)}</b>
      {!compact && (
        <div className="actions">
          <button onClick={() => updateTaskStatus(task.id, "Promise-to-Pay", "Promise to pay marked by Admin")}>Promise</button>
          <button onClick={() => updateTaskStatus(task.id, "Verification Pending", "Documents requested from owner")}>More info</button>
          <button onClick={() => updateTaskStatus(task.id, "Closed", "Closed by Admin")}>Close</button>
        </div>
      )}
    </article>
  );
}

function VehicleFinanceTable({ vehicles, importedAssets, clientName, onDeleteVehicle }) {
  const [selectedRow, setSelectedRow] = useState(null);
  const [vehicleSearch, setVehicleSearch] = useState("");
  const searchTerm = vehicleSearch.trim().toLowerCase();
  const tableRows = vehicles.flatMap((vehicle, index) => {
    const matchingAssets = importedAssets.filter((asset) => !isBodyRow(asset) && normalizeRegNo(baseRegNo(asset.regNo)) === normalizeRegNo(baseRegNo(vehicle.regNo)));
    const excelAsset = matchingAssets.find((asset) => asset.pdfImportedAt || asset.bankClosingPrincipal) ?? matchingAssets[0];
    const bodyAsset = importedAssets.find((asset) => isBodyRow(asset) && normalizeRegNo(baseRegNo(asset.regNo)) === normalizeRegNo(baseRegNo(vehicle.regNo)));
    const searchableText = [
      vehicle.regNo,
      vehicle.make,
      vehicle.model,
      clientName,
      excelAsset?.owner,
      excelAsset?.regNo,
      excelAsset?.loanAccount,
      bodyAsset?.loanAccount
    ].filter(Boolean).join(" ").toLowerCase();
    if (searchTerm && !searchableText.includes(searchTerm)) return [];
    const bodyDetail = bodyAsset ? {
      ...bodyAsset,
      displaySrNo: "",
      isBodyDetail: true
    } : null;
    const mainAsset = {
      ...makeVehicleDetail(vehicle, excelAsset, clientName),
      vehicleId: vehicle.id,
      displaySrNo: String(index + 1),
      isBodyDetail: false,
      bodyDetail
    };
    if (bodyDetail) bodyDetail.parentDetail = mainAsset;
    return [mainAsset, bodyDetail].filter(Boolean);
  });
  const firstRow = tableRows[0];

  return (
    <article className="vehicle-detail-card">
      <div className="card-head">
        <div>
          <strong>{tableRows.filter((row) => !row.isBodyDetail).length} Vehicles</strong>
          <span>{clientName || firstRow?.owner}</span>
        </div>
        <Badge label="All" />
      </div>
      <div className="vehicle-table-toolbar">
        <label className="vehicle-search-box">
          <Icon name="search" />
          <input
            type="search"
            value={vehicleSearch}
            onChange={(event) => setVehicleSearch(event.target.value)}
            placeholder="Search vehicle number or name"
            aria-label="Search vehicles by number or name"
          />
        </label>
      </div>
      <div className="excel-record-table">
        <table>
          <thead>
            <tr>
              <th>View</th>
              <th>Delete</th>
              <th>Regt. No.</th>
              <th>EMI Am.</th>
              <th>Tenure</th>
              <th>Paid Emi</th>
              <th>Closing Principal</th>
              <th>Policy Company</th>
              <th>Policy No.</th>
              <th>Policy End Date</th>
              <th>PUC Expired</th>
              <th>Fitness Expired</th>
              <th>Permit Expired</th>
            </tr>
          </thead>
          <tbody>
            {tableRows.map((row) => (
              <tr className={row.isBodyDetail ? "body-row" : ""} key={`${row.regNo || "vehicle"}-${row.loanAccount || row.displaySrNo}`}>
                <td>
                  {!row.isBodyDetail && (
                    <button className="icon-button view-button" type="button" title="View details" aria-label={`View details for ${row.regNo || row.loanAccount || "vehicle"}`} onClick={() => setSelectedRow(row)}>
                      <Icon name="eye" />
                    </button>
                  )}
                </td>
                <td>
                  {!row.isBodyDetail && (
                    <button className="icon-button delete-row-button" type="button" title="Delete vehicle" aria-label={`Delete ${row.regNo || row.loanAccount || "vehicle"}`} onClick={() => onDeleteVehicle?.(row)}>
                      <Icon name="trash" />
                    </button>
                  )}
                </td>
                <td>
                  {row.isBodyDetail ? (
                    <span className="body-reg"><span className="body-connector" />BODY</span>
                  ) : row.regNo || "-"}
                </td>
                <td>{formatPlainMoney(row.emiAmount)}</td>
                <td>{row.tenure || "-"}</td>
                <td>{effectivePaidEmi(row) || "-"}</td>
                <td>{formatAutoClosingPrincipal(row)}</td>
                <td>{row.policyCompany || "-"}</td>
                <td>{row.policyNo || "-"}</td>
                <td>{formatDisplayDate(row.policyEnd)}</td>
                <td>{formatDisplayDate(row.pucExpired)}</td>
                <td>{formatDisplayDate(row.fitnessExpired)}</td>
                <td>{formatDisplayDate(row.permitExpired)}</td>
              </tr>
            ))}
            {tableRows.length === 0 && (
              <tr>
                <td className="table-empty-cell" colSpan="13">No vehicle match found.</td>
              </tr>
            )}
          </tbody>
        </table>
      </div>
      {selectedRow && <VehicleDetailModal row={selectedRow} onClose={() => setSelectedRow(null)} />}
    </article>
  );
}

function VehicleDetailModal({ row, onClose }) {
  return (
    <div className="detail-modal-backdrop" role="presentation" onMouseDown={onClose}>
      <section className="detail-modal" role="dialog" aria-modal="true" aria-labelledby="vehicle-detail-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="detail-modal-head">
          <div>
            <h2 id="vehicle-detail-title">{row.regNo || `Agreement ${row.loanAccount || ""}`}</h2>
            <p>{row.owner || "-"}</p>
          </div>
          <div className="detail-modal-actions">
            <Badge label={row.financeStatus || "Detail"} />
            <button className="icon-button close-button" type="button" title="Close" aria-label="Close details" onClick={onClose}><Icon name="close" /></button>
          </div>
        </div>
        <div className="detail-summary-strip">
          <div>
            <span>Vehicle EMI</span>
            <strong>{formatPlainMoney(row.emiAmount)}</strong>
          </div>
          <div>
            <span>Vehicle Closing</span>
            <strong>{formatAutoClosingPrincipal(row)}</strong>
          </div>
          {row.bodyDetail && (
            <div>
              <span>Body Closing</span>
              <strong>{formatAutoClosingPrincipal(row.bodyDetail)}</strong>
            </div>
          )}
        </div>
        <div className="detail-card-grid">
          <DetailSection title="Registration" rows={[
            ["Sr. No.", row.srNo],
            ["Regt. No.", row.regNo],
            ["Regt. Owner", row.owner],
            ["Manufacturer", row.manufacturer],
            ["Model", row.model],
            ["Yr of mfg", row.yearOfMfg],
            ["Reg. Date", formatDisplayDate(row.regDate)]
          ]} />
          <DetailSection title="Finance" rows={[
            ["Free/Fin Fin", row.financeStatus],
            ["Financier's Name", row.financier],
            ["Loan Acc. No", row.loanAccount],
            ["Loan Am.", formatMaybeMoney(row.loanAmount)],
            ["EMI Am.", formatPlainMoney(row.emiAmount)],
            ["Interest Rate", row.interestRate ? `${row.interestRate}%` : "-"],
            ["Tenure", row.tenure],
            ["Paid Emi", effectivePaidEmi(row) || "-"],
            ["EMI Start", formatDisplayDate(row.emiStart)],
            ["EMI End", formatDisplayDate(row.emiEnd)],
            ["Closing Principal", formatAutoClosingPrincipal(row)]
          ]} />
          {row.bodyDetail && (
            <DetailSection title="Body Loan Details" rows={[
              ["Regt. No.", "BODY"],
              ["Free/Fin Fin", row.bodyDetail.financeStatus],
              ["Financier's Name", row.bodyDetail.financier],
              ["Loan Acc. No", row.bodyDetail.loanAccount],
              ["Loan Am.", formatMaybeMoney(row.bodyDetail.loanAmount)],
              ["EMI Am.", formatPlainMoney(row.bodyDetail.emiAmount)],
              ["Interest Rate", row.bodyDetail.interestRate ? `${row.bodyDetail.interestRate}%` : "-"],
              ["Tenure", row.bodyDetail.tenure],
              ["Paid Emi", effectivePaidEmi(row.bodyDetail) || "-"],
              ["EMI Start", formatDisplayDate(row.bodyDetail.emiStart)],
              ["EMI End", formatDisplayDate(row.bodyDetail.emiEnd)],
              ["Closing Principal", formatAutoClosingPrincipal(row.bodyDetail)]
            ]} />
          )}
          <DetailSection title="Policy" rows={[
            ["Policy Company", row.policyCompany],
            ["Policy No.", row.policyNo],
            ["Policy Start Date", formatDisplayDate(row.policyStart)],
            ["Policy End Date", formatDisplayDate(row.policyEnd)],
            ["PUC No.", row.pucNo],
            ["PUC Expired", formatDisplayDate(row.pucExpired)],
            ["Fitness Expired", formatDisplayDate(row.fitnessExpired)]
          ]} />
          <DetailSection title="Permit" rows={[
            ["Permit No.", row.permitNo],
            ["Permit Issue", formatDisplayDate(row.permitIssue)],
            ["Permit Expired", formatDisplayDate(row.permitExpired)],
            ["Permit Type", row.permitType],
            ["National Permit Expired", formatDisplayDate(row.nationalPermitExpired)]
          ]} />
        </div>
        <div className="loan-ledger-grid">
          <LoanSchedulePanel schedule={row.emiSchedule ?? []} />
          <LoanHistoryPanel history={row.emiHistory ?? []} />
        </div>
        <div className="loan-ledger-grid compliance-ledger-grid">
          <InsuranceHistoryPanel history={row.insuranceHistory ?? []} />
          <ComplianceHistoryPanel history={row.complianceHistory ?? []} />
        </div>
      </section>
    </div>
  );
}

function LoanSchedulePanel({ schedule }) {
  return (
    <article className="loan-ledger-panel">
      <div className="loan-ledger-head"><h3>EMI Schedule</h3><span>{schedule.length} instalments</span></div>
      {schedule.length > 0 ? (
        <div className="loan-ledger-table">
          <div className="loan-ledger-row loan-ledger-row-head"><span>#</span><span>Due date</span><span>Amount</span><span>Status</span></div>
          {schedule.map((entry) => (
            <div className="loan-ledger-row" key={`${entry.installment}-${entry.dueDate}`}>
              <span>{entry.installment}</span>
              <span>{formatDisplayDate(entry.dueDate)}</span>
              <span>{formatPlainMoney(entry.amount)}</span>
              <Badge label={entry.status} />
            </div>
          ))}
        </div>
      ) : <Empty text="Loan tenure and EMI details are not saved yet." />}
    </article>
  );
}

function LoanHistoryPanel({ history }) {
  return (
    <article className="loan-ledger-panel">
      <div className="loan-ledger-head"><h3>EMI Payment History</h3><span>{history.length} paid</span></div>
      {history.length > 0 ? (
        <div className="loan-ledger-table">
          <div className="loan-ledger-row loan-ledger-row-head"><span>#</span><span>Paid on</span><span>Amount</span><span>Reference</span></div>
          {history.map((entry) => (
            <div className="loan-ledger-row" key={`${entry.installment}-${entry.paidOn}-${entry.reference}`}>
              <span>{entry.installment}</span>
              <span>{formatDisplayDate(entry.paidOn)}</span>
              <span>{formatPlainMoney(entry.amount)}</span>
              <span>{entry.reference || "-"}</span>
            </div>
          ))}
        </div>
      ) : <Empty text="No EMI payment history saved yet." />}
    </article>
  );
}

function InsuranceHistoryPanel({ history }) {
  return (
    <article className="loan-ledger-panel">
      <div className="loan-ledger-head"><h3>Insurance History</h3><span>{history.length} policies</span></div>
      {history.length > 0 ? history.map((entry) => (
        <div className="history-entry" key={`${entry.policyNo}-${entry.expiryDate}`}>
          <strong>{entry.company} · {entry.policyNo}</strong>
          <span>{formatDisplayDate(entry.startDate)} to {formatDisplayDate(entry.expiryDate)}</span>
          <Badge label={entry.status} />
        </div>
      )) : <Empty text="No insurance policy history saved yet." />}
    </article>
  );
}

function ComplianceHistoryPanel({ history }) {
  return (
    <article className="loan-ledger-panel">
      <div className="loan-ledger-head"><h3>Compliance History</h3><span>{history.length} records</span></div>
      {history.length > 0 ? history.map((entry) => (
        <div className="history-entry" key={`${entry.kind}-${entry.documentNo}-${entry.expiryDate}`}>
          <strong>{entry.kind} · {entry.documentNo}</strong>
          <span>{entry.issueDate ? `${formatDisplayDate(entry.issueDate)} to ` : "Expiry "}{formatDisplayDate(entry.expiryDate)}</span>
          <Badge label={entry.status} />
        </div>
      )) : <Empty text="No compliance history saved yet." />}
    </article>
  );
}

function DetailSection({ title, rows }) {
  return (
    <article className="detail-section-card">
      <h3>{title}</h3>
      <div>
        {rows.map(([label, value]) => (
          <p key={label}>
            <span>{label}</span>
            <b>{value || "-"}</b>
          </p>
        ))}
      </div>
    </article>
  );
}

function Verification({ data, updateTaskStatus }) {
  return (
    <section className="grid-list">
      {data.verificationItems.map((item) => {
        const task = data.dueTasks.find((entry) => entry.id === item.taskId);
        const proofDocument = (data.documents ?? []).find((document) => document.taskId === item.taskId);
        const proofDetails = (item.details ?? []).map(([label, value]) => [
          label,
          /proof\s*file/i.test(label) && proofDocument?.fileName
            ? proofDocument.fileName
            : value
        ]);
        return (
          <article className="asset" key={item.id}>
            <div className="card-head">
              <div>
                <strong>{item.proofType}</strong>
                <span>{item.submittedBy} | {item.submittedAt}</span>
              </div>
              <Badge label={task?.status ?? "Review"} />
            </div>
            {proofDetails.map(([label, value]) => <Pair key={label} label={label} value={value} />)}
            <div className="verification-proof-file">
              <div className="verification-proof-identity">
                <span className="verification-proof-icon"><ProofFileIcon /></span>
                <div>
                  <strong>{proofDocument?.fileName || "Proof file"}</strong>
                  <small>{proofDocument ? `${proofDocument.mimeType || "Document"} · ${Math.max(1, Math.round((proofDocument.size || 0) / 1024))} KB` : "File record found, but content is unavailable."}</small>
                </div>
              </div>
              {proofDocument?.dataUrl && (
                <div className="verification-proof-actions">
                  <a className="button-link" href={proofDocument.dataUrl} target="_blank" rel="noreferrer"><Icon name="eye" />Open file</a>
                  <a className="button-link secondary" href={proofDocument.dataUrl} download={proofDocument.fileName}><Icon name="download" />Download</a>
                </div>
              )}
            </div>
            <div className="audit-box">
              {item.audit.map(([label, value]) => <Pair key={label} label={label} value={value} />)}
            </div>
            <div className="actions">
              <button onClick={() => updateTaskStatus(task.id, "Closed", "Admin approved submitted proof")}>Approve</button>
              <button onClick={() => updateTaskStatus(task.id, "Verification Pending", "Admin requested additional information")}>More info</button>
              <button className="danger" onClick={() => updateTaskStatus(task.id, "Due", "Admin rejected submitted proof")}>Reject</button>
            </div>
          </article>
        );
      })}
      {data.verificationItems.length === 0 && (
        <div className="verification-empty-state">
          <div className="verification-empty-icon"><Icon name="check" /></div>
          <span className="verification-empty-kicker">All clear</span>
          <h3>No proof is waiting for verification</h3>
          <p>Submitted payment proofs will appear here when they are ready for your review.</p>
        </div>
      )}
    </section>
  );
}

function CallerQueue({ data, saveCallerOutcome, openWhatsApp, updateWhatsAppStatus }) {
  const whatsappTemplates = (data.whatsappTemplates ?? initialWhatsAppTemplates).filter((item) => item.active);
  const whatsappLogs = data.whatsappLogs ?? [];
  return (
    <section className="stack">
      <section className="caller-page-header">
        <div>
          <p className="eyebrow">Collection workspace</p>
          <h2>Caller queue</h2>
          <p>Work through open dues and record the next customer follow-up.</p>
        </div>
        <div className="caller-queue-stats">
          <div><span>Open dues</span><strong>{data.dueTasks.filter((task) => task.status !== "Closed").length}</strong></div>
          <div><span>Follow-ups</span><strong>{data.callerActivities.length}</strong></div>
        </div>
      </section>
      <section className="grid-list caller-grid">
        {data.dueTasks.map((task) => {
          const client = getDataClient(data, task.clientId);
          const history = data.callerActivities.filter((item) => item.taskId === task.id);
          return (
            <article className="asset caller-task-card" key={task.id}>
              <div className="card-head caller-task-head">
                <div>
                  <div className="caller-customer-line"><span className="caller-customer-avatar">{(client?.name || "?").slice(0, 1)}</span><strong>{client?.name || "Unknown customer"}</strong></div>
                  <span className="caller-due-meta">{task.type} <i /> {formatDisplayDate(task.dueDate)}</span>
                </div>
                <div className="caller-amount-block"><span className={`caller-priority ${String(task.priority || "Medium").toLowerCase()}`}>{task.priority || "Medium"}</span><b>{formatMoney(task.amount)}</b></div>
              </div>
              <div className="whatsapp-action">
                <label>WhatsApp template<select defaultValue={whatsappTemplates[0]?.id || ""} disabled={!whatsappTemplates.length} id={`wa-template-${task.id}`}>
                  {whatsappTemplates.map((template) => <option key={template.id} value={template.id}>{template.name}</option>)}
                </select></label>
                <button type="button" disabled={!whatsappTemplates.length} onClick={() => openWhatsApp(task, document.getElementById(`wa-template-${task.id}`)?.value)}><Icon name="phone" />Open WhatsApp</button>
              </div>
              <form className="form-grid compact-form caller-outcome-form" onSubmit={(event) => saveCallerOutcome(event, task)}>
                <label>Outcome<select name="outcome">{callerOutcomes.map((item) => <option key={item.outcome}>{item.outcome}</option>)}</select></label>
                <label>Channel<select name="channel"><option>Call</option><option>WhatsApp</option><option>Manual</option></select></label>
                <label>Expected amount<input name="expectedAmount" placeholder="INR 54,000" /></label>
                <label>Next follow-up<input name="nextFollowUp" placeholder="2026-08-12" /></label>
                <label className="span-2">Notes<textarea name="notes" placeholder="Customer note" /></label>
                <button type="submit"><Icon name="phone" />Save outcome</button>
              </form>
              {history.length > 0 ? <div className="history caller-history">
                <span className="caller-history-label">Recent activity</span>
                {history.slice(0, 2).map((item) => (
                  <p key={item.id}><strong>{item.outcome}</strong><span>{item.notes || "No note added"}</span></p>
                ))}
              </div> : <div className="caller-history-empty"><Icon name="clock" />No previous activity</div>}
            </article>
          );
        })}
      </section>
      <section className="settings-workload whatsapp-log-panel">
        <div className="settings-section-head"><div><h3>WhatsApp message log</h3><p>Track the handoff and delivery status for each customer message.</p></div><strong>{whatsappLogs.length} messages</strong></div>
        {whatsappLogs.length ? <div className="whatsapp-log-list">{whatsappLogs.slice(0, 30).map((log) => {
          const client = getDataClient(data, log.clientId);
          const task = data.dueTasks.find((item) => item.id === log.taskId);
          return <article key={log.id}><div><strong>{client?.name || log.phone}</strong><span>{task?.type || "Task"} | {log.createdAt ? new Date(log.createdAt).toLocaleString("en-IN") : ""}</span><small>{log.body}</small></div><label>Status<select value={log.status} onChange={(event) => updateWhatsAppStatus(log.id, event.target.value)}><option>Prepared</option><option>Opened</option><option>Sent</option><option>Delivered</option><option>Failed</option></select></label></article>;
        })}</div> : <Empty text="No WhatsApp messages logged yet." />}
      </section>
    </section>
  );
}

function Marketplace({ data, updateListingStatus, saveSaleClosing, openFleet }) {
  const [filters, setFilters] = useState({ assetType: "All", condition: "All", finance: "All", insurance: "All", location: "", maxPrice: "" });
  const [showFilters, setShowFilters] = useState(false);
  const [selectedListing, setSelectedListing] = useState(null);
  const filteredListings = data.listings.filter((listing) => {
    const vehicle = getDataVehicle(data, listing.vehicleId);
    const vehicleLiability = vehicle ? liability(vehicle) : 0;
    const locationMatch = !filters.location.trim() || listing.location.toLowerCase().includes(filters.location.trim().toLowerCase());
    const priceMatch = !filters.maxPrice || listing.price <= Number(filters.maxPrice);
    const assetMatch = filters.assetType === "All" || vehicle?.type === filters.assetType;
    const conditionMatch = filters.condition === "All" || listing.condition === filters.condition;
    const financeMatch = filters.finance === "All" || (filters.finance === "Financed" ? vehicleLiability > 0 : vehicleLiability <= 0);
    const insuranceMatch = filters.insurance === "All" || (filters.insurance === "Insured" ? Boolean(vehicle?.insuranceExpiry) : !vehicle?.insuranceExpiry);
    return locationMatch && priceMatch && assetMatch && conditionMatch && financeMatch && insuranceMatch;
  });

  const resetFilters = () => setFilters({ assetType: "All", condition: "All", finance: "All", insurance: "All", location: "", maxPrice: "" });
  const activeFilterCount = Object.entries(filters).filter(([key, value]) => key === "location" || key === "maxPrice" ? Boolean(value) : value !== "All").length;

  return (
    <section className="stack admin-marketplace-page">
      <div className="section-intro admin-marketplace-hero">
        <div>
          <p className="eyebrow">Listings workspace</p>
          <h2>Vehicle marketplace</h2>
          <p>Approve listings and maintain the bank-confirmed closing record for every vehicle sale.</p>
        </div>
        <div className="section-intro-actions">
          <button className="filter-trigger" type="button" onClick={() => setShowFilters(true)}><Icon name="filter" />Filters{activeFilterCount ? ` (${activeFilterCount})` : ""}</button>
          <button type="button" onClick={openFleet}><Icon name="truck" />Open Fleet</button>
        </div>
      </div>
      <>
      {showFilters && <button className="filter-drawer-backdrop" type="button" aria-label="Close filters" onClick={() => setShowFilters(false)} />}
      <aside className={`filter-drawer ${showFilters ? "open" : ""}`} aria-label="Marketplace filters">
        <div className="filter-drawer-head"><div><p className="eyebrow">Refine results</p><h3>Marketplace filters</h3></div><button className="icon-button" type="button" aria-label="Close filters" onClick={() => setShowFilters(false)}><Icon name="close" /></button></div>
        <label>Asset type<select value={filters.assetType} onChange={(event) => setFilters((current) => ({ ...current, assetType: event.target.value }))}><option>All</option><option>Truck</option><option>Trailer</option></select></label>
        <label>Condition<select value={filters.condition} onChange={(event) => setFilters((current) => ({ ...current, condition: event.target.value }))}><option>All</option><option>Excellent</option><option>Good</option><option>Average</option></select></label>
        <label>Finance<select value={filters.finance} onChange={(event) => setFilters((current) => ({ ...current, finance: event.target.value }))}><option>All</option><option>Financed</option><option>Loan-free</option></select></label>
        <label>Insurance<select value={filters.insurance} onChange={(event) => setFilters((current) => ({ ...current, insurance: event.target.value }))}><option>All</option><option>Insured</option><option>Missing</option></select></label>
        <label>Location<input value={filters.location} onChange={(event) => setFilters((current) => ({ ...current, location: event.target.value }))} placeholder="Search location" /></label>
        <label>Max price<input type="number" min="0" value={filters.maxPrice} onChange={(event) => setFilters((current) => ({ ...current, maxPrice: event.target.value }))} placeholder="Any price" /></label>
        <div className="filter-drawer-actions"><button className="filter-reset" type="button" onClick={resetFilters}><Icon name="refresh" />Reset</button><button type="button" onClick={() => setShowFilters(false)}>Apply filters</button></div>
      </aside>
      {data.listings.length === 0 ? <section className="empty-state-panel">
        <Icon name="store" />
        <h3>No marketplace listings yet</h3>
        <p>Create a sale listing from Fleet first. It will appear here for approval and bank closing updates.</p>
        <button type="button" onClick={openFleet}><Icon name="plus" />Create listing from Fleet</button>
      </section> : filteredListings.length > 0 ? <section className="grid-list admin-marketplace-listings">
      {filteredListings.map((listing) => {
        const listingPhoto = listing.photos?.[0];
        const closing = (data.saleClosings ?? []).find((item) => item.listingId === listing.id);
        return (
          <article className="asset admin-marketplace-card" key={listing.id}>
            {listingPhoto?.dataUrl && <img className="listing-photo" src={listingPhoto.dataUrl} alt={listing.title} />}
            <div className="card-head">
              <div>
                <strong>{listing.title}</strong>
                <span>{listing.location} | {listing.condition}</span>
              </div>
              <Badge label={listing.status} />
            </div>
            <Pair label="Price" value={formatMoney(listing.price)} />
            <Pair label="Vehicle" value={getDataVehicle(data, listing.vehicleId)?.regNo ?? listing.vehicleId} />
            <Pair label="Photos" value={String(listing.photos?.length ?? 0)} />
            <Pair label="Estimated closing" value={closing ? formatMoney(closing.estimatedAmount) : formatMoney(getDataVehicle(data, listing.vehicleId) ? liability(getDataVehicle(data, listing.vehicleId)) : 0)} />
            <Pair label="Bank-confirmed closing" value={closing?.bankConfirmedAmount ? formatMoney(closing.bankConfirmedAmount) : "Pending bank confirmation"} />
            <Pair label="Chat threads" value={String((data.marketplaceThreads ?? []).filter((thread) => thread.listingId === listing.id).length)} />
            <div className="actions">
              <button type="button" onClick={() => setSelectedListing(listing)}><Icon name="eye" />View details</button>
              <button onClick={() => updateListingStatus(listing.id, "Active")}>Approve</button>
              <button onClick={() => updateListingStatus(listing.id, "Changes Required")}>Changes</button>
              <button className="danger" onClick={() => updateListingStatus(listing.id, "Rejected")}>Reject</button>
            </div>
            <details className="sale-closing-panel">
              <summary>Bank closing &amp; sale records</summary>
              <form className="form-grid compact-form" onSubmit={(event) => saveSaleClosing(event, listing.id)}>
                <label>Bank-confirmed amount<input name="bankConfirmedAmount" type="number" min="0" step="0.01" defaultValue={closing?.bankConfirmedAmount || ""} placeholder="384884" /></label>
                <label>Sold date<input name="soldDate" type="date" defaultValue={closing?.soldDate || ""} /></label>
                <label>Foreclosure statement<input name="foreclosureStatement" type="file" accept="application/pdf,image/*" /></label>
                <label>Bank NOC<input name="bankNoc" type="file" accept="application/pdf,image/*" /></label>
                <label className="span-2">Ownership transfer proof<input name="ownershipTransfer" type="file" accept="application/pdf,image/*" /></label>
                <button className="span-2" type="submit"><Icon name="check" />Save closing records</button>
              </form>
              {closing && <div className="closing-file-list">
                {[["Foreclosure", closing.foreclosureStatement], ["Bank NOC", closing.bankNoc], ["Transfer proof", closing.ownershipTransfer]].map(([label, file]) => file?.dataUrl && <a key={label} href={file.dataUrl} target="_blank" rel="noreferrer"><Icon name="upload" />Open {label}</a>)}
              </div>}
            </details>
          </article>
        );
      })}
      </section> : <section className="empty-state-panel compact-empty"><Icon name="search" /><h3>No listings match these filters</h3><p>Reset the filters or create another listing from Fleet.</p><button type="button" onClick={resetFilters}><Icon name="refresh" />Reset filters</button></section>}
      </>
      {selectedListing && <MarketplaceVehicleModal listing={selectedListing} vehicle={getDataVehicle(data, selectedListing.vehicleId)} closing={(data.saleClosings ?? []).find((item) => item.listingId === selectedListing.id)} isAdmin onClose={() => setSelectedListing(null)} />}
    </section>
  );
}

function MarketplaceVehicleModal({ listing, vehicle, closing, isAdmin = false, onClose }) {
  const estimatedAmount = closing?.estimatedAmount ?? (vehicle ? liability(vehicle) : 0);
  return (
    <div className="detail-modal-backdrop" role="presentation" onMouseDown={onClose}>
      <section className="detail-modal marketplace-detail-modal" role="dialog" aria-modal="true" aria-labelledby="marketplace-detail-title" onMouseDown={(event) => event.stopPropagation()}>
        <div className="detail-modal-head">
          <div><p className="eyebrow">Marketplace listing</p><h2 id="marketplace-detail-title">{listing.title}</h2><p>{listing.location} | {listing.condition} condition</p></div>
          <div className="detail-modal-actions"><Badge label={listing.status} /><button className="icon-button close-button" type="button" aria-label="Close details" onClick={onClose}><Icon name="close" /></button></div>
        </div>
        {listing.photos?.length > 0 && <div className="marketplace-detail-photos">{listing.photos.map((photo) => photo.dataUrl && <img key={photo.id} src={photo.dataUrl} alt={photo.fileName} />)}</div>}
        <div className="detail-summary-strip">
          <div><span>Asking price</span><strong>{formatMoney(listing.price)}</strong></div>
          <div><span>Estimated closing</span><strong>{formatMoney(estimatedAmount)}</strong></div>
          <div><span>Bank-confirmed closing</span><strong>{closing?.bankConfirmedAmount ? formatMoney(closing.bankConfirmedAmount) : "Pending"}</strong></div>
        </div>
        <div className="detail-card-grid">
          <DetailSection title="Vehicle" rows={[
            ["Registration", vehicle?.regNo],
            ["Asset type", vehicle?.type],
            ["Manufacturer", vehicle?.make],
            ["Model", vehicle?.model],
            ["Year", vehicle?.year],
            ["Kilometres", vehicle?.km ? `${vehicle.km.toLocaleString("en-IN")} km` : "-"],
            ["Asset status", vehicle?.status]
          ]} />
          <DetailSection title="Finance" rows={[
            ["Finance status", vehicle && liability(vehicle) > 0 ? "Financed" : "Loan-free"],
            ...(isAdmin ? [["Financier", vehicle?.financier], ["Loan account", vehicle?.loanAccount]] : []),
            ["Monthly EMI", vehicle?.emiAmount ? formatMoney(vehicle.emiAmount) : "-"],
            ["EMI progress", vehicle?.tenure ? `${vehicle.paidEmi || 0} / ${vehicle.tenure} paid` : "-"],
            ["Public finance view", "Private bank documents hidden"]
          ]} />
          <DetailSection title="Insurance & compliance" rows={[
            ["Insurance company", vehicle?.insuranceCompany],
            ["Insurance expiry", formatDisplayDate(vehicle?.insuranceExpiry)],
            ["Policy number", vehicle?.insurancePolicyNo ? "Available to verified users" : "Not available"],
            ["Permit expiry", formatDisplayDate(vehicle?.permitExpiry)],
            ["PUC expiry", formatDisplayDate(vehicle?.pucExpiry)],
            ["Fitness expiry", formatDisplayDate(vehicle?.fitnessExpiry)]
          ]} />
          <DetailSection title="Sale tracking" rows={[
            ["Listing location", listing.location],
            ["Listing condition", listing.condition],
            ["Closing status", closing?.status || "Estimated"],
            ["Sold date", closing?.soldDate ? formatDisplayDate(closing.soldDate) : "Not sold"],
            ["Photos", `${listing.photos?.length || 0} uploaded`]
          ]} />
        </div>
      </section>
    </div>
  );
}

function Reports({ data, activeReport, setActiveReport }) {
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const rows = getReportRows(activeReport, data);
  const filteredRows = rows.filter((row) => {
    if (!row.date) return !fromDate && !toDate;
    const date = String(row.date).slice(0, 10);
    return (!fromDate || date >= fromDate) && (!toDate || date <= toDate);
  });
  const exportRows = filteredRows.map(({ date, ...row }) => row);
  const exportName = activeReport.toLowerCase().replace(/[^a-z0-9]+/g, "-");

  const downloadExcel = () => {
    const sheet = XLSX.utils.json_to_sheet(exportRows.length ? exportRows : [{ Message: "No records found for the selected date range." }]);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, sheet, activeReport.slice(0, 31));
    XLSX.writeFile(workbook, `kuber-finance-${exportName}-report.xlsx`);
  };

  const printPdf = () => {
    const popup = window.open("", "_blank", "width=1100,height=760");
    if (!popup) return;
    const headers = Object.keys(exportRows[0] ?? { Name: "", Status: "", Amount: "", Detail: "" });
    const body = (exportRows.length ? exportRows : [{ Name: "No records", Status: "", Amount: "", Detail: "No records found for the selected date range." }])
      .map((row) => `<tr>${headers.map((header) => `<td>${escapeHtml(row[header])}</td>`).join("")}</tr>`).join("");
    popup.document.write(`<!doctype html><html><head><title>Kuber Finance - ${escapeHtml(activeReport)} Report</title><style>body{font-family:Arial,sans-serif;padding:28px;color:#152238}h1{margin:0 0 6px}p{color:#5d6b7d;margin:0 0 18px}table{border-collapse:collapse;width:100%;font-size:12px}th,td{border:1px solid #cdd6e0;padding:8px;text-align:left;vertical-align:top}th{background:#e8eef5}</style></head><body><h1>Kuber Finance - ${escapeHtml(activeReport)} Report</h1><p>${fromDate || "All dates"} to ${toDate || "All dates"} | ${filteredRows.length} records</p><table><thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join("")}</tr></thead><tbody>${body}</tbody></table></body></html>`);
    popup.document.close();
    popup.focus();
    popup.onafterprint = () => popup.close();
    popup.print();
  };

  return (
    <section className="two-column">
      <section className="grid-list reports">
        {reportGroups.map(([title, detail]) => (
          <article className={`asset ${activeReport === title ? "selected-card" : ""}`} key={title}>
            <strong>{title}</strong>
            <span>{detail}</span>
            <button onClick={() => setActiveReport(title)}><Icon name="chart" />Open report</button>
          </article>
        ))}
      </section>
      <Panel title={`${activeReport} Report`}>
        <div className="report-toolbar">
          <label>From date<input type="date" value={fromDate} onChange={(event) => setFromDate(event.target.value)} /></label>
          <label>To date<input type="date" value={toDate} onChange={(event) => setToDate(event.target.value)} /></label>
          <button type="button" onClick={downloadExcel}><Icon name="upload" />Excel</button>
          <button type="button" onClick={printPdf}><Icon name="upload" />PDF</button>
        </div>
        <p className="report-result-count">{filteredRows.length} record{filteredRows.length === 1 ? "" : "s"} in selected range</p>
        {activeReport === "Audit" ? (
          <div className="timeline">
            {filteredRows.map((log) => (
              <article key={log.id}>
                <strong>{log.name}</strong>
                <span>{log.status} | {log.amount}</span>
                <small>{log.detail}</small>
              </article>
            ))}
          </div>
        ) : (
          <div className="table-wrap embedded-table">
            <table>
              <thead><tr><th>Name</th><th>Status</th><th>Amount</th><th>Detail</th></tr></thead>
              <tbody>
                {filteredRows.map((row) => <tr key={row.id}><td>{row.name}</td><td><Badge label={row.status} /></td><td>{row.amount}</td><td>{row.detail}</td></tr>)}
              </tbody>
            </table>
          </div>
        )}
      </Panel>
    </section>
  );
}

function ImportRows({ data, importValidRows, lastSavedAt }) {
  const validRows = data.importRows.filter((row) => row.status === "Valid" && row.issue !== "Imported and saved");
  return (
    <section className="stack">
      <Panel title="Excel Import">
        <div className="import-callout">
          <div>
            <strong>{validRows.length} valid rows ready</strong>
            <span>Creates client, vehicle, loan and schedule records in the workspace.</span>
          </div>
          <button disabled={validRows.length === 0} onClick={importValidRows}><Icon name="upload" />Final import</button>
        </div>
        <small>Last saved: {lastSavedAt ?? "Not saved in this session"}</small>
      </Panel>
      <section className="table-wrap">
        <table>
          <thead><tr><th>Row</th><th>Reg No</th><th>Asset</th><th>Client</th><th>Loan</th><th>Status</th><th>Issue</th></tr></thead>
          <tbody>
            {data.importRows.map((row) => (
              <tr key={row.row}>
                <td>{row.row}</td>
                <td>{row.regNo}</td>
                <td>{row.assetType}</td>
                <td>{row.client}</td>
                <td>{row.loanAccount}</td>
                <td><Badge label={row.status} /></td>
                <td>{row.issue}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </section>
    </section>
  );
}

function Settings({ data, lastSavedAt, saveStatus, clearNotifications, resetDemoData, exportData, saveCommonPassword, saveReminderSettings, saveRolePermissions, createCallerAccount, deleteCallerAccount, runCallerAssignment, saveWhatsAppTemplates, setSection }) {
  const callers = (data.users ?? users).filter((user) => user.role === "Caller");
  const openDues = data.dueTasks.filter((task) => task.status !== "Closed");
  const unassignedDues = openDues.filter((task) => !task.callerId);
  const assignedDues = openDues.length - unassignedDues.length;
  const editablePermissionRows = Array.isArray(data.rolePermissions) ? data.rolePermissions : permissionRows;
  const accessRows = editablePermissionRows.map(([feature, admin, owner, caller, buyer]) => ({
    feature,
    Admin: admin,
    Owner: owner,
    Caller: caller,
    Customer: buyer
  }));
  const permissionOptions = ["Yes", "No", "Assigned only", "Own fleet", "Optional", "Reports only"];
  const roleActions = [
    { role: "Admin", section: "clients", icon: "people", label: "Manage clients", allowed: accessRows.filter((row) => row.Admin !== "No").length },
    { role: "Caller", section: "caller", icon: "phone", label: "Open queue", allowed: accessRows.filter((row) => row.Caller !== "No").length },
    { role: "Customer", section: "marketplace", icon: "store", label: "View listings", allowed: accessRows.filter((row) => row.Customer !== "No").length }
  ];

  return (
    <section className="settings-page">
      <section className="settings-hero">
        <div>
          <span>System control</span>
          <h2>Operations settings</h2>
          <p>Customer login, caller workload and access rules are controlled from here.</p>
        </div>
        <div className="settings-hero-actions">
          <button type="button" onClick={exportData}><Icon name="upload" />Export</button>
          <button type="button" onClick={clearNotifications}><Icon name="bell" />Clear alerts</button>
        </div>
      </section>

      <section className="settings-status-grid">
        <article><Icon name="cloud" /><span>Storage</span><strong>{saveStatus}</strong></article>
        <article><Icon name="check" /><span>Last saved</span><strong>{lastSavedAt ?? "Ready"}</strong></article>
        <article><Icon name="phone" /><span>Assigned dues</span><strong>{assignedDues}/{openDues.length}</strong></article>
        <article><Icon name="bell" /><span>Unassigned</span><strong>{unassignedDues.length}</strong></article>
      </section>

      <section className="settings-grid">
        <article className="settings-card">
          <div className="settings-card-head"><Icon name="people" /><div><h3>Customer login password</h3><p>Shared password for all customer accounts until they change it.</p></div></div>
          <form className="settings-form" onSubmit={saveCommonPassword}>
            <label>Common password<input name="commonPassword" type="password" placeholder="Set shared password" minLength={4} required autoComplete="new-password" /></label>
            <button type="submit"><Icon name="check" />Save password</button>
          </form>
        </article>

        <article className="settings-card">
          <div className="settings-card-head"><Icon name="phone" /><div><h3>Caller assignment</h3><p>Create callers and assign unassigned dues with one rule.</p></div></div>
          <form className="settings-form two" onSubmit={createCallerAccount} autoComplete="off">
            <label>Caller name<input name="newCallerName" placeholder="Caller name" autoComplete="off" /></label>
            <label>Caller email<input name="newCallerEmail" type="email" placeholder="caller@kuber.local" autoComplete="off" /></label>
            <label>Caller password<input name="newCallerPassword" type="password" placeholder="Set password" minLength={4} autoComplete="new-password" /></label>
            <button type="submit"><Icon name="phone" />Create caller</button>
          </form>
          <form className="settings-form assignment-form" onSubmit={runCallerAssignment}>
            <label>Assignment rule<select name="assignmentMode" defaultValue="permanent-client"><option value="permanent-client">Permanent client-wise</option><option value="round-robin">Round-robin</option><option value="location-wise">Location-wise</option><option value="category-wise">Category-wise</option></select></label>
            <button type="submit"><Icon name="check" />Assign now</button>
          </form>
        </article>

        <article className="settings-card">
          <div className="settings-card-head"><Icon name="calendar" /><div><h3>Reminder automation</h3><p>Control when EMI, insurance and compliance reminders are created.</p></div></div>
          <form className="settings-form" onSubmit={saveReminderSettings} key={JSON.stringify(data.reminderSettings ?? {})}>
            <label className="settings-checkbox"><input name="reminderEnabled" type="checkbox" defaultChecked={data.reminderSettings?.enabled !== false} /><span>Enable automated monitoring</span></label>
            <div className="settings-form two reminder-fields">
              <label>Scan frequency (hours)<input name="reminderIntervalHours" type="number" min="1" max="168" step="1" defaultValue={data.reminderSettings?.intervalHours ?? 24} required /></label>
              <label>Reminder window (days)<input name="reminderWindowDays" type="number" min="1" max="180" step="1" defaultValue={data.reminderSettings?.windowDays ?? 30} required /></label>
            </div>
            <button type="submit"><Icon name="check" />Save reminder settings</button>
          </form>
        </article>
        <WhatsAppTemplatesSettings templates={data.whatsappTemplates ?? initialWhatsAppTemplates} saveTemplates={saveWhatsAppTemplates} />
      </section>

      <section className="settings-workload">
        <div className="settings-section-head">
          <div><h3>Caller workload</h3><p>Live queue split from open due records.</p></div>
          <button type="button" onClick={() => setSection("caller")}><Icon name="phone" />Open queue</button>
        </div>
        <div className="caller-list">
          {callers.map((caller) => {
            const count = openDues.filter((task) => task.callerId === caller.id).length;
            return <article className="caller-workload-row" key={caller.id}><div className="caller-workload-identity"><span className="caller-avatar"><Icon name="phone" /></span><div><strong>{caller.name}</strong><span>{caller.email}</span></div></div><div className="caller-workload-meta"><b>{count} active</b><button className="caller-delete-button" type="button" title={`Delete ${caller.name}`} aria-label={`Delete ${caller.name}`} onClick={() => deleteCallerAccount(caller)}><Icon name="trash" />Delete</button></div></article>;
          })}
          {callers.length === 0 && <Empty text="No callers yet. Create a caller above, then run assignment." />}
        </div>
      </section>

      <form className="settings-access" onSubmit={saveRolePermissions}>
        <div className="settings-section-head">
          <div><h3>Role access</h3><p>Set which role can use each feature, then save permissions.</p></div>
          <div className="settings-section-actions">
            <button type="submit"><Icon name="check" />Save permissions</button>
            <button className="danger" type="button" onClick={resetDemoData}><Icon name="settings" />Reset data</button>
          </div>
        </div>
        <div className="role-card-grid">
          {roleActions.map((role) => (
            <article className="role-card" key={role.role}>
              <div className="role-card-top"><Icon name={role.icon} /><strong>{role.role}</strong><span>{role.allowed} allowed</span></div>
              <div className="permission-control-list">
                {accessRows.map((row) => {
                  const value = row[role.role] ?? "No";
                  const allowed = value !== "No";
                  return (
                    <label className={allowed ? "permission-control on" : "permission-control"} key={`${role.role}-${row.feature}`}>
                      <span>{row.feature}</span>
                      <select name={`permission-${role.role}-${row.feature}`} defaultValue={value}>
                        {permissionOptions.map((option) => <option value={option} key={option}>{option}</option>)}
                      </select>
                    </label>
                  );
                })}
              </div>
              <button type="button" onClick={() => setSection(role.section)}><Icon name={role.icon} />{role.label}</button>
            </article>
          ))}
        </div>
      </form>
    </section>
  );
}

// ─── Customer portal (post-login, shown when logging in as Customer) ────────
function WhatsAppTemplatesSettings({ templates, saveTemplates }) {
  const [drafts, setDrafts] = useState(templates);

  useEffect(() => setDrafts(templates), [templates]);

  const updateDraft = (id, field, value) => {
    setDrafts((current) => current.map((item) => item.id === id ? { ...item, [field]: value } : item));
  };

  const addDraft = () => {
    setDrafts((current) => [...current, { id: `custom-${Date.now()}`, name: "New template", body: "Hello {{customer}}, please contact Kuber Finance regarding {{vehicle}}.", active: true }]);
  };

  return (
    <article className="settings-card whatsapp-template-card">
      <div className="settings-card-head"><Icon name="phone" /><div><h3>WhatsApp templates</h3><p>Use customer, vehicle, type, dueDate and amount placeholders in reusable messages.</p></div></div>
      <div className="whatsapp-template-list">
        {drafts.map((template) => <div className="whatsapp-template-row" key={template.id}>
          <label>Name<input value={template.name} onChange={(event) => updateDraft(template.id, "name", event.target.value)} /></label>
          <label className="template-body">Message<textarea value={template.body} onChange={(event) => updateDraft(template.id, "body", event.target.value)} /></label>
          <label className="settings-checkbox template-active"><input type="checkbox" checked={template.active} onChange={(event) => updateDraft(template.id, "active", event.target.checked)} /><span>Active</span></label>
          <button className="danger" type="button" onClick={() => setDrafts((current) => current.filter((item) => item.id !== template.id))}><Icon name="trash" />Remove</button>
        </div>)}
      </div>
      <div className="settings-section-actions"><button type="button" onClick={addDraft}><Icon name="plus" />Add template</button><button type="button" onClick={() => saveTemplates(drafts)}><Icon name="check" />Save templates</button></div>
    </article>
  );
}

const customerNavItems = [
  ["dashboard", "Dashboard", "home"],
  ["fleet",     "My Fleet",  "truck"],
  ["dues",      "Dues",      "calendar"],
  ["documents", "Documents", "upload"],
  ["marketplace","Marketplace","store"],
  ["chats", "Chats", "bell"],
];

function CustomerPortal({ session, onLogout }) {
  const marketplaceChat = useMarketplaceChat(API_BASE, session.token);
  const [section, setSection] = useState("dashboard");
  const [data, setData] = useState(loadData);
  const [loadError, setLoadError] = useState("");
  const visibleNavItems = useMemo(
    () => customerNavItems.filter(([key]) => canOpenSection(data, session.role, key)),
    [data, session.role]
  );
  const openPortalSection = (key) => {
    if (canOpenSection(data, session.role, key)) {
      setSection(key);
      return;
    }
    setLoadError("Permission denied for this section.");
  };

  useEffect(() => {
    let active = true;
    const refreshFromDatabase = () => fetchBackendData(session.token)
      .then((backendData) => {
        if (!active) return;
        setData(backendData);
        setLoadError("");
      })
      .catch((error) => {
        if (!active) return;
        if (error?.status === 401) {
          clearSession();
          onLogout();
          return;
        }
        setLoadError(error.message || "Customer data could not be loaded from the server.");
      });
    refreshFromDatabase();
    const intervalId = window.setInterval(refreshFromDatabase, 10000);
    const handleFocus = () => refreshFromDatabase();
    window.addEventListener("focus", handleFocus);
    return () => {
      active = false;
      window.clearInterval(intervalId);
      window.removeEventListener("focus", handleFocus);
    };
  }, []);

  const myClient = data.clients.find((c) => c.id === session.clientId)
    ?? data.clients.find((c) => session.email && c.email?.trim().toLowerCase() === session.email.trim().toLowerCase())
    ?? data.clients.find((c) => c.name?.toLowerCase() === session.name?.toLowerCase())
    ?? clients.find((c) => c.id === session.clientId);
  const customerClientId = myClient?.id ?? session.clientId;

  const myVehicles = useMemo(
    () => data.vehicles.filter((v) => v.clientId === customerClientId),
    [data, customerClientId]
  );
  const activeVehicles = useMemo(
    () => myVehicles.filter((vehicle) => vehicle.status !== "Sold"),
    [myVehicles]
  );
  const soldVehicles = useMemo(
    () => myVehicles.filter((vehicle) => vehicle.status === "Sold"),
    [myVehicles]
  );
  const myDues = useMemo(
    () => {
      const soldVehicleIds = new Set(myVehicles.filter((vehicle) => vehicle.status === "Sold").map((vehicle) => vehicle.id));
      return data.dueTasks.filter((task) => task.clientId === customerClientId && !soldVehicleIds.has(task.vehicleId));
    },
    [data, customerClientId, myVehicles]
  );
  const myListings = useMemo(
    () => data.listings.filter((l) => myVehicles.some((v) => v.id === l.vehicleId)),
    [data, myVehicles]
  );
  const marketplaceListings = useMemo(
    () => data.listings.filter((l) => {
      if (!["Active", "Reserved"].includes(l.status)) return false;
      const vehicle = myVehicles.find((item) => item.id === l.vehicleId);
      const closing = (data.saleClosings ?? []).find((item) => item.listingId === l.id);
      return vehicle?.status !== "Sold" && closing?.status !== "Sold";
    }),
    [data, myVehicles]
  );
  const myImportedAssets = useMemo(
    () => (data.clientImports ?? [])
      .filter((item) => item.clientId === customerClientId)
      .flatMap((item) => (item.rows ?? []).map((row) => ({ ...row, importFile: item.fileName, importedAt: item.importedAt }))),
    [data, customerClientId]
  );

  const totalLiability = useMemo(
    () => activeVehicles.reduce((sum, v) => sum + customerVehicleClosingPrincipal(v, myImportedAssets), 0),
    [activeVehicles, myImportedAssets]
  );

  const openDues = myDues.filter((t) => t.status !== "Closed");
  const sectionAllowed = section === "dashboard" || canOpenSection(data, session.role, section);
  const myDocuments = useMemo(
    () => (data.documents ?? []).filter((document) => myVehicles.some((vehicle) => vehicle.id === document.vehicleId)),
    [data, myVehicles]
  );

  const persistCustomerData = async (nextData, message) => {
    setData(nextData);
    try {
      await syncDataToBackend(nextData, session.token);
    } catch (error) {
      setLoadError(error.message || "Customer data could not be saved.");
      return;
    }
    setLoadError(message || "");
  };



  const submitCustomerListing = async (event) => {
    event.preventDefault();
    if (!canUsePermission(data, session.role, "Create sale listing")) {
      setLoadError("Permission denied for sale listings.");
      return;
    }
    const form = event.currentTarget;
    const formData = new FormData(form);
    const vehicleId = String(formData.get("vehicleId") || "");
    const vehicle = myVehicles.find((item) => item.id === vehicleId);
    const price = Number(formData.get("price") || 0);
    const location = String(formData.get("location") || "").trim();
    const condition = String(formData.get("condition") || "Good");
    if (!vehicle || !price || !location) {
      setLoadError("Choose a vehicle and enter a valid price and location.");
      return;
    }
    const photos = await readListingPhotos(formData.getAll("photos"));
    const existing = data.listings.find((item) => item.vehicleId === vehicleId && !["Rejected", "Disabled"].includes(item.status));
    const listing = {
      id: existing?.id ?? `m-${Date.now()}`,
      vehicleId,
      title: `${vehicle.year || ""} ${vehicle.make || vehicle.type} ${vehicle.model || "Vehicle"}`.trim(),
      price,
      location,
      status: existing?.status ?? "Submitted",
      condition,
      photos
    };
    const updated = {
      ...data,
      listings: [listing, ...data.listings.filter((item) => item.id !== listing.id)],
      notifications: [
        { id: `n-${Date.now()}`, title: "Sale listing submitted", detail: `${listing.title} is waiting for Admin approval.`, target: "marketplace", unread: true },
        ...(data.notifications ?? [])
      ]
    };
    await persistCustomerData(updated, "Vehicle listing submitted for Admin approval.");
    form.reset();
  };

  const submitCustomerProof = async (task, file, details = {}) => {
    if (!canUsePermission(data, session.role, "View own fleet")) {
      setLoadError("Permission denied for proof upload.");
      return;
    }
    if (!file) return;
    if (file.size > 8 * 1024 * 1024) {
      setLoadError("Please upload a PDF or image smaller than 8 MB.");
      return;
    }
    const vehicle = data.vehicles.find((item) => item.id === task.vehicleId);
    const uploadedAt = new Date().toLocaleString("en-IN");
    const dataUrl = await readFileAsDataUrl(file);
    const document = {
      id: `doc-${Date.now()}`,
      clientId: task.clientId,
      vehicleId: task.vehicleId,
      taskId: task.id,
      type: task.type === "Insurance" ? "Insurance Policy" : task.type === "EMI" ? "Payment Receipt" : "Compliance Document",
      fileName: file.name,
      mimeType: file.type || "application/octet-stream",
      size: file.size,
      dataUrl,
      uploadedBy: myClient?.name ?? session.name,
      uploadedAt,
      note: details.customerNote || "Submitted from customer web portal"
    };
    const detailRows = Object.entries(details)
      .filter(([key, value]) => key !== "customerNote" && value)
      .map(([key, value]) => [key.replace(/([A-Z])/g, " $1").replace(/^./, (letter) => letter.toUpperCase()), String(value)]);
    const verificationItem = {
      id: `vf-${Date.now()}`,
      taskId: task.id,
      submittedBy: myClient?.name ?? session.name,
      submittedAt: uploadedAt,
      proofType: document.type,
      details: [
        ["Vehicle", vehicle?.regNo ?? "-"],
        ["Due Type", task.type],
        ["Due Amount", formatMoney(task.amount)],
        ...detailRows,
        ["Uploaded File", file.name]
      ],
      audit: [
        ["Previous Status", task.status],
        ["Current Status", "Proof Pending"],
        ["Caller Contact", "Paused until Admin decision"],
        ["Owner Note", "Submitted from customer web portal"]
      ]
    };
    const updated = {
      ...data,
      dueTasks: data.dueTasks.map((item) => item.id === task.id ? { ...item, status: "Proof Pending" } : item),
      verificationItems: [verificationItem, ...(data.verificationItems ?? []).filter((item) => item.taskId !== task.id)],
      documents: [document, ...(data.documents ?? [])],
      notifications: [
        { id: `n-${Date.now()}`, title: `${document.type} submitted`, detail: `${document.uploadedBy} uploaded ${file.name}`, target: "verification", unread: true },
        ...(data.notifications ?? [])
      ],
      auditLogs: [
        {
          id: `a-${Date.now()}`,
          module: "Owner Submission",
          action: document.type,
          record: task.id,
          oldValue: task.status,
          newValue: "Proof Pending",
          remark: file.name,
          at: uploadedAt
        },
        ...(data.auditLogs ?? [])
      ]
    };
    await persistCustomerData(updated, "Proof uploaded and sent for Admin verification.");
  };

  return (
    <main className="app-shell customer-shell">
      {loadError && <div className="toast error" role="alert">{loadError}</div>}
      <aside className="sidebar customer-sidebar">
        <div className="brand">
          <span className="brand-mark">
            <KuberBrandMark />
          </span>
          <div>
            <strong>My Fleet</strong>
            <small>Customer portal</small>
          </div>
        </div>
        <nav className="sidebar-nav" aria-label="Customer sections">
          {visibleNavItems.map((item) => (
            <button
              key={item[0]}
              className={section === item[0] ? "active" : ""}
              onClick={() => openPortalSection(item[0])}
            >
              {item[0] === "documents" ? (
                <span className="customer-nav-icon"><ProofFileIcon /></span>
              ) : (
                <Icon name={item[2]} />
              )}
              {item[1]}
            </button>
          ))}
        </nav>
        <div className="sidebar-footer">
          <div className="sidebar-footer-inner">
            <div className="sidebar-footer-info">
              <strong>{myClient?.name ?? session.name}</strong>
              <span>{myClient?.city ?? session.role}</span>
            </div>
            <span className="sidebar-footer-avatar" aria-hidden="true">
              {(myClient?.name ?? session.name).slice(0, 1)}
            </span>
          </div>
          <button className="logout-button" type="button" onClick={onLogout} title="Sign out" aria-label="Sign out">
            <Icon name="logout" />
            <span>Sign out</span>
          </button>
        </div>
      </aside>

      <section className="workspace">
        <header className="topbar">
          <div>
            <p>Customer portal</p>
            <h1>{customerNavItems.find(([k]) => k === section)?.[1] ?? "Dashboard"}</h1>
          </div>
          <div className="customer-profile-pill">
            <span className="customer-avatar-badge">{(myClient?.name ?? session.name).slice(0, 1)}</span>
            <div>
              <strong>{myClient?.name ?? session.name}</strong>
              <small>{myClient?.city ?? ""}</small>
            </div>
            <span className="customer-profile-caret" aria-hidden="true"><Icon name="chevron-down" /></span>
          </div>
        </header>

        {!sectionAllowed && <AccessDenied role={roleAccessName(session.role)} section={section} />}
        {sectionAllowed && section === "dashboard" && (
          <CustomerDashboard
            client={myClient}
            vehicles={activeVehicles}
            dues={myDues}
            openDues={openDues}
            totalLiability={totalLiability}
            importedAssets={myImportedAssets}
            setSection={openPortalSection}
          />
        )}
        {sectionAllowed && section === "fleet" && (
          <CustomerFleet vehicles={activeVehicles} soldVehicles={soldVehicles} saleClosings={data.saleClosings} client={myClient} importedAssets={myImportedAssets} />
        )}
        {sectionAllowed && section === "dues" && (
          <CustomerDues dues={myDues} vehicles={myVehicles} submitCustomerProof={submitCustomerProof} />
        )}
       {sectionAllowed && section === "documents" && (
          <CustomerDocuments documents={myDocuments} vehicles={myVehicles} setSection={openPortalSection} />
       )}
        {sectionAllowed && section === "marketplace" && (
          <CustomerMarketplace chat={marketplaceChat} listings={marketplaceListings} vehicles={data.vehicles} saleClosings={data.saleClosings} myVehicles={myVehicles} myVehicleIds={new Set(myVehicles.map((vehicle) => vehicle.id))} submitListing={submitCustomerListing} setSection={openPortalSection} />
        )}
        {sectionAllowed && section === "chats" && (
          <MarketplaceChat chat={marketplaceChat} clientId={customerClientId} />
        )}
      </section>
    </main>
  );
}

function CustomerDashboard({ client, vehicles, dues, openDues, totalLiability, importedAssets = [], setSection }) {
  const overdue = dues.filter((t) => ["Overdue", "Escalated"].includes(t.status));

  return (
    <section className="dashboard-page">
      <div className="customer-hero">
        <div className="customer-hero-glow" aria-hidden="true" />
        <div className="customer-hero-avatar">{(client?.name ?? "K").slice(0, 1)}</div>
        <div className="customer-hero-text">
          <span className="customer-hero-eyebrow">Welcome back 👋</span>
          <h2>{client?.name ?? "Customer"}</h2>
          <p>{client?.city ? `${client.city} · ` : ""}Here's the current state of your fleet and dues.</p>
        </div>
        <ul className="customer-hero-stats">
          <li><span>Vehicles</span><strong>{vehicles.length}</strong><Icon name="truck" /></li>
          <li><span>Open dues</span><strong>{openDues.length}</strong><Icon name="calendar" /></li>
          <li><span>Closing total</span><strong>{formatMoney(totalLiability)}</strong><Icon name="money" /></li>
        </ul>
      </div>

      <section className="metrics dashboard-metrics customer-metrics">
        <CustomerMetric label="Vehicles" value={vehicles.length} helper="Total vehicles" icon="truck" onClick={() => setSection("fleet")} />
        <CustomerMetric label="Open Dues" value={openDues.length} helper="Outstanding dues" icon="calendar" onClick={() => setSection("dues")} />
        <CustomerMetric label="Closing Total" value={formatMoney(totalLiability)} helper="Total outstanding amount" icon="money" accent />
      </section>

      {overdue.length > 0 && (
        <div className="customer-alert-banner" role="alert">
          <Icon name="bell" />
          <span>
            <strong>{overdue.length} overdue/escalated due{overdue.length > 1 ? "s" : ""}</strong> — please contact your relationship manager immediately.
          </span>
        </div>
      )}

      <section className="split dashboard-split">
        <section className="panel customer-summary-panel">
          <div className="customer-panel-head">
            <div><span className="customer-panel-icon"><Icon name="truck" /></span><h2>My Fleet Summary</h2></div>
            <button type="button" aria-label="View fleet" onClick={() => setSection("fleet")}><Icon name="chevron-right" /></button>
          </div>
          {vehicles.length > 0 && (
            <div className="customer-table-head customer-fleet-table-head">
              <span>Vehicle</span><span>Model</span><span>Status</span><span>Value</span>
            </div>
          )}
          {vehicles.slice(0, 5).map((v) => (
            <article className="customer-vehicle-row" key={v.id}>
              <div className="customer-vehicle-reg"><span><Icon name="truck" /></span>{v.regNo}</div>
              <div className="customer-vehicle-model">{v.year} {v.make} {v.model}</div>
              <div className="customer-vehicle-status">
                <Badge label={v.status} />
              </div>
              <div className="customer-vehicle-amount">{formatMoney(customerVehicleClosingPrincipal(v, importedAssets))}</div>
            </article>
          ))}
          {vehicles.length === 0 && <Empty text="No vehicles found for your account." />}
          {vehicles.length > 0 && <p className="customer-panel-foot">Showing {Math.min(vehicles.length, 5)} of {vehicles.length} vehicles</p>}
        </section>
        <section className="panel customer-summary-panel">
          <div className="customer-panel-head">
            <div><span className="customer-panel-icon"><Icon name="calendar" /></span><h2>Recent Dues</h2></div>
            <button type="button" aria-label="View dues" onClick={() => setSection("dues")}><Icon name="chevron-right" /></button>
          </div>
          {openDues.length > 0 && (
            <div className="customer-table-head customer-dues-table-head">
              <span>Due type</span><span>Details</span><span>Status</span><span>Amount</span>
            </div>
          )}
          {openDues.slice(0, 5).map((task) => {
            const v = vehicles.find((x) => x.id === task.vehicleId);
            return (
              <article className="customer-due-row" key={task.id}>
                <div>
                  <strong>{task.type}</strong>
                  <span>{v?.regNo ?? "—"} · Due {formatDisplayDate(task.dueDate)}</span>
                </div>
                <div className="customer-due-right">
                  <Badge label={task.status} />
                  <b>{formatMoney(task.amount)}</b>
                </div>
              </article>
            );
          })}
          {openDues.length === 0 && <Empty text="No open dues. All clear!" />}
          {openDues.length > 0 && <p className="customer-panel-foot">Showing {Math.min(openDues.length, 5)} of {openDues.length} dues</p>}
        </section>
      </section>
    </section>
  );
}

function CustomerFleet({ vehicles, soldVehicles = [], saleClosings = [], client, importedAssets = [] }) {
  const bodyByReg = useMemo(() => {
    const pairs = new Map();
    importedAssets
      .filter((asset) => isBodyRow(asset))
      .forEach((asset) => {
        const key = normalizeRegNo(baseRegNo(asset.regNo));
        if (key && !pairs.has(key)) pairs.set(key, asset);
      });
    return pairs;
  }, [importedAssets]);
  return (
    <section className="customer-fleet-page">
      <div className="customer-fleet-summary">
        <div>
          <span>Fleet linked to</span>
          <strong>{client?.name ?? "Customer"}</strong>
        </div>
        <div>
          <span>Total vehicles</span>
          <strong>{vehicles.length}</strong>
        </div>
        <div>
          <span>Sold vehicles</span>
          <strong>{soldVehicles.length}</strong>
        </div>
        <div>
          <span>Body records</span>
          <strong>{[...bodyByReg.keys()].length}</strong>
        </div>
      </div>
      <div className="customer-fleet-grid">
        {vehicles.map((v) => (
          <article className="customer-fleet-card" key={v.id}>
            <div className="customer-fleet-card-head">
              <div>
                <strong>{v.regNo}</strong>
                <span>{v.year} {v.make} {v.model}</span>
              </div>
              <Badge label={v.status} />
            </div>
            <dl className="customer-fleet-details">
              <div><dt>Type</dt><dd>{v.type}</dd></div>
              <div><dt>KM</dt><dd>{v.km.toLocaleString("en-IN")}</dd></div>
              <div><dt>Finance Closing Principal</dt><dd>{formatMoney(customerVehicleClosingPrincipal(v, importedAssets))}</dd></div>
              <div><dt>Insurance</dt><dd>{formatDisplayDate(v.insuranceExpiry)}</dd></div>
              <div><dt>Permit</dt><dd>{formatDisplayDate(v.permitExpiry)}</dd></div>
              <div><dt>Combination</dt><dd>{v.combinationId || "Unlinked"}</dd></div>
            </dl>
            <div className="customer-loan-summary">
              <div><span>Loan ID</span><strong>{v.loanId || `LOAN-${String(v.regNo || v.id).replace(/[^A-Z0-9]/gi, "").toUpperCase()}`}</strong></div>
              <div><span>Loan account</span><strong>{v.loanAccount || "-"}</strong></div>
              <div><span>EMI</span><strong>{v.emiAmount ? formatMoney(v.emiAmount) : "-"}</strong></div>
              <div><span>Progress</span><strong>{v.tenure ? `${customerVehiclePaidEmi(v, importedAssets)} / ${v.tenure} paid` : "-"}</strong></div>
            </div>
            <div className="customer-loan-history">
              <span>Payment history</span>
              <strong>{v.emiHistory?.length || 0} verified payments · {v.emiSchedule?.length || 0} schedule rows</strong>
            </div>
            {bodyByReg.get(normalizeRegNo(baseRegNo(v.regNo))) ? (
              <BodyFinanceBlock asset={bodyByReg.get(normalizeRegNo(baseRegNo(v.regNo)))} />
            ) : (
              <div className="customer-body-empty">No body finance record linked</div>
            )}
          </article>
        ))}
        {vehicles.length === 0 && soldVehicles.length === 0 && (
          <Empty text="No vehicles found for your account." />
        )}
      </div>
      {soldVehicles.length > 0 && (
        <section className="customer-sold-vehicles">
          <div className="customer-fleet-section-head">
            <div><span>Vehicle archive</span><h2>Sold vehicles</h2></div>
            <strong>{soldVehicles.length} archived</strong>
          </div>
          <div className="customer-fleet-grid">
            {soldVehicles.map((vehicle) => {
              const closing = saleClosings.find((item) => item.vehicleId === vehicle.id);
              return (
                <article className="customer-fleet-card customer-sold-card" key={`sold-${vehicle.id}`}>
                  <div className="customer-fleet-card-head">
                    <div>
                      <strong>{vehicle.regNo}</strong>
                      <span>{vehicle.year} {vehicle.make} {vehicle.model}</span>
                    </div>
                    <Badge label="Sold" />
                  </div>
                  <dl className="customer-fleet-details">
                    <div><dt>Finance Closing Principal</dt><dd>{formatMoney(customerVehicleClosingPrincipal(vehicle, importedAssets))}</dd></div>
                    <div><dt>Sold date</dt><dd>{formatDisplayDate(vehicle.soldDate || closing?.soldDate)}</dd></div>
                    <div><dt>Monthly EMI</dt><dd>{vehicle.emiAmount ? formatMoney(vehicle.emiAmount) : "-"}</dd></div>
                    <div><dt>Loan amount</dt><dd>{vehicle.loanAmount ? formatMoney(vehicle.loanAmount) : "-"}</dd></div>
                    <div><dt>Loan account</dt><dd>{vehicle.loanAccount || "-"}</dd></div>
                  </dl>
                </article>
              );
            })}
          </div>
        </section>
      )}
    </section>
  );
}

function BodyFinanceBlock({ asset }) {
  return (
    <div className="customer-body-block">
      <div className="customer-body-block-head">
        <span>Body finance</span>
        <strong>{formatAutoClosingPrincipal(asset)}</strong>
      </div>
      <div className="customer-body-grid">
        <div><span>Loan No.</span><strong>{asset.loanAccount || "-"}</strong></div>
        <div><span>Financier</span><strong>{asset.financier || "-"}</strong></div>
        <div><span>EMI</span><strong>{formatPlainMoney(asset.emiAmount)}</strong></div>
        <div><span>Paid EMI</span><strong>{effectivePaidEmi(asset) || "-"}</strong></div>
      </div>
    </div>
  );
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => reject(reader.error || new Error("File could not be read."));
    reader.readAsDataURL(file);
  });
}

async function readListingPhotos(files) {
  const validFiles = Array.from(files || [])
    .filter((file) => file && typeof file.type === "string" && file.type.startsWith("image/") && file.size <= 5 * 1024 * 1024)
    .slice(0, 6);

  return Promise.all(validFiles.map(async (file, index) => ({
    id: `lp-${Date.now()}-${index}`,
    fileName: file.name || "vehicle-photo.jpg",
    mimeType: file.type || "image/jpeg",
    size: file.size || 0,
    dataUrl: await readFileAsDataUrl(file),
    uploadedAt: new Date().toLocaleString("en-IN")
  })));
}

function CustomerDues({ dues, vehicles, submitCustomerProof }) {
  const open = dues.filter((t) => t.status !== "Closed");
  const closed = dues.filter((t) => t.status === "Closed");
  const totalOpen = open.reduce((s, t) => s + t.amount, 0);

  return (
    <section className="dues-board">
      <div className="dues-summary">
        <article className="dues-summary-card summary-alerts">
          <div className="dues-summary-icon"><AlertSvg name="calendar" /></div>
          <div><span>Open dues</span><strong>{open.length}</strong><small>Needs attention</small></div>
        </article>
        <article className="dues-summary-card summary-expired">
          <div className="dues-summary-icon"><AlertSvg name="check" /></div>
          <div><span>Closed</span><strong>{closed.length}</strong><small>Completed records</small></div>
        </article>
        <article className="dues-summary-card summary-payment">
          <div className="dues-summary-icon"><AlertSvg name="money" /></div>
          <div><span>Total outstanding</span><strong>{formatMoney(totalOpen)}</strong><small>Open due amount</small></div>
        </article>
      </div>
      <section className="smart-alert-panel">
        <div className="smart-alert-head">
          <h2>All Dues</h2>
          <span>{dues.length} records</span>
        </div>
        <div className="dues-list" style={{ padding: "14px" }}>
          {open.map((task) => {
            const v = vehicles.find((x) => x.id === task.vehicleId);
            return (
              <article className="due" key={task.id}>
                <div className="due-copy">
                  <strong>{task.type} — {v?.regNo ?? "—"}</strong>
                  <span>Due {formatDisplayDate(task.dueDate)} · {task.priority} priority</span>
                </div>
                <div className="due-status"><Badge label={task.status} /></div>
                <b className="due-amount">{formatMoney(task.amount)}</b>
                {task.status !== "Proof Pending" && <CustomerProofForm task={task} submitCustomerProof={submitCustomerProof} />}
                {task.status === "Proof Pending" && <span className="proof-pending-note">Proof submitted for Admin review</span>}
              </article>
            );
          })}
          {dues.length === 0 && <Empty text="No dues found for your account." />}
        </div>
      </section>
    </section>
  );
}

function CustomerProofForm({ task, submitCustomerProof }) {
  const isEmi = task.type === "EMI";
  const isInsurance = task.type === "Insurance";
  return (
    <details className="customer-proof-details">
      <summary><span className="proof-summary-icon"><ProofFileIcon /></span> Submit details &amp; proof <span className="proof-summary-toggle" aria-hidden="true"><Icon name="plus" /></span></summary>
      <form className="form-grid compact-form" onSubmit={(event) => {
        event.preventDefault();
        const formData = new FormData(event.currentTarget);
        const details = Object.fromEntries(formData.entries());
        const file = formData.get("proofFile");
        submitCustomerProof(task, file, details);
        event.currentTarget.reset();
      }}>
        {isEmi && <>
          <label>Payment date<input name="paymentDate" type="date" required /></label>
          <label>Payment amount<input name="paymentAmount" type="number" min="1" step="0.01" placeholder="28000" required /></label>
          <label>Payment method<select name="paymentMethod" defaultValue="UPI"><option>UPI</option><option>NEFT / RTGS</option><option>Cash</option><option>Cheque</option><option>Other</option></select></label>
          <label>Transaction reference<input name="transactionReference" placeholder="UTR / receipt number" required /></label>
        </>}
        {isInsurance && <>
          <label>Insurance company<input name="insuranceCompany" placeholder="Company name" required /></label>
          <label>Policy number<input name="policyNumber" placeholder="Policy number" required /></label>
          <label>Policy start date<input name="policyStartDate" type="date" required /></label>
          <label>Policy expiry date<input name="policyExpiryDate" type="date" required /></label>
          <label>Premium amount<input name="premiumAmount" type="number" min="1" step="0.01" placeholder="42000" required /></label>
          <label>Policy type<select name="policyType" defaultValue="Comprehensive"><option>Comprehensive</option><option>Third Party</option><option>Other</option></select></label>
        </>}
        {!isEmi && !isInsurance && <>
          <label>Document number<input name="documentNumber" placeholder="Permit / PUC / fitness number" required /></label>
          <label>Issue date<input name="issueDate" type="date" required /></label>
          <label>Expiry date<input name="expiryDate" type="date" required /></label>
        </>}
        <label className="span-2">Additional note<textarea name="customerNote" placeholder="Add a note for Admin (optional)" /></label>
        <label className="span-2">Proof document<input name="proofFile" type="file" accept="application/pdf,image/*" required /></label>
        <button className="span-2" type="submit"><ProofFileIcon />Submit for verification</button>
      </form>
    </details>
  );
}

function CustomerDocuments({ documents, vehicles, setSection }) {
  return (
    <section className="stack customer-documents-page">
      <div className="customer-section-header customer-documents-header">
        <div className="customer-documents-section-title">
          <p>Document center</p>
          <h2>Uploaded Documents</h2>
        </div>
        <span>{documents.length} document{documents.length !== 1 ? "s" : ""}</span>
      </div>
      <div className="customer-documents-grid">
        {documents.map((document) => {
          const vehicle = vehicles.find((item) => item.id === document.vehicleId);
          return (
            <article className="asset customer-document-card" key={document.id}>
              <div className="customer-document-card-head">
                <div className="customer-document-title">
                  <span className="customer-document-icon"><ProofFileIcon /></span>
                  <div>
                    <strong>{document.fileName}</strong>
                    <span>{document.type} <i aria-hidden="true">|</i> {vehicle?.regNo ?? "-"}</span>
                  </div>
                </div>
                <Badge label={`${Math.max(1, Math.round((document.size || 0) / 1024))} KB`} />
              </div>
              <div className="customer-document-meta">
                <Pair label="Uploaded by" value={document.uploadedBy || "-"} />
                <Pair label="Uploaded at" value={document.uploadedAt || "-"} />
                <Pair label="Note" value={document.note || "-"} />
              </div>
              {document.dataUrl && (
                <a className="button-link customer-document-open" href={document.dataUrl} target="_blank" rel="noreferrer">
                  <ProofFileIcon />
                  Open file
                </a>
              )}
            </article>
          );
        })}
        {documents.length === 0 && (
          <section className="customer-documents-empty">
            <span className="customer-documents-empty-icon"><ProofFileIcon /></span>
            <div className="customer-documents-empty-copy">
              <p>DOCUMENT CENTER</p>
              <h3>No documents uploaded yet</h3>
              <span>Payment receipts, insurance certificates and other proofs will appear here after submission.</span>
            </div>
            <button type="button" onClick={() => setSection?.("dues")}>
              <Icon name="upload" />
              Go to dues
            </button>
          </section>
        )}
      </div>
    </section>
  );
}

function CustomerMarketplace({ chat, listings, vehicles, saleClosings = [], myVehicles, myVehicleIds, submitListing, setSection }) {
  const [selectedListing, setSelectedListing] = useState(null);
  const [filters, setFilters] = useState({ assetType: "All", condition: "All", finance: "All", insurance: "All", location: "", maxPrice: "" });
  const [showFilters, setShowFilters] = useState(false);
  const publicListings = listings.filter((listing) => {
    const vehicle = vehicles.find((item) => item.id === listing.vehicleId);
    const closing = saleClosings.find((item) => item.listingId === listing.id);
    return listing.status !== "Sold" && vehicle?.status !== "Sold" && closing?.status !== "Sold";
  });
  const filteredListings = publicListings.filter((listing) => {
    const vehicle = vehicles.find((item) => item.id === listing.vehicleId);
    const vehicleLiability = vehicle ? liability(vehicle) : 0;
    const locationMatch = !filters.location.trim() || String(listing.location || "").toLowerCase().includes(filters.location.trim().toLowerCase());
    const priceMatch = !filters.maxPrice || Number(listing.price) <= Number(filters.maxPrice);
    const assetMatch = filters.assetType === "All" || vehicle?.type === filters.assetType;
    const conditionMatch = filters.condition === "All" || listing.condition === filters.condition;
    const financeMatch = filters.finance === "All" || (filters.finance === "Financed" ? vehicleLiability > 0 : vehicleLiability <= 0);
    const insuranceMatch = filters.insurance === "All" || (filters.insurance === "Insured" ? Boolean(vehicle?.insuranceExpiry) : !vehicle?.insuranceExpiry);
    return locationMatch && priceMatch && assetMatch && conditionMatch && financeMatch && insuranceMatch;
  });
  const resetFilters = () => setFilters({ assetType: "All", condition: "All", finance: "All", insurance: "All", location: "", maxPrice: "" });
  const activeFilterCount = Object.entries(filters).filter(([key, value]) => key === "location" || key === "maxPrice" ? Boolean(value) : value !== "All").length;
  return (
    <section className="stack customer-marketplace-page">
      <div className="customer-marketplace-hero">
        <span className="customer-marketplace-hero-icon"><MarketplaceStoreIcon /></span>
        <div className="customer-marketplace-hero-copy">
          <span className="customer-marketplace-eyebrow">Vehicle marketplace</span>
          <h2>Find your next vehicle</h2>
          <p>Explore active listings and share your vehicle with verified buyers.</p>
        </div>
        <div className="customer-marketplace-hero-stat">
          <strong>{listings.length}</strong>
          <span>Active listings</span>
        </div>
      </div>
      <div className="customer-marketplace-toolbar">
        <div>
          <span className="customer-marketplace-toolbar-label">Marketplace</span>
          <h2>Available listings</h2>
        </div>
        <div className="customer-marketplace-toolbar-actions">
          <span className="customer-marketplace-count">{filteredListings.length} listing{filteredListings.length !== 1 ? "s" : ""}</span>
          <button className="filter-trigger customer-filter-trigger" type="button" onClick={() => setShowFilters(true)}><Icon name="filter" />Filters{activeFilterCount ? ` (${activeFilterCount})` : ""}</button>
          <details className="customer-action-panel customer-sell-panel">
            <summary><Icon name="plus" /> Sell a vehicle</summary>
        <form className="form-grid compact-form" onSubmit={submitListing}>
          <label>Vehicle<select name="vehicleId" required defaultValue=""><option value="" disabled>Select your vehicle</option>{myVehicles.map((vehicle) => <option key={vehicle.id} value={vehicle.id}>{vehicle.regNo} | {vehicle.make} {vehicle.model}</option>)}</select></label>
          <label>Asking price<input name="price" type="number" min="1" step="0.01" placeholder="950000" required /></label>
          <label>Location<input name="location" placeholder="Ahmedabad" required /></label>
          <label>Condition<select name="condition" defaultValue="Good"><option>Excellent</option><option>Good</option><option>Average</option></select></label>
          <label className="span-2">Vehicle photos<input name="photos" type="file" accept="image/*" multiple /></label>
          <button className="span-2" type="submit"><Icon name="store" />Submit sale listing</button>
        </form>
        {myVehicles.length === 0 && <Empty text="No vehicle is available for a sale listing." />}
          </details>
        </div>
      </div>
      {showFilters && <button className="filter-drawer-backdrop" type="button" aria-label="Close filters" onClick={() => setShowFilters(false)} />}
      <aside className={`filter-drawer customer-filter-drawer ${showFilters ? "open" : ""}`} aria-label="Marketplace filters">
        <div className="filter-drawer-head"><div><p className="eyebrow">Refine results</p><h3>Marketplace filters</h3></div><button className="icon-button" type="button" aria-label="Close filters" onClick={() => setShowFilters(false)}><Icon name="close" /></button></div>
        <label>Asset type<select value={filters.assetType} onChange={(event) => setFilters((current) => ({ ...current, assetType: event.target.value }))}><option>All</option><option>Truck</option><option>Trailer</option></select></label>
        <label>Condition<select value={filters.condition} onChange={(event) => setFilters((current) => ({ ...current, condition: event.target.value }))}><option>All</option><option>Excellent</option><option>Good</option><option>Average</option></select></label>
        <label>Finance<select value={filters.finance} onChange={(event) => setFilters((current) => ({ ...current, finance: event.target.value }))}><option>All</option><option>Financed</option><option>Loan-free</option></select></label>
        <label>Insurance<select value={filters.insurance} onChange={(event) => setFilters((current) => ({ ...current, insurance: event.target.value }))}><option>All</option><option>Insured</option><option>Missing</option></select></label>
        <label>Location<input value={filters.location} onChange={(event) => setFilters((current) => ({ ...current, location: event.target.value }))} placeholder="Search location" /></label>
        <label>Max price<input type="number" min="0" value={filters.maxPrice} onChange={(event) => setFilters((current) => ({ ...current, maxPrice: event.target.value }))} placeholder="Any price" /></label>
        <div className="filter-drawer-actions"><button className="filter-reset" type="button" onClick={resetFilters}><Icon name="refresh" />Reset</button><button type="button" onClick={() => setShowFilters(false)}>Apply filters</button></div>
      </aside>
      {chat.error && <p className="mc-error" role="alert">{chat.error}</p>}
      <div className="customer-marketplace-listings">
        {filteredListings.map((l) => {
          const v = vehicles.find((x) => x.id === l.vehicleId);
          const closing = saleClosings.find((item) => item.listingId === l.id);
          const listingPhoto = l.photos?.[0];
          const mine = myVehicleIds?.has(l.vehicleId);
          return (
            <article className="asset" key={l.id}>
              {listingPhoto?.dataUrl && <img className="listing-photo" src={listingPhoto.dataUrl} alt={l.title} />}
              <div className="card-head">
                <div>
                  <strong>{l.title}</strong>
                  <span>{l.location} · {l.condition}</span>
                </div>
                <Badge label={mine ? "My listing" : l.status} />
              </div>
              <dl>
                <div><dt>Price</dt><dd>{formatMoney(l.price)}</dd></div>
                <div><dt>Vehicle</dt><dd>{v?.regNo ?? l.vehicleId}</dd></div>
                <div><dt>Photos</dt><dd>{l.photos?.length ?? 0}</dd></div>
                <div><dt>Estimated closing</dt><dd>{closing ? formatMoney(closing.estimatedAmount) : formatMoney(v ? liability(v) : 0)}</dd></div>
                <div><dt>Bank-confirmed closing</dt><dd>{closing?.bankConfirmedAmount ? formatMoney(closing.bankConfirmedAmount) : "Pending bank confirmation"}</dd></div>
              </dl>
              <div className="actions"><button type="button" onClick={() => setSelectedListing(l)}><Icon name="eye" />View details</button><ChatRequestButton chat={chat} listingId={l.id} mine={mine} openChats={() => setSection("chats")} /></div>
            </article>
          );
        })}
        {listings.length === 0 && (
          <section className="customer-marketplace-empty">
            <span className="customer-marketplace-empty-icon"><MarketplaceStoreIcon /></span>
            <div>
              <h3>No active listings yet</h3>
              <p>New vehicles approved for the marketplace will appear here.</p>
            </div>
            <button className="customer-marketplace-empty-action" type="button" onClick={() => setSection?.("fleet")}>
              <Icon name="truck" />
              View my fleet
            </button>
          </section>
        )}
        {listings.length > 0 && filteredListings.length === 0 && (
          <section className="customer-marketplace-empty customer-marketplace-filter-empty">
            <span className="customer-marketplace-empty-icon"><MarketplaceStoreIcon /></span>
            <div>
              <h3>No listings match these filters</h3>
              <p>Try a different filter or reset the current search.</p>
            </div>
            <button className="customer-marketplace-empty-action" type="button" onClick={resetFilters}>
              <Icon name="refresh" />
              Reset filters
            </button>
          </section>
        )}
      </div>
      {selectedListing && <MarketplaceVehicleModal listing={selectedListing} vehicle={vehicles.find((vehicle) => vehicle.id === selectedListing.vehicleId)} closing={saleClosings.find((item) => item.listingId === selectedListing.id)} onClose={() => setSelectedListing(null)} />}
    </section>
  );
}



function Metric({ label, value, icon, onClick }) {
  const content = (
    <>
      <span className="metric-icon"><Icon name={icon} /></span>
      <span className="metric-label">{label}</span>
      <strong>{value}</strong>
      <span className="metric-line" />
    </>
  );

  if (onClick) {
    return (
      <button className="metric metric-action" type="button" onClick={onClick}>
        {content}
      </button>
    );
  }

  return (
    <article className="metric">
      {content}
    </article>
  );
}

function CustomerMetric({ label, value, helper, icon, onClick, accent = false }) {
  const content = (
    <>
      <span className="customer-metric-icon"><Icon name={icon} /></span>
      <span className="customer-metric-label">{label}</span>
      <strong>{value}</strong>
      <small>{helper}</small>
    </>
  );

  if (onClick) {
    return <button className="customer-metric" type="button" onClick={onClick}>{content}</button>;
  }

  return <article className={`customer-metric${accent ? " accent" : ""}`}>{content}</article>;
}

function Panel({ title, children }) {
  return (
    <section className="panel">
      <h2>{title}</h2>
      {children}
    </section>
  );
}

function Pair({ label, value }) {
  return <p className="pair"><span>{label}</span><b>{displayPairValue(value)}</b></p>;
}

function displayPairValue(value) {
  if (value === null || value === undefined || value === "") return "-";
  if (typeof value === "object") return value.fileName || value.file_name || value.name || "File attached";
  const text = String(value);
  return ["[object File]", "[object Object]"].includes(text) ? "File attached" : text;
}

function ProofFileIcon() {
  return (
    <svg className="proof-file-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.8" aria-hidden="true">
      <path d="M6 3.5h8l4 4V20a1.5 1.5 0 0 1-1.5 1.5h-11A1.5 1.5 0 0 1 4 20V5A1.5 1.5 0 0 1 5.5 3.5Z" />
      <path d="M14 3.5V8h4" />
      <path d="M7.5 12h5M7.5 15.5h7" />
    </svg>
  );
}

function KuberBrandMark() {
  return <img className="kuber-brand-svg" src={myFleetLogo} alt="My Fleet" />;
}

function MarketplaceStoreIcon() {
  return (
    <svg className="marketplace-store-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.7" aria-hidden="true">
      <path d="M4 10.5h16V20H4z" />
      <path d="M3 10.5 5 4h14l2 6.5" />
      <path d="M3 10.5c.4 1.4 1.4 2.2 2.8 2.2s2.4-.8 2.8-2.2c.4 1.4 1.4 2.2 2.8 2.2s2.4-.8 2.8-2.2c.4 1.4 1.4 2.2 2.8 2.2s2.4-.8 2.8-2.2" />
      <path d="M8 20v-4h8v4" />
    </svg>
  );
}

function ChatBubbleIcon() {
  return (
    <svg className="chat-bubble-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.7" aria-hidden="true">
      <path d="M5 5.5h14a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H11l-4.5 3v-3H5a2 2 0 0 1-2-2v-8a2 2 0 0 1 2-2Z" />
      <path d="M7.5 10h9M7.5 13.5h5.5" />
    </svg>
  );
}

function PaperclipIcon() {
  return (
    <svg className="paperclip-svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.8" aria-hidden="true">
      <path d="m8.5 12.5 5.9-5.9a3.25 3.25 0 0 1 4.6 4.6l-7.7 7.7a5 5 0 0 1-7.1-7.1l7.1-7.1" />
      <path d="m9.4 14.6 6.1-6.1a1.5 1.5 0 0 1 2.1 2.1l-6.1 6.1a2.5 2.5 0 0 1-3.5-3.5l6.1-6.1" />
    </svg>
  );
}

function Empty({ text }) {
  return <p className="empty">{text}</p>;
}

function AccessDenied({ role, section }) {
  return (
    <section className="access-denied">
      <Icon name="settings" />
      <div>
        <h2>Access blocked</h2>
        <p>{role} does not have permission for {titleFor(section)}.</p>
      </div>
    </section>
  );
}

function Badge({ label }) {
  const tone = String(label).toLowerCase().replace(/\s+/g, "-");
  const toneClass = tone === "due" ? "status-due" : tone;
  return <span className={`badge ${toneClass}`}>{label}</span>;
}

const iconPaths = {
  home: <><path d="m3 11 9-8 9 8" /><path d="M5 10v10h14V10" /></>,
  people: <><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" /><circle cx="9" cy="7" r="4" /><path d="M22 21v-2a4 4 0 0 0-3-3.87M16 3.13a4 4 0 0 1 0 7.75" /></>,
  truck: <><path d="M10 17h4V5H2v12h3" /><path d="M14 8h4l4 4v5h-3" /><circle cx="7.5" cy="17.5" r="2.5" /><circle cx="16.5" cy="17.5" r="2.5" /></>,
  calendar: <><rect x="3" y="4" width="18" height="18" rx="2" /><path d="M16 2v4M8 2v4M3 10h18" /></>,
  clock: <><circle cx="12" cy="12" r="9" /><path d="M12 7v5l3 2" /></>,
  check: <path d="m5 12 4.2 4.2L19 6.5" />,
  phone: <path d="M22 16.92v3a2 2 0 0 1-2.18 2A19.8 19.8 0 0 1 3.1 5.18 2 2 0 0 1 5.11 3h3a2 2 0 0 1 2 1.72c.12.9.33 1.77.62 2.61a2 2 0 0 1-.45 2.11L9 10.72a16 16 0 0 0 4.28 4.28l1.28-1.28a2 2 0 0 1 2.11-.45c.84.29 1.71.5 2.61.62A2 2 0 0 1 22 16.92z" />,
  store: <><path d="M4 10v11h16V10" /><path d="M3 10h18L19 3H5z" /><path d="M9 21v-6h6v6" /></>,
  chart: <><path d="M3 3v18h18" /><path d="M7 16V9M12 16V5M17 16v-3" /></>,
  download: <><path d="M12 3v12" /><path d="m7 10 5 5 5-5" /><path d="M5 21h14" /></>,
  upload: <><path d="M12 16V3" /><path d="m7 8 5-5 5 5" /><path d="M4 14v6h16v-6" /></>,
  settings: <><circle cx="12" cy="12" r="3" /><path d="M19.4 15a1.7 1.7 0 0 0 .34 1.88l.06.06a2 2 0 1 1-2.83 2.83l-.06-.06A1.7 1.7 0 0 0 15 19.4a1.7 1.7 0 0 0-1 .6 1.7 1.7 0 0 0-.4 1.1V21a2 2 0 1 1-4 0v-.09A1.7 1.7 0 0 0 8.6 19.4a1.7 1.7 0 0 0-1.88.34l-.06.06a2 2 0 1 1-2.83-2.83l.06-.06A1.7 1.7 0 0 0 4.6 15a1.7 1.7 0 0 0-.6-1 1.7 1.7 0 0 0-1.1-.4H3a2 2 0 1 1 0-4h.09A1.7 1.7 0 0 0 4.6 8.6a1.7 1.7 0 0 0-.34-1.88l-.06-.06A2 2 0 1 1 7.03 3.83l.06.06A1.7 1.7 0 0 0 9 4.6a1.7 1.7 0 0 0 1-.6 1.7 1.7 0 0 0 .4-1.1V3a2 2 0 1 1 4 0v.09A1.7 1.7 0 0 0 15.4 4.6a1.7 1.7 0 0 0 1.88-.34l-.06-.06A2 2 0 1 1 20.11 7l-.06.06A1.7 1.7 0 0 0 19.4 9c.4.25.75.6 1 1 .28.45.68.7 1.1.7H21a2 2 0 1 1 0 4h-.09a1.7 1.7 0 0 0-1.51.9z" /></>,
  cloud: <path d="M17.5 19H7a5 5 0 1 1 1.1-9.88A7 7 0 0 1 21 12.5 3.5 3.5 0 0 1 17.5 19z" />,
  money: <><rect x="2" y="6" width="20" height="12" rx="2" /><circle cx="12" cy="12" r="2" /><path d="M6 12h.01M18 12h.01" /></>,
  plus: <path d="M12 5v14M5 12h14" />,
  bell: <><path d="M18 8a6 6 0 0 0-12 0c0 7-3 7-3 9h18c0-2-3-2-3-9" /><path d="M13.73 21a2 2 0 0 0-3.46 0" /></>,
  eye: <><path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12z" /><circle cx="12" cy="12" r="3" /></>,
  search: <><circle cx="11" cy="11" r="8" /><path d="m21 21-4.35-4.35" /></>,
  trash: <><path d="M3 6h18M8 6V4h8v2M19 6l-1 14H6L5 6" /><path d="M10 11v6M14 11v6" /></>,
  logout: <><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4" /><path d="m16 17 5-5-5-5M21 12H9" /></>,
  filter: <path d="M4 6h16M7 12h10M10 18h4" />,
  close: <path d="m6 6 12 12M18 6 6 18" />,
  refresh: <><path d="M20 11a8 8 0 0 0-14.7-3L3 11" /><path d="M3 4v7h7M4 13a8 8 0 0 0 14.7 3L21 13" /><path d="M21 20v-7h-7" /></>,
  "chevron-down": <path d="m7 9 5 5 5-5" />,
  "chevron-right": <path d="m9 6 6 6-6 6" />
};

function Icon({ name }) {
  return (
    <svg
      className={`icon icon-${name}`}
      viewBox="0 0 24 24"
      fill="none"
      stroke="currentColor"
      strokeLinecap="round"
      strokeLinejoin="round"
      strokeWidth="1.8"
      aria-hidden="true"
      focusable="false"
    >
      {iconPaths[name] ?? iconPaths.settings}
    </svg>
  );
}

function AlertSvg({ name }) {
  const common = { fill: "none", stroke: "currentColor", strokeLinecap: "round", strokeLinejoin: "round", strokeWidth: 1.8 };
  if (name === "check") {
    return <svg className="alert-svg" viewBox="0 0 24 24" aria-hidden="true" {...common}><path d="m5 12 4.2 4.2L19 6.5" /></svg>;
  }
  if (name === "bell") {
    return <svg className="alert-svg" viewBox="0 0 24 24" aria-hidden="true" {...common}><path d="M18 8a6 6 0 0 0-12 0c0 6.5-2.5 7.5-2.5 9h17c0-1.5-2.5-2.5-2.5-9Z" /><path d="M10 21h4" /></svg>;
  }
  if (name === "money") {
    return <svg className="alert-svg" viewBox="0 0 24 24" aria-hidden="true" {...common}><rect x="2.5" y="5.5" width="19" height="13" rx="2" /><circle cx="12" cy="12" r="2.4" /><path d="M6 9h.01M18 15h.01" /></svg>;
  }
  return <svg className="alert-svg" viewBox="0 0 24 24" aria-hidden="true" {...common}><rect x="3.5" y="5" width="17" height="16" rx="2" /><path d="M7.5 3v4M16.5 3v4M3.5 10h17" /></svg>;
}

function titleFor(section) {
  if (section === "client-profile") return "Client Profile";
  return navItems.find(([key]) => key === section)?.[1] ?? customerNavItems.find(([key]) => key === section)?.[1] ?? "Overview";
}

function roleAccessName(role) {
  return role || "Customer";
}

function permissionRoleName(role) {
  return role === "Customer" ? "Buyer" : (role || "Buyer");
}

function permissionValueFor(data, role, feature) {
  const rows = Array.isArray(data.rolePermissions) ? data.rolePermissions : permissionRows;
  const row = rows.find((item) => item[0] === feature);
  const index = rolePermissionIndex[permissionRoleName(role)] ?? rolePermissionIndex.Buyer;
  return row?.[index] ?? "No";
}

function canUsePermission(data, role, feature) {
  if (!feature) return true;
  if (role === "Admin") return true;
  return permissionValueFor(data, role, feature) !== "No";
}

function canOpenSection(data, role, section) {
  return canUsePermission(data, role, sectionPermissionMap[section]);
}

function getDataClient(data, clientId) {
  return data.clients.find((client) => client.id === clientId) ?? getClient(clientId);
}

function getDataVehicle(data, vehicleId) {
  return data.vehicles.find((vehicle) => vehicle.id === vehicleId) ?? getVehicle(vehicleId);
}

function toNumber(value) {
  const number = Number(String(value ?? "").replace(/[^\d.-]/g, ""));
  return Number.isFinite(number) ? number : 0;
}

function formatMaybeMoney(value) {
  const number = toNumber(value);
  return number > 0 ? formatMoney(number) : value || "-";
}

function formatPlainMoney(value) {
  const number = toNumber(value);
  return number > 0 ? new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(number) : value || "-";
}

function formatDisplayDate(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "-") return "-";
  const isoDateTimeMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:T|$)/);
  if (isoDateTimeMatch) {
    const [, year, month, day] = isoDateTimeMatch;
    return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year}`;
  }
  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (isoMatch) {
    const [, year, month, day] = isoMatch;
    return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year}`;
  }
  const localMatch = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$/);
  if (localMatch) {
    let [, day, month, year] = localMatch;
    // Accept both Indian DD-MM-YYYY and spreadsheet-style MM-DD-YYYY.
    if (Number(day) <= 12 && Number(month) > 12) [day, month] = [month, day];
    return `${day.padStart(2, "0")}-${month.padStart(2, "0")}-${year.length === 2 ? `20${year}` : year}`;
  }
  const namedMonthMatch = text.match(/^(\d{1,2})[-\s]([A-Za-z]{3,})[-\s](\d{2,4})$/);
  if (namedMonthMatch) {
    const [, day, monthName, year] = namedMonthMatch;
    const monthIndex = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"].indexOf(String(monthName || "").slice(0, 3).toLowerCase());
    if (monthIndex >= 0) return `${day.padStart(2, "0")}-${String(monthIndex + 1).padStart(2, "0")}-${year.length === 2 ? `20${year}` : year}`;
  }
  return text;
}

function parseDisplayDate(value) {
  const text = String(value ?? "").trim();
  if (!text || text === "-") return null;
  const formatted = formatDisplayDate(text);
  const match = formatted.match(/^(\d{2})-(\d{2})-(\d{4})$/);
  if (!match) return null;
  const [, day, month, year] = match;
  const date = new Date(Number(year), Number(month) - 1, Number(day));
  if (Number.isNaN(date.getTime())) return null;
  return date.getFullYear() === Number(year) &&
    date.getMonth() === Number(month) - 1 &&
    date.getDate() === Number(day) ? date : null;
}

function buildLoanPlan(fields) {
  const tenure = Math.min(Math.max(Math.round(toNumber(fields.tenure)), 0), 360);
  const emiAmount = toNumber(fields.emiAmount);
  const loanAmount = toNumber(fields.loanAmount);
  if (tenure <= 0 || emiAmount <= 0) return { schedule: [], history: [] };
  const start = parseDisplayDate(fields.emiStart) || new Date();
  const rate = toNumber(fields.interestRate);
  const paidEmi = Math.min(Math.max(Math.round(toNumber(fields.paidEmi)), 0), tenure);
  let openingPrincipal = loanAmount;
  const schedule = [];
  const history = [];
  const toIsoDate = (date) => {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  };
  for (let index = 0; index < tenure; index += 1) {
    const dueDate = new Date(start.getFullYear(), start.getMonth() + index, start.getDate());
    const interest = Math.round((openingPrincipal * rate) / 1200 * 100) / 100;
    const principal = Math.min(Math.max(emiAmount - interest, 0), Math.max(openingPrincipal, 0));
    const closing = Math.max(openingPrincipal - principal, 0);
    const paid = index < paidEmi;
    const entry = {
      installment: index + 1,
      dueDate: toIsoDate(dueDate),
      amount: emiAmount,
      principal,
      interest,
      status: paid ? "Paid" : "Due"
    };
    schedule.push(entry);
    if (paid) {
      history.push({
        installment: index + 1,
        amount: emiAmount,
        paidOn: toIsoDate(dueDate),
        reference: `EMI-${String(index + 1).padStart(3, "0")}`,
        status: "Verified"
      });
    }
    openingPrincipal = closing;
  }
  return { schedule, history };
}

function complianceStatusForDate(dateText) {
  const date = parseDisplayDate(dateText);
  if (!date) return "Pending Verification";
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return date < today ? "Expired" : "Active";
}

function buildInsuranceHistoryEntry(fields) {
  if (!fields.insuranceCompany && !fields.insurancePolicyNo && !fields.insuranceExpiry) return null;
  return {
    company: fields.insuranceCompany || "Unknown insurer",
    policyNo: fields.insurancePolicyNo || "Not provided",
    startDate: fields.insuranceStart || "",
    expiryDate: fields.insuranceExpiry || "",
    status: fields.insuranceStatus === "Pending Verification"
      ? "Pending Verification"
      : fields.insuranceStatus === "Expired"
        ? "Expired"
        : complianceStatusForDate(fields.insuranceExpiry)
  };
}

function buildComplianceHistoryEntries(fields) {
  const records = [
    ["Permit", fields.permitNo, fields.permitIssue, fields.permitExpiry],
    ["PUC", fields.pucNo, "", fields.pucExpiry],
    ["Fitness", "", "", fields.fitnessExpiry],
    ["Road Tax", "", "", fields.nationalPermitExpiry]
  ];
  return records
    .filter(([, documentNo, issueDate, expiryDate]) => documentNo || issueDate || expiryDate)
    .map(([kind, documentNo, issueDate, expiryDate]) => ({
      kind,
      documentNo: documentNo || "Not provided",
      issueDate: issueDate || "",
      expiryDate: expiryDate || "",
      status: fields.complianceStatus === "Pending Verification"
        ? "Pending Verification"
        : fields.complianceStatus === "Expired"
          ? "Expired"
          : complianceStatusForDate(expiryDate)
    }));
}

function upsertHistory(history, entry, keyFor) {
  const current = Array.isArray(history) ? history : [];
  const key = keyFor(entry);
  const index = current.findIndex((item) => keyFor(item) === key);
  if (index < 0) return [...current, entry];
  return current.map((item, itemIndex) => itemIndex === index ? entry : item);
}

function mergeComplianceHistory(history, entries) {
  return entries.reduce(
    (result, entry) => upsertHistory(result, entry, (item) => `${item.kind}:${item.documentNo}:${item.expiryDate}`),
    Array.isArray(history) ? history : []
  );
}

function smartAlertStatus(daysLeft) {
  if (daysLeft < 0) return "Expired";
  if (daysLeft === 0) return "Today";
  if (daysLeft <= 10) return "Soon";
  return "Upcoming";
}

function smartAlertDetail(daysLeft) {
  if (daysLeft < 0) return `${Math.abs(daysLeft)} days overdue`;
  if (daysLeft === 0) return "Expires today";
  return `${daysLeft} days left`;
}

function buildSmartAlerts(data, includeAll = false) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const alerts = [];
  const seen = new Set();
  const pushAlert = ({ clientId, clientName, vehicle, type, date }) => {
    const parsedDate = parseDisplayDate(date);
    if (!parsedDate) return;
    parsedDate.setHours(0, 0, 0, 0);
    const daysLeft = Math.round((parsedDate - today) / 86400000);
    const normalizedVehicle = baseRegNo(vehicle);
    const vehicleKey = normalizeRegNo(normalizedVehicle).replace(/[^A-Z0-9]/g, "");
    const expiryKey = formatDisplayDate(date).replace(/[^0-9]/g, "");
    const id = `${vehicleKey}-${type}-${expiryKey}`;
    if (seen.has(id)) return;
    seen.add(id);
    alerts.push({
      id,
      clientId,
      client: clientName || getDataClient(data, clientId)?.name || "-",
      vehicle: normalizedVehicle || "-",
      type,
      date,
      daysLeft,
      detail: smartAlertDetail(daysLeft),
      status: smartAlertStatus(daysLeft)
    });
  };

  data.vehicles.forEach((vehicle) => {
    const clientName = getDataClient(data, vehicle.clientId)?.name;
    pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "Policy", date: vehicle.insuranceExpiry || vehicle.policyEnd });
    pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "Permit", date: vehicle.permitExpiry || vehicle.permitExpired });
    pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "National Permit", date: vehicle.nationalPermitExpiry || vehicle.nationalPermitExpired });
    pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "PUC", date: vehicle.pucExpiry || vehicle.pucExpired });
    pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "Fitness", date: vehicle.fitnessExpiry || vehicle.fitnessExpired });

    (vehicle.insuranceHistory ?? []).forEach((entry) => {
      pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type: "Policy", date: entry.expiryDate || entry.expiry });
    });
    (vehicle.complianceHistory ?? []).forEach((entry) => {
      const type = entry.kind === "Road Tax" ? "National Permit" : entry.kind;
      if (["Permit", "National Permit", "PUC", "Fitness"].includes(type)) {
        pushAlert({ clientId: vehicle.clientId, clientName, vehicle: vehicle.regNo, type, date: entry.expiryDate || entry.expiry });
      }
    });
  });

  (data.clientImports ?? []).forEach((importItem) => {
    const clientName = getDataClient(data, importItem.clientId)?.name;
    (importItem.rows ?? []).filter((row) => row.regNo && !isBodyRow(row)).forEach((row) => {
      pushAlert({ clientId: importItem.clientId, clientName, vehicle: row.regNo, type: "Policy", date: row.policyEnd || row.policyExpiry || row.insuranceExpiry });
      pushAlert({ clientId: importItem.clientId, clientName, vehicle: row.regNo, type: "PUC", date: row.pucExpired || row.pucExpiry });
      pushAlert({ clientId: importItem.clientId, clientName, vehicle: row.regNo, type: "Fitness", date: row.fitnessExpired || row.fitnessExpiry });
      pushAlert({ clientId: importItem.clientId, clientName, vehicle: row.regNo, type: "Permit", date: row.permitExpired || row.permitExpiry });
      pushAlert({ clientId: importItem.clientId, clientName, vehicle: row.regNo, type: "National Permit", date: row.nationalPermitExpired || row.nationalPermitExpiry });
    });
  });

  return alerts
    .filter((alert) => includeAll || alert.daysLeft <= 10)
    .sort((first, second) => first.daysLeft - second.daysLeft || first.vehicle.localeCompare(second.vehicle));
}

function autoClosingPrincipal(row) {
  if (!row) return 0;
  if (isClosedFinanceRow(row)) return 0;
  if (row.bankClosingPrincipalUnknown === true) return 0;
  const bankClosing = toNumber(row.bankClosingPrincipal ?? row.bank_closing_principal);
  const savedClosing = toNumber(row.closingPrincipal ?? row.closing_principal);
  const loanAmount = toNumber(row.loanAmount);
  const emiAmount = toNumber(row.emiAmount);
  const tenure = toNumber(row.tenure);
  const paidEmi = toNumber(row.paidEmi);
  const enteredRate = String(row.interestRate ?? "").trim();
  const interestRate = enteredRate ? toNumber(enteredRate) : BANK_RELEASE_RATE_PERCENT;
  const paidCount = effectivePaidEmi(row);
  const snapshotClosing = bankClosing || savedClosing;
  if (snapshotClosing > 0) {
    const extraPaidCount = Math.max(paidCount - Math.min(Math.max(Math.round(paidEmi), 0), tenure || paidCount), 0);
    if (extraPaidCount <= 0 || emiAmount <= 0) return snapshotClosing;
    return amortizeClosingPrincipal(snapshotClosing, emiAmount, interestRate, extraPaidCount);
  }
  // If the bank closing column was unreadable, use the validated schedule's
  // principal breakup instead of carrying forward an older/stale POS value.
  // This is exact for AU schedules: loan amount minus principal paid in all
  // paid installments equals the bank's closing principal.
  const schedulePrincipalPaid = Array.isArray(row.emiSchedule)
    ? row.emiSchedule
      .filter((entry) => entry?.status === "Paid")
      .reduce((sum, entry) => sum + Math.max(toNumber(entry?.principal), 0), 0)
    : 0;
  if (loanAmount > 0 && schedulePrincipalPaid > 0) {
    return Math.max(loanAmount - schedulePrincipalPaid, 0);
  }
  if (loanAmount <= 0 || emiAmount <= 0 || tenure <= 0) return savedClosing;
  return amortizeClosingPrincipal(loanAmount, emiAmount, interestRate, paidCount);
}

function amortizeClosingPrincipal(openingPrincipal, emiAmount, interestRate, paidCount) {
  let closingPrincipal = openingPrincipal;
  for (let count = 0; count < paidCount; count += 1) {
    const monthlyInterest = (closingPrincipal * interestRate) / 1200;
    const principalPaid = emiAmount - monthlyInterest;
    closingPrincipal -= Math.max(principalPaid, 0);
  }
  return Math.max(closingPrincipal, 0);
}

function formatAutoClosingPrincipal(row) {
  if (isClosedFinanceRow(row)) return "0";
  const closing = autoClosingPrincipal(row);
  return closing > 0 ? new Intl.NumberFormat("en-IN", { maximumFractionDigits: 0 }).format(closing) : "-";
}

function effectivePaidEmi(row, asOf = new Date()) {
  if (!row) return 0;
  if (isClosedFinanceRow(row) && toNumber(row.tenure) > 0) return Math.round(toNumber(row.tenure));
  if (row.paidEmiUnknown === true) return 0;
  const tenure = Math.min(Math.max(Math.round(toNumber(row.tenure)), 0), 360);
  const savedPaid = Math.min(Math.max(Math.round(toNumber(row.paidEmi)), 0), tenure || 360);
  const today = new Date(asOf);
  today.setHours(0, 0, 0, 0);
  const schedule = Array.isArray(row.emiSchedule) ? row.emiSchedule : [];
  const dueFromSchedule = schedule.filter((entry) => {
    const dueDate = parseDisplayDate(entry?.dueDate);
    if (!dueDate) return false;
    dueDate.setHours(0, 0, 0, 0);
    return dueDate <= today;
  }).length;
  if (dueFromSchedule > 0) return Math.min(Math.max(savedPaid, dueFromSchedule), tenure || schedule.length || 360);
  const start = parseDisplayDate(row.emiStart);
  if (!start) return savedPaid;
  start.setHours(0, 0, 0, 0);
  if (today < start) return savedPaid;
  const maxCount = tenure || 360;
  let dueCount = 0;
  for (let index = 0; index < maxCount; index += 1) {
    const dueDate = new Date(start.getFullYear(), start.getMonth() + index, start.getDate());
    dueDate.setHours(0, 0, 0, 0);
    if (dueDate > today) break;
    dueCount += 1;
  }
  return Math.min(Math.max(savedPaid, dueCount), maxCount);
}

function importedFinanceAssetForVehicle(vehicle, importedAssets = []) {
  const vehicleReg = normalizeRegNo(baseRegNo(vehicle.regNo));
  const vehicleLoanAccount = normalizeRegNo(vehicle.loanAccount || "");
  const matches = importedAssets.filter((asset) => {
    if (isBodyRow(asset)) return false;
    const assetReg = normalizeRegNo(baseRegNo(asset.regNo));
    const assetLoanAccount = normalizeRegNo(asset.loanAccount || "");
    return (vehicleReg && assetReg === vehicleReg)
      || (vehicleLoanAccount && assetLoanAccount && assetLoanAccount === vehicleLoanAccount);
  });
  return matches.find((asset) => asset.pdfImportedAt || asset.bankClosingPrincipal || asset.bank_closing_principal) ?? matches[0];
}

function customerVehicleClosingPrincipal(vehicle, importedAssets = []) {
  const importedAsset = importedFinanceAssetForVehicle(vehicle, importedAssets);
  if (importedAsset?.bankClosingPrincipalUnknown === true) return 0;
  const importedClosing = autoClosingPrincipal(importedAsset);
  return importedClosing > 0 ? importedClosing : Number(vehicle.principal || 0);
}

function customerVehiclePaidEmi(vehicle, importedAssets = []) {
  const importedAsset = importedFinanceAssetForVehicle(vehicle, importedAssets);
  if (importedAsset?.paidEmiUnknown === true) return 0;
  return effectivePaidEmi(importedAsset) || Math.round(toNumber(vehicle.paidEmi));
}

function normalizeRegNo(value) {
  return String(value ?? "").replace(/\s+/g, "").toUpperCase();
}

function normalizeAgreement(value) {
  return String(value ?? "").replace(/[^a-z0-9]/gi, "").toUpperCase();
}

function normalizeAgreementLoose(value) {
  return normalizeAgreement(value)
    .replace(/O/g, "0")
    .replace(/I/g, "1")
    .replace(/L/g, "1");
}

function agreementMatches(first, second) {
  const left = normalizeAgreement(first);
  const right = normalizeAgreement(second);
  const looseLeft = normalizeAgreementLoose(first);
  const looseRight = normalizeAgreementLoose(second);
  if (!left || !right) return false;
  if (left === right) return true;
  if (looseLeft === looseRight) return true;
  if (left.length >= 8 && right.length >= 8 && (left.includes(right) || right.includes(left))) return true;
  if (looseLeft.length >= 8 && looseRight.length >= 8 && (looseLeft.includes(looseRight) || looseRight.includes(looseLeft))) return true;
  const leftTail = left.slice(-8);
  const rightTail = right.slice(-8);
  const looseLeftTail = looseLeft.slice(-8);
  const looseRightTail = looseRight.slice(-8);
  return (leftTail.length >= 6 && leftTail === rightTail) ||
    (looseLeftTail.length >= 6 && looseLeftTail === looseRightTail);
}

function isValidAgreementValue(value) {
  const normalized = normalizeAgreement(value);
  return normalized.length >= 6 && /\d/.test(normalized);
}

function isClosedFinanceRow(row = {}) {
  return /closed|foreclos|fully\s+paid|settled|no\s+dues|noc\s+issued/i.test([
    row.loanStatus,
    row.financeStatus,
    row.remarks
  ].filter(Boolean).join(" "));
}

function normalizePdfFinanceRow(row = {}) {
  const next = { ...row };
  const startDate = parseDisplayDate(next.emiStart);
  const endDate = parseDisplayDate(next.emiEnd);
  // A bank statement cannot have its maturity date before its first EMI date.
  // Vision models occasionally swap these two fields, so correct only this
  // unambiguous ordering error.
  if (startDate && endDate && endDate < startDate) {
    const originalStart = next.emiStart;
    next.emiStart = next.emiEnd;
    next.emiEnd = originalStart;
  }
  const tenure = toNumber(next.tenure);
  const paidEmi = toNumber(next.paidEmi);
  const scheduleFullyPaid = next.scheduleParsed === "yes" && tenure > 0 && paidEmi >= tenure && toNumber(next.bankClosingPrincipal) <= 0;
  if (isClosedFinanceRow(next) || scheduleFullyPaid) {
    if (toNumber(next.tenure) > 0) next.paidEmi = toNumber(next.tenure);
    next.financeStatus = "CLOSED";
    next.bankClosingPrincipal = 0;
    next.closingPrincipal = 0;
  }
  return next;
}

function sanitizePdfRowForMerge(row) {
  const scheduleParsed = row.scheduleParsed === "yes";
  const hasValidLoan = (toNumber(row.loanAmount) > 0 || toNumber(row.emiAmount) > 0);
  const trustedPdfFields = scheduleParsed || row._pdfFieldsTrusted === true || hasValidLoan;
  // EMI amount and dates are safe fallback inputs even when a bank's table layout is not recognized.
  const blockedWithoutSchedule = new Set(["loanAmount", "tenure", "paidEmi", "interestRate", "bankClosingPrincipal"]);
  return Object.fromEntries(Object.entries(row).filter(([key, value]) => {
    if (key.startsWith("_")) return false;
    // Preserve an explicit blank schedule closing value so a previous bad
    // import cannot survive a corrected PDF re-import.
    if (key === "bankClosingPrincipal" && scheduleParsed && value === "") return true;
    if (!value) return false;
    if (Array.isArray(value) && value.length === 0) return false;
    if (blockedWithoutSchedule.has(key) && !trustedPdfFields) return false;
    return true;
  }));
}

function mergePdfFinanceAsset(existing = {}, incoming = {}) {
  const next = { ...existing, ...incoming };
  // A finance import updates the loan; keep the user's saved owner and plate.
  for (const key of ["owner", "regNo"]) {
    if (String(existing[key] ?? "").trim() && existing[key] !== "-") next[key] = existing[key];
  }
  return next;
}

function formText(form, name) {
  return form.get(name)?.toString().trim() ?? "";
}

function mapManualVehicleForm(form) {
  return {
    regNo: formText(form, "regNo"),
    owner: formText(form, "owner"),
    manufacturer: formText(form, "manufacturer"),
    model: formText(form, "model"),
    yearOfMfg: formText(form, "yearOfMfg"),
    regDate: formText(form, "regDate"),
    loanId: formText(form, "loanId"),
    financeStatus: formText(form, "financeStatus"),
    financier: formText(form, "financier"),
    loanAccount: formText(form, "loanAccount"),
    loanAmount: formText(form, "loanAmount"),
    emiAmount: formText(form, "emiAmount"),
    interestRate: formText(form, "interestRate"),
    tenure: formText(form, "tenure"),
    paidEmi: formText(form, "paidEmi"),
    emiStart: formText(form, "emiStart"),
    emiEnd: formText(form, "emiEnd"),
    closingPrincipal: formText(form, "closingPrincipal"),
    bodyRegNo: formText(form, "bodyRegNo") || "BODY",
    bodyFinanceStatus: formText(form, "bodyFinanceStatus"),
    bodyFinancier: formText(form, "bodyFinancier"),
    bodyLoanAccount: formText(form, "bodyLoanAccount"),
    bodyLoanAmount: formText(form, "bodyLoanAmount"),
    bodyEmiAmount: formText(form, "bodyEmiAmount"),
    bodyInterestRate: formText(form, "bodyInterestRate"),
    bodyTenure: formText(form, "bodyTenure"),
    bodyPaidEmi: formText(form, "bodyPaidEmi"),
    bodyEmiStart: formText(form, "bodyEmiStart"),
    bodyEmiEnd: formText(form, "bodyEmiEnd"),
    bodyClosingPrincipal: formText(form, "bodyClosingPrincipal"),
    policyCompany: formText(form, "policyCompany"),
    policyNo: formText(form, "policyNo"),
    policyStart: formText(form, "policyStart"),
    policyEnd: formText(form, "policyEnd"),
    pucNo: formText(form, "pucNo"),
    pucExpired: formText(form, "pucExpired"),
    fitnessExpired: formText(form, "fitnessExpired"),
    permitNo: formText(form, "permitNo"),
    permitIssue: formText(form, "permitIssue"),
    permitExpired: formText(form, "permitExpired"),
    permitType: formText(form, "permitType"),
    nationalPermitExpired: formText(form, "nationalPermitExpired")
  };
}

function isBodyRegNo(value) {
  return /\bbody\b/i.test(String(value ?? ""));
}

function isBodyRow(row) {
  return isBodyRegNo(row?.regNo);
}

function baseRegNo(value) {
  return String(value ?? "").replace(/\s+body\b/i, "").trim();
}

function excelRowToVehicle(row, clientId, id) {
  const importedSchedule = Array.isArray(row.emiSchedule) ? row.emiSchedule.filter((entry) => entry?.dueDate && toNumber(entry.amount) > 0) : [];
  const loanPlan = importedSchedule.length
    ? {
        schedule: importedSchedule,
        history: importedSchedule.filter((entry) => entry.status === "Paid").map((entry) => ({
          installment: Number(entry.installment),
          amount: toNumber(entry.amount),
          paidOn: entry.paidAt || entry.dueDate,
          reference: entry.reference || `EMI-${String(entry.installment).padStart(3, "0")}`,
          status: "Verified"
        }))
      }
    : buildLoanPlan(row);
  const insuranceHistory = buildInsuranceHistoryEntry({
    insuranceCompany: row.policyCompany || "",
    insurancePolicyNo: row.policyNo || "",
    insuranceStart: row.policyStart || "",
    insuranceExpiry: row.policyEnd || "",
    insuranceStatus: "Active"
  });
  const complianceHistory = buildComplianceHistoryEntries({
    permitNo: row.permitNo || "",
    permitIssue: row.permitIssue || "",
    permitExpiry: row.permitExpired || "",
    nationalPermitExpiry: row.nationalPermitExpired || "",
    pucNo: row.pucNo || "",
    pucExpiry: row.pucExpired || "",
    fitnessExpiry: row.fitnessExpired || "",
    complianceStatus: "Active"
  });
  return {
    id,
    clientId,
    type: "Truck",
    regNo: baseRegNo(row.regNo),
    make: row.manufacturer || "Imported",
    model: row.model || "Vehicle",
    year: Number(String(row.yearOfMfg || "").match(/\d{4}/)?.[0]) || new Date().getFullYear(),
    km: 0,
    principal: autoClosingPrincipal(row) || toNumber(row.loanAmount),
    overdue: 0,
    penalty: 0,
    foreclosure: 0,
    loanId: row.loanId || `LOAN-${String(row.regNo || id).replace(/[^A-Z0-9]/gi, "").toUpperCase()}`,
    loanAccount: row.loanAccount || "",
    financier: row.financier || "",
    loanAmount: toNumber(row.loanAmount),
    emiAmount: toNumber(row.emiAmount),
    interestRate: toNumber(row.interestRate),
    tenure: toNumber(row.tenure),
    paidEmi: toNumber(row.paidEmi),
    emiStart: row.emiStart || "",
    emiEnd: row.emiEnd || "",
    emiSchedule: loanPlan.schedule,
    emiHistory: loanPlan.history,
    insuranceCompany: row.policyCompany || "",
    insurancePolicyNo: row.policyNo || "",
    insuranceStart: row.policyStart || "",
    insuranceHistory: insuranceHistory ? [insuranceHistory] : [],
    permitNo: row.permitNo || "",
    permitIssue: row.permitIssue || "",
    permitType: row.permitType || "",
    nationalPermitExpiry: row.nationalPermitExpired || "",
    pucNo: row.pucNo || "",
    pucExpiry: row.pucExpired || "",
    fitnessExpiry: row.fitnessExpired || "",
    complianceHistory,
    insuranceExpiry: row.policyEnd || "-",
    permitExpiry: row.permitExpired || row.nationalPermitExpired || "-",
    status: "Active"
  };
}

function mergePdfIntoVehicle(vehicle, row) {
  const next = { ...vehicle };
  const textFields = [
    ["loanId", "loanId"],
    ["loanAccount", "loanAccount"],
    ["financier", "financier"],
    ["make", "manufacturer"],
    ["model", "model"]
  ];
  textFields.forEach(([target, source]) => {
    if (String(row[source] ?? "").trim()) next[target] = String(row[source]).trim();
  });
  if ((!String(vehicle.regNo ?? "").trim() || String(vehicle.regNo).trim() === "-") && row.regNo) {
    next.regNo = baseRegNo(row.regNo);
  }
  [
    ["loanAmount", "loanAmount"],
    ["emiAmount", "emiAmount"],
    ["interestRate", "interestRate"],
    ["tenure", "tenure"],
    ["paidEmi", "paidEmi"]
  ].forEach(([target, source]) => {
    if (toNumber(row[source]) > 0) next[target] = toNumber(row[source]);
  });
  [
    ["emiStart", "emiStart"],
    ["emiEnd", "emiEnd"],
    ["insuranceCompany", "policyCompany"],
    ["insurancePolicyNo", "policyNo"],
    ["insuranceStart", "policyStart"],
    ["insuranceExpiry", "policyEnd"],
    ["permitNo", "permitNo"],
    ["permitIssue", "permitIssue"],
    ["permitExpiry", "permitExpired"],
    ["permitType", "permitType"],
    ["nationalPermitExpiry", "nationalPermitExpired"],
    ["pucNo", "pucNo"],
    ["pucExpiry", "pucExpired"],
    ["fitnessExpiry", "fitnessExpired"]
  ].forEach(([target, source]) => {
    if (String(row[source] ?? "").trim() && String(row[source]).trim() !== "-") next[target] = String(row[source]).trim();
  });
  const schedule = Array.isArray(row.emiSchedule) ? row.emiSchedule.filter((entry) => entry?.dueDate && toNumber(entry.amount) > 0) : [];
  if (schedule.length > 0) {
    next.emiSchedule = schedule;
    next.emiHistory = schedule.filter((entry) => entry.status === "Paid").map((entry) => ({
      installment: Number(entry.installment),
      amount: toNumber(entry.amount),
      paidOn: entry.paidAt || entry.dueDate,
      reference: entry.reference || `EMI-${String(entry.installment).padStart(3, "0")}`,
      status: "Verified"
    }));
    next.emiStart = schedule[0].dueDate;
    next.emiEnd = schedule[schedule.length - 1].dueDate;
    next.emiAmount = toNumber(row.emiAmount) || toNumber(schedule[0].amount);
    next.tenure = toNumber(row.tenure) || schedule.length;
    next.paidEmi = toNumber(row.paidEmi) || schedule.filter((entry) => entry.status === "Paid").length;
  }
  const calculatedPrincipal = autoClosingPrincipal(row);
  if (row.bankClosingPrincipalUnknown === true) next.principal = "";
  else if (isClosedFinanceRow(row) || (row.scheduleParsed === "yes" && toNumber(row.tenure) > 0 && toNumber(row.paidEmi) >= toNumber(row.tenure) && toNumber(row.bankClosingPrincipal) <= 0)) next.principal = 0;
  else if (calculatedPrincipal > 0) next.principal = calculatedPrincipal;
  if (toNumber(row.loanAmount) > 0) next.loanAmount = toNumber(row.loanAmount);
  if (toNumber(row.bankClosingPrincipal) > 0) next.principal = toNumber(row.bankClosingPrincipal);
  return next;
}

function scheduleDateToIso(value) {
  const date = parseDisplayDate(value);
  if (!date) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}-${String(date.getDate()).padStart(2, "0")}`;
}

function nextPdfScheduleEntry(schedule) {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return (Array.isArray(schedule) ? schedule : [])
    .filter((entry) => entry?.dueDate && entry.status !== "Paid" && parseDisplayDate(entry.dueDate))
    .map((entry) => ({ entry, date: parseDisplayDate(entry.dueDate) }))
    .sort((left, right) => left.date.getTime() - right.date.getTime())
    .find(({ date }) => date >= today)
    || (Array.isArray(schedule) ? schedule
      .filter((entry) => entry?.dueDate && entry.status !== "Paid" && parseDisplayDate(entry.dueDate))
      .map((entry) => ({ entry, date: parseDisplayDate(entry.dueDate) }))
      .sort((left, right) => right.date.getTime() - left.date.getTime())[0] : null);
}

function buildPdfEmiDueTask(row, clientId, vehicleId, callerId = "") {
  const scheduleEntry = nextPdfScheduleEntry(row.emiSchedule)?.entry;
  const dueDate = scheduleDateToIso(scheduleEntry?.dueDate || row.emiStart || row.emiEnd);
  const amount = toNumber(scheduleEntry?.amount) || toNumber(row.emiAmount);
  if (!dueDate || amount <= 0) return null;
  const date = new Date(`${dueDate}T00:00:00`);
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const days = Math.round((date.getTime() - today.getTime()) / 86400000);
  return {
    // Match the backend monitor's deterministic key so PDF import and
    // scheduled monitoring upsert one EMI task instead of creating two.
    id: `auto-emi-${String(vehicleId).replace(/[^a-z0-9]/gi, "")}-${dueDate}`,
    clientId,
    vehicleId,
    type: "EMI",
    amount,
    dueDate,
    status: days < 0 ? "Overdue" : "Due",
    callerId,
    priority: days <= 7 ? "High" : days <= 15 ? "Medium" : "Low"
  };
}

function excelRowToDueTask(row, clientId, vehicleId, id) {
  const scheduleEntry = nextPdfScheduleEntry(row.emiSchedule)?.entry;
  const dueDate = scheduleDateToIso(scheduleEntry?.dueDate || row.emiStart) || row.emiStart || new Date().toISOString().slice(0, 10);
  return {
    id,
    clientId,
    vehicleId,
    type: "EMI",
    amount: toNumber(scheduleEntry?.amount) || toNumber(row.emiAmount),
    dueDate,
    status: "Due",
    callerId: "",
    priority: "Medium"
  };
}

async function parseClientExcelFile(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: false });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const table = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: false, defval: "" });
  const [headers, ...bodyRows] = table.map((row) => Array.isArray(row) ? row : []);
  if (!headers?.some((header) => String(header).toLowerCase().includes("regt"))) {
    throw new Error("Excel format did not match. Please upload a vehicle sheet with a Regt. No. column.");
  }
  return bodyRows
    .map((row) => mapClientExcelRow(row))
    .filter((row) => row.regNo || row.loanAccount || row.policyNo);
}

function validateClientExcelRows(rows, data, clientId) {
  const seenRegistrations = new Set();
  const seenLoanAccounts = new Set();
  const existingVehicles = data.vehicles.filter((vehicle) => vehicle.clientId === clientId);
  const existingRegistrations = new Set(existingVehicles.map((vehicle) => normalizeRegNo(baseRegNo(vehicle.regNo))));
  const existingLoanAccounts = new Set(existingVehicles.map((vehicle) => normalizeAgreement(vehicle.loanAccount)));
  const hasValue = (value) => String(value ?? "").trim() !== "" && String(value ?? "").trim() !== "-";
  const addDateIssue = (issues, label, value) => {
    if (hasValue(value) && !parseDisplayDate(value)) issues.push(`${label} is invalid`);
  };

  return rows.map((row) => {
    const issues = [];
    const registration = normalizeRegNo(baseRegNo(row.regNo));
    const isBody = isBodyRow(row);
    if (!hasValue(row.regNo)) {
      issues.push("Registration number is required");
    } else if (!/^[A-Z0-9][A-Z0-9 -]{3,23}$/.test(registration)) {
      issues.push("Registration number format is invalid");
    }
    const registrationKey = `${registration}:${isBody ? "body" : "vehicle"}`;
    if (registration && seenRegistrations.has(registrationKey)) issues.push("Duplicate registration in this file");
    if (registration) seenRegistrations.add(registrationKey);
    if (registration && !isBody && existingRegistrations.has(registration)) issues.push("Registration already exists for this client");
    if (!hasValue(row.owner)) issues.push("Owner is required");

    const financeStatus = String(row.financeStatus || "").trim().toUpperCase().replace(/\s+/g, "");
    const hasFinanceValues = [row.financier, row.loanAccount, row.loanAmount, row.emiAmount, row.tenure, row.paidEmi, row.emiStart, row.emiEnd]
      .some(hasValue);
    const isLoanFree = financeStatus === "FREE" || financeStatus === "LOANFREE" || financeStatus === "N/A";
    if (hasValue(row.financeStatus) && !["FIN", "FINANCED", "FREE", "LOANFREE", "N/A"].includes(financeStatus)) {
      issues.push("Finance status must be FIN or FREE");
    }
    if (hasFinanceValues && !isLoanFree) {
      if (!hasValue(row.financier)) issues.push("Financier is required for a financed row");
      if (!hasValue(row.loanAccount)) issues.push("Loan account is required for a financed row");
      if (hasValue(row.loanAccount) && !isValidAgreementValue(row.loanAccount)) issues.push("Loan account is invalid");
      if (toNumber(row.loanAmount) <= 0) issues.push("Loan amount must be greater than zero");
      if (toNumber(row.emiAmount) <= 0) issues.push("EMI amount must be greater than zero");
      if (!Number.isInteger(toNumber(row.tenure)) || toNumber(row.tenure) <= 0) issues.push("Tenure must be a positive whole number");
      if (!Number.isInteger(toNumber(row.paidEmi)) || toNumber(row.paidEmi) < 0 || toNumber(row.paidEmi) > toNumber(row.tenure)) {
        issues.push("Paid EMI must be between zero and tenure");
      }
      if (!hasValue(row.emiStart) || !hasValue(row.emiEnd)) issues.push("EMI start and end dates are required");
      addDateIssue(issues, "EMI start date", row.emiStart);
      addDateIssue(issues, "EMI end date", row.emiEnd);
      if (parseDisplayDate(row.emiStart) && parseDisplayDate(row.emiEnd) && parseDisplayDate(row.emiEnd) < parseDisplayDate(row.emiStart)) {
        issues.push("EMI end date cannot be before start date");
      }
    }
    const loanAccount = normalizeAgreement(row.loanAccount);
    if (loanAccount && seenLoanAccounts.has(loanAccount)) issues.push("Duplicate loan account in this file");
    if (loanAccount) seenLoanAccounts.add(loanAccount);
    if (loanAccount && existingLoanAccounts.has(loanAccount)) issues.push("Loan account already exists for this client");

    const hasPolicy = [row.policyCompany, row.policyNo, row.policyStart, row.policyEnd].some(hasValue);
    if (hasPolicy) {
      if (!hasValue(row.policyCompany)) issues.push("Insurance company is required");
      if (!hasValue(row.policyNo)) issues.push("Policy number is required");
      if (!hasValue(row.policyStart) || !hasValue(row.policyEnd)) issues.push("Policy start and expiry dates are required");
    }
    addDateIssue(issues, "Registration date", row.regDate);
    addDateIssue(issues, "Policy start date", row.policyStart);
    addDateIssue(issues, "Policy expiry date", row.policyEnd);
    addDateIssue(issues, "PUC expiry date", row.pucExpired);
    addDateIssue(issues, "Fitness expiry date", row.fitnessExpired);
    addDateIssue(issues, "Permit issue date", row.permitIssue);
    addDateIssue(issues, "Permit expiry date", row.permitExpired);
    addDateIssue(issues, "National permit expiry date", row.nationalPermitExpired);
    if (hasValue(row.permitNo) && !hasValue(row.permitExpired) && !hasValue(row.nationalPermitExpired)) issues.push("Permit expiry date is required when permit number is present");
    if (hasValue(row.pucExpired) && !hasValue(row.pucNo)) issues.push("PUC number is required when PUC expiry is present");

    return {
      ...row,
      validationStatus: issues.length === 0 ? "Valid" : "Invalid",
      validationIssues: issues
    };
  });
}

function validateStagedImportRows(rows) {
  const seenRegistrations = new Set();
  const seenLoanAccounts = new Set();
  return rows.map((row) => {
    if (row.issue === "Imported and saved") return row;
    const issues = [];
    const registration = normalizeRegNo(baseRegNo(row.regNo));
    const loanAccount = normalizeAgreement(row.loanAccount);
    if (!registration) issues.push("Registration number is required");
    else if (!/^[A-Z0-9][A-Z0-9 -]{3,23}$/.test(registration)) issues.push("Registration number format is invalid");
    if (!['Truck', 'Trailer'].includes(row.assetType)) issues.push("Asset type must be Truck or Trailer");
    if (!String(row.client || "").trim()) issues.push("Client is required");
    if (!String(row.lender || "").trim()) issues.push("Lender is required");
    if (!isValidAgreementValue(row.loanAccount)) issues.push("Loan account is invalid");
    if (registration && seenRegistrations.has(registration)) issues.push("Duplicate registration in this batch");
    if (registration) seenRegistrations.add(registration);
    if (loanAccount && seenLoanAccounts.has(loanAccount)) issues.push("Duplicate loan account in this batch");
    if (loanAccount) seenLoanAccounts.add(loanAccount);
    return { ...row, status: issues.length ? "Invalid" : "Valid", issue: issues.join("; ") || "Ready for final import" };
  });
}

function mapClientExcelRow(row) {
  const dateColumns = new Set([6, 14, 15, 19, 20, 22, 23, 25, 26, 28]);
  const value = (index) => {
    const raw = row[index] ?? "";
    if (dateColumns.has(index)) return formatExcelDate(raw);
    return String(raw).trim();
  };
  return {
    srNo: value(0),
    regNo: value(1),
    owner: value(2),
    manufacturer: value(3),
    model: value(4),
    yearOfMfg: value(5),
    regDate: value(6),
    financeStatus: value(7),
    financier: value(8),
    loanAccount: value(9),
    loanAmount: value(10),
    emiAmount: value(11),
    tenure: value(12),
    paidEmi: value(13),
    emiStart: value(14),
    emiEnd: value(15),
    closingPrincipal: value(16),
    policyCompany: value(17),
    policyNo: value(18),
    policyStart: value(19),
    policyEnd: value(20),
    pucNo: value(21),
    pucExpired: value(22),
    fitnessExpired: value(23),
    permitNo: value(24),
    permitIssue: value(25),
    permitExpired: value(26),
    permitType: value(27),
    nationalPermitExpired: value(28),
    remarks: value(29)
  };
}

function formatExcelDate(value) {
  const text = String(value ?? "").trim();
  const serial = Number(text);
  if (!text) return "";
  if (!Number.isFinite(serial) || serial < 25000) return formatDisplayDate(text);
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return formatDisplayDate(date.toISOString().slice(0, 10));
}

async function ocrImageFile(file) {
  const worker = await createFinanceOcrWorker();
  try {
    await worker.setParameters({ preserve_interword_spaces: "1", tessedit_pageseg_mode: "6" });
    const { data: psm6 } = await worker.recognize(file);
    await worker.setParameters({ tessedit_pageseg_mode: "11" });
    const { data: psm11 } = await worker.recognize(file);
    await worker.setParameters({ tessedit_pageseg_mode: "4" });
    const { data: psm4 } = await worker.recognize(file);
    return normalizePdfText([psm6.text, psm11.text, psm4.text].join("\n"));
  } finally {
    await worker.terminate();
  }
}

async function extractPdfTextWithOcr(file) {
  if (file.type?.startsWith("image/") || /\.(png|jpe?g|webp|bmp|tiff)$/i.test(file.name)) {
    return ocrImageFile(file);
  }
  let pdfDocument;
  try {
    const buffer = await file.arrayBuffer();
    pdfDocument = await pdfjsLib.getDocument({ data: buffer }).promise;
  } catch {
    return ocrImageFile(file);
  }
  // Step 1: Try native pdfjs text extraction first (works great for digital PDFs)
  const nativeText = await extractPdfNativeText(pdfDocument);
  const nativeNormalized = normalizePdfText(nativeText);
  // If native text contains enough finance-relevant content, use it directly.
  // Digital PDFs (like Cholamandalam) have embedded text that pdfjs extracts perfectly.
  const hasFinanceContent = /(?:agreement|loan|emi|installment|instalment|instl|repayment|principal|interest|outstanding|tenure|schedule|amortization)/i.test(nativeNormalized);
  const hasScheduleRows = /\b\d{1,3}\s+\d{1,2}[.\/-](?:\d{1,2}|[A-Za-z]{3,})[.\/-]\d{2,4}/m.test(nativeNormalized);
  const hasEnoughText = nativeNormalized.length > 200;
  const needsOcrSupplement = shouldSupplementPdfOcr(file.name, nativeNormalized);
  if (hasFinanceContent && hasEnoughText) {
    // Native text is usually more accurate, but TVS/scanned repayment layouts
    // need OCR as an extra source because their visible table can differ from
    // the PDF text layer.
    if (hasScheduleRows && !needsOcrSupplement) return nativeNormalized;
    try {
      const ocrText = await ocrPdfPages(pdfDocument);
      return normalizePdfText(nativeNormalized + "\n" + ocrText);
    } catch {
      return nativeNormalized;
    }
  }
  // Step 2: Fall back to OCR for scanned/image-based PDFs
  const ocrText = await ocrPdfPages(pdfDocument);
  // Combine native + OCR in case native had partial text
  const combined = nativeNormalized ? normalizePdfText(nativeNormalized + "\n" + ocrText) : normalizePdfText(ocrText);
  return combined;
}

function shouldSupplementPdfOcr(fileName, text) {
  const sourceName = String(fileName ?? "").toLowerCase();
  const source = String(text ?? "");
  return sourceName.includes("tvc") ||
    /TVS\s*CREDIT/i.test(source) ||
    (/Repayment\s+Schedule/i.test(source) && /Balance\s+Princip(?:al|le)/i.test(source) && /Monthly\s+Due/i.test(source));
}

function isTataFinanceDocument(value) {
  const source = String(value ?? "").replace(/[^A-Z0-9]+/gi, " ").trim();
  return /\bTATA\s+MOTORS\s+FINANCE(?:\s+SOLUTIONS)?\b/i.test(source) ||
    /\bTATA\s+CAPITAL\b/i.test(source) ||
    (/\bTATA\b/i.test(source) && /\bAMORTIZATION\s+TABLE\b/i.test(source) && /\bCONTRACT\s+NO\b/i.test(source));
}

function isAuSmallFinanceDocument(value) {
  const source = String(value ?? "").replace(/[^A-Z0-9]+/gi, " ").trim();
  return /\bA\s*U\s+(?:SMALL\s+)?FINANCE(?:\s+BANK)?\b/i.test(source) ||
    /\bA\s*U\s+BANK\b/i.test(source);
}

async function extractPdfNativeText(pdfDocument) {
  const pageCount = pdfDocument.numPages;
  const pageTexts = [];
  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    const page = await pdfDocument.getPage(pageNumber);
    const textContent = await page.getTextContent();

    // Deduplicate overlapping items (same text string at identical x, y coordinates)
    const uniqueItems = [];
    for (const item of textContent.items) {
      if (!item.str) continue;
      const x = item.transform[4];
      const y = item.transform[5];
      const text = item.str.trim();
      if (!text) continue;

      const isDuplicate = uniqueItems.some(
        (existing) =>
          Math.abs(existing.y - y) <= 1.5 &&
          Math.abs(existing.x - x) <= 2.5 &&
          existing.text.trim() === text
      );

      if (!isDuplicate) {
        uniqueItems.push({ x, y, text: item.str, width: item.width });
      }
    }

    // Group text items by their y-position to reconstruct lines (within 3.5px tolerance)
    const linesBucket = [];
    for (const item of uniqueItems) {
      let bucket = linesBucket.find((b) => Math.abs(b.y - item.y) <= 3.5);
      if (!bucket) {
        bucket = { y: item.y, items: [] };
        linesBucket.push(bucket);
      }
      bucket.items.push(item);
    }
    // Sort lines top-to-bottom (higher y = higher on page in PDF coords)
    linesBucket.sort((a, b) => b.y - a.y);
    const lines = linesBucket.map((bucket) => {
      const items = bucket.items.sort((a, b) => a.x - b.x);
      // Insert spaces between items that are far apart
      let line = "";
      for (let i = 0; i < items.length; i++) {
        if (i > 0) {
          const gap = items[i].x - (items[i - 1].x + (items[i - 1].width || 0));
          line += gap > 5 ? "  " : gap > 1 ? " " : "";
        }
        line += items[i].text;
      }
      return line.trim();
    }).filter(Boolean);
    pageTexts.push(lines.join("\n"));
  }
  return pageTexts.join("\n");
}

async function ocrPdfPages(pdfDocument) {
  const worker = await createFinanceOcrWorker();
  await worker.setParameters({
    preserve_interword_spaces: "1",
    tessedit_pageseg_mode: "6"
  });
  const pageCount = pdfDocument.numPages;
  const texts = [];
  try {
    for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
      const page = await pdfDocument.getPage(pageNumber);
      // Very large scans can create canvases that Tesseract/Leptonica cannot
      // read in production browsers. Keep enough resolution for table OCR,
      // but cap the rendered page size to a safe browser limit.
      const baseViewport = page.getViewport({ scale: 1 });
      const maxPageDimension = 3200;
      const renderScale = Math.min(4.5, maxPageDimension / Math.max(baseViewport.width, baseViewport.height));
      const viewport = page.getViewport({ scale: Math.max(renderScale, 1.5) });
      const canvas = window.document.createElement("canvas");
      canvas.width = Math.floor(viewport.width);
      canvas.height = Math.floor(viewport.height);
      const context = canvas.getContext("2d", { willReadFrequently: true });
      context.fillStyle = "#ffffff";
      context.fillRect(0, 0, canvas.width, canvas.height);
      await page.render({ canvasContext: context, viewport }).promise;
      const safeCanvas = limitOcrCanvas(canvas, maxPageDimension);
      const cleanCanvas = cloneCanvas(safeCanvas);
      const sharpenedCanvas = cloneCanvas(safeCanvas);
      sharpenCanvasForOcr(sharpenedCanvas);
      await worker.setParameters({ tessedit_pageseg_mode: "6" });
      const { data } = await recognizeOcrCanvas(worker, cleanCanvas);
      await worker.setParameters({ tessedit_pageseg_mode: "11" });
      const { data: sparseData } = await recognizeOcrCanvas(worker, cleanCanvas);
      // PSM 4 is better for scanned repayment tables with fixed columns;
      // PSM 6/11 remain useful for the document summary and labels.
      await worker.setParameters({ tessedit_pageseg_mode: "4" });
      const { data: columnData } = await recognizeOcrCanvas(worker, sharpenedCanvas);
      // TVS repayment tables use fine grid lines; a second table-oriented
      // pass recovers rows that PSM 6/11 often merges into the borders.
      await worker.setParameters({
        tessedit_pageseg_mode: "6",
        tessedit_char_whitelist: "0123456789/-.,%ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
      });
      const { data: numericTableData } = await recognizeOcrCanvas(worker, sharpenedCanvas);
      texts.push(data.text, sparseData.text, columnData.text, numericTableData.text);
    }
  } finally {
    await worker.terminate();
  }
  return normalizePdfText(texts.join("\n"));
}

async function recognizeOcrCanvas(worker, canvas) {
  try {
    return await worker.recognize(canvas);
  } catch (firstError) {
    // Some Hostinger/browser combinations fail while serializing a canvas
    // directly for the worker. A PNG data URL is a compatible fallback.
    try {
      return await worker.recognize(canvas.toDataURL("image/png"));
    } catch {
      throw firstError;
    }
  }
}

function limitOcrCanvas(canvas, maxDimension = 3200) {
  const largestSide = Math.max(canvas.width, canvas.height);
  if (!largestSide || largestSide <= maxDimension) return canvas;
  const scale = maxDimension / largestSide;
  const resized = window.document.createElement("canvas");
  resized.width = Math.max(1, Math.floor(canvas.width * scale));
  resized.height = Math.max(1, Math.floor(canvas.height * scale));
  const context = resized.getContext("2d", { willReadFrequently: true });
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, resized.width, resized.height);
  context.drawImage(canvas, 0, 0, resized.width, resized.height);
  return resized;
}

function cloneCanvas(canvas) {
  const clone = window.document.createElement("canvas");
  clone.width = canvas.width;
  clone.height = canvas.height;
  clone.getContext("2d", { willReadFrequently: true }).drawImage(canvas, 0, 0);
  return clone;
}

function sharpenCanvasForOcr(canvas) {
  const context = canvas.getContext("2d", { willReadFrequently: true });
  const image = context.getImageData(0, 0, canvas.width, canvas.height);
  for (let index = 0; index < image.data.length; index += 4) {
    const gray = image.data[index] * 0.299 + image.data[index + 1] * 0.587 + image.data[index + 2] * 0.114;
    const value = gray > 185 ? 255 : 0;
    image.data[index] = value;
    image.data[index + 1] = value;
    image.data[index + 2] = value;
  }
  context.putImageData(image, 0, 0);
}

function normalizePdfText(value) {
  return String(value ?? "")
    .replace(/[ \t]+/g, " ")
    .replace(/\r\n?/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function formatOcrDate(rawDate) {
  let cleaned = String(rawDate ?? "")
    .replace(/\s*([./-])\s*/g, "$1")
    .trim();
  const matchNoDash = cleaned.match(/^(\d{1,2})[./-]?(\d{1,2})[./-]?(\d{4})$/);
  if (matchNoDash) {
    const day = matchNoDash[1].padStart(2, "0");
    const month = matchNoDash[2].padStart(2, "0");
    const year = matchNoDash[3];
    return `${day}-${month}-${year}`;
  }
  return cleaned;
}

function findTataSummaryValues(text) {
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!isTataFinanceDocument(source) && !/Contract\s+No/i.test(source)) return {};

  const amountAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(?:INR|RS\\.?|₹)?\\s*([\\d,]+(?:\\.\\d{1,2})?)`, "i"));
    return match?.[1]?.replace(/,/g, "") ?? "";
  };
  const numberAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,4})`, "i"));
    return match?.[1] ?? "";
  };
  const dateAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4})`, "i"));
    return match?.[1] ?? "";
  };

  const customerMatch = source.match(/(?:Customer|Client|Client\s+Name)\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]{2,60}?)(?=\s+Contract|\s+c\/o|\s+Vehicle|\s+Finance|\s+Rate|$)/i);
  const regMatch = source.match(/(?:Vehicle\s+No\.?|Registration\s+No\.?)\s*[:\-]?\s*([A-Z]{2}\s*\d{1,2}\s*[A-Z]{1,3}\s*[- ]?\s*\d{3,4})/i);
  const contractMatch = source.match(/(?:Contract\s+No\.?|Agreement\s+No\.?)\s*[:\-]?\s*([A-Z0-9]{6,25})/i);
  const rateMatch = source.match(/(?:Rate\s*\([^)]*\)|IRR|Rate\s+of\s+Interest)[^0-9]{0,20}([0-9]{1,2}\.[0-9]{1,5})/i);

  const rawClosing = amountAfter("(?:Closing\\s+Principal|Principal\\s+Outstanding|Outstanding\\s+Principal|O\\/S\\s+Principal|Balance\\s+Outstanding)");
  const customerName = (customerMatch?.[1] ?? "").replace(/\s+(?:Contract|Agreement|Vehicle|Finance|Rate|Tenure|EMI)\b.*$/i, "").trim();

  return {
    owner: cleanPdfPersonName(customerName),
    loanAccount: contractMatch?.[1] ?? "",
    regNo: regMatch?.[1]?.replace(/\s+/g, " ").trim().toUpperCase() ?? "",
    loanAmount: amountAfter("(?:Finance\\s+Amount|Finance\\s+Amt|Amount\\s+Financed|Loan\\s+Amount)"),
    emiAmount: amountAfter("(?:EMI\\s+Amount|EMI\\s+Amt|Instl\\.?\\s*Amount)"),
    tenure: numberAfter("(?:Tenure\\s*\\/\\s*EMIs?|Tenure|No\\.?\\s*of\\s*Instls?\\.?|No\\.?\\s*of\\s*Installments?|No\\.?\\s*of\\s*EMIs)"),
    interestRate: rateMatch?.[1] ?? findPdfRate(source),
    emiStart: dateAfter("(?:First\\s+Instl\\s+Date|First\\s+Installment\\s+Date|Contract\\s+Date)"),
    emiEnd: dateAfter("(?:Maturity\\s+Date|Last\\s+Instalment\\s+date)"),
    bankClosingPrincipal: toNumber(rawClosing) > 0 ? rawClosing : ""
  };
}

function findAuSmallFinanceSummaryValues(text) {
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!isAuSmallFinanceDocument(source)) return {};

  const amountAfter = (label) => source.match(new RegExp(`${label}\\s*[:\\-]?\\s*(?:INR|RS\\.?|₹)?\\s*([\\d,]+(?:\\.\\d{1,2})?)`, "i"))?.[1]?.replace(/,/g, "") ?? "";
  const periodMatch = source.match(/EMI\s+Period\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s*(?:to|[-–])\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/i);
  const tenureMatch = source.match(/Tenure\s*(?:\/\s*EMIs?)?\s*[:\-]?\s*(\d{1,4})\s*(?:months?|EMIs?)?/i);
  const ownerMatch = source.match(/Customer\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]{1,60}?)(?=\s+Loan\s+Account|\s+Loan\s+Amount|$)/i);
  const registrationMatch = source.match(/Vehicle\s+No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9 -]{4,20}?)(?=\s+Tenure|\s+EMI|$)/i);

  return {
    owner: ownerMatch?.[1]?.trim() ?? "",
    regNo: registrationMatch?.[1]?.replace(/\s+/g, " ").trim().toUpperCase() ?? "",
    loanAmount: amountAfter("Loan\\s+Amount"),
    interestRate: source.match(/Rate\s+of\s+Interest\s*[:\-]?\s*([\d]+(?:\.\d{1,2})?)/i)?.[1] ?? "",
    tenure: tenureMatch?.[1] ?? "",
    emiAmount: amountAfter("EMI\\s+Amount"),
    emiStart: periodMatch?.[1] ?? "",
    emiEnd: periodMatch?.[2] ?? ""
  };
}

function parseAuScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!isAuSmallFinanceDocument(normalized)) return [];

  const parseLineValues = (installment, dueDate, valueText) => {
    const values = extractSignedMoneyLikeNumbers(valueText);
    if (values.length < 4) return null;
    const installmentAmount = values[0];
    const principalPaid = values[1];
    const interest = values[2];
    const closingPrincipal = values[3];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  };

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      return parseLineValues(match[1], match[2], match[3]);
    })
    .filter(Boolean);

  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseLineValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
}

function findCholaSummaryValues(text) {
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!/CHOLAMANDALAM/i.test(source)) return {};

  const amountAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(?:INR|RS\\.?|₹)?\\s*([\\d,]+(?:\\.\\d{1,2})?)`, "i"));
    return match?.[1]?.replace(/,/g, "") ?? "";
  };
  const numberAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,4})`, "i"));
    return match?.[1] ?? "";
  };
  const dateAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}|\\d{4}-\\d{1,2}-\\d{1,2}|\\d{1,2}[-\\s][A-Za-z]{3,}[-\\s]\\d{2,4})`, "i"));
    return match?.[1] ?? "";
  };

  const rateMatch = source.match(/(?:Internal\s+Rate\s+of\s*Return|IRR|Interest\s+Rate)[^0-9]{0,40}([0-9]{1,2}\.[0-9]{1,5})/i);

  const rawClosing = amountAfter("(?:Closing\\s+Principal|Principal\\s+Outstanding|Outstanding\\s+Principal|Balance\\s+Outstanding|Current\\s+POS)");

  const values = {
    loanAmount: amountAfter("(?:Amount\\s+Financed|Loan\\s+Amount|Asset\\s+Cost|Net\\s+Disbursement|Loan\\s+Sanctioned\\s+Amount|Finance\\s+Amount)"),
    emiAmount: amountAfter("(?:EMI\\s+Amount|EMI\\s+Amt|Instl\\.?\\s*Amt|Instl\\.?\\s*Amount|Installment\\s+Amount|Instalment\\s+Amount|Monthly\\s+Installment|Monthly\\s+Instalment)"),
    tenure: numberAfter("(?:Tenure|No\\.?\\s*of\\s*Instls?\\.?|No\\.?\\s*of\\s*Installments?|No\\.?\\s*of\\s*Instalments?|Repayable\\s+in\\s+Instalments?|Repayable\\s+in\\s+Installments?)"),
    interestRate: rateMatch ? rateMatch[1] : findPdfRate(source),
    emiStart: dateAfter("(?:First\\s+Insta?l?ment\\s+date|First\\s+Insta?l?ment\\s+Date|First\\s+EMI\\s+Date|EMI\\s+Start\\s+Date|Loan\\s+Authorization\\s+Date)"),
    emiEnd: dateAfter("(?:Last\\s+Insta?l?ment\\s+date|Last\\s+Insta?l?ment\\s+Date|Maturity\\s+Date|Last\\s+EMI\\s+Date)"),
    bankClosingPrincipal: toNumber(rawClosing) >= 1000 ? rawClosing : ""
  };
  return Object.fromEntries(Object.entries(values).filter(([, value]) => value));
}

function isCholaScheduleFormat(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  return /CHOLAMANDALAM/i.test(normalized) &&
    (/(?:Repayment\s+Schedule|Amortization|EMI\s+Schedule|Instl\.?\s*No)/i.test(normalized)) &&
    (/\b(?:Opening\s+Principal|Principal|Interest|Closing\s+Principal|Outstanding)\b/i.test(normalized));
}

function parseCholaScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!isCholaScheduleFormat(rawText) && !/CHOLAMANDALAM/i.test(normalized)) return [];
  // Cholamandalam repayment schedule columns:
  // Instl No | Due Date | Opening Principal | EMI Amount | Principal | Interest | Closing Principal
  // Some PDFs also have extra columns like Insurance/Service Tax.

  // Try line-by-line parsing first
  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      return parseCholaScheduleValues(match[1], match[2], match[3]);
    })
    .filter(Boolean);
  if (lineRows.length > 0) return lineRows;

  // Try flattened/segment parsing
  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  const segmentRows = starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseCholaScheduleValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
  return segmentRows;
}

function parseCholaScheduleValues(installment, dueDate, valueText) {
  const values = extractSignedMoneyLikeNumbers(valueText);
  if (values.length < 4) return null;

  // Try multiple column arrangements that Cholamandalam uses:
  // Format 1: Opening Principal | EMI Amount | Principal | Interest | Closing Principal
  // Format 2: EMI Amount | Principal | Interest | Closing Principal
  // Format 3: Opening Principal | EMI Amount | Principal | Interest | Service Tax/Insurance | Closing Principal

  if (values.length >= 5) {
    // Check if first value is Opening Principal (should be large, > EMI)
    const firstVal = toNumber(values[0]);
    const secondVal = toNumber(values[1]);
    if (firstVal > secondVal && firstVal > 1000) {
      // Format 1 or 3: Opening Principal is first
      const openingPrincipal = values[0];
      const installmentAmount = values[1];
      const principalPaid = values[2];
      const interest = values[3];
      const closingPrincipal = values.length >= 6 ? values[values.length - 1] : values[4];
      const serviceTax = values.length >= 6 ? values[4] : "";
      // Validate: principal + interest should approximately equal EMI
      const emiBreakup = toNumber(principalPaid) + toNumber(interest);
      const emiAmt = toNumber(installmentAmount);
      const diff = Math.abs(emiAmt - emiBreakup);
      if (emiAmt > 0 && diff <= Math.max(emiAmt * 0.15, 50)) {
        return makeScheduleRow({
          installment,
          dueDate,
          openingPrincipal,
          installmentAmount,
          principalPaid,
          interest,
          serviceTax,
          closingPrincipal,
          rate: deriveAnnualRate(interest, openingPrincipal)
        });
      }
    }
  }

  // Format 2: EMI Amount | Principal | Interest | Closing Principal (no opening principal)
  if (values.length >= 4) {
    const installmentAmount = values[0];
    const principalPaid = values[1];
    const interest = values[2];
    const closingPrincipal = values[values.length >= 5 ? 4 : 3];
    const serviceTax = values.length >= 5 ? values[3] : "";
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    const emiBreakup = toNumber(principalPaid) + toNumber(interest);
    const emiAmt = toNumber(installmentAmount);
    const diff = Math.abs(emiAmt - emiBreakup);
    if (emiAmt >= 100 && toNumber(principalPaid) > 0 && toNumber(interest) > 0 && diff <= Math.max(emiAmt * 0.15, 50)) {
      return makeScheduleRow({
        installment,
        dueDate,
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        serviceTax,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    }
  }

  // Fallback: try the last 4 values as EMI | Principal | Interest | Closing
  if (values.length >= 4) {
    const slice = values.slice(-4);
    const installmentAmount = slice[0];
    const principalPaid = slice[1];
    const interest = slice[2];
    const closingPrincipal = slice[3];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    if (toNumber(installmentAmount) >= 100 && toNumber(principalPaid) > 0 && toNumber(interest) > 0 && toNumber(closingPrincipal) >= 0) {
      return makeScheduleRow({
        installment,
        dueDate,
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    }
  }
  return null;
}

function parseBankPdfText(text, fileName = "") {
  const agreement = findPdfAgreement(text) || findAgreementFallback(text);
  const regNo = findPdfRegistration(text);
  const indostarSummaryValues = findIndostarSummaryValues(text);
  const tvsSummaryValues = findTvsCreditSummaryValues(text);
  const auSummaryValues = findAuSmallFinanceSummaryValues(text);
  const tataSummaryValues = findTataSummaryValues(text);
  const cholaSummaryValues = findCholaSummaryValues(text);
  const auRows = parseAuScheduleRows(text);
  const scheduleValues = findScheduleTableValues(text);
  const tataScheduleDetected = scheduleValues.scheduleParsed === "yes" && isTataFinanceDocument(text);
  const ashokLeylandSummaryValues = Object.keys(scheduleValues).length > 0 ? {} : findAshokLeylandSummaryValues(text);
  const tableValues = auRows.length > 0 ? summarizeScheduleRows(auRows, text) :
    Object.keys(scheduleValues).length > 0 ? scheduleValues :
    Object.keys(ashokLeylandSummaryValues).length > 0 ? ashokLeylandSummaryValues :
      findFinanceTableValues(text);
  // This scanned layout often loses its shaded headers. Its validated rows
  // are more reliable than loose label matches (which can read a date as ROI).
  const poonawallaValues = isPoonawallaScheduleFormat(text) && scheduleValues.scheduleParsed === "yes"
    ? scheduleValues : {};

  const exactLoanAmount = findPdfExactAmount(text, ["Amount Financed", "Loan Amount", "Financed Amount", "Total Loan Sanctioned", "Total Loan Disbursed", "Asset Cost", "Net Disbursement", "Sanctioned Amount", "Finance Amount"]);
  const explicitLoanAmount = findPdfAmount(text, ["Amount Financed", "Loan Amount", "Financed Amount", "Total Loan Sanctioned", "Total Loan Disbursed", "Asset Cost", "Net Disbursement", "Sanctioned Amount", "Finance Amount"], { positive: true, min: 1000 });
  const explicitEmiAmount = findPdfAmount(text, ["EMI Amount", "Installment Amount", "Instalment Amount", "Monthly Installment", "Monthly Instalment", "Repayment Amount", "EMI", "Instl Amt", "Instl Amount"], { positive: true, min: 100 });
  const explicitClosingPrincipal = findPdfAmount(text, ["Closing Principal", "Principal Outstanding", "Outstanding Principal", "Current POS", "POS", "Foreclosure Amount", "Foreclosure Value", "Closure Amount", "Amount to be paid", "Payable Amount", "Balance Outstanding"], { positive: true, min: 1000 });

  const indostarEmiAmount = isIndostarScheduleFormat(text) ? findIndostarEmiAmount(text) : "";
  const preferScheduleAmounts = scheduleValues.scheduleParsed === "yes" &&
    (isInstlOutstandingScheduleFormat(text) || isDateFreeRepaymentFormat(text));

  const finalLoanAmount = tvsSummaryValues.loanAmount || indostarSummaryValues.loanAmount || auSummaryValues.loanAmount || cholaSummaryValues.loanAmount || tataSummaryValues.loanAmount || (preferScheduleAmounts ? tableValues.loanAmount : explicitLoanAmount || exactLoanAmount || tableValues.loanAmount);
  const finalEmiAmount = tvsSummaryValues.emiAmount || indostarSummaryValues.emiAmount || auSummaryValues.emiAmount || cholaSummaryValues.emiAmount || tataSummaryValues.emiAmount || explicitEmiAmount || tableValues.emiAmount || indostarEmiAmount;
  const rawScheduleClosingPrincipal = scheduleValues.scheduleParsed === "yes" ? tableValues.bankClosingPrincipal : "";
  const finalClosingPrincipal = scheduleValues.scheduleParsed === "yes"
    ? plausibleBankClosingPrincipal(rawScheduleClosingPrincipal, finalLoanAmount)
      ? rawScheduleClosingPrincipal
      : ""
    : (cholaSummaryValues.bankClosingPrincipal || tataSummaryValues.bankClosingPrincipal || (auRows.length > 0 ? tableValues.bankClosingPrincipal : (explicitClosingPrincipal || tableValues.bankClosingPrincipal)));

  const tataFieldsTrusted = Boolean(
    tataSummaryValues.loanAmount &&
    (tataSummaryValues.emiAmount || scheduleValues.emiAmount) &&
    (tataSummaryValues.tenure || scheduleValues.tenure)
  );
  const cholaFieldsTrusted = Boolean(
    cholaSummaryValues.loanAmount &&
    (cholaSummaryValues.emiAmount || scheduleValues.emiAmount) &&
    (cholaSummaryValues.tenure || scheduleValues.tenure)
  );
  const anyFieldsTrusted = Boolean(finalLoanAmount || finalEmiAmount || scheduleValues.scheduleParsed === "yes");

  const loanStatus = findPdfValue(text, ["Loan Status", "Account Status", "Contract Status", "Status"]);

  return {
    owner: tataSummaryValues.owner || findPoonawallaOwner(text) || indostarSummaryValues.owner || tvsSummaryValues.owner || auSummaryValues.owner || findPdfOwner(text),
    financeStatus: agreement ? /closed|foreclos|fully\s+paid|settled|no\s+dues/i.test(loanStatus) ? "CLOSED" : "FIN" : "",
    loanAccount: indostarSummaryValues.agreementNumber || tvsSummaryValues.loanAccount || (/^(?:transfer|agreement|number|no)$/i.test(agreement) ? "" : agreement),
    regNo: indostarSummaryValues.regNo || tvsSummaryValues.regNo || auSummaryValues.regNo || tataSummaryValues.regNo || regNo,
    financier: findPdfFinancier(text, fileName),
    manufacturer: findPdfValue(text, ["Manufacturer", "Make", "Asset Make", "Vehicle Make", "Asset Description"]),
    model: tvsSummaryValues.model || findPdfValue(text, ["Asset Model", "Vehicle Model", "Model"]),
    loanStatus,
    loanAmount: finalLoanAmount,
    emiAmount: finalEmiAmount,
    tenure: tataSummaryValues.tenure || tvsSummaryValues.tenure || indostarSummaryValues.tenure || auSummaryValues.tenure || cholaSummaryValues.tenure || tableValues.tenure || findPdfValue(text, ["Tenure(In Months)", "Tenure In Months", "Tenure", "Period in Months", "Total Installment", "No of Installments", "No. of Installments", "No of Instalments", "No. of Instalments", "Number of EMIs", "No of Instls", "No. of Instls", "Repayable in Instalments", "Repayable in Installments"], { requireDigit: true, minLength: 1 }),
    paidEmi: scheduleValues.scheduleParsed === "yes"
      ? (toNumber(tableValues.bankClosingPrincipal) <= 0 && toNumber(tableValues.tenure) > 0
        ? tableValues.tenure
        : (tableValues.emiSchedule?.filter((entry) => entry.status === "Paid").length || tableValues.paidEmi))
      : (tableValues.paidEmi || findPdfValue(text, ["Paid EMI", "EMI Paid", "Installments Paid", "Instalments Paid", "No of EMI Paid", "No. of EMI Paid", "Instls Paid", "Instl Paid"])),
    interestRate: tvsSummaryValues.interestRate || indostarSummaryValues.interestRate || auSummaryValues.interestRate || cholaSummaryValues.interestRate || tataSummaryValues.interestRate || findPdfRate(text) || tableValues.interestRate,
    emiStart: tataScheduleDetected ? tableValues.emiStart : (tvsSummaryValues.emiStart || indostarSummaryValues.emiStart || auSummaryValues.emiStart || cholaSummaryValues.emiStart || tataSummaryValues.emiStart || tableValues.emiStart || findPdfDate(text, ["EMI Start Date", "Installment Start Date", "Instalment Start Date", "First EMI Date", "First Instalment date", "First Installment date", "First Instl Date", "Loan Authorization Date"])),
    emiEnd: tataScheduleDetected ? tableValues.emiEnd : (tvsSummaryValues.emiEnd || indostarSummaryValues.emiEnd || auSummaryValues.emiEnd || cholaSummaryValues.emiEnd || tataSummaryValues.emiEnd || tableValues.emiEnd || findPdfDate(text, ["EMI End Date", "Maturity Date", "Last EMI Date", "Last Instalment date", "Last Installment date", "Last Instl Date"])),
    bankClosingPrincipal: finalClosingPrincipal,
    scheduleParsed: auRows.length > 0 || scheduleValues.scheduleParsed ? "yes" : "",
    emiSchedule: Array.isArray(tableValues.emiSchedule) ? tableValues.emiSchedule : [],
    ...poonawallaValues,
    _pdfFieldsTrusted: tataFieldsTrusted || cholaFieldsTrusted || anyFieldsTrusted,
    remarks: "Imported from bank PDF"
  };
}

function findPdfRegistration(text) {
  const labelledReg = String(text ?? "").match(/\b(?:Registration\s*Number|Registration\s*No|Regn\s*Number|Regn\s*No|Vehicle\s*Number|Vehicle\s*No|Vehicle\s*Reg\s*No)\s*[:\-]?\s*([A-Z]{2}\s*\d{1,5}\s*[A-Z]{1,4}\s*[- ]?\s*\d{3,8})/i);
  if (labelledReg) return labelledReg[1].replace(/\s*-\s*/g, " ").replace(/\s+/g, " ").trim().toUpperCase();
  const looseReg = String(text ?? "").match(/\b([A-Z]{2}\s*\d{1,5}\s*[A-Z]{1,4}\s*[- ]?\s*\d{3,8})\b/i);
  return looseReg ? looseReg[1].replace(/\s*-\s*/g, " ").replace(/\s+/g, " ").trim().toUpperCase() : "";
}

function findPdfAgreement(text) {
  const source = String(text ?? "").replace(/\s+/g, " ");
  const patterns = [
    /Repayment\s+Schedule\s+For\s+Agreement\s+No\.?\s*[:\-]{0,3}\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Repayment\s+Schedule\s+For\s+([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Agreement\s+Number\s*[:\-]?\s*(?:Date\s*:\s*)?([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /(?:Ref\s*:\s*Your\s+)?Agreement\s+(?:Number|No)\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Agmt\s+No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Loan\s+Account\s+(?:Number|No)\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Account\s+(?:Number|No)\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Contract\s+(?:Number|No)\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i,
    /Application\s+No\.?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{8,})/i
  ];
  for (const pattern of patterns) {
    const match = source.match(pattern);
    const value = match?.[1]?.trim() ?? "";
    if (isValidAgreementValue(value) && !/^(?:AGREEMENT|NUMBER|NO|ACCOUNT|CONTRACT)$/i.test(value)) return value;
  }
  const fallback = findPdfValue(text, ["Agreement Number", "AgreementNumber", "Agreement No", "AgreementNo", "Contract Number", "Contract No", "ContractNo", "Account Number", "Account No", "AccountNo", "Loan Account Number", "LoanAccountNumber", "LAN", "Loan No", "LoanNo"], { requireDigit: true, minLength: 6 });
  return /^(?:agreement|number|no|account|contract)$/i.test(fallback) ? "" : fallback;
}

function findPdfOwner(text) {
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (/HDFC\s+BANK/i.test(source)) {
    const hdfcOwner = cleanPdfPersonName(source.match(/\b(MR\.?\s+[A-Za-z]+(?:\s+[A-Za-z]+){1,3}?)\s+(?=A-\d|Customer|Loan)/i)?.[1] ?? "");
    if (hdfcOwner) return hdfcOwner;
  }
  if (/ICICI\s+BANK/i.test(source)) {
    const iciciOwner = cleanPdfPersonName(source.match(/Repayment\s+Schedule\s+(Mr\.?\s+[A-Za-z]+(?:\s+[A-Za-z]+){2,5})(?=\s+Loan\s+Account)/i)?.[1] ?? "");
    if (iciciOwner) return iciciOwner;
  }
  const bankSpecific = [
    [/HDFC\s+BANK/i, /\b(MR\.?\s+[A-Z][A-Za-z .'-]{4,80}?)\s+(?=Customer\s+|Loan\s+Type)/i],
    [/ICICI\s+BANK/i, /Repayment\s+Schedule\s+(MR\.?\s+[A-Za-z][A-Za-z .'-]{2,80}?)\s+Loan\s+Account/i],
    [/SUNDARAM\s+FINANCE/i, /\b(MR\s+[A-Z][A-Z .'-]{4,80}?)\s+S\/O/i],
    [/CHOLAMANDALAM/i, /Dear\s+Mr\.?\/Mrs\.?\s+(?:Dear\s+Mr\.?\/Mrs\.?\s+)?(.+?)(?=\s+Ref:)/i],
    [/CHOLAMANDALAM/i, /Customer\s+Name\s*[:\-]?\s*(.+?)(?=\s+(?:Agreement|Contract|Loan|Account|Product|CIF|Customer\s+Code|$))/i],
    [/CHOLAMANDALAM/i, /Borrower(?:'s)?\s+Name\s*[:\-]?\s*(.+?)(?=\s+(?:Co-?Borrower|Agreement|Contract|Loan|$))/i]
  ];
  for (const [bankPattern, ownerPattern] of bankSpecific) {
    if (!bankPattern.test(source)) continue;
    const specific = cleanPdfPersonName(source.match(ownerPattern)?.[1] ?? "");
    if (specific) return specific;
  }
  const patterns = [
    /Repayment\s+Schedule\s+(Mr\.?\s+[A-Za-z]+(?:\s+[A-Za-z]+){2,5})(?=\s+Loan\s+Account)/i,
    /Customer\s+Name\s*[:\-]?\s*(.+?)(?=\s+(?:CIF\s+Number|Registered\s+Email|Contract\s+Date|Processing\s+Fee|Proposal\s+No|Customer\s+Code)|$)/i,
    /Borrower\s+Name\s*[:\-]?\s*(.+?)(?=\s+(?:Co-?Borrower|Borrower\s+Communication|\d{1,2}[./-]\d{1,2}[./-]\d{2,4})|$)/i,
    /Name\s+of\s+the\s+Customer\s+(.+?)(?=\s+Customer\s+Code|\s+UCIC|$)/i,
    /Dear\s+Mr\.?\/Mrs\.?\s+(?:Dear\s+Mr\.?\/Mrs\.?\s+)?(.+?)(?=\s+Ref:|$)/i,
    /Customer\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{2,80}?)(?=\s+Loan\s+Type)/i,
    /Customer\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{2,80}?)(?=\s+Customer\s+)/i,
    /Customer\s+([A-Z][A-Z .&'-]{3,80}?)(?=\s+Product\s+)/i,
    /Repayment\s+Schedule\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{2,80}?)(?=\s+Loan\s+Account)/i,
    /Client\s+Name\s*[:\-]?\s*(.+?)(?=\s+Contract\s+Date|$)/i,
    /Details\s+of\s+Interest\s+and\s+Principal\s+Breakup\s+[A-Z0-9]+\s+Contract\s+No\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{4,}?)(?=\s+S\/O)/i,
    /Contract\s+No\.?\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{4,}?)(?=\s+S\/O)/i,
    /\b([A-Z][A-Z]{3,})\s+C\s+\d{1,2}\s+[A-Z]/i,
    /Contract\s+No\.?\s*[:\-]?\s*[A-Z0-9]+\s+(.+?)(?=\s+S\/O|\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|$)/i,
    /Repayment\s+Schedule\s+((?:MR|MS|MRS)\.?\s+[A-Z][A-Z .'-]{2,80}?)(?=\s+Loan\s+Account(?:\s+Number)?)/i,
    /(?:^|\s)(MR\.?\s+[A-Z][A-Z .'-]{4,})(?=\s+Loan\s+Sanctioned\s+Amount)/i
  ];
  for (const pattern of patterns) {
    const match = source.match(pattern);
    const value = cleanPdfPersonName(match?.[1] ?? "");
    if (value && !/^(?:ID|CARE|NAME|CUSTOMER)$/i.test(value)) return value;
  }
  return "";
}

function cleanPdfPersonName(value) {
  let clean = String(value ?? "")
    .replace(/\b(?:CIF|UCIC|Customer\s+Code|Customer\s+No|Processing\s+Fee|Agreement\s+No|Proposal\s+No)\b.*$/i, "")
    .replace(/\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}.*$/, "")
    .replace(/\s+/g, " ")
    .trim();
  const tokens = clean.split(" ");
  if (tokens.length >= 4 && tokens.length % 2 === 0) {
    const half = tokens.length / 2;
    if (tokens.slice(0, half).join(" ").toUpperCase() === tokens.slice(half).join(" ").toUpperCase()) {
      clean = tokens.slice(0, half).join(" ");
    }
  }
  return clean;
}

function findPdfFinancier(text, fileName = "") {
  const upperText = String(text ?? "").toUpperCase();
  // The bank named inside IndoStar's repayment rows is the payment bank.
  if (isIndostarRepaymentLayout(text)) return "INDOSTAR";
  // Normalize legal-name and OCR-spacing variants to the names used by the
  // finance table and database filters.
  if (isTataFinanceDocument(text)) return "TATA MOTORS FINANCE SOLUTIONS LTD";
  if (isAuSmallFinanceDocument(text)) return "AU SMALL FINANCE";
  const knownFinanciers = [
    "BAJAJ FINANCE",
    "MAHINDRA AND MAHINDRA FINANCIAL SERVICES LIMITED",
    "MAHINDRA FINANCE",
    "POONAWALLA FINCORP LTD",
    "POONAWALLA FINCORP LIMITED",
    "POONAWALLA FINCORP",
    "SUNDARAM FINANCE",
    "BANK OF BARODA",
    "BANDHAN BANK",
    "TATA MOTORS FINANCE SOLUTIONS LTD",
    "TATA MOTORS FINANCE SOLUTIONS",
    "TATA MOTORS FINANCE",
    "TATA CAPITAL",
    "AXIS BANK",
    "AXIS FINANCE",
    "HDFC BANK",
    "HDB FINANCIAL",
    "ICICI BANK",
    "ICICI LOMBARD",
    "KOTAK MAHINDRA",
    "CHOLAMANDALAM",
    "MAHINDRA FINANCE",
    "INDUSIND BANK",
    "YES BANK",
    "AU SMALL FINANCE",
    "SHRIRAM FINANCE"
  ];
  const knownMatch = knownFinanciers.find((name) => upperText.includes(name));
  if (knownMatch) return knownMatch;

  const financierAliases = [
    { pattern: /\b(?:INDUS\s*IND|INDUSIND|INDU[S5]IND|INDU\s*HAND|INUHAND|INDUHAND)\b/i, name: "INDUSIND BANK" },
    { pattern: /\bASHOK\s*LEYLAND\b/i, name: "ASHOK LEYLAND" }
  ];
  const aliasMatch = financierAliases.find((item) => item.pattern.test(upperText));
  if (aliasMatch) return aliasMatch.name;

  const sourceName = String(fileName ?? "").toLowerCase();
  if (sourceName.includes("indushnd")) return "INDUSIND BANK";
  if (sourceName.includes("indostar")) return "INDOSTAR";
  if (sourceName.includes("tvc")) return "TVS CREDIT SERVICES LIMITED";

  const headerCompanyMatch = String(text ?? "").match(/^\s*([A-Z][A-Za-z .&-]{4,80}(?:Limited|Ltd|Finance|Fincorp|Bank))\s*(?:\n|Note\b|!)/im);
  if (headerCompanyMatch) return headerCompanyMatch[1].replace(/\s+/g, " ").trim().toUpperCase();

  const bankLineMatch = String(text ?? "").match(/\bBank\s*:\s*\d{1,8}\s+([A-Za-z][A-Za-z .&-]{2,60}?)(?=\s{2,}|\n|Branch\b|$)/i);
  if (bankLineMatch) return bankLineMatch[1].replace(/\s+/g, " ").trim().toUpperCase();

  const labelled = findPdfValue(text, [
    "Financier's Name",
    "Financier Name",
    "Lender Name",
    "Financier",
    "Name of Financier",
    "Finance Company",
    "NBFC Name"
  ], { minLength: 3 });
  if (labelled && !/^(bank|name|n\/?a|na)$/i.test(labelled)) return labelled;
  return "";
}

function findProfileAgreementInText(text, rows) {
  const normalizedText = normalizeAgreement(text);
  const looseText = normalizeAgreementLoose(text);
  return (rows ?? [])
    .map((row) => row.loanAccount)
    .filter(Boolean)
    .find((loanAccount) => {
      const exact = normalizeAgreement(loanAccount);
      const loose = normalizeAgreementLoose(loanAccount);
      if (!exact || exact.length < 6) return false;
      return normalizedText.includes(exact) ||
        looseText.includes(loose) ||
        (exact.length >= 8 && normalizedText.includes(exact.slice(-8))) ||
        (loose.length >= 8 && looseText.includes(loose.slice(-8)));
    }) ?? "";
}

function findFinanceTableValues(text) {
  if (isIndostarScheduleFormat(text)) return {};
  const lines = String(text ?? "")
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .filter(Boolean);
  for (let index = 0; index < lines.length; index += 1) {
    const header = lines[index];
    const normalizedHeader = header.toLowerCase();
    const hasEmi = /\bemi\b/.test(normalizedHeader);
    const hasTenure = /tenure/.test(normalizedHeader);
    const hasPaid = /paid\s*emi|emi\s*paid|paid/.test(normalizedHeader);
    const hasClosing = /closing|principal|outstanding|foreclosure|pos/.test(normalizedHeader);
    if (!(hasEmi && hasTenure && hasPaid && hasClosing)) continue;

    const scan = [
      header,
      lines[index + 1],
      lines[index + 2],
      `${header} ${lines[index + 1] ?? ""}`
    ].filter(Boolean);
    for (const line of scan) {
      const values = extractMoneyLikeNumbers(line).filter((value) => toNumber(value) > 0);
      if (values.length >= 4) {
        return {
          emiAmount: values[0],
          tenure: values[1],
          paidEmi: values[2],
          bankClosingPrincipal: values[3]
        };
      }
    }
    const nearby = extractMoneyLikeNumbers([lines[index + 1], lines[index + 2], lines[index + 3]].filter(Boolean).join(" ")).filter((value) => toNumber(value) > 0);
    if (nearby.length >= 4) {
      return {
        emiAmount: nearby[0],
        tenure: nearby[1],
        paidEmi: nearby[2],
        bankClosingPrincipal: nearby[3]
      };
    }
  }
  return {};
}

function findTvsCreditSummaryValues(text) {
  if (isIndostarRepaymentLayout(text)) return {};
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!/TVS\s*CREDIT|TVS|Balance\s+Principle|Monthly\s+Due|Repayment\s+Schedule/i.test(source)) return {};

  const amountAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(?:INR|RS\\.?|₹)?\\s*([\\d,]+(?:\\.\\d{1,2})?)`, "i"));
    return match?.[1]?.replace(/,/g, "") ?? "";
  };
  const numberAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,4})`, "i"));
    return match?.[1] ?? "";
  };
  const dateAfter = (pattern) => {
    const match = source.match(new RegExp(`${pattern}\\s*(?:\\([^)]*\\))?\\s*[:\\-]?\\s*(\\d{1,2}[./-]?\\d{1,2}[./-]?\\d{2,4}|\\d{4}-\\d{1,2}-\\d{1,2}|\\d{1,2}[-\\s][A-Za-z]{3,}[-\\s]\\d{2,4})`, "i"));
    return match ? formatDisplayDate(match[1]) : "";
  };

  const agreementMatch = source.match(/\bAgreement\s+(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i) ||
    source.match(/\bAgmt\s+No\.?(?:\s+Mobile\s+No\.?\s+Customer\s+No\.?)?\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i) ||
    source.match(/\b(?:Contract|Loan\s+Account)\s*(?:No\.?|Number)\s*[:\-]?\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i);
  const ownerMatch = source.match(/(?:Customer|Borrower|Client)\s*Name?\s*[:\-]?\s*([A-Za-z][A-Za-z .'-]{2,60}?)(?=\s+Proposal|\s+Father|\s+Agreement|\s+Asset|\s+Contract|$)/i);
  const modelMatch = source.match(/(?:Asset|Vehicle)\s*Model?\s*[:\-]?\s*([A-Za-z0-9][A-Za-z0-9 .'-]{2,40}?)(?=\s+Customer|\s+Applicable|\s+Amount|$)/i);
  const rateMatch = source.match(/Rate\s*of\s*Interest(?:\s*\(Amort\s*IRR\)\s*P\.?A\.?)?[^0-9]{0,30}([0-9]{1,2}(?:\.[0-9]{1,4})?)/i) ||
    source.match(/(?:Amort\s*IRR|IRR|APR)[^0-9]{0,30}([0-9]{1,2}(?:\.[0-9]{1,4})?)/i);
  const scheduleRows = parseTvsScheduleRows(text);
  const scheduleEmi = mostCommonAmount(scheduleRows.map((row) => row.installmentAmount).filter((item) => toNumber(item) > 0));
  const scheduleTenure = scheduleRows.length ? String(Math.max(...scheduleRows.map((row) => Number(row.installment) || 0))) : "";
  const scheduleLoanAmount = scheduleRows[0]?.openingPrincipal ?? "";

  const amountFinanced = amountAfter("Amount\\s+Financed");
  const totalLoanAmount = amountAfter("Total\\s+Loan\\s+Amount");
  const loanAmt = amountFinanced || totalLoanAmount || amountAfter("(?:Finance\\s+Amount|Loan\\s+Amount)");
  const emiStart = dateAfter("(?:First\\s+Insta?l?ment\\s+Date|First\\s+EMI\\s+Date)");
  const emiEnd = dateAfter("(?:Last\\s+Insta?l?ment\\s+Date|Maturity\\s+Date)");

  const tvsFallback = findTvsOcrSummaryValues(text, scheduleRows);
  const safeOwner = tvsFallback.owner || ((ownerMatch?.[1]?.trim() && /^[A-Z][A-Z ]+$/.test(ownerMatch[1].trim()) && !/TVS\s+CREDIT|SERVICES|LIMITED/i.test(ownerMatch[1]) ? ownerMatch[1].trim() : ""));
  return {
    loanAccount: tvsFallback.loanAccount || (agreementMatch?.[1] && /\d/.test(agreementMatch[1]) ? agreementMatch[1] : ""),
    owner: safeOwner,
    regNo: tvsFallback.regNo,
    model: modelMatch?.[1]?.trim() || tvsFallback.model,
    loanAmount: loanAmt || tvsFallback.loanAmount || scheduleLoanAmount,
    emiAmount: scheduleEmi,
    tenure: numberAfter("(?:No\\.?\\s*of\\s*Insta?l?ments?|Tenure|No\\.?\\s*of\\s*EMIs)") || tvsFallback.tenure || scheduleTenure,
    interestRate: rateMatch?.[1] || tvsFallback.interestRate,
    emiStart: emiStart || tvsFallback.emiStart || (scheduleRows[0]?.dueDate ? formatDisplayDate(scheduleRows[0].dueDate) : ""),
    emiEnd: emiEnd || tvsFallback.emiEnd || (scheduleRows[scheduleRows.length - 1]?.dueDate ? formatDisplayDate(scheduleRows[scheduleRows.length - 1].dueDate) : "")
  };
}

function findTvsOcrSummaryValues(source, scheduleRows = []) {
  const value = String(source ?? "");
  if (!/TVS\s*CREDIT|TVSCREDIT/i.test(value)) return {};
  const agreement = [...value.matchAll(/\b((?:[A-Z]{2,8}|\d{4})[A-Z0-9]{6,20})\b/gi)]
    .map((match) => match[1]).find((candidate) => /\d/.test(candidate) && candidate.length >= 10) ?? "";
  const registration = value.match(/\b(GJ\s*\d{1,5}\s*[A-Z]{1,4}\s*\d{3,8})\b/i)?.[1]?.replace(/\s+/g, "").toUpperCase() ?? "";
  const dates = [...value.matchAll(/\b(\d{1,2}[-/]\s*[A-Za-z]{3,}|\d{1,2}[-/]\d{1,2})[-/]?(\d{2,4})\b/gi)]
    .map((match) => match[0].replace(/\s+/g, ""));
  const dateValues = dates.filter((date) => !/^(?:01|4)-01-2024$/i.test(date));
  const ownerLine = value.split(/\n+/).find((line) => /\d{4}[A-Z]{2}\d{6,}/.test(line));
  const owner = ownerLine?.match(/([A-Z]{3,}(?:\s+[A-Z]{3,}){2,4})\s+.*\d{4}[A-Z]{2}\d{6,}/)?.[1]?.trim() ??
    value.match(/(?:Customer|Borrower|ULES\]|llH)\s*(?:Name)?[^A-Za-z]{0,12}([A-Z][A-Z ]{8,60}?)(?=\s+(?:Agreement|3026|[A-Z]{2}\d{6,}))/)?.[1]?.trim() ??
    value.match(/\b([A-Z]{3,}(?:\s+[A-Z]{3,}){2,4})\s+.*?(?=(?:\d{4}[A-Z]{2}|[A-Z]{2,8}\d{6,}))/)?.[1]?.trim() ?? "";
  const amounts = [...value.matchAll(/\b(\d{1,3}(?:,?\d{2,3})+|\d{6,8})\b/g)].map((match) => match[1].replace(/,/g, "")).map(Number);
  const loanAmount = amounts.find((amount) => amount >= 500000 && amount <= 10000000 && amount % 1000 === 0)?.toString() ?? "";
  const tenure = value.match(/\b(\d{1,3})\b(?=\s*(?:EMIs?|instalments?|installments?))/i)?.[1] ??
    Math.max(...[...value.matchAll(/\b(\d{1,3})\b/g)].map((match) => Number(match[1])).filter((amount) => amount >= 12 && amount <= 60), 0).toString();
  const rateMatch = value.match(/(?:APR|IRR|Interest)[^\d]{0,25}(\d{1,2}\.\d{2})\s*%/i);
  const model = value.match(/\b(TATA\s+[A-Z0-9 -]{3,30})\b/i)?.[1]?.trim() ?? "";
  return {
    agreement, loanAccount: agreement, owner, regNo: registration, model,
    loanAmount, tenure, interestRate: rateMatch?.[1] ?? "",
    emiStart: dateValues.find((date) => /AUG|[/-]\d{2}/i.test(date)) ? formatDisplayDate(dateValues[0]) : "",
    emiEnd: dateValues.length > 1 ? formatDisplayDate(dateValues[1]) : ""
  };
}

function parseTvsScheduleRows(text) {
  if (isIndostarRepaymentLayout(text)) return [];
  const rawText = String(text ?? "");
  const scheduleStart = rawText.search(/(?:Repayment\s+Schedule\s*:\s*)?S\.?\s*No\.?\s+Due\s+Date\s+EMI|Repayment\s+Schedule\s*:/i);
  const scheduleText = scheduleStart >= 0 ? rawText.slice(scheduleStart) : rawText;
  const normalized = scheduleText.replace(/\s+/g, " ");
  if (!/TVS/i.test(normalized) && !/Balance\s+Principle/i.test(normalized) && !/Vehicle\s+Insurance/i.test(normalized) && !/Monthly\s+Due/i.test(normalized) && !/Repayment\s+Schedule/i.test(normalized)) return [];

  const parseTvsRowValues = (installment, rawDueDate, valueText) => {
    let values = extractSignedMoneyLikeNumbers(valueText);
    if (values.length < 4) return null;
    let dueDate = formatOcrDate(rawDueDate);
    let installmentAmount = values[0];
    let principalPaid = values[1];
    let interest = values[2];
    let closingPrincipal = values[3];
    let serviceTax = values.length >= 5 ? values[4] : "";

    let emi = toNumber(installmentAmount);
    let prin = toNumber(principalPaid);
    let intr = toNumber(interest);

    // Self-correct OCR missing decimal point or digit glitch
    if (emi > 100) {
      if (intr > 0 && intr < emi && (prin <= 0 || prin > emi || Math.abs(emi - (prin + intr)) > 50)) {
        prin = Math.max(0, emi - intr);
        principalPaid = String(prin);
      } else if (prin > 0 && prin < emi && (intr <= 0 || intr > emi || Math.abs(emi - (prin + intr)) > 50)) {
        intr = Math.max(0, emi - prin);
        interest = String(intr);
      }
    }

    if (toNumber(installmentAmount) <= 0 || toNumber(principalPaid) < 0) return null;
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  };

  const strictRows = [...normalized.matchAll(/(?:^|\s)(\d{1,3})\s+(\d{1,2}\s*[.\/-]\s*\d{1,2}\s*[.\/-]\s*\d{2,4})\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)/gi)]
    .map((match) => parseTvsRowValues(match[1], match[2], `${match[3]} ${match[4]} ${match[5]} ${match[6]} ${match[7]} ${match[8]}`))
    .filter(Boolean);

  if (strictRows.length > 0) return strictRows;

  const lineRows = scheduleText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(?:0*(\d{1,3})|(?:\/|\|)?\s*(\d{1,3}))\s+(?:[\[\{]\s*)?(\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const inst = match[1] || match[2];
      if (!inst || Number(inst) <= 0) return null;
      return parseTvsRowValues(inst, match[3], match[4]);
    })
    .filter(Boolean);

  if (lineRows.length > 0) return lineRows;

  // Some TVS scans lose the left S.No column but retain the date and all
  // financial columns. Assign those rows in document order.
  const dateOnlyRows = scheduleText.split(/\n+/).map((line) => {
    const match = line.replace(/\s+/g, " ").trim().match(/^(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
    if (!match) return null;
    const row = parseTvsRowValues("", match[1], match[2]);
    return row ? { ...row, installment: 0 } : null;
  }).filter(Boolean);
  if (dateOnlyRows.length > 0) return dateOnlyRows.map((row, index) => ({ ...row, installment: index + 1 }));

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(?:[\[\{]\s*)?(\d{1,2}\s*[./-]\s*\d{1,2}\s*[./-]\s*\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseTvsRowValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
}

function parseAuSmallFinanceScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!isAuSmallFinanceDocument(normalized)) return [];
  if (!/EMI\s+No\.?/i.test(normalized) || !/Closing\s+Principal/i.test(normalized)) return [];
  const parseRow = (installment, dueDate, valueText) => {
    const values = extractSignedMoneyLikeNumbers(valueText);
    if (values.length < 4) return null;
    const [installmentAmount, principalPaid, interest, closingPrincipal] = values;
    if (toNumber(installmentAmount) <= 0 || toNumber(principalPaid) < 0 || toNumber(interest) < 0 || toNumber(closingPrincipal) < 0) return null;
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  };
  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      return parseRow(match[1], match[2], match[3]);
    })
    .filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseRow(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
}

function parseIndusIndScheduleRows(text) {
  const rawText = String(text ?? "");
  if (!/TENTATIVE\s+LOAN\s+REPAYMENT|AMORTIZATION\s+SCHEDULE/i.test(rawText) || !/Outstanding\s+Principal/i.test(rawText)) return [];
  return rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const values = extractSignedMoneyLikeNumbers(match[3]);
      if (values.length < 4) return null;
      const [installmentAmount, interest, principalPaid, closingPrincipal] = values;
      if (toNumber(installmentAmount) < 0 || toNumber(principalPaid) < 0 || toNumber(interest) < 0 || toNumber(closingPrincipal) < 0) return null;
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({
        installment: match[1],
        dueDate: match[2],
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    })
    .filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0);
}

function parseBandhanScheduleRows(text) {
  const rawText = String(text ?? "");
  if (!/BANDHAN\s+BANK/i.test(rawText) || !/Installment\s+Start\s+Date/i.test(rawText)) return [];
  return rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})\s+(\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const values = extractSignedMoneyLikeNumbers(match[4]);
      if (values.length < 7) return null;
      const [rate, , principalPaid, interest, charge, installmentAmount, closingPrincipal] = values;
      // The first Bandhan row can be a valid BPI row with zero installment.
      if (toNumber(installmentAmount) < 0 || toNumber(principalPaid) < 0 || toNumber(interest) < 0 || toNumber(closingPrincipal) < 0) return null;
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({
        installment: match[1],
        dueDate: match[3],
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        serviceTax: charge,
        closingPrincipal,
        rate
      });
    })
    .filter((row) => row?.installment > 0 && toNumber(row.closingPrincipal) >= 0);
}

function parseTataScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!isTataFinanceDocument(normalized) && !/Balance\s+Outstanding/i.test(normalized) && !/O\/?S\s+Principal/i.test(normalized)) return [];

  const parseTataRowValues = (installment, dueDate, valueText) => {
    let values = extractSignedMoneyLikeNumbers(valueText);
    if (values.length < 4) return null;

    let installmentAmount = values[0];
    let principalPaid = values[1];
    let interest = values[2];
    let serviceTax = "";
    let closingPrincipal = values[values.length - 1];

    const val0 = toNumber(values[0]);
    const val1 = toNumber(values[1]);
    const val2 = toNumber(values[2]);

    if (values.length >= 5) {
      if (Math.abs(val0 - (val1 + val2)) <= Math.max(val0 * 0.15, 50)) {
        installmentAmount = values[0];
        principalPaid = values[1];
        interest = values[2];
        serviceTax = values.length >= 6 ? values[3] : "";
        closingPrincipal = values[values.length - 1];
      } else if (val0 > val1 && val0 > val1 * 2 && val0 > 1000) {
        installmentAmount = values[1];
        principalPaid = values[2];
        interest = values[3];
        closingPrincipal = values[values.length - 1];
      } else {
        installmentAmount = values[0];
        principalPaid = values[1];
        interest = values[2];
        serviceTax = values.length >= 6 ? values[3] : "";
        closingPrincipal = values[values.length - 1];
      }
    }

    let emi = toNumber(installmentAmount);
    let prin = toNumber(principalPaid);
    let intr = toNumber(interest);

    // Auto-fix OCR missing decimal point using EMI = Principal + Interest
    if (emi > 100) {
      if (intr > 0 && intr < emi && (prin <= 0 || prin > emi || Math.abs(emi - (prin + intr)) > 50)) {
        prin = Math.max(0, emi - intr);
        principalPaid = prin.toFixed(2);
      } else if (prin > 0 && prin < emi && (intr <= 0 || intr > emi || Math.abs(emi - (prin + intr)) > 50)) {
        intr = Math.max(0, emi - prin);
        interest = intr.toFixed(2);
      }
    }

    if (toNumber(installmentAmount) <= 0 || toNumber(principalPaid) < 0) return null;
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  };

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(?:0*(\d{1,3})|(?:\/|\|)?\s*(\d{1,3}))\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const inst = match[1] || match[2];
      if (!inst || Number(inst) <= 0) return null;
      return parseTataRowValues(inst, match[3], match[4]);
    })
    .filter(Boolean);

  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseTataRowValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
}

function isDateFreeRepaymentFormat(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  const hasInstallmentColumn = /\b(?:Instalmnt|Instalment|Installment|Instl\.?\s*(?:No|Number|Amt)|Inst\.?\s*(?:No|Number|Amt))\b/i.test(normalized);
  return hasInstallmentColumn &&
    /Outstanding\s+Principal/i.test(normalized) &&
    /\bPrincipal\b/i.test(normalized) &&
    /\bInterest\b/i.test(normalized) &&
    /\b(?:Instalmnt|Instalment|Installment)\b/i.test(normalized);
}

function parseDateFreeScheduleRows(text) {
  const rawText = String(text ?? "").replace(/\u00a0/g, " ");
  if (!isDateFreeRepaymentFormat(rawText)) return [];
  const moneyToken = "-?\\d[\\d,]*(?:\\.\\d{1,2})?";
  const rowPattern = new RegExp(`^\\s*(\\d{1,3})\\s+(${moneyToken})\\s+(${moneyToken})\\s+(${moneyToken})\\s+(${moneyToken})(?:\\s|$)`, "i");

  const parseRow = (match) => {
    if (!match) return null;
    const [installment, openingPrincipal, principalPaid, interest, installmentAmount] = match.slice(1);
    const total = toNumber(principalPaid) + toNumber(interest);
    const amount = toNumber(installmentAmount);
    if (Number(installment) <= 0 || toNumber(openingPrincipal) <= 0 || amount <= 0) return null;
    // The four columns are a repayment breakup. Reject header/noise matches
    // unless principal + interest agrees with the instalment amount.
    if (Math.abs(amount - total) > Math.max(10, amount * 0.12)) return null;
    return {
      installment: Number(installment),
      openingPrincipal,
      principalPaid,
      interest,
      installmentAmount
    };
  };

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => parseRow(line.replace(/\s+/g, " ").trim().match(rowPattern)))
    .filter(Boolean);
  const parsedRows = lineRows.length > 0 ? lineRows : [...rawText.replace(/\s+/g, " ").matchAll(
    new RegExp(`(?:^|\\s)(\\d{1,3})\\s+(${moneyToken})\\s+(${moneyToken})\\s+(${moneyToken})\\s+(${moneyToken})(?=\\s|$)`, "gi")
  )].map(parseRow).filter(Boolean);

  return parsedRows.map((row, index) => {
    const nextOpeningPrincipal = parsedRows[index + 1]?.openingPrincipal;
    const closingPrincipal = nextOpeningPrincipal ?? String(Math.max(
      toNumber(row.openingPrincipal) - toNumber(row.principalPaid),
      0
    ));
    return makeScheduleRow({
      installment: row.installment,
      dueDate: "",
      openingPrincipal: row.openingPrincipal,
      installmentAmount: row.installmentAmount,
      principalPaid: row.principalPaid,
      interest: row.interest,
      closingPrincipal,
      rate: deriveAnnualRate(row.interest, row.openingPrincipal)
    });
  });
}

function inferProfileAgreementFromPdf(pdfRow = {}, rows = []) {
  const candidates = (rows ?? []).filter((row) => !isBodyRow(row) && isValidAgreementValue(row.loanAccount));
  if (candidates.length === 1) return candidates[0].loanAccount;

  const normalizeMatchText = (value) => String(value ?? "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");
  const pdfCustomer = normalizeMatchText(pdfRow.owner);
  if (pdfCustomer.length >= 4) {
    const customerMatches = candidates.filter((row) => {
      const profileCustomer = normalizeMatchText(row.owner || row.clientName || row.customerName);
      return profileCustomer && (profileCustomer === pdfCustomer ||
        profileCustomer.includes(pdfCustomer) || pdfCustomer.includes(profileCustomer));
    });
    if (customerMatches.length === 1) return customerMatches[0].loanAccount;
  }

  const pdfRegNo = normalizeRegNo(baseRegNo(pdfRow.regNo));
  if (pdfRegNo) {
    const registrationMatches = candidates.filter((row) =>
      normalizeRegNo(baseRegNo(row.regNo)) === pdfRegNo
    );
    if (registrationMatches.length === 1) return registrationMatches[0].loanAccount;
  }
  return "";
}

function summarizeDateFreeScheduleRows(rows, text) {
  const uniqueRows = Array.from(new Map((rows ?? []).map((row) => [row.installment, row])).values())
    .sort((first, second) => first.installment - second.installment);
  const firstRow = uniqueRows[0];
  const lastRow = uniqueRows[uniqueRows.length - 1];
  if (!firstRow || !lastRow) return {};
  const emiAmount = mostCommonAmount(uniqueRows.map((row) => row.installmentAmount));
  const declaredTenure = findDeclaredTenure(text);
  const explicitPaidEmi = findPdfValue(text, ["Paid EMI", "Paid Instl", "Paid Instalment", "Paid Installment"], { requireDigit: true });
  const paidEmi = explicitPaidEmi === "" ? "" : String(Number(explicitPaidEmi.match(/\d{1,3}/)?.[0] ?? 0));
  const startDate = findPdfDate(text, ["EMI Start Date", "Installment Start Date", "Instalment Start Date", "First EMI Date"]);
  const endDate = findPdfDate(text, ["EMI End Date", "Maturity Date", "Last EMI Date", "Last Instalment date", "Last Installment date"]);
  return {
    scheduleParsed: "yes",
    loanAmount: firstRow.openingPrincipal,
    emiAmount,
    tenure: String(declaredTenure || lastRow.installment || uniqueRows.length),
    paidEmi,
    // Do not invent a rate: this format does not contain an interest-rate column.
    interestRate: findPdfRate(text) || "",
    emiStart: startDate,
    emiEnd: endDate,
    bankClosingPrincipal: lastRow.closingPrincipal,
    // This format has no date column, so it cannot safely create dated due tasks.
    emiSchedule: []
  };
}

function findScheduleTableValues(text) {
  // Recognize Poonawalla before generic parsers interpret Rate / Days as money.
  if (isPoonawallaScheduleFormat(text)) {
    const rows = parseOpeningBalanceScheduleRows(text);
    return rows.length > 0 ? summarizeScheduleRows(rows, text) : {};
  }
  if (isIndostarRepaymentLayout(text)) {
    const rows = parseIndostarScheduleRows(text);
    return rows.length > 0 ? summarizeScheduleRows(rows, text) : {};
  }
  const auRows = parseAuScheduleRows(text);
  if (auRows.length > 0) return summarizeScheduleRows(auRows, text);
  // Cholamandalam has a specific format that must be parsed before generic parsers.
  const cholaRows = parseCholaScheduleRows(text);
  if (cholaRows.length > 0) return summarizeScheduleRows(cholaRows, text);
  // This two-date format has a stable column order and must be handled
  // before generic repayment parsers can mistake its columns for another
  // bank's schedule.
  const ashokLeylandRows = parseTwoDateScheduleRows(text);
  if (ashokLeylandRows.length > 0) return summarizeScheduleRows(ashokLeylandRows, text);
  // Tata's six-column schedule also resembles generic TVS tables; select its
  // validated row parser first so the complete 33-row schedule is retained.
  if (isTataFinanceDocument(text)) {
    const tataRows = parseTataScheduleRows(text);
    if (tataRows.length > 0) return summarizeScheduleRows(tataRows, text);
  }
  const tvsRows = parseTvsScheduleRows(text);
  if (tvsRows.length > 0) return summarizeScheduleRows(tvsRows, text);
  const indostarRows = parseIndostarScheduleRows(text);
  if (indostarRows.length > 0) return summarizeScheduleRows(indostarRows, text);
  const tataRows = parseTataScheduleRows(text);
  if (tataRows.length > 0) return summarizeScheduleRows(tataRows, text);
  const dateFreeRows = parseDateFreeScheduleRows(text);
  if (dateFreeRows.length > 0) return summarizeDateFreeScheduleRows(dateFreeRows, text);
  const indusIndRows = parseIndusIndScheduleRows(text);
  if (indusIndRows.length > 0) return summarizeScheduleRows(indusIndRows, text);
  const bandhanRows = parseBandhanScheduleRows(text);
  if (bandhanRows.length > 0) return summarizeScheduleRows(bandhanRows, text);
  const bajajRows = parseBajajScheduleRows(text);
  if (bajajRows.length > 0) return summarizeScheduleRows(bajajRows, text);
  const mahindraRows = parseMahindraScheduleRows(text);
  if (mahindraRows.length > 0) return summarizeScheduleRows(mahindraRows, text);
  const openingBalanceRows = parseOpeningBalanceScheduleRows(text);
  if (openingBalanceRows.length > 0) return summarizeScheduleRows(openingBalanceRows, text);
  const sundaramRows = parseSundaramScheduleRows(text);
  if (sundaramRows.length > 0) return summarizeScheduleRows(sundaramRows, text);
  const instlOutstandingRows = parseInstlOutstandingScheduleRows(text);
  if (instlOutstandingRows.length > 0) return summarizeScheduleRows(instlOutstandingRows, text);
  const bankNameRows = parseBankNameScheduleRows(text);
  if (bankNameRows.length > 0) return summarizeScheduleRows(bankNameRows, text);
  if (isIndostarScheduleFormat(text)) return {};
  const accountRows = parseAccountStatementScheduleRows(text);
  if (accountRows.length > 0) return summarizeScheduleRows(accountRows, text);
  const lineRows = String(text ?? "")
    .split(/\n+/)
    .map(parseScheduleLine)
    .filter(Boolean);
  const flatRows = parseFlattenedScheduleRows(text);
  const rows = lineRows.length >= flatRows.length ? lineRows : flatRows;
  if (rows.length === 0) return {};
  return summarizeScheduleRows(rows, text);
}

function parseTwoDateScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = collapseAdjacentDuplicateTokens(rawText.replace(/\s+/g, " "));
  const lineRows = rawText
    .split(/\n+/)
    .map((line) => collapseAdjacentDuplicateTokens(line.replace(/\s+/g, " ").trim()))
    .map(parseAshokLeylandScheduleLine)
    .filter(Boolean);
  if (lineRows.length > 0) return lineRows;
  return parseAshokLeylandStrictRows(normalized);
}

function parseBajajScheduleRows(text) {
  const normalized = normalizeWrappedPdfNumbers(String(text ?? "").replace(/\s+/g, " "));
  if (!/BAJAJ\s+FINANCE/i.test(normalized) && !/Agreement\s+No/i.test(normalized)) return [];
  if (!/Opening\s+Principal/i.test(normalized) || !/Instl\.?\s*Amt/i.test(normalized) || !/Closing\s+Principal/i.test(normalized)) return [];
  const rowStartPattern = /(?:^|\s)(D|B\d+|\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/gi;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const installmentLabel = match[1].toUpperCase();
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const segment = normalized.slice(rowStart, rowEnd);
    const values = extractSignedMoneyLikeNumbers(segment);
    if (!/^\d+$/.test(installmentLabel) || values.length < 6) return null;
    const openingPrincipal = values[0];
    const installmentAmount = values[1];
    const principalPaid = values[2];
    const interest = values[3];
    const closingPrincipal = values[5];
    const rate = values.find((value) => toNumber(value) > 0 && toNumber(value) <= 40) || "";
    return makeScheduleRow({
      installment: installmentLabel,
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      closingPrincipal,
      rate
    });
  }).filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
}

function normalizeWrappedPdfNumbers(value) {
  return String(value ?? "")
    // A few Bajaj PDFs split the last two digits of dates and amounts into
    // separate PDF text objects: 10/09/20 26 and 45,38,02 4.00.
    .replace(/(\d{1,2}[./-]\d{1,2}[./-]\d{2})\s+(\d{2})\b/g, "$1$2")
    .replace(/(\d{1,3}(?:,\d{2,3})+)\s+(\d{1,3}\.\d{1,2})\b/g, "$1$2")
    .replace(/(\d[\d,]*\.)\s+(\d{1,2})\b/g, "$1$2")
    .replace(/(\d[\d,]*\.\d)\s+(\d)\b/g, "$1$2")
    .replace(/(\d+)\s*,\s*(\d+)/g, "$1,$2");
}

function parseMahindraScheduleRows(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  if (!/MAHINDRA\s+AND\s+MAHINDRA\s+FINANCIAL\s+SERVICES/i.test(normalized)) return [];
  if (!/INSTAL\.?\s*AMT/i.test(normalized) || !/PRINCIPAL\.?\s*O\/?S/i.test(normalized)) return [];
  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const values = extractSignedMoneyLikeNumbers(normalized.slice(rowStart, rowEnd));
    if (values.length < 5) return null;
    const [installmentAmount, principalPaid, income, unmaturedIncome, closingPrincipal] = values.slice(-5);
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest: income,
      serviceTax: unmaturedIncome,
      closingPrincipal,
      rate: deriveAnnualRate(income, openingPrincipal)
    });
  }).filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
}

function isPoonawallaScheduleFormat(text) {
  const source = String(text ?? "").replace(/\s+/g, " ");
  return (/POONAWALLA\s*FINCORP/i.test(source) && /(?:Repayment|Install?ment)\s+Schedule/i.test(source)) ||
    (/Opening\s+Balance/i.test(source) && /Effective\s+Rate/i.test(source) && /\bDays\b/i.test(source));
}

function findPoonawallaOwner(text) {
  if (!isPoonawallaScheduleFormat(text)) return "";
  const match = String(text ?? "").match(/Repayment\s+Schedule\s+For\s+(?:Agreement\s+No\.?\s*[:\-]?\s*)?[A-Z0-9][A-Z0-9_/-]*\s+([A-Za-z][A-Za-z .&'()-]{2,100}?)\s+Note\s*:/i);
  return cleanPdfPersonName(match?.[1] ?? "");
}

function parseOpeningBalanceScheduleRows(text) {
  const source = String(text ?? "");
  if (!isPoonawallaScheduleFormat(source) &&
    !(/Opening\s+Balance/i.test(source) && /Effective\s+Rate/i.test(source))) return [];
  const normalize = (value) => normalizeWrappedPdfNumbers(value.replace(/[|\u00a0]/g, " ").replace(/\s+/g, " "))
    .replace(/(\d{1,2})\s*[-/]\s*([A-Za-z]{3,})\s*[-/]\s*(\d{2,4})/g, "$1-$2-$3");
  const date = "\\d{1,2}(?:[- ][A-Za-z]{3,}[- ]|[./-]\\d{1,2}[./-])\\d{2,4}";
  const money = "\\d[\\d,]*(?:\\.\\d{1,2})?";
  // Instl | Date | Opening | EMI | Principal | Interest | Effective rate / Days.
  // Match only contiguous columns, never numbers from totals or the next pass.
  const pattern = new RegExp(`(?:^|\\s)(\\d{1,3})\\s+(${date})\\s+(${money})\\s+(${money})\\s+(${money})\\s+(${money})\\s+(\\d{1,2}(?:\\.\\d{1,4})?)\\s*%?\\s*/\\s*(\\d{1,3})(?=\\s|$)`, "gi");
  // Prefer intact table lines over sparse OCR's reordered cells.
  const lineMatches = source.split(/\n+/).flatMap((line) => [...normalize(line).matchAll(pattern)]);
  const matches = lineMatches.length > 0 ? lineMatches : [...normalize(source).matchAll(pattern)];
  const rows = new Map();
  for (const match of matches) {
    const [, installment, rawDate, openingPrincipal, installmentAmount, principalPaid, interest, rate, days] = match;
    const dueDate = formatDisplayDate(rawDate);
    const parsedDate = parseDisplayDate(dueDate);
    const opening = toNumber(openingPrincipal);
    const principal = toNumber(principalPaid);
    const emi = toNumber(installmentAmount);
    const interestAmount = toNumber(interest);
    if (!parsedDate || formatDisplayDate(scheduleDateToIso(dueDate)) !== dueDate ||
      Number(installment) <= 0 || opening <= 0 || emi <= 0 || principal > opening ||
      Math.abs(emi - principal - interestAmount) > 1 || Number(rate) > 40 ||
      Number(days) <= 0 || Number(days) > 366) continue;
    const key = `${Number(installment)}|${dueDate}`;
    // Keep the first validated reading; sparse OCR can reorder later copies.
    if (!rows.has(key)) rows.set(key, makeScheduleRow({
      installment, dueDate, openingPrincipal, installmentAmount, principalPaid, interest, rate,
      closingPrincipal: (opening - principal).toFixed(2)
    }));
  }
  return [...rows.values()].sort((first, second) => first.installment - second.installment);
}

function parseSundaramScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!/SUNDARAM\s+FINANCE/i.test(normalized) && !/Details\s+of\s+Interest\s+and\s+Principal\s+Breakup/i.test(normalized)) return [];
  if (!/Due\s+Amount/i.test(normalized) || !/Loan\s+O\/?S/i.test(normalized)) return [];
  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const values = extractSignedMoneyLikeNumbers(match[3]);
      if (values.length < 5) return null;
      const [installmentAmount, principalPaid, interest, insurance, closingPrincipal] = values.slice(-5);
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({
        installment: match[1],
        dueDate: match[2],
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        serviceTax: insurance,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    })
    .filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const values = extractSignedMoneyLikeNumbers(normalized.slice(rowStart, rowEnd));
    if (values.length < 5) return null;
    const [installmentAmount, principalPaid, interest, insurance, closingPrincipal] = values.slice(-5);
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax: insurance,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  }).filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
}

function parseAshokLeylandScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = collapseAdjacentDuplicateTokens(rawText.replace(/\s+/g, " "));
  if (!isAshokLeylandScheduleFormat(normalized)) return [];

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => collapseAdjacentDuplicateTokens(line.replace(/\s+/g, " ").trim()))
    .map(parseAshokLeylandScheduleLine)
    .filter(Boolean);
  if (lineRows.length > 0) return lineRows;

  const strictRows = parseAshokLeylandStrictRows(normalized);
  if (strictRows.length > 0) return strictRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(?:(?:P?E?M?I|P\s*EMI|EM1|FMI)\s+)?(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/gi;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts
    .map((match, index) => {
      const rowStart = match.index + match[0].length;
      const rowEnd = starts[index + 1]?.index ?? normalized.length;
      return parseAshokLeylandScheduleValues(match[1], match[3], normalized.slice(rowStart, rowEnd));
    })
    .filter(Boolean);
}

function parseAshokLeylandScheduleLine(line) {
  const match = String(line ?? "").match(/^(\d{1,3})\s+(?:(?:P?E?M?I(?:\s*\/\s*E?M?I)?|P\s*EMI|EM1|FMI)\s+)?\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/i);
  if (!match) return null;
  return parseAshokLeylandScheduleValues(match[1], match[2], match[3]);
}

function parseAshokLeylandStrictRows(text) {
  const rowPattern = /(?:^|\s)(\d{1,3})\s+(?:(?:P?E?M?I(?:\s*\/\s*E?M?I)?|P\s*EMI|EM1|FMI)\s+)?\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(\d{2,4}(?:\.\d{1,4})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)(?:\s+\d{1,3})?(?=\s|$)/gi;
  return [...String(text ?? "").matchAll(rowPattern)]
    .map((match) => parseAshokLeylandScheduleValues(match[1], match[2], match.slice(3).join(" ")))
    .filter(Boolean);
}

function parseAshokLeylandScheduleValues(installment, dueDate, valueText) {
  const values = extractSignedMoneyLikeNumbers(valueText);
  if (values.length < 5) return null;
  const [rawRate, rawPrincipalPaid, rawInterest, rawInstallmentAmount, rawClosingPrincipal] = values.slice(0, 5);
  const rate = normalizeAshokRate(rawRate);
  const principalPaid = normalizeAshokAmount(rawPrincipalPaid);
  const interest = normalizeAshokAmount(rawInterest);
  const installmentAmount = normalizeAshokAmount(rawInstallmentAmount);
  const closingPrincipal = normalizeAshokAmount(rawClosingPrincipal);
  const rateValue = toNumber(rate);
  const emiBreakupDiff = Math.abs(toNumber(installmentAmount) - toNumber(principalPaid) - toNumber(interest));
  const allowedDiff = Math.max(5, toNumber(installmentAmount) * 0.05);
  if (
    rateValue <= 0 ||
    rateValue > 40 ||
    toNumber(principalPaid) <= 0 ||
    toNumber(interest) <= 0 ||
    toNumber(installmentAmount) <= 0 ||
    toNumber(closingPrincipal) < 0 ||
    emiBreakupDiff > allowedDiff
  ) {
    return null;
  }
  const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
  return makeScheduleRow({
    installment,
    dueDate,
    openingPrincipal,
    installmentAmount,
    principalPaid,
    interest,
    closingPrincipal,
    rate
  });
}

function normalizeAshokRate(value) {
  const clean = String(value ?? "").replace(/[^\d.]/g, "");
  const number = Number(clean);
  if (Number.isFinite(number) && number > 0 && number <= 40) return clean;
  if (/^\d{3,4}$/.test(clean)) {
    const scaled = Number(clean) / 100;
    if (scaled > 0 && scaled <= 40) return scaled.toFixed(2);
  }
  return clean;
}

function normalizeAshokAmount(value) {
  const clean = String(value ?? "").replace(/[^\d.]/g, "");
  if (!clean) return "";
  if (clean.includes(".")) return clean;
  return clean;
}

function isAshokLeylandScheduleFormat(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  const hasAshokBrand = /ASHOK\s*LEYLAND|ASHOKLEYLAND/i.test(normalized);
  const hasScheduleTitle = /AMORTI[ZS]ATION|REPAYMENT\s+SCHEDULE|TENTATIVE\s+LOAN\s+REPAYMENT/i.test(normalized);
  const hasExactHeaders = /Install(?:ation|ment)\s+(?:Number|No)|Inst\s*No/i.test(normalized) &&
    /Start\s+Date/i.test(normalized) &&
    /Repayment\s+Date/i.test(normalized) &&
    /Interest\s+Rate|Rate\s*\(?%?\)?/i.test(normalized) &&
    /Total\s+Instal(?:l)?ment|Instal(?:l)?ment\s+Amount/i.test(normalized) &&
    /Outstanding\s+Balance|Outstanding\s+Principal/i.test(normalized);
  const hasAshokRows = /\b\d{1,3}\s+(?:(?:P?E?M?I(?:\s*\/\s*E?M?I)?|P\s*EMI|EM1|FMI)\s+)?\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s+\d{2,4}(?:\.\d{1,4})?\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?/i.test(normalized);
  return (hasAshokBrand || hasScheduleTitle || hasExactHeaders) && hasAshokRows;
}

function findAshokLeylandSummaryValues(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  if (!/ASHOK\s*LEYLAND|ASHOKLEYLAND|Repayment\s+Schedule|Total\s+Loan\s+Sanctioned/i.test(normalized)) return {};
  const values = {
    loanAmount: findPdfAmount(text, ["Total Loan Sanctioned", "Total Loan Disbursed"], { positive: true, min: 1000 }),
    tenure: findPdfValue(text, ["Period in Months"], { requireDigit: true, minLength: 1 }),
    interestRate: findPdfRate(text) || findPdfValue(text, ["Current Int Rate(%)", "Current Int Rate", "Interest Rate"], { requireDigit: true, minLength: 1 }),
    emiEnd: findPdfDate(text, ["Maturity Date"])
  };
  return Object.fromEntries(Object.entries(values).filter(([, value]) => value));
}

function parseInstlOutstandingScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = collapseAdjacentDuplicateTokens(rawText.replace(/\s+/g, " "));
  if (!isInstlOutstandingScheduleFormat(normalized)) return [];

  const columnRows = parseInstlOutstandingColumnRows(normalized);
  if (columnRows.length > 0) return columnRows;

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => collapseAdjacentDuplicateTokens(line.replace(/\s+/g, " ").trim()))
    .map(parseInstlOutstandingLine)
    .filter(Boolean);
  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  const segmentedRows = starts
    .map((match, index) => {
      const rowStart = match.index + match[0].length;
      const rowEnd = starts[index + 1]?.index ?? normalized.length;
      return parseInstlOutstandingValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
    })
    .filter(Boolean);
  if (segmentedRows.length > 0) return segmentedRows;

  return parseInstlOutstandingStrictRows(normalized);
}

function collapseAdjacentDuplicateTokens(value) {
  const tokens = String(value ?? "").split(/\s+/).filter(Boolean);
  const collapsed = [];
  for (let index = 0; index < tokens.length; index += 1) {
    const token = tokens[index];
    if (token && token === tokens[index + 1]) {
      collapsed.push(token);
      index += 1;
    } else {
      collapsed.push(token);
    }
  }
  return collapsed.join(" ");
}

function parseInstlOutstandingColumnRows(normalized) {
  const installmentNumbers = extractColumnNumbers(normalized, /Instl\.?\s*No/i, /Due\s*Date/i)
    .map((value) => Number(value))
    .filter((value) => Number.isFinite(value) && value > 0 && value <= 240);
  const dueDates = extractColumnDates(normalized, /Due\s*Date/i, /Cheque\s*No/i);
  const rowCount = Math.min(installmentNumbers.length, dueDates.length);
  if (rowCount < 3) return [];

  const emiValues = extractColumnMoney(normalized, /\bEMI\b/i, /\bPrincipal\b/i).slice(0, rowCount);
  const principalValues = extractColumnMoney(normalized, /\bPrincipal\b/i, /\bInterest\b/i).slice(0, rowCount);
  const interestValues = extractColumnMoney(normalized, /\bInterest\b/i, /Outstanding\s*Principal/i).slice(0, rowCount);
  const closingValues = extractColumnMoney(normalized, /Outstanding\s*Principal/i).slice(0, rowCount);

  if ([emiValues, principalValues, interestValues, closingValues].some((values) => values.length < rowCount)) return [];
  return Array.from({ length: rowCount }, (_, index) => parseInstlOutstandingValues(
    String(installmentNumbers[index]),
    dueDates[index],
    [emiValues[index], principalValues[index], interestValues[index], closingValues[index]].join(" ")
  )).filter(Boolean);
}

function parseInstlOutstandingLine(line) {
  const match = String(line ?? "").match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
  if (!match) return null;
  return parseInstlOutstandingValues(match[1], match[2], match[3]);
}

function parseInstlOutstandingValues(installment, dueDate, valueText) {
  const values = extractSignedMoneyLikeNumbers(valueText);
  if (values.length < 4) return null;
  const amounts = chooseInstlOutstandingAmounts(values);
  if (!amounts) return null;
  const [installmentAmount, principalPaid, interest, closingPrincipal] = amounts;
  const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
  return makeScheduleRow({
    installment,
    dueDate,
    openingPrincipal,
    installmentAmount,
    principalPaid,
    interest,
    closingPrincipal,
    rate: deriveAnnualRate(interest, openingPrincipal)
  });
}

function parseInstlOutstandingStrictRows(text) {
  const rowPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(?:N\/?A|NIA|NA|-|[A-Z0-9./-]+)\s+(?:N\/?A|NIA|NA|-|[A-Z0-9./-]+)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)(?=\s|$)/gi;
  return [...String(text ?? "").matchAll(rowPattern)]
    .map((match) => parseInstlOutstandingValues(match[1], match[2], match.slice(3).join(" ")))
    .filter(Boolean);
}

function isInstlOutstandingScheduleFormat(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  return /Instl\.?\s*No/i.test(normalized) &&
    /Cheque\s*No/i.test(normalized) &&
    /Bank\s*Name/i.test(normalized) &&
    /\bEMI\b/i.test(normalized) &&
    /Principal/i.test(normalized) &&
    /Interest/i.test(normalized) &&
    /Outstanding\s*Principal/i.test(normalized);
}

function chooseInstlOutstandingAmounts(values) {
  const cleanedValues = (values ?? []).map((value) => String(value).replace(/,/g, ""));
  for (let index = 0; index <= cleanedValues.length - 4; index += 1) {
    const candidate = cleanedValues.slice(index, index + 4);
    const [installmentAmount, principalPaid, interest, closingPrincipal] = candidate.map(toNumber);
    const emiBreakupDiff = Math.abs(installmentAmount - principalPaid - interest);
    const allowedDiff = Math.max(5, installmentAmount * 0.05);
    if (
      installmentAmount >= 100 &&
      principalPaid > 0 &&
      interest > 0 &&
      (closingPrincipal === 0 || closingPrincipal > principalPaid) &&
      emiBreakupDiff <= allowedDiff
    ) {
      return candidate;
    }
  }
  return null;
}

function extractColumnNumbers(text, startPattern, endPattern) {
  return extractMoneyLikeNumbers(extractColumnText(text, startPattern, endPattern));
}

function extractColumnDates(text, startPattern, endPattern) {
  return [...extractColumnText(text, startPattern, endPattern).matchAll(/\b\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/g)]
    .map((match) => match[0]);
}

function extractColumnMoney(text, startPattern, endPattern) {
  return extractSignedMoneyLikeNumbers(extractColumnText(text, startPattern, endPattern));
}

function extractColumnText(text, startPattern, endPattern) {
  const source = String(text ?? "");
  const startMatch = source.match(startPattern);
  if (!startMatch) return "";
  const start = (startMatch.index ?? 0) + startMatch[0].length;
  const tail = source.slice(start);
  const endMatch = endPattern ? tail.match(endPattern) : null;
  return endMatch ? tail.slice(0, endMatch.index) : tail;
}

function summarizeScheduleRows(rows, text) {
  // Some bank PDFs contain the same text twice because the visible table is
  // backed by two overlapping PDF text layers. Keep one row per installment.
  const uniqueRows = Array.from(new Map((rows ?? []).map((row) => [
    `${row.installment}|${row.dueDate}`,
    row
  ])).values());
  const sortedRows = uniqueRows.sort((first, second) => first.installment - second.installment);
  const lastRow = sortedRows[sortedRows.length - 1];
  const payableRows = sortedRows.filter((row) => toNumber(row.installmentAmount) > 0);
  const emiCandidates = sortedRows
    .map((row) => row.installmentAmount)
    .filter((value) => toNumber(value) > 0);
  const emiAmount = mostCommonAmount(emiCandidates) || lastRow.installmentAmount;
  const paidRows = payableRows.filter((row) => isScheduleDuePaid(row.dueDate));
  const paidRow = paidRows[paidRows.length - 1];
  const closingPrincipal = chooseScheduleClosingPrincipal(paidRows, emiAmount) || sortedRows[0].openingPrincipal;
  const declaredTenure = findDeclaredTenure(text);
  const dateTenure = findTenureFromPdfDates(text);
  const tenure = declaredTenure || dateTenure || Math.max(payableRows.length || lastRow.installment, paidRows.length);
  const paidEmi = Math.min(paidRows.length, tenure);
  const emiSchedule = sortedRows.map((row) => ({
    installment: Number(row.installment),
    dueDate: scheduleDateToIso(row.dueDate),
    amount: toNumber(row.installmentAmount),
    principal: toNumber(row.principalPaid),
    interest: toNumber(row.interest),
    status: isScheduleDuePaid(row.dueDate) ? "Paid" : "Due"
  })).filter((row) => row.dueDate && row.amount > 0);
  return {
    scheduleParsed: "yes",
    loanAmount: sortedRows[0].openingPrincipal,
    emiAmount,
    tenure: String(tenure),
    paidEmi: String(paidEmi),
    interestRate: paidRow?.rate || lastRow.rate || sortedRows.find((row) => row.rate)?.rate || "",
    emiStart: formatDisplayDate(sortedRows[0].dueDate),
    emiEnd: formatDisplayDate(lastRow.dueDate),
    bankClosingPrincipal: closingPrincipal,
    emiSchedule
  };
}

function chooseScheduleClosingPrincipal(paidRows, emiAmount) {
  if (!paidRows.length) return "";
  const emi = toNumber(emiAmount);
  const lastPaidRow = paidRows[paidRows.length - 1];
  if (toNumber(lastPaidRow?.closingPrincipal) === 0 && isScheduleDuePaid(lastPaidRow?.dueDate)) {
    return lastPaidRow.closingPrincipal;
  }
  const lastClosing = toNumber(lastPaidRow?.closingPrincipal);
  if (lastClosing > 0 && (emi <= 0 || lastClosing > emi * 1.05 || paidRows.length === 1)) {
    return paidRows[paidRows.length - 1].closingPrincipal;
  }
  const previous = [...paidRows].reverse().find((row) => toNumber(row.closingPrincipal) > emi * 1.05);
  return previous?.closingPrincipal || paidRows[paidRows.length - 1]?.closingPrincipal || "";
}

function findTenureFromPdfDates(text) {
  const start = parseDisplayDate(findPdfDate(text, ["EMI Start Date", "Installment Start Date", "Instalment Start Date", "First EMI Date", "First Instalment date", "First Installment date"]));
  const end = parseDisplayDate(findPdfDate(text, ["EMI End Date", "Maturity Date", "Last EMI Date", "Last Instalment date", "Last Installment date"]));
  if (!start || !end || end < start) return 0;
  const months = (end.getFullYear() - start.getFullYear()) * 12 + (end.getMonth() - start.getMonth()) + 1;
  return months > 0 && months <= 240 ? months : 0;
}

function parseAccountStatementScheduleRows(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  const rowPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})\s+(\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})\s+([0-9]+(?:\.[0-9]+)?)\s+\d{1,3}\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)(?=\s|$)/gi;
  return [...normalized.matchAll(rowPattern)].map((match) => {
    const principalPaid = match[5];
    const interest = match[6];
    const charge = match[7];
    const installmentAmount = match[8];
    const closingPrincipal = match[9];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[3],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax: charge,
      closingPrincipal,
      rate: match[4]
    });
  }).filter((row) => row.installment > 0 && toNumber(row.closingPrincipal) >= 0);
}

function parseBankNameScheduleRows(text) {
  const rawText = String(text ?? "");
  const normalized = rawText.replace(/\s+/g, " ");
  if (!isIndostarScheduleFormat(normalized) && (!/Bank\s*Name/i.test(normalized) || !/(?:Out\s*Stand(?:ing)?|Outstanding)\s*Principal/i.test(normalized))) return [];
  const segmentedRows = parseBankNameSegmentRows(normalized);
  if (segmentedRows.length > 0) return segmentedRows;

  const lineRows = rawText
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
      if (!match) return null;
      const values = extractSignedMoneyLikeNumbers(match[3]);
      if (values.length < 5) return null;
      const [installmentAmount, principalPaid, interest, mi, closingPrincipal] = values.slice(-5);
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({
        installment: match[1],
        dueDate: match[2],
        openingPrincipal,
        installmentAmount,
        principalPaid,
        interest,
        serviceTax: mi,
        closingPrincipal,
        rate: deriveAnnualRate(interest, openingPrincipal)
      });
    })
    .filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
  if (lineRows.length > 0) return lineRows;

  const rowPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(?:[A-Z][A-Z&.-]*\s+){1,8}?(-?\d[\d,]*(?:\.\d{1,2})?)\s+(-?\d[\d,]*(?:\.\d{1,2})?)\s+(-?\d[\d,]*(?:\.\d{1,2})?)\s+(-?\d[\d,]*(?:\.\d{1,2})?)\s+(-?\d[\d,]*(?:\.\d{1,2})?)(?=\s|$)/gi;
  return [...normalized.matchAll(rowPattern)].map((match) => {
    const installmentAmount = match[3];
    const principalPaid = match[4];
    const interest = match[5];
    const mi = match[6];
    const closingPrincipal = match[7];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax: mi,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  }).filter((row) => row.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
}

function isIndostarRepaymentLayout(text) {
  const source = String(text ?? "").replace(/\s+/g, " ");
  return /\bINDO\s*STAR\b/i.test(source) ||
    (/REPAYMENT\s+SCHEDULE\s+FOR\s+AGREEMENT/i.test(source) &&
      /Bank\s+Name/i.test(source) && /\bMI\b/i.test(source) && /Out\s*Standing/i.test(source)) ||
    (/EMI\s+No\.?\s+Due\s+Date\s+EMI\s+Amount\s+Principal\s+Interest\s+Balance\s+Outstanding/i.test(source) && !isTataFinanceDocument(source));
}

function isIndostarScheduleFormat(text) {
  const value = String(text ?? "");
  return isIndostarRepaymentLayout(value) || /REPAYMENT\s*SCHEDULE\s*FOR\s*AGREEMENT/i.test(value) ||
    /INDOSTAR/i.test(value) ||
    (/Balance\s+Outstanding/i.test(value) && /Due\s+Date/i.test(value) && /\bEMI\b/i.test(value)) ||
    (/Vehicle\s*Number/i.test(value) && /Out\s*Stand(?:ing)?\s*Principal/i.test(value) && /\bEMI\b/i.test(value));
}

function findLegacyIndostarSummaryValues(text) {
  const source = String(text ?? "").replace(/\s+/g, " ").trim();
  if (!/INDOSTAR/i.test(source)) return {};
  const amountAfter = (pattern) => source.match(new RegExp(`${pattern}\s*[:\-]?\s*(?:INR|RS\.?|₹)?\s*([\d,]+(?:\.\d{1,2})?)`, "i"))?.[1]?.replace(/,/g, "") ?? "";
  const numberAfter = (pattern) => source.match(new RegExp(`${pattern}\s*[:\-]?\s*(\d{1,4})`, "i"))?.[1] ?? "";
  const dateAfter = (pattern) => source.match(new RegExp(`${pattern}\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4}|\d{1,2}[-\s][A-Za-z]{3,}[-\s]\d{2,4})`, "i"))?.[1] ?? "";
  return {
    loanAmount: amountAfter("(?:Loan\\s+(?:Am|Amount)\\.?|Finance\\s+Amount|Amount\\s+Financed)"),
    emiAmount: amountAfter("(?:EMI\\s+(?:Am|Amount)\\.?|Instl\\.?\\s+(?:Am|Amount)\\.?|Installment\\s+Amount|Instalment\\s+Amount)"),
    tenure: numberAfter("(?:Tenure|No\\.?\\s*of\\s*EMIs?|No\\.?\\s*of\\s*Installments?|No\\.?\\s*of\\s*Instls?\\.?)"),
    interestRate: source.match(/(?:Rate\s+of\s+Interest|Interest\s+Rate|ROI)\s*[:\-]?\s*([\d]+(?:\.\d{1,4})?)/i)?.[1] ?? "",
    emiStart: dateAfter("(?:EMI\\s+Start|First\\s+Repayment\\s+Date|First\\s+EMI\\s+Date)"),
    emiEnd: dateAfter("(?:EMI\\s+End|Last\\s+Repayment\\s+Date|Last\\s+EMI\\s+Date|Maturity\\s+Date)")
  };
}

function findIndostarSummaryValues(text) {
  const source = normalizeWrappedPdfNumbers(String(text ?? "").replace(/[|\u00a0]/g, " ").replace(/\s+/g, " ").trim());
  if (!isIndostarScheduleFormat(source)) return {};
  const amount = (labels) => findPdfAmount(source, labels, { positive: true, min: 100 });
  const value = (labels) => findPdfValue(source, labels, { requireDigit: true, minLength: 1 });
  const agreementNumber = source.match(/Agreement\s+(?:Number|No)\.?\s*[:\-]*\s*([A-Z0-9][A-Z0-9_\/-]{5,})/i)?.[1] ?? "";
  const rows = parseIndostarScheduleRows(text);
  const scheduleEmi = mostCommonAmount(rows.map((row) => row.installmentAmount).filter((item) => toNumber(item) > 0));
  const rowNumbers = [...source.matchAll(/\b([1-9]\d{0,2})\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/gi)].map((match) => Number(match[1])).filter((number) => number > 0);
  const firstRowMatch = source.match(/\b1\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/i);
  const nextRowMatch = firstRowMatch ? source.slice((firstRowMatch.index ?? 0) + firstRowMatch[0].length).match(/\b2\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\b/i) : null;
  const firstRowText = firstRowMatch ? source.slice((firstRowMatch.index ?? 0) + firstRowMatch[0].length, (firstRowMatch.index ?? 0) + firstRowMatch[0].length + (nextRowMatch?.index ?? 240)) : "";
  const firstRowValues = extractSignedMoneyLikeNumbers(firstRowText);
  const fallbackEmi = firstRowValues[0] ?? "";
  const fallbackPrincipal = firstRowValues[1] ?? "";
  const fallbackClosing = firstRowValues.length >= 5 ? firstRowValues[4] : firstRowValues[3] ?? "";
  const fallbackLoanAmount = toNumber(fallbackPrincipal) > 0 && toNumber(fallbackClosing) >= 0 ? String(toNumber(fallbackPrincipal) + toNumber(fallbackClosing)) : "";
  const scheduleTenure = rows.length ? String(Math.max(...rows.map((row) => Number(row.installment) || 0))) : (rowNumbers.length ? String(Math.max(...rowNumbers)) : "");
  const loanAmountMatch = source.match(/Loan\s+(?:Amount|Am\.?)\s*[:\-]?\s*(?:INR|RS\.?|₹)?\s*([\d,]+(?:\.\d{1,2})?)/i);
  const emiAmountMatch = source.match(/(?:EMI|Instalment|Installment)\s+(?:Amount|Amt)\s*[:\-]?\s*(?:INR|RS\.?|₹)?\s*([\d,]+(?:\.\d{1,2})?)/i);
  const owner = source.match(/Customer(?:\s+Name)?\s*[:\-]?\s*(.+?)(?=\s+(?:Processing\s+Fee|Agreement|Loan\s+Amount|Rate\s+of\s+Interest|Vehicle\s+(?:No|Number)|EMI\s+Amount))/i)?.[1]?.trim() ?? "";
  const regNo = source.match(/Vehicle\s+(?:No|Number)\.?\s*[:\-]?\s*([A-Z]{2}\s*\d{1,2}\s*[A-Z]{1,3}\s*[- ]?\s*\d{3,4})/i)?.[1]?.replace(/\s+/g, " ").trim().toUpperCase() ?? "";
  const period = source.match(/EMI\s+Period\s*[:\-]?\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s*(?:to|[-–])\s*(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})/i);
  const startDate = period?.[1] ?? findPdfDate(source, ["EMI Period", "EMI Start Date", "First EMI Date"]);
  const endDate = period?.[2] ?? findPdfDate(source, ["EMI End Date", "Maturity Date", "Last EMI Date"]);
  return {
    agreementNumber,
    owner,
    regNo,
    loanAmount: loanAmountMatch?.[1]?.replace(/,/g, "") || amount(["Loan Amount", "Loan Am", "Finance Amount", "Amount Financed"]) || rows[0]?.openingPrincipal || fallbackLoanAmount,
    emiAmount: emiAmountMatch?.[1]?.replace(/,/g, "") || amount(["EMI Amount", "EMI Am", "Installment Amount", "Instalment Amount"]) || scheduleEmi || fallbackEmi,
    tenure: value(["Tenure / EMIs", "Tenure", "Number of EMIs", "No of Installments"]) || scheduleTenure,
    interestRate: source.match(/(?:Rate\s+of\s+Interest|Interest\s+Rate|ROI)\s*[:\-]?\s*([\d]+(?:\.\d{1,4})?)/i)?.[1] ?? "",
    emiStart: startDate ? formatDisplayDate(startDate) : "",
    emiEnd: endDate ? formatDisplayDate(endDate) : ""
  };
}

function findIndostarEmiAmount(text) {
  const rows = parseIndostarScheduleRows(text);
  const emiAmount = mostCommonAmount(rows.map((row) => row.installmentAmount).filter((value) => toNumber(value) > 0));
  if (emiAmount) return emiAmount;
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  const firstPayableRow = normalized.match(/\b[1-9]\d{0,2}\s+\d{1,2}[./-]\d{1,2}[./-]\d{2,4}\s+(?:[A-Z][A-Z&.-]*\s+){1,8}?(\d[\d,]*(?:\.\d{1,2})?)\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?\s+\d[\d,]*(?:\.\d{1,2})?/i);
  return firstPayableRow?.[1]?.replace(/,/g, "") ?? "";
}

function parseIndostarScheduleRows(text) {
  if (isIndostarRepaymentLayout(text)) return parseIndostarRepaymentRows(text);
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  if (!isIndostarScheduleFormat(normalized)) return [];
  const datePattern = "(\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}|\\d{1,2}[-\\s][A-Za-z]{3,}[-\\s]\\d{2,4})";
  const parseValues = (installment, dueDate, valueText) => {
    const values = extractSignedMoneyLikeNumbers(valueText);
    if (values.length < 4) return null;
    const makeCandidate = (installmentAmount, principalPaid, interest, closingPrincipal, serviceTax = "") => {
      const emiBreakup = Math.abs(toNumber(installmentAmount) - toNumber(principalPaid) - toNumber(interest));
      if (toNumber(installmentAmount) < 100 || toNumber(principalPaid) <= 0 || toNumber(interest) <= 0 || toNumber(closingPrincipal) < 0 || emiBreakup > Math.max(toNumber(installmentAmount) * 0.2, 100)) return null;
      const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
      return makeScheduleRow({ installment, dueDate, openingPrincipal, installmentAmount, principalPaid, interest, serviceTax, closingPrincipal, rate: deriveAnnualRate(interest, openingPrincipal) });
    };
    if (values.length >= 5) {
      const standard = makeCandidate(values[0], values[1], values[2], values[4], values[3]);
      if (standard) return standard;
      const breakupFirst = makeCandidate(values[3], values[0], values[1], values[4], values[2]);
      if (breakupFirst) return breakupFirst;
    }
    const direct = makeCandidate(values[0], values[1], values[2], values[3]);
    if (direct) return direct;
    return makeCandidate(values[2], values[0], values[1], values[3]);
  };
  const linePattern = new RegExp(`^(\\d{1,3})\\s+${datePattern}\\s+(.+)$`, "i");
  const lineRows = String(text ?? "")
    .split(/\n+/)
    .map((line) => line.replace(/\s+/g, " ").trim())
    .map((line) => {
      const match = line.match(linePattern);
      return match ? parseValues(match[1], match[2], match[3]) : null;
    })
    .filter(Boolean);
  if (lineRows.length > 0) return lineRows;

  const rowStartPattern = new RegExp(`(?:^|\\s)(\\d{1,3})\\s+${datePattern}(?=\\s)`, "gi");
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    return parseValues(match[1], match[2], normalized.slice(rowStart, rowEnd));
  }).filter(Boolean);
}

function parseIndostarRepaymentRows(text) {
  const normalize = (value) => normalizeWrappedPdfNumbers(String(value ?? "")
    .replace(/[|\u00a0]/g, " ").replace(/\s+/g, " "));
  const source = normalize(text);
  const hasMiColumn = /\bMI\b/i.test(source);
  const money = "-?\\d[\\d,]*(?:\\.\\d{1,2})?";
  const date = "\\d{1,2}\\s*[./-]\\s*\\d{1,2}\\s*[./-]?\\s*\\d{4}|\\d{1,2}[- ][A-Za-z]{3,}[- ]\\d{4}";
  // Original: EMI | Principal | Interest | MI | Outstanding.
  // Summary/screenshot: EMI | Principal | Interest | Outstanding.
  const columns = hasMiColumn ? 5 : 4;
  const amounts = Array.from({ length: columns }, () => `(${money})`).join("\\s+");
  const pattern = new RegExp(`(?:^|\\s)(\\d{1,3})\\s+(${date})\\s+(?:[A-Z][A-Z&.-]*\\s+){0,8}${amounts}(?=\\s|$)`, "gi");
  const explicitRate = source.match(/Rate\s+of\s+Interest\s*[:\-]?\s*(\d+(?:\.\d{1,4})?)/i)?.[1] ?? "";
  const rows = new Map();
  const addMatches = (value) => {
    for (const match of normalize(value).matchAll(pattern)) {
      const [, installment, rawDate, installmentAmount, principalPaid, interest] = match;
      const closingPrincipal = match[hasMiColumn ? 7 : 6];
      const serviceTax = hasMiColumn ? match[6] : "";
      const dueDate = formatDisplayDate(formatOcrDate(rawDate));
      const emi = toNumber(installmentAmount);
      const principal = toNumber(principalPaid);
      const closing = toNumber(closingPrincipal);
      if (Number(installment) <= 0 || !scheduleDateToIso(dueDate) ||
        formatDisplayDate(scheduleDateToIso(dueDate)) !== dueDate ||
        emi <= 0 || principal < 0 || toNumber(interest) < 0 || closing < 0 ||
        Math.abs(emi - principal - toNumber(interest)) > 1) continue;
      const key = `${Number(installment)}|${dueDate}`;
      if (!rows.has(key)) rows.set(key, makeScheduleRow({
        installment, dueDate, installmentAmount, principalPaid, interest, serviceTax,
        openingPrincipal: (closing + principal).toFixed(2), closingPrincipal,
        // Broken-period interest in the first EMI is not the contractual rate.
        rate: explicitRate
      }));
    }
  };
  // Prefer complete lines; recover wrapped rows without overwriting good reads.
  String(text ?? "").split(/\n+/).forEach(addMatches);
  addMatches(source);
  return [...rows.values()].sort((first, second) => first.installment - second.installment);
}

function parseBankNameSegmentRows(normalized) {
  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})(?=\s)/g;
  const starts = [...String(normalized ?? "").matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const values = extractSignedMoneyLikeNumbers(normalized.slice(rowStart, rowEnd));
    if (values.length < 5) return null;
    const [installmentAmount, principalPaid, interest, mi, closingPrincipal] = values.slice(-5);
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      serviceTax: mi,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  }).filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0 && toNumber(row.closingPrincipal) >= 0);
}

function findDeclaredTenure(text) {
  const value = findPdfValue(text, [
    "Tenure(In Months)",
    "Tenure In Months",
    "Tenure",
    "Period in Months",
    "Number of EMIs",
    "No of Installments",
    "No. of Installments",
    "No of Instalments",
    "No. of Instalments",
    "Repayable in Instalments",
    "Repayable in Installments",
    "Total Instl",
    "Total Instalments",
    "Total Installments",
  ], { requireDigit: true, minLength: 1 });
  const number = Number(String(value ?? "").match(/\d{1,3}/)?.[0]);
  return Number.isFinite(number) && number > 0 && number <= 240 ? number : 0;
}

function isScheduleDuePaid(value) {
  const dueDate = parseDisplayDate(value);
  if (!dueDate) return false;
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  dueDate.setHours(0, 0, 0, 0);
  return dueDate <= today;
}

function parseFlattenedScheduleRows(text) {
  const normalized = String(text ?? "").replace(/\s+/g, " ");
  const naRows = parseNaScheduleRows(normalized);
  if (naRows.length > 0) return naRows;
  const rowStartPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+/g;
  const starts = [...normalized.matchAll(rowStartPattern)];
  return starts.map((match, index) => {
    const rowStart = match.index + match[0].length;
    const rowEnd = starts[index + 1]?.index ?? normalized.length;
    const values = extractMoneyLikeNumbers(normalized.slice(rowStart, rowEnd));
    return classifyScheduleValues(match[1], match[2], values);
  }).filter((row) => row?.installment > 0 && toNumber(row.installmentAmount) > 0);
}

function parseNaScheduleRows(normalized) {
  const rowPattern = /(?:^|\s)(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(?:N\/?A|NIA|NA|-|[A-Z0-9]+)\s+(?:N\/?A|NIA|NA|-|[A-Z0-9]+)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)\s+(\d[\d,]*(?:\.\d{1,2})?)(?=\s|$)/gi;
  return [...String(normalized ?? "").matchAll(rowPattern)].map((match) => {
    const installmentAmount = match[3];
    const principalPaid = match[4];
    const interest = match[5];
    const closingPrincipal = match[6];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment: match[1],
      dueDate: match[2],
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest,
      closingPrincipal,
      rate: deriveAnnualRate(interest, openingPrincipal)
    });
  }).filter((row) => row.installment > 0 && toNumber(row.installmentAmount) > 0);
}

function parseScheduleLine(line) {
  const text = String(line ?? "").replace(/\s+/g, " ").trim();
  const match = text.match(/^(\d{1,3})\s+(\d{1,2}[./-]\d{1,2}[./-]\d{2,4})\s+(.+)$/);
  if (!match) return null;
  const [, installment, dueDate, rest] = match;
  const numbers = extractMoneyLikeNumbers(rest);
  if (numbers.length < 4) return null;
  return classifyScheduleValues(installment, dueDate, numbers);
}

function classifyScheduleValues(installment, dueDate, numbers) {
  const cleanedNumbers = (numbers ?? []).map((value) => String(value).replace(/,/g, ""));
  if (cleanedNumbers.length < 4) return null;
  const maybeRate = cleanedNumbers[cleanedNumbers.length - 1];
  const hasRate = toNumber(maybeRate) > 0 && toNumber(maybeRate) <= 40;
  const moneyValues = hasRate ? cleanedNumbers.slice(0, -1) : cleanedNumbers;
  if (moneyValues.length >= 6 && toNumber(moneyValues[4]) <= 100) {
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal: moneyValues[0],
      installmentAmount: moneyValues[1],
      principalPaid: moneyValues[2],
      interest: moneyValues[3],
      serviceTax: moneyValues[4],
      closingPrincipal: moneyValues[5],
      rate: hasRate ? maybeRate : ""
    });
  }
  if (moneyValues.length >= 6 && toNumber(moneyValues[3]) <= 100 && toNumber(moneyValues[4]) > 1000) {
    const installmentAmount = moneyValues[0];
    const principalPaid = moneyValues[1];
    const finCharges = moneyValues[2];
    const insurance = moneyValues[3];
    const closingPrincipal = moneyValues[4];
    const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal,
      installmentAmount,
      principalPaid,
      interest: finCharges,
      serviceTax: insurance,
      closingPrincipal,
      rate: deriveAnnualRate(finCharges, openingPrincipal)
    });
  }
  if (moneyValues.length >= 5) {
    return makeScheduleRow({
      installment,
      dueDate,
      openingPrincipal: moneyValues[0],
      installmentAmount: moneyValues[1],
      principalPaid: moneyValues[2],
      interest: moneyValues[3],
      closingPrincipal: moneyValues[4],
      rate: hasRate ? maybeRate : ""
    });
  }
  const [installmentAmount, principalPaid, interest, closingPrincipal] = moneyValues;
  const openingPrincipal = String(toNumber(closingPrincipal) + toNumber(principalPaid));
  return makeScheduleRow({
    installment,
    dueDate,
    openingPrincipal,
    installmentAmount,
    principalPaid,
    interest,
    closingPrincipal,
    rate: deriveAnnualRate(interest, openingPrincipal)
  });
}

function makeScheduleRow(row) {
  return {
    installment: Number(row.installment),
    dueDate: row.dueDate,
    openingPrincipal: String(row.openingPrincipal ?? "").replace(/,/g, ""),
    installmentAmount: String(row.installmentAmount ?? "").replace(/,/g, ""),
    principalPaid: String(row.principalPaid ?? "").replace(/,/g, ""),
    interest: String(row.interest ?? "").replace(/,/g, ""),
    serviceTax: String(row.serviceTax ?? "").replace(/,/g, ""),
    closingPrincipal: String(row.closingPrincipal ?? "").replace(/,/g, ""),
    rate: row.rate ? String(row.rate).replace(/^0+(\d)/, "$1") : ""
  };
}

function deriveAnnualRate(interest, openingPrincipal) {
  const principal = toNumber(openingPrincipal);
  const monthlyInterest = toNumber(interest);
  if (principal <= 0 || monthlyInterest <= 0) return "";
  return ((monthlyInterest * 1200) / principal).toFixed(2);
}

function mostCommonAmount(values) {
  const counts = new Map();
  values.forEach((value) => {
    const key = String(Math.round(toNumber(value)));
    counts.set(key, (counts.get(key) ?? 0) + 1);
  });
  return [...counts.entries()]
    .sort((first, second) => second[1] - first[1] || Number(second[0]) - Number(first[0]))[0]?.[0] ?? "";
}

function plausibleBankClosingPrincipal(value, loanAmount) {
  const closing = toNumber(value);
  const loan = toNumber(loanAmount);
  if (closing < 0) return false;
  if (!closing || !loan) return true;
  // OCR sometimes joins Indian comma groups or picks another numeric column
  // (for example 25,653,712 for a 26-lakh loan). Never persist that as POS.
  return closing <= loan * 1.5;
}

function extractMoneyLikeNumbers(value) {
  return [...String(value ?? "").matchAll(/\b\d[\d,]*(?:\.\d{1,2})?\b/g)]
    .map((match) => match[0].replace(/,/g, ""))
    .filter((item) => item.length <= 12);
}

function extractSignedMoneyLikeNumbers(value) {
  return [...String(value ?? "").matchAll(/-?\b\d[\d,]*(?:\.\d{1,2})?\b/g)]
    .map((match) => match[0].replace(/,/g, ""))
    .filter((item) => item.length <= 13);
}

function findAgreementFallback(text) {
  const patterns = [
    /\b(?:CVR|LNC|LAN|AG|AGR|HF|HL|TW|TR|LN)[A-Z0-9/-]{5,35}\b/gi,
    /\b[A-Z]{2,6}\d{6,25}\b/gi,
    /\b\d{8,25}\b/g
  ];
  for (const pattern of patterns) {
    const matches = [...text.matchAll(pattern)].map((match) => match[0]);
    const candidate = matches.find((value) => normalizeAgreement(value).length >= 8);
    if (candidate) return candidate;
  }
  return "";
}

function findPdfValue(text, labels, options = {}) {
  const allowsShortNumber = labels.some((label) => /tenure|paid|install/i.test(label));
  const requireDigit = options.requireDigit ?? allowsShortNumber;
  const minLength = options.minLength ?? (allowsShortNumber ? 1 : 2);
  for (const label of labels) {
    const valuePattern = allowsShortNumber ? "([A-Z0-9][A-Z0-9/\\-]{0,40})" : "([A-Z0-9][A-Z0-9/\\-]{1,40})";
    const pattern = new RegExp(`${looseLabelPattern(label)}\\s*[:\\-]?\\s*${valuePattern}`, "gi");
    const match = text.match(pattern);
    if (match) {
      for (const item of match) {
        const value = item.replace(new RegExp(`${looseLabelPattern(label)}\\s*[:\\-]?\\s*`, "i"), "").trim();
        const clean = cleanPdfValue(value);
        if ((!requireDigit || /\d/.test(clean)) && normalizeAgreement(clean).length >= minLength) return clean;
      }
    }
  }
  return "";
}

function findPdfAmount(text, labels, options = {}) {
  const minVal = options.min ?? (options.positive ? 1 : 0);
  const maxVal = options.max ?? 100000000;
  const source = String(text ?? "");

  const cleanNumber = (str) => str.replace(/,/g, "").trim();

  const isInvalidCandidate = (valStr, numVal) => {
    if (isNaN(numVal) || numVal < minVal || numVal > maxVal) return true;
    if (/^[6-9]\d{9}$/.test(valStr)) return true;
    if (/^(202[0-9]|203[0-5])$/.test(valStr)) return true;
    return false;
  };

  for (const label of labels) {
    const escaped = looseLabelPattern(label);

    const sameLinePattern = new RegExp(`${escaped}\\s*(?:\\([^)]*\\))?\\s*[:\\-=]?\\s*(?:INR|RS\\.?|₹)?\\s*([0-9][0-9,]*(?:\\.\\d{1,2})?)`, "gi");
    for (const match of source.matchAll(sameLinePattern)) {
      const valStr = cleanNumber(match[1]);
      const num = toNumber(valStr);
      if (!isInvalidCandidate(valStr, num)) return valStr;
    }

    const tightPattern = new RegExp(`${escaped}[^\\n\\r0-9]{0,40}(?:INR|RS\\.?|₹)?\\s*([0-9][0-9,]*(?:\\.\\d{1,2})?)`, "gi");
    for (const match of source.matchAll(tightPattern)) {
      const valStr = cleanNumber(match[1]);
      const num = toNumber(valStr);
      if (!isInvalidCandidate(valStr, num)) return valStr;
    }
  }

  return "";
}

function findPdfExactAmount(text, labels) {
  for (const label of labels) {
    const pattern = new RegExp(`${looseLabelPattern(label)}\\s*[:\\-]?\\s*(?:INR|Rs\\.?)?\\s*(-?[0-9][0-9,]*(?:\\.\\d{1,2})?)`, "i");
    const match = String(text ?? "").match(pattern);
    if (match && Math.abs(toNumber(match[1])) >= 100) return match[1].replace(/,/g, "");
  }
  return "";
}

function findPdfRate(text) {
  const source = String(text ?? "");
  const apr = source.match(/\bAPR\b\s*[:\-]?\s*([0-9]+(?:\.[0-9]+)?)\s*%/i);
  if (apr) return apr[1];
  const irrIndex = source.search(/\bIRR\b/i);
  if (irrIndex >= 0) {
    const irrValues = [...source.slice(irrIndex).matchAll(/(?<![\d.])\d+\.\d{3,}(?![\d.])/g)]
      .map((match) => match[0])
      .filter((value) => toNumber(value) > 0 && toNumber(value) <= 40);
    if (irrValues.length > 0 && isTataFinanceDocument(source)) return irrValues[irrValues.length - 1];
  }
  const kfsRate = source.match(/Rate\s+of\s+Interest\s*\(\s*Sl\s*No\.\s*13[^)]*\)\s*([0-9]+(?:\.[0-9]+)?)/i);
  if (kfsRate && toNumber(kfsRate[1]) > 0) return kfsRate[1];
  if (/SUNDARAM\s+FINANCE|Details\s+of\s+Interest\s+and\s+Principal\s+Breakup/i.test(source)) return "";
  const explicit = source.match(/(?:Annualized\s+Rate\s+of\s+Interest(?:\s*\([^)]*\))?|Current\s+Interest\s*\([^)]*\)|Current\s*Int\s*Rate|Rate\s*of\s*Interest(?:\s*\([^)]*\))?|Interest\s*Rate)[^0-9]{0,35}([0-9]+(?:\.[0-9]+)?)/i);
  if (explicit && toNumber(explicit[1]) > 0) return explicit[1];
  const irrValues = irrIndex >= 0
    ? [...source.slice(irrIndex).matchAll(/(?<![\d.])\d+\.\d{3,}(?![\d.])/g)].map((match) => match[0])
    : [];
  return irrValues[irrValues.length - 1] ?? "";
}

function findPdfDate(text, labels) {
  for (const label of labels) {
    const pattern = new RegExp(`${looseLabelPattern(label)}[^0-9A-Za-z]{0,50}(\\d{1,2}[./-]\\d{1,2}[./-]\\d{2,4}|\\d{4}-\\d{1,2}-\\d{1,2}|\\d{1,2}[-\\s][A-Za-z]{3,}[-\\s]\\d{2,4})`, "i");
    const match = text.match(pattern);
    if (match) return formatDisplayDate(match[1]);
  }
  return "";
}

function looseLabelPattern(label) {
  const parts = label
    .trim()
    .split(/\s+/)
    .map((word) => word.split("").map((char) => escapeRegExp(char)).join("\\s*"));
  return `\\b${parts.join("\\s+")}\\b`;
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function cleanPdfValue(value) {
  return String(value ?? "")
    .replace(/\s{2,}.*/, "")
    .replace(/\b(?:INR|Rs|Amount|Date)\b.*$/i, "")
    .trim();
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/\"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function getReportRows(report, data) {
  if (report === "Finance") {
    return data.dueTasks.map((task) => ({
      id: task.id,
      name: `${task.type} - ${getDataClient(data, task.clientId)?.name ?? task.clientId}`,
      status: task.status,
      amount: formatMoney(task.amount),
      detail: `Due ${formatDisplayDate(task.dueDate)}`,
      date: task.dueDate
    }));
  }
  if (report === "Insurance") {
    return data.vehicles.map((vehicle) => ({
      id: vehicle.id,
      name: vehicle.regNo,
      status: vehicle.insuranceExpiry <= "2026-08-31" ? "Due" : "Active",
      amount: formatMoney(vehicle.overdue),
      detail: `Insurance expiry ${formatDisplayDate(vehicle.insuranceExpiry)}`,
      date: vehicle.insuranceExpiry
    }));
  }
  if (report === "Compliance") {
    return data.vehicles.map((vehicle) => ({
      id: vehicle.id,
      name: vehicle.regNo,
      status: vehicle.permitExpiry <= "2026-08-31" ? "Due" : "Active",
      amount: "-",
      detail: `Permit expiry ${formatDisplayDate(vehicle.permitExpiry)}`,
      date: vehicle.permitExpiry
    }));
  }
  if (report === "Caller") {
    return data.callerActivities.map((activity) => ({
      id: activity.id,
      name: getDataClient(data, data.dueTasks.find((task) => task.id === activity.taskId)?.clientId)?.name ?? activity.taskId,
      status: activity.outcome,
      amount: activity.expectedAmount || "-",
      detail: `${activity.channel} | ${activity.nextFollowUp}`,
      date: activity.at || activity.nextFollowUp
    }));
  }
  if (report === "Marketplace") {
    return data.listings.map((listing) => ({
      id: listing.id,
      name: listing.title,
      status: listing.status,
      amount: formatMoney(listing.price),
      detail: `${listing.location} | ${listing.condition}`,
      date: listing.createdAt || ""
    }));
  }
  return data.auditLogs.map((log) => ({
    id: log.id,
    name: `${log.module}: ${log.action}`,
    status: `${log.record} | ${log.oldValue} to ${log.newValue}`,
    amount: "-",
    detail: `${log.at} - ${log.remark}`,
    date: log.at
  }));
}
