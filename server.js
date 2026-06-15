import http from "node:http";
import { stat } from "node:fs/promises";
import { createReadStream } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";
import { loadLocalEnv, normalizeEnvBoolean } from "./src/env.js";
import { parseCookies, readBody, redirect, sendJson, sendNoContent } from "./src/http.js";
import { createStorage } from "./src/storage.js";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");

loadLocalEnv(__dirname);

const bundledDataDir = path.join(__dirname, "data");
const dataDir = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : bundledDataDir;
const bookingsFile = path.join(dataDir, "bookings.json");
const clientsFile = path.join(dataDir, "clients.json");
const photographersFile = path.join(dataDir, "photographers.json");
const workFile = path.join(dataDir, "work-assignments.json");
const invoicesFile = path.join(dataDir, "invoices.json");
const wagesFile = path.join(dataDir, "wages.json");
const employeeWagesFile = path.join(dataDir, "employee-wages.json");
const usersFile = path.join(dataDir, "users.json");
const sendLogsFile = path.join(dataDir, "send-logs.json");
const seedFiles = {
  bookings: path.join(bundledDataDir, "bookings.json"),
  clients: path.join(bundledDataDir, "clients.json"),
  photographers: path.join(bundledDataDir, "photographers.json"),
  work: path.join(bundledDataDir, "work-assignments.json"),
  invoices: path.join(bundledDataDir, "invoices.json"),
  wages: path.join(bundledDataDir, "wages.json"),
  employeeWages: path.join(bundledDataDir, "employee-wages.json"),
  users: path.join(bundledDataDir, "users.json"),
  sendLogs: path.join(bundledDataDir, "send-logs.json")
};
const supabaseStorage = {
  url: (process.env.SUPABASE_URL || process.env.SUPABASE_PROJECT_URL || "").replace(/\/$/, ""),
  key: process.env.SUPABASE_SERVICE_ROLE_KEY || process.env.SUPABASE_SERVICE_KEY || "",
  table: process.env.SUPABASE_TABLE || "internal_booking_data"
};
const storageBackend = supabaseStorage.url && supabaseStorage.key ? "supabase" : (process.env.DATA_DIR ? "disk" : "app");

const port = Number(process.env.PORT || 4180);
const host = process.env.HOST || (process.env.PORT ? "0.0.0.0" : "127.0.0.1");
const cancelledBookingRetentionHours = Number(process.env.CANCELLED_BOOKING_RETENTION_HOURS || 12);
const cancelledBookingRetentionMs =
  (Number.isFinite(cancelledBookingRetentionHours) && cancelledBookingRetentionHours > 0 ? cancelledBookingRetentionHours : 12) * 60 * 60 * 1000;
const configuredWorkAttachmentCount = Number(process.env.WORK_ATTACHMENT_MAX_COUNT || 6);
const configuredWorkAttachmentBytes = Number(process.env.WORK_ATTACHMENT_MAX_BYTES || 450_000);
const maxWorkAttachmentCount = Number.isFinite(configuredWorkAttachmentCount) && configuredWorkAttachmentCount > 0
  ? configuredWorkAttachmentCount
  : 6;
const maxWorkAttachmentBytes = Number.isFinite(configuredWorkAttachmentBytes) && configuredWorkAttachmentBytes > 0
  ? configuredWorkAttachmentBytes
  : 450_000;
const publicAppUrl = (process.env.APP_PUBLIC_URL || "https://system.openframe.studio").replace(/\/$/, "");
const bookingTimeZone = "Australia/Sydney";
const larkConfig = {
  appId: process.env.LARK_APP_ID || "",
  appSecret: process.env.LARK_APP_SECRET || "",
  calendarId: process.env.LARK_CALENDAR_ID || "primary",
  organizerCalendarId: process.env.LARK_ORGANIZER_CALENDAR_ID || "",
  organizerUserId: process.env.LARK_ORGANIZER_USER_ID || "",
  senderEmail: process.env.LARK_SENDER_EMAIL || "admin@openframe.studio",
  senderName: process.env.LARK_SENDER_NAME || "admin@openframe.studio",
  timezone: bookingTimeZone,
  apiBase: (process.env.LARK_API_BASE || "https://open.larksuite.com/open-apis").replace(/\/$/, "")
};
const invoiceEmailUser = process.env.INVOICE_EMAIL_USER || process.env.SMTP_USER || "admin@openframe.studio";
const invoiceEmailPort = Number(process.env.INVOICE_EMAIL_PORT || process.env.SMTP_PORT || 465);
const invoiceEmailTimeoutMs = Number(process.env.INVOICE_EMAIL_TIMEOUT_MS || 15_000);
const invoiceEmailProviderValue = String(process.env.INVOICE_EMAIL_PROVIDER || "resend").trim().toLowerCase();
const invoiceEmailProvider = ["resend", "smtp"].includes(invoiceEmailProviderValue) ? invoiceEmailProviderValue : "resend";
const invoiceEmailConfig = {
  host: process.env.INVOICE_EMAIL_HOST || process.env.SMTP_HOST || "smtp.larksuite.com",
  port: Number.isFinite(invoiceEmailPort) ? invoiceEmailPort : 587,
  secure: normalizeEnvBoolean(process.env.INVOICE_EMAIL_SECURE || process.env.SMTP_SECURE, invoiceEmailPort === 465),
  user: invoiceEmailUser,
  pass: process.env.INVOICE_EMAIL_PASSWORD || process.env.INVOICE_MAIL_PASSWORD || process.env.SMTP_PASSWORD || "",
  from: process.env.INVOICE_EMAIL_FROM || `OpenFrame Studio <${invoiceEmailUser}>`,
  replyTo: process.env.INVOICE_EMAIL_REPLY_TO || "admin@openframe.studio",
  bcc: process.env.INVOICE_EMAIL_BCC || "",
  subjectPrefix: process.env.INVOICE_EMAIL_SUBJECT_PREFIX || "Tax Invoice",
  timeoutMs: Number.isFinite(invoiceEmailTimeoutMs) ? invoiceEmailTimeoutMs : 15_000
};
const resendConfig = {
  apiKey: process.env.RESEND_API_KEY || "",
  apiBase: (process.env.RESEND_API_BASE || "https://api.resend.com").replace(/\/$/, "")
};
const workInviteEmailConfig = {
  enabled: normalizeEnvBoolean(process.env.WORK_INVITE_EMAIL_ENABLED, true),
  to: process.env.WORK_INVITE_EMAIL_TO || process.env.EMPLOYEE_EMAIL || "",
  from: process.env.WORK_INVITE_EMAIL_FROM || invoiceEmailConfig.from,
  replyTo: process.env.WORK_INVITE_EMAIL_REPLY_TO || invoiceEmailConfig.replyTo,
  bcc: process.env.WORK_INVITE_EMAIL_BCC || "",
  subjectPrefix: process.env.WORK_INVITE_EMAIL_SUBJECT_PREFIX || "Work invite",
  timeoutMs: Number.isFinite(invoiceEmailTimeoutMs) ? invoiceEmailTimeoutMs : 15_000
};
const workLarkNotificationConfig = {
  enabled: normalizeEnvBoolean(process.env.WORK_LARK_NOTIFICATIONS_ENABLED, true),
  receiveIdType: process.env.WORK_LARK_RECEIVE_ID_TYPE || "email",
  receiveId: process.env.WORK_LARK_RECEIVE_ID || process.env.EMPLOYEE_LARK_RECEIVE_ID || "",
  subjectPrefix: process.env.WORK_LARK_NOTIFICATION_PREFIX || "New work"
};
const workCompletionLarkNotificationConfig = {
  enabled: normalizeEnvBoolean(process.env.WORK_COMPLETION_LARK_NOTIFICATIONS_ENABLED, true),
  receiveIdType: process.env.WORK_COMPLETION_LARK_RECEIVE_ID_TYPE || process.env.BOSS_LARK_RECEIVE_ID_TYPE || "email",
  receiveId: process.env.WORK_COMPLETION_LARK_RECEIVE_ID
    || process.env.BOSS_LARK_RECEIVE_ID
    || process.env.ADMIN_LARK_RECEIVE_ID
    || "barry.gao@openframe.studio",
  subjectPrefix: process.env.WORK_COMPLETION_LARK_NOTIFICATION_PREFIX || "Work finished"
};
const workCompletionEmailConfig = {
  enabled: normalizeEnvBoolean(process.env.WORK_COMPLETION_EMAIL_ENABLED, true),
  to: process.env.WORK_COMPLETION_EMAIL_TO || process.env.BOSS_EMAIL || "barry.gao@openframe.studio",
  from: process.env.WORK_COMPLETION_EMAIL_FROM || invoiceEmailConfig.from,
  replyTo: process.env.WORK_COMPLETION_EMAIL_REPLY_TO || invoiceEmailConfig.replyTo,
  bcc: process.env.WORK_COMPLETION_EMAIL_BCC || "",
  subjectPrefix: process.env.WORK_COMPLETION_EMAIL_SUBJECT_PREFIX || "Work finished",
  timeoutMs: Number.isFinite(invoiceEmailTimeoutMs) ? invoiceEmailTimeoutMs : 15_000
};
const calendarInviteEmailConfig = {
  enabled: normalizeEnvBoolean(process.env.CALENDAR_INVITE_EMAIL_ENABLED, true),
  from: process.env.CALENDAR_INVITE_EMAIL_FROM || invoiceEmailConfig.from,
  replyTo: process.env.CALENDAR_INVITE_EMAIL_REPLY_TO || invoiceEmailConfig.replyTo,
  bcc: process.env.CALENDAR_INVITE_EMAIL_BCC || "",
  subjectPrefix: process.env.CALENDAR_INVITE_EMAIL_SUBJECT_PREFIX || "Booking invitation",
  timeoutMs: Number.isFinite(invoiceEmailTimeoutMs) ? invoiceEmailTimeoutMs : 15_000
};
const authConfig = {
  username: process.env.ADMIN_USERNAME || "ShuhanGao",
  password: process.env.ADMIN_PASSWORD || "Sg1654723576"
};
const authUsers = [
  {
    username: authConfig.username,
    password: authConfig.password,
    role: "boss",
    label: "Boss / Team Leader",
    name: "Boss",
    apps: ["bookings", "clients", "photographers", "work", "invoices"],
    permissions: ["manage_bookings", "manage_directory", "manage_work", "sync_work_bookings", "view_work_messages", "manage_invoices", "manage_wages"]
  },
  {
    username: process.env.EMPLOYEE_USERNAME || "Faye",
    password: process.env.EMPLOYEE_PASSWORD || "1111",
    role: "employee",
    label: "Employee",
    name: "Faye",
    employeeId: "faye",
    apps: ["bookings", "work"],
    permissions: ["complete_work"]
  },
  {
    username: process.env.TEST_EMPLOYEE_USERNAME || "Test",
    password: process.env.TEST_EMPLOYEE_PASSWORD || "1111",
    role: "employee",
    label: "Employee",
    name: "Test",
    employeeId: "test",
    apps: ["bookings", "work"],
    permissions: ["complete_work"]
  }
];
const accountRoleOptions = [
  {
    value: "boss",
    label: "Boss / Team Leader",
    apps: ["bookings", "clients", "photographers", "work", "invoices"],
    permissions: ["manage_bookings", "manage_directory", "manage_work", "sync_work_bookings", "view_work_messages", "manage_invoices", "manage_wages"]
  },
  {
    value: "employee",
    label: "Employee",
    apps: ["bookings", "work"],
    permissions: ["complete_work"]
  }
];
const workDeskOrigins = new Set(
  (process.env.WORK_DESK_ORIGINS || "http://127.0.0.1:4173,http://localhost:4173")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean)
);
const sessionSigningSecret = process.env.SESSION_SECRET
  || process.env.AUTH_SECRET
  || process.env.SUPABASE_SERVICE_ROLE_KEY
  || larkConfig.appSecret
  || resendConfig.apiKey
  || invoiceEmailConfig.pass
  || authConfig.password;
const larkMessageReceiveIdTypes = new Set(["open_id", "user_id", "union_id", "email", "chat_id"]);
const larkInvalidReceiveIdCode = 230034;

let tenantTokenCache = null;
const sessions = new Map();
const sessionCookieName = "internalbooking_session";
const sessionCookieVersion = "v1";
const sessionMaxAgeSeconds = 30 * 24 * 60 * 60;
const addressSuggestionCache = new Map();
let lastAddressLookupAt = 0;

const contentTypes = {
  ".html": "text/html; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".webmanifest": "application/manifest+json; charset=utf-8",
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon"
};

const serviceCatalog = new Set(["Photography", "Floorplan", "Drone", "Siteplan"]);
const invoiceServicePrices = {
  Photography: 150,
  Floorplan: 75,
  Drone: 100,
  Siteplan: 25
};
const invoiceGstRate = 0.1;
const invoiceCurrency = "AUD";
const invoicePrefixOverrides = [
  { prefix: "9AMC", names: ["9AM Group", "9AM Production"] },
  { prefix: "BBS", names: ["Boteli Business Service"] },
  { prefix: "CP", names: ["Consortium Property"] },
  { prefix: "DY", names: ["D & Y Group"] },
  { prefix: "ES", names: ["Esery Studio"] },
  { prefix: "FS", names: ["Flipt Studio"] },
  { prefix: "HP", names: ["Helium Property"] },
  { prefix: "ID", names: ["Impact Displays"] },
  { prefix: "JD", names: ["Julie Dang"] },
  { prefix: "KG", names: ["KozyGuru", "The Guru Hub"] },
  { prefix: "KRE", names: ["Khy Real Estate"] },
  { prefix: "LM", names: ["Landmark Creations"] },
  { prefix: "MB", names: ["McConnell Bourn"] },
  { prefix: "MP", names: ["Masterpiece Holiday", "Masterpiece Homestay"] },
  { prefix: "MS", names: ["MS image"] },
  { prefix: "MW", names: ["Metawise BnB"] },
  { prefix: "NC", names: ["Nguyen Cameras"] },
  { prefix: "PA", names: ["Pico Australia"] },
  { prefix: "RTW", names: ["RTW Properties"] },
  { prefix: "SC", names: ["Sprite Clean"] },
  { prefix: "VE", names: ["Video Estate", "VideoEstate"] }
];
const wageCurrency = "AUD";
const employeeWageCurrency = "THB";
const larkImportDays = Number(process.env.LARK_IMPORT_DAYS || 120);
const defaultPhotographers = [
  {
    id: "default-barry",
    name: "Barry",
    email: process.env.DEFAULT_PHOTOGRAPHER_EMAIL || "",
    phone: "0403 007 853",
    gstIncluded: false,
    createdAt: "2026-05-16T00:00:00.000+10:00",
    updatedAt: "2026-05-16T00:00:00.000+10:00"
  }
];
const defaultWorkEmployee = {
  id: "faye",
  name: "Faye",
  email: parseGuestEmails(workInviteEmailConfig.to)[0] || process.env.FAYE_EMAIL || "faye.w@openframe.studio",
  larkReceiveId: workLarkNotificationConfig.receiveId,
  larkReceiveIdType: workLarkNotificationConfig.receiveIdType,
  role: "Editor / Admin",
  availability: "Mon-Fri, 12pm-8pm Australian time"
};
const testWorkEmployee = {
  id: "test",
  name: "Test",
  email: process.env.TEST_EMPLOYEE_EMAIL || "",
  larkReceiveId: "",
  larkReceiveIdType: "email",
  role: "Test Employee",
  availability: "Testing account"
};
const defaultWorkEmployees = [defaultWorkEmployee, testWorkEmployee];
const defaultWorkState = {
  employee: defaultWorkEmployee,
  employees: defaultWorkEmployees,
  assignments: [],
  messages: []
};
const defaultEmployeeWages = [
  {
    id: "employee-faye-monthly",
    wageNumber: "E001",
    employeeName: "Faye",
    employmentType: "full_time",
    payPeriod: "monthly",
    amount: 15000,
    currency: employeeWageCurrency,
    status: "draft",
    issuedAt: "2026-05-26T00:00:00.000+10:00",
    paymentSchedule: "last_day_of_month",
    nextPaymentAt: "2026-05-31T12:00:00.000Z",
    paidAt: "",
    voidedAt: "",
    notes: "Monthly wage for Faye. Payment is due on the last day of every month.",
    createdAt: "2026-05-26T00:00:00.000+10:00",
    updatedAt: "2026-05-26T00:00:00.000+10:00"
  }
];
const dataFiles = {
  bookings: {
    file: bookingsFile,
    seedFile: seedFiles.bookings,
    supabaseKey: "bookings",
    fallback: []
  },
  clients: {
    file: clientsFile,
    seedFile: seedFiles.clients,
    supabaseKey: "clients",
    fallback: []
  },
  photographers: {
    file: photographersFile,
    seedFile: seedFiles.photographers,
    supabaseKey: "photographers",
    fallback: defaultPhotographers
  },
  work: {
    file: workFile,
    seedFile: seedFiles.work,
    supabaseKey: "work",
    fallback: defaultWorkState
  },
  invoices: {
    file: invoicesFile,
    seedFile: seedFiles.invoices,
    supabaseKey: "invoices",
    fallback: []
  },
  wages: {
    file: wagesFile,
    seedFile: seedFiles.wages,
    supabaseKey: "wages",
    fallback: []
  },
  employeeWages: {
    file: employeeWagesFile,
    seedFile: seedFiles.employeeWages,
    supabaseKey: "employee-wages",
    fallback: defaultEmployeeWages
  },
  users: {
    file: usersFile,
    seedFile: seedFiles.users,
    supabaseKey: "users",
    fallback: []
  },
  sendLogs: {
    file: sendLogsFile,
    seedFile: seedFiles.sendLogs,
    supabaseKey: "send_logs",
    fallback: []
  }
};

const { prepareDataStorage, readStoredJson, writeStoredJson } = createStorage({
  dataDir,
  dataFiles,
  storageBackend,
  supabaseStorage
});

function applyCorsHeaders(req, res) {
  const origin = req.headers.origin || "";
  if (!workDeskOrigins.has(origin)) {
    return;
  }

  res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Access-Control-Allow-Credentials", "true");
  res.setHeader("Access-Control-Allow-Methods", "GET,POST,PUT,PATCH,DELETE,OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type,Accept");
  res.setHeader("Vary", "Origin");
}

function isSecureRequest(req) {
  return req.headers["x-forwarded-proto"] === "https" || Boolean(req.socket.encrypted);
}

function buildSessionCookie(req, token, maxAgeSeconds = sessionMaxAgeSeconds) {
  const parts = [
    `${sessionCookieName}=${encodeURIComponent(token)}`,
    "HttpOnly",
    "Path=/",
    "SameSite=Lax",
    `Max-Age=${maxAgeSeconds}`
  ];

  if (isSecureRequest(req)) {
    parts.push("Secure");
  }

  return parts.join("; ");
}

function clearSessionCookie(req) {
  return buildSessionCookie(req, "", 0);
}

function getSessionToken(req) {
  return parseCookies(req.headers.cookie || "")[sessionCookieName] || "";
}

function cleanupExpiredSessions() {
  const now = Date.now();
  for (const [token, session] of sessions.entries()) {
    if (session.expiresAt <= now) {
      sessions.delete(token);
    }
  }
}

function signSessionMessage(message) {
  return crypto
    .createHmac("sha256", sessionSigningSecret)
    .update(message)
    .digest("base64url");
}

function safeEqualsString(left, right) {
  const leftBuffer = Buffer.from(String(left || ""));
  const rightBuffer = Buffer.from(String(right || ""));
  return leftBuffer.length === rightBuffer.length && crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function accountRoleOption(value) {
  return accountRoleOptions.find((option) => option.value === value) || null;
}

function normalizeAccountRole(value) {
  return accountRoleOption(String(value || "").trim())?.value || "";
}

function canChangeAccountRole(user) {
  return (user?.baseRole || user?.role) === "boss";
}

function userWithAccountRole(user, roleValue = "") {
  const baseRole = user?.baseRole || user?.role || "";
  const roleMode = canChangeAccountRole({ ...user, baseRole })
    ? (normalizeAccountRole(roleValue) || normalizeAccountRole(user?.roleMode) || normalizeAccountRole(baseRole) || baseRole)
    : baseRole;
  const roleOption = accountRoleOption(roleMode) || accountRoleOption(baseRole);

  if (!roleOption) {
    return {
      ...user,
      baseRole,
      roleMode: baseRole,
      canChangeRole: false,
      roleOptions: []
    };
  }

  const roleCanChange = canChangeAccountRole({ ...user, baseRole });
  const usesBaseRole = roleOption.value === baseRole;

  return {
    ...user,
    name: usesBaseRole ? (user.name || user.username) : roleOption.label,
    role: roleOption.value,
    label: roleOption.label,
    baseRole,
    roleMode: roleOption.value,
    canChangeRole: roleCanChange,
    roleOptions: roleCanChange
      ? accountRoleOptions.map(({ value, label }) => ({ value, label }))
      : [],
    employeeId: roleOption.value === "employee" ? (user.employeeId || defaultWorkEmployee.id) : (user.employeeId || ""),
    apps: [...roleOption.apps],
    permissions: [...roleOption.permissions]
  };
}

function sessionUserPayload(user) {
  return {
    username: user.username,
    role: user.role,
    baseRole: user.baseRole || user.role,
    roleMode: user.roleMode || user.role,
    label: user.label,
    name: user.name || user.username,
    employeeId: user.employeeId || "",
    apps: Array.isArray(user.apps) ? user.apps : [],
    permissions: Array.isArray(user.permissions) ? user.permissions : [],
    canChangeRole: Boolean(user.canChangeRole),
    roleOptions: Array.isArray(user.roleOptions) ? user.roleOptions : []
  };
}

function createSignedSessionToken(user) {
  const baseRole = user?.baseRole || user?.role || "";
  const roleMode = normalizeAccountRole(user?.roleMode);
  const payload = {
    exp: Date.now() + sessionMaxAgeSeconds * 1000,
    username: user.username
  };

  if (roleMode && roleMode !== baseRole) {
    payload.roleMode = roleMode;
  }

  const encodedPayload = Buffer.from(JSON.stringify(payload)).toString("base64url");
  const message = `${sessionCookieVersion}.${encodedPayload}`;
  return `${message}.${signSessionMessage(message)}`;
}

function readSignedSession(token) {
  const [version, encodedPayload, signature, extra] = String(token || "").split(".");
  if (extra || version !== sessionCookieVersion || !encodedPayload || !signature) {
    return null;
  }

  const message = `${version}.${encodedPayload}`;
  if (!safeEqualsString(signature, signSessionMessage(message))) {
    return null;
  }

  let payload = null;
  try {
    payload = JSON.parse(Buffer.from(encodedPayload, "base64url").toString("utf8"));
  } catch {
    return null;
  }

  if (!payload?.username || !Number.isFinite(Number(payload.exp)) || Number(payload.exp) <= Date.now()) {
    return null;
  }

  const user = baseAuthUser(payload.username);
  if (!user) {
    return null;
  }

  return {
    expiresAt: Number(payload.exp),
    user: sessionUserPayload(userWithAccountRole(user, payload.roleMode))
  };
}

function createSession(user) {
  return createSignedSessionToken(user);
}

function currentSession(req) {
  const token = getSessionToken(req);
  const signedSession = readSignedSession(token);
  if (signedSession) {
    return signedSession;
  }

  cleanupExpiredSessions();
  const session = sessions.get(token);
  if (!session) {
    return null;
  }

  if (session.expiresAt <= Date.now()) {
    sessions.delete(token);
    return null;
  }

  return session;
}

function currentUser(req) {
  return currentSession(req)?.user || null;
}

function isAuthenticated(req) {
  return Boolean(currentSession(req));
}

function isBoss(req) {
  return currentUser(req)?.role === "boss";
}

function userCanAccessApp(user, app) {
  return Boolean(user?.apps?.includes(app));
}

function canAccessApp(req, app) {
  return userCanAccessApp(currentUser(req), app);
}

function hasPermission(req, permission) {
  return Boolean(currentUser(req)?.permissions?.includes(permission));
}

function requirePermission(req, res, permission, message = "You do not have access to that feature.") {
  if (hasPermission(req, permission)) {
    return true;
  }

  sendForbidden(res, message);
  return false;
}

function canCompleteWorkAssignment(user, assignment) {
  return user?.role === "boss" || assignment.employeeId === user?.employeeId;
}

function sendForbidden(res, message = "You do not have access to that feature.") {
  sendJson(res, 403, { errors: [message] });
}

function appForRoute(pathname) {
  if (pathname === "/bookings" || pathname === "/index.html") return "bookings";
  if (pathname === "/clients") return "clients";
  if (pathname === "/photographers") return "photographers";
  if (pathname === "/work" || pathname === "/work/") return "work";
  if (pathname === "/invoices") return "invoices";
  if (pathname === "/wages") return "wages";
  return "";
}

function safeEquals(left, right) {
  const leftBuffer = Buffer.from(String(left || ""));
  const rightBuffer = Buffer.from(String(right || ""));
  return leftBuffer.length === rightBuffer.length && crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function hashPassword(password) {
  const passwordSalt = crypto.randomBytes(16).toString("hex");
  const passwordHash = crypto.scryptSync(String(password || ""), passwordSalt, 64).toString("hex");
  return { passwordHash, passwordSalt };
}

function verifyHashedPassword(password, passwordSalt, passwordHash) {
  if (!passwordSalt || !passwordHash) return false;

  const calculated = crypto.scryptSync(String(password || ""), passwordSalt, 64);
  const saved = Buffer.from(String(passwordHash || ""), "hex");
  return saved.length === calculated.length && crypto.timingSafeEqual(saved, calculated);
}

function verifyAuthPassword(user, password) {
  if (user?.passwordHash && user?.passwordSalt) {
    return verifyHashedPassword(password, user.passwordSalt, user.passwordHash);
  }

  return safeEquals(password, user?.password);
}

function normalizePasswordRecords(raw) {
  const records = Array.isArray(raw) ? raw : (Array.isArray(raw?.users) ? raw.users : []);
  return records
    .map((record) => ({
      username: String(record?.username || "").trim(),
      passwordHash: String(record?.passwordHash || ""),
      passwordSalt: String(record?.passwordSalt || ""),
      updatedAt: String(record?.updatedAt || "")
    }))
    .filter((record) => record.username && record.passwordHash && record.passwordSalt);
}

async function loadPasswordRecords() {
  return normalizePasswordRecords(await readStoredJson(dataFiles.users));
}

async function savePasswordRecords(records) {
  await writeStoredJson(dataFiles.users, normalizePasswordRecords(records));
}

function baseAuthUser(username) {
  return authUsers.find((user) => safeEquals(username, user.username)) || null;
}

function passwordRecordFor(records, username) {
  return records.find((record) => safeEquals(record.username, username)) || null;
}

async function effectiveAuthUser(username) {
  const user = baseAuthUser(username);
  if (!user) return null;

  const records = await loadPasswordRecords();
  const savedPassword = passwordRecordFor(records, user.username);
  return savedPassword ? { ...user, ...savedPassword } : user;
}

async function findAuthUser(input) {
  const user = await effectiveAuthUser(String(input?.username || ""));
  if (!user || !verifyAuthPassword(user, input?.password)) {
    return null;
  }

  return user;
}

async function updateAuthUserPassword(username, newPassword) {
  const user = baseAuthUser(username);
  if (!user) return false;

  const records = await loadPasswordRecords();
  const nextRecord = {
    username: user.username,
    ...hashPassword(newPassword),
    updatedAt: new Date().toISOString()
  };
  const index = records.findIndex((record) => safeEquals(record.username, user.username));
  if (index === -1) {
    records.push(nextRecord);
  } else {
    records[index] = nextRecord;
  }
  await savePasswordRecords(records);
  return true;
}

function expireOtherSessionsForUser(username, currentToken) {
  for (const [token, session] of sessions.entries()) {
    if (token !== currentToken && safeEquals(session?.user?.username, username)) {
      sessions.delete(token);
    }
  }
}

async function loadBookings() {
  try {
    const bookings = await readStoredJson(dataFiles.bookings);
    const prunedBookings = pruneExpiredCancelledBookings(bookings);
    if (prunedBookings.length !== bookings.length) {
      await saveBookings(prunedBookings);
    }
    return prunedBookings;
  } catch (error) {
    throw error;
  }
}

async function saveBookings(bookings) {
  await writeStoredJson(dataFiles.bookings, bookings);
}

function canAutoRemoveCancelledBooking(booking, now = Date.now()) {
  if (booking?.status !== "cancelled") return false;
  if (booking.larkEventId && booking.larkStatus !== "cancelled") return false;

  const referenceTime = new Date(booking.cancelledAt || booking.updatedAt || booking.createdAt || booking.endAt || booking.startAt).getTime();
  return Number.isFinite(referenceTime) && now - referenceTime >= cancelledBookingRetentionMs;
}

function pruneExpiredCancelledBookings(bookings) {
  const now = Date.now();
  return bookings.filter((booking) => !canAutoRemoveCancelledBooking(booking, now));
}

async function loadClients() {
  try {
    return await readStoredJson(dataFiles.clients);
  } catch (error) {
    throw error;
  }
}

async function saveClients(clients) {
  await writeStoredJson(dataFiles.clients, clients);
}

function normalizePhotographer(photographer) {
  return {
    id: String(photographer?.id || crypto.randomUUID()),
    name: String(photographer?.name || photographer?.photographerName || "").trim(),
    email: String(photographer?.email || photographer?.photographerEmail || "").trim(),
    phone: String(photographer?.phone || photographer?.photographerPhone || "").trim(),
    gstIncluded: normalizeEnvBoolean(photographer?.gstIncluded ?? photographer?.photographerGstIncluded, false),
    createdAt: photographer?.createdAt || new Date().toISOString(),
    updatedAt: photographer?.updatedAt || photographer?.createdAt || new Date().toISOString()
  };
}

function normalizePhotographers(photographers) {
  return Array.isArray(photographers)
    ? photographers.map(normalizePhotographer).filter((photographer) => photographer.name)
    : [];
}

async function loadPhotographers() {
  try {
    const photographers = await readStoredJson(dataFiles.photographers);
    return normalizePhotographers(photographers);
  } catch (error) {
    throw error;
  }
}

async function savePhotographers(photographers) {
  await writeStoredJson(dataFiles.photographers, normalizePhotographers(photographers));
}

async function loadInvoices() {
  const invoices = await readStoredJson(dataFiles.invoices);
  return normalizeInvoices(invoices);
}

async function saveInvoices(invoices) {
  await writeStoredJson(dataFiles.invoices, normalizeInvoices(invoices));
}

async function loadWages() {
  const wages = await readStoredJson(dataFiles.wages);
  return normalizeWages(wages);
}

async function saveWages(wages) {
  await writeStoredJson(dataFiles.wages, normalizeWages(wages));
}

async function loadEmployeeWages() {
  const employeeWages = await readStoredJson(dataFiles.employeeWages);
  return normalizeEmployeeWages(employeeWages);
}

async function saveEmployeeWages(employeeWages) {
  await writeStoredJson(dataFiles.employeeWages, normalizeEmployeeWages(employeeWages));
}

function normalizeSendLog(log = {}) {
  const type = ["invoice", "wage", "calendar_invite", "work_notification", "work_email"].includes(log.type)
    ? log.type
    : "send";
  const status = ["success", "failed", "skipped"].includes(log.status) ? log.status : "failed";
  const recipients = Array.isArray(log.recipients)
    ? uniqueEmails(log.recipients.map((email) => String(email || "").trim()).filter(Boolean))
    : uniqueEmails(parseGuestEmails(log.recipients || ""));

  return {
    id: String(log.id || crypto.randomUUID()),
    type,
    status,
    title: String(log.title || "").trim().slice(0, 160),
    detail: String(log.detail || "").trim().slice(0, 240),
    provider: String(log.provider || "").trim().slice(0, 80),
    providerMessageId: String(log.providerMessageId || "").trim().slice(0, 160),
    from: String(log.from || "").trim().slice(0, 160),
    recipients,
    relatedId: String(log.relatedId || "").trim().slice(0, 120),
    relatedNumber: String(log.relatedNumber || "").trim().slice(0, 80),
    error: String(log.error || "").trim().slice(0, 700),
    createdAt: log.createdAt || new Date().toISOString()
  };
}

function normalizeSendLogs(logs) {
  return Array.isArray(logs)
    ? logs.map(normalizeSendLog).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt)).slice(0, 200)
    : [];
}

async function loadSendLogs() {
  return normalizeSendLogs(await readStoredJson(dataFiles.sendLogs));
}

async function saveSendLogs(logs) {
  await writeStoredJson(dataFiles.sendLogs, normalizeSendLogs(logs));
}

async function appendSendLogs(inputs) {
  try {
    const nextLogs = (Array.isArray(inputs) ? inputs : [inputs])
      .filter(Boolean)
      .map((input) => normalizeSendLog({
        ...input,
        id: input?.id || crypto.randomUUID(),
        createdAt: input?.createdAt || new Date().toISOString()
      }));

    if (!nextLogs.length) {
      return [];
    }

    const logs = await loadSendLogs();
    await saveSendLogs([...nextLogs, ...logs]);
    return nextLogs;
  } catch (error) {
    console.warn("Could not save send log:", error.message || error);
    return [];
  }
}

async function appendSendLog(input) {
  const [log] = await appendSendLogs([input]);
  return log || null;
}

function canViewSendLogs(req) {
  return ["manage_invoices", "manage_wages", "manage_bookings", "manage_work"].some((permission) => hasPermission(req, permission));
}

function normalizeWorkEmployee(employee = {}, fallback = {}) {
  const id = String(employee.id || fallback.id || defaultWorkEmployee.id).trim().toLowerCase().replace(/[^a-z0-9_-]+/g, "-");
  const configuredFayeReceiveId = String(workLarkNotificationConfig.receiveId || "").trim();
  const configuredFayeReceiveIdType = normalizeLarkMessageReceiveIdType(workLarkNotificationConfig.receiveIdType);
  const rawReceiveId = String(employee.larkReceiveId || fallback.larkReceiveId || "").trim();
  const rawReceiveIdType = normalizeLarkMessageReceiveIdType(employee.larkReceiveIdType || fallback.larkReceiveIdType || "email");
  const shouldUseConfiguredFayeLarkId =
    id === defaultWorkEmployee.id
    && configuredFayeReceiveId
    && configuredFayeReceiveIdType !== "email"
    && (rawReceiveIdType === "email" || !rawReceiveId || isValidEmail(rawReceiveId));

  return {
    ...fallback,
    ...employee,
    id: id || defaultWorkEmployee.id,
    name: String(employee.name || fallback.name || "Employee").trim(),
    email: String(employee.email || fallback.email || "").trim(),
    larkReceiveId: shouldUseConfiguredFayeLarkId ? configuredFayeReceiveId : rawReceiveId,
    larkReceiveIdType: shouldUseConfiguredFayeLarkId ? configuredFayeReceiveIdType : rawReceiveIdType,
    role: String(employee.role || fallback.role || "").trim(),
    availability: String(employee.availability || fallback.availability || "").trim()
  };
}

function normalizeWorkEmployees(raw) {
  const employees = new Map();
  const candidates = [
    ...defaultWorkEmployees,
    raw?.employee || null,
    ...(Array.isArray(raw?.employees) ? raw.employees : [])
  ].filter(Boolean);

  for (const candidate of candidates) {
    const previous = employees.get(String(candidate?.id || "").trim().toLowerCase());
    const employee = normalizeWorkEmployee(candidate, previous || {});
    if (!employee.id) {
      continue;
    }

    employees.set(employee.id, employee);
  }

  return employees.size ? [...employees.values()] : defaultWorkEmployees.map((employee) => normalizeWorkEmployee(employee));
}

function workEmployeeById(workState, employeeId) {
  const employees = Array.isArray(workState?.employees) && workState.employees.length
    ? workState.employees
    : [workState?.employee || defaultWorkEmployee];
  return employees.find((employee) => employee.id === employeeId) || employees[0] || defaultWorkEmployee;
}

function workAttachmentByteSize(dataUrl) {
  const base64 = String(dataUrl || "").split(",")[1] || "";
  return Math.ceil(base64.length * 3 / 4);
}

function normalizeWorkAttachment(attachment) {
  const dataUrl = String(attachment?.dataUrl || "").trim();
  const match = dataUrl.match(/^data:(image\/(?:jpeg|jpg|png|webp));base64,([a-z0-9+/=]+)$/i);
  if (!match) return null;

  const type = match[1].toLowerCase() === "image/jpg" ? "image/jpeg" : match[1].toLowerCase();
  const estimatedSize = workAttachmentByteSize(dataUrl);
  if (estimatedSize > maxWorkAttachmentBytes) return null;
  const providedSize = Number(attachment?.size || estimatedSize);

  return {
    id: String(attachment?.id || crypto.randomUUID()),
    name: String(attachment?.name || "Work photo").trim().slice(0, 120) || "Work photo",
    type,
    size: Number.isFinite(providedSize) ? Math.max(0, Math.min(providedSize, maxWorkAttachmentBytes)) : estimatedSize,
    dataUrl,
    createdAt: attachment?.createdAt || new Date().toISOString()
  };
}

function normalizeWorkAttachments(attachments) {
  return Array.isArray(attachments)
    ? attachments.slice(0, maxWorkAttachmentCount).map(normalizeWorkAttachment).filter(Boolean)
    : [];
}

function normalizeWorkState(raw) {
  const employees = normalizeWorkEmployees(raw);
  const employeeIds = new Set(employees.map((employee) => employee.id));
  const employee = employees[0];

  const assignments = Array.isArray(raw?.assignments)
    ? raw.assignments.map((assignment) => ({
        id: assignment.id || crypto.randomUUID(),
        employeeId: employeeIds.has(assignment.employeeId) ? assignment.employeeId : defaultWorkEmployee.id,
        title: String(assignment.title || "Untitled work"),
        dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(assignment.dueDate || ""))
          ? assignment.dueDate
          : toDateValue(new Date()),
        priority: ["high", "normal", "low"].includes(assignment.priority) ? assignment.priority : "normal",
        notes: String(assignment.notes || ""),
        status: assignment.status === "done" ? "done" : "open",
        source: String(assignment.source || ""),
        sourceId: String(assignment.sourceId || ""),
        attachments: normalizeWorkAttachments(assignment.attachments),
        inviteStatus: ["sent", "failed", "not_configured"].includes(assignment.inviteStatus) ? assignment.inviteStatus : "",
        inviteSentAt: assignment.inviteSentAt || "",
        inviteTo: uniqueEmails(parseGuestEmails(assignment.inviteTo || "")).join(", "),
        inviteEmailFrom: String(assignment.inviteEmailFrom || ""),
        inviteError: String(assignment.inviteError || ""),
        larkNotifyStatus: ["sent", "failed", "not_configured"].includes(assignment.larkNotifyStatus) ? assignment.larkNotifyStatus : "",
        larkNotifySentAt: assignment.larkNotifySentAt || "",
        larkNotifyTo: String(assignment.larkNotifyTo || ""),
        larkNotifyError: String(assignment.larkNotifyError || ""),
        completionNotifyStatus: ["sent", "failed", "not_configured"].includes(assignment.completionNotifyStatus) ? assignment.completionNotifyStatus : "",
        completionNotifySentAt: assignment.completionNotifySentAt || "",
        completionNotifyTo: String(assignment.completionNotifyTo || ""),
        completionNotifyError: String(assignment.completionNotifyError || ""),
        completionEmailStatus: ["sent", "failed", "not_configured"].includes(assignment.completionEmailStatus) ? assignment.completionEmailStatus : "",
        completionEmailSentAt: assignment.completionEmailSentAt || "",
        completionEmailTo: uniqueEmails(parseGuestEmails(assignment.completionEmailTo || "")).join(", "),
        completionEmailFrom: String(assignment.completionEmailFrom || ""),
        completionEmailError: String(assignment.completionEmailError || ""),
        completedAt: assignment.completedAt || "",
        createdAt: assignment.createdAt || new Date().toISOString(),
        updatedAt: assignment.updatedAt || assignment.createdAt || new Date().toISOString()
      }))
    : [];

  const messages = Array.isArray(raw?.messages) ? raw.messages : [];

  return { employee, employees, assignments, messages };
}

async function loadWorkState() {
  return normalizeWorkState(await readStoredJson(dataFiles.work));
}

async function saveWorkState(workState) {
  await writeStoredJson(dataFiles.work, normalizeWorkState(workState));
}

function visibleWorkStateForUser(workState, user) {
  if (user?.role === "boss") {
    return workState;
  }

  const employee = workEmployeeById(workState, user?.employeeId);

  return {
    ...workState,
    employee,
    employees: [employee],
    assignments: workState.assignments.filter((assignment) => assignment.employeeId === user?.employeeId),
    messages: []
  };
}

function timeZoneDateParts(date, timeZone = bookingTimeZone) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat("en-AU", {
    timeZone,
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hourCycle: "h23"
  }).formatToParts(date).map((part) => [part.type, part.value]));

  return {
    year: Number(parts.year),
    month: Number(parts.month),
    day: Number(parts.day),
    hour: Number(parts.hour),
    minute: Number(parts.minute),
    second: Number(parts.second)
  };
}

function parseDateParts(value) {
  const [year, month, day] = String(value || "").split("-").map(Number);
  return Number.isInteger(year) && Number.isInteger(month) && Number.isInteger(day)
    ? { year, month, day }
    : null;
}

function parseTimeParts(value) {
  const [hour, minute] = String(value || "").split(":").map(Number);
  return Number.isInteger(hour) && Number.isInteger(minute)
    ? { hour, minute }
    : null;
}

function dateValueFromTimeZone(date, timeZone = bookingTimeZone) {
  const parts = timeZoneDateParts(date, timeZone);
  return `${parts.year}-${String(parts.month).padStart(2, "0")}-${String(parts.day).padStart(2, "0")}`;
}

function zonedDateTimeToDate(dateValue, timeValue, timeZone = bookingTimeZone) {
  const dateParts = parseDateParts(dateValue);
  const timeParts = parseTimeParts(timeValue);
  if (!dateParts || !timeParts) return null;

  const targetUtc = Date.UTC(dateParts.year, dateParts.month - 1, dateParts.day, timeParts.hour, timeParts.minute, 0);
  let utc = targetUtc;
  for (let index = 0; index < 3; index += 1) {
    const parts = timeZoneDateParts(new Date(utc), timeZone);
    const displayedUtc = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second || 0);
    utc -= displayedUtc - targetUtc;
  }

  const result = new Date(utc);
  const check = timeZoneDateParts(result, timeZone);
  const matches = check.year === dateParts.year
    && check.month === dateParts.month
    && check.day === dateParts.day
    && check.hour === timeParts.hour
    && check.minute === timeParts.minute;

  return matches ? result : null;
}

function toDateValue(date) {
  return dateValueFromTimeZone(date, bookingTimeZone);
}

function parseDateValue(value) {
  const parts = parseDateParts(value);
  return parts ? new Date(parts.year, parts.month - 1, parts.day) : new Date(NaN);
}

function validateWorkAssignment(input, existing = null, workState = null) {
  const title = String(input.title || "").trim();
  const dueDate = String(input.dueDate || "").trim();
  const priority = ["high", "normal", "low"].includes(input.priority) ? input.priority : "normal";
  const notes = String(input.notes || "").trim();
  const rawAttachments = Object.prototype.hasOwnProperty.call(input, "attachments")
    ? input.attachments
    : existing?.attachments || [];
  const attachments = normalizeWorkAttachments(rawAttachments);
  const employees = Array.isArray(workState?.employees) && workState.employees.length
    ? workState.employees
    : defaultWorkState.employees;
  const requestedEmployeeId = String(input.employeeId || existing?.employeeId || defaultWorkEmployee.id).trim();
  const employeeId = employees.some((employee) => employee.id === requestedEmployeeId)
    ? requestedEmployeeId
    : defaultWorkEmployee.id;
  const errors = [];

  if (!title) errors.push("Enter the work title.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) errors.push("Choose a valid due date.");
  if (!employees.some((employee) => employee.id === employeeId)) errors.push("Choose a valid employee.");
  if (Array.isArray(rawAttachments) && rawAttachments.length > maxWorkAttachmentCount) {
    errors.push(`Attach up to ${maxWorkAttachmentCount} photos.`);
  }
  if (Array.isArray(rawAttachments) && attachments.length !== rawAttachments.length) {
    errors.push("One of the work photos is too large or is not a supported image.");
  }

  return {
    errors,
    assignment: {
      id: existing?.id || crypto.randomUUID(),
      employeeId,
      title,
      dueDate,
      priority,
      notes,
      attachments,
      status: existing?.status || "open",
      source: existing?.source || String(input.source || ""),
      sourceId: existing?.sourceId || String(input.sourceId || ""),
      inviteStatus: existing?.inviteStatus || "",
      inviteSentAt: existing?.inviteSentAt || "",
      inviteTo: existing?.inviteTo || "",
      inviteEmailFrom: existing?.inviteEmailFrom || "",
      inviteError: existing?.inviteError || "",
      larkNotifyStatus: existing?.larkNotifyStatus || "",
      larkNotifySentAt: existing?.larkNotifySentAt || "",
      larkNotifyTo: existing?.larkNotifyTo || "",
      larkNotifyError: existing?.larkNotifyError || "",
      completionNotifyStatus: existing?.completionNotifyStatus || "",
      completionNotifySentAt: existing?.completionNotifySentAt || "",
      completionNotifyTo: existing?.completionNotifyTo || "",
      completionNotifyError: existing?.completionNotifyError || "",
      completionEmailStatus: existing?.completionEmailStatus || "",
      completionEmailSentAt: existing?.completionEmailSentAt || "",
      completionEmailTo: existing?.completionEmailTo || "",
      completionEmailFrom: existing?.completionEmailFrom || "",
      completionEmailError: existing?.completionEmailError || "",
      completedAt: existing?.completedAt || "",
      createdAt: existing?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    }
  };
}

function bookingServiceLabel(booking) {
  if (Array.isArray(booking.services) && booking.services.length) {
    return booking.services
      .map((service) => service?.name || service)
      .filter(Boolean)
      .join(" + ");
  }

  return booking.service || "Booking";
}

function bookingWindowLabel(booking) {
  const start = new Date(booking.startAt);
  const end = new Date(booking.endAt || booking.startAt);
  if (Number.isNaN(start.getTime())) return "Time not set";

  const date = new Intl.DateTimeFormat("en-AU", {
    weekday: "short",
    day: "numeric",
    month: "short",
    timeZone: bookingTimeZone
  }).format(start);
  const time = new Intl.DateTimeFormat("en-AU", {
    hour: "numeric",
    minute: "2-digit",
    timeZone: bookingTimeZone
  });

  return Number.isNaN(end.getTime())
    ? `${date}, ${time.format(start)}`
    : `${date}, ${time.format(start)}-${time.format(end)}`;
}

function workdayBeforeBooking(booking) {
  const start = new Date(booking.startAt);
  if (Number.isNaN(start.getTime())) return toDateValue(new Date());

  const due = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  due.setDate(due.getDate() - 1);

  while (due.getDay() === 0 || due.getDay() === 6) {
    due.setDate(due.getDate() - 1);
  }

  const today = parseDateValue(toDateValue(new Date()));
  const bookingDay = new Date(start.getFullYear(), start.getMonth(), start.getDate());
  return toDateValue(bookingDay < today ? due : (due < today ? today : due));
}

function bookingWorkPriority(booking) {
  const due = parseDateValue(workdayBeforeBooking(booking));
  const today = parseDateValue(toDateValue(new Date()));
  const daysUntilDue = Math.ceil((due - today) / (24 * 60 * 60 * 1000));
  return daysUntilDue <= 1 ? "high" : "normal";
}

function isSyncableBooking(booking) {
  if (!booking?.id || booking.status === "cancelled") return false;
  const end = new Date(booking.endAt || booking.startAt);
  return !Number.isNaN(end.getTime());
}

function assignmentFromBooking(booking, existing = null) {
  const service = bookingServiceLabel(booking);
  const title = `Booking prep: ${booking.propertyAddress || service}`;
  const notes = [
    "Imported from the internal booking system.",
    `Booking: ${bookingWindowLabel(booking)}`,
    `Services: ${service}`,
    booking.propertyAddress ? `Address: ${booking.propertyAddress}` : "",
    booking.clientName ? `Client: ${booking.clientName}` : "",
    booking.agentName ? `Agent: ${booking.agentName}${booking.agentPhone ? `, ${booking.agentPhone}` : ""}` : "",
    booking.photographerName ? `Photographer: ${booking.photographerName}${booking.photographerPhone ? `, ${booking.photographerPhone}` : ""}` : "",
    booking.notes ? `Booking notes: ${booking.notes}` : ""
  ].filter(Boolean).join("\n");

  return {
    id: existing?.id || crypto.randomUUID(),
    employeeId: existing?.employeeId || defaultWorkEmployee.id,
    title,
    dueDate: workdayBeforeBooking(booking),
    priority: bookingWorkPriority(booking),
    notes,
    status: existing?.status || "open",
    source: "booking",
    sourceId: booking.id,
    attachments: normalizeWorkAttachments(existing?.attachments || []),
    inviteStatus: existing?.inviteStatus || "",
    inviteSentAt: existing?.inviteSentAt || "",
    inviteTo: existing?.inviteTo || "",
    inviteEmailFrom: existing?.inviteEmailFrom || "",
    inviteError: existing?.inviteError || "",
    larkNotifyStatus: existing?.larkNotifyStatus || "",
    larkNotifySentAt: existing?.larkNotifySentAt || "",
    larkNotifyTo: existing?.larkNotifyTo || "",
    larkNotifyError: existing?.larkNotifyError || "",
    completionNotifyStatus: existing?.completionNotifyStatus || "",
    completionNotifySentAt: existing?.completionNotifySentAt || "",
    completionNotifyTo: existing?.completionNotifyTo || "",
    completionNotifyError: existing?.completionNotifyError || "",
    completionEmailStatus: existing?.completionEmailStatus || "",
    completionEmailSentAt: existing?.completionEmailSentAt || "",
    completionEmailTo: existing?.completionEmailTo || "",
    completionEmailFrom: existing?.completionEmailFrom || "",
    completionEmailError: existing?.completionEmailError || "",
    completedAt: existing?.completedAt || "",
    createdAt: existing?.createdAt || new Date().toISOString(),
    updatedAt: new Date().toISOString()
  };
}

function mergeBookingWorkAssignments(workState, bookings) {
  let created = 0;
  let updated = 0;
  const createdAssignments = [];

  for (const booking of bookings.filter(isSyncableBooking)) {
    const existing = workState.assignments.find((assignment) => {
      return assignment.source === "booking" && assignment.sourceId === booking.id;
    });
    const next = assignmentFromBooking(booking, existing);

    if (existing) {
      workState.assignments = workState.assignments.map((assignment) =>
        assignment.id === existing.id ? next : assignment
      );
      updated += 1;
    } else {
      workState.assignments.push(next);
      createdAssignments.push(next);
      created += 1;
    }
  }

  return { created, updated, syncable: bookings.filter(isSyncableBooking).length, createdAssignments };
}

function normalizeInvoiceStatus(status) {
  return ["draft", "paid", "void"].includes(status) ? status : "draft";
}

function normalizeEditableInvoiceNumber(value) {
  return String(value || "").trim().toUpperCase().replace(/\s+/g, "");
}

function validateEditableInvoiceNumber(value) {
  const invoiceNumber = normalizeEditableInvoiceNumber(value);
  const errors = [];
  if (!invoiceNumber) {
    errors.push("Enter an invoice number.");
  } else if (invoiceNumber.length > 32) {
    errors.push("Invoice number is too long.");
  } else if (!/^[A-Z0-9][A-Z0-9-]*$/.test(invoiceNumber)) {
    errors.push("Use only letters, numbers, and hyphens.");
  }

  return { invoiceNumber, errors };
}

function invoiceNumberInUse(invoices, invoiceNumber, exceptId = "") {
  const key = normalizeEditableInvoiceNumber(invoiceNumber);
  if (!key) return false;

  return invoices.some((invoice) => (
    invoice.id !== exceptId
    && normalizeEditableInvoiceNumber(invoice.invoiceNumber) === key
  ));
}

function normalizeMoney(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 100) / 100 : 0;
}

function normalizeQuantity(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 100) / 100 : 0;
}

function calculateGst(subtotal) {
  return normalizeMoney(subtotal * invoiceGstRate);
}

function calculateInvoiceTotals(items) {
  const subtotal = normalizeMoney(items.reduce((sum, item) => sum + normalizeMoney(item.amount), 0));
  const gstAmount = calculateGst(subtotal);
  return {
    subtotal,
    gstAmount,
    total: normalizeMoney(subtotal + gstAmount)
  };
}

function normalizeInvoice(invoice) {
  const items = Array.isArray(invoice?.items)
    ? invoice.items.map((item) => {
        const quantity = normalizeQuantity(item.quantity ?? 1);
        const safeQuantity = quantity > 0 ? quantity : 1;
        const unitPrice = normalizeMoney(item.unitPrice ?? item.price ?? (item.amount ? Number(item.amount) / safeQuantity : 0));
        return {
          name: String(item.name || "Service").trim() || "Service",
          quantity: safeQuantity,
          unitPrice,
          amount: normalizeMoney(item.amount ?? safeQuantity * unitPrice)
        };
      })
    : [];
  const subtotal = normalizeMoney(invoice?.subtotal ?? items.reduce((sum, item) => sum + item.amount, 0));
  const gstRate = Number.isFinite(Number(invoice?.gstRate)) ? Number(invoice.gstRate) : invoiceGstRate;
  const gstAmount = normalizeMoney(invoice?.gstAmount ?? subtotal * gstRate);
  const hasStoredGst = invoice?.gstAmount !== undefined || invoice?.gstRate !== undefined;
  const total = normalizeMoney(hasStoredGst && invoice?.total !== undefined ? invoice.total : subtotal + gstAmount);

  return {
    id: invoice?.id || crypto.randomUUID(),
    invoiceNumber: String(invoice?.invoiceNumber || "").trim(),
    bookingId: String(invoice?.bookingId || ""),
    sourceInvoiceId: String(invoice?.sourceInvoiceId || ""),
    source: String(invoice?.source || ""),
    propertyAddress: String(invoice?.propertyAddress || ""),
    clientName: String(invoice?.clientName || ""),
    clientEmail: String(invoice?.clientEmail || ""),
    agentName: String(invoice?.agentName || ""),
    agentPhone: String(invoice?.agentPhone || ""),
    photographerName: String(invoice?.photographerName || ""),
    bookingStartAt: invoice?.bookingStartAt || "",
    bookingEndAt: invoice?.bookingEndAt || "",
    items,
    subtotal,
    discount: normalizeMoney(invoice?.discount || 0),
    gstRate,
    gstAmount,
    gstIncluded: normalizeEnvBoolean(invoice?.gstIncluded, true),
    total,
    currency: invoice?.currency || invoiceCurrency,
    status: normalizeInvoiceStatus(invoice?.status),
    issuedAt: invoice?.issuedAt || new Date().toISOString(),
    dueAt: invoice?.dueAt || addDays(new Date(), 7).toISOString(),
    paidAt: invoice?.paidAt || "",
    voidedAt: invoice?.voidedAt || "",
    sentAt: invoice?.sentAt || "",
    sentTo: uniqueEmails(Array.isArray(invoice?.sentTo) ? invoice.sentTo : parseGuestEmails(invoice?.sentTo || "")),
    notes: String(invoice?.notes || ""),
    detailsEditedAt: invoice?.detailsEditedAt || "",
    importedAt: invoice?.importedAt || "",
    pricesEditedAt: invoice?.pricesEditedAt || "",
    createdAt: invoice?.createdAt || new Date().toISOString(),
    updatedAt: invoice?.updatedAt || invoice?.createdAt || new Date().toISOString()
  };
}

function normalizeInvoices(invoices) {
  return Array.isArray(invoices)
    ? invoices.map(normalizeInvoice).filter((invoice) => invoice.bookingId || invoice.invoiceNumber)
    : [];
}

function addDays(date, days) {
  const next = new Date(date);
  next.setDate(next.getDate() + days);
  return next;
}

function bookingServiceNames(booking) {
  const services = Array.isArray(booking?.services) && booking.services.length
    ? booking.services.map((service) => service?.name || service).filter(Boolean)
    : String(booking?.service || "").split("+").map((service) => service.trim()).filter(Boolean);
  return [...new Set(services)];
}

function bookingInvoiceItems(booking) {
  return bookingServiceNames(booking).map((name) => {
    const unitPrice = normalizeMoney(invoiceServicePrices[name] ?? 0);
    return {
      name,
      quantity: 1,
      unitPrice,
      amount: unitPrice
    };
  });
}

function invoiceItemPriceKey(item) {
  return String(item?.name || "").trim().toLowerCase();
}

function mergeBookingInvoiceItemsWithExistingPrices(nextItems, existingItems) {
  const existingByName = new Map(
    (Array.isArray(existingItems) ? existingItems : [])
      .filter((item) => invoiceItemPriceKey(item))
      .map((item) => [invoiceItemPriceKey(item), item])
  );

  return nextItems.map((item) => {
    const existing = existingByName.get(invoiceItemPriceKey(item));
    if (!existing) return item;

    const quantity = Math.max(1, Number(item.quantity || existing.quantity || 1));
    const unitPrice = normalizeMoney(existing.unitPrice ?? existing.amount);
    return {
      ...item,
      quantity,
      unitPrice,
      amount: normalizeMoney(quantity * unitPrice)
    };
  });
}

function normalizeManualInvoiceItem(item) {
  const name = String(item?.name || "").trim();
  const quantity = normalizeQuantity(item?.quantity ?? 1);
  const unitPrice = normalizeMoney(item?.unitPrice ?? item?.price ?? item?.amount);
  return {
    name,
    quantity,
    unitPrice,
    amount: normalizeMoney(quantity * unitPrice)
  };
}

function validateEditableInvoiceItems(input) {
  const errors = [];
  const items = (Array.isArray(input?.items) ? input.items : [])
    .map(normalizeManualInvoiceItem)
    .filter((item) => item.name || item.amount > 0);

  if (!items.length) {
    errors.push("Add at least one invoice item.");
  }

  for (const item of items) {
    if (!item.name) errors.push("Enter a name for each invoice item.");
    if (item.name.length > 80) errors.push(`${item.name} is too long.`);
    if (!(item.quantity > 0)) errors.push(`${item.name || "Invoice item"} needs a quantity above 0.`);
    if (item.unitPrice < 0 || item.amount < 0) errors.push(`${item.name || "Invoice item"} cannot be below $0.`);
    if (item.amount > 100000) errors.push(`${item.name || "Invoice item"} is too high.`);
  }

  return {
    errors,
    items,
    ...calculateInvoiceTotals(items)
  };
}

function validateManualInvoice(input, invoices) {
  const errors = [];
  const clientName = String(input?.clientName || "").trim();
  const clientEmail = uniqueEmails(parseGuestEmails(input?.clientEmail || ""));
  const propertyAddress = String(input?.propertyAddress || input?.description || "").trim();
  const agentName = String(input?.agentName || "").trim();
  const agentPhone = String(input?.agentPhone || "").trim();
  const items = (Array.isArray(input?.items) ? input.items : [])
    .map(normalizeManualInvoiceItem)
    .filter((item) => item.name || item.amount > 0);
  const invoiceNumberInput = String(input?.invoiceNumber || "").trim();
  const invoiceNumberValidation = invoiceNumberInput
    ? validateEditableInvoiceNumber(invoiceNumberInput)
    : { invoiceNumber: "", errors: [] };

  if (!clientName) errors.push("Enter the client name.");
  if (!propertyAddress) errors.push("Enter the property or invoice description.");
  for (const email of clientEmail) {
    if (!isValidEmail(email)) errors.push(`Enter a valid client email for ${email}.`);
  }
  for (const item of items) {
    if (!item.name) errors.push("Enter a name for each invoice item.");
    if (!(item.quantity > 0)) errors.push(`${item.name || "Invoice item"} needs a quantity above 0.`);
    if (!(item.amount > 0)) errors.push(`${item.name || "Invoice item"} needs an amount above $0.`);
  }
  if (!items.length) errors.push("Add at least one invoice item.");
  errors.push(...invoiceNumberValidation.errors);

  let invoiceNumber = invoiceNumberValidation.invoiceNumber;
  if (!invoiceNumber && !errors.length) {
    invoiceNumber = nextInvoiceNumber(invoices, { clientName });
  }
  if (invoiceNumber && invoiceNumberInUse(invoices, invoiceNumber)) {
    errors.push(`${invoiceNumber} is already used by another invoice.`);
  }

  const issuedDate = new Date(input?.issuedAt || Date.now());
  const issuedAt = Number.isNaN(issuedDate.getTime()) ? new Date() : issuedDate;
  const dueDate = input?.dueAt ? new Date(input.dueAt) : addDays(issuedAt, 7);
  const dueAt = Number.isNaN(dueDate.getTime()) ? addDays(issuedAt, 7) : dueDate;
  const { subtotal, gstAmount, total } = calculateInvoiceTotals(items);
  const status = normalizeInvoiceStatus(input?.status);
  const now = new Date().toISOString();

  return {
    errors,
    invoice: normalizeInvoice({
      id: crypto.randomUUID(),
      invoiceNumber,
      source: "manual",
      propertyAddress,
      clientName,
      clientEmail: clientEmail.join(", "),
      agentName,
      agentPhone,
      items,
      subtotal,
      gstRate: invoiceGstRate,
      gstAmount,
      gstIncluded: true,
      total,
      currency: invoiceCurrency,
      status,
      issuedAt: issuedAt.toISOString(),
      dueAt: dueAt.toISOString(),
      paidAt: status === "paid" ? now : "",
      notes: String(input?.notes || "Created manually.").trim(),
      createdAt: now,
      updatedAt: now
    })
  };
}

function validateEditableInvoiceDetails(input, existing) {
  const errors = [];
  const invoiceNumberValidation = validateEditableInvoiceNumber(input?.invoiceNumber ?? existing?.invoiceNumber);
  const clientName = String(input?.clientName ?? existing?.clientName ?? "").trim();
  const clientEmail = uniqueEmails(parseGuestEmails(input?.clientEmail ?? existing?.clientEmail ?? ""));
  const propertyAddress = String(input?.propertyAddress ?? input?.description ?? existing?.propertyAddress ?? "").trim();
  const agentName = String(input?.agentName ?? existing?.agentName ?? "").trim();
  const agentPhone = String(input?.agentPhone ?? existing?.agentPhone ?? "").trim();
  const issuedDate = new Date(input?.issuedAt ?? existing?.issuedAt ?? Date.now());
  const issuedAt = Number.isNaN(issuedDate.getTime()) ? null : issuedDate;
  const dueDate = input?.dueAt ? new Date(input.dueAt) : (issuedAt ? addDays(issuedAt, 7) : null);
  const dueAt = dueDate && !Number.isNaN(dueDate.getTime()) ? dueDate : null;

  if (!clientName) errors.push("Enter the client name.");
  if (clientName.length > 120) errors.push("Client name is too long.");
  if (!propertyAddress) errors.push("Enter the property or invoice description.");
  if (propertyAddress.length > 180) errors.push("Property or invoice description is too long.");
  if (agentName.length > 120) errors.push("Agent name is too long.");
  if (agentPhone.length > 60) errors.push("Agent phone is too long.");
  if (!issuedAt) errors.push("Enter a valid issue date.");
  if (!dueAt) errors.push("Enter a valid due date.");
  for (const email of clientEmail) {
    if (!isValidEmail(email)) errors.push(`Enter a valid client email for ${email}.`);
  }
  errors.push(...invoiceNumberValidation.errors);

  return {
    errors,
    details: {
      invoiceNumber: invoiceNumberValidation.invoiceNumber,
      propertyAddress,
      clientName,
      clientEmail: clientEmail.join(", "),
      agentName,
      agentPhone,
      issuedAt: issuedAt ? issuedAt.toISOString() : "",
      dueAt: dueAt ? dueAt.toISOString() : ""
    }
  };
}

function normalizeInvoiceClientKey(value) {
  return String(value || "")
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/&/g, " and ")
    .replace(/\b(pty|ltd|limited|proprietary|company|co|the)\b/g, " ")
    .replace(/[^a-z0-9]+/g, " ")
    .trim()
    .replace(/\s+/g, " ");
}

function compactInvoiceClientKey(value) {
  return normalizeInvoiceClientKey(value).replace(/\s+/g, "");
}

function invoiceClientNames(source = {}) {
  const names = [
    source.clientName,
    source.name,
    source.billingName
  ]
    .map((value) => String(value || "").trim())
    .filter(Boolean);

  for (const name of [...names]) {
    const parentheticalMatches = [...name.matchAll(/\(([^)]+)\)/g)];
    parentheticalMatches.forEach((match) => {
      if (match[1]) names.push(match[1].trim());
    });
  }

  return [...new Set(names)];
}

function invoiceClientNameMatches(left, right) {
  const leftKey = normalizeInvoiceClientKey(left);
  const rightKey = normalizeInvoiceClientKey(right);
  if (!leftKey || !rightKey) return false;
  if (leftKey === rightKey) return true;
  if (leftKey.includes(rightKey) || rightKey.includes(leftKey)) return true;

  const leftCompact = compactInvoiceClientKey(left);
  const rightCompact = compactInvoiceClientKey(right);
  return Boolean(leftCompact && rightCompact && leftCompact === rightCompact);
}

function invoiceNumberParts(value) {
  const invoiceNumber = String(value || "").trim().toUpperCase();
  const prefixMatch = invoiceNumber.match(/^([A-Z0-9]*?[A-Z])(\d+)$/);
  if (prefixMatch) {
    return {
      prefix: prefixMatch[1],
      number: Number(prefixMatch[2]),
      width: prefixMatch[2].length
    };
  }

  const legacyMatch = invoiceNumber.match(/^INV-\d{4}-(\d{4,})$/);
  if (legacyMatch) {
    return {
      prefix: "R",
      number: Number(legacyMatch[1]),
      width: Math.max(3, legacyMatch[1].length)
    };
  }

  return null;
}

function matchingInvoicePrefixOverride(source) {
  const names = invoiceClientNames(source);
  return invoicePrefixOverrides.find((override) => {
    return override.names.some((overrideName) =>
      names.some((name) => invoiceClientNameMatches(name, overrideName))
    );
  })?.prefix || "";
}

function prefixFromInvoiceHistory(source, invoices) {
  const names = invoiceClientNames(source);
  if (!names.length) return "";

  const matches = new Map();
  for (const invoice of invoices) {
    const parts = invoiceNumberParts(invoice.invoiceNumber);
    if (!parts?.prefix) continue;

    const invoiceNames = invoiceClientNames(invoice);
    if (!invoiceNames.some((invoiceName) => names.some((name) => invoiceClientNameMatches(invoiceName, name)))) {
      continue;
    }

    const existing = matches.get(parts.prefix) || { count: 0, highest: 0 };
    matches.set(parts.prefix, {
      count: existing.count + 1,
      highest: Math.max(existing.highest, parts.number)
    });
  }

  return [...matches.entries()]
    .sort((left, right) => {
      if (right[1].count !== left[1].count) return right[1].count - left[1].count;
      return right[1].highest - left[1].highest;
    })[0]?.[0] || "";
}

function initialsInvoicePrefix(source) {
  const sourceName = invoiceClientNames(source)[0] || "Receipt";
  const normalized = normalizeInvoiceClientKey(sourceName)
    .replace(/\b(real estate|property|properties|agency|group|studio|australia|australian|service|services)\b/g, " ")
    .trim();
  const words = normalized.split(/\s+/).filter(Boolean);
  if (!words.length) return "R";

  const initials = words
    .map((word) => word[0])
    .join("")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
  if (initials.length >= 2) return initials.slice(0, 4);

  return words[0].slice(0, 3).toUpperCase() || "R";
}

function invoicePrefixForSource(source, invoices) {
  return matchingInvoicePrefixOverride(source)
    || prefixFromInvoiceHistory(source, invoices)
    || initialsInvoicePrefix(source);
}

function nextInvoiceNumber(invoices, source = {}) {
  const prefix = invoicePrefixForSource(source, invoices);
  const matchingNumbers = invoices
    .map((invoice) => invoiceNumberParts(invoice.invoiceNumber))
    .filter((parts) => parts?.prefix === prefix);

  const maxNumber = matchingNumbers.reduce((max, parts) => Math.max(max, parts.number), 0);
  const width = Math.max(3, ...matchingNumbers.map((parts) => parts.width));
  return `${prefix}${String(maxNumber + 1).padStart(width, "0")}`;
}

function invoiceNumberForBooking(existing, booking, invoices) {
  const otherInvoices = existing?.id
    ? invoices.filter((invoice) => invoice.id !== existing.id)
    : invoices;
  const nextNumber = () => nextInvoiceNumber(otherInvoices, booking);
  if (!existing?.invoiceNumber) {
    return nextNumber();
  }

  if (invoiceNumberInUse(otherInvoices, existing.invoiceNumber)) {
    return nextNumber();
  }

  const canRefreshDraftNumber =
    existing.status === "draft"
    && !existing.sentAt
    && !existing.sourceInvoiceId;
  if (!canRefreshDraftNumber) {
    return existing.invoiceNumber;
  }

  const currentPrefix = invoiceNumberParts(existing.invoiceNumber)?.prefix || "";
  const expectedPrefix = invoicePrefixForSource(booking, otherInvoices);
  return currentPrefix === expectedPrefix ? existing.invoiceNumber : nextNumber();
}

function importInvoiceId(sourceInvoiceId) {
  const safeId = String(sourceInvoiceId || "")
    .trim()
    .replace(/[^a-zA-Z0-9_-]+/g, "-")
    .replace(/^-+|-+$/g, "");
  return safeId ? `legacy-invoice-${safeId}` : crypto.randomUUID();
}

function uniqueImportInvoiceId(invoices, sourceInvoiceId) {
  const baseId = importInvoiceId(sourceInvoiceId);
  if (!invoices.some((invoice) => invoice.id === baseId)) {
    return baseId;
  }

  let suffix = 2;
  while (invoices.some((invoice) => invoice.id === `${baseId}-${suffix}`)) {
    suffix += 1;
  }
  return `${baseId}-${suffix}`;
}

function mergeImportedInvoice(existing, input, allInvoices, now = new Date().toISOString()) {
  const sourceInvoiceId = String(input?.sourceInvoiceId || "").trim();
  const status = normalizeInvoiceStatus(input?.status);
  const issuedAt = input?.issuedAt || existing?.issuedAt || now;
  const invoice = normalizeInvoice({
    ...(existing || {}),
    ...input,
    id: existing?.id || input?.id || uniqueImportInvoiceId(allInvoices, sourceInvoiceId),
    sourceInvoiceId,
    source: input?.source || existing?.source || "legacy-import",
    status,
    issuedAt,
    dueAt: input?.dueAt || existing?.dueAt || addDays(new Date(issuedAt), 7).toISOString(),
    paidAt: status === "paid" ? (input?.paidAt || existing?.paidAt || issuedAt) : "",
    voidedAt: status === "void" ? (input?.voidedAt || existing?.voidedAt || now) : "",
    importedAt: existing?.importedAt || input?.importedAt || now,
    createdAt: existing?.createdAt || input?.createdAt || issuedAt,
    updatedAt: now
  });

  return invoice;
}

function importInvoices(invoices, importedRecords) {
  const now = new Date().toISOString();
  const sourceIndex = new Map();
  invoices.forEach((invoice, index) => {
    if (invoice.sourceInvoiceId) {
      sourceIndex.set(invoice.sourceInvoiceId, index);
    }
  });

  let imported = 0;
  let updated = 0;
  const skipped = [];
  const duplicateInvoiceNumbers = [];

  for (const record of importedRecords) {
    const sourceInvoiceId = String(record?.sourceInvoiceId || "").trim();
    const invoiceNumberValidation = validateEditableInvoiceNumber(record?.invoiceNumber);
    const invoiceNumber = invoiceNumberValidation.invoiceNumber;

    if (!sourceInvoiceId || !invoiceNumber || invoiceNumberValidation.errors.length) {
      skipped.push({
        sourceInvoiceId,
        invoiceNumber: String(record?.invoiceNumber || "").trim(),
        reason: invoiceNumberValidation.errors.join(" ") || "Missing original invoice ID or invoice number."
      });
      continue;
    }

    const existingIndex = sourceIndex.get(sourceInvoiceId);
    const existing = existingIndex === undefined ? null : invoices[existingIndex];
    if (invoiceNumberInUse(invoices, invoiceNumber, existing?.id || "")) {
      duplicateInvoiceNumbers.push({ sourceInvoiceId, invoiceNumber });
      skipped.push({
        sourceInvoiceId,
        invoiceNumber,
        reason: `${invoiceNumber} is already used by another invoice.`
      });
      continue;
    }

    const invoice = mergeImportedInvoice(existing, { ...record, sourceInvoiceId, invoiceNumber }, invoices, now);
    if (existingIndex === undefined) {
      invoices.push(invoice);
      sourceIndex.set(sourceInvoiceId, invoices.length - 1);
      imported += 1;
    } else {
      invoices[existingIndex] = invoice;
      updated += 1;
    }
  }

  invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
  return { imported, updated, skipped, duplicateInvoiceNumbers };
}

function invoiceFromBooking(booking, invoices, existing = null) {
  const now = new Date().toISOString();
  const isPaid = existing?.status === "paid";
  const existingItems = Array.isArray(existing?.items) ? existing.items : [];
  const bookingItems = bookingInvoiceItems(booking);
  const items = isPaid ? existingItems : mergeBookingInvoiceItemsWithExistingPrices(bookingItems, existingItems);
  const { subtotal, gstAmount, total } = calculateInvoiceTotals(items);
  const invoiceDetails = existing?.detailsEditedAt ? existing : booking;
  const status = booking.status === "cancelled"
    ? (isPaid ? "paid" : "void")
    : (isPaid ? "paid" : "draft");

  return normalizeInvoice({
    ...(existing || {}),
    id: existing?.id || crypto.randomUUID(),
    invoiceNumber: invoiceNumberForBooking(existing, booking, invoices),
    bookingId: booking.id,
    propertyAddress: invoiceDetails.propertyAddress || "",
    clientName: invoiceDetails.clientName || "",
    clientEmail: invoiceDetails.clientEmail || "",
    agentName: invoiceDetails.agentName || "",
    agentPhone: invoiceDetails.agentPhone || "",
    photographerName: booking.photographerName || "",
    bookingStartAt: booking.startAt || "",
    bookingEndAt: booking.endAt || "",
    items,
    subtotal,
    gstRate: invoiceGstRate,
    gstAmount,
    total,
    currency: invoiceCurrency,
    status,
    issuedAt: existing?.issuedAt || now,
    dueAt: existing?.dueAt || addDays(new Date(), 7).toISOString(),
    paidAt: existing?.paidAt || "",
    voidedAt: status === "void" ? (existing?.voidedAt || now) : "",
    notes: "Generated from internal booking.",
    detailsEditedAt: existing?.detailsEditedAt || "",
    pricesEditedAt: existing?.pricesEditedAt || "",
    createdAt: existing?.createdAt || now,
    updatedAt: now
  });
}

function upsertInvoiceForBooking(invoices, booking) {
  if (!booking?.id || booking.larkOnly) {
    return { invoice: null, created: false, updated: false };
  }

  const index = invoices.findIndex((invoice) => invoice.bookingId === booking.id);
  const existing = index >= 0 ? invoices[index] : null;
  if (booking.status === "cancelled" && !existing) {
    return { invoice: null, created: false, updated: false };
  }

  const invoice = invoiceFromBooking(booking, invoices, existing);
  if (index >= 0) {
    invoices[index] = invoice;
    return { invoice, created: false, updated: true };
  }

  invoices.push(invoice);
  return { invoice, created: true, updated: false };
}

function syncInvoicesFromBookings(invoices, bookings) {
  let created = 0;
  let updated = 0;

  for (const booking of bookings.filter((item) => !item.larkOnly)) {
    const result = upsertInvoiceForBooking(invoices, booking);
    if (result.created) created += 1;
    if (result.updated) updated += 1;
  }

  return { created, updated, total: invoices.length };
}

function completedBookingAssignments(workState) {
  const assignments = Array.isArray(workState?.assignments) ? workState.assignments : [];
  return new Map(
    assignments
      .filter((assignment) => (
        assignment.status === "done"
        && assignment.source === "booking"
        && assignment.sourceId
      ))
      .map((assignment) => [assignment.sourceId, assignment])
  );
}

function applyCompletedJobDatesToInvoices(invoices, workState) {
  const assignmentsByBookingId = completedBookingAssignments(workState);
  let updated = 0;

  for (const invoice of invoices) {
    const assignment = assignmentsByBookingId.get(invoice.bookingId);
    if (!assignment) continue;

    const completedDate = new Date(assignment.completedAt || assignment.updatedAt || "");
    if (Number.isNaN(completedDate.getTime())) continue;

    const issuedAt = completedDate.toISOString();
    const dueAt = addDays(completedDate, 7).toISOString();
    if (invoice.issuedAt === issuedAt && invoice.dueAt === dueAt) continue;

    invoice.issuedAt = issuedAt;
    invoice.dueAt = dueAt;
    invoice.updatedAt = new Date().toISOString();
    updated += 1;
  }

  return { updated };
}

async function syncInvoiceDatesFromCompletedJobs(invoices, workState = null, options = {}) {
  try {
    const { persist = true } = options;
    const resolvedWorkState = workState || await loadWorkState();
    const result = applyCompletedJobDatesToInvoices(invoices, resolvedWorkState);
    if (persist && result.updated) {
      await saveInvoices(invoices);
    }
    return result;
  } catch {
    return { updated: 0 };
  }
}

function bookingHasPassed(booking, now = Date.now()) {
  const endSource = booking?.endAt || (
    booking?.startAt && booking?.durationMinutes
      ? getEndAt(booking.startAt, Number(booking.durationMinutes))
      : ""
  );
  const endDate = new Date(endSource);
  return !Number.isNaN(endDate.getTime()) && endDate.getTime() < now;
}

function markPastBookingLocal(booking) {
  booking.larkStatus = "past_local";
  booking.larkError = null;
  booking.larkAttendeeStatus = "not_needed";
  booking.larkAttendeeError = null;
  booking.calendarInviteStatus = "not_needed";
  booking.calendarInviteError = null;
  return booking;
}

function bookingIdFromApiPath(pathname, trailingAction = "") {
  const prefix = "/api/bookings/";
  if (!pathname.startsWith(prefix)) return "";

  let rawId = pathname.slice(prefix.length);
  if (trailingAction) {
    const suffix = `/${trailingAction}`;
    if (!rawId.endsWith(suffix)) return "";
    rawId = rawId.slice(0, -suffix.length);
  }

  if (!rawId) return "";
  try {
    return decodeURIComponent(rawId);
  } catch {
    return rawId;
  }
}

function invoiceIsForPastBooking(invoice, pastBookingIds, now = Date.now()) {
  if (invoice?.bookingId && pastBookingIds.has(invoice.bookingId)) {
    return true;
  }

  const endDate = new Date(invoice?.bookingEndAt || "");
  return !Number.isNaN(endDate.getTime()) && endDate.getTime() < now;
}

async function deletePastBookingsAndInvoices() {
  const now = Date.now();
  const bookings = await loadBookings();
  const pastBookingIds = new Set(
    bookings
      .filter((booking) => bookingHasPassed(booking, now))
      .map((booking) => booking.id)
      .filter(Boolean)
  );
  const nextBookings = bookings.filter((booking) => !pastBookingIds.has(booking.id));

  const invoices = await loadInvoices();
  const nextInvoices = invoices.filter((invoice) => !invoiceIsForPastBooking(invoice, pastBookingIds, now));

  if (nextBookings.length !== bookings.length) {
    await saveBookings(nextBookings);
  }

  if (nextInvoices.length !== invoices.length) {
    await saveInvoices(nextInvoices);
  }

  return {
    cutoff: new Date(now).toISOString(),
    removedBookingIds: [...pastBookingIds],
    removedBookings: bookings.length - nextBookings.length,
    removedInvoices: invoices.length - nextInvoices.length,
    remainingBookings: nextBookings.length,
    remainingInvoices: nextInvoices.length
  };
}

function invoiceEmailMissingSettings() {
  const missing = [];
  if (invoiceEmailProvider === "resend") {
    if (!resendConfig.apiKey) missing.push("RESEND_API_KEY");
    return missing;
  }

  if (!invoiceEmailConfig.host) missing.push("INVOICE_EMAIL_HOST");
  if (!invoiceEmailConfig.user) missing.push("INVOICE_EMAIL_USER");
  if (!invoiceEmailConfig.pass) missing.push("INVOICE_EMAIL_PASSWORD");
  return missing;
}

function isInvoiceEmailConfigured() {
  return invoiceEmailMissingSettings().length === 0;
}

function formatInvoiceMoney(value) {
  return new Intl.NumberFormat("en-AU", {
    style: "currency",
    currency: invoiceCurrency
  }).format(Number(value || 0));
}

function formatDocumentMoney(value, currency = invoiceCurrency) {
  const safeCurrency = /^[A-Z]{3}$/.test(String(currency || "")) ? String(currency).toUpperCase() : invoiceCurrency;
  const amount = Number(value || 0);
  if (safeCurrency === "THB") {
    return `${safeCurrency} ${amount.toLocaleString("en-AU", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2
    })}`;
  }

  try {
    return new Intl.NumberFormat("en-AU", {
      style: "currency",
      currency: safeCurrency,
      currencyDisplay: "symbol"
    }).format(amount).replace(/\u00a0/g, " ");
  } catch {
    return `${safeCurrency} ${amount.toFixed(2)}`;
  }
}

function formatInvoiceDocumentDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return new Date().toISOString().slice(0, 10);
  }

  return date.toISOString().slice(0, 10);
}

function formatInvoiceQuantity(value) {
  const quantity = normalizeQuantity(value ?? 1);
  if (Number.isInteger(quantity)) return String(quantity);
  return quantity.toLocaleString("en-AU", {
    minimumFractionDigits: 0,
    maximumFractionDigits: 2
  });
}

function invoicePdfItems(invoice) {
  const items = Array.isArray(invoice?.items) && invoice.items.length
    ? invoice.items
    : [{ name: "Service", quantity: 1, unitPrice: invoice.subtotal || invoice.total || 0, amount: invoice.subtotal || invoice.total || 0 }];
  const description = String(invoice?.propertyAddress || "").trim();

  return items.map((item) => {
    const quantity = normalizeQuantity(item.quantity || 1) || 1;
    const unitPrice = normalizeMoney(item.unitPrice ?? item.price ?? (item.amount ? Number(item.amount) / quantity : 0));
    const amount = normalizeMoney(item.amount ?? quantity * unitPrice);
    const name = String(item.name || "Service").trim() || "Service";
    return {
      product: [description, name].filter(Boolean).join("\n"),
      quantity,
      unitPrice,
      amount
    };
  });
}

function invoiceFileName(invoice) {
  const base = [invoice.invoiceNumber || "invoice", invoice.propertyAddress || ""]
    .filter(Boolean)
    .join("-")
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 96);

  return `${base || "invoice"}.pdf`;
}

function invoiceEmailRecipients(invoice, input = {}) {
  const requestedRecipients = Array.isArray(input.to || input.recipients)
    ? (input.to || input.recipients)
    : parseGuestEmails(input.to || input.recipients || "");
  const fallbackRecipients = parseGuestEmails(invoice.clientEmail || "");
  return uniqueEmails(requestedRecipients.length ? requestedRecipients : fallbackRecipients).filter(isValidEmail);
}

function buildInvoiceEmailSubject(invoice) {
  return `${invoiceEmailConfig.subjectPrefix} ${invoice.invoiceNumber} - ${invoice.propertyAddress || "OpenFrame Studio"}`;
}

function buildInvoiceEmailText(invoice) {
  return [
    "Hi,",
    "",
    `Please find attached tax invoice ${invoice.invoiceNumber} for ${invoice.propertyAddress || "your booking"}.`,
    `Total due: ${formatInvoiceMoney(invoice.total)} including GST.`,
    "",
    "Payment information:",
    "Name: Openframe Studio Pty Ltd",
    "BSB: 062-128",
    "Account: 11440602",
    `Please reference ${invoice.invoiceNumber} for the payment.`,
    "",
    "Thank you,",
    "OpenFrame Studio"
  ].join("\n");
}

function buildInvoiceEmailHtml(invoice) {
  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p>Hi,</p>
      <p>Please find attached tax invoice <strong>${escapeHtmlForEmail(invoice.invoiceNumber)}</strong> for ${escapeHtmlForEmail(invoice.propertyAddress || "your booking")}.</p>
      <p><strong>Total due: ${escapeHtmlForEmail(formatInvoiceMoney(invoice.total))} including GST.</strong></p>
      <p>Payment information:</p>
      <ul>
        <li>Name: Openframe Studio Pty Ltd</li>
        <li>BSB: 062-128</li>
        <li>Account: 11440602</li>
        <li>Please reference ${escapeHtmlForEmail(invoice.invoiceNumber)} for the payment.</li>
      </ul>
      <p>Thank you,<br />OpenFrame Studio</p>
    </div>
  `;
}

function escapeHtmlForEmail(value) {
  return String(value ?? "").replace(/[&<>"']/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  }[char]));
}

function drawPdfCell(doc, text, x, y, width, height, options = {}) {
  const paddingX = options.paddingX ?? 8;
  const paddingY = options.paddingY ?? 7;
  const textHeight = doc.heightOfString(text, {
    width: width - paddingX * 2,
    align: options.align || "left"
  });
  const textY = y + Math.max(paddingY, (height - textHeight) / 2);
  doc.text(text, x + paddingX, textY, {
    width: width - paddingX * 2,
    height: height - paddingY * 2,
    align: options.align || "left",
    ellipsis: true
  });
}

function normalizeBillingMatchValue(value) {
  return String(value || "").trim().toLowerCase();
}

function invoiceBillingClientScore(invoice, client) {
  let score = 0;
  const invoiceClientName = normalizeBillingMatchValue(invoice.clientName);
  const clientName = normalizeBillingMatchValue(client.name);
  const invoiceAgentName = normalizeBillingMatchValue(invoice.agentName);
  const clientAgentName = normalizeBillingMatchValue(client.agentName);
  const invoiceAgentPhone = normalizeBillingMatchValue(invoice.agentPhone).replace(/\D/g, "");
  const clientAgentPhone = normalizeBillingMatchValue(client.agentPhone).replace(/\D/g, "");
  const invoiceEmails = new Set(parseGuestEmails(invoice.clientEmail || "").map(normalizeBillingMatchValue));
  const clientEmails = parseGuestEmails(client.email || "").map(normalizeBillingMatchValue);

  if (invoiceClientName && clientName && invoiceClientName === clientName) score += 6;
  if (invoiceClientName && clientName && (invoiceClientName.includes(clientName) || clientName.includes(invoiceClientName))) score += 3;
  if (invoiceAgentName && clientAgentName && invoiceAgentName === clientAgentName) score += 3;
  if (invoiceAgentPhone && clientAgentPhone && invoiceAgentPhone === clientAgentPhone) score += 3;
  if (clientEmails.some((email) => invoiceEmails.has(email))) score += 4;
  return score;
}

async function findInvoiceBillingClient(invoice) {
  try {
    const clients = await loadClients();
    return clients
      .map((client) => ({ client, score: invoiceBillingClientScore(invoice, client) }))
      .filter((item) => item.score > 0)
      .sort((a, b) => b.score - a.score)[0]?.client || null;
  } catch {
    return null;
  }
}

function invoiceBillingLines(invoice, client = null) {
  const billingName = String(invoice.billingName || invoice.clientName || client?.name || "").trim();
  const billingAbn = String(invoice.clientAbn || invoice.abn || client?.abn || "").trim();
  const shouldShowAbn = Number(invoice.total || 0) > 1000 && billingAbn;

  const lines = [
    billingName,
    shouldShowAbn ? `ABN: ${billingAbn}` : ""
  ].filter(Boolean);

  return lines.length ? lines : [invoice.propertyAddress || "Client"];
}

function drawInvoicePdf(doc, invoice, billingClient = null) {
  const pageWidth = doc.page.width;
  const margin = 52;
  const logoPath = path.join(publicDir, "openframe-logo.png");

  doc.rect(0, 0, pageWidth, 8).fill("#2ecad3");

  doc.fillColor("#000").font("Helvetica-Bold").fontSize(34).text("Tax Invoice", margin, 54);
  try {
    doc.image(logoPath, pageWidth - margin - 88, 48, { width: 88, height: 88, fit: [88, 88] });
  } catch {
    doc.fontSize(11).text("OPENFRAME\nSTUDIO", pageWidth - margin - 88, 62, { width: 88, align: "center" });
  }

  doc.font("Helvetica").fontSize(11);
  doc.text("Invoice Number", margin, 158);
  doc.text(invoice.invoiceNumber || "", 198, 158);
  doc.text("Invoice Date", margin, 181);
  doc.text(formatInvoiceDocumentDate(invoice.issuedAt), 198, 181);

  doc.font("Helvetica-Bold").fontSize(13).text("OUR INFORMATION", margin, 242);
  doc.font("Helvetica").fontSize(10.5);
  const infoLines = [
    "OpenFrame Studio Pty Ltd",
    "23 Selborne St",
    "Burwood",
    "NSW 2134",
    "ABN: 35 687 073 114",
    "Email: openframeau@gmail.com"
  ];
  infoLines.forEach((line, index) => doc.text(line, margin, 269 + index * 17));

  const billingX = 324;
  doc.font("Helvetica-Bold").fontSize(13).text("BILLING TO", billingX, 242);
  doc.font("Helvetica").fontSize(10.5).text(invoiceBillingLines(invoice, billingClient).join("\n"), billingX, 269, {
    width: pageWidth - margin - billingX,
    height: 82,
    ellipsis: true
  });

  const tableX = margin;
  const itemRows = invoicePdfItems(invoice);
  const hasMultipleRows = itemRows.length > 1;
  const tableY = hasMultipleRows ? 370 : 392;
  const columnWidths = [182, 82, 112, 115];
  const rowHeight = itemRows.length > 3 ? 36 : (hasMultipleRows ? 44 : 58);
  const headerHeight = hasMultipleRows ? 30 : 34;
  const tableWidth = columnWidths.reduce((sum, width) => sum + width, 0);
  const headers = ["PRODUCT", "QUANTITY", "PRICE", "SUBTOTAL"];

  doc.rect(tableX, tableY, tableWidth, headerHeight).fill("#e7e7e7");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(10);
  let cursorX = tableX;
  headers.forEach((header, index) => {
    drawPdfCell(doc, header, cursorX, tableY, columnWidths[index], headerHeight, { align: "center" });
    cursorX += columnWidths[index];
  });

  doc.font("Helvetica").fontSize(hasMultipleRows ? 9 : 10);
  itemRows.forEach((item, index) => {
    cursorX = tableX;
    const rowY = tableY + headerHeight + index * rowHeight;
    drawPdfCell(doc, item.product, cursorX, rowY, columnWidths[0], rowHeight, { align: "center" });
    cursorX += columnWidths[0];
    drawPdfCell(doc, formatInvoiceQuantity(item.quantity), cursorX, rowY, columnWidths[1], rowHeight, { align: "center" });
    cursorX += columnWidths[1];
    drawPdfCell(doc, formatInvoiceMoney(item.unitPrice), cursorX, rowY, columnWidths[2], rowHeight, { align: "right" });
    cursorX += columnWidths[2];
    drawPdfCell(doc, formatInvoiceMoney(item.amount), cursorX, rowY, columnWidths[3], rowHeight, { align: "right" });
    doc.moveTo(tableX, rowY + rowHeight).lineTo(tableX + tableWidth, rowY + rowHeight).lineWidth(1.2).stroke("#000");
  });

  const tableBottom = tableY + headerHeight + itemRows.length * rowHeight;
  const totalY = hasMultipleRows ? Math.max(556, tableBottom + 34) : Math.max(524, tableBottom + 46);

  doc.save();
  const stampY = hasMultipleRows ? Math.max(584, tableBottom + 44) : Math.max(559, totalY + 42);
  doc.rotate(-10, { origin: [140, stampY + 16] });
  doc.rect(110, stampY, 96, 30).lineWidth(2).stroke("#f5a623");
  doc.fillColor("#f5a623").font("Helvetica-Bold").fontSize(18).text(invoice.status === "paid" ? "PAID" : invoice.status === "void" ? "VOID" : "UNPAID", 118, stampY + 6);
  doc.restore();

  const totalX = 324;
  const totalW = 219;
  const totalRowH = hasMultipleRows ? 30 : 33;
  const totalRows = [
    ["Total", formatInvoiceMoney(invoice.subtotal)],
    ["GST (10%)", formatInvoiceMoney(invoice.gstAmount)],
    ["Total Due", formatInvoiceMoney(invoice.total)]
  ];
  doc.lineWidth(1.4).strokeColor("#000").rect(totalX, totalY, totalW, totalRowH * totalRows.length).stroke();
  doc.font("Helvetica-Bold").fontSize(10).fillColor("#000");
  totalRows.forEach(([label, value], index) => {
    const y = totalY + index * totalRowH;
    if (index > 0) doc.moveTo(totalX, y).lineTo(totalX + totalW, y).stroke();
    doc.moveTo(totalX + totalW / 2, y).lineTo(totalX + totalW / 2, y + totalRowH).stroke();
    drawPdfCell(doc, label, totalX, y, totalW / 2, totalRowH, { align: "center" });
    drawPdfCell(doc, value, totalX + totalW / 2, y, totalW / 2, totalRowH, { align: "center" });
  });

  const paymentY = hasMultipleRows
    ? Math.max(686, tableBottom + 78)
    : Math.max(676, totalY + totalRowH * totalRows.length + 62);
  doc.font("Helvetica-Bold").fontSize(13).fillColor("#000").text("PAYMENT INFORMATION", margin, paymentY);
  doc.font("Helvetica").fontSize(10.5).fillColor("#787878");
  [
    "Bank Transfer:",
    "Name: Openframe Studio Pty Ltd",
    "BSB: 062-128",
    "Account: 11440602",
    `Please reference ${invoice.invoiceNumber} for the payment`
  ].forEach((line, index) => doc.text(line, margin, paymentY + 28 + index * 17, {
    width: pageWidth - margin * 2,
    height: 16,
    ellipsis: true
  }));
}

async function createInvoicePdfBuffer(invoice) {
  let PDFDocument;
  try {
    ({ default: PDFDocument } = await import("pdfkit"));
  } catch {
    throw new Error("Invoice PDF generation is still installing. Wait for the latest deploy to finish, then try again.");
  }

  const doc = new PDFDocument({
    size: "A4",
    margin: 0,
    info: {
      Title: `${invoice.invoiceNumber} Tax Invoice`,
      Author: "OpenFrame Studio"
    }
  });
  const chunks = [];
  const finished = new Promise((resolve, reject) => {
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  const billingClient = await findInvoiceBillingClient(invoice);
  drawInvoicePdf(doc, invoice, billingClient);
  doc.end();
  return finished;
}

async function createInvoiceTransport() {
  let nodemailer;
  try {
    ({ default: nodemailer } = await import("nodemailer"));
  } catch {
    throw new Error("Invoice email sending is still installing. Wait for the latest deploy to finish, then try again.");
  }

  return nodemailer.createTransport({
    host: invoiceEmailConfig.host,
    port: invoiceEmailConfig.port,
    secure: invoiceEmailConfig.secure,
    requireTLS: !invoiceEmailConfig.secure,
    connectionTimeout: invoiceEmailConfig.timeoutMs,
    greetingTimeout: invoiceEmailConfig.timeoutMs,
    socketTimeout: invoiceEmailConfig.timeoutMs,
    auth: {
      user: invoiceEmailConfig.user,
      pass: invoiceEmailConfig.pass
    }
  });
}

async function sendResendEmail(payload, timeoutMs, timeoutMessage = "Resend email request timed out.") {
  const response = await withTimeout(fetch(`${resendConfig.apiBase}/emails`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${resendConfig.apiKey}`,
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  }), timeoutMs, timeoutMessage);
  const result = await response.json().catch(() => ({}));

  if (!response.ok) {
    const message = result.message || result.error?.message || result.error || `Resend email request failed with ${response.status}.`;
    throw new Error(message);
  }

  return result;
}

async function sendInvoiceWithResend(invoice, recipients, pdf, bcc = []) {
  const payload = {
    from: invoiceEmailConfig.from,
    to: recipients,
    subject: buildInvoiceEmailSubject(invoice),
    text: buildInvoiceEmailText(invoice),
    html: buildInvoiceEmailHtml(invoice),
    reply_to: invoiceEmailConfig.replyTo,
    attachments: [
      {
        filename: invoiceFileName(invoice),
        content: pdf.toString("base64")
      }
    ]
  };

  if (bcc.length) {
    payload.bcc = bcc;
  }

  return sendResendEmail(payload, invoiceEmailConfig.timeoutMs, "Resend email request timed out.");
}

async function withTimeout(promise, milliseconds, message) {
  let timeoutId;
  const timeout = new Promise((_, reject) => {
    timeoutId = setTimeout(() => reject(new Error(message)), milliseconds);
  });

  try {
    return await Promise.race([promise, timeout]);
  } finally {
    clearTimeout(timeoutId);
  }
}

function invoiceEmailErrorMessage(error) {
  const message = error?.message || "";
  const code = error?.code || "";
  const timeoutish = /timed out|timeout|greeting never received|connection/i.test(message)
    || ["ETIMEDOUT", "ESOCKET", "ECONNECTION", "EHOSTUNREACH", "ECONNREFUSED"].includes(code);

  if (timeoutish) {
    return "Could not connect to Lark Mail SMTP. Render free services block SMTP ports 465/587, so invoice email needs a paid Render service or an HTTP email provider.";
  }

  return message || "Could not send invoice email.";
}

async function sendInvoiceEmail(invoice, recipients) {
  if (!isInvoiceEmailConfigured()) {
    throw new Error(`Invoice email is not set up yet. Add ${invoiceEmailMissingSettings().join(", ")} in Render.`);
  }

  const pdf = await createInvoicePdfBuffer(invoice);
  const bcc = uniqueEmails(parseGuestEmails(invoiceEmailConfig.bcc)).filter(isValidEmail);

  if (invoiceEmailProvider === "resend") {
    const result = await sendInvoiceWithResend(invoice, recipients, pdf, bcc);
    return {
      providerMessageId: String(result?.id || result?.data?.id || "")
    };
  }

  const transport = await createInvoiceTransport();
  const result = await withTimeout(transport.sendMail({
    from: invoiceEmailConfig.from,
    to: recipients,
    bcc,
    replyTo: invoiceEmailConfig.replyTo,
    subject: buildInvoiceEmailSubject(invoice),
    text: buildInvoiceEmailText(invoice),
    html: buildInvoiceEmailHtml(invoice),
    attachments: [
      {
        filename: invoiceFileName(invoice),
        content: pdf,
        contentType: "application/pdf"
      }
    ]
  }), invoiceEmailConfig.timeoutMs + 2_000, "Invoice email timed out connecting to Lark Mail.");
  return {
    providerMessageId: String(result?.messageId || "")
  };
}

function normalizeWageStatus(status) {
  return ["draft", "paid", "void"].includes(status) ? status : "draft";
}

function normalizeEmployeeWageStatus(status) {
  return ["draft", "paid", "void"].includes(status) ? status : "draft";
}

function normalizeEmploymentType(value) {
  const type = String(value || "").trim().toLowerCase().replace(/[\s-]+/g, "_");
  return ["full_time", "part_time"].includes(type) ? type : "full_time";
}

function normalizePayPeriod(value) {
  const period = String(value || "").trim().toLowerCase();
  return ["monthly", "weekly", "fortnightly", "hourly"].includes(period) ? period : "monthly";
}

function lastDayOfMonthIso(value = new Date()) {
  const date = new Date(value);
  const base = Number.isNaN(date.getTime()) ? new Date() : date;
  const year = base.getUTCFullYear();
  const month = base.getUTCMonth();
  return new Date(Date.UTC(year, month + 1, 0, 12, 0, 0)).toISOString();
}

function employeePaymentSchedule(payPeriod) {
  return normalizePayPeriod(payPeriod) === "monthly" ? "last_day_of_month" : "";
}

function employeePaymentDate(wage) {
  const payPeriod = normalizePayPeriod(wage?.payPeriod || wage?.period);
  if (payPeriod === "monthly") {
    return lastDayOfMonthIso(new Date());
  }
  return wage?.nextPaymentAt || wage?.dueAt || wage?.issuedAt || new Date().toISOString();
}

function nextEmployeeWageNumber(employeeWages) {
  const maxNumber = employeeWages.reduce((max, wage) => {
    const value = String(wage.wageNumber || "");
    const match = value.match(/^E(\d+)$/);
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);
  return `E${String(maxNumber + 1).padStart(3, "0")}`;
}

function normalizeEmployeeWage(wage) {
  const amount = normalizeMoney(wage?.amount ?? wage?.total ?? 0);
  const status = normalizeEmployeeWageStatus(wage?.status);
  const payPeriod = normalizePayPeriod(wage?.payPeriod || wage?.period);
  const paymentSchedule = employeePaymentSchedule(payPeriod);

  return {
    id: wage?.id || crypto.randomUUID(),
    wageNumber: String(wage?.wageNumber || "").trim(),
    employeeName: String(wage?.employeeName || wage?.name || "").trim(),
    employmentType: normalizeEmploymentType(wage?.employmentType || wage?.type),
    payPeriod,
    amount,
    total: amount,
    currency: String(wage?.currency || employeeWageCurrency).trim().toUpperCase(),
    status,
    issuedAt: wage?.issuedAt || new Date().toISOString(),
    paymentSchedule,
    nextPaymentAt: paymentSchedule === "last_day_of_month"
      ? employeePaymentDate({ ...wage, payPeriod })
      : (wage?.nextPaymentAt || wage?.dueAt || wage?.issuedAt || new Date().toISOString()),
    paidAt: status === "paid" ? (wage?.paidAt || wage?.updatedAt || new Date().toISOString()) : "",
    voidedAt: status === "void" ? (wage?.voidedAt || wage?.updatedAt || new Date().toISOString()) : "",
    notes: String(wage?.notes || ""),
    createdAt: wage?.createdAt || new Date().toISOString(),
    updatedAt: wage?.updatedAt || wage?.createdAt || new Date().toISOString()
  };
}

function normalizeEmployeeWages(employeeWages) {
  return Array.isArray(employeeWages)
    ? employeeWages.map(normalizeEmployeeWage).filter((wage) => wage.employeeName && wage.amount > 0)
    : defaultEmployeeWages.map(normalizeEmployeeWage);
}

function validateEmployeeWage(input, existing = null, employeeWages = []) {
  const employeeName = String(input.employeeName || input.name || "").trim();
  const employmentType = normalizeEmploymentType(input.employmentType || input.type);
  const payPeriod = normalizePayPeriod(input.payPeriod || input.period);
  const amount = normalizeMoney(input.amount ?? input.total);
  const currency = String(input.currency || employeeWageCurrency).trim().toUpperCase();
  const notes = String(input.notes || "").trim();
  const errors = [];

  if (!employeeName) errors.push("Enter the employee name.");
  if (amount <= 0) errors.push("Enter the wage amount.");
  if (!/^[A-Z]{3}$/.test(currency)) errors.push("Use a valid 3-letter currency code.");

  return {
    errors,
    employeeWage: normalizeEmployeeWage({
      ...(existing || {}),
      id: existing?.id || input.id || crypto.randomUUID(),
      wageNumber: existing?.wageNumber || input.wageNumber || nextEmployeeWageNumber(employeeWages),
      employeeName,
      employmentType,
      payPeriod,
      amount,
      total: amount,
      currency,
      status: existing?.status || normalizeEmployeeWageStatus(input.status),
      issuedAt: existing?.issuedAt || new Date().toISOString(),
      paymentSchedule: existing?.paymentSchedule || employeePaymentSchedule(payPeriod),
      nextPaymentAt: existing?.nextPaymentAt || "",
      paidAt: existing?.paidAt || "",
      voidedAt: existing?.voidedAt || "",
      notes,
      createdAt: existing?.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    })
  };
}

function employeeWageFileName(wage) {
  const base = [wage.wageNumber || "payslip", wage.employeeName || "employee", "Pay-Slip"]
    .filter(Boolean)
    .join("-")
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 96);

  return `${base || "employee-payslip"}.pdf`;
}

function formatWageLabel(value) {
  return String(value || "")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (letter) => letter.toUpperCase());
}

function employeePaymentScheduleLabel(wage) {
  return wage.paymentSchedule === "last_day_of_month"
    ? "Last day of every month"
    : "Manual";
}

function drawEmployeePayslipPdf(doc, wage) {
  const pageWidth = doc.page.width;
  const margin = 52;
  const logoPath = path.join(publicDir, "openframe-logo.png");
  const currency = wage.currency || employeeWageCurrency;
  const total = Number(wage.amount || wage.total || 0);

  doc.rect(0, 0, pageWidth, 8).fill("#111611");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(34).text("Pay Slip", margin, 54);
  doc.font("Helvetica-Bold").fontSize(12).fillColor("#117a5d").text("Employee wage", margin, 94);
  try {
    doc.image(logoPath, pageWidth - margin - 88, 48, { width: 88, height: 88, fit: [88, 88] });
  } catch {
    doc.fontSize(11).fillColor("#000").text("OPENFRAME\nSTUDIO", pageWidth - margin - 88, 62, { width: 88, align: "center" });
  }

  doc.fillColor("#000").font("Helvetica").fontSize(11);
  doc.text("Payslip Number", margin, 158);
  doc.text(wage.wageNumber || "", 198, 158);
  doc.text("Issue Date", margin, 181);
  doc.text(formatInvoiceDocumentDate(wage.issuedAt), 198, 181);
  doc.text("Pay Period", margin, 204);
  doc.text(formatWageLabel(normalizePayPeriod(wage.payPeriod)), 198, 204);
  doc.text("Payment Date", margin, 227);
  doc.text(formatInvoiceDocumentDate(wage.nextPaymentAt), 198, 227);

  doc.font("Helvetica-Bold").fontSize(13).text("EMPLOYER", margin, 260);
  doc.font("Helvetica").fontSize(10.5);
  [
    "OpenFrame Studio Pty Ltd",
    "23 Selborne St",
    "Burwood NSW 2134",
    "admin@openframe.studio"
  ].forEach((line, index) => doc.text(line, margin, 286 + index * 17));

  const employeeX = 324;
  doc.font("Helvetica-Bold").fontSize(13).text("EMPLOYEE", employeeX, 260);
  doc.font("Helvetica").fontSize(10.5);
  [
    wage.employeeName || "Employee",
    formatWageLabel(normalizeEmploymentType(wage.employmentType)),
    `${formatWageLabel(normalizePayPeriod(wage.payPeriod))} wage`,
    employeePaymentScheduleLabel(wage)
  ].forEach((line, index) => doc.text(line, employeeX, 286 + index * 17, {
    width: pageWidth - margin - employeeX,
    height: 16,
    ellipsis: true
  }));

  const tableX = margin;
  const tableY = 392;
  const columnWidths = [250, 82, 159];
  const headerHeight = 34;
  const rowHeight = 48;
  const tableWidth = columnWidths.reduce((sum, width) => sum + width, 0);
  const headers = ["DESCRIPTION", "PERIOD", "AMOUNT"];

  doc.rect(tableX, tableY, tableWidth, headerHeight).fill("#eef3f0");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(10);
  let cursorX = tableX;
  headers.forEach((header, index) => {
    drawPdfCell(doc, header, cursorX, tableY, columnWidths[index], headerHeight, { align: "center" });
    cursorX += columnWidths[index];
  });

  const rowY = tableY + headerHeight;
  cursorX = tableX;
  doc.font("Helvetica").fontSize(10);
  drawPdfCell(doc, `${wage.employeeName || "Employee"} wage`, cursorX, rowY, columnWidths[0], rowHeight, { align: "left" });
  cursorX += columnWidths[0];
  drawPdfCell(doc, formatWageLabel(normalizePayPeriod(wage.payPeriod)), cursorX, rowY, columnWidths[1], rowHeight, { align: "center" });
  cursorX += columnWidths[1];
  drawPdfCell(doc, formatDocumentMoney(total, currency), cursorX, rowY, columnWidths[2], rowHeight, { align: "right" });
  doc.moveTo(tableX, rowY + rowHeight).lineTo(tableX + tableWidth, rowY + rowHeight).lineWidth(1).stroke("#cbd7d1");

  const totalX = 324;
  const totalY = 526;
  const totalRowHeight = 30;
  doc.lineWidth(1.4).strokeColor("#000").rect(totalX, totalY, 219, totalRowHeight * 2).stroke();
  [
    ["Gross Pay", formatDocumentMoney(total, currency)],
    ["Net Pay", formatDocumentMoney(total, currency)]
  ].forEach(([label, value], index) => {
    const y = totalY + index * totalRowHeight;
    if (index > 0) doc.moveTo(totalX, y).lineTo(totalX + 219, y).lineWidth(1).stroke("#cbd7d1");
    doc.font(index === 1 ? "Helvetica-Bold" : "Helvetica").fontSize(10.5).fillColor("#000");
    drawPdfCell(doc, label, totalX, y, 110, totalRowHeight, { align: "left", paddingY: 5 });
    drawPdfCell(doc, value, totalX + 110, y, 109, totalRowHeight, { align: "right", paddingY: 5 });
  });

  doc.save();
  doc.rotate(-10, { origin: [144, 620] });
  const stampColor = wage.status === "paid" ? "#117a5d" : wage.status === "void" ? "#a6403a" : "#f5a623";
  doc.rect(108, 604, 108, 30).lineWidth(2).stroke(stampColor);
  doc.fillColor(stampColor)
    .font("Helvetica-Bold").fontSize(18)
    .text(wage.status === "paid" ? "PAID" : wage.status === "void" ? "VOID" : "DRAFT", 116, 610);
  doc.restore();

  doc.font("Helvetica").fontSize(10).fillColor("#555");
  const notes = String(wage.notes || "").trim();
  if (notes) {
    doc.font("Helvetica-Bold").fillColor("#000").text("NOTES", margin, 650);
    doc.font("Helvetica").fontSize(9.5).fillColor("#555").text(notes, margin, 674, {
      width: pageWidth - margin * 2,
      height: 58,
      ellipsis: true
    });
  }

  doc.font("Helvetica").fontSize(9.5).fillColor("#787878");
  doc.text("This payslip is generated from the OpenFrame internal wage system.", margin, 742, {
    width: pageWidth - margin * 2
  });
}

async function createEmployeePayslipPdfBuffer(wage) {
  let PDFDocument;
  try {
    ({ default: PDFDocument } = await import("pdfkit"));
  } catch {
    throw new Error("Payslip PDF generation is still installing. Wait for the latest deploy to finish, then try again.");
  }

  const doc = new PDFDocument({
    size: "A4",
    margin: 0,
    info: {
      Title: `${wage.wageNumber} ${wage.employeeName} Pay Slip`,
      Author: "OpenFrame Studio"
    }
  });
  const chunks = [];
  const finished = new Promise((resolve, reject) => {
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  drawEmployeePayslipPdf(doc, wage);
  doc.end();
  return finished;
}

function normalizeWage(wage) {
  const items = Array.isArray(wage?.items)
    ? wage.items.map((item) => ({
        name: String(item.name || "Wage").trim() || "Wage",
        quantity: Math.max(1, Number(item.quantity || 1)),
        unitPrice: normalizeMoney(item.unitPrice),
        amount: normalizeMoney(item.amount ?? Number(item.quantity || 1) * Number(item.unitPrice || 0))
      }))
    : [];
  const total = normalizeMoney(wage?.total ?? items.reduce((sum, item) => sum + item.amount, 0));
  const rawGstIncluded = wage?.photographerGstIncluded ?? wage?.gstIncluded;
  const photographerGstIncluded = rawGstIncluded === undefined || rawGstIncluded === null || rawGstIncluded === ""
    ? Number(wage?.gstAmount || 0) > 0
    : normalizeEnvBoolean(rawGstIncluded, false);
  const gstRate = photographerGstIncluded ? invoiceGstRate : 0;
  const calculatedGstAmount = photographerGstIncluded
    ? normalizeMoney(total - (total / (1 + invoiceGstRate)))
    : 0;
  const gstAmount = photographerGstIncluded
    ? normalizeMoney(wage?.gstAmount ?? calculatedGstAmount)
    : 0;
  const subtotal = photographerGstIncluded
    ? normalizeMoney(wage?.subtotal ?? total - gstAmount)
    : total;

  return {
    id: wage?.id || crypto.randomUUID(),
    wageNumber: String(wage?.wageNumber || "").trim(),
    bookingId: String(wage?.bookingId || ""),
    propertyAddress: String(wage?.propertyAddress || ""),
    clientName: String(wage?.clientName || ""),
    agentName: String(wage?.agentName || ""),
    photographerName: String(wage?.photographerName || ""),
    photographerEmail: String(wage?.photographerEmail || ""),
    photographerPhone: String(wage?.photographerPhone || ""),
    bookingStartAt: wage?.bookingStartAt || "",
    bookingEndAt: wage?.bookingEndAt || "",
    services: Array.isArray(wage?.services) ? wage.services.map(String).filter(Boolean) : [],
    items,
    subtotal,
    gstRate,
    gstAmount,
    total,
    currency: wage?.currency || wageCurrency,
    photographerGstIncluded,
    status: normalizeWageStatus(wage?.status),
    issuedAt: wage?.issuedAt || new Date().toISOString(),
    dueAt: wage?.dueAt || addDays(new Date(), 7).toISOString(),
    paidAt: wage?.paidAt || "",
    voidedAt: wage?.voidedAt || "",
    sentAt: wage?.sentAt || "",
    sentTo: uniqueEmails(Array.isArray(wage?.sentTo) ? wage.sentTo : parseGuestEmails(wage?.sentTo || "")),
    notes: String(wage?.notes || ""),
    createdAt: wage?.createdAt || new Date().toISOString(),
    updatedAt: wage?.updatedAt || wage?.createdAt || new Date().toISOString()
  };
}

function normalizeWages(wages) {
  return Array.isArray(wages)
    ? wages.map(normalizeWage).filter((wage) => wage.bookingId || wage.wageNumber)
    : [];
}

function photographerWageItems(booking) {
  const serviceSet = new Set(bookingServiceNames(booking));
  const hasPhotography = serviceSet.has("Photography");
  const hasFloorplan = serviceSet.has("Floorplan");
  const hasDrone = serviceSet.has("Drone");
  const hasSiteplan = serviceSet.has("Siteplan");
  const items = [];

  if (hasDrone) {
    items.push({
      name: hasFloorplan ? "Photography + Floorplan + Drone" : "Photography + Drone",
      quantity: 1,
      unitPrice: 170,
      amount: 170
    });
  } else if (hasFloorplan) {
    items.push({
      name: hasPhotography ? "Photography + Floorplan" : "Floorplan package",
      quantity: 1,
      unitPrice: 120,
      amount: 120
    });
  } else if (hasPhotography) {
    items.push({
      name: "Photography",
      quantity: 1,
      unitPrice: 90,
      amount: 90
    });
  }

  if (hasSiteplan) {
    items.push({
      name: "Siteplan add-on",
      quantity: 1,
      unitPrice: 30,
      amount: 30
    });
  }

  return items;
}

function nextWageNumber(wages) {
  const maxNumber = wages.reduce((max, wage) => {
    const value = String(wage.wageNumber || "");
    const match = value.match(/^W(\d+)$/);
    return match ? Math.max(max, Number(match[1])) : max;
  }, 0);
  return `W${String(maxNumber + 1).padStart(3, "0")}`;
}

function wageFromBooking(booking, wages, existing = null) {
  const now = new Date().toISOString();
  const isPaid = existing?.status === "paid";
  const existingItems = Array.isArray(existing?.items) ? existing.items : [];
  const items = isPaid ? existingItems : photographerWageItems(booking);
  const total = normalizeMoney(items.reduce((sum, item) => sum + item.amount, 0));
  const photographerGstIncluded = isPaid
    ? normalizeEnvBoolean(existing?.photographerGstIncluded ?? existing?.gstIncluded, false)
    : normalizeEnvBoolean(booking.photographerGstIncluded, false);
  const status = booking.status === "cancelled"
    ? (isPaid ? "paid" : "void")
    : (isPaid ? "paid" : "draft");

  return normalizeWage({
    ...(existing || {}),
    id: existing?.id || crypto.randomUUID(),
    wageNumber: existing?.wageNumber || nextWageNumber(wages),
    bookingId: booking.id,
    propertyAddress: booking.propertyAddress || "",
    clientName: booking.clientName || "",
    agentName: booking.agentName || "",
    photographerName: booking.photographerName || "",
    photographerEmail: booking.photographerEmail || "",
    photographerPhone: booking.photographerPhone || "",
    photographerGstIncluded,
    bookingStartAt: booking.startAt || "",
    bookingEndAt: booking.endAt || "",
    services: bookingServiceNames(booking),
    items,
    subtotal: undefined,
    gstRate: undefined,
    gstAmount: undefined,
    total,
    currency: wageCurrency,
    status,
    issuedAt: existing?.issuedAt || now,
    dueAt: existing?.dueAt || addDays(new Date(), 7).toISOString(),
    paidAt: existing?.paidAt || "",
    voidedAt: status === "void" ? (existing?.voidedAt || now) : "",
    notes: "Generated from internal booking.",
    createdAt: existing?.createdAt || now,
    updatedAt: now
  });
}

function upsertWageForBooking(wages, booking) {
  if (!booking?.id || booking.larkOnly) {
    return { wage: null, created: false, updated: false };
  }

  const index = wages.findIndex((wage) => wage.bookingId === booking.id);
  const existing = index >= 0 ? wages[index] : null;
  if (booking.status === "cancelled" && !existing) {
    return { wage: null, created: false, updated: false };
  }

  const wage = wageFromBooking(booking, wages, existing);
  if (!wage.items.length && !existing) {
    return { wage: null, created: false, updated: false };
  }

  if (index >= 0) {
    wages[index] = wage;
    return { wage, created: false, updated: true };
  }

  wages.push(wage);
  return { wage, created: true, updated: false };
}

function syncWagesFromBookings(wages, bookings) {
  let created = 0;
  let updated = 0;

  for (const booking of bookings.filter((item) => !item.larkOnly)) {
    const result = upsertWageForBooking(wages, booking);
    if (result.created) created += 1;
    if (result.updated) updated += 1;
  }

  return { created, updated, total: wages.length };
}

function wageFileName(wage) {
  const base = [wage.wageNumber || "wage", wage.propertyAddress || ""]
    .filter(Boolean)
    .join("-")
    .replace(/[^a-z0-9._-]+/gi, "-")
    .replace(/-+/g, "-")
    .replace(/^-|-$/g, "")
    .slice(0, 96);

  return `${base || "wage-proforma"}.pdf`;
}

function australianFinancialYearBounds(startYearValue) {
  const startYear = Number.parseInt(String(startYearValue || ""), 10);
  if (!Number.isInteger(startYear) || startYear < 2000 || startYear > 2100) {
    return null;
  }

  return {
    startYear,
    label: `${startYear}-${startYear + 1}`,
    startMs: Date.UTC(startYear, 6, 1, 0, 0, 0),
    endMs: Date.UTC(startYear + 1, 6, 1, 0, 0, 0)
  };
}

function recordInFinancialYear(record, bounds, dateFields = ["issuedAt"]) {
  for (const field of dateFields) {
    const value = record?.[field];
    if (!value) continue;
    const time = new Date(value).getTime();
    if (Number.isFinite(time)) {
      if (time >= bounds.startMs && time < bounds.endMs) {
        return true;
      }
    }
  }

  return false;
}

function safeArchiveName(value, fallback = "document") {
  const name = String(value || fallback)
    .replace(/[<>:"/\\|?*\x00-\x1f]+/g, "-")
    .replace(/\s+/g, " ")
    .replace(/-+/g, "-")
    .trim()
    .slice(0, 120);
  return name || fallback;
}

function uniqueArchiveEntryName(folder, filename, usedNames) {
  const safeFolder = safeArchiveName(folder, "Documents");
  const safeFile = safeArchiveName(filename, "document.pdf");
  const extensionIndex = safeFile.lastIndexOf(".");
  const stem = extensionIndex > 0 ? safeFile.slice(0, extensionIndex) : safeFile;
  const extension = extensionIndex > 0 ? safeFile.slice(extensionIndex) : "";
  let candidate = `${safeFolder}/${safeFile}`;
  let suffix = 2;

  while (usedNames.has(candidate.toLowerCase())) {
    candidate = `${safeFolder}/${stem}-${suffix}${extension}`;
    suffix += 1;
  }

  usedNames.add(candidate.toLowerCase());
  return candidate;
}

let zipCrcTable = null;

function crc32(buffer) {
  if (!zipCrcTable) {
    zipCrcTable = Array.from({ length: 256 }, (_, index) => {
      let crc = index;
      for (let bit = 0; bit < 8; bit += 1) {
        crc = crc & 1 ? (0xedb88320 ^ (crc >>> 1)) : (crc >>> 1);
      }
      return crc >>> 0;
    });
  }

  let crc = 0xffffffff;
  for (const byte of buffer) {
    crc = zipCrcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function zipDosDateTime(value) {
  const inputDate = value instanceof Date ? value : new Date(value || Date.now());
  const date = Number.isNaN(inputDate.getTime()) ? new Date() : inputDate;
  const year = Math.max(1980, date.getFullYear());
  return {
    time: (date.getHours() << 11) | (date.getMinutes() << 5) | Math.floor(date.getSeconds() / 2),
    date: ((year - 1980) << 9) | ((date.getMonth() + 1) << 5) | date.getDate()
  };
}

function createStoredZip(entries) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const entry of entries) {
    const data = Buffer.isBuffer(entry.data) ? entry.data : Buffer.from(String(entry.data || ""));
    const nameBuffer = Buffer.from(entry.name.replace(/\\/g, "/"), "utf8");
    const checksum = crc32(data);
    const { time, date } = zipDosDateTime(entry.date);

    const localHeader = Buffer.alloc(30);
    localHeader.writeUInt32LE(0x04034b50, 0);
    localHeader.writeUInt16LE(20, 4);
    localHeader.writeUInt16LE(0, 6);
    localHeader.writeUInt16LE(0, 8);
    localHeader.writeUInt16LE(time, 10);
    localHeader.writeUInt16LE(date, 12);
    localHeader.writeUInt32LE(checksum, 14);
    localHeader.writeUInt32LE(data.length, 18);
    localHeader.writeUInt32LE(data.length, 22);
    localHeader.writeUInt16LE(nameBuffer.length, 26);
    localHeader.writeUInt16LE(0, 28);
    localParts.push(localHeader, nameBuffer, data);

    const centralHeader = Buffer.alloc(46);
    centralHeader.writeUInt32LE(0x02014b50, 0);
    centralHeader.writeUInt16LE(20, 4);
    centralHeader.writeUInt16LE(20, 6);
    centralHeader.writeUInt16LE(0, 8);
    centralHeader.writeUInt16LE(0, 10);
    centralHeader.writeUInt16LE(time, 12);
    centralHeader.writeUInt16LE(date, 14);
    centralHeader.writeUInt32LE(checksum, 16);
    centralHeader.writeUInt32LE(data.length, 20);
    centralHeader.writeUInt32LE(data.length, 24);
    centralHeader.writeUInt16LE(nameBuffer.length, 28);
    centralHeader.writeUInt16LE(0, 30);
    centralHeader.writeUInt16LE(0, 32);
    centralHeader.writeUInt16LE(0, 34);
    centralHeader.writeUInt16LE(0, 36);
    centralHeader.writeUInt32LE(0, 38);
    centralHeader.writeUInt32LE(offset, 42);
    centralParts.push(centralHeader, nameBuffer);

    offset += localHeader.length + nameBuffer.length + data.length;
  }

  const centralDirectory = Buffer.concat(centralParts);
  const endRecord = Buffer.alloc(22);
  endRecord.writeUInt32LE(0x06054b50, 0);
  endRecord.writeUInt16LE(0, 4);
  endRecord.writeUInt16LE(0, 6);
  endRecord.writeUInt16LE(entries.length, 8);
  endRecord.writeUInt16LE(entries.length, 10);
  endRecord.writeUInt32LE(centralDirectory.length, 12);
  endRecord.writeUInt32LE(offset, 16);
  endRecord.writeUInt16LE(0, 20);

  return Buffer.concat([...localParts, centralDirectory, endRecord]);
}

function financialYearExportFileName(bounds) {
  return safeArchiveName(`OpenFrame Studio - FY ${bounds.label} - invoices and wages.zip`, "financial-year-export.zip");
}

async function createFinancialYearExportZip(bounds) {
  const [invoices, employeeWages, contractorWages] = await Promise.all([
    loadInvoices(),
    loadEmployeeWages(),
    loadWages()
  ]);
  const usedNames = new Set();
  const entries = [];

  const invoiceMatches = invoices
    .filter((invoice) => recordInFinancialYear(invoice, bounds, ["issuedAt", "createdAt"]))
    .sort((a, b) => new Date(a.issuedAt) - new Date(b.issuedAt));
  const employeeMatches = employeeWages
    .filter((wage) => recordInFinancialYear(wage, bounds, ["issuedAt", "nextPaymentAt", "createdAt"]))
    .sort((a, b) => new Date(a.issuedAt) - new Date(b.issuedAt));
  const contractorMatches = contractorWages
    .filter((wage) => recordInFinancialYear(wage, bounds, ["issuedAt", "bookingStartAt", "createdAt"]))
    .sort((a, b) => new Date(a.issuedAt) - new Date(b.issuedAt));

  for (const invoice of invoiceMatches) {
    entries.push({
      name: uniqueArchiveEntryName("Invoices", invoiceFileName(invoice), usedNames),
      data: await createInvoicePdfBuffer(invoice),
      date: invoice.issuedAt
    });
  }

  for (const wage of employeeMatches) {
    entries.push({
      name: uniqueArchiveEntryName("Employee wages", employeeWageFileName(wage), usedNames),
      data: await createEmployeePayslipPdfBuffer(wage),
      date: wage.issuedAt
    });
  }

  for (const wage of contractorMatches) {
    entries.push({
      name: uniqueArchiveEntryName("Contractor wages", wageFileName(wage), usedNames),
      data: await createWagePdfBuffer(wage),
      date: wage.issuedAt
    });
  }

  if (!entries.length) {
    entries.push({
      name: "README.txt",
      data: [
        `OpenFrame Studio financial year export: ${bounds.label}`,
        "",
        "No invoices, employee wages, or contractor wages were found for this Australian financial year.",
        "Australian financial years run from 1 July to 30 June."
      ].join("\n"),
      date: new Date()
    });
  }

  return {
    buffer: createStoredZip(entries),
    counts: {
      invoices: invoiceMatches.length,
      employeeWages: employeeMatches.length,
      contractorWages: contractorMatches.length,
      total: invoiceMatches.length + employeeMatches.length + contractorMatches.length
    }
  };
}

function wageEmailRecipients(wage, input = {}) {
  const requestedRecipients = Array.isArray(input.to || input.recipients)
    ? (input.to || input.recipients)
    : parseGuestEmails(input.to || input.recipients || "");
  const fallbackRecipients = parseGuestEmails(wage.photographerEmail || "");
  return uniqueEmails(requestedRecipients.length ? requestedRecipients : fallbackRecipients).filter(isValidEmail);
}

function buildWageEmailSubject(wage) {
  return `Photographer Proforma ${wage.wageNumber} - ${wage.propertyAddress || "OpenFrame Studio"}`;
}

function buildWageEmailText(wage) {
  const gstLine = wage.photographerGstIncluded
    ? `GST included in total: ${formatInvoiceMoney(wage.gstAmount)}.`
    : "GST: not included.";

  return [
    `Hi ${wage.photographerName || "there"},`,
    "",
    `Please find attached proforma ${wage.wageNumber} for ${wage.propertyAddress || "your booking"}.`,
    `Amount: ${formatInvoiceMoney(wage.total)}.`,
    gstLine,
    "",
    "Thank you,",
    "OpenFrame Studio"
  ].join("\n");
}

function buildWageEmailHtml(wage) {
  const gstLine = wage.photographerGstIncluded
    ? `GST included in total: ${formatInvoiceMoney(wage.gstAmount)}.`
    : "GST: not included.";

  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p>Hi ${escapeHtmlForEmail(wage.photographerName || "there")},</p>
      <p>Please find attached proforma <strong>${escapeHtmlForEmail(wage.wageNumber)}</strong> for ${escapeHtmlForEmail(wage.propertyAddress || "your booking")}.</p>
      <p><strong>Amount: ${escapeHtmlForEmail(formatInvoiceMoney(wage.total))}.</strong></p>
      <p>${escapeHtmlForEmail(gstLine)}</p>
      <p>Thank you,<br />OpenFrame Studio</p>
    </div>
  `;
}

function wageDescription(wage) {
  const serviceNames = wage.services?.length
    ? wage.services.join(" + ")
    : (wage.items || []).map((item) => item.name).join(" + ");
  return [wage.propertyAddress || "Booking", serviceNames].filter(Boolean).join("\n");
}

function drawWagePdf(doc, wage) {
  const pageWidth = doc.page.width;
  const margin = 52;
  const logoPath = path.join(publicDir, "openframe-logo.png");

  doc.rect(0, 0, pageWidth, 8).fill("#117a5d");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(34).text("Proforma", margin, 54);
  doc.font("Helvetica-Bold").fontSize(12).fillColor("#117a5d").text("Photographer wage", margin, 94);
  try {
    doc.image(logoPath, pageWidth - margin - 88, 48, { width: 88, height: 88, fit: [88, 88] });
  } catch {
    doc.fontSize(11).fillColor("#000").text("OPENFRAME\nSTUDIO", pageWidth - margin - 88, 62, { width: 88, align: "center" });
  }

  doc.fillColor("#000").font("Helvetica").fontSize(11);
  doc.text("Proforma Number", margin, 158);
  doc.text(wage.wageNumber || "", 198, 158);
  doc.text("Issue Date", margin, 181);
  doc.text(formatInvoiceDocumentDate(wage.issuedAt), 198, 181);
  doc.text("Booking Date", margin, 204);
  doc.text(formatInvoiceDocumentDate(wage.bookingStartAt), 198, 204);

  doc.font("Helvetica-Bold").fontSize(13).text("OPENFRAME", margin, 260);
  doc.font("Helvetica").fontSize(10.5);
  [
    "OpenFrame Studio Pty Ltd",
    "23 Selborne St",
    "Burwood NSW 2134",
    "admin@openframe.studio"
  ].forEach((line, index) => doc.text(line, margin, 286 + index * 17));

  const photographerX = 324;
  doc.font("Helvetica-Bold").fontSize(13).text("PHOTOGRAPHER", photographerX, 260);
  doc.font("Helvetica").fontSize(10.5);
  [
    wage.photographerName || "Photographer",
    wage.photographerEmail || "",
    wage.photographerPhone || ""
  ].filter(Boolean).forEach((line, index) => doc.text(line, photographerX, 286 + index * 17, {
    width: pageWidth - margin - photographerX,
    height: 16,
    ellipsis: true
  }));

  const tableX = margin;
  const tableY = 392;
  const columnWidths = [250, 82, 82, 77];
  const headerHeight = 34;
  const rowHeight = 40;
  const tableWidth = columnWidths.reduce((sum, width) => sum + width, 0);
  const headers = ["BOOKING / SERVICE", "QTY", "RATE", "AMOUNT"];

  doc.rect(tableX, tableY, tableWidth, headerHeight).fill("#e7f3ef");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(10);
  let cursorX = tableX;
  headers.forEach((header, index) => {
    drawPdfCell(doc, header, cursorX, tableY, columnWidths[index], headerHeight, { align: "center" });
    cursorX += columnWidths[index];
  });

  let rowY = tableY + headerHeight;
  doc.font("Helvetica").fontSize(10);
  for (const item of wage.items || []) {
    cursorX = tableX;
    drawPdfCell(doc, `${wageDescription(wage)}\n${item.name}`, cursorX, rowY, columnWidths[0], rowHeight, { align: "left" });
    cursorX += columnWidths[0];
    drawPdfCell(doc, String(item.quantity || 1), cursorX, rowY, columnWidths[1], rowHeight, { align: "center" });
    cursorX += columnWidths[1];
    drawPdfCell(doc, formatInvoiceMoney(item.unitPrice), cursorX, rowY, columnWidths[2], rowHeight, { align: "right" });
    cursorX += columnWidths[2];
    drawPdfCell(doc, formatInvoiceMoney(item.amount), cursorX, rowY, columnWidths[3], rowHeight, { align: "right" });
    doc.moveTo(tableX, rowY + rowHeight).lineTo(tableX + tableWidth, rowY + rowHeight).lineWidth(1).stroke("#d3ddd8");
    rowY += rowHeight;
  }

  const totalX = 332;
  const totalY = Math.max(rowY + 32, 540);
  const totalRowHeight = 24;
  const totalRows = [
    {
      label: wage.photographerGstIncluded ? "Subtotal ex GST" : "Subtotal",
      amount: formatInvoiceMoney(wage.subtotal ?? wage.total),
      bold: false
    },
    {
      label: "GST",
      amount: wage.photographerGstIncluded ? formatInvoiceMoney(wage.gstAmount) : "No GST",
      bold: false
    },
    {
      label: "Total Payable",
      amount: formatInvoiceMoney(wage.total),
      bold: true
    }
  ];
  doc.lineWidth(1.4).strokeColor("#000").rect(totalX, totalY, 211, totalRowHeight * totalRows.length).stroke();
  totalRows.forEach((row, index) => {
    const y = totalY + index * totalRowHeight;
    if (index > 0) doc.moveTo(totalX, y).lineTo(totalX + 211, y).lineWidth(1).stroke("#d3ddd8");
    doc.font(row.bold ? "Helvetica-Bold" : "Helvetica").fontSize(10.5).fillColor("#000");
    drawPdfCell(doc, row.label, totalX, y, 112, totalRowHeight, { align: "left", paddingY: 4 });
    drawPdfCell(doc, row.amount, totalX + 112, y, 99, totalRowHeight, { align: "right", paddingY: 4 });
  });

  doc.save();
  doc.rotate(-10, { origin: [144, 620] });
  doc.rect(108, 604, 108, 30).lineWidth(2).stroke(wage.status === "paid" ? "#117a5d" : wage.status === "void" ? "#a6403a" : "#f5a623");
  doc.fillColor(wage.status === "paid" ? "#117a5d" : wage.status === "void" ? "#a6403a" : "#f5a623")
    .font("Helvetica-Bold").fontSize(18)
    .text(wage.status === "paid" ? "PAID" : wage.status === "void" ? "VOID" : "DRAFT", 116, 610);
  doc.restore();

  doc.font("Helvetica").fontSize(9.5).fillColor("#787878");
  doc.text("This proforma is generated from the internal booking system for photographer wage tracking.", margin, 704, {
    width: pageWidth - margin * 2
  });
}

async function createWagePdfBuffer(wage) {
  let PDFDocument;
  try {
    ({ default: PDFDocument } = await import("pdfkit"));
  } catch {
    throw new Error("Wage PDF generation is still installing. Wait for the latest deploy to finish, then try again.");
  }

  const doc = new PDFDocument({
    size: "A4",
    margin: 0,
    info: {
      Title: `${wage.wageNumber} Photographer Proforma`,
      Author: "OpenFrame Studio"
    }
  });
  const chunks = [];
  const finished = new Promise((resolve, reject) => {
    doc.on("data", (chunk) => chunks.push(chunk));
    doc.on("end", () => resolve(Buffer.concat(chunks)));
    doc.on("error", reject);
  });

  drawWagePdf(doc, wage);
  doc.end();
  return finished;
}

async function sendWageWithResend(wage, recipients, pdf, bcc = []) {
  const payload = {
    from: invoiceEmailConfig.from,
    to: recipients,
    subject: buildWageEmailSubject(wage),
    text: buildWageEmailText(wage),
    html: buildWageEmailHtml(wage),
    reply_to: invoiceEmailConfig.replyTo,
    attachments: [
      {
        filename: wageFileName(wage),
        content: pdf.toString("base64")
      }
    ]
  };

  if (bcc.length) {
    payload.bcc = bcc;
  }

  return sendResendEmail(payload, invoiceEmailConfig.timeoutMs, "Resend wage email request timed out.");
}

async function sendWageEmail(wage, recipients) {
  if (!isInvoiceEmailConfigured()) {
    throw new Error(`Wage proforma email is not set up yet. Add ${invoiceEmailMissingSettings().join(", ")} in Render.`);
  }

  const pdf = await createWagePdfBuffer(wage);
  const bcc = uniqueEmails(parseGuestEmails(invoiceEmailConfig.bcc)).filter(isValidEmail);

  if (invoiceEmailProvider === "resend") {
    await sendWageWithResend(wage, recipients, pdf, bcc);
    return;
  }

  const transport = await createInvoiceTransport();
  await withTimeout(transport.sendMail({
    from: invoiceEmailConfig.from,
    to: recipients,
    bcc,
    replyTo: invoiceEmailConfig.replyTo,
    subject: buildWageEmailSubject(wage),
    text: buildWageEmailText(wage),
    html: buildWageEmailHtml(wage),
    attachments: [
      {
        filename: wageFileName(wage),
        content: pdf,
        contentType: "application/pdf"
      }
    ]
  }), invoiceEmailConfig.timeoutMs + 2_000, "Wage proforma email timed out connecting to Lark Mail.");
}

function workEmployeeForAssignment(workState = null, assignment = null) {
  return workEmployeeById(workState, assignment?.employeeId || defaultWorkEmployee.id);
}

function isDefaultWorkEmployee(employee) {
  return employee?.id === defaultWorkEmployee.id;
}

function workInviteRecipients(workState = null, assignment = null) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const employeeEmails = parseGuestEmails(employee?.email || "");
  const fallbackEmails = isDefaultWorkEmployee(employee) ? parseGuestEmails(workInviteEmailConfig.to) : [];
  return uniqueEmails(employeeEmails.length ? employeeEmails : fallbackEmails).filter(isValidEmail);
}

function workInviteEmailMissingSettings(workState = null, assignment = null) {
  const missing = [];
  if (!workInviteEmailConfig.enabled) missing.push("WORK_INVITE_EMAIL_ENABLED");
  if (!resendConfig.apiKey) missing.push("RESEND_API_KEY");
  if (!workInviteRecipients(workState, assignment).length) missing.push("WORK_INVITE_EMAIL_TO");
  return missing;
}

function isWorkInviteEmailConfigured() {
  return workInviteEmailMissingSettings().length === 0;
}

function normalizeLarkMessageReceiveIdType(value) {
  const receiveIdType = String(value || "email").trim().toLowerCase();
  return larkMessageReceiveIdTypes.has(receiveIdType) ? receiveIdType : "email";
}

function workLarkReceiveIdType(workState = null, assignment = null) {
  const employee = workEmployeeForAssignment(workState, assignment);
  return normalizeLarkMessageReceiveIdType(
    employee?.larkReceiveIdType || workLarkNotificationConfig.receiveIdType
  );
}

function workLarkReceiveId(workState = null, assignment = null) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const employeeId = String(employee?.larkReceiveId || "").trim();
  if (employeeId) return employeeId;

  const configuredId = String(workLarkNotificationConfig.receiveId || "").trim();
  if (configuredId && isDefaultWorkEmployee(employee)) return configuredId;

  if (workLarkReceiveIdType(workState, assignment) === "email") {
    return workInviteRecipients(workState, assignment)[0] || "";
  }

  return "";
}

function workLarkNotificationMissingSettings(workState = null, assignment = null) {
  const missing = [];
  if (!workLarkNotificationConfig.enabled) missing.push("WORK_LARK_NOTIFICATIONS_ENABLED");
  if (!larkConfig.appId) missing.push("LARK_APP_ID");
  if (!larkConfig.appSecret) missing.push("LARK_APP_SECRET");
  if (!workLarkReceiveId(workState, assignment)) {
    missing.push(workLarkReceiveIdType(workState, assignment) === "email" ? "WORK_LARK_RECEIVE_ID or WORK_INVITE_EMAIL_TO" : "WORK_LARK_RECEIVE_ID");
  }
  return missing;
}

function isWorkLarkNotificationConfigured() {
  return workLarkNotificationMissingSettings().length === 0;
}

function workCompletionLarkReceiveIdType() {
  return normalizeLarkMessageReceiveIdType(workCompletionLarkNotificationConfig.receiveIdType);
}

function workCompletionLarkReceiveId() {
  return String(workCompletionLarkNotificationConfig.receiveId || "").trim();
}

function workCompletionLarkNotificationMissingSettings() {
  const missing = [];
  if (!workCompletionLarkNotificationConfig.enabled) missing.push("WORK_COMPLETION_LARK_NOTIFICATIONS_ENABLED");
  if (!larkConfig.appId) missing.push("LARK_APP_ID");
  if (!larkConfig.appSecret) missing.push("LARK_APP_SECRET");
  if (!workCompletionLarkReceiveId()) missing.push("WORK_COMPLETION_LARK_RECEIVE_ID or BOSS_LARK_RECEIVE_ID");
  return missing;
}

function isWorkCompletionLarkNotificationConfigured() {
  return workCompletionLarkNotificationMissingSettings().length === 0;
}

function workCompletionEmailRecipients() {
  return uniqueEmails(parseGuestEmails(workCompletionEmailConfig.to)).filter(isValidEmail);
}

function workCompletionEmailMissingSettings() {
  const missing = [];
  if (!workCompletionEmailConfig.enabled) missing.push("WORK_COMPLETION_EMAIL_ENABLED");
  if (!resendConfig.apiKey) missing.push("RESEND_API_KEY");
  if (!workCompletionEmailRecipients().length) missing.push("WORK_COMPLETION_EMAIL_TO");
  if (!workCompletionEmailConfig.from) missing.push("WORK_COMPLETION_EMAIL_FROM");
  return missing;
}

function isWorkCompletionEmailConfigured() {
  return workCompletionEmailMissingSettings().length === 0;
}

function formatWorkInviteDueDate(assignment) {
  const dueDate = parseDateValue(assignment.dueDate);
  if (Number.isNaN(dueDate.getTime())) return assignment.dueDate || "Not set";

  return new Intl.DateTimeFormat("en-AU", {
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric"
  }).format(dueDate);
}

function buildWorkInviteEmailSubject(assignment) {
  return `${workInviteEmailConfig.subjectPrefix}: ${assignment.title}`;
}

function buildWorkInviteEmailText(assignment, workState) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const lines = [
    `Hi ${employee?.name || "Faye"},`,
    "",
    "A new work item has been assigned to you in the OpenFrame internal booking system.",
    "",
    `Work: ${assignment.title}`,
    `Due: ${formatWorkInviteDueDate(assignment)}`,
    `Priority: ${assignment.priority}`,
    ""
  ];

  if (assignment.notes) {
    lines.push("Details:", assignment.notes, "");
  }

  if (assignment.attachments?.length) {
    lines.push(`Photos: ${assignment.attachments.length} attached in the work desk.`, "");
  }

  lines.push(
    "Open the work desk:",
    `${publicAppUrl}/work/`,
    "",
    "Thank you,",
    "OpenFrame Studio"
  );

  return lines.join("\n");
}

function buildWorkLarkNotificationText(assignment, workState) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const lines = [
    `${workLarkNotificationConfig.subjectPrefix}: ${assignment.title}`,
    "",
    `Hi ${employee?.name || "Faye"},`,
    `Due: ${formatWorkInviteDueDate(assignment)}`,
    `Priority: ${assignment.priority}`,
    ""
  ];

  if (assignment.notes) {
    const notes = assignment.notes.length > 900 ? `${assignment.notes.slice(0, 900)}...` : assignment.notes;
    lines.push("Details:", notes, "");
  }

  if (assignment.attachments?.length) {
    lines.push(`Photos: ${assignment.attachments.length} attached in the work desk.`, "");
  }

  lines.push("Open work desk:", `${publicAppUrl}/work/`);
  return lines.join("\n");
}

function buildWorkCompletionLarkNotificationText(assignment, workState, completedBy) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const completedByName = completedBy?.name || employee?.name || "Employee";
  const lines = [
    `${workCompletionLarkNotificationConfig.subjectPrefix}: ${assignment.title}`,
    "",
    `${completedByName} marked this work as finished.`,
    `Assigned to: ${employee?.name || "Employee"}`,
    `Due: ${formatWorkInviteDueDate(assignment)}`,
    `Priority: ${assignment.priority}`,
    ""
  ];

  if (assignment.notes) {
    const notes = assignment.notes.length > 700 ? `${assignment.notes.slice(0, 700)}...` : assignment.notes;
    lines.push("Details:", notes, "");
  }

  lines.push("Open work desk:", `${publicAppUrl}/work/`);
  return lines.join("\n");
}

function buildWorkCompletionEmailSubject(assignment) {
  return `${workCompletionEmailConfig.subjectPrefix}: ${assignment.title}`;
}

function buildWorkCompletionEmailText(assignment, workState, completedBy) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const completedByName = completedBy?.name || employee?.name || "Employee";
  const lines = [
    `${completedByName} marked this work as finished.`,
    "",
    `Work: ${assignment.title}`,
    `Completed by: ${completedByName}`,
    `Assigned to: ${employee?.name || "Employee"}`,
    `Due: ${formatWorkInviteDueDate(assignment)}`,
    `Priority: ${assignment.priority}`,
    ""
  ];

  if (assignment.notes) {
    lines.push("Details:", assignment.notes, "");
  }

  lines.push("Open work desk:", `${publicAppUrl}/work/`, "", "OpenFrame Studio");
  return lines.join("\n");
}

function buildWorkCompletionEmailHtml(assignment, workState, completedBy) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const completedByName = completedBy?.name || employee?.name || "Employee";
  const notes = assignment.notes
    ? `<p><strong>Details:</strong><br />${escapeHtmlForEmail(assignment.notes).replace(/\n/g, "<br />")}</p>`
    : "";

  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p><strong>${escapeHtmlForEmail(completedByName)}</strong> marked this work as finished.</p>
      <p><strong>Work:</strong> ${escapeHtmlForEmail(assignment.title)}</p>
      <p><strong>Completed by:</strong> ${escapeHtmlForEmail(completedByName)}</p>
      <p><strong>Assigned to:</strong> ${escapeHtmlForEmail(employee?.name || "Employee")}</p>
      <p><strong>Due:</strong> ${escapeHtmlForEmail(formatWorkInviteDueDate(assignment))}</p>
      <p><strong>Priority:</strong> ${escapeHtmlForEmail(assignment.priority)}</p>
      ${notes}
      <p><a href="${escapeHtmlForEmail(`${publicAppUrl}/work/`)}">Open the work desk</a></p>
      <p>OpenFrame Studio</p>
    </div>
  `;
}

function isInvalidLarkReceiveIdResult(result) {
  const code = Number(result?.code);
  const message = String(result?.msg || result?.message || "").toLowerCase();
  return code === larkInvalidReceiveIdCode || message.includes("invalid receive_id");
}

function larkMessageErrorText(response, result) {
  const message = result?.msg || result?.message || `Lark message failed with ${response.status}.`;
  if (String(message).toLowerCase().includes("no availability")) {
    return "Lark bot is not available to this employee yet. Ask them to open or add the OpenFrame Lark app/bot, or use the email invite instead.";
  }
  if (String(message).includes("contact:user.employee_id:readonly")) {
    return "Lark needs the contact:user.employee_id:readonly permission before it can send notifications by User ID. Add that permission in the Lark Developer app, publish/save it, then retry.";
  }
  if (String(message).includes("contact:user.id:readonly")) {
    return "Lark needs the contact:user.id:readonly permission before it can look up a user from an email address. Add that permission in the Lark Developer app, publish/save it, then retry.";
  }
  return message;
}

async function resolveLarkOpenIdByEmail(email, token) {
  if (!isValidEmail(email)) {
    return "";
  }

  const params = new URLSearchParams({ user_id_type: "open_id" });
  const response = await fetch(`${larkConfig.apiBase}/contact/v3/users/batch_get_id?${params}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({
      emails: [email],
      include_resigned: false
    })
  });
  const result = await response.json().catch(() => ({}));

  if (!response.ok || (result.code !== undefined && result.code !== 0)) {
    throw new Error(result.msg || result.message || `Lark user lookup failed with ${response.status}.`);
  }

  const user = result.data?.user_list?.[0] || result.data?.users?.[0] || null;
  return user?.open_id || user?.user_id || user?.id || "";
}

async function postWorkLarkMessage(token, receiveIdType, receiveId, text) {
  const params = new URLSearchParams({ receive_id_type: receiveIdType });
  const response = await fetch(`${larkConfig.apiBase}/im/v1/messages?${params}`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify({
      receive_id: receiveId,
      msg_type: "text",
      content: JSON.stringify({ text })
    })
  });
  const result = await response.json().catch(() => ({}));

  return {
    ok: response.ok && (result.code === undefined || result.code === 0),
    invalidReceiveId: isInvalidLarkReceiveIdResult(result),
    message: larkMessageErrorText(response, result),
    result
  };
}

function buildWorkInviteEmailHtml(assignment, workState) {
  const employee = workEmployeeForAssignment(workState, assignment);
  const notes = assignment.notes
    ? `<p><strong>Details:</strong><br />${escapeHtmlForEmail(assignment.notes).replace(/\n/g, "<br />")}</p>`
    : "";
  const photos = assignment.attachments?.length
    ? `<p><strong>Photos:</strong> ${assignment.attachments.length} attached in the work desk.</p>`
    : "";

  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p>Hi ${escapeHtmlForEmail(employee?.name || "Faye")},</p>
      <p>A new work item has been assigned to you in the OpenFrame internal booking system.</p>
      <p><strong>Work:</strong> ${escapeHtmlForEmail(assignment.title)}</p>
      <p><strong>Due:</strong> ${escapeHtmlForEmail(formatWorkInviteDueDate(assignment))}</p>
      <p><strong>Priority:</strong> ${escapeHtmlForEmail(assignment.priority)}</p>
      ${notes}
      ${photos}
      <p><a href="${escapeHtmlForEmail(`${publicAppUrl}/work/`)}">Open the work desk</a></p>
      <p>Thank you,<br />OpenFrame Studio</p>
    </div>
  `;
}

async function sendWorkLarkNotification(assignment, workState) {
  const assignee = workEmployeeForAssignment(workState, assignment);
  const missing = workLarkNotificationMissingSettings(workState, assignment);
  if (missing.length) {
    return { sent: false, skipped: true, missing };
  }

  const token = await getTenantAccessToken();
  const receiveIdType = workLarkReceiveIdType(workState, assignment);
  const receiveId = workLarkReceiveId(workState, assignment);
  const text = buildWorkLarkNotificationText(assignment, workState);
  const firstAttempt = await postWorkLarkMessage(token, receiveIdType, receiveId, text);
  let resolvedOpenId = "";
  let result = firstAttempt.result;

  if (!firstAttempt.ok && receiveIdType === "email" && firstAttempt.invalidReceiveId) {
    try {
      resolvedOpenId = await resolveLarkOpenIdByEmail(receiveId, token);
    } catch (error) {
      throw new Error(
        `Lark rejected ${receiveId} as a message recipient, and the app could not look up ${assignee?.name || "the assignee"}'s Lark user ID by email: ${error.message}.`
      );
    }

    if (resolvedOpenId) {
      const retryAttempt = await postWorkLarkMessage(token, "open_id", resolvedOpenId, text);
      result = retryAttempt.result;
      if (!retryAttempt.ok) {
        throw new Error(retryAttempt.message);
      }
    } else {
      throw new Error(
        `Lark rejected ${receiveId} as a message recipient, and no Lark user was found for that email. Set ${assignee?.name || "the assignee"}'s WORK_LARK_RECEIVE_ID_TYPE to user_id and WORK_LARK_RECEIVE_ID to their Lark User ID, or use open_id if you have their Lark open_id.`
      );
    }
  } else if (!firstAttempt.ok) {
    throw new Error(firstAttempt.message);
  }

  return {
    sent: true,
    receiveId: resolvedOpenId || receiveId,
    receiveIdType: resolvedOpenId ? "open_id" : receiveIdType,
    messageId: result.data?.message_id || result.data?.messageId || ""
  };
}

async function sendWorkCompletionLarkNotification(assignment, workState, completedBy) {
  const missing = workCompletionLarkNotificationMissingSettings();
  if (missing.length) {
    return { sent: false, skipped: true, missing };
  }

  const token = await getTenantAccessToken();
  const receiveIdType = workCompletionLarkReceiveIdType();
  const receiveId = workCompletionLarkReceiveId();
  const text = buildWorkCompletionLarkNotificationText(assignment, workState, completedBy);
  const firstAttempt = await postWorkLarkMessage(token, receiveIdType, receiveId, text);
  let resolvedOpenId = "";
  let result = firstAttempt.result;

  if (!firstAttempt.ok && receiveIdType === "email" && firstAttempt.invalidReceiveId) {
    try {
      resolvedOpenId = await resolveLarkOpenIdByEmail(receiveId, token);
    } catch (error) {
      throw new Error(
        `Lark rejected ${receiveId} as the boss notification recipient, and the app could not look up that Lark user ID by email: ${error.message}.`
      );
    }

    if (resolvedOpenId) {
      const retryAttempt = await postWorkLarkMessage(token, "open_id", resolvedOpenId, text);
      result = retryAttempt.result;
      if (!retryAttempt.ok) {
        throw new Error(retryAttempt.message);
      }
    } else {
      throw new Error(
        `Lark rejected ${receiveId} as the boss notification recipient, and no Lark user was found for that email. Set WORK_COMPLETION_LARK_RECEIVE_ID_TYPE to user_id and WORK_COMPLETION_LARK_RECEIVE_ID to your Lark User ID, or use open_id if you have your Lark open_id.`
      );
    }
  } else if (!firstAttempt.ok) {
    throw new Error(firstAttempt.message);
  }

  return {
    sent: true,
    receiveId: resolvedOpenId || receiveId,
    receiveIdType: resolvedOpenId ? "open_id" : receiveIdType,
    messageId: result.data?.message_id || result.data?.messageId || ""
  };
}

async function sendWorkCompletionEmail(assignment, workState, completedBy) {
  const missing = workCompletionEmailMissingSettings();
  if (missing.length) {
    return { sent: false, skipped: true, missing };
  }

  const recipients = workCompletionEmailRecipients();
  const bcc = uniqueEmails(parseGuestEmails(workCompletionEmailConfig.bcc)).filter(isValidEmail);
  const payload = {
    from: workCompletionEmailConfig.from,
    to: recipients,
    subject: buildWorkCompletionEmailSubject(assignment),
    text: buildWorkCompletionEmailText(assignment, workState, completedBy),
    html: buildWorkCompletionEmailHtml(assignment, workState, completedBy),
    reply_to: workCompletionEmailConfig.replyTo
  };

  if (bcc.length) {
    payload.bcc = bcc;
  }

  await sendResendEmail(payload, workCompletionEmailConfig.timeoutMs, "Work completion email request timed out.");
  return { sent: true, recipients };
}

async function sendWorkInviteEmail(assignment, workState) {
  const missing = workInviteEmailMissingSettings(workState, assignment);
  if (missing.length) {
    return { sent: false, skipped: true, missing };
  }

  const recipients = workInviteRecipients(workState, assignment);
  const bcc = uniqueEmails(parseGuestEmails(workInviteEmailConfig.bcc)).filter(isValidEmail);
  const payload = {
    from: workInviteEmailConfig.from,
    to: recipients,
    subject: buildWorkInviteEmailSubject(assignment),
    text: buildWorkInviteEmailText(assignment, workState),
    html: buildWorkInviteEmailHtml(assignment, workState),
    reply_to: workInviteEmailConfig.replyTo
  };

  if (bcc.length) {
    payload.bcc = bcc;
  }

  await sendResendEmail(payload, workInviteEmailConfig.timeoutMs, "Work invite email request timed out.");
  return { sent: true, recipients };
}

async function sendWorkLarkNotifications(workState, assignments) {
  const sendLogs = [];
  const summary = {
    attempted: assignments.length,
    sent: 0,
    skipped: 0,
    failed: 0,
    recipients: [],
    errors: [],
    missing: []
  };

  for (const assignment of assignments) {
    try {
      const result = await sendWorkLarkNotification(assignment, workState);
      const now = new Date().toISOString();

      if (result.sent) {
        summary.sent += 1;
        const recipient = `${result.receiveIdType}:${result.receiveId}`;
        summary.recipients = [...new Set([...summary.recipients, recipient])];
        assignment.larkNotifyStatus = "sent";
        assignment.larkNotifySentAt = now;
        assignment.larkNotifyTo = recipient;
        assignment.larkNotifyError = "";
        sendLogs.push({
          type: "work_notification",
          status: "success",
          title: `Work notification: ${assignment.title}`,
          detail: workEmployeeForAssignment(workState, assignment).name,
          provider: "lark",
          recipients: [recipient],
          relatedId: assignment.id
        });
      } else {
        summary.skipped += 1;
        summary.missing = [...new Set([...summary.missing, ...(result.missing || [])])];
        assignment.larkNotifyStatus = "not_configured";
        assignment.larkNotifyError = result.missing?.length
          ? `Missing ${result.missing.join(", ")}`
          : "Lark work notifications are not configured.";
        sendLogs.push({
          type: "work_notification",
          status: "skipped",
          title: `Work notification: ${assignment.title}`,
          detail: workEmployeeForAssignment(workState, assignment).name,
          provider: "lark",
          relatedId: assignment.id,
          error: assignment.larkNotifyError
        });
      }
    } catch (error) {
      summary.failed += 1;
      const message = error.message || "Could not send Lark work notification.";
      summary.errors.push(message);
      assignment.larkNotifyStatus = "failed";
      assignment.larkNotifyError = message;
      sendLogs.push({
        type: "work_notification",
        status: "failed",
        title: `Work notification: ${assignment.title}`,
        detail: workEmployeeForAssignment(workState, assignment).name,
        provider: "lark",
        recipients: [workEmployeeForAssignment(workState, assignment).larkReceiveId].filter(Boolean),
        relatedId: assignment.id,
        error: message
      });
    }
  }

  await appendSendLogs(sendLogs);
  return summary;
}

async function sendWorkInviteEmails(workState, assignments) {
  const sendLogs = [];
  const summary = {
    attempted: assignments.length,
    sent: 0,
    skipped: 0,
    failed: 0,
    recipients: [],
    errors: [],
    missing: []
  };

  for (const assignment of assignments) {
    try {
      const result = await sendWorkInviteEmail(assignment, workState);
      const now = new Date().toISOString();

      if (result.sent) {
        summary.sent += 1;
        summary.recipients = uniqueEmails([...summary.recipients, ...result.recipients]);
        assignment.inviteStatus = "sent";
        assignment.inviteSentAt = now;
        assignment.inviteTo = result.recipients.join(", ");
        assignment.inviteEmailFrom = workInviteEmailConfig.from;
        assignment.inviteError = "";
        sendLogs.push({
          type: "work_email",
          status: "success",
          title: `Work email: ${assignment.title}`,
          detail: workEmployeeForAssignment(workState, assignment).name,
          provider: "resend",
          from: workInviteEmailConfig.from,
          recipients: result.recipients,
          relatedId: assignment.id
        });
      } else {
        summary.skipped += 1;
        summary.missing = uniqueEmails([...summary.missing, ...(result.missing || [])]);
        assignment.inviteStatus = "not_configured";
        assignment.inviteEmailFrom = workInviteEmailConfig.from;
        assignment.inviteError = result.missing?.length
          ? `Missing ${result.missing.join(", ")}`
          : "Work invite email is not configured.";
        sendLogs.push({
          type: "work_email",
          status: "skipped",
          title: `Work email: ${assignment.title}`,
          detail: workEmployeeForAssignment(workState, assignment).name,
          provider: "resend",
          from: workInviteEmailConfig.from,
          relatedId: assignment.id,
          error: assignment.inviteError
        });
      }
    } catch (error) {
      summary.failed += 1;
      const message = error.message || "Could not send work invite email.";
      summary.errors.push(message);
      assignment.inviteStatus = "failed";
      assignment.inviteEmailFrom = workInviteEmailConfig.from;
      assignment.inviteError = message;
      sendLogs.push({
        type: "work_email",
        status: "failed",
        title: `Work email: ${assignment.title}`,
        detail: workEmployeeForAssignment(workState, assignment).name,
        provider: "resend",
        from: workInviteEmailConfig.from,
        recipients: workInviteRecipients(workState, assignment),
        relatedId: assignment.id,
        error: message
      });
    }
  }

  await appendSendLogs(sendLogs);
  return summary;
}

function workLarkNotificationMessage(summary, context = "saved") {
  if (!summary?.attempted) return "";
  if (summary.sent) {
    return "Lark notification sent to the employee.";
  }
  if (summary.failed) {
    return context === "manual"
      ? `Lark notification could not send: ${summary.errors[0] || "unknown error"}`
      : `Work saved, but Lark notification could not send: ${summary.errors[0] || "unknown error"}`;
  }
  if (summary.skipped) {
    const missing = summary.missing || [];
    const needsLarkApp = missing.some((item) => ["LARK_APP_ID", "LARK_APP_SECRET", "WORK_LARK_NOTIFICATIONS_ENABLED"].includes(item));
    if (needsLarkApp) {
      return context === "manual"
        ? "Lark notification is not fully configured yet."
        : "Work saved. Lark notification is not fully configured yet.";
    }

    return context === "manual"
      ? "Add the employee's Lark or email details to notify them in Lark."
      : "Work saved. Add the employee's Lark or email details to notify them in Lark.";
  }
  return "";
}

function workAssignmentNotificationMessage(workLarkSummary, workEmailSummary) {
  const larkMessage = workLarkNotificationMessage(workLarkSummary);
  const emailMessage = workInviteMessage(workEmailSummary);

  if (workLarkSummary?.sent && workEmailSummary?.skipped) {
    return larkMessage;
  }

  if (workEmailSummary?.sent && workLarkSummary?.failed) {
    return `${emailMessage} Lark bot notification could not send: ${workLarkSummary.errors[0] || "unknown error"}`;
  }

  if (workEmailSummary?.sent && workLarkSummary?.skipped) {
    return emailMessage;
  }

  return [larkMessage, emailMessage].filter(Boolean).join(" ");
}

async function trySendWorkCompletionLarkNotification(workState, assignment, completedBy) {
  const summary = {
    attempted: completedBy?.role === "employee" ? 1 : 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    recipients: [],
    errors: [],
    missing: []
  };

  if (!summary.attempted) {
    return summary;
  }

  try {
    const result = await sendWorkCompletionLarkNotification(assignment, workState, completedBy);
    const now = new Date().toISOString();

    if (result.sent) {
      summary.sent = 1;
      const recipient = `${result.receiveIdType}:${result.receiveId}`;
      summary.recipients = [recipient];
      assignment.completionNotifyStatus = "sent";
      assignment.completionNotifySentAt = now;
      assignment.completionNotifyTo = recipient;
      assignment.completionNotifyError = "";
      await appendSendLog({
        type: "work_notification",
        status: "success",
        title: `Completion notification: ${assignment.title}`,
        detail: completedBy?.name || completedBy?.username || "Employee",
        provider: "lark",
        recipients: [recipient],
        relatedId: assignment.id
      });
    } else {
      summary.skipped = 1;
      summary.missing = result.missing || [];
      assignment.completionNotifyStatus = "not_configured";
      assignment.completionNotifyError = result.missing?.length
        ? `Missing ${result.missing.join(", ")}`
        : "Completion Lark notification is not configured.";
      await appendSendLog({
        type: "work_notification",
        status: "skipped",
        title: `Completion notification: ${assignment.title}`,
        detail: completedBy?.name || completedBy?.username || "Employee",
        provider: "lark",
        relatedId: assignment.id,
        error: assignment.completionNotifyError
      });
    }
  } catch (error) {
    summary.failed = 1;
    const message = error.message || "Could not send completion Lark notification.";
    summary.errors = [message];
    assignment.completionNotifyStatus = "failed";
    assignment.completionNotifyError = message;
    await appendSendLog({
      type: "work_notification",
      status: "failed",
      title: `Completion notification: ${assignment.title}`,
      detail: completedBy?.name || completedBy?.username || "Employee",
      provider: "lark",
      recipients: [workCompletionLarkReceiveId()].filter(Boolean),
      relatedId: assignment.id,
      error: message
    });
  }

  return summary;
}

async function trySendWorkCompletionEmailNotification(workState, assignment, completedBy) {
  const summary = {
    attempted: completedBy?.role === "employee" ? 1 : 0,
    sent: 0,
    skipped: 0,
    failed: 0,
    recipients: [],
    errors: [],
    missing: []
  };

  if (!summary.attempted) {
    return summary;
  }

  try {
    const result = await sendWorkCompletionEmail(assignment, workState, completedBy);
    const now = new Date().toISOString();

    if (result.sent) {
      summary.sent = 1;
      summary.recipients = result.recipients;
      assignment.completionEmailStatus = "sent";
      assignment.completionEmailSentAt = now;
      assignment.completionEmailTo = result.recipients.join(", ");
      assignment.completionEmailFrom = workCompletionEmailConfig.from;
      assignment.completionEmailError = "";
      await appendSendLog({
        type: "work_email",
        status: "success",
        title: `Completion email: ${assignment.title}`,
        detail: completedBy?.name || completedBy?.username || "Employee",
        provider: "resend",
        from: workCompletionEmailConfig.from,
        recipients: result.recipients,
        relatedId: assignment.id
      });
    } else {
      summary.skipped = 1;
      summary.missing = result.missing || [];
      assignment.completionEmailStatus = "not_configured";
      assignment.completionEmailFrom = workCompletionEmailConfig.from;
      assignment.completionEmailError = result.missing?.length
        ? `Missing ${result.missing.join(", ")}`
        : "Completion email is not configured.";
      await appendSendLog({
        type: "work_email",
        status: "skipped",
        title: `Completion email: ${assignment.title}`,
        detail: completedBy?.name || completedBy?.username || "Employee",
        provider: "resend",
        from: workCompletionEmailConfig.from,
        recipients: workCompletionEmailRecipients(),
        relatedId: assignment.id,
        error: assignment.completionEmailError
      });
    }
  } catch (error) {
    summary.failed = 1;
    const message = error.message || "Could not send completion email.";
    summary.errors = [message];
    assignment.completionEmailStatus = "failed";
    assignment.completionEmailFrom = workCompletionEmailConfig.from;
    assignment.completionEmailTo = workCompletionEmailRecipients().join(", ");
    assignment.completionEmailError = message;
    await appendSendLog({
      type: "work_email",
      status: "failed",
      title: `Completion email: ${assignment.title}`,
      detail: completedBy?.name || completedBy?.username || "Employee",
      provider: "resend",
      from: workCompletionEmailConfig.from,
      recipients: workCompletionEmailRecipients(),
      relatedId: assignment.id,
      error: message
    });
  }

  return summary;
}

function workCompletionNotificationMessage(larkSummary, emailSummary = null) {
  const messages = [];

  if (emailSummary?.attempted) {
    if (emailSummary.sent) {
      messages.push(`Completion email sent to ${emailSummary.recipients.join(", ")}.`);
    } else if (emailSummary.failed) {
      messages.push(`Finished, but Barry's completion email could not send: ${emailSummary.errors[0] || "unknown error"}`);
    } else if (emailSummary.skipped) {
      messages.push("Finished. Add RESEND_API_KEY in Render to email Barry when work is finished.");
    }
  }

  if (larkSummary?.attempted) {
    if (larkSummary.sent) {
      messages.push("Lark completion notification sent to Barry.");
    } else if (larkSummary.failed) {
      messages.push(`Lark completion notification could not send: ${larkSummary.errors[0] || "unknown error"}`);
    } else if (larkSummary.skipped) {
      messages.push("Add WORK_COMPLETION_LARK_RECEIVE_ID or BOSS_LARK_RECEIVE_ID in Render to notify Barry in Lark.");
    }
  }

  return messages.join(" ");
}

function workInviteMessage(summary) {
  if (!summary?.attempted) return "";
  if (summary.sent) {
    return `Work invite sent from ${workInviteEmailConfig.from} to ${summary.recipients.join(", ")}.`;
  }
  if (summary.failed) {
    return `Work saved, but the invite email could not send: ${summary.errors[0] || "unknown error"}`;
  }
  if (summary.skipped) {
    const missing = summary.missing || [];
    if (missing.includes("RESEND_API_KEY")) {
      return "Work saved. Email invite could not send because RESEND_API_KEY is missing in Render.";
    }
    if (missing.includes("WORK_INVITE_EMAIL_ENABLED")) {
      return "Work saved. Email invite sending is turned off.";
    }

    return `Work saved. Add the employee's email details to send work invite emails from ${workInviteEmailConfig.from}.`;
  }
  return "";
}

function calendarInviteEmailSetupMissingSettings() {
  const missing = [];
  if (!calendarInviteEmailConfig.enabled) missing.push("CALENDAR_INVITE_EMAIL_ENABLED");
  if (!resendConfig.apiKey) missing.push("RESEND_API_KEY");
  if (!calendarInviteEmailConfig.from) missing.push("CALENDAR_INVITE_EMAIL_FROM");
  return missing;
}

function isCalendarInviteEmailConfigured() {
  return calendarInviteEmailSetupMissingSettings().length === 0;
}

function calendarInviteRecipients(booking, recipientsOverride = null) {
  const source = recipientsOverride || booking?.guestEmails || [];
  const emails = Array.isArray(source) ? source : parseGuestEmails(source);
  return uniqueEmails(emails).filter(isValidEmail);
}

function calendarInviteEmailMissingSettings(booking, recipientsOverride = null) {
  const missing = calendarInviteEmailSetupMissingSettings();
  if (!calendarInviteRecipients(booking, recipientsOverride).length) {
    missing.push("GUEST_EMAILS");
  }
  return missing;
}

function shouldUseLarkCalendarNotifications() {
  const explicitValue = process.env.LARK_CALENDAR_NOTIFICATIONS_ENABLED;
  if (explicitValue !== undefined && explicitValue !== "") {
    return normalizeEnvBoolean(explicitValue, false);
  }

  return !isCalendarInviteEmailConfigured();
}

function shouldAddLarkAttendees() {
  const explicitValue = process.env.LARK_ATTENDEES_ENABLED;
  if (explicitValue !== undefined && explicitValue !== "") {
    return normalizeEnvBoolean(explicitValue, false);
  }

  return !isCalendarInviteEmailConfigured();
}

function emailAddressFromSender(sender) {
  const value = String(sender || "").trim();
  const match = value.match(/<([^>]+)>/);
  const email = (match ? match[1] : value).trim();
  return isValidEmail(email) ? email : invoiceEmailUser;
}

function displayNameFromSender(sender) {
  const value = String(sender || "").replace(/<[^>]+>/, "").replace(/^"|"$/g, "").trim();
  return value || "OpenFrame Studio";
}

function escapeIcsText(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/;/g, "\\;")
    .replace(/,/g, "\\,")
    .replace(/\r?\n/g, "\\n");
}

function escapeIcsParameter(value) {
  return String(value || "")
    .replace(/\\/g, "\\\\")
    .replace(/"/g, "\\\"")
    .replace(/\r?\n/g, " ");
}

function foldIcsLine(line) {
  const lines = [];
  let value = String(line);

  while (Buffer.byteLength(value, "utf8") > 73) {
    let take = Math.min(73, value.length);
    while (take > 1 && Buffer.byteLength(value.slice(0, take), "utf8") > 73) {
      take -= 1;
    }
    lines.push(value.slice(0, take));
    value = ` ${value.slice(take)}`;
  }

  lines.push(value);
  return lines.join("\r\n");
}

function formatIcsDateTime(value) {
  const date = value instanceof Date ? value : new Date(value);
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function formatCalendarInviteDateTime(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "Time not set";

  return new Intl.DateTimeFormat("en-AU", {
    weekday: "short",
    day: "numeric",
    month: "short",
    year: "numeric",
    hour: "numeric",
    minute: "2-digit",
    timeZone: larkConfig.timezone
  }).format(date);
}

function formatCalendarInviteWindow(booking) {
  return `${formatCalendarInviteDateTime(booking.startAt)} - ${formatCalendarInviteDateTime(booking.endAt)} (${larkConfig.timezone})`;
}

function calendarInviteUid(booking) {
  if (booking?.id) {
    return `${booking.id}@openframe.studio`;
  }

  const stableHash = crypto
    .createHash("sha256")
    .update(`${booking?.propertyAddress || ""}|${booking?.startAt || ""}`)
    .digest("hex")
    .slice(0, 24);
  return `${stableHash}@openframe.studio`;
}

function buildBookingCalendarInviteDescription(booking) {
  return [
    buildLarkDescription(booking),
    "",
    "Created by OpenFrame Studio internal booking system."
  ].join("\n").trim();
}

function buildBookingCalendarInviteIcs(booking, method = "REQUEST", recipients = calendarInviteRecipients(booking)) {
  const methodValue = method === "CANCEL" ? "CANCEL" : "REQUEST";
  const senderEmail = emailAddressFromSender(calendarInviteEmailConfig.from);
  const senderName = displayNameFromSender(calendarInviteEmailConfig.from);
  const status = methodValue === "CANCEL" ? "CANCELLED" : "CONFIRMED";
  const sequence = Math.max(0, Number(booking.calendarInviteSequence || 0));
  const location = booking.locationAddress || booking.locationName || booking.propertyAddress || "";
  const lines = [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "PRODID:-//OpenFrame Studio//Internal Booking//EN",
    "CALSCALE:GREGORIAN",
    `METHOD:${methodValue}`,
    "BEGIN:VEVENT",
    `UID:${calendarInviteUid(booking)}`,
    `DTSTAMP:${formatIcsDateTime(new Date())}`,
    `DTSTART:${formatIcsDateTime(booking.startAt)}`,
    `DTEND:${formatIcsDateTime(booking.endAt)}`,
    `SUMMARY:${escapeIcsText(booking.propertyAddress || "OpenFrame Studio booking")}`,
    `LOCATION:${escapeIcsText(location)}`,
    `DESCRIPTION:${escapeIcsText(buildBookingCalendarInviteDescription(booking))}`,
    `ORGANIZER;CN="${escapeIcsParameter(senderName)}":mailto:${senderEmail}`,
    ...recipients.map((email) =>
      `ATTENDEE;CN="${escapeIcsParameter(email)}";ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE:mailto:${email}`
    ),
    `STATUS:${status}`,
    `SEQUENCE:${sequence}`,
    "TRANSP:OPAQUE",
    "END:VEVENT",
    "END:VCALENDAR"
  ];

  return `${lines.map(foldIcsLine).join("\r\n")}\r\n`;
}

function buildBookingCalendarInviteSubject(booking, method = "REQUEST") {
  const prefix = method === "CANCEL" ? "Booking cancelled" : calendarInviteEmailConfig.subjectPrefix;
  return `${prefix}: ${booking.propertyAddress || "OpenFrame Studio booking"}`;
}

function buildBookingCalendarInviteEmailText(booking, method = "REQUEST") {
  const isCancel = method === "CANCEL";
  const lines = [
    isCancel ? "This booking has been cancelled." : "Your calendar invitation is attached.",
    "",
    `Booking: ${booking.propertyAddress || "OpenFrame Studio booking"}`,
    `Time: ${formatCalendarInviteWindow(booking)}`,
    `Address: ${booking.locationAddress || booking.propertyAddress || ""}`,
    booking.service ? `Services: ${booking.service}` : "",
    booking.agentName || booking.agentPhone ? `Agent: ${formatContact(booking.agentName, booking.agentPhone)}` : "",
    booking.photographerName || booking.photographerPhone ? `Photographer: ${formatContact(booking.photographerName, booking.photographerPhone)}` : "",
    "",
    "OpenFrame Studio"
  ];

  return lines.filter((line) => line !== "").join("\n");
}

function buildBookingCalendarInviteEmailHtml(booking, method = "REQUEST") {
  const isCancel = method === "CANCEL";
  const intro = isCancel ? "This booking has been cancelled." : "Your calendar invitation is attached.";

  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p>${escapeHtmlForEmail(intro)}</p>
      <p><strong>Booking:</strong> ${escapeHtmlForEmail(booking.propertyAddress || "OpenFrame Studio booking")}</p>
      <p><strong>Time:</strong> ${escapeHtmlForEmail(formatCalendarInviteWindow(booking))}</p>
      <p><strong>Address:</strong> ${escapeHtmlForEmail(booking.locationAddress || booking.propertyAddress || "")}</p>
      ${booking.service ? `<p><strong>Services:</strong> ${escapeHtmlForEmail(booking.service)}</p>` : ""}
      ${booking.agentName || booking.agentPhone ? `<p><strong>Agent:</strong> ${escapeHtmlForEmail(formatContact(booking.agentName, booking.agentPhone))}</p>` : ""}
      ${booking.photographerName || booking.photographerPhone ? `<p><strong>Photographer:</strong> ${escapeHtmlForEmail(formatContact(booking.photographerName, booking.photographerPhone))}</p>` : ""}
      <p>OpenFrame Studio</p>
    </div>
  `;
}

async function trySendBookingCalendarInviteEmail(booking, method = "REQUEST", recipientsOverride = null) {
  const recipients = calendarInviteRecipients(booking, recipientsOverride);
  const shouldTrack = !recipientsOverride;
  const now = new Date().toISOString();

  if (shouldTrack) {
    booking.calendarInviteEmailFrom = calendarInviteEmailConfig.from;
    booking.calendarInviteTo = recipients.join(", ");
  }

  if (!recipients.length) {
    if (shouldTrack) {
      booking.calendarInviteStatus = "not_needed";
      booking.calendarInviteError = null;
    }
    return { sent: false, skipped: true, missing: ["GUEST_EMAILS"] };
  }

  const missing = calendarInviteEmailMissingSettings(booking, recipients);
  if (missing.length) {
    if (shouldTrack) {
      booking.calendarInviteStatus = "not_configured";
      booking.calendarInviteError = `Missing ${missing.join(", ")}`;
    }
    await appendSendLog({
      type: "calendar_invite",
      status: "skipped",
      title: `Calendar invite: ${booking.propertyAddress || "Booking"}`,
      detail: method === "CANCEL" ? "Cancellation invite" : "Booking invite",
      provider: "resend",
      from: calendarInviteEmailConfig.from,
      recipients,
      relatedId: booking.id,
      error: `Missing ${missing.join(", ")}`
    });
    return { sent: false, skipped: true, missing };
  }

  booking.calendarInviteSequence = Math.max(0, Number(booking.calendarInviteSequence || 0)) + 1;
  const ics = buildBookingCalendarInviteIcs(booking, method, recipients);
  const methodValue = method === "CANCEL" ? "CANCEL" : "REQUEST";
  const bcc = uniqueEmails(parseGuestEmails(calendarInviteEmailConfig.bcc)).filter(isValidEmail);
  const payload = {
    from: calendarInviteEmailConfig.from,
    to: recipients,
    subject: buildBookingCalendarInviteSubject(booking, methodValue),
    text: buildBookingCalendarInviteEmailText(booking, methodValue),
    html: buildBookingCalendarInviteEmailHtml(booking, methodValue),
    reply_to: calendarInviteEmailConfig.replyTo,
    attachments: [
      {
        filename: `openframe-booking-${methodValue.toLowerCase()}-${booking.id || "invite"}.ics`,
        content: Buffer.from(ics, "utf8").toString("base64")
      }
    ],
    headers: {
      "Content-Class": "urn:content-classes:calendarmessage"
    }
  };

  if (bcc.length) {
    payload.bcc = bcc;
  }

  try {
    await sendResendEmail(payload, calendarInviteEmailConfig.timeoutMs, "Calendar invite email request timed out.");
    if (shouldTrack) {
      booking.calendarInviteStatus = methodValue === "CANCEL" ? "cancelled" : "sent";
      booking.calendarInviteSentAt = now;
      booking.calendarInviteTo = recipients.join(", ");
      booking.calendarInviteEmailFrom = calendarInviteEmailConfig.from;
      booking.calendarInviteError = null;
    }
    await appendSendLog({
      type: "calendar_invite",
      status: "success",
      title: `Calendar invite: ${booking.propertyAddress || "Booking"}`,
      detail: `${methodValue} ${formatCalendarInviteWindow(booking)}`,
      provider: "resend",
      from: calendarInviteEmailConfig.from,
      recipients,
      relatedId: booking.id
    });
    return { sent: true, recipients };
  } catch (error) {
    const message = error.message || "Could not send calendar invite email.";
    if (shouldTrack) {
      booking.calendarInviteStatus = "failed";
      booking.calendarInviteEmailFrom = calendarInviteEmailConfig.from;
      booking.calendarInviteError = message;
    }
    await appendSendLog({
      type: "calendar_invite",
      status: "failed",
      title: `Calendar invite: ${booking.propertyAddress || "Booking"}`,
      detail: `${methodValue} ${formatCalendarInviteWindow(booking)}`,
      provider: "resend",
      from: calendarInviteEmailConfig.from,
      recipients,
      relatedId: booking.id,
      error: message
    });
    return { sent: false, failed: true, error: message };
  }
}

function isLarkConfigured() {
  return Boolean(larkConfig.appId && larkConfig.appSecret && effectiveLarkCalendarId());
}

function effectiveLarkCalendarId() {
  return larkConfig.organizerCalendarId || larkConfig.calendarId;
}

function buildLarkOrganizerPayload() {
  const organizer = {};

  if (larkConfig.organizerUserId) {
    organizer.user_id = larkConfig.organizerUserId;
  }

  if (larkConfig.senderName || larkConfig.senderEmail) {
    organizer.display_name = larkConfig.senderName || larkConfig.senderEmail;
  }

  return Object.keys(organizer).length ? organizer : null;
}

function getEndAt(startAt, durationMinutes) {
  return new Date(new Date(startAt).getTime() + durationMinutes * 60_000).toISOString();
}

function hasBookingConflict(bookings, startAt, endAt, ignoredBookingId = "") {
  const nextStart = new Date(startAt).getTime();
  const nextEnd = new Date(endAt).getTime();

  return bookings.some((booking) => {
    if (booking.id === ignoredBookingId) {
      return false;
    }

    if (booking.status === "cancelled") {
      return false;
    }

    const existingStart = new Date(booking.startAt).getTime();
    const existingEnd = new Date(booking.endAt).getTime();
    return nextStart < existingEnd && nextEnd > existingStart;
  });
}

function parseGuestEmails(input) {
  const emailMatches = String(input || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  if (emailMatches?.length) {
    return emailMatches.map((email) => email.trim());
  }

  return String(input || "")
    .split(/[\s,;]+/)
    .map((email) => email.trim())
    .filter(Boolean);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function uniqueEmails(emails) {
  const seen = new Set();
  const unique = [];

  for (const email of emails) {
    const trimmed = String(email || "").trim();
    if (!trimmed) {
      continue;
    }

    const key = trimmed.toLowerCase();
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    unique.push(trimmed);
  }

  return unique;
}

function formatContact(name, phone) {
  if (name && phone) {
    return `${name} (${phone})`;
  }

  return name || phone || "";
}

function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function normalizeAddress(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function compactAddress(result) {
  const address = result.address || {};
  const parts = [
    [address.house_number, address.road].filter(Boolean).join(" "),
    address.suburb || address.city_district || address.town || address.city || address.village,
    address.state,
    address.postcode,
    address.country
  ].filter(Boolean);

  return normalizeAddress(parts.join(" "));
}

function uniqueAddressSuggestions(items) {
  const seen = new Set();
  const suggestions = [];

  for (const item of items) {
    const address = normalizeAddress(item.address || item.label || "");
    if (!address) {
      continue;
    }

    const key = address.toLowerCase();
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    suggestions.push({
      address,
      label: normalizeAddress(item.label || address)
    });
  }

  return suggestions;
}

async function fetchJsonWithTimeout(url, options = {}, timeoutMs = 6500) {
  const controller = new AbortController();
  const timeout = setTimeout(() => controller.abort(), timeoutMs);
  try {
    const response = await fetch(url, { ...options, signal: controller.signal });
    if (!response.ok) {
      throw new Error(`Address lookup failed with ${response.status}.`);
    }
    return await response.json().catch(() => ({}));
  } finally {
    clearTimeout(timeout);
  }
}

function normalizeArcgisAddress(value) {
  return normalizeAddress(String(value || "").replace(/,\s*AUS$/i, ", Australia"));
}

async function findArcgisAddressSuggestions(cleanQuery) {
  const params = new URLSearchParams({
    f: "json",
    text: cleanQuery,
    countryCode: "AUS",
    maxSuggestions: "6"
  });

  const result = await fetchJsonWithTimeout(`https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/suggest?${params}`);
  return uniqueAddressSuggestions((result.suggestions || []).map((suggestion) => ({
    address: normalizeArcgisAddress(suggestion.text),
    label: suggestion.text
  })));
}

async function findNominatimAddressSuggestions(cleanQuery) {
  const waitMs = Math.max(0, 1100 - (Date.now() - lastAddressLookupAt));
  if (waitMs) {
    await delay(waitMs);
  }
  lastAddressLookupAt = Date.now();

  const params = new URLSearchParams({
    format: "jsonv2",
    addressdetails: "1",
    countrycodes: "au",
    limit: "6",
    q: cleanQuery
  });

  const results = await fetchJsonWithTimeout(`https://nominatim.openstreetmap.org/search?${params}`, {
    headers: {
      "User-Agent": "OpenFrameInternalBooking/1.0 (internalbooking.openframe.studio)",
      "Accept-Language": "en-AU,en;q=0.9"
    }
  });

  return uniqueAddressSuggestions((Array.isArray(results) ? results : []).map((result) => ({
    address: compactAddress(result) || result.display_name,
    label: result.display_name
  })));
}

async function findAddressSuggestions(query) {
  const cleanQuery = normalizeAddress(query);
  if (cleanQuery.length < 4) {
    return [];
  }

  const cacheKey = cleanQuery.toLowerCase();
  const cached = addressSuggestionCache.get(cacheKey);
  if (cached && cached.expiresAt > Date.now()) {
    return cached.suggestions;
  }

  let suggestions = [];
  try {
    suggestions = await findArcgisAddressSuggestions(cleanQuery);
  } catch {
    suggestions = [];
  }

  if (!suggestions.length) {
    try {
      suggestions = await findNominatimAddressSuggestions(cleanQuery);
    } catch {
      suggestions = [];
    }
  }

  addressSuggestionCache.set(cacheKey, {
    suggestions,
    expiresAt: Date.now() + 24 * 60 * 60 * 1000
  });

  return suggestions;
}

function validateClient(input) {
  const errors = [];
  const addressLine1 = String(input.addressLine1 || input.billingAddressLine1 || "").trim();
  const addressLine2 = String(input.addressLine2 || input.billingAddressLine2 || "").trim();
  const city = String(input.city || input.billingCity || "").trim();
  const postcode = String(input.postcode || input.postalCode || input.billingPostcode || "").trim();
  const client = {
    id: String(input.id || "").trim(),
    name: String(input.name || input.clientName || "").trim(),
    email: String(input.email || input.clientEmail || "").trim(),
    agentName: String(input.agentName || "").trim(),
    agentPhone: String(input.agentPhone || input.phone || "").trim(),
    addressLine1,
    addressLine2,
    city,
    postcode,
    billingAddress: String(input.billingAddress || [addressLine1, addressLine2, city, postcode].filter(Boolean).join(", ")).trim(),
    abn: String(input.abn || input.ABN || "").trim(),
    customerStatus: String(input.customerStatus || input.status || "").trim(),
    sourceCustomerId: String(input.sourceCustomerId || input.customerId || "").trim()
  };
  const clientEmails = uniqueEmails(parseGuestEmails(client.email));

  if (!client.name) {
    errors.push("Enter the agency or client name.");
  }

  for (const email of clientEmails) {
    if (!isValidEmail(email)) {
      errors.push(`Enter a valid client email for ${email}.`);
    }
  }

  client.email = clientEmails.join(", ");
  return { errors, client };
}

function mergeClientRecord(existing, client, now) {
  const merged = {
    ...(existing || {}),
    name: client.name,
    email: client.email,
    agentName: client.agentName,
    agentPhone: client.agentPhone,
    updatedAt: now
  };
  const optionalFields = [
    "addressLine1",
    "addressLine2",
    "city",
    "postcode",
    "billingAddress",
    "abn",
    "customerStatus",
    "sourceCustomerId"
  ];

  for (const field of optionalFields) {
    if (client[field] || existing?.[field] === undefined) {
      merged[field] = client[field];
    }
  }

  return merged;
}

function clientIdentityKey(client) {
  return [
    client.name,
    client.agentName,
    client.email,
    client.agentPhone
  ]
    .map((value) => String(value || "").trim().toLowerCase())
    .join("|");
}

function validatePhotographer(input) {
  const errors = [];
  const photographer = {
    id: String(input.id || "").trim(),
    name: String(input.name || input.photographerName || "").trim(),
    email: String(input.email || input.photographerEmail || "").trim(),
    phone: String(input.phone || input.photographerPhone || "").trim(),
    gstIncluded: normalizeEnvBoolean(input.gstIncluded ?? input.photographerGstIncluded, false)
  };

  if (!photographer.name) {
    errors.push("Enter the photographer name.");
  }

  if (photographer.email && !isValidEmail(photographer.email)) {
    errors.push("Enter a valid photographer email address.");
  }

  return { errors, photographer };
}

function validateBooking(input) {
  const errors = [];
  const requestedServices = Array.isArray(input.services) ? input.services : [input.service].filter(Boolean);
  const services = [];
  const seenServices = new Set();
  const propertyAddress = String(input.propertyAddress || "").trim();
  const locationName = propertyAddress;
  const locationAddress = propertyAddress;
  const clientName = String(input.clientName || "").trim();
  const clientEmail = String(input.clientEmail || "").trim();
  const clientEmails = uniqueEmails(parseGuestEmails(clientEmail));
  const photographerName = String(input.photographerName || "Barry").trim();
  const photographerEmail = String(input.photographerEmail || "").trim();
  const photographerPhone = String(input.photographerPhone || "").trim();
  const photographerGstIncluded = normalizeEnvBoolean(input.photographerGstIncluded ?? input.gstIncluded, false);
  const agentName = String(input.agentName || "").trim();
  const agentPhone = String(input.agentPhone || "").trim();
  const rawGuestEmails = parseGuestEmails(input.guestEmails);
  const guestEmails = uniqueEmails([...clientEmails, photographerEmail, ...rawGuestEmails]);
  const notes = String(input.notes || "").trim();
  const bookingDate = String(input.date || input.bookingDate || "").trim();
  const bookingTime = String(input.time || input.bookingTime || "").trim();
  const startAt = bookingDate && bookingTime
    ? zonedDateTimeToDate(bookingDate, bookingTime, bookingTimeZone)?.toISOString() || ""
    : String(input.startAt || "").trim();
  const durationMinutes = Number(input.durationMinutes || 0);

  for (const requestedService of requestedServices) {
    const name = String(requestedService?.name || requestedService || "").trim();
    if (!name || seenServices.has(name)) {
      continue;
    }

    if (!serviceCatalog.has(name)) {
      errors.push(`${name || "Selected service"} is not available.`);
      continue;
    }

    seenServices.add(name);
    services.push({ name });
  }

  if (!services.length) errors.push("Choose at least one service.");
  if (!propertyAddress) errors.push("Enter the calendar title or property address.");
  if (!clientName) errors.push("Enter the agency or client name.");
  for (const email of clientEmails) {
    if (!isValidEmail(email)) {
      errors.push(`Enter a valid client email for ${email}.`);
    }
  }
  if (!photographerName) errors.push("Enter the photographer name.");
  if (photographerEmail && !isValidEmail(photographerEmail)) errors.push("Enter a valid photographer email address.");
  for (const email of rawGuestEmails) {
    if (!isValidEmail(email)) {
      errors.push(`Enter a valid guest email for ${email}.`);
    }
  }
  if (!Number.isInteger(durationMinutes) || durationMinutes < 15 || durationMinutes > 240) errors.push("Choose a valid duration.");

  const startDate = new Date(startAt);
  if (!startAt || Number.isNaN(startDate.getTime())) {
    errors.push("Choose a valid start time.");
  }

  return {
    errors,
    value: {
      service: services.map((serviceItem) => serviceItem.name).join(" + "),
      services,
      propertyAddress,
      locationName,
      locationAddress,
      clientName,
      clientEmail: clientEmails.join(", "),
      photographerName,
      photographerEmail,
      photographerPhone,
      photographerGstIncluded,
      agentName,
      agentPhone,
      guestEmails,
      notes,
      startAt: startDate.toISOString(),
      durationMinutes
    }
  };
}

async function getTenantAccessToken() {
  if (!isLarkConfigured()) {
    throw new Error("Lark is not configured.");
  }

  const now = Date.now();
  if (tenantTokenCache && tenantTokenCache.expiresAt > now + 60_000) {
    return tenantTokenCache.token;
  }

  const response = await fetch(`${larkConfig.apiBase}/auth/v3/tenant_access_token/internal`, {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify({
      app_id: larkConfig.appId,
      app_secret: larkConfig.appSecret
    })
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.code !== 0 || !result.tenant_access_token) {
    const message = result.msg || result.message || `Lark token request failed with ${response.status}.`;
    throw new Error(message);
  }

  tenantTokenCache = {
    token: result.tenant_access_token,
    expiresAt: now + Number(result.expire || 7200) * 1000
  };

  return tenantTokenCache.token;
}

function buildLarkDescription(booking) {
  const lines = [booking.clientName];

  if (booking.photographerName || booking.photographerPhone) {
    lines.push(`Photographer: ${formatContact(booking.photographerName, booking.photographerPhone)}`);
  }

  if (booking.agentName || booking.agentPhone) {
    lines.push(`Agent: ${formatContact(booking.agentName, booking.agentPhone)}`);
  }

  if (booking.service) {
    lines.push(`Services: ${booking.service}`);
  }

  if (booking.notes) {
    lines.push("", booking.notes);
  }

  return lines.join("\n");
}

function buildLarkEventPayload(booking) {
  const startTime = Math.floor(new Date(booking.startAt).getTime() / 1000).toString();
  const endTime = Math.floor(new Date(booking.endAt).getTime() / 1000).toString();
  const organizer = buildLarkOrganizerPayload();
  const calendarId = effectiveLarkCalendarId();

  const payload = {
    summary: booking.propertyAddress,
    description: buildLarkDescription(booking),
    need_notification: shouldUseLarkCalendarNotifications(),
    start_time: {
      timestamp: startTime,
      timezone: larkConfig.timezone
    },
    end_time: {
      timestamp: endTime,
      timezone: larkConfig.timezone
    },
    visibility: "default",
    attendee_ability: "can_see_others",
    free_busy_status: "busy",
    vchat: {
      vc_type: "no_meeting"
    },
    location: {
      name: booking.locationName || booking.propertyAddress,
      address: booking.locationAddress || booking.propertyAddress
    },
    reminders: [{ minutes: 30 }]
  };

  if (calendarId && calendarId !== "primary") {
    payload.organizer_calendar_id = calendarId;
  }

  if (organizer) {
    payload.event_organizer = organizer;
  }

  return payload;
}

function buildLarkAttendeesPayload(booking, emails = booking.guestEmails || []) {
  if (!emails.length) {
    return null;
  }

  return {
    need_notification: shouldUseLarkCalendarNotifications(),
    attendees: emails.map((email) => ({
      type: "third_party",
      third_party_email: email
    }))
  };
}

function buildLarkPreview(booking) {
  const eventPayload = buildLarkEventPayload(booking);
  const attendeesPayload = shouldAddLarkAttendees() ? buildLarkAttendeesPayload(booking) : null;

  return {
    title: eventPayload.summary,
    startAt: booking.startAt,
    endAt: booking.endAt,
    timezone: larkConfig.timezone,
    location: eventPayload.location,
    description: eventPayload.description,
    senderEmail: larkConfig.senderEmail,
    senderName: larkConfig.senderName,
    calendarInviteEmailFrom: calendarInviteEmailConfig.from,
    calendarInviteEmailConfigured: isCalendarInviteEmailConfigured(),
    larkNotificationsEnabled: shouldUseLarkCalendarNotifications(),
    larkAttendeesEnabled: shouldAddLarkAttendees(),
    calendarId: effectiveLarkCalendarId(),
    guestEmails: booking.guestEmails,
    eventPayload,
    attendeesPayload
  };
}

function parseLarkDateTime(time) {
  if (!time) {
    return null;
  }

  if (time.timestamp) {
    const timestamp = Number(time.timestamp);
    if (!Number.isNaN(timestamp)) {
      return new Date(timestamp * 1000).toISOString();
    }
  }

  if (time.date) {
    const parsed = zonedDateTimeToDate(time.date, "00:00", bookingTimeZone);
    if (parsed && !Number.isNaN(parsed.getTime())) {
      return parsed.toISOString();
    }
  }

  return null;
}

function readDescriptionLine(description, label) {
  const prefix = `${label}:`;
  const line = String(description || "")
    .split(/\r?\n/)
    .find((item) => item.trim().toLowerCase().startsWith(prefix.toLowerCase()));
  return line ? line.trim().slice(prefix.length).trim() : "";
}

function readDescriptionClient(description) {
  return String(description || "")
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line && !/^(photographer|agent|services|booking id):/i.test(line)) || "";
}

function splitContact(value) {
  const match = String(value || "").match(/^(.*?)\s*\((.*?)\)\s*$/);
  if (!match) {
    return { name: String(value || "").trim(), phone: "" };
  }

  return {
    name: match[1].trim(),
    phone: match[2].trim()
  };
}

function readDescriptionBookingId(description) {
  return readDescriptionLine(description, "Booking ID");
}

function normalizeLarkEvent(event) {
  const larkEventId = event.event_id || event.id || event.uid || "";
  const startAt = parseLarkDateTime(event.start_time || event.start);
  const endAt = parseLarkDateTime(event.end_time || event.end);
  if (!larkEventId || !startAt || !endAt) {
    return null;
  }

  const description = String(event.description || "").trim();
  const photographer = splitContact(readDescriptionLine(description, "Photographer"));
  const agent = splitContact(readDescriptionLine(description, "Agent"));
  const service = readDescriptionLine(description, "Services");
  const services = service
    ? service.split("+").map((name) => ({ name: name.trim() })).filter((item) => item.name)
    : [];
  const location = event.location || {};
  const propertyAddress = normalizeAddress(event.summary || location.address || location.name || "Lark event");
  const startMs = new Date(startAt).getTime();
  const endMs = new Date(endAt).getTime();
  const status = event.status === "cancelled" ? "cancelled" : "confirmed";

  return {
    id: `lark-${larkEventId}`,
    larkOnly: true,
    service,
    services,
    propertyAddress,
    locationName: normalizeAddress(location.name || propertyAddress),
    locationAddress: normalizeAddress(location.address || location.name || propertyAddress),
    clientName: readDescriptionClient(description),
    clientEmail: "",
    photographerName: photographer.name,
    photographerEmail: "",
    photographerPhone: photographer.phone,
    agentName: agent.name,
    agentPhone: agent.phone,
    guestEmails: [],
    notes: description,
    startAt,
    endAt,
    durationMinutes: Math.max(15, Math.round((endMs - startMs) / 60000)),
    status,
    larkStatus: "synced",
    larkEventId,
    larkError: null,
    larkAttendeeStatus: "synced",
    larkAttendeeError: null,
    createdAt: startAt,
    importedFromLark: true,
    sourceBookingId: readDescriptionBookingId(description)
  };
}

function getLarkEventItems(data) {
  if (Array.isArray(data?.items)) return data.items;
  if (Array.isArray(data?.events)) return data.events;
  if (Array.isArray(data?.event_infos)) return data.event_infos;
  if (Array.isArray(data?.event_list)) return data.event_list;
  return [];
}

async function fetchLarkImportedBookings() {
  if (!isLarkConfigured()) {
    return [];
  }

  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(effectiveLarkCalendarId());
  const imported = [];
  const firstStartMs = Date.now() - 24 * 60 * 60 * 1000;
  const finalEndMs = Date.now() + larkImportDays * 24 * 60 * 60 * 1000;
  const windowMs = 35 * 24 * 60 * 60 * 1000;

  for (let windowStartMs = firstStartMs; windowStartMs < finalEndMs && imported.length < 500; windowStartMs += windowMs) {
    const windowEndMs = Math.min(windowStartMs + windowMs, finalEndMs);
    let pageToken = "";

    do {
      const params = new URLSearchParams({
        start_time: Math.floor(windowStartMs / 1000).toString(),
        end_time: Math.floor(windowEndMs / 1000).toString(),
        page_size: "100"
      });
      if (pageToken) {
        params.set("page_token", pageToken);
      }

      const response = await fetch(`${larkConfig.apiBase}/calendar/v4/calendars/${calendarId}/events?${params}`, {
        headers: {
          Authorization: `Bearer ${token}`,
          "Content-Type": "application/json; charset=utf-8"
        }
      });

      const result = await response.json().catch(() => ({}));
      if (!response.ok || result.code !== 0) {
        const message = result.msg || result.message || `Lark event import failed with ${response.status}.`;
        throw new Error(message);
      }

      imported.push(...getLarkEventItems(result.data).map(normalizeLarkEvent).filter(Boolean));
      pageToken = result.data?.page_token || result.data?.next_page_token || "";
    } while (pageToken && imported.length < 500);
  }

  return [...new Map(imported.map((booking) => [booking.larkEventId, booking])).values()].filter((booking) => booking.status !== "cancelled");
}

function mergeLocalAndLarkBookings(localBookings, larkBookings) {
  const localIds = new Set(localBookings.map((booking) => booking.id));
  const localLarkIds = new Set(localBookings.map((booking) => booking.larkEventId).filter(Boolean));

  return [
    ...localBookings,
    ...larkBookings.filter((booking) => {
      return !localLarkIds.has(booking.larkEventId) && !localIds.has(booking.sourceBookingId);
    })
  ];
}

async function createLarkEvent(booking) {
  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(effectiveLarkCalendarId());

  const response = await fetch(`${larkConfig.apiBase}/calendar/v4/calendars/${calendarId}/events`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(buildLarkEventPayload(booking))
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.code !== 0) {
    const message = result.msg || result.message || `Lark event request failed with ${response.status}.`;
    throw new Error(message);
  }

  return result.data?.event || result.data || result;
}

async function updateLarkEvent(booking) {
  if (!booking.larkEventId) {
    throw new Error("No Lark event ID is saved for this booking.");
  }

  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(effectiveLarkCalendarId());
  const eventId = encodeURIComponent(booking.larkEventId);

  const response = await fetch(`${larkConfig.apiBase}/calendar/v4/calendars/${calendarId}/events/${eventId}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(buildLarkEventPayload(booking))
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.code !== 0) {
    const message = result.msg || result.message || `Lark event update failed with ${response.status}.`;
    throw new Error(message);
  }

  return result.data?.event || result.data || result;
}

async function deleteLarkEvent(booking) {
  if (!booking.larkEventId) {
    return null;
  }

  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(effectiveLarkCalendarId());
  const eventId = encodeURIComponent(booking.larkEventId);
  const params = new URLSearchParams({ need_notification: String(shouldUseLarkCalendarNotifications()) });

  const response = await fetch(`${larkConfig.apiBase}/calendar/v4/calendars/${calendarId}/events/${eventId}?${params}`, {
    method: "DELETE",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    }
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || (result.code !== undefined && result.code !== 0)) {
    const message = result.msg || result.message || `Lark event delete failed with ${response.status}.`;
    throw new Error(message);
  }

  return result.data || result;
}

function getAddedGuestEmails(previousEmails = [], nextEmails = []) {
  const previous = new Set(previousEmails.map((email) => email.toLowerCase()));
  return nextEmails.filter((email) => !previous.has(email.toLowerCase()));
}

function getRemovedGuestEmails(previousEmails = [], nextEmails = []) {
  const next = new Set(nextEmails.map((email) => email.toLowerCase()));
  return previousEmails.filter((email) => !next.has(email.toLowerCase()));
}

async function createLarkAttendees(booking, emails = booking.guestEmails || []) {
  if (!shouldAddLarkAttendees()) {
    return null;
  }

  if (!booking.larkEventId || !emails.length) {
    return null;
  }

  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(effectiveLarkCalendarId());
  const eventId = encodeURIComponent(booking.larkEventId);

  const response = await fetch(`${larkConfig.apiBase}/calendar/v4/calendars/${calendarId}/events/${eventId}/attendees`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json; charset=utf-8"
    },
    body: JSON.stringify(buildLarkAttendeesPayload(booking, emails))
  });

  const result = await response.json().catch(() => ({}));
  if (!response.ok || result.code !== 0) {
    const message = result.msg || result.message || `Lark attendee request failed with ${response.status}.`;
    throw new Error(message);
  }

  return result.data?.attendees || result.data || result;
}

async function syncBookingToLark(booking, previousBooking = null) {
  if (!isLarkConfigured()) {
    booking.larkStatus = "not_configured";
    booking.larkError = null;
    booking.larkAttendeeStatus = "not_configured";
    booking.larkAttendeeError = null;
    booking.calendarInviteStatus = booking.guestEmails?.length ? "not_sent" : "not_needed";
    booking.calendarInviteError = booking.guestEmails?.length ? "Lark is not configured, so the calendar invite email was not sent." : null;
    return booking;
  }

  const previousGuestEmails = previousBooking?.guestEmails || [];
  const hasExistingLarkEvent = Boolean(previousBooking?.larkEventId);
  const addedGuestEmails = hasExistingLarkEvent ? getAddedGuestEmails(previousGuestEmails, booking.guestEmails) : booking.guestEmails;
  const removedGuestEmails = hasExistingLarkEvent ? getRemovedGuestEmails(previousGuestEmails, booking.guestEmails) : [];

  try {
    if (booking.larkEventId) {
      await updateLarkEvent(booking);
    } else {
      const larkEvent = await createLarkEvent(booking);
      booking.larkEventId = larkEvent.event_id || larkEvent.id || null;
    }

    booking.larkStatus = "synced";
    booking.larkError = null;

    if (!shouldAddLarkAttendees()) {
      booking.larkAttendeeStatus = booking.guestEmails.length ? "handled_by_admin_email" : "not_needed";
      booking.larkAttendeeError = null;
    } else if (addedGuestEmails.length) {
      try {
        await createLarkAttendees(booking, addedGuestEmails);
        booking.larkAttendeeStatus = "synced";
        booking.larkAttendeeError = null;
      } catch (error) {
        booking.larkStatus = "attendees_failed";
        booking.larkAttendeeStatus = "failed";
        booking.larkAttendeeError = error.message;
      }
    } else if (!booking.guestEmails.length) {
      booking.larkAttendeeStatus = "not_needed";
      booking.larkAttendeeError = null;
    } else if (previousBooking) {
      booking.larkAttendeeStatus = previousBooking.larkAttendeeStatus || "synced";
      booking.larkAttendeeError = previousBooking.larkAttendeeError || null;
    } else {
      booking.larkAttendeeStatus = "not_needed";
      booking.larkAttendeeError = null;
    }

    if (removedGuestEmails.length) {
      if (shouldAddLarkAttendees()) {
        booking.larkAttendeeStatus = "needs_review";
        booking.larkAttendeeError = "Guest list changed. Remove old guests from Lark manually if needed.";
      }
      await trySendBookingCalendarInviteEmail(previousBooking || booking, "CANCEL", removedGuestEmails);
    }

    if (booking.larkStatus === "synced" || booking.larkStatus === "attendees_failed") {
      await trySendBookingCalendarInviteEmail(booking, "REQUEST");
    }
  } catch (error) {
    booking.larkStatus = "failed";
    booking.larkError = error.message;
    booking.larkAttendeeStatus = booking.guestEmails.length ? "not_sent" : "not_needed";
    booking.calendarInviteStatus = booking.guestEmails.length ? "not_sent" : "not_needed";
    booking.calendarInviteError = booking.guestEmails.length ? "Lark did not sync, so the calendar invite email was not sent." : null;
  }

  return booking;
}

async function cancelBookingInLark(booking) {
  if (!isLarkConfigured()) {
    booking.larkStatus = "not_configured";
    booking.larkError = null;
    booking.larkAttendeeStatus = "not_configured";
    booking.larkAttendeeError = null;
    return booking;
  }

  if (!booking.larkEventId) {
    booking.larkStatus = "not_needed";
    booking.larkError = null;
    booking.larkAttendeeStatus = "not_needed";
    booking.larkAttendeeError = null;
    return booking;
  }

  try {
    await deleteLarkEvent(booking);
    booking.larkStatus = "cancelled";
    booking.larkError = null;
    booking.larkAttendeeStatus = "cancelled";
    booking.larkAttendeeError = null;
  } catch (error) {
    booking.larkStatus = "delete_failed";
    booking.larkError = error.message;
    booking.larkAttendeeStatus = booking.guestEmails?.length ? "cancel_not_sent" : "not_needed";
    booking.larkAttendeeError = error.message;
  }

  return booking;
}

async function handleApi(req, res, url) {
  if (req.method === "GET" && url.pathname === "/api/session") {
    const user = currentUser(req);
    const headers = user ? { "Set-Cookie": buildSessionCookie(req, createSession(user)) } : {};
    sendJson(res, 200, { authenticated: Boolean(user), user }, headers);
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/login") {
    const input = await readBody(req);
    const user = await findAuthUser(input);
    if (!user) {
      sendJson(res, 401, { errors: ["Login details did not match."] });
      return;
    }

    const token = createSession(user);
    sendJson(res, 200, { ok: true, user: sessionUserPayload(user) }, { "Set-Cookie": buildSessionCookie(req, token) });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/logout") {
    const token = getSessionToken(req);
    if (token) {
      sessions.delete(token);
    }

    sendJson(res, 200, { ok: true }, { "Set-Cookie": clearSessionCookie(req) });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/status") {
    sendJson(res, 200, {
      larkConfigured: isLarkConfigured(),
      calendarId: effectiveLarkCalendarId(),
      senderEmail: larkConfig.senderEmail,
      senderName: larkConfig.senderName,
      organizerCalendarConfigured: Boolean(larkConfig.organizerCalendarId),
      invoiceEmailConfigured: isInvoiceEmailConfigured(),
      invoiceEmailProvider,
      wageEmailConfigured: isInvoiceEmailConfigured(),
      workInviteEmailConfigured: isWorkInviteEmailConfigured(),
      workInviteEmailFrom: workInviteEmailConfig.from,
      workLarkNotificationConfigured: isWorkLarkNotificationConfigured(),
      workLarkNotificationsEnabled: workLarkNotificationConfig.enabled,
      workLarkReceiveIdType: normalizeLarkMessageReceiveIdType(workLarkNotificationConfig.receiveIdType),
      workCompletionLarkNotificationConfigured: isWorkCompletionLarkNotificationConfigured(),
      workCompletionLarkNotificationsEnabled: workCompletionLarkNotificationConfig.enabled,
      workCompletionLarkReceiveIdType: workCompletionLarkReceiveIdType(),
      workCompletionEmailConfigured: isWorkCompletionEmailConfigured(),
      workCompletionEmailTo: workCompletionEmailRecipients().join(", "),
      calendarInviteEmailConfigured: isCalendarInviteEmailConfigured(),
      calendarInviteEmailFrom: calendarInviteEmailConfig.from,
      larkNotificationsEnabled: shouldUseLarkCalendarNotifications(),
      larkAttendeesEnabled: shouldAddLarkAttendees(),
      timezone: larkConfig.timezone,
      persistentStorage: storageBackend !== "app",
      storageBackend,
      supabaseConfigured: storageBackend === "supabase"
    });
    return;
  }

  if (!isAuthenticated(req)) {
    sendJson(res, 401, { errors: ["Log in first."] });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/change-password") {
    const input = await readBody(req);
    const currentPassword = String(input.currentPassword || "");
    const newPassword = String(input.newPassword || "");
    const confirmPassword = String(input.confirmPassword || "");
    const user = currentUser(req);

    if (!currentPassword || !newPassword) {
      sendJson(res, 400, { errors: ["Enter your current password and your new password."] });
      return;
    }

    if (newPassword.length < 4) {
      sendJson(res, 400, { errors: ["Use at least 4 characters for the new password."] });
      return;
    }

    if (confirmPassword !== newPassword) {
      sendJson(res, 400, { errors: ["The new passwords do not match."] });
      return;
    }

    const authUser = await effectiveAuthUser(user.username);
    if (!authUser || !verifyAuthPassword(authUser, currentPassword)) {
      sendJson(res, 400, { errors: ["Current password is not correct."] });
      return;
    }

    await updateAuthUserPassword(user.username, newPassword);
    expireOtherSessionsForUser(user.username, getSessionToken(req));
    sendJson(res, 200, { ok: true, message: "Password updated." });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/change-role") {
    const user = currentUser(req);
    if (!user?.canChangeRole) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const role = normalizeAccountRole(input.role);
    if (!role) {
      sendJson(res, 400, { errors: ["Choose a valid role."] });
      return;
    }

    const baseUser = baseAuthUser(user.username);
    if (!baseUser) {
      sendJson(res, 404, { errors: ["Account was not found."] });
      return;
    }

    const nextUser = userWithAccountRole(baseUser, role);
    const token = createSession(nextUser);
    sendJson(res, 200, { ok: true, user: sessionUserPayload(nextUser) }, { "Set-Cookie": buildSessionCookie(req, token) });
    return;
  }

  if (url.pathname === "/api/send-logs") {
    if (!canViewSendLogs(req)) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    if (req.method === "GET") {
      sendJson(res, 200, { logs: await loadSendLogs() });
      return;
    }

    if (req.method === "DELETE") {
      await saveSendLogs([]);
      sendJson(res, 200, { logs: [], message: "Sending logs cleared." });
      return;
    }
  }

  if (req.method === "POST" && url.pathname === "/api/cleanup/past-bookings-invoices") {
    if (!hasPermission(req, "manage_bookings") || !hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss only.");
      return;
    }

    const result = await deletePastBookingsAndInvoices();
    sendJson(res, 200, result);
    return;
  }

  const financialYearExportMatch = url.pathname.match(/^\/api\/exports\/financial-year\/(\d{4})$/);
  if (req.method === "GET" && financialYearExportMatch) {
    if (!hasPermission(req, "manage_invoices") || !hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss only.");
      return;
    }

    const bounds = australianFinancialYearBounds(financialYearExportMatch[1]);
    if (!bounds) {
      sendJson(res, 400, { errors: ["Choose a valid Australian financial year."] });
      return;
    }

    try {
      const { buffer, counts } = await createFinancialYearExportZip(bounds);
      const filename = financialYearExportFileName(bounds).replace(/["\\\r\n]/g, "");
      res.writeHead(200, {
        "Content-Type": "application/zip",
        "Content-Length": buffer.length,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store",
        "X-OpenFrame-Invoice-Count": String(counts.invoices),
        "X-OpenFrame-Employee-Wage-Count": String(counts.employeeWages),
        "X-OpenFrame-Contractor-Wage-Count": String(counts.contractorWages)
      });
      res.end(buffer);
    } catch (error) {
      sendJson(res, 500, { errors: [error.message || "Could not create financial year export."] });
    }
    return;
  }

  if (url.pathname.startsWith("/api/work") && !canAccessApp(req, "work")) {
    sendForbidden(res);
    return;
  }

  if (url.pathname.startsWith("/api/invoices") && !canAccessApp(req, "invoices")) {
    sendForbidden(res);
    return;
  }

  if (url.pathname.startsWith("/api/wages") && !canAccessApp(req, "wages")) {
    sendForbidden(res);
    return;
  }

  if (url.pathname.startsWith("/api/employee-wages") && !canAccessApp(req, "wages")) {
    sendForbidden(res);
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/work") {
    const workState = await loadWorkState();
    sendJson(res, 200, {
      ...visibleWorkStateForUser(workState, currentUser(req)),
      user: currentUser(req)
    });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/work/sync-bookings") {
    if (!requirePermission(req, res, "sync_work_bookings", "Boss or team leader only.")) {
      return;
    }

    const workState = await loadWorkState();
    const bookings = await loadBookings();
    const result = mergeBookingWorkAssignments(workState, bookings);
    const { createdAssignments, ...syncResult } = result;
    const workInvite = await sendWorkInviteEmails(workState, createdAssignments);
    const workLarkNotification = await sendWorkLarkNotifications(workState, createdAssignments);
    const workNotificationMessage = workAssignmentNotificationMessage(workLarkNotification, workInvite);
    await saveWorkState(workState);
    sendJson(res, 200, {
      ...visibleWorkStateForUser(workState, currentUser(req)),
      ...syncResult,
      workInvite,
      workLarkNotification,
      workLarkNotificationMessage: workLarkNotificationMessage(workLarkNotification),
      workInviteMessage: workNotificationMessage,
      user: currentUser(req)
    });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/work/assignments") {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const input = await readBody(req);
    const workState = await loadWorkState();
    const { errors, assignment } = validateWorkAssignment(input, null, workState);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    workState.assignments.push(assignment);
    const workInvite = await sendWorkInviteEmails(workState, [assignment]);
    const workLarkNotification = await sendWorkLarkNotifications(workState, [assignment]);
    const workNotificationMessage = workAssignmentNotificationMessage(workLarkNotification, workInvite);
    await saveWorkState(workState);
    sendJson(res, 201, {
      assignment,
      ...visibleWorkStateForUser(workState, currentUser(req)),
      workInvite,
      workLarkNotification,
      workLarkNotificationMessage: workLarkNotificationMessage(workLarkNotification),
      workInviteMessage: workNotificationMessage,
      user: currentUser(req)
    });
    return;
  }

  const notifyWorkMatch = url.pathname.match(/^\/api\/work\/assignments\/([^/]+)\/notify$/);
  if (req.method === "POST" && notifyWorkMatch) {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const assignmentId = decodeURIComponent(notifyWorkMatch[1]);
    const workState = await loadWorkState();
    const assignment = workState.assignments.find((item) => item.id === assignmentId);
    if (!assignment) {
      sendJson(res, 404, { errors: ["Work not found."] });
      return;
    }

    const workInvite = await sendWorkInviteEmails(workState, [assignment]);
    const workLarkNotification = await sendWorkLarkNotifications(workState, [assignment]);
    const workNotificationMessage = workAssignmentNotificationMessage(workLarkNotification, workInvite);
    await saveWorkState(workState);
    sendJson(res, 200, {
      ...visibleWorkStateForUser(workState, currentUser(req)),
      workInvite,
      workLarkNotification,
      workLarkNotificationMessage: workLarkNotificationMessage(workLarkNotification, "manual"),
      workInviteMessage: workNotificationMessage,
      user: currentUser(req)
    });
    return;
  }

  const completeWorkMatch = url.pathname.match(/^\/api\/work\/assignments\/([^/]+)\/complete$/);
  if (req.method === "POST" && completeWorkMatch) {
    const assignmentId = decodeURIComponent(completeWorkMatch[1]);
    const workState = await loadWorkState();
    const assignment = workState.assignments.find((item) => item.id === assignmentId);
    if (!assignment) {
      sendJson(res, 404, { errors: ["Work not found."] });
      return;
    }

    if (!canCompleteWorkAssignment(currentUser(req), assignment)) {
      sendForbidden(res);
      return;
    }

    const completedAt = new Date().toISOString();
    const completedAssignment = { ...assignment, status: "done", completedAt, updatedAt: completedAt };
    workState.messages.unshift({
      id: crypto.randomUUID(),
      text: `${workEmployeeForAssignment(workState, completedAssignment).name} finished: ${completedAssignment.title}`,
      createdAt: completedAt
    });
    workState.messages = workState.messages.slice(0, 12);
    const workCompletionNotification = await trySendWorkCompletionLarkNotification(
      workState,
      completedAssignment,
      currentUser(req)
    );
    const workCompletionEmailNotification = await trySendWorkCompletionEmailNotification(
      workState,
      completedAssignment,
      currentUser(req)
    );
    workState.assignments = workState.assignments.map((item) =>
      item.id === assignmentId ? completedAssignment : item
    );
    await saveWorkState(workState);
    const invoices = await loadInvoices();
    await syncInvoiceDatesFromCompletedJobs(invoices, workState);
    sendJson(res, 200, {
      ...visibleWorkStateForUser(workState, currentUser(req)),
      workCompletionNotification,
      workCompletionEmailNotification,
      workCompletionNotificationMessage: workCompletionNotificationMessage(workCompletionNotification, workCompletionEmailNotification),
      user: currentUser(req)
    });
    return;
  }

  const reopenWorkMatch = url.pathname.match(/^\/api\/work\/assignments\/([^/]+)\/reopen$/);
  if (req.method === "POST" && reopenWorkMatch) {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const assignmentId = decodeURIComponent(reopenWorkMatch[1]);
    const workState = await loadWorkState();
    if (!workState.assignments.some((item) => item.id === assignmentId)) {
      sendJson(res, 404, { errors: ["Work not found."] });
      return;
    }

    workState.assignments = workState.assignments.map((item) =>
      item.id === assignmentId
        ? { ...item, status: "open", completedAt: "", updatedAt: new Date().toISOString() }
        : item
    );
    await saveWorkState(workState);
    sendJson(res, 200, { ...visibleWorkStateForUser(workState, currentUser(req)), user: currentUser(req) });
    return;
  }

  const workAssignmentMatch = url.pathname.match(/^\/api\/work\/assignments\/([^/]+)$/);
  if ((req.method === "PUT" || req.method === "PATCH") && workAssignmentMatch) {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const assignmentId = decodeURIComponent(workAssignmentMatch[1]);
    const workState = await loadWorkState();
    const existing = workState.assignments.find((item) => item.id === assignmentId);
    if (!existing) {
      sendJson(res, 404, { errors: ["Work not found."] });
      return;
    }

    const input = await readBody(req);
    const { errors, assignment } = validateWorkAssignment(input, existing, workState);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    workState.assignments = workState.assignments.map((item) =>
      item.id === assignmentId ? assignment : item
    );
    await saveWorkState(workState);
    sendJson(res, 200, { assignment, ...visibleWorkStateForUser(workState, currentUser(req)), user: currentUser(req) });
    return;
  }

  if (req.method === "DELETE" && workAssignmentMatch) {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const assignmentId = decodeURIComponent(workAssignmentMatch[1]);
    const workState = await loadWorkState();
    const nextAssignments = workState.assignments.filter((item) => item.id !== assignmentId);
    if (nextAssignments.length === workState.assignments.length) {
      sendJson(res, 404, { errors: ["Work not found."] });
      return;
    }

    workState.assignments = nextAssignments;
    await saveWorkState(workState);
    sendJson(res, 200, { ...visibleWorkStateForUser(workState, currentUser(req)), user: currentUser(req) });
    return;
  }

  if (req.method === "DELETE" && url.pathname === "/api/work/messages") {
    if (!requirePermission(req, res, "view_work_messages", "Boss or team leader only.")) {
      return;
    }

    const workState = await loadWorkState();
    workState.messages = [];
    await saveWorkState(workState);
    sendJson(res, 200, { ...visibleWorkStateForUser(workState, currentUser(req)), user: currentUser(req) });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/invoices") {
    const invoices = await loadInvoices();
    await syncInvoiceDatesFromCompletedJobs(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoices });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/invoices") {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const invoices = await loadInvoices();
    const { errors, invoice } = validateManualInvoice(input, invoices);
    if (errors.length) {
      sendJson(res, errors.some((error) => error.includes("already used")) ? 409 : 400, { errors });
      return;
    }

    invoices.push(invoice);
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 201, { invoice, invoices });
    return;
  }

  const invoicePdfMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)\/pdf$/);
  if (req.method === "GET" && invoicePdfMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoicePdfMatch[1]);
    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }
    await syncInvoiceDatesFromCompletedJobs(invoices);

    try {
      const pdf = await createInvoicePdfBuffer(invoice);
      const filename = invoiceFileName(invoice).replace(/["\\\r\n]/g, "");
      const disposition = url.searchParams.get("preview") === "1" ? "inline" : "attachment";
      res.writeHead(200, {
        "Content-Type": "application/pdf",
        "Content-Length": pdf.length,
        "Content-Disposition": `${disposition}; filename="${filename}"`,
        "Cache-Control": "no-store, max-age=0",
        "X-Content-Type-Options": "nosniff"
      });
      res.end(pdf);
    } catch (error) {
      sendJson(res, 500, { errors: [error.message || "Could not create invoice PDF."] });
    }
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/invoices/sync") {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoices = await loadInvoices();
    const bookings = await loadBookings();
    const result = syncInvoicesFromBookings(invoices, bookings);
    await syncInvoiceDatesFromCompletedJobs(invoices, null, { persist: false });
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoices, ...result });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/invoices/import") {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const importedRecords = Array.isArray(input?.invoices) ? input.invoices : [];
    if (!importedRecords.length) {
      sendJson(res, 400, { errors: ["Add at least one invoice to import."] });
      return;
    }

    const invoices = await loadInvoices();
    const result = importInvoices(invoices, importedRecords);
    await saveInvoices(invoices);
    sendJson(res, 200, { invoices, ...result, total: invoices.length });
    return;
  }

  const invoiceItemsMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)\/items$/);
  if ((req.method === "PATCH" || req.method === "PUT") && invoiceItemsMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoiceItemsMatch[1]);
    const input = await readBody(req);
    const { errors, items, subtotal, gstAmount, total } = validateEditableInvoiceItems(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    if (invoice.status === "void") {
      sendJson(res, 400, { errors: ["Voided invoices cannot be edited."] });
      return;
    }

    const now = new Date().toISOString();
    invoice.items = items;
    invoice.subtotal = subtotal;
    invoice.gstRate = invoiceGstRate;
    invoice.gstAmount = gstAmount;
    invoice.gstIncluded = true;
    invoice.total = total;
    invoice.pricesEditedAt = now;
    invoice.updatedAt = now;
    await syncInvoiceDatesFromCompletedJobs(invoices, null, { persist: false });
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoice: normalizeInvoice(invoice), invoices });
    return;
  }

  const invoiceDetailsMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)\/details$/);
  if ((req.method === "PATCH" || req.method === "PUT") && invoiceDetailsMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoiceDetailsMatch[1]);
    const input = await readBody(req);
    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    if (invoice.status === "void") {
      sendJson(res, 400, { errors: ["Voided invoices cannot be edited."] });
      return;
    }

    const detailValidation = validateEditableInvoiceDetails(input, invoice);
    const itemValidation = Array.isArray(input?.items) ? validateEditableInvoiceItems(input) : null;
    const errors = [
      ...detailValidation.errors,
      ...(itemValidation?.errors || [])
    ];
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }
    const { details } = detailValidation;

    if (invoiceNumberInUse(invoices, details.invoiceNumber, invoiceId)) {
      sendJson(res, 409, { errors: [`${details.invoiceNumber} is already used by another invoice.`] });
      return;
    }

    const now = new Date().toISOString();
    Object.assign(invoice, details, {
      detailsEditedAt: now,
      updatedAt: now
    });
    if (itemValidation) {
      invoice.items = itemValidation.items;
      invoice.subtotal = itemValidation.subtotal;
      invoice.gstRate = invoiceGstRate;
      invoice.gstAmount = itemValidation.gstAmount;
      invoice.gstIncluded = true;
      invoice.total = itemValidation.total;
      invoice.pricesEditedAt = now;
    }
    await syncInvoiceDatesFromCompletedJobs(invoices, null, { persist: false });
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoice: normalizeInvoice(invoice), invoices });
    return;
  }

  const invoiceDeleteMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)$/);
  if ((req.method === "PATCH" || req.method === "PUT") && invoiceDeleteMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoiceDeleteMatch[1]);
    const input = await readBody(req);
    const { invoiceNumber, errors } = validateEditableInvoiceNumber(input?.invoiceNumber);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    if (invoiceNumberInUse(invoices, invoiceNumber, invoiceId)) {
      sendJson(res, 409, { errors: [`${invoiceNumber} is already used by another invoice.`] });
      return;
    }

    invoice.invoiceNumber = invoiceNumber;
    invoice.updatedAt = new Date().toISOString();
    await syncInvoiceDatesFromCompletedJobs(invoices, null, { persist: false });
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoice, invoices });
    return;
  }

  if (req.method === "DELETE" && invoiceDeleteMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoiceDeleteMatch[1]);
    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    if (invoice.status !== "void") {
      sendJson(res, 400, { errors: ["Only voided invoices can be deleted."] });
      return;
    }

    const nextInvoices = invoices.filter((item) => item.id !== invoiceId);
    await saveInvoices(nextInvoices);
    nextInvoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoices: nextInvoices, deletedInvoiceId: invoiceId });
    return;
  }

  const invoiceSendMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)\/send$/);
  if (req.method === "POST" && invoiceSendMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const invoiceId = decodeURIComponent(invoiceSendMatch[1]);
    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    if (invoice.status === "void") {
      sendJson(res, 400, { errors: ["Voided invoices cannot be sent."] });
      return;
    }

    const recipients = invoiceEmailRecipients(invoice, input);
    if (!recipients.length) {
      sendJson(res, 400, { errors: ["Add a valid client email before sending this invoice."] });
      return;
    }
    await syncInvoiceDatesFromCompletedJobs(invoices);

    let sendResult = null;
    try {
      sendResult = await sendInvoiceEmail(invoice, recipients);
    } catch (error) {
      const message = invoiceEmailErrorMessage(error);
      await appendSendLog({
        type: "invoice",
        status: "failed",
        title: `Invoice ${invoice.invoiceNumber}`,
        detail: invoice.propertyAddress || invoice.clientName || "",
        provider: invoiceEmailProvider,
        from: invoiceEmailConfig.from,
        recipients,
        relatedId: invoice.id,
        relatedNumber: invoice.invoiceNumber,
        error: message
      });
      sendJson(res, 502, { errors: [message] });
      return;
    }

    invoice.sentAt = new Date().toISOString();
    invoice.sentTo = recipients;
    invoice.updatedAt = invoice.sentAt;
    await saveInvoices(invoices);
    await appendSendLog({
      type: "invoice",
      status: "success",
      title: `Invoice ${invoice.invoiceNumber}`,
      detail: invoice.propertyAddress || invoice.clientName || "",
      provider: invoiceEmailProvider,
      providerMessageId: sendResult?.providerMessageId || "",
      from: invoiceEmailConfig.from,
      recipients,
      relatedId: invoice.id,
      relatedNumber: invoice.invoiceNumber
    });
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoice, invoices, message: `Invoice sent to ${recipients.join(", ")}.` });
    return;
  }

  const invoiceActionMatch = url.pathname.match(/^\/api\/invoices\/([^/]+)\/(paid|void|draft)$/);
  if (req.method === "POST" && invoiceActionMatch) {
    if (!hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const invoiceId = decodeURIComponent(invoiceActionMatch[1]);
    const status = invoiceActionMatch[2] === "paid" ? "paid" : invoiceActionMatch[2];
    const invoices = await loadInvoices();
    const invoice = invoices.find((item) => item.id === invoiceId);
    if (!invoice) {
      sendJson(res, 404, { errors: ["Invoice not found."] });
      return;
    }

    invoice.status = normalizeInvoiceStatus(status);
    invoice.updatedAt = new Date().toISOString();
    invoice.paidAt = invoice.status === "paid" ? (invoice.paidAt || invoice.updatedAt) : "";
    invoice.voidedAt = invoice.status === "void" ? (invoice.voidedAt || invoice.updatedAt) : "";
    await syncInvoiceDatesFromCompletedJobs(invoices, null, { persist: false });
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoice, invoices });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/wages") {
    const wages = await loadWages();
    wages.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { wages });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/employee-wages") {
    const employeeWages = await loadEmployeeWages();
    employeeWages.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    sendJson(res, 200, { employeeWages });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/employee-wages") {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const employeeWages = await loadEmployeeWages();
    const existingIndex = input.id
      ? employeeWages.findIndex((item) => item.id === input.id)
      : employeeWages.findIndex((item) => item.employeeName.toLowerCase() === String(input.employeeName || input.name || "").trim().toLowerCase());
    const existing = existingIndex >= 0 ? employeeWages[existingIndex] : null;
    const { errors, employeeWage } = validateEmployeeWage(input, existing, employeeWages);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    if (existingIndex >= 0) {
      employeeWages[existingIndex] = employeeWage;
    } else {
      employeeWages.push(employeeWage);
    }

    employeeWages.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    await saveEmployeeWages(employeeWages);
    sendJson(res, 200, { employeeWage, employeeWages });
    return;
  }

  const employeeWageActionMatch = url.pathname.match(/^\/api\/employee-wages\/([^/]+)\/(paid|void|draft)$/);
  if (req.method === "POST" && employeeWageActionMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const employeeWageId = decodeURIComponent(employeeWageActionMatch[1]);
    const status = employeeWageActionMatch[2];
    const employeeWages = await loadEmployeeWages();
    const employeeWage = employeeWages.find((item) => item.id === employeeWageId);
    if (!employeeWage) {
      sendJson(res, 404, { errors: ["Employee wage not found."] });
      return;
    }

    employeeWage.status = normalizeEmployeeWageStatus(status);
    employeeWage.updatedAt = new Date().toISOString();
    employeeWage.paidAt = employeeWage.status === "paid" ? (employeeWage.paidAt || employeeWage.updatedAt) : "";
    employeeWage.voidedAt = employeeWage.status === "void" ? (employeeWage.voidedAt || employeeWage.updatedAt) : "";
    await saveEmployeeWages(employeeWages);
    employeeWages.sort((a, b) => a.employeeName.localeCompare(b.employeeName));
    sendJson(res, 200, { employeeWage, employeeWages });
    return;
  }

  const employeePayslipPdfMatch = url.pathname.match(/^\/api\/employee-wages\/([^/]+)\/pdf$/);
  if (req.method === "GET" && employeePayslipPdfMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const employeeWageId = decodeURIComponent(employeePayslipPdfMatch[1]);
    const employeeWages = await loadEmployeeWages();
    const employeeWage = employeeWages.find((item) => item.id === employeeWageId);
    if (!employeeWage) {
      sendJson(res, 404, { errors: ["Employee wage not found."] });
      return;
    }

    try {
      const pdf = await createEmployeePayslipPdfBuffer(employeeWage);
      const filename = employeeWageFileName(employeeWage).replace(/["\\\r\n]/g, "");
      res.writeHead(200, {
        "Content-Type": "application/pdf",
        "Content-Length": pdf.length,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      });
      res.end(pdf);
    } catch (error) {
      sendJson(res, 500, { errors: [error.message || "Could not create employee payslip PDF."] });
    }
    return;
  }

  const employeeWageDeleteMatch = url.pathname.match(/^\/api\/employee-wages\/([^/]+)$/);
  if (req.method === "DELETE" && employeeWageDeleteMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const employeeWageId = decodeURIComponent(employeeWageDeleteMatch[1]);
    const employeeWages = await loadEmployeeWages();
    const nextEmployeeWages = employeeWages.filter((item) => item.id !== employeeWageId);
    if (nextEmployeeWages.length === employeeWages.length) {
      sendJson(res, 404, { errors: ["Employee wage not found."] });
      return;
    }

    await saveEmployeeWages(nextEmployeeWages);
    sendJson(res, 200, { employeeWages: nextEmployeeWages });
    return;
  }

  const wagePdfMatch = url.pathname.match(/^\/api\/wages\/([^/]+)\/pdf$/);
  if (req.method === "GET" && wagePdfMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const wageId = decodeURIComponent(wagePdfMatch[1]);
    const wages = await loadWages();
    const wage = wages.find((item) => item.id === wageId);
    if (!wage) {
      sendJson(res, 404, { errors: ["Wage proforma not found."] });
      return;
    }

    try {
      const pdf = await createWagePdfBuffer(wage);
      const filename = wageFileName(wage).replace(/["\\\r\n]/g, "");
      res.writeHead(200, {
        "Content-Type": "application/pdf",
        "Content-Length": pdf.length,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
      });
      res.end(pdf);
    } catch (error) {
      sendJson(res, 500, { errors: [error.message || "Could not create wage proforma PDF."] });
    }
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/wages/sync") {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const wages = await loadWages();
    const bookings = await loadBookings();
    const result = syncWagesFromBookings(wages, bookings);
    await saveWages(wages);
    wages.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { wages, ...result });
    return;
  }

  const wageSendMatch = url.pathname.match(/^\/api\/wages\/([^/]+)\/send$/);
  if (req.method === "POST" && wageSendMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const wageId = decodeURIComponent(wageSendMatch[1]);
    const wages = await loadWages();
    const wage = wages.find((item) => item.id === wageId);
    if (!wage) {
      sendJson(res, 404, { errors: ["Wage proforma not found."] });
      return;
    }

    if (wage.status === "void") {
      sendJson(res, 400, { errors: ["Voided wage proformas cannot be sent."] });
      return;
    }

    const recipients = wageEmailRecipients(wage, input);
    if (!recipients.length) {
      sendJson(res, 400, { errors: ["Add a valid photographer email before sending this proforma."] });
      return;
    }

    try {
      await sendWageEmail(wage, recipients);
    } catch (error) {
      const message = invoiceEmailErrorMessage(error);
      await appendSendLog({
        type: "wage",
        status: "failed",
        title: `Wage ${wage.wageNumber}`,
        detail: wage.propertyAddress || wage.photographerName || "",
        provider: invoiceEmailProvider,
        from: invoiceEmailConfig.from,
        recipients,
        relatedId: wage.id,
        relatedNumber: wage.wageNumber,
        error: message
      });
      sendJson(res, 502, { errors: [message] });
      return;
    }

    wage.sentAt = new Date().toISOString();
    wage.sentTo = recipients;
    wage.updatedAt = wage.sentAt;
    await saveWages(wages);
    await appendSendLog({
      type: "wage",
      status: "success",
      title: `Wage ${wage.wageNumber}`,
      detail: wage.propertyAddress || wage.photographerName || "",
      provider: invoiceEmailProvider,
      from: invoiceEmailConfig.from,
      recipients,
      relatedId: wage.id,
      relatedNumber: wage.wageNumber
    });
    wages.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { wage, wages, message: `Proforma sent to ${recipients.join(", ")}.` });
    return;
  }

  const wageActionMatch = url.pathname.match(/^\/api\/wages\/([^/]+)\/(paid|void|draft)$/);
  if (req.method === "POST" && wageActionMatch) {
    if (!hasPermission(req, "manage_wages")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const wageId = decodeURIComponent(wageActionMatch[1]);
    const status = wageActionMatch[2] === "paid" ? "paid" : wageActionMatch[2];
    const wages = await loadWages();
    const wage = wages.find((item) => item.id === wageId);
    if (!wage) {
      sendJson(res, 404, { errors: ["Wage proforma not found."] });
      return;
    }

    wage.status = normalizeWageStatus(status);
    wage.updatedAt = new Date().toISOString();
    wage.paidAt = wage.status === "paid" ? (wage.paidAt || wage.updatedAt) : "";
    wage.voidedAt = wage.status === "void" ? (wage.voidedAt || wage.updatedAt) : "";
    await saveWages(wages);
    wages.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { wage, wages });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/bookings") {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const bookings = await loadBookings();
    let larkBookings = [];
    let larkImportError = null;

    try {
      larkBookings = await fetchLarkImportedBookings();
    } catch (error) {
      larkImportError = error.message;
    }

    const mergedBookings = mergeLocalAndLarkBookings(bookings, larkBookings)
      .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    sendJson(res, 200, {
      bookings: mergedBookings,
      larkImportedCount: larkBookings.length,
      larkImportError
    });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/address-suggestions") {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const query = url.searchParams.get("q") || "";
    try {
      const suggestions = await findAddressSuggestions(query);
      sendJson(res, 200, { suggestions });
    } catch (error) {
      sendJson(res, 502, { errors: [error.message || "Address lookup failed."] });
    }
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/clients") {
    if (!canAccessApp(req, "bookings") && !canAccessApp(req, "clients")) {
      sendForbidden(res);
      return;
    }

    const clients = await loadClients();
    clients.sort((a, b) => a.name.localeCompare(b.name));
    sendJson(res, 200, { clients });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/clients") {
    if (!hasPermission(req, "manage_directory")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const { errors, client } = validateClient(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const clients = await loadClients();
    const existingIndex = client.id
      ? clients.findIndex((item) => item.id === client.id)
      : clients.findIndex((item) => clientIdentityKey(item) === clientIdentityKey(client));
    const now = new Date().toISOString();

    if (existingIndex >= 0) {
      clients[existingIndex] = mergeClientRecord(clients[existingIndex], client, now);
    } else {
      clients.push({
        ...mergeClientRecord(null, client, now),
        id: crypto.randomUUID(),
        createdAt: now,
      });
    }

    clients.sort((a, b) => a.name.localeCompare(b.name));
    await saveClients(clients);
    const savedClient = existingIndex >= 0
      ? clients[existingIndex]
      : clients.find((item) => clientIdentityKey(item) === clientIdentityKey(client));
    sendJson(res, 200, { client: savedClient, clients });
    return;
  }

  const clientIdMatch = url.pathname.match(/^\/api\/clients\/([^/]+)$/);
  if (req.method === "DELETE" && clientIdMatch) {
    if (!hasPermission(req, "manage_directory")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const clientId = decodeURIComponent(clientIdMatch[1]);
    const clients = await loadClients();
    const nextClients = clients.filter((item) => item.id !== clientId);

    if (nextClients.length === clients.length) {
      sendJson(res, 404, { errors: ["Client not found."] });
      return;
    }

    await saveClients(nextClients);
    sendJson(res, 200, { clients: nextClients });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/photographers") {
    if (!canAccessApp(req, "bookings") && !canAccessApp(req, "photographers")) {
      sendForbidden(res);
      return;
    }

    const photographers = await loadPhotographers();
    photographers.sort((a, b) => a.name.localeCompare(b.name));
    sendJson(res, 200, { photographers });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/photographers") {
    if (!hasPermission(req, "manage_directory")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const input = await readBody(req);
    const { errors, photographer } = validatePhotographer(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const photographers = await loadPhotographers();
    const existingIndex = photographer.id
      ? photographers.findIndex((item) => item.id === photographer.id)
      : photographers.findIndex((item) => item.name.toLowerCase() === photographer.name.toLowerCase());
    const now = new Date().toISOString();

    if (existingIndex >= 0) {
      photographers[existingIndex] = {
        ...photographers[existingIndex],
        name: photographer.name,
        email: photographer.email,
        phone: photographer.phone,
        gstIncluded: photographer.gstIncluded,
        updatedAt: now
      };
    } else {
      photographers.push({
        id: crypto.randomUUID(),
        name: photographer.name,
        email: photographer.email,
        phone: photographer.phone,
        gstIncluded: photographer.gstIncluded,
        createdAt: now,
        updatedAt: now
      });
    }

    photographers.sort((a, b) => a.name.localeCompare(b.name));
    await savePhotographers(photographers);
    const savedPhotographer = photographer.id
      ? photographers.find((item) => item.id === photographer.id)
      : photographers.find((item) => item.name.toLowerCase() === photographer.name.toLowerCase());
    sendJson(res, 200, { photographer: savedPhotographer, photographers });
    return;
  }

  const photographerIdMatch = url.pathname.match(/^\/api\/photographers\/([^/]+)$/);
  if (req.method === "DELETE" && photographerIdMatch) {
    if (!hasPermission(req, "manage_directory")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    const photographerId = decodeURIComponent(photographerIdMatch[1]);
    const photographers = await loadPhotographers();
    const nextPhotographers = photographers.filter((item) => item.id !== photographerId);

    if (nextPhotographers.length === photographers.length) {
      sendJson(res, 404, { errors: ["Photographer not found."] });
      return;
    }

    await savePhotographers(nextPhotographers);
    sendJson(res, 200, { photographers: nextPhotographers });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/lark/preview") {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const input = await readBody(req);
    const { errors, value } = validateBooking(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const booking = {
      id: "preview",
      ...value,
      endAt: getEndAt(value.startAt, value.durationMinutes)
    };

    sendJson(res, 200, { preview: buildLarkPreview(booking) });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/bookings") {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const input = await readBody(req);
    const { errors, value } = validateBooking(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const bookings = await loadBookings();
    const endAt = getEndAt(value.startAt, value.durationMinutes);
    if (hasBookingConflict(bookings, value.startAt, endAt)) {
      sendJson(res, 409, { errors: ["That time is already booked. Choose another slot."] });
      return;
    }

    const booking = {
      id: crypto.randomUUID(),
      ...value,
      endAt,
      status: "confirmed",
      larkStatus: isLarkConfigured() ? "pending" : "not_configured",
      larkEventId: null,
      larkError: null,
      larkAttendeeStatus: isLarkConfigured() ? (value.guestEmails.length ? "pending" : "not_needed") : "not_configured",
      larkAttendeeError: null,
      calendarInviteStatus: value.guestEmails.length ? "pending" : "not_needed",
      calendarInviteSentAt: "",
      calendarInviteTo: value.guestEmails.join(", "),
      calendarInviteEmailFrom: calendarInviteEmailConfig.from,
      calendarInviteError: null,
      calendarInviteSequence: 0,
      createdAt: new Date().toISOString()
    };

    if (bookingHasPassed(booking)) {
      markPastBookingLocal(booking);
    } else {
      await syncBookingToLark(booking);
    }

    bookings.push(booking);
    await saveBookings(bookings);
    const invoices = await loadInvoices();
    const invoiceResult = upsertInvoiceForBooking(invoices, booking);
    await saveInvoices(invoices);
    const wages = await loadWages();
    const wageResult = upsertWageForBooking(wages, booking);
    await saveWages(wages);
    sendJson(res, 201, { booking, invoice: invoiceResult.invoice, wage: wageResult.wage });
    return;
  }

  if ((req.method === "PUT" || req.method === "PATCH") && url.pathname.startsWith("/api/bookings/")) {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const id = bookingIdFromApiPath(url.pathname);
    if (!id) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const input = await readBody(req);
    const { errors, value } = validateBooking(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const bookings = await loadBookings();
    const bookingIndex = bookings.findIndex((item) => item.id === id);
    if (bookingIndex === -1) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const existingBooking = bookings[bookingIndex];
    if (existingBooking.status === "cancelled") {
      sendJson(res, 400, { errors: ["Cancelled bookings cannot be edited."] });
      return;
    }

    const endAt = getEndAt(value.startAt, value.durationMinutes);
    if (hasBookingConflict(bookings, value.startAt, endAt, id)) {
      sendJson(res, 409, { errors: ["That time is already booked. Choose another slot."] });
      return;
    }

    const booking = {
      ...existingBooking,
      ...value,
      id: existingBooking.id,
      endAt,
      status: existingBooking.status || "confirmed",
      larkStatus: existingBooking.larkStatus || (isLarkConfigured() ? "pending" : "not_configured"),
      larkEventId: existingBooking.larkEventId || null,
      larkError: existingBooking.larkError || null,
      larkAttendeeStatus:
        existingBooking.larkAttendeeStatus || (isLarkConfigured() ? (value.guestEmails.length ? "pending" : "not_needed") : "not_configured"),
      larkAttendeeError: existingBooking.larkAttendeeError || null,
      calendarInviteStatus: existingBooking.calendarInviteStatus || (value.guestEmails.length ? "pending" : "not_needed"),
      calendarInviteSentAt: existingBooking.calendarInviteSentAt || "",
      calendarInviteTo: value.guestEmails.join(", "),
      calendarInviteEmailFrom: existingBooking.calendarInviteEmailFrom || calendarInviteEmailConfig.from,
      calendarInviteError: existingBooking.calendarInviteError || null,
      calendarInviteSequence: Number(existingBooking.calendarInviteSequence || 0),
      createdAt: existingBooking.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    if (bookingHasPassed(booking)) {
      markPastBookingLocal(booking);
    } else {
      await syncBookingToLark(booking, existingBooking);
    }

    bookings[bookingIndex] = booking;
    await saveBookings(bookings);
    const invoices = await loadInvoices();
    const invoiceResult = upsertInvoiceForBooking(invoices, booking);
    await saveInvoices(invoices);
    const wages = await loadWages();
    const wageResult = upsertWageForBooking(wages, booking);
    await saveWages(wages);
    sendJson(res, 200, { booking, invoice: invoiceResult.invoice, wage: wageResult.wage });
    return;
  }

  if (req.method === "POST" && url.pathname.startsWith("/api/bookings/") && url.pathname.endsWith("/cancel")) {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const id = bookingIdFromApiPath(url.pathname, "cancel");
    if (!id) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const bookings = await loadBookings();
    const booking = bookings.find((item) => item.id === id);

    if (!booking) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    booking.status = "cancelled";
    booking.cancelledAt = new Date().toISOString();
    await cancelBookingInLark(booking);
    await trySendBookingCalendarInviteEmail(booking, "CANCEL");
    await saveBookings(bookings);
    const invoices = await loadInvoices();
    const invoiceResult = upsertInvoiceForBooking(invoices, booking);
    await saveInvoices(invoices);
    const wages = await loadWages();
    const wageResult = upsertWageForBooking(wages, booking);
    await saveWages(wages);
    sendJson(res, 200, { booking, invoice: invoiceResult.invoice, wage: wageResult.wage });
    return;
  }

  if (req.method === "DELETE" && url.pathname.startsWith("/api/bookings/")) {
    if (!canAccessApp(req, "bookings")) {
      sendForbidden(res);
      return;
    }

    const id = bookingIdFromApiPath(url.pathname);
    if (!id) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const bookings = await loadBookings();
    const bookingIndex = bookings.findIndex((item) => item.id === id);
    if (bookingIndex === -1) {
      if (id.startsWith("lark-")) {
        sendJson(res, 200, { removedBookingId: id, bookings });
        return;
      }

      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const booking = bookings[bookingIndex];
    if (booking.status !== "cancelled") {
      sendJson(res, 400, { errors: ["Cancel the booking before removing it from the system."] });
      return;
    }

    const forceLocalDelete = ["1", "true", "yes"].includes(String(url.searchParams.get("force") || "").toLowerCase());
    if (!forceLocalDelete && booking.larkEventId && booking.larkStatus !== "cancelled") {
      await cancelBookingInLark(booking);
      if (booking.larkStatus === "delete_failed") {
        await saveBookings(bookings);
        sendJson(res, 502, {
          booking,
          errors: [booking.larkError || "Calendar cancellation failed. Try removing again after the calendar cancellation succeeds."]
        });
        return;
      }
    }

    if (!forceLocalDelete && booking.calendarInviteStatus !== "cancelled") {
      await trySendBookingCalendarInviteEmail(booking, "CANCEL");
    }

    bookings.splice(bookingIndex, 1);
    await saveBookings(bookings);
    sendJson(res, 200, { removedBookingId: id, bookings });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/lark/test") {
    if (!hasPermission(req, "manage_bookings")) {
      sendForbidden(res, "Boss or team leader only.");
      return;
    }

    if (!isLarkConfigured()) {
      sendJson(res, 400, { ok: false, message: "Add LARK_APP_ID, LARK_APP_SECRET, and LARK_CALENDAR_ID or LARK_ORGANIZER_CALENDAR_ID first." });
      return;
    }

    try {
      await getTenantAccessToken();
      sendJson(res, 200, { ok: true, message: "Lark credentials returned a tenant access token." });
    } catch (error) {
      sendJson(res, 502, { ok: false, message: error.message });
    }
    return;
  }

  if (req.method === "OPTIONS") {
    sendNoContent(res);
    return;
  }

  sendJson(res, 404, { errors: ["API route not found."] });
}

async function serveStatic(req, res, url) {
  const routePath = url.pathname === "/login" ? "/login.html" : url.pathname;
  const routeFiles = {
    "/": "/portal.html",
    "/bookings": "/index.html",
    "/invoices": "/index.html",
    "/wages": "/wages.html",
    "/settings": "/settings.html",
    "/work": "/work/index.html",
    "/work/": "/work/index.html"
  };
  const appPath = routeFiles[routePath] || routePath;
  const requestedPath = decodeURIComponent(appPath);
  let filePath = path.normalize(path.join(publicDir, requestedPath));

  if (!filePath.startsWith(publicDir)) {
    res.writeHead(403);
    res.end("Forbidden");
    return;
  }

  try {
    const fileStat = await stat(filePath);
    if (!fileStat.isFile()) {
      throw new Error("Not a file");
    }
  } catch {
    filePath = path.join(publicDir, "index.html");
  }

  const extension = path.extname(filePath);
  const stream = createReadStream(filePath);

  stream.on("error", () => {
    if (!res.headersSent) {
      sendJson(res, 404, { errors: ["File not found."] });
    } else {
      res.end();
    }
  });

  res.writeHead(200, {
    "Content-Type": contentTypes[extension] || "application/octet-stream",
    "Cache-Control": "no-store"
  });
  stream.pipe(res);
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);

    if (url.pathname.startsWith("/api/")) {
      applyCorsHeaders(req, res);
      if (req.method === "OPTIONS") {
        sendNoContent(res);
        return;
      }

      await handleApi(req, res, url);
      return;
    }

    const isLoginPage = url.pathname === "/login" || url.pathname === "/login.html";
    const publicAssets = new Set([
      "/apple-touch-icon.png",
      "/favicon-32.png",
      "/favicon-48.png",
      "/favicon.ico",
      "/icons/icon-192.png",
      "/icons/icon-512.png",
      "/install.js",
      "/openframe-logo.png",
      "/session-keepalive.js",
      "/service-worker.js",
      "/site.webmanifest",
      "/styles.css"
    ]);
    const isPublicAsset = isLoginPage || publicAssets.has(url.pathname);

    if (!isAuthenticated(req) && !isPublicAsset) {
      redirect(res, "/login");
      return;
    }

    if (isAuthenticated(req) && isLoginPage) {
      redirect(res, "/");
      return;
    }

    const requestedApp = appForRoute(url.pathname);
    if (requestedApp && !userCanAccessApp(currentUser(req), requestedApp)) {
      const fallbackApp = currentUser(req)?.apps?.[0] || "bookings";
      redirect(res, fallbackApp === "work" ? "/work/" : "/");
      return;
    }

    await serveStatic(req, res, url);
  } catch (error) {
    sendJson(res, 500, { errors: [error.message || "Unexpected server error."] });
  }
});

await prepareDataStorage();

server.listen(port, host, () => {
  console.log(`Booking app running at http://${host}:${port}`);
});
