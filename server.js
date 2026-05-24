import http from "node:http";
import { mkdir, readFile, rename, stat, writeFile } from "node:fs/promises";
import { createReadStream, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");

loadLocalEnv();

const bundledDataDir = path.join(__dirname, "data");
const dataDir = process.env.DATA_DIR ? path.resolve(process.env.DATA_DIR) : bundledDataDir;
const bookingsFile = path.join(dataDir, "bookings.json");
const clientsFile = path.join(dataDir, "clients.json");
const photographersFile = path.join(dataDir, "photographers.json");
const workFile = path.join(dataDir, "work-assignments.json");
const invoicesFile = path.join(dataDir, "invoices.json");
const wagesFile = path.join(dataDir, "wages.json");
const usersFile = path.join(dataDir, "users.json");
const seedFiles = {
  bookings: path.join(bundledDataDir, "bookings.json"),
  clients: path.join(bundledDataDir, "clients.json"),
  photographers: path.join(bundledDataDir, "photographers.json"),
  work: path.join(bundledDataDir, "work-assignments.json"),
  invoices: path.join(bundledDataDir, "invoices.json"),
  wages: path.join(bundledDataDir, "wages.json"),
  users: path.join(bundledDataDir, "users.json")
};
const githubStorage = {
  token: process.env.GITHUB_STORAGE_TOKEN || "",
  repo: process.env.GITHUB_STORAGE_REPO || "OpenFrameStudio/InternalBooking",
  branch: process.env.GITHUB_STORAGE_BRANCH || "data-store",
  pathPrefix: (process.env.GITHUB_STORAGE_PATH || "data").replace(/^\/+|\/+$/g, "")
};
const storageBackend = githubStorage.token && githubStorage.repo ? "github" : (process.env.DATA_DIR ? "disk" : "app");

const port = Number(process.env.PORT || 4180);
const host = process.env.HOST || (process.env.PORT ? "0.0.0.0" : "127.0.0.1");
const cancelledBookingRetentionHours = Number(process.env.CANCELLED_BOOKING_RETENTION_HOURS || 12);
const cancelledBookingRetentionMs =
  (Number.isFinite(cancelledBookingRetentionHours) && cancelledBookingRetentionHours > 0 ? cancelledBookingRetentionHours : 12) * 60 * 60 * 1000;
const publicAppUrl = (process.env.APP_PUBLIC_URL || "https://system.openframe.studio").replace(/\/$/, "");
const larkConfig = {
  appId: process.env.LARK_APP_ID || "",
  appSecret: process.env.LARK_APP_SECRET || "",
  calendarId: process.env.LARK_CALENDAR_ID || "primary",
  organizerCalendarId: process.env.LARK_ORGANIZER_CALENDAR_ID || "",
  organizerUserId: process.env.LARK_ORGANIZER_USER_ID || "",
  senderEmail: process.env.LARK_SENDER_EMAIL || "admin@openframe.studio",
  senderName: process.env.LARK_SENDER_NAME || "admin@openframe.studio",
  timezone: process.env.LARK_TIMEZONE || "Australia/Sydney",
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
    apps: ["bookings", "clients", "photographers", "work", "invoices", "wages"],
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
  }
];
const workDeskOrigins = new Set(
  (process.env.WORK_DESK_ORIGINS || "http://127.0.0.1:4173,http://localhost:4173")
    .split(",")
    .map((origin) => origin.trim())
    .filter(Boolean)
);

let tenantTokenCache = null;
const sessions = new Map();
const sessionCookieName = "internalbooking_session";
const sessionMaxAgeSeconds = 7 * 24 * 60 * 60;
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
const wageCurrency = "AUD";
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
const defaultWorkState = {
  employee: {
    id: "faye",
    name: "Faye",
    email: parseGuestEmails(workInviteEmailConfig.to)[0] || "",
    role: "Editor / Admin",
    availability: "Mon-Fri, 12pm-8pm Australian time"
  },
  assignments: [],
  messages: []
};
const dataFiles = {
  bookings: {
    file: bookingsFile,
    seedFile: seedFiles.bookings,
    githubPath: githubDataPath("bookings.json"),
    fallback: []
  },
  clients: {
    file: clientsFile,
    seedFile: seedFiles.clients,
    githubPath: githubDataPath("clients.json"),
    fallback: []
  },
  photographers: {
    file: photographersFile,
    seedFile: seedFiles.photographers,
    githubPath: githubDataPath("photographers.json"),
    fallback: defaultPhotographers
  },
  work: {
    file: workFile,
    seedFile: seedFiles.work,
    githubPath: githubDataPath("work-assignments.json"),
    fallback: defaultWorkState
  },
  invoices: {
    file: invoicesFile,
    seedFile: seedFiles.invoices,
    githubPath: githubDataPath("invoices.json"),
    fallback: []
  },
  wages: {
    file: wagesFile,
    seedFile: seedFiles.wages,
    githubPath: githubDataPath("wages.json"),
    fallback: []
  },
  users: {
    file: usersFile,
    seedFile: seedFiles.users,
    githubPath: githubDataPath("users.json"),
    fallback: []
  }
};

function loadLocalEnv() {
  try {
    const raw = readFileSync(path.join(__dirname, ".env"), "utf8");
    for (const line of raw.split(/\r?\n/)) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith("#")) {
        continue;
      }

      const equalsAt = trimmed.indexOf("=");
      if (equalsAt === -1) {
        continue;
      }

      const key = trimmed.slice(0, equalsAt).trim();
      let value = trimmed.slice(equalsAt + 1).trim();
      if ((value.startsWith('"') && value.endsWith('"')) || (value.startsWith("'") && value.endsWith("'"))) {
        value = value.slice(1, -1);
      }

      if (key && process.env[key] === undefined) {
        process.env[key] = value;
      }
    }
  } catch (error) {
    if (error.code !== "ENOENT") {
      console.warn(`Could not load .env: ${error.message}`);
    }
  }
}

function normalizeEnvBoolean(value, fallback = false) {
  if (value === undefined || value === null || value === "") {
    return fallback;
  }

  return ["1", "true", "yes", "on"].includes(String(value).trim().toLowerCase());
}

function sendJson(res, status, payload, headers = {}) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    ...headers
  });
  res.end(body);
}

function sendNoContent(res) {
  res.writeHead(204);
  res.end();
}

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

function redirect(res, location) {
  res.writeHead(302, {
    Location: location,
    "Cache-Control": "no-store"
  });
  res.end();
}

function parseCookies(header = "") {
  const cookies = {};

  for (const part of header.split(";")) {
    const [rawKey, ...rawValue] = part.trim().split("=");
    if (!rawKey) {
      continue;
    }

    try {
      cookies[rawKey] = decodeURIComponent(rawValue.join("="));
    } catch {
      cookies[rawKey] = rawValue.join("=");
    }
  }

  return cookies;
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

function sessionUserPayload(user) {
  return {
    username: user.username,
    role: user.role,
    label: user.label,
    name: user.name || user.username,
    employeeId: user.employeeId || "",
    apps: Array.isArray(user.apps) ? user.apps : [],
    permissions: Array.isArray(user.permissions) ? user.permissions : []
  };
}

function createSession(user) {
  cleanupExpiredSessions();
  const token = crypto.randomBytes(32).toString("base64url");
  sessions.set(token, {
    expiresAt: Date.now() + sessionMaxAgeSeconds * 1000,
    user: sessionUserPayload(user)
  });
  return token;
}

function currentSession(req) {
  const token = getSessionToken(req);
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

function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    req.on("data", (chunk) => {
      data += chunk;
      if (data.length > 1_000_000) {
        reject(new Error("Request body is too large."));
        req.destroy();
      }
    });
    req.on("end", () => {
      if (!data) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(data));
      } catch {
        reject(new Error("Request body must be valid JSON."));
      }
    });
    req.on("error", reject);
  });
}

async function readJsonFile(filePath, fallback) {
  try {
    const raw = await readFile(filePath, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === "ENOENT") {
      return fallback;
    }
    throw error;
  }
}

async function writeJsonFile(filePath, value) {
  await mkdir(path.dirname(filePath), { recursive: true });
  const tempFile = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  await writeFile(tempFile, `${JSON.stringify(value, null, 2)}\n`, "utf8");
  await rename(tempFile, filePath);
}

function githubDataPath(filename) {
  return [githubStorage.pathPrefix, filename].filter(Boolean).join("/");
}

function githubContentsUrl(filePath) {
  const parts = githubStorage.repo.split("/");
  if (parts.length !== 2 || !parts[0] || !parts[1]) {
    throw new Error("GITHUB_STORAGE_REPO must look like owner/repo.");
  }

  const repoPath = parts.map((part) => encodeURIComponent(part)).join("/");
  const contentPath = filePath.split("/").map((part) => encodeURIComponent(part)).join("/");
  return `https://api.github.com/repos/${repoPath}/contents/${contentPath}`;
}

async function parseGithubError(response) {
  const data = await response.json().catch(() => ({}));
  return data.message || `GitHub storage request failed with ${response.status}.`;
}

async function getGithubContent(filePath) {
  const url = `${githubContentsUrl(filePath)}?ref=${encodeURIComponent(githubStorage.branch)}`;
  const response = await fetch(url, {
    headers: {
      Accept: "application/vnd.github+json",
      Authorization: `Bearer ${githubStorage.token}`,
      "X-GitHub-Api-Version": "2022-11-28",
      "User-Agent": "OpenFrameInternalBooking/1.0"
    }
  });

  if (response.status === 404) {
    return null;
  }

  if (!response.ok) {
    throw new Error(await parseGithubError(response));
  }

  return response.json();
}

async function readGithubJson(filePath, fallback) {
  const content = await getGithubContent(filePath);
  if (!content) {
    return fallback;
  }

  if (content.type !== "file" || !content.content) {
    throw new Error(`GitHub storage path ${filePath} is not a file.`);
  }

  const raw = Buffer.from(content.content.replace(/\s/g, ""), "base64").toString("utf8");
  return JSON.parse(raw);
}

async function writeGithubJson(filePath, value) {
  const json = `${JSON.stringify(value, null, 2)}\n`;
  const content = Buffer.from(json, "utf8").toString("base64");

  for (let attempt = 0; attempt < 2; attempt += 1) {
    const current = await getGithubContent(filePath);
    const body = {
      message: `Update ${filePath}`,
      content,
      branch: githubStorage.branch
    };
    if (current?.sha) {
      body.sha = current.sha;
    }

    const response = await fetch(githubContentsUrl(filePath), {
      method: "PUT",
      headers: {
        Accept: "application/vnd.github+json",
        Authorization: `Bearer ${githubStorage.token}`,
        "Content-Type": "application/json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "OpenFrameInternalBooking/1.0"
      },
      body: JSON.stringify(body)
    });

    if (response.ok) {
      return;
    }

    if (response.status === 409 && attempt === 0) {
      continue;
    }

    throw new Error(await parseGithubError(response));
  }
}

async function fileExists(filePath) {
  try {
    await stat(filePath);
    return true;
  } catch (error) {
    if (error.code === "ENOENT") {
      return false;
    }
    throw error;
  }
}

async function seedDataFile(targetFile, seedFile, fallback) {
  if (await fileExists(targetFile)) {
    return;
  }

  const seed = await readJsonFile(seedFile, fallback);
  await writeJsonFile(targetFile, seed);
}

async function readStoredJson(dataFile) {
  if (storageBackend === "github") {
    return readGithubJson(dataFile.githubPath, dataFile.fallback);
  }

  return readJsonFile(dataFile.file, dataFile.fallback);
}

async function writeStoredJson(dataFile, value) {
  if (storageBackend === "github") {
    await writeGithubJson(dataFile.githubPath, value);
    return;
  }

  await writeJsonFile(dataFile.file, value);
}

async function seedStoredDataFile(dataFile) {
  if (storageBackend === "github") {
    const existing = await getGithubContent(dataFile.githubPath);
    if (existing) {
      return;
    }

    const seed = await readJsonFile(dataFile.seedFile, dataFile.fallback);
    await writeGithubJson(dataFile.githubPath, seed);
    return;
  }

  await seedDataFile(dataFile.file, dataFile.seedFile, dataFile.fallback);
}

async function prepareDataStorage() {
  if (storageBackend !== "github") {
    await mkdir(dataDir, { recursive: true });
  }

  await Promise.all([
    seedStoredDataFile(dataFiles.bookings),
    seedStoredDataFile(dataFiles.clients),
    seedStoredDataFile(dataFiles.photographers),
    seedStoredDataFile(dataFiles.work),
    seedStoredDataFile(dataFiles.invoices),
    seedStoredDataFile(dataFiles.wages),
    seedStoredDataFile(dataFiles.users)
  ]);
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

function normalizeWorkState(raw) {
  const employee = {
    ...defaultWorkState.employee,
    ...(raw?.employee || {}),
    email: String(raw?.employee?.email || defaultWorkState.employee.email || "").trim()
  };

  const assignments = Array.isArray(raw?.assignments)
    ? raw.assignments.map((assignment) => ({
        id: assignment.id || crypto.randomUUID(),
        employeeId: defaultWorkState.employee.id,
        title: String(assignment.title || "Untitled work"),
        dueDate: /^\d{4}-\d{2}-\d{2}$/.test(String(assignment.dueDate || ""))
          ? assignment.dueDate
          : toDateValue(new Date()),
        priority: ["high", "normal", "low"].includes(assignment.priority) ? assignment.priority : "normal",
        notes: String(assignment.notes || ""),
        status: assignment.status === "done" ? "done" : "open",
        source: String(assignment.source || ""),
        sourceId: String(assignment.sourceId || ""),
        inviteStatus: ["sent", "failed", "not_configured"].includes(assignment.inviteStatus) ? assignment.inviteStatus : "",
        inviteSentAt: assignment.inviteSentAt || "",
        inviteTo: uniqueEmails(parseGuestEmails(assignment.inviteTo || "")).join(", "),
        inviteEmailFrom: String(assignment.inviteEmailFrom || ""),
        inviteError: String(assignment.inviteError || ""),
        completedAt: assignment.completedAt || "",
        createdAt: assignment.createdAt || new Date().toISOString(),
        updatedAt: assignment.updatedAt || assignment.createdAt || new Date().toISOString()
      }))
    : [];

  const messages = Array.isArray(raw?.messages) ? raw.messages : [];

  return { employee, assignments, messages };
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

  return {
    ...workState,
    assignments: workState.assignments.filter((assignment) => assignment.employeeId === user?.employeeId),
    messages: []
  };
}

function toDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function parseDateValue(value) {
  const [year, month, day] = String(value || "").split("-").map(Number);
  return new Date(year, month - 1, day);
}

function validateWorkAssignment(input, existing = null) {
  const title = String(input.title || "").trim();
  const dueDate = String(input.dueDate || "").trim();
  const priority = ["high", "normal", "low"].includes(input.priority) ? input.priority : "normal";
  const notes = String(input.notes || "").trim();
  const errors = [];

  if (!title) errors.push("Enter the work title.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDate)) errors.push("Choose a valid due date.");

  return {
    errors,
    assignment: {
      id: existing?.id || crypto.randomUUID(),
      employeeId: defaultWorkState.employee.id,
      title,
      dueDate,
      priority,
      notes,
      status: existing?.status || "open",
      source: existing?.source || String(input.source || ""),
      sourceId: existing?.sourceId || String(input.sourceId || ""),
      inviteStatus: existing?.inviteStatus || "",
      inviteSentAt: existing?.inviteSentAt || "",
      inviteTo: existing?.inviteTo || "",
      inviteEmailFrom: existing?.inviteEmailFrom || "",
      inviteError: existing?.inviteError || "",
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
    month: "short"
  }).format(start);
  const time = new Intl.DateTimeFormat("en-AU", {
    hour: "numeric",
    minute: "2-digit"
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
  return toDateValue(due < today ? today : due);
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
  if (Number.isNaN(end.getTime())) return false;
  return end >= parseDateValue(toDateValue(new Date()));
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
    employeeId: defaultWorkState.employee.id,
    title,
    dueDate: workdayBeforeBooking(booking),
    priority: bookingWorkPriority(booking),
    notes,
    status: existing?.status || "open",
    source: "booking",
    sourceId: booking.id,
    inviteStatus: existing?.inviteStatus || "",
    inviteSentAt: existing?.inviteSentAt || "",
    inviteTo: existing?.inviteTo || "",
    inviteEmailFrom: existing?.inviteEmailFrom || "",
    inviteError: existing?.inviteError || "",
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

function normalizeMoney(value) {
  const number = Number(value);
  return Number.isFinite(number) ? Math.round(number * 100) / 100 : 0;
}

function calculateGst(subtotal) {
  return normalizeMoney(subtotal * invoiceGstRate);
}

function normalizeInvoice(invoice) {
  const items = Array.isArray(invoice?.items)
    ? invoice.items.map((item) => ({
        name: String(item.name || "Service").trim() || "Service",
        quantity: Math.max(1, Number(item.quantity || 1)),
        unitPrice: normalizeMoney(item.unitPrice),
        amount: normalizeMoney(item.amount ?? Number(item.quantity || 1) * Number(item.unitPrice || 0))
      }))
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
    gstRate,
    gstAmount,
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

function nextInvoiceNumber(invoices) {
  const maxNumber = invoices.reduce((max, invoice) => {
    const value = String(invoice.invoiceNumber || "");
    const receiptMatch = value.match(/^R(\d+)$/);
    if (receiptMatch) return Math.max(max, Number(receiptMatch[1]));

    const legacyMatch = value.match(/^INV-\d{4}-(\d{4,})$/);
    if (legacyMatch) return Math.max(max, Number(legacyMatch[1]));
    return max;
  }, 0);
  return `R${String(maxNumber + 1).padStart(3, "0")}`;
}

function invoiceFromBooking(booking, invoices, existing = null) {
  const now = new Date().toISOString();
  const isPaid = existing?.status === "paid";
  const existingItems = Array.isArray(existing?.items) ? existing.items : [];
  const items = isPaid ? existingItems : bookingInvoiceItems(booking);
  const subtotal = normalizeMoney(items.reduce((sum, item) => sum + item.amount, 0));
  const gstAmount = calculateGst(subtotal);
  const status = booking.status === "cancelled"
    ? (isPaid ? "paid" : "void")
    : (isPaid ? "paid" : "draft");

  return normalizeInvoice({
    ...(existing || {}),
    id: existing?.id || crypto.randomUUID(),
    invoiceNumber: existing?.invoiceNumber || nextInvoiceNumber(invoices),
    bookingId: booking.id,
    propertyAddress: booking.propertyAddress || "",
    clientName: booking.clientName || "",
    clientEmail: booking.clientEmail || "",
    agentName: booking.agentName || "",
    agentPhone: booking.agentPhone || "",
    photographerName: booking.photographerName || "",
    bookingStartAt: booking.startAt || "",
    bookingEndAt: booking.endAt || "",
    items,
    subtotal,
    gstRate: invoiceGstRate,
    gstAmount,
    total: normalizeMoney(subtotal + gstAmount),
    currency: invoiceCurrency,
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

function bookingHasPassed(booking, now = Date.now()) {
  const endSource = booking?.endAt || (
    booking?.startAt && booking?.durationMinutes
      ? getEndAt(booking.startAt, Number(booking.durationMinutes))
      : ""
  );
  const endDate = new Date(endSource);
  return !Number.isNaN(endDate.getTime()) && endDate.getTime() < now;
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

function formatInvoiceDocumentDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) {
    return new Date().toISOString().slice(0, 10);
  }

  return date.toISOString().slice(0, 10);
}

function invoiceProductDescription(invoice) {
  const serviceNames = (invoice.items || []).map((item) => item.name).filter(Boolean).join(" + ");
  return [invoice.propertyAddress || "Booking", serviceNames].filter(Boolean).join("\n");
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

function drawInvoicePdf(doc, invoice) {
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
  doc.font("Helvetica").fontSize(10.5).text(invoice.propertyAddress || invoice.clientName || "Client", billingX, 269, {
    width: pageWidth - margin - billingX,
    height: 82,
    ellipsis: true
  });

  const tableX = margin;
  const tableY = 392;
  const columnWidths = [182, 82, 112, 115];
  const rowHeight = 58;
  const headerHeight = 34;
  const tableWidth = columnWidths.reduce((sum, width) => sum + width, 0);
  const headers = ["PRODUCT", "QUANTITY", "PRICE", "SUBTOTAL"];
  const product = invoiceProductDescription(invoice);

  doc.rect(tableX, tableY, tableWidth, headerHeight).fill("#e7e7e7");
  doc.fillColor("#000").font("Helvetica-Bold").fontSize(10);
  let cursorX = tableX;
  headers.forEach((header, index) => {
    drawPdfCell(doc, header, cursorX, tableY, columnWidths[index], headerHeight, { align: "center" });
    cursorX += columnWidths[index];
  });

  doc.font("Helvetica").fontSize(10);
  cursorX = tableX;
  const rowY = tableY + headerHeight;
  drawPdfCell(doc, product, cursorX, rowY, columnWidths[0], rowHeight, { align: "center" });
  cursorX += columnWidths[0];
  drawPdfCell(doc, "1", cursorX, rowY, columnWidths[1], rowHeight, { align: "center" });
  cursorX += columnWidths[1];
  drawPdfCell(doc, formatInvoiceMoney(invoice.subtotal), cursorX, rowY, columnWidths[2], rowHeight, { align: "right" });
  cursorX += columnWidths[2];
  drawPdfCell(doc, formatInvoiceMoney(invoice.subtotal), cursorX, rowY, columnWidths[3], rowHeight, { align: "right" });
  doc.moveTo(tableX, rowY + rowHeight).lineTo(tableX + tableWidth, rowY + rowHeight).lineWidth(1.2).stroke("#000");

  doc.save();
  doc.rotate(-10, { origin: [140, 575] });
  doc.rect(110, 559, 96, 30).lineWidth(2).stroke("#f5a623");
  doc.fillColor("#f5a623").font("Helvetica-Bold").fontSize(18).text(invoice.status === "paid" ? "PAID" : invoice.status === "void" ? "VOID" : "UNPAID", 118, 565);
  doc.restore();

  const totalX = 324;
  const totalY = 524;
  const totalW = 219;
  const totalRowH = 33;
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

  doc.font("Helvetica-Bold").fontSize(13).fillColor("#000").text("PAYMENT INFORMATION", margin, 676);
  doc.font("Helvetica").fontSize(10.5).fillColor("#787878");
  [
    "Bank Transfer:",
    "Name: Openframe Studio Pty Ltd",
    "BSB: 062-128",
    "Account: 11440602",
    `Please reference ${invoice.invoiceNumber} for the payment`
  ].forEach((line, index) => doc.text(line, margin, 704 + index * 17, {
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

  drawInvoicePdf(doc, invoice);
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
    await sendInvoiceWithResend(invoice, recipients, pdf, bcc);
    return;
  }

  const transport = await createInvoiceTransport();
  await withTimeout(transport.sendMail({
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
}

function normalizeWageStatus(status) {
  return ["draft", "paid", "void"].includes(status) ? status : "draft";
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

function workInviteRecipients(workState = null) {
  return uniqueEmails([
    ...parseGuestEmails(workInviteEmailConfig.to),
    ...parseGuestEmails(workState?.employee?.email || "")
  ]).filter(isValidEmail);
}

function workInviteEmailMissingSettings(workState = null) {
  const missing = [];
  if (!workInviteEmailConfig.enabled) missing.push("WORK_INVITE_EMAIL_ENABLED");
  if (!resendConfig.apiKey) missing.push("RESEND_API_KEY");
  if (!workInviteRecipients(workState).length) missing.push("WORK_INVITE_EMAIL_TO");
  return missing;
}

function isWorkInviteEmailConfigured() {
  return workInviteEmailMissingSettings().length === 0;
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
  const lines = [
    `Hi ${workState?.employee?.name || "Faye"},`,
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

  lines.push(
    "Open the work desk:",
    `${publicAppUrl}/work/`,
    "",
    "Thank you,",
    "OpenFrame Studio"
  );

  return lines.join("\n");
}

function buildWorkInviteEmailHtml(assignment, workState) {
  const notes = assignment.notes
    ? `<p><strong>Details:</strong><br />${escapeHtmlForEmail(assignment.notes).replace(/\n/g, "<br />")}</p>`
    : "";

  return `
    <div style="font-family:Arial,sans-serif;color:#111611;line-height:1.5">
      <p>Hi ${escapeHtmlForEmail(workState?.employee?.name || "Faye")},</p>
      <p>A new work item has been assigned to you in the OpenFrame internal booking system.</p>
      <p><strong>Work:</strong> ${escapeHtmlForEmail(assignment.title)}</p>
      <p><strong>Due:</strong> ${escapeHtmlForEmail(formatWorkInviteDueDate(assignment))}</p>
      <p><strong>Priority:</strong> ${escapeHtmlForEmail(assignment.priority)}</p>
      ${notes}
      <p><a href="${escapeHtmlForEmail(`${publicAppUrl}/work/`)}">Open the work desk</a></p>
      <p>Thank you,<br />OpenFrame Studio</p>
    </div>
  `;
}

async function sendWorkInviteEmail(assignment, workState) {
  const missing = workInviteEmailMissingSettings(workState);
  if (missing.length) {
    return { sent: false, skipped: true, missing };
  }

  const recipients = workInviteRecipients(workState);
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

async function sendWorkInviteEmails(workState, assignments) {
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
      } else {
        summary.skipped += 1;
        summary.missing = uniqueEmails([...summary.missing, ...(result.missing || [])]);
        assignment.inviteStatus = "not_configured";
        assignment.inviteEmailFrom = workInviteEmailConfig.from;
        assignment.inviteError = result.missing?.length
          ? `Missing ${result.missing.join(", ")}`
          : "Work invite email is not configured.";
      }
    } catch (error) {
      summary.failed += 1;
      const message = error.message || "Could not send work invite email.";
      summary.errors.push(message);
      assignment.inviteStatus = "failed";
      assignment.inviteEmailFrom = workInviteEmailConfig.from;
      assignment.inviteError = message;
    }
  }

  return summary;
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
    return `Work saved. Add WORK_INVITE_EMAIL_TO in Render to email Faye from ${workInviteEmailConfig.from}.`;
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
    return { sent: true, recipients };
  } catch (error) {
    if (shouldTrack) {
      booking.calendarInviteStatus = "failed";
      booking.calendarInviteEmailFrom = calendarInviteEmailConfig.from;
      booking.calendarInviteError = error.message || "Could not send calendar invite email.";
    }
    return { sent: false, failed: true, error: error.message || "Could not send calendar invite email." };
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
  const client = {
    id: String(input.id || "").trim(),
    name: String(input.name || input.clientName || "").trim(),
    email: String(input.email || input.clientEmail || "").trim(),
    agentName: String(input.agentName || "").trim(),
    agentPhone: String(input.agentPhone || "").trim()
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
  const startAt = String(input.startAt || "").trim();
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

  if (startDate.getTime() < Date.now() - 60_000) {
    errors.push("Choose a future time.");
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
    const parsed = new Date(`${time.date}T00:00:00`);
    if (!Number.isNaN(parsed.getTime())) {
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

  return [...new Map(imported.map((booking) => [booking.larkEventId, booking])).values()];
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
    sendJson(res, 200, { authenticated: Boolean(user), user });
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
      calendarInviteEmailConfigured: isCalendarInviteEmailConfigured(),
      calendarInviteEmailFrom: calendarInviteEmailConfig.from,
      larkNotificationsEnabled: shouldUseLarkCalendarNotifications(),
      larkAttendeesEnabled: shouldAddLarkAttendees(),
      timezone: larkConfig.timezone,
      persistentStorage: storageBackend !== "app",
      storageBackend
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

  if (req.method === "POST" && url.pathname === "/api/cleanup/past-bookings-invoices") {
    if (!hasPermission(req, "manage_bookings") || !hasPermission(req, "manage_invoices")) {
      sendForbidden(res, "Boss only.");
      return;
    }

    const result = await deletePastBookingsAndInvoices();
    sendJson(res, 200, result);
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
    await saveWorkState(workState);
    sendJson(res, 200, {
      ...visibleWorkStateForUser(workState, currentUser(req)),
      ...syncResult,
      workInvite,
      workInviteMessage: workInviteMessage(workInvite),
      user: currentUser(req)
    });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/work/assignments") {
    if (!requirePermission(req, res, "manage_work", "Boss or team leader only.")) {
      return;
    }

    const input = await readBody(req);
    const { errors, assignment } = validateWorkAssignment(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const workState = await loadWorkState();
    workState.assignments.push(assignment);
    const workInvite = await sendWorkInviteEmails(workState, [assignment]);
    await saveWorkState(workState);
    sendJson(res, 201, {
      assignment,
      ...visibleWorkStateForUser(workState, currentUser(req)),
      workInvite,
      workInviteMessage: workInviteMessage(workInvite),
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
    workState.assignments = workState.assignments.map((item) =>
      item.id === assignmentId
        ? { ...item, status: "done", completedAt, updatedAt: completedAt }
        : item
    );
    workState.messages.unshift({
      id: crypto.randomUUID(),
      text: `${workState.employee.name} finished: ${assignment.title}`,
      createdAt: completedAt
    });
    workState.messages = workState.messages.slice(0, 12);
    await saveWorkState(workState);
    sendJson(res, 200, { ...visibleWorkStateForUser(workState, currentUser(req)), user: currentUser(req) });
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
    const { errors, assignment } = validateWorkAssignment(input, existing);
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
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoices });
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

    try {
      const pdf = await createInvoicePdfBuffer(invoice);
      const filename = invoiceFileName(invoice).replace(/["\\\r\n]/g, "");
      res.writeHead(200, {
        "Content-Type": "application/pdf",
        "Content-Length": pdf.length,
        "Content-Disposition": `attachment; filename="${filename}"`,
        "Cache-Control": "no-store"
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
    await saveInvoices(invoices);
    invoices.sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
    sendJson(res, 200, { invoices, ...result });
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

    try {
      await sendInvoiceEmail(invoice, recipients);
    } catch (error) {
      sendJson(res, 502, { errors: [invoiceEmailErrorMessage(error)] });
      return;
    }

    invoice.sentAt = new Date().toISOString();
    invoice.sentTo = recipients;
    invoice.updatedAt = invoice.sentAt;
    await saveInvoices(invoices);
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
      sendJson(res, 502, { errors: [invoiceEmailErrorMessage(error)] });
      return;
    }

    wage.sentAt = new Date().toISOString();
    wage.sentTo = recipients;
    wage.updatedAt = wage.sentAt;
    await saveWages(wages);
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
      clients[existingIndex] = {
        ...clients[existingIndex],
        name: client.name,
        email: client.email,
        agentName: client.agentName,
        agentPhone: client.agentPhone,
        updatedAt: now
      };
    } else {
      clients.push({
        id: crypto.randomUUID(),
        name: client.name,
        email: client.email,
        agentName: client.agentName,
        agentPhone: client.agentPhone,
        createdAt: now,
        updatedAt: now
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

    await syncBookingToLark(booking);

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

    const [, apiPart, bookingsPart, id, extra] = url.pathname.split("/");
    if (apiPart !== "api" || bookingsPart !== "bookings" || !id || extra) {
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

    await syncBookingToLark(booking, existingBooking);

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

    const id = url.pathname.split("/")[3];
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

    const [, apiPart, bookingsPart, id, extra] = url.pathname.split("/");
    if (apiPart !== "api" || bookingsPart !== "bookings" || !id || extra) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const bookings = await loadBookings();
    const bookingIndex = bookings.findIndex((item) => item.id === id);
    if (bookingIndex === -1) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    const booking = bookings[bookingIndex];
    if (booking.status !== "cancelled") {
      sendJson(res, 400, { errors: ["Cancel the booking before removing it from the system."] });
      return;
    }

    if (booking.larkEventId && booking.larkStatus !== "cancelled") {
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

    if (booking.calendarInviteStatus !== "cancelled") {
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
  const appPath =
    routePath === "/" ? "/portal.html"
      : routePath === "/bookings" ? "/index.html"
        : routePath === "/invoices" ? "/index.html"
          : routePath === "/wages" ? "/index.html"
        : routePath === "/work" || routePath === "/work/" ? "/work/index.html"
          : routePath;
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
      "/favicon.ico",
      "/icons/icon-192.png",
      "/icons/icon-512.png",
      "/install.js",
      "/openframe-logo.png",
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
