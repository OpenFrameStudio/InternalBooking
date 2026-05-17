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
const seedFiles = {
  bookings: path.join(bundledDataDir, "bookings.json"),
  clients: path.join(bundledDataDir, "clients.json"),
  photographers: path.join(bundledDataDir, "photographers.json")
};

const port = Number(process.env.PORT || 4180);
const host = process.env.HOST || (process.env.PORT ? "0.0.0.0" : "127.0.0.1");
const larkConfig = {
  appId: process.env.LARK_APP_ID || "",
  appSecret: process.env.LARK_APP_SECRET || "",
  calendarId: process.env.LARK_CALENDAR_ID || "primary",
  timezone: process.env.LARK_TIMEZONE || "Australia/Sydney",
  apiBase: (process.env.LARK_API_BASE || "https://open.larksuite.com/open-apis").replace(/\/$/, "")
};
const authConfig = {
  username: process.env.ADMIN_USERNAME || "OpenFrame",
  password: process.env.ADMIN_PASSWORD || "Studio"
};

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
  ".svg": "image/svg+xml",
  ".png": "image/png",
  ".ico": "image/x-icon"
};

const serviceCatalog = new Set(["Photography", "Floorplan", "Drone"]);
const larkImportDays = Number(process.env.LARK_IMPORT_DAYS || 120);
const defaultPhotographers = [
  {
    id: "default-barry",
    name: "Barry",
    email: process.env.DEFAULT_PHOTOGRAPHER_EMAIL || "",
    phone: "0403 007 853",
    createdAt: "2026-05-16T00:00:00.000+10:00",
    updatedAt: "2026-05-16T00:00:00.000+10:00"
  }
];

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

function createSession() {
  cleanupExpiredSessions();
  const token = crypto.randomBytes(32).toString("base64url");
  sessions.set(token, { expiresAt: Date.now() + sessionMaxAgeSeconds * 1000 });
  return token;
}

function isAuthenticated(req) {
  const token = getSessionToken(req);
  const session = sessions.get(token);
  if (!session) {
    return false;
  }

  if (session.expiresAt <= Date.now()) {
    sessions.delete(token);
    return false;
  }

  return true;
}

function safeEquals(left, right) {
  const leftBuffer = Buffer.from(String(left || ""));
  const rightBuffer = Buffer.from(String(right || ""));
  return leftBuffer.length === rightBuffer.length && crypto.timingSafeEqual(leftBuffer, rightBuffer);
}

function credentialsAreValid(input) {
  return safeEquals(input.username, authConfig.username) && safeEquals(input.password, authConfig.password);
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

async function prepareDataStorage() {
  await mkdir(dataDir, { recursive: true });
  await Promise.all([
    seedDataFile(bookingsFile, seedFiles.bookings, []),
    seedDataFile(clientsFile, seedFiles.clients, []),
    seedDataFile(photographersFile, seedFiles.photographers, defaultPhotographers)
  ]);
}

async function loadBookings() {
  try {
    return await readJsonFile(bookingsFile, []);
  } catch (error) {
    throw error;
  }
}

async function saveBookings(bookings) {
  await writeJsonFile(bookingsFile, bookings);
}

async function loadClients() {
  try {
    return await readJsonFile(clientsFile, []);
  } catch (error) {
    throw error;
  }
}

async function saveClients(clients) {
  await writeJsonFile(clientsFile, clients);
}

async function loadPhotographers() {
  try {
    return await readJsonFile(photographersFile, defaultPhotographers);
  } catch (error) {
    throw error;
  }
}

async function savePhotographers(photographers) {
  await writeJsonFile(photographersFile, photographers);
}

function isLarkConfigured() {
  return Boolean(larkConfig.appId && larkConfig.appSecret && larkConfig.calendarId);
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

  const response = await fetch(`https://nominatim.openstreetmap.org/search?${params}`, {
    headers: {
      "User-Agent": "OpenFrameInternalBooking/1.0 (internalbooking.openframe.studio)",
      "Accept-Language": "en-AU,en;q=0.9"
    }
  });

  if (!response.ok) {
    throw new Error(`Address lookup failed with ${response.status}.`);
  }

  const results = await response.json().catch(() => []);
  const suggestions = uniqueAddressSuggestions(
    results.map((result) => ({
      address: compactAddress(result) || result.display_name,
      label: result.display_name
    }))
  );

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

function validatePhotographer(input) {
  const errors = [];
  const photographer = {
    id: String(input.id || "").trim(),
    name: String(input.name || input.photographerName || "").trim(),
    email: String(input.email || input.photographerEmail || "").trim(),
    phone: String(input.phone || input.photographerPhone || "").trim()
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

  lines.push("", `Booking ID: ${booking.id}`);
  return lines.join("\n");
}

function buildLarkEventPayload(booking) {
  const startTime = Math.floor(new Date(booking.startAt).getTime() / 1000).toString();
  const endTime = Math.floor(new Date(booking.endAt).getTime() / 1000).toString();

  return {
    summary: booking.propertyAddress,
    description: buildLarkDescription(booking),
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
    location: {
      name: booking.locationName || booking.propertyAddress,
      address: booking.locationAddress || booking.propertyAddress
    },
    reminders: [{ minutes: 30 }]
  };
}

function buildLarkAttendeesPayload(booking, emails = booking.guestEmails || []) {
  if (!emails.length) {
    return null;
  }

  return {
    need_notification: true,
    attendees: emails.map((email) => ({
      type: "third_party",
      third_party_email: email
    }))
  };
}

function buildLarkPreview(booking) {
  const eventPayload = buildLarkEventPayload(booking);
  const attendeesPayload = buildLarkAttendeesPayload(booking);

  return {
    title: eventPayload.summary,
    startAt: booking.startAt,
    endAt: booking.endAt,
    timezone: larkConfig.timezone,
    location: eventPayload.location,
    description: eventPayload.description,
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
  const calendarId = encodeURIComponent(larkConfig.calendarId);
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
  const calendarId = encodeURIComponent(larkConfig.calendarId);

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
  const calendarId = encodeURIComponent(larkConfig.calendarId);
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

function getAddedGuestEmails(previousEmails = [], nextEmails = []) {
  const previous = new Set(previousEmails.map((email) => email.toLowerCase()));
  return nextEmails.filter((email) => !previous.has(email.toLowerCase()));
}

function getRemovedGuestEmails(previousEmails = [], nextEmails = []) {
  const next = new Set(nextEmails.map((email) => email.toLowerCase()));
  return previousEmails.filter((email) => !next.has(email.toLowerCase()));
}

async function createLarkAttendees(booking, emails = booking.guestEmails || []) {
  if (!booking.larkEventId || !emails.length) {
    return null;
  }

  const token = await getTenantAccessToken();
  const calendarId = encodeURIComponent(larkConfig.calendarId);
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

    if (addedGuestEmails.length) {
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
      booking.larkAttendeeStatus = "needs_review";
      booking.larkAttendeeError = "Guest list changed. Remove old guests from Lark manually if needed.";
    }
  } catch (error) {
    booking.larkStatus = "failed";
    booking.larkError = error.message;
    booking.larkAttendeeStatus = booking.guestEmails.length ? "not_sent" : "not_needed";
  }

  return booking;
}

async function handleApi(req, res, url) {
  if (req.method === "GET" && url.pathname === "/api/session") {
    sendJson(res, 200, { authenticated: isAuthenticated(req) });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/login") {
    const input = await readBody(req);
    if (!credentialsAreValid(input)) {
      sendJson(res, 401, { errors: ["Login details did not match."] });
      return;
    }

    const token = createSession();
    sendJson(res, 200, { ok: true }, { "Set-Cookie": buildSessionCookie(req, token) });
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
      calendarId: larkConfig.calendarId,
      timezone: larkConfig.timezone,
      persistentStorage: Boolean(process.env.DATA_DIR)
    });
    return;
  }

  if (!isAuthenticated(req)) {
    sendJson(res, 401, { errors: ["Log in first."] });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/bookings") {
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
    const clients = await loadClients();
    clients.sort((a, b) => a.name.localeCompare(b.name));
    sendJson(res, 200, { clients });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/clients") {
    const input = await readBody(req);
    const { errors, client } = validateClient(input);
    if (errors.length) {
      sendJson(res, 400, { errors });
      return;
    }

    const clients = await loadClients();
    const existingIndex = client.id
      ? clients.findIndex((item) => item.id === client.id)
      : clients.findIndex((item) => item.name.toLowerCase() === client.name.toLowerCase());
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
    const savedClient = clients.find((item) => item.name.toLowerCase() === client.name.toLowerCase());
    sendJson(res, 200, { client: savedClient, clients });
    return;
  }

  const clientIdMatch = url.pathname.match(/^\/api\/clients\/([^/]+)$/);
  if (req.method === "DELETE" && clientIdMatch) {
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
    const photographers = await loadPhotographers();
    photographers.sort((a, b) => a.name.localeCompare(b.name));
    sendJson(res, 200, { photographers });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/photographers") {
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
        updatedAt: now
      };
    } else {
      photographers.push({
        id: crypto.randomUUID(),
        name: photographer.name,
        email: photographer.email,
        phone: photographer.phone,
        createdAt: now,
        updatedAt: now
      });
    }

    photographers.sort((a, b) => a.name.localeCompare(b.name));
    await savePhotographers(photographers);
    const savedPhotographer = photographers.find((item) => item.name.toLowerCase() === photographer.name.toLowerCase());
    sendJson(res, 200, { photographer: savedPhotographer, photographers });
    return;
  }

  const photographerIdMatch = url.pathname.match(/^\/api\/photographers\/([^/]+)$/);
  if (req.method === "DELETE" && photographerIdMatch) {
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
      createdAt: new Date().toISOString()
    };

    await syncBookingToLark(booking);

    bookings.push(booking);
    await saveBookings(bookings);
    sendJson(res, 201, { booking });
    return;
  }

  if ((req.method === "PUT" || req.method === "PATCH") && url.pathname.startsWith("/api/bookings/")) {
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
      createdAt: existingBooking.createdAt || new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };

    await syncBookingToLark(booking, existingBooking);

    bookings[bookingIndex] = booking;
    await saveBookings(bookings);
    sendJson(res, 200, { booking });
    return;
  }

  if (req.method === "POST" && url.pathname.startsWith("/api/bookings/") && url.pathname.endsWith("/cancel")) {
    const id = url.pathname.split("/")[3];
    const bookings = await loadBookings();
    const booking = bookings.find((item) => item.id === id);

    if (!booking) {
      sendJson(res, 404, { errors: ["Booking not found."] });
      return;
    }

    booking.status = "cancelled";
    booking.cancelledAt = new Date().toISOString();
    await saveBookings(bookings);
    sendJson(res, 200, { booking });
    return;
  }

  if (req.method === "POST" && url.pathname === "/api/lark/test") {
    if (!isLarkConfigured()) {
      sendJson(res, 400, { ok: false, message: "Add LARK_APP_ID, LARK_APP_SECRET, and LARK_CALENDAR_ID first." });
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
  const requestedPath = decodeURIComponent(routePath === "/" ? "/index.html" : routePath);
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
      await handleApi(req, res, url);
      return;
    }

    const isLoginPage = url.pathname === "/login" || url.pathname === "/login.html";
    const isPublicAsset = isLoginPage || url.pathname === "/styles.css" || url.pathname === "/favicon.ico";

    if (!isAuthenticated(req) && !isPublicAsset) {
      redirect(res, "/login");
      return;
    }

    if (isAuthenticated(req) && isLoginPage) {
      redirect(res, "/");
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
