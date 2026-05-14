import http from "node:http";
import { readFile, stat, writeFile } from "node:fs/promises";
import { createReadStream, readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import crypto from "node:crypto";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const publicDir = path.join(__dirname, "public");
const bookingsFile = path.join(__dirname, "data", "bookings.json");
const clientsFile = path.join(__dirname, "data", "clients.json");

loadLocalEnv();

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

async function loadBookings() {
  try {
    const raw = await readFile(bookingsFile, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
}

async function saveBookings(bookings) {
  await writeFile(bookingsFile, `${JSON.stringify(bookings, null, 2)}\n`, "utf8");
}

async function loadClients() {
  try {
    const raw = await readFile(clientsFile, "utf8");
    return JSON.parse(raw);
  } catch (error) {
    if (error.code === "ENOENT") {
      return [];
    }
    throw error;
  }
}

async function saveClients(clients) {
  await writeFile(clientsFile, `${JSON.stringify(clients, null, 2)}\n`, "utf8");
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

function validateClient(input) {
  const errors = [];
  const client = {
    id: String(input.id || "").trim(),
    name: String(input.name || input.clientName || "").trim(),
    email: String(input.email || input.clientEmail || "").trim(),
    agentName: String(input.agentName || "").trim(),
    agentPhone: String(input.agentPhone || "").trim()
  };

  if (!client.name) {
    errors.push("Enter the agency or client name.");
  }

  if (client.email && !isValidEmail(client.email)) {
    errors.push("Enter a valid client email address.");
  }

  return { errors, client };
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
  const photographerName = String(input.photographerName || "Barry").trim();
  const photographerPhone = String(input.photographerPhone || "").trim();
  const agentName = String(input.agentName || "").trim();
  const agentPhone = String(input.agentPhone || "").trim();
  const guestEmails = uniqueEmails([clientEmail, ...parseGuestEmails(input.guestEmails)]);
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
  if (clientEmail && !isValidEmail(clientEmail)) errors.push("Enter a valid client email address.");
  if (!photographerName) errors.push("Enter the photographer name.");
  for (const email of guestEmails) {
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
      clientEmail,
      photographerName,
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
      timezone: larkConfig.timezone
    });
    return;
  }

  if (!isAuthenticated(req)) {
    sendJson(res, 401, { errors: ["Log in first."] });
    return;
  }

  if (req.method === "GET" && url.pathname === "/api/bookings") {
    const bookings = await loadBookings();
    bookings.sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
    sendJson(res, 200, { bookings });
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

server.listen(port, host, () => {
  console.log(`Booking app running at http://${host}:${port}`);
});
