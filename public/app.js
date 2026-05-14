const bookingForm = document.querySelector("#bookingForm");
const bookingList = document.querySelector("#bookingList");
const bookingTemplate = document.querySelector("#bookingTemplate");
const clientTemplate = document.querySelector("#clientTemplate");
const formMessage = document.querySelector("#formMessage");
const clientMessage = document.querySelector("#clientMessage");
const bookingPage = document.querySelector("#bookingPage");
const clientsPage = document.querySelector("#clientsPage");
const navLinks = [...document.querySelectorAll("[data-route]")];
const dateInput = document.querySelector("#dateInput");
const timeInput = document.querySelector("#timeInput");
const durationInput = document.querySelector("#durationInput");
const serviceInputs = [...document.querySelectorAll('input[name="services"]')];
const clientSelect = document.querySelector("#clientSelect");
const clientNameInput = document.querySelector("#clientNameInput");
const clientEmailInput = document.querySelector("#clientEmailInput");
const agentNameInput = document.querySelector("#agentNameInput");
const agentPhoneInput = document.querySelector("#agentPhoneInput");
const guestEmailsInput = document.querySelector("#guestEmailsInput");
const invitationEmails = document.querySelector("#invitationEmails");
const clientForm = document.querySelector("#clientForm");
const clientList = document.querySelector("#clientList");
const directoryClientId = document.querySelector("#directoryClientId");
const directoryClientName = document.querySelector("#directoryClientName");
const directoryClientEmail = document.querySelector("#directoryClientEmail");
const directoryAgentName = document.querySelector("#directoryAgentName");
const directoryAgentPhone = document.querySelector("#directoryAgentPhone");
const newClientButton = document.querySelector("#newClientButton");
const refreshButton = document.querySelector("#refreshButton");
const testLarkButton = document.querySelector("#testLarkButton");
const previewLarkButton = document.querySelector("#previewLarkButton");
const logoutButton = document.querySelector("#logoutButton");
const larkPreview = document.querySelector("#larkPreview");
const previewTitle = document.querySelector("#previewTitle");
const previewTime = document.querySelector("#previewTime");
const previewLocation = document.querySelector("#previewLocation");
const previewGuests = document.querySelector("#previewGuests");
const previewDescription = document.querySelector("#previewDescription");
const larkDot = document.querySelector("#larkDot");
const larkTitle = document.querySelector("#larkTitle");
const larkDetail = document.querySelector("#larkDetail");
const todayCount = document.querySelector("#todayCount");
const weekCount = document.querySelector("#weekCount");

const dateFormatter = new Intl.DateTimeFormat(undefined, {
  weekday: "short",
  month: "short",
  day: "numeric"
});

const timeFormatter = new Intl.DateTimeFormat(undefined, {
  hour: "numeric",
  minute: "2-digit"
});

const monthFormatter = new Intl.DateTimeFormat(undefined, { month: "short" });
const dayFormatter = new Intl.DateTimeFormat(undefined, { day: "2-digit" });

let bookings = [];
let clients = [];
let syncedClientEmail = "";
let isRestoringBookingDraft = false;

const bookingDraftKey = "openframe.bookingDraft.v2";
const clientsCacheKey = "openframe.clients.v1";

const clientExamples = [
  {
    agency: "Stonebridge Collective",
    email: "ava@stonebridge.example",
    agent: "Ava Brooks",
    phone: "0412 684 209"
  },
  {
    agency: "Northline Estates",
    email: "leo@northline.example",
    agent: "Leo Tran",
    phone: "0426 318 774"
  },
  {
    agency: "Cedar & Coast Realty",
    email: "nina@cedarcoast.example",
    agent: "Nina Patel",
    phone: "0437 902 146"
  },
  {
    agency: "Blue Lantern Property",
    email: "eli@bluelantern.example",
    agent: "Eli Morris",
    phone: "0408 571 663"
  },
  {
    agency: "Park Row Agency",
    email: "sophie@parkrow.example",
    agent: "Sophie Lin",
    phone: "0419 246 805"
  }
];

function readStoredJson(key) {
  try {
    return JSON.parse(localStorage.getItem(key) || "null");
  } catch {
    return null;
  }
}

function writeStoredJson(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Draft saving is a convenience; booking creation should still work without it.
  }
}

function removeStoredJson(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    // Ignore browsers that block local storage.
  }
}

function getRandomClientExample() {
  return clientExamples[Math.floor(Math.random() * clientExamples.length)];
}

function applyClientExamplePlaceholders(example = getRandomClientExample()) {
  clientNameInput.placeholder = example.agency;
  clientEmailInput.placeholder = example.email;
  agentNameInput.placeholder = example.agent;
  agentPhoneInput.placeholder = example.phone;
  directoryClientName.placeholder = example.agency;
  directoryClientEmail.placeholder = example.email;
  directoryAgentName.placeholder = example.agent;
  directoryAgentPhone.placeholder = example.phone;
}

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));

  if (response.status === 401) {
    window.location.href = "/login";
    throw new Error("Log in first.");
  }

  return { response, data };
}

function parseEmailList(value) {
  const emailMatches = String(value || "").match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  if (emailMatches?.length) {
    return emailMatches.map((email) => email.trim());
  }

  return String(value || "")
    .split(/[\s,;]+/)
    .map((email) => email.trim())
    .filter(Boolean);
}

function uniqueEmails(emails) {
  const seen = new Set();
  const unique = [];

  for (const email of emails) {
    const key = email.toLowerCase();
    if (seen.has(key)) {
      continue;
    }

    seen.add(key);
    unique.push(email);
  }

  return unique;
}

function isEmailish(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function getSelectedServices() {
  return serviceInputs
    .filter((input) => input.checked)
    .map((input) => ({ name: input.value }));
}

function getRecommendedDurationMinutes(services = getSelectedServices()) {
  if (services.length >= 3) {
    return 90;
  }

  if (services.length === 2) {
    return 60;
  }

  if (services.length === 1) {
    return 45;
  }

  return Number(durationInput.value) || 45;
}

function updateDurationForServices() {
  durationInput.value = String(getRecommendedDurationMinutes());
}

function getCachedClients() {
  const cachedClients = readStoredJson(clientsCacheKey);
  return Array.isArray(cachedClients) ? cachedClients.filter((client) => client?.name) : [];
}

function mergeClientLists(...clientLists) {
  const merged = new Map();

  for (const list of clientLists) {
    for (const client of list || []) {
      const name = String(client?.name || "").trim();
      if (!name) {
        continue;
      }

      const key = name.toLowerCase();
      const current = merged.get(key);
      const next = {
        id: String(client.id || current?.id || ""),
        name,
        email: String(client.email || ""),
        agentName: String(client.agentName || ""),
        agentPhone: String(client.agentPhone || ""),
        createdAt: client.createdAt || current?.createdAt || new Date().toISOString(),
        updatedAt: client.updatedAt || current?.updatedAt || new Date().toISOString()
      };

      if (!current || new Date(next.updatedAt).getTime() >= new Date(current.updatedAt).getTime()) {
        merged.set(key, next);
      }
    }
  }

  return [...merged.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function storeClientCache(nextClients) {
  writeStoredJson(clientsCacheKey, nextClients);
}

function getBookingDraft() {
  const data = Object.fromEntries(new FormData(bookingForm).entries());
  return {
    propertyAddress: data.propertyAddress || "",
    selectedClientId: clientSelect.value,
    clientName: clientNameInput.value,
    clientEmail: clientEmailInput.value,
    agentName: agentNameInput.value,
    agentPhone: agentPhoneInput.value,
    services: getSelectedServices().map((service) => service.name),
    date: dateInput.value,
    time: timeInput.value,
    durationMinutes: durationInput.value,
    photographerName: data.photographerName || "",
    photographerPhone: data.photographerPhone || "",
    guestEmails: guestEmailsInput.value,
    notes: data.notes || ""
  };
}

function saveBookingDraft() {
  if (isRestoringBookingDraft) {
    return;
  }

  writeStoredJson(bookingDraftKey, getBookingDraft());
}

function restoreBookingDraft() {
  const draft = readStoredJson(bookingDraftKey);
  if (!draft || typeof draft !== "object") {
    return false;
  }

  isRestoringBookingDraft = true;

  bookingForm.elements.propertyAddress.value = draft.propertyAddress || "";
  clientSelect.value = draft.selectedClientId || "";
  clientNameInput.value = draft.clientName || "";
  clientEmailInput.value = draft.clientEmail || "";
  agentNameInput.value = draft.agentName || "";
  agentPhoneInput.value = draft.agentPhone || "";
  dateInput.value = draft.date || dateInput.value;
  timeInput.value = draft.time || timeInput.value;
  durationInput.value = draft.durationMinutes || durationInput.value;
  bookingForm.elements.photographerName.value = draft.photographerName || "Barry";
  bookingForm.elements.photographerPhone.value = draft.photographerPhone || "0403 007 853";
  guestEmailsInput.value = draft.guestEmails || "";
  bookingForm.elements.notes.value = draft.notes || "";

  if (Array.isArray(draft.services) && draft.services.length) {
    const selectedServices = new Set(draft.services);
    for (const input of serviceInputs) {
      input.checked = selectedServices.has(input.value);
    }
  }

  syncedClientEmail = isEmailish(clientEmailInput.value.trim()) ? clientEmailInput.value.trim() : "";
  updateInvitationSummary();
  isRestoringBookingDraft = false;
  return true;
}

function getBookingServiceLabel(booking) {
  if (Array.isArray(booking.services) && booking.services.length) {
    return booking.services.map((service) => service.name).join(" + ");
  }

  return booking.service || "Booking";
}

function formatContact(name, phone) {
  if (name && phone) {
    return `${name} (${phone})`;
  }

  return name || phone || "";
}

function toDateValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function setInitialDateTime() {
  const now = new Date();
  now.setMinutes(Math.ceil(now.getMinutes() / 15) * 15, 0, 0);
  if (now.getHours() >= 18) {
    now.setDate(now.getDate() + 1);
    now.setHours(9, 0, 0, 0);
  }

  dateInput.min = toDateValue(new Date());
  dateInput.value = toDateValue(now);
  timeInput.value = `${String(now.getHours()).padStart(2, "0")}:${String(now.getMinutes()).padStart(2, "0")}`;
}

function setMessage(message, tone = "") {
  formMessage.textContent = message;
  formMessage.className = tone;
}

function setClientMessage(message, tone = "") {
  clientMessage.textContent = message;
  clientMessage.className = tone;
}

function setRoute(route, push = true) {
  const nextRoute = route === "/clients" ? "/clients" : "/";
  bookingPage.hidden = nextRoute !== "/";
  clientsPage.hidden = nextRoute !== "/clients";

  for (const link of navLinks) {
    link.classList.toggle("active", link.dataset.route === nextRoute);
  }

  if (push && window.location.pathname !== nextRoute) {
    history.pushState({ route: nextRoute }, "", nextRoute);
  }

  if (nextRoute === "/clients") {
    renderClientList();
  }
}

function syncClientEmailToGuests() {
  const nextEmail = clientEmailInput.value.trim();
  let emails = parseEmailList(guestEmailsInput.value);

  if (syncedClientEmail && syncedClientEmail.toLowerCase() !== nextEmail.toLowerCase()) {
    emails = emails.filter((email) => email.toLowerCase() !== syncedClientEmail.toLowerCase());
  }

  if (isEmailish(nextEmail)) {
    emails = uniqueEmails([nextEmail, ...emails]);
    syncedClientEmail = nextEmail;
  } else {
    syncedClientEmail = "";
  }

  guestEmailsInput.value = emails.join(", ");
  updateInvitationSummary();
}

function updateInvitationSummary() {
  const emails = uniqueEmails([clientEmailInput.value.trim(), ...parseEmailList(guestEmailsInput.value)]).filter(isEmailish);
  invitationEmails.textContent = emails.length ? emails.join(", ") : "No invitation emails yet.";
}

function renderClientOptions(selectedId = clientSelect.value) {
  clientSelect.innerHTML = "";

  const emptyOption = document.createElement("option");
  emptyOption.value = "";
  emptyOption.textContent = "New client";
  clientSelect.append(emptyOption);

  const sortedClients = [...clients].sort((a, b) => a.name.localeCompare(b.name));
  for (const client of sortedClients) {
    const option = document.createElement("option");
    option.value = client.id;
    option.textContent = client.agentName ? `${client.name} · ${client.agentName}` : client.name;
    clientSelect.append(option);
  }

  clientSelect.value = selectedId || "";
}

function renderClientList() {
  clientList.innerHTML = "";

  if (!clients.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No saved clients yet. Add one above and it will appear here.";
    clientList.append(empty);
    return;
  }

  const sortedClients = [...clients].sort((a, b) => a.name.localeCompare(b.name));
  for (const client of sortedClients) {
    const item = clientTemplate.content.firstElementChild.cloneNode(true);
    item.querySelector("h3").textContent = client.name;
    item.querySelector(".client-email").textContent = client.email || "";
    item.querySelector(".client-agent").textContent = client.agentName
      ? `Agent: ${formatContact(client.agentName, client.agentPhone)}`
      : "";
    item.querySelector(".edit-client-button").addEventListener("click", () => editClient(client.id));
    clientList.append(item);
  }
}

function applySelectedClient() {
  const client = clients.find((item) => item.id === clientSelect.value);
  if (!client) {
    return;
  }

  clientNameInput.value = client.name || "";
  clientEmailInput.value = client.email || "";
  agentNameInput.value = client.agentName || "";
  agentPhoneInput.value = client.agentPhone || "";
  syncClientEmailToGuests();
  saveBookingDraft();
}

function resetClientForm() {
  clientForm.reset();
  directoryClientId.value = "";
  applyClientExamplePlaceholders();
  setClientMessage("");
  directoryClientName.focus();
}

function editClient(id) {
  const client = clients.find((item) => item.id === id);
  if (!client) {
    return;
  }

  directoryClientId.value = client.id;
  directoryClientName.value = client.name || "";
  directoryClientEmail.value = client.email || "";
  directoryAgentName.value = client.agentName || "";
  directoryAgentPhone.value = client.agentPhone || "";
  setClientMessage("Editing saved client.", "");
}

function getSelectedStartDate() {
  if (!dateInput.value || !timeInput.value) {
    return null;
  }

  const date = new Date(`${dateInput.value}T${timeInput.value}:00`);
  return Number.isNaN(date.getTime()) ? null : date;
}

function buildBookingPayload() {
  const start = getSelectedStartDate();
  if (!start) {
    return { error: "Choose a valid date and time." };
  }

  syncClientEmailToGuests();
  const data = Object.fromEntries(new FormData(bookingForm).entries());
  data.clientEmail = clientEmailInput.value.trim();
  data.guestEmails = guestEmailsInput.value;
  data.services = getSelectedServices();
  data.durationMinutes = Number(durationInput.value);
  data.startAt = start.toISOString();

  if (!data.services.length) {
    return { error: "Choose at least one service." };
  }

  return { data };
}

function getSyncLabel(booking) {
  if (booking.status === "cancelled") {
    return "cancelled";
  }

  const labels = {
    synced: "in Lark",
    failed: "Lark failed",
    attendees_failed: "guests failed",
    pending: "syncing",
    not_configured: "local only"
  };

  return labels[booking.larkStatus] || "local";
}

function renderBookings() {
  const upcoming = bookings
    .filter((booking) => new Date(booking.endAt).getTime() >= Date.now() || booking.status === "cancelled")
    .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));

  bookingList.innerHTML = "";

  if (!upcoming.length) {
    const empty = document.createElement("div");
    empty.className = "empty-state";
    empty.textContent = "No bookings yet. Create one above and it will appear here.";
    bookingList.append(empty);
    updateStats();
    return;
  }

  for (const booking of upcoming) {
    const start = new Date(booking.startAt);
    const end = new Date(booking.endAt);
    const item = bookingTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle("cancelled", booking.status === "cancelled");

    item.querySelector(".month").textContent = monthFormatter.format(start);
    item.querySelector(".day").textContent = dayFormatter.format(start);
    item.querySelector("h3").textContent = booking.propertyAddress || getBookingServiceLabel(booking);
    item.querySelector(".booking-time").textContent = `${dateFormatter.format(start)}, ${timeFormatter.format(start)}-${timeFormatter.format(end)}`;
    item.querySelector(".booking-location").textContent = booking.locationAddress || booking.locationName || "";
    item.querySelector(".booking-client").textContent = [booking.clientName, getBookingServiceLabel(booking)].filter(Boolean).join(" · ");
    item.querySelector(".booking-contact").textContent = [
      booking.photographerName ? `Photographer: ${formatContact(booking.photographerName, booking.photographerPhone)}` : "",
      booking.agentName ? `Agent: ${formatContact(booking.agentName, booking.agentPhone)}` : ""
    ].filter(Boolean).join(" · ");
    item.querySelector(".booking-notes").textContent = booking.notes || "";

    const pill = item.querySelector(".sync-pill");
    pill.textContent = getSyncLabel(booking);
    pill.classList.toggle("synced", booking.larkStatus === "synced");
    pill.classList.toggle("failed", booking.larkStatus === "failed" || booking.larkStatus === "attendees_failed");
    pill.classList.toggle("not-configured", booking.larkStatus === "not_configured");
    if (booking.larkError || booking.larkAttendeeError) {
      pill.title = [booking.larkError, booking.larkAttendeeError].filter(Boolean).join(" ");
    }

    const cancelButton = item.querySelector(".cancel-button");
    cancelButton.disabled = booking.status === "cancelled";
    cancelButton.addEventListener("click", () => cancelBooking(booking.id));

    bookingList.append(item);
  }

  updateStats();
}

function updateStats() {
  const now = new Date();
  const startOfToday = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const endOfToday = startOfToday + 24 * 60 * 60 * 1000;
  const nextWeek = now.getTime() + 7 * 24 * 60 * 60 * 1000;
  const active = bookings.filter((booking) => booking.status !== "cancelled");

  todayCount.textContent = active.filter((booking) => {
    const start = new Date(booking.startAt).getTime();
    return start >= startOfToday && start < endOfToday;
  }).length;

  weekCount.textContent = active.filter((booking) => {
    const start = new Date(booking.startAt).getTime();
    return start >= now.getTime() && start < nextWeek;
  }).length;
}

async function loadStatus() {
  const { data: status } = await fetchJson("/api/status");

  larkDot.className = `status-dot ${status.larkConfigured ? "ready" : "offline"}`;
  larkTitle.textContent = status.larkConfigured ? "Lark connected" : "Lark not connected";
  larkDetail.textContent = status.larkConfigured
    ? `Calendar ${status.calendarId}, ${status.timezone}`
    : "Add your Lark app settings to enable calendar sync.";
}

async function loadBookings() {
  const { data } = await fetchJson("/api/bookings");
  bookings = data.bookings || [];
  renderBookings();
}

async function loadClients() {
  const cachedClients = getCachedClients();

  try {
    const { data } = await fetchJson("/api/clients");
    clients = mergeClientLists(cachedClients, data.clients || []);
  } catch {
    clients = cachedClients;
  }

  storeClientCache(clients);
  renderClientOptions();
  renderClientList();
}

async function saveDirectoryClient(event) {
  event.preventDefault();

  const payload = {
    id: directoryClientId.value || undefined,
    name: directoryClientName.value.trim(),
    email: directoryClientEmail.value.trim(),
    agentName: directoryAgentName.value.trim(),
    agentPhone: directoryAgentPhone.value.trim()
  };

  if (!payload.name) {
    setClientMessage("Enter the agency or client name before saving.", "error");
    return;
  }

  const submitButton = clientForm.querySelector("button[type='submit']");
  submitButton.disabled = true;
  submitButton.textContent = "Saving";
  setClientMessage("Saving client...");

  try {
    const { response, data: result } = await fetchJson("/api/clients", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (!response.ok) {
      setClientMessage((result.errors || ["Could not save client."]).join(" "), "error");
      return;
    }

    clients = mergeClientLists(getCachedClients(), result.clients || clients, [result.client]);
    storeClientCache(clients);
    renderClientOptions(result.client.id);
    renderClientList();
    directoryClientId.value = result.client.id;
    setClientMessage("Client saved for bookings.", "success");
  } catch {
    setClientMessage("Could not reach the client directory.", "error");
  } finally {
    submitButton.disabled = false;
    submitButton.textContent = "Save client";
  }
}

async function submitBooking(event) {
  event.preventDefault();

  const { data, error } = buildBookingPayload();
  if (error) {
    setMessage(error, "error");
    return;
  }

  setMessage("Creating booking...");
  bookingForm.querySelector("button[type='submit']").disabled = true;

  try {
    const { response, data: result } = await fetchJson("/api/bookings", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      setMessage((result.errors || ["Could not create booking."]).join(" "), "error");
      return;
    }

    bookings.push(result.booking);
    renderBookings();
    bookingForm.reset();
    clientSelect.value = "";
    syncedClientEmail = "";
    removeStoredJson(bookingDraftKey);
    setInitialDateTime();
    updateDurationForServices();
    updateInvitationSummary();
    setMessage(
      result.booking.larkStatus === "synced"
        ? "Booking created and sent to Lark."
        : "Booking created locally. Add Lark settings when ready.",
      "success"
    );
  } catch {
    setMessage("Could not reach the booking server.", "error");
  } finally {
    bookingForm.querySelector("button[type='submit']").disabled = false;
  }
}

function renderLarkPreview(preview) {
  const start = new Date(preview.startAt);
  const end = new Date(preview.endAt);

  previewTitle.textContent = preview.title;
  previewTime.textContent = `${dateFormatter.format(start)}, ${timeFormatter.format(start)}-${timeFormatter.format(end)} (${preview.timezone})`;
  previewLocation.textContent = preview.location.address || preview.location.name || "";
  previewGuests.textContent = preview.guestEmails.length ? preview.guestEmails.join(", ") : "No guests";
  previewDescription.textContent = preview.description;
  larkPreview.hidden = false;
}

async function previewLarkEvent() {
  const { data, error } = buildBookingPayload();
  if (error) {
    setMessage(error, "error");
    return;
  }

  previewLarkButton.disabled = true;
  previewLarkButton.textContent = "Previewing";
  setMessage("Building Lark preview...");

  try {
    const { response, data: result } = await fetchJson("/api/lark/preview", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(data)
    });

    if (!response.ok) {
      setMessage((result.errors || ["Could not build preview."]).join(" "), "error");
      return;
    }

    renderLarkPreview(result.preview);
    setMessage("Lark preview generated. Nothing was sent.", "success");
  } catch {
    setMessage("Could not reach the preview service.", "error");
  } finally {
    previewLarkButton.disabled = false;
    previewLarkButton.textContent = "Preview Lark event";
  }
}

async function cancelBooking(id) {
  const { response, data: result } = await fetchJson(`/api/bookings/${id}/cancel`, { method: "POST" });

  if (response.ok) {
    bookings = bookings.map((booking) => (booking.id === id ? result.booking : booking));
    renderBookings();
  }
}

async function testLark() {
  testLarkButton.disabled = true;
  testLarkButton.textContent = "Testing";

  try {
    const { data: result } = await fetchJson("/api/lark/test", { method: "POST" });
    larkDot.className = `status-dot ${result.ok ? "ready" : "offline"}`;
    larkTitle.textContent = result.ok ? "Lark test passed" : "Lark test failed";
    larkDetail.textContent = result.message;
  } catch {
    larkDot.className = "status-dot offline";
    larkTitle.textContent = "Lark test failed";
    larkDetail.textContent = "The booking server could not complete the test.";
  } finally {
    testLarkButton.disabled = false;
    testLarkButton.textContent = "Test Lark";
  }
}

async function logout() {
  logoutButton.disabled = true;

  try {
    await fetch("/api/logout", { method: "POST" });
  } finally {
    window.location.href = "/login";
  }
}

bookingForm.addEventListener("submit", submitBooking);
bookingForm.addEventListener("input", saveBookingDraft);
bookingForm.addEventListener("change", saveBookingDraft);
clientForm.addEventListener("submit", saveDirectoryClient);
refreshButton.addEventListener("click", loadBookings);
testLarkButton.addEventListener("click", testLark);
previewLarkButton.addEventListener("click", previewLarkEvent);
logoutButton.addEventListener("click", logout);
clientSelect.addEventListener("change", applySelectedClient);
clientEmailInput.addEventListener("input", syncClientEmailToGuests);
clientEmailInput.addEventListener("change", syncClientEmailToGuests);
clientEmailInput.addEventListener("blur", syncClientEmailToGuests);
guestEmailsInput.addEventListener("input", updateInvitationSummary);
newClientButton.addEventListener("click", resetClientForm);
for (const input of serviceInputs) {
  input.addEventListener("change", () => {
    updateDurationForServices();
    saveBookingDraft();
  });
}

for (const link of navLinks) {
  link.addEventListener("click", (event) => {
    event.preventDefault();
    setRoute(link.dataset.route);
  });
}

window.addEventListener("popstate", () => {
  setRoute(window.location.pathname, false);
});

setInitialDateTime();
updateDurationForServices();
applyClientExamplePlaceholders();
await Promise.all([loadStatus(), loadClients(), loadBookings()]);
restoreBookingDraft();
updateInvitationSummary();
setRoute(window.location.pathname, false);
