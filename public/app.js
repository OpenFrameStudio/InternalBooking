const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const workNoticeDismissedKey = 'openframe.workNoticeDismissed.v1';

const el = {
  bookingForm: $('#bookingForm'),
  bookingList: $('#bookingList'),
  bookingTemplate: $('#bookingTemplate'),
  clientTemplate: $('#clientTemplate'),
  photographerTemplate: $('#photographerTemplate'),
  invoiceTemplate: $('#invoiceTemplate'),
  wageTemplate: $('#wageTemplate'),
  bookingPage: $('#bookingPage'),
  clientsPage: $('#clientsPage'),
  photographersPage: $('#photographersPage'),
  invoicesPage: $('#invoicesPage'),
  wagesPage: $('#wagesPage'),
  propertyAddress: $('#propertyAddressInput'),
  formMessage: $('#formMessage'),
  clientMessage: $('#clientMessage'),
  photographerMessage: $('#photographerMessage'),
  date: $('#dateInput'),
  time: $('#timeInput'),
  duration: $('#durationInput'),
  serviceInputs: $$('input[name=services]'),
  clientDropdown: $('#clientDropdown'),
  clientDropdownButton: $('#clientDropdownButton'),
  clientDropdownLabel: $('#clientDropdownLabel'),
  clientDropdownMenu: $('#clientDropdownMenu'),
  clientSearch: $('#clientSearchInput'),
  clientSearchResults: $('#clientSearchResults'),
  clientSelect: $('#clientSelect'),
  clientName: $('#clientNameInput'),
  clientEmail: $('#clientEmailInput'),
  agentName: $('#agentNameInput'),
  agentPhone: $('#agentPhoneInput'),
  photographerDropdown: $('#photographerDropdown'),
  photographerDropdownButton: $('#photographerDropdownButton'),
  photographerDropdownLabel: $('#photographerDropdownLabel'),
  photographerDropdownMenu: $('#photographerDropdownMenu'),
  photographerSelect: $('#photographerSelect'),
  photographerName: $('#photographerNameInput'),
  photographerEmail: $('#photographerEmailInput'),
  photographerPhone: $('#photographerPhoneInput'),
  photographerGstIncluded: $('#photographerGstIncludedInput'),
  guestEmails: $('#guestEmailsInput'),
  invitationEmails: $('#invitationEmails'),
  clientForm: $('#clientForm'),
  clientList: $('#clientList'),
  directoryClientId: $('#directoryClientId'),
  directoryClientName: $('#directoryClientName'),
  directoryClientEmail: $('#directoryClientEmail'),
  directoryAgentName: $('#directoryAgentName'),
  directoryAgentPhone: $('#directoryAgentPhone'),
  photographerForm: $('#photographerForm'),
  photographerList: $('#photographerList'),
  directoryPhotographerId: $('#directoryPhotographerId'),
  directoryPhotographerName: $('#directoryPhotographerName'),
  directoryPhotographerEmail: $('#directoryPhotographerEmail'),
  directoryPhotographerPhone: $('#directoryPhotographerPhone'),
  directoryPhotographerGstIncluded: $('#directoryPhotographerGstIncluded'),
  invoiceList: $('#invoiceList'),
  invoiceMessage: $('#invoiceMessage'),
  invoiceDraftCount: $('#invoiceDraftCount'),
  invoicePaidCount: $('#invoicePaidCount'),
  invoiceTotalValue: $('#invoiceTotalValue'),
  invoiceStatusFilter: $('#invoiceStatusFilter'),
  invoiceClientFilter: $('#invoiceClientFilter'),
  syncInvoicesButton: $('#syncInvoicesButton'),
  refreshInvoicesButton: $('#refreshInvoicesButton'),
  wageList: $('#wageList'),
  wageMessage: $('#wageMessage'),
  wageDraftCount: $('#wageDraftCount'),
  wagePaidCount: $('#wagePaidCount'),
  wageTotalValue: $('#wageTotalValue'),
  syncWagesButton: $('#syncWagesButton'),
  refreshWagesButton: $('#refreshWagesButton'),
  invoicePrintSheet: $('#invoicePrintSheet'),
  refreshButton: $('#refreshButton'),
  testLarkButton: $('#testLarkButton'),
  previewLarkButton: $('#previewLarkButton'),
  submitButton: $('#bookingSubmitButton'),
  cancelEditButton: $('#cancelEditButton'),
  changePasswordButton: $('#changePasswordButton'),
  logoutButton: $('#logoutButton'),
  passwordDialog: $('#passwordDialog'),
  passwordForm: $('#passwordForm'),
  currentPassword: $('#currentPasswordInput'),
  newPassword: $('#newPasswordInput'),
  confirmPassword: $('#confirmPasswordInput'),
  passwordMessage: $('#passwordMessage'),
  cancelPasswordButton: $('#cancelPasswordButton'),
  savePasswordButton: $('#savePasswordButton'),
  newClientButton: $('#newClientButton'),
  newPhotographerButton: $('#newPhotographerButton'),
  larkPreview: $('#larkPreview'),
  previewTitle: $('#previewTitle'),
  previewTime: $('#previewTime'),
  previewLocation: $('#previewLocation'),
  previewGuests: $('#previewGuests'),
  previewDescription: $('#previewDescription'),
  larkDot: $('#larkDot'),
  larkTitle: $('#larkTitle'),
  larkDetail: $('#larkDetail'),
  mainAssignmentNotice: $('#mainAssignmentNotice'),
  mainAssignmentNoticeButton: $('#mainAssignmentNoticeButton'),
  mainAssignmentNoticeDetail: $('#mainAssignmentNoticeDetail'),
  mainAssignmentNoticeTitle: $('#mainAssignmentNoticeTitle'),
  todayCount: $('#todayCount'),
  weekCount: $('#weekCount'),
  modeEyebrow: $('#bookingModeEyebrow'),
  modeTitle: $('#bookingModeTitle'),
  appLinks: $$('[data-app-link]')
};

const state = {
  bookings: [],
  clients: [],
  photographers: [],
  invoices: [],
  wages: [],
  workAssignments: [],
  bookingForm: {
    selectedClientId: '',
    selectedPhotographerId: '',
    clientDropdownOpen: false,
    photographerDropdownOpen: false,
    clientSearchQuery: '',
    clientSearchOpen: false
  },
  invoiceFilters: {
    status: 'all',
    query: ''
  },
  editingBookingId: '',
  restoringDraft: false,
  syncedClientEmail: [],
  syncedPhotographerEmail: [],
  workNoticeDismissedSignature: readWorkNoticeDismissedSignature(),
  user: null
};

const storageKeys = {
  draft: 'openframe.bookingDraft.v2',
  bookings: 'openframe.bookings.v1',
  clients: 'openframe.clients.v1',
  photographers: 'openframe.photographers.v1',
  invoices: 'openframe.invoices.v1',
  wages: 'openframe.wages.v1',
  addresses: 'openframe.addresses.v1'
};

const clientExamples = [
  ['Stonebridge Collective', 'ava@stonebridge.example', 'Ava Brooks', '0412 684 209'],
  ['Northline Estates', 'leo@northline.example', 'Leo Tran', '0426 318 774'],
  ['Cedar & Coast Realty', 'nina@cedarcoast.example', 'Nina Patel', '0437 902 146'],
  ['Blue Lantern Property', 'eli@bluelantern.example', 'Eli Morris', '0408 571 663'],
  ['Park Row Agency', 'sophie@parkrow.example', 'Sophie Lin', '0419 246 805']
];

const photographerExamples = [
  ['Maya Hart', 'maya@openframe.studio', '0416 208 771'],
  ['Ethan Vale', 'ethan@openframe.studio', '0428 615 904'],
  ['Lina Park', 'lina@openframe.studio', '0435 772 018'],
  ['Noah Kim', 'noah@openframe.studio', '0407 394 226'],
  ['Iris Chen', 'iris@openframe.studio', '0419 830 642']
];

const bookingTimeZone = 'Australia/Sydney';
const dateFormatter = new Intl.DateTimeFormat('en-AU', { weekday: 'short', month: 'short', day: 'numeric', timeZone: bookingTimeZone });
const timeFormatter = new Intl.DateTimeFormat('en-AU', { hour: 'numeric', minute: '2-digit', timeZone: bookingTimeZone });
const monthFormatter = new Intl.DateTimeFormat('en-AU', { month: 'short', timeZone: bookingTimeZone });
const dayFormatter = new Intl.DateTimeFormat('en-AU', { day: '2-digit', timeZone: bookingTimeZone });
const invoiceDateFormatter = new Intl.DateTimeFormat('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
const currencyFormatter = new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD' });
const bookingUpcomingGraceMs = 24 * 60 * 60 * 1000;
const timeZonePartsFormatter = new Intl.DateTimeFormat('en-AU', {
  timeZone: bookingTimeZone,
  year: 'numeric',
  month: '2-digit',
  day: '2-digit',
  hour: '2-digit',
  minute: '2-digit',
  second: '2-digit',
  hourCycle: 'h23'
});

function isTruthy(value) {
  return value === true || ['1', 'true', 'yes', 'on'].includes(String(value || '').trim().toLowerCase());
}

function wageGstLabel(value) {
  return isTruthy(value) ? 'GST included' : 'No GST';
}

function readStored(key, fallback) {
  try {
    const value = JSON.parse(localStorage.getItem(key) || 'null');
    return value ?? fallback;
  } catch {
    return fallback;
  }
}

function writeStored(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Local storage is a convenience only.
  }
}

function removeStored(key) {
  try {
    localStorage.removeItem(key);
  } catch {
    // Ignore browsers that block local storage.
  }
}

function readWorkNoticeDismissedSignature() {
  try {
    return localStorage.getItem(workNoticeDismissedKey) || '';
  } catch {
    return '';
  }
}

function writeWorkNoticeDismissedSignature(signature) {
  try {
    if (signature) localStorage.setItem(workNoticeDismissedKey, signature);
    else localStorage.removeItem(workNoticeDismissedKey);
  } catch {
    // This only hides the small in-app banner.
  }
}

function cacheBookings() {
  writeStored(storageKeys.bookings, state.bookings);
}

function cacheInvoices() {
  if (userCanAccess('invoices')) writeStored(storageKeys.invoices, state.invoices);
}

function cacheWages() {
  if (userCanAccess('wages')) writeStored(storageKeys.wages, state.wages);
}

function hydrateCachedData() {
  const cachedBookings = readStored(storageKeys.bookings, []).filter((booking) => booking?.id);
  if (cachedBookings.length) {
    state.bookings = cachedBookings;
  }
  renderBookings();

  const cachedClients = readStored(storageKeys.clients, []).filter((client) => client?.name);
  if (cachedClients.length) {
    state.clients = mergeDirectoryRecords(state.clients, cachedClients);
  }
  renderClientOptions();
  renderClientList();

  const cachedPhotographers = readStored(storageKeys.photographers, []).filter((photographer) => photographer?.name);
  if (cachedPhotographers.length) {
    state.photographers = mergeDirectoryRecords(state.photographers, cachedPhotographers);
  }
  renderPhotographerOptions();
  renderPhotographerList();

  const cachedInvoices = userCanAccess('invoices')
    ? readStored(storageKeys.invoices, []).filter((invoice) => invoice?.id)
    : [];
  if (cachedInvoices.length) {
    state.invoices = cachedInvoices;
  }
  if (userCanAccess('invoices')) renderInvoices();

  const cachedWages = userCanAccess('wages')
    ? readStored(storageKeys.wages, []).filter((wage) => wage?.id)
    : [];
  if (cachedWages.length) {
    state.wages = cachedWages;
  }
  if (userCanAccess('wages')) renderWages();
}

function refreshLiveData() {
  loadStatus().catch(() => {
    el.larkDot.className = 'status-dot offline';
    el.larkTitle.textContent = 'Connection delayed';
    el.larkDetail.textContent = 'Showing saved data while the live calendar catches up.';
  });

  Promise.allSettled([
    loadClients(),
    loadPhotographers(),
    loadBookings(),
    loadInvoices(),
    loadWages(),
    loadWorkAssignments()
  ]);
}

async function fetchJson(url, options) {
  const response = await fetch(url, options);
  const data = await response.json().catch(() => ({}));
  if (response.status === 401) {
    window.location.href = '/login';
    throw new Error('Log in first.');
  }
  return { response, data };
}

function setMessage(target, message, tone = '') {
  target.textContent = message;
  target.className = tone;
}

async function loadSession() {
  const { data } = await fetchJson('/api/session');
  state.user = data.user || null;
  applyAppAccess();
}

function userCanAccess(app) {
  return Boolean(state.user?.apps?.includes(app));
}

function userHasPermission(permission) {
  return Boolean(state.user?.permissions?.includes(permission));
}

function applyAppAccess() {
  el.appLinks.forEach((link) => {
    link.hidden = !userCanAccess(link.dataset.appLink);
  });
  el.testLarkButton.hidden = !userHasPermission('manage_bookings');

  if (!userCanAccess('clients') && window.location.pathname === '/clients') {
    setRoute('/bookings');
  }

  if (!userCanAccess('photographers') && window.location.pathname === '/photographers') {
    setRoute('/bookings');
  }

  if (!userCanAccess('invoices') && window.location.pathname === '/invoices') {
    setRoute('/bookings');
  }

  if (!userCanAccess('wages') && window.location.pathname === '/wages') {
    setRoute('/bookings');
  }
}

function randomItem(items) {
  return items[Math.floor(Math.random() * items.length)];
}

function applyExamplePlaceholders() {
  const [agency, email, agent, phone] = randomItem(clientExamples);
  el.clientName.placeholder = agency;
  el.clientEmail.placeholder = email;
  el.agentName.placeholder = agent;
  el.agentPhone.placeholder = phone;
  el.directoryClientName.placeholder = agency;
  el.directoryClientEmail.placeholder = email;
  el.directoryAgentName.placeholder = agent;
  el.directoryAgentPhone.placeholder = phone;

  const [name, photographerEmail, photographerPhone] = randomItem(photographerExamples);
  el.photographerName.placeholder = name;
  el.photographerEmail.placeholder = photographerEmail;
  el.photographerPhone.placeholder = photographerPhone;
  el.directoryPhotographerName.placeholder = name;
  el.directoryPhotographerEmail.placeholder = photographerEmail;
  el.directoryPhotographerPhone.placeholder = photographerPhone;
}

function normalizeAddress(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function uniqueAddresses(addresses) {
  const seen = new Set();
  const unique = [];
  for (const address of addresses.map(normalizeAddress)) {
    const key = address.toLowerCase();
    if (!address || seen.has(key)) continue;
    seen.add(key);
    unique.push(address);
  }
  return unique;
}

function rememberAddress(address) {
  const cleanAddress = normalizeAddress(address);
  if (!cleanAddress) return;
  const stored = readStored(storageKeys.addresses, []);
  writeStored(storageKeys.addresses, uniqueAddresses([cleanAddress, ...stored]).slice(0, 30));
}

function parseEmails(value) {
  const matches = String(value || '').match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/gi);
  if (matches?.length) {
    return matches.map((email) => email.trim());
  }
  return String(value || '').split(/[\s,;]+/).map((email) => email.trim()).filter(Boolean);
}

function uniqueEmails(emails) {
  const seen = new Set();
  return emails.map((email) => String(email || '').trim()).filter((email) => {
    const key = email.toLowerCase();
    if (!email || seen.has(key)) {
      return false;
    }
    seen.add(key);
    return true;
  });
}

function isEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(value);
}

function formatContact(name, phone) {
  if (name && phone) return `${name} (${phone})`;
  return name || phone || '';
}

function emailsOverlap(left, right) {
  const rightEmails = new Set(parseEmails(right).map((email) => email.toLowerCase()));
  return parseEmails(left).some((email) => rightEmails.has(email.toLowerCase()));
}

function selectedServices() {
  return el.serviceInputs.filter((input) => input.checked).map((input) => ({ name: input.value }));
}

function updateDurationForServices() {
  const count = selectedServices().length;
  if (count >= 3) el.duration.value = '90';
  else if (count === 2) el.duration.value = '60';
  else if (count === 1) el.duration.value = '45';
}

function clientIdentityKey(client) {
  return [
    client?.name,
    client?.agentName,
    client?.email,
    client?.agentPhone
  ]
    .map((value) => String(value || '').trim().toLowerCase())
    .join('|');
}

function recordIdentityKey(item) {
  return item?.agentName !== undefined || item?.agentPhone !== undefined
    ? clientIdentityKey(item)
    : String(item?.name || '').trim().toLowerCase();
}

function mergeDirectoryRecords(...lists) {
  const merged = new Map();
  for (const list of lists) {
    for (const item of list || []) {
      const name = String(item?.name || '').trim();
      if (!name) continue;
      const key = item?.id || recordIdentityKey(item);
      const current = merged.get(key);
      const next = { ...current, ...item, name };
      if (!current || new Date(next.updatedAt || 0) >= new Date(current.updatedAt || 0)) {
        merged.set(key, next);
      }
    }
  }
  return [...merged.values()].sort((a, b) => a.name.localeCompare(b.name));
}

function updateBookingFormState(patch, { renderSelections = true } = {}) {
  state.bookingForm = { ...state.bookingForm, ...patch };
  if (renderSelections) {
    renderClientOptions();
    renderPhotographerOptions();
  }
}

function selectedClient() {
  return state.clients.find((item) => item.id === state.bookingForm.selectedClientId) || null;
}

function selectedPhotographer() {
  return state.photographers.find((item) => item.id === state.bookingForm.selectedPhotographerId) || null;
}

function reconcileDirectorySelections() {
  if (state.bookingForm.selectedClientId && !selectedClient()) {
    state.bookingForm.selectedClientId = '';
  }

  if (state.bookingForm.selectedPhotographerId && !selectedPhotographer()) {
    state.bookingForm.selectedPhotographerId = '';
  }
}

function bookingServiceLabel(booking) {
  if (Array.isArray(booking.services) && booking.services.length) {
    return booking.services.map((service) => service.name).join(' + ');
  }
  return booking.service || 'Booking';
}

function workPriorityWeight(priority) {
  return { high: 3, normal: 2, low: 1 }[priority] || 2;
}

function parseDateParts(value) {
  const [year, month, day] = String(value || '').split('-').map(Number);
  return Number.isInteger(year) && Number.isInteger(month) && Number.isInteger(day)
    ? { year, month, day }
    : null;
}

function parseTimeParts(value) {
  const [hour, minute] = String(value || '').split(':').map(Number);
  return Number.isInteger(hour) && Number.isInteger(minute)
    ? { hour, minute }
    : null;
}

function parseDateValue(value) {
  const parts = parseDateParts(value);
  return parts ? new Date(parts.year, parts.month - 1, parts.day) : new Date(NaN);
}

function isTodayDateValue(value) {
  return value === toDateValue(new Date());
}

function isOverdueWorkAssignment(assignment) {
  return assignment.status !== 'done' && parseDateValue(assignment.dueDate) < parseDateValue(toDateValue(new Date()));
}

function openWorkAssignments() {
  return [...state.workAssignments]
    .filter((assignment) => assignment.status !== 'done')
    .sort((a, b) => {
      const overdueSort = Number(isOverdueWorkAssignment(b)) - Number(isOverdueWorkAssignment(a));
      if (overdueSort) return overdueSort;
      if (a.dueDate !== b.dueDate) return String(a.dueDate || '').localeCompare(String(b.dueDate || ''));
      return workPriorityWeight(b.priority) - workPriorityWeight(a.priority);
    });
}

function workNoticeSignature(assignments) {
  return assignments
    .map((assignment) => `${assignment.id}:${assignment.updatedAt || assignment.createdAt || ''}`)
    .sort()
    .join('|');
}

function getAustraliaTimeParts(date) {
  const parts = Object.fromEntries(timeZonePartsFormatter.formatToParts(date).map((part) => [part.type, part.value]));
  return {
    year: Number(parts.year),
    month: Number(parts.month),
    day: Number(parts.day),
    hour: Number(parts.hour),
    minute: Number(parts.minute),
    second: Number(parts.second)
  };
}

function dateValueFromParts(parts) {
  return `${parts.year}-${String(parts.month).padStart(2, '0')}-${String(parts.day).padStart(2, '0')}`;
}

function timeValueFromMinutes(minutes) {
  return `${String(Math.floor(minutes / 60)).padStart(2, '0')}:${String(minutes % 60).padStart(2, '0')}`;
}

function addDaysToDateValue(value, days) {
  const parts = parseDateParts(value);
  if (!parts) return value;
  const date = new Date(Date.UTC(parts.year, parts.month - 1, parts.day + days, 12, 0, 0));
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}-${String(date.getUTCDate()).padStart(2, '0')}`;
}

function toDateValue(date) {
  return dateValueFromParts(getAustraliaTimeParts(date));
}

function toTimeValue(date) {
  const parts = getAustraliaTimeParts(date);
  return `${String(parts.hour).padStart(2, '0')}:${String(parts.minute).padStart(2, '0')}`;
}

function australiaDateTimeToUtcDate(dateValue, timeValue) {
  const dateParts = parseDateParts(dateValue);
  const timeParts = parseTimeParts(timeValue);
  if (!dateParts || !timeParts) return null;

  const targetUtc = Date.UTC(dateParts.year, dateParts.month - 1, dateParts.day, timeParts.hour, timeParts.minute, 0);
  let utc = targetUtc;
  for (let index = 0; index < 3; index += 1) {
    const parts = getAustraliaTimeParts(new Date(utc));
    const displayedUtc = Date.UTC(parts.year, parts.month - 1, parts.day, parts.hour, parts.minute, parts.second || 0);
    utc -= displayedUtc - targetUtc;
  }

  const result = new Date(utc);
  const check = getAustraliaTimeParts(result);
  const matches = check.year === dateParts.year
    && check.month === dateParts.month
    && check.day === dateParts.day
    && check.hour === timeParts.hour
    && check.minute === timeParts.minute;

  return matches ? result : null;
}

function setInitialDateTime() {
  const nowParts = getAustraliaTimeParts(new Date());
  const today = dateValueFromParts(nowParts);
  let dateValue = today;
  let minutes = Math.ceil((nowParts.hour * 60 + nowParts.minute) / 15) * 15;

  if (minutes < 8 * 60) {
    minutes = 9 * 60;
  } else if (minutes >= 18 * 60) {
    dateValue = addDaysToDateValue(dateValue, 1);
    minutes = 9 * 60;
  }

  el.date.min = today;
  el.date.value = dateValue;
  el.time.value = timeValueFromMinutes(minutes);
}

function ensureDurationOption(minutes) {
  if ([...el.duration.options].some((option) => option.value === String(minutes))) return;
  const option = document.createElement('option');
  option.value = String(minutes);
  option.textContent = `${minutes} min`;
  el.duration.append(option);
}

function updateInvitationSummary() {
  const emails = uniqueEmails([
    ...parseEmails(el.clientEmail.value),
    ...parseEmails(el.photographerEmail.value),
    ...parseEmails(el.guestEmails.value)
  ]).filter(isEmail);
  el.invitationEmails.textContent = emails.length ? emails.join(', ') : 'No invitation emails yet.';
}

function syncManagedGuestEmail(input, key) {
  const nextEmails = uniqueEmails(parseEmails(input.value)).filter(isEmail);
  const nextKeys = new Set(nextEmails.map((email) => email.toLowerCase()));
  const previousEmails = Array.isArray(state[key]) ? state[key] : parseEmails(state[key]);
  let emails = parseEmails(el.guestEmails.value);

  for (const previousEmail of previousEmails) {
    if (!nextKeys.has(previousEmail.toLowerCase())) {
      emails = emails.filter((email) => email.toLowerCase() !== previousEmail.toLowerCase());
    }
  }

  emails = uniqueEmails([...nextEmails, ...emails]);
  state[key] = nextEmails;
  el.guestEmails.value = emails.join(', ');
  updateInvitationSummary();
  saveDraft();
}

function syncAutoEmails() {
  syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail');
  syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail');
}

function renderOptions(select, items, emptyLabel, selectedId = '') {
  select.innerHTML = '';
  select.append(new Option(emptyLabel, ''));
  for (const item of [...items].sort((a, b) => a.name.localeCompare(b.name))) {
    const subtitle = item.agentName || item.phone || '';
    select.append(new Option(subtitle ? `${item.name} - ${subtitle}` : item.name, item.id));
  }
  select.value = selectedId || '';
}

function clientOptionLabel(client) {
  return [client.name, client.agentName].filter(Boolean).join(' - ');
}

function selectedClientLabel() {
  const client = selectedClient();
  return client ? clientOptionLabel(client) : 'New client / enter manually';
}

function selectedPhotographerLabel() {
  const photographer = selectedPhotographer();
  return photographer ? photographer.name : 'New photographer';
}

function clientSearchText(client) {
  return [
    client.name,
    client.agentName,
    client.email,
    client.agentPhone
  ].join(' ').toLowerCase();
}

function filteredClients() {
  const query = normalizeAddress(state.bookingForm.clientSearchQuery).toLowerCase();
  if (!query) return state.clients;
  const terms = query.split(/\s+/).filter(Boolean);
  return state.clients.filter((client) => terms.every((term) => clientSearchText(client).includes(term)));
}

function renderClientOptions() {
  reconcileDirectorySelections();
  renderOptions(el.clientSelect, state.clients, 'New client / enter manually', state.bookingForm.selectedClientId);
  renderClientDropdown();
  renderClientSearchResults();
}

function setClientDropdownOpen(isOpen) {
  updateBookingFormState({ clientDropdownOpen: Boolean(isOpen) });
}

function closeClientDropdown() {
  setClientDropdownOpen(false);
}

function selectClient(clientId) {
  updateBookingFormState({
    selectedClientId: clientId || '',
    clientDropdownOpen: false
  }, { renderSelections: false });
  applySelectedClient();
  renderClientOptions();
}

function renderClientDropdown() {
  el.clientDropdownLabel.textContent = selectedClientLabel();
  el.clientDropdownMenu.innerHTML = '';
  el.clientDropdownMenu.hidden = !state.bookingForm.clientDropdownOpen;
  el.clientDropdownButton.setAttribute('aria-expanded', String(state.bookingForm.clientDropdownOpen));
  el.clientDropdown.classList.toggle('open', state.bookingForm.clientDropdownOpen);

  const options = [
    { id: '', title: 'New client / enter manually', meta: 'Clear the saved client fields' },
    ...[...state.clients]
      .sort((a, b) => clientOptionLabel(a).localeCompare(clientOptionLabel(b)))
      .map((client) => ({
        id: client.id,
        title: clientOptionLabel(client) || client.name,
        meta: [client.email, client.agentPhone].filter(Boolean).join(' · ')
      }))
  ];

  for (const option of options) {
    const button = document.createElement('button');
    const isSelected = option.id === state.bookingForm.selectedClientId;
    button.className = `client-dropdown-option${isSelected ? ' selected' : ''}`;
    button.type = 'button';
    button.role = 'option';
    button.setAttribute('aria-selected', String(isSelected));
    button.dataset.clientId = option.id;

    const title = document.createElement('strong');
    title.textContent = option.title;
    const meta = document.createElement('span');
    meta.textContent = option.meta;
    button.append(title, meta);
    button.addEventListener('click', () => selectClient(option.id));
    el.clientDropdownMenu.append(button);
  }
}

function renderPhotographerOptions() {
  reconcileDirectorySelections();
  renderOptions(el.photographerSelect, state.photographers, 'New photographer', state.bookingForm.selectedPhotographerId);
  renderPhotographerDropdown();
}

function setPhotographerDropdownOpen(isOpen) {
  updateBookingFormState({ photographerDropdownOpen: Boolean(isOpen) });
}

function closePhotographerDropdown() {
  setPhotographerDropdownOpen(false);
}

function selectPhotographer(photographerId) {
  updateBookingFormState({
    selectedPhotographerId: photographerId || '',
    photographerDropdownOpen: false
  }, { renderSelections: false });
  applySelectedPhotographer();
  renderPhotographerOptions();
}

function renderPhotographerDropdown() {
  el.photographerDropdownLabel.textContent = selectedPhotographerLabel();
  el.photographerDropdownMenu.innerHTML = '';
  el.photographerDropdownMenu.hidden = !state.bookingForm.photographerDropdownOpen;
  el.photographerDropdownButton.setAttribute('aria-expanded', String(state.bookingForm.photographerDropdownOpen));
  el.photographerDropdown.classList.toggle('open', state.bookingForm.photographerDropdownOpen);

  const options = [
    { id: '', title: 'New photographer', meta: 'Enter photographer details manually' },
    ...[...state.photographers]
      .sort((a, b) => a.name.localeCompare(b.name))
      .map((photographer) => ({
        id: photographer.id,
        title: photographer.name,
        meta: [photographer.email, photographer.phone, wageGstLabel(photographer.gstIncluded)].filter(Boolean).join(' · ')
      }))
  ];

  for (const option of options) {
    const button = document.createElement('button');
    const isSelected = option.id === state.bookingForm.selectedPhotographerId;
    button.className = `client-dropdown-option${isSelected ? ' selected' : ''}`;
    button.type = 'button';
    button.role = 'option';
    button.setAttribute('aria-selected', String(isSelected));
    button.dataset.photographerId = option.id;

    const title = document.createElement('strong');
    title.textContent = option.title;
    const meta = document.createElement('span');
    meta.textContent = option.meta;
    button.append(title, meta);
    button.addEventListener('click', () => selectPhotographer(option.id));
    el.photographerDropdownMenu.append(button);
  }
}

function renderClientSearchResults() {
  if (el.clientSearch.value !== state.bookingForm.clientSearchQuery) {
    el.clientSearch.value = state.bookingForm.clientSearchQuery;
  }

  const query = normalizeAddress(state.bookingForm.clientSearchQuery);
  el.clientSearchResults.innerHTML = '';

  if (!state.bookingForm.clientSearchOpen || !query) {
    el.clientSearchResults.hidden = true;
    return;
  }

  const matches = filteredClients().slice(0, 6);
  if (!matches.length) {
    const empty = document.createElement('div');
    empty.className = 'client-search-empty';
    empty.textContent = 'No saved clients match this search.';
    el.clientSearchResults.append(empty);
    el.clientSearchResults.hidden = false;
    return;
  }

  for (const client of matches) {
    const button = document.createElement('button');
    button.className = 'client-search-result';
    button.type = 'button';

    const title = document.createElement('strong');
    title.textContent = clientOptionLabel(client) || client.name;
    const meta = document.createElement('span');
    meta.textContent = [client.email, client.agentPhone].filter(Boolean).join(' · ');

    button.append(title, meta);
    button.addEventListener('click', () => {
      selectClient(client.id);
    });
    el.clientSearchResults.append(button);
  }

  el.clientSearchResults.hidden = false;
}

function renderClientList() {
  el.clientList.innerHTML = '';
  if (!state.clients.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = 'No saved clients yet. Add one above and it will appear here.';
    el.clientList.append(empty);
    return;
  }
  for (const client of state.clients) {
    const item = el.clientTemplate.content.firstElementChild.cloneNode(true);
    item.querySelector('h3').textContent = client.name;
    item.querySelector('.client-email').textContent = client.email || '';
    item.querySelector('.client-agent').textContent = client.agentName ? `Agent: ${formatContact(client.agentName, client.agentPhone)}` : '';
    item.querySelector('.edit-client-button').addEventListener('click', () => editClient(client.id));
    item.querySelector('.delete-client-button').addEventListener('click', () => deleteClient(client.id));
    el.clientList.append(item);
  }
}

function renderPhotographerList() {
  el.photographerList.innerHTML = '';
  if (!state.photographers.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = 'No saved photographers yet. Add one above and it will appear here.';
    el.photographerList.append(empty);
    return;
  }
  for (const photographer of state.photographers) {
    const item = el.photographerTemplate.content.firstElementChild.cloneNode(true);
    item.querySelector('h3').textContent = photographer.name;
    item.querySelector('.photographer-email').textContent = photographer.email || '';
    item.querySelector('.photographer-phone').textContent = [photographer.phone ? `Phone: ${photographer.phone}` : '', wageGstLabel(photographer.gstIncluded)].filter(Boolean).join(' · ');
    item.querySelector('.edit-photographer-button').addEventListener('click', () => editPhotographer(photographer.id));
    item.querySelector('.delete-photographer-button').addEventListener('click', () => deletePhotographer(photographer.id));
    el.photographerList.append(item);
  }
}

function applySelectedClient() {
  const client = selectedClient();
  if (!client) {
    state.bookingForm.clientSearchQuery = '';
    state.bookingForm.clientSearchOpen = false;
    el.clientSearchResults.hidden = true;
    el.clientName.value = '';
    el.clientEmail.value = '';
    el.agentName.value = '';
    el.agentPhone.value = '';
    syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail');
    return;
  }
  state.bookingForm.clientSearchQuery = clientOptionLabel(client);
  state.bookingForm.clientSearchOpen = false;
  el.clientSearchResults.hidden = true;
  el.clientName.value = client.name || '';
  el.clientEmail.value = client.email || '';
  el.agentName.value = client.agentName || '';
  el.agentPhone.value = client.agentPhone || '';
  syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail');
  renderClientDropdown();
}

function applySelectedPhotographer() {
  const photographer = selectedPhotographer();
  if (!photographer) {
    renderPhotographerDropdown();
    return;
  }
  el.photographerName.value = photographer.name || '';
  el.photographerEmail.value = photographer.email || '';
  el.photographerPhone.value = photographer.phone || '';
  el.photographerGstIncluded.checked = isTruthy(photographer.gstIncluded);
  syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail');
  renderPhotographerDropdown();
}

function getDraft() {
  const data = Object.fromEntries(new FormData(el.bookingForm).entries());
  return {
    propertyAddress: el.propertyAddress.value || data.propertyAddress || '',
    selectedClientId: state.bookingForm.selectedClientId,
    clientName: el.clientName.value,
    clientEmail: el.clientEmail.value,
    agentName: el.agentName.value,
    agentPhone: el.agentPhone.value,
    services: selectedServices().map((service) => service.name),
    date: el.date.value,
    time: el.time.value,
    durationMinutes: el.duration.value,
    selectedPhotographerId: state.bookingForm.selectedPhotographerId,
    photographerName: el.photographerName.value,
    photographerEmail: el.photographerEmail.value,
    photographerPhone: el.photographerPhone.value,
    photographerGstIncluded: el.photographerGstIncluded.checked,
    guestEmails: el.guestEmails.value,
    notes: data.notes || ''
  };
}

function saveDraft() {
  removeStored(storageKeys.draft);
}

function restoreDraft() {
  removeStored(storageKeys.draft);
  return false;
}

function setBookingMode(booking = null) {
  state.editingBookingId = booking?.id || '';
  el.modeEyebrow.textContent = state.editingBookingId ? 'Edit booking' : 'New booking';
  el.modeTitle.textContent = state.editingBookingId ? 'Update property booking' : 'Create a property booking';
  el.submitButton.textContent = state.editingBookingId ? 'Update booking' : 'Create booking';
  el.cancelEditButton.hidden = !state.editingBookingId;
}

function setSelectedServices(booking) {
  const selected = new Set(
    Array.isArray(booking.services) && booking.services.length
      ? booking.services.map((service) => service.name)
      : String(booking.service || '').split('+').map((service) => service.trim())
  );
  el.serviceInputs.forEach((input) => {
    input.checked = selected.has(input.value);
  });
}

function fillBookingForm(booking) {
  state.restoringDraft = true;
  const matchingClient = state.clients.find((client) => {
    return client.name.toLowerCase() === String(booking.clientName || '').toLowerCase()
      || (client.email && booking.clientEmail && emailsOverlap(client.email, booking.clientEmail));
  });
  const matchingPhotographer = state.photographers.find((photographer) => {
    return photographer.name.toLowerCase() === String(booking.photographerName || '').toLowerCase()
      || (photographer.email && booking.photographerEmail && photographer.email.toLowerCase() === booking.photographerEmail.toLowerCase());
  });

  el.propertyAddress.value = booking.propertyAddress || '';
  updateBookingFormState({
    selectedClientId: matchingClient?.id || '',
    selectedPhotographerId: matchingPhotographer?.id || '',
    clientDropdownOpen: false,
    photographerDropdownOpen: false,
    clientSearchQuery: matchingClient ? clientOptionLabel(matchingClient) : '',
    clientSearchOpen: false
  }, { renderSelections: false });
  renderClientOptions();
  el.clientName.value = booking.clientName || '';
  el.clientEmail.value = booking.clientEmail || '';
  el.agentName.value = booking.agentName || '';
  el.agentPhone.value = booking.agentPhone || '';
  renderPhotographerOptions();
  el.photographerName.value = booking.photographerName || matchingPhotographer?.name || 'Barry';
  el.photographerEmail.value = booking.photographerEmail || matchingPhotographer?.email || '';
  el.photographerPhone.value = booking.photographerPhone || matchingPhotographer?.phone || '0403 007 853';
  el.photographerGstIncluded.checked = isTruthy(
    booking.photographerGstIncluded !== undefined
      ? booking.photographerGstIncluded
      : matchingPhotographer?.gstIncluded
  );
  setSelectedServices(booking);

  const start = new Date(booking.startAt);
  if (!Number.isNaN(start.getTime())) {
    el.date.value = toDateValue(start);
    el.time.value = toTimeValue(start);
  }
  ensureDurationOption(booking.durationMinutes || 60);
  el.duration.value = String(booking.durationMinutes || 60);
  el.guestEmails.value = Array.isArray(booking.guestEmails) ? booking.guestEmails.join(', ') : '';
  el.bookingForm.elements.notes.value = booking.notes || '';
  state.syncedClientEmail = uniqueEmails(parseEmails(el.clientEmail.value)).filter(isEmail);
  state.syncedPhotographerEmail = uniqueEmails(parseEmails(el.photographerEmail.value)).filter(isEmail);
  updateInvitationSummary();
  el.larkPreview.hidden = true;
  state.restoringDraft = false;
}

function resetBookingForm(message = '') {
  setBookingMode();
  el.bookingForm.reset();
  el.clientSearchResults.hidden = true;
  updateBookingFormState({
    selectedClientId: '',
    selectedPhotographerId: '',
    clientDropdownOpen: false,
    photographerDropdownOpen: false,
    clientSearchQuery: '',
    clientSearchOpen: false
  });
  state.syncedClientEmail = [];
  state.syncedPhotographerEmail = [];
  el.photographerGstIncluded.checked = false;
  setInitialDateTime();
  updateDurationForServices();
  applyExamplePlaceholders();
  updateInvitationSummary();
  el.larkPreview.hidden = true;
  setMessage(el.formMessage, message);
}

function setRoute(route, push = true) {
  const nextRoute = ['/clients', '/photographers', '/invoices', '/wages'].includes(route) ? route : '/bookings';
  el.bookingPage.hidden = nextRoute !== '/bookings';
  el.clientsPage.hidden = nextRoute !== '/clients';
  el.photographersPage.hidden = nextRoute !== '/photographers';
  el.invoicesPage.hidden = nextRoute !== '/invoices';
  el.wagesPage.hidden = nextRoute !== '/wages';
  $$('[data-route]').forEach((link) => link.classList.toggle('active', link.dataset.route === nextRoute));
  if (push && window.location.pathname !== nextRoute) history.pushState({ route: nextRoute }, '', nextRoute);
  if (nextRoute === '/clients') renderClientList();
  if (nextRoute === '/photographers') renderPhotographerList();
  if (nextRoute === '/invoices') renderInvoices();
  if (nextRoute === '/wages') renderWages();
}

function buildBookingPayload() {
  const start = el.date.value && el.time.value ? australiaDateTimeToUtcDate(el.date.value, el.time.value) : null;
  if (!start || Number.isNaN(start.getTime())) return { error: 'Choose a valid date and time.' };
  syncAutoEmails();
  const data = Object.fromEntries(new FormData(el.bookingForm).entries());
  data.clientEmail = el.clientEmail.value.trim();
  data.photographerName = el.photographerName.value.trim();
  data.photographerEmail = el.photographerEmail.value.trim();
  data.photographerPhone = el.photographerPhone.value.trim();
  data.photographerGstIncluded = el.photographerGstIncluded.checked;
  data.guestEmails = el.guestEmails.value;
  data.services = selectedServices();
  data.durationMinutes = Number(el.duration.value);
  data.date = el.date.value;
  data.time = el.time.value;
  data.timezone = bookingTimeZone;
  data.startAt = start.toISOString();
  if (!data.services.length) return { error: 'Choose at least one service.' };
  return { data };
}

function syncLabel(booking) {
  if (booking.status === 'cancelled' && booking.larkStatus === 'delete_failed') return 'calendar cancel failed';
  if (booking.status === 'cancelled' && booking.larkStatus === 'cancelled') return 'calendar cancelled';
  if (booking.status === 'cancelled') return 'cancelled';
  if (booking.larkOnly) return 'from Lark';
  if (booking.larkAttendeeStatus === 'needs_review') return 'check guests';
  if (booking.larkAttendeeStatus === 'failed') return 'guests failed';
  return {
    synced: 'in Lark',
    failed: 'Lark failed',
    attendees_failed: 'guests failed',
    pending: 'syncing',
    not_configured: 'local only'
  }[booking.larkStatus] || 'local';
}

function isVisibleOnUpcomingPage(booking) {
  const endTime = new Date(booking.endAt).getTime();
  if (!Number.isFinite(endTime)) return true;
  return endTime >= Date.now() - bookingUpcomingGraceMs;
}

function renderBookings() {
  const upcoming = state.bookings
    .filter(isVisibleOnUpcomingPage)
    .sort((a, b) => new Date(a.startAt) - new Date(b.startAt));
  el.bookingList.innerHTML = '';
  if (!upcoming.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = 'No bookings yet. Create one above and it will appear here.';
    el.bookingList.append(empty);
    updateStats();
    return;
  }
  for (const booking of upcoming) {
    const start = new Date(booking.startAt);
    const end = new Date(booking.endAt);
    const item = el.bookingTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle('cancelled', booking.status === 'cancelled');
    item.querySelector('.month').textContent = monthFormatter.format(start);
    item.querySelector('.day').textContent = dayFormatter.format(start);
    item.querySelector('h3').textContent = booking.propertyAddress || bookingServiceLabel(booking);
    item.querySelector('.booking-time').textContent = `${dateFormatter.format(start)}, ${timeFormatter.format(start)}-${timeFormatter.format(end)}`;
    item.querySelector('.booking-location').textContent = booking.locationAddress || booking.locationName || '';
    item.querySelector('.booking-client').textContent = [booking.clientName, bookingServiceLabel(booking)].filter(Boolean).join(' - ');
    item.querySelector('.booking-contact').textContent = [
      booking.photographerName ? `Photographer: ${formatContact(booking.photographerName, booking.photographerPhone)}` : '',
      booking.agentName ? `Agent: ${formatContact(booking.agentName, booking.agentPhone)}` : ''
    ].filter(Boolean).join(' - ');
    item.querySelector('.booking-notes').textContent = booking.notes || '';

    const pill = item.querySelector('.sync-pill');
    pill.textContent = syncLabel(booking);
    pill.classList.toggle('synced', booking.larkStatus === 'synced');
    pill.classList.toggle('failed', ['failed', 'attendees_failed'].includes(booking.larkStatus) || ['failed', 'needs_review'].includes(booking.larkAttendeeStatus));
    pill.classList.toggle('not-configured', booking.larkStatus === 'not_configured');
    pill.title = [booking.larkError, booking.larkAttendeeError].filter(Boolean).join(' ');

    const editButton = item.querySelector('.edit-booking-button');
    editButton.disabled = booking.status === 'cancelled' || booking.larkOnly;
    editButton.title = booking.larkOnly ? 'This booking was imported from Lark.' : '';
    editButton.addEventListener('click', () => startBookingEdit(booking.id));
    const cancelButton = item.querySelector('.cancel-button');
    const canRetryCalendarCancel = booking.status === 'cancelled' && booking.larkEventId && booking.larkStatus !== 'cancelled';
    const canRemoveCancelledBooking = booking.status === 'cancelled' && !canRetryCalendarCancel && !booking.larkOnly;
    cancelButton.disabled = (booking.status === 'cancelled' && !canRetryCalendarCancel && !canRemoveCancelledBooking) || booking.larkOnly;
    cancelButton.setAttribute('aria-label', canRemoveCancelledBooking ? 'Remove cancelled booking' : 'Cancel booking');
    cancelButton.title = booking.larkOnly
      ? 'Cancel this in Lark.'
      : canRemoveCancelledBooking
        ? 'Remove this cancelled booking from the system.'
        : canRetryCalendarCancel
          ? 'Send calendar cancellation again.'
          : 'Cancel booking';
    cancelButton.addEventListener('click', () => {
      if (canRemoveCancelledBooking) {
        removeCancelledBooking(booking.id);
        return;
      }
      cancelBooking(booking.id);
    });
    el.bookingList.append(item);
  }
  updateStats();
}

function renderMainAssignmentNotice() {
  if (!el.mainAssignmentNotice) return;

  const open = openWorkAssignments();
  const signature = workNoticeSignature(open);
  el.mainAssignmentNotice.hidden = open.length === 0 || signature === state.workNoticeDismissedSignature;

  if (!open.length || signature === state.workNoticeDismissedSignature) {
    el.mainAssignmentNoticeTitle.textContent = 'New work waiting';
    el.mainAssignmentNoticeDetail.textContent = '';
    return;
  }

  const nextAssignment = open[0];
  const countLabel = open.length === 1 ? '1 new assignment' : `${open.length} new assignments`;
  const audienceLabel = state.user?.role === 'employee' ? 'assigned to you' : 'in the work queue';
  const dueLabel = isOverdueWorkAssignment(nextAssignment)
    ? 'overdue'
    : isTodayDateValue(nextAssignment.dueDate)
      ? 'due today'
      : `due ${dateFormatter.format(parseDateValue(nextAssignment.dueDate))}`;

  el.mainAssignmentNoticeTitle.textContent = `${countLabel} ${audienceLabel}`;
  el.mainAssignmentNoticeDetail.textContent = `Next: ${nextAssignment.title} - ${dueLabel}`;
}

function formatMoney(value) {
  return currencyFormatter.format(Number(value || 0));
}

function formatInvoiceDate(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? 'Not set' : invoiceDateFormatter.format(date);
}

function formatInvoiceIsoDate(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? '' : toDateValue(date);
}

function invoiceStatusLabel(status) {
  return {
    draft: 'Draft',
    paid: 'Paid',
    void: 'Void'
  }[status] || 'Draft';
}

function invoiceBookingLabel(invoice) {
  const start = new Date(invoice.bookingStartAt);
  if (Number.isNaN(start.getTime())) return invoice.propertyAddress || 'Booking';
  return `${invoice.propertyAddress || 'Booking'} - ${dateFormatter.format(start)}, ${timeFormatter.format(start)}`;
}

function invoiceProductDescription(invoice) {
  const serviceNames = (invoice.items || []).map((item) => item.name).filter(Boolean).join(' + ');
  return [invoice.propertyAddress || 'Booking', serviceNames].filter(Boolean).join('\n');
}

function invoiceStampLabel(invoice) {
  if (invoice.status === 'paid') return 'PAID';
  if (invoice.status === 'void') return 'VOID';
  return 'UNPAID';
}

function invoicePrintTitle(invoice) {
  return [invoice.invoiceNumber, invoice.propertyAddress, 'Tax Invoice'].filter(Boolean).join(' - ');
}

function invoiceRecipientEmails(invoice) {
  return uniqueEmails(parseEmails(invoice.clientEmail || '')).filter(isEmail);
}

function wageRecipientEmails(wage) {
  return uniqueEmails(parseEmails(wage.photographerEmail || '')).filter(isEmail);
}

function updateInvoiceStats() {
  const drafts = state.invoices.filter((invoice) => invoice.status === 'draft');
  const paid = state.invoices.filter((invoice) => invoice.status === 'paid');
  el.invoiceDraftCount.textContent = drafts.length;
  el.invoicePaidCount.textContent = paid.length;
  el.invoiceTotalValue.textContent = formatMoney(drafts.reduce((sum, invoice) => sum + Number(invoice.total || 0), 0));
}

function updateWageStats() {
  const drafts = state.wages.filter((wage) => wage.status === 'draft');
  const paid = state.wages.filter((wage) => wage.status === 'paid');
  el.wageDraftCount.textContent = drafts.length;
  el.wagePaidCount.textContent = paid.length;
  el.wageTotalValue.textContent = formatMoney(drafts.reduce((sum, wage) => sum + Number(wage.total || 0), 0));
}

function invoiceMatchesFilters(invoice) {
  if (state.invoiceFilters.status !== 'all' && invoice.status !== state.invoiceFilters.status) {
    return false;
  }

  const query = state.invoiceFilters.query.trim().toLowerCase();
  if (!query) return true;

  return [
    invoice.clientName,
    invoice.agentName,
    invoice.propertyAddress,
    invoice.invoiceNumber,
    invoice.clientEmail
  ].some((value) => String(value || '').toLowerCase().includes(query));
}

function renderWages() {
  if (!userCanAccess('wages')) return;

  el.wageList.innerHTML = '';
  const wages = [...state.wages].sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
  updateWageStats();

  if (!wages.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = 'No wages yet. Create or sync booking wages and they will appear here.';
    el.wageList.append(empty);
    return;
  }

  for (const wage of wages) {
    const item = el.wageTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle('paid', wage.status === 'paid');
    item.classList.toggle('void', wage.status === 'void');
    item.querySelector('h3').textContent = `${wage.wageNumber} - ${wage.propertyAddress || 'Photographer wage'}`;
    item.querySelector('.invoice-status').textContent = invoiceStatusLabel(wage.status);
    item.querySelector('.invoice-status').classList.toggle('paid', wage.status === 'paid');
    item.querySelector('.invoice-status').classList.toggle('void', wage.status === 'void');
    item.querySelector('.wage-photographer').textContent = [wage.photographerName, wage.photographerEmail].filter(Boolean).join(' - ') || 'No photographer email';
    item.querySelector('.wage-booking').textContent = `Issued ${formatInvoiceDate(wage.issuedAt)} - ${invoiceBookingLabel(wage)}`;
    item.querySelector('.wage-services').textContent = [
      (wage.items || []).map((item) => `${item.name} ${formatMoney(item.amount)}`).join(' + '),
      isTruthy(wage.photographerGstIncluded) ? `GST included ${formatMoney(wage.gstAmount || 0)}` : 'No GST'
    ].filter(Boolean).join(' - ') || 'No wage items';
    const sentLine = item.querySelector('.wage-sent');
    if (wage.sentAt) {
      const sentTo = Array.isArray(wage.sentTo) && wage.sentTo.length ? ` to ${wage.sentTo.join(', ')}` : '';
      sentLine.textContent = `Sent ${formatInvoiceDate(wage.sentAt)}${sentTo}`;
      sentLine.hidden = false;
    } else {
      sentLine.textContent = '';
      sentLine.hidden = true;
    }
    item.querySelector('.invoice-total strong').textContent = formatMoney(wage.total);
    item.querySelector('.invoice-total small').textContent = `${wage.currency || 'AUD'} ${wageGstLabel(wage.photographerGstIncluded).toLowerCase()}`;

    item.querySelector('.print-wage-button').addEventListener('click', () => printWage(wage.id));
    const sendButton = item.querySelector('.send-wage-button');
    sendButton.hidden = wage.status === 'void';
    sendButton.addEventListener('click', () => sendWage(wage.id));
    const paidButton = item.querySelector('.paid-wage-button');
    paidButton.hidden = wage.status !== 'draft';
    paidButton.addEventListener('click', () => updateWageStatus(wage.id, 'paid'));
    const voidButton = item.querySelector('.void-wage-button');
    voidButton.hidden = wage.status === 'void';
    voidButton.addEventListener('click', () => updateWageStatus(wage.id, 'void'));
    el.wageList.append(item);
  }
}

function renderInvoices() {
  if (!userCanAccess('invoices')) return;

  el.invoiceList.innerHTML = '';
  const invoices = [...state.invoices]
    .filter(invoiceMatchesFilters)
    .sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
  updateInvoiceStats();

  if (!invoices.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = state.invoices.length
      ? 'No invoices match this filter.'
      : 'No invoices yet. Create or sync booking invoices and they will appear here.';
    el.invoiceList.append(empty);
    return;
  }

  for (const invoice of invoices) {
    const item = el.invoiceTemplate.content.firstElementChild.cloneNode(true);
    item.classList.toggle('paid', invoice.status === 'paid');
    item.classList.toggle('void', invoice.status === 'void');
    item.querySelector('h3').textContent = `${invoice.invoiceNumber} - ${invoice.propertyAddress || 'Booking invoice'}`;
    item.querySelector('.invoice-status').textContent = invoiceStatusLabel(invoice.status);
    item.querySelector('.invoice-status').classList.toggle('paid', invoice.status === 'paid');
    item.querySelector('.invoice-status').classList.toggle('void', invoice.status === 'void');
    item.querySelector('.invoice-client').textContent = [invoice.clientName, invoice.agentName].filter(Boolean).join(' - ') || 'No client saved';
    item.querySelector('.invoice-booking').textContent = `Issued ${formatInvoiceDate(invoice.issuedAt)} - due ${formatInvoiceDate(invoice.dueAt)}`;
    item.querySelector('.invoice-services').textContent = [
      (invoice.items || []).map((item) => `${item.name} ${formatMoney(item.amount)}`).join(' + '),
      `GST 10% ${formatMoney(invoice.gstAmount)}`
    ].filter(Boolean).join(' - ');
    const sentLine = item.querySelector('.invoice-sent');
    if (invoice.sentAt) {
      const sentTo = Array.isArray(invoice.sentTo) && invoice.sentTo.length ? ` to ${invoice.sentTo.join(', ')}` : '';
      sentLine.textContent = `Sent ${formatInvoiceDate(invoice.sentAt)}${sentTo}`;
      sentLine.hidden = false;
    } else {
      sentLine.textContent = '';
      sentLine.hidden = true;
    }
    item.querySelector('.invoice-total strong').textContent = formatMoney(invoice.total);
    item.querySelector('.invoice-total small').textContent = `${invoice.currency || 'AUD'} incl. GST`;

    item.querySelector('.print-invoice-button').addEventListener('click', () => printInvoice(invoice.id));
    const editButton = item.querySelector('.edit-invoice-button');
    const sourceBooking = state.bookings.find((booking) => booking.id === invoice.bookingId);
    const canEditSourceBooking = Boolean(invoice.bookingId) && invoice.status !== 'void' && (!sourceBooking || sourceBooking.status !== 'cancelled');
    editButton.disabled = !canEditSourceBooking;
    editButton.title = canEditSourceBooking
      ? 'Edit the booking that created this invoice.'
      : 'Refresh bookings, then edit the source booking for this invoice.';
    editButton.addEventListener('click', () => startInvoiceEdit(invoice.id));
    const sendButton = item.querySelector('.send-invoice-button');
    sendButton.hidden = invoice.status === 'void';
    sendButton.addEventListener('click', () => sendInvoice(invoice.id));
    const paidButton = item.querySelector('.paid-invoice-button');
    paidButton.hidden = invoice.status !== 'draft';
    paidButton.addEventListener('click', () => updateInvoiceStatus(invoice.id, 'paid'));
    const deleteButton = item.querySelector('.delete-invoice-button');
    deleteButton.hidden = invoice.status !== 'void';
    deleteButton.addEventListener('click', () => deleteInvoice(invoice.id));
    const voidButton = item.querySelector('.void-invoice-button');
    voidButton.hidden = invoice.status === 'void';
    voidButton.addEventListener('click', () => updateInvoiceStatus(invoice.id, 'void'));
    el.invoiceList.append(item);
  }
}

function upsertStateInvoice(invoice) {
  if (!invoice || !userCanAccess('invoices')) return;
  const index = state.invoices.findIndex((item) => item.id === invoice.id || item.bookingId === invoice.bookingId);
  if (index >= 0) state.invoices[index] = invoice;
  else state.invoices.unshift(invoice);
  cacheInvoices();
  renderInvoices();
}

function upsertStateWage(wage) {
  if (!wage || !userCanAccess('wages')) return;
  const index = state.wages.findIndex((item) => item.id === wage.id || item.bookingId === wage.bookingId);
  if (index >= 0) state.wages[index] = wage;
  else state.wages.unshift(wage);
  cacheWages();
  renderWages();
}

function escapeHtml(value) {
  return String(value ?? '').replace(/[&<>"']/g, (char) => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[char]));
}

async function printInvoice(id) {
  const invoice = state.invoices.find((item) => item.id === id);
  if (!invoice) return;

  setMessage(el.invoiceMessage, `Preparing ${invoice.invoiceNumber} PDF...`);
  try {
    const response = await fetch(`/api/invoices/${encodeURIComponent(id)}/pdf`, {
      credentials: 'same-origin'
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      setMessage(el.invoiceMessage, (data.errors || ['Could not create invoice PDF.']).join(' '), 'error');
      return;
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${invoicePrintTitle(invoice).replace(/[\\/:*?"<>|]/g, ' - ')}.pdf`;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    setMessage(el.invoiceMessage, `${invoice.invoiceNumber} PDF downloaded.`, 'success');
  } catch {
    setMessage(el.invoiceMessage, 'Could not reach the invoice PDF maker.', 'error');
  }
}

async function printWage(id) {
  const wage = state.wages.find((item) => item.id === id);
  if (!wage) return;

  setMessage(el.wageMessage, `Preparing ${wage.wageNumber} PDF...`);
  try {
    const response = await fetch(`/api/wages/${encodeURIComponent(id)}/pdf`, {
      credentials: 'same-origin'
    });
    if (!response.ok) {
      const data = await response.json().catch(() => ({}));
      setMessage(el.wageMessage, (data.errors || ['Could not create wage PDF.']).join(' '), 'error');
      return;
    }

    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `${[wage.wageNumber, wage.propertyAddress, 'Photographer Proforma'].filter(Boolean).join(' - ').replace(/[\\/:*?"<>|]/g, ' - ')}.pdf`;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 1500);
    setMessage(el.wageMessage, `${wage.wageNumber} PDF downloaded.`, 'success');
  } catch {
    setMessage(el.wageMessage, 'Could not reach the wage PDF maker.', 'error');
  }
}

function startBookingEdit(id) {
  const booking = state.bookings.find((item) => item.id === id);
  if (!booking || booking.status === 'cancelled' || booking.larkOnly) {
    const message = booking?.larkOnly ? 'This booking came from Lark. Edit it in Lark for now.' : (booking ? 'Cancelled bookings cannot be edited.' : 'Booking not found.');
    setMessage(el.formMessage, message, 'error');
    return;
  }
  setRoute('/bookings');
  setBookingMode(booking);
  fillBookingForm(booking);
  setMessage(el.formMessage, 'Editing booking. Update booking to save changes.');
  el.bookingForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

async function startInvoiceEdit(id) {
  const invoice = state.invoices.find((item) => item.id === id);
  if (!invoice) {
    setMessage(el.invoiceMessage, 'Invoice not found.', 'error');
    return;
  }

  if (!invoice.bookingId) {
    setMessage(el.invoiceMessage, 'This invoice is not linked to a booking, so it cannot be edited here.', 'error');
    return;
  }

  let booking = state.bookings.find((item) => item.id === invoice.bookingId);
  if (!booking) {
    setMessage(el.invoiceMessage, 'Finding the booking for this invoice...');
    await loadBookings();
    booking = state.bookings.find((item) => item.id === invoice.bookingId);
  }

  if (!booking) {
    setMessage(el.invoiceMessage, 'The booking for this invoice could not be found.', 'error');
    return;
  }

  if (booking.status === 'cancelled') {
    setMessage(el.invoiceMessage, 'This invoice belongs to a cancelled booking, so it cannot be edited.', 'error');
    return;
  }

  startBookingEdit(booking.id);
  setMessage(el.formMessage, `Editing booking for ${invoice.invoiceNumber}. Save it to refresh the invoice.`, 'success');
}

function updateStats() {
  const now = new Date();
  const today = toDateValue(now);
  const tomorrow = addDaysToDateValue(today, 1);
  const todayStart = australiaDateTimeToUtcDate(today, '00:00').getTime();
  const todayEnd = australiaDateTimeToUtcDate(tomorrow, '00:00').getTime();
  const weekEnd = now.getTime() + 7 * 24 * 60 * 60 * 1000;
  const active = state.bookings.filter((booking) => booking.status !== 'cancelled');
  el.todayCount.textContent = active.filter((booking) => {
    const start = new Date(booking.startAt).getTime();
    return start >= todayStart && start < todayEnd;
  }).length;
  el.weekCount.textContent = active.filter((booking) => {
    const start = new Date(booking.startAt).getTime();
    return start >= now.getTime() && start < weekEnd;
  }).length;
}

async function loadStatus() {
  const { data } = await fetchJson('/api/status');
  el.larkDot.className = `status-dot ${data.larkConfigured ? 'ready' : 'offline'}`;
  el.larkTitle.textContent = data.larkConfigured ? 'Lark connected' : 'Lark not connected';
  el.larkDetail.textContent = data.larkConfigured ? `Calendar ${data.calendarId}, ${data.timezone}` : 'Add your Lark app settings to enable calendar sync.';
}

async function loadBookings() {
  try {
    const { response, data } = await fetchJson('/api/bookings');
    if (!response.ok) {
      el.larkDot.className = 'status-dot offline';
      el.larkTitle.textContent = 'Bookings refresh delayed';
      el.larkDetail.textContent = (data.errors || ['Showing saved bookings for now.']).join(' ');
      return;
    }

    state.bookings = data.bookings || [];
    cacheBookings();
    renderBookings();
    if (data.larkImportError) {
      el.larkDot.className = 'status-dot offline';
      el.larkTitle.textContent = 'Lark import needs checking';
      el.larkDetail.textContent = data.larkImportError;
    }
  } catch {
    el.larkDot.className = 'status-dot offline';
    el.larkTitle.textContent = 'Bookings refresh delayed';
    el.larkDetail.textContent = 'Showing saved bookings while the live calendar catches up.';
  }
}

async function loadWorkAssignments() {
  if (!userCanAccess('work')) {
    state.workAssignments = [];
    renderMainAssignmentNotice();
    return;
  }

  try {
    const { response, data } = await fetchJson('/api/work');
    if (!response.ok) return;
    state.workAssignments = Array.isArray(data.assignments) ? data.assignments : [];
    renderMainAssignmentNotice();
  } catch {
    renderMainAssignmentNotice();
  }
}

async function loadClients() {
  const cached = readStored(storageKeys.clients, []).filter((client) => client?.name);
  if (cached.length) {
    state.clients = mergeDirectoryRecords(state.clients, cached);
    renderClientOptions();
    renderClientList();
  }

  try {
    const { response, data } = await fetchJson('/api/clients');
    if (!response.ok) return;
    state.clients = mergeDirectoryRecords(cached, data.clients || []);
  } catch {
    if (!state.clients.length) state.clients = cached;
  }
  writeStored(storageKeys.clients, state.clients);
  renderClientOptions();
  renderClientList();
}

async function loadPhotographers() {
  const cached = readStored(storageKeys.photographers, []).filter((photographer) => photographer?.name);
  if (cached.length) {
    state.photographers = mergeDirectoryRecords(state.photographers, cached);
    renderPhotographerOptions();
    renderPhotographerList();
  }

  try {
    const { response, data } = await fetchJson('/api/photographers');
    if (!response.ok) return;
    state.photographers = mergeDirectoryRecords(cached, data.photographers || []);
  } catch {
    if (!state.photographers.length) state.photographers = cached;
  }
  writeStored(storageKeys.photographers, state.photographers);
  renderPhotographerOptions();
  renderPhotographerList();
}

async function loadInvoices() {
  if (!userCanAccess('invoices')) return;
  const cached = readStored(storageKeys.invoices, []).filter((invoice) => invoice?.id);
  if (cached.length) {
    state.invoices = cached;
    renderInvoices();
  }

  try {
    const { response, data } = await fetchJson('/api/invoices');
    if (!response.ok) {
      setMessage(el.invoiceMessage, (data.errors || ['Could not load invoices.']).join(' '), 'error');
      return;
    }
    state.invoices = data.invoices || [];
    cacheInvoices();
    renderInvoices();
  } catch {
    if (!state.invoices.length) state.invoices = cached;
    renderInvoices();
    setMessage(el.invoiceMessage, 'Showing saved invoices while the live list catches up.', 'error');
  }
}

async function syncInvoices() {
  if (!userCanAccess('invoices')) return;
  el.syncInvoicesButton.disabled = true;
  setMessage(el.invoiceMessage, 'Creating missing invoices...');
  try {
    const { response, data } = await fetchJson('/api/invoices/sync', { method: 'POST' });
    if (!response.ok) {
      setMessage(el.invoiceMessage, (data.errors || ['Could not sync invoices.']).join(' '), 'error');
      return;
    }
    state.invoices = data.invoices || [];
    cacheInvoices();
    renderInvoices();
    setMessage(el.invoiceMessage, `Invoices synced: ${data.created || 0} created, ${data.updated || 0} updated.`, 'success');
  } catch {
    setMessage(el.invoiceMessage, 'Could not reach the invoice app.', 'error');
  } finally {
    el.syncInvoicesButton.disabled = false;
  }
}

async function loadWages() {
  if (!userCanAccess('wages')) return;
  const cached = readStored(storageKeys.wages, []).filter((wage) => wage?.id);
  if (cached.length) {
    state.wages = cached;
    renderWages();
  }

  try {
    const { response, data } = await fetchJson('/api/wages');
    if (!response.ok) {
      setMessage(el.wageMessage, (data.errors || ['Could not load wages.']).join(' '), 'error');
      return;
    }
    state.wages = data.wages || [];
    cacheWages();
    renderWages();
  } catch {
    if (!state.wages.length) state.wages = cached;
    renderWages();
    setMessage(el.wageMessage, 'Showing saved wages while the live list catches up.', 'error');
  }
}

async function syncWages() {
  if (!userCanAccess('wages')) return;
  el.syncWagesButton.disabled = true;
  setMessage(el.wageMessage, 'Creating missing wages...');
  try {
    const { response, data } = await fetchJson('/api/wages/sync', { method: 'POST' });
    if (!response.ok) {
      setMessage(el.wageMessage, (data.errors || ['Could not sync wages.']).join(' '), 'error');
      return;
    }
    state.wages = data.wages || [];
    cacheWages();
    renderWages();
    setMessage(el.wageMessage, `Wages synced: ${data.created || 0} created, ${data.updated || 0} updated.`, 'success');
  } catch {
    setMessage(el.wageMessage, 'Could not reach the wage app.', 'error');
  } finally {
    el.syncWagesButton.disabled = false;
  }
}

async function sendInvoice(id) {
  const invoice = state.invoices.find((item) => item.id === id);
  if (!invoice) return;

  const recipients = invoiceRecipientEmails(invoice);
  if (!recipients.length) {
    setMessage(el.invoiceMessage, 'Add a client email before sending this invoice.', 'error');
    return;
  }

  if (!window.confirm(`Send ${invoice.invoiceNumber} to ${recipients.join(', ')}?`)) {
    return;
  }

  setMessage(el.invoiceMessage, `Sending ${invoice.invoiceNumber}...`);
  try {
    const { response, data } = await fetchJson(`/api/invoices/${encodeURIComponent(id)}/send`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to: recipients })
    });
    if (!response.ok) {
      setMessage(el.invoiceMessage, (data.errors || ['Could not send invoice.']).join(' '), 'error');
      return;
    }

    state.invoices = data.invoices || state.invoices.map((invoiceItem) => invoiceItem.id === id ? data.invoice : invoiceItem);
    cacheInvoices();
    renderInvoices();
    setMessage(el.invoiceMessage, data.message || 'Invoice sent.', 'success');
  } catch {
    setMessage(el.invoiceMessage, 'Could not reach the invoice email sender.', 'error');
  }
}

async function sendWage(id) {
  const wage = state.wages.find((item) => item.id === id);
  if (!wage) return;

  const recipients = wageRecipientEmails(wage);
  if (!recipients.length) {
    setMessage(el.wageMessage, 'Add a photographer email before sending this proforma.', 'error');
    return;
  }

  if (!window.confirm(`Send ${wage.wageNumber} to ${recipients.join(', ')}?`)) {
    return;
  }

  setMessage(el.wageMessage, `Sending ${wage.wageNumber}...`);
  try {
    const { response, data } = await fetchJson(`/api/wages/${encodeURIComponent(id)}/send`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ to: recipients })
    });
    if (!response.ok) {
      setMessage(el.wageMessage, (data.errors || ['Could not send proforma.']).join(' '), 'error');
      return;
    }

    state.wages = data.wages || state.wages.map((wageItem) => wageItem.id === id ? data.wage : wageItem);
    cacheWages();
    renderWages();
    setMessage(el.wageMessage, data.message || 'Proforma sent.', 'success');
  } catch {
    setMessage(el.wageMessage, 'Could not reach the wage email sender.', 'error');
  }
}

async function updateInvoiceStatus(id, status) {
  const action = status === 'paid' ? 'Marking invoice paid...' : 'Voiding invoice...';
  setMessage(el.invoiceMessage, action);
  try {
    const { response, data } = await fetchJson(`/api/invoices/${encodeURIComponent(id)}/${status}`, { method: 'POST' });
    if (!response.ok) {
      setMessage(el.invoiceMessage, (data.errors || ['Could not update invoice.']).join(' '), 'error');
      return;
    }
    state.invoices = data.invoices || state.invoices.map((invoice) => invoice.id === id ? data.invoice : invoice);
    cacheInvoices();
    renderInvoices();
    setMessage(el.invoiceMessage, status === 'paid' ? 'Invoice marked paid.' : 'Invoice voided.', 'success');
  } catch {
    setMessage(el.invoiceMessage, 'Could not reach the invoice app.', 'error');
  }
}

async function deleteInvoice(id) {
  const invoice = state.invoices.find((item) => item.id === id);
  if (!invoice) return;

  if (invoice.status !== 'void') {
    setMessage(el.invoiceMessage, 'Only voided invoices can be deleted.', 'error');
    return;
  }

  if (!window.confirm(`Delete voided invoice ${invoice.invoiceNumber}? This removes it from the invoice list.`)) {
    return;
  }

  setMessage(el.invoiceMessage, `Deleting ${invoice.invoiceNumber}...`);
  try {
    const { response, data } = await fetchJson(`/api/invoices/${encodeURIComponent(id)}`, {
      method: 'DELETE'
    });
    if (!response.ok) {
      setMessage(el.invoiceMessage, (data.errors || ['Could not delete invoice.']).join(' '), 'error');
      return;
    }
    state.invoices = data.invoices || state.invoices.filter((invoiceItem) => invoiceItem.id !== id);
    cacheInvoices();
    renderInvoices();
    setMessage(el.invoiceMessage, `${invoice.invoiceNumber} deleted.`, 'success');
  } catch {
    setMessage(el.invoiceMessage, 'Could not reach the invoice app.', 'error');
  }
}

async function updateWageStatus(id, status) {
  const action = status === 'paid' ? 'Marking wage paid...' : 'Voiding wage...';
  setMessage(el.wageMessage, action);
  try {
    const { response, data } = await fetchJson(`/api/wages/${encodeURIComponent(id)}/${status}`, { method: 'POST' });
    if (!response.ok) {
      setMessage(el.wageMessage, (data.errors || ['Could not update wage.']).join(' '), 'error');
      return;
    }
    state.wages = data.wages || state.wages.map((wage) => wage.id === id ? data.wage : wage);
    cacheWages();
    renderWages();
    setMessage(el.wageMessage, status === 'paid' ? 'Wage marked paid.' : 'Wage voided.', 'success');
  } catch {
    setMessage(el.wageMessage, 'Could not reach the wage app.', 'error');
  }
}

async function saveDirectoryClient(event) {
  event.preventDefault();
  const payload = {
    id: el.directoryClientId.value || undefined,
    name: el.directoryClientName.value.trim(),
    email: el.directoryClientEmail.value.trim(),
    agentName: el.directoryAgentName.value.trim(),
    agentPhone: el.directoryAgentPhone.value.trim()
  };
  if (!payload.name) {
    setMessage(el.clientMessage, 'Enter the agency or client name before saving.', 'error');
    return;
  }
  const button = el.clientForm.querySelector('button[type=submit]');
  button.disabled = true;
  button.textContent = 'Saving';
  setMessage(el.clientMessage, 'Saving client...');
  try {
    const { response, data } = await fetchJson('/api/clients', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      setMessage(el.clientMessage, (data.errors || ['Could not save client.']).join(' '), 'error');
      return;
    }
    state.clients = data.clients || state.clients;
    writeStored(storageKeys.clients, state.clients);
    updateBookingFormState({
      selectedClientId: data.client.id,
      clientSearchQuery: data.client ? clientOptionLabel(data.client) : '',
      clientSearchOpen: false
    }, { renderSelections: false });
    renderClientOptions();
    renderClientList();
    applySelectedClient();
    el.clientForm.reset();
    el.directoryClientId.value = '';
    applyExamplePlaceholders();
    setMessage(el.clientMessage, 'Client saved for bookings.', 'success');
  } catch {
    setMessage(el.clientMessage, 'Could not reach the client directory.', 'error');
  } finally {
    button.disabled = false;
    button.textContent = 'Save client';
  }
}

async function saveDirectoryPhotographer(event) {
  event.preventDefault();
  const payload = {
    id: el.directoryPhotographerId.value || undefined,
    name: el.directoryPhotographerName.value.trim(),
    email: el.directoryPhotographerEmail.value.trim(),
    phone: el.directoryPhotographerPhone.value.trim(),
    gstIncluded: el.directoryPhotographerGstIncluded.checked
  };
  if (!payload.name) {
    setMessage(el.photographerMessage, 'Enter the photographer name before saving.', 'error');
    return;
  }
  const button = el.photographerForm.querySelector('button[type=submit]');
  button.disabled = true;
  button.textContent = 'Saving';
  setMessage(el.photographerMessage, 'Saving photographer...');
  try {
    const { response, data } = await fetchJson('/api/photographers', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    if (!response.ok) {
      setMessage(el.photographerMessage, (data.errors || ['Could not save photographer.']).join(' '), 'error');
      return;
    }
    state.photographers = data.photographers || state.photographers;
    writeStored(storageKeys.photographers, state.photographers);
    updateBookingFormState({ selectedPhotographerId: data.photographer.id }, { renderSelections: false });
    renderPhotographerOptions();
    renderPhotographerList();
    applySelectedPhotographer();
    el.photographerForm.reset();
    el.directoryPhotographerId.value = '';
    el.directoryPhotographerGstIncluded.checked = false;
    applyExamplePlaceholders();
    setMessage(el.photographerMessage, 'Photographer saved for bookings.', 'success');
  } catch {
    setMessage(el.photographerMessage, 'Could not reach the photographer directory.', 'error');
  } finally {
    button.disabled = false;
    button.textContent = 'Save photographer';
  }
}

function editClient(id) {
  const client = state.clients.find((item) => item.id === id);
  if (!client) return;
  el.directoryClientId.value = client.id;
  el.directoryClientName.value = client.name || '';
  el.directoryClientEmail.value = client.email || '';
  el.directoryAgentName.value = client.agentName || '';
  el.directoryAgentPhone.value = client.agentPhone || '';
  setMessage(el.clientMessage, 'Editing saved client.');
}

function editPhotographer(id) {
  const photographer = state.photographers.find((item) => item.id === id);
  if (!photographer) return;
  el.directoryPhotographerId.value = photographer.id;
  el.directoryPhotographerName.value = photographer.name || '';
  el.directoryPhotographerEmail.value = photographer.email || '';
  el.directoryPhotographerPhone.value = photographer.phone || '';
  el.directoryPhotographerGstIncluded.checked = isTruthy(photographer.gstIncluded);
  setMessage(el.photographerMessage, 'Editing saved photographer.');
}

async function deleteClient(id) {
  const client = state.clients.find((item) => item.id === id);
  if (!client) return;
  if (!window.confirm(`Delete ${client.name}?`)) return;

  setMessage(el.clientMessage, 'Deleting client...');
  try {
    const { response, data } = await fetchJson(`/api/clients/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (!response.ok) {
      setMessage(el.clientMessage, (data.errors || ['Could not delete client.']).join(' '), 'error');
      return;
    }

    state.clients = data.clients || [];
    writeStored(storageKeys.clients, state.clients);
    if (state.bookingForm.selectedClientId === id) {
      updateBookingFormState({
        selectedClientId: '',
        clientSearchQuery: '',
        clientSearchOpen: false
      }, { renderSelections: false });
      applySelectedClient();
    }
    if (el.directoryClientId.value === id) {
      el.clientForm.reset();
      el.directoryClientId.value = '';
    }
    renderClientOptions();
    renderClientList();
    setMessage(el.clientMessage, 'Client deleted.', 'success');
  } catch {
    setMessage(el.clientMessage, 'Could not reach the client directory.', 'error');
  }
}

async function deletePhotographer(id) {
  const photographer = state.photographers.find((item) => item.id === id);
  if (!photographer) return;
  if (!window.confirm(`Delete ${photographer.name}?`)) return;

  setMessage(el.photographerMessage, 'Deleting photographer...');
  try {
    const { response, data } = await fetchJson(`/api/photographers/${encodeURIComponent(id)}`, { method: 'DELETE' });
    if (!response.ok) {
      setMessage(el.photographerMessage, (data.errors || ['Could not delete photographer.']).join(' '), 'error');
      return;
    }

    state.photographers = data.photographers || [];
    writeStored(storageKeys.photographers, state.photographers);
    if (state.bookingForm.selectedPhotographerId === id) {
      updateBookingFormState({ selectedPhotographerId: '' }, { renderSelections: false });
      applySelectedPhotographer();
    }
    if (el.directoryPhotographerId.value === id) {
      el.photographerForm.reset();
      el.directoryPhotographerId.value = '';
      el.directoryPhotographerGstIncluded.checked = false;
    }
    renderPhotographerOptions();
    renderPhotographerList();
    setMessage(el.photographerMessage, 'Photographer deleted.', 'success');
  } catch {
    setMessage(el.photographerMessage, 'Could not reach the photographer directory.', 'error');
  }
}

function bookingSavedMessage(booking, action) {
  if (booking.calendarInviteStatus === 'sent') return `Booking ${action}. Calendar invite sent from admin@openframe.studio.`;
  if (booking.calendarInviteStatus === 'failed') return `Booking ${action}, but the admin invite email could not send.`;
  if (booking.calendarInviteStatus === 'not_configured') return `Booking ${action}. Add Resend settings to send invites from admin@openframe.studio.`;
  if (booking.larkAttendeeStatus === 'needs_review') return `Booking ${action}. Check guest removals in Lark.`;
  if (booking.larkStatus === 'synced') return `Booking ${action} and synced to Lark.`;
  if (booking.larkStatus === 'attendees_failed') return `Booking ${action}, but Lark guest invitations need checking.`;
  if (booking.larkStatus === 'failed') return `Booking ${action} locally, but Lark did not update.`;
  return `Booking ${action} locally. Add Lark settings when ready.`;
}

async function submitBooking(event) {
  event.preventDefault();
  const { data, error } = buildBookingPayload();
  if (error) {
    setMessage(el.formMessage, error, 'error');
    return;
  }
  const isEditing = Boolean(state.editingBookingId);
  const endpoint = isEditing ? `/api/bookings/${state.editingBookingId}` : '/api/bookings';
  setMessage(el.formMessage, isEditing ? 'Updating booking...' : 'Creating booking...');
  el.submitButton.disabled = true;
  try {
    const { response, data: result } = await fetchJson(endpoint, {
      method: isEditing ? 'PUT' : 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!response.ok) {
      setMessage(el.formMessage, (result.errors || ['Could not save booking.']).join(' '), 'error');
      return;
    }
    if (isEditing) {
      state.bookings = state.bookings.map((booking) => booking.id === result.booking.id ? result.booking : booking);
    } else {
      state.bookings.push(result.booking);
    }
    upsertStateInvoice(result.invoice);
    upsertStateWage(result.wage);
    cacheBookings();
    rememberAddress(result.booking.propertyAddress);
    renderBookings();
    removeStored(storageKeys.draft);
    resetBookingForm(bookingSavedMessage(result.booking, isEditing ? 'updated' : 'created'));
    el.formMessage.className = 'success';
  } catch {
    setMessage(el.formMessage, 'Could not reach the booking server.', 'error');
  } finally {
    el.submitButton.disabled = false;
  }
}

function renderLarkPreview(preview) {
  const start = new Date(preview.startAt);
  const end = new Date(preview.endAt);
  el.previewTitle.textContent = preview.title;
  el.previewTime.textContent = `${dateFormatter.format(start)}, ${timeFormatter.format(start)}-${timeFormatter.format(end)} (${preview.timezone})`;
  el.previewLocation.textContent = preview.location.address || preview.location.name || '';
  el.previewGuests.textContent = preview.guestEmails.length ? preview.guestEmails.join(', ') : 'No guests';
  el.previewDescription.textContent = preview.description;
  el.larkPreview.hidden = false;
}

async function previewLarkEvent() {
  const { data, error } = buildBookingPayload();
  if (error) {
    setMessage(el.formMessage, error, 'error');
    return;
  }
  el.previewLarkButton.disabled = true;
  el.previewLarkButton.textContent = 'Previewing';
  setMessage(el.formMessage, 'Building Lark preview...');
  try {
    const { response, data: result } = await fetchJson('/api/lark/preview', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(data)
    });
    if (!response.ok) {
      setMessage(el.formMessage, (result.errors || ['Could not build preview.']).join(' '), 'error');
      return;
    }
    renderLarkPreview(result.preview);
    setMessage(el.formMessage, 'Lark preview generated. Nothing was sent.', 'success');
  } catch {
    setMessage(el.formMessage, 'Could not reach the preview service.', 'error');
  } finally {
    el.previewLarkButton.disabled = false;
    el.previewLarkButton.textContent = 'Preview Lark event';
  }
}

async function cancelBooking(id) {
  const { response, data } = await fetchJson(`/api/bookings/${id}/cancel`, { method: 'POST' });
  if (!response.ok) return;
  state.bookings = state.bookings.map((booking) => booking.id === id ? data.booking : booking);
  upsertStateInvoice(data.invoice);
  upsertStateWage(data.wage);
  cacheBookings();
  if (state.editingBookingId === id) resetBookingForm('Edit cancelled because the booking was cancelled.');
  renderBookings();
}

async function removeCancelledBooking(id) {
  const { response, data } = await fetchJson(`/api/bookings/${id}`, { method: 'DELETE' });
  if (!response.ok) {
    if (data.booking) {
      state.bookings = state.bookings.map((booking) => booking.id === id ? data.booking : booking);
      renderBookings();
    }
    setMessage(el.formMessage, (data.errors || ['Could not remove this booking.']).join(' '), 'error');
    return;
  }

  state.bookings = Array.isArray(data.bookings)
    ? data.bookings
    : state.bookings.filter((booking) => booking.id !== id);
  cacheBookings();
  if (state.editingBookingId === id) resetBookingForm('Removed cancelled booking.');
  renderBookings();
  setMessage(el.formMessage, 'Cancelled booking removed. That time can be booked again.', 'success');
}

function cancelBookingEdit() {
  resetBookingForm('Edit cancelled.');
  removeStored(storageKeys.draft);
}

async function testLark() {
  el.testLarkButton.disabled = true;
  el.testLarkButton.textContent = 'Testing';
  try {
    const { data } = await fetchJson('/api/lark/test', { method: 'POST' });
    el.larkDot.className = `status-dot ${data.ok ? 'ready' : 'offline'}`;
    el.larkTitle.textContent = data.ok ? 'Lark test passed' : 'Lark test failed';
    el.larkDetail.textContent = data.message;
  } catch {
    el.larkDot.className = 'status-dot offline';
    el.larkTitle.textContent = 'Lark test failed';
    el.larkDetail.textContent = 'The booking server could not complete the test.';
  } finally {
    el.testLarkButton.disabled = false;
    el.testLarkButton.textContent = 'Test Lark';
  }
}

function openPasswordDialog() {
  el.passwordForm.reset();
  setMessage(el.passwordMessage, '');
  el.savePasswordButton.disabled = false;

  if (typeof el.passwordDialog.showModal === 'function') {
    el.passwordDialog.showModal();
  } else {
    el.passwordDialog.setAttribute('open', '');
  }

  requestAnimationFrame(() => el.currentPassword.focus());
}

function closePasswordDialog() {
  el.passwordForm.reset();
  setMessage(el.passwordMessage, '');

  if (typeof el.passwordDialog.close === 'function' && el.passwordDialog.open) {
    el.passwordDialog.close();
  } else {
    el.passwordDialog.removeAttribute('open');
  }
}

async function changePassword(event) {
  event.preventDefault();
  const currentPassword = el.currentPassword.value;
  const newPassword = el.newPassword.value;
  const confirmPassword = el.confirmPassword.value;

  if (newPassword.length < 4) {
    setMessage(el.passwordMessage, 'Use at least 4 characters.', 'error');
    return;
  }

  if (newPassword !== confirmPassword) {
    setMessage(el.passwordMessage, 'The new passwords do not match.', 'error');
    return;
  }

  el.savePasswordButton.disabled = true;
  el.savePasswordButton.textContent = 'Saving';

  try {
    const { response, data } = await fetchJson('/api/change-password', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ currentPassword, newPassword, confirmPassword })
    });

    if (!response.ok) {
      setMessage(el.passwordMessage, data.errors?.[0] || 'Password could not be changed.', 'error');
      return;
    }

    setMessage(el.passwordMessage, 'Password updated.', 'success');
    setTimeout(closePasswordDialog, 650);
  } catch {
    setMessage(el.passwordMessage, 'Password could not be changed.', 'error');
  } finally {
    el.savePasswordButton.disabled = false;
    el.savePasswordButton.textContent = 'Save password';
  }
}

async function logout() {
  el.logoutButton.disabled = true;
  try {
    await fetch('/api/logout', { method: 'POST' });
  } finally {
    window.location.href = '/login';
  }
}

function viewOpenWorkAssignments() {
  const signature = workNoticeSignature(openWorkAssignments());
  state.workNoticeDismissedSignature = signature;
  writeWorkNoticeDismissedSignature(signature);
  renderMainAssignmentNotice();
  window.location.href = '/work/';
}

el.bookingForm.addEventListener('submit', submitBooking);
el.bookingForm.addEventListener('input', saveDraft);
el.bookingForm.addEventListener('change', saveDraft);
el.clientForm.addEventListener('submit', saveDirectoryClient);
el.photographerForm.addEventListener('submit', saveDirectoryPhotographer);
el.refreshButton.addEventListener('click', loadBookings);
el.refreshInvoicesButton.addEventListener('click', loadInvoices);
el.syncInvoicesButton.addEventListener('click', syncInvoices);
el.invoiceStatusFilter.addEventListener('change', () => {
  state.invoiceFilters.status = el.invoiceStatusFilter.value;
  renderInvoices();
});
el.invoiceClientFilter.addEventListener('input', () => {
  state.invoiceFilters.query = el.invoiceClientFilter.value;
  renderInvoices();
});
el.refreshWagesButton.addEventListener('click', loadWages);
el.syncWagesButton.addEventListener('click', syncWages);
el.testLarkButton.addEventListener('click', testLark);
el.previewLarkButton.addEventListener('click', previewLarkEvent);
el.cancelEditButton.addEventListener('click', cancelBookingEdit);
el.changePasswordButton.addEventListener('click', openPasswordDialog);
el.passwordForm.addEventListener('submit', changePassword);
el.cancelPasswordButton.addEventListener('click', closePasswordDialog);
el.passwordDialog.addEventListener('click', (event) => {
  if (event.target === el.passwordDialog) closePasswordDialog();
});
el.logoutButton.addEventListener('click', logout);
el.mainAssignmentNoticeButton?.addEventListener('click', viewOpenWorkAssignments);
el.clientDropdownButton.addEventListener('click', () => {
  setClientDropdownOpen(!state.bookingForm.clientDropdownOpen);
});
el.photographerDropdownButton.addEventListener('click', () => {
  setPhotographerDropdownOpen(!state.bookingForm.photographerDropdownOpen);
});
el.clientSearch.addEventListener('input', () => {
  updateBookingFormState({
    selectedClientId: '',
    clientSearchQuery: el.clientSearch.value,
    clientSearchOpen: true
  }, { renderSelections: false });
  renderClientOptions();
});
el.clientSearch.addEventListener('focus', () => {
  updateBookingFormState({ clientSearchOpen: true }, { renderSelections: false });
  renderClientSearchResults();
});
el.clientSelect.addEventListener('change', () => {
  selectClient(el.clientSelect.value);
});
el.photographerSelect.addEventListener('change', () => {
  selectPhotographer(el.photographerSelect.value);
});
document.addEventListener('click', (event) => {
  if (!el.clientDropdown.contains(event.target)) closeClientDropdown();
  if (!el.photographerDropdown.contains(event.target)) closePhotographerDropdown();
});
document.addEventListener('keydown', (event) => {
  if (event.key === 'Escape') {
    closeClientDropdown();
    closePhotographerDropdown();
  }
});
el.clientEmail.addEventListener('input', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.clientEmail.addEventListener('change', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.clientEmail.addEventListener('blur', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.photographerEmail.addEventListener('input', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.photographerEmail.addEventListener('change', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.photographerEmail.addEventListener('blur', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.guestEmails.addEventListener('input', updateInvitationSummary);
el.newClientButton.addEventListener('click', () => {
  el.clientForm.reset();
  el.directoryClientId.value = '';
  applyExamplePlaceholders();
  setMessage(el.clientMessage, '');
});
el.newPhotographerButton.addEventListener('click', () => {
  el.photographerForm.reset();
  el.directoryPhotographerId.value = '';
  el.directoryPhotographerGstIncluded.checked = false;
  applyExamplePlaceholders();
  setMessage(el.photographerMessage, '');
});
el.serviceInputs.forEach((input) => input.addEventListener('change', () => {
  updateDurationForServices();
  saveDraft();
}));
$$('[data-route]').forEach((link) => link.addEventListener('click', (event) => {
  event.preventDefault();
  setRoute(link.dataset.route);
}));
window.addEventListener('popstate', () => setRoute(window.location.pathname, false));

setInitialDateTime();
updateDurationForServices();
applyExamplePlaceholders();
await loadSession();
applyAppAccess();
hydrateCachedData();
restoreDraft();
updateInvitationSummary();
setRoute(window.location.pathname, false);
refreshLiveData();
