const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const el = {
  bookingForm: $('#bookingForm'),
  bookingList: $('#bookingList'),
  bookingTemplate: $('#bookingTemplate'),
  clientTemplate: $('#clientTemplate'),
  photographerTemplate: $('#photographerTemplate'),
  bookingPage: $('#bookingPage'),
  clientsPage: $('#clientsPage'),
  photographersPage: $('#photographersPage'),
  propertyAddress: $('#propertyAddressInput'),
  addressDatalist: $('#addressSuggestionList'),
  addressSuggestions: $('#addressSuggestions'),
  addressMessage: $('#addressMessage'),
  formMessage: $('#formMessage'),
  clientMessage: $('#clientMessage'),
  photographerMessage: $('#photographerMessage'),
  date: $('#dateInput'),
  time: $('#timeInput'),
  duration: $('#durationInput'),
  serviceInputs: $$('input[name=services]'),
  clientSearch: $('#clientSearchInput'),
  clientSelect: $('#clientSelect'),
  clientName: $('#clientNameInput'),
  clientEmail: $('#clientEmailInput'),
  agentName: $('#agentNameInput'),
  agentPhone: $('#agentPhoneInput'),
  photographerSelect: $('#photographerSelect'),
  photographerName: $('#photographerNameInput'),
  photographerEmail: $('#photographerEmailInput'),
  photographerPhone: $('#photographerPhoneInput'),
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
  refreshButton: $('#refreshButton'),
  searchAddressButton: $('#searchAddressButton'),
  testLarkButton: $('#testLarkButton'),
  previewLarkButton: $('#previewLarkButton'),
  submitButton: $('#bookingSubmitButton'),
  cancelEditButton: $('#cancelEditButton'),
  logoutButton: $('#logoutButton'),
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
  todayCount: $('#todayCount'),
  weekCount: $('#weekCount'),
  modeEyebrow: $('#bookingModeEyebrow'),
  modeTitle: $('#bookingModeTitle')
};

const state = {
  bookings: [],
  clients: [],
  photographers: [],
  editingBookingId: '',
  restoringDraft: false,
  syncedClientEmail: [],
  syncedPhotographerEmail: []
};

const storageKeys = {
  draft: 'openframe.bookingDraft.v2',
  clients: 'openframe.clients.v1',
  photographers: 'openframe.photographers.v1',
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

const dateFormatter = new Intl.DateTimeFormat(undefined, { weekday: 'short', month: 'short', day: 'numeric' });
const timeFormatter = new Intl.DateTimeFormat(undefined, { hour: 'numeric', minute: '2-digit' });
const monthFormatter = new Intl.DateTimeFormat(undefined, { month: 'short' });
const dayFormatter = new Intl.DateTimeFormat(undefined, { day: '2-digit' });

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

function localAddressSuggestions(query = el.propertyAddress.value) {
  const search = normalizeAddress(query).toLowerCase();
  const saved = readStored(storageKeys.addresses, []);
  const bookingAddresses = state.bookings.flatMap((booking) => [
    booking.propertyAddress,
    booking.locationAddress,
    booking.locationName
  ]);
  return uniqueAddresses([...saved, ...bookingAddresses])
    .filter((address) => !search || address.toLowerCase().includes(search))
    .slice(0, 8);
}

function normalizeRemoteAddress(value) {
  return normalizeAddress(String(value || '').replace(/,\s*AUS$/i, ', Australia'));
}

async function fetchBrowserAddressSuggestions(query) {
  const params = new URLSearchParams({
    f: 'json',
    text: query,
    countryCode: 'AUS',
    maxSuggestions: '6'
  });
  const response = await fetch(`https://geocode.arcgis.com/arcgis/rest/services/World/GeocodeServer/suggest?${params}`);
  if (!response.ok) return [];
  const data = await response.json().catch(() => ({}));
  return (data.suggestions || []).map((suggestion) => normalizeRemoteAddress(suggestion.text)).filter(Boolean);
}

function renderAddressSuggestions(addresses, message = '') {
  const suggestions = uniqueAddresses(addresses).slice(0, 8);
  el.addressDatalist.innerHTML = '';
  el.addressSuggestions.innerHTML = '';

  for (const address of suggestions) {
    el.addressDatalist.append(new Option(address, address));
    const button = document.createElement('button');
    button.className = 'address-suggestion';
    button.type = 'button';
    button.textContent = address;
    button.addEventListener('click', () => {
      el.propertyAddress.value = address;
      rememberAddress(address);
      renderAddressSuggestions(localAddressSuggestions(address));
      saveDraft();
    });
    el.addressSuggestions.append(button);
  }

  el.addressSuggestions.hidden = !suggestions.length;
  if (message) {
    setMessage(el.addressMessage, message);
  }
}

function updateAddressSuggestions() {
  renderAddressSuggestions(localAddressSuggestions(), '');
}

async function searchAddressSuggestions() {
  const query = normalizeAddress(el.propertyAddress.value);
  if (query.length < 4) {
    setMessage(el.addressMessage, 'Type at least 4 characters.', 'error');
    return;
  }

  el.searchAddressButton.disabled = true;
  el.searchAddressButton.textContent = 'Searching';
  setMessage(el.addressMessage, 'Searching addresses...');
  try {
    let remoteAddresses = await fetchBrowserAddressSuggestions(query);
    if (!remoteAddresses.length) {
      const { response, data } = await fetchJson(`/api/address-suggestions?q=${encodeURIComponent(query)}`);
      if (!response.ok) {
        setMessage(el.addressMessage, (data.errors || ['Could not find address suggestions.']).join(' '), 'error');
        return;
      }
      remoteAddresses = (data.suggestions || []).map((suggestion) => suggestion.address || suggestion.label);
    }
    const addresses = uniqueAddresses([...remoteAddresses, ...localAddressSuggestions(query)]);
    renderAddressSuggestions(addresses, addresses.length ? 'Choose an address above.' : 'No suggestions found.');
    if (addresses.length) {
      setMessage(el.addressMessage, 'Choose an address above.', 'success');
    }
  } catch {
    setMessage(el.addressMessage, 'Could not reach address lookup.', 'error');
  } finally {
    el.searchAddressButton.disabled = false;
    el.searchAddressButton.textContent = 'Find address suggestions';
  }
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

function bookingServiceLabel(booking) {
  if (Array.isArray(booking.services) && booking.services.length) {
    return booking.services.map((service) => service.name).join(' + ');
  }
  return booking.service || 'Booking';
}

function toDateValue(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function toTimeValue(date) {
  return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
}

function setInitialDateTime() {
  const next = new Date();
  next.setMinutes(Math.ceil(next.getMinutes() / 15) * 15, 0, 0);
  if (next.getHours() >= 18) {
    next.setDate(next.getDate() + 1);
    next.setHours(9, 0, 0, 0);
  }
  el.date.min = toDateValue(new Date());
  el.date.value = toDateValue(next);
  el.time.value = toTimeValue(next);
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

function renderOptions(select, items, emptyLabel, selectedId = select.value) {
  select.innerHTML = '';
  select.append(new Option(emptyLabel, ''));
  for (const item of [...items].sort((a, b) => a.name.localeCompare(b.name))) {
    const subtitle = item.agentName || item.phone || '';
    select.append(new Option(subtitle ? `${item.name} - ${subtitle}` : item.name, item.id));
  }
  select.value = selectedId || '';
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
  const query = normalizeAddress(el.clientSearch.value).toLowerCase();
  if (!query) return state.clients;
  const terms = query.split(/\s+/).filter(Boolean);
  return state.clients.filter((client) => terms.every((term) => clientSearchText(client).includes(term)));
}

function renderClientOptions(selectedId = el.clientSelect.value) {
  const clients = filteredClients();
  const hasSelected = state.clients.some((client) => client.id === selectedId);
  const selectedIsVisible = clients.some((client) => client.id === selectedId);
  const options = hasSelected && selectedId && !selectedIsVisible
    ? [state.clients.find((client) => client.id === selectedId), ...clients]
    : clients;
  renderOptions(el.clientSelect, options.filter(Boolean), clients.length ? 'New client' : 'No matching clients', selectedId);
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
    item.querySelector('.photographer-phone').textContent = photographer.phone ? `Phone: ${photographer.phone}` : '';
    item.querySelector('.edit-photographer-button').addEventListener('click', () => editPhotographer(photographer.id));
    item.querySelector('.delete-photographer-button').addEventListener('click', () => deletePhotographer(photographer.id));
    el.photographerList.append(item);
  }
}

function applySelectedClient() {
  const client = state.clients.find((item) => item.id === el.clientSelect.value);
  if (!client) return;
  el.clientSearch.value = '';
  renderClientOptions(client.id);
  el.clientName.value = client.name || '';
  el.clientEmail.value = client.email || '';
  el.agentName.value = client.agentName || '';
  el.agentPhone.value = client.agentPhone || '';
  syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail');
}

function applySelectedPhotographer() {
  const photographer = state.photographers.find((item) => item.id === el.photographerSelect.value);
  if (!photographer) return;
  el.photographerName.value = photographer.name || '';
  el.photographerEmail.value = photographer.email || '';
  el.photographerPhone.value = photographer.phone || '';
  syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail');
}

function getDraft() {
  const data = Object.fromEntries(new FormData(el.bookingForm).entries());
  return {
    propertyAddress: el.propertyAddress.value || data.propertyAddress || '',
    selectedClientId: el.clientSelect.value,
    clientName: el.clientName.value,
    clientEmail: el.clientEmail.value,
    agentName: el.agentName.value,
    agentPhone: el.agentPhone.value,
    services: selectedServices().map((service) => service.name),
    date: el.date.value,
    time: el.time.value,
    durationMinutes: el.duration.value,
    selectedPhotographerId: el.photographerSelect.value,
    photographerName: el.photographerName.value,
    photographerEmail: el.photographerEmail.value,
    photographerPhone: el.photographerPhone.value,
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
  el.clientSelect.value = matchingClient?.id || '';
  el.clientName.value = booking.clientName || '';
  el.clientEmail.value = booking.clientEmail || '';
  el.agentName.value = booking.agentName || '';
  el.agentPhone.value = booking.agentPhone || '';
  el.photographerSelect.value = matchingPhotographer?.id || '';
  el.photographerName.value = booking.photographerName || matchingPhotographer?.name || 'Barry';
  el.photographerEmail.value = booking.photographerEmail || matchingPhotographer?.email || '';
  el.photographerPhone.value = booking.photographerPhone || matchingPhotographer?.phone || '0403 007 853';
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
  updateAddressSuggestions();
  el.larkPreview.hidden = true;
  state.restoringDraft = false;
}

function resetBookingForm(message = '') {
  setBookingMode();
  el.bookingForm.reset();
  el.clientSearch.value = '';
  el.clientSelect.value = '';
  el.photographerSelect.value = '';
  state.syncedClientEmail = [];
  state.syncedPhotographerEmail = [];
  setInitialDateTime();
  updateDurationForServices();
  applyExamplePlaceholders();
  updateInvitationSummary();
  updateAddressSuggestions();
  el.larkPreview.hidden = true;
  setMessage(el.formMessage, message);
}

function setRoute(route, push = true) {
  const nextRoute = route === '/clients' || route === '/photographers' ? route : '/';
  el.bookingPage.hidden = nextRoute !== '/';
  el.clientsPage.hidden = nextRoute !== '/clients';
  el.photographersPage.hidden = nextRoute !== '/photographers';
  $$('[data-route]').forEach((link) => link.classList.toggle('active', link.dataset.route === nextRoute));
  if (push && window.location.pathname !== nextRoute) history.pushState({ route: nextRoute }, '', nextRoute);
  if (nextRoute === '/clients') renderClientList();
  if (nextRoute === '/photographers') renderPhotographerList();
}

function buildBookingPayload() {
  const start = el.date.value && el.time.value ? new Date(`${el.date.value}T${el.time.value}:00`) : null;
  if (!start || Number.isNaN(start.getTime())) return { error: 'Choose a valid date and time.' };
  syncAutoEmails();
  const data = Object.fromEntries(new FormData(el.bookingForm).entries());
  data.clientEmail = el.clientEmail.value.trim();
  data.photographerName = el.photographerName.value.trim();
  data.photographerEmail = el.photographerEmail.value.trim();
  data.photographerPhone = el.photographerPhone.value.trim();
  data.guestEmails = el.guestEmails.value;
  data.services = selectedServices();
  data.durationMinutes = Number(el.duration.value);
  data.startAt = start.toISOString();
  if (!data.services.length) return { error: 'Choose at least one service.' };
  return { data };
}

function syncLabel(booking) {
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

function renderBookings() {
  const upcoming = state.bookings
    .filter((booking) => new Date(booking.endAt).getTime() >= Date.now() || booking.status === 'cancelled')
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
    cancelButton.disabled = booking.status === 'cancelled' || booking.larkOnly;
    cancelButton.title = booking.larkOnly ? 'Cancel this in Lark.' : '';
    cancelButton.addEventListener('click', () => cancelBooking(booking.id));
    el.bookingList.append(item);
  }
  updateStats();
}

function startBookingEdit(id) {
  const booking = state.bookings.find((item) => item.id === id);
  if (!booking || booking.status === 'cancelled' || booking.larkOnly) {
    const message = booking?.larkOnly ? 'This booking came from Lark. Edit it in Lark for now.' : (booking ? 'Cancelled bookings cannot be edited.' : 'Booking not found.');
    setMessage(el.formMessage, message, 'error');
    return;
  }
  setRoute('/');
  setBookingMode(booking);
  fillBookingForm(booking);
  setMessage(el.formMessage, 'Editing booking. Update booking to save changes.');
  el.bookingForm.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function updateStats() {
  const now = new Date();
  const todayStart = new Date(now.getFullYear(), now.getMonth(), now.getDate()).getTime();
  const todayEnd = todayStart + 24 * 60 * 60 * 1000;
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
  const { data } = await fetchJson('/api/bookings');
  state.bookings = data.bookings || [];
  renderBookings();
  updateAddressSuggestions();
  if (data.larkImportError) {
    el.larkDot.className = 'status-dot offline';
    el.larkTitle.textContent = 'Lark import needs checking';
    el.larkDetail.textContent = data.larkImportError;
  }
}

async function loadClients() {
  const cached = readStored(storageKeys.clients, []).filter((client) => client?.name);
  try {
    const { data } = await fetchJson('/api/clients');
    state.clients = mergeDirectoryRecords(cached, data.clients || []);
  } catch {
    state.clients = cached;
  }
  writeStored(storageKeys.clients, state.clients);
  renderClientOptions();
  renderClientList();
}

async function loadPhotographers() {
  const cached = readStored(storageKeys.photographers, []).filter((photographer) => photographer?.name);
  try {
    const { data } = await fetchJson('/api/photographers');
    state.photographers = mergeDirectoryRecords(cached, data.photographers || []);
  } catch {
    state.photographers = cached;
  }
  writeStored(storageKeys.photographers, state.photographers);
  renderOptions(el.photographerSelect, state.photographers, 'New photographer');
  renderPhotographerList();
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
    el.clientSearch.value = '';
    renderClientOptions(data.client.id);
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
    phone: el.directoryPhotographerPhone.value.trim()
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
    renderOptions(el.photographerSelect, state.photographers, 'New photographer', data.photographer.id);
    renderPhotographerList();
    applySelectedPhotographer();
    el.photographerForm.reset();
    el.directoryPhotographerId.value = '';
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
    if (el.clientSelect.value === id) el.clientSelect.value = '';
    if (el.directoryClientId.value === id) {
      el.clientForm.reset();
      el.directoryClientId.value = '';
    }
    renderClientOptions(el.clientSelect.value);
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
    if (el.photographerSelect.value === id) el.photographerSelect.value = '';
    if (el.directoryPhotographerId.value === id) {
      el.photographerForm.reset();
      el.directoryPhotographerId.value = '';
    }
    renderOptions(el.photographerSelect, state.photographers, 'New photographer', el.photographerSelect.value);
    renderPhotographerList();
    setMessage(el.photographerMessage, 'Photographer deleted.', 'success');
  } catch {
    setMessage(el.photographerMessage, 'Could not reach the photographer directory.', 'error');
  }
}

function bookingSavedMessage(booking, action) {
  if (booking.larkAttendeeStatus === 'needs_review') return `Booking ${action}. Check guest removals in Lark.`;
  if (booking.larkStatus === 'synced') return `Booking ${action} and sent to Lark.`;
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
  if (state.editingBookingId === id) resetBookingForm('Edit cancelled because the booking was cancelled.');
  renderBookings();
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

async function logout() {
  el.logoutButton.disabled = true;
  try {
    await fetch('/api/logout', { method: 'POST' });
  } finally {
    window.location.href = '/login';
  }
}

el.bookingForm.addEventListener('submit', submitBooking);
el.bookingForm.addEventListener('input', saveDraft);
el.bookingForm.addEventListener('change', saveDraft);
el.clientForm.addEventListener('submit', saveDirectoryClient);
el.photographerForm.addEventListener('submit', saveDirectoryPhotographer);
el.refreshButton.addEventListener('click', loadBookings);
el.searchAddressButton.addEventListener('click', searchAddressSuggestions);
el.testLarkButton.addEventListener('click', testLark);
el.previewLarkButton.addEventListener('click', previewLarkEvent);
el.cancelEditButton.addEventListener('click', cancelBookingEdit);
el.logoutButton.addEventListener('click', logout);
el.clientSearch.addEventListener('input', () => renderClientOptions(''));
el.clientSelect.addEventListener('change', applySelectedClient);
el.photographerSelect.addEventListener('change', applySelectedPhotographer);
el.clientEmail.addEventListener('input', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.clientEmail.addEventListener('change', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.clientEmail.addEventListener('blur', () => syncManagedGuestEmail(el.clientEmail, 'syncedClientEmail'));
el.photographerEmail.addEventListener('input', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.photographerEmail.addEventListener('change', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.photographerEmail.addEventListener('blur', () => syncManagedGuestEmail(el.photographerEmail, 'syncedPhotographerEmail'));
el.guestEmails.addEventListener('input', updateInvitationSummary);
el.propertyAddress.addEventListener('input', updateAddressSuggestions);
el.propertyAddress.addEventListener('focus', updateAddressSuggestions);
el.newClientButton.addEventListener('click', () => {
  el.clientForm.reset();
  el.directoryClientId.value = '';
  applyExamplePlaceholders();
  renderClientOptions();
  setMessage(el.clientMessage, '');
});
el.newPhotographerButton.addEventListener('click', () => {
  el.photographerForm.reset();
  el.directoryPhotographerId.value = '';
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
await Promise.all([loadStatus(), loadClients(), loadPhotographers(), loadBookings()]);
restoreDraft();
updateInvitationSummary();
updateAddressSuggestions();
setRoute(window.location.pathname, false);
