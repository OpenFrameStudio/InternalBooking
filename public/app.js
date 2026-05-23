const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const el = {
  bookingForm: $('#bookingForm'),
  bookingList: $('#bookingList'),
  bookingTemplate: $('#bookingTemplate'),
  clientTemplate: $('#clientTemplate'),
  photographerTemplate: $('#photographerTemplate'),
  invoiceTemplate: $('#invoiceTemplate'),
  bookingPage: $('#bookingPage'),
  clientsPage: $('#clientsPage'),
  photographersPage: $('#photographersPage'),
  invoicesPage: $('#invoicesPage'),
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
  invoiceList: $('#invoiceList'),
  invoiceMessage: $('#invoiceMessage'),
  invoiceDraftCount: $('#invoiceDraftCount'),
  invoicePaidCount: $('#invoicePaidCount'),
  invoiceTotalValue: $('#invoiceTotalValue'),
  syncInvoicesButton: $('#syncInvoicesButton'),
  refreshInvoicesButton: $('#refreshInvoicesButton'),
  invoicePrintSheet: $('#invoicePrintSheet'),
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
  modeTitle: $('#bookingModeTitle'),
  appLinks: $$('[data-app-link]')
};

const state = {
  bookings: [],
  clients: [],
  photographers: [],
  invoices: [],
  bookingForm: {
    selectedClientId: '',
    selectedPhotographerId: '',
    clientDropdownOpen: false,
    photographerDropdownOpen: false,
    clientSearchQuery: '',
    clientSearchOpen: false
  },
  editingBookingId: '',
  restoringDraft: false,
  syncedClientEmail: [],
  syncedPhotographerEmail: [],
  user: null
};

const storageKeys = {
  draft: 'openframe.bookingDraft.v2',
  bookings: 'openframe.bookings.v1',
  clients: 'openframe.clients.v1',
  photographers: 'openframe.photographers.v1',
  invoices: 'openframe.invoices.v1',
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
const invoiceDateFormatter = new Intl.DateTimeFormat('en-AU', { day: 'numeric', month: 'short', year: 'numeric' });
const currencyFormatter = new Intl.NumberFormat('en-AU', { style: 'currency', currency: 'AUD' });
const bookingUpcomingGraceMs = 24 * 60 * 60 * 1000;

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

function cacheBookings() {
  writeStored(storageKeys.bookings, state.bookings);
}

function cacheInvoices() {
  if (userCanAccess('invoices')) writeStored(storageKeys.invoices, state.invoices);
}

function hydrateCachedData() {
  const cachedBookings = readStored(storageKeys.bookings, []).filter((booking) => booking?.id);
  if (cachedBookings.length) {
    state.bookings = cachedBookings;
    updateAddressSuggestions();
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
    loadInvoices()
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
        meta: [photographer.email, photographer.phone].filter(Boolean).join(' · ')
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
    item.querySelector('.photographer-phone').textContent = photographer.phone ? `Phone: ${photographer.phone}` : '';
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
  setInitialDateTime();
  updateDurationForServices();
  applyExamplePlaceholders();
  updateInvitationSummary();
  updateAddressSuggestions();
  el.larkPreview.hidden = true;
  setMessage(el.formMessage, message);
}

function setRoute(route, push = true) {
  const nextRoute = ['/clients', '/photographers', '/invoices'].includes(route) ? route : '/bookings';
  el.bookingPage.hidden = nextRoute !== '/bookings';
  el.clientsPage.hidden = nextRoute !== '/clients';
  el.photographersPage.hidden = nextRoute !== '/photographers';
  el.invoicesPage.hidden = nextRoute !== '/invoices';
  $$('[data-route]').forEach((link) => link.classList.toggle('active', link.dataset.route === nextRoute));
  if (push && window.location.pathname !== nextRoute) history.pushState({ route: nextRoute }, '', nextRoute);
  if (nextRoute === '/clients') renderClientList();
  if (nextRoute === '/photographers') renderPhotographerList();
  if (nextRoute === '/invoices') renderInvoices();
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

function updateInvoiceStats() {
  const drafts = state.invoices.filter((invoice) => invoice.status === 'draft');
  const paid = state.invoices.filter((invoice) => invoice.status === 'paid');
  el.invoiceDraftCount.textContent = drafts.length;
  el.invoicePaidCount.textContent = paid.length;
  el.invoiceTotalValue.textContent = formatMoney(drafts.reduce((sum, invoice) => sum + Number(invoice.total || 0), 0));
}

function renderInvoices() {
  if (!userCanAccess('invoices')) return;

  el.invoiceList.innerHTML = '';
  const invoices = [...state.invoices].sort((a, b) => new Date(b.issuedAt) - new Date(a.issuedAt));
  updateInvoiceStats();

  if (!invoices.length) {
    const empty = document.createElement('div');
    empty.className = 'empty-state';
    empty.textContent = 'No invoices yet. Create or sync booking invoices and they will appear here.';
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
    const sendButton = item.querySelector('.send-invoice-button');
    sendButton.hidden = invoice.status === 'void';
    sendButton.addEventListener('click', () => sendInvoice(invoice.id));
    const paidButton = item.querySelector('.paid-invoice-button');
    paidButton.hidden = invoice.status !== 'draft';
    paidButton.addEventListener('click', () => updateInvoiceStatus(invoice.id, 'paid'));
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
    updateAddressSuggestions();
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
    updateBookingFormState({ selectedPhotographerId: data.photographer.id }, { renderSelections: false });
    renderPhotographerOptions();
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
el.refreshInvoicesButton.addEventListener('click', loadInvoices);
el.syncInvoicesButton.addEventListener('click', syncInvoices);
el.searchAddressButton.addEventListener('click', searchAddressSuggestions);
el.testLarkButton.addEventListener('click', testLark);
el.previewLarkButton.addEventListener('click', previewLarkEvent);
el.cancelEditButton.addEventListener('click', cancelBookingEdit);
el.logoutButton.addEventListener('click', logout);
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
el.propertyAddress.addEventListener('input', updateAddressSuggestions);
el.propertyAddress.addEventListener('focus', updateAddressSuggestions);
el.newClientButton.addEventListener('click', () => {
  el.clientForm.reset();
  el.directoryClientId.value = '';
  applyExamplePlaceholders();
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
await loadSession();
applyAppAccess();
hydrateCachedData();
restoreDraft();
updateInvitationSummary();
updateAddressSuggestions();
setRoute(window.location.pathname, false);
refreshLiveData();
