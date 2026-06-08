const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];
const defaultWorkEmployeeId = "faye";
const maxWorkAttachmentCount = 6;
const maxWorkAttachmentBytes = 450_000;
const maxWorkPhotoDimension = 1280;
const workNoticeDismissedKey = "openframe.workNoticeDismissed.v1";

const el = {
  addAssignmentButton: $("#addAssignmentButton"),
  assignmentDialog: $("#assignmentDialog"),
  assignmentDialogTitle: $("#assignmentDialogTitle"),
  assignmentDueDate: $("#assignmentDueDate"),
  assignmentEmployee: $("#assignmentEmployee"),
  assignmentError: $("#assignmentError"),
  assignmentForm: $("#assignmentForm"),
  assignmentId: $("#assignmentId"),
  assignmentList: $("#assignmentList"),
  assignmentNotice: $("#assignmentNotice"),
  assignmentNoticeButton: $("#assignmentNoticeButton"),
  assignmentNoticeDetail: $("#assignmentNoticeDetail"),
  assignmentNoticeTitle: $("#assignmentNoticeTitle"),
  assignmentNotes: $("#assignmentNotes"),
  assignmentPhotoDropzone: $("#assignmentPhotoDropzone"),
  assignmentPhotoList: $("#assignmentPhotoList"),
  assignmentPhotos: $("#assignmentPhotos"),
  assignmentPriority: $("#assignmentPriority"),
  assignmentTitle: $("#assignmentTitle"),
  appLinks: $$("[data-app-link]"),
  bookingSyncPanel: $("#bookingSyncPanel"),
  bookingSyncStatus: $("#bookingSyncStatus"),
  cancelAssignmentButton: $("#cancelAssignmentButton"),
  clearMessagesButton: $("#clearMessagesButton"),
  closeAssignmentDialogButton: $("#closeAssignmentDialogButton"),
  deleteAssignmentButton: $("#deleteAssignmentButton"),
  doneCount: $("#doneCount"),
  dueTodayCount: $("#dueTodayCount"),
  employeeAvailability: $("#employeeAvailability"),
  employeeName: $("#employeeName"),
  employeeRole: $("#employeeRole"),
  employeeStatusDot: $("#employeeStatusDot"),
  enableMessagesButton: $("#enableMessagesButton"),
  fayeDoneCount: $("#fayeDoneCount"),
  fayeDoneThisMonthCount: $("#fayeDoneThisMonthCount"),
  fayeHistoryList: $("#fayeHistoryList"),
  fayeHistorySearchInput: $("#fayeHistorySearchInput"),
  fayeLatestDone: $("#fayeLatestDone"),
  filterButtons: $$("[data-filter]"),
  logoutButton: $("#logoutButton"),
  messageList: $("#messageList"),
  messagePanel: $("#messagePanel"),
  openCount: $("#openCount"),
  sessionPill: $("#sessionPill"),
  summaryLine: $("#summaryLine"),
  syncBookingsButton: $("#syncBookingsButton"),
  toastRegion: $("#toastRegion"),
};

const initialBookingSyncStatus = "Sync upcoming bookings into the work queue.";

const state = {
  employee: {
    id: "faye",
    name: "Faye",
    role: "Editor / Admin",
    availability: "Mon-Fri, 12pm-8pm Australian time",
  },
  employees: [
    {
      id: "faye",
      name: "Faye",
      role: "Editor / Admin",
      availability: "Mon-Fri, 12pm-8pm Australian time",
    },
    {
      id: "test",
      name: "Test",
      role: "Test Employee",
      availability: "Testing account",
    },
  ],
  assignments: [],
  messages: [],
  user: null,
  ui: {
    assignmentDialogOpen: false,
    assignmentDraft: emptyAssignmentDraft(),
    assignmentError: "",
    assignmentNoticeDismissedSignature: readWorkNoticeDismissedSignature(),
    bookingSyncStatus: initialBookingSyncStatus,
    fayeHistoryQuery: "",
    filter: "open",
    loggingOut: false,
    notifyingAssignmentId: "",
    notificationPermission: getNotificationPermission(),
    syncInProgress: false,
    toasts: [],
  },
};

function emptyAssignmentDraft() {
  return {
    id: "",
    employeeId: defaultWorkEmployeeId,
    title: "",
    dueDate: toISODate(new Date()),
    priority: "normal",
    notes: "",
    attachments: [],
  };
}

function defaultEmployeeId() {
  return state?.employees?.[0]?.id || state?.employee?.id || "faye";
}

function employeeById(employeeId) {
  return state.employees.find((employee) => employee.id === employeeId) || state.employee;
}

function getSydneyWorkTime(date = new Date()) {
  const parts = new Intl.DateTimeFormat("en-AU", {
    timeZone: "Australia/Sydney",
    weekday: "short",
    hour: "numeric",
    minute: "numeric",
    hourCycle: "h23",
  }).formatToParts(date);

  const valueFor = (type) => parts.find((part) => part.type === type)?.value || "";
  return {
    weekday: valueFor("weekday"),
    minutes: Number(valueFor("hour")) * 60 + Number(valueFor("minute")),
  };
}

function employeeAvailabilityStatus(employee) {
  const availability = (employee?.availability || "").toLowerCase();
  const hasFayeSchedule =
    availability.includes("mon-fri") &&
    availability.includes("12pm") &&
    availability.includes("8pm");

  if (!hasFayeSchedule) {
    return { online: false, label: "Offline now" };
  }

  const sydneyTime = getSydneyWorkTime();
  const workDayIndex = {
    Mon: 1,
    Tue: 2,
    Wed: 3,
    Thu: 4,
    Fri: 5,
  }[sydneyTime.weekday];
  const { minutes } = sydneyTime;
  const online = Boolean(workDayIndex) && minutes >= 12 * 60 && minutes < 20 * 60;

  return {
    online,
    label: online ? "Online now" : "Offline now",
  };
}

function assignmentDraftFromAssignment(assignment = null) {
  if (!assignment) return emptyAssignmentDraft();

  return {
    id: assignment.id || "",
    employeeId: assignment.employeeId || defaultEmployeeId(),
    title: assignment.title || "",
    dueDate: assignment.dueDate || toISODate(new Date()),
    priority: assignment.priority || "normal",
    notes: assignment.notes || "",
    attachments: Array.isArray(assignment.attachments) ? assignment.attachments : [],
  };
}

function setState(patch = {}) {
  Object.assign(state, patch);
  render();
}

function setUiState(patch = {}) {
  state.ui = {
    ...state.ui,
    ...patch,
  };
  render();
}

function setWorkData(data, uiPatch = {}) {
  setState({
    employee: data.employee || state.employee,
    employees: Array.isArray(data.employees) && data.employees.length ? data.employees : state.employees,
    assignments: Array.isArray(data.assignments) ? data.assignments : state.assignments,
    messages: Array.isArray(data.messages) ? data.messages : state.messages,
    user: data.user || state.user,
    ui: {
      ...state.ui,
      ...uiPatch,
    },
  });
}

function userCanAccess(app) {
  return Boolean(state.user?.apps?.includes(app));
}

function userHasPermission(permission) {
  return Boolean(state.user?.permissions?.includes(permission));
}

function canManageWork() {
  return userHasPermission("manage_work");
}

function canSyncBookings() {
  return userHasPermission("sync_work_bookings");
}

function canViewWorkMessages() {
  return userHasPermission("view_work_messages");
}

function canCompleteAssignment(assignment) {
  return canManageWork() || assignment?.employeeId === state.user?.employeeId;
}

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
    headers: {
      Accept: "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {}),
    },
    ...options,
  });

  const data = await response.json().catch(() => ({}));

  if (response.status === 401) {
    window.location.href = "/login";
    throw new Error("Log in first.");
  }

  if (!response.ok) {
    throw new Error((data.errors || ["Something went wrong."]).join(" "));
  }

  return data;
}

function jsonRequest(method, body = null) {
  return {
    method,
    ...(body ? { body: JSON.stringify(body) } : {}),
  };
}

function assignmentEndpoint(id, action = "") {
  const base = `/api/work/assignments/${encodeURIComponent(id)}`;
  return action ? `${base}/${action}` : base;
}

const workApi = {
  load: () => apiFetch("/api/work"),
  syncBookings: () => apiFetch("/api/work/sync-bookings", jsonRequest("POST")),
  createAssignment: (payload) => apiFetch("/api/work/assignments", jsonRequest("POST", payload)),
  updateAssignment: (id, payload) => apiFetch(assignmentEndpoint(id), jsonRequest("PUT", payload)),
  completeAssignment: (id) => apiFetch(assignmentEndpoint(id, "complete"), jsonRequest("POST")),
  notifyAssignment: (id) => apiFetch(assignmentEndpoint(id, "notify"), jsonRequest("POST")),
  reopenAssignment: (id) => apiFetch(assignmentEndpoint(id, "reopen"), jsonRequest("POST")),
  deleteAssignment: (id) => apiFetch(assignmentEndpoint(id), jsonRequest("DELETE")),
  clearMessages: () => apiFetch("/api/work/messages", jsonRequest("DELETE")),
};

async function loadWork() {
  setWorkData(await workApi.load());
}

function parseISODate(value) {
  if (!value) return new Date(NaN);
  const [year, month, day] = value.split("-").map(Number);
  return new Date(year, month - 1, day);
}

function toISODate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function formatDate(date, options) {
  return new Intl.DateTimeFormat("en-AU", options).format(date);
}

function formatMessageTime(value) {
  return new Intl.DateTimeFormat("en-AU", {
    day: "numeric",
    month: "short",
    hour: "numeric",
    minute: "2-digit",
  }).format(new Date(value));
}

function formatHistoryDate(value) {
  if (!value) return "No completion date";
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "No completion date";
  return new Intl.DateTimeFormat("en-AU", {
    day: "numeric",
    month: "short",
    year: "numeric",
  }).format(date);
}

function formatFileSize(bytes) {
  const size = Number(bytes || 0);
  if (size < 1024) return `${Math.max(0, Math.round(size))} B`;
  if (size < 1024 * 1024) return `${Math.round(size / 102.4) / 10} KB`;
  return `${Math.round(size / 1024 / 102.4) / 10} MB`;
}

function workAttachmentByteSize(dataUrl) {
  const base64 = String(dataUrl || "").split(",")[1] || "";
  return Math.ceil(base64.length * 3 / 4);
}

function isToday(dateISO) {
  return dateISO === toISODate(new Date());
}

function isOverdue(assignment) {
  return assignment.status !== "done" && parseISODate(assignment.dueDate) < parseISODate(toISODate(new Date()));
}

function assignmentCompletedTime(assignment) {
  const value = assignment.completedAt || assignment.updatedAt || assignment.createdAt || "";
  const time = new Date(value).getTime();
  return Number.isNaN(time) ? 0 : time;
}

function fayePastAssignments() {
  return [...state.assignments]
    .filter((assignment) => assignment.employeeId === defaultWorkEmployeeId && assignment.status === "done")
    .sort((a, b) => assignmentCompletedTime(b) - assignmentCompletedTime(a));
}

function assignmentSearchText(assignment) {
  const assignee = employeeById(assignment.employeeId);
  return [
    assignment.title,
    assignment.notes,
    assignment.priority,
    assignment.status,
    assignment.dueDate,
    assignment.completedAt,
    assignment.createdAt,
    assignment.updatedAt,
    assignment.source,
    assignment.sourceId,
    assignee?.name,
  ].join(" ").toLowerCase();
}

function priorityWeight(priority) {
  return { high: 3, normal: 2, low: 1 }[priority] || 2;
}

function readWorkNoticeDismissedSignature() {
  try {
    return localStorage.getItem(workNoticeDismissedKey) || "";
  } catch {
    return "";
  }
}

function writeWorkNoticeDismissedSignature(signature) {
  try {
    if (signature) localStorage.setItem(workNoticeDismissedKey, signature);
    else localStorage.removeItem(workNoticeDismissedKey);
  } catch {
    // This only controls whether the small in-app banner is hidden.
  }
}

function getOpenAssignments() {
  return [...state.assignments]
    .filter((assignment) => assignment.status !== "done")
    .sort((a, b) => {
      const overdueSort = Number(isOverdue(b)) - Number(isOverdue(a));
      if (overdueSort) return overdueSort;
      if (a.dueDate !== b.dueDate) return a.dueDate.localeCompare(b.dueDate);
      const prioritySort = priorityWeight(b.priority) - priorityWeight(a.priority);
      if (prioritySort) return prioritySort;
      return new Date(b.createdAt || 0) - new Date(a.createdAt || 0);
    });
}

function workNoticeSignature(assignments) {
  return assignments
    .map((assignment) => `${assignment.id}:${assignment.updatedAt || assignment.createdAt || ""}`)
    .sort()
    .join("|");
}

function getNotificationPermission() {
  if (!("Notification" in window)) return "unavailable";
  return Notification.permission;
}

function render() {
  renderProfile();
  renderMetrics();
  renderFayeHistory();
  renderAssignmentNotice();
  renderRoleAccess();
  renderFilters();
  renderMessages();
  renderAssignments();
  renderAssignmentDialog();
  renderBookingSync();
  renderNotificationButton();
  renderToasts();
  refreshIcons();
}

function renderProfile() {
  const availabilityStatus = employeeAvailabilityStatus(state.employee);
  el.sessionPill.textContent = state.user?.label || "Logged in";
  el.employeeName.textContent = state.employee.name;
  el.employeeRole.textContent = state.employee.role;
  el.employeeAvailability.textContent = state.employee.availability;
  el.employeeStatusDot.classList.toggle("ready", availabilityStatus.online);
  el.employeeStatusDot.classList.toggle("offline", !availabilityStatus.online);
  el.employeeStatusDot.title = availabilityStatus.label;
  el.employeeStatusDot.setAttribute("aria-label", availabilityStatus.label);
  el.logoutButton.disabled = state.ui.loggingOut;
}

function renderRoleAccess() {
  el.appLinks.forEach((link) => {
    link.hidden = !userCanAccess(link.dataset.appLink);
  });
  el.addAssignmentButton.hidden = !canManageWork();
  el.syncBookingsButton.hidden = !canSyncBookings();
  el.bookingSyncPanel.hidden = !canSyncBookings();
  el.enableMessagesButton.hidden = !canViewWorkMessages();
  el.clearMessagesButton.hidden = !canViewWorkMessages();
}

function renderMetrics() {
  const open = getOpenAssignments();
  const done = state.assignments.filter((assignment) => assignment.status === "done");
  const dueToday = open.filter((assignment) => isToday(assignment.dueDate));
  const overdue = open.filter(isOverdue);

  el.openCount.textContent = open.length;
  el.doneCount.textContent = done.length;
  el.dueTodayCount.textContent = dueToday.length;
  el.summaryLine.textContent = `${open.length} open - ${dueToday.length} due today - ${overdue.length} overdue`;
}

function renderFayeHistory() {
  const completed = fayePastAssignments();
  const query = state.ui.fayeHistoryQuery.trim().toLowerCase();
  const visible = query
    ? completed.filter((assignment) => assignmentSearchText(assignment).includes(query))
    : completed;
  const monthKey = toISODate(new Date()).slice(0, 7);
  const completedThisMonth = completed.filter((assignment) => {
    const value = assignment.completedAt || assignment.updatedAt || "";
    return value.slice(0, 7) === monthKey;
  });

  el.fayeDoneCount.textContent = completed.length;
  el.fayeDoneThisMonthCount.textContent = completedThisMonth.length;
  el.fayeLatestDone.textContent = completed.length
    ? formatHistoryDate(completed[0].completedAt || completed[0].updatedAt)
    : "No finished work yet";

  if (el.fayeHistorySearchInput.value !== state.ui.fayeHistoryQuery) {
    el.fayeHistorySearchInput.value = state.ui.fayeHistoryQuery;
  }

  if (!completed.length) {
    el.fayeHistoryList.innerHTML = `
      <div class="history-empty">
        <strong>No past work yet</strong>
        <span>Finished work assigned to Faye will appear here.</span>
      </div>
    `;
    return;
  }

  if (!visible.length) {
    el.fayeHistoryList.innerHTML = `
      <div class="history-empty">
        <strong>No matches</strong>
        <span>Try a different title, note, or date.</span>
      </div>
    `;
    return;
  }

  el.fayeHistoryList.replaceChildren(...visible.slice(0, 8).map((assignment) => {
    const item = document.createElement("button");
    item.className = "history-item";
    item.type = "button";
    item.dataset.historyAssignment = assignment.id;
    item.innerHTML = `
      <strong></strong>
      <span class="history-date"></span>
      <span class="history-notes"></span>
    `;
    item.querySelector("strong").textContent = assignment.title;
    item.querySelector(".history-date").textContent = formatHistoryDate(assignment.completedAt || assignment.updatedAt);
    item.querySelector(".history-notes").textContent = assignment.notes || "No details added.";
    return item;
  }));
}

function renderAssignmentNotice() {
  const open = getOpenAssignments();
  const signature = workNoticeSignature(open);
  el.assignmentNotice.hidden = open.length === 0 || signature === state.ui.assignmentNoticeDismissedSignature;

  if (!open.length || signature === state.ui.assignmentNoticeDismissedSignature) {
    el.assignmentNoticeTitle.textContent = "New work waiting";
    el.assignmentNoticeDetail.textContent = "";
    return;
  }

  const nextAssignment = open[0];
  const countLabel = open.length === 1 ? "1 new assignment" : `${open.length} new assignments`;
  const audienceLabel = state.user?.role === "employee" ? "assigned to you" : "in the queue";
  const dueLabel = isOverdue(nextAssignment)
    ? "overdue"
    : isToday(nextAssignment.dueDate)
      ? "due today"
      : `due ${formatDate(parseISODate(nextAssignment.dueDate), { weekday: "short", day: "numeric", month: "short" })}`;
  const priorityLabel = `${nextAssignment.priority.charAt(0).toUpperCase()}${nextAssignment.priority.slice(1)} priority`;

  el.assignmentNoticeTitle.textContent = `${countLabel} ${audienceLabel}`;
  el.assignmentNoticeDetail.textContent = `Next: ${nextAssignment.title} - ${dueLabel} - ${priorityLabel}`;
}

function renderFilters() {
  el.filterButtons.forEach((button) => {
    button.checked = button.dataset.filter === state.ui.filter;
  });
}

function getVisibleAssignments() {
  return [...state.assignments]
    .filter((assignment) => {
      if (state.ui.filter === "faye-done") {
        return assignment.employeeId === defaultWorkEmployeeId && assignment.status === "done";
      }
      if (state.ui.filter === "done") return assignment.status === "done";
      if (state.ui.filter === "overdue") return isOverdue(assignment);
      if (state.ui.filter === "open") return assignment.status !== "done";
      return true;
    })
    .sort((a, b) => {
      if (state.ui.filter === "faye-done" || state.ui.filter === "done") {
        return assignmentCompletedTime(b) - assignmentCompletedTime(a);
      }
      if (a.status !== b.status) return a.status === "done" ? 1 : -1;
      if (a.dueDate !== b.dueDate) return a.dueDate.localeCompare(b.dueDate);
      return priorityWeight(b.priority) - priorityWeight(a.priority);
    });
}

function renderAssignments() {
  const assignments = getVisibleAssignments();

  if (!assignments.length) {
    el.assignmentList.innerHTML = `
      <div class="empty-state">
        <strong>No work assigned</strong>
        <span></span>
      </div>
    `;
    const emptyMessages = {
      done: "No finished work yet.",
      "faye-done": "No finished work assigned to Faye yet.",
      overdue: "No overdue work.",
      open: "Ready for the next assignment."
    };
    el.assignmentList.querySelector("span").textContent = emptyMessages[state.ui.filter] || "No work assigned.";
    return;
  }

  el.assignmentList.replaceChildren(...assignments.map(renderAssignmentCard));
}

function renderAssignmentCard(assignment) {
  const assignee = employeeById(assignment.employeeId);
  const card = document.createElement("article");
  card.className = `assignment-card ${assignment.status === "done" ? "done" : "open"} ${
    isOverdue(assignment) ? "overdue" : ""
  }`;
  card.dataset.assignmentId = assignment.id;
  card.innerHTML = `
    <div class="assignment-content">
      <div class="assignment-kicker">
        <span class="status-pill"></span>
        <span class="priority-pill"></span>
      </div>
      <h3></h3>
      <p class="assignment-notes"></p>
      <div class="assignment-meta">
        <span><i data-lucide="user-round"></i><strong></strong></span>
        <span><i data-lucide="calendar-days"></i><em></em></span>
      </div>
    </div>
    <div class="assignment-actions"></div>
  `;

  card.querySelector(".status-pill").textContent = assignment.status === "done" ? "Finished" : "Open";
  card.querySelector(".priority-pill").textContent = assignment.priority;
  card.querySelector("h3").textContent = assignment.title;
  card.querySelector(".assignment-notes").textContent = assignment.notes || "No details added.";
  card.querySelector(".assignment-meta strong").textContent = assignee?.name || "Employee";
  card.querySelector(".assignment-meta em").textContent = `Due ${formatDate(parseISODate(assignment.dueDate), {
    weekday: "short",
    day: "numeric",
    month: "short",
  })}`;
  renderAssignmentAttachments(card, assignment.attachments || []);

  const actions = card.querySelector(".assignment-actions");

  if (assignment.status !== "done" && canCompleteAssignment(assignment)) {
    const completeButton = document.createElement("button");
    completeButton.className = "ghost-button complete-button";
    completeButton.type = "button";
    completeButton.textContent = state.user?.role === "employee" ? "Mark Finished" : "Finished";
    completeButton.dataset.completeAssignment = assignment.id;
    actions.append(completeButton);
  } else if (assignment.status === "done" && canManageWork()) {
    const reopenButton = document.createElement("button");
    reopenButton.className = "ghost-button complete-button";
    reopenButton.type = "button";
    reopenButton.textContent = "Reopen";
    reopenButton.dataset.reopenAssignment = assignment.id;
    actions.append(reopenButton);
  } else {
    const finishedLabel = document.createElement("span");
    finishedLabel.className = "finished-label";
    finishedLabel.textContent = "Finished";
    actions.append(finishedLabel);
  }

  if (assignment.status !== "done" && canManageWork()) {
    const notifyButton = document.createElement("button");
    const isNotifying = state.ui.notifyingAssignmentId === assignment.id;
    notifyButton.className = "ghost-button notify-button";
    notifyButton.type = "button";
    notifyButton.disabled = isNotifying;
    notifyButton.title = `Notify ${assignee?.name || "employee"}`;
    notifyButton.setAttribute("aria-label", `Notify ${assignee?.name || "employee"}`);
    notifyButton.dataset.notifyAssignment = assignment.id;
    notifyButton.innerHTML = `<i data-lucide="bell-ring"></i><span>${isNotifying ? "Sending" : "Notify"}</span>`;
    actions.append(notifyButton);
  }

  if (canManageWork()) {
    const editButton = document.createElement("button");
    editButton.className = "ghost-button icon-ghost";
    editButton.type = "button";
    editButton.title = "Edit work";
    editButton.setAttribute("aria-label", "Edit work");
    editButton.dataset.editAssignment = assignment.id;
    editButton.innerHTML = `<i data-lucide="settings-2"></i>`;
    actions.append(editButton);
  }

  return card;
}

function renderAssignmentAttachments(card, attachments) {
  if (!attachments.length) return;

  const gallery = document.createElement("div");
  gallery.className = "assignment-photo-gallery";
  gallery.replaceChildren(...attachments.map((attachment) => {
    const link = document.createElement("a");
    link.href = attachment.dataUrl;
    link.target = "_blank";
    link.rel = "noopener";
    link.download = attachment.name || "work-photo.jpg";
    link.title = attachment.name || "Open photo";
    link.innerHTML = `<img alt="" loading="lazy" />`;
    link.querySelector("img").src = attachment.dataUrl;
    link.querySelector("img").alt = attachment.name || "Work photo";
    return link;
  }));

  card.querySelector(".assignment-content").insertBefore(gallery, card.querySelector(".assignment-meta"));
}

function renderMessages() {
  if (!canViewWorkMessages()) {
    el.messagePanel.hidden = true;
    el.messageList.replaceChildren();
    return;
  }

  const messages = state.messages.slice(0, 6);
  el.messagePanel.hidden = messages.length === 0;

  if (!messages.length) {
    el.messageList.replaceChildren();
    return;
  }

  el.messageList.replaceChildren(...messages.map((message) => {
    const item = document.createElement("div");
    item.className = "message-row";
    item.innerHTML = `<strong></strong><span></span>`;
    item.querySelector("strong").textContent = message.text;
    item.querySelector("span").textContent = formatMessageTime(message.createdAt);
    return item;
  }));
}

function renderAssignmentDialog() {
  const draft = state.ui.assignmentDraft;
  const wasOpen = el.assignmentDialog.open;

  el.assignmentDialogTitle.textContent = draft.id ? "Edit Work" : "Assign Work";
  el.assignmentId.value = draft.id;
  renderEmployeeOptions(draft.employeeId);
  el.assignmentTitle.value = draft.title;
  el.assignmentDueDate.value = draft.dueDate;
  el.assignmentPriority.value = draft.priority;
  el.assignmentNotes.value = draft.notes;
  el.assignmentError.textContent = state.ui.assignmentError;
  el.deleteAssignmentButton.hidden = !draft.id;
  renderAssignmentPhotoList();

  if (state.ui.assignmentDialogOpen && !wasOpen) {
    el.assignmentDialog.showModal();
    el.assignmentTitle.focus();
  }

  if (!state.ui.assignmentDialogOpen && wasOpen) {
    el.assignmentDialog.close();
  }
}

function renderAssignmentPhotoList() {
  const attachments = state.ui.assignmentDraft.attachments || [];

  if (!attachments.length) {
    el.assignmentPhotoList.innerHTML = `<p class="work-photo-empty">No photos attached.</p>`;
    return;
  }

  el.assignmentPhotoList.replaceChildren(...attachments.map((attachment) => {
    const item = document.createElement("figure");
    item.className = "work-photo-chip";
    item.innerHTML = `
      <img alt="" loading="lazy" />
      <figcaption>
        <strong></strong>
        <span></span>
      </figcaption>
      <button class="icon-button" type="button" title="Remove photo" aria-label="Remove photo">
        <i data-lucide="x"></i>
      </button>
    `;
    item.querySelector("img").src = attachment.dataUrl;
    item.querySelector("img").alt = attachment.name || "Work photo";
    item.querySelector("strong").textContent = attachment.name || "Work photo";
    item.querySelector("span").textContent = formatFileSize(attachment.size || workAttachmentByteSize(attachment.dataUrl));
    item.querySelector("button").dataset.removeAssignmentPhoto = attachment.id;
    return item;
  }));
}

function renderEmployeeOptions(selectedEmployeeId) {
  const employees = state.employees.length ? state.employees : [state.employee];
  el.assignmentEmployee.replaceChildren(...employees.map((employee) => {
    const option = document.createElement("option");
    option.value = employee.id;
    option.textContent = employee.name || "Employee";
    option.selected = employee.id === selectedEmployeeId;
    return option;
  }));
}

function renderBookingSync() {
  el.syncBookingsButton.disabled = state.ui.syncInProgress || !canSyncBookings();
  el.bookingSyncStatus.textContent = state.ui.bookingSyncStatus;
}

function renderNotificationButton() {
  if (state.ui.notificationPermission === "unavailable") {
    el.enableMessagesButton.disabled = true;
    el.enableMessagesButton.title = "Browser messages unavailable";
    return;
  }

  const enabled = state.ui.notificationPermission === "granted";
  el.enableMessagesButton.disabled = !canViewWorkMessages();
  el.enableMessagesButton.classList.toggle("active", enabled);
  el.enableMessagesButton.title = enabled ? "Messages enabled" : "Enable messages";
  el.enableMessagesButton.setAttribute("aria-label", enabled ? "Messages enabled" : "Enable messages");
}

function renderToasts() {
  el.toastRegion.replaceChildren(...state.ui.toasts.map((toast) => {
    const item = document.createElement("div");
    item.className = "toast";
    item.dataset.toastId = toast.id;
    item.textContent = toast.text;
    return item;
  }));
}

function openAssignmentDialog(assignment = null) {
  if (!canManageWork()) return;

  setUiState({
    assignmentDialogOpen: true,
    assignmentDraft: assignmentDraftFromAssignment(assignment),
    assignmentError: "",
  });
}

function closeAssignmentDialog() {
  setUiState({
    assignmentDialogOpen: false,
    assignmentDraft: emptyAssignmentDraft(),
    assignmentError: "",
  });
}

function updateAssignmentDraft(event) {
  const fieldById = {
    assignmentEmployee: "employeeId",
    assignmentTitle: "title",
    assignmentDueDate: "dueDate",
    assignmentPriority: "priority",
    assignmentNotes: "notes",
  };
  const field = fieldById[event.target.id];
  if (!field || !state.ui.assignmentDialogOpen) return;

  setUiState({
    assignmentDraft: {
      ...state.ui.assignmentDraft,
      [field]: event.target.value,
    },
    assignmentError: "",
  });
}

function assignmentPayloadFromDraft(draft) {
  return {
    employeeId: draft.employeeId || defaultEmployeeId(),
    title: draft.title.trim(),
    dueDate: draft.dueDate,
    priority: draft.priority,
    notes: draft.notes.trim(),
    attachments: Array.isArray(draft.attachments) ? draft.attachments : [],
  };
}

function readFileAsDataUrl(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.addEventListener("load", () => resolve(reader.result));
    reader.addEventListener("error", () => reject(new Error("Photo could not be read.")));
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl) {
  return new Promise((resolve, reject) => {
    const image = new Image();
    image.addEventListener("load", () => resolve(image));
    image.addEventListener("error", () => reject(new Error("Photo could not be loaded.")));
    image.src = dataUrl;
  });
}

function canvasToDataUrl(canvas, quality) {
  return canvas.toDataURL("image/jpeg", quality);
}

async function compressPhoto(file) {
  if (!/^image\/(jpeg|jpg|png|webp)$/i.test(file.type)) {
    throw new Error("Only JPEG, PNG, or WebP photos can be attached.");
  }

  const originalDataUrl = await readFileAsDataUrl(file);
  const image = await loadImage(originalDataUrl);
  const scale = Math.min(1, maxWorkPhotoDimension / Math.max(image.width, image.height));
  const width = Math.max(1, Math.round(image.width * scale));
  const height = Math.max(1, Math.round(image.height * scale));
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  const context = canvas.getContext("2d");
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, width, height);
  context.drawImage(image, 0, 0, width, height);

  for (const quality of [0.74, 0.64, 0.52]) {
    const dataUrl = canvasToDataUrl(canvas, quality);
    if (workAttachmentByteSize(dataUrl) <= maxWorkAttachmentBytes) return dataUrl;
  }

  throw new Error(`${file.name} is too large. Please upload a smaller photo.`);
}

function attachmentId() {
  if (typeof crypto !== "undefined" && crypto.randomUUID) return crypto.randomUUID();
  return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
}

async function attachmentFromFile(file) {
  const dataUrl = await compressPhoto(file);
  return {
    id: attachmentId(),
    name: file.name || "Work photo.jpg",
    type: "image/jpeg",
    size: workAttachmentByteSize(dataUrl),
    dataUrl,
    createdAt: new Date().toISOString(),
  };
}

async function processAssignmentPhotos(files) {
  if (!state.ui.assignmentDialogOpen) return;

  const photoFiles = [...files].filter((file) => file?.type?.startsWith("image/"));
  if (!photoFiles.length) {
    setUiState({ assignmentError: "Drop or choose JPEG, PNG, or WebP photos." });
    return;
  }

  const current = state.ui.assignmentDraft.attachments || [];
  if (current.length + photoFiles.length > maxWorkAttachmentCount) {
    setUiState({ assignmentError: `Attach up to ${maxWorkAttachmentCount} photos.` });
    return;
  }

  try {
    const attachments = [];
    for (const file of photoFiles) {
      attachments.push(await attachmentFromFile(file));
    }

    setUiState({
      assignmentDraft: {
        ...state.ui.assignmentDraft,
        attachments: [...current, ...attachments],
      },
      assignmentError: "",
    });
  } catch (error) {
    setUiState({ assignmentError: error.message });
  }
}

async function addAssignmentPhotos(event) {
  const files = [...event.target.files];
  event.target.value = "";
  await processAssignmentPhotos(files);
}

function setPhotoDropActive(active) {
  el.assignmentPhotoDropzone.classList.toggle("drag-over", active);
}

function handlePhotoDrag(event) {
  event.preventDefault();
  if (!state.ui.assignmentDialogOpen) return;

  event.dataTransfer.dropEffect = "copy";
  setPhotoDropActive(true);
}

function handlePhotoDragLeave(event) {
  if (event.currentTarget.contains(event.relatedTarget)) return;
  setPhotoDropActive(false);
}

async function handlePhotoDrop(event) {
  event.preventDefault();
  setPhotoDropActive(false);
  await processAssignmentPhotos(event.dataTransfer.files);
}

function removeAssignmentPhoto(id) {
  const attachments = state.ui.assignmentDraft.attachments || [];
  setUiState({
    assignmentDraft: {
      ...state.ui.assignmentDraft,
      attachments: attachments.filter((attachment) => attachment.id !== id),
    },
    assignmentError: "",
  });
}

async function handleAssignmentSubmit(event) {
  event.preventDefault();
  if (!canManageWork()) return;

  const draft = state.ui.assignmentDraft;
  const payload = assignmentPayloadFromDraft(draft);

  if (!payload.title || !payload.dueDate) {
    setUiState({ assignmentError: "Work and due date are required." });
    return;
  }

  try {
    const data = draft.id
      ? await workApi.updateAssignment(draft.id, payload)
      : await workApi.createAssignment(payload);
    setWorkData(data, {
      assignmentDialogOpen: false,
      assignmentDraft: emptyAssignmentDraft(),
      assignmentError: "",
    });
    if (!draft.id && data.workInviteMessage) showToast(data.workInviteMessage);
  } catch (error) {
    setUiState({ assignmentError: error.message });
  }
}

async function completeAssignment(id) {
  const assignmentBeforeUpdate = state.assignments.find((item) => item.id === id);
  if (!canCompleteAssignment(assignmentBeforeUpdate)) return;

  const data = await workApi.completeAssignment(id);
  setWorkData(data);
  const assignment = data.assignments?.find((item) => item.id === id);
  if (assignment) notifyCompletion(assignment);
  if (data.workCompletionNotificationMessage) showToast(data.workCompletionNotificationMessage);
}

async function notifyAssignment(id) {
  if (!canManageWork()) return;

  setUiState({ notifyingAssignmentId: id });
  try {
    const data = await workApi.notifyAssignment(id);
    setWorkData(data, { notifyingAssignmentId: "" });
    if (data.workInviteMessage || data.workLarkNotificationMessage) {
      showToast(data.workInviteMessage || data.workLarkNotificationMessage);
    }
  } catch (error) {
    setUiState({ notifyingAssignmentId: "" });
    showToast(error.message);
  }
}

async function reopenAssignment(id) {
  if (!canManageWork()) return;
  setWorkData(await workApi.reopenAssignment(id));
}

async function deleteCurrentAssignment() {
  if (!canManageWork()) return;

  const id = state.ui.assignmentDraft.id;
  if (!id) return;
  setWorkData(await workApi.deleteAssignment(id), {
    assignmentDialogOpen: false,
    assignmentDraft: emptyAssignmentDraft(),
    assignmentError: "",
  });
}

async function clearMessages() {
  if (!canViewWorkMessages()) return;
  setWorkData(await workApi.clearMessages());
}

function viewOpenAssignments() {
  const signature = workNoticeSignature(getOpenAssignments());
  writeWorkNoticeDismissedSignature(signature);
  setUiState({ filter: "open", assignmentNoticeDismissedSignature: signature });
  window.requestAnimationFrame(() => {
    el.assignmentList.scrollIntoView({ behavior: "smooth", block: "start" });
  });
}

async function syncBookings() {
  if (!canSyncBookings()) return;

  setUiState({
    syncInProgress: true,
    bookingSyncStatus: "Checking upcoming bookings...",
  });

  try {
    const data = await workApi.syncBookings();
    const message = data.syncable === 0
      ? "No upcoming bookings found to assign."
      : `Synced ${data.syncable} upcoming bookings: ${data.created} new, ${data.updated} updated.`;
    const fullMessage = data.workInviteMessage ? `${message} ${data.workInviteMessage}` : message;

    setWorkData(data, {
      bookingSyncStatus: fullMessage,
      filter: "open",
      syncInProgress: false,
    });
    showToast(fullMessage);
  } catch (error) {
    setUiState({
      bookingSyncStatus: error.message,
      syncInProgress: false,
    });
    showToast(error.message);
  }
}

function notifyCompletion(assignment) {
  const assignee = employeeById(assignment.employeeId);
  const assigneeName = assignee?.name || state.employee.name;
  showToast(`${assigneeName} finished: ${assignment.title}`);

  if (!("Notification" in window) || state.ui.notificationPermission !== "granted") return;

  try {
    new Notification("Work finished", { body: `${assigneeName} finished: ${assignment.title}` });
  } catch {
    showToast("Browser message could not be shown.");
  }
}

function showToast(text) {
  const id = `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  setUiState({
    toasts: [...state.ui.toasts, { id, text }],
  });
  window.setTimeout(() => {
    setUiState({
      toasts: state.ui.toasts.filter((toast) => toast.id !== id),
    });
  }, 5200);
}

async function enableNotifications() {
  if (!canViewWorkMessages()) return;

  if (!("Notification" in window)) {
    showToast("Browser messages are not available here.");
    return;
  }

  const permission = await Notification.requestPermission();
  setUiState({ notificationPermission: permission });
  showToast(permission === "granted" ? "Messages enabled." : "Messages were not enabled.");
}

async function logout() {
  setUiState({ loggingOut: true });
  try {
    await fetch("/api/logout", { method: "POST", credentials: "include" });
  } finally {
    window.location.href = "/login";
  }
}

function refreshIcons() {
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

function wireEvents() {
  el.logoutButton.addEventListener("click", logout);
  el.addAssignmentButton.addEventListener("click", () => openAssignmentDialog());
  el.assignmentNoticeButton.addEventListener("click", viewOpenAssignments);
  el.syncBookingsButton.addEventListener("click", syncBookings);
  el.enableMessagesButton.addEventListener("click", enableNotifications);
  el.clearMessagesButton.addEventListener("click", clearMessages);

  el.filterButtons.forEach((button) => {
    button.addEventListener("change", () => {
      setUiState({ filter: button.dataset.filter });
    });
  });

  el.fayeHistorySearchInput.addEventListener("input", () => {
    setUiState({ fayeHistoryQuery: el.fayeHistorySearchInput.value });
  });

  el.fayeHistoryList.addEventListener("click", (event) => {
    const target = event.target.closest("[data-history-assignment]");
    if (!target) return;

    setUiState({ filter: "faye-done" });
    window.requestAnimationFrame(() => {
      const escapedId = CSS.escape(target.dataset.historyAssignment);
      const card = el.assignmentList.querySelector(`[data-assignment-id="${escapedId}"]`);
      card?.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  });

  el.assignmentList.addEventListener("click", async (event) => {
    const completeTarget = event.target.closest("[data-complete-assignment]");
    if (completeTarget) {
      await completeAssignment(completeTarget.dataset.completeAssignment);
      return;
    }

    const notifyTarget = event.target.closest("[data-notify-assignment]");
    if (notifyTarget) {
      await notifyAssignment(notifyTarget.dataset.notifyAssignment);
      return;
    }

    const reopenTarget = event.target.closest("[data-reopen-assignment]");
    if (reopenTarget) {
      await reopenAssignment(reopenTarget.dataset.reopenAssignment);
      return;
    }

    const editTarget = event.target.closest("[data-edit-assignment]");
    if (editTarget) {
      const assignment = state.assignments.find((item) => item.id === editTarget.dataset.editAssignment);
      if (assignment) openAssignmentDialog(assignment);
    }
  });

  el.assignmentForm.addEventListener("input", updateAssignmentDraft);
  el.assignmentForm.addEventListener("change", updateAssignmentDraft);
  el.assignmentForm.addEventListener("submit", handleAssignmentSubmit);
  el.assignmentPhotos.addEventListener("change", addAssignmentPhotos);
  el.assignmentPhotoDropzone.addEventListener("dragenter", handlePhotoDrag);
  el.assignmentPhotoDropzone.addEventListener("dragover", handlePhotoDrag);
  el.assignmentPhotoDropzone.addEventListener("dragleave", handlePhotoDragLeave);
  el.assignmentPhotoDropzone.addEventListener("drop", handlePhotoDrop);
  el.assignmentPhotoList.addEventListener("click", (event) => {
    const removeButton = event.target.closest("[data-remove-assignment-photo]");
    if (removeButton) removeAssignmentPhoto(removeButton.dataset.removeAssignmentPhoto);
  });
  el.assignmentDialog.addEventListener("cancel", (event) => {
    event.preventDefault();
    closeAssignmentDialog();
  });
  el.assignmentDialog.addEventListener("close", () => {
    if (state.ui.assignmentDialogOpen) closeAssignmentDialog();
  });
  el.deleteAssignmentButton.addEventListener("click", deleteCurrentAssignment);
  el.cancelAssignmentButton.addEventListener("click", closeAssignmentDialog);
  el.closeAssignmentDialogButton.addEventListener("click", closeAssignmentDialog);
}

wireEvents();
render();
loadWork().catch((error) => showToast(error.message));
window.setInterval(renderProfile, 60 * 1000);
window.addEventListener("load", refreshIcons);
