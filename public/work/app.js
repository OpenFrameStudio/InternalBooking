const $ = (selector) => document.querySelector(selector);
const $$ = (selector) => [...document.querySelectorAll(selector)];

const el = {
  addAssignmentButton: $("#addAssignmentButton"),
  assignmentDialog: $("#assignmentDialog"),
  assignmentDialogTitle: $("#assignmentDialogTitle"),
  assignmentDueDate: $("#assignmentDueDate"),
  assignmentError: $("#assignmentError"),
  assignmentForm: $("#assignmentForm"),
  assignmentId: $("#assignmentId"),
  assignmentList: $("#assignmentList"),
  assignmentNotes: $("#assignmentNotes"),
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
  enableMessagesButton: $("#enableMessagesButton"),
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

const initialBookingSyncStatus = "Sync upcoming bookings into Faye's work queue.";

const state = {
  employee: {
    id: "faye",
    name: "Faye",
    role: "Editor / Admin",
    availability: "Mon-Fri, 12pm-8pm Australian time",
  },
  assignments: [],
  messages: [],
  user: null,
  ui: {
    assignmentDialogOpen: false,
    assignmentDraft: emptyAssignmentDraft(),
    assignmentError: "",
    bookingSyncStatus: initialBookingSyncStatus,
    filter: "open",
    loggingOut: false,
    notificationPermission: getNotificationPermission(),
    syncInProgress: false,
    toasts: [],
  },
};

function emptyAssignmentDraft() {
  return {
    id: "",
    title: "",
    dueDate: toISODate(new Date()),
    priority: "normal",
    notes: "",
  };
}

function assignmentDraftFromAssignment(assignment = null) {
  if (!assignment) return emptyAssignmentDraft();

  return {
    id: assignment.id || "",
    title: assignment.title || "",
    dueDate: assignment.dueDate || toISODate(new Date()),
    priority: assignment.priority || "normal",
    notes: assignment.notes || "",
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
  reopenAssignment: (id) => apiFetch(assignmentEndpoint(id, "reopen"), jsonRequest("POST")),
  deleteAssignment: (id) => apiFetch(assignmentEndpoint(id), jsonRequest("DELETE")),
  clearMessages: () => apiFetch("/api/work/messages", jsonRequest("DELETE")),
};

async function loadWork() {
  setWorkData(await workApi.load());
}

function parseISODate(value) {
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

function isToday(dateISO) {
  return dateISO === toISODate(new Date());
}

function isOverdue(assignment) {
  return assignment.status !== "done" && parseISODate(assignment.dueDate) < parseISODate(toISODate(new Date()));
}

function priorityWeight(priority) {
  return { high: 3, normal: 2, low: 1 }[priority] || 2;
}

function getNotificationPermission() {
  if (!("Notification" in window)) return "unavailable";
  return Notification.permission;
}

function render() {
  renderProfile();
  renderMetrics();
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
  el.sessionPill.textContent = state.user?.label || "Logged in";
  el.employeeName.textContent = state.employee.name;
  el.employeeRole.textContent = state.employee.role;
  el.employeeAvailability.textContent = state.employee.availability;
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
  const open = state.assignments.filter((assignment) => assignment.status !== "done");
  const done = state.assignments.filter((assignment) => assignment.status === "done");
  const dueToday = open.filter((assignment) => isToday(assignment.dueDate));
  const overdue = open.filter(isOverdue);

  el.openCount.textContent = open.length;
  el.doneCount.textContent = done.length;
  el.dueTodayCount.textContent = dueToday.length;
  el.summaryLine.textContent = `${open.length} open - ${dueToday.length} due today - ${overdue.length} overdue`;
}

function renderFilters() {
  el.filterButtons.forEach((button) => {
    button.checked = button.dataset.filter === state.ui.filter;
  });
}

function getVisibleAssignments() {
  return [...state.assignments]
    .filter((assignment) => {
      if (state.ui.filter === "done") return assignment.status === "done";
      if (state.ui.filter === "open") return assignment.status !== "done";
      return true;
    })
    .sort((a, b) => {
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
    el.assignmentList.querySelector("span").textContent =
      state.ui.filter === "done" ? "No finished work yet." : "Ready for the next assignment.";
    return;
  }

  el.assignmentList.replaceChildren(...assignments.map(renderAssignmentCard));
}

function renderAssignmentCard(assignment) {
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
  card.querySelector(".assignment-meta strong").textContent = state.employee.name;
  card.querySelector(".assignment-meta em").textContent = `Due ${formatDate(parseISODate(assignment.dueDate), {
    weekday: "short",
    day: "numeric",
    month: "short",
  })}`;

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
  el.assignmentTitle.value = draft.title;
  el.assignmentDueDate.value = draft.dueDate;
  el.assignmentPriority.value = draft.priority;
  el.assignmentNotes.value = draft.notes;
  el.assignmentError.textContent = state.ui.assignmentError;
  el.deleteAssignmentButton.hidden = !draft.id;

  if (state.ui.assignmentDialogOpen && !wasOpen) {
    el.assignmentDialog.showModal();
    el.assignmentTitle.focus();
  }

  if (!state.ui.assignmentDialogOpen && wasOpen) {
    el.assignmentDialog.close();
  }
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
    title: draft.title.trim(),
    dueDate: draft.dueDate,
    priority: draft.priority,
    notes: draft.notes.trim(),
  };
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

async function syncBookings() {
  if (!canSyncBookings()) return;

  setUiState({
    syncInProgress: true,
    bookingSyncStatus: "Checking upcoming bookings...",
  });

  try {
    const data = await workApi.syncBookings();
    const message = data.syncable === 0
      ? "No upcoming bookings found to assign to Faye."
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
  showToast(`${state.employee.name} finished: ${assignment.title}`);

  if (!("Notification" in window) || state.ui.notificationPermission !== "granted") return;

  try {
    new Notification("Work finished", { body: `${state.employee.name} finished: ${assignment.title}` });
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
  el.syncBookingsButton.addEventListener("click", syncBookings);
  el.enableMessagesButton.addEventListener("click", enableNotifications);
  el.clearMessagesButton.addEventListener("click", clearMessages);

  el.filterButtons.forEach((button) => {
    button.addEventListener("change", () => {
      setUiState({ filter: button.dataset.filter });
    });
  });

  el.assignmentList.addEventListener("click", async (event) => {
    const completeTarget = event.target.closest("[data-complete-assignment]");
    if (completeTarget) {
      await completeAssignment(completeTarget.dataset.completeAssignment);
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
window.addEventListener("load", refreshIcons);
