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

const state = {
  employee: {
    id: "faye",
    name: "Faye",
    role: "Editor / Admin",
    availability: "Mon-Fri, 12pm-8pm Australian time",
  },
  assignments: [],
  messages: [],
  filter: "open",
  user: null,
};

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

function applyWorkData(data) {
  state.employee = data.employee || state.employee;
  state.assignments = Array.isArray(data.assignments) ? data.assignments : [];
  state.messages = Array.isArray(data.messages) ? data.messages : [];
  state.user = data.user || state.user;
}

async function loadWork() {
  applyWorkData(await workApi.load());
  render();
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

function render() {
  renderProfile();
  renderMetrics();
  renderRoleAccess();
  renderFilters();
  renderMessages();
  renderAssignments();
  updateNotificationButton();
  refreshIcons();
}

function renderProfile() {
  el.sessionPill.textContent = state.user?.label || "Logged in";
  el.employeeName.textContent = state.employee.name;
  el.employeeRole.textContent = state.employee.role;
  el.employeeAvailability.textContent = state.employee.availability;
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
  el.summaryLine.textContent = `${open.length} open · ${dueToday.length} due today · ${overdue.length} overdue`;
}

function renderFilters() {
  el.filterButtons.forEach((button) => {
    button.checked = button.dataset.filter === state.filter;
  });
}

function getVisibleAssignments() {
  return [...state.assignments]
    .filter((assignment) => {
      if (state.filter === "done") return assignment.status === "done";
      if (state.filter === "open") return assignment.status !== "done";
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
      state.filter === "done" ? "No finished work yet." : "Ready for the next assignment.";
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

function openAssignmentDialog(assignment = null) {
  if (!canManageWork()) return;

  el.assignmentError.textContent = "";
  el.assignmentId.value = assignment?.id || "";
  el.assignmentTitle.value = assignment?.title || "";
  el.assignmentDueDate.value = assignment?.dueDate || toISODate(new Date());
  el.assignmentPriority.value = assignment?.priority || "normal";
  el.assignmentNotes.value = assignment?.notes || "";
  el.assignmentDialogTitle.textContent = assignment ? "Edit Work" : "Assign Work";
  el.deleteAssignmentButton.hidden = !assignment;
  el.assignmentDialog.showModal();
  el.assignmentTitle.focus();
}

function closeAssignmentDialog() {
  el.assignmentDialog.close();
}

async function handleAssignmentSubmit(event) {
  event.preventDefault();
  if (!canManageWork()) return;

  const id = el.assignmentId.value;
  const payload = {
    title: el.assignmentTitle.value.trim(),
    dueDate: el.assignmentDueDate.value,
    priority: el.assignmentPriority.value,
    notes: el.assignmentNotes.value.trim(),
  };

  if (!payload.title || !payload.dueDate) {
    el.assignmentError.textContent = "Work and due date are required.";
    return;
  }

  try {
    const data = id
      ? await workApi.updateAssignment(id, payload)
      : await workApi.createAssignment(payload);
    applyWorkData(data);
    closeAssignmentDialog();
    if (!id && data.workInviteMessage) showToast(data.workInviteMessage);
    render();
  } catch (error) {
    el.assignmentError.textContent = error.message;
  }
}

async function completeAssignment(id) {
  const assignmentBeforeUpdate = state.assignments.find((item) => item.id === id);
  if (!canCompleteAssignment(assignmentBeforeUpdate)) return;

  const data = await workApi.completeAssignment(id);
  applyWorkData(data);
  render();
  const assignment = state.assignments.find((item) => item.id === id);
  if (assignment) notifyCompletion(assignment);
}

async function reopenAssignment(id) {
  if (!canManageWork()) return;
  const data = await workApi.reopenAssignment(id);
  applyWorkData(data);
  render();
}

async function deleteCurrentAssignment() {
  if (!canManageWork()) return;

  const id = el.assignmentId.value;
  if (!id) return;
  const data = await workApi.deleteAssignment(id);
  applyWorkData(data);
  closeAssignmentDialog();
  render();
}

async function clearMessages() {
  if (!canViewWorkMessages()) return;
  const data = await workApi.clearMessages();
  applyWorkData(data);
  render();
}

async function syncBookings() {
  if (!canSyncBookings()) return;

  el.syncBookingsButton.disabled = true;
  el.bookingSyncStatus.textContent = "Checking upcoming bookings...";

  try {
    const data = await workApi.syncBookings();
    applyWorkData(data);
    state.filter = "open";
    const message = data.syncable === 0
      ? "No upcoming bookings found to assign to Faye."
      : `Synced ${data.syncable} upcoming bookings: ${data.created} new, ${data.updated} updated.`;
    const fullMessage = data.workInviteMessage ? `${message} ${data.workInviteMessage}` : message;
    el.bookingSyncStatus.textContent = fullMessage;
    showToast(fullMessage);
    render();
  } catch (error) {
    el.bookingSyncStatus.textContent = error.message;
    showToast(error.message);
  } finally {
    el.syncBookingsButton.disabled = false;
  }
}

function notifyCompletion(assignment) {
  const text = `${state.employee.name} finished: ${assignment.title}`;
  showToast(text);

  if (!("Notification" in window) || Notification.permission !== "granted") return;

  try {
    new Notification("Work finished", { body: text });
  } catch {
    showToast("Browser message could not be shown.");
  }
}

function showToast(text) {
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = text;
  el.toastRegion.append(toast);
  window.setTimeout(() => toast.remove(), 5200);
}

async function enableNotifications() {
  if (!canViewWorkMessages()) return;

  if (!("Notification" in window)) {
    showToast("Browser messages are not available here.");
    return;
  }

  const permission = await Notification.requestPermission();
  updateNotificationButton();
  showToast(permission === "granted" ? "Messages enabled." : "Messages were not enabled.");
}

function updateNotificationButton() {
  if (!("Notification" in window)) {
    el.enableMessagesButton.disabled = true;
    el.enableMessagesButton.title = "Browser messages unavailable";
    return;
  }

  const enabled = Notification.permission === "granted";
  el.enableMessagesButton.classList.toggle("active", enabled);
  el.enableMessagesButton.title = enabled ? "Messages enabled" : "Enable messages";
  el.enableMessagesButton.setAttribute("aria-label", enabled ? "Messages enabled" : "Enable messages");
}

async function logout() {
  el.logoutButton.disabled = true;
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
      state.filter = button.dataset.filter;
      render();
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

  el.assignmentForm.addEventListener("submit", handleAssignmentSubmit);
  el.deleteAssignmentButton.addEventListener("click", deleteCurrentAssignment);
  el.cancelAssignmentButton.addEventListener("click", closeAssignmentDialog);
  el.closeAssignmentDialogButton.addEventListener("click", closeAssignmentDialog);
}

wireEvents();
loadWork().catch((error) => showToast(error.message));
window.addEventListener("load", refreshIcons);
