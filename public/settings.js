const state = {
  user: null,
  status: null,
  savingPassword: false,
  savingRole: false
};

const el = {
  accountName: document.querySelector("#accountName"),
  accountLogin: document.querySelector("#accountLogin"),
  accountRole: document.querySelector("#accountRole"),
  accountApps: document.querySelector("#accountApps"),
  roleForm: document.querySelector("#roleForm"),
  roleSelect: document.querySelector("#roleSelect"),
  roleMessage: document.querySelector("#roleMessage"),
  saveRoleButton: document.querySelector("#saveRoleButton"),
  statusList: document.querySelector("#statusList"),
  passwordForm: document.querySelector("#passwordForm"),
  currentPassword: document.querySelector("#currentPasswordInput"),
  newPassword: document.querySelector("#newPasswordInput"),
  confirmPassword: document.querySelector("#confirmPasswordInput"),
  passwordMessage: document.querySelector("#passwordMessage"),
  savePasswordButton: document.querySelector("#savePasswordButton"),
  logoutButton: document.querySelector("#logoutButton")
};

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
    headers: {
      Accept: "application/json",
      ...(options.body ? { "Content-Type": "application/json" } : {}),
      ...(options.headers || {})
    },
    ...options
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

function appLabel(app) {
  return {
    bookings: "Bookings",
    clients: "Clients",
    photographers: "Photographers",
    work: "Work",
    invoices: "Invoices"
  }[app] || app;
}

function renderAccount() {
  const user = state.user || {};
  el.accountName.textContent = user.name || user.username || "Account";
  el.accountLogin.textContent = user.username || "-";
  el.accountRole.textContent = user.label || user.role || "-";
  el.accountApps.textContent = Array.isArray(user.apps) && user.apps.length
    ? user.apps.map(appLabel).join(", ")
    : "-";
  renderRoleForm();
}

function renderRoleForm() {
  const user = state.user || {};
  const roleOptions = Array.isArray(user.roleOptions) ? user.roleOptions : [];
  el.roleForm.hidden = !user.canChangeRole || roleOptions.length === 0;
  if (el.roleForm.hidden) return;

  const currentOptions = [...el.roleSelect.options].map((option) => option.value).join("|");
  const nextOptions = roleOptions.map((option) => option.value).join("|");
  if (currentOptions !== nextOptions) {
    el.roleSelect.replaceChildren(...roleOptions.map((option) => {
      const item = document.createElement("option");
      item.value = option.value;
      item.textContent = option.label;
      return item;
    }));
  }

  el.roleSelect.value = user.roleMode || user.role || "";
  updateRoleButton();
}

function updateRoleButton() {
  const user = state.user || {};
  el.saveRoleButton.disabled = state.savingRole || el.roleSelect.value === (user.roleMode || user.role || "");
  el.saveRoleButton.textContent = state.savingRole ? "Saving" : "Save role";
}

function statusRows() {
  const status = state.status || {};
  return [
    ["Lark calendar", status.larkConfigured],
    ["Booking invite email", status.calendarInviteEmailConfigured],
    ["Invoice email", status.invoiceEmailConfigured],
    ["Work Lark notifications", status.workLarkNotificationConfigured],
    ["Supabase storage", status.supabaseConfigured],
    ["Calendar notifications", status.larkNotificationsEnabled]
  ];
}

function renderStatus() {
  el.statusList.replaceChildren(
    ...statusRows().map(([label, enabled]) => {
      const row = document.createElement("div");
      row.className = "settings-status-row";
      row.innerHTML = `
        <span>${label}</span>
        <strong class="settings-pill ${enabled ? "good" : "quiet"}">${enabled ? "On" : "Off"}</strong>
      `;
      return row;
    })
  );
}

function setPasswordMessage(message, type = "") {
  el.passwordMessage.textContent = message;
  el.passwordMessage.className = `settings-message${type ? ` ${type}` : ""}`;
}

function setRoleMessage(message, type = "") {
  el.roleMessage.textContent = message;
  el.roleMessage.className = `settings-message${type ? ` ${type}` : ""}`;
}

async function changeRole(event) {
  event.preventDefault();
  if (state.savingRole) return;

  state.savingRole = true;
  updateRoleButton();
  setRoleMessage("Saving role...");

  try {
    const data = await apiFetch("/api/change-role", {
      method: "POST",
      body: JSON.stringify({ role: el.roleSelect.value })
    });
    state.user = data.user || state.user;
    renderAccount();
    setRoleMessage("Role updated.", "success");
  } catch (error) {
    setRoleMessage(error.message || "Role could not be changed.", "error");
  } finally {
    state.savingRole = false;
    updateRoleButton();
  }
}

async function changePassword(event) {
  event.preventDefault();
  if (state.savingPassword) return;

  const currentPassword = el.currentPassword.value;
  const newPassword = el.newPassword.value;
  const confirmPassword = el.confirmPassword.value;

  if (newPassword.length < 4) {
    setPasswordMessage("Use at least 4 characters.", "error");
    return;
  }

  if (newPassword !== confirmPassword) {
    setPasswordMessage("The new passwords do not match.", "error");
    return;
  }

  state.savingPassword = true;
  el.savePasswordButton.disabled = true;
  el.savePasswordButton.textContent = "Saving";
  setPasswordMessage("Saving password...");

  try {
    await apiFetch("/api/change-password", {
      method: "POST",
      body: JSON.stringify({ currentPassword, newPassword, confirmPassword })
    });
    el.passwordForm.reset();
    setPasswordMessage("Password updated.", "success");
  } catch (error) {
    setPasswordMessage(error.message || "Password could not be changed.", "error");
  } finally {
    state.savingPassword = false;
    el.savePasswordButton.disabled = false;
    el.savePasswordButton.textContent = "Save password";
  }
}

async function logout() {
  el.logoutButton.disabled = true;
  try {
    await fetch("/api/logout", { method: "POST", credentials: "include" });
  } finally {
    window.location.href = "/login";
  }
}

async function loadSettings() {
  const [session, status] = await Promise.all([
    apiFetch("/api/session"),
    apiFetch("/api/status")
  ]);
  state.user = session.user;
  state.status = status;
  renderAccount();
  renderStatus();
  if (window.lucide) {
    window.lucide.createIcons();
  }
}

el.passwordForm.addEventListener("submit", changePassword);
el.roleForm.addEventListener("submit", changeRole);
el.roleSelect.addEventListener("change", () => {
  setRoleMessage("");
  updateRoleButton();
});
el.logoutButton.addEventListener("click", logout);

await loadSettings();

window.addEventListener("load", () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
});
