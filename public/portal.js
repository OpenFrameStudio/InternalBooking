const appCatalog = [
  {
    id: "bookings",
    name: "Bookings",
    href: "/bookings",
    icon: "calendar-days",
    tone: "green",
  },
  {
    id: "work",
    name: "Work",
    href: "/work/",
    icon: "clipboard-check",
    tone: "blue",
  },
  {
    id: "invoices",
    name: "Invoices",
    href: "/invoices",
    icon: "receipt-text",
    tone: "violet",
  },
  {
    id: "wages",
    name: "Wages",
    href: "/wages",
    icon: "wallet",
    tone: "amber",
  },
  {
    id: "clients",
    name: "Clients",
    href: "/clients",
    icon: "users-round",
    tone: "amber",
  },
  {
    id: "photographers",
    name: "Photographers",
    href: "/photographers",
    icon: "camera",
    tone: "pink",
  },
];

const workNoticeDismissedKey = "openframe.workNoticeDismissed.v1";
const launcherGrid = document.querySelector("#launcherGrid");
const sessionLabel = document.querySelector("#sessionLabel");
const logoutButton = document.querySelector("#logoutButton");

const state = {
  user: null,
  workAssignments: [],
};

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
    cache: "no-store",
    headers: { Accept: "application/json", ...(options.headers || {}) },
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

function renderApps(user) {
  const allowedApps = new Set(user?.apps || []);
  const apps = appCatalog.filter((app) => allowedApps.has(app.id));
  const workNotice = workNoticeSummary();

  sessionLabel.textContent = user?.label || "Logged in";
  launcherGrid.replaceChildren(
    ...apps.map((app) => {
      const link = document.createElement("a");
      link.className = "launcher-app";
      link.href = app.href;
      if (app.id === "work" && workNotice.count) {
        link.setAttribute("aria-label", `Work, ${workNotice.count} open assignment${workNotice.count === 1 ? "" : "s"}`);
        link.addEventListener("click", () => dismissWorkNotice(workNotice.signature));
      }
      link.innerHTML = `
        <span class="launcher-icon ${app.tone}">
          <i data-lucide="${app.icon}"></i>
          ${app.id === "work" && workNotice.count ? `<span class="launcher-badge">${workNotice.label}</span>` : ""}
        </span>
        <span>${app.name}</span>
      `;
      return link;
    }),
  );

  if (window.lucide) {
    window.lucide.createIcons();
  }
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
    // This only controls the app launcher badge.
  }
}

function dismissWorkNotice(signature) {
  writeWorkNoticeDismissedSignature(signature);
}

function workNoticeSignature(assignments) {
  return assignments
    .map((assignment) => `${assignment.id}:${assignment.updatedAt || assignment.createdAt || ""}`)
    .sort()
    .join("|");
}

function workNoticeSummary() {
  const openAssignments = state.workAssignments.filter((assignment) => assignment.status !== "done");
  const signature = workNoticeSignature(openAssignments);
  const dismissedSignature = readWorkNoticeDismissedSignature();
  const count = signature && signature !== dismissedSignature ? openAssignments.length : 0;

  return {
    count,
    signature,
    label: count > 9 ? "9+" : String(count),
  };
}

async function loadWorkNotice() {
  if (!state.user?.apps?.includes("work")) return;

  try {
    const data = await apiFetch("/api/work");
    state.workAssignments = Array.isArray(data.assignments) ? data.assignments : [];
    renderApps(state.user);
  } catch {
    state.workAssignments = [];
  }
}

async function logout() {
  logoutButton.disabled = true;
  try {
    await fetch("/api/logout", { method: "POST", credentials: "include" });
  } finally {
    window.location.href = "/login";
  }
}

logoutButton.addEventListener("click", logout);

const session = await apiFetch("/api/session");
state.user = session.user || null;
renderApps(state.user);
loadWorkNotice();

window.addEventListener("load", () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
});
