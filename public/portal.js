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

const launcherGrid = document.querySelector("#launcherGrid");
const sessionLabel = document.querySelector("#sessionLabel");
const logoutButton = document.querySelector("#logoutButton");

async function apiFetch(url, options = {}) {
  const response = await fetch(url, {
    credentials: "include",
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

  sessionLabel.textContent = user?.label || "Logged in";
  launcherGrid.replaceChildren(
    ...apps.map((app) => {
      const link = document.createElement("a");
      link.className = "launcher-app";
      link.href = app.href;
      link.innerHTML = `
        <span class="launcher-icon ${app.tone}">
          <i data-lucide="${app.icon}"></i>
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
renderApps(session.user);

window.addEventListener("load", () => {
  if (window.lucide) {
    window.lucide.createIcons();
  }
});
