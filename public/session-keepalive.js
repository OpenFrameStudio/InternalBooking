const sessionKeepAliveIntervalMs = 13 * 60 * 1000;

async function refreshSession() {
  try {
    const response = await fetch(`/api/session?_of=${Date.now().toString(36)}`, {
      credentials: "include",
      cache: "no-store"
    });
    const data = await response.json().catch(() => ({}));
    if (!response.ok || !data.authenticated) {
      window.location.assign("/login");
    }
  } catch {
    // Ignore temporary network wake-up failures; the next timer can retry.
  }
}

window.setInterval(refreshSession, sessionKeepAliveIntervalMs);
