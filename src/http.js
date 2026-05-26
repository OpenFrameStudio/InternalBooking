export function sendJson(res, status, payload, headers = {}) {
  const body = JSON.stringify(payload);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(body),
    ...headers
  });
  res.end(body);
}

export function sendNoContent(res) {
  res.writeHead(204);
  res.end();
}

export function redirect(res, location) {
  res.writeHead(302, {
    Location: location,
    "Cache-Control": "no-store"
  });
  res.end();
}

export function parseCookies(header = "") {
  const cookies = {};

  for (const part of header.split(";")) {
    const [rawKey, ...rawValue] = part.trim().split("=");
    if (!rawKey) {
      continue;
    }

    try {
      cookies[rawKey] = decodeURIComponent(rawValue.join("="));
    } catch {
      cookies[rawKey] = rawValue.join("=");
    }
  }

  return cookies;
}

export function readBody(req) {
  return new Promise((resolve, reject) => {
    let data = "";
    const configuredMaxBytes = Number(process.env.JSON_BODY_LIMIT_BYTES || 4_500_000);
    const maxBytes = Number.isFinite(configuredMaxBytes) && configuredMaxBytes > 0 ? configuredMaxBytes : 4_500_000;
    req.on("data", (chunk) => {
      data += chunk;
      if (data.length > maxBytes) {
        reject(new Error("Request body is too large."));
        req.destroy();
      }
    });
    req.on("end", () => {
      if (!data) {
        resolve({});
        return;
      }

      try {
        resolve(JSON.parse(data));
      } catch {
        reject(new Error("Request body must be valid JSON."));
      }
    });
    req.on("error", reject);
  });
}
