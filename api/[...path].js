const BACKEND_BASE = "https://erp.aakashfinance.com/api";

function setCors(res, origin) {
  const allowed = origin === "https://aakashfinance.com"
    || origin === "https://www.aakashfinance.com"
    || origin === "https://app-webappication.vercel.app"
    || /^https?:\/\/localhost(?::\d+)?$/.test(origin || "");
  if (allowed) res.setHeader("Access-Control-Allow-Origin", origin);
  res.setHeader("Vary", "Origin");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type, Authorization, Cache-Control");
  res.setHeader("Access-Control-Allow-Methods", "GET, HEAD, POST, PUT, PATCH, DELETE, OPTIONS");
}

function getPath(req) {
  const value = req.query?.path;
  if (Array.isArray(value)) return value.join("/");
  return String(value || "");
}

function getBody(req) {
  if (req.body == null || req.body === "") return undefined;
  if (typeof req.body === "string" || Buffer.isBuffer(req.body)) return req.body;
  return JSON.stringify(req.body);
}

export default async function handler(req, res) {
  const origin = String(req.headers.origin || "");
  setCors(res, origin);
  if (req.method === "OPTIONS") return res.status(204).end();

  const path = getPath(req);
  if (!path) return res.status(404).json({ error: "API path is required." });

  const target = `${BACKEND_BASE}/${path}${req.url.includes("?") ? req.url.slice(req.url.indexOf("?")) : ""}`;
  const headers = {
    Accept: String(req.headers.accept || "application/json"),
    ...(req.headers["content-type"] ? { "Content-Type": req.headers["content-type"] } : {}),
    ...(req.headers.authorization ? { Authorization: req.headers.authorization } : {}),
    ...(req.headers["cache-control"] ? { "Cache-Control": req.headers["cache-control"] } : {})
  };

  try {
    const upstream = await fetch(target, {
      method: req.method,
      headers,
      body: ["GET", "HEAD"].includes(req.method) ? undefined : getBody(req)
    });
    const contentType = upstream.headers.get("content-type");
    if (contentType) res.setHeader("Content-Type", contentType);
    const cacheControl = upstream.headers.get("cache-control");
    if (cacheControl) res.setHeader("Cache-Control", cacheControl);
    const payload = Buffer.from(await upstream.arrayBuffer());
    return res.status(upstream.status).send(payload);
  } catch (error) {
    console.error("API proxy error:", error);
    return res.status(502).json({ error: "Backend proxy could not reach the API." });
  }
}
