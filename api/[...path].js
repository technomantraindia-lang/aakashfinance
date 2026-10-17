import https from "node:https";

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
  if (value) return String(value);
  const pathname = String(req.url || "").split("?", 1)[0];
  return decodeURIComponent(pathname.replace(/^\/api\/?/, "").replace(/^\/+|\/+$/g, ""));
}

function getBody(req) {
  if (req.body == null || req.body === "") return undefined;
  if (typeof req.body === "string" || Buffer.isBuffer(req.body)) return req.body;
  return JSON.stringify(req.body);
}

function requestBackend(method, path, headers, body) {
  return new Promise((resolve, reject) => {
    const request = https.request({
      hostname: "erp.aakashfinance.com",
      port: 443,
      servername: "erp.aakashfinance.com",
      family: 4,
      method,
      path,
      headers
    }, (response) => {
      const chunks = [];
      response.on("data", (chunk) => chunks.push(chunk));
      response.on("end", () => resolve({
        status: response.statusCode || 502,
        headers: response.headers,
        body: Buffer.concat(chunks)
      }));
    });
    request.on("error", reject);
    if (body !== undefined) request.write(body);
    request.end();
  });
}

export default async function handler(req, res) {
  const origin = String(req.headers.origin || "");
  setCors(res, origin);
  if (req.method === "OPTIONS") return res.status(204).end();

  const path = getPath(req);
  if (!path) return res.status(404).json({ error: "API path is required." });

  const query = String(req.url || "").includes("?")
    ? String(req.url).slice(String(req.url).indexOf("?"))
    : "";
  const targetPath = `/api/${path}${query}`;
  const body = ["GET", "HEAD"].includes(req.method) ? undefined : getBody(req);
  const headers = {
    Accept: String(req.headers.accept || "application/json"),
    "User-Agent": "Mozilla/5.0 (compatible; KuberFinanceProxy/1.0)",
    ...(req.headers["content-type"] ? { "Content-Type": req.headers["content-type"] } : {}),
    ...(req.headers.authorization ? { Authorization: req.headers.authorization } : {}),
    ...(req.headers["cache-control"] ? { "Cache-Control": req.headers["cache-control"] } : {}),
    ...(body !== undefined ? { "Content-Length": Buffer.byteLength(body) } : {})
  };

  try {
    const upstream = await requestBackend(req.method, targetPath, headers, body);
    const contentType = upstream.headers["content-type"];
    if (contentType) res.setHeader("Content-Type", contentType);
    const cacheControl = upstream.headers["cache-control"];
    if (cacheControl) res.setHeader("Cache-Control", cacheControl);
    return res.status(upstream.status).send(upstream.body);
  } catch (error) {
    console.error("API proxy error:", error);
    return res.status(502).json({ error: "Backend proxy could not reach the API." });
  }
}
