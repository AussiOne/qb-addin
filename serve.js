#!/usr/bin/env node
// Minimal HTTPS static server for local testing: serves ./src at https://localhost:3000
// Uses the trusted dev certificate created by:  npm run certs
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = Number(process.env.PORT) || 3000;
const ROOT = path.join(__dirname, "src");
const certDir = path.join(os.homedir(), ".office-addin-dev-certs");
const keyFile = path.join(certDir, "localhost.key");
const crtFile = path.join(certDir, "localhost.crt");

if (!fs.existsSync(keyFile) || !fs.existsSync(crtFile)) {
  console.error("Dev certificate not found. Run:  npm run certs");
  process.exit(1);
}

const TYPES = { ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".css": "text/css",
  ".png": "image/png", ".svg": "image/svg+xml", ".json": "application/json", ".ico": "image/x-icon" };

https.createServer({ key: fs.readFileSync(keyFile), cert: fs.readFileSync(crtFile) }, (req, res) => {
  const urlPath = decodeURIComponent(req.url.split("?")[0]);
  let file = path.normalize(path.join(ROOT, urlPath === "/" ? "index.html" : urlPath));
  if (!file.startsWith(ROOT)) { res.writeHead(403); return res.end(); }
  fs.readFile(file, (err, data) => {
    if (err) { res.writeHead(404); return res.end("Not found"); }
    res.writeHead(200, { "Content-Type": TYPES[path.extname(file)] || "application/octet-stream",
                         "Cache-Control": "no-cache" });
    res.end(data);
  });
}).listen(PORT, () => console.log(`Serving ${ROOT} at https://localhost:${PORT}/index.html`));
