#!/usr/bin/env node
// Generates a ready-to-install manifest from manifest.xml.
//
//   node configure.js <baseUrl> <realm>
//
// Examples:
//   node configure.js https://yourname.github.io/qb-addin acme     -> manifest.prod.xml
//   node configure.js https://localhost:3000 acme                  -> manifest.dev.xml
//
// <baseUrl> is the HTTPS URL that serves the contents of the /src folder
// (index.html must be reachable at <baseUrl>/index.html).

const fs = require("fs");
const path = require("path");

const [, , baseArg, realmArg] = process.argv;
if (!baseArg || !realmArg) {
  console.error("Usage: node configure.js <baseUrl> <realm>");
  process.exit(1);
}
const base = baseArg.replace(/\/+$/, "");
const realm = realmArg.replace(/\.quickbase\.com.*$/i, "").replace(/^https?:\/\//i, "");
if (!/^https:\/\//i.test(base)) {
  console.error("baseUrl must start with https:// (Office add-ins require HTTPS).");
  process.exit(1);
}

const isDev = /localhost|127\.0\.0\.1/i.test(base);
let xml = fs.readFileSync(path.join(__dirname, "manifest.xml"), "utf8");
xml = xml.split("https://YOUR-HOST/qb-addin").join(base).split("YOURREALM").join(realm);

if (isDev) {
  // Separate identity so the dev and prod add-ins can be installed side by side.
  xml = xml
    .replace("<Id>81a85707-81ea-473e-b267-276868a54bcf</Id>", "<Id>81a85707-81ea-473e-b267-276868a5d0e0</Id>")
    .replace('DisplayName DefaultValue="Quickbase Live Dashboard"', 'DisplayName DefaultValue="Quickbase Live Dashboard (Dev)"');
}
// Drop the instructions comment from the generated file.
xml = xml.replace(/<!--[\s\S]*?-->\s*/, "");

const out = path.join(__dirname, isDev ? "manifest.dev.xml" : "manifest.prod.xml");
fs.writeFileSync(out, xml);
console.log("Wrote " + path.basename(out));
console.log("  Add-in URL : " + base + "/index.html");
console.log("  Realm      : https://" + realm + ".quickbase.com");
