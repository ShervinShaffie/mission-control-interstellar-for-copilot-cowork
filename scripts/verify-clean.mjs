#!/usr/bin/env node
// Privacy check for Mission Control Interstellar. Run before every release: `node scripts/verify-clean.mjs`
// Fails (exit 1) if the repository contains anything that looks like personal or organization-specific data:
// email addresses, GUIDs (tenant, environment, app or user IDs), connection IDs, or tenant-specific URLs.
// Plain Node.js, no packages to install.
import { readdirSync, readFileSync, statSync } from "node:fs";
import { join, relative, extname } from "node:path";

const root = new URL("..", import.meta.url).pathname;
const SKIP_DIRS = new Set([".git", "node_modules", "dist", ".vite-cache"]);
const BINARY = new Set([".jpg", ".jpeg", ".png", ".gif", ".webp", ".ico", ".zip", ".pdf", ".woff", ".woff2"]);

// Placeholders and documentation examples that are safe to publish.
const ALLOWED_EMAIL_DOMAINS = /@(contoso\.(com|onmicrosoft\.com)|example\.(com|org|net)|x\.com|domain\.com|thread\.(tacv|v2|skype))$/i;
const ALLOWED_GUIDS = new Set(["00000000-0000-0000-0000-000000000000", "11111111-2222-3333-4444-555555555555", "00000000-1111-2222-3333-444444444444"]);

const checks = [
  { name: "email address", re: /[A-Za-z0-9._%+-]+@[A-Za-z0-9-]+(?:\.[A-Za-z0-9-]+)*\.[A-Za-z]{2,}/g, allow: m => ALLOWED_EMAIL_DOMAINS.test(m) },
  { name: "GUID (tenant, environment, app or user ID)", re: /\b[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}\b/gi, allow: m => ALLOWED_GUIDS.has(m.toLowerCase()) },
  { name: "connection ID", re: /\b[0-9a-f]{32}\b/gi, allow: () => false },
  { name: "tenant-specific URL", re: /https?:\/\/[a-z0-9-]+(-my)?\.sharepoint\.com\/[^\s"')]*|https?:\/\/[a-z0-9-]+\.crm[0-9]*\.dynamics\.com[^\s"')]*|https?:\/\/(make|apps)\.powerapps\.com\/environments\/[^\s"')]+|https?:\/\/play\.managedapps\.cloud\.microsoft\/apps\/[^\s"')]+/gi, allow: m => /^https?:\/\/(contoso(-my)?|[a-z])\.sharepoint\.com\//i.test(m) || m.includes("${") || [...ALLOWED_GUIDS].some(g => m.includes(g)) },
  { name: "tenant domain", re: /\b[a-z0-9-]+\.onmicrosoft\.com\b/gi, allow: m => /^contoso\./i.test(m) },
];

function* files(dir) {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) { if (!SKIP_DIRS.has(name)) yield* files(path); }
    else if (!BINARY.has(extname(name).toLowerCase())) yield path;
  }
}

const hits = [];
for (const path of files(root)) {
  const rel = relative(root, path);
  if (rel === "scripts/verify-clean.mjs") continue;
  readFileSync(path, "utf8").split(/\r?\n/).forEach((line, i) => {
    for (const check of checks) {
      for (const match of line.match(check.re) ?? []) {
        if (/^@odata/i.test(match.slice(match.indexOf("@")))) continue; // OData annotations, not emails
        if (!check.allow(match)) hits.push(`${rel}:${i + 1}  ${check.name}: ${match}`);
      }
    }
  });
}

if (hits.length) {
  console.error(`verify-clean: FAILED, ${hits.length} item(s) look like personal or organization data:\n` + hits.map(h => `  ${h}`).join("\n"));
  process.exit(1);
}
console.log("verify-clean: passed. No personal or organization-specific data found.");
