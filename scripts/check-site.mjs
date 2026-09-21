// Serves public/ on a random port and asserts the landing page loads with
// expected content. This is the literal "does the website work" check used by
// the pre-commit guardrail and by `npm run verify` — it has no Chrome
// dependency and always runs.
import { createServer } from "node:http";
import { readFile, stat } from "node:fs/promises";
import { extname, join, resolve } from "node:path";

const root = resolve("public");

const TYPES = {
  ".html": "text/html",
  ".js": "text/javascript",
  ".css": "text/css",
  ".png": "image/png",
  ".svg": "image/svg+xml",
};

async function tryRead(relPath) {
  const safe = relPath.replace(/\.\.+/g, "").replace(/^\/+/, "");
  const full = safe === "" ? "index.html" : safe;
  const target = join(root, full);
  const resolved = resolve(target);
  if (!resolved.startsWith(root + "/") && resolved !== join(root, "index.html")) {
    return null;
  }
  try {
    await stat(resolved);
    return await readFile(resolved);
  } catch {
    return null;
  }
}

const server = createServer(async (req, res) => {
  const url = decodeURIComponent((req.url ?? "/").split("?")[0]);
  const body = await tryRead(url);
  if (body == null) {
    res.writeHead(404).end("not found");
    return;
  }
  res.writeHead(200, { "Content-Type": TYPES[extname(url)] ?? "application/octet-stream" });
  res.end(body);
});

await new Promise((r) => server.listen(0, r));
const port = server.address().port;

let failed = false;
try {
  const resp = await fetch(`http://localhost:${port}/`);
  if (resp.status !== 200) {
    console.error(`FAIL: / returned status ${resp.status}`);
    failed = true;
  } else {
    const text = await resp.text();
    if (!text.includes("Event Extractor")) {
      console.error("FAIL: / response is missing 'Event Extractor'");
      failed = true;
    } else {
      console.log("PASS: landing page returns 200 and contains 'Event Extractor'");
    }
  }
} catch (err) {
  console.error(`FAIL: could not fetch / — ${err?.message ?? String(err)}`);
  failed = true;
} finally {
  server.close();
}

if (failed) process.exit(1);
