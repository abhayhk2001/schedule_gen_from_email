// Loads the task pane with a stubbed Office host in headless Chrome, clicks
// "Extract events from this email", and fails on any console error or unhandled
// rejection. `node --check` only proves the file parses; this proves the module
// actually initialises and that every identifier it references exists.
import { spawn } from "node:child_process";
import { mkdtempSync, writeFileSync, cpSync, rmSync, existsSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { extname } from "node:path";

const CHROME = process.env.CHROME_PATH
  ?? "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

if (!existsSync(CHROME)) {
  console.log(`smoke-pane: skipped, no Chrome at ${CHROME} (set CHROME_PATH to override)`);
  process.exit(0);
}

const root = mkdtempSync(join(tmpdir(), "pane-smoke-"));
cpSync("public/outlook-addin", join(root, "outlook-addin"), { recursive: true });

// Stub Office before app.js runs, and replace the CDN office.js (which blanks
// the document when it cannot find a host).
const stub = `
window.__errors = [];
window.onerror = (m) => window.__errors.push("onerror: " + m);
window.addEventListener("unhandledrejection", (e) =>
  window.__errors.push("unhandledrejection: " + (e.reason && e.reason.message || e.reason)));
const origError = console.error;
console.error = (...a) => { window.__errors.push("console.error: " + a.join(" ")); origError(...a); };
window.Office = {
  HostType: { Outlook: "Outlook" },
  CoercionType: { Text: "text" },
  AsyncResultStatus: { Succeeded: "succeeded" },
  EventType: { ItemChanged: "olkItemSelectedChanged", ThemeChanged: "officeThemeChanged" },
  onReady: (cb) => setTimeout(() => cb({ host: "Outlook", platform: "Mac" }), 0),
  context: {
    officeTheme: { bodyBackgroundColor: "#1f1f1f", bodyForegroundColor: "#f0f0f0",
      controlBackgroundColor: "#292929", controlForegroundColor: "#e6e6e6", controlBorderColor: "#3d3d3d" },
    ui: { displayDialogAsync: () => {} },
    addHandlerAsync: (t, h, cb) => cb && cb({ status: "succeeded" }),
    mailbox: {
      addHandlerAsync: (t, h, cb) => cb && cb({ status: "succeeded" }),
      item: {
        subject: "Bargaining Session #9",
        sender: { emailAddress: "geo@example.org" },
        body: { getAsync: (type, cb) => cb({ status: "succeeded", value: "Session #10 on 2026-09-24 09:00-13:00" }) },
      },
    },
  },
};
`;
let html = await readFile("public/outlook-addin/index.html", "utf8");
html = html.replace(
  '<script src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>',
  `<script>${stub}</script>`,
);
writeFileSync(join(root, "outlook-addin", "index.html"), html);

// Stand in for /api/extract-event so the click path runs end to end.
const TYPES = { ".html": "text/html", ".js": "text/javascript", ".css": "text/css", ".png": "image/png" };
const server = createServer(async (req, res) => {
  if (req.url.startsWith("/api/extract-event")) {
    res.writeHead(200, { "Content-Type": "application/json" });
    res.end(JSON.stringify({ events: [{ event_name: "Bargaining Session #10", date: "2026-09-24",
      time: "09:00", end_time: "13:00", timezone: "America/Chicago", description: "Extracted." }] }));
    return;
  }
  try {
    const body = await readFile(join(root, decodeURIComponent(req.url.split("?")[0])));
    res.writeHead(200, { "Content-Type": TYPES[extname(req.url)] ?? "application/octet-stream" });
    res.end(body);
  } catch {
    res.writeHead(404).end("not found");
  }
});
await new Promise((r) => server.listen(0, r));
const port = server.address().port;

const out = join(root, "dom.txt");
const page = `http://localhost:${port}/outlook-addin/index.html`;
const driver = `
  (async () => {
    await new Promise(r => setTimeout(r, 300));
    document.getElementById("extract-btn").click();
    await new Promise(r => setTimeout(r, 900));
    const ev = document.querySelectorAll("#events .event").length;
    document.title = JSON.stringify({ errors: window.__errors, eventCards: ev,
      status: document.getElementById("status").textContent,
      createHidden: document.getElementById("create-btn").classList.contains("hidden") });
  })();
`;
writeFileSync(join(root, "outlook-addin", "drive.js"), driver);
writeFileSync(
  join(root, "outlook-addin", "index.html"),
  html.replace("</body>", '<script type="module" src="drive.js"></script></body>'),
);

const chrome = spawn(CHROME, ["--headless=new", "--disable-gpu", "--dump-dom",
  "--virtual-time-budget=4000", page], { stdio: ["ignore", "pipe", "ignore"] });
let dom = "";
chrome.stdout.on("data", (d) => (dom += d));
await new Promise((r) => chrome.on("exit", r));
server.close();

const m = dom.match(/<title>([\s\S]*?)<\/title>/);
if (!m) { console.error("FAIL: page never reported (module likely threw on load)"); rmSync(root, { recursive: true, force: true }); process.exit(1); }
const result = JSON.parse(m[1].replace(/&quot;/g, '"').replace(/&amp;/g, "&"));
rmSync(root, { recursive: true, force: true });

console.log(JSON.stringify(result, null, 2));
const bad = result.errors.length > 0 || result.eventCards !== 1 || result.createHidden;
console.log(bad ? "\nFAIL" : "\nPASS: extract click rendered 1 event card, no console errors");
process.exit(bad ? 1 : 0);
