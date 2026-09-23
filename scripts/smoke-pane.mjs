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
window.__graphPosts = [];
window.__graphViews = 0;
window.__graphMessageGets = 0;
const SOURCE_WEB_LINK = "https://outlook.office365.com/owa/?ItemID=SMOKE%3D&exvsurl=1";
// The "&" must arrive HTML-escaped inside the href attribute.
const EXPECTED_HREF = 'href="https://outlook.office365.com/owa/?ItemID=SMOKE%3D&amp;exvsurl=1"';
// Stands in for the user's calendar: one event that exactly matches the
// "Upcoming Session" fixture (with a punctuation difference, so the subject
// normaliser is exercised) and nothing matching the "Past Session" one.
const CANNED_CALENDAR = [{
  id: "existing-1",
  subject: "Upcoming Session!",
  start: { dateTime: "2099-09-24T09:00:00.0000000", timeZone: "America/Chicago" },
  end: { dateTime: "2099-09-24T13:00:00.0000000", timeZone: "America/Chicago" },
  isAllDay: false,
  webLink: "https://outlook.example/existing-1",
}];
const origFetch = window.fetch.bind(window);
window.fetch = async (url, init) => {
  const href = typeof url === "string" ? url : url.url;
  if (href.startsWith("https://graph.microsoft.com/")) {
    if (href.includes("/me/messages/")) {
      window.__graphMessageGets += 1;
      return {
        ok: true,
        status: 200,
        headers: { get: () => null },
        json: async () => ({ webLink: SOURCE_WEB_LINK }),
      };
    }
    if (href.includes("/calendarView")) {
      window.__graphViews += 1;
      return {
        ok: true,
        status: 200,
        // The pane refuses to trust a window Graph did not return in the
        // requested timezone, so the stub must echo the applied preference.
        headers: { get: (h) => (h === "Preference-Applied" ? 'outlook.timezone="America/Chicago"' : null) },
        json: async () => ({ value: CANNED_CALENDAR }),
      };
    }
    window.__graphPosts.push(JSON.parse(init.body));
    return { ok: true, status: 201, json: async () => ({ id: "smoke-1", webLink: "#" }) };
  }
  return origFetch(url, init);
};
window.Office = {
  HostType: { Outlook: "Outlook" },
  MailboxEnums: { RestVersion: { v2_0: "v2.0" } },
  auth: { getAccessToken: (opts, cb) => cb({ status: "succeeded", value: "smoke-token" }) },
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
      convertToRestId: (id) => "rest-" + id,
      item: {
        itemId: "EWS-ITEM-1",
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
    res.end(JSON.stringify({ events: [
      { event_name: "Upcoming Session", date: "2099-09-24", time: "09:00",
        end_time: "13:00", timezone: "America/Chicago", location: "Room 100", description: "Upcoming." },
      { event_name: "Past Session", date: "2000-01-01", time: "09:00",
        end_time: "13:00", timezone: "America/Chicago", location: "Room 200", description: "Past." },
    ] }));
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
    const upcoming = document.querySelectorAll("#events-upcoming .event").length;
    const past = document.querySelectorAll("#events-past .event").length;
    // Snapshot the post-extract state before creating: a successful create
    // swaps the card list for results and hides the create button.
    const snapshot = {
      upcomingCards: upcoming, pastCards: past,
      upcomingSectionHidden: document.getElementById("upcoming-section").classList.contains("hidden"),
      pastSectionHidden: document.getElementById("past-section").classList.contains("hidden"),
      upcomingCreateHidden: document.getElementById("create-upcoming-btn").classList.contains("hidden"),
      pastCreateHidden: document.getElementById("create-past-btn").classList.contains("hidden"),
      status: document.getElementById("status").textContent,
    };
    // 1. The upcoming fixture is already on the canned calendar: expect no POST
    //    and a duplicate row offering the override.
    document.getElementById("create-upcoming-btn").click();
    await new Promise(r => setTimeout(r, 900));
    const afterDuplicate = {
      postsAfterDuplicate: window.__graphPosts.length,
      duplicateRows: document.querySelectorAll("#events-upcoming .event.result.duplicate").length,
      duplicateLink: !!document.querySelector("#events-upcoming .event.result.duplicate a"),
    };

    // 2. The past fixture matches nothing: the check must not block it.
    document.getElementById("create-past-btn").click();
    await new Promise(r => setTimeout(r, 900));
    const afterPast = {
      postsAfterPast: window.__graphPosts.length,
      pastSubject: window.__graphPosts.length ? window.__graphPosts[0].subject : null,
    };

    // 3. "Create anyway" overrides the duplicate and posts it after all.
    const anyway = document.querySelector("#events-upcoming .event.result.duplicate .retry-btn");
    if (anyway) anyway.click();
    await new Promise(r => setTimeout(r, 900));

    document.title = JSON.stringify({ errors: window.__errors, ...snapshot,
      ...afterDuplicate, ...afterPast,
      graphViews: window.__graphViews,
      graphMessageGets: window.__graphMessageGets,
      // The duplicate row never posts, so the first POST is the past event and
      // the second is the "Create anyway" override. Both must carry the link.
      bodyTypes: window.__graphPosts.map(p => p.body && p.body.contentType),
      bodiesWithLink: window.__graphPosts
        .filter(p => p.body && String(p.body.content).includes(EXPECTED_HREF))
        .length,
      bodyEscapesSubject: window.__graphPosts.every(
        p => !String(p.body && p.body.content).includes("<script")),
      postsAfterOverride: window.__graphPosts.length,
      graphPosts: window.__graphPosts.length,
      graphLocations: window.__graphPosts.map(p => p.location ? p.location.displayName : null) });
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
const bad =
  result.errors.length > 0 ||
  result.upcomingCards !== 1 ||
  result.pastCards !== 1 ||
  result.postsAfterDuplicate !== 0 ||
  result.duplicateRows !== 1 ||
  !result.duplicateLink ||
  result.postsAfterPast !== 1 ||
  result.pastSubject !== "Past Session" ||
  result.postsAfterOverride !== 2 ||
  result.graphViews < 2 ||
  result.graphLocations.length !== 2 ||
  result.graphLocations[0] !== "Room 200" ||
  result.graphLocations[1] !== "Room 100" ||
  result.bodyTypes.join(",") !== "HTML,HTML" ||
  result.bodiesWithLink !== 2 ||
  !result.bodyEscapesSubject ||
  result.graphMessageGets !== 1 ||
  result.upcomingSectionHidden ||
  result.pastSectionHidden ||
  result.upcomingCreateHidden ||
  result.pastCreateHidden;
console.log(bad ? "\nFAIL" : "\nPASS: duplicate skipped (no POST, override offered), non-matching event still created, \"Create anyway\" posted it, locations carried through, both bodies link back to the source email via a single lookup, no console errors");
process.exit(bad ? 1 : 0);
