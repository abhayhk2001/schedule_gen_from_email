import { initTheme, setThemeOverride, effectiveMode } from "./theme.js";

const $ = (id) => document.getElementById(id);

const btn = $("extract-btn");
const createBtn = $("create-btn");
const status = $("status");
const results = $("results");
const eventsEl = $("events");
const themeBtn = $("theme-toggle");
const themeIcon = $("theme-icon");

const FAST_LANE_BROKEN_KEY = "addCalEvent.fastLaneBroken";
const FAST_LANE_TIMEOUT_MS = 4_000;

// Diagnostics go to the console only. The in-pane debug panel existed to
// diagnose the Office SSO hang from a host with no DevTools; that is solved,
// and the panel was taking up most of the task pane.
function pushLog(level, text) {
  const fn = level === "error" ? console.error
    : level === "warn" ? console.warn
    : level === "success" || level === "info" ? console.info
    : console.debug;
  fn(`[addCalEvent:${level}] ${text}`);
}

const GRAPH_RESOURCE = "https://graph.microsoft.com";
const GRAPH_DEFAULT_SCOPES = ["openid", "profile", "offline_access", "User.Read", "Calendars.ReadWrite"];

const SUN_SVG =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="4"/><path d="M12 2v2"/><path d="M12 20v2"/><path d="m4.93 4.93 1.41 1.41"/><path d="m17.66 17.66 1.41 1.41"/><path d="M2 12h2"/><path d="M20 12h2"/><path d="m4.93 19.07 1.41-1.41"/><path d="m17.66 6.34 1.41-1.41"/></svg>';
const MOON_SVG =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>';

function paintThemeButton() {
  if (!themeBtn || !themeIcon) return;
  const mode = effectiveMode();
  const isDark = mode === "dark";
  themeIcon.innerHTML = isDark ? SUN_SVG : MOON_SVG;
  themeBtn.setAttribute("aria-label", isDark ? "Switch to light theme" : "Switch to dark theme");
  themeBtn.title = isDark ? "Switch to light" : "Switch to dark";
}

function onThemeToggleClick() {
  const next = effectiveMode() === "dark" ? "light" : "dark";
  setThemeOverride(next);
  paintThemeButton();
}

if (themeBtn) {
  themeBtn.addEventListener("click", onThemeToggleClick);
  paintThemeButton();
}

const state = {
  events: [],
  removed: new Set(),
  results: null,
};

function setStatus(kind, text) {
  status.className = `status ${kind}`;
  status.textContent = text;
  status.classList.remove("hidden");
}

function clearStatus() {
  status.classList.add("hidden");
  status.textContent = "";
}

function formatTime(ev) {
  if (ev.whole_day) return "All day";
  const start = ev.time ?? "?";
  if (ev.end_time) return `${start} – ${ev.end_time}`;
  return start;
}

function pad(n) {
  return String(n).padStart(2, "0");
}

function addHoursToHHMM(hhmm, hours) {
  const [h, m] = hhmm.split(":").map(Number);
  const total = h * 60 + m + Math.round(hours * 60);
  const nh = Math.floor((total / 60) % 24);
  const nm = total % 60;
  return `${pad(nh)}:${pad(nm)}`;
}

function addDaysToISO(yyyymmdd, days) {
  const d = new Date(`${yyyymmdd}T00:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

function getLocalTimeZone() {
  try {
    return Intl.DateTimeFormat().resolvedOptions().timeZone || "UTC";
  } catch {
    return "UTC";
  }
}

function mapToGraphFields(ev) {
  const subject = ev.event_name || "Untitled event";
  const body = ev.description || "";
  const timeZone = ev.timezone || getLocalTimeZone();

  if (ev.whole_day) {
    return {
      subject,
      body: { contentType: "Text", content: body },
      start: { dateTime: ev.date, timeZone },
      end: {
        dateTime: ev.end_date || addDaysToISO(ev.date, 1),
        timeZone,
      },
      isAllDay: true,
    };
  }

  const startTime = ev.time || "00:00";
  const endTime = ev.end_time || addHoursToHHMM(startTime, 1);

  return {
    subject,
    body: { contentType: "Text", content: body },
    start: { dateTime: `${ev.date}T${startTime}:00`, timeZone },
    end: {
      dateTime: `${ev.end_date || ev.date}T${endTime}:00`,
      timeZone,
    },
    isAllDay: false,
  };
}

// Office.auth.getAccessToken never calls back on Outlook for Mac. Once we have
// seen that on this host, stop paying the timeout on every subsequent event.
function markFastLaneBroken(why) {
  try {
    window.localStorage.setItem(FAST_LANE_BROKEN_KEY, "1");
  } catch {}
  console.debug(`[sso] fast lane marked broken on this host (${why})`);
}

function isFastLaneBroken() {
  try {
    return window.localStorage.getItem(FAST_LANE_BROKEN_KEY) === "1";
  } catch {
    return false;
  }
}

function getOfficeAccessToken(scopes) {
  return new Promise((resolve, reject) => {
    if (!Office?.auth?.getAccessToken) {
      markFastLaneBroken("not available");
      pushLog("info", "Office.auth.getAccessToken unavailable; skipping fast lane");
      reject(new Error("Office.auth.getAccessToken is not available in this host."));
      return;
    }
    const options = {
      allowSignInPrompt: true,
      allowConsentPrompt: true,
      allowMultipleSignInPrompt: false,
      scopes,
    };
    options.forMSGraphAccess = true;
    pushLog(
      "info",
      `fast lane: Office.auth.getAccessToken timeoutMs=${FAST_LANE_TIMEOUT_MS}`,
    );

    let settled = false;
    const timer = setTimeout(() => {
      if (settled) return;
      settled = true;
      markFastLaneBroken("timed out");
      pushLog(
        "warn",
        `fast lane timed out after ${FAST_LANE_TIMEOUT_MS}ms — using the Office dialog instead`,
      );
      reject(new Error(`SSO timeout after ${FAST_LANE_TIMEOUT_MS}ms (no callback fired)`));
    }, FAST_LANE_TIMEOUT_MS);

    Office.auth.getAccessToken(options, (result) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      if (result?.status === "succeeded") {
        pushLog("success", `fast lane: Office token acquired`);
        resolve(result.value);
      } else {
        const err = result?.error ?? {};
        markFastLaneBroken(`code ${err.code ?? "?"}`);
        pushLog(
          "warn",
          `fast lane failed code=${err.code ?? "?"} name=${err.name ?? "?"} — using the Office dialog instead`,
        );
        reject(
          new Error(`SSO error (${err.code ?? "?"}): ${err.message ?? "Unknown"}`),
        );
      }
    });
  });
}

async function getGraphToken(scopes, opts = {}) {
  let useFastLane = !opts.skipFastLane;
  if (useFastLane && isFastLaneBroken()) {
    useFastLane = false;
    pushLog("info", "fast lane skipped (Office SSO already known to fail on this host)");
  }
  if (useFastLane) {
    try {
      return await getOfficeAccessToken(scopes);
    } catch (err) {
      console.warn("[sso] fast lane failed", err);
    }
  }

  pushLog("info", "auth: opening the Office dialog sign-in flow");
  const msalModule = await import("./msal.js");
  return msalModule.msalLogin(scopes, {
    onProgress(stage) {
      pushLog("info", `msal: ${stage}`);
    },
  });
}

async function createGraphEvent(ev, { retry = true, skipFastLane = false } = {}) {
  const event = mapToGraphFields(ev);
  const token = await getGraphToken(GRAPH_DEFAULT_SCOPES, { skipFastLane });

  pushLog("info", `POST ${GRAPH_RESOURCE}/v1.0/me/events`);
  let resp;
  try {
    resp = await fetch(`${GRAPH_RESOURCE}/v1.0/me/events`, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify(event),
    });
  } catch (err) {
    pushLog("error", `Graph POST network error: ${err?.message ?? String(err)}`);
    throw err;
  }

  pushLog("info", `Graph response status=${resp.status}`);

  if (resp.status === 401 && retry) {
    pushLog("warn", "Graph returned 401; discarding the access token and retrying once");
    const msalModule = await import("./msal.js");
    // Keep the refresh token — the grant is usually still valid, so the retry
    // can renew silently. Skip the fast lane, since an Office SSO token is
    // what produced this 401 whenever that lane is the one in use.
    msalModule.invalidateAccessToken();
    return createGraphEvent(ev, { retry: false, skipFastLane: true });
  }

  if (!resp.ok) {
    const payload = await resp.json().catch(() => ({}));
    const code = payload?.error?.code ?? `HTTP ${resp.status}`;
    const message = payload?.error?.message ?? resp.statusText;
    pushLog(
      "error",
      `Graph error code=${code} message="${message}"`,
    );
    throw new Error(`${code}: ${message}`);
  }

  const data = await resp.json();
  pushLog("success", `Graph event created id=${data.id ?? "?"}`);
  return { itemId: data.id ?? null, webLink: data.webLink ?? null };
}

async function createEvents(events) {
  const results = [];
  for (const ev of events) {
    try {
      const { itemId, webLink } = await createGraphEvent(ev);
      results.push({ ev, status: "success", itemId, webLink });
    } catch (err) {
      results.push({ ev, status: "error", error: err?.message ?? String(err) });
    }
  }
  return results;
}

function renderEvents() {
  results.classList.remove("hidden");
  eventsEl.innerHTML = "";

  if (!state.events.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "No events detected in this email.";
    eventsEl.appendChild(empty);
    createBtn.classList.add("hidden");
    return;
  }

  state.events.forEach((ev, idx) => {
    const removed = state.removed.has(idx);
    const card = document.createElement("article");
    card.className = `event ${removed ? "removed" : ""}`;

    const row = document.createElement("div");
    row.className = "event-row";

    const checkbox = document.createElement("input");
    checkbox.type = "checkbox";
    checkbox.checked = !removed;
    checkbox.disabled = removed;
    checkbox.addEventListener("change", () => {
      if (checkbox.checked) state.removed.delete(idx);
      else state.removed.add(idx);
      renderEvents();
    });
    row.appendChild(checkbox);

    const body = document.createElement("div");
    body.className = "event-body";

    const h3 = document.createElement("h3");
    h3.textContent = ev.event_name ?? "Untitled event";
    body.appendChild(h3);

    if (ev.location) {
      const location = document.createElement("p");
      location.className = "location";
      location.textContent = ev.location;
      body.appendChild(location);
    }

    const when = document.createElement("p");
    when.className = "when";
    const date = ev.date ?? "Unknown date";
    const time = formatTime(ev);
    when.textContent = `${date} · ${time}`;
    if (ev.timezone) {
      const tzEl = document.createElement("span");
      tzEl.className = "tz";
      tzEl.textContent = ev.timezone;
      when.appendChild(tzEl);
    }
    body.appendChild(when);

    row.appendChild(body);

    const removeBtn = document.createElement("button");
    removeBtn.type = "button";
    removeBtn.className = "remove-btn";
    removeBtn.title = removed ? "Restore" : "Remove";
    removeBtn.textContent = removed ? "+" : "×";
    removeBtn.addEventListener("click", () => {
      if (state.removed.has(idx)) state.removed.delete(idx);
      else state.removed.add(idx);
      renderEvents();
      renderCreateButton();
    });
    row.appendChild(removeBtn);

    card.appendChild(row);
    eventsEl.appendChild(card);
  });

  renderCreateButton();
}

function renderCreateButton() {
  const remaining = state.events.length - state.removed.size;
  if (remaining > 0) {
    createBtn.classList.remove("hidden");
    createBtn.textContent = `Create ${remaining} event${remaining === 1 ? "" : "s"} in calendar`;
    createBtn.disabled = false;
  } else {
    createBtn.classList.add("hidden");
  }
}

function renderResults() {
  if (!state.results) return;
  results.classList.remove("hidden");
  eventsEl.innerHTML = "";

  state.results.forEach((r, idx) => {
    const card = document.createElement("article");
    card.className = `event result ${r.status}`;

    const marker = document.createElement("div");
    marker.className = "result-marker";
    marker.textContent = r.status === "success" ? "✓" : "✗";
    card.appendChild(marker);

    const body = document.createElement("div");
    body.className = "event-body";

    const h3 = document.createElement("h3");
    h3.textContent = r.ev.event_name ?? "Untitled event";
    body.appendChild(h3);

    if (r.ev.location) {
      const location = document.createElement("p");
      location.className = "location";
      location.textContent = r.ev.location;
      body.appendChild(location);
    }

    const when = document.createElement("p");
    when.className = "when";
    when.textContent = `${r.ev.date ?? "?"} · ${formatTime(r.ev)}`;
    body.appendChild(when);

    if (r.status === "success") {
      const note = document.createElement("p");
      note.className = "desc";
      note.textContent = "Created — open your Outlook calendar to view.";
      body.appendChild(note);
      if (r.webLink) {
        const link = document.createElement("a");
        link.href = r.webLink;
        link.target = "_blank";
        link.rel = "noopener noreferrer";
        link.textContent = "Open in Outlook on the web";
        body.appendChild(link);
      }
    } else if (r.status === "error") {
      const err = document.createElement("p");
      err.className = "error-text";
      err.textContent = r.error;
      body.appendChild(err);

      const retry = document.createElement("button");
      retry.type = "button";
      retry.className = "retry-btn";
      retry.textContent = "Retry";
      retry.addEventListener("click", () => retryOne(idx));
      body.appendChild(retry);
    }

    card.appendChild(body);
    eventsEl.appendChild(card);
  });
}

async function retryOne(idx) {
  const r = state.results[idx];
  if (!r || r.status !== "error") return;
  state.results[idx] = { ev: r.ev, status: "pending" };
  renderResults();
  try {
    const { itemId, webLink } = await createGraphEvent(r.ev);
    state.results[idx] = { ev: r.ev, status: "success", itemId, webLink };
  } catch (err) {
    state.results[idx] = {
      ev: r.ev,
      status: "error",
      error: err?.message ?? String(err),
    };
  }
  renderResults();
}

// Clears everything the pane is showing about the current message. The pinned
// pane survives moving to another email, so its contents have to be reset when
// the selected item changes or it would still be describing the old one.
function resetPane() {
  state.events = [];
  state.removed.clear();
  state.results = null;
  createBtn.classList.add("hidden");
  results.classList.add("hidden");
  eventsEl.innerHTML = "";
  clearStatus();
}

Office.onReady((info) => {
  pushLog(
    "info",
    `Office.onReady host=${info.host} platform=${info.platform ?? "?"}`,
  );
  pushLog(
    "info",
    `auth.getAccessToken available: ${typeof Office?.auth?.getAccessToken === "function"}${
      isFastLaneBroken() ? " (known broken on this host — skipping it)" : ""
    }`,
  );

  if (info.host !== Office.HostType.Outlook) {
    pushLog("warn", `not in an Outlook host (info.host=${info.host}); aborting`);
    return;
  }

  initTheme();
  paintThemeButton();

  // With SupportsPinning the pane stays open across messages, so react to the
  // selection changing instead of being torn down and rebuilt each time.
  if (typeof Office.context.mailbox.addHandlerAsync === "function") {
    Office.context.mailbox.addHandlerAsync(
      Office.EventType.ItemChanged,
      () => {
        pushLog("info", "selected item changed; clearing the pane");
        resetPane();
        btn.disabled = false;
      },
      (result) => {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          pushLog(
            "warn",
            `ItemChanged handler not registered: ${result.error?.message ?? "?"}`,
          );
        }
      },
    );
  }

  btn.disabled = false;
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    resetPane();

    try {
      const item = Office.context.mailbox.item;
      const subject = item.subject ?? "";
      const sender = item.sender?.emailAddress ?? "";

      setStatus("loading", "Reading email body…");

      const body = await new Promise((resolve, reject) => {
        item.body.getAsync(
          Office.CoercionType.Text,
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value ?? "");
            } else {
              reject(new Error(result.error?.message ?? "Failed to read body"));
            }
          },
        );
      });

      const composed = `Subject: ${subject}\nFrom: ${sender}\n\n${body}`;
      setStatus("loading", "Calling API…");

      const resp = await fetch("/api/extract-event", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ model: "gpt-4o-mini", email: composed }),
      });

      const data = await resp.json().catch(() => ({}));

      if (!resp.ok) {
        throw new Error(data?.error ?? `HTTP ${resp.status}`);
      }

      state.events = data.events ?? [];
      clearStatus();
      renderEvents();
    } catch (err) {
      setStatus("error", err?.message ?? "Unknown error");
    } finally {
      btn.disabled = false;
    }
  });

  createBtn.addEventListener("click", async () => {
    const remaining = state.events.filter((_, i) => !state.removed.has(i));
    if (!remaining.length) return;
    createBtn.disabled = true;
    btn.disabled = true;
    setStatus("loading", `Creating ${remaining.length} event(s)…`);
    try {
      state.results = await createEvents(remaining);
      clearStatus();
      renderResults();
    } catch (err) {
      setStatus("error", err?.message ?? "Unknown error");
    } finally {
      createBtn.disabled = false;
      btn.disabled = false;
    }
  });

});
