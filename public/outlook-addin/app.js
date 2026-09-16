const $ = (id) => document.getElementById(id);

const btn = $("extract-btn");
const createBtn = $("create-btn");
const status = $("status");
const results = $("results");
const eventsEl = $("events");

const state = {
  events: [],
  removed: new Set(),
  results: null,
};

const GRAPH_BASE = "https://graph.microsoft.com/v1.0";

function setStatus(kind, text) {
  status.className = `status ${kind}`;
  status.textContent = text;
  status.classList.remove("hidden");
}

function clearStatus() {
  status.classList.add("hidden");
  status.textContent = "";
}

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, (c) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;",
  }[c]));
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

function mapToGraphEvent(ev) {
  const tz = ev.timezone || "UTC";
  const subject = ev.event_name || "Untitled event";
  const body = {
    contentType: "text",
    content: ev.description || "",
  };
  if (ev.whole_day) {
    const startDate = ev.date;
    const endDate = ev.end_date || addDaysToISO(ev.date, 1);
    return {
      subject,
      body,
      start: { dateTime: startDate, timeZone: tz },
      end: { dateTime: endDate, timeZone: tz },
      isAllDay: true,
    };
  }
  const startDateTime = `${ev.date}T${ev.time}:00`;
  const endDate = ev.end_date || ev.date;
  const endTime = ev.end_time || addHoursToHHMM(ev.time || "00:00", 1);
  const endDateTime = `${endDate}T${endTime}:00`;
  return {
    subject,
    body,
    start: { dateTime: startDateTime, timeZone: tz },
    end: { dateTime: endDateTime, timeZone: tz },
    isAllDay: false,
  };
}

async function getAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync(
      { isRest: true },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(
            new Error(
              result.error?.message ?? "Failed to acquire callback token",
            ),
          );
        }
      },
    );
  });
}

async function postEvent(token, payload) {
  const resp = await fetch(`${GRAPH_BASE}/me/events`, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify(payload),
  });
  if (!resp.ok) {
    let detail = "";
    try {
      const j = await resp.json();
      detail = j?.error?.message ?? "";
    } catch {
      detail = await resp.text().catch(() => "");
    }
    throw new Error(`HTTP ${resp.status}${detail ? `: ${detail}` : ""}`);
  }
  return resp.json();
}

async function createEvents(events) {
  const token = await getAccessToken();
  const results = [];
  for (const ev of events) {
    try {
      const payload = mapToGraphEvent(ev);
      const data = await postEvent(token, payload);
      results.push({ ev, status: "success", webLink: data.webLink });
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

    if (ev.description) {
      const desc = document.createElement("p");
      desc.className = "desc";
      desc.textContent = ev.description;
      body.appendChild(desc);
    }

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

    const when = document.createElement("p");
    when.className = "when";
    when.textContent = `${r.ev.date ?? "?"} · ${formatTime(r.ev)}`;
    body.appendChild(when);

    if (r.status === "success" && r.webLink) {
      const link = document.createElement("a");
      link.href = r.webLink;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      link.textContent = "Open in Outlook";
      body.appendChild(link);
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
    const token = await getAccessToken();
    const data = await postEvent(token, mapToGraphEvent(r.ev));
    state.results[idx] = { ev: r.ev, status: "success", webLink: data.webLink };
  } catch (err) {
    state.results[idx] = {
      ev: r.ev,
      status: "error",
      error: err?.message ?? String(err),
    };
  }
  renderResults();
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  btn.disabled = false;
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    createBtn.classList.add("hidden");
    clearStatus();
    results.classList.add("hidden");
    state.events = [];
    state.removed.clear();
    state.results = null;

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
