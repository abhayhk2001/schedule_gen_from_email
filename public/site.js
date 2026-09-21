const $ = (id) => document.getElementById(id);

const form = $("extract-form");
const modelEl = $("model");
const emailEl = $("email");
const submitBtn = $("submit-btn");
const sampleBtn = $("sample-btn");
const statusEl = $("status");
const resultsEl = $("results");
const eventsEl = $("events");

const SAMPLE_EMAIL = `Subject: Hackathon
From: organizer@uni.edu

Our 3-day hackathon runs Oct 10-12, 2026, 9am-5pm each day in Siebel Center.
Lunch is provided.

There is also a kickoff dinner on Oct 9 at 6:30pm in the Illini Union.

Registration closes Sept 30 — please sign up before then.`;

const monthNames = [
  "Jan", "Feb", "Mar", "Apr", "May", "Jun",
  "Jul", "Aug", "Sep", "Oct", "Nov", "Dec",
];

function formatDate(ev) {
  if (!ev.date) return "Unknown date";
  const [, m, d] = ev.date.split("-").map(Number);
  if (!m || !d) return ev.date;
  return `${monthNames[m - 1]} ${d}`;
}

function formatTime(ev) {
  if (ev.whole_day) return "All day";
  const start = ev.time ?? "?";
  if (ev.end_time) return `${start} – ${ev.end_time}`;
  return start;
}

function escapeText(s) {
  return String(s ?? "");
}

function setStatus(kind, text) {
  statusEl.className = `status ${kind}`;
  statusEl.textContent = text;
}
function clearStatus() {
  statusEl.className = "status hidden";
  statusEl.textContent = "";
}

function renderEvents(events) {
  eventsEl.innerHTML = "";

  if (!events.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "No events detected.";
    eventsEl.appendChild(empty);
    return;
  }

  for (const ev of events) {
    const card = document.createElement("article");
    card.className = "event";

    const title = document.createElement("h4");
    title.textContent = ev.event_name ?? "Untitled event";
    card.appendChild(title);

    if (ev.location) {
      const location = document.createElement("p");
      location.className = "location";
      location.textContent = ev.location;
      card.appendChild(location);
    }

    const meta = document.createElement("p");
    meta.className = "when";
    meta.textContent = `${formatDate(ev)} · ${formatTime(ev)}`;
    card.appendChild(meta);

    if (ev.timezone) {
      const tz = document.createElement("span");
      tz.className = "muted small";
      tz.textContent = ` ${ev.timezone}`;
      meta.appendChild(tz);
    }

    if (ev.description) {
      const desc = document.createElement("p");
      desc.className = "desc";
      desc.textContent = ev.description;
      card.appendChild(desc);
    }

    eventsEl.appendChild(card);
  }
}

async function extract(email, model) {
  const resp = await fetch("/api/extract-event", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ email, model }),
  });

  const data = await resp.json().catch(() => ({}));
  if (!resp.ok) {
    throw new Error(data?.error ?? `HTTP ${resp.status}`);
  }
  return data;
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();
  const email = emailEl.value.trim();
  const model = modelEl.value;
  if (!email) return;

  submitBtn.disabled = true;
  sampleBtn.disabled = true;
  resultsEl.classList.add("hidden");
  setStatus("loading", "Calling API…");

  try {
    const data = await extract(email, model);
    const events = Array.isArray(data.events) ? data.events : [];
    renderEvents(events);
    resultsEl.classList.remove("hidden");
    clearStatus();
  } catch (err) {
    setStatus("error", err?.message ?? "Unknown error");
  } finally {
    submitBtn.disabled = false;
    sampleBtn.disabled = false;
  }
});

sampleBtn.addEventListener("click", () => {
  emailEl.value = SAMPLE_EMAIL;
  emailEl.focus();
});
