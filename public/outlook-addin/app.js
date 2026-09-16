const $ = (id) => document.getElementById(id);

const btn = $("extract-btn");
const status = $("status");
const results = $("results");
const eventsEl = $("events");

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

function renderEvents(events) {
  results.classList.remove("hidden");
  eventsEl.innerHTML = "";

  if (!events.length) {
    const empty = document.createElement("div");
    empty.className = "empty";
    empty.textContent = "No events detected in this email.";
    eventsEl.appendChild(empty);
    return;
  }

  for (const ev of events) {
    const card = document.createElement("article");
    card.className = "event";

    const h3 = document.createElement("h3");
    h3.textContent = ev.event_name ?? "Untitled event";
    card.appendChild(h3);

    const when = document.createElement("p");
    when.className = "when";
    const date = ev.date ?? "Unknown date";
    const time = formatTime(ev);
    when.textContent = `${date} · ${time}`;
    if (ev.timezone) {
      const tz = document.createElement("span");
      tz.className = "tz";
      tz.textContent = ev.timezone;
      when.appendChild(tz);
    }
    card.appendChild(when);

    if (ev.description) {
      const desc = document.createElement("p");
      desc.className = "desc";
      desc.textContent = ev.description;
      card.appendChild(desc);
    }

    eventsEl.appendChild(card);
  }
}

Office.onReady((info) => {
  if (info.host !== Office.HostType.Outlook) return;

  btn.disabled = false;
  btn.addEventListener("click", async () => {
    btn.disabled = true;
    clearStatus();
    results.classList.add("hidden");

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

      clearStatus();
      renderEvents(data.events ?? []);
    } catch (err) {
      setStatus("error", err?.message ?? "Unknown error");
    } finally {
      btn.disabled = false;
    }
  });
});
