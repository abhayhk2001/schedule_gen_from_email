// Builds the Graph event body, kept free of DOM and Office globals so the
// escaping rules can be unit-tested directly (tests/eventbody.test.ts).
//
// When the source email could be resolved the body becomes HTML carrying a
// clickable link back to it. Everything interpolated into that HTML — the
// LLM-written description, the email subject, the URL — is untrusted and is
// escaped before it goes in.

export function escapeHtml(value) {
  if (typeof value !== "string") return "";
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

// Only http(s) may reach an href. A javascript: or data: URL in that slot would
// execute when the user clicks the link in their calendar.
function safeHref(url) {
  if (typeof url !== "string") return null;
  const trimmed = url.trim();
  return /^https?:\/\//i.test(trimmed) ? trimmed : null;
}

// `source` is state.source: { restId, subject, webLink }. Anything falsy, or a
// source whose webLink could not be resolved, yields exactly the plain-text
// body the add-in produced before this feature existed.
export function buildEventBody(description, source) {
  const text = typeof description === "string" ? description : "";
  const href = safeHref(source?.webLink);

  if (!href) {
    return { contentType: "Text", content: text };
  }

  const label = escapeHtml(source.subject || "").trim() || "Open the email";
  const parts = [];
  if (text) parts.push(`<p>${escapeHtml(text)}</p>`);
  parts.push("<hr>");
  parts.push(
    `<p>From email: <a href="${escapeHtml(href)}">${label}</a></p>`,
  );

  return { contentType: "HTML", content: parts.join("") };
}
