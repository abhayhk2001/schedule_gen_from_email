// Duplicate detection for calendar events, kept free of DOM and Office globals
// so the matching rules can be unit-tested directly (tests/dedupe.test.ts).
//
// The rule: an existing calendar event is a duplicate only when it occupies the
// exact same slot — same all-day flag, same start and end to the minute — AND
// carries a near-identical title. A different time is never a duplicate, however
// the event is named, so a genuine double-booking is never mistaken for one.

const SIMILARITY_THRESHOLD = 0.8;

// Graph returns "2026-10-01T17:30:00.0000000"; mapToGraphFields emits
// "2026-10-01T17:30:00". Comparing to the minute puts both in one shape and
// drops the sub-second precision Graph pads on.
export function toMinute(dateTime) {
  return typeof dateTime === "string" ? dateTime.slice(0, 16) : "";
}

export function normalizeSubject(subject) {
  if (typeof subject !== "string") return "";
  return subject
    .toLowerCase()
    .normalize("NFKD")
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

export function subjectsSimilar(a, b) {
  const na = normalizeSubject(a);
  const nb = normalizeSubject(b);
  if (!na || !nb) return false;
  if (na === nb) return true;

  // Containment catches a subject that was truncated or extended on one side,
  // e.g. "After Hours" vs "After Hours in Chicago". Safe because the caller has
  // already required an exact time match.
  if (na.includes(nb) || nb.includes(na)) return true;

  const ta = new Set(na.split(" "));
  const tb = new Set(nb.split(" "));
  let shared = 0;
  for (const token of ta) if (tb.has(token)) shared += 1;
  return shared / Math.min(ta.size, tb.size) >= SIMILARITY_THRESHOLD;
}

function sameSlot(candidate, target) {
  return (
    Boolean(candidate?.isAllDay) === Boolean(target?.isAllDay) &&
    toMinute(candidate?.start?.dateTime) === toMinute(target?.start?.dateTime) &&
    toMinute(candidate?.end?.dateTime) === toMinute(target?.end?.dateTime)
  );
}

// `candidates` is the `value` array of a calendarView response, `target` the
// body mapToGraphFields built. Both sides must already be expressed in the same
// timezone — the caller guarantees that with a Prefer: outlook.timezone header.
export function findDuplicate(candidates, target) {
  if (!Array.isArray(candidates) || !target) return null;
  if (!toMinute(target.start?.dateTime)) return null;
  return (
    candidates.find(
      (c) => sameSlot(c, target) && subjectsSimilar(c?.subject, target.subject),
    ) ?? null
  );
}
