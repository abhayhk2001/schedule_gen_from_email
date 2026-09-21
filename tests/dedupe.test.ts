import { describe, it, expect } from "vitest";
import {
  normalizeSubject,
  subjectsSimilar,
  findDuplicate,
} from "../public/outlook-addin/dedupe.js";

// Shapes a calendarView entry the way Graph returns it, with the padded
// sub-second precision the real API sends.
const existing = (
  subject: string,
  start: string,
  end: string,
  isAllDay = false,
) => ({
  id: "existing-1",
  subject,
  start: { dateTime: `${start}:00.0000000`, timeZone: "America/Chicago" },
  end: { dateTime: `${end}:00.0000000`, timeZone: "America/Chicago" },
  isAllDay,
  webLink: "https://outlook.office.com/calendar/item/existing-1",
});

// Shapes the payload mapToGraphFields builds.
const target = (
  subject: string,
  start: string,
  end: string,
  isAllDay = false,
) => ({
  subject,
  start: { dateTime: `${start}:00`, timeZone: "America/Chicago" },
  end: { dateTime: `${end}:00`, timeZone: "America/Chicago" },
  isAllDay,
});

describe("normalizeSubject", () => {
  it("strips case and punctuation", () => {
    expect(normalizeSubject("After Hours In Chicago!")).toBe(
      "after hours in chicago",
    );
  });

  it("collapses whitespace runs", () => {
    expect(normalizeSubject("  Team   standup\n\n")).toBe("team standup");
  });

  it("returns empty for non-strings", () => {
    expect(normalizeSubject(null)).toBe("");
    expect(normalizeSubject(undefined)).toBe("");
  });
});

describe("subjectsSimilar", () => {
  it("matches identical subjects", () => {
    expect(subjectsSimilar("Team standup", "Team standup")).toBe(true);
  });

  it("matches across case and punctuation differences", () => {
    expect(
      subjectsSimilar("After Hours in Chicago", "After Hours In Chicago!"),
    ).toBe(true);
  });

  it("matches when one subject contains the other", () => {
    expect(subjectsSimilar("After Hours", "After Hours in Chicago")).toBe(true);
  });

  it("rejects unrelated subjects", () => {
    expect(subjectsSimilar("Team standup", "Dentist appointment")).toBe(false);
  });

  it("never matches an empty subject", () => {
    expect(subjectsSimilar("", "Team standup")).toBe(false);
    expect(subjectsSimilar(null, null)).toBe(false);
  });
});

describe("findDuplicate", () => {
  const slot = ["2026-10-01T17:30", "2026-10-01T19:30"] as const;

  it("matches an identical event in the same slot", () => {
    const found = findDuplicate(
      [existing("After Hours in Chicago", ...slot)],
      target("After Hours in Chicago", ...slot),
    );
    expect(found?.webLink).toBe(
      "https://outlook.office.com/calendar/item/existing-1",
    );
  });

  it("matches despite case and punctuation differences", () => {
    expect(
      findDuplicate(
        [existing("After Hours In Chicago!", ...slot)],
        target("After Hours in Chicago", ...slot),
      ),
    ).not.toBeNull();
  });

  it("does NOT match an unrelated event occupying the same slot", () => {
    expect(
      findDuplicate(
        [existing("Dentist appointment", ...slot)],
        target("After Hours in Chicago", ...slot),
      ),
    ).toBeNull();
  });

  it("does NOT match the same title at a different start", () => {
    expect(
      findDuplicate(
        [existing("After Hours in Chicago", "2026-10-01T18:00", "2026-10-01T19:30")],
        target("After Hours in Chicago", ...slot),
      ),
    ).toBeNull();
  });

  it("does NOT match the same title with a different end", () => {
    expect(
      findDuplicate(
        [existing("After Hours in Chicago", "2026-10-01T17:30", "2026-10-01T20:00")],
        target("After Hours in Chicago", ...slot),
      ),
    ).toBeNull();
  });

  it("does NOT match an all-day event against a timed one", () => {
    expect(
      findDuplicate(
        [existing("Offsite", "2026-10-05T00:00", "2026-10-06T00:00", true)],
        target("Offsite", "2026-10-05T00:00", "2026-10-06T00:00", false),
      ),
    ).toBeNull();
  });

  it("matches two all-day events on the same date", () => {
    expect(
      findDuplicate(
        [existing("Offsite", "2026-10-05T00:00", "2026-10-06T00:00", true)],
        target("Offsite", "2026-10-05T00:00", "2026-10-06T00:00", true),
      ),
    ).not.toBeNull();
  });

  it("picks the matching entry out of a busy window", () => {
    const found = findDuplicate(
      [
        existing("Dentist appointment", ...slot),
        existing("After Hours in Chicago", ...slot),
      ],
      target("After Hours in Chicago", ...slot),
    );
    expect(found?.subject).toBe("After Hours in Chicago");
  });

  it("returns null for an empty or malformed window", () => {
    expect(findDuplicate([], target("Anything", ...slot))).toBeNull();
    expect(findDuplicate(null, target("Anything", ...slot))).toBeNull();
  });

  it("returns null when the target has no start", () => {
    expect(
      findDuplicate([existing("Team standup", ...slot)], {
        subject: "Team standup",
        start: {},
        end: {},
        isAllDay: false,
      }),
    ).toBeNull();
  });
});
