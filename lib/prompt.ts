import { SUPPORTED_MODELS, type SupportedModel } from "./types.js";

export { SUPPORTED_MODELS, type SupportedModel };

export const SYSTEM_PROMPT = `You are an event extractor. Given the body of an email, identify every distinct calendar event and return them as a JSON object with an "events" array.

Fields per event:
- date:        ISO YYYY-MM-DD, the start date.
- time:        24h HH:MM, the start time. Null when whole_day=true.
- timezone:    IANA name inferred from location/sender (e.g. "America/Chicago"). Null if not inferable.
- event_name:  short title.
- location:    venue, address, or room. Literal "Virtual" for online-only events. Null if neither can be inferred.
- description: one-sentence summary.
- whole_day:   true when the entry has NO specific time-of-day on that date.
- end_date:    ISO YYYY-MM-DD, the end date. Null when the entry spans only one day.
- end_time:    24h HH:MM, the end time. Null when whole_day=true or no end time is given.

Rules:
1. Same-day event with explicit end time ("2pm to 3pm", "9am-5pm"):
     One entry. whole_day=false, time=<start>, end_time=<end>, end_date=null.

2. Date range with no times ("Oct 5 to Oct 8", "Oct 10-20"):
     One entry per calendar day in the range. Each entry has whole_day=true,
     time=null, end_time=null, end_date=null. The date carries the day.

3. Date range with "each day" or "daily" times ("Oct 10-12, 9am-5pm each day",
   "Mon Oct 5 to Wed Oct 7, daily 10am-4pm"):
     One entry per calendar day. Each entry has whole_day=false,
     time=<start-of-day>, end_time=<end-of-day>, end_date=null.

4. Independent events on different dates ("Day 1: ... Day 2: ...", or two
   distinct events described separately):
     One entry per event. Each entry gets its own date/time/end_time/whole_day
     based on rules 1-3 applied to that single event.

5. A bare registration deadline, RSVP date, or "registration closes on ..." is
   NOT a calendar event. Skip it.

6. If the event is online only — Zoom, Teams, Google Meet, WebEx, or text like
   "virtual", "online", "join via the link" with no physical venue mentioned —
   set location to the literal string "Virtual". If both a venue and a virtual
   option are mentioned, prefer the physical venue.

7. If no event is mentioned, return { "events": [] }.

8. Resolve relative phrases ("tomorrow", "next Tuesday") using today's date,
   provided below. Convert "3:00 PM" to "15:00". Never invent fields.

Return ONLY the JSON object, no prose or markdown.`;

export const responseSchema = {
  type: "object",
  additionalProperties: false,
  required: ["events"],
  properties: {
    events: {
      type: "array",
      items: {
        type: "object",
        additionalProperties: false,
        required: [
          "date",
          "time",
          "timezone",
          "event_name",
          "location",
          "description",
          "whole_day",
          "end_date",
          "end_time",
        ],
        properties: {
          date: {
            type: ["string", "null"],
            description: "ISO 8601 date in YYYY-MM-DD format. Start date.",
          },
          time: {
            type: ["string", "null"],
            description:
              "24-hour time in HH:MM format (e.g. 15:00). Null when whole_day=true.",
          },
          timezone: {
            type: ["string", "null"],
            description:
              'IANA timezone name (e.g. "America/Chicago"). Null if it cannot be inferred.',
          },
          event_name: {
            type: ["string", "null"],
            description: "Short title for the event.",
          },
          location: {
            type: ["string", "null"],
            description:
              'Venue, address, or room. Literal "Virtual" for online-only events. Null if neither can be inferred.',
          },
          description: {
            type: ["string", "null"],
            description: "One-sentence summary of the event.",
          },
          whole_day: {
            type: "boolean",
            description:
              "True when the entry has no specific time-of-day on that date.",
          },
          end_date: {
            type: ["string", "null"],
            description:
              "ISO 8601 end date YYYY-MM-DD. Null when the entry spans only one day.",
          },
          end_time: {
            type: ["string", "null"],
            description:
              "24-hour end time HH:MM. Null when whole_day=true or no end time is given.",
          },
        },
      },
    },
  },
} as const;

export function isSupportedModel(value: unknown): value is SupportedModel {
  return (
    typeof value === "string" &&
    (SUPPORTED_MODELS as readonly string[]).includes(value)
  );
}
