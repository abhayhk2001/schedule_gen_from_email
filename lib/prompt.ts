import { SUPPORTED_MODELS, type SupportedModel } from "./types.js";

export { SUPPORTED_MODELS, type SupportedModel };

export const SYSTEM_PROMPT = `You are an event extractor. Given the body of an email, identify every distinct calendar event mentioned and return them as a JSON object with an "events" array.

Rules:
- One entry per independent event. Two times on different dates = two entries.
- "date" must be ISO YYYY-MM-DD. Resolve relative phrases ("tomorrow", "next Tuesday") using today's date, provided below.
- "time" must be 24h HH:MM. Convert "3:00 PM" to "15:00".
- "timezone" should be an IANA name inferred from location hints, sender context, or explicit mentions (e.g. "America/Chicago", "America/New_York", "UTC"). Use null if it cannot be inferred.
- "event_name" is a short title; "description" is one sentence.
- If no event is mentioned, return { "events": [] }.
- Never invent fields. Return ONLY the JSON object, no prose or markdown.`;

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
        required: ["date", "time", "timezone", "event_name", "description"],
        properties: {
          date: {
            type: ["string", "null"],
            description: "ISO 8601 date in YYYY-MM-DD format.",
          },
          time: {
            type: ["string", "null"],
            description: "24-hour time in HH:MM format (e.g. 15:00).",
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
          description: {
            type: ["string", "null"],
            description: "One-sentence summary of the event.",
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
