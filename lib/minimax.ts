import OpenAI from "openai";
import type { Event, ExtractResponse } from "./types.js";

const MINIMAX_BASE_URL =
  process.env.MINIMAX_BASE_URL ?? "https://api.minimax.io/v1";
const MINIMAX_MODEL = process.env.MINIMAX_MODEL ?? "MiniMax-M3";

const eventItemSchema = {
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
} as const;

const responseSchema = {
  type: "object",
  additionalProperties: false,
  required: ["events"],
  properties: {
    events: {
      type: "array",
      items: eventItemSchema,
    },
  },
} as const;

const SYSTEM_PROMPT = `You are an event extractor. Given the body of an email, identify every distinct calendar event mentioned and return them as a JSON object with an "events" array.

Rules:
- One entry per independent event. Two times on different dates = two entries.
- "date" must be ISO YYYY-MM-DD. Resolve relative phrases ("tomorrow", "next Tuesday") using today's date, provided below.
- "time" must be 24h HH:MM. Convert "3:00 PM" to "15:00".
- "timezone" should be an IANA name inferred from location hints, sender context, or explicit mentions (e.g. "America/Chicago", "America/New_York", "UTC"). Use null if it cannot be inferred.
- "event_name" is a short title; "description" is one sentence.
- If no event is mentioned, return { "events": [] }.
- Never invent fields. Return ONLY the JSON object, no prose or markdown.`;

export async function extractEvents(email: string): Promise<ExtractResponse> {
  const apiKey = process.env.MINIMAX_API_KEY;
  if (!apiKey) {
    throw new Error("MINIMAX_API_KEY is not configured");
  }

  const client = new OpenAI({ apiKey, baseURL: MINIMAX_BASE_URL });
  const today = new Date().toISOString().slice(0, 10);

  type MiniMaxBody = Parameters<typeof client.chat.completions.create>[0] & {
    extra_body?: Record<string, unknown>;
  };

  const completion = await client.chat.completions.create({
    model: MINIMAX_MODEL,
    response_format: {
      type: "json_schema",
      json_schema: {
        name: "extract_events",
        strict: true,
        schema: responseSchema,
      },
    },
    messages: [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content: `Today's date is ${today}.\n\nEmail:\n"""\n${email}\n"""`,
      },
    ],
    temperature: 0,
    // MiniMax-specific: disable thinking so the response is clean JSON
    // instead of being prefixed with <thinking>...</thinking> tags.
    extra_body: { thinking: { type: "disabled" } },
  } as MiniMaxBody);

  const raw = (completion as OpenAI.Chat.ChatCompletion).choices[0]?.message
    ?.content;
  if (!raw) {
    return { events: [] };
  }

  const parsed = JSON.parse(raw) as { events: Event[] };
  return { events: Array.isArray(parsed.events) ? parsed.events : [] };
}
