import OpenAI from "openai";
import type { Event, ExtractResponse } from "./types.js";
import { SYSTEM_PROMPT, responseSchema } from "./prompt.js";

export async function extractEventsWithOpenAI(
  email: string,
  model: string,
): Promise<ExtractResponse> {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new Error("OPENAI_API_KEY is not configured");
  }

  const client = new OpenAI({ apiKey });
  const today = new Date().toISOString().slice(0, 10);

  const completion = await client.chat.completions.create({
    model,
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
  });

  const raw = completion.choices[0]?.message?.content;
  if (!raw) {
    return { events: [] };
  }

  const parsed = JSON.parse(raw) as { events: Event[] };
  return { events: Array.isArray(parsed.events) ? parsed.events : [] };
}
