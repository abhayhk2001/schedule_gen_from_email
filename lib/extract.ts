import OpenAI from "openai";
import type { Event, ExtractResponse } from "./types.js";
import { SYSTEM_PROMPT, responseSchema } from "./prompt.js";

export async function runExtraction(
  client: OpenAI,
  model: string,
  email: string,
  extraBody?: Record<string, unknown>,
): Promise<ExtractResponse> {
  const today = new Date().toISOString().slice(0, 10);

  const body = {
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
    ...(extraBody ? { extra_body: extraBody } : {}),
  } as Parameters<typeof client.chat.completions.create>[0] & {
    extra_body?: Record<string, unknown>;
  };

  const completion = await client.chat.completions.create(body);
  const raw = (completion as OpenAI.Chat.ChatCompletion).choices[0]?.message
    ?.content;

  if (!raw) {
    return { events: [] };
  }

  try {
    const parsed = JSON.parse(raw) as { events: Event[] };
    return { events: Array.isArray(parsed.events) ? parsed.events : [] };
  } catch {
    return { events: [] };
  }
}
