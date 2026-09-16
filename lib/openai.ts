import OpenAI from "openai";
import { runExtraction } from "./extract.js";
import { ConfigError } from "./errors.js";

export async function extractEventsWithOpenAI(
  email: string,
  model: string,
) {
  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    throw new ConfigError("OPENAI_API_KEY is not configured");
  }

  const client = new OpenAI({ apiKey });
  return runExtraction(client, model, email);
}
