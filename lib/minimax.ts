import OpenAI from "openai";
import { runExtraction } from "./extract.js";
import { ConfigError } from "./errors.js";

const MINIMAX_BASE_URL =
  process.env.MINIMAX_BASE_URL ?? "https://api.minimax.io/v1";

export async function extractEventsWithMiniMax(
  email: string,
  model: string,
) {
  const apiKey = process.env.MINIMAX_API_KEY;
  if (!apiKey) {
    throw new ConfigError("MINIMAX_API_KEY is not configured");
  }

  const client = new OpenAI({ apiKey, baseURL: MINIMAX_BASE_URL });
  return runExtraction(client, model, email, {
    thinking: { type: "disabled" },
  });
}
