import type { ExtractResponse, SupportedModel } from "./types.js";
import { extractEventsWithOpenAI } from "./openai.js";
import { extractEventsWithMiniMax } from "./minimax.js";
import { SUPPORTED_MODELS } from "./prompt.js";

export async function extractEvents(
  email: string,
  model: SupportedModel,
): Promise<ExtractResponse> {
  switch (model) {
    case "gpt-4o-mini":
      return extractEventsWithOpenAI(email, model);
    case "MiniMax-M3":
      return extractEventsWithMiniMax(email, model);
    default: {
      const _exhaustive: never = model;
      throw new Error(
        `Unsupported model: ${String(_exhaustive)}. Supported: ${SUPPORTED_MODELS.join(", ")}`,
      );
    }
  }
}
