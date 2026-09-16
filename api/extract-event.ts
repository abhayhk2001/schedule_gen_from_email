import type { VercelRequest, VercelResponse } from "@vercel/node";
import { extractEvents } from "../lib/router.js";
import { isSupportedModel } from "../lib/prompt.js";
import { SUPPORTED_MODELS } from "../lib/types.js";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  const body = (typeof req.body === "object" && req.body !== null
    ? req.body
    : {}) as Record<string, unknown>;

  const email = body.email;
  const model = body.model;

  if (typeof email !== "string" || email.trim().length === 0) {
    return res.status(400).json({
      error: 'Request body must be JSON with a non-empty "email" string field.',
    });
  }

  if (typeof model !== "string" || !isSupportedModel(model)) {
    return res.status(400).json({
      error: `Request body must include a "model" field. Supported values: ${SUPPORTED_MODELS.join(", ")}.`,
    });
  }

  try {
    const result = await extractEvents(email, model);
    return res.status(200).json(result);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return res.status(500).json({ error: message });
  }
}
