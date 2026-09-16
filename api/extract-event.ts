import type { VercelRequest, VercelResponse } from "@vercel/node";
import { extractEvents } from "../lib/minimax.js";

export default async function handler(req: VercelRequest, res: VercelResponse) {
  if (req.method !== "POST") {
    res.setHeader("Allow", "POST");
    return res.status(405).json({ error: "Method Not Allowed" });
  }

  const email =
    typeof req.body === "string" ? req.body : (req.body?.email as unknown);

  if (typeof email !== "string" || email.trim().length === 0) {
    return res.status(400).json({
      error: 'Request body must be JSON with a non-empty "email" string field.',
    });
  }

  try {
    const result = await extractEvents(email);
    return res.status(200).json(result);
  } catch (err) {
    const message = err instanceof Error ? err.message : "Unknown error";
    return res.status(500).json({ error: message });
  }
}
