import type { VercelRequest, VercelResponse } from "@vercel/node";
import {
  MANIFEST_TEMPLATE,
  MANIFEST_PLACEHOLDER,
} from "../lib/manifest-template";

export default function handler(_req: VercelRequest, res: VercelResponse) {
  const clientId = process.env.AZURE_CLIENT_ID?.trim();

  if (!clientId) {
    res.setHeader("Content-Type", "application/xml; charset=utf-8");
    res.status(503);
    return res.send(
      `<?xml version="1.0" encoding="UTF-8"?>\n<error>AZURE_CLIENT_ID is not configured on this deployment. Set it via \`vercel env add AZURE_CLIENT_ID production\` and redeploy.</error>\n`,
    );
  }

  const xml = MANIFEST_TEMPLATE.split(MANIFEST_PLACEHOLDER).join(clientId);

  res.setHeader("Content-Type", "application/xml; charset=utf-8");
  res.setHeader("Cache-Control", "no-store");
  res.status(200);
  res.send(xml);
}
