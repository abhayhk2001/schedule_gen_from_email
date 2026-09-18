import { readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, resolve } from "node:path";

const here = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(here, "..");

const templatePath = resolve(repoRoot, "lib/manifest-template.ts");
const outDir = resolve(repoRoot, "public/outlook-addin");
const outFile = resolve(outDir, "manifest.xml");
const placeholder = "__AZURE_CLIENT_ID__";

const clientId = (process.env.AZURE_CLIENT_ID ?? "").trim();

if (!clientId) {
  console.error(
    "[build-manifest] AZURE_CLIENT_ID is empty. Run `vercel env add AZURE_CLIENT_ID production --type config` (this value is a public client ID, not a secret).",
  );
  process.exit(1);
}

if (!/^[0-9a-fA-F-]{32,}$/.test(clientId)) {
  console.error(
    `[build-manifest] AZURE_CLIENT_ID does not look like a GUID: "${clientId}"`,
  );
  process.exit(1);
}

const source = readFileSync(templatePath, "utf8");

const literalMatch = source.match(/MANIFEST_TEMPLATE\s*=\s*`([\s\S]*?)`\s*;/);
if (!literalMatch) {
  console.error(
    "[build-manifest] could not locate MANIFEST_TEMPLATE template literal in",
    templatePath,
  );
  process.exit(1);
}

const template = literalMatch[1];
const xml = template.split(placeholder).join(clientId);

mkdirSync(outDir, { recursive: true });
writeFileSync(outFile, xml, "utf8");

console.log(
  `[build-manifest] wrote ${outFile} (${xml.length} bytes) with AZURE_CLIENT_ID substituted`,
);
