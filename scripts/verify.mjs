// "Basic working of the website" guardrail. Runs every check that should pass
// before a commit (or a deploy) is allowed through. Invoked by the pre-commit
// hook and by `npm run verify`.
import { execSync } from "node:child_process";

const DUMMY_GUID = "00000000-0000-0000-0000-000000000000";

function run(label, cmd, env) {
  console.log(`\n>>> ${label}`);
  execSync(cmd, { stdio: "inherit", env: { ...process.env, ...(env ?? {}) } });
}

run("typecheck", "npm run typecheck");
run("unit tests", "npm test");
run("add-in smoke", "npm run smoke");
run("site health (landing page)", "node scripts/check-site.mjs");

// Build script must not block deploys when the env var is missing — this is
// the 404 guardrail. If anyone reintroduces process.exit(1) on the empty path,
// this step fails and the commit is blocked.
run(
  "build script (valid GUID — must generate manifest)",
  "node scripts/build-manifest.mjs",
  { AZURE_CLIENT_ID: DUMMY_GUID },
);
run(
  "build script (empty GUID — must NOT block deploy)",
  "node scripts/build-manifest.mjs",
  { AZURE_CLIENT_ID: "" },
);

console.log("\nAll checks passed.");
