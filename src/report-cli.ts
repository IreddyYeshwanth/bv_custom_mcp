import "dotenv/config";
import { getConfiguredInstances } from "./bazaarvoice.js";
import { generateRrKeyMetricsTemplateReport } from "./rr-key-metrics-report.js";
import { generateSampleReportExcel } from "./rr-report.js";
import { generateInstanceAuditReport } from "./instance-audit-report.js";

// Usage examples:
//   npm run report -- rr-full   pampers-en-us
//   npm run report -- rr-full   pampers-en-us,alwaysdiscreet
//   npm run report -- rr-sample pampers-en-us
//   npm run report -- audit     pampers-en-us,alwaysdiscreet
//   npm run report -- audit     (no instance → all configured instances)
//
// Dedicated npm scripts (pass instance name(s) as first positional arg):
//   npm run report:rr:full   -- pampers-en-us
//   npm run report:rr:sample -- pampers-en-us
//   npm run report:audit     -- pampers-en-us,alwaysdiscreet

const npmScript = process.env.npm_lifecycle_event ?? "";
const modeArg   = process.argv[2]?.toLowerCase();

type ReportMode = "rr-full" | "rr-sample" | "audit";

function detectMode(): ReportMode {
  if (modeArg === "rr-full" || modeArg === "rr-sample" || modeArg === "audit") return modeArg;
  if (npmScript.includes("audit"))  return "audit";
  if (npmScript.includes("sample")) return "rr-sample";
  return "rr-full";
}

/**
 * Resolve instance name(s) from CLI args.
 * Supports comma-separated lists: "pampers-en-us,alwaysdiscreet"
 * Falls back to all configured instances when no argument is given.
 */
function detectInstanceNames(): string[] {
  // When called as `npm run report -- <mode> <instances>`
  const rawArg = (modeArg === "rr-full" || modeArg === "rr-sample" || modeArg === "audit")
    ? process.argv[3]
    : process.argv[2]; // dedicated script: first arg is the instance list

  if (rawArg?.trim()) {
    return rawArg.split(",").map((s) => s.trim()).filter(Boolean);
  }

  // No instance specified — default to all configured instances
  return getConfiguredInstances().map((i) => i.name);
}

const mode          = detectMode();
const instanceNames = detectInstanceNames();

console.error(`[report-cli] mode=${mode} instances=[${instanceNames.join(", ")}]`);

let report;

if (mode === "audit") {
  report = await generateInstanceAuditReport({ instanceNames });
} else if (mode === "rr-sample") {
  // Sample report is per-instance; collect all results
  const results = await Promise.all(instanceNames.map((n) => generateSampleReportExcel(n)));
  report = results.length === 1 ? results[0] : results;
} else {
  report = await generateRrKeyMetricsTemplateReport(instanceNames);
}

console.log(JSON.stringify(report, null, 2));
