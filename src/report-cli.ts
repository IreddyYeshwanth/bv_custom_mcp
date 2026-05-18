import "dotenv/config";
import { generateRrKeyMetricsTemplateReport } from "./rr-key-metrics-report.js";
import { generateSampleReportExcel } from "./rr-report.js";
import { generateInstanceAuditReport } from "./instance-audit-report.js";

// Usage:
// npm run report:rr:full -- alwaysdiscreet      -> R&R Key Metrics
// npm run report:rr:sample -- pampers-en-us     -> R&R Sample
// npm run report:audit -- alwaysdiscreet        -> Audit Report
// npm run report -- audit alwaysdiscreet        -> Dynamic mode + instance
// npm run report -- rr-full alwaysdiscreet      -> Dynamic mode + instance
// npm run report -- rr-sample pampers-en-us     -> Dynamic mode + instance

const npmScript = process.env.npm_lifecycle_event || "";
const modeArg = process.argv[2]?.toLowerCase();

type ReportMode = "rr-full" | "rr-sample" | "audit";

function detectMode(): ReportMode {
  if (modeArg === "rr-full" || modeArg === "rr-sample" || modeArg === "audit") {
    return modeArg;
  }

  if (npmScript.includes("audit")) {
    return "audit";
  }

  if (npmScript.includes("sample") || process.env.NODE_ENV === "sample") {
    return "rr-sample";
  }

  return "rr-full";
}

function detectInstanceName(): string {
  // When called as `npm run report -- <mode> <instance>`
  if (modeArg === "rr-full" || modeArg === "rr-sample" || modeArg === "audit") {
    return process.argv[3] || "pampers-en-us";
  }

  // When called via dedicated scripts, first arg is always the instance.
  return process.argv[2] || "pampers-en-us";
}

let report;
const mode = detectMode();
const instanceName = detectInstanceName();

if (mode === "audit") {
  report = await generateInstanceAuditReport({ instanceName });
} else if (mode === "rr-sample") {
  report = await generateSampleReportExcel(instanceName);
} else {
  report = await generateRrKeyMetricsTemplateReport(instanceName);
}

console.log(JSON.stringify(report, null, 2));