import { appendFile, mkdir } from "node:fs/promises";
import path from "node:path";

// ─── Types ─────────────────────────────────────────────────────────────────────

type LogLevel = "INFO" | "SUCCESS" | "WARN" | "ERROR";

type LogEntry = {
  timestamp: string;
  level: LogLevel;
  event: string;
  details?: Record<string, unknown>;
  error?: string;
};

// ─── Internal writer ──────────────────────────────────────────────────────────

const LOGS_DIR = path.resolve(process.cwd(), "logs");

async function writeEntry(entry: LogEntry): Promise<void> {
  const dateStamp = entry.timestamp.substring(0, 10).replace(/-/g, ""); // e.g. "20260819"
  const logFile   = path.join(LOGS_DIR, `run-${dateStamp}.log`);

  await mkdir(LOGS_DIR, { recursive: true });
  await appendFile(logFile, JSON.stringify(entry) + "\n", "utf8");
}

function toStderr(level: LogLevel, event: string, extra?: string): void {
  const prefix = `[${level.padEnd(7)}]`;
  const suffix = extra ? ` — ${extra}` : "";
  console.error(`${prefix} ${event}${suffix}`);
}

// ─── Public API ───────────────────────────────────────────────────────────────

/**
 * Log an informational event (e.g. report run started).
 */
export async function logInfo(
  event: string,
  details?: Record<string, unknown>,
): Promise<void> {
  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level: "INFO",
    event,
    ...(details && { details }),
  };
  toStderr("INFO", event);
  await writeEntry(entry);
}

/**
 * Log a successful operation (e.g. report generated, instance loaded).
 */
export async function logSuccess(
  event: string,
  details?: Record<string, unknown>,
): Promise<void> {
  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level: "SUCCESS",
    event,
    ...(details && { details }),
  };
  toStderr("SUCCESS", event);
  await writeEntry(entry);
}

/**
 * Log a non-fatal warning (e.g. one instance skipped, partial data).
 */
export async function logWarn(
  event: string,
  details?: Record<string, unknown>,
): Promise<void> {
  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level: "WARN",
    event,
    ...(details && { details }),
  };
  toStderr("WARN", event);
  await writeEntry(entry);
}

/**
 * Log a fatal or caught error. Pass the original Error or a string message.
 */
export async function logError(
  event: string,
  error: unknown,
  details?: Record<string, unknown>,
): Promise<void> {
  const message = error instanceof Error ? error.message : String(error);
  const entry: LogEntry = {
    timestamp: new Date().toISOString(),
    level: "ERROR",
    event,
    error: message,
    ...(details && { details }),
  };
  toStderr("ERROR", event, message);
  await writeEntry(entry);
}
