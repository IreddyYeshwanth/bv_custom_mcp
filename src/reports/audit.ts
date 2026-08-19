import { mkdir } from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import { findInstanceByName, listProducts } from "../api/bazaarvoice.js";
import { isProductActive, resolveProductStatus } from "../utils/product-status.js";
import type { BvInstance, Product } from "../types/index.js";

const DEFAULT_REPORTS_DIR = path.resolve(process.cwd(), "reports");

const DOCS_LINK =
  "https://pgone.sharepoint.com/:p:/s/GlobalProductServices/IQAhg8ackh-HQ77DjbpXzz43ASkqCZKDUUbpZ9z4fYTL8-M?e=ELbrVw";

// ─── Colour palette ───────────────────────────────────────────────────────────

const C = {
  DARK_BLUE_TEXT: "1F4E79",
  LIGHT_BLUE_BG:  "D6E6F4",
  HEADER_BLUE_BG: "4472C4",
  WHITE:          "FFFFFF",
  CRUCIAL:        "FF6B6B",
  HIGH:           "FFA500",
  MEDIUM:         "FFD93D",
  LOW:            "6BCB77",
  TOTAL_ROW_BG:   "E2EFDA",
  NAVY_BG:        "002060",
  ALERT_LABEL:    "FFE699",
  ALERT_VALUE:    "FFF9E6",
  ALERT_BANNER:   "C00000",
} as const;

/** Maximum columns across any section table. Section 2 = 14 cols. */
const MAX_COLS = 14;

// ─── Types ────────────────────────────────────────────────────────────────────

export type AuditReportOptions = {
  /** Single instance name — backward-compatible shorthand. */
  instanceName?: string;
  /** One or more instance names to combine into a single report. */
  instanceNames?: string[];
  outputDir?: string;
};

export type AuditReportResult = {
  instanceNames: string[];
  outputPath: string;
  totalProducts: number;
  generatedAt: string;
};

type ProductMetrics = {
  instanceName: string;
  displayName: string;
  product: Product;
  isActive: boolean;
  reviewCount: number;
  hasDescription: boolean;
  hasImageUrl: boolean;
  hasPageUrl: boolean;
  hasUpc: boolean;
  hasEan: boolean;
  hasName: boolean;
  hasBrand: boolean;
};

type InstanceSummary = {
  instanceName: string;
  displayName: string;
  isNorthAmerica: boolean;
  totalProducts: number;
  activeProducts: number;
  inactiveProducts: number;
  percentInactive: number;
  totalReviews: number;
  /** Reviews associated with inactive products (currently hidden). */
  reviewsHiddenOnInactive: number;
  percentReviewsHidden: number;
  /** NA: missing UPC count. Others: missing EAN count. */
  missingUpcOrEan: number;
  missingDescription: number;
  missingImageUrl: number;
  missingName: number;
  missingPageUrl: number;
  missingBrand: number;
};

type InstanceAuditData = {
  instanceName: string;
  displayName: string;
  isNorthAmerica: boolean;
  metrics: ProductMetrics[];
  summary: InstanceSummary;
};

// ─── Region helpers ───────────────────────────────────────────────────────────

const NA_SUBSTRINGS = ["-us", "-ca", "_us", "_ca", "en-us", "en-ca", "fr-ca"];
const EU_SUBSTRINGS = [
  "-uk", "-gb", "-de", "-fr", "-be", "-nl", "-it", "-es-es", "-pt",
  "-at", "-ch", "-se", "-no", "-dk", "-fi", "-ie", "-pl", "-cz",
  "-sk", "-hu", "-ro", "-bg", "-hr", "-si", "-ee", "-lv", "-lt",
];

function heuristicIsNA(name: string): boolean {
  return NA_SUBSTRINGS.some((s) => name.toLowerCase().includes(s));
}
function heuristicIsEU(name: string): boolean {
  return EU_SUBSTRINGS.some((s) => name.toLowerCase().includes(s));
}

function resolveRegion(instance: BvInstance): { order: number; isNA: boolean } {
  if (instance.region === "na")    return { order: 0, isNA: true };
  if (instance.region === "eu")    return { order: 1, isNA: false };
  if (instance.region === "other") return { order: 2, isNA: false };
  if (heuristicIsNA(instance.name)) return { order: 0, isNA: true };
  if (heuristicIsEU(instance.name)) return { order: 1, isNA: false };
  return { order: 2, isNA: false };
}

function sortByRegion<T extends { instanceName: string; isNorthAmerica: boolean }>(
  items: T[],
  regionOrders: Map<string, number>,
): T[] {
  return [...items].sort((a, b) => {
    const diff = (regionOrders.get(a.instanceName) ?? 2) - (regionOrders.get(b.instanceName) ?? 2);
    return diff !== 0 ? diff : a.instanceName.localeCompare(b.instanceName);
  });
}

function resolveDisplayName(instance: BvInstance): string {
  if (instance.displayName) return instance.displayName;
  return instance.name
    .split(/[-_]/)
    .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
    .join(" ");
}

// ─── Utility helpers ──────────────────────────────────────────────────────────

function round(value: number, decimals = 4): number {
  return Math.round(value * 10 ** decimals) / 10 ** decimals;
}

/** Returns a 0–1 fraction (raw), matching benchmark storage format. */
function fracPct(part: number, total: number): number {
  return total > 0 ? part / total : 0;
}

function formatTimestampForFilename(date: Date): string {
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function formatRunDateShort(date: Date): string {
  return date.toISOString().replace("T", " ").substring(0, 10) + " UTC";
}

function formatRunDate(date: Date): string {
  const month   = date.toLocaleString("en-US", { month: "short", timeZone: "UTC" });
  const day     = String(date.getUTCDate()).padStart(2, "0");
  const year    = date.getUTCFullYear();
  const hours   = String(date.getUTCHours() % 12 || 12).padStart(2, "0");
  const minutes = String(date.getUTCMinutes()).padStart(2, "0");
  const ampm    = date.getUTCHours() >= 12 ? "PM" : "AM";
  return `Run Date: ${month} ${day}, ${year} ${hours}:${minutes} ${ampm} UTC`;
}

function safeCellValue(value: unknown): string | number {
  if (value === null || value === undefined) return "";
  if (typeof value === "number") return value;
  const s = String(value);
  return s.length > 32700 ? `${s.substring(0, 32700)}...` : s;
}

function priorityColor(level: string): string {
  switch (level.toUpperCase()) {
    case "CRUCIAL": return C.CRUCIAL;
    case "HIGH":    return C.HIGH;
    case "MEDIUM":  return C.MEDIUM;
    case "LOW":     return C.LOW;
    default:        return C.WHITE;
  }
}

// ─── Cell-level helpers ───────────────────────────────────────────────────────

interface CellStyle {
  bold?:      boolean;
  italic?:    boolean;
  fontColor?: string;
  bgColor?:   string;
  fontSize?:  number;
  wrapText?:  boolean;
}

function sc(
  ws: xlsx.WorkSheet,
  r: number,
  c: number,
  value: string | number | null,
  style?: CellStyle,
): void {
  const addr = xlsx.utils.encode_cell({ r, c });
  const v    = value ?? "";
  ws[addr]   = { v, t: typeof v === "number" ? "n" : "s" };

  if (style?.bgColor) ws[addr].fill = { patternType: "solid", fgColor: { rgb: style.bgColor } };

  const font: Record<string, unknown> = {};
  if (style?.bold)      font.bold   = true;
  if (style?.italic)    font.italic = true;
  if (style?.fontColor) font.color  = { rgb: style.fontColor };
  if (style?.fontSize)  font.sz     = style.fontSize;
  if (Object.keys(font).length) ws[addr].font = font;

  const align: Record<string, unknown> = { vertical: "top" };
  if (style?.wrapText) align.wrapText = true;
  ws[addr].alignment = align;
}

/** Write a section-heading row spanning MAX_COLS. Returns next row index. */
function sectionHeading(
  ws: xlsx.WorkSheet,
  merges: xlsx.Range[],
  row: number,
  text: string,
): number {
  sc(ws, row, 0, text, { bold: true, fontSize: 12, fontColor: C.DARK_BLUE_TEXT, bgColor: C.LIGHT_BLUE_BG });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.LIGHT_BLUE_BG });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  return row + 1;
}

/** Write a table header row. Returns next row index. */
function tHeaders(ws: xlsx.WorkSheet, row: number, headers: string[]): number {
  for (const [c, h] of headers.entries()) {
    sc(ws, row, c, h, { bold: true, fontColor: C.WHITE, bgColor: C.HEADER_BLUE_BG, wrapText: true });
  }
  return row + 1;
}

/** Write a standard data row. Returns next row index. */
function dRow(
  ws: xlsx.WorkSheet,
  row: number,
  values: (string | number | null)[],
  style?: CellStyle,
): number {
  for (const [c, val] of values.entries()) {
    sc(ws, row, c, val, style ?? { wrapText: true });
  }
  return row + 1;
}

/** Write a row whose background is determined by the value at `priorityCol`. Returns next row. */
function pRow(
  ws: xlsx.WorkSheet,
  row: number,
  values: (string | number | null)[],
  priorityCol: number,
): number {
  const bg = priorityColor(String(values[priorityCol] ?? ""));
  return dRow(ws, row, values, { bgColor: bg, wrapText: true });
}

/** Write a bold TOTAL row. Returns next row. */
function totRow(ws: xlsx.WorkSheet, row: number, values: (string | number | null)[]): number {
  return dRow(ws, row, values, { bold: true, bgColor: C.TOTAL_ROW_BG, wrapText: true });
}

// ─── Data fetching ────────────────────────────────────────────────────────────

async function listAllProducts(instanceName: string): Promise<Product[]> {
  const all: Product[] = [];
  const pageSize = 100;
  for (let offset = 0; ; offset += pageSize) {
    const page = await listProducts(instanceName, { limit: pageSize, offset, includeStats: true });
    all.push(...page.Results);
    if (offset + page.Results.length >= page.TotalResults || page.Results.length < pageSize) break;
  }
  return all;
}

function buildMetrics(instanceName: string, displayName: string, products: Product[]): ProductMetrics[] {
  return products.map((product) => ({
    instanceName,
    displayName,
    product,
    isActive:       isProductActive(product),
    reviewCount:    product.ReviewStatistics?.TotalReviewCount ?? 0,
    hasDescription: Boolean(product.Description?.trim()),
    hasImageUrl:    Boolean(product.ImageUrl?.trim()),
    hasPageUrl:     Boolean(product.ProductPageUrl?.trim()),
    hasUpc:         Boolean(product.UPCs?.length),
    hasEan:         Boolean(product.EANs?.length),
    hasName:        Boolean(product.Name?.trim()),
    hasBrand:       Boolean(product.Brand?.Name?.trim()),
  }));
}

function buildSummary(
  instanceName: string,
  displayName: string,
  isNA: boolean,
  metrics: ProductMetrics[],
): InstanceSummary {
  const total    = metrics.length;
  const active   = metrics.filter((m) => m.isActive).length;
  const inactive = total - active;
  const totalRev = metrics.reduce((s, m) => s + m.reviewCount, 0);
  const hiddenRev = metrics.filter((m) => !m.isActive).reduce((s, m) => s + m.reviewCount, 0);

  return {
    instanceName,
    displayName,
    isNorthAmerica:          isNA,
    totalProducts:           total,
    activeProducts:          active,
    inactiveProducts:        inactive,
    percentInactive:         fracPct(inactive, total),
    totalReviews:            totalRev,
    reviewsHiddenOnInactive: hiddenRev,
    percentReviewsHidden:    fracPct(hiddenRev, totalRev),
    missingUpcOrEan:         isNA ? metrics.filter((m) => !m.hasUpc).length : metrics.filter((m) => !m.hasEan).length,
    missingDescription:      metrics.filter((m) => !m.hasDescription).length,
    missingImageUrl:         metrics.filter((m) => !m.hasImageUrl).length,
    missingName:             metrics.filter((m) => !m.hasName).length,
    missingPageUrl:          metrics.filter((m) => !m.hasPageUrl).length,
    missingBrand:            metrics.filter((m) => !m.hasBrand).length,
  };
}

// ─── Product Data sheet (per-instance) ───────────────────────────────────────

function createProductDataSheet(
  instanceName: string,
  displayName: string,
  metrics: ProductMetrics[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  sc(ws, row++, 0, "Bazaarvoice");
  sc(ws, row++, 0, `R&R Key Metrics ${displayName}`);
  sc(ws, row++, 0, formatRunDate(new Date(generatedAt)));
  sc(ws, row++, 0, `Instance: ${instanceName}`);
  sc(ws, row++, 0, "Date range: All Time");
  sc(ws, row++, 0, "Date field: Review Submission Date");
  row++;

  const headers = [
    "Brand ID", "Brand", "Product ID", "Product Status", "Product Name",
    "Description", "Product Page URL", "Image URL", "UPC", "EAN",
    "Product Family", "Product Family - Expand", "# Approved Reviews",
    "# Family Reviews", "Family Average Rating",
  ];
  row = tHeaders(ws, row, headers);

  for (const m of metrics) {
    const p   = m.product;
    const avg = p.ReviewStatistics?.AverageOverallRating;
    row = dRow(ws, row, [
      safeCellValue(p.BrandExternalId ?? p.Brand?.Id ?? ""),
      safeCellValue(p.Brand?.Name ?? ""),
      safeCellValue(p.Id),
      safeCellValue(resolveProductStatus(p)),
      safeCellValue(p.Name),
      safeCellValue(p.Description ?? ""),
      safeCellValue(p.ProductPageUrl ?? ""),
      safeCellValue(p.ImageUrl ?? ""),
      safeCellValue(p.UPCs?.[0] ?? ""),
      safeCellValue(p.EANs?.[0] ?? ""),
      safeCellValue(p.FamilyIds?.[0] ?? ""),
      "",
      m.reviewCount,
      m.reviewCount,
      typeof avg === "number" ? round(avg, 2) : "",
    ]);
  }

  ws["!ref"]  = xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row - 1, c: 14 } });
  ws["!cols"] = [15, 12, 18, 12, 30, 50, 30, 30, 15, 15, 15, 20, 18, 18, 18].map((wch) => ({ wch }));
  return ws;
}

// ─── Main Audit Report sheet ──────────────────────────────────────────────────

function createAuditReportSheet(
  allData: InstanceAuditData[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  const merges: xlsx.Range[] = [];
  let row = 0;

  // ── Title block ─────────────────────────────────────────────────────────────
  sc(ws, row, 0, "Bazaarvoice Product Catalog Audit Report", {
    bold: true, fontSize: 14, fontColor: C.WHITE, bgColor: C.NAVY_BG,
  });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.NAVY_BG });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row++;

  sc(ws, row, 0, `All regional instances | All-time catalog snapshot | Run date in source: ${formatRunDateShort(new Date(generatedAt))}`);
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "");
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row++;

  sc(ws, row, 0, `Documentation: ${DOCS_LINK}`);
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "");
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row += 2;

  // ── Aggregate totals ─────────────────────────────────────────────────────────
  const allInactiveWithReviews = allData.flatMap((d) => d.metrics.filter((m) => !m.isActive && m.reviewCount > 0));
  const totalHiddenReviews     = allInactiveWithReviews.reduce((s, m) => s + m.reviewCount, 0);
  const totalReviewsGlobal     = allData.reduce((s, d) => s + d.summary.totalReviews, 0);

  // ── TOP PRIORITY ALERT ───────────────────────────────────────────────────────
  sc(ws, row, 0, "TOP PRIORITY ALERT", { bold: true, fontSize: 12, fontColor: C.WHITE, bgColor: C.ALERT_BANNER });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.ALERT_BANNER });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row++;

  const alertRows: [string, string | number][] = [
    ["Most Critical Issue", "Reviews are currently hidden because they are associated with inactive products."],
    ["Total Impact", totalHiddenReviews],
    ["Number of Affected Products", allInactiveWithReviews.length],
    ["Number of Reviews Hidden/Tagged to Inactive Products", totalHiddenReviews],
    ["% Reviews Hidden/Tagged to Inactive Products", fracPct(totalHiddenReviews, totalReviewsGlobal)],
    ["Required Action Summary", "Prioritize reviewed inactive products."],
    ["Recommended Fix", "Map IDs to live PDPs; then activate."],
    ["Suggestion Impact Level", "CRUCIAL"],
  ];
  for (const [label, value] of alertRows) {
    sc(ws, row, 0, label, { bold: true, bgColor: C.ALERT_LABEL });
    sc(ws, row, 1, value, { wrapText: true, bgColor: C.ALERT_VALUE });
    for (let c = 2; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.ALERT_VALUE });
    row++;
  }
  row++;

  // ── SECTION 1: CRITICAL FIELD DEFINITIONS ───────────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 1: CRITICAL FIELD DEFINITIONS FOR AUDIT");

  row = tHeaders(ws, row, [
    "Critical Field", "Column Reference", "Why Critical",
    "Business Impact if Missing", "Recommended Action", "Documentation Reference",
  ]);

  const fieldDefs: string[][] = [
    ["Product ID",              "C", "Primary catalog matching key.",      "Matching and governance fail.",             "Use one stable canonical ID.",       DOCS_LINK],
    ["UPC/EAN",                 "I / J", "Regional syndication identifier.", "Syndication can be blocked.",           "Populate UPC in NA; EAN elsewhere.", DOCS_LINK],
    ["Product Name",            "E", "Customer-facing product identity.",  "PDP clarity is reduced.",                  "Supply the approved local name.",     DOCS_LINK],
    ["Product Description",     "F", "Core product content.",              "PDP completeness is reduced.",             "Supply approved local copy.",         DOCS_LINK],
    ["Product Page URL / PDP URL", "G", "Links reviews to the live PDP.", "Reviews may not render on PDP.",           "Use the canonical live PDP URL.",     DOCS_LINK],
    ["Image URL",               "H", "Provides product imagery.",          "Catalog presentation is weakened.",        "Use a reachable canonical image.",    DOCS_LINK],
    ["Brand",                   "B", "Supports attribution and reporting.", "Brand reporting can fragment.",           "Use approved brand taxonomy.",        DOCS_LINK],
    ["Product Status",          "D", "Controls product visibility.",        "Inactive items hide reviews.",            "Validate Active/Inactive status.",    DOCS_LINK],
  ];
  for (const fd of fieldDefs) row = dRow(ws, row, fd);
  row++;

  // ── SECTION 2: VOLUME-BASED SUMMARY BY INSTANCE ─────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 2: VOLUME-BASED SUMMARY BY INSTANCE");

  row = tHeaders(ws, row, [
    "Instance / Region", "Total Products", "Active Products", "Inactive Products",
    "% Inactive", "Total Reviews", "Reviews Hidden/Tagged to Inactive Product",
    "% Reviews Hidden/Tagged to Inactive Product", "Missing UPC/EAN",
    "Missing Description", "Missing Image URL",
    "Business Impact", "Recommended Action", "Suggestion Impact",
  ]);

  let totProd = 0, totAct = 0, totInact = 0, totRev = 0, totHid = 0;
  let totMisUE = 0, totMisDesc = 0, totMisImg = 0;

  for (const d of allData) {
    const s = d.summary;
    totProd    += s.totalProducts;
    totAct     += s.activeProducts;
    totInact   += s.inactiveProducts;
    totRev     += s.totalReviews;
    totHid     += s.reviewsHiddenOnInactive;
    totMisUE   += s.missingUpcOrEan;
    totMisDesc += s.missingDescription;
    totMisImg  += s.missingImageUrl;

    const imp = s.reviewsHiddenOnInactive > 0 ? "CRUCIAL" : s.missingUpcOrEan > 0 ? "HIGH" : "MEDIUM";
    row = pRow(ws, row, [
      d.displayName,
      s.totalProducts, s.activeProducts, s.inactiveProducts,
      s.percentInactive, s.totalReviews, s.reviewsHiddenOnInactive,
      s.percentReviewsHidden, s.missingUpcOrEan, s.missingDescription, s.missingImageUrl,
      "Hidden reviews reduce social proof.",
      "Map reviewed IDs; then activate.",
      imp,
    ], 13);
  }

  row = totRow(ws, row, [
    "TOTAL", totProd, totAct, totInact,
    fracPct(totInact, totProd),
    totRev, totHid,
    fracPct(totHid, totRev),
    totMisUE, totMisDesc, totMisImg,
    `Hidden reviews: ${totHid.toLocaleString()}`,
    `Prioritize ${totHid.toLocaleString()} hidden reviews.`,
    totHid > 0 ? "CRUCIAL" : "MEDIUM",
  ]);
  row++;

  // Immediate attention note
  const naInstances = allData.filter((d) => d.isNorthAmerica).map((d) => d.displayName);
  sc(ws, row, 0,
    "Immediate attention: prioritize instances with the largest hidden-review volume; " +
    `North American instances are listed first.${naInstances.length === 0 ? " No North American instance is present in this workbook." : ""}`,
    { italic: true, fontColor: C.DARK_BLUE_TEXT, wrapText: true },
  );
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "");
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row += 2;

  // ── SECTION 3: ISSUES WITH BUSINESS IMPACT ───────────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 3: ISSUES WITH BUSINESS IMPACT, RECOMMENDATIONS & HOW TO FIX");

  row = tHeaders(ws, row, [
    "Instance / Region", "#", "Issue Category", "Critical Field",
    "Volume Impacted", "Business Impact", "Recommended Action", "How to Fix", "Suggestion Impact",
  ]);

  for (const d of allData) {
    const s    = d.summary;
    const disp = d.displayName;
    const field = s.isNorthAmerica ? "UPC" : "EAN";
    const col   = s.isNorthAmerica ? "I" : "J";

    const inactiveWithReviews = d.metrics.filter((m) => !m.isActive && m.reviewCount > 0);
    const hiddenCount         = inactiveWithReviews.reduce((sum, m) => sum + m.reviewCount, 0);

    const numericIds      = d.metrics.filter((m) => /^\d+$/.test(m.product.Id));
    const alphanumericIds = d.metrics.filter((m) => !/^\d+$/.test(m.product.Id));

    // Unique products with at least one missing crucial field
    const missingFieldCount = d.metrics.filter(
      (m) => !m.hasDescription || !m.hasImageUrl || !m.hasPageUrl || !m.hasName || !m.hasBrand,
    ).length;

    // Brand variation count
    const uniqueBrandNames = new Set(
      d.metrics.map((m) => m.product.Brand?.Name?.trim()).filter(Boolean),
    );
    const brandVarCount = uniqueBrandNames.size;

    let n = 1;

    // 1 — Inactive → hidden reviews (CRUCIAL)
    row = pRow(ws, row, [
      disp, n++,
      "Inactive Products → Reviews Hidden/Tagged to Inactive Product", "Product Status",
      `${d.metrics.filter((m) => !m.isActive).length} products / ${hiddenCount} reviews`,
      "Hidden reviews reduce social proof.",
      "Reactivate or migrate reviewed products.",
      "Map IDs to live PDPs; then activate.",
      "CRUCIAL",
    ], 8);

    // 2 — Product ID issues (HIGH)
    row = pRow(ws, row, [
      disp, n++,
      "Product ID Issues", "Product ID",
      `Numeric ${numericIds.length}; alphanumeric ${alphanumericIds.length}; blank 0`,
      "Mixed IDs hinder catalog matching.",
      "Adopt one canonical ID standard.",
      "Map legacy IDs; publish canonical IDs.",
      "HIGH",
    ], 8);

    // 3 — UPC/EAN missing (HIGH)
    row = pRow(ws, row, [
      disp, n++,
      `${field} Data Missing Issues`, field,
      `${s.missingUpcOrEan} products`,
      "Missing codes can block syndication.",
      `Populate approved ${field} values.`,
      `Validate source codes; update column ${col}.`,
      "HIGH",
    ], 8);

    // 4 — Missing crucial fields (MEDIUM)
    row = pRow(ws, row, [
      disp, n++,
      "Missing Crucial Fields", "Name/PDP/Description/Image/Brand",
      `${missingFieldCount} products`,
      "Missing content weakens PDP quality.",
      "Complete required catalog fields.",
      "Backfill from approved product masters.",
      "MEDIUM",
    ], 8);

    // 5 — Brand name issues (MEDIUM) — only if >1 variation
    if (brandVarCount > 1) {
      row = pRow(ws, row, [
        disp, n++,
        "Brand Name Issues", "Brand",
        `${brandVarCount} variations`,
        "Brand variants fragment reporting.",
        "Use one approved brand label.",
        "Normalize case, spacing, and aliases.",
        "MEDIUM",
      ], 8);
    }
  }
  row++;

  // ── SECTION 4: INACTIVE PRODUCTS WITH REVIEWS – MIGRATION CANDIDATES ─────────
  row = sectionHeading(ws, merges, row, "SECTION 4: INACTIVE PRODUCTS WITH REVIEWS – MIGRATION CANDIDATES");

  row = tHeaders(ws, row, [
    "Instance / Region", "Product ID", "Product Name", "Brand", "Product Status",
    "Total Reviews Hidden/Tagged to Inactive Product", "Reviews on Product",
    "Suggested Migration Target / Action", "Business Impact", "How to Fix", "Suggestion Impact",
  ]);

  const migrationCandidates = allData
    .flatMap((d) => d.metrics.filter((m) => !m.isActive && m.reviewCount > 0))
    .sort((a, b) => b.reviewCount - a.reviewCount);

  if (migrationCandidates.length > 0) {
    for (const m of migrationCandidates) {
      const p = m.product;
      row = dRow(ws, row, [
        m.displayName,
        safeCellValue(p.Id),
        safeCellValue(p.Name) || 0,
        safeCellValue(p.Brand?.Name ?? "") || 0,
        resolveProductStatus(p),
        m.reviewCount,
        m.reviewCount,
        "Find active successor; otherwise activate.",
        "Reviews are currently hidden because they are associated with inactive products.",
        "Map IDs to a live PDP; migrate or activate.",
        "CRUCIAL",
      ]);
    }
  } else {
    row = dRow(ws, row, ["No inactive products with reviews found.", ...Array(10).fill("")]);
  }
  row++;

  // ── SECTION 5: PRODUCT ID STANDARDISATION ISSUES ──────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 5: PRODUCT ID STANDARDIZATION ISSUES");

  row = tHeaders(ws, row, [
    "Instance / Region", "Issue Type", "Examples", "Products Affected",
    "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ]);

  for (const d of allData) {
    const numericIds      = d.metrics.filter((m) => /^\d+$/.test(m.product.Id));
    const alphanumericIds = d.metrics.filter((m) => !/^\d+$/.test(m.product.Id));

    if (numericIds.length > 0) {
      row = pRow(ws, row, [
        d.displayName, "Numeric-only Product IDs",
        numericIds.slice(0, 3).map((m) => m.product.Id).join(", "),
        numericIds.length,
        "Mixed formats hinder matching.",
        "Adopt one canonical ID pattern.",
        "Map numeric IDs to canonical IDs.",
        "HIGH",
      ], 7);
    }

    if (alphanumericIds.length > 0) {
      row = pRow(ws, row, [
        d.displayName, "Alphanumeric/mixed Product IDs",
        alphanumericIds.slice(0, 3).map((m) => m.product.Id).join(", "),
        alphanumericIds.length,
        "Mixed formats hinder matching.",
        "Adopt one canonical ID pattern.",
        "Normalize separators and casing.",
        "HIGH",
      ], 7);
    }
  }
  row++;

  // ── SECTION 6: UPC/EAN DATA QUALITY ANALYSIS ──────────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 6: UPC/EAN DATA QUALITY ANALYSIS");

  row = tHeaders(ws, row, [
    "Instance / Region", "Issue Type", "Products Affected", "Examples",
    "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ]);

  let sec6HasData = false;
  for (const d of allData) {
    const s = d.summary;
    if (s.missingUpcOrEan === 0) continue;
    sec6HasData = true;
    const field    = s.isNorthAmerica ? "UPC" : "EAN";
    const col      = s.isNorthAmerica ? "I" : "J";
    const affected = s.isNorthAmerica
      ? d.metrics.filter((m) => !m.hasUpc)
      : d.metrics.filter((m) => !m.hasEan);

    row = pRow(ws, row, [
      d.displayName, `Missing ${field}`,
      s.missingUpcOrEan,
      affected.slice(0, 3).map((m) => m.product.Id).join(", "),
      "Missing codes can block syndication.",
      `Populate approved ${field} values.`,
      `Validate source; update column ${col}.`,
      "HIGH",
    ], 7);
  }
  if (!sec6HasData) {
    row = dRow(ws, row, ["No UPC/EAN data quality issues found.", ...Array(7).fill("")]);
  }
  row++;

  sc(ws, row, 0,
    "Assessment rule: multiple UPC/EAN values are compliant and are not used as an audit criterion.",
    { italic: true, fontColor: C.DARK_BLUE_TEXT },
  );
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "");
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row += 2;

  // ── SECTION 7: MISSING CRUCIAL FIELDS ─────────────────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 7: MISSING CRUCIAL FIELDS");

  row = tHeaders(ws, row, [
    "Instance / Region", "Missing Image URLs", "Missing Product Names",
    "Missing PDP URLs", "Missing Descriptions", "Missing Brand",
    "Total Missing Fields", "Business Impact", "How to Fix", "Suggestion Impact",
  ]);

  let s7Img = 0, s7Name = 0, s7Url = 0, s7Desc = 0, s7Brand = 0;
  for (const d of allData) {
    const s   = d.summary;
    s7Img    += s.missingImageUrl;
    s7Name   += s.missingName;
    s7Url    += s.missingPageUrl;
    s7Desc   += s.missingDescription;
    s7Brand  += s.missingBrand;
    const tot = s.missingImageUrl + s.missingName + s.missingPageUrl + s.missingDescription + s.missingBrand;
    const imp = tot > 50 ? "HIGH" : tot > 10 ? "MEDIUM" : "LOW";

    row = pRow(ws, row, [
      d.displayName,
      s.missingImageUrl, s.missingName, s.missingPageUrl,
      s.missingDescription, s.missingBrand, tot,
      tot > 0 ? "Missing fields weaken PDP quality." : "No missing fields",
      tot > 0 ? "Backfill from approved product masters." : "No remediation required.",
      imp,
    ], 9);
  }

  const totMissingFields = s7Img + s7Name + s7Url + s7Desc + s7Brand;
  row = totRow(ws, row, [
    "TOTAL", s7Img, s7Name, s7Url, s7Desc, s7Brand, totMissingFields,
    `Missing field cells: ${totMissingFields}`,
    "Complete all required fields.",
    "MEDIUM",
  ]);
  row++;

  // ── SECTION 8: BRAND CONSISTENCY CHECK ────────────────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 8: BRAND CONSISTENCY CHECK");

  row = tHeaders(ws, row, [
    "Instance / Region", "Brand Variations Found", "Count",
    "Issue", "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ]);

  for (const d of allData) {
    // Count occurrences of each unique brand name
    const brandCountMap = new Map<string, number>();
    for (const m of d.metrics) {
      const b = m.product.Brand?.Name?.trim();
      if (b) brandCountMap.set(b, (brandCountMap.get(b) ?? 0) + 1);
    }
    const uniqueBrands = [...brandCountMap.entries()].sort((a, b) => b[1] - a[1]);
    const brandSummary = uniqueBrands.map(([name, cnt]) => `${name} (${cnt})`).join(", ") || "—";
    const varCount     = uniqueBrands.length;
    const hasVariation = varCount > 1;

    row = pRow(ws, row, [
      d.displayName, brandSummary, varCount,
      hasVariation
        ? "Capitalization or regional naming differs."
        : "No within-instance variation found.",
      hasVariation
        ? "Brand variants fragment reporting."
        : "No variation impact found.",
      hasVariation
        ? "Use one approved regional label."
        : "Maintain current naming.",
      hasVariation
        ? "Normalize case, spaces, and aliases."
        : "No remediation required.",
      hasVariation ? "MEDIUM" : "LOW",
    ], 7);
  }
  row++;

  // ── SECTION 9: PRIORITY ACTION ITEMS BY INSTANCE ──────────────────────────────
  row = sectionHeading(ws, merges, row, "SECTION 9: PRIORITY ACTION ITEMS BY INSTANCE");

  row = tHeaders(ws, row, [
    "Instance / Region", "Priority Rank", "Issue",
    "Impact Level", "Action Required", "How to Fix", "Expected Benefit",
  ]);

  for (const d of allData) {
    const s     = d.summary;
    const field  = s.isNorthAmerica ? "UPC" : "EAN";
    const col    = s.isNorthAmerica ? "I" : "J";
    const uniqueBrands = new Set(d.metrics.map((m) => m.product.Brand?.Name?.trim()).filter(Boolean));
    const hasBrandVar  = uniqueBrands.size > 1;
    let rank = 1;

    row = pRow(ws, row, [
      d.displayName, rank++,
      "Resolve inactive products with reviews",
      "CRUCIAL",
      "Migrate reviews or activate products.",
      "Map IDs to live PDPs; then activate.",
      "Restores visible review content.",
    ], 3);

    row = pRow(ws, row, [
      d.displayName, rank++,
      "Standardize Product IDs",
      "HIGH",
      "Adopt a canonical Product ID pattern.",
      "Map legacy IDs; normalize future feeds.",
      "Improves matching and governance.",
    ], 3);

    row = pRow(ws, row, [
      d.displayName, rank++,
      `Complete missing ${field} values`,
      "HIGH",
      `Populate approved ${field} values.`,
      `Validate source; update column ${col}.`,
      "Improves syndication readiness.",
    ], 3);

    row = pRow(ws, row, [
      d.displayName, rank++,
      "Complete missing crucial fields",
      "MEDIUM",
      "Backfill required catalog content.",
      "Use approved product master content.",
      "Improves PDP completeness.",
    ], 3);

    if (hasBrandVar) {
      row = pRow(ws, row, [
        d.displayName, rank++,
        "Normalize brand names",
        "MEDIUM",
        "Use one approved regional label.",
        "Normalize case, spaces, and aliases.",
        "Unifies reporting and attribution.",
      ], 3);
    }
  }

  // ── Finalise sheet ────────────────────────────────────────────────────────────
  ws["!ref"]    = xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row - 1, c: MAX_COLS - 1 } });
  ws["!merges"] = merges;
  ws["!cols"]   = [22, 13, 30, 22, 16, 36, 40, 45, 16, 36, 36, 36, 25, 16].map((wch) => ({ wch }));

  return ws;
}

// ─── Public API ───────────────────────────────────────────────────────────────

function resolveInstanceNames(options: AuditReportOptions): string[] {
  if (options.instanceNames?.length) return [...new Set(options.instanceNames)];
  if (options.instanceName)          return [options.instanceName];
  throw new Error("Provide instanceName or instanceNames in AuditReportOptions.");
}

/**
 * Generate a multi-instance audit workbook.
 *
 * Workbook layout:
 *   Sheet 1: "Audit Report" — comprehensive 9-section analysis across all instances.
 *   Sheet N: per-instance product data sheets, named after the instance.
 *
 * Accepts a single instanceName (backward-compatible) or an array via instanceNames.
 * Instances are sorted NA → EU → Other, then alphabetically within each region.
 */
export async function generateInstanceAuditReport(
  options: AuditReportOptions = {},
): Promise<AuditReportResult> {
  const outputDir   = options.outputDir ?? DEFAULT_REPORTS_DIR;
  const generatedAt = new Date().toISOString();
  const names       = resolveInstanceNames(options);

  // Fetch data for all instances — continue even if individual instances fail
  const regionOrders = new Map<string, number>();

  const settled = await Promise.allSettled(
    names.map(async (instanceName) => {
      const instance   = findInstanceByName(instanceName);
      const { order, isNA } = resolveRegion(instance);
      const displayName = resolveDisplayName(instance);
      regionOrders.set(instanceName, order);

      const products = await listAllProducts(instanceName);
      const metrics  = buildMetrics(instanceName, displayName, products);
      const summary  = buildSummary(instanceName, displayName, isNA, metrics);

      return { instanceName, displayName, isNorthAmerica: isNA, metrics, summary } as InstanceAuditData;
    }),
  );

  const allData: InstanceAuditData[] = [];
  for (let i = 0; i < settled.length; i++) {
    const result = settled[i];
    if (result.status === "fulfilled") {
      allData.push(result.value);
    } else {
      console.error(`[audit] Skipping instance "${names[i]}": ${result.reason?.message ?? result.reason}`);
    }
  }

  if (allData.length === 0) {
    throw new Error("All instances failed to load. Check passkeys and connectivity.");
  }

  // Sort: NA → EU → Other, then alpha by instanceName within region
  const sorted = sortByRegion(allData, regionOrders);

  const workbook = xlsx.utils.book_new();

  // Sheet 1: single comprehensive audit report
  xlsx.utils.book_append_sheet(workbook, createAuditReportSheet(sorted, generatedAt), "Audit Report");

  // Per-instance product data sheets
  for (const d of sorted) {
    const sheetName = d.instanceName.replace(/[^a-zA-Z0-9-_ ]/g, "_").slice(0, 31);
    xlsx.utils.book_append_sheet(
      workbook,
      createProductDataSheet(d.instanceName, d.displayName, d.metrics, generatedAt),
      sheetName,
    );
  }

  await mkdir(outputDir, { recursive: true });

  const label    = names.length === 1 ? names[0] : `multi-instance_${names.length}instances`;
  const filename = `${label}_Audit_Report_${formatTimestampForFilename(new Date(generatedAt))}.xlsx`;
  const outputPath = path.join(outputDir, filename);
  xlsx.writeFile(workbook, outputPath);

  return {
    instanceNames: sorted.map((d) => d.instanceName),
    outputPath,
    totalProducts: sorted.reduce((s, d) => s + d.summary.totalProducts, 0),
    generatedAt,
  };
}
