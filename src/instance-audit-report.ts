import { mkdir } from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import { listProducts } from "./bazaarvoice.js";
import { isProductActive, resolveProductStatus } from "./product-status.js";
import type { Product } from "./types.js";

const DEFAULT_REPORTS_DIR = path.resolve(process.cwd(), "reports");

const DOCS_LINK =
  "https://pgone.sharepoint.com/:p:/s/GlobalProductServices/IQAhg8ackh-HQ77DjbpXzz43ASkqCZKDUUbpZ9z4fYTL8-M?e=vhPjwX&xsdata=MDV8MDJ8fGU2MDE3ZGRkZmM5ODRiZTcwMGQ4MDhkZTlhYmYwNjE1fGZmMzU1Mjg5NzIxZTRkZDdhNjYzYWZlYzYyYWI5ZDU0fDB8MHw2MzkxMTgzNDIzNjQ2NTk5Mjh8VW5rbm9";

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

/** Maximum columns across any section table (Section 2 = 14 cols). */
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
  isNorthAmerica: boolean;
  totalProducts: number;
  activeProducts: number;
  inactiveProducts: number;
  percentInactive: number;
  totalReviews: number;
  /** Sum of reviews on inactive products — these reviews are currently hidden. */
  reviewsHiddenOnInactive: number;
  percentReviewsHidden: number;
  /** NA instances: missing UPC count.  All other instances: missing EAN count. */
  missingUpcOrEan: number;
  missingDescription: number;
  missingImageUrl: number;
  missingName: number;
  missingPageUrl: number;
  missingBrand: number;
};

type InstanceAuditData = {
  instanceName: string;
  metrics: ProductMetrics[];
  summary: InstanceSummary;
};

// ─── Region detection ─────────────────────────────────────────────────────────

const NA_SUBSTRINGS = ["-us", "-ca", "_us", "_ca", "en-us", "en-ca", "fr-ca"];
const EU_SUBSTRINGS = [
  "-uk", "-gb", "-de", "-fr", "-es", "-it", "-nl", "-pl", "-pt",
  "-be", "-at", "-ch", "-se", "-no", "-dk", "-fi", "-ie", "-cz",
  "-sk", "-hu", "-ro", "-bg", "-hr", "-si", "-ee", "-lv", "-lt",
];

function isNorthAmericaInstance(name: string): boolean {
  const lower = name.toLowerCase();
  return NA_SUBSTRINGS.some((s) => lower.includes(s));
}

function regionOrder(name: string): number {
  if (isNorthAmericaInstance(name)) return 0;
  const lower = name.toLowerCase();
  if (EU_SUBSTRINGS.some((s) => lower.includes(s))) return 1;
  return 2;
}

function sortByRegion<T extends { instanceName: string }>(items: T[]): T[] {
  return [...items].sort((a, b) => {
    const diff = regionOrder(a.instanceName) - regionOrder(b.instanceName);
    return diff !== 0 ? diff : a.instanceName.localeCompare(b.instanceName);
  });
}

// ─── Formatting helpers ───────────────────────────────────────────────────────

function round(value: number, decimals = 2): number {
  return Math.round(value * 10 ** decimals) / 10 ** decimals;
}

function pct(part: number, total: number, decimals = 1): number {
  return total > 0 ? round((part / total) * 100, decimals) : 0;
}

function formatTimestampForFilename(date: Date): string {
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
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

// ─── Cell-writing primitives ──────────────────────────────────────────────────

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
  const v = value ?? "";
  ws[addr] = { v, t: typeof v === "number" ? "n" : "s" };

  if (style?.bgColor) {
    ws[addr].fill = { patternType: "solid", fgColor: { rgb: style.bgColor } };
  }

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

/**
 * Write a section heading that spans MAX_COLS columns.
 * Adds a merge record and returns the next available row.
 */
function heading(ws: xlsx.WorkSheet, merges: xlsx.Range[], row: number, text: string): number {
  sc(ws, row, 0, text, { bold: true, fontSize: 12, fontColor: C.DARK_BLUE_TEXT, bgColor: C.LIGHT_BLUE_BG });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.LIGHT_BLUE_BG });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  return row + 2; // blank line after heading
}

/** Write a coloured sub-section banner. Returns next row. */
function subHeading(
  ws: xlsx.WorkSheet,
  merges: xlsx.Range[],
  row: number,
  text: string,
  bgColor: string,
  fontColor: string = C.WHITE,
): number {
  sc(ws, row, 0, text, { bold: true, fontColor, bgColor });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  return row + 1;
}

/** Write a table header row. Returns next row. */
function tHeaders(ws: xlsx.WorkSheet, row: number, headers: string[]): number {
  for (const [c, h] of headers.entries()) {
    sc(ws, row, c, h, { bold: true, fontColor: C.WHITE, bgColor: C.HEADER_BLUE_BG, wrapText: true });
  }
  return row + 1;
}

/** Write a plain data row. Returns next row. */
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

/**
 * Write a data row whose background colour is determined by
 * the string value at `priorityCol`.
 */
function pRow(
  ws: xlsx.WorkSheet,
  row: number,
  values: (string | number | null)[],
  priorityCol: number,
): number {
  const bg = priorityColor(String(values[priorityCol] ?? ""));
  return dRow(ws, row, values, { bgColor: bg, wrapText: true });
}

/** Write a bold TOTAL row with green background. Returns next row. */
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

function buildMetrics(instanceName: string, products: Product[]): ProductMetrics[] {
  return products.map((product) => ({
    instanceName,
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

function buildSummary(instanceName: string, metrics: ProductMetrics[]): InstanceSummary {
  const isNA           = isNorthAmericaInstance(instanceName);
  const totalProducts  = metrics.length;
  const active         = metrics.filter((m) => m.isActive).length;
  const inactive       = totalProducts - active;
  const totalReviews   = metrics.reduce((s, m) => s + m.reviewCount, 0);
  const hiddenReviews  = metrics.filter((m) => !m.isActive).reduce((s, m) => s + m.reviewCount, 0);

  return {
    instanceName,
    isNorthAmerica:          isNA,
    totalProducts,
    activeProducts:          active,
    inactiveProducts:        inactive,
    percentInactive:         pct(inactive, totalProducts),
    totalReviews,
    reviewsHiddenOnInactive: hiddenReviews,
    percentReviewsHidden:    pct(hiddenReviews, totalReviews),
    // NA: missing UPC — all others: missing EAN
    missingUpcOrEan:    isNA ? metrics.filter((m) => !m.hasUpc).length
                             : metrics.filter((m) => !m.hasEan).length,
    missingDescription: metrics.filter((m) => !m.hasDescription).length,
    missingImageUrl:    metrics.filter((m) => !m.hasImageUrl).length,
    missingName:        metrics.filter((m) => !m.hasName).length,
    missingPageUrl:     metrics.filter((m) => !m.hasPageUrl).length,
    missingBrand:       metrics.filter((m) => !m.hasBrand).length,
  };
}

// ─── Per-instance PRODUCT DATA sheet ─────────────────────────────────────────

function createProductDataSheet(
  instanceName: string,
  metrics: ProductMetrics[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  const displayName = instanceName
    .split("-")
    .map((p) => p.charAt(0).toUpperCase() + p.slice(1))
    .join(" ");

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
    "Product Family", "Product Family - Expand",
    "# Approved Reviews", "# Family Reviews", "Family Average Rating",
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
      typeof avg === "number" ? round(avg) : "",
    ]);
  }

  ws["!ref"]  = xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row - 1, c: 14 } });
  ws["!cols"] = [15, 12, 18, 12, 30, 50, 30, 30, 15, 15, 15, 20, 18, 18, 18].map((wch) => ({ wch }));
  return ws;
}

// ─── Single "Audit Report" sheet (9 sections) ────────────────────────────────

function createAuditReportSheet(allData: InstanceAuditData[], generatedAt: string): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  const merges: xlsx.Range[] = [];
  let row = 0;

  // Data is already sorted by region (NA → EU → Other)
  const data = allData;

  // ── Report title ──────────────────────────────────────────────────────────
  sc(ws, row, 0, "Bazaarvoice Catalog Audit Report", {
    bold: true, fontSize: 16, fontColor: C.WHITE, bgColor: C.NAVY_BG,
  });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.NAVY_BG });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row++;
  sc(ws, row, 0, formatRunDate(new Date(generatedAt)));
  sc(ws, row, 2, `Instances: ${data.map((d) => d.instanceName).join(", ")}`);
  row += 2;

  // ── Pre-calculate cross-instance aggregates ────────────────────────────────
  const allInactiveWithReviews = data.flatMap((d) =>
    d.metrics.filter((m) => !m.isActive && m.reviewCount > 0),
  );
  const totalHiddenReviews = allInactiveWithReviews.reduce((s, m) => s + m.reviewCount, 0);
  const totalReviewsGlobal = data.reduce((s, d) => s + d.summary.totalReviews, 0);
  const pctHiddenGlobal    = pct(totalHiddenReviews, totalReviewsGlobal);

  // ── TOP PRIORITY ALERT ───────────────────────────────────────────────────────
  sc(ws, row, 0, "🔴  TOP PRIORITY ALERT", {
    bold: true, fontSize: 13, fontColor: C.WHITE, bgColor: C.ALERT_BANNER,
  });
  for (let c = 1; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.ALERT_BANNER });
  merges.push({ s: { r: row, c: 0 }, e: { r: row, c: MAX_COLS - 1 } });
  row++;

  const alertRows: [string, string | number][] = [
    ["Most Critical Issue",
      "Inactive Products with Reviews — Reviews are currently hidden because they are associated with inactive products."],
    ["Total Products Impacted",           allInactiveWithReviews.length],
    ["Total Reviews Currently Hidden",    totalHiddenReviews],
    ["% Reviews Hidden",                  `${pctHiddenGlobal}%`],
    ["Required Action Summary",
      `Reactivate or update visibility of ${allInactiveWithReviews.length} inactive product(s) across ${data.length} instance(s).`],
    ["Recommended Fix",
      "Bulk update product status to Active in Bazaarvoice Workbench / Admin. Re-index catalog after update."],
    ["Suggestion Impact", "CRUCIAL"],
  ];
  for (const [label, value] of alertRows) {
    sc(ws, row, 0, label, { bold: true, bgColor: C.ALERT_LABEL });
    sc(ws, row, 1, value, { wrapText: true, bgColor: C.ALERT_VALUE });
    for (let c = 2; c < MAX_COLS; c++) sc(ws, row, c, "", { bgColor: C.ALERT_VALUE });
    row++;
  }
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 1: CRITICAL FIELD DEFINITIONS FOR AUDIT
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 1: CRITICAL FIELD DEFINITIONS FOR AUDIT");

  const sec1H = [
    "Critical Field", "Column Reference", "Why Critical",
    "Business Impact if Missing", "Recommended Action", "Documentation Reference",
  ];
  row = tHeaders(ws, row, sec1H);

  const fieldDefs: string[][] = [
    ["Product ID", "C",
      "Unique identifier for tracking and linking",
      "Cannot identify or track products; data chaos and reporting failures",
      "Ensure every product has a unique, non-null Product ID", DOCS_LINK],
    ["UPC / EAN", "I / J",
      "Links products to inventory, commerce, and syndication systems",
      "Inventory mismatch; syndication failures; lost sales",
      "Add missing UPC (North America) or EAN (other regions) for all products", DOCS_LINK],
    ["Product Name", "E",
      "Customer-facing display name; drives search and discoverability",
      "Low search ranking; poor user experience; reduced discoverability",
      "Ensure all products have descriptive, non-empty names", DOCS_LINK],
    ["Product Description", "F",
      "SEO content and customer understanding of the product",
      "Reduced conversion rates; low organic search ranking",
      "Enrich all product descriptions from master content source", DOCS_LINK],
    ["Product Page URL / PDP URL", "G",
      "Primary navigation link between reviews and the product detail page",
      "Broken user journey; lost sales attribution",
      "Verify and update all Product Page URLs to current PDP URLs", DOCS_LINK],
    ["Image URL", "H",
      "Visual representation of the product",
      "Poor conversion; high bounce rate; reduced customer trust",
      "Upload or link product images for all catalogue entries", DOCS_LINK],
    ["Brand", "B",
      "Brand classification for filtering and compliance",
      "Poor filtering; brand compliance issues; incorrect attribution",
      "Ensure correct, consistent brand names for all products", DOCS_LINK],
    ["Product Status", "D",
      "Controls product visibility and review display in Bazaarvoice",
      "Hidden reviews; inactive products consuming catalog slots",
      "Review and update status for all inactive products that have reviews", DOCS_LINK],
  ];
  for (const fd of fieldDefs) row = dRow(ws, row, fd);
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 2: VOLUME-BASED SUMMARY BY INSTANCE
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 2: VOLUME-BASED SUMMARY BY INSTANCE");

  const sec2H = [
    "Instance / Region",
    "Total Products",
    "Active Products",
    "Inactive Products",
    "% Inactive",
    "Total Reviews",
    "Reviews Hidden / Tagged to Inactive Product",
    "% Reviews Hidden / Tagged",
    "Missing UPC (NA) / EAN (Other)",
    "Missing Description",
    "Missing Image URL",
    "Business Impact",
    "Recommended Action",
    "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec2H);

  let s2TotProd = 0, s2TotAct = 0, s2TotInact = 0, s2TotRev = 0, s2TotHid = 0;
  let s2TotUE = 0, s2TotDesc = 0, s2TotImg = 0;

  for (const d of data) {
    const s = d.summary;
    s2TotProd  += s.totalProducts;
    s2TotAct   += s.activeProducts;
    s2TotInact += s.inactiveProducts;
    s2TotRev   += s.totalReviews;
    s2TotHid   += s.reviewsHiddenOnInactive;
    s2TotUE    += s.missingUpcOrEan;
    s2TotDesc  += s.missingDescription;
    s2TotImg   += s.missingImageUrl;

    const imp = s.reviewsHiddenOnInactive > 0 ? "CRUCIAL"
              : s.missingUpcOrEan > 0          ? "HIGH"
              : "MEDIUM";

    row = pRow(ws, row, [
      d.instanceName,
      s.totalProducts, s.activeProducts, s.inactiveProducts,
      `${s.percentInactive}%`,
      s.totalReviews,
      s.reviewsHiddenOnInactive,
      `${s.percentReviewsHidden}%`,
      s.missingUpcOrEan, s.missingDescription, s.missingImageUrl,
      s.reviewsHiddenOnInactive > 0
        ? `${s.reviewsHiddenOnInactive} reviews are currently hidden due to inactive products`
        : "No reviews currently hidden",
      s.reviewsHiddenOnInactive > 0
        ? "Reactivate inactive products immediately"
        : s.missingUpcOrEan > 0
          ? "Add missing UPC/EAN codes"
          : "Monitor and maintain catalog quality",
      imp,
    ], 13);
  }

  // TOTAL row
  row = totRow(ws, row, [
    "TOTAL",
    s2TotProd, s2TotAct, s2TotInact,
    `${pct(s2TotInact, s2TotProd)}%`,
    s2TotRev, s2TotHid,
    `${pct(s2TotHid, s2TotRev)}%`,
    s2TotUE, s2TotDesc, s2TotImg,
    s2TotHid > 0 ? `${s2TotHid} reviews hidden across all instances` : "No reviews hidden",
    s2TotHid > 0 ? "Immediate reactivation required" : "Maintain catalog quality",
    s2TotHid > 0 ? "CRUCIAL" : "MEDIUM",
  ]);
  row++;

  const urgentInstances = data.filter((d) => d.summary.reviewsHiddenOnInactive > 0).map((d) => d.instanceName);
  if (urgentInstances.length > 0) {
    sc(ws, row, 0,
      `⚠️  Instances requiring immediate attention: ${urgentInstances.join(", ")}`,
      { bold: true, fontColor: C.ALERT_BANNER, wrapText: true },
    );
    row++;
  }
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 3: ISSUES WITH BUSINESS IMPACT, RECOMMENDATIONS & HOW TO FIX
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 3: ISSUES WITH BUSINESS IMPACT, RECOMMENDATIONS & HOW TO FIX");

  const sec3H = [
    "Instance / Region", "#", "Issue Category", "Critical Field",
    "Volume Impacted", "Business Impact", "Recommended Action", "How to Fix", "Suggestion Impact",
  ];

  // — 3.1  Inactive Products → Hidden Reviews (CRUCIAL) ─────────────────────
  row = subHeading(ws, merges, row,
    "1.  Inactive Products → Reviews Hidden / Tagged to Inactive Product (CRUCIAL)",
    C.CRUCIAL);
  row = tHeaders(ws, row, sec3H);
  {
    let n = 1;
    let hasData = false;
    for (const d of data) {
      const inactive = d.metrics.filter((m) => !m.isActive && m.reviewCount > 0);
      if (inactive.length === 0) continue;
      hasData = true;
      const hidden = inactive.reduce((s, m) => s + m.reviewCount, 0);
      row = pRow(ws, row, [
        d.instanceName, n++, "Inactive Products with Hidden Reviews", "Product Status",
        `${inactive.length} products / ${hidden} reviews hidden`,
        "Reviews are currently hidden because they are associated with inactive products. " +
          "This directly reduces social proof, customer trust, and conversion rates.",
        `Reactivate ${inactive.length} product(s) or update their visibility settings immediately.`,
        "1. Log in to Bazaarvoice Workbench.  " +
          "2. Navigate to Catalog → Products.  " +
          "3. Filter by Status = Inactive.  " +
          "4. Bulk update Status to Active.  " +
          "5. Re-index catalog.  " +
          "6. Verify reviews are visible on the product page.",
        "CRUCIAL",
      ], 8);
    }
    if (!hasData) row = dRow(ws, row, ["No inactive products with hidden reviews found.", ...Array(8).fill("")]);
  }
  row++;

  // — 3.2  Product ID Issues (HIGH) ─────────────────────────────────────────
  row = subHeading(ws, merges, row, "2.  Product ID Issues (HIGH)", C.HIGH);
  row = tHeaders(ws, row, sec3H);
  {
    let n = 1;
    let hasData = false;
    for (const d of data) {
      const numeric = d.metrics.filter((m) => /^\d+$/.test(m.product.Id));
      if (numeric.length === 0) continue;
      hasData = true;
      row = pRow(ws, row, [
        d.instanceName, n++, "Purely Numeric Product IDs", "Product ID",
        numeric.length,
        "Purely numeric Product IDs risk collision with retailer or ERP numeric keys, " +
          "causing data mapping errors and reporting inaccuracies.",
        "Migrate Product IDs to an alphanumeric, brand-prefixed format (e.g., BRAND-12345).",
        "1. Export catalog.  " +
          "2. Identify all purely numeric IDs.  " +
          "3. Coordinate with IT and Bazaarvoice support to plan a safe ID migration.  " +
          "4. Ensure all downstream systems are updated before migrating.",
        "HIGH",
      ], 8);
    }
    if (!hasData) row = dRow(ws, row, ["No Product ID standardisation issues found.", ...Array(8).fill("")]);
  }
  row++;

  // — 3.3  UPC/EAN Missing (HIGH) ───────────────────────────────────────────
  row = subHeading(ws, merges, row,
    "3.  UPC / EAN Data Missing Issues (HIGH)  [North America: UPC | All other regions: EAN]",
    C.HIGH);
  row = tHeaders(ws, row, sec3H);
  {
    let n = 1;
    let hasData = false;
    for (const d of data) {
      if (d.summary.missingUpcOrEan === 0) continue;
      hasData = true;
      const field  = d.summary.isNorthAmerica ? "UPC" : "EAN";
      const region = d.summary.isNorthAmerica ? "North American" : "non-North American";
      row = pRow(ws, row, [
        d.instanceName, n++, `Missing ${field} Values`, `${field} (${region} region)`,
        d.summary.missingUpcOrEan,
        `Missing ${field} values prevent inventory synchronisation, syndication, and commerce platform integrations ` +
          `for this ${region} instance.`,
        `Add ${field} codes to all ${d.summary.missingUpcOrEan} affected products.`,
        `1. Export catalog for ${d.instanceName}.  ` +
          `2. Filter products where ${field} is blank.  ` +
          `3. Obtain ${field} from product master data or supplier.  ` +
          "4. Upload via Bazaarvoice catalog feed.  " +
          "5. Validate after re-index.",
        "HIGH",
      ], 8);
    }
    if (!hasData) row = dRow(ws, row, ["No UPC/EAN data quality issues found.", ...Array(8).fill("")]);
  }
  row++;

  // — 3.4  Missing Crucial Fields (MEDIUM) ──────────────────────────────────
  row = subHeading(ws, merges, row, "4.  Missing Crucial Fields (MEDIUM)", C.MEDIUM, "000000");
  row = tHeaders(ws, row, sec3H);
  {
    let n = 1;
    let hasData = false;
    for (const d of data) {
      const s = d.summary;
      if (s.missingDescription > 0) {
        hasData = true;
        row = pRow(ws, row, [
          d.instanceName, n++, "Missing Product Description", "Description",
          s.missingDescription,
          "Missing descriptions reduce organic search ranking and customer understanding, leading to lower conversion rates.",
          "Enrich product descriptions for all affected products.",
          "1. Export catalog.  2. Filter where Description is blank.  " +
            "3. Source from PIM or master data.  4. Upload via catalog feed.",
          "MEDIUM",
        ], 8);
      }
      if (s.missingImageUrl > 0) {
        hasData = true;
        row = pRow(ws, row, [
          d.instanceName, n++, "Missing Image URL", "Image URL",
          s.missingImageUrl,
          "Missing product images reduce customer engagement, increase bounce rate, and negatively affect conversion.",
          "Upload or link product images for all affected products.",
          "1. Export catalog.  2. Filter where Image URL is blank.  " +
            "3. Upload images to CDN or content repository.  4. Update Image URL in catalog feed.",
          "MEDIUM",
        ], 8);
      }
      if (s.missingPageUrl > 0) {
        hasData = true;
        row = pRow(ws, row, [
          d.instanceName, n++, "Missing Product Page URL (PDP URL)", "Product Page URL",
          s.missingPageUrl,
          "Missing PDP URLs break the user journey from reviews to purchase, causing lost attribution and lost sales.",
          "Update Product Page URLs for all affected products.",
          "1. Export catalog.  2. Filter where Product Page URL is blank.  " +
            "3. Map Product IDs to live PDP URLs.  4. Update via catalog feed.",
          "MEDIUM",
        ], 8);
      }
      if (s.missingName > 0) {
        hasData = true;
        row = pRow(ws, row, [
          d.instanceName, n++, "Missing Product Name", "Product Name",
          s.missingName,
          "Missing product names prevent proper display, search, and filtering functionality.",
          "Add product names for all affected products.",
          "1. Export catalog.  2. Filter where Product Name is blank.  " +
            "3. Source from PIM or master data.  4. Upload via feed.",
          "MEDIUM",
        ], 8);
      }
    }
    if (!hasData) row = dRow(ws, row, ["No missing crucial field issues found.", ...Array(8).fill("")]);
  }
  row++;

  // — 3.5  Brand Name Issues (MEDIUM) ──────────────────────────────────────
  row = subHeading(ws, merges, row, "5.  Brand Name Issues (MEDIUM)", C.MEDIUM, "000000");
  row = tHeaders(ws, row, sec3H);
  {
    let n = 1;
    let hasData = false;
    for (const d of data) {
      // Detect capitalisation / spelling variants
      const lowerMap = new Map<string, string[]>();
      for (const m of d.metrics) {
        const b = m.product.Brand?.Name?.trim();
        if (!b) continue;
        const key = b.toLowerCase();
        const arr = lowerMap.get(key) ?? [];
        if (!arr.includes(b)) arr.push(b);
        lowerMap.set(key, arr);
      }
      const inconsistencies = [...lowerMap.values()].filter((v) => v.length > 1);
      for (const variants of inconsistencies) {
        hasData = true;
        const count = d.metrics.filter((m) => variants.includes(m.product.Brand?.Name?.trim() ?? "")).length;
        row = pRow(ws, row, [
          d.instanceName, n++, "Brand Name Inconsistency", "Brand",
          count,
          "Inconsistent brand naming reduces filtering accuracy, affects brand compliance reporting, and creates data integrity issues.",
          `Standardise to a single canonical brand name. Detected variations: ${variants.join(", ")}`,
          "1. Export catalog.  2. Identify all brand name variants.  " +
            "3. Agree on canonical brand name.  4. Bulk update via catalog feed.",
          "MEDIUM",
        ], 8);
      }
      if (inconsistencies.length === 0 && d.summary.missingBrand > 0) {
        hasData = true;
        row = pRow(ws, row, [
          d.instanceName, n++, "Missing Brand Name", "Brand",
          d.summary.missingBrand,
          "Products without a brand name cannot be properly filtered or attributed in reporting.",
          "Add brand names to all products missing a brand.",
          "1. Export catalog.  2. Filter where Brand is blank.  " +
            "3. Map to correct brand.  4. Upload via feed.",
          "MEDIUM",
        ], 8);
      }
    }
    if (!hasData) row = dRow(ws, row, ["No brand name issues found.", ...Array(8).fill("")]);
  }
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 4: INACTIVE PRODUCTS WITH REVIEWS – MIGRATION CANDIDATES
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 4: INACTIVE PRODUCTS WITH REVIEWS – MIGRATION CANDIDATES");

  const sec4H = [
    "Instance / Region", "Product ID", "Product Name", "Brand", "Product Status",
    "Total Reviews Hidden / Tagged to Inactive Product", "Reviews on Product",
    "Suggested Migration Target / Action",
    "Business Impact", "How to Fix", "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec4H);

  const sec4Rows = data
    .flatMap((d) => d.metrics.filter((m) => !m.isActive && m.reviewCount > 0))
    .sort((a, b) => b.reviewCount - a.reviewCount);

  if (sec4Rows.length > 0) {
    for (const m of sec4Rows) {
      const p = m.product;
      row = dRow(ws, row, [
        m.instanceName,
        safeCellValue(p.Id),
        safeCellValue(p.Name),
        safeCellValue(p.Brand?.Name ?? ""),
        resolveProductStatus(p),
        m.reviewCount,
        m.reviewCount,
        "Reactivate product or reassign reviews to active replacement SKU",
        "Reviews are currently hidden because this product is inactive. " +
          "This reduces social proof and customer engagement.",
        "1. Locate active replacement product (if applicable).  " +
          "2. If same product: change status to Active.  " +
          "3. If discontinued: contact Bazaarvoice support to migrate reviews to successor SKU.",
        "CRUCIAL",
      ]);
    }
  } else {
    row = dRow(ws, row, ["No inactive products with reviews found.", ...Array(10).fill("")]);
  }
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 5: PRODUCT ID STANDARDISATION ISSUES
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 5: PRODUCT ID STANDARDISATION ISSUES");

  const sec5H = [
    "Instance / Region", "Issue Type", "Examples", "Products Affected",
    "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec5H);

  let sec5HasData = false;
  for (const d of data) {
    const numeric = d.metrics.filter((m) => /^\d+$/.test(m.product.Id));
    if (numeric.length > 0) {
      sec5HasData = true;
      row = pRow(ws, row, [
        d.instanceName, "Purely Numeric Product ID",
        numeric.slice(0, 3).map((m) => m.product.Id).join(", "),
        numeric.length,
        "Numeric IDs risk collision with retailer/ERP numeric keys, causing product mapping errors.",
        "Migrate to an alphanumeric, brand-prefixed Product ID format (e.g., BRAND-12345).",
        "1. Plan ID migration with Bazaarvoice support.  " +
          "2. Update catalog feed with new IDs.  " +
          "3. Ensure all downstream systems are updated simultaneously.",
        "HIGH",
      ], 7);
    }
  }
  if (!sec5HasData) row = dRow(ws, row, ["No Product ID standardisation issues found.", ...Array(7).fill("")]);
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 6: UPC/EAN DATA QUALITY ANALYSIS
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 6: UPC/EAN DATA QUALITY ANALYSIS");
  sc(ws, row, 0,
    "Note: Multiple EAN/UPC values per product are valid and are NOT flagged as issues.",
    { italic: true, fontColor: C.DARK_BLUE_TEXT },
  );
  row += 2;

  const sec6H = [
    "Instance / Region", "Issue Type", "Products Affected", "Examples (Product IDs)",
    "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec6H);

  let sec6HasData = false;
  for (const d of data) {
    if (d.summary.missingUpcOrEan === 0) continue;
    sec6HasData = true;
    const field    = d.summary.isNorthAmerica ? "UPC" : "EAN";
    const affected = d.summary.isNorthAmerica
      ? d.metrics.filter((m) => !m.hasUpc)
      : d.metrics.filter((m) => !m.hasEan);
    row = pRow(ws, row, [
      d.instanceName,
      `Missing ${field}`,
      d.summary.missingUpcOrEan,
      affected.slice(0, 3).map((m) => m.product.Id).join(", "),
      `Missing ${field} values prevent inventory synchronisation, commerce platform integrations, ` +
        "and accurate syndication for this instance.",
      `Add ${field} codes to all ${d.summary.missingUpcOrEan} affected products using product master data or supplier feed.`,
      `1. Export catalog.  2. Filter products where ${field} is blank.  ` +
        `3. Obtain ${field} from supplier/PIM.  ` +
        "4. Update via Bazaarvoice catalog feed.  5. Validate post-index.",
      "HIGH",
    ], 7);
  }
  if (!sec6HasData) row = dRow(ws, row, ["No UPC/EAN data quality issues found.", ...Array(7).fill("")]);
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 7: MISSING CRUCIAL FIELDS
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 7: MISSING CRUCIAL FIELDS");

  const sec7H = [
    "Instance / Region",
    "Missing Image URLs", "Missing Product Names", "Missing PDP URLs",
    "Missing Descriptions", "Missing Brand",
    "Total Missing Fields", "Business Impact", "How to Fix", "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec7H);

  let s7Img = 0, s7Name = 0, s7Url = 0, s7Desc = 0, s7Brand = 0;
  for (const d of data) {
    const s  = d.summary;
    s7Img   += s.missingImageUrl;
    s7Name  += s.missingName;
    s7Url   += s.missingPageUrl;
    s7Desc  += s.missingDescription;
    s7Brand += s.missingBrand;
    const tot = s.missingImageUrl + s.missingName + s.missingPageUrl + s.missingDescription + s.missingBrand;
    const imp = tot > 50 ? "HIGH" : tot > 10 ? "MEDIUM" : "LOW";
    row = pRow(ws, row, [
      d.instanceName,
      s.missingImageUrl, s.missingName, s.missingPageUrl,
      s.missingDescription, s.missingBrand, tot,
      tot > 0 ? "Missing fields reduce discoverability, conversion rate, and data integrity." : "No missing fields",
      tot > 0 ? "Export catalog → identify gaps → source data from PIM → upload via feed." : "N/A",
      imp,
    ], 9);
  }
  row = totRow(ws, row, [
    "TOTAL", s7Img, s7Name, s7Url, s7Desc, s7Brand,
    s7Img + s7Name + s7Url + s7Desc + s7Brand,
    "", "", "",
  ]);
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 8: BRAND CONSISTENCY CHECK
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 8: BRAND CONSISTENCY CHECK");

  const sec8H = [
    "Instance / Region", "Brand Variations Found", "Count", "Issue",
    "Business Impact", "Recommendation", "How to Fix", "Suggestion Impact",
  ];
  row = tHeaders(ws, row, sec8H);

  let sec8HasData = false;
  for (const d of data) {
    const lowerMap = new Map<string, string[]>();
    for (const m of d.metrics) {
      const b = m.product.Brand?.Name?.trim();
      if (!b) continue;
      const key = b.toLowerCase();
      const arr = lowerMap.get(key) ?? [];
      if (!arr.includes(b)) arr.push(b);
      lowerMap.set(key, arr);
    }
    const issues = [...lowerMap.values()].filter((v) => v.length > 1);
    for (const variants of issues) {
      sec8HasData = true;
      const count = d.metrics.filter((m) => variants.includes(m.product.Brand?.Name?.trim() ?? "")).length;
      row = pRow(ws, row, [
        d.instanceName, variants.join(" / "), count,
        `Capitalisation or spelling variation detected: ${variants.join(", ")}`,
        "Inconsistent brand names affect filtering accuracy, brand compliance reporting, and data integrity.",
        `Standardise to: "${variants[0]}" across all product entries.`,
        "1. Export catalog.  2. Identify all brand name variants.  " +
          "3. Agree on canonical brand name.  4. Bulk update via catalog feed.",
        "MEDIUM",
      ], 7);
    }
  }
  if (!sec8HasData) row = dRow(ws, row, ["No brand consistency issues found.", ...Array(7).fill("")]);
  row++;

  // ═══════════════════════════════════════════════════════════════════════════
  // SECTION 9: PRIORITY ACTION ITEMS BY INSTANCE
  // ═══════════════════════════════════════════════════════════════════════════
  row = heading(ws, merges, row, "SECTION 9: PRIORITY ACTION ITEMS BY INSTANCE");

  const sec9H = [
    "Instance / Region", "Priority Rank", "Issue",
    "Impact Level", "Action Required", "How to Fix", "Expected Benefit",
  ];
  row = tHeaders(ws, row, sec9H);

  for (const d of data) {
    const s          = d.summary;
    let rank         = 1;
    const field      = s.isNorthAmerica ? "UPC" : "EAN";
    const inactive   = d.metrics.filter((m) => !m.isActive && m.reviewCount > 0);
    const hiddenCnt  = inactive.reduce((sum, m) => sum + m.reviewCount, 0);
    const numericCnt = d.metrics.filter((m) => /^\d+$/.test(m.product.Id)).length;
    const missFlds   = s.missingDescription + s.missingImageUrl + s.missingPageUrl + s.missingName;

    if (inactive.length > 0) {
      row = pRow(ws, row, [
        d.instanceName, rank++,
        `${inactive.length} Inactive Products — ${hiddenCnt} Reviews Currently Hidden`,
        "CRUCIAL",
        `Reactivate ${inactive.length} product(s) to make ${hiddenCnt} reviews visible`,
        "Workbench → Catalog → Products → Filter Inactive → Bulk Activate → Re-index",
        `Restore ${hiddenCnt} hidden reviews; recover social proof and conversion impact`,
      ], 3);
    }
    if (numericCnt > 0) {
      row = pRow(ws, row, [
        d.instanceName, rank++,
        `${numericCnt} Products with Purely Numeric Product IDs`,
        "HIGH",
        "Migrate Product IDs to alphanumeric, brand-prefixed format",
        "Plan migration with BV support; update catalog feed and all downstream systems",
        "Prevent data mapping conflicts; improve cross-system data integrity",
      ], 3);
    }
    if (s.missingUpcOrEan > 0) {
      row = pRow(ws, row, [
        d.instanceName, rank++,
        `${s.missingUpcOrEan} Products Missing ${field}`,
        "HIGH",
        `Add ${field} codes to all affected products`,
        `Export → filter blank ${field} → source from PIM/supplier → upload via feed → validate`,
        "Restore inventory sync; enable commerce platform integrations",
      ], 3);
    }
    if (missFlds > 0) {
      row = pRow(ws, row, [
        d.instanceName, rank++,
        `${missFlds} Products with Missing Crucial Fields (description, image, URL, name)`,
        "MEDIUM",
        "Enrich missing product fields across all affected products",
        "Export catalog → filter blanks by field → source from PIM → upload via feed",
        "Improve SEO ranking, conversion rates, and customer experience",
      ], 3);
    }
    if (s.missingBrand > 0) {
      row = pRow(ws, row, [
        d.instanceName, rank++,
        `${s.missingBrand} Products Missing Brand Name`,
        "MEDIUM",
        "Add brand name to all products missing a brand",
        "Export → filter blank Brand → map to correct brand → upload via feed",
        "Improve brand compliance, filtering, and reporting accuracy",
      ], 3);
    }
  }
  row++;

  // ── Finalise ──────────────────────────────────────────────────────────────
  ws["!ref"]    = xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: row - 1, c: MAX_COLS - 1 } });
  ws["!merges"] = merges;
  // Column widths tuned for Section 2 (widest section)
  ws["!cols"] = [22, 13, 30, 22, 16, 36, 40, 45, 16, 36, 36, 36, 25, 16].map((wch) => ({ wch }));

  return ws;
}

// ─── Public API ───────────────────────────────────────────────────────────────

function resolveInstanceNames(options: AuditReportOptions): string[] {
  if (options.instanceNames?.length) return [...new Set(options.instanceNames)];
  if (options.instanceName) return [options.instanceName];
  throw new Error("Provide instanceName or instanceNames in AuditReportOptions.");
}

/**
 * Generate a multi-instance catalog audit workbook.
 *
 * Output:
 * - Sheet 1 "Audit Report"  — single comprehensive sheet with 9 analysis sections.
 * - Sheet 2…N               — one PRODUCT DATA sheet per instance, named after the instance.
 *
 * Accepts a single instanceName (backward-compatible) or an array via instanceNames.
 * Instances are sorted: North America → Europe → Other, then alphabetically within each region.
 */
export async function generateInstanceAuditReport(
  options: AuditReportOptions = {},
): Promise<AuditReportResult> {
  const outputDir   = options.outputDir ?? DEFAULT_REPORTS_DIR;
  const generatedAt = new Date().toISOString();

  const names = resolveInstanceNames(options);

  // Fetch all instances concurrently — scales to any number of instances.
  const allData: InstanceAuditData[] = await Promise.all(
    names.map(async (instanceName) => {
      const products = await listAllProducts(instanceName);
      const metrics  = buildMetrics(instanceName, products);
      const summary  = buildSummary(instanceName, metrics);
      return { instanceName, metrics, summary };
    }),
  );

  // Sort: NA → EU → Other (then alphabetical within region)
  const sorted = sortByRegion(allData);

  const workbook = xlsx.utils.book_new();

  // Sheet 1: comprehensive Audit Report
  xlsx.utils.book_append_sheet(workbook, createAuditReportSheet(sorted, generatedAt), "Audit Report");

  // Sheets 2…N: per-instance product data, named after the instance
  for (const d of sorted) {
    const sheetName = d.instanceName.replace(/[^a-zA-Z0-9-_ ]/g, "_").slice(0, 31);
    xlsx.utils.book_append_sheet(
      workbook,
      createProductDataSheet(d.instanceName, d.metrics, generatedAt),
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
