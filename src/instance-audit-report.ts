import { mkdir } from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import { findInstanceByName, listProducts } from "./bazaarvoice.js";
import { isProductActive, resolveProductStatus } from "./product-status.js";
import type { Product } from "./types.js";

const DEFAULT_INSTANCE_NAME = "pampers-en-us";
const DEFAULT_REPORTS_DIR = path.resolve(process.cwd(), "reports");

type AuditReportOptions = {
  instanceName?: string;
  outputDir?: string;
};

type AuditReportResult = {
  instanceName: string;
  outputPath: string;
  totalProducts: number;
  activeProducts: number;
  inactiveProducts: number;
  generatedAt: string;
};

type ProductMetrics = {
  product: Product;
  isActive: boolean;
  reviewCount: number;
  hasDescription: boolean;
  hasImageUrl: boolean;
  hasPageUrl: boolean;
  hasUpc: boolean;
  hasEan: boolean;
};

type InstanceSummary = {
  instanceName: string;
  totalProducts: number;
  activeProducts: number;
  inactiveProducts: number;
  percentInactive: number;
  totalReviews: number;
  reviewsOnInactive: number;
  percentReviewsAtRisk: number;
  missingUpcEan: number;
  missingDescription: number;
  missingImageUrl: number;
};

function formatRunDate(date: Date): string {
  const month = date.toLocaleString("en-US", { month: "short", timeZone: "UTC" });
  const day = String(date.getUTCDate()).padStart(2, "0");
  const year = date.getUTCFullYear();
  const hours = String(date.getUTCHours() % 12 || 12).padStart(2, "0");
  const minutes = String(date.getUTCMinutes()).padStart(2, "0");
  const meridiem = date.getUTCHours() >= 12 ? "PM" : "AM";
  return `Run Date: ${month} ${day}, ${year} ${hours}:${minutes} ${meridiem} UTC`;
}

function formatTimestampForFilename(date: Date): string {
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function round(value: number, decimals = 2): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

// Safely truncate cell values to Excel's 32,767 character limit
function safeCellValue(value: unknown): string | number {
  if (value === null || value === undefined) {
    return "";
  }
  if (typeof value === "number") {
    return value;
  }
  const stringValue = String(value);
  if (stringValue.length > 32700) {
    return stringValue.substring(0, 32700) + "...";
  }
  return stringValue;
}

async function listAllProducts(instanceName: string): Promise<Product[]> {
  const allProducts: Product[] = [];
  const pageSize = 100;

  for (let offset = 0; ; offset += pageSize) {
    const page = await listProducts(instanceName, {
      limit: pageSize,
      offset,
      includeStats: true,
    });

    allProducts.push(...page.Results);

    // Check if we've fetched all products
    // Use actual results length instead of page.Limit to handle partial final pages
    const totalFetched = offset + page.Results.length;
    if (totalFetched >= page.TotalResults || page.Results.length < pageSize) {
      break;
    }
  }

  return allProducts;
}

function getProductMetrics(product: Product): ProductMetrics {
  const isActive = isProductActive(product);
  const reviewCount = product.ReviewStatistics?.TotalReviewCount ?? 0;
  const hasDescription = Boolean(product.Description?.trim());
  const hasImageUrl = Boolean(product.ImageUrl?.trim());
  const hasPageUrl = Boolean(product.ProductPageUrl?.trim());
  const hasUpc = Boolean(product.UPCs && product.UPCs.length > 0);
  const hasEan = Boolean(product.EANs && product.EANs.length > 0);

  return {
    product,
    isActive,
    reviewCount,
    hasDescription,
    hasImageUrl,
    hasPageUrl,
    hasUpc,
    hasEan,
  };
}

function calculateInstanceSummary(metrics: ProductMetrics[]): InstanceSummary & { dummy?: string } {
  const totalProducts = metrics.length;
  const activeProducts = metrics.filter((m) => m.isActive).length;
  const inactiveProducts = totalProducts - activeProducts;
  const totalReviews = metrics.reduce((sum, m) => sum + m.reviewCount, 0);
  const reviewsOnInactive = metrics.filter((m) => !m.isActive && m.reviewCount > 0).reduce((sum, m) => sum + m.reviewCount, 0);
  const missingUpcEan = metrics.filter((m) => !m.hasUpc && !m.hasEan).length;
  const missingDescription = metrics.filter((m) => !m.hasDescription).length;
  const missingImageUrl = metrics.filter((m) => !m.hasImageUrl).length;

  return {
    instanceName: "",
    totalProducts,
    activeProducts,
    inactiveProducts,
    percentInactive: totalProducts > 0 ? round((inactiveProducts / totalProducts) * 100, 1) : 0,
    totalReviews,
    reviewsOnInactive,
    percentReviewsAtRisk: totalReviews > 0 ? round((reviewsOnInactive / totalReviews) * 100, 1) : 0,
    missingUpcEan,
    missingDescription,
    missingImageUrl,
  };
}

function applyHeaderStyle(sheet: xlsx.WorkSheet, row: number, cols: number, bgColor: string = "0070C0") {
  for (let col = 0; col < cols; col++) {
    const addr = xlsx.utils.encode_cell({ r: row, c: col });
    sheet[addr] = sheet[addr] || {};
    sheet[addr].fill = { type: "solid", fgColor: { rgb: bgColor } };
    sheet[addr].font = { bold: true, color: { rgb: "FFFFFF" } };
    sheet[addr].alignment = { horizontal: "left", vertical: "center", wrapText: true };
  }
}

function finalizeSheet(sheet: xlsx.WorkSheet, maxRow: number, maxCol: number) {
  sheet["!ref"] = xlsx.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: maxRow - 1, c: maxCol - 1 } });
}

export async function generateInstanceAuditReport(
  options: AuditReportOptions = {},
): Promise<AuditReportResult> {
  const instanceName = options.instanceName ?? DEFAULT_INSTANCE_NAME;
  const outputDir = options.outputDir ?? DEFAULT_REPORTS_DIR;
  const generatedAt = new Date().toISOString();

  const instance = findInstanceByName(instanceName);
  const products = await listAllProducts(instanceName);
  const metrics = products.map((p) => getProductMetrics(p));
  const summary = calculateInstanceSummary(metrics);
  summary.instanceName = instanceName;

  const workbook = xlsx.utils.book_new();

  // Create all sheets
  const auditDashboardSheet = createAuditDashboardSheet(instanceName, summary, metrics, generatedAt);
  const catalogAnalysisSheet = createCatalogAnalysisSheet(instanceName, summary, metrics, generatedAt);
  const suggestionsSheet = createSuggestionsSheet(instanceName, summary, metrics, generatedAt);
  const pivotSheet = createPivotSheet(instanceName, metrics);
  const productDataSheet = createProductDataSheet(instanceName, metrics, generatedAt);

  // Append sheets to workbook
  xlsx.utils.book_append_sheet(workbook, auditDashboardSheet, "AUDIT DASHBOARD");
  xlsx.utils.book_append_sheet(workbook, catalogAnalysisSheet, "CATALOG ANALYSIS");
  xlsx.utils.book_append_sheet(workbook, suggestionsSheet, "SUGGESTIONS SUMMARY");
  xlsx.utils.book_append_sheet(workbook, pivotSheet, "PIVOT DATA");
  xlsx.utils.book_append_sheet(workbook, productDataSheet, "PRODUCT DATA");

  // Write file
  await mkdir(outputDir, { recursive: true });
  const filename = `${instanceName}_Audit_Report_${formatTimestampForFilename(new Date(generatedAt))}.xlsx`;
  const outputPath = path.join(outputDir, filename);
  xlsx.writeFile(workbook, outputPath);

  return {
    instanceName,
    outputPath,
    totalProducts: summary.totalProducts,
    activeProducts: summary.activeProducts,
    inactiveProducts: summary.inactiveProducts,
    generatedAt,
  };
}

// ─── SHEET 1: AUDIT DASHBOARD ─────────────────────────────────────────────

function createAuditDashboardSheet(
  instanceName: string,
  summary: InstanceSummary & { dummy?: string },
  metrics: ProductMetrics[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  // Title
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "AUDIT DASHBOARD", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, color: { rgb: "FFFFFF" }, size: 14 };
  row += 2;

  // Top Priority Alert
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "🔴 TOP PRIORITY ALERT", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12, color: { rgb: "FF0000" } };
  row++;

  const inactiveWithReviews = metrics.filter((m) => !m.isActive && m.reviewCount > 0);
  const totalReviewsAtRisk = inactiveWithReviews.reduce((sum, m) => sum + m.reviewCount, 0);

  const alertData = [
    ["Issue", "Inactive Products with Reviews (Reviews not visible)"],
    ["Products Impacted", inactiveWithReviews.length],
    ["Reviews at Risk", totalReviewsAtRisk],
    ["% Impact", summary.percentReviewsAtRisk],
    ["Business Impact", "Loss of customer reviews and engagement; reduced social proof"],
    ["Required Action", `Migrate ${inactiveWithReviews.length} products back to Active or update visibility settings`],
  ];

  for (const [col, [label, value]] of alertData.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: label, t: "s" };
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "FFE699" } };
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true };

    ws[xlsx.utils.encode_cell({ r: row, c: 1 })] = {
      v: value,
      t: typeof value === "number" ? "n" : "s",
    };
    row++;
  }

  row += 2;

  // Critical Field Definitions
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "1️⃣ CRITICAL FIELD DEFINITIONS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12, color: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "E7E6E6" } };
  row++;

  const fieldDefs = [
    ["Field", "Why Critical", "Business Impact if Missing"],
    ["Product ID", "Unique identifier for tracking and linking", "Cannot identify products; data chaos"],
    ["Brand", "Filtering and organization", "Poor filtering; brand compliance issues"],
    ["Product Name", "Customer-facing critical field", "Low search ranking; poor UX"],
    ["Description", "SEO and customer understanding", "Reduced conversion; low engagement"],
    ["Product Page URL", "Navigation and attribution", "Broken user journey; lost sales"],
    ["Image URL", "Visual appeal and trust", "Poor conversion; high bounce rate"],
    ["UPC/EAN", "Inventory sync and commerce", "Inventory mismatch; fulfillment errors"],
    ["Product Status", "Visibility and active selling", "Wrong products shown; lost revenue"],
  ];

  applyHeaderStyle(ws, row, 3);
  for (const [i, item] of fieldDefs.entries()) {
    for (const [j, val] of item.entries()) {
      ws[xlsx.utils.encode_cell({ r: row + i, c: j })] = { v: val, t: "s" };
    }
  }

  row += fieldDefs.length + 2;

  // Instance-wise Summary
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "2️⃣ INSTANCE-WISE SUMMARY", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12, color: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "E7E6E6" } };
  row++;

  const summaryHeaders = [
    "Instance",
    "Total Products",
    "Active Products",
    "Inactive Products",
    "% Inactive",
    "Total Reviews",
    "Reviews on Inactive",
    "% Reviews at Risk",
    "Missing UPC/EAN",
    "Missing Description",
    "Missing Image URL",
  ];

  applyHeaderStyle(ws, row, summaryHeaders.length);
  for (const [i, header] of summaryHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const summaryRow = [
    instanceName,
    summary.totalProducts,
    summary.activeProducts,
    summary.inactiveProducts,
    summary.percentInactive,
    summary.totalReviews,
    summary.reviewsOnInactive,
    summary.percentReviewsAtRisk,
    summary.missingUpcEan,
    summary.missingDescription,
    summary.missingImageUrl,
  ];

  for (const [i, val] of summaryRow.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
  }
  row++;
  row++;

  // Issues, Business Impact & Fixes
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "3️⃣ ISSUES, BUSINESS IMPACT & FIXES", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12, color: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "E7E6E6" } };
  row++;

  const issuesHeaders = ["Issue #", "Issue", "Volume", "Business Impact", "Recommended Action", "Suggestion Impact"];
  applyHeaderStyle(ws, row, issuesHeaders.length);
  for (const [i, header] of issuesHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const issues = [
    [
      1,
      "Inactive Products with Reviews",
      inactiveWithReviews.length,
      "Loss of social proof; customer engagement lost",
      "Reactivate products or update visibility",
      "CRUCIAL",
    ],
    [
      2,
      "Missing UPC/EAN",
      summary.missingUpcEan,
      "Inventory sync failures; commerce integration issues",
      "Bulk add UPC/EAN codes from supplier",
      "HIGH",
    ],
    [
      3,
      "Missing Product Description",
      summary.missingDescription,
      "Low SEO ranking; poor conversion",
      "Enrich product descriptions from source",
      "MEDIUM",
    ],
    [
      4,
      "Missing Image URL",
      summary.missingImageUrl,
      "Low engagement; high bounce rate",
      "Upload or link product images",
      "MEDIUM",
    ],
  ];

  for (const issue of issues) {
    for (const [i, val] of issue.entries()) {
      const cell = xlsx.utils.encode_cell({ r: row, c: i });
      ws[cell] = { v: val, t: typeof val === "number" ? "n" : "s" };

      // Color code by impact
      const impact = issue[5] as string;
      if (impact === "CRUCIAL") ws[cell].fill = { fgColor: { rgb: "FF0000" } };
      else if (impact === "HIGH") ws[cell].fill = { fgColor: { rgb: "FFC000" } };
      else if (impact === "MEDIUM") ws[cell].fill = { fgColor: { rgb: "FFFF00" } };
    }
    row++;
  }

  // Set column widths
  ws["!cols"] = [15, 35, 12, 35, 30, 15].map((w) => ({ wch: w }));

  finalizeSheet(ws, row, 6);
  return ws;
}

// ─── SHEET 2: CATALOG ANALYSIS ────────────────────────────────────────────

function createCatalogAnalysisSheet(
  instanceName: string,
  summary: InstanceSummary & { dummy?: string },
  metrics: ProductMetrics[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  // Title
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "CATALOG ANALYSIS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, color: { rgb: "FFFFFF" }, size: 14 };
  row += 2;

  // Catalog Overview
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "1️⃣ CATALOG OVERVIEW", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const overviewHeaders = [
    "Instance",
    "Total Products",
    "Active",
    "Inactive",
    "% Inactive",
    "Total Reviews",
    "Reviews on Inactive",
    "% at Risk",
  ];
  applyHeaderStyle(ws, row, overviewHeaders.length);
  for (const [i, header] of overviewHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const overviewRow = [
    instanceName,
    summary.totalProducts,
    summary.activeProducts,
    summary.inactiveProducts,
    summary.percentInactive,
    summary.totalReviews,
    summary.reviewsOnInactive,
    summary.percentReviewsAtRisk,
  ];

  for (const [i, val] of overviewRow.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
  }
  row += 3;

  // Top Inactive Products with High Reviews
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "2️⃣ TOP INACTIVE PRODUCTS WITH REVIEWS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const topInactive = metrics
    .filter((m) => !m.isActive && m.reviewCount > 0)
    .sort((a, b) => b.reviewCount - a.reviewCount)
    .slice(0, 15);

  const inactiveHeaders = ["Product ID", "Product Name", "Brand", "Total Reviews", "Status"];
  applyHeaderStyle(ws, row, inactiveHeaders.length);
  for (const [i, header] of inactiveHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  for (const m of topInactive) {
    const p = m.product;
    const data = [safeCellValue(p.Id), safeCellValue(p.Name), safeCellValue(p.Brand?.Name ?? ""), safeCellValue(m.reviewCount), safeCellValue("Inactive")];
    for (const [i, val] of data.entries()) {
      ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
    }
    row++;
  }

  row += 2;

  // Missing Crucial Fields
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "3️⃣ MISSING CRUCIAL FIELDS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const missingHeaders = ["Instance", "Missing Images", "Missing Names", "Missing URLs", "Missing Descriptions", "Total Missing"];
  applyHeaderStyle(ws, row, missingHeaders.length);
  for (const [i, header] of missingHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const missingCounts = [
    instanceName,
    summary.missingImageUrl,
    metrics.filter((m) => !m.product.Name || !m.product.Name.trim()).length,
    metrics.filter((m) => !m.hasPageUrl).length,
    summary.missingDescription,
    summary.missingImageUrl + summary.missingDescription + summary.missingUpcEan,
  ];

  for (const [i, val] of missingCounts.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
  }
  row += 3;

  // UPC / EAN Issues
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "4️⃣ UPC / EAN ISSUES", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const upcHeaders = ["Issue Type", "Products Affected"];
  applyHeaderStyle(ws, row, upcHeaders.length);
  for (const [i, header] of upcHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const upcIssues = [["Missing UPC/EAN", summary.missingUpcEan]];

  for (const issue of upcIssues) {
    for (const [i, val] of issue.entries()) {
      ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
    }
    row++;
  }

  ws["!cols"] = [15, 30, 20, 15, 15, 15].map((w) => ({ wch: w }));

  finalizeSheet(ws, row, 6);
  return ws;
}

// ─── SHEET 3: SUGGESTIONS SUMMARY ─────────────────────────────────────────

function createSuggestionsSheet(
  instanceName: string,
  summary: InstanceSummary & { dummy?: string },
  metrics: ProductMetrics[],
  generatedAt: string,
): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  // Title
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "SUGGESTIONS SUMMARY", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, color: { rgb: "FFFFFF" }, size: 14 };
  row += 2;

  // Top Priority Task
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "🔴 TOP PRIORITY TASK", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12, color: { rgb: "FF0000" } };
  row++;

  const inactiveWithReviews = metrics.filter((m) => !m.isActive && m.reviewCount > 0);
  const totalReviewsAtRisk = inactiveWithReviews.reduce((sum, m) => sum + m.reviewCount, 0);

  const priorityData = [
    ["Most Critical Issue", "Inactive Products with Reviews (Reviews not visible)"],
    ["Immediate Action", `Reactivate or update visibility of ${inactiveWithReviews.length} products`],
    ["How to Fix", "Bulk update product status in Bazaarvoice Admin; Re-index catalog"],
  ];

  for (const [label, value] of priorityData) {
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: label, t: "s" };
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "FFE699" } };
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true };

    ws[xlsx.utils.encode_cell({ r: row, c: 1 })] = { v: value, t: "s" };
    row++;
  }

  row += 2;

  // Quick Stats
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "📊 QUICK STATS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const statsData = [
    ["Total Products", summary.totalProducts],
    ["Active Products (%)", `${summary.activeProducts} (${round(((summary.activeProducts / summary.totalProducts) * 100), 1)}%)`],
    ["Inactive Products (%)", `${summary.inactiveProducts} (${summary.percentInactive}%)`],
    ["Total Reviews", summary.totalReviews],
    ["Reviews on Inactive Products (%)", `${summary.reviewsOnInactive} (${summary.percentReviewsAtRisk}%)`],
  ];

  for (const [label, value] of statsData) {
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: label, t: "s" };
    ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true };
    ws[xlsx.utils.encode_cell({ r: row, c: 1 })] = { v: value, t: "s" };
    row++;
  }

  row += 2;

  // Issues & Recommendations
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "📌 ISSUES & RECOMMENDATIONS", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, size: 12 };
  row++;

  const issueHeaders = ["Issue #", "Issue", "Recommendation", "Impact", "Priority"];
  applyHeaderStyle(ws, row, issueHeaders.length);
  for (const [i, header] of issueHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  const suggestions = [
    [1, "Inactive Products with Reviews", `Reactivate ${inactiveWithReviews.length} products`, `${totalReviewsAtRisk} reviews at risk`, "CRUCIAL"],
    [2, "Missing UPC/EAN Codes", `Add codes to ${summary.missingUpcEan} products`, "Inventory issues", "HIGH"],
    [3, "Missing Product Descriptions", `Enrich ${summary.missingDescription} products`, "SEO impact", "MEDIUM"],
    [4, "Missing Product Images", `Add images to ${summary.missingImageUrl} products`, "Conversion impact", "MEDIUM"],
  ];

  for (const suggestion of suggestions) {
    for (const [i, val] of suggestion.entries()) {
      const cell = xlsx.utils.encode_cell({ r: row, c: i });
      ws[cell] = { v: val, t: typeof val === "number" ? "n" : "s" };

      const priority = suggestion[4] as string;
      if (priority === "CRUCIAL") ws[cell].fill = { fgColor: { rgb: "FF0000" } };
      else if (priority === "HIGH") ws[cell].fill = { fgColor: { rgb: "FFC000" } };
      else if (priority === "MEDIUM") ws[cell].fill = { fgColor: { rgb: "FFFF00" } };
    }
    row++;
  }

  ws["!cols"] = [12, 35, 35, 25, 15].map((w) => ({ wch: w }));

  finalizeSheet(ws, row, 5);
  return ws;
}

// ─── SHEET 4: PIVOT TABLE DATA ────────────────────────────────────────────

function createPivotSheet(instanceName: string, metrics: ProductMetrics[]): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  // Title
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "PIVOT DATA - MISSING FIELDS BY INSTANCE & PRODUCT", t: "s" };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].fill = { fgColor: { rgb: "002060" } };
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })].font = { bold: true, color: { rgb: "FFFFFF" }, size: 12 };
  row += 2;

  // Pivot table headers
  const pivotHeaders = [
    "Instance",
    "Product ID",
    "Product Name",
    "Missing Description",
    "Missing Image URL",
    "Missing UPC/EAN",
    "Missing Page URL",
    "Status",
  ];

  applyHeaderStyle(ws, row, pivotHeaders.length);
  for (const [i, header] of pivotHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  // Add product rows with missing field indicators
  for (const m of metrics) {
    const p = m.product;
    const status = m.isActive ? "Active" : "Inactive";

    const data = [
      instanceName,
      p.Id,
      p.Name,
      m.hasDescription ? "" : "YES",
      m.hasImageUrl ? "" : "YES",
      m.hasUpc && m.hasEan ? "" : "YES",
      m.hasPageUrl ? "" : "YES",
      status,
    ];

    for (const [i, val] of data.entries()) {
      const cell = xlsx.utils.encode_cell({ r: row, c: i });
      ws[cell] = { v: val, t: "s" };

      // Highlight missing fields
      if (val === "YES") {
        ws[cell].fill = { fgColor: { rgb: "FF0000" } };
        ws[cell].font = { color: { rgb: "FFFFFF" }, bold: true };
      }
    }
    row++;
  }

  ws["!cols"] = [15, 20, 25, 18, 18, 18, 18, 12].map((w) => ({ wch: w }));

  finalizeSheet(ws, row, 8);
  return ws;
}

// ─── SHEET 5: PRODUCT DATA ────────────────────────────────────────────────

function createProductDataSheet(instanceName: string, metrics: ProductMetrics[], generatedAt: string): xlsx.WorkSheet {
  const ws: xlsx.WorkSheet = {};
  let row = 0;

  // Header
  const instanceDisplay = instanceName.split("-").map((p) => p.charAt(0).toUpperCase() + p.slice(1)).join(" ");

  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "Bazaarvoice" };
  row++;
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: `R&R Key Metrics ${instanceDisplay}` };
  row++;
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: formatRunDate(new Date(generatedAt)) };
  row++;
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: `Instance: ${instanceName}` };
  row++;
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "Date range: All Time" };
  row++;
  ws[xlsx.utils.encode_cell({ r: row, c: 0 })] = { v: "Date field: Review Submission Date" };
  row += 2;

  // Product data table headers
  const dataHeaders = [
    "Brand ID",
    "Brand",
    "Product ID",
    "Product Status",
    "Product Name",
    "Description",
    "Product Page URL",
    "Image URL",
    "UPC",
    "EAN",
    "Product Family",
    "Product Family - Expand",
    "# Approved Reviews",
    "# Family Reviews",
    "Family Average Rating",
  ];

  applyHeaderStyle(ws, row, dataHeaders.length);
  for (const [i, header] of dataHeaders.entries()) {
    ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: header, t: "s" };
  }
  row++;

  // Add product rows
  for (const m of metrics) {
    const p = m.product;
    const status = resolveProductStatus(p);

    const data = [
      safeCellValue(p.BrandExternalId ?? p.Brand?.Id ?? ""),
      safeCellValue(p.Brand?.Name ?? ""),
      safeCellValue(p.Id),
      safeCellValue(status),
      safeCellValue(p.Name),
      safeCellValue(p.Description ?? ""),
      safeCellValue(p.ProductPageUrl ?? ""),
      safeCellValue(p.ImageUrl ?? ""),
      safeCellValue(p.UPCs?.[0] ?? ""),
      safeCellValue(p.EANs?.[0] ?? ""),
      safeCellValue(p.FamilyIds?.[0] ?? ""),
      safeCellValue(""),
      safeCellValue(m.reviewCount),
      safeCellValue(0),
      safeCellValue(""),
    ];

    for (const [i, val] of data.entries()) {
      ws[xlsx.utils.encode_cell({ r: row, c: i })] = { v: val, t: typeof val === "number" ? "n" : "s" };
    }
    row++;
  }

  ws["!cols"] = [15, 12, 15, 12, 30, 50, 30, 30, 15, 15, 15, 15, 18, 18, 18].map((w) => ({ wch: w }));

  finalizeSheet(ws, row, 15);
  return ws;
}
