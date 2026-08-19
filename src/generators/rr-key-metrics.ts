import { mkdir } from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import { listProducts } from "../api/bazaarvoice.js";
import { resolveProductStatus } from "../utils/product-status.js";
import type { Product } from "../types/index.js";

const DEFAULT_REPORTS_DIR = path.resolve(process.cwd(), "reports");

type RrKeyMetricsTemplateReport = {
  instanceNames: string[];
  outputPath: string;
  totalProducts: number;
  generatedAt: string;
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

    const totalFetched = offset + page.Results.length;
    if (totalFetched >= page.TotalResults || page.Results.length < pageSize) {
      break;
    }
  }

  return allProducts;
}

function buildSheetRows(instanceName: string, generatedAt: string, products: Product[]): (string | number)[][] {
  const instanceDisplayName = instanceName
    .split("-")
    .map((part) => part.charAt(0).toUpperCase() + part.slice(1).toLowerCase())
    .join(" ");

  const rows: (string | number)[][] = [
    ["Bazaarvoice"],
    [`R&R Key Metrics ${instanceDisplayName}`],
    [formatRunDate(new Date(generatedAt))],
    [`Instance: ${instanceName}`],
    ["Date range: All Time"],
    ["Date field: Review Submission Date"],
    [
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
    ],
  ];

  for (const product of products) {
    const reviewCount = product.ReviewStatistics?.TotalReviewCount ?? 0;
    const familyAvgRating = product.ReviewStatistics?.AverageOverallRating;

    rows.push([
      product.BrandExternalId ?? product.Brand?.Id ?? "",
      product.Brand?.Name ?? "",
      product.Id,
      resolveProductStatus(product),
      product.Name ?? "",
      product.Description ?? "",
      product.ProductPageUrl ?? "",
      product.ImageUrl ?? "",
      product.UPCs?.[0] ?? "",
      product.EANs?.[0] ?? "",
      product.FamilyIds?.[0] ?? "",
      "",
      reviewCount,
      reviewCount,
      typeof familyAvgRating === "number" ? round(familyAvgRating) : "",
    ]);
  }

  return rows;
}

export async function generateRrKeyMetricsTemplateReport(
  instanceInput: string | string[],
  outputDir = DEFAULT_REPORTS_DIR,
): Promise<RrKeyMetricsTemplateReport> {
  const instanceNames = Array.isArray(instanceInput) ? [...new Set(instanceInput)] : [instanceInput];
  const generatedAt   = new Date().toISOString();

  const workbook = xlsx.utils.book_new();
  let totalProducts = 0;

  for (const instanceName of instanceNames) {
    const products = await listAllProducts(instanceName);
    totalProducts += products.length;
    const rows = buildSheetRows(instanceName, generatedAt, products);
    const sheet = xlsx.utils.aoa_to_sheet(rows);
    sheet["!cols"] = [15, 12, 15, 12, 30, 50, 30, 30, 15, 15, 15, 20, 18, 18, 18].map((wch) => ({ wch }));
    const sheetName = instanceName.replace(/[^a-zA-Z0-9-_ ]/g, "_").slice(0, 31) || "R&R_Key_Metrics";
    xlsx.utils.book_append_sheet(workbook, sheet, sheetName);
  }

  await mkdir(outputDir, { recursive: true });

  const label      = instanceNames.length === 1 ? instanceNames[0] : `multi-instance_${instanceNames.length}instances`;
  const outputPath = path.join(outputDir, `${label}_R&R_key_metrics_${formatTimestampForFilename(new Date(generatedAt))}.xlsx`);
  xlsx.writeFile(workbook, outputPath);

  return {
    instanceNames,
    outputPath,
    totalProducts,
    generatedAt,
  };
}
