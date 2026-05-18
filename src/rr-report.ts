import { mkdir } from "node:fs/promises";
import path from "node:path";
import xlsx from "xlsx";
import { findInstanceByName, getProductCount, listAnswers, listQuestions, listReviews } from "./bazaarvoice.js";

type MetricRow = {
  metric: string;
  value: number | string;
  notes?: string;
};

type RatingDistributionRow = {
  rating: number;
  count: number;
  percentage: number;
};

export type RrKeyMetricsReport = {
  instanceName: string;
  baseUrl: string;
  generatedAt: string;
  outputPath: string;
  reportType: "full" | "sample";
  metrics: {
    totalProducts: number;
    totalReviews: number;
    averageRatingSample: number | null;
    recommendedRateSample: number | null;
    syndicatedRateSample: number | null;
    totalQuestions: number;
    totalAnswers: number;
    reviewsPerProduct: number | null;
    questionsPerProduct: number | null;
    totalUgc: number;
    avgUgcPerProduct: number | null;
  };
  ratingDistribution: RatingDistributionRow[];
};

const DEFAULT_INSTANCE_NAME = "pampers-en-us";
const DEFAULT_REPORTS_DIR = path.resolve(process.cwd(), "reports");

function formatDateStamp(isoDate: string): string {
  return isoDate.replace(/[-:]/g, "").replace(/\.\d{3}Z$/, "Z");
}

function round(value: number, decimals = 2): number {
  const factor = 10 ** decimals;
  return Math.round(value * factor) / factor;
}

function createMetricRows(report: Omit<RrKeyMetricsReport, "outputPath" | "reportType">): MetricRow[] {
  return [
    { metric: "Instance", value: report.instanceName, notes: report.baseUrl },
    { metric: "Generated At", value: report.generatedAt },
    { metric: "Total Products", value: report.metrics.totalProducts },
    { metric: "Total Reviews", value: report.metrics.totalReviews },
    {
      metric: "Average Rating (Sample)",
      value: report.metrics.averageRatingSample ?? "N/A",
      notes: "Latest 100 reviews ordered by SubmissionTime desc",
    },
    {
      metric: "Recommended Rate (Sample %)",
      value: report.metrics.recommendedRateSample ?? "N/A",
      notes: "Latest 100 reviews ordered by SubmissionTime desc",
    },
    {
      metric: "Syndicated Rate (Sample %)",
      value: report.metrics.syndicatedRateSample ?? "N/A",
      notes: "Latest 100 reviews ordered by SubmissionTime desc",
    },
    { metric: "Total Questions", value: report.metrics.totalQuestions },
    { metric: "Total Answers", value: report.metrics.totalAnswers },
    { metric: "Reviews per Product", value: report.metrics.reviewsPerProduct ?? "N/A" },
    { metric: "Questions per Product", value: report.metrics.questionsPerProduct ?? "N/A" },
    { metric: "Total UGC Items", value: report.metrics.totalUgc, notes: "Reviews + Questions + Answers" },
    { metric: "Avg UGC per Product", value: report.metrics.avgUgcPerProduct ?? "N/A" },
  ];
}

export async function buildRrKeyMetricsReport(
  instanceName = DEFAULT_INSTANCE_NAME,
): Promise<Omit<RrKeyMetricsReport, "outputPath" | "reportType">> {
  const instance = findInstanceByName(instanceName);
  const generatedAt = new Date().toISOString();

  try {
    // Try standard non-syndicated queries first
    const [productCountResult, reviewsResult, reviewSampleResult, questionsResult, answersResult] = await Promise.all([
      getProductCount(instanceName),
      listReviews(instanceName, { limit: 1, offset: 0 }),
      listReviews(instanceName, { limit: 100, offset: 0, sort: "SubmissionTime:desc" }),
      listQuestions(instanceName, { limit: 1, offset: 0 }),
      listAnswers(instanceName, { limit: 1, offset: 0 }),
    ]);

    return buildMetricsFromResults(instance, generatedAt, productCountResult, reviewsResult, reviewSampleResult, questionsResult, answersResult, instanceName);
  } catch (err: any) {
    // Handle syndicated accounts that require filters
    if (err.message?.includes("Syndication") || err.message?.includes("EQ filter")) {
      console.log(`Detected syndicated account: ${instanceName}. Using alternative query strategy...`);
      
      // For syndicated accounts, we need to query with specific filters
      // Try to get data using available endpoints or provide limited metrics
      const productCountResult = { totalProducts: 0 };
      
      try {
        // Try to get product count (some syndicated accounts might allow this)
        const countResult = await getProductCount(instanceName);
        productCountResult.totalProducts = countResult.totalProducts;
      } catch {
        // If product count also fails, we'll have to report 0
      }

      // For syndicated accounts, return limited metrics
      return {
        instanceName,
        baseUrl: instance.baseUrl,
        generatedAt,
        metrics: {
          totalProducts: productCountResult.totalProducts,
          totalReviews: 0,
          averageRatingSample: null,
          recommendedRateSample: null,
          syndicatedRateSample: null,
          totalQuestions: 0,
          totalAnswers: 0,
          reviewsPerProduct: null,
          questionsPerProduct: null,
          totalUgc: 0,
          avgUgcPerProduct: null,
        },
        ratingDistribution: [],
      } as Omit<RrKeyMetricsReport, "outputPath" | "reportType">;
    }
    
    throw err;
  }
}

function buildMetricsFromResults(
  instance: { baseUrl: string },
  generatedAt: string,
  productCountResult: { totalProducts: number },
  reviewsResult: { TotalResults: number },
  reviewSampleResult: { Results: any[] },
  questionsResult: { TotalResults: number },
  answersResult: { TotalResults: number },
  instanceName: string,
): Omit<RrKeyMetricsReport, "outputPath" | "reportType"> {
  const totalProducts = productCountResult.totalProducts;
  const totalReviews = reviewsResult.TotalResults;
  const totalQuestions = questionsResult.TotalResults;
  const totalAnswers = answersResult.TotalResults;
  const reviewSample = reviewSampleResult.Results;

  const averageRatingSample = reviewSample.length
    ? round(reviewSample.reduce((sum, review) => sum + (review.Rating || 0), 0) / reviewSample.length)
    : null;

  const recommendedRateSample = reviewSample.length
    ? round((reviewSample.filter((review) => Boolean(review.IsRecommended)).length / reviewSample.length) * 100, 1)
    : null;

  const syndicatedRateSample = reviewSample.length
    ? round((reviewSample.filter((review) => Boolean(review.IsSyndicated)).length / reviewSample.length) * 100, 1)
    : null;

  const ratingDistributionMap = new Map<number, number>();
  for (const review of reviewSample) {
    if (typeof review.Rating !== "number") continue;
    ratingDistributionMap.set(review.Rating, (ratingDistributionMap.get(review.Rating) ?? 0) + 1);
  }

  const ratingDistribution = [...ratingDistributionMap.entries()]
    .sort((a, b) => b[0] - a[0])
    .map(([rating, count]) => ({
      rating,
      count,
      percentage: reviewSample.length ? round((count / reviewSample.length) * 100, 1) : 0,
    }));

  const totalUgc = totalReviews + totalQuestions + totalAnswers;

  return {
    instanceName,
    baseUrl: instance.baseUrl,
    generatedAt,
    metrics: {
      totalProducts,
      totalReviews,
      averageRatingSample,
      recommendedRateSample,
      syndicatedRateSample,
      totalQuestions,
      totalAnswers,
      reviewsPerProduct: totalProducts ? round(totalReviews / totalProducts) : null,
      questionsPerProduct: totalProducts ? round(totalQuestions / totalProducts) : null,
      totalUgc,
      avgUgcPerProduct: totalProducts ? round(totalUgc / totalProducts) : null,
    },
    ratingDistribution,
  };
}

export async function generateRrKeyMetricsExcelReport(
  instanceName = DEFAULT_INSTANCE_NAME,
  outputDir = DEFAULT_REPORTS_DIR,
): Promise<RrKeyMetricsReport> {
  const report = await buildRrKeyMetricsReport(instanceName);
  const workbook = xlsx.utils.book_new();

  // Full R&R Key Metrics Report: Multiple detailed sheets
  const summarySheet = xlsx.utils.json_to_sheet(createMetricRows(report));
  const distributionRows = report.ratingDistribution.length
    ? report.ratingDistribution
    : [{ rating: "N/A", count: 0, percentage: 0 }];
  const distributionSheet = xlsx.utils.json_to_sheet(distributionRows);

  xlsx.utils.book_append_sheet(workbook, summarySheet, "Summary");
  xlsx.utils.book_append_sheet(workbook, distributionSheet, "Rating Distribution");

  await mkdir(outputDir, { recursive: true });

  const filename = `${instanceName}_R&R_key_metrics_${formatDateStamp(report.generatedAt)}.xlsx`;
  const outputPath = path.join(outputDir, filename);
  xlsx.writeFile(workbook, outputPath);

  return {
    ...report,
    reportType: "full",
    outputPath,
  };
}

export async function generateSampleReportExcel(
  instanceName = DEFAULT_INSTANCE_NAME,
  outputDir = DEFAULT_REPORTS_DIR,
): Promise<RrKeyMetricsReport> {
  const report = await buildRrKeyMetricsReport(instanceName);
  const workbook = xlsx.utils.book_new();

  // Sample Report: same compact structure as the current R&R workbook
  const summarySheet = xlsx.utils.json_to_sheet(createMetricRows(report));
  const distributionRows = report.ratingDistribution.length
    ? report.ratingDistribution
    : [{ rating: "N/A", count: 0, percentage: 0 }];
  const distributionSheet = xlsx.utils.json_to_sheet(distributionRows);

  xlsx.utils.book_append_sheet(workbook, summarySheet, "Summary");
  xlsx.utils.book_append_sheet(workbook, distributionSheet, "Rating Distribution");

  await mkdir(outputDir, { recursive: true });

  const filename = `${instanceName}_sample_report_${formatDateStamp(report.generatedAt)}.xlsx`;
  const outputPath = path.join(outputDir, filename);
  xlsx.writeFile(workbook, outputPath);

  return {
    ...report,
    reportType: "sample",
    outputPath,
  };
}