import "dotenv/config";

// Usage:  node scripts/generate-rr-report.mjs [instanceName]
// If no instance name is passed, the first configured instance is used.

const instances = JSON.parse(process.env.BV_INSTANCES_JSON ?? "[]");
if (instances.length === 0) {
  throw new Error("No instances found in BV_INSTANCES_JSON");
}

const targetName = process.argv[2]?.trim();
const instance   = targetName
  ? instances.find((item) => item.name === targetName)
  : instances[0];

if (!instance) {
  const names = instances.map((i) => i.name).join(", ");
  throw new Error(
    `Instance "${targetName}" not found in BV_INSTANCES_JSON. Available: ${names}`,
  );
}

const timeoutMs = Number(process.env.BV_TIMEOUT_MS ?? "12000");

async function bvFetch(endpoint, params) {
  const url = new URL(endpoint, instance.baseUrl);
  url.searchParams.set("passkey", instance.passkey);
  url.searchParams.set("apiversion", instance.apiversion ?? "5.4");

  for (const [key, value] of Object.entries(params)) {
    url.searchParams.set(key, value);
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url, {
      method: "GET",
      signal: controller.signal,
      headers: { Accept: "application/json" },
    });

    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    return await response.json();
  } finally {
    clearTimeout(timer);
  }
}

console.log(
  "\n╔════════════════════════════════════════════════════════════════════════════╗",
);
console.log("║                    BAZAARVOICE R&R KEY METRICS REPORT                      ║");
console.log(
  "╚════════════════════════════════════════════════════════════════════════════╝\n",
);
console.log(`Instance: ${instance.name} (${instance.baseUrl})`);
console.log(`Generated: ${new Date().toISOString()}\n`);

// ─── Product Overview ───

const productCountResult = await bvFetch("/data/products.json", { Limit: "1", Offset: "0" });
const totalProducts = productCountResult.TotalResults;
console.log("📊 PRODUCT OVERVIEW");
console.log("───────────────────────────────────────────────────────────────────────────");
console.log(`Total Products:              ${totalProducts}`);

// ─── Review Metrics ───

console.log("\n💬 REVIEW METRICS");
console.log("───────────────────────────────────────────────────────────────────────────");

const reviewsResult = await bvFetch("/data/reviews.json", { Limit: "1", Offset: "0" });
const totalReviews = reviewsResult.TotalResults;
console.log(`Total Reviews:               ${totalReviews}`);

if (totalReviews === 0) {
  console.log("  (No reviews data available)");
} else {
  // Get sample of reviews to calculate engagement
  const reviewSampleResult = await bvFetch("/data/reviews.json", {
    Limit: "100",
    Offset: "0",
    Sort: "SubmissionTime:desc",
  });

  if (reviewSampleResult.Results.length > 0) {
    const avgRating =
      reviewSampleResult.Results.reduce((sum, r) => sum + (r.Rating || 0), 0) /
      reviewSampleResult.Results.length;
    console.log(`Average Rating (sample):     ${avgRating.toFixed(2)} / 5.0`);

    const recommendedCount = reviewSampleResult.Results.filter((r) => r.IsRecommended).length;
    const recommendedPct = ((recommendedCount / reviewSampleResult.Results.length) * 100).toFixed(1);
    console.log(`Recommended Rate (sample):   ${recommendedCount}/${reviewSampleResult.Results.length} (${recommendedPct}%)`);

    const syndicatedCount = reviewSampleResult.Results.filter((r) => r.IsSyndicated).length;
    console.log(`Syndicated Reviews (sample): ${syndicatedCount}/${reviewSampleResult.Results.length}`);
  }
}

// ─── Rating Distribution ───

console.log("\n⭐ RATING DISTRIBUTION");
console.log("───────────────────────────────────────────────────────────────────────────");

// Get statistics for a sample product to show rating breakdown
const statsResult = await bvFetch("/data/statistics.json", { Limit: "100", Offset: "0" });
if (statsResult.Results && statsResult.Results.length > 0) {
  const firstStats = statsResult.Results[0]?.ProductStatistics?.ReviewStatistics;
  if (firstStats?.RatingsDistribution) {
    const totalRatings = firstStats.RatingsDistribution.reduce((sum, r) => sum + r.Count, 0);
    console.log("Rating breakdown (sample products):");
    for (const dist of firstStats.RatingsDistribution) {
      const pct = ((dist.Count / totalRatings) * 100).toFixed(1);
      const bar = "█".repeat(Math.round((dist.Count / totalRatings) * 20));
      console.log(
        `  ${dist.RatingValue}★ ${bar.padEnd(20, "░")} ${dist.Count} (${pct}%)`,
      );
    }
  }
}

// ─── Q&A Metrics ───

console.log("\n❓ Q&A METRICS");
console.log("───────────────────────────────────────────────────────────────────────────");

const questionsResult = await bvFetch("/data/questions.json", { Limit: "1", Offset: "0" });
const totalQuestions = questionsResult.TotalResults;
console.log(`Total Questions:             ${totalQuestions}`);

const answersResult = await bvFetch("/data/answers.json", { Limit: "1", Offset: "0" });
const totalAnswers = answersResult.TotalResults;
console.log(`Total Answers:               ${totalAnswers}`);

if (totalQuestions > 0) {
  const avgAnswersPerQuestion = (totalAnswers / totalQuestions).toFixed(2);
  console.log(`Avg Answers per Question:    ${avgAnswersPerQuestion}`);
}

// ─── Engagement Summary ───

console.log("\n📈 ENGAGEMENT SUMMARY");
console.log("───────────────────────────────────────────────────────────────────────────");

const reviewsPerProduct = totalProducts > 0 ? (totalReviews / totalProducts).toFixed(2) : "N/A";
const questionsPerProduct = totalProducts > 0 ? (totalQuestions / totalProducts).toFixed(2) : "N/A";
const totalUGC = totalReviews + totalQuestions + totalAnswers;

console.log(`Reviews per Product:         ${reviewsPerProduct}`);
console.log(`Questions per Product:       ${questionsPerProduct}`);
console.log(`Total UGC Items (Reviews+Q&A): ${totalUGC}`);

if (totalProducts > 0) {
  const ugcPerProduct = (totalUGC / totalProducts).toFixed(2);
  console.log(`Avg UGC per Product:         ${ugcPerProduct}`);
}

console.log("\n" + "═".repeat(76) + "\n");
