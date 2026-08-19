import "dotenv/config";
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { CallToolRequestSchema, ListToolsRequestSchema } from "@modelcontextprotocol/sdk/types.js";
import { z } from "zod";
import {
  countInstancesByOwner,
  getConfiguredInstances,
  getProductCount,
  getProductStatistics,
  listAnswers,
  listProducts,
  listQuestions,
  listReviews,
} from "../api/bazaarvoice.js";
import { generateRrKeyMetricsTemplateReport } from "../generators/rr-key-metrics.js";
import { generateSampleReportExcel } from "../generators/rr-sample.js";
import { generateInstanceAuditReport } from "../generators/audit.js";

// ─── Server definition ─────────────────────────────────────────────────────────

const server = new Server(
  { name: "bv-mcp-server", version: "0.2.0" },
  { capabilities: { tools: {} } },
);

// ─── Input schemas (Zod) ──────────────────────────────────────────────────────

const instanceNameSchema = z.object({
  instanceName: z.string().min(1),
});

const paginationFields = {
  limit: z.number().int().min(1).max(100).default(10).optional(),
  offset: z.number().int().min(0).default(0).optional(),
};

const productCountSchema = instanceNameSchema;

const listProductsSchema = instanceNameSchema.extend({
  ...paginationFields,
  search: z.string().optional(),
  sort: z.string().optional(),
  includeStats: z.boolean().default(false).optional(),
});

const productStatisticsSchema = instanceNameSchema.extend({
  productId: z.string().min(1),
});

const listReviewsSchema = instanceNameSchema.extend({
  ...paginationFields,
  productId: z.string().optional(),
  sort: z.string().optional(),
  search: z.string().optional(),
});

const listQuestionsSchema = instanceNameSchema.extend({
  ...paginationFields,
  productId: z.string().optional(),
  sort: z.string().optional(),
  includeAnswers: z.boolean().default(false).optional(),
});

const listAnswersSchema = instanceNameSchema.extend({
  ...paginationFields,
  questionId: z.string().optional(),
  sort: z.string().optional(),
});

const countByOwnerSchema = z.object({
  owner: z.string().min(1),
});

/**
 * Resolves the list of instance names from a tool call.
 * Accepts either a single `instanceName` string or an `instanceNames` array.
 * Falls back to all configured instances when neither is provided.
 */
function resolveInstanceNamesFromArgs(
  args: Record<string, unknown>,
): string[] {
  if (Array.isArray(args.instanceNames) && (args.instanceNames as unknown[]).length > 0) {
    return (args.instanceNames as unknown[]).map(String);
  }
  if (typeof args.instanceName === "string" && args.instanceName.trim()) {
    return [args.instanceName.trim()];
  }
  // No instance specified — run against all configured instances.
  return getConfiguredInstances().map((i) => i.name);
}

const instanceNamesInputSchema = {
  type: "object" as const,
  properties: {
    instanceName: {
      type: "string",
      description: "Single configured instance name (e.g. pampers-en-us). Omit to run against all instances.",
    },
    instanceNames: {
      type: "array",
      items: { type: "string" },
      description: "Array of configured instance names to combine into one report.",
    },
  },
  additionalProperties: false,
} as const;

// ─── Tool registry ────────────────────────────────────────────────────────────

server.setRequestHandler(ListToolsRequestSchema, async () => ({
  tools: [
    // — Instance discovery
    {
      name: "bv_list_instances",
      description:
        "List all Bazaarvoice instances configured in BV_INSTANCES_JSON. No network call is made.",
      inputSchema: { type: "object", properties: {}, additionalProperties: false },
    },
    {
      name: "bv_count_instances_by_owner",
      description:
        "Count and list configured Bazaarvoice instances filtered by the owner field (e.g. 'P&G').",
      inputSchema: {
        type: "object",
        properties: {
          owner: { type: "string", description: "Owner name to match, e.g. P&G" },
        },
        required: ["owner"],
        additionalProperties: false,
      },
    },

    // — Products
    {
      name: "bv_get_product_count",
      description: "Return the total number of products in a Bazaarvoice instance via the Conversations API.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string", description: "Configured instance name" },
        },
        required: ["instanceName"],
        additionalProperties: false,
      },
    },
    {
      name: "bv_list_products",
      description:
        "List products from a Bazaarvoice instance using the Conversations API (/data/products.json). Supports pagination, text search, and optional review statistics.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string" },
          limit: { type: "number", description: "Results per page (1–100, default 10)" },
          offset: { type: "number", description: "Zero-based page offset (default 0)" },
          search: { type: "string", description: "Full-text search across product names" },
          sort: { type: "string", description: "Sort expression e.g. 'Name:asc'" },
          includeStats: {
            type: "boolean",
            description: "Include ReviewStatistics (avg rating, count) inline with each product",
          },
        },
        required: ["instanceName"],
        additionalProperties: false,
      },
    },

    // — Statistics
    {
      name: "bv_get_product_statistics",
      description:
        "Fetch aggregate review statistics (average rating, total review count, rating distribution) for a specific product using the Conversations API /data/statistics.json endpoint.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string" },
          productId: { type: "string", description: "The Bazaarvoice product identifier" },
        },
        required: ["instanceName", "productId"],
        additionalProperties: false,
      },
    },

    // — Reviews
    {
      name: "bv_list_reviews",
      description:
        "Retrieve user reviews from a Bazaarvoice instance via /data/reviews.json. Filter by product, paginate, sort (e.g. 'Rating:desc'), or full-text search.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string" },
          productId: { type: "string", description: "Filter reviews to a specific product" },
          limit: { type: "number", description: "Results per page (1–100, default 10)" },
          offset: { type: "number", description: "Zero-based page offset" },
          sort: {
            type: "string",
            description: "Sort expression e.g. 'Rating:desc', 'SubmissionTime:desc' (default)",
          },
          search: { type: "string", description: "Full-text keyword search across review title and body" },
        },
        required: ["instanceName"],
        additionalProperties: false,
      },
    },

    // — Questions
    {
      name: "bv_list_questions",
      description:
        "Retrieve Q&A questions from a Bazaarvoice instance via /data/questions.json. Filter by product and optionally inline the first page of answers.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string" },
          productId: { type: "string", description: "Filter questions to a specific product" },
          limit: { type: "number", description: "Results per page (1–100, default 10)" },
          offset: { type: "number", description: "Zero-based page offset" },
          sort: { type: "string", description: "Sort expression e.g. 'SubmissionTime:desc' (default)" },
          includeAnswers: { type: "boolean", description: "Inline answers in the response via Include=Answers" },
        },
        required: ["instanceName"],
        additionalProperties: false,
      },
    },

    // — Answers
    {
      name: "bv_list_answers",
      description:
        "Retrieve answers from a Bazaarvoice instance via /data/answers.json. Filter by question ID.",
      inputSchema: {
        type: "object",
        properties: {
          instanceName: { type: "string" },
          questionId: { type: "string", description: "Filter answers to a specific question" },
          limit: { type: "number", description: "Results per page (1–100, default 10)" },
          offset: { type: "number", description: "Zero-based page offset" },
          sort: { type: "string", description: "Sort expression e.g. 'SubmissionTime:desc' (default)" },
        },
        required: ["instanceName"],
        additionalProperties: false,
      },
    },
    {
      name: "bv_generate_rr_key_metrics_report",
      description:
        "Generate an Excel R&R Key Metrics report for one or more Bazaarvoice instances. " +
        "Each instance becomes its own sheet. Pass instanceName for a single instance, " +
        "instanceNames array to combine multiple, or omit to include all configured instances.",
      inputSchema: instanceNamesInputSchema,
    },
    {
      name: "bv_generate_sample_report",
      description:
        "Generate an Excel sample R&R report (Summary + Rating Distribution) for one or more instances. " +
        "Pass instanceName, instanceNames array, or omit to include all configured instances.",
      inputSchema: instanceNamesInputSchema,
    },
    {
      name: "bv_generate_instance_audit_report",
      description:
        "Generate a comprehensive catalog audit Excel workbook for one or more Bazaarvoice instances. " +
        "Produces a single 'Audit Report' sheet (9 sections) plus one product-data sheet per instance. " +
        "Pass instanceName, instanceNames array, or omit to audit all configured instances.",
      inputSchema: instanceNamesInputSchema,
    },
  ],
}));

// ─── Tool handlers ─────────────────────────────────────────────────────────────

function ok(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
}

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  try {
    const args = request.params.arguments ?? {};

    switch (request.params.name) {
      // — Instance discovery
      case "bv_list_instances": {
        const instances = getConfiguredInstances().map((item) => ({
          name: item.name,
          baseUrl: item.baseUrl,
          apiversion: item.apiversion ?? "5.4",
          owner: item.owner ?? null,
        }));

        return ok({ count: instances.length, instances });
      }

      case "bv_count_instances_by_owner": {
        const { owner } = countByOwnerSchema.parse(args);
        return ok(countInstancesByOwner(owner));
      }

      // — Products
      case "bv_get_product_count": {
        const { instanceName } = productCountSchema.parse(args);
        return ok(await getProductCount(instanceName));
      }

      case "bv_list_products": {
        const { instanceName, ...opts } = listProductsSchema.parse(args);
        const result = await listProducts(instanceName, opts);

        return ok({
          totalResults: result.TotalResults,
          limit: result.Limit,
          offset: result.Offset,
          products: result.Results,
        });
      }

      // — Statistics
      case "bv_get_product_statistics": {
        const { instanceName, productId } = productStatisticsSchema.parse(args);
        const result = await getProductStatistics(instanceName, productId);
        const stats = result.Results[0]?.ProductStatistics;

        if (!stats) {
          return ok({ productId, message: "No statistics found for this product." });
        }

        return ok({
          productId: stats.ProductId,
          reviewStatistics: stats.ReviewStatistics,
          nativeReviewStatistics: stats.NativeReviewStatistics ?? null,
        });
      }

      // — Reviews
      case "bv_list_reviews": {
        const { instanceName, ...opts } = listReviewsSchema.parse(args);
        const result = await listReviews(instanceName, opts);

        return ok({
          totalResults: result.TotalResults,
          limit: result.Limit,
          offset: result.Offset,
          reviews: result.Results,
        });
      }

      // — Questions
      case "bv_list_questions": {
        const { instanceName, ...opts } = listQuestionsSchema.parse(args);
        const result = await listQuestions(instanceName, opts);

        return ok({
          totalResults: result.TotalResults,
          limit: result.Limit,
          offset: result.Offset,
          questions: result.Results,
          includes: result.Includes ?? null,
        });
      }

      // — Answers
      case "bv_list_answers": {
        const { instanceName, ...opts } = listAnswersSchema.parse(args);
        const result = await listAnswers(instanceName, opts);

        return ok({
          totalResults: result.TotalResults,
          limit: result.Limit,
          offset: result.Offset,
          answers: result.Results,
        });
      }

      case "bv_generate_rr_key_metrics_report": {
        const names = resolveInstanceNamesFromArgs(args);
        return ok(await generateRrKeyMetricsTemplateReport(names));
      }

      case "bv_generate_sample_report": {
        // Sample report runs per-instance and returns an array of results
        const names   = resolveInstanceNamesFromArgs(args);
        const results = await Promise.all(names.map((n) => generateSampleReportExcel(n)));
        return ok(results.length === 1 ? results[0] : results);
      }

      case "bv_generate_instance_audit_report": {
        const names = resolveInstanceNamesFromArgs(args);
        return ok(await generateInstanceAuditReport({ instanceNames: names }));
      }

      default:
        throw new Error(`Unknown tool: ${request.params.name}`);
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : "Unknown error";
    return { isError: true, content: [{ type: "text" as const, text: message }] };
  }
});

// ─── Start ───────────────────────────────────────────────────────────────────

async function main() {
  console.error("[bv-mcp-server] Starting MCP server...");

  const transport = new StdioServerTransport();
  console.error("[bv-mcp-server] Connecting to stdio transport...");

  await server.connect(transport);

  // MCP servers are long-running and wait for client requests over stdio.
  console.error("[bv-mcp-server] Ready. Waiting for MCP client requests on stdio.");
}

main().catch((error) => {
  const message = error instanceof Error ? error.message : String(error);
  console.error(`MCP server failed to start: ${message}`);
  process.exit(1);
});
