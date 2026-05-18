import { z } from "zod";
import type {
  Answer,
  BvInstance,
  ConversationsResponse,
  Product,
  Question,
  StatisticsResult,
  Review,
} from "./types.js";

// ─── Instance config ─────────────────────────────────────────────────────────

const instancesSchema = z.array(
  z.object({
    name: z.string().min(1),
    baseUrl: z.string().url(),
    passkey: z.string().min(1),
    apiversion: z.string().default("5.4").optional(),
    owner: z.string().optional(),
  }),
);

function parseInstancesFromEnv(): BvInstance[] {
  const raw = process.env.BV_INSTANCES_JSON;

  if (!raw) {
    throw new Error("Missing BV_INSTANCES_JSON. Add it to your .env file.");
  }

  let parsedUnknown: unknown;
  try {
    parsedUnknown = JSON.parse(raw);
  } catch {
    throw new Error("BV_INSTANCES_JSON is not valid JSON.");
  }

  const parsed = instancesSchema.safeParse(parsedUnknown);
  if (!parsed.success) {
    throw new Error(`BV_INSTANCES_JSON schema error: ${parsed.error.message}`);
  }

  return parsed.data;
}

export function getConfiguredInstances(): BvInstance[] {
  return parseInstancesFromEnv();
}

export function findInstanceByName(instanceName: string): BvInstance {
  const instance = getConfiguredInstances().find((item) => item.name === instanceName);

  if (!instance) {
    throw new Error(`No Bazaarvoice instance named '${instanceName}' was found in BV_INSTANCES_JSON.`);
  }

  return instance;
}

export function countInstancesByOwner(owner: string): { owner: string; count: number; instances: string[] } {
  const normalizedOwner = owner.trim().toLowerCase();

  const matched = getConfiguredInstances().filter(
    (item) => (item.owner ?? "").trim().toLowerCase() === normalizedOwner,
  );

  return {
    owner,
    count: matched.length,
    instances: matched.map((item) => item.name),
  };
}

// ─── Central HTTP helper ─────────────────────────────────────────────────────────

/**
 * Calls a Bazaarvoice Conversations API endpoint and returns the parsed JSON.
 * Enforces the timeout configured via BV_TIMEOUT_MS.
 * Throws an Error if the HTTP status is not OK or if the API reports errors.
 */
async function bvFetch<T>(
  instance: BvInstance,
  endpoint: string,
  params: Record<string, string>,
): Promise<ConversationsResponse<T>> {
  const timeoutMs = Number(process.env.BV_TIMEOUT_MS ?? "12000");

  const url = new URL(endpoint, instance.baseUrl);
  url.searchParams.set("passkey", instance.passkey);
  url.searchParams.set("apiversion", instance.apiversion ?? "5.4");

  for (const [key, value] of Object.entries(params)) {
    url.searchParams.set(key, value);
  }

  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const response = await fetch(url.toString(), {
      method: "GET",
      signal: controller.signal,
      headers: { Accept: "application/json" },
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`Bazaarvoice API returned ${response.status}: ${body}`);
    }

    const json = (await response.json()) as ConversationsResponse<T>;

    if (json.HasErrors && json.Errors?.length) {
      const messages = json.Errors.map((e) => `[${e.Code}] ${e.Message}`).join("; ");
      throw new Error(`Bazaarvoice API error: ${messages}`);
    }

    return json;
  } finally {
    clearTimeout(timer);
  }
}

// ─── Products ──────────────────────────────────────────────────────────────────

export async function getProductCount(instanceName: string): Promise<{
  instanceName: string;
  totalProducts: number;
}> {
  const instance = findInstanceByName(instanceName);
  const result = await bvFetch<Product>(instance, "/data/products.json", { Limit: "1", Offset: "0" });

  return { instanceName, totalProducts: result.TotalResults };
}

export async function listProducts(
  instanceName: string,
  opts: {
    limit?: number;
    offset?: number;
    search?: string;
    sort?: string;
    includeStats?: boolean;
  } = {},
): Promise<ConversationsResponse<Product>> {
  const instance = findInstanceByName(instanceName);

  const params: Record<string, string> = {
    Limit: String(Math.min(opts.limit ?? 10, 100)),
    Offset: String(opts.offset ?? 0),
  };

  if (opts.search) params["Search"] = opts.search;
  if (opts.sort) params["Sort"] = opts.sort;
  if (opts.includeStats) params["Stats"] = "Reviews";

  return bvFetch<Product>(instance, "/data/products.json", params);
}

// ─── Statistics ───────────────────────────────────────────────────────────────

/**
 * Fetches aggregate review statistics (avg rating, count, distribution) for a product.
 * Uses the /data/statistics.json endpoint with Stats=Reviews,NativeReviews.
 */
export async function getProductStatistics(
  instanceName: string,
  productId: string,
): Promise<ConversationsResponse<StatisticsResult>> {
  const instance = findInstanceByName(instanceName);

  return bvFetch<StatisticsResult>(instance, "/data/statistics.json", {
    Filter: `ProductId:eq:${productId}`,
    Stats: "Reviews,NativeReviews",
  });
}

// ─── Reviews ──────────────────────────────────────────────────────────────────

export type ListReviewsOptions = {
  productId?: string;
  limit?: number;
  offset?: number;
  /** e.g. "Rating:desc", "SubmissionTime:desc" */
  sort?: string;
  /** Full-text keyword search across review title and body */
  search?: string;
  /** Extra Conversations API Filter expressions appended with AND */
  extraFilters?: string[];
};

export async function listReviews(
  instanceName: string,
  opts: ListReviewsOptions = {},
): Promise<ConversationsResponse<Review>> {
  const instance = findInstanceByName(instanceName);

  const params: Record<string, string> = {
    Limit: String(Math.min(opts.limit ?? 10, 100)),
    Offset: String(opts.offset ?? 0),
    Sort: opts.sort ?? "SubmissionTime:desc",
  };

  const filters: string[] = [];
  if (opts.productId) filters.push(`ProductId:eq:${opts.productId}`);
  if (opts.extraFilters?.length) filters.push(...opts.extraFilters);
  if (filters.length) params["Filter"] = filters.join("&Filter=");

  if (opts.search) params["Search"] = opts.search;

  return bvFetch<Review>(instance, "/data/reviews.json", params);
}

// ─── Questions ───────────────────────────────────────────────────────────────

export async function listQuestions(
  instanceName: string,
  opts: {
    productId?: string;
    limit?: number;
    offset?: number;
    sort?: string;
    /** Pass true to inline the first page of answers in the response via Include=Answers */
    includeAnswers?: boolean;
  } = {},
): Promise<ConversationsResponse<Question>> {
  const instance = findInstanceByName(instanceName);

  const params: Record<string, string> = {
    Limit: String(Math.min(opts.limit ?? 10, 100)),
    Offset: String(opts.offset ?? 0),
    Sort: opts.sort ?? "SubmissionTime:desc",
  };

  if (opts.productId) params["Filter"] = `ProductId:eq:${opts.productId}`;
  if (opts.includeAnswers) params["Include"] = "Answers";

  return bvFetch<Question>(instance, "/data/questions.json", params);
}

// ─── Answers ──────────────────────────────────────────────────────────────────

export async function listAnswers(
  instanceName: string,
  opts: {
    questionId?: string;
    limit?: number;
    offset?: number;
    sort?: string;
  } = {},
): Promise<ConversationsResponse<Answer>> {
  const instance = findInstanceByName(instanceName);

  const params: Record<string, string> = {
    Limit: String(Math.min(opts.limit ?? 10, 100)),
    Offset: String(opts.offset ?? 0),
    Sort: opts.sort ?? "SubmissionTime:desc",
  };

  if (opts.questionId) params["Filter"] = `QuestionId:eq:${opts.questionId}`;

  return bvFetch<Answer>(instance, "/data/answers.json", params);
}

