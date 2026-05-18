# Bazaarvoice MCP Server (Starter)

Hello! This project is a starter MCP server for connecting VS Code to Bazaarvoice.

Current capabilities:
- `bv_list_instances`: list configured BV instances from local config
- `bv_get_product_count`: get total product count for one BV instance
- `bv_count_instances_by_owner`: local metadata count by owner (for future P&G-focused reporting)
- `bv_generate_rr_key_metrics_report`: generate an R&R Key Metrics workbook similar to Bazaarvoice_Pampers_RR_Key_Metrics.xlsx
- `bv_generate_instance_audit_report`: generate an Audit workbook similar to Bazaarvoice_Pampers_Audit_Report.xlsx

## Capability Reference

This section documents all tools currently exposed by this MCP server.

### Instance management

1. `bv_list_instances`
  - Purpose: list all Bazaarvoice instances configured in `BV_INSTANCES_JSON`.
  - Input: none.
  - Output: `{ count, instances[] }` where each instance includes `name`, `baseUrl`, `apiversion`, and `owner`.
  - Notes: local config only, no network call.

2. `bv_count_instances_by_owner`
  - Purpose: count configured instances by owner label (for example `P&G`).
  - Input: `owner` (string).
  - Output: `{ owner, count, instances[] }`.
  - Notes: local config only, case-insensitive owner match.

### Products

3. `bv_get_product_count`
  - Purpose: return total product count for an instance catalog.
  - Input: `instanceName` (string).
  - Output: `{ instanceName, totalProducts }`.
  - Notes: API call to `/data/products.json` with `Limit=1` and uses `TotalResults`.

4. `bv_list_products`
  - Purpose: list products with paging, search, and optional inline review statistics.
  - Input:
    - `instanceName` (required)
    - `limit` (optional, `1-100`, default `10`)
    - `offset` (optional, default `0`)
    - `search` (optional)
    - `sort` (optional, for example `Name:asc`)
    - `includeStats` (optional boolean)
  - Output: `{ totalResults, limit, offset, products[] }`.

### Product statistics

5. `bv_get_product_statistics`
  - Purpose: fetch aggregate review statistics for one product.
  - Input: `instanceName` (string), `productId` (string).
  - Output: `{ productId, reviewStatistics, nativeReviewStatistics }`.
  - Notes: uses `/data/statistics.json` with `Stats=Reviews,NativeReviews`.

### Reviews

6. `bv_list_reviews`
  - Purpose: list reviews with filtering, paging, sort, and full-text search.
  - Input:
    - `instanceName` (required)
    - `productId` (optional)
    - `limit` (optional, `1-100`, default `10`)
    - `offset` (optional, default `0`)
    - `sort` (optional, default `SubmissionTime:desc`)
    - `search` (optional)
  - Output: `{ totalResults, limit, offset, reviews[] }`.

### Questions and answers

7. `bv_list_questions`
  - Purpose: list Q&A questions, optionally including answers inline.
  - Input:
    - `instanceName` (required)
    - `productId` (optional)
    - `limit` (optional, `1-100`, default `10`)
    - `offset` (optional, default `0`)
    - `sort` (optional, default `SubmissionTime:desc`)
    - `includeAnswers` (optional boolean)
  - Output: `{ totalResults, limit, offset, questions[], includes }`.

8. `bv_list_answers`
  - Purpose: list answers, optionally filtered by question ID.
  - Input:
    - `instanceName` (required)
    - `questionId` (optional)
    - `limit` (optional, `1-100`, default `10`)
    - `offset` (optional, default `0`)
    - `sort` (optional, default `SubmissionTime:desc`)
  - Output: `{ totalResults, limit, offset, answers[] }`.

### Reporting

9. `bv_generate_rr_key_metrics_report`
  - Purpose: generate a local Excel `.xlsx` report with key R&R metrics for one instance.
  - Input: `instanceName` (optional, defaults to `pampers-en-us`).
  - Output: `{ instanceName, generatedAt, outputPath, totalProducts, sheetName }`.
  - Notes: writes the file under `reports/` in the workspace.

10. `bv_generate_instance_audit_report`
  - Purpose: generate an Audit workbook with dashboard, analysis, suggestions, and pivot sheets.
  - Input: `instanceName` (optional, defaults to `pampers-en-us`).
  - Output: `{ instanceName, outputPath, totalProducts, activeProducts, inactiveProducts, generatedAt }`.
  - Notes: writes the file under `reports/` in the workspace.

### Cross-cutting behavior and limits

- API mode: read-only Conversations API calls (GET endpoints under `/data/*`).
- Writes are not supported: no product catalog mutation (for example UPC updates), no create/update/delete operations.
- Config source: all instance credentials come from `BV_INSTANCES_JSON`.
- Timeout: all HTTP calls are bounded by `BV_TIMEOUT_MS` (default `12000` ms).
- Error handling: Bazaarvoice API errors are surfaced as MCP tool errors with BV code and message.

## 1) Install

```bash
npm install
```

## 2) Configure environment

Copy `.env.example` to `.env` and add your instance details.

Example:

```env
BV_INSTANCES_JSON=[{"name":"pg-prod-us","baseUrl":"https://api.bazaarvoice.com","passkey":"YOUR_PASSKEY","apiversion":"5.4","owner":"P&G"}]
BV_TIMEOUT_MS=12000
```

## 3) Run locally

```bash
npm run dev
```

Or build and run:

```bash
npm run build
npm start
```

Generate the RR Key Metrics workbook directly:

```bash
npm run report:rr:full -- pampers-en-us
```

Generate the Audit workbook directly:

```bash
npm run report:audit -- pampers-en-us
```

Dynamic consolidated CLI (single entry point for all report types):

```bash
npm run report -- rr-full pampers-en-us
npm run report -- rr-sample pampers-en-us
npm run report -- audit pampers-en-us
```

## 4) Connect from VS Code MCP

Add this server in your VS Code MCP client config (path depends on your setup):

```json
{
  "mcpServers": {
    "bazaarvoice": {
      "command": "node",
      "args": ["dist/index.js"],
      "cwd": "c:/Vs code/Custom-tool",
      "env": {
        "BV_INSTANCES_JSON": "[{\"name\":\"pg-prod-us\",\"baseUrl\":\"https://api.bazaarvoice.com\",\"passkey\":\"YOUR_PASSKEY\",\"apiversion\":\"5.4\",\"owner\":\"P&G\"}]",
        "BV_TIMEOUT_MS": "12000"
      }
    }
  }
}
```

If your MCP client supports dotenv loading in `cwd`, you can keep secrets in `.env` instead of inline `env`.

## Notes

- Product count currently calls: `/data/products.json` with `Limit=1` and reads `TotalResults`.
- Owner counting is local to `BV_INSTANCES_JSON` metadata right now.
- For the next phase ("number of BV instances for P&G" from upstream APIs), you can replace `bv_count_instances_by_owner` internals with a real Bazaarvoice account/discovery API call when credentials and endpoint details are confirmed.
