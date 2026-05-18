import type { Product } from "./types.js";

export type ProductStatus = "Active" | "Inactive" | "Disabled";

function toBooleanLike(value: unknown): boolean | undefined {
  if (typeof value === "boolean") {
    return value;
  }

  if (typeof value === "string") {
    const normalized = value.trim().toLowerCase();
    if (normalized === "true") return true;
    if (normalized === "false") return false;
  }

  return undefined;
}

export function resolveProductStatus(product: Product): ProductStatus {
  const disabled = toBooleanLike((product as Product & { Disabled?: unknown }).Disabled);
  if (disabled === true) {
    return "Disabled";
  }

  const active = toBooleanLike((product as Product & { Active?: unknown }).Active);
  if (active === true) {
    return "Active";
  }
  if (active === false) {
    return "Inactive";
  }

  // Bazaarvoice can omit Active/Disabled on product payloads.
  // Fallback to ProductPageUrl presence, which best matches legacy reports.
  const hasPageUrl = Boolean(product.ProductPageUrl?.trim());
  return hasPageUrl ? "Active" : "Inactive";
}

export function isProductActive(product: Product): boolean {
  return resolveProductStatus(product) === "Active";
}
