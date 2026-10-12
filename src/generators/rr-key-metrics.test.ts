import assert from "node:assert/strict";
import test from "node:test";

import { resolveValidInstanceNames } from "./rr-key-metrics.js";

test("resolveValidInstanceNames skips invalid instances but keeps valid ones", async () => {
  const names = ["bad-passkey", "pampers-en-us", "also-bad"];

  const valid = await resolveValidInstanceNames(names, async (instanceName) => {
    if (instanceName === "pampers-en-us") {
      return { Results: [], TotalResults: 0 };
    }

    throw new Error(`invalid passkey for ${instanceName}`);
  });

  assert.deepEqual(valid, ["pampers-en-us"]);
});
