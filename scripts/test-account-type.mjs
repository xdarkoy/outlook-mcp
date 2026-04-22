// Unit tests for auth/accountType.ts using the pure classifier.
import assert from "node:assert/strict";
import { classifyAccount } from "../dist/auth/accountType.js";

let passed = 0;
let failed = 0;

function t(name, fn) {
  try {
    fn();
    console.log("  ok  ", name);
    passed++;
  } catch (e) {
    console.log("  FAIL", name, "—", e.message);
    failed++;
  }
}

console.log("classifyAccount:");

t("MSA sentinel tenant → 'msa'", () => {
  assert.equal(
    classifyAccount({ tenantId: "9188040d-6c67-4c5b-b112-36a304b66dad" }),
    "msa",
  );
});

t("Work tenant GUID → 'aad'", () => {
  assert.equal(
    classifyAccount({ tenantId: "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee" }),
    "aad",
  );
});

t("No account (not signed in) → 'aad' (fail-open)", () => {
  assert.equal(classifyAccount(null), "aad");
});

t("Undefined account → 'aad' (fail-open)", () => {
  assert.equal(classifyAccount(undefined), "aad");
});

t("Account with missing tenantId → 'aad'", () => {
  assert.equal(classifyAccount({}), "aad");
});

t("MSA tenant case-sensitivity (must be lowercase)", () => {
  // MSAL always returns lowercase — uppercase must not accidentally match.
  assert.equal(
    classifyAccount({ tenantId: "9188040D-6C67-4C5B-B112-36A304B66DAD" }),
    "aad",
  );
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
