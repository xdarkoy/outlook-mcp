// Unit tests for list_emails' buildFilter — particularly the
// InefficientFilter workaround. Graph rejects `startsWith(from/…, …)` +
// `orderby receivedDateTime desc` unless the filter also touches
// receivedDateTime, so we inject a baseline `receivedDateTime ge` clause
// whenever `from` is set without a user-supplied `since`.
import assert from "node:assert/strict";
import { buildFilter } from "../dist/tools/listEmails.js";

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

console.log("buildFilter:");

t("no args → undefined (no filter clause)", () => {
  assert.equal(buildFilter({}), undefined);
});

t("from-only → injects baseline receivedDateTime ge (InefficientFilter workaround)", () => {
  const out = buildFilter({ from: "acme.com" });
  assert.ok(out.includes("startsWith(from/emailAddress/address, 'acme.com')"));
  assert.ok(
    out.includes("receivedDateTime ge 1900-01-01T00:00:00Z"),
    `expected baseline receivedDateTime ge clause; got: ${out}`,
  );
});

t("from + since → uses user's since, no baseline injected", () => {
  const out = buildFilter({
    from: "acme.com",
    since: "2026-04-01T00:00:00Z",
  });
  assert.ok(out.includes("receivedDateTime ge 2026-04-01T00:00:00Z"));
  assert.ok(
    !out.includes("1900-01-01"),
    `baseline should NOT be injected when since is provided; got: ${out}`,
  );
});

t("since-only → no baseline (no startsWith means no InefficientFilter risk)", () => {
  const out = buildFilter({ since: "2026-04-01T00:00:00Z" });
  assert.ok(out.includes("receivedDateTime ge 2026-04-01T00:00:00Z"));
  assert.ok(!out.includes("1900-01-01"));
});

t("until-only → no baseline needed", () => {
  const out = buildFilter({ until: "2026-05-01T00:00:00Z" });
  assert.ok(out.includes("receivedDateTime lt 2026-05-01T00:00:00Z"));
  assert.ok(!out.includes("1900-01-01"));
});

t("unreadOnly-only → no baseline (isRead filter doesn't trigger inefficient-filter)", () => {
  const out = buildFilter({ unreadOnly: true });
  assert.equal(out, "isRead eq false");
});

t("from + unread → still injects baseline (from drives the requirement)", () => {
  const out = buildFilter({ from: "acme.com", unreadOnly: true });
  assert.ok(out.includes("1900-01-01"));
  assert.ok(out.includes("isRead eq false"));
});

t("from + since + until + unread → all four clauses, no baseline", () => {
  const out = buildFilter({
    from: "acme.com",
    since: "2026-04-01T00:00:00Z",
    until: "2026-05-01T00:00:00Z",
    unreadOnly: true,
  });
  assert.ok(out.includes("startsWith(from/emailAddress/address, 'acme.com')"));
  assert.ok(out.includes("receivedDateTime ge 2026-04-01T00:00:00Z"));
  assert.ok(out.includes("receivedDateTime lt 2026-05-01T00:00:00Z"));
  assert.ok(out.includes("isRead eq false"));
  assert.ok(!out.includes("1900-01-01"));
});

t("from value with single quote is OData-escaped to ''", () => {
  // Pathological but valid: an apostrophe in the address must be doubled.
  const out = buildFilter({ from: "o'brien@acme.com" });
  assert.ok(out.includes("'o''brien@acme.com'"));
});

t("clause order is stable (filter → date → state)", () => {
  // The Graph SDK does not require any particular order, but a stable
  // shape makes the request log easier to read and tests easier to write.
  const out = buildFilter({
    from: "acme.com",
    since: "2026-04-01T00:00:00Z",
    unreadOnly: true,
  });
  const a = out.indexOf("startsWith");
  const b = out.indexOf("receivedDateTime");
  const c = out.indexOf("isRead");
  assert.ok(a < b && b < c, `order should be from → date → state; got: ${out}`);
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
