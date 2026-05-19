// Unit tests for searchEmails MSA-branch query escaping.
// The wrapper is exported precisely so a bad escape doesn't ship silently.
import assert from "node:assert/strict";
import { escapeMsaSearchQuery } from "../dist/tools/searchEmails.js";

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

console.log("escapeMsaSearchQuery:");

t("plain query is just quoted", () => {
  assert.equal(escapeMsaSearchQuery("invoice"), '"invoice"');
});

t("KQL operators pass through unchanged", () => {
  assert.equal(
    escapeMsaSearchQuery("from:acme.com hasAttachment:true"),
    '"from:acme.com hasAttachment:true"',
  );
});

t("embedded double quote is backslash-escaped", () => {
  // Important: the escape is \" not "" (Graph $search uses backslash).
  assert.equal(escapeMsaSearchQuery('say "hi"'), '"say \\"hi\\""');
});

t("multiple embedded quotes all escaped", () => {
  assert.equal(escapeMsaSearchQuery('"a" "b"'), '"\\"a\\" \\"b\\""');
});

t("control chars (newline, tab, NUL) replaced with space", () => {
  assert.equal(escapeMsaSearchQuery("foo\nbar"), '"foo bar"');
  assert.equal(escapeMsaSearchQuery("a\tb"), '"a b"');
  assert.equal(escapeMsaSearchQuery("x\x00y"), '"x y"');
});

t("DEL (0x7f) replaced with space", () => {
  assert.equal(escapeMsaSearchQuery("x\x7fy"), '"x y"');
});

t("backslash is stripped (avoids broken KQL escape sequences)", () => {
  // Raw backslashes cause ambiguity with KQL's own \ escape sequences.
  // Strategy: replace with space rather than try to second-guess KQL.
  // Input: a, \, ", b  →  expected: ", a, space, \, ", b, "
  assert.equal(escapeMsaSearchQuery('a\\"b'), '"a \\"b"');
});

t("lone backslashes are stripped without affecting non-quote text", () => {
  assert.equal(escapeMsaSearchQuery("path\\to\\thing"), '"path to thing"');
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
