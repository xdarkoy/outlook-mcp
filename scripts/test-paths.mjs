// Quick sanity checks for sanitizeFilename / resolveWritePath.
// Not a full test suite — just coverage for the security-critical edges.
import assert from "node:assert/strict";
import path from "node:path";
import os from "node:os";
import { promises as fs } from "node:fs";
import {
  sanitizeFilename,
  allowedDir,
  resolveWritePath,
  suffixedPath,
} from "../dist/util/paths.js";

const tmp = path.join(os.tmpdir(), "outlook-mcp-test-" + Date.now());
process.env.OUTLOOK_MCP_ALLOWED_DIR = tmp;

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

async function ta(name, fn) {
  try {
    await fn();
    console.log("  ok  ", name);
    passed++;
  } catch (e) {
    console.log("  FAIL", name, "—", e.message);
    failed++;
  }
}

console.log("sanitizeFilename:");

t("plain filename", () => assert.equal(sanitizeFilename("report.pdf"), "report.pdf"));
t("strips forward slash", () => assert.equal(sanitizeFilename("a/b.pdf"), "a_b.pdf"));
t("strips backslash", () => assert.equal(sanitizeFilename("a\\b.pdf"), "a_b.pdf"));
t("strips colon (ADS/drive)", () => assert.equal(sanitizeFilename("file.txt:stream"), "file.txt_stream"));
t("strips drive letter relative", () => assert.equal(sanitizeFilename("C:foo.pdf"), "C_foo.pdf"));
t("rejects empty", () => assert.throws(() => sanitizeFilename("")));
t("rejects '.'", () => assert.throws(() => sanitizeFilename(".")));
t("rejects '..'", () => assert.throws(() => sanitizeFilename("..")));
t("rejects CON", () => assert.throws(() => sanitizeFilename("CON")));
t("rejects con.txt", () => assert.throws(() => sanitizeFilename("con.txt")));
t("rejects 'CON '", () => assert.throws(() => sanitizeFilename("CON ")));
t("rejects 'CON.'", () => assert.throws(() => sanitizeFilename("CON.")));
t("rejects 'NUL'", () => assert.throws(() => sanitizeFilename("NUL")));
t("rejects 'lpt1.log'", () => assert.throws(() => sanitizeFilename("lpt1.log")));
t("allows innocent leading dot", () => assert.equal(sanitizeFilename(".bashrc"), ".bashrc"));
t("strips pipe/question/asterisk", () =>
  assert.equal(sanitizeFilename("a|b?c*.txt"), "a_b_c_.txt"));

console.log("\nresolveWritePath:");

await ta("resolves inside allowed dir", async () => {
  const p = await resolveWritePath("invoice.pdf");
  assert.ok(p.startsWith(allowedDir() + path.sep) || p === path.join(allowedDir(), "invoice.pdf"));
});

await ta("rejects traversal via backslash", async () => {
  // sanitizer turns \ into _, so this becomes a legal filename — not an escape.
  const p = await resolveWritePath("..\\..\\etc\\passwd");
  assert.ok(p.startsWith(allowedDir() + path.sep));
});

await ta("resolveWritePath is deterministic (collision handled in writer)", async () => {
  // resolveWritePath returns the logical target; collision handling moved
  // into the saveAttachment writer via O_CREAT|O_EXCL retry. So two calls
  // for the same filename now return the SAME path — the caller is
  // responsible for atomic create.
  const a = await resolveWritePath("dup.txt");
  await fs.writeFile(a, "first");
  const b = await resolveWritePath("dup.txt");
  assert.equal(a, b, "resolveWritePath should be pure / deterministic now");
});

t("suffixedPath generates collision-avoiding variants", () => {
  assert.ok(suffixedPath("C:/tmp/invoice.pdf", 2).endsWith("invoice (2).pdf"));
  assert.ok(suffixedPath("/tmp/invoice.pdf", 3).endsWith("invoice (3).pdf"));
});

t("sanitizeFilename strips zero-width chars", () => {
  // U+200B ZERO WIDTH SPACE between 'a' and 'b'
  const zw = "a​b.txt";
  assert.equal(sanitizeFilename(zw), "ab.txt");
});

t("sanitizeFilename strips RLO override (invoice_RLO_fdp.exe visual attack)", () => {
  // U+202E RIGHT-TO-LEFT OVERRIDE
  const rlo = "invoice‮fdp.exe";
  const cleaned = sanitizeFilename(rlo);
  assert.ok(!cleaned.includes("‮"));
  assert.equal(cleaned, "invoicefdp.exe");
});

t("sanitizeFilename strips BOM", () => {
  const bom = "﻿report.pdf";
  assert.equal(sanitizeFilename(bom), "report.pdf");
});

await fs.rm(tmp, { recursive: true, force: true });

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
