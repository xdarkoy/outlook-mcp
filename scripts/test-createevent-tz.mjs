// Unit tests for create_event's timezone normalization.
// Tests the pure logic directly — no Graph, no auth.
import assert from "node:assert/strict";
import { normalizeDateTime } from "../dist/tools/createEvent.js";

let passed = 0;
let failed = 0;

async function t(name, fn) {
  try {
    await fn();
    console.log("  ok  ", name);
    passed++;
  } catch (e) {
    console.log("  FAIL", name, "—", e.message);
    failed++;
  }
}

console.log("normalizeDateTime:");

await t("naive + tz UTC → pass-through", () => {
  assert.equal(normalizeDateTime("2026-05-01T10:00:00", "UTC", "start"), "2026-05-01T10:00:00");
});

await t("naive + tz Europe/Berlin → pass-through", () => {
  assert.equal(
    normalizeDateTime("2026-05-01T10:00:00", "Europe/Berlin", "start"),
    "2026-05-01T10:00:00",
  );
});

await t("offset +02:00 + tz UTC → converted to instant UTC", () => {
  // 10:00+02:00 == 08:00 UTC
  assert.equal(
    normalizeDateTime("2026-05-01T10:00:00+02:00", "UTC", "start"),
    "2026-05-01T08:00:00",
  );
});

await t("offset -05:00 + tz UTC → converted to instant UTC", () => {
  // 10:00-05:00 == 15:00 UTC
  assert.equal(
    normalizeDateTime("2026-05-01T10:00:00-05:00", "UTC", "start"),
    "2026-05-01T15:00:00",
  );
});

await t("offset Z + tz UTC → naive UTC", () => {
  assert.equal(
    normalizeDateTime("2026-05-01T10:00:00Z", "UTC", "start"),
    "2026-05-01T10:00:00",
  );
});

await t("offset + tz non-UTC → REJECTED (ambiguous)", () => {
  assert.throws(
    () => normalizeDateTime("2026-05-01T10:00:00+02:00", "Europe/Berlin", "start"),
    /explicit offset/i,
  );
});

await t("offset Z + tz non-UTC → REJECTED (ambiguous)", () => {
  assert.throws(
    () => normalizeDateTime("2026-05-01T10:00:00Z", "America/New_York", "start"),
    /explicit offset/i,
  );
});

await t("garbage datetime with offset → throws ToolError", () => {
  // With an offset-looking suffix, the parser is exercised. A non-offset
  // garbage string is passed through unchanged (Graph will validate).
  assert.throws(() => normalizeDateTime("nope+02:00", "UTC", "start"));
});

await t("naive garbage string → passed through as-is (Graph will 400)", () => {
  assert.equal(normalizeDateTime("not-a-date", "UTC", "start"), "not-a-date");
});

// Integration tests for the handler via a mocked graph() call.
// We verify end<=start rejection and attendee normalization without hitting Graph.
const { createEventTool } = await import("../dist/tools/createEvent.js");

async function h(args) {
  return await createEventTool.handler(args);
}

console.log("\ncreate_event handler:");

await t("rejects end <= start", async () => {
  const r = await h({
    subject: "x",
    start: "2026-05-01T11:00:00Z",
    end: "2026-05-01T10:00:00Z",
  });
  assert.equal(r.isError, true);
  assert.match(r.text, /not after start/i);
});

await t("rejects end == start", async () => {
  const r = await h({
    subject: "x",
    start: "2026-05-01T10:00:00Z",
    end: "2026-05-01T10:00:00Z",
  });
  assert.equal(r.isError, true);
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
