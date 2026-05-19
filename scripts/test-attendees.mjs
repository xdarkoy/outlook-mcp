// Unit tests for createEvent's normalizeAttendees (also used by update_event).
// Pure function — no Graph, no auth.
import assert from "node:assert/strict";
import { normalizeAttendees } from "../dist/tools/createEvent.js";

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

console.log("normalizeAttendees:");

t("bare email becomes required attendee", () => {
  const out = normalizeAttendees(["alice@acme.com"]);
  assert.deepEqual(out, [
    { emailAddress: { address: "alice@acme.com" }, type: "required" },
  ]);
});

t("object form preserves type", () => {
  const out = normalizeAttendees([
    { address: "room1@acme.com", type: "resource" },
  ]);
  assert.deepEqual(out, [
    { emailAddress: { address: "room1@acme.com" }, type: "resource" },
  ]);
});

t("missing type defaults to required", () => {
  const out = normalizeAttendees([{ address: "bob@acme.com" }]);
  assert.equal(out[0].type, "required");
});

t("dedup is case-insensitive on address", () => {
  const out = normalizeAttendees([
    "Alice@acme.com",
    { address: "alice@ACME.com", type: "optional" },
  ]);
  assert.equal(out.length, 1);
  // Required (from the first entry, which is the bare string) wins by priority.
  assert.equal(out[0].type, "required");
});

t("tiebreaker: required wins over optional", () => {
  const out = normalizeAttendees([
    { address: "x@y.z", type: "optional" },
    { address: "x@y.z", type: "required" },
  ]);
  assert.equal(out.length, 1);
  assert.equal(out[0].type, "required");
});

t("tiebreaker: required wins regardless of order", () => {
  const out = normalizeAttendees([
    { address: "x@y.z", type: "required" },
    { address: "x@y.z", type: "optional" },
  ]);
  assert.equal(out[0].type, "required");
});

t("tiebreaker: optional wins over resource", () => {
  const out = normalizeAttendees([
    { address: "x@y.z", type: "resource" },
    { address: "x@y.z", type: "optional" },
  ]);
  assert.equal(out[0].type, "optional");
});

t("tiebreaker: required wins over resource (3-way)", () => {
  const out = normalizeAttendees([
    { address: "x@y.z", type: "resource" },
    { address: "x@y.z", type: "required" },
    { address: "x@y.z", type: "optional" },
  ]);
  assert.equal(out[0].type, "required");
});

t("preserves the original case of the address (only key is lowercased)", () => {
  const out = normalizeAttendees([
    { address: "Alice@Acme.COM", type: "required" },
  ]);
  assert.equal(out[0].emailAddress.address, "Alice@Acme.COM");
});

t("empty input → empty output (does not crash)", () => {
  const out = normalizeAttendees([]);
  assert.deepEqual(out, []);
});

t("multiple distinct attendees preserved in input order", () => {
  const out = normalizeAttendees([
    "a@x.com",
    { address: "b@x.com", type: "optional" },
    { address: "c@x.com", type: "resource" },
  ]);
  assert.equal(out.length, 3);
  assert.equal(out[0].emailAddress.address, "a@x.com");
  assert.equal(out[1].emailAddress.address, "b@x.com");
  assert.equal(out[2].emailAddress.address, "c@x.com");
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
