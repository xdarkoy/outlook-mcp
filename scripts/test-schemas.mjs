// Schema-level validation tests. Each test parses a known-good or
// known-bad input and asserts the right outcome — these run without
// touching Graph or MSAL.
import assert from "node:assert/strict";
import {
  SaveAttachmentInput,
  MarkEmailInput,
  MoveEmailInput,
  ListFoldersInput,
  CreateEventInput,
  UpdateEventInput,
  DeleteEventInput,
} from "../dist/schemas.js";

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

console.log("SaveAttachmentInput.targetFilename:");

t("plain filename accepted", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
    targetFilename: "invoice.pdf",
  });
  assert.equal(r.success, true);
});

t("targetFilename omitted is accepted (optional)", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
  });
  assert.equal(r.success, true);
});

t("forward-slash REJECTED at schema level", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
    targetFilename: "a/b.pdf",
  });
  assert.equal(r.success, false);
});

t("backslash REJECTED at schema level", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
    targetFilename: "a\\b.pdf",
  });
  assert.equal(r.success, false);
});

t("traversal pattern (..) is REJECTED via separators", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
    targetFilename: "../etc/passwd",
  });
  assert.equal(r.success, false);
});

t("empty string is rejected (min(1))", () => {
  const r = SaveAttachmentInput.safeParse({
    messageId: "m1",
    attachmentId: "a1",
    targetFilename: "",
  });
  assert.equal(r.success, false);
});

console.log("\nMarkEmailInput:");

t("happy path", () => {
  const r = MarkEmailInput.safeParse({ messageId: "m1", read: true });
  assert.equal(r.success, true);
});

t("missing 'read' field rejected", () => {
  const r = MarkEmailInput.safeParse({ messageId: "m1" });
  assert.equal(r.success, false);
});

t("non-boolean 'read' rejected", () => {
  const r = MarkEmailInput.safeParse({ messageId: "m1", read: "yes" });
  assert.equal(r.success, false);
});

console.log("\nMoveEmailInput:");

t("wellknown destination accepted", () => {
  const r = MoveEmailInput.safeParse({
    messageId: "m1",
    destinationFolder: "archive",
  });
  assert.equal(r.success, true);
});

t("arbitrary folder ID accepted (loose validation)", () => {
  const r = MoveEmailInput.safeParse({
    messageId: "m1",
    destinationFolder: "AAMkAGI2THVS...opaqueID...",
  });
  assert.equal(r.success, true);
});

t("empty destination rejected", () => {
  const r = MoveEmailInput.safeParse({
    messageId: "m1",
    destinationFolder: "",
  });
  assert.equal(r.success, false);
});

console.log("\nListFoldersInput:");

t("no args is valid (lists top-level)", () => {
  const r = ListFoldersInput.safeParse({});
  assert.equal(r.success, true);
});

t("parentFolderId accepted", () => {
  const r = ListFoldersInput.safeParse({ parentFolderId: "fid123" });
  assert.equal(r.success, true);
});

t("limit > 200 rejected", () => {
  const r = ListFoldersInput.safeParse({ limit: 201 });
  assert.equal(r.success, false);
});

console.log("\nCreateEventInput datetime handling:");

t("naive datetime + Europe/Berlin accepted (was rejected pre-fix)", () => {
  const r = CreateEventInput.safeParse({
    subject: "Lunch",
    start: "2026-05-01T12:00:00",
    end: "2026-05-01T13:00:00",
    timeZone: "Europe/Berlin",
  });
  assert.equal(r.success, true);
});

t("offset datetime (Z) + UTC accepted", () => {
  const r = CreateEventInput.safeParse({
    subject: "Sync",
    start: "2026-05-01T10:00:00Z",
    end: "2026-05-01T10:30:00Z",
  });
  assert.equal(r.success, true);
});

t("offset datetime (+02:00) + UTC accepted at schema (handler normalizes)", () => {
  const r = CreateEventInput.safeParse({
    subject: "Sync",
    start: "2026-05-01T10:00:00+02:00",
    end: "2026-05-01T10:30:00+02:00",
    timeZone: "UTC",
  });
  assert.equal(r.success, true);
});

t("malformed datetime rejected", () => {
  const r = CreateEventInput.safeParse({
    subject: "x",
    start: "not-a-date",
    end: "2026-05-01T10:30:00Z",
  });
  assert.equal(r.success, false);
});

console.log("\nUpdateEventInput:");

t("eventId-only is valid (handler will reject empty payload)", () => {
  // Schema allows it — semantic check (no fields to update) lives in the handler.
  const r = UpdateEventInput.safeParse({ eventId: "e1" });
  assert.equal(r.success, true);
});

t("partial update with subject only is valid", () => {
  const r = UpdateEventInput.safeParse({ eventId: "e1", subject: "New" });
  assert.equal(r.success, true);
});

t("invalid datetime rejected", () => {
  const r = UpdateEventInput.safeParse({
    eventId: "e1",
    start: "not-a-date",
  });
  assert.equal(r.success, false);
});

t("UpdateEventInput naive datetime accepted (parity with CreateEventInput)", () => {
  const r = UpdateEventInput.safeParse({
    eventId: "e1",
    start: "2026-05-01T12:00:00",
    timeZone: "Europe/Berlin",
  });
  assert.equal(r.success, true);
});

t("attendees: empty array valid (means: remove all attendees)", () => {
  const r = UpdateEventInput.safeParse({ eventId: "e1", attendees: [] });
  assert.equal(r.success, true);
});

t("attendees: invalid email rejected", () => {
  const r = UpdateEventInput.safeParse({
    eventId: "e1",
    attendees: ["not-an-email"],
  });
  assert.equal(r.success, false);
});

console.log("\nDeleteEventInput:");

t("eventId required", () => {
  const r = DeleteEventInput.safeParse({});
  assert.equal(r.success, false);
});

t("eventId valid", () => {
  const r = DeleteEventInput.safeParse({ eventId: "e1" });
  assert.equal(r.success, true);
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
