// Regression tests for the outbound-mail trust boundary.
// No auth required: this checks configured scopes and the advertised MCP tools.
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { DEFAULT_SCOPES } from "../dist/auth/msal.js";

let passed = 0;
let failed = 0;

async function t(name, fn) {
  try {
    await fn();
    console.log("  ok  ", name);
    passed++;
  } catch (e) {
    console.log("  FAIL", name, "-", e.message);
    failed++;
  }
}

function startServer() {
  const child = spawn("node", ["dist/index.js"], {
    stdio: ["pipe", "pipe", "inherit"],
  });

  let buf = "";
  const pending = new Map();
  const hardTimeout = setTimeout(() => {
    child.kill();
    for (const reject of pending.values()) reject(new Error("Timed out after 8s"));
    pending.clear();
  }, 8000);

  child.stdout.on("data", (chunk) => {
    buf += chunk.toString("utf-8");
    let idx;
    while ((idx = buf.indexOf("\n")) !== -1) {
      const line = buf.slice(0, idx).trim();
      buf = buf.slice(idx + 1);
      if (!line) continue;
      const msg = JSON.parse(line);
      const waiter = pending.get(msg.id);
      if (waiter) {
        pending.delete(msg.id);
        waiter.resolve(msg);
      }
    }
  });

  function send(obj) {
    child.stdin.write(JSON.stringify(obj) + "\n");
    return new Promise((resolve, reject) => pending.set(obj.id, { resolve, reject }));
  }

  function stop() {
    clearTimeout(hardTimeout);
    child.kill();
  }

  return { send, stop };
}

async function listTools() {
  const server = startServer();
  try {
    await server.send({
      jsonrpc: "2.0",
      id: 1,
      method: "initialize",
      params: {
        protocolVersion: "2024-11-05",
        capabilities: {},
        clientInfo: { name: "mail-trust-boundary-test", version: "0.0.1" },
      },
    });
    const resp = await server.send({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
    return resp.result?.tools ?? [];
  } finally {
    server.stop();
  }
}

console.log("mail trust boundary:");

await t("OAuth scopes do not include Mail.Send", () => {
  assert.equal(DEFAULT_SCOPES.includes("Mail.Send"), false);
});

await t("create_draft remains advertised as draft-only", async () => {
  const tools = await listTools();
  const draft = tools.find((tool) => tool.name === "create_draft");
  assert.ok(draft, "create_draft tool missing");
  assert.match(draft.description, /never sends/i);
});

await t("no advertised tool is named like a send action", async () => {
  const tools = await listTools();
  const sendLike = tools.map((tool) => tool.name).filter((name) => /(^|_)send(_|$)/i.test(name));
  assert.deepEqual(sendLike, []);
});

console.log(`\n${passed} passed, ${failed} failed`);
process.exit(failed ? 1 : 0);
