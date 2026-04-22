// Live test against Microsoft Graph.
// Starts the server, sends initialize + list_emails, streams stderr (so the
// device-code URL is visible), and prints the result.
//
// Usage:
//   $env:OUTLOOK_MCP_CLIENT_ID = "<your-client-id>"
//   node scripts/live-test.mjs
import { spawn } from "node:child_process";

if (!process.env.OUTLOOK_MCP_CLIENT_ID) {
  console.error("Set OUTLOOK_MCP_CLIENT_ID first.");
  process.exit(2);
}

const child = spawn("node", ["dist/index.js"], {
  stdio: ["pipe", "pipe", "pipe"],
  env: process.env,
});

// Forward the server's stderr straight to our terminal. This is where
// MSAL writes the device-code prompt ("open https://microsoft.com/devicelogin
// and enter code ABCD-1234").
child.stderr.on("data", (chunk) => {
  process.stderr.write(chunk);
});

let buf = "";
const pending = new Map();

child.stdout.on("data", (chunk) => {
  buf += chunk.toString("utf-8");
  let idx;
  while ((idx = buf.indexOf("\n")) !== -1) {
    const line = buf.slice(0, idx).trim();
    buf = buf.slice(idx + 1);
    if (!line) continue;
    try {
      const msg = JSON.parse(line);
      const r = pending.get(msg.id);
      if (r) {
        pending.delete(msg.id);
        r(msg);
      }
    } catch {
      // ignore non-JSON
    }
  }
});

function send(obj) {
  child.stdin.write(JSON.stringify(obj) + "\n");
  return new Promise((resolve) => pending.set(obj.id, resolve));
}

try {
  console.log("[live-test] initializing...");
  await send({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "live-test", version: "0.0.1" },
    },
  });

  console.log("[live-test] calling list_emails (limit: 5)...");
  console.log("[live-test] FIRST CALL: if you see a device-code URL above,");
  console.log("[live-test] open it in a browser and sign in. DO NOT CTRL+C.");
  console.log("[live-test] The call will finish once you've signed in.");
  console.log("");

  const resp = await send({
    jsonrpc: "2.0",
    id: 2,
    method: "tools/call",
    params: { name: "list_emails", arguments: { limit: 5 } },
  });

  console.log("\n[live-test] response:");
  console.log(JSON.stringify(resp, null, 2));

  child.kill();
  process.exit(0);
} catch (e) {
  console.error("[live-test] error:", e);
  child.kill();
  process.exit(1);
}
