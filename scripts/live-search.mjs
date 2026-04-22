// One-off: call search_emails with a job-related query to verify the MSA fix.
import { spawn } from "node:child_process";

if (!process.env.OUTLOOK_MCP_CLIENT_ID) {
  console.error("Set OUTLOOK_MCP_CLIENT_ID first.");
  process.exit(2);
}

const child = spawn("node", ["dist/index.js"], {
  stdio: ["pipe", "pipe", "pipe"],
  env: process.env,
});
child.stderr.on("data", (c) => process.stderr.write(c));

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
      if (r) { pending.delete(msg.id); r(msg); }
    } catch {}
  }
});

function send(obj) {
  child.stdin.write(JSON.stringify(obj) + "\n");
  return new Promise((r) => pending.set(obj.id, r));
}

const query = process.argv[2] ?? "invoice OR receipt";

try {
  await send({
    jsonrpc: "2.0", id: 1, method: "initialize",
    params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "live-search", version: "0.0.1" } },
  });

  const resp = await send({
    jsonrpc: "2.0", id: 2, method: "tools/call",
    params: { name: "search_emails", arguments: { query, limit: 25 } },
  });

  console.log("\n[result]");
  console.log(JSON.stringify(resp, null, 2));
  child.kill();
  process.exit(0);
} catch (e) {
  console.error(e);
  child.kill();
  process.exit(1);
}
