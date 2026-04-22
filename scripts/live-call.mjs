// Call an arbitrary tool on a fresh server instance and print the result.
// Usage: node scripts/live-call.mjs <tool_name> '<json_args>'
import { spawn } from "node:child_process";

if (!process.env.OUTLOOK_MCP_CLIENT_ID) {
  console.error("Set OUTLOOK_MCP_CLIENT_ID first.");
  process.exit(2);
}
const [toolName, rawArgs = "{}"] = process.argv.slice(2);
if (!toolName) {
  console.error("Usage: node scripts/live-call.mjs <tool_name> '<json_args>'");
  process.exit(2);
}
let args;
try { args = JSON.parse(rawArgs); }
catch (e) { console.error("Invalid JSON args:", e.message); process.exit(2); }

const child = spawn("node", ["dist/index.js"], { stdio: ["pipe", "pipe", "pipe"], env: process.env });
child.stderr.on("data", (c) => process.stderr.write(c));
let buf = ""; const pending = new Map();
child.stdout.on("data", (chunk) => {
  buf += chunk.toString("utf-8");
  let idx;
  while ((idx = buf.indexOf("\n")) !== -1) {
    const line = buf.slice(0, idx).trim(); buf = buf.slice(idx + 1);
    if (!line) continue;
    try { const msg = JSON.parse(line); const r = pending.get(msg.id); if (r) { pending.delete(msg.id); r(msg); } } catch {}
  }
});
function send(obj) { child.stdin.write(JSON.stringify(obj) + "\n"); return new Promise((r) => pending.set(obj.id, r)); }
try {
  await send({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "live-call", version: "0.0.1" } } });
  const resp = await send({ jsonrpc: "2.0", id: 2, method: "tools/call", params: { name: toolName, arguments: args } });
  const text = resp.result?.content?.[0]?.text ?? "";
  if (resp.result?.isError) { console.error("ERROR:\n" + text); child.kill(); process.exit(1); }
  console.log(text);
  child.kill(); process.exit(0);
} catch (e) { console.error(e); child.kill(); process.exit(1); }
