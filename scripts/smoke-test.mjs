// Minimal smoke test: spawn the server, initialize, then list tools.
// Sends the next request only after the previous one's reply has arrived —
// no flaky setTimeout race. No auth required; this exercises the MCP
// transport and registry only.
import { spawn } from "node:child_process";

const child = spawn("node", ["dist/index.js"], {
  stdio: ["pipe", "pipe", "inherit"],
});

let buf = "";
const pending = new Map(); // id -> resolve
const hardTimeout = setTimeout(() => {
  console.error("Timed out after 8s");
  child.kill();
  process.exit(1);
}, 8000);

child.stdout.on("data", (chunk) => {
  buf += chunk.toString("utf-8");
  let idx;
  while ((idx = buf.indexOf("\n")) !== -1) {
    const line = buf.slice(0, idx).trim();
    buf = buf.slice(idx + 1);
    if (!line) continue;
    try {
      const msg = JSON.parse(line);
      console.log("<<", JSON.stringify(msg, null, 2));
      const r = pending.get(msg.id);
      if (r) {
        pending.delete(msg.id);
        r(msg);
      }
    } catch {
      console.log("<< (non-JSON)", line);
    }
  }
});

function send(obj) {
  const s = JSON.stringify(obj) + "\n";
  console.log(">>", s.trim());
  child.stdin.write(s);
  return new Promise((resolve) => pending.set(obj.id, resolve));
}

try {
  await send({
    jsonrpc: "2.0",
    id: 1,
    method: "initialize",
    params: {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "smoke-test", version: "0.0.1" },
    },
  });
  await send({ jsonrpc: "2.0", id: 2, method: "tools/list", params: {} });
  clearTimeout(hardTimeout);
  child.kill();
  process.exit(0);
} catch (e) {
  console.error(e);
  clearTimeout(hardTimeout);
  child.kill();
  process.exit(1);
}
