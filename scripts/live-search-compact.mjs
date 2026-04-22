// Same as live-search.mjs but prints only a compact table (date, from, subject).
// Useful for scanning history without drowning in JSON.
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
const limit = Number(process.argv[3] ?? 100);

try {
  await send({
    jsonrpc: "2.0", id: 1, method: "initialize",
    params: { protocolVersion: "2024-11-05", capabilities: {}, clientInfo: { name: "live-search-compact", version: "0.0.1" } },
  });
  const resp = await send({
    jsonrpc: "2.0", id: 2, method: "tools/call",
    params: { name: "search_emails", arguments: { query, limit } },
  });

  const text = resp.result?.content?.[0]?.text ?? "";
  if (resp.result?.isError) {
    console.error("ERROR:", text);
    child.kill();
    process.exit(1);
  }
  const data = JSON.parse(text);
  console.log(`query: ${data.query}`);
  console.log(`backend: ${data.backend}`);
  console.log(`returned: ${data.returned}`);
  console.log("");

  // Sort by date descending if received is present
  const msgs = [...data.messages].sort((a, b) => {
    const ta = a.received ? Date.parse(a.received) : 0;
    const tb = b.received ? Date.parse(b.received) : 0;
    return tb - ta;
  });

  for (const m of msgs) {
    const dt = m.received ? m.received.slice(0, 10) : "????-??-??";
    const from = (m.from ?? "").padEnd(40).slice(0, 40);
    const subj = (m.subject ?? "").slice(0, 90);
    console.log(`${dt} | ${from} | ${subj}`);
  }

  child.kill();
  process.exit(0);
} catch (e) {
  console.error(e);
  child.kill();
  process.exit(1);
}
