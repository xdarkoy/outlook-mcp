#!/usr/bin/env node
import { Server } from "@modelcontextprotocol/sdk/server/index.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
  type CallToolResult,
} from "@modelcontextprotocol/sdk/types.js";
import { zodToJsonSchema } from "zod-to-json-schema";

import { runLogin } from "./cli/login.js";
import { listEmailsTool } from "./tools/listEmails.js";
import { readEmailTool } from "./tools/readEmail.js";
import { searchEmailsTool } from "./tools/searchEmails.js";
import { saveAttachmentTool } from "./tools/saveAttachment.js";
import { listCalendarEventsTool } from "./tools/listCalendarEvents.js";
import { createEventTool } from "./tools/createEvent.js";
import { createDraftTool } from "./tools/createDraft.js";
import { toAny, type AnyToolDef } from "./tools/types.js";

/**
 * STDOUT is the MCP JSON-RPC channel. Anything written there that is not a
 * valid JSON-RPC frame corrupts the stream and breaks the client. Common
 * culprits are transitive dependencies that call `console.log` directly
 * (azure-identity has been observed doing this). Redirect console.log and
 * console.info to stderr at the earliest possible moment.
 */
function guardStdout(): void {
  const toStderr = (prefix: string) => (...args: unknown[]) => {
    process.stderr.write(
      `[outlook-mcp:${prefix}] ` +
        args
          .map((a) => (typeof a === "string" ? a : JSON.stringify(a)))
          .join(" ") +
        "\n",
    );
  };
  console.log = toStderr("log");
  console.info = toStderr("info");
}
guardStdout();

// Unhandled errors must never reach stdout. Node's default for
// unhandledRejection is to crash with a stderr dump (safe), but we intercept
// to format cleanly and avoid double-logging.
process.on("unhandledRejection", (reason) => {
  process.stderr.write(`[outlook-mcp] unhandledRejection: ${reason}\n`);
});
process.on("uncaughtException", (err) => {
  process.stderr.write(`[outlook-mcp] uncaughtException: ${err?.stack ?? err}\n`);
});

// Tool registry — adding a new tool means: one file in tools/, one line here.
const tools: AnyToolDef[] = [
  toAny(listEmailsTool),
  toAny(readEmailTool),
  toAny(searchEmailsTool),
  toAny(saveAttachmentTool),
  toAny(listCalendarEventsTool),
  toAny(createEventTool),
  toAny(createDraftTool),
];

/**
 * Precompute each tool's JSON schema at startup. zodToJsonSchema is not free
 * and was previously called on every tools/list request. Precomputing also
 * surfaces schema-conversion errors at boot rather than on first list.
 */
interface AdvertisedTool {
  name: string;
  description: string;
  inputSchema: Record<string, unknown>;
}
const advertisedTools: AdvertisedTool[] = tools.map((t) => ({
  name: t.name,
  description: t.description,
  inputSchema: zodToJsonSchema(t.schema, {
    $refStrategy: "none",
  }) as Record<string, unknown>,
}));

const server = new Server(
  { name: "outlook-mcp", version: "0.1.0" },
  { capabilities: { tools: {} } },
);

server.setRequestHandler(ListToolsRequestSchema, async () => {
  return { tools: advertisedTools };
});

server.setRequestHandler(CallToolRequestSchema, async (req) => {
  const tool = tools.find((t) => t.name === req.params.name);
  if (!tool) {
    const result: CallToolResult = {
      isError: true,
      content: [{ type: "text", text: `Unknown tool: ${req.params.name}` }],
    };
    return result;
  }

  const parsed = tool.schema.safeParse(req.params.arguments ?? {});
  if (!parsed.success) {
    const result: CallToolResult = {
      isError: true,
      content: [
        {
          type: "text",
          text: `Invalid arguments for ${tool.name}:\n${parsed.error.toString()}`,
        },
      ],
    };
    return result;
  }

  try {
    const out = await tool.handler(parsed.data);
    // Only attach isError: true on failure — omit on success (per MCP spec,
    // absence signals success more cleanly than `isError: false`).
    const result: CallToolResult = {
      content: [{ type: "text", text: out.text }],
    };
    if (out.isError) result.isError = true;
    return result;
  } catch (err) {
    const msg = err instanceof Error ? err.message : String(err);
    const result: CallToolResult = {
      isError: true,
      content: [{ type: "text", text: `Tool ${tool.name} crashed: ${msg}` }],
    };
    return result;
  }
});

async function main() {
  // Subcommand routing. The default (no args) is MCP stdio server mode.
  //   npx outlook-mcp-local            → stdio server (this file's main)
  //   npx outlook-mcp-local login      → interactive device-code login
  //
  // We treat only explicit first-arg matches as subcommands, so spurious
  // flags from MCP clients (e.g. "--stdio") fall through harmlessly.
  const sub = process.argv[2];
  if (sub === "login") {
    process.exit(await runLogin());
  }
  if (sub === "help" || sub === "--help" || sub === "-h") {
    process.stderr.write(
      "outlook-mcp-local — MCP tools for Outlook / Microsoft 365.\n\n" +
        "Usage:\n" +
        "  outlook-mcp-local          Run as MCP stdio server (default).\n" +
        "  outlook-mcp-local login    Interactive one-time sign-in.\n" +
        "  outlook-mcp-local help     Show this help.\n\n" +
        "Required env: OUTLOOK_MCP_CLIENT_ID (your Azure AD app registration).\n" +
        "See docs/setup-admin.md for the one-time Azure setup.\n",
    );
    process.exit(0);
  }

  const transport = new StdioServerTransport();
  await server.connect(transport);
  process.stderr.write(
    `[outlook-mcp] server started, ${tools.length} tool(s) registered\n`,
  );
}

main().catch((err) => {
  process.stderr.write(`[outlook-mcp] fatal: ${err}\n`);
  process.exit(1);
});
