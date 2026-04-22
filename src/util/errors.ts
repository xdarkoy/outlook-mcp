/**
 * Translate opaque Graph / MSAL / filesystem errors into messages that are
 * actionable inside an LLM tool-calling loop. An LLM that gets a clear error
 * can self-correct; one that gets `Error: Request failed with status 403`
 * usually cannot.
 */

import { LoginRequiredError } from "../auth/msal.js";

export class ToolError extends Error {
  constructor(message: string, public readonly hint?: string) {
    super(message);
    this.name = "ToolError";
  }
}

interface GraphErrorShape {
  statusCode?: number;
  code?: string;
  body?: string;
  message?: string;
}

/**
 * The Graph SDK surfaces the most useful error text inside err.body as a JSON
 * string (e.g. '{"error":{"code":"...","message":"<the real message>"}}').
 * err.message is often just "Bad Request". Extract the inner message when we
 * can.
 */
function bestErrorMessage(e: GraphErrorShape): string {
  if (e.body) {
    try {
      const parsed = JSON.parse(e.body) as { error?: { message?: string } };
      if (parsed?.error?.message) return parsed.error.message;
    } catch {
      // body wasn't JSON — fall through.
    }
  }
  return e.message ?? "Unknown Graph error.";
}

export function explainGraphError(err: unknown): ToolError {
  // Pre-auth hasn't happened yet — bubble the actionable instruction.
  if (err instanceof LoginRequiredError) {
    return new ToolError(err.message);
  }

  const e = err as GraphErrorShape;
  const status = e?.statusCode;
  const code = e?.code;
  const raw = bestErrorMessage(e);

  if (status === 401 || code === "InvalidAuthenticationToken") {
    return new ToolError(
      "Authentication failed or expired.",
      "Delete the local token cache (~/.outlook-mcp/cache.json) and restart the MCP server to trigger a fresh sign-in."
    );
  }
  if (status === 403 || code === "ErrorAccessDenied" || code === "Forbidden") {
    return new ToolError(
      "Permission denied by Microsoft Graph.",
      "Your Azure AD app registration is missing a required scope, or the tenant admin has not granted consent. See docs/setup-admin.md."
    );
  }
  if (status === 404 || code === "ErrorItemNotFound") {
    return new ToolError(
      "The requested item was not found.",
      "The message, attachment, or calendar event ID may be stale. Re-run list_emails or search_emails to refresh IDs."
    );
  }
  if (status === 429) {
    return new ToolError(
      "Microsoft Graph rate limit hit.",
      "Wait a few seconds and retry. For bulk work, reduce batch size or add delays."
    );
  }
  return new ToolError(`Graph request failed: ${raw}`);
}
