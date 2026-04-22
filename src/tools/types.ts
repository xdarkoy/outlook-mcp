import type { ZodTypeAny, z } from "zod";

/**
 * Every tool module exports a single object of this shape. `index.ts` builds
 * the MCP tool registry from an array of these — adding a new tool is one
 * file + one line in the registry.
 *
 * The erased `AnyToolDef` form is what the registry stores; call sites can
 * still use `ToolDef<MySchema>` for their own type safety.
 */
export interface ToolDef<S extends ZodTypeAny> {
  name: string;
  description: string;
  schema: S;
  handler: (args: z.infer<S>) => Promise<ToolResult>;
}

export type AnyToolDef = {
  name: string;
  description: string;
  schema: ZodTypeAny;
  handler: (args: unknown) => Promise<ToolResult>;
};

/**
 * Narrow a typed ToolDef<S> to the erased AnyToolDef registry entry without
 * an ugly `as unknown as` double-cast at each call site. Safe because the
 * handler only ever receives values that were freshly parsed by `schema`.
 */
export function toAny<S extends ZodTypeAny>(t: ToolDef<S>): AnyToolDef {
  return t as unknown as AnyToolDef;
}

export interface ToolResult {
  /** Primary payload. LLMs parse this as text; keep it compact + structured. */
  text: string;
  /** If true, surface this result to the caller as an error (MCP isError=true). */
  isError?: boolean;
}

export function ok(payload: unknown): ToolResult {
  if (payload === undefined) return { text: "ok" };
  return {
    text: typeof payload === "string" ? payload : JSON.stringify(payload, null, 2),
  };
}

export function fail(message: string, hint?: string): ToolResult {
  return {
    text: hint ? `${message}\n\nHint: ${hint}` : message,
    isError: true,
  };
}
