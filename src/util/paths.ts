import path from "node:path";
import os from "node:os";
import { promises as fs } from "node:fs";
import { ToolError } from "./errors.js";

/**
 * Path-safety utilities for save_attachment.
 *
 * Threat model: the LLM is untrusted input. Never let it write outside the
 * allowed directory. Never overwrite existing files silently.
 *
 * The allowed directory is fixed at server startup via OUTLOOK_MCP_ALLOWED_DIR.
 * Default: ~/Downloads/outlook-mcp/. This is deliberately not LLM-controllable —
 * even a tool that lets the LLM specify a path would still resolve against
 * this root and reject escapes.
 */

const WINDOWS_RESERVED = new Set([
  "con", "prn", "aux", "nul",
  "com1", "com2", "com3", "com4", "com5", "com6", "com7", "com8", "com9",
  "lpt1", "lpt2", "lpt3", "lpt4", "lpt5", "lpt6", "lpt7", "lpt8", "lpt9",
]);

export function allowedDir(): string {
  const override = process.env.OUTLOOK_MCP_ALLOWED_DIR;
  if (override) return path.resolve(override);
  return path.join(os.homedir(), "Downloads", "outlook-mcp");
}

/**
 * Resolve symlinks in the allowed root BEFORE prefix-checking the write
 * target. Without this, a user (or attacker who controls the env) could
 * point OUTLOOK_MCP_ALLOWED_DIR at a symlink whose real target is /etc or
 * C:\Windows: `candidate.startsWith(rootWithSep)` would pass in path-space
 * but the actual write would land outside. Cached: realpath() is a syscall
 * and this is called on every save_attachment.
 */
let _realAllowedDir: string | null = null;
async function resolveAllowedDir(): Promise<string> {
  if (_realAllowedDir) return _realAllowedDir;
  const logical = allowedDir();
  await fs.mkdir(logical, { recursive: true });
  try {
    _realAllowedDir = await fs.realpath(logical);
  } catch {
    // realpath may fail on brand-new dirs briefly on some filesystems —
    // fall back to the logical path; next call will succeed.
    _realAllowedDir = logical;
  }
  return _realAllowedDir;
}

export function sanitizeFilename(name: string): string {
  if (!name) throw new ToolError("Empty filename.");

  // Strip any path separators — we only accept bare filenames.
  // The `:` replacement in the control-char regex below is LOAD-BEARING on
  // Windows: it neutralizes both drive-relative paths ("C:foo") and NTFS
  // Alternate Data Streams ("file.txt:stream"). Do not remove it.
  const bare = name.replace(/[\\/]/g, "_");

  // Strip control chars and the classic unsafe set (<, >, :, ", |, ?, *).
  // eslint-disable-next-line no-control-regex
  let cleaned = bare.replace(/[\x00-\x1f<>:"|?*]/g, "_");

  // Strip Unicode formatting characters that can be used to visually
  // disguise filenames:
  //   U+200B..U+200D  zero-width space/joiner/non-joiner
  //   U+200E..U+200F  LTR / RTL marks
  //   U+202A..U+202E  embedding / override (RLO attack: "invoice‮fdp.exe"
  //                   displays as "invoiceexe.pdf")
  //   U+2066..U+2069  isolate controls
  //   U+FEFF          BOM / zero-width no-break space
  cleaned = cleaned.replace(
    /[\u200B-\u200F\u202A-\u202E\u2066-\u2069\uFEFF]/g,
    "",
  );

  // Windows strips trailing dots and spaces from filenames ("CON." → "CON"),
  // so a reserved-name check BEFORE this step would miss "CON " or "NUL.".
  cleaned = cleaned.trim().replace(/[. ]+$/g, "");

  if (!cleaned || cleaned === "." || cleaned === "..") {
    throw new ToolError(`Invalid filename '${name}'.`);
  }

  // Reserved name check runs on the sanitized form.
  const stem = cleaned.split(".")[0]?.toLowerCase().trim() ?? "";
  if (WINDOWS_RESERVED.has(stem)) {
    throw new ToolError(`Filename '${name}' is a reserved name on Windows.`);
  }

  return cleaned;
}

/**
 * Resolve `filename` inside the allowed directory and verify the final
 * absolute path is still under the root. Returns a candidate path whose
 * directory is the realpath-resolved root. The caller is responsible for
 * using an atomic O_EXCL create when writing — see saveAttachment.ts.
 *
 * We no longer probe-and-return a non-colliding name here: the probe is
 * TOCTOU, because another write could create the file between our
 * existence check and the writeFile. Collision handling moves to the
 * writer, which retries with an incremented suffix on EEXIST.
 */
export async function resolveWritePath(filename: string): Promise<string> {
  const root = await resolveAllowedDir();
  const safe = sanitizeFilename(filename);
  const candidate = path.resolve(root, safe);

  const rootWithSep = root.endsWith(path.sep) ? root : root + path.sep;
  if (!candidate.startsWith(rootWithSep) && candidate !== root) {
    throw new ToolError(
      `Resolved path escapes the allowed directory (${root}).`,
      "This is almost certainly a path-traversal attempt. Aborting.",
    );
  }
  return candidate;
}

/**
 * Generate the next collision-suffixed name for a given path.
 * "invoice.pdf" → "invoice (2).pdf" → "invoice (3).pdf" …
 */
export function suffixedPath(original: string, n: number): string {
  const parsed = path.parse(original);
  return path.join(parsed.dir, `${parsed.name} (${n})${parsed.ext}`);
}
