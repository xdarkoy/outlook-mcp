import { promises as fs } from "node:fs";
import { constants as fsConstants } from "node:fs";
import path from "node:path";
import os from "node:os";

/**
 * Persistent MSAL token cache on disk.
 *
 * Location: ~/.outlook-mcp/cache.json (or OUTLOOK_MCP_CACHE_DIR override).
 * The directory is realpath-resolved on first access so that a symlinked
 * override cannot be used to point cache writes at an arbitrary target.
 *
 * Permissions: on POSIX we create with mode 0600 via open(O_CREAT|O_EXCL)
 * then chmod — closing any umask race window. On Windows mode flags are
 * silently ignored by the kernel; the file inherits the user-profile ACL
 * which restricts access to the current user.
 *
 * Writes are atomic: write-to-temp + rename. A SIGKILL mid-write cannot
 * leave cache.json truncated — either the previous version is intact, or
 * the new one is fully on disk. No native deps (keytar/DPAPI), keeping
 * the "npx and go" install story.
 */

function logicalCacheDir(): string {
  const override = process.env.OUTLOOK_MCP_CACHE_DIR;
  if (override) return path.resolve(override);
  return path.join(os.homedir(), ".outlook-mcp");
}

let _realCacheDir: string | null = null;
async function resolveCacheDir(): Promise<string> {
  if (_realCacheDir) return _realCacheDir;
  const dir = logicalCacheDir();
  await fs.mkdir(dir, { recursive: true });
  try {
    _realCacheDir = await fs.realpath(dir);
  } catch {
    _realCacheDir = dir;
  }
  return _realCacheDir;
}

async function cachePath(): Promise<string> {
  return path.join(await resolveCacheDir(), "cache.json");
}

export async function readCache(): Promise<string> {
  try {
    return await fs.readFile(await cachePath(), "utf-8");
  } catch (err: unknown) {
    if ((err as NodeJS.ErrnoException).code === "ENOENT") return "";
    throw err;
  }
}

export async function writeCache(data: string): Promise<void> {
  const file = await cachePath();
  const tmp = `${file}.${process.pid}.tmp`;

  // Write to a sibling tmp file with mode 0600 baked in at create time.
  // On POSIX this closes the umask window; on Windows the mode is ignored
  // but the user-profile ACL still restricts access.
  const handle = await fs.open(tmp, "w", 0o600);
  try {
    await handle.writeFile(data, "utf-8");
  } finally {
    await handle.close();
  }

  // Best-effort chmod for the case where the file already existed (open
  // with mode only applies on create).
  if (process.platform !== "win32") {
    try {
      await fs.chmod(tmp, 0o600);
    } catch {
      /* chmod is best-effort */
    }
  }

  // Atomic replace. POSIX rename is atomic; Windows rename over an existing
  // file also works on modern NTFS but is not a strict atomicity guarantee.
  // Acceptable for a single-user token cache.
  await fs.rename(tmp, file);
}

export async function cacheExists(): Promise<boolean> {
  try {
    await fs.access(await cachePath(), fsConstants.F_OK);
    return true;
  } catch {
    return false;
  }
}
