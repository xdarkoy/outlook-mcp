import { getSignedInAccount } from "./msal.js";

/**
 * Detect whether the signed-in user is a personal Microsoft Account (MSA)
 * or a work/school account (AAD).
 *
 * Why it matters:
 *   - POST /search/query (Microsoft Search) is AAD-only; MSA returns
 *     "This API is not supported for MSA accounts".
 *   - /me/messages?$search="..." works on BOTH.
 *   - Several other endpoints and headers also differ.
 *
 * Detection: MSAL exposes the signed-in account's tenantId directly from
 * cached account info. All MSA-issued tokens come from the special
 * consumers tenant with id 9188040d-6c67-4c5b-b112-36a304b66dad; any
 * other tenantId is an AAD work/school tenant.
 *
 * NOTE: we intentionally do NOT decode the access token as a JWT here —
 * MSA-issued Graph access tokens are opaque (not JWTs). Only AAD ones are.
 * MSAL's `account.tenantId` is the single source of truth that works for
 * both.
 *
 * Cached per process: tenantId never changes for a given signed-in account.
 */

const MSA_TENANT_ID = "9188040d-6c67-4c5b-b112-36a304b66dad";

export type AccountType = "msa" | "aad";

let _cachedType: AccountType | null = null;

export async function getAccountType(): Promise<AccountType> {
  if (_cachedType) return _cachedType;
  const account = await getSignedInAccount();
  _cachedType = classifyAccount(account);
  return _cachedType;
}

/**
 * Pure classifier, exported so tests don't need to mock MSAL.
 */
export function classifyAccount(
  account: { tenantId?: string } | null | undefined,
): AccountType {
  return account?.tenantId === MSA_TENANT_ID ? "msa" : "aad";
}

export function _resetAccountTypeCache(): void {
  _cachedType = null;
}
