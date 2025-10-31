import { useRBAC } from '../context/RBACContext';

/**
 * usePerm: simple helper to test a permission key.
 * - Returns true for Super Admins unconditionally
 * - Otherwise returns perms[key] boolean (false if missing)
 */
export function usePerm(key: string): boolean {
  const { isSuperAdmin, perms } = useRBAC();
  if (isSuperAdmin) return true;
  return !!(perms && perms[key]);
}
