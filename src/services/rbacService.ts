import { getApiBase } from '../utils/runtimeConfig';

export type PermissionDef = { key: string; label: string; description?: string; category?: string };
export type RolePermission = { role: string; permKey: string; value: boolean };
export type UserPermission = { email: string; permKey: string; value: boolean };

const api = () => (getApiBase() as string).replace(/\/$/, '');

export async function getPermissionCatalog(): Promise<PermissionDef[]> {
  const res = await fetch(`${api()}/api/rbac/permissions`, { cache: 'no-store' });
  if (!res.ok) throw new Error('perm_catalog_failed');
  return res.json();
}

export async function getRolePermissions(role?: string): Promise<RolePermission[]> {
  const url = role ? `${api()}/api/rbac/role-permissions?role=${encodeURIComponent(role)}` : `${api()}/api/rbac/role-permissions`;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error('role_perms_failed');
  return res.json();
}

export async function setRolePermissions(role: string, mapping: Record<string, boolean>): Promise<void> {
  const res = await fetch(`${api()}/api/rbac/role-permissions`, {
    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ role, mapping })
  });
  if (!res.ok) throw new Error('role_perms_update_failed');
}

export async function getUserPermissions(email?: string): Promise<UserPermission[]> {
  const url = email ? `${api()}/api/rbac/user-permissions?email=${encodeURIComponent(email)}` : `${api()}/api/rbac/user-permissions`;
  const res = await fetch(url, { cache: 'no-store' });
  if (!res.ok) throw new Error('user_perms_failed');
  return res.json();
}

export async function setUserPermissions(email: string, mapping: Record<string, boolean>): Promise<void> {
  const res = await fetch(`${api()}/api/rbac/user-permissions`, {
    method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ email, mapping })
  });
  if (!res.ok) throw new Error('user_perms_update_failed');
}

export async function getEffectivePermissions(email: string): Promise<{ roles: string[]; permissions: Record<string, boolean>}> {
  const res = await fetch(`${api()}/api/rbac/effective?email=${encodeURIComponent(email)}`, { cache: 'no-store' });
  if (!res.ok) throw new Error('effective_perms_failed');
  return res.json();
}
