/**
 * RBACContext: Determines user role and permissions.
 *
 * - Reads Azure AD group memberships via Graph using the MSAL token.
 * - Environment-based roles: Super admins, admins, and managers can be configured in .env
 * - Exposes simple booleans (canSeeAdmin, canEditAdmin) for component gating.
 */
import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';
import { useAuth } from './AuthContext';
import { fetchUserGroups } from '../services/graphService';
import { getRoles, type DbRole } from '../services/dbService';
import { getPermissionCatalog, getEffectivePermissions } from '../services/rbacService';
import { isSQLiteEnabled } from '../utils/runtimeConfig';

type RBAC = { 
  role: 'SuperAdmin'|'Admin'|'Manager'|'Employee', 
  canSeeAdmin: boolean, 
  canEditAdmin: boolean, 
  isSuperAdmin: boolean,
  perms: Record<string, boolean>
};
const defaultRBAC: RBAC = { role: 'Employee', canSeeAdmin: false, canEditAdmin: false, isSuperAdmin: false, perms: {} };
export const RBACContext = createContext<RBAC>(defaultRBAC);
export const useRBAC = () => useContext(RBACContext);

const config = {
  Admin: { groups: ['Sunbeth-Portal-Admins','HR-Managers'] },
  Manager: { groups: ['Sunbeth-Dept-Managers'] },
  Employee: { groups: ['Sunbeth-Employees'] }
};

// Environment-based role configuration
const getEnvEmails = (envVar: string): string[] => {
  const emails = process.env[envVar];
  return emails ? emails.split(',').map(email => email.trim().toLowerCase()).filter(email => email.length > 0) : [];
};

const SUPER_ADMINS = getEnvEmails('REACT_APP_SUPER_ADMINS');
const ADMINS = getEnvEmails('REACT_APP_ADMINS');
const MANAGERS = getEnvEmails('REACT_APP_MANAGERS');

// DB-backed role caches (populated at runtime when SQLite API is enabled)
let DB_SUPER_ADMINS: string[] = [];
let DB_ADMINS: string[] = [];
let DB_MANAGERS: string[] = [];

// Helper function to determine role from email and groups
const determineRole = (userEmail: string, groups: string[]): RBAC['role'] => {
  const normalizedEmail = userEmail.toLowerCase();
  
  // DB roles take precedence if present, else fall back to environment lists
  if (DB_SUPER_ADMINS.includes(normalizedEmail) || SUPER_ADMINS.includes(normalizedEmail)) return 'SuperAdmin';
  if (DB_ADMINS.includes(normalizedEmail) || ADMINS.includes(normalizedEmail)) return 'Admin';
  if (DB_MANAGERS.includes(normalizedEmail) || MANAGERS.includes(normalizedEmail)) return 'Manager';
  
  // Check group-based roles
  if (groups.some(g => config.Admin.groups.includes(g))) return 'Admin';
  if (groups.some(g => config.Manager.groups.includes(g))) return 'Manager';
  
  return 'Employee';
};

export const RBACProvider: React.FC<{children: React.ReactNode}> = ({ children }) => {
  const { token, account } = useAuth();
  const [role, setRole] = useState<RBAC['role']>('Employee');
  const [perms, setPerms] = useState<Record<string, boolean>>({});

  // Fetch DB roles (if enabled) and then groups when token is available
  useEffect(() => {
    if (!token || !account) { setRole('Employee'); return; }
    let active = true;
    const userEmail = account.username || '';

    (async () => {
      try {
        if (isSQLiteEnabled()) {
          const roles = await getRoles();
          if (!active) return;
          DB_SUPER_ADMINS = roles.filter(r => r.role === 'SuperAdmin').map(r => r.email.toLowerCase());
          DB_ADMINS = roles.filter(r => r.role === 'Admin').map(r => r.email.toLowerCase());
          DB_MANAGERS = roles.filter(r => r.role === 'Manager').map(r => r.email.toLowerCase());
        } else {
          DB_SUPER_ADMINS = [];
          DB_ADMINS = [];
          DB_MANAGERS = [];
        }
      } catch {
        // ignore and use env only
      }

      // After DB roles are loaded, check immediate role from DB/env
      if (userEmail) {
        const immediate = determineRole(userEmail, []);
        if (immediate !== 'Employee') {
          if (active) setRole(immediate);
          // Still fall through to group-based for potential elevation to Admin via groups when neither DB nor env set
        }
      }

      // Then check group-based roles as final fallback
      try {
        const groups = await fetchUserGroups(token);
        if (!active) return;
        const finalRole = determineRole(userEmail, groups);
        setRole(finalRole);
      } catch {
        if (active && !userEmail) setRole('Employee');
      }
    })();
    return () => { active = false; };
  }, [token, account]);

  // Load effective permissions (from server) or fallback to defaults by role
  useEffect(() => {
    let cancelled = false;
    (async () => {
      const email = account?.username || '';
      const sqlite = isSQLiteEnabled();
      try {
        if (sqlite && email) {
          const eff = await getEffectivePermissions(email);
          if (!cancelled) setPerms(eff.permissions || {});
          return;
        }
      } catch {}
      // Fallback defaults by role
      try {
        const catalog = await getPermissionCatalog().catch(() => []);
        const keys = Array.isArray(catalog) && catalog.length > 0 ? catalog.map(p => p.key) : [
          'viewAdmin','manageSettings','viewDebugLogs','exportAnalytics','viewAnalytics','createBatch','editBatch','deleteBatch','manageRecipients','manageDocuments','sendNotifications','uploadDocuments','manageBusinesses','manageRoles','managePermissions'
        ];
        const allTrue = Object.fromEntries(keys.map(k => [k, true]));
        const allFalse = Object.fromEntries(keys.map(k => [k, false]));
        let mapping = allFalse;
        if (role === 'SuperAdmin') mapping = allTrue;
        else if (role === 'Admin') mapping = { ...allTrue, deleteBatch: true };
        else if (role === 'Manager') {
          mapping = { ...allFalse };
          const allow = ['viewAdmin','viewAnalytics','exportAnalytics','createBatch','editBatch','manageRecipients','manageDocuments','sendNotifications','uploadDocuments'];
          for (const k of allow) mapping[k] = true;
        }
        if (!cancelled) setPerms(mapping);
      } catch {}
    })();
    return () => { cancelled = true; };
  }, [role, account]);

  const value: RBAC = useMemo(() => ({
    role,
    canSeeAdmin: role === 'SuperAdmin' || role === 'Admin' || role === 'Manager',
    canEditAdmin: role === 'SuperAdmin' || role === 'Admin',
    isSuperAdmin: role === 'SuperAdmin',
    perms
  }), [role, perms]);

  return <RBACContext.Provider value={value}>{children}</RBACContext.Provider>;
};
