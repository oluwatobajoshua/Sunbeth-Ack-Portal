/**
 * RBACContext: Determines user role and permissions.
 *
 * - Live mode: Reads Azure AD group memberships via Graph using the MSAL token.
 * - Environment-based roles: Super admins, admins, and managers can be configured in .env
 * - Mock mode: Reads groups from localStorage (mock_user_groups) to enable UI testing.
 * - Exposes simple booleans (canSeeAdmin, canEditAdmin) for component gating.
 */
import React, { createContext, useContext, useEffect, useMemo, useState } from 'react';
import { useAuth } from './AuthContext';
import { fetchUserGroups } from '../services/graphService';
import { useRuntimeMock } from '../utils/runtimeMock';

type RBAC = { role: 'SuperAdmin'|'Admin'|'Manager'|'Employee', canSeeAdmin: boolean, canEditAdmin: boolean, isSuperAdmin: boolean };
const defaultRBAC: RBAC = { role: 'Employee', canSeeAdmin: false, canEditAdmin: false, isSuperAdmin: false };
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

// Helper function to determine role from email and groups
const determineRole = (userEmail: string, groups: string[]): RBAC['role'] => {
  const normalizedEmail = userEmail.toLowerCase();
  
  // Check environment-based roles first (highest priority)
  if (SUPER_ADMINS.includes(normalizedEmail)) return 'SuperAdmin';
  if (ADMINS.includes(normalizedEmail)) return 'Admin';
  if (MANAGERS.includes(normalizedEmail)) return 'Manager';
  
  // Check group-based roles
  if (groups.some(g => config.Admin.groups.includes(g))) return 'Admin';
  if (groups.some(g => config.Manager.groups.includes(g))) return 'Manager';
  
  return 'Employee';
};

export const RBACProvider: React.FC<{children: React.ReactNode}> = ({ children }) => {
  const { token, account } = useAuth();
  const runtimeMock = useRuntimeMock();
  const [role, setRole] = useState<RBAC['role']>('Employee');

  const computeFromMock = () => {
    try {
      const raw = localStorage.getItem('mock_user_groups');
      const groups: string[] = raw ? JSON.parse(raw) : [];
      // In mock mode, check environment variables first if user email is available
      if (account?.username) {
        const envRole = determineRole(account.username, groups);
        if (envRole !== 'Employee') return envRole;
      }
      if (groups.some(g => config.Admin.groups.includes(g))) return 'Admin' as RBAC['role'];
      if (groups.some(g => config.Manager.groups.includes(g))) return 'Manager' as RBAC['role'];
    } catch (e) { /* fallthrough */ }
    return 'Employee' as RBAC['role'];
  };

  // Mock mode: derive role from localStorage and respond immediately to changes
  useEffect(() => {
    if (!runtimeMock) return;
    const setFromMock = () => setRole(computeFromMock());
    setFromMock();
    const onStorage = (e: StorageEvent) => {
      if (e.key && e.key !== 'mock_user_groups') return;
      setFromMock();
      try { window.dispatchEvent(new CustomEvent('sunbeth:toast', { detail: { message: `Role set to ${computeFromMock()} (mock)` } })); } catch {}
    };
    const onRoleChange = () => setFromMock();
    window.addEventListener('storage', onStorage as EventListener);
    window.addEventListener('sunbeth:roleChange', onRoleChange as EventListener);
    return () => {
      window.removeEventListener('storage', onStorage as EventListener);
      window.removeEventListener('sunbeth:roleChange', onRoleChange as EventListener);
    };
  }, [runtimeMock]);

  // Live mode: fetch groups from Graph when token is available
  useEffect(() => {
    if (runtimeMock) return;
    if (!token || !account) { setRole('Employee'); return; }
    let active = true;
    
    // First check environment-based roles
    const userEmail = account.username;
    if (userEmail) {
      const envRole = determineRole(userEmail, []);
      if (envRole !== 'Employee') {
        if (active) setRole(envRole);
        return;
      }
    }
    
    // Then check group-based roles
    fetchUserGroups(token).then(groups => {
      if (!active) return;
      const finalRole = determineRole(userEmail || '', groups);
      setRole(finalRole);
    }).catch(() => { if (active) setRole('Employee'); });
    return () => { active = false; };
  }, [runtimeMock, token, account]);

  const value: RBAC = useMemo(() => ({
    role,
    canSeeAdmin: role === 'SuperAdmin' || role === 'Admin' || role === 'Manager',
    canEditAdmin: role === 'SuperAdmin' || role === 'Admin',
    isSuperAdmin: role === 'SuperAdmin'
  }), [role]);

  return <RBACContext.Provider value={value}>{children}</RBACContext.Provider>;
};
