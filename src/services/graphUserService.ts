/**
 * Microsoft Graph User and Group Management Service
 * Provides comprehensive user/group operations for admin batch assignment
 */
import { info, warn, error as logError } from '../diagnostics/logger';

export interface GraphUser {
  id: string;
  displayName: string;
  userPrincipalName: string;
  mail?: string;
  jobTitle?: string;
  department?: string;
  officeLocation?: string;
  businessPhones?: string[];
  mobilePhone?: string;
}

export interface GraphGroup {
  id: string;
  displayName: string;
  description?: string;
  mail?: string;
  groupTypes: string[];
  memberCount?: number;
}

export interface UserSearchFilters {
  department?: string;
  jobTitle?: string;
  location?: string;
  search?: string;
}

/**
 * Fetches all users in the organization with optional filtering
 */
export const getUsers = async (token: string, filters?: UserSearchFilters): Promise<GraphUser[]> => {
  try {
    // Base query with reasonable page size; we will follow @odata.nextLink for pagination
    const base = 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName,mail,jobTitle,department,officeLocation,businessPhones,mobilePhone&$orderby=displayName&$top=200';
    const escapeOData = (s: string) => s.replace(/'/g, "''");
    const filterStr = filters?.search ? `(startswith(displayName,'${escapeOData(filters.search)}') or startswith(userPrincipalName,'${escapeOData(filters.search)}') or startswith(mail,'${escapeOData(filters.search)}'))` : '';
    let url = filterStr ? `${base}&$filter=${filterStr}` : base;

    const users: GraphUser[] = [];
    while (url) {
      const response: Response = await fetch(url, {
        headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'ConsistencyLevel': 'eventual' }
      });

      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new Error(`Graph API error: ${response.status} ${response.statusText}${text ? ' — ' + text : ''}`);
      }

      const data: any = await response.json();
      const page = Array.isArray(data.value) ? (data.value as GraphUser[]) : [];
      users.push(...page);
      url = data['@odata.nextLink'] || '';
    }

    // Apply client-side filters for better UX
    let filtered = users;
    if (filters?.department) {
      filtered = filtered.filter(u => (u.department || '').toLowerCase().includes(filters.department!.toLowerCase()));
    }
    if (filters?.jobTitle) {
      filtered = filtered.filter(u => (u.jobTitle || '').toLowerCase().includes(filters.jobTitle!.toLowerCase()));
    }
    if (filters?.location) {
      filtered = filtered.filter(u => (u.officeLocation || '').toLowerCase().includes(filters.location!.toLowerCase()));
    }

    info('graphUserService: fetched users', { count: filtered.length, filters });
    return filtered;
  } catch (e) {
    logError('graphUserService: failed to fetch users', e);
    throw e;
  }
};

/**
 * Fetches all groups in the organization
 */
export const getGroups = async (token: string): Promise<GraphGroup[]> => {
  try {
    let url: string | null = 'https://graph.microsoft.com/v1.0/groups?$select=id,displayName,description,mail,groupTypes&$orderby=displayName&$top=200';
    const groups: GraphGroup[] = [];
    while (url) {
      const response: Response = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'application/json', 'ConsistencyLevel': 'eventual' } });
      if (!response.ok) {
        const text = await response.text().catch(() => '');
        throw new Error(`Graph API error: ${response.status} ${response.statusText}${text ? ' — ' + text : ''}`);
      }
      const data: any = await response.json();
      const page = Array.isArray(data.value) ? (data.value as GraphGroup[]) : [];
      groups.push(...page);
      url = data['@odata.nextLink'] || null;
    }

    // Fetch member counts for each group (in batches to avoid rate limits)
    const groupsWithCounts = await Promise.all(
      groups.map(async (group) => {
        try {
          const membersResponse = await fetch(
            `https://graph.microsoft.com/v1.0/groups/${group.id}/members/$count`,
            { headers: { Authorization: `Bearer ${token}`, 'ConsistencyLevel': 'eventual' } }
          );
          if (membersResponse.ok) {
            group.memberCount = parseInt(await membersResponse.text());
          }
        } catch (e) {
          // Ignore count errors
        }
        return group;
      })
    );

    info('graphUserService: fetched groups', { count: groupsWithCounts.length });
    return groupsWithCounts;
  } catch (e) {
    logError('graphUserService: failed to fetch groups', e);
    throw e;
  }
};

/** Get total count of users (approximate). */
export const getUsersCount = async (token: string): Promise<number> => {
  const url = 'https://graph.microsoft.com/v1.0/users/$count?$filter=accountEnabled eq true';
  const res = await fetch(url, { headers: { Authorization: `Bearer ${token}`, Accept: 'text/plain', 'ConsistencyLevel': 'eventual' } });
  if (!res.ok) {
    const text = await res.text().catch(() => '');
    throw new Error(`Graph API error (users count): ${res.status} ${res.statusText}${text ? ' — ' + text : ''}`);
  }
  const text = await res.text();
  const n = parseInt(text, 10);
  info('graphUserService: users count', { count: n });
  return isNaN(n) ? 0 : n;
};

/**
 * Gets group members
 */
export const getGroupMembers = async (token: string, groupId: string): Promise<GraphUser[]> => {
  try {
    const all: GraphUser[] = [];
    let url: string | null = `https://graph.microsoft.com/v1.0/groups/${groupId}/members?$select=id,displayName,userPrincipalName,mail,jobTitle,department&$top=200`;
    while (url) {
      const response: Response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
      if (!response.ok) {
        throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
      }
      const data: any = await response.json();
      const pageMembers = (data.value || []).filter((m: any) => m['@odata.type'] === '#microsoft.graph.user') as GraphUser[];
      all.push(...pageMembers);
      url = data['@odata.nextLink'] || null;
    }
    info('graphUserService: fetched group members', { groupId, count: all.length });
    return all;
  } catch (e) {
    logError('graphUserService: failed to fetch group members', e);
    throw e;
  }
};

/**
 * Gets organizational hierarchy (manager/direct reports)
 */
export const getUserHierarchy = async (token: string, userId: string): Promise<{ manager?: GraphUser; directReports: GraphUser[] }> => {
  try {
    const [managerResponse, reportsResponse] = await Promise.all([
      fetch(`https://graph.microsoft.com/v1.0/users/${userId}/manager?$select=id,displayName,userPrincipalName,mail,jobTitle`, {
        headers: { Authorization: `Bearer ${token}` }
      }),
      fetch(`https://graph.microsoft.com/v1.0/users/${userId}/directReports?$select=id,displayName,userPrincipalName,mail,jobTitle`, {
        headers: { Authorization: `Bearer ${token}` }
      })
    ]);

    const manager = managerResponse.ok ? await managerResponse.json() : null;
    const reportsData = reportsResponse.ok ? await reportsResponse.json() : { value: [] };

    return {
      manager: manager as GraphUser,
      directReports: reportsData.value as GraphUser[]
    };
  } catch (e) {
    logError('graphUserService: failed to fetch user hierarchy', e);
    return { directReports: [] };
  }
};

/**
 * Gets departments and job titles for filtering
 */
export const getOrganizationStructure = async (token: string): Promise<{ departments: string[]; jobTitles: string[]; locations: string[] }> => {
  try {
    const response = await fetch(
      'https://graph.microsoft.com/v1.0/users?$select=department,jobTitle,officeLocation&$top=999',
      { headers: { Authorization: `Bearer ${token}` } }
    );

    if (!response.ok) {
      throw new Error(`Graph API error: ${response.status} ${response.statusText}`);
    }

    const data = await response.json();
    const users = data.value;

    const departments = [...new Set(users.map((u: any) => u.department).filter(Boolean))].sort() as string[];
    const jobTitles = [...new Set(users.map((u: any) => u.jobTitle).filter(Boolean))].sort() as string[];
    const locations = [...new Set(users.map((u: any) => u.officeLocation).filter(Boolean))].sort() as string[];

    info('graphUserService: fetched organization structure', { departments: departments.length, jobTitles: jobTitles.length, locations: locations.length });
    return { departments, jobTitles, locations };
  } catch (e) {
    logError('graphUserService: failed to fetch organization structure', e);
    return { departments: [], jobTitles: [], locations: [] };
  }
};