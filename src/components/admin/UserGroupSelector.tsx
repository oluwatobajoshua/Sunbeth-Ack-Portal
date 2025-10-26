import React, { useEffect, useState } from 'react';
import { useAuth as useAuthCtx } from '../../context/AuthContext';
import {
  GraphUser,
  GraphGroup,
  getUsers,
  getGroups,
  getOrganizationStructure,
  UserSearchFilters,
  getGroupMembers
} from '../../services/graphUserService';
import { showToast } from '../../utils/alerts';

export const UserGroupSelector: React.FC<{ onSelectionChange: (selection: any) => void }> = ({ onSelectionChange }) => {
  const { getToken, login, account } = useAuthCtx();
  const [loading, setLoading] = useState(false);
  const [hadError, setHadError] = useState<string | null>(null);
  const [tab, setTab] = useState<'users' | 'groups' | 'structure'>('users');
  const [users, setUsers] = useState<GraphUser[]>([]);
  const [groups, setGroups] = useState<GraphGroup[]>([]);
  const [orgStructure, setOrgStructure] = useState<{ departments: string[]; jobTitles: string[]; locations: string[] }>({ departments: [], jobTitles: [], locations: [] });
  const [filters, setFilters] = useState<UserSearchFilters>({});
  const [localSearch, setLocalSearch] = useState<string>('');
  const [selectedUsers, setSelectedUsers] = useState<Set<string>>(new Set());
  const [selectedGroups, setSelectedGroups] = useState<Set<string>>(new Set());
  const [usersPage, setUsersPage] = useState<number>(1);
  const [groupsPage, setGroupsPage] = useState<number>(1);
  const [groupSearch, setGroupSearch] = useState<string>('');
  const pageSize = 50;

  const loadData = async () => {
    setLoading(true);
    setHadError(null);
    try {
      const token = await getToken(['User.Read.All', 'Group.Read.All']);
      if (!token) throw new Error('No token available');

      const [usersData, groupsData, structureData] = await Promise.all([
        getUsers(token, filters),
        getGroups(token),
        getOrganizationStructure(token)
      ]);

      setUsers(usersData);
      setGroups(groupsData);
      setOrgStructure(structureData);
    } catch (error: any) {
      console.error('Failed to load user/group data:', error);
      const msg = typeof error?.message === 'string' ? error.message : '';
      const hint = msg.includes('No active account')
        ? 'Please sign in to continue.'
        : 'Ask your admin to grant Microsoft Graph permissions (User.Read.All and Group.Read.All) to this app.';
      setHadError(`${msg || 'Failed to load user data.'} ${hint}`.trim());
      showToast(`Failed to load user data. ${hint}`, 'error');
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    void loadData();
    setUsersPage(1);
    setGroupsPage(1);
  }, [filters]);

  // Debounce search input before applying to filters
  useEffect(() => {
    const h = setTimeout(() => {
      setFilters(prev => ({ ...prev, search: localSearch || undefined }));
    }, 300);
    return () => clearTimeout(h);
  }, [localSearch]);

  useEffect(() => {
    onSelectionChange({
      users: Array.from(selectedUsers).map(id => users.find(u => u.id === id)!).filter(Boolean),
      groups: Array.from(selectedGroups).map(id => groups.find(g => g.id === id)!).filter(Boolean)
    });
  }, [selectedUsers, selectedGroups, users, groups]);

  const toggleUser = (userId: string) => {
    const newSelection = new Set(selectedUsers);
    if (newSelection.has(userId)) {
      newSelection.delete(userId);
    } else {
      newSelection.add(userId);
    }
    setSelectedUsers(newSelection);
  };

  const toggleGroup = (groupId: string) => {
    const newSelection = new Set(selectedGroups);
    if (newSelection.has(groupId)) {
      newSelection.delete(groupId);
    } else {
      newSelection.add(groupId);
    }
    setSelectedGroups(newSelection);
  };

  return (
    <div style={{ border: '1px solid #e0e0e0', borderRadius: 8, padding: 16 }}>
      <h3 style={{ margin: '0 0 16px 0', fontSize: 16 }}>Assign to Users & Groups</h3>
      <div style={{ marginBottom: 12 }}>
        {!account && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#fff3cd', padding: 8, borderRadius: 6, border: '1px solid #ffeeba' }}>
            <span>You're not signed in.</span>
            <button className="btn sm" onClick={() => login().then(() => loadData())}>Sign in</button>
          </div>
        )}
        {hadError && (
          <div className="small" style={{ display: 'flex', gap: 8, alignItems: 'center', background: '#f8d7da', padding: 8, borderRadius: 6, border: '1px solid #f5c6cb', marginTop: 8 }}>
            <span style={{ flex: 1 }}>{hadError}</span>
            <button className="btn ghost sm" onClick={() => loadData()}>Retry</button>
          </div>
        )}
      </div>
      
      {/* Tab Navigation */}
      <div style={{ display: 'flex', gap: 8, marginBottom: 16, borderBottom: '1px solid #e0e0e0' }}>
        {(['users', 'groups', 'structure'] as const).map(t => (
          <button 
            key={t}
            className={tab === t ? 'btn sm' : 'btn ghost sm'}
            onClick={() => setTab(t)}
          >
            {t === 'users' ? `Users (${users.length})` : t === 'groups' ? `Groups (${groups.length})` : 'Filters'}
          </button>
        ))}
      </div>

      {loading && <div className="small muted">Loading...</div>}

      {/* Filters Tab */}
      {tab === 'structure' && (
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 12 }}>
          <div>
            <label className="small">Search:</label>
            <input 
              type="text" 
              placeholder="Name, email..." 
              value={localSearch}
              onChange={e => setLocalSearch(e.target.value)}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            />
          </div>
          <div>
            <label className="small">Department:</label>
            <select 
              value={filters.department || ''} 
              onChange={e => setFilters({...filters, department: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Departments</option>
              {orgStructure.departments.map(dept => <option key={dept} value={dept}>{dept}</option>)}
            </select>
          </div>
          <div>
            <label className="small">Job Title:</label>
            <select 
              value={filters.jobTitle || ''} 
              onChange={e => setFilters({...filters, jobTitle: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Titles</option>
              {orgStructure.jobTitles.map(title => <option key={title} value={title}>{title}</option>)}
            </select>
          </div>
          <div>
            <label className="small">Location:</label>
            <select 
              value={filters.location || ''} 
              onChange={e => setFilters({...filters, location: e.target.value || undefined})}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            >
              <option value="">All Locations</option>
              {orgStructure.locations.map(loc => <option key={loc} value={loc}>{loc}</option>)}
            </select>
          </div>
        </div>
      )}

      {/* Users Tab */}
      {tab === 'users' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          {/* Users search */}
          <div style={{ marginBottom: 12 }}>
            <input 
              type="text"
              placeholder="Search users (name or email)"
              value={localSearch}
              onChange={e => setLocalSearch(e.target.value)}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            />
          </div>
          <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
            <button className="btn ghost sm" onClick={() => setSelectedUsers(new Set(users.map(u => u.id)))}>Select All</button>
            <button className="btn ghost sm" onClick={() => setSelectedUsers(new Set())}>Clear</button>
            <span className="small muted">Selected: {selectedUsers.size}</span>
          </div>
          {users.slice(0, usersPage * pageSize).map(user => (
            <div
              key={user.id}
              onClick={() => toggleUser(user.id)}
              role="button"
              style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0', cursor: 'pointer' }}
            >
              <input
                type="checkbox"
                checked={selectedUsers.has(user.id)}
                onClick={e => e.stopPropagation()}
                onChange={() => toggleUser(user.id)}
              />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 500 }}>{user.displayName}</div>
                <div className="small muted">{user.userPrincipalName}</div>
                {user.department && <div className="small muted">{user.department} • {user.jobTitle}</div>}
              </div>
            </div>
          ))}
          {(usersPage * pageSize) < users.length && (
            <div style={{ padding: 8, textAlign: 'center' }}>
              <button className="btn ghost sm" onClick={() => setUsersPage(p => p + 1)}>Load more</button>
              <div className="small muted" style={{ marginTop: 6 }}>{Math.min(usersPage * pageSize, users.length)} of {users.length}</div>
            </div>
          )}
        </div>
      )}

      {/* Groups Tab */}
      {tab === 'groups' && (
        <div style={{ maxHeight: 300, overflowY: 'auto' }}>
          {/* Groups search (client-side filter) */}
          <div style={{ marginBottom: 12 }}>
            <input 
              type="text"
              placeholder="Search groups"
              value={groupSearch}
              onChange={e => { setGroupSearch(e.target.value); setGroupsPage(1); }}
              style={{ width: '100%', padding: 6, border: '1px solid #ddd', borderRadius: 4 }}
            />
          </div>
          <div style={{ display: 'flex', gap: 8, marginBottom: 12 }}>
            <button className="btn ghost sm" onClick={() => setSelectedGroups(new Set(groups.map(g => g.id)))}>Select All</button>
            <button className="btn ghost sm" onClick={() => setSelectedGroups(new Set())}>Clear</button>
            <span className="small muted">Selected: {selectedGroups.size}</span>
          </div>
          {groups
            .filter(g => {
              if (!groupSearch.trim()) return true;
              const q = groupSearch.toLowerCase();
              return (g.displayName || '').toLowerCase().includes(q) || (g.description || '').toLowerCase().includes(q) || (g.mail || '').toLowerCase().includes(q);
            })
            .slice(0, groupsPage * pageSize)
            .map(group => (
            <div key={group.id} style={{ display: 'flex', alignItems: 'center', gap: 8, padding: 8, borderBottom: '1px solid #f0f0f0' }}>
              <input 
                type="checkbox" 
                checked={selectedGroups.has(group.id)} 
                onChange={() => toggleGroup(group.id)} 
              />
              <div style={{ flex: 1 }}>
                <div style={{ fontWeight: 500 }}>{group.displayName}</div>
                {group.description && <div className="small muted">{group.description}</div>}
                <div className="small muted">{group.memberCount || 0} members</div>
              </div>
            </div>
          ))}
          {(groupsPage * pageSize) < groups.length && (
            <div style={{ padding: 8, textAlign: 'center' }}>
              <button className="btn ghost sm" onClick={() => setGroupsPage(p => p + 1)}>Load more</button>
              <div className="small muted" style={{ marginTop: 6 }}>{Math.min(groupsPage * pageSize, groups.length)} of {groups.length}</div>
            </div>
          )}
        </div>
      )}
    </div>
  );
};
