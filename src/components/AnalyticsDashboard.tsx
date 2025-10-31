import React, { useEffect, useState } from 'react';
import { useRBAC } from '../context/RBACContext';
import { exportAnalyticsExcel } from '../utils/excelExport';
import { exportAnalyticsCsvFull } from '../utils/csvExport';
import { useAuth } from '../context/AuthContext';
import { getBusinesses } from '../services/dbService';

// Types for analytics data
interface KPIData {
  totalBatches: number;
  activeBatches: number;
  totalUsers: number;
  completionRate: number;
  overdueBatches: number;
  avgCompletionTime: number;
  lastUpdated: string;
}

interface ComplianceData {
  department: string;
  totalUsers: number;
  completed: number;
  pending: number;
  overdue: number;
  completionRate: number;
}

interface TrendData {
  date: string;
  completions: number;
  newBatches: number;
  activeUsers: number;
}

interface DocumentStats {
  documentName: string;
  batchName: string;
  totalAssigned: number;
  acknowledged: number;
  pending: number;
  avgTimeToComplete: number;
}

// KPI Card Component
const KPICard: React.FC<{ title: string; value: string | number; change?: string; color?: string; icon?: string }> = ({ 
  title, value, change, color = 'var(--primary)', icon = '📊' 
}) => (
  <div className="card" style={{ padding: 20, textAlign: 'center', background: 'linear-gradient(135deg, #fff 0%, #f8f9fa 100%)' }}>
    <div style={{ fontSize: 24, marginBottom: 8 }}>{icon}</div>
    <div style={{ fontSize: 28, fontWeight: 'bold', color, marginBottom: 4 }}>{value}</div>
    <div style={{ fontSize: 14, color: '#666', marginBottom: 4 }}>{title}</div>
    {change && (
      <div style={{ 
        fontSize: 12, 
        color: change.startsWith('+') ? '#28a745' : change.startsWith('-') ? '#dc3545' : '#666',
        fontWeight: 500
      }}>
        {change}
      </div>
    )}
  </div>
);

// Chart Component (simplified)
const SimpleChart: React.FC<{ data: any[]; type: 'line' | 'bar'; height?: number }> = ({ data, type, height = 200 }) => (
  <div style={{ 
    height, 
    border: '1px solid #e0e0e0', 
    borderRadius: 8, 
    padding: 16, 
    background: '#f8f9fa',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center',
    fontSize: 14,
    color: '#666'
  }}>
    📈 {type === 'line' ? 'Trend' : 'Bar'} Chart ({data.length} data points)
    <div className="small" style={{ marginLeft: 8 }}>Interactive charts with Chart.js/D3 would render here</div>
  </div>
);

// Data Table Component
const DataTable: React.FC<{ data: any[]; columns: Array<{ key: string; label: string; format?: (val: any) => string }> }> = ({ 
  data, columns 
}) => (
  <div style={{ overflowX: 'auto' }}>
    <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 14 }}>
      <thead>
        <tr style={{ backgroundColor: '#f8f9fa' }}>
          {columns.map(col => (
            <th key={col.key} style={{ padding: 12, textAlign: 'left', borderBottom: '2px solid #dee2e6', fontWeight: 600 }}>
              {col.label}
            </th>
          ))}
        </tr>
      </thead>
      <tbody>
        {data.map((row, i) => (
          <tr key={i} style={{ borderBottom: '1px solid #dee2e6' }}>
            {columns.map(col => (
              <td key={col.key} style={{ padding: 12 }}>
                {col.format ? col.format(row[col.key]) : row[col.key]}
              </td>
            ))}
          </tr>
        ))}
      </tbody>
    </table>
  </div>
);

// Helpers for safe formatting and pagination
const safeNum = (v: any) => {
  const n = Number(v);
  return isFinite(n) ? n : 0;
};

function paginate<T>(arr: T[], page: number, pageSize: number) {
  const total = Array.isArray(arr) ? arr.length : 0;
  const totalPages = Math.max(1, Math.ceil((total || 1) / pageSize));
  const p = Math.min(Math.max(1, page), totalPages);
  const start = (p - 1) * pageSize;
  const items = (Array.isArray(arr) ? arr : []).slice(start, start + pageSize);
  return { items, page: p, totalPages, total };
}

// Helper: format relative time like "2 min ago"
function formatRelative(ts?: string) {
  if (!ts) return '';
  const t = Date.parse(ts);
  if (isNaN(t)) return new Date(ts).toLocaleString();
  const diffMs = Date.now() - t;
  const sec = Math.floor(diffMs / 1000);
  if (sec < 60) return `${sec}s ago`;
  const min = Math.floor(sec / 60);
  if (min < 60) return `${min} min ago`;
  const hr = Math.floor(min / 60);
  if (hr < 24) return `${hr} hr${hr === 1 ? '' : 's'} ago`;
  const d = Math.floor(hr / 24);
  return `${d} day${d === 1 ? '' : 's'} ago`;
}

// Filter Component
const FilterPanel: React.FC<{ onFilterChange: (filters: any) => void; liveOptions?: { businesses: Array<{ id: string; name: string }>; departments: string[]; groups: string[] } }> = ({ onFilterChange, liveOptions }) => {
  const [filters, setFilters] = useState({
    dateRange: '30d',
    businessId: 'all',
    department: 'all',
    group: 'all',
    status: 'all',
    batchType: 'all'
  });

  const updateFilter = (key: string, value: string) => {
    const newFilters = { ...filters, [key]: value };
    setFilters(newFilters);
    onFilterChange(newFilters);
  };

  return (
    <div style={{ 
      display: 'flex', 
      gap: 16, 
      padding: 16, 
      backgroundColor: '#f8f9fa', 
      borderRadius: 8, 
      marginBottom: 24,
      flexWrap: 'wrap'
    }}>
      <div>
  <label htmlFor="f-dateRange" className="small" style={{ display: 'block', marginBottom: 4 }}>Date Range:</label>
  <select id="f-dateRange" value={filters.dateRange} onChange={e => updateFilter('dateRange', e.target.value)} className="form-control">
          <option value="7d">Last 7 days</option>
          <option value="30d">Last 30 days</option>
          <option value="90d">Last 90 days</option>
          <option value="1y">Last year</option>
        </select>
      </div>
      {liveOptions && (
        <>
          <div>
            <label htmlFor="f-business" className="small" style={{ display: 'block', marginBottom: 4 }}>Business:</label>
            <select id="f-business" value={filters.businessId} onChange={e => updateFilter('businessId', e.target.value)} className="form-control">
              <option value="all">All Businesses</option>
              {liveOptions.businesses.map(b => (
                <option key={b.id} value={b.id}>{b.name}</option>
              ))}
            </select>
          </div>
          <div>
            <label htmlFor="f-dept" className="small" style={{ display: 'block', marginBottom: 4 }}>Department:</label>
            <select id="f-dept" value={filters.department} onChange={e => updateFilter('department', e.target.value)} className="form-control">
              <option value="all">All Departments</option>
              {liveOptions.departments.map(d => (
                <option key={d} value={d}>{d}</option>
              ))}
            </select>
          </div>
          <div>
            <label htmlFor="f-group" className="small" style={{ display: 'block', marginBottom: 4 }}>Group:</label>
            <select id="f-group" value={filters.group} onChange={e => updateFilter('group', e.target.value)} className="form-control">
              <option value="all">All Groups</option>
              {liveOptions.groups.map(g => (
                <option key={g} value={g}>{g}</option>
              ))}
            </select>
          </div>
        </>
      )}
      <div>
  <label htmlFor="f-status" className="small" style={{ display: 'block', marginBottom: 4 }}>Status:</label>
  <select id="f-status" value={filters.status} onChange={e => updateFilter('status', e.target.value)} className="form-control">
          <option value="all">All Statuses</option>
          <option value="completed">Completed</option>
          <option value="pending">Pending</option>
          <option value="overdue">Overdue</option>
        </select>
      </div>
      <div>
  <label htmlFor="f-type" className="small" style={{ display: 'block', marginBottom: 4 }}>Batch Type:</label>
  <select id="f-type" value={filters.batchType} onChange={e => updateFilter('batchType', e.target.value)} className="form-control">
          <option value="all">All Types</option>
          <option value="policy">Policy Updates</option>
          <option value="training">Training Materials</option>
          <option value="compliance">Compliance Documents</option>
        </select>
      </div>
      <div style={{ alignSelf: 'end' }}>
        <button className="btn ghost sm">📊 Export Report</button>
      </div>
    </div>
  );
};

// Main Analytics Dashboard Component
const AnalyticsDashboard: React.FC = () => {
  const [loading, setLoading] = useState(false);
  const { isSuperAdmin, perms } = useRBAC();
  const [data, setData] = useState<{
    kpis: KPIData;
    compliance: ComplianceData[];
    trends: TrendData[];
    documents: DocumentStats[];
  } | null>(null);
  const [filters, setFilters] = useState({});
  const [reportYear, setReportYear] = useState<number>(new Date().getFullYear());
  const [liveOptions, setLiveOptions] = useState<{ businesses: Array<{ id: string; name: string }>; departments: string[]; groups: string[] }>({ businesses: [], departments: [], groups: [] });
  const [recipients, setRecipients] = useState<any[]>([]);
  const [activities, setActivities] = useState<any[]>([]);
  const sqliteEnabled = (process.env.REACT_APP_ENABLE_SQLITE === 'true') && !!process.env.REACT_APP_API_BASE;
  const apiBase = sqliteEnabled ? (process.env.REACT_APP_API_BASE as string).replace(/\/$/, '') : '';
  const getApiBases = () => {
    const envBase = (process.env.REACT_APP_API_BASE || '').replace(/\/$/, '');
    // @ts-ignore window dynamic hints from runtime or index.html
    const hinted = (typeof window !== 'undefined' && ((window as any).__API_BASE__ || (window as any).API_BASE)) ? String((window as any).__API_BASE__ || (window as any).API_BASE).replace(/\/$/, '') : '';
    const local = 'http://127.0.0.1:4000';
    return Array.from(new Set([envBase, hinted, local].filter(Boolean)));
  };
  const tryFetchJson = async (path: string) => {
    const bases = getApiBases();
    let lastErr: any = null;
    for (const b of bases) {
      try {
        const r = await fetch(`${b}${path}`);
        if (r.ok) return await r.json();
      } catch (e) { lastErr = e; }
    }
    if (lastErr) throw lastErr;
    throw new Error('All API base candidates failed: ' + bases.join(', '));
  };
  // Simple search & pagination states
  const [compSearch, setCompSearch] = useState('');
  const [compPage, setCompPage] = useState(1);
  const compPageSize = 10;
  const [docSearch, setDocSearch] = useState('');
  const [docPage, setDocPage] = useState(1);
  const docPageSize = 10;
  const { account } = useAuth();

  const loadAnalyticsData = async () => {
    setLoading(true);
    try {
      try {
          // Build filter query
          const q: string[] = [];
          const bf = (filters as any).businessId; if (bf && bf !== 'all') q.push(`businessId=${encodeURIComponent(bf)}`);
          const df = (filters as any).department; if (df && df !== 'all') q.push(`department=${encodeURIComponent(df)}`);
          const gf = (filters as any).group; if (gf && gf !== 'all') q.push(`primaryGroup=${encodeURIComponent(gf)}`);
          const qs = q.length ? `?${q.join('&')}` : '';
          const [statsRes, recRes, bizRes, compRes, docRes, trendRes, actRes] = await Promise.all([
            tryFetchJson(`/api/stats${qs}`),
            tryFetchJson(`/api/recipients${qs}`),
            getBusinesses().catch(() => []),
            tryFetchJson(`/api/compliance${qs}`).catch(() => []),
            tryFetchJson(`/api/doc-stats${qs}`).catch(() => []),
            tryFetchJson(`/api/trends${qs}`).catch(() => ({ completions: [], newBatches: [], activeUsers: [] })),
            tryFetchJson(`/api/activity/recent?limit=20`).catch(() => [])
          ]);
          setRecipients(Array.isArray(recRes) ? recRes : []);
          setActivities(Array.isArray(actRes) ? actRes : []);
          // Live options
          const deptSet = new Set<string>();
          const groupSet = new Set<string>();
          for (const r of (Array.isArray(recRes) ? recRes : [])) {
            if (r.department) deptSet.add(String(r.department));
            if (r.primaryGroup) groupSet.add(String(r.primaryGroup));
          }
          const businesses = Array.isArray(bizRes)
            ? bizRes.map((b: any) => ({ id: String(b.id || b.toba_businessid || ''), name: String(b.name || b.toba_name || '') }))
            : [];
          setLiveOptions({ businesses, departments: Array.from(deptSet).sort(), groups: Array.from(groupSet).sort() });

          // Normalize document stats to expected shape
          const docs: DocumentStats[] = Array.isArray(docRes)
            ? (docRes as any[]).map((d: any) => {
                const totalAssigned = Number(d.totalAssigned || 0);
                const acknowledged = Number((d.acknowledged ?? d.completed) || 0);
                const pending = Math.max(0, totalAssigned - acknowledged);
                const avgTimeToComplete = Number(d.avgTimeToComplete ?? d.avgDays ?? 0);
                const documentName = String(d.documentName || d.title || d.toba_title || d.documentId || 'Document');
                const batchName = String(d.batchName || d.toba_batchname || '—');
                return { documentName, batchName, totalAssigned, acknowledged, pending, avgTimeToComplete };
              })
            : [];

          const live = {
            kpis: { ...statsRes, lastUpdated: new Date().toISOString() },
            compliance: Array.isArray(compRes) ? compRes : [],
            trends: Array.isArray((trendRes as any).completions) ? (trendRes as any).completions.map((row: any, idx: number) => ({
              date: row.date,
              completions: Number(row.count || 0),
              newBatches: Number(((trendRes as any).newBatches?.[idx]?.count) || 0),
              activeUsers: Number(((trendRes as any).activeUsers?.[idx]?.count) || 0)
            })) : [],
            documents: docs
          };
          setData(live);
          try { (window as any).__analyticsData = { ...live, __recipients: Array.isArray(recRes) ? recRes : [] }; } catch {}
          setLoading(false);
          return;
      } catch (e) {
        setData({
          kpis: { totalBatches: 0, activeBatches: 0, totalUsers: 0, completionRate: 0, overdueBatches: 0, avgCompletionTime: 0, lastUpdated: new Date().toISOString() },
          compliance: [], trends: [], documents: []
        });
        try { (window as any).__analyticsData = null; } catch {}
        setLoading(false);
        return;
      }
    } catch (error) {
      console.error('Failed to load analytics data:', error);
      setLoading(false);
    }
  };

  useEffect(() => {
    loadAnalyticsData();
  }, [filters]);

  if (loading) {
    return (
      <div className="container">
        <div className="card" style={{ textAlign: 'center', padding: 40 }}>
          <div style={{ fontSize: 24, marginBottom: 16 }}>📊</div>
          <div>Loading analytics dashboard...</div>
          <div className="small muted" style={{ marginTop: 8 }}>
            {sqliteEnabled ? 'Fetching data from local SQLite API' : 'Fetching data from Microsoft Graph and configured backend'}
          </div>
        </div>
      </div>
    );
  }

  if (!data) {
    return (
      <div className="container">
        <div className="card">
          <div style={{ color: '#dc3545' }}>Failed to load analytics data</div>
        </div>
      </div>
    );
  }

  return (
    <div className="container">
      {/* Header */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', marginBottom: 24 }}>
        <div>
          <h1 style={{ margin: 0, fontSize: 28, color: 'var(--primary)' }}>📊 Analytics Dashboard</h1>
          <p className="small muted">Last updated: {new Date(data.kpis.lastUpdated).toLocaleString()}</p>
        </div>
        <div style={{ display: 'flex', gap: 8 }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <label htmlFor="report-year" className="small muted">Year:</label>
            <select id="report-year" className="form-control" value={reportYear}
              onChange={e => setReportYear(Number(e.target.value))}>
              {Array.from({ length: 6 }, (_, i) => new Date().getFullYear() - i).map(y => (
                <option key={y} value={y}>{y}</option>
              ))}
            </select>
          </div>
          <button className="btn ghost sm" onClick={() => loadAnalyticsData()}>🔄 Refresh</button>
          <button className="btn ghost sm">📋 Schedule Report</button>
          {(isSuperAdmin || (perms && (perms as any).exportAnalytics)) && (
            <button className="btn sm" onClick={async () => { try { await exportAnalyticsExcel({ year: reportYear, adminEmail: account?.username }); } catch (e) { console.warn('Excel export failed', e); } }}>📘 Export Excel</button>
          )}
          <button className="btn sm" onClick={async () => { try { await exportAnalyticsCsvFull({ year: reportYear, adminEmail: account?.username }); } catch (e) { console.warn('CSV export failed', e); } }}>📤 Export CSV</button>
        </div>
      </div>

  {/* Filters */}
  <FilterPanel onFilterChange={setFilters} liveOptions={liveOptions} />

      {/* KPIs */}
      <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(200px, 1fr))', gap: 16, marginBottom: 32 }}>
        <KPICard title="Total Batches" value={data.kpis.totalBatches} change="+5 this month" icon="📋" />
        <KPICard title="Active Batches" value={data.kpis.activeBatches} color="#17a2b8" icon="⚡" />
        <KPICard title="Assigned Recipients" value={Number(data.kpis.totalUsers).toLocaleString()} color="#28a745" icon="👥" />
        <KPICard title="Completion Rate" value={`${data.kpis.completionRate}%`} color="#ffc107" icon="✅" />
        <KPICard title="Overdue Batches" value={data.kpis.overdueBatches} color="#dc3545" icon="⚠️" />
        <KPICard title="Avg. Completion Time" value={`${data.kpis.avgCompletionTime} days`} color="#6f42c1" icon="⏱️" />
      </div>

      {/* Charts Section */}
      <div style={{ display: 'grid', gridTemplateColumns: '2fr 1fr', gap: 24, marginBottom: 32 }}>
        <div className="card" style={{ padding: 20 }}>
          <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>📈 Completion Trends (30 days)</h3>
          <SimpleChart data={data.trends} type="line" height={250} />
        </div>
        <div className="card" style={{ padding: 20 }}>
          <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>🎯 Compliance Status</h3>
          <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: 12, backgroundColor: '#d4edda', borderRadius: 6 }}>
              <span>Compliant</span>
              <strong style={{ color: '#155724' }}>{data.kpis.completionRate}%</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: 12, backgroundColor: '#fff3cd', borderRadius: 6 }}>
              <span>Pending</span>
              <strong style={{ color: '#856404' }}>{Math.max(0, 100 - data.kpis.completionRate)}%</strong>
            </div>
            <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', padding: 12, backgroundColor: '#f8d7da', borderRadius: 6 }}>
              <span>Overdue</span>
              <strong style={{ color: '#721c24' }}>{data.kpis.overdueBatches}%</strong>
            </div>
          </div>
        </div>
      </div>

      {/* Department Compliance Table (live: currently empty until we derive from recipients/acks) */}
      <div className="card" style={{ padding: 20, marginBottom: 32 }}>
        <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>🏢 Department Compliance Overview</h3>
        {/* Search + pagination controls */}
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
          <input
            type="search"
            placeholder="Search department..."
            value={compSearch}
            onChange={(e) => { setCompSearch(e.target.value); setCompPage(1); }}
            className="form-control"
            style={{ maxWidth: 260 }}
          />
          <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}>
            <button className="btn ghost sm" disabled={compPage <= 1} onClick={() => setCompPage(p => Math.max(1, p - 1))}>◀ Prev</button>
            <span className="small muted">Page {compPage}</span>
            <button className="btn ghost sm" disabled={compPage * compPageSize >= (data.compliance || []).filter(c => !compSearch || String(c.department||'').toLowerCase().includes(compSearch.toLowerCase())).length} onClick={() => setCompPage(p => p + 1)}>Next ▶</button>
          </div>
        </div>
        {data.compliance.length === 0 ? (
          <div className="small muted">No live compliance breakdown yet.</div>
        ) : (
          (() => {
            const filtered = (data.compliance || [])
              .filter(c => !compSearch || String(c.department||'').toLowerCase().includes(compSearch.toLowerCase()));
            const { items } = paginate(filtered, compPage, compPageSize);
            return (
              <DataTable
                data={items}
                columns={[
                  { key: 'department', label: 'Department' },
                  { key: 'totalUsers', label: 'Total Users', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'completed', label: 'Completed', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'pending', label: 'Pending', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'overdue', label: 'Overdue', format: (val) => safeNum(val).toLocaleString() },
                  { 
                    key: 'completionRate', 
                    label: 'Completion Rate', 
                    format: (val) => (
                      <span style={{ 
                        color: safeNum(val) >= 90 ? '#28a745' : safeNum(val) >= 75 ? '#ffc107' : '#dc3545',
                        fontWeight: 'bold'
                      }}>
                        {safeNum(val)}%
                      </span>
                    ) as any
                  }
                ]}
              />
            );
          })()
        )}
      </div>

      {/* Live recipients preview (filtered) */}
      {recipients && recipients.length > 0 && (
        <div className="card" style={{ padding: 20, marginBottom: 32 }}>
          <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>👥 Recipients (Filtered)</h3>
          <DataTable
            data={recipients.map((r: any) => ({
              displayName: r.displayName || r.toba_DisplayName || r.email || r.toba_Email,
              email: r.email || r.toba_Email,
              department: r.department || r.toba_Department || '—',
              group: r.primaryGroup || r.toba_PrimaryGroup || '—'
            }))}
            columns={[
              { key: 'displayName', label: 'Name' },
              { key: 'email', label: 'Email' },
              { key: 'department', label: 'Department' },
              { key: 'group', label: 'Group' }
            ]}
          />
        </div>
      )}

      {/* Document Performance */}
      <div className="card" style={{ padding: 20, marginBottom: 32 }}>
        <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>📄 Document Performance</h3>
        {/* Search + pagination controls */}
        <div style={{ display: 'flex', gap: 12, alignItems: 'center', marginBottom: 12 }}>
          <input
            type="search"
            placeholder="Search document or batch..."
            value={docSearch}
            onChange={(e) => { setDocSearch(e.target.value); setDocPage(1); }}
            className="form-control"
            style={{ maxWidth: 300 }}
          />
          <div style={{ marginLeft: 'auto', display: 'flex', alignItems: 'center', gap: 8 }}>
            <button className="btn ghost sm" disabled={docPage <= 1} onClick={() => setDocPage(p => Math.max(1, p - 1))}>◀ Prev</button>
            <span className="small muted">Page {docPage}</span>
            <button className="btn ghost sm" disabled={docPage * docPageSize >= (data.documents || []).filter(d => !docSearch || String(d.documentName||'').toLowerCase().includes(docSearch.toLowerCase()) || String(d.batchName||'').toLowerCase().includes(docSearch.toLowerCase())).length} onClick={() => setDocPage(p => p + 1)}>Next ▶</button>
          </div>
        </div>
        {data.documents.length === 0 ? (
          <div className="small muted">No live document stats yet.</div>
        ) : (
          (() => {
            const filtered = (data.documents || [])
              .filter(d => !docSearch || String(d.documentName||'').toLowerCase().includes(docSearch.toLowerCase()) || String(d.batchName||'').toLowerCase().includes(docSearch.toLowerCase()));
            const { items } = paginate(filtered, docPage, docPageSize);
            return (
              <DataTable 
                data={items}
                columns={[
                  { key: 'documentName', label: 'Document' },
                  { key: 'batchName', label: 'Batch' },
                  { key: 'totalAssigned', label: 'Assigned', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'acknowledged', label: 'Acknowledged', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'pending', label: 'Pending', format: (val) => safeNum(val).toLocaleString() },
                  { key: 'avgTimeToComplete', label: 'Avg. Time', format: (val) => `${safeNum(val)} days` }
                ]}
              />
            );
          })()
        )}
      </div>

      {/* Real-time Activity Feed (DB-backed in live mode) */}
      <div className="card" style={{ padding: 20 }}>
        <h3 style={{ margin: '0 0 16px 0', fontSize: 18 }}>🔄 Recent Activity</h3>
        <div style={{ display: 'flex', flexDirection: 'column', gap: 12 }}>
          {(sqliteEnabled && activities.length === 0) && (
            <div className="small muted">No recent activity yet.</div>
          )}
          {activities.map((activity: any, i: number) => {
            const type = activity.type || (activity.action === 'acknowledged' ? 'success' : activity.action === 'created batch' ? 'info' : 'info');
            const rel = formatRelative(activity.timestamp);
            const labelUser = activity.user || 'System';
            const labelDoc = activity.document || activity.batch || '';
            return (
              <div key={i} style={{ 
                display: 'flex', 
                alignItems: 'center', 
                gap: 12, 
                padding: 12, 
                backgroundColor: '#f8f9fa', 
                borderRadius: 6,
                borderLeft: `4px solid ${
                  type === 'success' ? '#28a745' : 
                  type === 'warning' ? '#ffc107' : 
                  type === 'info' ? '#17a2b8' : '#6c757d'
                }`
              }}>
                <div style={{ fontSize: 12, color: '#666', minWidth: 70 }}>{rel}</div>
                <div style={{ flex: 1 }}>
                  <strong>{labelUser}</strong> {activity.action} {labelDoc && (<em>{labelDoc}</em>)}
                </div>
                <div style={{ 
                  fontSize: 10, 
                  padding: '2px 6px', 
                  backgroundColor: type === 'success' ? '#d4edda' : 
                                  type === 'warning' ? '#fff3cd' : 
                                  type === 'info' ? '#d1ecf1' : '#e2e3e5',
                  color: type === 'success' ? '#155724' : 
                         type === 'warning' ? '#856404' : 
                         type === 'info' ? '#0c5460' : '#383d41',
                  borderRadius: 4,
                  textTransform: 'uppercase',
                  fontWeight: 'bold'
                }}>
                  {type}
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};

export default AnalyticsDashboard;

// --- helpers: CSV export ---
function toCSV(rows: any[], headers: string[]): string {
  const esc = (v: any) => '"' + String(v ?? '').replace(/"/g, '""') + '"';
  const head = headers.map(esc).join(',');
  const body = rows.map(r => headers.map(h => esc(r[h])).join(',')).join('\n');
  return head + '\n' + body;
}

function download(filename: string, content: string, mime = 'text/csv;charset=utf-8;') {
  const blob = new Blob([content], { type: mime });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = filename; a.click(); URL.revokeObjectURL(url);
}

function exportCsv(this: any) {
  try {
    // @ts-ignore - access component state via closure when bound
    const data = (window as any).__analyticsData as any;
    if (!data) { console.warn('No analytics data to export'); return; }
    // Recipients (if present)
    if (Array.isArray(data.__recipients) && data.__recipients.length > 0) {
      const rows = data.__recipients.map((r: any) => ({
        displayName: r.displayName || r.toba_DisplayName || r.email || r.toba_Email,
        email: r.email || r.toba_Email,
        department: r.department || r.toba_Department || '',
        group: r.primaryGroup || r.toba_PrimaryGroup || ''
      }));
      const csv = toCSV(rows, ['displayName','email','department','group']);
      download('recipients.csv', csv);
    }
    // Document performance
    if (Array.isArray(data.documents) && data.documents.length > 0) {
      const csv = toCSV(data.documents, ['documentName','batchName','totalAssigned','acknowledged','pending','avgTimeToComplete']);
      download('documents.csv', csv);
    }
  } catch (e) { console.warn('CSV export failed', e); }
}