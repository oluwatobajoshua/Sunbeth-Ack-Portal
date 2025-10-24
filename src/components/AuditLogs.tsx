import React, { useEffect, useMemo, useRef, useState } from 'react';
import { getApiBase } from '../utils/runtimeConfig';
import Modal from './Modal';

// Simple date helpers
const isoStartOfDay = (yyyyMmDd: string): string => {
  if (!yyyyMmDd) return '';
  const d = new Date(yyyyMmDd);
  d.setHours(0,0,0,0);
  return d.toISOString();
};
const isoEndOfDay = (yyyyMmDd: string): string => {
  if (!yyyyMmDd) return '';
  const d = new Date(yyyyMmDd);
  d.setHours(23,59,59,999);
  return d.toISOString();
};

export type AuditLogRow = {
  id: number;
  ts: string;
  event: string;
  email: string;
  ip?: string;
  ua?: string;
  result?: string;
  details?: string | null;
};

const pageSizeOptions = [25, 50, 100, 200];

const AuditLogs: React.FC = () => {
  const apiBase = (getApiBase() as string) || '';
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [rows, setRows] = useState<AuditLogRow[]>([]);

  // Filters
  const [q, setQ] = useState('');
  const [event, setEvent] = useState('');
  const [email, setEmail] = useState('');
  const [result, setResult] = useState('');
  const [ip, setIp] = useState('');
  const [sinceDate, setSinceDate] = useState(''); // yyyy-mm-dd
  const [untilDate, setUntilDate] = useState(''); // yyyy-mm-dd

  // Paging
  const [pageSize, setPageSize] = useState<number>(50);
  const [page, setPage] = useState<number>(1);
  const [autoRefresh, setAutoRefresh] = useState<boolean>(false);
  const refreshTimer = useRef<number | null>(null);

  // Details modal
  const [openRow, setOpenRow] = useState<AuditLogRow | null>(null);
  const [sort, setSort] = useState<{ key: keyof AuditLogRow; dir: 'asc' | 'desc' }>({ key: 'id', dir: 'desc' });
  const [filtersOpen, setFiltersOpen] = useState(false);
  const firstFilterInputRef = useRef<HTMLInputElement | null>(null);

  // Focus first input when filters modal opens
  useEffect(() => {
    if (filtersOpen) {
      const id = window.setTimeout(() => {
        try { firstFilterInputRef.current?.focus(); } catch {}
      }, 50);
      return () => window.clearTimeout(id);
    }
  }, [filtersOpen]);

  const params = useMemo(() => {
    const p: Record<string, string> = { limit: String(pageSize), offset: String((page - 1) * pageSize) };
    if (q.trim()) p.q = q.trim();
    if (event) p.event = event;
    if (email.trim()) p.email = email.trim();
    if (result) p.result = result;
    if (ip.trim()) p.ip = ip.trim();
    if (sinceDate) p.since = isoStartOfDay(sinceDate);
    if (untilDate) p.until = isoEndOfDay(untilDate);
    return p;
  }, [q, event, email, result, ip, sinceDate, untilDate, page, pageSize]);

  const load = async () => {
    setLoading(true); setError(null);
    try {
      const sp = new URLSearchParams(params as any).toString();
      const res = await fetch(`${apiBase}/api/audit-logs?${sp}`, { cache: 'no-store' });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const j = await res.json();
      const list: AuditLogRow[] = Array.isArray(j?.logs) ? j.logs : [];
      // Client sort (server already returns newest first by id)
      const sorted = [...list].sort((a, b) => {
        const { key, dir } = sort;
        let av: any = (a as any)[key];
        let bv: any = (b as any)[key];
        if (key === 'ts') { av = new Date(a.ts).getTime(); bv = new Date(b.ts).getTime(); }
        if (av < bv) return dir === 'asc' ? -1 : 1;
        if (av > bv) return dir === 'asc' ? 1 : -1;
        return 0;
      });
      setRows(sorted);
    } catch (e: any) {
      setError('Failed to load audit logs');
    } finally {
      setLoading(false);
    }
  };

  // Debounced reload on param changes
  useEffect(() => {
    const h = setTimeout(() => { setPage(1); void load(); }, 250);
    return () => clearTimeout(h);
  }, [params]);

  // Initialize last 7 days
  useEffect(() => {
    if (!sinceDate && !untilDate) {
      const now = new Date();
      const start = new Date(now);
      start.setDate(now.getDate() - 7);
      const pad = (n: number) => String(n).padStart(2, '0');
      const fmt = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
      setSinceDate(fmt(start));
      setUntilDate(fmt(now));
    }
  }, []);

  // Auto refresh handler
  useEffect(() => {
    if (autoRefresh) {
      // refresh every 30s
      const id = window.setInterval(() => { void load(); }, 30000);
      refreshTimer.current = id;
      return () => { window.clearInterval(id); refreshTimer.current = null; };
    } else {
      if (refreshTimer.current) { window.clearInterval(refreshTimer.current); refreshTimer.current = null; }
    }
    return () => {};
  }, [autoRefresh]);

  const timeAgo = (ts: string) => {
    const d = new Date(ts).getTime();
    const diff = Date.now() - d;
    const s = Math.floor(diff/1000);
    if (s < 60) return `${s}s ago`;
    const m = Math.floor(s/60);
    if (m < 60) return `${m}m ago`;
    const h = Math.floor(m/60);
    if (h < 24) return `${h}h ago`;
    const days = Math.floor(h/24);
    return `${days}d ago`;
  };

  const badge = (text: string, kind: 'ok'|'warn'|'error'|'info' = 'info') => {
    const colors: Record<string, { bg: string; fg: string; bd: string }> = {
      ok: { bg: '#e8f5e9', fg: '#1b5e20', bd: '#c8e6c9' },
      warn: { bg: '#fff8e1', fg: '#8d6e63', bd: '#ffe0b2' },
      error: { bg: '#ffebee', fg: '#b71c1c', bd: '#ffcdd2' },
      info: { bg: '#e3f2fd', fg: '#0d47a1', bd: '#bbdefb' },
    };
    const c = colors[kind] || colors.info;
    return <span className="small" style={{ background: c.bg, color: c.fg, border: `1px solid ${c.bd}`, padding: '2px 6px', borderRadius: 999 }}>{text}</span>;
  };

  const resultBadge = (r?: string) => {
    const v = String(r || '').toLowerCase();
    if (v === 'ok') return badge('ok', 'ok');
    if (v.includes('rate') || v.includes('limit')) return badge(v, 'warn');
    if (v.includes('locked') || v.includes('invalid') || v.includes('error')) return badge(v, 'error');
    return badge(v || '—', 'info');
  };

  const setQuickRange = (days: number | '24h' | 'all') => {
    if (days === 'all') { setSinceDate(''); setUntilDate(''); return; }
    const now = new Date();
    const start = new Date(now);
    if (days === '24h') { start.setDate(now.getDate()); start.setHours(now.getHours()-24); }
    else { start.setDate(now.getDate() - days); start.setHours(0,0,0,0); }
    const pad = (n: number) => String(n).padStart(2, '0');
    const fmt = (d: Date) => `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
    setSinceDate(fmt(start));
    setUntilDate(fmt(now));
  };

  const copy = async (text: string) => {
    try { await navigator.clipboard.writeText(text); } catch {}
  };

  const Chip: React.FC<{ label: string; onClick: () => void; active?: boolean; title?: string }>
    = ({ label, onClick, active, title }) => (
    <button
      className={active ? 'btn sm' : 'btn ghost sm'}
      onClick={onClick}
      title={title}
      style={{ borderRadius: 999 }}
    >{label}</button>
  );

  const HeaderStat: React.FC<{ label: string; value: string | number }>
    = ({ label, value }) => (
    <div className="card" style={{ padding: 12, textAlign: 'center' }}>
      <div style={{ fontSize: 18, fontWeight: 700, color: 'var(--primary)' }}>{value}</div>
      <div className="small muted">{label}</div>
    </div>
  );

  const parseDetails = (s?: string | null): string => {
    if (!s) return '';
    try { const o = JSON.parse(s); return typeof o === 'object' ? JSON.stringify(o) : String(s); } catch { return String(s); }
  };

  const exportCsv = () => {
    const header = ['id','ts','event','email','ip','result','details'];
    const esc = (val: any) => {
      const str = String(val ?? '');
      const doubled = str.replace(/"/g, '""');
      return `"${doubled}"`;
    };
    const body = rows.map(r => [r.id, r.ts, r.event, r.email, r.ip || '', r.result || '', parseDetails(r.details)]
      .map(esc).join(',')
    ).join('\n');
    const csv = header.map(h => `"${h}"`).join(',') + '\n' + body;
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'audit-logs.csv'; a.click();
    URL.revokeObjectURL(url);
  };

  const activeFilterCount = (
    (event ? 1 : 0) + (email ? 1 : 0) + (result ? 1 : 0) + (ip ? 1 : 0) + (sinceDate ? 1 : 0) + (untilDate ? 1 : 0)
  );

  return (
    <div className="card" style={{ padding: 16 }}>
      {/* Title */}
      <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 12, flexWrap: 'wrap', marginBottom: 12 }}>
        <div>
          <div style={{ fontWeight: 700, color: 'var(--primary)' }}>Audit Logs</div>
          <div className="small muted">Security events for external auth (login/reset/MFA/onboard)</div>
        </div>
        <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
          <input
            placeholder="Search (event/email/result/ip/details)"
            value={q}
            onChange={e => setQ(e.target.value)}
            style={{ minWidth: 240, width: 360, maxWidth: '40vw', padding: 6, border: '1px solid #ddd', borderRadius: 6 }}
          />
          <button className="btn ghost sm" onClick={() => setFiltersOpen(true)} title="Open filters">
            Filters{activeFilterCount ? ` (${activeFilterCount})` : ''}
          </button>
          <select value={pageSize} onChange={e => { setPageSize(Number(e.target.value)); setPage(1); }}>
            {pageSizeOptions.map(n => <option key={n} value={n}>{n}/page</option>)}
          </select>
          <button className="btn ghost sm" onClick={() => void load()} disabled={loading} title="Reload data">Refresh</button>
          <label className="small" style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
            <input type="checkbox" checked={autoRefresh} onChange={e => setAutoRefresh(e.target.checked)} /> Auto refresh
          </label>
          <button className="btn sm" onClick={exportCsv} disabled={rows.length === 0}>Export CSV</button>
        </div>
      </div>

      {/* Error / Loading */}
      {error && (
        <div className="small" style={{ background: '#ffebee', border: '1px solid #ffcdd2', color: '#b71c1c', padding: 10, borderRadius: 6, marginBottom: 8 }}>
          <div style={{ display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
            <span>{error}</span>
            <button className="btn ghost sm" onClick={() => void load()}>Retry</button>
          </div>
        </div>
      )}
      {loading && (
        <div className="small" style={{ marginBottom: 8 }}>
          <div className="progressBar" aria-hidden="true"><i style={{ width: '60%' }} /></div>
        </div>
      )}

      {/* Table */}
      <div style={{ overflowX: 'auto' }}>
        <table className="small" style={{ width: '100%', borderCollapse: 'collapse' }}>
          <thead>
            <tr style={{ textAlign: 'left', position: 'sticky', top: 0, background: '#fff', zIndex: 1 }}>
              <th style={{ padding: 8, borderBottom: '1px solid #eee', cursor: 'pointer' }} onClick={() => setSort(s => ({ key: 'ts', dir: s.dir === 'asc' ? 'desc' : 'asc' }))}>Time</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}>Event</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}>Email</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}>IP</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}>Result</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}>Details</th>
              <th style={{ padding: 8, borderBottom: '1px solid #eee' }}></th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 && !loading && (
              <tr><td colSpan={7} className="small muted" style={{ padding: 12 }}>No audit entries.</td></tr>
            )}
            {rows.map(r => (
              <tr key={r.id} style={{ background: (r.result || '').toLowerCase() === 'ok' ? 'transparent' : 'rgba(255, 235, 238, 0.3)' }}>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5', whiteSpace: 'nowrap' }}>
                  <div style={{ display: 'flex', gap: 6, alignItems: 'center' }}>
                    <span title={new Date(r.ts).toLocaleString()}>{timeAgo(r.ts)}</span>
                  </div>
                </td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5' }}>{r.event}</td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span>{r.email}</span>
                    {r.email && <button className="btn ghost sm" title="Copy email" onClick={() => copy(r.email)}>Copy</button>}
                  </div>
                </td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5' }}>
                  <div style={{ display: 'flex', alignItems: 'center', gap: 6 }}>
                    <span>{r.ip || ''}</span>
                    {r.ip && <button className="btn ghost sm" title="Copy IP" onClick={() => r.ip && copy(r.ip)}>Copy</button>}
                  </div>
                </td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5' }}>{resultBadge(r.result)}</td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5', maxWidth: 420, overflow: 'hidden', textOverflow: 'ellipsis' }}>{parseDetails(r.details)}</td>
                <td style={{ padding: 8, borderBottom: '1px solid #f5f5f5' }}>
                  <button className="btn ghost sm" onClick={() => setOpenRow(r)}>View</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>

      {/* Pager (simple next/prev) */}
      <div style={{ display: 'flex', justifyContent: 'space-between', marginTop: 12 }}>
        <button className="btn ghost sm" onClick={() => setPage(p => Math.max(1, p - 1))} disabled={page <= 1}>Prev</button>
        <div className="small muted">Page {page}</div>
        <button className="btn ghost sm" onClick={() => setPage(p => p + 1)} disabled={rows.length < pageSize}>Next</button>
      </div>

      {/* Filters Modal */}
      <Modal open={filtersOpen} onClose={() => setFiltersOpen(false)} title="Filters" maxWidth="90vw">
        <div className="small" style={{ display: 'grid', gap: 12 }}>
          <div style={{ display: 'grid', gap: 8, gridTemplateColumns: 'repeat(auto-fit, minmax(220px, 1fr))' }}>
            <input ref={firstFilterInputRef} placeholder="Event (e.g., login)" value={event} onChange={e => setEvent(e.target.value)} />
            <input placeholder="Email" value={email} onChange={e => setEmail(e.target.value)} />
            <input placeholder="Result (ok/error/locked)" value={result} onChange={e => setResult(e.target.value)} />
            <input placeholder="IP" value={ip} onChange={e => setIp(e.target.value)} />
            <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 8 }}>
              <div>
                <label className="small muted">Since</label>
                <input type="date" value={sinceDate} onChange={e => setSinceDate(e.target.value)} />
              </div>
              <div>
                <label className="small muted">Until</label>
                <input type="date" value={untilDate} onChange={e => setUntilDate(e.target.value)} />
              </div>
            </div>
          </div>

          <div style={{ display: 'flex', gap: 8, alignItems: 'center', flexWrap: 'wrap' }}>
            <span className="muted">Quick ranges:</span>
            <Chip label="24h" active={!!sinceDate && !!untilDate && (new Date().getTime() - new Date(sinceDate).getTime()) <= 24*60*60*1000} onClick={() => setQuickRange('24h')} />
            <Chip label="7d" active={!!sinceDate && !!untilDate && (new Date().getTime() - new Date(sinceDate).getTime()) <= 7*24*60*60*1000} onClick={() => setQuickRange(7)} />
            <Chip label="30d" active={!!sinceDate && !!untilDate && (new Date().getTime() - new Date(sinceDate).getTime()) <= 30*24*60*60*1000} onClick={() => setQuickRange(30)} />
            <Chip label="All" active={!sinceDate && !untilDate} onClick={() => setQuickRange('all')} />
          </div>

          <div>
            <div className="small muted" style={{ marginBottom: 6 }}>Events</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {['login','password_reset_request','password_reset','mfa_setup','mfa_verify','onboard_set_password'].map(ev => (
                <Chip key={ev} label={ev} active={event === ev} onClick={() => setEvent(ev)} />
              ))}
              {event && <button className="btn ghost sm" onClick={() => setEvent('')}>Clear event</button>}
            </div>
          </div>

          <div>
            <div className="small muted" style={{ marginBottom: 6 }}>Result</div>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              {['ok','error','invalid_password','invalid_code','locked','rate_limited'].map(r => (
                <Chip key={r} label={r} active={result === r} onClick={() => setResult(r)} />
              ))}
              {result && <button className="btn ghost sm" onClick={() => setResult('')}>Clear result</button>}
            </div>
          </div>

          <div style={{ position: 'sticky', bottom: 0, background: '#fff', paddingTop: 8, display: 'flex', justifyContent: 'space-between', alignItems: 'center', gap: 8 }}>
            <button className="btn ghost sm" onClick={() => { setQ(''); setEvent(''); setEmail(''); setResult(''); setIp(''); setSinceDate(''); setUntilDate(''); }}>Reset all</button>
            <div style={{ display: 'flex', gap: 8, flexWrap: 'wrap' }}>
              <button className="btn ghost sm" onClick={() => setFiltersOpen(false)}>Close</button>
              <button className="btn sm" onClick={() => setFiltersOpen(false)}>Apply & Close</button>
            </div>
          </div>
        </div>
      </Modal>

      {/* Details Modal */}
      <Modal open={!!openRow} onClose={() => setOpenRow(null)} title="Audit Entry">
        {openRow && (
          <div className="small" style={{ display: 'grid', rowGap: 8 }}>
            <div><strong>Time:</strong> {new Date(openRow.ts).toLocaleString()} ({timeAgo(openRow.ts)})</div>
            <div><strong>Event:</strong> {openRow.event}</div>
            <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
              <div><strong>Email:</strong> {openRow.email}</div>
              <button className="btn ghost sm" onClick={() => copy(openRow.email)}>Copy</button>
            </div>
            {openRow.ip && (
              <div style={{ display: 'flex', gap: 8, alignItems: 'center' }}>
                <div><strong>IP:</strong> {openRow.ip}</div>
                <button className="btn ghost sm" onClick={() => copy(openRow.ip!)}>Copy</button>
              </div>
            )}
            {openRow.result && (<div><strong>Result:</strong> {resultBadge(openRow.result)}</div>)}
            {openRow.ua && (<div><strong>User Agent:</strong> <span style={{ wordBreak: 'break-all' }}>{openRow.ua}</span></div>)}
            {openRow.details && (
              <div>
                <div style={{ marginBottom: 4 }}><strong>Details</strong></div>
                <pre style={{ background: '#f8f9fa', padding: 8, borderRadius: 6, maxHeight: 320, overflow: 'auto' }}>{
                  (() => { try { return JSON.stringify(JSON.parse(openRow.details!), null, 2); } catch { return openRow.details; } })()
                }</pre>
              </div>
            )}
          </div>
        )}
      </Modal>
    </div>
  );
};

export default AuditLogs;