/*
  Minimal SQLite API using sql.js (pure WASM, no native build). 
  - DB file: ./data/sunbeth.db (created if missing)
  - Endpoints cover app features: batches, documents, recipients, acks, progress, businesses.
*/
const path = require('path');
const fs = require('fs');
const express = require('express');
const cors = require('cors');
const initSqlJs = require('sql.js');
const http = require('http');
const https = require('https');

// Permission catalog for RBAC matrix (extendable)
const PERMISSIONS = [
  { key: 'viewAdmin', label: 'View Admin Panel', description: 'Access the Admin route and dashboards', category: 'General' },
  { key: 'manageSettings', label: 'Manage Settings', description: 'Change system settings in Admin', category: 'General' },
  { key: 'viewDebugLogs', label: 'View Debug Logs', description: 'Access troubleshooting console and logs', category: 'General' },
  { key: 'exportAnalytics', label: 'Export Analytics', description: 'Export analytics to Excel/CSV', category: 'Analytics' },
  { key: 'viewAnalytics', label: 'View Analytics', description: 'Access analytics dashboards', category: 'Analytics' },
  { key: 'createBatch', label: 'Create Batch', description: 'Create acknowledgement batches', category: 'Batches' },
  { key: 'editBatch', label: 'Edit Batch', description: 'Update batch metadata and content', category: 'Batches' },
  { key: 'deleteBatch', label: 'Delete Batch', description: 'Remove batches and related records', category: 'Batches' },
  { key: 'manageRecipients', label: 'Manage Recipients', description: 'Add/remove batch recipients', category: 'Batches' },
  { key: 'manageDocuments', label: 'Manage Documents', description: 'Add/remove documents in a batch', category: 'Batches' },
  { key: 'sendNotifications', label: 'Send Notifications', description: 'Send email notifications via Graph', category: 'Communications' },
  { key: 'uploadDocuments', label: 'Upload Documents', description: 'Upload to SharePoint libraries', category: 'Content' },
  { key: 'manageBusinesses', label: 'Manage Businesses', description: 'Create/edit/delete businesses', category: 'Data' },
  { key: 'manageRoles', label: 'Manage Roles', description: 'Add/remove Admins and Managers', category: 'Security' },
  { key: 'managePermissions', label: 'Manage Permissions', description: 'Edit RBAC matrix (role/user overrides)', category: 'Security' }
];

// Enhanced logging utility for batch operations
const createLogger = (requestId) => {
  const log = (level, operation, message, data = null) => {
    const timestamp = new Date().toISOString();
    const logEntry = {
      timestamp,
      level,
      operation,
      message,
      requestId,
      data
    };
    
    const logMessage = `[${timestamp}] [${level.toUpperCase()}] [${requestId}] ${operation}: ${message}`;
    
    switch (level) {
      case 'error':
        console.error(logMessage, data || '');
        break;
      case 'warn':
        console.warn(logMessage, data || '');
        break;
      case 'debug':
        console.debug(logMessage, data || '');
        break;
      default:
        console.log(logMessage, data || '');
    }
    
    return logEntry;
  };
  
  return {
    info: (operation, message, data) => log('info', operation, message, data),
    warn: (operation, message, data) => log('warn', operation, message, data),
    error: (operation, message, data) => log('error', operation, message, data),
    debug: (operation, message, data) => log('debug', operation, message, data)
  };
};

// Generate unique request ID
const generateRequestId = () => {
  return Math.random().toString(36).substr(2, 9) + Date.now().toString(36);
};

const DATA_DIR = path.join(__dirname, 'data');
const DB_PATH = path.join(DATA_DIR, 'sunbeth.db');

const PORT = process.env.PORT || 4000;

async function start() {
  if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });
  const SQL = await initSqlJs();

  let db;
  if (fs.existsSync(DB_PATH)) {
    const filebuffer = fs.readFileSync(DB_PATH);
    db = new SQL.Database(filebuffer);
  } else {
    db = new SQL.Database();
    try { db.run('PRAGMA foreign_keys = ON'); } catch {}
    bootstrapSchema(db);
    persist(db);
  }
  try { db.run('PRAGMA foreign_keys = ON'); } catch {}

  // Attempt lightweight migrations for existing databases
  try { migrateSchema(db); } catch (e) { console.warn('Schema migration warning (non-fatal):', e); }


  // Utilities (move up so 'one' and 'all' are defined before use)
  const exec = (sql, params = []) => {
    try { db.run(sql, params); persist(db); return true; } catch (e) { console.error(e); return false; }
  };
  const all = (sql, params = []) => {
    const stmt = db.prepare(sql);
    const rows = [];
    try {
      stmt.bind(params);
      while (stmt.step()) {
        const row = stmt.getAsObject();
        rows.push(row);
      }
    } finally { stmt.free(); }
    return rows;
  };
  const one = (sql, params = []) => all(sql, params)[0] || null;

  // Seed DB roles from environment (.env) for Admins/Managers (idempotent)
  try {
    const parseList = (s) => String(s || '')
      .split(',')
      .map(x => String(x).trim().toLowerCase())
      .filter(x => x && x.includes('@'));
    const admins = parseList(process.env.REACT_APP_ADMINS);
    const managers = parseList(process.env.REACT_APP_MANAGERS);
    if ((admins.length + managers.length) > 0) {
      db.run('BEGIN');
      const now = new Date().toISOString();
      try {
        for (const e of admins) {
          db.run('INSERT OR IGNORE INTO roles (email, role, createdAt) VALUES (?, ?, ?)', [e, 'Admin', now]);
        }
        for (const e of managers) {
          db.run('INSERT OR IGNORE INTO roles (email, role, createdAt) VALUES (?, ?, ?)', [e, 'Manager', now]);
        }
        db.run('COMMIT');
        persist(db);
      } catch (e) {
        try { db.run('ROLLBACK'); } catch {}
        console.warn('Env roles seed failed (non-fatal):', e?.message || e);
      }
    }
  } catch (e) {
    console.warn('Env roles parse failed (non-fatal):', e?.message || e);
  }

  // Seed default role permissions if none exist
  try {
    const has = one('SELECT COUNT(*) as c FROM role_permissions')?.c || 0;
    if (has === 0) {
      const allowAll = (keys) => Object.fromEntries(keys.map(k => [k, 1]));
      const denyAll = (keys) => Object.fromEntries(keys.map(k => [k, 0]));
      const keys = PERMISSIONS.map(p => p.key);
      const adminDefaults = allowAll(keys);
      // Manager defaults: allow most, restrict destructive and security
      const managerDefaults = allowAll(keys);
      for (const k of ['deleteBatch','manageRoles','managePermissions','manageSettings','viewDebugLogs','manageBusinesses']) managerDefaults[k] = 0;
      const employeeDefaults = denyAll(keys);
      const seedRole = (role, mapping) => {
        for (const k of keys) {
          const v = mapping[k] ? 1 : 0;
          db.run('INSERT OR IGNORE INTO role_permissions (role, permKey, value) VALUES (?, ?, ?)', [role, k, v]);
        }
      };
      db.run('BEGIN');
      try {
        seedRole('Admin', adminDefaults);
        seedRole('Manager', managerDefaults);
        seedRole('Employee', employeeDefaults);
        db.run('COMMIT');
        persist(db);
      } catch (e) { try { db.run('ROLLBACK'); } catch {} }
    }
  } catch (e) { console.warn('Default role-permissions seed failed (non-fatal):', e?.message || e); }

  const app = express();
  app.use(cors());
  app.use(express.json({ limit: '2mb' }));
  
  // Request logging middleware
  app.use((req, res, next) => {
    req.requestId = generateRequestId();
    req.logger = createLogger(req.requestId);
    
    const startTime = Date.now();
    req.logger.info('request', `${req.method} ${req.url}`, {
      ip: req.ip,
      userAgent: req.get('User-Agent'),
      contentType: req.get('Content-Type')
    });
    
    // Override res.json to log responses
    const originalJson = res.json.bind(res);
    res.json = function(data) {
      const duration = Date.now() - startTime;
      req.logger.info('response', `${req.method} ${req.url} - ${res.statusCode}`, {
        duration: `${duration}ms`,
        status: res.statusCode,
        dataSize: JSON.stringify(data).length
      });
      return originalJson(data);
    };
    
    // Override res.status for error logging
    const originalStatus = res.status.bind(res);
    res.status = function(code) {
      if (code >= 400) {
        const duration = Date.now() - startTime;
        req.logger.error('response', `${req.method} ${req.url} - ${code}`, {
          duration: `${duration}ms`,
          status: code
        });
      }
      return originalStatus(code);
    };
    
    next();
  });

  // Notification Emails API
  app.get('/api/notification-emails', (req, res) => {
    try {
      const stmt = db.prepare('SELECT email FROM notification_emails ORDER BY email ASC');
      const emails = [];
      while (stmt.step()) {
        const row = stmt.getAsObject();
        emails.push(row.email);
      }
      stmt.free();
      res.json({ emails });
    } catch (e) {
      res.status(500).json({ error: 'Failed to load notification emails', details: e?.message || e });
    }
  });

  app.post('/api/notification-emails', (req, res) => {
    const emails = Array.isArray(req.body.emails) ? req.body.emails : [];
    try {
      db.run('BEGIN');
      db.run('DELETE FROM notification_emails');
      for (const email of emails) {
        if (typeof email === 'string' && email.includes('@')) {
          db.run('INSERT OR IGNORE INTO notification_emails (email) VALUES (?)', [email.trim().toLowerCase()]);
        }
      }
      db.run('COMMIT');
      persist(db);
      res.json({ success: true });
    } catch (e) {
      try { db.run('ROLLBACK'); } catch {}
      res.status(500).json({ error: 'Failed to save notification emails', details: e?.message || e });
    }
  });


  // Ensure at least one business exists (after helpers available)
  try {
    const cnt = one('SELECT COUNT(*) as c FROM businesses')?.c || 0;
    if (cnt === 0) {
      db.run("INSERT INTO businesses (name, code, isActive, description) VALUES ('Default Business', 'DEF', 1, 'Auto-created')");
      persist(db);
    }
  } catch (e) { console.warn('Business seed check failed (non-fatal):', e); }

  // Routes
  // RBAC: permissions catalog
  app.get('/api/rbac/permissions', (_req, res) => {
    res.json(PERMISSIONS);
  });
  // RBAC: role permissions (get)
  app.get('/api/rbac/role-permissions', (req, res) => {
    const role = (req.query.role || '').toString();
    const rows = role
      ? all('SELECT role, permKey, value FROM role_permissions WHERE role=?', [role])
      : all('SELECT role, permKey, value FROM role_permissions');
    res.json(rows.map(r => ({ role: r.role, permKey: r.permKey, value: !!r.value })));
  });
  // RBAC: role permissions (set mapping for a role)
  app.put('/api/rbac/role-permissions', (req, res) => {
    try {
      const { role, mapping } = req.body || {};
      if (!role || typeof mapping !== 'object') return res.status(400).json({ error: 'invalid_payload' });
      db.run('BEGIN');
      try {
        // Upsert each perm
        for (const p of PERMISSIONS) {
          if (!(p.key in mapping)) continue;
          const val = mapping[p.key] ? 1 : 0;
          db.run('INSERT INTO role_permissions (role, permKey, value) VALUES (?, ?, ?) ON CONFLICT(LOWER(role), permKey) DO UPDATE SET value=excluded.value', [role, p.key, val]);
        }
        db.run('COMMIT'); persist(db);
        res.json({ ok: true });
      } catch (e) { try { db.run('ROLLBACK'); } catch {}; throw e; }
    } catch (e) { console.error('role-permissions update failed', e); res.status(500).json({ error: 'update_failed' }); }
  });
  // RBAC: user permissions (get)
  app.get('/api/rbac/user-permissions', (req, res) => {
    const email = (req.query.email || '').toString().trim().toLowerCase();
    const rows = email
      ? all('SELECT email, permKey, value FROM user_permissions WHERE LOWER(email)=LOWER(?)', [email])
      : all('SELECT email, permKey, value FROM user_permissions');
    res.json(rows.map(r => ({ email: (r.email || '').toLowerCase(), permKey: r.permKey, value: !!r.value })));
  });
  // RBAC: user permissions (set mapping for a user)
  app.put('/api/rbac/user-permissions', (req, res) => {
    try {
      const { email, mapping } = req.body || {};
      const e = String(email || '').trim().toLowerCase();
      if (!e || !e.includes('@') || typeof mapping !== 'object') return res.status(400).json({ error: 'invalid_payload' });
      db.run('BEGIN');
      try {
        for (const p of PERMISSIONS) {
          if (!(p.key in mapping)) continue;
          const val = mapping[p.key] ? 1 : 0;
          db.run('INSERT INTO user_permissions (email, permKey, value) VALUES (?, ?, ?) ON CONFLICT(LOWER(email), permKey) DO UPDATE SET value=excluded.value', [e, p.key, val]);
        }
        db.run('COMMIT'); persist(db);
        res.json({ ok: true });
      } catch (e) { try { db.run('ROLLBACK'); } catch {}; throw e; }
    } catch (e) { console.error('user-permissions update failed', e); res.status(500).json({ error: 'update_failed' }); }
  });
  // RBAC: effective permissions for a user (email required)
  app.get('/api/rbac/effective', (req, res) => {
    try {
      const email = (req.query.email || '').toString().trim().toLowerCase();
      if (!email || !email.includes('@')) return res.status(400).json({ error: 'email_required' });
      const roles = resolveUserRoles(email, db);
      // SuperAdmin shortcut: everything true
      const effective = {};
      for (const p of PERMISSIONS) effective[p.key] = false;
      if (roles.includes('SuperAdmin')) {
        for (const p of PERMISSIONS) effective[p.key] = true;
        return res.json({ roles, permissions: effective });
      }
      // Apply role defaults/mapping
      const roleRows = all('SELECT role, permKey, value FROM role_permissions WHERE LOWER(role) IN (' + roles.map(() => 'LOWER(?)').join(',') + ')', roles);
      for (const r of roleRows) {
        effective[r.permKey] = effective[r.permKey] || !!r.value; // OR semantics across roles
      }
      // Apply user overrides (can set true/false explicitly)
      const userRows = all('SELECT permKey, value FROM user_permissions WHERE LOWER(email)=LOWER(?)', [email]);
      for (const u of userRows) {
        effective[u.permKey] = !!u.value;
      }
      res.json({ roles, permissions: effective });
    } catch (e) { console.error('effective perms failed', e); res.status(500).json({ error: 'failed' }); }
  });
  // Health
  app.get('/api/health', (_req, res) => res.json({ ok: true }));
  // Root helper
  app.get('/', (_req, res) => {
    res.type('text/plain').send(
      'Sunbeth SQLite API is running.\n' +
      'Try GET /api/health or call the app at http://localhost:3000.\n' +
      'Available endpoints: /api/batches, /api/batches/:id/documents, /api/batches/:id/acks, /api/batches/:id/progress, /api/ack, /api/seed' 
    );
  });

  // Stats for dashboards/overview (supports filters via query: businessId, department, primaryGroup)
  app.get('/api/stats', (req, res) => {
    const filters = [];
    const params = [];
    const hasFilters = () => filters.length > 0;
    if (req.query.businessId) { filters.push('r.businessId = ?'); params.push(Number(req.query.businessId)); }
    if (req.query.department) { filters.push('LOWER(r.department) = ?'); params.push(String(req.query.department).toLowerCase()); }
    if (req.query.primaryGroup) { filters.push('LOWER(r.primaryGroup) = ?'); params.push(String(req.query.primaryGroup).toLowerCase()); }
    const where = hasFilters() ? `WHERE ${filters.join(' AND ')}` : '';

    let totalRecipients = 0;
    if (hasFilters()) {
      totalRecipients = one(`SELECT COUNT(*) as c FROM recipients r ${where}`, params)?.c || 0;
    } else {
      totalRecipients = one('SELECT COUNT(*) as c FROM recipients')?.c || 0;
    }

    let ackTrue = 0;
    if (hasFilters()) {
      ackTrue = one(
        `SELECT COUNT(*) as c FROM acks a 
         JOIN recipients r ON r.batchId=a.batchId AND LOWER(r.email)=LOWER(a.email)
         ${where} AND a.acknowledged=1`, params
      )?.c || 0;
    } else {
      ackTrue = one('SELECT COUNT(*) as c FROM acks WHERE acknowledged=1')?.c || 0;
    }

    const completionRate = totalRecipients > 0 ? Math.round((ackTrue / totalRecipients) * 1000) / 10 : 0;

    let totalBatches = 0;
    let activeBatches = 0;
    if (hasFilters()) {
      totalBatches = one(
        `SELECT COUNT(DISTINCT r.batchId) as c FROM recipients r ${where}`,
        params
      )?.c || 0;
      activeBatches = one(
        `SELECT COUNT(DISTINCT r.batchId) as c FROM recipients r 
         JOIN batches b ON b.id=r.batchId 
         ${where} AND b.status=1`, params
      )?.c || 0;
    } else {
      totalBatches = one('SELECT COUNT(*) as c FROM batches')?.c || 0;
      activeBatches = one('SELECT COUNT(*) as c FROM batches WHERE status=1')?.c || totalBatches;
    }

    res.json({ totalBatches, activeBatches, totalUsers: totalRecipients, completionRate, overdueBatches: 0, avgCompletionTime: 0 });
  });

  // Compliance breakdown by department (supports filters)
  app.get('/api/compliance', (req, res) => {
    const filters = [];
    const params = [];
    if (req.query.businessId) { filters.push('r.businessId = ?'); params.push(Number(req.query.businessId)); }
    if (req.query.department) { filters.push('LOWER(r.department) = ?'); params.push(String(req.query.department).toLowerCase()); }
    if (req.query.primaryGroup) { filters.push('LOWER(r.primaryGroup) = ?'); params.push(String(req.query.primaryGroup).toLowerCase()); }
    const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';
    // Totals per department
    const totals = all(`SELECT COALESCE(r.department,'Unspecified') as department, COUNT(*) as totalUsers FROM recipients r ${where} GROUP BY COALESCE(r.department,'Unspecified')`, params);
    // Acks per department
    const acks = all(
      `SELECT COALESCE(r.department,'Unspecified') as department, COUNT(*) as completed
       FROM acks a 
       JOIN recipients r ON r.batchId=a.batchId AND LOWER(r.email)=LOWER(a.email)
       ${where} AND a.acknowledged=1
       GROUP BY COALESCE(r.department,'Unspecified')`, params
    );
    const ackMap = new Map(acks.map(r => [String(r.department), Number(r.completed)]));
    const rows = totals.map(t => {
      const totalUsers = Number(t.totalUsers) || 0;
      const completed = Number(ackMap.get(String(t.department)) || 0);
      const pending = Math.max(0, totalUsers - completed);
      const overdue = 0;
      const completionRate = totalUsers > 0 ? Math.round((completed / totalUsers) * 1000) / 10 : 0;
      return { department: String(t.department), totalUsers, completed, pending, overdue, completionRate };
    });
    res.json(rows);
  });

  // Document performance stats (supports filters)
  app.get('/api/doc-stats', (req, res) => {
    const filters = [];
    const params = [];
    if (req.query.businessId) { filters.push('r.businessId = ?'); params.push(Number(req.query.businessId)); }
    if (req.query.department) { filters.push('LOWER(r.department) = ?'); params.push(String(req.query.department).toLowerCase()); }
    if (req.query.primaryGroup) { filters.push('LOWER(r.primaryGroup) = ?'); params.push(String(req.query.primaryGroup).toLowerCase()); }
    const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';

    // Total assigned = recipients per batch for each document
    const assigned = all(
      `SELECT d.id as documentId, d.title as documentName, b.name as batchName, COUNT(r.id) as totalAssigned
       FROM documents d
       JOIN batches b ON b.id=d.batchId
       JOIN recipients r ON r.batchId=d.batchId
       ${where}
       GROUP BY d.id, d.title, b.name
       ORDER BY d.id DESC`, params
    );
    // Acknowledged per document (filtered via recipients)
    const acked = all(
      `SELECT d.id as documentId, COUNT(a.id) as acknowledged
       FROM documents d
       JOIN acks a ON a.documentId=d.id AND a.acknowledged=1
       JOIN recipients r ON r.batchId=d.batchId AND LOWER(r.email)=LOWER(a.email)
       ${where}
       GROUP BY d.id`, params
    );
    const ackMap = new Map(acked.map(r => [Number(r.documentId), Number(r.acknowledged)]));
    const rows = assigned.map(a => {
      const acknowledged = Number(ackMap.get(Number(a.documentId)) || 0);
      const totalAssigned = Number(a.totalAssigned) || 0;
      const pending = Math.max(0, totalAssigned - acknowledged);
      const avgTimeToComplete = 0;
      return {
        documentName: String(a.documentName),
        batchName: String(a.batchName),
        totalAssigned,
        acknowledged,
        pending,
        avgTimeToComplete
      };
    });
    res.json(rows);
  });

  // Trends over the last 30 days (supports filters)
  app.get('/api/trends', (req, res) => {
    const filters = [];
    const params = [];
    if (req.query.businessId) { filters.push('r.businessId = ?'); params.push(Number(req.query.businessId)); }
    if (req.query.department) { filters.push('LOWER(r.department) = ?'); params.push(String(req.query.department).toLowerCase()); }
    if (req.query.primaryGroup) { filters.push('LOWER(r.primaryGroup) = ?'); params.push(String(req.query.primaryGroup).toLowerCase()); }
    const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';

    // Completions per day
    const completions = all(
      `SELECT substr(a.ackDate,1,10) as date, COUNT(*) as cnt
       FROM acks a
       ${filters.length ? 'JOIN recipients r ON r.batchId=a.batchId AND LOWER(r.email)=LOWER(a.email)' : ''}
       ${filters.length ? where + ' AND ' : 'WHERE '} a.ackDate >= date('now','-29 day')
       GROUP BY substr(a.ackDate,1,10)
       ORDER BY date`, params
    );
    // New batches per day (use startDate as proxy)
    let newBatches = [];
    if (filters.length) {
      newBatches = all(
        `SELECT b.startDate as date, COUNT(DISTINCT b.id) as cnt
         FROM batches b
         JOIN recipients r ON r.batchId=b.id
         ${where} AND b.startDate IS NOT NULL AND b.startDate >= date('now','-29 day')
         GROUP BY b.startDate
         ORDER BY b.startDate`, params
      );
    } else {
      newBatches = all(
        `SELECT startDate as date, COUNT(*) as cnt
         FROM batches 
         WHERE startDate IS NOT NULL AND startDate >= date('now','-29 day')
         GROUP BY startDate
         ORDER BY startDate`
      );
    }
    // Active users per day (distinct emails with acks)
    const activeUsers = all(
      `SELECT substr(a.ackDate,1,10) as date, COUNT(DISTINCT LOWER(a.email)) as cnt
       FROM acks a
       ${filters.length ? 'JOIN recipients r ON r.batchId=a.batchId AND LOWER(r.email)=LOWER(a.email)' : ''}
       ${filters.length ? where + ' AND ' : 'WHERE '} a.ackDate >= date('now','-29 day')
       GROUP BY substr(a.ackDate,1,10)
       ORDER BY date`, params
    );

    // Normalize to last 30 days, fill zeros for missing days
    const days = Array.from({ length: 30 }, (_, i) => new Date(Date.now() - (29 - i) * 24*60*60*1000).toISOString().slice(0,10));
    const mapRows = (rows) => {
      const m = new Map(rows.map(r => [String(r.date), Number(r.cnt)]));
      return days.map(d => ({ date: d, count: Number(m.get(d) || 0) }));
    };
    const series = {
      completions: mapRows(completions),
      newBatches: mapRows(newBatches),
      activeUsers: mapRows(activeUsers)
    };
    res.json(series);
  });

  // Recipients listing with optional filters for analytics filter panels
  app.get('/api/recipients', (req, res) => {
    const filters = [];
    const params = [];
    if (req.query.businessId) { filters.push('businessId = ?'); params.push(Number(req.query.businessId)); }
    if (req.query.department) { filters.push('LOWER(department) = ?'); params.push(String(req.query.department).toLowerCase()); }
    if (req.query.primaryGroup) { filters.push('LOWER(primaryGroup) = ?'); params.push(String(req.query.primaryGroup).toLowerCase()); }
    const where = filters.length ? `WHERE ${filters.join(' AND ')}` : '';
    const rows = all(`SELECT id, batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup FROM recipients ${where} ORDER BY id DESC`, params);
    res.json(rows);
  });

  // Recent activity feed: acknowledgements and batch creations (via startDate)
  app.get('/api/activity/recent', (req, res) => {
    try {
      const limit = Math.max(1, Math.min(Number(req.query.limit || 20), 100));

      // Latest acknowledgements
      const ackRows = all(
        `SELECT a.ackDate AS timestamp,
                LOWER(a.email) AS email,
                COALESCE(r.displayName, a.email) AS displayName,
                d.title AS documentTitle,
                b.name AS batchName
         FROM acks a
         JOIN documents d ON d.id = a.documentId
         JOIN batches b ON b.id = a.batchId
         LEFT JOIN recipients r ON r.batchId = a.batchId AND LOWER(r.email) = LOWER(a.email)
         WHERE a.ackDate IS NOT NULL
         ORDER BY a.ackDate DESC
         LIMIT ?`, [limit]
      ).map(r => ({
        timestamp: r.timestamp,
        type: 'success',
        action: 'acknowledged',
        user: r.displayName || r.email,
        email: r.email,
        document: r.documentTitle,
        batch: r.batchName
      }));

      // Recent batch creations (use startDate as proxy for creation date if available)
      const batchRows = all(
        `SELECT b.startDate AS timestamp,
                b.name AS batchName
         FROM batches b
         WHERE b.startDate IS NOT NULL
         ORDER BY b.startDate DESC
         LIMIT ?`, [Math.max(1, Math.floor(limit / 2))]
      ).map(r => ({
        timestamp: r.timestamp,
        type: 'info',
        action: 'created batch',
        user: null,
        email: null,
        document: r.batchName,
        batch: r.batchName
      }));

      // Merge and sort by timestamp desc
      const combined = [...ackRows, ...batchRows]
        .filter(ev => !!ev.timestamp)
        .sort((a, b) => String(b.timestamp).localeCompare(String(a.timestamp)))
        .slice(0, limit);

      res.json(combined);
    } catch (e) {
      console.error('recent activity failed', e);
      res.status(500).json({ error: 'activity_failed' });
    }
  });

  // Simple streaming proxy to bypass X-Frame-Options/CSP on third-party hosts when embedding
  // Usage: GET /api/proxy?url=https%3A%2F%2Fexample.com%2Ffile.pdf
  app.get('/api/proxy', (req, res) => {
    try {
      const raw = (req.query.url || '').toString();
      if (!raw) return res.status(400).json({ error: 'url_required' });
      let target;
      try { target = new URL(raw); } catch { return res.status(400).json({ error: 'invalid_url' }); }
      if (!/^https?:$/.test(target.protocol)) return res.status(400).json({ error: 'unsupported_protocol' });

      const forward = (urlObj, redirects = 0) => {
        const client = urlObj.protocol === 'https:' ? https : http;
        const reqOpts = { method: 'GET', headers: { 'User-Agent': 'Sunbeth-Proxy/1.0' } };
        const r = client.request(urlObj, reqOpts, (upstream) => {
          // Handle simple redirects up to 3 hops
          if (upstream.statusCode >= 300 && upstream.statusCode < 400 && upstream.headers.location && redirects < 3) {
            try { const next = new URL(upstream.headers.location, urlObj); forward(next, redirects + 1); } catch { res.status(502).end(); }
            return;
          }
          // Propagate content-type if available; force inline disposition
          const ct = upstream.headers['content-type'] || 'application/octet-stream';
          res.setHeader('Content-Type', ct);
          res.setHeader('Cache-Control', 'no-store');
          res.removeHeader && res.removeHeader('X-Frame-Options');
          upstream.on('error', () => { try { res.destroy(); } catch {} });
          upstream.pipe(res);
        });
        r.on('error', () => { if (!res.headersSent) res.status(502).json({ error: 'upstream_error' }); else try { res.destroy(); } catch {} });
        r.end();
      };
      forward(target);
    } catch (e) {
      console.error('Proxy error', e);
      res.status(500).json({ error: 'proxy_failed' });
    }
  });

  // Authenticated Microsoft Graph streaming proxy for SharePoint files
  // Usage (query):
  //   - GET /api/proxy/graph?driveId=...&itemId=...&token=...
  //   - GET /api/proxy/graph?url=https%3A%2F%2Fcontoso.sharepoint.com%2F...&token=...
  // Token is optional if supplied via Authorization header; otherwise required as query param.
  app.get('/api/proxy/graph', (req, res) => {
    try {
      const driveId = (req.query.driveId || '').toString();
      const itemId = (req.query.itemId || '').toString();
      const rawUrl = (req.query.url || '').toString();
      const qToken = (req.query.token || '').toString();
      const hdrAuth = (req.headers['authorization'] || '').toString();
      const bearer = qToken ? `Bearer ${qToken}` : (hdrAuth && /^Bearer\s+/i.test(hdrAuth) ? hdrAuth : '');
      const download = (req.query.download || '').toString() === '1';
      if (!bearer) return res.status(401).json({ error: 'token_required' });

      let url;
      if (driveId && itemId) {
        url = new URL(`https://graph.microsoft.com/v1.0/drives/${encodeURIComponent(driveId)}/items/${encodeURIComponent(itemId)}/content`);
      } else if (rawUrl) {
        // Build shares URL id: 'u!' + base64urlencode(originalUrl)
        const b64 = Buffer.from(rawUrl, 'utf8').toString('base64').replace(/=/g, '').replace(/\+/g, '-').replace(/\//g, '_');
        const shareId = `u!${b64}`;
        url = new URL(`https://graph.microsoft.com/v1.0/shares/${shareId}/driveItem/content`);
      } else {
        return res.status(400).json({ error: 'missing_ids_or_url' });
      }

      const follow = (targetUrl, redirects = 0) => {
        const opts = { method: 'GET', headers: { 'Authorization': bearer, 'User-Agent': 'Sunbeth-Graph-Proxy/1.0' } };
        const r = https.request(targetUrl, opts, (upstream) => {
          // Handle Graph 302 redirect to a pre-authenticated blob URL
          if (upstream.statusCode >= 300 && upstream.statusCode < 400 && upstream.headers.location && redirects < 3) {
            try {
              const next = new URL(upstream.headers.location, targetUrl);
              const client = next.protocol === 'https:' ? https : http;
              const r2 = client.request(next, { method: 'GET', headers: { 'User-Agent': 'Sunbeth-Graph-Proxy/1.0' } }, (up2) => {
                const ct2 = up2.headers['content-type'] || 'application/octet-stream';
                res.setHeader('Content-Type', ct2);
                res.setHeader('Cache-Control', 'no-store');
                if (download) {
                  try {
                    const cd = up2.headers['content-disposition'];
                    let name = null;
                    if (cd && /filename\*=utf-8''([^;]+)|filename="?([^";]+)"?/i.test(String(cd))) {
                      const m = String(cd).match(/filename\*=utf-8''([^;]+)|filename="?([^";]+)"?/i);
                      name = decodeURIComponent(m[1] || m[2] || 'file');
                    }
                    if (!name) {
                      const candidate = decodeURIComponent((next.pathname || '').split('/').pop() || '').trim();
                      name = candidate && candidate !== 'content' ? candidate : 'document';
                    }
                    // If we know it's a PDF and there's no extension, add .pdf for better UX
                    if (/application\/pdf/i.test(ct2) && !/\.pdf$/i.test(name)) name = `${name}.pdf`;
                    res.setHeader('Content-Disposition', `attachment; filename="${name}"`);
                  } catch {}
                }
                up2.on('error', () => { try { res.destroy(); } catch {} });
                up2.pipe(res);
              });
              r2.on('error', () => { if (!res.headersSent) res.status(502).json({ error: 'upstream_error' }); else try { res.destroy(); } catch {} });
              r2.end();
              return;
            } catch {
              return res.status(502).json({ error: 'redirect_failed' });
            }
          }
          // No redirect: stream as-is
          const ct = upstream.headers['content-type'] || 'application/octet-stream';
          res.setHeader('Content-Type', ct);
          res.setHeader('Cache-Control', 'no-store');
          if (download) {
            try {
              const cd = upstream.headers['content-disposition'];
              let name = null;
              if (cd && /filename\*=utf-8''([^;]+)|filename="?([^";]+)"?/i.test(String(cd))) {
                const m = String(cd).match(/filename\*=utf-8''([^;]+)|filename="?([^";]+)"?/i);
                name = decodeURIComponent(m[1] || m[2] || 'file');
              }
              if (!name) name = 'document';
              if (/application\/pdf/i.test(ct) && !/\.pdf$/i.test(name)) name = `${name}.pdf`;
              res.setHeader('Content-Disposition', `attachment; filename="${name}"`);
            } catch {}
          }
          upstream.on('error', () => { try { res.destroy(); } catch {} });
          upstream.pipe(res);
        });
        r.on('error', () => { if (!res.headersSent) res.status(502).json({ error: 'upstream_error' }); else try { res.destroy(); } catch {} });
        r.end();
      };
      follow(url);
    } catch (e) {
      console.error('Graph proxy error', e);
      res.status(500).json({ error: 'proxy_failed' });
    }
  });

  // Businesses
  app.get('/api/businesses', (_req, res) => {
    const rows = all('SELECT id, name, code, isActive, description FROM businesses ORDER BY name');
    res.json(rows.map(r => ({ id: r.id, name: r.name, code: r.code, isActive: !!r.isActive, description: r.description })));
  });
  // Create business
  app.post('/api/businesses', (req, res) => {
    const { name, code = null, isActive = true, description = null } = req.body || {};
    if (!name || String(name).trim().length === 0) return res.status(400).json({ error: 'name_required' });
    try {
      db.run('INSERT INTO businesses (name, code, isActive, description) VALUES (?, ?, ?, ?)', [String(name).trim(), code, isActive ? 1 : 0, description]);
      const id = one('SELECT last_insert_rowid() as id')?.id;
      persist(db);
      res.json({ id });
    } catch (e) {
      console.error('Create business failed', e);
      res.status(500).json({ error: 'insert_failed' });
    }
  });
  // Update business
  app.put('/api/businesses/:id', (req, res) => {
    const id = Number(req.params.id);
    const { name, code, isActive, description } = req.body || {};
    try {
      const current = one('SELECT id, name, code, isActive, description FROM businesses WHERE id=?', [id]);
      if (!current) return res.status(404).json({ error: 'not_found' });
      const next = {
        name: name != null ? String(name).trim() : current.name,
        code: code != null ? code : current.code,
        isActive: isActive != null ? (isActive ? 1 : 0) : current.isActive,
        description: description != null ? description : current.description
      };
      db.run('UPDATE businesses SET name=?, code=?, isActive=?, description=? WHERE id=?', [next.name, next.code, next.isActive, next.description, id]);
      persist(db);
      res.json({ ok: true });
    } catch (e) {
      console.error('Update business failed', e);
      res.status(500).json({ error: 'update_failed' });
    }
  });
  // Delete business (sets recipients.businessId = NULL for references)
  app.delete('/api/businesses/:id', (req, res) => {
    const id = Number(req.params.id);
    try {
      db.run('BEGIN');
      db.run('UPDATE recipients SET businessId=NULL WHERE businessId=?', [id]);
      db.run('DELETE FROM businesses WHERE id=?', [id]);
      db.run('COMMIT');
      persist(db);
      res.json({ ok: true });
    } catch (e) {
      try { db.run('ROLLBACK'); } catch {}
      console.error('Delete business failed', e);
      res.status(500).json({ error: 'delete_failed' });
    }
  });

  // Batches assigned to a user (via recipients)
  app.get('/api/batches', (req, res) => {
    const email = (req.query.email || '').toString().trim().toLowerCase();
    if (!email) {
      // return all (admin view) if no email specified
      const rows = all('SELECT id, name, startDate, dueDate, status, description FROM batches ORDER BY id DESC');
      return res.json(rows.map(mapBatch));
    }
    const rows = all(
      `SELECT DISTINCT b.id, b.name, b.startDate, b.dueDate, b.status, b.description
       FROM batches b
       JOIN recipients r ON r.batchId=b.id
       WHERE LOWER(r.email)=? ORDER BY b.id DESC`, [email]
    );
    res.json(rows.map(mapBatch));
  });

  // Documents by batch
  app.get('/api/batches/:id/documents', (req, res) => {
    const id = Number(req.params.id);
    const rows = all('SELECT id, batchId, title, url, version, requiresSignature, driveId, itemId, source FROM documents WHERE batchId=? ORDER BY id', [id]);
    res.json(rows.map(mapDoc));
  });

  // Recipients by batch (convenience for verification and UI)
  app.get('/api/batches/:id/recipients', (req, res) => {
    const id = Number(req.params.id);
    const rows = all('SELECT id, batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup FROM recipients WHERE batchId=? ORDER BY id DESC', [id]);
    res.json(rows);
  });

  // Completion status for all recipients in a batch
  // Returns [{ email, displayName, department, jobTitle, location, primaryGroup, businessId, businessName, acknowledged, total, completed, completionAt }]
  app.get('/api/batches/:id/completions', (req, res) => {
    try {
      const id = Number(req.params.id);
      if (!Number.isFinite(id)) return res.status(400).json({ error: 'invalid_batch_id' });
      const totalRow = one('SELECT COUNT(*) as c FROM documents WHERE batchId=?', [id]);
      const total = totalRow?.c || 0;
      if (total === 0) return res.json([]);

      const rows = all(
        `SELECT LOWER(r.email) as email,
                COALESCE(r.displayName, r.email) as displayName,
                r.department, r.jobTitle, r.location, r.primaryGroup,
                r.businessId,
                b.name as businessName,
                COUNT(CASE WHEN a.acknowledged=1 THEN 1 END) as acknowledged,
                MAX(a.ackDate) as lastAckDate
         FROM recipients r
         LEFT JOIN acks a ON a.batchId=r.batchId AND LOWER(a.email)=LOWER(r.email)
         LEFT JOIN businesses b ON b.id=r.businessId
         WHERE r.batchId=?
         GROUP BY LOWER(r.email), r.displayName, r.department, r.jobTitle, r.location, r.primaryGroup, r.businessId, b.name`,
        [id]
      );
      const mapped = rows.map(r => {
        const acknowledged = Number(r.acknowledged) || 0;
        const completed = acknowledged >= total;
        const completionAt = completed ? (r.lastAckDate || null) : null;
        return {
          email: String(r.email || ''),
          displayName: r.displayName || r.email || '',
          department: r.department || null,
          jobTitle: r.jobTitle || null,
          location: r.location || null,
          primaryGroup: r.primaryGroup || null,
          businessId: r.businessId != null ? Number(r.businessId) : null,
          businessName: r.businessName || null,
          acknowledged,
          total,
          completed,
          completionAt
        };
      });
      res.json(mapped);
    } catch (e) {
      console.error('completions endpoint failed', e);
      res.status(500).json({ error: 'completions_failed' });
    }
  });

  // Acked doc ids for user
  app.get('/api/batches/:id/acks', (req, res) => {
    const id = Number(req.params.id);
    const email = (req.query.email || '').toString().toLowerCase();
    const rows = all('SELECT documentId FROM acks WHERE batchId=? AND LOWER(email)=? AND acknowledged=1', [id, email]);
    res.json({ ids: rows.map(r => String(r.documentId)) });
  });

  // Progress for a user in a batch
  app.get('/api/batches/:id/progress', (req, res) => {
    const id = Number(req.params.id);
    const email = (req.query.email || '').toString().toLowerCase();
    const totalRow = one('SELECT COUNT(*) as c FROM documents WHERE batchId=?', [id]);
    const ackRow = one('SELECT COUNT(*) as c FROM acks WHERE batchId=? AND LOWER(email)=? AND acknowledged=1', [id, email]);
    const total = totalRow?.c || 0;
    const acknowledged = ackRow?.c || 0;
    const percent = total === 0 ? 0 : Math.round((acknowledged / total) * 100);
    res.json({ acknowledged, total, percent });
  });

  // Admin: create batch
  app.post('/api/batches', (req, res) => {
    const { logger } = req;
    
    try {
      logger.info('batch-create', 'Starting batch creation process');
      
      const { name, startDate = null, dueDate = null, description = null, status = 1 } = req.body || {};
      
      // Validate required fields
      if (!name || typeof name !== 'string' || !name.trim()) {
        logger.error('batch-create', 'Validation failed: name is required', { providedName: name });
        return res.status(400).json({ error: 'name_required', message: 'Batch name is required and must be a non-empty string' });
      }
      
      // Validate optional fields
      if (startDate !== null && startDate !== '' && typeof startDate !== 'string') {
        logger.error('batch-create', 'Validation failed: invalid startDate format', { startDate });
        return res.status(400).json({ error: 'invalid_start_date', message: 'Start date must be a string or null' });
      }
      
      if (dueDate !== null && dueDate !== '' && typeof dueDate !== 'string') {
        logger.error('batch-create', 'Validation failed: invalid dueDate format', { dueDate });
        return res.status(400).json({ error: 'invalid_due_date', message: 'Due date must be a string or null' });
      }
      
      const trimmedName = name.trim();
      const finalStartDate = (startDate === '' || startDate === null) ? null : startDate;
      const finalDueDate = (dueDate === '' || dueDate === null) ? null : dueDate;
      const finalDescription = (description === '' || description === null) ? null : description;
      const finalStatus = Number.isInteger(status) ? status : 1;
      
      logger.debug('batch-create', 'Validated input parameters', {
        name: trimmedName,
        startDate: finalStartDate,
        dueDate: finalDueDate,
        description: finalDescription ? 'provided' : 'null',
        status: finalStatus
      });
      
      // Insert into database
      logger.info('batch-create', 'Inserting batch into database');
      const sql = 'INSERT INTO batches (name, startDate, dueDate, status, description) VALUES (?, ?, ?, ?, ?)';
      const params = [trimmedName, finalStartDate, finalDueDate, finalStatus, finalDescription];
      
      logger.debug('batch-create', 'Executing SQL insert', { sql, params });
      
      const ok = exec(sql, params);
      if (!ok) {
        logger.error('batch-create', 'Database insert failed', { sql, params });
        return res.status(500).json({ error: 'insert_failed', message: 'Failed to insert batch into database' });
      }
      
      // Get the generated ID
      const idResult = one('SELECT last_insert_rowid() as id');
      const id = idResult?.id;
      
      if (!id) {
        logger.error('batch-create', 'Failed to retrieve generated batch ID');
        return res.status(500).json({ error: 'id_retrieval_failed', message: 'Batch created but ID could not be retrieved' });
      }
      
      // Verify batch exists and is accessible
      const verifyBatch = one('SELECT id, name FROM batches WHERE id = ?', [id]);
      if (!verifyBatch) {
        logger.error('batch-create', 'Batch verification failed - batch not found after creation', { batchId: id });
        return res.status(500).json({ error: 'verification_failed', message: 'Batch created but verification failed' });
      }
      
      logger.info('batch-create', 'Batch created and verified successfully', {
        batchId: id,
        name: trimmedName,
        startDate: finalStartDate,
        dueDate: finalDueDate,
        verifiedName: verifyBatch.name
      });
      
      res.json({ id, batchId: id });
      
    } catch (error) {
      logger.error('batch-create', 'Unexpected error during batch creation', {
        error: error.message,
        stack: error.stack
      });
      res.status(500).json({ error: 'internal_error', message: 'An unexpected error occurred during batch creation' });
    }
  });

  // Admin: create batch WITH documents and recipients atomically
  app.post('/api/batches/full', (req, res) => {
    const { logger } = req;
    try {
      logger.info('batch-full-create', 'Starting full batch creation process');

      const body = req.body || {};
      const { name, startDate = null, dueDate = null, description = null, status = 1 } = body.batch || body;
      const documents = Array.isArray(body.documents) ? body.documents : [];
      const recipients = Array.isArray(body.recipients) ? body.recipients : [];

      // Validate batch
      if (!name || typeof name !== 'string' || !name.trim()) {
        logger.error('batch-full-create', 'Validation failed: name is required', { providedName: name });
        return res.status(400).json({ error: 'name_required', message: 'Batch name is required and must be a non-empty string' });
      }
      const trimmedName = name.trim();
      const finalStartDate = (startDate === '' || startDate === null) ? null : startDate;
      const finalDueDate = (dueDate === '' || dueDate === null) ? null : dueDate;
      const finalDescription = (description === '' || description === null) ? null : description;
      const finalStatus = Number.isInteger(status) ? status : 1;

      // Enforce at least one document and one recipient when creating
      if (documents.length === 0) {
        logger.warn('batch-full-create', 'No documents provided in full create');
        return res.status(400).json({ error: 'documents_required', message: 'At least one document is required to create a batch' });
      }
      if (recipients.length === 0) {
        logger.warn('batch-full-create', 'No recipients provided in full create');
        return res.status(400).json({ error: 'recipients_required', message: 'At least one recipient is required to create a batch' });
      }

      // Begin transaction
      logger.debug('batch-full-create', 'Beginning DB transaction');
      db.run('BEGIN');
      let newBatchId = null;
      let docsInserted = 0;
      let recsInserted = 0;

      try {
        // Insert batch
        db.run('INSERT INTO batches (name, startDate, dueDate, status, description) VALUES (?, ?, ?, ?, ?)',
          [trimmedName, finalStartDate, finalDueDate, finalStatus, finalDescription]);
        const idRow = one('SELECT last_insert_rowid() as id');
        newBatchId = idRow?.id;
        if (!newBatchId) throw new Error('failed_to_create_batch');

        // Insert documents
        for (let i = 0; i < documents.length; i++) {
          const d = documents[i] || {};
          const { title, url, version = 1, requiresSignature = 0, driveId = null, itemId = null, source = null } = d;
          if (!title || !url) continue;
          db.run('INSERT OR IGNORE INTO documents (batchId, title, url, version, requiresSignature, driveId, itemId, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?)',
            [newBatchId, String(title), String(url), Number(version) || 1, requiresSignature ? 1 : 0, driveId, itemId, source]);
          docsInserted++;
        }

        // Insert recipients
        const processedEmails = new Set();
        for (let i = 0; i < recipients.length; i++) {
          const r = recipients[i] || {};
          const { businessId = null, user = null, email = null, displayName = null, department = null, jobTitle = null, location = null, primaryGroup = null } = r;
          const emailLower = String(email || user || '').trim().toLowerCase();
          if (!emailLower || !emailLower.includes('@') || emailLower.length < 5) continue;
          if (processedEmails.has(emailLower)) continue;
          processedEmails.add(emailLower);
          db.run(`INSERT OR IGNORE INTO recipients (batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup)
                  VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`,
            [newBatchId, businessId, emailLower, emailLower, displayName, department, jobTitle, location, primaryGroup]);
          recsInserted++;
        }

        // Ensure we created relations
        if (docsInserted === 0) throw new Error('no_documents_created');
        if (recsInserted === 0) throw new Error('no_recipients_created');

        // Commit
        db.run('COMMIT');
        persist(db);
        logger.info('batch-full-create', 'Full batch creation successful', { batchId: newBatchId, docsInserted, recsInserted });
        return res.json({ id: newBatchId, batchId: newBatchId, documentsInserted: docsInserted, recipientsInserted: recsInserted });
      } catch (txErr) {
        // Rollback on any error
        try { db.run('ROLLBACK'); } catch {}
        logger.error('batch-full-create', 'Transaction failed, rolled back', { error: txErr?.message || String(txErr) });
        const code = (txErr?.message === 'no_documents_created') ? 400
                  : (txErr?.message === 'no_recipients_created') ? 400
                  : 500;
        return res.status(code).json({ error: txErr?.message || 'tx_failed' });
      }
    } catch (error) {
      logger.error('batch-full-create', 'Unexpected error during full batch creation', { error: error?.message || String(error) });
      return res.status(500).json({ error: 'internal_error' });
    }
  });

  // Admin: update batch
  app.put('/api/batches/:id', (req, res) => {
    const id = Number(req.params.id);
    const { name, startDate = null, dueDate = null, status, description = null } = req.body || {};
    try {
      const current = one('SELECT id, name, startDate, dueDate, status, description FROM batches WHERE id=?', [id]);
      if (!current) return res.status(404).json({ error: 'not_found' });
      const next = {
        name: name != null ? String(name).trim() : current.name,
        startDate: startDate !== undefined ? startDate : current.startDate,
        dueDate: dueDate !== undefined ? dueDate : current.dueDate,
        status: status != null ? Number(status) : current.status,
        description: description !== undefined ? description : current.description
      };
      db.run('UPDATE batches SET name=?, startDate=?, dueDate=?, status=?, description=? WHERE id=?', [next.name, next.startDate, next.dueDate, next.status, next.description, id]);
      persist(db);
      res.json({ ok: true });
    } catch (e) {
      console.error('Update batch failed', e);
      res.status(500).json({ error: 'update_failed' });
    }
  });

  // Admin: bulk add documents
  app.post('/api/batches/:id/documents', (req, res) => {
    const { logger } = req;
    
    try {
      logger.info('documents-create', 'Starting bulk document addition process');
      
      const id = Number(req.params.id);
      if (!Number.isInteger(id) || id <= 0) {
        logger.error('documents-create', 'Invalid batch ID provided', { providedId: req.params.id, parsedId: id });
        return res.status(400).json({ error: 'invalid_batch_id', message: 'Batch ID must be a positive integer' });
      }
      
      // Check if batch exists
      const batchExists = one('SELECT id, name FROM batches WHERE id = ?', [id]);
      if (!batchExists) {
        logger.error('documents-create', 'Batch not found', { batchId: id });
        return res.status(404).json({ error: 'batch_not_found', message: 'Specified batch does not exist' });
      }
      
      logger.info('documents-create', 'Batch verified for document insertion', { 
        batchId: id, 
        batchName: batchExists.name 
      });
      
      const docs = Array.isArray(req.body?.documents) ? req.body.documents : [];
      logger.info('documents-create', 'Processing document list', { 
        batchId: id, 
        totalDocuments: docs.length 
      });
      
      if (docs.length === 0) {
        logger.warn('documents-create', 'No documents provided in request');
        return res.json({ inserted: 0, message: 'No documents provided' });
      }
      
      let count = 0;
      let skipped = 0;
      const errors = [];
      
      logger.debug('documents-create', 'Starting database transaction');
      db.run('BEGIN');
      
      try {
        for (let i = 0; i < docs.length; i++) {
          const d = docs[i];
          const { title, url, version = 1, requiresSignature = 0, driveId = null, itemId = null, source = null } = d || {};
          
          // Validate document fields
          if (!title || !url) {
            logger.warn('documents-create', `Document ${i + 1} missing required fields`, { 
              index: i, 
              hasTitle: !!title, 
              hasUrl: !!url 
            });
            skipped++;
            errors.push(`Document ${i + 1}: Missing title or URL`);
            continue;
          }
          
          if (typeof title !== 'string' || typeof url !== 'string') {
            logger.warn('documents-create', `Document ${i + 1} has invalid field types`, { 
              index: i, 
              titleType: typeof title, 
              urlType: typeof url 
            });
            skipped++;
            errors.push(`Document ${i + 1}: Title and URL must be strings`);
            continue;
          }
          
          logger.debug('documents-create', `Processing document ${i + 1}`, {
            title: title.substring(0, 50) + (title.length > 50 ? '...' : ''),
            url: url.substring(0, 100) + (url.length > 100 ? '...' : ''),
            version,
            requiresSignature: !!requiresSignature
          });
          
          try {
            db.run('INSERT OR IGNORE INTO documents (batchId, title, url, version, requiresSignature, driveId, itemId, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?)', 
              [id, String(title), String(url), Number(version) || 1, requiresSignature ? 1 : 0, driveId, itemId, source]);
            count++;
          } catch (docError) {
            logger.error('documents-create', `Failed to insert document ${i + 1}`, {
              index: i,
              error: docError.message,
              title: title.substring(0, 50)
            });
            errors.push(`Document ${i + 1}: ${docError.message}`);
            skipped++;
          }
        }
        
        logger.debug('documents-create', 'Committing database transaction');
        db.run('COMMIT');
        persist(db);
        
        logger.info('documents-create', 'Document addition completed', {
          batchId: id,
          inserted: count,
          skipped: skipped,
          totalProcessed: docs.length
        });
        
        const result = { inserted: count };
        if (skipped > 0) {
          result.skipped = skipped;
          result.errors = errors;
        }
        
        res.json(result);
        
      } catch (transactionError) {
        logger.error('documents-create', 'Transaction failed, rolling back', {
          error: transactionError.message,
          batchId: id,
          processedCount: count
        });
        db.run('ROLLBACK');
        throw transactionError;
      }
      
    } catch (e) {
      logger.error('documents-create', 'Unexpected error during document creation', {
        error: e.message,
        stack: e.stack,
        batchId: req.params.id
      });
      
      try {
        db.run('ROLLBACK');
      } catch (rollbackError) {
        logger.error('documents-create', 'Failed to rollback transaction', { error: rollbackError.message });
      }
      
      res.status(500).json({ error: 'insert_failed', message: 'Failed to add documents to batch' });
    }
  });

  // Admin: bulk add recipients
  app.post('/api/batches/:id/recipients', (req, res) => {
    const { logger } = req;
    
    try {
      logger.info('recipients-create', 'Starting bulk recipient addition process');
      
      const id = Number(req.params.id);
      if (!Number.isInteger(id) || id <= 0) {
        logger.error('recipients-create', 'Invalid batch ID provided', { providedId: req.params.id, parsedId: id });
        return res.status(400).json({ error: 'invalid_batch_id', message: 'Batch ID must be a positive integer' });
      }
      
      // Check if batch exists
      const batchExists = one('SELECT id, name FROM batches WHERE id = ?', [id]);
      if (!batchExists) {
        logger.error('recipients-create', 'Batch not found', { batchId: id });
        return res.status(404).json({ error: 'batch_not_found', message: 'Specified batch does not exist' });
      }
      
      logger.info('recipients-create', 'Batch verified for recipient insertion', { 
        batchId: id, 
        batchName: batchExists.name 
      });
      
      const list = Array.isArray(req.body?.recipients) ? req.body.recipients : [];
      logger.info('recipients-create', 'Processing recipient list', { 
        batchId: id, 
        totalRecipients: list.length 
      });
      
      if (list.length === 0) {
        logger.warn('recipients-create', 'No recipients provided in request');
        return res.json({ inserted: 0, message: 'No recipients provided' });
      }
      
      let count = 0;
      let skipped = 0;
      const errors = [];
      const processedEmails = new Set();
      
      logger.debug('recipients-create', 'Starting database transaction');
      db.run('BEGIN');
      
      try {
        for (let i = 0; i < list.length; i++) {
          const r = list[i];
          const { businessId = null, user = null, email = null, displayName = null, department = null, jobTitle = null, location = null, primaryGroup = null } = r || {};
          
          const emailRaw = email || user || '';
          const emailLower = String(emailRaw).trim().toLowerCase();
          
          // Validate email
          if (!emailLower) {
            logger.warn('recipients-create', `Recipient ${i + 1} missing email`, { 
              index: i,
              providedEmail: emailRaw,
              providedUser: user
            });
            skipped++;
            errors.push(`Recipient ${i + 1}: Missing email address`);
            continue;
          }
          
          // Basic email format validation
          if (!emailLower.includes('@') || emailLower.length < 5) {
            logger.warn('recipients-create', `Recipient ${i + 1} has invalid email format`, { 
              index: i,
              email: emailLower
            });
            skipped++;
            errors.push(`Recipient ${i + 1}: Invalid email format`);
            continue;
          }
          
          // Check for duplicates in current batch
          if (processedEmails.has(emailLower)) {
            logger.warn('recipients-create', `Recipient ${i + 1} is duplicate in current request`, { 
              index: i,
              email: emailLower
            });
            skipped++;
            errors.push(`Recipient ${i + 1}: Duplicate email in request`);
            continue;
          }
          
          processedEmails.add(emailLower);
          
          logger.debug('recipients-create', `Processing recipient ${i + 1}`, {
            email: emailLower,
            displayName: displayName || 'none',
            businessId: businessId || 'none',
            department: department || 'none'
          });
          
          try {
            db.run(`INSERT OR IGNORE INTO recipients (batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`, 
                    [id, businessId, emailLower, emailLower, displayName, department, jobTitle, location, primaryGroup]);
            count++;
          } catch (recipientError) {
            logger.error('recipients-create', `Failed to insert recipient ${i + 1}`, {
              index: i,
              email: emailLower,
              error: recipientError.message
            });
            errors.push(`Recipient ${i + 1}: ${recipientError.message}`);
            skipped++;
          }
        }
        
        logger.debug('recipients-create', 'Committing database transaction');
        db.run('COMMIT');
        persist(db);
        
        logger.info('recipients-create', 'Recipient addition completed', {
          batchId: id,
          inserted: count,
          skipped: skipped,
          totalProcessed: list.length,
          uniqueEmails: processedEmails.size
        });
        
        const result = { inserted: count };
        if (skipped > 0) {
          result.skipped = skipped;
          result.errors = errors;
        }
        
        res.json(result);
        
      } catch (transactionError) {
        logger.error('recipients-create', 'Transaction failed, rolling back', {
          error: transactionError.message,
          batchId: id,
          processedCount: count
        });
        db.run('ROLLBACK');
        throw transactionError;
      }
      
    } catch (e) {
      logger.error('recipients-create', 'Unexpected error during recipient creation', {
        error: e.message,
        stack: e.stack,
        batchId: req.params.id
      });
      
      try {
        db.run('ROLLBACK');
      } catch (rollbackError) {
        logger.error('recipients-create', 'Failed to rollback transaction', { error: rollbackError.message });
      }
      
      res.status(500).json({ error: 'insert_failed', message: 'Failed to add recipients to batch' });
    }
  });

  // Admin: delete batch (cascade delete related data)
  app.delete('/api/batches/:id', (req, res) => {
    const id = Number(req.params.id);
    try {
      db.run('BEGIN');
      db.run('DELETE FROM acks WHERE batchId=?', [id]);
      db.run('DELETE FROM documents WHERE batchId=?', [id]);
      db.run('DELETE FROM recipients WHERE batchId=?', [id]);
      db.run('DELETE FROM batches WHERE id=?', [id]);
      db.run('COMMIT');
      persist(db);
      res.json({ ok: true });
    } catch (e) {
      try { db.run('ROLLBACK'); } catch {}
      console.error('Delete batch failed', e);
      res.status(500).json({ error: 'delete_failed' });
    }
  });

  // Roles management API
  // List roles
  app.get('/api/roles', (_req, res) => {
    try {
      const rows = all('SELECT id, email, role, createdAt FROM roles ORDER BY role, LOWER(email)');
      res.json(rows.map(r => ({ id: r.id, email: String(r.email).toLowerCase(), role: String(r.role), createdAt: r.createdAt })));
    } catch (e) {
      console.error('List roles failed', e);
      res.status(500).json({ error: 'list_failed' });
    }
  });
  // Create role
  app.post('/api/roles', (req, res) => {
    try {
      const { email, role } = req.body || {};
      const e = String(email || '').trim().toLowerCase();
      const r = String(role || '').trim();
      if (!e || !e.includes('@')) return res.status(400).json({ error: 'invalid_email' });
      // Allow only Admin or Manager via API to avoid accidental grant of SuperAdmin; env remains authoritative for SuperAdmin
      if (!['Admin','Manager'].includes(r)) return res.status(400).json({ error: 'invalid_role' });
      const now = new Date().toISOString();
      const ok = exec('INSERT OR IGNORE INTO roles (email, role, createdAt) VALUES (?, ?, ?)', [e, r, now]);
      if (!ok) return res.status(400).json({ error: 'insert_failed' });
      const id = one('SELECT last_insert_rowid() as id')?.id;
      res.json({ id, email: e, role: r, createdAt: now });
    } catch (e) {
      console.error('Create role failed', e);
      res.status(500).json({ error: 'insert_failed' });
    }
  });
  // Delete role
  app.delete('/api/roles/:id', (req, res) => {
    try {
      const id = Number(req.params.id);
      if (!Number.isInteger(id) || id <= 0) return res.status(400).json({ error: 'invalid_id' });
      const ok = exec('DELETE FROM roles WHERE id=?', [id]);
      if (!ok) return res.status(400).json({ error: 'delete_failed' });
      res.json({ ok: true });
    } catch (e) {
      console.error('Delete role failed', e);
      res.status(500).json({ error: 'delete_failed' });
    }
  });

  // Acknowledge a document
  app.post('/api/ack', (req, res) => {
    const { batchId, documentId, email } = req.body || {};
    if (!batchId || !documentId || !email) return res.status(400).json({ error: 'missing_fields' });
    // Idempotent: delete existing then insert
    db.run('DELETE FROM acks WHERE batchId=? AND documentId=? AND LOWER(email)=?', [batchId, documentId, String(email).toLowerCase()]);
    const now = new Date().toISOString();
    const ok = exec('INSERT INTO acks (batchId, documentId, email, acknowledged, ackDate) VALUES (?, ?, ?, 1, ?)', [batchId, documentId, String(email).toLowerCase(), now]);
    if (!ok) return res.status(400).json({ error: 'insert_failed' });
    res.json({ ok: true });
  });

  // Seed sample data for a specific user email
  app.post('/api/seed', (req, res) => {
    const email = (req.query.email || req.body?.email || '').toString().trim().toLowerCase();
    if (!email) return res.status(400).json({ error: 'email_required' });
    try {
      db.run('BEGIN');
      // Create a batch
      const name = 'Demo Batch';
      const startDate = new Date().toISOString().substring(0,10);
      const dueDate = new Date(Date.now() + 7*24*60*60*1000).toISOString().substring(0,10);
      db.run('INSERT INTO batches (name, startDate, dueDate, status, description) VALUES (?, ?, ?, 1, ?)', [name, startDate, dueDate, 'Seeded demo batch']);
      const batchId = one('SELECT last_insert_rowid() as id')?.id;
      // Add two docs
      db.run('INSERT INTO documents (batchId, title, url, version, requiresSignature) VALUES (?, ?, ?, ?, ?)', [batchId, 'Code of Conduct', 'https://www.w3.org/WAI/ER/tests/xhtml/testfiles/resources/pdf/dummy.pdf', 1, 0]);
      db.run('INSERT INTO documents (batchId, title, url, version, requiresSignature) VALUES (?, ?, ?, ?, ?)', [batchId, 'IT Security Policy', 'https://www.africau.edu/images/default/sample.pdf', 1, 0]);
      // Add recipient (user)
      db.run(`INSERT INTO recipients (batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup)
              VALUES (?, NULL, ?, ?, ?, NULL, NULL, NULL, NULL)`, [batchId, email, email, 'Demo User']);
      db.run('COMMIT');
      persist(db);
      res.json({ ok: true, batchId });
    } catch (e) {
      try { db.run('ROLLBACK'); } catch {}
      console.error('Seed failed', e);
      res.status(500).json({ error: 'seed_failed' });
    }
  });

  app.listen(PORT, () => {
    console.log(`SQLite API listening on http://localhost:${PORT}`);
  });
}

function mapBatch(r) {
  return {
    toba_batchid: String(r.id),
    toba_name: r.name,
    toba_startdate: r.startDate || null,
    toba_duedate: r.dueDate || null,
    toba_status: r.status != null ? String(r.status) : null
  };
}
function mapDoc(r) {
  return {
    toba_documentid: String(r.id),
    toba_title: r.title,
    toba_version: r.version != null ? String(r.version) : '1',
    toba_requiressignature: !!r.requiresSignature,
    toba_fileurl: r.url,
    toba_driveid: r.driveId || null,
    toba_itemid: r.itemId || null,
    toba_source: r.source || null
  };
}

function bootstrapSchema(db) {
  db.run(`CREATE TABLE IF NOT EXISTS notification_emails (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL UNIQUE
  );`);
  try { db.run('PRAGMA foreign_keys = ON'); } catch {}
  db.run(`CREATE TABLE IF NOT EXISTS businesses (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    code TEXT,
    isActive INTEGER DEFAULT 1,
    description TEXT
  );`);
  db.run(`CREATE TABLE IF NOT EXISTS batches (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    startDate TEXT,
    dueDate TEXT,
    status INTEGER DEFAULT 1,
    description TEXT
  );`);
  db.run(`CREATE TABLE IF NOT EXISTS documents (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batchId INTEGER NOT NULL,
    title TEXT NOT NULL,
    url TEXT NOT NULL,
    version INTEGER DEFAULT 1,
    requiresSignature INTEGER DEFAULT 0,
    driveId TEXT,
    itemId TEXT,
    source TEXT,
    FOREIGN KEY (batchId) REFERENCES batches(id) ON DELETE CASCADE
  );`);
  db.run(`CREATE TABLE IF NOT EXISTS recipients (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batchId INTEGER NOT NULL,
    businessId INTEGER,
    user TEXT,
    email TEXT,
    displayName TEXT,
    department TEXT,
    jobTitle TEXT,
    location TEXT,
    primaryGroup TEXT,
    FOREIGN KEY (batchId) REFERENCES batches(id) ON DELETE CASCADE,
    FOREIGN KEY (businessId) REFERENCES businesses(id) ON DELETE SET NULL
  );`);
  db.run(`CREATE TABLE IF NOT EXISTS acks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    batchId INTEGER NOT NULL,
    documentId INTEGER NOT NULL,
    email TEXT NOT NULL,
    acknowledged INTEGER DEFAULT 1,
    ackDate TEXT,
    FOREIGN KEY (batchId) REFERENCES batches(id) ON DELETE CASCADE,
    FOREIGN KEY (documentId) REFERENCES documents(id) ON DELETE CASCADE
  );`);

  // Indexes and uniqueness constraints
  db.run(`CREATE INDEX IF NOT EXISTS idx_documents_batch ON documents(batchId);`);
  db.run(`CREATE INDEX IF NOT EXISTS idx_recipients_batch ON recipients(batchId);`);
  db.run(`CREATE INDEX IF NOT EXISTS idx_acks_batch ON acks(batchId);`);
  db.run(`CREATE INDEX IF NOT EXISTS idx_acks_doc ON acks(documentId);`);
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_recipients_batch_email ON recipients(batchId, LOWER(email));`);
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_documents_batch_url ON documents(batchId, url);`);

  // Roles table for RBAC overrides (DB-managed roles)
  db.run(`CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL,
    role TEXT NOT NULL,
    createdAt TEXT
  );`);
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_roles_email_role ON roles(LOWER(email), role);`);

  // Role-based and user-based permissions
  db.run(`CREATE TABLE IF NOT EXISTS role_permissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    role TEXT NOT NULL,
    permKey TEXT NOT NULL,
    value INTEGER NOT NULL DEFAULT 1
  );`);
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_role_perm ON role_permissions(LOWER(role), permKey);`);
  db.run(`CREATE TABLE IF NOT EXISTS user_permissions (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL,
    permKey TEXT NOT NULL,
    value INTEGER NOT NULL DEFAULT 1
  );`);
  db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_user_perm ON user_permissions(LOWER(email), permKey);`);

  // Seed a default business for convenience
  db.run("INSERT INTO businesses (name, code, isActive, description) VALUES ('Default Business', 'DEF', 1, 'Auto-created')");
}

// Best-effort migrations for existing databases (adds new columns if missing)
function migrateSchema(db) {
  try { db.run("ALTER TABLE documents ADD COLUMN driveId TEXT"); } catch {}
  try { db.run("ALTER TABLE documents ADD COLUMN itemId TEXT"); } catch {}
  try { db.run("ALTER TABLE documents ADD COLUMN source TEXT"); } catch {}
  // roles table added in later versions
  try { db.run(`CREATE TABLE IF NOT EXISTS roles (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    email TEXT NOT NULL,
    role TEXT NOT NULL,
    createdAt TEXT
  );`); } catch {}
  try { db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_roles_email_role ON roles(LOWER(email), role);`); } catch {}
  // Permissions tables
  try { db.run(`CREATE TABLE IF NOT EXISTS role_permissions (id INTEGER PRIMARY KEY AUTOINCREMENT, role TEXT NOT NULL, permKey TEXT NOT NULL, value INTEGER NOT NULL DEFAULT 1);`); } catch {}
  try { db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_role_perm ON role_permissions(LOWER(role), permKey);`); } catch {}
  try { db.run(`CREATE TABLE IF NOT EXISTS user_permissions (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT NOT NULL, permKey TEXT NOT NULL, value INTEGER NOT NULL DEFAULT 1);`); } catch {}
  try { db.run(`CREATE UNIQUE INDEX IF NOT EXISTS ux_user_perm ON user_permissions(LOWER(email), permKey);`); } catch {}
}

function persist(db) {
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(DB_PATH, buffer);
}

// Resolve user roles using DB roles and environment. Groups not evaluated server-side.
function resolveUserRoles(email, db) {
  const e = String(email || '').trim().toLowerCase();
  const roles = ['Employee'];
  try {
    const envList = (s) => String(s || '')
      .split(',')
      .map(x => String(x).trim().toLowerCase())
      .filter(x => x && x.includes('@'));
    const superAdmins = envList(process.env.REACT_APP_SUPER_ADMINS);
    if (superAdmins.includes(e)) roles.push('SuperAdmin');
  } catch {}
  try {
    const rows = db ? db.prepare && db : null;
    const fromDb = (() => {
      try { return (db ? (function(){ const rows = []; const stmt = db.prepare('SELECT role FROM roles WHERE LOWER(email)=LOWER(?)'); try { stmt.bind([e]); while (stmt.step()) rows.push(stmt.getAsObject()); } finally { stmt.free(); } return rows; })() : []); } catch { return []; }
    })();
    for (const r of fromDb) {
      const role = String(r.role);
      if (!roles.includes(role)) roles.push(role);
    }
  } catch {}
  return roles;
}

start().catch(err => {
  console.error('Failed to start SQLite API', err);
  process.exit(1);
});
