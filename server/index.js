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

  const app = express();
  app.use(cors());
  app.use(express.json({ limit: '2mb' }));

  // Utilities
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

  // Ensure at least one business exists (after helpers available)
  try {
    const cnt = one('SELECT COUNT(*) as c FROM businesses')?.c || 0;
    if (cnt === 0) {
      db.run("INSERT INTO businesses (name, code, isActive, description) VALUES ('Default Business', 'DEF', 1, 'Auto-created')");
      persist(db);
    }
  } catch (e) { console.warn('Business seed check failed (non-fatal):', e); }

  // Routes
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

      const r = https.request(url, { method: 'GET', headers: { 'Authorization': bearer, 'User-Agent': 'Sunbeth-Graph-Proxy/1.0' } }, (upstream) => {
        // Forward content type and stream
        const ct = upstream.headers['content-type'] || 'application/octet-stream';
        res.setHeader('Content-Type', ct);
        res.setHeader('Cache-Control', 'no-store');
        upstream.on('error', () => { try { res.destroy(); } catch {} });
        upstream.pipe(res);
      });
      r.on('error', () => { if (!res.headersSent) res.status(502).json({ error: 'upstream_error' }); else try { res.destroy(); } catch {} });
      r.end();
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
    const { name, startDate = null, dueDate = null, description = null, status = 1 } = req.body || {};
    const ok = exec('INSERT INTO batches (name, startDate, dueDate, status, description) VALUES (?, ?, ?, ?, ?)', [name, startDate, dueDate, status, description]);
    if (!ok) return res.status(400).json({ error: 'insert_failed' });
    const id = one('SELECT last_insert_rowid() as id')?.id;
    res.json({ id });
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
    const id = Number(req.params.id);
    const docs = Array.isArray(req.body?.documents) ? req.body.documents : [];
    let count = 0;
    try {
      db.run('BEGIN');
      for (const d of docs) {
        const { title, url, version = 1, requiresSignature = 0, driveId = null, itemId = null, source = null } = d || {};
        if (!title || !url) continue;
        db.run('INSERT OR IGNORE INTO documents (batchId, title, url, version, requiresSignature, driveId, itemId, source) VALUES (?, ?, ?, ?, ?, ?, ?, ?)', [id, String(title), String(url), Number(version) || 1, requiresSignature ? 1 : 0, driveId, itemId, source]);
        count++;
      }
      db.run('COMMIT');
      persist(db);
      res.json({ inserted: count });
    } catch (e) {
      db.run('ROLLBACK');
      console.error(e);
      res.status(400).json({ error: 'insert_failed' });
    }
  });

  // Admin: bulk add recipients
  app.post('/api/batches/:id/recipients', (req, res) => {
    const id = Number(req.params.id);
    const list = Array.isArray(req.body?.recipients) ? req.body.recipients : [];
    let count = 0;
    try {
      db.run('BEGIN');
      for (const r of list) {
        const { businessId = null, user = null, email = null, displayName = null, department = null, jobTitle = null, location = null, primaryGroup = null } = r || {};
        const emailLower = String(email || user || '').trim().toLowerCase();
        if (!emailLower) continue;
        db.run(`INSERT OR IGNORE INTO recipients (batchId, businessId, user, email, displayName, department, jobTitle, location, primaryGroup)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)`, [id, businessId, emailLower, emailLower, displayName, department, jobTitle, location, primaryGroup]);
        count++;
      }
      db.run('COMMIT');
      persist(db);
      res.json({ inserted: count });
    } catch (e) {
      db.run('ROLLBACK');
      console.error(e);
      res.status(400).json({ error: 'insert_failed' });
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

  // Seed a default business for convenience
  db.run("INSERT INTO businesses (name, code, isActive, description) VALUES ('Default Business', 'DEF', 1, 'Auto-created')");
}

// Best-effort migrations for existing databases (adds new columns if missing)
function migrateSchema(db) {
  try { db.run("ALTER TABLE documents ADD COLUMN driveId TEXT"); } catch {}
  try { db.run("ALTER TABLE documents ADD COLUMN itemId TEXT"); } catch {}
  try { db.run("ALTER TABLE documents ADD COLUMN source TEXT"); } catch {}
}

function persist(db) {
  const data = db.export();
  const buffer = Buffer.from(data);
  fs.writeFileSync(DB_PATH, buffer);
}

start().catch(err => {
  console.error('Failed to start SQLite API', err);
  process.exit(1);
});
