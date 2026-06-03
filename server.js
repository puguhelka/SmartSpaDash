// SmartSpaDash Server — Express + SQLite (sql.js)
// Lelap Mom Baby Care Salatiga

const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
app.use(cors());
app.use(express.json());

// ── Database (sql.js — pure JS, no native build) ──
const initSqlJs = require('sql.js');
const DB_PATH = path.join(__dirname, 'data.db');

let db;

async function initDB() {
  const SQL = await initSqlJs();
  
  // Load existing database or create new
  if (fs.existsSync(DB_PATH)) {
    const fileBuffer = fs.readFileSync(DB_PATH);
    db = new SQL.Database(fileBuffer);
    console.log('Database loaded from disk');
  } else {
    db = new SQL.Database();
    console.log('New database created');
  }
  
  // Create schema
  db.run(`
    CREATE TABLE IF NOT EXISTS store (
      id TEXT PRIMARY KEY,
      resource TEXT NOT NULL,
      data TEXT NOT NULL DEFAULT '{}',
      created_at TEXT NOT NULL DEFAULT (datetime('now')),
      updated_at TEXT NOT NULL DEFAULT (datetime('now'))
    )
  `);
  db.run('CREATE INDEX IF NOT EXISTS idx_resource ON store(resource)');
  
  // Persist to disk periodically and on important writes
  saveDB();
}

function saveDB() {
  try {
    const data = db.export();
    const buffer = Buffer.from(data);
    fs.writeFileSync(DB_PATH, buffer);
  } catch(e) {
    console.error('Failed to save DB:', e.message);
  }
}

// Auto-save every 60 seconds
setInterval(saveDB, 60000);

// ── Helpers ──
function uid() {
  return Date.now().toString(36) + crypto.randomBytes(6).toString('hex');
}

function nowISO() {
  return new Date().toISOString();
}

function readAll(resource) {
  const stmt = db.prepare('SELECT * FROM store WHERE resource = ?');
  stmt.bind([resource]);
  const items = [];
  while (stmt.step()) {
    const row = stmt.getAsObject();
    try { items.push(JSON.parse(row.data)); } catch { items.push({}); }
  }
  stmt.free();
  items.sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
  return items;
}

function getOne(resource, id) {
  const stmt = db.prepare('SELECT * FROM store WHERE resource = ? AND id = ?');
  stmt.bind([resource, id]);
  if (stmt.step()) {
    const row = stmt.getAsObject();
    stmt.free();
    try { return JSON.parse(row.data); } catch { return null; }
  }
  stmt.free();
  return null;
}

function saveOne(resource, id, data) {
  const iso = nowISO();
  const existing = getOne(resource, id);
  if (existing) {
    const merged = { ...existing, ...data, updated_at: iso };
    db.run('UPDATE store SET data = ?, updated_at = ? WHERE resource = ? AND id = ?', 
      [JSON.stringify(merged), iso, resource, id]);
    saveDB();
    return merged;
  } else {
    const item = { id, ...data, created_at: iso, updated_at: iso };
    db.run('INSERT OR REPLACE INTO store (id, resource, data, created_at, updated_at) VALUES (?, ?, ?, ?, ?)',
      [id, resource, JSON.stringify(item), iso, iso]);
    saveDB();
    return item;
  }
}

function deleteOne(resource, id) {
  db.run('DELETE FROM store WHERE resource = ? AND id = ?', [resource, id]);
  saveDB();
}

function verifyToken(req) {
  const auth = (req.headers.authorization || '').replace('Bearer ', '');
  if (!auth) return null;
  try {
    return JSON.parse(Buffer.from(auth, 'base64').toString());
  } catch { return null; }
}

function findUserByEmail(email) {
  const stmt = db.prepare("SELECT * FROM store WHERE resource = 'users'");
  const rows = [];
  while (stmt.step()) rows.push(stmt.getAsObject());
  stmt.free();
  return rows.find(r => {
    try { const d = JSON.parse(r.data); return d.email === email; } catch { return false; }
  });
}

// ── Auth Routes ──
app.post('/api/auth', (req, res) => {
  const { action, email, password, token } = req.body;
  
  if (action === 'login') {
    const user = findUserByEmail(email);
    if (!user) return res.status(401).json({ error: 'Email/password salah' });
    const ud = JSON.parse(user.data);
    if (ud.password !== password) return res.status(401).json({ error: 'Email/password salah' });
    const tok = Buffer.from(JSON.stringify({ id: user.id, role: ud.role, name: ud.name })).toString('base64');
    return res.json({ token: tok, user: { id: user.id, name: ud.name, email: ud.email, role: ud.role } });
  }
  
  if (action === 'me') {
    const tok = (token || '').replace('Bearer ', '');
    if (!tok) return res.status(401).json({ error: 'Unauthorized' });
    try {
      const decoded = JSON.parse(Buffer.from(tok, 'base64').toString());
      const user = getOne('users', decoded.id);
      if (!user) return res.status(401).json({ error: 'User not found' });
      return res.json({ user: { id: decoded.id, name: user.name, email: user.email, role: user.role } });
    } catch { return res.status(401).json({ error: 'Invalid token' }); }
  }
  
  return res.status(400).json({ error: 'Invalid action' });
});

// Backward compatibility
app.post('/api/login', (req, res) => {
  const { email, password } = req.body;
  const user = findUserByEmail(email);
  if (!user) return res.status(401).json({ error: 'Email/password salah' });
  const ud = JSON.parse(user.data);
  if (ud.password !== password) return res.status(401).json({ error: 'Email/password salah' });
  const tok = Buffer.from(JSON.stringify({ id: user.id, role: ud.role, name: ud.name })).toString('base64');
  return res.json({ token: tok, user: { id: user.id, name: ud.name, email: ud.email, role: ud.role } });
});

app.post('/api/me', (req, res) => {
  const tok = (req.headers.authorization || '').replace('Bearer ', '');
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  try {
    const decoded = JSON.parse(Buffer.from(tok, 'base64').toString());
    const user = getOne('users', decoded.id);
    if (!user) return res.status(401).json({ error: 'User not found' });
    return res.json({ user: { id: decoded.id, name: user.name, email: user.email, role: user.role } });
  } catch { return res.status(401).json({ error: 'Invalid token' }); }
});

// ── Dashboard ──
app.get('/api/dashboard', (req, res) => {
  const clients = readAll('clients');
  const apps = readAll('appointments');
  const services = readAll('services');
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const monthStart = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  const uniqueMap = {};
  clients.forEach(c => {
    const key = (c.name || '') + '|' + (c.phone || '');
    if (key !== '|') uniqueMap[key] = c;
  });
  const uniqueClients = Object.values(uniqueMap);

  res.json({
    totalBookings: apps.filter(a => a.status === 'Selesai').length,
    bookingsBulanIni: apps.filter(a => a.status === 'Selesai' && a.date && a.date.startsWith(monthStart)).length,
    bookingsHariIni: apps.filter(a => a.status === 'Selesai' && a.date === today).length,
    draftBookings: apps.filter(a => (a.status === 'Pending' || a.status === 'Booking' || a.status === 'Menunggu') && a.date === today).length,
    totalServices: services.length,
    totalClients: uniqueClients.length,
    clientsBulanIni: uniqueClients.filter(c => (c.created_at || '').startsWith(monthStart)).length,
    recentAppointments: apps.slice(0, 5)
  });
});

// ── CRUD for all resources ──
const resources = ['clients', 'appointments', 'services', 'staff', 'products', 'transactions', 'reports', 'users', 'homecare', 'customer_types'];

resources.forEach(resource => {
  const base = '/api/' + resource;
  
  app.get(base, (req, res) => {
    res.json(readAll(resource));
  });
  
  app.get(base + '/:id', (req, res) => {
    const item = getOne(resource, req.params.id);
    if (!item) return res.status(404).json({ error: 'Not found' });
    res.json(item);
  });
  
  app.post(base, (req, res) => {
    if (resource === 'users') {
      const tok = verifyToken(req);
      if (!tok) return res.status(401).json({ error: 'Unauthorized' });
      const user = getOne('users', tok.id);
      if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    const item = saveOne(resource, uid(), req.body);
    res.status(201).json(item);
  });
  
  app.put(base + '/:id', (req, res) => {
    if (resource === 'users') {
      const tok = verifyToken(req);
      if (!tok) return res.status(401).json({ error: 'Unauthorized' });
      const user = getOne('users', tok.id);
      if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    const existing = getOne(resource, req.params.id);
    if (!existing) return res.status(404).json({ error: 'Not found' });
    const item = saveOne(resource, req.params.id, req.body);
    res.json(item);
  });
  
  app.delete(base + '/:id', (req, res) => {
    if (resource === 'users') {
      const tok = verifyToken(req);
      if (!tok) return res.status(401).json({ error: 'Unauthorized' });
      const user = getOne('users', tok.id);
      if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    deleteOne(resource, req.params.id);
    res.json({ success: true });
  });
});

// ── Serve Static Frontend ──
app.use(express.static(path.join(__dirname, 'public')));

// SPA fallback
app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Start ──
const PORT = process.env.PORT || 3000;
initDB().then(() => {
  // Ensure default admin
  const admin = findUserByEmail('puguh.legowo.k@gmail.com');
  if (!admin) {
    saveOne('users', uid(), { name: 'Admin', email: 'puguh.legowo.k@gmail.com', password: 'Admin123!', role: 'admin' });
    console.log('Default admin created');
  }
  
  app.listen(PORT, '0.0.0.0', () => {
    console.log(`SmartSpaDash running on port ${PORT}`);
  });
}).catch(err => {
  console.error('Failed to initialize database:', err);
  process.exit(1);
});
