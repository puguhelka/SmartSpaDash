// SmartSpaDash Server — Express + JSON File Storage
// Lelap Mom Baby Care Salatiga

const express = require('express');
const cors = require('cors');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const app = express();
app.use(cors());
app.use(express.json());

// ── Storage ──
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function getFilePath(resource) {
  const dir = path.join(DATA_DIR, resource);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}

function uid() {
  return Date.now().toString(36) + crypto.randomBytes(6).toString('hex');
}

function nowISO() {
  return new Date().toISOString();
}

function readAll(resource) {
  const dir = getFilePath(resource);
  const items = [];
  try {
    const files = fs.readdirSync(dir);
    for (const file of files) {
      if (!file.endsWith('.json')) continue;
      try {
        const data = JSON.parse(fs.readFileSync(path.join(dir, file), 'utf8'));
        items.push(data);
      } catch {}
    }
  } catch {}
  items.sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
  return items;
}

function getOne(resource, id) {
  const file = path.join(getFilePath(resource), id + '.json');
  try {
    return JSON.parse(fs.readFileSync(file, 'utf8'));
  } catch {
    return null;
  }
}

function saveOne(resource, id, data) {
  const iso = nowISO();
  const existing = getOne(resource, id);
  const item = existing 
    ? { ...existing, ...data, updated_at: iso }
    : { id, ...data, created_at: iso, updated_at: iso };
  const file = path.join(getFilePath(resource), id + '.json');
  fs.writeFileSync(file, JSON.stringify(item, null, 2));
  return item;
}

function deleteOne(resource, id) {
  const file = path.join(getFilePath(resource), id + '.json');
  try { fs.unlinkSync(file); return true; } catch { return false; }
}

function verifyToken(req) {
  const auth = (req.headers.authorization || '').replace('Bearer ', '');
  if (!auth) return null;
  try {
    return JSON.parse(Buffer.from(auth, 'base64').toString());
  } catch { return null; }
}

function findUserByEmail(email) {
  const all = readAll('users');
  return all.find(u => u.email === email);
}

// ── Ensure default admin ──
const admin = findUserByEmail('puguh.legowo.k@gmail.com');
if (!admin) {
  saveOne('users', uid(), { name: 'Admin', email: 'puguh.legowo.k@gmail.com', password: 'Admin123!', role: 'admin' });
  console.log('Default admin created');
}

// ── Auth ──
app.post('/api/auth', (req, res) => {
  const { action, email, password, token } = req.body;
  
  if (action === 'login') {
    const user = findUserByEmail(email);
    if (!user || user.password !== password) return res.status(401).json({ error: 'Email/password salah' });
    const tok = Buffer.from(JSON.stringify({ id: user.id, role: user.role, name: user.name })).toString('base64');
    return res.json({ token: tok, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
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

app.post('/api/login', (req, res) => {
  const { email, password } = req.body;
  const user = findUserByEmail(email);
  if (!user || user.password !== password) return res.status(401).json({ error: 'Email/password salah' });
  const tok = Buffer.from(JSON.stringify({ id: user.id, role: user.role, name: user.name })).toString('base64');
  return res.json({ token: tok, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
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

// ── CRUD ──
const resources = ['clients', 'appointments', 'services', 'staff', 'products', 'transactions', 'reports', 'users', 'homecare', 'customer_types'];

resources.forEach(resource => {
  const base = '/api/' + resource;
  
  app.get(base, (req, res) => res.json(readAll(resource)));
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
    res.status(201).json(saveOne(resource, uid(), req.body));
  });
  app.put(base + '/:id', (req, res) => {
    if (resource === 'users') {
      const tok = verifyToken(req);
      if (!tok) return res.status(401).json({ error: 'Unauthorized' });
      const user = getOne('users', tok.id);
      if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    if (!getOne(resource, req.params.id)) return res.status(404).json({ error: 'Not found' });
    res.json(saveOne(resource, req.params.id, req.body));
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


// ── Settings ──
const SETTINGS_FILE = path.join(DATA_DIR, 'settings.json');
const defaultSettings = {
  spa_name: 'Lelap Mom Baby Care Salatiga',
  address: 'Jl Taman Pahlawan Salatiga',
  tagline: 'Perawatan Profesional dan Hangat untuk Kesehatan Mama dan Buah Hati',
  whatsapp: '',
  open_time: '08:00',
  close_time: '20:00'
};

function getSettings() {
  try { return { ...defaultSettings, ...JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8')) }; }
  catch { return { ...defaultSettings }; }
}

app.get('/api/settings', (req, res) => res.json(getSettings()));

app.put('/api/settings', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  const user = getOne('users', tok.id);
  if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner' });
  const current = getSettings();
  const updated = { ...current, ...req.body };
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(updated, null, 2));
  res.json(updated);
});

// ── Change Password ──
app.post('/api/change-password', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  const { old_password, new_password } = req.body;
  if (!old_password || !new_password) return res.status(400).json({ error: 'Old and new password required' });
  if (new_password.length < 6) return res.status(400).json({ error: 'Password minimal 6 karakter' });
  const user = getOne('users', tok.id);
  if (!user || user.password !== old_password) return res.status(400).json({ error: 'Password lama salah' });
  saveOne('users', tok.id, { password: new_password });
  res.json({ success: true });
});

// ── Backup ──
app.get('/api/backup/download', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  const user = getOne('users', tok.id);
  if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner' });
  
  const backup = {};
  ['clients','appointments','services','staff','products','transactions','reports','users','homecare','customer_types'].forEach(r => {
    backup[r] = readAll(r);
  });
  backup.settings = getSettings();
  backup.exported_at = nowISO();
  
  res.setHeader('Content-Type', 'application/json');
  res.setHeader('Content-Disposition', 'attachment; filename="lelapsapadash-backup-' + new Date().toISOString().split('T')[0] + '.json"');
  res.json(backup);
});

app.post('/api/backup/restore', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  const user = getOne('users', tok.id);
  if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner' });
  
  const backup = req.body;
  if (!backup || !backup.exported_at) return res.status(400).json({ error: 'Invalid backup file' });
  
  let count = 0;
  ['clients','appointments','services','staff','products','transactions','reports','users','homecare','customer_types'].forEach(r => {
    if (Array.isArray(backup[r])) {
      // Clear existing
      try {
        const dir = getDir(r);
        fs.readdirSync(dir).forEach(f => { if (f.endsWith('.json')) fs.unlinkSync(path.join(dir, f)); });
      } catch {}
      // Restore
      backup[r].forEach(item => {
        if (item.id) saveOne(r, item.id, item);
        else saveOne(r, uid(), item);
        count++;
      });
    }
  });
  if (backup.settings) fs.writeFileSync(SETTINGS_FILE, JSON.stringify(backup.settings, null, 2));
  res.json({ success: true, restored: count });
});

// ── Reset Data ──
app.post('/api/reset-data', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized' });
  const user = getOne('users', tok.id);
  if (!user || user.role !== 'admin') return res.status(403).json({ error: 'Only Owner' });
  const { confirm, scope } = req.body;
  if (confirm !== 'YA SAYA YAKIN') return res.status(400).json({ error: 'Ketik "YA SAYA YAKIN" untuk konfirmasi' });
  
  const scopes = (scope === 'all') 
    ? ['clients','appointments','services','staff','products','transactions','reports','homecare','customer_types','users']
    : ['appointments','transactions'];
  
  let deleted = 0;
  scopes.forEach(r => {
    try {
      const dir = getDir(r);
      fs.readdirSync(dir).forEach(f => {
        if (f.endsWith('.json')) { fs.unlinkSync(path.join(dir, f)); deleted++; }
      });
    } catch {}
  });
  
  // Re-create admin if users were deleted
  if (scopes.includes('users')) {
    if (!findUserByEmail('puguh.legowo.k@gmail.com')) {
      saveOne('users', uid(), { name: 'Admin', email: 'puguh.legowo.k@gmail.com', password: 'Admin123!', role: 'admin' });
    }
  }
  
  res.json({ success: true, deleted });
});


// ── Static ──
app.use(express.static(path.join(__dirname, 'public')));

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Start ──
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`SmartSpaDash running on port ${PORT}`);
});
