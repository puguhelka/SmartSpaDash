// SmartSpaDash Server — Express + JSON File Storage
// Lelap Mom Baby Care Salatiga

const express = require('express');
const cors = require('cors');
const helmet = require('helmet');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');
const multer = require('multer');
require('dotenv').config();

const app = express();
app.use(helmet({
  contentSecurityPolicy: {
    directives: {
      defaultSrc: ["'self'"],
      scriptSrc: ["'self'", "'unsafe-inline'"],       // dashboard needs inline scripts
      scriptSrcAttr: ["'unsafe-inline'"],               // dashboard needs onclick handlers
      styleSrc: ["'self'", "'unsafe-inline'", "https:"],
      imgSrc: ["'self'", "data:"],
      fontSrc: ["'self'", "https:", "data:"],
      connectSrc: ["'self'"],
      objectSrc: ["'none'"],
      frameAncestors: ["'self'"],
      formAction: ["'self'"],
      upgradeInsecureRequests: [],
    }
  }
}));
app.disable('x-powered-by');
app.use((req, res, next) => {
  res.setHeader('Permissions-Policy', 'camera=(), microphone=(), geolocation=self');
  next();
});
app.use(cors({
  origin: function(origin, callback) {
    // Allow all trycloudflare.com tunnels + Firebase + localhost
    if (!origin) return callback(null, true);
    if (origin.endsWith('.trycloudflare.com')) return callback(null, true);
    const allowed = [
      'https://lelap-booking-care.firebaseapp.com',
      'https://lelap-booking-care.web.app',
      'http://localhost:8081',
      'http://localhost:3000',
    ];
    if (allowed.indexOf(origin) !== -1) return callback(null, true);
    callback(null, false); // reject silently (no CORS headers)
  },
  methods: ['GET', 'POST', 'PUT', 'DELETE', 'OPTIONS'],
  allowedHeaders: ['Content-Type', 'Authorization']
}));
app.use(express.json({ limit: '10mb' }));

// RATE LIMITING — simple in-memory
const rateLimitMap = new Map();
const RATE_LIMIT_WINDOW = 60000; // 1 minute
const RATE_LIMIT_AUTH = 10;      // 10 req/min for auth
const RATE_LIMIT_GENERAL = 60;    // 60 req/min for others
app.use((req, res, next) => {
  const ip = req.headers['x-forwarded-for'] || req.socket.remoteAddress || 'unknown';
  const key = ip + ':' + (req.path.includes('/auth/') ? 'auth' : 'general');
  const limit = req.path.includes('/auth/') ? RATE_LIMIT_AUTH : RATE_LIMIT_GENERAL;
  const now = Date.now();
  if (!rateLimitMap.has(key)) rateLimitMap.set(key, []);
  const window = rateLimitMap.get(key).filter(t => now - t < RATE_LIMIT_WINDOW);
  if (window.length >= limit) {
    return res.status(429).json({ error: 'Too many requests', code: 429 });
  }
  window.push(now);
  rateLimitMap.set(key, window);
  next();
});

// REQUEST LOGGING
app.use(function(req, res, next){
  console.log(new Date().toISOString().replace('T',' ').substring(0,19)+' '+req.method+' '+req.url+' from '+req.headers['x-forwarded-for']||req.socket.remoteAddress);
  next();
});

// ── Storage ──
const DATA_DIR = process.env.DATA_DIR || path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR, { recursive: true });

function getFilePath(resource) {
  const dir = path.join(DATA_DIR, resource);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  return dir;
}
const getDir = getFilePath;

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

// ── Auto-seed master data from Google Sheets ──
const SHEET_ID = '1TwBM_zb-kfX3IVvf6CblmFnpsiqoTQLxhwb3ARcYelg';
const SHEET_URL = 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/export?format=csv';

async function seedFromSheet() {
  try {
    const https = require('https');
    const csvData = await new Promise((resolve, reject) => {
      https.get(SHEET_URL, (res) => {
        let data = '';
        res.on('data', chunk => data += chunk);
        res.on('end', () => resolve(data));
        res.on('error', reject);
      }).on('error', reject);
    });
    
    const lines = csvData.split('\n').filter(l => l.trim());
    if (lines.length < 2) return console.log('Sheet empty, skip seed');
    
    const header = lines[0].split(',');
    const staffMap = {};
    let servicesSeeded = 0;
    
    for (let i = 1; i < lines.length; i++) {
      const cols = lines[i].split(',');
      if (cols.length < 3 || !cols[2]) continue;
      
      const kode = (cols[0]||'').replace(/"/g,'').trim();
      const kategori = (cols[1]||'').replace(/"/g,'').trim();
      const nama = (cols[2]||'').replace(/"/g,'').trim();
      const durasi = parseInt(cols[3]) || 60;
      const harga = parseInt(cols[4]) || 0;
      const fee_sa = parseInt(cols[5]) || 0;
      const fee_terapis = parseInt(cols[6]) || 0;
      const nama_terapis = (cols[8]||'').replace(/"/g,'').trim();
      
      if (!nama) continue;
      
      // Only seed if not already exists
      const existing = readAll('services');
      const exists = existing.some(s => s.code === kode);
      if (!exists) {
        saveOne('services', kode.toLowerCase(), {
          code: kode, name: nama, category: kategori,
          duration: durasi, price: harga,
          fee_sa: fee_sa, fee_terapis: fee_terapis,
        });
        servicesSeeded++;
      }
      
      if (nama_terapis && nama_terapis !== 'Owner' && !staffMap[nama_terapis]) {
        staffMap[nama_terapis] = true;
        const staffExist = readAll('staff').some(s => s.name === nama_terapis);
        if (!staffExist) {
          saveOne('staff', nama_terapis.toLowerCase(), {
            name: nama_terapis, role: 'Terapis'
          });
        }
      }
    }
    
    // Ensure Owner exists
    const staffExist2 = readAll('staff').some(s => s.name === 'Owner');
    if (!staffExist2) {
      saveOne('staff', 'owner', { name: 'Owner', role: 'Owner' });
    }
    
    console.log('Sheet seed: ' + servicesSeeded + ' services, ' + Object.keys(staffMap).length + ' staff');
  } catch(e) {
    console.log('Sheet seed error (non-fatal): ' + e.message);
  }
}

// Run auto-seed after startup
setTimeout(seedFromSheet, 2000);

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
    if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    try {
      const decoded = JSON.parse(Buffer.from(tok, 'base64').toString());
      const user = getOne('users', decoded.id);
      if (!user) return res.status(401).json({ error: 'User not found' });
      return res.json({ user: { id: decoded.id, name: user.name, email: user.email, role: user.role } });
    } catch { return res.status(401).json({ error: 'Invalid token' }); }
  }
  
  return res.status(400).json({ error: 'Invalid action' });
});

app.post('/api/auth', (req, res) => {
  const { action, email, password } = req.body;
  if (action !== 'login') return res.status(400).json({ error: 'Invalid action' });
  
  const user = findUserByEmail(email);
  if (!user || user.password !== password) return res.status(401).json({ error: 'Email/password salah' });
  const tok = Buffer.from(JSON.stringify({ id: user.id, role: user.role, name: user.name })).toString('base64');
  return res.json({ token: tok, user: { id: user.id, name: user.name, email: user.email, role: user.role } });
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
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  try {
    const decoded = JSON.parse(Buffer.from(tok, 'base64').toString());
    const user = getOne('users', decoded.id);
    if (!user) return res.status(401).json({ error: 'User not found' });
    return res.json({ user: { id: decoded.id, name: user.name, email: user.email, role: user.role } });
  } catch { return res.status(401).json({ error: 'Invalid token' }); }
});

// ── Dashboard (protected) ──
app.get('/api/dashboard', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const apps = readAll('appointments');
  const services = readAll('services');
  const now = new Date();
  const today = now.toISOString().split('T')[0];
  const monthStart = now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0');

  // Unique clients from SELESAI appointments only
  const uniqueMap = {};
  const bulanIniMap = {};
  apps.forEach(a => {
    if (a.status !== 'Selesai') return;
    const key = (a.client_name || '') + '|' + (a.wa || '');
    if (key !== '|') {
      uniqueMap[key] = true;
      if (a.date && a.date.startsWith(monthStart)) {
        bulanIniMap[key] = true;
      }
    }
  });

  res.json({
    totalBookings: apps.filter(a => a.status === 'Selesai').length,
    bookingsBulanIni: apps.filter(a => a.status === 'Selesai' && a.date && a.date.startsWith(monthStart)).length,
    bookingsHariIni: apps.filter(a => a.status === 'Selesai' && a.date === today).length,
    draftBookings: apps.filter(a => (a.status === 'Pending' || a.status === 'Booking' || a.status === 'Menunggu') && a.date === today).length,
    totalServices: services.length,
    totalClients: Object.keys(uniqueMap).length,
    clientsBulanIni: Object.keys(bulanIniMap).length,
    recentAppointments: apps.slice(0, 5)
  });
});

// ── Approval helpers ──
const APPROVALS_FILE = path.join(DATA_DIR, 'approvals.json');
function readApprovals(){try{return JSON.parse(fs.readFileSync(APPROVALS_FILE,'utf8'))}catch(e){return[]}}
function saveApprovals(d){fs.writeFileSync(APPROVALS_FILE,JSON.stringify(d,null,2))}

// ── Override appointments DELETE: Selesai need Owner approval ──
app.delete('/api/appointments/:id', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  if (!user) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const appt = getOne('appointments', req.params.id);
  if (!appt) return res.status(404).json({ error: 'Not found' });
  const role = (user.role || '').toLowerCase();

  if (appt.status === 'Selesai' || appt.status === 'Lunas') {
    if (role === 'owner' || role === 'admin') {
      const kw = (req.body || {}).keyword || '';
      if (kw !== 'HAPUS PERMANEN') {
        return res.status(400).json({ error: 'Ketik "HAPUS PERMANEN" untuk hapus booking selesai', needKeyword: true });
      }
      deleteOne('appointments', req.params.id);
      return res.json({ success: true, message: 'Booking selesai dihapus permanen' });
    }
    // Non-owner → create approval request
    const approvals = readApprovals();
    approvals.push({
      id: uid(),
      appointment_id: req.params.id,
      booking_code: appt.booking_code || '',
      client_name: appt.client_name || '',
      service: appt.service || '',
      date: appt.date || '',
      requested_by: user.name || user.email,
      requested_by_id: tok.id,
      created_at: new Date().toISOString(),
      status: 'pending'
    });
    saveApprovals(approvals);
    return res.json({ success: true, message: 'Permintaan hapus dikirim ke Owner', pendingApproval: true });
  }

  // Non-Selesai — normal delete
  deleteOne('appointments', req.params.id);
  res.json({ success: true });
});

app.get('/api/approvals', (req, res) => {
  res.json(readApprovals().filter(function(a){return a.status==='pending'}));
});

app.post('/api/approvals/:id/approve', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  const role = (user.role || '').toLowerCase();
  if (role !== 'owner' && role !== 'admin') return res.status(403).json({ error: 'Only Owner' });
  const approvals = readApprovals();
  const approval = approvals.find(function(a){return a.id===req.params.id});
  if (!approval) return res.status(404).json({ error: 'Not found' });
  const kw = (req.body || {}).keyword || '';
  if (kw !== 'HAPUS PERMANEN') {
    return res.status(400).json({ error: 'Ketik "HAPUS PERMANEN" untuk konfirmasi', needKeyword: true });
  }
  approval.status = 'approved';
  approval.approved_by = user.name || user.email;
  approval.approved_at = new Date().toISOString();
  saveApprovals(approvals);
  deleteOne('appointments', approval.appointment_id);
  res.json({ success: true, message: 'Booking dihapus permanen' });
});

// ── LAPORAN: Komisi Terapis & Fee SA (protected) ──
app.get('/api/reports/commission', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const month = req.query.month;
  if (!month) return res.status(400).json({ error: 'month required (YYYY-MM)' });
  const apps = readAll('appointments').filter(a => a.date && a.date.startsWith(month) && (a.status === 'Selesai' || a.status === 'Lunas'));
  const svcs = readAll('services');
  const feeMap = {}; svcs.forEach(s => { feeMap[s.name] = { fee_terapis: s.fee_terapis || 0, fee_sa: s.fee_sa || 0 }; });
  const therapistMap = {}, details = [];
  apps.forEach(a => {
    const therapist = a.staff || a.therapist || 'Tidak diketahui';
    const svc = feeMap[a.service] || { fee_terapis: 0 };
    let feePerBooking = Number(svc.fee_terapis) || 0;
    if (a.session_total > 1) feePerBooking = Math.round(feePerBooking / a.session_total);
    const transport = Number(a.transport) || 0;
    if (!therapistMap[therapist]) therapistMap[therapist] = { name: therapist, bookings: 0, fee: 0, transport: 0 };
    therapistMap[therapist].bookings++;
    therapistMap[therapist].fee += feePerBooking;
    therapistMap[therapist].transport += transport;
    details.push({ date: a.date, booking_code: a.booking_code || '', client_name: a.client_name || '', service: a.service || '', therapist, fee_terapis: feePerBooking, transport, session: a.session_total ? `${a.session_index}/${a.session_total}` : null });
  });
  const therapists = Object.values(therapistMap).sort((a, b) => b.fee - a.fee);
  const totalFee = therapists.reduce((s, t) => s + t.fee, 0);
  const totalTransport = therapists.reduce((s, t) => s + t.transport, 0);
  res.json({ month, total_bookings: apps.length, total_fee: totalFee, total_transport: totalTransport, total_commission: totalFee + totalTransport, therapists, details: details.sort((a, b) => a.date.localeCompare(b.date)) });
});

app.get('/api/reports/fee-sa', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const month = req.query.month;
  if (!month) return res.status(400).json({ error: 'month required (YYYY-MM)' });
  const apps = readAll('appointments').filter(a => a.date && a.date.startsWith(month) && (a.status === 'Selesai' || a.status === 'Lunas'));
  const svcs = readAll('services');
  const feeMap = {}; svcs.forEach(s => { feeMap[s.name] = { fee_sa: s.fee_sa || 0 }; });
  let totalFee = 0;
  const details = apps.map(a => {
    const svc = feeMap[a.service] || { fee_sa: 0 };
    let feePerBooking = Number(svc.fee_sa) || 0;
    if (a.session_total > 1) feePerBooking = Math.round(feePerBooking / a.session_total);
    totalFee += feePerBooking;
    return { date: a.date, booking_code: a.booking_code || '', client_name: a.client_name || '', service: a.service || '', fee_sa: feePerBooking, session: a.session_total ? `${a.session_index}/${a.session_total}` : null };
  });
  res.json({ month, total_bookings: apps.length, total_fee_sa: totalFee, details: details.sort((a, b) => a.date.localeCompare(b.date)) });
});

// ── CRUD ──
const resources = ['clients', 'appointments', 'services', 'staff', 'products', 'transactions', 'reports', 'users', 'homecare', 'customer_types'];

// Client search (before CRUD to avoid route collision) — protected
app.get('/api/clients/search', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const q = (req.query.q || '').toLowerCase().trim();
  if (!q || q.length < 2) return res.json([]);
  const clients = readAll('clients');
  const apps = readAll('appointments');
  const results = clients
    .filter(c => {
      const name = (c.name || '').toLowerCase();
      const phone = (c.phone || c.wa || '').toLowerCase().replace(/[^0-9]/g, '');
      const qClean = q.replace(/[^0-9a-z]/g, '');
      return name.includes(qClean) || phone.includes(qClean);
    })
    .slice(0, 15)
    .map(c => {
      const clientApps = apps.filter(a => a.client_id === c.id || (a.client_name || '').toLowerCase() === (c.name || '').toLowerCase());
      const selesai = clientApps.filter(a => a.status === 'Selesai' || a.status === 'Lunas');
      const total_spending = c.total_spending || 0;
      const tierPts = Math.floor(total_spending / 10000);
      let tier = 'non-tier';
      if (tierPts >= 130) tier = 'platinum';
      else if (tierPts >= 100) tier = 'gold';
      else if (tierPts >= 70) tier = 'silver';
      else if (tierPts >= 30) tier = 'bronze';
      return {
        id: c.id,
        name: c.name || '',
        phone: c.phone || c.wa || '',
        address: c.address || '',
        total_bookings: selesai.length,
        tier,
        tier_label: tier.charAt(0).toUpperCase() + tier.slice(1),
        points: c.loyalty_points || 0,
        total_spending
      };
    });
  res.json(results);
});

resources.forEach(resource => {
  const base = '/api/' + resource;
  
  // Auth guard: protect sensitive resources, keep services/staff public
  const requiresAuth = ['clients', 'appointments', 'products', 'transactions', 'reports', 'users', 'homecare', 'customer_types'];
  const isPublic = ['services', 'staff']; // always public
  
  function guard(req, res, next) {
    if (isPublic.includes(resource) && req.method === 'GET') return next();
    // Allow GET on clients with PII stripping (handled in route)
    if (resource === 'clients' && req.method === 'GET') return next();
    if (!requiresAuth.includes(resource) && !['POST','PUT','DELETE'].includes(req.method)) return next();
    const tok = verifyToken(req);
    if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const user = getOne('users', tok.id);
    if (!user) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    req.user = user;
    next();
  }
  
  app.get(base, guard, (req, res) => {
    let data = readAll(resource);
    // PII stripping: when unauthenticated, strip sensitive fields from clients
    if (resource === 'clients' && !req.user) {
      data = data.map(c => ({
        id: c.id,
        name: c.name || '',
        loyalty_points: c.loyalty_points || 0,
        total_spending: c.total_spending || 0,
        created_at: c.created_at
      }));
    }
    res.json(data);
  });
  app.get(base + '/:id', guard, (req, res) => {
    const item = getOne(resource, req.params.id);
    if (!item) return res.status(404).json({ error: 'Not found' });
    // PII stripping for single client without auth
    if (resource === 'clients' && !req.user) {
      const { id, name, loyalty_points, total_spending, created_at } = item;
      return res.json({ id, name: name || '', loyalty_points: loyalty_points || 0, total_spending: total_spending || 0, created_at });
    }
    res.json(item);
  });
  app.post(base, guard, (req, res) => {
    if (resource === 'users') {
      const tok = verifyToken(req);
      if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
      const user = getOne('users', tok.id);
      if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    // Input validation for appointments
    if (resource === 'appointments') {
      const errors = [];
      const svc = (req.body.service || '').trim();
      const dt = (req.body.date || '').trim();
      const tm = (req.body.time || '').trim();
      if (!svc) errors.push('service wajib diisi');
      if (!dt) errors.push('date wajib diisi');
      else if (!/^\d{4}-\d{2}-\d{2}$/.test(dt)) errors.push('date format YYYY-MM-DD');
      else {
        const today = new Date().toISOString().split('T')[0];
        if (dt < today) errors.push('date tidak boleh di masa lalu');
      }
      if (!tm) errors.push('time wajib diisi');
      else if (!/^\d{2}:\d{2}$/.test(tm)) errors.push('time format HH:MM');
      if (req.body.type && !['Inhouse','Homecare'].includes(req.body.type)) errors.push('type harus Inhouse atau Homecare');
      if (errors.length > 0) return res.status(400).json({ error: 'Validasi gagal', details: errors });
    }
    res.status(201).json(saveOne(resource, uid(), req.body));
  });
  app.put(base + '/:id', guard, (req, res) => {
    if (resource === 'users') {
      const user = req.user || getOne('users', (verifyToken(req) || {}).id);
      if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    const old = getOne(resource, req.params.id);
    if (!old) return res.status(404).json({ error: 'Not found' });
    
    // === AUTO LOYALTY POINTS: when appointment status → Selesai ===
    let pointsAwarded = null;
    if (resource === 'appointments' && req.body.status === 'Selesai' && old.status !== 'Selesai' && !old.completed_at) {
      // Require payment_method when completing
      if (!req.body.payment_method) {
        return res.status(400).json({ error: 'payment_method wajib diisi (cash/qris/transfer/dp)' });
      }
      req.body.completed_at = new Date().toISOString();
      const updated = saveOne(resource, req.params.id, { ...req.body, payment: req.body.payment_method });
      if (updated.client_id) {
        pointsAwarded = awardBookingPoints(updated, updated.client_id);
      }
      return res.json({ ...updated, points: pointsAwarded });
    }
    
    // Auto-confirm: assign therapist to Menunggu booking
    if (resource === 'appointments' && old.status === 'Menunggu' && (req.body.staff || req.body.therapist)) {
      req.body.status = 'Confirmed';
      req.body.confirmed_at = new Date().toISOString();
    }
    
    res.json(saveOne(resource, req.params.id, req.body));
  });
  app.delete(base + '/:id', guard, (req, res) => {
    if (resource === 'users') {
      const user = req.user || getOne('users', (verifyToken(req) || {}).id);
      if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner can manage users' });
    }
    deleteOne(resource, req.params.id);
    res.json({ success: true });
  });
});

// ── Slot Availability (public API mirror) ──
app.get('/api/slots', (req, res) => {
  const date = req.query.date || new Date().toISOString().split('T')[0];
  const serviceId = req.query.service;
  const therapist = req.query.therapist || null;
  
  if (!serviceId) {
    return res.status(400).json({ error: 'service parameter required (service code or ID)' });
  }
  
  // Find service by code or ID
  const services = readAll('services');
  const service = services.find(s => s.code === serviceId || s.id === serviceId);
  if (!service) {
    return res.json({ date, service_id: serviceId, slots: [], error: 'Service not found' });
  }
  
  const allAppointments = readAll('appointments');
  const settings = getSettings();
  const slots = calculateSlots(date, service, therapist, allAppointments, settings);
  
  res.json({ date, service_id: serviceId, service_name: service.name, duration: service.duration, price: service.price, slots });
});

// ── API Aliases (frontend compatibility) ──
const aliases = { 'layanan': 'services', 'booking': 'appointments', 'pelanggan': 'clients' };
Object.entries(aliases).forEach(([alias, resource]) => {
  const isPublicAlias = resource === 'services';
  
  function aliasGuard(req, res, next) {
    if (isPublicAlias && req.method === 'GET') return next();
    const tok = verifyToken(req);
    if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const user = getOne('users', tok.id);
    if (!user) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    req.user = user;
    next();
  }
  
  app.get('/api/' + alias, aliasGuard, (req, res) => res.json(readAll(resource)));
  app.get('/api/' + alias + '/:id', aliasGuard, (req, res) => {
    const item = getOne(resource, req.params.id);
    if (!item) return res.status(404).json({ error: 'Not found' });
    res.json(item);
  });
  app.post('/api/' + alias, aliasGuard, (req, res) => res.status(201).json(saveOne(resource, uid(), req.body)));
  app.put('/api/' + alias + '/:id', aliasGuard, (req, res) => {
    const existing = getOne(resource, req.params.id);
    if (!existing) return res.status(404).json({ error: 'Not found' });
    
    // Require payment_method when completing
    if (resource === 'appointments' && req.body.status === 'Selesai' && existing.status !== 'Selesai') {
      if (!req.body.payment_method) {
        return res.status(400).json({ error: 'payment_method wajib diisi (cash/qris/transfer/dp)' });
      }
      req.body.payment = req.body.payment_method;
      req.body.completed_at = new Date().toISOString();
    }
    
    // Auto-confirm: assign therapist to Menunggu booking
    if (resource === 'appointments' && existing.status === 'Menunggu' && (req.body.staff || req.body.therapist)) {
      req.body.status = 'Confirmed';
      req.body.confirmed_at = new Date().toISOString();
    }
    
    const updated = saveOne(resource, req.params.id, req.body);
    
    // === Auto-earn points when booking marked Selesai via dashboard ===
    if (resource === 'appointments' && req.body.status === 'Selesai' && existing.status !== 'Selesai') {
      try {
        let clientId = updated.client_id;
        // Find client by name+WA if no client_id
        if (!clientId) {
          const clientName = (updated.client_name || '').trim();
          const clientWa = (updated.wa || updated.phone || '').trim();
          if (clientName) {
            const allClients = readAll('clients');
            const match = allClients.find(c => 
              (c.name || '').trim().toLowerCase() === clientName.toLowerCase() &&
              (c.phone || c.wa || '').trim() === clientWa
            );
            if (match) {
              clientId = match.id;
            } else {
              // Create new client record
              clientId = uid();
              saveOne('clients', clientId, {
                id: clientId, name: clientName, phone: clientWa,
                loyalty_points: 0, total_spending: 0,
                created_at: new Date().toISOString()
              });
            }
          }
        }
        if (clientId) {
          const baseAmount = (updated.amount || 0) + (updated.discount || 0);
          const c = getOne('clients', clientId);
          const oldSpending = (c ? c.total_spending : 0) || 0;
          const newSpending = oldSpending + baseAmount;
          const oldPts = Math.floor(oldSpending / 10000);
          const newPts = Math.floor(newSpending / 10000);
          const pts = Math.max(0, newPts - oldPts);
          
          // Update client total_spending
          saveOne('clients', clientId, { total_spending: newSpending });
          
          if (pts > 0) {
            earnPoints(clientId, pts, 'booking',
              `Akumulasi: +Rp${baseAmount.toLocaleString('id-ID')} → total Rp${newSpending.toLocaleString('id-ID')} (${newPts} pts)`,
              updated.id);
          }
        }
      } catch (e) { console.error('Points award failed:', e.message); }
    }
    
    res.json(updated);
  });
  app.delete('/api/' + alias + '/:id', aliasGuard, (req, res) => {
    deleteOne(resource, req.params.id);
    res.json({ success: true });
  });
});


// ── Import CSV ──
app.get('/api/import-template', (req, res) => {
  const header = 'date,type,time,therapist,service,client_name,wa,child,age,address,discount,transport,deposit,payment,notes,client_type,status';
  const example = '2026-06-05,Inhouse,09:00,Salsa,B001 - Baby Relaxation Massage,Nama Klien,08123456789,Nama Anak,6 bln,Jl. Contoh No 123,0,0,0,Cash,,Anak,Menunggu';
  res.setHeader('Content-Type', 'text/csv; charset=utf-8');
  res.setHeader('Content-Disposition', 'attachment; filename="template-import-booking.csv"');
  res.send('\uFEFF' + header + '\n' + example);
});

app.post('/api/import-csv', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  const role = (user.role || '').toLowerCase();
  if (role !== 'owner' && role !== 'admin') return res.status(403).json({ error: 'Only Owner' });

  const { data: csvText } = req.body || {};
  if (!csvText) return res.status(400).json({ error: 'No CSV data' });

  const lines = csvText.split('\n').filter(function(l){return l.trim()});
  if (lines.length < 2) return res.status(400).json({ error: 'CSV kosong atau hanya header' });

  const header = lines[0].split(',').map(function(h){return h.trim()});
  const imported = [];
  const skipped = [];

  for (var i = 1; i < lines.length; i++) {
    var vals = lines[i].split(',').map(function(v){return v.trim()});
    if (vals.length < 2) continue;
    var row = {};
    header.forEach(function(h, j){row[h] = vals[j] || ''});

    if (!row.client_name && !row.service) { skipped.push(i); continue; }

    // Parse service code if present
    var svcName = row.service || '';
    var svcCode = '';
    var match = svcName.match(/^([A-Z0-9]+)\s*-\s*/);
    if (match) { svcCode = match[1]; svcName = svcName.substring(match[0].length); }

    var appointment = {
      id: uid(),
      date: row.date || new Date().toISOString().split('T')[0],
      type: row.type || 'Inhouse',
      time: row.time || '',
      staff: row.therapist || '',
      service: svcName,
      client_name: row.client_name || '',
      wa: row.wa || '',
      child: row.child || '',
      age: row.age || '',
      address: row.address || '',
      discount: parseInt(row.discount) || 0,
      transport: parseInt(row.transport) || 0,
      deposit: parseInt(row.deposit) || 0,
      payment: row.payment || 'Cash',
      notes: row.notes || '',
      client_type: row.client_type || 'Anak',
      status: row.status || 'Menunggu',
      amount: 0,
      booking_code: 'MBS-' + new Date().toISOString().split('T')[0].replace(/-/g,'').substring(2) + '-' + String(Date.now() % 1000 + i).padStart(3, '0'),
      created_at: new Date().toISOString(),
      updated_at: new Date().toISOString()
    };

    // Look up service price
    var svcs = readAll('services');
    var svc = svcs.find(function(s){return s.name === svcName});
    if (svc) {
      appointment.amount = svc.price || 0;
      appointment.duration = svc.duration || 60;
    }

    saveOne('appointments', appointment.id, appointment);
    imported.push(appointment.booking_code);
  }

  res.json({ success: true, imported: imported.length, skipped: skipped.length, codes: imported.slice(0, 10) });
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
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner' });
  const current = getSettings();
  const updated = { ...current, ...req.body };
  fs.writeFileSync(SETTINGS_FILE, JSON.stringify(updated, null, 2));
  res.json(updated);
});

// ── Change Password ──
app.post('/api/change-password', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
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
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner' });
  
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
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const user = getOne('users', tok.id);
  if (!user || (user.role||'').toLowerCase() !== 'admin' && (user.role||'').toLowerCase() !== 'owner') return res.status(403).json({ error: 'Only Owner' });
  
  const backup = req.body;
  if (!backup || !backup.exported_at) return res.status(400).json({ error: 'Invalid backup file' });
  
  let count = 0;
  ['clients','appointments','services','staff','products','transactions','reports','users','homecare','customer_types'].forEach(r => {
    if (Array.isArray(backup[r])) {
      // Clear existing
      try {
        const dir = getFilePath(r);
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

// ── Static Files (Dashboard SPA) ──
app.use(express.static(path.join(__dirname, 'public')));
app.use('/data/ig_images', express.static(path.join(__dirname, 'data', 'ig_images')));
app.use('/data/service_images', express.static(path.join(__dirname, 'data', 'service_images')));

// ── Public API (Client App) ── MUST be before catch-all


// ── JWT Helpers (Simple HMAC) ──
function createJWT(payload) {
  const header = Buffer.from(JSON.stringify({ alg: 'HS256', typ: 'JWT' })).toString('base64url');
  const body = Buffer.from(JSON.stringify(payload)).toString('base64url');
  const secret = get_settings_data().jwt_secret || 'lelap-secret-key-2024';
  const hmac = crypto.createHmac('sha256', secret).update(header + '.' + body).digest('base64url');
  return header + '.' + body + '.' + hmac;
}
function verifyJWT(token) {
  try {
    const [header, body, signature] = token.split('.');
    const secret = get_settings_data().jwt_secret || 'lelap-secret-key-2024';
    const expected = crypto.createHmac('sha256', secret).update(header + '.' + body).digest('base64url');
    if (signature !== expected) return null;
    return JSON.parse(Buffer.from(body, 'base64url').toString());
  } catch { return null; }
}
function get_settings_data() {
  try { return JSON.parse(fs.readFileSync(SETTINGS_FILE, 'utf8')); } catch { return {}; }
}

// Quick test route

// DEBUG: simulate services response
app.get('/api/public/debug-services', (req, res) => {
  let services = readAll('services');
  const { category, search } = req.query;
  if (category && category !== 'ALL') {
    services = services.filter(s => (s.category || '').toUpperCase() === category.toUpperCase());
  }
  if (search) {
    const q = search.toLowerCase();
    services = services.filter(s => (s.name || '').toLowerCase().includes(q));
  }
  const categories = [...new Set(services.map(s => s.category || 'OTHER'))];
  res.json({ services, categories: ['ALL', ...categories] });
});

app.get('/api/public/ping', (req, res) => res.json({ ping: 'pong', time: new Date().toISOString() }));

// ═══ Public API (fully inlined) ═══
// ── Public API Module ──
// Lelap Booking Care — Client App Backend
// Dibangun oleh Hermes Agent — 6 Juni 2026

// (already declared)
// (already declared)
// (already declared)

// Pake shared helpers dari server.js (readAll, saveOne, getOne, deleteOne, uid, nowISO, getFilePath)
// Function ini akan di-pass dari server.js saat init

// helpers passed from server.js

// ── JWT ──
const JWT_SECRET = (process.env.JWT_SECRET || 'lelapsalatigasecret2026');
const JWT_EXPIRY = 24 * 60 * 60 * 1000; // 24 jam

function createJWT(payload) {
  const header = Buffer.from(JSON.stringify({ alg: 'HS256', typ: 'JWT' })).toString('base64url');
  const now = Math.floor(Date.now() / 1000);
  const body = Buffer.from(JSON.stringify({ ...payload, iat: now, exp: now + JWT_EXPIRY / 1000 })).toString('base64url');
  const signature = crypto.createHmac('sha256', JWT_SECRET).update(header + '.' + body).digest('base64url');
  return header + '.' + body + '.' + signature;
}

function verifyJWT(token) {
  try {
    const parts = token.split('.');
    if (parts.length !== 3) return null;
    const payload = JSON.parse(Buffer.from(parts[1], 'base64url').toString());
    if (payload.exp < Math.floor(Date.now() / 1000)) return null;
    const expectedSig = crypto.createHmac('sha256', JWT_SECRET).update(parts[0] + '.' + parts[1]).digest('base64url');
    if (parts[2] !== expectedSig) return null;
    return payload;
  } catch { return null; }
}

// ── Google Token Verify ──
async function verifyGoogleToken(idToken) {
  // Call Google's tokeninfo endpoint
  const https = require('https');
  return new Promise((resolve, reject) => {
    https.get(`https://oauth2.googleapis.com/tokeninfo?id_token=${encodeURIComponent(idToken)}`, (resp) => {
      let data = '';
      resp.on('data', chunk => data += chunk);
      resp.on('end', () => {
        try {
          const parsed = JSON.parse(data);
          if (parsed.error) resolve(null);
          else resolve(parsed);
        } catch { resolve(null); }
      });
    }).on('error', () => resolve(null));
  });
}

// ── Auth Middleware ──
function publicAuth(req, res, next) {
  const token = (req.headers.authorization || '').replace('Bearer ', '');
  if (!token) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const decoded = verifyJWT(token);
  if (!decoded) return res.status(401).json({ error: 'Invalid or expired token' });
  req.client = decoded;
  next();
}

// ── GPS Radius Validation ──
const LELAP_LAT = -7.3326;
const LELAP_LNG = 110.5069;

function haversineDistance(lat1, lng1, lat2, lng2) {
  const R = 6371; // Earth radius in km
  const dLat = (lat2 - lat1) * Math.PI / 180;
  const dLng = (lng2 - lng1) * Math.PI / 180;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * Math.PI / 180) * Math.cos(lat2 * Math.PI / 180) *
    Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
}

// ── Radius Validation ──
const SERVICE_CITIES = ['salatiga', 'kota salatiga'];
const SERVICE_DISTRICTS = [
  'sidorejo', 'sidomukti', 'argomulyo', 'tingkir', // Salatiga
  'ambarawa', 'banyubiru', 'bawen', 'bandungan', 'bergas', 'bringin', 'bancak',
  'jambu', 'pabelan', 'pringapus', 'sumowono', 'suruh', 'susukan', 'tengaran',
  'tuntang', 'ungaran barat', 'ungaran timur', 'kaliwungu', // Kab. Semarang
];

function validateLocation(city, district) {
  const c = (city || '').toLowerCase().trim();
  const d = (district || '').toLowerCase().trim();
  if (SERVICE_CITIES.some(sc => c.includes(sc))) return true;
  if (SERVICE_DISTRICTS.some(sd => d.includes(sd) || c.includes(sd))) return true;
  return false;
}

// ── Slot Calculator ──
function calculateSlots(date, service, therapist, allAppointments, settings) {
  const openTime = (settings && settings.open_time) || '08:00';
  const closeTime = (settings && settings.close_time) || '20:00';
  const duration = service.duration || 60;
  const slots = [];
  
  // Parse open/close
  const [oh, om] = openTime.split(':').map(Number);
  const [ch, cm] = closeTime.split(':').map(Number);
  const startMin = oh * 60 + om;
  const endMin = ch * 60 + cm;
  
  // Generate slots every 15 minutes
  for (let m = startMin; m + duration <= endMin; m += 15) {
    const h = Math.floor(m / 60);
    const min = m % 60;
    const timeStr = `${String(h).padStart(2, '0')}:${String(min).padStart(2, '0')}`;
    const endH = Math.floor((m + duration) / 60);
    const endM = (m + duration) % 60;
    const endStr = `${String(endH).padStart(2, '0')}:${String(endM).padStart(2, '0')}`;
    
    // Check if slot is available — max 3 therapists per slot
    let available = true;
    let reason = '';
    
    // Check existing appointments for this date+time+therapist
    const overlapping = allAppointments.filter(a => {
      if (a.date !== date) return false;
      if (a.status === 'cancelled' || a.status === 'Dibatalkan' || a.status === 'Ditolak') return false;
      if (a.status === 'Menunggu') return false;
      if (therapist && a.staff !== therapist && a.therapist !== therapist) return false;
      // Check time overlap
      const aStart = timeToMinutes(a.time);
      const aDuration = a.duration || 60;
      const aEnd = aStart + aDuration;
      const slotStart = m;
      const slotEnd = m + duration;
      return slotStart < aEnd && slotEnd > aStart;
    });
    
    const bookedCount = overlapping.length;
    const freeCount = 3 - bookedCount;
    
    if (therapist) {
      // Specific therapist mode: slot full if that therapist is booked
      if (bookedCount > 0) {
        available = false;
        reason = 'full';
      }
    } else {
      // Any therapist mode: slot full only if all 3 therapists are booked
      if (bookedCount >= 3) {
        available = false;
        reason = 'full';
      }
    }
    
    // Special rule: PRENATAL YOGA
    if (service.name && service.name.toUpperCase().includes('PRENATAL YOGA') || 
        (service.category && service.category.toUpperCase().includes('YOGA'))) {
      const dayOfWeek = new Date(date).getDay(); // 0=Sun, 1=Mon
      if (dayOfWeek !== 1) { // Not Monday
        available = false;
        reason = 'prenatal_yoga_monday_only';
      } else if (timeStr !== '16:00') {
        available = false;
        reason = 'prenatal_yoga_1600_only';
      } else if (therapist !== 'Owner' && therapist !== 'owner') {
        available = false;
        reason = 'prenatal_yoga_owner_only';
      } else {
        // Max 6 participants
        const existingCount = overlapping.filter(a => 
          (a.service || '').toUpperCase().includes('PRENATAL') || 
          (a.service || '').toUpperCase().includes('YOGA')
        ).length;
        if (existingCount >= 6) {
          available = false;
          reason = 'prenatal_yoga_full';
        }
      }
    }
    
    slots.push({
      time: timeStr,
      end_time: endStr,
      available,
      booked_count: bookedCount,
      free_count: freeCount,
      reason: reason || (available ? null : 'full')
    });
  }
  
  return slots;
}

// ── Slot Summary for AI Chat ──
function getSlotSummary() {
  const services = readAll('services');
  const allAppointments = readAll('appointments');
  const settings = getSettings();
  const today = new Date().toISOString().split('T')[0];
  const tomorrow = new Date(Date.now() + 86400000).toISOString().split('T')[0];
  
  const days = [today, tomorrow];
  const dayLabels = ['HARI INI', 'BESOK'];
  const dayNames = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  
  let summary = '';
  
  for (let di = 0; di < days.length; di++) {
    const date = days[di];
    const dow = new Date(date).getDay();
    const dayName = dayNames[dow];
    
    // Monday — most services closed, only Prenatal
    if (dow === 1) {
      summary += `\n📅 ${dayLabels[di]} (${dayName}, ${date}):\n`;
      summary += `   ⚠️ HARI LIBUR. Hanya PRENATAL YOGA pukul 16:00-17:00 yang beroperasi.\n`;
      continue;
    }
    
    const activeServices = services.filter(s => s.active !== false && (s.duration || s.duration_minutes));
    if (activeServices.length === 0) {
      summary += `\n📅 ${dayLabels[di]} (${dayName}, ${date}): Tidak ada layanan tersedia.\n`;
      continue;
    }
    
    summary += `\n📅 ${dayLabels[di]} (${dayName}, ${date}):\n`;
    
    // Check overall slot availability across all services
    const allSlots = {};
    for (const svc of activeServices) {
      const slots = calculateSlots(date, svc, null, allAppointments, settings);
      for (const s of slots) {
        if (!allSlots[s.time]) {
          allSlots[s.time] = { available: 0, total: 0 };
        }
        allSlots[s.time].total++;
        if (s.available) allSlots[s.time].available++;
      }
    }
    
    // Summarize hourly
    const hours = {};
    for (const [time, data] of Object.entries(allSlots)) {
      const hour = time.split(':')[0] + ':00';
      if (!hours[hour]) hours[hour] = { available: 0, total: 0 };
      hours[hour].available += data.available;
      hours[hour].total += data.total;
    }
    
    const hourKeys = Object.keys(hours).sort();
    let line = '   ';
    for (const h of hourKeys) {
      const { available, total } = hours[h];
      if (available === 0) line += `❌${h} | `;
      else if (available === total) line += `✅${h} | `;
      else line += `⚠️${h}(${available}/${total}) | `;
    }
    summary += line.trimEnd() + '\n';
  }
  
  summary += '\n✅ = semua layanan tersedia | ⚠️ = sebagian tersedia | ❌ = penuh';
  return summary;
}

// Slot-related keywords
const SLOT_KEYWORDS = [
  'slot', 'jadwal', 'kosong', 'penuh', 'tersedia', 'booking',
  'jam', 'pagi', 'siang', 'sore', 'malam',
  'kapan', 'hari ini', 'besok', 'minggu',
];

function hasSlotIntent(text) {
  const lower = text.toLowerCase();
  // Exclude FAQ patterns that happen to match slot keywords
  const faqPatterns = ['jam buka', 'jam tutup', 'jam operasional', 'buka jam'];
  if (faqPatterns.some(p => lower.includes(p))) return false;
  
  const matchCount = SLOT_KEYWORDS.filter(kw => lower.includes(kw)).length;
  // Need at least 2 slot keywords to avoid false positives
  return matchCount >= 2;
}

function timeToMinutes(timeStr) {
  const [h, m] = (timeStr || '00:00').split(':').map(Number);
  return h * 60 + m;
}

// ═══════════════════════════════════════════════════════════
// MEMBERSHIP & LOYALTY MODULE — Lelap Mom Baby Care
// ═══════════════════════════════════════════════════════════

// ═══════════════════════════════════════════════════════════
// MEMBERSHIP SCHEME — loaded from data/membership-scheme.json
// ═══════════════════════════════════════════════════════════
function loadMembershipScheme() {
  try {
    const raw = fs.readFileSync(path.join(DATA_DIR, 'membership-scheme.json'), 'utf8');
    return JSON.parse(raw);
  } catch (e) {
    console.error('Failed to load membership-scheme.json:', e.message);
    return null;
  }
}

function getTierConfig(tier) {
  const scheme = loadMembershipScheme();
  if (!scheme) return { label: 'Non-Tier', color: '#999', voucher_label: 'Belum ada voucher' };
  if (tier === 'non-tier') return scheme.non_tier;
  const t = scheme.tiers.find(t => t.tier === tier);
  return t || scheme.non_tier;
}

function calculateTier(spendingPoints) {
  const scheme = loadMembershipScheme();
  if (!scheme || !scheme.tiers) return 'non-tier';
  // Tiers sorted by min_points descending
  const sorted = [...scheme.tiers].sort((a, b) => b.min_points - a.min_points);
  for (const t of sorted) {
    if (spendingPoints >= t.min_points) return t.tier;
  }
  return 'non-tier';
}

function getNextTier(tier) {
  const scheme = loadMembershipScheme();
  if (!scheme || !scheme.tiers) return null;
  const order = ['non-tier', ...scheme.tiers.map(t => t.tier)];
  const idx = order.indexOf(tier);
  return idx < order.length - 1 ? order[idx + 1] : null;
}

function hasUsedVoucherThisMonth(clientId) {
  const now = new Date();
  const thisMonth = now.getMonth();
  const thisYear = now.getFullYear();
  const bookings = readAll('appointments');
  return bookings.some(b => 
    b.client_id === clientId && 
    b.voucher_type && b.voucher_type !== '' && b.voucher_type !== 'membership' &&
    b.date && new Date(b.date + 'T00:00:00+07:00').getMonth() === thisMonth &&
    new Date(b.date + 'T00:00:00+07:00').getFullYear() === thisYear
  );
}

function hasUsedWelcome(clientId) {
  const bookings = readAll('appointments');
  return bookings.some(b => b.client_id === clientId && b.voucher_type === 'welcome');
}

function hasUsedBirthdayThisYear(clientId) {
  const thisYear = new Date().getFullYear();
  const bookings = readAll('appointments');
  return bookings.some(b => 
    b.client_id === clientId && b.voucher_type === 'birthday' &&
    b.date && new Date(b.date + 'T00:00:00+07:00').getFullYear() === thisYear
  );
}

// Points transactions stored in data/points_tx/{id}.json
// Each: { id, client_id, type: 'earn'|'redeem'|'expire'|'welcome', 
//         amount: +/-, source: 'booking'|'review'|'birthday'|'redeem'|'welcome'|'expiry',
//         description, booking_id, created_at, expires_at }

function getMembership(clientId) {
  const client = getOne('clients', clientId);
  if (!client) return { tier: 'non-tier', tier_label: 'Non-Tier', tier_color: '#999', 
    voucher_label: 'Belum ada voucher', points_balance: 0, spending_points: 0, 
    next_tier: 'bronze', next_tier_label: 'Bronze', points_to_next_tier: 30, 
    voucher_used_this_month: false, available_loyalty_vouchers: [],
    welcome_bonus_available: false, birthday_bonus_available: false };
  
  // Lifetime tier from total_spending (never decreases)
  const lifetimePts = Math.floor((client.total_spending || 0) / 10000);
  const tier = calculateTier(lifetimePts);
  const tierCfg = getTierConfig(tier);
  const nextTier = getNextTier(tier);
  const nextTierCfg = nextTier ? getTierConfig(nextTier) : null;
  
  // Spendable balance
  const spendable = client.loyalty_points || 0;
  
  // Check if any voucher already used this calendar month
  const voucherUsedThisMonth = hasUsedVoucherThisMonth(clientId);
  
  // Build available loyalty vouchers (tiers user can afford with spendable points)
  const availableTiers = [];
  const tierOrder = ['bronze', 'silver', 'gold', 'platinum'];
  for (const t of tierOrder) {
    const cfg = getTierConfig(t);
    if (spendable >= cfg.voucher_cost && !voucherUsedThisMonth) {
      availableTiers.push({
        tier: t,
        label: cfg.label,
        discount_percent: cfg.voucher_pct,
        cost: cfg.voucher_cost,
      });
    }
  }
  
  return {
    tier, tier_label: tierCfg.label, tier_color: tierCfg.color,
    voucher_label: tierCfg.voucher_label,
    points_balance: spendable,
    spending_points: lifetimePts,
    next_tier: nextTier,
    next_tier_label: nextTierCfg ? nextTierCfg.label : null,
    points_to_next_tier: nextTierCfg ? Math.max(0, nextTierCfg.min - lifetimePts) : 0,
    voucher_used_this_month: voucherUsedThisMonth,
    available_loyalty_vouchers: availableTiers,
    welcome_bonus_available: !voucherUsedThisMonth && !hasUsedWelcome(clientId),
    birthday_bonus_available: !voucherUsedThisMonth && !hasUsedBirthdayThisYear(clientId),
  };
}

function earnPoints(clientId, amount, source, description, bookingId) {
  const id = uid();
  const now = new Date();
  // Expiry: 12 months from end of current month
  const expires = new Date(now.getFullYear(), now.getMonth() + 12, 0, 23, 59, 59);
  
  const tx = {
    id, client_id: clientId, type: 'earn', amount, source,
    description, booking_id: bookingId || null,
    created_at: now.toISOString(), expires_at: expires.toISOString()
  };
  saveOne('points_tx', id, tx);
  
  // Update client's loyalty_points
  const client = getOne('clients', clientId);
  if (client) {
    saveOne('clients', clientId, {
      loyalty_points: (client.loyalty_points || 0) + amount
    });
  }
  
  return tx;
}

// === Auto award loyalty points when booking is completed (global — used by admin CRUD + client complete) ===
function awardBookingPoints(booking, clientId) {
  const baseAmount = (booking.amount || 0) + (parseInt(booking.discount) || 0);
  const client = getOne('clients', clientId);
  if (!client) return { earned: 0 };
  
  const oldSpending = client.total_spending || 0;
  const newSpending = oldSpending + baseAmount;
  const oldPts = Math.floor(oldSpending / 10000);
  const newPts = Math.floor(newSpending / 10000);
  const pointsEarned = Math.max(0, newPts - oldPts);
  
  // Update client total_spending
  saveOne('clients', clientId, { total_spending: newSpending });
  
  if (pointsEarned > 0) {
    earnPoints(clientId, pointsEarned, 'booking', 
      `Akumulasi: +Rp${baseAmount.toLocaleString('id-ID')} → total Rp${newSpending.toLocaleString('id-ID')} (${newPts} pts)`, 
      booking.id);
  }
  
  const mem = getMembership(clientId);
  return {
    earned: pointsEarned,
    tier: mem.tier_label,
    tier_voucher: mem.voucher_label,
    balance: mem.points_balance,
    tier_upgraded: calculateTier(oldPts) !== calculateTier(newPts) ? mem.tier_label : null
  };
}

function redeemPoints(clientId, points) {
  const mem = getMembership(clientId);
  if (mem.points_balance < points) return { error: 'Poin tidak mencukupi', balance: mem.points_balance };
  
  const discountAmount = points * 1000; // 1 poin = Rp1.000
  const id = uid();
  const tx = {
    id, client_id: clientId, type: 'redeem', amount: -points,
    source: 'redeem', description: `Tukar ${points} poin → diskon Rp${discountAmount.toLocaleString('id-ID')}`,
    created_at: new Date().toISOString(), expires_at: null
  };
  saveOne('points_tx', id, tx);
  
  return { success: true, points_redeemed: points, discount_amount: discountAmount, balance_remaining: mem.points_balance - points };
}

function expireOldPoints() {
  const now = new Date();
  const all = readAll('points_tx');
  let expired = 0;
  for (const tx of all) {
    if (!tx.expires_at || tx.type === 'redeem' || tx.type === 'expire') continue;
    if (new Date(tx.expires_at) < now && tx.amount > 0) {
      // Mark as expired
      saveOne('points_tx', tx.id, { amount: 0, type: 'expire', description: (tx.description||'')+' [EXPIRED]' });
      expired++;
    }
  }
  return expired;
}

// ═══════════════════════════════════════════════════════════
// AI CONSULTATION — Non-medical FAQ only
// ═══════════════════════════════════════════════════════════

const CLINIC_FAQ = {
  'jam buka': 'Klinik Lelap Mom Baby Care buka setiap hari pukul 08.00–17.00 WIB. Khusus Sabtu-Minggu buka pukul 08.00–14.00.',
  'jam operasional': 'Klinik Lelap Mom Baby Care buka setiap hari pukul 08.00–17.00 WIB. Khusus Sabtu-Minggu buka pukul 08.00–14.00.',
  'buka': 'Klinik Lelap Mom Baby Care buka setiap hari pukul 08.00–17.00 WIB. Khusus Sabtu-Minggu buka pukul 08.00–14.00.',
  'tutup': 'Klinik Lelap tutup di luar jam operasional (08.00–17.00 Senin-Jumat, 08.00–14.00 Sabtu-Minggu).',
  'alamat': 'Lelap Mom Baby Care berlokasi di Jl. Taman Pahlawan, Salatiga, Jawa Tengah. Tepatnya di depan taman, area pusat kota.',
  'lokasi': 'Lelap Mom Baby Care berlokasi di Jl. Taman Pahlawan, Salatiga, Jawa Tengah.',
  'homecare': 'Layanan Homecare tersedia untuk area Salatiga dan Kabupaten Semarang (radius 20 km). Bidan kami akan datang ke rumah Anda. Biaya transport dihitung berdasarkan jarak.',
  'home care': 'Layanan Homecare tersedia untuk area Salatiga dan Kabupaten Semarang (radius 20 km). Bidan kami akan datang ke rumah Anda.',
  'layanan': 'Lelap menyediakan: Pijat Relaksasi Bayi, Renang & Terapi Air Hangat, Pijat Ibu Hamil, Pijat Ibu Nifas, Perawatan Tali Pusat, Baby Spa, dan Homecare Bidan. Detail cek di menu Layanan.',
  'harga': 'Harga layanan bervariasi mulai dari Rp50.000 (renang bayi) hingga Rp200.000+ (paket lengkap). Cek menu Layanan untuk daftar lengkap.',
  'biaya': 'Biaya layanan bervariasi mulai dari Rp50.000. Silakan cek halaman Layanan untuk detail harga per treatment.',
  'daftar': 'Untuk mendaftar, klik "Login with Google" di aplikasi. Isi data diri dan Anda otomatis jadi member Non-Tier. Gratis!',
  'member': 'Program Membership Lelap memiliki 4 tier: Bronze, Silver, Gold, dan Platinum. Setiap tier memberi kamu hak menggunakan voucher diskon (5%-30%) dengan menukar poin. 1 voucher per bulan!',
  'poin': 'Poin didapat dari setiap booking selesai (Rp10.000 = 1 poin) dan Google review (10 poin, 1x seumur hidup). 1 poin = Rp1.000. Poin bisa ditukar voucher diskon sesuai tier. Poin berlaku 12 bulan.',
  'diskon': 'Voucher diskon: Welcome 10% (1x seumur hidup), Ultah 10% (1x/tahun), dan Voucher Loyalty sesuai tier — Bronze 5% (30 poin), Silver 12% (70 poin), Gold 20% (100 poin), Platinum 30% (130 poin). 1 voucher per bulan!',
  'konsultasi': 'Konsultasi tentang klinik dan layanan tersedia gratis via chat ini. Untuk konsultasi medis, silakan hubungi bidan kami langsung.',
  'kontak': 'Hubungi kami via WhatsApp di nomor yang tertera di halaman utama aplikasi, atau datang langsung ke klinik.',
  'booking': 'Booking bisa dilakukan langsung via aplikasi. Pilih layanan, tanggal, jam, dan metode pembayaran. Konfirmasi instan!',
  'reschedule': 'Reschedule bisa dilakukan maksimal H-1 sebelum jadwal via menu Riwayat Booking > pilih booking > Ubah Jadwal.',
  'cancel': 'Pembatalan gratis jika dilakukan H-1 atau lebih. Kurang dari H-1, permintaan akan direview admin.',
  'pembayaran': 'Kami menerima pembayaran Cash, QRIS, dan Transfer Bank.',
  'bayi': 'Kami melayani bayi usia 0-24 bulan. Layanan unggulan: pijat relaksasi, baby spa, renang air hangat.',
  'ibu hamil': 'Kami menyediakan pijat ibu hamil, senam hamil, dan konsultasi laktasi.',
  'nifas': 'Layanan pijat ibu nifas dan perawatan pasca melahirkan tersedia di klinik maupun homecare.',
};

const MEDICAL_KEYWORDS = [
  'sakit', 'demam', 'batuk', 'pilek', 'diare', 'muntah', 'alergi', 'ruam',
  'obat', 'dosis', 'resep', 'diagnosa', 'gejala', 'penyakit', 'infeksi',
  'darah', 'luka', 'bengkak', 'kejang', 'sesak', 'napas', 'asma',
  'imunisasi', 'vaksin', 'campak', 'cacar', 'dbd', 'tipes', 'kuning',
  'lahir', 'operasi', 'caesar', 'prematur', 'stunting', 'gizi buruk',
  'konsultasi dokter', 'resep obat', 'surat dokter', 'rujukan',
];

function isMedicalQuestion(text) {
  const lower = text.toLowerCase();
  return MEDICAL_KEYWORDS.some(kw => lower.includes(kw));
}

function isRecommendationQuery(text) {
  const lower = text.toLowerCase();
  if (lower.includes('rekomendasi') || lower.includes('saran') || lower.includes('cocok') ||
      lower.includes('apa saja') || lower.includes('layanan apa')) return true;
  // Age-based: "bayi X bulan", "anak X tahun", "umur X"
  if (/(?:bayi|anak|balita|umur|usia)\s*\d+/.test(lower)) return true;
  if (/\d+\s*(?:bulan|tahun)/.test(lower)) return true;
  return false;
}

function answerRecommendation(text) {
  const lower = text.toLowerCase();
  const svcs = readAll('services');
  
  // Detect age
  let ageMatch = text.match(/(\d+)\s*(?:bulan|tahun)/);
  if (!ageMatch) ageMatch = text.match(/(?:bayi|anak|balita|umur|usia)\s*(\d+)/);
  
  let filtered, label;
  if (ageMatch) {
    const num = parseInt(ageMatch[1]);
    const isMonths = text.includes('bulan') || (ageMatch[0].includes('bayi') && num <= 12);
    if (isMonths && num <= 12) {
      filtered = svcs.filter(s => s.category === 'BABY' || s.category === 'HOMECARE BABY KID');
      label = `👶 Untuk bayi ${num} bulan, rekomendasi kami:`;
    } else if (num >= 2 && num <= 6) {
      filtered = svcs.filter(s => s.category === 'KID' || s.category === 'HOMECARE BABY KID');
      label = `🧒 Untuk anak ${num} tahun, rekomendasi kami:`;
    } else {
      filtered = svcs.filter(s => s.category === 'BABY' || s.category === 'MOM' || s.category === 'KID' || s.category === 'HOMECARE BABY KID' || s.category === 'HOMECARE MOM');
      label = `📋 Rekomendasi layanan:`;
    }
  } else {
    // Generic recommendation
    filtered = svcs.filter(s => s.category === 'BABY' || s.category === 'MOM' || s.category === 'KID');
    label = `📋 Rekomendasi layanan Lelap:`;
  }
  
  if (!filtered.length) {
    return { type: 'general', message: 'Maaf, tidak ada layanan yang cocok. Coba tanyakan lebih spesifik ya, Ma~ 😊' };
  }
  
  let msg = `${label}\n\n`;
  const top = filtered.slice(0, 8);
  top.forEach(s => {
    const sess = s.sessions > 1 ? ` • Paket ${s.sessions}×` : '';
    msg += `✅ **${s.code}** — ${s.name}\n   ⏱ ${s.duration} menit | 💰 Rp ${(s.price||0).toLocaleString('id-ID')}${sess}\n   ℹ️ ${(s.description||'').replace(/\\n/g, ' ')}\n\n`;
  });
  msg += '📲 Detail lengkap & booking: cek menu **Layanan** ya, Ma~ 😊';
  return { type: 'faq', message: msg };
}

// ═══════════════════════════════════════════════════════════
// AI CONSULTATION — OpenRouter API
// ═══════════════════════════════════════════════════════════
const OPENROUTER_KEY = process.env.OPENROUTER_API_KEY || '';
const AI_MODEL = 'google/gemini-2.5-flash-lite';

// ── Build real-time service catalog for AI context ──
function getServiceCatalog() {
  const svcs = readAll('services');
  const cats = { 'BABY': '👶 Layanan Bayi (Inhouse)', 'MOM': '🤰 Layanan Ibu (Inhouse)', 
                 'KID': '🧒 Layanan Anak (Inhouse)', 'HOMECARE BABY KID': '🏠 Homecare Bayi & Anak',
                 'HOMECARE MOM': '🏠 Homecare Ibu' };
  let out = '===== KATALOG LAYANAN LELAP =====\n';
  for (const [cat, label] of Object.entries(cats)) {
    const items = svcs.filter(s => s.category === cat).sort((a,b) => (a.code||'').localeCompare(b.code||''));
    if (!items.length) continue;
    out += `\n📌 ${label}:\n`;
    for (const s of items) {
      const desc = (s.description || '').replace(/\n/g, ' ');
      const sess = s.sessions > 1 ? ` • ${s.sessions}× sesi paket` : '';
      out += `  ${s.code} — ${s.name} | ${s.duration} menit | Rp ${(s.price||0).toLocaleString('id-ID')}${sess}\n    ${desc}\n`;
    }
  }
  return out;
}

async function askAI(question, context) {
  if (!OPENROUTER_KEY || OPENROUTER_KEY === '***') return null;
  try {
    const now = new Date();
    const options = { weekday: 'long', year: 'numeric', month: 'long', day: 'numeric', timeZone: 'Asia/Jakarta' };
    const todayStr = now.toLocaleDateString('id-ID', options);
    const dayName = now.toLocaleDateString('id-ID', { weekday: 'long', timeZone: 'Asia/Jakarta' });
    const isMonday = dayName === 'Senin';
    
    const scheduleInfo = isMonday
      ? 'HARI INI: ' + dayName + ' (' + todayStr + ') — TUTUP (LIBUR). HANYA Prenatal Class 16.00-17.00.'
      : 'HARI INI: ' + dayName + ' (' + todayStr + ') — BUKA 08.00-16.00.';
    
    const serviceCatalog = getServiceCatalog();
    const pointsKnowledge = '\\n===== POIN & MEMBERSHIP =====\\nLOYALTY FORMULA: floor(total_spending / 10000) — setiap kelipatan Rp10.000 = 1 poin lifetime.\\n\\nCARA DAPAT POIN:\\n- Booking selesai: 1 poin per Rp10.000 spending (otomatis)\\n- Google Review: 10 poin (1x seumur hidup, harus diverifikasi admin)\\n- Welcome Bonus: diskon 10% untuk member baru (1x seumur hidup, saat Non-Tier)\\n- Birthday Bonus: diskon 10% (1x per tahun, klaim manual)\\n\\nTIER MEMBERSHIP (berdasarkan spending_points / total poin seumur hidup):\\n- Non-Tier: 0-29 poin — belum dapat voucher loyalty\\n- Bronze: 30-69 poin — voucher 5%, tukar 30 poin\\n- Silver: 70-99 poin — voucher 12%, tukar 70 poin\\n- Gold: 100-129 poin — voucher 20%, tukar 100 poin\\n- Platinum: 130+ poin — voucher 30%, tukar 130 poin\\n\\nPENTING:\\n- 1 poin = Rp1.000 nilai diskon saat redeem voucher\\n- Voucher loyalty maksimal 1x per bulan\\n- Poin berlaku 12 bulan dari perolehan\\n- Naik tier otomatis, tidak bisa turun\\n- Voucher ditukar dengan poin (bukan gratis) — poin berkurang saat redeem\\n';
    const systemMsg = 'Kamu asisten AI Lelap Mom Baby Care, Jl. Taman Pahlawan Salatiga. Bantu info layanan, harga, booking, membership. JANGAN jawab pertanyaan medis/kesehatan → arahkan ke WhatsApp 081-313-99-636. Panggil customer "Mama". IG @Lelap.Salatiga.\\n\\n' + scheduleInfo + '\\n\\nJAM OPERASIONAL: Senin LIBUR (kecuali Prenatal Class 16.00-17.00). Selasa-Minggu 08.00-16.00. Layanan maksimal 19.00. Booking setelah 16.00 WAJIB sebelum 15.00 di hari sama.\\n\\n' + serviceCatalog + pointsKnowledge + '\\n===== CONTOH RESPONS YANG BAIK =====\\nSebutkan kode layanan (misal B001), nama lengkap, durasi, dan harga saat menjawab pertanyaan. Jika Mama tanya rekomendasi, berikan 2-3 opsi yang relevan dengan penjelasan singkat. Untuk pertanyaan poin/membership, jelaskan tier, cara dapat poin, dan cara tukar voucher.\\n\\n' + (context || '');
    
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 8000);
    
    const resp = await fetch('https://openrouter.ai/api/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + OPENROUTER_KEY,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        model: AI_MODEL,
        messages: [
          { role: 'system', content: systemMsg },
          { role: 'user', content: question }
        ],
        max_tokens: 600,
        temperature: 0.4,
      }),
      signal: controller.signal
    });
    clearTimeout(timeout);
    const data = await resp.json();
    return data.choices?.[0]?.message?.content || null;
  } catch (e) {
    console.error('AI consultation error:', e.message);
    return null;
  }
}

function answerFAQ(text) {
  const lower = text.toLowerCase();
  
  // Dynamic: current day
  if (lower.includes('hari ini') || lower.includes('hari apa') || lower.includes('sekarang hari')) {
    const now = new Date();
    const dayName = now.toLocaleDateString('id-ID', { weekday: 'long', timeZone: 'Asia/Jakarta' });
    const dateStr = now.toLocaleDateString('id-ID', { year: 'numeric', month: 'long', day: 'numeric', timeZone: 'Asia/Jakarta' });
    const isMonday = dayName === 'Senin';
    const msg = isMonday
      ? `Hari ini ${dayName}, ${dateStr}.\\n\\n⚠️ Klinik TUTUP (LIBUR). HANYA Prenatal Class pukul 16.00-17.00.\\n\\nLayanan normal buka lagi besok Selasa pukul 08.00-16.00.`
      : `Hari ini ${dayName}, ${dateStr}.\\n\\n✅ Klinik BUKA pukul 08.00-16.00.\\n\\nUntuk booking hari ini, pastikan sebelum pukul 15.00 ya, Ma~ 😊`;
    return { type: 'faq', message: msg };
  }
  
  // Check medical first
  if (isMedicalQuestion(text)) {
    return {
      type: 'rejected',
      message: '⚠️ Maaf, saya tidak bisa menjawab pertanyaan medis. Untuk konsultasi kesehatan, silakan:\n\n✅ Hubungi bidan Lelap langsung via WhatsApp\n✅ Booking sesi konsultasi bidan di aplikasi (menu Layanan)\n✅ Kunjungi klinik untuk pemeriksaan langsung\n\nSaya hanya bisa membantu pertanyaan seputar layanan, jam buka, harga, dan informasi klinik ya, Ma~ 😊'
    };
  }
  
  // Match FAQ
  for (const [keyword, answer] of Object.entries(CLINIC_FAQ)) {
    if (lower.includes(keyword)) {
      return { type: 'faq', message: answer };
    }
  }
  
  // Recommendation / age-based queries → build from catalog
  if (lower.includes('rekomendasi') || lower.includes('saran') || lower.includes('cocok') ||
      /(?:umur|usia|bulan|tahun)\s*\d+/.test(lower)) {
    const svcs = readAll('services');
    const ageMatch = text.match(/(\d+)\s*(?:bulan|tahun)/);
    const isBaby = ageMatch && ((text.includes('bulan') && parseInt(ageMatch[1]) <= 12) || parseInt(ageMatch[1]) <= 1);
    const isKid = ageMatch && (text.includes('tahun') && parseInt(ageMatch[1]) >= 2);
    
    let recs = [];
    if (isBaby) recs = svcs.filter(s => s.category === 'BABY' || s.category === 'HOMECARE BABY KID');
    else if (isKid) recs = svcs.filter(s => s.category === 'KID' || s.category === 'HOMECARE BABY KID');
    else recs = svcs.filter(s => s.category === 'BABY' || s.category === 'MOM' || s.category === 'KID');
    
    if (recs.length) {
      let msg = '📋 Rekomendasi layanan:\n\n';
      recs.slice(0, 6).forEach(s => {
        msg += `• **${s.code}** — ${s.name} | ${s.duration} menit | Rp ${(s.price||0).toLocaleString('id-ID')}\n   ${(s.description||'').replace(/\n/g, ' ')}\n\n`;
      });
      msg += 'Detail lengkap cek menu Layanan ya, Ma~ 😊';
      return { type: 'faq', message: msg };
    }
  }
  
  // Generic non-medical
  return {
    type: 'general',
    message: 'Terima kasih pertanyaannya! 🤗\n\nUntuk informasi lebih detail, silakan:\n• Cek menu Layanan di aplikasi\n• Hubungi kami via WhatsApp\n• Atau kunjungi klinik langsung di Jl. Taman Pahlawan, Salatiga\n\nAda yang bisa saya bantu lagi?'
  };
}

// ═══════════════════════════════════════════════════════════

// ── API Routes ──
function setupPublicRoutes(app, helpers) {
  const { readAll, saveOne, getOne, uid, getSettings, createJWT, verifyJWT } = helpers;
  console.log('DEBUG helpers loaded:', typeof readAll, typeof saveOne, typeof getOne);
  
  // === AUTH ===
  app.post('/api/public/auth/google', async (req, res) => {
    const { firebase_token, name, phone, lat, lng, birth_date } = req.body;
    if (!firebase_token) return res.status(400).json({ error: 'firebase_token required' });
    
    const googleUser = await verifyGoogleToken(firebase_token);
    if (!googleUser) return res.status(401).json({ error: 'Invalid Google token' });
    
    const email = googleUser.email;
    const googleId = googleUser.sub;
    
    // ── GPS Radius Check (for NEW users only) ──
    let distanceKm = null;
    let withinRadius = true;
    
    if (lat !== undefined && lng !== undefined) {
      distanceKm = haversineDistance(LELAP_LAT, LELAP_LNG, parseFloat(lat), parseFloat(lng));
      distanceKm = Math.round(distanceKm * 10) / 10;
      const maxRadius = (loadTransportRates ? loadTransportRates().max_radius_km : 20) || 20;
      withinRadius = distanceKm <= maxRadius;
    }
    
    // Find existing client
    let client = null;
    const allClients = readAll('clients');
    client = allClients.find(c => c.google_id === googleId || c.email === email);
    
    // Existing users always allowed
    if (client) {
      const jwt = createJWT({ id: client.id, email: client.email, name: client.name });
      return res.json({ token: jwt, user: client, new_user: false, distance_km: distanceKm });
    }
    
    // New user — validate phone
    if (!phone || !phone.toString().trim()) {
      return res.status(400).json({ error: 'Nomor WhatsApp wajib diisi saat registrasi.' });
    }
    
    // New user — must be within radius
    if (!withinRadius) {
      const maxR = (loadTransportRates().max_radius_km || 20);
      return res.status(403).json({
        error: `Maaf, saat ini Lelap hanya melayani area dalam radius ${maxR} km dari Salatiga. Jarak Anda ±${distanceKm} km.`,
        distance_km: distanceKm,
        max_radius_km: maxR,
        need_gps: !lat || !lng
      });
    }
    
    // Also validate via city/district if no GPS
    if (req.body.city && req.body.district) {
      if (!validateLocation(req.body.city, req.body.district)) {
        return res.status(403).json({
          error: 'Maaf, saat ini Lelap hanya melayani area Salatiga dan Kabupaten Semarang.',
          distance_km: distanceKm
        });
      }
    }
    
    let newUser = false;
    if (!client) {
      const id = uid();
      client = saveOne('clients', id, {
        google_id: googleId,
        email: email,
        name: name || googleUser.name || email.split('@')[0],
        phone: phone || '',
        address: '',
        city: '',
        district: '',
        profiles: [{ id: 'prof1', name: name || googleUser.name || email.split('@')[0], type: 'adult', gender: 'female', birth_date: birth_date || '', birthday_edited: false }],
        loyalty_points: 0
      });
      newUser = true;
    }
    
    const jwt = createJWT({ id: client.id, email: client.email, name: client.name });
    res.json({ token: jwt, user: client, new_user: newUser });
  });
  
  app.get('/api/public/profile', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    res.json(client);
  });
  
  app.put('/api/public/profile', publicAuth, (req, res) => {
    const { name, phone, address, city, district, birth_date } = req.body;
    const updates = {};
    if (name !== undefined) updates.name = name;
    if (phone !== undefined) updates.phone = phone;
    if (address !== undefined) updates.address = address;
    if (city !== undefined) updates.city = city;
    if (district !== undefined) updates.district = district;
    if (birth_date !== undefined) {
      const client = getOne('clients', req.client.id);
      const profiles = (client?.profiles || []).map(p => {
        if (p.type === 'adult') return { ...p, birth_date };
        return p;
      });
      updates.profiles = profiles;
    }
    
    const client = saveOne('clients', req.client.id, updates);
    res.json(client);
  });
  
  // === PROFILE PHOTO ===
  app.post('/api/public/profile/photo', publicAuth, (req, res) => {
    const { photo } = req.body; // base64 or data URL
    if (!photo) return res.status(400).json({ error: 'photo required' });
    // Max ~500KB after base64
    if (photo.length > 700000) return res.status(400).json({ error: 'Foto terlalu besar (max 500KB)' });
    const client = saveOne('clients', req.client.id, { photo });
    res.json({ photo_url: client.photo ? `data:image/jpeg;base64,${client.photo.replace(/^data:image\/\w+;base64,/, '')}` : null });
  });
  
  app.delete('/api/public/profile/photo', publicAuth, (req, res) => {
    saveOne('clients', req.client.id, { photo: null });
    res.json({ photo_url: null });
  });
  
  // === MULTI-PROFILE ===
  app.post('/api/public/profiles', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const profiles = client.profiles || [];
    const profileType = req.body.type || 'child';
    
    // MAX 3 CHILDREN
    if (profileType === 'child') {
      const childCount = profiles.filter(p => p.type === 'child').length;
      if (childCount >= 3) {
        return res.status(400).json({ error: 'Maksimal 3 anak per akun. Tidak bisa menambahkan lagi.' });
      }
    }
    
    const newProfile = {
      id: 'prof' + (profiles.length + 1),
      name: req.body.name,
      type: profileType,
      gender: req.body.gender || 'female',
      birth_date: req.body.birth_date || '',
      birthday_edited: false,
      notes: req.body.notes || ''
    };
    profiles.push(newProfile);
    saveOne('clients', req.client.id, { profiles });
    res.json(newProfile);
  });

  app.delete('/api/public/profiles/:id', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    const profiles = (client.profiles || []).filter(p => p.id !== req.params.id);
    saveOne('clients', req.client.id, { profiles });
    res.json({ status: 'deleted' });
  });

  // Edit profile — anak: birthday 1x only, nama locked; mama: name + birth_date
  app.put('/api/public/profiles/:id', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const profiles = client.profiles || [];
    const idx = profiles.findIndex(p => p.id === req.params.id);
    if (idx === -1) return res.status(404).json({ error: 'Profile not found' });
    
    const profile = profiles[idx];
    const { name, birth_date } = req.body;
    
    // CHILD: name locked, birthday only 1x
    if (profile.type === 'child') {
      if (name !== undefined) {
        return res.status(400).json({ error: 'Nama anak tidak bisa diubah.' });
      }
      if (birth_date !== undefined) {
        if (profile.birthday_edited) {
          return res.status(400).json({ error: 'Tanggal lahir anak hanya bisa diedit 1 kali.' });
        }
        profile.birth_date = birth_date;
        profile.birthday_edited = true;
      }
    } else {
      // MAMA (adult): can edit name and birth_date
      if (name !== undefined) profile.name = name;
      if (birth_date !== undefined) profile.birth_date = birth_date;
    }
    
    profiles[idx] = profile;
    saveOne('clients', req.client.id, { profiles });
    res.json(profile);
  });

  // Add family member (sama seperti profiles — untuk kompatibilitas)
  app.post('/api/public/family-members', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const profiles = client.profiles || [];
    const profileType = req.body.type || 'child';
    
    // MAX 3 CHILDREN
    if (profileType === 'child') {
      const childCount = profiles.filter(p => p.type === 'child').length;
      if (childCount >= 3) {
        return res.status(400).json({ error: 'Maksimal 3 anak per akun. Tidak bisa menambahkan lagi.' });
      }
    }
    
    const newProfile = {
      id: 'prof' + (profiles.length + 1),
      name: req.body.name,
      type: profileType,
      gender: req.body.gender || 'female',
      birth_date: req.body.birth_date || '',
      birthday_edited: false,
      notes: req.body.notes || ''
    };
    profiles.push(newProfile);
    saveOne('clients', req.client.id, { profiles });
    res.json(newProfile);
  });

  // === SERVICES ===
  app.get('/api/public/services', (req, res) => {
    const allRaw = readAll('services');
    console.log('DEBUG services count:', allRaw.length);
    let services = allRaw;
    const { category, search } = req.query;
    
    if (category && category !== 'ALL') {
      services = services.filter(s => (s.category || '').toUpperCase() === category.toUpperCase());
    }
    if (search) {
      const q = search.toLowerCase();
      services = services.filter(s => (s.name || '').toLowerCase().includes(q));
    }
    
    // Get categories
    const categories = [...new Set(services.map(s => s.category || 'OTHER'))];
    
    res.json({ services, categories: ['ALL', ...categories] });
  });
  
  app.get('/api/public/services/:id', publicAuth, (req, res) => {
    const service = getOne('services', req.params.id);
    if (!service) return res.status(404).json({ error: 'Service not found' });
    res.json(service);
  });
  
  // === INSTAGRAM FEED ===
  app.get('/api/public/ig-feed', (req, res) => {
    try {
      const feedPath = path.join(DATA_DIR, 'ig_feed.json');
      if (!fs.existsSync(feedPath)) {
        return res.json({ profile: 'lelap.salatiga', post_count: 0, posts: [], fetched_at: null });
      }
      const feed = JSON.parse(fs.readFileSync(feedPath, 'utf8'));
      // Only return last 20
      res.json({ ...feed, posts: (feed.posts || []).slice(0, 20) });
    } catch (e) {
      res.json({ profile: 'lelap.salatiga', post_count: 0, posts: [], error: e.message });
    }
  });
  
  // === STAFF ===
  app.get('/api/public/staff', publicAuth, (req, res) => {
    const staff = readAll('staff');
    res.json(staff);
  });
  
  // === SLOTS ===
  app.get('/api/public/slots', publicAuth, (req, res) => {
    const { date, service_id, therapist } = req.query;
    if (!date || !service_id) return res.status(400).json({ error: 'date and service_id required' });
    
    const service = getOne('services', service_id);
    if (!service) return res.status(404).json({ error: 'Service not found' });
    
    const allAppointments = readAll('appointments');
    const settings = getOne('settings', 'settings') || {};
    
    const slots = calculateSlots(date, service, therapist || null, allAppointments, settings);
    res.json({ date, service_id, slots });
  });
  
  // === BOOKINGS ===
  app.post('/api/public/bookings', publicAuth, async (req, res) => {
    const { service_id, date, time, dates, therapist, profile_id, payment_method, deposit,
            booking_type, client_lat, client_lng, client_address, voucher } = req.body;
    if (!service_id) return res.status(400).json({ error: 'service_id required' });
    
    const service = getOne('services', service_id);
    if (!service) return res.status(404).json({ error: 'Service not found' });
    
    const sessions = service.sessions || 1;
    const isMultiSession = sessions > 1;
    
    // Multi-session requires dates array, single requires date+time
    if (isMultiSession) {
      if (!dates || !Array.isArray(dates) || dates.length !== sessions) {
        return res.status(400).json({ error: `dates array with ${sessions} dates required for this package` });
      }
      if (!time) return res.status(400).json({ error: 'time required' });
    } else {
      if (!date || !time) return res.status(400).json({ error: 'date, time required' });
    }
    
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const profile = (client.profiles || []).find(p => p.id === profile_id) || client.profiles?.[0] || {};
    
    // Validate slot
    const allAppointments = readAll('appointments');
    const settings = getOne('settings', 'settings') || {};
    const slots = calculateSlots(date, service, therapist || null, allAppointments, settings);
    const slot = slots.find(s => s.time === time);
    if (!slot || !slot.available) {
      return res.status(409).json({ error: 'Slot tidak tersedia. Silakan pilih jam atau hari lain.', slot });
    }
    
    // Location validation: frontend already validates via GPS bounding box
    // Homecare bookings have client_lat/client_lng for transport calculation
    
    // === TIME VALIDATION: today booking restrictions ===
    const nowWIB = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta' }));
    const todayWIB = nowWIB.toISOString().substring(0, 10);
    const hourWIB = nowWIB.getHours();
    const minWIB = nowWIB.getMinutes();
    
    function validateDateTime(d, tm) {
      if (d === todayWIB) {
        if (hourWIB >= 15) {
          return 'Booking untuk hari ini hanya sampai jam 15:00. Silakan pilih besok atau hari lain.';
        }
        const [sh, sm] = (tm || '00:00').split(':').map(Number);
        const slotEndMin = sh * 60 + sm + (service.duration || 60);
        const nowMin = hourWIB * 60 + minWIB;
        if (slotEndMin <= nowMin) {
          return 'Jam sudah lewat. Silakan pilih jam yang lebih siang.';
        }
      }
      return null;
    }
    
    // Build list of dates to book
    const bookDates = isMultiSession ? dates : [date];
    
    // Validate time for first date (multi-session uses same time for all)
    const timeErr = validateDateTime(bookDates[0], time);
    if (timeErr) return res.status(400).json({ error: timeErr });
    
    // Validate all multi-session dates have available slots
    if (isMultiSession) {
      for (const d of bookDates) {
        const err = validateDateTime(d, time);
        if (err) return res.status(400).json({ error: `Tanggal ${d}: ${err}` });
        const daySlots = calculateSlots(d, service, therapist || null, readAll('appointments'), getOne('settings', 'settings') || {});
        const daySlot = daySlots.find(s => s.time === time);
        if (!daySlot || !daySlot.available) {
          return res.status(400).json({ error: `Slot ${time} tidak tersedia di tanggal ${d}. Silakan pilih tanggal lain.` });
        }
      }
    }
    
    const isHomecare = booking_type === 'homecare';
    const fixTherapist = service.name && service.name.toUpperCase().includes('PRENATAL YOGA') ? 'Owner' : (therapist || '');
    
    // === TRANSPORT CALCULATION (once, shared across sessions) ===
    let transportPrice = 0;
    let transportStatus = null;
    let distanceKm = null;
    let clientLat = null;
    let clientLng = null;
    
    if (isHomecare && client_lat && client_lng) {
      clientLat = parseFloat(client_lat);
      clientLng = parseFloat(client_lng);
      const clinicSettings = getOne('settings', 'settings') || {};
      const clinicLat = clinicSettings.clinic_lat || LELAP_LAT;
      const clinicLng = clinicSettings.clinic_lng || LELAP_LNG;
      
      let roadKm = null;
      try {
        const https = require('https');
        const osrmResult = await new Promise((resolve) => {
          const osrmUrl = `https://router.project-osrm.org/route/v1/driving/${clinicLng},${clinicLat};${clientLng},${clientLat}?overview=false`;
          const req = https.get(osrmUrl, { timeout: 6000 }, (resp) => {
            let data = '';
            resp.on('data', chunk => data += chunk);
            resp.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve(null); } });
          });
          req.on('error', () => resolve(null));
          req.on('timeout', () => { req.destroy(); resolve(null); });
        });
        if (osrmResult && osrmResult.routes && osrmResult.routes[0]) {
          roadKm = osrmResult.routes[0].distance / 1000;
        }
      } catch (_) {}
      
      if (roadKm && roadKm > 0) {
        distanceKm = Math.ceil(roadKm * 1.15 * 10) / 10;
      } else {
        const straight = haversineDistance(clinicLat, clinicLng, clientLat, clientLng);
        distanceKm = Math.ceil(straight * 1.6 * 10) / 10;
      }
      
      const transport = getTransportPrice(distanceKm);
      transportPrice = transport.price;
      transportStatus = transport.needs_approval ? 'pending_approval' : 'auto';
    }
    
    // === VOUCHER: 1 voucher per calendar month (pick one) ===
    let voucherLabel = '';
    let voucherDiscount = 0;
    const effectiveVoucher = isMultiSession ? null : voucher;
    const amount = getPrice(service, bookDates[0]);
    
    if (effectiveVoucher && !isMultiSession) {
      const mem = getMembership(client.id);
      
      // Check monthly limit first
      if (mem.voucher_used_this_month) {
        return res.status(400).json({ error: 'Kamu sudah menggunakan voucher bulan ini. Coba lagi bulan depan!' });
      }
      
      if (effectiveVoucher === 'birthday') {
        if (hasUsedBirthdayThisYear(client.id)) {
          return res.status(400).json({ error: 'Voucher ultah sudah digunakan tahun ini.' });
        }
        voucherLabel = '🎂 Ultah 10%';
        voucherDiscount = Math.round(amount * 0.10);
      } else if (effectiveVoucher === 'welcome') {
        if (hasUsedWelcome(client.id)) {
          return res.status(400).json({ error: 'Voucher welcome sudah pernah digunakan.' });
        }
        voucherLabel = '🎁 Welcome 10%';
        voucherDiscount = Math.round(amount * 0.10);
      } else if (effectiveVoucher.startsWith('loyalty_')) {
        const tierName = effectiveVoucher.replace('loyalty_', '');
        const cfg = getTierConfig(tierName);
        if (!cfg || tierName === 'non-tier') {
          return res.status(400).json({ error: 'Voucher loyalty tidak valid.' });
        }
        if (mem.points_balance < cfg.voucher_cost) {
          return res.status(400).json({ error: `Poin tidak cukup. Butuh ${cfg.voucher_cost} poin, kamu punya ${mem.points_balance} poin.` });
        }
        // Deduct points
        saveOne('clients', client.id, {
          loyalty_points: mem.points_balance - cfg.voucher_cost
        });
        // Record transaction
        saveOne('points_tx', uid(), {
          client_id: client.id, type: 'redeem', amount: -cfg.voucher_cost,
          source: 'voucher_loyalty',
          description: `Voucher ${cfg.label} ${cfg.voucher_pct}% (-${cfg.voucher_cost} poin)`,
          created_at: new Date().toISOString(), expires_at: null
        });
        voucherLabel = `💎 ${cfg.label} ${cfg.voucher_pct}%`;
        voucherDiscount = Math.round(amount * cfg.voucher_pct / 100);
      }
    }
    
    // NO automatic membership discount — user picks one voucher above
    const appliedDiscount = voucherDiscount;
    
    // === Per-session amounts ===
    const totalAmount = getPrice(service, bookDates[0]);
    const perSessionAmount = isMultiSession ? Math.round(totalAmount / sessions) : totalAmount;
    const perSessionTransport = transportPrice; // transport price is per-visit, multiply by sessions for total
    const perSessionDiscount = isMultiSession ? 0 : appliedDiscount;
    
    // === Create bookings ===
    const sessionGroup = isMultiSession ? uid() : null;
    const bookings = [];
    
    for (let i = 0; i < bookDates.length; i++) {
      const d = bookDates[i];
      const bookingCode = 'MBS-' + d.replace(/-/g, '').substring(2) + '-' + Math.floor(Math.random() * 900 + 100);
      const finalAmount = Math.max(0, perSessionAmount - perSessionDiscount + perSessionTransport);
      
      const booking = saveOne('appointments', uid(), {
        date: d,
        type: isHomecare ? 'Homecare' : 'Inhouse',
        time,
        therapist: fixTherapist,
        kode: service.code || '',
        kategori: service.category || '',
        service: service.name,
        mother: profile.type === 'child' ? client.name : (profile.name || client.name),
        client_name: client.name,
        wa: client.phone || '',
        child: profile.type === 'child' ? profile.name : '',
        age: profile.birth_date || '',
        address: client_address || client.address || '',
        discount: String(perSessionDiscount),
        voucher: i === 0 ? voucherLabel : '',
        voucher_type: i === 0 ? (effectiveVoucher || '') : '',
        transport: String(perSessionTransport),
        transport_price: perSessionTransport,
        transport_status: transportStatus,
        deposit: deposit || '0',
        payment: payment_method === 'qris' ? 'QRIS' : payment_method === 'transfer' ? 'Transfer' : 'Cash',
        notes: isMultiSession ? `Sesi ${i + 1}/${sessions}` : '',
        client_type: profile.type === 'child' ? 'Anak' : 'Dewasa',
        staff: fixTherapist,
        status: 'Menunggu',
        amount: finalAmount,
        service_amount: perSessionAmount,
        booking_code: bookingCode,
        client_id: client.id,
        profile_id: profile_id || '',
        duration: service.duration || 60,
        source: 'app',
        client_lat: clientLat,
        client_lng: clientLng,
        distance_km: distanceKm,
        session_group: sessionGroup,
        session_index: isMultiSession ? i + 1 : null,
        session_total: isMultiSession ? sessions : null,
      });
      bookings.push(booking);
    }
    
    // Auto reverse-geocode client location
    if (clientLat && clientLng && (!client_address || client_address === '')) {
      const https = require('https');
      const geoUrl = `https://nominatim.openstreetmap.org/reverse?format=json&lat=${clientLat}&lon=${clientLng}`;
      https.get(geoUrl, { headers: { 'User-Agent': 'LelapBookingCare/1.0' } }, (resp) => {
        let gdata = '';
        resp.on('data', chunk => gdata += chunk);
        resp.on('end', () => {
          try {
            const result = JSON.parse(gdata);
            if (result && result.display_name) {
              bookings.forEach(b => saveOne('appointments', b.id, { address: result.display_name }));
            }
          } catch {}
        });
      }).on('error', () => {});
    }
    
    // Loyalty points hanya dari complete endpoint (cumulative: Rp10k = 1pt)
    
    // Send WhatsApp notification
    const b0 = bookings[0];
    const displayTotal = isMultiSession ? ` (${sessions}x sesi)` : '';
    const totalTransport = isMultiSession ? transportPrice * sessions : transportPrice;
    sendWA((getOne('settings', 'settings') || {}).whatsapp || '',
      `🔔 *Booking Baru!*\\n📋 ${b0.booking_code}${isMultiSession ? ' +' + (sessions-1) + ' sesi' : ''}\\n👤 ${profile.name || client.name}\\n💆 ${service.name}${displayTotal}\\n📅 ${isMultiSession ? bookDates.join(', ') : bookDates[0]} ${time}\\n📍 ${isHomecare ? 'HOMECARE' : 'Inhouse'}\\n💰 Rp ${(perSessionAmount * sessions + totalTransport).toLocaleString('id-ID')}${voucherLabel ? '\\n' + voucherLabel + ' -Rp ' + voucherDiscount.toLocaleString('id-ID') : ''}\\n🏠 ${client_address || client.address || '-'}\\n💳 ${payment_method === 'qris' ? 'QRIS' : payment_method === 'transfer' ? 'Transfer' : 'Cash'}${transportStatus === 'pending_approval' ? '\\n⚠️ Transport >15km — menunggu persetujuan SA' : ''}`);

    res.status(201).json({ 
      bookings, 
      booking_code: b0.booking_code,
      is_multi_session: isMultiSession,
      session_count: isMultiSession ? sessions : 1,
      per_session_amount: perSessionAmount,
      transport: { status: transportStatus, price: perSessionTransport, total: totalTransport, distance_km: distanceKm },
      voucher: voucherLabel, voucher_discount: perSessionDiscount 
    });
  });
  
  app.get('/api/public/bookings', publicAuth, (req, res) => {
    const all = readAll('appointments');
    const status = req.query.status;
    let mine = all.filter(a => a.client_id === req.client.id);
    
    // Tab filter
    if (status === 'upcoming') {
      mine = mine.filter(a => a.status !== 'Selesai' && a.status !== 'cancelled' && a.status !== 'Dibatalkan' &&
        new Date(a.date + 'T' + (a.time || '00:00') + ':00+07:00') >= new Date());
    } else if (status === 'completed') {
      mine = mine.filter(a => a.status === 'Selesai' || a.status === 'Completed');
    } else if (status === 'cancelled') {
      mine = mine.filter(a => a.status === 'cancelled' || a.status === 'Dibatalkan' || a.status === 'pending_cancel');
    }
    
    mine.sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
    res.json(mine);
  });
  
  app.get('/api/public/bookings/:id/invoice', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    const serviceInfo = booking.kode ? getOne('services', booking.kode.toLowerCase()) : null;
    const client = getOne('clients', req.client.id);
    
    // Check if points were earned from this booking
    const ptsTx = readAll('points_tx').filter(t => t.client_id === req.client.id && t.booking_id === booking.id && t.type === 'earn');
    const pointsEarned = ptsTx.reduce((sum, t) => sum + (t.amount || 0), 0);
    
    res.json({
      invoice_no: booking.booking_code || `INV-${booking.id.substring(0,8).toUpperCase()}`,
      date: booking.date,
      time: booking.time,
      service: booking.service,
      category: booking.kategori || serviceInfo?.category || '',
      type: booking.type || (booking.is_homecare ? 'Homecare' : 'Inhouse'),
      therapist: booking.therapist || booking.staff || '',
      client_name: booking.client_name || client?.name || '',
      child: booking.child || '',
      amount: parseInt(booking.service_amount) || parseInt(booking.amount) || 0,
      discount: parseInt(booking.discount) || 0,
      transport: parseInt(booking.transport) || 0,
      deposit: parseInt(booking.deposit) || 0,
      voucher: booking.voucher || '',
      total: parseInt(booking.amount) || 0,
      payment: booking.payment || 'Cash',
      status: booking.status,
      booking_date: booking.created_at,
      completed_at: booking.completed_at || null,
      points_earned: pointsEarned
    });
  });
  
  app.get('/api/public/bookings/:id', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    res.json(booking);
  });
  
  app.put('/api/public/bookings/:id/cancel', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    // Check if H-1 or more
    const bookingDate = new Date(booking.date + 'T' + (booking.time || '00:00') + ':00+07:00');
    const now = new Date();
    const hoursUntil = (bookingDate.getTime() - now.getTime()) / (1000 * 60 * 60);
    
    if (hoursUntil >= 24) {
      // Free cancel
      saveOne('appointments', req.params.id, { status: 'cancelled', cancelled_at: new Date().toISOString(), cancel_method: 'app_auto' });
      res.json({ status: 'cancelled', message: 'Booking berhasil dibatalkan.' });
    } else {
      // Need admin approval
      saveOne('appointments', req.params.id, { status: 'pending_cancel', cancel_requested_at: new Date().toISOString() });
      res.json({ status: 'pending_cancel', message: 'Permintaan pembatalan dikirim ke admin. Tim kami akan menghubungi Anda.' });
    }
  });
  
  app.put('/api/public/bookings/:id/reschedule', publicAuth, (req, res) => {
    const { date, time } = req.body;
    if (!date || !time) return res.status(400).json({ error: 'date and time required' });
    
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    const bookingDate = new Date(booking.date + 'T' + (booking.time || '00:00') + ':00+07:00');
    const now = new Date();
    const hoursUntil = (bookingDate.getTime() - now.getTime()) / (1000 * 60 * 60);
    
    if (hoursUntil < 24) {
      return res.status(400).json({ error: 'Reschedule hanya bisa dilakukan H-1 atau lebih. Silakan hubungi admin.' });
    }
    
    saveOne('appointments', req.params.id, { date, time, rescheduled_at: new Date().toISOString() });
    res.json({ status: 'rescheduled', message: 'Jadwal berhasil diubah.' });
  });
  
  // === REVIEWS ===
  // Check if booking already reviewed
  app.get('/api/bookings/:id/review', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    const review = readAll('reviews').find(r => r.booking_id === req.params.id && r.client_id === req.client.id);
    res.json({ reviewed: !!review, review: review || null });
  });

  app.post('/api/public/reviews', publicAuth, (req, res) => {
    const { booking_id, rating, comment, therapist_rating, photos } = req.body;
    if (!booking_id || !rating) return res.status(400).json({ error: 'booking_id and rating required' });
    
    // DEBUG
    console.log('REVIEW DEBUG - photos type:', typeof photos, 'isArray:', Array.isArray(photos), 'length:', photos ? photos.length : 'undefined');
    if (Array.isArray(photos) && photos.length > 0) {
      console.log('REVIEW DEBUG - photo sizes:', photos.map(p => (p || '').length));
    }
    
    const booking = getOne('appointments', booking_id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    // Check if already reviewed
    const existingReviews = readAll('reviews').filter(r => r.booking_id === booking_id && r.client_id === req.client.id);
    if (existingReviews.length > 0) return res.status(400).json({ error: 'Anda sudah memberikan review untuk booking ini' });
    
    const client = getOne('clients', req.client.id);
    const review = saveOne('reviews', uid(), {
      booking_id,
      client_id: req.client.id,
      client_name: client?.name || '',
      client_photo: client?.photo || null,
      rating,
      comment: comment || '',
      therapist_rating: therapist_rating || rating,
      therapist: booking.staff || '',
      service: booking.service,
      date: booking.date,
      photos: photos || [],
      google_review_claimed: false,
      google_review_verified: false,
      created_at: new Date().toISOString()
    });
    
    // Points only from admin Google review verification (not automatic)
    const mem = getMembership(req.client.id);
    
    res.status(201).json({
      ...review,
      points_earned: 0,
      tier: mem.tier_label,
      points_balance: mem.points_balance,
      google_review_link: 'https://www.google.com/maps/place/Lelap+Mom+Baby+Care+Salatiga/@-7.3285,110.5100,17z/data=!4m8!3m7!1s0x2e7a79a87d92c81f:0x61c355bf9ad20e62!8m2!3d-7.3285488!4d110.5100338!9m1!1b1'
    });
  });

  // Client claims they reviewed on Google (pending admin verification)
  app.post('/api/public/reviews/:id/claim-google', publicAuth, (req, res) => {
    const review = getOne('reviews', req.params.id);
    if (!review) return res.status(404).json({ error: 'Review not found' });
    if (review.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    if (review.google_review_verified) return res.status(400).json({ error: 'Google review sudah diverifikasi' });
    if (review.google_review_claimed) return res.status(400).json({ error: 'Anda sudah klaim Google review' });
    
    saveOne('reviews', req.params.id, {
      ...review,
      google_review_claimed: true,
      google_claimed_at: new Date().toISOString()
    });
    
    res.json({ success: true, message: 'Klaim diterima, menunggu verifikasi admin' });
  });

  // ═══════════════════════════════════════════════════════════
  // SERVICE IMAGE UPLOAD
  // ═══════════════════════════════════════════════════════════
  const serviceImageUpload = multer({
    storage: multer.diskStorage({
      destination: path.join(__dirname, 'data', 'service_images'),
      filename: (req, file, cb) => {
        const ext = path.extname(file.originalname).toLowerCase();
        cb(null, req.params.id + ext);
      }
    }),
    limits: { fileSize: 500 * 1024 }, // 500KB
    fileFilter: (req, file, cb) => {
      const allowed = ['.jpg', '.jpeg', '.png', '.webp'];
      const ext = path.extname(file.originalname).toLowerCase();
      if (allowed.includes(ext)) cb(null, true);
      else cb(new Error('Hanya JPG, PNG, WebP yang diizinkan'));
    }
  });

  app.post('/api/admin/services/:id/image', (req, res) => {
    serviceImageUpload.single('image')(req, res, (err) => {
      if (err) {
        if (err.code === 'LIMIT_FILE_SIZE') return res.status(400).json({ error: 'File terlalu besar. Maksimal 500KB.' });
        return res.status(400).json({ error: err.message });
      }
      if (!req.file) return res.status(400).json({ error: 'Tidak ada file yang diupload' });
      
      // Update service record
      const svc = getOne('services', req.params.id);
      if (!svc) return res.status(404).json({ error: 'Layanan tidak ditemukan' });
      const ext = path.extname(req.file.originalname).toLowerCase();
      saveOne('services', req.params.id, { ...svc, image: req.params.id + ext });
      
      res.json({ success: true, filename: req.params.id + ext, size: req.file.size });
    });
  });

  app.delete('/api/admin/services/:id/image', (req, res) => {
    const svc = getOne('services', req.params.id);
    if (!svc) return res.status(404).json({ error: 'Layanan tidak ditemukan' });
    if (!svc.image) return res.status(400).json({ error: 'Layanan tidak memiliki gambar' });
    
    const imgPath = path.join(__dirname, 'data', 'service_images', svc.image);
    if (fs.existsSync(imgPath)) fs.unlinkSync(imgPath);
    saveOne('services', req.params.id, { ...svc, image: null });
    
    res.json({ success: true });
  });

  // Admin: list reviews pending Google verification
  app.get('/api/admin/reviews/pending-google', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const pending = readAll('reviews').filter(r => r.google_review_claimed && !r.google_review_verified);
    res.json(pending);
  });

  // ═══════════════════════════════════════════════
  // MEMBERSHIP SCHEME API
  // ═══════════════════════════════════════════════
  app.get('/api/admin/membership-scheme', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || (tok.role !== 'admin' && tok.role !== 'owner')) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const scheme = loadMembershipScheme();
    if (!scheme) return res.status(500).json({ error: 'Failed to load scheme' });
    res.json(scheme);
  });

  app.put('/api/admin/membership-scheme', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || (tok.role !== 'admin' && tok.role !== 'owner')) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const { tiers, non_tier, welcome_voucher_pct, birthday_voucher_pct, max_voucher_per_month, loyalty_formula } = req.body;
    if (!tiers || !Array.isArray(tiers)) return res.status(400).json({ error: 'tiers array is required' });
    const scheme = {
      version: 1,
      last_updated: new Date().toISOString(),
      loyalty_formula: loyalty_formula || 'floor(total_spending / 10000)',
      max_voucher_per_month: max_voucher_per_month || 1,
      welcome_voucher_pct: welcome_voucher_pct || 10,
      birthday_voucher_pct: birthday_voucher_pct || 10,
      tiers,
      non_tier: non_tier || { label: 'Non-Tier', color: '#999', voucher_label: 'Belum ada voucher' }
    };
    try {
      fs.writeFileSync(path.join(DATA_DIR, 'membership-scheme.json'), JSON.stringify(scheme, null, 2));
      res.json({ success: true, message: 'Membership scheme updated', scheme });
    } catch (e) {
      res.status(500).json({ error: 'Failed to save scheme: ' + e.message });
    }
  });

  // ═══════════════════════════════════════════════
  // TRANSPORT RATES API
  // ═══════════════════════════════════════════════
  app.get('/api/admin/transport-rates', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || (tok.role !== 'admin' && tok.role !== 'owner')) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const cfg = loadTransportRates();
    res.json(cfg);
  });

  app.put('/api/admin/transport-rates', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || (tok.role !== 'admin' && tok.role !== 'owner')) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const { rates, max_radius_km, approval_threshold_km } = req.body;
    if (!rates || !Array.isArray(rates)) return res.status(400).json({ error: 'rates array is required' });
    const cfg = {
      version: 1,
      last_updated: new Date().toISOString(),
      max_radius_km: max_radius_km || 20,
      approval_threshold_km: approval_threshold_km || 15,
      approval_message: 'Jarak >' + (approval_threshold_km || 15) + 'km — harga transport akan dinegosiasikan dengan admin setelah booking.',
      rates
    };
    try {
      fs.writeFileSync(path.join(DATA_DIR, 'transport-rates.json'), JSON.stringify(cfg, null, 2));
      res.json({ success: true, message: 'Transport rates updated', config: cfg });
    } catch (e) {
      res.status(500).json({ error: 'Failed to save transport rates: ' + e.message });
    }
  });

  // Admin: verify Google review → give 10 points (1x per client lifetime)
  app.post('/api/admin/reviews/:id/verify-google', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const review = getOne('reviews', req.params.id);
    if (!review) return res.status(404).json({ error: 'Review not found' });
    if (!review.google_review_claimed) return res.status(400).json({ error: 'Client belum klaim Google review' });
    if (review.google_review_verified) return res.status(400).json({ error: 'Sudah diverifikasi' });
    
    // Check if client already got Google review points (1x lifetime)
    const alreadyGotPoints = readAll('points_tx').some(tx =>
      tx.client_id === review.client_id && tx.source === 'google_review'
    );
    
    if (alreadyGotPoints) {
      saveOne('reviews', req.params.id, {
        ...review,
        google_review_verified: true,
        google_verified_at: new Date().toISOString(),
        google_verified_by: tok.name || 'Admin'
      });
      return res.json({ success: true, points_earned: 0, message: 'Diverifikasi (poin sudah pernah diberikan sebelumnya)' });
    }
    
    // Give 10 points
    earnPoints(review.client_id, 10, 'google_review',
      `Google Review untuk ${review.service} (${'⭐'.repeat(review.rating)})`,
      review.booking_id);
    
    saveOne('reviews', req.params.id, {
      ...review,
      google_review_verified: true,
      google_verified_at: new Date().toISOString(),
      google_verified_by: tok.name || 'Admin'
    });
    
    res.json({ success: true, points_earned: 10, message: 'Diverifikasi! +10 poin diberikan' });
  });
  
  app.get('/api/public/reviews/:service_id', (req, res) => {
    const all = readAll('reviews');
    const serviceReviews = all.filter(r => {
      // Direct service_code match (template reviews)
      if (r.service_code === req.params.service_id) return true;
      // Match via booking
      const booking = getOne('appointments', r.booking_id);
      return booking && (booking.kode === req.params.service_id || booking.service === req.params.service_id);
    });
    const avg = serviceReviews.length > 0 
      ? Math.round(serviceReviews.reduce((s, r) => s + r.rating, 0) / serviceReviews.length * 10) / 10
      : 0;
    res.json({ reviews: serviceReviews, average_rating: avg, count: serviceReviews.length });
  });
  
  // === BOOKING STATUS LIFECYCLE (Admin) ===
  // Confirm booking (admin action)
  app.put('/api/public/bookings/:id/confirm', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    // Client can confirm their own booking, or admin via internal API
    if (booking.client_id !== req.client.id) {
      // Allow admin override later via internal API
      return res.status(403).json({ error: 'Forbidden — use admin dashboard' });
    }
    if (booking.status === 'cancelled' || booking.status === 'Dibatalkan') {
      return res.status(400).json({ error: 'Booking sudah dibatalkan' });
    }
    saveOne('appointments', req.params.id, { status: 'Confirmed', confirmed_at: new Date().toISOString() });
    res.json({ status: 'Confirmed', message: 'Booking dikonfirmasi. Sampai jumpa! ✨' });
  });

  // Mark booking as completed (after service done)
  app.put('/api/public/bookings/:id/complete', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) {
      return res.status(403).json({ error: 'Forbidden — use admin dashboard' });
    }
    if (booking.status === 'cancelled' || booking.status === 'Dibatalkan') {
      return res.status(400).json({ error: 'Booking sudah dibatalkan' });
    }
    saveOne('appointments', req.params.id, {
      status: 'Selesai',
      completed_at: new Date().toISOString()
    });

    // === AUTO LOYALTY POINTS ===
    const pointsResult = awardBookingPoints(booking, req.client.id);

    res.json({
      status: 'Selesai',
      message: `Sesi selesai! ✨ ${pointsResult.earned > 0 ? `Anda mendapat ${pointsResult.earned} poin!` : ''} Jangan lupa kasih ulasan ya, Ma~`,
      points: pointsResult
    });
  });

  // === LOYALTY REDEMPTION (disabled) ===
  // (replaced by use-discount below)
  
  // === USE TIER DISCOUNT (1x, resets spendable points to 0) ===
  app.post('/api/public/loyalty/use-discount', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const mem = getMembership(req.client.id);
    if (!mem.discount_available) {
      return res.status(400).json({ 
        error: `Poin belum cukup untuk diskon ${mem.tier_label}. Butuh ${getTierConfig(mem.tier).min} poin, saat ini: ${mem.points_balance} poin.`,
        tier: mem.tier_label,
        required: getTierConfig(mem.tier).min,
        current: mem.points_balance
      });
    }
    
    const discountPct = mem.discount_percent;
    
    // Record redeem transaction
    const txId = uid();
    const now = new Date();
    saveOne('points_tx', txId, {
      id: txId, client_id: req.client.id, type: 'redeem',
      amount: mem.points_balance, source: 'discount',
      description: `Pakai diskon ${mem.tier_label} ${discountPct}% (${mem.points_balance} pts)`,
      created_at: now.toISOString(), expires_at: null
    });
    
    // Reset spendable points to 0
    saveOne('clients', req.client.id, {
      loyalty_points: 0,
      last_discount: { tier: mem.tier_label, discount: discountPct, used_at: now.toISOString() }
    });
    
    res.json({
      success: true,
      discount_percent: discountPct,
      tier: mem.tier_label,
      message: `Diskon ${mem.tier_label} ${discountPct}% siap digunakan! Poin direset ke 0. Kumpulkan ${getTierConfig(mem.tier).min} poin lagi untuk diskon berikutnya.`
    });
  });

  // === WHATSAPP NOTIFICATION (via Fonnte / Wablas) ===
  const WA_API_URL = process.env.WA_API_URL || '';
  const WA_API_KEY = process.env.WA_API_KEY || '';

  async function sendWA(phone, message) {
    if (!WA_API_URL || !WA_API_KEY || !phone) return;
    try {
      const https = require('https');
      const url = new URL(WA_API_URL);
      const payload = JSON.stringify({ target: phone.replace(/\D/g, ''), message });
      await new Promise((resolve) => {
        const r = https.request({
          hostname: url.hostname,
          path: url.pathname,
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': WA_API_KEY
          }
        }, (resp) => { resp.on('end', resolve); resp.resume(); });
        r.on('error', resolve);
        r.write(payload);
        r.end();
      });
    } catch {}
  }

  // === MORNING REMINDER: sends WA to all clients with bookings today ===
  app.post('/api/cron/morning-reminders', (req, res) => {
    const today = new Date().toLocaleDateString('en-CA'); // YYYY-MM-DD in local time
    const apps = readAll('appointments').filter(a => 
      a.date === today && 
      a.status !== 'cancelled' && a.status !== 'Dibatalkan' &&
      a.wa && a.wa.trim()
    );
    
    if (!apps.length) return res.json({ sent: 0, message: 'No bookings today' });
    
    // Deduplicate by WA number (one notification per client per day)
    const seen = new Set();
    const unique = [];
    apps.forEach(a => {
      const wa = a.wa.replace(/\D/g, '');
      if (!seen.has(wa)) {
        seen.add(wa);
        unique.push(a);
      }
    });
    
    let sent = 0;
    const results = [];
    
    unique.forEach(a => {
      // Count all bookings for this client today
      const clientApps = apps.filter(x => x.wa.replace(/\D/g, '') === a.wa.replace(/\D/g, ''));
      const scheduleLines = clientApps.map(x => `• ${x.time} — ${x.service} (${x.child ? 'Anak: ' + x.child : 'Ibu'})`).join('\n');
      
      const msg = `🌅 *Selamat Pagi!*\n\nJadwal Anda hari ini di Lelap Mom Baby Care:\n${scheduleLines}\n\n📍 Jl. Taman Pahlawan, Salatiga\n📞 Info: 081-313-99-636\n\n_Sampai jumpa! 💚_`;
      
      sendWA(a.wa, msg);
      sent++;
      results.push({ wa: a.wa.replace(/\D/g, '').slice(-4), name: a.client_name, bookings: clientApps.length });
    });
    
    res.json({ sent, total_bookings: apps.length, clients: results });
  });

  // ── Auto-cancel past Menunggu bookings ──
  app.post('/api/cron/auto-cancel', (req, res) => {
    const nowWIB = new Date(new Date().toLocaleString('en-US', { timeZone: 'Asia/Jakarta' }));
    const today = nowWIB.toISOString().substring(0, 10);
    const hourWIB = nowWIB.getHours();
    const minWIB = nowWIB.getMinutes();
    const nowMin = hourWIB * 60 + minWIB;
    
    const apps = readAll('appointments').filter(a => a.status === 'Menunggu');
    let cancelled = 0;
    
    apps.forEach(a => {
      if (!a.date || !a.time) return;
      const [sh, sm] = a.time.split(':').map(Number);
      const slotMin = sh * 60 + sm;
      
      // Cancel if date is past OR today but time has passed
      if (a.date < today || (a.date === today && slotMin < nowMin)) {
        saveOne('appointments', a.id, { 
          status: 'Dibatalkan',
          cancelled_at: new Date().toISOString(),
          cancel_reason: 'Auto: tidak dikonfirmasi sebelum jadwal'
        });
        cancelled++;
      }
    });
    
    res.json({ cancelled, message: cancelled > 0 ? `${cancelled} booking otomatis dibatalkan` : 'No bookings to cancel' });
  });

  console.log('✅ Backend complete: payment, status lifecycle, loyalty, WA ready');
  app.get('/api/public/loyalty', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    res.json({ points: client.loyalty_points || 0 });
  });
  
  // === CHAT BOT ===
  app.get('/api/public/chat', publicAuth, (req, res) => {
    const { message } = req.query;
    if (!message) return res.status(400).json({ error: 'message required', code: 400 });
    const msg = message.toLowerCase().trim();
    // FAQ pattern matching
    const faq = {
      'jam buka': 'Lelap buka setiap hari pukul 08:00 - 20:00 WIB, Ma. Hari Minggu & libur nasional tetap buka ya!',
      'alamat': 'Lelap Mom Baby Care berlokasi di Jl. Diponegoro No. 10, Salatiga. Dekat Alun-Alun Pancasila, Ma.',
      'parkir': 'Tersedia parkir gratis untuk mobil & motor di halaman Lelap, Ma. Aman dan luas.',
      'newborn': 'Kami melayani baby massage untuk newborn mulai usia 2 minggu ya, Ma. Terapis kami tersertifikasi.',
      'harga': 'Harga layanan bervariasi mulai Rp75.000 - Rp350.000. Bisa lihat lengkap di menu Layanan setelah login, Ma.',
      'pembayaran': 'Kami terima pembayaran via QRIS, transfer bank, dan tunai di tempat. Bisa DP 50% dulu, Ma.',
      'metode pembayaran': 'Tersedia QRIS, transfer bank BCA/Mandiri, dan bayar di tempat. Fleksibel, Ma!',
      'reschedule': 'Bisa reschedule via aplikasi maksimal H-1 sebelum jadwal. Untuk H yang sama, hubungi admin ya, Ma.',
      'cancel': 'Cancel gratis maksimal H-1. Kurang dari itu, hubungi admin kami untuk bantuan, Ma.',
      'testimoni': 'Banyak Mama yang puas dengan layanan kami! Rating rata-rata 4.8 dari 500+ review. Bisa lihat di menu Review ya.',
      'bpjs': 'Saat ini kami belum menerima BPJS, Ma. Tapi harga kami terjangkau dan bisa dicicil.',
      'diskon': 'Kami punya program loyalitas: setiap 100 poin dapat diskon 10%. Poin didapat dari setiap booking, Ma!',
      'kontak': 'Butuh bantuan? Hubungi admin kami via WhatsApp di 0812-3456-7890 atau chat di sini ya, Ma.',
    };
    // Check for keyword match
    for (const [keyword, answer] of Object.entries(faq)) {
      if (msg.includes(keyword)) {
        return res.json({ answer, type: 'faq', keyword });
      }
    }
    // No match — escalate
    res.json({ 
      answer: 'Pertanyaan Mama sudah kami catat. Admin kami akan segera merespon ya, Ma. Untuk pertanyaan mendesak, silakan hubungi WhatsApp 0812-3456-7890.', 
      type: 'escalated',
      ticket_id: 'faq_' + Date.now()
    });
  });

  app.post('/api/public/chat', publicAuth, (req, res) => {
    const { message } = req.body;
    const msg = (message || '').toLowerCase().trim();
    
    // Slot intent detection
    if (hasSlotIntent(msg)) {
      const summary = getSlotSummary();
      return res.json({ reply: '📊 *Cek Slot Jadwal*\n' + summary + '\n\nBooking langsung via aplikasi ya, Ma~ 💚' });
    }
    
    const faq = {
      'jam buka': 'Lelap buka setiap hari, jam 08:00 - 20:00 WIB, Ma~ 🕐',
      'buka': 'Lelap buka setiap hari, jam 08:00 - 20:00 WIB, Ma~ 🕐',
      'hari': (() => { const d=new Date(); const dn=d.toLocaleDateString('id-ID',{weekday:'long',timeZone:'Asia/Jakarta'}); const ds=d.toLocaleDateString('id-ID',{year:'numeric',month:'long',day:'numeric',timeZone:'Asia/Jakarta'}); return dn==='Senin' ? `Hari ini ${dn}, ${ds}. ⚠️ TUTUP (LIBUR). Hanya Prenatal Class 16.00-17.00.` : `Hari ini ${dn}, ${ds}. ✅ BUKA 08.00-16.00.`; })(),
      'alamat': 'Jl. Taman Pahlawan Salatiga. Ada di Google Maps juga! 📍',
      'lokasi': 'Jl. Taman Pahlawan Salatiga. Ada di Google Maps juga! 📍',
      'parkir': 'Parkir luas dan gratis, Ma~ 🚗',
      'bayi': 'Lelap spesialis perawatan bayi! Ada pijat relaksasi, renang air hangat, gym bayi, dan masih banyak lagi. Cek menu layanan ya 👶',
      'hamil': 'Ada PRENATAL YOGA setiap Senin jam 16:00 khusus untuk Mama hamil, diajar langsung oleh Owner kami 🧘‍♀️',
      'harga': 'Harga mulai dari Rp 20.000 untuk gym bayi, sampai Rp 150.000 untuk paket lengkap. Cek menu layanan untuk detailnya 💰',
      'cancel': 'Mama bisa batalkan booking sendiri di app jika masih H-1 ya. Kalau sudah dekat jadwal, hubungi admin kami.',
      'reschedule': 'Reschedule bisa dilakukan sendiri lewat app maksimal H-1 sebelum jadwal.',
    };
    
    let response = null;
    for (const [key, value] of Object.entries(faq)) {
      if (msg.includes(key)) {
        response = value;
        break;
      }
    }
    
    if (!response) {
      response = 'Maaf, Kak Lela belum ngerti pertanyaan Mama nih 🥺 Coba tanya tentang jam buka, alamat, layanan, harga, atau cara cancel/reschedule ya~ Atau nanti admin kami yang bantu jawab!';
    }
    
    res.json({ reply: response });
  });
  
  // === PRICE (NO WEEKEND SURCHARGE) ===
  function getPrice(service, date) {
    return service.price || 0;
  }

  // === MIDTRANS SNAP ===
  const MIDTRANS_SERVER_KEY = process.env.MIDTRANS_SERVER_KEY || '';
  const MIDTRANS_IS_PRODUCTION = process.env.MIDTRANS_IS_PRODUCTION === 'true';

  app.post('/api/public/payment/snap', publicAuth, async (req, res) => {
    const { booking_id } = req.body;
    if (!booking_id) return res.status(400).json({ error: 'booking_id required' });

    const booking = getOne('appointments', booking_id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });

    const client = getOne('clients', req.client.id);
    const amount = booking.amount || 0;

    if (!MIDTRANS_SERVER_KEY) {
      // No Midtrans key — fallback: mark as "Menunggu Pembayaran"
      saveOne('appointments', booking_id, { status: 'Menunggu Pembayaran', payment_method: 'transfer' });
      return res.json({
        snap_token: null,
        fallback: true,
        message: 'Silakan transfer ke rekening Lelap. Admin akan konfirmasi.',
        booking_code: booking.booking_code,
        amount
      });
    }

    try {
      const https = require('https');
      const midtransReq = {
        transaction_details: {
          order_id: booking.booking_code,
          gross_amount: amount
        },
        customer_details: {
          first_name: client.name || 'Pelanggan',
          email: client.email || '',
          phone: client.phone || ''
        },
        item_details: [{
          id: booking.kode || booking_id,
          price: amount,
          quantity: 1,
          name: booking.service || 'Layanan Lelap'
        }]
      };

      const snapResp = await new Promise((resolve, reject) => {
        const payload = JSON.stringify(midtransReq);
        const options = {
          hostname: MIDTRANS_IS_PRODUCTION ? 'app.midtrans.com' : 'app.sandbox.midtrans.com',
          path: '/snap/v1/transactions',
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Basic ' + Buffer.from(MIDTRANS_SERVER_KEY + ':').toString('base64'),
            'Content-Length': Buffer.byteLength(payload)
          }
        };
        const r = https.request(options, (resp) => {
          let data = '';
          resp.on('data', chunk => data += chunk);
          resp.on('end', () => {
            try { resolve(JSON.parse(data)); } catch { resolve(null); }
          });
        });
        r.on('error', () => resolve(null));
        r.write(payload);
        r.end();
      });

      if (snapResp && snapResp.token) {
        saveOne('appointments', booking_id, { payment_method: 'midtrans', snap_token: snapResp.token });
        return res.json({ snap_token: snapResp.token, redirect_url: snapResp.redirect_url });
      }

      saveOne('appointments', booking_id, { status: 'Menunggu Pembayaran', payment_method: 'transfer' });
      res.json({ snap_token: null, fallback: true, message: 'Gagal membuat pembayaran online. Silakan bayar di tempat.', booking_code: booking.booking_code, amount });
    } catch (e) {
      saveOne('appointments', booking_id, { status: 'Menunggu Pembayaran', payment_method: 'transfer' });
      res.json({ snap_token: null, fallback: true, message: 'Pembayaran online error: ' + e.message, booking_code: booking.booking_code, amount });
    }
  });
  app.post('/api/public/payment/callback', (req, res) => {
    const { order_id, transaction_status, fraud_status } = req.body;
    // Validate required fields
    if (!order_id) return res.status(400).json({ error: 'order_id required', code: 400 });
    if (!transaction_status) return res.status(400).json({ error: 'transaction_status required', code: 400 });
    console.log('Payment callback:', req.body);
    
    // Find booking by order_id (booking_code)
    const bookings = readAll('appointments');
    const booking = bookings.find(b => b.booking_code === order_id);
    
    if (booking) {
      let status = booking.status;
      if (transaction_status === 'capture' || transaction_status === 'settlement') {
        status = 'Confirmed';
      } else if (transaction_status === 'pending') {
        status = 'Menunggu Pembayaran';
      } else if (transaction_status === 'deny' || transaction_status === 'expire' || transaction_status === 'cancel') {
        status = 'cancelled';
      }
      saveOne('appointments', booking.id, { status, payment_status: transaction_status });
    }
    
    res.json({ status: 'OK' });
  });
  
  // === SETTINGS (PUBLIC - LIMITED) ===
  app.get('/api/public/settings', (req, res) => {
    const settings = getOne('settings', 'settings');
    // Only return safe settings
    res.json({
      spa_name: settings?.spa_name || 'Lelap Mom Baby Care Salatiga',
      address: settings?.address || 'Jl Taman Pahlawan Salatiga',
      tagline: settings?.tagline || '',
      open_time: settings?.open_time || '08:00',
      close_time: settings?.close_time || '20:00',
      whatsapp: settings?.whatsapp || ''
    });
  });
  
  // === HOMECARE & TRANSPORT ===
  // ── Transport rates (loaded from data/transport-rates.json) ──
  function loadTransportRates() {
    try {
      const raw = fs.readFileSync(path.join(DATA_DIR, 'transport-rates.json'), 'utf8');
      return JSON.parse(raw);
    } catch (e) {
      console.error('Failed to load transport-rates.json:', e.message);
      return { rates: [{max_km:5,price:15000},{max_km:7,price:25000},{max_km:10,price:35000},{max_km:15,price:50000}], max_radius_km: 20, approval_threshold_km: 15 };
    }
  }

  function getTransportPrice(distanceKm) {
    const cfg = loadTransportRates();
    const rates = cfg.rates || [];
    const threshold = cfg.approval_threshold_km || 15;
    for (const tier of rates) {
      if (distanceKm <= tier.max_km) return { price: tier.price, tier: `0-${tier.max_km} km`, needs_approval: false };
    }
    return { price: 0, tier: '>' + threshold + ' km', needs_approval: true };
  }

  function getMaxRadius() {
    const cfg = loadTransportRates();
    return cfg.max_radius_km || 20;
  }

  // Geocoding: address -> coordinates (Nominatim)
  app.get('/api/public/geocode', publicAuth, async (req, res) => {
    const { address } = req.query;
    if (!address) return res.status(400).json({ error: 'address required' });
    try {
      const https = require('https');
      const result = await new Promise((resolve) => {
        const url = `https://nominatim.openstreetmap.org/search?format=json&q=${encodeURIComponent(address)}&limit=1`;
        https.get(url, { headers: { 'User-Agent': 'LelapBookingCare/1.0' } }, (resp) => {
          let data = '';
          resp.on('data', chunk => data += chunk);
          resp.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve([]); } });
        }).on('error', () => resolve([]));
      });
      if (result.length > 0) {
        res.json({ lat: parseFloat(result[0].lat), lng: parseFloat(result[0].lon), display_name: result[0].display_name });
      } else {
        res.json({ error: 'Alamat tidak ditemukan' });
      }
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  // Reverse geocoding: coordinates -> address (Nominatim)
  app.get('/api/public/reverse-geocode', publicAuth, async (req, res) => {
    const { lat, lng } = req.query;
    if (!lat || !lng) return res.status(400).json({ error: 'lat and lng required' });
    try {
      const https = require('https');
      const result = await new Promise((resolve) => {
        const url = `https://nominatim.openstreetmap.org/reverse?format=json&lat=${lat}&lon=${lng}`;
        https.get(url, { headers: { 'User-Agent': 'LelapBookingCare/1.0' } }, (resp) => {
          let data = '';
          resp.on('data', chunk => data += chunk);
          resp.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve(null); } });
        }).on('error', () => resolve(null));
      });
      if (result && result.display_name) {
        res.json({ display_name: result.display_name, address: result.address });
      } else {
        res.json({ error: 'Lokasi tidak ditemukan' });
      }
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  // Road distance calculation (OSRM)
  app.get('/api/public/distance', publicAuth, async (req, res) => {
    const { from_lat, from_lng, to_lat, to_lng } = req.query;
    if (!from_lat || !from_lng || !to_lat || !to_lng) {
      return res.status(400).json({ error: 'from_lat, from_lng, to_lat, to_lng required' });
    }
    try {
      const https = require('https');
      const result = await new Promise((resolve) => {
        const url = `https://router.project-osrm.org/route/v1/driving/${from_lng},${from_lat};${to_lng},${to_lat}?overview=false`;
        const req = https.get(url, { timeout: 6000 }, (resp) => {
          let data = '';
          resp.on('data', chunk => data += chunk);
          resp.on('end', () => { try { resolve(JSON.parse(data)); } catch { resolve(null); } });
        });
        req.on('error', () => resolve(null));
        req.on('timeout', () => { req.destroy(); resolve(null); });
      });
      if (result && result.routes && result.routes.length > 0) {
        const roadKm = result.routes[0].distance / 1000;
        const distanceKm = Math.ceil(roadKm * 1.15 * 10) / 10; // +15% safety margin
        const transport = getTransportPrice(distanceKm);
        res.json({
          distance_km: distanceKm,
          distance_meters: Math.round(roadKm * 1000),
          transport_price: transport.price,
          transport_tier: transport.tier,
          needs_approval: transport.needs_approval,
          note: 'Jarak jalan + margin aman 15%',
          rates: (loadTransportRates().rates || []),
          max_radius_km: getMaxRadius()
        });
      } else {
        // Fallback: Haversine with conservative multiplier
        const straightDist = haversineDistance(parseFloat(from_lat), parseFloat(from_lng), parseFloat(to_lat), parseFloat(to_lng));
        const roadEstimate = Math.ceil(straightDist * 1.6 * 10) / 10; // 60% extra for roads
        const transport = getTransportPrice(roadEstimate);
        res.json({
          distance_km: roadEstimate,
          distance_meters: Math.round(roadEstimate * 1000),
          transport_price: transport.price,
          transport_tier: transport.tier,
          needs_approval: transport.needs_approval,
          fallback: true,
          note: 'Estimasi jarak + margin aman (rute jalan tidak tersedia)',
          rates: (loadTransportRates().rates || []),
          max_radius_km: getMaxRadius()
        });
      }
    } catch(e) { res.status(500).json({ error: e.message }); }
  });

  // Clinic location (public)
  app.get('/api/public/clinic-location', (req, res) => {
    const settings = getOne('settings', 'settings') || {};
    res.json({
      lat: settings.clinic_lat || -7.3326,
      lng: settings.clinic_lng || 110.5069,
      address: settings.address || 'Jl Taman Pahlawan Salatiga'
    });
  });

  // === TRANSPORT APPROVAL (Owner/SA) ===
  app.get('/api/public/transport-approvals', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    const role = (client || {}).role || 'client';
    const all = readAll('appointments');
    let pending;
    if (role === 'owner' || role === 'admin' || role === 'sa') {
      pending = all.filter(a => a.transport_status === 'pending_approval')
        .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
    } else {
      pending = all.filter(a => a.client_id === req.client.id && a.transport_status === 'pending_approval')
        .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
    }
    res.json(pending);
  });

  // SA sets transport price
  app.put('/api/public/transport-approvals/:id/price', publicAuth, (req, res) => {
    const tok = verifyToken(req);
    if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
    const user = getOne('users', tok.id);
    const role = (user.role || '').toLowerCase();
    if (role !== 'owner' && role !== 'admin' && role !== 'sa') return res.status(403).json({ error: 'Forbidden' });

    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.transport_status !== 'pending_approval') return res.status(400).json({ error: 'Not pending approval' });

    const { transport_price } = req.body;
    if (!transport_price || transport_price < 0) return res.status(400).json({ error: 'transport_price required' });

    saveOne('appointments', req.params.id, {
      transport_price: parseInt(transport_price),
      transport_status: 'price_proposed',
      transport_updated_by: user.name || user.email,
      transport_updated_at: new Date().toISOString()
    });
    res.json({ status: 'price_proposed', transport_price: parseInt(transport_price) });
  });

  // Client approves/rejects transport price
  app.put('/api/public/transport-approvals/:id/respond', publicAuth, (req, res) => {
    const booking = getOne('appointments', req.params.id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });

    const { action } = req.body; // 'approve' or 'reject'
    if (action === 'approve') {
      saveOne('appointments', req.params.id, {
        transport_status: 'approved',
        transport_confirmed_at: new Date().toISOString()
      });
      res.json({ status: 'approved', message: 'Harga transport disetujui. Booking dilanjutkan.' });
    } else if (action === 'reject') {
      saveOne('appointments', req.params.id, {
        transport_status: 'rejected',
        transport_rejected_at: new Date().toISOString()
      });
      res.json({ status: 'rejected', message: 'Harga transport ditolak. Silakan hubungi admin.' });
    } else {
      return res.status(400).json({ error: 'action must be approve or reject' });
    }
  });

  // ═══════════════════════════════════════════════════════════
  // MEMBERSHIP & LOYALTY API
  // ═══════════════════════════════════════════════════════════

  // Get membership status
  app.get('/api/public/membership', publicAuth, (req, res) => {
    // Auto-heal: sync client.loyalty_points with points_tx to prevent divergence
    const client = getOne('clients', req.client.id);
    if (client) {
      const tx = readAll('points_tx').filter(t => t.client_id === req.client.id);
      const txBalance = tx.reduce((sum, t) => {
        if (t.type === 'earn') return sum + (t.amount || 0);
        if (t.type === 'redeem') return sum - (t.amount || 0);
        return sum;
      }, 0);
      if (client.loyalty_points !== txBalance) {
        saveOne('clients', req.client.id, { loyalty_points: txBalance });
      }
    }
    
    const mem = getMembership(req.client.id);
    
    // Auto-expire old points on read
    if (Math.random() < 0.1) expireOldPoints(); // 10% chance to trigger cleanup
    
    res.json(mem);
  });

  // Get points transaction history
  app.get('/api/public/points/history', publicAuth, (req, res) => {
    const tx = readAll('points_tx')
      .filter(t => t.client_id === req.client.id)
      .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
    res.json(tx);
  });

  // Redeem points for discount
  app.post('/api/public/points/redeem', publicAuth, (req, res) => {
    const { points } = req.body;
    if (!points || points < 1) return res.status(400).json({ error: 'Jumlah poin harus minimal 1' });
    
    const result = redeemPoints(req.client.id, parseInt(points));
    if (result.error) return res.status(400).json(result);
    
    const mem = getMembership(req.client.id);
    res.json({
      ...result,
      tier: mem.tier_label,
      points_balance: mem.points_balance
    });
  });

  // Claim birthday voucher (10% discount, 1x/year/account)
  app.post('/api/public/birthday-bonus', publicAuth, (req, res) => {
    const mem = getMembership(req.client.id);
    if (!mem.birthday_bonus_available) {
      return res.status(400).json({ error: 'Voucher ulang tahun sudah diklaim tahun ini atau belum waktunya. 🎂' });
    }
    
    // Record voucher claim marker (amount: 0, no points earned)
    const id = uid();
    saveOne('points_tx', id, {
      id, client_id: req.client.id, type: 'welcome', amount: 0,
      source: 'birthday', description: 'Voucher ultah diskon 10% (1x/tahun)',
      created_at: new Date().toISOString(), expires_at: null
    });
    
    res.json({
      success: true,
      message: 'Voucher ultah siap digunakan! Diskon 10% untuk booking berikutnya. 🎂',
      voucher_type: 'birthday',
      discount_percent: 10
    });
  });

  // Welcome bonus info (10% discount for non-tier, one-time)
  app.get('/api/public/welcome-bonus', publicAuth, (req, res) => {
    const mem = getMembership(req.client.id);
    res.json({
      available: mem.welcome_bonus_available,
      discount_percent: 10,
      message: mem.welcome_bonus_available 
        ? 'Selamat datang! 🎉 Anda mendapat diskon 10% untuk booking pertama.' 
        : 'Welcome bonus sudah digunakan atau tier Anda sudah naik.'
    });
  });

  // Claim welcome bonus (marks as used, doesn't auto-apply — applied during first booking)
  app.post('/api/public/welcome-bonus/claim', publicAuth, (req, res) => {
    const mem = getMembership(req.client.id);
    if (!mem.welcome_bonus_available) {
      return res.status(400).json({ error: 'Welcome bonus sudah diklaim atau tidak tersedia.' });
    }
    
    // Record welcome bonus as a points transaction (non-expiring, non-redeemable marker)
    const id = uid();
    saveOne('points_tx', id, {
      id, client_id: req.client.id, type: 'welcome', amount: 0,
      source: 'welcome', description: 'Welcome bonus diskon 10% (sekali pakai)',
      created_at: new Date().toISOString(), expires_at: null
    });
    
    res.json({
      success: true,
      discount_percent: 10,
      message: 'Welcome bonus siap digunakan! Diskon 10% otomatis terpakai di booking pertama Anda. 🎉'
    });
  });

  // ═══════════════════════════════════════════════════════════
  // AI CONSULTATION API
  // ═══════════════════════════════════════════════════════════

  app.post('/api/public/consultation', publicAuth, async (req, res) => {
    const { question, category } = req.body;
    if (!question || question.trim().length < 2) {
      return res.status(400).json({ error: 'Pertanyaan terlalu singkat. Silakan ketik lebih detail.' });
    }
    
    let result;
    const q = question.trim();
    
    // Medical questions → reject immediately
    if (isMedicalQuestion(q)) {
      result = {
        type: 'rejected',
        message: '⚠️ Maaf, saya tidak bisa menjawab pertanyaan medis. Untuk konsultasi kesehatan, silakan:\n\n✅ Hubungi bidan Lelap langsung via WhatsApp\n✅ Booking sesi konsultasi bidan di aplikasi (menu Layanan)\n✅ Kunjungi klinik untuk pemeriksaan langsung\n\nSaya hanya bisa membantu pertanyaan seputar layanan, jam buka, harga, dan informasi klinik ya, Ma~ 😊'
      };
    } else if (isRecommendationQuery(q)) {
      // Recommendation/age queries → use DB directly (no AI, no truncation)
      result = answerRecommendation(q);
    } else {
      // Detect slot inquiry → inject real-time data
      let context = '';
      if (hasSlotIntent(q)) {
        context = '\n===== DATA SLOT REAL-TIME =====\n' + getSlotSummary() + '\nGunakan data ini untuk menjawab pertanyaan slot/jadwal.\n';
      }
      // Try AI first, fallback to template
      const aiAnswer = await askAI(q, context);
      if (aiAnswer) {
        result = { type: 'ai', message: aiAnswer };
      } else if (hasSlotIntent(q)) {
        // AI failed on slot query → use slot summary directly
        const summary = getSlotSummary();
        result = { type: 'general', message: '📊 *Cek Slot Jadwal*\n' + summary + '\n\nBooking langsung via aplikasi ya, Ma~ 💚' };
      } else {
        result = answerFAQ(q);
      }
    }
    
    // Save to consultation log
    const consultation = {
      id: uid(),
      client_id: req.client.id,
      question: question.trim(),
      answer_type: result.type,
      answer: result.message,
      category: category || 'general',
      created_at: new Date().toISOString()
    };
    saveOne('consultations', consultation.id, consultation);
    
    res.json({
      id: consultation.id,
      type: result.type,
      answer: result.message,
      created_at: consultation.created_at
    });
  });

  // Get consultation history
  app.get('/api/public/consultation', publicAuth, (req, res) => {
    const all = readAll('consultations')
      .filter(c => c.client_id === req.client.id)
      .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''))
      .slice(0, 50);
    res.json(all);
  });

  // === MEDICAL CONSULTATION CHAT (Bronze+ only) ===
  app.post('/api/public/consultation/medical', publicAuth, (req, res) => {
    const { message } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Pesan tidak boleh kosong' });
    
    // Bronze+ gate
    const mem = getMembership(req.client.id);
    if (mem.tier === 'non-tier') {
      return res.status(403).json({ error: 'Konsultasi medis hanya tersedia untuk member Bronze ke atas. Upgrade membership Anda terlebih dahulu.', tier: mem.tier_label });
    }
    
    const client = getOne('clients', req.client.id);
    const chat = saveOne('medical_chats', uid(), {
      client_id: req.client.id,
      client_name: client?.name || 'Client',
      client_photo: client?.photo || null,
      message: message.trim(),
      sender: 'client',
      admin_name: null,
      created_at: new Date().toISOString(),
      read: false
    });
    
    res.status(201).json(chat);
  });
  
  app.get('/api/public/consultation/medical', publicAuth, (req, res) => {
    // Bronze+ gate
    const mem = getMembership(req.client.id);
    if (mem.tier === 'non-tier') {
      return res.status(403).json({ error: 'Konsultasi medis hanya tersedia untuk member Bronze ke atas.', tier: mem.tier_label });
    }
    
    const chats = readAll('medical_chats')
      .filter(c => c.client_id === req.client.id)
      .sort((a, b) => (a.created_at || '').localeCompare(b.created_at || ''));
    
    // Mark admin messages as read
    chats.forEach(c => { if (c.sender === 'admin' && !c.read) {
      saveOne('medical_chats', c.id, { read: true });
    }});
    
    res.json({ chats, tier: mem.tier_label, tier_discount: mem.discount_pct });
  });

  // === ADMIN: Medical Chat Dashboard ===
  app.get('/api/admin/medical-chats', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const all = readAll('medical_chats');
    // Group by client_id, get latest message per client
    const clientMap = {};
    all.forEach(c => {
      if (!clientMap[c.client_id] || c.created_at > clientMap[c.client_id].created_at) {
        clientMap[c.client_id] = c;
      }
    });
    
    // Count unread per client
    const clients = Object.values(clientMap).map(c => {
      const client = getOne('clients', c.client_id);
      const clientName = c.client_name || client?.name || 'Unknown';
      return {
        client_id: c.client_id,
        client_name: clientName,
        client_photo: c.client_photo || client?.photo || null,
        last_message: c.message?.substring(0, 80),
        last_sender: c.sender,
        last_time: c.created_at,
        unread: all.filter(m => m.client_id === c.client_id && m.sender === 'client' && !m.read).length
      };
    }).sort((a, b) => (b.last_time || '').localeCompare(a.last_time || ''));
    
    res.json(clients);
  });
  
  app.get('/api/admin/medical-chats/:client_id', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const chats = readAll('medical_chats')
      .filter(c => c.client_id === req.params.client_id)
      .sort((a, b) => (a.created_at || '').localeCompare(b.created_at || ''));
    
    // Mark client messages as read
    let marked = 0;
    chats.forEach(c => { if (c.sender === 'client' && !c.read) {
      saveOne('medical_chats', c.id, { read: true });
      marked++;
    }});
    
    res.json({ chats, marked });
  });
  
  app.post('/api/admin/medical-chats/:client_id', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const { message, admin_name } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Pesan tidak boleh kosong' });
    
    const chat = saveOne('medical_chats', uid(), {
      client_id: req.params.client_id,
      client_name: null,
      message: message.trim(),
      sender: 'admin',
      admin_name: admin_name || 'Bidan Lelap',
      created_at: new Date().toISOString(),
      read: false
    });
    
    res.status(201).json(chat);
  });

  // === BOOKING CHAT (Client ↔ Admin tentang masalah booking) ===
  
  // Client: get all their chat messages
  app.get('/api/public/booking-chats', publicAuth, (req, res) => {
    const all = readAll('booking_chats');
    const messages = all.filter(c => c.client_id === req.client.id)
      .sort((a, b) => (a.created_at || '').localeCompare(b.created_at || ''));
    
    messages.forEach(c => { if (c.sender === 'admin' && !c.read) {
      saveOne('booking_chats', c.id, { read: true });
    }});
    
    res.json({ messages, room_id: req.client.id });
  });

  app.post('/api/public/booking-chats', publicAuth, (req, res) => {
    const { message } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Pesan tidak boleh kosong' });
    
    const { booking_id } = req.body;
    const client = getOne('clients', req.client.id);
    const chat = saveOne('booking_chats', uid(), {
      client_id: req.client.id,
      client_name: client?.name || req.client.name || 'Client',
      message: message.trim(),
      sender: 'client',
      created_at: new Date().toISOString(),
      read: false,
      ...(booking_id ? { booking_id } : {})
    });
    
    // Auto-reply untuk chat PERTAMA dari client
    const allClientChats = readAll('booking_chats').filter(c => c.client_id === req.client.id && c.sender === 'client');
    if (allClientChats.length === 1) {
      // Cari nama mama: client.name → profil type=mama → email
      let mamaName = client?.name || req.client.name || null;
      if (!mamaName) {
        const profiles = readAll('profiles').filter(p => p.client_id === req.client.id && (p.type || p.relation || '').toLowerCase() === 'mama');
        mamaName = profiles[0]?.name || null;
      }
      if (!mamaName) {
        mamaName = req.client.email || client?.email || '';
      }
      
      let greeting = mamaName ? `Hi mama ${mamaName}, kami akan hubungi sesaat lagi, terimakasih` : 'Hi mama, kami akan hubungi sesaat lagi, terimakasih';
      
      // Jika ada booking_id, tambahkan info booking
      const { booking_id } = req.body;
      if (booking_id) {
        const booking = getOne('appointments', booking_id) || getOne('bookings', booking_id);
        if (booking) {
          const bService = booking.service || booking.service_name || '-';
          const bDate = booking.date || '-';
          const bTime = booking.time || '-';
          const bCode = booking.booking_code || booking.id || '-';
          greeting = mamaName
            ? `Hi mama ${mamaName}, kami akan hubungi sesaat lagi.\\n\\n📋 Info Booking:\\nKode: ${bCode}\\nLayanan: ${bService}\\nTanggal: ${bDate} ${bTime}\\n\\nTerimakasih`
            : `Hi mama, kami akan hubungi sesaat lagi.\\n\\n📋 Info Booking:\\nKode: ${bCode}\\nLayanan: ${bService}\\nTanggal: ${bDate} ${bTime}\\n\\nTerimakasih`;
        }
      }
      
      saveOne('booking_chats', uid(), {
        client_id: req.client.id,
        client_name: client?.name || req.client.name || 'Client',
        message: greeting,
        sender: 'admin',
        admin_name: 'Lelap Care',
        created_at: new Date().toISOString(),
        read: false
      });
    }
    
    res.status(201).json(chat);
  });

  // === ADMIN: Booking Chat ===
  app.get('/api/admin/booking-chats', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const all = readAll('booking_chats');
    const search = (req.query.search || '').toLowerCase().trim();
    const allProfiles = readAll('profiles');
    const allClients = readAll('clients');
    
    const clientMap = {};
    all.forEach(c => {
      if (!clientMap[c.client_id] || c.created_at > clientMap[c.client_id].created_at) {
        clientMap[c.client_id] = c;
      }
    });
    
    const rooms = Object.values(clientMap).map(c => {
      // Cari nama mama dari profil
      const mamaProfile = allProfiles.find(p => p.client_id === c.client_id && (p.type || p.relation || '').toLowerCase() === 'mama');
      const mamaName = mamaProfile?.name || '';
      const client = allClients.find(cl => cl.id === c.client_id);
      const clientName = c.client_name || client?.name || 'Client';

      // Cari info booking dari chat record (gunakan booking_id TERBARU)
      const bkChat = [...all].reverse().find(m => m.client_id === c.client_id && m.booking_id);
      let bookingInfo = null;
      if (bkChat?.booking_id) {
        const booking = getOne('appointments', bkChat.booking_id) || getOne('bookings', bkChat.booking_id);
        if (booking) {
          bookingInfo = {
            booking_id: bkChat.booking_id,
            booking_code: booking.booking_code || bkChat.booking_id,
            service: booking.service || booking.service_name || '-',
            date: booking.date || '-',
            time: booking.time || '-',
            type: booking.type || '-',
            status: booking.status || '-'
          };
        }
      }

      return {
        client_id: c.client_id,
        client_name: clientName,
        mama_name: mamaName,
        last_message: (c.message || '').substring(0, 80),
        last_sender: c.sender,
        last_time: c.created_at,
        unread: all.filter(m => m.client_id === c.client_id && m.sender === 'client' && !m.read).length,
        booking: bookingInfo
      };
    }).sort((a, b) => (b.last_time || '').localeCompare(a.last_time || ''));
    
    // Filter by search (client_name OR mama_name)
    const filtered = search
      ? rooms.filter(r => r.client_name.toLowerCase().includes(search) || r.mama_name.toLowerCase().includes(search))
      : rooms;
    
    res.json(filtered);
  });

  app.get('/api/admin/booking-chats/:client_id', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const chats = readAll('booking_chats')
      .filter(c => c.client_id === req.params.client_id)
      .sort((a, b) => (a.created_at || '').localeCompare(b.created_at || ''));
    
    chats.forEach(c => { if (c.sender === 'client' && !c.read) {
      saveOne('booking_chats', c.id, { read: true });
    }});
    
    // Cari booking info (gunakan booking_id TERBARU)
    const bkChat = [...chats].reverse().find(c => c.booking_id);
    let bookingInfo = null;
    if (bkChat?.booking_id) {
      const booking = getOne('appointments', bkChat.booking_id);
      const finalBooking = booking || getOne('bookings', bkChat.booking_id);
      if (finalBooking) {
        bookingInfo = {
          booking_id: bkChat.booking_id,
          booking_code: finalBooking.booking_code || bkChat.booking_id,
          service: finalBooking.service || finalBooking.service_name || '-',
          date: finalBooking.date || '-',
          time: finalBooking.time || '-',
          type: finalBooking.type || '-',
          status: finalBooking.status || '-'
        };
      }
    }

    res.json({ messages: chats, client_name: chats[0]?.client_name || 'Client', booking: bookingInfo });
  });

  app.post('/api/admin/booking-chats/init/:booking_id', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });

    const booking = getOne('appointments', req.params.booking_id) || getOne('bookings', req.params.booking_id);
    if (!booking) return res.status(404).json({ error: 'Booking tidak ditemukan' });

    // Cari client_id dari booking
    const clientId = booking.client_id || booking.clientId;
    if (!clientId) return res.status(400).json({ error: 'Booking tidak memiliki client_id' });

    // Cek apakah sudah ada chat room
    const existingChats = readAll('booking_chats').filter(c => c.client_id === clientId);
    const hasBookingChat = existingChats.some(c => c.booking_id === req.params.booking_id);

    // Kalau belum ada chat dengan booking_id ini, buat system message
    if (!hasBookingChat) {
      const client = getOne('clients', clientId);
      const clientName = booking.client_name || client?.name || 'Client';

      // Simpan system message untuk inisiasi room
      saveOne('booking_chats', uid(), {
        client_id: clientId,
        client_name: clientName,
        message: `📋 Chat dari booking #${booking.booking_code || req.params.booking_id}`,
        sender: 'system',
        admin_name: tok.name || 'Admin',
        created_at: new Date().toISOString(),
        read: true,
        booking_id: req.params.booking_id
      });
    }

    const bService = booking.service || booking.service_name || '-';
    const bDate = booking.date || '-';
    const bTime = booking.time || '-';
    const bCode = booking.booking_code || req.params.booking_id;

    res.json({
      client_id: clientId,
      booking: {
        booking_id: req.params.booking_id,
        booking_code: bCode,
        service: bService,
        date: bDate,
        time: bTime,
        type: booking.type || '-',
        status: booking.status || '-'
      }
    });
  });

  app.post('/api/admin/booking-chats/:client_id', (req, res) => {
    const tok = verifyToken(req);
    if (!tok || tok.role !== 'admin') return res.status(401).json({ error: 'Unauthorized', code: 401 });
    
    const { message } = req.body;
    if (!message || !message.trim()) return res.status(400).json({ error: 'Pesan tidak boleh kosong' });
    
    const existing = readAll('booking_chats').find(c => c.client_id === req.params.client_id);
    
    const chat = saveOne('booking_chats', uid(), {
      client_id: req.params.client_id,
      client_name: existing?.client_name || 'Client',
      message: message.trim(),
      sender: 'admin',
      admin_name: tok.name || 'Admin Lelap',
      created_at: new Date().toISOString(),
      read: false
    });
    
    res.status(201).json(chat);
  });

  console.log('✅ Membership, Loyalty & Consultation APIs mounted');

  console.log('Public API routes mounted');
}

// Activate routes
setupPublicRoutes(app, { readAll, saveOne, getOne, uid, getSettings, createJWT, verifyJWT });
console.log('✅ Public API routes INLINE mounted');

// ── CRM: Client list with points ──
app.get('/api/crm/clients', (req, res) => {
  const tok = verifyToken(req);
  if (!tok) return res.status(401).json({ error: 'Unauthorized', code: 401 });
  const clients = readAll('clients');
  const apps = readAll('appointments');
  const pointsTx = readAll('points_tx');
  
  const result = clients.map(c => {
    // Count completed bookings
    const clientApps = apps.filter(a => 
      (a.client_id === c.id) || 
      ((a.client_name || '').trim().toLowerCase() === (c.name || '').trim().toLowerCase() &&
       (a.wa || a.phone || '').trim() === (c.phone || c.wa || '').trim())
    );
    const selesai = clientApps.filter(a => a.status === 'Selesai' || a.status === 'Lunas');
    const sorted = selesai.sort((a, b) => (b.date || '').localeCompare(a.date || ''));
    
    // Calculate points balance from points_tx (booking + google_review only)
    const now = new Date();
    let balance = 0;
    const clientTx = pointsTx.filter(t => t.client_id === c.id);
    for (const t of clientTx) {
      if (t.type === 'redeem') { balance -= (t.amount || 0); continue; }
      if (!t.expires_at) continue;
      if (new Date(t.expires_at) < now) continue;
      if (t.source === 'booking' || t.source === 'google_review') balance += (t.amount || 0);
    }
    // spending_points from client.total_spending (cumulative Rp10k = 1pt)
    const spendingPoints = Math.floor((c.total_spending || 0) / 10000);
    
    return {
      id: c.id,
      name: c.name || '',
      phone: c.phone || c.wa || '',
      address: c.address || '',
      orders: selesai.length,
      lastDate: sorted[0]?.date || '',
      firstDate: selesai[selesai.length-1]?.date || '',
      loyalty_points: c.loyalty_points || balance,
      spending_points: spendingPoints,
      tier: c.tier || null,
      total_spending: c.total_spending || 0
    };
  });
  
  res.json(result);
});

// ── 404 for unmatched API routes (must be before catch-all dashboard) ──
app.all('/api/*', (req, res) => {
  res.status(404).json({ error: 'Not found', code: 404, path: req.originalUrl });
});

// ── Catch-all: serve dashboard (must be LAST) ──
app.get('*', (req, res) => {
  // Only serve dashboard to authenticated admin users
  const tok = verifyToken(req);
  if (!tok || (tok.role !== 'admin' && tok.role !== 'owner')) {
    // Redirect to login page instead of exposing dashboard HTML
    return res.status(401).json({ error: 'Unauthorized', code: 401 });
  }
  res.set('Cache-Control', 'no-store, no-cache, must-revalidate');
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ── Global error handler (jangan crash karena bad JSON) ──
app.use((err, req, res, next) => {
  if (err.type === 'entity.parse.failed' || err instanceof SyntaxError) {
    return res.status(400).json({ error: 'Invalid JSON' });
  }
  console.error('Unhandled error:', err.message);
  res.status(500).json({ error: 'Internal server error' });
});

// ── Start ──
const PORT = process.env.PORT || 3000;
app.listen(PORT, '0.0.0.0', () => {
  console.log(`SmartSpaDash running on port ${PORT}`);
});
