// ── Public API Module ──
// Lelap Booking Care — Client App Backend
// Dibangun oleh Hermes Agent — 6 Juni 2026

const crypto = require('crypto');
const path = require('path');
const fs = require('fs');

// Pake shared helpers dari server.js (readAll, saveOne, getOne, deleteOne, uid, nowISO, getFilePath)
// Function ini akan di-pass dari server.js saat init

let readAll, saveOne, getOne, deleteOne, uid, nowISO, getFilePath, DATA_DIR;

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
  if (!token) return res.status(401).json({ error: 'Unauthorized' });
  const decoded = verifyJWT(token);
  if (!decoded) return res.status(401).json({ error: 'Invalid or expired token' });
  req.client = decoded;
  next();
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
    
    // Check if slot is available
    let available = true;
    let reason = '';
    
    // Check existing appointments for this date+time+therapist
    const overlapping = allAppointments.filter(a => {
      if (a.date !== date) return false;
      if (a.status === 'cancelled' || a.status === 'Dibatalkan') return false;
      if (therapist && a.staff !== therapist && a.therapist !== therapist) return false;
      // Check time overlap
      const aStart = timeToMinutes(a.time);
      const aDuration = a.duration || 60;
      const aEnd = aStart + aDuration;
      const slotStart = m;
      const slotEnd = m + duration;
      return slotStart < aEnd && slotEnd > aStart;
    });
    
    if (overlapping.length > 0) {
      available = false;
      reason = 'full';
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
      reason: reason || (available ? null : 'full')
    });
  }
  
  return slots;
}

function timeToMinutes(timeStr) {
  const [h, m] = (timeStr || '00:00').split(':').map(Number);
  return h * 60 + m;
}

// ── API Routes ──
function setupPublicRoutes(app, helpers) {
  const { readAll, saveOne, getOne, uid, getSettings, createJWT, verifyJWT } = helpers;
  console.log('DEBUG helpers loaded:', typeof readAll, typeof saveOne, typeof getOne);
  
  // === AUTH ===
  app.post('/api/public/auth/google', async (req, res) => {
    const { firebase_token, name, phone, city, district, latitude, longitude } = req.body;
    if (!firebase_token) return res.status(400).json({ error: 'firebase_token required' });
    
    const googleUser = await verifyGoogleToken(firebase_token);
    if (!googleUser) return res.status(401).json({ error: 'Invalid Google token' });
    
    const email = googleUser.email;
    const googleId = googleUser.sub;
    
    // Find or create client
    let client = null;
    const allClients = readAll('clients');
    client = allClients.find(c => c.google_id === googleId || c.email === email);
    
    if (!client) {
      // ── VALIDASI LOKASI SAAT REGISTRASI ──
      const clientCity = city || '';
      const clientDistrict = district || '';
      const clientLat = parseFloat(latitude) || null;
      const clientLng = parseFloat(longitude) || null;
      
      if (!validateLocation(clientCity, clientDistrict)) {
        return res.status(400).json({ 
          error: 'Maaf, saat ini Lelap hanya melayani area Salatiga dan Kabupaten Semarang.',
          error_code: 'LOCATION_OUT_OF_SERVICE'
        });
      }
      
      const id = uid();
      client = saveOne('clients', id, {
        google_id: googleId,
        email: email,
        name: name || googleUser.name || email.split('@')[0],
        phone: phone || '',
        city: clientCity,
        district: clientDistrict,
        latitude: clientLat,
        longitude: clientLng,
        profiles: [{ id: 'prof1', name: name || googleUser.name || email.split('@')[0], type: 'adult', gender: 'female' }],
        loyalty_points: 0
      });
    } else {
      // ── UPDATE LOKASI SAAT LOGIN ULANG ──
      if (city !== undefined || district !== undefined || latitude !== undefined) {
        const updates = {};
        if (city !== undefined) updates.city = city;
        if (district !== undefined) updates.district = district;
        if (latitude !== undefined) updates.latitude = parseFloat(latitude) || null;
        if (longitude !== undefined) updates.longitude = parseFloat(longitude) || null;
        client = saveOne('clients', client.id, updates);
      }
    }
    
    const jwt = createJWT({ id: client.id, email: client.email, name: client.name });
    res.json({ token: jwt, user: client });
  });
  
  app.get('/api/public/profile', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    res.json(client);
  });
  
  app.put('/api/public/profile', publicAuth, (req, res) => {
    const { name, phone, address, city, district } = req.body;
    const updates = {};
    if (name !== undefined) updates.name = name;
    if (phone !== undefined) updates.phone = phone;
    if (address !== undefined) updates.address = address;
    if (city !== undefined) updates.city = city;
    if (district !== undefined) updates.district = district;
    
    const client = saveOne('clients', req.client.id, updates);
    res.json(client);
  });
  
  // === MULTI-PROFILE ===
  app.post('/api/public/profiles', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    
    const profiles = client.profiles || [];
    const newProfile = {
      id: 'prof' + (profiles.length + 1),
      name: req.body.name,
      type: req.body.type || 'child',
      gender: req.body.gender || 'male',
      birth_date: req.body.birth_date || '',
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
  // === BOOKINGS ===
  app.post('/api/public/bookings', publicAuth, (req, res) => {
    const { service_id, date, time, therapist, profile_id, payment_method, deposit, latitude, longitude } = req.body;
    if (!service_id || !date || !time) return res.status(400).json({ error: 'service_id, date, time required' });
    
    const service = getOne('services', service_id);
    if (!service) return res.status(404).json({ error: 'Service not found' });
    
    // ── HC (Homecare) = wajib GPS ──
    const isHC = (service.category || '').toUpperCase() === 'HOMECARE' ||
                 (service.code || '').toUpperCase().includes('HC') || 
                 (service.name || '').toUpperCase().includes('HC');
    if (isHC) {
      const lat = parseFloat(latitude);
      const lng = parseFloat(longitude);
      if (!lat || !lng || isNaN(lat) || isNaN(lng)) {
        return res.status(400).json({ 
          error: 'Layanan Homecare (HC) memerlukan lokasi GPS. Mohon aktifkan GPS Anda.',
          error_code: 'HC_GPS_REQUIRED'
        });
      }
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
    
    // Validate location
    if (client.address && client.city) {
      if (!validateLocation(client.city, client.district || '')) {
        return res.status(400).json({ error: 'Maaf, saat ini Lelap hanya melayani area Salatiga dan Kabupaten Semarang.' });
      }
    }
    
    const amount = service.price || 0;
    const bookingCode = 'MBS-' + date.replace(/-/g, '').substring(2) + '-' + Math.floor(Math.random() * 900 + 100);
    const fixTherapist = service.name && service.name.toUpperCase().includes('PRENATAL YOGA') ? 'Owner' : (therapist || '');
    
    const booking = saveOne('appointments', uid(), {
      date,
      type: isHC ? 'Homecare' : 'Inhouse',
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
      address: client.address || '',
      discount: '0',
      transport: '0',
      deposit: deposit || '0',
      payment: payment_method === 'qris' ? 'QRIS' : payment_method === 'transfer' ? 'Transfer' : 'Cash',
      notes: '',
      client_type: profile.type === 'child' ? 'Anak' : 'Dewasa',
      staff: fixTherapist,
      status: (payment_method === 'cash' || payment_method === 'dp') ? 'Menunggu Pembayaran' : 'Confirmed',
      amount,
      booking_code: bookingCode,
      client_id: client.id,
      profile_id: profile_id || '',
      duration: service.duration || 60,
      source: 'app',
      latitude: isHC ? parseFloat(latitude) || null : null,
      longitude: isHC ? parseFloat(longitude) || null : null,
      is_homecare: isHC
    });
    
    // Add loyalty points
    const points = (client.loyalty_points || 0) + 10 + Math.floor(amount / 1000);
    saveOne('clients', client.id, { loyalty_points: points });
    
    res.status(201).json({ booking, booking_code: bookingCode });
  });
  
  app.get('/api/public/bookings', publicAuth, (req, res) => {
    const all = readAll('appointments');
    const mine = all.filter(a => a.client_id === req.client.id)
      .sort((a, b) => (b.created_at || '').localeCompare(a.created_at || ''));
    res.json(mine);
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
  app.post('/api/public/reviews', publicAuth, (req, res) => {
    const { booking_id, rating, comment, therapist_rating, photos } = req.body;
    if (!booking_id || !rating) return res.status(400).json({ error: 'booking_id and rating required' });
    
    const booking = getOne('appointments', booking_id);
    if (!booking) return res.status(404).json({ error: 'Booking not found' });
    if (booking.client_id !== req.client.id) return res.status(403).json({ error: 'Forbidden' });
    
    // Check if already reviewed
    const existingReview = readAll('reviews').find(r => r.booking_id === booking_id && r.client_id === req.client.id);
    if (existingReview) return res.status(409).json({ error: 'Booking sudah direview', review: existingReview });
    
    const review = saveOne('reviews', uid(), {
      booking_id,
      client_id: req.client.id,
      client_name: req.client.name || '',
      rating,
      comment: comment || '',
      photos: Array.isArray(photos) ? photos.slice(0, 3) : [], // max 3 photos (base64)
      therapist_rating: therapist_rating || rating,
      therapist: booking.staff || '',
      service: booking.service,
      date: booking.date,
      created_at: new Date().toISOString()
    });
    
    // Award loyalty points for review
    try {
      earnPoints(req.client.id, 10, 'review', 'Ulasan untuk ' + (booking.service || 'booking'));
    } catch(_) {}
    
    res.status(201).json({
      ...review,
      points_earned: 10,
      google_review_link: 'https://www.google.com/maps/place/Lelap+Mom+Baby+Care+Salatiga/@-7.3285,110.5100,17z/data=!4m8!3m7!1s0x2e7a79a87d92c81f:0x61c355bf9ad20e62!8m2!3d-7.3285488!4d110.5100338!9m1!1b1'
    });
  });
  
  // Check if booking already reviewed
  app.get('/api/public/reviews/booking/:booking_id', publicAuth, (req, res) => {
    const review = readAll('reviews').find(r => 
      r.booking_id === req.params.booking_id && r.client_id === req.client.id
    );
    res.json({ reviewed: !!review, review: review || null });
  });
  
  app.get('/api/public/reviews/:service_id', (req, res) => {
    const all = readAll('reviews');
    const serviceReviews = all.filter(r => {
      const booking = getOne('appointments', r.booking_id);
      return booking && (booking.kode === req.params.service_id || booking.service === req.params.service_id);
    });
    const avg = serviceReviews.length > 0 
      ? Math.round(serviceReviews.reduce((s, r) => s + r.rating, 0) / serviceReviews.length * 10) / 10
      : 0;
    res.json({ reviews: serviceReviews, average_rating: avg, count: serviceReviews.length });
  });
  
  // === LOYALTY ===
  app.get('/api/public/loyalty', publicAuth, (req, res) => {
    const client = getOne('clients', req.client.id);
    if (!client) return res.status(404).json({ error: 'Client not found' });
    res.json({ points: client.loyalty_points || 0 });
  });
  
  // === CHAT BOT ===
  app.post('/api/public/chat', publicAuth, (req, res) => {
    const { message } = req.body;
    const msg = (message || '').toLowerCase().trim();
    
    const faq = {
      'jam buka': 'Lelap buka setiap hari, jam 08:00 - 20:00 WIB, Ma~ 🕐',
      'buka': 'Lelap buka setiap hari, jam 08:00 - 20:00 WIB, Ma~ 🕐',
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
  
  // === MIDTRANS CALLBACK ===
  app.post('/api/public/payment/callback', (req, res) => {
    const { order_id, transaction_status, fraud_status } = req.body;
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
  
  console.log('Public API routes mounted');
}

module.exports = { setupPublicRoutes };
