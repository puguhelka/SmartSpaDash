
/**
 * SmartSpaDash Premium Add-On
 * Version: 1.0
 * Purpose:
 *  - Adds safe status separation (payment vs session done)
 *  - Adds system log
 *  - Adds user role session + lightweight login
 *  - Adds sheet/range protection helpers
 *  - Adds lock-safe booking ID generator
 *  - Adds CRM favorite service recalculation
 *
 * This file is ADDITIVE. It is designed to sit beside the existing SmartSpaDash_2203.gs
 * and minimize changes to the current flow.
 */

var SSD_PREMIUM = {
  sheets: {
    users: 'USER ACCESS',
    systemLog: 'SYSTEM LOG',
    statusMap: 'STATUS MAP'
  },
  login: {
    sheetName: 'LOGIN',
    emailCell: 'C6',
    pinCell: 'F6',
    roleCell: 'C8',
    statusCell: 'F8',
    lastLoginCell: 'C9'
  },
  properties: {
    sessionRole: 'SSD_SESSION_ROLE',
    sessionEmail: 'SSD_SESSION_EMAIL',
    selectionGuard: 'SSD_SELECTION_GUARD'
  },
  actionCells: {
    markPaid: 'H24',
    markSessionDone: 'J24'
  },
  input: {
    bookingCodeCell: 'J6',
    searchResultBookingCodeCell: 'K6'
  },
  allowedRoles: ['OWNER', 'ADMIN'],
  version: '1.2'
};

function ssdPremiumOnOpen_() {
  ssdPremiumEnsureSupportSheets_();
  ssdPremiumSetLoginSheetStatus_();
  ssdPremiumBuildMenu_();
}

function ssdPremiumBuildMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('SmartSpaDash Premium')
    .addItem('Otorisasi Akun Google', 'ssdPremiumSetup_')
    .addSeparator()
    .addItem('Sinkron Favorit CRM', 'ssdPremiumRefreshFavoriteService_')
    .addSeparator()
    .addItem('Terapkan Proteksi Sheet', 'ssdPremiumProtectSheets_')
    .addItem('Lepaskan Proteksi Sheet', 'ssdPremiumUnprotectOwnedProtections_')
    .addToUi();
}

function ssdPremiumSetup_() {
  ssdPremiumEnsureSupportSheets_();
  ssdPremiumSeedStatusMap_();
  ssdPremiumSeedUsersGuide_();
  SpreadsheetApp.getActive().toast('SmartSpaDash : Otorisasi akun Google berhasil disiapkan.', 'SmartSpaDash', 5);
}

function ssdPremiumHandleSelection_(e) {
  if (!e || !e.range) return false;
  if (typeof isSelectionGuardActive_ === 'function' && isSelectionGuardActive_()) return false;

  var sh = e.range.getSheet();
  var a1 = e.range.getA1Notation();

  if (sh.getName() === SSD_PREMIUM.login.sheetName) {
    if (ssdPremiumRangeHit_(e.range, 'H6:I7')) {
      var loginMsg = ssdPremiumLoginValidated_();
      if (typeof toastAndResetSelection_ === 'function') {
        toastAndResetSelection_(sh, SSD_PREMIUM.login.emailCell, loginMsg, 5);
      } else {
        SpreadsheetApp.getActive().toast('SmartSpaDash : ' + loginMsg, 'SmartSpaDash', 5);
      }
      return true;
    }

    if (ssdPremiumRangeHit_(e.range, 'H8:I9')) {
      var logoutMsg = ssdPremiumLogoutValidated_();
      if (typeof toastAndResetSelection_ === 'function') {
        toastAndResetSelection_(sh, SSD_PREMIUM.login.emailCell, logoutMsg, 5);
      } else {
        SpreadsheetApp.getActive().toast('SmartSpaDash : ' + logoutMsg, 'SmartSpaDash', 5);
      }
      return true;
    }

    return false;
  }

  if (sh.getName() !== SSD.sheets.input) return false;

  if (a1 === SSD_PREMIUM.actionCells.markPaid) {
    var msg1 = ssdPremiumUpdatePaymentValidated_();
    if (typeof toastAndResetSelection_ === 'function') {
      toastAndResetSelection_(sh, SSD.cells.bookingDate, msg1, 5);
    }
    return true;
  }

  if (a1 === SSD_PREMIUM.actionCells.markSessionDone) {
    var msg2 = ssdPremiumUpdateServiceDoneValidated_();
    if (typeof toastAndResetSelection_ === 'function') {
      toastAndResetSelection_(sh, SSD.cells.bookingDate, msg2, 5);
    }
    return true;
  }

  return false;
}

function ssdPremiumRangeHit_(clickedRange, targetA1) {
  var sh = clickedRange.getSheet();
  var target = sh.getRange(targetA1);

  var r1 = clickedRange.getRow();
  var c1 = clickedRange.getColumn();
  var r2 = r1 + clickedRange.getNumRows() - 1;
  var c2 = c1 + clickedRange.getNumColumns() - 1;

  var tr1 = target.getRow();
  var tc1 = target.getColumn();
  var tr2 = tr1 + target.getNumRows() - 1;
  var tc2 = tc1 + target.getNumColumns() - 1;

  return !(r2 < tr1 || r1 > tr2 || c2 < tc1 || c1 > tc2);
}

function ssdPremiumLoginValidated_() {
  var loginSh = SpreadsheetApp.getActive().getSheetByName(SSD_PREMIUM.login.sheetName);
  if (!loginSh) throw new Error('Sheet LOGIN tidak ditemukan.');

  var email = String(loginSh.getRange(SSD_PREMIUM.login.emailCell).getValue() || '').trim().toLowerCase();
  var pin = String(loginSh.getRange(SSD_PREMIUM.login.pinCell).getValue() || '').trim();

  if (!email) throw new Error('Email login wajib diisi.');
  if (!pin) throw new Error('PIN login wajib diisi.');

  var usersSh = ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.users);
  var rows = ssdPremiumGetObjects_(usersSh);
  var found = rows.find(function(r) {
    return String(r.email || '').trim().toLowerCase() === email && String(r.pin || '').trim() === pin && String(r.is_active || 'Y').trim().toUpperCase() === 'Y';
  });

  if (!found) throw new Error('Email / PIN tidak cocok atau user nonaktif.');

  var role = String(found.role || 'ADMIN').trim().toUpperCase();
  var userProps = PropertiesService.getUserProperties();
  userProps.setProperty(SSD_PREMIUM.properties.sessionEmail, email);
  userProps.setProperty(SSD_PREMIUM.properties.sessionRole, role);

  loginSh.getRange(SSD_PREMIUM.login.roleCell).setValue(role);
  loginSh.getRange(SSD_PREMIUM.login.statusCell).setValue('LOGIN OK');
  loginSh.getRange(SSD_PREMIUM.login.lastLoginCell).setValue(new Date());

  ssdPremiumAppendSystemLog_({
    action: 'LOGIN_OK',
    bookingId: '',
    actorEmail: email,
    actorRole: role,
    targetSheet: SSD_PREMIUM.login.sheetName,
    details: 'Login berhasil.'
  });

  SpreadsheetApp.getActive().toast('SmartSpaDash : Login berhasil sebagai ' + role, 'SmartSpaDash', 4);
  return 'Login berhasil sebagai ' + role;
}

function ssdPremiumLogoutValidated_() {
  var userProps = PropertiesService.getUserProperties();
  var email = userProps.getProperty(SSD_PREMIUM.properties.sessionEmail) || '';
  var role = userProps.getProperty(SSD_PREMIUM.properties.sessionRole) || '';

  userProps.deleteProperty(SSD_PREMIUM.properties.sessionEmail);
  userProps.deleteProperty(SSD_PREMIUM.properties.sessionRole);

  var loginSh = SpreadsheetApp.getActive().getSheetByName(SSD_PREMIUM.login.sheetName);
  if (loginSh) {
    loginSh.getRange(SSD_PREMIUM.login.roleCell).clearContent();
    loginSh.getRange(SSD_PREMIUM.login.statusCell).setValue('LOGOUT');
  }

  ssdPremiumAppendSystemLog_({
    action: 'LOGOUT',
    bookingId: '',
    actorEmail: email,
    actorRole: role,
    targetSheet: SSD_PREMIUM.login.sheetName,
    details: 'Logout manual.'
  });

  SpreadsheetApp.getActive().toast('SmartSpaDash : Logout berhasil.', 'SmartSpaDash', 4);
  return 'Logout berhasil.';
}

function ssdPremiumSetLoginSheetStatus_() {
  var loginSh = SpreadsheetApp.getActive().getSheetByName(SSD_PREMIUM.login.sheetName);
  if (!loginSh) return;

  var userProps = PropertiesService.getUserProperties();
  var email = userProps.getProperty(SSD_PREMIUM.properties.sessionEmail) || '';
  var role = userProps.getProperty(SSD_PREMIUM.properties.sessionRole) || '';

  loginSh.getRange(SSD_PREMIUM.login.roleCell).setValue(role || '');
  loginSh.getRange(SSD_PREMIUM.login.statusCell).setValue(email ? 'SESSION AKTIF' : 'BELUM LOGIN');
}

function ssdPremiumRequireRole_(allowedRoles) {
  var role = String(PropertiesService.getUserProperties().getProperty(SSD_PREMIUM.properties.sessionRole) || '').trim().toUpperCase();
  if (!role) throw new Error('Silakan login dahulu pada sheet LOGIN.');
  if (allowedRoles && allowedRoles.length && allowedRoles.indexOf(role) === -1) {
    throw new Error('Role ' + role + ' tidak diizinkan untuk aksi ini.');
  }
  return role;
}

function ssdPremiumGetCurrentEmail_() {
  return String(PropertiesService.getUserProperties().getProperty(SSD_PREMIUM.properties.sessionEmail) || '').trim().toLowerCase();
}


function ssdPremiumUpdatePaymentValidated_() {
  ssdPremiumRequireRole_(SSD_PREMIUM.allowedRoles);
  return ssdPremiumWithLock_(function() {
    var bookingId = ssdPremiumResolveBookingIdFromInput_();
    var match = ssdPremiumFindLogRowByBookingId_(bookingId);
    if (!match) throw new Error('Kode booking ' + bookingId + ' tidak ditemukan di TRANSAKSI_LOG.');

    var sh = match.sheet;
    var rowIndex = match.rowIndex;
    var idx = match.idx;
    var row = match.row;
    var inputSh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.input);

    var grandTotal = Number(row[idx.grand_total] || 0);
    var oldPaymentStatus = row[idx.payment_status] || '';
    var oldBalance = Number(row[idx.remaining_balance] || 0);

    sh.getRange(rowIndex, idx.payment_status + 1).setValue('Lunas');
    sh.getRange(rowIndex, idx.remaining_balance + 1).setValue(0);

    if (inputSh && SSD && SSD.cells && SSD.cells.searchResultRemainingBalance) {
      inputSh.getRange(SSD.cells.searchResultRemainingBalance).setValue(0);
    }

    ssdPremiumAppendSystemLog_({
      action: 'MARK_PAID',
      bookingId: row[idx.booking_id] || bookingId,
      oldValue: JSON.stringify({ payment_status: oldPaymentStatus, remaining_balance: oldBalance }),
      newValue: JSON.stringify({ payment_status: 'Lunas', remaining_balance: 0 }),
      targetSheet: sh.getName(),
      targetRow: rowIndex,
      details: 'Pelunasan manual berbasis hasil pencarian booking.'
    });

    if (typeof refreshRekapHarian_ === 'function') refreshRekapHarian_();
    if (typeof refreshRekapBulanan_ === 'function') refreshRekapBulanan_();
    if (typeof refreshPendingBookingList_ === 'function') refreshPendingBookingList_();
    if (typeof refreshDashboardV2_ === 'function') refreshDashboardV2_();
    if (typeof setStatusMessage_ === 'function') setStatusMessage_('Booking ' + (row[idx.booking_id] || bookingId) + ' ditandai Lunas.');

    return 'Booking ' + (row[idx.booking_id] || bookingId) + ' ditandai Lunas. Total: ' + grandTotal;
  });
}

function ssdPremiumUpdateServiceDoneValidated_() {
  ssdPremiumRequireRole_(SSD_PREMIUM.allowedRoles);
  return ssdPremiumWithLock_(function() {
    var bookingId = ssdPremiumResolveBookingIdFromInput_();
    var match = ssdPremiumFindLogRowByBookingId_(bookingId);
    if (!match) throw new Error('Kode booking ' + bookingId + ' tidak ditemukan di TRANSAKSI_LOG.');

    var sh = match.sheet;
    var rowIndex = match.rowIndex;
    var idx = match.idx;
    var row = match.row;

    var oldBookingStatus = row[idx.booking_status] || '';
    var oldServiceDone = row[idx.service_done_status] || '';
    var paymentStatus = String(row[idx.payment_status] || '').trim();

    sh.getRange(rowIndex, idx.service_done_status + 1).setValue('DONE');
    sh.getRange(rowIndex, idx.booking_status + 1).setValue('Done');

    ssdPremiumAppendSystemLog_({
      action: 'MARK_SESSION_DONE',
      bookingId: row[idx.booking_id] || bookingId,
      oldValue: JSON.stringify({ booking_status: oldBookingStatus, service_done_status: oldServiceDone, payment_status: paymentStatus }),
      newValue: JSON.stringify({ booking_status: 'Done', service_done_status: 'DONE', payment_status: paymentStatus }),
      targetSheet: sh.getName(),
      targetRow: rowIndex,
      details: 'Sesi selesai diproses berbasis hasil pencarian booking.'
    });

    if (typeof refreshRekapHarian_ === 'function') refreshRekapHarian_();
    if (typeof refreshRekapBulanan_ === 'function') refreshRekapBulanan_();
    if (typeof refreshPendingBookingList_ === 'function') refreshPendingBookingList_();
    if (typeof refreshDashboardV2_ === 'function') refreshDashboardV2_();
    if (typeof setStatusMessage_ === 'function') setStatusMessage_('Booking ' + (row[idx.booking_id] || bookingId) + ' ditandai Sesi Selesai.');

    return 'Booking ' + (row[idx.booking_id] || bookingId) + ' ditandai Sesi Selesai.';
  });
}

function ssdPremiumResolveBookingIdFromInput_() {
  var sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.input);
  if (!sh) throw new Error('Sheet input tidak ditemukan: ' + SSD.sheets.input);

  var resultBookingCodeCell = (SSD && SSD.cells && SSD.cells.searchResultBookingCode) ? SSD.cells.searchResultBookingCode : SSD_PREMIUM.input.searchResultBookingCodeCell;
  var bookingId = String(sh.getRange(resultBookingCodeCell).getDisplayValue() || '').trim().toUpperCase();

  if (!bookingId) {
    throw new Error('Lakukan pencarian booking dulu melalui F16 lalu klik J20. Hasil booking harus muncul di K6 sebelum proses dilanjutkan.');
  }
  return bookingId;
}

function ssdPremiumFindLogRowByBookingId_(bookingId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.log);
  if (!sh) throw new Error('Sheet log tidak ditemukan: ' + SSD.sheets.log);

  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return null;

  var idx = ssdPremiumIndexMap_(vals[0]);
  for (var i = vals.length - 1; i >= 1; i--) {
    var row = vals[i];
    if (String(row[idx.booking_id] || '').trim().toUpperCase() === bookingId &&
        ssdPremiumCleanText_(row[idx.record_status] || 'Active') === 'Active') {
      return { sheet: sh, rowIndex: i + 1, idx: idx, row: row };
    }
  }
  return null;
}

function ssdPremiumBookingCodeExistsInSystemLog_(bookingId) {
  var sh = SpreadsheetApp.getActive().getSheetByName(SSD_PREMIUM.sheets.systemLog);
  if (!sh || sh.getLastRow() < 2) return false;
  var values = sh.getDataRange().getValues();
  var idx = ssdPremiumIndexMap_(values[0]);
  for (var i = values.length - 1; i >= 1; i--) {
    var row = values[i];
    if (String(row[idx.booking_id] || '').trim().toUpperCase() === bookingId) return true;
  }
  return false;
}

function ssdPremiumFindLogRowByForm_(form) {
  var bookingId = '';
  try { bookingId = ssdPremiumResolveBookingIdFromInput_(); } catch (err) {}
  if (bookingId) return ssdPremiumFindLogRowByBookingId_(bookingId);

  var sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.log);
  if (!sh) throw new Error('Sheet log tidak ditemukan: ' + SSD.sheets.log);

  var vals = sh.getDataRange().getValues();
  if (!vals || vals.length < 2) return null;

  var idx = ssdPremiumIndexMap_(vals[0]);

  for (var i = vals.length - 1; i >= 1; i--) {
    var row = vals[i];
    if (
      ssdPremiumNormalizePhone_(row[idx.phone]) === ssdPremiumNormalizePhone_(form.phone) &&
      ssdPremiumSameDate_(row[idx.booking_date], form.bookingDate) &&
      ssdPremiumSameTime_(row[idx.start_time], form.startTime) &&
      ssdPremiumCleanText_(row[idx.therapist_name]) === ssdPremiumCleanText_(form.therapistName) &&
      ssdPremiumCleanText_(row[idx.service_name]) === ssdPremiumCleanText_(form.serviceName) &&
      ssdPremiumCleanText_(row[idx.record_status] || 'Active') === 'Active'
    ) {
      return { sheet: sh, rowIndex: i + 1, idx: idx, row: row };
    }
  }

  return null;
}

function ssdPremiumSafeBuildBookingId_(bookingDate) {
  return ssdPremiumWithLock_(function() {
    var docProps = PropertiesService.getDocumentProperties();
    var dt = ssdPremiumParseDateValue_(bookingDate) || new Date();
    var prefix = Utilities.formatDate(dt, SSD.timezone || 'Asia/Jakarta', 'yyyyMMdd');
    var key = 'SSD_BOOKING_SEQ_GLOBAL';

    var sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.log);
    var vals = sh ? sh.getDataRange().getValues() : [];
    var existingMax = 0;

    for (var i = 1; i < vals.length; i++) {
      var bookingId = String(vals[i][0] || '').trim().toUpperCase();
      var match = bookingId.match(/^BK-(\d{8})-(\d{3,})$/);
      if (match) {
        existingMax = Math.max(existingMax, Number(match[2]) || 0);
      }
    }

    var stored = Number(docProps.getProperty(key) || 0);
    var current = Math.max(existingMax, stored) + 1;

    if (vals.length < 2) current = 1;

    docProps.setProperty(key, String(current));
    return 'BK-' + prefix + '-' + ('000' + current).slice(-3);
  });
}

function ssdPremiumWrapSaveBooking_(originalFn) {
  return function() {
    return ssdPremiumWithLock_(function() {
      return originalFn();
    });
  };
}

function ssdPremiumRefreshFavoriteService_() {
  var crmSh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.pelanggan);
  var logSh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.log);
  if (!crmSh || !logSh) throw new Error('Sheet CRM atau log tidak ditemukan.');

  var crmRows = ssdPremiumGetObjects_(crmSh);
  var logRows = ssdPremiumGetObjects_(logSh).filter(function(r) {
    return ssdPremiumCleanText_(r.record_status || 'Active') === 'Active';
  });

  if (!crmRows.length || !logRows.length) return 'Tidak ada data CRM/log untuk diproses.';

  var crmVals = crmSh.getDataRange().getValues();
  var crmIdx = ssdPremiumIndexMap_(crmVals[0]);

  crmRows.forEach(function(cust, offset) {
    var phone = ssdPremiumNormalizePhone_(cust.phone);
    if (!phone) return;

    var serviceCounter = {};
    logRows.forEach(function(rec) {
      if (ssdPremiumNormalizePhone_(rec.phone) !== phone) return;
      var svc = String(rec.service_name || '').trim();
      if (!svc) return;
      serviceCounter[svc] = (serviceCounter[svc] || 0) + 1;
    });

    var favorite = '';
    var maxCount = 0;
    Object.keys(serviceCounter).forEach(function(name) {
      if (serviceCounter[name] > maxCount) {
        maxCount = serviceCounter[name];
        favorite = name;
      }
    });

    if (crmIdx.favorite_service != null) {
      crmSh.getRange(offset + 2, crmIdx.favorite_service + 1).setValue(favorite);
    }
    if (crmIdx.total_visits != null) {
      crmSh.getRange(offset + 2, crmIdx.total_visits + 1).setValue(Object.values(serviceCounter).reduce(function(a,b){ return a+b; }, 0));
    }
  });

  ssdPremiumAppendSystemLog_({
    action: 'CRM_FAVORITE_REFRESH',
    bookingId: '',
    targetSheet: crmSh.getName(),
    details: 'favorite_service dihitung ulang dari layanan tersering.'
  });

  return 'Favorit layanan CRM berhasil dihitung ulang.';
}

function ssdPremiumProtectSheets_() {
  var role = ssdPremiumRequireRole_(['OWNER']);
  var ss = SpreadsheetApp.getActive();
  var currentEmail = ssdPremiumGetCurrentEmail_() || Session.getEffectiveUser().getEmail();
  var protectedSheets = [SSD.sheets.log, SSD.sheets.pelanggan, 'REKAP HARIAN', 'REKAP BULANAN', 'DASHBOARD', 'INVOICE'];

  protectedSheets.forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) return;
    var protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    var protection = protections && protections.length ? protections[0] : sh.protect();
    protection.setDescription('SmartSpaDash Premium: ' + name);
    var editors = protection.getEditors();
    editors.forEach(function(editor) {
      try { protection.removeEditor(editor); } catch (err) {}
    });
    if (currentEmail) protection.addEditor(currentEmail);
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
  });

  ssdPremiumAppendSystemLog_({
    action: 'APPLY_PROTECTION',
    actorRole: role,
    targetSheet: 'MULTI',
    details: 'Proteksi sheet inti diterapkan.'
  });

  return 'Proteksi sheet inti berhasil diterapkan.';
}

function ssdPremiumUnprotectOwnedProtections_() {
  var ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(function(sh) {
    var protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(function(protection) {
      try {
        if (protection.canEdit()) protection.remove();
      } catch (err) {}
    });
  });
  return 'Proteksi yang dimiliki user aktif telah dilepas.';
}

function ssdPremiumAppendSystemLog_(payload) {
  var sh = ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.systemLog);
  var headers = [
    'timestamp', 'actor_email', 'actor_role', 'action',
    'booking_id', 'target_sheet', 'target_row', 'old_value',
    'new_value', 'details'
  ];

  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.setFrozenRows(1);
  } else {
    var existingHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function(x){ return String(x || '').trim().toLowerCase(); });
    if (existingHeaders.join('|') !== headers.join('|')) {
      sh.clear();
      sh.getRange(1, 1, 1, headers.length).setValues([headers]);
      sh.setFrozenRows(1);
    }
  }

  var actorEmail = payload.actorEmail || ssdPremiumGetCurrentEmail_() || Session.getEffectiveUser().getEmail() || '';
  var actorRole = payload.actorRole || PropertiesService.getUserProperties().getProperty(SSD_PREMIUM.properties.sessionRole) || '';

  sh.appendRow([
    new Date(),
    actorEmail,
    actorRole,
    payload.action || '',
    payload.bookingId || '',
    payload.targetSheet || '',
    payload.targetRow || '',
    payload.oldValue || '',
    payload.newValue || '',
    payload.details || ''
  ]);

  var lastRow = sh.getLastRow();
  sh.getRange(lastRow, 1).setNumberFormat('yyyy-MM-dd HH:mm:ss');
}

function ssdPremiumEnsureSupportSheets_() {
  var ss = SpreadsheetApp.getActive();
  ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.users);
  ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.systemLog);
  ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.statusMap);
}

function ssdPremiumSeedUsersGuide_() {
  var sh = ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.users);
  if (sh.getLastRow() > 1) return;

  var headers = ['email', 'pin', 'role', 'display_name', 'is_active', 'notes'];
  var sample = [
    ['owner@company.com', '1234', 'OWNER', 'Owner Utama', 'Y', 'Ganti PIN setelah setup'],
    ['admin@company.com', '4321', 'ADMIN', 'Admin Frontdesk', 'Y', '']
  ];

  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, sample.length, headers.length).setValues(sample);
  sh.setFrozenRows(1);
}

function ssdPremiumSeedStatusMap_() {
  var sh = ssdPremiumGetOrCreateSheet_(SSD_PREMIUM.sheets.statusMap);
  var headers = ['domain', 'status_code', 'label_id', 'meaning', 'owner_metric'];
  var rows = [
    ['booking_status', 'Active', 'Active', 'Booking aktif dan terjadwal', 'booking_open'],
    ['booking_status', 'Done', 'Done', 'Sesi selesai / closed by operation', 'booking_done'],
    ['booking_status', 'Cancelled', 'Cancelled', 'Dibatalkan', 'booking_cancelled'],
    ['payment_status', 'Belum Bayar', 'Belum Bayar', 'Belum ada pembayaran masuk', 'ar_open'],
    ['payment_status', 'DP', 'DP', 'Deposit masuk tetapi belum lunas', 'ar_partial'],
    ['payment_status', 'Lunas', 'Lunas', 'Tagihan selesai', 'cash_closed'],
    ['service_done_status', 'PENDING', 'Pending', 'Sesi belum dilaksanakan', 'session_open'],
    ['service_done_status', 'ONGOING', 'On Going', 'Sesi sedang berjalan', 'session_live'],
    ['service_done_status', 'DONE', 'Done', 'Sesi sudah selesai', 'session_closed'],
    ['record_status', 'Active', 'Active', 'Baris transaksi aktif', 'included'],
    ['record_status', 'Archived', 'Archived', 'Baris disimpan / tidak aktif', 'excluded']
  ];

  sh.clear();
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sh.setFrozenRows(1);
}

function ssdPremiumGetOrCreateSheet_(name) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(name);
  return sh || ss.insertSheet(name);
}

function ssdPremiumGetObjects_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];
  var idx = ssdPremiumIndexMap_(values[0]);
  return values.slice(1).map(function(row) {
    var out = {};
    Object.keys(idx).forEach(function(key) {
      out[key] = row[idx[key]];
    });
    return out;
  });
}

function ssdPremiumIndexMap_(headerRow) {
  var map = {};
  headerRow.forEach(function(h, i) {
    map[String(h || '').trim().toLowerCase().replace(/\s+/g, '_')] = i;
  });
  return map;
}

function ssdPremiumWithLock_(fn) {
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
}

function ssdPremiumCleanText_(value) {
  return String(value == null ? '' : value).trim();
}

function ssdPremiumNormalizePhone_(value) {
  return String(value == null ? '' : value).replace(/[^\d]/g, '');
}

function ssdPremiumParseDateValue_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) return value;
  if (typeof value === 'number') return new Date(value);
  var txt = String(value || '').trim();
  if (!txt) return null;
  if (/^\d{4}-\d{2}-\d{2}$/.test(txt)) return new Date(txt + 'T00:00:00');
  if (/^\d{2}\/\d{2}\/\d{4}$/.test(txt)) {
    var p = txt.split('/');
    return new Date(Number(p[2]), Number(p[1]) - 1, Number(p[0]));
  }
  var dt = new Date(txt);
  return isNaN(dt) ? null : dt;
}

function ssdPremiumSameDate_(a, b) {
  var da = ssdPremiumParseDateValue_(a);
  var db = ssdPremiumParseDateValue_(b);
  if (!da || !db) return false;
  return Utilities.formatDate(da, SSD.timezone || 'Asia/Jakarta', 'yyyy-MM-dd') ===
         Utilities.formatDate(db, SSD.timezone || 'Asia/Jakarta', 'yyyy-MM-dd');
}

function ssdPremiumSameTime_(a, b) {
  return ssdPremiumTimeToMinutes_(a) === ssdPremiumTimeToMinutes_(b);
}

function ssdPremiumTimeToMinutes_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return value.getHours() * 60 + value.getMinutes();
  }
  var txt = String(value || '').trim();
  if (!txt) return null;
  var m = txt.match(/^(\d{1,2}):(\d{2})/);
  if (!m) return null;
  return Number(m[1]) * 60 + Number(m[2]);
}
