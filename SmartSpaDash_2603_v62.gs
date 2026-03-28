const SSD = {
  sheets: {
    input: 'INPUT BOOKING',
    log: 'TRANSAKSI_LOG',
    pelanggan: 'CRM PELANGGAN',
    invoice: 'INVOICE',
    dashboardV2: 'DASHBOARD V2',
    dashboardV2Data: 'DASHBOARD V2 DATA'
  },
  cells: {
    bookingDate: 'C6',
    bookingType: 'F6',
    therapist: 'C7',
    bookingStatus: 'F7',
    service: 'C8',
    startTime: 'C9',
    area: 'F8',
    promo: 'F9',
    parentName: 'C12',
    phone: 'F12',
    childName: 'C13',
    childAge: 'F13',
    address: 'C14',
    childWeight: 'F14',
    notes: 'C15',
    paymentMethod: 'F15',
    deposit: 'C16',
    category: 'J7',
    duration: 'J8',
    price: 'J9',
    discountValue: 'J10',
    transportFee: 'J11',
    grandTotal: 'J12',
    therapistFee: 'J13',
    adminFee: 'J14',
    remainingBalance: 'J15',
    bookingCode: 'J6',
    searchKey: 'F16',
    searchResultBookingCode: 'K6',
    searchResultCategory: 'K7',
    searchResultDuration: 'K8',
    searchResultPrice: 'K9',
    searchResultDiscountValue: 'K10',
    searchResultTransportFee: 'K11',
    searchResultGrandTotal: 'K12',
    searchResultTherapistFee: 'K13',
    searchResultAdminFee: 'K14',
    searchResultRemainingBalance: 'K15',
    searchResultParentName: 'K16',
    searchResultBookingDate: 'K17',
    mirrorParentName: 'J16',
    mirrorBookingDate: 'J17',
    statusMessage: 'J27'
  },
  actionCells: {
    save: 'H18',
    reset: 'J18',
    conflict: 'H20',
    findCustomer: 'J20',
    invoice: 'J22',
    revertStatus: 'H22',
    dashboardReset: 'K5:L6'
  },
  timezone: 'Asia/Jakarta'
};

function onOpen() {
  showToast_('SmartSpaDash siap.', 'SmartSpaDash', 4);
  try {
    refreshRekapHarian_();
    refreshRekapBulanan_();
    refreshPendingBookingList_();
    refreshDashboardV2_();
  } catch (err) {}
  ssdPremiumOnOpen_();
  if (typeof ssdNavOnOpen_ === 'function') ssdNavOnOpen_();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sh = e.range.getSheet();
  const a1 = e.range.getA1Notation();

  try {
    if (sh.getName() === 'REKAP HARIAN' && (a1 === 'B3' || a1 === 'C3')) {
      refreshRekapHarian_();
      showToast_('Rekap harian berhasil diperbarui.', 'SmartSpaDash', 4);
      return;
    }

    if (sh.getName() === 'REKAP BULANAN' && a1 === 'B3') {
      refreshRekapBulanan_();
      showToast_('Rekap bulanan berhasil diperbarui.', 'SmartSpaDash', 4);
      return;
    }

    if (sh.getName() === SSD.sheets.dashboardV2 && (a1 === 'C5' || a1 === 'F5' || a1 === 'I5')) {
      const msg = refreshDashboardV2_();
      showToast_(msg, 'SmartSpaDash', 4);
      return;
    }
  } catch (err) {
    showToast_(err.message || String(err), 'SmartSpaDash', 6);
  }
}

function onSelectionChange(e) {
  if (!e || !e.range) return;
  if (isSelectionGuardActive_()) return;

  const sh = e.range.getSheet();

  try {
    if (typeof ssdPremiumHandleSelection_ === 'function') {
      if (ssdPremiumHandleSelection_(e)) return;
    }
    if (typeof ssdNavHandleSelection_ === 'function') {
      if (ssdNavHandleSelection_(e)) return;
    }

    if (sh.getName() === SSD.sheets.input) {
      const a1 = e.range.getA1Notation();

      if (a1 === SSD.actionCells.save) {
        const msg = saveBookingValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 5);
        return;
      }
      if (a1 === SSD.actionCells.reset) {
        const msg = resetFormValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 4);
        return;
      }
      if (a1 === SSD.actionCells.conflict) {
        const msg = checkConflictValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 5);
        return;
      }
      if (a1 === SSD.actionCells.findCustomer) {
        const msg = findCustomerValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 5);
        return;
      }
      if (a1 === SSD.actionCells.invoice) {
        const msg = generateInvoiceValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 5);
        return;
      }
      if (isRangeInsideButton_(e.range, SSD.actionCells.revertStatus)) {
        const msg = revertBookingStatusValidated_();
        toastAndResetSelection_(sh, 'C6', msg, 5);
        return;
      }
    }

    if (sh.getName() === SSD.sheets.dashboardV2 && isRangeInsideButton_(e.range, SSD.actionCells.dashboardReset)) {
      const msg = resetDashboardV2Filters_();
      toastAndResetSelection_(sh, 'C5', msg, 4);
      return;
    }

    // MASTER TERAPIS -> tombol merged K2:L3
    if (sh.getName() === 'MASTER_TERAPIS' && isRangeInsideButton_(e.range, 'K2:L3')) {
      const msg = syncKomisiTerapisValidated_();
      toastAndResetSelection_(sh, 'A1', msg, 5);
      return;
    }

    // MASTER SA -> tombol merged H2:I3
    if (sh.getName() === 'MASTER_SA' && isRangeInsideButton_(e.range, 'H2:I3')) {
      const msg = syncKomisiSAValidated_();
      toastAndResetSelection_(sh, 'A1', msg, 5);
      return;
    }

  } catch (err) {
    const msg = (err && err.message) ? err.message : String(err);
    showToast_('SmartSpaDash : ' + msg, 'SmartSpaDash', 6);
  }
}

function isRangeInsideButton_(clickedRange, buttonA1) {
  const sh = clickedRange.getSheet();
  const btn = sh.getRange(buttonA1);

  const r1 = clickedRange.getRow();
  const c1 = clickedRange.getColumn();
  const r2 = r1 + clickedRange.getNumRows() - 1;
  const c2 = c1 + clickedRange.getNumColumns() - 1;

  const br1 = btn.getRow();
  const bc1 = btn.getColumn();
  const br2 = br1 + btn.getNumRows() - 1;
  const bc2 = bc1 + btn.getNumColumns() - 1;

  const overlap =
    !(r2 < br1 || r1 > br2 || c2 < bc1 || c1 > bc2);

  return overlap;
}

function showToast_(message, title, seconds) {
  var msg = String(message || '');
  if (!/^SmartSpaDash\s*:/i.test(msg)) msg = 'SmartSpaDash : ' + msg;
  var ttl = 'SmartSpaDash';
  var secs = Math.max(6, Number(seconds || 6));
  SpreadsheetApp.getActive().toast(msg, ttl, secs);
  SpreadsheetApp.flush();
  return msg;
}

function toastAndResetSelection_(sheet, targetA1, message, seconds) {
  var finalMsg = showToast_(message, 'SmartSpaDash', seconds || 6);
  try {
    if (sheet && sheet.getName && sheet.getName() === SSD.sheets.input && typeof setStatusMessage_ === 'function') {
      setStatusMessage_(finalMsg);
    }
  } catch (err) {}
  setSelectionGuard_(true);
  try {
    Utilities.sleep(1200);
    sheet.setActiveSelection(targetA1);
    SpreadsheetApp.flush();
  } finally {
    Utilities.sleep(200);
    setSelectionGuard_(false);
  }
}

function isSelectionGuardActive_() {
  return PropertiesService.getDocumentProperties().getProperty('SSD_SELECTION_GUARD') === '1';
}

function setSelectionGuard_(flag) {
  PropertiesService.getDocumentProperties().setProperty('SSD_SELECTION_GUARD', flag ? '1' : '0');
}

function saveBookingValidated_() {
  const form = getFormData_();
  validateForm_(form, 'save');

  const conflict = findConflict_(form);
  if (conflict) throw new Error('Bentrok dengan ' + conflict.booking_id);

  const bookingId = ssdPremiumSafeBuildBookingId_(form.bookingDate);
  const endTime = addMinutes_(form.startTime, form.duration);
  const paymentStatus = form.deposit <= 0 ? 'Belum Bayar' : (form.deposit >= form.grandTotal ? 'Lunas' : 'DP');

  const record = [
    bookingId,
    Utilities.formatDate(new Date(), SSD.timezone, 'yyyy-MM-dd HH:mm:ss'),
    formatBookingDate_(form.bookingDate),
    monthNameId_(form.bookingDate),
    form.startTime,
    endTime,
    form.bookingType,
    form.serviceName,
    form.category,
    form.therapistName,
    form.parentName,
    form.phone,
    form.childName,
    form.childAge,
    form.address,
    form.listPrice,
    form.promoName,
    form.discountValue,
    form.area,
    form.transportFee,
    form.grandTotal,
    form.deposit,
    form.remainingBalance,
    form.paymentMethod,
    paymentStatus,
    form.bookingStatus || 'Active',
    'PENDING',
    form.notes,
    form.therapistFee,
    form.adminFee,
    'Active',
    Number(Utilities.formatDate(new Date(form.bookingDate), SSD.timezone, 'M')),
    Number(Utilities.formatDate(new Date(form.bookingDate), SSD.timezone, 'yyyy'))
  ];

  const sh = getSheet_(SSD.sheets.log);
  sh.appendRow(record);
  ssdPremiumAppendSystemLog_({
  action: 'SAVE_BOOKING',
  bookingId: bookingId,
  targetSheet: sh.getName(),
  targetRow: sh.getLastRow(),
  newValue: JSON.stringify({
    booking_status: form.bookingStatus || 'Active',
    payment_status: paymentStatus,
    service_done_status: 'PENDING'
  }),
  details: 'Booking baru tersimpan dari INPUT BOOKING.'
});

  const lr = sh.getLastRow();
  sh.getRange(lr, 2).setNumberFormat('yyyy-MM-dd HH:mm:ss');
  sh.getRange(lr, 3).setNumberFormat('@');
  sh.getRange(lr, 5, 1, 2).setNumberFormat('HH:mm:ss');

  syncCustomerByIdentity_(form.parentName, form.phone, form.childName);
  refreshRekapHarian_();
  refreshRekapBulanan_();
  refreshPendingBookingList_();
  if (typeof refreshDashboardV2_ === 'function') refreshDashboardV2_();

  try {
    getSheet_(SSD.sheets.input).getRange(SSD.cells.bookingCode).setValue(bookingId).setNumberFormat('@');
  } catch (err) {}

  const msg = 'Booking berhasil disimpan: ' + bookingId;
  setStatusMessage_(msg);
  return msg;
}

function checkConflictValidated_() {
  const form = getFormData_();
  validateForm_(form, 'check');

  const conflict = findConflict_(form);
  const msg = conflict ? 'Bentrok dengan ' + conflict.booking_id : 'Aman. Tidak ada bentrok jadwal.';
  setStatusMessage_(msg);
  return msg;
}

function findCustomerValidated_() {
  const shIn = getSheet_(SSD.sheets.input);
  const searchKey = cleanText_(shIn.getRange(SSD.cells.searchKey).getDisplayValue());

  if (!searchKey) throw new Error('Isi No WA atau kode booking di F16 dulu.');

  clearSearchResultValidated_();

  const rows = getObjects_(getSheet_(SSD.sheets.log));
  const normalizedSearchPhone = normalizePhone_(searchKey);
  const normalizedSearchCode = searchKey.toUpperCase();

  const activeRows = rows.filter(r => cleanText_(r.record_status || 'Active') === 'Active');
  let found = null;

  if (normalizedSearchCode && /^BK-/i.test(normalizedSearchCode)) {
    found = activeRows.slice().reverse().find(r =>
      cleanText_(r.booking_id).toUpperCase() === normalizedSearchCode
    ) || null;
  } else {
    const matchedPhoneRows = activeRows.filter(r => normalizePhone_(r.phone) === normalizedSearchPhone);
    found = matchedPhoneRows.slice().reverse().find(r => {
      const paymentStatus = cleanText_(r.payment_status);
      const serviceDone = cleanText_(r.service_done_status).toUpperCase();
      return paymentStatus !== 'Lunas' && serviceDone !== 'DONE';
    }) || matchedPhoneRows.slice().reverse()[0] || null;
  }

  if (!found) {
    const msg = 'Data booking tidak ditemukan di TRANSAKSI_LOG.';
    setStatusMessage_(msg);
    return msg;
  }

  populateSearchResultsFromLogRow_(found);

  const msg = 'Booking ditemukan: ' + (found.booking_id || '-') + ' | WA: ' + (found.phone || '-');
  setStatusMessage_(msg);
  return msg;
}

function resetFormValidated_() {
  const sh = getSheet_(SSD.sheets.input);
  ['F6','C7','F7','C8','C9','F8','F9','C12','F12','C13','F13','C14','F14','C15','F15','C16','F16','J6']
    .forEach(a1 => sh.getRange(a1).clearContent());

  clearSearchResultValidated_();
  refreshPendingBookingList_();

  const msg = 'Form berhasil di-reset.';
  setStatusMessage_(msg);
  return msg;
}


function clearSearchResultValidated_() {
  const sh = getSheet_(SSD.sheets.input);
  sh.getRange('K6:K17').clearContent();
  sh.getRange('J16:J17').clearContent();
}

function computeDurationMinutesFromLogRow_(startValue, endValue) {
  if (!startValue || !endValue) return 0;
  const startMinutes = toMinutes_(startValue);
  const endMinutes = toMinutes_(endValue);
  if (isNaN(startMinutes) || isNaN(endMinutes)) return 0;
  return Math.max(0, endMinutes - startMinutes);
}


function clearPendingBookingListValidated_() {
  const sh = getSheet_(SSD.sheets.input);
  sh.getRange('B21:D200').clearContent();
}

function refreshPendingBookingList_() {
  const shInput = getSheet_(SSD.sheets.input);
  const logSh = getSheet_(SSD.sheets.log);
  clearPendingBookingListValidated_();

  const vals = logSh.getDataRange().getValues();
  if (!vals || vals.length < 2) return 0;

  const idx = indexMap_(vals[0]);
  const out = [];

  for (let i = 1; i < vals.length; i++) {
    const row = vals[i];
    const recordStatus = cleanText_(row[idx.record_status] || 'Active');
    const paymentStatus = cleanText_(row[idx.payment_status] || '');
    const serviceDone = cleanText_(row[idx.service_done_status] || '').toUpperCase();

    if (recordStatus !== 'Active') continue;
    if (paymentStatus === 'Lunas' && serviceDone === 'DONE') continue;

    out.push([
      row[idx.booking_id] || '',
      row[idx.payment_status] || '',
      row[idx.service_done_status] || ''
    ]);
  }

  if (out.length) {
    shInput.getRange(21, 2, out.length, 3).setValues(out);
  }
  return out.length;
}


function syncKomisiTerapisValidated_() {
  const shMaster = getSheet_('MASTER_TERAPIS');
  const shKomisi = getSheet_('KOMISI TERAPIS');

  // pakai display values agar aman untuk sheet hasil import / formula
  const values = shMaster.getDataRange().getDisplayValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const code = cleanText_(values[i][0]);     // kolom A
    const name = cleanText_(values[i][1]);     // kolom B
    const helper = cleanText_(values[i][8]);   // kolom I

    if (!helper || !code || !name) continue;
    rows.push([code, name]);
  }

  shKomisi.getRange('A5:G500').clearContent();

  if (!rows.length) {
    return 'SmartSpaDash : Tidak ada data terapis aktif dari MASTER_TERAPIS kolom I.';
  }

  const ts = Utilities.formatDate(new Date(), SSD.timezone, 'dd/MM/yyyy HH:mm');

  const out = rows.map((item, idx) => {
    const r = idx + 5;
    return [
      item[0],
      item[1],
      '=IFNA(COUNTIFS(TRANSAKSI_LOG!J:J;B' + r + ';TRANSAKSI_LOG!Y:Y;"lunas";TRANSAKSI_LOG!AA:AA;"Done";TRANSAKSI_LOG!AF:AF;$A$3;TRANSAKSI_LOG!AG:AG;$B$3);0)',
      '=IFNA(SUMIFS(TRANSAKSI_LOG!AC:AC;TRANSAKSI_LOG!J:J;B' + r + ';TRANSAKSI_LOG!Y:Y;"lunas";TRANSAKSI_LOG!AA:AA;"Done";TRANSAKSI_LOG!AF:AF;$A$3;TRANSAKSI_LOG!AG:AG;$B$3);0)',
      '=IFNA(SUMIFS(TRANSAKSI_LOG!T:T;TRANSAKSI_LOG!J:J;B' + r + ';TRANSAKSI_LOG!G:G;"Homecare";TRANSAKSI_LOG!Y:Y;"lunas";TRANSAKSI_LOG!AA:AA;"Done";TRANSAKSI_LOG!AF:AF;$A$3;TRANSAKSI_LOG!AG:AG;$B$3);0)',
      '=D' + r + '+E' + r,
      ts
    ];
  });

  shKomisi.getRange(5, 1, out.length, 7).setValues(out);
  SpreadsheetApp.flush();

  return 'SmartSpaDash : Data terapis berhasil ditarik ke KOMISI TERAPIS.';
}

function syncKomisiSAValidated_() {
  const shMaster = getSheet_('MASTER_SA');
  const shKomisi = getSheet_('KOMISI SA');

  // pakai display values agar aman untuk sheet hasil import / formula
  const values = shMaster.getDataRange().getDisplayValues();
  const rows = [];

  for (let i = 1; i < values.length; i++) {
    const code = cleanText_(values[i][0]);    // kolom A
    const name = cleanText_(values[i][1]);    // kolom B
    const helper = cleanText_(values[i][5]);  // kolom F

    if (!helper || !code || !name) continue;
    rows.push([code, name]);
  }

  shKomisi.getRange('A5:G500').clearContent();

  if (!rows.length) {
    return 'SmartSpaDash : Tidak ada data SA aktif dari MASTER_SA kolom F.';
  }

  const ts = Utilities.formatDate(new Date(), SSD.timezone, 'dd/MM/yyyy HH:mm');

  const out = rows.map((item, idx) => {
    const r = idx + 5;
    return [
      item[0],
      item[1],
      '=IFNA(COUNTIFS(TRANSAKSI_LOG!AE:AE;"Active";TRANSAKSI_LOG!Y:Y;"lunas";TRANSAKSI_LOG!AA:AA;"Done";TRANSAKSI_LOG!AF:AF;$A$3;TRANSAKSI_LOG!AG:AG;$B$3);0)',
      '=IFNA(SUMIFS(TRANSAKSI_LOG!AD:AD;TRANSAKSI_LOG!AE:AE;"Active";TRANSAKSI_LOG!Y:Y;"lunas";TRANSAKSI_LOG!AA:AA;"Done";TRANSAKSI_LOG!AF:AF;$A$3;TRANSAKSI_LOG!AG:AG;$B$3);0)',
      '=0',
      '=D' + r + '+E' + r,
      ts
    ];
  });

  shKomisi.getRange(5, 1, out.length, 7).setValues(out);
  SpreadsheetApp.flush();

  return 'SmartSpaDash : Data SA berhasil ditarik ke KOMISI SA.';
}

function markDoneValidated_() {
  return ssdPremiumUpdateServiceDoneValidated_();
}

function revertBookingStatusValidated_() {
  if (typeof ssdPremiumRequireRole_ !== 'function') throw new Error('SmartSpaDash : Modul login tidak tersedia.');
  ssdPremiumRequireRole_(['OWNER', 'ADMIN']);

  var executor = (typeof ssdPremiumWithLock_ === 'function')
    ? ssdPremiumWithLock_
    : function(fn) { return fn(); };

  return executor(function() {
    var inputSh = getSheet_(SSD.sheets.input);
    var bookingId = cleanText_(inputSh.getRange(SSD.cells.searchResultBookingCode).getDisplayValue()).toUpperCase();
    if (!bookingId) {
      throw new Error('SmartSpaDash : Lakukan pencarian booking dulu sampai hasil muncul di K6:K17.');
    }

    var rows = getObjects_(getSheet_(SSD.sheets.log));
    var rec = rows.slice().reverse().find(function(r) {
      return cleanText_(r.booking_id).toUpperCase() === bookingId &&
        cleanText_(r.record_status || 'Active') === 'Active';
    });
    if (!rec) throw new Error('SmartSpaDash : Booking ' + bookingId + ' tidak ditemukan di TRANSAKSI_LOG.');

    var sh = getSheet_(SSD.sheets.log);
    var vals = sh.getDataRange().getValues();
    var headers = vals[0];
    var idx = indexMap_(headers);
    var rowIndex = -1;

    for (var i = vals.length - 1; i >= 1; i--) {
      if (cleanText_(vals[i][idx.booking_id]).toUpperCase() === bookingId &&
          cleanText_(vals[i][idx.record_status] || 'Active') === 'Active') {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex < 2) throw new Error('SmartSpaDash : Row booking tidak ditemukan.');

    var grandTotal = parseNumber_(vals[rowIndex - 1][idx.grand_total] || 0);
    var deposit = parseNumber_(vals[rowIndex - 1][idx.deposit_amount] || 0);
    var oldPaymentStatus = vals[rowIndex - 1][idx.payment_status] || '';
    var oldBookingStatus = vals[rowIndex - 1][idx.booking_status] || '';
    var oldServiceDone = vals[rowIndex - 1][idx.service_done_status] || '';
    var oldRemaining = parseNumber_(vals[rowIndex - 1][idx.remaining_balance] || 0);
    var newRemaining = Math.max(0, grandTotal - deposit);

    sh.getRange(rowIndex, idx.payment_status + 1).setValue('Belum Bayar');
    sh.getRange(rowIndex, idx.service_done_status + 1).setValue('PENDING');
    sh.getRange(rowIndex, idx.booking_status + 1).setValue('Active');
    sh.getRange(rowIndex, idx.remaining_balance + 1).setValue(newRemaining);

    var refreshed = getObjects_(sh).slice().reverse().find(function(r) {
      return cleanText_(r.booking_id).toUpperCase() === bookingId &&
        cleanText_(r.record_status || 'Active') === 'Active';
    });
    if (refreshed) {
      populateSearchResultsFromLogRow_(refreshed);
    }

    if (typeof ssdPremiumAppendSystemLog_ === 'function') {
      ssdPremiumAppendSystemLog_({
        action: 'REVERT_STATUS',
        bookingId: bookingId,
        oldValue: JSON.stringify({
          payment_status: oldPaymentStatus,
          booking_status: oldBookingStatus,
          service_done_status: oldServiceDone,
          remaining_balance: oldRemaining
        }),
        newValue: JSON.stringify({
          payment_status: 'Belum Bayar',
          booking_status: 'Active',
          service_done_status: 'PENDING',
          remaining_balance: newRemaining
        }),
        targetSheet: sh.getName(),
        targetRow: rowIndex,
        details: 'Status booking dikembalikan dari hasil pencarian K6:K17.'
      });
    }

    refreshPendingBookingList_();
    if (typeof refreshRekapHarian_ === 'function') refreshRekapHarian_();
    if (typeof refreshRekapBulanan_ === 'function') refreshRekapBulanan_();
    if (typeof refreshDashboardV2_ === 'function') refreshDashboardV2_();

    var msg = 'SmartSpaDash : Booking ' + bookingId + ' dikembalikan ke Belum Bayar & PENDING.';
    setStatusMessage_(msg);
    return msg;
  });
}

function populateSearchResultsFromLogRow_(found) {
  var shIn = getSheet_(SSD.sheets.input);
  var durationMinutes = computeDurationMinutesFromLogRow_(found.start_time, found.end_time);

  shIn.getRange(SSD.cells.searchResultBookingCode).setValue(found.booking_id || '');
  shIn.getRange(SSD.cells.searchResultCategory).setValue(found.service_category || '');
  shIn.getRange(SSD.cells.searchResultDuration).setValue(durationMinutes || '');
  shIn.getRange(SSD.cells.searchResultPrice).setValue(parseNumber_(found.list_price || 0));
  shIn.getRange(SSD.cells.searchResultDiscountValue).setValue(parseNumber_(found.discount_value || 0));
  shIn.getRange(SSD.cells.searchResultTransportFee).setValue(parseNumber_(found.transport_fee || 0));
  shIn.getRange(SSD.cells.searchResultGrandTotal).setValue(parseNumber_(found.grand_total || 0));
  shIn.getRange(SSD.cells.searchResultTherapistFee).setValue(parseNumber_(found.therapist_fee || 0));
  shIn.getRange(SSD.cells.searchResultAdminFee).setValue(parseNumber_(found.admin_fee || 0));
  shIn.getRange(SSD.cells.searchResultRemainingBalance).setValue(parseNumber_(found.remaining_balance || 0));
  shIn.getRange(SSD.cells.searchResultParentName).setValue(found.parent_name || '');
  shIn.getRange(SSD.cells.searchResultBookingDate).setValue(found.booking_date || '');
  shIn.getRange(SSD.cells.mirrorParentName).setValue(found.parent_name || '');
  shIn.getRange(SSD.cells.mirrorBookingDate).setValue(found.booking_date || '');
  shIn.getRange(SSD.cells.bookingCode).setValue(found.booking_id || '').setNumberFormat('@');
}


function generateInvoiceValidated_() {
  const shInput = getSheet_(SSD.sheets.input);
  const bookingId = cleanText_(shInput.getRange(SSD.cells.searchResultBookingCode).getDisplayValue()).toUpperCase();

  if (!bookingId) {
    throw new Error('Lakukan pencarian booking dulu sampai hasil muncul di K6:K17 sebelum membuat invoice.');
  }

  const rows = getObjects_(getSheet_(SSD.sheets.log));
  const rec = rows.slice().reverse().find(r =>
    cleanText_(r.booking_id).toUpperCase() === bookingId &&
    cleanText_(r.record_status || 'Active') === 'Active'
  );

  if (!rec) throw new Error('Booking ' + bookingId + ' belum ditemukan di TRANSAKSI_LOG.');

  const inv = getSheet_(SSD.sheets.invoice);
  inv.getRange('C4').setValue('INV-' + rec.booking_id);
  inv.getRange('C5').setValue(new Date());
  inv.getRange('C6').setValue(rec.parent_name || '');
  inv.getRange('C7').setValue(rec.child_name || '');
  inv.getRange('C8').setValue(rec.booking_date || '');
  inv.getRange('C9').setValue(rec.start_time || '');
  inv.getRange('C10').setValue(rec.service_name || '');
  inv.getRange('C11').setValue(rec.therapist_name || '');
  inv.getRange('C12').setValue(rec.list_price || 0);
  inv.getRange('C13').setValue(rec.discount_value || 0);
  inv.getRange('C14').setValue(rec.transport_fee || 0);
  inv.getRange('C15').setValue(rec.grand_total || 0);
  inv.getRange('C16').setValue(rec.deposit_amount || 0);
  inv.getRange('C17').setValue(rec.remaining_balance || 0);
  inv.getRange('C18').setValue(rec.payment_status || '');

  SpreadsheetApp.getActive().setActiveSheet(inv);
  const msg = 'Invoice berhasil dibuat untuk ' + rec.booking_id;
  setStatusMessage_(msg);
  return msg;
}

function validateForm_(form, mode) {
  const errors = [];

  if (!form.bookingDate) errors.push('Tanggal booking wajib diisi.');
  if (!form.bookingType) errors.push('Jenis booking wajib dipilih.');
  if (!form.therapistName) errors.push('Terapis wajib dipilih.');
  if (!form.serviceName) errors.push('Layanan wajib dipilih.');
  if (!form.startTime) errors.push('Jam mulai wajib dipilih.');
  if (!form.parentName) errors.push('Nama ibu / ortu wajib diisi.');
  if (!form.phone) errors.push('No WhatsApp wajib diisi.');
  if (!form.childName) errors.push('Nama anak wajib diisi.');
  if (!form.childAge) errors.push('Usia anak wajib diisi.');
  if (!form.address) errors.push('Alamat wajib diisi.');
  if (form.childWeight === '' || form.childWeight === null) errors.push('Berat anak wajib diisi.');
  if (!form.paymentMethod && mode === 'save') errors.push('Metode bayar wajib dipilih.');
  if (form.bookingType === 'Homecare' && !form.area) errors.push('Area homecare wajib dipilih.');

  if (errors.length) throw new Error(errors.join(' '));
}


function isQualifiedClosedBooking_(r) {
  return (
    cleanText_(r.record_status || 'Active') === 'Active' &&
    cleanText_(r.payment_status).toLowerCase() === 'lunas' &&
    cleanText_(r.service_done_status).toUpperCase() === 'DONE'
  );
}

function refreshRekapHarian_() {
  const sh = getSheet_('REKAP HARIAN');
  const monthValue = parseNumber_(sh.getRange('B3').getValue());
  const yearValue = parseNumber_(sh.getRange('C3').getValue());
  const useFilter = monthValue >= 1 && monthValue <= 12 && yearValue >= 1900;

  const rows = getObjects_(getSheet_(SSD.sheets.log)).filter(r =>
    isQualifiedClosedBooking_(r) &&
    parseDateValue_(r.booking_date)
  );

  const startRow = 5;
  const clearRows = Math.max(0, sh.getMaxRows() - startRow + 1);
  if (clearRows > 0) sh.getRange(startRow, 1, clearRows, 8).clearContent();
  if (!rows.length) return;

  const summaryMap = buildDailySummaryMap_(rows);
  const dates = buildRekapDateList_(rows, monthValue, yearValue, useFilter);
  if (!dates.length) return;

  const timestamp = Utilities.formatDate(new Date(), SSD.timezone, 'dd/MM/yyyy HH:mm');
  const output = dates.map(d => {
    const key = Utilities.formatDate(d, SSD.timezone, 'yyyy-MM-dd');
    const s = summaryMap[key] || {
      total_booking: 0,
      homecare_count: 0,
      inhouse_count: 0,
      gross_revenue: 0,
      pending_payment: 0,
      done_count: 0
    };

    return [
      formatBookingDateForSheet_(d),
      s.total_booking,
      s.homecare_count,
      s.inhouse_count,
      s.gross_revenue,
      s.pending_payment,
      s.done_count,
      timestamp
    ];
  });

  sh.getRange(startRow, 1, output.length, 8).setValues(output);
  sh.getRange(startRow, 1, output.length, 1).setNumberFormat('@');
  sh.getRange(startRow, 2, output.length, 3).setNumberFormat('0');
  sh.getRange(startRow, 5, output.length, 2).setNumberFormat('#,##0');
  sh.getRange(startRow, 7, output.length, 1).setNumberFormat('0');
  sh.getRange(startRow, 8, output.length, 1).setNumberFormat('@');
}

function buildDailySummaryMap_(rows) {
  const map = {};

  rows.forEach(r => {
    if (!isQualifiedClosedBooking_(r)) return;

    const d = parseDateValue_(r.booking_date);
    if (!d) return;

    const key = Utilities.formatDate(d, SSD.timezone, 'yyyy-MM-dd');
    if (!map[key]) {
      map[key] = {
        total_booking: 0,
        homecare_count: 0,
        inhouse_count: 0,
        gross_revenue: 0,
        pending_payment: 0,
        done_count: 0
      };
    }

    const s = map[key];
    const bookingType = cleanText_(r.booking_type).toLowerCase();

    s.total_booking += 1;
    if (bookingType === 'homecare') s.homecare_count += 1;
    if (bookingType === 'inhouse') s.inhouse_count += 1;
    s.gross_revenue += parseNumber_(r.grand_total);
    s.pending_payment += 0;
    s.done_count += 1;
  });

  return map;
}

function buildRekapDateList_(rows, monthValue, yearValue, useFilter) {
  const parsedDates = rows
    .map(r => parseDateValue_(r.booking_date))
    .filter(Boolean)
    .sort((a, b) => a.getTime() - b.getTime());

  if (!parsedDates.length) return [];

  if (useFilter) {
    return enumerateDates_(
      new Date(yearValue, monthValue - 1, 1),
      new Date(yearValue, monthValue, 0)
    );
  }

  return enumerateDates_(parsedDates[0], parsedDates[parsedDates.length - 1]);
}

function enumerateDates_(startDate, endDate) {
  const dates = [];
  const start = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
  const end = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

  for (let d = new Date(start); d.getTime() <= end.getTime(); d.setDate(d.getDate() + 1)) {
    dates.push(new Date(d.getFullYear(), d.getMonth(), d.getDate()));
  }

  return dates;
}

function refreshRekapBulanan_() {
  const sh = getSheet_('REKAP BULANAN');
  const yearValue = parseNumber_(sh.getRange('B3').getValue());
  const useFilter = yearValue >= 1900;

  const rows = getObjects_(getSheet_(SSD.sheets.log)).filter(r =>
    isQualifiedClosedBooking_(r) &&
    parseDateValue_(r.booking_date)
  );

  const startRow = 5;
  const clearRows = Math.max(0, sh.getMaxRows() - startRow + 1);
  if (clearRows > 0) sh.getRange(startRow, 1, clearRows, 15).clearContent();
  if (!rows.length) return;

  const summaryMap = buildMonthlySummaryMap_(rows);
  const periods = buildRekapBulananPeriods_(rows, yearValue, useFilter);
  if (!periods.length) return;

  const timestamp = Utilities.formatDate(new Date(), SSD.timezone, 'dd/MM/yyyy HH:mm');
  const monthNames = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];

  const output = periods.map(p => {
    const key = p.year + '-' + Utilities.formatString('%02d', p.month);
    const s = summaryMap[key] || {
      total_booking: 0,
      homecare_count: 0,
      inhouse_count: 0,
      gross_revenue: 0,
      pending_payment: 0,
      done_count: 0,
      avg_ticket: 0,
      top_therapist: '',
      top_therapist_qty: 0,
      jumlah_layanan: 0,
      layanan_paling_banyak: '',
      jumlah_layanan_terbanyak: 0
    };

    return [
      monthNames[p.month - 1],
      p.year,
      s.total_booking,
      s.homecare_count,
      s.inhouse_count,
      s.gross_revenue,
      s.pending_payment,
      s.done_count,
      s.avg_ticket,
      s.top_therapist,
      s.top_therapist_qty,
      s.jumlah_layanan,
      s.layanan_paling_banyak,
      s.jumlah_layanan_terbanyak,
      timestamp
    ];
  });

  sh.getRange(startRow, 1, output.length, 15).setValues(output);
  sh.getRange(startRow, 2, output.length, 4).setNumberFormat('0');
  sh.getRange(startRow, 6, output.length, 2).setNumberFormat('#,##0');
  sh.getRange(startRow, 8, output.length, 1).setNumberFormat('0');
  sh.getRange(startRow, 9, output.length, 1).setNumberFormat('#,##0');
  sh.getRange(startRow, 11, output.length, 1).setNumberFormat('0');
  sh.getRange(startRow, 12, output.length, 1).setNumberFormat('0');
  sh.getRange(startRow, 14, output.length, 1).setNumberFormat('0');
  sh.getRange(startRow, 15, output.length, 1).setNumberFormat('@');
}

function buildMonthlySummaryMap_(rows) {
  const map = {};

  rows.forEach(r => {
    if (!isQualifiedClosedBooking_(r)) return;

    const d = parseDateValue_(r.booking_date);
    if (!d) return;

    const month = d.getMonth() + 1;
    const year = d.getFullYear();
    const key = year + '-' + Utilities.formatString('%02d', month);

    if (!map[key]) {
      map[key] = {
        total_booking: 0,
        homecare_count: 0,
        inhouse_count: 0,
        gross_revenue: 0,
        pending_payment: 0,
        done_count: 0,
        therapistCounts: {},
        serviceCounts: {}
      };
    }

    const s = map[key];
    const bookingType = cleanText_(r.booking_type).toLowerCase();
    const therapist = cleanText_(r.therapist_name);
    const service = cleanText_(r.service_name);

    s.total_booking += 1;
    if (bookingType === 'homecare') s.homecare_count += 1;
    if (bookingType === 'inhouse') s.inhouse_count += 1;
    s.gross_revenue += parseNumber_(r.grand_total);
    s.pending_payment += 0;
    s.done_count += 1;

    if (therapist) s.therapistCounts[therapist] = (s.therapistCounts[therapist] || 0) + 1;
    if (service) s.serviceCounts[service] = (s.serviceCounts[service] || 0) + 1;
  });

  Object.keys(map).forEach(key => {
    const s = map[key];
    const topTherapist = getTopLabelAndCount_(s.therapistCounts);
    const topService = getTopLabelAndCount_(s.serviceCounts);

    s.avg_ticket = s.total_booking ? Math.round(s.gross_revenue / s.total_booking) : 0;
    s.top_therapist = topTherapist.label;
    s.top_therapist_qty = topTherapist.count;
    s.jumlah_layanan = Object.keys(s.serviceCounts).length;
    s.layanan_paling_banyak = topService.label;
    s.jumlah_layanan_terbanyak = topService.count;
  });

  return map;
}

function getTopLabelAndCount_(obj) {
  const entries = Object.keys(obj).map(k => ({ label: k, count: obj[k] }));
  if (!entries.length) return { label: '', count: 0 };

  entries.sort((a, b) => {
    if (b.count !== a.count) return b.count - a.count;
    return a.label.localeCompare(b.label);
  });

  return entries[0];
}

function buildRekapBulananPeriods_(rows, yearValue, useFilter) {
  const dates = rows
    .map(r => parseDateValue_(r.booking_date))
    .filter(Boolean)
    .sort((a, b) => a.getTime() - b.getTime());

  if (!dates.length) return [];

  const periods = [];

  if (useFilter) {
    for (let month = 1; month <= 12; month++) periods.push({ year: yearValue, month: month });
    return periods;
  }

  const minYear = dates[0].getFullYear();
  const maxYear = dates[dates.length - 1].getFullYear();

  for (let year = minYear; year <= maxYear; year++) {
    for (let month = 1; month <= 12; month++) {
      periods.push({ year: year, month: month });
    }
  }

  return periods;
}

function getFormData_() {
  const sh = getSheet_(SSD.sheets.input);
  const durationRaw = sh.getRange(SSD.cells.duration).getValue();
  let durationInMinutes = 0;

  if (durationRaw instanceof Date) {
    durationInMinutes = (durationRaw.getHours() * 60) + durationRaw.getMinutes();
  } else {
    durationInMinutes = parseNumber_(durationRaw);
  }

  return {
    bookingDate: sh.getRange(SSD.cells.bookingDate).getValue(),
    bookingType: cleanText_(sh.getRange(SSD.cells.bookingType).getDisplayValue()),
    therapistName: cleanText_(sh.getRange(SSD.cells.therapist).getDisplayValue()),
    bookingStatus: cleanText_(sh.getRange(SSD.cells.bookingStatus).getDisplayValue()),
    serviceName: cleanText_(sh.getRange(SSD.cells.service).getDisplayValue()),
    startTime: sh.getRange(SSD.cells.startTime).getValue(),
    area: cleanText_(sh.getRange(SSD.cells.area).getDisplayValue()),
    promoName: cleanText_(sh.getRange(SSD.cells.promo).getDisplayValue()),
    parentName: cleanText_(sh.getRange(SSD.cells.parentName).getValue()),
    phone: normalizePhone_(sh.getRange(SSD.cells.phone).getValue()),
    childName: cleanText_(sh.getRange(SSD.cells.childName).getValue()),
    childAge: cleanText_(sh.getRange(SSD.cells.childAge).getValue()),
    address: cleanText_(sh.getRange(SSD.cells.address).getValue()),
    childWeight: sh.getRange(SSD.cells.childWeight).getValue(),
    notes: cleanText_(sh.getRange(SSD.cells.notes).getValue()),
    paymentMethod: cleanText_(sh.getRange(SSD.cells.paymentMethod).getDisplayValue()),
    deposit: parseNumber_(sh.getRange(SSD.cells.deposit).getValue()),
    category: cleanText_(sh.getRange(SSD.cells.category).getDisplayValue()),
    duration: durationInMinutes,
    listPrice: parseNumber_(sh.getRange(SSD.cells.price).getValue()),
    discountValue: parseNumber_(sh.getRange(SSD.cells.discountValue).getValue()),
    transportFee: parseNumber_(sh.getRange(SSD.cells.transportFee).getValue()),
    grandTotal: parseNumber_(sh.getRange(SSD.cells.grandTotal).getValue()),
    therapistFee: parseNumber_(sh.getRange(SSD.cells.therapistFee).getValue()),
    adminFee: parseNumber_(sh.getRange(SSD.cells.adminFee).getValue()),
    remainingBalance: parseNumber_(sh.getRange(SSD.cells.remainingBalance).getValue())
  };
}

function findConflict_(form) {
  const rows = getObjects_(getSheet_(SSD.sheets.log));
  const targetStart = toMinutes_(form.startTime);
  const targetEnd = targetStart + Number(form.duration || 0);

  for (const r of rows) {
    if (cleanText_(r.record_status) !== 'Active') continue;
    if (!sameDate_(r.booking_date, form.bookingDate)) continue;
    if (cleanText_(r.therapist_name) !== form.therapistName) continue;

    const rowStart = toMinutes_(r.start_time);
    const rowEnd = toMinutes_(r.end_time);
    if (targetStart < rowEnd && targetEnd > rowStart) return r;
  }

  return null;
}

function syncCustomerByIdentity_(parentName, phone, childName) {
  parentName = cleanText_(parentName);
  phone = normalizePhone_(phone);
  childName = cleanText_(childName);

  const shLog = getSheet_(SSD.sheets.log);
  const logRows = getObjects_(shLog).filter(r =>
    normalizePhone_(r.phone) === phone &&
    cleanText_(r.parent_name) === parentName &&
    cleanText_(r.child_name) === childName &&
    cleanText_(r.record_status || 'Active') === 'Active'
  );

  if (!logRows.length) return;

  let totalSpending = 0;
  let latest = null;

  logRows.forEach(r => {
    totalSpending += parseNumber_(r.grand_total);
    if (!latest || compareBookingMoment_(r, latest) > 0) latest = r;
  });

  const visitCount = logRows.length;
  const avgPerVisit = visitCount ? totalSpending / visitCount : 0;
  const customerType = visitCount === 1 ? 'Baru' : (visitCount <= 10 ? 'Repeat' : 'VIP');

  const sh = getSheet_(SSD.sheets.pelanggan);
  const vals = sh.getDataRange().getValues();
  if (!vals.length) return;

  const idx = indexMap_(vals[0]);

  let foundRow = -1;
  for (let i = 1; i < vals.length; i++) {
    if (
      normalizePhone_(vals[i][idx.phone]) === phone &&
      cleanText_(vals[i][idx.parent_name]) === parentName &&
      cleanText_(vals[i][idx.child_name]) === childName
    ) {
      foundRow = i + 1;
      break;
    }
  }

  if (foundRow === -1) foundRow = Math.max(2, sh.getLastRow() + 1);

  const customerIdCell = sh.getRange(foundRow, idx.customer_id + 1);
  if (!cleanText_(customerIdCell.getValue())) {
    customerIdCell.setValue(nextCustomerId_(sh, idx.customer_id));
  }
  customerIdCell.setNumberFormat('@');

  sh.getRange(foundRow, idx.parent_name + 1).setValue(latest.parent_name || parentName);
  sh.getRange(foundRow, idx.phone + 1).setValue(phone);
  sh.getRange(foundRow, idx.address + 1).setValue(latest.address || '');
  sh.getRange(foundRow, idx.child_name + 1).setValue(latest.child_name || childName);
  sh.getRange(foundRow, idx.child_age + 1).setValue(latest.child_age || '');

  if (idx.visit_count != null) sh.getRange(foundRow, idx.visit_count + 1).setValue(visitCount);
  if (idx.total_spending != null) sh.getRange(foundRow, idx.total_spending + 1).setValue(totalSpending);
  if (idx.avg_per_visit != null) sh.getRange(foundRow, idx.avg_per_visit + 1).setValue(avgPerVisit);
  if (idx.last_visit != null) sh.getRange(foundRow, idx.last_visit + 1).setValue(formatBookingDateForSheet_(latest.booking_date));
  if (idx.customer_type != null) sh.getRange(foundRow, idx.customer_type + 1).setValue(customerType);
  if (idx.favorite_service != null) sh.getRange(foundRow, idx.favorite_service + 1).setValue(latest.service_name || '');
}

function nextCustomerId_(sh, customerIdIndexZeroBased) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 'L0001';

  const values = sh.getRange(2, customerIdIndexZeroBased + 1, lastRow - 1, 1).getDisplayValues().flat();
  let maxSeq = 0;

  values.forEach(v => {
    const m = cleanText_(v).match(/^L(\d+)$/i);
    if (m) maxSeq = Math.max(maxSeq, Number(m[1]));
  });

  return 'L' + Utilities.formatString('%04d', maxSeq + 1);
}

function compareBookingMoment_(a, b) {
  const da = parseDateValue_(a.booking_date);
  const db = parseDateValue_(b.booking_date);
  const dayA = da ? da.getTime() : 0;
  const dayB = db ? db.getTime() : 0;
  if (dayA !== dayB) return dayA - dayB;
  return toMinutes_(a.start_time) - toMinutes_(b.start_time);
}

function getObjects_(sh) {
  const vals = sh.getDataRange().getValues();
  if (vals.length < 2) return [];

  const headers = vals[0];
  return vals.slice(1)
    .filter(r => r.some(v => v !== '' && v != null))
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => obj[String(h).trim()] = r[i]);
      return obj;
    });
}

function getSheet_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh) throw new Error('Sheet tidak ditemukan: ' + name);
  return sh;
}

function setStatusMessage_(msg) {
  getSheet_(SSD.sheets.input).getRange(SSD.cells.statusMessage).setValue(msg);
}

function getBookingCodeInput_() {
  const raw = getSheet_(SSD.sheets.input).getRange(SSD.cells.bookingCode).getDisplayValue();
  return cleanText_(raw).toUpperCase();
}

function buildBookingId_(bookingDate) {
  const sh = getSheet_(SSD.sheets.log);
  const seq = Math.max(1, sh.getLastRow());
  return 'SSD-' + Utilities.formatDate(new Date(bookingDate), SSD.timezone, 'yyMMdd') + '-' + ('000' + seq).slice(-3);
}

function formatBookingDate_(d) {
  return Utilities.formatDate(new Date(d), SSD.timezone, 'd-M-yyyy');
}

function formatBookingDateForSheet_(v) {
  const d = parseDateValue_(v);
  return d ? Utilities.formatDate(d, SSD.timezone, 'd-M-yyyy') : '';
}

function monthNameId_(d) {
  return ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'][new Date(d).getMonth()];
}

function parseDateValue_(v) {
  if (v instanceof Date && !isNaN(v)) {
    return new Date(v.getFullYear(), v.getMonth(), v.getDate());
  }

  const s = cleanText_(v);
  if (!s) return null;

  let m = s.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));

  const iso = new Date(v);
  if (!isNaN(iso)) return new Date(iso.getFullYear(), iso.getMonth(), iso.getDate());

  return null;
}

function addMinutes_(timeVal, mins) {
  if (!timeVal || isNaN(new Date(timeVal).getTime())) return timeVal;
  const d = new Date(timeVal.getTime());
  d.setMinutes(d.getMinutes() + Number(mins || 0));
  return d;
}

function indexMap_(headers) {
  const m = {};
  headers.forEach((h, i) => m[String(h).trim()] = i);
  return m;
}

function cleanText_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
}

function parseNumber_(v) {
  if (v === '' || v == null) return 0;
  if (typeof v === 'number') return v;
  return Number(String(v).replace(/[^\d.-]/g, '')) || 0;
}

function normalizePhone_(v) {
  let s = String(v == null ? '' : v).replace(/\D/g, '');
  if (!s) return '';
  if (s.startsWith('62')) return s;
  if (s.startsWith('0')) return '62' + s.slice(1);
  return s;
}

function sameDate_(a, b) {
  const da = parseDateValue_(a);
  const db = parseDateValue_(b);
  return !!(da && db && da.getTime() === db.getTime());
}

function sameTime_(a, b) {
  return toMinutes_(a) === toMinutes_(b);
}

function toMinutes_(v) {
  if (v instanceof Date && !isNaN(v)) return v.getHours() * 60 + v.getMinutes();
  const d = new Date(v);
  if (!isNaN(d)) return d.getHours() * 60 + d.getMinutes();
  const m = String(v).match(/(\d{1,2}):(\d{2})/);
  return m ? Number(m[1]) * 60 + Number(m[2]) : 0;
}

function monthNameToNumber_(monthName) {
  const map = {
    'Januari': 1, 'Februari': 2, 'Maret': 3, 'April': 4,
    'Mei': 5, 'Juni': 6, 'Juli': 7, 'Agustus': 8,
    'September': 9, 'Oktober': 10, 'November': 11, 'Desember': 12
  };
  return map[cleanText_(monthName)] || 0;
}

function ensureDashboardV2SupportSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SSD.sheets.dashboardV2Data);
  if (!sh) {
    sh = ss.insertSheet(SSD.sheets.dashboardV2Data);
  }
  sh.hideSheet();
  return sh;
}

function getDashboardV2Filters_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.dashboardV2);
  if (!sh) return null;

  const yearText = cleanText_(sh.getRange('C5').getDisplayValue());
  const monthText = cleanText_(sh.getRange('F5').getDisplayValue());
  const dayText = cleanText_(sh.getRange('I5').getDisplayValue());

  return {
    yearText: yearText || 'Semua Tahun',
    monthText: monthText || 'Semua Bulan',
    dayText: dayText || 'Semua Hari',
    yearNum: yearText === 'Semua Tahun' ? 0 : parseNumber_(yearText),
    monthNum: monthText === 'Semua Bulan' ? 0 : monthNameToNumber_(monthText),
    dayNum: dayText === 'Semua Hari' ? 0 : parseNumber_(dayText)
  };
}

function matchesDashboardV2Filters_(r, filters) {
  if (!isQualifiedClosedBooking_(r)) return false;
  const d = parseDateValue_(r.booking_date);
  if (!d) return false;

  if (filters.yearNum && d.getFullYear() !== filters.yearNum) return false;
  if (filters.monthNum && (d.getMonth() + 1) !== filters.monthNum) return false;
  if (filters.dayNum && d.getDate() !== filters.dayNum) return false;
  return true;
}

function getDashboardV2PeriodLabel_(dateObj, filters) {
  if (!dateObj) return '';
  if (filters.dayNum) {
    return Utilities.formatDate(dateObj, SSD.timezone, 'dd/MM/yyyy');
  }
  if (filters.monthNum) {
    return Utilities.formatDate(dateObj, SSD.timezone, 'dd/MM');
  }
  if (filters.yearNum) {
    return Utilities.formatDate(dateObj, SSD.timezone, 'MMM');
  }
  return Utilities.formatDate(dateObj, SSD.timezone, 'yyyy-MM');
}

function buildDashboardV2Data_() {
  const filters = getDashboardV2Filters_();
  if (!filters) return null;

  const rows = getObjects_(getSheet_(SSD.sheets.log));
  const filtered = rows.filter(r => matchesDashboardV2Filters_(r, filters));

  const topMap = {};
  const omzetTrend = {};
  const bookingTrend = {};
  const typeMap = { Homecare: 0, Inhouse: 0 };
  const allYears = [];

  rows.forEach(r => {
    const d = parseDateValue_(r.booking_date);
    if (d) allYears.push(d.getFullYear());
  });

  filtered.forEach(r => {
    const service = cleanText_(r.service_name);
    const category = cleanText_(r.service_category);
    const bookingType = cleanText_(r.booking_type);
    const total = parseNumber_(r.grand_total);
    const d = parseDateValue_(r.booking_date);
    const label = getDashboardV2PeriodLabel_(d, filters);

    if (!topMap[service]) {
      topMap[service] = { count: 0, categoryCounts: {}, typeCounts: {} };
    }
    topMap[service].count += 1;
    if (category) topMap[service].categoryCounts[category] = (topMap[service].categoryCounts[category] || 0) + 1;
    if (bookingType) topMap[service].typeCounts[bookingType] = (topMap[service].typeCounts[bookingType] || 0) + 1;

    omzetTrend[label] = (omzetTrend[label] || 0) + total;
    bookingTrend[label] = (bookingTrend[label] || 0) + 1;

    if (bookingType === 'Homecare') typeMap.Homecare += 1;
    if (bookingType === 'Inhouse') typeMap.Inhouse += 1;
  });

  const top10 = Object.keys(topMap)
    .map(name => {
      const item = topMap[name];
      const bestCategory = Object.keys(item.categoryCounts).sort((a, b) => {
        if (item.categoryCounts[b] !== item.categoryCounts[a]) return item.categoryCounts[b] - item.categoryCounts[a];
        return a.localeCompare(b);
      })[0] || '';
      const bestType = Object.keys(item.typeCounts).sort((a, b) => {
        if (item.typeCounts[b] !== item.typeCounts[a]) return item.typeCounts[b] - item.typeCounts[a];
        return a.localeCompare(b);
      })[0] || '';

      return [0, name, bestCategory, bestType, item.count];
    })
    .sort((a, b) => {
      if (b[4] !== a[4]) return b[4] - a[4];
      return a[1].localeCompare(b[1]);
    })
    .slice(0, 10)
    .map((row, idx) => [idx + 1, row[1], row[2], row[3], row[4]]);

  const periodLabels = Object.keys(omzetTrend).sort((a, b) => {
    return a.localeCompare(b);
  });

  const omzetRows = periodLabels.map(label => [label, omzetTrend[label] || 0]);
  const bookingRows = periodLabels.map(label => [label, bookingTrend[label] || 0]);
  const typeRows = [['Homecare', typeMap.Homecare], ['Inhouse', typeMap.Inhouse]];

  const totalBooking = filtered.length;
  const omzetClosed = filtered.reduce((sum, r) => sum + parseNumber_(r.grand_total), 0);
  const avgTicket = totalBooking ? Math.round(omzetClosed / totalBooking) : 0;

  return {
    filters: filters,
    top10: top10,
    omzetRows: omzetRows,
    bookingRows: bookingRows,
    typeRows: typeRows,
    metrics: {
      totalBooking: totalBooking,
      omzetClosed: omzetClosed,
      homecareClosed: typeMap.Homecare,
      inhouseClosed: typeMap.Inhouse,
      avgTicket: avgTicket
    },
    allYears: Array.from(new Set(allYears)).sort()
  };
}

function syncDashboardV2Top10Formulas_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.dashboardV2);
  if (!sh) return;

  for (var r = 20; r <= 29; r++) {
    var helperRow = r - 18;
    sh.getRange(r, 5).setFormula("=IF('DASHBOARD V2 DATA'!$A$" + helperRow + "=\"\" ;\"\" ;'DASHBOARD V2 DATA'!$A$" + helperRow + ")");
    sh.getRange(r, 6).setFormula("=IF($E" + r + "=\"\" ;\"\" ;'DASHBOARD V2 DATA'!$B$" + helperRow + ")");
    sh.getRange(r, 7).setFormula("=IF($E" + r + "=\"\" ;\"\" ;'DASHBOARD V2 DATA'!$C$" + helperRow + ")");
    sh.getRange(r, 8).setFormula("=IF($E" + r + "=\"\" ;\"\" ;'DASHBOARD V2 DATA'!$D$" + helperRow + ")");
    sh.getRange(r, 9).setFormula("=IF($E" + r + "=\"\" ;\"\" ;'DASHBOARD V2 DATA'!$E$" + helperRow + ")");
  }
}

function setupDashboardV2Validation_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SSD.sheets.dashboardV2);
  const helper = ss.getSheetByName(SSD.sheets.dashboardV2Data);
  if (!dash || !helper) return;

  const yearLastRow = Math.max(2, helper.getLastRow());
  const yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(helper.getRange('Y2:Y' + yearLastRow), true)
    .setAllowInvalid(false)
    .build();

  const monthRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Semua Bulan','Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'], true)
    .setAllowInvalid(false)
    .build();

  const dayList = ['Semua Hari'];
  for (var i = 1; i <= 31; i++) dayList.push(String(i));
  const dayRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(dayList, true)
    .setAllowInvalid(false)
    .build();

  dash.getRange('C5').setDataValidation(yearRule);
  dash.getRange('F5').setDataValidation(monthRule);
  dash.getRange('I5').setDataValidation(dayRule);
}

function rebuildDashboardV2Charts_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SSD.sheets.dashboardV2);
  const helper = ss.getSheetByName(SSD.sheets.dashboardV2Data);
  if (!dash || !helper) return;

  const charts = dash.getCharts();
  charts.forEach(function(ch) { dash.removeChart(ch); });

  const omzetLastRow = Math.max(2, helper.getRange('G2:G').getDisplayValues().flat().filter(String).length + 1);
  const bookingLastRow = Math.max(2, helper.getRange('K2:K').getDisplayValues().flat().filter(String).length + 1);

  const omzetChart = dash.newChart()
    .asLineChart()
    .addRange(helper.getRange('G1:H' + omzetLastRow))
    .setPosition(33, 2, 0, 0)
    .setOption('title', 'Trend Omzet Closed')
    .setOption('legend', { position: 'none' })
    .setOption('width', 650)
    .setOption('height', 280)
    .build();

  const bookingChart = dash.newChart()
    .asColumnChart()
    .addRange(helper.getRange('K1:L' + bookingLastRow))
    .setPosition(33, 10, 0, 0)
    .setOption('title', 'Trend Booking Closed')
    .setOption('legend', { position: 'none' })
    .setOption('width', 650)
    .setOption('height', 280)
    .build();

  const typeChart = dash.newChart()
    .asPieChart()
    .addRange(helper.getRange('O1:P3'))
    .setPosition(33, 18, 0, 0)
    .setOption('title', 'Homecare vs Inhouse')
    .setOption('pieHole', 0.4)
    .setOption('width', 420)
    .setOption('height', 280)
    .build();

  dash.insertChart(omzetChart);
  dash.insertChart(bookingChart);
  dash.insertChart(typeChart);
}

function refreshDashboardV2_() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SSD.sheets.dashboardV2);
  if (!dash) return 'Dashboard V2 belum tersedia.';

  const helper = ensureDashboardV2SupportSheet_();
  const built = buildDashboardV2Data_();
  if (!built) return 'Dashboard V2 belum tersedia.';

  helper.clearContents();

  helper.getRange('A1:E1').setValues([['Rank','Layanan','Category','Inhouse/Homecare','Jumlah Booking']]);
  helper.getRange('G1:H1').setValues([['Periode','Omzet Closed']]);
  helper.getRange('K1:L1').setValues([['Periode','Booking Closed']]);
  helper.getRange('O1:P1').setValues([['Jenis','Jumlah']]);
  helper.getRange('S1:W1').setValues([['Total Booking Closed','Omzet Closed','Homecare Closed','Inhouse Closed','Avg Ticket Closed']]);
  helper.getRange('Y1').setValue('Year Filter');

  if (built.top10.length) helper.getRange(2, 1, built.top10.length, 5).setValues(built.top10);
  if (built.omzetRows.length) helper.getRange(2, 7, built.omzetRows.length, 2).setValues(built.omzetRows);
  if (built.bookingRows.length) helper.getRange(2, 11, built.bookingRows.length, 2).setValues(built.bookingRows);
  helper.getRange(2, 15, built.typeRows.length, 2).setValues(built.typeRows);
  helper.getRange('S2:W2').setValues([[built.metrics.totalBooking, built.metrics.omzetClosed, built.metrics.homecareClosed, built.metrics.inhouseClosed, built.metrics.avgTicket]]);

  const yearValues = [['Semua Tahun']].concat(built.allYears.map(function(y) { return [y]; }));
  helper.getRange(2, 25, yearValues.length, 1).setValues(yearValues);

  syncDashboardV2Top10Formulas_();
  setupDashboardV2Validation_();
  rebuildDashboardV2Charts_();

  dash.getRange('B8').setFormula("='DASHBOARD V2 DATA'!$S$2");
  dash.getRange('E8').setFormula("='DASHBOARD V2 DATA'!$T$2");
  dash.getRange('H8').setFormula("='DASHBOARD V2 DATA'!$U$2");
  dash.getRange('B13').setFormula("='DASHBOARD V2 DATA'!$V$2");
  dash.getRange('E13').setFormula("='DASHBOARD V2 DATA'!$W$2");
  dash.getRange('H13').setFormula('=IF(C5="Semua Tahun";"Semua Tahun";C5)&" | "&IF(F5="Semua Bulan";"Semua Bulan";F5)&" | "&IF(I5="Semua Hari";"Semua Hari";I5)');

  return 'Dashboard V2 berhasil diperbarui.';
}

function resetDashboardV2Filters_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SSD.sheets.dashboardV2);
  if (!sh) throw new Error('Sheet DASHBOARD V2 tidak ditemukan.');

  sh.getRange('C5').setValue('Semua Tahun');
  sh.getRange('F5').setValue('Semua Bulan');
  sh.getRange('I5').setValue('Semua Hari');

  return refreshDashboardV2_();
}
