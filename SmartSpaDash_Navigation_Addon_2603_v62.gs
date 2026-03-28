
var SSD_NAV = {
  sheets: ['LOGIN','DASHBOARD V2','INPUT BOOKING','CRM PELANGGAN','TRANSAKSI_LOG','REKAP HARIAN','REKAP BULANAN','INVOICE','JADWAL HARIAN'],
  buttons: {
    'LOGIN': [
      {range:'B1:C2', target:'DASHBOARD V2'},
      {range:'D1:E2', target:'INPUT BOOKING'},
      {range:'F1:G2', target:'TRANSAKSI_LOG'},
      {range:'H1:I2', target:'CRM PELANGGAN'},
      {range:'J1:K2', target:'REKAP HARIAN'},
      {range:'L1:M2', target:'REKAP BULANAN'},
      {range:'N1:O2', target:'INVOICE'},
      {range:'P1:Q2', target:'JADWAL HARIAN'},
      {range:'R1:S2', target:'LOGIN'}
    ],
    'DASHBOARD V2': [
      {range:'B1:C2', target:'DASHBOARD V2'},
      {range:'D1:E2', target:'INPUT BOOKING'},
      {range:'F1:G2', target:'TRANSAKSI_LOG'},
      {range:'H1:I2', target:'CRM PELANGGAN'},
      {range:'J1:K2', target:'REKAP HARIAN'},
      {range:'L1:M2', target:'REKAP BULANAN'},
      {range:'N1:O2', target:'INVOICE'},
      {range:'P1:Q2', target:'JADWAL HARIAN'},
      {range:'R1:S2', target:'LOGIN'}
    ],
    'INPUT BOOKING': [
      {range:'B1:C2', target:'DASHBOARD V2'},
      {range:'D1:E2', target:'INPUT BOOKING'},
      {range:'F1:G2', target:'TRANSAKSI_LOG'},
      {range:'H1:I2', target:'CRM PELANGGAN'},
      {range:'J1:K2', target:'REKAP HARIAN'},
      {range:'L1:M2', target:'REKAP BULANAN'},
      {range:'N1:O2', target:'INVOICE'},
      {range:'P1:Q2', target:'JADWAL HARIAN'},
      {range:'R1:S2', target:'LOGIN'}
    ],
    'CRM PELANGGAN': [
      {range:'L1:M2', target:'DASHBOARD V2'},
      {range:'N1:O2', target:'INPUT BOOKING'},
      {range:'P1:Q2', target:'TRANSAKSI_LOG'},
      {range:'R1:S2', target:'CRM PELANGGAN'},
      {range:'T1:U2', target:'REKAP HARIAN'},
      {range:'V1:W2', target:'REKAP BULANAN'},
      {range:'X1:Y2', target:'INVOICE'},
      {range:'Z1:AA2', target:'JADWAL HARIAN'},
      {range:'AB1:AC2', target:'LOGIN'}
    ],
    'TRANSAKSI_LOG': [
      {range:'AJ1:AK2', target:'DASHBOARD V2'},
      {range:'AL1:AM2', target:'INPUT BOOKING'},
      {range:'AN1:AO2', target:'TRANSAKSI_LOG'},
      {range:'AP1:AQ2', target:'CRM PELANGGAN'},
      {range:'AR1:AS2', target:'REKAP HARIAN'},
      {range:'AT1:AU2', target:'REKAP BULANAN'},
      {range:'AV1:AW2', target:'INVOICE'},
      {range:'AX1:AY2', target:'JADWAL HARIAN'},
      {range:'AZ1:BA2', target:'LOGIN'}
    ],
    'REKAP HARIAN': [
      {range:'J1:K2', target:'DASHBOARD V2'},
      {range:'L1:M2', target:'INPUT BOOKING'},
      {range:'N1:O2', target:'TRANSAKSI_LOG'},
      {range:'P1:Q2', target:'CRM PELANGGAN'},
      {range:'R1:S2', target:'REKAP HARIAN'},
      {range:'T1:U2', target:'REKAP BULANAN'},
      {range:'V1:W2', target:'INVOICE'},
      {range:'X1:Y2', target:'JADWAL HARIAN'},
      {range:'Z1:AA2', target:'LOGIN'}
    ],
    'REKAP BULANAN': [
      {range:'Q1:R2', target:'DASHBOARD V2'},
      {range:'S1:T2', target:'INPUT BOOKING'},
      {range:'U1:V2', target:'TRANSAKSI_LOG'},
      {range:'W1:X2', target:'CRM PELANGGAN'},
      {range:'Y1:Z2', target:'REKAP HARIAN'},
      {range:'AA1:AB2', target:'REKAP BULANAN'},
      {range:'AC1:AD2', target:'INVOICE'},
      {range:'AE1:AF2', target:'JADWAL HARIAN'},
      {range:'AG1:AH2', target:'LOGIN'}
    ],
    'INVOICE': [
      {range:'B1:C2', target:'DASHBOARD V2'},
      {range:'D1:E2', target:'INPUT BOOKING'},
      {range:'F1:G2', target:'TRANSAKSI_LOG'},
      {range:'H1:I2', target:'CRM PELANGGAN'},
      {range:'J1:K2', target:'REKAP HARIAN'},
      {range:'L1:M2', target:'REKAP BULANAN'},
      {range:'N1:O2', target:'INVOICE'},
      {range:'P1:Q2', target:'JADWAL HARIAN'},
      {range:'R1:S2', target:'LOGIN'}
    ],
    'JADWAL HARIAN': [
      {range:'B1:C2', target:'DASHBOARD V2'},
      {range:'D1:E2', target:'INPUT BOOKING'},
      {range:'F1:G2', target:'TRANSAKSI_LOG'},
      {range:'H1:I2', target:'CRM PELANGGAN'},
      {range:'J1:K2', target:'REKAP HARIAN'},
      {range:'L1:M2', target:'REKAP BULANAN'},
      {range:'N1:O2', target:'INVOICE'},
      {range:'P1:Q2', target:'JADWAL HARIAN'},
      {range:'R1:S2', target:'LOGIN'}
    ]
  }
};

function ssdNavOnOpen_() {
  return true;
}

function ssdNavHandleSelection_(e) {
  if (!e || !e.range) return false;

  var sh = e.range.getSheet();
  var sheetName = sh.getName();
  var buttons = SSD_NAV.buttons[sheetName];
  if (!buttons || !buttons.length) return false;

  for (var i = 0; i < buttons.length; i++) {
    var btn = buttons[i];
    if (ssdNavRangeHit_(e.range, btn.range)) {
      ssdNavGoToSheet_(btn.target);
      return true;
    }
  }
  return false;
}

function ssdNavGoToSheet_(sheetName) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(sheetName);
  if (!sh) return false;

  if (typeof setSelectionGuard_ === 'function') setSelectionGuard_(true);
  try {
    ss.setActiveSheet(sh);
    sh.setActiveSelection('A1');
    SpreadsheetApp.flush();
  } finally {
    Utilities.sleep(50);
    if (typeof setSelectionGuard_ === 'function') setSelectionGuard_(false);
  }
  return true;
}

function ssdNavRangeHit_(clickedRange, buttonA1) {
  var sh = clickedRange.getSheet();
  var btn = sh.getRange(buttonA1);

  var r1 = clickedRange.getRow();
  var c1 = clickedRange.getColumn();
  var r2 = r1 + clickedRange.getNumRows() - 1;
  var c2 = c1 + clickedRange.getNumColumns() - 1;

  var br1 = btn.getRow();
  var bc1 = btn.getColumn();
  var br2 = br1 + btn.getNumRows() - 1;
  var bc2 = bc1 + btn.getNumColumns() - 1;

  return !(r2 < br1 || r1 > br2 || c2 < bc1 || c1 > bc2);
}
