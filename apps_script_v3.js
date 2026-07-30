function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action || 'hasil';

  if (action === 'daftar') {
    return handleDaftar(ss, e);
  }
  return handleHasil(ss, e);
}

function doGet(e) {
  var callback = e.parameter.callback;
  var result;

  if (e.parameter.action === 'verify') {
    result = verifyLogin(e.parameter.id, e.parameter.kode);
  } else {
    result = { error: "unknown action" };
  }

  var output = callback ? (callback + "(" + JSON.stringify(result) + ")") : JSON.stringify(result);
  var mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(output).setMimeType(mime);
}

// ================== PENDAFTARAN ==================
function handleDaftar(ss, e) {
  var sheet = ss.getSheetByName("Pendaftaran");
  if (!sheet) {
    sheet = ss.insertSheet("Pendaftaran");
  }
  if (sheet.getLastRow() === 0) {
    var headers = ["Waktu Daftar", "Nama", "Kelas", "Kontak (WA/Email)", "Status", "Kode Akses"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  sheet.appendRow([
    new Date(),
    e.parameter.nama,
    e.parameter.kelas,
    e.parameter.kontak,
    "Menunggu",
    ""
  ]);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  sheet.getRange(1, 1, lastRow, lastCol).setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  sheet.getRange(1, 1, lastRow, lastCol).createFilter();

  // Warnai baris yang statusnya masih "Menunggu"
  sheet.clearConditionalFormatRules();
  var statusRange = sheet.getRange(2, 5, Math.max(lastRow - 1, 1), 1);
  var ruleMenunggu = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Menunggu")
    .setBackground("#F7EFD9").setFontColor("#6B5416")
    .setRanges([statusRange]).build();
  var ruleDisetujui = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Disetujui")
    .setBackground("#DCE8E4").setFontColor("#153029")
    .setRanges([statusRange]).build();
  sheet.setConditionalFormatRules([ruleMenunggu, ruleDisetujui]);

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" })).setMimeType(ContentService.MimeType.JSON);
}

function verifyLogin(idPengguna, kode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pendaftaran");
  if (!sheet || !idPengguna || !kode) return { valid: false };

  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowNama = String(row[1]).trim().toLowerCase();
    var rowKode = String(row[5]).trim();
    var status = String(row[4]).trim();
    if (
      rowNama === String(idPengguna).trim().toLowerCase() &&
      rowKode !== "" &&
      rowKode.toLowerCase() === String(kode).trim().toLowerCase() &&
      status === "Disetujui"
    ) {
      return { valid: true, nama: row[1], kelas: row[2] };
    }
  }
  return { valid: false };
}

// Menu bantu di spreadsheet: pilih baris siswa di sheet "Pendaftaran", lalu klik menu ini
// untuk otomatis menyetujui + membuat kode akses acak.
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Ruang Ujian')
    .addItem('Setujui baris terpilih & buat kode akses', 'setujuiBarisTerpilih')
    .addToUi();
}

function setujuiBarisTerpilih() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pendaftaran");
  var ui = SpreadsheetApp.getUi();
  if (!sheet || ss.getActiveSheet().getName() !== "Pendaftaran") {
    ui.alert('Buka sheet "Pendaftaran" dan pilih baris siswa terlebih dahulu.');
    return;
  }
  var row = sheet.getActiveCell().getRow();
  if (row === 1) {
    ui.alert('Pilih baris data siswa, bukan baris header.');
    return;
  }
  var existingKode = sheet.getRange(row, 6).getValue();
  var kode = existingKode ? existingKode : Utilities.getUuid().substring(0, 6).toUpperCase();
  sheet.getRange(row, 5).setValue("Disetujui");
  sheet.getRange(row, 6).setValue(kode);
  ui.alert('Baris ' + row + ' disetujui.\nID Pengguna (Nama): ' + sheet.getRange(row, 2).getValue() + '\nKata Sandi: ' + kode + '\n\nSampaikan ID Pengguna dan Kata Sandi ini ke siswa yang bersangkutan.');
}

// ================== HASIL UJIAN ==================
function handleHasil(ss, e) {
  var sheet = ss.getSheetByName("Hasil");
  if (!sheet) {
    sheet = ss.insertSheet("Hasil");
  }
  if (sheet.getLastRow() === 0) {
    var headers = ["Waktu", "Nama", "Kelas", "Paket", "Mapel", "Skor", "Benar", "Salah", "Kosong", "Total Soal"];
    sheet.appendRow(headers);
  }

  sheet.appendRow([
    new Date(),
    e.parameter.nama,
    e.parameter.kelas,
    e.parameter.paket,
    e.parameter.mapel,
    Number(e.parameter.skor),
    Number(e.parameter.benar),
    Number(e.parameter.salah),
    Number(e.parameter.kosong),
    Number(e.parameter.total)
  ]);

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  var headerRange = sheet.getRange(1, 1, 1, lastCol);
  headerRange.setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);

  fullRange.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");
  }

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  fullRange.createFilter();

  sheet.clearConditionalFormatRules();
  var skorRange = sheet.getRange(2, 6, Math.max(lastRow - 1, 1), 1);
  var ruleLulus = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(75)
    .setBackground("#DCE8E4").setFontColor("#153029")
    .setRanges([skorRange]).build();
  var ruleBelum = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(75)
    .setBackground("#F3DED6").setFontColor("#7A2E1A")
    .setRanges([skorRange]).build();
  sheet.setConditionalFormatRules([ruleLulus, ruleBelum]);

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" })).setMimeType(ContentService.MimeType.JSON);
}
