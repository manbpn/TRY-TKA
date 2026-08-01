// ============ KONFIGURASI ============
// Ganti dengan link GitHub Pages ujian Anda, dipakai untuk pesan WA.
var EXAM_URL = "https://manbpn.github.io/TRY-TKA/";

function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var action = e.parameter.action || 'hasil';
  if (action === 'daftar') return handleDaftar(ss, e);
  return handleHasil(ss, e);
}

function doGet(e) {
  var callback = e.parameter.callback;
  var action = e.parameter.action;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var result;

  if (action === 'verify') {
    result = verifyLogin(e.parameter.id, e.parameter.kode);
  } else if (action === 'admin_data') {
    result = checkAdminPassword(e.parameter.password)
      ? { ok: true, pendaftaran: adminGetPendaftaran(ss), hasil: adminGetHasil(ss) }
      : { ok: false, error: 'unauthorized' };
  } else if (action === 'admin_approve') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Pendaftaran");
      ensureWaHeader(sheet);
      var hasil = setujuiSatuBaris(sheet, Number(e.parameter.row));
      formatPendaftaranSheet(sheet);
      return { ok: true, nama: hasil.nama, kode: hasil.kode };
    });
  } else if (action === 'admin_approve_all') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Pendaftaran");
      if (!sheet) return { ok: true, count: 0 };
      ensureWaHeader(sheet);
      var lastRow = sheet.getLastRow();
      var count = 0;
      for (var row = 2; row <= lastRow; row++) {
        var status = String(sheet.getRange(row, 5).getValue()).trim();
        if (status === "Menunggu") { setujuiSatuBaris(sheet, row); count++; }
      }
      formatPendaftaranSheet(sheet);
      return { ok: true, count: count };
    });
  } else if (action === 'admin_reject') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Pendaftaran");
      sheet.getRange(Number(e.parameter.row), 5).setValue("Ditolak");
      formatPendaftaranSheet(sheet);
      return { ok: true };
    });
  } else if (action === 'admin_reset_kode') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Pendaftaran");
      var row = Number(e.parameter.row);
      sheet.getRange(row, 6).setValue(""); // kosongkan supaya kode baru dibuat
      var hasil = setujuiSatuBaris(sheet, row);
      formatPendaftaranSheet(sheet);
      return { ok: true, nama: hasil.nama, kode: hasil.kode };
    });
  } else if (action === 'admin_delete_pendaftaran') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Pendaftaran");
      sheet.deleteRow(Number(e.parameter.row));
      formatPendaftaranSheet(sheet);
      return { ok: true };
    });
  } else if (action === 'admin_delete_hasil') {
    result = adminAction(e.parameter.password, function () {
      var sheet = ss.getSheetByName("Hasil");
      sheet.deleteRow(Number(e.parameter.row));
      return { ok: true };
    });
  } else {
    result = { error: "unknown action" };
  }

  var output = callback ? (callback + "(" + JSON.stringify(result) + ")") : JSON.stringify(result);
  var mime = callback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(output).setMimeType(mime);
}

// ================== ADMIN ==================
// Password default (bisa diganti lewat menu "Ruang Ujian > Atur Password Admin").
var ADMIN_PASSWORD_DEFAULT = "ubah-password-ini";

function checkAdminPassword(pw) {
  var stored = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD') || ADMIN_PASSWORD_DEFAULT;
  return !!pw && pw === stored;
}

function adminAction(password, fn) {
  if (!checkAdminPassword(password)) return { ok: false, error: 'unauthorized' };
  try {
    return fn();
  } catch (err) {
    return { ok: false, error: String(err) };
  }
}

function adminGetPendaftaran(ss) {
  var sheet = ss.getSheetByName("Pendaftaran");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue;
    out.push({
      row: i + 1,
      waktu: fmtTanggal(row[0]),
      nama: row[1], kelas: row[2], kontak: row[3], status: row[4], kode: row[5]
    });
  }
  return out;
}

function adminGetHasil(ss) {
  var sheet = ss.getSheetByName("Hasil");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row[1]) continue;
    out.push({
      row: i + 1,
      waktu: fmtTanggal(row[0]),
      nama: row[1], kelas: row[2], paket: row[3], mapel: row[4],
      skor: row[5], benar: row[6], salah: row[7], kosong: row[8], total: row[9]
    });
  }
  return out;
}

function fmtTanggal(val) {
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
  }
  return String(val || "");
}

// Menu bantu: atur/ganti password admin tanpa perlu edit kode.
function aturPasswordAdmin() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt(
    'Atur Password Admin',
    'Masukkan password baru untuk menu Admin di aplikasi ujian (dipakai guru/panitia untuk mengelola pendaftaran & hasil):',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp.getSelectedButton() == ui.Button.OK) {
    var pw = resp.getResponseText().trim();
    if (pw) {
      PropertiesService.getScriptProperties().setProperty('ADMIN_PASSWORD', pw);
      ui.alert('Password admin berhasil disimpan.\n\nGunakan password ini saat membuka menu "Masuk sebagai Admin" di aplikasi ujian.');
    } else {
      ui.alert('Password tidak boleh kosong.');
    }
  }
}

// ================== PENDAFTARAN ==================
function handleDaftar(ss, e) {
  var sheet = ss.getSheetByName("Pendaftaran");
  if (!sheet) sheet = ss.insertSheet("Pendaftaran");

  if (sheet.getLastRow() === 0) {
    var headers = ["Waktu Daftar", "Nama", "Kelas", "Kontak (WA/Email)", "Status", "Kode Akses", "Link Kirim WA"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
  }
  ensureWaHeader(sheet);

  sheet.appendRow([new Date(), e.parameter.nama, e.parameter.kelas, e.parameter.kontak, "Menunggu", "", ""]);

  formatPendaftaranSheet(sheet);

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" })).setMimeType(ContentService.MimeType.JSON);
}

function ensureWaHeader(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 7 || sheet.getRange(1, 7).getValue() !== "Link Kirim WA") {
    sheet.getRange(1, 7).setValue("Link Kirim WA")
      .setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
  }
}

function formatPendaftaranSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 7);
  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  fullRange.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  fullRange.createFilter();

  sheet.clearConditionalFormatRules();
  var statusRange = sheet.getRange(2, 5, Math.max(lastRow - 1, 1), 1);
  var ruleMenunggu = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Menunggu").setBackground("#F7EFD9").setFontColor("#6B5416")
    .setRanges([statusRange]).build();
  var ruleDisetujui = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Disetujui").setBackground("#DCE8E4").setFontColor("#153029")
    .setRanges([statusRange]).build();
  sheet.setConditionalFormatRules([ruleMenunggu, ruleDisetujui]);
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
    if (rowNama === String(idPengguna).trim().toLowerCase() && rowKode !== "" &&
        rowKode.toLowerCase() === String(kode).trim().toLowerCase() && status === "Disetujui") {
      return { valid: true, nama: row[1], kelas: row[2] };
    }
  }
  return { valid: false };
}

// ---------- Nomor WA ----------
function formatNomorWa(raw) {
  var digits = String(raw || "").replace(/[^0-9]/g, "");
  if (digits.indexOf("0") === 0) digits = "62" + digits.substring(1);
  else if (digits.indexOf("62") !== 0) digits = "62" + digits;
  return digits;
}

function buatLinkWa(nama, kelas, kode) {
  var pesan = "Halo " + nama + ", pendaftaran tryout TKA kamu sudah disetujui.\n" +
              "ID Pengguna: " + nama + "\n" +
              "Kata Sandi: " + kode + "\n" +
              "Silakan masuk lewat: " + EXAM_URL;
  return encodeURIComponent(pesan);
}

// ---------- Menu & persetujuan ----------
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Ruang Ujian')
    .addItem('Setujui baris terpilih & buat kode akses', 'setujuiBarisTerpilih')
    .addItem('Setujui SEMUA pendaftar yang menunggu', 'setujuiSemuaMenunggu')
    .addSeparator()
    .addItem('Atur Password Admin (untuk menu Admin di web ujian)', 'aturPasswordAdmin')
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
  if (row === 1) { ui.alert('Pilih baris data siswa, bukan baris header.'); return; }

  ensureWaHeader(sheet);
  var hasil = setujuiSatuBaris(sheet, row);
  formatPendaftaranSheet(sheet);

  ui.alert('Baris ' + row + ' disetujui.\n\nID Pengguna: ' + hasil.nama + '\nKata Sandi: ' + hasil.kode +
           '\n\nLink kirim WA sudah otomatis dibuat di kolom G baris ini — tinggal klik linknya untuk membuka WhatsApp dengan pesan siap kirim.');
}

function setujuiSemuaMenunggu() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Pendaftaran");
  var ui = SpreadsheetApp.getUi();
  if (!sheet) { ui.alert('Sheet "Pendaftaran" belum ada.'); return; }

  ensureWaHeader(sheet);
  var lastRow = sheet.getLastRow();
  var count = 0;
  for (var row = 2; row <= lastRow; row++) {
    var status = String(sheet.getRange(row, 5).getValue()).trim();
    if (status === "Menunggu") {
      setujuiSatuBaris(sheet, row);
      count++;
    }
  }
  formatPendaftaranSheet(sheet);
  ui.alert(count + ' pendaftar berhasil disetujui.\n\nBuka kolom "Link Kirim WA" di setiap baris untuk mengirim ID Pengguna & Kata Sandi ke masing-masing siswa lewat WhatsApp (klik link-nya satu per satu, WhatsApp akan membuka pesan yang sudah siap kirim).');
}

function setujuiSatuBaris(sheet, row) {
  var nama = sheet.getRange(row, 2).getValue();
  var kelas = sheet.getRange(row, 3).getValue();
  var kontak = sheet.getRange(row, 4).getValue();
  var existingKode = sheet.getRange(row, 6).getValue();
  var kode = existingKode ? existingKode : Utilities.getUuid().substring(0, 6).toUpperCase();

  sheet.getRange(row, 5).setValue("Disetujui");
  sheet.getRange(row, 6).setValue(kode);

  var nomorWa = formatNomorWa(kontak);
  var pesanEncoded = buatLinkWa(nama, kelas, kode);
  var link = "https://wa.me/" + nomorWa + "?text=" + pesanEncoded;
  var cell = sheet.getRange(row, 7);
  cell.setFormula('=HYPERLINK("' + link + '","Kirim WA ke ' + nama + '")');

  return { nama: nama, kode: kode };
}

// ================== HASIL UJIAN ==================
function cariKontakByNama(ss, nama) {
  var sheet = ss.getSheetByName("Pendaftaran");
  if (!sheet) return "";
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim().toLowerCase() === String(nama).trim().toLowerCase()) {
      return data[i][3]; // kolom D = Kontak
    }
  }
  return "";
}

function buatLinkWaHasil(nama, paket, mapel, skor, benar, salah, kosong, total, kontak) {
  var nomorWa = formatNomorWa(kontak);
  if (!nomorWa) return "";
  var pesan = "Halo " + nama + ", berikut hasil tryout kamu:\n" +
              "Paket: " + paket + " (" + mapel + ")\n" +
              "Skor: " + skor + "\n" +
              "Benar: " + benar + " | Salah: " + salah + " | Kosong: " + kosong + " dari " + total + " soal\n" +
              "Terus semangat belajar!";
  var link = "https://wa.me/" + nomorWa + "?text=" + encodeURIComponent(pesan);
  return '=HYPERLINK("' + link + '","Kirim hasil ke ' + nama + '")';
}

function handleHasil(ss, e) {
  var sheet = ss.getSheetByName("Hasil");
  if (!sheet) sheet = ss.insertSheet("Hasil");
  if (sheet.getLastRow() === 0) {
    var headers = ["Waktu", "Nama", "Kelas", "Paket", "Mapel", "Skor", "Benar", "Salah", "Kosong", "Total Soal", "Link Kirim WA"];
    sheet.appendRow(headers);
  }
  if (sheet.getRange(1, 11).getValue() !== "Link Kirim WA") {
    sheet.getRange(1, 11).setValue("Link Kirim WA").setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
  }

  var nama = e.parameter.nama, paket = e.parameter.paket, mapel = e.parameter.mapel;
  var skor = Number(e.parameter.skor), benar = Number(e.parameter.benar);
  var salah = Number(e.parameter.salah), kosong = Number(e.parameter.kosong), total = Number(e.parameter.total);

  sheet.appendRow([new Date(), nama, e.parameter.kelas, paket, mapel, skor, benar, salah, kosong, total, ""]);

  var lastRow = sheet.getLastRow();
  var kontak = cariKontakByNama(ss, nama);
  var linkFormula = buatLinkWaHasil(nama, paket, mapel, skor, benar, salah, kosong, total, kontak);
  if (linkFormula) sheet.getRange(lastRow, 11).setFormula(linkFormula);

  var lastCol = sheet.getLastColumn();
  var fullRange = sheet.getRange(1, 1, lastRow, lastCol);

  sheet.getRange(1, 1, 1, lastCol).setFontWeight("bold").setBackground("#2B5D50").setFontColor("#FFFFFF");
  sheet.setFrozenRows(1);
  fullRange.setBorder(true, true, true, true, true, true, "#CCCCCC", SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, lastCol);

  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, 1).setNumberFormat("dd/mm/yyyy hh:mm");

  var existingFilter = sheet.getFilter();
  if (existingFilter) existingFilter.remove();
  fullRange.createFilter();

  sheet.clearConditionalFormatRules();
  var skorRange = sheet.getRange(2, 6, Math.max(lastRow - 1, 1), 1);
  var ruleLulus = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThanOrEqualTo(75).setBackground("#DCE8E4").setFontColor("#153029")
    .setRanges([skorRange]).build();
  var ruleBelum = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThan(75).setBackground("#F3DED6").setFontColor("#7A2E1A")
    .setRanges([skorRange]).build();
  sheet.setConditionalFormatRules([ruleLulus, ruleBelum]);

  return ContentService.createTextOutput(JSON.stringify({ status: "ok" })).setMimeType(ContentService.MimeType.JSON);
}
