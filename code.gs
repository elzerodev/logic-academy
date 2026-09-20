/**
 * Excel Logic Academy — Google Apps Script Web App
 * ---------------------------------------------------
 * Struktur:
 *   index.html -> landing page (perkenalan modul)
 *   soal.html  -> modul interaktif (peta level, lab spreadsheet, kuis, sertifikat)
 *
 * doGet menentukan file mana yang ditampilkan lewat parameter URL ?page=
 *   (kosong)   -> index.html
 *   ?page=soal -> soal.html
 */

/* Link donasi/tip kopi (Saweria, Trakteer, dll) yang ditampilkan opsional
   di layar sertifikat soal.html. Kosongkan ('') untuk menyembunyikan
   tombolnya. Disimpan di sini (server) supaya gampang diganti tanpa
   menyentuh file HTML. */
var COFFEE_TIP_URL = 'https://tiptap.gg/dikilabs';

function doGet(e) {
  var page = (e && e.parameter && e.parameter.page) ? String(e.parameter.page) : 'index';
  var fileName = (page === 'soal') ? 'soal' : 'index';

  var template = HtmlService.createTemplateFromFile(fileName);
  // scriptUrl dipakai di dalam HTML (lewat <?= scriptUrl ?>) untuk membuat
  // tautan antar halaman tetap valid setelah aplikasi di-deploy.
  template.scriptUrl = ScriptApp.getService().getUrl();
  // coffeeTipUrl dipakai di soal.html (lewat <?= coffeeTipUrl ?>) untuk
  // menampilkan tombol tip kopi opsional di layar sertifikat.
  template.coffeeTipUrl = COFFEE_TIP_URL;

  return template.evaluate()
    .setTitle('Excel Logic Academy — Basic to Master')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper opsional — sediakan kalau nanti CSS/JS ingin dipecah jadi file
 * terpisah (mis. "styles.html") dan disisipkan lewat <?!= include('styles'); ?>
 * Tidak dipakai secara default karena index.html & soal.html sudah mandiri.
 */
function include(fileName) {
  return HtmlService.createHtmlOutputFromFile(fileName).getContent();
}

/* ============================= PENYIMPANAN HASIL (Google Sheets) =============================
 * Dipanggil dari soal.html (lewat google.script.run) setiap kali seseorang
 * membuat sertifikat di layar terakhir. Menulis satu baris berisi:
 *   datetime | nama | nilaifinal | pesan | tips
 * ke Google Sheet dengan ID di bawah. Sheet (tab) dibuat otomatis kalau
 * belum ada, lengkap dengan header di baris pertama.
 * ============================================================================================= */

var RESULT_SPREADSHEET_ID = '18dsEgSIO9_tRbbvADdo8zkFh-kewAvWx2g4NYRGr_Dk';
var RESULT_SHEET_NAME = 'Hasil Excel Logic Academy';
var RESULT_HEADERS = ['datetime', 'nama', 'nilaifinal', 'pesan', 'tips'];

/**
 * Mengambil sheet tujuan di dalam spreadsheet RESULT_SPREADSHEET_ID.
 * Kalau sheet-nya belum ada, buat baru sekaligus tulis header di baris 1.
 * Kalau sheet sudah ada tapi baris headernya kosong/tidak lengkap, header
 * akan ditulis ulang supaya urutan kolom tetap konsisten.
 */
function getOrCreateResultSheet_() {
  var ss = SpreadsheetApp.openById(RESULT_SPREADSHEET_ID);
  var sheet = ss.getSheetByName(RESULT_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(RESULT_SHEET_NAME);
  }

  var firstRow = sheet.getRange(1, 1, 1, RESULT_HEADERS.length).getValues()[0];
  var headerOk = RESULT_HEADERS.every(function (h, i) { return firstRow[i] === h; });
  if (!headerOk) {
    sheet.getRange(1, 1, 1, RESULT_HEADERS.length).setValues([RESULT_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, RESULT_HEADERS.length).setFontWeight('bold');
  }

  return sheet;
}

/**
 * Fungsi yang dipanggil dari client (soal.html) lewat:
 *   google.script.run.saveCertificateResult({ nama, nilaifinal, pesan, tips })
 *
 * data: {
 *   nama:       string  -> nama yang diketik di layar sertifikat
 *   nilaifinal: number  -> skor akhir gabungan seluruh level (0-100)
 *   pesan:      string  -> isi jurnal refleksi dari Level terakhir
 *   tips:       string  -> tips otomatis berdasarkan skor akhir
 * }
 * Mengembalikan { ok:true } jika berhasil, atau { ok:false, error:'...' } jika gagal
 * (client tidak boleh menganggap kegagalan ini sebagai error fatal — lihat soal.html).
 */
function saveCertificateResult(data) {
  try {
    data = data || {};
    var sheet = getOrCreateResultSheet_();

    var nama = data.nama ? String(data.nama).trim() : '';
    var nilaifinal = (typeof data.nilaifinal === 'number' && !isNaN(data.nilaifinal)) ? data.nilaifinal : '';
    var pesan = data.pesan ? String(data.pesan).trim() : '';
    var tips = data.tips ? String(data.tips).trim() : '';

    sheet.appendRow([new Date(), nama, nilaifinal, pesan, tips]);

    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}