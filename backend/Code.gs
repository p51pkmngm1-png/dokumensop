/**
 * GUDANG KENDALI DOKUMEN — INPUT SOP
 * Backend: Google Apps Script Web App + Google Sheet sebagai database
 *
 * ============================================================
 * CARA SETUP
 * ============================================================
 * 1. Sheet database: "Database Gudang Kendali Dokumen" (di folder Drive kamu).
 *    SHEET_ID di bawah udah otomatis diisi sesuai sheet itu.
 * 2. Buka sheet itu → Extensions > Apps Script → hapus isi lama → tempel SELURUH isi file ini.
 * 3. Jalankan function `setup` sekali (pilih "setup" di dropdown run, klik Run).
 *    Ini bikin tab Dokumen, Counter, & Referensi otomatis (Referensi diisi
 *    data awal contoh — SILAKAN EDIT LANGSUNG DI SHEET sesuai klaster/program
 *    Puskesmas kamu yang sebenarnya, gak perlu ubah kode).
 * 4. Deploy > Manage deployments > Edit (pensil) > New version > Deploy.
 *    (Karena ini UPDATE dari deployment lama, bukan New deployment — biar
 *    URL /exec yang udah dipasang di frontend tetap sama, gak perlu ganti lagi.)
 * ============================================================
 */

const SHEET_ID = "13PgodNVw9P_f1WDO9XxlGtAsOCNTneRX9G1Fn7LdoSU";
const SHEET_DOKUMEN = "Dokumen";
const SHEET_COUNTER = "Counter";
const SHEET_REFERENSI = "Referensi";

const DOKUMEN_HEADERS = [
  "NoDokumen", "Judul", "Klaster", "SubKlaster", "Status", "TglTerbit", "NoRevisi",
  "NoDokumenLama", "AlasanGanti", "NamaKepalaPuskesmas", "NIPKepalaPuskesmas",
  "ItemsJSON", "DibuatPada", "DiubahPada",
];

const COUNTER_HEADERS = ["Klaster", "SubKlaster", "Tahun", "UrutanTerakhir"];
const REFERENSI_HEADERS = ["KodeKlaster", "NamaKlaster", "SubKlaster", "KodeSubKlaster"];

// Data awal tab Referensi — cuma dipakai SEKALI pas pertama kali setup().
// Setelahnya, edit langsung di Sheet, gak usah balik ke sini.
// Kolom ke-4 (KodeSubKlaster) itu yang jadi angka ".01"/".02" dst di nomor
// dokumen — isi manual per baris, unik per klaster (boleh sama antar klaster
// beda, misal K1.01 dan K2.01 sama-sama boleh "01").
const REFERENSI_SEED = [
  ["K1", "Manajemen", "Manajemen Inti Puskesmas", "01"],
  ["K1", "Manajemen", "Manajemen Arsip & Dokumen", "02"],
  ["K1", "Manajemen", "Manajemen Sumber Daya Manusia", "03"],
  ["K1", "Manajemen", "Manajemen Sarana, Prasarana & Perbekalan Kesehatan", "04"],
  ["K1", "Manajemen", "Manajemen Mutu Pelayanan", "05"],
  ["K1", "Manajemen", "Manajemen Keuangan & Aset/BMD", "06"],
  ["K1", "Manajemen", "Manajemen Sistem Informasi & Digital", "07"],
  ["K1", "Manajemen", "Manajemen Jejaring", "08"],
  ["K2", "Ibu dan Anak", "Ibu Hamil", "01"],
  ["K2", "Ibu dan Anak", "Ibu Bersalin & Nifas", "02"],
  ["K2", "Ibu dan Anak", "Bayi Baru Lahir (BBL)", "03"],
  ["K2", "Ibu dan Anak", "Balita", "04"],
  ["K2", "Ibu dan Anak", "Anak Usia Sekolah & Remaja", "05"],
  ["K2", "Ibu dan Anak", "Gizi", "06"],
  ["K3", "Usia Produktif dan Lansia", "Usia Produktif", "01"],
  ["K3", "Usia Produktif dan Lansia", "Lansia", "02"],
  ["K3", "Usia Produktif dan Lansia", "Penyakit Tidak Menular (PTM)", "03"],
  ["K3", "Usia Produktif dan Lansia", "Kesehatan Jiwa", "04"],
  ["K3", "Usia Produktif dan Lansia", "Kesehatan Reproduksi (Kespro)", "05"],
  ["K3", "Usia Produktif dan Lansia", "Keluarga Berencana (KB)", "06"],
  ["K3", "Usia Produktif dan Lansia", "Upaya Berhenti Merokok (UBM)", "07"],
  ["K4", "Penanggulangan Penyakit", "Pencegahan & Pengendalian Penyakit (P2)", "01"],
  ["K4", "Penanggulangan Penyakit", "Surveilans", "02"],
  ["K4", "Penanggulangan Penyakit", "Kesehatan Lingkungan", "03"],
  ["K4", "Penanggulangan Penyakit", "Imunisasi", "04"],
  ["K5", "Lintas Klaster", "Farmasi", "01"],
  ["K5", "Lintas Klaster", "Laboratorium", "02"],
  ["K5", "Lintas Klaster", "Kesehatan Gigi & Mulut", "03"],
  ["K5", "Lintas Klaster", "Gizi", "04"],
  ["K5", "Lintas Klaster", "Psikologi", "05"],
  ["K5", "Lintas Klaster", "Fisioterapi", "06"],
  ["K5", "Lintas Klaster", "IGD", "07"],
  ["K5", "Lintas Klaster", "Rawat Inap", "08"],
  ["K5", "Lintas Klaster", "Kesehatan Kerja & Olahraga", "09"],
];

/* ------------------------------------------------------------
   SETUP — jalankan sekali manual dari editor Apps Script
   ------------------------------------------------------------ */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  ensureSheet_(ss, SHEET_DOKUMEN, DOKUMEN_HEADERS);
  ensureSheet_(ss, SHEET_COUNTER, COUNTER_HEADERS);
  const refSheet = ensureSheet_(ss, SHEET_REFERENSI, REFERENSI_HEADERS);
  if (refSheet.getLastRow() === 1) {
    refSheet.getRange(2, 1, REFERENSI_SEED.length, REFERENSI_HEADERS.length).setValues(REFERENSI_SEED);
  }
  Logger.log("Setup selesai. Tab Dokumen, Counter, & Referensi siap dipakai.");
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  // Selalu samain baris header (baris 1) ke versi kode terbaru — aman dijalanin
  // berkali-kali, cuma nulis ulang label kolom, data di baris bawahnya gak kesentuh.
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (sheet.getFrozenRows() < 1) sheet.setFrozenRows(1);
  return sheet;
}

function getSheet_(name, headers) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  return ensureSheet_(ss, name, headers);
}

/* ------------------------------------------------------------
   ENTRY POINTS
   ------------------------------------------------------------
   Frontend kirim POST dengan Content-Type: text/plain (bukan
   application/json) supaya browser gak kirim CORS preflight
   (Apps Script Web App gak bisa jawab OPTIONS). Body-nya tetap
   string JSON biasa, di sini di-parse manual.
*/
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action;
    let result;

    switch (action) {
      case "generateNomor":
        result = generateNomor_(body.klaster, body.subKlaster, body.tahun);
        break;
      case "saveDokumen":
        result = saveDokumen_(body.dokumen);
        break;
      case "finalisasiDokumen":
        result = finalisasiDokumen_(body.noDokumenLama, body.klaster, body.subKlaster, body.tahun, body.noDokumenDigantikan);
        break;
      case "getDokumenByNomor":
        result = getDokumenByNomor_(body.noDokumen);
        break;
      case "getReferensi":
        result = getReferensi_();
        break;
      default:
        return jsonResponse_({ ok: false, error: "Action tidak dikenali: " + action });
    }

    return jsonResponse_({ ok: true, data: result });
  } catch (err) {
    return jsonResponse_({ ok: false, error: String(err) });
  }
}

// GET dipakai buat hal-hal read-only ringan (query lewat URL, gak kena CORS preflight).
// Dukung JSONP juga: kalau ada parameter "callback", respons dibungkus jadi
// pemanggilan fungsi JS (bukan JSON polos) — ini AKALIN batasan Apps Script
// yang gak bisa nambahin header CORS ke responsnya, jadi fetch() browser
// selalu gagal baca hasilnya biar pun request-nya sukses. Lewat tag <script>
// (yang dipakai JSONP), CORS itu gak berlaku sama sekali.
function doGet(e) {
  const callback = e.parameter.callback;
  try {
    const action = e.parameter.action;
    let result;

    switch (action) {
      case "ping":
        result = { message: "Backend Gudang Kendali Dokumen aktif." };
        break;
      case "getDokumenByNomor":
        result = getDokumenByNomor_(e.parameter.noDokumen);
        break;
      case "getReferensi":
        result = getReferensi_();
        break;
      case "listDokumen":
        result = listDokumen_();
        break;
      case "finalisasiDokumen":
        result = finalisasiDokumen_(
          e.parameter.noDokumenLama,
          e.parameter.klaster,
          e.parameter.subKlaster,
          e.parameter.tahun,
          e.parameter.noDokumenDigantikan || undefined
        );
        break;
      default:
        return jsonpResponse_({ ok: false, error: "Action tidak dikenali: " + action }, callback);
    }

    return jsonpResponse_({ ok: true, data: result }, callback);
  } catch (err) {
    return jsonpResponse_({ ok: false, error: String(err) }, callback);
  }
}

function jsonResponse_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// Kalau ada "callback" (dipanggil dari <script> tag ala JSONP), bungkus jadi
// pemanggilan fungsi JS — ini yang bikin browser gak nge-block hasilnya walau
// Apps Script gak bisa ngasih header CORS. Kalau gak ada callback (dipanggil
// langsung/dari Postman/dsb), tetap balikin JSON polos seperti biasa.
function jsonpResponse_(obj, callback) {
  if (callback) {
    const safe = String(callback).replace(/[^a-zA-Z0-9_$.]/g, "");
    return ContentService.createTextOutput(safe + "(" + JSON.stringify(obj) + ");").setMimeType(
      ContentService.MimeType.JAVASCRIPT
    );
  }
  return jsonResponse_(obj);
}

/* ------------------------------------------------------------
   REFERENSI — daftar Klaster & Sub-Klaster (Program), diambil
   langsung dari tab "Referensi" di Sheet. Edit datanya langsung
   di Sheet, gak perlu ubah kode ini.
   Dikembalikan berbentuk:
   {
     klasterList: [{ kode, nama }, ...],           // unik per klaster
     subKlasterByKlaster: { K1: ["Umum", ...], ... }
   }
   ------------------------------------------------------------ */
/* Cari index kolom berdasarkan NAMA header (bukan posisi tetap) — biar
   aman walau kolom di tab Referensi digeser/diinsert manual di Sheet. */
function findCol_(headerRow, name) {
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i] || "").trim() === name) return i;
  }
  return -1;
}

function getReferensi_() {
  const sheet = getSheet_(SHEET_REFERENSI, REFERENSI_HEADERS);
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const cKode = findCol_(header, "KodeKlaster");
  const cNama = findCol_(header, "NamaKlaster");
  const cSub = findCol_(header, "SubKlaster");
  const cKodeSub = findCol_(header, "KodeSubKlaster");

  const klasterMap = {}; // kode -> nama
  const subKlasterByKlaster = {};
  const kodeSubKlasterByKlaster = {}; // { K1: { "Umum": "01", ... } }

  for (let i = 1; i < data.length; i++) {
    const kode = String(data[i][cKode] || "").trim();
    const nama = cNama > -1 ? String(data[i][cNama] || "").trim() : "";
    const sub = cSub > -1 ? String(data[i][cSub] || "").trim() : "";
    const kodeSub = cKodeSub > -1 ? String(data[i][cKodeSub] || "").trim() : "";
    if (!kode) continue;

    if (!klasterMap[kode]) klasterMap[kode] = nama;
    if (!subKlasterByKlaster[kode]) subKlasterByKlaster[kode] = [];
    if (!kodeSubKlasterByKlaster[kode]) kodeSubKlasterByKlaster[kode] = {};
    if (sub && subKlasterByKlaster[kode].indexOf(sub) === -1) {
      subKlasterByKlaster[kode].push(sub);
      kodeSubKlasterByKlaster[kode][sub] = kodeSub;
    }
  }

  const klasterList = Object.keys(klasterMap).map((kode) => ({ kode: kode, nama: klasterMap[kode] }));

  return {
    klasterList: klasterList,
    subKlasterByKlaster: subKlasterByKlaster,
    kodeSubKlasterByKlaster: kodeSubKlasterByKlaster,
  };
}

/* ------------------------------------------------------------
   Cari KodeSubKlaster buat pasangan klaster + sub-klaster tertentu —
   dipakai buat bagian ".01" / ".02" dst di nomor dokumen. Baca kolom
   berdasarkan nama header, bukan posisi, biar tahan urutan kolom diubah.
   ------------------------------------------------------------ */
function getKodeSubKlaster_(klaster, subKlaster) {
  const sheet = getSheet_(SHEET_REFERENSI, REFERENSI_HEADERS);
  const data = sheet.getDataRange().getValues();
  const header = data[0] || [];
  const cKode = findCol_(header, "KodeKlaster");
  const cSub = findCol_(header, "SubKlaster");
  const cKodeSub = findCol_(header, "KodeSubKlaster");

  for (let i = 1; i < data.length; i++) {
    const kode = String(data[i][cKode] || "").trim();
    const sub = cSub > -1 ? String(data[i][cSub] || "").trim() : "";
    if (kode === klaster && sub === subKlaster) {
      const kodeSub = cKodeSub > -1 ? String(data[i][cKodeSub] || "").trim() : "";
      return kodeSub || "00";
    }
  }
  // Sub-klaster belum terdaftar di tab Referensi — tetap kasih nomor,
  // pakai "00" biar gak nge-block generate, tapi ini kasus perlu dicek manual.
  return "00";
}

/* ------------------------------------------------------------
   GENERATE NOMOR DOKUMEN
   Format: SOP.{2 digit tahun}/{Klaster}.{2 digit urutan sub-klaster}/{3 digit urutan dokumen}
   Contoh: SOP.26/K1.01/003
   Urutan dokumen RESET tiap tahun, per kombinasi klaster+sub-klaster
   (disimpan di tab Counter). Pakai LockService biar aman kalau ada
   2 orang generate bareng.
   ------------------------------------------------------------ */
function generateNomor_(klaster, subKlaster, tahun) {
  if (!klaster || !subKlaster || !tahun) {
    throw new Error("klaster, subKlaster, dan tahun wajib diisi.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(SHEET_COUNTER, COUNTER_HEADERS);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    let urutanTerakhir = 0;
    for (let i = 1; i < data.length; i++) {
      if (
        String(data[i][0]) === String(klaster) &&
        String(data[i][1]) === String(subKlaster) &&
        String(data[i][2]) === String(tahun)
      ) {
        rowIndex = i + 1; // 1-based, +1 karena header
        urutanTerakhir = Number(data[i][3]) || 0;
        break;
      }
    }

    const urutanBaru = urutanTerakhir + 1;

    if (rowIndex === -1) {
      sheet.appendRow([klaster, subKlaster, tahun, urutanBaru]);
    } else {
      sheet.getRange(rowIndex, 4).setValue(urutanBaru);
    }

    const kodeSub = getKodeSubKlaster_(klaster, subKlaster);
    const tahun2 = String(tahun).slice(-2);
    const nomor = `SOP.${tahun2}/${klaster}.${kodeSub.padStart(2, "0")}/${String(urutanBaru).padStart(3, "0")}`;

    return {
      nomor: nomor,
      urutan: urutanBaru,
      klaster: klaster,
      subKlaster: subKlaster,
      kodeSubKlaster: kodeSub,
      tahun: tahun,
    };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
   SIMPAN DOKUMEN (create kalau nomor belum ada, update kalau sudah)
   dokumen (dari frontend) diharapkan berbentuk:
   {
     noDokumen, judul, klaster, subKlaster, status, tglTerbit, noRevisi,
     noDokumenLama, alasanGanti, namaKepalaPuskesmas, nipKepalaPuskesmas,
     items: { pengertian, tujuan, ..., catatanRevisi: [...] }
   }
   ------------------------------------------------------------ */
/* ------------------------------------------------------------
   FORMAT JUDUL DOKUMEN — setiap judul yang disimpan otomatis
   diformat jadi "SOP <Proper Case>". Contoh:
   "pelayanan imunisasi dasar lengkap" -> "SOP Pelayanan Imunisasi Dasar Lengkap"
   Kalau user udah nulis "SOP ..." di depan sendiri, gak didobel —
   prefix-nya cuma dirapiin jadi "SOP" huruf besar semua.
   Cuma huruf pertama tiap kata yang dikapital — sisanya dibiarin
   apa adanya, biar singkatan kayak "(PTM)", "BMD", "(P2)" gak
   ke-rusak jadi huruf kecil.
   ------------------------------------------------------------ */
function toProperTitle_(str) {
  return String(str || "").replace(/\S+/g, function (word) {
    return word.charAt(0).toUpperCase() + word.slice(1);
  });
}

function formatJudulDokumen_(judul) {
  const trimmed = String(judul || "").trim();
  if (!trimmed) return trimmed;
  const proper = toProperTitle_(trimmed);
  if (/^SOP\s+/i.test(proper)) {
    return "SOP " + proper.replace(/^SOP\s+/i, "");
  }
  return "SOP " + proper;
}

function saveDokumen_(dokumen) {
  if (!dokumen || !dokumen.noDokumen) {
    throw new Error("Data dokumen atau noDokumen kosong.");
  }

  dokumen.judul = formatJudulDokumen_(dokumen.judul);

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
    const data = sheet.getDataRange().getValues();
    const now = new Date().toISOString();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(dokumen.noDokumen)) {
        rowIndex = i + 1;
        break;
      }
    }

    const itemsJSON = JSON.stringify(dokumen.items || {});

    const row = [
      dokumen.noDokumen,
      dokumen.judul || "",
      dokumen.klaster || "",
      dokumen.subKlaster || "",
      dokumen.status || "Draft",
      dokumen.tglTerbit || "",
      dokumen.noRevisi != null ? dokumen.noRevisi : 0,
      dokumen.noDokumenLama || "",
      dokumen.alasanGanti || "",
      dokumen.namaKepalaPuskesmas || "",
      dokumen.nipKepalaPuskesmas || "",
      itemsJSON,
      rowIndex === -1 ? now : data[rowIndex - 1][12], // DibuatPada tetap, gak berubah kalau update
      now, // DiubahPada
    ];

    if (rowIndex === -1) {
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    }

    return { noDokumen: dokumen.noDokumen, tersimpan: true };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
   FINALISASI DOKUMEN — dipanggil pas dokumen lolos SEMUA tilik.
   Dokumen "Baru" gak lagi punya nomor resmi dari awal (cuma ID
   sementara kayak "Draft-1758..."), nomor resmi baru diterbitkan
   di sini: generate nomor beneran, lalu GANTI key NoDokumen baris
   itu di Sheet dari ID sementara ke nomor resmi + set status
   Terkendali + NoRevisi 0. Isi dokumen (ItemsJSON dkk) gak disentuh,
   cuma key & status/revisi-nya yang berubah.
   ------------------------------------------------------------ */
function finalisasiDokumen_(noDokumenLama, klaster, subKlaster, tahun, noDokumenDigantikan) {
  if (!noDokumenLama || !klaster || !subKlaster || !tahun) {
    throw new Error("noDokumenLama, klaster, subKlaster, dan tahun wajib diisi.");
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
    const data = sheet.getDataRange().getValues();

    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(noDokumenLama)) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) {
      throw new Error("Dokumen dengan ID sementara '" + noDokumenLama + "' tidak ditemukan.");
    }

    const nomorInfo = generateNomor_(klaster, subKlaster, tahun);
    const now = new Date().toISOString();

    sheet.getRange(rowIndex, 1).setValue(nomorInfo.nomor); // NoDokumen
    sheet.getRange(rowIndex, 4).setValue(subKlaster); // SubKlaster (jaga-jaga kalau beda)
    sheet.getRange(rowIndex, 3).setValue(klaster); // Klaster
    sheet.getRange(rowIndex, 5).setValue("Terkendali"); // Status
    sheet.getRange(rowIndex, 7).setValue(0); // NoRevisi
    sheet.getRange(rowIndex, 14).setValue(now); // DiubahPada

    // Dokumen "Baru" ini ceklis "Ganti" & isi No SOP Dokumen Lama — begitu
    // dokumen ini resmi terbit (Terkendali), dokumen lama yang digantikan
    // otomatis di-set Kadaluarsa juga, dalam transaksi yang sama.
    let dokumenDigantikanExpired = false;
    if (noDokumenDigantikan) {
      let oldRowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(noDokumenDigantikan)) {
          oldRowIndex = i + 1;
          break;
        }
      }
      if (oldRowIndex !== -1) {
        sheet.getRange(oldRowIndex, 5).setValue("Kadaluarsa"); // Status
        sheet.getRange(oldRowIndex, 14).setValue(now); // DiubahPada
        dokumenDigantikanExpired = true;
      }
      // Kalau dokumen lamanya gak ketemu di Sheet, gak nge-block finalisasi
      // dokumen baru — cukup dicatat aja gak ke-expire otomatis.
    }

    return {
      noDokumenLama: noDokumenLama,
      noDokumenBaru: nomorInfo.nomor,
      status: "Terkendali",
      dokumenDigantikanExpired: dokumenDigantikanExpired,
    };
  } finally {
    lock.releaseLock();
  }
}

/* ------------------------------------------------------------
   AMBIL DOKUMEN BY NOMOR — dipakai buat lookup Revisi/Ganti
   ------------------------------------------------------------ */
function getDokumenByNomor_(noDokumen) {
  if (!noDokumen) {
    throw new Error("noDokumen wajib diisi.");
  }

  const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(noDokumen)) {
      let items = {};
      try {
        items = JSON.parse(data[i][11] || "{}");
      } catch (e) {
        items = {};
      }
      return {
        noDokumen: data[i][0],
        judul: data[i][1],
        klaster: data[i][2],
        subKlaster: data[i][3],
        status: data[i][4],
        tglTerbit: data[i][5]
          ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd")
          : "",
        noRevisi: data[i][6],
        noDokumenLama: data[i][7],
        alasanGanti: data[i][8],
        namaKepalaPuskesmas: data[i][9],
        nipKepalaPuskesmas: data[i][10],
        items: items,
      };
    }
  }

  return null; // gak ketemu — frontend nampilin "notfound"
}

/* ------------------------------------------------------------
   LIST SEMUA DOKUMEN — buat isi dashboard. Sengaja gak ikut parse
   ItemsJSON (dashboard cuma butuh ringkasan), biar cepet & ringan
   walau datanya udah banyak baris.
   ------------------------------------------------------------ */
/* ------------------------------------------------------------
   PERBAIKAN DATA — jalankan SEKALI manual dari editor Apps Script
   kalau ada NoDokumen yang formatnya rusak (misal urutan cuma 2
   digit karena keketik manual, bukan hasil generateNomor_ asli).
   Aman dijalankan berkali-kali — cuma benerin yang formatnya gak
   sesuai pola SOP.YY/K{n}.{SS}/{NNN} (urutan wajib 3 digit).
   ------------------------------------------------------------ */
function perbaikiFormatNomorDokumen() {
  const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
  const data = sheet.getDataRange().getValues();
  const pola = /^(SOP\.\d{2}\/K\d+\.\d{2})\/(\d+)$/;
  let jumlahDiperbaiki = 0;

  for (let i = 1; i < data.length; i++) {
    const noDokumen = String(data[i][0] || "").trim();
    const match = noDokumen.match(pola);
    if (!match) continue; // ID sementara (Draft-...) atau format lain, skip

    const urutanAsli = match[2];
    if (urutanAsli.length === 3) continue; // udah bener, skip

    const nomorBenar = `${match[1]}/${urutanAsli.padStart(3, "0")}`;
    sheet.getRange(i + 1, 1).setValue(nomorBenar);
    Logger.log(`Diperbaiki: "${noDokumen}" -> "${nomorBenar}"`);
    jumlahDiperbaiki++;
  }

  Logger.log(`Selesai. ${jumlahDiperbaiki} nomor dokumen diperbaiki.`);
}

function listDokumen_() {
  const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
  const data = sheet.getDataRange().getValues();
  const list = [];

  for (let i = 1; i < data.length; i++) {
    const noDokumen = String(data[i][0] || "").trim();
    if (!noDokumen) continue; // baris kosong, skip
    list.push({
      noDokumen: noDokumen,
      judul: data[i][1],
      klaster: data[i][2],
      subKlaster: data[i][3],
      status: data[i][4],
      tglTerbit: data[i][5]
        ? Utilities.formatDate(new Date(data[i][5]), Session.getScriptTimeZone(), "yyyy-MM-dd")
        : "",
      noRevisi: data[i][6],
      noDokumenLama: data[i][7],
    });
  }

  return list;
}

/* ------------------------------------------------------------
   SEED DATA DUMMY — jalankan SEKALI manual dari editor Apps Script
   buat masukin 11 dokumen contoh (yang tadinya cuma ada di frontend)
   ke tab Dokumen. Aman dijalankan berkali-kali — noDokumen yang
   sudah ada bakal di-UPDATE (bukan dobel), bukan di-skip.
   Catatan: gambar bagan alir (base64) sengaja TIDAK disertakan biar
   sel Sheet nggak kegedean — cuma nama filenya yang disimpan.
   ------------------------------------------------------------ */
const DUMMY_SEED_DOCS = [
  {
    "noDokumen": "SOP.24/K3.02/001",
    "judul": "Pelayanan Lansia di Posbindu",
    "klaster": "K3",
    "subKlaster": "Lansia",
    "status": "Draft",
    "tglTerbit": "2024-07-19",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Pelayanan lansia di Posbindu adalah kegiatan deteksi dini dan pemantauan kesehatan lanjut usia di masyarakat.",
      "tujuan": "Meningkatkan kualitas hidup lansia melalui deteksi dini penyakit tidak menular.",
      "kebijakan": "SK Kepala Puskesmas tentang Pelayanan Kesehatan Usia Lanjut.",
      "referensi": "Pedoman Pelayanan Kesehatan Lanjut Usia, Kemenkes RI.",
      "prosedur": "1. Kader mendaftar peserta.\n2. Petugas mengukur tekanan darah, gula darah, dan status gizi.\n3. Petugas memberikan edukasi dan rujukan bila diperlukan.",
      "baganAlir": {
        "name": "bagan-alir-posbindu-lansia.png",
        "pages": 1
      },
      "halPerlu": "Perhatikan lansia dengan riwayat penyakit kronis.",
      "unitTerkait": "Poli Lansia, Kader Posbindu, Poli Umum.",
      "dokumenTerkait": "Kartu Menuju Sehat Lansia, Buku Register Posbindu.",
      "catatanRevisi": []
    }
  },
  {
    "noDokumen": "SOP.26/K4.03/002",
    "judul": "Penanganan Limbah Medis B3",
    "klaster": "K4",
    "subKlaster": "Kesehatan Lingkungan",
    "status": "Draft",
    "tglTerbit": "2026-04-02",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {}
  },
  {
    "noDokumen": "SOP.24/K5.03/001",
    "judul": "Pelayanan Kesehatan Gigi & Mulut Dasar",
    "klaster": "K5",
    "subKlaster": "Kesehatan Gigi & Mulut",
    "status": "Revisi",
    "tglTerbit": "2024-09-02",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Pelayanan kesehatan gigi dan mulut dasar adalah upaya promotif, preventif, dan kuratif sederhana terhadap masalah kesehatan gigi dan mulut pasien.",
      "tujuan": "Menurunkan angka karies dan penyakit gigi mulut di wilayah kerja Puskesmas.",
      "kebijakan": "SK Kepala Puskesmas tentang Pelayanan Kesehatan Gigi dan Mulut.",
      "referensi": "Pedoman Pelayanan Kesehatan Gigi dan Mulut di Puskesmas, Kemenkes RI.",
      "prosedur": "1. Petugas melakukan anamnesa keluhan gigi pasien.\n2. Dokter gigi melakukan pemeriksaan intraoral.\n3. Dokter gigi memberikan tindakan sesuai indikasi (tambal, cabut, scaling) atau rujukan.\n4. Petugas mencatat tindakan di rekam medis gigi.",
      "baganAlir": {
        "name": "bagan-alir-kesgimul-dasar.png",
        "pages": 1
      },
      "halPerlu": "Pastikan sterilisasi alat sebelum dan sesudah tindakan.",
      "unitTerkait": "Poli Gigi, Ruang Sterilisasi, Ruang Farmasi.",
      "dokumenTerkait": "Rekam Medis Gigi, Odontogram, Informed Consent Tindakan.",
      "catatanRevisi": []
    }
  },
  {
    "noDokumen": "SOP.25/K2.02/001",
    "judul": "Pelayanan Ibu Bersalin & Nifas Normal",
    "klaster": "K2",
    "subKlaster": "Ibu Bersalin & Nifas",
    "status": "Revisi",
    "tglTerbit": "2025-04-14",
    "noRevisi": 1,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Pelayanan ibu bersalin dan nifas normal adalah asuhan yang diberikan pada ibu selama proses persalinan hingga masa nifas tanpa penyulit.",
      "tujuan": "Memastikan proses persalinan dan nifas berjalan aman bagi ibu dan bayi.",
      "kebijakan": "SK Kepala Puskesmas tentang Pelayanan Persalinan Normal di Puskesmas PONED.",
      "referensi": "Buku Saku Pelayanan Kesehatan Ibu di Fasilitas Kesehatan Dasar dan Rujukan, Kemenkes RI.",
      "prosedur": "1. Bidan melakukan penilaian kala I persalinan.\n2. Bidan mendampingi proses persalinan kala II-IV sesuai APN.\n3. Bidan melakukan observasi nifas 2 jam pertama.\n4. Bidan memberikan edukasi ASI eksklusif dan perawatan nifas.",
      "baganAlir": {
        "name": "bagan-alir-persalinan-nifas.png",
        "pages": 1
      },
      "halPerlu": "Segera rujuk apabila ditemukan tanda bahaya persalinan/nifas.",
      "unitTerkait": "Ruang Bersalin, Poli KIA, Ambulans Rujukan.",
      "dokumenTerkait": "Partograf, Buku KIA, Kohort Ibu Bersalin.",
      "catatanRevisi": [
        {
          "no": 1,
          "yangDirubah": "Prosedur",
          "isiPerubahan": "Penambahan langkah observasi nifas 2 jam pertama.",
          "tgl": "14 April 2025"
        }
      ]
    }
  },
  {
    "noDokumen": "SOP.26/K4.03/001",
    "judul": "Inspeksi Kesehatan Lingkungan Puskesmas",
    "klaster": "K4",
    "subKlaster": "Kesehatan Lingkungan",
    "status": "Menunggu Verifikasi",
    "tglTerbit": "2026-01-05",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Inspeksi kesehatan lingkungan adalah kegiatan pemeriksaan sanitasi dan lingkungan fisik puskesmas secara berkala.",
      "tujuan": "Memastikan lingkungan puskesmas memenuhi standar kesehatan lingkungan.",
      "kebijakan": "SK Kepala Puskesmas tentang Kesehatan Lingkungan Puskesmas.",
      "referensi": "Permenkes tentang Persyaratan Kesehatan Lingkungan Puskesmas.",
      "prosedur": "1. Petugas kesling menyusun jadwal inspeksi.\n2. Petugas melakukan pemeriksaan menggunakan checklist.\n3. Petugas menyusun laporan & rekomendasi tindak lanjut.",
      "baganAlir": {
        "name": "bagan-alir-inspeksi-kesling.png",
        "pages": 1
      },
      "halPerlu": "Dokumen masih menunggu verifikasi, belum disahkan.",
      "unitTerkait": "Tim Kesling, Tim Mutu, Manajemen Puskesmas.",
      "dokumenTerkait": "Checklist Inspeksi Kesling, Formulir Tindak Lanjut.",
      "catatanRevisi": []
    }
  },
  {
    "noDokumen": "SOP.26/K2.04/001",
    "judul": "Pemantauan Tumbuh Kembang Balita",
    "klaster": "K2",
    "subKlaster": "Balita",
    "status": "Menunggu Verifikasi",
    "tglTerbit": "2026-02-10",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Pemantauan tumbuh kembang balita adalah kegiatan pengukuran dan penilaian pertumbuhan serta perkembangan anak usia 0-59 bulan secara berkala.",
      "tujuan": "Mendeteksi dini gangguan pertumbuhan (stunting, gizi kurang) dan perkembangan balita.",
      "kebijakan": "SK Kepala Puskesmas tentang Pelayanan Kesehatan Anak Balita.",
      "referensi": "Buku Kesehatan Ibu dan Anak (KIA), Kemenkes RI.",
      "prosedur": "1. Kader/petugas mengukur berat badan, tinggi/panjang badan, dan lingkar kepala.\n2. Petugas memplot hasil pada grafik pertumbuhan.\n3. Petugas melakukan skrining perkembangan menggunakan KPSP.\n4. Petugas memberi rujukan bila ditemukan penyimpangan.",
      "baganAlir": {
        "name": "bagan-alir-tumbuh-kembang-balita.png",
        "pages": 1
      },
      "halPerlu": "Dokumen masih menunggu verifikasi, belum disahkan.",
      "unitTerkait": "Posyandu, Poli KIA, Poli Gizi.",
      "dokumenTerkait": "Buku KIA, KMS, Formulir KPSP.",
      "catatanRevisi": []
    }
  },
  {
    "noDokumen": "SOP.23/K2.01/001",
    "judul": "Pemeriksaan Antenatal Care (ANC) Terpadu",
    "klaster": "K2",
    "subKlaster": "Ibu Hamil",
    "status": "Terkendali",
    "tglTerbit": "2023-03-10",
    "noRevisi": 1,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "ANC terpadu adalah pelayanan pemeriksaan kehamilan yang komprehensif dan berkualitas, diberikan pada seluruh ibu hamil.",
      "tujuan": "Memantau kemajuan kehamilan, mengenali secara dini kelainan/penyakit yang diderita ibu hamil, dan mempersiapkan persalinan yang aman.",
      "kebijakan": "SK Kepala Puskesmas Ngemplak I tentang Pelayanan Klinis Ibu dan Anak.",
      "referensi": "Pedoman Pelayanan Antenatal Terpadu, Kemenkes RI, 2020.",
      "prosedur": "1. Petugas memanggil pasien sesuai nomor urut.\n2. Petugas melakukan anamnesa.\n3. Petugas melakukan pemeriksaan fisik dan penunjang.\n4. Petugas mencatat hasil pada buku KIA.",
      "baganAlir": {
        "name": "bagan-alir-anc-terpadu.png",
        "pages": 1
      },
      "halPerlu": "Perhatikan tanda bahaya kehamilan.",
      "unitTerkait": "Poli KIA, Laboratorium, Ruang Farmasi.",
      "dokumenTerkait": "Buku KIA, Rekam Medis, Kohort Ibu.",
      "catatanRevisi": [
        {
          "no": 1,
          "yangDirubah": "Prosedur",
          "isiPerubahan": "Penyesuaian alur pemeriksaan laboratorium.",
          "tgl": "10 Maret 2023"
        }
      ]
    }
  },
  {
    "noDokumen": "SOP.25/K3.02/001",
    "judul": "Skrining Kesehatan Lansia di Posyandu",
    "klaster": "K3",
    "subKlaster": "Lansia",
    "status": "Terkendali",
    "tglTerbit": "2025-06-11",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {}
  },
  {
    "noDokumen": "SOP.22/K5.04/001",
    "judul": "Pelayanan Konsultasi Gizi Lintas Program",
    "klaster": "K5",
    "subKlaster": "Gizi",
    "status": "Kadaluarsa",
    "tglTerbit": "2022-04-08",
    "noRevisi": 2,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Konsultasi gizi lintas program adalah layanan konseling gizi yang melibatkan lebih dari satu program (KIA, TB, Lansia, dll).",
      "tujuan": "Memberikan edukasi gizi yang terintegrasi sesuai kondisi klinis pasien.",
      "kebijakan": "SK Kepala Puskesmas tentang Pelayanan Gizi Terintegrasi.",
      "referensi": "Pedoman Gizi Seimbang, Kemenkes RI, 2019.",
      "prosedur": "1. Petugas program merujuk pasien ke konsultasi gizi.\n2. Ahli gizi melakukan asesmen status gizi.\n3. Ahli gizi memberikan rencana diet & edukasi.",
      "baganAlir": {
        "name": "bagan-alir-konsultasi-gizi.png",
        "pages": 1
      },
      "halPerlu": "Dokumen ini sudah digantikan versi terbaru, tidak berlaku lagi.",
      "unitTerkait": "Poli Gizi, Poli KIA, Poli TB, Poli Lansia.",
      "dokumenTerkait": "Formulir Asesmen Gizi, Kartu Status Gizi.",
      "catatanRevisi": [
        {
          "no": 1,
          "yangDirubah": "Prosedur",
          "isiPerubahan": "Penambahan alur rujukan internal antar poli.",
          "tgl": "12 Juni 2022"
        },
        {
          "no": 2,
          "yangDirubah": "Kebijakan, Referensi",
          "isiPerubahan": "Update rujukan ke pedoman gizi seimbang terbaru.",
          "tgl": "20 Oktober 2022"
        }
      ]
    }
  },
  {
    "noDokumen": "SOP.21/K1.01/001",
    "judul": "Orientasi Pegawai Baru",
    "klaster": "K1",
    "subKlaster": "Manajemen Inti Puskesmas",
    "status": "Kadaluarsa",
    "tglTerbit": "2021-05-20",
    "noRevisi": 0,
    "noDokumenLama": "",
    "alasanGanti": "",
    "namaKepalaPuskesmas": "dr. Vivin Ratri Kurniasari Agustiyani Sulistyo",
    "nipKepalaPuskesmas": "19830827 201001 2 010",
    "items": {
      "pengertian": "Orientasi pegawai baru adalah kegiatan pengenalan lingkungan kerja, tugas pokok, dan budaya kerja bagi pegawai yang baru bergabung di Puskesmas.",
      "tujuan": "Memastikan pegawai baru memahami tugas, tanggung jawab, dan tata nilai organisasi sebelum bertugas mandiri.",
      "kebijakan": "SK Kepala Puskesmas tentang Orientasi Pegawai Baru.",
      "referensi": "Pedoman Manajemen Sumber Daya Manusia Puskesmas, Kemenkes RI.",
      "prosedur": "1. Bagian kepegawaian menyusun jadwal orientasi.\n2. Pegawai baru mengikuti pengenalan struktur organisasi & unit kerja.\n3. Pegawai baru didampingi mentor selama masa orientasi.\n4. Bagian kepegawaian melakukan evaluasi akhir masa orientasi.",
      "baganAlir": {
        "name": "bagan-alir-orientasi-pegawai.png",
        "pages": 1
      },
      "halPerlu": "Dokumen ini sudah digantikan versi terbaru, tidak berlaku lagi.",
      "unitTerkait": "Bagian Kepegawaian, Seluruh Unit Kerja.",
      "dokumenTerkait": "Formulir Evaluasi Orientasi, Daftar Hadir Orientasi.",
      "catatanRevisi": []
    }
  }
];

const DUMMY_SEED_COUNTER = [
  [
    "K3",
    "Lansia",
    "2024",
    1
  ],
  [
    "K4",
    "Kesehatan Lingkungan",
    "2026",
    2
  ],
  [
    "K5",
    "Kesehatan Gigi & Mulut",
    "2024",
    1
  ],
  [
    "K2",
    "Ibu Bersalin & Nifas",
    "2025",
    1
  ],
  [
    "K2",
    "Balita",
    "2026",
    1
  ],
  [
    "K2",
    "Ibu Hamil",
    "2023",
    1
  ],
  [
    "K3",
    "Lansia",
    "2025",
    1
  ],
  [
    "K5",
    "Gizi",
    "2022",
    1
  ],
  [
    "K1",
    "Manajemen Inti Puskesmas",
    "2021",
    1
  ]
];

function seedDummyData() {
  let count = 0;
  DUMMY_SEED_DOCS.forEach(function (doc) {
    saveDokumen_(doc);
    count++;
  });

  // Samain tab Counter biar generate nomor beneran nanti nggak nabrak
  // urutan yang udah dipakai dokumen dummy ini.
  const counterSheet = getSheet_(SHEET_COUNTER, COUNTER_HEADERS);
  const counterData = counterSheet.getDataRange().getValues();
  DUMMY_SEED_COUNTER.forEach(function (row) {
    const klaster = row[0], subKlaster = row[1], tahun = row[2], urutan = row[3];
    let rowIndex = -1;
    for (let i = 1; i < counterData.length; i++) {
      if (
        String(counterData[i][0]) === String(klaster) &&
        String(counterData[i][1]) === String(subKlaster) &&
        String(counterData[i][2]) === String(tahun)
      ) {
        rowIndex = i + 1;
        break;
      }
    }
    if (rowIndex === -1) {
      counterSheet.appendRow([klaster, subKlaster, tahun, urutan]);
    } else {
      const existing = Number(counterData[rowIndex - 1][3]) || 0;
      counterSheet.getRange(rowIndex, 4).setValue(Math.max(existing, urutan));
    }
  });

  Logger.log(count + " dokumen dummy berhasil dimasukkan/diupdate ke tab Dokumen, Counter disinkronkan.");
}


/* ------------------------------------------------------------
   BERSIHIN DUMMY LAMA — jalankan SEKALI manual kalau tab Dokumen
   masih ada sisa baris dummy versi lama yang udah gak dipakai lagi
   (nomor dokumennya gak ada di daftar 10 dummy terbaru). Dokumen
   ASLI yang udah kamu buat sendiri TIDAK kesentuh — cuma yang nomor
   dokumennya persis cocok sama daftar dummy LAMA yang dihapus.
   ------------------------------------------------------------ */
const NOMOR_DUMMY_TERBARU = [
  "SOP.24/K3.02/001",
  "SOP.26/K4.03/002",
  "SOP.24/K5.03/001",
  "SOP.25/K2.02/001",
  "SOP.26/K4.03/001",
  "SOP.26/K2.04/001",
  "SOP.23/K2.01/001",
  "SOP.25/K3.02/001",
  "SOP.22/K5.04/001",
  "SOP.21/K1.01/001"
];

function hapusDummyLama() {
  const sheet = getSheet_(SHEET_DOKUMEN, DOKUMEN_HEADERS);
  const data = sheet.getDataRange().getValues();
  const nomorLamaYangDihapus = [
    "SOP/014/K2/2023", "SOP/006/K1/2024", "SOP/021/K2/2024", "SOP/009/K3/2024",
    "SOP/003/K4/2025", "SOP/012/K5/2022", "SOP/017/K3/2025", "SOP/002/K1/2026",
    "SOP/025/K4/2026", "SOP/030/K1/2026", "SOP/018/K3/2025",
    "SOP.24/K1.02/001", "SOP.24/K4.04/001", "SOP.25/K3.06/001",
    "SOP.26/K3.03/001", "SOP.25/K4.01/001", "SOP.24/K5.02/001"
  ];
  let hapusCount = 0;
  // Hapus dari baris paling bawah biar index gak geser pas dihapus.
  for (let i = data.length - 1; i >= 1; i--) {
    const noDok = String(data[i][0]);
    if (nomorLamaYangDihapus.indexOf(noDok) !== -1) {
      sheet.deleteRow(i + 1);
      hapusCount++;
    }
  }
  Logger.log(hapusCount + " baris dummy lama berhasil dihapus dari tab Dokumen.");
}
