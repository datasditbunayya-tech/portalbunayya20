// ============================================================
//  BUNAYYA ISLAMIC SCHOOL — Google Apps Script (kode.gs)
//  Spreadsheet ID: 1NyNLjFhDOqsi4KYGM-lmB3ZGupt7amn2W67kYUIac-Y
//
//  CARA PASANG:
//  1. Buka Google Sheet → Extensions → Apps Script
//  2. Paste seluruh kode ini, Save (Ctrl+S)
//  3. Klik "Deploy" → "New Deployment"
//  4. Type: Web App | Execute as: Me | Who has access: Anyone
//  5. Deploy → Salin Web App URL → Tempel ke SCRIPT_URL di portal
//  6. Jalankan fungsi setupSheets() SATU KALI untuk buat semua sheet
// ============================================================

const SPREADSHEET_ID = "1NyNLjFhDOqsi4KYGM-lmB3ZGupt7amn2W67kYUIac-Y";

// ── Sheet names & header definitions ────────────────────────────────────────
const SHEETS = {
  absensiSiswa: {
    name: "Absensi Siswa",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Tanggal", "Status", "Waktu Input"]
  },
  absensiGuru: {
    name: "Absensi Guru",
    headers: ["ID", "Guru ID", "Nama Guru", "Kelas", "Tanggal", "Status", "Waktu Input"]
  },
  jurnal: {
    name: "Jurnal Mengajar",
    headers: ["ID", "Guru ID", "Nama Guru", "Kelas", "Mata Pelajaran", "Tanggal", "Topik", "Kegiatan", "Catatan", "Waktu Input"]
  },
  harian: {
    name: "Laporan Harian",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Tanggal", "Bahasan", "Halaman", "Hasil", "Catatan", "Waktu Input"]
  },
  nilai: {
    name: "Rekap Nilai",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Mata Pelajaran", "Jenis Nilai", "Skor", "Tanggal", "Waktu Input"]
  },
  kegiatan: {
    name: "Laporan Kegiatan",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Mata Pelajaran", "Penilaian", "Tanggal", "Waktu Input"]
  },
  target: {
    name: "Target Pencapaian",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Mata Pelajaran", "Progress Saat Ini", "Progress Total", "Satuan", "Hasil", "Catatan", "Tanggal", "Waktu Input"]
  },
  ziyadah: {
    name: "Ziyadah Hafalan",
    headers: ["ID", "Siswa ID", "Nama Siswa", "Kelas", "Juz", "Surat", "Progress", "Hasil", "Catatan", "Tanggal", "Waktu Input"]
  },
  guru: {
    name: "Master Guru",
    headers: ["ID", "NIP", "Nama", "Email", "Kelas", "Status", "Waktu Input"]
  },
  siswa: {
    name: "Master Siswa",
    headers: ["ID", "NIS", "Nama", "Kelas", "Waktu Input"]
  },
  log: {
    name: "Log Aktivitas",
    headers: ["Waktu", "Aksi", "Sheet", "Data", "User"]
  }
};

// ── Helper: ambil atau buat sheet ────────────────────────────────────────────
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    styleHeader(sheet, headers.length);
  }
  return sheet;
}

// ── Helper: style header row ─────────────────────────────────────────────────
function styleHeader(sheet, colCount) {
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  headerRange.setBackground("#1d7a5f");
  headerRange.setFontColor("#ffffff");
  headerRange.setFontWeight("bold");
  headerRange.setHorizontalAlignment("center");
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 60);   // ID
  for (let i = 2; i <= colCount; i++) sheet.setColumnWidth(i, 140);
}

// ── Helper: tulis log ────────────────────────────────────────────────────────
function writeLog(ss, aksi, sheetName, data, user) {
  try {
    const logSheet = getOrCreateSheet(ss, SHEETS.log.name, SHEETS.log.headers);
    logSheet.appendRow([
      new Date().toLocaleString("id-ID"),
      aksi, sheetName,
      JSON.stringify(data).substring(0, 300),
      user || "portal"
    ]);
  } catch (e) { /* log error tidak menghentikan operasi */ }
}

// ── Helper: cari baris by ID ─────────────────────────────────────────────────
function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) return i + 1; // 1-indexed
  }
  return -1;
}

// ── SETUP: jalankan SATU KALI untuk inisialisasi semua sheet ─────────────────
function setupSheets() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Object.values(SHEETS).forEach(s => {
    getOrCreateSheet(ss, s.name, s.headers);
  });

  // Hapus Sheet1 default jika masih ada
  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  Logger.log("✅ Setup selesai! Semua sheet berhasil dibuat.");
}

// ── CORS Headers ─────────────────────────────────────────────────────────────
function setCorsHeaders(output) {
  return output
    .setMimeType(ContentService.MimeType.JSON)
    .addHeader("Access-Control-Allow-Origin", "*")
    .addHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .addHeader("Access-Control-Allow-Headers", "Content-Type");
}

// ── GET: baca semua data dari satu sheet ─────────────────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetKey = e.parameter.sheet;
    const action   = e.parameter.action || "read";

    if (action === "readAll") {
      // Baca semua sheet sekaligus
      const result = {};
      Object.entries(SHEETS).forEach(([key, cfg]) => {
        const sheet = ss.getSheetByName(cfg.name);
        if (!sheet) { result[key] = []; return; }
        const rows = sheet.getDataRange().getValues();
        if (rows.length <= 1) { result[key] = []; return; }
        const headers = rows[0];
        result[key] = rows.slice(1).map(row => {
          const obj = {};
          headers.forEach((h, i) => { obj[h] = row[i]; });
          return obj;
        });
      });
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: "ok", data: result })));
    }

    if (!sheetKey || !SHEETS[sheetKey]) {
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "Sheet tidak ditemukan: " + sheetKey })
      ));
    }

    const cfg   = SHEETS[sheetKey];
    const sheet = ss.getSheetByName(cfg.name);
    if (!sheet || sheet.getLastRow() <= 1) {
      return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: "ok", data: [] })));
    }

    const rows    = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data    = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

    return setCorsHeaders(ContentService.createTextOutput(JSON.stringify({ status: "ok", data })));

  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ));
  }
}

// ── POST: tulis / hapus data ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
    // Support both application/json and text/plain (untuk CORS compatibility)
    let rawBody = "";
    try { rawBody = e.postData.contents; } catch(ex) { rawBody = "{}"; }
    if (!rawBody || rawBody.trim() === "") rawBody = "{}";
    const body = JSON.parse(rawBody);
    const { action, sheet: sheetKey, data, id } = body;

    // clearSheet tidak perlu validasi sheetKey di sini (sudah di atas)
    if (!sheetKey || !SHEETS[sheetKey]) {
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "Sheet key tidak valid: " + sheetKey })
      ));
    }

    const cfg   = SHEETS[sheetKey];
    const sheet = getOrCreateSheet(ss, cfg.name, cfg.headers);
    const ts    = new Date().toLocaleString("id-ID");

    // ── INSERT ─────────────────────────────────────────────────────────────
    if (action === "insert") {
      if (!data || !Array.isArray(data)) {
        return setCorsHeaders(ContentService.createTextOutput(
          JSON.stringify({ status: "error", message: "Data harus berupa array" })
        ));
      }
      const inserted = [];
      data.forEach(row => {
        const rowData = buildRow(sheetKey, row, ts);
        sheet.appendRow(rowData);
        inserted.push(row.id || row.ID);
      });
      writeLog(ss, "INSERT", cfg.name, { count: inserted.length }, body.user);
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "ok", inserted: inserted.length, message: `${inserted.length} baris berhasil disimpan ke ${cfg.name}` })
      ));
    }

    // ── DELETE ─────────────────────────────────────────────────────────────
    if (action === "delete") {
      if (!id) {
        return setCorsHeaders(ContentService.createTextOutput(
          JSON.stringify({ status: "error", message: "ID diperlukan untuk delete" })
        ));
      }
      const rowNum = findRowById(sheet, id);
      if (rowNum === -1) {
        return setCorsHeaders(ContentService.createTextOutput(
          JSON.stringify({ status: "error", message: "Data dengan ID " + id + " tidak ditemukan" })
        ));
      }
      sheet.deleteRow(rowNum);
      writeLog(ss, "DELETE", cfg.name, { id }, body.user);
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "ok", message: "Data berhasil dihapus", id })
      ));
    }

    // ── CLEAR SHEET (hapus semua baris kecuali header) ──────────────────────
    if (action === "clearSheet") {
      if (sheetKey && SHEETS[sheetKey]) {
        const s = getOrCreateSheet(ss, SHEETS[sheetKey].name, SHEETS[sheetKey].headers);
        if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
        writeLog(ss, "CLEAR", SHEETS[sheetKey].name, {}, body.user);
        return setCorsHeaders(ContentService.createTextOutput(
          JSON.stringify({ status: "ok", message: "Sheet dikosongkan: " + SHEETS[sheetKey].name })
        ));
      }
    }

    // ── SYNC SEMUA (bulk replace seluruh sheet) ───────────────────────────
    if (action === "syncAll") {
      const allData = body.allData; // { absensiSiswa: [...], guru: [...], ... }
      if (!allData) return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "error", message: "allData diperlukan untuk syncAll" })
      ));

      const results = {};
      Object.entries(allData).forEach(([key, rows]) => {
        if (!SHEETS[key] || !Array.isArray(rows)) return;
        const s   = getOrCreateSheet(ss, SHEETS[key].name, SHEETS[key].headers);
        const ts2 = new Date().toLocaleString("id-ID");
        // Hapus isi lama (kecuali header)
        if (s.getLastRow() > 1) s.deleteRows(2, s.getLastRow() - 1);
        // Tulis ulang
        rows.forEach(row => { s.appendRow(buildRow(key, row, ts2)); });
        results[key] = rows.length;
      });

      writeLog(ss, "SYNC_ALL", "ALL", results, body.user);
      return setCorsHeaders(ContentService.createTextOutput(
        JSON.stringify({ status: "ok", synced: results })
      ));
    }

    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: "Action tidak dikenal: " + action })
    ));

  } catch (err) {
    return setCorsHeaders(ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ));
  }
}

// ── BUILD ROW: mapping data object → array sesuai urutan kolom ───────────────
function buildRow(sheetKey, d, ts) {
  switch (sheetKey) {
    case "absensiSiswa":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.tanggal, d.status, ts];

    case "absensiGuru":
      return [d.id, d.teacherId, d.namaGuru || d["Nama Guru"] || "", d.kelas, d.tanggal, d.status, ts];

    case "jurnal":
      return [d.id, d.teacherId, d.namaGuru || d["Nama Guru"] || "", d.kelas, d.mataPelajaran, d.tanggal, d.topik, d.kegiatan || "", d.catatan || "", ts];

    case "harian":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.tanggal, d.bahasan, d.halaman || "", d.hasil, d.catatan || "", ts];

    case "nilai":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.mataPelajaran, d.jenisNilai, d.skor, d.tanggal, ts];

    case "kegiatan":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.kategori || d.mataPelajaran, d.penilaian, d.tanggal, ts];

    case "target":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.mataPelajaran,
              d.progressSaatIni || 0, d.progressTotal || 100, d.satuan || "",
              d.hasil, d.catatan || "", d.tanggal, ts];

    case "ziyadah":
      return [d.id, d.studentId, d.namaSiswa || d["Nama Siswa"] || "", d.kelas, d.juz, d.surat || "", d.progress || "", d.hasil, d.catatan || "", d.tanggal, ts];

    case "guru":
      return [d.id, d.nip, d.nama, d.email || "", d.kelas, d.status, ts];

    case "siswa":
      return [d.id, d.nis, d.nama, d.class_name, ts];

    case "log":
      return [ts, d.aksi, d.sheet, d.data, d.user || "portal"];

    default:
      return Object.values(d);
  }
}

// ── Utilitas: test koneksi (jalankan manual dari Apps Script editor) ──────────
function testKoneksi() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log("✅ Koneksi berhasil ke: " + ss.getName());
  Logger.log("📋 Sheet yang ada: " + ss.getSheets().map(s => s.getName()).join(", "));
}

// ╔══════════════════════════════════════════════════════════════╗
// ║              MENU BUNAYYA DI GOOGLE SHEET                   ║
// ╚══════════════════════════════════════════════════════════════╝

// ── Buat menu otomatis saat Spreadsheet dibuka ────────────────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🕌 Bunayya")
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("⚙️ Setup")
        .addItem("🚀 Inisialisasi Semua Sheet",        "menuSetupSheets")
        .addItem("🎨 Rapikan Format Semua Sheet",       "menuRapikanFormat")
        .addItem("🔗 Tampilkan Web App URL",            "menuTampilkanUrl")
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("📊 Laporan")
        .addItem("📋 Rekap Absensi Siswa",             "menuRekapAbsensiSiswa")
        .addItem("🗒️ Rekap Absensi Guru",              "menuRekapAbsensiGuru")
        .addItem("📈 Ringkasan Nilai per Kelas",        "menuRingkasanNilai")
        .addItem("📖 Rekap Ziyadah per Kelas",          "menuRekapZiyadah")
        .addItem("🎯 Rekap Target Pencapaian",          "menuRekapTarget")
    )
    .addSeparator()
    .addSubMenu(
      SpreadsheetApp.getUi().createMenu("🛠️ Utilitas")
        .addItem("✅ Test Koneksi",                    "menuTestKoneksi")
        .addItem("🗑️ Kosongkan Sheet (Hapus Data)",    "menuKosongkanSheet")
        .addItem("📤 Export Log ke Sheet Baru",         "menuExportLog")
        .addItem("🔢 Hitung Total Data",               "menuHitungTotal")
    )
    .addToUi();
}

// ══════════════════════════════════════════════════════════════
// ⚙️  SETUP MENU HANDLERS
// ══════════════════════════════════════════════════════════════

function menuSetupSheets() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "🚀 Inisialisasi Sheet",
    "Fungsi ini akan membuat semua sheet yang diperlukan portal Bunayya.\n\nLanjutkan?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dibuat = [], sudahAda = [];

  Object.values(SHEETS).forEach(s => {
    const existing = ss.getSheetByName(s.name);
    if (existing) {
      sudahAda.push(s.name);
    } else {
      getOrCreateSheet(ss, s.name, s.headers);
      dibuat.push(s.name);
    }
  });

  const defaultSheet = ss.getSheetByName("Sheet1");
  if (defaultSheet && ss.getSheets().length > 1) ss.deleteSheet(defaultSheet);

  let msg = "✅ Setup selesai!\n\n";
  if (dibuat.length)   msg += "📋 Dibuat baru:\n• " + dibuat.join("\n• ") + "\n\n";
  if (sudahAda.length) msg += "✔️ Sudah ada:\n• " + sudahAda.join("\n• ");

  ui.alert("Setup Selesai", msg, ui.ButtonSet.OK);
}

function menuRapikanFormat() {
  const ss = SpreadsheetApp.getUi();
  const sp = SpreadsheetApp.openById(SPREADSHEET_ID);

  Object.values(SHEETS).forEach(cfg => {
    const sheet = sp.getSheetByName(cfg.name);
    if (!sheet) return;
    // Re-apply header style
    styleHeader(sheet, cfg.headers.length);
    // Alternating row colors
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      for (let r = 2; r <= lastRow; r++) {
        const bg = r % 2 === 0 ? "#f0fdf4" : "#ffffff";
        sheet.getRange(r, 1, 1, cfg.headers.length).setBackground(bg);
      }
    }
    // Auto-resize
    sheet.autoResizeColumns(1, cfg.headers.length);
  });

  SpreadsheetApp.getUi().alert("✅ Format semua sheet berhasil dirapikan!", "", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuTampilkanUrl() {
  const url = ScriptApp.getService().getUrl();
  const ui  = SpreadsheetApp.getUi();
  if (!url) {
    ui.alert(
      "⚠️ Belum Di-Deploy",
      "Script belum di-deploy sebagai Web App.\n\nLakukan:\nDeploy → New Deployment → Web App → Anyone → Deploy",
      ui.ButtonSet.OK
    );
    return;
  }
  ui.alert(
    "🔗 Web App URL",
    "Salin URL berikut dan tempel ke portal:\n\n" + url + "\n\nTempel ke variabel SCRIPT_URL di bunayya-portal.jsx",
    ui.ButtonSet.OK
  );
}

// ══════════════════════════════════════════════════════════════
// 📊  LAPORAN MENU HANDLERS
// ══════════════════════════════════════════════════════════════

function menuRekapAbsensiSiswa() {
  const sp    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sp.getSheetByName(SHEETS.absensiSiswa.name);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data absensi siswa.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1);

  // Hitung per kelas per status
  const kelasIdx  = headers.indexOf("Kelas");
  const statusIdx = headers.indexOf("Status");
  const namaIdx   = headers.indexOf("Nama Siswa");

  const rekap = {};
  data.forEach(row => {
    const kelas  = row[kelasIdx]  || "—";
    const status = row[statusIdx] || "—";
    if (!rekap[kelas]) rekap[kelas] = { Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0, total: 0 };
    rekap[kelas][status] = (rekap[kelas][status] || 0) + 1;
    rekap[kelas].total++;
  });

  // Buat atau ganti sheet Rekap Absensi
  let rekapSheet = sp.getSheetByName("📋 Rekap Absensi Siswa");
  if (rekapSheet) sp.deleteSheet(rekapSheet);
  rekapSheet = sp.insertSheet("📋 Rekap Absensi Siswa");

  const rekapHeaders = ["Kelas", "Hadir", "Sakit", "Izin", "Alpa", "Total", "% Hadir"];
  rekapSheet.appendRow(rekapHeaders);
  styleHeader(rekapSheet, rekapHeaders.length);

  Object.entries(rekap)
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([kelas, d]) => {
      const pct = d.total > 0 ? ((d.Hadir / d.total) * 100).toFixed(1) + "%" : "0%";
      rekapSheet.appendRow([kelas, d.Hadir || 0, d.Sakit || 0, d.Izin || 0, d.Alpa || 0, d.total, pct]);
    });

  // Total row
  const lastR = rekapSheet.getLastRow();
  rekapSheet.appendRow(["TOTAL",
    `=SUM(B2:B${lastR})`, `=SUM(C2:C${lastR})`,
    `=SUM(D2:D${lastR})`, `=SUM(E2:E${lastR})`,
    `=SUM(F2:F${lastR})`, `=IFERROR(TEXT(B${lastR+1}/F${lastR+1},"0.0%"),"—")`
  ]);
  rekapSheet.getRange(lastR + 1, 1, 1, rekapHeaders.length)
    .setFontWeight("bold").setBackground("#fff3cd");

  rekapSheet.autoResizeColumns(1, rekapHeaders.length);
  sp.setActiveSheet(rekapSheet);
  SpreadsheetApp.getUi().alert("✅ Rekap absensi siswa berhasil dibuat!", "Sheet '📋 Rekap Absensi Siswa' sudah tersedia.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRekapAbsensiGuru() {
  const sp    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sp.getSheetByName(SHEETS.absensiGuru.name);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data absensi guru.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1);

  const namaIdx   = headers.indexOf("Nama Guru");
  const statusIdx = headers.indexOf("Status");

  const rekap = {};
  data.forEach(row => {
    const nama   = row[namaIdx]   || "—";
    const status = row[statusIdx] || "—";
    if (!rekap[nama]) rekap[nama] = { Hadir: 0, Sakit: 0, Izin: 0, Alpa: 0, total: 0 };
    rekap[nama][status] = (rekap[nama][status] || 0) + 1;
    rekap[nama].total++;
  });

  let rekapSheet = sp.getSheetByName("🗒️ Rekap Absensi Guru");
  if (rekapSheet) sp.deleteSheet(rekapSheet);
  rekapSheet = sp.insertSheet("🗒️ Rekap Absensi Guru");

  const rekapHeaders = ["Nama Guru", "Hadir", "Sakit", "Izin", "Alpa", "Total", "% Hadir"];
  rekapSheet.appendRow(rekapHeaders);
  styleHeader(rekapSheet, rekapHeaders.length);

  Object.entries(rekap).sort().forEach(([nama, d]) => {
    const pct = d.total > 0 ? ((d.Hadir / d.total) * 100).toFixed(1) + "%" : "0%";
    rekapSheet.appendRow([nama, d.Hadir || 0, d.Sakit || 0, d.Izin || 0, d.Alpa || 0, d.total, pct]);
  });

  rekapSheet.autoResizeColumns(1, rekapHeaders.length);
  sp.setActiveSheet(rekapSheet);
  SpreadsheetApp.getUi().alert("✅ Rekap absensi guru berhasil dibuat!", "", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRingkasanNilai() {
  const sp    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sp.getSheetByName(SHEETS.nilai.name);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data nilai.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1);

  const kelasIdx = headers.indexOf("Kelas");
  const mapelIdx = headers.indexOf("Mata Pelajaran");
  const skorIdx  = headers.indexOf("Skor");
  const namaIdx  = headers.indexOf("Nama Siswa");

  // Rata-rata per kelas per mapel
  const rekap = {}; // { kelas: { mapel: [skor,...] } }
  data.forEach(row => {
    const kelas = row[kelasIdx] || "—";
    const mapel = row[mapelIdx] || "—";
    const skor  = parseFloat(row[skorIdx]) || 0;
    if (!rekap[kelas]) rekap[kelas] = {};
    if (!rekap[kelas][mapel]) rekap[kelas][mapel] = [];
    rekap[kelas][mapel].push(skor);
  });

  let rekapSheet = sp.getSheetByName("📈 Ringkasan Nilai");
  if (rekapSheet) sp.deleteSheet(rekapSheet);
  rekapSheet = sp.insertSheet("📈 Ringkasan Nilai");

  rekapSheet.appendRow(["Kelas", "Mata Pelajaran", "Jumlah Data", "Nilai Terendah", "Nilai Tertinggi", "Rata-rata"]);
  styleHeader(rekapSheet, 6);

  Object.entries(rekap).sort().forEach(([kelas, mapels]) => {
    Object.entries(mapels).sort().forEach(([mapel, scores]) => {
      const min = Math.min(...scores);
      const max = Math.max(...scores);
      const avg = (scores.reduce((a, b) => a + b, 0) / scores.length).toFixed(1);
      rekapSheet.appendRow([kelas, mapel, scores.length, min, max, avg]);
    });
  });

  rekapSheet.autoResizeColumns(1, 6);
  sp.setActiveSheet(rekapSheet);
  SpreadsheetApp.getUi().alert("✅ Ringkasan nilai berhasil dibuat!", "Sheet '📈 Ringkasan Nilai' sudah tersedia.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRekapZiyadah() {
  const sp    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sp.getSheetByName(SHEETS.ziyadah.name);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data ziyadah.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1);

  const kelasIdx  = headers.indexOf("Kelas");
  const hasilIdx  = headers.indexOf("Hasil");
  const namaIdx   = headers.indexOf("Nama Siswa");

  const rekap = {};
  data.forEach(row => {
    const kelas = row[kelasIdx] || "—";
    const hasil = row[hasilIdx] || "—";
    if (!rekap[kelas]) rekap[kelas] = { total: 0, hasil: {} };
    rekap[kelas].total++;
    rekap[kelas].hasil[hasil] = (rekap[kelas].hasil[hasil] || 0) + 1;
  });

  let rekapSheet = sp.getSheetByName("📖 Rekap Ziyadah");
  if (rekapSheet) sp.deleteSheet(rekapSheet);
  rekapSheet = sp.insertSheet("📖 Rekap Ziyadah");

  rekapSheet.appendRow(["Kelas", "Total Setoran", "Sangat Baik", "Baik", "Cukup", "Kurang"]);
  styleHeader(rekapSheet, 6);

  Object.entries(rekap).sort().forEach(([kelas, d]) => {
    rekapSheet.appendRow([
      kelas, d.total,
      d.hasil["Sangat Baik"] || 0,
      d.hasil["Baik"]        || 0,
      d.hasil["Cukup"]       || 0,
      d.hasil["Kurang"]      || 0
    ]);
  });

  rekapSheet.autoResizeColumns(1, 6);
  sp.setActiveSheet(rekapSheet);
  SpreadsheetApp.getUi().alert("✅ Rekap ziyadah berhasil dibuat!", "", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuRekapTarget() {
  const sp    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = sp.getSheetByName(SHEETS.target.name);
  if (!sheet || sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data target pencapaian.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const rows    = sheet.getDataRange().getValues();
  const headers = rows[0];
  const data    = rows.slice(1);

  const kelasIdx  = headers.indexOf("Kelas");
  const mapelIdx  = headers.indexOf("Mata Pelajaran");
  const hasilIdx  = headers.indexOf("Hasil");
  const progIdx   = headers.indexOf("Progress Saat Ini");
  const totalIdx  = headers.indexOf("Progress Total");

  const rekap = {};
  data.forEach(row => {
    const key   = (row[kelasIdx] || "—") + " | " + (row[mapelIdx] || "—");
    const hasil = row[hasilIdx] || "—";
    const prog  = parseFloat(row[progIdx]) || 0;
    const tot   = parseFloat(row[totalIdx]) || 0;
    if (!rekap[key]) rekap[key] = { kelas: row[kelasIdx], mapel: row[mapelIdx], totalSiswa: 0, totalProg: 0, totalTot: 0, hasil: {} };
    rekap[key].totalSiswa++;
    rekap[key].totalProg += prog;
    rekap[key].totalTot  += tot;
    rekap[key].hasil[hasil] = (rekap[key].hasil[hasil] || 0) + 1;
  });

  let rekapSheet = sp.getSheetByName("🎯 Rekap Target");
  if (rekapSheet) sp.deleteSheet(rekapSheet);
  rekapSheet = sp.insertSheet("🎯 Rekap Target");

  rekapSheet.appendRow(["Kelas", "Mata Pelajaran", "Jml Siswa", "Rata-rata Progress (%)", "Sangat Baik", "Baik", "Cukup", "Kurang"]);
  styleHeader(rekapSheet, 8);

  Object.entries(rekap).sort().forEach(([, d]) => {
    const avgPct = d.totalTot > 0 ? ((d.totalProg / d.totalTot) * 100).toFixed(1) + "%" : "0%";
    rekapSheet.appendRow([
      d.kelas, d.mapel, d.totalSiswa, avgPct,
      d.hasil["Sangat Baik"] || 0,
      d.hasil["Baik"]        || 0,
      d.hasil["Cukup"]       || 0,
      d.hasil["Kurang"]      || 0
    ]);
  });

  rekapSheet.autoResizeColumns(1, 8);
  sp.setActiveSheet(rekapSheet);
  SpreadsheetApp.getUi().alert("✅ Rekap target pencapaian berhasil dibuat!", "", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ══════════════════════════════════════════════════════════════
// 🛠️  UTILITAS MENU HANDLERS
// ══════════════════════════════════════════════════════════════

function menuTestKoneksi() {
  const sp = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheets = sp.getSheets().map(s => s.getName());
  const sheetDef = Object.values(SHEETS).map(s => s.name);
  const missing  = sheetDef.filter(n => !sheets.includes(n));
  const ada      = sheetDef.filter(n =>  sheets.includes(n));

  let msg = "✅ Koneksi OK — " + sp.getName() + "\n\n";
  msg += "📋 Sheet tersedia (" + ada.length + "/" + sheetDef.length + "):\n";
  ada.forEach(n => { msg += "  ✔ " + n + "\n"; });
  if (missing.length) {
    msg += "\n⚠️ Sheet belum dibuat:\n";
    missing.forEach(n => { msg += "  ✘ " + n + "\n"; });
    msg += "\nJalankan Setup → Inisialisasi Semua Sheet";
  }

  SpreadsheetApp.getUi().alert("🔍 Test Koneksi", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuKosongkanSheet() {
  const ui   = SpreadsheetApp.getUi();
  const resp = ui.alert(
    "⚠️ Konfirmasi Hapus Data",
    "Ini akan menghapus SEMUA DATA (kecuali header) di seluruh sheet portal.\n\nTindakan ini TIDAK DAPAT DIBATALKAN!\n\nLanjutkan?",
    ui.ButtonSet.YES_NO
  );
  if (resp !== ui.Button.YES) return;

  const resp2 = ui.alert(
    "⚠️ Konfirmasi Terakhir",
    "Anda yakin ingin menghapus semua data?",
    ui.ButtonSet.YES_NO
  );
  if (resp2 !== ui.Button.YES) return;

  const sp = SpreadsheetApp.openById(SPREADSHEET_ID);
  const dikosongkan = [];

  Object.values(SHEETS).forEach(cfg => {
    const sheet = sp.getSheetByName(cfg.name);
    if (!sheet || sheet.getLastRow() <= 1) return;
    sheet.deleteRows(2, sheet.getLastRow() - 1);
    dikosongkan.push(cfg.name);
  });

  ui.alert("🗑️ Selesai", "Data berhasil dikosongkan dari " + dikosongkan.length + " sheet:\n• " + dikosongkan.join("\n• "), ui.ButtonSet.OK);
}

function menuExportLog() {
  const sp        = SpreadsheetApp.openById(SPREADSHEET_ID);
  const logSheet  = sp.getSheetByName(SHEETS.log.name);
  if (!logSheet || logSheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("ℹ️", "Belum ada data log.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const timestamp  = Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyyMMdd_HHmm");
  const exportName = "Log " + timestamp;
  const existing   = sp.getSheetByName(exportName);
  if (existing) sp.deleteSheet(existing);

  const exportSheet = logSheet.copyTo(sp);
  exportSheet.setName(exportName);
  sp.setActiveSheet(exportSheet);

  SpreadsheetApp.getUi().alert("✅ Log Diekspor", "Sheet baru '" + exportName + "' berhasil dibuat sebagai salinan log.", SpreadsheetApp.getUi().ButtonSet.OK);
}

function menuHitungTotal() {
  const sp = SpreadsheetApp.openById(SPREADSHEET_ID);
  let msg  = "🔢 Total Data per Sheet\n" + "─".repeat(35) + "\n";
  let grandTotal = 0;

  Object.values(SHEETS).forEach(cfg => {
    const sheet = sp.getSheetByName(cfg.name);
    if (!sheet) { msg += `  ✘ ${cfg.name}: belum dibuat\n`; return; }
    const count = Math.max(0, sheet.getLastRow() - 1);
    grandTotal += count;
    msg += `  • ${cfg.name}: ${count} baris\n`;
  });

  msg += "─".repeat(35) + "\n";
  msg += `  TOTAL: ${grandTotal} baris`;

  SpreadsheetApp.getUi().alert("🔢 Hitung Total Data", msg, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ── Handle OPTIONS preflight (CORS) ──────────────────────────────────────────
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.TEXT)
    .addHeader("Access-Control-Allow-Origin", "*")
    .addHeader("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
    .addHeader("Access-Control-Allow-Headers", "Content-Type, Authorization");
}
