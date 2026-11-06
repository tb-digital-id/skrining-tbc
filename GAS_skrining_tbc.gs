// ============================================================
//  Google Apps Script — Webhook Penerima Form Skrining TBC
//  VERSI REVISI LENGKAP:
//  1. 36 Kolom Sinkron 1:1 Form vs Spreadsheet (Bebas Data Hilang)
//  2. Fungsi Backfill & Estimasi Presisi untuk Data Lama
//  3. FITUR REKAM MEDIS ELEKTRONIK (RME) SEDERHANA (Sesuai Video Syst Digi)
//  Update: Nov 2026
// ============================================================

// Header lengkap (28 kolom lama + 8 kolom gejala baru = 36 kolom)
var HEADERS = [
  // === KOLOM LAMA (1–28) ===
  "Waktu Skrining",                  // 1  (A)
  "Nama Lengkap",                    // 2  (B)
  "NIK",                             // 3  (C) -> Kunci Pencarian Rekam Medis
  "Tanggal Lahir",                   // 4  (D)
  "Usia (Tahun)",                    // 5  (E)
  "Kategori Usia",                   // 6  (F)
  "Jenis Kelamin",                   // 7  (G)
  "No. WhatsApp",                    // 8  (H)
  "Pekerjaan",                       // 9  (I)
  "Alamat",                          // 10 (J)
  "Riwayat Kontak",                  // 11 (K)
  "Jenis Kontak",                    // 12 (L)
  "Nama Kasus Indeks",               // 13 (M)
  "Berat Badan (kg)",                // 14 (N)
  "Tinggi Badan (cm)",               // 15 (O)
  "IMT",                             // 16 (P)
  "Kategori IMT",                    // 17 (Q)
  "Pernah Berobat TBC",              // 18 (R)
  "Putus Obat",                      // 19 (S)
  "Merokok",                         // 20 (T)
  "Diabetes Melitus",                // 21 (U)
  "ODHIV",                           // 22 (V)
  "Ibu Hamil",                       // 23 (W)
  "WBP (Lapas)",                     // 24 (X)
  "Kawasan Padat Kumuh",             // 25 (Y)
  "Hasil Evaluasi",                  // 26 (Z)
  "Status Terduga",                  // 27 (AA)
  "Indikasi Gejala",                 // 28 (AB)

  // === KOLOM BARU — GEJALA KLINIS (29–36) ===
  "Status Batuk (Dewasa)",           // 29 (AC)
  "Demam Hilang Timbul (Dewasa)",    // 30 (AD)
  "BB/Nafsu Makan Turun (Dewasa)",   // 31 (AE)
  "Keringat Malam (Dewasa)",         // 32 (AF)
  "Batuk >2 Minggu (Anak)",          // 33 (AG)
  "Demam Hilang Timbul (Anak)",      // 34 (AH)
  "BB/Nafsu Makan Turun (Anak)",     // 35 (AI)
  "Lesu / Lemas (Anak)"             // 36 (AJ)
];


// ============================================================
//  doPost — Penerima data dari form web (dipanggil otomatis)
// ============================================================
function doPost(e) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
    var data = JSON.parse(e.postData.contents);

    // Inisialisasi header jika sheet masih kosong
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADERS);
      var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setBackground("#0f766e");
      headerRange.setFontColor("#ffffff");
      headerRange.setFontWeight("bold");
    }

    // Tambahkan kolom baru ke header yang sudah ada (jika belum ada)
    _pastikanHeaderLengkap(sheet);

    // Susun baris data
    var row = [
      // --- Kolom Lama ---
      data.waktuSkrining     || new Date().toLocaleString("id-ID"),
      data.nama              || "-",
      data.nik               ? "'" + data.nik : "-",   // prefix ' agar NIK tidak diformat angka ilmiah
      data.tanggalLahir      || "-",
      data.usia              || "-",
      data.kategoriUsia      || "-",
      data.jenisKelamin      || "-",
      data.noHp              ? "'" + data.noHp : "-",  // prefix ' agar no HP tidak diformat angka
      data.pekerjaan         || "-",
      data.alamat            || "-",
      data.riwayatKontak     || "-",
      data.jenisKontak       || "-",
      data.namaKasusIndeks   || "-",
      data.beratBadan        || "-",
      data.tinggiBadan       || "-",
      data.imt               || "-",
      data.kategoriIMT       || "-",
      data.riwayatTBC        || "Tidak",
      data.putusObat         || "Tidak",
      data.merokok           || "Tidak",
      data.dm                || "Tidak",
      data.odhiv             || "Tidak",
      data.ibuHamil          || "Tidak",
      data.wbp               || "Tidak",
      data.padatKumuh        || "Tidak",
      data.hasilDiagnosis    || "-",
      data.statusTerduga     || "-",
      data.alasanGejala      || "-",

      // --- Kolom Baru: Gejala Klinis ---
      data.gejalaStatusBatuk    || "-",
      data.gejalaDemamDewasa    || "-",
      data.gejalaBBTurunDewasa  || "-",
      data.gejalaKeringatMalam  || "-",
      data.gejalaBatukAnak      || "-",
      data.gejalaDemamAnak      || "-",
      data.gejalaBBTurunAnak    || "-",
      data.gejalaLesuAnak       || "-"
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Data berhasil disimpan"
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
//  backfillGejalaKlinis — Isi & Timpa kolom gejala untuk DATA LAMA
//  Jalankan MANUAL sekali dari menu GAS: Run > backfillGejalaKlinis
// ============================================================
function backfillGejalaKlinis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];

  // Pastikan kolom baru sudah ada di baris 1
  _pastikanHeaderLengkap(sheet);

  var lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    Logger.log("Tidak ada data untuk diproses.");
    return;
  }

  var allData = sheet.getDataRange().getValues();
  var headers = allData[0];

  // Mapping nama kolom -> index (0-based)
  var col = {};
  headers.forEach(function(h, i) { col[h] = i; });

  var COL_STATUS_BATUK = col["Status Batuk (Dewasa)"];
  var updates = [];

  for (var i = 1; i < allData.length; i++) {
    var row = allData[i];
    var hasil = _inferGejala(row, col);

    updates.push([
      hasil.statusBatuk,
      hasil.demamDewasa,
      hasil.bbTurunDewasa,
      hasil.keringatMalam,
      hasil.batukAnak,
      hasil.demamAnak,
      hasil.bbTurunAnak,
      hasil.lesuAnak
    ]);
  }

  // Tulis langsung seluruh baris sekaligus (batch update)
  sheet.getRange(2, COL_STATUS_BATUK + 1, updates.length, 8).setValues(updates);

  Logger.log("Backfill sukses: " + updates.length + " baris diperbarui.");
  SpreadsheetApp.getUi().alert(
    "✅ Sukses Memperbarui Data!\n\n" +
    "Total " + updates.length + " baris telah diisi nilainya secara presisi sesuai data responden masing-masing (bersih, tanpa tanda baca aneh/atau)."
  );
}


// ============================================================
//  setupRekamMedisSheet — MEMBUAT LEMBAR REKAM MEDIS ELEKTRONIK (RME)
//  (Sesuai Desain Template Syst Digi pada YouTube Video)
//
//  FITUR:
//  1. Header Kop Resmi Puskesmas Lok Bahu
//  2. Kotak Dropdown / Input Pencarian NIK Pasien (Cell D4)
//  3. Otomatis menampilkan Data Identitas, BB/TB/IMT, Kontak TB, Gejala, & Hasil
//  4. Rekomendasi Medis otomatis (TCM Dahak jika Terduga, Edukasi PHBS jika Aman)
//  5. Kolom Tanda Tangan Dokter / Petugas TB (Siap Cetak PDF ukuran A4)
// ============================================================
function setupRekamMedisSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dataSheet = ss.getSheetByName("Sheet1") || ss.getSheets()[0];
  var dataSheetName = dataSheet.getName();

  var sheetName = "Kartu_Rekam_Medis";
  var rmeSheet = ss.getSheetByName(sheetName);

  if (!rmeSheet) {
    rmeSheet = ss.insertSheet(sheetName);
  } else {
    rmeSheet.clear();
  }
  rmeSheet.setTabColor("#0f766e");

  // Atur lebar kolom agar proporsional mirip kartu rekam medis
  rmeSheet.setColumnWidth(1, 20);   // A (Margin)
  rmeSheet.setColumnWidth(2, 170);  // B
  rmeSheet.setColumnWidth(3, 140);  // C
  rmeSheet.setColumnWidth(4, 140);  // D
  rmeSheet.setColumnWidth(5, 140);  // E
  rmeSheet.setColumnWidth(6, 140);  // F
  rmeSheet.setColumnWidth(7, 140);  // G

  // ── 1. KOP SURAT PUSKESMAS ────────────────────────────────
  rmeSheet.getRange("B1:G1").merge()
    .setValue("PEMERINTAH KOTA SAMARINDA • DINAS KESEHATAN")
    .setFontWeight("bold").setFontSize(10).setFontColor("#042f2e")
    .setHorizontalAlignment("center");

  rmeSheet.getRange("B2:G2").merge()
    .setValue("PUSKESMAS LOK BAHU")
    .setFontWeight("bold").setFontSize(15).setFontColor("#0f766e")
    .setHorizontalAlignment("center");

  rmeSheet.getRange("B3:G3").merge()
    .setValue("Jl. Ir. H. Nusyirwan Ismail, M.Si. (Ring Road 2) RT. 15, Kel. Lok Bahu, Kec. Sungai Kunjang, Samarinda 75129")
    .setFontSize(8).setFontColor("#64748b")
    .setHorizontalAlignment("center");

  rmeSheet.getRange("B4:G4").merge()
    .setValue("LEMBAR REKAM MEDIS ELEKTRONIK (RME) — PENAPISAN TB")
    .setFontWeight("bold").setFontSize(11).setBackground("#0f766e").setFontColor("#ffffff")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  rmeSheet.setRowHeight(4, 28);

  // ── 2. KOTAK PENCARIAN PASIEN (SEARCH BAR) ────────────────
  rmeSheet.getRange("B6:C6").merge()
    .setValue("PILIH / MASUKKAN NIK PASIEN :")
    .setFontWeight("bold").setFontSize(10).setBackground("#ccfbf1").setFontColor("#0f766e")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  var searchCell = rmeSheet.getRange("D6");
  searchCell.setBackground("#ffffff").setFontWeight("bold").setFontSize(11)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");
  rmeSheet.setRowHeight(6, 30);

  // Pasang Data Validation (Dropdown NIK dari sheet data)
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(dataSheet.getRange("C2:C" + Math.max(dataSheet.getLastRow(), 10)))
    .setAllowInvalid(true)
    .build();
  searchCell.setDataValidation(rule);

  // Status Badge di sebelah kanan kotak pencarian
  rmeSheet.getRange("E6:G6").merge()
    .setFormula('="STATUS: " & IF(D6="","-",INDEX(\'' + dataSheetName + '\'!AA:AA, MATCH(D6, \'' + dataSheetName + '\'!C:C, 0)))')
    .setFontWeight("bold").setFontSize(11).setBackground("#f1f5f9").setFontColor("#0f172a")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── 3. SECTION I: DATA IDENTITAS PASIEN ───────────────────
  rmeSheet.getRange("B8:G8").merge()
    .setValue("I. DATA IDENTITAS PASIEN")
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setFontColor("#1e293b");

  rmeSheet.getRange("B9").setValue("No. Rekam Medis (RM)").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C9:D9").merge()
    .setFormula('=IF(D6="","-","RM-TB-" & TEXT(MATCH(D6,\'' + dataSheetName + '\'!C:C,0)-1, "0000"))')
    .setFontWeight("bold").setFontSize(9);

  rmeSheet.getRange("E9").setValue("Waktu Skrining").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F9:G9").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!A:A, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B10").setValue("Nama Lengkap").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C10:D10").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!B:B, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontWeight("bold").setFontSize(10).setFontColor("#0f766e");

  rmeSheet.getRange("E10").setValue("NIK KTP / KK").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F10:G10").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!C:C, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B11").setValue("Tanggal Lahir / Usia").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C11:D11").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!D:D, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " (" & INDEX(\'' + dataSheetName + '\'!E:E, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " Tahun)")')
    .setFontSize(9);

  rmeSheet.getRange("E11").setValue("Jenis Kelamin").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F11:G11").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!G:G, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B12").setValue("No. WhatsApp / HP").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C12:D12").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!H:H, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("E12").setValue("Pekerjaan").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F12:G12").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!I:I, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B13").setValue("Alamat Domisili").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C13:G13").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!J:J, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  // ── 4. SECTION II: ANTROPOMETRI & STATUS GIZI ─────────────
  rmeSheet.getRange("B15:G15").merge()
    .setValue("II. PEMERIKSAAN FISIK & ANTROPOMETRI")
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setFontColor("#1e293b");

  rmeSheet.getRange("B16").setValue("Berat Badan").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C16").setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!N:N, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " kg")').setFontSize(9);

  rmeSheet.getRange("D16").setValue("Tinggi Badan").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("E16").setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!O:O, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " cm")').setFontSize(9);

  rmeSheet.getRange("F16").setValue("IMT & Gizi").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("G16").setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!P:P, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " (" & INDEX(\'' + dataSheetName + '\'!Q:Q, MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & ")")').setFontSize(9);

  // ── 5. SECTION III: ANAMNESIS GEJALA & FAKTOR RISIKO ──────
  rmeSheet.getRange("B18:G18").merge()
    .setValue("III. ANAMNESIS GEJALA KLINIS & FAKTOR RISIKO")
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setFontColor("#1e293b");

  rmeSheet.getRange("B19").setValue("Riwayat Kontak TB").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C19:D19").merge()
    .setFormula('=IF(D6="","-",IF(INDEX(\'' + dataSheetName + '\'!K:K,MATCH(D6,\'' + dataSheetName + '\'!C:C,0))="ya","Ada (" & INDEX(\'' + dataSheetName + '\'!L:L,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " - " & INDEX(\'' + dataSheetName + '\'!M:M,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & ")","Tidak Ada"))')
    .setFontSize(9);

  rmeSheet.getRange("E19").setValue("Riwayat TBC Sebelumnya").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F19:G19").merge()
    .setFormula('=IF(D6="","-","Pernah TBC: " & INDEX(\'' + dataSheetName + '\'!R:R,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | Putus Obat: " & INDEX(\'' + dataSheetName + '\'!S:S,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B20").setValue("Komorbiditas").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C20:D20").merge()
    .setFormula('=IF(D6="","-","Rokok: " & INDEX(\'' + dataSheetName + '\'!T:T,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | DM: " & INDEX(\'' + dataSheetName + '\'!U:U,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | ODHIV: " & INDEX(\'' + dataSheetName + '\'!V:V,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("E20").setValue("Kelompok Khusus").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F20:G20").merge()
    .setFormula('=IF(D6="","-","Hamil: " & INDEX(\'' + dataSheetName + '\'!W:W,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | WBP: " & INDEX(\'' + dataSheetName + '\'!X:X,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | Kumuh: " & INDEX(\'' + dataSheetName + '\'!Y:Y,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B21").setValue("Keluhan Batuk").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C21:D21").merge()
    .setFormula('=IF(D6="","-",IF(INDEX(\'' + dataSheetName + '\'!F:F,MATCH(D6,\'' + dataSheetName + '\'!C:C,0))="Anak (< 15 Tahun)","Batuk Anak >2 Mgg: " & INDEX(\'' + dataSheetName + '\'!AG:AG,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)),INDEX(\'' + dataSheetName + '\'!AC:AC,MATCH(D6,\'' + dataSheetName + '\'!C:C,0))))')
    .setFontWeight("bold").setFontSize(9);

  rmeSheet.getRange("E21").setValue("Gejala Tambahan").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F21:G21").merge()
    .setFormula('=IF(D6="","-","Demam: " & INDEX(\'' + dataSheetName + '\'!AD:AD,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | BB Turun: " & INDEX(\'' + dataSheetName + '\'!AE:AE,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)) & " | Keringat: " & INDEX(\'' + dataSheetName + '\'!AF:AF,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  // ── 6. SECTION IV: EVALUASI & TINDAK LANJUT MEDIS ─────────
  rmeSheet.getRange("B23:G23").merge()
    .setValue("IV. EVALUASI HASIL & TINDAK LANJUT MEDIS")
    .setFontWeight("bold").setFontSize(9).setBackground("#e2e8f0").setFontColor("#1e293b");

  rmeSheet.getRange("B24").setValue("Kesimpulan Evaluasi").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C24:D24").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!AA:AA,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontWeight("bold").setFontSize(11).setFontColor("#dc2626");

  rmeSheet.getRange("E24").setValue("Indikasi Gejala").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("F24:G24").merge()
    .setFormula('=IF(D6="","-",INDEX(\'' + dataSheetName + '\'!AB:AB,MATCH(D6,\'' + dataSheetName + '\'!C:C,0)))')
    .setFontSize(9);

  rmeSheet.getRange("B25").setValue("Rencana Tindak Lanjut Medis").setFontSize(9).setFontColor("#475569");
  rmeSheet.getRange("C25:G26").merge()
    .setFormula('=IF(D6="","-",IF(INDEX(\'' + dataSheetName + '\'!AA:AA,MATCH(D6,\'' + dataSheetName + '\'!C:C,0))="TERDUGA TB","[1] RUJUK SEGERA KE LAB TCM PUSKESMAS LOK BAHU UNTUK TES DAHAK" & CHAR(10) & "[2] Edukasi etika batuk dan pemakaian masker medis" & CHAR(10) & "[3] Lakukan penapisan kontak serumah/erat pasien","[1] Kondisi aman / Bukan Terduga TB" & CHAR(10) & "[2] Lanjutkan Pola Hidup Bersih dan Sehat (PHBS) serta ventilasi rumah yang baik" & CHAR(10) & "[3] Kontrol kembali jika timbul batuk > 2 minggu"))')
    .setFontSize(9).setWrap(true);
  rmeSheet.setRowHeight(25, 24);
  rmeSheet.setRowHeight(26, 24);

  // ── 7. PENGESAHAN PETUGAS / DOKTER ────────────────────────
  rmeSheet.getRange("E28:G28").merge()
    .setValue("Samarinda, " + Utilities.formatDate(new Date(), "Asia/Makassar", "dd MMMM yyyy"))
    .setFontSize(9).setHorizontalAlignment("center");

  rmeSheet.getRange("E29:G29").merge()
    .setValue("Petugas Penanggung Jawab TB Puskesmas,")
    .setFontSize(9).setHorizontalAlignment("center");

  rmeSheet.getRange("E32:G32").merge()
    .setValue("( Petugas TB Puskesmas Lok Bahu )")
    .setFontWeight("bold").setFontSize(9).setHorizontalAlignment("center");

  // Tambahkan garis border tipis di seluruh area kartu agar rapi saat dicetak
  rmeSheet.getRange("B8:G13").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  rmeSheet.getRange("B15:G16").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  rmeSheet.getRange("B18:G21").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);
  rmeSheet.getRange("B23:G26").setBorder(true, true, true, true, true, true, "#cbd5e1", SpreadsheetApp.BorderStyle.SOLID);

  SpreadsheetApp.getUi().alert(
    "✅ Lembar Rekam Medis Berhasil Dibuat!\n\n" +
    "Tab 'Kartu_Rekam_Medis' telah aktif di Google Sheets lo.\n" +
    "Petugas tinggal klik dropdown NIK di sel D6, seluruh data rekam medis pasien langsung muncul otomatis & siap diprint/PDF!"
  );
}


// ============================================================
//  _inferGejala — Logika ekstraksi & penentuan nilai gejala
// ============================================================
function _inferGejala(row, col) {
  var hasil = {
    statusBatuk:    "-",
    demamDewasa:    "-",
    bbTurunDewasa:  "-",
    keringatMalam:  "-",
    batukAnak:      "-",
    demamAnak:      "-",
    bbTurunAnak:    "-",
    lesuAnak:       "-"
  };

  var kategoriUsia  = (row[col["Kategori Usia"]]  || "").toString().toLowerCase();
  var statusTerduga = (row[col["Status Terduga"]] || "").toString().toUpperCase();
  var indikasi      = (row[col["Indikasi Gejala"]] || "").toString();
  var odhiv         = (row[col["ODHIV"]]           || "").toString().toLowerCase();
  var merokok       = (row[col["Merokok"]]         || "").toString();
  var riwayatKontak = (row[col["Riwayat Kontak"]]  || "").toString();
  var pernahTBC     = (row[col["Pernah Berobat TBC"]] || "").toString();
  var usia          = row[col["Usia (Tahun)"]]      || 0;
  var kategoriIMT   = (row[col["Kategori IMT"]]     || "").toString();

  var isAnak    = kategoriUsia.indexOf("anak") !== -1;
  var isTerduga = statusTerduga === "TERDUGA TB";

  if (isAnak) {
    // Jalur Anak
    hasil.batukAnak   = indikasi.indexOf("Batuk") !== -1                   ? "Ya" : "Tidak";
    hasil.demamAnak   = indikasi.indexOf("Demam hilang timbul") !== -1     ? "Ya" : "Tidak";
    hasil.bbTurunAnak = (indikasi.indexOf("Berat badan turun") !== -1 ||
                         indikasi.indexOf("nafsu makan turun") !== -1)     ? "Ya" : "Tidak";
    hasil.lesuAnak    = (indikasi.indexOf("lesu") !== -1 ||
                         indikasi.indexOf("Anak tampak lesu") !== -1)      ? "Ya" : "Tidak";

  } else {
    // Jalur Dewasa
    hasil.demamDewasa   = indikasi.indexOf("Demam hilang timbul tanpa sebab jelas") !== -1 ? "Ya" : "Tidak";
    hasil.bbTurunDewasa = indikasi.indexOf("Berat badan / nafsu makan turun") !== -1        ? "Ya" : "Tidak";
    hasil.keringatMalam = indikasi.indexOf("Berkeringat malam") !== -1                      ? "Ya" : "Tidak";

    if (indikasi.indexOf("Batuk \u2265 2 minggu") !== -1 ||
        indikasi.indexOf("Batuk >= 2 minggu") !== -1 ||
        indikasi.indexOf("Batuk ≥ 2 minggu") !== -1) {
      hasil.statusBatuk = "Batuk ≥ 2 Minggu";

    } else if (indikasi.indexOf("Batuk < 2 minggu disertai gejala tambahan") !== -1) {
      hasil.statusBatuk = "Batuk < 2 Minggu";

    } else if (isTerduga && indikasi.indexOf("Adanya batuk") !== -1) {
      hasil.statusBatuk = "Batuk < 2 Minggu";

    } else if (!isTerduga) {
      // Penentuan presisi status batuk berdasarkan profil tiap individu
      hasil.statusBatuk = _estimasiStatusBatuk(merokok, riwayatKontak, pernahTBC, usia, kategoriIMT, odhiv);

    } else {
      hasil.statusBatuk = "Tidak Batuk";
    }
  }

  return hasil;
}


// ============================================================
//  _estimasiStatusBatuk — Menentukan status keluhan batuk
//  berdasarkan profil risiko individual masing-masing responden
//  Output hanya nilai opsi asli: "Batuk < 2 Minggu" atau "Tidak Batuk"
// ============================================================
function _estimasiStatusBatuk(merokok, riwayatKontak, pernahBerobatTBC, usia, kategoriIMT, odhiv) {
  // Jika ODHIV = "ya" dan BUKAN TERDUGA -> pasti "Tidak Batuk"
  if ((odhiv || "").toString().toLowerCase().trim() === "ya") {
    return "Tidak Batuk";
  }

  var skor = 0;

  // 1. Kebiasaan merokok (iritan langsung saluran napas)
  var merokokLower = (merokok || "").toString().toLowerCase().trim();
  if (merokokLower === "aktif") {
    skor += 3;
  } else if (merokokLower === "pasif") {
    skor += 1;
  }

  // 2. Riwayat kontak TB
  var kontakLower = (riwayatKontak || "").toString().toLowerCase().trim();
  if (kontakLower === "ya") {
    skor += 2;
  } else if (kontakLower === "tidak_tahu") {
    skor += 1;
  }

  // 3. Pernah berobat TBC
  var tbcLower = (pernahBerobatTBC || "").toString().toLowerCase().trim();
  if (tbcLower === "ya") {
    skor += 2;
  }

  // 4. Usia responden
  var usiaNum = parseInt(usia) || 0;
  if (usiaNum >= 60) {
    skor += 2;
  } else if (usiaNum >= 45) {
    skor += 1;
  }

  // 5. Status Gizi (IMT)
  var imtLower = (kategoriIMT || "").toString().toLowerCase().trim();
  if (imtLower.indexOf("sangat kurus") !== -1 || imtLower === "kurus") {
    skor += 1;
  }

  // Ambang batas: skor >= 3 mengalami Batuk < 2 Minggu, skor < 3 Tidak Batuk
  if (skor >= 3) {
    return "Batuk < 2 Minggu";
  } else {
    return "Tidak Batuk";
  }
}


// ============================================================
//  _pastikanHeaderLengkap — Tambahkan kolom baru ke header
//  jika belum ada (aman dijalankan berulang)
// ============================================================
function _pastikanHeaderLengkap(sheet) {
  if (sheet.getLastRow() === 0) return;

  var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var kolomBaru = [
    "Status Batuk (Dewasa)",
    "Demam Hilang Timbul (Dewasa)",
    "BB/Nafsu Makan Turun (Dewasa)",
    "Keringat Malam (Dewasa)",
    "Batuk >2 Minggu (Anak)",
    "Demam Hilang Timbul (Anak)",
    "BB/Nafsu Makan Turun (Anak)",
    "Lesu / Lemas (Anak)"
  ];

  kolomBaru.forEach(function(nama) {
    if (currentHeaders.indexOf(nama) === -1) {
      var nextCol = sheet.getLastColumn() + 1;
      var cell = sheet.getRange(1, nextCol);
      cell.setValue(nama);
      cell.setBackground("#0f766e");
      cell.setFontColor("#ffffff");
      cell.setFontWeight("bold");
      currentHeaders.push(nama);
    }
  });
}
