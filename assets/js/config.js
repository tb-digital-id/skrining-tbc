// ============================================================
//  config.js — KONFIGURASI TERPUSAT PUSKESMAS
//  Ganti isi variabel di bawah sesuai data Puskesmas lo!
// ============================================================

window.PUSKESMAS_CONFIG = {
    // Nama resmi Puskesmas
    namaPuskesmas: "Puskesmas Lok Bahu",

    // Nama petugas atau jabatan program TB
    namaPetugas: "Petugas TB Puskesmas Lok Bahu",

    // Nomor WhatsApp petugas (format internasional, tanpa + atau spasi)
    // Contoh: 6281234567890
    nomorWA: "6285753551401",

    // Alamat lengkap Puskesmas
    alamat: "Jl. Ir. H. Nusyirwan Ismail, M.Si. (Ring Road 2) RT. 15 (Belakang SMKN 15), Kelurahan Lok Bahu, Kecamatan Sungai Kunjang, Kota Samarinda, Kalimantan Timur 75129",

    // Jam pelayanan skrining & lab dahak TCM
    jamLayanan: "Senin \u2013 Sabtu, 07.30 \u2013 12.00 WITA",

    // ============================================================
    //  KONFIGURASI GOOGLE SHEETS (via Google Apps Script)
    //  Langkah: Deploy GAS sebagai Web App → salin URL di bawah
    // ============================================================
    // Ganti dengan URL Web App GAS lo setelah deploy
    // Contoh: "https://script.google.com/macros/s/AKfyc.../exec"
    // Biarkan string kosong ("") jika belum diatur / tidak dipakai
    gasEndpointURL: "https://script.google.com/macros/s/AKfycbza0met0LA34padiJ3i3gwjt_d0065sm_ff8-CQiw_W6gdOH26WM2P2608oJKkhNt1g/exec",

    // Fungsi generate pesan WhatsApp otomatis (jangan diubah)
    buatPesanWhatsApp(dataPasien, hasil) {
        let teks = `Halo Petugas ${this.namaPuskesmas}, saya telah mengisi formulir Skrining TB Mandiri melalui website.%0A%0A`;
        teks += `*Ringkasan Data Skrining:*%0A`;
        teks += `\u2022 Nama: ${dataPasien.nama || "-"}%0A`;
        teks += `\u2022 NIK: ${dataPasien.nik || "-"}%0A`;
        teks += `\u2022 Tanggal Lahir: ${dataPasien.tanggalLahir || "-"} (${dataPasien.usia || "-"} Tahun)%0A`;
        teks += `\u2022 No. HP: ${dataPasien.noHp || "-"}%0A`;
        teks += `\u2022 Tanggal Skrining: ${new Date().toLocaleDateString("id-ID", { day: "2-digit", month: "long", year: "numeric" })}%0A`;
        teks += `\u2022 Hasil Skrining: *${hasil.label ? hasil.label.toUpperCase() : "-"}*%0A`;
        if (hasil.alasan && hasil.alasan.length > 0) {
            teks += `\u2022 Gejala / Indikasi: ${hasil.alasan.join(", ")}%0A`;
        }
        teks += `\u2022 Riwayat Kontak TB: ${dataPasien.riwayatKontak === "ya" ? (dataPasien.jenisKontak === "serumah" ? "Kontak Serumah" : "Kontak Erat") : "Tidak Ada"}%0A`;
        teks += `%0AMohon arahan dan informasi mengenai jadwal pemeriksaan dahak (TCM) di ${this.namaPuskesmas}. Terima kasih.`;
        return teks;
    },

    // Fungsi buat link WhatsApp
    getLinkWA(dataPasien, hasil) {
        const pesan = this.buatPesanWhatsApp(dataPasien, hasil);
        return `https://wa.me/${this.nomorWA}?text=${pesan}`;
    },

    // ============================================================
    //  Fungsi kirim data ke Google Sheets via GAS
    //  Dipanggil otomatis setelah submit form (jika gasEndpointURL diisi)
    // ============================================================
    async kirimKeGoogleSheets(dataPasien, hasil) {
        if (!this.gasEndpointURL || this.gasEndpointURL.trim() === "") {
            console.info("[GAS] gasEndpointURL belum diisi. Data tidak dikirim ke Google Sheets.");
            return { sukses: false, pesan: "Endpoint GAS belum dikonfigurasi." };
        }

        // ============================================================
        //  Helper: konversi boolean null-safe ke label Ya/Tidak
        // ============================================================
        const yaAtauTidak = (val) => val === true ? "Ya" : val === false ? "Tidak" : "-";

        // ============================================================
        //  Gejala Klinis — Dewasa (3 pertanyaan gejala tambahan + status batuk)
        // ============================================================
        const isAnak = dataPasien.kategoriUsia === "anak";

        // Label status batuk (dewasa)
        const labelStatusBatuk = {
            tidak_batuk:  "Tidak Batuk",
            batuk_kurang2: "Batuk < 2 Minggu",
            batuk_lebih2:  "Batuk ≥ 2 Minggu",
        }[dataPasien.statusBatuk] || "-";

        const payload = {
            waktuSkrining: new Date().toLocaleString("id-ID"),
            nama: dataPasien.nama || "",
            nik: dataPasien.nik || "",
            tanggalLahir: dataPasien.tanggalLahir || "",
            usia: dataPasien.usia || "",
            kategoriUsia: hasil.kategoriUsia || "",
            jenisKelamin: dataPasien.jenisKelamin === "L" ? "Laki-laki" : "Perempuan",
            noHp: dataPasien.noHp || "",
            pekerjaan: dataPasien.pekerjaan || "",
            alamat: dataPasien.alamat || "",
            riwayatKontak: dataPasien.riwayatKontak || "tidak",
            jenisKontak: dataPasien.jenisKontak || "",
            namaKasusIndeks: dataPasien.namaKasusIndeks || "",
            beratBadan: dataPasien.beratBadan || "",
            tinggiBadan: dataPasien.tinggiBadan || "",
            imt: dataPasien.imt || "",
            kategoriIMT: dataPasien.kategoriIMT || "",
            riwayatTBC: dataPasien.riwayatTBC ? "Ya" : "Tidak",
            putusObat: dataPasien.putusObat ? "Ya" : "Tidak",
            merokok: dataPasien.merokok || "tidak",
            dm: dataPasien.dm || "tidak",
            odhiv: dataPasien.odhiv || "tidak",
            ibuHamil: dataPasien.ibuHamil ? "Ya" : "Tidak",
            wbp: dataPasien.wbp ? "Ya" : "Tidak",
            padatKumuh: dataPasien.padatKumuh ? "Ya" : "Tidak",
            hasilDiagnosis: hasil.label || "",
            statusTerduga: hasil.terduga ? "TERDUGA TB" : "BUKAN TERDUGA TB",
            alasanGejala: (hasil.alasan || []).join("; "),

            // ============================================================
            //  Gejala Klinis — BARU (sebelumnya tidak dikirim ke Sheets)
            //  Kolom baru di spreadsheet diperlukan mulai baris ini.
            //  Data lama (sebelum revisi) akan kosong di kolom-kolom ini.
            // ============================================================

            // --- Jalur Dewasa (≥ 15 tahun) ---
            // Diisi hanya jika responden dewasa; "-" jika anak / tidak relevan
            gejalaStatusBatuk:    isAnak ? "-" : labelStatusBatuk,
            gejalaDemamDewasa:    isAnak ? "-" : yaAtauTidak(dataPasien.demamDewasa),
            gejalaBBTurunDewasa:  isAnak ? "-" : yaAtauTidak(dataPasien.bbTurunDewasa),
            gejalaKeringatMalam:  isAnak ? "-" : yaAtauTidak(dataPasien.keringatMalam),

            // --- Jalur Anak (< 15 tahun) ---
            // Diisi hanya jika responden anak; "-" jika dewasa / tidak relevan
            gejalaBatukAnak:   isAnak ? yaAtauTidak(dataPasien.batukAnak)   : "-",
            gejalaDemamAnak:   isAnak ? yaAtauTidak(dataPasien.demamAnak)   : "-",
            gejalaBBTurunAnak: isAnak ? yaAtauTidak(dataPasien.bbTurunAnak) : "-",
            gejalaLesuAnak:    isAnak ? yaAtauTidak(dataPasien.lesuAnak)    : "-",
        };

        try {
            const response = await fetch(this.gasEndpointURL, {
                method: "POST",
                mode: "no-cors", // GAS CORS workaround
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });
            console.info("[GAS] Data berhasil dikirim ke Google Sheets.");
            return { sukses: true, pesan: "Data berhasil dikirim." };
        } catch (err) {
            console.error("[GAS] Gagal mengirim ke Google Sheets:", err);
            return { sukses: false, pesan: "Gagal mengirim data." };
        }
    }
};
