  //     ========     Hapus Data Menggunakan Filter 1 Sheet    ========
    function HapusDataDenganKriteria() {
      modulFungsiHapusDataDenganKriteria() ;    }
  //     ========     Hapus Data Menggunakan Filter 1 Sheet    ========
  
  //     ========     Hapus Data Menggunakan Filter Semua Sheet    ========
    function HapusDataDenganKriteriaALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
      modulFungsiHapusDataDenganKriteriaALLSHEET(sheetMulai) ;    }
  //     ========     Hapus Data Menggunakan Filter Semua Sheet    ========






  //     ========     Fungsi Utama Hapus Data Menggunakan Filter  1 Sheet   ========
    function modulFungsiHapusDataDenganKriteria() {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getActiveSheet();
        const name = sheet.getName();
        const namaPenerbit = name.replace(/^\d+\.\s*/, "").trim();
        

        // Lewati jika termasuk sheet yang di-skip
        if (SHEET_KECUALI.includes(name)) {
          Logger.log("⏭ Melewati sheet: " + name);
          return;
        }

        Logger.log("🔍 Memproses sheet: " + name);

        // Konfigurasi filter
        const filterOptions = {
          kodeRef: { aktif: false, nilaiDiperbolehkan: [''] },
          tahun: { aktif: false, min: 2021, max: 2025 },
          halaman: { aktif: false, min: 151, max: 2025 },
          harga: { aktif: false, min: 14800, max: 2025 },
          kategori: { aktif: false, nilaiDiperbolehkan: [''] },  // ['fiksi', 'bahasa', 'puisi','cerpen','novel']}
        };

        const startRow = 10;
        const headerRow = 9;
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();

        if (lastRow < startRow) {
          Logger.log("ℹ️ Tidak ada data yang diproses.");
          return;
        }

        // Ambil header & data
        const headers = sheet.getRange(headerRow, 1, 1, lastCol)
          .getValues()[0]
          .map(h => String(h).trim().toLowerCase());
        const data = sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).getValues();

        // Index kolom
        const getIndex = (colName) => headers.indexOf(colName.toLowerCase());
        const idx = {
          kodeRef: getIndex("kode referensi"),
          kategori: getIndex("kategori*"),
          tahun: getIndex("tahun terbit digital*"),
          halaman: getIndex("jumlah halaman*"),
          harga: getIndex("harga satuan")
        };

        // Fungsi pengecekan baris lolos filter
        const isLolosFilter = (row) => {
        let lolosKodeRef = false;
        let lolosKategori = false;

        // ===== KODE REFERENSI =====
        if (filterOptions.kodeRef.aktif && idx.kodeRef !== -1) {
          const val = String(row[idx.kodeRef]).trim();
          lolosKodeRef = filterOptions.kodeRef.nilaiDiperbolehkan.includes(val);
        }

        // ===== KATEGORI =====
        if (filterOptions.kategori.aktif) {
          if (idx.kategori !== -1) {
            const val = String(row[idx.kategori]).trim().toLowerCase();
            const allowed = filterOptions.kategori.nilaiDiperbolehkan.map(v => v.toLowerCase());
            lolosKategori = allowed.includes(val);
          }
        }

        // ===== OR LOGIC UTAMA =====
        if (filterOptions.kodeRef.aktif || filterOptions.kategori.aktif) {
          if (!(lolosKodeRef || lolosKategori)) return false;
        }

        // ===== FILTER LAIN (AND) =====
        if (filterOptions.tahun.aktif && idx.tahun !== -1) {
          const val = Number(row[idx.tahun]);
          if (isNaN(val) || val < filterOptions.tahun.min || val > filterOptions.tahun.max) return false;
        }

        if (filterOptions.halaman.aktif && idx.halaman !== -1) {
          const val = Number(row[idx.halaman]);
          if (isNaN(val) || val < filterOptions.halaman.min || val > filterOptions.halaman.max) return false;
        }

        if (filterOptions.harga.aktif && idx.harga !== -1) {
          const val = Number(row[idx.harga]);
          if (isNaN(val) || val < filterOptions.harga.min || val > filterOptions.harga.max) return false;
        }

        return true;
      };


        // Filter data
        const dataLolos = data.filter(isLolosFilter);

        // Jika semua baris gagal → hapus sheet
        if (dataLolos.length === 0) {
          const allSheets = ss.getSheets();
          if (allSheets.length > 1) {
            Logger.log("🗑 Semua data tidak memenuhi syarat. Menghapus sheet: " + name);
            modulHapusBarisDariHasilSeleksi(namaPenerbit);
            ss.deleteSheet(sheet);
          } else {
            Logger.log("⚠️ Tidak bisa menghapus sheet karena hanya ada satu sheet.");
          }
          return;
        }

        // Kosongkan isi data lama
        sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();

        // Tulis ulang data yang lolos
        sheet.getRange(startRow, 1, dataLolos.length, lastCol).setValues(dataLolos);

        // Bersihkan baris kosong & format di bawah data
        const lastDataRow = startRow + dataLolos.length - 1;
        if (lastDataRow < sheet.getMaxRows()) {
          const rangeToClear = sheet.getRange(lastDataRow + 1, 1, sheet.getMaxRows() - lastDataRow, lastCol);
          rangeToClear.clear({ contentsOnly: false }).setBorder(false, false, false, false, false, false);
          Logger.log(`🧹 Menghapus format & isi dari ${sheet.getMaxRows() - lastDataRow} baris di bawah data.`);
        }

        // Update tampilan sheet
        modulFungsiTampilanSheetPenerbit();
        Logger.log(`✅ Selesai. Dihapus ${data.length - dataLolos.length} baris, tersisa ${dataLolos.length}.`);
    }

    function modulHapusBarisDariHasilSeleksi(namaPenerbit) {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Hasil Seleksi");
        const data = sheet.getDataRange().getValues();

        const barisUntukHapus = [];

        for (let i = 0; i < data.length; i++) {
          const nama = String(data[i][1] || "").replace(/^\d+\.\s*/, "").trim();
          if (nama.toLowerCase() === namaPenerbit.toLowerCase()) {
            barisUntukHapus.push(i + 1);
          }
        }

        barisUntukHapus.reverse().forEach(row => sheet.deleteRow(row));
    }
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter  1 Sheet   ========

  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Semua Sheet    ======== 
      function modulFungsiHapusDataDenganKriteriaALLSHEET(sheetMulai) {
        const spreadsheet = SpreadsheetApp.getActive();
        const semuaSheet = spreadsheet.getSheets();
        const mulai = sheetMulai + 2;
        const sheetDikecualikan = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID'];
        if (mulai < 0 || mulai >= semuaSheet.length) {
          Logger.log("Nomor sheet tidak valid.");
          return;
        }
        semuaSheet.slice(mulai).forEach(sheet => {
          const namaSheet = sheet.getName();
          if (sheetDikecualikan.includes(namaSheet)) return;
          spreadsheet.setActiveSheet(sheet);
          modulFungsiHapusDataDenganKriteria();
        });
      }
  //     ========     Fungsi Utama Hapus Data Menggunakan Filter Semua Sheet    ========  