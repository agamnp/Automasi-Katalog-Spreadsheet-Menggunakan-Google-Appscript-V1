  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID 1 Sheet     ========
    function HapusPembeliansebelumdenganUUID(){
      ModulHapusSudahdibeli("UUID");
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID 1 Sheet     ========     

  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID Semua Sheet     ========
    function HapusPembeliansebelumdenganUUIDALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
      ModulHapusSudahdibeliALLSHEET("UUID",sheetMulai);
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar UUID Semua Sheet     ========  

  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN 1 Sheet     ========
    function HapusPembeliansebelumdenganISBN(){
      ModulHapusSudahdibeli("ISBN");
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN 1 Sheet     ======== 

  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN Semua Sheet     ========
    function HapusPembeliansebelumdenganISBNALLSHEET() {
    //PERLU DIRUBAH ============================================================================== 
    const sheetMulai = 0; 
    //PERLU DIRUBAH ==============================================================================
    
      ModulHapusSudahdibeliALLSHEET("ISBN",sheetMulai);
    }
  //     ========     MengHapus Pembelian Sebelumnya berdasar ISBN & e-ISBN Semua Sheet     ========






  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya 1 Sheet    ========
    function ModulHapusSudahdibeli(type) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetData = ss.getActiveSheet();
      const nama = sheetData.getName();

      // ✅ Skip sheet non-penerbit
      if (SHEET_KECUALI.includes(nama)) {
        Logger.log('Sheet dilewati: ' + nama);
        return;
      }

      Logger.log('▶ Menjalankan HapusPembelianSebelumnya pada: ' + nama);

      const headerRow = 9;
      const startRow = 10;
      const lastRow = sheetData.getLastRow();
      const numRows = lastRow - startRow + 1;

      if (numRows < 1) {
        Logger.log("⚠️ Tidak ada data untuk diproses.");
        return;
      }

      // ✅ Ambil header
      const headers = sheetData.getRange(headerRow, 1, 1, sheetData.getLastColumn())
        .getValues()[0]
        .map(h => String(h).toLowerCase().trim());
      const data = sheetData.getRange(startRow, 1, numRows, sheetData.getLastColumn()).getValues();

      let filteredData = [];

      if (type === "UUID") {
        const sheetUUID = ss.getSheetByName("DaftarUUID");
        if (!sheetUUID) throw new Error("Sheet 'DaftarUUID' tidak ditemukan.");
        const uuidRef = sheetUUID.getRange("A:A").getValues().flat().filter(val => val);
        const uuidSet = new Set(uuidRef.map(val => String(val).trim()));

        const uuidIndex = headers.findIndex(h => h.includes("uuid"));
        if (uuidIndex === -1) throw new Error(`Kolom UUID tidak ditemukan di baris ${headerRow}`);

        filteredData = data.filter(row => {
          const uuid = String(row[uuidIndex]).trim();
          return !uuidSet.has(uuid);
        });
      }

      if (type === "ISBN") {
        const sheetISBN = ss.getSheetByName("DaftarISBN");
        if (!sheetISBN) throw new Error("Sheet 'DaftarISBN' tidak ditemukan.");
        const rawISBNData = sheetISBN.getRange("A:B").getValues();
        const isbnSet = new Set();
        rawISBNData.forEach(([isbnCetak, isbnElektronik]) => {
          const clean1 = String(isbnCetak).replace(/\D/g, "");
          const clean2 = String(isbnElektronik).replace(/\D/g, "");
          if (clean1) isbnSet.add(clean1);
          if (clean2) isbnSet.add(clean2);
        });

        const isbnCetakIndex = headers.findIndex(h => h.includes("isbn cetak"));
        const isbnElektronikIndex = headers.findIndex(h => h.includes("isbn elektronik"));
        if (isbnCetakIndex === -1 && isbnElektronikIndex === -1) {
          throw new Error("Kolom ISBN tidak ditemukan di baris " + headerRow);
        }

        filteredData = data.filter(row => {
          const valCetak = isbnCetakIndex !== -1 ? String(row[isbnCetakIndex]).replace(/\D/g, "") : "";
          const valElektronik = isbnElektronikIndex !== -1 ? String(row[isbnElektronikIndex]).replace(/\D/g, "") : "";
          return !(isbnSet.has(valCetak) || isbnSet.has(valElektronik));
        });
      }

      // Kosongkan area lama
      sheetData.getRange(startRow, 1, numRows, sheetData.getLastColumn()).clearContent();

      // ✅ Tulis ulang
      if (filteredData.length > 0) {
        sheetData.getRange(startRow, 1, filteredData.length, sheetData.getLastColumn())
          .setValues(filteredData);
      }

      const deletedCount = data.length - filteredData.length;

      // ✅ Hapus sheet kalau kosong
      if (filteredData.length === 0) {
        const namaPenerbit = sheetData.getName().replace(/^\d+\.\s*/, "").trim();
        ss.deleteSheet(sheetData);
        ModulHapusBarisDariHasilSeleksi(namaPenerbit);
        Logger.log(`🗑 Sheet '${nama}' dihapus karena kosong.`);
        return;
      }

      Logger.log(`✅ [${type}] Selesai. Total: ${data.length}, dihapus: ${deletedCount}, tersisa: ${filteredData.length}`);
      modulFungsiTampilanSheetPenerbit();
    }
  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya 1 Sheet    ========

  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Semua Sheet    ========
    function ModulHapusSudahdibeliALLSHEET(mode,sheetMulai) {
      const spreadsheet = SpreadsheetApp.getActive();
      const semuaSheet = spreadsheet.getSheets();



        const mulai = sheetMulai + 2;
        const sheetDikecualikan = [
          'Form Pengadaan', 
          'Hasil Seleksi', 
          'Referensi', 
          'DaftarISBN', 
          'DaftarUUID'
        ];

        if (mulai < 0 || mulai >= semuaSheet.length) {
          Logger.log("❌ Nomor sheet tidak valid.");
          return;
        }
        semuaSheet.slice(mulai).forEach(sheet => {
          const namaSheet = sheet.getName();
          if (sheetDikecualikan.includes(namaSheet)) {
            Logger.log("➡️ Skip sheet: " + namaSheet);
            return;
          }
          spreadsheet.setActiveSheet(sheet);
          try {
            if (mode === "ISBN") {
              HapusPembeliansebelumdenganISBN();   // ✅ jalankan berdasarkan ISBN
            } else if (mode === "UUID") {
              HapusPembeliansebelumdenganUUID();           // ✅ jalankan berdasarkan UUID
            } else {
              Logger.log("⚠️ Mode tidak dikenali: " + mode);
            }
          } catch (err) {
            Logger.log("⚠️ Error di sheet " + namaSheet + ": " + err.message);
          }
        });
        Logger.log("✅ Selesai Hapus Data Sudah Di beli Menggunakan : " + mode);
    }

  //     ========     Fungsi Utama MengHapus Pembelian Sebelumnya Semua Sheet    ========