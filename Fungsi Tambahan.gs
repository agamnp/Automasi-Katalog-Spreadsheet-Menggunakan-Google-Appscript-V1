  //     ========     Fungsi Cari UUID dari Katalog General     ========
    
    //buat sheet baru "CariUUID"
    //struktur tabel No	Judul	Pengarang	Penerbit	Perusahaan	E-ISBN	UUID
    
    function CariUUID() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const targetSheet = ss.getSheetByName("CariUUID"); // nama sheet hasil pencarian
      const targetData = targetSheet.getDataRange().getValues();

      const eisbnColTarget = 5; // kolom F = indeks ke-5 (0-based)
      const uuidColTarget = 6;  // kolom G = indeks ke-6 (0-based)
      const headerRowTarget = 1; // baris header di sheet CariUUID

      // Buat map ISBN → UUID dari semua sheet lain
      const sheets = ss.getSheets();
      const uuidMap = {};

      sheets.forEach(sh => {
        const name = sh.getName();
        if (name === "CariUUID") return; // skip sheet target

        const lastRow = sh.getLastRow();
        const lastCol = sh.getLastColumn();
        if (lastRow < 10) return; // sheet kosong / belum ada data

        const data = sh.getRange(10, 1, lastRow - 9, lastCol).getValues();

        const eisbnColSrc = 7;  // kolom G = ke-7
        const uuidColSrc = 29;  // kolom AC = ke-29

        for (let i = 0; i < data.length; i++) {
          const eisbn = data[i][eisbnColSrc - 1];
          const uuid = data[i][uuidColSrc - 1];
          if (eisbn && uuid) {
            uuidMap[eisbn.toString().trim()] = uuid;
          }
        }
      });

      // Isi UUID ke sheet target
      let filledCount = 0;
      for (let i = headerRowTarget; i < targetData.length; i++) {
        const eisbn = targetData[i][eisbnColTarget];
        if (eisbn && uuidMap[eisbn.toString().trim()]) {
          targetSheet.getRange(i + 1, uuidColTarget + 1)
            .setValue(uuidMap[eisbn.toString().trim()]);
          filledCount++;
        }
      }

      SpreadsheetApp.getActiveSpreadsheet()
        .toast(`✅ UUID berhasil diisi untuk ${filledCount} baris dari semua sheet!`, "Selesai");
    }
  //     ========     Fungsi Cari UUID dari Katalog General     ========

  //     ========     Fungsi Hapus Data Di Kolom Harga Satuan     ========

    function HapusDataHargaSatuan() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();

      const skipSheets = [
        "Form Pengadaan",
        "Hasil Seleksi",
        "Referensi",
        "DaftarISBN",
        "DaftarUUID"
      ];

      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        if (!skipSheets.includes(sheet.getName())) {

          const headerRow = 9;
          const lastColumn = sheet.getLastColumn();
          const lastRow = sheet.getLastRow();

          // Ambil header baris 9 untuk mencari kolom "HARGA SATUAN"
          const headers = sheet.getRange(headerRow, 1, 1, lastColumn).getValues()[0];
          const targetCol = headers.indexOf("HARGA SATUAN") + 1;

          if (targetCol > 0 && lastRow > headerRow) {
            // Hapus isi dari baris 10 hingga terakhir pada kolom "HARGA SATUAN"
            sheet.getRange(headerRow + 1, targetCol, lastRow - headerRow, 1).clearContent();
            Logger.log(`Data di sheet ${sheet.getName()} pada kolom HARGA SATUAN telah dihapus.`);
          }
        }
      });
    }

  //     ========     Fungsi Hapus Data Di Kolom Harga Satuan     ========

  //     ========     Fungsi Cek Ketersediaan di Katalog General Berdasarkan UUID     ========

    //buat sheet baru "CekKetersediaan"
    //struktur tabel no	judul	anak judul	ISBN Cetak	ISBN Elektronik	Perusahaan	UUID	Ketersediaan
  
    function CekKetersediaanFastbyUUID() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetCek = ss.getSheetByName("CekKetersediaan");
      const dataCek = sheetCek.getDataRange().getValues();

      const uuidIndexCek = 6;     // kolom F (index=5)
      const ketersediaanCol = 8;  // kolom H output

      const uuidColumnIndex = 29; // kolom AC
      const dataStartRow = 10;

      let uuidIndex = {};   // tempat menyimpan semua uuid
      Logger.log("🔍 Load semua UUID dari 95+ sheet...");

      // === Tahap 1 → Scan semua sheet SEKALI AJA ===
      for (let sh of ss.getSheets()) {
        if (sh.getName() === "CekKetersediaan") continue;

        const lastRow = sh.getLastRow();
        if (lastRow < dataStartRow) continue;

        const uuidList = sh.getRange(dataStartRow, uuidColumnIndex, lastRow - dataStartRow + 1, 1).getValues();

        for (let i = 0; i < uuidList.length; i++) {
          const uuid = uuidList[i][0];
          if (uuid) {
            uuidIndex[uuid.toString().trim()] = true; // hanya perlu tanda ada/ tidak
          }
        }
      }

      Logger.log("⚡ Index selesai dibuat. Mulai proses pengecekan...");

      let found = 0;

      // === Tahap 2 → Lookup cepat (tanpa scan ulang sheet) ===
      for (let i = 1; i < dataCek.length; i++) {
        const uuid = dataCek[i][uuidIndexCek];
        if (!uuid) continue;

        if (uuidIndex[uuid.toString().trim()]) {
          sheetCek.getRange(i+1, ketersediaanCol).setValue("Tersedia Di Katalog General");
          found++;
        }
      }

      Logger.log(`🚀 Selesai — UUID ditemukan pada ${found} baris.`);
      SpreadsheetApp.getActive().toast(`Selesai! ${found} UUID ditemukan.`);
    }
  //     ========     Fungsi Cek Ketersediaan di Katalog General Berdasarkan UUID     ========  