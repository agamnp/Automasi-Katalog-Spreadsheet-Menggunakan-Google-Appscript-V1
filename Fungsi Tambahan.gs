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

  //     ========     Fungsi Menggabungkan Semua Sheet Katalog menjadi 1 sheet     ========
    function Gabungsemuasheet() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const targetName = "Master";
      const logName = "LogProses";
      const sheetExcluded = ['Form Pengadaan', 'Hasil Seleksi', 'Referensi', 'DaftarISBN', 'DaftarUUID', targetName, logName];
      const publisherColumn = 20; // Kolom T

      const props = PropertiesService.getScriptProperties();
      const startIndex = Number(props.getProperty("sheetIndex")) || 0;

      const sheets = ss.getSheets().filter(s => !sheetExcluded.includes(s.getName()));
      if (startIndex >= sheets.length) {
        props.deleteProperty("sheetIndex"); // Semua sheet selesai
        return;
      }

      const targetSheet = ss.getSheetByName(targetName) || ss.insertSheet(targetName);
      const logSheet = ss.getSheetByName(logName) || ss.insertSheet(logName);

      // Batch pertama → reset Master & Log, tulis header
      if (startIndex === 0) {
        targetSheet.clearContents();
        logSheet.clearContents();

        const firstSheet = sheets[0];
        const lastCol = firstSheet.getLastColumn();
        const header = firstSheet.getRange(9, 1, 1, lastCol).getValues()[0];
        targetSheet.getRange(1, 1, 1, header.length).setValues([header]);

        logSheet.appendRow(["Timestamp", "Sheet", "Penerbit (Kolom T)", "Jumlah Baris", "Status"]);
      }

      const batchSize = 10;
      const endIndex = Math.min(startIndex + batchSize, sheets.length);
      const startTime = new Date().getTime();
      let output = [];
      let currentLastRow = targetSheet.getLastRow();

      for (let i = startIndex; i < endIndex; i++) {
        const sheet = sheets[i];
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();

        let rowsAdded = 0;
        let publisher = "";

        if (lastRow >= 10) {
          const data = sheet.getRange(10, 1, lastRow - 9, lastCol).getValues();

          // Filter hanya baris yang kolom C (kolom ke-3) ada isi
          const filteredData = data.filter(row => row[2] !== "" && row[2] !== null);

          rowsAdded = filteredData.length;
          if (rowsAdded > 0) publisher = filteredData[0][publisherColumn - 1] || "";

          output = output.concat(filteredData);
        }

        // Log sheet dengan status Done
        logSheet.appendRow([new Date(), sheet.getName(), publisher, rowsAdded, "Done"]);

        // Anti-timeout safety (opsional, 5 sheet pasti aman)
        const now = new Date().getTime();
        if (now - startTime > 250000) break;
      }

      // Tulis hasil batch ke Master
      if (output.length > 0) {
        targetSheet
          .getRange(currentLastRow + 1, 1, output.length, output[0].length)
          .setValues(output);
      }

      // Update progress
      props.setProperty("sheetIndex", endIndex);
   }
  //     ========     Fungsi Menggabungkan Semua Sheet Katalog menjadi 1 sheet     ========

  //     ========     Fungsi List Nama Sheet     ========
    function ListNamaSheet() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      
      // Nama sheet output
      const outputName = "DaftarSheet";
      let outputSheet = ss.getSheetByName(outputName);
      
      // Jika sheet belum ada, buat
      if (!outputSheet) {
        outputSheet = ss.insertSheet(outputName);
      }
      
      // Bersihkan isi lama
      outputSheet.clear();
      
      // Header
      outputSheet.getRange(1, 1, 1, 2).setValues([["No", "Nama Sheet"]]);
      
      // Isi data
      const data = sheets.map((sheet, index) => [
        index + 1,
        sheet.getName()
      ]);
      
      outputSheet.getRange(2, 1, data.length, 2).setValues(data);
   }
  //     ========     Fungsi List Nama Sheet     ========

  //     ========     Fungsi List Kategori     ========
    function rekapKategori() {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheets = ss.getSheets();
      const outputName = "DaftarKategori";
      
      let outputSheet = ss.getSheetByName(outputName);
      if (!outputSheet) {
        outputSheet = ss.insertSheet(outputName);
      }
      
      outputSheet.clear();
      outputSheet.getRange("A1:B1").setValues([["Kategori", "Jumlah"]]);
      
      const kategoriCount = {};
      
      sheets.forEach(sheet => {
        const name = sheet.getName();
        
        // Skip sheet output sendiri
        if (name === outputName) return;
        
        const lastRow = sheet.getLastRow();
        if (lastRow < 10) return;
        
        const values = sheet
          .getRange(10, 8, lastRow - 9, 1) // H10 ke bawah
          .getValues()
          .flat();
        
        values.forEach(val => {
          if (val !== "" && val !== null) {
            kategoriCount[val] = (kategoriCount[val] || 0) + 1;
          }
        });
      });
      
      const hasil = Object.entries(kategoriCount)
        .sort() // optional: urut alfabet
        .map(([kategori, jumlah]) => [kategori, jumlah]);
      
      if (hasil.length > 0) {
        outputSheet.getRange(2, 1, hasil.length, 2).setValues(hasil);
      }
    }
  //     ========     Fungsi List Kategori     ========
