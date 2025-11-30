  //     ========     Fungsi MengHapus Sheet Penerbit PPN     ========
    function HapusSemuaSheetPPN() {
      ModulHapusSheetSaja("ppn");
      ModulHapusBarisSaja("ppn");
    }
  //     ========     Fungsi MengHapus Sheet Penerbit PPN     ========  

  //     ========     Fungsi MengHapus Sheet Penerbit Non PPN     ========
    function HapusSemuaSheetNonPPN() {
      ModulHapusSheetSaja("non ppn");
      ModulHapusBarisSaja("non ppn");
    }
  //     ========     Fungsi MengHapus Sheet Penerbit Non PPN     ======== 

  //     ========     Fungsi Utama MengHapus Sheet Penerbit PPN atau Non PPN     ========
    function ModulHapusSheetSaja(jenisYangDihapus) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheetSeleksi = ss.getSheetByName("Hasil Seleksi");
      var data = sheetSeleksi.getDataRange().getValues();
      var namaPenerbitYangAkanDihapus = {};

      for (var i = 0; i < data.length; i++) {
        var namaPenerbit = data[i][1]; // Kolom B
        var jenisPPN = data[i][3];     // Kolom D

        if (
          namaPenerbit &&
          jenisPPN &&
          jenisPPN.toLowerCase().trim() === jenisYangDihapus.toLowerCase().trim()
        ) {
          var namaTrim = namaPenerbit.trim();
          namaPenerbitYangAkanDihapus[namaTrim] = true;
        }
      }

      var allSheets = ss.getSheets();
      allSheets.forEach(function (sheet) {
        var rawName = sheet.getName();
        var cleanName = rawName.replace(/^\d+\.\s*/, "").trim(); // hapus prefix angka jika ada

        if (namaPenerbitYangAkanDihapus[cleanName]) {
          Logger.log("Menghapus sheet: " + cleanName);
          ss.deleteSheet(sheet);
        }
      });
    }

    function ModulHapusBarisSaja(jenisYangDihapus) {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var sheet = ss.getSheetByName("Hasil Seleksi");
      var data = sheet.getDataRange().getValues();
      var barisUntukDihapus = [];

      for (var i = 0; i < data.length; i++) {
        var namaPenerbit = data[i][1]; // Kolom B
        var jenisPPN = data[i][3];     // Kolom D

        if (
          namaPenerbit &&
          jenisPPN &&
          jenisPPN.toLowerCase().trim() === jenisYangDihapus.toLowerCase().trim()
        ) {
          barisUntukDihapus.push(i + 1); // +1 karena indeks baris mulai dari 1
        }
      }

      // Hapus dari baris terbawah agar indeks tidak bergeser
      barisUntukDihapus.reverse().forEach(function (rowIndex) {
        sheet.deleteRow(rowIndex);
      });

      Logger.log("Selesai Menghapus Daftar Sheet Di Sheet Hasil Seleksi");
    }
  //     ========     Fungsi Utama MengHapus Sheet Penerbit PPN atau Non PPN     ========  