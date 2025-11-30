//     ========     Fungsi Mengisi Sheet Referensi     ========
  function FungsiIsiSheetReferensi() {
    // ✅ Cek Sheet Aktif
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetAktif = ss.getActiveSheet();
      if (sheetAktif.getName().toUpperCase() !== 'REFERENSI') {
        Logger.log('Fungsi ini hanya bisa dijalankan di sheet "REFERENSI".');
        return;
      }
    // 
    Logger.log('Menjalankan Fungsi Isi Sheet Referensi');
    const sheetHasilSeleksi = ss.getSheetByName("Hasil Seleksi");
    const sheetRef = ss.getSheetByName("REFERENSI");
    const semuaSheet = ss.getSheets();
    const startCol = 5; // kolom E
    const lastCol = sheetRef.getLastColumn();
    const rowHyperlink = 4; // baris hyperlink
    let lastHyperlinkCol = startCol;
    for (let c = startCol; c <= lastCol; c++) {
    const formula = sheetRef.getRange(rowHyperlink, c).getFormula();
    if (formula && formula.includes("HYPERLINK")) {
    lastHyperlinkCol = c;}} // update posisi terakhir
    const rowAwalKode = 11; // baris awal kode referensi
    
    // Range pertama: E4 sampai kolom terakhir di baris 119
    const range1 = sheetRef.getRange(4, 5, 129, lastCol - 4); 
    range1.clearContent();

    // Range kedua: C11:C129
    const range2 = sheetRef.getRange("C11:C129");
    range2.clearContent();

  
    // ✅ 1. Modul Hyperlink Penerbit
      Logger.log('Mengisi Data Penerbit');
      const lastRowHS = sheetHasilSeleksi.getLastRow();
      const dataNo = sheetHasilSeleksi.getRange(1, 1, lastRowHS).getValues(); // kol A
      const dataPenerbit = sheetHasilSeleksi.getRange(1, 2, lastRowHS).getValues(); // kol B
      let col = startCol;

      for (let i = 0; i < lastRowHS; i++) {
        const nomor = dataNo[i][0];
        const penerbit = dataPenerbit[i][0];
        if (typeof nomor === "number" && !isNaN(nomor) && penerbit) {
          const sheetCocok = semuaSheet.find(s =>
            s.getName().toLowerCase().includes(penerbit.toString().toLowerCase())
          );
          if (sheetCocok) {
            const gid = sheetCocok.getSheetId();
            const label = `${nomor}`;
            const formula = `=HYPERLINK("#gid=${gid}"; "${label}")`;
            sheetRef.getRange(rowHyperlink, col).setFormula(formula);
            col++;
          }
        }    }
    //  

    // ✅ 2. Modul Isi Data Referensi
      Logger.log('Mengisi Data Referensi');
        const barisTarget = [];
        const dataKode = sheetRef.getRange(rowAwalKode, 1, sheetRef.getLastRow() - rowAwalKode + 1, 1).getValues();

        for (let i = 0; i < dataKode.length; i++) {
          const isi = dataKode[i][0];
          if (isi && typeof isi === "string" && !isi.toLowerCase().includes("total")) {
            barisTarget.push(rowAwalKode + i);   }    
          }
        function colLetter(colIndex) {
          let letter = '';
          while (colIndex > 0) {
            let mod = (colIndex - 1) % 26;
            letter = String.fromCharCode(65 + mod) + letter;
            colIndex = Math.floor((colIndex - mod) / 26);
          }
          return letter;
        }
        const sumFormulas = {
          5: '=SUM({col}14;{col}25;{col}38;{col}50;{col}57;{col}64;{col}72;{col}84;{col}118;{col}129)',
          14: '=SUM({col}11:{col}13)',
          25: '=SUM({col}19:{col}24)',
          38: '=SUM({col}30:{col}37)',
          50: '=SUM({col}43:{col}49)',
          57: '=SUM({col}55:{col}56)',
          64: '=SUM({col}62:{col}63)',
          72: '=SUM({col}69:{col}71)',
          84: '=SUM({col}77:{col}83)',
          118: '=SUM({col}89:{col}117)',
          129: '=SUM({col}123:{col}128)'
        };

        // ✅ 3. Isi kolom C 
      const lastColLetter = colLetter(lastCol);
      barisTarget.forEach(row => {
      sheetRef.getRange(row, 3).setFormula(`=SUM(E${row}:${lastColLetter}${row})`);
      });
    //  

        // ✅  4. Isi total di kolom C (C14, C25, dst) sesuai sumFormulas
      for (let rowStr in sumFormulas) {
        const rowNum = parseInt(rowStr, 10);
        const formula = sumFormulas[rowNum].replace(/{col}/g, 'C');
        sheetRef.getRange(rowNum, 3).setFormula(formula);
      }
    // 
        for (let col = startCol; col <= lastHyperlinkCol; col++) {
          const formulaHyperlink = sheetRef.getRange(rowHyperlink, col).getFormula();
          const gidMatch = formulaHyperlink.match(/gid=(\d+)/);
          if (!gidMatch){
            Logger.log(`⏭ Kolom ${col} dilewati (tidak ada hyperlink)`); 
            continue;
          }
          const gid = parseInt(gidMatch[1], 10);
          const sheetTarget = semuaSheet.find(s => s.getSheetId() === gid);
          if (!sheetTarget){
            Logger.log(`❌ Kolom ${col} - Sheet target dengan gid ${gid} tidak ditemukan`);
            continue;
          }
          const namaSheet = sheetTarget.getName();
          const namaRange = namaSheet.replace(/[^A-Za-z]/g, "");
          const colLet = colLetter(col);
          Logger.log(`✅ Proses kolom ${colLet} (${col}) → Sheet: ${namaSheet}`);

          // ✅ Baris 6: Baris yang belum teridentifikasi
          sheetRef.getRange(6, col).setFormula(`=COUNTIF(${namaRange};"")`);

          // ✅ Baris yang sudah teridentifikasi
          barisTarget.forEach(row => {
            sheetRef.getRange(row, col).setFormula(`=COUNTIF(${namaRange};$A${row})`);
          });

          // ✅ Baris SUM
          for (let rowStr in sumFormulas) {
            const rowNum = parseInt(rowStr, 10);
            const formula = sumFormulas[rowNum].replace(/{col}/g, colLet);
            sheetRef.getRange(rowNum, col).setFormula(formula);
          }
        }
    //
    Logger.log('Selesai Menjalankan Fungsi Isi Sheet Referensi');
 } 
//     ========     Fungsi Mengisi Sheet Referensi     ========