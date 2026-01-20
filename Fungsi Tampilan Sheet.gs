//     ========     Atur Tampilan 1 Sheet    ========
function AturTampilan1Sheet(sheet) {
  modulFungsiTampilanSheetPenerbit(sheet);
}
//     ========     Atur Tampilan 1 Sheet    ========

//     ========     Atur Tampilan Semua Sheet    ========
function AturTampilanSemuaSheet() {
  //PERLU DIRUBAH ==============================================================================
  const sheetMulai = 0;
  //PERLU DIRUBAH ==============================================================================
  modulFungsiTampilanSheetPenerbitALLSHEET(sheetMulai);
}
//     ========     Atur Tampilan Semua Sheet    ========

//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit    ========
function modulFungsiTampilanSheetPenerbit(sheet) {
  // ✅ Memeriksa Sheet Yang Aktif
  if (!sheet) {
    try {
      sheet = SpreadsheetApp.getActiveSheet();
    } catch (e) {
      Logger.log("❌ Tidak ada sheet aktif. Fungsi dihentikan.");
      return;
    }
  }
  if (!sheet) return;
  //

  // ✅ Variabel yang Banyak Di Gunakan
  const nama = sheet.getSheetName();
  const startRow = 10;
  const spreadsheet = sheet.getParent();
  const lastRow = sheet.getLastRow();
  const totalRows = sheet.getMaxRows();
  Logger.log("▶ Menjalankan FungsiTampilanSheetPenerbit pada: " + nama);

  //

  // ✅ Melewati Sheet Form Pengadaan , Hasil Seleksi , Referensi , DaftarISBN , DaftarUUID
  if (SHEET_KECUALI.includes(nama)) {
    Logger.log("Sheet dilewati: " + nama);
    return;
  }
  //

  // ✅ Melepas Freeze & Filter jika ada
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);
  const filter = sheet.getFilter();
  if (filter) filter.remove();
  //

  // ✅ Menghapusformat seluruh area data dulu
  if (lastRow >= startRow) {
    const dataRange = sheet.getRange(`A10:AC${lastRow}`);
    dataRange.clearFormat(); // hapus format
  }
  //

  //✅ Mengatur ukuran kolom
  const ukuranKolom = [
    44, 119, 369, 129, 127, 134, 124, 125, 109, 109, 109, 100, 100, 100, 100,
    100, 100, 100, 100, 100, 100, 100, 100, 100, 100, 150, 80, 115, 266,
  ];

  ukuranKolom.forEach((width, i) => {
    sheet.setColumnWidth(i + 1, width);
  });
  //

  // ✅ Mengisi Rumus Tabel Ketersediaan Katalog dan Hasil Seleksi
  const formulaCells = [
    ["G2", `=COUNTA(C10:C${lastRow})`],
    ["G3", `=SUM(Z10:Z${lastRow})`],
    ["G4", `=AVERAGE(Z10:Z${lastRow})`],
    ["J2", `=COUNTA(AA10:AA${lastRow})`],
    ["J3", `=SUM(AA10:AA${lastRow})`],
    ["J4", `=SUM(AB10:AB${lastRow})`],
    ["J5", `=AVERAGEIF(AA10:AA${lastRow}; ">0"; AB10:AB${lastRow})`],
  ];
  formulaCells.forEach(([cell, formula]) =>
    sheet.getRange(cell).setFormula(formula),
  );
  //

  // ✅ Mengisi autofill nomor Urut , Preview Konten ,Referensi,Sub Referensi ,Total Harga
  function modulclearAndAutoFillColumn(colLetter, formulaOrValue) {
    const col = sheet.getRange(colLetter + "1").getColumn();
    const range = sheet.getRange(startRow, col, lastRow - startRow + 1);
    range.clear({ contentsOnly: true, skipFilteredRows: true });

    const firstCell = sheet.getRange(startRow, col);
    if (formulaOrValue.startsWith("=")) {
      firstCell.setFormula(formulaOrValue);
    } else {
      firstCell.setValue(formulaOrValue);
    }
    if (lastRow > startRow) {
      firstCell.autoFill(
        sheet.getRange(startRow, col, lastRow - startRow + 1),
        SpreadsheetApp.AutoFillSeries.ALTERNATE_SERIES,
      );
    }
  }
  modulclearAndAutoFillColumn("A", "1");
  modulclearAndAutoFillColumn(
    "B",
    '=HYPERLINK("https://mocostore.moco.co.id/catalog/"&AC10;"Klik Disini")',
  );
  modulclearAndAutoFillColumn("AB", "=Z10*AA10");
  modulclearAndAutoFillColumn(
    "K",
    '=IFERROR(VLOOKUP(J10; Referensi!A:B; 2; FALSE); "")',
  );
  modulclearAndAutoFillColumn(
    "I",
    '=IFERROR(VLOOKUP(LEFT(J10;3); Referensi!A:B; 2; FALSE); "")',
  );
  //

  // ✅ Hapus baris kosong
  if (lastRow < totalRows) {
    sheet.deleteRows(lastRow + 1, totalRows - lastRow);
    Logger.log(`🗑 Menghapus ${totalRows - lastRow} baris kosong.`);
  }
  //

  // ✅ Mengatur Format border , Font , alignment
  const borderStyle = SpreadsheetApp.BorderStyle.SOLID;
  sheet.getRange("F1:J5").setBorder(false, false, false, false, false, false);
  sheet
    .getRangeList(["F1:G4", "I1:J5"])
    .setBorder(true, true, true, true, true, true, "#000", borderStyle);
  sheet
    .getRange(`A10:AC${lastRow}`)
    .setBorder(true, true, true, true, true, true, "#000", borderStyle);
  sheet
    .getRange("F1:J5")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");
  sheet
    .getRangeList([
      `A10:B${lastRow}`,
      `F10:G${lastRow}`,
      `J10:J${lastRow}`,
      `Y10:AC${lastRow}`,
    ])
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontStyle("normal")
    .setFontWeight("normal");
  sheet
    .getRangeList([`C10:E${lastRow}`, `H10:I${lastRow}`, `K10:X${lastRow}`])
    .setHorizontalAlignment("left")
    .setVerticalAlignment("middle")
    .setFontStyle("normal")
    .setFontWeight("normal");
  sheet.getRangeList(["J4", "G3", "G4"]).setNumberFormat("[$Rp-421] #,##0");
  //

  // ✅ Update alternating color range
  function updateAlternatingColor(sheet, lastRow) {
    sheet.getBandings().forEach((b) => b.remove());
    const banding = sheet
      .getRange(`A9:AC${lastRow}`)
      .applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

    // banding.setHeaderRowColor(null); // biar baris 9 tetap polos
  }
  updateAlternatingColor(sheet, lastRow);
  //

  // ✅ Pasang ulang freeze
  sheet.setFrozenColumns(10);
  sheet.setFrozenRows(9);
  //

  // ✅ Pasang filter
  const dataRange = sheet.getRange(`A9:AC${lastRow}`);
  if (!dataRange.getFilter()) dataRange.createFilter();
  //

  // ✅ Mengatur tingg Baris
  if (lastRow >= startRow) {
    sheet.setRowHeightsForced(startRow, lastRow - startRow + 1, 20);
  }
  //

  // ✅ Ganti named range
  const cleanNamerange = nama.replace(/[0-9().\-]/g, "").replace(/\s/g, "");
  spreadsheet.setNamedRange(cleanNamerange, sheet.getRange(`J10:J${lastRow}`));
  //

  // ✅ Ganti nama sheet
  const cleanNamesheet = nama.replace(/[0-9.]/g, "");
  const sheetIndex = sheet.getIndex() - 3;
  const newName = `${sheetIndex}.${cleanNamesheet}`;
  if (spreadsheet.getSheets().every((s) => s.getName() !== newName)) {
    sheet.setName(newName);
  }
  //

  Logger.log(
    "✅ Selesai Menjalankan FungsiTampilanSheetPenerbit pada: " + nama,
  );
}

//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit    ========

//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit semua Sheet     ========
function modulFungsiTampilanSheetPenerbitALLSHEET(sheetMulai) {
  const spreadsheet = SpreadsheetApp.getActive();
  const semuaSheet = spreadsheet.getSheets();

  const mulai = sheetMulai + 2;
  semuaSheet.slice(mulai).forEach((sheet) => {
    if (!SHEET_KECUALI.includes(sheet.getName())) {
      spreadsheet.setActiveSheet(sheet); // 🔹 pindah ke sheet yang diproses
      modulFungsiTampilanSheetPenerbit(sheet);
    }
  });
}
//     ========     Fungsi Utama Mengatur Tampilan Sheet Penerbit semua Sheet     ========
