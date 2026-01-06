//     ========     Fungsi Mengisi Sheet Hasil Seleksi     ========
 function FungsiIsiSheetHasilSeleksi() {

  // PERLU DIRUBAH ==============================================================================
    const blokAwal = [20, 36];// baris awal untuk tiap tabel
  // PERLU DIRUBAH ==============================================================================
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetUtama = ss.getActiveSheet();
  const semuaSheet = ss.getSheets();
  const namaSheets = semuaSheet.map(s => s.getName());
  const colPenerbit = 2; // kolom B
  const colJumlahJudul = 3; // kolom C
  const colTotalHarga = 5; // kolom E
  const coljumjudlsel = 9; // kolom I
  const coltotcopsel = 10; // kolom J
  const coltotharsel = 11; // kolom K

  // ✅ Cek Sheet Yang Aktif
    if (sheetUtama.getName() !== 'Hasil Seleksi') {
      Logger.log('Fungsi ini hanya dapat dijalankan di sheet "Hasil Seleksi".');
      return;
    }
  // 

  // ✅ Proses lop setiap blok
    blokAwal.forEach(startRow => {
      let row = startRow;
      while (true) {
        const penerbit = sheetUtama.getRange(row, colPenerbit).getValue();
          if (!penerbit || penerbit.toString().trim() === "") break; // berhenti kalau kosong  
        const sheetCocok = namaSheets.find(namaSheet =>namaSheet.toLowerCase().includes(penerbit.toString().toLowerCase()));
         if (sheetCocok) {
          Logger.log(`✅ Mengisi data untuk: ${penerbit} (sheet: ${sheetCocok})`);
          sheetUtama.getRange(row, colJumlahJudul).setFormula(`='${sheetCocok}'!G2`);
          sheetUtama.getRange(row, colTotalHarga).setFormula(`='${sheetCocok}'!G3`);
          sheetUtama.getRange(row, coljumjudlsel).setFormula(`='${sheetCocok}'!J2`);
          sheetUtama.getRange(row, coltotcopsel).setFormula(`='${sheetCocok}'!J3`);
          sheetUtama.getRange(row, coltotharsel).setFormula(`='${sheetCocok}'!J4`);
         } else {
          Logger.log(`⚠️ Sheet untuk penerbit "${penerbit}" tidak ditemukan!`);
          sheetUtama.getRange(row, colJumlahJudul).setValue('Sheet tidak ditemukan');
          sheetUtama.getRange(row, colTotalHarga).setValue('Sheet tidak ditemukan');
          sheetUtama.getRange(row, coljumjudlsel).setValue('Sheet tidak ditemukan');
          sheetUtama.getRange(row, coltotcopsel).setValue('Sheet tidak ditemukan');
          sheetUtama.getRange(row, coltotharsel).setValue('Sheet tidak ditemukan');
         } row++;}});
 }
      
 function AddSheetLinks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Hasil Seleksi");
  const column = 2; // kolom B

  // daftar blok baris yang mau diproses
  const rowRanges = [
  { start: 20, end: 28 },   // baris 20 sampai 28
  { start: 36, end: 119 }, // baris 36 sampai 119
  ];

  // ambil semua nama sheet di file
  const allSheetNames = ss.getSheets().map(s => s.getName());

  rowRanges.forEach(rangeInfo => {
    const startRow = rangeInfo.start;
    const endRow = rangeInfo.end;
    const numRows = endRow - startRow + 1;

    const range = sheet.getRange(startRow, column, numRows, 1);
    const values = range.getValues();

    for (let i = 0; i < values.length; i++) {
      const publisherName = values[i][0];
      if (!publisherName) continue;

      // cari sheet yang namanya mengandung teks publisherName
      const matchingSheetName = allSheetNames.find(name => name.includes(publisherName));
      if (!matchingSheetName) continue;

      const targetSheet = ss.getSheetByName(matchingSheetName);
      const link = `https://docs.google.com/spreadsheets/d/${ss.getId()}/edit#gid=${targetSheet.getSheetId()}`;

      // bikin rich text link tanpa underline
      const style = SpreadsheetApp.newTextStyle()
        .setUnderline(false)
        .build();

      const richText = SpreadsheetApp.newRichTextValue()
        .setText(publisherName)
        .setTextStyle(style)
        .setLinkUrl(link)
        .build();

      range.getCell(i + 1, 1).setRichTextValue(richText);
    }
  });

  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Link otomatis berhasil dibuat tanpa underline!");
 }
//     ========     Fungsi Mengisi Sheet Hasil Seleksi     ========