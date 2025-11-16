//buat sheet baru "CariUUID"
//struktur tabel No	Judul	Pengarang	Penerbit	Perusahaan	E-ISBN	UUID

function isiUUID_dariSemuaSheet() {
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
