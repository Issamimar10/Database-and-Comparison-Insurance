function doGet(e) {
  // Dukungan untuk mengambil base64 file Google Drive langsung tanpa kendala CORS
  if (e && e.parameter && e.parameter.action === "get_file_base64" && e.parameter.fileId) {
    try {
      const file = DriveApp.getFileById(e.parameter.fileId);
      const b64 = Utilities.base64Encode(file.getBlob().getBytes());
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        fileName: file.getName(),
        mimeType: file.getMimeType(),
        base64: b64
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService.createTextOutput("Comparison Maker Backend Active")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Server sibuk, coba lagi dalam beberapa detik" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Batas karakter per sel Google Sheets
    const MAX_CELL_CHARS = 50000;

    // Helper: Ambil daftar PROTECTED_COLS dari sheet Config secara dinamis
    function getProtectedCols() {
      const defaultCols = [
        "product name", "company name", "launch date", "tanggal rilis", "par/non par", "product type",
        "distribution channel", "currency",
        "death benefit", "accidental death benefit", "terminal illness benefit",
        "sum assured bonus", "cash payout bonus", "maturity benefit",
        "additional benefit",
        "entry age", "coverage period", "premium payment period (ppp)",
        "min premium", "min sum assured", "underwriting method", "available riders",
        "aqd", "akad", "alokasi kontribusi", "surplus underwriting"
      ];

      try {
        const configSheet = ss.getSheetByName("Config");
        if (!configSheet) return defaultCols;

        const data = configSheet.getDataRange().getValues();
        const customCols = [];
        for (let i = 1; i < data.length; i++) {
          const type = String(data[i][0] || "").trim().toLowerCase();
          const val = String(data[i][1] || "").trim().toLowerCase();
          if (type === "non-benefit" && val) {
            customCols.push(val);
          }
        }
        return customCols.length > 0 ? Array.from(new Set(defaultCols.concat(customCols))) : defaultCols;
      } catch (err) {
        return defaultCols;
      }
    }

    const PROTECTED_COLS = getProtectedCols();

    // Helper: Ambil atau Gunakan Folder Google Drive untuk Dokumen (Folder "IDCM Test")
    function getOrCreateRootDocFolder() {
      const configSheet = ss.getSheetByName("Config");
      let folderId = "";

      if (configSheet) {
        const cfgData = configSheet.getDataRange().getValues();
        for (let i = 1; i < cfgData.length; i++) {
          const key = String(cfgData[i][1] || "").trim().toLowerCase();
          if (key === "drive folder id") {
            folderId = String(cfgData[i][2] || "").trim();
            break;
          }
        }
      }

      let rootFolder = null;
      if (folderId) {
        try {
          rootFolder = DriveApp.getFolderById(folderId);
        } catch (e) {
          rootFolder = null;
        }
      }

      if (!rootFolder) {
        // 1. Cek apakah ada folder bernama "IDCM Test" yang sudah dibuat user
        const userFolders = DriveApp.getFoldersByName("IDCM Test");
        if (userFolders.hasNext()) {
          rootFolder = userFolders.next();
        } else {
          // 2. Cek apakah ada folder "Comparison Maker Documents"
          const fallbackFolders = DriveApp.getFoldersByName("Comparison Maker Documents");
          if (fallbackFolders.hasNext()) {
            rootFolder = fallbackFolders.next();
          } else {
            // 3. Buat folder baru
            rootFolder = DriveApp.createFolder("IDCM Test");
          }
        }

        // Coba aktifkan sharing (jika diizinkan oleh akun Google, abaikan jika dibatasi kebijakan domain)
        try {
          rootFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (e) {}

        // Catat otomatis ID folder ini ke sheet Config
        if (configSheet) {
          const cfgData = configSheet.getDataRange().getValues();
          let written = false;
          for (let i = 1; i < cfgData.length; i++) {
            const key = String(cfgData[i][1] || "").trim().toLowerCase();
            if (key === "drive folder id") {
              configSheet.getRange(i + 1, 3).setValue(rootFolder.getId());
              written = true;
              break;
            }
          }
          if (!written) {
            configSheet.appendRow(["Setting", "Drive Folder ID", rootFolder.getId()]);
          }
        }
      }

      return rootFolder;
    }

    // Helper: Pastikan semua header dari web ada di Sheet
    function syncSheetHeaders(targetSheet, webHeaders) {
      let lastCol = targetSheet.getLastColumn();
      let sheetHeaders = lastCol > 0
        ? targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
        : [];

      if (!webHeaders || !Array.isArray(webHeaders)) return sheetHeaders;

      // Jika sheet masih kosong total -> tulis seluruh header dari web di baris 1
      if (sheetHeaders.length === 0 || sheetHeaders.every(h => h === "")) {
        const cleanHeaders = webHeaders.map(h => String(h || "").trim()).filter(h => h !== "");
        if (cleanHeaders.length > 0) {
          const headerRange = targetSheet.getRange(1, 1, 1, cleanHeaders.length);
          headerRange.setNumberFormat("@");
          headerRange.setValues([cleanHeaders]);
        }
        return cleanHeaders;
      }

      webHeaders.forEach(function(webHeader) {
        const cleanWebHeader = String(webHeader || "").trim();
        if (!cleanWebHeader) return;

        const exists = sheetHeaders.some(sh => sh.toLowerCase() === cleanWebHeader.toLowerCase());
        if (!exists) {
          const lowerH = cleanWebHeader.toLowerCase();

          if (lowerH === "launch date" || lowerH === "tanggal rilis") {
            // Sisipkan Launch Date setelah Currency
            const currIdx = sheetHeaders.findIndex(sh => sh.toLowerCase() === "currency");
            if (currIdx !== -1) {
              targetSheet.insertColumnAfter(currIdx + 1);
              targetSheet.getRange(1, currIdx + 2).setValue("Launch Date");
            } else {
              const insertAfter = Math.min(6, sheetHeaders.length);
              targetSheet.insertColumnAfter(insertAfter);
              targetSheet.getRange(1, insertAfter + 1).setValue("Launch Date");
            }
          } else if (lowerH === "additional benefit") {
            // Sisipkan "Additional Benefit" tepat sebelum "Entry Age"
            const entryAgeIdx = sheetHeaders.findIndex(sh => sh.toLowerCase() === "entry age");
            if (entryAgeIdx !== -1) {
              targetSheet.insertColumnBefore(entryAgeIdx + 1);
              targetSheet.getRange(1, entryAgeIdx + 1).setValue("Additional Benefit");
            } else {
              const insertAfter = Math.max(1, Math.min(12, sheetHeaders.length));
              targetSheet.insertColumnAfter(insertAfter);
              targetSheet.getRange(1, insertAfter + 1).setValue("Additional Benefit");
            }
          } else {
            // Kolom baru lainnya disisipkan setelah kolom ke-11
            const insertAfterCol = Math.max(1, Math.min(11, sheetHeaders.length));
            targetSheet.insertColumnAfter(insertAfterCol);
            targetSheet.getRange(1, insertAfterCol + 1).setValue(cleanWebHeader);
          }

          lastCol = targetSheet.getLastColumn();
          sheetHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
        }
      });

      return sheetHeaders;
    }

    // Helper: Hapus kolom manfaat tambahan (di luar PROTECTED_COLS) yang kosong di seluruh baris
    function cleanupEmptyBenefitColumns(targetSheet) {
      const updatedData = targetSheet.getDataRange().getValues();
      if (updatedData.length <= 1) return;

      const latestHeaders = updatedData[0].map(h => String(h).trim());
      for (let colIdx = latestHeaders.length - 1; colIdx >= 0; colIdx--) {
        const colNameLower = latestHeaders[colIdx].toLowerCase();
        if (!colNameLower || PROTECTED_COLS.indexOf(colNameLower) !== -1) continue;

        let isAllEmpty = true;
        for (let r = 1; r < updatedData.length; r++) {
          const cellVal = String(updatedData[r][colIdx] || "").trim();
          if (cellVal !== "" && cellVal !== "-" && cellVal.toLowerCase() !== "n/a") {
            isAllEmpty = false;
            break;
          }
        }

        if (isAllEmpty) {
          targetSheet.deleteColumn(colIdx + 1);
        }
      }
    }

    // Helper: Tulis 1 baris sebagai TEKS MURNI (@)
    function writeRowAsPlainText(targetSheet, rowIndex, rowValues, headers) {
      const cleanValues = rowValues.map(function(v) {
        return (v === null || v === undefined) ? "-" : v;
      });

      cleanValues.forEach(function(v, i) {
        if (String(v).length > MAX_CELL_CHARS) {
          throw new Error("Kolom '" + (headers[i] || ("#" + (i + 1))) + "' melebihi " + MAX_CELL_CHARS + " karakter");
        }
      });

      const range = targetSheet.getRange(rowIndex, 1, 1, cleanValues.length);
      range.setNumberFormat("@");
      range.setValues([cleanValues]);
    }

    // Helper: Susun baris sesuai urutan header di sheet
    function alignRowToSheetHeaders(webHeaders, rowData, sheetHeaders) {
      const valueMap = {};
      webHeaders.forEach(function(h, idx) {
        valueMap[String(h).trim().toLowerCase()] = rowData[idx];
      });
      return sheetHeaders.map(function(sh) {
        const key = sh.toLowerCase();
        return (valueMap[key] !== undefined) ? valueMap[key] : "-";
      });
    }

    function jsonResponse(obj) {
      return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ====================================================
    // 1. ACTION: UPLOAD_DOCUMENT (Simpan ke Google Drive & Catat di Sheet 'Documents')
    // ====================================================
    if (payload.action === "upload_document") {
      const rootFolder = getOrCreateRootDocFolder();
      const cleanProdName = String(payload.productName || "General").trim();

      // Buat atau gunakan subfolder per produk
      const prodFolders = rootFolder.getFoldersByName(cleanProdName);
      const prodFolder = prodFolders.hasNext() ? prodFolders.next() : rootFolder.createFolder(cleanProdName);
      try {
        prodFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}

      // Decode base64 dan simpan file ke Drive
      const decodedBytes = Utilities.base64Decode(payload.base64Data);
      const cleanFileName = String(payload.fileName || "dokumen").trim();
      const blob = Utilities.newBlob(decodedBytes, payload.mimeType || "application/octet-stream", cleanFileName);
      const file = prodFolder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}

      const fileId = file.getId();
      const viewUrl = file.getUrl();
      const downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
      const fileSize = file.getSize();
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");

      // Catat ke sheet Documents
      let docSheet = ss.getSheetByName("Documents");
      if (!docSheet) {
        docSheet = ss.insertSheet("Documents");
        const docHeaders = ["Product Name", "Doc Type", "File Name", "File ID", "Drive URL", "Download URL", "File Size", "Upload Date"];
        const hRange = docSheet.getRange(1, 1, 1, docHeaders.length);
        hRange.setNumberFormat("@");
        hRange.setValues([docHeaders]);
      }

      const docRow = [
        cleanProdName,
        payload.docType || "Dokumen",
        cleanFileName,
        fileId,
        viewUrl,
        downloadUrl,
        String(fileSize),
        nowStr
      ];
      writeRowAsPlainText(docSheet, docSheet.getLastRow() + 1, docRow, [
        "Product Name", "Doc Type", "File Name", "File ID", "Drive URL", "Download URL", "File Size", "Upload Date"
      ]);

      return jsonResponse({
        status: "success",
        action: "upload_document",
        fileId: fileId,
        viewUrl: viewUrl,
        downloadUrl: downloadUrl,
        fileName: cleanFileName,
        fileSize: fileSize,
        productName: cleanProdName,
        docType: payload.docType || "Dokumen"
      });
    }

    // ====================================================
    // 2. ACTION: DELETE_DOCUMENT (Hapus dari Drive & Baris di Sheet 'Documents')
    // ====================================================
    if (payload.action === "delete_document") {
      if (payload.fileId) {
        try {
          const file = DriveApp.getFileById(payload.fileId);
          file.setTrashed(true);
        } catch (e) {
          console.error("Gagal trash file di Drive:", e);
        }
      }

      const docSheet = ss.getSheetByName("Documents");
      if (docSheet) {
        const data = docSheet.getDataRange().getValues();
        for (let r = data.length - 1; r >= 1; r--) {
          const rowFileId = String(data[r][3] || "").trim();
          if (rowFileId === String(payload.fileId).trim()) {
            docSheet.deleteRow(r + 1);
          }
        }
      }

      return jsonResponse({ status: "success", action: "delete_document" });
    }

    // ====================================================
    // 3. ACTION: GET_FILE_BASE64 (Untuk Ekspor ZIP Tanpa Kendala CORS)
    // ====================================================
    if (payload.action === "get_file_base64" && payload.fileId) {
      const file = DriveApp.getFileById(payload.fileId);
      const b64 = Utilities.base64Encode(file.getBlob().getBytes());
      return jsonResponse({
        status: "success",
        fileName: file.getName(),
        mimeType: file.getMimeType(),
        base64: b64
      });
    }

    // ====================================================
    // 4. ACTION: DELETE (Hapus 1 atau Banyak Produk)
    // ====================================================
    if (payload.action === "delete" && Array.isArray(payload.items)) {
      const affectedSheets = {};

      payload.items.forEach(function(item) {
        const sheet = ss.getSheetByName(item.sheetName);
        if (!sheet) return;

        const dataRange = sheet.getDataRange().getValues();
        const targetName = String(item.productName || "").trim().toLowerCase();

        for (let r = dataRange.length - 1; r >= 1; r--) {
          const rowProdName = String(dataRange[r][0] || "").trim().toLowerCase();
          if (rowProdName === targetName) {
            sheet.deleteRow(r + 1);
            affectedSheets[item.sheetName] = sheet;
          }
        }
      });

      Object.keys(affectedSheets).forEach(function(sName) {
        cleanupEmptyBenefitColumns(affectedSheets[sName]);
      });

      return jsonResponse({ status: "success", action: "delete" });
    }

    let sheet = ss.getSheetByName(payload.sheetName);
    if (!sheet && payload.sheetName === "Documents") {
      sheet = ss.insertSheet("Documents");
      const docHeaders = ["Product Name", "Doc Type", "File Name", "File ID", "Drive URL", "Download URL", "File Size", "Upload Date"];
      const hRange = sheet.getRange(1, 1, 1, docHeaders.length);
      hRange.setNumberFormat("@");
      hRange.setValues([docHeaders]);
    }
    if (!sheet) {
      return jsonResponse({ status: "error", message: "Sheet tidak ditemukan: " + payload.sheetName });
    }

    // ====================================================
    // 5. ACTION: APPEND (Upload Batch dari File Excel)
    // ====================================================
    if (payload.action === "append") {
      if (payload.headers && Array.isArray(payload.headers)) {
        const sheetHeaders = syncSheetHeaders(sheet, payload.headers);

        payload.data.forEach(function(row) {
          const alignedRow = alignRowToSheetHeaders(payload.headers, row, sheetHeaders);
          writeRowAsPlainText(sheet, sheet.getLastRow() + 1, alignedRow, sheetHeaders);
        });
      } else {
        payload.data.forEach(function(row) {
          writeRowAsPlainText(sheet, sheet.getLastRow() + 1, row, payload.data[0] || []);
        });
      }

      return jsonResponse({ status: "success", action: "append" });
    }

    // ====================================================
    // 6. ACTION: UPDATE / CREATE MANUAL
    // ====================================================
    if (payload.action === "update") {
      const webHeaders = (payload.headers && Array.isArray(payload.headers)) ? payload.headers : [];
      const sheetHeaders = syncSheetHeaders(sheet, webHeaders);

      const alignedRow = alignRowToSheetHeaders(webHeaders, payload.data, sheetHeaders);

      const allValues = sheet.getDataRange().getValues();
      let rowIndex = -1;
      for (let r = 1; r < allValues.length; r++) {
        if (String(allValues[r][0]).trim().toLowerCase() === String(payload.productName).trim().toLowerCase()) {
          rowIndex = r + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        writeRowAsPlainText(sheet, rowIndex, alignedRow, sheetHeaders);
      } else {
        writeRowAsPlainText(sheet, sheet.getLastRow() + 1, alignedRow, sheetHeaders);
      }

      cleanupEmptyBenefitColumns(sheet);

      return jsonResponse({ status: "success", action: "update" });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali" });

  } catch (err) {
    console.error("doPost error:", err);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

// ====================================================
// FUNGSI KHUSUS UNTUK MEMICU POP-UP IZIN TULIS (WRITE) GOOGLE DRIVE
// ====================================================
function authorizeDrive() {
  // Menulis file uji coba sementara agar Google memicu pop-up izin penuh (Create, Edit, Delete)
  const tempFile = DriveApp.createFile("test_permission.txt", "OK");
  tempFile.setTrashed(true);
  console.log("Izin Google Drive Full Access (Tulis & Buat Folder) Berhasil Aktif!");
}
