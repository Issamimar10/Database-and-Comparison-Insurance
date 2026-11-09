// ====================================================
// COMPARISON MAKER - GOOGLE APPS SCRIPT BACKEND
// Versi: 2.1 (Penyimpanan Dokumen Langsung ke Google Drive Tanpa Sheet 'Documents')
// ====================================================

// Helper Global: Ambil atau Gunakan Folder Google Drive untuk Dokumen (Folder "IDCM Test")
function getOrCreateRootDocFolder() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
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
  }
  return rootFolder;
}

function doGet(e) {
  // 1. Ambil daftar seluruh dokumen langsung dari folder Google Drive (Tanpa Sheet 'Documents')
  if (e && e.parameter && e.parameter.action === "get_all_documents") {
    try {
      const rootFolder = getOrCreateRootDocFolder();
      const docs = [];
      const subFolders = rootFolder.getFolders();
      while (subFolders.hasNext()) {
        const sub = subFolders.next();
        const prodName = sub.getName();
        if (prodName === "__TABLE_IMAGES__") continue;
        const files = sub.getFiles();
        while (files.hasNext()) {
          const f = files.next();
          if (f.isTrashed()) continue;
          if (f.getDescription() === "__TABLE_IMAGE__") continue;
          docs.push({
            productName: prodName,
            docType: f.getDescription() || "Dokumen",
            fileName: f.getName(),
            fileId: f.getId(),
            driveUrl: f.getUrl(),
            downloadUrl: "https://drive.google.com/uc?export=download&id=" + f.getId(),
            fileSize: f.getSize(),
            uploadDate: Utilities.formatDate(f.getDateCreated(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss")
          });
        }
      }
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        documents: docs
      })).setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // 2. Dukungan untuk mengambil base64 file Google Drive langsung tanpa kendala CORS
  // Mendukung konversi otomatis file Office (.docx / .xlsx) ke PDF untuk pratinjau instan di web
  if (e && e.parameter && e.parameter.action === "get_file_base64" && e.parameter.fileId) {
    try {
      const file = DriveApp.getFileById(e.parameter.fileId);
      const asPdf = e.parameter.asPdf === "true" || e.parameter.asPdf === true;
      let blob;
      let mimeType;
      
      if (asPdf) {
        try {
          blob = file.getAs("application/pdf");
          mimeType = "application/pdf";
        } catch (convErr) {
          blob = file.getBlob();
          mimeType = file.getMimeType();
        }
      } else {
        blob = file.getBlob();
        mimeType = file.getMimeType();
      }

      const b64 = Utilities.base64Encode(blob.getBytes());
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        fileName: file.getName(),
        mimeType: mimeType,
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
    function getProtectedCols(sheetName) {
      const isSyariah = String(sheetName || "").trim().toLowerCase() === "syariah";
      const defaultCols = [
        "product name", "company name", "launch date", "tanggal rilis", "par/non par", "product type",
        "distribution channel", "currency",
        "death benefit", "accidental death benefit", "terminal illness benefit",
        "sum assured bonus", "cash payout bonus", "maturity benefit",
        "additional benefit",
        "entry age", "coverage period", "premium payment period (ppp)",
        "min premium", "min sum assured", "underwriting method", "available riders"
      ].concat(isSyariah ? ["aqd", "akad", "alokasi kontribusi", "surplus underwriting"] : []);

      try {
        const configSheet = ss.getSheetByName("Config");
        if (!configSheet) return defaultCols;

        const data = configSheet.getDataRange().getValues();
        const customCols = [];
        for (let i = 1; i < data.length; i++) {
          const type = String(data[i][0] || "").trim().toLowerCase();
          const val = String(data[i][1] || "").trim().toLowerCase();
          if (type === "non-benefit" && val) {
            if (!isSyariah && ["aqd", "akad", "alokasi kontribusi", "surplus underwriting"].includes(val)) {
              continue;
            }
            customCols.push(val);
          }
        }
        return customCols.length > 0 ? Array.from(new Set(defaultCols.concat(customCols))) : defaultCols;
      } catch (err) {
        return defaultCols;
      }
    }

    // Helper: Sinkronisasi Header Kolom Sheet dengan Header Baru dari Web / Excel
    function syncSheetHeaders(targetSheet, incomingHeaders) {
      let sheetHeaders = [];
      const lastCol = targetSheet.getLastColumn();
      if (lastCol > 0) {
        sheetHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
      }

      if (sheetHeaders.length === 0 || sheetHeaders.every(h => h === "")) {
        targetSheet.getRange(1, 1, 1, incomingHeaders.length).setValues([incomingHeaders]);
        return incomingHeaders;
      }

      const lowerSheetHeaders = sheetHeaders.map(h => h.toLowerCase());
      const newColsToAdd = [];

      incomingHeaders.forEach(function(h) {
        const cleanH = String(h || "").trim();
        if (!cleanH) return;
        if (!lowerSheetHeaders.includes(cleanH.toLowerCase())) {
          newColsToAdd.push(cleanH);
          lowerSheetHeaders.push(cleanH.toLowerCase());
        }
      });

      if (newColsToAdd.length > 0) {
        const startCol = sheetHeaders.length + 1;
        targetSheet.getRange(1, startCol, 1, newColsToAdd.length).setValues([newColsToAdd]);
        sheetHeaders = sheetHeaders.concat(newColsToAdd);
      }

      return sheetHeaders;
    }

    // Helper: Petakan data array berdasarkan urutan nama kolom pada sheet
    function alignRowToSheetHeaders(incomingHeaders, incomingRow, sheetHeaders) {
      const lowerIncoming = incomingHeaders.map(h => String(h || "").trim().toLowerCase());
      return sheetHeaders.map(function(sHeader) {
        const sLower = String(sHeader || "").trim().toLowerCase();
        const idx = lowerIncoming.indexOf(sLower);
        if (idx !== -1 && incomingRow[idx] !== undefined && incomingRow[idx] !== null) {
          let cellStr = String(incomingRow[idx]).trim();
          if (cellStr.length > MAX_CELL_CHARS) {
            cellStr = cellStr.substring(0, MAX_CELL_CHARS - 15) + "...[TRUNCATED]";
          }
          return cellStr;
        }
        return "-";
      });
    }

    // Helper: Tulis satu baris sebagai Plain Text murni (format '@')
    function writeRowAsPlainText(targetSheet, rowNum, rowData, headers) {
      if (!rowData || rowData.length === 0) return;
      const range = targetSheet.getRange(rowNum, 1, 1, rowData.length);
      const formats = [];
      const formatRow = [];
      for (let c = 0; c < rowData.length; c++) {
        formatRow.push("@");
      }
      formats.push(formatRow);
      range.setNumberFormats(formats);

      const strRowData = rowData.map(function(val) {
        if (val === undefined || val === null) return "-";
        return String(val);
      });
      range.setValues([strRowData]);
    }

    // Helper: Pembersihan Kolom Manfaat Kosong
    function cleanupEmptyBenefitColumns(targetSheet) {
      try {
        if (String(targetSheet.getName() || "").trim().toLowerCase() === "config") return;
        const lastCol = targetSheet.getLastColumn();
        const lastRow = targetSheet.getLastRow();
        if (lastCol <= 1 || lastRow <= 1) return;

        const protectedCols = getProtectedCols(targetSheet.getName());
        const headers = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const data = targetSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

        for (let c = lastCol - 1; c >= 0; c--) {
          const hName = String(headers[c] || "").trim().toLowerCase();
          if (!hName || protectedCols.includes(hName)) {
            continue;
          }

          let hasValue = false;
          for (let r = 0; r < data.length; r++) {
            const cellVal = String(data[r][c] || "").trim();
            if (cellVal !== "" && cellVal !== "-") {
              hasValue = true;
              break;
            }
          }

          if (!hasValue) {
            targetSheet.deleteColumn(c + 1);
          }
        }
      } catch (err) {
        console.error("Gagal membersihkan kolom kosong:", err);
      }
    }

    function jsonResponse(obj) {
      return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // ====================================================
    // 1. ACTION: UPLOAD_DOCUMENT (Simpan ke Google Drive langsung tanpa Sheet 'Documents')
    // ====================================================
    if (payload.action === "upload_document") {
      const rootFolder = getOrCreateRootDocFolder();
      const cleanProdName = String(payload.productName || "General").trim();

      // Buat atau gunakan subfolder per produk di Google Drive
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

      // Simpan tipe dokumen ke metadata description file di Drive
      if (payload.docType) {
        file.setDescription(String(payload.docType));
      }

      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}

      const fileId = file.getId();
      const viewUrl = file.getUrl();
      const downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
      const fileSize = file.getSize();

      // Hapus tab sheet 'Documents' jika pernah ada di Master Spreadsheet
      const oldDocSheet = ss.getSheetByName("Documents");
      if (oldDocSheet) {
        try { ss.deleteSheet(oldDocSheet); } catch (e) {}
      }

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
    // 2. ACTION: DELETE_DOCUMENT (Hapus langsung dari Google Drive)
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

      // Hapus tab sheet 'Documents' jika masih ada di Master Spreadsheet
      const oldDocSheet = ss.getSheetByName("Documents");
      if (oldDocSheet) {
        try { ss.deleteSheet(oldDocSheet); } catch (e) {}
      }

      return jsonResponse({ status: "success", action: "delete_document" });
    }

    // ====================================================
    // 3. ACTION: GET_FILE_BASE64 (Untuk Ekspor ZIP & Preview Inline dengan Opsi Konversi PDF)
    // ====================================================
    if (payload.action === "get_file_base64" && payload.fileId) {
      const file = DriveApp.getFileById(payload.fileId);
      const asPdf = payload.asPdf === true || payload.asPdf === "true";
      let blob;
      let mimeType;

      if (asPdf) {
        try {
          blob = file.getAs("application/pdf");
          mimeType = "application/pdf";
        } catch (convErr) {
          blob = file.getBlob();
          mimeType = file.getMimeType();
        }
      } else {
        blob = file.getBlob();
        mimeType = file.getMimeType();
      }

      const b64 = Utilities.base64Encode(blob.getBytes());
      return jsonResponse({
        status: "success",
        fileName: file.getName(),
        mimeType: mimeType,
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
    if (!sheet && payload.sheetName) {
      sheet = ss.insertSheet(String(payload.sheetName).trim());
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

      cleanupEmptyBenefitColumns(sheet);

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
  const tempFile = DriveApp.createFile("test_permission.txt", "OK");
  tempFile.setTrashed(true);
  console.log("Izin Google Drive Full Access (Tulis & Buat Folder) Berhasil Aktif!");
}

// ====================================================
// FUNGSI UNTUK MENGHAPUS SHEET 'Documents' JIKA MASIH ADA DI MASTER SPREADSHEET
// ====================================================
function deleteDocumentsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const docSheet = ss.getSheetByName("Documents");
  if (docSheet) {
    ss.deleteSheet(docSheet);
    console.log("Sheet 'Documents' berhasil dihapus dari Master Spreadsheet!");
  } else {
    console.log("Sheet 'Documents' memang sudah tidak ada.");
  }
}
