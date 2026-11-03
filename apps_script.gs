function doPost(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Server sibuk, coba lagi" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const payload = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Batas karakter per sel Google Sheets
    const MAX_CELL_CHARS = 50000;

    // Daftar kolom standar yang TIDAK BOLEH dihapus otomatis meskipun kosong
    const PROTECTED_COLS = [
      "product name", "company name", "par/non par", "product type",
      "distribution channel", "currency",
      "death benefit", "accidental death benefit", "terminal illness benefit",
      "sum assured bonus", "cash payout bonus", "maturity benefit",
      "additional benefit",
      "entry age", "coverage period", "premium payment period (ppp)",
      "min premium", "min sum assured", "underwriting method", "available riders"
    ];

    // Helper: Pastikan semua header dari web (termasuk Additional Benefit di Col M) ada di Sheet
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
          if (cleanWebHeader.toLowerCase() === "additional benefit") {
            // Sisipkan "Additional Benefit" tepat sebelum "Entry Age" (di Column M)
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
            // Manfaat baru lainnya (termasuk kolom lanjutan seperti "Alokasi Kontribusi [2]")
            // disisipkan tepat setelah Column K (Kolom ke-11)
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

    // Helper: Tulis 1 baris sebagai TEKS MURNI
    // - Cek batas 50.000 karakter per sel (pesan error jelas menyebut nama kolom)
    // - Format "@" mencegah teks "- Aqd ..." jadi #ERROR! dan "5/10/15" berubah jadi tanggal
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

    // Helper: Susun baris sesuai urutan header di sheet berdasarkan nama header dari web
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
    // 1. ACTION: DELETE (Hapus 1 atau Banyak Produk Sekaligus)
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

    const sheet = ss.getSheetByName(payload.sheetName);
    if (!sheet) {
      return jsonResponse({ status: "error", message: "Sheet tidak ditemukan" });
    }

    // ====================================================
    // 2. ACTION: APPEND (Upload Batch dari File Excel)
    // ====================================================
    if (payload.action === "append") {
      if (payload.headers && Array.isArray(payload.headers)) {
        const sheetHeaders = syncSheetHeaders(sheet, payload.headers);

        payload.data.forEach(function(row) {
          const alignedRow = alignRowToSheetHeaders(payload.headers, row, sheetHeaders);
          writeRowAsPlainText(sheet, sheet.getLastRow() + 1, alignedRow, sheetHeaders);
        });
      } else {
        // Tanpa properti headers: baris pertama data = header (untuk sheet yang masih kosong)
        payload.data.forEach(function(row) {
          writeRowAsPlainText(sheet, sheet.getLastRow() + 1, row, payload.data[0] || []);
        });
      }

      return jsonResponse({ status: "success", action: "append" });
    }

    // ====================================================
    // 3. ACTION: UPDATE / CREATE MANUAL
    // ====================================================
    if (payload.action === "update") {
      const webHeaders = (payload.headers && Array.isArray(payload.headers)) ? payload.headers : [];
      const sheetHeaders = syncSheetHeaders(sheet, webHeaders);
      const newRow = alignRowToSheetHeaders(webHeaders, payload.data || [], sheetHeaders);

      const allValues = sheet.getDataRange().getValues();
      const targetName = String(payload.productName || "").trim().toLowerCase();
      let rowIndex = -1;
      for (let r = 1; r < allValues.length; r++) {
        if (String(allValues[r][0]).trim().toLowerCase() === targetName) {
          rowIndex = r + 1;
          break;
        }
      }

      if (rowIndex !== -1) {
        writeRowAsPlainText(sheet, rowIndex, newRow, sheetHeaders);
      } else {
        writeRowAsPlainText(sheet, sheet.getLastRow() + 1, newRow, sheetHeaders);
      }

      cleanupEmptyBenefitColumns(sheet);

      return jsonResponse({ status: "success", action: "update" });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali: " + payload.action });

  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

