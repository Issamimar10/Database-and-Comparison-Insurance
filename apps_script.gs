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

// Helper Global: Ambil seluruh konfigurasi aman & data produk dari Master Spreadsheet
// (Memungkinkan Google Spreadsheet dikunci 'Restricted / Private' tanpa perlu dibuka ke publik via /gviz/tq)
function buildAllProductDataResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DEFAULT_TABS = ["Conventional", "Sharia", "Health & Surgical", "DPLK"];
  const configuredTabs = [];
  const nonBenefitColumns = [];
  const benefitAliases = [];
  let driveFolderId = "";

  const configSheet = ss.getSheetByName("Config");
  if (configSheet && configSheet.getLastRow() > 0 && configSheet.getLastColumn() > 0) {
    const cfgValues = configSheet.getDataRange().getDisplayValues();
    for (let i = 1; i < cfgValues.length; i++) {
      const row = cfgValues[i];
      const type = String(row[0] || "").trim().toLowerCase();
      const val = String(row[1] || "").trim();
      const extra = String(row[2] || "").trim();

      if ((type === "sheet" || type === "tab" || type === "kategori" || type === "category") && val) {
        const lowerVal = val.toLowerCase();
        if (
          lowerVal !== "config" &&
          lowerVal !== "documents" &&
          lowerVal !== "users" &&
          !configuredTabs.some(function(t) { return t.toLowerCase() === lowerVal; })
        ) {
          configuredTabs.push(val);
        }
      } else if (type === "non-benefit" && val) {
        if (!nonBenefitColumns.some(function(c) { return c.toLowerCase() === val.toLowerCase(); })) {
          nonBenefitColumns.push(val);
        }
      } else if (type === "benefit" && val) {
        // Baris alias manfaat: Benefit | Nama Standar | Alias 1, Alias 2, ...
        benefitAliases.push({
          canonical: val,
          aliases: splitBenefitAliasList_(extra)
        });
      } else if (type === "setting" && val.toLowerCase() === "drive folder id" && extra) {
        driveFolderId = extra;
      }
      // Catatan Keamanan: Baris 'Invite Code' TIDAK dimasukkan ke response agar tetap rahasia di server
    }
  }

  const tabsToLoad = configuredTabs.length > 0 ? configuredTabs : DEFAULT_TABS;
  const sheetsData = [];

  tabsToLoad.forEach(function(tabName) {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet || sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
      sheetsData.push({
        sheetName: tabName,
        exists: Boolean(sheet),
        headers: [],
        rows: []
      });
      return;
    }

    const displayVals = sheet.getDataRange().getDisplayValues();
    const headers = (displayVals[0] || []).map(function(h) { return String(h || "").trim(); });
    const rows = displayVals.slice(1);

    sheetsData.push({
      sheetName: tabName,
      exists: true,
      headers: headers,
      rows: rows
    });
  });

  return {
    status: "success",
    action: "get_all_product_data",
    config: {
      tabs: tabsToLoad,
      nonBenefitColumns: nonBenefitColumns,
      benefitAliases: benefitAliases,
      driveFolderId: driveFolderId
    },
    sheets: sheetsData
  };
}

// ====================================================
// HELPER: ALIAS NAMA MANFAAT (Config: Benefit | Nama Standar | Alias 1, Alias 2, ...)
// ====================================================
function splitBenefitAliasList_(raw) {
  return String(raw || "")
    .split(/[,;\n]/)
    .map(function(a) { return a.trim(); })
    .filter(function(a) { return a !== ""; });
}

// Kunci pembanding nama manfaat: huruf kecil, tanpa bullet / titik dua di akhir, spasi rapat, bentuk jamak diabaikan
// (HARUS sama dengan normalizeBenefitNameKey() di web)
function normalizeBenefitAliasKey_(name) {
  return String(name || "")
    .toLowerCase()
    .replace(/^[\s\-•*]+/, "")
    .replace(/[:：]\s*$/, "")
    .replace(/\s+/g, " ")
    .trim()
    .split(" ")
    .map(function(w) {
      return (w.length > 3 && /s$/.test(w) && !/(ss|us|is)$/.test(w)) ? w.slice(0, -1) : w;
    })
    .join(" ");
}

// Simpan pemetaan alias -> nama standar ke sheet Config.
// Satu alias hanya boleh menunjuk ke satu nama standar: jika alias sudah terdaftar di nama standar lain, alias dipindahkan.
function saveBenefitAliases_(ss, mappings) {
  let configSheet = ss.getSheetByName("Config");
  if (!configSheet) {
    configSheet = ss.insertSheet("Config");
    configSheet.appendRow(["Type", "Value", "Extra"]);
  }

  const lastRow = configSheet.getLastRow();
  const data = lastRow > 0 ? configSheet.getRange(1, 1, lastRow, 3).getValues() : [];
  const benefitRows = [];
  for (let i = 1; i < data.length; i++) {
    const type = String(data[i][0] || "").trim().toLowerCase();
    const canonical = String(data[i][1] || "").trim();
    if (type === "benefit" && canonical) {
      benefitRows.push({
        rowNum: i + 1,
        canonical: canonical,
        aliases: splitBenefitAliasList_(data[i][2]),
        dirty: false
      });
    }
  }

  let added = 0;
  (mappings || []).forEach(function(m) {
    const canonical = String((m && m.canonical) || "").trim();
    const alias = String((m && m.alias) || "").trim().replace(/[:：]\s*$/, "").trim();
    if (!canonical || !alias) return;
    const canonicalKey = normalizeBenefitAliasKey_(canonical);
    const aliasKey = normalizeBenefitAliasKey_(alias);
    if (!aliasKey || aliasKey === canonicalKey) return;

    // Hapus alias yang sama dari nama standar lain agar tidak bentrok
    benefitRows.forEach(function(r) {
      if (normalizeBenefitAliasKey_(r.canonical) === canonicalKey) return;
      const before = r.aliases.length;
      r.aliases = r.aliases.filter(function(a) { return normalizeBenefitAliasKey_(a) !== aliasKey; });
      if (r.aliases.length !== before) r.dirty = true;
    });

    let target = benefitRows.find(function(r) { return normalizeBenefitAliasKey_(r.canonical) === canonicalKey; });
    if (!target) {
      target = { rowNum: null, canonical: canonical, aliases: [], dirty: true };
      benefitRows.push(target);
    }
    if (!target.aliases.some(function(a) { return normalizeBenefitAliasKey_(a) === aliasKey; })) {
      target.aliases.push(alias);
      target.dirty = true;
      added++;
    }
  });

  benefitRows.forEach(function(r) {
    if (!r.dirty) return;
    const aliasStr = r.aliases.join(", ");
    let rowNum = r.rowNum;
    if (!rowNum) {
      rowNum = configSheet.getLastRow() + 1;
      configSheet.getRange(rowNum, 1, 1, 2).setValues([["Benefit", r.canonical]]);
    }
    const cell = configSheet.getRange(rowNum, 3);
    cell.setNumberFormat("@");
    cell.setValue(aliasStr);
  });

  return added;
}

function doGet(e) {
  // 0. Ambil seluruh data produk & konfigurasi aman langsung via Apps Script (Tanpa perlu /gviz/tq publik)
  if (e && e.parameter && e.parameter.action === "get_all_product_data") {
    try {
      const result = buildAllProductDataResponse();
      return ContentService.createTextOutput(JSON.stringify(result))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

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
      const isSyariah = /^(sharia|syariah)$/i.test(String(sheetName || "").trim());
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
        const sNameLower = String(targetSheet.getName() || "").trim().toLowerCase();
        if (sNameLower === "config" || sNameLower === "users") return;
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

    // 00. ACTION: GET_ALL_PRODUCT_DATA (Via POST)
    if (payload.action === "get_all_product_data") {
      return jsonResponse(buildAllProductDataResponse());
    }

    // ====================================================
    // 0. SISTEM AUTENTIKASI (SIGN UP, SIGN IN, FORGOT PASSWORD OTP EMAIL)
    // ====================================================
    function hashPasswordWithSalt(password, salt) {
      const raw = String(salt || "") + "::IDCM_AUTH::" + String(password || "");
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
      return digest.map(function(b) {
        const v = (b < 0) ? b + 256 : b;
        return ("0" + v.toString(16)).slice(-2);
      }).join("");
    }

    function getOrCreateUsersSheet() {
      let uSheet = ss.getSheetByName("Users");
      const headers = ["Username", "Email", "PasswordHash", "Salt", "CreatedAt", "ResetOTP", "OTPExpiry", "LastLogin"];
      if (!uSheet) {
        uSheet = ss.insertSheet("Users");
        uSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        uSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      } else if (uSheet.getLastRow() === 0) {
        uSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
        uSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
      }
      return uSheet;
    }

    function getSignupInviteCode() {
      const DEFAULT_INVITE_CODE = "IDCM2026";
      try {
        const configSheet = ss.getSheetByName("Config");
        if (!configSheet) return DEFAULT_INVITE_CODE;
        const cfgData = configSheet.getDataRange().getValues();
        for (let i = 1; i < cfgData.length; i++) {
          const key = String(cfgData[i][1] || "").trim().toLowerCase();
          if (key === "invite code" || key === "signup invite code" || key === "kode akses") {
            const codeVal = String(cfgData[i][2] || "").trim();
            if (codeVal) return codeVal;
          }
        }
        // Jika belum ada baris Invite Code di Sheet Config, tambahkan otomatis agar mudah diubah admin
        configSheet.appendRow(["Setting", "Invite Code", DEFAULT_INVITE_CODE]);
      } catch (e) {
        console.warn("Gagal membaca Invite Code dari Config:", e);
      }
      return DEFAULT_INVITE_CODE;
    }

    function maskEmailAddress(email) {
      const str = String(email || "").trim();
      const atIdx = str.indexOf("@");
      if (atIdx <= 1) return str;
      const namePart = str.substring(0, atIdx);
      const domainPart = str.substring(atIdx);
      const visibleLen = Math.min(2, namePart.length);
      return namePart.substring(0, visibleLen) + "***" + domainPart;
    }

    // 0A. ACTION: AUTH_SIGNUP
    if (payload.action === "auth_signup") {
      const username = String(payload.username || "").trim();
      const email = String(payload.email || "").trim().toLowerCase();
      const password = String(payload.password || "");
      const inviteCode = String(payload.inviteCode || "").trim();

      if (username.length < 3) {
        return jsonResponse({ status: "error", message: "Username minimal harus 3 karakter." });
      }
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        return jsonResponse({ status: "error", message: "Format alamat email tidak valid." });
      }
      if (password.length < 6) {
        return jsonResponse({ status: "error", message: "Password minimal harus 6 karakter." });
      }

      const validInviteCode = getSignupInviteCode();
      if (inviteCode !== validInviteCode) {
        return jsonResponse({ status: "error", message: "Kode Akses Tim (Invite Code) tidak sesuai." });
      }

      const uSheet = getOrCreateUsersSheet();
      const uData = uSheet.getDataRange().getValues();
      for (let r = 1; r < uData.length; r++) {
        const existingUser = String(uData[r][0] || "").trim().toLowerCase();
        const existingEmail = String(uData[r][1] || "").trim().toLowerCase();
        if (existingUser === username.toLowerCase()) {
          return jsonResponse({ status: "error", message: "Username sudah digunakan. Silakan pilih username lain." });
        }
        if (existingEmail === email) {
          return jsonResponse({ status: "error", message: "Alamat email sudah terdaftar. Silakan Sign In atau gunakan fitur Lupa Password." });
        }
      }

      const salt = Utilities.getUuid();
      const pwdHash = hashPasswordWithSalt(password, salt);
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");
      writeRowAsPlainText(uSheet, uSheet.getLastRow() + 1, [username, email, pwdHash, salt, nowStr, "", "", nowStr]);

      return jsonResponse({
        status: "success",
        action: "auth_signup",
        user: { username: username, email: email }
      });
    }

    // 0B. ACTION: AUTH_SIGNIN
    if (payload.action === "auth_signin") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      const password = String(payload.password || "");

      if (!identifier || !password) {
        return jsonResponse({ status: "error", message: "Mohon isi Username/Email dan Password." });
      }

      const uSheet = getOrCreateUsersSheet();
      const uData = uSheet.getDataRange().getValues();
      let foundRowIdx = -1;
      for (let r = 1; r < uData.length; r++) {
        const uName = String(uData[r][0] || "").trim().toLowerCase();
        const uEmail = String(uData[r][1] || "").trim().toLowerCase();
        if (uName === identifier || uEmail === identifier) {
          foundRowIdx = r;
          break;
        }
      }

      if (foundRowIdx === -1) {
        return jsonResponse({ status: "error", message: "Akun tidak ditemukan. Periksa kembali Username/Email Anda atau daftar akun baru." });
      }

      const row = uData[foundRowIdx];
      const storedHash = String(row[2] || "").trim();
      const storedSalt = String(row[3] || "").trim();
      const computedHash = hashPasswordWithSalt(password, storedSalt);

      if (computedHash !== storedHash) {
        return jsonResponse({ status: "error", message: "Password yang Anda masukkan salah." });
      }

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");
      uSheet.getRange(foundRowIdx + 1, 8).setNumberFormat("@").setValue(nowStr);

      return jsonResponse({
        status: "success",
        action: "auth_signin",
        user: {
          username: String(row[0] || "").trim(),
          email: String(row[1] || "").trim()
        }
      });
    }

    // 0C. ACTION: AUTH_REQUEST_OTP (Kirim Kode OTP 6 Digit ke Email)
    if (payload.action === "auth_request_otp") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      if (!identifier) {
        return jsonResponse({ status: "error", message: "Masukkan Email atau Username akun Anda." });
      }

      const uSheet = getOrCreateUsersSheet();
      const uData = uSheet.getDataRange().getValues();
      let foundRowIdx = -1;
      for (let r = 1; r < uData.length; r++) {
        const uName = String(uData[r][0] || "").trim().toLowerCase();
        const uEmail = String(uData[r][1] || "").trim().toLowerCase();
        if (uName === identifier || uEmail === identifier) {
          foundRowIdx = r;
          break;
        }
      }

      if (foundRowIdx === -1) {
        return jsonResponse({ status: "error", message: "Akun dengan Email/Username tersebut tidak ditemukan." });
      }

      const targetUsername = String(uData[foundRowIdx][0] || "").trim();
      const targetEmail = String(uData[foundRowIdx][1] || "").trim();
      const otpCode = String(Math.floor(100000 + Math.random() * 900000));
      const expiryMs = String(Date.now() + 10 * 60 * 1000); // Berlaku 10 menit

      uSheet.getRange(foundRowIdx + 1, 6, 1, 2).setNumberFormats([["@", "@"]]).setValues([[otpCode, expiryMs]]);

      const emailHtml = [
        "<div style='font-family:Segoe UI,Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px;border:1px solid #e2e8f0;border-radius:12px;background:#ffffff;color:#1e293b;'>",
        "<h2 style='margin:0 0 12px 0;color:#1e3a8a;font-size:18px;'>Reset Password — Comparison Maker</h2>",
        "<p style='font-size:14px;line-height:1.5;margin:0 0 16px 0;'>Halo <strong>" + targetUsername + "</strong>,</p>",
        "<p style='font-size:14px;line-height:1.5;margin:0 0 16px 0;'>Berikut adalah kode OTP 6 digit untuk mereset password akun Anda (berlaku selama <strong>10 menit</strong>):</p>",
        "<div style='background:#eff6ff;border:1px dashed #3b82f6;border-radius:10px;padding:16px;text-align:center;font-size:28px;font-weight:bold;letter-spacing:6px;color:#1d4ed8;margin-bottom:16px;'>" + otpCode + "</div>",
        "<p style='font-size:12px;color:#64748b;margin:0;'>Jika Anda tidak merasa meminta pengaturan ulang password, abaikan email ini.</p>",
        "</div>"
      ].join("");

      MailApp.sendEmail({
        to: targetEmail,
        subject: "[" + otpCode + "] Kode OTP Reset Password - Life Insurance Comparison Maker",
        htmlBody: emailHtml
      });

      return jsonResponse({
        status: "success",
        action: "auth_request_otp",
        emailMasked: maskEmailAddress(targetEmail),
        message: "Kode OTP 6 digit telah dikirim ke " + maskEmailAddress(targetEmail)
      });
    }

    // 0D. ACTION: AUTH_RESET_PASSWORD (Verifikasi OTP & Simpan Password Baru)
    if (payload.action === "auth_reset_password") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      const otpInput = String(payload.otp || "").trim();
      const newPassword = String(payload.newPassword || "");

      if (!identifier || !otpInput || !newPassword) {
        return jsonResponse({ status: "error", message: "Lengkapi Kode OTP dan Password Baru." });
      }
      if (newPassword.length < 6) {
        return jsonResponse({ status: "error", message: "Password baru minimal harus 6 karakter." });
      }

      const uSheet = getOrCreateUsersSheet();
      const uData = uSheet.getDataRange().getValues();
      let foundRowIdx = -1;
      for (let r = 1; r < uData.length; r++) {
        const uName = String(uData[r][0] || "").trim().toLowerCase();
        const uEmail = String(uData[r][1] || "").trim().toLowerCase();
        if (uName === identifier || uEmail === identifier) {
          foundRowIdx = r;
          break;
        }
      }

      if (foundRowIdx === -1) {
        return jsonResponse({ status: "error", message: "Akun tidak ditemukan." });
      }

      const savedOtp = String(uData[foundRowIdx][5] || "").trim();
      const savedExpiry = Number(uData[foundRowIdx][6] || 0);

      if (!savedOtp || savedOtp !== otpInput) {
        return jsonResponse({ status: "error", message: "Kode OTP tidak valid. Periksa kembali kode 6 digit di email Anda." });
      }
      if (Date.now() > savedExpiry) {
        return jsonResponse({ status: "error", message: "Kode OTP sudah kedaluwarsa (lebih dari 10 menit). Silakan minta kode OTP baru." });
      }

      const newSalt = Utilities.getUuid();
      const newHash = hashPasswordWithSalt(newPassword, newSalt);
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");

      uSheet.getRange(foundRowIdx + 1, 3, 1, 2).setNumberFormats([["@", "@"]]).setValues([[newHash, newSalt]]);
      uSheet.getRange(foundRowIdx + 1, 6, 1, 3).setNumberFormats([["@", "@", "@"]]).setValues([["", "", nowStr]]);

      return jsonResponse({
        status: "success",
        action: "auth_reset_password",
        user: {
          username: String(uData[foundRowIdx][0] || "").trim(),
          email: String(uData[foundRowIdx][1] || "").trim()
        }
      });
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

    // Helper: Cek apakah folder masih memiliki file atau subfolder aktif (tidak di-trash)
    function isFolderCompletelyEmpty(folder) {
      if (!folder) return false;
      const files = folder.getFiles();
      while (files.hasNext()) {
        const f = files.next();
        if (!f.isTrashed()) return false;
      }
      const subFolders = folder.getFolders();
      while (subFolders.hasNext()) {
        const sf = subFolders.next();
        if (!sf.isTrashed()) return false;
      }
      return true;
    }

    // ====================================================
    // 2. ACTION: DELETE_DOCUMENT (Hapus langsung dari Google Drive + Hapus Folder Produk jika Kosong)
    // ====================================================
    if (payload.action === "delete_document") {
      const rootFolder = getOrCreateRootDocFolder();
      const rootFolderId = rootFolder ? rootFolder.getId() : "";
      const parentFoldersToCheck = [];

      if (payload.fileId) {
        try {
          const file = DriveApp.getFileById(payload.fileId);
          const parents = file.getParents();
          while (parents.hasNext()) {
            parentFoldersToCheck.push(parents.next());
          }
          file.setTrashed(true);
        } catch (e) {
          console.error("Gagal trash file di Drive:", e);
        }
      }

      // Jika seluruh dokumen di dalam folder produk telah terhapus (kosong), hapus juga foldernya
      parentFoldersToCheck.forEach(function(parentFolder) {
        try {
          if (
            parentFolder &&
            parentFolder.getId() !== rootFolderId &&
            parentFolder.getName() !== "__TABLE_IMAGES__" &&
            isFolderCompletelyEmpty(parentFolder)
          ) {
            parentFolder.setTrashed(true);
          }
        } catch (e) {
          console.error("Gagal mengecek/menghapus folder kosong:", e);
        }
      });

      // Fallback: jika productName dikirim dan foldernya kosong, hapus foldernya
      if (payload.productName && rootFolder) {
        try {
          const targetProd = String(payload.productName).trim().toLowerCase();
          const subFolders = rootFolder.getFolders();
          while (subFolders.hasNext()) {
            const sf = subFolders.next();
            if (!sf.isTrashed() && sf.getName().trim().toLowerCase() === targetProd && isFolderCompletelyEmpty(sf)) {
              sf.setTrashed(true);
            }
          }
        } catch (e) {}
      }

      // Hapus tab sheet 'Documents' jika masih ada di Master Spreadsheet
      const oldDocSheet = ss.getSheetByName("Documents");
      if (oldDocSheet) {
        try { ss.deleteSheet(oldDocSheet); } catch (e) {}
      }

      return jsonResponse({ status: "success", action: "delete_document" });
    }

    // ====================================================
    // 3. ACTION: GET_FILE_BASE64 (Untuk Ekspor ZIP & Preview Inline)
    //    asPdf: true -> file Word/PowerPoint dikonversi ke PDF khusus untuk preview
    // ====================================================
    if (payload.action === "get_file_base64" && payload.fileId) {
      const file = DriveApp.getFileById(payload.fileId);
      const asPdf = payload.asPdf === true || payload.asPdf === "true";
      let blob = null;
      let mimeType = "";
      let pdfConversionError = "";

      if (asPdf) {
        try {
          blob = convertOfficeFileToPdfBlob(file);
          mimeType = "application/pdf";
        } catch (convErr) {
          pdfConversionError = String(convErr && convErr.message ? convErr.message : convErr);
          console.warn("Konversi PDF gagal, mengembalikan file asli:", pdfConversionError);
          blob = null;
        }
      }

      if (!blob) {
        blob = file.getBlob();
        mimeType = file.getMimeType();
      }

      const resObj = {
        status: "success",
        fileName: file.getName(),
        mimeType: mimeType,
        base64: Utilities.base64Encode(blob.getBytes())
      };
      if (pdfConversionError) resObj.pdfConversionError = pdfConversionError;
      return jsonResponse(resObj);
    }

    // ====================================================
    // 4. ACTION: DELETE (Hapus 1 atau Banyak Produk + Opsi Hapus Gambar Tabel & Folder Dokumen Drive)
    // ====================================================
    if (payload.action === "delete" && Array.isArray(payload.items)) {
      const affectedSheets = {};
      const deleteTableImages = payload.deleteTableImages === true || payload.deleteTableImages === "true";
      const deleteProductDocs = payload.deleteProductDocs === true || payload.deleteProductDocs === "true";
      const tableImageIdsToDelete = {};

      if (deleteTableImages && Array.isArray(payload.tableImageFileIds)) {
        payload.tableImageFileIds.forEach(function(fid) {
          const cleanId = String(fid || "").trim();
          if (cleanId) tableImageIdsToDelete[cleanId] = true;
        });
      }

      const deletedProductNamesLower = {};

      payload.items.forEach(function(item) {
        const targetName = String(item.productName || "").trim().toLowerCase();
        if (targetName) deletedProductNamesLower[targetName] = true;

        const sheet = ss.getSheetByName(item.sheetName);
        if (!sheet) return;

        const dataRange = sheet.getDataRange().getValues();

        for (let r = dataRange.length - 1; r >= 1; r--) {
          const rowProdName = String(dataRange[r][0] || "").trim().toLowerCase();
          if (rowProdName === targetName) {
            // Jika opsi hapus gambar tabel aktif, ekstrak juga seluruh token [IMG_ID:...] dari baris yang dihapus
            if (deleteTableImages) {
              for (let c = 1; c < dataRange[r].length; c++) {
                const cellStr = String(dataRange[r][c] || "");
                const regex = /\[IMG_ID:([a-zA-Z0-9_-]+)\]|\[IMG:drive:([a-zA-Z0-9_-]+)\]/gi;
                let m;
                while ((m = regex.exec(cellStr)) !== null) {
                  const foundId = m[1] || m[2];
                  if (foundId) tableImageIdsToDelete[foundId] = true;
                }
              }
            }
            sheet.deleteRow(r + 1);
            affectedSheets[item.sheetName] = sheet;
          }
        }
      });

      // 1. Hapus file Gambar Tabel di Google Drive jika dicentang
      if (deleteTableImages) {
        Object.keys(tableImageIdsToDelete).forEach(function(imgFileId) {
          try {
            DriveApp.getFileById(imgFileId).setTrashed(true);
          } catch (e) {
            console.error("Gagal menghapus gambar tabel di Drive:", imgFileId, e);
          }
        });
      }

      // 2. Hapus seluruh Dokumen & Folder Produk di Google Drive jika dicentang
      if (deleteProductDocs && Object.keys(deletedProductNamesLower).length > 0) {
        try {
          const rootFolder = getOrCreateRootDocFolder();
          if (rootFolder) {
            const subFolders = rootFolder.getFolders();
            while (subFolders.hasNext()) {
              const sub = subFolders.next();
              if (sub.isTrashed()) continue;
              const folderNameLower = String(sub.getName() || "").trim().toLowerCase();
              if (folderNameLower === "__table_images__") continue;

              if (deletedProductNamesLower[folderNameLower]) {
                const files = sub.getFiles();
                while (files.hasNext()) {
                  try {
                    files.next().setTrashed(true);
                  } catch (fErr) {}
                }
                sub.setTrashed(true);
              }
            }
          }
        } catch (docErr) {
          console.error("Gagal menghapus dokumen/folder produk di Drive:", docErr);
        }
      }

      Object.keys(affectedSheets).forEach(function(sName) {
        cleanupEmptyBenefitColumns(affectedSheets[sName]);
      });

      return jsonResponse({ status: "success", action: "delete" });
    }

    // ====================================================
    // 4b. ACTION: SIMPAN ALIAS NAMA MANFAAT KE CONFIG (dari Review Benefit saat Upload Excel)
    // ====================================================
    if (payload.action === "save_benefit_aliases" && Array.isArray(payload.mappings)) {
      const addedCount = saveBenefitAliases_(ss, payload.mappings);
      return jsonResponse({ status: "success", action: "save_benefit_aliases", added: addedCount });
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
// HELPER: KONVERSI FILE WORD/POWERPOINT (.DOC/.DOCX/.PPT/.PPTX) KE PDF (KHUSUS PREVIEW)
// Alur: salin file sebagai Google Docs/Slides (konversi otomatis Drive) -> ekspor PDF -> hapus salinan sementara.
// Hasil PDF disimpan di cache folder "_pdf_preview_cache" agar preview berikutnya instan.
// ====================================================
function convertOfficeFileToPdfBlob(file) {
  const fileName = String(file.getName() || "");
  const ext = fileName.indexOf(".") !== -1 ? fileName.split(".").pop().toLowerCase() : "";
  const origMime = String(file.getMimeType() || "").toLowerCase();

  // File yang sudah berupa Google Docs/Slides bisa langsung diekspor ke PDF
  if (origMime === "application/vnd.google-apps.document" || origMime === "application/vnd.google-apps.presentation") {
    return file.getAs("application/pdf");
  }

  let targetGoogleMime = "";
  if (ext === "doc" || ext === "docx" || origMime.indexOf("word") !== -1) {
    targetGoogleMime = "application/vnd.google-apps.document";
  } else if (ext === "ppt" || ext === "pptx" || origMime.indexOf("presentation") !== -1 || origMime.indexOf("powerpoint") !== -1) {
    targetGoogleMime = "application/vnd.google-apps.presentation";
  }
  if (!targetGoogleMime) {
    throw new Error("Format file tidak didukung untuk konversi PDF: " + (ext || origMime));
  }

  // 1. Bersihkan cache PDF lama secara otomatis (maksimal 1x per hari) lalu cek cache
  maybeCleanupPdfPreviewCacheDaily_();
  const cacheName = "pdfprev_" + file.getId() + "_" + file.getLastUpdated().getTime() + ".pdf";
  const cacheFolder = getPdfPreviewCacheFolder_();
  const cached = cacheFolder.getFilesByName(cacheName);
  if (cached.hasNext()) {
    return cached.next().getBlob().setContentType("application/pdf");
  }

  // 2. Buat salinan sementara dalam format Google (konversi otomatis oleh Drive)
  let tempFileId = null;
  const errors = [];
  try {
    // 2a. Advanced Drive Service (jika sudah ditambahkan di menu Services)
    if (typeof Drive !== "undefined" && Drive.Files && typeof Drive.Files.copy === "function") {
      try {
        // Drive API v3
        const copiedV3 = Drive.Files.copy(
          { name: "_temp_pdf_" + fileName, mimeType: targetGoogleMime },
          file.getId(),
          { supportsAllDrives: true }
        );
        if (copiedV3 && copiedV3.id) tempFileId = copiedV3.id;
      } catch (v3Err) {
        errors.push("Drive v3: " + (v3Err.message || v3Err));
        try {
          // Drive API v2
          const copiedV2 = Drive.Files.copy(
            { title: "_temp_pdf_" + fileName },
            file.getId(),
            { convert: true, supportsAllDrives: true }
          );
          if (copiedV2 && copiedV2.id) tempFileId = copiedV2.id;
        } catch (v2Err) {
          errors.push("Drive v2: " + (v2Err.message || v2Err));
        }
      }
    }

    // Jika salinan dari Advanced Service tidak terkonversi, buang dan lanjut ke fallback REST
    if (tempFileId) {
      try {
        if (String(DriveApp.getFileById(tempFileId).getMimeType()) !== targetGoogleMime) {
          DriveApp.getFileById(tempFileId).setTrashed(true);
          errors.push("Advanced Drive copy was not converted");
          tempFileId = null;
        }
      } catch (chkErr) { }
    }

    // 2b. Fallback: Drive REST API v3 via UrlFetchApp (tidak perlu Advanced Service)
    if (!tempFileId) {
      const copyUrl = "https://www.googleapis.com/drive/v3/files/" + encodeURIComponent(file.getId()) + "/copy?supportsAllDrives=true&fields=id";
      const resp = UrlFetchApp.fetch(copyUrl, {
        method: "post",
        contentType: "application/json",
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ name: "_temp_pdf_" + fileName, mimeType: targetGoogleMime }),
        muteHttpExceptions: true
      });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        const parsed = JSON.parse(resp.getContentText());
        if (parsed && parsed.id) tempFileId = parsed.id;
      } else {
        errors.push("Drive REST HTTP " + code + ": " + resp.getContentText().substring(0, 300));
      }
    }

    if (!tempFileId) {
      throw new Error(errors.join(" | ") || "Gagal membuat salinan sementara untuk konversi PDF.");
    }

    // Pastikan salinan benar-benar sudah terkonversi ke format Google (bukan salinan Office biasa)
    const tempFile = DriveApp.getFileById(tempFileId);
    if (String(tempFile.getMimeType()) !== targetGoogleMime) {
      throw new Error("Salinan tidak terkonversi ke format Google (mime: " + tempFile.getMimeType() + "). " + errors.join(" | "));
    }

    // 3. Ekspor salinan Google Docs/Slides ke PDF & simpan ke cache
    const pdfBlob = tempFile.getAs("application/pdf");
    pdfBlob.setName(cacheName);
    try {
      cacheFolder.createFile(pdfBlob);
    } catch (cacheErr) {
      console.warn("Gagal menyimpan cache PDF:", cacheErr);
    }
    return pdfBlob;
  } finally {
    // 4. Hapus salinan sementara
    if (tempFileId) {
      try { DriveApp.getFileById(tempFileId).setTrashed(true); } catch (cleanupErr) { }
    }
  }
}

// Folder cache PDF preview (dicari berdasarkan nama di seluruh Drive, sehingga aman dipindahkan ke folder main secara manual)
function getPdfPreviewCacheFolder_() {
  const folderName = "_pdf_preview_cache";
  const it = DriveApp.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(folderName);
}

// Pembersihan otomatis ringan maksimal 1x per hari (tanpa perlu setup Trigger manual)
function maybeCleanupPdfPreviewCacheDaily_() {
  try {
    const props = PropertiesService.getScriptProperties();
    const todayKey = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "Asia/Jakarta", "yyyy-MM-dd");
    const lastDay = props.getProperty("LAST_PDF_CACHE_CLEANUP_DAY");
    if (lastDay === todayKey) return;
    props.setProperty("LAST_PDF_CACHE_CLEANUP_DAY", todayKey);
    cleanupPdfPreviewCache(24);
  } catch (e) {
    console.warn("Gagal cek pembersihan cache PDF harian:", e);
  }
}

// Hapus isi file cache PDF yang berumur >= maxAgeHours (default 24 jam).
// Foldernya sendiri TIDAK dihapus agar lokasi manual di folder main tetap terjaga.
// Bisa juga dipasang di Apps Script Triggers (Time-driven -> Day timer) untuk pembersihan rutin jam malam.
function cleanupPdfPreviewCache(maxAgeHours) {
  const hours = (typeof maxAgeHours === "number" && !isNaN(maxAgeHours)) ? maxAgeHours : 24;
  const cutoffMs = Date.now() - (hours * 60 * 60 * 1000);
  const folder = getPdfPreviewCacheFolder_();
  const files = folder.getFiles();
  let deletedCount = 0;
  while (files.hasNext()) {
    const f = files.next();
    try {
      if (hours <= 0 || f.getLastUpdated().getTime() <= cutoffMs) {
        f.setTrashed(true);
        deletedCount++;
      }
    } catch (e) { }
  }
  console.log("cleanupPdfPreviewCache selesai. File cache dihapus:", deletedCount);
  return deletedCount;
}

// ====================================================
// FUNGSI KHUSUS UNTUK MEMICU POP-UP IZIN GOOGLE DRIVE & EMAIL (MAILAPP)
// ====================================================
function authorizeDrive() {
  const tempFile = DriveApp.createFile("test_permission.txt", "OK");
  tempFile.setTrashed(true);
  const quota = MailApp.getRemainingDailyQuota();
  // Memicu izin "Connect to an external service" (UrlFetchApp) yang dibutuhkan konversi Word/PPT -> PDF
  const resp = UrlFetchApp.fetch("https://www.googleapis.com/drive/v3/about?fields=user", {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });
  getPdfPreviewCacheFolder_();
  console.log("Izin Google Drive, MailApp & UrlFetchApp Berhasil Aktif! Sisa kuota email harian:", quota, "| Drive API HTTP:", resp.getResponseCode());
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
