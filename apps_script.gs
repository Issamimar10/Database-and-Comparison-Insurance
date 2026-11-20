// ====================================================
// COMPARISON MAKER - GOOGLE APPS SCRIPT BACKEND
// Versi: 2.1 (Penyimpanan Dokumen Langsung ke Google Drive Tanpa Sheet 'Documents')
// ====================================================

// Helper Global: Ambil atau Gunakan Folder Google Drive untuk Dokumen (Folder "IDCM Test")
// Di-cache 10 menit di ScriptCache agar aksi Drive-only tidak perlu membuka SpreadsheetApp setiap request.
function getOrCreateRootDocFolder() {
  const cache = CacheService.getScriptCache();
  const CACHE_KEY = "ROOT_DOC_FOLDER_ID_V1";
  try {
    const cachedId = cache.get(CACHE_KEY);
    if (cachedId) {
      const cachedFolder = DriveApp.getFolderById(cachedId);
      if (cachedFolder && !cachedFolder.isTrashed()) {
        return cachedFolder;
      }
    }
  } catch (e) {}

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
      const candidate = DriveApp.getFolderById(folderId);
      if (candidate && !candidate.isTrashed()) {
        rootFolder = candidate;
      }
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

  if (rootFolder) {
    try { cache.put(CACHE_KEY, rootFolder.getId(), 600); } catch (e) {}
  }
  return rootFolder;
}

// ====================================================
// HELPER KEAMANAN GLOBAL (SESSION, VALIDASI FILE, RESERVED SHEETS, ESCAPING, ANTI FORMULA INJECTION)
// ====================================================
const RESERVED_SYSTEM_SHEETS_ = {
  "config": true,
  "users": true,
  "documents": true
};
const MAX_UPLOAD_BASE64_CHARS = 35 * 1024 * 1024; // ~25 MB biner
const MAX_APPEND_ROWS = 2000;

// Cegah Spreadsheet Formula Injection (Google Sheets & Excel/CSV):
// Google Sheets tetap mengevaluasi rumus "=..." saat ditulis via setValue/setValues meskipun format sel "@".
// Memberi awalan apostrof (') memaksa Google Sheets menyimpan nilai sebagai teks literal tanpa dievaluasi.
function sanitizeSheetCell_(val) {
  if (val === undefined || val === null) return "";
  const s = String(val);
  if (s === "" || s === "-") return s;
  if (/^[\t\r\n]*[=+\-@]/.test(s) || /^[\t\r]/.test(s)) {
    return "'" + s;
  }
  return s;
}

function sanitizeSheetRow_(row) {
  return (Array.isArray(row) ? row : []).map(sanitizeSheetCell_);
}

function sanitizeSheetMatrix_(rows) {
  return (Array.isArray(rows) ? rows : []).map(sanitizeSheetRow_);
}
const SESSION_DEFAULT_TTL_MS = 24 * 60 * 60 * 1000;
const SESSION_REMEMBER_TTL_MS = 7 * 24 * 60 * 60 * 1000;

function isReservedSystemSheet_(sheetName) {
  const lower = String(sheetName || "").trim().toLowerCase();
  return Boolean(lower && RESERVED_SYSTEM_SHEETS_[lower]);
}

function escapeHtmlServer_(str) {
  return String(str === undefined || str === null ? "" : str)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function constantTimeEquals_(a, b) {
  const strA = String(a === undefined || a === null ? "" : a);
  const strB = String(b === undefined || b === null ? "" : b);
  const maxLen = Math.max(strA.length, strB.length);
  let diff = strA.length ^ strB.length;
  for (let i = 0; i < maxLen; i++) {
    const ca = i < strA.length ? strA.charCodeAt(i) : 0;
    const cb = i < strB.length ? strB.charCodeAt(i) : 0;
    diff |= (ca ^ cb);
  }
  return diff === 0;
}

function bytesToHex_(bytes) {
  return bytes.map(function(b) {
    const v = (b < 0) ? b + 256 : b;
    return ("0" + v.toString(16)).slice(-2);
  }).join("");
}

function getOrCreateSessionSecret_() {
  const props = PropertiesService.getScriptProperties();
  let secret = props.getProperty("IDCM_SESSION_HMAC_SECRET");
  if (!secret) {
    secret = Utilities.getUuid() + ":" + Utilities.getUuid();
    props.setProperty("IDCM_SESSION_HMAC_SECRET", secret);
  }
  return secret;
}

const MAX_ACTIVE_SESSIONS_PER_USER = 5;

function pruneAndCapUserSessions_(props, targetUsernameLower) {
  try {
    const all = props.getProperties();
    const now = Date.now();
    const userSessions = [];
    Object.keys(all).forEach(function(k) {
      if (k.indexOf("SESSION_") === 0) {
        try {
          const parsed = JSON.parse(all[k]);
          if (!parsed || !parsed.expiresAt || now > Number(parsed.expiresAt)) {
            props.deleteProperty(k);
          } else if (targetUsernameLower && String(parsed.username || "").trim().toLowerCase() === targetUsernameLower) {
            userSessions.push({ key: k, issuedAt: Number(parsed.issuedAt) || 0 });
          }
        } catch (e) {
          props.deleteProperty(k);
        }
      }
    });
    if (targetUsernameLower && userSessions.length > MAX_ACTIVE_SESSIONS_PER_USER) {
      userSessions.sort(function(a, b) { return a.issuedAt - b.issuedAt; });
      const excessCount = userSessions.length - MAX_ACTIVE_SESSIONS_PER_USER;
      const cache = CacheService.getScriptCache();
      for (let i = 0; i < excessCount; i++) {
        props.deleteProperty(userSessions[i].key);
        try { cache.remove(userSessions[i].key); } catch (e) {}
      }
    }
  } catch (e) {}
}

function pruneExpiredSessions_(props) {
  pruneAndCapUserSessions_(props, "");
}

function normalizeUserRole_(rawRole) {
  const r = String(rawRole || "").trim().toLowerCase();
  if (r === "admin" || r === "administrator") return "admin";
  if (r === "editor") return "editor";
  return "viewer";
}

function revokeAllSessionsForUser_(targetUsernameLower) {
  const cleanTarget = String(targetUsernameLower || "").trim().toLowerCase();
  if (!cleanTarget) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const cache = CacheService.getScriptCache();
    const all = props.getProperties();
    Object.keys(all).forEach(function(k) {
      if (k.indexOf("SESSION_") === 0) {
        try {
          const parsed = JSON.parse(all[k]);
          if (parsed && String(parsed.username || "").trim().toLowerCase() === cleanTarget) {
            props.deleteProperty(k);
            try { cache.remove(k); } catch (e) {}
          }
        } catch (e) {}
      }
    });
  } catch (e) {}
}

function issueSessionToken_(username, email, ttlMs, role) {
  const cleanUser = String(username || "").trim();
  const cleanEmail = String(email || "").trim().toLowerCase();
  const cleanRole = normalizeUserRole_(role);
  const duration = (typeof ttlMs === "number" && ttlMs > 0)
    ? Math.min(ttlMs, SESSION_REMEMBER_TTL_MS)
    : SESSION_DEFAULT_TTL_MS;
  const now = Date.now();
  const expiresAt = now + duration;
  const secret = getOrCreateSessionSecret_();
  const entropy = Utilities.getUuid() + "::" + Utilities.getUuid() + "::" + cleanUser + "::" + now;
  const sigBytes = Utilities.computeHmacSha256Signature(entropy, secret);
  const token = bytesToHex_(sigBytes);
  const record = {
    username: cleanUser,
    email: cleanEmail,
    role: cleanRole,
    issuedAt: now,
    expiresAt: expiresAt
  };
  const serialized = JSON.stringify(record);
  const key = "SESSION_" + token;

  try {
    const cache = CacheService.getScriptCache();
    cache.put(key, serialized, Math.min(21600, Math.max(60, Math.floor(duration / 1000))));
  } catch (e) {}

  const props = PropertiesService.getScriptProperties();
  props.setProperty(key, serialized);
  pruneAndCapUserSessions_(props, cleanUser.toLowerCase());

  return { token: token, expiresAt: expiresAt, role: cleanRole };
}

function verifySessionToken_(token) {
  const cleanToken = String(token || "").trim();
  if (!/^[a-f0-9]{64}$/i.test(cleanToken)) return null;
  const key = "SESSION_" + cleanToken.toLowerCase();

  try {
    const cached = CacheService.getScriptCache().get(key);
    if (cached) {
      const parsedC = JSON.parse(cached);
      if (parsedC && parsedC.username && Date.now() <= Number(parsedC.expiresAt)) {
        parsedC.role = normalizeUserRole_(parsedC.role);
        return parsedC;
      }
    }
  } catch (e) {}

  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty(key);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (!parsed || !parsed.username || Date.now() > Number(parsed.expiresAt)) {
      props.deleteProperty(key);
      return null;
    }
    parsed.role = normalizeUserRole_(parsed.role);
    try {
      const remSec = Math.min(21600, Math.max(60, Math.floor((Number(parsed.expiresAt) - Date.now()) / 1000)));
      CacheService.getScriptCache().put(key, raw, remSec);
    } catch (e) {}
    return parsed;
  } catch (e) {
    return null;
  }
}

function revokeSessionToken_(token) {
  const cleanToken = String(token || "").trim();
  if (!/^[a-f0-9]{64}$/i.test(cleanToken)) return;
  const key = "SESSION_" + cleanToken.toLowerCase();
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
  try { PropertiesService.getScriptProperties().deleteProperty(key); } catch (e) {}
}

function revokeAllUserSessions_(username) {
  const targetLower = String(username || "").trim().toLowerCase();
  if (!targetLower) return;
  try {
    const props = PropertiesService.getScriptProperties();
    const all = props.getProperties();
    const cache = CacheService.getScriptCache();
    Object.keys(all).forEach(function(k) {
      if (k.indexOf("SESSION_") === 0) {
        try {
          const parsed = JSON.parse(all[k]);
          if (parsed && String(parsed.username || "").trim().toLowerCase() === targetLower) {
            props.deleteProperty(k);
            try { cache.remove(k); } catch (ce) {}
          }
        } catch (e) {}
      }
    });
  } catch (e) {}
}

function isFileInsideRootFolder_(file, rootFolder) {
  if (!file || !rootFolder) return false;
  const rootId = rootFolder.getId();
  const visited = {};
  const queue = [];

  try {
    const parents = file.getParents();
    while (parents.hasNext()) {
      const p = parents.next();
      if (p) queue.push({ folder: p, depth: 1 });
    }
  } catch (e) {
    return false;
  }

  while (queue.length > 0) {
    const item = queue.shift();
    const f = item.folder;
    const fid = f.getId();
    if (fid === rootId) return true;
    if (visited[fid] || item.depth >= 4) continue;
    visited[fid] = true;
    try {
      const gp = f.getParents();
      while (gp.hasNext()) {
        const nextF = gp.next();
        if (nextF && !visited[nextF.getId()]) {
          queue.push({ folder: nextF, depth: item.depth + 1 });
        }
      }
    } catch (e) {}
  }
  return false;
}

function assertAccessibleProductFile_(fileId) {
  const cleanId = String(fileId || "").trim();
  if (!cleanId) throw new Error("fileId tidak valid.");
  const file = DriveApp.getFileById(cleanId);
  if (file.isTrashed()) {
    throw new Error("File sudah dihapus (trashed).");
  }
  const parents = file.getParents();
  while (parents.hasNext()) {
    const pName = String(parents.next().getName() || "").trim().toLowerCase();
    if (pName === "_pdf_preview_cache") {
      throw new Error("Akses ditolak: folder cache internal tidak dapat diakses langsung.");
    }
  }
  const rootFolder = getOrCreateRootDocFolder();
  if (!isFileInsideRootFolder_(file, rootFolder)) {
    throw new Error("Akses ditolak: file berada di luar folder dokumen aplikasi.");
  }
  return file;
}

function purgePdfPreviewCacheForFile_(fileId) {
  const cleanId = String(fileId || "").trim();
  if (!cleanId) return;
  try {
    const cacheFolder = getPdfPreviewCacheFolder_();
    const prefix = "pdfprev_" + cleanId + "_";
    const files = cacheFolder.getFiles();
    while (files.hasNext()) {
      const f = files.next();
      try {
        if (String(f.getName() || "").indexOf(prefix) === 0) {
          f.setTrashed(true);
        }
      } catch (e) {}
    }
  } catch (e) {}
}

function secureTrashDriveFile_(file) {
  if (!file) return;
  const fid = file.getId();
  try {
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE);
  } catch (e) {}
  purgePdfPreviewCacheForFile_(fid);
  file.setTrashed(true);
}

// Helper: Baca seluruh tab produk sekaligus via Google Sheets Advanced Service (batchGet),
// atau fallback otomatis ke sheet.getDataRange().getDisplayValues() jika Sheets API belum diaktifkan.
function readSheetsBatchOrFallback_(ss, tabNames) {
  const allSheets = ss.getSheets();
  const sheetByLowerName = {};
  allSheets.forEach(function(s) {
    sheetByLowerName[String(s.getName() || "").trim().toLowerCase()] = s;
  });

  const resultMap = {};
  const existingTabs = [];

  tabNames.forEach(function(tabName) {
    const exact = ss.getSheetByName(tabName) || sheetByLowerName[String(tabName || "").trim().toLowerCase()] || null;
    if (!exact) {
      resultMap[tabName] = { exists: false, values: [] };
    } else {
      existingTabs.push({ tabName: tabName, sheet: exact, actualName: exact.getName() });
    }
  });

  if (existingTabs.length === 0) {
    return resultMap;
  }

  // Coba batch read 1x panggilan via Google Sheets Advanced Service
  if (typeof Sheets !== "undefined" && Sheets.Spreadsheets && Sheets.Spreadsheets.Values && typeof Sheets.Spreadsheets.Values.batchGet === "function") {
    try {
      const ranges = existingTabs.map(function(item) {
        const escaped = String(item.actualName || "").replace(/'/g, "''");
        return "'" + escaped + "'!A:ZZ";
      });
      const batchRes = Sheets.Spreadsheets.Values.batchGet(ss.getId(), {
        ranges: ranges,
        valueRenderOption: "FORMATTED_VALUE"
      });
      const valueRanges = (batchRes && Array.isArray(batchRes.valueRanges)) ? batchRes.valueRanges : [];
      if (valueRanges.length === existingTabs.length) {
        existingTabs.forEach(function(item, idx) {
          const rawVals = (valueRanges[idx] && Array.isArray(valueRanges[idx].values)) ? valueRanges[idx].values : [];
          if (rawVals.length === 0) {
            resultMap[item.tabName] = { exists: true, values: [] };
            return;
          }
          const headerLen = (rawVals[0] || []).length;
          const padded = rawVals.map(function(row) {
            const out = new Array(headerLen);
            for (let c = 0; c < headerLen; c++) {
              out[c] = (row && row[c] !== undefined && row[c] !== null) ? String(row[c]) : "";
            }
            return out;
          });
          resultMap[item.tabName] = { exists: true, values: padded };
        });
        return resultMap;
      }
    } catch (batchErr) {
      console.warn("Sheets.Spreadsheets.Values.batchGet fallback ke SpreadsheetApp:", batchErr);
    }
  }

  // Fallback standar SpreadsheetApp per tab
  existingTabs.forEach(function(item) {
    const sheet = item.sheet;
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
      resultMap[item.tabName] = { exists: true, values: [] };
      return;
    }
    resultMap[item.tabName] = {
      exists: true,
      values: sheet.getDataRange().getDisplayValues()
    };
  });

  return resultMap;
}

// Helper: Pangkas baris sheet menjadi payload ringkas (mode="summary") untuk render cepat kartu awal.
// Hanya menyertakan kolom identitas inti + 1 snippet manfaat preview (atau flag "✓" untuk DPLK).
function buildSummaryRowsForSheet_(tabName, headers, rows, nonBenefitColumns) {
  const CORE_COLS_LOWER = {
    "product name": true,
    "company name": true,
    "launch date": true,
    "tanggal rilis": true,
    "distribution channel": true,
    "currency": true,
    "product type": true,
    "par/non par": true,
    "underwriting method": true
  };

  const STANDARD_NON_BENEFIT_LOWER = {
    "launch date": true,
    "tanggal rilis": true,
    "additional benefit": true,
    "aqd": true,
    "akad": true,
    "aqd / akad": true,
    "alokasi kontribusi": true,
    "surplus underwriting": true,
    "entry age": true,
    "coverage period": true,
    "premium payment period (ppp)": true,
    "min premium": true,
    "min sum assured": true,
    "underwriting method": true,
    "available riders": true,
    "sales tools": true,
    "after sales tools": true,
    "remarks": true
  };

  const customNonBenefitLower = {};
  (nonBenefitColumns || []).forEach(function(nb) {
    const lw = String(nb || "").trim().toLowerCase();
    if (lw) customNonBenefitLower[lw] = true;
  });

  const isDplkTab = /^dplk$/i.test(String(tabName || "").trim());
  const colMeta = headers.map(function(h) {
    const clean = String(h || "").trim();
    const lower = clean.toLowerCase();
    const isCont = /\[\d+\]$/.test(clean);
    const isCore = Boolean(CORE_COLS_LOWER[lower]);
    const isNonBen = Boolean(STANDARD_NON_BENEFIT_LOWER[lower] || customNonBenefitLower[lower]);
    const isCharge = /^biaya\b/i.test(lower);
    const isBenefit = Boolean(clean) && !isCore && !isCont && !isNonBen && !isCharge;
    return {
      lower: lower,
      isCore: isCore,
      isBenefit: isBenefit,
      isDeathBenefit: isBenefit && lower.indexOf("death benefit") !== -1
    };
  });

  const isMeaningfulCell = function(val) {
    const s = String(val === undefined || val === null ? "" : val).trim();
    if (!s || s === "-" || /^n\/?a\.?$/i.test(s) || s.toLowerCase() === "null" || s.toLowerCase() === "undefined") {
      return false;
    }
    return true;
  };

  const stripImgTokens = function(val) {
    return String(val || "")
      .replace(/\[IMG_ID:[^\]]+\]/gi, "")
      .replace(/\[IMG:[^\]]+\]/gi, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  };

  return rows.map(function(row) {
    const out = new Array(headers.length);
    for (let c = 0; c < headers.length; c++) {
      const meta = colMeta[c];
      const rawCell = (row && row[c] !== undefined && row[c] !== null) ? String(row[c]) : "";
      if (meta && meta.isCore) {
        out[c] = rawCell;
      } else if (isDplkTab && meta && meta.isBenefit) {
        const trimmed = rawCell.trim();
        const isChecked = /^(✓|✅)/.test(trimmed) || /^(v|check|checked)(?=\s*$|\s*\||\s*:|\r?\n)/i.test(trimmed);
        out[c] = isChecked ? "✓" : "";
      } else {
        out[c] = "";
      }
    }

    if (!isDplkTab) {
      let chosenIdx = -1;
      let chosenCleanText = "";

      for (let c = 0; c < headers.length; c++) {
        if (!colMeta[c] || !colMeta[c].isDeathBenefit) continue;
        const rawVal = (row && row[c] !== undefined && row[c] !== null) ? String(row[c]) : "";
        if (!isMeaningfulCell(rawVal)) continue;
        const stripped = stripImgTokens(rawVal);
        if (isMeaningfulCell(stripped)) {
          chosenIdx = c;
          chosenCleanText = stripped;
          break;
        }
      }

      if (chosenIdx === -1) {
        for (let c = 0; c < headers.length; c++) {
          if (!colMeta[c] || !colMeta[c].isBenefit) continue;
          const rawVal = (row && row[c] !== undefined && row[c] !== null) ? String(row[c]) : "";
          if (!isMeaningfulCell(rawVal)) continue;
          const stripped = stripImgTokens(rawVal);
          if (isMeaningfulCell(stripped)) {
            chosenIdx = c;
            chosenCleanText = stripped;
            break;
          }
        }
      }

      if (chosenIdx !== -1 && chosenCleanText) {
        out[chosenIdx] = chosenCleanText.length > 240
          ? chosenCleanText.slice(0, 240) + "..."
          : chosenCleanText;
      }
    }

    return out;
  });
}

// ====================================================
// DEFAULT KONFIGURASI SHEET 'Config' (4 KOLOM: Kategori | Nama Utama | Nilai / Nama Lain | Keterangan)
// Mudah diedit oleh orang awam langsung dari Google Sheets tanpa perlu mengubah kode function.
// ====================================================
const CONFIG_HEADERS_4COL_ = ["Kategori", "Nama Utama", "Nilai / Nama Lain", "Keterangan"];

const DEFAULT_CONFIG_ROWS_ = [
  // 1. Pengaturan Sistem
  ["Setting", "Drive Folder ID", "", "ID folder utama Google Drive tempat menyimpan dokumen produk (kosongkan untuk otomatis memakai folder IDCM Test)"],
  ["Setting", "Invite Code", "", "Kode akses rahasia untuk pendaftaran akun baru (wajib diisi oleh Admin sebelum Sign Up dapat digunakan)"],

  // 2. Daftar Sheet / Kategori Produk
  ["Sheet", "Conventional", "Traditional, Konven", "Nama tab sheet produk konvensional di Master Spreadsheet"],
  ["Sheet", "Sharia", "Syariah", "Nama tab sheet produk syariah di Master Spreadsheet"],
  ["Sheet", "Health & Surgical", "Health, Kesehatan", "Nama tab sheet produk kesehatan di Master Spreadsheet"],
  ["Sheet", "DPLK", "Dana Pensiun", "Nama tab sheet produk DPLK di Master Spreadsheet"],

  // 3. Kolom Identitas Inti (Core)
  ["Core", "Product Name", "Nama Produk", "Kolom identitas utama nama produk"],
  ["Core", "Company Name", "Nama Perusahaan, Company", "Kolom identitas nama perusahaan asuransi"],
  ["Core", "Par/Non Par", "Par / Non Par", "Kategori partisipasi produk (Participating / Non-Participating)"],
  ["Core", "Product Type", "Tipe Produk", "Jenis produk asuransi (Whole Life, Term Life, Endowment, Unit Link, dll.)"],
  ["Core", "Distribution Channel", "Channel Distribution", "Jalur distribusi pemasaran produk"],
  ["Core", "Currency", "Mata Uang", "Mata uang polis (IDR, USD, dll.)"],
  ["Core", "Launch Date", "Tanggal Rilis, Launching Date, Release Date", "Tanggal atau tahun peluncuran produk"],
  ["Core", "Remarks", "Catatan", "Catatan tambahan produk"],

  // 4. Kolom Non-Benefit (Ketentuan Produk Umum — Di-lock / Tidak dihapus otomatis)
  ["Non-Benefit", "Entry Age", "Insured Entry Age, Age Definition, Usia Masuk", "Ketentuan usia masuk tertanggung / pemegang polis"],
  ["Non-Benefit", "Coverage Period", "Coverage Term, Masa Pertanggungan", "Masa pertanggungan asuransi"],
  ["Non-Benefit", "Premium Payment Period (PPP)", "PPP, Premium Payment Period, Premium Term, Masa Pembayaran Premi", "Masa pembayaran premi / kontribusi"],
  ["Non-Benefit", "Min Premium", "Premium (Iuran), Iuran, Premium Limit, Minimum Premium, Premium", "Ketentuan minimum premi / iuran"],
  ["Non-Benefit", "Min Sum Assured", "Sum Assured Limit, Minimum Sum Assured, Sum Assured", "Ketentuan minimum uang pertanggungan (UP)"],
  ["Non-Benefit", "Underwriting Method", "Underwriting", "Metode seleksi risiko / underwriting (GIO, SIO, Full Underwriting, NMU)"],
  ["Non-Benefit", "Available Riders", "Available Rider, Riders, Rider", "Daftar asuransi tambahan (Rider) yang tersedia"],
  ["Non-Benefit", "Additional Benefit", "Additional Feature (if any), Additional Feature, Additional Features, Other Benefit, Other Feature, Fitur Tambahan", "Fitur atau manfaat tambahan di luar manfaat utama"],

  // 5. Kolom Non-Benefit Khusus Syariah & DPLK (Di-lock / Tidak dihapus otomatis)
  ["Non-Benefit (Syariah)", "Aqd", "Akad, Aqd / Akad, Akad / Aqd", "Jenis akad produk Syariah"],
  ["Non-Benefit (Syariah)", "Alokasi Kontribusi", "Tabarru, Tanahud, Ujrah, Nisbah", "Ketentuan alokasi kontribusi / dana Tabarru & Ujrah Syariah"],
  ["Non-Benefit (Syariah)", "Surplus Underwriting", "Underwriting Surplus", "Ketentuan pembagian surplus underwriting Syariah"],
  ["Non-Benefit (DPLK)", "Sales Tools", "Sales Tool", "Alat bantu penjualan khusus produk DPLK"],
  ["Non-Benefit (DPLK)", "After Sales Tools", "After Sales Tool", "Alat bantu layanan purna jual khusus produk DPLK"],

  // 6. Daftar Benefit Umum & Nama Lainnya (Otomatis dihapus jika sudah tidak dipakai produk mana pun)
  ["Benefit", "Death Benefit", "Death Benefits, DB", "Manfaat meninggal dunia"],
  ["Benefit", "Accidental Death Benefit", "Accident Death Benefit, Accidental Death Benefits, ADB", "Manfaat meninggal dunia akibat kecelakaan"],
  ["Benefit", "Accident Death During Hajj (Pilgrim) Benefit", "Hajj Death Benefit", "Manfaat meninggal dunia saat menunaikan ibadah Haji / Umrah"],
  ["Benefit", "Total Permanent Disability Benefit", "TPD, Total Permanent Disability, TPD Benefit", "Manfaat cacat tetap total"],
  ["Benefit", "Terminal Illness Benefit", "Terminal Illness", "Manfaat penyakit terminal"],
  ["Benefit", "Critical Illness Benefit", "CI Benefit, CI Benefits, Critical Illness", "Manfaat penyakit kritis"],
  ["Benefit", "Additional Death Benefit", "Additional Death Benefits", "Manfaat tambahan meninggal dunia"],
  ["Benefit", "Maturity Benefit", "Maturity Benefits", "Manfaat akhir masa polis / jatuh tempo"],
  ["Benefit", "Cash Payout Bonus", "Cash Payout Benefit", "Manfaat tahapan / pembayaran tunai berkala"],
  ["Benefit", "Sum Assured Bonus", "SA Bonus, Sum Assured Benefit", "Bonus uang pertanggungan"],
  ["Benefit", "Loyalty Bonus", "Bonus Loyalitas", "Manfaat bonus loyalitas polis"],
  ["Benefit", "Booster Benefit", "Investment Booster", "Manfaat booster nilai investasi / uang pertanggungan"],
  ["Benefit", "Survival Benefit", "Manfaat Hidup", "Manfaat hidup berkala"],
  ["Benefit", "Manfaat Dana Tahapan", "Dana Tahapan", "Manfaat pembayaran dana tahapan terjadwal"],
  ["Benefit", "Education Fund Benefit", "Manfaat Dana Pendidikan", "Manfaat dana tahapan pendidikan"],
  ["Benefit", "Hospitalization Benefit", "Rawat Inap", "Manfaat santunan rawat inap rumah sakit"],
  ["Benefit", "Premium Waiver", "Waiver of Premium", "Manfaat pembebasan premi"],
  ["Benefit", "Remaining Premium Waived", "Waived Remaining Premium", "Manfaat pembebasan sisa premi yang belum jatuh tempo"],

  // 7. Daftar Benefit Khusus DPLK (Mode Checklist "✓")
  ["Benefit (DPLK)", "Pensiun Normal", "pensiun normal, normal", "Manfaat pensiun normal (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Ditunda", "pensiun ditunda, dana pensiun ditunda, ditunda", "Manfaat pensiun ditunda (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Cacat / Disabilitas", "cacat, disabilitas", "Manfaat pensiun karena cacat / disabilitas (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Dipercepat", "dipercepat", "Manfaat pensiun dipercepat (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Resign / Berhenti Kerja", "resign, berhenti, phk", "Manfaat saat resign / berhenti kerja (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Meninggal Dunia", "meninggal, janda, duda", "Manfaat pensiun meninggal dunia / janda / duda (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun WNI Pindah Kewarganegaraan", "wni, kewarganegaraan", "Manfaat bagi WNI yang pindah kewarganegaraan (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun WNA Kembali ke Negara Asal", "wna, negara asal, expatriate, ekspatriat", "Manfaat bagi WNA kembali ke negara asal (checklist DPLK)"],
  ["Benefit (DPLK)", "Pensiun Sakit Keras", "sakit keras, kritis", "Manfaat pensiun karena sakit keras (checklist DPLK)"],
  ["Benefit (DPLK)", "Pengalihan ke DPLK Lain", "pengalihan dplk", "Fasilitas pengalihan dana ke DPLK lain (checklist DPLK)"],
  ["Benefit (DPLK)", "Manfaat Lain (Dana Kompensasi Pasca Kerja)", "manfaat lain, kompensasi pasca kerja, dkp, asuransi tambahan, penarikan", "Manfaat tambahan / DKP pasca kerja (checklist DPLK)"],

  // 8. Daftar Komponen Biaya DPLK
  ["Biaya (DPLK)", "Biaya Pendaftaran", "pendaftaran, kepesertaan baru, registrasi", "Komponen biaya pendaftaran peserta baru DPLK"],
  ["Biaya (DPLK)", "Biaya Administrasi Kepesertaan", "administrasi, administration, biaya admin", "Komponen biaya administrasi kepesertaan DPLK"],
  ["Biaya (DPLK)", "Biaya Administrasi Iuran", "iuran, kontribusi, top up", "Komponen biaya administrasi atas iuran / kontribusi DPLK"],
  ["Biaya (DPLK)", "Biaya Pengelolaan Dana Peserta", "pengelolaan", "Komponen biaya pengelolaan dana peserta DPLK"],
  ["Biaya (DPLK)", "Biaya Perubahan Pilihan Investasi (Switching)", "perubahan, switching, pengalihan investasi, pengalihan paket", "Komponen biaya perubahan paket investasi (switching) DPLK"],
  ["Biaya (DPLK)", "Biaya Penarikan Sebagian Iuran", "penarikan", "Komponen biaya penarikan sebagian dana iuran DPLK"],
  ["Biaya (DPLK)", "Biaya Administrasi Pembayaran Manfaat Pensiun", "pembayaran manfaat, pencairan", "Komponen biaya administrasi pencairan manfaat pensiun DPLK"],
  ["Biaya (DPLK)", "Biaya Pengelolaan Investasi", "pengelolaan investasi, biaya investasi, management fee", "Komponen biaya pengelolaan investasi (management fee) DPLK"],
  ["Biaya (DPLK)", "Biaya Pengalihan Dana ke DPLK Lain", "pengalihan, transfer ke dplk", "Komponen biaya transfer kepesertaan ke DPLK lain"],
  ["Biaya (DPLK)", "Biaya Pencetakan Laporan Rekening", "pencetakan, laporan", "Komponen biaya pencetakan laporan rekening DPLK"],
  ["Biaya (DPLK)", "Biaya Kartu Kepesertaan", "kartu, buku", "Komponen biaya pembuatan kartu / buku kepesertaan DPLK"],
  ["Biaya (DPLK)", "Biaya Penghentian Kerja Sama / Pemutusan", "penghentian, pemutusan, penutupan, terminasi", "Komponen biaya penutupan / pemutusan kepesertaan DPLK"],
  ["Biaya (DPLK)", "Biaya Lain (Pihak Ketiga)", "pihak ketiga, kustodian, pajak, materai, biaya lain", "Komponen biaya pihak ketiga (kustodian, pajak, materai, dll.)"],

  // 9. Pengelompokan Filter Distribution Channel
  ["Filter Channel", "Bancassurance", "banca, bancassurance", "Kelompok filter jalur distribusi Bank (Bancassurance)"],
  ["Filter Channel", "Agency", "agency, agen", "Kelompok filter jalur distribusi Keagenan (Agency)"],
  ["Filter Channel", "Direct", "direct, digital", "Kelompok filter jalur distribusi Langsung / Digital"],
  ["Filter Channel", "Financial Advisor", "financial advisor, fa", "Kelompok filter jalur distribusi Financial Advisor"],
  ["Filter Channel", "Telemarketing", "telemarketing, tele", "Kelompok filter jalur distribusi Telemarketing"],
  ["Filter Channel", "Corporate", "corporate, group, kumpulan, employee benefit", "Kelompok filter jalur distribusi Korporasi / Kumpulan"],

  // 10. Pengelompokan Filter Underwriting Method
  ["Filter Underwriting", "GIO", "gio, guaranteed issue offer, guaranteed issuance", "Kelompok filter Underwriting Guaranteed Issue Offer (GIO)"],
  ["Filter Underwriting", "SIO", "sio, simplified issue offer, simplified issuance", "Kelompok filter Underwriting Simplified Issue Offer (SIO)"],
  ["Filter Underwriting", "Full Underwriting", "full underwriting, fuw, full uw", "Kelompok filter Full Underwriting"],
  ["Filter Underwriting", "NMU / Non-Medical", "nmu, non-medical, non medical", "Kelompok filter Non-Medical Underwriting (NMU)"],
  ["Filter Underwriting", "Medical Underwriting", "medical", "Kelompok filter Medical Underwriting"],
  ["Filter Underwriting", "Clean Case", "clean case", "Kelompok filter Clean Case"],

  // 11. Kata Kunci Deteksi Otomatis Nama Perusahaan dari Nama Produk saat Upload Excel
  ["Company Keyword", "Astra Life", "ava, astra, asli, asuransi jiwa astra", "Kata kunci deteksi otomatis perusahaan Astra Life dari nama produk"],
  ["Company Keyword", "Manulife", "manulife", "Kata kunci deteksi otomatis perusahaan Manulife dari nama produk"],
  ["Company Keyword", "AIA", "aia", "Kata kunci deteksi otomatis perusahaan AIA dari nama produk"],
  ["Company Keyword", "Allianz", "allianz, allisya", "Kata kunci deteksi otomatis perusahaan Allianz dari nama produk"],
  ["Company Keyword", "Prudential", "pru, prudential", "Kata kunci deteksi otomatis perusahaan Prudential dari nama produk"],
  ["Company Keyword", "AXA", "mandiri, axa, maestro", "Kata kunci deteksi otomatis perusahaan AXA / AXA Mandiri dari nama produk"],
  ["Company Keyword", "FWD", "fwd", "Kata kunci deteksi otomatis perusahaan FWD dari nama produk"],
  ["Company Keyword", "Great Eastern", "great, great eastern", "Kata kunci deteksi otomatis perusahaan Great Eastern dari nama produk"],
  ["Company Keyword", "Generali", "generali, isalaam", "Kata kunci deteksi otomatis perusahaan Generali dari nama produk"],
  ["Company Keyword", "Sunlife", "sun, sunlife, salam, berkah", "Kata kunci deteksi otomatis perusahaan Sun Life dari nama produk"]
];

function getDefaultConfigDescription_(category, canonicalName) {
  const catLower = String(category || "").trim().toLowerCase();
  const nameKey = normalizeBenefitAliasKey_(canonicalName);
  for (let i = 0; i < DEFAULT_CONFIG_ROWS_.length; i++) {
    const dRow = DEFAULT_CONFIG_ROWS_[i];
    if (String(dRow[0]).toLowerCase() === catLower && normalizeBenefitAliasKey_(dRow[1]) === nameKey) {
      return dRow[3];
    }
  }
  if (catLower === "benefit") return "Manfaat produk (otomatis dihapus jika tidak ada produk yang menggunakan)";
  if (catLower === "benefit (dplk)") return "Manfaat pensiun DPLK (mode checklist)";
  if (catLower === "biaya (dplk)") return "Komponen biaya DPLK";
  if (catLower.indexOf("non-benefit") === 0) return "Ketentuan produk Non-Benefit (di-lock / tidak dihapus otomatis)";
  if (catLower === "core") return "Kolom identitas utama produk";
  if (catLower === "sheet" || catLower === "tab" || catLower === "kategori" || catLower === "category") return "Nama tab sheet kategori produk di Master Spreadsheet";
  if (catLower === "filter channel") return "Pengelompokan filter Distribution Channel";
  if (catLower === "filter underwriting") return "Pengelompokan filter Underwriting Method";
  if (catLower === "company keyword") return "Kata kunci deteksi nama perusahaan dari nama produk";
  if (catLower === "setting") return "Pengaturan sistem aplikasi";
  return "";
}

// Fungsi inisialisasi / pelengkap otomatis Sheet Config (4 Kolom) tanpa menimpa data/alias yang sudah diedit user
function ensureDefaultConfigSheet_(ss, forceSeedAll) {
  const targetSs = ss || SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = targetSs.getSheetByName("Config");
  if (!configSheet) {
    configSheet = targetSs.insertSheet("Config");
  }

  const lastRow = configSheet.getLastRow();
  const lastCol = configSheet.getLastColumn();

  if (lastRow === 0 || lastCol === 0) {
    const allRows = [CONFIG_HEADERS_4COL_].concat(DEFAULT_CONFIG_ROWS_);
    const range = configSheet.getRange(1, 1, allRows.length, 4);
    range.setNumberFormat("@");
    range.setValues(allRows);
    configSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
    return configSheet;
  }

  const readCols = Math.max(4, lastCol);
  const rawValues = configSheet.getRange(1, 1, lastRow, readCols).getDisplayValues();
  const h0 = String(rawValues[0][0] || "").trim();
  const h1 = String(rawValues[0][1] || "").trim();
  const h2 = String(rawValues[0][2] || "").trim();
  const h3 = String(rawValues[0][3] || "").trim();

  // Perbarui baris header ke 4 kolom bahasa Indonesia jika masih memakai format 3 kolom lama
  if (h0 !== CONFIG_HEADERS_4COL_[0] || h1 !== CONFIG_HEADERS_4COL_[1] || h2 !== CONFIG_HEADERS_4COL_[2] || h3 !== CONFIG_HEADERS_4COL_[3]) {
    configSheet.getRange(1, 1, 1, 4).setValues([CONFIG_HEADERS_4COL_]);
    configSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  }

  // Cek apakah kategori-kategori baru sudah pernah di-seed ke sheet Config
  const existingKeys = {};
  const existingCategories = {};
  let hasBlankDescriptions = false;

  for (let r = 1; r < rawValues.length; r++) {
    const catRaw = String(rawValues[r][0] || "").trim();
    const nameRaw = String(rawValues[r][1] || "").trim();
    const descRaw = String(rawValues[r][3] || "").trim();
    if (!catRaw || !nameRaw) continue;
    const catLower = catRaw.toLowerCase();
    const normCat = (catLower === "tab" || catLower === "kategori" || catLower === "category") ? "sheet" : catLower;
    existingCategories[normCat] = true;
    existingKeys[normCat + "||" + normalizeBenefitAliasKey_(nameRaw)] = r;
    if (!descRaw) hasBlankDescriptions = true;
  }

  const requiredCategories = ["core", "non-benefit", "non-benefit (syariah)", "non-benefit (dplk)", "benefit (dplk)", "biaya (dplk)", "filter channel", "filter underwriting", "company keyword"];
  let missingCategory = false;
  for (let c = 0; c < requiredCategories.length; c++) {
    if (!existingCategories[requiredCategories[c]]) {
      missingCategory = true;
      break;
    }
  }

  if (!forceSeedAll && !missingCategory && !hasBlankDescriptions) {
    return configSheet;
  }

  // Lengkapi kolom D (Keterangan) yang masih kosong pada baris yang sudah ada
  if (hasBlankDescriptions) {
    const descColValues = [];
    for (let r = 1; r < rawValues.length; r++) {
      const catRaw = String(rawValues[r][0] || "").trim();
      const nameRaw = String(rawValues[r][1] || "").trim();
      let descRaw = String(rawValues[r][3] || "").trim();
      if (!descRaw && catRaw && nameRaw) {
        descRaw = getDefaultConfigDescription_(catRaw, nameRaw);
      }
      descColValues.push([descRaw]);
    }
    if (descColValues.length > 0) {
      configSheet.getRange(2, 4, descColValues.length, 1).setValues(descColValues);
    }
  }

  // Tambahkan baris default yang belum ada (tanpa mengubah baris yang sudah ada)
  const rowsToAppend = [];
  for (let i = 0; i < DEFAULT_CONFIG_ROWS_.length; i++) {
    const defRow = DEFAULT_CONFIG_ROWS_[i];
    const defCatLower = String(defRow[0]).trim().toLowerCase();
    const defKey = defCatLower + "||" + normalizeBenefitAliasKey_(defRow[1]);
    if (existingKeys[defKey] === undefined) {
      if (forceSeedAll || !existingCategories[defCatLower] || defCatLower !== "benefit") {
        rowsToAppend.push(defRow);
        existingKeys[defKey] = -1;
      }
    }
  }

  if (rowsToAppend.length > 0) {
    const startRow = configSheet.getLastRow() + 1;
    const appendRange = configSheet.getRange(startRow, 1, rowsToAppend.length, 4);
    appendRange.setNumberFormat("@");
    appendRange.setValues(rowsToAppend);
  }

  return configSheet;
}

// Fungsi yang bisa dijalankan manual dari Editor Google Apps Script kapan saja untuk melengkapi/merapikan Sheet Config
function setupDefaultConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureDefaultConfigSheet_(ss, true);
  invalidateSummaryCache_();
  console.log("Sheet Config berhasil dilengkapi dengan struktur 4 kolom (Kategori | Nama Utama | Nilai / Nama Lain | Keterangan)!");
}

// ====================================================
// CACHE RINGKAS STARTUP (CacheService) AGAR LOADING AWAL JAUH LEBIH CEPAT
// ====================================================
const SUMMARY_CACHE_PREFIX_ = "IDCM_SUMMARY_V2_";
const SUMMARY_CACHE_TTL_SEC_ = 21600; // 6 jam (otomatis di-invalidate setiap ada perubahan data/config)

function getCachedSummaryResponseText_() {
  try {
    const cache = CacheService.getScriptCache();
    const metaRaw = cache.get(SUMMARY_CACHE_PREFIX_ + "META");
    if (!metaRaw) return null;
    const meta = JSON.parse(metaRaw);
    const chunks = Number(meta && meta.chunks) || 0;
    if (chunks <= 0) return null;
    const keys = [];
    for (let i = 0; i < chunks; i++) keys.push(SUMMARY_CACHE_PREFIX_ + "CHUNK_" + i);
    const vals = cache.getAll(keys);
    let joined = "";
    for (let i = 0; i < chunks; i++) {
      const part = vals[SUMMARY_CACHE_PREFIX_ + "CHUNK_" + i];
      if (!part) return null;
      joined += part;
    }
    if (meta && meta.encoding === "gzip_b64") {
      const gzBytes = Utilities.base64Decode(joined);
      const gzBlob = Utilities.newBlob(gzBytes, "application/x-gzip", "summary.json.gz");
      return Utilities.ungzip(gzBlob).getDataAsString("UTF-8");
    }
    return joined;
  } catch (e) {
    return null;
  }
}

function setCachedSummaryResponseText_(jsonText) {
  try {
    const text = String(jsonText || "");
    if (!text) return;
    const cache = CacheService.getScriptCache();
    const CHUNK_SIZE = 90000; // di bawah batas 100KB per entri CacheService
    let payloadStr = text;
    let encoding = "plain";
    try {
      const rawBlob = Utilities.newBlob(text, "application/json", "summary.json");
      const gzBlob = Utilities.gzip(rawBlob);
      const b64 = Utilities.base64Encode(gzBlob.getBytes());
      if (b64 && b64.length < text.length) {
        payloadStr = b64;
        encoding = "gzip_b64";
      }
    } catch (gzErr) {
      console.warn("Kompresi gzip summary cache dilewati:", gzErr);
    }
    const chunks = Math.ceil(payloadStr.length / CHUNK_SIZE);
    if (chunks > 10) return; // tetap patuhi batas aman 10 chunk CacheService
    const entries = {};
    for (let i = 0; i < chunks; i++) {
      entries[SUMMARY_CACHE_PREFIX_ + "CHUNK_" + i] = payloadStr.substring(i * CHUNK_SIZE, (i + 1) * CHUNK_SIZE);
    }
    entries[SUMMARY_CACHE_PREFIX_ + "META"] = JSON.stringify({ chunks: chunks, encoding: encoding, updatedAt: Date.now() });
    cache.putAll(entries, SUMMARY_CACHE_TTL_SEC_);
  } catch (e) {
    console.warn("Gagal menyimpan summary cache:", e);
  }
}

function invalidateSummaryCache_() {
  try {
    const cache = CacheService.getScriptCache();
    const keys = [SUMMARY_CACHE_PREFIX_ + "META"];
    for (let i = 0; i < 10; i++) keys.push(SUMMARY_CACHE_PREFIX_ + "CHUNK_" + i);
    cache.removeAll(keys);
  } catch (e) {}
}

// Helper: Baca cepat sheet Config tanpa memicu write ke spreadsheet kecuali benar-benar belum lengkap
function getFastConfigDisplayValues_(ss) {
  let cfgValues = null;
  if (typeof Sheets !== "undefined" && Sheets.Spreadsheets && Sheets.Spreadsheets.Values && typeof Sheets.Spreadsheets.Values.get === "function") {
    try {
      const res = Sheets.Spreadsheets.Values.get(ss.getId(), "'Config'!A:D", { valueRenderOption: "FORMATTED_VALUE" });
      if (res && Array.isArray(res.values) && res.values.length > 0) {
        cfgValues = res.values.map(function(r) {
          return [
            r[0] !== undefined ? String(r[0]) : "",
            r[1] !== undefined ? String(r[1]) : "",
            r[2] !== undefined ? String(r[2]) : "",
            r[3] !== undefined ? String(r[3]) : ""
          ];
        });
      }
    } catch (e) {}
  }

  let needsSetup = !cfgValues || cfgValues.length === 0;
  if (!needsSetup) {
    const h = cfgValues[0];
    const has4ColHeader = String(h[0] || "").trim() === CONFIG_HEADERS_4COL_[0] &&
      String(h[1] || "").trim() === CONFIG_HEADERS_4COL_[1] &&
      String(h[2] || "").trim() === CONFIG_HEADERS_4COL_[2] &&
      String(h[3] || "").trim() === CONFIG_HEADERS_4COL_[3];
    if (!has4ColHeader) {
      needsSetup = true;
    } else {
      let hasCore = false;
      let hasCompanyKw = false;
      for (let i = 1; i < cfgValues.length; i++) {
        const t = String(cfgValues[i][0] || "").trim().toLowerCase();
        if (t === "core" || t === "identitas") hasCore = true;
        if (t === "company keyword" || t === "company") hasCompanyKw = true;
        if (hasCore && hasCompanyKw) break;
      }
      if (!hasCore || !hasCompanyKw) needsSetup = true;
    }
  }

  if (needsSetup) {
    const configSheet = ensureDefaultConfigSheet_(ss, false);
    if (configSheet && configSheet.getLastRow() > 0 && configSheet.getLastColumn() > 0) {
      cfgValues = configSheet.getDataRange().getDisplayValues();
    }
  }
  return cfgValues || [];
}

function parseProductSectionExtra_(rawExtra) {
  const txt = String(rawExtra || "").trim();
  if (!txt) return { mode: "standard", items: [] };
  const pipeIdx = txt.indexOf("|");
  let modePart = txt;
  let itemsPart = "";
  if (pipeIdx !== -1) {
    modePart = txt.substring(0, pipeIdx).trim();
    itemsPart = txt.substring(pipeIdx + 1).trim();
  }
  const modeMatch = modePart.match(/mode\s*=\s*(checklist|standard)/i);
  const mode = modeMatch ? modeMatch[1].toLowerCase() : (/checklist/i.test(modePart) ? "checklist" : "standard");
  const items = itemsPart
    ? splitBenefitAliasList_(itemsPart)
    : (!modeMatch && txt.indexOf("=") === -1 ? splitBenefitAliasList_(txt) : []);
  return { mode: mode, items: items };
}

// Helper Global: Ambil seluruh konfigurasi aman & data produk dari Master Spreadsheet
// (Memungkinkan Google Spreadsheet dikunci 'Restricted / Private' tanpa perlu dibuka ke publik via /gviz/tq)
// Parameter mode: "summary" (ringkas untuk loading awal cepat) atau "full" (default, lengkap).
function buildAllProductDataResponse(mode, bypassCache) {
  const isSummary = String(mode || "").trim().toLowerCase() === "summary";
  if (isSummary && !bypassCache) {
    const cachedText = getCachedSummaryResponseText_();
    if (cachedText) {
      return { __rawCachedJsonText: cachedText };
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const DEFAULT_TABS = ["Conventional", "Sharia", "Health & Surgical", "DPLK"];
  const configuredTabs = [];
  const sheetAliases = [];
  const coreAliases = [];
  const nonBenefitColumns = [];
  const syariahNonBenefitColumns = [];
  const dplkNonBenefitColumns = [];
  const nonBenefitAliases = [];
  const benefitAliases = [];
  const dplkBenefits = [];
  const dplkCharges = [];
  const channelFilters = [];
  const underwritingFilters = [];
  const companyKeywords = [];
  const categoryConfigs = {};
  const ensureCatCfg = function(catName) {
    const key = String(catName || "").trim();
    if (!key) return null;
    if (!categoryConfigs[key]) {
      categoryConfigs[key] = {
        generalProvisions: null,
        specialProvisions: [],
        productSections: []
      };
    }
    return categoryConfigs[key];
  };
  let driveFolderId = "";

  const cfgValues = getFastConfigDisplayValues_(ss);
  for (let i = 1; i < cfgValues.length; i++) {
    const row = cfgValues[i];
    const type = String(row[0] || "").trim().toLowerCase();
    const val = String(row[1] || "").trim();
    const extra = String(row[2] || "").trim();
    const desc = String(row[3] || "").trim();
    if (!type || !val) continue;
    const aliasesList = splitBenefitAliasList_(extra);

    if (type === "sheet" || type === "tab" || type === "kategori" || type === "category") {
      const lowerVal = val.toLowerCase();
      if (
        lowerVal !== "config" &&
        lowerVal !== "documents" &&
        lowerVal !== "users" &&
        !configuredTabs.some(function(t) { return t.toLowerCase() === lowerVal; })
      ) {
        configuredTabs.push(val);
        sheetAliases.push({ canonical: val, aliases: aliasesList, description: desc });
      }
    } else if (type === "core" || type === "identitas") {
      coreAliases.push({ canonical: val, aliases: aliasesList, description: desc });
    } else if (type === "category provision" || type === "category provisions") {
      const catEntry = ensureCatCfg(val);
      if (catEntry) {
        catEntry.generalProvisions = aliasesList;
      }
    } else if (type.indexOf("special provision") === 0) {
      const spMatch = type.match(/^special\s+provisions?\s*\((.+)\)$/i);
      const targetCatName = spMatch ? spMatch[1].trim() : "";
      if (targetCatName) {
        const matchedCat = configuredTabs.find(function(t) { return t.toLowerCase() === targetCatName.toLowerCase(); }) || targetCatName;
        const catEntry = ensureCatCfg(matchedCat);
        const allowImg = /\[image\]/i.test(desc);
        const cleanDesc = desc.replace(/\[image\]/gi, "").trim();
        if (catEntry) {
          catEntry.specialProvisions.push({
            name: val,
            aliases: aliasesList,
            allowImage: allowImg,
            description: cleanDesc
          });
        }
      }
      if (!nonBenefitColumns.some(function(c) { return c.toLowerCase() === val.toLowerCase(); })) {
        nonBenefitColumns.push(val);
      }
      nonBenefitAliases.push({ canonical: val, aliases: aliasesList, scope: "special", description: desc });
    } else if (type.indexOf("product section") === 0) {
      const secMatch = type.match(/^product\s+sections?\s*\((.+)\)$/i);
      const targetCatName = secMatch ? secMatch[1].trim() : "";
      if (targetCatName) {
        const matchedCat = configuredTabs.find(function(t) { return t.toLowerCase() === targetCatName.toLowerCase(); }) || targetCatName;
        const catEntry = ensureCatCfg(matchedCat);
        const parsedSec = parseProductSectionExtra_(extra);
        if (catEntry) {
          catEntry.productSections.push({
            title: val,
            mode: parsedSec.mode,
            items: parsedSec.items,
            description: desc
          });
        }
      }
    } else if (type === "non-benefit" || type === "non-benefit (syariah)" || type === "non-benefit (dplk)") {
      if (!nonBenefitColumns.some(function(c) { return c.toLowerCase() === val.toLowerCase(); })) {
        nonBenefitColumns.push(val);
      }
      let scope = "general";
      if (type === "non-benefit (syariah)") {
        scope = "syariah";
        if (!syariahNonBenefitColumns.some(function(c) { return c.toLowerCase() === val.toLowerCase(); })) {
          syariahNonBenefitColumns.push(val);
        }
      } else if (type === "non-benefit (dplk)") {
        scope = "dplk";
        if (!dplkNonBenefitColumns.some(function(c) { return c.toLowerCase() === val.toLowerCase(); })) {
          dplkNonBenefitColumns.push(val);
        }
      }
      nonBenefitAliases.push({
        canonical: val,
        aliases: aliasesList,
        scope: scope,
        description: desc
      });
    } else if (type === "benefit") {
      benefitAliases.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "benefit (dplk)" || type === "dplk benefit") {
      dplkBenefits.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "biaya (dplk)" || type === "dplk charge" || type === "charge (dplk)") {
      dplkCharges.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "filter channel" || type === "channel") {
      channelFilters.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "filter underwriting" || type === "underwriting") {
      underwritingFilters.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "company keyword" || type === "company") {
      companyKeywords.push({
        canonical: val,
        aliases: aliasesList,
        description: desc
      });
    } else if (type === "setting" && val.toLowerCase() === "drive folder id" && extra) {
      driveFolderId = extra;
    }
    // Catatan Keamanan: Baris 'Invite Code' TIDAK dimasukkan ke response agar tetap rahasia di server
  }

  const tabsToLoad = configuredTabs.length > 0 ? configuredTabs : DEFAULT_TABS;
  const batchData = readSheetsBatchOrFallback_(ss, tabsToLoad);
  const sheetsData = [];

  tabsToLoad.forEach(function(tabName) {
    const sheetResult = batchData[tabName];
    const displayVals = (sheetResult && Array.isArray(sheetResult.values)) ? sheetResult.values : [];
    if (!sheetResult || !sheetResult.exists || displayVals.length === 0) {
      sheetsData.push({
        sheetName: tabName,
        exists: Boolean(sheetResult && sheetResult.exists),
        headers: [],
        rows: []
      });
      return;
    }

    const headers = (displayVals[0] || []).map(function(h) { return String(h || "").trim(); });
    if (headers.length === 0 || headers.every(function(h) { return h === ""; })) {
      sheetsData.push({
        sheetName: tabName,
        exists: true,
        headers: [],
        rows: []
      });
      return;
    }

    const fullRows = displayVals.slice(1);
    const outRows = isSummary
      ? buildSummaryRowsForSheet_(tabName, headers, fullRows, nonBenefitColumns)
      : fullRows;

    sheetsData.push({
      sheetName: tabName,
      exists: true,
      headers: headers,
      rows: outRows
    });
  });

  const resultObj = {
    status: "success",
    action: "get_all_product_data",
    mode: isSummary ? "summary" : "full",
    config: {
      tabs: tabsToLoad,
      sheetAliases: sheetAliases,
      coreAliases: coreAliases,
      nonBenefitColumns: nonBenefitColumns,
      syariahNonBenefitColumns: syariahNonBenefitColumns,
      dplkNonBenefitColumns: dplkNonBenefitColumns,
      nonBenefitAliases: nonBenefitAliases,
      benefitAliases: benefitAliases,
      dplkBenefits: dplkBenefits,
      dplkCharges: dplkCharges,
      channelFilters: channelFilters,
      underwritingFilters: underwritingFilters,
      companyKeywords: companyKeywords,
      categoryConfigs: categoryConfigs,
      driveFolderId: driveFolderId
    },
    sheets: sheetsData
  };

  if (isSummary) {
    setCachedSummaryResponseText_(JSON.stringify(resultObj));
  }

  return resultObj;
}

// Simpan konfigurasi Sheet / Kategori beserta Provision, Special Provision, & Product Sections ke Master Spreadsheet & Config
function saveSheetCategoryConfig_(ss, payload) {
  const sheetName = String(payload.sheetName || "").trim();
  if (!sheetName) {
    throw new Error("Nama sheet / kategori tidak boleh kosong.");
  }
  const lowerSheetName = sheetName.toLowerCase();
  if (lowerSheetName === "config" || lowerSheetName === "users" || lowerSheetName === "documents") {
    throw new Error("Nama sheet tidak diizinkan (reserved system sheet): " + sheetName);
  }

  const aliases = Array.isArray(payload.aliases) ? payload.aliases.map(function(s) { return String(s || "").trim(); }).filter(Boolean) : [];
  const generalProvisions = Array.isArray(payload.generalProvisions) ? payload.generalProvisions.map(function(s) { return String(s || "").trim(); }).filter(Boolean) : [];
  const specialProvisions = Array.isArray(payload.specialProvisions)
    ? payload.specialProvisions
        .map(function(sp) {
          return {
            name: String((sp && sp.name) || "").trim(),
            aliases: Array.isArray(sp && sp.aliases) ? sp.aliases.map(function(a) { return String(a || "").trim(); }).filter(Boolean) : [],
            allowImage: Boolean(sp && sp.allowImage),
            description: String((sp && sp.description) || "").trim()
          };
        })
        .filter(function(sp) { return sp.name !== ""; })
    : [];
  const productSections = Array.isArray(payload.productSections)
    ? payload.productSections
        .map(function(sec) {
          const title = String((sec && sec.title) || "").trim();
          const mode = String((sec && sec.mode) || "standard").trim().toLowerCase() === "checklist" ? "checklist" : "standard";
          const items = Array.isArray(sec && sec.items) ? sec.items.map(function(x) { return String(x || "").trim(); }).filter(Boolean) : [];
          return { title: title, mode: mode, items: items };
        })
        .filter(function(sec) { return sec.title !== ""; })
    : [];

  // 1. Pastikan tab sheet produk tersedia dan memiliki header kolom yang diperlukan (tanpa menghapus data existing)
  let targetSheet = ss.getSheetByName(sheetName);
  if (!targetSheet) {
    const existingByLower = ss.getSheets().find(function(s) { return String(s.getName() || "").trim().toLowerCase() === lowerSheetName; });
    targetSheet = existingByLower || ss.insertSheet(sheetName);
  }

  const requiredHeaders = [
    "Product Name", "Company Name", "Par/Non Par", "Product Type",
    "Distribution Channel", "Currency", "Launch Date"
  ];
  generalProvisions.forEach(function(gp) {
    if (!requiredHeaders.some(function(h) { return h.toLowerCase() === gp.toLowerCase(); })) {
      requiredHeaders.push(gp);
    }
  });
  specialProvisions.forEach(function(sp) {
    if (!requiredHeaders.some(function(h) { return h.toLowerCase() === sp.name.toLowerCase(); })) {
      requiredHeaders.push(sp.name);
    }
  });
  if (!requiredHeaders.some(function(h) { return h.toLowerCase() === "additional benefit"; })) {
    requiredHeaders.push("Additional Benefit");
  }

  const lastCol = targetSheet.getLastColumn();
  if (lastCol === 0 || targetSheet.getLastRow() === 0) {
    const hdrRange = targetSheet.getRange(1, 1, 1, requiredHeaders.length);
    hdrRange.setNumberFormat("@");
    hdrRange.setValues([sanitizeSheetRow_(requiredHeaders)]);
    hdrRange.setFontWeight("bold");
  } else {
    const existingHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || "").trim(); });
    const existingLower = existingHeaders.map(function(h) { return h.toLowerCase(); });
    const colsToAdd = [];
    requiredHeaders.forEach(function(h) {
      if (existingLower.indexOf(h.toLowerCase()) === -1) {
        colsToAdd.push(h);
        existingLower.push(h.toLowerCase());
      }
    });
    if (colsToAdd.length > 0) {
      const startCol = existingHeaders.length + 1;
      const addHdrRange = targetSheet.getRange(1, startCol, 1, colsToAdd.length);
      addHdrRange.setNumberFormat("@");
      addHdrRange.setValues([sanitizeSheetRow_(colsToAdd)]);
      addHdrRange.setFontWeight("bold");
    }
  }

  // 2. Perbarui sheet Config (4 Kolom) secara rapi tanpa menghilangkan baris pengaturan lain
  const configSheet = ensureDefaultConfigSheet_(ss, false);
  const cfgLastRow = configSheet.getLastRow();
  const cfgCols = Math.max(4, configSheet.getLastColumn());
  const cfgData = cfgLastRow > 0 ? configSheet.getRange(1, 1, cfgLastRow, cfgCols).getDisplayValues() : [CONFIG_HEADERS_4COL_];

  const spCatLabel = "Special Provision (" + sheetName + ")";
  const spCatLower = spCatLabel.toLowerCase();
  const secCatLabel = "Product Section (" + sheetName + ")";
  const secCatLower = secCatLabel.toLowerCase();
  const isShariaAlias = /^(sharia|syariah)$/i.test(sheetName);
  const isDplkSheet = /^dplk$/i.test(sheetName);

  const nonBenExistingKeys = {};
  const retainedRows = [CONFIG_HEADERS_4COL_];
  let sheetRowUpdated = false;
  let catProvRowUpdated = false;

  for (let r = 1; r < cfgData.length; r++) {
    const row = cfgData[r];
    const c0 = String(row[0] || "").trim();
    const c0L = c0.toLowerCase();
    const c1 = String(row[1] || "").trim();
    const c1L = c1.toLowerCase();
    const c2 = String(row[2] || "").trim();
    const c3 = String(row[3] || "").trim();

    if (!c0 && !c1 && !c2 && !c3) continue;

    // Hapus baris Special Provision / Product Section lama milik kategori ini (akan digantikan yang terbaru)
    if (c0L === spCatLower || c0L === secCatLower) {
      continue;
    }
    if (isShariaAlias && c0L === "non-benefit (syariah)") continue;
    if (isDplkSheet && c0L === "non-benefit (dplk)") continue;

    if ((c0L === "sheet" || c0L === "tab" || c0L === "kategori" || c0L === "category") && c1L === lowerSheetName) {
      retainedRows.push(["Sheet", sheetName, aliases.join(", "), c3 || ("Nama tab sheet produk " + sheetName + " di Master Spreadsheet")]);
      sheetRowUpdated = true;
      continue;
    }

    if ((c0L === "category provision" || c0L === "category provisions") && c1L === lowerSheetName) {
      retainedRows.push(["Category Provision", sheetName, generalProvisions.join(", "), "Daftar General Provision aktif untuk kategori " + sheetName]);
      catProvRowUpdated = true;
      continue;
    }

    if (c0L === "non-benefit" && c1) {
      nonBenExistingKeys[normalizeBenefitAliasKey_(c1)] = true;
    }

    retainedRows.push([c0, c1, c2, c3]);
  }

  if (!sheetRowUpdated) {
    retainedRows.push(["Sheet", sheetName, aliases.join(", "), "Nama tab sheet produk " + sheetName + " di Master Spreadsheet"]);
  }
  if (!catProvRowUpdated) {
    retainedRows.push(["Category Provision", sheetName, generalProvisions.join(", "), "Daftar General Provision aktif untuk kategori " + sheetName]);
  }

  specialProvisions.forEach(function(sp) {
    const descBase = sp.description || ("Special Provision khusus kategori " + sheetName + " (di-lock / tidak dihapus otomatis)");
    const finalDesc = sp.allowImage ? (descBase + " [Image]") : descBase;
    retainedRows.push([spCatLabel, sp.name, sp.aliases.join(", "), finalDesc]);
  });

  productSections.forEach(function(sec) {
    const itemsStr = sec.items.join(", ");
    const valCol = itemsStr ? ("mode=" + sec.mode + " | " + itemsStr) : ("mode=" + sec.mode);
    const modeLabel = sec.mode === "checklist" ? "Checklist + Catatan" : "Teks & Gambar";
    retainedRows.push([secCatLabel, sec.title, valCol, "Section " + sec.title + " pada " + sheetName + " (" + modeLabel + ")"]);
  });

  // Daftarkan General Provision kustom baru ke kategori "Non-Benefit" agar tetap diproteksi
  const builtInCoreOrGen = {
    "product name": true, "company name": true, "par/non par": true, "product type": true,
    "distribution channel": true, "currency": true, "launch date": true, "tanggal rilis": true,
    "entry age": true, "coverage period": true, "premium payment period (ppp)": true,
    "min premium": true, "min sum assured": true, "underwriting method": true,
    "available riders": true, "additional benefit": true, "remarks": true
  };
  generalProvisions.forEach(function(gp) {
    const gpKey = normalizeBenefitAliasKey_(gp);
    if (!builtInCoreOrGen[gp.toLowerCase()] && !nonBenExistingKeys[gpKey]) {
      retainedRows.push(["Non-Benefit", gp, "", getDefaultConfigDescription_("Non-Benefit", gp)]);
      nonBenExistingKeys[gpKey] = true;
    }
  });

  const writeRange = configSheet.getRange(1, 1, retainedRows.length, 4);
  writeRange.setNumberFormat("@");
  writeRange.setValues(sanitizeSheetMatrix_(retainedRows));
  configSheet.getRange(1, 1, 1, 4).setFontWeight("bold");
  if (cfgLastRow > retainedRows.length) {
    configSheet.getRange(retainedRows.length + 1, 1, cfgLastRow - retainedRows.length, cfgCols).clearContent();
  }

  invalidateSummaryCache_();
  return buildAllProductDataResponse("summary", true).config;
}

// ====================================================
// HELPER: ALIAS NAMA MANFAAT & NON-BENEFIT (Config 4 Kolom: Kategori | Nama Utama | Nilai / Nama Lain | Keterangan)
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

// Simpan pemetaan alias -> nama standar ke sheet Config (mendukung tipe "Benefit" dan "Non-Benefit").
// Satu alias pada tipe yang sama hanya boleh menunjuk ke satu nama standar; baris baru juga dapat didaftarkan tanpa alias.
function saveBenefitAliases_(ss, mappings) {
  const configSheet = ensureDefaultConfigSheet_(ss, false);
  const lastRow = configSheet.getLastRow();
  const lastCol = Math.max(4, configSheet.getLastColumn());
  const data = lastRow > 0 ? configSheet.getRange(1, 1, lastRow, lastCol).getValues() : [];
  const configRowsByType = { "benefit": [], "non-benefit": [] };

  for (let i = 1; i < data.length; i++) {
    const rawType = String(data[i][0] || "").trim();
    const typeLower = rawType.toLowerCase();
    const canonical = String(data[i][1] || "").trim();
    const desc = String(data[i][3] || "").trim();
    const isNbType = typeLower === "non-benefit" || typeLower === "non-benefit (syariah)" || typeLower === "non-benefit (dplk)";
    const isBenType = typeLower === "benefit" || typeLower === "benefit (dplk)";
    if ((isBenType || isNbType) && canonical) {
      const groupKey = isNbType ? "non-benefit" : "benefit";
      configRowsByType[groupKey].push({
        rowNum: i + 1,
        entryType: rawType || (isNbType ? "Non-Benefit" : "Benefit"),
        canonical: canonical,
        aliases: splitBenefitAliasList_(data[i][2]),
        description: desc,
        dirty: false
      });
    }
  }

  let added = 0;
  (mappings || []).forEach(function(m) {
    if (!m) return;
    const isNonBenefit = String(m.entryType || "").trim().toLowerCase().indexOf("non-benefit") === 0;
    const typeKey = isNonBenefit ? "non-benefit" : "benefit";
    const entryTypeLabel = isNonBenefit ? "Non-Benefit" : "Benefit";
    const typeRows = configRowsByType[typeKey];

    const canonical = String(m.canonical || "").trim();
    const alias = String(m.alias || "").trim().replace(/[:：]\s*$/, "").trim();
    if (!canonical) return;
    const canonicalKey = normalizeBenefitAliasKey_(canonical);
    const aliasKey = alias ? normalizeBenefitAliasKey_(alias) : "";

    let target = typeRows.find(function(r) { return normalizeBenefitAliasKey_(r.canonical) === canonicalKey; });
    if (!target) {
      target = {
        rowNum: null,
        entryType: entryTypeLabel,
        canonical: canonical,
        aliases: [],
        description: getDefaultConfigDescription_(entryTypeLabel, canonical),
        dirty: true
      };
      typeRows.push(target);
      added++;
    }

    if (!aliasKey || aliasKey === canonicalKey) return;

    // Hapus alias yang sama dari nama standar lain dalam tipe yang sama agar tidak bentrok
    typeRows.forEach(function(r) {
      if (normalizeBenefitAliasKey_(r.canonical) === canonicalKey) return;
      const before = r.aliases.length;
      r.aliases = r.aliases.filter(function(a) { return normalizeBenefitAliasKey_(a) !== aliasKey; });
      if (r.aliases.length !== before) r.dirty = true;
    });

    if (!target.aliases.some(function(a) { return normalizeBenefitAliasKey_(a) === aliasKey; })) {
      target.aliases.push(alias);
      target.dirty = true;
      added++;
    }
  });

  ["benefit", "non-benefit"].forEach(function(typeKey) {
    configRowsByType[typeKey].forEach(function(r) {
      if (!r.dirty) return;
      const aliasStr = r.aliases.join(", ");
      let rowNum = r.rowNum;
      if (!rowNum) {
        rowNum = configSheet.getLastRow() + 1;
        const rowRange = configSheet.getRange(rowNum, 1, 1, 4);
        rowRange.setNumberFormat("@");
        rowRange.setValues([sanitizeSheetRow_([r.entryType, r.canonical, aliasStr, r.description || getDefaultConfigDescription_(r.entryType, r.canonical)])]);
        return;
      }
      const cell = configSheet.getRange(rowNum, 3);
      cell.setNumberFormat("@");
      cell.setValue(sanitizeSheetCell_(aliasStr));
      if (!r.description) {
        const descCell = configSheet.getRange(rowNum, 4);
        descCell.setNumberFormat("@");
        descCell.setValue(sanitizeSheetCell_(getDefaultConfigDescription_(r.entryType, r.canonical)));
      }
    });
  });

  return added;
}

// Kelola (List / Upsert / Delete) entri Alias Benefit & Non-Benefit di sheet Config secara aman dari frontend
function manageBenefitAliases_(ss, payload) {
  const configSheet = ensureDefaultConfigSheet_(ss, false);
  const op = String((payload && payload.op) || "list").trim().toLowerCase();

  function readEntries_() {
    const lastRow = configSheet.getLastRow();
    const lastCol = Math.max(4, configSheet.getLastColumn());
    const data = lastRow > 0 ? configSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues() : [];
    const entries = [];
    for (let i = 1; i < data.length; i++) {
      const rawType = String(data[i][0] || "").trim();
      const typeLower = rawType.toLowerCase();
      const canonical = String(data[i][1] || "").trim();
      const isNbType = typeLower === "non-benefit" || typeLower === "non-benefit (syariah)" || typeLower === "non-benefit (dplk)";
      const isBenType = typeLower === "benefit" || typeLower === "benefit (dplk)";
      if ((isBenType || isNbType) && canonical) {
        entries.push({
          rowNum: i + 1,
          entryType: rawType,
          groupKey: isNbType ? "non-benefit" : "benefit",
          canonical: canonical,
          aliases: splitBenefitAliasList_(data[i][2]),
          description: String(data[i][3] || "").trim()
        });
      }
    }
    return entries;
  }

  if (op === "list") {
    return { entries: readEntries_() };
  }

  if (op === "upsert") {
    const rawType = String(payload.entryType || "Benefit").trim();
    const typeLower = rawType.toLowerCase();
    const isNb = typeLower.indexOf("non-benefit") === 0;
    const safeType = isNb
      ? (typeLower === "non-benefit (syariah)" ? "Non-Benefit (Syariah)" : (typeLower === "non-benefit (dplk)" ? "Non-Benefit (DPLK)" : "Non-Benefit"))
      : (typeLower === "benefit (dplk)" ? "Benefit (DPLK)" : "Benefit");
    const canonical = String(payload.canonical || "").trim();
    if (!canonical) {
      throw new Error("Nama utama (canonical) tidak boleh kosong.");
    }
    const rawAliases = Array.isArray(payload.aliases)
      ? payload.aliases.join(", ")
      : String(payload.aliases || "");
    const cleanAliases = [];
    const seenAlias = {};
    const canonicalKey = normalizeBenefitAliasKey_(canonical);
    splitBenefitAliasList_(rawAliases).forEach(function(a) {
      const aClean = a.replace(/[:：]\s*$/, "").trim();
      const aKey = normalizeBenefitAliasKey_(aClean);
      if (aKey && aKey !== canonicalKey && !seenAlias[aKey]) {
        seenAlias[aKey] = true;
        cleanAliases.push(aClean);
      }
    });
    const aliasStr = cleanAliases.join(", ");
    const desc = String(payload.description || "").trim() || getDefaultConfigDescription_(safeType, canonical);
    const reqRowNum = parseInt(payload.rowNum, 10) || 0;

    const entries = readEntries_();
    let targetRowNum = 0;
    if (reqRowNum >= 2) {
      const matchByRow = entries.find(function(e) { return e.rowNum === reqRowNum; });
      if (matchByRow) targetRowNum = matchByRow.rowNum;
    }
    if (!targetRowNum) {
      const groupKey = isNb ? "non-benefit" : "benefit";
      const matchByKey = entries.find(function(e) {
        return e.groupKey === groupKey && normalizeBenefitAliasKey_(e.canonical) === canonicalKey;
      });
      if (matchByKey) targetRowNum = matchByKey.rowNum;
    }

    if (targetRowNum >= 2) {
      const r = configSheet.getRange(targetRowNum, 1, 1, 4);
      r.setNumberFormat("@");
      r.setValues([sanitizeSheetRow_([safeType, canonical, aliasStr, desc])]);
    } else {
      const newRowNum = configSheet.getLastRow() + 1;
      const r = configSheet.getRange(newRowNum, 1, 1, 4);
      r.setNumberFormat("@");
      r.setValues([sanitizeSheetRow_([safeType, canonical, aliasStr, desc])]);
    }

    invalidateSummaryCache_();
    return {
      entries: readEntries_(),
      config: buildAllProductDataResponse("summary", true).config
    };
  }

  if (op === "delete") {
    const reqRowNum = parseInt(payload.rowNum, 10) || 0;
    const reqCanonical = String(payload.canonical || "").trim().toLowerCase();
    const entries = readEntries_();
    const target = entries.find(function(e) {
      return e.rowNum === reqRowNum && (!reqCanonical || e.canonical.toLowerCase() === reqCanonical);
    });
    if (!target) {
      throw new Error("Entri alias tidak ditemukan atau bukan baris Benefit/Non-Benefit.");
    }
    configSheet.deleteRow(target.rowNum);
    invalidateSummaryCache_();
    return {
      entries: readEntries_(),
      config: buildAllProductDataResponse("summary", true).config
    };
  }

  throw new Error("Operasi manage_benefit_aliases tidak valid.");
}

// Helper: Ambil seluruh daftar dokumen produk dengan Drive Advanced Service (mendukung v2 & v3),
// atau fallback otomatis ke DriveApp jika Drive API belum diaktifkan.
function listAllProductDocumentsFast_() {
  const rootFolder = getOrCreateRootDocFolder();
  const rootId = rootFolder.getId();
  const tz = Session.getScriptTimeZone() || "GMT+7";

  if (typeof Drive !== "undefined" && Drive.Files && typeof Drive.Files.list === "function") {
    try {
      let detectedVersion = null; // "v3" | "v2"
      const driveListAll_ = function(q, fieldsV3, fieldsV2) {
        const collected = [];
        let pageToken = null;
        do {
          let res = null;
          if (detectedVersion === "v2") {
            const opts2 = { q: q, maxResults: 1000, fields: fieldsV2 };
            if (pageToken) opts2.pageToken = pageToken;
            res = Drive.Files.list(opts2);
          } else if (detectedVersion === "v3") {
            const opts3 = { q: q, pageSize: 1000, fields: fieldsV3 };
            if (pageToken) opts3.pageToken = pageToken;
            res = Drive.Files.list(opts3);
          } else {
            try {
              const opts3 = { q: q, pageSize: 1000, fields: fieldsV3 };
              if (pageToken) opts3.pageToken = pageToken;
              res = Drive.Files.list(opts3);
              detectedVersion = "v3";
            } catch (v3Err) {
              const opts2 = { q: q, maxResults: 1000, fields: fieldsV2 };
              if (pageToken) opts2.pageToken = pageToken;
              res = Drive.Files.list(opts2);
              detectedVersion = "v2";
            }
          }

          const batchItems = (res && Array.isArray(res.files)) ? res.files : ((res && Array.isArray(res.items)) ? res.items : []);
          batchItems.forEach(function(item) { collected.push(item); });
          pageToken = res ? (res.nextPageToken || null) : null;
        } while (pageToken);

        return collected;
      };

      const folderQuery = "'" + rootId + "' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false";
      const folders = driveListAll_(
        folderQuery,
        "nextPageToken, files(id, name)",
        "nextPageToken, items(id, title)"
      );

      const folderIdToName = {};
      const folderIds = [];
      const folderNames = [];
      folders.forEach(function(folder) {
        const fid = folder.id;
        const fname = String(folder.name || folder.title || "").trim();
        if (!fid || !fname || fname === "__TABLE_IMAGES__" || fname === "_pdf_preview_cache") return;
        folderIdToName[fid] = fname;
        folderIds.push(fid);
        if (!folderNames.some(function(n) { return n.toLowerCase() === fname.toLowerCase(); })) {
          folderNames.push(fname);
        }
      });

      if (folderIds.length === 0) {
        return { documents: [], folders: folderNames };
      }

      const docs = [];
      const CHUNK_SIZE = 35;
      for (let i = 0; i < folderIds.length; i += CHUNK_SIZE) {
        const chunk = folderIds.slice(i, i + CHUNK_SIZE);
        const parentsExpr = chunk.map(function(fid) { return "'" + fid + "' in parents"; }).join(" or ");
        const filesQuery = "(" + parentsExpr + ") and mimeType != 'application/vnd.google-apps.folder' and trashed = false";
        const files = driveListAll_(
          filesQuery,
          "nextPageToken, files(id, name, mimeType, description, webViewLink, webContentLink, size, createdTime, parents)",
          "nextPageToken, items(id, title, mimeType, description, alternateLink, webContentLink, fileSize, createdDate, parents(id))"
        );

        files.forEach(function(f) {
          if (!f || !f.id) return;
          if (String(f.description || "") === "__TABLE_IMAGE__") return;

          let prodName = "";
          const parentsList = Array.isArray(f.parents) ? f.parents : [];
          for (let pIdx = 0; pIdx < parentsList.length; pIdx++) {
            const pEntry = parentsList[pIdx];
            const pid = (typeof pEntry === "string") ? pEntry : (pEntry && pEntry.id ? pEntry.id : "");
            if (pid && folderIdToName[pid]) {
              prodName = folderIdToName[pid];
              break;
            }
          }
          if (!prodName) return;

          const createdRaw = f.createdTime || f.createdDate || "";
          const uploadDate = createdRaw
            ? Utilities.formatDate(new Date(createdRaw), tz, "yyyy-MM-dd HH:mm:ss")
            : "";

          docs.push({
            productName: prodName,
            docType: f.description || "Dokumen",
            fileName: f.name || f.title || "",
            fileId: f.id,
            driveUrl: f.webViewLink || f.alternateLink || ("https://drive.google.com/file/d/" + f.id + "/view"),
            downloadUrl: "https://drive.google.com/uc?export=download&id=" + f.id,
            fileSize: Number(f.size != null ? f.size : (f.fileSize || 0)) || 0,
            uploadDate: uploadDate
          });
        });
      }

      return { documents: docs, folders: folderNames };
    } catch (driveErr) {
      console.warn("Drive.Files.list fallback ke DriveApp:", driveErr);
    }
  }

  // Fallback standar DriveApp
  const docs = [];
  const folderNames = [];
  const subFolders = rootFolder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    const prodName = sub.getName();
    if (prodName === "__TABLE_IMAGES__" || prodName === "_pdf_preview_cache") continue;
    if (!folderNames.some(function(n) { return n.toLowerCase() === prodName.toLowerCase(); })) {
      folderNames.push(prodName);
    }
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
        uploadDate: Utilities.formatDate(f.getDateCreated(), tz, "yyyy-MM-dd HH:mm:ss")
      });
    }
  }
  return { documents: docs, folders: folderNames };
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : "";
  const deprecatedGetActions = {
    "get_all_product_data": true,
    "get_all_documents": true,
    "get_file_base64": true
  };

  if (deprecatedGetActions[action]) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      code: "METHOD_NOT_ALLOWED",
      message: "Akses GET untuk data terproteksi telah dinonaktifkan demi keamanan (hindari token di URL). Gunakan POST JSON dengan sessionToken di body."
    })).setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService.createTextOutput("Comparison Maker Backend Active")
    .setMimeType(ContentService.MimeType.TEXT);
}

function doPost(e) {
  let lock = null;
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Permintaan tidak valid (body kosong)." }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    const payload = JSON.parse(e.postData.contents);
    let _cachedSs = null;
    function getSs() {
      if (!_cachedSs) _cachedSs = SpreadsheetApp.getActiveSpreadsheet();
      return _cachedSs;
    }

    // Batas karakter per sel Google Sheets
    const MAX_CELL_CHARS = 50000;

    function jsonResponse(obj) {
      return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const PUBLIC_AUTH_ACTIONS = {
      "auth_signup": true,
      "auth_signin": true,
      "auth_request_otp": true,
      "auth_reset_password": true,
      "auth_logout": true
    };

    // Hanya aksi penulisan Google Sheets & pengiriman OTP yang perlu ScriptLock (mencegah TOCTOU race condition)
    const WRITE_LOCK_ACTIONS = {
      "append": true,
      "update": true,
      "delete": true,
      "save_sheet_category_config": true,
      "save_benefit_aliases": true,
      "manage_benefit_aliases": true,
      "admin_manage_users": true,
      "auth_signup": true,
      "auth_request_otp": true,
      "auth_reset_password": true
    };

    // RBAC: Pembagian izin berbasis Role (viewer < editor < admin)
    const EDITOR_ACTIONS = {
      "append": true,
      "update": true,
      "upload_document": true,
      "save_benefit_aliases": true
    };
    const ADMIN_ACTIONS = {
      "delete": true,
      "delete_document": true,
      "save_sheet_category_config": true,
      "manage_benefit_aliases": true,
      "admin_manage_users": true
    };

    const reqAction = String((payload && payload.action) || "").trim();
    let activeSession = null;
    if (!PUBLIC_AUTH_ACTIONS[reqAction]) {
      activeSession = verifySessionToken_(payload && payload.sessionToken);
      if (!activeSession) {
        return jsonResponse({
          status: "error",
          code: "UNAUTHORIZED",
          message: "Sesi tidak valid atau telah berakhir. Silakan Sign In kembali."
        });
      }
      const userRole = normalizeUserRole_(activeSession.role);
      if (ADMIN_ACTIONS[reqAction] && userRole !== "admin") {
        return jsonResponse({
          status: "error",
          code: "FORBIDDEN",
          message: "Akses ditolak: aksi ini hanya diizinkan untuk Admin."
        });
      }
      if (EDITOR_ACTIONS[reqAction] && userRole !== "editor" && userRole !== "admin") {
        return jsonResponse({
          status: "error",
          code: "FORBIDDEN",
          message: "Akses ditolak: akun Viewer hanya dapat melihat dan membandingkan data."
        });
      }
    }

    if (reqAction === "auth_logout") {
      revokeSessionToken_(payload && payload.sessionToken);
      return jsonResponse({ status: "success", action: "auth_logout" });
    }

    if (WRITE_LOCK_ACTIONS[reqAction]) {
      lock = LockService.getScriptLock();
      if (!lock.tryLock(15000)) {
        lock = null;
        return jsonResponse({ status: "error", message: "Server sibuk, coba lagi dalam beberapa detik" });
      }
    }

    // Helper: Ambil daftar PROTECTED_COLS (Seluruh kolom Core & Non-Benefit — baik bawaan maupun yang baru ditambahkan di Config — di-lock / tidak dihapus otomatis; sedangkan Benefit akan dihapus otomatis jika kosong di semua produk)
    function getProtectedCols(sheetName) {
      const isSyariah = /^(sharia|syariah)$/i.test(String(sheetName || "").trim());
      const isDplk = /^dplk$/i.test(String(sheetName || "").trim());
      const baseProtected = [
        "product name", "company name", "launch date", "tanggal rilis", "par/non par", "product type",
        "distribution channel", "currency", "remarks",
        "additional benefit",
        "entry age", "coverage period", "premium payment period (ppp)",
        "min premium", "min sum assured", "underwriting method", "available riders"
      ].concat(isSyariah ? ["aqd", "akad", "aqd / akad", "alokasi kontribusi", "surplus underwriting"] : [])
       .concat(isDplk ? ["sales tools", "after sales tools"] : []);

      try {
        const configSheet = getSs().getSheetByName("Config");
        if (configSheet && configSheet.getLastRow() > 1) {
          const cfgVals = configSheet.getRange(2, 1, configSheet.getLastRow() - 1, 2).getDisplayValues();
          for (let i = 0; i < cfgVals.length; i++) {
            const cat = String(cfgVals[i][0] || "").trim().toLowerCase();
            const name = String(cfgVals[i][1] || "").trim().toLowerCase();
            if (!name) continue;
            if (
              cat === "core" ||
              cat === "identitas" ||
              cat === "non-benefit" ||
              cat === "non-benefit (syariah)" ||
              cat === "non-benefit (dplk)" ||
              cat.indexOf("special provision") === 0 ||
              cat === "biaya (dplk)" ||
              cat === "dplk charge"
            ) {
              if (baseProtected.indexOf(name) === -1) {
                baseProtected.push(name);
              }
            } else if ((cat === "category provision" || cat === "category provisions") && name === String(sheetName || "").trim().toLowerCase()) {
              // Pastikan juga seluruh general provisions kategori ini diproteksi
            }
          }
        }
      } catch (cfgErr) {
        console.warn("Gagal membaca Config pada getProtectedCols:", cfgErr);
      }

      return baseProtected;
    }

    // Helper: Sinkronisasi Header Kolom Sheet dengan Header Baru dari Web / Excel
    function syncSheetHeaders(targetSheet, incomingHeaders) {
      let sheetHeaders = [];
      const lastCol = targetSheet.getLastColumn();
      if (lastCol > 0) {
        sheetHeaders = targetSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || "").trim());
      }

      if (sheetHeaders.length === 0 || sheetHeaders.every(h => h === "")) {
        const initHdrRange = targetSheet.getRange(1, 1, 1, incomingHeaders.length);
        initHdrRange.setNumberFormat("@");
        initHdrRange.setValues([sanitizeSheetRow_(incomingHeaders)]);
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
        const addHdrRange = targetSheet.getRange(1, startCol, 1, newColsToAdd.length);
        addHdrRange.setNumberFormat("@");
        addHdrRange.setValues([sanitizeSheetRow_(newColsToAdd)]);
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
        return sanitizeSheetCell_(String(val));
      });
      range.setValues([strRowData]);
    }

    // Helper: Pembersihan Kolom Manfaat (Benefit) Kosong dari Sheet Produk & Config
    // Catatan: Kolom Non-Benefit (baik existing maupun baru) sudah masuk getProtectedCols sehingga TIDAK dihapus.
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
        const deletedBenefitKeys = {};
        let hasDeletedAny = false;

        for (let c = lastCol - 1; c >= 0; c--) {
          const rawHeader = String(headers[c] || "").trim();
          const hName = rawHeader.toLowerCase();
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
            const normKey = normalizeBenefitAliasKey_(rawHeader);
            if (normKey) {
              deletedBenefitKeys[normKey] = true;
              hasDeletedAny = true;
            }
          }
        }

        // Jika ada kolom Benefit yang dihapus, cek apakah Benefit tersebut masih dipakai di sheet produk lain.
        // Apabila sudah tidak dipakai di produk mana pun, hapus juga dari sheet Config.
        if (hasDeletedAny) {
          const stillUsedBenefitKeys = {};
          const allSheets = getSs().getSheets();
          for (let s = 0; s < allSheets.length; s++) {
            const sh = allSheets[s];
            const shNameL = String(sh.getName() || "").trim().toLowerCase();
            if (shNameL === "config" || shNameL === "users" || shNameL === "documents") continue;
            const shLastCol = sh.getLastColumn();
            const shLastRow = sh.getLastRow();
            if (shLastCol <= 1 || shLastRow <= 1) continue;
            const shHeaders = sh.getRange(1, 1, 1, shLastCol).getValues()[0];
            const shData = sh.getRange(2, 1, shLastRow - 1, shLastCol).getValues();
            for (let colIdx = 0; colIdx < shHeaders.length; colIdx++) {
              const k = normalizeBenefitAliasKey_(shHeaders[colIdx]);
              if (!k || !deletedBenefitKeys[k]) continue;
              for (let rIdx = 0; rIdx < shData.length; rIdx++) {
                const v = String(shData[rIdx][colIdx] || "").trim();
                if (v !== "" && v !== "-") {
                  stillUsedBenefitKeys[k] = true;
                  break;
                }
              }
            }
          }

          const configSheet = getSs().getSheetByName("Config");
          if (configSheet && configSheet.getLastRow() > 1) {
            const cfgVals = configSheet.getRange(1, 1, configSheet.getLastRow(), 2).getDisplayValues();
            for (let r = cfgVals.length - 1; r >= 1; r--) {
              const catL = String(cfgVals[r][0] || "").trim().toLowerCase();
              const canonicalKey = normalizeBenefitAliasKey_(cfgVals[r][1]);
              if (catL === "benefit" && canonicalKey && deletedBenefitKeys[canonicalKey] && !stillUsedBenefitKeys[canonicalKey]) {
                configSheet.deleteRow(r + 1);
              }
            }
          }
        }
      } catch (err) {
        console.error("Gagal membersihkan kolom kosong:", err);
      }
    }

    // 00. ACTION: GET_ALL_PRODUCT_DATA (Via POST — unwrap __rawCachedJsonText identik dengan doGet)
    if (payload.action === "get_all_product_data") {
      const bypassCache = payload.nocache === true || payload.nocache === "1" || payload.nocache === "true";
      const result = buildAllProductDataResponse(payload.mode, bypassCache);
      if (result && result.__rawCachedJsonText) {
        return ContentService.createTextOutput(result.__rawCachedJsonText)
          .setMimeType(ContentService.MimeType.JSON);
      }
      return jsonResponse(result);
    }

    // 00B. ACTION: GET_ALL_DOCUMENTS (Via POST agar sessionToken tidak lewat query URL)
    if (payload.action === "get_all_documents") {
      const docResult = listAllProductDocumentsFast_();
      return jsonResponse({
        status: "success",
        documents: docResult.documents,
        folders: docResult.folders
      });
    }

    // ====================================================
    // 0. SISTEM AUTENTIKASI (SIGN UP, SIGN IN, FORGOT PASSWORD OTP EMAIL)
    // ====================================================
    const PASSWORD_HASH_ITERATIONS = 10000;
    const AUTH_RATE_LIMIT_LOCK_SEC = 900; // 15 menit
    const AUTH_SIGNIN_MAX_FAILS = 5;
    const AUTH_INVITE_MAX_FAILS = 10;
    const OTP_GLOBAL_MAX_PER_WINDOW = 15;
    const OTP_GLOBAL_WINDOW_SEC = 600; // 10 menit

    // Pure V8 JS SHA-256 / HMAC-SHA256 / PBKDF2-HMAC-SHA256 helpers (hindari 10.000+ RPC Utilities bridge yang lambat)
    const SHA256_K_JS_ = new Int32Array([
      0x428a2f98, 0x71374491, 0xb5c0fbcf, 0xe9b5dba5, 0x3956c25b, 0x59f111f1, 0x923f82a4, 0xab1c5ed5,
      0xd807aa98, 0x12835b01, 0x243185be, 0x550c7dc3, 0x72be5d74, 0x80deb1fe, 0x9bdc06a7, 0xc19bf174,
      0xe49b69c1, 0xefbe4786, 0x0fc19dc6, 0x240ca1cc, 0x2de92c6f, 0x4a7484aa, 0x5cb0a9dc, 0x76f988da,
      0x983e5152, 0xa831c66d, 0xb00327c8, 0xbf597fc7, 0xc6e00bf3, 0xd5a79147, 0x06ca6351, 0x14292967,
      0x27b70a85, 0x2e1b2138, 0x4d2c6dfc, 0x53380d13, 0x650a7354, 0x766a0abb, 0x81c2c92e, 0x92722c85,
      0xa2bfe8a1, 0xa81a664b, 0xc24b8b70, 0xc76c51a3, 0xd192e819, 0xd6990624, 0xf40e3585, 0x106aa070,
      0x19a4c116, 0x1e376c08, 0x2748774c, 0x34b0bcb5, 0x391c0cb3, 0x4ed8aa4a, 0x5b9cca4f, 0x682e6ff3,
      0x748f82ee, 0x78a5636f, 0x84c87814, 0x8cc70208, 0x90befffa, 0xa4506ceb, 0xbef9a3f7, 0xc67178f2
    ]);
    const SHA256_W_JS_ = new Int32Array(64);

    function utf8ToBytesJs_(str) {
      const s = String(str === undefined || str === null ? "" : str);
      const out = [];
      for (let i = 0; i < s.length; i++) {
        let cp = s.charCodeAt(i);
        if (cp < 0x80) {
          out.push(cp);
        } else if (cp < 0x800) {
          out.push(0xc0 | (cp >> 6), 0x80 | (cp & 0x3f));
        } else if (cp >= 0xd800 && cp <= 0xdbff && i + 1 < s.length) {
          const low = s.charCodeAt(i + 1);
          if (low >= 0xdc00 && low <= 0xdfff) {
            cp = 0x10000 + (((cp - 0xd800) << 10) | (low - 0xdc00));
            i++;
            out.push(
              0xf0 | (cp >> 18),
              0x80 | ((cp >> 12) & 0x3f),
              0x80 | ((cp >> 6) & 0x3f),
              0x80 | (cp & 0x3f)
            );
          } else {
            out.push(0xef, 0xbf, 0xbd);
          }
        } else if (cp >= 0xd800 && cp <= 0xdfff) {
          out.push(0xef, 0xbf, 0xbd);
        } else {
          out.push(0xe0 | (cp >> 12), 0x80 | ((cp >> 6) & 0x3f), 0x80 | (cp & 0x3f));
        }
      }
      return new Uint8Array(out);
    }

    function sha256BytesJs_(msgBytes) {
      const len = msgBytes.length;
      const paddedLen = (((len + 8) >> 6) + 1) << 6;
      const padded = new Uint8Array(paddedLen);
      padded.set(msgBytes, 0);
      padded[len] = 0x80;
      const bitLenHi = Math.floor(len / 0x20000000);
      const bitLenLo = (len << 3) >>> 0;
      padded[paddedLen - 8] = (bitLenHi >>> 24) & 0xff;
      padded[paddedLen - 7] = (bitLenHi >>> 16) & 0xff;
      padded[paddedLen - 6] = (bitLenHi >>> 8) & 0xff;
      padded[paddedLen - 5] = bitLenHi & 0xff;
      padded[paddedLen - 4] = (bitLenLo >>> 24) & 0xff;
      padded[paddedLen - 3] = (bitLenLo >>> 16) & 0xff;
      padded[paddedLen - 2] = (bitLenLo >>> 8) & 0xff;
      padded[paddedLen - 1] = bitLenLo & 0xff;

      let h0 = 0x6a09e667 | 0;
      let h1 = 0xbb67ae85 | 0;
      let h2 = 0x3c6ef372 | 0;
      let h3 = 0xa54ff53a | 0;
      let h4 = 0x510e527f | 0;
      let h5 = 0x9b05688c | 0;
      let h6 = 0x1f83d9ab | 0;
      let h7 = 0x5be0cd19 | 0;

      const W = SHA256_W_JS_;
      const K = SHA256_K_JS_;

      for (let offset = 0; offset < paddedLen; offset += 64) {
        for (let i = 0; i < 16; i++) {
          const j = offset + (i << 2);
          W[i] = (padded[j] << 24) | (padded[j + 1] << 16) | (padded[j + 2] << 8) | padded[j + 3];
        }
        for (let i = 16; i < 64; i++) {
          const w15 = W[i - 15];
          const w2 = W[i - 2];
          const s0 = ((w15 >>> 7) | (w15 << 25)) ^ ((w15 >>> 18) | (w15 << 14)) ^ (w15 >>> 3);
          const s1 = ((w2 >>> 17) | (w2 << 15)) ^ ((w2 >>> 19) | (w2 << 13)) ^ (w2 >>> 10);
          W[i] = (W[i - 16] + s0 + W[i - 7] + s1) | 0;
        }

        let a = h0, b = h1, c = h2, d = h3, e = h4, f = h5, g = h6, h = h7;
        for (let i = 0; i < 64; i++) {
          const S1 = ((e >>> 6) | (e << 26)) ^ ((e >>> 11) | (e << 21)) ^ ((e >>> 25) | (e << 7));
          const ch = (e & f) ^ (~e & g);
          const temp1 = (h + S1 + ch + K[i] + W[i]) | 0;
          const S0 = ((a >>> 2) | (a << 30)) ^ ((a >>> 13) | (a << 19)) ^ ((a >>> 22) | (a << 10));
          const maj = (a & b) ^ (a & c) ^ (b & c);
          const temp2 = (S0 + maj) | 0;

          h = g;
          g = f;
          f = e;
          e = (d + temp1) | 0;
          d = c;
          c = b;
          b = a;
          a = (temp1 + temp2) | 0;
        }

        h0 = (h0 + a) | 0;
        h1 = (h1 + b) | 0;
        h2 = (h2 + c) | 0;
        h3 = (h3 + d) | 0;
        h4 = (h4 + e) | 0;
        h5 = (h5 + f) | 0;
        h6 = (h6 + g) | 0;
        h7 = (h7 + h) | 0;
      }

      const out = new Uint8Array(32);
      const hs = [h0, h1, h2, h3, h4, h5, h6, h7];
      for (let i = 0; i < 8; i++) {
        const v = hs[i];
        out[i * 4] = (v >>> 24) & 0xff;
        out[i * 4 + 1] = (v >>> 16) & 0xff;
        out[i * 4 + 2] = (v >>> 8) & 0xff;
        out[i * 4 + 3] = v & 0xff;
      }
      return out;
    }

    function hmacSha256BytesJs_(keyBytes, msgBytes) {
      let key = keyBytes;
      if (key.length > 64) key = sha256BytesJs_(key);
      const inner = new Uint8Array(64 + msgBytes.length);
      const outer = new Uint8Array(64 + 32);
      for (let i = 0; i < 64; i++) {
        const k = i < key.length ? key[i] : 0;
        inner[i] = k ^ 0x36;
        outer[i] = k ^ 0x5c;
      }
      inner.set(msgBytes, 64);
      outer.set(sha256BytesJs_(inner), 64);
      return sha256BytesJs_(outer);
    }

    function bytesToHexUnsigned_(uint8) {
      let hex = "";
      for (let i = 0; i < uint8.length; i++) {
        const b = uint8[i];
        hex += (b < 16 ? "0" : "") + b.toString(16);
      }
      return hex;
    }

    const MAX_PASSWORD_LENGTH = 256;

    function pbkdf2HmacSha256HexJs_(password, salt, iterations) {
      const rawPwdBytes = utf8ToBytesJs_(password);
      // Normalisasi key dan precompute blok ipad/opad 1x di luar loop (anti CPU DoS & lebih cepat)
      const normKey = rawPwdBytes.length > 64 ? sha256BytesJs_(rawPwdBytes) : rawPwdBytes;
      const ipad = new Uint8Array(64);
      const opad = new Uint8Array(64);
      for (let i = 0; i < 64; i++) {
        const k = i < normKey.length ? normKey[i] : 0;
        ipad[i] = k ^ 0x36;
        opad[i] = k ^ 0x5c;
      }

      const saltPrefixBytes = utf8ToBytesJs_(String(salt || "") + "::IDCM_AUTH::");
      const firstInner = new Uint8Array(64 + saltPrefixBytes.length + 4);
      firstInner.set(ipad, 0);
      firstInner.set(saltPrefixBytes, 64);
      firstInner[64 + saltPrefixBytes.length + 3] = 1; // INT(1)

      const outerBuf = new Uint8Array(96);
      outerBuf.set(opad, 0);
      outerBuf.set(sha256BytesJs_(firstInner), 64);
      let u = sha256BytesJs_(outerBuf);

      const dk = new Uint8Array(u);
      const innerBuf = new Uint8Array(96);
      innerBuf.set(ipad, 0);

      for (let i = 1; i < iterations; i++) {
        innerBuf.set(u, 64);
        outerBuf.set(sha256BytesJs_(innerBuf), 64);
        u = sha256BytesJs_(outerBuf);
        for (let j = 0; j < 32; j++) {
          dk[j] ^= u[j];
        }
      }
      return bytesToHexUnsigned_(dk);
    }

    function shortCacheHash_(str) {
      return bytesToHexUnsigned_(sha256BytesJs_(utf8ToBytesJs_(str))).substring(0, 24);
    }

    function legacySingleSha256_(password, salt) {
      const raw = String(salt || "") + "::IDCM_AUTH::" + String(password || "");
      return bytesToHexUnsigned_(sha256BytesJs_(utf8ToBytesJs_(raw)));
    }

    function legacyIteratedSha256_(password, salt, iterations) {
      const iters = (typeof iterations === "number" && iterations >= 1000) ? iterations : 2500;
      const prefix = String(salt || "") + "::IDCM_AUTH::" + String(password || "");
      let block = sha256BytesJs_(utf8ToBytesJs_(prefix));
      for (let i = 1; i < iters; i++) {
        block = sha256BytesJs_(block);
      }
      return "pbkdf2_sha256_" + iters + "$" + bytesToHexUnsigned_(block);
    }

    function legacyChainedHmacSha256_(password, salt, iterations) {
      const iters = (typeof iterations === "number" && iterations >= 1000) ? iterations : PASSWORD_HASH_ITERATIONS;
      const pwdStr = String(password || "");
      const keyBytes = utf8ToBytesJs_(pwdStr);
      let block = hmacSha256BytesJs_(keyBytes, utf8ToBytesJs_(String(salt || "") + "::IDCM_AUTH::" + pwdStr));
      for (let i = 1; i < iters; i++) {
        block = hmacSha256BytesJs_(keyBytes, block);
      }
      return "hmac_sha256_" + iters + "$" + bytesToHexUnsigned_(block);
    }

    function hashPasswordWithSalt(password, salt, iterations) {
      const iters = (typeof iterations === "number" && iterations >= 1000) ? iterations : PASSWORD_HASH_ITERATIONS;
      return "pbkdf2_hmac_sha256_" + iters + "$" + pbkdf2HmacSha256HexJs_(password, salt, iters);
    }

    function verifyPasswordAutoMigrate_(password, storedHash, storedSalt) {
      const stored = String(storedHash || "").trim();
      const pbkdf2Match = stored.match(/^pbkdf2_hmac_sha256_(\d+)\$([a-f0-9]+)$/i);
      if (pbkdf2Match) {
        const iters = Math.max(1000, Math.min(50000, parseInt(pbkdf2Match[1], 10) || PASSWORD_HASH_ITERATIONS));
        const candidate = hashPasswordWithSalt(password, storedSalt, iters);
        const ok = constantTimeEquals_(candidate, stored);
        return { ok: ok, needsRehash: ok && iters < PASSWORD_HASH_ITERATIONS };
      }
      const hmacMatch = stored.match(/^hmac_sha256_(\d+)\$([a-f0-9]+)$/i);
      if (hmacMatch) {
        const iters = Math.max(1000, Math.min(50000, parseInt(hmacMatch[1], 10) || PASSWORD_HASH_ITERATIONS));
        const candidate = legacyChainedHmacSha256_(password, storedSalt, iters);
        const ok = constantTimeEquals_(candidate, stored);
        return { ok: ok, needsRehash: ok };
      }
      const oldIterMatch = stored.match(/^pbkdf2_sha256_(\d+)\$([a-f0-9]+)$/i);
      if (oldIterMatch) {
        const iters = Math.max(1000, Math.min(20000, parseInt(oldIterMatch[1], 10) || 2500));
        const candidate = legacyIteratedSha256_(password, storedSalt, iters);
        const ok = constantTimeEquals_(candidate, stored);
        return { ok: ok, needsRehash: ok };
      }
      const legacyCandidate = legacySingleSha256_(password, storedSalt);
      const ok = constantTimeEquals_(legacyCandidate, stored);
      return { ok: ok, needsRehash: ok };
    }

    function findUserRowFast_(uSheet, identifierLower) {
      const cleanId = String(identifierLower || "").trim().toLowerCase();
      if (!cleanId) return null;
      const lastRow = uSheet.getLastRow();
      if (lastRow < 2) return null;

      const cache = CacheService.getScriptCache();
      const idxCacheKey = "USER_ROW_IDX_" + shortCacheHash_(cleanId);
      try {
        const cachedRowNum = parseInt(cache.get(idxCacheKey) || "0", 10);
        if (cachedRowNum >= 2 && cachedRowNum <= lastRow) {
          const rowVals = uSheet.getRange(cachedRowNum, 1, 1, 9).getValues()[0];
          const uName = String(rowVals[0] || "").trim().toLowerCase();
          const uEmail = String(rowVals[1] || "").trim().toLowerCase();
          if (uName === cleanId || uEmail === cleanId) {
            return { rowNum: cachedRowNum, row: rowVals };
          }
        }
      } catch (e) {}

      const foundRange = uSheet.getRange(2, 1, lastRow - 1, 2)
        .createTextFinder(cleanId)
        .matchEntireCell(true)
        .matchCase(false)
        .findNext();

      if (foundRange) {
        const rowNum = foundRange.getRow();
        const rowVals = uSheet.getRange(rowNum, 1, 1, 9).getValues()[0];
        try {
          cache.put(idxCacheKey, String(rowNum), 600);
        } catch (e) {}
        return { rowNum: rowNum, row: rowVals };
      }

      // Fallback untuk akun lama jika ada spasi tersembunyi di sel Username/Email pada sheet Users
      const nameEmailRows = uSheet.getRange(2, 1, lastRow - 1, 2).getValues();
      for (let i = 0; i < nameEmailRows.length; i++) {
        const uName = String(nameEmailRows[i][0] || "").trim().toLowerCase();
        const uEmail = String(nameEmailRows[i][1] || "").trim().toLowerCase();
        if (uName === cleanId || uEmail === cleanId) {
          const rowNum = i + 2;
          const rowVals = uSheet.getRange(rowNum, 1, 1, 9).getValues()[0];
          try {
            cache.put(idxCacheKey, String(rowNum), 600);
          } catch (e) {}
          return { rowNum: rowNum, row: rowVals };
        }
      }
      return null;
    }

    function hashOtp_(otp, salt) {
      const raw = String(salt || "") + "::IDCM_OTP::" + String(otp || "");
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw, Utilities.Charset.UTF_8);
      return "otp_sha256$" + bytesToHex_(digest);
    }

    function generateSecureOtp6_() {
      const seed = Utilities.getUuid() + "::" + Utilities.getUuid() + "::" + Date.now();
      const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, seed, Utilities.Charset.UTF_8);
      const b0 = (digest[0] < 0) ? digest[0] + 256 : digest[0];
      const b1 = (digest[1] < 0) ? digest[1] + 256 : digest[1];
      const b2 = (digest[2] < 0) ? digest[2] + 256 : digest[2];
      const b3 = (digest[3] < 0) ? digest[3] + 256 : digest[3];
      const uint32 = (((b0 << 24) | (b1 << 16) | (b2 << 8) | b3) >>> 0);
      return String(100000 + (uint32 % 900000));
    }

    function getOrCreateUsersSheet() {
      let uSheet = getSs().getSheetByName("Users");
      const headers = ["Username", "Email", "PasswordHash", "Salt", "CreatedAt", "ResetOTP", "OTPExpiry", "LastLogin", "Role"];
      if (!uSheet) {
        uSheet = getSs().insertSheet("Users");
        const r = uSheet.getRange(1, 1, 1, headers.length);
        r.setNumberFormat("@");
        r.setValues([headers]);
        r.setFontWeight("bold");
      } else if (uSheet.getLastRow() === 0) {
        const r = uSheet.getRange(1, 1, 1, headers.length);
        r.setNumberFormat("@");
        r.setValues([headers]);
        r.setFontWeight("bold");
      } else if (uSheet.getLastColumn() < headers.length) {
        const roleHdr = uSheet.getRange(1, 9);
        roleHdr.setNumberFormat("@");
        roleHdr.setValue("Role");
        roleHdr.setFontWeight("bold");
      }
      return uSheet;
    }

    function getSignupInviteCode() {
      try {
        const configSheet = getSs().getSheetByName("Config");
        if (!configSheet) return "";
        const cfgData = configSheet.getDataRange().getValues();
        for (let i = 1; i < cfgData.length; i++) {
          const key = String(cfgData[i][1] || "").trim().toLowerCase();
          if (key === "invite code" || key === "signup invite code" || key === "kode akses") {
            const codeVal = String(cfgData[i][2] || "").trim();
            if (codeVal) return codeVal;
          }
        }
      } catch (e) {
        console.warn("Gagal membaca Invite Code dari Config:", e);
      }
      return "";
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

      if (username.length < 3 || username.length > 64) {
        return jsonResponse({ status: "error", message: "Username harus antara 3 hingga 64 karakter." });
      }
      if (/^[=+\-@\t\r]/.test(username) || /[<>]/.test(username) || /[\x00-\x1F]/.test(username)) {
        return jsonResponse({ status: "error", message: "Username mengandung karakter yang tidak diizinkan." });
      }
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) {
        return jsonResponse({ status: "error", message: "Format alamat email tidak valid." });
      }
      if (password.length < 6 || password.length > MAX_PASSWORD_LENGTH) {
        return jsonResponse({ status: "error", message: "Password harus antara 6 hingga " + MAX_PASSWORD_LENGTH + " karakter." });
      }

      const authCache = CacheService.getScriptCache();
      const inviteLockKey = "INVITE_LOCK_GLOBAL";
      const inviteFailKey = "INVITE_FAIL_GLOBAL";
      if (authCache.get(inviteLockKey)) {
        return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Kode Akses salah. Silakan coba lagi dalam 15 menit." });
      }

      const validInviteCode = getSignupInviteCode();
      if (!validInviteCode) {
        return jsonResponse({ status: "error", message: "Pendaftaran ditutup sementara: Invite Code belum dikonfigurasi oleh Administrator di sheet Config." });
      }
      if (!constantTimeEquals_(inviteCode, validInviteCode)) {
        const invFails = (parseInt(authCache.get(inviteFailKey) || "0", 10) || 0) + 1;
        if (invFails >= AUTH_INVITE_MAX_FAILS) {
          authCache.put(inviteLockKey, "1", AUTH_RATE_LIMIT_LOCK_SEC);
          authCache.remove(inviteFailKey);
          return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Kode Akses salah. Pendaftaran dikunci sementara selama 15 menit." });
        }
        authCache.put(inviteFailKey, String(invFails), AUTH_RATE_LIMIT_LOCK_SEC);
        return jsonResponse({ status: "error", message: "Kode Akses Tim (Invite Code) tidak sesuai." });
      }
      authCache.remove(inviteFailKey);

      const uSheet = getOrCreateUsersSheet();
      if (findUserRowFast_(uSheet, username.toLowerCase())) {
        return jsonResponse({ status: "error", message: "Username sudah digunakan. Silakan pilih username lain." });
      }
      if (findUserRowFast_(uSheet, email)) {
        return jsonResponse({ status: "error", message: "Alamat email sudah terdaftar. Silakan Sign In atau gunakan fitur Lupa Password." });
      }

      const salt = Utilities.getUuid();
      const pwdHash = hashPasswordWithSalt(password, salt, PASSWORD_HASH_ITERATIONS);
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");
      const newRowNum = uSheet.getLastRow() + 1;
      const defaultRole = "viewer";
      writeRowAsPlainText(uSheet, newRowNum, [username, email, pwdHash, salt, nowStr, "", "", nowStr, defaultRole]);
      try {
        authCache.put("USER_ROW_IDX_" + shortCacheHash_(username.toLowerCase()), String(newRowNum), 600);
        authCache.put("USER_ROW_IDX_" + shortCacheHash_(email), String(newRowNum), 600);
      } catch (e) {}

      const sessionObj = issueSessionToken_(username, email, SESSION_REMEMBER_TTL_MS, defaultRole);
      return jsonResponse({
        status: "success",
        action: "auth_signup",
        sessionToken: sessionObj.token,
        expiresAt: sessionObj.expiresAt,
        user: { username: username, email: email, role: defaultRole }
      });
    }

    // 0B. ACTION: AUTH_SIGNIN
    if (payload.action === "auth_signin") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      const password = String(payload.password || "");
      const rememberMe = payload.rememberMe !== false;

      if (!identifier || !password) {
        return jsonResponse({ status: "error", message: "Mohon isi Username/Email dan Password." });
      }
      if (password.length > MAX_PASSWORD_LENGTH) {
        return jsonResponse({ status: "error", message: "Username/Email atau Password tidak sesuai." });
      }

      const authCache = CacheService.getScriptCache();
      const idHash = shortCacheHash_(identifier);
      const loginLockKey = "LOGIN_LOCK_" + idHash;
      const loginFailKey = "LOGIN_FAIL_" + idHash;
      if (authCache.get(loginLockKey)) {
        return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Sign In gagal. Akun dikunci sementara selama 15 menit." });
      }

      const uSheet = getOrCreateUsersSheet();
      const found = findUserRowFast_(uSheet, identifier);

      if (!found) {
        // Dummy PBKDF2 untuk menyamakan waktu komputasi (anti timing enumeration)
        pbkdf2HmacSha256HexJs_(password, "idcm_dummy_timing_salt", PASSWORD_HASH_ITERATIONS);
        const fails = (parseInt(authCache.get(loginFailKey) || "0", 10) || 0) + 1;
        if (fails >= AUTH_SIGNIN_MAX_FAILS) {
          authCache.put(loginLockKey, "1", AUTH_RATE_LIMIT_LOCK_SEC);
          authCache.remove(loginFailKey);
          return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Sign In gagal (5x). Silakan coba lagi dalam 15 menit." });
        }
        authCache.put(loginFailKey, String(fails), AUTH_RATE_LIMIT_LOCK_SEC);
        return jsonResponse({ status: "error", message: "Username/Email atau Password tidak sesuai." });
      }

      const row = found.row;
      const canonicalUserLower = String(row[0] || "").trim().toLowerCase();
      const canonicalHash = shortCacheHash_(canonicalUserLower || identifier);
      const canonLockKey = "LOGIN_LOCK_U_" + canonicalHash;
      const canonFailKey = "LOGIN_FAIL_U_" + canonicalHash;
      if (authCache.get(canonLockKey)) {
        return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Sign In gagal. Akun dikunci sementara selama 15 menit." });
      }

      const storedHash = String(row[2] || "").trim();
      const storedSalt = String(row[3] || "").trim();
      const pwdCheck = verifyPasswordAutoMigrate_(password, storedHash, storedSalt);

      if (!pwdCheck.ok) {
        const rawFails = (parseInt(authCache.get(loginFailKey) || "0", 10) || 0) + 1;
        const canonFails = (parseInt(authCache.get(canonFailKey) || "0", 10) || 0) + 1;
        const effectiveFails = Math.max(rawFails, canonFails);
        if (effectiveFails >= AUTH_SIGNIN_MAX_FAILS) {
          authCache.put(loginLockKey, "1", AUTH_RATE_LIMIT_LOCK_SEC);
          authCache.put(canonLockKey, "1", AUTH_RATE_LIMIT_LOCK_SEC);
          authCache.remove(loginFailKey);
          authCache.remove(canonFailKey);
          return jsonResponse({ status: "error", message: "Terlalu banyak percobaan Sign In gagal (5x). Silakan coba lagi dalam 15 menit." });
        }
        authCache.put(loginFailKey, String(rawFails), AUTH_RATE_LIMIT_LOCK_SEC);
        authCache.put(canonFailKey, String(canonFails), AUTH_RATE_LIMIT_LOCK_SEC);
        return jsonResponse({ status: "error", message: "Username/Email atau Password tidak sesuai." });
      }

      authCache.remove(loginFailKey);
      authCache.remove(loginLockKey);
      authCache.remove(canonFailKey);
      authCache.remove(canonLockKey);

      if (pwdCheck.needsRehash) {
        const newSalt = Utilities.getUuid();
        const upgradedHash = hashPasswordWithSalt(password, newSalt, PASSWORD_HASH_ITERATIONS);
        uSheet.getRange(found.rowNum, 3, 1, 2).setNumberFormats([["@", "@"]]).setValues([[upgradedHash, newSalt]]);
      }

      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");
      uSheet.getRange(found.rowNum, 8).setNumberFormat("@").setValue(nowStr);

      const uname = String(row[0] || "").trim();
      const uemail = String(row[1] || "").trim();
      const urole = normalizeUserRole_(row[8]);
      const ttlMs = rememberMe ? SESSION_REMEMBER_TTL_MS : SESSION_DEFAULT_TTL_MS;
      const sessionObj = issueSessionToken_(uname, uemail, ttlMs, urole);

      return jsonResponse({
        status: "success",
        action: "auth_signin",
        sessionToken: sessionObj.token,
        expiresAt: sessionObj.expiresAt,
        user: {
          username: uname,
          email: uemail,
          role: urole
        }
      });
    }

    // 0E. ACTION: ADMIN_MANAGE_USERS (Khusus Admin: List User, Ubah Role, Hapus User & Cabut Sesi)
    if (payload.action === "admin_manage_users") {
      const uSheet = getOrCreateUsersSheet();
      const op = String(payload.op || "list").trim().toLowerCase();
      const actorLower = String((activeSession && activeSession.username) || "").trim().toLowerCase();

      function readAllUsersInternal_() {
        const lastRow = uSheet.getLastRow();
        if (lastRow < 2) return [];
        const vals = uSheet.getRange(2, 1, lastRow - 1, 9).getDisplayValues();
        const out = [];
        for (let i = 0; i < vals.length; i++) {
          const u = String(vals[i][0] || "").trim();
          if (!u) continue;
          out.push({
            rowNum: i + 2,
            username: u,
            emailLower: String(vals[i][1] || "").trim().toLowerCase(),
            passwordHash: String(vals[i][2] || "").trim(),
            salt: String(vals[i][3] || "").trim(),
            createdAt: String(vals[i][4] || "").trim(),
            lastLogin: String(vals[i][7] || "").trim(),
            role: normalizeUserRole_(vals[i][8])
          });
        }
        return out;
      }

      function toPublicUsersList_(internalUsers) {
        return internalUsers.map(function(u) {
          return {
            rowNum: u.rowNum,
            username: u.username,
            createdAt: u.createdAt,
            lastLogin: u.lastLogin,
            role: u.role
          };
        });
      }

      // Re-verifikasi role aktor di dalam ScriptLock (mencegah race condition jika 2 Admin saling demote bersamaan)
      const allUsersNow = readAllUsersInternal_();
      const actorRecord = allUsersNow.find(function(u) { return u.username.toLowerCase() === actorLower; });
      if (!actorRecord || actorRecord.role !== "admin") {
        revokeAllSessionsForUser_(actorLower);
        return jsonResponse({
          status: "error",
          code: "UNAUTHORIZED",
          message: "Hak akses Admin Anda telah dicabut atau berubah. Silakan Sign In kembali."
        });
      }

      if (op === "list") {
        return jsonResponse({ status: "success", action: "admin_manage_users", users: toPublicUsersList_(allUsersNow) });
      }

      if (op === "set_role") {
        const targetUsername = String(payload.username || "").trim().toLowerCase();
        const nextRole = normalizeUserRole_(payload.role);
        if (!targetUsername) {
          return jsonResponse({ status: "error", message: "Username target tidak valid." });
        }
        const target = allUsersNow.find(function(u) { return u.username.toLowerCase() === targetUsername; });
        if (!target) {
          return jsonResponse({ status: "error", message: "User tidak ditemukan." });
        }
        if (target.role === nextRole) {
          return jsonResponse({ status: "success", action: "admin_manage_users", selfDemoted: false, users: toPublicUsersList_(allUsersNow) });
        }

        // Validasi ketat Last Admin: tolak tanpa mencabut sesi jika hanya tersisa 1 Admin
        if (target.role === "admin" && nextRole !== "admin") {
          const adminCount = allUsersNow.filter(function(u) { return u.role === "admin"; }).length;
          if (adminCount <= 1) {
            return jsonResponse({ status: "error", message: "Tidak dapat menurunkan role Admin terakhir. Promosikan user lain ke Admin terlebih dahulu." });
          }
        }

        // Opsi 2: Setiap perubahan yang melibatkan role Admin (promote ke Admin / demote Admin / demote diri sendiri) wajib verifikasi password Admin
        const involvesAdminRole = (target.role === "admin" || nextRole === "admin");
        if (involvesAdminRole) {
          const adminPwd = String(payload.adminPassword || "");
          if (!adminPwd || adminPwd.length > MAX_PASSWORD_LENGTH) {
            return jsonResponse({ status: "error", code: "INVALID_ADMIN_PASSWORD", message: "Password Admin wajib diisi untuk mengonfirmasi perubahan role Admin." });
          }
          const pwdCheck = verifyPasswordAutoMigrate_(adminPwd, actorRecord.passwordHash, actorRecord.salt);
          if (!pwdCheck.ok) {
            return jsonResponse({ status: "error", code: "INVALID_ADMIN_PASSWORD", message: "Password Admin tidak sesuai." });
          }
        }

        uSheet.getRange(target.rowNum, 9).setNumberFormat("@").setValue(sanitizeSheetCell_(nextRole));
        revokeAllSessionsForUser_(targetUsername);
        const selfDemoted = (targetUsername === actorLower && nextRole !== "admin");
        return jsonResponse({
          status: "success",
          action: "admin_manage_users",
          selfDemoted: selfDemoted,
          users: toPublicUsersList_(readAllUsersInternal_())
        });
      }

      if (op === "delete_user") {
        const targetUsername = String(payload.username || "").trim().toLowerCase();
        if (!targetUsername) {
          return jsonResponse({ status: "error", message: "Username target tidak valid." });
        }
        if (targetUsername === actorLower) {
          return jsonResponse({ status: "error", message: "Anda tidak dapat menghapus akun Anda sendiri." });
        }
        const target = allUsersNow.find(function(u) { return u.username.toLowerCase() === targetUsername; });
        if (!target) {
          return jsonResponse({ status: "error", message: "User tidak ditemukan." });
        }
        uSheet.deleteRow(target.rowNum);
        revokeAllSessionsForUser_(targetUsername);
        try {
          const c = CacheService.getScriptCache();
          c.remove("USER_ROW_IDX_" + shortCacheHash_(targetUsername));
          if (target.emailLower) c.remove("USER_ROW_IDX_" + shortCacheHash_(target.emailLower));
        } catch (e) {}
        return jsonResponse({ status: "success", action: "admin_manage_users", users: toPublicUsersList_(readAllUsersInternal_()) });
      }

      return jsonResponse({ status: "error", message: "Operasi admin_manage_users tidak dikenali." });
    }

    // 0C. ACTION: AUTH_REQUEST_OTP (Kirim Kode OTP 6 Digit ke Email — Anti-Enumerasi & Global Rate Limit)
    if (payload.action === "auth_request_otp") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      if (!identifier) {
        return jsonResponse({ status: "error", message: "Masukkan Email atau Username akun Anda." });
      }

      const genericOtpResp = {
        status: "success",
        action: "auth_request_otp",
        message: "Jika akun terdaftar, kode OTP 6 digit telah dikirim ke email terkait. Silakan cek kotak masuk atau folder spam Anda."
      };

      const cache = CacheService.getScriptCache();
      const globalOtpKey = "OTP_GLOBAL_COUNT";
      const globalOtpCount = parseInt(cache.get(globalOtpKey) || "0", 10) || 0;
      if (globalOtpCount >= OTP_GLOBAL_MAX_PER_WINDOW || MailApp.getRemainingDailyQuota() <= 5) {
        return jsonResponse(genericOtpResp);
      }

      const uSheet = getOrCreateUsersSheet();
      const found = findUserRowFast_(uSheet, identifier);
      if (!found) {
        return jsonResponse(genericOtpResp);
      }

      const targetUsername = String(found.row[0] || "").trim();
      const targetEmail = String(found.row[1] || "").trim();
      const userSalt = String(found.row[3] || "").trim();

      const cooldownKey = "OTP_COOLDOWN_" + shortCacheHash_(targetUsername.toLowerCase());
      if (cache.get(cooldownKey)) {
        // Kembalikan respons generik agar tidak membocorkan status keberadaan akun
        return jsonResponse(genericOtpResp);
      }

      const otpCode = generateSecureOtp6_();
      const otpHash = hashOtp_(otpCode, userSalt);
      const expiryMs = String(Date.now() + 10 * 60 * 1000); // Berlaku 10 menit

      uSheet.getRange(found.rowNum, 6, 1, 2).setNumberFormats([["@", "@"]]).setValues([[otpHash, expiryMs]]);
      cache.put(cooldownKey, "1", 60);
      cache.put(globalOtpKey, String(globalOtpCount + 1), OTP_GLOBAL_WINDOW_SEC);
      cache.remove("OTP_FAILS_" + targetUsername.toLowerCase());

      const safeDisplayUsername = escapeHtmlServer_(targetUsername);
      const emailHtml = [
        "<div style='font-family:Segoe UI,Arial,sans-serif;max-width:480px;margin:0 auto;padding:24px;border:1px solid #e2e8f0;border-radius:12px;background:#ffffff;color:#1e293b;'>",
        "<h2 style='margin:0 0 12px 0;color:#1e3a8a;font-size:18px;'>Reset Password — Comparison Maker</h2>",
        "<p style='font-size:14px;line-height:1.5;margin:0 0 16px 0;'>Halo <strong>" + safeDisplayUsername + "</strong>,</p>",
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

      return jsonResponse(genericOtpResp);
    }

    // 0D. ACTION: AUTH_RESET_PASSWORD (Verifikasi OTP & Simpan Password Baru)
    if (payload.action === "auth_reset_password") {
      const identifier = String(payload.identifier || "").trim().toLowerCase();
      const otpInput = String(payload.otp || "").trim();
      const newPassword = String(payload.newPassword || "");

      if (!identifier || !otpInput || !newPassword) {
        return jsonResponse({ status: "error", message: "Lengkapi Kode OTP dan Password Baru." });
      }
      if (newPassword.length < 6 || newPassword.length > MAX_PASSWORD_LENGTH) {
        return jsonResponse({ status: "error", message: "Password baru harus antara 6 hingga " + MAX_PASSWORD_LENGTH + " karakter." });
      }

      const uSheet = getOrCreateUsersSheet();
      const found = findUserRowFast_(uSheet, identifier);
      const genericOtpInvalidMsg = "Kode OTP tidak valid atau sudah kedaluwarsa.";

      if (!found) {
        return jsonResponse({ status: "error", message: genericOtpInvalidMsg });
      }

      const targetUsername = String(found.row[0] || "").trim();
      const targetEmail = String(found.row[1] || "").trim();
      const userSalt = String(found.row[3] || "").trim();
      const savedOtp = String(found.row[5] || "").trim();
      const savedExpiry = Number(found.row[6] || 0);

      if (!savedOtp || Date.now() > savedExpiry) {
        uSheet.getRange(found.rowNum, 6, 1, 2).setNumberFormats([["@", "@"]]).setValues([["", ""]]);
        return jsonResponse({ status: "error", message: genericOtpInvalidMsg });
      }

      const expectedOtpHash = hashOtp_(otpInput, userSalt);
      const isOtpMatch = savedOtp.indexOf("otp_sha256$") === 0
        ? constantTimeEquals_(savedOtp, expectedOtpHash)
        : constantTimeEquals_(savedOtp, otpInput);

      if (!isOtpMatch) {
        const cache = CacheService.getScriptCache();
        const failKey = "OTP_FAILS_" + targetUsername.toLowerCase();
        const failCount = (parseInt(cache.get(failKey) || "0", 10) || 0) + 1;
        if (failCount >= 5) {
          uSheet.getRange(found.rowNum, 6, 1, 2).setNumberFormats([["@", "@"]]).setValues([["", ""]]);
          cache.remove(failKey);
        } else {
          cache.put(failKey, String(failCount), 600);
        }
        return jsonResponse({ status: "error", message: genericOtpInvalidMsg });
      }

      CacheService.getScriptCache().remove("OTP_FAILS_" + targetUsername.toLowerCase());
      const newSalt = Utilities.getUuid();
      const newHash = hashPasswordWithSalt(newPassword, newSalt, PASSWORD_HASH_ITERATIONS);
      const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT+7", "yyyy-MM-dd HH:mm:ss");

      uSheet.getRange(found.rowNum, 3, 1, 2).setNumberFormats([["@", "@"]]).setValues([[newHash, newSalt]]);
      uSheet.getRange(found.rowNum, 6, 1, 3).setNumberFormats([["@", "@", "@"]]).setValues([["", "", nowStr]]);

      revokeAllUserSessions_(targetUsername);
      const sessionObj = issueSessionToken_(targetUsername, targetEmail, SESSION_DEFAULT_TTL_MS);

      return jsonResponse({
        status: "success",
        action: "auth_reset_password",
        sessionToken: sessionObj.token,
        expiresAt: sessionObj.expiresAt,
        user: {
          username: targetUsername,
          email: targetEmail
        }
      });
    }

    // ====================================================
    // 1. ACTION: UPLOAD_DOCUMENT (Simpan ke Google Drive langsung tanpa Sheet 'Documents')
    // ====================================================
    if (payload.action === "upload_document") {
      const rawBase64 = String(payload.base64Data || "");
      if (!rawBase64) {
        return jsonResponse({ status: "error", message: "Data file kosong." });
      }
      if (rawBase64.length > MAX_UPLOAD_BASE64_CHARS) {
        return jsonResponse({ status: "error", message: "Ukuran file melebihi batas maksimal 25 MB." });
      }

      const rootFolder = getOrCreateRootDocFolder();
      const cleanProdName = String(payload.productName || "General").trim();
      const isTableImgUpload = String(payload.docType || "").trim() === "__TABLE_IMAGE__";
      if (/^_pdf_preview_cache$/i.test(cleanProdName) || (!isTableImgUpload && /^__table_images__$/i.test(cleanProdName))) {
        return jsonResponse({ status: "error", message: "Nama folder dokumen tidak diizinkan." });
      }

      // Buat atau gunakan subfolder per produk di Google Drive (tetap privat, diakses melalui backend terautentikasi)
      const prodFolders = rootFolder.getFoldersByName(cleanProdName);
      const prodFolder = prodFolders.hasNext() ? prodFolders.next() : rootFolder.createFolder(cleanProdName);

      // Decode base64 dan simpan file ke Drive (validasi MIME type terhadap allowlist aman)
      const ALLOWED_UPLOAD_MIMES_ = {
        "application/pdf": true,
        "application/msword": true,
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document": true,
        "application/vnd.ms-excel": true,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": true,
        "application/vnd.ms-powerpoint": true,
        "application/vnd.openxmlformats-officedocument.presentationml.presentation": true,
        "image/png": true,
        "image/jpeg": true,
        "image/jpg": true,
        "image/webp": true,
        "image/gif": true,
        "text/plain": true,
        "text/csv": true,
        "application/octet-stream": true
      };
      const reqMime = String(payload.mimeType || "application/octet-stream").trim().toLowerCase();
      if (!ALLOWED_UPLOAD_MIMES_[reqMime]) {
        return jsonResponse({ status: "error", message: "Tipe file (MIME) tidak diizinkan untuk diunggah." });
      }
      const cleanFileName = String(payload.fileName || "dokumen").trim();
      if (/\.(html?|xhtml|svg|js|mjs|exe|bat|cmd|sh|ps1|vbs|scr)$/i.test(cleanFileName)) {
        return jsonResponse({ status: "error", message: "Ekstensi file tidak diizinkan untuk diunggah." });
      }
      const decodedBytes = Utilities.base64Decode(rawBase64);
      const blob = Utilities.newBlob(decodedBytes, reqMime, cleanFileName);
      const file = prodFolder.createFile(blob);

      // Simpan tipe dokumen ke metadata description file di Drive
      if (payload.docType) {
        file.setDescription(String(payload.docType));
      }

      const fileId = file.getId();
      const viewUrl = file.getUrl();
      const downloadUrl = "https://drive.google.com/uc?export=download&id=" + fileId;
      const fileSize = file.getSize();

      // Hapus tab sheet 'Documents' legacy maksimal 1x per 6 jam agar aksi Drive tidak membuka SpreadsheetApp tiap kali
      try {
        const legacyCache = CacheService.getScriptCache();
        if (!legacyCache.get("LEGACY_DOCS_SHEET_CLEANED")) {
          const oldDocSheet = getSs().getSheetByName("Documents");
          if (oldDocSheet) {
            try { getSs().deleteSheet(oldDocSheet); } catch (e) {}
          }
          legacyCache.put("LEGACY_DOCS_SHEET_CLEANED", "1", 21600);
        }
      } catch (e) {}

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
          const file = assertAccessibleProductFile_(payload.fileId);
          const parents = file.getParents();
          while (parents.hasNext()) {
            parentFoldersToCheck.push(parents.next());
          }
          secureTrashDriveFile_(file);
        } catch (e) {
          console.error("Gagal trash file di Drive:", e);
          return jsonResponse({ status: "error", message: "Gagal menghapus dokumen di Google Drive." });
        }
      }

      // Jika seluruh dokumen di dalam folder produk telah terhapus (kosong), hapus juga foldernya
      parentFoldersToCheck.forEach(function(parentFolder) {
        try {
          const pfNameLower = parentFolder ? String(parentFolder.getName() || "").trim().toLowerCase() : "";
          if (
            parentFolder &&
            parentFolder.getId() !== rootFolderId &&
            pfNameLower !== "__table_images__" &&
            pfNameLower !== "_pdf_preview_cache" &&
            isFolderCompletelyEmpty(parentFolder)
          ) {
            try { parentFolder.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (se) {}
            parentFolder.setTrashed(true);
          }
        } catch (e) {
          console.error("Gagal mengecek/menghapus folder kosong:", e);
        }
      });

      // Fallback: jika productName dikirim dan foldernya kosong, hapus foldernya (kecuali folder internal)
      if (payload.productName && rootFolder) {
        try {
          const targetProd = String(payload.productName).trim().toLowerCase();
          if (targetProd && targetProd !== "__table_images__" && targetProd !== "_pdf_preview_cache") {
            const subFolders = rootFolder.getFolders();
            while (subFolders.hasNext()) {
              const sf = subFolders.next();
              if (!sf.isTrashed() && sf.getName().trim().toLowerCase() === targetProd && isFolderCompletelyEmpty(sf)) {
                try { sf.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (se) {}
                sf.setTrashed(true);
              }
            }
          }
        } catch (e) {}
      }

      // Hapus tab sheet 'Documents' legacy maksimal 1x per 6 jam agar aksi Drive tidak membuka SpreadsheetApp tiap kali
      try {
        const legacyCache = CacheService.getScriptCache();
        if (!legacyCache.get("LEGACY_DOCS_SHEET_CLEANED")) {
          const oldDocSheet = getSs().getSheetByName("Documents");
          if (oldDocSheet) {
            try { getSs().deleteSheet(oldDocSheet); } catch (e) {}
          }
          legacyCache.put("LEGACY_DOCS_SHEET_CLEANED", "1", 21600);
        }
      } catch (e) {}

      return jsonResponse({ status: "success", action: "delete_document" });
    }

    // ====================================================
    // 3. ACTION: GET_FILE_BASE64 (Untuk Ekspor ZIP & Preview Inline)
    //    asPdf: true -> file Word/PowerPoint dikonversi ke PDF khusus untuk preview
    // ====================================================
    if (payload.action === "get_file_base64" && payload.fileId) {
      const file = assertAccessibleProductFile_(payload.fileId);
      const asPdf = payload.asPdf === true || payload.asPdf === "true";
      let blob = null;
      let mimeType = "";
      let pdfConversionError = "";

      if (asPdf) {
        try {
          blob = convertOfficeFileToPdfBlob(file);
          mimeType = "application/pdf";
        } catch (convErr) {
          console.warn("Konversi PDF gagal, mengembalikan file asli:", String(convErr && convErr.message ? convErr.message : convErr));
          pdfConversionError = "Konversi preview PDF di server tidak tersedia; menampilkan file asli.";
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
      const itemsBySheet = {};

      for (let idx = 0; idx < payload.items.length; idx++) {
        const item = payload.items[idx];
        if (!item || isReservedSystemSheet_(item.sheetName)) {
          return jsonResponse({ status: "error", message: "Akses ditolak: sheet sistem tidak dapat dihapus/dimodifikasi via aksi delete." });
        }
        const sName = String(item.sheetName || "").trim();
        const targetName = String(item.productName || "").trim().toLowerCase();
        if (!sName || !targetName) continue;
        deletedProductNamesLower[targetName] = true;
        if (!itemsBySheet[sName]) itemsBySheet[sName] = {};
        itemsBySheet[sName][targetName] = true;
      }

      Object.keys(itemsBySheet).forEach(function(sName) {
        const sheet = getSs().getSheetByName(sName);
        if (!sheet) return;
        const targetMap = itemsBySheet[sName];
        const dataRange = sheet.getDataRange().getValues();

        for (let r = dataRange.length - 1; r >= 1; r--) {
          const rowProdName = String(dataRange[r][0] || "").trim().toLowerCase();
          if (targetMap[rowProdName]) {
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
            affectedSheets[sName] = sheet;
          }
        }
      });

      // 1. Hapus file Gambar Tabel di Google Drive jika dicentang (validasi ketat: hanya file __TABLE_IMAGE__ di dalam folder __TABLE_IMAGES__)
      if (deleteTableImages) {
        Object.keys(tableImageIdsToDelete).forEach(function(imgFileId) {
          try {
            const imgFile = assertAccessibleProductFile_(imgFileId);
            if (String(imgFile.getDescription() || "").trim() !== "__TABLE_IMAGE__") {
              console.warn("Abaikan hapus tableImageFileId karena bukan __TABLE_IMAGE__:", imgFileId);
              return;
            }
            let inTableImagesFolder = false;
            const imgParents = imgFile.getParents();
            while (imgParents.hasNext()) {
              if (String(imgParents.next().getName() || "").trim() === "__TABLE_IMAGES__") {
                inTableImagesFolder = true;
                break;
              }
            }
            if (!inTableImagesFolder) {
              console.warn("Abaikan hapus tableImageFileId karena berada di luar folder __TABLE_IMAGES__:", imgFileId);
              return;
            }
            secureTrashDriveFile_(imgFile);
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
              if (folderNameLower === "__table_images__" || folderNameLower === "_pdf_preview_cache") continue;

              if (deletedProductNamesLower[folderNameLower]) {
                const files = sub.getFiles();
                while (files.hasNext()) {
                  try {
                    secureTrashDriveFile_(files.next());
                  } catch (fErr) {}
                }
                try { sub.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.NONE); } catch (se) {}
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

      invalidateSummaryCache_();
      return jsonResponse({ status: "success", action: "delete" });
    }

    // ====================================================
    // 4b. ACTION: SIMPAN ALIAS NAMA MANFAAT KE CONFIG (dari Review Benefit saat Upload Excel)
    // ====================================================
    if (payload.action === "save_benefit_aliases" && Array.isArray(payload.mappings)) {
      const addedCount = saveBenefitAliases_(getSs(), payload.mappings);
      invalidateSummaryCache_();
      return jsonResponse({ status: "success", action: "save_benefit_aliases", added: addedCount });
    }

    if (payload.action === "manage_benefit_aliases") {
      const result = manageBenefitAliases_(getSs(), payload);
      return jsonResponse({
        status: "success",
        action: "manage_benefit_aliases",
        entries: result.entries,
        config: result.config || null
      });
    }

    // ====================================================
    // 4c. ACTION: SIMPAN PENGATURAN SHEET / KATEGORI (General Provisions, Special Provisions, Product Sections)
    // ====================================================
    if (payload.action === "save_sheet_category_config") {
      const updatedConfig = saveSheetCategoryConfig_(getSs(), payload);
      return jsonResponse({
        status: "success",
        action: "save_sheet_category_config",
        sheetName: payload.sheetName,
        config: updatedConfig
      });
    }

    if (payload.action !== "append" && payload.action !== "update") {
      return jsonResponse({ status: "error", message: "Action tidak dikenali" });
    }

    if (isReservedSystemSheet_(payload.sheetName)) {
      return jsonResponse({ status: "error", message: "Akses ditolak: sheet sistem terproteksi (" + payload.sheetName + ")." });
    }

    let sheet = getSs().getSheetByName(payload.sheetName);
    if (!sheet && payload.sheetName) {
      sheet = getSs().insertSheet(String(payload.sheetName).trim());
    }
    if (!sheet) {
      return jsonResponse({ status: "error", message: "Sheet tidak ditemukan: " + payload.sheetName });
    }

    // ====================================================
    // 5. ACTION: APPEND (Upload Batch dari File Excel)
    // ====================================================
    if (payload.action === "append") {
      if (!Array.isArray(payload.data) || payload.data.length === 0) {
        return jsonResponse({ status: "error", message: "Data append kosong atau tidak valid." });
      }
      if (payload.data.length > MAX_APPEND_ROWS) {
        return jsonResponse({ status: "error", message: "Jumlah baris melebihi batas maksimal (" + MAX_APPEND_ROWS + ")." });
      }

      // Batch write seluruh baris sekaligus dengan satu panggilan setValues() + sanitasi anti formula injection
      let rowsToWrite = [];
      let colCount = 0;

      if (payload.headers && Array.isArray(payload.headers)) {
        const sheetHeaders = syncSheetHeaders(sheet, payload.headers);
        colCount = sheetHeaders.length;
        rowsToWrite = payload.data.map(function(row) {
          const alignedRow = alignRowToSheetHeaders(payload.headers, row, sheetHeaders);
          return alignedRow.map(function(cell) {
            let cellStr = (cell === undefined || cell === null) ? "-" : String(cell);
            if (cellStr.length > MAX_CELL_CHARS) {
              cellStr = cellStr.substring(0, MAX_CELL_CHARS - 15) + "...[TRUNCATED]";
            }
            return sanitizeSheetCell_(cellStr);
          });
        });
      } else {
        rowsToWrite = payload.data.map(function(row) {
          const arr = Array.isArray(row) ? row : [];
          if (arr.length > colCount) colCount = arr.length;
          return arr.map(function(cell) {
            let cellStr = (cell === undefined || cell === null) ? "-" : String(cell).trim();
            if (cellStr.length > MAX_CELL_CHARS) {
              cellStr = cellStr.substring(0, MAX_CELL_CHARS - 15) + "...[TRUNCATED]";
            }
            return sanitizeSheetCell_(cellStr);
          });
        });
        rowsToWrite = rowsToWrite.map(function(r) {
          while (r.length < colCount) r.push("-");
          return r;
        });
      }

      if (rowsToWrite.length > 0 && colCount > 0) {
        const startRow = sheet.getLastRow() + 1;
        const formatsToWrite = rowsToWrite.map(function() {
          return new Array(colCount).fill("@");
        });
        sheet.getRange(startRow, 1, rowsToWrite.length, colCount)
          .setNumberFormats(formatsToWrite)
          .setValues(rowsToWrite);
      }

      invalidateSummaryCache_();

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
      invalidateSummaryCache_();

      return jsonResponse({ status: "success", action: "update" });
    }

    return jsonResponse({ status: "error", message: "Action tidak dikenali" });

  } catch (err) {
    console.error("doPost error:", err);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Terjadi kesalahan internal saat memproses permintaan." }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (lock) {
      try { lock.releaseLock(); } catch (e) {}
    }
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

// Folder cache PDF preview di-scope di dalam rootFolder agar terisolasi dari Google Drive pribadi pemilik
function getPdfPreviewCacheFolder_() {
  const rootFolder = getOrCreateRootDocFolder();
  const folderName = "_pdf_preview_cache";
  const it = rootFolder.getFoldersByName(folderName);
  if (it.hasNext()) return it.next();
  return rootFolder.createFolder(folderName);
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

// Hapus SEMUA file di folder "_pdf_preview_cache" tanpa menunggu umur 24 jam.
// Pasang fungsi ini di Apps Script Triggers (Time-driven -> Day timer -> Midnight to 1am).
function clearAllPdfPreviewCacheDaily() {
  return cleanupPdfPreviewCache(0);
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
