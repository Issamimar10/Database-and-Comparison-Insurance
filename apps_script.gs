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
        return "'" + escaped + "'!A1:ZZ1000";
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
  ["Setting", "Invite Code", "IDCM2026", "Kode akses rahasia untuk pendaftaran akun baru (Sign Up)"],

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
  console.log("Sheet Config berhasil dilengkapi dengan struktur 4 kolom (Kategori | Nama Utama | Nilai / Nama Lain | Keterangan)!");
}

// Helper Global: Ambil seluruh konfigurasi aman & data produk dari Master Spreadsheet
// (Memungkinkan Google Spreadsheet dikunci 'Restricted / Private' tanpa perlu dibuka ke publik via /gviz/tq)
// Parameter mode: "summary" (ringkas untuk loading awal cepat) atau "full" (default, lengkap).
function buildAllProductDataResponse(mode) {
  const isSummary = String(mode || "").trim().toLowerCase() === "summary";
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
  let driveFolderId = "";

  const configSheet = ensureDefaultConfigSheet_(ss, false);
  if (configSheet && configSheet.getLastRow() > 0 && configSheet.getLastColumn() > 0) {
    const cfgValues = configSheet.getDataRange().getDisplayValues();
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

  return {
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
      driveFolderId: driveFolderId
    },
    sheets: sheetsData
  };
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
        rowRange.setValues([[r.entryType, r.canonical, aliasStr, r.description || getDefaultConfigDescription_(r.entryType, r.canonical)]]);
        return;
      }
      const cell = configSheet.getRange(rowNum, 3);
      cell.setNumberFormat("@");
      cell.setValue(aliasStr);
      if (!r.description) {
        configSheet.getRange(rowNum, 4).setValue(getDefaultConfigDescription_(r.entryType, r.canonical));
      }
    });
  });

  return added;
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
  // 0. Ambil seluruh data produk & konfigurasi aman langsung via Apps Script (Tanpa perlu /gviz/tq publik)
  if (e && e.parameter && e.parameter.action === "get_all_product_data") {
    try {
      const result = buildAllProductDataResponse(e.parameter.mode);
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
      const docResult = listAllProductDocumentsFast_();
      return ContentService.createTextOutput(JSON.stringify({
        status: "success",
        documents: docResult.documents,
        folders: docResult.folders
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
        const configSheet = ss.getSheetByName("Config");
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
              cat === "biaya (dplk)" ||
              cat === "dplk charge"
            ) {
              if (baseProtected.indexOf(name) === -1) {
                baseProtected.push(name);
              }
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
          const allSheets = ss.getSheets();
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

          const configSheet = ss.getSheetByName("Config");
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

    function jsonResponse(obj) {
      return ContentService.createTextOutput(JSON.stringify(obj))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 00. ACTION: GET_ALL_PRODUCT_DATA (Via POST)
    if (payload.action === "get_all_product_data") {
      return jsonResponse(buildAllProductDataResponse(payload.mode));
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
        configSheet.appendRow(["Setting", "Invite Code", DEFAULT_INVITE_CODE, "Kode akses rahasia untuk pendaftaran akun baru (Sign Up)"]);
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
