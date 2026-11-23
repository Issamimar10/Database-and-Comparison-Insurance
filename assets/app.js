const SHEET_ID = "1NF4sv1CT_C4Vv6tJC5ppplVir3abi8_9GYHAwhdM8YI";
const DEFAULT_SHEET_TABS = ["Conventional", "Sharia", "Health & Surgical", "DPLK"];
let SHEET_TABS = [...DEFAULT_SHEET_TABS];
const APPS_SCRIPT_URL = "/api";

function syncSheetTabSelectors() {
    const optionsHtml = SHEET_TABS
        .map(tab => `<option value="${escapeHtmlAttr(tab)}">${escapeHtml(tab)}</option>`)
        .join('');

    ["uploadTargetSheet", "manual_category", "edit_category"].forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        const prevVal = el.value;
        el.innerHTML = optionsHtml;
        if (prevVal && SHEET_TABS.includes(prevVal)) {
            el.value = prevVal;
        } else if (SHEET_TABS.length > 0) {
            el.value = SHEET_TABS[0];
        }
    });
}

let products = [];
let selectedProductIds = [];
let currentlyVisibleProducts = [];
let dynamicColumns = [];
let sheetContinuationColumns = [];
let currentComparisonFeatures = [];
let currentComparisonProducts = [];

let currentActiveProduct = null;
let isEditMode = false;

let activeFilters = {
    category: [],
    company: [],
    type: [],
    channel: [],
    currency: [],
    year: [],
    underwriting: []
};

const CUSTOM_FILTER_STORAGE_KEY = "idcm_custom_filters";
let customFilterColumns = [];
try {
    const savedCustomFilters = JSON.parse(localStorage.getItem(CUSTOM_FILTER_STORAGE_KEY) || "[]");
    if (Array.isArray(savedCustomFilters)) {
        customFilterColumns = savedCustomFilters.filter(c => typeof c === "string" && c.trim() !== "");
        customFilterColumns.forEach(c => {
            activeFilters[`custom:${c}`] = [];
        });
    }
} catch (e) { }
let dropdownSearchQueries = {};
let addFilterSearchQuery = "";

let parsedUploadRows = [];
let parsedDuplicateRows = [];

let productDocuments = [];
let productDriveFolders = [];
let currentSortOption = "default";
let bulkDocStagingFiles = [];
let systemConfig = {
    driveFolderId: ""
};

// State Staged / Lazy Loading (Summary dulu -> Full & Documents menyusul)
let currentSheetLoadId = 0;
let isFullProductDataLoaded = false;
let fullProductDataPromise = null;
let isDocumentsLoaded = false;
let documentsLoadPromise = null;
let currentLoadUsedLegacyGviz = false;

// Alias nama manfaat & ketentuan non-benefit dari sheet Config
let BENEFIT_ALIAS_MAP = {};            // normalizedKey -> Nama Standar
let BENEFIT_CONFIG_CANONICALS = new Set(); // normalizedKey dari Nama Standar di Config
let BENEFIT_CONFIG_CANONICAL_LIST = []; // daftar Nama Standar Benefit dari Config (case asli)
let NON_BENEFIT_ALIAS_MAP = {};        // normalizedKey -> Nama Standar Non-Benefit
let NON_BENEFIT_CONFIG_CANONICALS = new Set();

// State Review Benefit & Non-Benefit saat Upload Excel
let benefitReviewItems = [];
let benefitReviewDecisions = {};       // decisionKey -> { mode: "pending" | "new" | "ignore", name, kind }
let uploadParsedSchema = [];
let uploadParsedSchemaStructured = false;
let uploadSubmitBaseEnabled = false;

// --- LOGIKA DARK / LIGHT MODE ---
// Ikon (lingkaran <-> bulan sabit) digambar sekali di HTML dan dianimasikan lewat CSS
// mengikuti class "dark" pada <html>. JS hanya memperbarui tooltip tombol.
function updateThemeUI() {
    const isDark = document.documentElement.classList.contains('dark');
    const metaCS = document.querySelector('meta[name="color-scheme"]');
    if (metaCS) {
        const stored = localStorage.getItem('theme');
        metaCS.content = stored === 'dark' ? 'dark' : (stored === 'light' ? 'light' : 'light dark');
    }
    const themeBtn = document.getElementById('themeToggleBtn');
    if (!themeBtn) return;

    themeBtn.title = isDark
        ? "Switch to Light Mode"
        : "Switch to Dark Mode";
}

let themeSwitchTimer = null;
function toggleTheme() {
    const htmlEl = document.documentElement;
    htmlEl.classList.add('theme-switching');
    htmlEl.classList.toggle('dark');
    const isDark = htmlEl.classList.contains('dark');
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
    updateThemeUI();

    if (themeSwitchTimer) clearTimeout(themeSwitchTimer);
    themeSwitchTimer = setTimeout(() => {
        htmlEl.classList.remove('theme-switching');
    }, 60);
}

// --- DATE & DOCUMENT FORMAT HELPERS ---
const INDO_MONTHS = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
];

function excelSerialToDate(serial) {
    const num = Number(serial);
    if (!Number.isFinite(num) || num < 20000 || num > 70000) return null;
    // Excel epoch: 1899-12-30 (offset 25569 hari terhadap Unix epoch 1970-01-01); gunakan UTC agar tidak bergeser ±1 hari karena timezone
    const utc = new Date(Math.round((num - 25569) * 86400 * 1000));
    if (isNaN(utc.getTime())) return null;
    return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
}

function parseDateValue(dateStr) {
    if (!dateStr || dateStr === "-") return null;
    if (dateStr instanceof Date) return isNaN(dateStr.getTime()) ? null : dateStr;
    const clean = String(dateStr).trim();
    // Angka serial tanggal Excel (5 digit: mis. 45313 -> 22 Jan 2024, 46054 -> 1 Feb 2026)
    if (/^\d{5}(?:\.\d+)?$/.test(clean)) {
        const exDate = excelSerialToDate(parseFloat(clean));
        if (exDate) return exDate;
    }
    // Recovery untuk data lama yang sempat tersimpan sebagai "01-01-46054" atau "46054-01-01"
    const corruptedSerial = clean.match(/^(?:0?1[-\/]0?1[-\/](\d{5})|(\d{5})[-\/]0?1[-\/]0?1)$/);
    if (corruptedSerial) {
        const exDate = excelSerialToDate(parseInt(corruptedSerial[1] || corruptedSerial[2], 10));
        if (exDate) return exDate;
    }
    // Google Visualization API Date(yyyy, m, d)
    const gvizDate = clean.match(/^Date\((\d+),\s*(\d+),\s*(\d+)(?:,\s*(\d+),\s*(\d+),\s*(\d+))?\)$/i);
    if (gvizDate) {
        const year = parseInt(gvizDate[1], 10);
        const month = parseInt(gvizDate[2], 10);
        const day = parseInt(gvizDate[3], 10);
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }
    // DD-MM-YYYY or DD/MM/YYYY
    const dmy = clean.match(/^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/);
    if (dmy) {
        const day = parseInt(dmy[1], 10);
        const month = parseInt(dmy[2], 10) - 1;
        const year = parseInt(dmy[3], 10);
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }
    // YYYY-MM-DD
    const ymd = clean.match(/^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/);
    if (ymd) {
        const year = parseInt(ymd[1], 10);
        const month = parseInt(ymd[2], 10) - 1;
        const day = parseInt(ymd[3], 10);
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }
    // Tahun saja (4 digit, mis. "2024")
    if (/^(19|20)\d{2}$/.test(clean)) {
        return new Date(parseInt(clean, 10), 0, 1);
    }
    const ts = Date.parse(clean);
    return isNaN(ts) ? null : new Date(ts);
}

function formatDateToDDMMYYYY(val) {
    if (!val || val === "-") return "-";
    const d = (val instanceof Date) ? val : parseDateValue(String(val));
    if (!d) return String(val).trim() || "-";
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${day}-${month}-${year}`;
}

function formatDateForInput(val) {
    if (!val || val === "-") return "";
    const d = (val instanceof Date) ? val : parseDateValue(String(val));
    if (!d) return "";
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const year = d.getFullYear();
    return `${year}-${month}-${day}`;
}

function formatDisplayDate(val) {
    if (!val || val === "-") return "-";
    const d = (val instanceof Date) ? val : parseDateValue(String(val));
    if (!d) return String(val);
    return `${d.getDate()} ${INDO_MONTHS[d.getMonth()]} ${d.getFullYear()}`;
}

function formatFileSize(bytes) {
    const num = Number(bytes);
    if (isNaN(num) || num <= 0) return "";
    if (num < 1024) return num + " B";
    if (num < 1024 * 1024) return (num / 1024).toFixed(1) + " KB";
    return (num / (1024 * 1024)).toFixed(1) + " MB";
}

function getDocFileIcon(fileName) {
    const ext = String(fileName || "").split('.').pop().toLowerCase();
    if (ext === 'pdf') return { icon: 'fa-file-pdf', color: 'text-red-500 bg-red-50 dark:bg-red-950/40 border-red-200 dark:border-red-800' };
    if (['doc', 'docx'].includes(ext)) return { icon: 'fa-file-word', color: 'text-blue-500 bg-blue-50 dark:bg-neutral-800/80 border-blue-200 dark:border-neutral-600' };
    if (['xls', 'xlsx', 'csv'].includes(ext)) return { icon: 'fa-file-excel', color: 'text-emerald-500 bg-emerald-50 dark:bg-emerald-950/40 border-emerald-200 dark:border-emerald-800' };
    if (['ppt', 'pptx'].includes(ext)) return { icon: 'fa-file-powerpoint', color: 'text-amber-500 bg-amber-50 dark:bg-amber-950/40 border-amber-200 dark:border-amber-800' };
    if (['png', 'jpg', 'jpeg', 'webp', 'gif'].includes(ext)) return { icon: 'fa-file-image', color: 'text-purple-500 bg-purple-50 dark:bg-purple-950/40 border-purple-200 dark:border-purple-800' };
    return { icon: 'fa-file-lines', color: 'text-gray-500 bg-gray-50 dark:bg-neutral-800 border-gray-200 dark:border-neutral-700' };
}

function getDocumentsForProduct(productName) {
    if (!productName) return [];
    const cleanTarget = String(productName).trim().toLowerCase();
    return productDocuments.filter(d => String(d.productName || "").trim().toLowerCase() === cleanTarget);
}

function setSortOption(optionKey, e) {
    if (e) {
        e.preventDefault();
        e.stopPropagation();
    }
    currentSortOption = optionKey;
    const labels = {
        default: "Sort",
        launch_desc: "Newest Release",
        launch_asc: "Oldest Release",
        name_asc: "Name (A-Z)",
        name_desc: "Name (Z-A)",
        company_asc: "Company (A-Z)"
    };
    const btnText = document.getElementById("sortBtnText");
    if (btnText) btnText.textContent = labels[optionKey] || "Sort";

    ["default", "launch_desc", "launch_asc", "name_asc", "company_asc"].forEach(k => {
        const checkEl = document.getElementById(`check_sort_${k}`);
        if (checkEl) {
            if (k === optionKey) checkEl.classList.remove("hidden");
            else checkEl.classList.add("hidden");
        }
    });

    closeAllDropdowns();
    filterData();
}

// --- HELPER: SANITIZE N/A TO DASH ---
function sanitizeValue(val) {
    if (!val) return "-";
    let str = String(val).trim();
    if (/^n\/?a\.?$/i.test(str) || str === "" || str === "null" || str === "undefined") {
        return "-";
    }
    return str;
}

function isEmptyVal(val) {
    const clean = sanitizeValue(val);
    return clean === "-";
}

// --- FULL-RESOLUTION TABLE IMAGE CACHE & LOADER (GOOGLE DRIVE) ---
const driveImageCache = {};
const driveImagePendingPromises = {};
const pendingAttachmentUploads = new WeakMap();
let hydrateDriveImagesTimer = null;

async function fetchDriveImageDataUrl(fileId) {
    if (!fileId) return "";
    if (driveImageCache[fileId]) return driveImageCache[fileId];
    if (driveImagePendingPromises[fileId]) return driveImagePendingPromises[fileId];

    const promise = (async () => {
        try {
            const resp = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(withSessionToken({
                    action: "get_file_base64",
                    fileId: fileId,
                    asPdf: false
                }))
            });
            const data = await resp.json();
            if (handleUnauthorizedResponse(data)) return "";
            if (data && data.status === "success" && data.base64) {
                const mime = data.mimeType && data.mimeType.startsWith("image/") ? data.mimeType : "image/png";
                const dataUrl = `data:${mime};base64,${data.base64}`;
                driveImageCache[fileId] = dataUrl;
                return dataUrl;
            }
        } catch (err) {
            console.warn("Failed to load full-resolution image from Google Drive:", fileId, err);
        } finally {
            delete driveImagePendingPromises[fileId];
        }
        return "";
    })();

    driveImagePendingPromises[fileId] = promise;
    return promise;
}

function hydrateDriveImagesInDOM(root = document) {
    const imgs = root.querySelectorAll('img[data-drive-img-id]');
    imgs.forEach(imgEl => {
        const fileId = imgEl.getAttribute('data-drive-img-id');
        if (!fileId) return;
        if (driveImageCache[fileId]) {
            imgEl.src = driveImageCache[fileId];
            imgEl.classList.remove('hidden');
            imgEl.removeAttribute('data-drive-img-id');
            const loader = imgEl.parentElement ? imgEl.parentElement.querySelector('.drive-img-loader') : null;
            if (loader) loader.remove();
            return;
        }
        if (imgEl.dataset.hydrating === "1") return;
        imgEl.dataset.hydrating = "1";

        fetchDriveImageDataUrl(fileId).then(dataUrl => {
            imgEl.dataset.hydrating = "0";
            if (!dataUrl) return;
            const safeSelectorId = (typeof CSS !== "undefined" && typeof CSS.escape === "function") ? CSS.escape(fileId) : String(fileId).replace(/[^a-zA-Z0-9_-]/g, "");
            document.querySelectorAll(`img[data-drive-img-id="${safeSelectorId}"]`).forEach(targetImg => {
                targetImg.src = dataUrl;
                targetImg.classList.remove('hidden');
                targetImg.removeAttribute('data-drive-img-id');
                const loader = targetImg.parentElement ? targetImg.parentElement.querySelector('.drive-img-loader') : null;
                if (loader) loader.remove();
            });
        });
    });
}

function scheduleHydrateDriveImages() {
    if (hydrateDriveImagesTimer) return;
    hydrateDriveImagesTimer = requestAnimationFrame(() => {
        hydrateDriveImagesTimer = null;
        hydrateDriveImagesInDOM();
    });
}

// Helper to parse DPLK Benefit cell value ("✓", "✅", or "✓ | Remark text")
function parseDplkBenefitValue(rawVal) {
    const str = String(rawVal || "").trim();
    if (!str || str === "-") return { isChecked: false, checked: false, remark: "" };
    if (str === "✓" || str === "✅") return { isChecked: true, checked: true, remark: "" };
    const match = str.match(/^(?:✓|✅)(?:\s*\|\s*|\s*:\s*|\r?\n|\s+)([\s\S]+)$/);
    if (match) {
        const rem = (match[1] || "").trim();
        return { isChecked: true, checked: true, remark: rem };
    }
    return { isChecked: false, checked: false, remark: "" };
}

// Converts lines starting with "-" into bullet points while keeping mid-sentence "-" intact
function formatBulletLinesHtml(rawText) {
    if (!rawText || rawText === "-") return "-";
    const lines = String(rawText).split(/\r\n|\r|\n/);
    return lines.map(line => {
        const bulletMatch = line.match(/^\s*-\s*(.+)$/);
        if (bulletMatch) {
            return `<div class="flex items-start gap-1.5 my-0.5"><span class="select-none font-bold leading-snug">&bull;</span><span class="flex-1">${escapeHtml(bulletMatch[1])}</span></div>`;
        }
        if (line.trim() === "") {
            return `<div class="h-2"></div>`;
        }
        return `<div>${escapeHtml(line)}</div>`;
    }).join("");
}

function formatListText(val) {
    const clean = sanitizeValue(val);
    if (clean === "-") return "-";

    const images = [];
    const textWithoutImgs = clean
        .replace(/\[IMG_ID:([a-zA-Z0-9_-]+)\]/gi, (match, fileId) => {
            images.push(`drive:${fileId}`);
            return "";
        })
        .replace(/\[IMG:(data:image\/[^\]]+|https?:\/\/[^\]]+|drive:[a-zA-Z0-9_-]+)\]/gi, (match, src) => {
            images.push(src);
            return "";
        })
        .replace(/\n{3,}/g, "\n\n").trim();

    let html = "";
    const dplkParsed = parseDplkBenefitValue(textWithoutImgs);
    if (dplkParsed.checked && images.length === 0) {
        if (!dplkParsed.remark) {
            return `<span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-emerald-100 dark:bg-emerald-900/60 text-emerald-600 dark:text-emerald-400 font-bold shadow-sm" title="Available"><i class="fa-solid fa-check"></i></span>`;
        }
        return `
                    <div class="inline-flex items-center gap-2.5 text-left">
                        <span class="inline-flex items-center justify-center w-6 h-6 rounded-full bg-emerald-100 dark:bg-emerald-900/60 text-emerald-600 dark:text-emerald-400 font-bold shadow-sm shrink-0" title="Available"><i class="fa-solid fa-check"></i></span>
                        <div class="text-gray-800 dark:text-neutral-200 leading-relaxed">${formatBulletLinesHtml(dplkParsed.remark)}</div>
                    </div>
                `;
    }
    if (textWithoutImgs && textWithoutImgs !== "-") {
        html += `<div class="text-left">${formatBulletLinesHtml(textWithoutImgs)}</div>`;
    }

    if (images.length > 0) {
        let needsHydration = false;
        const imgsHtml = images.map((src, idx) => {
            const isDrive = String(src).startsWith("drive:");
            const driveFileId = isDrive ? String(src).slice(6) : "";
            const resolvedSrc = isDrive ? (driveImageCache[driveFileId] || "") : src;
            const isUnresolvedDrive = isDrive && !resolvedSrc;
            if (isUnresolvedDrive) needsHydration = true;

            const rawActionSrc = isDrive ? `drive:${driveFileId}` : src;
            const safeActionSrc = escapeHtmlAttr(rawActionSrc);
            const safeDisplaySrc = escapeHtmlAttr(resolvedSrc || "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");

            return `
                        <div class="relative inline-flex flex-col items-center group cursor-zoom-in rounded-lg overflow-hidden border border-gray-200 dark:border-neutral-600 bg-white p-1 shadow-sm hover:shadow-md transition my-1 w-fit max-w-full"
                             data-img-action="lightbox" data-img-src="${safeActionSrc}" title="Click to enlarge table image">
                            ${isUnresolvedDrive ? `
                                <div class="drive-img-loader flex items-center gap-2 px-4 py-5 text-xs text-blue-600 dark:text-blue-400 font-medium bg-gray-50 dark:bg-neutral-800/60 rounded">
                                    <i class="fa-solid fa-circle-notch fa-spin"></i> Loading full-resolution image...
                                </div>
                            ` : ''}
                            <img src="${safeDisplaySrc}" ${isUnresolvedDrive ? `data-drive-img-id="${escapeHtmlAttr(driveFileId)}"` : ''} alt="Product Table ${idx + 1}" class="${isUnresolvedDrive ? 'hidden ' : ''}w-auto max-w-full h-auto block rounded" style="image-rendering: -webkit-optimize-contrast;" loading="lazy">
                            <div class="absolute top-1.5 right-1.5 opacity-90 group-hover:opacity-100 transition z-10">
                                <button type="button" data-img-action="copy" data-img-src="${safeActionSrc}"
                                    class="bg-slate-900/85 hover:bg-emerald-600 text-white text-[10px] font-semibold px-2 py-0.5 rounded-md transition flex items-center gap-1 shadow cursor-pointer" title="Copy image to clipboard">
                                    <i class="fa-solid fa-copy text-emerald-300"></i> Copy
                                </button>
                            </div>
                            <div class="absolute bottom-1.5 right-1.5 bg-slate-900/80 text-white text-[10px] px-2 py-0.5 rounded-md opacity-85 group-hover:opacity-100 transition flex items-center gap-1 pointer-events-none shadow">
                                <i class="fa-solid fa-magnifying-glass-plus text-blue-300"></i> Enlarge
                            </div>
                        </div>
                    `;
        }).join("");
        html += `<div class="flex flex-col items-start gap-2 ${textWithoutImgs ? 'mt-2' : ''}">${imgsHtml}</div>`;
        if (needsHydration) scheduleHydrateDriveImages();
    }

    return html || "-";
}

// Separates plain text and [IMG:...] / [IMG_ID:...] tokens from an attribute string
function splitTextAndImages(rawVal) {
    if (!rawVal || rawVal === "-") return { text: "", images: [] };
    const images = [];
    const textOnly = String(rawVal)
        .replace(/\[IMG_ID:([a-zA-Z0-9_-]+)\]/gi, (match, fileId) => {
            images.push(`drive:${fileId}`);
            return "";
        })
        .replace(/\[IMG:(data:image\/[^\]]+|https?:\/\/[^\]]+|drive:[a-zA-Z0-9_-]+)\]/gi, (match, src) => {
            images.push(src);
            return "";
        })
        .replace(/\n{3,}/g, "\n\n").trim();
    return {
        text: textOnly === "-" ? "" : textOnly,
        images
    };
}

// Strips [IMG:...] tokens when exporting to Excel so cells do not contain raw Base64 strings
function stripImageTokensForExcel(rawVal) {
    if (!rawVal || rawVal === "-") return "-";
    const { text, images } = splitTextAndImages(rawVal);
    if (images.length === 0) return sanitizeValue(rawVal);
    const imgNote = images.length === 1
        ? "[Table Image - View on Web]"
        : `[${images.length} Table Images - View on Web]`;
    return text ? `${text}\n${imgNote}` : imgNote;
}

// --- TABLE IMAGE LIGHTBOX PREVIEW (STATIC FRAME & ENLARGE TOGGLE LIKE COMPARISON MODAL) ---
let currentLightboxScale = 1;
let defaultFitLightboxScale = 1;
let isLightboxModalEnlarged = false;

function setEnlargeLightboxModalState(enlarged) {
    isLightboxModalEnlarged = Boolean(enlarged);
    const modal = document.getElementById("imageLightboxModal");
    const dialog = document.getElementById("imageLightboxDialog");
    const icon = document.getElementById("iconEnlargeLightboxModal");
    const text = document.getElementById("textEnlargeLightboxModal");
    const btn = document.getElementById("btnEnlargeLightboxModal");
    if (!modal || !dialog) return;

    if (isLightboxModalEnlarged) {
        modal.classList.remove("p-4");
        modal.classList.add("p-1", "sm:p-2");
        dialog.classList.remove("max-w-6xl", "h-[88vh]", "max-h-[90vh]", "rounded-2xl");
        dialog.classList.add("max-w-none", "w-[99vw]", "h-[98vh]", "max-h-[98vh]", "rounded-xl");
        if (icon) icon.className = "fa-solid fa-compress text-amber-400";
        if (text) text.textContent = "Restore";
        if (btn) btn.title = "Restore Normal Frame Size";
    } else {
        modal.classList.remove("p-1", "sm:p-2");
        modal.classList.add("p-4");
        dialog.classList.remove("max-w-none", "w-[99vw]", "h-[98vh]", "max-h-[98vh]", "rounded-xl");
        dialog.classList.add("max-w-6xl", "h-[88vh]", "max-h-[90vh]", "rounded-2xl");
        if (icon) icon.className = "fa-solid fa-expand text-blue-400";
        if (text) text.textContent = "Enlarge";
        if (btn) btn.title = "Enlarge / Fullscreen";
    }
}

function toggleEnlargeLightboxModal(event) {
    if (event) event.stopPropagation();
    setEnlargeLightboxModalState(!isLightboxModalEnlarged);
}

function applyLightboxZoom(scale) {
    const imgEl = document.getElementById("imageLightboxTarget");
    const inputZoom = document.getElementById("inputLightboxZoom");
    if (!imgEl || !imgEl.naturalWidth) return;

    currentLightboxScale = Math.max(0.1, Math.min(5.0, scale));
    const targetW = Math.round(imgEl.naturalWidth * currentLightboxScale);
    const targetH = Math.round(imgEl.naturalHeight * currentLightboxScale);

    imgEl.style.width = `${targetW}px`;
    imgEl.style.height = `${targetH}px`;
    imgEl.style.maxWidth = "none";

    if (inputZoom) {
        inputZoom.value = Math.round(currentLightboxScale * 100);
    }
}

function handleLightboxZoomInput(inputEl) {
    if (!inputEl) return;
    const parsed = parseFloat(inputEl.value);
    if (!isNaN(parsed) && parsed > 0) {
        applyLightboxZoom(parsed / 100);
    } else {
        inputEl.value = Math.round(currentLightboxScale * 100);
    }
}

async function openImageLightbox(src) {
    if (!src) return;
    const modal = document.getElementById("imageLightboxModal");
    const viewport = document.getElementById("imageLightboxViewport");
    const imgEl = document.getElementById("imageLightboxTarget");
    const dimInfo = document.getElementById("imageLightboxDimInfo");
    if (!modal || !imgEl) return;

    setEnlargeLightboxModalState(false);

    imgEl.style.width = "auto";
    imgEl.style.height = "auto";
    imgEl.style.maxWidth = "100%";

    modal.classList.remove("hidden");

    imgEl.onload = () => {
        const nw = imgEl.naturalWidth || 1;
        const nh = imgEl.naturalHeight || 1;
        if (dimInfo) dimInfo.textContent = `(${nw} × ${nh} px)`;

        const availW = Math.max(320, (viewport ? viewport.clientWidth : Math.floor(window.innerWidth * 0.85)) - 32);
        const availH = Math.max(240, (viewport ? viewport.clientHeight : Math.floor(window.innerHeight * 0.75)) - 32);

        const fitRatio = Math.min(availW / nw, availH / nh);
        defaultFitLightboxScale = fitRatio < 1 ? fitRatio : Math.min(1.25, fitRatio);
        applyLightboxZoom(defaultFitLightboxScale);
    };

    if (String(src).startsWith("drive:")) {
        const fileId = String(src).slice(6);
        if (driveImageCache[fileId]) {
            imgEl.src = driveImageCache[fileId];
        } else {
            if (dimInfo) dimInfo.textContent = "(Loading full-resolution image...)";
            imgEl.src = "";
            const resolved = await fetchDriveImageDataUrl(fileId);
            if (resolved) {
                imgEl.src = resolved;
            } else if (dimInfo) {
                dimInfo.textContent = "(Failed to load image)";
            }
        }
    } else {
        imgEl.src = src;
    }
}

function adjustLightboxZoom(delta, event) {
    if (event) event.stopPropagation();
    applyLightboxZoom(currentLightboxScale + delta);
}

function resetLightboxZoom(event) {
    if (event) event.stopPropagation();
    if (Math.abs(currentLightboxScale - 1) > 0.05) {
        applyLightboxZoom(1);
    } else {
        applyLightboxZoom(defaultFitLightboxScale);
    }
}

function handleLightboxWheel(event) {
    if (!event) return;
    if (event.ctrlKey || event.metaKey) {
        event.preventDefault();
        const delta = event.deltaY < 0 ? 0.15 : -0.15;
        applyLightboxZoom(currentLightboxScale + delta);
    }
}

function closeImageLightbox(event = null) {
    if (event && event.target && event.target.id !== "imageLightboxModal") return;
    const modal = document.getElementById("imageLightboxModal");
    const imgEl = document.getElementById("imageLightboxTarget");
    if (modal) modal.classList.add("hidden");
    setEnlargeLightboxModalState(false);
    if (imgEl) {
        imgEl.onload = null;
        imgEl.src = "";
    }
}

// --- DIRECT IMAGE COPY & PASTE HELPERS ---
async function copyImageSrcToClipboard(imgSrc, btnElement = null) {
    if (!imgSrc) return;
    try {
        if (!navigator.clipboard || !window.ClipboardItem) {
            showToast("Your browser does not support ClipboardItem for copying images directly. Please right-click on the image and select 'Copy Image'.");
            return;
        }

        if (String(imgSrc).startsWith("drive:")) {
            const fileId = String(imgSrc).slice(6);
            imgSrc = driveImageCache[fileId] || await fetchDriveImageDataUrl(fileId);
            if (!imgSrc) throw new Error("Failed to fetch image from Google Drive");
        }

        let blob;
        if (imgSrc.startsWith("data:image/png;base64,")) {
            const byteString = atob(imgSrc.split(",")[1]);
            const ab = new ArrayBuffer(byteString.length);
            const ia = new Uint8Array(ab);
            for (let i = 0; i < byteString.length; i++) {
                ia[i] = byteString.charCodeAt(i);
            }
            blob = new Blob([ia], { type: "image/png" });
        } else {
            const res = await fetch(imgSrc);
            const rawBlob = await res.blob();
            const img = await loadImageFromBlob(rawBlob);
            const canvas = document.createElement("canvas");
            canvas.width = img.naturalWidth;
            canvas.height = img.naturalHeight;
            const ctx = canvas.getContext("2d");
            ctx.drawImage(img, 0, 0);
            blob = await new Promise((resolve, reject) => {
                canvas.toBlob(b => b ? resolve(b) : reject(new Error("Failed to convert image")), "image/png");
            });
        }

        await navigator.clipboard.write([
            new ClipboardItem({ "image/png": blob })
        ]);

        if (btnElement) {
            const originalHTML = btnElement.innerHTML;
            btnElement.innerHTML = `<i class="fa-solid fa-check text-emerald-300"></i> Copied!`;
            btnElement.classList.add("ring-2", "ring-emerald-400");
            setTimeout(() => {
                btnElement.innerHTML = originalHTML;
                btnElement.classList.remove("ring-2", "ring-emerald-400");
            }, 2000);
        }
    } catch (err) {
        console.error("Failed to copy image to clipboard:", err);
        showToast("Failed to copy image to clipboard. Please make sure your browser allows clipboard access.");
    }
}

async function copyCurrentLightboxImage(btnEl = null) {
    const imgEl = document.getElementById("imageLightboxTarget");
    if (!imgEl || !imgEl.src) return;
    await copyImageSrcToClipboard(imgEl.src, btnEl || document.getElementById("btnCopyLightboxImage"));
}

async function handlePasteImageFromClipboardBtn(btnEl) {
    const wrapper = btnEl ? btnEl.closest('.field-with-images-wrapper') : null;
    if (!wrapper) return;

    try {
        if (!navigator.clipboard || !navigator.clipboard.read) {
            showToast("Your browser does not support automatic clipboard reading. Please click the text box and press Ctrl+V.");
            return;
        }

        const clipboardItems = await navigator.clipboard.read();
        let pasted = false;

        for (const item of clipboardItems) {
            const imageType = item.types.find(t => t.startsWith("image/"));
            if (imageType) {
                const blob = await item.getType(imageType);
                await appendImageFileToWrapper(blob, wrapper);
                pasted = true;
            }
        }

        if (!pasted) {
            showToast("No image found in your clipboard.\n\nTip: Copy an image or table screenshot first (Win + Shift + S or Copy Image), then click 'Paste Image' again or press Ctrl+V.");
        } else {
            const originalHTML = btnEl.innerHTML;
            btnEl.innerHTML = `<i class="fa-solid fa-check text-emerald-300"></i> Pasted!`;
            setTimeout(() => {
                btnEl.innerHTML = originalHTML;
            }, 1500);
        }
    } catch (err) {
        console.warn("Clipboard read error:", err);
        showToast("Clipboard read access is restricted by your browser. Please click the text box and press Ctrl+V.");
    }
}

function getChannelIcon(channelName) {
    const channel = String(channelName).toLowerCase();
    if (channel.includes("banca")) return "fa-building-columns";
    if (channel.includes("agency") || channel.includes("agen")) return "fa-user-tie";
    if (channel.includes("direct")) return "fa-paper-plane";
    if (channel.includes("corporate") || channel.includes("group")) return "fa-users";
    return "fa-diagram-project";
}

// Daftar default aturan pengelompokan Filter Channel & Underwriting (otomatis diperbarui dari sheet Config)
const DEFAULT_CHANNEL_FILTER_RULES = [
    { canonical: "Bancassurance", aliases: ["banca", "bancassurance"] },
    { canonical: "Agency", aliases: ["agency", "agen"] },
    { canonical: "Direct", aliases: ["direct", "digital"] },
    { canonical: "Financial Advisor", aliases: ["financial advisor", "fa"] },
    { canonical: "Telemarketing", aliases: ["telemarketing", "tele"] },
    { canonical: "Corporate", aliases: ["corporate", "group", "kumpulan", "employee benefit"] }
];
let CHANNEL_FILTER_RULES = DEFAULT_CHANNEL_FILTER_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));

const DEFAULT_UNDERWRITING_FILTER_RULES = [
    { canonical: "GIO", aliases: ["gio", "guaranteed issue offer", "guaranteed issuance"] },
    { canonical: "SIO", aliases: ["sio", "simplified issue offer", "simplified issuance"] },
    { canonical: "Full Underwriting", aliases: ["full underwriting", "fuw", "full uw"] },
    { canonical: "NMU / Non-Medical", aliases: ["nmu", "non-medical", "non medical"] },
    { canonical: "Medical Underwriting", aliases: ["medical"] },
    { canonical: "Clean Case", aliases: ["clean case"] },
    { canonical: "Health Declaration", aliases: ["health declaration", "simple question", "questionnaire", "deklarasi kesehatan"] },
    { canonical: "Free Cover Limit (FCL)", aliases: ["free cover", "fcl"] }
];
let UNDERWRITING_FILTER_RULES = DEFAULT_UNDERWRITING_FILTER_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));

// Helper untuk mengelompokkan nama Distribution Channel ke kategori dasar berdasarkan daftar di sheet Config
function extractNormalizedChannels(rawChannel) {
    if (!rawChannel || isEmptyVal(rawChannel)) return [];
    // Hapus keterangan di dalam tanda kurung terlebih dahulu, misal "(Allianz Star Network)" atau "(Maybank)"
    const withoutParens = String(rawChannel).replace(/\([^)]*\)/g, ' ');
    // Pisahkan jika produk memiliki lebih dari 1 channel (dipisahkan koma, garis miring, &, atau titik koma)
    const parts = withoutParens.split(/[,/&;]+/).map(s => s.trim()).filter(Boolean);
    const rules = (Array.isArray(CHANNEL_FILTER_RULES) && CHANNEL_FILTER_RULES.length > 0)
        ? CHANNEL_FILTER_RULES
        : DEFAULT_CHANNEL_FILTER_RULES;

    const normalized = [];
    parts.forEach(part => {
        // Hapus deskripsi setelah tanda hubung ber-spasi, misal "Bancassurance - Bank DBS" -> "Bancassurance"
        const basePart = part.split(/\s+[-–]\s+|\s*–\s*/)[0].trim();
        const lower = basePart.toLowerCase();
        const fullLower = part.toLowerCase();
        if (!lower || lower === "-") return;

        let mapped = basePart;
        for (let i = 0; i < rules.length; i++) {
            const rule = rules[i];
            const canonicalLower = String(rule.canonical || "").trim().toLowerCase();
            const kws = [canonicalLower, ...(Array.isArray(rule.aliases) ? rule.aliases.map(a => String(a || "").trim().toLowerCase()) : [])].filter(Boolean);
            if (kws.some(kw => kw.length <= 2
                ? new RegExp(`\\b${kw.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}\\b`, 'i').test(fullLower)
                : (lower.includes(kw) || fullLower.includes(kw))
            )) {
                mapped = rule.canonical;
                break;
            }
        }

        if (mapped && !normalized.includes(mapped)) {
            normalized.push(mapped);
        }
    });
    return normalized;
}

// Helper untuk memecah mata uang majemuk (misal "IDR & USD" atau "IDR / USD") menjadi mata uang tunggal
function extractNormalizedCurrencies(rawCurrency) {
    if (!rawCurrency || isEmptyVal(rawCurrency)) return [];
    const withoutParens = String(rawCurrency).replace(/\([^)]*\)/g, ' ');
    const parts = withoutParens
        .split(/\s*(?:&|\/|,|;|\+|\band\b|\bdan\b)\s*/i)
        .map(s => s.trim().toUpperCase())
        .filter(s => s && s !== "-");
    return [...new Set(parts)];
}

// Helper untuk mengekstrak tipe underwriting utama berdasarkan daftar "Filter Underwriting" di sheet Config
function extractNormalizedUnderwriting(rawUw) {
    if (!rawUw || isEmptyVal(rawUw)) return [];
    const cleanText = String(rawUw).replace(/\[IMG(?:_ID)?:[^\]]*\]/gi, "").trim();
    if (!cleanText || isEmptyVal(cleanText)) return [];

    const tags = [];
    const addTag = (t) => {
        if (t && !tags.includes(t)) tags.push(t);
    };

    const rules = (Array.isArray(UNDERWRITING_FILTER_RULES) && UNDERWRITING_FILTER_RULES.length > 0)
        ? UNDERWRITING_FILTER_RULES
        : DEFAULT_UNDERWRITING_FILTER_RULES;

    const segments = cleanText
        .split(/[\r\n,;]+|(?<=\S)\s*\/\s*(?=\S)/)
        .map(s => s.replace(/^[-•*]\s*/, "").trim())
        .filter(Boolean);

    segments.forEach(seg => {
        const lw = seg.toLowerCase();
        let matchedKnown = false;
        const isNonMedical = /\bnmu\b|\bnon[-\s]?medical\b/i.test(lw);

        rules.forEach(rule => {
            const canonical = String(rule.canonical || "").trim();
            if (!canonical) return;
            // Hindari "Non-Medical" ikut terdorong ke "Medical Underwriting"
            if (isNonMedical && canonical.toLowerCase() === "medical underwriting") return;

            const kws = [canonical.toLowerCase(), ...(Array.isArray(rule.aliases) ? rule.aliases.map(a => String(a || "").trim().toLowerCase()) : [])].filter(Boolean);
            const matched = kws.some(kw => {
                if (!kw) return false;
                if (kw.length <= 4 || !kw.includes(" ")) {
                    const esc = kw.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
                    return new RegExp(`\\b${esc}\\b`, "i").test(lw);
                }
                return lw.includes(kw);
            });
            if (matched) {
                addTag(canonical);
                matchedKnown = true;
            }
        });

        if (!matchedKnown && seg !== "-") {
            const concise = seg.length > 48 ? seg.slice(0, 45) + "..." : seg;
            addTag(concise);
        }
    });

    return tags;
}

function getCustomFilterKey(colName) {
    return `custom:${colName}`;
}

function isCustomFilterKey(key) {
    return String(key || "").startsWith("custom:");
}

function getProductAttrCaseInsensitive(p, colName) {
    if (!p || !p.attributes || !colName) return "-";
    const targetL = String(colName).trim().toLowerCase();
    if (targetL === "par/non par") return p.par || "-";
    const actualKey = Object.keys(p.attributes).find(k => k.toLowerCase() === targetL);
    return actualKey ? p.attributes[actualKey] : "-";
}

function getProductValuesForFilter(p, filterKey) {
    if (!p) return [];
    if (filterKey === "category") return p.category ? [p.category] : [];
    if (filterKey === "company") return (p.company && p.company !== "-") ? [p.company] : [];
    if (filterKey === "type") return (p.productType && p.productType !== "-") ? [p.productType] : [];
    if (filterKey === "channel") return extractNormalizedChannels(p.channel);
    if (filterKey === "currency") return extractNormalizedCurrencies(p.currency);
    if (filterKey === "year") {
        const y = getLaunchYear(p);
        return y ? [y] : [];
    }
    if (filterKey === "underwriting") {
        return extractNormalizedUnderwriting(getProductAttrCaseInsensitive(p, "Underwriting Method"));
    }
    if (isCustomFilterKey(filterKey)) {
        const colName = filterKey.slice("custom:".length);
        const rawVal = getProductAttrCaseInsensitive(p, colName);
        if (!rawVal || isEmptyVal(rawVal)) return [];
        const stripped = String(rawVal)
            .replace(/\[IMG(?:_ID)?:[^\]]*\]/gi, "")
            .replace(/\s+/g, " ")
            .trim();
        if (!stripped && /\[IMG(?:_ID)?:/i.test(String(rawVal))) return ["[Table Image]"];
        if (!stripped || isEmptyVal(stripped)) return [];
        const dplk = parseDplkBenefitValue(stripped);
        if (dplk.isChecked) return ["Available (✓)"];
        return [stripped.length > 70 ? stripped.slice(0, 67) + "..." : stripped];
    }
    return [];
}

function saveCustomFilterColumnsToStorage() {
    try {
        localStorage.setItem(CUSTOM_FILTER_STORAGE_KEY, JSON.stringify(customFilterColumns));
    } catch (e) { }
}

function getCustomDropdownDomId(colName) {
    const idx = customFilterColumns.indexOf(colName);
    return `customFilterDropdown_${idx}`;
}

function getCustomBtnTextDomId(colName) {
    const idx = customFilterColumns.indexOf(colName);
    return `customFilterBtnText_${idx}`;
}

function getDropdownDomId(filterKey) {
    if (isCustomFilterKey(filterKey)) {
        return getCustomDropdownDomId(filterKey.slice("custom:".length));
    }
    return `${filterKey}Dropdown`;
}

function getBtnTextDomId(filterKey) {
    if (isCustomFilterKey(filterKey)) {
        return getCustomBtnTextDomId(filterKey.slice("custom:".length));
    }
    return `${filterKey}BtnText`;
}

function renderCustomFilterControlsShells() {
    const container = document.getElementById("customFiltersContainer");
    if (!container) return;
    const colsSig = JSON.stringify(customFilterColumns);
    if (container.dataset.renderedCols === colsSig) return;
    container.dataset.renderedCols = colsSig;
    container.innerHTML = customFilterColumns.map((colName, idx) => `
                <div class="relative custom-select">
                    <button type="button"
                        class="w-full md:w-auto min-w-[95px] max-w-[190px] border border-blue-200 dark:border-neutral-600 text-xs sm:text-sm rounded-lg px-2.5 py-1.5 bg-blue-50/40 dark:bg-neutral-700 text-gray-700 dark:text-neutral-200 text-left flex justify-between items-center gap-1 hover:bg-blue-50 dark:hover:bg-neutral-600 transition"
                        data-ev-click="a111" data-ev-click-args="${idcmPackArgs([colName])}">
                        <span id="customFilterBtnText_${idx}" class="truncate font-medium">${escapeHtmlAttr(colName)}</span>
                        <i class="fa-solid fa-chevron-down text-[10px] ml-1 text-gray-400 shrink-0"></i>
                    </button>
                    <div id="customFilterDropdown_${idx}" popover="auto"
                        class="dropdown-content hidden absolute right-0 z-50 mt-1 w-64 bg-white dark:bg-neutral-800 border dark:border-neutral-700 rounded-lg shadow-xl max-h-72 overflow-y-auto p-2">
                    </div>
                </div>
            `).join("");
    container.querySelectorAll('.dropdown-content[popover]').forEach(el => {
        el.addEventListener('toggle', (ev) => {
            if (ev.newState === 'closed') {
                el.classList.add('hidden');
                resetDropdownPositionStyles(el);
            }
        });
    });
}

async function toggleCustomFilterDropdown(colName, event) {
    const domId = getCustomDropdownDomId(colName);
    toggleDropdown(domId, event);
    const target = document.getElementById(domId);
    if (target && !target.classList.contains("hidden") && !isFullProductDataLoaded && colName.toLowerCase() !== "underwriting method" && colName.toLowerCase() !== "par/non par") {
        target.innerHTML = `<div class="p-3 text-center text-xs text-gray-500 dark:text-neutral-400"><i class="fa-solid fa-spinner fa-spin me-1"></i>Loading options...</div>`;
        await ensureFullProductDataLoaded();
        populateAllDropdowns();
        filterData();
    }
}

function getDynamicAddFilterPool() {
    const search = (document.getElementById("searchInput")?.value || "").toLowerCase();
    const baseKeys = ["category", "company", "type", "channel", "currency", "year", "underwriting"];
    return products.filter(p => {
        if (search) {
            const matchSearch = String(p.name || "").toLowerCase().includes(search) || String(p.company || "").toLowerCase().includes(search);
            if (!matchSearch) return false;
        }
        for (const k of baseKeys) {
            if (!productMatchesFilterKey(p, k)) return false;
        }
        return true;
    });
}

function isColumnRelevantToCurrentContext(colName, contextPool) {
    if (!colName) return false;
    const lw = String(colName).trim().toLowerCase();
    if (customFilterColumns.some(c => c.toLowerCase() === lw)) return true;

    const activeCategories = (Array.isArray(activeFilters.category) && activeFilters.category.length > 0)
        ? activeFilters.category
        : [...new Set(contextPool.map(p => p.category).filter(Boolean))];

    if (activeCategories.length === 0) return false;

    const hasSharia = activeCategories.some(cat => /^(sharia|syariah)$/i.test(String(cat || "").trim()));
    const hasDplk = activeCategories.some(cat => /^dplk$/i.test(String(cat || "").trim()));
    const hasNonDplk = activeCategories.some(cat => !/^dplk$/i.test(String(cat || "").trim()));

    if (lw === "par/non par") {
        if (isFullProductDataLoaded) {
            return contextPool.some(p => p.par && !isEmptyVal(p.par));
        }
        return hasNonDplk;
    }

    if (!hasSharia && isSyariahSpecColumn(colName)) return false;
    if (!hasDplk && (isDplkSpecColumn(colName) || isDplkChargeColumn(colName) || isDplkBenefitColumn(colName))) return false;
    if (!hasNonDplk && ["coverage period", "premium payment period (ppp)", "min sum assured", "available riders"].includes(lw)) return false;

    const existsInActiveSheet = activeCategories.some(cat => {
        const cols = sheetColumnsMap[cat];
        return Array.isArray(cols) && cols.includes(lw);
    });
    if (!existsInActiveSheet) return false;

    if (isFullProductDataLoaded) {
        return contextPool.some(p => {
            const val = getProductAttrCaseInsensitive(p, colName);
            return val && !isEmptyVal(val);
        });
    }

    return true;
}

async function toggleAddFilterMenu(event) {
    const menuId = "addFilterDropdown";
    const menu = document.getElementById(menuId);
    if (!menu) return;
    const willOpen = menu.classList.contains('hidden');
    if (willOpen) {
        renderAddFilterMenu();
    }
    toggleDropdown(menuId, event);
    if (willOpen && !isFullProductDataLoaded) {
        await ensureFullProductDataLoaded();
        if (!menu.classList.contains('hidden')) {
            renderAddFilterMenu();
            positionDropdownInViewport(menu);
        }
    }
}

function renderAddFilterMenu() {
    const menu = document.getElementById("addFilterDropdown");
    if (!menu) return;

    const prevSearchEl = document.getElementById("addFilterMenuSearchInput");
    const hadSearchFocus = prevSearchEl && document.activeElement === prevSearchEl;
    const selStart = hadSearchFocus ? prevSearchEl.selectionStart : null;

    const builtInCols = new Set(["product name", "company name", "launch date", "tanggal rilis", "product type", "distribution channel", "currency", "underwriting method"]);
    const contextPool = getDynamicAddFilterPool();
    const candidateCols = [];
    const seenLower = new Set();
    const addCand = (col, group) => {
        if (!col) return;
        const lw = String(col).trim().toLowerCase();
        if (!lw || builtInCols.has(lw) || seenLower.has(lw)) return;
        if (!isColumnRelevantToCurrentContext(col, contextPool)) return;
        seenLower.add(lw);
        candidateCols.push({ col, group });
    };

    addCand("Par/Non Par", "General");
    dynamicColumns.filter(col => isNonBenefitColumn(col)).forEach(col => addCand(col, "Provisions"));
    dynamicColumns.filter(col => !isNonBenefitColumn(col) && !isDplkChargeColumn(col)).forEach(col => addCand(col, "Benefits"));
    dynamicColumns.filter(col => isDplkChargeColumn(col)).forEach(col => addCand(col, "Charges"));

    const q = addFilterSearchQuery.trim().toLowerCase();
    const filtered = q ? candidateCols.filter(x => x.col.toLowerCase().includes(q) || x.group.toLowerCase().includes(q)) : candidateCols;

    menu.innerHTML = `
                <div class="sticky top-0 bg-white dark:bg-neutral-800 pb-2 border-b dark:border-neutral-700 space-y-1.5 z-10">
                    <div class="flex items-center justify-between">
                        <span class="font-bold text-[11px] text-gray-700 dark:text-neutral-200 uppercase tracking-wider">Add / Remove Filters (${candidateCols.length})</span>
                        ${customFilterColumns.length > 0 ? `<button type="button" data-ev-click="a112" class="text-[11px] font-semibold text-red-500 hover:underline">Remove All (${customFilterColumns.length})</button>` : ""}
                    </div>
                    <input type="text" id="addFilterMenuSearchInput" value="${escapeHtmlAttr(addFilterSearchQuery)}"
                           data-ev-input="a113"
                           data-ev-click="a85"
                           placeholder="Search database column..."
                           class="w-full px-2 py-1 text-xs border dark:border-neutral-600 rounded bg-gray-50 dark:bg-neutral-700 text-gray-800 dark:text-neutral-100 outline-none focus:ring-1 focus:ring-blue-500">
                </div>
                <div class="pt-1.5 space-y-0.5">
                    ${filtered.length === 0 ? `<p class="text-center text-gray-400 py-3 text-xs">No matching columns for current filters.</p>` : filtered.map(item => {
        const isAdded = customFilterColumns.some(c => c.toLowerCase() === item.col.toLowerCase());
        return `
                            <label class="flex items-center justify-between gap-2 p-1.5 rounded hover:bg-gray-100 dark:hover:bg-neutral-700 cursor-pointer" data-ev-click="a85">
                                <span class="flex items-center gap-2 min-w-0">
                                    <input type="checkbox" ${isAdded ? 'checked' : ''}
                                           data-ev-change="a114" data-ev-change-args="${idcmPackArgs([item.col])}"
                                           class="w-3.5 h-3.5 accent-blue-600 rounded cursor-pointer shrink-0">
                                    <span class="truncate text-gray-700 dark:text-neutral-200 font-medium" title="${escapeHtmlAttr(item.col)}">${escapeHtmlAttr(item.col)}</span>
                                </span>
                                <span class="text-[10px] px-1.5 py-0.2 rounded bg-gray-100 dark:bg-neutral-700 text-gray-500 dark:text-neutral-400 shrink-0">${item.group}</span>
                            </label>`;
    }).join("")}
                </div>
            `;

    if (hadSearchFocus) {
        const restored = document.getElementById("addFilterMenuSearchInput");
        if (restored) {
            restored.focus();
            if (selStart !== null) restored.setSelectionRange(selStart, selStart);
        }
    }
}

async function toggleCustomFilterColumn(colName, enabled, event) {
    if (event) event.stopPropagation();
    const key = getCustomFilterKey(colName);
    if (enabled) {
        if (!customFilterColumns.some(c => c.toLowerCase() === colName.toLowerCase())) {
            customFilterColumns.push(colName);
            activeFilters[key] = [];
        }
        if (!isFullProductDataLoaded && colName.toLowerCase() !== "par/non par") {
            await ensureFullProductDataLoaded();
        }
    } else {
        customFilterColumns = customFilterColumns.filter(c => c.toLowerCase() !== colName.toLowerCase());
        delete activeFilters[key];
    }
    saveCustomFilterColumnsToStorage();
    renderCustomFilterControlsShells();
    populateAllDropdowns();
    filterData();
    renderAddFilterMenu();
}

function removeAllCustomFilters(event) {
    if (event) event.stopPropagation();
    customFilterColumns.forEach(c => {
        delete activeFilters[getCustomFilterKey(c)];
    });
    customFilterColumns = [];
    saveCustomFilterColumnsToStorage();
    renderCustomFilterControlsShells();
    populateAllDropdowns();
    filterData();
    renderAddFilterMenu();
}

// --- DROPDOWN MULTI-SELECT & MOBILE COLLAPSIBLE FILTER PANEL ---
function resetDropdownPositionStyles(el) {
    if (!el) return;
    el.style.top = "";
    el.style.bottom = "";
    el.style.left = "";
    el.style.right = "";
    el.style.margin = "";
    el.style.inset = "";
}

function positionDropdownInViewport(dropdownEl) {
    if (!dropdownEl || dropdownEl.classList.contains('hidden')) return;
    resetDropdownPositionStyles(dropdownEl);

    const triggerWrapper = dropdownEl.closest('.custom-select');
    const triggerBtn = triggerWrapper ? triggerWrapper.querySelector('button') : null;
    if (!triggerBtn) return;

    const vw = window.innerWidth || document.documentElement.clientWidth;
    const vh = window.innerHeight || document.documentElement.clientHeight;
    const btnRect = triggerBtn.getBoundingClientRect();
    const isPopover = dropdownEl.hasAttribute('popover');

    // Smartphone (< 640px): place horizontally inset and vertically below/above trigger
    if (vw < 640) {
        const dropHeight = Math.min(dropdownEl.scrollHeight || 260, Math.round(vh * 0.56), 304);
        const spaceBelow = vh - btnRect.bottom - 12;
        const spaceAbove = btnRect.top - 12;
        if (isPopover) {
            dropdownEl.style.margin = "0";
            dropdownEl.style.left = "0.75rem";
            dropdownEl.style.right = "0.75rem";
            dropdownEl.style.bottom = "auto";
        }
        if (spaceBelow < 180 && spaceAbove > spaceBelow) {
            const topPos = Math.max(12, btnRect.top - dropHeight - 6);
            dropdownEl.style.top = `${Math.round(topPos)}px`;
        } else {
            dropdownEl.style.top = `${Math.round(btnRect.bottom + 6)}px`;
        }
        return;
    }

    // Desktop / Tablet (>= 640px):
    // 1. If native CSS Anchor Positioning is supported on a [popover], let CSS position-area + position-try-fallbacks handle it
    if (isPopover && typeof CSS !== "undefined" && CSS.supports && CSS.supports("position-area: block-end")) {
        return;
    }

    // 2. Fallback for [popover] without CSS Anchor Positioning (e.g. Firefox): tether in viewport coords
    if (isPopover) {
        dropdownEl.style.margin = "0";
        dropdownEl.style.inset = "auto";
        const dropRect = dropdownEl.getBoundingClientRect();
        const dropWidth = dropRect.width || 240;
        const dropHeight = Math.min(dropdownEl.scrollHeight || dropRect.height || 260, 288);
        const spaceBelow = vh - btnRect.bottom - 12;
        const spaceAbove = btnRect.top - 12;
        const topPos = (spaceBelow < dropHeight && spaceAbove > spaceBelow)
            ? Math.max(8, btnRect.top - dropHeight - 6)
            : Math.round(btnRect.bottom + 6);
        const leftPos = Math.min(Math.max(8, Math.round(btnRect.left)), Math.max(8, vw - Math.round(dropWidth) - 8));
        dropdownEl.style.top = `${topPos}px`;
        dropdownEl.style.left = `${leftPos}px`;
        return;
    }

    // 3. Non-popover relative dropdowns: prevent left/right viewport clipping
    const dropRect = dropdownEl.getBoundingClientRect();
    if (dropRect.left < 8) {
        dropdownEl.style.left = "0px";
        dropdownEl.style.right = "auto";
    } else if (dropRect.right > vw - 8) {
        dropdownEl.style.right = "0px";
        dropdownEl.style.left = "auto";
    }
}

function toggleMobileFilterPanel(event) {
    if (event) event.stopPropagation();
    const container = document.getElementById("filterContainer");
    const btn = document.getElementById("btnMobileFilterToggle");
    const chevron = document.getElementById("mobileFilterChevron");
    if (!container) return;

    const isCollapsed = container.classList.toggle("mobile-filters-collapsed");
    const isExpanded = !isCollapsed;
    if (btn) btn.setAttribute("aria-expanded", String(isExpanded));
    if (chevron) chevron.classList.toggle("rotate-180", isExpanded);
    if (!isExpanded) closeAllDropdowns();
}

function closeAllDropdowns() {
    document.querySelectorAll('.dropdown-content').forEach(el => {
        if (el.hasAttribute('popover') && typeof el.hidePopover === 'function') {
            try { if (el.matches(':popover-open')) el.hidePopover(); } catch (e) { }
        }
        el.classList.add('hidden');
        resetDropdownPositionStyles(el);
    });
}

function toggleDropdown(id, event) {
    if (event) event.stopPropagation();
    document.querySelectorAll('.dropdown-content').forEach(el => {
        if (el.id !== id) {
            if (el.hasAttribute('popover') && typeof el.hidePopover === 'function') {
                try { if (el.matches(':popover-open')) el.hidePopover(); } catch (e) { }
            }
            el.classList.add('hidden');
            resetDropdownPositionStyles(el);
        }
    });
    const target = document.getElementById(id);
    if (target) {
        const isCurrentlyHidden = target.classList.contains('hidden');
        if (isCurrentlyHidden) {
            const triggerWrapper = target.closest('.custom-select');
            const triggerBtn = (event && event.currentTarget && event.currentTarget.tagName === 'BUTTON')
                ? event.currentTarget
                : (triggerWrapper ? triggerWrapper.querySelector('button') : null);
            if (triggerBtn && target.hasAttribute('popover')) {
                const anchorId = `--dd-anchor-${id.replace(/[^a-zA-Z0-9_-]/g, '')}`;
                triggerBtn.style.anchorName = anchorId;
                target.style.positionAnchor = anchorId;
            }
            target.classList.remove('hidden');
            if (target.hasAttribute('popover') && typeof target.showPopover === 'function') {
                try {
                    if (triggerBtn) {
                        target.showPopover({ source: triggerBtn });
                    } else {
                        target.showPopover();
                    }
                } catch (e) {
                    try { target.showPopover(); } catch (e2) { }
                }
            }
            positionDropdownInViewport(target);
        } else {
            if (target.hasAttribute('popover') && typeof target.hidePopover === 'function') {
                try { if (target.matches(':popover-open')) target.hidePopover(); } catch (e) { }
            }
            target.classList.add('hidden');
            resetDropdownPositionStyles(target);
        }
    }
}

// Modern Web Guidance: Native <dialog> top-layer sync + Safari closedby fallback + popover sync
(function initNativeDialogsAndPopovers() {
    const dialogs = document.querySelectorAll('dialog.idcm-native-dialog');
    dialogs.forEach(dialog => {
        const syncDialogState = () => {
            const shouldBeOpen = !dialog.classList.contains('hidden');
            if (shouldBeOpen && !dialog.open) {
                try { dialog.showModal(); } catch (e) { }
            } else if (!shouldBeOpen && dialog.open) {
                try { dialog.close(); } catch (e) { }
            }
        };
        const observer = new MutationObserver(syncDialogState);
        observer.observe(dialog, { attributes: true, attributeFilter: ['class'] });

        dialog.addEventListener('close', () => {
            dialog.returnValue = '';
            if (!dialog.classList.contains('hidden')) {
                dialog.classList.add('hidden');
            }
            if (dialog.id === 'uploadModal' && typeof cleanupUploadModalState === 'function') {
                cleanupUploadModalState();
            }
        });
    });

    // Mandatory fallback for browsers without <dialog closedby="any"> support (e.g. Safari)
    if (!('closedBy' in HTMLDialogElement.prototype)) {
        const outside = (e, dialog) => {
            const panel = dialog.firstElementChild || dialog;
            const r = panel.getBoundingClientRect();
            return e.clientX < r.left || e.clientX > r.right || e.clientY < r.top || e.clientY > r.bottom;
        };
        for (const dialog of document.querySelectorAll('dialog[closedby="any"]')) {
            let pressedOutside = false;
            dialog.addEventListener('pointerdown', (e) => { pressedOutside = outside(e, dialog); });
            dialog.addEventListener('pointerup', (e) => {
                if (pressedOutside && outside(e, dialog)) dialog.close();
            });
        }
    }

    // Keep .hidden class synchronized when native popover light-dismisses
    document.querySelectorAll('.dropdown-content[popover]').forEach(el => {
        el.addEventListener('toggle', (ev) => {
            if (ev.newState === 'closed') {
                el.classList.add('hidden');
                resetDropdownPositionStyles(el);
            }
        });
    });
})();

document.addEventListener('click', (e) => {
    if (!e.target.closest('.custom-select')) {
        closeAllDropdowns();
    }
});

window.addEventListener('scroll', () => {
    if (window.innerWidth < 640) {
        const openDrop = document.querySelector('#filterStickyWrapper .dropdown-content:not(.hidden)');
        if (openDrop) positionDropdownInViewport(openDrop);
    }
}, { passive: true });

// Event Keyboard Shortcut (Ctrl+C untuk salin gambar saat preview; Esc ditangani native oleh <dialog> Top Layer)
document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'c') {
        const imageLightboxModal = document.getElementById('imageLightboxModal');
        if (imageLightboxModal && !imageLightboxModal.classList.contains('hidden')) {
            const btnCopy = document.getElementById('btnCopyLightboxImage');
            copyCurrentLightboxImage(btnCopy);
            e.preventDefault();
            return;
        }
    }
});

function updateFilterState(filterKey, value, isChecked) {
    if (!activeFilters[filterKey]) activeFilters[filterKey] = [];
    if (isChecked) {
        if (!activeFilters[filterKey].includes(value)) {
            activeFilters[filterKey].push(value);
        }
    } else {
        activeFilters[filterKey] = activeFilters[filterKey].filter(v => v !== value);
    }
    updateFilterButtonText(filterKey);
    filterData();
}

function toggleSelectAllFilter(filterKey, selectAllChecked, event) {
    if (event) event.stopPropagation();
    const options = getDynamicFilterOptions(filterKey);
    activeFilters[filterKey] = selectAllChecked ? options.slice() : [];
    updateFilterButtonText(filterKey);
    filterData();
}

function clearSingleFilter(filterKey, event) {
    if (event) event.stopPropagation();
    activeFilters[filterKey] = [];
    dropdownSearchQueries[filterKey] = "";
    updateFilterButtonText(filterKey);
    filterData();
}

function clearAllFilters() {
    const searchEl = document.getElementById("searchInput");
    if (searchEl) searchEl.value = "";
    Object.keys(activeFilters).forEach(k => {
        activeFilters[k] = [];
        updateFilterButtonText(k);
    });
    dropdownSearchQueries = {};
    filterData();
}

function updateResetAllFiltersVisibility() {
    const btn = document.getElementById("btnResetAllFilters");
    const searchVal = (document.getElementById("searchInput")?.value || "").trim();
    const activeGroupCount = Object.values(activeFilters).filter(arr => Array.isArray(arr) && arr.length > 0).length;
    const hasActive = Boolean(searchVal) || activeGroupCount > 0;
    if (btn) {
        btn.classList.toggle("hidden", !hasActive);
        btn.classList.toggle("flex", hasActive);
    }

    const mobileBadge = document.getElementById("mobileActiveFilterBadge");
    const mobileBtn = document.getElementById("btnMobileFilterToggle");
    const badgeCount = activeGroupCount + (currentSortOption !== "default" ? 1 : 0);
    if (mobileBadge) {
        mobileBadge.textContent = String(badgeCount);
        mobileBadge.classList.toggle("hidden", badgeCount === 0);
    }
    if (mobileBtn) {
        mobileBtn.classList.toggle("ring-2", badgeCount > 0);
        mobileBtn.classList.toggle("ring-blue-400", badgeCount > 0);
    }
}

function updateFilterButtonText(key) {
    const btn = document.getElementById(getBtnTextDomId(key));
    if (!btn) return;
    const count = (activeFilters[key] || []).length;
    const defaultLabels = {
        category: "Category",
        company: "Company",
        type: "Type",
        channel: "Channel",
        currency: "Currency",
        year: "Launch Year",
        underwriting: "Underwriting"
    };
    const baseLabel = isCustomFilterKey(key) ? key.slice("custom:".length) : (defaultLabels[key] || key);

    if (count === 0) {
        btn.textContent = baseLabel;
        btn.classList.remove("text-blue-600", "dark:text-blue-400", "font-bold");
    } else if (count === 1) {
        btn.textContent = activeFilters[key][0];
        btn.classList.add("text-blue-600", "dark:text-blue-400", "font-bold");
    } else {
        btn.textContent = `${baseLabel} (${count})`;
        btn.classList.add("text-blue-600", "dark:text-blue-400", "font-bold");
    }
}

// Sinkronisasi posisi top sticky filter wrapper agar menempel pas di bawah header navbar
let headerResizeObserverInitialized = false;
function updateStickyFilterPosition() {
    const header = document.getElementById("mainHeader");
    const filterWrapper = document.getElementById("filterStickyWrapper");
    if (header && filterWrapper) {
        filterWrapper.style.top = `${header.offsetHeight}px`;
        if (!headerResizeObserverInitialized && typeof ResizeObserver !== "undefined") {
            headerResizeObserverInitialized = true;
            const ro = new ResizeObserver(() => {
                filterWrapper.style.top = `${header.offsetHeight}px`;
            });
            ro.observe(header);
        }
    }
}
window.addEventListener("resize", () => {
    updateStickyFilterPosition();
    const openDrop = document.querySelector('#filterStickyWrapper .dropdown-content:not(.hidden)');
    if (openDrop) positionDropdownInViewport(openDrop);
});

function buildDropdownUI(filterKey, optionsArray) {
    const container = document.getElementById(getDropdownDomId(filterKey));
    if (!container) return;

    // Pertahankan seleksi yang aktif meskipun sementara tidak ada produk lain yang cocok
    if (!Array.isArray(activeFilters[filterKey])) activeFilters[filterKey] = [];
    const selectedArr = activeFilters[filterKey];
    const displayOptions = [...optionsArray];
    selectedArr.forEach(sel => {
        if (!displayOptions.includes(sel)) displayOptions.push(sel);
    });
    updateFilterButtonText(filterKey);

    // Jika pencarian inline di dalam dropdown sedang difokus, jaga agar input tidak kehilangan fokus
    const prevSearchEl = document.getElementById(`filterSearchInput_${filterKey}`);
    const hadSearchFocus = prevSearchEl && document.activeElement === prevSearchEl;
    const selStart = hadSearchFocus ? prevSearchEl.selectionStart : null;

    container.replaceChildren();

    if (displayOptions.length === 0) {
        container.innerHTML = `<p class="text-xs text-gray-400 text-center py-3">No options match current filters.</p>`;
        return;
    }

    const q = String(dropdownSearchQueries[filterKey] || "").trim().toLowerCase();
    const visibleOptions = q ? displayOptions.filter(o => o.toLowerCase().includes(q)) : displayOptions;

    const allChecked = displayOptions.length > 0 && displayOptions.every(opt => selectedArr.includes(opt));
    const someChecked = selectedArr.length > 0 && !allChecked;

    const headerDiv = document.createElement("div");
    headerDiv.className = "sticky top-0 bg-white dark:bg-neutral-800 pb-1.5 mb-1 border-b dark:border-neutral-700 space-y-1.5 z-10";
    headerDiv.onclick = (e) => e.stopPropagation();

    const topBar = document.createElement("div");
    topBar.className = "flex items-center justify-between gap-2 px-1.5 py-1";

    const selectAllLabel = document.createElement("label");
    selectAllLabel.className = "flex items-center gap-2 cursor-pointer text-xs font-bold text-gray-800 dark:text-neutral-100 select-none";
    const selectAllCb = document.createElement("input");
    selectAllCb.type = "checkbox";
    selectAllCb.checked = allChecked;
    selectAllCb.indeterminate = someChecked;
    selectAllCb.className = "rounded text-blue-600 focus:ring-blue-500 w-4 h-4 accent-blue-600 cursor-pointer";
    selectAllCb.onchange = (e) => toggleSelectAllFilter(filterKey, e.target.checked, e);
    const selectAllSpan = document.createElement("span");
    selectAllSpan.textContent = "(Select All)";
    selectAllLabel.appendChild(selectAllCb);
    selectAllLabel.appendChild(selectAllSpan);

    const clearBtn = document.createElement("button");
    clearBtn.type = "button";
    clearBtn.textContent = "Clear All";
    clearBtn.className = `text-[11px] font-semibold px-1.5 py-0.5 rounded transition ${selectedArr.length > 0 ? "text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-950/40" : "text-gray-400 dark:text-neutral-500 cursor-default"}`;
    clearBtn.disabled = selectedArr.length === 0;
    clearBtn.onclick = (e) => clearSingleFilter(filterKey, e);

    topBar.appendChild(selectAllLabel);
    topBar.appendChild(clearBtn);
    headerDiv.appendChild(topBar);

    if (displayOptions.length > 6 || q) {
        const searchInputEl = document.createElement("input");
        searchInputEl.type = "text";
        searchInputEl.id = `filterSearchInput_${filterKey}`;
        searchInputEl.value = dropdownSearchQueries[filterKey] || "";
        searchInputEl.placeholder = "Search option...";
        searchInputEl.className = "w-full px-2 py-1 text-xs border dark:border-neutral-600 rounded bg-gray-50 dark:bg-neutral-700 text-gray-800 dark:text-neutral-100 outline-none focus:ring-1 focus:ring-blue-500";
        searchInputEl.onclick = (e) => e.stopPropagation();
        searchInputEl.oninput = (e) => {
            dropdownSearchQueries[filterKey] = e.target.value;
            buildDropdownUI(filterKey, optionsArray);
        };
        headerDiv.appendChild(searchInputEl);
    }

    container.appendChild(headerDiv);

    if (visibleOptions.length === 0) {
        const emptyP = document.createElement("p");
        emptyP.className = "text-xs text-gray-400 text-center py-2";
        emptyP.textContent = "No matching option.";
        container.appendChild(emptyP);
    } else {
        visibleOptions.forEach(opt => {
            const isChecked = selectedArr.includes(opt);
            const label = document.createElement('label');
            label.className = "flex items-start space-x-2.5 p-1.5 hover:bg-gray-100 dark:hover:bg-neutral-700 cursor-pointer rounded text-xs sm:text-sm text-gray-700 dark:text-neutral-200 transition";
            const cb = document.createElement('input');
            cb.type = "checkbox";
            cb.value = opt;
            cb.checked = isChecked;
            cb.className = "mt-0.5 rounded text-blue-600 focus:ring-blue-500 w-4 h-4 accent-blue-600 cursor-pointer shrink-0";
            cb.onchange = (e) => updateFilterState(filterKey, e.target.value, e.target.checked);
            const span = document.createElement('span');
            span.className = "flex-1 break-words";
            span.textContent = opt;
            label.appendChild(cb);
            label.appendChild(span);
            container.appendChild(label);
        });
    }

    if (hadSearchFocus) {
        const restored = document.getElementById(`filterSearchInput_${filterKey}`);
        if (restored) {
            restored.focus();
            if (selStart !== null) restored.setSelectionRange(selStart, selStart);
        }
    }
}

// --- SORTING URUTAN PRODUK ---
function sortProducts(list) {
    const categoryOrder = {};
    SHEET_TABS.forEach((tab, idx) => { categoryOrder[tab] = idx + 1; });
    return list.toSorted((a, b) => {
        if (currentSortOption === "launch_desc") {
            // Tanggal Rilis: Terbaru -> Terlama
            const dateA = parseDateValue(a.launchDate);
            const dateB = parseDateValue(b.launchDate);
            if (dateA && dateB) return dateA.getTime() - dateB.getTime();
            if (dateA && !dateB) return -1;
            if (!dateA && dateB) return 1;
            return String(a.name || "").localeCompare(String(b.name || ""));
        }
        if (currentSortOption === "launch_asc") {
            // Tanggal Rilis: Terlama -> Terbaru
            const dateA = parseDateValue(a.launchDate);
            const dateB = parseDateValue(b.launchDate);
            if (dateA && dateB) return dateB.getTime() - dateA.getTime();
            if (dateA && !dateB) return -1;
            if (!dateA && dateB) return 1;
            return String(a.name || "").localeCompare(String(b.name || ""));
        }
        if (currentSortOption === "name_asc") {
            return String(a.name || "").localeCompare(String(b.name || ""));
        }
        if (currentSortOption === "company_asc") {
            const compA = String(a.company || "");
            const compB = String(b.company || "");
            const cmp = compA.localeCompare(compB);
            if (cmp !== 0) return cmp;
            return String(a.name || "").localeCompare(String(b.name || ""));
        }

        // Default: 1. Kategori -> 2. Perusahaan (A-Z) -> 3. Nama Produk (A-Z)
        let catA = categoryOrder[a.category] || 99;
        let catB = categoryOrder[b.category] || 99;
        if (catA !== catB) return catA - catB;

        let compA = String(a.company || "").toLowerCase();
        let compB = String(b.company || "").toLowerCase();
        if (compA !== compB) return compA.localeCompare(compB);

        let nameA = String(a.name || "").toLowerCase();
        let nameB = String(b.name || "").toLowerCase();
        return nameA.localeCompare(nameB);
    });
}

// --- LOAD & REFRESH DATA DARI GOOGLE SHEETS ---
let sheetsProcessed = 0;
let sheetsFailed = 0;
let sheetHasHeaders = {};
let sheetColumnsMap = {}; // kolom non-inti per tab sheet (huruf kecil), untuk membatasi header yang dikirim saat simpan

// Helper untuk menambahkan baris ke Sheet tujuan sekaligus membuat baris Header jika Sheet tujuan masih kosong
async function appendToSheetWithHeaderCheck(targetSheet, masterSchema, rowsArray) {
    if (!sheetHasHeaders[targetSheet]) {
        // Jika sheet tujuan (misal Syariah / HS) masih benar-benar kosong tanpa header:
        // Kirim action "append" tanpa properti "headers" dengan baris pertama = masterSchema (Header)
        // diikuti data produk, agar Google Apps Script tidak error pada getLastColumn() == 0
        const initPayload = {
            action: "append",
            sheetName: targetSheet,
            data: [masterSchema, ...rowsArray]
        };
        await postAuthenticatedAppsScript(initPayload);
        sheetHasHeaders[targetSheet] = true;
    } else {
        const appendPayload = {
            action: "append",
            sheetName: targetSheet,
            headers: masterSchema,
            data: rowsArray
        };
        await postAuthenticatedAppsScript(appendPayload);
    }
}

// Helper untuk memecah sel atribut yang memiliki banyak gambar (> 47.000 karakter) ke kolom lanjutan otomatis (misal "Alokasi Kontribusi [2]")
function isContinuationColumn(colName) {
    return /^.+?\s*\[\d+\]$/.test(String(colName || "").trim());
}

function splitCellValueIntoChunks(rawVal, maxChunkLen = 47000) {
    const clean = sanitizeValue(rawVal);
    if (clean === "-" || clean.length <= maxChunkLen) return [clean];

    const lines = clean.split("\n");
    const chunks = [];
    let current = "";

    for (const line of lines) {
        if (!current) {
            current = line;
        } else if (current.length + 1 + line.length <= maxChunkLen) {
            current += "\n" + line;
        } else {
            chunks.push(current);
            current = line;
        }
    }
    if (current) chunks.push(current);
    return chunks.length > 0 ? chunks : ["-"];
}

function expandRowsWithContinuationColumns(baseMasterSchema, rowsMatrix) {
    const extraColsSet = new Set();
    const expandedMapsPerRow = rowsMatrix.map(row => {
        const rowMap = {};
        baseMasterSchema.forEach((colName, idx) => {
            const val = row[idx] !== undefined ? String(row[idx]) : "-";
            if (idx < 7) {
                rowMap[colName] = val;
            } else {
                const chunks = splitCellValueIntoChunks(val, 47000);
                rowMap[colName] = chunks[0] || "-";
                for (let p = 1; p < chunks.length; p++) {
                    const contColName = `${colName} [${p + 1}]`;
                    extraColsSet.add(contColName);
                    rowMap[contColName] = chunks[p];
                }
            }
        });
        return rowMap;
    });

    extraColsSet.forEach(c => {
        if (!sheetContinuationColumns.includes(c)) sheetContinuationColumns.push(c);
    });
    const batchContinuationCols = Array.from(extraColsSet);
    const finalSchema = [...baseMasterSchema, ...batchContinuationCols.filter(c => !baseMasterSchema.includes(c))];
    const finalRows = expandedMapsPerRow.map(rowMap => finalSchema.map(col => rowMap[col] || "-"));
    return { finalSchema, finalRows };
}

function ensureDocumentsLoaded(forceRefresh = false) {
    if (isDocumentsLoaded && !forceRefresh) {
        return Promise.resolve(productDocuments);
    }
    if (documentsLoadPromise && !forceRefresh) {
        return documentsLoadPromise;
    }

    const loadId = currentSheetLoadId;
    documentsLoadPromise = fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(withSessionToken({ action: "get_all_documents" }))
    })
        .then(res => res.json())
        .then(data => {
            if (loadId !== currentSheetLoadId) return productDocuments;
            if (handleUnauthorizedResponse(data)) return productDocuments;
            if (data && data.status === "success" && Array.isArray(data.documents)) {
                productDocuments = data.documents.filter(d =>
                    d.productName !== "__TABLE_IMAGES__" && d.docType !== "__TABLE_IMAGE__"
                );
                productDriveFolders = Array.isArray(data.folders)
                    ? data.folders.map(f => String(f || "").trim()).filter(Boolean)
                    : [...new Set(productDocuments.map(d => String(d.productName || "").trim()).filter(Boolean))];
                isDocumentsLoaded = true;
                const detailModalEl = document.getElementById("detailModal");
                if (currentActiveProduct && !isEditMode && isFullProductDataLoaded && detailModalEl && !detailModalEl.classList.contains("hidden")) {
                    renderDetailBodyView(currentActiveProduct);
                    scheduleHydrateDriveImages();
                }
                if (currentlyVisibleProducts && currentlyVisibleProducts.length > 0) {
                    renderDeck(currentlyVisibleProducts);
                }
            }
            return productDocuments;
        })
        .catch(err => {
            console.warn("Gagal memuat dokumen langsung dari Google Drive:", err);
            return productDocuments;
        })
        .finally(() => {
            documentsLoadPromise = null;
        });

    return documentsLoadPromise;
}

function applyFullSheetDataToProducts(sheetsArray) {
    if (!Array.isArray(sheetsArray)) return;
    const CORE_IDENTITY_COLS = ["product name", "company name", "launch date", "tanggal rilis", "distribution channel", "currency", "product type", "par/non par"];
    const isCoreProductCol = c => CORE_IDENTITY_COLS.includes(String(c || "").trim().toLowerCase());

    // Indeks cepat ke produk yang sudah ada di memory tanpa mengubah p.id maupun selectedProductIds
    const prodByTripleKey = {};
    const prodByDoubleKey = {};
    products.forEach(p => {
        const catL = String(p.category || "").trim().toLowerCase();
        const nameL = String(p.name || "").trim().toLowerCase();
        const compL = String(p.company || "").trim().toLowerCase();
        const k3 = `${catL}||${nameL}||${compL}`;
        const k2 = `${catL}||${nameL}`;
        if (!prodByTripleKey[k3]) prodByTripleKey[k3] = [];
        if (!prodByDoubleKey[k2]) prodByDoubleKey[k2] = [];
        prodByTripleKey[k3].push(p);
        prodByDoubleKey[k2].push(p);
    });

    sheetsArray.forEach(sheetItem => {
        const tabName = sheetItem.sheetName;
        const headers = Array.isArray(sheetItem.headers) ? sheetItem.headers.map(h => String(h || "").trim()) : [];
        const rows = Array.isArray(sheetItem.rows) ? sheetItem.rows : [];
        if (!sheetItem.exists || headers.length === 0) return;

        sheetHasHeaders[tabName] = headers.some(c => c && c.toLowerCase() === "product name");
        sheetColumnsMap[tabName] = headers.filter(c => c && !isCoreProductCol(c) && !isContinuationColumn(c)).map(c => c.toLowerCase());

        headers.forEach(col => {
            if (col && !isCoreProductCol(col)) {
                if (isContinuationColumn(col)) {
                    if (!sheetContinuationColumns.includes(col)) sheetContinuationColumns.push(col);
                } else {
                    if (!dynamicColumns.includes(col)) dynamicColumns.push(col);
                }
            }
        });

        const colIndexByLower = {};
        headers.forEach((col, idx) => {
            const lw = col.toLowerCase();
            if (colIndexByLower[lw] === undefined) colIndexByLower[lw] = idx;
        });

        const usedProducts = new Set();
        rows.forEach(row => {
            if (!Array.isArray(row)) return;
            const getVal = (colName) => {
                const cIdx = colIndexByLower[String(colName || "").toLowerCase()];
                if (cIdx !== undefined && row[cIdx] !== null && row[cIdx] !== undefined) {
                    return sanitizeValue(String(row[cIdx]));
                }
                return "-";
            };

            const prodName = getVal("Product Name");
            const compName = getVal("Company Name");
            if (prodName === "-" && compName === "-") return;

            const catL = String(tabName || "").trim().toLowerCase();
            const nameL = prodName.trim().toLowerCase();
            const compL = compName.trim().toLowerCase();

            const candidates3 = prodByTripleKey[`${catL}||${nameL}||${compL}`] || [];
            const candidates2 = prodByDoubleKey[`${catL}||${nameL}`] || [];
            const targetProd = candidates3.find(p => !usedProducts.has(p)) || candidates2.find(p => !usedProducts.has(p));
            if (!targetProd) return;
            usedProducts.add(targetProd);

            const fullAttrs = {};
            const contColsForProduct = [];
            headers.forEach(col => {
                if (col && !isCoreProductCol(col)) {
                    if (isContinuationColumn(col)) {
                        contColsForProduct.push(col);
                    } else {
                        fullAttrs[col] = getVal(col);
                    }
                }
            });

            // Gabungkan kolom lanjutan (misal "Alokasi Kontribusi [2]") ke atribut utamanya
            contColsForProduct
                .map(col => {
                    const m = col.match(/^(.+?)\s*\[(\d+)\]$/);
                    return m ? { col, baseCol: m[1].trim(), partNum: parseInt(m[2], 10) } : null;
                })
                .filter(Boolean)
                .sort((a, b) => a.partNum - b.partNum)
                .forEach(({ col, baseCol }) => {
                    const extraVal = getVal(col);
                    if (extraVal && extraVal !== "-") {
                        const matchedBase = Object.keys(fullAttrs).find(k => k.toLowerCase() === baseCol.toLowerCase()) || baseCol;
                        if (fullAttrs[matchedBase] && fullAttrs[matchedBase] !== "-") {
                            fullAttrs[matchedBase] += "\n" + extraVal;
                        } else {
                            fullAttrs[matchedBase] = extraVal;
                        }
                    }
                });

            targetProd.attributes = fullAttrs;
        });
    });
}

function ensureFullProductDataLoaded(loadId = currentSheetLoadId) {
    if (isFullProductDataLoaded) {
        return Promise.resolve(true);
    }
    if (fullProductDataPromise) {
        return fullProductDataPromise;
    }

    fullProductDataPromise = fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(withSessionToken({ action: "get_all_product_data", mode: "full" }))
    })
        .then(res => res.json())
        .then(data => {
            if (loadId !== currentSheetLoadId) return false;
            if (handleUnauthorizedResponse(data)) return false;
            if (!data || data.status !== "success" || !Array.isArray(data.sheets)) {
                throw new Error("Gagal memuat detail lengkap produk");
            }

            applyFullSheetDataToProducts(data.sheets);
            isFullProductDataLoaded = true;
            populateAllDropdowns();
            if (customFilterColumns.some(c => Array.isArray(activeFilters[getCustomFilterKey(c)]) && activeFilters[getCustomFilterKey(c)].length > 0)) {
                filterData();
            }

            const detailModalEl = document.getElementById("detailModal");
            if (currentActiveProduct && !isEditMode && detailModalEl && !detailModalEl.classList.contains("hidden")) {
                renderDetailBodyView(currentActiveProduct);
                scheduleHydrateDriveImages();
            }

            // Hangatkan daftar dokumen Drive di background setelah detail penuh siap
            ensureDocumentsLoaded();
            return true;
        })
        .catch(err => {
            console.warn("Gagal memuat detail lengkap produk:", err);
            return false;
        })
        .finally(() => {
            fullProductDataPromise = null;
        });

    return fullProductDataPromise;
}

const SUMMARY_CACHE_STORAGE_KEY = "idcm_summary_cache_v1";

function readLocalSummaryCache() {
    try {
        const raw = sessionStorage.getItem(SUMMARY_CACHE_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (!parsed || !parsed.payload || !Array.isArray(parsed.payload.sheets)) return null;
        return parsed.payload;
    } catch (e) {
        return null;
    }
}

function writeLocalSummaryCache(payload) {
    try {
        if (!payload || payload.status !== "success" || !Array.isArray(payload.sheets)) return;
        sessionStorage.setItem(SUMMARY_CACHE_STORAGE_KEY, JSON.stringify({
            savedAt: Date.now(),
            payload
        }));
    } catch (e) { }
}

function invalidateLocalSummaryCache() {
    try {
        sessionStorage.removeItem(SUMMARY_CACHE_STORAGE_KEY);
        localStorage.removeItem(SUMMARY_CACHE_STORAGE_KEY);
    } catch (e) { }
}

function setBackgroundSyncBannerVisible(visible) {
    const banner = document.getElementById("backgroundSyncBanner");
    if (banner) banner.classList.toggle("hidden", !visible);
}

function applySummaryResponsePayload(data, { isFromCache = false, loadId = currentSheetLoadId } = {}) {
    if (loadId !== currentSheetLoadId || !data || !Array.isArray(data.sheets)) return;

    products = [];
    dynamicColumns = [];
    resetDynamicConfigListsToDefault();
    sheetContinuationColumns = [];
    sheetsProcessed = 0;
    sheetsFailed = 0;
    sheetHasHeaders = {};
    sheetColumnsMap = {};

    if (data.config) {
        applyDynamicConfigFromBackend(data.config);
    } else {
        SHEET_TABS = [...DEFAULT_SHEET_TABS];
    }

    syncSheetTabSelectors();

    data.sheets.forEach(sheetItem => {
        const tabName = sheetItem.sheetName;
        const headers = Array.isArray(sheetItem.headers) ? sheetItem.headers : [];
        const rows = Array.isArray(sheetItem.rows) ? sheetItem.rows : [];

        if (!sheetItem.exists || headers.length === 0) {
            sheetHasHeaders[tabName] = false;
            sheetColumnsMap[tabName] = [];
            checkAllSheetsLoaded();
            return;
        }

        const gvizCompatibleData = {
            table: {
                cols: headers.map(h => ({ label: String(h || "").trim() })),
                rows: rows.map(r => ({
                    c: (Array.isArray(r) ? r : []).map(cellVal => ({ v: cellVal }))
                }))
            }
        };

        processSheetData(gvizCompatibleData, tabName);
        checkAllSheetsLoaded();
    });

    if (!isFromCache) {
        setBackgroundSyncBannerVisible(false);
        if (data.mode === "summary") {
            ensureFullProductDataLoaded(loadId);
        } else {
            // Kompatibilitas mundur jika Apps Script belum di-redeploy dan langsung mengembalikan data lengkap
            isFullProductDataLoaded = true;
            ensureDocumentsLoaded();
        }
    }
}

function fetchAllSheetData(forceRefresh = false) {
    const icon = document.getElementById("refreshIcon");
    if (icon) icon.classList.add("fa-spin");

    const loadId = ++currentSheetLoadId;
    isFullProductDataLoaded = false;
    fullProductDataPromise = null;
    isDocumentsLoaded = false;
    documentsLoadPromise = null;
    currentLoadUsedLegacyGviz = false;

    const cachedPayload = !forceRefresh ? readLocalSummaryCache() : null;
    let renderedFromCache = false;

    if (cachedPayload) {
        try {
            applySummaryResponsePayload(cachedPayload, { isFromCache: true, loadId });
            renderedFromCache = products.length > 0;
            if (renderedFromCache) {
                setBackgroundSyncBannerVisible(true);
            }
        } catch (cacheErr) {
            console.warn("Cache lokal diabaikan:", cacheErr);
            renderedFromCache = false;
        }
    }

    if (!renderedFromCache) {
        products = [];
        dynamicColumns = [];
        sheetContinuationColumns = [];
        productDocuments = [];
        productDriveFolders = [];
        sheetsProcessed = 0;
        sheetsFailed = 0;
        sheetHasHeaders = {};
        sheetColumnsMap = {};
        setBackgroundSyncBannerVisible(false);
        document.getElementById("loadingIndicator").classList.remove("hidden");
        document.getElementById("deckGrid").classList.add("hidden");
    }

    // Muat Payload Ringkas (mode=summary) terlebih dahulu via POST JSON body agar sessionToken tidak lewat query URL
    fetch(APPS_SCRIPT_URL, {
        method: "POST",
        headers: { "Content-Type": "text/plain;charset=utf-8" },
        body: JSON.stringify(withSessionToken({
            action: "get_all_product_data",
            mode: "summary",
            nocache: !!forceRefresh
        }))
    })
        .then(res => res.json())
        .then(data => {
            if (loadId !== currentSheetLoadId) return;
            if (handleUnauthorizedResponse(data)) return;
            if (!data || data.status !== "success" || !Array.isArray(data.sheets)) {
                throw new Error((data && data.message) || "Endpoint get_all_product_data belum aktif di Apps Script");
            }

            writeLocalSummaryCache(data);
            applySummaryResponsePayload(data, { isFromCache: false, loadId });
        })
        .catch(err => {
            if (loadId !== currentSheetLoadId) return;
            setBackgroundSyncBannerVisible(false);
            const iconEl = document.getElementById("refreshIcon");
            if (iconEl) iconEl.classList.remove("fa-spin");
            if (renderedFromCache) {
                console.warn("Sinkronisasi background gagal (mempertahankan data cache):", err);
                return;
            }
            const loadingEl = document.getElementById("loadingIndicator");
            if (loadingEl) {
                loadingEl.innerHTML = `<p class="text-red-500 font-semibold"><i class="fa-solid fa-triangle-exclamation me-1"></i> ${escapeHtml(err && err.message ? err.message : "Failed to load data from server.")}</p>`;
            }
        });
}

// ====================================================
// SISTEM AUTENTIKASI (SIGN IN, SIGN UP, & FORGOT PASSWORD OTP)
// ====================================================
function getActiveAuthSession() {
    try {
        // Bersihkan sisa token legacy di localStorage (migrasi ke HttpOnly cookie)
        try { localStorage.removeItem(AUTH_STORAGE_KEY); } catch (e) { }
        const raw = sessionStorage.getItem(AUTH_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (!parsed || !parsed.username || !parsed.expiresAt) return null;
        if (Date.now() > Number(parsed.expiresAt)) {
            sessionStorage.removeItem(AUTH_STORAGE_KEY);
            return null;
        }
        return parsed;
    } catch (e) {
        return null;
    }
}

function saveAuthSession(userObj, rememberMe = true, _sessionToken = "", serverExpiresAt = null) {
    const durationMs = rememberMe ? (7 * 24 * 60 * 60 * 1000) : (24 * 60 * 60 * 1000);
    // Simpan hanya metadata non-sensitif di sessionStorage; kredensial sesi sesungguhnya ada di HttpOnly cookie
    const sessionData = {
        username: userObj.username,
        email: userObj.email || "",
        role: String(userObj.role || "viewer").trim().toLowerCase(),
        loginAt: Date.now(),
        expiresAt: serverExpiresAt ? Number(serverExpiresAt) : (Date.now() + durationMs)
    };
    try { localStorage.removeItem(AUTH_STORAGE_KEY); } catch (e) { }
    try { sessionStorage.setItem(AUTH_STORAGE_KEY, JSON.stringify(sessionData)); } catch (e) { }
    return sessionData;
}

function showAuthAlert(message, type = "error") {
    const box = document.getElementById("authAlertBox");
    if (!box) return;
    if (!message) {
        box.className = "hidden mx-5 mt-4 p-3 rounded-xl text-xs font-medium border";
        box.replaceChildren();
        return;
    }
    const safeMsg = escapeHtml(message);
    if (type === "success") {
        box.className = "mx-5 mt-4 p-3 rounded-xl text-xs font-medium border bg-emerald-50 dark:bg-emerald-950/60 text-emerald-800 dark:text-emerald-200 border-emerald-200 dark:border-emerald-800 flex items-start gap-2";
        box.innerHTML = `<i class="fa-solid fa-circle-check text-emerald-600 dark:text-emerald-400 mt-0.5"></i><span>${safeMsg}</span>`;
    } else if (type === "info") {
        box.className = "mx-5 mt-4 p-3 rounded-xl text-xs font-medium border bg-blue-50 dark:bg-neutral-800/80 text-blue-800 dark:text-blue-200 border-blue-200 dark:border-neutral-600 flex items-start gap-2";
        box.innerHTML = `<i class="fa-solid fa-circle-info text-blue-600 dark:text-blue-400 mt-0.5"></i><span>${safeMsg}</span>`;
    } else {
        box.className = "mx-5 mt-4 p-3 rounded-xl text-xs font-medium border bg-red-50 dark:bg-red-950/60 text-red-700 dark:text-red-200 border-red-200 dark:border-red-800 flex items-start gap-2";
        box.innerHTML = `<i class="fa-solid fa-triangle-exclamation text-red-600 dark:text-red-400 mt-0.5"></i><span>${safeMsg}</span>`;
    }
}

function switchAuthTab(mode) {
    showAuthAlert("");
    const tabSwitcher = document.getElementById("authTabSwitcher");
    const btnIn = document.getElementById("tabBtnAuthSignIn");
    const btnUp = document.getElementById("tabBtnAuthSignUp");
    const viewIn = document.getElementById("authViewSignIn");
    const viewUp = document.getElementById("authViewSignUp");
    const viewForgot = document.getElementById("authViewForgot");

    const activeBtnClass = "py-2 text-xs font-bold rounded-lg transition bg-white dark:bg-neutral-900 text-blue-700 dark:text-blue-400 shadow-sm";
    const inactiveBtnClass = "py-2 text-xs font-bold rounded-lg transition text-gray-600 dark:text-neutral-400 hover:text-gray-800 dark:hover:text-neutral-200";

    if (mode === "signup") {
        if (tabSwitcher) tabSwitcher.classList.remove("hidden");
        if (btnIn) btnIn.className = inactiveBtnClass;
        if (btnUp) btnUp.className = activeBtnClass;
        if (viewIn) viewIn.classList.add("hidden");
        if (viewUp) viewUp.classList.remove("hidden");
        if (viewForgot) viewForgot.classList.add("hidden");
    } else if (mode === "forgot") {
        if (tabSwitcher) tabSwitcher.classList.add("hidden");
        if (viewIn) viewIn.classList.add("hidden");
        if (viewUp) viewUp.classList.add("hidden");
        if (viewForgot) viewForgot.classList.remove("hidden");
        const step2 = document.getElementById("authForgotStep2");
        if (step2) step2.classList.add("hidden");
    } else {
        if (tabSwitcher) tabSwitcher.classList.remove("hidden");
        if (btnIn) btnIn.className = activeBtnClass;
        if (btnUp) btnUp.className = inactiveBtnClass;
        if (viewIn) viewIn.classList.remove("hidden");
        if (viewUp) viewUp.classList.add("hidden");
        if (viewForgot) viewForgot.classList.add("hidden");
    }
}

function toggleAuthPasswordVisibility(inputId, btnEl) {
    const input = document.getElementById(inputId);
    if (!input) return;
    const isPwd = input.type === "password";
    input.type = isPwd ? "text" : "password";
    if (btnEl) {
        btnEl.innerHTML = isPwd ? `<i class="fa-solid fa-eye-slash"></i>` : `<i class="fa-solid fa-eye"></i>`;
    }
}

function unlockAppWithSession(session) {
    const overlay = document.getElementById("authGateOverlay");
    if (overlay) overlay.classList.add("hidden");

    const navContainer = document.getElementById("navUserContainer");
    const navLabel = document.getElementById("navUsernameLabel");
    if (navContainer && navLabel) {
        navLabel.textContent = session.username;
        navLabel.title = session.email ? `${session.username} (${session.email})` : session.username;
        navContainer.classList.remove("hidden");
        navContainer.classList.add("flex");
    }

    applyRoleBasedUI();
    updateStickyFilterPosition();
    if (products.length === 0) {
        fetchAllSheetData();
    }
}

function handleLogout() {
    if (!confirm("Are you sure you want to Sign Out of this session?")) return;
    try {
        postAuthenticatedAppsScript({ action: "auth_logout" }, { requireSuccess: false }).catch(() => { });
    } catch (e) { }
    localStorage.removeItem(AUTH_STORAGE_KEY);
    sessionStorage.removeItem(AUTH_STORAGE_KEY);
    clearAllSensitiveClientCaches();

    const navContainer = document.getElementById("navUserContainer");
    if (navContainer) {
        navContainer.classList.add("hidden");
        navContainer.classList.remove("flex");
    }

    const pwdIn = document.getElementById("authSignInPassword");
    if (pwdIn) pwdIn.value = "";
    switchAuthTab("signin");

    const overlay = document.getElementById("authGateOverlay");
    if (overlay) overlay.classList.remove("hidden");
}

async function handleSignInSubmit(e) {
    e.preventDefault();
    showAuthAlert("");

    const identifier = document.getElementById("authSignInIdentifier").value.trim();
    const password = document.getElementById("authSignInPassword").value;
    const rememberMe = document.getElementById("authRememberMe").checked;
    const btn = document.getElementById("btnSubmitSignIn");
    const origHtml = btn.innerHTML;

    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Verifying Account...`;

    try {
        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            credentials: "same-origin",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            signal: (typeof AbortSignal !== "undefined" && typeof AbortSignal.timeout === "function") ? AbortSignal.timeout(30000) : undefined,
            body: JSON.stringify({
                action: "auth_signin",
                identifier,
                password,
                rememberMe
            })
        });
        const res = await resp.json();
        if (!res || res.status !== "success" || !res.user) {
            throw new Error((res && res.message) || "Sign In failed. Please check your Username/Email and Password.");
        }

        const session = saveAuthSession(res.user, rememberMe, "", res.expiresAt);
        unlockAppWithSession(session);
    } catch (err) {
        if (err && (err.name === "TimeoutError" || err.name === "AbortError")) {
            showAuthAlert("Koneksi ke server autentikasi melebihi batas waktu (30 detik). Pastikan Apps Script versi terbaru sudah di-deploy dan coba lagi.", "error");
        } else {
            showAuthAlert(err.message || "Failed to connect to authentication server.", "error");
        }
    } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
    }
}

async function handleSignUpSubmit(e) {
    e.preventDefault();
    showAuthAlert("");

    const username = document.getElementById("authSignUpUsername").value.trim();
    const email = document.getElementById("authSignUpEmail").value.trim();
    const password = document.getElementById("authSignUpPassword").value;
    const confirmPwd = document.getElementById("authSignUpConfirm").value;
    const inviteCode = document.getElementById("authSignUpInviteCode").value.trim();
    const btn = document.getElementById("btnSubmitSignUp");
    const origHtml = btn.innerHTML;

    if (password !== confirmPwd) {
        showAuthAlert("Password confirmation does not match the password entered.", "error");
        return;
    }

    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Registering Account...`;

    try {
        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            credentials: "same-origin",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
                action: "auth_signup",
                username,
                email,
                password,
                inviteCode
            })
        });
        const res = await resp.json();
        if (!res || res.status !== "success" || !res.user) {
            throw new Error((res && res.message) || "Failed to register new account.");
        }

        const session = saveAuthSession(res.user, true, "", res.expiresAt);
        unlockAppWithSession(session);
    } catch (err) {
        showAuthAlert(err.message || "Failed to connect to registration server.", "error");
    } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
    }
}

async function handleForgotRequestOtp(e) {
    e.preventDefault();
    showAuthAlert("");

    const identifier = document.getElementById("authForgotIdentifier").value.trim();
    const btn = document.getElementById("btnRequestOtp");
    const origHtml = btn.innerHTML;

    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Sending OTP Code to Email...`;

    try {
        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
                action: "auth_request_otp",
                identifier
            })
        });
        const res = await resp.json();
        if (!res || res.status !== "success") {
            throw new Error((res && res.message) || "Failed to send OTP code.");
        }

        showAuthAlert(res.message || "A 6-digit OTP code has been sent to your email.", "success");
        document.getElementById("authForgotStep2").classList.remove("hidden");
        document.getElementById("authForgotOtpCode").focus();
    } catch (err) {
        showAuthAlert(err.message || "Failed to send OTP.", "error");
    } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
    }
}

async function handleForgotResetPassword(e) {
    e.preventDefault();
    showAuthAlert("");

    const identifier = document.getElementById("authForgotIdentifier").value.trim();
    const otp = document.getElementById("authForgotOtpCode").value.trim();
    const newPassword = document.getElementById("authForgotNewPassword").value;
    const confirmPassword = document.getElementById("authForgotConfirmPassword").value;
    const btn = document.getElementById("btnResetPasswordSubmit");
    const origHtml = btn.innerHTML;

    if (newPassword !== confirmPassword) {
        showAuthAlert("New password confirmation does not match.", "error");
        return;
    }

    btn.disabled = true;
    btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving New Password...`;

    try {
        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({
                action: "auth_reset_password",
                identifier,
                otp,
                newPassword
            })
        });
        const res = await resp.json();
        if (!res || res.status !== "success") {
            throw new Error((res && res.message) || "Failed to reset password.");
        }

        switchAuthTab("signin");
        document.getElementById("authSignInIdentifier").value = (res.user && res.user.username) ? res.user.username : identifier;
        showAuthAlert("Password updated successfully! Please Sign In with your new password.", "success");
    } catch (err) {
        showAuthAlert(err.message || "Failed to reset password.", "error");
    } finally {
        btn.disabled = false;
        btn.innerHTML = origHtml;
    }
}

document.addEventListener("DOMContentLoaded", async () => {
    updateThemeUI();
    updateStickyFilterPosition();
    // Verifikasi sesi HttpOnly cookie langsung ke server (/api -> auth_me)
    try {
        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            credentials: "same-origin",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify({ action: "auth_me" })
        });
        if (resp.ok) {
            const res = await resp.json();
            if (res && res.status === "success" && res.user) {
                const session = saveAuthSession(res.user, true, "", res.expiresAt);
                unlockAppWithSession(session);
                return;
            }
        }
    } catch (e) { }
    try { localStorage.removeItem(AUTH_STORAGE_KEY); } catch (e) { }
    try { sessionStorage.removeItem(AUTH_STORAGE_KEY); } catch (e) { }
});

function processSheetData(jsonData, categoryName) {
    try {
        let cols = jsonData.table.cols.map(c => c ? (c.label || "").trim() : "");
        let rows = jsonData.table.rows;

        const hasValidLabels = cols.some(l => l && l !== "" && !/^[A-Z]+$/.test(l));
        if (!hasValidLabels && rows.length > 0) {
            cols = rows[0].c.map(cell => (cell && cell.v !== null && cell.v !== undefined) ? String(cell.v).trim() : "");
            rows = rows.slice(1);
        }

        // Tandai apakah sheet ini sudah memiliki header kolom ("Product Name", dll.)
        sheetHasHeaders[categoryName] = cols.some(c => c && c.toLowerCase() === "product name");

        const CORE_IDENTITY_COLS = ["product name", "company name", "launch date", "tanggal rilis", "distribution channel", "currency", "product type", "par/non par"];
        const isCoreProductCol = c => CORE_IDENTITY_COLS.includes(String(c || "").trim().toLowerCase());

        sheetColumnsMap[categoryName] = cols.filter(c => c && !isCoreProductCol(c) && !isContinuationColumn(c)).map(c => c.toLowerCase());

        cols.forEach(col => {
            if (col && !isCoreProductCol(col)) {
                if (isContinuationColumn(col)) {
                    if (!sheetContinuationColumns.includes(col)) sheetContinuationColumns.push(col);
                } else {
                    if (!dynamicColumns.includes(col)) dynamicColumns.push(col);
                }
            }
        });

        rows.forEach((row) => {
            if (!row || !row.c) return;

            const getVal = (colName) => {
                const cIdx = cols.findIndex(c => c.toLowerCase() === colName.toLowerCase());
                if (cIdx !== -1 && row.c[cIdx] && row.c[cIdx].v !== null && row.c[cIdx].v !== undefined) {
                    return sanitizeValue(row.c[cIdx].f ? row.c[cIdx].f : String(row.c[cIdx].v));
                }
                return "-";
            };

            const prodName = getVal("Product Name");
            const compName = getVal("Company Name");
            const rawLaunchDate = getVal("Launch Date") !== "-" ? getVal("Launch Date") : getVal("Tanggal Rilis");

            if (prodName === "-" && compName === "-") return;

            let pObj = {
                id: products.length + 1,
                category: categoryName,
                company: compName,
                name: prodName,
                launchDate: rawLaunchDate !== "-" ? rawLaunchDate : "",
                channel: getVal("Distribution Channel"),
                currency: getVal("Currency"),
                productType: getVal("Product Type"),
                par: getVal("Par/Non Par"),
                attributes: {}
            };

            const contColsForProduct = [];
            cols.forEach(col => {
                if (col && !isCoreProductCol(col)) {
                    if (isContinuationColumn(col)) {
                        contColsForProduct.push(col);
                    } else {
                        pObj.attributes[col] = getVal(col);
                    }
                }
            });

            // Gabungkan kembali nilai dari kolom lanjutan (misal "Alokasi Kontribusi [2]") ke atribut utamanya
            contColsForProduct
                .map(col => {
                    const m = col.match(/^(.+?)\s*\[(\d+)\]$/);
                    return m ? { col, baseCol: m[1].trim(), partNum: parseInt(m[2], 10) } : null;
                })
                .filter(Boolean)
                .sort((a, b) => a.partNum - b.partNum)
                .forEach(({ col, baseCol }) => {
                    const extraVal = getVal(col);
                    if (extraVal && extraVal !== "-") {
                        const matchedBase = Object.keys(pObj.attributes).find(k => k.toLowerCase() === baseCol.toLowerCase()) || baseCol;
                        if (pObj.attributes[matchedBase] && pObj.attributes[matchedBase] !== "-") {
                            pObj.attributes[matchedBase] += "\n" + extraVal;
                        } else {
                            pObj.attributes[matchedBase] = extraVal;
                        }
                    }
                });

            products.push(pObj);
        });
    } catch (err) {
        console.error(`Error tab ${categoryName}:`, err);
    }
}

function checkAllSheetsLoaded() {
    sheetsProcessed++;
    if (sheetsProcessed + sheetsFailed === SHEET_TABS.length) {
        const icon = document.getElementById("refreshIcon");
        if (icon) icon.classList.remove("fa-spin");

        // Deteksi jika ada nama sheet baru di Config yang tab-nya belum dibuat di Google Sheets
        // (karena Google Visualization API otomatis mengembalikan tab pertama / Konven jika nama sheet tidak ditemukan)
        if (SHEET_TABS.length > 1) {
            const firstTab = SHEET_TABS[0];
            const firstTabNames = products.filter(p => p.category === firstTab).map(p => p.name);
            if (firstTabNames.length > 0) {
                SHEET_TABS.slice(1).forEach(tab => {
                    const tabProds = products.filter(p => p.category === tab);
                    if (tabProds.length === firstTabNames.length && tabProds.every((p, idx) => p.name === firstTabNames[idx])) {
                        products = products.filter(p => p.category !== tab);
                        sheetHasHeaders[tab] = false;
                        sheetColumnsMap[tab] = [];
                    }
                });
            }
        }

        if (products.length === 0) {
            document.getElementById("loadingIndicator").innerHTML = `
                        <p class="text-red-500 font-semibold"><i class="fa-solid fa-triangle-exclamation me-1"></i> Data is empty or tabs not found (${SHEET_TABS.map(escapeHtml).join(", ")}).</p>`;
            if (currentLoadUsedLegacyGviz) {
                isFullProductDataLoaded = true;
                ensureDocumentsLoaded();
            }
            return;
        }

        document.getElementById("loadingIndicator").classList.add("hidden");
        document.getElementById("deckGrid").classList.remove("hidden");

        // Sort products hierarchically and update ID numbering
        products = sortProducts(products);
        products.forEach((p, idx) => { p.id = idx + 1; });

        populateAllDropdowns();
        filterData();

        if (currentLoadUsedLegacyGviz) {
            isFullProductDataLoaded = true;
            ensureDocumentsLoaded();
        }
    }
}

// Product launch year (string, e.g. "2023") from Launch Date column; null if empty/invalid
function getLaunchYear(p) {
    const d = parseDateValue(p.launchDate);
    return d ? String(d.getFullYear()) : null;
}

function getAllFilterKeys() {
    return [
        "category", "company", "type", "channel", "currency", "year", "underwriting",
        ...customFilterColumns.map(getCustomFilterKey)
    ];
}

function productMatchesFilterKey(p, filterKey) {
    const selected = activeFilters[filterKey];
    if (!Array.isArray(selected) || selected.length === 0) return true;
    const prodVals = getProductValuesForFilter(p, filterKey);
    return selected.some(v => prodVals.includes(v));
}

function getProductsMatchingAllExcept(excludedKey) {
    const search = (document.getElementById("searchInput")?.value || "").toLowerCase();
    const keys = getAllFilterKeys();
    return products.filter(p => {
        if (search) {
            const matchSearch = String(p.name || "").toLowerCase().includes(search) || String(p.company || "").toLowerCase().includes(search);
            if (!matchSearch) return false;
        }
        for (const k of keys) {
            if (k === excludedKey) continue;
            if (!productMatchesFilterKey(p, k)) return false;
        }
        return true;
    });
}

function getDynamicFilterOptions(filterKey) {
    const pool = getProductsMatchingAllExcept(filterKey);
    const rawOpts = [...new Set(pool.flatMap(p => getProductValuesForFilter(p, filterKey)))].filter(Boolean);
    if (filterKey === "category") {
        const order = {};
        SHEET_TABS.forEach((tab, idx) => { order[tab] = idx + 1; });
        return rawOpts.sort((a, b) => (order[a] || 999) - (order[b] || 999) || a.localeCompare(b));
    }
    if (filterKey === "year") {
        return rawOpts.sort((a, b) => Number(b) - Number(a));
    }
    return rawOpts.sort((a, b) => a.localeCompare(b));
}

function populateAllDropdowns() {
    renderCustomFilterControlsShells();
    getAllFilterKeys().forEach(filterKey => {
        const opts = getDynamicFilterOptions(filterKey);
        buildDropdownUI(filterKey, opts);
    });
    const addMenu = document.getElementById("addFilterDropdown");
    if (addMenu && !addMenu.classList.contains("hidden")) {
        renderAddFilterMenu();
    }
    updateResetAllFiltersVisibility();
}

function getCategoryBadge(category) {
    const c = String(category || "").trim().toLowerCase();
    const safeCat = escapeHtml(category);
    if (c === "conventional" || c === "konven") return `<span class="bg-blue-100 text-blue-800 dark:bg-neutral-700/70 dark:text-blue-300 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${safeCat}</span>`;
    if (c === "sharia" || c === "syariah") return `<span class="bg-green-100 text-green-800 dark:bg-green-900/60 dark:text-green-300 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${safeCat}</span>`;
    if (c === "health & surgical" || c === "hs") return `<span class="bg-purple-100 text-purple-800 dark:bg-purple-900/60 dark:text-purple-300 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${safeCat}</span>`;
    if (c === "dplk") return `<span class="bg-amber-100 text-amber-800 dark:bg-amber-900/60 dark:text-amber-300 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${safeCat}</span>`;
    return `<span class="bg-teal-100 text-teal-800 dark:bg-teal-900/60 dark:text-teal-300 text-[10px] font-bold px-2 py-0.5 rounded uppercase tracking-wider">${safeCat}</span>`;
}

function renderDeck(data) {
    const grid = document.getElementById("deckGrid");
    document.getElementById("totalProductCount").textContent = `${data.length} Products`;
    grid.replaceChildren();

    if (data.length === 0) {
        grid.innerHTML = `<div class="col-span-full text-center py-12 text-gray-500 dark:text-neutral-400">No products found matching your filters</div>`;
        return;
    }

    const benefitCols = getBenefitColumns();
    const frag = document.createDocumentFragment();

    data.forEach(p => {
        const isSelected = selectedProductIds.includes(p.id);
        const card = document.createElement("div");
        card.className = `product-card bg-white dark:bg-neutral-800 rounded-xl shadow-sm border ${isSelected ? 'border-blue-500 ring-2 ring-blue-200 dark:ring-blue-900' : 'border-gray-200 dark:border-neutral-700'} hover:shadow-md transition-shadow p-5 flex flex-col justify-between relative`;

        let previewBenefitKey = benefitCols.find(k => k.toLowerCase().includes("death benefit") && !isEmptyVal(p.attributes[k])) || benefitCols.find(k => !isEmptyVal(p.attributes[k])) || "";
        let previewVal = previewBenefitKey ? p.attributes[previewBenefitKey] : "-";
        const checklistBenefits = benefitCols.filter(k => {
            return parseDplkBenefitValue(p.attributes[k]).isChecked;
        });
        const previewCardHtml = checklistBenefits.length > 0
            ? `<div class="text-xs bg-gray-50 dark:bg-neutral-700/60 p-2.5 rounded-lg text-gray-600 dark:text-neutral-300 mb-3 line-clamp-3">
                            <strong class="text-gray-800 dark:text-neutral-200 block"><i class="fa-solid fa-list-check text-emerald-600 dark:text-emerald-400 me-1"></i>Pension Benefits (${checklistBenefits.length}):</strong> ${checklistBenefits.map(escapeHtml).join(", ")}
                       </div>`
            : (previewBenefitKey && !isEmptyVal(previewVal)
                ? `<div class="text-xs bg-gray-50 dark:bg-neutral-700/60 p-2.5 rounded-lg text-gray-600 dark:text-neutral-300 mb-3 line-clamp-3">
                            <strong class="text-gray-800 dark:text-neutral-200 block">${escapeHtml(previewBenefitKey)}:</strong> ${formatListText(previewVal)}
                           </div>`
                : '');

        card.innerHTML = `
                    <div class="flex-1 cursor-pointer group/cardbody" data-ev-click="a115" data-ev-click-args="${idcmPackArgs([p.id])}">
                        <div class="flex justify-between items-start mb-2">
                            <span class="text-xs font-semibold px-2.5 py-0.5 rounded-full bg-gray-100 dark:bg-neutral-700 text-gray-700 dark:text-neutral-200 border border-gray-200 dark:border-neutral-600">${escapeHtml(p.company)}</span>
                            <div class="flex flex-col items-end gap-1">
                                ${getCategoryBadge(p.category)}
                                <span class="text-xs text-gray-500 dark:text-neutral-400 font-medium">${escapeHtml(p.currency)}</span>
                            </div>
                        </div>
                        <h3 class="font-bold text-gray-800 dark:text-neutral-100 group-hover/cardbody:text-blue-600 dark:group-hover/cardbody:text-blue-400 transition-colors text-base mb-1">${escapeHtml(p.name)}</h3>
                        <div class="flex flex-wrap items-center gap-x-2 gap-y-1 text-xs text-gray-500 dark:text-neutral-400 mb-3">
                            <span><i class="fa-solid ${getChannelIcon(p.channel)} me-1"></i>${escapeHtml(p.channel)}</span>
                            <span>&bull;</span>
                            <span class="italic">${escapeHtml(p.productType)}</span>
                            ${p.launchDate && p.launchDate !== '-' ? `
                                <span>&bull;</span>
                                <span class="text-[11px] text-blue-600 dark:text-blue-400 font-medium" title="Launch Date">
                                    <i class="fa-regular fa-calendar-check me-1"></i>${escapeHtml(formatDisplayDate(p.launchDate))}
                                </span>
                            ` : ''}
                            ${(() => {
                const docCount = getDocumentsForProduct(p.name).length;
                return docCount > 0 ? `
                                    <span>&bull;</span>
                                    <span class="text-[11px] text-emerald-600 dark:text-emerald-400 font-medium" title="${docCount} Documents Available">
                                        <i class="fa-solid fa-paperclip me-1"></i>${docCount} Doc${docCount > 1 ? 's' : ''}
                                    </span>
                                ` : '';
            })()}
                        </div>
                        
                        ${previewCardHtml}
                    </div>

                    <div class="pt-3 border-t dark:border-neutral-700 flex justify-between items-center text-xs">
                        <button data-ev-click="a115" data-ev-click-args="${idcmPackArgs([p.id])}" class="text-blue-600 dark:text-blue-400 hover:underline font-medium">View Details</button>
                        <button data-ev-click="a116" data-ev-click-args="${idcmPackArgs([p.id])}" class="px-3 py-1.5 rounded-lg border text-xs font-semibold flex items-center gap-1 transition-colors ${isSelected ? 'bg-blue-600 text-white border-blue-600' : 'bg-white dark:bg-neutral-700 text-gray-700 dark:text-neutral-200 border-gray-200 dark:border-neutral-600 hover:bg-gray-50 dark:hover:bg-neutral-600'}">
                            <i class="fa-solid ${isSelected ? 'fa-check' : 'fa-plus'}"></i> ${isSelected ? 'Selected' : 'Select'}
                        </button>
                    </div>
                `;
        frag.appendChild(card);
    });
    grid.appendChild(frag);
}

function filterData() {
    currentlyVisibleProducts = getProductsMatchingAllExcept(null);
    currentlyVisibleProducts = sortProducts(currentlyVisibleProducts);
    populateAllDropdowns();
    renderDeck(currentlyVisibleProducts);
    updateCompareBar();
    updateResetAllFiltersVisibility();
}

function toggleSelectProduct(id, e) {
    e.stopPropagation();
    if (selectedProductIds.includes(id)) {
        selectedProductIds = selectedProductIds.filter(pid => pid !== id);
    } else {
        selectedProductIds.push(id);
    }
    updateCompareBar();
    renderDeck(currentlyVisibleProducts);
}

function selectAllVisible() {
    if (!currentlyVisibleProducts || currentlyVisibleProducts.length === 0) return;

    const allVisibleSelected = currentlyVisibleProducts.every(p => selectedProductIds.includes(p.id));
    if (allVisibleSelected) {
        const visibleSet = new Set(currentlyVisibleProducts.map(p => p.id));
        selectedProductIds = selectedProductIds.filter(id => !visibleSet.has(id));
    } else {
        currentlyVisibleProducts.forEach(p => {
            if (!selectedProductIds.includes(p.id)) {
                selectedProductIds.push(p.id);
            }
        });
    }
    updateCompareBar();
    renderDeck(currentlyVisibleProducts);
}

function updateCompareBar() {
    const bar = document.getElementById("compareBar");
    document.getElementById("selectedCount").textContent = selectedProductIds.length;
    if (selectedProductIds.length > 0) {
        bar.classList.remove("translate-y-32");
    } else {
        bar.classList.add("translate-y-32");
    }

    const btnSelectAll = document.getElementById("btnSelectAllVisible");
    if (btnSelectAll) {
        const allVisibleSelected = currentlyVisibleProducts.length > 0 &&
            currentlyVisibleProducts.every(p => selectedProductIds.includes(p.id));
        if (allVisibleSelected) {
            btnSelectAll.innerHTML = `<i class="fa-solid fa-xmark me-1" aria-hidden="true"></i> Deselect All`;
            btnSelectAll.className = "col-span-2 sm:col-auto justify-self-end sm:justify-self-auto text-[11px] sm:text-xs md:text-sm bg-amber-600/90 hover:bg-amber-600 px-2 sm:px-3 py-1.5 rounded-lg transition text-white font-medium border border-amber-400/50 whitespace-nowrap";
        } else {
            btnSelectAll.innerHTML = `<i class="fa-solid fa-check-double me-1" aria-hidden="true"></i> Select All`;
            btnSelectAll.className = "col-span-2 sm:col-auto justify-self-end sm:justify-self-auto text-[11px] sm:text-xs md:text-sm bg-gray-700 hover:bg-gray-600 px-2 sm:px-3 py-1.5 rounded-lg transition text-gray-200 font-medium border border-gray-600 whitespace-nowrap";
        }
    }
}

function clearSelection() {
    selectedProductIds = [];
    updateCompareBar();
    renderDeck(currentlyVisibleProducts);
}

let isCompareModalEnlarged = false;
let currentCompareTableZoom = 1;
let compareFreezeRows = 2;
let compareFreezeCols = 1;

function sortComparisonProductsAstraFirst(list) {
    if (!Array.isArray(list)) return [];
    return list.toSorted((a, b) => {
        const aAstra = /astra/i.test(String((a && a.company) || ""));
        const bAstra = /astra/i.test(String((b && b.company) || ""));
        if (aAstra !== bAstra) return aAstra ? -1 : 1;
        return 0;
    });
}

function onComparisonFreezeChange() {
    const rowSel = document.getElementById("selectFreezeRows");
    const colSel = document.getElementById("selectFreezeCols");
    compareFreezeRows = Math.max(0, parseInt(rowSel ? rowSel.value : "2", 10) || 0);
    compareFreezeCols = Math.max(0, parseInt(colSel ? colSel.value : "1", 10) || 0);
    applyComparisonTableFreeze();
}

function applyComparisonTableFreeze() {
    const table = document.getElementById("comparisonTable");
    if (!table) return;
    const rows = Array.from(table.querySelectorAll("tr"));
    if (rows.length === 0) return;

    // Reset sticky inline styles terlebih dahulu
    rows.forEach(tr => {
        Array.from(tr.children).forEach(cell => {
            cell.style.position = "";
            cell.style.top = "";
            cell.style.left = "";
            cell.style.zIndex = "";
            cell.style.boxShadow = "";
        });
    });

    const maxRowCount = Math.min(compareFreezeRows, rows.length);
    const firstRowCells = Array.from(rows[0].children);
    const maxColCount = Math.min(compareFreezeCols, firstRowCells.length);

    const zoom = currentCompareTableZoom || 1;
    const rowTops = [];
    let accTop = 0;
    for (let r = 0; r < maxRowCount; r++) {
        rowTops.push(accTop);
        const h = rows[r].offsetHeight || Math.round(rows[r].getBoundingClientRect().height / zoom);
        accTop += h;
    }

    const colLefts = [];
    let accLeft = 0;
    for (let c = 0; c < maxColCount; c++) {
        colLefts.push(accLeft);
        const w = firstRowCells[c].offsetWidth || Math.round(firstRowCells[c].getBoundingClientRect().width / zoom);
        accLeft += w;
    }

    const isDark = document.documentElement.classList.contains("dark");
    rows.forEach((tr, rIdx) => {
        const isFrozenRow = rIdx < maxRowCount;
        Array.from(tr.children).forEach((cell, cIdx) => {
            const isFrozenCol = cIdx < maxColCount;
            if (!isFrozenRow && !isFrozenCol) return;

            cell.style.position = "sticky";
            if (isFrozenRow) {
                cell.style.top = `${rowTops[rIdx]}px`;
            }
            if (isFrozenCol) {
                cell.style.left = `${colLefts[cIdx]}px`;
            }

            if (isFrozenRow && isFrozenCol) {
                cell.style.zIndex = "30";
            } else if (isFrozenRow) {
                cell.style.zIndex = "20";
            } else {
                cell.style.zIndex = "10";
            }

            // Pastikan sel beku memiliki latar belakang solid agar teks di belakangnya tidak tembus saat di-scroll
            const compBg = getComputedStyle(cell).backgroundColor;
            if (!compBg || compBg === "rgba(0, 0, 0, 0)" || compBg === "transparent") {
                cell.style.backgroundColor = isFrozenRow ? (isDark ? "#374151" : "#f3f4f6") : (isDark ? "#1f2937" : "#ffffff");
            }

            const shadows = [];
            if (isFrozenCol && cIdx === maxColCount - 1) {
                shadows.push("inset -2px 0 0 rgba(59, 130, 246, 0.35)");
            }
            if (isFrozenRow && rIdx === maxRowCount - 1) {
                shadows.push("inset 0 -2px 0 rgba(59, 130, 246, 0.35)");
            }
            if (shadows.length > 0) {
                cell.style.boxShadow = shadows.join(", ");
            }
        });
    });
}

window.addEventListener("resize", () => {
    const modal = document.getElementById("compareModal");
    if (modal && !modal.classList.contains("hidden")) {
        requestAnimationFrame(() => applyComparisonTableFreeze());
    }
});

function toggleEnlargeCompareModal() {
    const modal = document.getElementById("compareModal");
    const dialog = document.getElementById("compareModalDialog");
    const btn = document.getElementById("btnEnlargeCompareModal");
    if (!modal || !dialog) return;

    isCompareModalEnlarged = !isCompareModalEnlarged;
    if (isCompareModalEnlarged) {
        modal.classList.remove("p-2", "sm:p-4", "p-4", "md:p-6");
        modal.classList.add("p-0");
        dialog.classList.remove("max-w-6xl", "max-w-7xl", "max-h-[92dvh]", "sm:max-h-[90vh]", "max-h-[90vh]", "rounded-2xl");
        dialog.classList.add("w-screen", "h-[100dvh]", "max-w-none", "max-h-[100dvh]", "rounded-none");
        if (btn) {
            btn.innerHTML = `<i class="fa-solid fa-compress text-amber-400" aria-hidden="true"></i> <span class="hidden sm:inline">Restore</span>`;
            btn.title = "Restore Normal Size";
        }
    } else {
        modal.classList.remove("p-0");
        modal.classList.add("p-2", "sm:p-4");
        dialog.classList.remove("w-screen", "h-[100dvh]", "h-screen", "max-w-none", "max-h-[100dvh]", "max-h-screen", "rounded-none");
        dialog.classList.add("max-w-6xl", "max-h-[92dvh]", "sm:max-h-[90vh]", "rounded-2xl");
        if (btn) {
            btn.innerHTML = `<i class="fa-solid fa-expand text-blue-400" aria-hidden="true"></i> <span class="hidden sm:inline">Enlarge</span>`;
            btn.title = "Enlarge / Restore Comparison Table Screen";
        }
    }
    requestAnimationFrame(() => applyComparisonTableFreeze());
}

function applyCompareTableZoom(scale) {
    const table = document.getElementById("comparisonTable");
    const zoomInput = document.getElementById("inputCompareTableZoom");
    const clamped = Math.min(2.5, Math.max(0.4, Math.round(scale * 100) / 100));
    currentCompareTableZoom = clamped;

    if (table) {
        table.style.zoom = String(clamped);
    }
    if (zoomInput) {
        zoomInput.value = Math.round(clamped * 100);
    }
    requestAnimationFrame(() => applyComparisonTableFreeze());
}

function adjustCompareTableZoom(delta) {
    applyCompareTableZoom(currentCompareTableZoom + delta);
}

function handleCompareTableZoomInput(inputEl) {
    if (!inputEl) return;
    let rawVal = parseFloat(inputEl.value);
    if (isNaN(rawVal)) {
        inputEl.value = Math.round(currentCompareTableZoom * 100);
        return;
    }
    const clampedPercent = Math.min(250, Math.max(40, rawVal));
    applyCompareTableZoom(clampedPercent / 100);
}

function resetCompareTableZoom() {
    applyCompareTableZoom(1);
}

function handleCompareTableWheel(event) {
    if (!event.ctrlKey) return;
    event.preventDefault();
    const delta = event.deltaY < 0 ? 0.1 : -0.1;
    adjustCompareTableZoom(delta);
}

function closeCategoryMismatchModal() {
    const modal = document.getElementById("categoryMismatchModal");
    if (modal) modal.classList.add("hidden");
}

function proceedCrossCategoryComparison() {
    closeCategoryMismatchModal();
    openComparisonModal(true);
}

async function openComparisonModal(forceProceed = false) {
    const container = document.getElementById("modalTableContainer");
    const compareModal = document.getElementById("compareModal");
    currentComparisonProducts = sortComparisonProductsAstraFirst(
        products.filter(p => selectedProductIds.includes(p.id))
    );
    if (currentComparisonProducts.length === 0) return;

    // Check if user is comparing Insurance Products (Konven/Syariah/HS/etc.) with DPLK Products
    if (!forceProceed) {
        const dplkProducts = currentComparisonProducts.filter(
            p => String(p.category || "").trim().toUpperCase() === "DPLK"
        );
        const insuranceProducts = currentComparisonProducts.filter(
            p => String(p.category || "").trim().toUpperCase() !== "DPLK"
        );

        if (dplkProducts.length > 0 && insuranceProducts.length > 0) {
            const listEl = document.getElementById("categoryMismatchProductList");
            if (listEl) {
                const renderGroupItems = (items, badgeClass) => items.map(p => `
                            <div class="flex items-center justify-between gap-2 bg-white/80 dark:bg-neutral-800/80 px-2.5 py-1.5 rounded-lg border border-gray-200/70 dark:border-neutral-700">
                                <div class="min-w-0">
                                    <div class="font-bold text-gray-800 dark:text-neutral-100 truncate">${escapeHtml(p.name || "-")}</div>
                                    <div class="text-[11px] text-gray-500 dark:text-neutral-400 truncate">${escapeHtml(p.company || "-")}</div>
                                </div>
                                <span class="text-[10px] font-bold px-2 py-0.5 rounded shrink-0 ${badgeClass}">
                                    ${escapeHtml(p.category || "-")}
                                </span>
                            </div>
                        `).join("");

                listEl.innerHTML = `
                            <div class="space-y-1.5">
                                <div class="text-[11px] font-extrabold uppercase tracking-wider text-blue-800 dark:text-blue-300 flex items-center gap-1.5">
                                    <i class="fa-solid fa-shield-halved"></i> Insurance Products (${insuranceProducts.length})
                                </div>
                                <div class="space-y-1">
                                    ${renderGroupItems(insuranceProducts, "bg-blue-100 text-blue-800 dark:bg-neutral-700/70 dark:text-blue-200")}
                                </div>
                            </div>
                            <div class="space-y-1.5 pt-1 border-t border-amber-200/70 dark:border-amber-800/50">
                                <div class="text-[11px] font-extrabold uppercase tracking-wider text-amber-800 dark:text-amber-300 flex items-center gap-1.5">
                                    <i class="fa-solid fa-piggy-bank"></i> Pension Fund / DPLK Products (${dplkProducts.length})
                                </div>
                                <div class="space-y-1">
                                    ${renderGroupItems(dplkProducts, "bg-amber-100 text-amber-800 dark:bg-amber-900/60 dark:text-amber-200")}
                                </div>
                            </div>
                        `;
            }
            const mismatchModal = document.getElementById("categoryMismatchModal");
            if (mismatchModal) {
                mismatchModal.classList.remove("hidden");
                return;
            }
        }
    }

    if (!isFullProductDataLoaded) {
        container.innerHTML = `
                    <div class="py-16 text-center text-gray-500 dark:text-neutral-400 text-sm space-y-2.5">
                        <i class="fa-solid fa-spinner fa-spin text-3xl text-blue-600"></i>
                        <p class="font-medium">Loading full comparison details...</p>
                    </div>`;
        compareModal.classList.remove("hidden");

        await ensureFullProductDataLoaded();
        if (compareModal.classList.contains("hidden")) return;
        currentComparisonProducts = sortComparisonProductsAstraFirst(
            products.filter(p => selectedProductIds.includes(p.id))
        );
        if (currentComparisonProducts.length === 0) return;
    }

    currentComparisonProducts.forEach(p => {
        p.launchDateFormatted = formatDisplayDate(p.launchDate);
    });

    let fixedFeatures = [
        { label: "Category System", key: "category", isAttr: false },
        { label: "Launch Date", key: "launchDateFormatted", isAttr: false },
        { label: "Distribution Channel", key: "channel", isAttr: false },
        { label: "Product Type", key: "productType", isAttr: false }
    ];

    let allFeatures = [...fixedFeatures];
    const specialColSetLower = new Set(["aqd", "akad", "alokasi kontribusi", "surplus underwriting"]);
    const customSectionByColLower = new Map();
    const customSectionOrder = [];
    currentComparisonProducts.forEach(p => {
        const cfg = getCategorySchemaConfig(p.category);
        (cfg.specialProvisions || []).forEach(sp => {
            if (sp.name) specialColSetLower.add(sp.name.toLowerCase());
        });
        (cfg.productSections || []).forEach((sec, secIdx) => {
            if (sec.isChargeLike || secIdx === 0) return;
            if (!customSectionOrder.includes(sec.title)) customSectionOrder.push(sec.title);
            (sec.items || []).forEach(it => {
                if (it && !customSectionByColLower.has(it.toLowerCase())) {
                    customSectionByColLower.set(it.toLowerCase(), sec.title);
                }
            });
        });
    });

    // Order: General Provisions -> Special Provisions -> Primary Benefits -> Additional Benefit -> Custom Sections -> Charges -> Available Riders
    const generalSpecCols = dynamicColumns.filter(col =>
        isNonBenefitColumn(col) &&
        !specialColSetLower.has(col.toLowerCase()) &&
        col.toLowerCase() !== "available riders" &&
        col.toLowerCase() !== "additional benefit" &&
        col.toLowerCase() !== "remarks"
    );
    const specialSpecCols = dynamicColumns.filter(col => specialColSetLower.has(col.toLowerCase()));
    const regularBenefitCols = getRegularBenefitColumns().filter(col => !customSectionByColLower.has(col.toLowerCase()));
    const chargeCols = getChargeColumns();
    const additionalBenefitCols = dynamicColumns.filter(col => col.toLowerCase() === "additional benefit");
    const riderCols = dynamicColumns.filter(col => col.toLowerCase() === "available riders");

    generalSpecCols.forEach(col => {
        allFeatures.push({ label: col, key: col, isAttr: true, isHighlight: false });
    });
    specialSpecCols.forEach(col => {
        const displayLabel = (col.toLowerCase() === "aqd" || col.toLowerCase() === "akad") ? "Aqd / Contract" : col;
        allFeatures.push({ label: displayLabel, key: col, isAttr: true, isHighlight: false, isSpecialProvision: true });
    });
    regularBenefitCols.forEach(col => {
        allFeatures.push({ label: col, key: col, isAttr: true, isHighlight: true, isBenefit: true });
    });
    additionalBenefitCols.forEach(col => {
        allFeatures.push({ label: "Additional Benefit", key: col, isAttr: true, isHighlight: true, isBenefit: true });
    });
    customSectionOrder.forEach(secTitle => {
        dynamicColumns.forEach(col => {
            if (customSectionByColLower.get(col.toLowerCase()) === secTitle) {
                allFeatures.push({ label: col, key: col, isAttr: true, isHighlight: true, customSectionTitle: secTitle });
            }
        });
    });
    chargeCols.forEach(col => {
        allFeatures.push({ label: col, key: col, isAttr: true, isHighlight: true, isCharge: true });
    });
    riderCols.forEach(col => {
        allFeatures.push({ label: col, key: col, isAttr: true, isHighlight: false });
    });

    let filteredFeatures = allFeatures.filter(feat => {
        return currentComparisonProducts.some(p => {
            let val = feat.isAttr ? p.attributes[feat.key] : p[feat.key];
            return !isEmptyVal(val);
        });
    });

    // Insert section header rows right before the first item of each group
    currentComparisonFeatures = [];
    let hasInsertedSpecialHeader = false;
    let hasInsertedBenefitHeader = false;
    let hasInsertedChargeHeader = false;
    const insertedCustomHeaders = new Set();
    filteredFeatures.forEach(feat => {
        if (feat.isSpecialProvision && !hasInsertedSpecialHeader) {
            currentComparisonFeatures.push({
                label: "Special Provisions",
                isSectionHeader: true
            });
            hasInsertedSpecialHeader = true;
        }
        if (feat.isBenefit && !hasInsertedBenefitHeader) {
            currentComparisonFeatures.push({
                label: "Product Benefits",
                isSectionHeader: true
            });
            hasInsertedBenefitHeader = true;
        }
        if (feat.customSectionTitle && !insertedCustomHeaders.has(feat.customSectionTitle)) {
            currentComparisonFeatures.push({
                label: feat.customSectionTitle,
                isSectionHeader: true
            });
            insertedCustomHeaders.add(feat.customSectionTitle);
        }
        if (feat.isCharge && !hasInsertedChargeHeader) {
            currentComparisonFeatures.push({
                label: "Charges",
                isSectionHeader: true,
                isChargeSection: true
            });
            hasInsertedChargeHeader = true;
        }
        currentComparisonFeatures.push(feat);
    });

    const prodCount = currentComparisonProducts.length;

    container.innerHTML = `
                <table class="w-full table-fixed text-xs sm:text-sm text-left border-collapse" style="min-width: calc(var(--cmp-feature-col-w, 220px) + ${prodCount} * var(--cmp-prod-col-min-w, 260px));" id="comparisonTable">
                    <colgroup>
                        <col style="width: var(--cmp-feature-col-w, 220px);" />
                        ${currentComparisonProducts.map(() => `<col style="width: calc((100% - var(--cmp-feature-col-w, 220px)) / ${prodCount});" />`).join('')}
                    </colgroup>
                    <thead>
                        <tr class="bg-gray-100 dark:bg-neutral-700 border-b dark:border-neutral-600">
                            <th scope="row" class="p-2 sm:p-3 font-semibold text-gray-700 dark:text-neutral-200 bg-gray-100 dark:bg-neutral-700 break-words" style="width: var(--cmp-feature-col-w, 220px);">Company Name</th>
                            ${currentComparisonProducts.map(p => `
                                <th scope="col" class="p-2 sm:p-3 font-semibold text-blue-700 dark:text-blue-300 border-l dark:border-neutral-600 bg-gray-100 dark:bg-neutral-700 break-words">
                                    ${escapeHtml(p.company)}
                                </th>
                            `).join('')}
                        </tr>
                        <tr class="bg-gray-100 dark:bg-neutral-700 border-b-2 dark:border-neutral-600">
                            <th scope="row" class="p-2 sm:p-3 font-semibold text-gray-700 dark:text-neutral-200 bg-gray-100 dark:bg-neutral-700 break-words" style="width: var(--cmp-feature-col-w, 220px);">Product Name</th>
                            ${currentComparisonProducts.map(p => `
                                <th scope="col" class="p-2 sm:p-3 font-bold text-blue-900 dark:text-blue-100 border-l dark:border-neutral-600 bg-gray-100 dark:bg-neutral-700 text-sm sm:text-base break-words">
                                    ${escapeHtml(p.name)}
                                </th>
                            `).join('')}
                        </tr>
                    </thead>
                    <tbody class="divide-y dark:divide-neutral-700">
                        ${currentComparisonFeatures.map(feat => {
        if (feat.isSectionHeader) {
            const secBg = feat.isChargeSection
                ? "bg-amber-50 dark:bg-neutral-900 border-y-2 border-amber-200 dark:border-amber-900/80"
                : "bg-blue-50 dark:bg-neutral-900 border-y-2 border-blue-200 dark:border-neutral-700";
            const secText = feat.isChargeSection
                ? "text-amber-900 dark:text-amber-300 bg-amber-50 dark:bg-neutral-900"
                : "text-blue-900 dark:text-blue-300 bg-blue-50 dark:bg-neutral-900";
            const secCellBg = feat.isChargeSection
                ? "bg-amber-50 dark:bg-neutral-900"
                : "bg-blue-50 dark:bg-neutral-900";
            return `
                                <tr class="${secBg}">
                                    <td class="p-2 sm:p-3 font-extrabold ${secText} uppercase tracking-wider text-[11px] sm:text-xs break-words">
                                        ${escapeHtml(feat.label)}
                                    </td>
                                    ${currentComparisonProducts.map(() => `<td class="p-2 sm:p-3 border-l dark:border-neutral-800 ${secCellBg}"></td>`).join('')}
                                </tr>
                    `;
        }

        const rowBgClass = feat.isCharge
            ? 'font-bold text-amber-900 dark:text-amber-300 bg-amber-50 dark:bg-neutral-900 border-t-2 border-amber-100 dark:border-amber-900/60'
            : (feat.isHighlight
                ? 'font-bold text-blue-800 dark:text-blue-300 bg-blue-50 dark:bg-neutral-900 border-t-2 border-blue-100 dark:border-neutral-700'
                : 'font-medium text-gray-600 dark:text-neutral-300 bg-gray-50 dark:bg-neutral-800');
        const cellBgClass = feat.isCharge
            ? 'text-xs bg-white dark:bg-neutral-800 text-gray-800 dark:text-neutral-200 border-t-2 border-amber-100 dark:border-amber-900/60'
            : (feat.isHighlight
                ? 'text-xs bg-white dark:bg-neutral-800 text-gray-800 dark:text-neutral-200 border-t-2 border-blue-100 dark:border-neutral-700'
                : 'bg-white dark:bg-neutral-800 text-gray-700 dark:text-neutral-300');

        return `
                                <tr>
                                    <td class="p-2 sm:p-3 align-middle break-words ${rowBgClass}">${escapeHtml(feat.label)}</td>
                                    ${currentComparisonProducts.map(p => {
            let val = feat.isAttr ? p.attributes[feat.key] : p[feat.key];
            const renderedVal = formatListText(val);
            const isDashEmpty = isEmptyVal(val) || renderedVal === "-";
            const alignClass = isDashEmpty
                ? "text-center align-middle font-semibold text-gray-400 dark:text-neutral-500"
                : "align-middle";
            return `<td class="p-2 sm:p-3 border-l dark:border-neutral-700 break-words ${cellBgClass} ${alignClass}">${renderedVal}</td>`;
        }).join('')}
                                </tr>
                            `;
    }).join('')}
                    </tbody>
                </table>
            `;
    applyCompareTableZoom(currentCompareTableZoom);
    compareModal.classList.remove("hidden");
    requestAnimationFrame(() => applyComparisonTableFreeze());
}

function closeComparisonModal() {
    document.getElementById("compareModal").classList.add("hidden");
}

async function resolveImageForExcelExport(rawImgSrc) {
    if (!rawImgSrc) return null;
    try {
        let src = String(rawImgSrc).trim();
        if (src.startsWith("drive:")) {
            const fileId = src.slice(6);
            src = driveImageCache[fileId] || await fetchDriveImageDataUrl(fileId);
            if (!src) return null;
        }

        const img = await new Promise((resolve, reject) => {
            const el = new Image();
            el.crossOrigin = "anonymous";
            el.onload = () => resolve(el);
            el.onerror = (e) => reject(e);
            el.src = src;
        });

        const nw = img.naturalWidth || img.width || 240;
        const nh = img.naturalHeight || img.height || 160;
        const canvas = document.createElement("canvas");
        canvas.width = nw;
        canvas.height = nh;
        const ctx = canvas.getContext("2d");
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, nw, nh);
        ctx.drawImage(img, 0, 0, nw, nh);

        const pngDataUrl = canvas.toDataURL("image/png");
        const b64 = pngDataUrl.split(",")[1];
        if (!b64) return null;

        const binStr = atob(b64);
        const bytes = new Uint8Array(binStr.length);
        for (let i = 0; i < binStr.length; i++) {
            bytes[i] = binStr.charCodeAt(i);
        }

        return {
            pngBytes: bytes,
            width: nw,
            height: nh
        };
    } catch (err) {
        console.warn("Gagal menyiapkan gambar untuk ekspor Excel:", err);
        return null;
    }
}

function buildFreezeSheetViewsXml(freezeRows = 0, freezeCols = 0) {
    const xSplit = Math.max(0, parseInt(freezeCols, 10) || 0);
    const ySplit = Math.max(0, parseInt(freezeRows, 10) || 0);
    if (xSplit === 0 && ySplit === 0) return "";
    const colRef = XLSX.utils.encode_col(xSplit);
    const rowRef = ySplit + 1;
    const topLeftCell = `${colRef}${rowRef}`;
    const activePane = (xSplit > 0 && ySplit > 0)
        ? "bottomRight"
        : (ySplit > 0 ? "bottomLeft" : "topRight");
    const paneAttrs = [
        xSplit > 0 ? `xSplit="${xSplit}"` : "",
        ySplit > 0 ? `ySplit="${ySplit}"` : "",
        `topLeftCell="${topLeftCell}"`,
        `activePane="${activePane}"`,
        `state="frozen"`
    ].filter(Boolean).join(" ");
    return `<sheetViews><sheetView tabSelected="1" workbookViewId="0"><pane ${paneAttrs}/><selection pane="${activePane}" activeCell="${topLeftCell}" sqref="${topLeftCell}"/></sheetView></sheetViews>`;
}

async function injectImagesIntoXlsxBuffer(xlsxArrayBuffer, excelDrawings, freezeRows = 0, freezeCols = 0) {
    const hasDrawings = Array.isArray(excelDrawings) && excelDrawings.length > 0;
    const hasFreeze = (freezeRows > 0 || freezeCols > 0);
    if ((!hasDrawings && !hasFreeze) || typeof JSZip === "undefined") {
        return xlsxArrayBuffer;
    }

    const zip = await JSZip.loadAsync(xlsxArrayBuffer);

    // 1. Terapkan Freeze Pane (<sheetViews><pane ... state="frozen"/></sheetViews>) jika aktif
    const sheetPath = "xl/worksheets/sheet1.xml";
    if (hasFreeze && zip.file(sheetPath)) {
        let sheetXml = await zip.file(sheetPath).async("string");
        const freezeXml = buildFreezeSheetViewsXml(freezeRows, freezeCols);
        if (freezeXml) {
            if (/<sheetViews\b[^>]*\/>|<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/i.test(sheetXml)) {
                sheetXml = sheetXml.replace(/<sheetViews\b[^>]*\/>|<sheetViews\b[^>]*>[\s\S]*?<\/sheetViews>/i, freezeXml);
            } else if (sheetXml.includes("<sheetFormatPr")) {
                sheetXml = sheetXml.replace("<sheetFormatPr", `${freezeXml}<sheetFormatPr`);
            } else if (sheetXml.includes("<cols")) {
                sheetXml = sheetXml.replace("<cols", `${freezeXml}<cols`);
            } else if (sheetXml.includes("<sheetData")) {
                sheetXml = sheetXml.replace("<sheetData", `${freezeXml}<sheetData`);
            }
            zip.file(sheetPath, sheetXml);
        }
    }

    if (!hasDrawings) {
        return await zip.generateAsync({ type: "uint8array" });
    }

    const EMU_PER_PX = 9525;

    let drawingAnchorsXml = "";
    let drawingRelsXml = "";

    excelDrawings.forEach((item, idx) => {
        const imgNum = idx + 1;
        const rId = `rId_img_${imgNum}`;
        zip.file(`xl/media/image${imgNum}.png`, item.pngBytes, { binary: true });

        drawingRelsXml += `<Relationship Id="${rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image${imgNum}.png"/>`;

        const colOffEmu = Math.max(0, Math.round(item.colOffPx * EMU_PER_PX));
        const rowOffEmu = Math.max(0, Math.round(item.rowOffPx * EMU_PER_PX));
        const cxEmu = Math.max(EMU_PER_PX * 16, Math.round(item.widthPx * EMU_PER_PX));
        const cyEmu = Math.max(EMU_PER_PX * 16, Math.round(item.heightPx * EMU_PER_PX));

        drawingAnchorsXml += `
                    <xdr:oneCellAnchor>
                        <xdr:from>
                            <xdr:col>${item.col}</xdr:col>
                            <xdr:colOff>${colOffEmu}</xdr:colOff>
                            <xdr:row>${item.row}</xdr:row>
                            <xdr:rowOff>${rowOffEmu}</xdr:rowOff>
                        </xdr:from>
                        <xdr:ext cx="${cxEmu}" cy="${cyEmu}"/>
                        <xdr:pic>
                            <xdr:nvPicPr>
                                <xdr:cNvPr id="${imgNum + 10}" name="TableImage_${imgNum}"/>
                                <xdr:cNvPicPr>
                                    <a:picLocks noChangeAspect="1"/>
                                </xdr:cNvPicPr>
                            </xdr:nvPicPr>
                            <xdr:blipFill>
                                <a:blip r:embed="${rId}"/>
                                <a:stretch><a:fillRect/></a:stretch>
                            </xdr:blipFill>
                            <xdr:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="${cxEmu}" cy="${cyEmu}"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                            </xdr:spPr>
                        </xdr:pic>
                        <xdr:clientData/>
                    </xdr:oneCellAnchor>
                `.trim();
    });

    const drawingXmlContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">` +
        drawingAnchorsXml +
        `</xdr:wsDr>`;
    zip.file("xl/drawings/drawing1.xml", drawingXmlContent);

    const drawingRelsContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
        `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        drawingRelsXml +
        `</Relationships>`;
    zip.file("xl/drawings/_rels/drawing1.xml.rels", drawingRelsContent);

    const sheetRelsPath = "xl/worksheets/_rels/sheet1.xml.rels";
    let sheetRelsXml = "";
    if (zip.file(sheetRelsPath)) {
        sheetRelsXml = await zip.file(sheetRelsPath).async("string");
        if (!sheetRelsXml.includes("drawings/drawing1.xml")) {
            sheetRelsXml = sheetRelsXml.replace(
                "</Relationships>",
                `<Relationship Id="rId_drawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/></Relationships>`
            );
        }
    } else {
        sheetRelsXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
            `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
            `<Relationship Id="rId_drawing1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing1.xml"/>` +
            `</Relationships>`;
    }
    zip.file(sheetRelsPath, sheetRelsXml);

    if (zip.file(sheetPath)) {
        let sheetXml = await zip.file(sheetPath).async("string");
        if (!sheetXml.includes('xmlns:r=')) {
            sheetXml = sheetXml.replace(
                '<worksheet ',
                '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
            );
        }
        if (!sheetXml.includes('<drawing ')) {
            if (sheetXml.includes('<legacyDrawing')) {
                sheetXml = sheetXml.replace('<legacyDrawing', '<drawing r:id="rId_drawing1"/><legacyDrawing');
            } else if (sheetXml.includes('<extLst')) {
                sheetXml = sheetXml.replace('<extLst', '<drawing r:id="rId_drawing1"/><extLst');
            } else {
                sheetXml = sheetXml.replace('</worksheet>', '<drawing r:id="rId_drawing1"/></worksheet>');
            }
        }
        zip.file(sheetPath, sheetXml);
    }

    const contentTypesPath = "[Content_Types].xml";
    if (zip.file(contentTypesPath)) {
        let ctXml = await zip.file(contentTypesPath).async("string");
        if (!/Extension="png"/i.test(ctXml)) {
            ctXml = ctXml.replace(/<Types[^>]*>/, match => `${match}<Default Extension="png" ContentType="image/png"/>`);
        }
        if (!ctXml.includes('/xl/drawings/drawing1.xml')) {
            ctXml = ctXml.replace(
                '</Types>',
                `<Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/></Types>`
            );
        }
        zip.file(contentTypesPath, ctXml);
    }

    return await zip.generateAsync({ type: "uint8array" });
}

async function downloadComparisonExcel(btnEl = null) {
    if (currentComparisonProducts.length === 0) return;
    currentComparisonProducts = sortComparisonProductsAstraFirst(currentComparisonProducts);

    const exportBtn = btnEl || document.getElementById("btnExportComparisonExcel");
    const origBtnHtml = exportBtn ? exportBtn.innerHTML : "";
    if (exportBtn) {
        exportBtn.disabled = true;
        exportBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Preparing Excel...`;
    }

    try {
        await loadExcelLibs();
        let ws_data = [];
        const FIRST_COL_WCH = 28;
        const PROD_COL_WCH = 38;
        const PROD_COL_WPX = 270;
        const MAX_IMG_DISPLAY_W = 248;

        // Row 1: Company Name
        let companyRow = ["Company Name"];
        currentComparisonProducts.forEach(p => companyRow.push(p.company || "-"));
        ws_data.push(companyRow);

        // Row 2: Product Name
        let productRow = ["Product Name"];
        currentComparisonProducts.forEach(p => productRow.push(p.name || "-"));
        ws_data.push(productRow);

        const excelDrawings = [];
        const cellRequiredHeightPx = {};
        const cellHasEmbeddedImages = {};

        const formatExcelCellText = (rawText) => {
            if (!rawText) return "";
            const dplkParsed = parseDplkBenefitValue(rawText);
            if (dplkParsed.isChecked) {
                return dplkParsed.remark ? `✓ ${dplkParsed.remark}` : "✓";
            }
            return String(rawText).replace(/\r\n|\r|\n/g, '\n').trim();
        };

        for (let fIdx = 0; fIdx < currentComparisonFeatures.length; fIdx++) {
            const feat = currentComparisonFeatures[fIdx];
            const rIdx = ws_data.length;

            if (feat.isSectionHeader) {
                let sectionRow = [feat.label, ...currentComparisonProducts.map(() => "")];
                ws_data.push(sectionRow);
                continue;
            }

            let rowData = [feat.label];
            cellRequiredHeightPx[rIdx] = [];

            for (let pIdx = 0; pIdx < currentComparisonProducts.length; pIdx++) {
                const p = currentComparisonProducts[pIdx];
                const cIdx = pIdx + 1;
                let rawCellVal = feat.isAttr ? p.attributes[feat.key] : p[feat.key];
                const { text, images } = splitTextAndImages(rawCellVal);

                let cleanText = text ? formatExcelCellText(text) : "";

                if (!images || images.length === 0) {
                    const fallbackVal = cleanText || formatExcelCellText(sanitizeValue(rawCellVal));
                    rowData.push(String(fallbackVal).replace(/\r\n|\r|\n/g, '\n'));
                    continue;
                }

                const resolvedImgs = [];
                for (const imgSrc of images) {
                    const resImg = await resolveImageForExcelExport(imgSrc);
                    if (resImg) resolvedImgs.push(resImg);
                }

                if (resolvedImgs.length === 0) {
                    let fallbackStr = formatExcelCellText(stripImageTokensForExcel(rawCellVal));
                    rowData.push(String(fallbackStr).replace(/\r\n|\r|\n/g, '\n'));
                    continue;
                }

                // When images are embedded, keep only cleanText (without placeholder text)
                rowData.push(cleanText);
                cellHasEmbeddedImages[`${rIdx},${cIdx}`] = true;

                // Calculate text space above image (if cell has both text and image)
                let textOffsetPx = 6;
                if (cleanText) {
                    const effectiveCharLimit = 46;
                    const paragraphs = cleanText.split('\n');
                    let totalTextLines = 0;
                    paragraphs.forEach(line => {
                        totalTextLines += Math.max(1, Math.ceil(line.length / effectiveCharLimit));
                    });
                    textOffsetPx = totalTextLines * 18 + 8;
                }

                // Calculate display size of each image inside the cell
                const sizedImgs = resolvedImgs.map(im => {
                    let dispW = Math.min(MAX_IMG_DISPLAY_W, im.width);
                    if (im.width < 150) {
                        dispW = Math.min(MAX_IMG_DISPLAY_W, Math.round(im.width * 1.15));
                    }
                    let dispH = Math.max(24, Math.round(im.height * (dispW / im.width)));
                    return { ...im, dispW, dispH };
                });

                // Cap total height per cell within Excel's max row height limit (409 pt ~ 540 px)
                const maxAvailImgHeightPx = Math.max(85, 525 - textOffsetPx - 10);
                const rawTotalImgHeightPx = sizedImgs.reduce((sum, im) => sum + im.dispH, 0) + (sizedImgs.length - 1) * 6;
                if (rawTotalImgHeightPx > maxAvailImgHeightPx) {
                    const scaleDown = maxAvailImgHeightPx / rawTotalImgHeightPx;
                    sizedImgs.forEach(im => {
                        im.dispW = Math.max(36, Math.round(im.dispW * scaleDown));
                        im.dispH = Math.max(20, Math.round(im.dispH * scaleDown));
                    });
                }

                let curY = textOffsetPx;
                sizedImgs.forEach(im => {
                    const colOffPx = cleanText
                        ? 6
                        : Math.max(6, Math.round((PROD_COL_WPX - im.dispW) / 2));
                    excelDrawings.push({
                        col: cIdx,
                        row: rIdx,
                        colOffPx,
                        rowOffPx: curY,
                        widthPx: im.dispW,
                        heightPx: im.dispH,
                        pngBytes: im.pngBytes
                    });
                    curY += im.dispH + 6;
                });

                cellRequiredHeightPx[rIdx].push(curY + 6);
            }

            ws_data.push(rowData);
        }

        const ws = XLSX.utils.aoa_to_sheet(ws_data);

        // 1. Column Widths
        ws['!cols'] = [
            { wch: FIRST_COL_WCH },
            ...currentComparisonProducts.map(() => ({ wch: PROD_COL_WCH, wpx: PROD_COL_WPX }))
        ];

        // 2. Dynamic & Tight Row Heights: measure exact wrapped lines using Calibri 10pt metrics,
        //    fit to the tallest cell in each row, and add +10 px breathing room.
        const measureCanvas = document.createElement('canvas');
        const measureCtx = measureCanvas.getContext('2d');
        const countWrappedLines = (text, maxWidthPx, isBold) => {
            const str = String(text || "").trim();
            if (!str) return 1;
            if (!measureCtx) {
                const approxChars = Math.max(20, Math.floor(maxWidthPx / 5.2));
                return str.split('\n').reduce((acc, ln) => acc + Math.max(1, Math.ceil(ln.trim().length / approxChars)), 0);
            }
            measureCtx.font = `${isBold ? 'bold ' : ''}10pt Calibri, Arial, sans-serif`;
            let totalLines = 0;
            const paragraphs = str.split('\n');
            paragraphs.forEach(para => {
                const trimmed = para.replace(/\s+$/g, '');
                if (!trimmed) {
                    totalLines += 1;
                    return;
                }
                if (measureCtx.measureText(trimmed).width <= maxWidthPx) {
                    totalLines += 1;
                    return;
                }
                const words = trimmed.split(/(\s+)/);
                let currentLine = "";
                let paraLines = 1;
                for (let i = 0; i < words.length; i++) {
                    const token = words[i];
                    if (!token.trim()) {
                        currentLine += token;
                        continue;
                    }
                    const testLine = currentLine + token;
                    if (measureCtx.measureText(testLine).width <= maxWidthPx) {
                        currentLine = testLine;
                    } else {
                        if (currentLine.trim().length > 0) {
                            paraLines += 1;
                            currentLine = token;
                        } else {
                            currentLine = token;
                        }
                        // Handle single long word/URL wider than the column
                        while (measureCtx.measureText(currentLine).width > maxWidthPx && currentLine.length > 1) {
                            let cut = currentLine.length - 1;
                            while (cut > 1 && measureCtx.measureText(currentLine.slice(0, cut)).width > maxWidthPx) {
                                cut--;
                            }
                            paraLines += 1;
                            currentLine = currentLine.slice(cut);
                        }
                    }
                }
                totalLines += paraLines;
            });
            return Math.max(1, totalLines);
        };

        ws['!rows'] = ws_data.map((row, rIdx) => {
            const featObj = rIdx >= 2 ? currentComparisonFeatures[rIdx - 2] : null;
            if (featObj && featObj.isSectionHeader) {
                return { hpt: 22 };
            }

            let maxWrappedLines = 1;
            row.forEach((cellText, cIdx) => {
                const str = String(cellText || "").trim();
                if (!str) return;
                // Match Excel wch: 38 (~271px) / col 0 (wch: 28, ~170px)
                const usableWidthPx = (cIdx === 0) ? 170 : PROD_COL_WPX;
                const isBoldCell = (cIdx === 0 || rIdx === 0 || rIdx === 1);
                const lines = countWrappedLines(str, usableWidthPx, isBoldCell);
                if (lines > maxWrappedLines) {
                    maxWrappedLines = lines;
                }
            });

            // Excel AutoFit height for Calibri 10pt is 12.75pt/line (= 16 * 79.6875%), 11pt header is ~14pt/line.
            // Omit hpx so xlsx-js-style writes hpt directly to <row ht="...">.
            const linePt = (rIdx === 1) ? 14 : 12.75;
            let rowHpt = maxWrappedLines * linePt;

            const imgHeightsPx = cellRequiredHeightPx[rIdx] || [];
            if (imgHeightsPx.length > 0) {
                const maxImgPx = Math.max(...imgHeightsPx);
                rowHpt = Math.max(rowHpt, Math.ceil(maxImgPx * 0.75));
            }

            // Add +2.5pt breathing room and cap at Excel max row height (409pt)
            rowHpt += 2.5;
            rowHpt = Math.min(409, Math.round(rowHpt * 100) / 100);
            return { hpt: rowHpt };
        });

        // 3. Apply Wrap Text, Alignment, and Black Borders
        const thinBorder = {
            top: { style: "thin", color: { rgb: "000000" } },
            bottom: { style: "thin", color: { rgb: "000000" } },
            left: { style: "thin", color: { rgb: "000000" } },
            right: { style: "thin", color: { rgb: "000000" } }
        };

        const entryAgeRowIdxFound = ws_data.findIndex(r => String(r[0] || "").trim().toLowerCase() === "entry age");
        const topAlignStartRow = entryAgeRowIdxFound !== -1 ? entryAgeRowIdxFound : 5;
        const totalCols = companyRow.length;
        for (let r = 0; r < ws_data.length; r++) {
            const featObj = r >= 2 ? currentComparisonFeatures[r - 2] : null;
            const isSectionRow = Boolean(featObj && featObj.isSectionHeader);
            const isHeaderRow = (r === 0 || r === 1);
            const useTopAlign = (r >= topAlignStartRow && !isSectionRow);

            for (let c = 0; c < totalCols; c++) {
                const cellRef = XLSX.utils.encode_cell({ r, c });
                if (!ws[cellRef]) {
                    ws[cellRef] = { t: 's', v: '' };
                }

                const isBold = isHeaderRow || (c === 0) || isSectionRow;
                const cellTextVal = String(ws[cellRef].v || "").trim();
                const hasImgInCell = Boolean(cellHasEmbeddedImages[`${r},${c}`]);
                const isCenterCell = (c > 0 && !isHeaderRow && !isSectionRow && !hasImgInCell && (cellTextVal === "-" || cellTextVal === "✓" || cellTextVal === "✅"));
                ws[cellRef].s = {
                    font: {
                        name: "Calibri",
                        sz: (r === 1 && c > 0) ? 11 : 10,
                        bold: isBold
                    },
                    alignment: {
                        wrapText: true,
                        vertical: (useTopAlign || (hasImgInCell && cellTextVal)) ? "top" : "center",
                        horizontal: isCenterCell ? "center" : "left"
                    },
                    border: thinBorder
                };
            }
        }

        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, "Comparison");

        const today = new Date().toISOString().split('T')[0];
        const fileName = `Product_Comparison_${today}.xlsx`;

        const rawXlsxArray = XLSX.write(wb, { type: 'array', bookType: 'xlsx' });
        const effectiveFreezeRows = Math.min(compareFreezeRows, Math.max(0, ws_data.length - 1));
        const effectiveFreezeCols = Math.min(compareFreezeCols, Math.max(0, totalCols - 1));
        const finalXlsxBytes = await injectImagesIntoXlsxBuffer(rawXlsxArray, excelDrawings, effectiveFreezeRows, effectiveFreezeCols);

        if (window.pywebview) {
            let binary = "";
            const bytes = new Uint8Array(finalXlsxBytes);
            for (let i = 0; i < bytes.byteLength; i++) {
                binary += String.fromCharCode(bytes[i]);
            }
            const b64Data = btoa(binary);
            window.pywebview.api.save_excel(b64Data, fileName);
        } else {
            const outBlob = new Blob([finalXlsxBytes], {
                type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            });
            triggerBlobDownload(outBlob, fileName);
        }
    } catch (err) {
        console.error("Failed to export comparison Excel:", err);
        showToast("An error occurred while exporting the Excel file: " + err.message);
    } finally {
        if (exportBtn) {
            exportBtn.disabled = false;
            exportBtn.innerHTML = origBtnHtml;
        }
    }
}

const DEFAULT_SYARIAH_SPEC_COLUMNS = [
    "Aqd",
    "Alokasi Kontribusi",
    "Surplus Underwriting"
];
let SYARIAH_SPEC_COLUMNS = [...DEFAULT_SYARIAH_SPEC_COLUMNS];

const DEFAULT_DPLK_SPEC_COLUMNS = [
    "Sales Tools",
    "After Sales Tools"
];
let DPLK_SPEC_COLUMNS = [...DEFAULT_DPLK_SPEC_COLUMNS];

const DEFAULT_DPLK_BENEFIT_RULES = [
    { canonical: "Pensiun Normal", aliases: ["pensiun normal", "normal"] },
    { canonical: "Pensiun Ditunda", aliases: ["pensiun ditunda", "dana pensiun ditunda", "ditunda"] },
    { canonical: "Pensiun Cacat / Disabilitas", aliases: ["cacat", "disabilitas"] },
    { canonical: "Pensiun Dipercepat", aliases: ["dipercepat"] },
    { canonical: "Pensiun Resign / Berhenti Kerja", aliases: ["resign", "berhenti", "phk"] },
    { canonical: "Pensiun Meninggal Dunia", aliases: ["meninggal", "janda", "duda"] },
    { canonical: "Pensiun WNI Pindah Kewarganegaraan", aliases: ["wni", "kewarganegaraan"] },
    { canonical: "Pensiun WNA Kembali ke Negara Asal", aliases: ["wna", "negara asal", "expatriate", "ekspatriat"] },
    { canonical: "Pensiun Sakit Keras", aliases: ["sakit keras", "kritis"] },
    { canonical: "Pengalihan ke DPLK Lain", aliases: ["pengalihan dplk"] },
    { canonical: "Manfaat Lain (Dana Kompensasi Pasca Kerja)", aliases: ["manfaat lain", "kompensasi pasca kerja", "dkp", "asuransi tambahan", "penarikan"] }
];
let DPLK_BENEFIT_RULES = DEFAULT_DPLK_BENEFIT_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
let DPLK_BENEFIT_COLUMNS = DPLK_BENEFIT_RULES.map(r => r.canonical);

const DEFAULT_DPLK_CHARGE_RULES = [
    { canonical: "Biaya Pendaftaran", aliases: ["pendaftaran", "kepesertaan baru", "registrasi"] },
    { canonical: "Biaya Administrasi Pembayaran Manfaat Pensiun", aliases: ["pembayaran manfaat", "pencairan"] },
    { canonical: "Biaya Administrasi Iuran", aliases: ["iuran", "kontribusi", "top up"] },
    { canonical: "Biaya Administrasi Kepesertaan", aliases: ["administrasi", "administration", "biaya admin"] },
    { canonical: "Biaya Perubahan Pilihan Investasi (Switching)", aliases: ["perubahan", "switching", "pengalihan investasi", "pengalihan paket"] },
    { canonical: "Biaya Pengelolaan Investasi", aliases: ["pengelolaan investasi", "biaya investasi", "management fee"] },
    { canonical: "Biaya Pengelolaan Dana Peserta", aliases: ["pengelolaan"] },
    { canonical: "Biaya Penarikan Sebagian Iuran", aliases: ["penarikan"] },
    { canonical: "Biaya Pengalihan Dana ke DPLK Lain", aliases: ["pengalihan", "transfer ke dplk"] },
    { canonical: "Biaya Pencetakan Laporan Rekening", aliases: ["pencetakan", "laporan"] },
    { canonical: "Biaya Kartu Kepesertaan", aliases: ["kartu", "buku"] },
    { canonical: "Biaya Penghentian Kerja Sama / Pemutusan", aliases: ["penghentian", "pemutusan", "penutupan", "terminasi"] },
    { canonical: "Biaya Lain (Pihak Ketiga)", aliases: ["pihak ketiga", "kustodian", "pajak", "materai", "biaya lain"] }
];
let DPLK_CHARGE_RULES = DEFAULT_DPLK_CHARGE_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
let DPLK_CHARGE_COLUMNS = DPLK_CHARGE_RULES.map(r => r.canonical);

const DEFAULT_STANDARD_BENEFIT_COLUMNS = [
    "Accident Death During Hajj (Pilgrim) Benefit",
    "Total Permanent Disability Benefit",
    "Additional Death Benefit",
    "Accidental Death Benefit",
    "Accident Death Benefit",
    "Terminal Illness Benefit",
    "Critical Illness Benefit",
    "Remaining Premium Waived",
    "Education Fund Benefit",
    "Hospitalization Benefit",
    "Manfaat Dana Tahapan",
    "Sum Assured Bonus",
    "Cash Payout Bonus",
    "Maturity Benefit",
    "Survival Benefit",
    "Booster Benefit",
    "Premium Waiver",
    "Death Benefit",
    "Loyalty Bonus"
];
let STANDARD_BENEFIT_COLUMNS = [...DEFAULT_STANDARD_BENEFIT_COLUMNS];

const DEFAULT_COMPANY_KEYWORD_RULES = [
    { canonical: "Manulife", aliases: ["manulife"] },
    { canonical: "Astra Life", aliases: ["ava", "astra", "asli", "asuransi jiwa astra"] },
    { canonical: "AIA", aliases: ["aia"] },
    { canonical: "Allianz", aliases: ["allianz", "allisya"] },
    { canonical: "Prudential", aliases: ["pru", "prudential"] },
    { canonical: "AXA", aliases: ["mandiri", "axa", "maestro"] },
    { canonical: "FWD", aliases: ["fwd"] },
    { canonical: "Great Eastern", aliases: ["great", "great eastern"] },
    { canonical: "Generali", aliases: ["generali", "isalaam"] },
    { canonical: "Sunlife", aliases: ["sun", "sunlife", "salam", "berkah"] }
];
let COMPANY_KEYWORD_RULES = DEFAULT_COMPANY_KEYWORD_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));

const DEFAULT_CORE_FIELD_RULES = [
    { canonical: "Product Name", aliases: ["nama produk"] },
    { canonical: "Company Name", aliases: ["nama perusahaan", "company"] },
    { canonical: "Par/Non Par", aliases: ["par / non par"] },
    { canonical: "Product Type", aliases: ["tipe produk"] },
    { canonical: "Distribution Channel", aliases: ["channel distribution"] },
    { canonical: "Currency", aliases: ["mata uang"] },
    { canonical: "Launch Date", aliases: ["tanggal rilis", "launching date", "release date", "estimation launching date", "estimated launching date", "est. launching date", "est launching date"] },
    { canonical: "Remarks", aliases: ["catatan"] }
];
let CORE_FIELD_RULES = DEFAULT_CORE_FIELD_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));

const DEFAULT_NON_BENEFIT_ALIAS_RULES = [
    { canonical: "Entry Age", aliases: ["insured entry age", "age definition", "usia masuk"] },
    { canonical: "Coverage Period", aliases: ["coverage term", "masa pertanggungan"] },
    { canonical: "Premium Payment Period (PPP)", aliases: ["ppp", "premium payment period", "premium term", "masa pembayaran premi"] },
    { canonical: "Min Premium", aliases: ["premium (iuran)", "iuran", "premium limit", "minimum premium", "premium"] },
    { canonical: "Min Sum Assured", aliases: ["sum assured limit", "minimum sum assured", "sum assured"] },
    { canonical: "Underwriting Method", aliases: ["underwriting"] },
    { canonical: "Available Riders", aliases: ["available rider", "riders", "rider"] },
    { canonical: "Additional Benefit", aliases: ["additional feature (if any)", "additional feature", "additional features", "other benefit", "other feature", "fitur tambahan"] },
    { canonical: "Aqd", aliases: ["akad", "aqd / akad", "akad / aqd"] },
    { canonical: "Alokasi Kontribusi", aliases: ["tabarru", "tanahud", "ujrah", "nisbah"] },
    { canonical: "Surplus Underwriting", aliases: ["underwriting surplus"] },
    { canonical: "Sales Tools", aliases: ["sales tool"] },
    { canonical: "After Sales Tools", aliases: ["after sales tool"] }
];

const DEFAULT_NON_BENEFIT_COLUMNS = [
    "Launch Date",
    "Tanggal Rilis",
    "Additional Benefit",
    "Aqd",
    "Akad",
    "Aqd / Akad",
    "Alokasi Kontribusi",
    "Surplus Underwriting",
    "Entry Age",
    "Coverage Period",
    "Premium Payment Period (PPP)",
    "Min Premium",
    "Min Sum Assured",
    "Underwriting Method",
    "Available Riders",
    "Sales Tools",
    "After Sales Tools",
    "Remarks"
];
let NON_BENEFIT_COLUMNS = [...DEFAULT_NON_BENEFIT_COLUMNS];
const STANDARD_FIXED_SPEC_COLS = new Set(DEFAULT_NON_BENEFIT_COLUMNS.map(c => c.toLowerCase()));

let CATEGORY_CONFIGS = {};

const DEFAULT_GENERAL_PROVISION_DEFS = [
    { col: "Entry Age", label: "Entry Age", placeholder: "e.g. Min 18 years old or married / 18 - 70 years old", fullWidth: false },
    { col: "Coverage Period", label: "Coverage Period", placeholder: "e.g. Up to age 100", fullWidth: false },
    { col: "Premium Payment Period (PPP)", label: "Premium Payment Period (PPP)", placeholder: "e.g. Single / 5 / 10 years", fullWidth: false },
    { col: "Underwriting Method", label: "Underwriting Method", placeholder: "e.g. Full Underwriting / GIO", fullWidth: false },
    { col: "Min Premium", label: "Min Premium", dplkLabel: "Min Premium / Contribution", placeholder: "e.g. Rp 100,000 / month or Rp 5,000,000 / year", fullWidth: false },
    { col: "Min Sum Assured", label: "Min Sum Assured", placeholder: "e.g. Rp 100,000,000", fullWidth: false },
    { col: "Available Riders", label: "Available Riders", placeholder: "Enter available riders (if any)...", fullWidth: true }
];

function isSyariahSpecColumn(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    return ["aqd", "akad", "aqd / akad", "alokasi kontribusi", "surplus underwriting"].includes(lower)
        || SYARIAH_SPEC_COLUMNS.some(c => c.toLowerCase() === lower);
}

function isDplkSpecColumn(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    return ["sales tools", "after sales tools"].includes(lower)
        || DPLK_SPEC_COLUMNS.some(c => c.toLowerCase() === lower);
}

function isDplkBenefitColumn(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    if (DPLK_BENEFIT_COLUMNS.some(b => b.toLowerCase() === lower)) return true;
    return Object.values(CATEGORY_CONFIGS || {}).some(cfg =>
        (cfg.productSections || []).some(sec => sec.mode === "checklist" && (sec.items || []).some(it => it.toLowerCase() === lower))
    );
}

function isDplkChargeColumn(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    if (!lower || isNonBenefitColumn(colName)) return false;
    return DPLK_CHARGE_COLUMNS.some(c => c.toLowerCase() === lower) || /^biaya\b/i.test(lower);
}

function isNonBenefitColumn(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    return NON_BENEFIT_COLUMNS.some(nb => nb.toLowerCase() === lower) || isSpecialProvisionColumnAnyCategory(colName);
}

function isSpecialProvisionColumnAnyCategory(colName) {
    const lower = String(colName || "").trim().toLowerCase();
    if (!lower) return false;
    if (isSyariahSpecColumn(lower) || isDplkSpecColumn(lower)) return true;
    return Object.values(CATEGORY_CONFIGS || {}).some(cfg =>
        (cfg.specialProvisions || []).some(sp => String(sp.name || "").trim().toLowerCase() === lower)
    );
}

function getCategorySchemaConfig(categoryName) {
    const rawCat = String(categoryName || "").trim();
    const lowerCat = rawCat.toLowerCase();
    const isSharia = /^(sharia|syariah)$/i.test(rawCat);
    const isDplk = /^dplk$/i.test(rawCat);

    const matchedKey = Object.keys(CATEGORY_CONFIGS || {}).find(k => k.toLowerCase() === lowerCat);
    const stored = matchedKey ? CATEGORY_CONFIGS[matchedKey] : null;

    const defaultGeneral = isDplk
        ? ["Entry Age", "Min Premium"]
        : DEFAULT_GENERAL_PROVISION_DEFS.map(d => d.col);

    const defaultSpecial = isSharia
        ? [
            { name: "Aqd", aliases: ["Akad", "Aqd / Akad", "Akad / Aqd"], allowImage: false, description: "Jenis akad produk Syariah" },
            { name: "Alokasi Kontribusi", aliases: ["Tabarru", "Tanahud", "Ujrah", "Nisbah"], allowImage: true, description: "Ketentuan alokasi kontribusi Syariah" },
            { name: "Surplus Underwriting", aliases: ["Underwriting Surplus"], allowImage: false, description: "Ketentuan pembagian surplus underwriting Syariah" }
        ]
        : (isDplk
            ? [
                { name: "Sales Tools", aliases: ["Sales Tool"], allowImage: false, description: "Alat bantu penjualan khusus DPLK" },
                { name: "After Sales Tools", aliases: ["After Sales Tool"], allowImage: false, description: "Alat bantu layanan purna jual khusus DPLK" }
            ]
            : []);

    const defaultSections = isDplk
        ? [
            { title: "Product Benefits", mode: "checklist", items: [...DPLK_BENEFIT_COLUMNS], isChargeLike: false, isPrimaryBenefits: true },
            { title: "DPLK Charge Components (Charges)", mode: "standard", items: [...DPLK_CHARGE_COLUMNS], isChargeLike: true, isPrimaryBenefits: false }
        ]
        : [
            { title: "Product Benefits", mode: "standard", items: [], isChargeLike: false, isPrimaryBenefits: true }
        ];

    const generalProvisions = (stored && Array.isArray(stored.generalProvisions))
        ? [...stored.generalProvisions]
        : [...defaultGeneral];

    const specialProvisions = (stored && Array.isArray(stored.specialProvisions))
        ? stored.specialProvisions.map(sp => ({
            name: String(sp.name || "").trim(),
            aliases: Array.isArray(sp.aliases) ? [...sp.aliases] : [],
            allowImage: Boolean(sp.allowImage),
            description: String(sp.description || "").trim()
        })).filter(sp => sp.name)
        : defaultSpecial;

    let productSections = (stored && Array.isArray(stored.productSections) && stored.productSections.length > 0)
        ? stored.productSections.map((sec, idx) => {
            const title = String(sec.title || "").trim() || `Section ${idx + 1}`;
            const mode = String(sec.mode || "standard").toLowerCase() === "checklist" ? "checklist" : "standard";
            const items = Array.isArray(sec.items) ? [...sec.items] : [];
            const isChargeLike = /charge|biaya/i.test(title) || (items.length > 0 && items.every(it => isDplkChargeColumn(it)));
            const isPrimaryBenefits = !isChargeLike && (/benefit|manfaat/i.test(title) || idx === 0);
            return { title, mode, items, isChargeLike, isPrimaryBenefits };
        })
        : defaultSections;

    // Pastikan jika ini DPLK tanpa productSections eksplisit tetap membawa defaultSections
    if (isDplk && (!stored || !Array.isArray(stored.productSections) || stored.productSections.length === 0)) {
        productSections = defaultSections;
    }

    return {
        sheetName: rawCat || (SHEET_TABS[0] || "Conventional"),
        generalProvisions,
        specialProvisions,
        productSections
    };
}

function getAvailableGeneralProvisions() {
    const base = DEFAULT_GENERAL_PROVISION_DEFS.map(d => d.col);
    const custom = getCustomNonBenefitColumns();
    const merged = [...base];
    custom.forEach(c => {
        if (!merged.some(m => m.toLowerCase() === c.toLowerCase())) merged.push(c);
    });
    return merged;
}

function getAllKnownNonBenefitColumns() {
    const specFromCats = [];
    Object.values(CATEGORY_CONFIGS || {}).forEach(cfg => {
        (cfg.specialProvisions || []).forEach(sp => {
            if (sp.name) specFromCats.push(sp.name);
        });
    });
    return [...new Set([
        ...NON_BENEFIT_COLUMNS,
        ...specFromCats,
        ...dynamicColumns.filter(col => isNonBenefitColumn(col))
    ])];
}

function getCustomNonBenefitColumns() {
    const coreCols = new Set(["product name", "company name", "par/non par", "product type", "distribution channel", "currency", "launch date", "tanggal rilis"]);
    CORE_FIELD_RULES.forEach(r => {
        if (r && r.canonical) coreCols.add(String(r.canonical).trim().toLowerCase());
        (Array.isArray(r.aliases) ? r.aliases : []).forEach(a => coreCols.add(String(a || "").trim().toLowerCase()));
    });
    return getAllKnownNonBenefitColumns().filter(col => {
        const lw = String(col || "").trim().toLowerCase();
        return lw && !coreCols.has(lw) && !STANDARD_FIXED_SPEC_COLS.has(lw) && !isSpecialProvisionColumnAnyCategory(lw);
    });
}

function getSectionAssignedNonPrimaryColumnsForCategory(categoryName) {
    const cfg = getCategorySchemaConfig(categoryName);
    const set = new Set();
    (cfg.productSections || []).forEach((sec, idx) => {
        if (idx === 0 && !sec.isChargeLike) return;
        (sec.items || []).forEach(it => set.add(String(it || "").trim().toLowerCase()));
    });
    return set;
}

function getBenefitColumns() {
    return dynamicColumns.filter(col => !isNonBenefitColumn(col));
}

function getRegularBenefitColumns() {
    return dynamicColumns.filter(col => !isNonBenefitColumn(col) && !isDplkChargeColumn(col));
}

function getChargeColumns() {
    return dynamicColumns.filter(col => !isNonBenefitColumn(col) && isDplkChargeColumn(col));
}

// Kolom yang dikirim ke sheet tujuan saat Simpan: hanya kolom yang sudah ada di sheet itu, kolom ketentuan
// standar / spesial milik kategori itu, atau kolom yang benar-benar terisi.
function getColumnsForTargetSheet(targetSheet, attrValues) {
    const catCfg = getCategorySchemaConfig(targetSheet);
    const isSharia = /^(sharia|syariah)$/i.test(String(targetSheet || "").trim());
    const isDplk = /^dplk$/i.test(String(targetSheet || "").trim());
    const sheetCols = sheetColumnsMap[targetSheet];
    const inSheet = sheetCols ? new Set(sheetCols) : null;
    const catSpecialSet = new Set(catCfg.specialProvisions.map(sp => sp.name.toLowerCase()));
    const standard = [
        "additional benefit",
        ...catCfg.generalProvisions.map(c => c.toLowerCase()),
        ...catCfg.specialProvisions.map(sp => sp.name.toLowerCase())
    ];

    return dynamicColumns.filter(col => {
        const lower = col.toLowerCase();
        const hasVal = attrValues && !isEmptyVal(attrValues[col]);
        if (isSpecialProvisionColumnAnyCategory(col) && !catSpecialSet.has(lower)) {
            return hasVal;
        }
        if (!isSharia && isSyariahSpecColumn(col) && !catSpecialSet.has(lower)) {
            return hasVal;
        }
        if (!isDplk && (isDplkSpecColumn(col) || isDplkChargeColumn(col) || isDplkBenefitColumn(col)) && !catSpecialSet.has(lower)) {
            return hasVal;
        }
        if (!inSheet) return true;
        return inSheet.has(lower) || standard.includes(lower) || hasVal;
    });
}

function filterSchemaAndRowsForTargetSheet(targetSheet, baseMasterSchema, rowsMatrix) {
    const catCfg = getCategorySchemaConfig(targetSheet);
    const isSharia = /^(sharia|syariah)$/i.test(String(targetSheet || "").trim());
    const isDplk = /^dplk$/i.test(String(targetSheet || "").trim());
    const sheetCols = sheetColumnsMap[targetSheet];
    const inSheet = sheetCols ? new Set(sheetCols) : null;
    const catSpecialSet = new Set(catCfg.specialProvisions.map(sp => sp.name.toLowerCase()));
    const standard = new Set([
        "additional benefit",
        ...catCfg.generalProvisions.map(c => c.toLowerCase()),
        ...catCfg.specialProvisions.map(sp => sp.name.toLowerCase())
    ]);

    const keepIndices = [];
    baseMasterSchema.forEach((col, idx) => {
        if (idx < 7) {
            keepIndices.push(idx);
            return;
        }
        const lower = String(col || "").trim().toLowerCase();
        const hasAnyVal = rowsMatrix.some(r => !isEmptyVal(r[idx]));
        if (isSpecialProvisionColumnAnyCategory(col) && !catSpecialSet.has(lower)) {
            if (hasAnyVal) keepIndices.push(idx);
            return;
        }
        if (!isSharia && isSyariahSpecColumn(col) && !catSpecialSet.has(lower)) {
            if (hasAnyVal) keepIndices.push(idx);
            return;
        }
        if (!isDplk && (isDplkSpecColumn(col) || isDplkChargeColumn(col) || isDplkBenefitColumn(col)) && !catSpecialSet.has(lower)) {
            if (hasAnyVal) keepIndices.push(idx);
            return;
        }
        if (hasAnyVal || standard.has(lower) || (inSheet && inSheet.has(lower))) {
            keepIndices.push(idx);
        }
    });

    const filteredSchema = keepIndices.map(i => baseMasterSchema[i]);
    const filteredRows = rowsMatrix.map(r => {
        const nr = keepIndices.map(i => r[i]);
        if (r.__sheetName) nr.__sheetName = r.__sheetName;
        return nr;
    });
    return { filteredSchema, filteredRows };
}

async function offloadOversizedCellImagesToDrive(rowsMatrix, onProgress) {
    const imgRegex = /\[IMG:(data:image\/[^\]]+)\]/g;
    for (let r = 0; r < rowsMatrix.length; r++) {
        const row = rowsMatrix[r];
        for (let c = 7; c < row.length; c++) {
            const cellStr = String(row[c] || "");
            if (cellStr.length <= 45000 || !cellStr.includes("[IMG:data:image/")) continue;
            const matches = [...cellStr.matchAll(imgRegex)];
            if (matches.length === 0) continue;
            if (onProgress) onProgress(String(row[0] || ""));
            let updatedStr = cellStr;
            for (const m of matches) {
                const fullToken = m[0];
                const dataUrl = m[1];
                const fileId = await uploadTableImageToDrive(null, dataUrl);
                updatedStr = updatedStr.replace(fullToken, `[IMG_ID:${fileId}]`);
            }
            row[c] = updatedStr;
        }
    }
}

function ensureNonBenefitColumnsInDynamicSchema(includeSyariah = true, includeDplk = false) {
    const allSpecials = [];
    Object.values(CATEGORY_CONFIGS || {}).forEach(cfg => {
        (cfg.specialProvisions || []).forEach(sp => {
            if (sp.name) allSpecials.push(sp.name);
        });
    });
    const baseSpecs = [
        "Additional Benefit",
        ...(includeSyariah ? SYARIAH_SPEC_COLUMNS : []),
        "Entry Age",
        "Coverage Period",
        "Premium Payment Period (PPP)",
        "Min Premium",
        "Min Sum Assured",
        "Underwriting Method",
        "Available Riders",
        ...(includeDplk ? DPLK_SPEC_COLUMNS : []),
        ...allSpecials
    ];
    baseSpecs.forEach(specCol => {
        if (!dynamicColumns.some(c => c.toLowerCase() === specCol.toLowerCase())) {
            const entryAgeIdx = dynamicColumns.findIndex(c => c.toLowerCase() === "entry age");
            if (entryAgeIdx !== -1 && ["Additional Benefit", ...SYARIAH_SPEC_COLUMNS].includes(specCol)) {
                dynamicColumns.splice(entryAgeIdx, 0, specCol);
            } else {
                dynamicColumns.push(specCol);
            }
        }
    });
}

// =====================================================================
// PENGELOLAAN SHEET / KATEGORI (GENERAL PROVISIONS, SPECIAL PROVISIONS, & PRODUCT SECTIONS)
// =====================================================================
let sheetConfigModalState = null;

function openSheetConfigModal(mode = "edit", targetCategory = "") {
    const modalEl = document.getElementById("sheetConfigModal");
    if (!modalEl) return;
    const isNew = mode === "new" || mode === "create";
    const activeCat = isNew ? "" : (targetCategory || SHEET_TABS[0] || "Conventional");
    const cfg = isNew
        ? {
            sheetName: "",
            generalProvisions: DEFAULT_GENERAL_PROVISION_DEFS.map(d => d.col),
            specialProvisions: [],
            productSections: [
                { title: "Product Benefits", mode: "standard", items: [], isChargeLike: false, isPrimaryBenefits: true }
            ]
        }
        : getCategorySchemaConfig(activeCat);

    const sheetAliasObj = (SHEET_ALIAS_RULES || []).find(r => r.canonical.toLowerCase() === String(activeCat).toLowerCase());
    sheetConfigModalState = {
        mode: isNew ? "new" : "edit",
        originalSheetName: isNew ? "" : cfg.sheetName,
        sheetName: cfg.sheetName,
        aliasesText: isNew ? "" : ((sheetAliasObj && Array.isArray(sheetAliasObj.aliases)) ? sheetAliasObj.aliases.join(", ") : ""),
        generalProvisions: [...cfg.generalProvisions],
        extraGeneralPool: getAvailableGeneralProvisions(),
        specialProvisions: cfg.specialProvisions.map(sp => ({
            name: sp.name,
            aliasesText: Array.isArray(sp.aliases) ? sp.aliases.join(", ") : "",
            allowImage: Boolean(sp.allowImage),
            description: sp.description || ""
        })),
        productSections: cfg.productSections.map(sec => ({
            title: sec.title,
            mode: sec.mode === "checklist" ? "checklist" : "standard",
            itemsText: Array.isArray(sec.items) ? sec.items.join("\n") : ""
        }))
    };

    const titleEl = document.getElementById("sheetConfigModalTitle");
    if (titleEl) {
        titleEl.innerHTML = isNew
            ? `<i class="fa-solid fa-folder-plus text-emerald-400"></i> Create New Category / Google Sheet`
            : `<i class="fa-solid fa-sliders text-sky-400"></i> Configure Category: ${escapeHtmlAttr(cfg.sheetName)}`;
    }
    const btnSave = document.getElementById("btnSaveSheetCategoryConfig");
    if (btnSave) {
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Sheet &amp; Config`;
    }

    renderSheetConfigModalBody();
    modalEl.classList.remove("hidden");
}

function closeSheetConfigModal() {
    const modalEl = document.getElementById("sheetConfigModal");
    if (modalEl) modalEl.classList.add("hidden");
    sheetConfigModalState = null;
}

function syncSheetConfigInputsToState() {
    if (!sheetConfigModalState) return;
    const body = document.getElementById("sheetConfigModalBody");
    if (!body) return;

    const nameEl = document.getElementById("cfgSheetNameInput");
    if (nameEl) sheetConfigModalState.sheetName = nameEl.value.trim();

    const aliasEl = document.getElementById("cfgSheetAliasesInput");
    if (aliasEl) sheetConfigModalState.aliasesText = aliasEl.value;

    const checkedGen = [];
    body.querySelectorAll('.cfg-gen-prov-chk:checked').forEach(chk => {
        const v = chk.getAttribute('data-prov-name');
        if (v) checkedGen.push(v);
    });
    sheetConfigModalState.generalProvisions = checkedGen;

    const updatedSpecials = [];
    body.querySelectorAll('.cfg-special-prov-row').forEach(row => {
        const n = (row.querySelector('.cfg-sp-name')?.value || "").trim();
        const al = (row.querySelector('.cfg-sp-aliases')?.value || "").trim();
        const img = Boolean(row.querySelector('.cfg-sp-allow-img')?.checked);
        updatedSpecials.push({ name: n, aliasesText: al, allowImage: img });
    });
    sheetConfigModalState.specialProvisions = updatedSpecials;

    const updatedSections = [];
    body.querySelectorAll('.cfg-product-section-card').forEach(card => {
        const title = (card.querySelector('.cfg-sec-title')?.value || "").trim();
        const mode = card.querySelector('.cfg-sec-mode:checked')?.value === "checklist" ? "checklist" : "standard";
        const itemsText = card.querySelector('.cfg-sec-items')?.value || "";
        updatedSections.push({ title, mode, itemsText });
    });
    sheetConfigModalState.productSections = updatedSections;
}

function renderSheetConfigModalBody() {
    const body = document.getElementById("sheetConfigModalBody");
    if (!body || !sheetConfigModalState) return;
    const st = sheetConfigModalState;
    const genLowerSet = new Set(st.generalProvisions.map(g => g.toLowerCase()));
    const allGenPool = [...new Set([...st.extraGeneralPool, ...st.generalProvisions])];

    const specialRowsHtml = st.specialProvisions.length === 0
        ? `<div class="p-3 rounded-lg border border-dashed border-gray-300 dark:border-neutral-600 text-center text-[11px] text-gray-500 dark:text-neutral-400">No special provisions configured &mdash; Special Provisions section will not appear in forms for this category.</div>`
        : st.specialProvisions.map((sp, idx) => `
                    <div class="cfg-special-prov-row grid grid-cols-1 sm:grid-cols-12 gap-2 items-center p-2 rounded-lg bg-emerald-50/50 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-800/60">
                        <div class="sm:col-span-5">
                            <input type="text" value="${escapeHtmlAttr(sp.name)}" placeholder="Field name (e.g. Hospital Network)..."
                                class="cfg-sp-name w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded px-2 py-1.5 text-xs outline-none focus:ring-1 focus:ring-emerald-500">
                        </div>
                        <div class="sm:col-span-4">
                            <input type="text" value="${escapeHtmlAttr(sp.aliasesText || '')}" placeholder="Aliases (comma-separated, opt)..."
                                class="cfg-sp-aliases w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded px-2 py-1.5 text-xs outline-none focus:ring-1 focus:ring-emerald-500">
                        </div>
                        <label class="sm:col-span-2 flex items-center gap-1 cursor-pointer text-[11px] font-medium text-gray-700 dark:text-neutral-200">
                            <input type="checkbox" class="cfg-sp-allow-img w-3.5 h-3.5 accent-emerald-600 rounded" ${sp.allowImage ? 'checked' : ''}>
                            <span>Allow Img</span>
                        </label>
                        <div class="sm:col-span-1 flex justify-end">
                            <button type="button" data-ev-click="a117" data-ev-click-args="${idcmPackArgs([idx])}" class="text-gray-400 hover:text-red-500 transition p-1" title="Remove Special Provision">
                                <i class="fa-solid fa-trash-can"></i>
                            </button>
                        </div>
                    </div>
                `).join('');

    const sectionCardsHtml = st.productSections.map((sec, idx) => `
                <div class="cfg-product-section-card p-3 rounded-xl border border-blue-200 dark:border-neutral-600 bg-blue-50/40 dark:bg-neutral-800/80 space-y-2.5">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <span class="text-[11px] font-bold uppercase tracking-wider text-blue-800 dark:text-blue-300">
                            Section #${idx + 1}
                        </span>
                        ${st.productSections.length > 1 ? `
                            <button type="button" data-ev-click="a118" data-ev-click-args="${idcmPackArgs([idx])}" class="text-[11px] text-red-500 hover:text-red-400 font-semibold flex items-center gap-1">
                                <i class="fa-solid fa-trash-can"></i> Remove
                            </button>
                        ` : `<span class="text-[10px] text-gray-400">Primary section</span>`}
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Section Title <span class="text-red-500">*</span></label>
                            <input type="text" value="${escapeHtmlAttr(sec.title)}" placeholder="e.g. Product Benefits, Charges, Investment Funds..."
                                class="cfg-sec-title w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg px-2.5 py-1.5 text-xs outline-none focus:ring-1 focus:ring-blue-500">
                        </div>
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Input Style</label>
                            <div class="flex items-center gap-3 pt-1.5">
                                <label class="inline-flex items-center gap-1.5 cursor-pointer font-medium text-gray-800 dark:text-neutral-200">
                                    <input type="radio" name="cfgSecMode_${idx}" value="standard" class="cfg-sec-mode accent-blue-600" ${sec.mode !== 'checklist' ? 'checked' : ''}>
                                    <span>Standard Text &amp; Image</span>
                                </label>
                                <label class="inline-flex items-center gap-1.5 cursor-pointer font-medium text-emerald-700 dark:text-emerald-300">
                                    <input type="radio" name="cfgSecMode_${idx}" value="checklist" class="cfg-sec-mode accent-emerald-600" ${sec.mode === 'checklist' ? 'checked' : ''}>
                                    <span>Checklist + Notes</span>
                                </label>
                            </div>
                        </div>
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Preset / Suggestion Items <span class="font-normal text-gray-500">(1 per line or comma-separated; required for Checklist mode, optional suggestions for Standard mode)</span></label>
                        <textarea rows="3" placeholder="e.g.&#10;Pensiun Normal&#10;Pensiun Ditunda"
                            class="cfg-sec-items w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs outline-none focus:ring-1 focus:ring-blue-500">${escapeHtmlAttr(sec.itemsText || '')}</textarea>
                    </div>
                </div>
            `).join('');

    body.innerHTML = `
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Sheet / Category Name <span class="text-red-500">*</span></label>
                        ${st.mode === "edit" ? `
                            <div class="flex items-center gap-2">
                                <select data-ev-change="a119"
                                    class="w-full border dark:border-neutral-600 bg-gray-50 dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs font-bold outline-none">
                                    ${SHEET_TABS.map(tab => `<option value="${escapeHtmlAttr(tab)}" ${tab.toLowerCase() === st.sheetName.toLowerCase() ? 'selected' : ''}>${escapeHtmlAttr(tab)}</option>`).join('')}
                                </select>
                            </div>
                            <input type="hidden" id="cfgSheetNameInput" value="${escapeHtmlAttr(st.sheetName)}">
                        ` : `
                            <input type="text" id="cfgSheetNameInput" value="${escapeHtmlAttr(st.sheetName)}" placeholder="e.g. Unit Link, Bancassurance, Corporate..."
                                class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs font-semibold outline-none focus:ring-2 focus:ring-emerald-500">
                        `}
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Sheet Aliases <span class="font-normal text-gray-500">(auto-detect Excel tab)</span></label>
                        <input type="text" id="cfgSheetAliasesInput" value="${escapeHtmlAttr(st.aliasesText)}" placeholder="e.g. UnitLink, UL, Endowment..."
                            class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs outline-none focus:ring-2 focus:ring-blue-500">
                    </div>
                </div>

                <!-- 2. General Provisions -->
                <div class="p-3.5 rounded-xl bg-slate-50 dark:bg-neutral-800/60 border border-slate-200 dark:border-neutral-700 space-y-2.5">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <div>
                            <h4 class="font-bold text-slate-800 dark:text-neutral-200 text-xs uppercase tracking-wider">2. General Provisions (Fixed Fields in Forms)</h4>
                            <p class="text-[11px] text-gray-500 dark:text-neutral-400">Select which Non-Benefit Provisions appear by default for this category.</p>
                        </div>
                    </div>
                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-1.5">
                        ${allGenPool.map(prov => `
                            <label class="flex items-center gap-2 p-1.5 rounded-lg bg-white dark:bg-neutral-700/60 border border-gray-200 dark:border-neutral-600 cursor-pointer hover:border-blue-400 transition">
                                <input type="checkbox" class="cfg-gen-prov-chk w-4 h-4 accent-blue-600 rounded shrink-0" data-prov-name="${escapeHtmlAttr(prov)}" ${genLowerSet.has(prov.toLowerCase()) ? 'checked' : ''}>
                                <span class="text-xs font-medium text-gray-800 dark:text-neutral-100 truncate">${escapeHtmlAttr(prov)}</span>
                            </label>
                        `).join('')}
                    </div>
                    <div class="flex items-center gap-2 pt-1">
                        <input type="text" id="cfgQuickNewGenProv" placeholder="Add new general provision name (e.g. Grace Period)..."
                            class="flex-1 border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg px-2.5 py-1.5 text-xs outline-none focus:ring-1 focus:ring-blue-500">
                        <button type="button" data-ev-click="a120"
                            class="px-3 py-1.5 bg-slate-700 hover:bg-slate-600 text-white rounded-lg text-xs font-bold transition whitespace-nowrap">
                            <i class="fa-solid fa-plus me-1"></i>Add Provision
                        </button>
                    </div>
                </div>

                <!-- 3. Special Provisions -->
                <div class="p-3.5 rounded-xl bg-emerald-50/40 dark:bg-emerald-950/20 border border-emerald-200 dark:border-emerald-800/70 space-y-2.5">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <div>
                            <h4 class="font-bold text-emerald-900 dark:text-emerald-300 text-xs uppercase tracking-wider">3. Special Provisions (Category-Specific)</h4>
                            <p class="text-[11px] text-gray-500 dark:text-neutral-400">E.g. <em>Aqd, Alokasi Kontribusi</em> in Sharia, <em>Sales Tools</em> in DPLK, or <em>Hospital Network</em> in H&amp;S. Hidden automatically if empty.</p>
                        </div>
                        <button type="button" data-ev-click="a121"
                            class="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm">
                            <i class="fa-solid fa-plus"></i> Add Special Provision
                        </button>
                    </div>
                    <div class="space-y-2">${specialRowsHtml}</div>
                </div>

                <!-- 4. Dynamic Product Sections -->
                <div class="space-y-2.5">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <div>
                            <h4 class="font-bold text-blue-900 dark:text-blue-300 text-xs uppercase tracking-wider">4. Dynamic Product Sections (Benefits, Charges, &amp; Custom)</h4>
                            <p class="text-[11px] text-gray-500 dark:text-neutral-400">Choose whether each repeatable feature section uses Standard Text/Image or Checklist + Notes (like DPLK).</p>
                        </div>
                        <button type="button" data-ev-click="a122"
                            class="px-2.5 py-1.5 bg-blue-600 hover:bg-blue-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm">
                            <i class="fa-solid fa-plus"></i> Add Product Section
                        </button>
                    </div>
                    <div class="space-y-2.5">${sectionCardsHtml}</div>
                </div>
            `;
}

function addQuickGeneralProvisionInModal() {
    syncSheetConfigInputsToState();
    const inp = document.getElementById("cfgQuickNewGenProv");
    const val = (inp ? inp.value : "").trim();
    if (!val) {
        if (inp) inp.focus();
        return;
    }
    if (!sheetConfigModalState.extraGeneralPool.some(x => x.toLowerCase() === val.toLowerCase())) {
        sheetConfigModalState.extraGeneralPool.push(val);
    }
    if (!sheetConfigModalState.generalProvisions.some(x => x.toLowerCase() === val.toLowerCase())) {
        sheetConfigModalState.generalProvisions.push(val);
    }
    renderSheetConfigModalBody();
}

function addSheetConfigSpecialRow() {
    syncSheetConfigInputsToState();
    sheetConfigModalState.specialProvisions.push({ name: "", aliasesText: "", allowImage: false });
    renderSheetConfigModalBody();
    const names = document.querySelectorAll('#sheetConfigModalBody .cfg-sp-name');
    if (names.length > 0) names[names.length - 1].focus();
}

function removeSheetConfigSpecialRow(idx) {
    syncSheetConfigInputsToState();
    sheetConfigModalState.specialProvisions.splice(idx, 1);
    renderSheetConfigModalBody();
}

function addSheetConfigSectionBlock() {
    syncSheetConfigInputsToState();
    sheetConfigModalState.productSections.push({ title: "", mode: "standard", itemsText: "" });
    renderSheetConfigModalBody();
    const titles = document.querySelectorAll('#sheetConfigModalBody .cfg-sec-title');
    if (titles.length > 0) titles[titles.length - 1].focus();
}

function removeSheetConfigSectionBlock(idx) {
    syncSheetConfigInputsToState();
    if (sheetConfigModalState.productSections.length <= 1) return;
    sheetConfigModalState.productSections.splice(idx, 1);
    renderSheetConfigModalBody();
}

async function saveSheetCategoryConfig() {
    if (!sheetConfigModalState) return;
    syncSheetConfigInputsToState();

    const sheetName = sheetConfigModalState.sheetName.trim();
    if (!sheetName) {
        showToast("Please enter a Sheet / Category Name.");
        document.getElementById("cfgSheetNameInput")?.focus();
        return;
    }
    if (/^(config|users|documents)$/i.test(sheetName)) {
        showToast(`"${sheetName}" is a reserved system sheet name. Please choose another name.`);
        return;
    }
    if (sheetConfigModalState.mode === "new" && SHEET_TABS.some(t => t.toLowerCase() === sheetName.toLowerCase())) {
        showToast(`Sheet "${sheetName}" already exists. Use "Edit Category" to modify its provisions or sections.`);
        return;
    }

    const cleanedSpecials = sheetConfigModalState.specialProvisions
        .filter(sp => sp.name.trim() !== "")
        .map(sp => ({
            name: sp.name.trim(),
            aliases: sp.aliasesText.split(/[,;\n]/).map(s => s.trim()).filter(Boolean),
            allowImage: Boolean(sp.allowImage)
        }));

    const cleanedSections = [];
    for (const sec of sheetConfigModalState.productSections) {
        const title = sec.title.trim();
        if (!title) {
            showToast("Please fill in Section Title for every Product Section.");
            return;
        }
        const items = sec.itemsText.split(/[\r\n,;]+/).map(s => s.trim()).filter(Boolean);
        if (sec.mode === "checklist" && items.length === 0) {
            showToast(`Section "${title}" uses Checklist mode. Please provide at least 1 preset item.`);
            return;
        }
        cleanedSections.push({ title, mode: sec.mode, items });
    }

    const aliases = sheetConfigModalState.aliasesText.split(/[,;\n]/).map(s => s.trim()).filter(Boolean);
    const btnSave = document.getElementById("btnSaveSheetCategoryConfig");
    if (btnSave) {
        btnSave.disabled = true;
        btnSave.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving Sheet &amp; Config...`;
    }

    try {
        const payload = {
            action: "save_sheet_category_config",
            isNewSheet: sheetConfigModalState.mode === "new",
            sheetName,
            aliases,
            generalProvisions: sheetConfigModalState.generalProvisions,
            specialProvisions: cleanedSpecials,
            productSections: cleanedSections
        };

        const resp = await fetch(APPS_SCRIPT_URL, {
            method: "POST",
            headers: { "Content-Type": "text/plain;charset=utf-8" },
            body: JSON.stringify(withSessionToken(payload))
        });
        const res = await resp.json();
        if (handleUnauthorizedResponse(res)) {
            throw new Error(res.message || "Session expired.");
        }
        if (!res || res.status !== "success") {
            throw new Error((res && res.message) || "Failed to save category configuration.");
        }

        // Update local state immediately
        if (res.config) {
            resetDynamicConfigListsToDefault();
            applyDynamicConfigFromBackend(res.config);
        }
        if (!CATEGORY_CONFIGS[sheetName]) {
            CATEGORY_CONFIGS[sheetName] = {
                sheetName,
                generalProvisions: [...sheetConfigModalState.generalProvisions],
                specialProvisions: cleanedSpecials,
                productSections: cleanedSections
            };
        }
        if (!SHEET_TABS.some(t => t.toLowerCase() === sheetName.toLowerCase())) {
            SHEET_TABS.push(sheetName);
        }

        invalidateLocalSummaryCache();
        sheetHasHeaders[sheetName] = true;
        syncSheetTabSelectors();

        // Sinkronkan dropdown Upload Excel / Manual Add / Edit Product ke sheet yang baru disimpan
        const uploadSel = document.getElementById("uploadTargetSheet");
        if (uploadSel) uploadSel.value = sheetName;
        const manualCatSel = document.getElementById("manual_category");
        if (manualCatSel) manualCatSel.value = sheetName;
        const editCatSel = document.getElementById("edit_category");
        if (editCatSel && isEditMode) editCatSel.value = sheetName;

        closeSheetConfigModal();

        const manualDyn = document.getElementById("manualDynamicCategorySections");
        if (manualDyn) delete manualDyn.dataset.renderedCat;
        const editDyn = document.getElementById("editDynamicCategorySections");
        if (editDyn) delete editDyn.dataset.renderedCat;

        if (document.getElementById("manualFormBody") && !document.getElementById("uploadModal")?.classList.contains("hidden")) {
            validateManualCategoryAndToggleSyariah();
        }
        if (isEditMode && currentActiveProduct) {
            validateEditProductIdentity();
        }

        filterData();
    } catch (err) {
        console.error("saveSheetCategoryConfig error:", err);
        showToast("Failed to save Sheet / Category configuration: " + (err.message || err));
    } finally {
        if (btnSave) {
            btnSave.disabled = false;
            btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Sheet &amp; Config`;
        }
    }
}

// =====================================================================
// RENDERER DINAMIS PROVISION, SPECIAL PROVISION, & PRODUCT SECTIONS (EDIT & MANUAL)
// =====================================================================
function captureScopeFormValues(scopeRoot) {
    const values = {};
    if (!scopeRoot) return values;
    scopeRoot.querySelectorAll('.spec-field-input[data-spec-col]').forEach(el => {
        const col = el.getAttribute('data-spec-col');
        if (!col) return;
        const wrapper = el.closest('.field-with-images-wrapper');
        const imgs = [];
        if (wrapper) {
            wrapper.querySelectorAll('.attached-img-item img').forEach(img => {
                if (img.dataset.origSrc || img.src) imgs.push(img.dataset.origSrc || img.src);
            });
        }
        values[col.toLowerCase()] = { text: el.value, images: imgs };
    });
    return values;
}

function addNewSectionItemRow(containerId, noticeId, sectionTitle, datalistId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const notice = noticeId ? document.getElementById(noticeId) : null;
    if (notice) notice.classList.add("hidden");

    const scopeRoot = container.closest('#manualFormBody, #detailBody') || document;
    const rowDiv = document.createElement("div");
    rowDiv.className = "new-section-row field-with-images-wrapper p-3 bg-indigo-50/70 dark:bg-neutral-800/80 border border-indigo-200 dark:border-neutral-600 rounded-xl space-y-2 relative transition";
    rowDiv.setAttribute("data-section-title", sectionTitle || "Feature");

    rowDiv.innerHTML = `
                <div class="flex justify-between items-center">
                    <span class="text-[11px] font-bold uppercase tracking-wider text-indigo-700 dark:text-indigo-300">
                        <i class="fa-solid fa-layer-group me-1"></i> Add ${escapeHtmlAttr(sectionTitle || "Item")}
                    </span>
                    <div class="flex items-center gap-1.5">
                        <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-700 dark:text-indigo-300 bg-indigo-100 hover:bg-indigo-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Paste screenshot / image">
                            <i class="fa-solid fa-paste"></i> Paste
                        </button>
                        <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-700 dark:text-indigo-300 bg-indigo-100 hover:bg-indigo-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Upload image">
                            <i class="fa-solid fa-upload"></i> Upload
                            <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                        </label>
                        <button type="button" data-ev-click="a125" data-ev-click-args="${idcmPackArgs([noticeId || ''])}" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded" title="Cancel adding">
                            <i class="fa-solid fa-xmark"></i> Cancel
                        </button>
                    </div>
                </div>
                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Item Name</label>
                    <input type="text" list="${datalistId || ''}" data-ev-input="a126"
                           placeholder="Type or select item name..."
                           class="new-section-item-name w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-indigo-500 outline-none">
                    <div class="section-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                </div>
                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Description / Details</label>
                    <textarea rows="2" data-ev-paste="a127" placeholder="Enter details or paste (Ctrl+V) table image..."
                              class="new-section-item-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-indigo-500 outline-none"></textarea>
                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden"></div>
                </div>
            `;

    container.appendChild(rowDiv);
    rowDiv.querySelector('.new-section-item-name')?.focus();
    refreshSectionItemDatalist(scopeRoot);
}

function getUsedSectionItemNamesLower(excludeInputEl = null, scopeContainer = null) {
    const root = scopeContainer || (excludeInputEl && excludeInputEl.closest('#manualFormBody, #detailBody')) || document;
    const names = [];
    root.querySelectorAll('.existing-section-row, .existing-benefit-row, .existing-charge-row').forEach(r => {
        const c = r.getAttribute('data-col-name');
        if (c) names.push(c.trim().toLowerCase());
    });
    root.querySelectorAll('.new-section-item-name, .new-benefit-name, .new-charge-name').forEach(inp => {
        if (inp !== excludeInputEl && inp.value.trim()) names.push(inp.value.trim().toLowerCase());
    });
    return names;
}

function refreshSectionItemDatalist(scopeContainer = null) {
    const root = scopeContainer || document;
    const used = new Set(getUsedSectionItemNamesLower(null, root));
    root.querySelectorAll('datalist.dynamic-section-datalist').forEach(dl => {
        const allItems = String(dl.getAttribute('data-preset-items') || "").split("||").filter(Boolean);
        dl.innerHTML = allItems
            .filter(it => !used.has(it.toLowerCase()))
            .map(it => `<option value="${escapeHtmlAttr(it)}"></option>`)
            .join('');
    });
}

function validateSectionItemName(inputEl) {
    const rawVal = inputEl.value.trim();
    const lowerVal = rawVal.toLowerCase();
    const feedbackEl = inputEl.parentElement.querySelector('.section-name-feedback');
    const scopeRoot = inputEl.closest('#manualFormBody, #detailBody') || document;

    if (!rawVal) {
        if (feedbackEl) { feedbackEl.className = "section-name-feedback hidden"; feedbackEl.replaceChildren(); }
        inputEl.dataset.status = "empty";
        refreshSectionItemDatalist(scopeRoot);
        return;
    }
    if (isNonBenefitColumn(rawVal)) {
        feedbackEl.className = "section-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtmlAttr(rawVal)}"</strong> is already a Non-Benefit Provision column.`;
        inputEl.dataset.status = "duplicate";
        return;
    }
    if (getUsedSectionItemNamesLower(inputEl, scopeRoot).includes(lowerVal)) {
        feedbackEl.className = "section-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtmlAttr(rawVal)}"</strong> is already added in this product!`;
        inputEl.dataset.status = "duplicate";
        return;
    }
    feedbackEl.className = "section-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
    feedbackEl.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> Ready to save under this section.`;
    inputEl.dataset.status = "valid";
    refreshSectionItemDatalist(scopeRoot);
}

function removeSectionItemRow(btnEl, noticeId) {
    const scopeRoot = btnEl.closest('#manualFormBody, #detailBody') || document;
    const row = btnEl.closest('.existing-section-row, .new-section-row');
    const container = row ? row.parentElement : null;
    if (row) row.remove();
    if (container && noticeId) {
        const notice = document.getElementById(noticeId);
        if (notice && container.querySelectorAll('.existing-section-row, .new-section-row').length === 0) {
            notice.classList.remove("hidden");
        }
    }
    scopeRoot.querySelectorAll('.new-section-item-name').forEach(inp => validateSectionItemName(inp));
    refreshSectionItemDatalist(scopeRoot);
}

// Validasi terpusat untuk seluruh baris baru (Provision, Benefit, Charge, & Custom Section):
// Jika kolom Name terisi namun Description & Gambar sama-sama kosong, munculkan alert dan batalkan simpan.
function validateAllNewRepeatableRows(rootEl) {
    if (!rootEl) return true;
    const groups = [
        {
            selector: '.new-nonbenefit-row',
            nameClass: '.new-nonbenefit-name',
            descClass: '.new-nonbenefit-desc',
            kindLabel: 'Provision',
            nameFieldLabel: 'Provision Name',
            validateFn: typeof validateNonBenefitName === 'function' ? validateNonBenefitName : null
        },
        {
            selector: '.new-benefit-row',
            nameClass: '.new-benefit-name',
            descClass: '.new-benefit-desc',
            kindLabel: 'Benefit',
            nameFieldLabel: 'Benefit Name',
            validateFn: typeof validateBenefitName === 'function' ? validateBenefitName : null
        },
        {
            selector: '.new-charge-row',
            nameClass: '.new-charge-name',
            descClass: '.new-charge-desc',
            kindLabel: 'Charge component',
            nameFieldLabel: 'Charge Name',
            validateFn: typeof validateChargeName === 'function' ? validateChargeName : null
        },
        {
            selector: '.new-section-row',
            nameClass: '.new-section-item-name',
            descClass: '.new-section-item-desc',
            kindLabel: 'Feature item',
            nameFieldLabel: 'Item Name',
            validateFn: typeof validateSectionItemName === 'function' ? validateSectionItemName : null
        }
    ];

    for (const g of groups) {
        const rows = Array.from(rootEl.querySelectorAll(g.selector));
        for (const row of rows) {
            const nameInput = row.querySelector(g.nameClass);
            const descInput = row.querySelector(g.descClass);
            if (!nameInput || !descInput) continue;
            const hasAttachedImgs = row.querySelectorAll('.attached-img-item').length > 0;
            const nameVal = nameInput.value.trim();
            const descVal = descInput.value.trim();

            if (!nameVal && !descVal && !hasAttachedImgs) continue;
            if (!nameVal && (descVal || hasAttachedImgs)) {
                showToast(`Please fill in '${g.nameFieldLabel}' for every description or image added.`);
                nameInput.focus();
                return false;
            }
            if (nameVal && !descVal && !hasAttachedImgs) {
                showToast(`Please fill in description or attach an image for ${g.kindLabel} "${nameVal}" before saving.`);
                descInput.focus();
                return false;
            }
            if (g.validateFn) {
                g.validateFn(nameInput);
                if (nameInput.dataset.status === "duplicate") {
                    showToast(`${g.kindLabel} "${nameVal}" is invalid or already exists in this form. Please check again.`);
                    nameInput.focus();
                    return false;
                }
            }
        }
    }

    // Validasi item custom checklist yang dicentang agar namanya tidak kosong
    const customChkRows = Array.from(rootEl.querySelectorAll('.dplk-custom-benefit-item'));
    for (const row of customChkRows) {
        const chk = row.querySelector('.dplk-benefit-chk');
        const customNameEl = row.querySelector('.dplk-custom-benefit-name');
        if (chk && chk.checked && customNameEl && !customNameEl.value.trim()) {
            showToast("Please fill in the item name for every checked custom checklist item.");
            customNameEl.focus();
            return false;
        }
    }

    return true;
}

function renderCategoryDynamicFormSections(scopePrefix, categoryName, productObj = null, preservedValues = null) {
    const isEdit = scopePrefix === "edit";
    const catCfg = getCategorySchemaConfig(categoryName);
    const isDplkCat = /^dplk$/i.test(categoryName || "");
    const focusRing = isEdit ? "focus:ring-blue-500" : "focus:ring-green-500";

    const getFieldState = (colName) => {
        const lowerCol = colName.toLowerCase();
        if (preservedValues && preservedValues[lowerCol]) {
            return preservedValues[lowerCol];
        }
        if (productObj && productObj.attributes) {
            const actualKey = Object.keys(productObj.attributes).find(k => k.toLowerCase() === lowerCol) || colName;
            const rawVal = productObj.attributes[actualKey];
            if (rawVal && rawVal !== "-") return splitTextAndImages(rawVal);
        }
        return { text: "", images: [] };
    };

    // 1. General Provisions Grid (sesuai daftar yang dicentang di Category Config + field existing yang sudah terisi)
    const activeGenNames = [...catCfg.generalProvisions];
    if (productObj && productObj.attributes) {
        DEFAULT_GENERAL_PROVISION_DEFS.forEach(def => {
            if (!activeGenNames.some(g => g.toLowerCase() === def.col.toLowerCase())) {
                const st = getFieldState(def.col);
                if (!isEmptyVal(st.text) || st.images.length > 0) {
                    activeGenNames.push(def.col);
                }
            }
        });
    }

    const genDefMap = {};
    DEFAULT_GENERAL_PROVISION_DEFS.forEach(d => { genDefMap[d.col.toLowerCase()] = d; });
    const gridGenFields = [];
    let fullWidthGenHtml = "";
    activeGenNames.forEach(col => {
        const def = genDefMap[col.toLowerCase()] || { col, label: col, placeholder: `Enter ${col} details...`, fullWidth: false };
        const st = getFieldState(col);
        const lblText = (isDplkCat && def.dplkLabel) ? def.dplkLabel : def.label;
        const fieldHtml = `
                    <div class="field-with-images-wrapper">
                        <div class="flex justify-between items-center mb-1">
                            <label class="font-bold text-gray-700 dark:text-neutral-300">${escapeHtmlAttr(lblText)}</label>
                            <button type="button" data-ev-click="a128" class="text-[10px] text-gray-400 hover:text-red-500 transition" title="Clear this field">Clear</button>
                        </div>
                        <textarea data-spec-col="${escapeHtmlAttr(col)}" rows="2" class="spec-field-input w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs ${focusRing} outline-none" placeholder="${escapeHtmlAttr(def.placeholder)}">${escapeHtmlAttr(st.text).replace(/&quot;/g, '"')}</textarea>
                    </div>
                `;
        if (def.fullWidth) fullWidthGenHtml += fieldHtml;
        else gridGenFields.push(fieldHtml);
    });

    // Custom Non-Benefits (di luar General Provisions aktif dan Special Provisions)
    const activeGenLowerSet = new Set(activeGenNames.map(c => c.toLowerCase()));
    const existingCustomNbCols = (isEdit && productObj)
        ? getCustomNonBenefitColumns().filter(col => {
            if (activeGenLowerSet.has(col.toLowerCase())) return false;
            const st = getFieldState(col);
            return !isEmptyVal(st.text) || st.images.length > 0;
        })
        : [];
    const customNbContainerId = isEdit ? "editNewNonBenefitsContainer" : "manualNewNonBenefitsContainer";
    const existingNbListId = isEdit ? "existingNonBenefitNamesList" : "manualNonBenefitNamesList";

    // 2. Special Provisions Box — hanya dirender jika categoryConfig punya >= 1 Special Provision!
    const specialsToShow = catCfg.specialProvisions.map(sp => {
        let st = getFieldState(sp.name);
        if (!st.text && st.images.length === 0 && sp.name.toLowerCase() === "aqd") {
            const alt = getFieldState("Akad");
            if (alt.text || alt.images.length > 0) st = alt;
        }
        return { ...sp, st };
    });

    const specialProvisionsHtml = specialsToShow.length === 0 ? "" : `
                <div class="p-3.5 rounded-xl bg-emerald-50/70 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-800/80 space-y-3 transition-all">
                    <div class="flex items-center justify-between">
                        <h5 class="font-bold text-emerald-900 dark:text-emerald-300 text-xs uppercase tracking-wider flex items-center gap-1.5">
                            <i class="fa-solid fa-star-of-life text-emerald-600 dark:text-emerald-400"></i> ${escapeHtmlAttr(catCfg.sheetName)} Special Provisions
                        </h5>
                        <span class="text-[10px] font-semibold bg-emerald-100 dark:bg-emerald-900/60 text-emerald-800 dark:text-emerald-200 px-2 py-0.5 rounded-full">${escapeHtmlAttr(catCfg.sheetName)} Only</span>
                    </div>
                    <div class="space-y-2.5">
                        ${specialsToShow.map(sp => {
        const showImgs = sp.allowImage || sp.st.images.length > 0;
        const lbl = sp.name.toLowerCase() === "aqd" ? "Aqd / Contract"
            : (sp.name.toLowerCase() === "alokasi kontribusi" ? "Contribution Allocation (Alokasi Kontribusi)"
                : (sp.name.toLowerCase() === "surplus underwriting" ? "Underwriting Surplus (Surplus Underwriting)" : sp.name));
        return `
                                <div class="field-with-images-wrapper">
                                    <div class="flex flex-wrap justify-between items-center gap-1.5 mb-1">
                                        <label class="block font-bold text-gray-700 dark:text-neutral-300">${escapeHtmlAttr(lbl)}</label>
                                        ${showImgs ? `
                                            <div class="flex items-center gap-1.5">
                                                <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 dark:text-emerald-300 bg-emerald-100 hover:bg-emerald-200 dark:bg-emerald-900/50 dark:hover:bg-emerald-800/60 px-2 py-0.5 rounded-md transition" title="Paste image (Ctrl+V)">
                                                    <i class="fa-solid fa-paste"></i> Paste Image
                                                </button>
                                                <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 dark:text-emerald-300 bg-emerald-100 hover:bg-emerald-200 dark:bg-emerald-900/50 dark:hover:bg-emerald-800/60 px-2 py-0.5 rounded-md transition" title="Upload image">
                                                    <i class="fa-solid fa-upload"></i> Upload
                                                    <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                                                </label>
                                            </div>
                                        ` : ''}
                                    </div>
                                    <textarea data-spec-col="${escapeHtmlAttr(sp.name)}" rows="${showImgs ? 3 : 2}" ${showImgs ? 'onpaste="handleFieldImagePaste(event, this)"' : ''} class="spec-field-input w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-emerald-500 outline-none" placeholder="Enter ${escapeHtmlAttr(sp.name)} details${showImgs ? ' or paste image...' : '...'}">${escapeHtmlAttr(sp.st.text).replace(/&quot;/g, '"')}</textarea>
                                    ${showImgs ? `<div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(sp.st.images)}</div>` : ''}
                                </div>
                            `;
    }).join('')}
                    </div>
                </div>
            `;

    // 3. Dynamic Product Sections (Checklist + Notes ATAU Standard Text & Image)
    const assignedNonPrimaryCols = getSectionAssignedNonPrimaryColumnsForCategory(categoryName);
    const regularBenefitCols = getRegularBenefitColumns().filter(col => {
        if (assignedNonPrimaryCols.has(col.toLowerCase())) return false;
        if (!productObj) return true;
        const v = productObj.attributes[col];
        if (isEmptyVal(v)) return false;
        if (isDplkBenefitColumn(col) && parseDplkBenefitValue(v).isChecked) return false;
        return true;
    });

    let hasRenderedChecklistId = false;
    const sectionsHtml = catCfg.productSections.map((sec, secIdx) => {
        if (sec.mode === "checklist") {
            const checklistId = !hasRenderedChecklistId
                ? (isEdit ? "editDplkBenefitChecklist" : "manualDplkBenefitChecklist")
                : `${scopePrefix}ChecklistSec_${secIdx}`;
            const badgeId = !hasRenderedChecklistId
                ? (isEdit ? "editDplkChecklistCount" : "manualDplkChecklistCount")
                : `${scopePrefix}ChecklistCount_${secIdx}`;
            hasRenderedChecklistId = true;
            const presetItems = sec.items || [];
            const presetLowerSet = new Set(presetItems.map(it => it.toLowerCase()));

            const checkedCount = presetItems.filter(bCol => {
                if (!productObj || !productObj.attributes) return false;
                const actualKey = Object.keys(productObj.attributes).find(k => k.toLowerCase() === bCol.toLowerCase()) || bCol;
                return parseDplkBenefitValue(productObj.attributes[actualKey]).isChecked;
            }).length;

            const extraCustomChecklistCols = (isEdit && productObj)
                ? getRegularBenefitColumns().filter(col => {
                    if (presetLowerSet.has(col.toLowerCase()) || assignedNonPrimaryCols.has(col.toLowerCase())) return false;
                    const v = productObj.attributes[col];
                    return !isEmptyVal(v);
                })
                : [];

            return `
                        <hr class="my-3 dark:border-neutral-700">
                        <div class="flex justify-between items-center mb-2">
                            <h5 class="font-bold text-blue-900 dark:text-blue-300 text-sm">${escapeHtmlAttr(sec.title)}</h5>
                        </div>
                        <div id="${checklistId}" class="p-3.5 rounded-xl bg-emerald-50/60 dark:bg-emerald-950/30 border border-emerald-200 dark:border-emerald-800/80 space-y-2.5 mb-2.5">
                            <div class="flex flex-wrap items-center justify-between gap-2">
                                <div class="flex items-center gap-1.5">
                                    <i class="fa-solid fa-list-check text-emerald-600 dark:text-emerald-400"></i>
                                    <span class="font-bold text-emerald-900 dark:text-emerald-300 text-xs">${escapeHtmlAttr(sec.title)} (Checklist)</span>
                                    <span id="${badgeId}" class="text-[10px] font-bold bg-emerald-100 dark:bg-emerald-900/70 text-emerald-800 dark:text-emerald-200 px-2 py-0.5 rounded-full">${checkedCount} / ${presetItems.length} selected</span>
                                </div>
                                <div class="flex items-center gap-1.5">
                                    <button type="button" data-ev-click="a129" data-ev-click-args="${idcmPackArgs([checklistId])}" class="text-[10px] font-semibold px-2 py-0.5 rounded bg-emerald-600 hover:bg-emerald-500 text-white transition">Select All</button>
                                    <button type="button" data-ev-click="a130" data-ev-click-args="${idcmPackArgs([checklistId])}" class="text-[10px] font-semibold px-2 py-0.5 rounded bg-gray-200 hover:bg-gray-300 dark:bg-neutral-700 dark:hover:bg-neutral-600 text-gray-700 dark:text-neutral-200 transition">Reset</button>
                                </div>
                            </div>
                            <p class="text-[11px] text-gray-600 dark:text-neutral-400">Check available items and optionally add a note beside each checked item:</p>
                            <div class="space-y-1.5">
                                ${presetItems.map(bCol => {
                let isChecked = false;
                let rem = "";
                if (productObj && productObj.attributes) {
                    const actualKey = Object.keys(productObj.attributes).find(k => k.toLowerCase() === bCol.toLowerCase()) || bCol;
                    const parsed = parseDplkBenefitValue(productObj.attributes[actualKey]);
                    isChecked = parsed.isChecked;
                    rem = parsed.remark || "";
                }
                return `
                                        <div class="dplk-benefit-item flex flex-col sm:flex-row sm:items-center gap-2 p-2 rounded-lg border border-emerald-200/80 dark:border-emerald-800/60 bg-white dark:bg-neutral-800 hover:bg-emerald-50/40 dark:hover:bg-emerald-900/20 transition">
                                            <label class="flex items-center gap-2 cursor-pointer sm:w-5/12 shrink-0">
                                                <input type="checkbox" class="dplk-benefit-chk w-4 h-4 rounded accent-emerald-600 cursor-pointer shrink-0" data-benefit-col="${escapeHtmlAttr(bCol)}" ${isChecked ? 'checked' : ''} data-ev-change="a131" data-ev-change-args="${idcmPackArgs([checklistId])}">
                                                <span class="text-xs font-medium text-gray-800 dark:text-neutral-200">${escapeHtmlAttr(bCol)}</span>
                                            </label>
                                            <input type="text" ${isChecked ? '' : 'disabled'} class="dplk-benefit-remark flex-1 w-full rounded px-2.5 py-1.5 text-xs outline-none transition border ${isChecked ? 'bg-white dark:bg-neutral-700 border-emerald-400 dark:border-emerald-500 text-gray-900 dark:text-neutral-100 focus:ring-2 focus:ring-emerald-500' : 'bg-gray-100 dark:bg-neutral-900/60 border-gray-200 dark:border-neutral-700 text-gray-400 dark:text-neutral-500 opacity-60 cursor-not-allowed'}" data-benefit-col="${escapeHtmlAttr(bCol)}" value="${escapeHtmlAttr(rem)}" placeholder="Add note (optional)...">
                                        </div>
                                    `;
            }).join('')}
                                <div class="dplk-custom-benefits-list space-y-1.5 empty:hidden">
                                    ${extraCustomChecklistCols.map(col => {
                const rawCustom = productObj.attributes[col] || "";
                const parsedCustom = parseDplkBenefitValue(rawCustom);
                const customRemark = parsedCustom.isChecked ? parsedCustom.remark : splitTextAndImages(rawCustom).text;
                return `
                                            <div class="dplk-benefit-item dplk-custom-benefit-item flex flex-col sm:flex-row sm:items-center gap-2 p-2 rounded-lg border border-emerald-300 dark:border-emerald-700 bg-white dark:bg-neutral-800 transition">
                                                <div class="flex items-center gap-2 sm:w-5/12 shrink-0">
                                                    <input type="checkbox" class="dplk-benefit-chk w-4 h-4 rounded accent-emerald-600 cursor-pointer shrink-0" checked data-ev-change="a131" data-ev-change-args="${idcmPackArgs([checklistId])}">
                                                    <input type="text" class="dplk-custom-benefit-name flex-1 min-w-0 border border-gray-300 dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded px-2 py-1 text-xs font-medium focus:ring-1 focus:ring-emerald-500 outline-none" value="${escapeHtmlAttr(col)}" placeholder="Item name..." data-ev-input="a131" data-ev-input-args="${idcmPackArgs([checklistId])}">
                                                    <button type="button" data-ev-click="a132" data-ev-click-args="${idcmPackArgs([checklistId])}" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition px-1" title="Remove item">
                                                        <i class="fa-solid fa-trash-can text-xs"></i>
                                                    </button>
                                                </div>
                                                <input type="text" class="dplk-benefit-remark flex-1 w-full rounded px-2.5 py-1.5 text-xs outline-none transition border bg-white dark:bg-neutral-700 border-emerald-400 dark:border-emerald-500 text-gray-900 dark:text-neutral-100 focus:ring-2 focus:ring-emerald-500" value="${escapeHtmlAttr(customRemark)}" placeholder="Add note (optional)...">
                                            </div>
                                        `;
            }).join('')}
                                </div>
                            </div>
                            <button type="button" data-ev-click="a133" data-ev-click-args="${idcmPackArgs([checklistId])}" class="w-full py-1.5 border border-dashed border-emerald-400 dark:border-emerald-700 hover:border-emerald-500 text-emerald-700 dark:text-emerald-300 hover:bg-emerald-100/50 dark:hover:bg-emerald-900/30 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5">
                                <i class="fa-solid fa-circle-plus"></i> Add Item
                            </button>
                        </div>
                    `;
        }

        // Mode Standard (Charge-like, Primary Benefits, atau Custom Section)
        if (sec.isChargeLike) {
            const chargeContainerId = isEdit ? "editNewChargesContainer" : "manualNewChargesContainer";
            const chargeNoticeId = isEdit ? "editEmptyChargeNotice" : "manualEmptyChargeNotice";
            const chargeDatalistId = isEdit ? "existingChargeNamesList" : "manualChargeNamesList";
            const chargeSuggestions = [...new Set([...(sec.items || []), ...DPLK_CHARGE_COLUMNS, ...getChargeColumns()])];
            const filledCharges = (isEdit && productObj)
                ? [...new Set([...(sec.items || []), ...getChargeColumns()])].filter(col => !isEmptyVal(productObj.attributes[col]))
                : [];

            return `
                        <div class="pt-2 space-y-2.5">
                            <hr class="my-3 dark:border-neutral-700">
                            <div class="flex justify-between items-center mb-1">
                                <h5 class="font-bold text-amber-800 dark:text-amber-300 text-sm flex items-center gap-1.5">
                                    <i class="fa-solid fa-receipt text-amber-600 dark:text-amber-400"></i> ${escapeHtmlAttr(sec.title)}
                                </h5>
                                <button type="button" data-ev-click="a134" data-ev-click-args="${idcmPackArgs([chargeContainerId, chargeNoticeId])}" class="px-2.5 py-1.5 bg-amber-600 hover:bg-amber-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm shrink-0">
                                    <i class="fa-solid fa-plus"></i> Add Charge
                                </button>
                            </div>
                            <datalist id="${chargeDatalistId}">${chargeSuggestions.map(c => `<option value="${escapeHtmlAttr(c)}"></option>`).join('')}</datalist>
                            <div ${isEdit ? 'id="existingChargesContainer"' : ''} class="space-y-2.5">
                                ${filledCharges.length === 0 ? `
                                    <div id="${chargeNoticeId}" class="text-center py-4 bg-amber-50/40 dark:bg-amber-950/20 rounded-lg border border-dashed border-amber-300 dark:border-amber-800 text-gray-500 dark:text-neutral-400">
                                        No charge details added yet. Click <strong>+ Add Charge</strong> to fill in fee components.
                                    </div>
                                ` : filledCharges.map(col => {
                const cSplit = splitTextAndImages(productObj.attributes[col] || "");
                return `
                                        <div class="existing-charge-row field-with-images-wrapper p-3 bg-amber-50/40 dark:bg-amber-950/20 border border-amber-200 dark:border-amber-800/70 rounded-lg" data-col-name="${escapeHtmlAttr(col)}">
                                            <div class="flex flex-wrap justify-between items-center mb-1.5 gap-1.5">
                                                <label class="font-bold text-amber-900 dark:text-amber-200">${escapeHtmlAttr(col)}</label>
                                                <div class="flex items-center gap-1.5">
                                                    <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-amber-800 dark:text-amber-300 bg-amber-100 hover:bg-amber-200 dark:bg-amber-900/50 dark:hover:bg-amber-800/60 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-paste"></i> Paste</button>
                                                    <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-amber-800 dark:text-amber-300 bg-amber-100 hover:bg-amber-200 dark:bg-amber-900/50 dark:hover:bg-amber-800/60 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-upload"></i> Upload<input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124"></label>
                                                    <button type="button" data-ev-click="a135" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded"><i class="fa-solid fa-trash-can"></i> Remove</button>
                                                </div>
                                            </div>
                                            <textarea data-ev-paste="a127" class="existing-charge-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-amber-500 outline-none" rows="2">${escapeHtmlAttr(cSplit.text).replace(/&quot;/g, '"')}</textarea>
                                            <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(cSplit.images)}</div>
                                        </div>
                                    `;
            }).join('')}
                            </div>
                            <div id="${chargeContainerId}" class="space-y-3 pt-1"></div>
                            <button type="button" data-ev-click="a134" data-ev-click-args="${idcmPackArgs([chargeContainerId, chargeNoticeId])}" class="w-full py-2 border-2 border-dashed border-amber-300 dark:border-amber-700 hover:border-amber-500 text-amber-700 dark:text-amber-400 hover:bg-amber-50/50 dark:hover:bg-amber-900/20 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 mt-2">
                                <i class="fa-solid fa-circle-plus"></i> Add Charge
                            </button>
                        </div>
                    `;
        }

        if (secIdx === 0 && !sec.isChargeLike) {
            const benContainerId = isEdit ? "newBenefitsContainer" : "manualNewBenefitsContainer";
            const benNoticeId = isEdit ? "emptyBenefitNotice" : "manualEmptyBenefitNotice";
            const benDatalistId = isEdit ? "existingBenefitNamesList" : "manualBenefitNamesList";
            const filledBenefits = (isEdit && productObj) ? regularBenefitCols : [];
            const presetList = [...new Set([...(sec.items || []), ...getRegularBenefitColumns()])];

            return `
                        <hr class="my-3 dark:border-neutral-700">
                        <div class="flex justify-between items-center mb-2">
                            <h5 class="font-bold text-blue-900 dark:text-blue-300 text-sm">${escapeHtmlAttr(sec.title)}</h5>
                            <button type="button" data-ev-click="a136" data-ev-click-args="${idcmPackArgs([benContainerId, benNoticeId])}" class="px-2.5 py-1.5 bg-green-600 hover:bg-green-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm">
                                <i class="fa-solid fa-plus"></i> Add Benefit
                            </button>
                        </div>
                        <datalist id="${benDatalistId}">${presetList.map(c => `<option value="${escapeHtmlAttr(c)}"></option>`).join('')}</datalist>
                        <div ${isEdit ? 'id="existingBenefitsContainer"' : ''} class="space-y-2.5">
                            ${filledBenefits.length === 0 ? `
                                <div id="${benNoticeId}" class="text-center py-4 bg-gray-50 dark:bg-neutral-700/40 rounded-lg border border-dashed border-gray-300 dark:border-neutral-600 text-gray-500 dark:text-neutral-400">
                                    No benefits added yet. Click <strong>+ Add Benefit</strong> to fill in product benefits.
                                </div>
                            ` : filledBenefits.map(col => {
                const benSplit = splitTextAndImages(productObj.attributes[col] || "");
                return `
                                    <div class="existing-benefit-row field-with-images-wrapper p-3 bg-gray-50 dark:bg-neutral-700/40 border border-gray-200 dark:border-neutral-700 rounded-lg" data-col-name="${escapeHtmlAttr(col)}">
                                        <div class="flex flex-wrap justify-between items-center mb-1.5 gap-1.5">
                                            <label class="font-bold text-gray-800 dark:text-neutral-200">${escapeHtmlAttr(col)}</label>
                                            <div class="flex items-center gap-1.5">
                                                <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-paste"></i> Paste</button>
                                                <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-upload"></i> Upload<input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124"></label>
                                                <button type="button" data-ev-click="a137" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded"><i class="fa-solid fa-trash-can"></i> Remove</button>
                                            </div>
                                        </div>
                                        <textarea data-ev-paste="a127" class="existing-benefit-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-blue-500 outline-none" rows="3">${escapeHtmlAttr(benSplit.text).replace(/&quot;/g, '"')}</textarea>
                                        <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(benSplit.images)}</div>
                                    </div>
                                `;
            }).join('')}
                        </div>
                        <div id="${benContainerId}" class="space-y-3 pt-1"></div>
                        <button type="button" data-ev-click="a136" data-ev-click-args="${idcmPackArgs([benContainerId, benNoticeId])}" class="w-full py-2 border-2 border-dashed border-blue-300 dark:border-blue-700 hover:border-blue-500 text-blue-600 dark:text-blue-400 hover:bg-blue-50/50 dark:hover:bg-neutral-700/70 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 mt-2">
                            <i class="fa-solid fa-circle-plus"></i> Add Benefit
                        </button>
                    `;
        }

        // Custom Repeatable Section (Standard Text & Image)
        const customSecContainerId = `${scopePrefix}CustomSecNew_${secIdx}`;
        const customSecNoticeId = `${scopePrefix}CustomSecNotice_${secIdx}`;
        const customSecDatalistId = `${scopePrefix}CustomSecDatalist_${secIdx}`;
        const presetStr = (sec.items || []).join("||");
        const filledSectionItems = (isEdit && productObj)
            ? (sec.items || []).filter(col => !isEmptyVal(productObj.attributes[col]))
            : [];

        return `
                    <div class="dynamic-section-block pt-2 space-y-2.5" data-section-title="${escapeHtmlAttr(sec.title)}">
                        <hr class="my-3 dark:border-neutral-700">
                        <div class="flex justify-between items-center mb-1">
                            <h5 class="font-bold text-indigo-800 dark:text-indigo-300 text-sm flex items-center gap-1.5">
                                <i class="fa-solid fa-layer-group text-indigo-600 dark:text-indigo-400"></i> ${escapeHtmlAttr(sec.title)}
                            </h5>
                            <button type="button" data-ev-click="a138" data-ev-click-args="${idcmPackArgs([customSecContainerId, customSecNoticeId, sec.title, customSecDatalistId])}" class="px-2.5 py-1.5 bg-indigo-600 hover:bg-indigo-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm shrink-0">
                                <i class="fa-solid fa-plus"></i> Add ${escapeHtmlAttr(sec.title)}
                            </button>
                        </div>
                        <datalist id="${customSecDatalistId}" class="dynamic-section-datalist" data-preset-items="${escapeHtmlAttr(presetStr)}">
                            ${(sec.items || []).map(it => `<option value="${escapeHtmlAttr(it)}"></option>`).join('')}
                        </datalist>
                        <div class="space-y-2.5">
                            ${filledSectionItems.length === 0 ? `
                                <div id="${customSecNoticeId}" class="text-center py-4 bg-indigo-50/40 dark:bg-neutral-800/80 rounded-lg border border-dashed border-indigo-300 dark:border-neutral-600 text-gray-500 dark:text-neutral-400">
                                    No ${escapeHtmlAttr(sec.title)} items added yet. Click <strong>+ Add ${escapeHtmlAttr(sec.title)}</strong> to fill in details.
                                </div>
                            ` : filledSectionItems.map(col => {
            const sSplit = splitTextAndImages(productObj.attributes[col] || "");
            return `
                                    <div class="existing-section-row field-with-images-wrapper p-3 bg-indigo-50/40 dark:bg-neutral-800/80 border border-indigo-200 dark:border-neutral-600 rounded-lg" data-col-name="${escapeHtmlAttr(col)}">
                                        <div class="flex flex-wrap justify-between items-center mb-1.5 gap-1.5">
                                            <label class="font-bold text-indigo-900 dark:text-indigo-200">${escapeHtmlAttr(col)}</label>
                                            <div class="flex items-center gap-1.5">
                                                <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-700 dark:text-indigo-300 bg-indigo-100 hover:bg-indigo-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-paste"></i> Paste</button>
                                                <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-indigo-700 dark:text-indigo-300 bg-indigo-100 hover:bg-indigo-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-upload"></i> Upload<input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124"></label>
                                                <button type="button" data-ev-click="a125" data-ev-click-args="${idcmPackArgs([customSecNoticeId])}" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded"><i class="fa-solid fa-trash-can"></i> Remove</button>
                                            </div>
                                        </div>
                                        <textarea data-ev-paste="a127" class="existing-section-item-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-indigo-500 outline-none" rows="2">${escapeHtmlAttr(sSplit.text).replace(/&quot;/g, '"')}</textarea>
                                        <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(sSplit.images)}</div>
                                    </div>
                                `;
        }).join('')}
                        </div>
                        <div id="${customSecContainerId}" class="space-y-3 pt-1"></div>
                        <button type="button" data-ev-click="a138" data-ev-click-args="${idcmPackArgs([customSecContainerId, customSecNoticeId, sec.title, customSecDatalistId])}" class="w-full py-2 border-2 border-dashed border-indigo-300 dark:border-indigo-700 hover:border-indigo-500 text-indigo-700 dark:text-indigo-400 hover:bg-indigo-50/50 dark:hover:bg-neutral-700/70 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5 mt-2">
                            <i class="fa-solid fa-circle-plus"></i> Add ${escapeHtmlAttr(sec.title)}
                        </button>
                    </div>
                `;
    }).join('');

    const addBenState = getFieldState("Additional Benefit");

    return `
                <!-- SECTION 2: GENERAL PROVISIONS & SPECIFICATIONS -->
                <hr class="my-3 dark:border-neutral-700">
                <div class="flex justify-between items-center mb-2">
                    <h5 class="font-bold text-blue-900 dark:text-blue-300 text-sm">Product Provisions &amp; Specifications</h5>
                    <button type="button" data-ev-click="a139" data-ev-click-args="${idcmPackArgs([customNbContainerId])}" class="px-2.5 py-1 bg-slate-600 hover:bg-slate-500 text-white rounded-lg text-xs font-bold transition flex items-center gap-1 shadow-sm">
                        <i class="fa-solid fa-plus"></i> Add Provision
                    </button>
                </div>

                ${gridGenFields.length > 0 ? `<div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">${gridGenFields.join('')}</div>` : ''}
                ${fullWidthGenHtml}

                <!-- CUSTOM / EXTRA NON-BENEFIT PROVISIONS -->
                <datalist id="${existingNbListId}">${getAllKnownNonBenefitColumns().map(col => `<option value="${escapeHtmlAttr(col)}"></option>`).join('')}</datalist>
                ${isEdit ? `
                    <div id="editExistingNonBenefitsContainer" class="space-y-2.5">
                        ${existingCustomNbCols.map(col => {
        const nbSplit = getFieldState(col);
        return `
                                <div class="existing-nonbenefit-row field-with-images-wrapper p-3 bg-slate-50 dark:bg-neutral-800/60 border border-slate-200 dark:border-neutral-700 rounded-lg" data-col-name="${escapeHtmlAttr(col)}">
                                    <div class="flex flex-wrap justify-between items-center mb-1.5 gap-1.5">
                                        <label class="font-bold text-slate-800 dark:text-neutral-200">${escapeHtmlAttr(col)}</label>
                                        <div class="flex items-center gap-1.5">
                                            <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-700 dark:text-neutral-300 bg-slate-200 hover:bg-slate-300 dark:bg-neutral-700 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-paste"></i> Paste</button>
                                            <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-slate-700 dark:text-neutral-300 bg-slate-200 hover:bg-slate-300 dark:bg-neutral-700 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition"><i class="fa-solid fa-upload"></i> Upload<input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124"></label>
                                            <button type="button" data-ev-click="a140" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded"><i class="fa-solid fa-trash-can"></i> Remove</button>
                                        </div>
                                    </div>
                                    <textarea data-ev-paste="a127" class="existing-nonbenefit-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-slate-500 outline-none" rows="2">${escapeHtmlAttr(nbSplit.text).replace(/&quot;/g, '"')}</textarea>
                                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(nbSplit.images)}</div>
                                </div>
                            `;
    }).join('')}
                    </div>
                ` : ''}
                <div id="${customNbContainerId}" class="space-y-2.5 empty:hidden"></div>
                <button type="button" data-ev-click="a139" data-ev-click-args="${idcmPackArgs([customNbContainerId])}" class="w-full py-1.5 border border-dashed border-slate-300 dark:border-neutral-600 hover:border-slate-500 text-slate-600 dark:text-neutral-300 hover:bg-slate-50 dark:hover:bg-neutral-800/40 rounded-lg text-xs font-bold transition flex items-center justify-center gap-1.5">
                    <i class="fa-solid fa-circle-plus"></i> Add Provision (Non-Benefit)
                </button>

                ${specialProvisionsHtml}

                ${sectionsHtml}

                <!-- SECTION 4: ADDITIONAL BENEFIT -->
                <hr class="my-3 dark:border-neutral-700">
                <div class="field-with-images-wrapper">
                    <div class="flex flex-wrap justify-between items-center gap-1.5 mb-1">
                        <label class="block font-bold text-blue-900 dark:text-blue-300 text-sm">Additional / Other Benefit</label>
                        <div class="flex items-center gap-1.5">
                            <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Paste image">
                                <i class="fa-solid fa-paste"></i> Paste Image
                            </button>
                            <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Upload image">
                                <i class="fa-solid fa-upload"></i> Upload
                                <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                            </label>
                        </div>
                    </div>
                    <textarea data-spec-col="Additional Benefit" rows="3" data-ev-paste="a127" class="spec-field-input w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs ${focusRing} outline-none" placeholder="Enter Additional / Other Benefit details or paste table image...">${escapeHtmlAttr(addBenState.text).replace(/&quot;/g, '"')}</textarea>
                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden">${renderImageAttachmentItemsHTML(addBenState.images)}</div>
                </div>
            `;
}

// --- DETAIL & EDIT MODAL PRODUK ---
async function openDetailModal(id) {
    const p = products.find(prod => prod.id === id);
    if (!p) return;

    currentActiveProduct = p;
    isEditMode = false;

    const btnToggleEdit = document.getElementById("btnToggleEdit");
    const btnDelSingle = document.getElementById("btnDeleteSingleDetail");
    const detailModalEl = document.getElementById("detailModal");

    document.getElementById("detailEditFooter").classList.add("hidden");
    document.getElementById("detailTitle").textContent = p.name;
    detailModalEl.classList.remove("hidden");

    // Muat daftar dokumen Drive secara on-demand jika belum selesai
    ensureDocumentsLoaded();

    if (!isFullProductDataLoaded) {
        btnToggleEdit.classList.add("hidden");
        if (btnDelSingle) btnDelSingle.classList.add("hidden");
        document.getElementById("detailBody").innerHTML = `
                    <div class="py-12 text-center text-gray-500 dark:text-neutral-400 text-xs space-y-2">
                        <i class="fa-solid fa-spinner fa-spin text-2xl text-blue-600"></i>
                        <p class="font-medium">Loading full product details...</p>
                    </div>`;

        await ensureFullProductDataLoaded();
        if (!currentActiveProduct || currentActiveProduct.id !== id || isEditMode || detailModalEl.classList.contains("hidden")) {
            return;
        }
    }

    if (canEditData()) {
        btnToggleEdit.classList.remove("hidden");
    } else {
        btnToggleEdit.classList.add("hidden");
    }
    if (btnDelSingle) {
        btnDelSingle.classList.toggle("hidden", !isAdminUser());
    }
    renderDetailBodyView(p);
    scheduleHydrateDriveImages();
}

function renderDetailBodyView(p) {
    const getSpecVal = (targetCol) => {
        const actualKey = Object.keys(p.attributes).find(k => k.toLowerCase() === targetCol.toLowerCase()) || targetCol;
        return p.attributes[actualKey] || "-";
    };

    let basicDetailsHTML = `
                <div class="detail-row"><span class="detail-label">Company:</span> ${escapeHtml(p.company)}</div>
                <div class="detail-row"><span class="detail-label">Launch Date:</span> ${escapeHtml(formatDisplayDate(p.launchDate))}</div>
                <div class="detail-row"><span class="detail-label">Category System:</span> ${escapeHtml(p.category)}</div>
                <div class="detail-row"><span class="detail-label">Channel:</span> ${escapeHtml(p.channel)}</div>
                <div class="detail-row"><span class="detail-label">Product Type:</span> ${escapeHtml(p.productType)}</div>
                <div class="detail-row"><span class="detail-label">Currency:</span> ${escapeHtml(p.currency)}</div>
            `;

    const catCfg = getCategorySchemaConfig(p.category);
    const activeGenNames = [...catCfg.generalProvisions];
    DEFAULT_GENERAL_PROVISION_DEFS.forEach(def => {
        if (!activeGenNames.some(g => g.toLowerCase() === def.col.toLowerCase()) && !isEmptyVal(getSpecVal(def.col))) {
            activeGenNames.push(def.col);
        }
    });
    activeGenNames.forEach(col => {
        if (col.toLowerCase() === "available riders") return;
        const val = getSpecVal(col);
        if (!isEmptyVal(val)) {
            const displayLbl = (p.category === "DPLK" && col.toLowerCase() === "min premium") ? "Min Premium (Contribution)" : col;
            basicDetailsHTML += `<div class="detail-row"><span class="detail-label">${escapeHtml(displayLbl)}:</span><br>${formatListText(val)}</div>`;
        }
    });

    // Custom Non-Benefit Provisions (di luar General Provisions aktif & di luar Special Provisions)
    const activeGenLowerSet = new Set(activeGenNames.map(c => c.toLowerCase()));
    const catSpecialLowerSet = new Set((catCfg.specialProvisions || []).map(sp => sp.name.toLowerCase()));
    getCustomNonBenefitColumns().forEach(customCol => {
        const lc = customCol.toLowerCase();
        if (activeGenLowerSet.has(lc) || catSpecialLowerSet.has(lc) || isSpecialProvisionColumnAnyCategory(customCol)) return;
        const val = getSpecVal(customCol);
        if (!isEmptyVal(val)) {
            basicDetailsHTML += `<div class="detail-row"><span class="detail-label">${escapeHtml(customCol)}:</span><br>${formatListText(val)}</div>`;
        }
    });

    // Special Provisions Section — hanya tampil jika kategori memiliki Special Provisions yg terisi!
    let specialRowsHTML = "";
    (catCfg.specialProvisions || []).forEach(sp => {
        let val = getSpecVal(sp.name);
        if (isEmptyVal(val) && sp.name.toLowerCase() === "aqd") {
            val = getSpecVal("Akad");
        }
        if (!isEmptyVal(val)) {
            const spLbl = sp.name.toLowerCase() === "aqd" ? "Aqd / Contract"
                : (sp.name.toLowerCase() === "alokasi kontribusi" ? "Contribution Allocation"
                    : (sp.name.toLowerCase() === "surplus underwriting" ? "Underwriting Surplus" : sp.name));
            specialRowsHTML += `<div class="detail-row"><span class="detail-label">${escapeHtml(spLbl)}:</span><br>${formatListText(val)}</div>`;
        }
    });
    const isSyariahIcon = /^(sharia|syariah)$/i.test(catCfg.sheetName || "");
    const specialSectionHTML = specialRowsHTML ? `
                <div class="pt-4 mt-2 border-t border-emerald-200 dark:border-emerald-800/70">
                    <h4 class="font-bold text-emerald-800 dark:text-emerald-400 text-base mb-2 flex items-center gap-1.5">
                        <i class="fa-solid ${isSyariahIcon ? 'fa-mosque' : 'fa-star-of-life'} text-sm"></i> ${escapeHtml(catCfg.sheetName)} Special Provisions
                    </h4>
                    ${specialRowsHTML}
                </div>
            ` : '';

    // Helper renderer untuk daftar kolom deskriptif / checklist
    const renderDetailFeatureItemsHTML = (colsList, preferChecklist = false) => {
        let checklistHTML = "";
        let descriptiveHTML = "";
        colsList.forEach(attrKey => {
            const actualKey = Object.keys(p.attributes).find(k => k.toLowerCase() === attrKey.toLowerCase()) || attrKey;
            const val = p.attributes[actualKey];
            if (isEmptyVal(val)) return;
            const dplkParsed = parseDplkBenefitValue(val);
            if (dplkParsed.isChecked || preferChecklist) {
                const remarkText = dplkParsed.isChecked ? dplkParsed.remark : (isEmptyVal(val) ? "" : String(val));
                const remarkHtml = remarkText
                    ? `<span class="text-gray-600 dark:text-neutral-300 font-normal ms-1.5">&mdash; ${formatBulletLinesHtml(remarkText)}</span>`
                    : '';
                checklistHTML += `
                            <div class="flex items-start gap-2.5 px-3 py-2 rounded-lg bg-emerald-50/80 dark:bg-emerald-950/40 border border-emerald-200 dark:border-emerald-800/70 text-xs font-semibold text-gray-800 dark:text-neutral-200">
                                <span class="inline-flex items-center justify-center w-5 h-5 rounded-full bg-emerald-500 text-white text-[11px] shrink-0 mt-0.5"><i class="fa-solid fa-check"></i></span>
                                <div class="min-w-0 leading-relaxed"><span>${escapeHtml(actualKey)}</span>${remarkHtml}</div>
                            </div>
                        `;
            } else {
                descriptiveHTML += `
                            <div class="mt-3">
                                <span class="benefit-label">${escapeHtml(actualKey)}:</span>
                                <span class="benefit-value">${formatListText(val)}</span>
                            </div>
                        `;
            }
        });
        return `
                    ${checklistHTML ? `<div class="grid grid-cols-1 sm:grid-cols-2 gap-2 mt-2">${checklistHTML}</div>` : ''}
                    ${descriptiveHTML}
                `.trim();
    };

    const additionalBenefitVal = getSpecVal("Additional Benefit");
    const additionalBenefitHTML = !isEmptyVal(additionalBenefitVal) ? `
                <div class="pt-4 mt-4 border-t border-gray-200 dark:border-neutral-700">
                    <h4 class="font-bold text-blue-800 dark:text-blue-400 text-base mb-2">Additional / Other Benefit</h4>
                    <div class="benefit-value text-sm">${formatListText(additionalBenefitVal)}</div>
                </div>
            ` : '';

    // Dynamic Product Sections (Primary Benefits, Additional Benefit, Custom Sections, Charges)
    const assignedNonPrimaryCols = getSectionAssignedNonPrimaryColumnsForCategory(p.category);
    const primaryBenefitCols = getRegularBenefitColumns().filter(col => !assignedNonPrimaryCols.has(col.toLowerCase()));
    let dynamicSectionsHTML = "";
    let renderedAdditionalBenefit = false;

    (catCfg.productSections || []).forEach((sec, secIdx) => {
        if (sec.isChargeLike) {
            const chargeCols = [...new Set([...(sec.items || []), ...getChargeColumns()])];
            let chargesRowsHTML = "";
            chargeCols.forEach(chargeKey => {
                const val = getSpecVal(chargeKey);
                if (!isEmptyVal(val)) {
                    chargesRowsHTML += `
                                <div class="mt-3">
                                    <span class="benefit-label text-amber-800 dark:text-amber-400">${escapeHtml(chargeKey)}:</span>
                                    <span class="benefit-value">${formatListText(val)}</span>
                                </div>
                            `;
                }
            });
            if (chargesRowsHTML) {
                dynamicSectionsHTML += `
                            <div class="pt-4 mt-4 border-t border-amber-200 dark:border-amber-800/70">
                                <h4 class="font-bold text-amber-800 dark:text-amber-400 text-base mb-2 flex items-center gap-1.5">
                                    <i class="fa-solid fa-receipt text-sm"></i> ${escapeHtml(sec.title)}
                                </h4>
                                ${chargesRowsHTML}
                            </div>
                        `;
            }
            return;
        }

        if (secIdx === 0) {
            const primaryHtml = renderDetailFeatureItemsHTML(primaryBenefitCols, false);
            dynamicSectionsHTML += `
                        <div class="pt-4 mt-2 border-t border-gray-200 dark:border-neutral-700">
                            <h4 class="font-bold text-blue-800 dark:text-blue-400 text-base mb-2">${escapeHtml(sec.title || "Product Benefits")}</h4>
                            ${primaryHtml || '<p class="text-xs text-gray-500 dark:text-neutral-400">No additional benefit details.</p>'}
                        </div>
                    `;
            if (additionalBenefitHTML) {
                dynamicSectionsHTML += additionalBenefitHTML;
                renderedAdditionalBenefit = true;
            }
            return;
        }

        const customSecHtml = renderDetailFeatureItemsHTML(sec.items || [], sec.mode === "checklist");
        if (customSecHtml) {
            dynamicSectionsHTML += `
                        <div class="pt-4 mt-4 border-t border-indigo-200 dark:border-neutral-600">
                            <h4 class="font-bold text-indigo-800 dark:text-indigo-300 text-base mb-2 flex items-center gap-1.5">
                                <i class="fa-solid fa-layer-group text-sm"></i> ${escapeHtml(sec.title)}
                            </h4>
                            ${customSecHtml}
                        </div>
                    `;
        }
    });

    if (!renderedAdditionalBenefit && additionalBenefitHTML) {
        dynamicSectionsHTML += additionalBenefitHTML;
    }

    // Fallback jika suatu produk punya Charges di sheet padahal categoryConfig-nya tidak mencantumkan section Charge
    if (!(catCfg.productSections || []).some(s => s.isChargeLike)) {
        let fallbackChargesHTML = "";
        getChargeColumns().forEach(chargeKey => {
            const val = p.attributes[chargeKey];
            if (!isEmptyVal(val)) {
                fallbackChargesHTML += `<div class="mt-3"><span class="benefit-label text-amber-800 dark:text-amber-400">${escapeHtml(chargeKey)}:</span><span class="benefit-value">${formatListText(val)}</span></div>`;
            }
        });
        if (fallbackChargesHTML) {
            dynamicSectionsHTML += `<div class="pt-4 mt-4 border-t border-amber-200 dark:border-amber-800/70"><h4 class="font-bold text-amber-800 dark:text-amber-400 text-base mb-2 flex items-center gap-1.5"><i class="fa-solid fa-receipt text-sm"></i> Charges</h4>${fallbackChargesHTML}</div>`;
        }
    }

    // Available Riders Section (separate below)
    const ridersVal = getSpecVal("Available Riders");
    const ridersHTML = !isEmptyVal(ridersVal) ? `
                <div class="pt-4 mt-4 border-t border-gray-200 dark:border-neutral-700">
                    <h4 class="font-bold text-blue-800 dark:text-blue-400 text-base mb-2">Available Riders</h4>
                    <div class="benefit-value text-sm">${formatListText(ridersVal)}</div>
                </div>
            ` : '';

    const pDocs = getDocumentsForProduct(p.name);
    const documentsSectionHTML = renderDocumentsSectionView(p, pDocs);

    // Remarks Section - right below Product Documents
    const remarksVal = getSpecVal("Remarks");
    const remarksSectionHTML = `
                <div class="pt-4 mt-4 border-t border-gray-200 dark:border-neutral-700" id="detailRemarksSection">
                    <h4 class="font-bold text-gray-800 dark:text-neutral-200 text-base mb-2 flex items-center gap-2">
                        <i class="fa-solid fa-note-sticky text-amber-500"></i> Remarks
                    </h4>
                    ${!isEmptyVal(remarksVal)
            ? `<div class="benefit-value text-sm">${formatListText(remarksVal)}</div>`
            : `<p class="text-xs text-gray-500 dark:text-neutral-400">No remarks for this product yet.</p>`}
                </div>
            `;

    document.getElementById("detailBody").innerHTML = `
                ${basicDetailsHTML}
                ${specialSectionHTML}
                ${dynamicSectionsHTML}
                ${ridersHTML}
                ${documentsSectionHTML}
                ${remarksSectionHTML}
            `;
}

function renderDocumentsSectionView(p, pDocs) {
    const safeJsStr = (s) => escapeHtmlAttr(
        String(s || "")
            .replace(/\\/g, "\\\\")
            .replace(/'/g, "\\'")
            .replace(/\r/g, "\\r")
            .replace(/\n/g, "\\n")
            .replace(/\u2028/g, "\\u2028")
            .replace(/\u2029/g, "\\u2029")
    );
    const safeName = safeJsStr(p.name);
    return `
                <div class="pt-4 mt-4 border-t border-gray-200 dark:border-neutral-700" id="detailDocumentsSection">
                    <div class="flex flex-wrap justify-between items-center gap-2 mb-3">
                        <h4 class="font-bold text-gray-800 dark:text-neutral-200 text-sm sm:text-base flex items-center gap-2">
                            <i class="fa-solid fa-folder-open text-blue-600 dark:text-blue-400" aria-hidden="true"></i> Product Documents
                            <span class="text-xs px-2 py-0.5 rounded-full bg-blue-100 dark:bg-neutral-700/70 text-blue-800 dark:text-blue-300 font-semibold">
                                ${!isDocumentsLoaded ? `<i class="fa-solid fa-spinner fa-spin me-1"></i>Loading...` : `${pDocs.length} file${pDocs.length === 1 ? '' : 's'}`}
                            </span>
                        </h4>
                        ${canEditData() ? `<button type="button" data-ev-click="a141" data-ev-click-args="${idcmPackArgs([idcmUnescapeJsStr(safeName)])}"
                            class="modal-touch-btn text-xs font-semibold px-3 py-1.5 rounded-lg bg-blue-600 hover:bg-blue-500 text-white transition flex items-center gap-1.5 shadow-sm">
                            <i class="fa-solid fa-cloud-arrow-up" aria-hidden="true"></i> Upload Document
                        </button>` : ''}
                    </div>
                    ${!isDocumentsLoaded ? `
                        <div class="p-4 border border-dashed border-gray-300 dark:border-neutral-700 rounded-xl text-center text-xs text-gray-500 dark:text-neutral-400 bg-gray-50/50 dark:bg-neutral-800/40">
                            <i class="fa-solid fa-spinner fa-spin text-xl text-blue-500 mb-1.5 block"></i>
                            Checking product documents in Google Drive...
                        </div>
                    ` : pDocs.length === 0 ? `
                        <div class="p-4 border border-dashed border-gray-300 dark:border-neutral-700 rounded-xl text-center text-xs text-gray-500 dark:text-neutral-400 bg-gray-50/50 dark:bg-neutral-800/40">
                            <i class="fa-solid fa-file-circle-plus text-2xl text-gray-400 mb-1.5 block"></i>
                            No documents uploaded for this product yet.
                        </div>
                    ` : `
                        <div class="space-y-2">
                            ${pDocs.map(doc => {
        const fileInfo = getDocFileIcon(doc.fileName);
        const metaParts = [];
        if (doc.fileSize) metaParts.push(`<span>${formatFileSize(doc.fileSize)}</span>`);
        if (doc.uploadDate) metaParts.push(`<span>${escapeHtml(doc.uploadDate.split(' ')[0])}</span>`);
        const safeFileId = safeJsStr(doc.fileId);
        const safeFileNameJs = safeJsStr(doc.fileName);
        const safeDownloadUrlJs = safeJsStr(doc.downloadUrl || "");
        return `
                                    <div class="flex flex-col sm:flex-row sm:items-center justify-between gap-2 p-2.5 bg-gray-50 dark:bg-neutral-700/50 rounded-xl border border-gray-200 dark:border-neutral-600 hover:border-blue-400 transition">
                                        <div class="flex items-center gap-2.5 min-w-0">
                                            <div class="w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border ${fileInfo.color}">
                                                <i class="fa-solid ${fileInfo.icon} text-sm" aria-hidden="true"></i>
                                            </div>
                                            <div class="min-w-0 flex-1">
                                                <div class="text-xs font-bold text-gray-800 dark:text-neutral-200 truncate" title="${escapeHtmlAttr(doc.fileName)}">${escapeHtml(doc.fileName)}</div>
                                                ${metaParts.length > 0 ? `
                                                <div class="flex items-center gap-1.5 text-[10px] text-gray-500 dark:text-neutral-400 mt-0.5">
                                                    ${metaParts.join('<span>&bull;</span>')}
                                                </div>` : ''}
                                            </div>
                                        </div>
                                        <div class="flex items-center justify-end gap-1.5 shrink-0">
                                            <button type="button" data-ev-click="a142" data-ev-click-args="${idcmPackArgs([idcmUnescapeJsStr(safeFileId), idcmUnescapeJsStr(safeFileNameJs), idcmUnescapeJsStr(safeDownloadUrlJs)])}" class="modal-touch-btn px-2.5 py-1.5 sm:py-1 text-xs font-semibold rounded-lg bg-blue-50 dark:bg-neutral-700/70 text-blue-700 dark:text-blue-300 hover:bg-blue-100 dark:hover:bg-neutral-700/70 transition flex items-center gap-1" title="Preview Document Inline">
                                                <i class="fa-solid fa-eye text-[10px]" aria-hidden="true"></i> View
                                            </button>
                                            <button type="button" data-ev-click="a143" data-ev-click-args="${idcmPackArgs([idcmUnescapeJsStr(safeFileId), idcmUnescapeJsStr(safeFileNameJs), idcmUnescapeJsStr(safeDownloadUrlJs)])}" class="modal-touch-btn px-2.5 py-1.5 sm:py-1 text-xs font-semibold rounded-lg bg-emerald-50 dark:bg-emerald-950/40 text-emerald-700 dark:text-emerald-300 hover:bg-emerald-100 dark:hover:bg-emerald-950/60 disabled:opacity-60 transition flex items-center gap-1" title="Download Original File Directly">
                                                <i class="fa-solid fa-download text-[10px]" aria-hidden="true"></i> <span>Download</span>
                                            </button>
                                            ${isAdminUser() ? `<button type="button" data-ev-click="a144" data-ev-click-args="${idcmPackArgs([idcmUnescapeJsStr(safeFileId), idcmUnescapeJsStr(safeName)])}" aria-label="Delete Document" class="modal-touch-btn w-8 h-8 sm:w-7 sm:h-7 rounded-lg text-red-500 hover:bg-red-50 dark:hover:bg-red-950/40 flex items-center justify-center transition" title="Delete Document">
                                                <i class="fa-solid fa-trash-can text-xs" aria-hidden="true"></i>
                                            </button>` : ''}
                                        </div>
                                    </div>
                                `;
    }).join('')}
                        </div>
                    `}
                </div>
            `;
}

function toggleEditMode() {
    if (!currentActiveProduct) return;
    isEditMode = true;

    document.getElementById("btnToggleEdit").classList.add("hidden");
    const btnDelSingle = document.getElementById("btnDeleteSingleDetail");
    if (btnDelSingle) btnDelSingle.classList.add("hidden");
    document.getElementById("detailEditFooter").classList.remove("hidden");

    const btnSave = document.getElementById("btnSaveProductEdit");
    if (btnSave) {
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Changes`;
    }

    const p = currentActiveProduct;

    // Helper to retrieve value from non-benefit provision columns
    const getSpecAttr = (targetCol) => {
        const actualKey = Object.keys(p.attributes).find(k => k.toLowerCase() === targetCol.toLowerCase()) || targetCol;
        const val = p.attributes[actualKey];
        return (!val || val === "-") ? "" : val;
    };

    let html = `
                <div class="space-y-3 text-xs">
                    <div class="bg-blue-50 dark:bg-neutral-700/70 p-2.5 rounded-lg border border-blue-200 dark:border-neutral-600 text-blue-800 dark:text-blue-200 font-semibold mb-3">
                        <i class="fa-solid fa-circle-info me-1"></i> Editing Product: <strong>${escapeHtml(p.name)}</strong> (${escapeHtml(p.category)})
                    </div>

                    <!-- Target Category & Product Name -->
                    <div class="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Category / Sheet Tab <span class="text-red-500">*</span></label>
                            <div class="flex items-center gap-1.5">
                                <select id="edit_category" data-ev-change="a145" class="flex-1 min-w-0 border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                                    ${SHEET_TABS.map(tab => `<option value="${escapeHtmlAttr(tab)}" ${p.category === tab ? 'selected' : ''}>${escapeHtmlAttr(tab)}</option>`).join('')}
                                </select>
                                <button type="button" data-ev-click="a146" class="px-2 py-2 bg-gray-100 hover:bg-gray-200 dark:bg-neutral-700 dark:hover:bg-neutral-600 border dark:border-neutral-600 text-gray-700 dark:text-neutral-200 rounded-lg text-xs transition shrink-0" title="Edit this category's provisions & sections">
                                    <i class="fa-solid fa-gear"></i>
                                </button>
                                <button type="button" data-ev-click="a46" class="px-2 py-2 bg-blue-50 hover:bg-blue-100 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 border border-blue-200 dark:border-neutral-600 text-blue-700 dark:text-blue-300 rounded-lg text-xs transition shrink-0" title="Create a new category sheet">
                                    <i class="fa-solid fa-folder-plus"></i>
                                </button>
                            </div>
                            <div id="editCategoryFeedback" class="hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                        </div>
                        <div class="sm:col-span-2">
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Product Name <span class="text-red-500">*</span></label>
... [truncated for diff preview]
                            <input type="text" id="edit_productName" value="${escapeHtmlAttr(p.name)}" data-ev-input="a145" placeholder="Enter product name..." class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                            <div id="editProductNameFeedback" class="hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                        </div>
                    </div>

                    <!-- SECTION 1: BASIC INFORMATION -->
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Company Name <span class="text-red-500">*</span></label>
                        <input type="text" id="edit_company" value="${escapeHtmlAttr(p.company)}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Product Type</label>
                            <input type="text" id="edit_productType" value="${escapeHtmlAttr(p.productType)}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                        </div>
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Distribution Channel</label>
                            <input type="text" id="edit_channel" value="${escapeHtmlAttr(p.channel)}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                        </div>
                    </div>

                    <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Currency</label>
                            <input type="text" id="edit_currency" value="${escapeHtmlAttr(p.currency)}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                        </div>
                        <div>
                            <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Launch Date</label>
                            <input type="date" id="edit_launchDate" value="${formatDateForInput(p.launchDate)}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-blue-500 outline-none">
                        </div>
                    </div>

                    <div id="editDynamicCategorySections" data-rendered-cat="${escapeHtmlAttr(p.category)}">
                        ${renderCategoryDynamicFormSections('edit', p.category, p)}
                    </div>

                    <!-- SECTION 5: REMARKS -->
                    <hr class="my-3 dark:border-neutral-700">
                    <div class="field-with-images-wrapper">
                        <label class="block font-bold text-blue-900 dark:text-blue-300 text-sm mb-1">Remarks</label>
                        <textarea data-spec-col="Remarks" rows="3" class="spec-field-input w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-blue-500 outline-none" placeholder="Notes about this product...">${getSpecAttr("Remarks").replace(/&/g, "&amp;").replace(/</g, "&lt;")}</textarea>
                    </div>
                </div>
            `;

    const detailBodyEl = document.getElementById("detailBody");
    detailBodyEl.innerHTML = html;
    refreshBenefitDatalist(detailBodyEl);
    refreshChargeDatalist(detailBodyEl);
    refreshNonBenefitDatalist(detailBodyEl);
    refreshSectionItemDatalist(detailBodyEl);
    validateEditProductIdentity();
}

function onDplkBenefitCheckboxChange(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const activeClasses = ['bg-white', 'dark:bg-neutral-700', 'border-emerald-400', 'dark:border-emerald-500', 'text-gray-900', 'dark:text-neutral-100', 'focus:ring-2', 'focus:ring-emerald-500'];
    const disabledClasses = ['bg-gray-100', 'dark:bg-neutral-900/60', 'border-gray-200', 'dark:border-neutral-700', 'text-gray-400', 'dark:text-neutral-500', 'opacity-60', 'cursor-not-allowed'];
    container.querySelectorAll('.dplk-benefit-item').forEach(itemEl => {
        const chk = itemEl.querySelector('.dplk-benefit-chk');
        const remInput = itemEl.querySelector('.dplk-benefit-remark');
        if (chk && remInput) {
            remInput.disabled = !chk.checked;
            if (chk.checked) {
                remInput.classList.remove(...disabledClasses);
                remInput.classList.add(...activeClasses);
            } else {
                remInput.classList.remove(...activeClasses);
                remInput.classList.add(...disabledClasses);
            }
        }
    });
    const allCount = container.querySelectorAll('.dplk-benefit-chk').length;
    const checkedCount = container.querySelectorAll('.dplk-benefit-chk:checked').length;
    const badgeId = containerId === 'manualDplkBenefitChecklist'
        ? 'manualDplkChecklistCount'
        : (containerId === 'editDplkBenefitChecklist'
            ? 'editDplkChecklistCount'
            : containerId.replace('ChecklistSec_', 'ChecklistCount_'));
    const badgeEl = document.getElementById(badgeId);
    if (badgeEl) {
        badgeEl.textContent = `${checkedCount} / ${allCount} selected`;
    }
    const scopeRoot = container.closest('#manualFormBody, #detailBody') || document;
    refreshBenefitDatalist(scopeRoot);
}

function addDplkCustomBenefitRow(containerId) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const listEl = container.querySelector('.dplk-custom-benefits-list');
    if (!listEl) return;
    const row = document.createElement('div');
    row.className = "dplk-benefit-item dplk-custom-benefit-item flex flex-col sm:flex-row sm:items-center gap-2 p-2 rounded-lg border border-emerald-300 dark:border-emerald-700 bg-white dark:bg-neutral-800 transition";
    row.innerHTML = `
                <div class="flex items-center gap-2 sm:w-5/12 shrink-0">
                    <input type="checkbox" class="dplk-benefit-chk w-4 h-4 rounded accent-emerald-600 cursor-pointer shrink-0" checked data-ev-change="a131" data-ev-change-args="${idcmPackArgs([containerId])}">
                    <input type="text" class="dplk-custom-benefit-name flex-1 min-w-0 border border-gray-300 dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded px-2 py-1 text-xs font-medium focus:ring-1 focus:ring-emerald-500 outline-none" placeholder="Benefit name..." data-ev-input="a131" data-ev-input-args="${idcmPackArgs([containerId])}">
                    <button type="button" data-ev-click="a132" data-ev-click-args="${idcmPackArgs([containerId])}" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition px-1" title="Remove benefit">
                        <i class="fa-solid fa-trash-can text-xs"></i>
                    </button>
                </div>
                <input type="text" class="dplk-benefit-remark flex-1 w-full rounded px-2.5 py-1.5 text-xs outline-none transition border bg-white dark:bg-neutral-700 border-emerald-400 dark:border-emerald-500 text-gray-900 dark:text-neutral-100 focus:ring-2 focus:ring-emerald-500" placeholder="Add note (optional)...">
            `;
    listEl.appendChild(row);
    onDplkBenefitCheckboxChange(containerId);
    const nameInput = row.querySelector('.dplk-custom-benefit-name');
    if (nameInput) nameInput.focus();
}

function toggleAllDplkBenefitCheckboxes(containerId, checkedState) {
    const container = document.getElementById(containerId);
    if (!container) return;
    container.querySelectorAll('.dplk-benefit-chk').forEach(chk => {
        chk.checked = Boolean(checkedState);
    });
    onDplkBenefitCheckboxChange(containerId);
}

// Memperbarui opsi autocomplete <datalist> hanya dengan kolom MANFAAT yang belum dipakai di form aktif
function refreshBenefitDatalist(scopeContainer = null) {
    const root = scopeContainer || document;
    const datalist = root.querySelector('#manualBenefitNamesList, #existingBenefitNamesList') || document.getElementById("existingBenefitNamesList");
    if (!datalist) return;

    const usedNames = getUsedBenefitNamesLower(null, root);
    datalist.innerHTML = getRegularBenefitColumns()
        .filter(col => !usedNames.includes(col.toLowerCase()))
        .map(col => `<option value="${escapeHtmlAttr(col)}"></option>`)
        .join('');
}

function refreshChargeDatalist(scopeContainer = null) {
    const root = scopeContainer || document;
    const datalist = root.querySelector('#manualChargeNamesList, #existingChargeNamesList') || document.getElementById("existingChargeNamesList");
    if (!datalist) return;

    const usedCharges = getUsedChargeNamesLower(null, root);
    const allKnownCharges = [...new Set([...DPLK_CHARGE_COLUMNS, ...getChargeColumns()])];
    datalist.innerHTML = allKnownCharges
        .filter(col => !usedCharges.includes(col.toLowerCase()))
        .map(col => `<option value="${escapeHtmlAttr(col)}"></option>`)
        .join('');
}

function refreshNonBenefitDatalist(scopeContainer = null) {
    const root = scopeContainer || document;
    const datalist = root.querySelector('#manualNonBenefitNamesList, #existingNonBenefitNamesList') || document.getElementById("existingNonBenefitNamesList");
    if (!datalist) return;

    const usedNonBenefits = getUsedNonBenefitNamesLower(null, root);
    datalist.innerHTML = getAllKnownNonBenefitColumns()
        .filter(col => !usedNonBenefits.includes(col.toLowerCase()))
        .map(col => `<option value="${escapeHtmlAttr(col)}"></option>`)
        .join('');
}

function getUsedNonBenefitNamesLower(excludeInputEl = null, scopeContainer = null) {
    const root = scopeContainer
        || (excludeInputEl && excludeInputEl.closest('#manualFormBody, #detailBody'))
        || document;
    const names = [
        "product name", "company name", "par/non par", "product type",
        "distribution channel", "currency", "launch date", "tanggal rilis"
    ];
    root.querySelectorAll('.spec-field-input[data-spec-col]').forEach(el => {
        const col = el.getAttribute('data-spec-col');
        if (col && !el.closest('.hidden')) {
            names.push(col.trim().toLowerCase());
        }
    });
    root.querySelectorAll('.existing-nonbenefit-row').forEach(row => {
        const colName = row.getAttribute('data-col-name');
        if (colName) names.push(colName.trim().toLowerCase());
    });
    root.querySelectorAll('.new-nonbenefit-name').forEach(input => {
        if (input !== excludeInputEl) {
            const val = input.value.trim().toLowerCase();
            if (val) names.push(val);
        }
    });
    return names;
}

function clearSpecFieldInput(btnEl) {
    if (!btnEl) return;
    const wrapper = btnEl.closest('div')?.parentElement;
    if (!wrapper) return;
    const ta = wrapper.querySelector('textarea.spec-field-input');
    if (ta) {
        ta.value = "";
        ta.focus();
    }
}

function addNewNonBenefitRow(targetContainerId = "editNewNonBenefitsContainer") {
    const container = document.getElementById(targetContainerId);
    if (!container) return;
    const scopeRoot = container.closest('#manualFormBody, #detailBody') || document;
    const datalistId = scopeRoot.id === 'manualFormBody' ? 'manualNonBenefitNamesList' : 'existingNonBenefitNamesList';

    const rowDiv = document.createElement("div");
    rowDiv.className = "new-nonbenefit-row field-with-images-wrapper p-3 bg-slate-50 dark:bg-neutral-800/80 border border-slate-300 dark:border-neutral-600 rounded-xl space-y-2 relative transition";

    rowDiv.innerHTML = `
                <div class="flex justify-between items-center">
                    <span class="text-[11px] font-bold uppercase tracking-wider text-slate-700 dark:text-neutral-300">
                        <i class="fa-solid fa-sliders me-1"></i> Add Provision (Non-Benefit)
                    </span>
                    <div class="flex items-center gap-1.5">
                        <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-slate-700 dark:text-neutral-300 bg-slate-200 hover:bg-slate-300 dark:bg-neutral-700 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Paste screenshot / image">
                            <i class="fa-solid fa-paste"></i> Paste
                        </button>
                        <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-slate-700 dark:text-neutral-300 bg-slate-200 hover:bg-slate-300 dark:bg-neutral-700 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Upload image file">
                            <i class="fa-solid fa-upload"></i> Upload
                            <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                        </label>
                        <button type="button" data-ev-click="a140" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded" title="Cancel adding this provision">
                            <i class="fa-solid fa-xmark"></i> Cancel
                        </button>
                    </div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Provision / Non-Benefit Name</label>
                    <input type="text" list="${datalistId}" data-ev-input="a147"
                           placeholder="Type or select provision name (e.g. Billing Frequency, Grace Period)..."
                           class="new-nonbenefit-name w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-slate-500 outline-none">
                    <div class="nonbenefit-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Provision Value / Details</label>
                    <textarea rows="2" data-ev-paste="a127" placeholder="Enter details or paste (Ctrl+V) table image..."
                              class="new-nonbenefit-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-slate-500 outline-none"></textarea>
                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden"></div>
                </div>
            `;

    container.appendChild(rowDiv);
    const nameInput = rowDiv.querySelector('.new-nonbenefit-name');
    if (nameInput) nameInput.focus();
    refreshNonBenefitDatalist(scopeRoot);
}

function validateNonBenefitName(inputEl) {
    const rawVal = inputEl.value.trim();
    const lowerVal = rawVal.toLowerCase();
    const feedbackEl = inputEl.parentElement.querySelector('.nonbenefit-name-feedback');
    const scopeRoot = inputEl.closest('#manualFormBody, #detailBody') || document;

    if (!rawVal) {
        feedbackEl.className = "nonbenefit-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
        feedbackEl.replaceChildren();
        inputEl.dataset.status = "empty";
        delete inputEl.dataset.canonicalName;
        refreshNonBenefitDatalist(scopeRoot);
        return;
    }

    if (isDplkChargeColumn(rawVal)) {
        feedbackEl.className = "nonbenefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtmlAttr(rawVal)}"</strong> is a Charge component, not a Non-Benefit Provision.`;
        inputEl.dataset.status = "duplicate";
        refreshNonBenefitDatalist(scopeRoot);
        return;
    }

    const matchedBenefit = getRegularBenefitColumns().find(col => col.toLowerCase() === lowerVal) || resolveKnownBenefitName(rawVal);
    if (matchedBenefit && !isNonBenefitColumn(rawVal)) {
        feedbackEl.className = "nonbenefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtmlAttr(rawVal)}"</strong> belongs to Product Benefits (matches <em>${escapeHtmlAttr(matchedBenefit)}</em>), not Non-Benefits.`;
        inputEl.dataset.status = "duplicate";
        refreshNonBenefitDatalist(scopeRoot);
        return;
    }

    const usedInProduct = getUsedNonBenefitNamesLower(inputEl, scopeRoot);
    if (usedInProduct.includes(lowerVal)) {
        feedbackEl.className = "nonbenefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Provision <strong>"${escapeHtmlAttr(rawVal)}"</strong> is already shown or added in this form!`;
        inputEl.dataset.status = "duplicate";
        refreshNonBenefitDatalist(scopeRoot);
        return;
    }

    const resolvedExisting = resolveKnownNonBenefitName(rawVal);
    if (resolvedExisting) {
        feedbackEl.className = "nonbenefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> <strong>Registered Non-Benefit:</strong> Already exists in database columns (<em>${escapeHtmlAttr(resolvedExisting)}</em>).`;
        inputEl.dataset.status = "existing";
        inputEl.dataset.canonicalName = resolvedExisting;
    } else {
        feedbackEl.className = "nonbenefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-wand-magic-sparkles me-1"></i> <strong>New Non-Benefit:</strong> Will be added as a Provision column &amp; registered in Config.`;
        inputEl.dataset.status = "new";
        inputEl.dataset.canonicalName = rawVal;
    }

    refreshNonBenefitDatalist(scopeRoot);
}

function removeNonBenefitRow(btnEl) {
    const scopeRoot = btnEl.closest('#manualFormBody, #detailBody') || document;
    const row = btnEl.closest('.existing-nonbenefit-row, .new-nonbenefit-row');
    if (row) row.remove();

    scopeRoot.querySelectorAll('.new-nonbenefit-name').forEach(input => validateNonBenefitName(input));
    refreshNonBenefitDatalist(scopeRoot);
}

// Mengambil semua nama manfaat yang sedang aktif ditampilkan di form (Edit Produk atau Tambah Manual)
function getUsedBenefitNamesLower(excludeInputEl = null, scopeContainer = null) {
    const root = scopeContainer
        || (excludeInputEl && excludeInputEl.closest('#manualFormBody, #detailBody'))
        || document;
    const names = [];
    root.querySelectorAll('.existing-benefit-row').forEach(row => {
        const colName = row.getAttribute('data-col-name');
        if (colName) names.push(colName.trim().toLowerCase());
    });
    root.querySelectorAll('.dplk-benefit-chk:checked').forEach(chk => {
        const colName = chk.getAttribute('data-benefit-col');
        if (colName) names.push(colName.trim().toLowerCase());
    });
    root.querySelectorAll('.new-benefit-name').forEach(input => {
        if (input !== excludeInputEl) {
            const val = input.value.trim().toLowerCase();
            if (val) names.push(val);
        }
    });
    return names;
}

function getUsedChargeNamesLower(excludeInputEl = null, scopeContainer = null) {
    const root = scopeContainer
        || (excludeInputEl && excludeInputEl.closest('#manualFormBody, #detailBody'))
        || document;
    const names = [];
    root.querySelectorAll('.existing-charge-row').forEach(row => {
        const colName = row.getAttribute('data-col-name');
        if (colName) names.push(colName.trim().toLowerCase());
    });
    root.querySelectorAll('.new-charge-name').forEach(input => {
        if (input !== excludeInputEl) {
            const val = input.value.trim().toLowerCase();
            if (val) names.push(val);
        }
    });
    return names;
}

// --- HELPER LAMPIRAN GAMBAR TABEL PADA FORM EDIT & TAMBAH MANUAL (RESOLUSI PENUH / ZERO COMPRESSION) ---
function readFileAsDataUrl(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
    });
}

async function uploadTableImageToDrive(fileOrBlob, fullDataUrl) {
    const base64Data = fullDataUrl.split(",")[1];
    const mimeType = (fileOrBlob && fileOrBlob.type)
        || (fullDataUrl.match(/^data:([^;]+);/) || [])[1]
        || "image/png";
    const ext = (mimeType.split("/")[1] || "png").replace(/[^a-z0-9]/gi, "");
    const fileName = (fileOrBlob && fileOrBlob.name && fileOrBlob.name !== "image.png")
        ? fileOrBlob.name
        : `table_img_${Date.now()}_${Math.random().toString(36).slice(2, 7)}.${ext}`;

    const resp = await fetch(APPS_SCRIPT_URL, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify(withSessionToken({
            action: "upload_document",
            productName: "__TABLE_IMAGES__",
            docType: "__TABLE_IMAGE__",
            fileName: fileName,
            mimeType: mimeType,
            base64Data: base64Data
        }))
    });
    const res = await resp.json();
    if (handleUnauthorizedResponse(res)) {
        throw new Error(res.message || "Session expired.");
    }
    if (res && res.status === "success" && res.fileId) {
        driveImageCache[res.fileId] = fullDataUrl;
        return res.fileId;
    }
    throw new Error((res && res.message) || "Failed to upload full-resolution image to Google Drive");
}

function renderImageAttachmentItemsHTML(imagesArray) {
    if (!imagesArray || imagesArray.length === 0) return "";
    let needsHydration = false;
    const html = imagesArray.map((src, idx) => {
        const isDrive = String(src).startsWith("drive:");
        const driveFileId = isDrive ? String(src).slice(6) : "";
        const resolvedSrc = isDrive ? (driveImageCache[driveFileId] || "") : src;
        const isUnresolvedDrive = isDrive && !resolvedSrc;
        if (isUnresolvedDrive) needsHydration = true;

        const rawActionSrc = isDrive ? `drive:${driveFileId}` : src;
        const safeActionSrc = escapeHtmlAttr(rawActionSrc);
        const safeDisplaySrc = escapeHtmlAttr(resolvedSrc || "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///yH5BAEAAAAALAAAAAABAAEAAAIBRAA7");

        return `
                    <div class="attached-img-item relative group rounded-lg overflow-hidden border border-gray-300 dark:border-neutral-600 bg-white shadow-sm p-1" data-attached-img="${safeActionSrc}" ${driveFileId ? `data-drive-file-id="${escapeHtmlAttr(driveFileId)}"` : ''}>
                        ${isUnresolvedDrive ? `
                            <div class="drive-img-loader h-20 w-28 flex flex-col items-center justify-center gap-1 text-[10px] text-blue-600 dark:text-blue-400 bg-gray-50 dark:bg-neutral-800 rounded px-2 text-center">
                                <i class="fa-solid fa-circle-notch fa-spin text-sm"></i>
                                <span>Loading HD...</span>
                            </div>
                        ` : ''}
                        <img src="${safeDisplaySrc}" ${isUnresolvedDrive ? `data-drive-img-id="${escapeHtmlAttr(driveFileId)}"` : ''} alt="Table Attachment ${idx + 1}" class="${isUnresolvedDrive ? 'hidden ' : ''}h-20 w-auto max-w-[180px] object-contain cursor-zoom-in block rounded" data-img-action="lightbox" data-img-src="${safeActionSrc}" title="Click to enlarge (Full Resolution)">
                        <span class="img-hd-badge absolute bottom-1 left-1 bg-emerald-600/90 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow pointer-events-none ${driveFileId ? '' : 'hidden'}">
                            <i class="fa-solid fa-check"></i> Original HD
                        </span>
                        <div class="absolute top-1 right-1 flex items-center gap-1 opacity-90 group-hover:opacity-100 transition">
                            <button type="button" data-img-action="copy" data-img-src="${safeActionSrc}" class="bg-slate-900/80 hover:bg-emerald-600 text-white rounded-full w-5 h-5 flex items-center justify-center text-[10px] shadow" title="Copy image to clipboard">
                                <i class="fa-solid fa-copy"></i>
                            </button>
                            <button type="button" data-ev-click="a148" class="bg-red-600 hover:bg-red-500 text-white rounded-full w-5 h-5 flex items-center justify-center text-[10px] shadow" title="Remove this image">
                                <i class="fa-solid fa-xmark"></i>
                            </button>
                        </div>
                    </div>
                `;
    }).join("");
    if (needsHydration) scheduleHydrateDriveImages();
    return html;
}

async function appendImageFileToWrapper(file, wrapperEl) {
    if (!file || !wrapperEl) return;
    const container = wrapperEl.querySelector('.field-images-container');
    if (!container) return;

    try {
        // Read 100% original image file (Zero Compression, full resolution)
        const fullDataUrl = await readFileAsDataUrl(file);
        container.insertAdjacentHTML('beforeend', renderImageAttachmentItemsHTML([fullDataUrl]));

        const itemEl = container.lastElementChild;
        if (!itemEl) return;

        const badgeEl = itemEl.querySelector('.img-hd-badge');
        if (badgeEl) {
            badgeEl.className = "img-hd-badge absolute bottom-1 left-1 bg-blue-600/90 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow pointer-events-none";
            badgeEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Uploading HD...`;
        }

        // Upload original image to Google Drive in background so it's ready when Save is clicked
        const uploadPromise = uploadTableImageToDrive(file, fullDataUrl)
            .then(fileId => {
                itemEl.setAttribute('data-drive-file-id', fileId);
                itemEl.setAttribute('data-attached-img', `drive:${fileId}`);
                if (badgeEl) {
                    badgeEl.className = "img-hd-badge absolute bottom-1 left-1 bg-emerald-600/90 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow pointer-events-none";
                    badgeEl.innerHTML = `<i class="fa-solid fa-check"></i> Original HD`;
                }
                return fileId;
            })
            .catch(err => {
                console.warn("Failed to upload HD image to Google Drive, using fallback:", err);
                if (badgeEl) {
                    badgeEl.className = "img-hd-badge absolute bottom-1 left-1 bg-amber-600/90 text-white text-[9px] font-bold px-1.5 py-0.5 rounded shadow pointer-events-none";
                    badgeEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation"></i> Local`;
                }
                return "";
            });

        pendingAttachmentUploads.set(itemEl, uploadPromise);
    } catch (err) {
        console.error("Failed to process attached image:", err);
        showToast("Failed to read image file. Please make sure the image format is valid (PNG/JPG/WebP).");
    }
}

async function handleFieldImageUpload(fileInputEl) {
    if (!fileInputEl || !fileInputEl.files || fileInputEl.files.length === 0) return;
    const wrapper = fileInputEl.closest('.field-with-images-wrapper');
    const files = Array.from(fileInputEl.files);
    for (const f of files) {
        await appendImageFileToWrapper(f, wrapper);
    }
    fileInputEl.value = "";
}

async function handleFieldImagePaste(event, textareaEl) {
    if (!event || !event.clipboardData || !event.clipboardData.items) return;
    const items = Array.from(event.clipboardData.items);
    const imageItems = items.filter(item => item.type && item.type.startsWith("image/"));
    if (imageItems.length === 0) return;

    event.preventDefault();
    const wrapper = textareaEl.closest('.field-with-images-wrapper');
    for (const item of imageItems) {
        const blob = item.getAsFile();
        if (blob) {
            await appendImageFileToWrapper(blob, wrapper);
        }
    }
}

async function collectFieldValueWithImages(wrapperEl, textareaEl) {
    const rawText = textareaEl ? textareaEl.value.trim() : "";
    const cleanText = (rawText && rawText !== "-") ? rawText : "";
    if (!wrapperEl) return sanitizeValue(cleanText);

    const imgItems = Array.from(wrapperEl.querySelectorAll('.attached-img-item'));
    if (imgItems.length === 0) return sanitizeValue(cleanText);

    const tokens = [];
    for (const itemEl of imgItems) {
        // Tunggu jika gambar sedang dalam proses upload ke Google Drive
        const pendingUpload = pendingAttachmentUploads.get(itemEl);
        if (pendingUpload) {
            await pendingUpload;
        }

        const driveFileId = itemEl.getAttribute('data-drive-file-id');
        if (driveFileId) {
            tokens.push(`[IMG_ID:${driveFileId}]`);
            continue;
        }

        const rawSrc = itemEl.getAttribute('data-attached-img') || "";
        if (!rawSrc) continue;

        if (rawSrc.startsWith("drive:")) {
            tokens.push(`[IMG_ID:${rawSrc.slice(6)}]`);
        } else {
            // Fallback jika upload Drive gagal / gambar lama
            const safeSrc = await ensureImageSrcFitsCell(rawSrc, 45000);
            tokens.push(`[IMG:${safeSrc}]`);
        }
    }

    if (tokens.length === 0) return sanitizeValue(cleanText);
    const combined = cleanText ? `${cleanText}\n${tokens.join("\n")}` : tokens.join("\n");
    return sanitizeValue(combined);
}

// Add a new benefit row ("Benefit Name" & "Benefit Description")
function addNewBenefitRow(targetContainerId = "newBenefitsContainer", emptyNoticeId = "emptyBenefitNotice") {
    const emptyNotice = document.getElementById(emptyNoticeId);
    if (emptyNotice) emptyNotice.remove();

    const container = document.getElementById(targetContainerId);
    if (!container) return;
    const scopeRoot = container.closest('#manualFormBody, #detailBody') || document;
    const datalistId = scopeRoot.id === 'manualFormBody' ? 'manualBenefitNamesList' : 'existingBenefitNamesList';

    const rowDiv = document.createElement("div");
    rowDiv.className = "new-benefit-row field-with-images-wrapper p-3 bg-blue-50/50 dark:bg-neutral-800/80 border border-blue-200 dark:border-neutral-600 rounded-xl space-y-2 relative transition";

    rowDiv.innerHTML = `
                <div class="flex justify-between items-center">
                    <span class="text-[11px] font-bold uppercase tracking-wider text-blue-700 dark:text-blue-400">
                        <i class="fa-solid fa-sparkles me-1"></i> Add Benefit
                    </span>
                    <div class="flex items-center gap-1.5">
                        <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Paste screenshot / image from clipboard (or press Ctrl+V)">
                            <i class="fa-solid fa-paste"></i> Paste
                        </button>
                        <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-blue-700 dark:text-blue-300 bg-blue-100 hover:bg-blue-200 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 px-2 py-0.5 rounded-md transition" title="Upload image file from computer">
                            <i class="fa-solid fa-upload"></i> Upload
                            <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                        </label>
                        <button type="button" data-ev-click="a137" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded" title="Cancel adding this benefit">
                            <i class="fa-solid fa-xmark"></i> Cancel
                        </button>
                    </div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Benefit Name</label>
                    <input type="text" list="${datalistId}" data-ev-input="a149" 
                           placeholder="Type or select benefit name (e.g. Accidental Death Benefit)..."
                           class="new-benefit-name w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-blue-500 outline-none">
                    <div class="benefit-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Benefit Description</label>
                    <textarea rows="3" data-ev-paste="a127" placeholder="Enter benefit description or paste (Ctrl+V) table image..."
                              class="new-benefit-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-blue-500 outline-none"></textarea>
                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden"></div>
                </div>
            `;

    container.appendChild(rowDiv);
    const nameInput = rowDiv.querySelector('.new-benefit-name');
    if (nameInput) nameInput.focus();
    refreshBenefitDatalist(scopeRoot);
}

// Real-time validation to check if Benefit Name already exists or is new
function validateBenefitName(inputEl) {
    const rawVal = inputEl.value.trim();
    const lowerVal = rawVal.toLowerCase();
    const feedbackEl = inputEl.parentElement.querySelector('.benefit-name-feedback');
    const scopeRoot = inputEl.closest('#manualFormBody, #detailBody') || document;

    if (!rawVal) {
        feedbackEl.className = "benefit-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
        feedbackEl.replaceChildren();
        inputEl.dataset.status = "empty";
        refreshBenefitDatalist(scopeRoot);
        return;
    }

    // 0. Check if typed name is a Product Provision (Non-Benefit) column
    if (isNonBenefitColumn(rawVal)) {
        feedbackEl.className = "benefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtml(rawVal)}"</strong> belongs to Product Provisions &amp; Specifications above, not Benefits.`;
        inputEl.dataset.status = "duplicate";
        refreshBenefitDatalist(scopeRoot);
        return;
    }

    // 0b. Check if typed name is a DPLK Charge column
    if (isDplkChargeColumn(rawVal)) {
        feedbackEl.className = "benefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-receipt me-1"></i> <strong>"${escapeHtml(rawVal)}"</strong> is a Charge component. Please use the <strong>DPLK Charge Components (Charges)</strong> section below.`;
        inputEl.dataset.status = "duplicate";
        refreshBenefitDatalist(scopeRoot);
        return;
    }

    // 1. Check if already used/checked in this product form (duplicate)
    const usedInProduct = getUsedBenefitNamesLower(inputEl, scopeRoot);
    if (usedInProduct.includes(lowerVal)) {
        feedbackEl.className = "benefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Benefit <strong>"${escapeHtml(rawVal)}"</strong> is already added or checked in this form!`;
        inputEl.dataset.status = "duplicate";
        refreshBenefitDatalist(scopeRoot);
        return;
    }

    // 2. Check if benefit name already exists in database benefit columns
    const matchedExistingCol = getRegularBenefitColumns().find(col => col.toLowerCase() === lowerVal);
    if (matchedExistingCol) {
        feedbackEl.className = "benefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> <strong>Registered Benefit:</strong> Already exists in database columns (<em>${escapeHtml(matchedExistingCol)}</em>).`;
        inputEl.dataset.status = "existing";
        inputEl.dataset.canonicalName = matchedExistingCol;
    } else {
        // 3. Brand-new benefit name
        feedbackEl.className = "benefit-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-wand-magic-sparkles me-1"></i> <strong>New Benefit:</strong> Not yet in the database (will be added after Column K).`;
        inputEl.dataset.status = "new";
        inputEl.dataset.canonicalName = rawVal;
    }

    refreshBenefitDatalist(scopeRoot);
}

// Remove benefit row
function removeBenefitRow(btnEl) {
    const scopeRoot = btnEl.closest('#manualFormBody, #detailBody') || document;
    const row = btnEl.closest('.existing-benefit-row, .new-benefit-row');
    if (row) row.remove();

    scopeRoot.querySelectorAll('.new-benefit-name').forEach(input => validateBenefitName(input));
    refreshBenefitDatalist(scopeRoot);
}

// Add a new charge row ("Charge Name" & "Charge Details") for DPLK
function addNewChargeRow(targetContainerId = "editNewChargesContainer", emptyNoticeId = "editEmptyChargeNotice") {
    const emptyNotice = document.getElementById(emptyNoticeId);
    if (emptyNotice) emptyNotice.remove();

    const container = document.getElementById(targetContainerId);
    if (!container) return;
    const scopeRoot = container.closest('#manualFormBody, #detailBody') || document;
    const datalistId = scopeRoot.id === 'manualFormBody' ? 'manualChargeNamesList' : 'existingChargeNamesList';

    const rowDiv = document.createElement("div");
    rowDiv.className = "new-charge-row field-with-images-wrapper p-3 bg-amber-50/60 dark:bg-amber-950/30 border border-amber-300 dark:border-amber-800 rounded-xl space-y-2 relative transition";

    rowDiv.innerHTML = `
                <div class="flex justify-between items-center">
                    <span class="text-[11px] font-bold uppercase tracking-wider text-amber-800 dark:text-amber-300">
                        <i class="fa-solid fa-receipt me-1"></i> Add Charge Component
                    </span>
                    <div class="flex items-center gap-1.5">
                        <button type="button" data-ev-click="a123" class="inline-flex items-center gap-1 text-[11px] font-semibold text-amber-800 dark:text-amber-300 bg-amber-100 hover:bg-amber-200 dark:bg-amber-900/50 dark:hover:bg-amber-800/60 px-2 py-0.5 rounded-md transition" title="Paste screenshot / table image">
                            <i class="fa-solid fa-paste"></i> Paste
                        </button>
                        <label class="cursor-pointer inline-flex items-center gap-1 text-[11px] font-semibold text-amber-800 dark:text-amber-300 bg-amber-100 hover:bg-amber-200 dark:bg-amber-900/50 dark:hover:bg-amber-800/60 px-2 py-0.5 rounded-md transition" title="Upload image file">
                            <i class="fa-solid fa-upload"></i> Upload
                            <input type="file" accept="image/*" multiple class="hidden" data-ev-change="a124">
                        </label>
                        <button type="button" data-ev-click="a135" class="text-gray-400 hover:text-red-500 dark:hover:text-red-400 transition text-xs px-1.5 py-0.5 rounded" title="Cancel adding this charge">
                            <i class="fa-solid fa-xmark"></i> Cancel
                        </button>
                    </div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Charge Name</label>
                    <input type="text" list="${datalistId}" data-ev-input="a150" 
                           placeholder="Select or type charge name (e.g. Biaya Administrasi Kepesertaan)..."
                           class="new-charge-name w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-amber-500 outline-none">
                    <div class="charge-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                </div>

                <div>
                    <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Charge Details / Amount</label>
                    <textarea rows="2" data-ev-paste="a127" placeholder="e.g. Rp 2,000 / month or Max 1% per year of managed funds..."
                              class="new-charge-desc w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs focus:ring-2 focus:ring-amber-500 outline-none"></textarea>
                    <div class="field-images-container flex flex-wrap gap-2 mt-2 empty:hidden"></div>
                </div>
            `;

    container.appendChild(rowDiv);
    const nameInput = rowDiv.querySelector('.new-charge-name');
    if (nameInput) nameInput.focus();
    refreshChargeDatalist(scopeRoot);
}

function validateChargeName(inputEl) {
    const rawVal = inputEl.value.trim();
    const lowerVal = rawVal.toLowerCase();
    const feedbackEl = inputEl.parentElement.querySelector('.charge-name-feedback');
    const scopeRoot = inputEl.closest('#manualFormBody, #detailBody') || document;

    if (!rawVal) {
        feedbackEl.className = "charge-name-feedback hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
        feedbackEl.replaceChildren();
        inputEl.dataset.status = "empty";
        refreshChargeDatalist(scopeRoot);
        return;
    }

    if (isNonBenefitColumn(rawVal)) {
        feedbackEl.className = "charge-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtml(rawVal)}"</strong> is a Product Provision &amp; Specification column, not a Charge.`;
        inputEl.dataset.status = "duplicate";
        refreshChargeDatalist(scopeRoot);
        return;
    }

    if (isDplkBenefitColumn(rawVal)) {
        feedbackEl.className = "charge-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> <strong>"${escapeHtml(rawVal)}"</strong> is a DPLK Pension Benefit. Please check it in the Benefit Checklist section above.`;
        inputEl.dataset.status = "duplicate";
        refreshChargeDatalist(scopeRoot);
        return;
    }

    const usedCharges = getUsedChargeNamesLower(inputEl, scopeRoot);
    if (usedCharges.includes(lowerVal)) {
        feedbackEl.className = "charge-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Charge component <strong>"${escapeHtml(rawVal)}"</strong> is already added in this form!`;
        inputEl.dataset.status = "duplicate";
        refreshChargeDatalist(scopeRoot);
        return;
    }

    const allKnownCharges = [...new Set([...DPLK_CHARGE_COLUMNS, ...getChargeColumns()])];
    const matchedCharge = allKnownCharges.find(col => col.toLowerCase() === lowerVal);
    if (matchedCharge) {
        feedbackEl.className = "charge-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> <strong>Registered DPLK Charge:</strong> Matches standard fee component (<em>${escapeHtml(matchedCharge)}</em>).`;
        inputEl.dataset.status = "existing";
        inputEl.dataset.canonicalName = matchedCharge;
    } else {
        const canonicalNewCharge = /^biaya\b/i.test(rawVal) ? rawVal : `Biaya ${rawVal}`;
        feedbackEl.className = "charge-name-feedback block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-wand-magic-sparkles me-1"></i> <strong>New Charge Component:</strong> Will be saved as <em>"${escapeHtml(canonicalNewCharge)}"</em> in the Charges section.`;
        inputEl.dataset.status = "new";
        inputEl.dataset.canonicalName = canonicalNewCharge;
    }

    refreshChargeDatalist(scopeRoot);
}

function removeChargeRow(btnEl) {
    const scopeRoot = btnEl.closest('#manualFormBody, #detailBody') || document;
    const row = btnEl.closest('.existing-charge-row, .new-charge-row');
    if (row) row.remove();

    scopeRoot.querySelectorAll('.new-charge-name').forEach(input => validateChargeName(input));
    refreshChargeDatalist(scopeRoot);
}

// Real-time validation for Product Name & Category in Edit Product form
function validateEditProductIdentity() {
    if (!currentActiveProduct) return { validName: false, validCategory: false };

    const p = currentActiveProduct;
    const nameInput = document.getElementById("edit_productName");
    const catInput = document.getElementById("edit_category");
    const nameFeedback = document.getElementById("editProductNameFeedback");
    const catFeedback = document.getElementById("editCategoryFeedback");

    let validName = true;
    let validCategory = true;

    // 1. Validate Category / Sheet Tab & dynamically re-render sections if category changed
    if (catInput && catFeedback) {
        const rawCat = catInput.value.trim();
        const matchedTab = SHEET_TABS.find(t => t.toLowerCase() === rawCat.toLowerCase());
        const detailBodyEl = document.getElementById("detailBody");
        const dynWrap = document.getElementById("editDynamicCategorySections");

        if (!rawCat || !matchedTab) {
            catFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
            catFeedback.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Invalid category! Choose: ${SHEET_TABS.map(t => `<strong>${escapeHtmlAttr(t)}</strong>`).join(", ")}.`;
            validCategory = false;
        } else {
            if (dynWrap && dynWrap.dataset.renderedCat !== matchedTab && detailBodyEl) {
                const preserved = captureScopeFormValues(detailBodyEl);
                dynWrap.innerHTML = renderCategoryDynamicFormSections('edit', matchedTab, p, preserved);
                dynWrap.dataset.renderedCat = matchedTab;
                refreshBenefitDatalist(detailBodyEl);
                refreshChargeDatalist(detailBodyEl);
                refreshNonBenefitDatalist(detailBodyEl);
                refreshSectionItemDatalist(detailBodyEl);
            }

            const catCfg = getCategorySchemaConfig(matchedTab);
            const specialCount = (catCfg.specialProvisions || []).length;
            const checklistSecs = (catCfg.productSections || []).filter(s => s.mode === "checklist").map(s => s.title);

            if (matchedTab !== p.category) {
                catFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800 font-medium";
                catFeedback.innerHTML = `<i class="fa-solid fa-arrow-right-arrow-left me-1"></i> Category Change: <strong>${escapeHtmlAttr(p.category)}</strong> &rarr; <strong>${escapeHtmlAttr(matchedTab)}</strong> (will be moved to sheet <strong>${escapeHtmlAttr(matchedTab)}</strong>).`;
            } else if (specialCount > 0 || checklistSecs.length > 0) {
                const summaryParts = [];
                if (specialCount > 0) summaryParts.push(`${specialCount} Special Provision(s)`);
                if (checklistSecs.length > 0) summaryParts.push(`Checklist: ${checklistSecs.join(", ")}`);
                catFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
                catFeedback.innerHTML = `<i class="fa-solid fa-layer-group me-1"></i> <strong>${escapeHtmlAttr(matchedTab)}</strong> template active (${escapeHtmlAttr(summaryParts.join(" · "))}).`;
            } else {
                catFeedback.className = "hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
                catFeedback.replaceChildren();
            }
        }
    }

    // 2. Validate Product Name
    if (nameInput && nameFeedback) {
        const rawName = nameInput.value.trim();
        const selectedCat = catInput ? catInput.value.trim() : p.category;

        if (!rawName || rawName === "-") {
            nameFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
            nameFeedback.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Product name cannot be empty!`;
            validName = false;
        } else {
            const duplicateProd = products.find(
                other => other.id !== p.id && other.name.toLowerCase().trim() === rawName.toLowerCase()
            );

            if (duplicateProd) {
                nameFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
                nameFeedback.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Product <strong>"${escapeHtml(duplicateProd.name)}"</strong> is already registered under category <strong>${escapeHtml(duplicateProd.category)}</strong>!`;
                validName = false;
            } else if (rawName !== p.name) {
                nameFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
                nameFeedback.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> New product name is available (<strong>"${escapeHtml(p.name)}"</strong> &rarr; <strong>"${escapeHtml(rawName)}"</strong>).`;
            } else if (selectedCat !== p.category) {
                nameFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
                nameFeedback.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> Product ready to be moved to category <strong>${escapeHtml(selectedCat)}</strong>.`;
            } else {
                nameFeedback.className = "hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
                nameFeedback.replaceChildren();
            }
        }
    }

    return { validName, validCategory };
}

function cancelEditMode() {
    if (currentActiveProduct) {
        openDetailModal(currentActiveProduct.id);
    }
}

async function collectAllDynamicFormData(rootEl, targetCategory) {
    const catCfg = getCategorySchemaConfig(targetCategory);
    const hasSpecial = (colName) => (catCfg.specialProvisions || []).some(sp => sp.name.toLowerCase() === colName.toLowerCase());
    const isShariaMode = /^(sharia|syariah)$/i.test(targetCategory || "") || hasSpecial("Aqd") || hasSpecial("Alokasi Kontribusi") || hasSpecial("Surplus Underwriting");
    const isDplkMode = /^dplk$/i.test(targetCategory || "") ||
        (catCfg.generalProvisions || []).some(g => /^(sales tools|after sales tools)$/i.test(g)) ||
        (catCfg.productSections || []).some(s => s.mode === "checklist" || s.isChargeLike);

    ensureNonBenefitColumnsInDynamicSchema(isShariaMode, isDplkMode);

    (catCfg.generalProvisions || []).forEach(g => {
        if (g && !isNonBenefitColumn(g)) NON_BENEFIT_COLUMNS.push(g);
        if (g && !dynamicColumns.some(c => c.toLowerCase() === g.toLowerCase())) dynamicColumns.push(g);
    });
    (catCfg.specialProvisions || []).forEach(sp => {
        if (sp.name && !isNonBenefitColumn(sp.name)) NON_BENEFIT_COLUMNS.push(sp.name);
        if (sp.name && !dynamicColumns.some(c => c.toLowerCase() === sp.name.toLowerCase())) {
            const entryAgeIdx = dynamicColumns.findIndex(c => c.toLowerCase() === "entry age");
            if (entryAgeIdx !== -1) dynamicColumns.splice(entryAgeIdx, 0, sp.name);
            else dynamicColumns.push(sp.name);
        }
    });

    let updatedAttr = {};
    dynamicColumns.forEach(col => {
        updatedAttr[col] = "-";
    });

    const activeSpecialLower = new Set((catCfg.specialProvisions || []).map(sp => sp.name.toLowerCase()));
    if (activeSpecialLower.has("aqd")) activeSpecialLower.add("akad");

    // 1. General Provisions, Special Provisions & Additional Benefit (.spec-field-input)
    const specInputs = Array.from(rootEl.querySelectorAll('.spec-field-input'));
    for (const input of specInputs) {
        const targetSpecCol = input.getAttribute('data-spec-col');
        if (!targetSpecCol) continue;
        const lowerCol = targetSpecCol.toLowerCase();
        if (isSpecialProvisionColumnAnyCategory(targetSpecCol) && !activeSpecialLower.has(lowerCol)) {
            continue;
        }
        const wrapper = input.closest('.field-with-images-wrapper');
        const valWithImgs = await collectFieldValueWithImages(wrapper, input);
        const actualCol = dynamicColumns.find(k => k.toLowerCase() === lowerCol) || targetSpecCol;
        if (!dynamicColumns.includes(actualCol)) {
            if (lowerCol === "additional benefit" || activeSpecialLower.has(lowerCol)) {
                const entryAgeIdx = dynamicColumns.findIndex(c => c.toLowerCase() === "entry age");
                if (entryAgeIdx !== -1) dynamicColumns.splice(entryAgeIdx, 0, actualCol);
                else dynamicColumns.splice(Math.min(6, dynamicColumns.length), 0, actualCol);
            } else {
                dynamicColumns.push(actualCol);
            }
        }
        updatedAttr[actualCol] = valWithImgs;
    }

    // 1a. Existing & newly added Custom Non-Benefit Provisions
    const customNonBenefitsToRegister = [];
    for (const row of Array.from(rootEl.querySelectorAll('.existing-nonbenefit-row'))) {
        const colName = row.getAttribute('data-col-name');
        const descTextarea = row.querySelector('.existing-nonbenefit-desc');
        const valWithImgs = await collectFieldValueWithImages(row, descTextarea);
        if (colName) {
            const actualCol = dynamicColumns.find(k => k.toLowerCase() === colName.toLowerCase()) || colName;
            if (!isNonBenefitColumn(actualCol)) NON_BENEFIT_COLUMNS.push(actualCol);
            if (!dynamicColumns.includes(actualCol)) {
                const addBenIdx = dynamicColumns.findIndex(c => c.toLowerCase() === "additional benefit");
                if (addBenIdx !== -1) dynamicColumns.splice(addBenIdx, 0, actualCol);
                else dynamicColumns.push(actualCol);
            }
            updatedAttr[actualCol] = valWithImgs;
        }
    }

    for (const row of Array.from(rootEl.querySelectorAll('.new-nonbenefit-row'))) {
        const nbNameInput = row.querySelector('.new-nonbenefit-name');
        const nbDescInput = row.querySelector('.new-nonbenefit-desc');
        const rawName = (nbNameInput.dataset.canonicalName || nbNameInput.value).trim();
        const valWithImgs = await collectFieldValueWithImages(row, nbDescInput);

        if (rawName && !isEmptyVal(valWithImgs)) {
            const matchedCol = dynamicColumns.find(col => col.toLowerCase() === rawName.toLowerCase());
            const finalColName = matchedCol || rawName;
            if (!isNonBenefitColumn(finalColName)) NON_BENEFIT_COLUMNS.push(finalColName);
            if (!dynamicColumns.includes(finalColName)) {
                const addBenIdx = dynamicColumns.findIndex(c => c.toLowerCase() === "additional benefit");
                if (addBenIdx !== -1) dynamicColumns.splice(addBenIdx, 0, finalColName);
                else dynamicColumns.push(finalColName);
            }
            updatedAttr[finalColName] = valWithImgs;
            if (!NON_BENEFIT_CONFIG_CANONICALS.has(normalizeBenefitNameKey(finalColName))) {
                customNonBenefitsToRegister.push({ entryType: "Non-Benefit", canonical: finalColName, alias: "" });
            }
        }
    }

    // 1b. Checklist Items (untuk kategori mana pun yg memakai mode Checklist + Optional Note)
    const checkedBoxes = Array.from(rootEl.querySelectorAll('.dplk-benefit-chk:checked'));
    checkedBoxes.forEach(chk => {
        const itemEl = chk.closest('.dplk-benefit-item');
        const customNameInput = itemEl ? itemEl.querySelector('.dplk-custom-benefit-name') : null;
        const bCol = customNameInput ? customNameInput.value.trim() : chk.getAttribute('data-benefit-col');
        if (!bCol) return;
        const actualCol = dynamicColumns.find(c => c.toLowerCase() === bCol.toLowerCase()) || bCol;
        if (!dynamicColumns.includes(actualCol)) {
            const insertIdx = Math.min(5, dynamicColumns.length);
            dynamicColumns.splice(insertIdx, 0, actualCol);
        }
        const remInput = itemEl ? itemEl.querySelector('.dplk-benefit-remark') : null;
        const remVal = remInput ? remInput.value.trim() : "";
        updatedAttr[actualCol] = remVal ? `✓ | ${remVal}` : "✓";
    });

    // 2. Existing & newly added Primary Benefits
    for (const row of Array.from(rootEl.querySelectorAll('.existing-benefit-row'))) {
        const colName = row.getAttribute('data-col-name');
        const descTextarea = row.querySelector('.existing-benefit-desc');
        const valWithImgs = await collectFieldValueWithImages(row, descTextarea);
        if (colName) {
            const actualCol = dynamicColumns.find(k => k.toLowerCase() === colName.toLowerCase()) || colName;
            if (!dynamicColumns.includes(actualCol)) dynamicColumns.push(actualCol);
            updatedAttr[actualCol] = valWithImgs;
        }
    }

    for (const row of Array.from(rootEl.querySelectorAll('.new-benefit-row'))) {
        const bNameInput = row.querySelector('.new-benefit-name');
        const bDescInput = row.querySelector('.new-benefit-desc');
        const rawName = bNameInput.value.trim();
        const valWithImgs = await collectFieldValueWithImages(row, bDescInput);

        if (rawName && !isEmptyVal(valWithImgs)) {
            const matchedCol = dynamicColumns.find(col => col.toLowerCase() === rawName.toLowerCase());
            const finalColName = matchedCol || rawName;
            if (!dynamicColumns.includes(finalColName)) {
                const insertIdx = Math.min(5, dynamicColumns.length);
                dynamicColumns.splice(insertIdx, 0, finalColName);
            }
            updatedAttr[finalColName] = valWithImgs;
            if (!BENEFIT_CONFIG_CANONICALS.has(normalizeBenefitNameKey(finalColName))) {
                customNonBenefitsToRegister.push({ entryType: "Benefit", canonical: finalColName, alias: "" });
            }
        }
    }

    // 3. Existing & newly added Charges
    for (const row of Array.from(rootEl.querySelectorAll('.existing-charge-row'))) {
        const colName = row.getAttribute('data-col-name');
        const descTextarea = row.querySelector('.existing-charge-desc');
        const valWithImgs = await collectFieldValueWithImages(row, descTextarea);
        if (colName) {
            const actualCol = dynamicColumns.find(c => c.toLowerCase() === colName.toLowerCase()) || colName;
            if (!dynamicColumns.includes(actualCol)) dynamicColumns.push(actualCol);
            updatedAttr[actualCol] = valWithImgs;
        }
    }

    for (const row of Array.from(rootEl.querySelectorAll('.new-charge-row'))) {
        const cNameInput = row.querySelector('.new-charge-name');
        const cDescInput = row.querySelector('.new-charge-desc');
        const rawName = (cNameInput.dataset.canonicalName || cNameInput.value).trim();
        const valWithImgs = await collectFieldValueWithImages(row, cDescInput);

        if (rawName && !isEmptyVal(valWithImgs)) {
            const matchedCol = dynamicColumns.find(col => col.toLowerCase() === rawName.toLowerCase());
            const finalColName = matchedCol || rawName;
            if (!dynamicColumns.includes(finalColName)) dynamicColumns.push(finalColName);
            updatedAttr[finalColName] = valWithImgs;
        }
    }

    // 4. Existing & newly added Custom Product Section items
    for (const row of Array.from(rootEl.querySelectorAll('.existing-section-row'))) {
        const colName = row.getAttribute('data-col-name');
        const descTextarea = row.querySelector('.existing-section-item-desc');
        const valWithImgs = await collectFieldValueWithImages(row, descTextarea);
        if (colName) {
            const actualCol = dynamicColumns.find(c => c.toLowerCase() === colName.toLowerCase()) || colName;
            if (!dynamicColumns.includes(actualCol)) dynamicColumns.push(actualCol);
            updatedAttr[actualCol] = valWithImgs;
        }
    }

    for (const row of Array.from(rootEl.querySelectorAll('.new-section-row'))) {
        const sNameInput = row.querySelector('.new-section-item-name');
        const sDescInput = row.querySelector('.new-section-item-desc');
        const rawName = (sNameInput ? sNameInput.value : "").trim();
        const valWithImgs = await collectFieldValueWithImages(row, sDescInput);

        if (rawName && !isEmptyVal(valWithImgs)) {
            const matchedCol = dynamicColumns.find(col => col.toLowerCase() === rawName.toLowerCase());
            const finalColName = matchedCol || rawName;
            if (!dynamicColumns.includes(finalColName)) {
                dynamicColumns.push(finalColName);
            }
            updatedAttr[finalColName] = valWithImgs;
        }
    }

    return { updatedAttr, customNonBenefitsToRegister };
}

async function saveProductEdit() {
    if (!currentActiveProduct) return;

    const p = currentActiveProduct;
    const detailBodyEl = document.getElementById("detailBody");

    const nameInput = document.getElementById("edit_productName");
    const catInput = document.getElementById("edit_category");
    const compInput = document.getElementById("edit_company");

    // Validate Product Name & Category first
    const { validName, validCategory } = validateEditProductIdentity();
    if (!validCategory) {
        showToast(`Invalid product category! Please make sure the category is one of: ${SHEET_TABS.join(", ")}.`);
        if (catInput) catInput.focus();
        return;
    }
    if (!validName) {
        showToast("Product name is invalid or already used by another product. Please check again.");
        if (nameInput) nameInput.focus();
        return;
    }
    if (!compInput.value.trim()) {
        showToast("Please fill in 'Company Name' first.");
        compInput.focus();
        return;
    }

    const newProductName = nameInput.value.trim();
    const newCategory = SHEET_TABS.find(t => t.toLowerCase() === catInput.value.trim().toLowerCase()) || p.category;
    const isCategoryChanged = newCategory !== p.category;

    if (!validateAllNewRepeatableRows(detailBodyEl)) return;

    // Confirm if product category was changed (moves sheet tab)
    if (isCategoryChanged) {
        const confirmed = confirm(
            `Product category changed from "${p.category}" to "${newCategory}".\n\n` +
            `Product "${newProductName}" will be moved from sheet tab "${p.category}" to sheet tab "${newCategory}". Continue?`
        );
        if (!confirmed) return;
    }

    const btnSave = document.getElementById("btnSaveProductEdit");
    btnSave.disabled = true;
    btnSave.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving...`;

    // Collect primary form fields
    const newCompany = compInput.value.trim() || "-";
    const newType = document.getElementById("edit_productType").value.trim() || "-";
    const newChannel = document.getElementById("edit_channel").value.trim() || "-";
    const newCurr = document.getElementById("edit_currency").value.trim() || "-";
    const editLaunchEl = document.getElementById("edit_launchDate");
    const newLaunchDate = (editLaunchEl && editLaunchEl.value) ? formatDateToDDMMYYYY(editLaunchEl.value) : (p.launchDate || "-");

    const { updatedAttr, customNonBenefitsToRegister } = await collectAllDynamicFormData(detailBodyEl, newCategory);

    // Build complete row according to Master Schema
    let baseMasterSchema = [
        "Product Name", "Company Name", "Par/Non Par", "Product Type",
        "Distribution Channel", "Currency", "Launch Date", ...getColumnsForTargetSheet(newCategory, updatedAttr)
    ];

    let baseRowData = baseMasterSchema.map(colName => {
        if (colName === "Product Name") return newProductName;
        if (colName === "Company Name") return newCompany;
        if (colName === "Par/Non Par") return p.par || "-";
        if (colName === "Product Type") return newType;
        if (colName === "Distribution Channel") return newChannel;
        if (colName === "Currency") return newCurr;
        if (colName === "Launch Date") return newLaunchDate;
        return updatedAttr[colName] || "-";
    });

    await offloadOversizedCellImagesToDrive([baseRowData]);
    const { finalSchema: masterSchema, finalRows } = expandRowsWithContinuationColumns(baseMasterSchema, [baseRowData]);
    const rowData = finalRows[0];

    const oversizedEdit = findOversizedCells(masterSchema, rowData);
    if (oversizedEdit.length > 0) {
        showToast(`Failed to save: the following columns exceed Google Sheets' 50,000-character limit per cell:\n- ${oversizedEdit.join("\n- ")}\n\nPlease shorten the text or use smaller images.`);
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Changes`;
        return;
    }

    try {
        if (!isCategoryChanged) {
            const payload = {
                action: "update",
                sheetName: p.category,
                productName: p.name,
                headers: masterSchema,
                data: rowData
            };

            await postAuthenticatedAppsScript(payload);
        } else {
            await appendToSheetWithHeaderCheck(newCategory, masterSchema, [rowData]);

            const deletePayload = {
                action: "delete",
                items: [{
                    sheetName: p.category,
                    productName: p.name
                }]
            };

            await postAuthenticatedAppsScript(deletePayload);
        }

        if (customNonBenefitsToRegister.length > 0) {
            try {
                await postAuthenticatedAppsScript({ action: "save_benefit_aliases", mappings: customNonBenefitsToRegister });
                customNonBenefitsToRegister.forEach(m => {
                    if (!m || !m.canonical) return;
                    const cKey = normalizeBenefitNameKey(m.canonical);
                    if (String(m.entryType || "").toLowerCase() === "non-benefit") {
                        NON_BENEFIT_CONFIG_CANONICALS.add(cKey);
                    } else {
                        BENEFIT_CONFIG_CANONICALS.add(cKey);
                        if (!BENEFIT_CONFIG_CANONICAL_LIST.some(b => b.toLowerCase() === String(m.canonical).toLowerCase())) {
                            BENEFIT_CONFIG_CANONICAL_LIST.push(String(m.canonical).trim());
                        }
                    }
                });
                invalidateReviewLookupCaches();
            } catch (e) {
                if (e && e.code === "UNAUTHORIZED") throw e;
                console.warn("Failed to save custom benefit/non-benefit to Config:", e);
            }
        }

        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Changes`;

        closeDetailModal();
        invalidateLocalSummaryCache();
        fetchAllSheetData(true);

    } catch (err) {
        console.error(err);
        if (!err || err.code !== "UNAUTHORIZED") {
            showToast("Failed to update data to Google Sheets. Please check your internet connection.");
        }
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save Changes`;
    }
}

function closeDetailModal() {
    document.getElementById("detailModal").classList.add("hidden");
    currentActiveProduct = null;
}

// --- UPLOAD EXCEL & TAMBAHKAN MANUAL LOGIC (DENGAN TEMA DINAMIS KONVEN/SYARIAH/HS) ---
let currentAddModalTab = 'excel';

function getUploadModalTheme(category) {
    const c = String(category || "").trim().toLowerCase();
    if (c === "sharia" || c === "syariah") {
        return {
            headerBg: "bg-green-700",
            tabActiveText: "text-green-700",
            tabActiveTextDark: "dark:text-green-400",
            btnSubmitBg: "bg-green-600 hover:bg-green-500",
            fileInputClass: "w-full text-xs text-gray-500 dark:text-neutral-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-green-50 dark:file:bg-green-900/40 file:text-green-700 dark:file:text-green-300 hover:file:bg-green-100 dark:hover:file:bg-green-900/60 cursor-pointer border dark:border-neutral-600 rounded-lg p-1 transition",
            focusRing: "focus:ring-green-500",
            bannerClass: "bg-green-50 dark:bg-green-900/30 border border-green-200 dark:border-green-800 text-green-800 dark:text-green-200 font-semibold p-2.5 rounded-lg mb-2"
        };
    }
    if (c === "health & surgical" || c === "hs") {
        return {
            headerBg: "bg-purple-900",
            tabActiveText: "text-purple-800",
            tabActiveTextDark: "dark:text-purple-400",
            btnSubmitBg: "bg-purple-800 hover:bg-purple-700",
            fileInputClass: "w-full text-xs text-gray-500 dark:text-neutral-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-purple-100 dark:file:bg-purple-950/70 file:text-purple-900 dark:file:text-purple-300 hover:file:bg-purple-200 dark:hover:file:bg-purple-900/80 cursor-pointer border dark:border-neutral-600 rounded-lg p-1 transition",
            focusRing: "focus:ring-purple-800",
            bannerClass: "bg-purple-50 dark:bg-purple-900/30 border border-purple-200 dark:border-purple-800 text-purple-800 dark:text-purple-200 font-semibold p-2.5 rounded-lg mb-2"
        };
    }
    if (c === "dplk") {
        return {
            headerBg: "bg-amber-600",
            tabActiveText: "text-amber-700",
            tabActiveTextDark: "dark:text-amber-400",
            btnSubmitBg: "bg-amber-600 hover:bg-amber-500",
            fileInputClass: "w-full text-xs text-gray-500 dark:text-neutral-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-amber-50 dark:file:bg-amber-900/40 file:text-amber-700 dark:file:text-amber-300 hover:file:bg-amber-100 dark:hover:file:bg-amber-900/60 cursor-pointer border dark:border-neutral-600 rounded-lg p-1 transition",
            focusRing: "focus:ring-amber-500",
            bannerClass: "bg-amber-50 dark:bg-amber-900/30 border border-amber-200 dark:border-amber-800 text-amber-800 dark:text-amber-200 font-semibold p-2.5 rounded-lg mb-2"
        };
    }
    if (c && c !== "conventional" && c !== "konven") {
        return {
            headerBg: "bg-teal-700",
            tabActiveText: "text-teal-700",
            tabActiveTextDark: "dark:text-teal-400",
            btnSubmitBg: "bg-teal-600 hover:bg-teal-500",
            fileInputClass: "w-full text-xs text-gray-500 dark:text-neutral-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-teal-50 dark:file:bg-teal-900/40 file:text-teal-700 dark:file:text-teal-300 hover:file:bg-teal-100 dark:hover:file:bg-teal-900/60 cursor-pointer border dark:border-neutral-600 rounded-lg p-1 transition",
            focusRing: "focus:ring-teal-500",
            bannerClass: "bg-teal-50 dark:bg-teal-900/30 border border-teal-200 dark:border-teal-800 text-teal-800 dark:text-teal-200 font-semibold p-2.5 rounded-lg mb-2"
        };
    }
    // Konven (Default): Biru Muda (Sky)
    return {
        headerBg: "bg-sky-600",
        tabActiveText: "text-sky-600",
        tabActiveTextDark: "dark:text-sky-400",
        btnSubmitBg: "bg-sky-600 hover:bg-sky-500",
        fileInputClass: "w-full text-xs text-gray-500 dark:text-neutral-300 file:mr-3 file:py-2 file:px-4 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-sky-50 dark:file:bg-sky-900/40 file:text-sky-700 dark:file:text-sky-300 hover:file:bg-sky-100 dark:hover:file:bg-sky-900/60 cursor-pointer border dark:border-neutral-600 rounded-lg p-1 transition",
        focusRing: "focus:ring-sky-500",
        bannerClass: "bg-sky-50 dark:bg-neutral-700/70 border border-sky-200 dark:border-neutral-600 text-sky-800 dark:text-sky-200 font-semibold p-2.5 rounded-lg mb-2"
    };
}

function updateUploadModalTheme(targetCat) {
    const selectEl = document.getElementById("uploadTargetSheet");
    const cat = targetCat || (selectEl ? selectEl.value : (SHEET_TABS[0] || "Conventional")) || (SHEET_TABS[0] || "Conventional");
    const theme = getUploadModalTheme(cat);

    // 1. Header Background
    const header = document.getElementById("uploadModalHeader");
    if (header) {
        header.classList.remove("bg-green-700", "bg-sky-600", "bg-purple-700", "bg-purple-900", "bg-amber-600", "bg-teal-700");
        header.classList.add(theme.headerBg);
    }

    // 2. Tab switcher buttons
    const btnExcel = document.getElementById("tabBtnExcel");
    const btnManual = document.getElementById("tabBtnManual");
    const allTabTextClasses = [
        "text-green-700", "dark:text-green-400",
        "text-sky-600", "dark:text-sky-400",
        "text-purple-600", "text-purple-800", "dark:text-purple-400",
        "text-amber-700", "dark:text-amber-400",
        "text-teal-700", "dark:text-teal-400"
    ];

    if (btnExcel && btnManual) {
        [btnExcel, btnManual].forEach(btn => btn.classList.remove(...allTabTextClasses));
        if (currentAddModalTab === 'excel') {
            btnExcel.classList.add(theme.tabActiveText, theme.tabActiveTextDark);
        } else {
            btnManual.classList.add(theme.tabActiveText, theme.tabActiveTextDark);
        }
    }

    // 3. Choose File input button
    const fileInput = document.getElementById("excelFileInput");
    if (fileInput) {
        fileInput.className = theme.fileInputClass;
    }

    // 4. Action buttons: Upload ke Google Sheets & Simpan Produk Baru
    const allBtnBgClasses = [
        "bg-green-600", "hover:bg-green-500",
        "bg-sky-600", "hover:bg-sky-500",
        "bg-purple-600", "hover:bg-purple-500", "bg-purple-800", "hover:bg-purple-700",
        "bg-amber-600", "hover:bg-amber-500",
        "bg-teal-600", "hover:bg-teal-500"
    ];
    const btnSubmit = document.getElementById("btnSubmitUpload");
    if (btnSubmit) {
        btnSubmit.classList.remove(...allBtnBgClasses);
        btnSubmit.classList.add(...theme.btnSubmitBg.split(' '));
    }
    const btnSaveManual = document.getElementById("btnSaveManualProduct");
    if (btnSaveManual) {
        btnSaveManual.classList.remove(...allBtnBgClasses);
        btnSaveManual.classList.add(...theme.btnSubmitBg.split(' '));
    }

    // 5. Manual form banner
    const manualBanner = document.getElementById("manualFormBanner");
    if (manualBanner) {
        manualBanner.className = theme.bannerClass;
    }

    // 6. Focus rings (termasuk seluruh input di form Tambahkan Manual)
    const allFocusRings = ["focus:ring-green-500", "focus:ring-sky-500", "focus:ring-purple-500", "focus:ring-purple-800", "focus:ring-amber-500", "focus:ring-teal-500"];
    const manualBody = document.getElementById("manualFormBody");
    if (manualBody) {
        manualBody.querySelectorAll("input, select, textarea").forEach(el => {
            if (el.classList.contains("syariah-spec-input") || el.classList.contains("dplk-spec-input")) return;
            if (allFocusRings.some(cls => el.classList.contains(cls))) {
                el.classList.remove(...allFocusRings);
                el.classList.add(theme.focusRing);
            }
        });
    }
    if (selectEl) {
        selectEl.classList.remove(...allFocusRings);
        selectEl.classList.add(theme.focusRing);
    }
    const manualCatSelect = document.getElementById("manual_category");
    if (manualCatSelect) {
        manualCatSelect.classList.remove(...allFocusRings);
        manualCatSelect.classList.add(theme.focusRing);
    }
}

function onUploadTargetSheetChange() {
    const selectEl = document.getElementById("uploadTargetSheet");
    const val = selectEl ? selectEl.value : "Konven";
    updateUploadModalTheme(val);

    // Sinkronkan ke tab manual
    const manualCat = document.getElementById("manual_category");
    if (manualCat && manualCat.value !== val) {
        manualCat.value = val;
        validateManualCategoryAndToggleSyariah();
    }

    // Re-render pratinjau duplikasi jika file excel sudah ada
    const fileInput = document.getElementById("excelFileInput");
    if (uploadWorkbookCache) {
        renderUploadSheetSelector();
        analyzeSelectedUploadSheets();
    } else if (fileInput && fileInput.files && fileInput.files.length > 0) {
        previewExcelFile();
    }
}

let uploadModalSessionId = 0;

function cleanupUploadModalState() {
    uploadModalSessionId++;
    if (uploadWorkbookCache) {
        if (Array.isArray(uploadWorkbookCache.baseColumns)) {
            dynamicColumns = uploadWorkbookCache.baseColumns.slice();
        }
        if (Array.isArray(uploadWorkbookCache.baseNonBenefits)) {
            NON_BENEFIT_COLUMNS = uploadWorkbookCache.baseNonBenefits.slice();
        }
    }
    const fileIn = document.getElementById("excelFileInput");
    if (fileIn) fileIn.value = "";
    const previewEl = document.getElementById("uploadPreviewArea");
    if (previewEl) {
        previewEl.classList.add("hidden");
        previewEl.replaceChildren();
    }
    resetUploadSheetSelector();
    resetBenefitReviewState();
    uploadSubmitBaseEnabled = false;
    parsedUploadRows = [];
    parsedDuplicateRows = [];
    const submitBtn = document.getElementById("btnSubmitUpload");
    if (submitBtn) {
        submitBtn.disabled = true;
        submitBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Upload to Google Sheets`;
    }
}

function openUploadModal() {
    loadExcelLibs().catch(() => { });
    cleanupUploadModalState();
    document.getElementById("uploadModal").classList.remove("hidden");

    renderManualProductForm();
    switchAddModalTab('excel');
    updateUploadModalTheme();
}

function closeUploadModal() {
    document.getElementById("uploadModal").classList.add("hidden");
    cleanupUploadModalState();
}

function switchAddModalTab(mode) {
    currentAddModalTab = mode;
    const tabExcel = document.getElementById("uploadTabExcel");
    const tabManual = document.getElementById("uploadTabManual");
    const btnExcel = document.getElementById("tabBtnExcel");
    const btnManual = document.getElementById("tabBtnManual");
    const selectEl = document.getElementById("uploadTargetSheet");
    const theme = getUploadModalTheme(selectEl ? selectEl.value : "Konven");

    const activeBase = ["bg-white", "dark:bg-neutral-800", "shadow-sm"];
    const inactiveClasses = ["text-gray-600", "dark:text-neutral-400", "hover:text-gray-900", "dark:hover:text-neutral-200"];
    const allTabTextClasses = [
        "text-green-700", "dark:text-green-400",
        "text-sky-600", "dark:text-sky-400",
        "text-purple-600", "text-purple-800", "dark:text-purple-400",
        "text-amber-700", "dark:text-amber-400",
        "text-teal-700", "dark:text-teal-400"
    ];

    if (mode === 'manual') {
        tabExcel.classList.add("hidden");
        tabManual.classList.remove("hidden");
        btnManual.classList.remove(...inactiveClasses, ...allTabTextClasses);
        btnManual.classList.add(...activeBase, theme.tabActiveText, theme.tabActiveTextDark);

        btnExcel.classList.remove(...activeBase, ...allTabTextClasses);
        btnExcel.classList.add(...inactiveClasses);

        const prodInput = document.getElementById("manual_productName");
        if (prodInput) prodInput.focus();
    } else {
        tabManual.classList.add("hidden");
        tabExcel.classList.remove("hidden");
        btnExcel.classList.remove(...inactiveClasses, ...allTabTextClasses);
        btnExcel.classList.add(...activeBase, theme.tabActiveText, theme.tabActiveTextDark);

        btnManual.classList.remove(...activeBase, ...allTabTextClasses);
        btnManual.classList.add(...inactiveClasses);
    }
}

function renderManualProductForm() {
    const container = document.getElementById("manualFormBody");
    if (!container) return;

    const selectEl = document.getElementById("uploadTargetSheet");
    const currentCat = (selectEl ? selectEl.value : (SHEET_TABS[0] || "Conventional")) || (SHEET_TABS[0] || "Conventional");
    const theme = getUploadModalTheme(currentCat);

    const btnSave = document.getElementById("btnSaveManualProduct");
    if (btnSave) {
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save New Product`;
        btnSave.className = `px-4 py-2 ${theme.btnSubmitBg} text-white rounded-lg text-xs font-bold transition flex items-center gap-1.5`;
    }

    container.innerHTML = `
                <div id="manualFormBanner" class="${theme.bannerClass}">
                    <i class="fa-solid fa-file-circle-plus me-1"></i> Add New Product Manually to Google Sheets
                </div>

                <!-- Target Category & Product Name -->
                <div class="grid grid-cols-1 sm:grid-cols-3 gap-2.5">
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Category / Sheet Tab <span class="text-red-500">*</span></label>
                        <div class="flex items-center gap-1.5">
                            <select id="manual_category" data-ev-change="a151" class="flex-1 min-w-0 border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                                ${SHEET_TABS.map(tab => `<option value="${escapeHtmlAttr(tab)}">${escapeHtmlAttr(tab)}</option>`).join('')}
                            </select>
                            <button type="button" data-ev-click="a152" class="px-2 py-2 bg-gray-100 hover:bg-gray-200 dark:bg-neutral-700 dark:hover:bg-neutral-600 border dark:border-neutral-600 text-gray-700 dark:text-neutral-200 rounded-lg text-xs transition shrink-0" title="Edit this category's provisions & sections">
                                <i class="fa-solid fa-gear"></i>
                            </button>
                            <button type="button" data-ev-click="a46" class="px-2 py-2 bg-blue-50 hover:bg-blue-100 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 border border-blue-200 dark:border-neutral-600 text-blue-700 dark:text-blue-300 rounded-lg text-xs transition shrink-0" title="Create a new category sheet">
                                <i class="fa-solid fa-folder-plus"></i>
                            </button>
                        </div>
                        <div id="manualCategoryFeedback" class="hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                    </div>
                    <div class="sm:col-span-2">
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Product Name <span class="text-red-500">*</span></label>
                        <input type="text" id="manual_productName" data-ev-input="a153" placeholder="e.g. DPLK AstraLife - PPIP Individu / MiFuture Income Protector..." class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                        <div id="manualProductNameFeedback" class="hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border"></div>
                    </div>
                </div>

                <!-- SECTION 1: BASIC INFORMATION -->
                <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Company Name <span class="text-red-500">*</span></label>
                        <input type="text" id="manual_company" placeholder="e.g. DPLK AstraLife / Manulife / Allianz..." class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Currency</label>
                        <input type="text" id="manual_currency" placeholder="e.g. IDR / USD" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Launch Date</label>
                        <input type="date" id="manual_launchDate" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                    </div>
                </div>

                <div class="grid grid-cols-1 sm:grid-cols-2 gap-2.5">
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Product Type</label>
                        <input type="text" id="manual_productType" placeholder="e.g. PPIP / DKP / Endowment / Traditional" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                    </div>
                    <div>
                        <label class="block font-bold text-gray-700 dark:text-neutral-300 mb-1">Distribution Channel</label>
                        <input type="text" id="manual_channel" placeholder="e.g. Direct / Corporate / Agency / Bancassurance" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 focus:ring-2 focus:ring-green-500 outline-none">
                    </div>
                </div>

                <div id="manualDynamicCategorySections" data-rendered-cat="${escapeHtmlAttr(currentCat)}">
                    ${renderCategoryDynamicFormSections('manual', currentCat, null)}
                </div>
            `;

    const manualCatSel = document.getElementById("manual_category");
    if (manualCatSel) manualCatSel.value = currentCat;
    refreshBenefitDatalist(container);
    refreshChargeDatalist(container);
    refreshNonBenefitDatalist(container);
    refreshSectionItemDatalist(container);
    validateManualCategoryAndToggleSyariah();
}

function validateManualCategoryAndToggleSyariah() {
    const catSelect = document.getElementById("manual_category");
    const catFeedback = document.getElementById("manualCategoryFeedback");
    const manualFormEl = document.getElementById("manualFormBody");
    if (!catSelect) return true;

    const rawCat = catSelect.value.trim();
    const matchedTab = SHEET_TABS.find(t => t.toLowerCase() === rawCat.toLowerCase());

    if (!matchedTab) {
        if (catFeedback) {
            catFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
            catFeedback.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Invalid category! Choose: ${SHEET_TABS.map(t => escapeHtml(t)).join(", ")}.`;
        }
        return false;
    }

    const dynWrap = document.getElementById("manualDynamicCategorySections");
    if (dynWrap && dynWrap.dataset.renderedCat !== matchedTab && manualFormEl) {
        const preserved = captureScopeFormValues(manualFormEl);
        dynWrap.innerHTML = renderCategoryDynamicFormSections('manual', matchedTab, null, preserved);
        dynWrap.dataset.renderedCat = matchedTab;
        refreshBenefitDatalist(manualFormEl);
        refreshChargeDatalist(manualFormEl);
        refreshNonBenefitDatalist(manualFormEl);
        refreshSectionItemDatalist(manualFormEl);
    }

    const uploadTarget = document.getElementById("uploadTargetSheet");
    if (uploadTarget && uploadTarget.value !== matchedTab) {
        uploadTarget.value = matchedTab;
    }
    updateUploadModalTheme(matchedTab);

    const catCfg = getCategorySchemaConfig(matchedTab);
    const specialCount = (catCfg.specialProvisions || []).length;
    const checklistSecs = (catCfg.productSections || []).filter(s => s.mode === "checklist").map(s => s.title);
    const chargeSecs = (catCfg.productSections || []).filter(s => s.isChargeLike).map(s => s.title);

    if (catFeedback) {
        if (specialCount > 0 || checklistSecs.length > 0 || chargeSecs.length > 0) {
            const summaryParts = [];
            if (specialCount > 0) summaryParts.push(`${specialCount} Special Provision(s)`);
            if (checklistSecs.length > 0) summaryParts.push(`Checklist: ${checklistSecs.join(", ")}`);
            if (chargeSecs.length > 0) summaryParts.push(chargeSecs.join(", "));
            catFeedback.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
            catFeedback.innerHTML = `<i class="fa-solid fa-layer-group me-1"></i> <strong>${escapeHtmlAttr(matchedTab)}</strong> category active: ${escapeHtmlAttr(summaryParts.join(" · "))}.`;
        } else {
            catFeedback.className = "hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
            catFeedback.replaceChildren();
        }
    }
    return true;
}

function validateManualProductName(inputEl) {
    const rawName = inputEl.value.trim();
    const feedbackEl = document.getElementById("manualProductNameFeedback");
    if (!feedbackEl) return true;

    if (!rawName) {
        feedbackEl.className = "hidden mt-1.5 text-[11px] px-2.5 py-1 rounded-md border";
        feedbackEl.replaceChildren();
        return false;
    }

    const existingProd = products.find(p => p.name.toLowerCase().trim() === rawName.toLowerCase());
    if (existingProd) {
        feedbackEl.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-triangle-exclamation me-1"></i> Product <strong>"${escapeHtml(existingProd.name)}"</strong> is already registered under category <strong>${escapeHtml(existingProd.category)}</strong>!`;
        return false;
    } else {
        feedbackEl.className = "block mt-1.5 text-[11px] px-2.5 py-1 rounded-md border bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800 font-medium";
        feedbackEl.innerHTML = `<i class="fa-solid fa-circle-check me-1"></i> Product name is available and not yet in the database.`;
        return true;
    }
}

async function saveManualProduct() {
    const manualFormEl = document.getElementById("manualFormBody");
    const prodNameInput = document.getElementById("manual_productName");
    const compNameInput = document.getElementById("manual_company");
    const targetCategory = document.getElementById("manual_category").value;

    const prodName = prodNameInput.value.trim();
    const compName = compNameInput.value.trim();

    if (!validateManualCategoryAndToggleSyariah()) {
        showToast(`Invalid product category! Please choose one of: ${SHEET_TABS.join(", ")}.`);
        return;
    }

    if (!prodName) {
        showToast("Please fill in 'Product Name' first.");
        prodNameInput.focus();
        return;
    }

    if (!validateManualProductName(prodNameInput)) {
        showToast(`Product "${prodName}" already exists in the database. Please use the Edit Product feature to modify its data.`);
        prodNameInput.focus();
        return;
    }

    if (!compName) {
        showToast("Please fill in 'Company Name' first.");
        compNameInput.focus();
        return;
    }

    if (!validateAllNewRepeatableRows(manualFormEl)) return;

    const btnSave = document.getElementById("btnSaveManualProduct");
    btnSave.disabled = true;
    btnSave.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving...`;

    const newType = document.getElementById("manual_productType").value.trim() || "-";
    const newChannel = document.getElementById("manual_channel").value.trim() || "-";
    const newCurr = document.getElementById("manual_currency").value.trim() || "-";
    const manualLaunchEl = document.getElementById("manual_launchDate");
    const newLaunchDate = (manualLaunchEl && manualLaunchEl.value) ? formatDateToDDMMYYYY(manualLaunchEl.value) : "-";

    const { updatedAttr, customNonBenefitsToRegister } = await collectAllDynamicFormData(manualFormEl, targetCategory);

    const targetCols = getColumnsForTargetSheet(targetCategory, updatedAttr);
    let baseMasterSchema = [
        "Product Name", "Company Name", "Par/Non Par", "Product Type",
        "Distribution Channel", "Currency", "Launch Date", ...targetCols
    ];

    let baseRowData = baseMasterSchema.map(colName => {
        if (colName === "Product Name") return prodName;
        if (colName === "Company Name") return compName;
        if (colName === "Par/Non Par") return "-";
        if (colName === "Product Type") return newType;
        if (colName === "Distribution Channel") return newChannel;
        if (colName === "Currency") return newCurr;
        if (colName === "Launch Date") return newLaunchDate;
        return updatedAttr[colName] || "-";
    });

    await offloadOversizedCellImagesToDrive([baseRowData]);
    const { finalSchema: masterSchema, finalRows } = expandRowsWithContinuationColumns(baseMasterSchema, [baseRowData]);
    const rowData = finalRows[0];

    const oversizedManual = findOversizedCells(masterSchema, rowData);
    if (oversizedManual.length > 0) {
        showToast(`Failed to save: the following columns exceed Google Sheets' 50,000-character limit per cell:\n- ${oversizedManual.join("\n- ")}\n\nPlease shorten the text or use smaller images.`);
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save New Product`;
        return;
    }

    try {
        if (!sheetHasHeaders[targetCategory]) {
            await appendToSheetWithHeaderCheck(targetCategory, masterSchema, [rowData]);
        } else {
            const payload = {
                action: "update",
                sheetName: targetCategory,
                productName: prodName,
                headers: masterSchema,
                data: rowData
            };

            await postAuthenticatedAppsScript(payload);
        }

        if (customNonBenefitsToRegister.length > 0) {
            try {
                await postAuthenticatedAppsScript({ action: "save_benefit_aliases", mappings: customNonBenefitsToRegister });
                customNonBenefitsToRegister.forEach(m => {
                    if (!m || !m.canonical) return;
                    const cKey = normalizeBenefitNameKey(m.canonical);
                    if (String(m.entryType || "").toLowerCase() === "non-benefit") {
                        NON_BENEFIT_CONFIG_CANONICALS.add(cKey);
                    } else {
                        BENEFIT_CONFIG_CANONICALS.add(cKey);
                        if (!BENEFIT_CONFIG_CANONICAL_LIST.some(b => b.toLowerCase() === String(m.canonical).toLowerCase())) {
                            BENEFIT_CONFIG_CANONICAL_LIST.push(String(m.canonical).trim());
                        }
                    }
                });
                invalidateReviewLookupCaches();
            } catch (e) {
                if (e && e.code === "UNAUTHORIZED") throw e;
                console.warn("Failed to save custom benefit/non-benefit to Config:", e);
            }
        }

        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save New Product`;

        closeUploadModal();
        invalidateLocalSummaryCache();
        fetchAllSheetData(true);

    } catch (err) {
        console.error(err);
        if (!err || err.code !== "UNAUTHORIZED") {
            showToast("Failed to save new product to Google Sheets. Please check your internet connection.");
        }
        btnSave.disabled = false;
        btnSave.innerHTML = `<i class="fa-solid fa-floppy-disk"></i> Save New Product`;
    }
}

// --- DELETE PRODUCTS (SINGLE & MULTI-SELECT WITH CONFIRMATION + TABLE IMAGES & DRIVE DOCS CLEANUP) ---
let pendingDeleteProductIds = [];

function extractProductTableImagesInfo(targetProducts) {
    let totalCount = 0;
    const driveIdSet = new Set();

    (targetProducts || []).forEach(p => {
        if (!p || !p.attributes) return;
        Object.values(p.attributes).forEach(val => {
            if (!val || val === "-") return;
            const { images } = splitTextAndImages(val);
            totalCount += images.length;
            images.forEach(src => {
                if (String(src).startsWith("drive:")) {
                    const fId = String(src).slice(6).trim();
                    if (fId) driveIdSet.add(fId);
                }
            });
        });
    });

    return {
        totalCount,
        driveFileIds: Array.from(driveIdSet)
    };
}

function hasDriveFolderForProduct(productName) {
    const cleanTarget = String(productName || "").trim().toLowerCase();
    if (!cleanTarget) return false;
    if (productDriveFolders.some(f => String(f || "").trim().toLowerCase() === cleanTarget)) {
        return true;
    }
    return getDocumentsForProduct(productName).length > 0;
}

function extractProductDocumentsInfo(targetProducts) {
    let totalDocs = 0;
    let matchingFoldersCount = 0;
    const docIdSet = new Set();

    (targetProducts || []).forEach(p => {
        if (!p || !p.name) return;
        const docs = getDocumentsForProduct(p.name);
        totalDocs += docs.length;
        if (hasDriveFolderForProduct(p.name)) {
            matchingFoldersCount++;
        }
        docs.forEach(d => {
            if (d && d.fileId) docIdSet.add(d.fileId);
        });
    });

    return {
        totalDocs,
        docFileIds: Array.from(docIdSet),
        matchingFoldersCount
    };
}

async function openDeleteConfirmModal(idsArray, event = null) {
    if (event) event.stopPropagation();
    if (!idsArray || idsArray.length === 0) return;

    let targets = products.filter(p => idsArray.includes(p.id));
    if (targets.length === 0) return;

    pendingDeleteProductIds = targets.map(p => p.id);

    const summaryEl = document.getElementById("deleteConfirmSummary");
    const listEl = document.getElementById("deleteConfirmList");
    const btnExec = document.getElementById("btnExecuteDelete");
    const modalEl = document.getElementById("deleteConfirmModal");

    summaryEl.innerHTML = `Are you sure you want to delete the following <strong>${targets.length} product${targets.length > 1 ? 's' : ''}</strong> from the database?`;

    const chkImgs = document.getElementById("chkDeleteTableImages");
    const badgeImgs = document.getElementById("deleteTableImagesBadge");
    const descImgs = document.getElementById("deleteTableImagesDesc");
    const lblImgs = document.getElementById("lblDeleteTableImages");
    const chkDocs = document.getElementById("chkDeleteProductDocs");
    const badgeDocs = document.getElementById("deleteProductDocsBadge");
    const descDocs = document.getElementById("deleteProductDocsDesc");
    const lblDocs = document.getElementById("lblDeleteProductDocs");

    if (!isFullProductDataLoaded || !isDocumentsLoaded) {
        btnExec.disabled = true;
        btnExec.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Checking attachments...`;
        listEl.innerHTML = targets.map(p => `
                    <div class="flex items-center justify-between gap-2 py-1.5 border-b last:border-b-0 border-red-200/60 dark:border-red-800/40">
                        <div class="min-w-0">
                            <span class="font-bold text-gray-900 dark:text-neutral-100 block truncate">${escapeHtml(p.name)}</span>
                            <div class="text-[11px] text-gray-500 dark:text-neutral-400">${escapeHtml(p.company)} &bull; Checking images &amp; documents...</div>
                        </div>
                        ${getCategoryBadge(p.category)}
                    </div>`).join('');
        if (badgeImgs) badgeImgs.textContent = "Checking...";
        if (badgeDocs) badgeDocs.textContent = "Checking...";
        modalEl.classList.remove("hidden");

        await Promise.all([ensureFullProductDataLoaded(), ensureDocumentsLoaded()]);
        if (modalEl.classList.contains("hidden") || pendingDeleteProductIds.length === 0) return;
        targets = products.filter(p => pendingDeleteProductIds.includes(p.id));
        if (targets.length === 0) {
            closeDeleteConfirmModal();
            return;
        }
    }

    btnExec.disabled = false;
    btnExec.innerHTML = `<i class="fa-solid fa-trash-can"></i> Yes, Delete (${targets.length})`;

    listEl.innerHTML = targets.map(p => {
        const pImgs = extractProductTableImagesInfo([p]).totalCount;
        const pDocs = getDocumentsForProduct(p.name).length;
        return `
                    <div class="flex items-center justify-between gap-2 py-1.5 border-b last:border-b-0 border-red-200/60 dark:border-red-800/40">
                        <div class="min-w-0">
                            <span class="font-bold text-gray-900 dark:text-neutral-100 block truncate">${escapeHtml(p.name)}</span>
                            <div class="flex flex-wrap items-center gap-1.5 text-[11px] text-gray-500 dark:text-neutral-400">
                                <span>${escapeHtml(p.company)}</span>
                                ${pImgs > 0 ? `<span class="text-[10px] px-1.5 py-0.2 rounded bg-blue-100/80 dark:bg-neutral-700/70 text-blue-700 dark:text-blue-300 font-semibold"><i class="fa-solid fa-images me-0.5"></i>${pImgs} Image${pImgs > 1 ? 's' : ''}</span>` : ''}
                                ${pDocs > 0 ? `<span class="text-[10px] px-1.5 py-0.2 rounded bg-amber-100/80 dark:bg-amber-900/50 text-amber-800 dark:text-amber-300 font-semibold"><i class="fa-solid fa-paperclip me-0.5"></i>${pDocs} Doc${pDocs > 1 ? 's' : ''}</span>` : ''}
                            </div>
                        </div>
                        ${getCategoryBadge(p.category)}
                    </div>
                `;
    }).join('');

    const imgInfo = extractProductTableImagesInfo(targets);
    const docInfo = extractProductDocumentsInfo(targets);

    if (chkImgs && badgeImgs && descImgs && lblImgs) {
        badgeImgs.textContent = `${imgInfo.totalCount} image${imgInfo.totalCount === 1 ? '' : 's'}`;
        if (imgInfo.totalCount > 0) {
            chkImgs.checked = true;
            chkImgs.disabled = false;
            lblImgs.classList.remove("opacity-60");
            descImgs.innerHTML = `Deletes <strong>${imgInfo.totalCount} table image${imgInfo.totalCount > 1 ? 's' : ''}</strong> in Benefits, Specifications, Contribution Allocation, or Charges for the selected product(s) (including ${imgInfo.driveFileIds.length} HD file(s) in Google Drive).`;
        } else {
            chkImgs.checked = false;
            chkImgs.disabled = true;
            lblImgs.classList.add("opacity-60");
            descImgs.textContent = "No table images found in the selected product(s).";
        }
    }

    if (chkDocs && badgeDocs && descDocs && lblDocs) {
        badgeDocs.textContent = `${docInfo.totalDocs} doc${docInfo.totalDocs === 1 ? '' : 's'}`;
        if (docInfo.totalDocs > 0) {
            chkDocs.disabled = false;
            chkDocs.checked = true;
            lblDocs.classList.remove("opacity-60");
            descDocs.innerHTML = `Deletes all <strong>${docInfo.totalDocs} product document${docInfo.totalDocs > 1 ? 's' : ''}</strong> in Google Drive along with their <strong>product folder</strong> automatically.`;
        } else if (docInfo.matchingFoldersCount > 0) {
            chkDocs.disabled = false;
            chkDocs.checked = true;
            lblDocs.classList.remove("opacity-60");
            descDocs.innerHTML = `No documents detected, but ${docInfo.matchingFoldersCount === 1 ? 'an empty product folder exists' : `${docInfo.matchingFoldersCount} empty product folders exist`} in Google Drive and will be cleaned up.`;
        } else {
            chkDocs.checked = false;
            chkDocs.disabled = true;
            lblDocs.classList.add("opacity-60");
            descDocs.textContent = "No product folder or documents found in Google Drive for the selected product(s).";
        }
    }

    modalEl.classList.remove("hidden");
}

function closeDeleteConfirmModal() {
    document.getElementById("deleteConfirmModal").classList.add("hidden");
    pendingDeleteProductIds = [];
}

async function executeDeleteProducts() {
    if (pendingDeleteProductIds.length === 0) return;

    const targets = products.filter(p => pendingDeleteProductIds.includes(p.id));
    if (targets.length === 0) {
        closeDeleteConfirmModal();
        return;
    }

    const deleteTableImages = Boolean(document.getElementById("chkDeleteTableImages")?.checked);
    const deleteProductDocs = Boolean(document.getElementById("chkDeleteProductDocs")?.checked);

    const imgInfo = extractProductTableImagesInfo(targets);
    const docInfo = extractProductDocumentsInfo(targets);

    const btnExec = document.getElementById("btnExecuteDelete");
    btnExec.disabled = true;
    btnExec.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Deleting...`;

    try {
        // 1. Delete HD table images in Google Drive if checked
        if (deleteTableImages && imgInfo.driveFileIds.length > 0) {
            btnExec.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Deleting Table Images (${imgInfo.driveFileIds.length})...`;
            const imgDelResults = await Promise.allSettled(
                imgInfo.driveFileIds.map(fileId => {
                    delete driveImageCache[fileId];
                    return postAuthenticatedAppsScript({
                        action: "delete_document",
                        fileId: fileId
                    }, { requireSuccess: false });
                })
            );
            const unauthImg = imgDelResults.find(r => r.status === "rejected" && r.reason && r.reason.code === "UNAUTHORIZED");
            if (unauthImg) throw unauthImg.reason;
        }

        // 2. Delete product documents & folders in Google Drive if checked
        if (deleteProductDocs && docInfo.docFileIds.length > 0) {
            btnExec.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Deleting Drive Docs & Folders (${docInfo.docFileIds.length})...`;
            const docDelResults = await Promise.allSettled(
                docInfo.docFileIds.map(fileId =>
                    postAuthenticatedAppsScript({
                        action: "delete_document",
                        fileId: fileId,
                        deleteEmptyParentFolder: true
                    }, { requireSuccess: false })
                )
            );
            const unauthDoc = docDelResults.find(r => r.status === "rejected" && r.reason && r.reason.code === "UNAUTHORIZED");
            if (unauthDoc) throw unauthDoc.reason;
            const deletedDocIdSet = new Set(docInfo.docFileIds);
            productDocuments = productDocuments.filter(d => !deletedDocIdSet.has(d.fileId));
        }

        // 3. Delete product rows from Master Google Sheets
        btnExec.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Deleting Products from Sheets...`;
        const payload = {
            action: "delete",
            deleteTableImages: deleteTableImages,
            tableImageFileIds: deleteTableImages ? imgInfo.driveFileIds : [],
            deleteProductDocs: deleteProductDocs,
            items: targets.map(p => ({
                sheetName: p.category,
                productName: p.name
            }))
        };

        await postAuthenticatedAppsScript(payload);

        selectedProductIds = selectedProductIds.filter(id => !pendingDeleteProductIds.includes(id));
        updateCompareBar();

        if (currentActiveProduct && pendingDeleteProductIds.includes(currentActiveProduct.id)) {
            closeDetailModal();
        }

        btnExec.disabled = false;
        btnExec.innerHTML = `<i class="fa-solid fa-trash-can"></i> Yes, Delete Permanently`;

        closeDeleteConfirmModal();
        invalidateLocalSummaryCache();
        fetchAllSheetData(true);

    } catch (err) {
        console.error(err);
        if (!err || err.code !== "UNAUTHORIZED") {
            showToast("Failed to delete products from Google Sheets. Please check your internet connection.");
        }
        btnExec.disabled = false;
        btnExec.innerHTML = `<i class="fa-solid fa-trash-can"></i> Yes, Delete Permanently`;
    }
}

// --- EKSTRAKSI & KOMPRESI GAMBAR EMBEDDED DARI FILE EXCEL (.XLSX) ---
function loadImageFromBlob(blob) {
    return new Promise((resolve, reject) => {
        const url = URL.createObjectURL(blob);
        const img = new Image();
        img.onload = () => {
            URL.revokeObjectURL(url);
            resolve(img);
        };
        img.onerror = (err) => {
            URL.revokeObjectURL(url);
            reject(err);
        };
        img.src = url;
    });
}

function compressSingleImageToWebpDataUrl(img, maxCharBudget = 45000) {
    // Turunkan skala secara bertahap sampai DIJAMIN muat di bawah batas 1 sel Google Sheets (50.000 karakter).
    // (Sebelumnya fallback mengembalikan versi 100% yang masih kebesaran -> sel > 50.000 karakter -> update gagal diam-diam)
    const scales = [1.0, 0.9, 0.8, 0.7, 0.62, 0.55, 0.48, 0.42, 0.36, 0.3, 0.25, 0.2];
    const qualities = [0.95, 0.9, 0.85, 0.8, 0.74, 0.68, 0.6];
    let bestFallbackUrl = "";
    const natW = img.naturalWidth || img.width || 1;
    const natH = img.naturalHeight || img.height || 1;

    for (const scale of scales) {
        const w = Math.max(1, Math.round(natW * scale));
        const h = Math.max(1, Math.round(natH * scale));
        const canvas = document.createElement("canvas");
        canvas.width = w;
        canvas.height = h;
        const ctx = canvas.getContext("2d");
        if (scale === 1.0) {
            ctx.imageSmoothingEnabled = false;
        } else {
            ctx.imageSmoothingEnabled = true;
            ctx.imageSmoothingQuality = "high";
        }
        ctx.fillStyle = "#FFFFFF";
        ctx.fillRect(0, 0, w, h);
        ctx.drawImage(img, 0, 0, w, h);

        // Jika pada ukuran asli (100%) format PNG lossless sudah muat di bawah batas per-sel, gunakan PNG lossless langsung
        if (scale === 1.0) {
            const pngUrl = canvas.toDataURL("image/png");
            if (pngUrl.length <= maxCharBudget) {
                return pngUrl;
            }
        }

        for (const q of qualities) {
            const dataUrl = canvas.toDataURL("image/webp", q);
            // Simpan hasil TERKECIL sebagai fallback terakhir
            if (!bestFallbackUrl || dataUrl.length < bestFallbackUrl.length) {
                bestFallbackUrl = dataUrl;
            }
            if (dataUrl.length <= maxCharBudget) {
                return dataUrl;
            }
        }
    }
    return bestFallbackUrl;
}

// Pastikan setiap data URL gambar muat di 1 sel Google Sheets; kompres ulang bila kebesaran
async function ensureImageSrcFitsCell(src, maxCharBudget = 45000) {
    if (!src || src.length <= maxCharBudget || !src.startsWith("data:image/")) return src;
    try {
        const res = await fetch(src);
        const blob = await res.blob();
        const img = await loadImageFromBlob(blob);
        return compressSingleImageToWebpDataUrl(img, maxCharBudget);
    } catch (e) {
        console.warn("Gagal mengompres ulang gambar:", e);
        return src;
    }
}

// Cek sel yang masih melebihi batas Google Sheets sebelum dikirim (agar tidak gagal diam-diam)
function findOversizedCells(schema, row, limit = 49500) {
    const result = [];
    schema.forEach((col, i) => {
        if (String(row[i] ?? "").length > limit) result.push(col);
    });
    return result;
}

// Kompresi setiap gambar pada resolusi asli (100%) tanpa saling mengorbankan kualitas antar gambar
// (Jika total beberapa gambar dalam 1 atribut > 47.000 karakter, otomatis dibagi ke kolom lanjutan seperti "Alokasi Kontribusi [2]")
function compressImageGroupToTokens(imgList) {
    if (!imgList || imgList.length === 0) return [];
    const PER_IMAGE_SAFE_BUDGET = 45000;

    return imgList.map(img => {
        const dataUrl = compressSingleImageToWebpDataUrl(img, PER_IMAGE_SAFE_BUDGET);
        return `[IMG:${dataUrl}]`;
    });
}

function findAllByLocalName(rootNode, localName) {
    if (!rootNode) return [];
    const results = [];
    const all = rootNode.getElementsByTagName("*");
    for (let i = 0; i < all.length; i++) {
        const el = all[i];
        const ln = el.localName || el.nodeName.split(":").pop();
        if (ln === localName) results.push(el);
    }
    return results;
}

function findFirstByLocalName(rootNode, localName) {
    const list = findAllByLocalName(rootNode, localName);
    return list.length > 0 ? list[0] : null;
}

function getBlipEmbedId(blipEl) {
    if (!blipEl || !blipEl.attributes) return "";
    for (let i = 0; i < blipEl.attributes.length; i++) {
        const attr = blipEl.attributes[i];
        if (attr.name === "r:embed" || attr.localName === "embed" || attr.name.endsWith(":embed")) {
            return attr.value;
        }
    }
    return "";
}

async function extractImagesFromXlsxBuffer(arrayBuffer) {
    const sheetImagesMap = {};
    if (typeof JSZip === "undefined") return sheetImagesMap;

    const MAX_ZIP_ENTRIES = 2000;
    const MAX_UNCOMPRESSED_BYTES = 100 * 1024 * 1024; // 100 MB batas total dekompresi aman

    try {
        const zip = await JSZip.loadAsync(arrayBuffer);
        const entries = Object.values(zip.files || {});
        if (entries.length > MAX_ZIP_ENTRIES) {
            throw new Error(`Archive contains too many entries (${entries.length} > ${MAX_ZIP_ENTRIES}).`);
        }
        let totalUncompressed = 0;
        for (const entry of entries) {
            const uSize = Number(entry && entry._data && entry._data.uncompressedSize) || 0;
            totalUncompressed += Math.max(0, uSize);
            if (totalUncompressed > MAX_UNCOMPRESSED_BYTES) {
                throw new Error("Archive uncompressed size exceeds the 100 MB safety limit.");
            }
        }
        let actualDecompressedBytes = 0;
        const safeReadZipBytes = async (zipEntry) => {
            const bytes = await zipEntry.async("uint8array");
            actualDecompressedBytes += bytes.byteLength;
            if (actualDecompressedBytes > MAX_UNCOMPRESSED_BYTES) {
                throw new Error("Actual decompressed archive size exceeds the 100 MB safety limit.");
            }
            return bytes;
        };
        const safeReadZipString = async (zipEntry) => {
            const bytes = await safeReadZipBytes(zipEntry);
            return new TextDecoder().decode(bytes);
        };
        const parser = new DOMParser();

        const wbXmlFile = zip.file("xl/workbook.xml");
        const wbRelsFile = zip.file("xl/_rels/workbook.xml.rels");
        if (!wbXmlFile || !wbRelsFile) return sheetImagesMap;

        const wbDoc = parser.parseFromString(await safeReadZipString(wbXmlFile), "application/xml");
        const wbRelsDoc = parser.parseFromString(await safeReadZipString(wbRelsFile), "application/xml");

        const wbRels = {};
        findAllByLocalName(wbRelsDoc, "Relationship").forEach(rel => {
            const id = rel.getAttribute("Id");
            const target = rel.getAttribute("Target") || "";
            if (id && target) {
                wbRels[id] = target.startsWith("xl/") ? target : `xl/${target.replace(/^\/+/, "")}`;
            }
        });

        const sheetsInfo = [];
        findAllByLocalName(wbDoc, "sheet").forEach(sh => {
            const name = sh.getAttribute("name");
            let rId = sh.getAttribute("r:id") || "";
            if (!rId && sh.attributes) {
                for (let i = 0; i < sh.attributes.length; i++) {
                    if (sh.attributes[i].localName === "id") {
                        rId = sh.attributes[i].value;
                        break;
                    }
                }
            }
            if (name && rId && wbRels[rId]) {
                sheetsInfo.push({ name, path: wbRels[rId] });
            }
        });

        for (const sheet of sheetsInfo) {
            const sheetFileName = sheet.path.split("/").pop();
            const sheetRelsPath = `xl/worksheets/_rels/${sheetFileName}.rels`;
            const sheetRelsFile = zip.file(sheetRelsPath);
            if (!sheetRelsFile) continue;

            const sheetRelsDoc = parser.parseFromString(await safeReadZipString(sheetRelsFile), "application/xml");
            const drawingPaths = [];
            findAllByLocalName(sheetRelsDoc, "Relationship").forEach(rel => {
                const type = rel.getAttribute("Type") || "";
                const target = rel.getAttribute("Target") || "";
                if (type.includes("/drawing") && target) {
                    const cleanTarget = target.replace(/^(\.\.\/)+/, "");
                    drawingPaths.push(cleanTarget.startsWith("xl/") ? cleanTarget : `xl/${cleanTarget}`);
                }
            });

            const cellImgMap = {};
            const seenAnchorSignatures = new Set();

            for (const drawingPath of drawingPaths) {
                const drawingFile = zip.file(drawingPath);
                if (!drawingFile) continue;

                const drawingFileName = drawingPath.split("/").pop();
                const drawingRelsPath = `xl/drawings/_rels/${drawingFileName}.rels`;
                const drawingRelsFile = zip.file(drawingRelsPath);
                if (!drawingRelsFile) continue;

                const drawingRelsDoc = parser.parseFromString(await safeReadZipString(drawingRelsFile), "application/xml");
                const mediaRels = {};
                findAllByLocalName(drawingRelsDoc, "Relationship").forEach(rel => {
                    const id = rel.getAttribute("Id");
                    const target = rel.getAttribute("Target") || "";
                    if (id && target) {
                        const cleanTarget = target.replace(/^(\.\.\/)+/, "");
                        mediaRels[id] = cleanTarget.startsWith("xl/") ? cleanTarget : `xl/${cleanTarget}`;
                    }
                });

                const drawingDoc = parser.parseFromString(await safeReadZipString(drawingFile), "application/xml");
                const anchors = [
                    ...findAllByLocalName(drawingDoc, "twoCellAnchor"),
                    ...findAllByLocalName(drawingDoc, "oneCellAnchor")
                ];

                for (const anchor of anchors) {
                    const fromEl = findFirstByLocalName(anchor, "from");
                    const toEl = findFirstByLocalName(anchor, "to");
                    const blipEl = findFirstByLocalName(anchor, "blip");
                    if (!fromEl || !blipEl) continue;

                    const rId = getBlipEmbedId(blipEl);
                    const mediaPath = mediaRels[rId];
                    if (!rId || !mediaPath) continue;

                    const fCol = parseInt(findFirstByLocalName(fromEl, "col")?.textContent || "0", 10);
                    const fColOff = parseInt(findFirstByLocalName(fromEl, "colOff")?.textContent || "0", 10);
                    const fRow = parseInt(findFirstByLocalName(fromEl, "row")?.textContent || "0", 10);
                    const fRowOff = parseInt(findFirstByLocalName(fromEl, "rowOff")?.textContent || "0", 10);

                    const tCol = toEl ? parseInt(findFirstByLocalName(toEl, "col")?.textContent || String(fCol), 10) : fCol;
                    const tRow = toEl ? parseInt(findFirstByLocalName(toEl, "row")?.textContent || String(fRow), 10) : fRow;

                    // Lewati duplikat gambar yang tidak sengaja ter-copy dengan rId, row, dan rowOff identik di kolom lain
                    const dupSig = `${rId}|${fRow}|${fRowOff}`;
                    if (seenAnchorSignatures.has(dupSig)) continue;
                    seenAnchorSignatures.add(dupSig);

                    // Koreksi posisi visual jika anchor berada di tepi paling kanan kolom sebelumnya atau tepi bawah baris sebelumnya
                    const visCol = (tCol > fCol && fColOff > 1500000) ? fCol + 1 : fCol;
                    const visRow = (tRow > fRow && fRowOff > 500000) ? fRow + 1 : fRow;

                    const mediaZipObj = zip.file(mediaPath);
                    if (!mediaZipObj) continue;

                    const ext = mediaPath.split(".").pop().toLowerCase();
                    const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : (ext === "gif" ? "image/gif" : "image/png");
                    const uint8 = await safeReadZipBytes(mediaZipObj);
                    const blob = new Blob([uint8], { type: mime });
                    const imgElement = await loadImageFromBlob(blob);

                    if (!cellImgMap[visCol]) cellImgMap[visCol] = {};
                    if (!cellImgMap[visCol][visRow]) cellImgMap[visCol][visRow] = [];
                    cellImgMap[visCol][visRow].push({
                        imgElement,
                        sortY: fRow * 10000000 + fRowOff
                    });
                }
            }

            // Urutkan gambar di dalam sel berdasarkan posisi vertikalnya (dari atas ke bawah)
            Object.keys(cellImgMap).forEach(cKey => {
                Object.keys(cellImgMap[cKey]).forEach(rKey => {
                    cellImgMap[cKey][rKey].sort((a, b) => a.sortY - b.sortY);
                });
            });

            sheetImagesMap[sheet.name] = cellImgMap;
        }
    } catch (err) {
        console.warn("Ekstraksi gambar Excel dilewati:", err);
    }

    return sheetImagesMap;
}

// --- CACHE WORKBOOK & PEMILIH SHEET (MULTI-SHEET UPLOAD) ---
let uploadWorkbookCache = null; // { workbook, images, baseColumns, sheetCounts }
let selectedUploadSheets = [];
let uploadSheetTargets = {}; // { [sheetName]: "" (ikuti Target Tab) | "Konven" | "Syariah" | "HS" }

// Tebak Target Tab otomatis dari nama sheet Excel (mendukung daftar Sheet & Nilai / Nama Lain di Config)
function guessTargetTabFromSheetName(name) {
    const n = String(name || "").trim().toLowerCase();
    if (!n) return "";
    // 1. Cek kecocokan langsung atau substring dengan daftar tab di SHEET_TABS (dari Config)
    const exactMatch = SHEET_TABS.find(t => t.toLowerCase() === n);
    if (exactMatch) return exactMatch;
    const subMatch = SHEET_TABS.find(t => t.length >= 2 && n.includes(t.toLowerCase()));
    if (subMatch) return subMatch;

    // 1b. Cek daftar alias sheet dari Config (SHEET_ALIAS_RULES)
    if (Array.isArray(SHEET_ALIAS_RULES) && SHEET_ALIAS_RULES.length > 0) {
        for (let i = 0; i < SHEET_ALIAS_RULES.length; i++) {
            const rule = SHEET_ALIAS_RULES[i];
            const aliases = Array.isArray(rule.aliases) ? rule.aliases : [];
            if (aliases.some(a => {
                const al = String(a || "").trim().toLowerCase();
                return al && (n === al || (al.length >= 3 && n.includes(al)));
            })) {
                const matchedTab = SHEET_TABS.find(t => t.toLowerCase() === rule.canonical.toLowerCase());
                if (matchedTab) return matchedTab;
            }
        }
    }

    // 2. General aliases (matched dynamically against SHEET_TABS)
    if (/syari|sharia|takaful|\bsya\b/.test(n)) {
        const shariaTab = SHEET_TABS.find(t => /^(sharia|syariah)$/i.test(t));
        if (shariaTab) return shariaTab;
    }
    if (/\bhs\b|health|kesehatan|hospital|surgical/.test(n)) {
        const healthTab = SHEET_TABS.find(t => /^(health|hs)/i.test(t));
        if (healthTab) return healthTab;
    }
    if (/\bdplk\b|\bppip\b|\bdkp\b|pensiun|pension|winpension/.test(n)) {
        const dplkTab = SHEET_TABS.find(t => /^dplk$/i.test(t));
        if (dplkTab) return dplkTab;
    }
    if (/konven|conven|conv\b|tradisional/.test(n)) {
        const convTab = SHEET_TABS.find(t => /^(conventional|konven)$/i.test(t));
        if (convTab) return convTab;
    }
    return "";
}

// Target Tab akhir untuk sebuah sheet (jika tidak diatur, ikuti Target Tab utama)
function resolveSheetTarget(sheetName) {
    const globalTarget = (document.getElementById("uploadTargetSheet") || {}).value || SHEET_TABS[0] || "Conventional";
    const isMulti = uploadWorkbookCache && uploadWorkbookCache.workbook.SheetNames.length > 1;
    if (!isMulti) return globalTarget;
    return uploadSheetTargets[sheetName] || globalTarget;
}

function onUploadSheetTargetChange(idx, value) {
    if (!uploadWorkbookCache) return;
    const name = uploadWorkbookCache.workbook.SheetNames[idx];
    uploadSheetTargets[name] = value;
    preserveUploadScroll(() => analyzeSelectedUploadSheets());
}

const TARGET_TAB_BADGE = {
    Conventional: "bg-sky-100 text-sky-800 dark:bg-neutral-700/70 dark:text-sky-300",
    Sharia: "bg-green-100 text-green-800 dark:bg-green-900/50 dark:text-green-300",
    "Health & Surgical": "bg-purple-100 text-purple-900 dark:bg-purple-950/70 dark:text-purple-300",
    DPLK: "bg-amber-100 text-amber-800 dark:bg-amber-900/50 dark:text-amber-300",
    Konven: "bg-sky-100 text-sky-800 dark:bg-neutral-700/70 dark:text-sky-300",
    Syariah: "bg-green-100 text-green-800 dark:bg-green-900/50 dark:text-green-300",
    HS: "bg-purple-100 text-purple-900 dark:bg-purple-950/70 dark:text-purple-300"
};

let uploadSheetAnalyzeTimer = null;

function resetUploadSheetSelector() {
    if (uploadSheetAnalyzeTimer) {
        clearTimeout(uploadSheetAnalyzeTimer);
        uploadSheetAnalyzeTimer = null;
    }
    uploadWorkbookCache = null;
    selectedUploadSheets = [];
    uploadSheetTargets = {};
    const sel = document.getElementById("uploadSheetSelector");
    if (sel) {
        sel.replaceChildren();
        sel.classList.add("hidden");
    }
}

function countExtractedImagesInRows(rows) {
    let count = 0;
    if (!Array.isArray(rows)) return 0;
    for (let r = 0; r < rows.length; r++) {
        const row = rows[r];
        if (!Array.isArray(row)) continue;
        for (let c = 0; c < row.length; c++) {
            const v = row[c];
            if (typeof v === "string" && v.indexOf("[IMG:data:image/") !== -1) {
                const m = v.match(/\[IMG:data:image\//g);
                if (m) count += m.length;
            }
        }
    }
    return count;
}

function parseAndCacheSingleSheet(workbook, sheetName, images, baseColumns, baseNonBenefits) {
    dynamicColumns = baseColumns.slice();
    NON_BENEFIT_COLUMNS = baseNonBenefits.slice();
    let rows = [];
    try {
        rows = parseExcelToSheetsSchema({ SheetNames: [sheetName], Sheets: workbook.Sheets }, images);
    } catch (err) {
        console.warn("Gagal mem-parse produk pada sheet", sheetName, err);
        rows = [];
    }
    const sheetDynCols = dynamicColumns.slice();
    const sheetNbCols = NON_BENEFIT_COLUMNS.slice();
    const isFlat = rows.length > 0 && rows[0].length === (6 + sheetDynCols.length);
    const schema = isFlat
        ? ["Product Name", "Company Name", "Par/Non Par", "Product Type", "Distribution Channel", "Currency", ...sheetDynCols]
        : ["Product Name", "Company Name", "Par/Non Par", "Product Type", "Distribution Channel", "Currency", "Launch Date", ...sheetDynCols];
    return {
        rows,
        schema,
        isFlat,
        dynamicCols: sheetDynCols,
        nonBenefitCols: sheetNbCols,
        imageCount: countExtractedImagesInRows(rows)
    };
}

// Hitung jumlah produk per sheet & simpan cache hasil parse per-sheet agar tick/untick sheet instan
function countProductsPerSheet(workbook, images, baseColumns, baseNonBenefits = NON_BENEFIT_COLUMNS.slice(), parsedBySheetOut = null) {
    const counts = {};
    workbook.SheetNames.forEach(name => {
        const cached = parseAndCacheSingleSheet(workbook, name, images, baseColumns, baseNonBenefits);
        counts[name] = cached.rows.length;
        if (parsedBySheetOut) parsedBySheetOut[name] = cached;
    });
    dynamicColumns = baseColumns.slice();
    NON_BENEFIT_COLUMNS = baseNonBenefits.slice();
    return counts;
}

function mergeCachedSelectedSheets(cacheObj, selectedNames) {
    if (!cacheObj.parsedBySheet) cacheObj.parsedBySheet = {};
    const baseCols = cacheObj.baseColumns || [];
    const baseNbs = Array.isArray(cacheObj.baseNonBenefits) ? cacheObj.baseNonBenefits : NON_BENEFIT_COLUMNS.slice();

    const entries = [];
    selectedNames.forEach(name => {
        if (!cacheObj.parsedBySheet[name]) {
            cacheObj.parsedBySheet[name] = parseAndCacheSingleSheet(
                cacheObj.workbook, name, cacheObj.images, baseCols, baseNbs
            );
        }
        entries.push(cacheObj.parsedBySheet[name]);
    });

    // Gabungkan dynamicColumns & NON_BENEFIT_COLUMNS hanya dari sheet yang dipilih (jaga urutan)
    const mergedDyn = baseCols.slice();
    const mergedDynLower = new Set(mergedDyn.map(c => String(c || "").toLowerCase()));
    const mergedNb = baseNbs.slice();
    const mergedNbLower = new Set(mergedNb.map(c => String(c || "").toLowerCase()));

    entries.forEach(entry => {
        entry.dynamicCols.forEach(col => {
            const low = String(col || "").toLowerCase();
            if (low && !mergedDynLower.has(low)) {
                mergedDyn.push(col);
                mergedDynLower.add(low);
            }
        });
        entry.nonBenefitCols.forEach(col => {
            const low = String(col || "").toLowerCase();
            if (low && !mergedNbLower.has(low)) {
                mergedNb.push(col);
                mergedNbLower.add(low);
            }
        });
    });

    dynamicColumns = mergedDyn;
    NON_BENEFIT_COLUMNS = mergedNb;

    const nonEmptyEntries = entries.filter(e => e.rows.length > 0);
    const allFlat = nonEmptyEntries.length > 0 && nonEmptyEntries.every(e => e.isFlat);
    const mergedSchema = allFlat
        ? ["Product Name", "Company Name", "Par/Non Par", "Product Type", "Distribution Channel", "Currency", ...mergedDyn]
        : ["Product Name", "Company Name", "Par/Non Par", "Product Type", "Distribution Channel", "Currency", "Launch Date", ...mergedDyn];

    const mergedRows = [];
    let totalExtractedImages = 0;

    nonEmptyEntries.forEach(entry => {
        totalExtractedImages += entry.imageCount || 0;
        // Jika skema sheet sudah identik panjang & kolomnya, langsung salin referensi baris
        const sameSchema = entry.schema.length === mergedSchema.length &&
            entry.schema.every((col, idx) => col === mergedSchema[idx]);
        if (sameSchema) {
            entry.rows.forEach(r => mergedRows.push(r));
            return;
        }
        const colIdxMap = new Map();
        entry.schema.forEach((col, idx) => {
            const low = String(col || "").toLowerCase();
            if (!colIdxMap.has(low)) colIdxMap.set(low, idx);
        });
        const remapIndices = mergedSchema.map(col => {
            const low = String(col || "").toLowerCase();
            return colIdxMap.has(low) ? colIdxMap.get(low) : -1;
        });
        entry.rows.forEach(r => {
            const mapped = remapIndices.map(srcIdx => (srcIdx !== -1 && r[srcIdx] !== undefined ? r[srcIdx] : "-"));
            mapped.__sheetName = r.__sheetName;
            mergedRows.push(mapped);
        });
    });

    return { allParsed: mergedRows, totalExtractedImages };
}

function renderUploadSheetSelector() {
    const sel = document.getElementById("uploadSheetSelector");
    if (!sel || !uploadWorkbookCache) return;
    const names = uploadWorkbookCache.workbook.SheetNames;
    if (names.length <= 1) {
        sel.replaceChildren();
        sel.classList.add("hidden");
        return;
    }
    const counts = uploadWorkbookCache.sheetCounts || {};
    const allChecked = names.every(n => selectedUploadSheets.includes(n));
    const globalTarget = (document.getElementById("uploadTargetSheet") || {}).value || SHEET_TABS[0] || "Conventional";

    const items = names.map((name, idx) => {
        const cnt = counts[name] || 0;
        const checked = selectedUploadSheets.includes(name);
        const disabled = cnt === 0;
        const targetVal = uploadSheetTargets[name] || "";
        return `
                    <label class="flex items-center gap-2 px-2 py-1.5 rounded-md ${disabled ? 'opacity-50 cursor-not-allowed' : 'hover:bg-gray-100 dark:hover:bg-neutral-700 cursor-pointer'} transition">
                        <input type="checkbox" class="upload-sheet-chk w-4 h-4 rounded cursor-pointer accent-blue-600" data-sheet-idx="${idx}" ${checked ? 'checked' : ''} ${disabled ? 'disabled' : ''} data-ev-change="a154" data-ev-change-args="${idcmPackArgs([idx])}">
                        <i class="fa-solid fa-table-cells text-gray-400"></i>
                        <span class="flex-1 truncate text-gray-800 dark:text-neutral-200 font-medium" title="${escapeHtmlAttr(name)}">${escapeHtmlAttr(name)}</span>
                        <span class="text-[10px] font-semibold px-2 py-0.5 rounded-full ${cnt > 0 ? 'bg-blue-100 text-blue-700 dark:bg-neutral-700/70 dark:text-blue-300' : 'bg-gray-200 text-gray-500 dark:bg-neutral-700 dark:text-neutral-400'}">${cnt > 0 ? cnt + (cnt === 1 ? ' product' : ' products') : 'no products'}</span>
                        ${disabled ? '' : `
                        <select data-ev-click="a85" data-ev-change="a155" data-ev-change-args="${idcmPackArgs([idx])}" title="Target Google Sheets tab for this sheet"
                            class="text-[11px] border dark:border-neutral-600 rounded-md px-1.5 py-0.5 bg-white dark:bg-neutral-700 text-gray-800 dark:text-neutral-100 outline-none ${checked ? '' : 'opacity-50'}">
                            <option value="" ${!targetVal ? 'selected' : ''}>→ Follow Target Tab (${escapeHtmlAttr(globalTarget)})</option>
                            ${SHEET_TABS.map(tab => `<option value="${escapeHtmlAttr(tab)}" ${targetVal === tab ? 'selected' : ''}>→ ${escapeHtmlAttr(tab)}</option>`).join('')}
                        </select>`}
                    </label>`;
    }).join("");

    sel.innerHTML = `
                <div class="flex items-center justify-between mb-1">
                    <label class="block text-xs font-bold text-gray-700 dark:text-neutral-300">Select Sheets to Upload (${names.length} ${names.length === 1 ? 'sheet' : 'sheets'} detected):</label>
                    <label class="flex items-center gap-1.5 text-[11px] font-semibold text-gray-600 dark:text-neutral-300 cursor-pointer">
                        <input type="checkbox" class="w-3.5 h-3.5 rounded cursor-pointer accent-blue-600" ${allChecked ? 'checked' : ''} data-ev-change="a156"> Select All
                    </label>
                </div>
                <div id="uploadSheetList" class="border dark:border-neutral-600 rounded-lg p-1.5 bg-gray-50 dark:bg-neutral-700/40 text-xs space-y-0.5 max-h-40 overflow-y-auto">
                    ${items}
                </div>
                <p class="text-[10px] text-gray-500 dark:text-neutral-400 mt-1"><i class="fa-solid fa-circle-info me-1"></i>Set each sheet's target tab using the dropdown on the right. Sheets whose names match the tab list (${SHEET_TABS.map(escapeHtmlAttr).join("/")}) are automatically routed to the matching tab.</p>`;
    sel.classList.remove("hidden");
}

// Jalankan render ulang tanpa membuat daftar sheet / isi modal melompat ke atas
function preserveUploadScroll(fn) {
    const modalBody = document.getElementById("uploadTabExcel");
    const sheetList = document.getElementById("uploadSheetList");
    const modalTop = modalBody ? modalBody.scrollTop : 0;
    const listTop = sheetList ? sheetList.scrollTop : 0;

    fn();

    const restore = () => {
        const mb = document.getElementById("uploadTabExcel");
        const sl = document.getElementById("uploadSheetList");
        if (sl) sl.scrollTop = listTop;
        if (mb) mb.scrollTop = modalTop;
    };
    restore();
    requestAnimationFrame(restore); // jaga-jaga jika layout baru selesai di frame berikutnya
}

function syncUploadSheetSelectorVisuals() {
    if (!uploadWorkbookCache) return;
    const names = uploadWorkbookCache.workbook.SheetNames;
    const sel = document.getElementById("uploadSheetSelector");
    if (!sel) return;
    const allChk = sel.querySelector('input[onchange^="onUploadSheetToggleAll"]');
    if (allChk) {
        const counts = uploadWorkbookCache.sheetCounts || {};
        const selectable = names.filter(n => (counts[n] || 0) > 0);
        allChk.checked = selectable.length > 0 && selectable.every(n => selectedUploadSheets.includes(n));
    }
    sel.querySelectorAll('.upload-sheet-chk').forEach(chk => {
        const idx = parseInt(chk.dataset.sheetIdx, 10);
        const sheetName = names[idx];
        const isChecked = selectedUploadSheets.includes(sheetName);
        chk.checked = isChecked;
        const rowSelect = chk.closest('label')?.querySelector('select');
        if (rowSelect) rowSelect.classList.toggle('opacity-50', !isChecked);
    });
}

function scheduleAnalyzeSelectedUploadSheets(delayMs = 60) {
    if (uploadSheetAnalyzeTimer) clearTimeout(uploadSheetAnalyzeTimer);
    uploadSheetAnalyzeTimer = setTimeout(() => {
        uploadSheetAnalyzeTimer = null;
        preserveUploadScroll(() => analyzeSelectedUploadSheets());
    }, delayMs);
}

function onUploadSheetToggle(idx, checked) {
    if (!uploadWorkbookCache) return;
    const name = uploadWorkbookCache.workbook.SheetNames[idx];
    if (checked && !selectedUploadSheets.includes(name)) selectedUploadSheets.push(name);
    if (!checked) selectedUploadSheets = selectedUploadSheets.filter(n => n !== name);
    // Jaga urutan sesuai urutan sheet di file Excel
    selectedUploadSheets = uploadWorkbookCache.workbook.SheetNames.filter(n => selectedUploadSheets.includes(n));
    syncUploadSheetSelectorVisuals();
    scheduleAnalyzeSelectedUploadSheets(60);
}

function onUploadSheetToggleAll(checked) {
    if (!uploadWorkbookCache) return;
    const counts = uploadWorkbookCache.sheetCounts || {};
    selectedUploadSheets = checked
        ? uploadWorkbookCache.workbook.SheetNames.filter(n => (counts[n] || 0) > 0)
        : [];
    syncUploadSheetSelectorVisuals();
    scheduleAnalyzeSelectedUploadSheets(60);
}

function previewExcelFile() {
    const fileInput = document.getElementById("excelFileInput");
    const previewArea = document.getElementById("uploadPreviewArea");
    const submitBtn = document.getElementById("btnSubmitUpload");

    submitBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Upload to Google Sheets`;
    resetUploadSheetSelector();
    resetBenefitReviewState();
    uploadSubmitBaseEnabled = false;

    if (!fileInput.files || fileInput.files.length === 0) {
        previewArea.classList.add("hidden");
        submitBtn.disabled = true;
        return;
    }

    const file = fileInput.files[0];
    const reader = new FileReader();
    const mySessionId = ++uploadModalSessionId;

    previewArea.innerHTML = `<p class="text-blue-600 dark:text-blue-400 font-medium"><i class="fa-solid fa-spinner fa-spin me-1"></i> Reading data & extracting table images from Excel file...</p>`;
    previewArea.classList.remove("hidden");
    submitBtn.disabled = true;

    reader.onload = async function (e) {
        try {
            await loadExcelLibs();
            if (mySessionId !== uploadModalSessionId) return;
            const arrayBuffer = e.target.result;
            const data = new Uint8Array(arrayBuffer);
            // Batasi pembacaan maksimal 1000 baris pertama agar file Excel yang memiliki
            // sel kosong terformat di baris 1.048.576 (seperti example2.xlsx) tidak menyebabkan crash memori
            const xlsxReader = window.XLSX_CORE || XLSX;
            const workbook = xlsxReader.read(data, { type: 'array', sheetRows: 1000 });

            // Ekstrak gambar tabel yang tertanam di dalam file .xlsx (jika ada)
            const excelImagesBySheet = await extractImagesFromXlsxBuffer(arrayBuffer);
            if (mySessionId !== uploadModalSessionId) return;

            const baseColumns = dynamicColumns.slice();
            const baseNonBenefits = NON_BENEFIT_COLUMNS.slice();
            const parsedBySheet = {};
            const sheetCounts = workbook.SheetNames.length > 1
                ? countProductsPerSheet(workbook, excelImagesBySheet, baseColumns, baseNonBenefits, parsedBySheet)
                : {};

            uploadWorkbookCache = { workbook, images: excelImagesBySheet, baseColumns, baseNonBenefits, sheetCounts, parsedBySheet };

            // Default: pilih semua sheet yang berisi produk
            selectedUploadSheets = workbook.SheetNames.length > 1
                ? workbook.SheetNames.filter(n => (sheetCounts[n] || 0) > 0)
                : workbook.SheetNames.slice();

            // Tebak Target Tab tiap sheet dari namanya (misal "Syariah" -> tab Syariah)
            uploadSheetTargets = {};
            workbook.SheetNames.forEach(n => { uploadSheetTargets[n] = guessTargetTabFromSheetName(n); });

            renderUploadSheetSelector();
            analyzeSelectedUploadSheets();
        } catch (err) {
            console.error(err);
            previewArea.innerHTML = `<p class="text-red-500"><i class="fa-solid fa-circle-xmark"></i> Failed to read Excel file. Make sure the .xlsx / .xls format is valid.</p>`;
            previewArea.classList.remove("hidden");
            submitBtn.disabled = true;
        }
    };

    reader.readAsArrayBuffer(file);
}

// Parse hanya sheet yang dipilih, lalu tampilkan pratinjau & cek duplikasi
function analyzeSelectedUploadSheets() {
    const previewArea = document.getElementById("uploadPreviewArea");
    const submitBtn = document.getElementById("btnSubmitUpload");
    if (!uploadWorkbookCache) return;

    submitBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Upload to Google Sheets`;
    previewArea.classList.remove("hidden");

    if (selectedUploadSheets.length === 0) {
        parsedUploadRows = [];
        parsedDuplicateRows = [];
        benefitReviewItems = [];
        renderBenefitReview();
        uploadSubmitBaseEnabled = false;
        previewArea.innerHTML = `<p class="text-amber-600 dark:text-amber-400 font-semibold"><i class="fa-solid fa-hand-pointer me-1"></i> Select at least 1 sheet to upload.</p>`;
        submitBtn.disabled = true;
        return;
    }

    try {
        // Gunakan cache hasil parse per-sheet agar tick/untick sheet instan tanpa re-parse seluruh workbook
        const { allParsed, totalExtractedImages } = mergeCachedSelectedSheets(uploadWorkbookCache, selectedUploadSheets);

        // Simpan skema hasil parse (untuk Review Benefit) & kumpulkan nama manfaat yang belum dikenali
        uploadParsedSchema = [
            "Product Name", "Company Name", "Par/Non Par", "Product Type",
            "Distribution Channel", "Currency", "Launch Date", ...dynamicColumns
        ];
        uploadParsedSchemaStructured = allParsed.length > 0 && allParsed[0].length === uploadParsedSchema.length;
        buildBenefitReviewItems(allParsed);

        // --- FITUR CEK DUPLIKASI PRODUK ---
        let existingProductNameSet = new Set(products.map(p => p.name.toLowerCase().trim()));
        parsedUploadRows = [];
        parsedDuplicateRows = [];
        let duplicateCount = 0;

        allParsed.forEach(row => {
            let prodName = String(row[0]).toLowerCase().trim();
            if (existingProductNameSet.has(prodName)) {
                duplicateCount++;
                parsedDuplicateRows.push(row);
            } else {
                parsedUploadRows.push(row);
            }
        });

        if (allParsed.length === 0) {
            previewArea.innerHTML = `<p class="text-red-500 font-semibold"><i class="fa-solid fa-triangle-exclamation"></i> No product data detected.</p>`;
            submitBtn.disabled = true;
        } else if (parsedUploadRows.length === 0 && duplicateCount > 0) {
            let dupSampleNames = parsedDuplicateRows.map(r => r[0]);
            let sampleDupText = dupSampleNames.length <= 5
                ? dupSampleNames.join(", ")
                : dupSampleNames.slice(0, 4).join(", ") + ` (+${dupSampleNames.length - 4} more)`;

            previewArea.innerHTML = `
                        <p class="font-bold text-amber-600 dark:text-amber-400"><i class="fa-solid fa-arrows-rotate"></i> All ${duplicateCount} products in this file already exist in the database.</p>
                        ${totalExtractedImages > 0 ? `<p class="text-emerald-700 dark:text-emerald-400 font-medium mt-0.5"><i class="fa-solid fa-images me-1"></i>${totalExtractedImages} full-resolution Table Image(s) (100% sharp) ready to update</p>` : ''}
                        <p class="text-gray-600 dark:text-neutral-300 italic mt-1.5"><span class="font-semibold text-amber-700 dark:text-amber-300 not-italic">Existing Products Preview (${duplicateCount}):</span> ${escapeHtmlAttr(sampleDupText)}</p>
                        <label class="flex items-center gap-2 mt-2 p-2 rounded-lg bg-amber-50 dark:bg-amber-950/40 border border-amber-300 dark:border-amber-700 cursor-pointer">
                            <input type="checkbox" id="chkOverwriteDuplicates" checked data-ev-change="a157" class="w-4 h-4 accent-blue-600 rounded cursor-pointer">
                            <span class="text-xs font-semibold text-gray-800 dark:text-neutral-200">Update / overwrite ${duplicateCount} existing product(s) with the latest data & images</span>
                        </label>
                    `;
            submitBtn.disabled = false;
        } else {
            let newSampleNames = parsedUploadRows.map(r => r[0]);
            let sampleNewText = newSampleNames.length <= 4
                ? newSampleNames.join(", ")
                : newSampleNames.slice(0, 3).join(", ") + ` (+${newSampleNames.length - 3} more)`;

            let dupSampleNames = parsedDuplicateRows.map(r => r[0]);
            let sampleDupText = dupSampleNames.length <= 5
                ? dupSampleNames.join(", ")
                : dupSampleNames.slice(0, 4).join(", ") + ` (+${dupSampleNames.length - 4} more)`;

            previewArea.innerHTML = `
                        <p class="font-bold text-green-700 dark:text-green-400"><i class="fa-solid fa-circle-check"></i> Valid File!</p>
                        <p><strong>Ready to Upload:</strong> ${parsedUploadRows.length} New Product(s) (out of ${allParsed.length} total product(s) in Excel)</p>
                        ${totalExtractedImages > 0 ? `<p class="text-emerald-700 dark:text-emerald-400 font-medium"><i class="fa-solid fa-images me-1"></i>${totalExtractedImages} full-resolution Table Image(s) (100% sharp) successfully extracted</p>` : ''}
                        <p class="text-gray-500 dark:text-neutral-400 italic mt-0.5"><span class="font-semibold text-gray-700 dark:text-neutral-300 not-italic">New Products Preview:</span> ${escapeHtmlAttr(sampleNewText)}</p>
                        ${duplicateCount > 0 ? `
                            <div class="mt-2 pt-2 border-t border-gray-200 dark:border-neutral-600">
                                <p class="text-amber-600 dark:text-amber-400 font-medium"><i class="fa-solid fa-ban me-1"></i>${duplicateCount} Product(s) already exist in the database</p>
                                <p class="text-gray-600 dark:text-neutral-300 italic mt-0.5"><span class="font-semibold text-amber-700 dark:text-amber-300 not-italic">Existing Products Preview (${duplicateCount}):</span> ${escapeHtmlAttr(sampleDupText)}</p>
                                <label class="flex items-center gap-2 mt-1.5 p-2 rounded-lg bg-amber-50 dark:bg-amber-950/40 border border-amber-300 dark:border-amber-700 cursor-pointer">
                                    <input type="checkbox" id="chkOverwriteDuplicates" class="w-4 h-4 accent-blue-600 rounded cursor-pointer">
                                    <span class="text-xs font-semibold text-gray-800 dark:text-neutral-200">Also update / overwrite ${duplicateCount} existing product(s)</span>
                                </label>
                            </div>
                        ` : ''}
                    `;
            submitBtn.disabled = false;
        }

        if (uploadWorkbookCache.workbook.SheetNames.length > 1 && allParsed.length > 0) {
            // Summary of product counts per Target Tab
            const perTarget = {};
            allParsed.forEach(row => {
                const t = resolveSheetTarget(row.__sheetName);
                perTarget[t] = (perTarget[t] || 0) + 1;
            });
            const badges = Object.keys(perTarget).map(t =>
                `<span class="inline-block text-[10px] font-semibold px-2 py-0.5 rounded-full ${TARGET_TAB_BADGE[t] || 'bg-teal-100 text-teal-800 dark:bg-teal-900/50 dark:text-teal-300'}">${escapeHtmlAttr(t)}: ${perTarget[t]} ${perTarget[t] === 1 ? 'product' : 'products'}</span>`
            ).join(" ");
            previewArea.insertAdjacentHTML("afterbegin", `
                            <p class="text-[11px] text-gray-500 dark:text-neutral-400 mb-0.5"><i class="fa-solid fa-layer-group me-1"></i>Selected sheets: <strong>${selectedUploadSheets.map(escapeHtmlAttr).join(", ")}</strong></p>
                            <p class="mb-1 flex flex-wrap items-center gap-1"><span class="text-[11px] text-gray-500 dark:text-neutral-400">Target:</span> ${badges}</p>`);
        }
        previewArea.classList.remove("hidden");

        // Tombol Upload hanya aktif jika file valid DAN semua kartu Review Benefit sudah diputuskan
        uploadSubmitBaseEnabled = !submitBtn.disabled;
        renderBenefitReview();
    } catch (err) {
        console.error(err);
        benefitReviewItems = [];
        renderBenefitReview();
        uploadSubmitBaseEnabled = false;
        previewArea.innerHTML = `<p class="text-red-500"><i class="fa-solid fa-circle-xmark"></i> Failed to process selected sheets: ${escapeHtmlAttr(err.message || err)}</p>`;
        submitBtn.disabled = true;
    }
}

// =====================================================================
// ALIAS NAMA MANFAAT & NON-BENEFIT + REVIEW SAAT UPLOAD EXCEL
// =====================================================================

// Kunci pembanding nama manfaat: huruf kecil, tanpa bullet / titik dua di akhir, spasi rapat, bentuk jamak diabaikan
// (HARUS sama dengan normalizeBenefitAliasKey_() di Apps Script)
function normalizeBenefitNameKey(name) {
    return String(name || "")
        .toLowerCase()
        .replace(/^[\s\-•*]+/, "")
        .replace(/[:：]\s*$/, "")
        .replace(/\s+/g, " ")
        .trim()
        .split(" ")
        .map(w => (w.length > 3 && /s$/.test(w) && !/(ss|us|is)$/.test(w)) ? w.slice(0, -1) : w)
        .join(" ");
}

let SHEET_ALIAS_RULES = [];

function resetDynamicConfigListsToDefault() {
    NON_BENEFIT_COLUMNS = [...DEFAULT_NON_BENEFIT_COLUMNS];
    SYARIAH_SPEC_COLUMNS = [...DEFAULT_SYARIAH_SPEC_COLUMNS];
    DPLK_SPEC_COLUMNS = [...DEFAULT_DPLK_SPEC_COLUMNS];
    NON_BENEFIT_ALIAS_MAP = {};
    NON_BENEFIT_CONFIG_CANONICALS = new Set();
    BENEFIT_ALIAS_MAP = {};
    BENEFIT_CONFIG_CANONICALS = new Set();
    BENEFIT_CONFIG_CANONICAL_LIST = [];
    STANDARD_BENEFIT_COLUMNS = [...DEFAULT_STANDARD_BENEFIT_COLUMNS];
    DPLK_BENEFIT_RULES = DEFAULT_DPLK_BENEFIT_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    DPLK_BENEFIT_COLUMNS = DPLK_BENEFIT_RULES.map(r => r.canonical);
    DPLK_CHARGE_RULES = DEFAULT_DPLK_CHARGE_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    DPLK_CHARGE_COLUMNS = DPLK_CHARGE_RULES.map(r => r.canonical);
    CHANNEL_FILTER_RULES = DEFAULT_CHANNEL_FILTER_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    UNDERWRITING_FILTER_RULES = DEFAULT_UNDERWRITING_FILTER_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    COMPANY_KEYWORD_RULES = DEFAULT_COMPANY_KEYWORD_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    CORE_FIELD_RULES = DEFAULT_CORE_FIELD_RULES.map(r => ({ canonical: r.canonical, aliases: [...r.aliases] }));
    SHEET_ALIAS_RULES = [];
    CATEGORY_CONFIGS = {};
}

function applyDynamicConfigFromBackend(cfg) {
    if (!cfg || typeof cfg !== "object") {
        SHEET_TABS = [...DEFAULT_SHEET_TABS];
        return;
    }

    if (Array.isArray(cfg.categoryConfigs)) {
        CATEGORY_CONFIGS = {};
        cfg.categoryConfigs.forEach(item => {
            if (item && item.sheetName) {
                CATEGORY_CONFIGS[String(item.sheetName).trim()] = structuredClone(item);
            }
        });
    } else {
        CATEGORY_CONFIGS = (cfg.categoryConfigs && typeof cfg.categoryConfigs === "object")
            ? structuredClone(cfg.categoryConfigs)
            : {};
    }

    if (Array.isArray(cfg.tabs) && cfg.tabs.length > 0) {
        SHEET_TABS = cfg.tabs.filter(t => {
            const lw = String(t || "").trim().toLowerCase();
            return lw && lw !== "config" && lw !== "documents" && lw !== "users";
        });
    } else if (!Array.isArray(SHEET_TABS) || SHEET_TABS.length === 0) {
        SHEET_TABS = [...DEFAULT_SHEET_TABS];
    }
    Object.keys(CATEGORY_CONFIGS || {}).forEach(catKey => {
        if (catKey && !SHEET_TABS.some(t => t.toLowerCase() === catKey.toLowerCase())) {
            SHEET_TABS.push(catKey);
        }
    });

    if (Array.isArray(cfg.sheetAliases) && cfg.sheetAliases.length > 0) {
        SHEET_ALIAS_RULES = cfg.sheetAliases
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
    }

    if (Array.isArray(cfg.coreAliases) && cfg.coreAliases.length > 0) {
        CORE_FIELD_RULES = cfg.coreAliases
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
    }

    if (Array.isArray(cfg.syariahNonBenefitColumns) && cfg.syariahNonBenefitColumns.length > 0) {
        SYARIAH_SPEC_COLUMNS = [...new Set([...DEFAULT_SYARIAH_SPEC_COLUMNS, ...cfg.syariahNonBenefitColumns])];
    }
    if (Array.isArray(cfg.dplkNonBenefitColumns) && cfg.dplkNonBenefitColumns.length > 0) {
        DPLK_SPEC_COLUMNS = [...new Set([...DEFAULT_DPLK_SPEC_COLUMNS, ...cfg.dplkNonBenefitColumns])];
    }

    if (Array.isArray(cfg.nonBenefitColumns)) {
        cfg.nonBenefitColumns.forEach(val => {
            if (val && !NON_BENEFIT_COLUMNS.some(c => c.toLowerCase() === val.toLowerCase())) {
                NON_BENEFIT_COLUMNS.push(val);
            }
        });
    }

    if (cfg.driveFolderId) {
        systemConfig.driveFolderId = cfg.driveFolderId;
    }

    loadBenefitAliasesFromConfig(cfg.benefitAliases);
    loadNonBenefitAliasesFromConfig(cfg.nonBenefitAliases);

    if (Array.isArray(cfg.dplkBenefits) && cfg.dplkBenefits.length > 0) {
        DPLK_BENEFIT_RULES = cfg.dplkBenefits
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
        DPLK_BENEFIT_COLUMNS = DPLK_BENEFIT_RULES.map(r => r.canonical);
    }

    if (Array.isArray(cfg.dplkCharges) && cfg.dplkCharges.length > 0) {
        DPLK_CHARGE_RULES = cfg.dplkCharges
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
        DPLK_CHARGE_COLUMNS = DPLK_CHARGE_RULES.map(r => r.canonical);
    }

    if (Array.isArray(cfg.channelFilters) && cfg.channelFilters.length > 0) {
        CHANNEL_FILTER_RULES = cfg.channelFilters
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
    }

    if (Array.isArray(cfg.underwritingFilters) && cfg.underwritingFilters.length > 0) {
        UNDERWRITING_FILTER_RULES = cfg.underwritingFilters
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
    }

    if (Array.isArray(cfg.companyKeywords) && cfg.companyKeywords.length > 0) {
        COMPANY_KEYWORD_RULES = cfg.companyKeywords
            .filter(r => r && r.canonical)
            .map(r => ({ canonical: String(r.canonical).trim(), aliases: Array.isArray(r.aliases) ? r.aliases : [] }));
    }
}

function getConfigAliasesForColumn(colName, fallbackAliases = []) {
    const targetKey = normalizeBenefitNameKey(colName);
    const list = [colName, ...fallbackAliases];
    const addUnique = (val) => {
        const v = String(val || "").trim();
        if (v && !list.some(item => item.toLowerCase() === v.toLowerCase())) {
            list.push(v);
        }
    };

    CORE_FIELD_RULES.forEach(r => {
        if (normalizeBenefitNameKey(r.canonical) === targetKey) {
            addUnique(r.canonical);
            (Array.isArray(r.aliases) ? r.aliases : []).forEach(addUnique);
        }
    });

    DEFAULT_NON_BENEFIT_ALIAS_RULES.forEach(r => {
        if (normalizeBenefitNameKey(r.canonical) === targetKey) {
            addUnique(r.canonical);
            (Array.isArray(r.aliases) ? r.aliases : []).forEach(addUnique);
        }
    });

    Object.keys(NON_BENEFIT_ALIAS_MAP || {}).forEach(aKey => {
        const mapped = NON_BENEFIT_ALIAS_MAP[aKey];
        if (normalizeBenefitNameKey(mapped) === targetKey) {
            addUnique(mapped);
            addUnique(aKey);
        }
    });

    return list;
}

function loadBenefitAliasesFromConfig(aliasRows) {
    BENEFIT_ALIAS_MAP = {};
    BENEFIT_CONFIG_CANONICALS = new Set();
    BENEFIT_CONFIG_CANONICAL_LIST = [];
    STANDARD_BENEFIT_COLUMNS = [...DEFAULT_STANDARD_BENEFIT_COLUMNS];
    if (!Array.isArray(aliasRows)) return;
    aliasRows.forEach(item => {
        const canonical = String((item && item.canonical) || "").trim();
        if (!canonical) return;
        const cKey = normalizeBenefitNameKey(canonical);
        BENEFIT_CONFIG_CANONICALS.add(cKey);
        if (!BENEFIT_CONFIG_CANONICAL_LIST.some(b => b.toLowerCase() === canonical.toLowerCase())) {
            BENEFIT_CONFIG_CANONICAL_LIST.push(canonical);
        }
        if (!STANDARD_BENEFIT_COLUMNS.some(b => b.toLowerCase() === canonical.toLowerCase())) {
            STANDARD_BENEFIT_COLUMNS.push(canonical);
        }
        if (!BENEFIT_ALIAS_MAP[cKey]) BENEFIT_ALIAS_MAP[cKey] = canonical;
        (Array.isArray(item.aliases) ? item.aliases : []).forEach(a => {
            const aKey = normalizeBenefitNameKey(a);
            if (aKey && !BENEFIT_ALIAS_MAP[aKey]) BENEFIT_ALIAS_MAP[aKey] = canonical;
        });
    });
}

function loadNonBenefitAliasesFromConfig(aliasRows, merge = false) {
    if (!merge) {
        NON_BENEFIT_ALIAS_MAP = {};
        NON_BENEFIT_CONFIG_CANONICALS = new Set();
    }
    if (!Array.isArray(aliasRows)) return;
    aliasRows.forEach(item => {
        const canonical = String((item && item.canonical) || "").trim();
        if (!canonical) return;
        const cKey = normalizeBenefitNameKey(canonical);
        NON_BENEFIT_CONFIG_CANONICALS.add(cKey);
        if (!NON_BENEFIT_COLUMNS.some(c => c.toLowerCase() === canonical.toLowerCase())) {
            NON_BENEFIT_COLUMNS.push(canonical);
        }
        if (!NON_BENEFIT_ALIAS_MAP[cKey]) NON_BENEFIT_ALIAS_MAP[cKey] = canonical;
        (Array.isArray(item.aliases) ? item.aliases : []).forEach(a => {
            const aKey = normalizeBenefitNameKey(a);
            if (aKey && !NON_BENEFIT_ALIAS_MAP[aKey]) NON_BENEFIT_ALIAS_MAP[aKey] = canonical;
        });
    });
}

function isBenefitLikeColumn(col) {
    return Boolean(col) && !isNonBenefitColumn(col) && !isDplkChargeColumn(col);
}

// Cari nama kolom manfaat yang sudah dikenal (kolom yang ada / nama standar / alias Config) untuk nama mentah dari Excel
function resolveKnownBenefitName(rawName, extraCandidates = []) {
    const key = normalizeBenefitNameKey(rawName);
    if (!key) return null;
    const candidates = [...dynamicColumns.filter(isBenefitLikeColumn), ...extraCandidates];
    const direct = candidates.find(c => normalizeBenefitNameKey(c) === key);
    if (direct) return direct;

    const aliasCanonical = BENEFIT_ALIAS_MAP[key];
    if (aliasCanonical) {
        const aKey = normalizeBenefitNameKey(aliasCanonical);
        return candidates.find(c => normalizeBenefitNameKey(c) === aKey) || aliasCanonical;
    }
    return null;
}

// Cari nama kolom Non-Benefit yang sudah dikenal (Config / default / alias)
function resolveKnownNonBenefitName(rawName) {
    const key = normalizeBenefitNameKey(rawName);
    if (!key) return null;
    const baseNb = (uploadWorkbookCache && Array.isArray(uploadWorkbookCache.baseNonBenefits))
        ? uploadWorkbookCache.baseNonBenefits
        : NON_BENEFIT_COLUMNS;
    const direct = baseNb.find(c => normalizeBenefitNameKey(c) === key)
        || NON_BENEFIT_COLUMNS.find(c => normalizeBenefitNameKey(c) === key);
    if (direct) return direct;

    const aliasCanonical = NON_BENEFIT_ALIAS_MAP[key];
    if (aliasCanonical) {
        const aKey = normalizeBenefitNameKey(aliasCanonical);
        return baseNb.find(c => normalizeBenefitNameKey(c) === aKey) || aliasCanonical;
    }

    for (let i = 0; i < DEFAULT_NON_BENEFIT_ALIAS_RULES.length; i++) {
        const rule = DEFAULT_NON_BENEFIT_ALIAS_RULES[i];
        if (normalizeBenefitNameKey(rule.canonical) === key || (Array.isArray(rule.aliases) && rule.aliases.some(a => normalizeBenefitNameKey(a) === key))) {
            return rule.canonical;
        }
    }
    return null;
}

function getBaseNonBenefitColumns() {
    return (uploadWorkbookCache && Array.isArray(uploadWorkbookCache.baseNonBenefits))
        ? uploadWorkbookCache.baseNonBenefits
        : NON_BENEFIT_COLUMNS;
}

// Cache lookup O(1) untuk Master Columns & hasil suggestions Review Benefit agar tidak rebuild berulang
let _reviewMasterCache = null;
let _reviewSuggestionsCache = new Map();
let _benefitReviewPeerRefreshTimer = null;

function invalidateReviewLookupCaches() {
    _reviewMasterCache = null;
    _reviewSuggestionsCache.clear();
    if (_benefitReviewPeerRefreshTimer) {
        clearTimeout(_benefitReviewPeerRefreshTimer);
        _benefitReviewPeerRefreshTimer = null;
    }
}

function getReviewMasterLookupCache() {
    const base = (uploadWorkbookCache && Array.isArray(uploadWorkbookCache.baseColumns))
        ? uploadWorkbookCache.baseColumns
        : dynamicColumns;
    const baseNbRef = (uploadWorkbookCache && Array.isArray(uploadWorkbookCache.baseNonBenefits))
        ? uploadWorkbookCache.baseNonBenefits
        : NON_BENEFIT_COLUMNS;
    if (
        _reviewMasterCache &&
        _reviewMasterCache.baseRef === base &&
        _reviewMasterCache.baseLen === base.length &&
        _reviewMasterCache.baseNbLen === baseNbRef.length &&
        _reviewMasterCache.cfgBenLen === BENEFIT_CONFIG_CANONICAL_LIST.length
    ) {
        return _reviewMasterCache;
    }

    const baseNb = getBaseNonBenefitColumns();
    const baseNbSet = new Set(baseNb.map(c => c.toLowerCase()));
    const benefitCols = [];
    const seenBenLower = new Set();
    // Gabungkan Benefit dari Config + header yang sudah ada di Master Sheet (sehingga tetap muncul di dropdown & valid walau tab produk kosong)
    [...BENEFIT_CONFIG_CANONICAL_LIST, ...base].forEach(col => {
        const clean = String(col || "").trim();
        if (!clean) return;
        const low = clean.toLowerCase();
        if (baseNbSet.has(low) || isDplkChargeColumn(clean) || seenBenLower.has(low)) return;
        seenBenLower.add(low);
        benefitCols.push(clean);
    });

    const excluded = new Set(["launch date", "tanggal rilis", "remarks", "product name", "company name", "par/non par", "product type", "distribution channel", "currency"]);
    const mergedNb = [...baseNb];
    const mergedNbLower = new Set(mergedNb.map(m => m.toLowerCase()));
    base.forEach(c => {
        const low = String(c || "").toLowerCase();
        if (baseNbSet.has(low) && !mergedNbLower.has(low)) {
            mergedNb.push(c);
            mergedNbLower.add(low);
        }
    });
    const nonBenefitCols = mergedNb.filter(col => col && !excluded.has(col.toLowerCase()));

    const benefitKeyMap = new Map();
    benefitCols.forEach(c => {
        const k = normalizeBenefitNameKey(c);
        if (k && !benefitKeyMap.has(k)) benefitKeyMap.set(k, c);
    });
    const nonBenefitKeyMap = new Map();
    nonBenefitCols.forEach(c => {
        const k = normalizeBenefitNameKey(c);
        if (k && !nonBenefitKeyMap.has(k)) nonBenefitKeyMap.set(k, c);
    });

    _reviewMasterCache = {
        baseRef: base,
        baseLen: base.length,
        baseNbLen: baseNbRef.length,
        cfgBenLen: BENEFIT_CONFIG_CANONICAL_LIST.length,
        benefitCols,
        nonBenefitCols,
        benefitKeyMap,
        nonBenefitKeyMap
    };
    return _reviewMasterCache;
}

// Kolom manfaat yang SUDAH ADA di Master Sheet (sebelum file Excel ini diproses)
function getExistingMasterBenefitColumns() {
    return getReviewMasterLookupCache().benefitCols;
}

// Kolom Non-Benefit yang SUDAH ADA di Master Sheet / Config (sebelum file Excel ini diproses)
function getExistingMasterNonBenefitColumns() {
    return getReviewMasterLookupCache().nonBenefitCols;
}

function getExistingMasterColumnsByKind(kind) {
    return kind === "non-benefit" ? getExistingMasterNonBenefitColumns() : getExistingMasterBenefitColumns();
}

// Cocokkan nama ketikan user ke kolom yang sudah ada di Master Sheet (O(1) lookup langsung atau lewat alias Config)
function findExistingMasterColumnByKind(name, kind = "benefit") {
    const key = normalizeBenefitNameKey(name);
    if (!key) return null;
    const cache = getReviewMasterLookupCache();
    const keyMap = kind === "non-benefit" ? cache.nonBenefitKeyMap : cache.benefitKeyMap;
    const direct = keyMap.get(key);
    if (direct) return direct;
    const aliasMap = kind === "non-benefit" ? NON_BENEFIT_ALIAS_MAP : BENEFIT_ALIAS_MAP;
    const aliasCanonical = aliasMap[key];
    if (aliasCanonical) {
        const aKey = normalizeBenefitNameKey(aliasCanonical);
        return keyMap.get(aKey) || null;
    }
    return null;
}

function findExistingMasterBenefit(name) {
    return findExistingMasterColumnByKind(name, "benefit");
}

function bigramDice(a, b) {
    if (a === b) return 1;
    if (a.length < 2 || b.length < 2) return 0;
    const grams = s => { const m = new Map(); for (let i = 0; i < s.length - 1; i++) { const g = s.substr(i, 2); m.set(g, (m.get(g) || 0) + 1); } return m; };
    const ga = grams(a), gb = grams(b);
    let overlap = 0;
    ga.forEach((cnt, g) => { overlap += Math.min(cnt, gb.get(g) || 0); });
    return (2 * overlap) / (a.length - 1 + b.length - 1);
}

// Skor kemiripan 0..1 antar nama manfaat (per kata; singkatan seperti "Add" ~ "Additional" dihitung mirip)
function benefitNameSimilarity(a, b) {
    const ta = normalizeBenefitNameKey(a).split(/[^a-z0-9]+/).filter(Boolean);
    const tb = normalizeBenefitNameKey(b).split(/[^a-z0-9]+/).filter(Boolean);
    if (!ta.length || !tb.length) return 0;
    const tokenScore = (x, y) => {
        if (x === y) return 1;
        if (x.length >= 2 && y.length >= 2 && (x.startsWith(y) || y.startsWith(x))) return 0.85;
        return bigramDice(x, y);
    };
    const dir = (A, B) => A.reduce((sum, x) => sum + Math.max(...B.map(y => tokenScore(x, y))), 0) / A.length;
    return (dir(ta, tb) + dir(tb, ta)) / 2;
}

function isValidNewTargetNameForKind(name, kind) {
    const trimmed = String(name || "").trim();
    if (!trimmed || isDplkChargeColumn(trimmed)) return false;
    const key = normalizeBenefitNameKey(trimmed);
    if (!key) return false;
    if (kind === "benefit") {
        if (findExistingMasterColumnByKind(trimmed, "non-benefit")) return false;
    } else {
        if (findExistingMasterColumnByKind(trimmed, "benefit")) return false;
    }
    return true;
}

// Ambil target "New" dari kartu review lain agar kartu berikutnya otomatis tervalidasi terhadap nama baru tsb (Poin 3)
function getPeerNewTargetsInReview(excludeKey, kind, onlyEarlier = false) {
    const seen = new Map(); // normalizedKey -> { name, sourceCol }
    for (const item of benefitReviewItems) {
        if (item.key === excludeKey) {
            if (onlyEarlier) break;
            continue;
        }
        const d = benefitReviewDecisions[item.key];
        if (!d || d.mode !== "new") continue;
        const dKind = d.kind || item.detectedKind || "benefit";
        if (dKind !== kind) continue;
        const targetName = String(d.name || "").trim();
        if (!isValidNewTargetNameForKind(targetName, kind)) continue;
        if (findExistingMasterColumnByKind(targetName, kind)) continue;
        const tKey = normalizeBenefitNameKey(targetName);
        if (!seen.has(tKey)) {
            seen.set(tKey, { name: targetName, sourceCol: item.col });
        }
    }
    return Array.from(seen.values());
}

function getReviewSuggestions(name, kind, excludeKey = null, limit = 3) {
    const masterCandidates = getExistingMasterColumnsByKind(kind);
    const peerCandidates = getPeerNewTargetsInReview(excludeKey, kind).map(p => p.name);
    const cacheKey = `${kind}|${limit}|${peerCandidates.join("\u0001")}|${name}`;
    if (_reviewSuggestionsCache.has(cacheKey)) {
        return _reviewSuggestionsCache.get(cacheKey);
    }
    const allCandidates = peerCandidates.length > 0
        ? [...new Set([...masterCandidates, ...peerCandidates])]
        : masterCandidates;
    const res = allCandidates
        .map(col => ({ col, score: benefitNameSimilarity(name, col) }))
        .filter(x => x.score >= 0.4)
        .sort((a, b) => b.score - a.score)
        .slice(0, limit)
        .map(x => x.col);
    if (_reviewSuggestionsCache.size > 500) _reviewSuggestionsCache.clear();
    _reviewSuggestionsCache.set(cacheKey, res);
    return res;
}

function getBenefitSuggestions(name, limit = 3) {
    return getReviewSuggestions(name, "benefit", null, limit);
}

function stripImageTokensForSnippet(val) {
    return String(val || "").replace(/\[IMG(?:_ID)?:[^\]]*\]/g, "").trim();
}

function extractImageTokensFromCell(val) {
    return (String(val || "").match(/\[IMG(?:_ID)?:[^\]]*\]/g) || []);
}

// Kumpulkan nama Benefit DAN Non-Benefit dari Excel yang belum ada di Master Sheet maupun di Config
function buildBenefitReviewItems(allParsed) {
    invalidateReviewLookupCaches();
    benefitReviewItems = [];
    if (!uploadParsedSchemaStructured || !allParsed || allParsed.length === 0) return;

    const existingBenefitKeys = new Set(getExistingMasterBenefitColumns().map(normalizeBenefitNameKey));
    const existingNonBenefitKeys = new Set(getExistingMasterNonBenefitColumns().map(normalizeBenefitNameKey));
    const schema = uploadParsedSchema;

    schema.slice(7).forEach(col => {
        if (!col || isDplkChargeColumn(col)) return;
        const lowerCol = col.toLowerCase();
        if (STANDARD_FIXED_SPEC_COLS.has(lowerCol)) return;

        const isNb = isNonBenefitColumn(col);
        const key = normalizeBenefitNameKey(col);
        if (!key) return;

        if (isNb) {
            if (existingNonBenefitKeys.has(key) || NON_BENEFIT_CONFIG_CANONICALS.has(key)) return;
        } else {
            if (existingBenefitKeys.has(key) || BENEFIT_CONFIG_CANONICALS.has(key)) return;
        }

        const colIdx = schema.indexOf(col);
        const usedBy = [];
        allParsed.forEach(row => {
            const v = row[colIdx];
            if (!isEmptyVal(v)) usedBy.push({ name: String(row[0] || "-"), value: String(v) });
        });
        if (usedBy.length === 0) return;

        const detectedKind = isNb ? "non-benefit" : "benefit";
        if (!benefitReviewDecisions[key]) {
            benefitReviewDecisions[key] = { mode: "pending", name: col, kind: detectedKind };
        } else if (!benefitReviewDecisions[key].kind) {
            benefitReviewDecisions[key].kind = detectedKind;
        }
        const activeKind = benefitReviewDecisions[key].kind || detectedKind;
        benefitReviewItems.push({
            col,
            key,
            detectedKind,
            usedBy,
            suggestions: getReviewSuggestions(col, activeKind, key)
        });
    });
}

// Evaluasi keputusan user untuk satu kartu review (mendukung Benefit, Non-Benefit, serta referensi ke New Benefit/Non-Benefit lain)
function evaluateBenefitReviewItem(item) {
    const d = benefitReviewDecisions[item.key] || { mode: "pending", name: item.col, kind: item.detectedKind || "benefit" };
    const kind = d.kind || item.detectedKind || "benefit";
    const kindLabel = kind === "non-benefit" ? "Non-Benefit" : "Benefit";
    const oppLabel = kind === "non-benefit" ? "Benefit" : "Non-Benefit Provision";
    const name = String(d.name || "").trim();

    if (d.mode === "ignore") {
        return { valid: true, action: "ignore", target: null, kind, tone: "gray", msg: `Skipped — this ${kindLabel.toLowerCase()} will not be imported.` };
    }
    if (!name) {
        return { valid: false, kind, tone: "red", msg: `${kindLabel} name cannot be empty.` };
    }
    if (isDplkChargeColumn(name)) {
        return { valid: false, kind, tone: "red", msg: `<strong>"${escapeHtmlAttr(name)}"</strong> is a Charge component, not a ${kindLabel}.` };
    }

    if (kind === "benefit" && findExistingMasterColumnByKind(name, "non-benefit")) {
        return { valid: false, kind, tone: "red", msg: `<strong>"${escapeHtmlAttr(name)}"</strong> is a Non-Benefit Provision column. Switch type to <strong>Non-Benefit</strong> or pick a Benefit name.` };
    }
    if (kind === "non-benefit" && findExistingMasterColumnByKind(name, "benefit")) {
        return { valid: false, kind, tone: "red", msg: `<strong>"${escapeHtmlAttr(name)}"</strong> already exists in Product Benefits. Switch type to <strong>Benefit</strong> or pick a Provision name.` };
    }

    const existing = findExistingMasterColumnByKind(name, kind);
    if (existing) {
        const willAlias = normalizeBenefitNameKey(item.col) !== normalizeBenefitNameKey(existing);
        return {
            valid: true, action: "map", target: existing, kind, tone: "green",
            msg: `Will be merged into existing ${kindLabel} <strong>"${escapeHtmlAttr(existing)}"</strong>${willAlias ? ` &mdash; <em>"${escapeHtmlAttr(item.col)}"</em> will be saved as its alias in Config.` : "."}`
        };
    }

    // Cek apakah mengacu ke "New Benefit / New Non-Benefit" dari kartu lain (terutama kartu yang lebih dulu ditandai New)
    const nameKey = normalizeBenefitNameKey(name);
    const earlierPeer = getPeerNewTargetsInReview(item.key, kind, true).find(p => normalizeBenefitNameKey(p.name) === nameKey);
    const anyPeer = earlierPeer || (d.mode !== "new" ? getPeerNewTargetsInReview(item.key, kind, false).find(p => normalizeBenefitNameKey(p.name) === nameKey) : null);

    if (anyPeer) {
        const willAlias = normalizeBenefitNameKey(item.col) !== nameKey;
        return {
            valid: true, action: "map", target: anyPeer.name, kind, tone: "green",
            msg: `Will be merged into new ${kindLabel} <strong>"${escapeHtmlAttr(anyPeer.name)}"</strong> (from <em>"${escapeHtmlAttr(anyPeer.sourceCol)}"</em>)${willAlias ? ` &mdash; <em>"${escapeHtmlAttr(item.col)}"</em> saved as alias in Config.` : "."}`
        };
    }

    if (d.mode === "new") {
        const renamed = normalizeBenefitNameKey(item.col) !== nameKey;
        return {
            valid: true, action: "new", target: name, kind, tone: "amber",
            msg: `New ${kindLabel} column <strong>"${escapeHtmlAttr(name)}"</strong> will be added to the Master Sheet${kind === "non-benefit" || renamed ? ` &mdash; registered in Config.` : "."}`
        };
    }

    return {
        valid: false, kind, tone: "red",
        msg: `<strong>"${escapeHtmlAttr(name)}"</strong> does not exist in the Master Sheet (${kindLabel}). Pick an existing ${kindLabel}, match another new one, or tick <strong>New ${kindLabel}</strong>.`
    };
}

function isBenefitReviewAllValid() {
    return benefitReviewItems.every(item => evaluateBenefitReviewItem(item).valid);
}

function refreshUploadSubmitState() {
    const submitBtn = document.getElementById("btnSubmitUpload");
    if (!submitBtn) return;
    submitBtn.disabled = !(uploadSubmitBaseEnabled && isBenefitReviewAllValid());
}

const BENEFIT_REVIEW_TONE_CLASSES = {
    green: "bg-emerald-50 text-emerald-700 border-emerald-200 dark:bg-emerald-900/30 dark:text-emerald-300 dark:border-emerald-800",
    amber: "bg-amber-50 text-amber-700 border-amber-200 dark:bg-amber-900/30 dark:text-amber-300 dark:border-amber-800",
    red: "bg-red-50 text-red-600 border-red-200 dark:bg-red-900/30 dark:text-red-300 dark:border-red-800",
    gray: "bg-gray-100 text-gray-600 border-gray-200 dark:bg-neutral-800 dark:text-neutral-400 dark:border-neutral-700"
};
const BENEFIT_REVIEW_TONE_ICONS = { green: "fa-circle-check", amber: "fa-wand-magic-sparkles", red: "fa-triangle-exclamation", gray: "fa-eye-slash" };

function refreshReviewDatalists() {
    const benList = document.getElementById("benefitReviewDatalist");
    const nbList = document.getElementById("nonBenefitReviewDatalist");
    if (benList) {
        const benPeers = getPeerNewTargetsInReview(null, "benefit", false).map(p => p.name);
        benList.innerHTML = [...new Set([...getExistingMasterBenefitColumns(), ...benPeers])]
            .sort((a, b) => a.localeCompare(b))
            .map(c => `<option value="${escapeHtmlAttr(c)}"></option>`)
            .join("");
    }
    if (nbList) {
        const nbPeers = getPeerNewTargetsInReview(null, "non-benefit", false).map(p => p.name);
        nbList.innerHTML = [...new Set([...getExistingMasterNonBenefitColumns(), ...nbPeers])]
            .sort((a, b) => a.localeCompare(b))
            .map(c => `<option value="${escapeHtmlAttr(c)}"></option>`)
            .join("");
    }
}

function refreshReviewCardSuggestions(idx) {
    const item = benefitReviewItems[idx];
    const suggEl = document.getElementById(`benefitReviewSuggestions_${idx}`);
    if (!item || !suggEl) return;
    const d = benefitReviewDecisions[item.key] || { mode: "pending", name: item.col, kind: item.detectedKind };
    const kind = d.kind || item.detectedKind || "benefit";
    if (d.mode === "ignore") {
        suggEl.replaceChildren();
        return;
    }
    const suggestions = getReviewSuggestions(item.col, kind, item.key);
    item.suggestions = suggestions;
    if (suggestions.length === 0) {
        suggEl.replaceChildren();
        return;
    }
    suggEl.innerHTML = `
                <div class="flex flex-wrap items-center gap-1 mt-1.5">
                    <span class="text-[10px] text-gray-500 dark:text-neutral-400">Similar:</span>
                    ${suggestions.map((s, sIdx) =>
        `<button type="button" data-ev-click="a158" data-ev-click-args="${idcmPackArgs([idx, sIdx])}" class="text-[10px] font-semibold px-2 py-0.5 rounded-md bg-blue-100 hover:bg-blue-200 text-blue-700 dark:bg-neutral-700/70 dark:hover:bg-neutral-600 dark:text-blue-300 transition">${escapeHtmlAttr(s)}</button>`
    ).join("")}
                </div>`;
}

function updateBenefitReviewCardStatus(idx, options = {}) {
    const item = benefitReviewItems[idx];
    const statusEl = document.getElementById(`benefitReviewStatus_${idx}`);
    const cardEl = document.getElementById(`benefitReviewCard_${idx}`);
    if (!item || !statusEl) return;
    const ev = evaluateBenefitReviewItem(item);
    statusEl.className = `mt-1.5 text-[11px] px-2.5 py-1 rounded-md border font-medium ${BENEFIT_REVIEW_TONE_CLASSES[ev.tone] || BENEFIT_REVIEW_TONE_CLASSES.gray}`;
    statusEl.innerHTML = `<i class="fa-solid ${BENEFIT_REVIEW_TONE_ICONS[ev.tone] || "fa-circle-info"} me-1"></i>${ev.msg}`;
    if (cardEl) {
        cardEl.classList.toggle("border-red-300", !ev.valid);
        cardEl.classList.toggle("dark:border-red-700", !ev.valid);
    }
    if (!options.skipSummary) {
        updateBenefitReviewSummary();
        refreshUploadSubmitState();
    }
}

function refreshAllReviewCardsDynamicUI() {
    refreshReviewDatalists();
    benefitReviewItems.forEach((_, idx) => {
        refreshReviewCardSuggestions(idx);
        updateBenefitReviewCardStatus(idx, { skipSummary: true });
    });
    updateBenefitReviewSummary();
    refreshUploadSubmitState();
}

let benefitReviewShowPendingOnly = false;

function toggleBenefitReviewPendingOnly() {
    benefitReviewShowPendingOnly = !benefitReviewShowPendingOnly;
    applyBenefitReviewVisibilityFilter();
}

function applyBenefitReviewVisibilityFilter() {
    let pendingCount = 0;
    benefitReviewItems.forEach((item, idx) => {
        const isValid = evaluateBenefitReviewItem(item).valid;
        if (!isValid) pendingCount++;
        const cardEl = document.getElementById(`benefitReviewCard_${idx}`);
        if (cardEl) {
            cardEl.classList.toggle("hidden", benefitReviewShowPendingOnly && isValid);
        }
    });

    const btn = document.getElementById("btnTogglePendingOnlyReview");
    if (btn) {
        btn.className = benefitReviewShowPendingOnly
            ? "text-[11px] font-bold px-2.5 py-1 rounded-md bg-red-600 hover:bg-red-700 text-white transition flex items-center gap-1 shadow-sm"
            : "text-[11px] font-semibold px-2.5 py-1 rounded-md border border-red-300 dark:border-red-700/80 text-red-700 dark:text-red-300 bg-white dark:bg-neutral-800 hover:bg-red-50 dark:hover:bg-neutral-700 transition flex items-center gap-1";
        btn.innerHTML = `<i class="fa-solid ${benefitReviewShowPendingOnly ? "fa-filter-circle-xmark" : "fa-filter"}"></i> ${benefitReviewShowPendingOnly ? "Showing Pending" : "Pending Only"} (${pendingCount})`;
    }

    const emptyBanner = document.getElementById("benefitReviewPendingEmpty");
    if (emptyBanner) {
        emptyBanner.classList.toggle("hidden", !(benefitReviewShowPendingOnly && pendingCount === 0 && benefitReviewItems.length > 0));
    }
}

function updateBenefitReviewSummary() {
    const el = document.getElementById("benefitReviewSummary");
    if (!el) return;
    const total = benefitReviewItems.length;
    const done = benefitReviewItems.filter(item => evaluateBenefitReviewItem(item).valid).length;
    el.className = `text-[11px] font-semibold ${done === total ? "text-emerald-700 dark:text-emerald-400" : "text-red-600 dark:text-red-400"}`;
    el.innerHTML = done === total
        ? `<i class="fa-solid fa-circle-check me-1"></i>All ${total} item(s) reviewed. Ready to upload.`
        : `<i class="fa-solid fa-hourglass-half me-1"></i>${done} of ${total} reviewed &mdash; resolve the remaining card(s) to enable upload.`;
    applyBenefitReviewVisibilityFilter();
}

function bulkSetBenefitReviewMode(mode) {
    benefitReviewItems.forEach(item => {
        const prev = benefitReviewDecisions[item.key] || { mode: "pending", name: item.col, kind: item.detectedKind || "benefit" };
        if (mode === "pending") {
            benefitReviewDecisions[item.key] = { mode: "pending", name: item.col, kind: item.detectedKind || "benefit", valueOverrides: {}, activeProdIdx: 0 };
        } else {
            benefitReviewDecisions[item.key] = {
                mode,
                name: String(prev.name || "").trim() ? prev.name : item.col,
                kind: prev.kind || item.detectedKind || "benefit",
                valueOverrides: prev.valueOverrides || {},
                activeProdIdx: prev.activeProdIdx || 0
            };
        }
    });
    preserveUploadScroll(() => renderBenefitReview());
}

function onBenefitReviewKindChange(idx, nextKind) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    benefitReviewDecisions[item.key].kind = nextKind;
    preserveUploadScroll(() => renderBenefitReview());
}

function renderBenefitReview() {
    const area = document.getElementById("uploadBenefitReviewArea");
    if (!area) return;
    if (benefitReviewItems.length === 0) {
        area.replaceChildren();
        area.classList.add("hidden");
        refreshUploadSubmitState();
        return;
    }

    const cardsHtml = benefitReviewItems.map((item, idx) => {
        const d = benefitReviewDecisions[item.key];
        const kind = d.kind || item.detectedKind || "benefit";
        const isNb = kind === "non-benefit";
        const kindLabel = isNb ? "Non-Benefit" : "Benefit";
        const listId = isNb ? "nonBenefitReviewDatalist" : "benefitReviewDatalist";
        const isIgnored = d.mode === "ignore";
        const isNew = d.mode === "new";

        const maxBadges = 6;
        const badges = item.usedBy.slice(0, maxBadges).map(u =>
            `<span class="inline-block text-[10px] font-semibold px-2 py-0.5 rounded-full bg-sky-100 text-sky-800 dark:bg-neutral-700/70 dark:text-sky-300">${escapeHtmlAttr(u.name)}</span>`
        ).join(" ") + (item.usedBy.length > maxBadges
            ? ` <span class="text-[10px] text-gray-500 dark:text-neutral-400">+${item.usedBy.length - maxBadges} more</span>`
            : "");

        const activeProdIdx = Math.min(Math.max(0, d.activeProdIdx || 0), Math.max(0, item.usedBy.length - 1));
        const activeProd = item.usedBy[activeProdIdx] || { name: "-", value: "" };
        const origDescText = stripImageTokensForSnippet(activeProd.value);
        const hasOverride = d.valueOverrides && Object.prototype.hasOwnProperty.call(d.valueOverrides, activeProd.name);
        const currentDescText = hasOverride ? d.valueOverrides[activeProd.name] : origDescText;
        const activeImgCount = extractImageTokensFromCell(activeProd.value).length;
        const isDescEdited = hasOverride && currentDescText !== origDescText;

        const prodSelectHtml = item.usedBy.length > 1
            ? `<select id="benefitReviewDescProd_${idx}" ${isIgnored ? "disabled" : ""}
                               data-ev-change="a159" data-ev-change-args="${idcmPackArgs([idx])}"
                               class="text-[10px] border dark:border-neutral-600 rounded px-1.5 py-0.5 bg-white dark:bg-neutral-700 text-gray-700 dark:text-neutral-200 outline-none max-w-[200px] truncate disabled:opacity-50">
                            ${item.usedBy.map((u, uIdx) => {
                const uOrig = stripImageTokensForSnippet(u.value);
                const uEdited = d.valueOverrides && Object.prototype.hasOwnProperty.call(d.valueOverrides, u.name) && d.valueOverrides[u.name] !== uOrig;
                return `<option value="${uIdx}" ${uIdx === activeProdIdx ? "selected" : ""}>${uEdited ? "✎ " : ""}${escapeHtmlAttr(u.name)}</option>`;
            }).join("")}
                       </select>`
            : "";

        return `
                    <div id="benefitReviewCard_${idx}" class="p-3 rounded-xl border bg-white dark:bg-neutral-800/80 border-gray-200 dark:border-neutral-600 transition ${isIgnored ? "opacity-60" : ""}">
                        <div class="flex flex-wrap items-center justify-between gap-1.5 mb-1.5">
                            <div class="flex flex-wrap items-center gap-1 min-w-0">
                                <span class="text-[10px] font-bold uppercase tracking-wider text-gray-500 dark:text-neutral-400 me-0.5">${item.usedBy.length > 1 ? "Products" : "Product"}:</span>
                                ${badges}
                            </div>
                            <button type="button" ${isIgnored ? "disabled" : ""}
                                    data-ev-click="a160" data-ev-click-args="${idcmPackArgs([idx, isNb ? "benefit" : "non-benefit"])}"
                                    class="text-[10px] font-bold px-2 py-0.5 rounded-full border transition shrink-0 disabled:opacity-50 ${isNb ? "bg-slate-100 text-slate-700 border-slate-300 hover:bg-slate-200 dark:bg-neutral-700 dark:text-neutral-200 dark:border-neutral-600" : "bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100 dark:bg-neutral-700/70 dark:text-blue-300 dark:border-neutral-600"}"
                                    title="Click to switch between Benefit and Non-Benefit Provision">
                                ${isNb ? "Non-Benefit (Provision)" : "Benefit"} &rlarr;
                            </button>
                        </div>
                        <div class="flex items-center justify-between gap-2 mb-1">
                            <span class="text-[11px] text-gray-500 dark:text-neutral-400 truncate" title="${escapeHtmlAttr(item.col)}">Detected: <strong class="text-gray-700 dark:text-neutral-200">"${escapeHtmlAttr(item.col)}"</strong></span>
                            <div class="flex items-center gap-2 shrink-0">
                                <label class="inline-flex items-center gap-1 text-[11px] font-semibold text-amber-700 dark:text-amber-300 cursor-pointer select-none ${isIgnored ? "pointer-events-none opacity-50" : ""}">
                                    <input type="checkbox" ${isNew ? "checked" : ""} ${isIgnored ? "disabled" : ""} data-ev-change="a161" data-ev-change-args="${idcmPackArgs([idx])}" class="w-3.5 h-3.5 accent-amber-500 rounded cursor-pointer">
                                    New ${kindLabel}
                                </label>
                                <button type="button" data-ev-click="a162" data-ev-click-args="${idcmPackArgs([idx, isIgnored ? "pending" : "ignore"])}" class="text-[11px] font-semibold px-2 py-0.5 rounded-md transition ${isIgnored ? "bg-gray-200 dark:bg-neutral-700 text-gray-700 dark:text-neutral-200" : "text-gray-400 hover:text-red-500 dark:hover:text-red-400"}">
                                    <i class="fa-solid ${isIgnored ? "fa-rotate-left" : "fa-eye-slash"} me-0.5"></i>${isIgnored ? "Undo" : "Ignore"}
                                </button>
                            </div>
                        </div>
                        <input type="text" list="${listId}" value="${escapeHtmlAttr(d.name)}" ${isIgnored ? "disabled" : ""}
                               data-ev-input="a163" data-ev-input-args="${idcmPackArgs([idx])}"
                               placeholder="Type or select an existing ${kindLabel.toLowerCase()} name..."
                               class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 rounded-lg p-2 text-xs font-bold focus:ring-2 focus:ring-blue-500 outline-none disabled:opacity-60">
                        <div id="benefitReviewSuggestions_${idx}"></div>
                        <div id="benefitReviewStatus_${idx}"></div>
                        <div class="mt-2 pt-1.5 border-t border-gray-100 dark:border-neutral-700/80">
                            <div class="flex flex-wrap items-center justify-between gap-1.5 mb-1">
                                <div class="flex items-center gap-1.5 flex-wrap">
                                    <span class="text-[10px] font-semibold text-gray-500 dark:text-neutral-400"><i class="fa-solid fa-pen-to-square me-0.5"></i>Description:</span>
                                    ${prodSelectHtml}
                                    <span id="benefitReviewDescImgBadge_${idx}" class="${activeImgCount > 0 ? "" : "hidden"} text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 dark:bg-emerald-900/40 text-emerald-700 dark:text-emerald-300 font-medium"><i class="fa-solid fa-image me-0.5"></i>+${activeImgCount} img preserved</span>
                                </div>
                                <div id="benefitReviewDescMeta_${idx}" class="flex items-center gap-1.5 ${isDescEdited ? "" : "hidden"}">
                                    <span class="text-[10px] font-semibold text-blue-600 dark:text-blue-400">Edited</span>
                                    <button type="button" ${isIgnored ? "disabled" : ""} data-ev-click="a164" data-ev-click-args="${idcmPackArgs([idx])}" class="text-[10px] text-gray-400 hover:text-red-500 dark:hover:text-red-400 underline">Reset</button>
                                </div>
                            </div>
                            <textarea id="benefitReviewDesc_${idx}" rows="2" ${isIgnored ? "disabled" : ""}
                                      data-ev-input="a165" data-ev-input-args="${idcmPackArgs([idx])}"
                                      placeholder="Edit benefit description..."
                                      class="w-full text-[11px] text-gray-700 dark:text-neutral-200 bg-gray-50 dark:bg-neutral-900/50 border border-gray-200 dark:border-neutral-600 rounded-md p-2 resize-y focus:ring-2 focus:ring-blue-500 focus:bg-white dark:focus:bg-neutral-800 outline-none disabled:opacity-60 leading-relaxed">${escapeHtml(currentDescText)}</textarea>
                        </div>
                    </div>`;
    }).join("");

    const nbCount = benefitReviewItems.filter(i => (benefitReviewDecisions[i.key]?.kind || i.detectedKind) === "non-benefit").length;
    const benCount = benefitReviewItems.length - nbCount;
    const titleParts = [];
    if (benCount > 0) titleParts.push(`${benCount} benefit(s)`);
    if (nbCount > 0) titleParts.push(`${nbCount} non-benefit provision(s)`);

    area.innerHTML = `
                <div class="rounded-lg border border-amber-300 dark:border-amber-700 bg-amber-50/60 dark:bg-amber-950/30 p-3 space-y-2">
                    <div class="flex flex-wrap items-center justify-between gap-2">
                        <p class="text-xs font-bold text-amber-700 dark:text-amber-300"><i class="fa-solid fa-list-check me-1"></i>Review ${titleParts.join(" &amp; ")}</p>
                        <div class="flex flex-wrap items-center gap-1.5">
                            <button type="button" id="btnTogglePendingOnlyReview" data-ev-click="a166" class="text-[11px] font-semibold px-2.5 py-1 rounded-md border border-red-300 dark:border-red-700/80 text-red-700 dark:text-red-300 bg-white dark:bg-neutral-800 hover:bg-red-50 dark:hover:bg-neutral-700 transition flex items-center gap-1" title="Show only unresolved (red) review cards">
                                <i class="fa-solid fa-filter"></i> Pending Only
                            </button>
                            <button type="button" data-ev-click="a167" class="text-[11px] font-bold px-2.5 py-1 rounded-md bg-amber-500 hover:bg-amber-600 text-white transition flex items-center gap-1 shadow-sm" title="Mark all items as New Benefit / New Non-Benefit">
                                <i class="fa-solid fa-check-double"></i> Select All New
                            </button>
                            <button type="button" data-ev-click="a168" class="text-[11px] font-bold px-2.5 py-1 rounded-md bg-gray-600 hover:bg-gray-700 text-white transition flex items-center gap-1 shadow-sm" title="Ignore all unrecognized items">
                                <i class="fa-solid fa-eye-slash"></i> Ignore All
                            </button>
                            <button type="button" data-ev-click="a169" class="text-[11px] font-semibold px-2 py-1 rounded-md border border-gray-300 dark:border-neutral-600 text-gray-600 dark:text-neutral-300 hover:bg-white dark:hover:bg-neutral-700 transition" title="Reset all review cards">
                                Reset
                            </button>
                        </div>
                    </div>
                    <p class="text-[11px] text-gray-600 dark:text-neutral-300 leading-relaxed">
                        Map to an <strong>existing column</strong> (or to another <strong>New</strong> card above), tick <strong>New Benefit / Non-Benefit</strong> to add a new column, or <strong>Ignore</strong> to skip.
                    </p>
                    <p id="benefitReviewSummary"></p>
                    <div id="benefitReviewPendingEmpty" class="hidden p-3 rounded-lg bg-emerald-50 dark:bg-neutral-800 border border-emerald-200 dark:border-emerald-700/70 text-emerald-700 dark:text-emerald-300 text-xs font-semibold text-center">
                        <i class="fa-solid fa-circle-check me-1"></i>All items resolved! <button type="button" data-ev-click="a166" class="underline ml-1">Show all cards</button>
                    </div>
                    <div class="space-y-2">${cardsHtml}</div>
                    <datalist id="benefitReviewDatalist"></datalist>
                    <datalist id="nonBenefitReviewDatalist"></datalist>
                </div>`;
    area.classList.remove("hidden");

    refreshAllReviewCardsDynamicUI();
}

function onBenefitReviewNameInput(idx, value) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    const d = benefitReviewDecisions[item.key];
    d.name = value;
    // Langsung update kartu aktif + summary (O(1) DOM update, tanpa lag saat mengetik)
    updateBenefitReviewCardStatus(idx);
    // Jika kartu ini berstatus "new", kartu lain bisa merujuk ke nama barunya -> refresh peer dengan debounce
    if (d.mode === "new") {
        if (_benefitReviewPeerRefreshTimer) clearTimeout(_benefitReviewPeerRefreshTimer);
        _benefitReviewPeerRefreshTimer = setTimeout(() => {
            _benefitReviewPeerRefreshTimer = null;
            refreshAllReviewCardsDynamicUI();
        }, 180);
    }
}

function onBenefitReviewModeChange(idx, mode) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    benefitReviewDecisions[item.key].mode = mode;
    preserveUploadScroll(() => renderBenefitReview());
}

function onBenefitReviewPickSuggestion(idx, sIdx) {
    const item = benefitReviewItems[idx];
    if (!item || !item.suggestions[sIdx]) return;
    const pickedName = item.suggestions[sIdx];
    const prev = benefitReviewDecisions[item.key] || {};
    const prevKind = prev.kind || item.detectedKind || "benefit";
    const nextOverrides = { ...(prev.valueOverrides || {}) };
    // Jika deskripsi produk belum diedit manual dan baris pertamanya adalah judul duplikat/alias dari nama target baru, bersihkan otomatis
    if (prevKind === "benefit" && Array.isArray(item.usedBy)) {
        item.usedBy.forEach(prod => {
            if (!prod || Object.prototype.hasOwnProperty.call(nextOverrides, prod.name)) return;
            const origText = stripImageTokensForSnippet(prod.value);
            const cleaned = stripDuplicateLeadingBenefitTitle(pickedName, origText, item.originalName);
            if (cleaned && cleaned !== origText) {
                nextOverrides[prod.name] = cleaned;
            }
        });
    }
    benefitReviewDecisions[item.key] = {
        mode: "pending",
        name: pickedName,
        kind: prevKind,
        valueOverrides: nextOverrides,
        activeProdIdx: prev.activeProdIdx || 0
    };
    preserveUploadScroll(() => renderBenefitReview());
}

function onBenefitReviewDescInput(idx, value) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    const d = benefitReviewDecisions[item.key];
    if (!d) return;
    if (!d.valueOverrides) d.valueOverrides = {};
    const pIdx = Math.min(Math.max(0, d.activeProdIdx || 0), Math.max(0, item.usedBy.length - 1));
    const prod = item.usedBy[pIdx];
    if (!prod) return;
    const origText = stripImageTokensForSnippet(prod.value);
    if (value === origText) {
        delete d.valueOverrides[prod.name];
    } else {
        d.valueOverrides[prod.name] = value;
    }
    const isEdited = Object.prototype.hasOwnProperty.call(d.valueOverrides, prod.name);
    const metaEl = document.getElementById(`benefitReviewDescMeta_${idx}`);
    if (metaEl) metaEl.classList.toggle("hidden", !isEdited);
    const selEl = document.getElementById(`benefitReviewDescProd_${idx}`);
    if (selEl && selEl.options[pIdx]) {
        selEl.options[pIdx].textContent = (isEdited ? "✎ " : "") + prod.name;
    }
}

function onBenefitReviewDescProductChange(idx, prodIdxRaw) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    const d = benefitReviewDecisions[item.key];
    if (!d) return;
    const pIdx = Math.min(Math.max(0, parseInt(prodIdxRaw, 10) || 0), Math.max(0, item.usedBy.length - 1));
    d.activeProdIdx = pIdx;
    const prod = item.usedBy[pIdx];
    if (!prod) return;
    const origText = stripImageTokensForSnippet(prod.value);
    const hasOverride = d.valueOverrides && Object.prototype.hasOwnProperty.call(d.valueOverrides, prod.name);
    const currentText = hasOverride ? d.valueOverrides[prod.name] : origText;
    const imgCount = extractImageTokensFromCell(prod.value).length;

    const ta = document.getElementById(`benefitReviewDesc_${idx}`);
    if (ta) ta.value = currentText;
    const metaEl = document.getElementById(`benefitReviewDescMeta_${idx}`);
    if (metaEl) metaEl.classList.toggle("hidden", !(hasOverride && currentText !== origText));
    const imgBadge = document.getElementById(`benefitReviewDescImgBadge_${idx}`);
    if (imgBadge) {
        imgBadge.classList.toggle("hidden", imgCount === 0);
        imgBadge.innerHTML = `<i class="fa-solid fa-image me-0.5"></i>+${imgCount} img preserved`;
    }
}

function onBenefitReviewDescReset(idx) {
    const item = benefitReviewItems[idx];
    if (!item) return;
    const d = benefitReviewDecisions[item.key];
    if (!d) return;
    const pIdx = Math.min(Math.max(0, d.activeProdIdx || 0), Math.max(0, item.usedBy.length - 1));
    const prod = item.usedBy[pIdx];
    if (!prod) return;
    const origText = stripImageTokensForSnippet(prod.value);
    if (d.valueOverrides) delete d.valueOverrides[prod.name];
    const ta = document.getElementById(`benefitReviewDesc_${idx}`);
    if (ta) ta.value = origText;
    const metaEl = document.getElementById(`benefitReviewDescMeta_${idx}`);
    if (metaEl) metaEl.classList.add("hidden");
    const selEl = document.getElementById(`benefitReviewDescProd_${idx}`);
    if (selEl && selEl.options[pIdx]) {
        selEl.options[pIdx].textContent = prod.name;
    }
}

function resetBenefitReviewState() {
    invalidateReviewLookupCaches();
    benefitReviewShowPendingOnly = false;
    benefitReviewItems = [];
    benefitReviewDecisions = {};
    uploadParsedSchema = [];
    uploadParsedSchemaStructured = false;
    const area = document.getElementById("uploadBenefitReviewArea");
    if (area) { area.replaceChildren(); area.classList.add("hidden"); }
}

// Terapkan keputusan Review Benefit & Non-Benefit ke baris upload: gabung ke kolom yang ada, ganti nama, atau buang kolom.
// Mengembalikan baris baru + daftar alias/kolom baru yang perlu disimpan ke Config.
function applyBenefitReviewDecisions(rows) {
    const aliasMappings = [];
    if (!uploadParsedSchemaStructured || benefitReviewItems.length === 0) {
        return { rows, aliasMappings };
    }

    const schema = uploadParsedSchema.slice();
    const newRows = rows.map(r => {
        const copy = r.slice();
        copy.__sheetName = r.__sheetName;
        return copy;
    });
    const dropIdx = new Set();
    const baseNbSet = new Set(getBaseNonBenefitColumns().map(c => c.toLowerCase()));
    const seenMappingKeys = new Set();
    const pushMapping = (entryType, canonical, alias) => {
        const sig = `${entryType}||${canonical.toLowerCase()}||${String(alias || "").toLowerCase()}`;
        if (seenMappingKeys.has(sig)) return;
        seenMappingKeys.add(sig);
        aliasMappings.push({ entryType, canonical, alias });
    };

    // Proses item "new" terlebih dahulu agar item "map" yang mengacu ke nama baru tsb menemukan kolomnya
    const evaluated = benefitReviewItems.map(item => ({ item, ev: evaluateBenefitReviewItem(item) }));
    evaluated.sort((a, b) => {
        const rank = act => act === "new" ? 0 : (act === "map" ? 1 : 2);
        return rank(a.ev.action) - rank(b.ev.action);
    });

    evaluated.forEach(({ item, ev }) => {
        const srcIdx = schema.findIndex((c, i) => c === item.col && !dropIdx.has(i));
        if (srcIdx === -1 || !ev.valid) return;

        if (ev.action === "ignore") {
            dropIdx.add(srcIdx);
            if (!baseNbSet.has(item.col.toLowerCase())) {
                NON_BENEFIT_COLUMNS = NON_BENEFIT_COLUMNS.filter(c => c.toLowerCase() !== item.col.toLowerCase());
            }
            return;
        }

        // Terapkan edit deskripsi/nilai benefit yang diubah user di kartu Review (pertahankan token gambar [IMG...] jika ada)
        const dec = benefitReviewDecisions[item.key];
        if (dec && dec.valueOverrides && Object.keys(dec.valueOverrides).length > 0) {
            newRows.forEach(r => {
                const prodName = String(r[0] || "-");
                if (!Object.prototype.hasOwnProperty.call(dec.valueOverrides, prodName)) return;
                const editedText = String(dec.valueOverrides[prodName] || "").trim();
                const imgTokens = extractImageTokensFromCell(r[srcIdx]);
                if (editedText && imgTokens.length > 0) {
                    r[srcIdx] = `${editedText}\n${imgTokens.join("\n")}`;
                } else if (editedText) {
                    r[srcIdx] = editedText;
                } else if (imgTokens.length > 0) {
                    r[srcIdx] = imgTokens.join("\n");
                } else {
                    r[srcIdx] = "-";
                }
            });
        }

        const kind = ev.kind || "benefit";
        const entryType = kind === "non-benefit" ? "Non-Benefit" : "Benefit";
        const target = ev.target;
        const isRenamedOrAliased = normalizeBenefitNameKey(item.col) !== normalizeBenefitNameKey(target);

        if (kind === "non-benefit") {
            if (!NON_BENEFIT_COLUMNS.some(c => c.toLowerCase() === target.toLowerCase())) {
                NON_BENEFIT_COLUMNS.push(target);
            }
            if (!baseNbSet.has(item.col.toLowerCase()) && item.col.toLowerCase() !== target.toLowerCase()) {
                NON_BENEFIT_COLUMNS = NON_BENEFIT_COLUMNS.filter(c => c.toLowerCase() !== item.col.toLowerCase());
            }
            if (isRenamedOrAliased || !NON_BENEFIT_CONFIG_CANONICALS.has(normalizeBenefitNameKey(target))) {
                pushMapping("Non-Benefit", target, isRenamedOrAliased ? item.col : "");
            }
        } else {
            if (!baseNbSet.has(item.col.toLowerCase())) {
                NON_BENEFIT_COLUMNS = NON_BENEFIT_COLUMNS.filter(c => c.toLowerCase() !== item.col.toLowerCase());
            }
            if (isRenamedOrAliased || !BENEFIT_CONFIG_CANONICALS.has(normalizeBenefitNameKey(target))) {
                pushMapping("Benefit", target, isRenamedOrAliased ? item.col : "");
            }
        }

        if (target === item.col) return;

        const tgtIdx = schema.findIndex((c, i) => i !== srcIdx && !dropIdx.has(i) && c.toLowerCase() === target.toLowerCase());
        if (tgtIdx === -1) {
            // Nama kolom baru yang sudah diperbaiki: cukup ganti nama kolomnya
            schema[srcIdx] = target;
            return;
        }

        // Gabungkan isi ke kolom target yang sudah ada
        newRows.forEach(r => {
            const sv = r[srcIdx];
            if (isEmptyVal(sv)) return;
            const tv = r[tgtIdx];
            r[tgtIdx] = isEmptyVal(tv) ? sv : `${tv}\n${sv}`;
        });
        dropIdx.add(srcIdx);
    });

    const keepIdx = schema.map((_, i) => i).filter(i => !dropIdx.has(i));
    const finalRows = newRows.map(r => {
        const out = keepIdx.map(i => r[i]);
        out.__sheetName = r.__sheetName;
        return out;
    });
    dynamicColumns = keepIdx.map(i => schema[i]).slice(7);

    return { rows: finalRows, aliasMappings };
}

// Helper untuk mengekstrak Nama Manfaat & Deskripsi Manfaat dari sel Excel (dengan atau tanpa tanda ":")
function extractBenefitKeyAndValue(rawCellText) {
    const text = String(rawCellText || "").trim();
    if (!text || isEmptyVal(text)) return null;

    // Normalisasi alias nama manfaat dari database lama ke nama kolom standar Master Sheet
    const normalizeBenefitKey = (rawKey, rawVal) => {
        let k = String(rawKey || "").replace(/^[-•*]\s*/, '').replace(/:$/, '').trim();
        let v = String(rawVal || "").trim();
        const lowerK = k.toLowerCase();

        if (lowerK === "accident death benefit" || lowerK === "accidental death benefit") {
            return { key: "Accidental Death Benefit", value: v };
        }

        // Cocokkan dengan kolom manfaat yang sudah ada / nama standar / alias di Config
        // (abaikan huruf besar-kecil, titik dua di akhir, spasi ganda, dan bentuk jamak: "Maturity Benefits" -> "Maturity Benefit")
        const resolvedKnown = resolveKnownBenefitName(k, standardBenefits.filter(b => b !== "Accident Death Benefit"));
        if (resolvedKnown) {
            return { key: resolvedKnown, value: v };
        }

        if (/^maturity benefits?\b/i.test(k)) {
            const extraTag = k.replace(/^maturity benefits?\b/i, "").trim();
            return {
                key: "Maturity Benefit",
                value: extraTag ? `${extraTag}\n${v}`.trim() : v
            };
        }
        if (/^ci benefits?\b/i.test(k)) {
            const extraTag = k.replace(/^ci benefits?\b/i, "").trim();
            return {
                key: "Critical Illness Benefit",
                value: extraTag ? `${extraTag}\n${v}`.trim() : v
            };
        }
        if (/^additional death benefits?\b/i.test(k) && !/^additional death benefits?$/i.test(k)) {
            const extraTag = k.replace(/^additional death benefits?\b/i, "").trim();
            return {
                key: "Additional Death Benefit",
                value: extraTag ? `${extraTag}:\n${v}`.trim() : v
            };
        }
        return { key: k, value: v };
    };

    // Daftar nama manfaat standar dari Config + kolom manfaat yang sudah ada di database (diurutkan dari yang terpanjang)
    const standardBenefits = (Array.isArray(STANDARD_BENEFIT_COLUMNS) && STANDARD_BENEFIT_COLUMNS.length > 0)
        ? STANDARD_BENEFIT_COLUMNS
        : DEFAULT_STANDARD_BENEFIT_COLUMNS;
    const knownBenefits = [...new Set([...getBenefitColumns(), ...standardBenefits])]
        .sort((a, b) => b.length - a.length);

    // 1. Pisahkan per baris terlebih dahulu (Format paling umum di DB lama & baru: Baris 1 = Nama Manfaat, Baris 2+ = Deskripsi)
    const lines = text.split(/\r\n|\r|\n/).map(l => l.trim()).filter(l => l !== "");
    if (lines.length === 0) return null;

    const firstLine = lines[0].replace(/^[-•*]\s*/, '').trim();

    // 1a. Jika di baris pertama terdapat tanda ":"
    const firstColonIdx = firstLine.indexOf(":");
    if (firstColonIdx > 0 && firstColonIdx < 100) {
        const key = firstLine.substring(0, firstColonIdx).trim();
        const restFirstLine = firstLine.substring(firstColonIdx + 1).trim();
        const descLines = restFirstLine ? [restFirstLine, ...lines.slice(1)] : lines.slice(1);
        return normalizeBenefitKey(key, descLines.join("\n").trim());
    }

    // 1b. Jika multi-baris TANPA tanda ":" di baris pertama (Baris 1 = Judul Manfaat, Baris 2+ = Isi Manfaat)
    // Contoh DB Lama:
    // "Death Benefit\n* Y1-Y2: 105% Total premium paid\n* >Y2: 100% SA"
    // "Dana Tahapan Haji\n1st Stage: 50% SA after PPP..."
    if (lines.length >= 2 && firstLine.length <= 100) {
        const key = firstLine;
        const desc = lines.slice(1).join("\n").trim();
        return normalizeBenefitKey(key, desc);
    }

    // 2. Cek apakah diawali dengan salah satu nama manfaat yang sudah dikenal dalam 1 baris
    const lowerText = text.toLowerCase();
    for (const benefitName of knownBenefits) {
        const lowerBen = benefitName.toLowerCase();
        if (lowerText.startsWith(lowerBen)) {
            const remainder = text.substring(benefitName.length);
            if (!remainder || /^[\s:\r\n\-–]+/.test(remainder)) {
                const cleanDesc = remainder.replace(/^[\s:\r\n\-–]+/, '').trim();
                return normalizeBenefitKey(benefitName, cleanDesc);
            }
        }
    }

    // 3. Jika hanya 1 baris tanpa ":", cek pemisah " - " atau kata kunci "Benefit / Bonus / Waiver"
    if (lines.length === 1) {
        const dashMatch = firstLine.match(/^([A-Za-z][A-Za-z0-9\s/()&]{2,75}?)\s+[-–]\s+(.+)$/);
        if (dashMatch) {
            return normalizeBenefitKey(dashMatch[1].trim(), dashMatch[2].trim());
        }

        const kwMatch = firstLine.match(/^([A-Za-z][A-Za-z0-9\s/()&]{2,75}?(?:Benefit|Benefits|Bonus|Payout|Waiver|Waived))\s+(.+)$/i);
        if (kwMatch) {
            return normalizeBenefitKey(kwMatch[1].trim(), kwMatch[2].trim());
        }
    }

    return null;
}

// Helper untuk mengecek apakah label baris adalah bagian dari Fitur Tambahan (Additional Feature / Benefit / Other Benefit)
function isAdditionalBenefitLabel(lowerLabel) {
    if (!lowerLabel) return false;
    const cfgAliases = getConfigAliasesForColumn("Additional Benefit", [
        "additional feature", "additional benefit", "other benefit", "other feature", "fitur tambahan"
    ]).map(a => a.toLowerCase());
    return cfgAliases.some(a => lowerLabel === a || lowerLabel.startsWith(a));
}

// Helper untuk mengecek label baris khusus Syariah (Aqd / Akad, Alokasi Kontribusi, Surplus Underwriting)
function isAqdLabel(lowerLabel) {
    if (!lowerLabel) return false;
    const cfgAliases = getConfigAliasesForColumn("Aqd", ["aqd", "akad", "aqd / akad", "akad / aqd", "aqd / contract", "aqd/contract"]).map(a => a.toLowerCase());
    return cfgAliases.includes(lowerLabel);
}

function isAlokasiKontribusiLabel(lowerLabel) {
    if (!lowerLabel) return false;
    if (lowerLabel.startsWith("alokasi kontribusi")) return true;
    const cfgAliases = getConfigAliasesForColumn("Alokasi Kontribusi", []).map(a => a.toLowerCase());
    return cfgAliases.includes(lowerLabel);
}

function isSurplusUnderwritingLabel(lowerLabel) {
    if (!lowerLabel) return false;
    if (lowerLabel.startsWith("surplus underwriting") || lowerLabel.startsWith("underwriting surplus")) return true;
    const cfgAliases = getConfigAliasesForColumn("Surplus Underwriting", []).map(a => a.toLowerCase());
    return cfgAliases.includes(lowerLabel);
}

function isBenefitSectionLabel(lowerLabel) {
    if (!lowerLabel) return false;
    return (
        lowerLabel === "benefits" ||
        lowerLabel === "benefit" ||
        lowerLabel === "benefit(s)" ||
        lowerLabel === "product benefit" ||
        lowerLabel === "product benefits" ||
        lowerLabel === "product benefit(s)" ||
        lowerLabel === "manfaat produk" ||
        lowerLabel === "manfaat" ||
        lowerLabel.startsWith("benefit")
    );
}

function isUmbrellaBenefitLabel(lowerLabel) {
    if (!lowerLabel) return false;
    return (
        lowerLabel === "living benefit" ||
        lowerLabel === "living benefits" ||
        lowerLabel === "manfaat hidup" ||
        lowerLabel === "other benefit" ||
        lowerLabel === "other benefits" ||
        lowerLabel === "manfaat lainnya"
    );
}

function normalizeBenefitCoreWords(name) {
    const stop = new Set(["benefit", "manfaat", "and", "of", "the", "dan"]);
    return normalizeBenefitNameKey(name)
        .replace(/&/g, " ")
        .split(/[^a-z0-9]+/)
        .filter(w => w && !stop.has(w))
        .join(" ");
}

function getBenefitAcronyms(name) {
    const cleaned = normalizeBenefitNameKey(name).replace(/&/g, " ");
    const allWords = cleaned.split(/[^a-z0-9]+/).filter(Boolean);
    const stop = new Set(["benefit", "manfaat", "and", "of", "the", "dan"]);
    const coreWords = allWords.filter(w => !stop.has(w));
    const out = new Set();
    if (coreWords.length >= 2) {
        out.add(coreWords.map(w => w[0]).join(""));
    }
    const withBenefitWords = allWords.filter(w => w !== "and" && w !== "the" && w !== "dan");
    if (withBenefitWords.length >= 2) {
        out.add(withBenefitWords.map(w => w[0]).join(""));
    }
    return out;
}

function isEquivalentBenefitTitle(titleA, titleB) {
    const keyA = normalizeBenefitNameKey(titleA);
    const keyB = normalizeBenefitNameKey(titleB);
    if (!keyA || !keyB) return false;
    if (keyA === keyB) return true;

    const coreA = normalizeBenefitCoreWords(titleA);
    const coreB = normalizeBenefitCoreWords(titleB);
    if (coreA && coreB && coreA === coreB) return true;

    // Cek singkatan/akronim (misal "TPD Benefit" <-> "Total and Permanent Disability", "CI Benefit" <-> "Critical Illness")
    if (/^[a-z]{2,5}$/.test(coreA) && getBenefitAcronyms(titleB).has(coreA)) return true;
    if (/^[a-z]{2,5}$/.test(coreB) && getBenefitAcronyms(titleA).has(coreB)) return true;

    // Cek melalui alias Config / nama Master Sheet yang dikenal (termasuk varian dengan/tanpa kata "Benefit")
    const resolvedA = resolveKnownBenefitName(titleA) || findExistingMasterBenefit(titleA);
    const resolvedB = resolveKnownBenefitName(titleB) || findExistingMasterBenefit(titleB);
    const canonA = resolvedA ? normalizeBenefitCoreWords(resolvedA) : coreA;
    const canonB = resolvedB ? normalizeBenefitCoreWords(resolvedB) : coreB;
    if (canonA && canonB && canonA === canonB) return true;
    if (resolvedA && /^[a-z]{2,5}$/.test(coreB) && getBenefitAcronyms(resolvedA).has(coreB)) return true;
    if (resolvedB && /^[a-z]{2,5}$/.test(coreA) && getBenefitAcronyms(resolvedB).has(coreA)) return true;

    return false;
}

function stripDuplicateLeadingBenefitTitle(colName, rawVal, altLabel = "") {
    const text = String(rawVal || "").trim();
    if (!text) return "";
    const lines = text.split(/\r?\n/);
    if (lines.length < 2) return text;
    const firstLineClean = lines[0].replace(/[:：]\s*$/, "").trim();
    if (!firstLineClean) return text;
    if (
        isEquivalentBenefitTitle(firstLineClean, colName) ||
        (altLabel && isEquivalentBenefitTitle(firstLineClean, altLabel))
    ) {
        const rest = lines.slice(1).join("\n").trim();
        return rest || text;
    }
    return text;
}

function isChargesSectionLabel(lowerLabel) {
    if (!lowerLabel) return false;
    return (
        lowerLabel === "charges" ||
        lowerLabel === "charge" ||
        lowerLabel === "charge(s)" ||
        lowerLabel === "biaya" ||
        lowerLabel === "biaya-biaya" ||
        lowerLabel.startsWith("charges")
    );
}

function extractDplkBenefitColumn(rawCellText) {
    if (!rawCellText || typeof rawCellText !== "string") return null;
    const cleaned = rawCellText.replace(/\r\n|\r|\n/g, ' ').replace(/\s+/g, ' ').trim();
    if (!cleaned || cleaned === "-" || isEmptyVal(cleaned)) return null;
    const lower = cleaned.toLowerCase();

    const rules = (Array.isArray(DPLK_BENEFIT_RULES) && DPLK_BENEFIT_RULES.length > 0)
        ? DPLK_BENEFIT_RULES
        : DEFAULT_DPLK_BENEFIT_RULES;

    for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        const canonical = String(rule.canonical || "").trim();
        if (!canonical) continue;
        if (lower === canonical.toLowerCase()) return canonical;
        const aliases = Array.isArray(rule.aliases) ? rule.aliases : [];
        for (let j = 0; j < aliases.length; j++) {
            const kw = String(aliases[j] || "").trim().toLowerCase();
            if (!kw) continue;
            if (kw === "normal" || kw === "ditunda") {
                if (lower === kw || lower.includes(`pensiun ${kw}`)) return canonical;
            } else if (kw === "pengalihan dplk") {
                if (lower.includes("pengalihan") && lower.includes("dplk")) return canonical;
            } else if (lower.includes(kw)) {
                return canonical;
            }
        }
    }

    return null;
}

function extractDplkChargeKeyAndValue(rawCellText) {
    if (!rawCellText || typeof rawCellText !== "string") return null;
    const trimmed = rawCellText.trim();
    if (!trimmed || trimmed === "-" || isEmptyVal(trimmed)) return null;

    const colonIdx = trimmed.indexOf(":");
    let rawKey = colonIdx !== -1 ? trimmed.substring(0, colonIdx).trim() : trimmed;
    let rawVal = colonIdx !== -1 ? trimmed.substring(colonIdx + 1).trim() : trimmed;

    const lowerKey = rawKey.toLowerCase().replace(/\s+/g, ' ').trim();
    let canonicalKey = null;

    const rules = (Array.isArray(DPLK_CHARGE_RULES) && DPLK_CHARGE_RULES.length > 0)
        ? DPLK_CHARGE_RULES
        : DEFAULT_DPLK_CHARGE_RULES;

    for (let i = 0; i < rules.length; i++) {
        const rule = rules[i];
        const canonical = String(rule.canonical || "").trim();
        if (!canonical) continue;
        if (lowerKey === canonical.toLowerCase()) {
            canonicalKey = canonical;
            break;
        }
        const aliases = Array.isArray(rule.aliases) ? rule.aliases : [];
        const matched = aliases.some(a => {
            const kw = String(a || "").trim().toLowerCase();
            if (!kw) return false;
            if (kw === "pengalihan investasi" || kw === "pengalihan paket") {
                return lowerKey.includes(kw) || (lowerKey.includes("pengalihan") && (lowerKey.includes("investasi") || lowerKey.includes("paket")));
            }
            return lowerKey.includes(kw);
        });
        if (matched) {
            canonicalKey = canonical;
            break;
        }
    }

    if (!canonicalKey) {
        if (lowerKey.startsWith("biaya ")) {
            canonicalKey = rawKey.replace(/\s+/g, ' ').trim();
        } else if (colonIdx !== -1 && rawKey.length <= 50) {
            canonicalKey = /^biaya\b/i.test(rawKey) ? rawKey.replace(/\s+/g, ' ').trim() : `Biaya ${rawKey.replace(/\s+/g, ' ').trim()}`;
        }
    }

    if (!canonicalKey) return null;
    return {
        key: canonicalKey,
        value: rawVal || trimmed
    };
}

// SMART DYNAMIC PARSER (MENDUKUNG DATABASE LAMA, EXCEL VERTIKAL BARU, SYARIAH + GAMBAR TABEL, DPLK, & EXCEL HORIZONTAL)
function parseExcelToSheetsSchema(workbook, excelImagesBySheet = {}) {
    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) return [];

    // Pastikan seluruh kolom non-benefit (termasuk Aqd, Alokasi Kontribusi, Surplus Underwriting, & DPLK) terdaftar di dynamicColumns
    ensureNonBenefitColumnsInDynamicSchema(true, true);

    let allParsedProducts = [];
    let flatRawSheets = [];

    workbook.SheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) return;
        const sheetImgMap = excelImagesBySheet[sheetName] || {};

        // Perbaiki dimensi "!ref" pada file Excel yang memiliki baris kosong di baris 1.048.576
        let maxRow = 0;
        let maxCol = 0;
        let hasAnyCell = false;
        Object.keys(worksheet).forEach(key => {
            if (key[0] === '!') return;
            const cell = worksheet[key];
            if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
                const decoded = XLSX.utils.decode_cell(key);
                if (decoded.r <= 2000 && decoded.c <= 2000) {
                    if (decoded.r > maxRow) maxRow = decoded.r;
                    if (decoded.c > maxCol) maxCol = decoded.c;
                    hasAnyCell = true;
                }
            }
        });

        if (!hasAnyCell) return;
        worksheet['!ref'] = XLSX.utils.encode_range({
            s: { r: 0, c: 0 },
            e: { r: maxRow, c: maxCol }
        });

        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
        if (!rawData || rawData.length === 0) return;

        // Case A: Excel Flat / Horizontal (Baris pertama memiliki kolom "Product Name" DAN "Company Name" sekaligus)
        const headerRow = (rawData[0] || []).map(c => String(c).trim().toLowerCase());
        if (headerRow.includes("product name") && headerRow.includes("company name")) {
            flatRawSheets.push({ headerRow, rawData, sheetName });
            return;
        }

        // Case B: Excel Vertikal Custom (Database Lama, example2.xlsx, Database Syariah, DPLK, & example3.xlsx)
        // Deteksi kolom label paling kiri yang memuat label baris (Kolom A, B, dst.) menggunakan daftar dari Config + default
        const KNOWN_LABELS = [...new Set([
            "product name", "nama produk", "company name", "nama perusahaan", "company",
            "distribution channel", "channel distribution", "product type", "tipe produk",
            "currency", "mata uang", "par/non par", "product benefit", "benefit(s)", "benefits",
            "additional feature (if any)", "additional feature", "additional benefit", "other benefit", "fitur tambahan",
            "charges", "charge(s)", "biaya", "premium (iuran)", "entry age", "sales tools", "after sales tools",
            "administration guideline", "underwriting method", "insured entry age", "coverage period",
            ...CORE_FIELD_RULES.flatMap(r => [r.canonical, ...(Array.isArray(r.aliases) ? r.aliases : [])]),
            ...NON_BENEFIT_COLUMNS,
            ...Object.keys(NON_BENEFIT_ALIAS_MAP || {})
        ].map(s => String(s || "").trim().toLowerCase()).filter(Boolean))];

        let labelColIdx = 0;
        let maxLabelMatches = 0;
        const maxColToCheck = Math.min(6, rawData.reduce((m, r) => Math.max(m, r ? r.length : 0), 0));

        for (let c = 0; c < maxColToCheck; c++) {
            let matches = 0;
            for (let r = 0; r < Math.min(rawData.length, 40); r++) {
                const cellVal = String((rawData[r] && rawData[r][c]) || "").replace(/:$/, '').trim().toLowerCase();
                if (cellVal && KNOWN_LABELS.some(lbl => cellVal === lbl || cellVal.startsWith(lbl))) {
                    matches++;
                }
            }
            if (matches > maxLabelMatches) {
                maxLabelMatches = matches;
                labelColIdx = c;
            }
        }

        // Cari baris di mana labelColIdx adalah "Product Name", "Company Name", dan "Distribution Channel" (mendukung alias dari Config)
        let prodNameRowIdx = -1;
        let compNameRowIdx = -1;
        let distChannelRowIdx = -1;
        const prodNameAliases = getConfigAliasesForColumn("Product Name", ["nama produk"]).map(a => a.toLowerCase());
        const compNameAliases = getConfigAliasesForColumn("Company Name", ["nama perusahaan", "company"]).map(a => a.toLowerCase());
        const distChanAliases = getConfigAliasesForColumn("Distribution Channel", ["channel distribution"]).map(a => a.toLowerCase());

        for (let r = 0; r < rawData.length; r++) {
            const rowLabel = String((rawData[r] && rawData[r][labelColIdx]) || "").replace(/:$/, '').trim().toLowerCase();
            if (prodNameAliases.includes(rowLabel) && prodNameRowIdx === -1) prodNameRowIdx = r;
            if (compNameAliases.includes(rowLabel) && compNameRowIdx === -1) compNameRowIdx = r;
            if (distChanAliases.includes(rowLabel) && distChannelRowIdx === -1) distChannelRowIdx = r;
        }

        // Abaikan sheet dokumentasi/catatan yang tidak memiliki Product Name maupun Company Name (seperti sheet "Doc" di DPLK)
        if (prodNameRowIdx === -1 && compNameRowIdx === -1 && maxLabelMatches < 2) {
            return;
        }

        const isDplkFormatSheet = (prodNameRowIdx === -1 && compNameRowIdx !== -1);

        // Jika tidak ada label "Product Name" dan bukan format DPLK (Company Name only), gunakan baris pertama yang berisi data produk
        if (prodNameRowIdx === -1 && !isDplkFormatSheet) {
            for (let r = 0; r < rawData.length; r++) {
                if (rawData[r] && rawData[r].slice(labelColIdx + 1).some(cell => String(cell || "").trim() !== "")) {
                    prodNameRowIdx = r;
                    break;
                }
            }
        }
        if (prodNameRowIdx === -1 && !isDplkFormatSheet) return;

        let colCount = rawData.reduce((max, r) => Math.max(max, r ? r.length : 0), 0);
        let lastBenefitTargetByRow = {};

        // Hitung frekuensi Company Name pada sheet format DPLK untuk membedakan kolom jika ada nama perusahaan yang sama (misal Sheet1: Direct vs GB)
        const companyCountInSheet = {};
        if (isDplkFormatSheet && rawData[compNameRowIdx]) {
            for (let c = labelColIdx + 1; c < colCount; c++) {
                const rawComp = String(rawData[compNameRowIdx][c] || "").replace(/\r\n|\r|\n/g, ' ').trim();
                if (!rawComp || rawComp === "-" || /^source(\s+of\s+information)?$/i.test(rawComp)) continue;
                const normComp = rawComp.toLowerCase();
                companyCountInSheet[normComp] = (companyCountInSheet[normComp] || 0) + 1;
            }
        }

        for (let c = labelColIdx + 1; c < colCount; c++) {
            // Abaikan kolom "Source of Information" (seperti kolom C, E, G, I, K, M, O pada Comparison_DPLK.xlsm)
            let isSourceColumn = false;
            const headerCheckLimit = Math.max(prodNameRowIdx, compNameRowIdx, 3);
            for (let rCheck = 0; rCheck <= Math.min(rawData.length - 1, headerCheckLimit); rCheck++) {
                const checkCell = String((rawData[rCheck] && rawData[rCheck][c]) || "").replace(/\r\n|\r|\n/g, ' ').trim();
                if (/^source(\s+of\s+information)?$/i.test(checkCell)) {
                    isSourceColumn = true;
                    break;
                }
            }
            if (isSourceColumn) continue;

            let prodName = "";
            let company = "-";

            if (isDplkFormatSheet) {
                company = sanitizeValue(String((rawData[compNameRowIdx] && rawData[compNameRowIdx][c]) || "").replace(/\r\n|\r|\n/g, ' ').trim());
                if (!company || company === "-" || company === "undefined" || company === "null") continue;

                const distChannelVal = (distChannelRowIdx !== -1 && rawData[distChannelRowIdx])
                    ? sanitizeValue(String(rawData[distChannelRowIdx][c] || "").replace(/\r\n|\r|\n/g, ' ').trim())
                    : "-";
                const hasDuplicateCompany = (companyCountInSheet[company.toLowerCase()] || 0) > 1;
                const topLeftTitle = String((rawData[0] && rawData[0][labelColIdx]) || "").replace(/\r\n|\r|\n/g, ' ').trim();

                if (/^sheet\d*$/i.test(sheetName) && topLeftTitle && topLeftTitle.toLowerCase() !== "company name") {
                    prodName = (hasDuplicateCompany && distChannelVal !== "-")
                        ? `${company} - ${topLeftTitle} (${distChannelVal})`
                        : `${company} - ${topLeftTitle}`;
                } else {
                    prodName = (hasDuplicateCompany && distChannelVal !== "-")
                        ? `${company} - ${sheetName} (${distChannelVal})`
                        : `${company} - ${sheetName}`;
                }
            } else {
                prodName = String((rawData[prodNameRowIdx] && rawData[prodNameRowIdx][c]) || "")
                    .replace(/\r\n|\r|\n/g, ' ')
                    .trim();
                if (!prodName || prodName === "-" || prodName === "undefined" || prodName === "null") continue;
                if (/^source(\s+of\s+information)?$/i.test(prodName)) continue;

                // Ambil Company Name dari baris "Company Name" jika tersedia di Excel
                if (compNameRowIdx !== -1 && rawData[compNameRowIdx]) {
                    company = sanitizeValue(String(rawData[compNameRowIdx][c] || "").trim());
                }
            }

            // Fallback deteksi nama perusahaan dari nama produk menggunakan daftar "Company Keyword" di sheet Config
            if (company === "-") {
                const lowerName = prodName.toLowerCase();
                const compRules = (Array.isArray(COMPANY_KEYWORD_RULES) && COMPANY_KEYWORD_RULES.length > 0)
                    ? COMPANY_KEYWORD_RULES
                    : DEFAULT_COMPANY_KEYWORD_RULES;
                for (let i = 0; i < compRules.length; i++) {
                    const rule = compRules[i];
                    const canonical = String(rule.canonical || "").trim();
                    if (!canonical) continue;
                    const kws = [canonical.toLowerCase(), ...(Array.isArray(rule.aliases) ? rule.aliases.map(a => String(a || "").trim().toLowerCase()) : [])].filter(Boolean);
                    if (kws.some(kw => lowerName.includes(kw))) {
                        company = canonical;
                        break;
                    }
                }
            }

            let attrMap = {};
            let attrImagesMap = {};
            let lastSectionLabel = "";
            let currentSection = "identity";

            const appendTextToAttr = (key, textVal) => {
                const sanitized = sanitizeValue(textVal);
                if (sanitized === "-") return;
                if (attrMap[key] && attrMap[key] !== "-") {
                    if (attrMap[key] === "✓" && sanitized === "✓") return;
                    attrMap[key] += "\n\n" + sanitized;
                } else {
                    attrMap[key] = sanitized;
                }
            };

            const appendImagesToAttr = (key, imgEntries) => {
                if (!imgEntries || imgEntries.length === 0) return;
                if (!attrImagesMap[key]) attrImagesMap[key] = [];
                imgEntries.forEach(entry => {
                    if (entry && entry.imgElement) {
                        attrImagesMap[key].push(entry.imgElement);
                    }
                });
            };

            for (let r = 0; r < rawData.length; r++) {
                if (r === prodNameRowIdx || r === compNameRowIdx) continue;

                let label = String((rawData[r] && rawData[r][labelColIdx]) || "").trim();
                let cleanLabel = label.replace(/:$/, '').trim();
                let lowerLabel = cleanLabel.toLowerCase();

                // Hentikan pembacaan jika sudah memasuki tabel ilustrasi dummy atau catatan bawah di Excel
                if (
                    lowerLabel.startsWith("gender, smoking") ||
                    (lowerLabel.startsWith("gender") && lowerLabel.includes("smoking")) ||
                    lowerLabel.startsWith("tambah ppip") ||
                    lowerLabel.startsWith("sheet dipisah") ||
                    lowerLabel === "source"
                ) {
                    break;
                }

                // Perbarui status section berdasarkan label baris
                const KNOWN_SPEC_EXIT_LABELS = new Set([
                    "product name", "nama produk", "company name", "nama perusahaan", "company",
                    "category system", "par/non par", "product type", "tipe produk",
                    "distribution channel", "channel distribution", "currency", "mata uang",
                    "launch date", "tanggal rilis", "launching date", "release date", "estimation launching date", "estimated launching date", "est. launching date", "est launching date",
                    "entry age", "insured entry age", "age definition",
                    "coverage period", "coverage term",
                    "premium payment period (ppp)", "premium payment period", "premium term", "ppp",
                    "min premium", "premium (iuran)", "iuran", "premium limit", "minimum premium", "premium",
                    "min sum assured", "sum assured limit", "minimum sum assured", "sum assured",
                    "underwriting method", "underwriting",
                    "available riders", "available rider", "riders", "rider", "nfo",
                    "sales tools", "sales tool", "after sales tools", "after sales tool", "remarks",
                    ...CORE_FIELD_RULES.flatMap(r => [r.canonical, ...(Array.isArray(r.aliases) ? r.aliases : [])]).map(s => String(s || "").trim().toLowerCase()).filter(Boolean)
                ]);

                if (cleanLabel) {
                    lastSectionLabel = lowerLabel;
                    if (isBenefitSectionLabel(lowerLabel)) {
                        currentSection = "benefits";
                    } else if (isChargesSectionLabel(lowerLabel)) {
                        currentSection = "charges";
                    } else if (isAdditionalBenefitLabel(lowerLabel)) {
                        // Batas akhir zona benefit: Additional Feature (if any) / Other Benefit langsung masuk ke Fitur Tambahan
                        currentSection = "additional_feature";
                    } else if (lowerLabel.startsWith("administration guideline") || lowerLabel.startsWith("admin guideline") || lowerLabel.includes("guideline")) {
                        currentSection = "admin";
                    } else if (isAlokasiKontribusiLabel(lowerLabel)) {
                        currentSection = "alokasi_kontribusi";
                    } else if (isAqdLabel(lowerLabel)) {
                        currentSection = "aqd";
                    } else if (isSurplusUnderwritingLabel(lowerLabel)) {
                        currentSection = "surplus_underwriting";
                    } else if (KNOWN_SPEC_EXIT_LABELS.has(lowerLabel) || isNonBenefitColumn(cleanLabel) || Boolean(resolveKnownNonBenefitName(cleanLabel))) {
                        currentSection = "spec";
                    } else if (currentSection === "benefits") {
                        // Tetap berada di zona benefits ketika baris memiliki nama benefit di kolom kiri (misal Death Benefit, Maturity Benefit, Living Benefit)
                        currentSection = "benefits";
                    } else if (Boolean(resolveKnownBenefitName(cleanLabel)) || Boolean(findExistingMasterBenefit(cleanLabel))) {
                        // Auto-deteksi: jika label kolom kiri sudah ada di daftar Benefit Master Sheet / Config, langsung perlakukan sebagai Benefit
                        currentSection = "benefits";
                    } else {
                        currentSection = "spec";
                    }
                }

                let val = String((rawData[r] && rawData[r][c]) || "").trim();
                const hasText = Boolean(val && val !== "-" && val !== "null" && !isEmptyVal(val));
                const rowImgs = (sheetImgMap[c] && sheetImgMap[c][r]) ? sheetImgMap[c][r] : [];
                const hasImgs = rowImgs.length > 0;

                if (!hasText && !hasImgs) continue;

                // 1. Bagian Alokasi Kontribusi (Syariah: menggabungkan seluruh sub-baris Tabarru', Tanahud, Ujrah, Nisbah, & Gambar Tabel)
                if (currentSection === "alokasi_kontribusi" || isAlokasiKontribusiLabel(lowerLabel)) {
                    if (hasText) appendTextToAttr("Alokasi Kontribusi", val);
                    if (hasImgs) appendImagesToAttr("Alokasi Kontribusi", rowImgs);
                }
                // 2. Bagian Aqd / Akad (Syariah)
                else if (currentSection === "aqd" || isAqdLabel(lowerLabel)) {
                    if (hasText) appendTextToAttr("Aqd", val);
                    if (hasImgs) appendImagesToAttr("Aqd", rowImgs);
                }
                // 3. Bagian Surplus Underwriting (Syariah)
                else if (currentSection === "surplus_underwriting" || isSurplusUnderwritingLabel(lowerLabel)) {
                    if (hasText) appendTextToAttr("Surplus Underwriting", val);
                    if (hasImgs) appendImagesToAttr("Surplus Underwriting", rowImgs);
                }
                // 4. Bagian Additional Feature (if any) / Additional Benefit / Other Benefit / Fitur Tambahan
                else if (currentSection === "additional_feature" || isAdditionalBenefitLabel(lowerLabel)) {
                    if (hasText) appendTextToAttr("Additional Benefit", val);
                    if (hasImgs) appendImagesToAttr("Additional Benefit", rowImgs);
                }
                // 5. Bagian Charges / Biaya-Biaya (DPLK)
                else if (currentSection === "charges" || isChargesSectionLabel(lowerLabel)) {
                    let targetChargeCol = null;
                    let chargeVal = "";

                    if (hasText) {
                        const extractedCharge = extractDplkChargeKeyAndValue(val);
                        if (extractedCharge && extractedCharge.key) {
                            targetChargeCol = extractedCharge.key;
                            chargeVal = extractedCharge.value || val;
                        } else {
                            targetChargeCol = "Biaya Lain (Pihak Ketiga)";
                            chargeVal = val;
                        }
                    }

                    if (targetChargeCol) {
                        const existingCol = dynamicColumns.find(col => col.toLowerCase() === targetChargeCol.toLowerCase());
                        const finalChargeCol = existingCol || targetChargeCol;
                        if (!dynamicColumns.includes(finalChargeCol)) {
                            dynamicColumns.push(finalChargeCol);
                        }
                        if (chargeVal) {
                            appendTextToAttr(finalChargeCol, chargeVal);
                        }
                        if (hasImgs) {
                            appendImagesToAttr(finalChargeCol, rowImgs);
                        }
                    }
                }
                // 6. Bagian Benefit(s) / Product Benefit (Zona Manfaat produk hingga sebelum Additional Feature / Other Benefit)
                else if (currentSection === "benefits" || isBenefitSectionLabel(lowerLabel)) {
                    // Cek apakah ini manfaat DPLK (mode Checklist "✓")
                    const dplkBenCol = hasText ? extractDplkBenefitColumn(val) : null;
                    if (dplkBenCol && (isDplkFormatSheet || !val.includes(":"))) {
                        const existingCol = dynamicColumns.find(col => col.toLowerCase() === dplkBenCol.toLowerCase())
                            || resolveKnownBenefitName(dplkBenCol);
                        const finalColName = existingCol || dplkBenCol;
                        if (!dynamicColumns.includes(finalColName) && !isNonBenefitColumn(finalColName)) {
                            const insertIdx = Math.min(5, dynamicColumns.length);
                            dynamicColumns.splice(insertIdx, 0, finalColName);
                        }
                        attrMap[finalColName] = "✓";
                        if (hasImgs) appendImagesToAttr(finalColName, rowImgs);
                        continue;
                    }

                    let targetBenefitCol = null;
                    let benefitVal = "";
                    const hasRowBenefitLabel = Boolean(cleanLabel && !isBenefitSectionLabel(lowerLabel));

                    if (hasText) {
                        if (hasRowBenefitLabel) {
                            // Nama benefit ada di kolom kiri (misal Format Export Comparison & Format Syariah)
                            const resolvedRowBen = resolveKnownBenefitName(cleanLabel) || findExistingMasterBenefit(cleanLabel) || cleanLabel;
                            const isRegisteredDirectBenefit =
                                BENEFIT_CONFIG_CANONICALS.has(normalizeBenefitNameKey(cleanLabel)) ||
                                Boolean(findExistingMasterBenefit(cleanLabel));
                            if (isUmbrellaBenefitLabel(lowerLabel) && !isRegisteredDirectBenefit) {
                                // Kategori payung (misal "Living Benefit" / "Other Benefit"): gunakan nama spesifik dari dalam sel (Cash Benefit, Manfaat Dana Tahapan, dll.)
                                const extracted = extractBenefitKeyAndValue(val);
                                if (extracted && extracted.key) {
                                    targetBenefitCol = extracted.key;
                                    benefitVal = (extracted.value && extracted.value.trim() !== "") ? extracted.value : val;
                                } else {
                                    targetBenefitCol = resolvedRowBen;
                                    benefitVal = stripDuplicateLeadingBenefitTitle(resolvedRowBen, val, cleanLabel);
                                }
                            } else {
                                // Kolom kiri sudah memuat nama benefit spesifik (misal "Additional Cash Payout", "Death Benefit", "Maturity Benefit")
                                // Jangan biarkan deskripsi 1 baris dipecah menjadi nama kolom baru oleh extractBenefitKeyAndValue
                                targetBenefitCol = resolvedRowBen;
                                benefitVal = stripDuplicateLeadingBenefitTitle(resolvedRowBen, val, cleanLabel);
                            }
                        } else {
                            // Format Excel vertikal lama (kolom kiri kosong, nama & nilai ada di dalam sel)
                            const extracted = extractBenefitKeyAndValue(val);
                            if (extracted && extracted.key) {
                                targetBenefitCol = extracted.key;
                                benefitVal = (extracted.value && extracted.value.trim() !== "") ? extracted.value : extracted.key;
                            } else if (lastBenefitTargetByRow[r]) {
                                targetBenefitCol = lastBenefitTargetByRow[r];
                                benefitVal = val;
                            }
                        }
                    } else if (hasRowBenefitLabel && hasImgs) {
                        targetBenefitCol = resolveKnownBenefitName(cleanLabel) || findExistingMasterBenefit(cleanLabel) || cleanLabel;
                    }

                    if (targetBenefitCol) {
                        const existingCol = dynamicColumns.find(col => col.toLowerCase() === targetBenefitCol.toLowerCase())
                            || resolveKnownBenefitName(targetBenefitCol);
                        const finalColName = existingCol || targetBenefitCol;
                        targetBenefitCol = finalColName;
                        lastBenefitTargetByRow[r] = finalColName;

                        const dedupedVal = stripDuplicateLeadingBenefitTitle(finalColName, benefitVal, cleanLabel);
                        const sanitizedVal = sanitizeValue(dedupedVal);
                        if (!dynamicColumns.includes(finalColName) && (sanitizedVal !== "-" || hasImgs) && !isNonBenefitColumn(finalColName)) {
                            // Sisipkan kolom manfaat baru tepat setelah Column K (index ke-5 pada dynamicColumns)
                            const insertIdx = Math.min(5, dynamicColumns.length);
                            dynamicColumns.splice(insertIdx, 0, finalColName);
                        }

                        if (sanitizedVal !== "-") {
                            appendTextToAttr(finalColName, sanitizedVal);
                        }
                    }

                    if (!targetBenefitCol && hasImgs) {
                        targetBenefitCol = lastBenefitTargetByRow[r] || "Manfaat Dana Tahapan";
                        const existingCol = dynamicColumns.find(col => col.toLowerCase() === targetBenefitCol.toLowerCase());
                        targetBenefitCol = existingCol || targetBenefitCol;
                        if (!dynamicColumns.includes(targetBenefitCol) && !isNonBenefitColumn(targetBenefitCol)) {
                            const insertIdx = Math.min(5, dynamicColumns.length);
                            dynamicColumns.splice(insertIdx, 0, targetBenefitCol);
                        }
                        lastBenefitTargetByRow[r] = targetBenefitCol;
                    }

                    if (targetBenefitCol && hasImgs) {
                        appendImagesToAttr(targetBenefitCol, rowImgs);
                    }
                }
                // 7. Baris atribut / spesifikasi biasa (Entry Age, Billing Frequency, Insured Entry Age, Premium (Iuran), Sales Tools, After Sales Tools, dll)
                else {
                    const targetSpecKey = cleanLabel || lastSectionLabel;
                    if (targetSpecKey) {
                        const lowerSpecKey = targetSpecKey.toLowerCase().trim();
                        const KNOWN_IDENTITY_OR_ALIAS_SPECS = new Set([
                            "product name", "nama produk", "company name", "nama perusahaan", "company",
                            "par/non par", "product type", "tipe produk",
                            "distribution channel", "channel distribution", "currency", "mata uang",
                            "launch date", "tanggal rilis", "launching date", "release date", "estimation launching date", "estimated launching date", "est. launching date", "est launching date",
                            "entry age", "insured entry age", "age definition",
                            "coverage period", "coverage term",
                            "premium payment period (ppp)", "premium payment period", "premium term", "ppp",
                            "min premium", "premium (iuran)", "iuran", "premium limit", "minimum premium", "premium",
                            "min sum assured", "sum assured limit", "minimum sum assured", "sum assured",
                            "underwriting method", "underwriting",
                            "available riders", "available rider", "riders", "rider",
                            "sales tools", "sales tool", "after sales tools", "after sales tool",
                            "aqd", "akad", "aqd / akad", "akad / aqd", "aqd / contract", "aqd/contract",
                            "alokasi kontribusi", "surplus underwriting", "underwriting surplus",
                            "additional benefit", "additional feature (if any)", "additional feature", "additional features",
                            "other benefit", "other feature", "fitur tambahan", "remarks",
                            ...CORE_FIELD_RULES.flatMap(r => [r.canonical, ...(Array.isArray(r.aliases) ? r.aliases : [])]).map(s => String(s || "").trim().toLowerCase()).filter(Boolean)
                        ]);

                        if (lowerSpecKey === "category system") {
                            // Metadata kategori tampilan (tidak perlu masuk sebagai kolom Non-Benefit baru)
                            continue;
                        } else if (KNOWN_IDENTITY_OR_ALIAS_SPECS.has(lowerSpecKey)) {
                            if (hasText) attrMap[targetSpecKey] = sanitizeValue(val);
                            if (hasImgs) appendImagesToAttr(targetSpecKey, rowImgs);
                        } else {
                            const resolvedNb = resolveKnownNonBenefitName(targetSpecKey);
                            const existingCol = dynamicColumns.find(col => col.toLowerCase() === (resolvedNb || targetSpecKey).toLowerCase());
                            const finalSpecCol = existingCol || resolvedNb || targetSpecKey;

                            if (!isNonBenefitColumn(finalSpecCol) && !isDplkChargeColumn(finalSpecCol)) {
                                NON_BENEFIT_COLUMNS.push(finalSpecCol);
                            }
                            if (!dynamicColumns.includes(finalSpecCol) && (hasText || hasImgs) && !isDplkChargeColumn(finalSpecCol)) {
                                const addBenIdx = dynamicColumns.findIndex(col => col.toLowerCase() === "additional benefit");
                                if (addBenIdx !== -1) dynamicColumns.splice(addBenIdx, 0, finalSpecCol);
                                else dynamicColumns.push(finalSpecCol);
                            }
                            if (hasText) appendTextToAttr(finalSpecCol, val);
                            if (hasImgs) appendImagesToAttr(finalSpecCol, rowImgs);
                        }
                    }
                }
            }

            // Kompres dan sematkan token gambar [IMG:data:image/webp;base64,...] ke masing-masing atribut produk
            Object.keys(attrImagesMap).forEach(attrKey => {
                const imgs = attrImagesMap[attrKey];
                if (!imgs || imgs.length === 0) return;
                const existingText = (attrMap[attrKey] && attrMap[attrKey] !== "-") ? attrMap[attrKey] : "";
                const imgTokens = compressImageGroupToTokens(imgs, existingText.length);
                if (imgTokens.length > 0) {
                    attrMap[attrKey] = existingText
                        ? `${existingText}\n${imgTokens.join("\n")}`
                        : imgTokens.join("\n");
                }
            });

            allParsedProducts.push({ prodName, company, attrMap, sheetName });
        }
    });

    // Jika seluruh sheet adalah Excel Flat / Horizontal
    if (allParsedProducts.length === 0 && flatRawSheets.length > 0) {
        let masterSchema = [
            "Product Name", "Company Name", "Par/Non Par", "Product Type",
            "Distribution Channel", "Currency", ...dynamicColumns
        ];
        let result = [];
        flatRawSheets.forEach(({ headerRow, rawData, sheetName }) => {
            for (let r = 1; r < rawData.length; r++) {
                let row = rawData[r];
                if (!row || row.every(cell => String(cell).trim() === "")) continue;

                let mappedRow = masterSchema.map(colName => {
                    let lowerCol = colName.toLowerCase();
                    let idx = headerRow.findIndex(h => h === lowerCol);
                    if (idx === -1 && lowerCol === "additional benefit") {
                        idx = headerRow.findIndex(h => isAdditionalBenefitLabel(h));
                    } else if (idx === -1 && isAqdLabel(lowerCol)) {
                        idx = headerRow.findIndex(h => isAqdLabel(h));
                    } else if (idx === -1 && isAlokasiKontribusiLabel(lowerCol)) {
                        idx = headerRow.findIndex(h => isAlokasiKontribusiLabel(h));
                    } else if (idx === -1 && isSurplusUnderwritingLabel(lowerCol)) {
                        idx = headerRow.findIndex(h => isSurplusUnderwritingLabel(h));
                    }
                    let val = (idx !== -1 && row[idx] !== undefined) ? String(row[idx]).trim() : "-";
                    return sanitizeValue(val);
                });
                mappedRow.__sheetName = sheetName; // Penanda sheet asal (untuk Target Tab per sheet)
                result.push(mappedRow);
            }
        });
        return result;
    }

    if (allParsedProducts.length === 0) return [];

    // Susun masterSchema setelah semua kolom manfaat & biaya dari Excel terdaftar di dynamicColumns
    let masterSchema = [
        "Product Name", "Company Name", "Par/Non Par", "Product Type",
        "Distribution Channel", "Currency", "Launch Date", ...dynamicColumns
    ];

    let resultRows = allParsedProducts.map(({ prodName, company, attrMap }) => {
        // Helper pencari nilai attrMap secara case-insensitive dengan daftar alias nama baris
        const getByAliases = (aliases) => {
            for (const alias of aliases) {
                const foundKey = Object.keys(attrMap).find(k => k.toLowerCase() === alias.toLowerCase());
                if (foundKey && !isEmptyVal(attrMap[foundKey])) {
                    return sanitizeValue(attrMap[foundKey]);
                }
            }
            return "-";
        };

        return masterSchema.map(colName => {
            const lowerCol = colName.toLowerCase();

            if (colName === "Product Name") return prodName;
            if (colName === "Company Name") return company;
            if (colName === "Par/Non Par") return getByAliases(getConfigAliasesForColumn("Par/Non Par", ["Par / Non Par"]));
            if (colName === "Product Type") return getByAliases(getConfigAliasesForColumn("Product Type", ["Tipe Produk"]));
            if (colName === "Distribution Channel") return getByAliases(getConfigAliasesForColumn("Distribution Channel", ["Channel Distribution"]));
            if (colName === "Currency") return getByAliases(getConfigAliasesForColumn("Currency", ["Mata Uang"]));
            if (colName === "Launch Date" || colName === "Tanggal Rilis") {
                const rawD = getByAliases(getConfigAliasesForColumn("Launch Date", ["Tanggal Rilis", "Launching Date", "Release Date", "Estimation Launching Date", "Estimated Launching Date", "Est. Launching Date", "Est Launching Date"]));
                return rawD !== "-" ? formatDateToDDMMYYYY(rawD) : "-";
            }

            // Pemetaan alias kolom Spesifikasi, Ketentuan Syariah, DPLK, & Fitur Tambahan (mendukung alias dari Config)
            if (isAqdLabel(lowerCol)) {
                return getByAliases(getConfigAliasesForColumn("Aqd", ["Akad", "Aqd / Akad", "Akad / Aqd", "Aqd / Contract", "Aqd/Contract"]));
            }
            if (isAlokasiKontribusiLabel(lowerCol)) {
                return getByAliases(getConfigAliasesForColumn("Alokasi Kontribusi", []));
            }
            if (isSurplusUnderwritingLabel(lowerCol)) {
                return getByAliases(getConfigAliasesForColumn("Surplus Underwriting", ["Underwriting Surplus"]));
            }
            if (lowerCol === "additional benefit") {
                return getByAliases(getConfigAliasesForColumn("Additional Benefit", [
                    "Additional Feature (if any)",
                    "Additional Feature",
                    "Additional Features",
                    "Other Benefit",
                    "Other Feature",
                    "Fitur Tambahan"
                ]));
            }
            if (lowerCol === "entry age") {
                const entryAliases = getConfigAliasesForColumn("Entry Age", ["Insured Entry Age"]).filter(a => a.toLowerCase() !== "age definition");
                const ageVal = getByAliases(entryAliases);
                const ageDef = getByAliases(["Age Definition"]);
                if (ageVal !== "-" && ageDef !== "-") {
                    return `${ageVal}\n(${ageDef})`;
                }
                return ageVal;
            }
            if (lowerCol === "coverage period") {
                return getByAliases(getConfigAliasesForColumn("Coverage Period", ["Coverage Term"]));
            }
            if (lowerCol === "premium payment period (ppp)") {
                return getByAliases(getConfigAliasesForColumn("Premium Payment Period (PPP)", ["Premium Payment Period", "Premium Term", "PPP"]));
            }
            if (lowerCol === "min premium") {
                return getByAliases(getConfigAliasesForColumn("Min Premium", ["Premium (Iuran)", "Iuran", "Premium Limit", "Minimum Premium", "Premium"]));
            }
            if (lowerCol === "min sum assured") {
                return getByAliases(getConfigAliasesForColumn("Min Sum Assured", ["Sum Assured Limit", "Minimum Sum Assured", "Sum Assured"]));
            }
            if (lowerCol === "underwriting method") {
                return getByAliases(getConfigAliasesForColumn("Underwriting Method", ["Underwriting"]));
            }
            if (lowerCol === "available riders") {
                return getByAliases(getConfigAliasesForColumn("Available Riders", ["Available Rider", "Riders", "Rider"]));
            }
            if (lowerCol === "sales tools") {
                return getByAliases(getConfigAliasesForColumn("Sales Tools", ["Sales Tool"]));
            }
            if (lowerCol === "after sales tools") {
                return getByAliases(getConfigAliasesForColumn("After Sales Tools", ["After Sales Tool"]));
            }

            const customColAliases = getConfigAliasesForColumn(colName, []);
            if (customColAliases.length > 1) {
                const viaAlias = getByAliases(customColAliases);
                if (viaAlias !== "-") return viaAlias;
            }
            let matchedKey = Object.keys(attrMap).find(k => k.toLowerCase() === lowerCol);
            return matchedKey ? sanitizeValue(attrMap[matchedKey]) : "-";
        });
    });

    // Penanda sheet asal pada setiap baris (untuk Target Tab per sheet)
    resultRows.forEach((row, i) => { row.__sheetName = allParsedProducts[i].sheetName; });

    return resultRows;
}

async function processAndUploadExcel() {
    const chkOverwrite = document.getElementById("chkOverwriteDuplicates");
    const includeDuplicates = Boolean(chkOverwrite && chkOverwrite.checked);
    const rawRowsToUpload = includeDuplicates
        ? [...parsedUploadRows, ...parsedDuplicateRows]
        : [...parsedUploadRows];

    if (rawRowsToUpload.length === 0) return;
    if (!isBenefitReviewAllValid()) {
        refreshUploadSubmitState();
        return;
    }

    // Terapkan keputusan Review Benefit (gabung ke benefit yang ada / ganti nama / abaikan)
    const { rows: rowsToUpload, aliasMappings } = applyBenefitReviewDecisions(rawRowsToUpload);

    const submitBtn = document.getElementById("btnSubmitUpload");
    const previewArea = document.getElementById("uploadPreviewArea");
    const targetSheet = document.getElementById("uploadTargetSheet").value;

    submitBtn.disabled = true;
    submitBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Sending to Google Sheets...`;

    try {
        // Jika opsi perbarui produk duplikat dicentang, hapus baris lama produk tersebut terlebih dahulu
        if (includeDuplicates && parsedDuplicateRows.length > 0) {
            const dupNamesLower = parsedDuplicateRows.map(r => String(r[0]).toLowerCase().trim());
            const existingTargets = products.filter(p => dupNamesLower.includes(p.name.toLowerCase().trim()));
            if (existingTargets.length > 0) {
                const deletePayload = {
                    action: "delete",
                    items: existingTargets.map(p => ({
                        sheetName: p.category,
                        productName: p.name
                    }))
                };
                await postAuthenticatedAppsScript(deletePayload);
            }
        }

        let baseMasterSchema = [
            "Product Name", "Company Name", "Par/Non Par", "Product Type",
            "Distribution Channel", "Currency", "Launch Date", ...dynamicColumns
        ];

        // Kelompokkan baris berdasarkan Target Tab masing-masing sheet asal (Modern JS: Object.groupBy)
        const rowsByTarget = Object.groupBy(
            rowsToUpload,
            row => resolveSheetTarget(row.__sheetName) || targetSheet
        );

        const uploadSummary = [];
        for (const tabName of Object.keys(rowsByTarget)) {
            const { filteredSchema, filteredRows } = filterSchemaAndRowsForTargetSheet(tabName, baseMasterSchema, rowsByTarget[tabName]);
            await offloadOversizedCellImagesToDrive(filteredRows, (prodName) => {
                submitBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Uploading full image (${escapeHtml(prodName)})...`;
            });
            submitBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Sending to tab ${escapeHtml(tabName)}...`;
            const { finalSchema: masterSchema, finalRows } = expandRowsWithContinuationColumns(filteredSchema, filteredRows);
            await appendToSheetWithHeaderCheck(tabName, masterSchema, finalRows);
            uploadSummary.push({ tabName, count: finalRows.length });
        }

        // Simpan alias nama manfaat & non-benefit hasil Review ke sheet Config (agar upload berikutnya otomatis dikenali)
        let savedAliasCount = 0;
        if (aliasMappings.length > 0) {
            submitBtn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving aliases to Config...`;
            try {
                await postAuthenticatedAppsScript({ action: "save_benefit_aliases", mappings: aliasMappings });
                savedAliasCount = aliasMappings.length;
                aliasMappings.forEach(m => {
                    const aKey = normalizeBenefitNameKey(m.alias);
                    if (String(m.entryType || "").toLowerCase() === "non-benefit") {
                        if (aKey) NON_BENEFIT_ALIAS_MAP[aKey] = m.canonical;
                        if (m.canonical) NON_BENEFIT_CONFIG_CANONICALS.add(normalizeBenefitNameKey(m.canonical));
                    } else {
                        if (aKey) BENEFIT_ALIAS_MAP[aKey] = m.canonical;
                        if (m.canonical) {
                            BENEFIT_CONFIG_CANONICALS.add(normalizeBenefitNameKey(m.canonical));
                            if (!BENEFIT_CONFIG_CANONICAL_LIST.some(b => b.toLowerCase() === String(m.canonical).toLowerCase())) {
                                BENEFIT_CONFIG_CANONICAL_LIST.push(String(m.canonical).trim());
                            }
                        }
                    }
                });
            } catch (aliasErr) {
                if (aliasErr && aliasErr.code === "UNAUTHORIZED") throw aliasErr;
                console.warn("Failed to save aliases to Config:", aliasErr);
            }
        }

        const totalUploaded = uploadSummary.reduce((sum, x) => sum + x.count, 0);
        previewArea.innerHTML = `
                    <p class="font-bold text-green-700 dark:text-green-400 text-sm"><i class="fa-solid fa-circle-check"></i> Successfully Uploaded!</p>
                    <p class="text-xs text-gray-600 dark:text-neutral-300">${totalUploaded} product(s) successfully saved to: ${uploadSummary.map(x => `tab <strong>${escapeHtml(x.tabName)}</strong> (${x.count})`).join(", ")}.</p>
                    ${savedAliasCount > 0 ? `<p class="text-xs text-emerald-700 dark:text-emerald-400"><i class="fa-solid fa-tags me-1"></i>${savedAliasCount} mapping(s) / provision(s) saved to Config.</p>` : ""}
                `;
        resetBenefitReviewState();

        setTimeout(() => {
            submitBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Upload to Google Sheets`;
            closeUploadModal();
            invalidateLocalSummaryCache();
            fetchAllSheetData(true);
        }, 2000);

    } catch (err) {
        console.error(err);
        if (err && err.code === "UNAUTHORIZED") {
            previewArea.innerHTML = `<p class="text-red-500"><i class="fa-solid fa-lock"></i> Session expired. Please sign in again.</p>`;
        } else {
            previewArea.innerHTML = `<p class="text-red-500"><i class="fa-solid fa-circle-xmark"></i> Failed to connect to Google Sheets API. Please check your internet connection.</p>`;
        }
        submitBtn.disabled = false;
        submitBtn.innerHTML = `<i class="fa-solid fa-cloud-arrow-up"></i> Try Again`;
    }
}
// ====================================================
// CONFIG ALIAS MANAGER (LIHAT, TAMBAH, EDIT, HAPUS ALIAS BENEFIT & NON-BENEFIT)
// ====================================================
let aliasManagerEntries = [];
let aliasManagerFilterType = "all";

async function openAliasManagerModal() {
    const modal = document.getElementById("aliasManagerModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    closeAliasEditorForm();
    const searchEl = document.getElementById("aliasManagerSearch");
    if (searchEl) searchEl.value = "";
    setAliasManagerFilter("all", false);

    const listEl = document.getElementById("aliasManagerList");
    const infoEl = document.getElementById("aliasManagerCountInfo");
    if (listEl) {
        listEl.innerHTML = `<div class="py-10 text-center text-gray-500 dark:text-neutral-400"><i class="fa-solid fa-circle-notch fa-spin text-lg text-blue-500 mb-2"></i><p>Loading aliases from Config...</p></div>`;
    }
    if (infoEl) infoEl.textContent = "Loading...";

    try {
        const res = await postAuthenticatedAppsScript({ action: "manage_benefit_aliases", op: "list" });
        aliasManagerEntries = Array.isArray(res.entries) ? res.entries : [];
        renderAliasManagerList();
    } catch (err) {
        console.error("Failed to load aliases:", err);
        if (!err || err.code !== "UNAUTHORIZED") {
            if (listEl) {
                listEl.innerHTML = `<div class="py-8 text-center text-red-500 font-semibold"><i class="fa-solid fa-circle-xmark me-1"></i>Failed to load aliases from server.</div>`;
            }
            showToast("Failed to load aliases from Config.", "error");
        }
    }
}

function closeAliasManagerModal() {
    const modal = document.getElementById("aliasManagerModal");
    if (modal) modal.classList.add("hidden");
}

function setAliasManagerFilter(type, rerender = true) {
    aliasManagerFilterType = type;
    const activeCls = "px-2.5 py-1.5 bg-blue-600 text-white transition";
    const inactiveCls = "px-2.5 py-1.5 bg-white dark:bg-neutral-800 text-gray-700 dark:text-neutral-300 hover:bg-gray-100 dark:hover:bg-neutral-700 transition";
    const btnAll = document.getElementById("aliasFilterAll");
    const btnBen = document.getElementById("aliasFilterBenefit");
    const btnNb = document.getElementById("aliasFilterNonBenefit");
    if (btnAll) btnAll.className = type === "all" ? activeCls : inactiveCls;
    if (btnBen) btnBen.className = type === "benefit" ? activeCls : inactiveCls;
    if (btnNb) btnNb.className = type === "non-benefit" ? activeCls : inactiveCls;
    if (rerender) renderAliasManagerList();
}

function renderAliasManagerList() {
    const listEl = document.getElementById("aliasManagerList");
    const infoEl = document.getElementById("aliasManagerCountInfo");
    if (!listEl) return;

    const q = String(document.getElementById("aliasManagerSearch")?.value || "").trim().toLowerCase();
    const filtered = aliasManagerEntries.filter(entry => {
        if (aliasManagerFilterType !== "all" && entry.groupKey !== aliasManagerFilterType) return false;
        if (!q) return true;
        return entry.canonical.toLowerCase().includes(q)
            || entry.entryType.toLowerCase().includes(q)
            || (entry.aliases || []).some(a => a.toLowerCase().includes(q));
    });

    if (infoEl) {
        infoEl.textContent = `Showing ${filtered.length} of ${aliasManagerEntries.length} Config entries`;
    }

    if (filtered.length === 0) {
        listEl.innerHTML = `<div class="py-8 text-center text-gray-400 dark:text-neutral-500">No matching alias entries found.</div>`;
        return;
    }

    listEl.innerHTML = filtered.map(entry => {
        const isNb = entry.groupKey === "non-benefit";
        const badgeCls = isNb
            ? "bg-slate-100 text-slate-700 border-slate-300 dark:bg-neutral-700 dark:text-neutral-200 dark:border-neutral-600"
            : "bg-blue-50 text-blue-700 border-blue-200 dark:bg-neutral-700/70 dark:text-blue-300 dark:border-neutral-600";
        const aliasChips = (entry.aliases || []).length > 0
            ? entry.aliases.map(a => `<span class="inline-block text-[10px] px-2 py-0.5 rounded-full bg-gray-100 dark:bg-neutral-700 text-gray-700 dark:text-neutral-200 border border-gray-200 dark:border-neutral-600">${escapeHtml(a)}</span>`).join(" ")
            : `<span class="text-[10px] italic text-gray-400 dark:text-neutral-500">No aliases yet</span>`;

        return `
                    <div class="p-2.5 rounded-xl border border-gray-200 dark:border-neutral-700 bg-white dark:bg-neutral-800/80 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div class="min-w-0 flex-1 space-y-1">
                            <div class="flex flex-wrap items-center gap-1.5">
                                <span class="text-[10px] font-bold px-2 py-0.5 rounded-full border ${badgeCls}">${escapeHtml(entry.entryType)}</span>
                                <span class="font-bold text-gray-900 dark:text-neutral-100 text-xs break-words">${escapeHtml(entry.canonical)}</span>
                            </div>
                            <div class="flex flex-wrap items-center gap-1">${aliasChips}</div>
                        </div>
                        <div class="flex items-center gap-1.5 shrink-0 self-end sm:self-center">
                            <button type="button" data-ev-click="a170" data-ev-click-args="${idcmPackArgs([entry.rowNum])}"
                                class="px-2.5 py-1 rounded-lg border border-gray-300 dark:border-neutral-600 text-gray-700 dark:text-neutral-200 hover:bg-gray-100 dark:hover:bg-neutral-700 font-semibold transition">
                                <i class="fa-solid fa-pen-to-square me-1" aria-hidden="true"></i>Edit
                            </button>
                            <button type="button" data-ev-click="a171" data-ev-click-args="${idcmPackArgs([entry.rowNum])}"
                                class="px-2.5 py-1 rounded-lg border border-red-200 dark:border-red-800/70 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-950/40 font-semibold transition"
                                title="Delete entry from Config">
                                <i class="fa-solid fa-trash-can" aria-hidden="true"></i>
                            </button>
                        </div>
                    </div>`;
    }).join("");
}

function openAliasEditorForm(rowNum) {
    const box = document.getElementById("aliasEditorFormBox");
    const titleEl = document.getElementById("aliasEditorFormTitle");
    const rowInput = document.getElementById("aliasEditRowNum");
    const typeInput = document.getElementById("aliasEditType");
    const canonInput = document.getElementById("aliasEditCanonical");
    const aliasInput = document.getElementById("aliasEditAliases");
    const descInput = document.getElementById("aliasEditDesc");
    if (!box) return;

    const entry = rowNum ? aliasManagerEntries.find(e => e.rowNum === rowNum) : null;
    if (entry) {
        titleEl.textContent = `Edit Alias — ${entry.canonical}`;
        rowInput.value = String(entry.rowNum);
        if (typeInput && entry.entryType && !Array.from(typeInput.options).some(o => o.value === entry.entryType)) {
            const opt = document.createElement("option");
            opt.value = entry.entryType;
            opt.textContent = entry.entryType;
            typeInput.appendChild(opt);
        }
        typeInput.value = entry.entryType;
        canonInput.value = entry.canonical;
        aliasInput.value = (entry.aliases || []).join(", ");
        descInput.value = entry.description || "";
    } else {
        titleEl.textContent = "Add New Alias Entry";
        rowInput.value = "";
        typeInput.value = aliasManagerFilterType === "non-benefit" ? "Non-Benefit" : "Benefit";
        canonInput.value = "";
        aliasInput.value = "";
        descInput.value = "";
    }
    box.classList.remove("hidden");
    canonInput.focus();
}

function closeAliasEditorForm() {
    const box = document.getElementById("aliasEditorFormBox");
    if (box) box.classList.add("hidden");
}

function syncClientConfigFromAliasManager(cfg) {
    if (!cfg) return;
    if (typeof applyDynamicConfigFromBackend === "function") {
        applyDynamicConfigFromBackend(cfg);
    }
    if (uploadWorkbookCache) {
        uploadWorkbookCache.baseNonBenefits = NON_BENEFIT_COLUMNS.slice();
    }
    invalidateReviewLookupCaches();
    invalidateLocalSummaryCache();
    if (uploadWorkbookCache && selectedUploadSheets.length > 0) {
        analyzeSelectedUploadSheets();
    }
}

async function saveAliasManagerEntry() {
    const rowNum = parseInt(document.getElementById("aliasEditRowNum")?.value || "0", 10) || 0;
    const entryType = document.getElementById("aliasEditType")?.value || "Benefit";
    const canonical = String(document.getElementById("aliasEditCanonical")?.value || "").trim();
    const aliases = String(document.getElementById("aliasEditAliases")?.value || "").trim();
    const description = String(document.getElementById("aliasEditDesc")?.value || "").trim();

    if (!canonical) {
        showToast("Please enter a Canonical Column Name.", "warning");
        return;
    }

    const btn = document.getElementById("btnSaveAliasEntry");
    if (btn) {
        btn.disabled = true;
        btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Saving...`;
    }

    try {
        const res = await postAuthenticatedAppsScript({
            action: "manage_benefit_aliases",
            op: "upsert",
            rowNum: rowNum || undefined,
            entryType,
            canonical,
            aliases,
            description
        });
        aliasManagerEntries = Array.isArray(res.entries) ? res.entries : [];
        syncClientConfigFromAliasManager(res.config);
        closeAliasEditorForm();
        renderAliasManagerList();
        showToast(`Alias "${canonical}" saved to Config.`, "success");
    } catch (err) {
        console.error("Failed to save alias:", err);
        if (!err || err.code !== "UNAUTHORIZED") {
            showToast("Failed to save alias: " + (err.message || err), "error");
        }
    } finally {
        if (btn) {
            btn.disabled = false;
            btn.innerHTML = `<i class="fa-solid fa-floppy-disk" aria-hidden="true"></i> Save Alias`;
        }
    }
}

async function deleteAliasManagerEntry(rowNum) {
    const entry = aliasManagerEntries.find(e => e.rowNum === rowNum);
    if (!entry) return;
    if (!confirm(`Delete "${entry.canonical}" (${entry.entryType}) and its aliases from Config?`)) return;

    try {
        const res = await postAuthenticatedAppsScript({
            action: "manage_benefit_aliases",
            op: "delete",
            rowNum: entry.rowNum,
            canonical: entry.canonical
        });
        aliasManagerEntries = Array.isArray(res.entries) ? res.entries : [];
        syncClientConfigFromAliasManager(res.config);
        renderAliasManagerList();
        showToast(`Deleted "${entry.canonical}" from Config.`, "success");
    } catch (err) {
        console.error("Failed to delete alias:", err);
        if (!err || err.code !== "UNAUTHORIZED") {
            showToast("Failed to delete alias: " + (err.message || err), "error");
        }
    }
}
// ====================================================
// ADMIN PANEL: MANAGE USERS & ROLES (RBAC)
// ====================================================
let adminUsersEntries = [];
let pendingAdminRoleChange = null;

async function openAdminUsersModal() {
    if (!isAdminUser()) {
        showToast("Akses ditolak: hanya Admin yang dapat mengelola user.", "error");
        return;
    }
    const modal = document.getElementById("adminUsersModal");
    if (!modal) return;
    modal.classList.remove("hidden");
    const listEl = document.getElementById("adminUsersList");
    const infoEl = document.getElementById("adminUsersCountInfo");
    if (listEl) {
        listEl.innerHTML = `<div class="py-10 text-center text-gray-500 dark:text-neutral-400"><i class="fa-solid fa-circle-notch fa-spin text-lg text-amber-500 mb-2"></i><p>Loading registered users...</p></div>`;
    }
    if (infoEl) infoEl.textContent = "Loading...";

    try {
        const res = await postAuthenticatedAppsScript({ action: "admin_manage_users", op: "list" });
        adminUsersEntries = Array.isArray(res.users) ? res.users : [];
        renderAdminUsersList();
    } catch (err) {
        console.error("Failed to load users:", err);
        if (!err || (err.code !== "UNAUTHORIZED" && err.code !== "FORBIDDEN")) {
            if (listEl) listEl.innerHTML = `<div class="py-8 text-center text-red-500 font-semibold">Failed to load users.</div>`;
            showToast("Failed to load registered users.", "error");
        }
    }
}

function closeAdminUsersModal() {
    const modal = document.getElementById("adminUsersModal");
    if (modal) modal.classList.add("hidden");
}

function renderAdminUsersList() {
    const listEl = document.getElementById("adminUsersList");
    const infoEl = document.getElementById("adminUsersCountInfo");
    if (!listEl) return;

    const currentSession = typeof getActiveAuthSession === "function" ? getActiveAuthSession() : null;
    const meLower = String((currentSession && currentSession.username) || "").trim().toLowerCase();
    const adminCount = adminUsersEntries.filter(u => u.role === "admin").length;
    if (infoEl) infoEl.textContent = `Total ${adminUsersEntries.length} user(s)`;

    if (adminUsersEntries.length === 0) {
        listEl.innerHTML = `<div class="py-8 text-center text-gray-400 dark:text-neutral-500">No users found.</div>`;
        return;
    }

    const cardsHtml = adminUsersEntries.map(u => {
        const isMe = u.username.toLowerCase() === meLower;
        const isSoleAdmin = u.role === "admin" && adminCount <= 1;
        const safeUserJs = u.username.replace(/\\/g, "\\\\").replace(/'/g, "\\'");
        return `
                    <div class="p-2.5 rounded-xl border border-gray-200 dark:border-neutral-700 bg-white dark:bg-neutral-800/80 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
                        <div class="min-w-0 flex-1">
                            <div class="flex flex-wrap items-center gap-1.5">
                                <span class="font-bold text-gray-900 dark:text-neutral-100">${escapeHtml(u.username)}</span>
                                ${isMe ? `<span class="text-[10px] px-1.5 py-0.5 rounded bg-emerald-100 dark:bg-emerald-900/40 text-emerald-700 dark:text-emerald-300 font-bold">You</span>` : ""}
                            </div>
                            <div class="text-[10px] text-gray-400 dark:text-neutral-500 mt-0.5">Last login: ${escapeHtml(u.lastLogin || "-")}</div>
                            ${isSoleAdmin ? `<div class="text-[10px] text-amber-600 dark:text-amber-400 mt-0.5"><i class="fa-solid fa-lock me-1" aria-hidden="true"></i>Cannot demote the only Admin account. Promote another user to Admin first.</div>` : ""}
                        </div>
                        <div class="flex items-center gap-1.5 shrink-0 self-end sm:self-center">
                            <select ${isSoleAdmin ? "disabled" : ""} data-ev-change="a172" data-ev-change-args="${idcmPackArgs([idcmUnescapeJsStr(safeUserJs)])}" aria-label="Role for ${escapeHtmlAttr(u.username)}"
                                title="${isSoleAdmin ? "Cannot demote the only Admin account" : "Change user role"}"
                                class="border dark:border-neutral-600 rounded-lg px-2 py-1 text-xs font-semibold bg-white dark:bg-neutral-700 text-gray-900 dark:text-neutral-100 outline-none focus:ring-2 focus:ring-amber-500 disabled:opacity-60 disabled:cursor-not-allowed">
                                <option value="viewer" ${u.role === "viewer" ? "selected" : ""}>Viewer</option>
                                <option value="editor" ${u.role === "editor" ? "selected" : ""}>Editor</option>
                                <option value="admin" ${u.role === "admin" ? "selected" : ""}>Admin</option>
                            </select>
                            ${!isMe ? `<button type="button" data-ev-click="a173" data-ev-click-args="${idcmPackArgs([idcmUnescapeJsStr(safeUserJs)])}" aria-label="Delete user ${escapeHtmlAttr(u.username)}"
                                class="px-2.5 py-1 rounded-lg border border-red-200 dark:border-red-800/70 text-red-600 dark:text-red-400 hover:bg-red-50 dark:hover:bg-red-950/40 font-semibold transition" title="Delete User & Revoke Sessions">
                                <i class="fa-solid fa-trash-can" aria-hidden="true"></i>
                            </button>` : ""}
                        </div>
                    </div>`;
    }).join("");

    const onlySelf = adminUsersEntries.length === 1 && adminUsersEntries[0].username.toLowerCase() === meLower;
    listEl.innerHTML = cardsHtml + (onlySelf
        ? `<div class="mt-2 p-3 rounded-xl border border-dashed border-gray-300 dark:border-neutral-700 text-center text-[11px] text-gray-500 dark:text-neutral-400">No other users registered yet. New accounts signing up with the Invite Code will appear here as <strong>Viewer</strong>.</div>`
        : "");
}

function requestAdminUserRoleChange(username, nextRole) {
    const target = adminUsersEntries.find(u => u.username.toLowerCase() === String(username || "").toLowerCase());
    if (!target || target.role === nextRole) {
        renderAdminUsersList();
        return;
    }
    const adminCount = adminUsersEntries.filter(u => u.role === "admin").length;
    if (target.role === "admin" && nextRole !== "admin" && adminCount <= 1) {
        showToast("Cannot demote the only Admin account. Promote another user to Admin first.", "warning");
        renderAdminUsersList();
        return;
    }

    // Opsi 2: Perubahan yang melibatkan role Admin (promote ke Admin / demote Admin / demote diri sendiri) wajib konfirmasi + password Admin
    const involvesAdmin = (target.role === "admin" || nextRole === "admin");
    if (involvesAdmin) {
        pendingAdminRoleChange = { username: target.username, oldRole: target.role, nextRole };
        const currentSession = typeof getActiveAuthSession === "function" ? getActiveAuthSession() : null;
        const isSelf = currentSession && currentSession.username.toLowerCase() === target.username.toLowerCase();
        const actionWord = nextRole === "admin" ? "promote" : "demote";
        const summaryEl = document.getElementById("adminRoleConfirmSummary");
        if (summaryEl) {
            summaryEl.innerHTML = `Are you sure you want to <strong>${actionWord}</strong> <strong>${escapeHtml(target.username)}${isSelf ? " (You)" : ""}</strong> from <span class="uppercase font-bold">${escapeHtml(target.role)}</span> &rarr; <span class="uppercase font-bold">${escapeHtml(nextRole)}</span>?`
                + (isSelf ? `<br><span class="text-[11px] text-amber-600 dark:text-amber-400">You will be signed out immediately after demoting yourself.</span>` : "");
        }
        const pwdEl = document.getElementById("adminRoleConfirmPassword");
        if (pwdEl) { pwdEl.value = ""; pwdEl.type = "password"; }
        const errEl = document.getElementById("adminRoleConfirmError");
        if (errEl) { errEl.textContent = ""; errEl.classList.add("hidden"); }
        const modal = document.getElementById("adminRoleConfirmModal");
        if (modal) modal.classList.remove("hidden");
        setTimeout(() => { if (pwdEl) pwdEl.focus(); }, 50);
        return;
    }

    executeAdminUserRoleChange(target.username, nextRole, "");
}

function cancelAdminRoleConfirm() {
    pendingAdminRoleChange = null;
    const modal = document.getElementById("adminRoleConfirmModal");
    if (modal) modal.classList.add("hidden");
    const pwdEl = document.getElementById("adminRoleConfirmPassword");
    if (pwdEl) pwdEl.value = "";
    renderAdminUsersList();
}

async function submitAdminRoleConfirm(e) {
    e.preventDefault();
    if (!pendingAdminRoleChange) return;
    const pwdEl = document.getElementById("adminRoleConfirmPassword");
    const errEl = document.getElementById("adminRoleConfirmError");
    const btn = document.getElementById("btnSubmitAdminRoleConfirm");
    const adminPassword = pwdEl ? pwdEl.value : "";

    if (!adminPassword) {
        if (errEl) { errEl.textContent = "Please enter your Admin password."; errEl.classList.remove("hidden"); }
        return;
    }
    if (errEl) { errEl.textContent = ""; errEl.classList.add("hidden"); }
    if (btn) { btn.disabled = true; btn.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Verifying...`; }

    const { username, nextRole } = pendingAdminRoleChange;
    const ok = await executeAdminUserRoleChange(username, nextRole, adminPassword, (msg) => {
        if (errEl) { errEl.textContent = msg; errEl.classList.remove("hidden"); }
    });

    if (btn) { btn.disabled = false; btn.innerHTML = `<i class="fa-solid fa-check" aria-hidden="true"></i> Confirm Role`; }
    if (ok) {
        pendingAdminRoleChange = null;
        const modal = document.getElementById("adminRoleConfirmModal");
        if (modal) modal.classList.add("hidden");
        if (pwdEl) pwdEl.value = "";
    }
}

async function executeAdminUserRoleChange(username, nextRole, adminPassword = "", onInlineError = null) {
    try {
        const payload = {
            action: "admin_manage_users",
            op: "set_role",
            username,
            role: nextRole
        };
        if (adminPassword) payload.adminPassword = adminPassword;
        const res = await postAuthenticatedAppsScript(payload, { requireSuccess: false });
        if (!res || res.status !== "success") {
            const msg = (res && res.message) || "Failed to update user role.";
            if (onInlineError) {
                onInlineError(msg);
            } else {
                showToast(msg, "error");
                renderAdminUsersList();
            }
            return false;
        }
        adminUsersEntries = Array.isArray(res.users) ? res.users : [];
        renderAdminUsersList();
        showToast(`Role "${username}" updated to ${nextRole.toUpperCase()}.`, "success");

        if (res.selfDemoted) {
            closeAdminUsersModal();
            localStorage.removeItem(AUTH_STORAGE_KEY);
            sessionStorage.removeItem(AUTH_STORAGE_KEY);
            clearAllSensitiveClientCaches();
            const overlay = document.getElementById("authGateOverlay");
            if (overlay) overlay.classList.remove("hidden");
            if (typeof showAuthAlert === "function") {
                showAuthAlert(`Role Anda telah diubah menjadi ${nextRole.toUpperCase()}. Silakan Sign In kembali.`, "info");
            }
        }
        return true;
    } catch (err) {
        console.error("Failed to update user role:", err);
        if (!err || (err.code !== "UNAUTHORIZED" && err.code !== "FORBIDDEN")) {
            const msg = (err && err.message) || "Failed to update user role.";
            if (onInlineError) onInlineError(msg);
            else showToast(msg, "error");
        }
        if (!onInlineError) renderAdminUsersList();
        return false;
    }
}

async function deleteAdminUser(username) {
    if (!confirm(`Delete user "${username}" and immediately revoke all active sessions?`)) return;
    try {
        const res = await postAuthenticatedAppsScript({
            action: "admin_manage_users",
            op: "delete_user",
            username
        });
        adminUsersEntries = Array.isArray(res.users) ? res.users : [];
        renderAdminUsersList();
        showToast(`User "${username}" deleted.`, "success");
    } catch (err) {
        console.error("Failed to delete user:", err);
        if (!err || (err.code !== "UNAUTHORIZED" && err.code !== "FORBIDDEN")) {
            showToast("Failed to delete user: " + (err.message || err), "error");
        }
    }
}
