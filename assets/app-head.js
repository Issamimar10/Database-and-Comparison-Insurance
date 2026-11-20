// Initialize theme & color-scheme early to prevent FOUC (modern-web-guidance: dark-mode-toggle)
        (function () {
            try {
                const storedTheme = localStorage.getItem('theme');
                const metaCS = document.querySelector('meta[name="color-scheme"]');
                if (storedTheme === 'dark') {
                    document.documentElement.classList.add('dark');
                    if (metaCS) metaCS.content = 'dark';
                } else if (storedTheme === 'light') {
                    document.documentElement.classList.remove('dark');
                    if (metaCS) metaCS.content = 'light';
                } else {
                    if (metaCS) metaCS.content = 'light dark';
                    if (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches) {
                        document.documentElement.classList.add('dark');
                    }
                }
            } catch (e) { }
        })();

        // ====================================================
        // GLOBAL SECURITY & SESSION TOKEN HELPERS
        // ====================================================
        const AUTH_STORAGE_KEY = "idcm_auth_session_v1";

        function escapeHtml(str) {
            return String(str ?? "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;")
                .replace(/'/g, "&#39;");
        }

        function escapeHtmlAttr(str) {
            return escapeHtml(str);
        }

        // Autentikasi menggunakan HttpOnly Cookie via Cloudflare Pages Function (/api).
        // Tidak ada sessionToken yang disimpan di localStorage/sessionStorage.
        function getAuthSessionToken() {
            return "";
        }

        function getCurrentUserRole() {
            try {
                const raw = sessionStorage.getItem(AUTH_STORAGE_KEY);
                if (!raw) return "viewer";
                const parsed = JSON.parse(raw);
                const r = String((parsed && parsed.role) || "viewer").trim().toLowerCase();
                if (r === "admin" || r === "administrator") return "admin";
                if (r === "editor") return "editor";
                return "viewer";
            } catch (e) {
                return "viewer";
            }
        }

        function canEditData() {
            const r = getCurrentUserRole();
            return r === "editor" || r === "admin";
        }

        function isAdminUser() {
            return getCurrentUserRole() === "admin";
        }

        function applyRoleBasedUI() {
            const role = getCurrentUserRole();
            const canEdit = role === "editor" || role === "admin";
            const isAdmin = role === "admin";

            document.querySelectorAll(".rbac-editor-only").forEach(el => {
                el.classList.toggle("hidden", !canEdit);
            });
            document.querySelectorAll(".rbac-admin-only").forEach(el => {
                el.classList.toggle("hidden", !isAdmin);
            });

            const badge = document.getElementById("navUserRoleBadge");
            if (badge) {
                badge.textContent = role.toUpperCase();
                badge.className = role === "admin"
                    ? "text-[9px] font-extrabold px-1.5 py-0.5 rounded bg-amber-500/20 text-amber-300 border border-amber-400/40 tracking-wider"
                    : (role === "editor"
                        ? "text-[9px] font-extrabold px-1.5 py-0.5 rounded bg-sky-500/20 text-sky-300 border border-sky-400/40 tracking-wider"
                        : "text-[9px] font-extrabold px-1.5 py-0.5 rounded bg-neutral-500/25 text-neutral-300 border border-neutral-400/30 tracking-wider");
            }
        }

        function withSessionToken(payloadObj) {
            return Object.assign({}, payloadObj);
        }

        function clearAllSensitiveClientCaches() {
            try {
                if (typeof invalidateLocalSummaryCache === "function") invalidateLocalSummaryCache();
                localStorage.removeItem("idcm_summary_cache_v1");
                sessionStorage.removeItem("idcm_summary_cache_v1");
            } catch (e) { }
            try {
                if (typeof driveImageCache !== "undefined" && driveImageCache) {
                    Object.keys(driveImageCache).forEach(k => delete driveImageCache[k]);
                }
                if (typeof directDocBlobCache !== "undefined" && directDocBlobCache) {
                    Object.keys(directDocBlobCache).forEach(k => delete directDocBlobCache[k]);
                }
                if (typeof pdfPreviewBlobCache !== "undefined" && pdfPreviewBlobCache) {
                    Object.keys(pdfPreviewBlobCache).forEach(k => delete pdfPreviewBlobCache[k]);
                }
                if (typeof currentInlineDocBlobUrl !== "undefined" && currentInlineDocBlobUrl) {
                    URL.revokeObjectURL(currentInlineDocBlobUrl);
                    currentInlineDocBlobUrl = null;
                }
                if (typeof currentInlineDocMeta !== "undefined") currentInlineDocMeta = null;
                if (typeof products !== "undefined") products = [];
                if (typeof productDocuments !== "undefined") productDocuments = [];
                if (typeof productDriveFolders !== "undefined") productDriveFolders = [];
                if (typeof selectedProductIds !== "undefined") selectedProductIds = [];
                if (typeof currentlyVisibleProducts !== "undefined") currentlyVisibleProducts = [];
                if (typeof currentActiveProduct !== "undefined") currentActiveProduct = null;
                if (typeof isFullProductDataLoaded !== "undefined") isFullProductDataLoaded = false;
                if (typeof isDocumentsLoaded !== "undefined") isDocumentsLoaded = false;
            } catch (e) { }
            const gridEl = document.getElementById("deckGrid");
            if (gridEl) gridEl.replaceChildren();
            const detailBodyEl = document.getElementById("detailBody");
            if (detailBodyEl) detailBodyEl.replaceChildren();
            const inlineViewerEl = document.getElementById("inlineDocViewerContent");
            if (inlineViewerEl) inlineViewerEl.replaceChildren();
            const compareTableEl = document.getElementById("modalTableContainer");
            if (compareTableEl) compareTableEl.replaceChildren();
            const lightboxImg = document.getElementById("imageLightboxTarget");
            if (lightboxImg) lightboxImg.src = "";
            document.querySelectorAll("dialog.idcm-native-dialog[open]").forEach(d => {
                try { d.close(); } catch (e) {}
                d.classList.add("hidden");
            });
        }

        function handleUnauthorizedResponse(resObj) {
            if (resObj && resObj.code === "UNAUTHORIZED") {
                try {
                    localStorage.removeItem(AUTH_STORAGE_KEY);
                    sessionStorage.removeItem(AUTH_STORAGE_KEY);
                } catch (e) { }
                clearAllSensitiveClientCaches();
                const overlay = document.getElementById("authGateOverlay");
                if (overlay) overlay.classList.remove("hidden");
                if (typeof showAuthAlert === "function") {
                    showAuthAlert("Sesi Anda telah berakhir atau tidak valid. Silakan Sign In kembali.", "error");
                }
                return true;
            }
            return false;
        }

        async function postAuthenticatedAppsScript(payloadObj, { requireSuccess = true, explicitToken = null } = {}) {
            const bodyObj = explicitToken !== null
                ? Object.assign({}, payloadObj, { sessionToken: explicitToken })
                : withSessionToken(payloadObj);
            const resp = await fetch(APPS_SCRIPT_URL, {
                method: 'POST',
                credentials: 'same-origin',
                headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                body: JSON.stringify(bodyObj)
            });
            if (!resp.ok) {
                throw new Error(`HTTP ${resp.status} from Apps Script backend`);
            }
            const res = await resp.json();
            if (handleUnauthorizedResponse(res)) {
                const authErr = new Error(res.message || "Session expired or unauthorized.");
                authErr.code = "UNAUTHORIZED";
                throw authErr;
            }
            if (res && res.code === "FORBIDDEN") {
                showToast(res.message || "Akses ditolak untuk peran akun Anda.", "error");
                const forbErr = new Error(res.message || "Forbidden.");
                forbErr.code = "FORBIDDEN";
                throw forbErr;
            }
            if (requireSuccess && (!res || res.status !== "success")) {
                throw new Error((res && res.message) || "Backend operation failed.");
            }
            return res;
        }

        // Modern Web Guidance: Non-intrusive Top-Layer Toast Notifications (popover="manual" + ariaNotify + live-region fallback)
        function showToast(message, type = "auto", options = {}) {
            const text = String(message || "").trim();
            if (!text) return;

            let resolvedType = type;
            if (resolvedType === "auto") {
                const lower = text.toLowerCase();
                if (/^(success|berhasil|saved|tercopy|tersalin)/i.test(lower) || lower.includes("successfully")) {
                    resolvedType = "success";
                } else if (/(failed|error|gagal|invalid|exceed|cannot|tidak valid|tidak diizinkan|ditolak|expired)/i.test(lower)) {
                    resolvedType = "error";
                } else if (/(please |mohon |no |already exists|reserved|not support|restricted|tip:)/i.test(lower)) {
                    resolvedType = "warning";
                } else {
                    resolvedType = "info";
                }
            }

            const region = document.getElementById("toastRegion");
            if (!region) {
                console.warn("[Toast fallback]", resolvedType, text);
                return;
            }

            // Pastikan container toast berada di Top Layer (di atas modal <dialog> yang sedang terbuka)
            try {
                if (typeof region.showPopover === "function") {
                    if (region.matches(":popover-open")) {
                        region.hidePopover();
                    }
                    region.showPopover();
                } else {
                    region.classList.add("is-open-fallback");
                }
            } catch (e) {
                region.classList.add("is-open-fallback");
            }

            const palette = {
                success: {
                    box: "bg-emerald-50 dark:bg-neutral-800 border-emerald-300 dark:border-emerald-700 text-emerald-900 dark:text-emerald-200",
                    icon: "fa-circle-check text-emerald-600 dark:text-emerald-400"
                },
                error: {
                    box: "bg-red-50 dark:bg-neutral-800 border-red-300 dark:border-red-700 text-red-900 dark:text-red-200",
                    icon: "fa-circle-xmark text-red-600 dark:text-red-400"
                },
                warning: {
                    box: "bg-amber-50 dark:bg-neutral-800 border-amber-300 dark:border-amber-700 text-amber-900 dark:text-amber-200",
                    icon: "fa-triangle-exclamation text-amber-600 dark:text-amber-400"
                },
                info: {
                    box: "bg-sky-50 dark:bg-neutral-800 border-sky-300 dark:border-neutral-600 text-sky-900 dark:text-neutral-100",
                    icon: "fa-circle-info text-sky-600 dark:text-sky-400"
                }
            };
            const p = palette[resolvedType] || palette.info;

            // Batasi maksimal 4 toast aktif agar layar tetap rapi
            while (region.children.length >= 4) {
                region.firstElementChild.remove();
            }

            const toastEl = document.createElement("div");
            toastEl.className = `idcm-toast-item pointer-events-auto flex items-start gap-2.5 p-3 rounded-xl border shadow-xl text-xs leading-relaxed ${p.box}`;
            toastEl.setAttribute("role", resolvedType === "error" ? "alert" : "status");

            const iconEl = document.createElement("i");
            iconEl.className = `fa-solid ${p.icon} text-sm mt-0.5 shrink-0`;
            iconEl.setAttribute("aria-hidden", "true");

            const bodyEl = document.createElement("div");
            bodyEl.className = "flex-1 min-w-0 whitespace-pre-line break-words font-medium";
            bodyEl.textContent = text;

            const closeBtn = document.createElement("button");
            closeBtn.type = "button";
            closeBtn.setAttribute("aria-label", "Close notification");
            closeBtn.className = "shrink-0 w-6 h-6 -mr-1 -mt-0.5 rounded-md flex items-center justify-center opacity-70 hover:opacity-100 hover:bg-black/5 dark:hover:bg-white/10 transition text-sm leading-none";
            closeBtn.textContent = "×";

            toastEl.appendChild(iconEl);
            toastEl.appendChild(bodyEl);
            toastEl.appendChild(closeBtn);
            region.appendChild(toastEl);

            // Screen reader announcement: ariaNotify() dengan fallback aria-live region
            try {
                if (typeof document.ariaNotify === "function") {
                    document.ariaNotify(text, { priority: resolvedType === "error" ? "high" : "normal" });
                } else {
                    const srLive = document.getElementById("toastSrLive");
                    if (srLive) {
                        srLive.textContent = "";
                        requestAnimationFrame(() => { srLive.textContent = text; });
                    }
                }
            } catch (e) {}

            let dismissed = false;
            let timerId = null;
            const durationMs = typeof options.duration === "number"
                ? options.duration
                : (resolvedType === "error" ? 6500 : (resolvedType === "warning" ? 5000 : 3800));

            const dismiss = () => {
                if (dismissed) return;
                dismissed = true;
                if (timerId) { clearTimeout(timerId); timerId = null; }
                const wasFocused = toastEl.contains(document.activeElement);
                const nextFocusBtn = toastEl.nextElementSibling?.querySelector("button") || toastEl.previousElementSibling?.querySelector("button");
                toastEl.remove();
                if (region.children.length === 0) {
                    try {
                        if (typeof region.hidePopover === "function" && region.matches(":popover-open")) {
                            region.hidePopover();
                        }
                    } catch (e) {}
                    region.classList.remove("is-open-fallback");
                } else if (wasFocused && nextFocusBtn) {
                    nextFocusBtn.focus();
                }
            };

            const scheduleDismiss = () => {
                if (durationMs <= 0 || dismissed) return;
                if (timerId) clearTimeout(timerId);
                timerId = setTimeout(() => {
                    if (toastEl.contains(document.activeElement) || toastEl.matches(":hover")) {
                        scheduleDismiss();
                        return;
                    }
                    dismiss();
                }, durationMs);
            };

            closeBtn.addEventListener("click", dismiss);
            toastEl.addEventListener("mouseenter", () => { if (timerId) { clearTimeout(timerId); timerId = null; } });
            toastEl.addEventListener("mouseleave", scheduleDismiss);
            toastEl.addEventListener("focusin", () => { if (timerId) { clearTimeout(timerId); timerId = null; } });
            toastEl.addEventListener("focusout", scheduleDismiss);

            scheduleDismiss();
        }

        // Event delegation aman untuk klik gambar tabel & tombol salin gambar (anti DOM-XSS inline onclick)
        document.addEventListener("click", (e) => {
            const copyBtn = e.target.closest('[data-img-action="copy"][data-img-src]');
            if (copyBtn) {
                e.stopPropagation();
                if (typeof copyImageSrcToClipboard === "function") {
                    copyImageSrcToClipboard(copyBtn.dataset.imgSrc, copyBtn);
                }
                return;
            }
            const lightboxTrigger = e.target.closest('[data-img-action="lightbox"][data-img-src]');
            if (lightboxTrigger) {
                if (typeof openImageLightbox === "function") {
                    openImageLightbox(lightboxTrigger.dataset.imgSrc);
                }
            }
        });

        // ====================================================
        // DOCUMENT MANAGEMENT & GOOGLE DRIVE FUNCTIONS
        // ====================================================

        async function openBulkDownloadModal() {
            const selectedProds = products.filter(p => selectedProductIds.includes(p.id));
            if (selectedProds.length === 0) {
                showToast("Please select at least 1 product first to download documents.");
                return;
            }

            const listContainer = document.getElementById("bulkDownloadListContainer");
            const summaryBox = document.getElementById("bulkDownloadSummaryBox");
            const btnStart = document.getElementById("btnStartBulkDownload");
            const btnCancel = document.getElementById("btnCancelBulkDownload");
            const modalEl = document.getElementById("bulkDownloadConfirmModal");

            document.getElementById("bulkDownloadProgressBar").classList.add("hidden");
            btnCancel.disabled = false;

            if (!isDocumentsLoaded) {
                listContainer.innerHTML = `
                    <div class="py-8 text-center text-gray-500 dark:text-neutral-400 text-xs space-y-2">
                        <i class="fa-solid fa-spinner fa-spin text-xl text-emerald-600"></i>
                        <p class="font-medium">Checking documents in Google Drive...</p>
                    </div>`;
                summaryBox.innerHTML = `<span class="text-xs text-gray-500 dark:text-neutral-400">Scanning documents for ${selectedProds.length} selected product(s)...</span>`;
                btnStart.disabled = true;
                modalEl.classList.remove("hidden");

                await ensureDocumentsLoaded();
                if (modalEl.classList.contains("hidden")) return;
            }

            let totalDocsCount = 0;

            listContainer.innerHTML = selectedProds.map(p => {
                const docs = getDocumentsForProduct(p.name);
                totalDocsCount += docs.length;
                return `
                    <div class="flex items-center justify-between p-3 bg-gray-50 dark:bg-neutral-700/50 rounded-xl border border-gray-200 dark:border-neutral-600">
                        <div>
                            <div class="text-xs font-bold text-gray-800 dark:text-neutral-100">${escapeHtml(p.name)}</div>
                            <div class="text-[11px] text-gray-500 dark:text-neutral-400">${escapeHtml(p.company)} &bull; ${escapeHtml(p.category)}</div>
                        </div>
                        <div>
                            ${docs.length > 0 ? `
                                <span class="text-xs font-bold text-emerald-600 dark:text-emerald-400 bg-emerald-100 dark:bg-emerald-950/60 px-2.5 py-1 rounded-full flex items-center gap-1">
                                    <i class="fa-solid fa-check"></i> ${docs.length} Document${docs.length > 1 ? 's' : ''}
                                </span>
                            ` : `
                                <span class="text-xs font-medium text-gray-400 dark:text-neutral-500 bg-gray-100 dark:bg-neutral-800 px-2.5 py-1 rounded-full">
                                    0 Documents
                                </span>
                            `}
                        </div>
                    </div>
                `;
            }).join('');

            summaryBox.innerHTML = `
                <div class="flex items-center gap-2">
                    <i class="fa-solid fa-circle-info text-emerald-600 text-sm"></i>
                    <span>Found a total of <strong>${totalDocsCount} document${totalDocsCount !== 1 ? 's' : ''}</strong> from <strong>${selectedProds.length} selected product${selectedProds.length !== 1 ? 's' : ''}</strong>.</span>
                </div>
            `;

            btnStart.disabled = (totalDocsCount === 0);
            modalEl.classList.remove("hidden");
        }

        function closeBulkDownloadModal() {
            document.getElementById("bulkDownloadConfirmModal").classList.add("hidden");
        }

        async function executeBulkDownloadZip() {
            const selectedProds = products.filter(p => selectedProductIds.includes(p.id));
            const docsToDownload = [];
            selectedProds.forEach(p => {
                const pDocs = getDocumentsForProduct(p.name);
                pDocs.forEach(d => {
                    docsToDownload.push({ prodName: p.name, doc: d });
                });
            });

            if (docsToDownload.length === 0) {
                showToast("No documents available to download for the selected products.");
                return;
            }

            const pBar = document.getElementById("bulkDownloadProgressBar");
            const pText = document.getElementById("bulkDownloadProgressText");
            const pPercent = document.getElementById("bulkDownloadProgressPercent");
            const pFill = document.getElementById("bulkDownloadProgressFill");
            const btnStart = document.getElementById("btnStartBulkDownload");
            const btnCancel = document.getElementById("btnCancelBulkDownload");

            pBar.classList.remove("hidden");
            btnStart.disabled = true;
            btnCancel.disabled = true;

            try {
                const zip = new JSZip();

                for (let i = 0; i < docsToDownload.length; i++) {
                    const item = docsToDownload[i];
                    const percent = Math.round(((i) / docsToDownload.length) * 100);
                    pText.textContent = `Downloading (${i + 1}/${docsToDownload.length}): ${item.doc.fileName}...`;
                    pPercent.textContent = `${percent}%`;
                    pFill.style.width = `${percent}%`;

                    let fileBase64 = null;
                    try {
                        const resp = await fetch(APPS_SCRIPT_URL, {
                            method: 'POST',
                            headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                            body: JSON.stringify(withSessionToken({ action: "get_file_base64", fileId: item.doc.fileId, asPdf: false }))
                        });
                        const resData = await resp.json();
                        if (handleUnauthorizedResponse(resData)) throw new Error(resData.message || "Unauthorized");
                        if (resData && resData.status === "success" && resData.base64) {
                            fileBase64 = resData.base64;
                        }
                    } catch (e) {
                        console.warn("Failed to fetch base64 file via Apps Script:", e);
                    }

                    if (fileBase64) {
                        const sanitizeZipSegment = (val) => String(val || "unnamed")
                            .replace(/[\\/:*?"<>|]+/g, "_")
                            .replace(/^[.\s]+|[.\s]+$/g, "")
                            .replace(/\.{2,}/g, "_") || "unnamed";
                        const cleanFolder = sanitizeZipSegment(item.prodName);
                        const cleanFileName = sanitizeZipSegment(item.doc.fileName);
                        zip.folder(cleanFolder).file(cleanFileName, fileBase64, { base64: true });
                    }
                }

                pText.textContent = "Packing ZIP archive...";
                pPercent.textContent = "95%";
                pFill.style.width = "95%";

                const contentBlob = await zip.generateAsync({ type: "blob" });

                pText.textContent = "Completed! Downloading...";
                pPercent.textContent = "100%";
                pFill.style.width = "100%";

                const link = document.createElement("a");
                link.href = URL.createObjectURL(contentBlob);
                link.download = `Insurance_Documents_${Date.now()}.zip`;
                document.body.appendChild(link);
                link.click();
                document.body.removeChild(link);
                setTimeout(() => URL.revokeObjectURL(link.href), 1000);

                setTimeout(() => {
                    closeBulkDownloadModal();
                }, 1200);

            } catch (err) {
                console.error(err);
                showToast("An error occurred while creating the ZIP archive: " + err.message);
            } finally {
                btnStart.disabled = false;
                btnCancel.disabled = false;
            }
        }

        function openBulkDocUploadModal() {
            bulkDocStagingFiles = [];
            document.getElementById("bulkDocFileInput").value = "";
            document.getElementById("bulkDocPreviewArea").classList.add("hidden");
            document.getElementById("bulkDocUploadProgressBar").classList.add("hidden");
            document.getElementById("btnSubmitBulkDocUpload").disabled = true;
            document.getElementById("bulkDocUploadModal").classList.remove("hidden");
            ensureDocumentsLoaded();
        }

        function closeBulkDocUploadModal() {
            document.getElementById("bulkDocUploadModal").classList.add("hidden");
        }

        function handleBulkDocDragOver(e) {
            e.preventDefault();
            e.stopPropagation();
            const dropzone = document.getElementById("bulkDocDropzone");
            if (dropzone) {
                dropzone.classList.add("border-indigo-500", "bg-indigo-50/70", "dark:bg-neutral-800/80");
            }
        }

        function handleBulkDocDragLeave(e) {
            e.preventDefault();
            e.stopPropagation();
            const dropzone = document.getElementById("bulkDocDropzone");
            if (dropzone) {
                dropzone.classList.remove("border-indigo-500", "bg-indigo-50/70", "dark:bg-neutral-800/80");
            }
        }

        function handleBulkDocDrop(e) {
            e.preventDefault();
            e.stopPropagation();
            const dropzone = document.getElementById("bulkDocDropzone");
            if (dropzone) {
                dropzone.classList.remove("border-indigo-500", "bg-indigo-50/70", "dark:bg-neutral-800/80");
            }
            if (e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                handleBulkDocFileSelect(e.dataTransfer.files);
            }
        }

        function handleBulkDocFileSelect(inputOrFiles) {
            const fileList = (inputOrFiles && inputOrFiles.files) ? inputOrFiles.files : inputOrFiles;
            if (!fileList || fileList.length === 0) return;
            bulkDocStagingFiles = [];

            const productNames = products.map(p => p.name);

            Array.from(fileList).forEach((file, idx) => {
                const originalName = file.name;
                const dotIdx = originalName.lastIndexOf(".");
                const baseName = dotIdx !== -1 ? originalName.substring(0, dotIdx) : originalName;

                let detectedProdName = "";

                const underscoreIdx = baseName.indexOf("_");
                if (underscoreIdx > 0) {
                    const pPart = baseName.substring(0, underscoreIdx).trim();

                    const matched = productNames.find(pn => pn.toLowerCase() === pPart.toLowerCase()) ||
                        productNames.find(pn => pn.toLowerCase().includes(pPart.toLowerCase()) || pPart.toLowerCase().includes(pn.toLowerCase()));
                    if (matched) {
                        detectedProdName = matched;
                    }
                } else {
                    const matched = productNames.find(pn => baseName.toLowerCase().includes(pn.toLowerCase()));
                    if (matched) detectedProdName = matched;
                }

                bulkDocStagingFiles.push({
                    fileIndex: idx,
                    file: file,
                    fileName: originalName,
                    targetProdName: detectedProdName || (productNames[0] || ""),
                    docType: "",
                    size: file.size,
                    isMatched: !!detectedProdName
                });
            });

            renderBulkDocStagingPreview();
        }

        function renderBulkDocStagingPreview() {
            const previewArea = document.getElementById("bulkDocPreviewArea");
            const listEl = document.getElementById("bulkDocMatchList");
            const badgeEl = document.getElementById("bulkDocFileCountBadge");
            const submitBtn = document.getElementById("btnSubmitBulkDocUpload");

            if (bulkDocStagingFiles.length === 0) {
                previewArea.classList.add("hidden");
                submitBtn.disabled = true;
                return;
            }

            badgeEl.textContent = `${bulkDocStagingFiles.length} file${bulkDocStagingFiles.length !== 1 ? 's' : ''} selected`;
            previewArea.classList.remove("hidden");
            submitBtn.disabled = false;

            const productNames = products.map(p => p.name);

            listEl.innerHTML = bulkDocStagingFiles.map((item, idx) => {
                const iconInfo = getDocFileIcon(item.fileName);
                return `
                    <div class="p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5 text-xs">
                        <div class="flex items-center gap-2.5 min-w-0 flex-1">
                            <div class="w-7 h-7 rounded-lg flex items-center justify-center shrink-0 border ${iconInfo.color}">
                                <i class="fa-solid ${iconInfo.icon}"></i>
                            </div>
                            <div class="min-w-0">
                                <div class="font-bold text-gray-800 dark:text-neutral-200 truncate" title="${escapeHtmlAttr(item.fileName)}">${escapeHtml(item.fileName)}</div>
                                <div class="text-[11px] text-gray-500 dark:text-neutral-400">${formatFileSize(item.size)}</div>
                            </div>
                        </div>
                        <div class="flex flex-wrap items-center gap-2 w-full sm:w-auto">
                            <div class="w-full sm:w-56 flex-1 sm:flex-initial">
                                <select data-ev-change="a107" data-ev-change-args="${idcmPackArgs([idx])}" class="w-full border dark:border-neutral-600 bg-white dark:bg-neutral-700 text-gray-800 dark:text-neutral-200 rounded-lg p-2 sm:p-1.5 text-xs outline-none min-h-[38px] sm:min-h-0">
                                    ${productNames.map(pn => `
                                        <option value="${escapeHtmlAttr(pn)}" ${pn === item.targetProdName ? 'selected' : ''}>${escapeHtml(pn)}</option>
                                    `).join('')}
                                </select>
                            </div>
                            <span class="${item.isMatched ? 'text-emerald-600 dark:text-emerald-400 bg-emerald-50 dark:bg-emerald-950/40' : 'text-amber-600 dark:text-amber-400 bg-amber-50 dark:bg-amber-950/40'} text-[10px] font-semibold px-2 py-1 rounded">
                                ${item.isMatched ? 'Matched ✓' : 'Manual'}
                            </span>
                        </div>
                    </div>
                `;
            }).join('');
        }

        function updateStagingProdName(idx, val) {
            if (bulkDocStagingFiles[idx]) {
                bulkDocStagingFiles[idx].targetProdName = val;
                bulkDocStagingFiles[idx].isMatched = true;
            }
        }

        function readFileAsBase64(file) {
            return new Promise((resolve, reject) => {
                const reader = new FileReader();
                reader.onload = () => {
                    const res = reader.result;
                    const base64 = res.split(',')[1];
                    resolve(base64);
                };
                reader.onerror = reject;
                reader.readAsDataURL(file);
            });
        }

        async function startBulkDocUploadToDrive() {
            if (bulkDocStagingFiles.length === 0) return;

            const pBar = document.getElementById("bulkDocUploadProgressBar");
            const pText = document.getElementById("bulkDocUploadStatusText");
            const pPercent = document.getElementById("bulkDocUploadPercent");
            const pFill = document.getElementById("bulkDocUploadProgressFill");
            const btnSubmit = document.getElementById("btnSubmitBulkDocUpload");
            const btnCancel = document.getElementById("btnCancelBulkDocUpload");

            pBar.classList.remove("hidden");
            btnSubmit.disabled = true;
            btnCancel.disabled = true;

            let successCount = 0;

            for (let i = 0; i < bulkDocStagingFiles.length; i++) {
                const item = bulkDocStagingFiles[i];
                const pct = Math.round(((i) / bulkDocStagingFiles.length) * 100);
                pText.textContent = `Uploading (${i + 1}/${bulkDocStagingFiles.length}): ${item.fileName}...`;
                pPercent.textContent = `${pct}%`;
                pFill.style.width = `${pct}%`;

                try {
                    const b64 = await readFileAsBase64(item.file);
                    const payload = {
                        action: "upload_document",
                        productName: item.targetProdName,
                        docType: "",
                        fileName: item.fileName,
                        mimeType: item.file.type || "application/octet-stream",
                        base64Data: b64
                    };

                    const resp = await fetch(APPS_SCRIPT_URL, {
                        method: 'POST',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify(withSessionToken(payload))
                    });

                    const res = await resp.json();
                    if (handleUnauthorizedResponse(res)) break;
                    if (res && res.status === "success") {
                        successCount++;
                        productDocuments.push({
                            productName: res.productName,
                            docType: "",
                            fileName: res.fileName,
                            fileId: res.fileId,
                            driveUrl: res.viewUrl,
                            downloadUrl: res.downloadUrl,
                            fileSize: res.fileSize,
                            uploadDate: new Date().toISOString()
                        });
                    }
                } catch (e) {
                    console.error("Failed to upload file:", item.fileName, e);
                }
            }

            pText.textContent = `Completed! Successfully uploaded ${successCount} document${successCount !== 1 ? 's' : ''} to Google Drive.`;
            pPercent.textContent = "100%";
            pFill.style.width = "100%";

            renderDeck(currentlyVisibleProducts);

            setTimeout(() => {
                closeBulkDocUploadModal();
                showToast(`Success! ${successCount} document${successCount !== 1 ? 's' : ''} uploaded and linked to their respective products in Google Drive.`);
            }, 1000);
        }

        function openSingleDocUploadModal(productName) {
            document.getElementById("singleDocTargetProdName").value = productName;
            document.getElementById("singleDocFileInput").value = "";
            const statusEl = document.getElementById("singleDocUploadStatus");
            statusEl.className = "hidden p-2 rounded-lg text-[11px] font-medium";
            statusEl.textContent = "";
            document.getElementById("btnSubmitSingleDoc").disabled = false;
            document.getElementById("btnSubmitSingleDoc").innerHTML = `<i class="fa-solid fa-upload"></i> Upload`;
            document.getElementById("singleDocUploadModal").classList.remove("hidden");
        }

        function closeSingleDocUploadModal() {
            document.getElementById("singleDocUploadModal").classList.add("hidden");
        }

        function handleSingleDocFileChange(input) {
            // Kept for file input change event compatibility
        }

        async function submitSingleDocUpload() {
            const prodName = document.getElementById("singleDocTargetProdName").value.trim();
            const fileInput = document.getElementById("singleDocFileInput");
            const statusEl = document.getElementById("singleDocUploadStatus");
            const btnSubmit = document.getElementById("btnSubmitSingleDoc");

            if (!fileInput.files || fileInput.files.length === 0) {
                showToast("Please select a document file first.");
                return;
            }

            const file = fileInput.files[0];
            btnSubmit.disabled = true;
            btnSubmit.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> Uploading...`;

            statusEl.className = "p-2 rounded-lg text-[11px] font-medium bg-blue-50 dark:bg-neutral-800/80 text-blue-700 dark:text-blue-300 block";
            statusEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin me-1"></i> Uploading file to Google Drive...`;

            try {
                const b64 = await readFileAsBase64(file);
                const payload = {
                    action: "upload_document",
                    productName: prodName,
                    docType: "",
                    fileName: file.name,
                    mimeType: file.type || "application/octet-stream",
                    base64Data: b64
                };

                const resp = await fetch(APPS_SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify(withSessionToken(payload))
                });

                const res = await resp.json();
                if (handleUnauthorizedResponse(res)) {
                    throw new Error(res.message || "Session expired");
                }
                if (res && res.status === "success") {
                    productDocuments.push({
                        productName: res.productName,
                        docType: "",
                        fileName: res.fileName,
                        fileId: res.fileId,
                        driveUrl: res.viewUrl,
                        downloadUrl: res.downloadUrl,
                        fileSize: res.fileSize,
                        uploadDate: new Date().toISOString()
                    });

                    statusEl.className = "p-2 rounded-lg text-[11px] font-medium bg-green-50 dark:bg-green-950/40 text-green-700 dark:text-green-300 block";
                    statusEl.textContent = "Successfully uploaded to Google Drive!";

                    if (currentActiveProduct && currentActiveProduct.name === prodName) {
                        renderDetailBodyView(currentActiveProduct);
                    }
                    renderDeck(currentlyVisibleProducts);

                    setTimeout(() => {
                        closeSingleDocUploadModal();
                    }, 800);
                } else {
                    throw new Error((res && res.message) || "Upload failed");
                }
            } catch (e) {
                console.error(e);
                statusEl.className = "p-2 rounded-lg text-[11px] font-medium bg-red-50 dark:bg-red-950/40 text-red-700 dark:text-red-300 block";
                statusEl.textContent = "Upload failed: " + e.message;
                btnSubmit.disabled = false;
                btnSubmit.innerHTML = `<i class="fa-solid fa-upload"></i> Retry`;
            }
        }

        async function deleteSingleProductDoc(fileId, productName) {
            const currentDocs = getDocumentsForProduct(productName);
            const isLastDoc = currentDocs.length === 1 && currentDocs[0].fileId === fileId;
            const confirmMsg = isLastDoc
                ? `Delete this last document from Google Drive?\n\nSince no documents will remain for "${productName}", its product folder in Google Drive will also be deleted automatically.`
                : `Delete this document from Google Drive?`;

            if (!confirm(confirmMsg)) return;

            try {
                const delResp = await fetch(APPS_SCRIPT_URL, {
                    method: 'POST',
                    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                    body: JSON.stringify(withSessionToken({
                        action: "delete_document",
                        fileId: fileId,
                        productName: productName,
                        deleteEmptyParentFolder: true
                    }))
                });
                const delJson = await delResp.json().catch(() => null);
                if (handleUnauthorizedResponse(delJson)) return;
                if (delJson && delJson.status === "error") {
                    throw new Error(delJson.message || "Failed to delete document.");
                }

                productDocuments = productDocuments.filter(d => d.fileId !== fileId);

                if (currentActiveProduct) {
                    renderDetailBodyView(currentActiveProduct);
                }
                renderDeck(currentlyVisibleProducts);

            } catch (e) {
                console.error("Failed to delete document:", e);
                showToast("Failed to delete document: " + e.message);
            }
        }

        // ====================================================
        // IN-APP DOCUMENT VIEWER (INLINE WEB PREVIEW & DIRECT ORIGINAL FILE DOWNLOAD)
        // ====================================================

        let currentInlineDocBlobUrl = null;
        let currentInlineDocMeta = null;
        let currentInlineExcelWorkbook = null;
        let currentInlineExcelSheetIdx = 0;
        let currentInlineExcelZoom = 1;
        const directDocBlobCache = {};
        const pdfPreviewBlobCache = {};

        function triggerBlobDownload(blob, fileName) {
            const safeName = String(fileName || "document").trim();
            const blobUrl = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = blobUrl;
            a.download = safeName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            setTimeout(() => URL.revokeObjectURL(blobUrl), 10000);
        }

        function base64ToUint8Array(b64Str) {
            const byteCharacters = atob(b64Str);
            const byteNumbers = new Array(byteCharacters.length);
            for (let i = 0; i < byteCharacters.length; i++) {
                byteNumbers[i] = byteCharacters.charCodeAt(i);
            }
            return new Uint8Array(byteNumbers);
        }

        function inferMimeTypeFromFileName(fileName, fallbackMime) {
            const ext = String(fileName || "").split('.').pop().toLowerCase();
            const map = {
                pdf: "application/pdf",
                xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                xlsm: "application/vnd.ms-excel.sheet.macroEnabled.12",
                xls: "application/vnd.ms-excel",
                csv: "text/csv",
                docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                doc: "application/msword",
                pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
                ppt: "application/vnd.ms-powerpoint",
                png: "image/png",
                jpg: "image/jpeg",
                jpeg: "image/jpeg",
                webp: "image/webp"
            };
            return map[ext] || fallbackMime || "application/octet-stream";
        }

        function isSafeGoogleDriveUrl(url) {
            return typeof url === "string" && /^https:\/\/(?:drive|docs)\.google\.com\//i.test(url.trim());
        }

        async function downloadDocumentDirect(fileId, fileName, btnEl = null, fallbackUrl = "") {
            const safeFallbackUrl = isSafeGoogleDriveUrl(fallbackUrl) ? fallbackUrl.trim() : "";
            if (!fileId) {
                if (safeFallbackUrl) window.open(safeFallbackUrl, "_blank", "noopener,noreferrer");
                return;
            }

            if (directDocBlobCache[fileId] && directDocBlobCache[fileId].blob) {
                triggerBlobDownload(directDocBlobCache[fileId].blob, fileName || directDocBlobCache[fileId].fileName);
                return;
            }

            const originalHtml = btnEl ? btnEl.innerHTML : "";
            if (btnEl) {
                btnEl.disabled = true;
                btnEl.innerHTML = `<i class="fa-solid fa-spinner fa-spin"></i> <span>Downloading...</span>`;
            }

            let wasUnauthorized = false;
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
                if (handleUnauthorizedResponse(data)) {
                    wasUnauthorized = true;
                    throw new Error(data.message || "Session expired.");
                }
                if (!data || data.status !== "success" || !data.base64) {
                    throw new Error((data && data.message) || "Failed to retrieve file from server.");
                }

                const finalName = fileName || data.fileName || "document";
                const mimeType = inferMimeTypeFromFileName(finalName, data.mimeType);
                const byteArray = base64ToUint8Array(data.base64);
                const blob = new Blob([byteArray], { type: mimeType });

                directDocBlobCache[fileId] = { blob, fileName: finalName, mimeType, byteArray };
                triggerBlobDownload(blob, finalName);

                if (btnEl) {
                    btnEl.innerHTML = `<i class="fa-solid fa-check"></i> <span>Downloaded!</span>`;
                    setTimeout(() => {
                        btnEl.disabled = false;
                        btnEl.innerHTML = originalHtml;
                    }, 1200);
                }
            } catch (err) {
                console.error("Failed to download file directly:", err);
                if (btnEl) {
                    btnEl.disabled = false;
                    btnEl.innerHTML = originalHtml;
                }
                if (!wasUnauthorized) {
                    const targetUrl = safeFallbackUrl || `https://drive.google.com/uc?export=download&id=${encodeURIComponent(fileId)}`;
                    window.open(targetUrl, "_blank", "noopener,noreferrer");
                }
            }
        }

        function handleInlineDocViewerDownload(btnEl) {
            if (!currentInlineDocMeta) return;
            downloadDocumentDirect(
                currentInlineDocMeta.fileId,
                currentInlineDocMeta.fileName,
                btnEl,
                currentInlineDocMeta.downloadUrl
            );
        }

        function closeInlineDocViewer() {
            const modal = document.getElementById("inlineDocViewerModal");
            const content = document.getElementById("inlineDocViewerContent");
            if (modal) modal.classList.add("hidden");
            if (content) content.replaceChildren();
            if (currentInlineDocBlobUrl) {
                URL.revokeObjectURL(currentInlineDocBlobUrl);
                currentInlineDocBlobUrl = null;
            }
            currentInlineExcelWorkbook = null;
        }

        function renderInlineExcelSheetTable(ws) {
            const xlsxLib = window.XLSX_CORE || window.XLSX;
            if (!ws || !xlsxLib) {
                return `<div class="p-8 text-center text-gray-500 dark:text-neutral-400 text-xs">Empty sheet.</div>`;
            }

            let maxRow = 0;
            let maxCol = 0;
            let hasAnyCell = false;
            Object.keys(ws).forEach(key => {
                if (key[0] === '!') return;
                const cell = ws[key];
                if (cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== "") {
                    const dec = xlsxLib.utils.decode_cell(key);
                    if (dec.r <= 1500 && dec.c <= 120) {
                        if (dec.r > maxRow) maxRow = dec.r;
                        if (dec.c > maxCol) maxCol = dec.c;
                        hasAnyCell = true;
                    }
                }
            });

            const merges = Array.isArray(ws['!merges']) ? ws['!merges'] : [];
            merges.forEach(m => {
                if (m.e.r <= 1500 && m.e.c <= 120) {
                    if (m.e.r > maxRow) maxRow = m.e.r;
                    if (m.e.c > maxCol) maxCol = m.e.c;
                }
            });

            if (!hasAnyCell && merges.length === 0) {
                return `<div class="p-8 text-center text-gray-500 dark:text-neutral-400 text-xs">This sheet has no data.</div>`;
            }

            const mergeStartMap = {};
            const mergedCoveredSet = new Set();
            merges.forEach(m => {
                const startKey = `${m.s.r},${m.s.c}`;
                const rowspan = m.e.r - m.s.r + 1;
                const colspan = m.e.c - m.s.c + 1;
                mergeStartMap[startKey] = { rowspan, colspan };
                for (let r = m.s.r; r <= m.e.r; r++) {
                    for (let c = m.s.c; c <= m.e.c; c++) {
                        if (r === m.s.r && c === m.s.c) continue;
                        mergedCoveredSet.add(`${r},${c}`);
                    }
                }
            });

            const esc = (s) => String(s ?? "")
                .replace(/&/g, "&amp;")
                .replace(/</g, "&lt;")
                .replace(/>/g, "&gt;")
                .replace(/"/g, "&quot;");

            let colHeadersHtml = `<th class="sticky top-0 left-0 z-20 bg-slate-200 dark:bg-neutral-800 border border-gray-300 dark:border-neutral-600 px-2 py-1 text-[10px] font-bold text-gray-600 dark:text-neutral-300 text-center w-10">#</th>`;
            for (let c = 0; c <= maxCol; c++) {
                const colLetter = xlsxLib.utils.encode_col(c);
                colHeadersHtml += `<th class="sticky top-0 z-10 bg-slate-100 dark:bg-neutral-800 border border-gray-300 dark:border-neutral-600 px-3 py-1 text-[11px] font-bold text-gray-700 dark:text-neutral-300 text-center min-w-[120px]">${colLetter}</th>`;
            }

            let rowsHtml = "";
            for (let r = 0; r <= maxRow; r++) {
                let rowCellsHtml = `<td class="sticky left-0 z-10 bg-slate-100 dark:bg-neutral-800 border border-gray-300 dark:border-neutral-600 px-2 py-1 text-[10px] font-bold text-gray-500 dark:text-neutral-400 text-center select-none">${r + 1}</td>`;
                for (let c = 0; c <= maxCol; c++) {
                    const coordKey = `${r},${c}`;
                    if (mergedCoveredSet.has(coordKey)) continue;

                    const cellRef = xlsxLib.utils.encode_cell({ r, c });
                    const cellObj = ws[cellRef];
                    const formattedText = cellObj ? (cellObj.w !== undefined ? String(cellObj.w) : String(cellObj.v ?? "")) : "";
                    const mergeInfo = mergeStartMap[coordKey];
                    const spanAttrs = mergeInfo
                        ? `${mergeInfo.rowspan > 1 ? ` rowspan="${mergeInfo.rowspan}"` : ''}${mergeInfo.colspan > 1 ? ` colspan="${mergeInfo.colspan}"` : ''}`
                        : '';
                    const isFirstColOrHeader = (r === 0 || c === 0);
                    const cellTone = isFirstColOrHeader && formattedText.trim()
                        ? "bg-slate-50/90 dark:bg-neutral-800/70 font-semibold text-gray-900 dark:text-neutral-100"
                        : "bg-white dark:bg-neutral-900 text-gray-800 dark:text-neutral-200";

                    rowCellsHtml += `<td${spanAttrs} class="border border-gray-200 dark:border-neutral-700 px-2.5 py-1.5 text-xs align-top whitespace-pre-wrap break-words max-w-[360px] ${cellTone}">${esc(formattedText)}</td>`;
                }
                rowsHtml += `<tr>${rowCellsHtml}</tr>`;
            }

            return `
                <table id="inlineExcelPreviewTable" class="border-collapse text-left w-max min-w-full" style="zoom: ${currentInlineExcelZoom};">
                    <thead><tr>${colHeadersHtml}</tr></thead>
                    <tbody>${rowsHtml}</tbody>
                </table>
            `;
        }

        function renderInlineExcelWorkbook(wb, sheetIdx = 0) {
            const content = document.getElementById("inlineDocViewerContent");
            if (!content || !wb || !wb.SheetNames || wb.SheetNames.length === 0) return;

            currentInlineExcelSheetIdx = Math.max(0, Math.min(wb.SheetNames.length - 1, sheetIdx));
            const activeSheetName = wb.SheetNames[currentInlineExcelSheetIdx];
            const ws = wb.Sheets[activeSheetName];

            const tabsHtml = wb.SheetNames.map((sName, idx) => {
                const isActive = idx === currentInlineExcelSheetIdx;
                return `
                    <button type="button" data-ev-click="a108" data-ev-click-args="${idcmPackArgs([idx])}"
                        class="px-3 py-1.5 text-xs font-bold rounded-t-lg border-t border-x transition flex items-center gap-1.5 shrink-0 ${isActive
                        ? 'bg-emerald-600 text-white border-emerald-600 shadow-sm'
                        : 'bg-gray-200 dark:bg-neutral-800 text-gray-700 dark:text-neutral-300 border-gray-300 dark:border-neutral-700 hover:bg-gray-300 dark:hover:bg-neutral-700'}">
                        <i class="fa-solid fa-table-cells text-[10px]"></i>
                        <span>${escapeHtml(sName)}</span>
                    </button>
                `;
            }).join('');

            content.innerHTML = `
                <div class="w-full h-full flex flex-col bg-white dark:bg-neutral-900 rounded-xl border border-gray-300 dark:border-neutral-700 overflow-hidden shadow-inner">
                    <div class="px-2.5 sm:px-3 pt-2 bg-gray-100 dark:bg-neutral-900 border-b border-gray-300 dark:border-neutral-700 flex flex-wrap items-center justify-between gap-2 shrink-0">
                        <div class="flex items-center gap-1 overflow-x-auto no-scrollbar max-w-full">
                            ${tabsHtml}
                        </div>
                        <div class="flex items-center gap-1 pb-1.5 shrink-0 ms-auto">
                            <button type="button" data-ev-click="a109" aria-label="Zoom Out Table" class="px-2.5 py-1 sm:px-2 sm:py-0.5 min-w-[34px] min-h-[34px] sm:min-w-0 sm:min-h-0 text-xs font-bold rounded bg-white dark:bg-neutral-800 border border-gray-300 dark:border-neutral-600 text-gray-700 dark:text-neutral-200 hover:bg-gray-50" title="Zoom Out Table">-</button>
                            <span id="inlineExcelZoomBadge" class="text-[11px] font-bold text-gray-700 dark:text-neutral-300 px-1.5">${Math.round(currentInlineExcelZoom * 100)}%</span>
                            <button type="button" data-ev-click="a110" aria-label="Zoom In Table" class="px-2.5 py-1 sm:px-2 sm:py-0.5 min-w-[34px] min-h-[34px] sm:min-w-0 sm:min-h-0 text-xs font-bold rounded bg-white dark:bg-neutral-800 border border-gray-300 dark:border-neutral-600 text-gray-700 dark:text-neutral-200 hover:bg-gray-50" title="Zoom In Table">+</button>
                        </div>
                    </div>
                    <div id="inlineExcelTableViewport" tabindex="0" role="region" aria-label="Excel Sheet Preview" class="flex-1 overflow-auto">
                        ${renderInlineExcelSheetTable(ws)}
                    </div>
                </div>
            `;
        }

        function switchInlineExcelSheet(idx) {
            if (!currentInlineExcelWorkbook) return;
            renderInlineExcelWorkbook(currentInlineExcelWorkbook, idx);
        }

        function adjustInlineExcelZoom(delta) {
            currentInlineExcelZoom = Math.min(2.0, Math.max(0.5, Math.round((currentInlineExcelZoom + delta) * 100) / 100));
            const tbl = document.getElementById("inlineExcelPreviewTable");
            const badge = document.getElementById("inlineExcelZoomBadge");
            if (tbl) tbl.style.zoom = String(currentInlineExcelZoom);
            if (badge) badge.textContent = `${Math.round(currentInlineExcelZoom * 100)}%`;
        }

        async function previewDocumentInline(fileId, fileName, docType, downloadUrl) {
            const modal = document.getElementById("inlineDocViewerModal");
            const titleEl = document.getElementById("inlineDocViewerTitle");
            const badgeEl = document.getElementById("inlineDocViewerBadge");
            const iconEl = document.getElementById("inlineDocViewerIcon");
            const loader = document.getElementById("inlineDocViewerLoader");
            const loaderText = document.getElementById("inlineDocViewerLoaderText");
            const errorBox = document.getElementById("inlineDocViewerError");
            const errorMsg = document.getElementById("inlineDocViewerErrorMsg");
            const content = document.getElementById("inlineDocViewerContent");

            if (!modal) return;

            if (currentInlineDocBlobUrl) {
                URL.revokeObjectURL(currentInlineDocBlobUrl);
                currentInlineDocBlobUrl = null;
            }
            currentInlineExcelWorkbook = null;
            currentInlineExcelZoom = 1;

            const effectiveDlUrl = downloadUrl || `https://drive.google.com/uc?export=download&id=${fileId}`;
            currentInlineDocMeta = {
                fileId,
                fileName: fileName || "Document",
                docType: docType || "",
                downloadUrl: effectiveDlUrl
            };

            titleEl.textContent = fileName || "Document Preview";
            titleEl.title = fileName || "";
            if (badgeEl) {
                badgeEl.classList.add("hidden");
            }

            const fileIconInfo = getDocFileIcon(fileName);
            iconEl.className = `w-8 h-8 rounded-lg flex items-center justify-center shrink-0 border ${fileIconInfo.color} text-sm`;
            iconEl.innerHTML = `<i class="fa-solid ${fileIconInfo.icon}"></i>`;

            content.replaceChildren();
            errorBox.classList.add("hidden");
            loader.classList.remove("hidden");
            modal.classList.remove("hidden");

            const ext = String(fileName || "").split('.').pop().toLowerCase();
            const isExcel = ['xls', 'xlsx', 'xlsm', 'xlsb', 'csv'].includes(ext);
            const isDocx = ext === 'docx';
            const isPptx = ext === 'pptx';
            // Word & PowerPoint are converted to PDF on the server (Apps Script) so they open in the same PDF viewer as PDFs
            const isOfficeDoc = ['doc', 'docx', 'ppt', 'pptx'].includes(ext);

            // Loader shows a generic animated "Loading..." text (see .loading-dots CSS)
            loaderText.innerHTML = 'Loading<span class="loading-dots"></span>';

            try {
                let byteArray = null;
                let mimeType = "";
                let serverPdfError = "";

                if (isOfficeDoc && pdfPreviewBlobCache[fileId] && pdfPreviewBlobCache[fileId].byteArray) {
                    byteArray = pdfPreviewBlobCache[fileId].byteArray;
                    mimeType = "application/pdf";
                } else if (!isOfficeDoc && directDocBlobCache[fileId] && directDocBlobCache[fileId].byteArray) {
                    byteArray = directDocBlobCache[fileId].byteArray;
                    mimeType = directDocBlobCache[fileId].mimeType;
                } else {
                    const resp = await fetch(APPS_SCRIPT_URL, {
                        method: 'POST',
                        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
                        body: JSON.stringify(withSessionToken({
                            action: "get_file_base64",
                            fileId: fileId,
                            asPdf: isOfficeDoc
                        }))
                    });

                    const data = await resp.json();
                    if (handleUnauthorizedResponse(data)) {
                        throw new Error(data.message || "Session expired.");
                    }
                    if (!data || data.status !== "success" || !data.base64) {
                        throw new Error(data && data.message ? data.message : "File data was not received from server.");
                    }
                    if (data.pdfConversionError) {
                        serverPdfError = data.pdfConversionError;
                        console.warn("Server PDF conversion failed:", serverPdfError);
                    }

                    byteArray = base64ToUint8Array(data.base64);
                    mimeType = data.mimeType || inferMimeTypeFromFileName(fileName);

                    // Detect if bytes start with %PDF- header
                    const isPdfBytes = byteArray.length > 4 &&
                        byteArray[0] === 0x25 && byteArray[1] === 0x50 &&
                        byteArray[2] === 0x44 && byteArray[3] === 0x46;

                    if (isOfficeDoc && (mimeType === "application/pdf" || isPdfBytes)) {
                        // Converted PDF is only used for preview; Download still fetches the original file
                        mimeType = "application/pdf";
                        pdfPreviewBlobCache[fileId] = {
                            byteArray,
                            mimeType: "application/pdf"
                        };
                    } else {
                        // Original file (non-Office, or Office file whose PDF conversion failed)
                        const rawMime = inferMimeTypeFromFileName(fileName, data.mimeType);
                        mimeType = rawMime;
                        directDocBlobCache[fileId] = {
                            blob: new Blob([byteArray], { type: rawMime }),
                            fileName: fileName || data.fileName || "document",
                            mimeType: rawMime,
                            byteArray
                        };
                    }
                }

                if (isExcel) {
                    await loadExcelLibs();
                    const xlsxLib = window.XLSX_CORE || window.XLSX;
                    if (!xlsxLib) throw new Error("Excel reader library (SheetJS) is not ready.");
                    const wb = xlsxLib.read(byteArray, { type: 'array' });
                    currentInlineExcelWorkbook = wb;
                    renderInlineExcelWorkbook(wb, 0);
                    loader.classList.add("hidden");
                    return;
                }

                // Fallback (only when server PDF conversion failed): render DOCX inside isolated <iframe sandbox>
                if (isDocx && mimeType !== "application/pdf") {
                    if (!window.docx || typeof window.docx.renderAsync !== "function") {
                        throw new Error("Word preview library (docx-preview) is not loaded yet.");
                    }
                    const detachedHost = document.createElement("div");
                    const detachedStyleHost = document.createElement("div");
                    await window.docx.renderAsync(byteArray, detachedHost, detachedStyleHost, {
                        inWrapper: true,
                        ignoreWidth: false,
                        ignoreHeight: false
                    });
                    detachedHost.querySelectorAll("script, iframe, object, embed, base, meta").forEach(el => el.remove());
                    detachedHost.querySelectorAll("*").forEach(el => {
                        Array.from(el.attributes).forEach(attr => {
                            if (/^on/i.test(attr.name) || (attr.name === "href" && /^\s*javascript:/i.test(attr.value))) {
                                el.removeAttribute(attr.name);
                            }
                        });
                    });
                    const sandboxIframe = document.createElement("iframe");
                    sandboxIframe.setAttribute("sandbox", "");
                    sandboxIframe.setAttribute("title", "Sandboxed Word Document Preview");
                    sandboxIframe.className = "w-full h-full border-0 rounded-xl bg-white";
                    const sandboxCsp = `<meta http-equiv="Content-Security-Policy" content="default-src 'none'; style-src 'unsafe-inline'; img-src data: blob:; font-src data:;">`;
                    sandboxIframe.srcdoc = `<!DOCTYPE html><html><head><meta charset="utf-8">${sandboxCsp}<style>body{margin:0;padding:12px;background:#e5e7eb;font-family:sans-serif;}\x3C/style>${detachedStyleHost.innerHTML}\x3C/head><body>${detachedHost.innerHTML}\x3C/body>\x3C/html>`;
                    content.replaceChildren(sandboxIframe);
                    loader.classList.add("hidden");
                    return;
                }

                // Fallback (only when server PDF conversion failed): render PPTX slides directly in the browser
                if (isPptx && mimeType !== "application/pdf") {
                    await loadPptxViewerLib();
                    const nativeFileReader = window.FileReader;
                    content.innerHTML = `<div id="inlinePptxRenderContainer" class="w-full h-full overflow-auto bg-gray-200 dark:bg-neutral-900 p-2 sm:p-4 rounded-xl flex flex-col items-center gap-3 sm:gap-4"></div>`;
                    const pptxContainer = document.getElementById("inlinePptxRenderContainer");
                    const initCanvas = document.createElement("canvas");
                    initCanvas.width = 1280;
                    initCanvas.height = 720;
                    try {
                        const viewer = new window.PptxViewJS.PPTXViewer({ canvas: initCanvas, autoExposeGlobals: false });
                        // PptxViewJS only accepts File, ArrayBuffer, or Uint8Array (not Blob)
                        await viewer.loadFile(byteArray);
                        const totalSlides = typeof viewer.getSlideCount === "function" ? viewer.getSlideCount() : (viewer.slideCount || 1);
                        if (!totalSlides) throw new Error("No slides found in this presentation.");
                        for (let sIdx = 0; sIdx < totalSlides; sIdx++) {
                            const slideWrap = document.createElement("div");
                            slideWrap.className = "w-full max-w-4xl flex flex-col items-center gap-1";
                            const slideBadge = document.createElement("div");
                            slideBadge.className = "text-[11px] font-semibold text-gray-500 dark:text-neutral-400 self-start";
                            slideBadge.textContent = `Slide ${sIdx + 1} / ${totalSlides}`;
                            const slideCanvas = document.createElement("canvas");
                            slideCanvas.width = 1280;
                            slideCanvas.height = 720;
                            slideCanvas.className = "w-full h-auto bg-white rounded-lg shadow-md border dark:border-neutral-700";
                            slideWrap.appendChild(slideBadge);
                            slideWrap.appendChild(slideCanvas);
                            pptxContainer.appendChild(slideWrap);
                            await viewer.render(slideCanvas, { slideIndex: sIdx });
                        }
                    } finally {
                        if (window.FileReader !== nativeFileReader) window.FileReader = nativeFileReader;
                    }
                    loader.classList.add("hidden");
                    return;
                }

                if (ext === 'pdf') {
                    mimeType = "application/pdf";
                }

                const blob = new Blob([byteArray], { type: mimeType });
                currentInlineDocBlobUrl = URL.createObjectURL(blob);

                if (mimeType.startsWith("image/")) {
                    content.innerHTML = `
                        <div class="w-full h-full flex items-center justify-center p-2 sm:p-4 overflow-auto">
                            <img src="${currentInlineDocBlobUrl}" alt="${escapeHtmlAttr(fileName)}" class="max-w-full max-h-full object-contain rounded-lg shadow-md border dark:border-neutral-700">
                        </div>
                    `;
                } else if (mimeType === "application/pdf") {
                    content.innerHTML = `
                        <iframe src="${currentInlineDocBlobUrl}#toolbar=1&navpanes=0"
                                title="${escapeHtml(fileName || 'Document PDF Preview')}"
                                class="w-full h-full rounded-xl border dark:border-neutral-700 bg-white shadow-inner" 
                                style="border: none;"></iframe>
                    `;
                } else {
                    const origMime = inferMimeTypeFromFileName(fileName, mimeType);
                    directDocBlobCache[fileId] = {
                        blob: new Blob([byteArray], { type: origMime }),
                        fileName: fileName || "document",
                        mimeType: origMime,
                        byteArray
                    };
                    throw new Error(`Document format (${ext.toUpperCase()}) could not be converted to PDF automatically.${serverPdfError ? ' Server detail: ' + serverPdfError : ''} Please make sure the latest Apps Script is deployed, or click 'Download File Directly' below.`);
                }

                loader.classList.add("hidden");

            } catch (err) {
                console.error("Failed to render in-app viewer:", err);
                loader.classList.add("hidden");
                errorMsg.textContent = "Failed to load preview: " + (err.message || "Network error occurred.");
                errorBox.classList.remove("hidden");
            }
        }
    