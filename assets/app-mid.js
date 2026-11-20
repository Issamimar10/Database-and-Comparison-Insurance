let pptxViewerLibPromise = null;
        let excelLibsPromise = null;
        function loadScriptOnce(src, integrityHash) {
            return new Promise((resolve, reject) => {
                const s = document.createElement("script");
                s.src = src;
                if (integrityHash) {
                    s.integrity = integrityHash;
                    s.crossOrigin = "anonymous";
                }
                s.onload = () => resolve();
                s.onerror = () => reject(new Error("Failed to load script: " + src));
                document.head.appendChild(s);
            });
        }
        function loadExcelLibs() {
            if (window.XLSX_CORE && window.XLSX) return Promise.resolve();
            if (excelLibsPromise) return excelLibsPromise;
            excelLibsPromise = (async () => {
                try {
                    if (!window.XLSX_CORE) {
                        await loadScriptOnce(
                            "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
                            "sha384-EnyY0/GSHQGSxSgMwaIPzSESbqoOLSexfnSMN2AP+39Ckmn92stwABZynq1JyzdT"
                        );
                        window.XLSX_CORE = window.XLSX;
                    }
                    await loadScriptOnce(
                        "https://cdn.jsdelivr.net/npm/xlsx-js-style@1.2.0/dist/xlsx.bundle.js",
                        "sha384-OUW9euuUyxyHcAhTqbhI+Iyb8LMssXt/cpz0yXhs9UWG2/R/uaWdakx/4cfww7Vb"
                    );
                } catch (err) {
                    excelLibsPromise = null;
                    throw err;
                }
            })();
            return excelLibsPromise;
        }
        function loadPptxViewerLib() {
            if (window.PptxViewJS && typeof window.PptxViewJS.PPTXViewer === "function") return Promise.resolve();
            if (pptxViewerLibPromise) return pptxViewerLibPromise;
            pptxViewerLibPromise = (async () => {
                // Snapshot existing globals so the library cannot clobber native APIs (FileReader, etc.)
                const protectedNames = Object.getOwnPropertyNames(window);
                const snapshot = {};
                protectedNames.forEach(name => {
                    try { snapshot[name] = window[name]; } catch (e) { }
                });
                try {
                    if (!window.Chart) {
                        await loadScriptOnce("https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js", "sha384-9nhczxUqK87bcKHh20fSQcTGD4qq5GhayNYSYWqwBkINBhOfQLg/P5HG5lF1urn4");
                        snapshot.Chart = window.Chart;
                    }
                    await loadScriptOnce("https://cdn.jsdelivr.net/npm/pptxviewjs@1.1.9/dist/PptxViewJS.min.js", "sha384-bmEMthgLe+RjctA6cP56ldWv/0ZwdbYX3CNn/SeEv0XhFUk0K/NnMbhqTYVOpLtL");
                } finally {
                    Object.keys(snapshot).forEach(name => {
                        try {
                            if (window[name] !== snapshot[name]) window[name] = snapshot[name];
                        } catch (e) { }
                    });
                }
                if (!window.PptxViewJS || typeof window.PptxViewJS.PPTXViewer !== "function") {
                    pptxViewerLibPromise = null;
                    throw new Error("PowerPoint preview library (PptxViewJS) failed to initialize.");
                }
            })();
            pptxViewerLibPromise.catch(() => { pptxViewerLibPromise = null; });
            return pptxViewerLibPromise;
        }
    