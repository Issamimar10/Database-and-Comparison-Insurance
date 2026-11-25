// Auto-generated static CSP-compliant event dispatcher (zero 'unsafe-inline', zero eval)
function idcmPackArgs(arr) {
  const json = JSON.stringify(arr);
  return String(json ?? "[]")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function idcmUnescapeJsStr(s) {
  return String(s ?? "")
    .replace(/\\'/g, "'")
    .replace(/\\"/g, '"')
    .replace(/\\n/g, "\n")
    .replace(/\\r/g, "\r")
    .replace(/\\\\/g, "\\");
}

const IDCM_ACTIONS = {
  a1(el, event, args) { toggleTheme(event, el) },
  a2(el, event, args) { switchAuthTab('signin') },
  a3(el, event, args) { switchAuthTab('signup') },
  a4(el, event, args) { handleSignInSubmit(event) },
  a5(el, event, args) { switchAuthTab('forgot') },
  a6(el, event, args) { toggleAuthPasswordVisibility('authSignInPassword', el) },
  a7(el, event, args) { handleSignUpSubmit(event) },
  a8(el, event, args) { toggleAuthPasswordVisibility('authSignUpPassword', el) },
  a9(el, event, args) { toggleAuthPasswordVisibility('authSignUpConfirm', el) },
  a10(el, event, args) { toggleAuthPasswordVisibility('authSignUpInviteCode', el) },
  a11(el, event, args) { handleForgotRequestOtp(event) },
  a12(el, event, args) { handleForgotResetPassword(event) },
  a13(el, event, args) { toggleAuthPasswordVisibility('authForgotNewPassword', el) },
  a14(el, event, args) { toggleAuthPasswordVisibility('authForgotConfirmPassword', el) },
  a15(el, event, args) { handleLogout() },
  a16(el, event, args) { fetchAllSheetData(true) },
  a17(el, event, args) { openUploadModal() },
  a18(el, event, args) { openBulkDocUploadModal() },
  a19(el, event, args) { openAdminUsersModal() },
  a20(el, event, args) { scheduleFilterData() },
  a21(el, event, args) { toggleMobileFilterPanel(event) },
  a22(el, event, args) { toggleDropdown('categoryDropdown', event) },
  a23(el, event, args) { toggleDropdown('companyDropdown', event) },
  a24(el, event, args) { toggleDropdown('typeDropdown', event) },
  a25(el, event, args) { toggleDropdown('channelDropdown', event) },
  a26(el, event, args) { toggleDropdown('currencyDropdown', event) },
  a27(el, event, args) { toggleDropdown('yearDropdown', event) },
  a28(el, event, args) { toggleDropdown('underwritingDropdown', event) },
  a29(el, event, args) { toggleAddFilterMenu(event) },
  a30(el, event, args) { clearAllFilters() },
  a31(el, event, args) { toggleDropdown('sortDropdown', event) },
  a32(el, event, args) { setSortOption('default', event) },
  a33(el, event, args) { setSortOption('launch_desc', event) },
  a34(el, event, args) { setSortOption('launch_asc', event) },
  a35(el, event, args) { setSortOption('name_asc', event) },
  a36(el, event, args) { setSortOption('company_asc', event) },
  a37(el, event, args) { selectAllVisible() },
  a38(el, event, args) { openComparisonModal() },
  a39(el, event, args) { openBulkDownloadModal() },
  a40(el, event, args) { openDeleteConfirmModal(selectedProductIds) },
  a41(el, event, args) { clearSelection() },
  a42(el, event, args) { closeUploadModal() },
  a43(el, event, args) { switchAddModalTab('excel') },
  a44(el, event, args) { switchAddModalTab('manual') },
  a45(el, event, args) { onUploadTargetSheetChange() },
  a46(el, event, args) { openSheetConfigModal('new') },
  a47(el, event, args) { openSheetConfigModal('edit', document.getElementById('uploadTargetSheet')?.value) },
  a48(el, event, args) { openAliasManagerModal() },
  a49(el, event, args) { previewExcelFile() },
  a50(el, event, args) { processAndUploadExcel() },
  a51(el, event, args) { saveManualProduct() },
  a52(el, event, args) { closeSheetConfigModal() },
  a53(el, event, args) { saveSheetCategoryConfig() },
  a54(el, event, args) { closeAliasManagerModal() },
  a55(el, event, args) { setAliasManagerFilter('all') },
  a56(el, event, args) { setAliasManagerFilter('benefit') },
  a57(el, event, args) { setAliasManagerFilter('non-benefit') },
  a58(el, event, args) { renderAliasManagerList() },
  a59(el, event, args) { openAliasEditorForm(null) },
  a60(el, event, args) { closeAliasEditorForm() },
  a61(el, event, args) { saveAliasManagerEntry() },
  a62(el, event, args) { closeAdminUsersModal() },
  a63(el, event, args) { cancelAdminRoleConfirm() },
  a64(el, event, args) { submitAdminRoleConfirm(event) },
  a65(el, event, args) { toggleAuthPasswordVisibility('adminRoleConfirmPassword', el) },
  a66(el, event, args) { closeDeleteConfirmModal() },
  a67(el, event, args) { executeDeleteProducts() },
  a68(el, event, args) { closeCategoryMismatchModal() },
  a69(el, event, args) { proceedCrossCategoryComparison() },
  a70(el, event, args) { downloadComparisonExcel(el) },
  a71(el, event, args) { adjustCompareTableZoom(-0.1) },
  a72(el, event, args) { handleCompareTableZoomInput(el) },
  a73(el, event, args) { if(event.key==='Enter'){el.blur(); event.preventDefault();} },
  a74(el, event, args) { resetCompareTableZoom() },
  a75(el, event, args) { adjustCompareTableZoom(0.1) },
  a76(el, event, args) { onComparisonFreezeChange() },
  a77(el, event, args) { toggleEnlargeCompareModal() },
  a78(el, event, args) { closeComparisonModal() },
  a79(el, event, args) { if(currentActiveProduct) openDeleteConfirmModal([currentActiveProduct.id]) },
  a80(el, event, args) { toggleEditMode() },
  a81(el, event, args) { closeDetailModal() },
  a82(el, event, args) { cancelEditMode() },
  a83(el, event, args) { saveProductEdit() },
  a84(el, event, args) { closeImageLightbox(event) },
  a85(el, event, args) { event.stopPropagation() },
  a86(el, event, args) { copyCurrentLightboxImage(el) },
  a87(el, event, args) { adjustLightboxZoom(-0.25, event) },
  a88(el, event, args) { handleLightboxZoomInput(el) },
  a89(el, event, args) { resetLightboxZoom(event) },
  a90(el, event, args) { adjustLightboxZoom(0.25, event) },
  a91(el, event, args) { toggleEnlargeLightboxModal(event) },
  a92(el, event, args) { closeImageLightbox() },
  a93(el, event, args) { closeBulkDownloadModal() },
  a94(el, event, args) { executeBulkDownloadZip() },
  a95(el, event, args) { closeBulkDocUploadModal() },
  a96(el, event, args) { document.getElementById('bulkDocFileInput').click() },
  a97(el, event, args) { handleBulkDocDragOver(event) },
  a98(el, event, args) { handleBulkDocDragLeave(event) },
  a99(el, event, args) { handleBulkDocDrop(event) },
  a100(el, event, args) { handleBulkDocFileSelect(el) },
  a101(el, event, args) { startBulkDocUploadToDrive() },
  a102(el, event, args) { handleInlineDocViewerDownload(el) },
  a103(el, event, args) { closeInlineDocViewer() },
  a104(el, event, args) { closeSingleDocUploadModal() },
  a105(el, event, args) { handleSingleDocFileChange(el) },
  a106(el, event, args) { submitSingleDocUpload() },
  a107(el, event, args) { updateStagingProdName(args[0], el.value) },
  a108(el, event, args) { switchInlineExcelSheet(args[0]) },
  a109(el, event, args) { adjustInlineExcelZoom(-0.1) },
  a110(el, event, args) { adjustInlineExcelZoom(0.1) },
  a111(el, event, args) { toggleCustomFilterDropdown(args[0], event) },
  a112(el, event, args) { removeAllCustomFilters(event) },
  a113(el, event, args) { addFilterSearchQuery = el.value; renderAddFilterMenu(); },
  a114(el, event, args) { toggleCustomFilterColumn(args[0], el.checked, event) },
  a115(el, event, args) { openDetailModal(args[0]) },
  a116(el, event, args) { toggleSelectProduct(args[0], event) },
  a117(el, event, args) { removeSheetConfigSpecialRow(args[0]) },
  a118(el, event, args) { removeSheetConfigSectionBlock(args[0]) },
  a119(el, event, args) { openSheetConfigModal('edit', el.value) },
  a120(el, event, args) { addQuickGeneralProvisionInModal() },
  a121(el, event, args) { addSheetConfigSpecialRow() },
  a122(el, event, args) { addSheetConfigSectionBlock() },
  a123(el, event, args) { handlePasteImageFromClipboardBtn(el) },
  a124(el, event, args) { handleFieldImageUpload(el) },
  a125(el, event, args) { removeSectionItemRow(el, args[0]) },
  a126(el, event, args) { validateSectionItemName(el) },
  a127(el, event, args) { handleFieldImagePaste(event, el) },
  a128(el, event, args) { clearSpecFieldInput(el) },
  a129(el, event, args) { toggleAllDplkBenefitCheckboxes(args[0], true) },
  a130(el, event, args) { toggleAllDplkBenefitCheckboxes(args[0], false) },
  a131(el, event, args) { onDplkBenefitCheckboxChange(args[0]) },
  a132(el, event, args) { el.closest('.dplk-custom-benefit-item').remove(); onDplkBenefitCheckboxChange(args[0]); },
  a133(el, event, args) { addDplkCustomBenefitRow(args[0]) },
  a134(el, event, args) { addNewChargeRow(args[0], args[1]) },
  a135(el, event, args) { removeChargeRow(el) },
  a136(el, event, args) { addNewBenefitRow(args[0], args[1]) },
  a137(el, event, args) { removeBenefitRow(el) },
  a138(el, event, args) { addNewSectionItemRow(args[0], args[1], args[2], args[3]) },
  a139(el, event, args) { addNewNonBenefitRow(args[0]) },
  a140(el, event, args) { removeNonBenefitRow(el) },
  a141(el, event, args) { openSingleDocUploadModal(args[0]) },
  a142(el, event, args) { previewDocumentInline(args[0], args[1], '', args[2]) },
  a143(el, event, args) { downloadDocumentDirect(args[0], args[1], el, args[2]) },
  a144(el, event, args) { deleteSingleProductDoc(args[0], args[1]) },
  a145(el, event, args) { validateEditProductIdentity() },
  a146(el, event, args) { openSheetConfigModal('edit', document.getElementById('edit_category')?.value) },
  a147(el, event, args) { validateNonBenefitName(el) },
  a148(el, event, args) { el.closest('.attached-img-item').remove() },
  a149(el, event, args) { validateBenefitName(el) },
  a150(el, event, args) { validateChargeName(el) },
  a151(el, event, args) { validateManualCategoryAndToggleSyariah() },
  a152(el, event, args) { openSheetConfigModal('edit', document.getElementById('manual_category')?.value) },
  a153(el, event, args) { validateManualProductName(el) },
  a154(el, event, args) { onUploadSheetToggle(args[0], el.checked) },
  a155(el, event, args) { onUploadSheetTargetChange(args[0], el.value) },
  a156(el, event, args) { onUploadSheetToggleAll(el.checked) },
  a157(el, event, args) { uploadSubmitBaseEnabled = el.checked; refreshUploadSubmitState(); },
  a158(el, event, args) { onBenefitReviewPickSuggestion(args[0], args[1]) },
  a159(el, event, args) { onBenefitReviewDescProductChange(args[0], el.value) },
  a160(el, event, args) { onBenefitReviewKindChange(args[0], args[1]) },
  a161(el, event, args) { onBenefitReviewModeChange(args[0], el.checked ? 'new' : 'pending') },
  a162(el, event, args) { onBenefitReviewModeChange(args[0], args[1]) },
  a163(el, event, args) { onBenefitReviewNameInput(args[0], el.value) },
  a164(el, event, args) { onBenefitReviewDescReset(args[0]) },
  a165(el, event, args) { onBenefitReviewDescInput(args[0], el.value) },
  a166(el, event, args) { toggleBenefitReviewPendingOnly() },
  a167(el, event, args) { bulkSetBenefitReviewMode('new') },
  a168(el, event, args) { bulkSetBenefitReviewMode('ignore') },
  a169(el, event, args) { bulkSetBenefitReviewMode('pending') },
  a170(el, event, args) { openAliasEditorForm(args[0]) },
  a171(el, event, args) { deleteAliasManagerEntry(args[0]) },
  a172(el, event, args) { requestAdminUserRoleChange(args[0], el.value) },
  a173(el, event, args) { deleteAdminUser(args[0]) },
  a174(el, event, args) { goToDeckPage(args[0]) },
  a175(el, event, args) { setDeckPageSize(el.value) },
  a176(el, event, args) { goToComparePage(args[0]) },
  a177(el, event, args) { setComparePageSize(el.value) },
  a178(el, event, args) { toggleDesktopSecondaryFilters(event) },
  a179(el, event, args) { toggleUiAnimations() }
};

(function initIdcmEventDelegation() {
  const DELEGATED_EVENTS = ["click", "change", "input", "submit", "paste", "keydown", "keyup", "focus", "blur", "dragover", "dragleave", "drop"];
  DELEGATED_EVENTS.forEach((evType) => {
    const attrName = "data-ev-" + evType;
    const argsAttrName = "data-ev-" + evType + "-args";
    const useCapture = evType === "focus" || evType === "blur";
    document.addEventListener(evType, (event) => {
      let node = event.target instanceof Element ? event.target : (event.target && event.target.parentElement);
      while (node && node !== document) {
        if (node.hasAttribute(attrName)) {
          const actionId = node.getAttribute(attrName);
          const fn = IDCM_ACTIONS[actionId];
          if (typeof fn === "function") {
            let parsedArgs = [];
            const rawArgs = node.getAttribute(argsAttrName);
            if (rawArgs) {
              try { parsedArgs = JSON.parse(rawArgs); } catch (e) { parsedArgs = []; }
            }
            const ret = fn.call(node, node, event, parsedArgs);
            if (ret === false) event.preventDefault();
            if (event.cancelBubble) break;
          }
        }
        node = node.parentElement;
      }
    }, useCapture);
  });
})();
