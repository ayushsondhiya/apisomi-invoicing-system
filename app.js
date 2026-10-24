/**
 * Apisomi Invoicing & Automated Ledger Engine
 * Complete implementation of the user's Apisomi Tax Invoice & Non-GST specification.
 * Handles:
 * - Numbering convention: APS-INV-[PROJECT]-[NN] (e.g. APS-INV-SSV003-01)
 * - PDF file naming convention: Apisomi_[Project]_Invoice_[NN].pdf
 * - Auto tax determination: MP (23) -> CGST 9% + SGST 9%; Outside MP -> IGST 18%; Non-GST -> 0%
 * - SAC code (default 998314) & Unit columns (Lot, Hours, etc.)
 * - Discount, Taxable Value, Round-off, Grand Total, Advance Received, and Balance Due
 * - Payment Schedule & Installments with reserved Txn / UTR references
 * - Indian number-to-words converter (Rupees ... Lakhs ... Only)
 * - Automatic cumulative Excel sheet generator (.xlsx)
 */

// Application State
const state = {
  isGstMode: false,
  gstRate: 18,
  projectCode: 'SSV003',
  invSeq: '01',
  discountPct: 0,
  advanceReceived: 0,
  autoRound: true,
  items: [
    { id: 1, desc: 'UI/UX design — 12 screens & design system prototype', sac: '998314', qty: 1, unit: 'Lot', rate: 45000 },
    { id: 2, desc: 'Android & iOS telemetry app development — Sprint 1', sac: '998314', qty: 40, unit: 'Hours', rate: 1200 },
    { id: 3, desc: 'Cloud API, drone telemetry & automated database deployment', sac: '998314', qty: 1, unit: 'Lot', rate: 25000 }
  ],
  installments: [
    { id: 1, label: 'Milestone 1 — Architecture & Wireframes (50%)', dueDate: '2026-10-15', amount: 59000, txn: 'UTR-HDFC-9928192' },
    { id: 2, label: 'Milestone 2 — Design Sign-off & Sprint 1 (30%)', dueDate: '2026-10-25', amount: 35400, txn: '' },
    { id: 3, label: 'Milestone 3 — Production Handover (20%)', dueDate: '2026-11-15', amount: 23600, txn: '' }
  ],
  ledger: [],
  gsheetWebhookUrl: localStorage.getItem('apisomi_gsheet_webhook') || '',
  gsheetDirectUrl: localStorage.getItem('apisomi_gsheet_direct') || 'https://docs.google.com/spreadsheets/d/1fubQ-v4GZixmfQOOV2VnrJtoUDNjsSM-wdmwwXwkrQk/edit?usp=sharing'
};

const STORAGE_KEY = 'apisomi_invoices_ledger_v2';
const GSCRIPT_TEMPLATE = `function doPost(e) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = JSON.parse(e.postData.contents);
  
  // Set headers on first row if empty
  if (sheet.getLastRow() === 0) {
    var headers = [
      "Invoice No", "Date", "Due Date", "Project Code", "Project Name", 
      "Milestone", "Client Name", "Client State", "Client GSTIN", "Tax Regime", 
      "Subtotal (₹)", "Discount (₹)", "Taxable Value (₹)", "CGST (₹)", "SGST (₹)", 
      "IGST (₹)", "Round-off (₹)", "Grand Total (₹)", "Advance Paid (₹)", "Balance Due (₹)", 
      "Installments", "Logged Timestamp"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f1f5f9");
    sheet.setFrozenRows(1);
  }
  
  // Append new invoice row
  sheet.appendRow([
    data.invoiceNo, data.invoiceDate, data.dueDate, data.projectCode, data.projectName,
    data.milestone, data.clientName, data.clientState, data.clientGstin, data.taxRegime,
    data.subtotal, data.discount, data.taxableValue, data.cgst, data.sgst,
    data.igst, data.roundOff, data.grandTotal, data.advanceReceived, data.balanceDue,
    data.installmentsCount + " Tranches", data.loggedAt
  ]);
  
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Row added" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({ status: "success", message: "Apisomi Invoicing Webhook is Active!" }))
    .setMimeType(ContentService.MimeType.JSON);
}`;

// DOM Elements
const el = {
  taxModeToggle: document.getElementById('taxModeToggle'),
  taxModeBadge: document.getElementById('taxModeBadge'),
  projectCode: document.getElementById('projectCode'),
  invSeq: document.getElementById('invSeq'),
  invNumber: document.getElementById('invNumber'),
  invDate: document.getElementById('invDate'),
  invDueDate: document.getElementById('invDueDate'),
  docVersion: document.getElementById('docVersion'),
  proposalRef: document.getElementById('proposalRef'),
  milestoneDesc: document.getElementById('milestoneDesc'),

  // Google Sheets UI
  googleSheetBtn: document.getElementById('googleSheetBtn'),
  syncNowBtn: document.getElementById('syncNowBtn'),
  gsheetBtnText: document.getElementById('gsheetBtnText'),
  gsheetStatusDot: document.getElementById('gsheetStatusDot'),
  gsheetModal: document.getElementById('gsheetModal'),
  closeGsheetModalBtn: document.getElementById('closeGsheetModalBtn'),
  syncStatusCard: document.getElementById('syncStatusCard'),
  syncStatusTitle: document.getElementById('syncStatusTitle'),
  syncStatusSubtitle: document.getElementById('syncStatusSubtitle'),
  gsheetWebhookUrl: document.getElementById('gsheetWebhookUrl'),
  gsheetDirectUrl: document.getElementById('gsheetDirectUrl'),
  copyScriptBtn: document.getElementById('copyScriptBtn'),
  testGsheetSyncBtn: document.getElementById('testGsheetSyncBtn'),
  openSheetLink: document.getElementById('openSheetLink'),
  saveGsheetConfigBtn: document.getElementById('saveGsheetConfigBtn'),

  // Client
  clientName: document.getElementById('clientName'),
  clientContact: document.getElementById('clientContact'),
  clientAddress: document.getElementById('clientAddress'),
  clientStateSelect: document.getElementById('clientStateSelect'),
  clientGstin: document.getElementById('clientGstin'),
  clientPan: document.getElementById('clientPan'),
  clientEmailPhone: document.getElementById('clientEmailPhone'),

  // Apisomi
  apisomiName: document.getElementById('apisomiName'),
  apisomiContact: document.getElementById('apisomiContact'),
  apisomiAddress: document.getElementById('apisomiAddress'),
  apisomiGstin: document.getElementById('apisomiGstin'),
  apisomiPan: document.getElementById('apisomiPan'),
  apisomiState: document.getElementById('apisomiState'),
  gstinStatusLabel: document.getElementById('gstinStatusLabel'),

  // Project Header & Calc
  projectName: document.getElementById('projectName'),
  projectPeriod: document.getElementById('projectPeriod'),
  discountPercent: document.getElementById('discountPercent'),
  advanceReceivedInput: document.getElementById('advanceReceivedInput'),
  autoRoundSelect: document.getElementById('autoRoundSelect'),
  gstRateRow: document.getElementById('gstRateRow'),
  gstRateSelect: document.getElementById('gstRateSelect'),
  placeOfSupplyHint: document.getElementById('placeOfSupplyHint'),

  // Bank & Notes
  bankAccName: document.getElementById('bankAccName'),
  bankBranch: document.getElementById('bankBranch'),
  bankAccNo: document.getElementById('bankAccNo'),
  bankIfsc: document.getElementById('bankIfsc'),
  bankUpi: document.getElementById('bankUpi'),
  projectSpecificNote: document.getElementById('projectSpecificNote'),

  // Dynamic tables & items
  itemsTableBody: document.getElementById('itemsTableBody'),
  installmentsTableBody: document.getElementById('installmentsTableBody'),
  addItemBtn: document.getElementById('addItemBtn'),
  addInstallmentBtn: document.getElementById('addInstallmentBtn'),

  // Preview elements
  prevDocHeading: document.getElementById('prevDocHeading'),
  previewDocType: document.getElementById('previewDocType'),
  fileNamePreview: document.getElementById('fileNamePreview'),
  prevInvNum: document.getElementById('prevInvNum'),
  prevVersion: document.getElementById('prevVersion'),
  prevInvDate: document.getElementById('prevInvDate'),
  prevInvDueDate: document.getElementById('prevInvDueDate'),
  prevProposalRef: document.getElementById('prevProposalRef'),
  prevMilestone: document.getElementById('prevMilestone'),

  prevPartiesBilledByName: document.getElementById('prevPartiesBilledByName'),
  prevPartiesBilledByContact: document.getElementById('prevPartiesBilledByContact'),
  prevPartiesBilledByAddr: document.getElementById('prevPartiesBilledByAddr'),
  prevPartiesBilledByGst: document.getElementById('prevPartiesBilledByGst'),
  prevPartiesBilledByPan: document.getElementById('prevPartiesBilledByPan'),
  prevPartiesBilledByState: document.getElementById('prevPartiesBilledByState'),

  prevPartiesClientName: document.getElementById('prevPartiesClientName'),
  prevPartiesClientContact: document.getElementById('prevPartiesClientContact'),
  prevPartiesClientAddr: document.getElementById('prevPartiesClientAddr'),
  prevPartiesClientGst: document.getElementById('prevPartiesClientGst'),
  prevPartiesClientPan: document.getElementById('prevPartiesClientPan'),
  prevPartiesClientState: document.getElementById('prevPartiesClientState'),
  prevPartiesClientEmailPhone: document.getElementById('prevPartiesClientEmailPhone'),

  prevPlaceOfSupply: document.getElementById('prevPlaceOfSupply'),
  prevProjectName: document.getElementById('prevProjectName'),
  prevProjectCode: document.getElementById('prevProjectCode'),
  prevProjectPeriod: document.getElementById('prevProjectPeriod'),
  prevItemsTbody: document.getElementById('prevItemsTbody'),
  prevInstallmentsTbody: document.getElementById('prevInstallmentsTbody'),

  prevAmountWords: document.getElementById('prevAmountWords'),
  prevSummarySubtotal: document.getElementById('prevSummarySubtotal'),
  prevDiscountRow: document.getElementById('prevDiscountRow'),
  prevDiscountPct: document.getElementById('prevDiscountPct'),
  prevSummaryDiscount: document.getElementById('prevSummaryDiscount'),
  prevSummaryTaxable: document.getElementById('prevSummaryTaxable'),
  prevCgstRow: document.getElementById('prevCgstRow'),
  prevSummaryCgst: document.getElementById('prevSummaryCgst'),
  prevSgstRow: document.getElementById('prevSgstRow'),
  prevSummarySgst: document.getElementById('prevSummarySgst'),
  prevIgstRow: document.getElementById('prevIgstRow'),
  prevSummaryIgst: document.getElementById('prevSummaryIgst'),
  prevNonGstTaxRow: document.getElementById('prevNonGstTaxRow'),
  prevSummaryRoundoff: document.getElementById('prevSummaryRoundoff'),
  prevSummaryGrandTotal: document.getElementById('prevSummaryGrandTotal'),
  prevSummaryAdvance: document.getElementById('prevSummaryAdvance'),
  prevSummaryBalance: document.getElementById('prevSummaryBalance'),

  prevBankAccName: document.getElementById('prevBankAccName'),
  prevBankBranch: document.getElementById('prevBankBranch'),
  prevBankAccNo: document.getElementById('prevBankAccNo'),
  prevBankIfsc: document.getElementById('prevBankIfsc'),
  prevBankUpi: document.getElementById('prevBankUpi'),
  prevProjectNote: document.getElementById('prevProjectNote'),

  // Actions
  printBtn: document.getElementById('printBtn'),
  downloadPdfBtn: document.getElementById('downloadPdfBtn'),
  viewLedgerBtn: document.getElementById('viewLedgerBtn'),
  quickExcelBtn: document.getElementById('quickExcelBtn'),
  ledgerCount: document.getElementById('ledgerCount'),
  resetFormBtn: document.getElementById('resetFormBtn'),

  // Modal
  ledgerModal: document.getElementById('ledgerModal'),
  closeLedgerBtn: document.getElementById('closeLedgerBtn'),
  ledgerTbody: document.getElementById('ledgerTbody'),
  emptyLedgerMessage: document.getElementById('emptyLedgerMessage'),
  exportExcelBtn: document.getElementById('exportExcelBtn'),
  clearLedgerBtn: document.getElementById('clearLedgerBtn'),
  metricCount: document.getElementById('metricCount'),
  metricTotal: document.getElementById('metricTotal'),
  metricGstCount: document.getElementById('metricGstCount'),
  metricNonGstCount: document.getElementById('metricNonGstCount'),
  invoiceDocument: document.getElementById('invoiceDocument')
};

// Application Initialization
function init() {
  const today = new Date().toISOString().split('T')[0];
  const dueDate = new Date(Date.now() + 15 * 86400000).toISOString().split('T')[0];
  el.invDate.value = today;
  el.invDueDate.value = dueDate;

  loadLedger();
  initGoogleSheetSync();
  computeInvoiceNumber();
  renderItemsTable();
  renderInstallmentsTable();
  attachEvents();
  updateLivePreview();
}

// Compute standard document number and PDF filename
function computeInvoiceNumber() {
  const proj = (el.projectCode.value.trim() || 'PROJECT').toUpperCase();
  const seq = (el.invSeq.value.trim() || '01');
  const invNum = `APS-INV-${proj}-${seq}`;
  el.invNumber.value = invNum;

  // Proposal ref sync
  el.proposalRef.value = `APS-TCP-${proj}-${seq} · Client PO No. [N/A]`;

  // Filename format: Apisomi_[Project]_Invoice_[NN].pdf
  const pdfName = `Apisomi_${proj}_Invoice_${seq}.pdf`;
  el.fileNamePreview.textContent = pdfName;
}

// Event Listeners
function attachEvents() {
  // GST Toggle
  el.taxModeToggle.addEventListener('change', (e) => {
    state.isGstMode = e.target.checked;
    handleTaxModeToggle();
  });

  // Client State change (triggers auto CGST+SGST vs IGST)
  el.clientStateSelect.addEventListener('change', updateLivePreview);

  // Numbering changes
  el.projectCode.addEventListener('input', () => {
    computeInvoiceNumber();
    updateLivePreview();
  });
  el.invSeq.addEventListener('input', () => {
    computeInvoiceNumber();
    updateLivePreview();
  });

  // Discount, Advance, Rounding
  el.discountPercent.addEventListener('input', (e) => {
    state.discountPct = parseFloat(e.target.value) || 0;
    updateLivePreview();
  });
  el.advanceReceivedInput.addEventListener('input', (e) => {
    state.advanceReceived = parseFloat(e.target.value) || 0;
    updateLivePreview();
  });
  el.autoRoundSelect.addEventListener('change', (e) => {
    state.autoRound = (e.target.value === 'yes');
    updateLivePreview();
  });

  // General inputs sync
  const liveInputs = [
    el.invDate, el.invDueDate, el.docVersion, el.proposalRef, el.milestoneDesc,
    el.clientName, el.clientContact, el.clientAddress, el.clientGstin, el.clientPan, el.clientEmailPhone,
    el.apisomiName, el.apisomiContact, el.apisomiAddress, el.apisomiGstin, el.apisomiPan,
    el.projectName, el.projectPeriod, el.bankAccName, el.bankBranch, el.bankAccNo, el.bankIfsc, el.bankUpi,
    el.projectSpecificNote
  ];
  liveInputs.forEach(input => input.addEventListener('input', updateLivePreview));

  // Add Item
  el.addItemBtn.addEventListener('click', () => {
    state.items.push({
      id: Date.now(),
      desc: 'Technical Services / Milestone Deliverable',
      sac: '998314',
      qty: 1,
      unit: 'Lot',
      rate: 10000
    });
    renderItemsTable();
    updateLivePreview();
  });

  // Add Installment
  el.addInstallmentBtn.addEventListener('click', () => {
    state.installments.push({
      id: Date.now(),
      label: `Milestone ${state.installments.length + 1}`,
      dueDate: el.invDueDate.value,
      amount: 0,
      txn: ''
    });
    renderInstallmentsTable();
    updateLivePreview();
  });

  // Google Sheets Actions
  el.googleSheetBtn.addEventListener('click', openGsheetModal);
  el.syncNowBtn.addEventListener('click', handleManualSheetUpdate);
  el.closeGsheetModalBtn.addEventListener('click', closeGsheetModal);
  el.copyScriptBtn.addEventListener('click', copyGoogleAppsScript);
  el.saveGsheetConfigBtn.addEventListener('click', saveGoogleSheetConfig);
  el.testGsheetSyncBtn.addEventListener('click', testGoogleSheetConnection);

  el.gsheetModal.addEventListener('click', (e) => {
    if (e.target === el.gsheetModal) closeGsheetModal();
  });

  // Actions
  el.printBtn.addEventListener('click', () => window.print());
  el.downloadPdfBtn.addEventListener('click', exportPdfAndLog);
  el.viewLedgerBtn.addEventListener('click', openLedgerModal);
  el.quickExcelBtn.addEventListener('click', generateExcelWorkbook);
  el.closeLedgerBtn.addEventListener('click', closeLedgerModal);
  el.exportExcelBtn.addEventListener('click', generateExcelWorkbook);
  el.clearLedgerBtn.addEventListener('click', clearLedger);
  el.resetFormBtn.addEventListener('click', resetToDefaults);

  el.ledgerModal.addEventListener('click', (e) => {
    if (e.target === el.ledgerModal) closeLedgerModal();
  });
}

// Tax Mode Toggle
function handleTaxModeToggle() {
  if (state.isGstMode) {
    el.taxModeBadge.textContent = 'GST Invoice (Active)';
    el.taxModeBadge.classList.remove('non-gst');
    el.taxModeBadge.classList.add('gst-active');
    el.gstinStatusLabel.textContent = 'GST Active';
    el.gstinStatusLabel.style.background = 'rgba(16, 185, 129, 0.2)';
    el.gstinStatusLabel.style.color = '#34d399';
    el.gstRateRow.classList.remove('hidden');
    el.prevNonGstTaxRow.classList.add('hidden');
    el.prevDocHeading.textContent = 'APISOMI — TAX INVOICE';
    el.previewDocType.textContent = 'GST Tax Invoice';
  } else {
    el.taxModeBadge.textContent = 'Non-GST (Current)';
    el.taxModeBadge.classList.remove('gst-active');
    el.taxModeBadge.classList.add('non-gst');
    el.gstinStatusLabel.textContent = 'When Received';
    el.gstinStatusLabel.style.background = 'rgba(245, 158, 11, 0.15)';
    el.gstinStatusLabel.style.color = '#fbbf24';
    el.gstRateRow.classList.add('hidden');
    el.prevNonGstTaxRow.classList.remove('hidden');
    el.prevDocHeading.textContent = 'APISOMI — INVOICE';
    el.previewDocType.textContent = 'Non-GST Bill of Supply';
  }
  updateLivePreview();
}

// Render Line Items in Form
function renderItemsTable() {
  el.itemsTableBody.innerHTML = '';
  state.items.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" value="${escapeHtml(item.desc)}" data-field="desc" data-idx="${idx}" placeholder="Deliverable / Milestone"></td>
      <td><input type="text" value="${escapeHtml(item.sac)}" data-field="sac" data-idx="${idx}" class="mono"></td>
      <td><input type="number" min="0" step="any" value="${item.qty}" data-field="qty" data-idx="${idx}"></td>
      <td><input type="text" value="${escapeHtml(item.unit)}" data-field="unit" data-idx="${idx}" placeholder="Lot/Hrs"></td>
      <td><input type="number" min="0" step="any" value="${item.rate}" data-field="rate" data-idx="${idx}"></td>
      <td class="mono font-semibold" style="padding-top:10px;">₹ ${(item.qty * item.rate).toLocaleString('en-IN')}</td>
      <td>
        <button type="button" class="btn-del-row" data-action="del-item" data-idx="${idx}" title="Delete item">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
      </td>
    `;
    el.itemsTableBody.appendChild(tr);
  });

  // Input listeners
  el.itemsTableBody.querySelectorAll('input').forEach(input => {
    input.addEventListener('input', (e) => {
      const idx = e.target.dataset.idx;
      const f = e.target.dataset.field;
      if (f === 'desc') state.items[idx].desc = e.target.value;
      if (f === 'sac') state.items[idx].sac = e.target.value;
      if (f === 'qty') state.items[idx].qty = parseFloat(e.target.value) || 0;
      if (f === 'unit') state.items[idx].unit = e.target.value;
      if (f === 'rate') state.items[idx].rate = parseFloat(e.target.value) || 0;

      // Update total cell
      const row = e.target.closest('tr');
      row.children[5].textContent = `₹ ${(state.items[idx].qty * state.items[idx].rate).toLocaleString('en-IN')}`;
      updateLivePreview();
    });
  });

  // Delete handlers
  el.itemsTableBody.querySelectorAll('[data-action="del-item"]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      if (state.items.length <= 1) {
        alert('An invoice must contain at least 1 line item.');
        return;
      }
      const idx = parseInt(e.currentTarget.dataset.idx);
      state.items.splice(idx, 1);
      renderItemsTable();
      updateLivePreview();
    });
  });
}

// Render Installments in Form
function renderInstallmentsTable() {
  el.installmentsTableBody.innerHTML = '';
  state.installments.forEach((inst, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" value="${escapeHtml(inst.label)}" data-field="label" data-idx="${idx}" placeholder="e.g. Milestone 1"></td>
      <td><input type="date" value="${inst.dueDate || ''}" data-field="dueDate" data-idx="${idx}"></td>
      <td><input type="number" min="0" step="any" value="${inst.amount}" data-field="amount" data-idx="${idx}"></td>
      <td><input type="text" value="${escapeHtml(inst.txn || '')}" data-field="txn" data-idx="${idx}" placeholder="Empty or Txn/UTR #"></td>
      <td>
        <button type="button" class="btn-del-row" data-action="del-inst" data-idx="${idx}">
          <svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
      </td>
    `;
    el.installmentsTableBody.appendChild(tr);
  });

  el.installmentsTableBody.querySelectorAll('input').forEach(input => {
    input.addEventListener('input', (e) => {
      const idx = e.target.dataset.idx;
      const f = e.target.dataset.field;
      if (f === 'label') state.installments[idx].label = e.target.value;
      if (f === 'dueDate') state.installments[idx].dueDate = e.target.value;
      if (f === 'amount') state.installments[idx].amount = parseFloat(e.target.value) || 0;
      if (f === 'txn') state.installments[idx].txn = e.target.value;
      updateLivePreview();
    });
  });

  el.installmentsTableBody.querySelectorAll('[data-action="del-inst"]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt(e.currentTarget.dataset.idx);
      state.installments.splice(idx, 1);
      renderInstallmentsTable();
      updateLivePreview();
    });
  });
}

// Financial Computations
function calculateFinancials() {
  const subtotal = state.items.reduce((acc, it) => acc + (it.qty * it.rate), 0);
  const discountAmount = (subtotal * state.discountPct) / 100;
  const taxableValue = Math.max(0, subtotal - discountAmount);

  // Parse state code
  const [clientCode, clientStateName] = el.clientStateSelect.value.split('|');
  const isIntraState = (clientCode === '23'); // Madhya Pradesh

  let cgst = 0;
  let sgst = 0;
  let igst = 0;
  let totalTax = 0;

  if (state.isGstMode) {
    if (isIntraState) {
      cgst = (taxableValue * 9) / 100;
      sgst = (taxableValue * 9) / 100;
      totalTax = cgst + sgst;
    } else {
      igst = (taxableValue * 18) / 100;
      totalTax = igst;
    }
  }

  const rawGrandTotal = taxableValue + totalTax;
  let roundOff = 0;
  let grandTotal = rawGrandTotal;

  if (state.autoRound) {
    grandTotal = Math.round(rawGrandTotal);
    roundOff = grandTotal - rawGrandTotal;
  }

  const balanceDue = Math.max(0, grandTotal - state.advanceReceived);

  return {
    subtotal,
    discountAmount,
    taxableValue,
    isIntraState,
    clientCode,
    clientStateName,
    cgst,
    sgst,
    igst,
    totalTax,
    roundOff,
    grandTotal,
    advanceReceived: state.advanceReceived,
    balanceDue
  };
}

// Update Live Preview Paper
function updateLivePreview() {
  const calc = calculateFinancials();

  // Document metadata
  el.prevInvNum.textContent = el.invNumber.value || 'APS-INV-PROJECT-01';
  el.prevVersion.textContent = el.docVersion.value || 'v1.0 · Issued';
  el.prevInvDate.textContent = formatApisomiDate(el.invDate.value);
  el.prevInvDueDate.textContent = `${formatApisomiDate(el.invDueDate.value)} (Net 15)`;
  el.prevProposalRef.textContent = el.proposalRef.value || '-';
  el.prevMilestone.textContent = el.milestoneDesc.value || '-';

  // Parties: Apisomi
  el.prevPartiesBilledByName.textContent = el.apisomiName.value;
  el.prevPartiesBilledByContact.textContent = el.apisomiContact.value;
  el.prevPartiesBilledByAddr.textContent = el.apisomiAddress.value;
  el.prevPartiesBilledByGst.textContent = state.isGstMode ? (el.apisomiGstin.value.trim() || 'Applied for / GST Active') : 'Applied for / Non-GST';
  el.prevPartiesBilledByPan.textContent = el.apisomiPan.value.trim() || '-';
  el.prevPartiesBilledByState.textContent = el.apisomiState.value;

  // Parties: Client
  el.prevPartiesClientName.textContent = el.clientName.value || 'Client Name';
  el.prevPartiesClientContact.textContent = el.clientContact.value || '-';
  el.prevPartiesClientAddr.textContent = el.clientAddress.value || '-';
  el.prevPartiesClientGst.textContent = el.clientGstin.value.trim() || 'Unregistered';
  el.prevPartiesClientPan.textContent = el.clientPan.value.trim() || '-';
  el.prevPartiesClientState.textContent = `${calc.clientStateName} (${calc.clientCode})`;
  el.prevPartiesClientEmailPhone.textContent = el.clientEmailPhone.value || '-';

  // Place of supply banner
  if (state.isGstMode) {
    if (calc.isIntraState) {
      el.prevPlaceOfSupply.innerHTML = `<strong>Place of Supply:</strong> Madhya Pradesh (23) — same state as Apisomi &rarr; CGST 9% + SGST 9%.`;
      el.placeOfSupplyHint.textContent = `Client in MP (23) → CGST (9%) + SGST (9%)`;
    } else {
      el.prevPlaceOfSupply.innerHTML = `<strong>Place of Supply:</strong> ${calc.clientStateName} (${calc.clientCode}) — other state &rarr; IGST 18%.`;
      el.placeOfSupplyHint.textContent = `Client in ${calc.clientStateName} (${calc.clientCode}) → IGST (18%)`;
    }
  } else {
    el.prevPlaceOfSupply.innerHTML = `<strong>Place of Supply:</strong> ${calc.clientStateName} (${calc.clientCode}) — Non-GST supply under threshold.`;
    el.placeOfSupplyHint.textContent = `Non-GST Mode: Tax rate 0%`;
  }

  // Project Header
  el.prevProjectName.textContent = el.projectName.value || 'Project Name';
  el.prevProjectCode.textContent = (el.projectCode.value.trim() || 'XXX000').toUpperCase();
  el.prevProjectPeriod.textContent = el.projectPeriod.value || '-';

  // Render Items in Preview
  el.prevItemsTbody.innerHTML = '';
  state.items.forEach((item, idx) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="color:#64748b; font-weight:600;">${idx + 1}</td>
      <td style="font-weight:600; color:#0f172a;">${escapeHtml(item.desc)}</td>
      <td class="mono-cell" style="color:#64748b;">${escapeHtml(item.sac)}</td>
      <td style="text-align: center;">${item.qty}</td>
      <td style="text-align: center; color:#64748b;">${escapeHtml(item.unit)}</td>
      <td style="text-align: right;" class="mono-cell">${formatINR(item.rate)}</td>
      <td style="text-align: right; font-weight:600;" class="mono-cell">${formatINR(item.qty * item.rate)}</td>
    `;
    el.prevItemsTbody.appendChild(tr);
  });

  // Render Installments in Preview
  el.prevInstallmentsTbody.innerHTML = '';
  if (state.installments.length === 0) {
    const tr = document.createElement('tr');
    tr.innerHTML = `<td colspan="4" style="text-align:center; color:#94a3b8; padding:8px;">100% Payment due within 15 days upon issuance.</td>`;
    el.prevInstallmentsTbody.appendChild(tr);
  } else {
    state.installments.forEach(inst => {
      const tr = document.createElement('tr');
      const hasTxn = inst.txn && inst.txn.trim().length > 0;
      tr.innerHTML = `
        <td style="font-weight:600; color:#0f172a;">${escapeHtml(inst.label || 'Tranche')}</td>
        <td style="color:#64748b;">${formatApisomiDate(inst.dueDate)}</td>
        <td style="text-align: right; font-weight:600;" class="mono-cell">${formatINR(inst.amount)}</td>
        <td style="text-align: center;">
          ${hasTxn 
            ? `<span class="txn-ref-badge">${escapeHtml(inst.txn)}</span>` 
            : `<span class="empty-txn-slot">Pending Txn / UTR</span>`}
        </td>
      `;
      el.prevInstallmentsTbody.appendChild(tr);
    });
  }

  // Summary Calculations
  el.prevSummarySubtotal.textContent = formatINR(calc.subtotal);
  el.prevDiscountPct.textContent = state.discountPct;
  el.prevSummaryDiscount.textContent = `– ${formatINR(calc.discountAmount)}`;
  el.prevSummaryTaxable.textContent = formatINR(calc.taxableValue);

  if (state.isGstMode) {
    if (calc.isIntraState) {
      el.prevCgstRow.classList.remove('hidden');
      el.prevSgstRow.classList.remove('hidden');
      el.prevIgstRow.classList.add('hidden');
      el.prevSummaryCgst.textContent = formatINR(calc.cgst);
      el.prevSummarySgst.textContent = formatINR(calc.sgst);
    } else {
      el.prevCgstRow.classList.add('hidden');
      el.prevSgstRow.classList.add('hidden');
      el.prevIgstRow.classList.remove('hidden');
      el.prevSummaryIgst.textContent = formatINR(calc.igst);
    }
  } else {
    el.prevCgstRow.classList.add('hidden');
    el.prevSgstRow.classList.add('hidden');
    el.prevIgstRow.classList.add('hidden');
  }

  const sign = calc.roundOff >= 0 ? '+' : '–';
  el.prevSummaryRoundoff.textContent = `${sign} ${Math.abs(calc.roundOff).toFixed(2)}`;
  el.prevSummaryGrandTotal.textContent = formatINR(calc.grandTotal);
  el.prevSummaryAdvance.textContent = `– ${formatINR(calc.advanceReceived)}`;
  el.prevSummaryBalance.textContent = formatINR(calc.balanceDue);

  // Amount in Words
  el.prevAmountWords.textContent = inrToWords(Math.round(calc.grandTotal));

  // Bank & Footer
  el.prevBankAccName.textContent = el.bankAccName.value;
  el.prevBankBranch.textContent = el.bankBranch.value;
  el.prevBankAccNo.textContent = el.bankAccNo.value;
  el.prevBankIfsc.textContent = el.bankIfsc.value;
  el.prevBankUpi.textContent = el.bankUpi.value;
  el.prevProjectNote.textContent = el.projectSpecificNote.value || 'None';
}

// PDF Export & Automatic Logging
async function exportPdfAndLog() {
  const calc = calculateFinancials();
  const proj = (el.projectCode.value.trim() || 'PROJECT').toUpperCase();
  const seq = (el.invSeq.value.trim() || '01');
  const invNumber = el.invNumber.value.trim();
  const pdfFilename = `Apisomi_${proj}_Invoice_${seq}.pdf`;

  // 1. Record in Ledger
  const record = {
    invoiceNo: invNumber,
    version: el.docVersion.value,
    invoiceDate: el.invDate.value,
    dueDate: el.invDueDate.value,
    projectCode: proj,
    projectName: el.projectName.value.trim(),
    milestone: el.milestoneDesc.value.trim(),
    clientName: el.clientName.value.trim(),
    clientState: `${calc.clientStateName} (${calc.clientCode})`,
    clientGstin: el.clientGstin.value.trim() || 'Unregistered',
    taxRegime: state.isGstMode ? (calc.isIntraState ? 'CGST (9%) + SGST (9%)' : 'IGST (18%)') : 'Non-GST',
    subtotal: calc.subtotal,
    discount: calc.discountAmount,
    taxableValue: calc.taxableValue,
    cgst: calc.cgst,
    sgst: calc.sgst,
    igst: calc.igst,
    roundOff: calc.roundOff,
    grandTotal: calc.grandTotal,
    advanceReceived: calc.advanceReceived,
    balanceDue: calc.balanceDue,
    installmentsCount: state.installments.length,
    loggedAt: new Date().toLocaleString('en-IN')
  };

  // Upsert
  const existingIdx = state.ledger.findIndex(r => r.invoiceNo === invNumber);
  if (existingIdx >= 0) {
    state.ledger[existingIdx] = record;
  } else {
    state.ledger.unshift(record);
  }
  saveLedger();
  updateLedgerCount();

  // 2. Export PDF
  const opt = {
    margin: [6, 6, 6, 6],
    filename: pdfFilename,
    image: { type: 'jpeg', quality: 0.98 },
    html2canvas: { scale: 2, useCORS: true, logging: false },
    jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
  };

  try {
    el.downloadPdfBtn.disabled = true;
    el.downloadPdfBtn.textContent = 'Generating PDF & Logging...';
    await html2pdf().set(opt).from(el.invoiceDocument).save();
    
    // Sync to Google Sheet if configured
    let gsheetMsg = '';
    if (state.gsheetWebhookUrl) {
      const gsheetSuccess = await pushToGoogleSheet(record);
      if (gsheetSuccess) {
        gsheetMsg = ' & synced to permanent Google Sheet 🟢';
      } else {
        gsheetMsg = ' (Google Sheet sync failed to reach URL)';
      }
    } else {
      gsheetMsg = ' (Connect Google Sheet in top nav for cloud sync)';
    }

    showToast(`Invoice ${invNumber} generated (${pdfFilename}) & logged in Excel tracker${gsheetMsg}!`);
  } catch (err) {
    console.error('PDF generation issue:', err);
    window.print();
  } finally {
    el.downloadPdfBtn.disabled = false;
    el.downloadPdfBtn.innerHTML = `
      <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
        <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"></path>
        <polyline points="7 10 12 15 17 10"></polyline>
        <line x1="12" y1="15" x2="12" y2="3"></line>
      </svg>
      Export PDF & Log
    `;
  }
}

// Complete Excel Workbook (.xlsx) Generator with cumulative totals
function generateExcelWorkbook() {
  if (state.ledger.length === 0) {
    alert('No invoice records to export yet. Generate an invoice first!');
    return;
  }

  const rows = state.ledger.map((r, i) => ({
    'Sr No': i + 1,
    'Invoice No': r.invoiceNo,
    'Date': r.invoiceDate,
    'Due Date': r.dueDate,
    'Project Code': r.projectCode,
    'Project Name': r.projectName,
    'Milestone': r.milestone,
    'Client Name': r.clientName,
    'Client State': r.clientState,
    'Client GSTIN': r.clientGstin,
    'Tax Regime': r.taxRegime,
    'Subtotal (₹)': r.subtotal,
    'Discount (₹)': r.discount,
    'Taxable Value (₹)': r.taxableValue,
    'CGST (₹)': r.cgst,
    'SGST (₹)': r.sgst,
    'IGST (₹)': r.igst,
    'Round-off (₹)': r.roundOff,
    'Grand Total (₹)': r.grandTotal,
    'Advance Paid (₹)': r.advanceReceived,
    'Balance Due (₹)': r.balanceDue,
    'Installments': `${r.installmentsCount} Tranches`,
    'Logged Timestamp': r.loggedAt
  }));

  // Summary row
  const sumTaxable = state.ledger.reduce((a, b) => a + b.taxableValue, 0);
  const sumTotal = state.ledger.reduce((a, b) => a + b.grandTotal, 0);
  const sumAdvance = state.ledger.reduce((a, b) => a + b.advanceReceived, 0);
  const sumBalance = state.ledger.reduce((a, b) => a + b.balanceDue, 0);

  rows.push({});
  rows.push({
    'Sr No': 'TOTALS',
    'Invoice No': `${state.ledger.length} Invoices`,
    'Taxable Value (₹)': sumTaxable,
    'Grand Total (₹)': sumTotal,
    'Advance Paid (₹)': sumAdvance,
    'Balance Due (₹)': sumBalance
  });

  const ws = XLSX.utils.json_to_sheet(rows);

  ws['!cols'] = [
    { wch: 6 },  // Sr No
    { wch: 20 }, // Invoice No
    { wch: 12 }, // Date
    { wch: 12 }, // Due Date
    { wch: 14 }, // Project Code
    { wch: 28 }, // Project Name
    { wch: 28 }, // Milestone
    { wch: 28 }, // Client Name
    { wch: 20 }, // Client State
    { wch: 18 }, // Client GSTIN
    { wch: 18 }, // Tax Regime
    { wch: 15 }, // Subtotal
    { wch: 14 }, // Discount
    { wch: 16 }, // Taxable Value
    { wch: 12 }, // CGST
    { wch: 12 }, // SGST
    { wch: 12 }, // IGST
    { wch: 12 }, // Round-off
    { wch: 16 }, // Grand Total
    { wch: 16 }, // Advance Paid
    { wch: 16 }, // Balance Due
    { wch: 14 }, // Installments
    { wch: 22 }  // Timestamp
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Invoices Ledger');

  const todayStr = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Apisomi_Invoices_Ledger_${todayStr}.xlsx`);
}

// LocalStorage
function saveLedger() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.ledger));
}

function loadLedger() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (raw) state.ledger = JSON.parse(raw);
  } catch (e) {
    state.ledger = [];
  }
  updateLedgerCount();
}

function updateLedgerCount() {
  el.ledgerCount.textContent = state.ledger.length;
}

// Ledger Modal
function openLedgerModal() {
  el.ledgerModal.classList.remove('hidden');
  renderLedgerTable();
}

function closeLedgerModal() {
  el.ledgerModal.classList.add('hidden');
}

function renderLedgerTable() {
  el.ledgerTbody.innerHTML = '';
  if (state.ledger.length === 0) {
    el.emptyLedgerMessage.classList.remove('hidden');
    el.metricCount.textContent = '0';
    el.metricTotal.textContent = '₹ 0.00';
    el.metricGstCount.textContent = '0';
    el.metricNonGstCount.textContent = '0';
    return;
  }

  el.emptyLedgerMessage.classList.add('hidden');

  let totalBilled = 0;
  let gstCount = 0;
  let nonGstCount = 0;

  state.ledger.forEach((rec, idx) => {
    totalBilled += rec.grandTotal;
    if (rec.taxRegime.includes('GST') && !rec.taxRegime.includes('Non-GST')) gstCount++;
    else nonGstCount++;

    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td class="mono font-semibold" style="color:#38bdf8;">${escapeHtml(rec.invoiceNo)}</td>
      <td>${rec.invoiceDate}</td>
      <td><strong>${escapeHtml(rec.projectCode)}</strong></td>
      <td>${escapeHtml(rec.clientName)}</td>
      <td><span class="mode-pill ${rec.taxRegime.includes('Non-GST') ? 'non-gst' : 'gst-active'}">${escapeHtml(rec.taxRegime)}</span></td>
      <td style="text-align: right;" class="mono">₹ ${formatINR(rec.taxableValue)}</td>
      <td style="text-align: right; font-weight:700;" class="mono">₹ ${formatINR(rec.grandTotal)}</td>
      <td style="text-align: right; color:#fbbf24;" class="mono">₹ ${formatINR(rec.balanceDue)}</td>
      <td style="text-align: center;">
        <button class="btn-del-row" data-action="del-record" data-idx="${idx}" title="Delete record">
          <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
            <polyline points="3 6 5 6 21 6"></polyline>
            <path d="M19 6v14a2 2 0 0 1-2 2H7a2 2 0 0 1-2-2V6m3 0V4a2 2 0 0 1 2-2h4a2 2 0 0 1 2 2v2"></path>
          </svg>
        </button>
      </td>
    `;
    el.ledgerTbody.appendChild(tr);
  });

  el.ledgerTbody.querySelectorAll('[data-action="del-record"]').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const idx = parseInt(e.currentTarget.dataset.idx);
      state.ledger.splice(idx, 1);
      saveLedger();
      updateLedgerCount();
      renderLedgerTable();
    });
  });

  el.metricCount.textContent = state.ledger.length;
  el.metricTotal.textContent = `₹ ${formatINR(totalBilled)}`;
  el.metricGstCount.textContent = gstCount;
  el.metricNonGstCount.textContent = nonGstCount;
}

function clearLedger() {
  if (confirm('Clear all invoice logs from this browser?')) {
    state.ledger = [];
    saveLedger();
    updateLedgerCount();
    renderLedgerTable();
  }
}

// Reset to default sample from user template
function resetToDefaults() {
  if (!confirm('Reset fields to the Apisomi template defaults?')) return;
  state.projectCode = 'SSV003';
  state.invSeq = '01';
  el.projectCode.value = 'SSV003';
  el.invSeq.value = '01';
  el.docVersion.value = 'v1.0 · Issued';
  el.clientName.value = 'Acme Digital Systems Pvt Ltd';
  el.clientContact.value = 'Vikram Sharma, Director of Technology';
  el.clientAddress.value = 'Plot 42, Sector 18, Udyog Vihar, Gurugram, Haryana – 122015';
  el.clientGstin.value = '06AAACA1234A1Z5';
  el.clientPan.value = 'AAACA1234A';
  el.clientStateSelect.value = '06|Haryana';
  el.projectName.value = 'Drone Fleet Telemetry & Cloud Platform';
  el.projectPeriod.value = '01 Oct 2026 – 25 Oct 2026';
  el.discountPercent.value = 0;
  state.discountPct = 0;
  el.advanceReceivedInput.value = 0;
  state.advanceReceived = 0;

  state.items = [
    { id: 1, desc: 'UI/UX design — 12 screens & design system prototype', sac: '998314', qty: 1, unit: 'Lot', rate: 45000 },
    { id: 2, desc: 'Android & iOS telemetry app development — Sprint 1', sac: '998314', qty: 40, unit: 'Hours', rate: 1200 },
    { id: 3, desc: 'Cloud API, drone telemetry & automated database deployment', sac: '998314', qty: 1, unit: 'Lot', rate: 25000 }
  ];

  state.installments = [
    { id: 1, label: 'Milestone 1 — Architecture & Wireframes (50%)', dueDate: '2026-10-15', amount: 59000, txn: 'UTR-HDFC-9928192' },
    { id: 2, label: 'Milestone 2 — Design Sign-off & Sprint 1 (30%)', dueDate: '2026-10-25', amount: 35400, txn: '' },
    { id: 3, label: 'Milestone 3 — Production Handover (20%)', dueDate: '2026-11-15', amount: 23600, txn: '' }
  ];

  computeInvoiceNumber();
  renderItemsTable();
  renderInstallmentsTable();
  updateLivePreview();
}

// Helpers
function formatINR(num) {
  if (isNaN(num)) return '0.00';
  return Number(num).toLocaleString('en-IN', {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2
  });
}

function formatApisomiDate(dateStr) {
  if (!dateStr) return '[DD Mon YYYY]';
  const d = new Date(dateStr);
  if (isNaN(d.getTime())) return dateStr;
  const day = String(d.getDate()).padStart(2, '0');
  const month = d.toLocaleString('en-IN', { month: 'short' });
  const year = d.getFullYear();
  return `${day} ${month} ${year}`;
}

function escapeHtml(str) {
  if (!str) return '';
  return str.toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// Indian Numbering to Words (Rupees ... Lakhs ... Only)
function inrToWords(amount) {
  if (amount === 0) return 'Indian Rupees Zero Only.';

  const units = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine'];
  const teens = ['Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
  const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];

  function convertTwoDigits(n) {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    const unit = n % 10;
    return tens[Math.floor(n / 10)] + (unit !== 0 ? ' ' + units[unit] : '');
  }

  function convertThreeDigits(n) {
    const hundred = Math.floor(n / 100);
    const remainder = n % 100;
    let str = '';
    if (hundred > 0) str += units[hundred] + ' Hundred';
    if (remainder > 0) {
      if (str !== '') str += ' and ';
      str += convertTwoDigits(remainder);
    }
    return str;
  }

  let crore = Math.floor(amount / 10000000);
  amount %= 10000000;
  let lakh = Math.floor(amount / 100000);
  amount %= 100000;
  let thousand = Math.floor(amount / 1000);
  amount %= 1000;
  let remainder = amount;

  let words = '';

  if (crore > 0) words += convertThreeDigits(crore) + ' Crore ';
  if (lakh > 0) words += convertThreeDigits(lakh) + ' Lakh ';
  if (thousand > 0) words += convertThreeDigits(thousand) + ' Thousand ';
  if (remainder > 0) words += convertThreeDigits(remainder);

  return 'Indian Rupees ' + words.trim() + ' Only.';
}

// =========================================================
// Google Sheets Integration & Webhook Sync Engine
// =========================================================

function initGoogleSheetSync() {
  el.gsheetWebhookUrl.value = state.gsheetWebhookUrl;
  el.gsheetDirectUrl.value = state.gsheetDirectUrl;
  updateGoogleSheetStatusUI();
}

function updateGoogleSheetStatusUI() {
  const isConnected = Boolean(state.gsheetWebhookUrl && state.gsheetWebhookUrl.trim());
  if (isConnected) {
    el.gsheetStatusDot.className = 'status-dot dot-active';
    el.gsheetBtnText.textContent = 'Google Sheet';
    el.syncStatusTitle.textContent = 'Google Sheet: Connected 🟢';
    el.syncStatusSubtitle.textContent = 'Each generated invoice will be automatically appended to your permanent Google Sheet.';
    el.syncStatusCard.style.borderColor = 'rgba(16, 185, 129, 0.4)';
    if (state.gsheetDirectUrl && state.gsheetDirectUrl.trim()) {
      el.openSheetLink.href = state.gsheetDirectUrl.trim();
      el.openSheetLink.classList.remove('hidden');
    } else {
      el.openSheetLink.classList.add('hidden');
    }
  } else {
    el.gsheetStatusDot.className = 'status-dot dot-inactive';
    el.gsheetBtnText.textContent = 'Link Google Sheet';
    el.syncStatusTitle.textContent = 'Google Sheet: Not Configured';
    el.syncStatusSubtitle.textContent = 'Connect your Google Sheet Webhook once to enable automatic cloud logging on every invoice export.';
    el.syncStatusCard.style.borderColor = 'var(--border-app)';
    el.openSheetLink.classList.add('hidden');
  }
}

function openGsheetModal() {
  el.gsheetModal.classList.remove('hidden');
}

function closeGsheetModal() {
  el.gsheetModal.classList.add('hidden');
}

function saveGoogleSheetConfig() {
  state.gsheetWebhookUrl = el.gsheetWebhookUrl.value.trim();
  state.gsheetDirectUrl = el.gsheetDirectUrl.value.trim();

  localStorage.setItem('apisomi_gsheet_webhook', state.gsheetWebhookUrl);
  localStorage.setItem('apisomi_gsheet_direct', state.gsheetDirectUrl);

  updateGoogleSheetStatusUI();
  showToast(state.gsheetWebhookUrl ? 'Google Sheet connection saved!' : 'Google Sheet URL cleared.');
  closeGsheetModal();
}

async function testGoogleSheetConnection() {
  const url = el.gsheetWebhookUrl.value.trim();
  if (!url) {
    alert('Please enter a Google Apps Script Web App URL first.');
    return;
  }

  el.testGsheetSyncBtn.disabled = true;
  el.testGsheetSyncBtn.textContent = 'Testing...';

  try {
    const testPayload = {
      invoiceNo: 'APS-TEST-001',
      invoiceDate: new Date().toISOString().split('T')[0],
      dueDate: new Date().toISOString().split('T')[0],
      projectCode: 'TEST',
      projectName: 'Connection Test Entry',
      milestone: 'Test Row',
      clientName: 'Apisomi Test Client',
      clientState: 'Madhya Pradesh (23)',
      clientGstin: 'Unregistered',
      taxRegime: 'Non-GST',
      subtotal: 1000,
      discount: 0,
      taxableValue: 1000,
      cgst: 0,
      sgst: 0,
      igst: 0,
      roundOff: 0,
      grandTotal: 1000,
      advanceReceived: 0,
      balanceDue: 1000,
      installmentsCount: 1,
      loggedAt: new Date().toLocaleString('en-IN')
    };

    // Use mode: 'no-cors' to avoid browser cross-origin redirect block on Google Apps Script Web App
    await fetch(url, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(testPayload)
    });

    alert('✅ Test signal sent to your Google Sheet! Please check your sheet to confirm the test row was added.');
    state.gsheetWebhookUrl = url;
    localStorage.setItem('apisomi_gsheet_webhook', url);
    updateGoogleSheetStatusUI();
  } catch (err) {
    console.error('Google Sheet Test Error:', err);
    alert('Could not reach Webhook URL. Please verify your Google Apps Script deployment settings (Who has access: Anyone).');
  } finally {
    el.testGsheetSyncBtn.disabled = false;
    el.testGsheetSyncBtn.textContent = 'Test Connection';
  }
}

async function pushToGoogleSheet(record) {
  if (!state.gsheetWebhookUrl) return false;

  try {
    await fetch(state.gsheetWebhookUrl, {
      method: 'POST',
      mode: 'no-cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(record)
    });
    return true;
  } catch (err) {
    console.error('Google Sheet push error:', err);
    return false;
  }
}

async function handleManualSheetUpdate() {
  if (!state.gsheetWebhookUrl) {
    openGsheetModal();
    return;
  }

  const calc = calculateFinancials();
  const proj = (el.projectCode.value.trim() || 'PROJECT').toUpperCase();
  const invNumber = el.invNumber.value.trim();

  const record = {
    invoiceNo: invNumber,
    version: el.docVersion.value,
    invoiceDate: el.invDate.value,
    dueDate: el.invDueDate.value,
    projectCode: proj,
    projectName: el.projectName.value.trim(),
    milestone: el.milestoneDesc.value.trim(),
    clientName: el.clientName.value.trim(),
    clientState: `${calc.clientStateName} (${calc.clientCode})`,
    clientGstin: el.clientGstin.value.trim() || 'Unregistered',
    taxRegime: state.isGstMode ? (calc.isIntraState ? 'CGST (9%) + SGST (9%)' : 'IGST (18%)') : 'Non-GST',
    subtotal: calc.subtotal,
    discount: calc.discountAmount,
    taxableValue: calc.taxableValue,
    cgst: calc.cgst,
    sgst: calc.sgst,
    igst: calc.igst,
    roundOff: calc.roundOff,
    grandTotal: calc.grandTotal,
    advanceReceived: calc.advanceReceived,
    balanceDue: calc.balanceDue,
    installmentsCount: state.installments.length,
    loggedAt: new Date().toLocaleString('en-IN')
  };

  el.syncNowBtn.disabled = true;
  el.syncNowBtn.textContent = 'Updating...';

  const ok = await pushToGoogleSheet(record);

  el.syncNowBtn.disabled = false;
  el.syncNowBtn.innerHTML = `
    <svg width="17" height="17" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
      <polyline points="23 4 23 10 17 10"></polyline>
      <polyline points="1 20 1 14 7 14"></polyline>
      <path d="M3.51 9a9 9 0 0 1 14.85-3.36L23 10M1 14l4.64 4.36A9 9 0 0 0 20.49 15"></path>
    </svg>
    Update Sheet
  `;

  if (ok) {
    showToast(`Invoice ${invNumber} pushed to Google Sheet (invoice_tracking) 🟢!`);
  } else {
    alert('Failed to update Google Sheet. Please check your Webhook URL or permissions.');
  }
}

function copyGoogleAppsScript() {
  navigator.clipboard.writeText(GSCRIPT_TEMPLATE).then(() => {
    el.copyScriptBtn.textContent = 'Copied to Clipboard! ✓';
    setTimeout(() => {
      el.copyScriptBtn.textContent = 'Copy Google Apps Script';
    }, 3000);
  }).catch(() => {
    prompt('Copy this Google Apps Script code manually:', GSCRIPT_TEMPLATE);
  });
}

function showToast(msg) {
  const toast = document.createElement('div');
  toast.style.position = 'fixed';
  toast.style.bottom = '20px';
  toast.style.right = '20px';
  toast.style.background = '#0f172a';
  toast.style.border = '1px solid #3b82f6';
  toast.style.color = '#fff';
  toast.style.padding = '12px 18px';
  toast.style.borderRadius = '8px';
  toast.style.boxShadow = '0 10px 25px rgba(0,0,0,0.4)';
  toast.style.fontSize = '0.84rem';
  toast.style.fontWeight = '600';
  toast.style.zIndex = '9999';
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 4500);
}

window.addEventListener('DOMContentLoaded', init);
