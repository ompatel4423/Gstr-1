/**
 * 👑 GSTR-1 ENTERPRISE ENGINE & GSTN SPECIFICATION PORTAL JSON COMPILER
 * Official Schema Parity: Direct Portal Upload (Clean Envelope / Optional v3.2.4),
 * Table 12 HSN (with mandatory 'val', Bifurcated & Unified modes), Table 13 doc_issue (v4.1 & v5.0)
 *
 * Repository: Gstr-1
 * Author: ompatel4423
 * License: MIT
 */

const GSTIN_REGEX = /^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z]{1}[1-9A-Z]{1}Z[0-9A-Z]{1}$/;
const INUM_REGEX = /^(?=.{1,16}$)([/\-0]*[a-zA-Z0-9/\-]*[a-zA-Z1-9]+[a-zA-Z0-9/\-]*)$/;
const STATE_CODE_REGEX = /^(3[0-8]|[12][0-9]|0[1-9]|9[67])$/;
const HSN_REGEX = /^[0-9]{2,8}$/;
const DOC_NUM_MAP = {
  'Invoices for outward supply': 1,
  'Invoices for inward supply from unregistered person': 2,
  'Revised Invoice': 3,
  'Debit Note': 4,
  'Credit Note': 5,
  'Receipt Voucher': 6,
  'Payment Voucher': 7,
  'Refund Voucher': 8,
  'Delivery Challan for job work': 9,
  'Delivery Challan for supply on approval': 10,
  'Delivery Challan in case of liquid gas': 11,
  'Delivery Challan in cases other than by way of supply': 12
};
const INV_TYPE_MAP = {
  REGULAR: 'R',
  REG: 'R',
  R: 'R',
  DE: 'DE',
  'DEEMED EXP': 'DE',
  SEWP: 'SEWP',
  SEWOP: 'SEWOP',
  CBW: 'CBW'
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('⚡ GSTR-1 Automation')
    .addItem('🚀 1. Process & Segregate All Tables', 'processGSTR1DataStatutory')
    .addItem('🔍 2. Run Audit & Cross-Reconciliation', 'runReconAudit')
    .addSeparator()
    .addItem('📦 3. Export GST Portal Upload Ready JSON', 'exportGSTR1Json')
    .addSeparator()
    .addItem('🧹 4. Reset & Clear All Data for Next Client', 'resetSheetForNewClient')
    .addToUi();
}

/**
 * 1. Statutory Processor & Segregator
 */
function processGSTR1DataStatutory() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getRawSheet(ss);
  const config = getConfigSettings(ss);

  if (!rawSheet) {
    SpreadsheetApp.getUi().alert("Error: 'Invoices_Raw' sheet tab not found! 🥀");
    return;
  }

  const data = rawSheet.getDataRange().getValues();
  if (data.length <= 1) {
    clearAllSummaryTabs(ss);
    SpreadsheetApp.getUi().alert("✨ 'Invoices_Raw' is empty. All summary tabs have been automatically cleared and reset! 💅🏻✌️");
    return;
  }

  const model = buildGSTR1Model(data.slice(1), config, ss);

  writeToSheet(ss, 'B2B_Summary', model.sheets.b2bRows, 2);
  writeToSheet(ss, 'B2C_Large', model.sheets.b2clRows, 2);
  writeToSheet(ss, 'B2C_Small', model.sheets.b2csRows, 2);
  writeToSheet(ss, 'HSN_Summary', model.sheets.hsnRows, 2);
  writeToSheet(ss, 'CDNR_Summary', model.sheets.cdnrRows, 2);
  writeToSheet(ss, 'Docs_Summary', model.sheets.docsRows, 2);

  if (model.auditWarnings.length > 0) {
    SpreadsheetApp.getUi().alert(
      `⚠️ Segregated with ${model.auditWarnings.length} alert(s):\n\n` +
      model.auditWarnings.slice(0, 5).join('\n')
    );
  } else {
    SpreadsheetApp.getUi().alert("✨ Slay! GSTR-1 tables successfully compiled and recalculated with 0 errors! 💅🏻✌️");
  }
}

/**
 * 2. Cross-Reconciliation & Audit
 */
function runReconAudit() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getRawSheet(ss);
  if (!rawSheet) return;

  const config = getConfigSettings(ss);
  const rawRows = rawSheet.getDataRange().getValues().slice(1);
  const model = buildGSTR1Model(rawRows, config, ss);

  let rawNetTaxable = 0;
  let rawOutwardGrossTaxable = 0;

  rawRows.forEach(r => {
    const invNo = String(r[0] || '').trim();
    const isCancelled = String(r[12] || 'N').trim().toUpperCase() === 'Y';
    if (!invNo || isCancelled) return;

    const docType = String(r[6] || 'Regular').trim().toUpperCase();
    const taxable = parseAmount(r[8]);
    if (docType === 'CRN') {
      rawNetTaxable -= taxable;
    } else if (docType === 'DBN') {
      rawNetTaxable += taxable;
    } else {
      rawNetTaxable += taxable;
      rawOutwardGrossTaxable += taxable;
    }
  });

  const getSum = (sheetName, colIdx) => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return 0;
    return data.slice(1).reduce((acc, r) => acc + parseAmount(r[colIdx]), 0);
  };

  const b2bTaxable = getSum('B2B_Summary', 9);
  const b2clTaxable = getSum('B2C_Large', 5);
  const b2csTaxable = getSum('B2C_Small', 3);

  let cdnrNetTaxable = 0;
  const cdnrSheet = ss.getSheetByName('CDNR_Summary');
  if (cdnrSheet && cdnrSheet.getLastRow() >= 2) {
    const cdnrData = cdnrSheet.getDataRange().getValues().slice(1);
    cdnrData.forEach(r => {
      const type = String(r[2] || '').toUpperCase();
      const val = parseAmount(r[11]);
      if (type === 'CRN') cdnrNetTaxable -= val;
      else if (type === 'DBN') cdnrNetTaxable += val;
    });
  }

  let cdnurNetTaxable = 0;
  model.json.cdnur.forEach(note => {
    const noteTaxable = note.itms.reduce((s, itm) => s + (itm.itm_det.txval || 0), 0);
    cdnurNetTaxable += (note.ntty === 'C' ? -noteTaxable : noteTaxable);
  });

  const hsnTaxable = getSum('HSN_Summary', 4);
  const totalTablesTaxable = round2(b2bTaxable + b2clTaxable + b2csTaxable + cdnrNetTaxable + cdnurNetTaxable);
  const diffTables = round2(Math.abs(rawNetTaxable - totalTablesTaxable));
  const diffHsn = round2(Math.abs(rawOutwardGrossTaxable - hsnTaxable));

  let msg = `📊 GSTR-1 STATUTORY AUDIT & RECONCILIATION:\n\n`;
  msg += `• Net Raw Active Taxable: ₹${round2(rawNetTaxable)}\n`;
  msg += `• Table 4 (B2B): ₹${round2(b2bTaxable)}\n`;
  msg += `• Table 5 (B2C Large): ₹${round2(b2clTaxable)}\n`;
  msg += `• Table 7 (B2C Small): ₹${round2(b2csTaxable)}\n`;
  if (cdnrNetTaxable !== 0) msg += `• Table 9B (CDNR Net): ₹${round2(cdnrNetTaxable)}\n`;
  if (cdnurNetTaxable !== 0) msg += `• Table 9B (CDNUR Net): ₹${round2(cdnurNetTaxable)}\n`;
  msg += `--------------------------------------\n`;
  msg += `• Total Net Segregated: ₹${totalTablesTaxable} (Diff: ₹${diffTables})\n`;
  msg += `• Table 12 (HSN Outward): ₹${round2(hsnTaxable)} vs Gross Outward ₹${round2(rawOutwardGrossTaxable)} (Diff: ₹${diffHsn})\n\n`;

  if (diffTables < 1 && diffHsn < 1) {
    msg += `✅ 100% RECONCILED! Ready for GST Portal filing. No mismatch detected! 💅🏻🔥`;
  } else {
    msg += `❌ RECONCILIATION MISMATCH! Please verify missing HSNs, cancelled bills, or state codes. 🥀`;
  }

  SpreadsheetApp.getUi().alert(msg);
}

/**
 * 3. Direct GST Portal Upload-Ready JSON Exporter
 */
function exportGSTR1Json() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getRawSheet(ss);
  const config = getConfigSettings(ss);

  if (!rawSheet) {
    SpreadsheetApp.getUi().alert("Missing 'Invoices_Raw' sheet tab!");
    return;
  }

  const rawRows = rawSheet.getDataRange().getValues().slice(1);
  const model = buildGSTR1Model(rawRows, config, ss);
  const docsSheet = ss.getSheetByName('Docs_Summary');

  // Build clean tables payload (without version/hash by default - Finexo benchmark)
  const tablesPayload = {};
  if (model.json.b2b.length > 0) tablesPayload.b2b = model.json.b2b;
  if (model.json.b2cl.length > 0) tablesPayload.b2cl = model.json.b2cl;
  if (model.json.b2cs.length > 0) tablesPayload.b2cs = model.json.b2cs;
  if (model.json.cdnr.length > 0) tablesPayload.cdnr = model.json.cdnr;
  if (model.json.cdnur.length > 0) tablesPayload.cdnur = model.json.cdnur;

  // HSN variants
  const hsnBifurcated = {};
  if (model.json.hsnB2b.length > 0) hsnBifurcated.hsn_b2b = model.json.hsnB2b;
  if (model.json.hsnB2c.length > 0) hsnBifurcated.hsn_b2c = model.json.hsnB2c;

  const hsnUnified = {};
  if (model.json.hsnUnified.length > 0) hsnUnified.data = model.json.hsnUnified;

  // Doc Issue variants
  const docIssueV41 = buildDocIssuePayload(docsSheet, model.docSummaryList, false);
  const docIssueV50 = buildDocIssuePayload(docsSheet, model.docSummaryList, true);

  const payloadBundle = {
    gstin: config.gstin,
    fp: config.fp,
    tables: tablesPayload,
    hsnBifurcated: Object.keys(hsnBifurcated).length > 0 ? hsnBifurcated : null,
    hsnUnified: Object.keys(hsnUnified).length > 0 ? hsnUnified : null,
    docIssueV41: docIssueV41,
    docIssueV50: docIssueV50
  };

  const safeBundleJson = JSON.stringify(payloadBundle).replace(/</g, '\\u003c');

  const htmlOutput = HtmlService.createHtmlOutput(
    `<!DOCTYPE html>
    <html>
      <head>
        <style>
          body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; padding: 14px; margin: 0; background: #FAFAFA; color: #202124; }
          .header { display: flex; justify-content: space-between; align-items: baseline; border-bottom: 1px solid #E0E0E0; padding-bottom: 8px; margin-bottom: 12px; }
          .header h3 { margin: 0; color: #1A73E8; font-size: 16px; font-weight: 600; }
          .info { font-size: 12px; color: #5F6368; }
          .control-panel { background: #FFFFFF; border: 1px solid #DADCE0; border-radius: 6px; padding: 10px 14px; margin-bottom: 10px; }
          .control-row { display: flex; gap: 18px; align-items: center; flex-wrap: wrap; font-size: 12px; color: #3C4043; }
          .control-row label { display: flex; align-items: center; gap: 5px; cursor: pointer; }
          .actions { display: flex; justify-content: space-between; align-items: center; margin-bottom: 10px; }
          .btn-download { background: #1A73E8; color: #FFF; border: none; padding: 8px 18px; border-radius: 4px; font-weight: 600; font-size: 13px; cursor: pointer; transition: background 0.2s; }
          .btn-download:hover { background: #1557B0; }
          .status-badge { font-size: 11px; padding: 3px 8px; border-radius: 12px; background: #E6F4EA; color: #137333; font-weight: 500; }
          textarea { width: 100%; height: 250px; font-family: 'SFMono-Regular', Consolas, Monaco, monospace; font-size: 11px; line-height: 1.4; padding: 8px; border: 1px solid #DADCE0; border-radius: 4px; box-sizing: border-box; background: #FFF; color: #1F2328; resize: none; }
        </style>
      </head>
      <body>
        <div class="header">
          <h3>📦 GSTR-1 Portal Upload Ready JSON</h3>
          <span class="info">GSTIN: <b>${escapeHtml(config.gstin)}</b> | Period: <b>${escapeHtml(config.fp)}</b></span>
        </div>

        <div class="control-panel">
          <div class="control-row">
            <label title="Direct upload format used by Finexo and accounting software without dummy hash/version">
              <input type="checkbox" id="cleanEnvelopeToggle" checked /> Clean Root (No Version/Hash - Finexo Mode)
            </label>
            <label title="Bifurcated (hsn_b2b & hsn_b2c) vs Unified single data array">
              <input type="checkbox" id="hsnBifurcateToggle" checked /> Bifurcated HSN (hsn_b2b / hsn_b2c)
            </label>
            <label title="Table 13 doc_det vs doc_det_new (API v5.0)">
              <input type="checkbox" id="v5Toggle" /> API v5.0 Table 13 (<code>doc_det_new</code>)
            </label>
          </div>
        </div>

        <div class="actions">
          <button id="dlBtn" class="btn-download">💾 Download .json File</button>
          <span class="status-badge" id="statsBadge">Calculating...</span>
        </div>

        <textarea id="jsonArea" readonly></textarea>

        <script>
          const bundle = ${safeBundleJson};
          const cleanToggle = document.getElementById('cleanEnvelopeToggle');
          const hsnToggle = document.getElementById('hsnBifurcateToggle');
          const v5Toggle = document.getElementById('v5Toggle');
          const area = document.getElementById('jsonArea');
          const statsBadge = document.getElementById('statsBadge');

          function compilePayload() {
            const out = {
              gstin: bundle.gstin,
              fp: bundle.fp
            };

            if (!cleanToggle.checked) {
              out.version = "GST3.2.4";
              out.hash = "hash";
            }

            // Copy invoice tables
            Object.keys(bundle.tables).forEach(k => {
              out[k] = bundle.tables[k];
            });

            // Handle HSN
            if (hsnToggle.checked) {
              if (bundle.hsnBifurcated) out.hsn = bundle.hsnBifurcated;
            } else {
              if (bundle.hsnUnified) out.hsn = bundle.hsnUnified;
            }

            // Handle Table 13
            const docIssue = v5Toggle.checked ? bundle.docIssueV50 : bundle.docIssueV41;
            if (docIssue) out.doc_issue = docIssue;

            return out;
          }

          function render() {
            const payload = compilePayload();
            const jsonText = JSON.stringify(payload, null, 2);
            area.value = jsonText;
            const b2bCount = payload.b2b ? payload.b2b.length : 0;
            const sizeKb = (new Blob([jsonText]).size / 1024).toFixed(2);
            statsBadge.textContent = 'Size: ' + sizeKb + ' KB | B2B Recipients: ' + b2bCount;
          }

          cleanToggle.addEventListener('change', render);
          hsnToggle.addEventListener('change', render);
          v5Toggle.addEventListener('change', render);
          render();

          document.getElementById('dlBtn').addEventListener('click', function() {
            const blob = new Blob([area.value], { type: 'application/json' });
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = '${escapeHtml(config.gstin)}_GSTR1_${escapeHtml(config.fp)}.json';
            a.click();
          });
        </script>
      </body>
    </html>`
  ).setWidth(680).setHeight(470);

  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'GSTR-1 JSON Generator');
}

/**
 * Shared compiler for both Sheet Segregation and JSON Export
 */
function buildGSTR1Model(rawRows, config, ss) {
  const invoiceGroups = {};
  const hsnSheetMap = {};
  const hsnB2bMap = {};
  const hsnB2cMap = {};
  const docTracker = {};
  const auditWarnings = [];

  rawRows.forEach((r, idx) => {
    const rowNum = idx + 2;
    const invNo = String(r[0] || '').trim();
    if (!invNo) return;

    if (!INUM_REGEX.test(invNo)) {
      auditWarnings.push(`Row ${rowNum} (${invNo}): Invoice No must be 1-16 chars (A-Z, 0-9, /, -; not all zeros).`);
    }

    const invDate = formatDate(r[1], ss);
    const buyerGstin = String(r[2] || '').trim().toUpperCase();
    const buyerName = String(r[3] || '').trim();
    const posRaw = String(r[4] || '').trim();
    const rcm = String(r[5] || 'N').trim().toUpperCase() === 'Y' ? 'Y' : 'N';
    const docTypeRaw = String(r[6] || 'Regular').trim().toUpperCase();
    const isCreditDebit = (docTypeRaw === 'CRN' || docTypeRaw === 'DBN');
    const rate = parseTaxRate(r[7]);
    const taxableVal = parseAmount(r[8]);
    const hsn = String(r[9] || '').trim();
    const uqc = cleanUQC(r[10]);
    const qty = round2(Number(r[11]) || 0);
    const isCancelled = String(r[12] || 'N').trim().toUpperCase() === 'Y';
    const origInvNo = String(r[13] || '').trim();
    const origInvDate = formatDate(r[14], ss);

    const docNature = docTypeRaw === 'CRN'
      ? 'Credit Note'
      : (docTypeRaw === 'DBN' ? 'Debit Note' : 'Invoices for outward supply');
    if (!docTracker[docNature]) {
      docTracker[docNature] = { nature: docNature, seen: {}, cancelled: {}, order: [] };
    }
    if (!docTracker[docNature].seen[invNo]) {
      docTracker[docNature].seen[invNo] = true;
      docTracker[docNature].order.push(invNo);
    }
    if (isCancelled) {
      docTracker[docNature].cancelled[invNo] = true;
      return;
    }

    const posMatch = posRaw.match(/\d+/);
    let posDigits = posMatch ? posMatch[0].padStart(2, '0') : config.supplierState;
    if (!STATE_CODE_REGEX.test(posDigits)) {
      auditWarnings.push(`Row ${rowNum} (${invNo}): Invalid POS state code "${posRaw}".`);
      posDigits = config.supplierState;
    }

    const invType = INV_TYPE_MAP[docTypeRaw] || 'R';
    const isInterState = Boolean(
      (config.supplierState && posDigits && config.supplierState !== posDigits) ||
      invType === 'SEWP' || invType === 'SEWOP'
    );

    let igst = 0, cgst = 0, sgst = 0;
    if (invType !== 'SEWOP') {
      if (isInterState) {
        igst = round2(taxableVal * (rate / 100));
      } else {
        cgst = round2(taxableVal * (rate / 200));
        sgst = round2(taxableVal * (rate / 200));
      }
    }
    const lineTotal = round2(taxableVal + igst + cgst + sgst);

    const hasValidStructure = Boolean(buyerGstin && GSTIN_REGEX.test(buyerGstin));
    if (buyerGstin && !hasValidStructure) {
      auditWarnings.push(`Row ${rowNum} (${invNo}): GSTIN "${buyerGstin}" has invalid format/length.`);
    } else if (buyerGstin && hasValidStructure && !validateGSTINChecksum(buyerGstin)) {
      auditWarnings.push(`Row ${rowNum} (${invNo}): GSTIN "${buyerGstin}" failed Modulus-36 checksum.`);
    }

    const groupKey = `${docTypeRaw}_${invNo}`;
    if (!invoiceGroups[groupKey]) {
      invoiceGroups[groupKey] = {
        invNo: invNo,
        invDate: invDate,
        buyerGstin: buyerGstin,
        hasValidGstin: hasValidStructure,
        buyerName: buyerName,
        posRaw: posRaw || posDigits,
        posDigits: posDigits,
        rcm: rcm,
        docType: docTypeRaw,
        invType: invType,
        isCreditDebit: isCreditDebit,
        isInterState: isInterState,
        origInvNo: origInvNo,
        origInvDate: origInvDate,
        totalInvValue: 0,
        rateMap: {},
        rawItems: []
      };
    }

    const grp = invoiceGroups[groupKey];
    grp.totalInvValue = round2(grp.totalInvValue + lineTotal);
    grp.rawItems.push({ rate: rate, taxable: taxableVal, igst: igst, cgst: cgst, sgst: sgst });

    const rateKey = String(rate);
    if (!grp.rateMap[rateKey]) {
      grp.rateMap[rateKey] = { rt: rate, txval: 0, iamt: 0, camt: 0, samt: 0 };
    }
    grp.rateMap[rateKey].txval = round2(grp.rateMap[rateKey].txval + taxableVal);
    grp.rateMap[rateKey].iamt = round2(grp.rateMap[rateKey].iamt + igst);
    grp.rateMap[rateKey].camt = round2(grp.rateMap[rateKey].camt + cgst);
    grp.rateMap[rateKey].samt = round2(grp.rateMap[rateKey].samt + sgst);

    if (hsn && !isCreditDebit) {
      if (!HSN_REGEX.test(hsn)) {
        auditWarnings.push(`Row ${rowNum} (${invNo}): HSN "${hsn}" must be 2-8 digits.`);
      }
      const hsnKey = `${hsn}_${uqc}_${rate}`;
      if (!hsnSheetMap[hsnKey]) {
        hsnSheetMap[hsnKey] = { hsn: hsn, uqc: uqc, qty: 0, totalVal: 0, taxable: 0, rate: rate, igst: 0, cgst: 0, sgst: 0 };
      }
      hsnSheetMap[hsnKey].qty = round2(hsnSheetMap[hsnKey].qty + qty);
      hsnSheetMap[hsnKey].totalVal = round2(hsnSheetMap[hsnKey].totalVal + lineTotal);
      hsnSheetMap[hsnKey].taxable = round2(hsnSheetMap[hsnKey].taxable + taxableVal);
      hsnSheetMap[hsnKey].igst = round2(hsnSheetMap[hsnKey].igst + igst);
      hsnSheetMap[hsnKey].cgst = round2(hsnSheetMap[hsnKey].cgst + cgst);
      hsnSheetMap[hsnKey].sgst = round2(hsnSheetMap[hsnKey].sgst + sgst);

      const targetMap = hasValidStructure ? hsnB2bMap : hsnB2cMap;
      if (!targetMap[hsnKey]) {
        targetMap[hsnKey] = { hsn_sc: hsn, uqc: uqc, qty: 0, val: 0, txval: 0, rt: rate, iamt: 0, camt: 0, samt: 0, csamt: 0 };
      }
      targetMap[hsnKey].qty = round2(targetMap[hsnKey].qty + qty);
      targetMap[hsnKey].val = round2(targetMap[hsnKey].val + lineTotal);
      targetMap[hsnKey].txval = round2(targetMap[hsnKey].txval + taxableVal);
      targetMap[hsnKey].iamt = round2(targetMap[hsnKey].iamt + igst);
      targetMap[hsnKey].camt = round2(targetMap[hsnKey].camt + cgst);
      targetMap[hsnKey].samt = round2(targetMap[hsnKey].samt + sgst);
    } else if (!hsn && !isCreditDebit) {
      auditWarnings.push(`Row ${rowNum} (${invNo}): Missing mandatory HSN code!`);
    }
  });

  const b2bRows = [];
  const b2clRows = [];
  const b2csSheetMap = {};
  const cdnrRows = [];

  const b2bCustomers = {};
  const b2clPosMap = {};
  const b2csJsonMap = {};
  const cdnrCustomers = {};
  const cdnurList = [];

  Object.values(invoiceGroups).forEach(inv => {
    const itmsList = Object.values(inv.rateMap).map((det, i) => ({
      num: i + 1,
      itm_det: buildItemDet(det, inv.isInterState)
    }));

    if (inv.isCreditDebit && inv.hasValidGstin) {
      inv.rawItems.forEach(item => {
        cdnrRows.push([
          inv.buyerGstin, inv.buyerName, inv.docType, inv.invNo, inv.invDate,
          inv.origInvNo, inv.origInvDate, round2(inv.totalInvValue), inv.posRaw, inv.rcm,
          item.rate + '%', round2(item.taxable), round2(item.igst), round2(item.cgst), round2(item.sgst)
        ]);
      });

      if (!cdnrCustomers[inv.buyerGstin]) cdnrCustomers[inv.buyerGstin] = { ctin: inv.buyerGstin, nt: [] };
      cdnrCustomers[inv.buyerGstin].nt.push({
        ntty: inv.docType === 'CRN' ? 'C' : 'D',
        nt_num: inv.invNo,
        nt_dt: inv.invDate,
        val: inv.totalInvValue,
        pos: inv.posDigits,
        rchrg: inv.rcm,
        inv_typ: inv.invType,
        itms: itmsList
      });
    } else if (inv.isCreditDebit && !inv.hasValidGstin && inv.isInterState && inv.totalInvValue > config.b2cThreshold) {
      cdnurList.push({
        typ: 'B2CL',
        ntty: inv.docType === 'CRN' ? 'C' : 'D',
        nt_num: inv.invNo,
        nt_dt: inv.invDate,
        val: inv.totalInvValue,
        pos: inv.posDigits,
        itms: itmsList
      });
    } else if (inv.hasValidGstin && !inv.isCreditDebit) {
      inv.rawItems.forEach(item => {
        b2bRows.push([
          inv.buyerGstin, inv.buyerName, inv.invNo, inv.invDate, round2(inv.totalInvValue),
          inv.posRaw, inv.rcm, inv.invType === 'R' ? 'Regular' : inv.invType, item.rate + '%',
          round2(item.taxable), round2(item.igst), round2(item.cgst), round2(item.sgst)
        ]);
      });

      if (!b2bCustomers[inv.buyerGstin]) b2bCustomers[inv.buyerGstin] = { ctin: inv.buyerGstin, inv: [] };
      b2bCustomers[inv.buyerGstin].inv.push({
        inum: inv.invNo,
        idt: inv.invDate,
        val: inv.totalInvValue,
        pos: inv.posDigits,
        rchrg: inv.rcm,
        inv_typ: inv.invType,
        itms: itmsList
      });
    } else if (!inv.isCreditDebit && inv.isInterState && inv.totalInvValue > config.b2cThreshold) {
      inv.rawItems.forEach(item => {
        b2clRows.push([
          inv.invNo, inv.invDate, round2(inv.totalInvValue), inv.posRaw,
          item.rate + '%', round2(item.taxable), round2(item.igst)
        ]);
      });

      if (!b2clPosMap[inv.posDigits]) b2clPosMap[inv.posDigits] = { pos: inv.posDigits, inv: [] };
      b2clPosMap[inv.posDigits].inv.push({
        inum: inv.invNo,
        idt: inv.invDate,
        val: inv.totalInvValue,
        itms: itmsList
      });
    } else {
      const sign = inv.docType === 'CRN' ? -1 : 1;
      inv.rawItems.forEach(item => {
        const sKey = `${inv.posRaw}_${item.rate}`;
        if (!b2csSheetMap[sKey]) {
          b2csSheetMap[sKey] = { type: 'OE', pos: inv.posRaw, rate: item.rate + '%', taxable: 0, igst: 0, cgst: 0, sgst: 0 };
        }
        b2csSheetMap[sKey].taxable = round2(b2csSheetMap[sKey].taxable + sign * item.taxable);
        b2csSheetMap[sKey].igst = round2(b2csSheetMap[sKey].igst + sign * item.igst);
        b2csSheetMap[sKey].cgst = round2(b2csSheetMap[sKey].cgst + sign * item.cgst);
        b2csSheetMap[sKey].sgst = round2(b2csSheetMap[sKey].sgst + sign * item.sgst);

        const jKey = `${inv.posDigits}_${inv.isInterState ? 'INTER' : 'INTRA'}_${item.rate}`;
        if (!b2csJsonMap[jKey]) {
          b2csJsonMap[jKey] = {
            sply_ty: inv.isInterState ? 'INTER' : 'INTRA',
            typ: 'OE',
            pos: inv.posDigits,
            rt: item.rate,
            txval: 0,
            iamt: 0,
            camt: 0,
            samt: 0,
            csamt: 0
          };
        }
        b2csJsonMap[jKey].txval = round2(b2csJsonMap[jKey].txval + sign * item.taxable);
        b2csJsonMap[jKey].iamt = round2(b2csJsonMap[jKey].iamt + sign * item.igst);
        b2csJsonMap[jKey].camt = round2(b2csJsonMap[jKey].camt + sign * item.cgst);
        b2csJsonMap[jKey].samt = round2(b2csJsonMap[jKey].samt + sign * item.sgst);
      });
    }
  });

  const b2csRows = Object.values(b2csSheetMap).map(i => [
    i.type, i.pos, i.rate, round2(i.taxable), round2(i.igst), round2(i.cgst), round2(i.sgst)
  ]);

  const b2csJson = Object.values(b2csJsonMap).map(item => {
    const out = { sply_ty: item.sply_ty, typ: item.typ, pos: item.pos, rt: item.rt, txval: item.txval };
    if (item.sply_ty === 'INTER') {
      out.iamt = item.iamt;
    } else {
      out.camt = item.camt;
      out.samt = item.samt;
    }
    out.csamt = 0;
    return out;
  });

  const hsnRows = Object.values(hsnSheetMap).map(h => [
    h.hsn, h.uqc, h.qty, round2(h.totalVal), round2(h.taxable), round2(h.igst), round2(h.cgst), round2(h.sgst)
  ]);

  // Bifurcated HSN items (Strictly includes mandatory 'val')
  const toHsnJsonList = (mapObj) => Object.values(mapObj).map((h, i) => ({
    num: i + 1,
    hsn_sc: h.hsn_sc,
    uqc: h.uqc,
    qty: h.qty,
    val: round2(h.val || (h.txval + h.iamt + h.camt + h.samt + h.csamt)),
    txval: h.txval,
    rt: h.rt,
    iamt: h.iamt,
    camt: h.camt,
    samt: h.samt,
    csamt: h.csamt
  }));

  // Unified HSN items (Standard portal format with 'val')
  const hsnUnifiedList = Object.values(hsnSheetMap).map((h, i) => ({
    num: i + 1,
    hsn_sc: h.hsn,
    uqc: h.uqc,
    qty: h.qty,
    val: round2(h.totalVal || (h.taxable + h.igst + h.cgst + h.sgst)),
    txval: h.taxable,
    rt: h.rate,
    iamt: h.igst,
    camt: h.cgst,
    samt: h.sgst,
    csamt: 0
  }));

  const docSummaryList = Object.values(docTracker)
    .filter(d => d.order.length > 0)
    .map(d => {
      const totalCount = d.order.length;
      const cancelledCount = Object.keys(d.cancelled).length;
      return {
        nature: d.nature,
        from: d.order[0],
        to: d.order[d.order.length - 1],
        totnum: totalCount,
        cancel: cancelledCount,
        net_issue: totalCount - cancelledCount
      };
    });

  const docsRows = docSummaryList.map(d => [d.nature, d.from, d.to, d.totnum, d.cancel, d.net_issue]);

  return {
    auditWarnings: auditWarnings,
    docSummaryList: docSummaryList,
    sheets: {
      b2bRows: b2bRows,
      b2clRows: b2clRows,
      b2csRows: b2csRows,
      hsnRows: hsnRows,
      cdnrRows: cdnrRows,
      docsRows: docsRows
    },
    json: {
      b2b: Object.values(b2bCustomers),
      b2cl: Object.values(b2clPosMap),
      b2cs: b2csJson,
      cdnr: Object.values(cdnrCustomers),
      cdnur: cdnurList,
      hsnB2b: toHsnJsonList(hsnB2bMap),
      hsnB2c: toHsnJsonList(hsnB2cMap),
      hsnUnified: hsnUnifiedList
    }
  };
}

function buildItemDet(det, isInterState) {
  const out = { rt: det.rt, txval: det.txval };
  if (isInterState) {
    out.iamt = det.iamt;
  } else {
    out.camt = det.camt;
    out.samt = det.samt;
  }
  out.csamt = 0;
  return out;
}

function buildDocIssuePayload(docsSheet, fallbackList, useV5) {
  let records = [];
  if (docsSheet && docsSheet.getLastRow() >= 2) {
    const rows = docsSheet.getDataRange().getValues().slice(1);
    records = rows
      .filter(r => Number(r[3]) > 0)
      .map(r => ({
        nature: String(r[0] || 'Invoices for outward supply').trim(),
        from: String(r[1] || '').trim(),
        to: String(r[2] || '').trim(),
        totnum: Number(r[3]) || 0,
        cancel: Number(r[4]) || 0,
        net_issue: Number(r[5]) || 0
      }));
  } else if (fallbackList && fallbackList.length > 0) {
    records = fallbackList;
  }

  if (records.length === 0) return null;

  const grouped = {};
  records.forEach(rec => {
    const docNumInt = DOC_NUM_MAP[rec.nature] || 1;
    const key = String(docNumInt);
    if (!grouped[key]) grouped[key] = { docNumInt: docNumInt, docs: [] };
    grouped[key].docs.push({
      num: grouped[key].docs.length + 1,
      from: rec.from,
      to: rec.to,
      totnum: rec.totnum,
      cancel: rec.cancel,
      net_issue: rec.net_issue
    });
  });

  const detList = Object.values(grouped).map(g => ({
    doc_num: useV5 ? String(g.docNumInt) : g.docNumInt,
    docs: g.docs
  }));

  return useV5 ? { doc_det_new: detList } : { doc_det: detList };
}

/**
 * 4. Manual Reset Utility for Next Client / Month
 */
function resetSheetForNewClient() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    '⚠️ RESET TOOL FOR NEW CLIENT',
    'This will wipe all invoice data in Invoices_Raw AND all summary tables to give you a fresh blank slate.\n\nHeaders and Config will be preserved.\n\nDo you want to continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const rawSheet = getRawSheet(ss);
  if (rawSheet && rawSheet.getLastRow() >= 2) {
    rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, Math.max(rawSheet.getLastColumn(), 1)).clearContent();
  }

  clearAllSummaryTabs(ss);
  ui.alert('✨ Fresh slate! All data wiped. Ready for your next client 💅🏻✌️');
}

function clearAllSummaryTabs(ss) {
  const summaryTabs = ['B2B_Summary', 'B2C_Large', 'B2C_Small', 'HSN_Summary', 'CDNR_Summary', 'Docs_Summary'];
  summaryTabs.forEach(tab => {
    const sh = ss.getSheetByName(tab);
    if (sh && sh.getLastRow() >= 2) {
      sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(sh.getLastColumn(), 1)).clearContent();
    }
  });
}

function getRawSheet(ss) {
  return ss.getSheetByName('Invoices_Raw') || ss.getSheetByName('Invoices_Ledger');
}

function validateGSTINChecksum(gstin) {
  if (!gstin || gstin.length !== 15) return false;
  const chars = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ';
  let total = 0;

  for (let i = 0; i < 14; i++) {
    const factor = (i % 2 === 0) ? 1 : 2;
    const codePoint = chars.indexOf(gstin.charAt(i));
    if (codePoint === -1) return false;
    const prod = factor * codePoint;
    total += Math.floor(prod / 36) + (prod % 36);
  }

  const remainder = total % 36;
  const checkCodePoint = (36 - remainder) % 36;
  return chars.charAt(checkCodePoint) === gstin.charAt(14);
}

function getConfigSettings(ss) {
  const cfg = { gstin: '33AXEPN4338D1ZR', fp: '102026', supplierState: '33', b2cThreshold: 100000 };
  const cfgSheet = ss.getSheetByName('Config');
  if (!cfgSheet) return cfg;

  const data = cfgSheet.getDataRange().getValues();
  data.forEach(row => {
    const label = String(row[0] || '').toLowerCase().trim();
    const val = row[1];
    if (label.includes('gstin')) cfg.gstin = String(val || '').trim().toUpperCase();
    if (label.includes('period') || label.includes('fp')) {
      if (val instanceof Date) {
        const m = String(val.getMonth() + 1).padStart(2, '0');
        cfg.fp = `${m}${val.getFullYear()}`;
      } else {
        const fpDigits = String(val || '').replace(/\D/g, '');
        cfg.fp = fpDigits.length === 5 ? fpDigits.padStart(6, '0') : String(val || '').trim();
      }
    }
    if (label.includes('state')) cfg.supplierState = String(val || '').trim().padStart(2, '0');
    if (label.includes('threshold')) {
      const cleaned = String(val || '').replace(/[^0-9.-]/g, '');
      cfg.b2cThreshold = Number(cleaned) || 100000;
    }
  });

  return cfg;
}

function parseTaxRate(rawRate) {
  if (typeof rawRate === 'string') {
    return parseFloat(rawRate.replace('%', '').trim()) || 0;
  }
  if (typeof rawRate === 'number' && !isNaN(rawRate)) {
    if (rawRate > 0 && rawRate < 1 && rawRate !== 0.1 && rawRate !== 0.25) {
      return round2(rawRate * 100);
    }
    return round2(rawRate);
  }
  return 0;
}

function parseAmount(val) {
  if (val == null) return 0;
  if (typeof val === 'number') return round2(val);
  const cleaned = String(val).replace(/[^0-9.-]/g, '').trim();
  return round2(Number(cleaned) || 0);
}

function cleanUQC(val) {
  const raw = String(val || 'OTH').trim().toUpperCase();
  const code = raw.split('-')[0].trim().replace(/[^A-Z]/g, '');
  return code || 'OTH';
}

function formatDate(d, ss) {
  if (!d) return '';

  if (d instanceof Date) {
    if (typeof Utilities !== 'undefined' && ss && ss.getSpreadsheetTimeZone) {
      return Utilities.formatDate(d, ss.getSpreadsheetTimeZone(), 'dd-MM-yyyy');
    }
    const day = String(d.getDate()).padStart(2, '0');
    const month = String(d.getMonth() + 1).padStart(2, '0');
    return `${day}-${month}-${d.getFullYear()}`;
  }

  const str = String(d).trim();
  const monthMap = {
    jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
    jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12'
  };

  const separator = str.includes('-') ? '-' : (str.includes('/') ? '/' : null);
  if (separator) {
    const parts = str.split(separator);
    if (parts.length === 3) {
      const p0 = parts[0].trim();
      const p1 = parts[1].trim();
      const p2 = parts[2].trim();

      if (p0.length === 4) {
        const month = monthMap[p1.toLowerCase().slice(0, 3)] || p1.padStart(2, '0');
        return `${p2.padStart(2, '0')}-${month}-${p0}`;
      }

      const day = p0.padStart(2, '0');
      const month = monthMap[p1.toLowerCase().slice(0, 3)] || p1.padStart(2, '0');
      const year = p2.length === 2 ? '20' + p2 : p2;
      return `${day}-${month}-${year}`;
    }
  }

  return str;
}

function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function round2(n) {
  return Math.round((Number(n) + Number.EPSILON) * 100) / 100;
}

function writeToSheet(ss, tabName, rows, startRow) {
  const sheet = ss.getSheetByName(tabName);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  if (lastRow >= startRow) {
    sheet.getRange(startRow, 1, lastRow - startRow + 1, lastCol).clearContent();
  }
  if (rows && rows.length > 0) {
    sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
  }
}
