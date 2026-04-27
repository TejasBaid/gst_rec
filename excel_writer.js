// excel_writer.js

const BORDER_THIN = { style: 'thin' };
const THIN_BORDER = { top: BORDER_THIN, left: BORDER_THIN, bottom: BORDER_THIN, right: BORDER_THIN };

const HEADER_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
const TOTAL_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFE2EFDA' } };
const NOTOK_FILL  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFCE4D6' } };
const HOLD_FILL   = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDDEBF7' } };
const PORTAL_FILL = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFEBF1DE' } };
const RCM_FILL    = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F2F2' } };

const BOLD = { bold: true };
const BOLD_SM = { bold: true, size: 9 };
const NORMAL = { size: 9 };
const TITLE_FONT = { bold: true, size: 11 };

const DATE_FMT = "[$-en-IN]dd mmm yyyy";
const NUM_FMT = "#,##0.00";

function setColWidths(ws, widths) {
    widths.forEach((w, i) => { ws.getColumn(i+1).width = w; });
}

function writeEntryRow(ws, rowIdx, entry, isTally, fill) {
    const row = ws.getRow(rowIdx);
    
    let date_val, doc_date;
    if (isTally) {
        date_val = entry.date; doc_date = entry.doc_date;
    } else {
        date_val = null;
        if (entry.invoice_date && typeof entry.invoice_date === 'string') {
            const p = entry.invoice_date.split('/');
            if(p.length===3) doc_date = new Date(p[2], p[1]-1, p[0]);
            else doc_date = entry.invoice_date;
        } else if (entry.note_date && !entry.invoice_date) {
            const p = entry.note_date.split('/');
            if(p.length===3) doc_date = new Date(p[2], p[1]-1, p[0]);
            else doc_date = entry.note_date;
        } else {
            doc_date = entry.invoice_date;
        }
    }

    row.getCell(1).value = date_val;
    row.getCell(2).value = isTally ? entry.party : (entry.name || entry.party);
    row.getCell(3).value = entry.gstin;
    row.getCell(4).value = entry.vch_type || "";
    row.getCell(5).value = isTally ? entry.doc_no : (entry.invoice_no || entry.note_no);
    row.getCell(6).value = doc_date;
    row.getCell(7).value = isTally ? entry.invoice_amount : (entry.invoice_value || entry.note_value);
    row.getCell(8).value = entry.taxable;
    row.getCell(9).value = entry.status;
    row.getCell(10).value = entry.igst || null;
    row.getCell(11).value = entry.cgst || null;
    row.getCell(12).value = entry.sgst || null;

    row.eachCell({ includeEmpty: true }, (c, col) => {
        if (col > 12) return;
        c.font = NORMAL;
        if ([1,6].includes(col)) c.numFmt = DATE_FMT;
        if ([7,8,10,11,12].includes(col)) c.numFmt = NUM_FMT;
        if (fill) c.fill = fill;
        c.border = THIN_BORDER;
    });
}

function writeBreakup(ws, result, month, year, company, rcm_inputs) {
    const COL_HEADERS = ["Date", "Particulars", "Party GSTIN/UIN", "Vch Type", "Doc No.",
        "Doc date", "Invoice", "Taxable Amount", "STATUS", "IGST", "CGST", "SGST/UTGST"];
    
    setColWidths(ws, [12, 30, 18, 12, 22, 12, 12, 12, 16, 11, 11, 11]);

    ws.getCell('A1').value = company;
    ws.getCell('A1').font = TITLE_FONT;

    ws.getCell('A2').value = "GSTR-2A Reconciliation - Voucher Register";
    ws.getCell('A2').font = BOLD;
    ws.getCell('I2').value = `Period: ${month} ${year}`;
    ws.getCell('I2').font = NORMAL;

    const rHeader = ws.getRow(3);
    COL_HEADERS.forEach((h, i) => {
        const c = rHeader.getCell(i+1);
        c.value = h; c.font = BOLD_SM; c.fill = HEADER_FILL;
        c.alignment = {horizontal:'center'}; c.border = THIN_BORDER;
    });

    let currentR = 4;
    
    // 1. MAIN
    const R_DATA_START = currentR;
    if (result.main_entries && result.main_entries.length > 0) {
        result.main_entries.forEach(e => {
            const fill = isNotInGst2bStatus(e.status) ? NOTOK_FILL : null;
            writeEntryRow(ws, currentR++, e, true, fill);
        });
    }
    const R_DATA_END = Math.max(R_DATA_START, currentR - 1);
    const R_TOTAL_ROW = currentR++;
    ws.getCell(`I${R_TOTAL_ROW}`).value = "TOTAL"; 
    ws.getCell(`I${R_TOTAL_ROW}`).font = BOLD; ws.getCell(`I${R_TOTAL_ROW}`).fill = TOTAL_FILL;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_TOTAL_ROW, col);
        if (result.main_entries.length > 0) c.value = { formula: `SUM(${l}${R_DATA_START}:${l}${R_DATA_END})` };
        else c.value = 0;
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });

    let currentBalanceRow = R_TOTAL_ROW;

    // 2. Less: Not in 2B (Net 1)
    currentR += 2;
    ws.getCell(`A${currentR-1}`).value = "Less : ITC not reflected in GSTR 2B";
    ws.getCell(`A${currentR-1}`).font = BOLD;
    const LESS_HDR_ROW = currentR++;
    COL_HEADERS.forEach((h, i) => { const c = ws.getRow(LESS_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = NOTOK_FILL; c.border = THIN_BORDER; });
    const R_LESS_START = currentR;
    const hasLess = result.not_in_2b && result.not_in_2b.length > 0;
    if (hasLess) result.not_in_2b.forEach(e => writeEntryRow(ws, currentR++, e, true, NOTOK_FILL));
    const R_LESS_END = Math.max(R_LESS_START, currentR - 1);
    const R_NET1_ROW = currentR++;
    ws.getCell(`I${R_NET1_ROW}`).value = "Net 1"; ws.getCell(`I${R_NET1_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET1_ROW, col);
        if (hasLess) c.value = { formula: `${l}${currentBalanceRow}-SUM(${l}${R_LESS_START}:${l}${R_LESS_END})` };
        else c.value = { formula: `${l}${currentBalanceRow}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    currentBalanceRow = R_NET1_ROW;

    // 3. Add: Hold (Net 2)
    currentR += 2;
    ws.getCell(`A${currentR-1}`).value = "Add : ITC hold earlier now reflected in GSTR 2B";
    ws.getCell(`A${currentR-1}`).font = BOLD;
    const HOLD_HDR_ROW = currentR++;
    COL_HEADERS.forEach((h, i) => { const c = ws.getRow(HOLD_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = HOLD_FILL; c.border = THIN_BORDER; });
    const R_HOLD_START = currentR;
    const hasHold = result.hold_entries && result.hold_entries.length > 0;
    if (hasHold) result.hold_entries.forEach(e => writeEntryRow(ws, currentR++, e, false, HOLD_FILL));
    const R_HOLD_END = Math.max(R_HOLD_START, currentR - 1);
    const R_NET2_ROW = currentR++;
    ws.getCell(`I${R_NET2_ROW}`).value = "Net 2"; ws.getCell(`I${R_NET2_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET2_ROW, col);
        if (hasHold) c.value = { formula: `${l}${currentBalanceRow}+SUM(${l}${R_HOLD_START}:${l}${R_HOLD_END})` };
        else c.value = { formula: `${l}${currentBalanceRow}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    currentBalanceRow = R_NET2_ROW;

    // 4. Add: Portal (Net 3)
    currentR += 2;
    ws.getCell(`A${currentR-1}`).value = "Add : ITC taken from GSTR 2B";
    ws.getCell(`A${currentR-1}`).font = BOLD;
    const PORTAL_HDR_ROW = currentR++;
    COL_HEADERS.forEach((h, i) => { const c = ws.getRow(PORTAL_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = PORTAL_FILL; c.border = THIN_BORDER; });
    const R_PORTAL_START = currentR;
    const hasPortal = result.portal_entries && result.portal_entries.length > 0;
    if (hasPortal) result.portal_entries.forEach(e => writeEntryRow(ws, currentR++, e, false, PORTAL_FILL));
    const R_PORTAL_END = Math.max(R_PORTAL_START, currentR - 1);
    const R_NET3_ROW = currentR++;
    ws.getCell(`I${R_NET3_ROW}`).value = "Net 3"; ws.getCell(`I${R_NET3_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET3_ROW, col);
        if (hasPortal) c.value = { formula: `${l}${currentBalanceRow}+SUM(${l}${R_PORTAL_START}:${l}${R_PORTAL_END})` };
        else c.value = { formula: `${l}${currentBalanceRow}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    currentBalanceRow = R_NET3_ROW;

    // 5. Add/Less: CDNR
    currentR += 2;
    ws.getCell(`A${currentR-1}`).value = "Add/Less  : ITC taken from GSTR  2B CDNRA";
    ws.getCell(`A${currentR-1}`).font = BOLD;
    const CDNR_HDR_ROW = currentR++;
    COL_HEADERS.forEach((h, i) => { const c = ws.getRow(CDNR_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = HOLD_FILL; c.border = THIN_BORDER; });
    const R_CDNR_START = currentR;
    const hasCdnr = result.cdnr_entries && result.cdnr_entries.length > 0;
    if (hasCdnr) {
        result.cdnr_entries.forEach(e => {
            const isCredit = (e.note_type || "").toLowerCase().includes("credit");
            const mult = isCredit ? -1 : 1;
            const eAdj = {...e};
            ['igst','cgst','sgst','cess'].forEach(k => { if(eAdj[k]!=null) eAdj[k]*=mult; });
            eAdj.vch_type = e.note_type || "";
            if(!eAdj.invoice_no) eAdj.invoice_no = e.note_no;
            if(!eAdj.invoice_date) eAdj.invoice_date = e.note_date;
            writeEntryRow(ws, currentR++, eAdj, false, HOLD_FILL);
        });
    }
    const R_CDNR_END = Math.max(R_CDNR_START, currentR - 1);
    const R_NET_CDNR_ROW = currentR++;
    ws.getCell(`I${R_NET_CDNR_ROW}`).value = "Net (After CDNR)"; ws.getCell(`I${R_NET_CDNR_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET_CDNR_ROW, col);
        if (hasCdnr) c.value = { formula: `${l}${currentBalanceRow}+SUM(${l}${R_CDNR_START}:${l}${R_CDNR_END})` };
        else c.value = { formula: `${l}${currentBalanceRow}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    currentBalanceRow = R_NET_CDNR_ROW;

    // 6. Less: Reverted
    currentR += 2;
    ws.getCell(`A${currentR-1}`).value = "Less : ITC in from GSTR 2B but reverted";
    ws.getCell(`A${currentR-1}`).font = BOLD;
    const REV_HDR_ROW = currentR++;
    COL_HEADERS.forEach((h, i) => { const c = ws.getRow(REV_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = NOTOK_FILL; c.border = THIN_BORDER; });
    const R_REV_START = currentR;
    const hasRev = result.portal_excluded && result.portal_excluded.length > 0;
    if (hasRev) result.portal_excluded.forEach(e => writeEntryRow(ws, currentR++, e, false, NOTOK_FILL));
    const R_REV_END = Math.max(R_REV_START, currentR - 1);
    const R_NET_REV_ROW = currentR++;
    ws.getCell(`I${R_NET_REV_ROW}`).value = "Net (After Reversion)"; ws.getCell(`I${R_NET_REV_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET_REV_ROW, col);
        if (hasRev) c.value = { formula: `${l}${currentBalanceRow}-SUM(${l}${R_REV_START}:${l}${R_REV_END})` };
        else c.value = { formula: `${l}${currentBalanceRow}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    currentBalanceRow = R_NET_REV_ROW;

    // RCM & Final
    currentR += 2;
    ws.getCell(`A${currentR}`).value = `Add: RCM ITC of ${month} ${year}`; ws.getCell(`A${currentR}`).font = BOLD; currentR++;
    const RCM_HDR_ROW = currentR++;
    const rcmHdrs = ["Particulars", "Party GSTIN/UIN", "Vch Type", "Doc No.", "Doc Date", "Invoice", "Taxable Value", "STATUS", "IGST", "CGST", "SGST/UTGST"];
    rcmHdrs.forEach((h, i) => { const c = ws.getRow(RCM_HDR_ROW).getCell(i+1); c.value = h; c.font = BOLD_SM; c.fill = RCM_FILL; c.border = THIN_BORDER; });
    const R_RCM_START = currentR;
    const rRows = [["RCM on Lease Rent @18%", rcm_inputs.lease_rent||0, "CS"], ["RCM on Office Rent @18%", rcm_inputs.office_rent||0, "CS"], ["Freight Charges Outstation @5%", rcm_inputs.freight_outstation||0, "I"], ["Freight Charges Local @2.5%", rcm_inputs.freight_local||0, "CS"]];
    rRows.forEach(([lbl, tx, mode]) => {
        const r = currentR++;
        ws.getCell(r,1).value = lbl; ws.getCell(r,7).value = tx; ws.getCell(r,9).value = "RCM";
        if(mode === "I") ws.getCell(r,10).value = { formula: `ROUND(G${r}*5/100,0)` };
        else { const rate = lbl.includes("18%") ? "9/100" : "2.5/100"; ws.getCell(r,11).value = { formula: `ROUND(G${r}*${rate},0)` }; ws.getCell(r,12).value = { formula: `K${r}` }; }
        ws.getRow(r).eachCell({includeEmpty:true}, (c,col)=>{ if(col<=12){ c.fill=RCM_FILL; c.border=THIN_BORDER; if(col>=7) c.numFmt=NUM_FMT; } });
    });
    const R_RCM_END = currentR - 1;
    const R_RCM_SUM_ROW = currentR++;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_RCM_SUM_ROW, col);
        c.value = { formula: `SUM(${l}${R_RCM_START}:${l}${R_RCM_END})` };
        c.font = BOLD; c.fill = RCM_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });
    const R_FINAL_ROW = currentR++;
    ws.getCell(`A${R_FINAL_ROW}`).value = `Final ITC taken in GSTR 3B of ${month.toUpperCase()} ${year}`; ws.getCell(`A${R_FINAL_ROW}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_FINAL_ROW, col);
        c.value = { formula: `${l}${currentBalanceRow}+${l}${R_RCM_SUM_ROW}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT; c.border = THIN_BORDER;
    });

    return { R_TOTAL: R_TOTAL_ROW, R_NET_CDNR: R_NET_CDNR_ROW, R_NET_LESS: R_NET1_ROW, R_NET_HOLD: R_NET2_ROW, R_NET_PORTAL: R_NET3_ROW, R_NET_REV: R_NET_REV_ROW, R_RCM_SUM: R_RCM_SUM_ROW, R_FINAL: R_FINAL_ROW };
}

function writeCalculation(ws, result, month, year, company, rcm, mIn, rowIdxs) {
    ws.mergeCells("B2:G2");
    const tC = ws.getCell("B2");
    tC.value = `${company} (GSTR 3B Calculation for ${month} ${year} )`;
    tC.font = TITLE_FONT; tC.alignment = {horizontal:'center'};

    ws.getCell('B3').value = "Details"; ws.getCell('B3').font = BOLD;
    const hdrs = [[3,"Total Taxable"],[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]];
    hdrs.forEach(([col,lbl]) => {
        const cell=ws.getCell(3,col); cell.value=lbl; cell.font=BOLD; cell.fill=HEADER_FILL;
        cell.alignment={horizontal:'center'}; cell.border=THIN_BORDER;
    });

    const addRow = (r, pfx, desc, iVal, bold=false, fill=null) => {
        if(pfx) ws.getCell(r,1).value = pfx;
        ws.getCell(r,2).value = desc;
        [3,4,5,6].forEach((col, idx) => {
            const v = iVal[idx];
            if(v === undefined || v === null) return;
            const cell = ws.getCell(r,col);
            if (typeof v === 'string' && v.startsWith('=')) cell.value = { formula: v.substring(1) };
            else if (v && typeof v === 'object' && v.formula) cell.value = v;
            else cell.value = v;
            cell.font = bold ? BOLD : NORMAL;
            if(fill) cell.fill = fill;
            if(col >= 3) cell.numFmt = NUM_FMT;
            cell.border = THIN_BORDER;
        });
        const g = ws.getCell(r,7);
        g.value = { formula: `SUM(D${r}:F${r})` };
        g.font = bold ? BOLD : NORMAL; g.numFmt = NUM_FMT;
        if(fill) g.fill = fill; g.border = THIN_BORDER;
    };

    const b = "'BREAK UP'!";
    // Row 4: Total Purchases (Books)
    addRow(4, "", `All types of Purchases Entered in Books in  ${month} ${year}`, [
        {formula: `${b}H${rowIdxs.R_TOTAL}`}, 
        {formula: `${b}J${rowIdxs.R_TOTAL}`}, 
        {formula: `${b}K${rowIdxs.R_TOTAL}`}, 
        {formula: `${b}L${rowIdxs.R_TOTAL}`}
    ]);

    // Row 5: Less Not in 2B (Ref Row 5)
    addRow(5, "Less : ", `ITC Which are Entered in books but not reflected in GTSR 2B of ${month} ${year}`, [
        null,
        {formula: `${b}J${rowIdxs.R_NET_LESS}-${b}J${rowIdxs.R_TOTAL}`},
        {formula: `${b}K${rowIdxs.R_NET_LESS}-${b}K${rowIdxs.R_TOTAL}`},
        {formula: `${b}L${rowIdxs.R_NET_LESS}-${b}L${rowIdxs.R_TOTAL}`}
    ]);

    // Row 6: Subtotal
    addRow(6, "", "", [null, {formula: "D4+D5"}, {formula: "E4+E5"}, {formula: "F4+F5"}]);

    // Row 7: Add Hold
    addRow(7, "Add:", `Purchase Invoice which were on hold from previous months now reflected in GSTR 2B of ${month} ${year}`, [
        null,
        {formula: `${b}J${rowIdxs.R_NET_HOLD}-${b}J${rowIdxs.R_NET_LESS}`},
        {formula: `${b}K${rowIdxs.R_NET_HOLD}-${b}K${rowIdxs.R_NET_LESS}`},
        {formula: `${b}L${rowIdxs.R_NET_HOLD}-${b}L${rowIdxs.R_NET_LESS}`}
    ]);

    // Row 8: Net 1
    addRow(8, "", "", [null, {formula: "D6+D7"}, {formula: "E6+E7"}, {formula: "F6+F7"}]);

    // Row 9: Add Portal (Ref Row 9)
    addRow(9, "Add:", `Purchase Invoice which are taken from Portal and invoices are to be collected`, [
        null,
        {formula: `${b}J${rowIdxs.R_NET_PORTAL}-${b}J${rowIdxs.R_NET_HOLD}`},
        {formula: `${b}K${rowIdxs.R_NET_PORTAL}-${b}K${rowIdxs.R_NET_HOLD}`},
        {formula: `${b}L${rowIdxs.R_NET_PORTAL}-${b}L${rowIdxs.R_NET_HOLD}`}
    ]);

    // Row 10: Net 2
    addRow(10, "", "", [null, {formula: "D8+D9"}, {formula: "E8+E9"}, {formula: "F8+F9"}]);

    // Row 11: Credit Note (CDNR) (Ref Row 11)
    addRow(11, "Adjust ", "Credit Note from Portal", [
        null,
        {formula: `${b}J${rowIdxs.R_NET_CDNR}-${b}J${rowIdxs.R_NET_PORTAL}`},
        {formula: `${b}K${rowIdxs.R_NET_CDNR}-${b}K${rowIdxs.R_NET_PORTAL}`},
        {formula: `${b}L${rowIdxs.R_NET_CDNR}-${b}L${rowIdxs.R_NET_PORTAL}`}
    ]);

    // Row 12: Less Reverted
    addRow(12, "Less : ", "ITC in from GSTR 2B but reverted", [
        null,
        {formula: `${b}J${rowIdxs.R_NET_REV}-${b}J${rowIdxs.R_NET_CDNR}`},
        {formula: `${b}K${rowIdxs.R_NET_REV}-${b}K${rowIdxs.R_NET_CDNR}`},
        {formula: `${b}L${rowIdxs.R_NET_REV}-${b}L${rowIdxs.R_NET_CDNR}`}
    ]);

    // Row 13: Eligible ITC (without RCM)
    addRow(13, "", `Eligible ITC in ${month} ${year} ( without RCM ITC)`, [
        null,
        {formula: "D10+D11+D12"}, {formula: "E10+E11+E12"}, {formula: "F10+F11+F12"}
    ], true, TOTAL_FILL);

    ws.getCell('B14').value = "RCM ITC"; ws.getCell('B14').font = BOLD;
    addRow(15, "Add:", "RCM on Lease Rent  @18%", [rcm.lease_rent||0, null, {formula:"C15*9/100"}, {formula:"E15"}]);
    addRow(16, "Add:", "RCM on Office Rent  @18%", [rcm.office_rent||0, null, {formula:"C16*9/100"}, {formula:"E16"}]);
    addRow(17, "Add:", "RCM Outstation for 5 @5%", [rcm.freight_outstation||0, {formula:"ROUND(C17*5/100,0)"}, null, null]);
    addRow(18, "Add:", "RCM Local for  @2.5%", [rcm.freight_local||0, null, {formula:"ROUND(C18*2.5/100,0)"}, {formula:"E18"}]);
    addRow(19, "", "", [{formula: "SUM(C15:C18)"}, {formula: "SUM(D15:D18)"}, {formula: "SUM(E15:E18)"}, {formula: "SUM(F15:F18)"}], false, RCM_FILL);

    addRow(20, "", `Eligible ITC in ${month} ${year} (including RCM ITC)`, [null, {formula: "D13+D19"}, {formula: "E13+E19"}, {formula: "F13+F19"}], true, TOTAL_FILL);

    ws.getCell('B22').value = "Credit Ledger Balance"; ws.getCell('B22').font = BOLD;
    addRow(23, "", `OPENING RCM ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, [null, mIn.opening_igst||0, mIn.opening_cgst||0, mIn.opening_sgst||0]);
    addRow(24, "", `OPENING  ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, [null, 0, 0, 0]);
    addRow(25, "", `CURRENT  ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, [null, {formula: "ROUND(D20,0)"}, {formula: "ROUND(E20,0)"}, {formula: "ROUND(F20,0)"}]);
    addRow(26, "", `TOTAL  ITC AVAILABLE IN CREDIT LEDGER in ${month.toUpperCase()} ${year}`, [null, {formula: "SUM(D23:D25)"}, {formula: "SUM(E23:E25)"}, {formula: "SUM(F23:F25)"}], true, TOTAL_FILL);

    ws.getCell('B29').value = "Sales Details"; ws.getCell('B29').font = BOLD;
    addRow(30, "", `SALES in ${month.toUpperCase()} ${year}`, [mIn.sales_taxable||0, mIn.sales_igst||0, mIn.sales_cgst_sgst||0, {formula: "E30"}]);
    ws.getCell('B33').value = "TAX Set Off Detail"; ws.getCell('B33').font=BOLD;
    addRow(34, "", "ALL OUTPUT GST LIABILITY", [null, {formula: "ROUND(D30,0)"}, {formula: "ROUND(E30,0)"}, {formula: "ROUND(F30,0)"}]);
    addRow(35, "", "IGST , CGST AND SGST OUTPUT SET OFF WITH IGST ITC", [null, {formula: "D34"}, {formula: "(MIN(D26-D35, E34))/2"}, {formula: "MIN(D26-D35-E35, F34)"}]);
    addRow(36, "", "CGST OUTPUT SET OFF WITH CGST ITC", [null, null, {formula: "MIN(E26, E34-E35)"}, null]);
    addRow(37, "", "SGST OUTPUT SET OFF WITH SGST ITC", [null, null, null, {formula: "MIN(F26, F34-F35)"}]);
    addRow(38, "", "GST PAYABLE", [null, {formula: "D34-D35"}, {formula: "E34-E35-E36"}, {formula: "F34-F35-F37"}], true, NOTOK_FILL);
    addRow(39, "", "RCM PAYABLE", [null, {formula: "D19"}, {formula: "E19"}, {formula: "F19"}]);
    addRow(40, "", `TOTAL GST PAYABLE in ${month.toUpperCase()} ${year} GSTR 3B`, [null, {formula: "IF(D38<=0,D39,D38+D39)"}, {formula: "IF(E38<=0,E39,E38+E39)"}, {formula: "IF(F38<=0,F39,F38+F39)"}], true, TOTAL_FILL);
    addRow(44, "", `ITC LEFT IN CREDIT LEDGER in ${month.toUpperCase()} ${year}`, [null, {formula: "IF(D38<=0,D38,\"\")"}, {formula: "IF(E38<=0,E38,\"\")"}, {formula: "IF(F38<=0,F38,\"\")"}], true, null);
}

async function generateFinal(result, month, year, company, rcm, manualIn) {
    const wb = new ExcelJS.Workbook();
    const wsB = wb.addWorksheet('BREAK UP');
    const wsC = wb.addWorksheet('CALCULATION');
    const rowIdxs = writeBreakup(wsB, result, month, year, company, rcm);
    writeCalculation(wsC, result, month, year, company, rcm, manualIn, rowIdxs);
    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
}
