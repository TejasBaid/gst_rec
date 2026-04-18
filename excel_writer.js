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
        } else {
            doc_date = entry.invoice_date;
        }
    }

    const e = entry;
    row.getCell(1).value = date_val;
    row.getCell(2).value = isTally ? e.party : e.name;
    row.getCell(3).value = e.gstin;
    row.getCell(4).value = isTally ? e.vch_type : "";
    row.getCell(5).value = isTally ? e.doc_no : e.invoice_no;
    row.getCell(6).value = doc_date;
    row.getCell(7).value = isTally ? e.invoice_amount : e.invoice_value;
    row.getCell(8).value = e.taxable;
    row.getCell(9).value = e.status;
    row.getCell(10).value = e.igst || null;
    row.getCell(11).value = e.cgst || null;
    row.getCell(12).value = e.sgst || null;

    row.eachCell((c, col) => {
        c.font = NORMAL;
        if ([1,6].includes(col)) c.numFmt = DATE_FMT;
        if ([7,8,10,11,12].includes(col)) c.numFmt = NUM_FMT;
        if (fill) c.fill = fill;
    });
}

function writeBreakup(ws, result, month, year, company, rcm_inputs) {
    const COL_HEADERS = ["Date", "Particulars", "Party GSTIN/UIN", "Vch Type", "Doc No.",
        "Doc date", "Invoice", "Taxable Amount", "STATUS", "IGST", "CGST", "SGST/UTGST"];
    
    setColWidths(ws, [12, 30, 18, 12, 22, 12, 12, 12, 16, 11, 11, 11]);

    const n = result.main_entries.length;
    const m = result.not_in_2b.length;
    const p = result.hold_entries.length;
    const q = result.portal_entries.length;
    const v = result.portal_excluded.length;

    const lastDay = new Date(year, ["January","February","March","April","May","June","July","August","September","October","November","December"].indexOf(month)+1, 0).getDate();
    const periodL = `1-${month.substring(0,3)}-${String(year).slice(-2)} to ${lastDay}`;

    ws.getCell('A1').value = company;
    ws.getCell('A1').font = TITLE_FONT;

    ws.getCell('A2').value = "GSTR-2A Reconciliation - Voucher Register";
    ws.getCell('A2').font = BOLD;
    ws.getCell('F2').value = "GST Registration:";
    ws.getCell('F2').font = BOLD;
    ws.getCell('I2').value = periodL;
    ws.getCell('I2').font = NORMAL;

    const rHeader = ws.getRow(3);
    COL_HEADERS.forEach((h, i) => {
        const c = rHeader.getCell(i+1);
        c.value = h; c.font = BOLD_SM; c.fill = HEADER_FILL;
        c.alignment = {horizontal:'center'}; c.border = THIN_BORDER;
    });

    let currentR = 4;
    const R_DATA_START = 4;
    result.main_entries.forEach(e => {
        const fill = isNotInGst2bStatus(e.status) ? NOTOK_FILL : null;
        writeEntryRow(ws, currentR++, e, true, fill);
    });
    const R_DATA_END = currentR - 1;

    currentR++; // blank
    const R_TOTAL = currentR++;
    ws.getCell(`A${R_TOTAL}`).value = "";
    const tc = ws.getCell(`I${R_TOTAL}`);
    tc.value = "TOTAL"; tc.font = BOLD; tc.fill = TOTAL_FILL;
    [7,8,10,11,12].forEach(col => {
        const c = ws.getCell(R_TOTAL, col);
        const l = ws.getColumn(col).letter;
        c.value = { formula: `SUM(${l}${R_DATA_START}:${l}${R_DATA_END})` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT;
    });

    currentR += 2;
    const R_LESS_SECTION = currentR++;
    ws.getCell(`A${R_LESS_SECTION}`).value = "Less : ITC not reflected in GSTR 2B";
    ws.getCell(`A${R_LESS_SECTION}`).font = BOLD;
    ws.getCell(`A${R_LESS_SECTION}`).fill = NOTOK_FILL;

    const R_LESS_HDR = currentR++;
    const rLessHdr = ws.getRow(R_LESS_HDR);
    COL_HEADERS.forEach((h, i) => {
        const c = rLessHdr.getCell(i+1);
        c.value = h; c.font = BOLD_SM; c.fill = NOTOK_FILL;
        c.alignment = {horizontal:'center'}; c.border = THIN_BORDER;
    });

    const R_LESS_START = currentR;
    result.not_in_2b.forEach(e => {
        writeEntryRow(ws, currentR++, e, true, NOTOK_FILL);
    });
    const R_LESS_END = Math.max(R_LESS_START, currentR - 1);
    const R_LESS_TOTAL = currentR++;

    if (m > 0) {
        [10,11,12].forEach(col => {
            const c = ws.getCell(R_LESS_TOTAL, col);
            const l = ws.getColumn(col).letter;
            c.value = { formula: `SUM(${l}${R_LESS_START}:${l}${R_LESS_END})` };
            c.font = BOLD; c.fill = NOTOK_FILL; c.numFmt = NUM_FMT;
        });
        ws.getCell(`H${R_LESS_TOTAL}`).value = { formula: `J${R_LESS_TOTAL}*100/9` };
        ws.getCell(`H${R_LESS_TOTAL}`).numFmt = NUM_FMT;
    }

    const R_NET1 = currentR++;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter;
        const c = ws.getCell(R_NET1, col);
        c.value = { formula: `${l}${R_TOTAL}-${l}${R_LESS_TOTAL}` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT;
    });

    let currentNetRow = R_NET1;

    // --- NEW CDNR SECTION ---
    if (result.cdnr_entries && result.cdnr_entries.length > 0) {
        currentR += 2;
        const R_CDNR_SECTION = currentR++;
        ws.getCell(`A${R_CDNR_SECTION}`).value = "Add/Sub : Debit & Credit Note from GSTR 2B";
        ws.getCell(`A${R_CDNR_SECTION}`).font = BOLD; ws.getCell(`A${R_CDNR_SECTION}`).fill = NOTOK_FILL;
        
        const rCdnrHdr = ws.getRow(currentR++);
        COL_HEADERS.forEach((h, i) => {
            const c = rCdnrHdr.getCell(i+1);
            c.value = h; c.font = BOLD_SM; c.fill = NOTOK_FILL; c.alignment = {horizontal:'center'}; c.border=THIN_BORDER;
        });

        const R_CDNR_START = currentR;
        result.cdnr_entries.forEach(e => {
            const isDebit = (e.note_type || "").toLowerCase().includes("debit");
            const mult = isDebit ? -1 : 1;
            
            const eAdjusted = { ...e };
            if (eAdjusted.invoice_value != null) eAdjusted.invoice_value = e.invoice_value * mult;
            if (eAdjusted.taxable != null) eAdjusted.taxable = e.taxable * mult;
            if (eAdjusted.igst != null) eAdjusted.igst = e.igst * mult;
            if (eAdjusted.cgst != null) eAdjusted.cgst = e.cgst * mult;
            if (eAdjusted.sgst != null) eAdjusted.sgst = e.sgst * mult;
            if (eAdjusted.cess != null) eAdjusted.cess = e.cess * mult;

            if (eAdjusted.note_no && !eAdjusted.invoice_no) eAdjusted.invoice_no = eAdjusted.note_no;
            if (eAdjusted.note_date && !eAdjusted.invoice_date) eAdjusted.invoice_date = eAdjusted.note_date;
            if (eAdjusted.note_value != null && eAdjusted.invoice_value == null) eAdjusted.invoice_value = eAdjusted.note_value * mult;

            writeEntryRow(ws, currentR++, eAdjusted, false, NOTOK_FILL); 
        });
        const R_CDNR_END = Math.max(R_CDNR_START, currentR - 1);
        const R_CDNR_TOTAL = currentR++;
        
        [10,11,12].forEach(col => {
            const l = ws.getColumn(col).letter;
            const c = ws.getCell(R_CDNR_TOTAL, col);
            c.value = { formula: `SUM(${l}${R_CDNR_START}:${l}${R_CDNR_END})` };
            c.font = BOLD; c.fill = NOTOK_FILL; c.numFmt = NUM_FMT;
        });

        const R_NET_CDNR = currentR++;
        [10,11,12].forEach(col => {
            const l = ws.getColumn(col).letter;
            const c = ws.getCell(R_NET_CDNR, col);
            c.value = { formula: `${l}${currentNetRow}+${l}${R_CDNR_TOTAL}` };
            c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT;
        });
        
        currentNetRow = R_NET_CDNR;
    }

    // Hold entries
    currentR += 2;
    const R_HOLD_SECTION = currentR++;
    ws.getCell(`A${R_HOLD_SECTION}`).value = "Add : ITC hold earlier now reflected in GSTR 2B";
    ws.getCell(`A${R_HOLD_SECTION}`).font = BOLD; ws.getCell(`A${R_HOLD_SECTION}`).fill = HOLD_FILL;
    
    const rHoldHdr = ws.getRow(currentR++);
    COL_HEADERS.forEach((h, i) => {
        const c = rHoldHdr.getCell(i+1);
        c.value = h; c.font = BOLD_SM; c.fill = HOLD_FILL; c.alignment = {horizontal:'center'}; c.border=THIN_BORDER;
    });

    const R_HOLD_START = currentR;
    result.hold_entries.forEach(e => { writeEntryRow(ws, currentR++, e, false, HOLD_FILL); });
    const R_HOLD_END = Math.max(R_HOLD_START, currentR - 1);
    
    const R_NET2 = currentR++;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter;
        const c = ws.getCell(R_NET2, col);
        c.value = { formula: `${l}${currentNetRow}+SUM(${l}${R_HOLD_START}:${l}${R_HOLD_END})` };
        c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT;
    });

    // Portal
    currentR += 2;
    const R_PORTAL_SECTION = currentR++;
    ws.getCell(`A${R_PORTAL_SECTION}`).value = "Add : ITC taken from GSTR 2B";
    ws.getCell(`A${R_PORTAL_SECTION}`).font = BOLD; ws.getCell(`A${R_PORTAL_SECTION}`).fill = PORTAL_FILL;

    const rPortalHdr = ws.getRow(currentR++);
    COL_HEADERS.forEach((h, i) => {
        const c = rPortalHdr.getCell(i+1); c.value=h; c.font=BOLD_SM; c.fill=PORTAL_FILL; c.alignment={horizontal:'center'}; c.border=THIN_BORDER;
    });
    const R_PORTAL_START = currentR;
    result.portal_entries.forEach(e => { writeEntryRow(ws, currentR++, e, false, PORTAL_FILL); });
    const R_PORTAL_END = Math.max(R_PORTAL_START, currentR - 1);
    const R_PORTAL_TOTAL = currentR++;
    if(q > 0) {
        [10,11,12].forEach(col => {
            const l = ws.getColumn(col).letter; const c = ws.getCell(R_PORTAL_TOTAL, col);
            c.value = {formula: `SUM(${l}${R_PORTAL_START}:${l}${R_PORTAL_END})`}; c.font=BOLD; c.fill=PORTAL_FILL; c.numFmt=NUM_FMT;
        });
    }

    const R_NET3 = currentR++;
    [10,11,12].forEach(col => {
        const l = ws.getColumn(col).letter; const c = ws.getCell(R_NET3, col);
        c.value = { formula: `${l}${R_NET2}+${l}${R_PORTAL_TOTAL}` }; c.font = BOLD; c.fill = TOTAL_FILL; c.numFmt = NUM_FMT;
    });

    // Reverted 
    if (v > 0) {
        currentR += 2;
        ws.getCell(`A${currentR}`).value = "Less : ITC in  from GSTR 2B but reverted";
        ws.getCell(`A${currentR}`).font = BOLD; ws.getCell(`A${currentR}`).fill = NOTOK_FILL; currentR++;
        const rH = ws.getRow(currentR++);
        COL_HEADERS.forEach((h, i) => {
            const c=rH.getCell(i+1); c.value=h; c.font=BOLD_SM; c.fill=NOTOK_FILL; c.border=THIN_BORDER; c.alignment={horizontal:'center'}
        });
        const R_REV_START = currentR;
        result.portal_excluded.forEach(e => { writeEntryRow(ws, currentR++, e, false, NOTOK_FILL); });
        const R_REV_END = currentR - 1;
        const R_REV_TOT = currentR++;
        [10,11,12].forEach(col => {
            const l = ws.getColumn(col).letter; const c = ws.getCell(R_REV_TOT, col);
            c.value = {formula:`SUM(${l}${R_REV_START}:${l}${R_REV_END})`}; c.font=BOLD; c.fill=NOTOK_FILL; c.numFmt=NUM_FMT;
        });
    }

    // RCM
    currentR += 2;
    ws.getCell(`A${currentR}`).value = `Add: RCM ITC of ${month} ${year}`;
    ws.getCell(`A${currentR}`).font=BOLD; ws.getCell(`A${currentR}`).fill=RCM_FILL; currentR++;
    const rcH = ws.getRow(currentR++);
    COL_HEADERS.forEach((h, i) => {
        const c=rcH.getCell(i+1); c.value=h; c.font=BOLD_SM; c.fill=RCM_FILL; c.border=THIN_BORDER; c.alignment={horizontal:'center'}
    });
    
    const R_RCM_START = currentR;
    const rcm_rows = [
        ["RCM on Lease Rent @18%", rcm_inputs.lease_rent||0, "CGST+SGST"],
        ["RCM on Office Rent @18%", rcm_inputs.office_rent||0, "CGST+SGST"],
        [`Freight Charges Outstation @5%`, rcm_inputs.freight_outstation||0, "IGST"],
        [`Freight Charges Local @2.5%`, rcm_inputs.freight_local||0, "CGST+SGST"]
    ];

    rcm_rows.forEach(([label, taxb, type]) => {
        const r = currentR++;
        ws.getCell(r, 1).value = label; ws.getCell(r, 1).font = NORMAL;
        ws.getCell(r, 7).value = taxb; ws.getCell(r, 7).numFmt = NUM_FMT;
        ws.getCell(r, 9).value = "RCM"; ws.getCell(r, 9).font = NORMAL;
        if(type === "IGST") {
            ws.getCell(r, 10).value = {formula: `ROUND(G${r}*5/100,0)`}; ws.getCell(r, 10).numFmt = NUM_FMT;
        } else {
            const rate = label.includes("18%") ? "9/100" : "2.5/100";
            ws.getCell(r, 11).value = {formula: `ROUND(G${r}*${rate},0)`}; ws.getCell(r, 11).numFmt = NUM_FMT;
            ws.getCell(r, 12).value = {formula: `K${r}`}; ws.getCell(r, 12).numFmt = NUM_FMT;
        }
        [7,10,11,12].forEach(c => ws.getCell(r, c).fill = RCM_FILL);
    });
    const R_RCM_END = currentR - 1;
    const R_RCM_TOTAL = currentR++;
    [7,10,11,12].forEach(col => {
        const l=ws.getColumn(col).letter; const c=ws.getCell(R_RCM_TOTAL, col);
        c.value={formula:`SUM(${l}${R_RCM_START}:${l}${R_RCM_END})`}; c.font=BOLD; c.fill=RCM_FILL; c.numFmt=NUM_FMT;
    });

    const R_FINAL_ITC = currentR++;
    ws.getCell(`A${R_FINAL_ITC}`).value = `Final ITC taken in GSTR 3B of ${month.toUpperCase()} ${year}`;
    ws.getCell(`A${R_FINAL_ITC}`).font = BOLD;
    [10,11,12].forEach(col => {
        const l=ws.getColumn(col).letter; const c=ws.getCell(R_FINAL_ITC, col);
        c.value={formula:`${l}${R_NET3}+${l}${R_RCM_TOTAL}`}; c.font=BOLD; c.fill=TOTAL_FILL; c.numFmt=NUM_FMT;
    });

    return {
        R_TOTAL, R_LESS_TOTAL, R_NET1, R_HOLD_START, R_HOLD_END, R_NET2, R_PORTAL_TOTAL, R_NET3,
        R_RCM_START, R_RCM_END, R_RCM_TOTAL, R_FINAL_ITC
    };
}

function writeCalculation(ws, result, month, year, company, rcm, mIn) {
    setColWidths(ws, [8, 55, 14, 14, 14, 14, 14]);

    const sumVal = (arr, k) => arr.reduce((a,b)=>a+(b[k]||0),0);
    const mTax = sumVal(result.main_entries, 'taxable'), mIgst = sumVal(result.main_entries, 'igst');
    const mCgst = sumVal(result.main_entries, 'cgst'), mSgst = sumVal(result.main_entries, 'sgst');
    const lIgst = sumVal(result.not_in_2b, 'igst'), lCgst = sumVal(result.not_in_2b, 'cgst'), lSgst = sumVal(result.not_in_2b, 'sgst');
    const hIgst = sumVal(result.hold_entries, 'igst'), hCgst = sumVal(result.hold_entries, 'cgst'), hSgst = sumVal(result.hold_entries, 'sgst');
    const pIgst = sumVal(result.portal_entries, 'igst'), pCgst = sumVal(result.portal_entries, 'cgst'), pSgst = sumVal(result.portal_entries, 'sgst');
    const cIgst = sumVal(result.cdnr_entries, 'igst'), cCgst = sumVal(result.cdnr_entries, 'cgst'), cSgst = sumVal(result.cdnr_entries, 'sgst');

    ws.mergeCells("B2:G2");
    const tC = ws.getCell("B2");
    tC.value = `${company} (GSTR 3B Calculation for ${month} ${year})`;
    tC.font = TITLE_FONT; tC.alignment = {horizontal:'center'};

    ws.getCell('B3').value = "Details"; ws.getCell('B3').font = BOLD;
    const hdrs = [[3,"Total Taxable"],[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]];
    hdrs.forEach(([c,lbl]) => {
        const cell=ws.getCell(3,c); cell.value=lbl; cell.font=BOLD; cell.fill=HEADER_FILL;
        cell.alignment={horizontal:'center'}; cell.border=THIN_BORDER;
    });

    const addRow = (r, pfx, desc, txb, i, c, s, bold=false, fill=null) => {
        const vals = [[1,pfx],[2,desc],[3,txb],[4,i],[5,c],[6,s]];
        vals.forEach(([col,val]) => {
            if(val == null) return;
            const cell = ws.getCell(r,col); cell.value=val; cell.font=bold?BOLD:NORMAL; cell.alignment={horizontal:'left'};
            if(fill) cell.fill=fill;
            if(col>=3) cell.numFmt=NUM_FMT;
        });
        const g = ws.getCell(r,7);
        g.value={formula:`SUM(D${r}:F${r})`}; g.font=bold?BOLD:NORMAL; g.numFmt=NUM_FMT;
        if(fill) g.fill=fill;
    };

    addRow(4, "", `All types of Purchases Entered in Books in ${month} ${year}`, mTax, mIgst, mCgst, mSgst);
    addRow(5, "Less : ", `ITC Which are Entered in books but not reflected in GTSR 2B of ${month} ${year}`, null, lIgst||null, lCgst||null, lSgst||null);
    addRow(6, "", "", null, {formula:`D4-D5`}, {formula:`E4-E5`}, {formula:`F4-F5`});
    addRow(7, "Add:", `Purchase Invoice which were on hold from previous months now reflected in GSTR 2B of ${month} ${year}`, null, hIgst||null, hCgst||null, hSgst||null);
    addRow(8, "", "", null, {formula:`D6+D7`}, {formula:`E6+E7`}, {formula:`F6+F7`});
    addRow(9, "Add:", `Purchase Invoice which are taken from Portal and invoices are to be collected`, null, pIgst||null, pCgst||null, pSgst||null);
    addRow(10, "", "", null, {formula:`D8+D9`}, {formula:`E8+E9`}, {formula:`F8+F9`});
    addRow(11, "Less : ", `Credit Note from Portal`, null, cIgst||null, cCgst||null, cSgst||null);
    addRow(12, "", `Eligible ITC in ${month} ${year} ( without RCM ITC)`, null, {formula:`D10-D11`}, {formula:`E10-E11`}, {formula:`F10-F11`}, true, TOTAL_FILL);

    ws.getCell('B13').value = "RCM ITC"; ws.getCell('B13').font = BOLD;
    addRow(14, "Add:", "RCM on Lease Rent  @18%", rcm.lease_rent||0, null, {formula:`C14*9/100`}, {formula:`E14`});
    addRow(15, "Add:", "RCM on Office Rent  @18%", rcm.office_rent||0, null, {formula:`C15*9/100`}, {formula:`E15`});
    addRow(16, "Add:", "RCM Outstation for 5 @5%", rcm.freight_outstation||0, {formula:`ROUND(C16*5/100,0)`}, null, null);
    addRow(17, "Add:", "RCM Local for  @2.5%", rcm.freight_local||0, null, {formula:`ROUND(C17*2.5/100,0)`}, {formula:`E17`});
    addRow(18, "", "", {formula:`SUM(C14:C17)`}, {formula:`SUM(D14:D17)`}, {formula:`SUM(E14:E17)`}, {formula:`SUM(F14:F17)`}, false, RCM_FILL);

    addRow(19, "", `Eligible ITC in ${month} ${year} (including RCM ITC)`, null, {formula:`D12+D18`}, {formula:`E12+E18`}, {formula:`F12+F18`}, true, TOTAL_FILL);

    ws.getCell('B22').value = "Credit Ledger Balance"; ws.getCell('B22').font = BOLD;
    [[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]].forEach(([c,l])=>{
        var cl=ws.getCell(22,c); cl.value=l; cl.font=BOLD; cl.fill=HEADER_FILL;
    });

    addRow(23, "", `OPENING RCM ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, null, mIn.opening_igst||null, mIn.opening_cgst||null, mIn.opening_sgst||null);
    addRow(24, "", `OPENING  ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, null, null, null, null);
    addRow(25, "", `CURRENT  ITC IN  THE MONTH OF ${month.toUpperCase()} ${year}`, null, {formula:`ROUND(D19,0)`}, {formula:`ROUND(E19,0)`}, {formula:`ROUND(F19,0)`});
    addRow(26, "", `TOTAL  ITC AVAILABLE IN CREDIT LEDGER in ${month.toUpperCase()} ${year}`, null, {formula:`SUM(D23:D25)`}, {formula:`SUM(E23:E25)`}, {formula:`SUM(F23:F25)`}, true, TOTAL_FILL);

    ws.getCell('B29').value = "Sales Details"; ws.getCell('B29').font = BOLD;
    [[3,"Total Taxable"],[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]].forEach(([c,l])=>{
        var cl=ws.getCell(29,c); cl.value=l; cl.font=BOLD; cl.fill=HEADER_FILL;
    });

    addRow(30, "", `SALES in ${month.toUpperCase()} ${year}`, mIn.sales_taxable||0, mIn.sales_igst||null, mIn.sales_cgst_sgst||null, {formula:`E30`});

    ws.getCell('B33').value = `TAX Set Off in GSTR 3B in ${month.toUpperCase()} ${year}`; ws.getCell('B33').font=BOLD;
    [[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]].forEach(([c,l])=>{
        var cl=ws.getCell(33,c); cl.value=l; cl.font=BOLD; cl.fill=HEADER_FILL;
    });

    addRow(34, "", "ALL OUTPUT GST LIABILITY", null, {formula:`ROUND(D30,0)`}, {formula:`ROUND(E30,0)`}, {formula:`ROUND(F30,0)`});
    addRow(35, "", "IGST , CGST AND SGST OUTPUT SET OFF WITH IGST ITC", null, {formula:`D34`}, {formula:`(D26-D35)/2`}, {formula:`D26-(D35+E35)`});
    addRow(36, "", "CGST OUTPUT SET OFF WITH CGST ITC", null, null, {formula:`E26`}, null);
    addRow(37, "", "SGST OUTPUT SET OFF WITH SGST ITC", null, null, null, {formula:`F26`});
    addRow(38, "", "GST PAYABLE", null, {formula:`D34-D35-D36-D37`}, {formula:`E34-E35-E36-E37`}, {formula:`F34-F35-F36-F37`}, true, NOTOK_FILL);
    addRow(39, "", "RCM PAYABLE", null, {formula:`D18`}, {formula:`E18`}, {formula:`F18`});
    addRow(40, "", `TOTAL GST PAYABLE in ${month.toUpperCase()} ${year} GSTR 3B`, null, {formula:`IF(D38<=0,D39,D38+D39)`}, {formula:`IF(E38<=0,E39,E38+E39)`}, {formula:`IF(F38<=0,F39,F38+F39)`}, true, TOTAL_FILL);

    ws.getCell('B43').value = "Credit Ledger Balance"; ws.getCell('B43').font = BOLD;
    [[4,"Integrated Tax"],[5,"Central Tax"],[6,"State Tax"],[7,"Total"]].forEach(([c,l])=>{
        var cl=ws.getCell(43,c); cl.value=l; cl.font=BOLD; cl.fill=HEADER_FILL;
    });

    addRow(44, "", `ITC LEFT IN CREDIT LEDGER in ${month.toUpperCase()} ${year}`, null, {formula:`IF(D38<=0,D38,"")`}, {formula:`IF(E38<=0,E38,"")`}, {formula:`IF(F38<=0,F38,"")`}, true, null);
}

async function generateFinal(result, month, year, company, rcm, manualIn) {
    const wb = new ExcelJS.Workbook();
    const wsB = wb.addWorksheet('BREAK UP');
    const wsC = wb.addWorksheet('CALCULATION');

    const rt = writeBreakup(wsB, result, month, year, company, rcm);
    writeCalculation(wsC, result, month, year, company, rcm, manualIn);

    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
}
