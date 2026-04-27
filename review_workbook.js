// review_workbook.js

const DECISION_HEADERS = [
    "RowId", "Include?", "Bucket", "Status", "BookDate", "PartyOrName", "GSTIN",
    "VchType", "VchNo", "DocOrInvoiceNo", "DocOrInvoiceDate", "InvoiceValue",
    "Taxable", "IGST", "CGST", "SGST", "Cess", "TaxAmount", "InvoiceType",
    "PlaceOfSupply", "ReverseCharge", "GSTR1Period", "FilingDate", "ITCAvailability",
    "CDNR_NoteType", "CDNR_SupplyType", "CDNR_NoteValue"
];

function includeForReviewRow(bucket, tallyStatus = "", gst = null, sheetStatus = "") {
    const b = (bucket || "").toUpperCase();
    if (b === "RCM_BOOK" || b === "HOLD") return "No";
    if (b === "MAIN") {
        if (isNotInGst2bStatus(tallyStatus)) return "No";
        const st = (tallyStatus || "").trim().toUpperCase();
        if (isIneligible(tallyStatus)) return "No";
        if (st.includes("INELIGIBLE") || st.includes("CAR EXP") || (st.includes("FOOD") && !st.includes("JAN"))) return "No";
        return "Yes";
    }
    if (b === "PORTAL") {
        if (gst && isRcm(gst)) return "No";
        const ss = (sheetStatus || "").trim().toUpperCase();
        if (isIneligible(sheetStatus)) return "No";
        if (ss.includes("INELIGIBLE") || ss.includes("CAR EXP") || (ss.includes("FOOD") && !ss.includes("JAN"))) return "No";
        return "Yes";
    }
    if (b === "CDNR") return "Yes";
    return "Yes";
}

function portalSheetStatus(g) {
    if (isRcm(g)) return (g.status || "").trim() || "RCM";
    return "In GSTR 2B";
}

function mainSheetStatus(t) {
    if (isNotInGst2bStatus(t.status)) return "Not in GST 2B";
    if (isIneligible(t.status)) return (t.status || "").trim();
    return "Ok";
}

function parseInclude(val) {
    if (val == null || String(val).trim() === "") return true;
    const s = String(val).trim().toUpperCase();
    if (["YES", "Y", "1", "TRUE", "INCLUDE"].includes(s)) return true;
    if (["NO", "N", "0", "FALSE", "EXCLUDE"].includes(s)) return false;
    return true;
}

function statusNeedsBreakupRouting(status) {
    const s = (status || "").trim().toLowerCase();
    return s === "not in books" || s === "in gstr 2b";
}

function statusIsRcmExcluded(status) {
    const s = (status || "").trim().toUpperCase();
    if (!s) return false;
    if (s === "RCM") return true;
    return s.startsWith("RCM");
}

async function exportReview(result, month, year, company) {
    const wb = new ExcelJS.Workbook();
    const wsD = wb.addWorksheet('Decisions');
    const wsM = wb.addWorksheet('Meta');
    const wsI = wb.addWorksheet('Instructions');
    
    // Meta
    const periodTag = currentPeriodTag(month, year);
    wsM.addRow(['Month', month]).font = {bold:true};
    wsM.addRow(['Year', year]).font = {bold:true};
    wsM.addRow(['Company', company]).font = {bold:true};
    wsM.addRow(['PeriodTag', periodTag]).font = {bold:true};
    wsM.addRow(['FormatVersion', '1']).font = {bold:true};

    // Instructions
    wsI.getCell('A1').value = "GST Reconciliation — Review workbook (Step 1 of 2)\n\n" +
      "1. Check the 'Status' column — this is what the tool computed from your Tally and GST files.\n" +
      "2. In the 'Include?' column, type Yes or No for each row.\n" +
      "   Only rows with Yes are used when you generate Final.xlsx (Step 2).\n" +
      "3. Save this file (Ctrl+S).\n" +
      "4. In the app, use Step 2 — select this saved file and click Generate Final.xlsx.\n\n" +
      "Include? accepts: Yes, Y, 1, true (include) or No, N, 0, false (exclude).\n" +
      "Do not rename sheets or the 'RowId' / 'Bucket' columns.\n" +
      "Buckets: MAIN = Tally purchases; Status 'Not in GST 2B' = in books but not in " +
      "GSTR-2B B2B; HOLD / PORTAL = GST lines (Step 2 asks Hold / Reverted / Portal when " +
      "Status is 'In GSTR 2B' or 'Not in Books' and Include?=Yes; other Status skips that); " +
      "CDNR = credit/debit notes (exported with Include?=No and Status 'Not in Books' " +
      "by default; set Include?=Yes to include in Final); RCM_BOOK = reverse-charge matched in books.\n";
    wsI.getCell('A1').alignment = { wrapText: true, vertical: 'top' };
    wsI.getColumn(1).width = 85;

    // Decisions Headings
    const hr = wsD.addRow(DECISION_HEADERS);
    hr.eachCell(c => {
        c.font = { bold: true };
        c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9E1F2' } };
    });

    let rid = 0;
    const addRow = (rowVals) => {
        const r = wsD.addRow(rowVals);
        // format numbers / dates
        r.getCell(5).numFmt = '[$-en-IN]dd mmm yyyy';
        r.getCell(11).numFmt = '[$-en-IN]dd mmm yyyy';
        for(let i=12; i<=27; i++){
            if(typeof rowVals[i-1] === 'number') r.getCell(i).numFmt = '#,##0.00';
        }
    };

    function parseInvDtToExcel(s) {
        if (!s) return null;
        if (s instanceof Date) return s;
        // expect dd/mm/yyyy
        const parts = s.split(/[-/]/);
        if (parts.length === 3) {
            return new Date(parts[2], parts[1]-1, parts[0]);
        }
        return s;
    }

    result.main_entries.forEach(t => {
        rid++;
        const st = mainSheetStatus(t);
        const inc = includeForReviewRow("MAIN", t.status);
        addRow([rid, inc, "MAIN", st, t.date, t.party, t.gstin, t.vch_type, t.vch_no,
            t.doc_no, t.doc_date, t.invoice_amount, t.taxable, t.igst, t.cgst, t.sgst,
            t.cess, t.tax_amount, "", "", "", "", "", "", "", "", ""]);
    });

    result.rcm_tally.forEach(t => {
        rid++;
        const inc = includeForReviewRow("RCM_BOOK", t.status);
        addRow([rid, inc, "RCM_BOOK", t.status, t.date, t.party, t.gstin, t.vch_type, t.vch_no,
            t.doc_no, t.doc_date, t.invoice_amount, t.taxable, t.igst, t.cgst, t.sgst,
            t.cess, t.tax_amount, "", "", "", "", "", "", "", "", ""]);
    });

    result.hold_entries.forEach(g => {
        rid++;
        const inc = includeForReviewRow("HOLD", "", g, g.status);
        addRow([rid, inc, "HOLD", g.status, null, g.name, g.gstin, "", "",
            g.invoice_no, parseInvDtToExcel(g.invoice_date), g.invoice_value, g.taxable,
            g.igst, g.cgst, g.sgst, g.cess, "", g.invoice_type, g.place_of_supply,
            g.reverse_charge, g.period, g.filing_date, g.itc_availability, "", "", ""]);
    });

    result.portal_entries.forEach(g => {
        rid++;
        const st = portalSheetStatus(g);
        const inc = includeForReviewRow("PORTAL", "", g, st);
        addRow([rid, inc, "PORTAL", st, null, g.name, g.gstin, "", "",
            g.invoice_no, parseInvDtToExcel(g.invoice_date), g.invoice_value, g.taxable,
            g.igst, g.cgst, g.sgst, g.cess, "", g.invoice_type, g.place_of_supply,
            g.reverse_charge, g.period, g.filing_date, g.itc_availability, "", "", ""]);
    });

    result.cdnr_entries.forEach(c => {
        rid++;
        addRow([rid, "Yes", "CDNR", "In GSTR 2B", null, c.name, c.gstin, "", "",
            c.note_no, parseInvDtToExcel(c.note_date), c.note_value, c.taxable,
            c.igst, c.cgst, c.sgst, c.cess, "", "", c.place_of_supply, c.reverse_charge,
            "", "", c.itc_availability, c.note_type, c.note_supply_type, c.note_value]);
    });

    wsD.views = [{state: 'frozen', xSplit: 6, ySplit: 1, topLeftCell: 'G2'}];
    wsD.autoFilter = `A1:${wsD.getColumn(DECISION_HEADERS.length).letter}${wsD.rowCount}`;

    const hidden = new Set(["RowId", "VchType", "VchNo", "InvoiceType", "PlaceOfSupply",
        "ReverseCharge", "GSTR1Period", "FilingDate", "ITCAvailability", "CDNR_NoteType",
        "CDNR_SupplyType", "CDNR_NoteValue"]);

    DECISION_HEADERS.forEach((h, i) => {
        const col = wsD.getColumn(i+1);
        if (hidden.has(h)) {
            col.hidden = true; col.width = 8;
        } else {
            col.width = 14;
        }
    });
    wsD.getColumn(6).width = 28; // PartyOrName

    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
}

async function loadReview(arrayBuffer) {
    const wb = new ExcelJS.Workbook();
    await wb.xlsx.load(arrayBuffer);
    
    if (!wb.getWorksheet('Meta')) throw new Error("Missing 'Meta' sheet.");
    if (!wb.getWorksheet('Decisions')) throw new Error("Missing 'Decisions' sheet.");

    const wsM = wb.getWorksheet('Meta');
    const metaMap = {};
    wsM.eachRow((r, n) => {
        if(n > 20) return;
        const vals = r.values.slice(1);
        if(vals[0]) metaMap[valStr(vals[0])] = vals[1];
    });

    const snap = {
        month: valStr(metaMap["Month"]) || "January",
        year: parseInt(metaMap["Year"], 10) || new Date().getFullYear(),
        company: valStr(metaMap["Company"]) || "Company",
        period_tag: valStr(metaMap["PeriodTag"]),
        rows: []
    };

    const wsD = wb.getWorksheet('Decisions');
    const nameToCol = {};
    const headers = wsD.getRow(1).values.slice(1);
    headers.forEach((h, i) => { if(h) nameToCol[valStr(h).trim()] = i+1; });

    ['RowId', 'Include?', 'Bucket'].forEach(req => {
        if(!nameToCol[req]) throw new Error(`Missing required column: ${req}`);
    });

    for(let r=2; r<=wsD.rowCount; r++) {
        const row = wsD.getRow(r);
        const d = {};
        DECISION_HEADERS.forEach(name => {
            if(nameToCol[name]) d[name] = row.getCell(nameToCol[name]).value;
        });
        if(d["Bucket"] == null && d["RowId"] == null) continue;
        if(!valStr(d["Bucket"])) continue;
        snap.rows.push(d);
    }
    return snap;
}

function rowsRequiringRouting(snap) {
    const out = [];
    snap.rows.forEach(d => {
        if(statusIsRcmExcluded(valStr(d["Status"]))) return;
        if(!parseInclude(d["Include?"])) return;
        if(!statusNeedsBreakupRouting(valStr(d["Status"]))) return;
        const bucket = valStr(d["Bucket"]).toUpperCase();
        if(bucket === "PORTAL" || bucket === "HOLD" || bucket === "CDNR") out.push(d);
    });
    return out;
}

function snapshotToReconciliation(snap, route = {}) {
    const mainE=[], holdE=[], portalE=[], portalExc=[], cdnrE=[], rcmTally=[];

    const excelCellStr = (val) => {
        const dt = parseExcelDate(val);
        if(dt) return `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
        return valStr(val);
    };

    const rowToTally = d => ({
        date: parseExcelDate(d["BookDate"]),
        party: valStr(d["PartyOrName"]),
        gstin: valStr(d["GSTIN"]).toUpperCase(),
        vch_type: valStr(d["VchType"]),
        vch_no: valStr(d["VchNo"]),
        doc_no: valStr(d["DocOrInvoiceNo"]),
        doc_date: parseExcelDate(d["DocOrInvoiceDate"]),
        taxable: valFloat(d["Taxable"]),
        igst: valFloat(d["IGST"]),
        cgst: valFloat(d["CGST"]),
        sgst: valFloat(d["SGST"]),
        cess: valFloat(d["Cess"]),
        tax_amount: valFloat(d["TaxAmount"]),
        invoice_amount: valFloat(d["InvoiceValue"]),
        status: valStr(d["Status"])
    });

    const rowToGst = d => ({
        gstin: valStr(d["GSTIN"]).toUpperCase(),
        name: valStr(d["PartyOrName"]),
        invoice_no: valStr(d["DocOrInvoiceNo"]),
        invoice_type: valStr(d["InvoiceType"]),
        invoice_date: excelCellStr(d["DocOrInvoiceDate"]),
        invoice_value: valFloat(d["InvoiceValue"]),
        place_of_supply: valStr(d["PlaceOfSupply"]),
        reverse_charge: valStr(d["ReverseCharge"]),
        taxable: valFloat(d["Taxable"]),
        status: valStr(d["Status"]),
        igst: valFloat(d["IGST"]),
        cgst: valFloat(d["CGST"]),
        sgst: valFloat(d["SGST"]),
        cess: valFloat(d["Cess"]),
        period: valStr(d["GSTR1Period"]),
        filing_date: excelCellStr(d["FilingDate"]),
        itc_availability: valStr(d["ITCAvailability"])
    });

    const rowToCdnr = d => ({
        gstin: valStr(d["GSTIN"]).toUpperCase(),
        name: valStr(d["PartyOrName"]),
        note_no: valStr(d["DocOrInvoiceNo"]),
        note_type: valStr(d["CDNR_NoteType"]),
        note_supply_type: valStr(d["CDNR_SupplyType"]),
        note_date: excelCellStr(d["DocOrInvoiceDate"]),
        note_value: valFloat(d["CDNR_NoteValue"]) || valFloat(d["InvoiceValue"]),
        place_of_supply: valStr(d["PlaceOfSupply"]),
        reverse_charge: valStr(d["ReverseCharge"]),
        taxable: valFloat(d["Taxable"]),
        igst: valFloat(d["IGST"]),
        cgst: valFloat(d["CGST"]),
        sgst: valFloat(d["SGST"]),
        cess: valFloat(d["Cess"]),
        itc_availability: valStr(d["ITCAvailability"])
    });

    snap.rows.forEach(d => {
        if(statusIsRcmExcluded(valStr(d["Status"]))) return;
        const bucket = valStr(d["Bucket"]).toUpperCase();
        const inc = parseInclude(d["Include?"]);
        const rid = parseInt(d["RowId"], 10);

        if (bucket === "MAIN") {
            if (inc || isNotInGst2bStatus(valStr(d["Status"]))) {
                mainE.push(rowToTally(d));
            }
            return;
        }

        if (bucket === "PORTAL" || bucket === "HOLD" || bucket === "CDNR") {
            if (inc) {
                if (statusNeedsBreakupRouting(valStr(d["Status"]))) {
                    const c = route[rid];
                    if (c === "hold") holdE.push(rowToGst(d));
                    else if (c === "reverted") portalExc.push(rowToGst(d));
                    else if (c === "cdnr") cdnrE.push(rowToCdnr(d));
                    else {
                        // default behavior
                        if (bucket === "CDNR") cdnrE.push(rowToCdnr(d));
                        else if (bucket === "HOLD") holdE.push(rowToGst(d));
                        else portalE.push(rowToGst(d));
                    }
                } else {
                    if (bucket === "CDNR") cdnrE.push(rowToCdnr(d));
                    else if (bucket === "HOLD") holdE.push(rowToGst(d));
                    else portalE.push(rowToGst(d));
                }
            } else {
                if (bucket === "PORTAL" || bucket === "HOLD" || bucket === "CDNR") {
                   portalExc.push(rowToGst(d));
                }
            }
            return;
        }

        if (!inc) return;

        if (bucket === "RCM_BOOK") {
            rcmTally.push(rowToTally(d));
        }
    });

    const not_in_2b = mainE.filter(e => isNotInGst2bStatus(e.status));

    return {
        main_entries: mainE,
        not_in_2b: not_in_2b,
        hold_entries: holdE,
        portal_entries: portalE,
        rcm_tally: rcmTally,
        cdnr_entries: cdnrE,
        current_period: snap.period_tag,
        portal_excluded: portalExc
    };
}
