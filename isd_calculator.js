// isd_calculator.js
// Generates output in EXACT isdbook.xlsx format

const GST_STATE_MAP = {
    "01": "Jammu and Kashmir", "02": "Himachal Pradesh", "03": "Punjab",
    "04": "Chandigarh", "05": "Uttarakhand", "06": "Haryana", "07": "Delhi",
    "08": "Rajasthan", "09": "Uttar Pradesh", "10": "Bihar", "11": "Sikkim",
    "12": "Arunachal Pradesh", "13": "Nagaland", "14": "Manipur", "15": "Mizoram",
    "16": "Tripura", "17": "Meghalaya", "18": "Assam", "19": "West Bengal",
    "20": "Jharkhand", "21": "Odisha", "22": "Chattisgarh", "23": "Madhya Pradesh",
    "24": "Gujarat", "25": "Daman and Diu", "26": "Dadra and Nagar Haveli",
    "27": "Maharashtra", "29": "Karnataka", "30": "Goa", "31": "Lakshadweep",
    "32": "Kerala", "33": "Tamil Nadu", "34": "Puducherry",
    "35": "Andaman and Nicobar Islands", "36": "Telangana", "37": "Andhra Pradesh",
    "38": "Ladakh"
};

// Aliases for state names that appear with different spellings in source data
const STATE_NAME_ALIASES = {
    "jammu & kashmir":    "01",
    "jammu and kashmir":  "01",
    "j&k":                "01",
    "himachal pradesh":   "02",
    "punjab":             "03",
    "chandigarh":         "04",
    "uttarakhand":        "05",
    "haryana":            "06",
    "delhi":              "07",
    "rajasthan":          "08",
    "uttar pradesh":      "09",
    "up":                 "09",
    "bihar":              "10",
    "assam":              "18",
    "west bengal":        "19",
    "jharkhand":          "20",
    "odisha":             "21",
    "chattisgarh":        "22",
    "chhattisgarh":       "22",
    "madhya pradesh":     "23",
    "mp":                 "23",
    "gujarat":            "24",
    "maharashtra":        "27",
    "karnataka":          "29",
    "goa":                "30",
    "kerala":             "32",
    "tamil nadu":         "33",
    "telangana":          "36",
    "andhra pradesh":     "37",
    "ladakh":             "38",
    "ladhakh":            "38"
};

// Resolve a state name string to a 2-digit state code (or null)
function resolveStateCode(name) {
    const key = name.trim().toLowerCase();
    return STATE_NAME_ALIASES[key] || null;
}

function n(v) {
    if (v == null || v === '' || (typeof v === 'number' && isNaN(v))) return 0;
    const f = parseFloat(v);
    return isNaN(f) ? 0 : f;
}

function s(v) {
    if (v == null) return '';
    if (typeof v === 'object' && v.richText) return v.richText.map(r => r.text).join('').trim();
    if (typeof v === 'object' && v.result != null) return String(v.result).trim();
    return String(v).trim();
}

function fmt(num) {
    return Math.round(num * 1e10) / 1e10; // round to 10 decimal places
}

// Calculate distribution - returns result object + per-row distributions
function calculateIsdDistribution(invoices, turnovers, isdStateCode) {
    // Filter states with turnover > 0
    const validStates = turnovers.filter(t => t.turnover > 0);
    const totalTurnover = turnovers.reduce((acc, t) => acc + t.turnover, 0);
    const grandTotal = turnovers.reduce((acc, t) => acc + t.turnover, 0); // includes 0-turnover states

    let totalPool = 0;
    invoices.forEach(inv => {
        totalPool += n(inv.igst) + n(inv.cgst) + n(inv.sgst);
    });

    return { totalPool, totalTurnover, validStates, grandTotal };
}

// Main export function - builds entire workbook
async function exportIsdToExcel(isdBuffer, turnovers, isdStateCode) {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(isdBuffer);
    const ws = workbook.worksheets[0];

    // ── 1. Locate header row ──────────────────────────────────────────
    let headerRow = -1;
    ws.eachRow((row, rNum) => {
        if (headerRow !== -1) return;
        const vals = row.values;
        if (!vals) return;
        const hasPN = vals.some(v => s(v).toUpperCase().includes('PARTY NAME'));
        const hasGST = vals.some(v => s(v).toUpperCase().includes('PARTY GST'));
        if (hasPN && hasGST) headerRow = rNum;
    });
    if (headerRow === -1) throw new Error('Could not find header row in input file.');

    // ── 2. Map existing column indices (1-based) ─────────────────────
    const hdr = ws.getRow(headerRow);
    const colOf = {};
    hdr.eachCell((cell, colNum) => {
        const v = s(cell.value).toLowerCase().trim();
        if (v === 'party name') colOf.partyName = colNum;
        else if (v === 'party gst no') colOf.gstin = colNum;
        else if (v === 'supplier invoice no') colOf.invNo = colNum;
        else if (v === 'supplier invoice date') colOf.invDate = colNum;
        else if (v === 'accounting date') colOf.accDate = colNum;
        else if (v === 'gst rate') colOf.gstRate = colNum;
        else if (v === 'basic amount') colOf.basic = colNum;
        else if (v === 'igst') colOf.igst = colNum;
        else if (v === 'cgst') colOf.cgst = colNum;
        else if (v === 'sgst') colOf.sgst = colNum;
        else if (v === 'total gst') colOf.totalGst = colNum;
        else if (v === 'total value') colOf.totalValue = colNum;
        else if (v === 'recoverable') colOf.recoverable = colNum;
        else if (v === 'pan india/ state') colOf.panState = colNum;
        else if (v === 'io name') colOf.ioName = colNum;
    });

    // Last column of the input section (1-based ExcelJS = hdr.values.length-1)
    let inputLastCol = 0;
    hdr.eachCell((cell, c) => { if (c > inputLastCol) inputLastCol = c; });

    // ── 3. Build state list from Turnovers ────────────────────────────
    const validStates = turnovers.filter(t => t.turnover > 0);
    const totalTurnover = turnovers.reduce((acc, t) => acc + t.turnover, 0);

    // ── 4. Define new column ranges ───────────────────────────────────
    // Intermediate columns (appended right after col 25/IO Name):
    //  col 26: GSTIN+INV
    //  col 27: Total Tax
    //  col 28: GSTR6A
    //  col 29: Diffrence
    //  col 30: Remark
    //  col 31: ITC Type
    const C_GSTIN_INV = inputLastCol + 1;  // 27 in isdbook (0-indexed=26)
    const C_TOTAL_TAX = C_GSTIN_INV + 1;   // 28
    const C_GSTR6A    = C_TOTAL_TAX + 1;   // 29
    const C_DIFF      = C_GSTR6A + 1;      // 30
    const C_REMARK    = C_DIFF + 1;        // 31
    const C_ITC_TYPE  = C_REMARK + 1;      // 32

    // State distribution columns start here (4 cols per state: IGST, CGST, SGST, Total)
    const STATE_START = C_ITC_TYPE + 1;    // 33

    // Checking/summary section (3 cols after all states)
    const numStates = validStates.length;
    const C_TOTAL_DIST = STATE_START + (numStates * 4);      // col after all states
    const C_DIFF_CHK   = C_TOTAL_DIST + 1;
    const C_CHK_IGST   = C_DIFF_CHK + 1;
    const C_CHK_CGST   = C_CHK_IGST + 1;
    const C_CHK_SGST   = C_CHK_CGST + 1;

    // Helper: column index for a state's IGST/CGST/SGST/Total
    function stateCol(stateIdx, taxType) {
        const base = STATE_START + stateIdx * 4;
        return base + { igst: 0, cgst: 1, sgst: 2, total: 3 }[taxType];
    }

    // ── 5. Write metadata rows (rows 1-7) ────────────────────────────
    // Row 3: Turnover label
    ws.getCell(3, C_ITC_TYPE).value = 'Turnover';
    // Row 4: Grand total + per-state turnovers
    ws.getCell(4, C_TOTAL_TAX).value = totalTurnover; // grand total under "Total Tax" col
    validStates.forEach((st, i) => {
        ws.getCell(4, stateCol(i, 'igst')).value = st.turnover;
    });
    // Row 5: GSTINs
    validStates.forEach((st, i) => {
        ws.getCell(5, stateCol(i, 'igst')).value = st.gstin;
    });
    ws.getCell(5, C_TOTAL_DIST).value = 'CHECKING';
    // Row 6: State names
    validStates.forEach((st, i) => {
        const code = st.gstin.substring(0, 2);
        const name = GST_STATE_MAP[code] || st.gstin;
        ws.getCell(6, stateCol(i, 'igst')).value = name;
    });
    ws.getCell(6, C_TOTAL_DIST).value = 'Total distributed';
    ws.getCell(6, C_DIFF_CHK).value = 'Diff.';
    ws.getCell(6, C_CHK_IGST).value = 'Total distributed';

    // Row 7: Summary totals (filled after data rows)
    // Row 8: Headers
    ws.getCell(headerRow, C_GSTIN_INV).value = 'GSTIN+INV';
    ws.getCell(headerRow, C_TOTAL_TAX).value = 'Total Tax';
    ws.getCell(headerRow, C_GSTR6A).value    = 'GSTR6A';
    ws.getCell(headerRow, C_DIFF).value      = 'Diffrence';
    ws.getCell(headerRow, C_REMARK).value    = 'Remark';
    ws.getCell(headerRow, C_ITC_TYPE).value  = 'ITC Type';
    validStates.forEach((st, i) => {
        ws.getCell(headerRow, stateCol(i, 'igst')).value  = 'IGST';
        ws.getCell(headerRow, stateCol(i, 'cgst')).value  = 'CGST';
        ws.getCell(headerRow, stateCol(i, 'sgst')).value  = 'SGST';
        ws.getCell(headerRow, stateCol(i, 'total')).value = 'Total';
    });
    ws.getCell(headerRow, C_CHK_IGST).value = 'IGST';
    ws.getCell(headerRow, C_CHK_CGST).value = 'CGST';
    ws.getCell(headerRow, C_CHK_SGST).value = 'SGST';

    // ── 6. Process data rows ─────────────────────────────────────────
    // Accumulators for summary row (row 7)
    const summaryTotals = {
        totalTax: 0,
        totalDist: 0,
        chkIgst: 0, chkCgst: 0, chkSgst: 0
    };
    const stateSummary = validStates.map(() => ({ igst: 0, cgst: 0, sgst: 0 }));

    for (let r = headerRow + 1; r <= ws.rowCount; r++) {
        const row = ws.getRow(r);
        const partyName = s(row.getCell(colOf.partyName).value);
        if (!partyName || partyName.toLowerCase() === 'total') continue;

        const gstinVal     = s(row.getCell(colOf.gstin).value);
        const invNoVal     = s(row.getCell(colOf.invNo).value);
        const igstVal      = n(row.getCell(colOf.igst).value);
        const cgstVal      = n(row.getCell(colOf.cgst).value);
        const sgstVal      = n(row.getCell(colOf.sgst).value);
        const totalTaxVal  = igstVal + cgstVal + sgstVal;
        const gstr6aVal    = n(row.getCell(colOf.totalGst).value);
        const panState     = s(row.getCell(colOf.panState).value).trim();
        const recov        = s(row.getCell(colOf.recoverable).value).trim().toUpperCase();
        const itcType      = recov === 'Y' ? 'Recoverable' : 'Non-Recoverable';

        // Intermediate columns
        row.getCell(C_GSTIN_INV).value = gstinVal + invNoVal;
        row.getCell(C_TOTAL_TAX).value = fmt(totalTaxVal);
        row.getCell(C_GSTR6A).value    = gstr6aVal;
        row.getCell(C_DIFF).value      = fmt(totalTaxVal - gstr6aVal);
        row.getCell(C_REMARK).value    = Math.abs(totalTaxVal - gstr6aVal) < 1 ? 'Matched' : 'Mismatch';
        row.getCell(C_ITC_TYPE).value  = itcType;

        // Only distribute if recoverable
        if (recov !== 'Y') {
            // Set all state cols to 0 / null
            validStates.forEach((st, i) => {
                row.getCell(stateCol(i, 'igst')).value  = 0;
                row.getCell(stateCol(i, 'cgst')).value  = 0;
                row.getCell(stateCol(i, 'sgst')).value  = 0;
                row.getCell(stateCol(i, 'total')).value = 0;
            });
            row.getCell(C_TOTAL_DIST).value = 0;
            row.getCell(C_DIFF_CHK).value   = 0;
            row.getCell(C_CHK_IGST).value   = 0;
            row.getCell(C_CHK_CGST).value   = 0;
            row.getCell(C_CHK_SGST).value   = 0;
            row.commit();
            continue;
        }

        // ── Distribution logic ────────────────────────────────────────
        // Determine distribution type:
        //   isPanIndia  → distribute to ALL states proportionally
        //   isMultiState → specific list of named states, distribute proportionally AMONG those states only
        //   isSingleState → 100% to one matching state only
        const panStateUpper = panState.toUpperCase();
        const isPanIndia = panStateUpper === 'PAN INDIA' || panState === '' || panState === 'o';
        const isMultiState = !isPanIndia && panState.includes(',');

        // For multi-state: resolve the listed state codes
        let multiStateCodes = [];
        if (isMultiState) {
            const parts = panState.split(',');
            parts.forEach(p => {
                const code = resolveStateCode(p.trim());
                if (code) multiStateCodes.push(code);
            });
        }

        // For multi-state: compute sub-turnover total (only among matching states)
        let multiStateTurnoverTotal = 0;
        if (isMultiState && multiStateCodes.length > 0) {
            validStates.forEach(st => {
                if (multiStateCodes.includes(st.gstin.substring(0, 2))) {
                    multiStateTurnoverTotal += st.turnover;
                }
            });
        }

        let rowTotalDist = 0;
        let rowChkIgst = 0, rowChkCgst = 0, rowChkSgst = 0;

        validStates.forEach((st, i) => {
            const stateCode = st.gstin.substring(0, 2);
            const isSameState = stateCode === isdStateCode;

            let distIgst = 0, distCgst = 0, distSgst = 0;

            if (isPanIndia) {
                // Distribute proportionally to ALL states
                const ratio = totalTurnover > 0 ? st.turnover / totalTurnover : 0;
                const igstAlloc = igstVal * ratio;
                const cgstAlloc = cgstVal * ratio;
                const sgstAlloc = sgstVal * ratio;

                if (isSameState) {
                    distIgst = igstAlloc;
                    distCgst = cgstAlloc;
                    distSgst = sgstAlloc;
                } else {
                    distIgst = igstAlloc + cgstAlloc + sgstAlloc;
                    distCgst = 0;
                    distSgst = 0;
                }
            } else if (isMultiState) {
                // Distribute proportionally AMONG the listed states only
                const isInList = multiStateCodes.includes(stateCode);
                if (isInList && multiStateTurnoverTotal > 0) {
                    const ratio = st.turnover / multiStateTurnoverTotal;
                    const igstAlloc = igstVal * ratio;
                    const cgstAlloc = cgstVal * ratio;
                    const sgstAlloc = sgstVal * ratio;

                    if (isSameState) {
                        distIgst = igstAlloc;
                        distCgst = cgstAlloc;
                        distSgst = sgstAlloc;
                    } else {
                        distIgst = igstAlloc + cgstAlloc + sgstAlloc;
                        distCgst = 0;
                        distSgst = 0;
                    }
                } else {
                    distIgst = 0; distCgst = 0; distSgst = 0;
                }
            } else {
                // Single state-specific: entire amount goes to matching state ONLY
                const targetCode = resolveStateCode(panState);
                const isMatch = targetCode ? (stateCode === targetCode) : false;

                if (isMatch) {
                    if (isSameState) {
                        distIgst = igstVal;
                        distCgst = cgstVal;
                        distSgst = sgstVal;
                    } else {
                        distIgst = igstVal + cgstVal + sgstVal;
                        distCgst = 0;
                        distSgst = 0;
                    }
                } else {
                    distIgst = 0; distCgst = 0; distSgst = 0;
                }
            }

            const distTotal = distIgst + distCgst + distSgst;

            row.getCell(stateCol(i, 'igst')).value  = fmt(distIgst)  || 0;
            row.getCell(stateCol(i, 'cgst')).value  = fmt(distCgst)  || 0;
            row.getCell(stateCol(i, 'sgst')).value  = fmt(distSgst)  || 0;
            row.getCell(stateCol(i, 'total')).value = fmt(distTotal) || 0;

            rowTotalDist += distTotal;
            rowChkIgst   += distIgst;
            rowChkCgst   += distCgst;
            rowChkSgst   += distSgst;

            // Accumulate for summary row
            stateSummary[i].igst += distIgst;
            stateSummary[i].cgst += distCgst;
            stateSummary[i].sgst += distSgst;
        });

        // Checking columns
        row.getCell(C_TOTAL_DIST).value = fmt(rowTotalDist);
        row.getCell(C_DIFF_CHK).value   = fmt(totalTaxVal - rowTotalDist);
        row.getCell(C_CHK_IGST).value   = fmt(rowChkIgst);
        row.getCell(C_CHK_CGST).value   = fmt(rowChkCgst);
        row.getCell(C_CHK_SGST).value   = fmt(rowChkSgst);

        summaryTotals.totalTax  += totalTaxVal;
        summaryTotals.totalDist += rowTotalDist;
        summaryTotals.chkIgst   += rowChkIgst;
        summaryTotals.chkCgst   += rowChkCgst;
        summaryTotals.chkSgst   += rowChkSgst;

        row.commit();
    }

    // ── 7. Write Row 7 summary totals ────────────────────────────────
    const summaryRowNum = headerRow - 1; // row 7 in isdbook (headerRow=8)
    ws.getCell(summaryRowNum, C_TOTAL_TAX).value = fmt(summaryTotals.totalTax);
    ws.getCell(summaryRowNum, C_TOTAL_DIST).value = fmt(summaryTotals.totalDist);
    ws.getCell(summaryRowNum, C_DIFF_CHK).value   = fmt(summaryTotals.totalTax - summaryTotals.totalDist);
    ws.getCell(summaryRowNum, C_CHK_IGST).value   = fmt(summaryTotals.chkIgst);
    ws.getCell(summaryRowNum, C_CHK_CGST).value   = fmt(summaryTotals.chkCgst);
    ws.getCell(summaryRowNum, C_CHK_SGST).value   = fmt(summaryTotals.chkSgst);

    validStates.forEach((st, i) => {
        ws.getCell(summaryRowNum, stateCol(i, 'igst')).value  = fmt(stateSummary[i].igst);
        ws.getCell(summaryRowNum, stateCol(i, 'cgst')).value  = fmt(stateSummary[i].cgst);
        ws.getCell(summaryRowNum, stateCol(i, 'sgst')).value  = fmt(stateSummary[i].sgst);
        ws.getCell(summaryRowNum, stateCol(i, 'total')).value = fmt(stateSummary[i].igst + stateSummary[i].cgst + stateSummary[i].sgst);
    });

    const buffer = await workbook.xlsx.writeBuffer();
    return buffer;
}

// Simple aggregated result for the dashboard display only
function calculateIsdDistribution(invoices, turnovers, isdStateCode) {
    const validStates = turnovers.filter(t => t.turnover > 0);
    const totalTurnover = turnovers.reduce((acc, t) => acc + t.turnover, 0);
    let totalPool = 0;
    invoices.forEach(inv => {
        totalPool += n(inv.igst || 0) + n(inv.cgst || 0) + n(inv.sgst || 0);
    });
    const distribution = validStates.map(st => {
        const ratio = totalTurnover > 0 ? st.turnover / totalTurnover : 0;
        const code = st.gstin.substring(0, 2);
        const name = GST_STATE_MAP[code] || st.gstin;
        return { gstin: st.gstin, stateName: name, turnover: st.turnover, ratio };
    });
    return { totalPool, totalTurnover, distribution };
}
