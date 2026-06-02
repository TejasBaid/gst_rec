// excel_writer.js

function setColWidths(ws, widths) {
    widths.forEach((w, i) => { ws.getColumn(i+1).width = w; });
}

function generateInward(ws, result, isUnmatched = false) {
    let currentR = 3;
    
    const writeSection = (entries) => {
        if(!entries || entries.length === 0) return;
        entries.forEach(e => {
            if (e.raw_row) {
                const row = ws.getRow(currentR++);
                e.raw_row.forEach((val, i) => {
                    row.getCell(i + 1).value = val;
                });
            }
        });
    };

    if (isUnmatched) {
        writeSection(result.not_in_2b);
    } else {
        // We only write perfectly matched Oracle ERP entries that are ELIGIBLE for ITC or RCM.
        // PLUS we write Outward Supply entries (TRX SELF = Y) because the DELHI 3B template explicitly relies on them being in this sheet for Row 13 Outward calculations.
        writeSection(result.matched_eligible_tally);
        writeSection(result.matched_rcm_tally);
        writeSection(result.outward_tally);
    }
}

function base64ToArrayBuffer(base64) {
    const binaryString = window.atob(base64);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
}

async function generateFinal(result, month, year, company, rcm, manualIn) {
    const wb = new ExcelJS.Workbook();
    
    // Load the enterprise template from base64 string provided in template.js
    // Note: TEMPLATE_BASE64 must be defined in the global scope (e.g. from template.js in the browser)
    const templateBuffer = base64ToArrayBuffer(TEMPLATE_BASE64);
    await wb.xlsx.load(templateBuffer);
    
    // 1. Populate 'DELHI 3B' Header (only touch month/year)
    const ws3B = wb.getWorksheet('DELHI 3B');
    if (ws3B) {
        for (let r = 1; r <= 5; r++) {
            const row = ws3B.getRow(r);
            row.eachCell((cell, colNumber) => {
                const val = (cell.value || "").toString().toLowerCase();
                if (val.includes('year')) {
                    const targetCell = row.getCell(colNumber + 1);
                    if (!targetCell.value) targetCell.value = year;
                } else if (val.includes('month')) {
                    const targetCell = row.getCell(colNumber + 1);
                    if (!targetCell.value) targetCell.value = month;
                }
            });
        }
        
        // Fill Outward Data (Overwrites formula with manual UI inputs if provided)
        if (manualIn) {
            for (let r = 7; r <= 14; r++) {
                const row = ws3B.getRow(r);
                const labelCell = row.getCell(1).value;
                if (typeof labelCell !== 'string') continue;
                
                const label = labelCell.replace(/\s+/g, ' ').toLowerCase().trim();
                
                if (label.includes('other than zero rated, nil rated')) {
                    if (manualIn.sales_taxable) row.getCell(2).value = manualIn.sales_taxable;
                    if (manualIn.sales_igst) row.getCell(3).value = manualIn.sales_igst;
                    if (manualIn.sales_cgst_sgst) row.getCell(4).value = manualIn.sales_cgst_sgst;
                    if (manualIn.sales_cgst_sgst) row.getCell(5).value = manualIn.sales_cgst_sgst;
                }
            }
        }
        
        // Fix native template bug where Row 25 CGST sums SGST (AL) instead of CGST (AM)
        const row25 = ws3B.getRow(25);
        const cgstCell = row25.getCell(3);
        if (cgstCell && cgstCell.formula && cgstCell.formula.includes('AL:AL')) {
            cgstCell.formula = cgstCell.formula.replace('AL:AL', 'AM:AM');
        }

        // We DO NOT overwrite the ITC values in 'DELHI 3B'. 
        // We preserve the native SUMIFS formulas that pull from the 'Inward' sheet.
    }
    
    // 2. Clear and Rewrite 'Inward'
    let wsInward = wb.getWorksheet('Inward');
    if (!wsInward) {
        wsInward = wb.addWorksheet('Inward');
    } else {
        // Clear all data rows (keep row 1 & 2 as headers)
        wsInward.spliceRows(3, 5000);
    }
    generateInward(wsInward, result, false);

    // 3. Output Unmatched Oracle Entries to a NEW sheet (so they don't break Inward SUMIFS)
    let wsUnmatched = wb.getWorksheet('Not in 2B');
    if (!wsUnmatched) {
        wsUnmatched = wb.addWorksheet('Not in 2B');
        
        // Copy headers from Inward sheet
        const origHeader = wsInward.getRow(2);
        const newHeader = wsUnmatched.getRow(2);
        if (origHeader && newHeader) {
            origHeader.eachCell((cell, colNumber) => {
                newHeader.getCell(colNumber).value = cell.value;
                newHeader.getCell(colNumber).font = { bold: true };
            });
        }
    } else {
        wsUnmatched.spliceRows(3, 5000);
    }
    generateInward(wsUnmatched, result, true);

    const buffer = await wb.xlsx.writeBuffer();
    return buffer;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = { generateFinal };
}
