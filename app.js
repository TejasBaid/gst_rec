const ui = {
    appContainer: document.getElementById('appContainer'),
    dashboardContainer: document.getElementById('dashboardContainer'),
    
    tallyPath: document.getElementById('tallyPath'),
    tallyFile: document.getElementById('tallyFile'),
    gstPath: document.getElementById('gstPath'),
    gstFile: document.getElementById('gstFile'),
    monthSelect: document.getElementById('monthSelect'),
    yearSelect: document.getElementById('yearSelect'),
    companyInput: document.getElementById('companyInput'),
    
    rcmLease: document.getElementById('rcmLease'),
    rcmOffice: document.getElementById('rcmOffice'),
    rcmFreightOut: document.getElementById('rcmFreightOut'),
    rcmFreightLocal: document.getElementById('rcmFreightLocal'),
    
    salesTaxable: document.getElementById('salesTaxable'),
    salesCgst: document.getElementById('salesCgst'),
    salesIgst: document.getElementById('salesIgst'),
    
    openIgst: document.getElementById('openIgst'),
    openCgst: document.getElementById('openCgst'),
    openSgst: document.getElementById('openSgst'),
    
    btnGenerate: document.getElementById('btnGenerate'),
    loader: document.getElementById('loader'),
    statusText: document.getElementById('statusText'),

    dashTotalMatched: document.getElementById('dashTotalMatched'),
    dashTotalEligible: document.getElementById('dashTotalEligible'),
    dashTotalMissing: document.getElementById('dashTotalMissing'),
    dashTotalPortal: document.getElementById('dashTotalPortal'),
    
    tbodyMatched: document.getElementById('tbodyMatched'),
    tbodyMissing: document.getElementById('tbodyMissing'),
    tbodyPortal: document.getElementById('tbodyPortal')
};

// Wizard Logic
let currentStep = 1;
const totalSteps = 4;
let cachedTallyEntries = null;
let cachedGstB2b = null;
let cachedGstCdnr = null;
let cachedGstIsd = null;
let cachedConflicts = null;

function nextStep(step) {
    if (step === 2) {
        if(!ui.tallyFile.files.length) return alert("Select Oracle ERP .xlsx");
        if(!ui.gstFile.files.length) return alert("Select GST portal .xlsx/.xlsb");
    }
    
    document.getElementById(`step${currentStep}`).classList.remove('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.remove('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.add('completed');
    
    currentStep = step;
    
    document.getElementById(`step${currentStep}`).classList.add('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.add('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.remove('completed');
}

function prevStep(step) {
    document.getElementById(`step${currentStep}`).classList.remove('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.remove('active');
    
    currentStep = step;
    
    document.getElementById(`step${currentStep}`).classList.add('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.add('active');
    document.getElementById(`stepIndicator${currentStep}`).classList.remove('completed');
}

function resetWizard() {
    ui.dashboardContainer.style.display = 'none';
    ui.appContainer.style.display = 'block';
    prevStep(1);
    document.getElementById('stepIndicator2').classList.remove('completed');
    document.getElementById('stepIndicator3').classList.remove('completed');
    ui.statusText.textContent = "Ready";
    ui.btnGenerate.disabled = false;
}

// Dashboard Tab Logic
function switchTab(tabId) {
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(content => content.classList.remove('active'));
    
    event.target.classList.add('active');
    document.getElementById(tabId).classList.add('active');
}

// Initialize Year Select
const currentYear = new Date().getFullYear();
const monthIndex = new Date().getMonth();
ui.monthSelect.selectedIndex = monthIndex;

for(let y=currentYear-2; y<=currentYear+2; y++) {
    const opt = document.createElement('option');
    opt.value = y; opt.text = y;
    if(y === currentYear) opt.selected = true;
    ui.yearSelect.appendChild(opt);
}

ui.tallyFile.addEventListener('change', e => {
    if(e.target.files.length) ui.tallyPath.value = e.target.files[0].name;
});
ui.gstFile.addEventListener('change', e => {
    if(e.target.files.length) ui.gstPath.value = e.target.files[0].name;
});

function setStatus(msg, isError=false) {
    ui.statusText.textContent = msg;
    ui.statusText.style.color = isError ? "var(--error)" : "var(--text-muted)";
}

function showLoader() { ui.loader.style.display = 'block'; }
function hideLoader() { ui.loader.style.display = 'none'; }

document.getElementById('btnGoToReview').addEventListener('click', async () => {
    if(!ui.tallyFile.files.length) return alert("Select Oracle ERP .xlsx");
    if(!ui.gstFile.files.length) return alert("Select GST portal .xlsx/.xlsb");

    nextStep(3);
    document.getElementById('reviewLoader').style.display = 'block';
    document.getElementById('reviewContent').style.display = 'none';

    try {
        const month = ui.monthSelect.value;
        const year = parseInt(ui.yearSelect.value, 10);
        
        if (!cachedTallyEntries) {
            const tallyBuffer = await ui.tallyFile.files[0].arrayBuffer();
            cachedTallyEntries = await readOracleERP(tallyBuffer);
        }
        if (!cachedGstB2b) {
            const gstBuffer = await ui.gstFile.files[0].arrayBuffer();
            const gstParsed = await readGstB2bAndCdnr(gstBuffer);
            cachedGstB2b = gstParsed.b2b;
            cachedGstCdnr = gstParsed.cdnr;
            cachedGstIsd = await readGstIsd(gstBuffer);
        }
        
        const result = reconcile(cachedTallyEntries, cachedGstB2b, cachedGstCdnr, cachedGstIsd, month, year, {});
        cachedConflicts = result.conflicts;
        cachedConflicts.not_in_2b = result.not_in_2b;
        
        renderConflicts();
        document.getElementById('reviewLoader').style.display = 'none';
        document.getElementById('reviewContent').style.display = 'block';
    } catch(err) {
        console.error(err);
        alert("Error reading files: " + err.message);
        prevStep(2);
    }
});

function renderConflicts() {
    const tbodyDup = document.querySelector('#tableDuplicates tbody');
    const tbodyOvr = document.querySelector('#tableOverrides tbody');
    tbodyDup.innerHTML = '';
    tbodyOvr.innerHTML = '';

    if (cachedConflicts.duplicates.length === 0) {
        tbodyDup.innerHTML = '<tr><td colspan="3" style="text-align:center">No duplicates found.</td></tr>';
    } else {
        cachedConflicts.duplicates.forEach(d => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${d.yRow.doc_no}</td>
                <td>${d.yRow.gstin}</td>
                <td>
                    <label><input type="radio" name="${d.id}" value="keepY" checked> Keep TRX REC=Y (Rec)</label><br>
                    <label><input type="radio" name="${d.id}" value="keepN"> Keep TRX REC=N (Non-Rec)</label><br>
                    <label><input type="radio" name="${d.id}" value="keepBoth"> Keep Both (Duplicate)</label>
                </td>
            `;
            tbodyDup.appendChild(tr);
        });
    }

    if (cachedConflicts.overrides.length === 0) {
        tbodyOvr.innerHTML = '<tr><td colspan="3" style="text-align:center">No ITC availability conflicts found.</td></tr>';
    } else {
        cachedConflicts.overrides.forEach(o => {
            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${o.row.doc_no}</td>
                <td>${o.row.gstin}</td>
                <td>
                    <label><input type="radio" name="${o.id}" value="forceNonRec" checked> Force Non-Rec (Match 2B)</label><br>
                    <label><input type="radio" name="${o.id}" value="keepRec"> Keep as Rec (Override 2B)</label>
                </td>
            `;
            tbodyOvr.appendChild(tr);
        });
    }

    const tbodyMiss = document.querySelector('#tableNotIn2B tbody');
    tbodyMiss.innerHTML = '';
    
    // Filter out 0 tax rows and sort by absolute tax amount descending
    const sortedMissing = cachedConflicts.not_in_2b.filter(m => {
        const tax = (m.igst || 0) + (m.cgst || 0) + (m.sgst || 0);
        return Math.abs(tax) > 0.01;
    }).sort((a, b) => {
        const taxA = Math.abs((a.igst || 0) + (a.cgst || 0) + (a.sgst || 0));
        const taxB = Math.abs((b.igst || 0) + (b.cgst || 0) + (b.sgst || 0));
        return taxB - taxA; // Descending
    });

    if (sortedMissing.length === 0) {
        tbodyMiss.innerHTML = '<tr><td colspan="5" style="text-align:center">All Oracle invoices matched successfully.</td></tr>';
    } else {
        sortedMissing.forEach((m, idx) => {
            const tr = document.createElement('tr');
            tr.className = 'missing-row';
            tr.setAttribute('data-search', `${m.doc_no} ${m.gstin} ${m.party_name}`.toLowerCase());
            
            const checkboxId = `force_${idx}`;
            const totalTax = (m.igst || 0) + (m.cgst || 0) + (m.sgst || 0);
            tr.innerHTML = `
                <td><input type="checkbox" class="force-include-cb" data-doc="${m.doc_no}" data-gstin="${m.gstin}"></td>
                <td>${m.doc_no}</td>
                <td>${m.gstin}</td>
                <td>${m.party_name}</td>
                <td>${formatCurrency(totalTax)}</td>
            `;
            tbodyMiss.appendChild(tr);
        });
    }
}

document.getElementById('searchNotIn2B')?.addEventListener('input', (e) => {
    const q = e.target.value.toLowerCase();
    document.querySelectorAll('.missing-row').forEach(tr => {
        if (tr.getAttribute('data-search').includes(q)) {
            tr.style.display = '';
        } else {
            tr.style.display = 'none';
        }
    });
});

function saveFile(buffer, filename) {
    const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function formatCurrency(val) {
    if (!val) return '₹0.00';
    return '₹' + parseFloat(val).toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function renderTable(tbody, entries, isPortal) {
    tbody.innerHTML = '';
    entries.forEach(e => {
        const tr = document.createElement('tr');
        
        let inv, party, date, igst, cgst, sgst;
        if (isPortal) {
            inv = e.invoice_no || e.note_num || '';
            party = e.trade_name || e.legal_name || '';
            date = e.invoice_date || e.note_date || '';
            igst = e.igst || 0;
            cgst = e.cgst || 0;
            sgst = e.sgst || 0;
        } else {
            inv = e.doc_no || '';
            party = e.party_name || '';
            date = e.doc_date || '';
            igst = e.igst || 0;
            cgst = e.cgst || 0;
            sgst = e.sgst || 0;
        }

        tr.innerHTML = `
            <td>${inv}</td>
            <td>${party}</td>
            <td>${date}</td>
            <td style="color:var(--primary)">${formatCurrency(igst)}</td>
            <td style="color:var(--primary)">${formatCurrency(cgst)}</td>
            <td style="color:var(--primary)">${formatCurrency(sgst)}</td>
        `;
        tbody.appendChild(tr);
    });
}

function renderDashboard(result) {
    ui.dashTotalMatched.textContent = result.matched_eligible_tally.length + result.matched_rcm_tally.length;
    
    let eligibleIgst = 0;
    result.matched_eligible_tally.forEach(t => { eligibleIgst += (parseFloat(t.igst) || 0); });
    result.matched_rcm_tally.forEach(t => { eligibleIgst += (parseFloat(t.igst) || 0); });
    
    ui.dashTotalEligible.textContent = formatCurrency(eligibleIgst);
    
    ui.dashTotalMissing.textContent = result.not_in_2b.length;
    ui.dashTotalPortal.textContent = result.portal_only.length;
    
    const matchedCombined = [...result.matched_eligible_tally, ...result.matched_rcm_tally];
    
    renderTable(ui.tbodyMatched, matchedCombined, false);
    renderTable(ui.tbodyMissing, result.not_in_2b, false);
    renderTable(ui.tbodyPortal, result.portal_only, true);
    
    // Transition UI
    ui.appContainer.style.display = 'none';
    ui.dashboardContainer.style.display = 'block';
}

ui.btnGenerate.addEventListener('click', async () => {
    ui.btnGenerate.disabled = true;
    document.getElementById('btnBackTo3').style.display = 'none';
    showLoader();
    try {
        const month = ui.monthSelect.value;
        const year = parseInt(ui.yearSelect.value, 10);
        const company = ui.companyInput.value.trim() || 'Company';

        setStatus("Applying resolutions and reconciling...");
        
        // Harvest resolutions from UI
        const resolutions = {
            forceIncludes: []
        };
        if (cachedConflicts) {
            cachedConflicts.duplicates.forEach(d => {
                const checked = document.querySelector(`input[name="${d.id}"]:checked`);
                if (checked) resolutions[d.id] = checked.value;
            });
            cachedConflicts.overrides.forEach(o => {
                const checked = document.querySelector(`input[name="${o.id}"]:checked`);
                if (checked) resolutions[o.id] = checked.value;
            });
            
            document.querySelectorAll('.force-include-cb:checked').forEach(cb => {
                resolutions.forceIncludes.push({
                    doc_no: cb.getAttribute('data-doc'),
                    gstin: cb.getAttribute('data-gstin')
                });
            });
        }

        const result = reconcile(cachedTallyEntries, cachedGstB2b, cachedGstCdnr, cachedGstIsd, month, year, resolutions);

        setStatus("Generating GSTR-3B using Template...");
        
        const rcmInput = {
            lease_rent: parseFloat(ui.rcmLease.value) || 0,
            office_rent: parseFloat(ui.rcmOffice.value) || 0,
            freight_outstation: parseFloat(ui.rcmFreightOut.value) || 0,
            freight_local: parseFloat(ui.rcmFreightLocal.value) || 0
        };

        const manualInput = {
            sales_taxable: parseFloat(ui.salesTaxable.value) || 0,
            sales_cgst_sgst: parseFloat(ui.salesCgst.value) || 0,
            sales_igst: parseFloat(ui.salesIgst.value) || 0,
            open_igst: parseFloat(ui.openIgst.value) || 0,
            open_cgst: parseFloat(ui.openCgst.value) || 0,
            open_sgst: parseFloat(ui.openSgst.value) || 0
        };

        const buffer = await generateFinal(result, month, year, company, rcmInput, manualInput);

        setStatus(`✅ GSTR-3B generated! Unlocking dashboard...`);
        saveFile(buffer, `Final_${month}_${year}.xlsx`);
        
        setTimeout(() => {
            renderDashboard(result);
        }, 1500);

    } catch(err) {
        console.error(err);
        setStatus(`❌ Error: ${err.message}`, true);
        ui.btnGenerate.disabled = false;
        document.getElementById('btnBackTo3').style.display = 'block';
    } finally {
        hideLoader();
    }
});
