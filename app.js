
const ui = {
    tallyPath: document.getElementById('tallyPath'),
    tallyFile: document.getElementById('tallyFile'),
    gstPath: document.getElementById('gstPath'),
    gstFile: document.getElementById('gstFile'),
    monthSelect: document.getElementById('monthSelect'),
    yearSelect: document.getElementById('yearSelect'),
    companyInput: document.getElementById('companyInput'),
    btnStep1: document.getElementById('btnStep1'),
    
    reviewPath: document.getElementById('reviewPath'),
    reviewFile: document.getElementById('reviewFile'),
    
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
    
    btnStep2: document.getElementById('btnStep2'),
    loader: document.getElementById('loader'),
    statusText: document.getElementById('statusText'),

    // modal
    routingModal: document.getElementById('routingModal'),
    routingTableBody: document.getElementById('routingTable').querySelector('tbody'),
    btnCancelRouting: document.getElementById('btnCancelRouting'),
    btnApplyRouting: document.getElementById('btnApplyRouting'),
};

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
ui.reviewFile.addEventListener('change', e => {
    if(e.target.files.length) ui.reviewPath.value = e.target.files[0].name;
});

function setStatus(msg, isError=false) {
    ui.statusText.textContent = msg;
    ui.statusText.style.color = isError ? "var(--error)" : "var(--text-muted)";
    console.log(msg);
}

function showLoader() { ui.loader.style.display = 'block'; }
function hideLoader() { ui.loader.style.display = 'none'; }

function parseNumberInput(val) {
    if(!val) return 0;
    const clean = val.replace(/,/g, '').trim();
    const f = parseFloat(clean);
    return isNaN(f) ? 0 : f;
}

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

ui.btnStep1.addEventListener('click', async () => {
    if(!ui.tallyFile.files.length) return alert("Select Tally.xlsx");
    if(!ui.gstFile.files.length) return alert("Select GST.xlsx");

    ui.btnStep1.disabled = true;
    showLoader();
    try {
        const month = ui.monthSelect.value;
        const year = parseInt(ui.yearSelect.value, 10);
        const company = ui.companyInput.value.trim() || 'Company';

        setStatus("Reading Tally file...");
        const tallyBuffer = await ui.tallyFile.files[0].arrayBuffer();
        const tallyEntries = await readTally(tallyBuffer);

        setStatus("Reading GST portal file...");
        const gstBuffer = await ui.gstFile.files[0].arrayBuffer();
        const gstB2b = await readGstB2b(gstBuffer);
        const gstCdnr = await readGstCdnr(gstBuffer);

        setStatus("Reconciling...");
        const result = reconcile(tallyEntries, gstB2b, gstCdnr, month, year);

        setStatus("Generating Review workbook...");
        const buffer = await exportReview(result, month, year, company);

        setStatus(`✅ Review generated successfully! Initiating download...`);
        saveFile(buffer, `Review_${month}_${year}.xlsx`);
    } catch(err) {
        console.error(err);
        setStatus(`❌ Error: ${err.message}`, true);
    } finally {
        hideLoader();
        ui.btnStep1.disabled = false;
    }
});

let pendingRoutingRows = null;
let activeReviewSnapshot = null;
let routingChoiceMap = {};

ui.btnStep2.addEventListener('click', async () => {
    if(!ui.reviewFile.files.length) return alert("Select Review.xlsx from Step 1");

    ui.btnStep2.disabled = true;
    showLoader();
    try {
        setStatus("Loading review file...");
        const revBuffer = await ui.reviewFile.files[0].arrayBuffer();
        activeReviewSnapshot = await loadReview(revBuffer);

        pendingRoutingRows = rowsRequiringRouting(activeReviewSnapshot);
        
        if (pendingRoutingRows && pendingRoutingRows.length > 0) {
            hideLoader();
            showRoutingModal(pendingRoutingRows);
        } else {
            setStatus("No classification routing needed. Proceeding...");
            await proceedGenerateFinal();
        }
    } catch (err) {
        console.error(err);
        setStatus(`❌ Error: ${err.message}`, true);
        hideLoader();
        ui.btnStep2.disabled = false;
    }
});

function showRoutingModal(rows) {
    ui.routingTableBody.innerHTML = '';
    
    rows.forEach((r, idx) => {
        const tr = document.createElement('tr');
        
        const party = valStr(r["PartyOrName"]) || "—";
        const doc = valStr(r["DocOrInvoiceNo"]) || "—";
        const st = valStr(r["Status"]) || "—";
        const rid = valStr(r["RowId"]);
        const buck = valStr(r["Bucket"]);
        
        tr.innerHTML = `
            <td>${rid}<br><small style="color:var(--text-muted)">${buck}</small></td>
            <td>${party.length > 50 ? party.substring(0,47)+'...' : party}</td>
            <td>${doc}</td>
            <td>${st}</td>
            <td>
                <select id="route_sel_${idx}" style="width: 100%">
                    <option value="" disabled selected>Select action...</option>
                    <option value="hold">Add : ITC hold earlier now reflected in GSTR 2B</option>
                    <option value="reverted">Less : ITC Reverted</option>
                    <option value="portal">Add : ITC taken from GSTR 2B</option>
                </select>
            </td>
        `;
        ui.routingTableBody.appendChild(tr);
    });

    ui.routingModal.classList.add('visible');
    ui.btnStep2.disabled = false;
}

ui.btnCancelRouting.addEventListener('click', () => {
    ui.routingModal.classList.remove('visible');
    setStatus("Final generation cancelled.");
});

ui.btnApplyRouting.addEventListener('click', async () => {
    routingChoiceMap = {};
    for(let i=0; i<pendingRoutingRows.length; i++) {
        const sel = document.getElementById(`route_sel_${i}`);
        if (!sel.value) {
            alert("Please select an action for all rows before applying.");
            return;
        }
        const rid = parseInt(pendingRoutingRows[i]["RowId"], 10);
        routingChoiceMap[rid] = sel.value;
    }
    
    ui.routingModal.classList.remove('visible');
    ui.btnStep2.disabled = true;
    showLoader();
    setStatus("Generating Final.xlsx from categorized review...");
    
    try {
        await proceedGenerateFinal();
    } catch (err) {
        console.error(err);
        setStatus(`❌ Error: ${err.message}`, true);
        hideLoader();
        ui.btnStep2.disabled = false;
    }
});

async function proceedGenerateFinal() {
    const rcm = {
        lease_rent: parseNumberInput(ui.rcmLease.value),
        office_rent: parseNumberInput(ui.rcmOffice.value),
        freight_outstation: parseNumberInput(ui.rcmFreightOut.value),
        freight_local: parseNumberInput(ui.rcmFreightLocal.value),
    };
    const manual = {
        sales_taxable: parseNumberInput(ui.salesTaxable.value),
        sales_igst: parseNumberInput(ui.salesIgst.value),
        sales_cgst_sgst: parseNumberInput(ui.salesCgst.value),
        opening_igst: parseNumberInput(ui.openIgst.value),
        opening_cgst: parseNumberInput(ui.openCgst.value),
        opening_sgst: parseNumberInput(ui.openSgst.value),
    };

    const result = snapshotToReconciliation(activeReviewSnapshot, routingChoiceMap);
    
    const { month, year, company } = activeReviewSnapshot;
    
    const buffer = await generateFinal(result, month, year, company, rcm, manual);
    
    setStatus(`✅ Final.xlsx generated successfully! Initiating download...`);
    saveFile(buffer, `Final_${month}_${year}.xlsx`);
    
    hideLoader();
    ui.btnStep2.disabled = false;
}
