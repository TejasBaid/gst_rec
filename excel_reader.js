// excel_reader.js

function parseExcelDate(val) {
  if (val instanceof Date) return val;
  if (typeof val === 'number') {
    // Excel date serial number
    const utc_days = Math.floor(val - 25569);
    const utc_value = utc_days * 86400; 
    const date_info = new Date(utc_value * 1000);
    return date_info;
  }
  if (typeof val === 'string') {
    const parts = val.split(/[-/]/);
    if (parts.length === 3) {
      if (parts[2].length === 4) {
        return new Date(parts[2], parts[1]-1, parts[0]);
      }
    }
  }
  return null;
}

function valFloat(val) {
  if (val == null) return 0;
  if (typeof val?.result === 'number') return val.result;
  const v = parseFloat(val);
  return isNaN(v) ? 0 : v;
}

function valStr(val) {
  if (val == null) return "";
  if (typeof val === 'object' && val.richText) {
    return val.richText.map(r => r.text).join('').trim();
  }
  if (typeof val?.result !== 'undefined') return String(val.result).trim();
  return String(val).trim();
}

function normHeader(h) {
  return String(h).trim().toLowerCase().replace(/[^a-z0-9]/g, '');
}

function getTallyField(header) {
  const n = normHeader(header);
  if (!n) return null;
  if (n === 'date') return 'date';
  if (n === 'particulars') return 'particulars';
  if (n.includes('gstin') || n.endsWith('gstinuin')) return 'gstin';
  if (n === 'vchtype') return 'vch_type';
  if (n === 'vchno') return 'vch_no';
  if (n === 'docno') return 'doc_no';
  if (n === 'docdate' || n === 'doc') return 'doc_date';
  if (n.includes('taxable')) return 'taxable';
  if (n === 'igst') return 'igst';
  if (n === 'cgst') return 'cgst';
  if (n.startsWith('sgst') || n === 'sgstutgst') return 'sgst';
  if (n === 'cess') return 'cess';
  if (n === 'tax' || n === 'taxamount') return 'tax_amount';
  if (n === 'invoice' || n === 'invoiceamount') return 'invoice_amount';
  return null;
}

async function readTally(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.getWorksheet('Sheet1') || workbook.worksheets[0];
  
  let headerRowIdx = null;
  let idxmap = {};
  
  ws.eachRow((row, rowNumber) => {
    if (headerRowIdx) return;
    const vals = row.values.slice(1);
    const v0 = valStr(vals[0]);
    const v1 = valStr(vals[1]);
    if (v0 === 'Date' && v1 === 'Particulars') {
      headerRowIdx = rowNumber;
      vals.forEach((v, i) => {
        const field = getTallyField(v);
        if (field && !(field in idxmap)) idxmap[field] = i;
      });
    }
  });

  if (!headerRowIdx) throw new Error("Could not find Tally header row with 'Date' and 'Particulars'");
  if (idxmap['gstin'] === undefined) throw new Error("Tally sheet no GSTIN column");

  const dataStart = headerRowIdx + 2;
  const entries = [];
  const gstinIdx = idxmap['gstin'];
  const particIdx = idxmap['particulars'] !== undefined ? idxmap['particulars'] : 1;

  for (let r = dataStart; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const vals = row.values.slice(1);
    if (!vals || vals.length === 0) continue;

    const p = valStr(vals[particIdx]);
    if (p.toLowerCase() === 'total' || p.toLowerCase() === 'totals') break;

    if (!p && vals[gstinIdx] == null) continue;
    if (p && p.toLowerCase() !== 'total') {
      if (!vals[gstinIdx] && !vals[0]) continue;
    }

    const g = valStr(vals[gstinIdx]);
    if (!g) continue;

    const dBook = parseExcelDate(vals[idxmap['date']]);
    let dDoc = parseExcelDate(vals[idxmap['doc_date']]);
    if (!dDoc) dDoc = dBook;

    entries.push({
      date: dBook,
      party: valStr(vals[idxmap['particulars']]),
      gstin: g.toUpperCase(),
      vch_type: valStr(vals[idxmap['vch_type']]),
      vch_no: valStr(vals[idxmap['vch_no']]),
      doc_no: valStr(vals[idxmap['doc_no']]),
      doc_date: dDoc,
      taxable: valFloat(vals[idxmap['taxable']]),
      igst: valFloat(vals[idxmap['igst']]),
      cgst: valFloat(vals[idxmap['cgst']]),
      sgst: valFloat(vals[idxmap['sgst']]),
      cess: valFloat(vals[idxmap['cess']]),
      tax_amount: valFloat(vals[idxmap['tax_amount']]),
      invoice_amount: valFloat(vals[idxmap['invoice_amount']]),
      status: ""
    });
  }
  return entries;
}

const B2B_LEGACY = {
  gstin: 0, name: 1, invoice_no: 2, invoice_type: 3, invoice_date: 4,
  invoice_value: 5, place_of_supply: 6, reverse_charge: 7, taxable: 8,
  status: 9, igst: 10, cgst: 11, sgst: 12, cess: 13, period: 14,
  filing_date: 15, itc_availability: 16
};

const B2B_TWO_ROW = {
  gstin: 0, name: 1, invoice_no: 2, invoice_type: 3, invoice_date: 4,
  invoice_value: 5, place_of_supply: 6, reverse_charge: 7, taxable: 8,
  igst: 9, cgst: 10, sgst: 11, cess: 12, period: 13,
  filing_date: 14, itc_availability: 15
};

async function readGstB2b(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  const ws = workbook.getWorksheet('B2B');
  if (!ws) throw new Error("Sheet 'B2B' not found in GST file.");

  let hr = null;
  for (let r = 1; r <= 30; r++) {
    const v = valStr(ws.getRow(r).getCell(1).value).toUpperCase();
    if (v.includes("GSTIN") && v.includes("SUPPLIER")) {
      hr = r; break;
    }
  }
  if (!hr) throw new Error("Could not find B2B header row with 'GSTIN of supplier'");

  const nextCell = valStr(ws.getRow(hr + 1).getCell(3).value).toLowerCase();
  const twoRow = nextCell.includes("invoice number");

  let idxmap = {};
  let dataStart;

  if (twoRow) {
    const r1 = ws.getRow(hr).values.slice(1);
    const r2 = ws.getRow(hr + 1).values.slice(1);
    for (let c = 0; c < Math.max(r1.length, r2.length); c++) {
      const a = valStr(r1[c]).toLowerCase();
      const b = valStr(r2[c]).toLowerCase();
      const au = valStr(r1[c]).toUpperCase();
      const bu = valStr(r2[c]).toUpperCase();

      if (au.trim() === 'STATUS') { idxmap['status'] = c; continue; }
      if (bu.trim() === 'STATUS') { idxmap['status'] = c; continue; }

      if (b.includes('invoice number')) idxmap['invoice_no'] = c;
      else if (b.includes('invoice type')) idxmap['invoice_type'] = c;
      else if (b.includes('invoice date')) idxmap['invoice_date'] = c;
      else if (b.includes('invoice value')) idxmap['invoice_value'] = c;
      else if (a.startsWith('gstin')) idxmap['gstin'] = c;
      else if (a.includes('trade') && a.includes('legal')) idxmap['name'] = c;
      else if (a.includes('place of supply')) idxmap['place_of_supply'] = c;
      else if (a.includes('supply attract') || a.includes('reverse charge')) idxmap['reverse_charge'] = c;
      else if (a.includes('taxable value')) idxmap['taxable'] = c;
      else if (b.includes('integrated tax')) idxmap['igst'] = c;
      else if (b.includes('central tax')) idxmap['cgst'] = c;
      else if (b.includes('state') && b.includes('tax')) idxmap['sgst'] = c;
      else if (b.includes('cess')) idxmap['cess'] = c;
      else if (a.includes('gstr-1') && a.includes('period')) idxmap['period'] = c;
      else if (a.includes('filing date')) idxmap['filing_date'] = c;
      else if (a.includes('itc availability')) idxmap['itc_availability'] = c;
    }
    Object.keys(B2B_TWO_ROW).forEach(k => { if (idxmap[k] === undefined) idxmap[k] = B2B_TWO_ROW[k]; });
    dataStart = hr + 2;
  } else {
    const r1 = ws.getRow(hr).values.slice(1);
    for (let c = 0; c < r1.length; c++) {
      const v = valStr(r1[c]);
      const n = normHeader(v);
      const lc = v.toLowerCase();
      if (n === 'gstinofthesupplier' || (n.startsWith('gstin') && lc.includes('supplier'))) idxmap['gstin']=c;
      else if (lc.includes('trade') && lc.includes('legal')) idxmap['name']=c;
      else if (lc.includes('invoice number')) idxmap['invoice_no']=c;
      else if (lc.includes('invoice type')) idxmap['invoice_type']=c;
      else if (lc.includes('invoice date')) idxmap['invoice_date']=c;
      else if (lc.includes('invoice value')) idxmap['invoice_value']=c;
      else if (lc.includes('place of supply')) idxmap['place_of_supply']=c;
      else if (lc.includes('reverse charge') || lc.includes('supply attract')) idxmap['reverse_charge']=c;
      else if (lc.includes('taxable')) idxmap['taxable']=c;
      else if (v.trim().toUpperCase() === 'STATUS') idxmap['status']=c;
      else if (lc.includes('integrated tax')) idxmap['igst']=c;
      else if (lc.includes('central tax')) idxmap['cgst']=c;
      else if (lc.includes('state') && lc.includes('tax')) idxmap['sgst']=c;
      else if (lc.includes('cess')) idxmap['cess']=c;
      else if (lc.includes('gstr-1') && lc.includes('period')) idxmap['period']=c;
      else if (lc.includes('filing date')) idxmap['filing_date']=c;
      else if (lc.includes('itc availability')) idxmap['itc_availability']=c;
    }
    Object.keys(B2B_LEGACY).forEach(k => { if (idxmap[k] === undefined) idxmap[k] = B2B_LEGACY[k]; });
    dataStart = hr + 1;
  }

  const entries = [];
  for (let r = dataStart; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values.slice(1);
    if (!vals || vals.length === 0 || vals[0] == null) continue;
    const gstin = valStr(vals[0]).toUpperCase();
    if (gstin === '' || gstin.startsWith('=')) continue;

    let formatDt = valStr(vals[idxmap['invoice_date']]);
    if (vals[idxmap['invoice_date']] instanceof Date) {
      const dt = vals[idxmap['invoice_date']];
      formatDt = `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
    }

    let filingDt = valStr(vals[idxmap['filing_date']]);
    if (vals[idxmap['filing_date']] instanceof Date) {
      const dt = vals[idxmap['filing_date']];
      filingDt = `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
    }

    entries.push({
      gstin,
      name: valStr(vals[idxmap['name']]),
      invoice_no: valStr(vals[idxmap['invoice_no']]),
      invoice_type: valStr(vals[idxmap['invoice_type']]),
      invoice_date: formatDt,
      invoice_value: valFloat(vals[idxmap['invoice_value']]),
      place_of_supply: valStr(vals[idxmap['place_of_supply']]),
      reverse_charge: valStr(vals[idxmap['reverse_charge']]),
      taxable: valFloat(vals[idxmap['taxable']]),
      status: valStr(vals[idxmap['status']]),
      igst: valFloat(vals[idxmap['igst']]),
      cgst: valFloat(vals[idxmap['cgst']]),
      sgst: valFloat(vals[idxmap['sgst']]),
      cess: valFloat(vals[idxmap['cess']]),
      period: valStr(vals[idxmap['period']]),
      filing_date: filingDt,
      itc_availability: valStr(vals[idxmap['itc_availability']])
    });
  }
  return entries;
}

async function readGstCdnr(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.worksheets.find(w => {
    const u = w.name.toUpperCase();
    return u.includes("CDNR") && u.includes("B2B") && !u.includes("-CDNRA") && !u.includes("(REJECTED)");
  });
  if (!ws) ws = workbook.getWorksheet("B2B-CDNR");
  if (!ws) return [];

  let hr = 1;
  for (let r = 1; r <= 20; r++) {
    const v = valStr(ws.getRow(r).getCell(1).value).toUpperCase();
    if (v.includes("GSTIN") && v.includes("SUPPLIER")) { hr = r; break; }
  }

  const entries = [];
  const start = hr + 2; 
  for (let r = start; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const vals = row.values.slice(1);
    if (!vals || vals.length < 13 || !vals[0]) continue;
    
    let dt = vals[5];
    let formatDt = "";
    if (dt instanceof Date) {
        formatDt = `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
    } else {
        formatDt = valStr(dt);
    }

    entries.push({
      gstin: valStr(vals[0]).toUpperCase(),
      name: valStr(vals[1]),
      note_no: valStr(vals[2]),
      note_type: valStr(vals[3]),
      note_supply_type: valStr(vals[4]),
      note_date: formatDt,
      note_value: valFloat(vals[6]),
      taxable: valFloat(vals[9]),
      igst: valFloat(vals[10]),
      cgst: valFloat(vals[11]),
      sgst: valFloat(vals[12]),
      cess: valFloat(vals[13])
    });
  }
  return entries;
}
