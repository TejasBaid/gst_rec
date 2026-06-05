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

function getOracleField(header) {
  const n = normHeader(header);
  if (!n) return null;
  if (n === 'accountingdate') return 'date';
  if (n === 'partyname') return 'party';
  if (n === 'tpregnno') return 'gstin';
  if (n === 'trxtype') return 'vch_type';
  if (n === 'trxnumber') return 'vch_no';
  if (n === 'supplierinvoicenum') return 'doc_no';
  if (n === 'supplierinvoicedate') return 'doc_date';
  if (n === 'gsttaxableamt') return 'taxable';
  if (n === 'igst') return 'igst';
  if (n === 'cgst') return 'cgst';
  if (n === 'sgst') return 'sgst';
  if (n === 'cess1') return 'cess1';
  if (n === 'cess2') return 'cess2';
  if (n === 'totalinvoiceamount') return 'invoice_amount';
  if (n === 'trxrec') return 'trx_rec';
  if (n === 'trxself') return 'trx_self';
  if (n === 'trxid') return 'trx_id';
  return null;
}

async function readOracleERP(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.getWorksheet('Sheet1') || workbook.worksheets[0];
  
  let headerRowIdx = null;
  let idxmap = {};
  
  ws.eachRow((row, rowNumber) => {
    if (headerRowIdx) return;
    const vals = row.values.slice(1);
    const hasAccDate = vals.some(v => normHeader(v) === 'accountingdate');
    const hasTpReg = vals.some(v => normHeader(v) === 'tpregnno');
    
    if (hasAccDate && hasTpReg) {
      headerRowIdx = rowNumber;
      vals.forEach((v, i) => {
        const field = getOracleField(v);
        if (field && !(field in idxmap)) idxmap[field] = i;
      });
    }
  });

  if (!headerRowIdx) throw new Error("Could not find Oracle ERP header row with 'ACCOUNTING DATE' and 'TP REGNNO'");
  if (idxmap['gstin'] === undefined) throw new Error("Oracle sheet missing TP REGNNO column");

  const dataStart = headerRowIdx + 1;
  const entries = [];
  const gstinIdx = idxmap['gstin'];
  const partyIdx = idxmap['party'] !== undefined ? idxmap['party'] : 1;

  for (let r = dataStart; r <= ws.rowCount; r++) {
    const row = ws.getRow(r);
    const vals = row.values.slice(1);
    if (!vals || vals.length === 0) continue;

    const g = valStr(vals[gstinIdx]);
    const isOutward = valStr(vals[idxmap['trx_self']]).trim().toUpperCase() === 'Y';
    
    // We can skip empty GSTINs UNLESS it is an outward supply (TRX SELF = Y)
    if (!isOutward && (!g || g.toLowerCase() === 'total')) continue;

    const dBook = parseExcelDate(vals[idxmap['date']]);
    let dDoc = parseExcelDate(vals[idxmap['doc_date']]);
    if (!dDoc) dDoc = dBook;

    const cessTotal = valFloat(vals[idxmap['cess1']]) + valFloat(vals[idxmap['cess2']]);

    entries.push({
      date: dBook,
      party: valStr(vals[partyIdx]),
      gstin: g.toUpperCase(),
      vch_type: valStr(vals[idxmap['vch_type']]),
      vch_no: valStr(vals[idxmap['vch_no']]),
      doc_no: valStr(vals[idxmap['doc_no']]),
      doc_date: dDoc,
      taxable: valFloat(vals[idxmap['taxable']]),
      igst: valFloat(vals[idxmap['igst']]),
      cgst: valFloat(vals[idxmap['cgst']]),
      sgst: valFloat(vals[idxmap['sgst']]),
      cess: cessTotal,
      tax_amount: valFloat(vals[idxmap['igst']]) + valFloat(vals[idxmap['cgst']]) + valFloat(vals[idxmap['sgst']]) + cessTotal,
      invoice_amount: valFloat(vals[idxmap['invoice_amount']]),
      trx_rec: valStr(vals[idxmap['trx_rec']]),
      trx_self: valStr(vals[idxmap['trx_self']]),
      trx_id: valStr(vals[idxmap['trx_id']]),
      status: "",
      raw_row: vals
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

async function readGstB2bAndCdnr(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.worksheets.find(w => w.name.toUpperCase().includes('B2B'));
  if (!ws) throw new Error("Sheet 'B2B' (or similar) not found in GST file. Note: Please ensure the file is saved as .xlsx and not .xlsb.");

  let hr = null;
  for (let r = 1; r <= 30; r++) {
    const rowVals = ws.getRow(r).values;
    if (!rowVals) continue;
    const isHeader = rowVals.some(v => {
      const s = valStr(v).toUpperCase();
      return s.includes("GSTIN") && s.includes("SUPPLIER");
    });
    if (isHeader) {
      hr = r; break;
    }
  }
  if (!hr) throw new Error("Could not find B2B header row with 'GSTIN of supplier'");

  let idxmap = {};
  const r1 = ws.getRow(hr).values.slice(1);
  for (let c = 0; c < r1.length; c++) {
    const v = valStr(r1[c]);
    const lc = v.toLowerCase();
    if (lc.includes('gstin') && lc.includes('supplier')) idxmap['gstin']=c;
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
    else if (v.trim().toUpperCase() === 'REASON') idxmap['reason']=c;
  }
  Object.keys(B2B_LEGACY).forEach(k => { if (idxmap[k] === undefined) idxmap[k] = B2B_LEGACY[k]; });

  const dataStart = hr + 1;
  const b2bEntries = [];
  const cdnrEntries = [];

  for (let r = dataStart; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values.slice(1);
    if (!vals || vals.length === 0 || vals[idxmap['gstin']] == null) continue;
    const gstin = valStr(vals[idxmap['gstin']]).toUpperCase();
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
    
    const invType = valStr(vals[idxmap['invoice_type']]).toLowerCase();
    const isCdnr = invType.includes('credit') || invType.includes('debit') || invType.includes('note') || invType.includes('cn') || invType.includes('dn');

    const entry = {
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
      itc_availability: valStr(vals[idxmap['itc_availability']]),
      reason: valStr(vals[idxmap['reason']])
    };

    b2bEntries.push(entry);
  }
  return { b2b: b2bEntries, cdnr: [] };
}

async function readGstIsd(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.worksheets.find(w => w.name.toUpperCase().includes('ISD'));
  if (!ws) return [];

  let hr = null;
  for (let r = 1; r <= 30; r++) {
    const rowVals = ws.getRow(r).values;
    if (!rowVals) continue;
    const isHeader = rowVals.some(v => valStr(v).toUpperCase().includes("ISD"));
    if (isHeader) { hr = r; break; }
  }
  if (!hr) return [];

  let idxmap = {};
  const r1 = ws.getRow(hr).values.slice(1);
  for (let c = 0; c < r1.length; c++) {
    const v = valStr(r1[c]).toLowerCase();
    if (v.includes('gstin of isd')) idxmap['gstin']=c;
    else if (v.includes('trade') || v.includes('legal')) idxmap['name']=c;
    else if (v.includes('document type')) idxmap['doc_type']=c;
    else if (v.includes('document number')) idxmap['doc_no']=c;
    else if (v.includes('document date')) idxmap['doc_date']=c;
    else if (v.includes('integrated tax')) idxmap['igst']=c;
    else if (v.includes('central tax')) idxmap['cgst']=c;
    else if (v.includes('state') && v.includes('tax')) idxmap['sgst']=c;
    else if (v.includes('cess')) idxmap['cess']=c;
  }

  const entries = [];
  const dataStart = hr + 1;
  for (let r = dataStart; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values.slice(1);
    if (!vals || vals.length === 0 || vals[idxmap['gstin']] == null) continue;
    
    let formatDt = valStr(vals[idxmap['doc_date']]);
    if (vals[idxmap['doc_date']] instanceof Date) {
      const dt = vals[idxmap['doc_date']];
      formatDt = `${String(dt.getDate()).padStart(2,'0')}/${String(dt.getMonth()+1).padStart(2,'0')}/${dt.getFullYear()}`;
    }

    entries.push({
      gstin: valStr(vals[idxmap['gstin']]).toUpperCase(),
      name: valStr(vals[idxmap['name']]),
      doc_no: valStr(vals[idxmap['doc_no']]),
      doc_type: valStr(vals[idxmap['doc_type']]),
      doc_date: formatDt,
      igst: valFloat(vals[idxmap['igst']]),
      cgst: valFloat(vals[idxmap['cgst']]),
      sgst: valFloat(vals[idxmap['sgst']]),
      cess: valFloat(vals[idxmap['cess']])
    });
  }
  return entries;
}

async function readIsdInputFile(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.worksheets[0]; // Usually the first sheet
  
  let hr = null;
  for (let r = 1; r <= 30; r++) {
    const rowVals = ws.getRow(r).values;
    if (!rowVals) continue;
    const isHeader = rowVals.some(v => {
      const s = valStr(v).toUpperCase();
      return s.includes("PARTY NAME") || s.includes("GST RATE") || s.includes("IGST");
    });
    if (isHeader) { hr = r; break; }
  }
  if (!hr) throw new Error("Could not find header row in ISD Input file.");

  let idxmap = {};
  const r1 = ws.getRow(hr).values.slice(1);
  for (let c = 0; c < r1.length; c++) {
    const v = valStr(r1[c]).toLowerCase();
    if (v.includes('party name')) idxmap['party_name']=c;
    else if (v.includes('party gst no')) idxmap['gstin']=c;
    else if (v.includes('basic amount')) idxmap['basic_amount']=c;
    else if (v === 'igst') idxmap['igst']=c;
    else if (v === 'cgst') idxmap['cgst']=c;
    else if (v === 'sgst') idxmap['sgst']=c;
    else if (v.includes('total gst')) idxmap['total_gst']=c;
    else if (v.includes('total value')) idxmap['total_value']=c;
  }

  const entries = [];
  const dataStart = hr + 1;
  for (let r = dataStart; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values.slice(1);
    if (!vals || vals.length === 0 || vals[idxmap['gstin']] == null) continue;
    
    entries.push({
      gstin: valStr(vals[idxmap['gstin']]).toUpperCase(),
      party_name: valStr(vals[idxmap['party_name']]),
      igst: valFloat(vals[idxmap['igst']]),
      cgst: valFloat(vals[idxmap['cgst']]),
      sgst: valFloat(vals[idxmap['sgst']]),
      total_gst: valFloat(vals[idxmap['total_gst']])
    });
  }
  return entries;
}

async function readTurnoversFile(arrayBuffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(arrayBuffer);
  let ws = workbook.worksheets[0];
  
  const entries = [];
  // Assuming simple two columns: GSTIN and Turnover
  // Sometimes it lacks headers or headers are on row 1
  for (let r = 1; r <= ws.rowCount; r++) {
    const vals = ws.getRow(r).values.slice(1);
    if (!vals || vals.length < 2) continue;
    
    const col1 = valStr(vals[0]);
    const col2 = valFloat(vals[1]);
    
    // Ignore header strings or empty GSTINs
    if (col1.length > 5 && col2 > 0 && !col1.toLowerCase().includes('gstin')) {
      entries.push({
        gstin: col1.toUpperCase().trim(),
        turnover: col2
      });
    }
  }
  return entries;
}
