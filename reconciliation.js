// reconciliation.js
const STATUS_NOT_IN_GST_2B = "Not in GST 2B";

function isNotInGst2bStatus(status) {
    const u = (status || "").trim().toUpperCase();
    if (u === "NOT IN 2B") return true;
    if (u.includes("NOT IN GST 2B") || u === "NOT IN GST2B") return true;
    return (status || "").trim() === STATUS_NOT_IN_GST_2B;
}

function normCompare(s) {
    if (!s) return "";
    return String(s).replace(/[^0-9A-Za-z]/g, "").toUpperCase();
}
function normInv(s) { return normCompare(s); }
function normGstin(s) { return normCompare(s); }

function invSubstringMatch(invA, invB) {
    const a = normCompare(invA), b = normCompare(invB);
    const short = a.length <= b.length ? a : b;
    const llong = a.length <= b.length ? b : a;
    return short.length >= 4 && llong.includes(short);
}

function invDigitSuffixMatch(invA, invB) {
    function trailingDigits(s) {
        const str = normCompare(s);
        let res = "";
        for (let i = str.length - 1; i >= 0; i--) {
            if (/[0-9]/.test(str[i])) res = str[i] + res;
            else break;
        }
        return res;
    }
    const tdA = trailingDigits(invA), tdB = trailingDigits(invB);
    return tdA.length >= 4 && tdA === tdB;
}

function currentPeriodTag(month, year) {
    const abbr = month.substring(0, 3);
    const abbrCap = abbr.charAt(0).toUpperCase() + abbr.slice(1).toLowerCase();
    const yr2 = String(year).slice(-2);
    return `${abbrCap}'${yr2}`;
}

function isOkStatus(status) {
    return (status || "").trim().toUpperCase().startsWith("OK");
}

function isIneligible(status) {
    const s = (status || "").trim().toUpperCase();
    return s.includes("INELIGIBLE") || s.includes("FOOD") || s.includes("CAR");
}

function isRcm(gstEntry) {
    const s = (gstEntry.status || "").trim().toUpperCase();
    return s === "RCM" || (gstEntry.reverse_charge || "").trim().toUpperCase() === "YES";
}

function reconcile(tallyEntries, gstB2b, gstCdnr, isdEntries, month, year, resolutions = {}) {
    const currentPeriod = currentPeriodTag(month, year);
    const gstByKey = {};
    const gstByInv = {};
    const gstByGstin = {};

    for (const g of gstB2b) {
        const key = `${normGstin(g.gstin)}|${normInv(g.invoice_no)}`;
        gstByKey[key] = g;
        gstByInv[normInv(g.invoice_no)] = g;
        const gin = normGstin(g.gstin);
        if (!gstByGstin[gin]) gstByGstin[gin] = [];
        gstByGstin[gin].push(g);
    }

    const matchedGstInvKeys = new Set();

    function findGstMatch(tGstin, tInv) {
        let hit = gstByKey[`${tGstin}|${tInv}`];
        if (hit) return hit;
        hit = gstByInv[tInv];
        if (hit) return hit;
        const arr = gstByGstin[tGstin] || [];
        for (const g of arr) {
            if (invSubstringMatch(tInv, g.invoice_no) || invDigitSuffixMatch(tInv, g.invoice_no)) {
                return g;
            }
        }
        return null;
    }

    const matchedTally = [];
    const matchedEligibleTally = [];
    const matchedRcmTally = [];
    const outwardTally = [];
    const notIn2bTally = [];
    
    // We already built gstByKey, etc.
    const matchedGstKeys = new Set();
    const currentYearStr = (parseInt(year, 10) >= 2026 && month !== 'January' && month !== 'February' && month !== 'March') ? 'FY ' + year + '-' + (parseInt(year,10)+1).toString().slice(2) : 'FY ' + (parseInt(year,10)-1) + '-' + year.toString().slice(2);
    // Rough logic for 2B Year

    // Pre-calculate Tally total tax per GSTIN+INV key to ensure Diff is 0 when invoice matches perfectly
    const tallyAgg = {};
    for (const t of tallyEntries) {
        const k = t.gstin + t.doc_no;
        const rowTax = (t.igst || 0) + (t.cgst || 0) + (t.sgst || 0);
        tallyAgg[k] = (tallyAgg[k] || 0) + rowTax;
    }

    const duplicates = [];
    const overrides = [];
    const skipIndices = new Set();
    
    // 1. Detect Tally Duplicates (same GSTIN, Doc No, Tax amounts, different TRX REC)
    const tallyGroups = {};
    for (let i = 0; i < tallyEntries.length; i++) {
        const t = tallyEntries[i];
        t._originalIndex = i; // Keep track for skipping
        const k = `${t.gstin}|${t.doc_no}|${t.igst || 0}|${t.cgst || 0}|${t.sgst || 0}`;
        if (!tallyGroups[k]) tallyGroups[k] = [];
        tallyGroups[k].push(t);
    }
    
    for (const key in tallyGroups) {
        const group = tallyGroups[key];
        if (group.length > 1) {
            const hasY = group.find(g => (g.trx_rec || "").trim().toUpperCase() === "Y");
            const hasN = group.find(g => (g.trx_rec || "").trim().toUpperCase() === "N");
            if (hasY && hasN) {
                // It's a duplicate pair
                const conflictId = `dup_${hasY.gstin}_${hasY.doc_no}`;
                duplicates.push({ id: conflictId, yRow: hasY, nRow: hasN });
                
                // Apply resolution if provided
                if (resolutions[conflictId] === 'keepY') {
                    skipIndices.add(hasN._originalIndex);
                } else if (resolutions[conflictId] === 'keepN') {
                    skipIndices.add(hasY._originalIndex);
                }
            }
        }
    }

    for (let i = 0; i < tallyEntries.length; i++) {
        const t = tallyEntries[i];
        if (skipIndices.has(i)) continue; // User opted to drop this duplicate

        const tInv = normInv(t.doc_no);
        const tGstin = normGstin(t.gstin);
        const gstMatch = findGstMatch(tGstin, tInv);

        let vLookup = null;
        let diff = null;
        let remark = "";
        let itcType = "";
        let isRecov = (t.trx_rec || "").trim().toUpperCase() === "Y";

        // Col 61-71 preparation
        const stateCode = '07';
        const stateName = 'Delhi';
        const keyVal = t.gstin + t.doc_no;
        const totalTaxRow = (t.igst || 0) + (t.cgst || 0) + (t.sgst || 0);

        // Inv Year based on doc_date
        let invYear = 'FY 2026-27';
        if (t.doc_date instanceof Date) {
            if (t.doc_date.getFullYear() < 2026 || (t.doc_date.getFullYear() === 2026 && t.doc_date.getMonth() < 3)) {
                invYear = 'FY 2025-26';
            }
        }

        if (gstMatch) {
            matchedGstKeys.add(normInv(gstMatch.invoice_no));
            t.status = "Matched";
            matchedTally.push(t);
            
            const avail = (gstMatch.itc_availability || "").trim().toLowerCase();
            const rcm = isRcm(gstMatch);
            
            // 2. Detect ITC Overrides (Tally says Y, GST says No)
            if (isRecov && avail === "no") {
                const conflictId = `ovr_${t.gstin}_${t.doc_no}`;
                overrides.push({ id: conflictId, row: t });
                
                if (resolutions[conflictId] === 'keepRec') {
                    // User says keep as Rec despite 2B saying No
                    // Don't force anything
                } else {
                    // Default or 'forceNonRec': obey 2B
                    isRecov = false; 
                }
            } else if (avail === "no") {
                isRecov = false; // Always obey 2B if no conflict
            }
            
            remark = rcm ? "RCM" : "Matched";
            if (rcm) {
                itcType = isRecov ? "RCM Rec" : "RCM Non Rec";
                matchedRcmTally.push(t);
            } else {
                itcType = isRecov ? "Fwd Rec" : "Fwd Non Rec";
                matchedEligibleTally.push(t);
            }
            
            vLookup = (gstMatch.igst || 0) + (gstMatch.cgst || 0) + (gstMatch.sgst || 0);
            diff = tallyAgg[keyVal] - vLookup;
        } else {
            // Check if it's an outward supply (TRX SELF = Y). These don't belong in 2B, but must be kept in the Inward sheet for Outward math!
            if ((t.trx_self || "").trim().toUpperCase() === "Y") {
                t.status = "Outward Supply";
                remark = "Outward Supply";
                itcType = "Outward";
                outwardTally.push(t);
            } else {
                t.status = "Not in GSTR 2B";
                remark = "Not in 2B";
                itcType = isRecov ? "Fwd Rec" : "Fwd Non Rec";
                notIn2bTally.push(t);
            }
        }

        if (t.raw_row) {
            while (t.raw_row.length < 71) t.raw_row.push(null);
            // Ensure TRX REC matches our final resolved isRecov
            t.raw_row[50] = isRecov ? "Y" : "N";
            
            t.raw_row[60] = stateCode;
            t.raw_row[61] = stateName;
            t.raw_row[62] = keyVal;
            t.raw_row[63] = totalTaxRow;
            t.raw_row[64] = vLookup;
            t.raw_row[65] = diff;
            t.raw_row[66] = remark;
            t.raw_row[67] = t.trx_id;
            t.raw_row[68] = itcType;
            t.raw_row[69] = invYear;
            t.raw_row[70] = 1; // Count
        }
    }

    // Process any 'Force Includes' requested by the user
    if (resolutions && resolutions.forceIncludes && resolutions.forceIncludes.length > 0) {
        const forceSet = new Set(resolutions.forceIncludes.map(f => String(f.gstin).toUpperCase() + "_" + String(f.doc_no).toUpperCase()));
        const remainingNotIn2b = [];
        
        for (const t of notIn2bTally) {
            const key = String(t.gstin).toUpperCase() + "_" + String(t.doc_no).toUpperCase();
            if (forceSet.has(key)) {
                t.status = "Matched";
                
                let isRecov = ((t.trx_rec || "").trim().toUpperCase() === "Y");
                let itcType = isRecov ? "Fwd Rec" : "Fwd Non Rec";
                t.remark = "Matched";
                
                if (t.raw_row) {
                    t.raw_row[50] = isRecov ? "Y" : "N";
                    t.raw_row[66] = "Matched";
                    t.raw_row[68] = itcType;
                }
                
                matchedEligibleTally.push(t);
            } else {
                remainingNotIn2b.push(t);
            }
        }
        
        notIn2bTally.length = 0;
        notIn2bTally.push(...remainingNotIn2b);
    }

    const eligibleItc = [];
    const ineligibleItc = [];
    const rcmItc = [];
    const portalOnly = [];

    for (const g of gstB2b) {
        const avail = (g.itc_availability || "").trim().toLowerCase();
        const isMatched = matchedGstKeys.has(normCompare(g.invoice_no));
        
        if (!isMatched) {
            portalOnly.push(g);
        }

        if (isRcm(g)) {
            rcmItc.push(g);
        } else if (avail === "no") {
            ineligibleItc.push(g);
        } else if (isMatched) {
            // ONLY include invoices perfectly matched in Oracle ERP
            eligibleItc.push(g);
        }
    }

    return {
        matched_tally: matchedTally,
        matched_eligible_tally: matchedEligibleTally,
        matched_rcm_tally: matchedRcmTally,
        outward_tally: outwardTally,
        not_in_2b: notIn2bTally,
        portal_only: portalOnly,
        eligible_itc: eligibleItc,
        ineligible_itc: ineligibleItc,
        rcm_itc: rcmItc,
        cdnr_entries: gstCdnr,
        isd_entries: isdEntries,
        current_period: currentPeriod,
        conflicts: { duplicates, overrides }
    };
}
