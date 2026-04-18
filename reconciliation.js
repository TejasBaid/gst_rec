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

function reconcile(tallyEntries, gstB2b, cdnrEntries, month, year) {
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

    const mainEntries = [];
    const rcmTally = [];

    for (const t of tallyEntries) {
        const tInv = normInv(t.doc_no);
        const tGstin = normGstin(t.gstin);
        const gstMatch = findGstMatch(tGstin, tInv);

        if (gstMatch) {
            matchedGstInvKeys.add(normInv(gstMatch.invoice_no));
            if (isRcm(gstMatch)) {
                t.status = gstMatch.status;
                rcmTally.push(t);
                continue;
            }
            t.status = gstMatch.status;
        } else {
            t.status = STATUS_NOT_IN_GST_2B;
        }
        mainEntries.push(t);
    }

    const notIn2b = mainEntries.filter(e => isNotInGst2bStatus(e.status));
    const holdEntries = [];
    const portalEntries = [];

    for (const g of gstB2b) {
        const invKey = normCompare(g.invoice_no);
        const alreadyMatched = matchedGstInvKeys.has(invKey);

        if (isRcm(g)) {
            if (!alreadyMatched) portalEntries.push(g);
            continue;
        }

        if ((g.period || "").trim() !== currentPeriod && isOkStatus(g.status)) {
            if (!alreadyMatched) holdEntries.push(g);
            continue;
        }

        if (!alreadyMatched) {
            portalEntries.push(g);
        }
    }

    return {
        main_entries: mainEntries,
        not_in_2b: notIn2b,
        hold_entries: holdEntries,
        portal_entries: portalEntries,
        rcm_tally: rcmTally,
        cdnr_entries: cdnrEntries,
        current_period: currentPeriod,
        portal_excluded: []
    };
}
