// =====================================================================================
// PARAMETER CHANGE DETECTION (Smart Recalculation)
// =====================================================================================
// This tracks if actual parameters changed, not just any cell in the workbook

interface CallSignature {
    params: string;
    result: any;
}

const lastCallCache = new Map<string, CallSignature>();

function shouldRecalculate(funcName: string, ...params: any[]): { shouldCalc: boolean; cachedResult?: any } {
    const key = `${funcName}`;
    const paramString = JSON.stringify(params);
    
    const lastCall = lastCallCache.get(key);
    
    if (!lastCall || lastCall.params !== paramString) {
        // Parameters changed or first call - need to recalculate
        return { shouldCalc: true };
    }
    
    // Parameters unchanged - return cached result
    return { shouldCalc: false, cachedResult: lastCall.result };
}

function saveCallResult(funcName: string, result: any, ...params: any[]): void {
    const key = `${funcName}`;
    const paramString = JSON.stringify(params);
    lastCallCache.set(key, { params: paramString, result });
}

// =====================================================================================
// Custom Function Namespace: NRB
// =====================================================================================

/**
 * Test NRB API Connection
 * @customfunction
 * @returns {Promise<string>}
 * @alias NRB.TESTCONNECTION
 */
export async function TestConnection(): Promise<string> {
    const check = shouldRecalculate("TESTCONNECTION");
    if (!check.shouldCalc) return check.cachedResult;

    const testUrl = `https://www.nrb.org.np/api/forex/v1/rates?from=2024-01-01&to=2024-01-01&page=1&per_page=1`;

    try {
        const response = await fetch(testUrl);

        if (response.ok) {
            const jsonResponse = await response.json();
            const status = jsonResponse.status?.code ?? jsonResponse.status;

            if (status === 200 || status === 1) {
                const result = "SUCCESS: NRB API is reachable and responded with status 200/1.";
                saveCallResult("TESTCONNECTION", result);
                return result;
            }

            const snippet = JSON.stringify(jsonResponse).substring(0, 500);
            const result = `API WARNING: HTTP OK but unexpected status (${status}). Snippet:\n${snippet}...`;
            saveCallResult("TESTCONNECTION", result);
            return result;
        }

        const result = `NETWORK ERROR: HTTP ${response.status} - ${response.statusText}`;
        saveCallResult("TESTCONNECTION", result);
        return result;

    } catch (e) {
        const result = `ERROR: Could not complete request. (${e instanceof Error ? e.message : String(e)})`;
        saveCallResult("TESTCONNECTION", result);
        return result;
    }
}


/**
 * Get NRB Forex Rate (single date or date range)
 * @customfunction
 * @param {any} fromDate
 * @param {any} [toDate]
 * @param {string} [currencyPair="USDNPR"]
 * @param {string} [rateType="S"]
 * @returns {Promise<string | number>}
 * @alias NRB.FOREXRATE
 */
export async function ForexRate(
    fromDate: any,
    toDate?: any,
    currencyPair: string = "USDNPR",
    rateType: string = "S",
): Promise<string | number> {
    // Check if parameters actually changed
    const check = shouldRecalculate("FOREXRATE", fromDate, toDate, currencyPair, rateType);
    if (!check.shouldCalc) return check.cachedResult;

    try {
        if (!fromDate) return "#ERROR: From Date is required";

        const fromStr = formatDate(fromDate);
        if (fromStr === "ERROR") return "#ERROR: Invalid From Date";

        let toStr = fromStr;
        let isRange = false;

        if (toDate !== null && toDate !== undefined && toDate !== "") {
            toStr = formatDate(toDate);
            if (toStr === "ERROR") return "#ERROR: Invalid To Date";
            isRange = fromStr !== toStr;
        }

        currencyPair = String(currencyPair || "USDNPR").trim().toUpperCase();
        if (currencyPair.length !== 6) return "#ERROR: Invalid currency (e.g., USDNPR)";

        const needsInversion = currencyPair.startsWith("NPR");
        const targetCurrency = needsInversion ? currencyPair.slice(3) : currencyPair.slice(0, 3);

        rateType = String(rateType || "S").trim().toUpperCase();
        if (!["B", "S"].includes(rateType)) return "#ERROR: Rate type must be 'B' or 'S'";

        // RANGE CASE
        if (isRange) {
            const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${fromStr}&to=${toStr}&page=1&per_page=100`;
            const response = await fetch(url);
            if (!response.ok) return `#ERROR: HTTP ${response.status}`;

            const json = await response.json();
            const payload = json.data?.payload || json.data;

            if (!Array.isArray(payload) || payload.length === 0)
                return `#ERROR: No data for ${fromStr} to ${toStr}`;

            const results = extractRatesFromPayload(payload, targetCurrency, rateType, needsInversion);
            if (results.length === 0) return `#ERROR: Currency ${currencyPair} not found`;

            const result = JSON.stringify([["Date", "Rate"], ...results]);
            saveCallResult("FOREXRATE", result, fromDate, toDate, currencyPair, rateType);
            return result;
        }

        // SINGLE DATE WITH 7-DAY BACKOFF
        let cur = new Date(fromStr);
        for (let i = 0; i < 7; i++) {
            const dtStr = formatDate(cur);
            const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${dtStr}&to=${dtStr}&page=1&per_page=1`;

            const response = await fetch(url);
            if (response.ok) {
                const json = await response.json();
                const payload = json.data?.payload || json.data;

                if (Array.isArray(payload) && payload.length > 0) {
                    const results = extractRatesFromPayload(payload, targetCurrency, rateType, needsInversion);

                    if (results.length > 0) {
                        const finalRate = results[0][1];
                        const result = dtStr === fromStr ? finalRate : finalRate * -1;
                        saveCallResult("FOREXRATE", result, fromDate, toDate, currencyPair, rateType);
                        return result;
                    }
                }
            }

            cur.setDate(cur.getDate() - 1);
        }

        const result = `#ERROR: No rate found for ${currencyPair} on ${fromStr} or 7 preceding days`;
        saveCallResult("FOREXRATE", result, fromDate, toDate, currencyPair, rateType);
        return result;

    } catch (err) {
        const result = `#ERROR: ${err instanceof Error ? err.message : String(err)}`;
        saveCallResult("FOREXRATE", result, fromDate, toDate, currencyPair, rateType);
        return result;
    }
}


/**
 * Convert NRB.FOREXRATE JSON string back to Excel table
 * @customfunction
 * @param {string} jsonString
 * @returns {any[][]}
 * @alias NRB.PARSEJSON
 */
export function ParseJSON(jsonString: string): any[][] {
    const check = shouldRecalculate("PARSEJSON", jsonString);
    if (!check.shouldCalc) return check.cachedResult;

    try {
        const parsed = JSON.parse(jsonString);
        if (Array.isArray(parsed) && Array.isArray(parsed[0])) {
            saveCallResult("PARSEJSON", parsed, jsonString);
            return parsed;
        }
        const result = [["#ERROR: Invalid JSON structure"]];
        saveCallResult("PARSEJSON", result, jsonString);
        return result;
    } catch {
        const result = [["#ERROR: Invalid JSON string"]];
        saveCallResult("PARSEJSON", result, jsonString);
        return result;
    }
}


// =====================================================================================
// Utility Functions
// =====================================================================================

function formatDate(input: any): string {
    try {
        if (input === undefined || input === null || input === "") return "ERROR";

        let d: Date;

        if (input instanceof Date) {
            d = input;
        } else if (typeof input === "number") {
            const epoch = new Date(Date.UTC(1899, 11, 30));
            d = new Date(epoch.getTime() + input * 86400000);
        } else if (typeof input === "string") {
            let s = input.trim().replace(/[.\/]/g, "-");
            const monthMap: any = {
                jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
                jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
            };

            const words = s.split(/[-\s]/).map(x => {
                const m = monthMap[x.toLowerCase()];
                return m ? m : x;
            });
            s = words.join("-");

            if (/^\d{4}-\d{1,2}-\d{1,2}$/.test(s)) {
                const [y, m, dday] = s.split("-");
                d = new Date(`${y}-${m.padStart(2, "0")}-${dday.padStart(2, "0")}`);
            } else if (/^\d{1,2}-\d{1,2}-\d{4}$/.test(s)) {
                const [dday, m, y] = s.split("-");
                d = new Date(`${y}-${m.padStart(2, "0")}-${dday.padStart(2, "0")}`);
            } else {
                d = new Date(s);
            }
        } else {
            return "ERROR";
        }

        if (isNaN(d.getTime())) return "ERROR";

        return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
    } catch {
        return "ERROR";
    }
}

function extractRatesFromPayload(
    payload: any[],
    target: string,
    rateType: string,
    invert: boolean
): any[][] {
    const results: any[][] = [];
    const field = rateType === "B" ? "buy" : "sell";

    for (const entry of payload) {
        const date = entry.date;
        const rates = entry.rates;
        if (!Array.isArray(rates)) continue;

        for (const r of rates) {
            const iso3 = r.currency?.iso3 || r.iso3;
            if (iso3 === target) {
                let val = parseFloat(String(r[field]).replace(/,/g, ""));
                if (!isNaN(val) && val > 0) {
                    results.push([date, invert ? 1 / val : val]);
                }
                break;
            }
        }
    }

    return results;
}


// =====================================================================================
// Custom Function Association
// =====================================================================================

Office.onReady(() => {
    if (typeof CustomFunctions !== "undefined") {
        try {
            CustomFunctions.associate("NRB.TESTCONNECTION", TestConnection);
            CustomFunctions.associate("NRB.FOREXRATE", ForexRate);
            CustomFunctions.associate("NRB.PARSEJSON", ParseJSON);
        } catch {}
    }
});