// =====================================================================================
// Excel Desktop Calculation Control - Prevents Flickering
// =====================================================================================

let calculationControlInitialized = false;

if (typeof Office !== "undefined") {
  Office.onReady(() => {
    if (!calculationControlInitialized && typeof Excel !== "undefined") {
      calculationControlInitialized = true;
      
      // Disable automatic calculation mode for Desktop
      try {
        Excel.run(async (context) => {
          const app = context.application;
          app.calculationMode = Excel.CalculationMode.manual;
          await context.sync();
        }).catch(() => {
          // Silently ignore if not available
        });
      } catch (e) {
        // Ignore errors
      }
    }
  });
}


// =====================================================================================
// Custom Function Namespace: NRB
// =====================================================================================

/**
 * Test NRB API Connection
 * @customfunction
 * @volatile false
 * @returns {Promise<string>} Connection status
 * @alias NRB.TESTCONNECTION
 */
export async function TestConnection(): Promise<string> {
    const testUrl = `https://www.nrb.org.np/api/forex/v1/rates?from=2024-01-01&to=2024-01-01&page=1&per_page=1`;

    try {
        const response = await fetch(testUrl, {
            headers: { 'User-Agent': 'Excel-NRB-Addon' } 
        });

        if (response.ok) {
            const jsonResponse = await response.json();

            const status = jsonResponse.status?.code ?? jsonResponse.status;

            if (status === 200 || status === 1) {
                return "SUCCESS: NRB API is reachable and responded with status 200/1.";
            }

            const debugDetail = JSON.stringify(jsonResponse, null, 2);
            return `API WARNING: HTTP OK but unexpected status (${status}). Response snippet:\n${debugDetail.substring(0, 500)}...`;
        }

        return `NETWORK ERROR: HTTP Status ${response.status} - ${response.statusText}`;

    } catch (e) {
        return `ERROR: Could not complete request. (${e instanceof Error ? e.message : String(e)})`;
    }
}


/**
 * Get NRB Forex Rate (smart single/range return)
 * @customfunction
 * @volatile false
 * @param {any} fromDate Start date
 * @param {any} [toDate] End date (optional)
 * @param {string} [currencyPair="USDNPR"] Currency pair (default: USDNPR)
 * @param {string} [rateType="S"] B for Buy, S for Sell (default: S)
 * @returns {Promise<string | number>} Forex rate (number for single date, JSON string for range)
 * @alias NRB.FOREXRATE
 */
export async function ForexRate(
    fromDate: any,
    toDate?: any,
    currencyPair: string = "USDNPR",
    rateType: string = "S",
): Promise<string | number> {
    try {
        if (!fromDate) return "#ERROR: From Date is required";

        const requestedFromStr = formatDate(fromDate);
        if (requestedFromStr === "ERROR") return "#ERROR: Invalid From Date";

        let toStr = requestedFromStr;
        let isRange = false;

        if (toDate !== null && toDate !== undefined && toDate !== "") {
            toStr = formatDate(toDate);
            if (toStr === "ERROR") return "#ERROR: Invalid To Date";
            isRange = requestedFromStr !== toStr;
        }

        currencyPair = String(currencyPair || "USDNPR").trim().toUpperCase();
        if (currencyPair.length !== 6) return "#ERROR: Invalid currency (e.g., USDNPR)";

        const needsInversion = currencyPair.startsWith("NPR");
        const targetCurrency = needsInversion ? currencyPair.slice(3) : currencyPair.slice(0, 3);

        rateType = String(rateType || "S").trim().toUpperCase();
        if (!["B", "S"].includes(rateType)) return "#ERROR: Rate type must be 'B' or 'S'";

        // =================== RANGE HANDLING ===================
        if (isRange) {
            const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${requestedFromStr}&to=${toStr}&page=1&per_page=100`;
            const response = await fetch(url, { headers: { 'User-Agent': 'Excel-NRB-Addon' } });
            if (!response.ok) return `#ERROR: HTTP ${response.status}`;
            const json = await response.json();

            const payload = json.data?.payload || json.data;
            if (!Array.isArray(payload) || payload.length === 0)
                return `#ERROR: No data from ${requestedFromStr} to ${toStr}`;

            const results = extractRatesFromPayload(payload, targetCurrency, rateType, needsInversion);
            if (results.length === 0) return `#ERROR: Currency ${currencyPair} not found`;

            return JSON.stringify([["Date", "Rate"], ...results]);
        }

        // =================== SINGLE DATE HANDLING ===================
        let currentDay = new Date(requestedFromStr);
        for (let i = 0; i < 7; i++) {
            const searchDateStr = formatDate(currentDay);
            const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${searchDateStr}&to=${searchDateStr}&page=1&per_page=1`;
            const response = await fetch(url, { headers: { 'User-Agent': 'Excel-NRB-Addon' } });

            if (response.ok) {
                const json = await response.json();
                const payload = json.data?.payload || json.data;
                if (Array.isArray(payload) && payload.length > 0) {
                    const results = extractRatesFromPayload(payload, targetCurrency, rateType, needsInversion);
                    if (results.length > 0) {
                        const finalRate = results[0][1];

                        if (searchDateStr === requestedFromStr) {
                            return finalRate;
                        } else {
                            return finalRate * -1;
                        }
                    }
                }
            }

            currentDay.setDate(currentDay.getDate() - 1);
        }

        return `#ERROR: No rate found for ${currencyPair} on ${requestedFromStr} or 7 preceding days`;
    } catch (err) {
        return `#ERROR: ${err instanceof Error ? err.message : String(err)}`;
    }
}

/**
 * Helper function to parse JSON string back into table
 * @customfunction
 * @volatile false
 * @param {string} jsonString JSON string returned by NRB.FOREXRATE
 * @returns {any[][]} Parsed table for Excel
 * @alias NRB.PARSEJSON
 */
export function ParseJSON(jsonString: string): any[][] {
    try {
        const parsed = JSON.parse(jsonString);
        if (Array.isArray(parsed) && Array.isArray(parsed[0])) return parsed;
        return [["#ERROR: Invalid JSON structure"]];
    } catch {
        return [["#ERROR: Invalid JSON string"]];
    }
}


// =====================================================================================
// HELPER FUNCTIONS 
// =====================================================================================

/**
 * Format date from various input types to 'yyyy-mm-dd' (UTC-safe)
 */
function formatDate(inputDate: any): string {
    try {
        if (inputDate === null || inputDate === undefined || inputDate === "") {
            return "ERROR";
        }

        let dateObj: Date;

        if (inputDate instanceof Date) {
            dateObj = inputDate;
        } else if (typeof inputDate === 'number') {
            const MS_PER_DAY = 24 * 60 * 60 * 1000;
            const excelEpoch = new Date(Date.UTC(1899, 11, 30));
            dateObj = new Date(excelEpoch.getTime() + inputDate * MS_PER_DAY);
        } else if (typeof inputDate === 'string') {
            const trimmed = inputDate.trim();
            if (trimmed === "") return "ERROR";

            let str = trimmed.replace(/[\s]+/g, ' ').replace(/[.]/g, '-').replace(/[\/]/g, '-');

            const monthNames = {
                jan: "01", feb: "02", mar: "03", apr: "04", may: "05", jun: "06",
                jul: "07", aug: "08", sep: "09", oct: "10", nov: "11", dec: "12"
            };

            const monthRegex = /(\b[a-zA-Z]{3,}\b)/;
            if (monthRegex.test(str)) {
                const parts = str.split(/[-\s]/);
                const mapped = parts.map(p => {
                    const lower = p.toLowerCase();
                    if (monthNames[lower as keyof typeof monthNames]) {
                        return monthNames[lower as keyof typeof monthNames];
                    }
                    return p;
                });
                str = mapped.join("-");
            }

            let isoMatch = str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
            if (isoMatch) {
                const [_, y, m, d] = isoMatch;
                dateObj = new Date(`${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`);
            } else if (/^(\d{1,2})-(\d{1,2})-(\d{4})$/.test(str)) {
                const [_, d, m, y] = str.match(/^(\d{1,2})-(\d{1,2})-(\d{4})$/)!;
                dateObj = new Date(`${y}-${m.padStart(2, '0')}-${d.padStart(2, '0')}`);
            } else {
                dateObj = new Date(str);
            }
        } else {
            return "ERROR";
        }

        if (isNaN(dateObj.getTime())) {
            return "ERROR";
        }

        const year = dateObj.getUTCFullYear();
        const month = String(dateObj.getUTCMonth() + 1).padStart(2, '0');
        const day = String(dateObj.getUTCDate()).padStart(2, '0');

        return `${year}-${month}-${day}`;
    } catch {
        return "ERROR";
    }
}

/**
 * Helper to extract rate from the payload array
 */
function extractRatesFromPayload(
    payload: any[], 
    targetCurrency: string, 
    rateType: string, 
    needsInversion: boolean
): any[][] {
    const results: any[][] = [];
    const rateField = rateType === "B" ? "buy" : "sell";

    for (const entry of payload) {
        const dateStr = entry.date;
        const rates = entry.rates;

        if (!Array.isArray(rates)) continue;

        for (const rateEntry of rates) {
            const iso3 = rateEntry.currency?.iso3 || rateEntry.iso3;
            
            if (iso3 === targetCurrency) {
                let rateValue = rateEntry[rateField];

                if (rateValue === null || rateValue === undefined || rateValue === "") continue;

                const rateStr = String(rateValue).replace(/,/g, '');
                const rateParsed = parseFloat(rateStr);

                if (isNaN(rateParsed) || rateParsed <= 0) continue;

                const finalRate = needsInversion ? (1 / rateParsed) : rateParsed;
                
                results.push([dateStr, finalRate]);
                break;
            }
        }
    }
    return results;
}

// =====================================================================================
// Custom Function Association
// =====================================================================================

if (typeof CustomFunctions !== "undefined") {
    CustomFunctions.associate("NRB.FOREXRATE", ForexRate);
    CustomFunctions.associate("NRB.TESTCONNECTION", TestConnection);
    CustomFunctions.associate("NRB.PARSEJSON", ParseJSON);
}