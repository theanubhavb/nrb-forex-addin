// =====================================================================================
// WINDOWS-ONLY CALCULATION CONTROL (SAFE FOR MAC)
// =====================================================================================
// This runs only in taskpane/commands, not in Custom Functions runtime.

if (typeof Office !== "undefined") {
  Office.onReady((info) => {
    try {
      if (info.platform === Office.PlatformType.PC) {
        Excel.run(async (context) => {
          try {
            context.application.calculationMode = Excel.CalculationMode.manual;
            await context.sync();
          } catch {}
        });
      }
    } catch {}
  });
}

// =====================================================================================
// CUSTOM FUNCTIONS RUNTIME STARTS HERE (MAC-SAFE)
// =====================================================================================
// Pure JS + fetch + JSON, no Excel.run

// =====================================================================================
// NAMESPACE: NRB
// =====================================================================================

/**
 * Test NRB API connection.
 * @customfunction
 * @volatile false
 * @returns {Promise<string>}
 * @alias NRB.TESTCONNECTION
 */
export async function TestConnection(): Promise<string> {
  const url = "https://www.nrb.org.np/api/forex/v1/rates?from=2024-01-01&to=2024-01-01&page=1&per_page=1";
  try {
    const response = await fetch(url);
    if (!response.ok) return `NETWORK ERROR: ${response.status} - ${response.statusText}`;
    
    const json = await response.json();
    const status = json.status?.code ?? json.status;
    if (status === 200 || status === 1) return "SUCCESS: NRB API reachable.";
    
    return `API WARNING: HTTP OK but unexpected status (${status}). Snippet: ${JSON.stringify(json).slice(0, 500)}...`;
  } catch (e) {
    return `ERROR: Could not complete request. (${e instanceof Error ? e.message : String(e)})`;
  }
}

/**
 * Get NRB Forex rate (single date or range).
 * @customfunction
 * @volatile false
 * @param {any} fromDate
 * @param {any} [toDate]
 * @param {string} [currencyPair="USDNPR"]
 * @param {string} [rateType="S"]
 * @returns {Promise<string|number>}
 * @alias NRB.FOREXRATE
 */
export async function ForexRate(
  fromDate: any,
  toDate?: any,
  currencyPair: string = "USDNPR",
  rateType: string = "S"
): Promise<string | number> {
  try {
    if (!fromDate) return "#ERROR: From Date is required";

    const fromStr = parseDate(fromDate);
    if (fromStr === "ERROR") return "#ERROR: Invalid From Date";

    let toStr = fromStr;
    let isRange = false;

    if (toDate !== undefined && toDate !== null && toDate !== "") {
      toStr = parseDate(toDate);
      if (toStr === "ERROR") return "#ERROR: Invalid To Date";
      isRange = fromStr !== toStr;
    }

    currencyPair = String(currencyPair).trim().toUpperCase();
    if (currencyPair.length !== 6) return "#ERROR: Invalid currency pair (e.g., USDNPR)";

    const needsInvert = currencyPair.startsWith("NPR");
    const targetCurrency = needsInvert ? currencyPair.slice(3) : currencyPair.slice(0, 3);

    rateType = String(rateType).trim().toUpperCase();
    if (!["B", "S"].includes(rateType)) return "#ERROR: Rate type must be B or S";

    // ================== RANGE ==================
    if (isRange) {
      const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${fromStr}&to=${toStr}&page=1&per_page=100`;
      const response = await fetch(url);
      if (!response.ok) return `#ERROR: HTTP ${response.status}`;
      const json = await response.json();
      const payload = json.data?.payload || json.data;
      if (!Array.isArray(payload) || payload.length === 0) return `#ERROR: No data for ${fromStr} to ${toStr}`;
      
      const rates = extractRates(payload, targetCurrency, rateType, needsInvert);
      if (rates.length === 0) return `#ERROR: Currency ${currencyPair} not found`;
      return JSON.stringify([["Date", "Rate"], ...rates]);
    }

    // ================== SINGLE DATE WITH 7-DAY BACKOFF ==================
    let current = new Date(fromStr);
    for (let i = 0; i < 7; i++) {
      const dateStr = formatDate(current);
      const url = `https://www.nrb.org.np/api/forex/v1/rates?from=${dateStr}&to=${dateStr}&page=1&per_page=1`;
      const response = await fetch(url);
      if (response.ok) {
        const json = await response.json();
        const payload = json.data?.payload || json.data;
        if (Array.isArray(payload) && payload.length > 0) {
          const rates = extractRates(payload, targetCurrency, rateType, needsInvert);
          if (rates.length > 0) return dateStr === fromStr ? rates[0][1] : rates[0][1] * -1;
        }
      }
      current.setDate(current.getDate() - 1);
    }

    return `#ERROR: No rate found for ${currencyPair} on ${fromStr} or previous 7 days`;

  } catch (err) {
    return `#ERROR: ${err instanceof Error ? err.message : String(err)}`;
  }
}

/**
 * Parse JSON string returned by ForexRate to Excel table
 * @customfunction
 * @volatile false
 * @param {string} jsonString
 * @returns {any[][]}
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
// UTILITIES
// =====================================================================================

function parseDate(input: any): string {
  if (input === null || input === undefined || input === "") return "ERROR";
  let d: Date;
  if (input instanceof Date) d = input;
  else if (typeof input === "number") {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    d = new Date(epoch.getTime() + input * 86400000);
  } else if (typeof input === "string") {
    d = new Date(input.replace(/[.\/]/g, "-").trim());
  } else return "ERROR";

  if (isNaN(d.getTime())) return "ERROR";
  return formatDate(d);
}

function formatDate(d: Date): string {
  return `${d.getUTCFullYear()}-${String(d.getUTCMonth() + 1).padStart(2, "0")}-${String(d.getUTCDate()).padStart(2, "0")}`;
}

function extractRates(payload: any[], target: string, rateType: string, invert: boolean): any[][] {
  const results: any[][] = [];
  const field = rateType === "B" ? "buy" : "sell";

  for (const entry of payload) {
    const date = entry.date;
    const rates = entry.rates;
    if (!Array.isArray(rates)) continue;

    for (const r of rates) {
      const iso3 = r.currency?.iso3 || r.iso3;
      if (iso3 === target) {
        const val = parseFloat(String(r[field]).replace(/,/g, ""));
        if (!isNaN(val) && val > 0) results.push([date, invert ? 1 / val : val]);
      }
    }
  }
  return results;
}

// =====================================================================================
// CUSTOM FUNCTION ASSOCIATION (MAC SAFE)
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

// =====================================================================================
// MAC: CUSTOM FUNCTIONS METADATA
// =====================================================================================

if (typeof self !== "undefined") {
  (self as any).__customFunctionsMetadata = {
    allowCustomDataForDataTypeAny: true,
    allowErrorForDataTypeAny: true
  };
}
