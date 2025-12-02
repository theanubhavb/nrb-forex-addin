/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Define the name of the hidden item used to force recalculation.
const API_REFRESH_TICKER = "API_REFRESH_TICKER";

// The initialize function must be run each time a new page is loaded
Office.onReady(async () => {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    
    // 1. Initialize the hidden trigger value
    await initializeApiTicker();

    // 2. Attach the new manual refresh handler to the 'run' button
    document.getElementById("run").onclick = handleRefreshButtonClick;
});

/**
 * Initializes the hidden Named Item that will be used to track the refresh count.
 * Custom Functions will be made dependent on this value to control recalculation.
 */
async function initializeApiTicker() {
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            
            // Try to get the Named Item.
            let namedItem = workbook.names.getItemOrNullObject(API_REFRESH_TICKER);
            namedItem.load("value");
            await context.sync();

            // If the Named Item doesn't exist, create it with a starting value of 1.
            if (namedItem.isNullObject) {
                // The value '1' is stored as the initial trigger count.
                workbook.names.add(API_REFRESH_TICKER, 1);
                await context.sync();
                console.log("Initialized API_REFRESH_TICKER named item.");
            } else {
                console.log("API_REFRESH_TICKER already initialized.");
            }
        });
    } catch (error) {
        console.error("Error during API Ticker initialization:", error);
    }
}

/**
 * Handles the button click: Increments the hidden API_REFRESH_TICKER value.
 * Changing this value forces all dependent Custom Functions (ForexRate, TestConnection) to recalculate.
 */
export async function handleRefreshButtonClick() {
    try {
        await Excel.run(async (context) => {
            const workbook = context.workbook;
            
            // Get the Named Item that controls the recalculation.
            let namedItem = workbook.names.getItem(API_REFRESH_TICKER);
            namedItem.load("value");
            await context.sync();

            // Read, increment, and set the new value.
            let currentTicker = namedItem.value as number;
            let newTicker = currentTicker + 1;
            
            // Setting the new value forces the recalculation.
            namedItem.value = newTicker;
            
            await context.sync();
            console.log(`Updated refresh ticker to: ${newTicker}. Functions should now recalculate.`);
        });
    } catch (error) {
        console.error("Error during manual refresh:", error);
    }
}