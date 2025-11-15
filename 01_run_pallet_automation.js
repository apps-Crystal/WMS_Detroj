function runPalletAutomation(){
  // 1. Run the transaction and get the ID of the newly processed pallet (or null if skipped/failed).
  const newPalletId = updatePalletTransactionLedger();
  
  // 2. Only run the sync if a new pallet transaction was actually recorded.
  if (newPalletId) {
    syncPalletStatus(newPalletId);
  } else {
    Logger.log("No new pallet added to Ledger (skipped due to duplicate or missing data). Skipping syncPalletStatus.");
  }
}

/**
 * Updates Pallet_Transaction_Ledger by extracting the latest record from 
 * Pallet_Build_IB_04. Includes a check to prevent duplicate "Built" entries.
 * * @returns {string | null} The Pallet_ID if a new transaction was recorded, otherwise null.
 */
function updatePalletTransactionLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildSheet = ss.getSheetByName("Pallet_Build_IB_04");
  const ledgerSheet = ss.getSheetByName("Pallet_Transaction_Ledger");

  // Get all data from both sheets to establish context
  const buildRange = buildSheet.getDataRange();
  const buildData = buildRange.getValues();
  const buildHeaders = buildData[0];
  
  const ledgerRange = ledgerSheet.getDataRange();
  const ledgerData = ledgerRange.getValues();
  const ledgerHeaders = ledgerData[0];

  // 🔍 Identify column indexes in Pallet_Build_IB_04
  const grnIdCol = buildHeaders.indexOf("GRN_ID");
  const palletIdCol = buildHeaders.indexOf("Pallet_ID");
  const palletGRNCol = buildHeaders.indexOf("Pallet_GRN");
  const timestampCol = buildHeaders.indexOf("Timestamp");
  const skuIdCol = buildHeaders.indexOf("SKU_ID");
  const skuDescCol = buildHeaders.indexOf("SKU_Description");
  const batchCol = buildHeaders.indexOf("Batch_Number");
  const qtyCol = buildHeaders.indexOf("Quantity_Boxes");

  if (grnIdCol === -1 || palletIdCol === -1 || palletGRNCol === -1 || timestampCol === -1) {
    Logger.log("⚠️ Missing required column in Pallet_Build_IB_04: GRN_ID, Pallet_ID, Pallet_GRN, or Timestamp.");
    return null;
  }

  // 🆕 Get last (newest) record from Pallet_Build_IB_04
  const lastRowIndex = buildData.length - 1;
  if (lastRowIndex <= 0) {
    Logger.log("⚠️ Pallet_Build_IB_04 appears empty or has only headers.");
    return null;
  }
  const newRow = buildData[lastRowIndex];

  // 📦 Extract values
  const grnId = newRow[grnIdCol];
  const palletId = newRow[palletIdCol];
  const palletGRN = newRow[palletGRNCol];
  const timestamp = newRow[timestampCol];
  const skuId = newRow[skuIdCol];
  const skuDesc = newRow[skuDescCol];
  const batchNo = newRow[batchCol];
  const qty = newRow[qtyCol];

  /***************************************************
   * DUPLICATE CHECK: Prevent adding the same "Built" transaction twice
   ***************************************************/
  const L_ACTION_COL = ledgerHeaders.indexOf("Action_Type");
  const L_PALLET_COL = ledgerHeaders.indexOf("Pallet_ID");
  const L_GRN_COL = ledgerHeaders.indexOf("GRN_ID");
  
  let isDuplicate = false;

  if (L_ACTION_COL !== -1 && L_PALLET_COL !== -1 && L_GRN_COL !== -1) {
    // Iterate from 1 to skip headers
    for (let i = 1; i < ledgerData.length; i++) { 
      const row = ledgerData[i];
      // Compare the key fields for duplication
      if (
        row[L_ACTION_COL] === "Built" &&
        // Use loose comparison (==) as one might be string and the other number
        row[L_PALLET_COL] == palletId && 
        row[L_GRN_COL] == grnId
      ) {
        isDuplicate = true;
        break;
      }
    }
  } else {
    Logger.log("⚠️ Could not perform duplicate check: Missing Action_Type, Pallet_ID, or GRN_ID in Ledger headers. Inserting without check.");
  }

  if (isDuplicate) {
    Logger.log(`🚫 SKIPPED: Duplicate "Built" transaction already found for Pallet: ${palletId} and GRN: ${grnId}`);
    return null; // Return null on skip
  }
  /***************************************************
   * END DUPLICATE CHECK
   ***************************************************/


  // 🧾 Prepare new row for Ledger
  const newLedgerRow = [];
  for (let i = 0; i < ledgerHeaders.length; i++) {
    const header = ledgerHeaders[i];
    switch (header) {
      case "Timestamp":
        // Ensure timestamp is a valid Date object for logging if null
        newLedgerRow.push(timestamp instanceof Date ? timestamp : (timestamp || new Date()));
        break;
      case "Action_Type":
        newLedgerRow.push("Built");
        break;
      case "Pallet ID_GRN ID":
        newLedgerRow.push(palletGRN);
        break;
      case "Pallet_ID":
        newLedgerRow.push(palletId);
        break;
      case "GRN_ID":
        newLedgerRow.push(grnId);
        break;
      case "SKU_ID":
        newLedgerRow.push(skuId || "");
        break;
      case "SKU_Description":
        newLedgerRow.push(skuDesc || "");
        break;
      case "Batch_No":
        newLedgerRow.push(batchNo || "");
        break;
      case "Qty_Change":
        newLedgerRow.push(qty || 0);
        break;
      case "Status":
        newLedgerRow.push("Ready For Putaway");
        break;
      default:
        newLedgerRow.push("");
    }
  }

  // 🪄 Insert row into Ledger
  const insertRow = ledgerSheet.getLastRow() + 1;
  ledgerSheet.getRange(insertRow, 1, 1, newLedgerRow.length).setValues([newLedgerRow]);

  Logger.log(`✅ Ledger Updated | GRN: ${grnId} | Pallet: ${palletId} | Status: Ready For Putaway`);
  return palletId; // Return the Pallet ID on successful insertion
}


/**********************************************************************
 * PALLET STATUS SYNCHRONIZATION ENGINE
 * This version is optimized to only update the Pallet_Status_02 row
 * for the specified targetPalletId.
 **********************************************************************/

/**
 * Syncs the status of a SINGLE target pallet from Ledger and Build sheets 
 * to the Status sheet.
 * * @param {string} targetPalletId The Pallet_ID to be updated.
 */
function syncPalletStatus(targetPalletId) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ledgerSheet = ss.getSheetByName("Pallet_Transaction_Ledger");
  const buildSheet  = ss.getSheetByName("Pallet_Build_IB_04");
  const statusSheet = ss.getSheetByName("Pallet_Status_02");

  // Load all data (headers + rows)
  const ledger = ledgerSheet.getDataRange().getValues();
  const build  = buildSheet.getDataRange().getValues();
  const status = statusSheet.getDataRange().getValues();

  // -------- LOGGER ----------
  let LOG = "";
  function log(msg) {
    LOG += msg + "\n";
    Logger.log(msg);
  }

  log("🚀 START — Pallet Status Sync Engine (Targeted)");
  log("==============================================");

  if (!targetPalletId) {
      log("🚫 Aborting Sync: No target Pallet ID provided.");
      return;
  }
  log(`🎯 TARGET SYNC: Processing only Pallet ID: ${targetPalletId}`);

  /***************************************************
   * REAL COLUMN MAPPING (BASED ON YOUR SCREENSHOTS)
   ***************************************************/
  
  // Pallet_Transaction_Ledger
  const LEDGER_COL = {
    TIMESTAMP: 0,        // A
    ACTION_TYPE: 1,      // B
    PALLET_ID_GRN: 2,    // C (not needed)
    PALLET_ID: 3,        // D
    GRN_ID: 4,           // E
    DN_ID: 5,            // F
    SKU_ID: 6,           // G
    SKU_DESCRIPTION: 7,  // H
    BATCH_NUMBER: 8,     // I
    QTY_CHANGE: 9,       // J
    STATUS: 10           // K (Assuming Status is the 11th column/Index 10)
  };

  // Pallet_Build_IB_04
  const BUILD_COL = {
    PALLET_ID: 3,      // D
    EXPIRY_DATE: 10    // K
  };

  // Pallet_Status_02
  const STATUS_COL = {
    PALLET_ID: 0,           // A
    OCCUPANCY_STATUS: 1,    // B
    GRN_ID: 2,              // C
    SKU_ID: 3,              // D
    SKU_DESCRIPTION: 4,     // E
    EXPIRY_DATE: 5,         // F
    BATCH_NUMBER: 6,        // G
    LOCATION_ID: 7,         // H
    CURRENT_QTY: 8,          // I
    LAST_UPDATED: 9,         // J
    ASSIGNMENT_STATUS: 10    // K
  };

  /***************************************************
   * BUILD FAST LOOKUP MAP OF STATUS SHEET
   ***************************************************/
  let statusRowByID = {};

  // Build the map based on the Pallet_ID column (Index 0)
  for (let i = 1; i < status.length; i++) {
    let palletID = status[i][STATUS_COL.PALLET_ID];
    if (palletID) {
      statusRowByID[palletID] = i; // Store the row index
    }
  }

  log("📌 Loaded Status Sheet Pallets: " + Object.keys(statusRowByID).length);
  const sRow = statusRowByID[targetPalletId];

  if (sRow === undefined) {
      log(`⚠️ ABORT — Target Pallet ID ${targetPalletId} not found in Pallet_Status_02. Check data or allocation logic.`);
      return;
  }
  
  log(`Found target Pallet ID ${targetPalletId} at row index ${sRow}.`);


  /***************************************************
   * 1) UPDATE FROM LEDGER (Find the latest transaction)
   ***************************************************/
  log("\n📘 STEP 1 — Applying Ledger Transactions...");
  let ledgerUpdates = 0;

  // Find the LAST transaction for the targetPalletId in the ledger
  let lastTransactionRow = null;
  // Iterate backwards to find the latest (newest) record first
  for (let r = ledger.length - 1; r >= 1; r--) {
    // Loose comparison (==) handles cases where one is string and the other is number
    if (ledger[r][LEDGER_COL.PALLET_ID] == targetPalletId) { 
        lastTransactionRow = ledger[r];
        break; // Found the latest one, stop searching
    }
  }

  if (lastTransactionRow) {
    let row = lastTransactionRow;
    
    // Apply the latest Ledger data to the Status row
    status[sRow][STATUS_COL.SKU_ID]          = row[LEDGER_COL.SKU_ID];
    status[sRow][STATUS_COL.SKU_DESCRIPTION] = row[LEDGER_COL.SKU_DESCRIPTION];
    status[sRow][STATUS_COL.BATCH_NUMBER]    = row[LEDGER_COL.BATCH_NUMBER];
    status[sRow][STATUS_COL.CURRENT_QTY]     = row[LEDGER_COL.QTY_CHANGE];
    status[sRow][STATUS_COL.GRN_ID]          = row[LEDGER_COL.GRN_ID];
    status[sRow][STATUS_COL.LAST_UPDATED]    = row[LEDGER_COL.TIMESTAMP];

    // Determine Occupancy and Assignment Status based on the transaction type
    const action = row[LEDGER_COL.ACTION_TYPE];
    if (action === "Built" || action === "Received" || action === "Putaway") {
        status[sRow][STATUS_COL.OCCUPANCY_STATUS] = "✅ Occupied";
        status[sRow][STATUS_COL.ASSIGNMENT_STATUS] = "Unassigned"; 
    } else if (action === "Shipped" || action === "Empty" || row[LEDGER_COL.QTY_CHANGE] == 0) {
         status[sRow][STATUS_COL.OCCUPANCY_STATUS] = "❌ Empty";
         status[sRow][STATUS_COL.ASSIGNMENT_STATUS] = "N/A";
         // Clear relevant details for an empty pallet
         status[sRow][STATUS_COL.SKU_ID]          = "";
         status[sRow][STATUS_COL.SKU_DESCRIPTION] = "";
         status[sRow][STATUS_COL.BATCH_NUMBER]    = "";
         status[sRow][STATUS_COL.CURRENT_QTY]     = 0;
         status[sRow][STATUS_COL.GRN_ID]          = "";
         status[sRow][STATUS_COL.EXPIRY_DATE]     = "";
         status[sRow][STATUS_COL.LOCATION_ID]     = ""; 
    }
    ledgerUpdates = 1;
    log(`🔄 Updated Status from Ledger for Pallet: ${targetPalletId} | Action: ${action}`);
  } else {
    log(`⚠️ SKIPPED — Target Pallet ID ${targetPalletId} not found in Ledger (should not happen if updateLedger was successful).`);
  }

  log("\n📘 Ledger Updates Completed: " + ledgerUpdates);

  /***************************************************
   * 2) UPDATE EXPIRY FROM BUILD SHEET (Find the latest build record)
   ***************************************************/
  log("\n📗 STEP 2 — Applying Expiry Date from Build...");
  let expiryUpdates = 0;

  // Find the LAST row in the Build sheet for the targetPalletId
  let lastBuildRow = null;
  for (let r = build.length - 1; r >= 1; r--) {
    if (build[r][BUILD_COL.PALLET_ID] == targetPalletId) {
        lastBuildRow = build[r];
        break;
    }
  }

  if (lastBuildRow) {
    let row = lastBuildRow;

    let currentExpiry = status[sRow][STATUS_COL.EXPIRY_DATE];
    let newExpiry = row[BUILD_COL.EXPIRY_DATE];

    // Only update if the expiry date is available and different
    if (newExpiry) {
        // We always take the latest expiry from the Build sheet if it exists, 
        // regardless of what's in Status, as Build is the source of truth for the SKU attributes.
        status[sRow][STATUS_COL.EXPIRY_DATE] = newExpiry;
        log(`🟢 Expiry Updated → ${targetPalletId} = ${newExpiry}`);
        expiryUpdates = 1;
    } else {
         log(`➖ Expiry data was missing for ${targetPalletId} in the Build sheet.`);
    }
  } else {
     log(`⚠️ SKIPPED — Target Pallet ID ${targetPalletId} not found in Build sheet.`);
  }

  log("\n📗 Build Expiry Updates Completed: " + expiryUpdates);

  /***************************************************
   * WRITE BACK TO SHEET (Only write back the modified rows if necessary)
   ***************************************************/
  if (ledgerUpdates > 0 || expiryUpdates > 0) {
      log("\n💾 Writing Updated Row Back to Pallet_Status_02...");
      
      // We only write back the range that corresponds to the single updated row.
      // The updated row is status[sRow], which is at row index sRow + 1 in the sheet (since sRow is 1-indexed relative to data array, 0 is header)
      statusSheet
        .getRange(sRow + 1, 1, 1, status[0].length)
        .setValues([status[sRow]]);
      
      log("✅ WRITE COMPLETE for single row.");
  } else {
      log("✏️ No changes were applied to Pallet_Status_02. No write needed.");
  }


  /***************************************************
   * SUMMARY
   ***************************************************/
  log("\n🎯 SUMMARY");
  log(`• Target Pallet: ${targetPalletId}`);
  log("• Ledger Updates Applied: " + ledgerUpdates);
  log("• Expiry Updates Applied: " + expiryUpdates);

  log("\n🏁 DONE — Pallet Status Sync Completed Successfully");
}