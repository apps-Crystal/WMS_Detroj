/**********************************************************************
 * PALLET STATUS SYNCHRONIZATION ENGINE (FINAL VERSION)
 * Updates Pallet_Status_02 using:
 * 1) Pallet_Transaction_Ledger
 * 2) Pallet_Build_IB_04
 * Key = Pallet_ID
 **********************************************************************/

function syncPalletStatus() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const ledgerSheet = ss.getSheetByName("Pallet_Transaction_Ledger");
  const buildSheet  = ss.getSheetByName("Pallet_Build_IB_04");
  const statusSheet = ss.getSheetByName("Pallet_Status_02");

  const ledger = ledgerSheet.getDataRange().getValues();
  const build  = buildSheet.getDataRange().getValues();
  const status = statusSheet.getDataRange().getValues();

  // -------- LOGGER ----------
  let LOG = "";
  function log(msg) {
    LOG += msg + "\n";
    Logger.log(msg);
  }

  log("🚀 START — Pallet Status Sync Engine");
  log("===================================");

  /***************************************************
   * REAL COLUMN MAPPING (BASED ON YOUR SCREENSHOTS)
   ***************************************************/
  
  // Pallet_Transaction_Ledger
  const LEDGER_COL = {
    TIMESTAMP: 0,        // A
    ACTION_TYPE: 1,      // B
    PALLET_ID_GRN: 2,    // C (not needed)
    PALLET_ID: 3,        // D
    GRN_ID: 4,           // E
    DN_ID: 5,            // F
    SKU_ID: 6,           // G
    SKU_DESCRIPTION: 7,  // H
    BATCH_NUMBER: 8,     // I
    QTY_CHANGE: 9        // J
  };

  // Pallet_Build_IB_04
  const BUILD_COL = {
    PALLET_ID: 3,      // D
    EXPIRY_DATE: 10    // K
  };

  // Pallet_Status_02
  const STATUS_COL = {
    PALLET_ID: 0,           // A
    OCCUPANCY_STATUS: 1,    // B
    GRN_ID: 2,              // C
    SKU_ID: 3,              // D
    SKU_DESCRIPTION: 4,     // E
    EXPIRY_DATE: 5,         // F
    BATCH_NUMBER: 6,        // G
    LOCATION_ID: 7,         // H
    CURRENT_QTY: 8,         // I
    LAST_UPDATED: 9         // J
  };

  /***************************************************
   * BUILD FAST LOOKUP MAP OF STATUS SHEET
   ***************************************************/
  let statusRowByID = {};

  for (let i = 1; i < status.length; i++) {
    let palletID = status[i][STATUS_COL.PALLET_ID];
    if (palletID) {
      statusRowByID[palletID] = i;
    }
  }

  log("📌 Loaded Status Sheet Pallets: " + Object.keys(statusRowByID).length);

  /***************************************************
   * 1) UPDATE FROM LEDGER
   ***************************************************/
  log("\n📘 STEP 1 — Applying Ledger Transactions...");
  let ledgerUpdates = 0;

  for (let r = 1; r < ledger.length; r++) {

    let row = ledger[r];
    let palletID = row[LEDGER_COL.PALLET_ID];

    if (!palletID) continue;

    let sRow = statusRowByID[palletID];

    if (!sRow) {
      log("⚠️ SKIPPED — Pallet_ID not in Status: " + palletID);
      continue;
    }

    log("-------------------------------------");
    log("🔄 Updating Pallet_ID: " + palletID);

    status[sRow][STATUS_COL.SKU_ID]          = row[LEDGER_COL.SKU_ID];
    status[sRow][STATUS_COL.SKU_DESCRIPTION] = row[LEDGER_COL.SKU_DESCRIPTION];
    status[sRow][STATUS_COL.BATCH_NUMBER]    = row[LEDGER_COL.BATCH_NUMBER];
    status[sRow][STATUS_COL.CURRENT_QTY]     = row[LEDGER_COL.QTY_CHANGE];
    status[sRow][STATUS_COL.GRN_ID]          = row[LEDGER_COL.GRN_ID];
    status[sRow][STATUS_COL.LAST_UPDATED]    = row[LEDGER_COL.TIMESTAMP];

    // Mark pallet as occupied
    status[sRow][STATUS_COL.OCCUPANCY_STATUS] = "✅ Occupied";

    log("  ✔ SKU: " + row[LEDGER_COL.SKU_ID]);
    log("  ✔ Batch: " + row[LEDGER_COL.BATCH_NUMBER]);
    log("  ✔ Qty: " + row[LEDGER_COL.QTY_CHANGE]);
    log("  ✔ GRN: " + row[LEDGER_COL.GRN_ID]);
    log("  ✔ Occupancy = Occupied");

    ledgerUpdates++;
  }

  log("\n📘 Ledger Updates Completed: " + ledgerUpdates);

  /***************************************************
   * 2) UPDATE EXPIRY FROM BUILD SHEET
   ***************************************************/
  log("\n📗 STEP 2 — Applying Expiry Date from Build...");
  let expiryUpdates = 0;

  for (let r = 1; r < build.length; r++) {

    let row = build[r];
    let palletID = row[BUILD_COL.PALLET_ID];

    if (!palletID) continue;

    let sRow = statusRowByID[palletID];

    if (!sRow) {
      log("⚠️ SKIPPED — Pallet_ID not in Status (Build): " + palletID);
      continue;
    }

    status[sRow][STATUS_COL.EXPIRY_DATE] = row[BUILD_COL.EXPIRY_DATE];

    log("🟢 Expiry Updated → " + palletID + " = " + row[BUILD_COL.EXPIRY_DATE]);

    expiryUpdates++;
  }

  log("\n📗 Build Expiry Updates Completed: " + expiryUpdates);

  /***************************************************
   * WRITE BACK TO SHEET
   ***************************************************/
  log("\n💾 Writing Updated Data Back to Pallet_Status_02...");

  statusSheet
    .getRange(1, 1, status.length, status[0].length)
    .setValues(status);

  log("✅ WRITE COMPLETE");

  /***************************************************
   * SUMMARY
   ***************************************************/
  log("\n🎯 SUMMARY");
  log("• Ledger Updates: " + ledgerUpdates);
  log("• Expiry Updates: " + expiryUpdates);
  log("• Total Updates: " + (ledgerUpdates + expiryUpdates));

  log("\n🏁 DONE — Pallet Status Sync Completed Successfully");
}
