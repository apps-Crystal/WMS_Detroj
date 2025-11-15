function updatePalletTransactionLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const buildSheet = ss.getSheetByName("Pallet_Build_IB_04");
  const ledgerSheet = ss.getSheetByName("Pallet_Transaction_Ledger");

  const buildHeaders = buildSheet.getDataRange().getValues()[0];
  const ledgerHeaders = ledgerSheet.getDataRange().getValues()[0];

  // 🔍 Identify column indexes
  const grnIdCol = buildHeaders.indexOf("GRN_ID");
  const palletIdCol = buildHeaders.indexOf("Pallet_ID");
  const palletGRNCol = buildHeaders.indexOf("Pallet_GRN");
  const timestampCol = buildHeaders.indexOf("Timestamp");
  const skuIdCol = buildHeaders.indexOf("SKU_ID");
  const skuDescCol = buildHeaders.indexOf("SKU_Description");
  const batchCol = buildHeaders.indexOf("Batch_Number");
  const qtyCol = buildHeaders.indexOf("Quantity_Boxes");

  if (grnIdCol === -1 || palletIdCol === -1 || palletGRNCol === -1 || timestampCol === -1) {
    Logger.log("⚠️ Missing required column in Pallet_Build_IB_04");
    return;
  }

  // 🆕 Get last (newest) record
  const lastRow = buildSheet.getLastRow();
  const newRow = buildSheet.getRange(lastRow, 1, 1, buildHeaders.length).getValues()[0];

  // 📦 Extract values
  const grnId = newRow[grnIdCol];
  const palletId = newRow[palletIdCol];
  const palletGRN = newRow[palletGRNCol];
  const timestamp = newRow[timestampCol];
  const skuId = newRow[skuIdCol];
  const skuDesc = newRow[skuDescCol];
  const batchNo = newRow[batchCol];
  const qty = newRow[qtyCol];

  // 🧾 Prepare new row for Ledger
  const newLedgerRow = [];
  for (let i = 0; i < ledgerHeaders.length; i++) {
    const header = ledgerHeaders[i];
    switch (header) {
      case "Timestamp":
        newLedgerRow.push(timestamp || new Date());
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
}
