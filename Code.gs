function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧰 Bid Tools')
    .addItem('Generate Bid Sheet', 'generateBidSheet')
    .addToUi();
}

function generateBidSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const takeoff = ss.getSheetByName('Takeoff');
  const materials = ss.getSheetByName('Materials');
  const labor = ss.getSheetByName('Labor');
  const inputs = ss.getSheetByName('Inputs');
  const bid = ss.getSheetByName('Bid Sheet');

  // Clear existing data in Bid Sheet
  bid.clear();
  bid.appendRow(["Item Name", "Description", "Quantity", "Unit", "Material Cost", "Labor Cost", "Total"]);

  // Get input values
  const taxRate = inputs.getRange("B1").getValue() / 100;
  const margin = inputs.getRange("B2").getValue() / 100;
  const overhead = inputs.getRange("B3").getValue();

  const takeoffData = takeoff.getRange(2, 1, takeoff.getLastRow() - 1, 6).getValues();
  const materialsData = materials.getRange(2, 1, materials.getLastRow() - 1, 4).getValues();
  const laborData = labor.getRange(2, 1, labor.getLastRow() - 1, 3).getValues();

  let subtotal = 0;

  takeoffData.forEach(row => {
    const [item, desc, unit, netQty, waste, category] = row;
    const qty = netQty * (1 + (waste || 0) / 100);

    const matRow = materialsData.find(m => m[0] === item);
    const labRow = laborData.find(l => l[0] === item);

    const matRate = matRow ? matRow[2] : 0;
    const labRate = labRow ? labRow[2] : 0;

    const matCost = qty * matRate;
    const labCost = qty * labRate;
    const total = matCost + labCost;

    subtotal += total;

    bid.appendRow([item, desc, qty, unit, matCost, labCost, total]);
  });

  const tax = subtotal * taxRate;
  const marginAmt = (subtotal + tax + overhead) * margin;
  const grandTotal = subtotal + tax + overhead + marginAmt;

  bid.appendRow([]);
  bid.appendRow(["", "", "", "", "Subtotal", "", subtotal]);
  bid.appendRow(["", "", "", "", "Tax", "", tax]);
  bid.appendRow(["", "", "", "", "Overhead", "", overhead]);
  bid.appendRow(["", "", "", "", "Margin", "", marginAmt]);
  bid.appendRow(["", "", "", "", "Grand Total", "", grandTotal]);
}
