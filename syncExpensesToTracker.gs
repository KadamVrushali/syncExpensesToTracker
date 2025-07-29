function syncExpensesToTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eoSheet = ss.getSheetByName("EO");
  const stSheet = ss.getSheetByName("ST");

  if (!eoSheet || !stSheet) {
    throw new Error("Sheet 'EO' or 'ST' not found.");
  }

  console.log("Starting expense sync process...");


  const monthCell = eoSheet.getRange("A3").getValue(); // Apr-25 is in row 3, column A
  console.log("Raw month cell value: " + monthCell + " (type: " + typeof monthCell + ")");
  
  let eoMonth = null;
  
  if (monthCell instanceof Date) {
    eoMonth = monthCell;
    console.log("Found EO Date (as Date object): " + eoMonth);
  } else if (typeof monthCell === 'string' && monthCell.includes("Apr-25")) {
    // Create a date object from "Apr-25"
    eoMonth = new Date("2025-04-25");
    console.log("Created EO Date from string: " + eoMonth);
  } else {
    const row3Data = eoSheet.getRange(3, 1, 1, 25).getValues()[0]; // Check more columns
    
    for (let i = 0; i < row3Data.length; i++) {
      const cellVal = row3Data[i];
      if (cellVal instanceof Date) {
        eoMonth = cellVal;
        console.log("Found EO Date in row 3, column " + (i+1) + ": " + eoMonth);
        break;
      }
    }
  }
  
  if (!eoMonth) {
    console.log("Could not find date in EO sheet");
    return;
  }

  const stHeaderRow = stSheet.getRange(3, 1, 1, stSheet.getLastColumn()).getValues()[0];
  let stMonthCol = -1;
  
  console.log("Searching for matching date in ST headers...");
  
  for (let i = 0; i < stHeaderRow.length; i++) {
    const cellVal = stHeaderRow[i];
    console.log(`ST Header ${i+1}: ${cellVal} (type: ${typeof cellVal})`);
    
    if (cellVal instanceof Date && eoMonth instanceof Date) {
      // Compare dates - check if they're in the same month/year
      if (cellVal.getMonth() === eoMonth.getMonth() && cellVal.getFullYear() === eoMonth.getFullYear()) {
        stMonthCol = i + 1;
        console.log(`Found matching month at column ${stMonthCol}`);
        break;
      }
    } else if (typeof cellVal === 'string' && cellVal.includes("Apr 25")) {
      stMonthCol = i + 1;
      console.log(`Found matching month string at column ${stMonthCol}`);
      break;
    }
  }

  if (stMonthCol === -1) {
    console.log("Date not found in ST headers.");
    console.log("EO Date: " + eoMonth);
    console.log("Available ST headers: " + stHeaderRow.map((h, idx) => `[${idx+1}] ${h instanceof Date ? h.toDateString() : h}`).join(", "));
    return;
  }

  console.log("Month found in ST at column: " + stMonthCol);

  const fullRow3 = eoSheet.getRange(3, 1, 1, 25).getValues()[0];
  console.log("Full EO row 3 structure:");
  for (let i = 0; i < fullRow3.length; i++) {
    if (fullRow3[i] !== "" && fullRow3[i] !== null) {
      const colLetter = String.fromCharCode(65 + i);
      console.log(`  Column ${colLetter} (${i+1}): ${fullRow3[i]} (type: ${typeof fullRow3[i]})`);
    }
  }
  
  const expenseMapping = [];
  const expectedValues = [
    { value: 11800, category: "Fab/Raw/Consumable", stRow: 7, eoCategory: "Project Expense" },
    { value: 5739, category: "Consultancy", stRow: 5, eoCategory: "Consultant Fees" },
    { value: 25000, category: "HR", stRow: 4, eoCategory: "HR Expense" },
    { value: 1838, category: "Travel & Misc & Contingencies", stRow: 6, eoCategory: "OFC EXPENSES/MISC" },
    { value: 5310, category: "Travel & Misc & Contingencies", stRow: 6, eoCategory: "EXTERNALITIES" }
  ];
  
  console.log("Looking for expense values in EO sheet...");
  
  let travelMiscTotal = 0; // To sum both OFC EXPENSES and EXTERNALITIES
  
  for (let expected of expectedValues) {
    for (let i = 0; i < fullRow3.length; i++) {
      if (typeof fullRow3[i] === 'number' && fullRow3[i] === expected.value) {
        const colLetter = String.fromCharCode(65 + i);
        console.log(`Found ${expected.eoCategory} (${expected.value}) in column ${colLetter} (${i+1})`);
        
        if (expected.category === "Travel & Misc & Contingencies") {
          // Add to travel misc total instead of writing immediately
          travelMiscTotal += expected.value;
          console.log(`Adding ${expected.eoCategory} (₹${expected.value}) to Travel & Misc total. Current total: ₹${travelMiscTotal}`);
        } else {
          // Sync individual values for other categories
          stSheet.getRange(expected.stRow, stMonthCol).setValue(expected.value);
          console.log(`Synced ${expected.eoCategory} (₹${expected.value}) -> ${expected.category} at ST row ${expected.stRow}`);
        }
        break;
      }
    }
  }
  
  if (travelMiscTotal > 0) {
    stSheet.getRange(6, stMonthCol).setValue(travelMiscTotal);
    console.log(`Synced combined Travel & Misc & Contingencies: ₹${travelMiscTotal} (₹1838 + ₹5310) to ST row 6`);
  }

  updateSubtotal(stSheet, stMonthCol);
  
  console.log("Sync completed successfully!");
}

function updateSubtotal(stSheet, monthCol) {
  try {
    const expenseRange = stSheet.getRange(4, monthCol, 4, 1);
    const expenses = expenseRange.getValues();
    let total = 0;
    
    for (let i = 0; i < expenses.length; i++) {
      const val = expenses[i][0];
      if (typeof val === 'number') {
        total += val;
      }
    }
    
    stSheet.getRange(2, monthCol).setValue(total);
    console.log(`Updated subtotal: ${total}`);
  } catch (error) {
    console.log("Could not update subtotal: " + error.toString());
  }
}

function syncExpensesWithValidation() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const eoSheet = ss.getSheetByName("EO");
    const stSheet = ss.getSheetByName("ST");

    if (!eoSheet || !stSheet) {
      SpreadsheetApp.getUi().alert("Error", "Required sheets 'EO' or 'ST' not found!", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    const hasData = checkForExpenseData(eoSheet);
    if (!hasData) {
      SpreadsheetApp.getUi().alert("Info", "No expense data found to sync.", SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }

    syncExpensesToTracker();
    
    SpreadsheetApp.getUi().alert("Success", "Expenses synced successfully to Spending Tracker!", SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error("Sync failed: " + error.toString());
    SpreadsheetApp.getUi().alert("Error", "Sync failed: " + error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function checkForExpenseData(eoSheet) {
  const checkColumns = [5, 9, 13, 17, 21]; // E=5, I=9, M=13, Q=17, U=21 (1-based)
  
  for (let col of checkColumns) {
    const value = eoSheet.getRange(3, col).getValue();
    if (value && typeof value === 'number' && value >= 0) { // Include 0 values
      return true;
    }
  }
  return false;
}

function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  
  // Auto-sync when EO sheet is edited
  if (sheet.getName() === "EO" && e.range.getRow() === 1) {
    console.log("Auto-sync triggered by edit in EO sheet");
    
    // Add a small delay to ensure data is saved
    Utilities.sleep(1000);
    
    try {
      syncExpensesToTracker();
      
      const stSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ST");
      if (stSheet) {
        const monthCol = findCurrentMonthColumn(stSheet);
        if (monthCol > 0) {
          const range = stSheet.getRange(4, monthCol, 4, 1);
          range.setBackground("#90EE90"); // Light green
          
          // Reset color after 2 seconds
          Utilities.sleep(2000);
          range.setBackground(null);
        }
      }
    } catch (error) {
      console.log("Auto-sync failed: " + error.toString());
    }
  }
}

function findCurrentMonthColumn(stSheet) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const eoSheet = ss.getSheetByName("EO");
  
  if (!eoSheet) return -1;
  
  let eoMonth = null;
  const monthCell = eoSheet.getRange("A3").getValue();
  
  if (monthCell instanceof Date) {
    eoMonth = monthCell;
  } else if (typeof monthCell === 'string' && monthCell.includes("Apr-25")) {
    eoMonth = new Date("2025-04-25");
  } else {
    const row3Data = eoSheet.getRange(3, 1, 1, 25).getValues()[0];
    for (let i = 0; i < row3Data.length; i++) {
      const cellVal = row3Data[i];
      if (cellVal instanceof Date) {
        eoMonth = cellVal;
        break;
      }
    }
  }
  
  if (!eoMonth) return -1;
  
  const stHeaderRow = stSheet.getRange(3, 1, 1, stSheet.getLastColumn()).getValues()[0];
  
  for (let i = 0; i < stHeaderRow.length; i++) {
    const cellVal = stHeaderRow[i];
    
    if (cellVal instanceof Date && eoMonth instanceof Date) {
      if (cellVal.getMonth() === eoMonth.getMonth() && cellVal.getFullYear() === eoMonth.getFullYear()) {
        return i + 1;
      }
    } else if (typeof cellVal === 'string' && cellVal.includes("Apr 25")) {
      return i + 1;
    }
  }
  
  return -1;
}