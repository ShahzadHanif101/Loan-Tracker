const PEOPLE_SHEET = "People";
const TX_SHEET = "Transactions";

function doGet() {
  return HtmlService.createTemplateFromFile("Dashboard")
    .evaluate()
    .setTitle("Loan Tracker")
    .addMetaTag("viewport", "width=device-width, initial-scale=1");
}

function include(file) {
  return HtmlService.createHtmlOutputFromFile(file).getContent();
}

// =========================================================
// PEOPLE
// =========================================================

function getPeople() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(PEOPLE_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return [];
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 2).getValues();

  return values
    .filter((row) => row[1] !== "")
    .map((row) => ({
      id: row[0],
      name: row[1],
    }));
}

// =========================================================
// GET NEXT TRANSACTION ID
// =========================================================
//
// IDs are NEVER reused.
//
// Example:
// TN-0001
// TN-0002
// TN-0003
//
// If TN-0002 is deleted,
// the next new transaction will still receive
// the next unused number.
//

function getNextTransactionId() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return "TN-0001";
  }

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  let highestNumber = 0;

  ids.forEach((id) => {
    const match = String(id).match(/^TN-(\d+)$/);

    if (match) {
      const number = Number(match[1]);

      if (number > highestNumber) {
        highestNumber = number;
      }
    }
  });

  return "TN-" + String(highestNumber + 1).padStart(4, "0");
}

// =========================================================
// SAVE NEW TRANSACTION
// =========================================================

function saveTransaction(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  // Generate a completely new ID
  const transactionId = getNextTransactionId();

  const transactionDate = new Date(data.date);

  transactionDate.setHours(0, 0, 0, 0);

  sheet.appendRow([
    transactionId,

    transactionDate,

    data.person,

    data.type,

    Number(data.amount),

    data.notes,
  ]);

  const newRow = sheet.getLastRow();

  // Date only
  sheet.getRange(newRow, 2).setNumberFormat("dd MMM yyyy");

  // Amount
  sheet.getRange(newRow, 5).setNumberFormat("#,##0.00");

  return {
    success: true,
    id: transactionId,
  };
}

// =========================================================
// GET TRANSACTIONS
// =========================================================

function getTransactions(person, fromDate, toDate) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return {
      transactions: [],

      totalGiven: 0,

      totalReturned: 0,

      balance: 0,
    };
  }

  // 6 columns:
  // ID | Date | Person | Type | Amount | Notes

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  const startDate = new Date(fromDate);

  startDate.setHours(0, 0, 0, 0);

  const endDate = new Date(toDate);

  endDate.setHours(23, 59, 59, 999);

  let transactions = [];

  let totalGiven = 0;

  let totalReturned = 0;

  values.forEach((row) => {
    const id = row[0];

    const date = row[1];

    const rowPerson = row[2];

    const type = row[3];

    const amount = Number(row[4]) || 0;

    const notes = row[5] || "";

    if (!(date instanceof Date)) {
      return;
    }

    // Person filter

    if (String(rowPerson) !== String(person)) {
      return;
    }

    // Date filter

    if (date < startDate || date > endDate) {
      return;
    }

    if (type === "Given") {
      totalGiven += amount;
    }

    if (type === "Returned") {
      totalReturned += amount;
    }

    transactions.push({
      id: String(id),

      date: Utilities.formatDate(
        date,
        Session.getScriptTimeZone(),
        "dd MMM yyyy",
      ),

      type: type,

      amount: amount,

      notes: String(notes),
    });
  });

  // Newest first

  transactions.reverse();

  return {
    transactions: transactions,

    totalGiven: totalGiven,

    totalReturned: totalReturned,

    balance: totalGiven - totalReturned,
  };
}

// =========================================================
// GET SINGLE TRANSACTION
// =========================================================

function getTransactionById(id) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet.getRange(2, 1, lastRow - 1, 6).getValues();

  for (let i = 0; i < values.length; i++) {
    const row = values[i];

    if (String(row[0]) === String(id)) {
      const date = row[1];

      return {
        id: String(row[0]),

        date: Utilities.formatDate(
          date,
          Session.getScriptTimeZone(),
          "yyyy-MM-dd",
        ),

        person: String(row[2]),

        type: String(row[3]),

        amount: Number(row[4]) || 0,

        notes: String(row[5] || ""),
      };
    }
  }

  return null;
}

// =========================================================
// UPDATE TRANSACTION
// =========================================================

function updateTransaction(data) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("Transaction not found.");
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let targetRow = -1;

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(data.id)) {
      targetRow = i + 2;

      break;
    }
  }

  if (targetRow === -1) {
    throw new Error("Transaction not found.");
  }

  const transactionDate = new Date(data.date);

  transactionDate.setHours(0, 0, 0, 0);

  // IMPORTANT:
  // ID stays exactly the same.

  sheet
    .getRange(targetRow, 2, 1, 5)
    .setValues([
      [
        transactionDate,

        data.person,

        data.type,

        Number(data.amount),

        data.notes,
      ],
    ]);

  // Date only

  sheet.getRange(targetRow, 2).setNumberFormat("dd MMM yyyy");

  // Amount

  sheet.getRange(targetRow, 5).setNumberFormat("#,##0.00");

  return true;
}

// =========================================================
// DELETE TRANSACTION
// =========================================================

function deleteTransaction(id) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    throw new Error("Transaction not found.");
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let targetRow = -1;

  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      targetRow = i + 2;

      break;
    }
  }

  if (targetRow === -1) {
    throw new Error("Transaction not found.");
  }

  // Delete the row.

  // The ID is NOT stored anywhere for reuse.
  // getNextTransactionId() always looks for
  // the highest number ever currently present.

  sheet.deleteRow(targetRow);

  return true;
}

function initializeTransactionCounter() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(TX_SHEET);

  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    PropertiesService.getScriptProperties().setProperty(
      "LAST_TRANSACTION_NUMBER",
      "0",
    );

    return;
  }

  const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();

  let highestNumber = 0;

  ids.forEach((row) => {
    const match = String(row[0]).match(/^TN-(\d+)$/);

    if (match) {
      const number = Number(match[1]);

      if (number > highestNumber) {
        highestNumber = number;
      }
    }
  });

  PropertiesService.getScriptProperties().setProperty(
    "LAST_TRANSACTION_NUMBER",
    String(highestNumber),
  );
}
