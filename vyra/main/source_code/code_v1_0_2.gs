//FIX - The suppliers' names have been fixed by making the first letter a capital, which fixes the supplier payment issues.

const SPREADSHEET_ID = "17gncuj2eLxYgtc7PJm5SzT63U5dldWs8JIspQuyutR0";

const SHEET_NAMES = {
  transactions: "Transactions",
  lists: "Lists",
  openingBalances: "Opening Balances"
};

// function doGet() {
//   return HtmlService.createHtmlOutputFromFile("Index")
//     .setTitle("VYRA Kitchen — Finance")
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }

// function doGet() {
//   return HtmlService.createHtmlOutputFromFile("Index")
//     .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
//     .setTitle("VYRA Kitchen — Finance")
//     .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
// }

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .addMetaTag("viewport", "width=device-width, initial-scale=1, viewport-fit=cover")
    .setTitle("VYRA Kitchen — Finance")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function supplierKey_(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLocaleLowerCase("en-IN");
}
function formatSupplierName_(value) {
  return String(value || "").trim().replace(/\s+/g, " ").split(" ").filter(Boolean)
    .map(word => word.charAt(0).toLocaleUpperCase("en-IN") + word.slice(1).toLocaleLowerCase("en-IN")).join(" ");
}
function getCanonicalSupplierName_(value) {
  const display = formatSupplierName_(value);
  if (!display) return "";
  const sheet = getSheet_(SHEET_NAMES.lists), values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const col = headers.findIndex(h => String(h).trim().toLowerCase() === "suppliers");
  if (col < 0) throw new Error('The Lists sheet must contain a "Suppliers" column.');
  const key = supplierKey_(display);
  const match = values.slice(1).map(row => String(row[col] || "").trim()).find(existing => supplierKey_(existing) === key);
  return match ? formatSupplierName_(match) : display;
}

function getSpreadsheet_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === "PASTE_YOUR_GOOGLE_SHEET_ID_HERE") {
    throw new Error("Paste your Google Sheet ID into the SPREADSHEET_ID constant at the top of Code.gs.");
  }
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getSheet_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (!sheet) throw new Error("Required sheet not found: " + name);
  return sheet;
}

// function getInitialData() {
//   const ss = getSpreadsheet_();
//   const txSheet = getSheet_(SHEET_NAMES.transactions);
//   const listSheet = getSheet_(SHEET_NAMES.lists);

//   const txValues = txSheet.getDataRange().getValues();
//   const headers = txValues.shift().map(String);
//   const transactions = txValues
//     .filter(row => row.some(value => value !== "" && value !== null))
//     .map(row => rowToTransaction_(headers, row));

//   const listValues = listSheet.getDataRange().getDisplayValues();
//   const listHeaders = listValues[0] || [];
//   const lists = {};
//   listHeaders.forEach((header, index) => {
//     lists[header] = listValues.slice(1).map(row => row[index]).filter(Boolean);
//   });

//   return {
//     spreadsheetName: ss.getName(),
//     transactions: transactions,
//     lists: lists,
//     serverTime: new Date().toISOString()
//   };
// }

function getInitialData() {
  const ss = getSpreadsheet_();
  const txSheet = getSheet_(SHEET_NAMES.transactions);
  const listSheet = getSheet_(SHEET_NAMES.lists);

  const txValues = txSheet.getDataRange().getValues();
  const headers = (txValues[0] || []).map(String);

  const transactions = txValues.slice(1)
    .filter(row => row.some(value => value !== "" && value !== null))
    .map((row, index) => {
      const tx = rowToTransaction_(headers, row);
      tx.sortOrder = index + 2; // actual sheet row; breaks timestamp ties
      return tx;
    })
    .sort((a, b) => {
      const timeDiff = Number(b.createdAt || 0) - Number(a.createdAt || 0);
      return timeDiff || Number(b.sortOrder || 0) - Number(a.sortOrder || 0);
    });

  const listValues = listSheet.getDataRange().getDisplayValues();
  const listHeaders = listValues[0] || [];
  const lists = {};
  listHeaders.forEach((header, index) => {
    lists[header] = listValues.slice(1).map(row => row[index]).filter(Boolean);
  });

  return {
    spreadsheetName: ss.getName(),
    transactions: transactions,
    lists: lists,
    serverTime: new Date().toISOString()
  };
}

function saveTransaction(transaction) {
  if (!transaction || typeof transaction !== "object") {
    throw new Error("Invalid transaction payload.");
  }

  const type = String(transaction.type || "").trim();
  const date = String(transaction.date || "").trim();
  const amount = Number(transaction.amount);
  const paid = Number(transaction.paid || 0);
  const allowedTypes = ["Purchase", "Expense", "Supplier Payment", "Sales", "Cash Adjustment"];

  if (!allowedTypes.includes(type)) throw new Error("Invalid transaction type.");
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error("Enter a valid transaction date.");
  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || "Asia/Kolkata",
    "yyyy-MM-dd"
  );
  if (date > today) {
    throw new Error("Future-dated transactions are not allowed. Choose today or an earlier date.");
  }
  if (!Number.isFinite(amount) || amount <= 0) throw new Error("Amount must be greater than zero.");
  if (!Number.isFinite(paid) || paid < 0 || paid > amount) throw new Error("Amount paid is invalid.");

  const supplier = (type === "Purchase" || type === "Supplier Payment") ? getCanonicalSupplierName_(transaction.supplier) : "";
  if ((type === "Purchase" || type === "Supplier Payment") && !supplier) {
    throw new Error("Supplier is required for purchases and supplier payments.");
  }

  const method = String(transaction.method || "").trim();
  if (type === "Purchase" && !["Cash", "UPI", "Credit"].includes(method)) {
    throw new Error("Choose Cash, UPI or Credit for the purchase payment method.");
  }
  if (type === "Expense" && !["Cash", "UPI"].includes(method)) {
    throw new Error("Choose Cash or UPI for the expense payment method.");
  }
  if (type === "Supplier Payment" && !["Cash", "UPI"].includes(method)) {
    throw new Error("Choose Cash or UPI for the supplier payment.");
  }
  if (type === "Cash Adjustment" && !["Cash", "UPI"].includes(method)) {
    throw new Error("Choose Cash or UPI for the cash adjustment.");
  }

  const sheet = getSheet_(SHEET_NAMES.transactions);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const now = new Date();
  const id = Utilities.getUuid();

  let category = String(transaction.category || "").trim();
  if (type === "Supplier Payment") category = "Supplier Settlement";
  if (type === "Sales") category = "Sales";
  if (type === "Cash Adjustment") category = method === "UPI" ? "Cash Adjustment — UPI" : "Cash Adjustment";

  // The existing agreed 16-column sheet is retained.
  // For Cash Adjustment, Item / Description stores the direction; Notes stores the explanation.
  let item = String(transaction.item || "").trim();
  if (type === "Cash Adjustment") {
    item = transaction.cashDirection === "out" ? "Cash withdrawn" : "Cash added";
  }

  const rowObject = {
    "Transaction ID": id,
    "Date": date,
    "Transaction Type": type,
    "Category": category,
    "Item / Description": item,
    "Supplier": supplier,
    "Quantity": String(transaction.qty || "").trim(),
    "Unit": String(transaction.unit || "").trim(),
    "Total Amount": amount,
    "Amount Paid Now": type === "Sales" ? Math.max(0, amount - Number(transaction.deduction || 0)) : paid,
    "Payment Method": method,
    "Staff": String(transaction.employee || "").trim(),
    "Sales Channel": type === "Sales" ? String(transaction.channel || "").trim() : "",
    "Notes": String(transaction.notes || "").trim(),
    "Entered By": String(transaction.enteredBy || "").trim(),
    "Created At": now
  };

  const row = headers.map(header => Object.prototype.hasOwnProperty.call(rowObject, header) ? rowObject[header] : "");
  sheet.appendRow(row);
  return { success: true, id: id, message: "Transaction saved." };
}

function addSupplier(supplierName) {
  const name = formatSupplierName_(supplierName);
  if (!name) throw new Error("Enter a supplier name.");
  const sheet = getSheet_(SHEET_NAMES.lists), values = sheet.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  const col = headers.findIndex(h => String(h).trim().toLowerCase() === "suppliers");
  if (col < 0) throw new Error('The Lists sheet must contain a "Suppliers" column.');
  const existing = values.slice(1).map(row => String(row[col] || "").trim()).filter(Boolean);
  const match = existing.find(x => supplierKey_(x) === supplierKey_(name));
  if (match) return {success:true, suppliers:existing, canonicalName:formatSupplierName_(match)};
  sheet.getRange(Math.max(2, sheet.getLastRow() + 1), col + 1).setValue(name);
  return {success:true, suppliers:[...existing,name], canonicalName:name};
}

function deleteAllTransactions() {
  throw new Error(
    "Bulk deletion is disabled in the VYRA app for data safety. " +
    "If a reset is genuinely required, export a JSON backup first, " +
    "confirm with the other owner, and delete transaction rows directly " +
    "in the Google Sheet while preserving row 1 headers."
  );
}

function restoreTransactions(transactions, confirmText) {
  if (confirmText !== "RESTORE TRANSACTIONS") {
    throw new Error("Confirmation text did not match. No data was changed.");
  }
  if (!Array.isArray(transactions)) throw new Error("Backup does not contain a transaction list.");

  const sheet = getSheet_(SHEET_NAMES.transactions);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const rows = transactions.map(tx => {
    const type = String(tx.type || "");
    const amount = Number(tx.amount || 0);
    const paid = Number(tx.paid || 0);
    const category = type === "Supplier Payment" ? "Supplier Settlement" :
      type === "Sales" ? "Sales" :
      type === "Cash Adjustment" ? (String(tx.category || "") === "Cash Adjustment — UPI" ? "Cash Adjustment — UPI" : "Cash Adjustment") : String(tx.category || "");
    const item = type === "Cash Adjustment" ?
      (tx.cashDirection === "out" ? "Cash withdrawn" : "Cash added") : String(tx.item || "");
    const obj = {
      "Transaction ID": String(tx.id || Utilities.getUuid()),
      "Date": String(tx.date || ""),
      "Transaction Type": type,
      "Category": category,
      "Item / Description": item,
      "Supplier": tx.supplier ? formatSupplierName_(tx.supplier) : "",
      "Quantity": String(tx.qty || ""),
      "Unit": String(tx.unit || ""),
      "Total Amount": amount,
      "Amount Paid Now": type === "Sales" ? Math.max(0, amount - Number(tx.deduction || 0)) : paid,
      "Payment Method": String(tx.method || ""),
      "Staff": String(tx.employee || ""),
      "Sales Channel": String(tx.channel || ""),
      "Notes": String(tx.notes || ""),
      "Entered By": String(tx.enteredBy || ""),
      "Created At": new Date()
    };
    return headers.map(header => Object.prototype.hasOwnProperty.call(obj, header) ? obj[header] : "");
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  return { success: true, importedRows: rows.length };
}

function rowToTransaction_(headers, row) {
  const o = {};
  headers.forEach((header, index) => o[header] = row[index]);
  const type = String(o["Transaction Type"] || "");
  const amount = Number(o["Total Amount"] || 0);
  const paidNow = Number(o["Amount Paid Now"] || 0);
  const method = String(o["Payment Method"] || "");
  const item = String(o["Item / Description"] || "");
  const deduction = type === "Sales" ? Math.max(0, amount - paidNow) : 0;

  const createdRaw = o["Created At"];
  const createdTime = createdRaw instanceof Date
    ? createdRaw.getTime()
    : (createdRaw ? new Date(createdRaw).getTime() : 0);

  return {
    id: String(o["Transaction ID"] || ""),
    date: normalizeDate_(o["Date"]),
    type: type,
    category: String(o["Category"] || ""),
    item: type === "Cash Adjustment" ? "" : item,
    supplier: formatSupplierName_(o["Supplier"]),
    qty: String(o["Quantity"] || ""),
    unit: String(o["Unit"] || ""),
    amount: amount,
    paid: paidNow,
    status: type === "Purchase" ? (paidNow <= 0 ? "Credit" : paidNow < amount ? "Partial" : "Paid") : "",
    method: method,
    employee: String(o["Staff"] || ""),
    channel: String(o["Sales Channel"] || ""),
    deduction: deduction,
    cashDirection: type === "Cash Adjustment" && item.toLowerCase().includes("withdrawn") ? "out" : "in",
    notes: String(o["Notes"] || ""),
    createdAt: Number.isFinite(createdTime) ? createdTime : 0,
    enteredBy: String(o["Entered By"] || "")
  };
}

function normalizeDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || "Asia/Kolkata", "yyyy-MM-dd");
  }
  const s = String(value || "");
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, "Asia/Kolkata", "yyyy-MM-dd");
}

function testCloudRead() {
  const data = getInitialData();

  Logger.log("Spreadsheet: " + data.spreadsheetName);
  Logger.log("Transactions loaded: " + data.transactions.length);
  Logger.log("Lists loaded: " + Object.keys(data.lists).length);
  Logger.log("Cloud read successful.");
}
