/**
 * ============================================================
 * ORDER MANAGEMENT PRO - BACKEND
 * Refactored for Stability & Performance
 * ============================================================
 */

const CONFIG = {
  SHEET_ORDERS: "Orders",
  SHEET_PRODUCTS: "Products",
  SHEET_COURIERS: "Courier"
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Order Management Pro')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=0');
}

// --- Standard Response Helper ---
function createResponse(success, data = null, message = "") {
  return { success, data, message };
}

// --- API: Get Master Data (Combined for speed) ---
function getInitData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Helper to get simple list from a sheet
    const getList = (sheetName, col) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return [];
      const lastRow = sheet.getLastRow();
      if (lastRow < 2) return [];
      return sheet.getRange(`${col}2:${col}${lastRow}`).getValues()
        .flat()
        .filter(item => item && item.toString().trim() !== "");
    };

    return createResponse(true, {
      products: getList(CONFIG.SHEET_PRODUCTS, "B"),
      couriers: getList(CONFIG.SHEET_COURIERS, "A")
    });
  } catch (e) {
    return createResponse(false, null, e.message);
  }
}

// --- Utility: Format Date for Sheet ---
function formatDateThai(dateStr) {
  if (!dateStr) return "";
  const parts = dateStr.split('-'); // Expecting YYYY-MM-DD
  if (parts.length === 3) return `'${parts[2]}/${parts[1]}/${parts[0]}`; // 'DD/MM/YYYY force text
  return "'" + dateStr;
}

// --- Utility: Format Phone for Sheet ---
function formatPhoneForSheet(p) {
  let str = (p || "").toString().trim().replace(/-|\s/g, "");
  if (str === "") return "";
  if (str.startsWith("66")) str = "0" + str.substring(2);
  // Force text format to keep leading zero
  return "'" + str; 
}

// --- API: Save Data ---
function saveData(data) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createResponse(false, null, "ระบบกำลังยุ่ง กรุณาลองใหม่");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    
    const phone = formatPhoneForSheet(data.phone);
    const dateStr = formatDateThai(data.date);
    const timestamp = new Date(); // Log Create Time if needed

    // Batch operations are faster, but sticking to row-by-row for the items logic
    const rows = data.items.map(item => [
      dateStr, 
      data.order_no.toString(), 
      data.set_name, 
      data.page_no,
      data.recipient_name, 
      data.address, 
      phone,
      item.name, 
      item.qty, 
      data.courier
    ]);

    // Append all rows at once
    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    }

    return createResponse(true, null, "บันทึกข้อมูลสำเร็จ");
  } catch (e) {
    return createResponse(false, null, "Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// --- API: Update Order (Improved: Edit in place) ---
function updateOrder(oldNo, newData) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createResponse(false, null, "ระบบกำลังยุ่ง");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    
    // 1. Get all data first
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues(); // Keep original values (not display values) to match types
    
    // 2. Find row indices to DELETE first (Strategy: Delete old, Insert new is cleaner for multi-row items)
    // *NOTE: Although Update-in-place is better for single rows, 
    // for One-Order-Multi-Items structure, Delete-And-Insert is often safer to avoid gaps.*
    // *However, to be "Perfectionist" and preserve table integrity, we will use Delete & Insert logic carefully.*
    
    const rowsToDelete = [];
    for (let i = values.length - 1; i >= 1; i--) { // Reverse loop
      if (values[i][1].toString() === oldNo.toString()) {
        rowsToDelete.push(i + 1); // 1-based index
      }
    }

    if (rowsToDelete.length === 0) return createResponse(false, null, "ไม่พบข้อมูลเดิม");

    // Delete rows
    rowsToDelete.forEach(rowIndex => sheet.deleteRow(rowIndex));

    // Prepare new data
    const phone = formatPhoneForSheet(newData.phone);
    const dateStr = formatDateThai(newData.date);
    const newRows = newData.items.map(item => [
      dateStr, newData.order_no.toString(), newData.set_name, newData.page_no,
      newData.recipient_name, newData.address, phone,
      item.name, item.qty, newData.courier
    ]);

    // Insert at the position of the first deleted row (to keep chronological order roughly)
    const insertPos = rowsToDelete[rowsToDelete.length - 1]; // The smallest index
    
    // Insert blank rows then set values (Faster and safer than insertRow one by one)
    sheet.insertRowsBefore(insertPos, newRows.length);
    sheet.getRange(insertPos, 1, newRows.length, newRows[0].length).setValues(newRows);

    return createResponse(true, null, "แก้ไขข้อมูลเรียบร้อย");
  } catch (e) {
    return createResponse(false, null, "Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// --- API: Delete Order ---
function deleteOrder(no) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return createResponse(false, null, "ระบบยุ่ง");

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    const data = sheet.getDataRange().getValues();
    
    // Batch delete is complex in GAS because indexes shift. 
    // Reverse loop delete is the standard reliable way.
    let deletedCount = 0;
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][1].toString() === no.toString()) { 
        sheet.deleteRow(i + 1); 
        deletedCount++;
      }
    }
    
    if (deletedCount === 0) return createResponse(false, null, "ไม่พบข้อมูล");
    return createResponse(true, null, `ลบข้อมูลเรียบร้อย (${deletedCount} รายการ)`);
  } catch (e) {
    return createResponse(false, null, "Error: " + e.message);
  } finally {
    lock.releaseLock();
  }
}

// --- API: Get Orders Data ---
function getOrderData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_ORDERS);
    if (!sheet) return [];
    
    // Use getDisplayValues to ensure we get what looks like dates/text properly
    const rows = sheet.getDataRange().getDisplayValues(); 
    if (rows.length <= 1) return [];

    const grouped = {};
    // Skip header (row 0)
    for (let i = 1; i < rows.length; i++) {
      const r = rows[i];
      const id = r[1]; // Order No
      if (!id) continue;

      if (!grouped[id]) {
        grouped[id] = {
          date: r[0], 
          no: id, 
          set: r[2], 
          pg: r[3], 
          nm: r[4],
          ad: r[5], 
          tel: r[6], 
          items: [], 
          cur: r[9]
        };
      }
      grouped[id].items.push({ name: r[7], qty: r[8] });
    }

    // Convert Object back to Array for frontend
    // Return newest first (Reverse Object.values isn't guaranteed, better sort)
    // But assuming appendRow, last is newest.
    const result = Object.values(grouped).reverse();
    return createResponse(true, result);
    
  } catch (e) {
    return createResponse(false, [], e.message);
  }
}
