// ============================================================
// File: PurchaseOrders.gs — Purchase Order Module
// Version: 2.4 | Unique Check Fixed & Cleaned
// ============================================================

function createPurchaseOrder(data, caller) {
  // Validate required fields first (before touching the sheet)
  if (!data.vendorName)   return error('Vendor Name is required');
  if (!data.fabricName)   return error('Fabric Name is required');
  if (!data.color)        return error('Color is required');
  if (!data.orderedQty || isNaN(data.orderedQty) || parseFloat(data.orderedQty) <= 0)
    return error('Valid Ordered Quantity is required');
  if (!data.deliveryDate) return error('Delivery Date is required');

  const sheet    = getSheet(SHEETS.PURCHASE_ORDERS);
  const poNumber = generateSequentialId('PO', sheet);
  
  // --- FIXED UNIQUE PO CHECK FOR AUTO-GENERATION ---
  const sheetData = sheet.getDataRange().getValues();
  if (sheetData.length > 1) {
    const poNumberStr = poNumber.toString().trim().toLowerCase();
    // Check if the auto-generated PO ID already exists in Column A
    const isDuplicate = sheetData.slice(1).some(row => {
      return row[0] && row[0].toString().trim().toLowerCase() === poNumberStr;
    });
    
    if (isDuplicate) {
      return error(`System Error: The auto-generated PO Number (${poNumber}) already exists in database. Please check your Sequential ID generator configuration.`);
    }
  }
  // -------------------------------------------------

  const row = [
    poNumber,                                    // col 1  PO_Number
    data.vendorName,                             // col 2  Vendor_Name
    data.articleNumber || '',                    // col 3  Article_Number
    data.fabricName,                             // col 4  Fabric_Name
    data.color,                                  // col 5  Color
    data.width        || '',                     // col 6  Width
    data.widthUnit    || 'Inch',                 // col 7  Width_Unit
    parseFloat(data.orderedQty),                 // col 8  Ordered_Qty
    data.unit         || 'Meter',                // col 9  Unit
    data.deliveryDate,                           // col 10 Delivery_Date
    'Pending',                                   // col 11 Status
    (caller && caller.username) || Session.getActiveUser().getEmail() || 'User', // col 12 Created_By
    new Date(),                                  // col 13 Created_At
    data.notes || ''                             // col 14 Notes
  ];

  sheet.appendRow(row);

  try {
    logActivity('CREATE', 'PurchaseOrder', poNumber, {
      vendor: data.vendorName, fabric: data.fabricName, qty: data.orderedQty
    }, caller && caller.username, caller && caller.role);
  } catch(e) {}

  // Auto-format Status cell
  const lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 11).setBackground('#FFF3CD');

  return success({ poNumber }, `Purchase Order ${poNumber} created successfully`);
}

function getPurchaseOrders(filters) {
  const sheet   = getSheet(SHEETS.PURCHASE_ORDERS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return success([]);

  const headers = data[0];
  let rows = data.slice(1).map(row => rowToObj(headers, row));

  // Apply filters — null-safe string ops
  if (filters) {
    if (filters.status && filters.status !== 'All')
      rows = rows.filter(r => r.Status === filters.status);
    if (filters.vendor)
      rows = rows.filter(r => String(r.Vendor_Name || '').toLowerCase().includes(filters.vendor.toLowerCase()));
    if (filters.articleNumber)
      rows = rows.filter(r => String(r.Article_Number || '').toLowerCase().includes(filters.articleNumber.toLowerCase()));
    if (filters.dateFrom)
      rows = rows.filter(r => r.Created_At && new Date(r.Created_At) >= new Date(filters.dateFrom));
    if (filters.dateTo)
      rows = rows.filter(r => r.Created_At && new Date(r.Created_At) <= new Date(filters.dateTo));
  }

  // Sort newest first — guard against invalid dates
  rows.sort((a, b) => {
    const da = a.Created_At ? new Date(a.Created_At) : 0;
    const db = b.Created_At ? new Date(b.Created_At) : 0;
    return db - da;
  });

  // Attach receipt summary
  rows = rows.map(po => {
    const receipts      = getReceiptsForPO(po.PO_Number);
    const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);
    return {
      ...po,
      Total_Received: totalReceived,
      Balance_Qty   : (parseFloat(po.Ordered_Qty) || 0) - totalReceived,
      Receipt_Count : receipts.length
    };
  });

  return success(rows);
}

function getPOById(poNumber) {
  if (!poNumber) return error('PO Number is required');

  const sheet   = getSheet(SHEETS.PURCHASE_ORDERS);
  const data    = sheet.getDataRange().getValues();
  if (data.length <= 1) return error('Purchase Order not found: ' + poNumber);

  const headers = data[0];
  const row     = data.slice(1).find(r => r[0] === poNumber);
  if (!row) return error('Purchase Order not found: ' + poNumber);

  const po            = rowToObj(headers, row);
  const receipts      = getReceiptsForPO(poNumber);
  const totalReceived = receipts.reduce((s, r) => s + (parseFloat(r.Received_Qty) || 0), 0);

  return success({
    ...po,
    receipts,
    Total_Received: totalReceived,
    Balance_Qty   : (parseFloat(po.Ordered_Qty) || 0) - totalReceived
  });
}

function updatePOStatus(data, caller) {
  if (!data || !data.poNumber || !data.status)
    return error('PO Number and Status are required');

  const sheet  = getSheet(SHEETS.PURCHASE_ORDERS);
  const values = sheet.getDataRange().getValues();
  const rowIdx = values.findIndex(r => r[0] === data.poNumber);
  if (rowIdx === -1) return error('Purchase Order not found: ' + data.poNumber);

  sheet.getRange(rowIdx + 1, 11).setValue(data.status);

  // FIX: Partial was #676E00 (near-black) — unreadable. Use legible colours.
  const colors = {
    Pending  : '#FFF3CD',
    Partial  : '#CCE5FF',  // light blue — readable
    Completed: '#D4EDDA',  // light green
    Cancelled: '#F8D7DA'   // light red
  };
  sheet.getRange(rowIdx + 1, 11).setBackground(colors[data.status] || '#FFFFFF');

  try {
    logActivity('UPDATE_STATUS', 'PurchaseOrder', data.poNumber,
      { status: data.status }, caller && caller.username, caller && caller.role);
  } catch(e) {}

  return success({ poNumber: data.poNumber, status: data.status });
}

// ── INTERNAL HELPER: receipts for a given PO ────────────────
function getReceiptsForPO(poNumber) {
  const sheet  = getSheet(SHEETS.FABRIC_RECEIPTS);
  const data   = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1)
    .filter(r => r[1] === poNumber)
    .map(r => rowToObj(headers, r));
}

// ── CANONICAL rowToObj — convert row array → keyed object ───
function rowToObj(headers, row) {
  const obj = {};
  headers.forEach((h, i) => { obj[h] = row[i] !== undefined ? row[i] : ''; });
  return obj;
}