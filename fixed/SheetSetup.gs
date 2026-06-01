// ============================================================
// SheetSetup.gs — Database Initialization
// Garment Manufacturing ERP
// Run once to create all sheets with correct headers
// ============================================================

function initializeDatabase() {
  const ss = SpreadsheetApp.openById('1gk88o8KDAIFPzestuZlJs9SQuuaWxlEIGQwpsLpjIYQ');

  const SHEETS = [
    {
      name: '_Users',
      headers: ['UserID','FullName','Email','PasswordHash','Role','Status',
                'LastLogin','FailedAttempts','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: true
    },
    {
      name: '_Sessions',
      headers: ['SessionID','UserID','Token','CreatedAt','LastActivity','IsActive','IPAddress'],
      hidden: true
    },
    {
      name: '_AuditLogs',
      headers: ['LogID','Timestamp','UserID','UserName','Action','Module',
                'RecordID','OldValues','NewValues','IPAddress'],
      hidden: true
    },
    {
      name: '_Settings',
      headers: ['SettingKey','SettingValue','Description','UpdatedBy','UpdatedAt'],
      hidden: true
    },
    {
      name: 'BuyerMaster',
      headers: ['BuyerID','BuyerName','BuyerCode','ContactPerson','Mobile',
                'Email','Address','Country','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'ArticleMaster',
      headers: ['ArticleID','ArticleCode','ArticleName','BuyerID','Category',
                'Description','Season','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'BOM',
      headers: ['BOMID','ArticleID','Size','Color','FabricID','FabricColorID',
                'ConsumptionPerPc','UOM','WastagePercent','EffectiveConsumption',
                'Remarks','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'FabricMaster',
      headers: ['FabricID','DesignName','DesignCode','Category','Width',
                'UOM','Remarks','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'FabricColors',
      headers: ['FabricColorID','FabricID','ColorName','ColorCode','Status'],
      hidden: false
    },
    {
      name: 'VendorMaster',
      headers: ['VendorID','VendorName','VendorCode','Address','GSTNumber',
                'ContactPerson','Mobile','Email','VendorType','Status',
                'CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'SalesOrders',
      headers: ['SOID','SONumber','BuyerID','SODate','DeliveryDate','Season','Remarks',
                'Status','TotalPieces','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'SalesOrderItems',
      headers: ['SOItemID','SOID','ArticleID','Size','Color','OrderedQty',
                'Rate','Amount'],
      hidden: false
    },
    {
      name: 'PurchaseOrders',
      headers: ['POID','PONumber','PODate','VendorID','Remarks','Status','TotalAmount',
                'ApprovedBy','ApprovedAt','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'PurchaseOrderItems',
      headers: ['POItemID','POID','FabricID','FabricColorID','OrderedQty',
                'ReceivedQty','PendingQty','UOM','Width','Rate','Amount','ExcessQty','ItemStatus'],
      hidden: false
    },
    {
      name: 'GRN',
      headers: ['GRNID','GRNNumber','GRNDate','VendorID','POID','ChallanNumber',
                'VehicleNumber','Remarks','CreatedBy','CreatedAt'],
      hidden: false
    },
    {
      name: 'GRNItems',
      headers: ['GRNItemID','GRNID','POItemID','FabricID','FabricColorID',
                'ReceivedQty','UOM','Remarks'],
      hidden: false
    },
    {
      name: 'Inventory',
      headers: ['InventoryID','FabricID','FabricColorID','PurchasedQty',
                'ReceivedQty','IssuedQty','ReturnedQty','AvailableQty','LastUpdated'],
      hidden: false
    },
    {
      name: 'InventoryLedger',
      headers: ['LedgerID','TransactionDate','TransactionType','ReferenceNumber',
                'FabricID','FabricColorID','InQty','OutQty','BalanceQty',
                'Remarks','CreatedBy'],
      hidden: false
    },
    {
      name: 'FabricIssue',
      headers: ['IssueID','IssueNumber','IssueDate','ArticleID','SOID','IssueFor',
                'FabricatorID','Remarks','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'FabricIssueItems',
      headers: ['IssueItemID','IssueID','FabricID','FabricColorID',
                'AvailableQty','IssueQty','UOM','ReturnedQty'],
      hidden: false
    },
    {
      name: 'ProductionPlan',
      headers: ['PlanID','SOID','ArticleID','PlannedStartDate','PlannedEndDate',
                'BundleSize','TotalPieces','Status','Notes',
                'CreatedBy','CreatedAt','UpdatedAt'],
      hidden: false
    },
    {
      name: 'Bundles',
      headers: ['BundleID','PlanID','SOID','ArticleID','Size','Color',
                'BundleNumber','Pieces','BarcodeValue','QRData',
                'CuttingStatus','CuttingOperator','CuttingDate',
                'StitchingStatus','StitchingOperator','StitchingDate',
                'FinishingStatus','FinishingOperator','FinishingDate',
                'PackingStatus','PackingOperator','PackingDate',
                'CurrentProcess','FabricatorOut','CreatedBy','CreatedAt','UpdatedAt'],
      hidden: false
    },
    {
      name: 'BundleTracking',
      headers: ['TrackID','BundleID','PlanID','Process','Status',
                'Operator','Notes','TrackedBy','TrackedAt'],
      hidden: false
    },
    {
      name: 'FabricatorChallan',
      headers: ['ChallanID','ChallanDate','ChallanType','FabricatorID','IssueID',
                'Remarks','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'FabricatorChallanItems',
      headers: ['ChallanItemID','ChallanID','FabricID','FabricColorID','Qty','UOM','Remarks'],
      hidden: false
    },
    {
      name: 'Dispatch',
      headers: ['DispatchID','DispatchNumber','DispatchDate','SOID','BuyerID','VehicleNumber',
                'DriverName','DriverMobile','TotalPieces','Remarks','Status',
                'CreatedBy','CreatedAt'],
      hidden: false
    },
    {
      name: 'DispatchItems',
      headers: ['DispatchItemID','DispatchID','BundleID','ArticleID','Size','Color','Qty'],
      hidden: false
    },
    {
      name: 'Returns',
      headers: ['ReturnID','ReturnDate','ReturnType','ReferenceID','BuyerID',
                'Reason','Status','CreatedBy','CreatedAt','UpdatedBy','UpdatedAt'],
      hidden: false
    },
    {
      name: 'ReturnItems',
      headers: ['ReturnItemID','ReturnID','FabricID','FabricColorID','BundleID',
                'Qty','UOM','Remarks'],
      hidden: false
    }
  ];

  let created = 0;
  let skipped = 0;

  SHEETS.forEach(function(sheetDef) {
    let sheet = ss.getSheetByName(sheetDef.name);
    if (!sheet) {
      sheet = ss.insertSheet(sheetDef.name);
      // Write headers
      sheet.getRange(1, 1, 1, sheetDef.headers.length).setValues([sheetDef.headers]);
      // Format header row
      const headerRange = sheet.getRange(1, 1, 1, sheetDef.headers.length);
      headerRange.setBackground('#1a1a2e');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      headerRange.setFontSize(11);
      sheet.setFrozenRows(1);
      // Auto-resize columns
      sheet.autoResizeColumns(1, sheetDef.headers.length);
      // Hide system sheets
      if (sheetDef.hidden) {
        sheet.hideSheet();
      }
      created++;
    } else {
      skipped++;
    }
  });

  // NOTE: NO sheet protection — protection blocks script writes from web app
  // The ERP enforces access control via session tokens and role checks in code

  // Seed default settings
  seedSettings(ss);

  // Create default Super Admin
  createDefaultSuperAdmin(ss);

  Logger.log('✅ Database Initialized — Created: ' + created + ' sheets, Skipped: ' + skipped);
  Logger.log('Default login → Email: admin@erp.com | Password: Admin@1234');
}

function seedSettings(ss) {
  const sheet = ss.getSheetByName('_Settings');
  if (!sheet) return;
  const existingData = sheet.getDataRange().getValues();
  if (existingData.length > 1) return; // Already seeded

  const now = new Date().toISOString();
  const settings = [
    ['COMPANY_NAME',       'My Company ERP', 'Company display name',           'SYSTEM', now],
    ['COMPANY_ADDRESS',    'New Delhi',                           'Company address',                 'SYSTEM', now],
    ['FINANCIAL_YEAR',     '2026-27',                    'Current financial year',          'SYSTEM', now],
    ['PO_PREFIX',          'PO',                         'Purchase Order number prefix',    'SYSTEM', now],
    ['GRN_PREFIX',         'GRN',                        'GRN number prefix',               'SYSTEM', now],
    ['SO_PREFIX',          'SO',                         'Sales Order number prefix',       'SYSTEM', now],
    ['ISS_PREFIX',         'ISS',                        'Issue number prefix',             'SYSTEM', now],
    ['BND_PREFIX',         'BND',                        'Bundle number prefix',            'SYSTEM', now],
    ['DSP_PREFIX',         'DSP',                        'Dispatch number prefix',          'SYSTEM', now],
    ['RTN_PREFIX',         'RTN',                        'Return number prefix',            'SYSTEM', now],
    ['PP_PREFIX',          'PP',                         'Production Plan prefix',          'SYSTEM', now],
    ['FJC_PREFIX',         'FJC',                        'Fabricator Challan prefix',       'SYSTEM', now],
    ['LOW_STOCK_THRESHOLD','50',                          'Alert when stock below this qty', 'SYSTEM', now],
    ['BUNDLE_SIZE_DEFAULT','12',                          'Default pieces per bundle',       'SYSTEM', now],
    ['SESSION_ACTIVE',     'true',                       'Session persists until logout',   'SYSTEM', now]
  ];

  sheet.getRange(2, 1, settings.length, 5).setValues(settings);
}

function createDefaultSuperAdmin(ss) {
  const sheet = ss.getSheetByName('_Users');
  if (!sheet) return;
  const existingData = sheet.getDataRange().getValues();
  if (existingData.length > 1) return; // User already exists

  const now = new Date().toISOString();
  const passwordHash = hashPassword('Admin@123');
  const userRow = [
    'USR-00001',
    'Super Administrator',
    'admin@erp.com',
    passwordHash,
    'Super Admin',
    'Active',
    '',        // LastLogin
    0,         // FailedAttempts
    'SYSTEM',  // CreatedBy
    now,       // CreatedAt
    '',        // UpdatedBy
    ''         // UpdatedAt
  ];
  sheet.getRange(2, 1, 1, userRow.length).setValues([userRow]);
}

function hashPassword(password) {
  const rawBytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    password,
    Utilities.Charset.UTF_8
  );
  return rawBytes.map(function(b) {
    return ('0' + (b & 0xFF).toString(16)).slice(-2);
  }).join('');
}

// ── Remove all sheet protections (run if sheets got locked accidentally) ──
function removeAllProtections() {
  const ss = SpreadsheetApp.openById('1gk88o8KDAIFPzestuZlJs9SQuuaWxlEIGQwpsLpjIYQ');
  const sheets = ss.getSheets();
  let removed = 0;
  sheets.forEach(function(sheet) {
    const protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
    protections.forEach(function(p) {
      p.remove();
      removed++;
    });
  });
  Logger.log('✅ Removed ' + removed + ' sheet protections');
}

// ── Clean expired sessions (run periodically for maintenance) ──
function cleanExpiredSessions() {
  const ss = SpreadsheetApp.openById('1gk88o8KDAIFPzestuZlJs9SQuuaWxlEIGQwpsLpjIYQ');
  const sheet = ss.getSheetByName('_Sessions');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return;

  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const isActive = data[i][5]; // IsActive column
    if (isActive === true || isActive === 'true') {
      rows.push(data[i]);
    }
  }
  sheet.clearContents();
  sheet.getRange(1, 1, 1, data[0].length).setValues([data[0]]);
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, data[0].length).setValues(rows);
  }
  Logger.log('✅ Sessions cleaned. Active sessions kept: ' + rows.length);
}