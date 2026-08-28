/**
 * Siriwong Sales, Designer, Foreman & Manager Workflow Backend
 * Google Apps Script Server Backend & Google Sheets Integration
 */

function doGet(e) {
  var htmlOutput = null;

  // 1. Try Template From File 'Index'
  try {
    htmlOutput = HtmlService.createTemplateFromFile('Index').evaluate();
  } catch (e1) {}

  // 2. Try Template From File 'index' (lowercase)
  if (!htmlOutput) {
    try {
      htmlOutput = HtmlService.createTemplateFromFile('index').evaluate();
    } catch (e2) {}
  }

  // 3. Try HtmlOutput From File 'Index'
  if (!htmlOutput) {
    try {
      htmlOutput = HtmlService.createHtmlOutputFromFile('Index');
    } catch (e3) {}
  }

  // 4. Try HtmlOutput From File 'index' (lowercase)
  if (!htmlOutput) {
    try {
      htmlOutput = HtmlService.createHtmlOutputFromFile('index');
    } catch (e4) {}
  }

  if (htmlOutput) {
    return htmlOutput
      .setTitle('ระบบจัดการงาน - Siriwong Manager')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
  }

  return HtmlService.createHtmlOutput('<div style="font-family:sans-serif;padding:20px;text-align:center;"><h2>⚠️ ไม่พบไฟล์ HTML ในระบบ Apps Script</h2><p>กรุณาตรวจสอบว่ามีไฟล์ชื่อ <b>Index.html</b> หรือ <b>index.html</b> ในแถบซ้ายมือหรือไม่</p></div>');
}

/**
 * Helper JSON Safe Parser
 */
function safeJsonParse(str, fallback) {
  if (!str) return fallback;
  try {
    if (typeof str === 'object') return str;
    return JSON.parse(str);
  } catch (e) {
    return fallback;
  }
}

/**
 * Get or Connect Google Spreadsheet Database
 */
function getDatabaseSpreadsheet() {
  var ss = null;
  var targetSheetId = '1gw81aJmpCRyOFDwv61bOK5NR32SRZHm8GxKaXiDvcq4';
  var userProperties = PropertiesService.getScriptProperties();

  // 1. Primary Target Spreadsheet (Provided by User)
  try {
    ss = SpreadsheetApp.openById(targetSheetId);
    if (ss) userProperties.setProperty('SPREADSHEET_ID', targetSheetId);
  } catch (e) {}

  // 2. Container bound check
  if (!ss) {
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      if (ss) userProperties.setProperty('SPREADSHEET_ID', ss.getId());
    } catch (e) {}
  }

  // 3. Script Property saved ID
  if (!ss) {
    var savedId = userProperties.getProperty('SPREADSHEET_ID');
    if (savedId) {
      try { ss = SpreadsheetApp.openById(savedId); } catch (e) {}
    }
  }

  // 4. Fallback create new spreadsheet
  if (!ss) {
    try {
      ss = SpreadsheetApp.create('Siriwong manager');
      userProperties.setProperty('SPREADSHEET_ID', ss.getId());
    } catch (e) {}
  }

  if (ss) setupSheetsIfMissing(ss);
  return ss;
}

/**
 * Helper: Create a Brand New Clean Google Sheet Database Automatically
 */
function createNewDatabaseSheet() {
  var ss = SpreadsheetApp.create('Siriwong_Workflow_Database_Clean');
  var userProperties = PropertiesService.getScriptProperties();
  userProperties.setProperty('SPREADSHEET_ID', ss.getId());
  setupSheetsIfMissing(ss);
  Logger.log('✅ สบสำเร็จ! สร้าง Google Sheet ใบใหม่เรียบร้อยแล้ว: ' + ss.getUrl());
  return ss.getUrl();
}

/**
 * Helper: Link an Existing Google Sheet by URL or ID
 */
function linkCustomSpreadsheet(urlOrId) {
  var id = urlOrId;
  if (urlOrId.indexOf('/d/') !== -1) {
    id = urlOrId.split('/d/')[1].split('/')[0];
  }
  var ss = SpreadsheetApp.openById(id);
  if (ss) {
    PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
    setupSheetsIfMissing(ss);
    Logger.log('✅ เชื่อมต่อ Google Sheet ใหม่สำเร็จ: ' + ss.getUrl());
  }
  return ss ? ss.getUrl() : null;
}

/**
 * Setup Headers and Tables
 */
function setupSheetsIfMissing(ss) {
  if (!ss) return;

  var projectsSheet = ss.getSheetByName('Projects');
  if (!projectsSheet) {
    try { projectsSheet = ss.insertSheet('Projects'); } catch(e) {}
  }

  if (projectsSheet && projectsSheet.getLastRow() === 0) {
    projectsSheet.appendRow([
      'ID', 'Code', 'CustomerName', 'CustomerPhone', 'Address', 'ProjectType',
      'Budget', 'StartDate', 'DueDate', 'Status', 'SalesName', 'DesignerName',
      'ForemanName', 'Notes', 'RevisionsJSON', 'FilesJSON', 'CommentsJSON',
      'CreatedAt', 'UpdatedAt'
    ]);
    try { projectsSheet.getRange(1, 1, 1, 19).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff'); } catch(e) {}

    var sampleProjects = getSampleProjects();
    sampleProjects.forEach(function(p) {
      projectsSheet.appendRow([
        p.id, p.code, p.customerName, p.customerPhone, p.address, p.projectType,
        p.budget, p.startDate, p.dueDate, p.status, p.salesName, p.designerName,
        p.foremanName, p.notes, JSON.stringify(p.revisions || []), JSON.stringify(p.files || []),
        JSON.stringify(p.comments || []), p.createdAt, p.updatedAt
      ]);
    });
  }

  var usersSheet = ss.getSheetByName('Users');
  if (!usersSheet) {
    try { usersSheet = ss.insertSheet('Users'); } catch(e) {}
  }

  if (usersSheet && usersSheet.getLastRow() === 0) {
    usersSheet.appendRow(['ID', 'Username', 'Password', 'Name', 'Role', 'RoleTitle', 'Avatar', 'PermissionsJSON']);
    try { usersSheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff'); } catch(e) {}

    var sampleUsers = getSampleUsers();
    sampleUsers.forEach(function(u) {
      usersSheet.appendRow([
        u.id, u.username, String(u.password), u.name, u.role, u.roleTitle, u.avatar, JSON.stringify(u.permissions || {})
      ]);
    });
  }

  var dailyTasksSheet = ss.getSheetByName('DailyTasks');
  if (!dailyTasksSheet) {
    try { dailyTasksSheet = ss.insertSheet('DailyTasks'); } catch(e) {}
  }

  var dailyTasksSheet = ss.getSheetByName('DailyTasks');
  if (!dailyTasksSheet) {
    try { dailyTasksSheet = ss.insertSheet('DailyTasks'); } catch(e) {}
  }

  if (dailyTasksSheet && dailyTasksSheet.getLastRow() === 0) {
    dailyTasksSheet.appendRow([
      'ID', 'Title', 'ProjectId', 'ProjectCode', 'ProjectName', 'AssignedBy',
      'AssigneeId', 'AssigneeName', 'AssigneeRole', 'AssigneeRoleTitle',
      'DueDate', 'Priority', 'Status', 'Notes', 'CreatedAt'
    ]);
    try { dailyTasksSheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff'); } catch(e) {}
  }

  var invoicesSheet = ss.getSheetByName('Invoices');
  if (!invoicesSheet) {
    try { invoicesSheet = ss.insertSheet('Invoices'); } catch(e) {}
  }

  if (invoicesSheet && invoicesSheet.getLastRow() === 0) {
    invoicesSheet.appendRow([
      'ID', 'DocType', 'DocNo', 'VendorName', 'TaxId', 'ProjectId',
      'ProjectCode', 'ProjectName', 'Amount', 'HasVat', 'VatAmount',
      'TotalAmount', 'DocDate', 'RecordedBy', 'FileUrl', 'Status',
      'Notes', 'CreatedAt'
    ]);
    try { invoicesSheet.getRange(1, 1, 1, 18).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff'); } catch(e) {}
  }

  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try { ss.deleteSheet(defaultSheet); } catch(e) {}
  }
}

/**
 * Client Initial Data API
 */
function getInitialData() {
  try {
    cleanDatabaseSheetDuplicates();

    var ss = getDatabaseSpreadsheet();
    if (!ss) {
      return { projects: deduplicateList(getSampleProjects(), 'id'), users: deduplicateList(getSampleUsers(), 'id'), dailyTasks: [], invoices: [], spreadsheetUrl: '#' };
    }

    var projectsSheet = ss.getSheetByName('Projects');
    var usersSheet = ss.getSheetByName('Users');
    var dailyTasksSheet = ss.getSheetByName('DailyTasks');
    var invoicesSheet = ss.getSheetByName('Invoices');

    var projects = [];
    if (projectsSheet && projectsSheet.getLastRow() > 1) {
      var pValues = projectsSheet.getDataRange().getValues();
      for (var i = 1; i < pValues.length; i++) {
        var row = pValues[i];
        if (!row[0]) continue;
        projects.push({
          id: String(row[0]),
          code: String(row[1] || ''),
          customerName: String(row[2] || ''),
          customerPhone: String(row[3] || ''),
          address: String(row[4] || ''),
          projectType: String(row[5] || ''),
          budget: Number(row[6]) || 0,
          startDate: String(row[7] || ''),
          dueDate: String(row[8] || ''),
          status: String(row[9] || 'lead'),
          salesName: String(row[10] || ''),
          designerName: String(row[11] || ''),
          foremanName: String(row[12] || ''),
          notes: String(row[13] || ''),
          revisions: safeJsonParse(row[14], []),
          files: safeJsonParse(row[15], []),
          comments: safeJsonParse(row[16], []),
          createdAt: String(row[17] || new Date().toISOString()),
          updatedAt: String(row[18] || new Date().toISOString()),
          progress: Number(row[19]) || 0,
          tasks: safeJsonParse(row[20], [])
        });
      }
    }

    var users = [];
    if (usersSheet && usersSheet.getLastRow() > 1) {
      var uValues = usersSheet.getDataRange().getValues();
      for (var j = 1; j < uValues.length; j++) {
        var uRow = uValues[j];
        if (!uRow[0]) continue;
        users.push({
          id: String(uRow[0]),
          username: String(uRow[1]),
          password: String(uRow[2]),
          name: String(uRow[3]),
          role: String(uRow[4]),
          roleTitle: String(uRow[5]),
          avatar: String(uRow[6]),
          permissions: safeJsonParse(uRow[7], {})
        });
      }
    }

    var dailyTasks = [];
    if (dailyTasksSheet && dailyTasksSheet.getLastRow() > 1) {
      var dtValues = dailyTasksSheet.getDataRange().getValues();
      for (var k = 1; k < dtValues.length; k++) {
        var dtRow = dtValues[k];
        if (!dtRow[0]) continue;
        dailyTasks.push({
          id: String(dtRow[0]),
          title: String(dtRow[1] || ''),
          projectId: String(dtRow[2] || ''),
          projectCode: String(dtRow[3] || ''),
          projectName: String(dtRow[4] || ''),
          assignedBy: String(dtRow[5] || ''),
          assigneeId: String(dtRow[6] || ''),
          assigneeName: String(dtRow[7] || ''),
          assigneeRole: String(dtRow[8] || ''),
          assigneeRoleTitle: String(dtRow[9] || ''),
          dueDate: String(dtRow[10] || ''),
          priority: String(dtRow[11] || 'normal'),
          status: String(dtRow[12] || 'pending'),
          notes: String(dtRow[13] || ''),
          createdAt: String(dtRow[14] || new Date().toISOString())
        });
      }
    }

    var invoices = [];
    if (invoicesSheet && invoicesSheet.getLastRow() > 1) {
      var invValues = invoicesSheet.getDataRange().getValues();
      for (var m = 1; m < invValues.length; m++) {
        var invRow = invValues[m];
        if (!invRow[0]) continue;
        invoices.push({
          id: String(invRow[0]),
          docType: String(invRow[1] || 'receipt'),
          docNo: String(invRow[2] || ''),
          vendorName: String(invRow[3] || ''),
          taxId: String(invRow[4] || ''),
          projectId: String(invRow[5] || ''),
          projectCode: String(invRow[6] || ''),
          projectName: String(invRow[7] || ''),
          amount: Number(invRow[8]) || 0,
          hasVat: String(invRow[9] || 'vat7'),
          vatAmount: Number(invRow[10]) || 0,
          totalAmount: Number(invRow[11]) || 0,
          docDate: String(invRow[12] || ''),
          recordedBy: String(invRow[13] || ''),
          fileUrl: String(invRow[14] || ''),
          status: String(invRow[15] || 'pending'),
          notes: String(invRow[16] || ''),
          createdAt: String(invRow[17] || new Date().toISOString())
        });
      }
    }

    if (users.length === 0) users = getSampleUsers();
    if (projects.length === 0) projects = getSampleProjects();

    return {
      projects: deduplicateList(projects, 'id'),
      users: deduplicateList(users, 'id'),
      dailyTasks: deduplicateList(dailyTasks, 'id'),
      invoices: deduplicateList(invoices, 'id'),
      spreadsheetUrl: ss.getUrl()
    };
  } catch (e) {
    return {
      projects: deduplicateList(getSampleProjects(), 'id'),
      users: deduplicateList(getSampleUsers(), 'id'),
      dailyTasks: [],
      invoices: [],
      spreadsheetUrl: '#'
    };
  }
}

/**
 * Auto-Clean All Duplicate Rows in Google Sheets (Users & Projects)
 */
function cleanDatabaseSheetDuplicates() {
  try {
    var ss = getDatabaseSpreadsheet();
    if (!ss) return;

    var usersSheet = ss.getSheetByName('Users');
    if (usersSheet && usersSheet.getLastRow() > 1) {
      var uValues = usersSheet.getDataRange().getValues();
      if (uValues.length > 1) {
        var header = uValues[0];
        var seen = {};
        var cleanRows = [header];

        for (var i = 1; i < uValues.length; i++) {
          var row = uValues[i];
          var key = String(row[0] || row[1]); // ID or Username
          if (key && !seen[key]) {
            seen[key] = true;
            cleanRows.push(row);
          }
        }

        if (cleanRows.length < uValues.length) {
          usersSheet.clearContents();
          usersSheet.getRange(1, 1, cleanRows.length, cleanRows[0].length).setValues(cleanRows);
          usersSheet.getRange(1, 1, 1, cleanRows[0].length).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
        }
      }
    }

    var projectsSheet = ss.getSheetByName('Projects');
    if (projectsSheet && projectsSheet.getLastRow() > 1) {
      var pValues = projectsSheet.getDataRange().getValues();
      if (pValues.length > 1) {
        var pHeader = pValues[0];
        var pSeen = {};
        var pCleanRows = [pHeader];

        for (var j = 1; j < pValues.length; j++) {
          var pRow = pValues[j];
          var pKey = String(pRow[0] || pRow[1]); // ID or Code
          if (pKey && !pSeen[pKey]) {
            pSeen[pKey] = true;
            pCleanRows.push(pRow);
          }
        }

        if (pCleanRows.length < pValues.length) {
          projectsSheet.clearContents();
          projectsSheet.getRange(1, 1, pCleanRows.length, pCleanRows[0].length).setValues(pCleanRows);
          projectsSheet.getRange(1, 1, 1, pCleanRows[0].length).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
        }
      }
    }
  } catch(err) {
    Logger.log('cleanDatabaseSheetDuplicates error: ' + err);
  }
}

/**
 * Save Projects API (Includes Progress % and Task Schedule)
 */
function apiSaveProjects(projectsData) {
  try {
    if (!projectsData || !Array.isArray(projectsData)) return { success: false };
    var cleanProjects = deduplicateList(projectsData, 'id');

    var ss = getDatabaseSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName('Projects') || ss.insertSheet('Projects');

    sheet.clearContents();
    sheet.appendRow([
      'ID', 'Code', 'CustomerName', 'CustomerPhone', 'Address', 'ProjectType',
      'Budget', 'StartDate', 'DueDate', 'Status', 'SalesName', 'DesignerName',
      'ForemanName', 'Notes', 'RevisionsJSON', 'FilesJSON', 'CommentsJSON',
      'CreatedAt', 'UpdatedAt', 'Progress', 'TasksJSON'
    ]);
    sheet.getRange(1, 1, 1, 21).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');

    cleanProjects.forEach(function(p) {
      sheet.appendRow([
        p.id, p.code, p.customerName, p.customerPhone, p.address, p.projectType,
        p.budget, p.startDate, p.dueDate, p.status, p.salesName, p.designerName,
        p.foremanName, p.notes, JSON.stringify(p.revisions || []), JSON.stringify(p.files || []),
        JSON.stringify(p.comments || []), p.createdAt, p.updatedAt,
        Number(p.progress) || 0, JSON.stringify(p.tasks || [])
      ]);
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Save Users API (Deduplicated)
 */
function apiSaveUsers(usersData) {
  try {
    if (!usersData || !Array.isArray(usersData)) return { success: false };
    var cleanUsers = deduplicateList(usersData, 'id');

    var ss = getDatabaseSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName('Users') || ss.insertSheet('Users');

    sheet.clearContents();
    sheet.appendRow(['ID', 'Username', 'Password', 'Name', 'Role', 'RoleTitle', 'Avatar', 'PermissionsJSON']);
    sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');

    cleanUsers.forEach(function(u) {
      sheet.appendRow([
        u.id, u.username, String(u.password), u.name, u.role, u.roleTitle, u.avatar, JSON.stringify(u.permissions || {})
      ]);
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * Save Daily Tasks API (Deduplicated)
 */
function apiSaveDailyTasks(dailyTasksData) {
  try {
    if (!dailyTasksData || !Array.isArray(dailyTasksData)) return { success: false };
    var cleanTasks = deduplicateList(dailyTasksData, 'id');

    var ss = getDatabaseSpreadsheet();
    if (!ss) return { success: false };
    var sheet = ss.getSheetByName('DailyTasks') || ss.insertSheet('DailyTasks');

    sheet.clearContents();
    sheet.appendRow([
      'ID', 'Title', 'ProjectId', 'ProjectCode', 'ProjectName', 'AssignedBy',
      'AssigneeId', 'AssigneeName', 'AssigneeRole', 'AssigneeRoleTitle',
      'DueDate', 'Priority', 'Status', 'Notes', 'CreatedAt'
    ]);
    sheet.getRange(1, 1, 1, 15).setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');

    cleanTasks.forEach(function(t) {
      sheet.appendRow([
        t.id, t.title, t.projectId, t.projectCode, t.projectName, t.assignedBy,
        t.assigneeId, t.assigneeName, t.assigneeRole, t.assigneeRoleTitle,
        t.dueDate, t.priority, t.status, t.notes, t.createdAt
      ]);
    });

    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function deduplicateList(arr, key) {
  if (!arr || !Array.isArray(arr)) return [];
  var seen = {};
  var result = [];
  for (var i = 0; i < arr.length; i++) {
    var item = arr[i];
    if (!item) continue;
    var primaryKey = String(item[key] || item.id || item.username || item.code || item.name || '').toLowerCase().trim();
    if (primaryKey && !seen[primaryKey]) {
      seen[primaryKey] = true;
      result.push(item);
    }
  }
  return result;
}

/**
 * 12 Real Staff Credentials List
 */
function getSampleUsers() {
  return [
    { id: 'usr_mgr_1', username: 'Runlalit', password: '0423', name: 'รันต์ลลิต', role: 'manager', roleTitle: 'ผู้จัดการ (Manager)', avatar: 'https://images.unsplash.com/photo-1573496359142-b8d87734a5a2?w=100' },
    { id: 'usr_mgr_2', username: 'Pakpoom', password: '2512', name: 'ภาคภูมิ', role: 'manager', roleTitle: 'ผู้จัดการ (Manager)', avatar: 'https://images.unsplash.com/photo-1472099645785-5658abf4ff4e?w=100' },
    { id: 'usr_mgr_3', username: 'Charif', password: '0609', name: 'ชาลีฟ', role: 'manager', roleTitle: 'ผู้จัดการ (Manager)', avatar: 'https://images.unsplash.com/photo-1519085360753-af0119f7cbe7?w=100' },
    { id: 'usr_sales_1', username: 'kalaya', password: '201', name: 'กัลยา', role: 'sales', roleTitle: 'เจ้าหน้าที่ฝ่ายขาย (Sales)', avatar: 'https://images.unsplash.com/photo-1580489944761-15a19d654956?w=100' },
    { id: 'usr_sales_2', username: 'sunisa', password: '202', name: 'สุนิสา', role: 'sales', roleTitle: 'เจ้าหน้าที่ฝ่ายขาย (Sales)', avatar: 'https://images.unsplash.com/photo-1567532939604-b6b5b0db2604?w=100' },
    { id: 'usr_sales_3', username: 'Nopparat', password: '203', name: 'นพรัตน์', role: 'sales', roleTitle: 'เจ้าหน้าที่ฝ่ายขาย (Sales)', avatar: 'https://images.unsplash.com/photo-1507003211169-0a1dd7228f2d?w=100' },
    { id: 'usr_sales_4', username: 'Weeradon', password: '204', name: 'วีรดนย์', role: 'sales', roleTitle: 'เจ้าหน้าที่ฝ่ายขาย (Sales)', avatar: 'https://images.unsplash.com/photo-1500648767791-00dcc994a43e?w=100' },
    { id: 'usr_sales_5', username: 'Aprisara', password: '205', name: 'อภิสรา', role: 'sales', roleTitle: 'เจ้าหน้าที่ฝ่ายขาย (Sales)', avatar: 'https://images.unsplash.com/photo-1534528741775-53994a69daeb?w=100' },
    { id: 'usr_des_1', username: 'Pakawat', password: '301', name: 'ภัควัฒน์', role: 'designer', roleTitle: 'มัณฑนากร / ฝ่ายออกแบบ (Designer)', avatar: 'https://images.unsplash.com/photo-1539571696357-5a69c17a67c6?w=100' },
    { id: 'usr_foreman_1', username: 'Sirikorn', password: '401', name: 'ศิริกร', role: 'foreman', roleTitle: 'วิศวกรคุมงาน / โฟร์แมน (Foreman)', avatar: 'https://images.unsplash.com/photo-1522075469751-3a6694fb2f61?w=100' },
    { id: 'usr_foreman_2', username: 'Thiti', password: '402', name: 'ฐิติ', role: 'foreman', roleTitle: 'วิศวกรคุมงาน / โฟร์แมน (Foreman)', avatar: 'https://images.unsplash.com/photo-1506794778202-cad84cf45f1d?w=100' },
    { id: 'usr_foreman_3', username: 'Apichart', password: '403', name: 'อภิชาติ', role: 'foreman', roleTitle: 'วิศวกรคุมงาน / โฟร์แมน (Foreman)', avatar: 'https://images.unsplash.com/photo-1517841905240-472988babdf9?w=100' }
  ];
}

function getSampleProjects() {
  return [
    {
      id: 'proj_101',
      code: 'SRW-2026-001',
      customerName: 'คุณอนันต์ สุขเสริฐ',
      customerPhone: '081-234-5678',
      address: '123/45 หมู่บ้านปัญญา อ่อนนุช กรุงเทพฯ',
      projectType: 'ตกแต่งภายในบ้านเดี่ยว 2 ชั้น',
      budget: 850000,
      startDate: '2026-08-10',
      dueDate: '2026-08-25',
      status: 'design_review',
      salesName: 'กัลยา (Sales)',
      designerName: 'ภัควัฒน์ (Designer)',
      foremanName: 'ศิริกร (Foreman)',
      notes: 'ลูกค้าเน้นสไตล์ Modern Minimalist',
      comments: []
    }
  ];
}
