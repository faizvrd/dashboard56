/**
 * ============================================================
 *  ESR - AROHERA PROJECT DASHBOARD — Backend (Google Apps Script)
 *  Cocok dengan index.html yang sudah berisi tampilan Calendar
 *  (menggantikan Action Log) beserta halaman PO Tracking, FAT &
 *  Inspection, Shipment Tracking, dan User Management.
 * ============================================================
 *  CARA PASANG:
 *  1. Buat/gunakan Google Sheet yang sudah berisi sheet:
 *     PO_Tracking, FAT_Schedule, Shipment_Tracking, Action_Log,
 *     User_Management  (lihat file "ESR_Arohera_Template.xlsx"
 *     untuk struktur kolom yang wajib sama persis).
 *  2. Extensions > Apps Script pada Google Sheet tersebut.
 *  3. Tempel isi file ini sebagai "Code.gs".
 *  4. Deploy > New deployment > Web app.
 *     - Execute as: Me
 *     - Who has access: Anyone
 *  5. Salin URL web app yang muncul (".../exec") lalu tempelkan
 *     ke variabel `apiUrl` pada array PROJECTS di index.html
 *     (dan/atau ke AUTH_API_URL jika sheet ini juga menyimpan
 *     User_Management).
 * ============================================================
 */

// ---------- KONFIGURASI ----------
const SHEETS = {
  PO_Tracking: 'purchaseOrderId',
  FAT_Schedule: 'FAT_ID',
  Shipment_Tracking: 'Shipment_ID',
  Action_Log: 'Action_ID',
  User_Management: 'User_ID',
  Milestones: 'Milestone_ID',
};

const DRIVE_FOLDER_NAME = 'ESR Arohera Uploads';

// ---------- ENTRY POINTS ----------
function doGet(e) {
  try {
    const action = e.parameter.action;
    const result = routeAction_(action, e.parameter);
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const result = routeAction_(body.action, body.payload || {});
    return jsonOut_(result);
  } catch (err) {
    return jsonOut_({ success: false, error: err.message });
  }
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(
    ContentService.MimeType.JSON
  );
}

// ---------- ROUTER ----------
function routeAction_(action, payload) {
  switch (action) {
    case 'login':
      return doLogin_(payload);
    case 'getAllPO':
      return { success: true, data: sheetToObjects_('PO_Tracking') };
    case 'getAllFAT':
      return { success: true, data: sheetToObjects_('FAT_Schedule') };
    case 'getAllShipments':
      return { success: true, data: sheetToObjects_('Shipment_Tracking') };
    case 'getAllActions':
      return { success: true, data: sheetToObjects_('Action_Log') };
    case 'getAllUsers':
      return { success: true, data: sheetToObjects_('User_Management') };
    case 'getAllMilestones':
      return { success: true, data: sheetExistsOrEmpty_('Milestones') };
    case 'getDashboardData':
      return { success: true, data: getDashboardData_(Number(payload.upcomingDays) || 30) };
    case 'getCurrencyRates':
      return { success: true, data: { USD: 1, IDR: 16300, JPY: 157 } };
    case 'addRow':
      return addRow_(payload.sheet, payload.rowData);
    case 'updateRow':
      return updateRow_(payload.sheet, payload.keyColumn, payload.keyValue, payload.newData);
    case 'deleteRow':
      return deleteRow_(payload.sheet, payload.keyColumn, payload.keyValue);
    case 'closeAction':
      return closeAction_(payload.actionId);
    case 'uploadFileToDrive':
      return uploadFileToDrive_(payload.base64Data, payload.fileName, payload.mimeType);
    default:
      return { success: false, error: 'Unknown action: ' + action };
  }
}

// ---------- HELPERS ----------
function ss_() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheet_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan.');
  return sheet;
}

function sheetExistsOrEmpty_(name) {
  const sheet = ss_().getSheetByName(name);
  return sheet ? sheetToObjects_(name) : [];
}

function sheetToObjects_(name) {
  const sheet = getSheet_(name);
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values
    .filter(row => row.join('') !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => {
        let v = row[i];
        if (v instanceof Date) v = formatDate_(v);
        obj[h] = v;
      });
      return obj;
    });
}

function formatDate_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

// ---------- LOGIN ----------
function doLogin_(payload) {
  const users = sheetToObjects_('User_Management');
  const found = users.find(
    u =>
      String(u.Username).trim() === String(payload.username).trim() &&
      String(u.Password) === String(payload.password)
  );
  if (!found) return { success: false, error: 'Username atau password salah.' };
  if (String(found.Status) === 'Inactive') {
    return { success: false, error: 'Akun tidak aktif. Hubungi admin.' };
  }
  return {
    success: true,
    role: found.Role,
    name: found.Name,
    project: found.Project_ID,
  };
}

// ---------- DASHBOARD KPI ----------
function getDashboardData_(upcomingDays) {
  const po = sheetToObjects_('PO_Tracking');
  const fat = sheetToObjects_('FAT_Schedule');
  const actions = sheetToObjects_('Action_Log');

  const criticalItems = po.filter(
    p => p.category === 'At Risk' || p.category === 'Delay'
  ).length;

  const now = new Date();
  const limit = new Date(now.getTime() + upcomingDays * 24 * 60 * 60 * 1000);
  const upcomingFAT = fat.filter(f => {
    if (!f.FAT_Date) return false;
    const d = new Date(f.FAT_Date);
    return !isNaN(d.getTime()) && d >= now && d <= limit && f.Status !== 'Complete';
  }).length;

  const openActions = actions.filter(
    a => a.Status === 'OPEN' || a.Status === 'PENDING'
  ).length;

  return {
    kpi: { criticalItems, upcomingFAT, openActions },
  };
}

// ---------- GENERIC CRUD ----------
function addRow_(sheetName, rowData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_(sheetName);
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    const keyCol = SHEETS[sheetName];

    // Auto-generate ID kalau kolom key kosong
    if (keyCol && (!rowData[keyCol] || String(rowData[keyCol]).trim() === '')) {
      rowData[keyCol] = generateId_(sheetName);
    }

    const row = headers.map(h => (rowData.hasOwnProperty(h) ? rowData[h] : ''));
    sheet.appendRow(row);
    return { success: true, message: 'Data berhasil disimpan.', id: rowData[keyCol] };
  } finally {
    lock.releaseLock();
  }
}

function updateRow_(sheetName, keyColumn, keyValue, newData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const keyIdx = headers.indexOf(keyColumn);
    if (keyIdx === -1) throw new Error('Kolom key "' + keyColumn + '" tidak ditemukan.');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][keyIdx]) === String(keyValue)) {
        headers.forEach((h, c) => {
          if (newData.hasOwnProperty(h)) {
            sheet.getRange(r + 1, c + 1).setValue(newData[h]);
          }
        });
        return { success: true, message: 'Data berhasil diupdate.' };
      }
    }
    return { success: false, error: 'Data dengan ' + keyColumn + '=' + keyValue + ' tidak ditemukan.' };
  } finally {
    lock.releaseLock();
  }
}

function deleteRow_(sheetName, keyColumn, keyValue) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sheet = getSheet_(sheetName);
    const values = sheet.getDataRange().getValues();
    const headers = values[0];
    const keyIdx = headers.indexOf(keyColumn);
    if (keyIdx === -1) throw new Error('Kolom key "' + keyColumn + '" tidak ditemukan.');

    for (let r = 1; r < values.length; r++) {
      if (String(values[r][keyIdx]) === String(keyValue)) {
        sheet.deleteRow(r + 1);
        return { success: true, message: 'Data berhasil dihapus.' };
      }
    }
    return { success: false, error: 'Data tidak ditemukan.' };
  } finally {
    lock.releaseLock();
  }
}

function generateId_(sheetName) {
  const prefix = {
    PO_Tracking: 'PO',
    FAT_Schedule: 'FAT',
    Shipment_Tracking: 'SH',
    Action_Log: 'AC',
    User_Management: 'U',
    Milestones: 'MS',
  }[sheetName] || 'ID';
  return prefix + new Date().getTime();
}

// ---------- ACTION LOG / CALENDAR SPECIFIC ----------
function closeAction_(actionId) {
  return updateRow_('Action_Log', 'Action_ID', actionId, {
    Status: 'CLOSED',
    Closed_Date: formatDate_(new Date()),
  });
}

// ---------- FILE UPLOAD ----------
function uploadFileToDrive_(base64Data, fileName, mimeType) {
  try {
    const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
    const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(DRIVE_FOLDER_NAME);
    const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return { success: true, url: file.getUrl() };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
