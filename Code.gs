const SPREADSHEET_ID   = '1IaEQLuCyzPwvshvqS36sEPqaHuRAw8p4Q-KZ2nsATRM';
const SHEET_NAME       = 'NY & LA';
const INVOICE_SHEET    = 'Contractor Invoices';
const ROOT_FOLDER_ID   = '1SaiY0KVhBgL-giWpCl5P47ZOQYKsbsUd';

const LOGISTICS_SS_ID     = '1ITOG8VPJa6U9W2WjWj79767ZQaCv8EAct2Uw5xOAZO4';
const LOGISTICS_IMG_FOLDER = 'Logistics Images';

const GALLERY_SHEET_ID  = '1ORhRQGnIALiDeAQ9Oa93k1s35zdFbqU8y3atMWR_FDk';
const SHIPPING_SHEET   = 'Shipping Log';

const BOOK_INVENTORY_SHEET = 'Book Inventory';
const BOOK_LOG_SHEET       = 'Book Log';
const BOOK_FOLDER_NAME     = 'Book Inventory';

// Tabs within the "Global-Data" spreadsheet
const RESELLER_CERT_SHEET  = 'Reseller Certificates';
const CR_SHEET             = 'Condition Reports';

// Finds the "Global-Data" spreadsheet by name (creates it if absent) and caches its ID.
function getGlobalDataSS() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('GLOBAL_DATA_SS_ID');
  if (ssId) {
    try { return SpreadsheetApp.openById(ssId); } catch(e) { props.deleteProperty('GLOBAL_DATA_SS_ID'); }
  }
  const token = ScriptApp.getOAuthToken();
  const q = "name='Global-Data' and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false";
  const resp  = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id)&pageSize=1',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const files = JSON.parse(resp.getContentText()).files || [];
  let ss;
  if (files.length) {
    ssId = files[0].id;
    ss = SpreadsheetApp.openById(ssId);
  } else {
    // Spreadsheet not found — create it
    ss = SpreadsheetApp.create('Global-Data');
    ssId = ss.getId();
  }
  props.setProperty('GLOBAL_DATA_SS_ID', ssId);
  return ss;
}

// Returns the named tab in Global-Data, creating it with headers if absent.
function getGlobalDataSheet(tabName, headers) {
  const ss = getGlobalDataSS();
  let sheet = ss.getSheetByName(tabName);
  if (!sheet) {
    sheet = ss.insertSheet(tabName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ─────────────────────────────────────────
function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'gallery') {
    const year = parseInt(e.parameter.year) || 2026;
    return serveGalleryData(year);
  }
  if (e && e.parameter && e.parameter.action === 'books') {
    return serveBookInventory();
  }
  if (e && e.parameter && e.parameter.action === 'logistics') {
    return serveLogistics(e.parameter.loc || 'LA');
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('SGG Expenses')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function formatSheetDate(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return m + '/' + d;
  }
  return String(val).trim();
}

function serveGalleryData(year) {
  try {
    const ss     = SpreadsheetApp.openById(GALLERY_SHEET_ID);
    const sheets = ss.getSheets();
    // Find tab by year name, fallback: 2026=index 0, 2027=index 1
    let sheet = sheets.find(s => s.getName().indexOf(String(year)) !== -1);
    if (!sheet) sheet = (year === 2027 && sheets.length > 1) ? sheets[1] : sheets[0];

    const data = sheet.getDataRange().getValues();
    const nyc = [], la = [];
    // Exclude financial summary rows and entries without dates
    // Uses specific phrases to avoid filtering show names like "Bennet Sales"
    const skipPattern = /\btotal\b|arbitrary\s+sales|sales\s+goal|\bconsignment\b|bennet\s+sales/i;

    for (let i = 1; i < data.length; i++) {
      const row     = data[i];
      const nyShow  = row[0] ? String(row[0]).trim() : '';
      const nyDates = formatSheetDate(row[1]);
      const laShow  = row[4] ? String(row[4]).trim() : '';
      const laDates = formatSheetDate(row[5]);
      // Include only if: has a name, has dates, and is not a summary row
      if (nyShow && !skipPattern.test(nyShow)) nyc.push({ show: nyShow, dates: nyDates });
      if (laShow && !skipPattern.test(laShow)) la.push({ show: laShow, dates: laDates });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ year: year, nyc: nyc, la: la }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const result = data.type === 'invoice'          ? submitInvoice(data)          :
                   data.type === 'shipping'         ? submitShipping(data)         :
                   data.type === 'book_log'         ? submitBookLog(data)          :
                   data.type === 'reseller-cert'    ? submitResellerCert(data)     :
                   data.type === 'cr-photos'        ? submitCRPhotos(data)         :
                   data.type === 'condition-report' ? submitConditionReport(data)  :
                   data.type === 'logistics-add'    ? submitLogisticsAdd(data)     :
                   data.type === 'logistics-update' ? submitLogisticsUpdate(data)  :
                   submitExpense(data);
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.TEXT);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────
// EXPENSES
// Folder path: Root → City → Receipts → Month Year
// ─────────────────────────────────────────
function submitExpense(data) {
  try {
    const token      = ScriptApp.getOAuthToken();
    const cityId     = findOrCreateFolder(token, ROOT_FOLDER_ID, data.location);
    const receiptsId = findOrCreateFolder(token, cityId, 'Receipts');
    const monthId    = findOrCreateFolder(token, receiptsId, getMonthYear(data.date));

    const uploadName = buildFileName(data.fileName, data.mimeType, data.vendor, data.date);
    const upload     = uploadFile(token, monthId, uploadName, data.mimeType, data.fileData);
    const fileUrl    = 'https://drive.google.com/file/d/' + upload.id + '/view';
    setPublicRead(token, upload.id);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error('Sheet "' + SHEET_NAME + '" not found');

    const displayDate = fmtDate(data.date);
    sheet.appendRow([
      data.account  || '',
      data.vendor   || '',
      displayDate,
      data.amount   ? parseFloat(data.amount) : '',
      data.category || '',
      data.location,
      upload.name,
      data.note     || ''
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 7).setFormula(
      '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + upload.name.replace(/"/g,'""') + '")'
    );

    logToSheet('EXPENSE OK: ' + (data.vendor || '?') + ' → ' + upload.name);
    return { success: true, fileUrl: fileUrl, fileName: upload.name, folder: data.location + ' / Receipts / ' + getMonthYear(data.date) };

  } catch (err) {
    logToSheet('EXPENSE ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// CONTRACTOR INVOICES
// Folder path: Root → City → Contractor Invoices → Contractor → Month Year
// ─────────────────────────────────────────
function submitInvoice(data) {
  try {
    const token        = ScriptApp.getOAuthToken();
    const cityId       = findOrCreateFolder(token, ROOT_FOLDER_ID, data.location);
    const invRootId    = findOrCreateFolder(token, cityId, 'Contractor Invoices');
    const contractorId = findOrCreateFolder(token, invRootId, data.contractor);
    const monthId      = findOrCreateFolder(token, contractorId, getMonthYear(data.date));

    const uploadName = buildInvoiceFileName(data.fileName, data.mimeType, data.contractor, data.invoiceNum, data.date);
    const upload     = uploadFile(token, monthId, uploadName, data.mimeType, data.fileData);
    const fileUrl    = 'https://drive.google.com/file/d/' + upload.id + '/view';
    setPublicRead(token, upload.id);

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet   = ss.getSheetByName(INVOICE_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(INVOICE_SHEET);
      sheet.appendRow(['Contractor', 'Invoice #', 'Date', 'Amount ($)', 'Description', 'City', 'File']);
      sheet.getRange(1, 1, 1, 7).setFontWeight('bold');
    }

    const displayDate = fmtDate(data.date);
    sheet.appendRow([
      data.contractor  || '',
      data.invoiceNum  || '',
      displayDate,
      data.amount      ? parseFloat(data.amount) : '',
      data.note        || '',
      data.location,
      upload.name
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 7).setFormula(
      '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + upload.name.replace(/"/g,'""') + '")'
    );

    logToSheet('INVOICE OK: ' + (data.contractor || '?') + ' → ' + upload.name);
    return { success: true, fileUrl: fileUrl, fileName: upload.name, folder: data.location + ' / Contractor Invoices / ' + data.contractor + ' / ' + getMonthYear(data.date) };

  } catch (err) {
    logToSheet('INVOICE ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}


// ─────────────────────────────────────────
// SHIPPING LOG
// Folder path: Root → Shipping Log → Year → Month Day
// ─────────────────────────────────────────
function submitShipping(data) {
  try {
    const token  = ScriptApp.getOAuthToken();
    const date   = new Date(data.date + 'T12:00:00');
    const year   = date.getFullYear().toString();
    const months = ['January','February','March','April','May','June',
                    'July','August','September','October','November','December'];
    const monthDay = months[date.getMonth()] + ' ' + date.getDate();

    // Folder: Root → Shipping Log → City → Year → Month Day
    // City derived from: SGNY→NYC, SGLA→LA, else infer from from/to
    const fromUp = (data.from || '').toUpperCase();
    const toUp   = (data.to   || '').toUpperCase();
    let   city   = 'NYC';
    if      (fromUp.indexOf('SGLA') !== -1 || toUp.indexOf('SGLA') !== -1) city = 'LA';
    else if (fromUp.indexOf('LA')   !== -1 || toUp.indexOf('LA')   !== -1) city = 'LA';
    const shippingRootId = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Shipping Log');
    const cityId         = findOrCreateFolder(token, shippingRootId, city);
    const yearId         = findOrCreateFolder(token, cityId, year);
    const dayId          = findOrCreateFolder(token, yearId, monthDay);

    let fileUrl = '', uploadName = '';
    if (data.fileData) {
      uploadName  = buildShippingFileName(data.fileName, data.mimeType, data.artwork, data.date);
      const upload = uploadFile(token, dayId, uploadName, data.mimeType, data.fileData);
      fileUrl     = 'https://drive.google.com/file/d/' + upload.id + '/view';
      setPublicRead(token, upload.id);
    }

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHIPPING_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(SHIPPING_SHEET);
      sheet.appendRow(['Date','Direction','Artist','Artwork','From','Via','To','Notes','Photo']);
      sheet.getRange(1,1,1,9).setFontWeight('bold');
    }

    const displayDate = fmtDate(data.date);
    sheet.appendRow([
      displayDate,
      data.direction || '',
      data.artist    || '',
      data.artwork   || '',
      data.from      || '',
      data.via       || '',
      data.to        || '',
      data.notes     || '',
      uploadName     || ''
    ]);
    if (fileUrl) {
      const lastRow = sheet.getLastRow();
      sheet.getRange(lastRow, 9).setFormula(
        '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + (uploadName || 'View Photo').replace(/"/g,'""') + '")'
      );
    }

    logToSheet('SHIPPING OK: ' + (data.direction||'?') + ' ' + (data.artwork||'?'));
    return { success: true, fileUrl: fileUrl, folder: 'Shipping Log / ' + city + ' / ' + year + ' / ' + monthDay };

  } catch (err) {
    logToSheet('SHIPPING ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

function buildShippingFileName(originalName, mimeType, artwork, date) {
  const generics = ['file','image','image.jpg','image.jpeg','image.png','image.heic','photo','photo.jpg'];
  if (originalName && !generics.includes(originalName.toLowerCase()) && originalName.length >= 5) return originalName;
  const ext = (mimeType || '').split('/')[1] || 'jpg';
  const art = (artwork || 'artwork').replace(/[^a-zA-Z0-9]/g,'_').substring(0,30);
  return art + '_' + date + '.' + ext;
}

// ─────────────────────────────────────────
// RESELLER CERTIFICATES
// Folder path: Root → Reseller Certificates
// Sheet: "Reseller Certificates" in Global-Data (or SPREADSHEET_ID)
// ─────────────────────────────────────────
function submitResellerCert(data) {
  try {
    const token    = ScriptApp.getOAuthToken();
    const folderId = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Reseller Certificates');

    const safeClient = (data.clientName || 'Client').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fileName   = safeClient + '_' + (data.date || '') + '.pdf';
    const upload     = uploadFile(token, folderId, fileName, 'application/pdf', data.fileData);
    const fileUrl    = 'https://drive.google.com/file/d/' + upload.id + '/view';
    setPublicRead(token, upload.id);

    const sheet = getGlobalDataSheet(RESELLER_CERT_SHEET, ['Date', 'Client', 'File']);

    const displayDate = data.date ? fmtDate(data.date) : '';
    sheet.appendRow([displayDate, data.clientName || '', fileName]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 3).setFormula(
      '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + fileName.replace(/"/g,'""') + '")'
    );

    logToSheet('RESELLER-CERT OK: ' + (data.clientName || '?') + ' → ' + fileName);
    return { success: true, fileUrl: fileUrl };

  } catch (err) {
    logToSheet('RESELLER-CERT ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// CONDITION REPORTS — step 1: upload photos
// Folder path: Root → Condition Reports → stockNum
// ─────────────────────────────────────────
function submitCRPhotos(data) {
  try {
    const token     = ScriptApp.getOAuthToken();
    const crRootId  = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Condition Reports');
    const stockId   = findOrCreateFolder(token, crRootId, data.stockNum || 'Unknown');

    const photoUrls = {};
    const photos    = data.photos || {};
    for (const label in photos) {
      const photo    = photos[label];
      if (!photo || !photo.fileData) continue;
      const ext      = (photo.mimeType || 'image/jpeg').split('/')[1] || 'jpg';
      const safeLbl  = label.replace(/[^a-zA-Z0-9]/g, '_');
      const fileName = (data.stockNum || 'CR') + '_' + safeLbl + '.' + ext;
      const upload   = uploadFile(token, stockId, fileName, photo.mimeType, photo.fileData);
      const fileUrl  = 'https://drive.google.com/file/d/' + upload.id + '/view';
      setPublicRead(token, upload.id);
      photoUrls[label] = fileUrl;
    }

    logToSheet('CR-PHOTOS OK: ' + (data.stockNum || '?') + ' — ' + Object.keys(photoUrls).length + ' photos');
    return { success: true, photoUrls: photoUrls };

  } catch (err) {
    logToSheet('CR-PHOTOS ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// CONDITION REPORTS — step 2: upload PDF + log
// Folder path: Root → Condition Reports (PDF goes here)
// Sheet: "Condition Reports" in Global-Data (or SPREADSHEET_ID)
// ─────────────────────────────────────────
function submitConditionReport(data) {
  try {
    const token    = ScriptApp.getOAuthToken();
    const crRootId = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Condition Reports');

    const safeArtist = (data.artist || 'Artist').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fileName   = (data.stockNum || 'CR') + '_' + safeArtist + '_CR_' + (data.dateIn || '') + '.pdf';
    const upload     = uploadFile(token, crRootId, fileName, 'application/pdf', data.pdfData);
    const fileUrl    = 'https://drive.google.com/file/d/' + upload.id + '/view';
    setPublicRead(token, upload.id);

    const sheet = getGlobalDataSheet(CR_SHEET, ['Date In', 'Stock #', 'Artist', 'Title', 'Condition', 'Signed By', 'CR Document']);

    const displayDate = data.dateIn ? fmtDate(data.dateIn) : '';
    sheet.appendRow([
      displayDate,
      data.stockNum   || '',
      data.artist     || '',
      data.title      || '',
      data.condition  || '',
      data.signedBy   || '',
      fileName
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 7).setFormula(
      '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + fileName.replace(/"/g,'""') + '")'
    );

    logToSheet('CR OK: ' + (data.stockNum || '?') + ' ' + (data.artist || '?') + ' → ' + fileName);
    return { success: true, pdfUrl: fileUrl };

  } catch (err) {
    logToSheet('CR ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// DRIVE HELPERS (REST API — bypasses DriveApp scope issue)
// ─────────────────────────────────────────
function findOrCreateFolder(token, parentId, name) {
  const q = "mimeType='application/vnd.google-apps.folder'" +
            " and name='" + name.replace(/'/g, "\\'") + "'" +
            " and '" + parentId + "' in parents" +
            " and trashed=false";
  const resp  = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id)',
    { headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const files = JSON.parse(resp.getContentText()).files;
  if (files && files.length > 0) return files[0].id;

  const create = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?fields=id',
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ name: name, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] }),
      muteHttpExceptions: true
    }
  );
  return JSON.parse(create.getContentText()).id;
}

function uploadFile(token, folderId, fileName, mimeType, base64Data) {
  const boundary = 'SGGBoundary';
  const resp = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name',
    {
      method: 'POST',
      headers: {
        'Authorization': 'Bearer ' + token,
        'Content-Type': 'multipart/related; boundary=' + boundary
      },
      payload: '--' + boundary + '\r\n' +
               'Content-Type: application/json; charset=UTF-8\r\n\r\n' +
               JSON.stringify({ name: fileName, parents: [folderId] }) +
               '\r\n--' + boundary + '\r\n' +
               'Content-Type: ' + (mimeType || 'image/jpeg') + '\r\n' +
               'Content-Transfer-Encoding: base64\r\n\r\n' +
               base64Data +
               '\r\n--' + boundary + '--',
      muteHttpExceptions: true
    }
  );
  if (resp.getResponseCode() !== 200) {
    throw new Error('Upload failed (' + resp.getResponseCode() + '): ' + resp.getContentText());
  }
  return JSON.parse(resp.getContentText()); // { id, name }
}

function setPublicRead(token, fileId) {
  UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files/' + fileId + '/permissions',
    {
      method: 'POST',
      headers: { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ role: 'reader', type: 'anyone' }),
      muteHttpExceptions: true
    }
  );
}

// ─────────────────────────────────────────
// UTILITIES
// ─────────────────────────────────────────
function buildFileName(originalName, mimeType, vendor, date) {
  const generics = ['file','image','image.jpg','image.jpeg','image.png','image.heic','photo','photo.jpg'];
  if (originalName && !generics.includes(originalName.toLowerCase()) && originalName.length >= 5) {
    return originalName;
  }
  const ext    = (mimeType || '').split('/')[1] || 'jpg';
  const vendor_ = (vendor || 'receipt').replace(/[^a-zA-Z0-9]/g, '_');
  return vendor_ + '_' + date + '.' + ext;
}

function buildInvoiceFileName(originalName, mimeType, contractor, invoiceNum, date) {
  const generics = ['file','image','image.jpg','image.jpeg','image.png','image.heic','photo','photo.jpg'];
  if (originalName && !generics.includes(originalName.toLowerCase()) && originalName.length >= 5) {
    return originalName;
  }
  const ext  = (mimeType || '').split('/')[1] || 'pdf';
  const name = (contractor || 'invoice').replace(/[^a-zA-Z0-9]/g, '_');
  const num  = invoiceNum ? '_' + invoiceNum.replace(/[^a-zA-Z0-9]/g, '_') : '';
  return name + num + '_' + date + '.' + ext;
}

function fmtDate(dateStr) {
  const p = dateStr.split('-');
  return parseInt(p[1]) + '/' + parseInt(p[2]) + '/' + p[0];
}

function getMonthYear(dateStr) {
  const p = dateStr.split('-');
  const d = new Date(parseInt(p[0]), parseInt(p[1]) - 1, parseInt(p[2]));
  const months = ['January','February','March','April','May','June',
                  'July','August','September','October','November','December'];
  return months[d.getMonth()] + ' ' + d.getFullYear();
}

function logToSheet(message) {
  try {
    const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
    let log   = ss.getSheetByName('_Debug');
    if (!log) log = ss.insertSheet('_Debug');
    log.appendRow([new Date(), message]);
  } catch(e) {}
}

// ─────────────────────────────────────────
// BOOK INVENTORY LOG
// Sheet: Book Inventory (master stock)
// Sheet: Book Log (transaction history)
// Drive: Root → Book Inventory → NYC|LA → [Month Year] (Google Sheet)
// ─────────────────────────────────────────

// GET action=books — return full book list with NYC/LA stock
function serveBookInventory() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(BOOK_INVENTORY_SHEET);
    if (!sheet) return ContentService
      .createTextOutput(JSON.stringify({ error: 'Book Inventory sheet not found' }))
      .setMimeType(ContentService.MimeType.JSON);

    const data  = sheet.getDataRange().getValues();
    const books = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      books.push({ title: String(row[0]).trim(), nyc: parseInt(row[1]) || 0, la: parseInt(row[2]) || 0 });
    }
    return ContentService
      .createTextOutput(JSON.stringify({ books: books }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// POST type=book_log — decrement stock, log transaction, update monthly report
function submitBookLog(data) {
  try {
    const ss     = SpreadsheetApp.openById(SPREADSHEET_ID);
    const loc    = (data.location || 'NYC').toUpperCase();
    const qty    = Math.abs(parseInt(data.qty) || 1);
    const colIdx = loc === 'LA' ? 3 : 2; // B=NYC(col 2), C=LA(col 3), 1-indexed

    // 1. Find book and decrement stock
    const invSheet = ss.getSheetByName(BOOK_INVENTORY_SHEET);
    if (!invSheet) throw new Error('Book Inventory sheet not found');
    const invData  = invSheet.getDataRange().getValues();
    let   bookRow  = -1;
    const bookName = String(data.book || '').trim().toLowerCase();
    for (let i = 1; i < invData.length; i++) {
      if (invData[i][0] && String(invData[i][0]).trim().toLowerCase() === bookName) {
        bookRow = i + 1; break;
      }
    }
    if (bookRow === -1) throw new Error('Book not found: ' + data.book);
    const currentStock = parseInt(invSheet.getRange(bookRow, colIdx).getValue()) || 0;
    const newStock     = Math.max(0, currentStock - qty);
    invSheet.getRange(bookRow, colIdx).setValue(newStock);

    // 2. Append to Book Log sheet
    let logSheet = ss.getSheetByName(BOOK_LOG_SHEET);
    if (!logSheet) {
      logSheet = ss.insertSheet(BOOK_LOG_SHEET);
      logSheet.appendRow(['Date','Book Title','Location','Type','Client','Qty','Price','Notes','Stock After']);
      logSheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#f3f3f3');
      logSheet.setFrozenRows(1);
    }
    logSheet.appendRow([
      fmtDate(data.date),
      data.book   || '',
      loc,
      data.txType || '',
      data.client || '',
      qty,
      data.price  || '',
      data.notes  || '',
      newStock
    ]);

    // 3. Create/update monthly Drive report
    updateMonthlyBookReport(data, loc, qty, newStock);

    logToSheet('BOOK OK: ' + (data.txType||'?') + ' — ' + (data.book||'?') + ' → ' + loc);
    return { success: true, newStock: newStock };

  } catch (err) {
    logToSheet('BOOK ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

function updateMonthlyBookReport(data, loc, qty, newStock) {
  const token     = ScriptApp.getOAuthToken();
  const monthYear = getMonthYear(data.date); // e.g. "April 2026"

  // Drive folder: Root → Book Inventory → NYC|LA
  const bookRootId = findOrCreateFolder(token, ROOT_FOLDER_ID, BOOK_FOLDER_NAME);
  const cityId     = findOrCreateFolder(token, bookRootId, loc);

  // Find or create the monthly Google Sheet for this city
  const q = "name='" + (monthYear + ' \u2014 ' + loc).replace(/'/g,"\\'") + "'" +
            " and '" + cityId + "' in parents" +
            " and mimeType='application/vnd.google-apps.spreadsheet'" +
            " and trashed=false";
  const searchResp = UrlFetchApp.fetch(
    'https://www.googleapis.com/drive/v3/files?q=' + encodeURIComponent(q) + '&fields=files(id)',
    { headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
  );
  const found = JSON.parse(searchResp.getContentText()).files || [];

  let fileId;
  if (found.length > 0) {
    fileId = found[0].id;
  } else {
    // Create new Google Sheet named e.g. "April 2026 — NYC"
    const createResp = UrlFetchApp.fetch('https://sheets.googleapis.com/v4/spreadsheets', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ properties: { title: monthYear + ' \u2014 ' + loc } }),
      muteHttpExceptions: true
    });
    fileId = JSON.parse(createResp.getContentText()).spreadsheetId;

    // Move into city folder
    UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId +
      '?addParents=' + cityId + '&removeParents=root&fields=id',
      { method: 'PATCH', headers: { Authorization: 'Bearer ' + token }, muteHttpExceptions: true }
    );

    // Write header row
    UrlFetchApp.fetch(
      'https://sheets.googleapis.com/v4/spreadsheets/' + fileId +
      '/values/Sheet1!A1:H1?valueInputOption=RAW',
      {
        method: 'PUT',
        headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ values: [['Date','Book Title','Type','Client','Qty','Price','Notes','Stock After']] }),
        muteHttpExceptions: true
      }
    );
  }

  // Append this transaction
  UrlFetchApp.fetch(
    'https://sheets.googleapis.com/v4/spreadsheets/' + fileId +
    '/values/Sheet1!A:H:append?valueInputOption=RAW&insertDataOption=INSERT_ROWS',
    {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      payload: JSON.stringify({ values: [[
        fmtDate(data.date),
        data.book   || '',
        data.txType || '',
        data.client || '',
        qty,
        data.price  || '',
        data.notes  || '',
        newStock
      ]] }),
      muteHttpExceptions: true
    }
  );
}

// ─────────────────────────────────────────
// LOGISTICS MODULE
// ─────────────────────────────────────────
function serveLogistics(loc) {
  try {
    const ss    = SpreadsheetApp.openById(LOGISTICS_SS_ID);
    const sheet = ss.getSheetByName(loc.toUpperCase());
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);
    const data = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0] && !r[2]) continue;
      rows.push({ row: i+1, client: String(r[0]||''), imageUrl: String(r[1]||''), title: String(r[2]||''), invoice: String(r[3]||''), price: r[4]||'', salePrice: r[5]||'', salesPayout: r[6]||'', artistPaid: String(r[7]||''), galleryPaid: String(r[8]||''), notes: String(r[9]||'') });
    }
    return ContentService.createTextOutput(JSON.stringify({rows})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({error:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitLogisticsAdd(data) {
  try {
    const token = ScriptApp.getOAuthToken();
    const ss    = SpreadsheetApp.openById(LOGISTICS_SS_ID);
    const sheet = ss.getSheetByName((data.loc||'LA').toUpperCase());
    let imageUrl = '';
    if (data.imageData) {
      const fId = findOrCreateFolder(token, ROOT_FOLDER_ID, LOGISTICS_IMG_FOLDER);
      const ext  = (data.imageMime||'image/jpeg').split('/')[1]||'jpg';
      const name = (data.client||'artwork').replace(/[^a-zA-Z0-9]/g,'_')+'_'+Date.now()+'.'+ext;
      const up   = uploadFile(token, fId, name, data.imageMime||'image/jpeg', data.imageData);
      setPublicRead(token, up.id);
      imageUrl = 'https://drive.google.com/uc?export=view&id='+up.id;
    }
    sheet.appendRow([data.client||'', imageUrl, data.title||'', data.invoice||'', data.price||'', data.salePrice||'', data.salesPayout||'', data.artistPaid||'', data.galleryPaid||'', data.notes||'']);
    return { success: true };
  } catch(err) { return { success:false, error:err.toString() }; }
}

function submitLogisticsUpdate(data) {
  try {
    const token = ScriptApp.getOAuthToken();
    const ss    = SpreadsheetApp.openById(LOGISTICS_SS_ID);
    const sheet = ss.getSheetByName((data.loc||'LA').toUpperCase());
    const row   = parseInt(data.row);
    let imageUrl = data.existingImageUrl || '';
    if (data.imageData) {
      const fId  = findOrCreateFolder(token, ROOT_FOLDER_ID, LOGISTICS_IMG_FOLDER);
      const ext  = (data.imageMime||'image/jpeg').split('/')[1]||'jpg';
      const name = (data.client||'artwork').replace(/[^a-zA-Z0-9]/g,'_')+'_'+Date.now()+'.'+ext;
      const up   = uploadFile(token, fId, name, data.imageMime||'image/jpeg', data.imageData);
      setPublicRead(token, up.id);
      imageUrl = 'https://drive.google.com/uc?export=view&id='+up.id;
    }
    sheet.getRange(row,1,1,10).setValues([[data.client||'', imageUrl, data.title||'', data.invoice||'', data.price||'', data.salePrice||'', data.salesPayout||'', data.artistPaid||'', data.galleryPaid||'', data.notes||'']]);
    return { success: true };
  } catch(err) { return { success:false, error:err.toString() }; }
}

// Run once from Apps Script editor → reformats Book Inventory sheet with proper headers
function setupBookInventorySheet() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(BOOK_INVENTORY_SHEET);
  if (!sheet) sheet = ss.insertSheet(BOOK_INVENTORY_SHEET);

  sheet.getRange(1,1,1,5).setValues([['Book Title','NYC Stock','LA Stock','Total','Notes']]);
  sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#f3f3f3');
  sheet.setFrozenRows(1);

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    for (let r = 2; r <= lastRow; r++) {
      sheet.getRange(r, 4).setFormula('=B' + r + '+C' + r);
    }
  }
  sheet.setColumnWidth(1, 200); sheet.setColumnWidth(2, 100);
  sheet.setColumnWidth(3, 100); sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 200);
  SpreadsheetApp.flush();
  Logger.log('Book Inventory sheet restructured.');
}
