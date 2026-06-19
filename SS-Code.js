const SPREADSHEET_ID   = '1IaEQLuCyzPwvshvqS36sEPqaHuRAw8p4Q-KZ2nsATRM';
const SHEET_NAME       = 'Expenses';
const INVOICE_SHEET    = 'Contractor Invoices';
const ROOT_FOLDER_ID   = '1SaiY0KVhBgL-giWpCl5P47ZOQYKsbsUd';
const CR_FOLDER_ID     = '1EWYGMqhFwNRtT_jWGXNBY2XB7FQvmgtc'; // ss@sebastiangladstone Drive

const LOGISTICS_SS_ID     = '1ITOG8VPJa6U9W2WjWj79767ZQaCv8EAct2Uw5xOAZO4';
const LOGISTICS_IMG_FOLDER = 'Logistics Images';

const GALLERY_SHEET_ID  = '1ORhRQGnIALiDeAQ9Oa93k1s35zdFbqU8y3atMWR_FDk';
const SHIPPING_SHEET   = 'Shipping Log';

const EXHIBITIONS_CAL_ID = 'c_1062df1d655deaa79281e515b51573f79a5fecc969a82d168339f6f6094d2dd5@group.calendar.google.com';

const BOOK_INVENTORY_SHEET = 'Book Inventory';
const BOOK_LOG_SHEET       = 'Book Log';
const BOOK_FOLDER_NAME     = 'Book Inventory';

// Tabs for Condition Reports and Reseller Certificates — stored in the main SPREADSHEET_ID sheet
const RESELLER_CERT_SHEET  = 'Reseller Certificates';
const CR_SHEET             = 'Condition Reports';

// Returns the named tab in the main spreadsheet (SPREADSHEET_ID), creating it with headers if absent.
function getGlobalDataSheet(tabName, headers) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(tabName);
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
  if (e && e.parameter && e.parameter.action === 'rebuild-logistics-cache') {
    buildLogisticsCache('LA');
    buildLogisticsCache('NY');
    return ContentService.createTextOutput(JSON.stringify({ok:true})).setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'handbook') {
    return ContentService
      .createTextOutput(JSON.stringify(getHandbook()))
      .setMimeType(ContentService.MimeType.JSON);
  }
  if (e && e.parameter && e.parameter.action === 'artists') {
    return serveArtists();
  }
  if (e && e.parameter && e.parameter.action === 'signatures') {
    return serveSignatures();
  }
  if (e && e.parameter && e.parameter.action === 'contractors') {
    return serveContractors();
  }
  if (e && e.parameter && e.parameter.action === 'ic') {
    return serveIC();
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
                   data.type === 'save_backtag'      ? submitBacktag(data)            :
                   data.type === 'artists-update'     ? submitArtistsUpdate(data)      :
                   data.type === 'signatures-update'   ? submitSignaturesUpdate(data)     :
                   data.type === 'contractors-update' ? submitContractorsUpdate(data)   :
                   data.type === 'parse_bol'          ? parseBOL(data)                  :
                   data.type === 'save_bol'           ? saveBOL(data)                   :
                   data.type === 'ic-update'          ? submitICUpdate(data)             :
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
    const expId      = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Expenses');
    const cityId     = findOrCreateFolder(token, expId, data.location);
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
// Folder path: Root → Expenses → City → Contractor Invoices → Contractor → Month Year
// ─────────────────────────────────────────
function submitInvoice(data) {
  try {
    const token        = ScriptApp.getOAuthToken();
    const expId        = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Expenses');
    const cityId       = findOrCreateFolder(token, expId, data.location);
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
      sheet.appendRow(['Date','Direction','Artist','Artwork','From','Via','To','Notes','Photo','✓']);
      sheet.getRange(1,1,1,10).setFontWeight('bold');
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
    const lastRow = sheet.getLastRow();
    if (fileUrl) {
      sheet.getRange(lastRow, 9).setFormula(
        '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + (uploadName || 'View Photo').replace(/"/g,'""') + '")'
      );
    }
    sheet.getRange(lastRow, 10).insertCheckboxes();

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
// BACK TAGS
// Folder path: Root → Back Tags
// ─────────────────────────────────────────
function submitBacktag(data) {
  try {
    const token    = ScriptApp.getOAuthToken();
    const folderId = findOrCreateFolder(token, ROOT_FOLDER_ID, 'Back Tags');
    const fileName = data.fileName || ('backtags_' + new Date().toISOString().slice(0,10) + '.pdf');
    const upload   = uploadFile(token, folderId, fileName, 'application/pdf', data.fileData);
    const fileUrl  = 'https://drive.google.com/file/d/' + upload.id + '/view';
    logToSheet('BACKTAG OK: ' + fileName);
    return { success: true, fileUrl: fileUrl };
  } catch (err) {
    logToSheet('BACKTAG ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// RESELLER CERTIFICATES
// Folder path: Root → Reseller Certificates
// Sheet: "Reseller Certificates" in SGG-Global-Data
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
    const stockId   = findOrCreateFolder(token, CR_FOLDER_ID, data.stockNum || 'Unknown');

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
// Sheet: "Condition Reports" in SGG-Global-Data
// ─────────────────────────────────────────
function submitConditionReport(data) {
  try {
    const token    = ScriptApp.getOAuthToken();

    const safeArtist = (data.artist || 'Artist').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fileName   = (data.stockNum || 'CR') + '_' + safeArtist + '_CR_' + (data.dateIn || '') + '.pdf';
    const upload     = uploadFile(token, CR_FOLDER_ID, fileName, 'application/pdf', data.pdfData);
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
  if (!dateStr) return '';
  const p = String(dateStr).split('-');
  if (p.length < 3) return String(dateStr);
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

// Extract Drive file ID from any Drive URL format
function extractDriveFileId(url) {
  if (!url) return null;
  var m = url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  m = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (m) return m[1];
  return null;
}

// Convert any Drive URL to the embeddable uc?export=view format
function normalizeDriveUrl(url, token) {
  if (!url) return '';
  if (url.includes('uc?export=view')) return url;
  var fileId = extractDriveFileId(url);
  if (!fileId) return url;
  // Make publicly readable so the app can display it without auth
  try { setPublicRead(token, fileId); } catch(e) {}
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

// ─── Logistics Cache ─────────────────────────────────────────────────────────

function buildLogisticsCache(loc) {
  try {
    const token = ScriptApp.getOAuthToken();
    const ss    = SpreadsheetApp.openById(LOGISTICS_SS_ID);
    const sheet = ss.getSheetByName(loc.toUpperCase());
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const rows = [];
    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      if (!r[0] && !r[2] && !r[3]) continue;
      const imageUrl = normalizeDriveUrl(String(r[1]||''), token);
      rows.push({
        row:                  i+1,
        client:               String(r[0]||''),
        imageUrl,
        artist:               String(r[2]||''),
        title:                String(r[3]||''),
        year:                 String(r[4]||''),
        sku:                  String(r[5]||''),
        dimensions:           String(r[6]||''),
        material:             String(r[7]||''),
        invoice:              String(r[8]||''),
        price:                r[9]||'',
        salePrice:            r[10]||'',
        artistPayoutAmount:   r[11]||'',
        artistPaymentStatus:  String(r[12]||''),
        artistPayoutRecords:  String(r[13]||''),
        galleryPaid:          String(r[14]||''),
        logisticsStatus:      String(r[15]||''),
        dueDate:              String(r[16]||''),
        notes:                String(r[17]||''),
      });
    }
    CacheService.getScriptCache().put('logistics_' + loc.toUpperCase(), JSON.stringify({rows}), 21600);
    Logger.log('✅ Logistics cache built for ' + loc + ' — ' + rows.length + ' rows');
  } catch(err) {
    Logger.log('❌ buildLogisticsCache(' + loc + '): ' + err.toString());
  }
}

// Installable trigger handler — fires when the Logistics sheet is edited by a user
function onLogisticsEdit(e) {
  buildLogisticsCache('LA');
  buildLogisticsCache('NY');
}

// Run once from the Apps Script editor to install the onEdit trigger on the Logistics spreadsheet
function installLogisticsTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'onLogisticsEdit') ScriptApp.deleteTrigger(t);
  });
  const ss = SpreadsheetApp.openById(LOGISTICS_SS_ID);
  ScriptApp.newTrigger('onLogisticsEdit').forSpreadsheet(ss).onEdit().create();
  Logger.log('✅ Logistics onEdit trigger installed on sheet ' + LOGISTICS_SS_ID);
}

function serveLogistics(loc) {
  try {
    // Serve from cache if available (populated by onEdit trigger or last write)
    const cached = CacheService.getScriptCache().get('logistics_' + loc.toUpperCase());
    if (cached) {
      return ContentService.createTextOutput(cached).setMimeType(ContentService.MimeType.JSON);
    }
    // Cache miss — build live and cache for next request
    buildLogisticsCache(loc);
    const fresh = CacheService.getScriptCache().get('logistics_' + loc.toUpperCase());
    if (fresh) {
      return ContentService.createTextOutput(fresh).setMimeType(ContentService.MimeType.JSON);
    }
    return ContentService.createTextOutput(JSON.stringify({rows:[]})).setMimeType(ContentService.MimeType.JSON);
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
    sheet.appendRow([
      data.client||'', imageUrl,
      data.artist||'', data.title||'', data.year||'', data.sku||'',
      data.dimensions||'', data.material||'',
      data.invoice||'', data.price||'', data.salePrice||'',
      data.artistPayoutAmount||'', data.artistPaymentStatus||'', data.artistPayoutRecords||'',
      data.galleryPaid||'', data.logisticsStatus||'', data.dueDate||'', data.notes||''
    ]);
    buildLogisticsCache(data.loc || 'LA');
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
    sheet.getRange(row,1,1,18).setValues([[
      data.client||'', imageUrl,
      data.artist||'', data.title||'', data.year||'', data.sku||'',
      data.dimensions||'', data.material||'',
      data.invoice||'', data.price||'', data.salePrice||'',
      data.artistPayoutAmount||'', data.artistPaymentStatus||'', data.artistPayoutRecords||'',
      data.galleryPaid||'', data.logisticsStatus||'', data.dueDate||'', data.notes||''
    ]]);
    buildLogisticsCache(data.loc || 'LA');
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

// ─────────────────────────────────────────
// COMPANY HANDBOOK
// Reads the handbook Google Doc and returns structured JSON
// ─────────────────────────────────────────
function getHandbook() {
  try {
    var doc  = DocumentApp.openById('1UyamIq8Hsgy4QUKm-X4-ghpO-FDpLeoh5qyRWVyCPrI');
    var body = doc.getBody();
    var items = [];
    var idCount = {};

    for (var i = 0; i < body.getNumChildren(); i++) {
      var el   = body.getChild(i);
      var type = el.getType();

      if (type === DocumentApp.ElementType.PARAGRAPH) {
        var para = el.asParagraph();
        var text = para.getText().trim();
        if (!text) continue;

        var heading = para.getHeading();
        var tag = 'p';
        if      (heading === DocumentApp.ParagraphHeading.HEADING1) tag = 'h1';
        else if (heading === DocumentApp.ParagraphHeading.HEADING2) tag = 'h2';
        else if (heading === DocumentApp.ParagraphHeading.HEADING3) tag = 'h3';

        var slug = text.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '').slice(0, 40);
        idCount[slug] = (idCount[slug] || 0) + 1;
        var id = idCount[slug] > 1 ? slug + '-' + idCount[slug] : slug;

        items.push({ tag: tag, text: text, id: id, html: paraToHtml(para) });

      } else if (type === DocumentApp.ElementType.LIST_ITEM) {
        var li   = el.asListItem();
        var text = li.getText().trim();
        if (!text) continue;
        items.push({ tag: 'li', text: text, id: '', html: paraToHtml(li) });
      }
    }

    return { items: items };
  } catch (err) {
    return { error: err.toString(), items: [] };
  }
}

function paraToHtml(el) {
  var html = '';
  for (var i = 0; i < el.getNumChildren(); i++) {
    var child = el.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TEXT) continue;
    var t   = child.asText();
    var raw = t.getText();
    if (!raw) continue;

    var j = 0;
    while (j < raw.length) {
      var attrs = t.getAttributes(j);
      var k = j + 1;
      while (k < raw.length) {
        var a2 = t.getAttributes(k);
        if (a2[DocumentApp.Attribute.BOLD]      !== attrs[DocumentApp.Attribute.BOLD]      ||
            a2[DocumentApp.Attribute.ITALIC]    !== attrs[DocumentApp.Attribute.ITALIC]    ||
            a2[DocumentApp.Attribute.UNDERLINE] !== attrs[DocumentApp.Attribute.UNDERLINE] ||
            a2[DocumentApp.Attribute.LINK_URL]  !== attrs[DocumentApp.Attribute.LINK_URL]) break;
        k++;
      }
      var chunk = raw.slice(j, k).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
      if (attrs[DocumentApp.Attribute.LINK_URL])  chunk = '<a href="' + attrs[DocumentApp.Attribute.LINK_URL] + '" target="_blank">' + chunk + '</a>';
      if (attrs[DocumentApp.Attribute.BOLD])       chunk = '<strong>' + chunk + '</strong>';
      if (attrs[DocumentApp.Attribute.ITALIC])     chunk = '<em>' + chunk + '</em>';
      if (attrs[DocumentApp.Attribute.UNDERLINE])  chunk = '<u>' + chunk + '</u>';
      html += chunk;
      j = k;
    }
  }
  return html;
}

// ─────────────────────────────────────────
// ARTIST LIST
// Stored in "Artists" tab of SGG-Global-Data (SPREADSHEET_ID)
// ─────────────────────────────────────────

function serveArtists() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Artists');
    if (!sheet) {
      sheet = ss.insertSheet('Artists');
      sheet.appendRow(['Artist Name']);
      sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
      // Seed with the existing hardcoded list
      const seed = ['Chad Murray','Clayton Schiff','Cosima zu Knyphausen','Darya Diamond',
        'Dustin Hodges','Eli Peng','Emma McMillan','Emmanuel Louisnord Desir',
        'Francis Picabia','Franne Davids','G.V. Rodriguez','Herman Cherry',
        'Jason Nocito','Jiang Chiang','Kate Spencer Stewart','Luis Bermudez',
        'Malcolm Kenter','Melvino Garretti','Nan Montgomery','Nevine Mahmoud',
        'Nick Angelo','Nick Hoecker','Nihura Montiel','Timo Fahler','Tristan Unrau'];
      seed.forEach(a => sheet.appendRow([a]));
    }
    const data    = sheet.getDataRange().getValues();
    const artists = data.slice(1).map(r => String(r[0]||'').trim()).filter(Boolean).sort();
    return ContentService.createTextOutput(JSON.stringify({ artists })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString(), artists: [] })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitArtistsUpdate(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Artists');
    if (!sheet) { sheet = ss.insertSheet('Artists'); }
    sheet.clearContents();
    sheet.appendRow(['Artist Name']);
    sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    const artists = (data.artists || []).map(a => String(a).trim()).filter(Boolean).sort();
    artists.forEach(a => sheet.appendRow([a]));
    logToSheet('ARTISTS UPDATED: ' + artists.length + ' artists');
    return { success: true, count: artists.length };
  } catch(err) {
    logToSheet('ARTISTS ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

function serveSignatures() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Signatures');
    if (!sheet) {
      sheet = ss.insertSheet('Signatures');
      sheet.appendRow(['Name']);
      sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
      ['Dylan Varner','Sebastian Gladstone','Stevie Soares'].forEach(n => sheet.appendRow([n]));
    }
    const data    = sheet.getDataRange().getValues();
    const names   = data.slice(1).map(r => String(r[0]||'').trim()).filter(Boolean).sort();
    return ContentService.createTextOutput(JSON.stringify({ signatures: names })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString(), signatures: [] })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitSignaturesUpdate(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Signatures');
    if (!sheet) { sheet = ss.insertSheet('Signatures'); }
    sheet.clearContents();
    sheet.appendRow(['Name']);
    sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    const names = (data.signatures || []).map(n => String(n).trim()).filter(Boolean).sort();
    names.forEach(n => sheet.appendRow([n]));
    logToSheet('SIGNATURES UPDATED: ' + names.length + ' names');
    return { success: true, count: names.length };
  } catch(err) {
    logToSheet('SIGNATURES ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

function submitICUpdate(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets().find(s => s.getSheetId() === 49023683);
    if (!sheet) return { success: false, error: 'IC sheet not found' };
    sheet.clearContents();
    sheet.appendRow(['Name', 'Type', 'Number']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    const contacts = data.contacts || [];
    contacts.forEach(c => sheet.appendRow([c.name || '', c.type || '', c.number || '']));
    return { success: true, count: contacts.length };
  } catch(err) {
    return { success: false, error: err.toString() };
  }
}

function serveIC() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets().find(s => s.getSheetId() === 49023683);
    if (!sheet) return ContentService.createTextOutput(JSON.stringify({ contacts: [] })).setMimeType(ContentService.MimeType.JSON);
    const data     = sheet.getDataRange().getValues();
    const contacts = data.slice(1).map(r => ({ name: String(r[0]||'').trim(), type: String(r[1]||'').trim(), number: String(r[2]||'').trim() })).filter(c => c.name);
    return ContentService.createTextOutput(JSON.stringify({ contacts })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString(), contacts: [] })).setMimeType(ContentService.MimeType.JSON);
  }
}

function serveContractors() {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Contractors');
    if (!sheet) {
      sheet = ss.insertSheet('Contractors');
      sheet.appendRow(['Contractor Name']);
      sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
      sheet.setFrozenRows(1);
      ['Chasen Wolcott','Jesus Cardenas','Jonathan Austin','Kial Hocker',
       'Matthew Debbaudt','Sheila Alcibar','Stevie Soares','Sydney Busic','Yatta']
        .forEach(n => sheet.appendRow([n]));
    }
    const data        = sheet.getDataRange().getValues();
    const contractors = data.slice(1).map(r => String(r[0]||'').trim()).filter(Boolean).sort();
    return ContentService.createTextOutput(JSON.stringify({ contractors })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ error: err.toString(), contractors: [] })).setMimeType(ContentService.MimeType.JSON);
  }
}

function submitContractorsUpdate(data) {
  try {
    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName('Contractors');
    if (!sheet) { sheet = ss.insertSheet('Contractors'); }
    sheet.clearContents();
    sheet.appendRow(['Contractor Name']);
    sheet.getRange(1,1,1,1).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    const contractors = (data.contractors || []).map(n => String(n).trim()).filter(Boolean).sort();
    contractors.forEach(n => sheet.appendRow([n]));
    logToSheet('CONTRACTORS UPDATED: ' + contractors.length + ' contractors');
    return { success: true, count: contractors.length };
  } catch(err) {
    logToSheet('CONTRACTORS ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

// ─────────────────────────────────────────
// Show Calendar Sync  (restored from v50, adapted for current sheet layout)
// Sheet cols: A=Show, B=Location (LA/NY), C=Dates (MM/DD-MM/DD)
// ─────────────────────────────────────────

function syncGalleryToCalendar() {
  try {
    const cal  = CalendarApp.getCalendarById(EXHIBITIONS_CAL_ID);
    const ss   = SpreadsheetApp.openById(GALLERY_SHEET_ID);
    // Sheet layout: col A = NY show, col B = NYC dates, col E = LA show, col F = LA dates
    const skipPattern = /\btotal\b|sales\s*goal|summer\s*break|gallery\s*closed|arbitrary|consignment\b|bennet\s*sales/i;
    let created = 0;

    const thisYear  = new Date().getFullYear();
    const syncYears = [thisYear - 1, thisYear, thisYear + 1];

    syncYears.forEach(year => {
      // Find the tab whose name contains the year
      const sheets = ss.getSheets();
      const sheet  = sheets.find(s => s.getName().indexOf(String(year)) !== -1);
      if (!sheet) return;

      // Wipe existing Opening/Closing events for this year, then rebuild
      const yearStart = new Date(year, 0, 1);
      const yearEnd   = new Date(year, 11, 31, 23, 59);
      cal.getEvents(yearStart, yearEnd).forEach(ev => {
        if (/,\s*(Opening|Closing)\s*\((NYC|LA)\)/i.test(ev.getTitle())) ev.deleteEvent();
      });

      const data = sheet.getDataRange().getValues();

      for (let i = 1; i < data.length; i++) {
        // ── NYC: col A (index 0) = show name, col B (index 1) = dates ──
        const nyRaw = String(data[i][0] || '').trim();
        if (nyRaw && !skipPattern.test(nyRaw)) {
          const nyName = nyRaw.replace(/^\d+\.\s*/, '').replace(/\s*\(consigned\)\s*/gi, '').replace(/\s+BOOK\s*$/i, '').trim();
          if (nyName) {
            const dates = parseGalleryDateRange(data[i][1], year);
            if (dates) {
              cal.createAllDayEvent(nyName + ', Opening (NYC)', dates.start);
              cal.createAllDayEvent(nyName + ', Closing (NYC)', dates.end);
              created += 2;
            }
          }
        }

        // ── LA: col E (index 4) = show name, col F (index 5) = dates ──
        const laRaw = String(data[i][4] || '').trim();
        if (laRaw && !skipPattern.test(laRaw)) {
          const laName = laRaw.replace(/^\d+\.\s*/, '').replace(/\s*\(consigned\)\s*/gi, '').replace(/\s+BOOK\s*$/i, '').trim();
          if (laName) {
            const dates = parseGalleryDateRange(data[i][5], year);
            if (dates) {
              cal.createAllDayEvent(laName + ', Opening (LA)', dates.start);
              cal.createAllDayEvent(laName + ', Closing (LA)', dates.end);
              created += 2;
            }
          }
        }
      }
    });

    logToSheet('CALENDAR SYNC OK: ' + created + ' events created');
    Logger.log('✅ Synced ' + created + ' events to Google Calendar');
  } catch(err) {
    logToSheet('CALENDAR SYNC ERR: ' + err.toString());
    Logger.log('❌ Sync error: ' + err.toString());
  }
}

function parseGalleryDateRange(val, year) {
  if (!val) return null;
  if (val instanceof Date) return { start: val, end: val };

  // Strip time info like "6-8PM"
  let str = String(val).trim().replace(/\s+\d{1,2}(?::\d{2})?\s*(?:[-–]\s*\d{1,2}(?::\d{2})?)?\s*[AaPp][Mm]/g, '').trim();
  if (!str) return null;

  // MM/DD-MM/DD or MM/DD – MM/DD
  const rangeMatch = str.match(/^(\d{1,2})\/(\d{1,2})\s*[-–]\s*(\d{1,2})\/(\d{1,2})/);
  if (rangeMatch) {
    const sM = parseInt(rangeMatch[1]), sD = parseInt(rangeMatch[2]);
    const eM = parseInt(rangeMatch[3]), eD = parseInt(rangeMatch[4]);
    return {
      start: new Date(year,             sM - 1, sD),
      end:   new Date(eM < sM ? year+1 : year, eM - 1, eD)
    };
  }

  // Single date MM/DD
  const single = str.match(/^(\d{1,2})\/(\d{1,2})/);
  if (single) {
    const d = new Date(year, parseInt(single[1]) - 1, parseInt(single[2]));
    return { start: d, end: d };
  }

  return null;
}

// Run once to install the nightly trigger
function createGallerySyncTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncGalleryToCalendar') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncGalleryToCalendar')
    .timeBased().atHour(0).nearMinute(0).everyDays(1)
    .inTimezone('America/New_York').create();
  Logger.log('✅ Gallery sync trigger created: runs nightly at midnight ET');
}

// ─────────────────────────────────────────
// BILL OF LADING — parse (AI) + save
// ─────────────────────────────────────────
const BOL_SHEET = 'BOL Log';

function parseBOL(data) {
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
    if (!apiKey) throw new Error('ANTHROPIC_API_KEY not set in Script Properties');

    const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      payload: JSON.stringify({
        model: 'claude-opus-4-8',
        max_tokens: 1024,
        messages: [{
          role: 'user',
          content: [
            {
              type: 'document',
              source: { type: 'base64', media_type: data.mimeType || 'application/pdf', data: data.fileData }
            },
            {
              type: 'text',
              text: 'Extract the following fields from this Bill of Lading and return ONLY valid JSON with these keys: bol_number, date (ISO 8601 YYYY-MM-DD if present), carrier, origin, destination, pieces (integer), artist, artwork, notes. Use null for missing fields.'
            }
          ]
        }]
      }),
      muteHttpExceptions: true
    });

    const body = JSON.parse(resp.getContentText());
    if (resp.getResponseCode() !== 200) throw new Error(body.error && body.error.message || 'Claude API error');

    const raw = body.content[0].text.trim();
    const jsonMatch = raw.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Could not parse AI response');
    const fields = JSON.parse(jsonMatch[0]);
    return { success: true, fields: fields };

  } catch (err) {
    logToSheet('BOL PARSE ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}

function saveBOL(data) {
  try {
    const token     = ScriptApp.getOAuthToken();
    const bolRootId = findOrCreateFolder(token, ROOT_FOLDER_ID, 'BOL');
    const safeArtist = (data.artist || 'Unknown').replace(/[^a-zA-Z0-9 _-]/g, '_');
    const fileName  = (data.bol_number ? data.bol_number + '_' : '') + safeArtist + '_BOL.pdf';
    const upload    = uploadFile(token, bolRootId, fileName, data.mimeType || 'application/pdf', data.fileData);
    const fileUrl   = 'https://drive.google.com/file/d/' + upload.id + '/view';
    setPublicRead(token, upload.id);

    const sheet = getGlobalDataSheet(BOL_SHEET, ['BOL #','Date','Artist','Artwork','Carrier','Origin','Destination','Pieces','Notes','File']);
    sheet.appendRow([
      data.bol_number  || '',
      data.date        || '',
      data.artist      || '',
      data.artwork     || '',
      data.carrier     || '',
      data.origin      || '',
      data.destination || '',
      data.pieces      || '',
      data.notes       || '',
      ''
    ]);
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 10).setFormula(
      '=HYPERLINK("' + fileUrl.replace(/"/g,'""') + '","' + fileName.replace(/"/g,'""') + '")'
    );

    logToSheet('BOL SAVE OK: ' + fileName);
    return { success: true, fileUrl: fileUrl, fileName: fileName };

  } catch (err) {
    logToSheet('BOL SAVE ERR: ' + err.toString());
    return { success: false, error: err.toString() };
  }
}
