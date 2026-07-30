const SPREADSHEET_ID = '1BBjshS0EIbnXDxUHbtuyRvbksn5FK7nMFDLO46IY-Y0';
const SHEET_NAME = 'Sheet1';

// Documents are uploaded to this Drive folder:
// https://drive.google.com/drive/folders/1zteHC5UFYn3tOQmbdYuMN1tcPxWibvBS
// Leave DRIVE_FOLDER_ID empty to auto-create a folder next to the spreadsheet instead.
const DRIVE_FOLDER_ID = '1zteHC5UFYn3tOQmbdYuMN1tcPxWibvBS';
const DRIVE_FOLDER_NAME = 'Osool - Team';

const TIMEZONE = 'Asia/Riyadh';

const HEADERS = [
  'submitted_at',
  'full_name',
  'department',
  'position',
  'company',
  'mobile',
  'instagram_account',
  'nationality',
  'document_type',
  'document_number',
  'document_file',
  'has_car',
  'plate_number',
  'car_type',
  'days',
  'department_head',
  'is_head',
  'team_count',
  'notes',
  'consent_accuracy',
  'consent_confidentiality',
  'consent_no_photography',
  'source'
];

function doPost(e) {
  try {
    const data = JSON.parse((e && e.postData && e.postData.contents) || '{}');

    // Upload the attached ID/passport to Drive and store its link in the sheet.
    if (data.document_file_data) {
      data.document_file = saveDocument(data);
    }

    const sheet = getSheet();
    ensureHeaders(sheet, HEADERS);

    // Format the target cells as plain text BEFORE writing, otherwise Sheets
    // reads values like "0532414582" as numbers and drops the leading zero.
    const row = sheet.getLastRow() + 1;
    applyTextFormats(sheet, row);
    sheet.getRange(row, 1, 1, HEADERS.length)
         .setValues([HEADERS.map(header => data[header] || '')]);

    return json({ok: true, file: data.document_file || ''});
  } catch (error) {
    // Surface the failure so the form can show a real error instead of a false success.
    return json({ok: false, error: String((error && error.message) || error)});
  }
}

// Columns that must stay text: phone numbers, ID numbers and plate numbers all
// lose meaning if Sheets treats them as numbers.
const TEXT_COLUMNS = ['mobile', 'document_number', 'plate_number', 'instagram_account'];

function applyTextFormats(sheet, row) {
  TEXT_COLUMNS.forEach(function (name) {
    const index = HEADERS.indexOf(name);
    if (index > -1) sheet.getRange(row, index + 1).setNumberFormat('@');
  });
}

function doGet() {
  return json({ok: true, service: 'Captains crew form', time: new Date().toISOString()});
}

function saveDocument(data) {
  const bytes = Utilities.base64Decode(data.document_file_data);
  const mime = data.document_file_type || 'application/octet-stream';
  const stamp = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd_HH-mm-ss');
  const name = [
    sanitize(data.full_name) || 'crew',
    sanitize(data.document_number),
    stamp
  ].filter(String).join('_');

  const blob = Utilities.newBlob(bytes, mime, name + extensionFor(data.document_file_name, mime));
  return getFolder().createFile(blob).getUrl();
}

function getFolder() {
  if (DRIVE_FOLDER_ID) return DriveApp.getFolderById(DRIVE_FOLDER_ID);

  const parent = spreadsheetParent();
  const existing = parent.getFoldersByName(DRIVE_FOLDER_NAME);
  return existing.hasNext() ? existing.next() : parent.createFolder(DRIVE_FOLDER_NAME);
}

function spreadsheetParent() {
  const parents = DriveApp.getFileById(SPREADSHEET_ID).getParents();
  return parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
}

function extensionFor(fileName, mime) {
  const match = String(fileName || '').match(/\.[a-z0-9]+$/i);
  if (match) return match[0].toLowerCase();
  if (mime.indexOf('pdf') > -1) return '.pdf';
  if (mime.indexOf('png') > -1) return '.png';
  if (mime.indexOf('jpeg') > -1 || mime.indexOf('jpg') > -1) return '.jpg';
  return '';
}

function sanitize(value) {
  return String(value || '').replace(/[\\/:*?"<>|]/g, '-').trim();
}

function json(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeaders(sheet, headers) {
  const firstRow = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  const hasHeaders = firstRow.some(value => value);
  if (!hasHeaders) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}
