const SPREADSHEET_ID = '1BBjshS0EIbnXDxUHbtuyRvbksn5FK7nMFDLO46IY-Y0';
const SHEET_NAME = 'Sheet1';

function doPost(e) {
  const sheet = getSheet();
  const data = JSON.parse(e.postData.contents || '{}');
  const headers = [
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

  ensureHeaders(sheet, headers);
  sheet.appendRow(headers.map(header => data[header] || ''));

  return ContentService
    .createTextOutput(JSON.stringify({ok: true}))
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
