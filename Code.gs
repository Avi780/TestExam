const SPREADSHEET_ID = '1Xg_QsBX2y-Gd_lmlfrSJbUCiL3EV_yb6';

function createPreviousYearSourceSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = 'RRB PYQ Source Index';
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
    sheet.appendRow(['Recruitment year', 'CEN / notification', 'Post scope', 'Exam date / shift', 'Source type', 'Source URL', 'Official cross-check', 'Notes']);
    sheet.setFrozenRows(1);
  }

  const rows = [
    ['2018', 'CEN 01/2018', 'Grade I Signal and Grade III listings', 'CBT 1, multiple dates/shifts', 'Secondary paper listing', 'https://prepp.in/rrb-technician-exam/question%20paper%202018', 'Not independently verified against an archived RRB response sheet', 'The listing covers multiple posts; verify that a specific paper is Grade I Signal before treating it as such.'],
    ['2019', 'CEN 01/2018', 'Trade-specific Technician CBT 2 listings', 'Jan 2019, multiple dates/shifts', 'Secondary paper listing', 'https://prepp.in/rrb-technician-exam/question%20paper%202018', 'Not independently verified against an archived RRB response sheet', 'These are trade papers, not Grade I Signal; excluded from the Grade I Signal practice set.'],
    ['2024', 'CEN 02/2024', 'Technician Grade I Signal', '19 Dec 2024, Shift 1', 'Prepp-hosted response-paper PDF', 'https://cdn-images.prepp.in/public/image/RRB_Technician_Grade_I_Dec_19_2024_Shift_1_English_ae564f2eee28ce51b5563a4f30cc5ca9.pdf', 'Secondary host; not verified against an archived RRB response sheet', 'Full paper is available as a direct PDF; the interactive preview exposes only the first 15 questions without login.'],
    ['2024', 'CEN 02/2024', 'Technician Grade I Signal', '19 Dec 2024, Shift 2', 'Prepp-hosted response-paper PDF', 'https://cdn-images.prepp.in/public/image/RRB_Technician_Grade_I_Dec_19_Shift_2_English_6d406dd5c2af785d2f88b6fdf0f76260.pdf', 'Secondary host; not verified against an archived RRB response sheet', 'Direct PDF source identified through the paper listing.'],
    ['2024', 'CEN 02/2024', 'Technician Grade I Signal', '19 Dec 2024, Shift 3', 'Prepp-hosted response-paper PDF', 'https://cdn-images.prepp.in/public/image/RRB_Technician_Grade_I_Dec_19_2024_Shift_3_English_ee1d99d21bca5c2c526da5459b46a195.pdf', 'Secondary host; not verified against an archived RRB response sheet', 'Direct PDF source identified through the paper listing.'],
    ['2024', 'CEN 02/2024', 'Technician Grade I Signal', '20 Dec 2024, Shift 1', 'Prepp-hosted response-paper PDF', 'https://cdn-images.prepp.in/public/image/RRB_Technician_Grade_I_Dec_20_2024_Shift_1_English_2ae41139e5af2846698a98686a2892fd.pdf', 'Secondary host; not verified against an archived RRB response sheet', 'Direct PDF source identified through the paper listing.'],
    ['2025', 'CEN 02/2025', 'Technician Grade I / Grade III', 'Listings show exam dates in March 2026', 'Secondary paper listing', 'https://www.oliveboard.in/blog/rrb-technician-previous-year-papers/', 'Not an exam conducted by the end of 2025', 'Excluded from the through-2025 paper set because the listed exam took place in 2026.']
  ];

  const existingUrls = sheet.getLastRow() > 1 ? sheet.getRange(2, 6, sheet.getLastRow() - 1, 1).getDisplayValues().flat() : [];
  const newRows = rows.filter(row => !existingUrls.includes(row[5]));
  if (newRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, rows[0].length).setValues(newRows);
  if (!sheet.getFilter()) sheet.getRange(1, 1, sheet.getLastRow(), rows[0].length).createFilter();
  sheet.autoResizeColumns(1, rows[0].length);
  sheet.setColumnWidth(6, 320);
  sheet.setColumnWidth(7, 360);
  sheet.setColumnWidth(8, 520);
  return newRows.length ? 'Source index updated' : 'Source index is already up to date';
}

function createLicensedQuestionSheet() {
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheetName = 'Licensed PYQ Questions';
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (sheet) return 'Licensed question sheet already exists';

  sheet = spreadsheet.insertSheet(sheetName);
  const headers = ['Exam year', 'CEN / notification', 'Post / grade', 'Exam date', 'Shift', 'Subject', 'Question', 'Option A', 'Option B', 'Option C', 'Option D', 'Correct answer', 'Hint', 'Source URL', 'Verification status'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).createFilter();
  sheet.getRange(2, 12, 999, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D'], true).setAllowInvalid(false).build()
  );
  sheet.setColumnWidth(7, 420);
  sheet.setColumnWidths(8, 4, 180);
  sheet.setColumnWidth(13, 300);
  sheet.setColumnWidth(14, 360);
  sheet.setColumnWidth(15, 220);
  return 'Licensed question sheet created';
}

function doPost(e) {
  const result = JSON.parse(e.postData.contents);
  const name = String(result.name || '').trim().slice(0, 80);
  const score = Number(result.score);
  const correct = Number(result.correct);
  const wrong = Number(result.wrong);
  const unattempted = Number(result.unattempted);

  if (!name || !Number.isFinite(score) || !Number.isInteger(correct) || !Number.isInteger(wrong) || !Number.isInteger(unattempted)) {
    throw new Error('Invalid test result');
  }

  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName('Test Results');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('Test Results');
    sheet.appendRow(['Submitted At', 'Name', 'Score', 'Correct', 'Wrong', 'Unattempted', 'Test ID']);
  }
  const testId = String(result.testId || '');
  const row = [new Date(), name, score, correct, wrong, unattempted, testId];
  const savedIds = sheet.getLastRow() > 1 ? sheet.getRange(2, 7, sheet.getLastRow() - 1, 1).getDisplayValues().flat() : [];
  const existingIndex = savedIds.indexOf(testId);
  if (existingIndex >= 0) sheet.getRange(existingIndex + 2, 1, 1, row.length).setValues([row]);
  else sheet.appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ok: true})).setMimeType(ContentService.MimeType.JSON);
}