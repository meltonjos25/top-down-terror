/**
 * READ-ONLY diagnostic. Makes no changes to the spreadsheet at all --
 * it only logs what is actually there, straight from SpreadsheetApp
 * (the same API a fix script uses), so nothing is assumed or guessed.
 *
 * This exists because two prior fix scripts both reported "0 rows
 * found" against tab names and exercise names that were built from a
 * Drive-based text export of the sheet, not from SpreadsheetApp itself.
 * That export can silently normalize things (whitespace, punctuation,
 * even mis-render certain characters), so this script instead:
 *
 *   1. Lists EVERY actual tab name in the spreadsheet, exactly as
 *      SpreadsheetApp sees it, wrapped in quotes so leading/trailing
 *      whitespace is visible.
 *   2. Checks each of the 8 expected Core tab names with
 *      getSheetByName() and reports FOUND or NOT FOUND for each --
 *      this alone will show if a tab name assumption was simply wrong.
 *   3. For every tab that IS found, dumps every non-blank row (Day,
 *      Block, Exercise, Sets, Target, Coaching Note) to the log.
 *   4. Flags any hidden/invisible/non-ASCII character found inside a
 *      Day, Block, or Exercise cell -- printing its position and
 *      Unicode code point -- since an invisible character is another
 *      way two "identical-looking" strings can fail to match.
 *
 * Run once by selecting diagnoseCoreWorkouts from the function dropdown
 * and clicking Run, then read the output at View > Logs (or open
 * Executions, click this run, and view its log). Paste that log back
 * so the exact names can be used to write a script that will actually
 * match.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';

var TARGET_TABS = [
  'Husband Core P1C1',
  'Husband Core P2C1',
  'Husband Core P3C1',
  'Husband Core P4C1',
  'Wife Core P1C1',
  'Wife Core P2C1',
  'Wife Core P3C1',
  'Wife Core P4C1'
];

function diagnoseCoreWorkouts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var allSheets = ss.getSheets();

  Logger.log('===== ALL TAB NAMES IN THIS SPREADSHEET (' + allSheets.length + ' total) =====');
  for (var i = 0; i < allSheets.length; i++) {
    var rawName = allSheets[i].getName();
    Logger.log((i + 1) + '. "' + rawName + '"' + describeHiddenChars(rawName));
  }

  Logger.log('');
  Logger.log('===== TARGET TAB LOOKUP (getSheetByName) =====');
  for (var t = 0; t < TARGET_TABS.length; t++) {
    var name = TARGET_TABS[t];
    var sheet = ss.getSheetByName(name);
    Logger.log('"' + name + '": ' + (sheet ? 'FOUND' : 'NOT FOUND'));
  }

  Logger.log('');
  Logger.log('===== FULL ROW DUMP FOR EACH FOUND TAB =====');
  for (var t2 = 0; t2 < TARGET_TABS.length; t2++) {
    var name2 = TARGET_TABS[t2];
    var sheet2 = ss.getSheetByName(name2);
    Logger.log('----- ' + name2 + ' -----');
    if (!sheet2) {
      Logger.log('  (tab does not exist -- skipped)');
      continue;
    }
    dumpSheet(sheet2);
  }

  Logger.log('');
  Logger.log('===== DONE. No changes were made to the spreadsheet. =====');
}

function dumpSheet(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    Logger.log('  (empty sheet)');
    return;
  }

  var data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRow = data[0].map(function (h) { return String(h); });
  Logger.log('  HEADERS (' + headerRow.length + '): ' + JSON.stringify(headerRow));

  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var setsCol = headerRow.indexOf('Sets');
  var targetCol = headerRow.indexOf('Target Reps / Time');
  var noteCol = headerRow.indexOf('Coaching Note');

  if (dayCol === -1 || blockCol === -1 || exerciseCol === -1) {
    Logger.log('  Day/Block/Exercise not found by exact header text -- dumping raw rows instead.');
    for (var r = 1; r < data.length; r++) {
      Logger.log('  [row ' + (r + 1) + '] RAW: ' + JSON.stringify(data[r]));
    }
    return;
  }

  for (var r2 = 1; r2 < data.length; r2++) {
    var row = data[r2];
    var day = String(row[dayCol] || '').trim();
    var block = String(row[blockCol] || '').trim();
    var exercise = String(row[exerciseCol] || '').trim();
    var sets = setsCol !== -1 ? String(row[setsCol] || '').trim() : '';
    var target = targetCol !== -1 ? String(row[targetCol] || '').trim() : '';
    var note = noteCol !== -1 ? String(row[noteCol] || '').trim() : '';
    if (!day && !exercise) continue;

    var flags = describeHiddenChars(day) + describeHiddenChars(block) + describeHiddenChars(exercise);
    Logger.log('  [row ' + (r2 + 1) + '] ' + day + ' | ' + block + ' | "' + exercise + '" | ' + sets + ' | ' + target + ' | note: ' + note + flags);
  }
}

function describeHiddenChars(str) {
  var flags = [];
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    if (code > 126 || (code < 32 && code !== 9)) {
      flags.push('pos ' + i + ':U+' + ('0000' + code.toString(16).toUpperCase()).slice(-4));
    }
  }
  return flags.length ? ' [HIDDEN/NON-ASCII: ' + flags.join(', ') + ']' : '';
}
