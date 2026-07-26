/**
 * Applies the Core-workout row-exercise fix to the live Google Sheet.
 *
 * IMPORTANT -- read this before running:
 * The original 8-row plan was built against the app's local
 * CONJUGATE_EXERCISES fallback data, which turned out to have diverged
 * significantly from what is actually on this Sheet. After reading the
 * live tab content directly, only ONE of the original 8 target rows
 * actually exists as named:
 *
 *   Wife Core P2C1, Monday, Block 2: "Seated Band Core Row Tan"
 *
 * The other 7 do not need any change on this Sheet:
 *   - Husband Core P1C1 (Monday Block 2, Friday Block 1), Husband Core
 *     P3C1 (Wednesday Block 1), and Wife Core P1C1 (Monday Block 2) never
 *     contained the named rowing exercise in the first place -- those
 *     slots have different exercises entirely and no rowing movement.
 *   - Husband Core P4C1 and Wife Core P4C1 do not exist yet (Phase 4
 *     has not started, so those tabs have not been created).
 *
 * "Side Plank with Band Row Blue" (Husband Core P2C1, Wednesday, Block 2)
 * is the only OTHER rowing-pattern exercise present in Core right now,
 * and it is intentionally left alone per prior confirmation -- it is a
 * side-plank anti-rotation hold with an incidental band row, not a back
 * exercise, and is NOT in the CHANGES list below.
 *
 * Also fixed from the prior version of this script: the real header row
 * on this Sheet is "Target Reps / Time" (spaces around the slash) and
 * "Exercise Demo Link" -- not "Target Reps/Time" / "Demo Link". That
 * mismatch alone caused every tab to fail validation and report "not
 * found" last time, regardless of exercise names.
 *
 * For the one real match, this script finds the row by matching Day +
 * Block + the CURRENT Exercise name, then updates that row's Exercise,
 * Sets, Target Reps / Time, and Coaching Note cells. The Exercise Demo
 * Link column is never touched. No tab is reordered, renamed, added, or
 * deleted, and no row other than this one is ever written to.
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD change
 *   (View > Logs, or Execution log) and edits nothing. Once the log shows
 *   the row was found, set DRY_RUN to false and run updateCoreWorkouts()
 *   again to actually apply it.
 *
 * Run once by selecting updateCoreWorkouts from the function dropdown
 * and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

var EXPECTED_HEADERS = ['Day', 'Block', 'Exercise', 'Sets', 'Target Reps / Time', 'Exercise Demo Link', 'Coaching Note'];

var CHANGES = [
  {
    tab: 'Wife Core P2C1', day: 'Monday', block: 'Block 2',
    oldName: 'Seated Band Core Row Tan', newName: 'Band Pallof Press Tan',
    sets: '3', target: '10/side',
    note: 'Anchor the tan band at chest height and sit or stand sideways to it. Press both hands straight out and hold 2 seconds -- resist any rotation. Breathe steadily and never hold your breath.'
  }
];

function updateCoreWorkouts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var results = [];

  for (var i = 0; i < CHANGES.length; i++) {
    results.push(applyChange(ss, CHANGES[i]));
  }

  var okCount = 0;
  var missingCount = 0;
  var lines = [];
  for (var j = 0; j < results.length; j++) {
    lines.push(results[j].message);
    if (results[j].status === 'OK') okCount++;
    else missingCount++;
  }

  Logger.log((DRY_RUN ? '[DRY RUN -- nothing changed]\n' : '[APPLIED]\n') + lines.join('\n'));
  Logger.log(okCount + ' of ' + CHANGES.length + ' rows matched' + (DRY_RUN ? ' (would be updated).' : ' and updated.'));
  if (missingCount > 0) {
    Logger.log('WARNING: ' + missingCount + ' row(s) could not be found -- check tab name, Day, Block, and current Exercise name for those rows above.');
  }
}

function applyChange(ss, change) {
  var label = change.tab + ' / ' + change.day + ' / ' + change.block + ' / "' + change.oldName + '"';

  var sheet = ss.getSheetByName(change.tab);
  if (!sheet) {
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (no tab named "' + change.tab + '")' };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < EXPECTED_HEADERS.length) {
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (tab has no data rows)' };
  }

  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  if (!headersMatch(headerRow, EXPECTED_HEADERS)) {
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (tab headers do not match the expected layout -- actual headers were: ' + headerRow.join(' | ') + ')' };
  }

  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var setsCol = headerRow.indexOf('Sets');
  var targetCol = headerRow.indexOf('Target Reps / Time');
  var noteCol = headerRow.indexOf('Coaching Note');

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();

    if (day === change.day && block === change.block && exercise === change.oldName) {
      var sheetRow = i + 2;
      if (!DRY_RUN) {
        sheet.getRange(sheetRow, exerciseCol + 1).setValue(change.newName);
        sheet.getRange(sheetRow, setsCol + 1).setValue(change.sets);
        sheet.getRange(sheetRow, targetCol + 1).setValue(change.target);
        sheet.getRange(sheetRow, noteCol + 1).setValue(change.note);
      }
      return { status: 'OK', message: change.tab + ' / ' + change.day + ' / row ' + sheetRow + ': OK -- "' + change.oldName + '" -> "' + change.newName + '"' };
    }
  }

  return { status: 'NOT FOUND', message: label + ': NOT FOUND (no row matched that Day + Block + Exercise name -- it may already be updated, or the sheet text differs slightly)' };
}

function headersMatch(actual, expected) {
  if (actual.length < expected.length) return false;
  for (var i = 0; i < expected.length; i++) {
    if (actual[i] !== expected[i]) return false;
  }
  return true;
}
