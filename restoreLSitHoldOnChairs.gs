/**
 * Restores "L-Sit Hold on Chairs" to Husband Core P2C1, Friday, Block 2,
 * undoing a swap that should not have been made there. Read the live
 * Sheet directly first to confirm current state before writing this:
 * that row currently reads "Bear Hold" (3, 20 sec), and the Friday
 * Finisher "Strength Core Gauntlet" currently references "Bear Hold
 * 20s" in its combo text.
 *
 * This script makes exactly two changes, both on tab "Husband Core
 * P2C1", both on Friday:
 *
 *   1. Block 2: "Bear Hold" -> "L-Sit Hold on Chairs" (3, 25 sec, with
 *      its original coaching note). Bear Hold is not moved anywhere
 *      else or kept anywhere in this tab -- restoring L-Sit Hold on
 *      Chairs to this row simply overwrites it, matching the tab's
 *      original content exactly.
 *   2. Finisher: "Strength Core Gauntlet" note text corrected back
 *      from "Bear Hold 20s" to "L-Sit 20s" so the combo description
 *      matches the restored exercise.
 *
 * "L-Sit Hold on Chairs" was checked and does not appear anywhere else
 * in the Husband Core P2C1 tab (Monday or Wednesday), so there is
 * nothing else to restore in this tab.
 *
 * The local app code (workout.html) was also checked directly and was
 * never modified for this exercise or this Finisher note in the first
 * place -- both already read "L-Sit Hold on Chairs" / "L-Sit 20s" --
 * so no local file changes are needed or made by this script.
 *
 * The Exercise Demo Link column is never touched. No other row, tab,
 * or file is touched by this script.
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD
 *   change (View > Logs) and edits nothing. Once the log confirms both
 *   rows were found, set DRY_RUN to false and run
 *   restoreLSitHoldOnChairs() again to actually apply it.
 *
 * Run once by selecting restoreLSitHoldOnChairs from the function
 * dropdown and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

var TAB = 'Husband Core P2C1';

var EXERCISE_RESTORE = {
  day: 'Friday', block: 'Block 2',
  oldName: 'Bear Hold', newName: 'L-Sit Hold on Chairs',
  sets: '3', target: '25 sec',
  note: 'Press down hard through your palms with arms fully locked straight and lift your hips off the chairs. Breathe steadily and never hold your breath.'
};

var NOTE_RESTORE = {
  day: 'Friday', block: 'Finisher', exercise: 'Strength Core Gauntlet',
  newNote: 'Copenhagen 20s/side -> Dragon Flag 5 -> L-Sit 20s, zero rest. Move immediately between exercises and maintain quality form throughout.'
};

function restoreLSitHoldOnChairs() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(TAB);
  if (!sheet) {
    Logger.log('NOT FOUND: no tab named "' + TAB + '" exists in this spreadsheet.');
    return;
  }

  var headerInfo = findHeaderRow(sheet);
  if (!headerInfo) {
    Logger.log('NOT FOUND: could not locate a header row with "Day" in column A on tab "' + TAB + '".');
    return;
  }

  var headerRow = headerInfo.headers;
  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var setsCol = headerRow.indexOf('Sets');
  var targetCol = headerRow.indexOf('Target Reps / Time');
  var noteCol = headerRow.indexOf('Coaching Note');

  if (dayCol === -1 || blockCol === -1 || exerciseCol === -1 || noteCol === -1) {
    Logger.log('NOT FOUND: required columns missing. Actual header: ' + JSON.stringify(headerRow));
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var firstDataRow = headerInfo.rowIndex + 1;
  var data = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  var exerciseDone = false;
  var noteDone = false;

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();
    var sheetRow = firstDataRow + i;

    if (!exerciseDone && day === EXERCISE_RESTORE.day && block === EXERCISE_RESTORE.block) {
      if (exercise === EXERCISE_RESTORE.newName) {
        Logger.log('EXERCISE: ALREADY DONE (row ' + sheetRow + ' already reads "' + EXERCISE_RESTORE.newName + '") -- skipped.');
        exerciseDone = true;
      } else if (exercise === EXERCISE_RESTORE.oldName) {
        Logger.log('EXERCISE: ' + (DRY_RUN ? 'WOULD RESTORE' : 'RESTORED') + ' row ' + sheetRow + ': "' + EXERCISE_RESTORE.oldName + '" -> "' + EXERCISE_RESTORE.newName + '"');
        if (!DRY_RUN) {
          sheet.getRange(sheetRow, exerciseCol + 1).setValue(EXERCISE_RESTORE.newName);
          if (setsCol !== -1) sheet.getRange(sheetRow, setsCol + 1).setValue(EXERCISE_RESTORE.sets);
          if (targetCol !== -1) sheet.getRange(sheetRow, targetCol + 1).setValue(EXERCISE_RESTORE.target);
          sheet.getRange(sheetRow, noteCol + 1).setValue(EXERCISE_RESTORE.note);
        }
        exerciseDone = true;
      } else {
        Logger.log('EXERCISE: NOT FOUND -- row ' + sheetRow + ' at ' + EXERCISE_RESTORE.day + ' / ' + EXERCISE_RESTORE.block + ' reads "' + exercise + '", expected "' + EXERCISE_RESTORE.oldName + '" or "' + EXERCISE_RESTORE.newName + '".');
      }
    }

    if (!noteDone && day === NOTE_RESTORE.day && block === NOTE_RESTORE.block && exercise === NOTE_RESTORE.exercise) {
      var currentNote = String(row[noteCol]).trim();
      if (currentNote === NOTE_RESTORE.newNote) {
        Logger.log('NOTE: ALREADY DONE (row ' + sheetRow + ' note already matches) -- skipped.');
      } else {
        Logger.log('NOTE: ' + (DRY_RUN ? 'WOULD RESTORE' : 'RESTORED') + ' row ' + sheetRow + ' note for "' + NOTE_RESTORE.exercise + '"');
        if (!DRY_RUN) {
          sheet.getRange(sheetRow, noteCol + 1).setValue(NOTE_RESTORE.newNote);
        }
      }
      noteDone = true;
    }
  }

  if (!exerciseDone) {
    Logger.log('EXERCISE: NOT FOUND -- no row matched ' + EXERCISE_RESTORE.day + ' / ' + EXERCISE_RESTORE.block + ' at all.');
  }
  if (!noteDone) {
    Logger.log('NOTE: NOT FOUND -- no row matched ' + NOTE_RESTORE.day + ' / ' + NOTE_RESTORE.block + ' / "' + NOTE_RESTORE.exercise + '".');
  }

  Logger.log(DRY_RUN ? '[DRY RUN -- nothing changed]' : '[APPLIED]');
}

function findHeaderRow(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var scanRows = Math.min(lastRow, 5);
  if (scanRows < 1) return null;
  var values = sheet.getRange(1, 1, scanRows, lastCol).getValues();
  for (var r = 0; r < values.length; r++) {
    if (String(values[r][0]).trim() === 'Day') {
      return { rowIndex: r + 1, headers: values[r].map(function (h) { return String(h).trim(); }) };
    }
  }
  return null;
}
