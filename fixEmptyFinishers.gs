/**
 * Fills in the coaching notes for two Husband Arms P3C1 Finisher rows
 * that currently have generic pep-talk text with zero specific
 * exercises named (confirmed by reading the live Sheet directly).
 *
 * Only two rows are touched, both on tab "Husband Arms P3C1", both
 * note-only (the Exercise name, Sets, and Target Reps / Time are left
 * exactly as they are -- only the Coaching Note cell changes):
 *
 *   1. Thursday, Finisher, "Power Shoulder Complex": replaces the
 *      generic note with a named, ordered, no-rest circuit of the
 *      day's three shoulder press movements (Fast Pike Push-Up,
 *      Explosive Elevated Pike Push-Up, Jump and Press Band Blue).
 *
 *   2. Saturday, Finisher, "Power Triset": replaces the generic note
 *      with a named, ordered, no-rest circuit of three full-arm
 *      explosive movements already present that day (Clap Push-Up,
 *      Plyometric Dip, Band Curl Fast Blue).
 *
 * A third fix -- replacing "Max Single Arm Push-Up" with "Max Single
 * Arm Plank Hold" on Husband Core P4C1 Monday's Finisher -- is NOT in
 * this script because that tab does not exist on the Sheet yet (Phase
 * 4 has not started). That fix was applied to the local app's fallback
 * data only; apply it to the Sheet once that tab is created.
 *
 * The Exercise Demo Link column is never touched. No other row, tab,
 * or file is touched by this script.
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD
 *   change (View > Logs) and edits nothing. Once the log confirms both
 *   rows were found, set DRY_RUN to false and run fixEmptyFinishers()
 *   again to actually apply it.
 *
 * Run once by selecting fixEmptyFinishers from the function dropdown
 * and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

var TAB = 'Husband Arms P3C1';

var NOTE_UPDATES = [
  {
    day: 'Thursday', block: 'Finisher', exercise: 'Power Shoulder Complex',
    newNote: 'Move immediately from one exercise to the next with zero rest between movements. Fast Pike Push-Up: hips high, lower your head toward the floor and press back up as fast as possible for 10 reps. Explosive Elevated Pike Push-Up: feet on a chair, hips high, drive up hard enough that your hands briefly leave the floor for 6 reps. Jump and Press Band Blue: stand on the blue band, squat slightly, then jump and press both hands overhead in one explosive motion for 10 reps. Breathe out on every exertion and never hold your breath.'
  },
  {
    day: 'Saturday', block: 'Finisher', exercise: 'Power Triset',
    newNote: 'Move from exercise to exercise with zero rest between movements. Clap Push-Up: press up explosively hard enough to clap your hands together before landing, for 8 reps. Plyometric Dip: push down through your palms so hard your arms fully extend and your body rises off the surface, for 8 reps. Band Curl Fast Blue: stand on the blue band and curl both hands to your shoulders as fast as possible for 15 reps. Breathe out on every rep and never hold your breath.'
  }
];

function fixEmptyFinishers() {
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
  var noteCol = headerRow.indexOf('Coaching Note');

  if (dayCol === -1 || blockCol === -1 || exerciseCol === -1 || noteCol === -1) {
    Logger.log('NOT FOUND: required columns missing. Actual header: ' + JSON.stringify(headerRow));
    return;
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var firstDataRow = headerInfo.rowIndex + 1;
  var data = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  var doneFlags = NOTE_UPDATES.map(function () { return false; });

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();
    var sheetRow = firstDataRow + i;

    for (var u = 0; u < NOTE_UPDATES.length; u++) {
      if (doneFlags[u]) continue;
      var upd = NOTE_UPDATES[u];
      if (day === upd.day && block === upd.block && exercise === upd.exercise) {
        var currentNote = String(row[noteCol]).trim();
        if (currentNote === upd.newNote) {
          Logger.log('"' + upd.exercise + '": ALREADY DONE (row ' + sheetRow + ' note already matches) -- skipped.');
        } else {
          Logger.log('"' + upd.exercise + '": ' + (DRY_RUN ? 'WOULD UPDATE' : 'UPDATED') + ' row ' + sheetRow + ' note.');
          if (!DRY_RUN) {
            sheet.getRange(sheetRow, noteCol + 1).setValue(upd.newNote);
          }
        }
        doneFlags[u] = true;
      }
    }
  }

  for (var v = 0; v < NOTE_UPDATES.length; v++) {
    if (!doneFlags[v]) {
      Logger.log('"' + NOTE_UPDATES[v].exercise + '": NOT FOUND -- no row matched ' + NOTE_UPDATES[v].day + ' / ' + NOTE_UPDATES[v].block + ' / "' + NOTE_UPDATES[v].exercise + '".');
    }
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
