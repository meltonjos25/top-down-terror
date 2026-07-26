/**
 * Applies the 8 confirmed Core-workout exercise replacements (removing
 * band-row / rowing-pattern exercises and replacing them with genuine
 * anti-rotation core exercises) directly to the live Google Sheet.
 *
 * Scope: exactly 8 rows, on exactly 6 named tabs:
 *   Husband Core P1C1, Husband Core P3C1, Husband Core P4C1,
 *   Wife Core P1C1, Wife Core P2C1, Wife Core P4C1
 * No other tab is read or written. No tab is reordered, renamed, added,
 * or deleted -- only specific cells on specific existing rows change.
 *
 * For each of the 8 rows, this script finds the row by matching Day +
 * Block + the CURRENT (old) Exercise name, then updates that row's
 * Exercise, Sets, Target Reps/Time, and Coaching Note cells to the new
 * values below. Updating Sets/Target Reps/Time along with the Exercise
 * name is intentional -- leaving the old sets/reps numbers attached to a
 * brand new exercise name would leave the sheet internally inconsistent
 * (e.g. "Bird Dog Hold" still showing the old "25 sec" from "Band Row
 * Hold" instead of its own prescribed "20 sec/side"). If you actually
 * want Sets/Target Reps/Time left untouched, delete the two setValue(...)
 * lines for setsCol and targetCol inside applyChange() below before
 * running.
 *
 * The Demo Link column is NEVER touched, on any row, under any
 * circumstance. No row other than these exact 8 is ever written to.
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD change
 *   (View > Logs, or Execution log) and edits nothing. Once the log looks
 *   right, set DRY_RUN to false and run updateCoreWorkouts() again to
 *   actually apply it.
 *
 * Run once by selecting updateCoreWorkouts from the function dropdown
 * and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

var EXPECTED_HEADERS = ['Day', 'Block', 'Exercise', 'Sets', 'Target Reps/Time', 'Demo Link', 'Coaching Note'];

var CHANGES = [
  {
    tab: 'Husband Core P1C1', day: 'Monday', block: 'Block 2',
    oldName: 'Band Row Hold', newName: 'Bird Dog Hold',
    sets: '3', target: '20 sec/side',
    note: 'Start on all fours and extend one arm and the opposite leg simultaneously. Hold for 2 seconds at full extension keeping your hips completely level and your lower back flat -- never let your hips rotate or sag.'
  },
  {
    tab: 'Husband Core P1C1', day: 'Friday', block: 'Block 1',
    oldName: 'Tall Kneeling Band Row', newName: 'Kneeling Pallof Press',
    sets: '3', target: '12/side',
    note: 'Kneel tall on both knees with the blue band anchored at chest height to your side. Press both hands straight out and hold 2 seconds -- resist any pull to rotate or lean. Keep hips square and breathe steadily.'
  },
  {
    tab: 'Husband Core P3C1', day: 'Wednesday', block: 'Block 1',
    oldName: 'Explosive Table Inverted Row', newName: 'Plank Reach Fast',
    sets: '3', target: '20 reps fast',
    note: 'In a forearm plank, alternate reaching each arm straight forward as fast as possible. Keep your hips completely level and still -- speed should never cause your hips to rock or rotate.'
  },
  {
    tab: 'Husband Core P4C1', day: 'Monday', block: 'Block 2',
    oldName: 'Single Arm Band Row', newName: 'Half-Kneeling Band Chop',
    sets: '3', target: '10/side',
    note: 'Kneel on one knee with the green band anchored high. Pull the band diagonally down across your body to the opposite hip. Resist rotation throughout -- your torso should stay square.'
  },
  {
    tab: 'Husband Core P4C1', day: 'Friday', block: 'Block 1',
    oldName: 'Single Arm Band Row Blue', newName: 'Single Arm Band Pallof Press',
    sets: '3', target: '10/side',
    note: 'Anchor the blue band at chest height and kneel or stand sideways to it. Press one hand straight out from your chest and hold 2 seconds -- resist any rotation. Switch sides and repeat.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Monday', block: 'Block 2',
    oldName: 'Seated Band Row', newName: 'Seated Band Pallof Press',
    sets: '3', target: '10/side',
    note: 'Sit tall with the band anchored at chest height to your side. Press both hands straight out and hold 2 seconds -- resist any pull to rotate your torso. Breathe steadily and never hold your breath.'
  },
  {
    tab: 'Wife Core P2C1', day: 'Monday', block: 'Block 2',
    oldName: 'Seated Band Core Row Tan', newName: 'Band Pallof Press Tan',
    sets: '3', target: '10/side',
    note: 'Anchor the tan band at chest height and sit or stand sideways to it. Press both hands straight out and hold 2 seconds -- resist any rotation. Breathe steadily and never hold your breath.'
  },
  {
    tab: 'Wife Core P4C1', day: 'Friday', block: 'Block 1',
    oldName: 'Single Arm Band Row Tan', newName: 'Single Arm Band Pallof Press Tan',
    sets: '3', target: '12/side',
    note: 'Anchor the tan band at chest height and stand sideways to it. Press one hand straight out and hold 2 seconds -- resist any pull to rotate. Breathe steadily and never hold your breath.'
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
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (tab headers do not match the expected Day/Block/Exercise/Sets/Target Reps/Time/Demo Link/Coaching Note layout)' };
  }

  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var setsCol = headerRow.indexOf('Sets');
  var targetCol = headerRow.indexOf('Target Reps/Time');
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
