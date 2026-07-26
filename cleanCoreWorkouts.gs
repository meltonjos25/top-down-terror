/**
 * Comprehensive Core-workout cleanup, built from a direct read of the
 * live Google Sheet (not assumptions from the app's local fallback data).
 *
 * Scope: all 6 Core tabs that currently exist --
 *   Husband Core P1C1, Husband Core P2C1, Husband Core P3C1,
 *   Wife Core P1C1, Wife Core P2C1, Wife Core P3C1
 * (Phase 4 tabs do not exist yet -- that phase has not started.)
 *
 * What this removes and replaces:
 *   - Squats, jump squats, and squat-pattern isometric holds (e.g. a
 *     wall sit)
 *   - Burpees, in any variation
 *   - Marching / high-knee / lateral-step "jumping movements"
 *   - Non-glute leg-dominant work: hip hinges, leg press, single-leg
 *     standing balance, leg slides driven by the leg rather than the
 *     core
 *   - Rowing-pattern exercises (Seated Band Core Row Tan)
 *   - "Hip Flexor Lunge Stretch" is renamed to "Kneeling Hip Flexor
 *     Stretch" everywhere it appears -- same stretch, no "lunge" in
 *     the name
 *
 * What this deliberately keeps (glute-focused work is fine, and moves
 * where the leg is incidental rather than the primary target are fine):
 *   - All glute bridge variations, Side-Lying Clam / Band Clamshell,
 *     Standing/Side-Lying Hip Abduction, Prone Hip Extension, Modified
 *     Clamshell, Fast Glute Bridge
 *   - Mountain Climbers (all variants), Flutter Kicks, Speed/Sprint
 *     Dead Bug, Plank Knee Drives, Speed Bird Dog -- core is the
 *     primary target even though the legs move
 *   - "Side Plank with Band Row Blue" (Husband Core P2C1, Wednesday) --
 *     kept per explicit prior confirmation; it is a side-plank
 *     anti-rotation hold with an incidental band row, not a back
 *     exercise
 *
 * Every replacement is a genuine core/anti-rotation exercise or a
 * bodyweight/band core movement matched to that day's phase tempo, and
 * no replacement repeats an exercise already present elsewhere in that
 * same day.
 *
 * For each change, this REPLACES the row in place (same row, same
 * Block) rather than deleting it, so row order is untouched. The
 * Exercise Demo Link column is NEVER touched, on any row. No tab is
 * reordered, renamed, added, or deleted, and no row outside this list
 * is ever written to.
 *
 * The real header row on this Sheet is:
 *   Day | Block | Exercise | Sets | Target Reps / Time | Exercise Demo Link | Coaching Note
 * (note the spaces around the slash, and "Exercise Demo Link" -- not
 * "Target Reps/Time" / "Demo Link" like earlier scripts assumed).
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD change
 *   (View > Logs, or Execution log) and edits nothing. Once the log
 *   looks right, set DRY_RUN to false and run cleanCoreWorkouts() again
 *   to actually apply it.
 *
 * Run once by selecting cleanCoreWorkouts from the function dropdown
 * and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

var EXPECTED_HEADERS = ['Day', 'Block', 'Exercise', 'Sets', 'Target Reps / Time', 'Exercise Demo Link', 'Coaching Note'];

// Exercise swaps: old name -> new name, sets/target/note.
var CHANGES = [
  // ---- Husband Core P1C1 ----
  {
    tab: 'Husband Core P1C1', day: 'Monday', block: 'Block 2',
    oldName: 'Reverse Hyper on Table', newName: 'Bird Dog Hold',
    sets: '3', target: '20 sec/side',
    note: 'On all fours, extend opposite arm and leg. Hold still, hips level, back flat -- never let your hips rotate or sag.'
  },
  {
    tab: 'Husband Core P1C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Banded Hip March', newName: 'Side Plank Hold',
    sets: '3', target: '20 sec/side',
    note: 'Forearm on floor, stack your feet, and lift your hips to form a straight line. Core braced, hips level throughout.'
  },
  {
    tab: 'Husband Core P1C1', day: 'Friday', block: 'Cool-Down',
    oldName: 'Hip Flexor Lunge Stretch', newName: 'Kneeling Hip Flexor Stretch',
    sets: '1', target: '30 sec/side',
    note: 'Kneel on your back knee, front foot planted. Tuck your pelvis slightly and keep your torso upright. Feel the stretch in the front of your hip.'
  },

  // ---- Husband Core P2C1 ----
  {
    tab: 'Husband Core P2C1', day: 'Monday', block: 'Warm-Up',
    oldName: 'Hip Flexor Lunge Stretch', newName: 'Kneeling Hip Flexor Stretch',
    sets: '1', target: '30 sec/side',
    note: 'Kneel on your back knee, front foot planted. Tuck your pelvis slightly and keep your torso upright. Feel the stretch in the front of your hip.'
  },
  {
    tab: 'Husband Core P2C1', day: 'Friday', block: 'Warm-Up',
    oldName: 'Hip Flexor Lunge Stretch', newName: 'Kneeling Hip Flexor Stretch',
    sets: '1', target: '30 sec/side',
    note: 'Kneel on your back knee, front foot planted. Tuck your pelvis slightly and keep your torso upright. Feel the stretch in the front of your hip.'
  },

  // ---- Husband Core P3C1 ----
  {
    tab: 'Husband Core P3C1', day: 'Wednesday', block: 'Finisher',
    oldName: 'Tabata Burpee', newName: 'Tabata Mountain Climbers',
    sets: '1', target: '8 rounds: 20s on/10s off',
    note: 'Drive your knees in fast and controlled for the full interval -- speed is the goal, and never hold your breath.'
  },
  {
    tab: 'Husband Core P3C1', day: 'Friday', block: 'Warm-Up',
    oldName: 'Hip Flexor Lunge Stretch', newName: 'Kneeling Hip Flexor Stretch',
    sets: '1', target: '30 sec/side',
    note: 'Kneel on your back knee, front foot planted. Tuck your pelvis slightly and keep your torso upright. Feel the stretch in the front of your hip.'
  },

  // ---- Wife Core P1C1 ----
  {
    tab: 'Wife Core P1C1', day: 'Monday', block: 'Block 1',
    oldName: 'Kneeling Hip Hinge', newName: 'Seated Band Pallof Press',
    sets: '3', target: '10/side',
    note: 'Sit tall with the band anchored at chest height to your side. Press both hands straight out and hold 2 seconds -- resist any pull to rotate. Breathe steadily and never hold your breath.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Wednesday', block: 'Block 1',
    oldName: 'Hip Hinge with Band', newName: 'Standing Band Pallof Press',
    sets: '3', target: '12/side',
    note: 'Stand tall with the band anchored at chest height to your side. Press both hands straight out and hold 2 seconds -- resist any pull to rotate. Breathe steadily.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Quadruped Leg Slide', newName: 'Quadruped Arm Reach',
    sets: '3', target: '10/side',
    note: 'On all fours, extend one arm straight forward. Hold 2 seconds, core braced, hips level throughout.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Seated Leg Press with Band', newName: 'Seated Band Chest Press',
    sets: '3', target: '12/side',
    note: 'Sit tall with the band anchored behind you at chest height. Press both hands straight forward, squeeze your chest, and control the return.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Wednesday', block: 'Finisher',
    oldName: 'Standing Balance Hold', newName: 'Plank Hold',
    sets: '1', target: 'To failure',
    note: 'Forearm plank, hold as long as possible with your hips level and core braced. Breathe steadily throughout.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Friday', block: 'Block 1',
    oldName: 'Standing March with Hold', newName: 'Bear Hold',
    sets: '3', target: '25 sec',
    note: 'On all fours, lift your knees 1 inch off the floor. Keep your back flat like a table and breathe steadily -- never hold your breath.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Friday', block: 'Block 1',
    oldName: 'Supported Single Leg Stand', newName: 'Forearm Plank Hold',
    sets: '3', target: '20 sec',
    note: 'Forearms on the floor, body in a straight line from head to heels. Brace your core, breathe steadily, and don\'t let your hips sag or pike.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Friday', block: 'Block 1',
    oldName: 'Band Hip Hinge', newName: 'Standing Band Pallof Press',
    sets: '3', target: '12/side',
    note: 'Stand tall with the band anchored at chest height to your side. Press both hands straight out and hold 2 seconds -- resist any pull to rotate. Breathe steadily.'
  },
  {
    tab: 'Wife Core P1C1', day: 'Friday', block: 'Finisher',
    oldName: 'Breathing Wall Sit', newName: 'Max Bear Hold',
    sets: '1', target: 'To failure',
    note: 'Lift your knees 1 inch off the floor, back completely flat. Breathe slowly and steadily, and hold until your form breaks.'
  },

  // ---- Wife Core P2C1 ----
  {
    tab: 'Wife Core P2C1', day: 'Monday', block: 'Block 2',
    oldName: 'Seated Band Core Row Tan', newName: 'Band Pallof Press Tan',
    sets: '3', target: '10/side',
    note: 'Anchor the tan band at chest height and sit or stand sideways to it. Press both hands straight out and hold 2 seconds -- resist any rotation. Breathe steadily and never hold your breath.'
  },

  // ---- Wife Core P3C1 ----
  {
    tab: 'Wife Core P3C1', day: 'Monday', block: 'Block 1',
    oldName: 'Modified Burpee No Jump', newName: 'Speed Sit-Up',
    sets: '3', target: '15 reps fast',
    note: 'Drive your torso up as fast as possible, hands behind your head. Keep the movement explosive but controlled, and breathe steadily throughout.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Monday', block: 'Block 2',
    oldName: 'Standing Speed March High Knees', newName: 'Speed Plank Reach',
    sets: '3', target: '20 reps fast',
    note: 'In a forearm plank, alternate reaching each arm forward as fast as possible. Keep your hips level and still -- never let speed cause rocking.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Monday', block: 'Block 2',
    oldName: 'Small Squat Jump Land Soft', newName: 'Band Pallof Press Fast',
    sets: '3', target: '10/side',
    note: 'Anchor the band at chest height, press out fast and return with control. Resist rotation throughout, and breathe steadily.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Wednesday', block: 'Block 1',
    oldName: 'Lateral Step to Plank', newName: 'Plank Shoulder Tap Fast',
    sets: '3', target: '20 reps fast',
    note: 'In a plank, alternate tapping each shoulder as fast as possible. Keep your hips level and still -- never let speed cause rocking.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Speed Standing March', newName: 'Fast Hollow Body Rock',
    sets: '3', target: '30 sec',
    note: 'Press your lower back into the floor in hollow position and rock forward and back as fast as possible. Maintain total body tension and breathe steadily.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Quick Lateral Steps', newName: 'Band Pallof Press Fast',
    sets: '3', target: '10/side',
    note: 'Anchor the band at chest height, press out fast and return with control. Resist rotation throughout, and breathe steadily.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Wednesday', block: 'Finisher',
    oldName: 'Power Tabata Modified', newName: 'Power Tabata Mountain Climbers',
    sets: '1', target: '8 rounds: 20s on/10s off',
    note: 'Go as hard as you can for each 20-second interval, driving your knees in fast and controlled. Breathe steadily throughout and never hold your breath.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Friday', block: 'Block 1',
    oldName: 'Small Squat Jump', newName: 'Speed Sit-Up',
    sets: '3', target: '15 reps fast',
    note: 'Drive your torso up as fast as possible, hands behind your head. Keep the movement explosive but controlled, and breathe steadily throughout.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Friday', block: 'Block 1',
    oldName: 'Modified Burpee', newName: 'Plank Shoulder Tap Fast',
    sets: '3', target: '20 reps fast',
    note: 'In a plank, alternate tapping each shoulder as fast as possible. Keep your hips level and still -- never let speed cause rocking.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Friday', block: 'Block 2',
    oldName: 'Speed March', newName: 'Band Pallof Press Fast',
    sets: '3', target: '10/side',
    note: 'Anchor the band at chest height, press out fast and return with control. Resist rotation throughout, and breathe steadily.'
  }
];

// Finisher rows whose exercise NAME is not changing, but whose coaching
// note referenced a now-removed exercise (e.g. "jumping and marching")
// and needs correcting so it stays accurate.
var NOTE_ONLY_UPDATES = [
  {
    tab: 'Wife Core P3C1', day: 'Friday', block: 'Finisher',
    exercise: 'Power Flow',
    newNote: 'Move immediately from one exercise to the next. Breathe steadily throughout, and never hold your breath.'
  }
];

function cleanCoreWorkouts() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var results = [];

  for (var i = 0; i < CHANGES.length; i++) {
    results.push(applyChange(ss, CHANGES[i]));
  }
  for (var k = 0; k < NOTE_ONLY_UPDATES.length; k++) {
    results.push(applyNoteUpdate(ss, NOTE_ONLY_UPDATES[k]));
  }

  var okCount = 0;
  var missingCount = 0;
  var lines = [];
  for (var j = 0; j < results.length; j++) {
    lines.push(results[j].message);
    if (results[j].status === 'OK') okCount++;
    else missingCount++;
  }

  var totalExpected = CHANGES.length + NOTE_ONLY_UPDATES.length;
  Logger.log((DRY_RUN ? '[DRY RUN -- nothing changed]\n' : '[APPLIED]\n') + lines.join('\n'));
  Logger.log(okCount + ' of ' + totalExpected + ' rows matched' + (DRY_RUN ? ' (would be updated).' : ' and updated.'));
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
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (tab headers do not match -- actual headers were: ' + headerRow.join(' | ') + ')' };
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

function applyNoteUpdate(ss, upd) {
  var label = upd.tab + ' / ' + upd.day + ' / ' + upd.block + ' / "' + upd.exercise + '" (note only)';

  var sheet = ss.getSheetByName(upd.tab);
  if (!sheet) {
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (no tab named "' + upd.tab + '")' };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) {
    return String(h).trim();
  });
  if (!headersMatch(headerRow, EXPECTED_HEADERS)) {
    return { status: 'NOT FOUND', message: label + ': NOT FOUND (tab headers do not match -- actual headers were: ' + headerRow.join(' | ') + ')' };
  }

  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var noteCol = headerRow.indexOf('Coaching Note');

  var data = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();

    if (day === upd.day && block === upd.block && exercise === upd.exercise) {
      var sheetRow = i + 2;
      if (!DRY_RUN) {
        sheet.getRange(sheetRow, noteCol + 1).setValue(upd.newNote);
      }
      return { status: 'OK', message: upd.tab + ' / ' + upd.day + ' / row ' + sheetRow + ': OK -- updated note for "' + upd.exercise + '"' };
    }
  }

  return { status: 'NOT FOUND', message: label + ': NOT FOUND (no row matched that Day + Block + Exercise name)' };
}

function headersMatch(actual, expected) {
  if (actual.length < expected.length) return false;
  for (var i = 0; i < expected.length; i++) {
    if (actual[i] !== expected[i]) return false;
  }
  return true;
}
