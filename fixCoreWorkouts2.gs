/**
 * Same 30 exercise swaps + 2 Finisher note fixes as fixCoreWorkouts.gs,
 * re-verified character-for-character against the diagnostic log from
 * diagnoseCoreWorkouts.gs. That prior script only matched 2 of 32 rows
 * despite the names being correct, which means the real problem is
 * something structural (header detection, column alignment, or
 * something else specific to most tabs) rather than a text typo.
 *
 * Two things are different in this version:
 *
 * 1. IDEMPOTENT: before trying to match the OLD exercise name, each
 *    change first checks whether the NEW name is already present at
 *    that Day + Block. If so, it is reported as ALREADY DONE and
 *    skipped -- this means the 2 rows that already got updated by the
 *    prior run are handled safely without needing to know which 2 they
 *    were, and the script is safe to re-run any number of times.
 *
 * 2. SELF-DIAGNOSING ON FAILURE: if a row still cannot be matched, the
 *    log does not just say "not found" -- it lists every Exercise name
 *    actually present at that Day + Block on the sheet right now, and
 *    for any that look similar to the expected name, it prints a
 *    character-by-character comparison with Unicode code points so an
 *    invisible/mismatched character becomes visible instead of hidden.
 *    It also reports plainly if the header row itself could not be
 *    located on that tab, which would explain every row on that tab
 *    failing at once.
 *
 * Scope, what is removed/replaced, and what is deliberately kept are
 * unchanged from fixCoreWorkouts.gs -- see that file's header comment
 * for the full explanation. Tab names: Husband Core P1C1/P2C1/P3C1,
 * Wife Core P1C1/P2C1/P3C1 (P4C1 tabs do not exist yet for either user).
 *
 * The Exercise Demo Link column is never touched. No tab is reordered,
 * renamed, added, or deleted, and no row outside this list is ever
 * written to.
 *
 * SAFETY:
 *   DRY_RUN defaults to true. The first run only logs what WOULD
 *   change (View > Logs) and edits nothing. Once the log shows every
 *   remaining row was found (or already done), set DRY_RUN to false
 *   and run fixCoreWorkouts2() again to actually apply it.
 *
 * Run once by selecting fixCoreWorkouts2 from the function dropdown
 * and clicking Run.
 */

var SPREADSHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';
var DRY_RUN = true; // set to false after reviewing the log, then re-run

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
    oldName: 'Side-Lying Hip Abduction', newName: 'Side Plank Hold',
    sets: '3', target: '20 sec/side',
    note: 'Forearm on floor, stack your feet, and lift your hips to form a straight line. Core braced, hips level throughout.'
  },
  {
    tab: 'Husband Core P1C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Prone Hip Extension', newName: 'Prone Cobra Hold',
    sets: '3', target: '30 sec',
    note: 'Face down, arms at your sides. Lift your chest and hands off the floor and hold. Breathe steadily throughout.'
  },
  {
    tab: 'Husband Core P1C1', day: 'Wednesday', block: 'Block 2',
    oldName: 'Banded Hip March', newName: 'Band Pull-Apart Hold',
    sets: '3', target: '20 sec',
    note: 'Tan band at chest height, arms straight. Pull apart and hold. Squeeze your shoulder blades together throughout.'
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
    tab: 'Husband Core P2C1', day: 'Friday', block: 'Block 2',
    oldName: 'L-Sit Hold on Chairs', newName: 'Bear Hold',
    sets: '3', target: '20 sec',
    note: 'Start on all fours and lift your knees just 1 inch off the floor. Keep your back flat like a table and breathe steadily -- never hold your breath.'
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

var NOTE_ONLY_UPDATES = [
  {
    tab: 'Husband Core P2C1', day: 'Friday', block: 'Finisher',
    exercise: 'Strength Core Gauntlet',
    newNote: 'Copenhagen 20s/side -> Dragon Flag 5 -> Bear Hold 20s, zero rest. Move immediately between exercises and maintain quality form throughout.'
  },
  {
    tab: 'Wife Core P3C1', day: 'Friday', block: 'Finisher',
    exercise: 'Power Flow',
    newNote: 'Move immediately from one exercise to the next. Breathe steadily throughout, and never hold your breath.'
  }
];

function fixCoreWorkouts2() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var results = [];

  for (var i = 0; i < CHANGES.length; i++) {
    results.push(applyChange(ss, CHANGES[i]));
  }
  for (var k = 0; k < NOTE_ONLY_UPDATES.length; k++) {
    results.push(applyNoteUpdate(ss, NOTE_ONLY_UPDATES[k]));
  }

  var okCount = 0, alreadyCount = 0, missingCount = 0;
  var lines = [];
  for (var j = 0; j < results.length; j++) {
    lines.push(results[j].message);
    if (results[j].status === 'OK') okCount++;
    else if (results[j].status === 'ALREADY_DONE') alreadyCount++;
    else missingCount++;
  }

  var totalExpected = CHANGES.length + NOTE_ONLY_UPDATES.length;
  Logger.log((DRY_RUN ? '[DRY RUN -- nothing changed]\n' : '[APPLIED]\n') + lines.join('\n'));
  Logger.log('');
  Logger.log(okCount + ' matched' + (DRY_RUN ? ' (would update)' : ' and updated') + ', ' + alreadyCount + ' already done, ' + missingCount + ' not found, out of ' + totalExpected + ' total.');
}

// Returns the header row info for a sheet by searching for the row
// whose first cell is exactly "Day" (row 1 on every Core tab is a
// merged title row, not the real header).
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

// Debug helper: shows a string with visible markers for any character
// outside normal printable ASCII, so an invisible/mismatched character
// becomes obvious instead of hidden.
function debugStr(str) {
  var out = '';
  for (var i = 0; i < str.length; i++) {
    var code = str.charCodeAt(i);
    if (code < 32 || code > 126) {
      out += '[U+' + ('0000' + code.toString(16).toUpperCase()).slice(-4) + ']';
    } else {
      out += str[i];
    }
  }
  return '"' + out + '" (len ' + str.length + ')';
}

function applyChange(ss, change) {
  var label = change.tab + ' / ' + change.day + ' / ' + change.block + ' / "' + change.oldName + '"';

  var sheet = ss.getSheetByName(change.tab);
  if (!sheet) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- no tab named "' + change.tab + '" exists in this spreadsheet.' };
  }

  var headerInfo = findHeaderRow(sheet);
  if (!headerInfo) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- could not locate a header row with "Day" in column A within the first 5 rows of tab "' + change.tab + '". This tab likely has a different layout than expected.' };
  }

  var headerRow = headerInfo.headers;
  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var setsCol = headerRow.indexOf('Sets');
  var targetCol = headerRow.indexOf('Target Reps / Time');
  var noteCol = headerRow.indexOf('Coaching Note');

  if (dayCol === -1 || blockCol === -1 || exerciseCol === -1) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- header row found at sheet row ' + headerInfo.rowIndex + ' but Day/Block/Exercise columns missing. Actual header: ' + JSON.stringify(headerRow) };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var firstDataRow = headerInfo.rowIndex + 1;
  if (lastRow < firstDataRow) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- no data rows below the header on tab "' + change.tab + '".' };
  }

  var data = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  var matchesAtDayBlock = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();
    var sheetRow = firstDataRow + i;

    if (day === change.day && block === change.block) {
      matchesAtDayBlock.push({ sheetRow: sheetRow, exercise: exercise });

      if (exercise === change.newName) {
        return { status: 'ALREADY_DONE', message: label + ': ALREADY DONE (row ' + sheetRow + ' already reads "' + change.newName + '") -- skipped.' };
      }
      if (exercise === change.oldName) {
        if (!DRY_RUN) {
          sheet.getRange(sheetRow, exerciseCol + 1).setValue(change.newName);
          if (setsCol !== -1) sheet.getRange(sheetRow, setsCol + 1).setValue(change.sets);
          if (targetCol !== -1) sheet.getRange(sheetRow, targetCol + 1).setValue(change.target);
          if (noteCol !== -1) sheet.getRange(sheetRow, noteCol + 1).setValue(change.note);
        }
        return { status: 'OK', message: change.tab + ' / ' + change.day + ' / row ' + sheetRow + ': OK -- "' + change.oldName + '" -> "' + change.newName + '"' };
      }
    }
  }

  var detail = 'NOT FOUND -- no exercise named "' + change.oldName + '" (or already "' + change.newName + '") at ' + change.day + ' / ' + change.block + ' on "' + change.tab + '".';
  if (matchesAtDayBlock.length === 0) {
    detail += ' No rows at all matched that Day + Block combination.';
  } else {
    detail += ' Rows that DID match that Day + Block:';
    for (var m = 0; m < matchesAtDayBlock.length; m++) {
      var found = matchesAtDayBlock[m].exercise;
      detail += '\n      row ' + matchesAtDayBlock[m].sheetRow + ': expected ' + debugStr(change.oldName) + ' vs actual ' + debugStr(found);
    }
  }
  return { status: 'NOT_FOUND', message: label + ': ' + detail };
}

function applyNoteUpdate(ss, upd) {
  var label = upd.tab + ' / ' + upd.day + ' / ' + upd.block + ' / "' + upd.exercise + '" (note only)';

  var sheet = ss.getSheetByName(upd.tab);
  if (!sheet) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- no tab named "' + upd.tab + '" exists.' };
  }

  var headerInfo = findHeaderRow(sheet);
  if (!headerInfo) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- could not locate a header row with "Day" in column A.' };
  }

  var headerRow = headerInfo.headers;
  var dayCol = headerRow.indexOf('Day');
  var blockCol = headerRow.indexOf('Block');
  var exerciseCol = headerRow.indexOf('Exercise');
  var noteCol = headerRow.indexOf('Coaching Note');

  if (dayCol === -1 || blockCol === -1 || exerciseCol === -1 || noteCol === -1) {
    return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- required columns missing. Actual header: ' + JSON.stringify(headerRow) };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var firstDataRow = headerInfo.rowIndex + 1;
  var data = sheet.getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol).getValues();

  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var day = String(row[dayCol]).trim();
    var block = String(row[blockCol]).trim();
    var exercise = String(row[exerciseCol]).trim();

    if (day === upd.day && block === upd.block && exercise === upd.exercise) {
      var sheetRow = firstDataRow + i;
      var currentNote = noteCol !== -1 ? String(row[noteCol]).trim() : '';
      if (currentNote === upd.newNote) {
        return { status: 'ALREADY_DONE', message: label + ': ALREADY DONE (row ' + sheetRow + ' note already matches) -- skipped.' };
      }
      if (!DRY_RUN) {
        sheet.getRange(sheetRow, noteCol + 1).setValue(upd.newNote);
      }
      return { status: 'OK', message: upd.tab + ' / ' + upd.day + ' / row ' + sheetRow + ': OK -- updated note for "' + upd.exercise + '"' };
    }
  }

  return { status: 'NOT_FOUND', message: label + ': NOT FOUND -- no row matched that Day + Block + Exercise name.' };
}
