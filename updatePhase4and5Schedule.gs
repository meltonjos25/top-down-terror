function updatePhase4and5Schedule() {
  var DRY_RUN = true; // review the log first, then set to false and run again to apply

  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  var EXPECTED_HEADER = ['Day', 'Block', 'Exercise', 'Sets', 'Target Reps / Time', 'Exercise Demo Link', 'Coaching Note'];

  // For each tab: keepDay stays untouched, renameFrom->renameTo relabels a whole
  // day block in place, removeDay is deleted entirely (rows removed).
  var JOBS = [
    { name: 'Husband Core P4C1', keepDay: 'Monday',  renameFrom: 'Wednesday', renameTo: 'Thursday', removeDay: 'Friday' },
    { name: 'Husband Core P5C1', keepDay: 'Monday',  renameFrom: 'Wednesday', renameTo: 'Thursday', removeDay: 'Friday' },
    { name: 'Husband Arms P4C1', keepDay: 'Tuesday', renameFrom: 'Thursday',  renameTo: 'Friday',   removeDay: 'Saturday' },
    { name: 'Husband Arms P5C1', keepDay: 'Tuesday', renameFrom: 'Thursday',  renameTo: 'Friday',   removeDay: 'Saturday' },
  ];

  function processSheet(job) {
    var name = job.name;
    var sheet = ss.getSheetByName(name);
    if (!sheet) { Logger.log('SKIP (tab not found): ' + name); return; }

    var lastRow = sheet.getLastRow();
    if (lastRow < 3) { Logger.log('SKIP (no data rows): ' + name); return; }

    var header = sheet.getRange(2, 1, 1, 7).getValues()[0];
    for (var h = 0; h < EXPECTED_HEADER.length; h++) {
      if (header[h] !== EXPECTED_HEADER[h]) {
        throw new Error(name + ': unexpected header in row 2, col ' + (h + 1) +
          '. Expected "' + EXPECTED_HEADER[h] + '", found "' + header[h] + '". Aborting without changing anything.');
      }
    }

    var numRows = lastRow - 2;
    var dayCol = sheet.getRange(3, 1, numRows, 1).getValues();

    var renameRows = [];
    var removeRows = [];
    var unexpectedDays = [];

    for (var i = 0; i < numRows; i++) {
      var rowNum = i + 3;
      var d = dayCol[i][0];
      if (d === job.keepDay || d === job.renameTo) continue; // already correct, nothing to do
      if (d === job.renameFrom) { renameRows.push(rowNum); continue; }
      if (d === job.removeDay) { removeRows.push(rowNum); continue; }
      unexpectedDays.push(rowNum + ': "' + d + '"');
    }

    if (unexpectedDays.length > 0) {
      throw new Error(name + ': found row(s) with an unexpected Day value, aborting without changing anything: ' + unexpectedDays.join('; '));
    }

    Logger.log(name + ' -- keep "' + job.keepDay + '" / already "' + job.renameTo + '": ' +
      (numRows - renameRows.length - removeRows.length) + ' row(s) unchanged');
    Logger.log(name + ' -- rename "' + job.renameFrom + '" -> "' + job.renameTo + '": rows ' +
      (renameRows.length ? renameRows.join(', ') : '(none)'));
    Logger.log(name + ' -- remove "' + job.removeDay + '": rows ' +
      (removeRows.length ? removeRows.join(', ') : '(none)'));

    if (DRY_RUN) {
      Logger.log(name + ': DRY RUN -- no changes applied.');
      return;
    }

    // Delete from the bottom up so earlier row numbers (including the rename
    // rows, which always sit above the removed block) never shift mid-loop.
    removeRows.sort(function(a, b) { return b - a; });
    for (var r = 0; r < removeRows.length; r++) {
      sheet.deleteRow(removeRows[r]);
    }

    for (var j = 0; j < renameRows.length; j++) {
      sheet.getRange(renameRows[j], 1).setValue(job.renameTo);
    }

    // Self-check: after applying, every remaining row's Day must be keepDay or renameTo.
    var finalRows = sheet.getLastRow() - 2;
    var finalDayCol = finalRows > 0 ? sheet.getRange(3, 1, finalRows, 1).getValues() : [];
    for (var k = 0; k < finalDayCol.length; k++) {
      var fd = finalDayCol[k][0];
      if (fd !== job.keepDay && fd !== job.renameTo) {
        throw new Error(name + ': post-check failed at row ' + (k + 3) + ' -- found "' + fd +
          '", expected only "' + job.keepDay + '" or "' + job.renameTo + '". Sheet may be in a bad state, please review manually.');
      }
    }

    Logger.log(name + ': applied -- ' + removeRows.length + ' row(s) deleted, ' +
      renameRows.length + ' row(s) relabeled "' + job.renameFrom + '" -> "' + job.renameTo + '".');
  }

  JOBS.forEach(processSheet);

  SpreadsheetApp.flush();
  Logger.log(DRY_RUN
    ? 'DRY RUN complete -- review the log above, then set DRY_RUN = false at the top of the function and run again to apply.'
    : 'Done -- Husband Core/Arms P4C1 and P5C1 now reflect the 2-day-per-week schedule (Core: Monday/Thursday, Arms: Tuesday/Friday). Wife tabs and Phase 1-3 tabs were not touched.');
}
