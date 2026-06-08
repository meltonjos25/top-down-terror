function populatePhase2Sheets() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  var BLOCK_BG = {
    'Warm-Up':   '#e8f5e9',
    'Block 1':   '#e3f2fd',
    'Block 2':   '#fce4ec',
    'Finisher':  '#fff8e1',
    'Cool-Down': '#e0f7fa'
  };

  function buildSheet(name, exercises) {
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    var sheet = ss.insertSheet(name);
    var rows = [['Day','Block','Exercise','Sets','Target Reps/Time','Exercise Demo Link','Coaching Note']];
    for (var i = 0; i < exercises.length; i++) {
      var e = exercises[i];
      rows.push([e[0], e[1], e[2], e[3], e[4], '', e[5]]);
    }
    sheet.getRange(1, 1, rows.length, 7).setValues(rows);
    var hdr = sheet.getRange(1, 1, 1, 7);
    hdr.setBackground('#263238');
    hdr.setFontColor('#ffffff');
    hdr.setFontWeight('bold');
    for (var r = 1; r < rows.length; r++) {
      var bg = BLOCK_BG[rows[r][1]] || '#ffffff';
      var rn = r + 1;
      sheet.getRange(rn, 1).setBackground('#fff3e0').setFontWeight('bold');
      sheet.getRange(rn, 2).setBackground(bg).setFontWeight('bold');
      sheet.getRange(rn, 3, 1, 5).setBackground(bg);
    }
    sheet.setColumnWidth(1, 110);
    sheet.setColumnWidth(2, 130);
    sheet.setColumnWidth(3, 270);
    sheet.setColumnWidth(4, 60);
    sheet.setColumnWidth(5, 200);
    sheet.setColumnWidth(6, 100);
    sheet.setColumnWidth(7, 330);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 4, rows.length, 1).setHorizontalAlignment('center');
    Logger.log('Created: ' + name);
  }

  // Each exercise row: [day, block, name, sets, target, coaching note]

  buildSheet('Husband Core P2C1', [
    ['Monday','Warm-Up',   'Cat-Camel Stretch',              1, '10 reps',           ''],
    ['Monday','Warm-Up',   'Dead Bug Arms Only',             1, '10/side',            ''],
    ['Monday','Warm-Up',   'Band Pull-Apart Green',          1, '15 reps',            ''],
    ['Monday','Warm-Up',   'Hip Flexor Lunge Stretch',       1, '30 sec/side',        ''],
    ['Monday','Block 1',   'Copenhagen Plank Knee Down',     3, '20 sec/side',        ''],
    ['Monday','Block 1',   'Hollow Body Hold',               3, '40 sec',             ''],
    ['Monday','Block 1',   'L-Sit Hold on Chairs',           3, '20 sec',             ''],
    ['Monday','Block 2',   'Dragon Flag Negative',           3, '5 reps slow',        ''],
    ['Monday','Block 2',   'Stir the Pot on Forearms',       3, '8/direction',        ''],
    ['Monday','Block 2',   'Band Pallof Press Blue',         3, '12/side',            ''],
    ['Monday','Finisher',  'Max L-Sit Hold on Chairs',       1, '1 set to failure',   ''],
    ['Monday','Cool-Down', "Child's Pose",                   1, '45 sec',             ''],
    ['Monday','Cool-Down', 'Cobra Stretch',                  1, '30 sec',             ''],
    ['Monday','Cool-Down', 'Seated Spinal Twist',            1, '30 sec/side',        ''],
    ['Wednesday','Warm-Up',   'Hip Circles',                     1, '10/direction',       ''],
    ['Wednesday','Warm-Up',   'Lying Knee Rocks',                1, '12 reps',            ''],
    ['Wednesday','Warm-Up',   'Bear Crawl In Place',             1, '30 sec',             ''],
    ['Wednesday','Warm-Up',   'Inchworm Walk-Out',               1, '6 reps',             ''],
    ['Wednesday','Block 1',   'Slow Negative Push-Up',           3, '8 reps',             '4 sec down'],
    ['Wednesday','Block 1',   'Band Resisted Crunch Blue',       3, '15 reps',            ''],
    ['Wednesday','Block 1',   'Jefferson Curl with Band Tan',    3, '8 reps',             ''],
    ['Wednesday','Block 2',   'Side Plank with Band Row Blue',   3, '10/side',            ''],
    ['Wednesday','Block 2',   'Hollow Body Rock',                3, '35 sec',             ''],
    ['Wednesday','Block 2',   'Dead Bug with Band Tan',          3, '8/side',             ''],
    ['Wednesday','Finisher',  'Max Slow Negative Push-Ups',      1, '1 set to failure',   ''],
    ['Wednesday','Cool-Down', 'Lying Spinal Twist',              1, '45 sec/side',        ''],
    ['Wednesday','Cool-Down', 'Figure-4 Stretch',                1, '30 sec/side',        ''],
    ['Wednesday','Cool-Down', "Child's Pose",                    1, '30 sec',             ''],
    ['Friday','Warm-Up',   'Cat-Camel Stretch',              1, '10 reps',           ''],
    ['Friday','Warm-Up',   'Band Pull-Apart Green',          1, '15 reps',           ''],
    ['Friday','Warm-Up',   'Hip Flexor Lunge Stretch',       1, '30 sec/side',       ''],
    ['Friday','Warm-Up',   'Arm Circles',                    1, '10 each direction', ''],
    ['Friday','Block 1',   'Copenhagen Plank Full',          3, '25 sec/side',       ''],
    ['Friday','Block 1',   'Dragon Flag Negative',           3, '6 reps',            ''],
    ['Friday','Block 1',   'Stir the Pot',                   3, '10/direction',      ''],
    ['Friday','Block 2',   'Band Resisted Hollow Body Blue', 3, '30 sec',            ''],
    ['Friday','Block 2',   'L-Sit Hold on Chairs',           3, '25 sec',            ''],
    ['Friday','Block 2',   'Band Pallof Press Blue',         3, '12/side',           ''],
    ['Friday','Finisher',  'Strength Core Gauntlet',         1, '1 round',           'Copenhagen 20s/side -> Dragon Flag 5 -> L-Sit 20s, no rest'],
    ['Friday','Cool-Down', "Child's Pose",                   1, '45 sec',            ''],
    ['Friday','Cool-Down', 'Supine Hamstring Stretch',       1, '30 sec/side',       ''],
    ['Friday','Cool-Down', 'Cobra to Upward Dog',            1, '30 sec',            ''],
  ]);

  buildSheet('Husband Arms P2C1', [
    ['Tuesday','Warm-Up',   'Arm Circles',                        1, '15 each direction', ''],
    ['Tuesday','Warm-Up',   'Band Pull-Apart Green',              1, '15 reps',           ''],
    ['Tuesday','Warm-Up',   'Wrist Circles',                      1, '10 each direction', ''],
    ['Tuesday','Warm-Up',   'Shoulder Cross-Body Stretch',        1, '20 sec/side',       ''],
    ['Tuesday','Block 1',   'Slow Negative Table Inverted Row',   3, '8 reps',            '4 sec down'],
    ['Tuesday','Block 1',   'Band Curl Blue',                     3, '12 reps',           ''],
    ['Tuesday','Block 1',   'Isometric Table Inverted Row Hold',  3, '30 sec',            ''],
    ['Tuesday','Block 2',   'Slow Negative Close-Grip Push-Up',   3, '8 reps',            '4 sec down'],
    ['Tuesday','Block 2',   'Band Tricep Pushdown Blue',          3, '15 reps',           ''],
    ['Tuesday','Block 2',   'Tricep Dip Negative',                3, '8 reps',            '4 sec down'],
    ['Tuesday','Finisher',  'Max Slow Negative Push-Ups',         1, '1 set to failure',  ''],
    ['Tuesday','Cool-Down', 'Overhead Tricep Stretch',            1, '30 sec/side',       ''],
    ['Tuesday','Cool-Down', 'Cross-Body Bicep Stretch',           1, '30 sec/side',       ''],
    ['Thursday','Warm-Up',   'Arm Circles',                       1, '15 each direction', ''],
    ['Thursday','Warm-Up',   'Band Pull-Apart Green',             1, '15 reps',           ''],
    ['Thursday','Warm-Up',   'Cat-Camel',                         1, '10 reps',           ''],
    ['Thursday','Warm-Up',   'Wrist Circles',                     1, '10 each direction', ''],
    ['Thursday','Block 1',   'Slow Negative Pike Push-Up',        3, '8 reps',            '4 sec down'],
    ['Thursday','Block 1',   'Wall Handstand Hold',               3, '25 sec',            ''],
    ['Thursday','Block 1',   'Band Overhead Press Blue',          3, '12 reps',           ''],
    ['Thursday','Block 2',   'Typewriter Push-Up',                3, '6/side',            ''],
    ['Thursday','Block 2',   'Pseudo Planche Lean',               3, '20 sec',            ''],
    ['Thursday','Block 2',   'Band Face Pull Green',              3, '15 reps',           ''],
    ['Thursday','Finisher',  'Max Typewriter Push-Ups',           1, '1 set to failure',  ''],
    ['Thursday','Cool-Down', 'Doorway Chest Stretch',             1, '30 sec',            ''],
    ['Thursday','Cool-Down', "Child's Pose Wide Arms",            1, '45 sec',            ''],
    ['Saturday','Warm-Up',   'Jumping Jacks',                     1, '30 sec',            ''],
    ['Saturday','Warm-Up',   'Band Pull-Apart Green',             1, '15 reps',           ''],
    ['Saturday','Warm-Up',   'Inchworm Walk-Out',                 1, '6 reps',            ''],
    ['Saturday','Warm-Up',   'Wrist Circles',                     1, '10 each direction', ''],
    ['Saturday','Block 1',   'Pseudo Planche Push-Up',            3, '8 reps',            ''],
    ['Saturday','Block 1',   'Slow Negative Table Inverted Row',  3, '8 reps',            ''],
    ['Saturday','Block 1',   'Band Curl to Press Blue',           3, '10 reps',           ''],
    ['Saturday','Block 2',   'Typewriter Push-Up',                3, '6/side',            ''],
    ['Saturday','Block 2',   'Band Tricep Pushdown Blue',         3, '15 reps',           ''],
    ['Saturday','Block 2',   'Pike Push-Up Negative',             3, '8 reps',            ''],
    ['Saturday','Finisher',  'Strength Triset',                   1, '1 round',           'Pseudo Planche Push-Up 8 -> Band Curl Blue 12 -> Typewriter Push-Up 5/side, no rest'],
    ['Saturday','Cool-Down', 'Overhead Tricep Stretch',           1, '30 sec/side',       ''],
    ['Saturday','Cool-Down', 'Doorway Chest Opener',              1, '30 sec',            ''],
    ['Saturday','Cool-Down', 'Wrist Flexor Stretch',              1, '20 sec/side',       ''],
  ]);

  buildSheet('Wife Core P2C1', [
    ['Monday','Warm-Up',   'Diaphragmatic Breathing',               1, '8 breaths',          ''],
    ['Monday','Warm-Up',   'Pelvic Tilts',                          1, '12 reps',            ''],
    ['Monday','Warm-Up',   'Cat-Camel',                             1, '10 reps',            ''],
    ['Monday','Warm-Up',   'Band Pull-Apart Tan',                   1, '12 reps',            ''],
    ['Monday','Block 1',   'Modified Copenhagen Plank Knee Down',   3, '20 sec/side',        ''],
    ['Monday','Block 1',   'Hollow Body Hold Bent Knees',           3, '25 sec',             ''],
    ['Monday','Block 1',   'Bear Hold',                             3, '20 sec',             ''],
    ['Monday','Block 2',   'Dead Bug with Tan Band',                3, '8/side',             ''],
    ['Monday','Block 2',   'Slow Heel Slides',                      3, '10/side',            ''],
    ['Monday','Block 2',   'Seated Band Core Row Tan',              3, '12 reps',            ''],
    ['Monday','Finisher',  'Max Bear Hold',                         1, '1 set to failure',   'Breathe throughout'],
    ['Monday','Cool-Down', "Child's Pose",                          1, '45 sec',             ''],
    ['Monday','Cool-Down', 'Seated Spinal Twist',                   1, '30 sec/side',        ''],
    ['Monday','Cool-Down', 'Knees-to-Chest Hug',                    1, '30 sec',             ''],
    ['Wednesday','Warm-Up',   'Diaphragmatic Breathing',                1, '6 breaths',          ''],
    ['Wednesday','Warm-Up',   'Hip Circles Standing',                   1, '10/direction',       ''],
    ['Wednesday','Warm-Up',   'Lying Knee Rocks',                       1, '12 reps',            ''],
    ['Wednesday','Warm-Up',   'Glute Bridge',                           1, '12 reps',            ''],
    ['Wednesday','Block 1',   'Slow Negative Push-Up Knees Down',       3, '8 reps',             '4 sec down'],
    ['Wednesday','Block 1',   'Band Resisted Bird Dog Tan',             3, '10/side',            ''],
    ['Wednesday','Block 1',   'Side Plank Knees Down Hold',             3, '30 sec/side',        ''],
    ['Wednesday','Block 2',   'Hollow Body Hold Bent Knees',            3, '30 sec',             ''],
    ['Wednesday','Block 2',   'Glute Bridge with Band Blue',            3, '15 reps',            ''],
    ['Wednesday','Block 2',   'Band Pallof Press Tan',                  3, '12/side',            ''],
    ['Wednesday','Finisher',  'Max Slow Negative Push-Ups Knees Down',  1, '1 set to failure',   ''],
    ['Wednesday','Cool-Down', 'Lying Spinal Twist',                     1, '40 sec/side',        ''],
    ['Wednesday','Cool-Down', 'Figure-4 Glute Stretch',                 1, '30 sec/side',        ''],
    ['Wednesday','Cool-Down', "Child's Pose",                           1, '30 sec',             ''],
    ['Friday','Warm-Up',   'Diaphragmatic Breathing',               1, '8 breaths',          ''],
    ['Friday','Warm-Up',   'Pelvic Tilts',                          1, '12 reps',            ''],
    ['Friday','Warm-Up',   'Cat-Camel',                             1, '10 reps',            ''],
    ['Friday','Warm-Up',   'Standing Hip Flexor Stretch',           1, '30 sec/side',        ''],
    ['Friday','Block 1',   'Modified Copenhagen Plank',             3, '25 sec/side',        ''],
    ['Friday','Block 1',   'Bear Hold',                             3, '25 sec',             ''],
    ['Friday','Block 1',   'Band Pallof Press Tan',                 3, '12/side',            ''],
    ['Friday','Block 2',   'Dead Bug with Tan Band',                3, '10/side',            ''],
    ['Friday','Block 2',   'Slow Negative Push-Up Knees Down',      3, '8 reps',             ''],
    ['Friday','Block 2',   'Glute Bridge with Band Green',          3, '15 reps',            ''],
    ['Friday','Finisher',  'Strength Flow',                         1, '1 round',            'Bear Hold 20s -> Modified Copenhagen 20s/side -> Hollow Body Bent Knees 25s, no rest'],
    ['Friday','Cool-Down', "Child's Pose",                          1, '1 min',              ''],
    ['Friday','Cool-Down', 'Supine Spinal Twist',                   1, '40 sec/side',        ''],
    ['Friday','Cool-Down', 'Happy Baby Pose',                       1, '30 sec',             ''],
  ]);

  // Wife Bands P2C1 — arms and shoulders only (no lower body)
  buildSheet('Wife Bands P2C1', [
    // Tuesday: biceps / triceps / upper back
    ['Tuesday','Warm-Up',   'Diaphragmatic Breathing',             1, '6 breaths',              ''],
    ['Tuesday','Warm-Up',   'Band Pull-Apart Green',               1, '12 reps',                ''],
    ['Tuesday','Warm-Up',   'Wrist Circles',                       1, '10/direction',            ''],
    ['Tuesday','Warm-Up',   'Shoulder Rolls',                      1, '10 each direction',       ''],
    ['Tuesday','Block 1',   'Band Bicep Curl Slow Negative Green', 3, '10 reps',                '4 sec down'],
    ['Tuesday','Block 1',   'Band Overhead Press Green',           3, '10 reps',                ''],
    ['Tuesday','Block 1',   'Band Upright Row Green',              3, '10 reps',                ''],
    ['Tuesday','Block 2',   'Band Bicep Isometric Hold Green',     3, '20 sec',                 ''],
    ['Tuesday','Block 2',   'Band Tricep Overhead Extension Green',3, '12 reps',                ''],
    ['Tuesday','Block 2',   'Band Reverse Fly Tan',                3, '12 reps',                ''],
    ['Tuesday','Finisher',  'Band Curl Burnout Green',             1, 'Max reps',               ''],
    ['Tuesday','Cool-Down', "Child's Pose",                        1, '45 sec',                 ''],
    ['Tuesday','Cool-Down', 'Overhead Tricep Stretch',             1, '30 sec/side',             ''],
    ['Tuesday','Cool-Down', 'Cross-Body Shoulder Stretch',         1, '20 sec/side',             ''],
    // Thursday: shoulders / upper back / rear delts
    ['Thursday','Warm-Up',   'Diaphragmatic Breathing',            1, '6 breaths',              ''],
    ['Thursday','Warm-Up',   'Band Pull-Apart Green',              1, '12 reps',                ''],
    ['Thursday','Warm-Up',   'Arm Circles',                        1, '10/direction',            ''],
    ['Thursday','Warm-Up',   'Shoulder Rolls',                     1, '10 each direction',       ''],
    ['Thursday','Block 1',   'Band Row Slow Negative Green',       3, '10 reps',                '4 sec down'],
    ['Thursday','Block 1',   'Band Face Pull Green',               3, '12 reps',                ''],
    ['Thursday','Block 1',   'Band Reverse Fly Tan',               3, '12 reps',                ''],
    ['Thursday','Block 2',   'Band Lateral Raise Green',           3, '12/side',                ''],
    ['Thursday','Block 2',   'Band Front Raise Tan',               3, '12 reps',                ''],
    ['Thursday','Block 2',   'Band External Rotation Tan',         3, '12/side',                ''],
    ['Thursday','Finisher',  'Row Burnout',                        1, 'Max reps',               'Blue band. Seated row to failure.'],
    ['Thursday','Cool-Down', 'Doorway Chest Stretch',              1, '30 sec',                 ''],
    ['Thursday','Cool-Down', 'Overhead Tricep Stretch',            1, '30 sec/side',             ''],
    ['Thursday','Cool-Down', 'Wrist Flexor Stretch',               1, '20 sec/side',             ''],
  ]);

  SpreadsheetApp.flush();
  Logger.log('Done — all 4 Phase 2 Strength tabs created successfully.');
}
