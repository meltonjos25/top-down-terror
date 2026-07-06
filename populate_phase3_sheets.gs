function populatePhase3Sheets() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  // Phase 3 Power red
  var PHASE_BG    = '#f85149';   // red title banner
  var DAY_TINT    = '#fde8e8';   // light red for Day column
  var HEADER_BG   = '#263238';   // dark slate header row

  // Block background colors (same across all phases)
  var BLOCK_BG = {
    'Warm-Up':   '#e8f5e9',
    'Block 1':   '#dbeafe',
    'Block 2':   '#fce7f3',
    'Finisher':  '#000000',
    'Cool-Down': '#e0f2f1'
  };

  // Block label text colors
  var BLOCK_FG = {
    'Warm-Up':   '#14532d',
    'Block 1':   '#1e3a5f',
    'Block 2':   '#831843',
    'Finisher':  '#ffffff',
    'Cool-Down': '#134e4a'
  };

  // Tab titles
  var TAB_TITLES = {
    'Husband Core P3C1': 'Husband Core Workout — P3C1 (Monday · Wednesday · Friday) | Power Phase',
    'Husband Arms P3C1': 'Husband Arms Workout — P3C1 (Tuesday · Thursday · Saturday) | Power Phase',
    'Wife Core P3C1':    'Wife Core Workout — P3C1 (Monday · Wednesday · Friday) | Power Phase · Breastfeeding-Safe',
    'Wife Bands P3C1':   'Wife Bands Workout — P3C1 (Tuesday · Thursday) | Power Phase · Breastfeeding-Safe'
  };

  function buildSheet(name, exercises) {
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    var sheet = ss.insertSheet(name);

    // ── ROW 1: Title banner ──────────────────────────────────────────────────
    sheet.setRowHeight(1, 40);
    var titleRange = sheet.getRange(1, 1, 1, 7);
    titleRange.merge();
    titleRange.setValue(TAB_TITLES[name] || name);
    titleRange.setBackground(PHASE_BG);
    titleRange.setFontColor('#ffffff');
    titleRange.setFontWeight('bold');
    titleRange.setFontSize(13);
    titleRange.setHorizontalAlignment('center');
    titleRange.setVerticalAlignment('middle');

    // ── ROW 2: Column headers ────────────────────────────────────────────────
    sheet.setRowHeight(2, 28);
    var hdrRange = sheet.getRange(2, 1, 1, 7);
    hdrRange.setValues([['Day', 'Block', 'Exercise', 'Sets', 'Target Reps / Time', 'Exercise Demo Link', 'Coaching Note']]);
    hdrRange.setBackground(HEADER_BG);
    hdrRange.setFontColor('#ffffff');
    hdrRange.setFontWeight('bold');
    hdrRange.setFontSize(11);
    hdrRange.setHorizontalAlignment('center');
    hdrRange.setVerticalAlignment('middle');

    // ── ROWS 3+: Exercise data ───────────────────────────────────────────────
    var dataRows = [];
    for (var i = 0; i < exercises.length; i++) {
      var e = exercises[i];
      // [day, block, exercise, sets, target, demo_link(blank), note]
      dataRows.push([e[0], e[1], e[2], e[3], e[4], '', e[5]]);
    }
    if (dataRows.length === 0) return;

    sheet.getRange(3, 1, dataRows.length, 7).setValues(dataRows);

    for (var r = 0; r < dataRows.length; r++) {
      var rowNum  = r + 3;
      var block   = dataRows[r][1];
      var blockBg = BLOCK_BG[block] || '#ffffff';
      var blockFg = BLOCK_FG[block] || '#000000';
      var bodyFg  = (block === 'Finisher') ? '#ffffff' : '#000000';

      // Day cell — red tint, bold, centered
      var dayCell = sheet.getRange(rowNum, 1);
      dayCell.setBackground(DAY_TINT);
      dayCell.setFontWeight('bold');
      dayCell.setFontColor('#333333');
      dayCell.setHorizontalAlignment('center');
      dayCell.setVerticalAlignment('middle');

      // Block cell — colored background, bold, block-specific text color, centered
      var blockCell = sheet.getRange(rowNum, 2);
      blockCell.setBackground(blockBg);
      blockCell.setFontWeight('bold');
      blockCell.setFontColor(blockFg);
      blockCell.setHorizontalAlignment('center');
      blockCell.setVerticalAlignment('middle');

      // Exercise + remaining 4 columns — match block bg, normal weight
      var bodyRange = sheet.getRange(rowNum, 3, 1, 5);
      bodyRange.setBackground(blockBg);
      bodyRange.setFontWeight('normal');
      bodyRange.setFontColor(bodyFg);
      bodyRange.setVerticalAlignment('middle');

      // Sets column centered
      sheet.getRange(rowNum, 4).setHorizontalAlignment('center');
    }

    // ── Column widths ────────────────────────────────────────────────────────
    sheet.setColumnWidth(1, 110);  // Day
    sheet.setColumnWidth(2, 130);  // Block
    sheet.setColumnWidth(3, 270);  // Exercise
    sheet.setColumnWidth(4,  60);  // Sets
    sheet.setColumnWidth(5, 200);  // Target Reps / Time
    sheet.setColumnWidth(6, 140);  // Exercise Demo Link
    sheet.setColumnWidth(7, 500);  // Coaching Note

    // Freeze title + header rows
    sheet.setFrozenRows(2);

    // Wrap text in Coaching Note column
    sheet.getRange(3, 7, dataRows.length, 1).setWrap(true);

    Logger.log('Created: ' + name + ' (' + dataRows.length + ' rows)');
  }


  // ── HUSBAND CORE P3C1 ── Mon / Wed / Fri ─────────────────────────────────

  buildSheet('Husband Core P3C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Jumping Jacks',              1, '30 sec',               ''],
    ['Monday', 'Warm-Up',   'Hip Circles',                1, '10/direction',          ''],
    ['Monday', 'Warm-Up',   'Glute Bridge',               1, '15 reps',               ''],
    ['Monday', 'Warm-Up',   'Inchworm Walk-Out',          1, '6 reps',                ''],
    ['Monday', 'Block 1',   'Explosive Push-Up',          3, '8 reps',                ''],
    ['Monday', 'Block 1',   'Plyo Push-Up',               3, '6 reps',                ''],
    ['Monday', 'Block 1',   'Burpee to Hollow Body',      3, '5 reps',                ''],
    ['Monday', 'Block 2',   'Mountain Climber Tabata',    3, '20 sec on/10 sec off',  ''],
    ['Monday', 'Block 2',   'Jump Squat to Plank',        3, '8 reps',                ''],
    ['Monday', 'Block 2',   'Sprint Sit-Up',              3, '15 reps',               ''],
    ['Monday', 'Finisher',  'Max Explosive Push-Ups',     1, '30 sec max',            ''],
    ['Monday', 'Cool-Down', "Child's Pose",               1, '45 sec',                ''],
    ['Monday', 'Cool-Down', 'Cobra Stretch',              1, '30 sec',                ''],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',        1, '30 sec/side',           ''],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'High Knees',              1, '30 sec',               ''],
    ['Wednesday', 'Warm-Up',   'Bear Crawl In Place',     1, '30 sec',               ''],
    ['Wednesday', 'Warm-Up',   'Glute Bridge March',      1, '10/side',              ''],
    ['Wednesday', 'Warm-Up',   'Arm Circles',             1, '10 each direction',    ''],
    ['Wednesday', 'Block 1',   'Lateral Bound to Plank',  3, '8/side',               ''],
    ['Wednesday', 'Block 1',   'Clap Push-Up',            3, '6 reps',               ''],
    ['Wednesday', 'Block 1',   'Speed Dead Bug',          3, '10/side fast',         ''],
    ['Wednesday', 'Block 2',   'Plyo Mountain Climber',   3, '20 reps',              ''],
    ['Wednesday', 'Block 2',   'Jump Squat Hold',         3, '8 reps',               ''],
    ['Wednesday', 'Block 2',   'Hollow Body Rock Fast',   3, '30 sec',               ''],
    ['Wednesday', 'Finisher',  'Tabata Burpee',           1, '8 rounds: 20s on/10s off', ''],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',      1, '45 sec/side',          ''],
    ['Wednesday', 'Cool-Down', 'Figure-4 Stretch',        1, '30 sec/side',          ''],
    ['Wednesday', 'Cool-Down', "Child's Pose",            1, '30 sec',               ''],

    // Friday
    ['Friday', 'Warm-Up',   'Jumping Jacks',              1, '30 sec',               ''],
    ['Friday', 'Warm-Up',   'Hip Flexor Lunge Stretch',   1, '30 sec/side',          ''],
    ['Friday', 'Warm-Up',   'High Knees',                 1, '30 sec',               ''],
    ['Friday', 'Warm-Up',   'Inchworm Walk-Out',          1, '6 reps',               ''],
    ['Friday', 'Block 1',   'Explosive Push-Up',          3, '8 reps',               ''],
    ['Friday', 'Block 1',   'Burpee to Hollow Body',      3, '6 reps',               ''],
    ['Friday', 'Block 1',   'Sprint Sit-Up',              3, '15 reps',              ''],
    ['Friday', 'Block 2',   'Plyo Push-Up',               3, '6 reps',               ''],
    ['Friday', 'Block 2',   'Speed Mountain Climber',     3, '30 sec',               ''],
    ['Friday', 'Block 2',   'Jump Squat to Plank',        3, '8 reps',               ''],
    ['Friday', 'Finisher',  'Power Gauntlet',             1, '1 round',              'Clap Push-Up 6 -> Burpee 6 -> Speed Dead Bug 10/side, no rest'],
    ['Friday', 'Cool-Down', "Child's Pose",               1, '45 sec',               ''],
    ['Friday', 'Cool-Down', 'Supine Hamstring Stretch',   1, '30 sec/side',          ''],
    ['Friday', 'Cool-Down', 'Cobra to Upward Dog',        1, '30 sec',               '']

  ]);


  // ── HUSBAND ARMS P3C1 ── Tue / Thu / Sat ─────────────────────────────────

  buildSheet('Husband Arms P3C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Arm Circles Fast',                  1, '15 each direction', ''],
    ['Tuesday', 'Warm-Up',   'Jumping Jacks',                     1, '30 sec',             ''],
    ['Tuesday', 'Warm-Up',   'Wrist Circles',                     1, '10 each direction',  ''],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Green',             1, '15 reps',            ''],
    ['Tuesday', 'Block 1',   'Explosive Table Inverted Row',      3, '8 reps',             ''],
    ['Tuesday', 'Block 1',   'Band Bicep Curl Fast Blue',         3, '15 reps fast',       ''],
    ['Tuesday', 'Block 1',   'Clap Push-Up',                      3, '6 reps',             ''],
    ['Tuesday', 'Block 2',   'Explosive Push-Up',                 3, '10 reps',            ''],
    ['Tuesday', 'Block 2',   'Plyometric Dip on Chair',           3, '8 reps',             ''],
    ['Tuesday', 'Block 2',   'Reactive Push-Up',                  3, '8 reps',             ''],
    ['Tuesday', 'Finisher',  'Push-Up Power Burnout',             1, '30 sec max',         ''],
    ['Tuesday', 'Cool-Down', 'Overhead Tricep Stretch',           1, '30 sec/side',        ''],
    ['Tuesday', 'Cool-Down', 'Cross-Body Bicep Stretch',          1, '30 sec/side',        ''],

    // Thursday
    ['Thursday', 'Warm-Up',   'High Knees',                       1, '30 sec',             ''],
    ['Thursday', 'Warm-Up',   'Arm Circles Fast',                 1, '15 each direction',  ''],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Green',            1, '15 reps',            ''],
    ['Thursday', 'Warm-Up',   'Wrist Circles',                    1, '10 each direction',  ''],
    ['Thursday', 'Block 1',   'Fast Pike Push-Up',                3, '10 reps fast',       ''],
    ['Thursday', 'Block 1',   'Explosive Elevated Pike Push-Up',  3, '6 reps',             ''],
    ['Thursday', 'Block 1',   'Jump and Press Band Blue',         3, '10 reps',            ''],
    ['Thursday', 'Block 2',   'Explosive Table Inverted Row',     3, '8 reps',             ''],
    ['Thursday', 'Block 2',   'Band Face Pull Fast Green',        3, '15 reps fast',       ''],
    ['Thursday', 'Block 2',   'Speed Superman Pulls',             3, '15 reps fast',       ''],
    ['Thursday', 'Finisher',  'Power Shoulder Complex',           1, '1 round',            'Fast Pike Push-Up 10 -> Explosive Row 8 -> Band Face Pull 15, no rest'],
    ['Thursday', 'Cool-Down', 'Doorway Chest Stretch',            1, '30 sec',             ''],
    ['Thursday', 'Cool-Down', "Child's Pose Wide Arms",           1, '45 sec',             ''],

    // Saturday
    ['Saturday', 'Warm-Up',   'Jumping Jacks',                    1, '30 sec',             ''],
    ['Saturday', 'Warm-Up',   'High Knees',                       1, '30 sec',             ''],
    ['Saturday', 'Warm-Up',   'Inchworm Walk-Out',                1, '6 reps',             ''],
    ['Saturday', 'Warm-Up',   'Band Pull-Apart Green',            1, '15 reps',            ''],
    ['Saturday', 'Block 1',   'Clap Push-Up',                     3, '8 reps',             ''],
    ['Saturday', 'Block 1',   'Explosive Table Inverted Row',     3, '8 reps',             ''],
    ['Saturday', 'Block 1',   'Plyometric Dip',                   3, '8 reps',             ''],
    ['Saturday', 'Block 2',   'Speed Push-Up',                    3, '20 reps fast',       ''],
    ['Saturday', 'Block 2',   'Band Curl Fast Blue',              3, '15 reps fast',       ''],
    ['Saturday', 'Block 2',   'Fast Pike Push-Up',                3, '12 reps fast',       ''],
    ['Saturday', 'Finisher',  'Power Triset',                     1, '1 round',            'Clap Push-Up 8 -> Explosive Row 8 -> Plyometric Dip 8, no rest'],
    ['Saturday', 'Cool-Down', 'Overhead Tricep Stretch',          1, '30 sec/side',        ''],
    ['Saturday', 'Cool-Down', 'Doorway Chest Opener',             1, '30 sec',             ''],
    ['Saturday', 'Cool-Down', 'Wrist Flexor Stretch',             1, '20 sec/side',        '']

  ]);


  // ── WIFE CORE P3C1 ── Mon / Wed / Fri ────────────────────────────────────

  buildSheet('Wife Core P3C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Diaphragmatic Breathing',          1, '6 breaths',         ''],
    ['Monday', 'Warm-Up',   'Glute Bridge',                     1, '12 reps',            ''],
    ['Monday', 'Warm-Up',   'Marching In Place',                1, '30 sec',             ''],
    ['Monday', 'Warm-Up',   'Cat-Camel',                        1, '10 reps',            ''],
    ['Monday', 'Block 1',   'Modified Burpee No Jump',          3, '8 reps',             ''],
    ['Monday', 'Block 1',   'Low Impact Mountain Climber Step', 3, '10/side',            ''],
    ['Monday', 'Block 1',   'Speed Bird Dog',                   3, '10/side fast',       ''],
    ['Monday', 'Block 2',   'Standing Speed March High Knees',  3, '30 sec',             ''],
    ['Monday', 'Block 2',   'Small Squat Jump Land Soft',       3, '10 reps',            ''],
    ['Monday', 'Block 2',   'Fast Heel Slide',                  3, '12/side',            ''],
    ['Monday', 'Finisher',  'Speed Circuit',                    1, '1 round',            'Speed March 30s -> Modified Burpee 8 -> Fast Heel Slide 10/side, no rest'],
    ['Monday', 'Cool-Down', "Child's Pose",                     1, '45 sec',             ''],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',              1, '30 sec/side',        ''],
    ['Monday', 'Cool-Down', 'Knees-to-Chest Hug',               1, '30 sec',             ''],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Diaphragmatic Breathing',       1, '6 breaths',          ''],
    ['Wednesday', 'Warm-Up',   'Marching In Place',             1, '30 sec',             ''],
    ['Wednesday', 'Warm-Up',   'Hip Circles Standing',          1, '10/direction',       ''],
    ['Wednesday', 'Warm-Up',   'Glute Bridge',                  1, '12 reps',            ''],
    ['Wednesday', 'Block 1',   'Lateral Step to Plank',         3, '8/side',             ''],
    ['Wednesday', 'Block 1',   'Low Plyo Push-Up Knees',        3, '8 reps',             ''],
    ['Wednesday', 'Block 1',   'Sprint Dead Bug',               3, '10/side fast',       ''],
    ['Wednesday', 'Block 2',   'Speed Standing March',          3, '30 sec',             ''],
    ['Wednesday', 'Block 2',   'Fast Glute Bridge',             3, '15 reps fast',       ''],
    ['Wednesday', 'Block 2',   'Quick Lateral Steps',           3, '30 sec',             ''],
    ['Wednesday', 'Finisher',  'Power Tabata Modified',         1, '8 rounds: Modified Burpee 20s/rest 10s', ''],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',            1, '40 sec/side',        ''],
    ['Wednesday', 'Cool-Down', 'Figure-4 Glute Stretch',        1, '30 sec/side',        ''],
    ['Wednesday', 'Cool-Down', "Child's Pose",                  1, '30 sec',             ''],

    // Friday
    ['Friday', 'Warm-Up',   'Diaphragmatic Breathing',          1, '6 breaths',          ''],
    ['Friday', 'Warm-Up',   'Glute Bridge March',               1, '10/side',            ''],
    ['Friday', 'Warm-Up',   'Marching In Place',                1, '30 sec',             ''],
    ['Friday', 'Warm-Up',   'Cat-Camel',                        1, '10 reps',            ''],
    ['Friday', 'Block 1',   'Small Squat Jump',                 3, '10 reps',            ''],
    ['Friday', 'Block 1',   'Speed Bird Dog',                   3, '10/side',            ''],
    ['Friday', 'Block 1',   'Modified Burpee',                  3, '8 reps',             ''],
    ['Friday', 'Block 2',   'Fast Heel Slide',                  3, '12/side',            ''],
    ['Friday', 'Block 2',   'Low Impact Mountain Climber',      3, '15/side',            ''],
    ['Friday', 'Block 2',   'Speed March',                      3, '30 sec',             ''],
    ['Friday', 'Finisher',  'Power Flow',                       1, '1 round',            'Small Squat Jump 10 -> Speed Bird Dog 10/side -> Low Impact Mountain Climber 15/side, no rest'],
    ['Friday', 'Cool-Down', "Child's Pose",                     1, '1 min',              ''],
    ['Friday', 'Cool-Down', 'Supine Spinal Twist',              1, '40 sec/side',        ''],
    ['Friday', 'Cool-Down', 'Happy Baby Pose',                  1, '30 sec',             '']

  ]);


  // ── WIFE BANDS P3C1 ── Tue / Thu ─────────────────────────────────────────

  buildSheet('Wife Bands P3C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Diaphragmatic Breathing',         1, '6 breaths',          ''],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Green',           1, '12 reps',            ''],
    ['Tuesday', 'Warm-Up',   'Marching In Place',               1, '30 sec',             ''],
    ['Tuesday', 'Warm-Up',   'Shoulder Rolls',                  1, '10 each direction',  ''],
    ['Tuesday', 'Block 1',   'Band Bicep Curl Fast Green',      3, '15 reps fast',       ''],
    ['Tuesday', 'Block 1',   'Band Overhead Press Fast Green',  3, '12 reps fast',       ''],
    ['Tuesday', 'Block 1',   'Band Front Raise Fast Tan',       3, '15 reps fast',       ''],
    ['Tuesday', 'Block 2',   'Band Squat Jump Small Green',     3, '10 reps land soft',  ''],
    ['Tuesday', 'Block 2',   'Band Glute Bridge Fast Blue',     3, '15 reps fast',       ''],
    ['Tuesday', 'Block 2',   'Band Side Step Fast Tan',         3, '20 steps fast',      ''],
    ['Tuesday', 'Finisher',  'Band Power Circuit',              1, '1 round',            'Fast Curl 15 -> Fast Squat 12 -> Fast Glute Bridge 15, no rest'],
    ['Tuesday', 'Cool-Down', "Child's Pose",                    1, '45 sec',             ''],
    ['Tuesday', 'Cool-Down', 'Seated Spinal Twist',             1, '30 sec/side',        ''],

    // Thursday
    ['Thursday', 'Warm-Up',   'Diaphragmatic Breathing',        1, '6 breaths',          ''],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Green',          1, '12 reps',            ''],
    ['Thursday', 'Warm-Up',   'Glute Bridge Fast',              1, '15 reps',            ''],
    ['Thursday', 'Warm-Up',   'Shoulder Rolls',                 1, '10 each direction',  ''],
    ['Thursday', 'Block 1',   'Band Row Fast Green',            3, '15 reps fast',       ''],
    ['Thursday', 'Block 1',   'Band Face Pull Fast Green',      3, '15 reps fast',       ''],
    ['Thursday', 'Block 1',   'Band Reverse Fly Fast Tan',      3, '15 reps fast',       ''],
    ['Thursday', 'Block 2',   'Band Squat Fast Green',          3, '15 reps fast',       ''],
    ['Thursday', 'Block 2',   'Band Donkey Kick Fast Tan',      3, '20/side fast',       ''],
    ['Thursday', 'Block 2',   'Band Kickback Fast Tan',         3, '20/side fast',       ''],
    ['Thursday', 'Finisher',  'Band Power Burnout',             1, '1 round',            'Fast Row 15 -> Fast Squat 15 -> Fast Glute Bridge 15, no rest'],
    ['Thursday', 'Cool-Down', 'Figure-4 Glute Stretch',         1, '30 sec/side',        ''],
    ['Thursday', 'Cool-Down', "Child's Pose",                   1, '30 sec',             '']

  ]);


  SpreadsheetApp.flush();
  Logger.log('Done — all 4 Phase 3 Power tabs created successfully.');
}
