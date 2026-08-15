function createPhase4Sheets() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  // Phase 4 Unilateral - purple (matches CONJUGATE_PHASE_META color #a371f7 in workout.html)
  var PHASE_BG    = '#a371f7';   // purple title banner
  var DAY_TINT    = '#f2e9fe';   // light purple for Day column
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
    'Husband Core P4C1': 'Husband Core Workout - P4C1 (Monday, Wednesday, Friday) - Unilateral Phase',
    'Husband Arms P4C1': 'Husband Arms Workout - P4C1 (Tuesday, Thursday, Saturday) - Unilateral Phase',
    'Wife Core P4C1':    'Wife Core Workout - P4C1 (Monday, Wednesday, Friday) - Unilateral Phase, Breastfeeding-Safe',
    'Wife Bands P4C1':   'Wife Bands Workout - P4C1 (Tuesday, Thursday) - Unilateral Phase, Breastfeeding-Safe'
  };

  // Coaching notes keyed by exact exercise name. Used as fallback whenever a
  // row does not carry its own inline note (6th array element).
  var NOTES = {
    'Cat-Camel Stretch': 'Move slowly between arching and rounding your spine, syncing each phase with a full breath in and out.',
    'Cat-Camel': 'Alternate between arching and rounding your spine on all fours, moving slowly and breathing with each phase.',
    'Dead Bug Arms Only': 'Keep your lower back pressed flat into the floor and lower your arms only as far as you can control without your ribs flaring.',
    'Single-Leg Hip Circle': 'Balance on one leg and trace slow, controlled circles with the other knee, keeping your standing hip level throughout.',
    'Lying Knee Rocks': 'Hug both knees to your chest and rock gently side to side to loosen your lower back before standing work.',
    'Single-Leg Glute Bridge': 'Bridge your hips up one leg at a time, keeping the resting knee bent and your hips square to the ceiling.',
    'Inchworm Walk-Out': 'Walk your hands out to a plank keeping your legs as straight as comfortable, then walk them back and stand tall.',
    'Arm Circles': 'Keep your arms fully extended and make slow, controlled circles in each direction to open up your shoulders.',
    'Wrist Circles': 'Rotate slowly through the full range of each wrist in both directions to prep the joints for load-bearing work.',
    'Single Arm Band Pull-Apart Tan': 'Anchor the tan band and pull one hand back at shoulder height, squeezing your shoulder blade before returning under control.',
    'Diaphragmatic Breathing': 'Place one hand on your belly and breathe deeply into it, keeping your chest and shoulders relaxed on every inhale.',
    'Pelvic Tilts': 'Gently rock your pelvis to flatten and then arch your lower back against the floor, moving slowly with your breath.',
    'Hip Circles Standing': 'Stand tall with hands on hips and trace slow, wide circles with your hips in each direction.',
    'Single Arm Plank': 'Shift your weight onto one hand and keep your hips completely square - do not let them rotate toward the lifted side.',
    'Single Leg Hollow Hold': 'Press your lower back into the floor and extend one leg straight out while keeping the other knee tucked tight to your chest.',
    'Unilateral Dead Bug': 'Extend the opposite arm and leg together, keeping your lower back glued to the floor for the entire rep.',
    'Archer Push-Up': 'Shift your weight to one side as you lower, keeping the opposite arm straight and barely touching the floor.',
    'Single Leg Glute Bridge': 'Drive through the heel of your working leg and squeeze your glute hard at the top, keeping your hips level.',
    'Single Arm Tricep Push-Up': 'Keep your elbow tucked close to your ribs as you lower on one arm, using the other arm only for light balance support.',
    'Lateral Plank Walk': 'Move sideways in a strict push-up plank one hand and foot at a time, keeping your hips level the whole way.',
    'Single Arm Band Pallof Press': 'Anchor the band at chest height and press straight out from your chest, resisting any pull to rotate your torso.',
    'Single Arm Table Inverted Row': 'Grip the table edge with one hand and pull your chest toward it, keeping your body in a straight line.',
    'Single Arm Band Curl Blue': 'Curl with one arm at a time, keeping your elbow pinned to your side and avoiding any body swing.',
    'One Arm Hang on Table Edge': 'Hang from the table edge with one arm, keeping your shoulder engaged and not letting it shrug up toward your ear.',
    'Single Arm Push-Up Progression': 'Lower with one hand doing most of the work while the other rests lightly behind your back, going only as deep as you control.',
    'Single Arm Tricep Dip': 'Support most of your weight on one arm as you dip, keeping your elbow tracking straight back, not flared out.',
    'Archer Row': 'Row to one side at a time, pulling your elbow back hard while the opposite arm stays extended and relaxed.',
    'Single Arm Pike Push-Up': 'Keep your hips high in a pike position and lower toward one hand, driving through your shoulder to press back up.',
    'Single Arm Wall Push': 'Press against the wall with one hand, keeping your body rigid and your core braced throughout the rep.',
    'Single Arm Band Overhead Press Blue': 'Press one arm straight overhead from shoulder height, keeping your ribs down and avoiding any lean to the side.',
    'Single Arm Band Face Pull Green': 'Pull the band toward your face with one arm, driving your elbow high and squeezing your shoulder blade back.',
    'Single Arm Table Row': 'Row with one arm under the table edge, keeping your torso still and pulling with your elbow, not your hand.',
    'Single Arm Bird Dog': 'Extend the opposite arm forward only, keeping your hips and shoulders square to the floor throughout.',
    'Single Leg Balance Hold': 'Stand tall on one leg with a soft knee, keeping your hips level and your gaze fixed on a point ahead.',
    'Single Arm Modified Plank': 'Hold a plank on your knees with one hand lifted, keeping your hips square and avoiding any rotation.',
    'Unilateral Heel Slide': 'Slide one heel out along the floor while keeping your lower back pressed down, then return under control.',
    'Modified Archer Push-Up': 'From your knees, shift your weight to one side as you lower, keeping the other arm straight for support only.',
    'Single Leg Calf Raise Hold': 'Rise onto the ball of one foot and hold, keeping your ankle stable and your standing leg fully engaged.',
    'Lateral Lunge': "Step wide to one side and sink your hips back and down, keeping the straight leg's foot flat on the floor.",
    'Single Arm Band Curl Green': 'Curl one arm at a time against the band, keeping your elbow fixed at your side throughout the movement.',
    'Single Arm Band Overhead Press Tan': 'Press one arm overhead from the band anchor, bracing your core to avoid leaning to either side.',
    'Single Arm Band Front Raise Tan': 'Raise one arm straight in front to shoulder height against the band, keeping the movement slow and controlled.',
    'Single Leg Band Glute Bridge Green': 'Bridge up on one leg with the band around your hips, driving through your heel and squeezing at the top.',
    'Single Leg Band Kickback Tan': 'Kick one leg straight back against the band tension, squeezing your glute hard at the top of the movement.',
    'Single Leg Band Clamshell Tan': 'Lying on your side with the band above your knees, open your top knee while keeping your feet together.',
    'Single Arm Band Row Green': 'Row one arm at a time against the band, driving your elbow back and squeezing your shoulder blade.',
    'Single Arm Band Face Pull Tan': 'Pull the band toward your face with one arm, leading with your elbow high and squeezing your upper back.',
    'Single Arm Band Reverse Fly Tan': 'With a slight bend in your elbow, sweep one arm out to the side against the band, squeezing your shoulder blade.',
    'Single Leg Band Romanian Deadlift Tan': 'Hinge at your hip over one leg while the band resists, keeping your back flat and reaching toward the floor.',
    'Single Leg Band Squat Tan': 'Squat down on one leg with the band providing resistance, keeping your knee tracking over your toes.',
    'Single Leg Band Side Step Tan': 'Step sideways against the band tension on one leg, keeping tension on the band the entire step.',
    'Max Archer Push-Ups': 'Perform as many controlled reps as possible shifting your weight fully to each side, resting only as needed between sides.',
    'Max Single Arm Push-Up': 'Perform as many controlled single-arm push-ups as possible per side, resetting your plank position between reps if needed.',
    'Max Single Arm Modified Plank': 'Hold each side as long as possible with strict form, stopping the moment your hips begin to rotate.',
    'Max Single Leg Band Glute Bridge': 'Perform as many controlled reps as possible per side against the band, keeping your hips level throughout.',
    "Child's Pose": 'Sink your hips back toward your heels and reach your arms forward, breathing deeply and letting your back release.',
    'Cobra Stretch': 'Lie face down and press your hands into the floor under your shoulders, straightening your arms slowly while keeping your hips down.',
    'Seated Spinal Twist': 'Sit tall and rotate your torso, using a gentle hand press on your knee to deepen the twist while breathing slowly.',
    'Lying Spinal Twist': 'Lie on your back and let your knees fall gently to one side, breathing deeply and letting gravity do the work.',
    'Figure-4 Stretch': 'Cross one ankle over the opposite knee and gently pull the bottom leg toward you, breathing steadily to deepen the stretch.',
    'Supine Hamstring Stretch': 'Lie on your back and gently pull one straight leg toward you, keeping the other leg relaxed on the floor.',
    'Cobra to Upward Dog': 'Flow from a low cobra into a full upward dog, pressing through your hands and lifting your chest with each breath.',
    'Doorway Chest Stretch': 'Place your forearm on the doorframe and gently lean forward until you feel a stretch across your chest.',
    "Child's Pose Wide Arms": 'Sink your hips back toward your heels with your arms spread wide, breathing deeply into your upper back and shoulders.',
    'Overhead Tricep Stretch': 'Reach one arm overhead and bend your elbow, gently pulling it back with your other hand to deepen the stretch.',
    'Cross-Body Bicep Stretch': 'Extend one arm across your body and gently pull it toward your chest with your other arm, feeling the stretch through your bicep.',
    'Doorway Chest Opener': 'Place both forearms on the doorframe and lean gently forward, opening across your chest and shoulders.',
    'Wrist Flexor Stretch': 'Extend one arm forward with your palm up and gently pull your fingers back with your other hand.',
    'Knees-to-Chest Hug': 'Pull both knees gently toward your chest and hold, letting your lower back relax fully into the floor.',
    'Figure-4 Glute Stretch': 'Cross one ankle over the opposite knee and gently press the knee away from you, feeling the stretch deep in your glute.',
    'Supine Spinal Twist': 'Lie on your back and let your knees fall gently to one side while keeping both shoulders on the floor.',
    'Happy Baby Pose': 'Lie on your back and hold the outside of your feet, gently rocking side to side to release your lower back and hips.'
  };

  function buildSheet(name, exercises) {
    var existing = ss.getSheetByName(name);
    if (existing) ss.deleteSheet(existing);
    var sheet = ss.insertSheet(name);

    // -- ROW 1: Title banner ---------------------------------------------------
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

    // -- ROW 2: Column headers --------------------------------------------------
    sheet.setRowHeight(2, 28);
    var hdrRange = sheet.getRange(2, 1, 1, 7);
    hdrRange.setValues([['Day', 'Block', 'Exercise', 'Sets', 'Target Reps / Time', 'Exercise Demo Link', 'Coaching Note']]);
    hdrRange.setBackground(HEADER_BG);
    hdrRange.setFontColor('#ffffff');
    hdrRange.setFontWeight('bold');
    hdrRange.setFontSize(11);
    hdrRange.setHorizontalAlignment('center');
    hdrRange.setVerticalAlignment('middle');

    // -- ROWS 3+: Exercise data --------------------------------------------------
    var dataRows = [];
    var missingNotes = [];
    for (var i = 0; i < exercises.length; i++) {
      var e = exercises[i];
      // e = [day, block, exercise, sets, target, inlineNote(optional)]
      var note = e[5] || NOTES[e[2]] || '';
      if (!note) missingNotes.push(name + ' / ' + e[0] + ' / ' + e[2]);
      dataRows.push([e[0], e[1], e[2], e[3], e[4], '', note]);
    }
    if (missingNotes.length > 0) {
      throw new Error('Missing coaching note(s): ' + missingNotes.join('; '));
    }
    if (dataRows.length === 0) return;

    sheet.getRange(3, 1, dataRows.length, 7).setValues(dataRows);

    for (var r = 0; r < dataRows.length; r++) {
      var rowNum  = r + 3;
      var block   = dataRows[r][1];
      var blockBg = BLOCK_BG[block] || '#ffffff';
      var blockFg = BLOCK_FG[block] || '#000000';
      var bodyFg  = (block === 'Finisher') ? '#ffffff' : '#000000';

      // Day cell - purple tint, bold, centered
      var dayCell = sheet.getRange(rowNum, 1);
      dayCell.setBackground(DAY_TINT);
      dayCell.setFontWeight('bold');
      dayCell.setFontColor('#333333');
      dayCell.setHorizontalAlignment('center');
      dayCell.setVerticalAlignment('middle');

      // Block cell - colored background, bold, block-specific text color, centered
      var blockCell = sheet.getRange(rowNum, 2);
      blockCell.setBackground(blockBg);
      blockCell.setFontWeight('bold');
      blockCell.setFontColor(blockFg);
      blockCell.setHorizontalAlignment('center');
      blockCell.setVerticalAlignment('middle');

      // Exercise + remaining 4 columns - match block bg, normal weight
      var bodyRange = sheet.getRange(rowNum, 3, 1, 5);
      bodyRange.setBackground(blockBg);
      bodyRange.setFontWeight('normal');
      bodyRange.setFontColor(bodyFg);
      bodyRange.setVerticalAlignment('middle');

      // Sets column centered
      sheet.getRange(rowNum, 4).setHorizontalAlignment('center');
    }

    // -- Column widths ------------------------------------------------------------
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


  // == HUSBAND CORE P4C1 == Mon / Wed / Fri ====================================

  buildSheet('Husband Core P4C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Cat-Camel Stretch',          1, '10 reps'],
    ['Monday', 'Warm-Up',   'Dead Bug Arms Only',         1, '10/side'],
    ['Monday', 'Warm-Up',   'Single-Leg Hip Circle',      1, '10/direction each leg'],
    ['Monday', 'Block 1',   'Single Arm Plank',           3, '20 sec/side'],
    ['Monday', 'Block 1',   'Single Leg Hollow Hold',     3, '25 sec/side'],
    ['Monday', 'Block 1',   'Unilateral Dead Bug',        3, '8/side'],
    ['Monday', 'Block 2',   'Archer Push-Up',             3, '6/side'],
    ['Monday', 'Block 2',   'Single Leg Glute Bridge',    3, '12/side'],
    ['Monday', 'Block 2',   'Half-Kneeling Band Chop',    3, '10/side', 'Kneel on one knee with the green band anchored high. Pull the band diagonally down across your body to the opposite hip. Resist rotation throughout - your torso should stay square.'],
    ['Monday', 'Finisher',  'Max Single Arm Plank Hold',  1, 'To failure each side', 'In a push-up plank position lift one hand off the floor and extend it forward. Hold as long as possible without letting your hips rotate or drop. Switch sides immediately and repeat. Log your combined total time.'],
    ['Monday', 'Cool-Down', "Child's Pose",               1, '45 sec'],
    ['Monday', 'Cool-Down', 'Cobra Stretch',              1, '30 sec'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',        1, '30 sec/side'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Lying Knee Rocks',              1, '12 reps'],
    ['Wednesday', 'Warm-Up',   'Single-Leg Glute Bridge',       1, '10/side'],
    ['Wednesday', 'Warm-Up',   'Inchworm Walk-Out',             1, '6 reps'],
    ['Wednesday', 'Block 1',   'Single Arm Tricep Push-Up',     3, '8/side'],
    ['Wednesday', 'Block 1',   'Lateral Plank Walk',            3, '10 steps each direction'],
    ['Wednesday', 'Block 1',   'Archer Push-Up',                3, '8/side'],
    ['Wednesday', 'Block 2',   'Unilateral Dead Bug',           3, '10/side'],
    ['Wednesday', 'Block 2',   'Single Arm Plank',              3, '25 sec/side'],
    ['Wednesday', 'Block 2',   'Single Arm Band Pallof Press',  3, '10/side'],
    ['Wednesday', 'Finisher',  'Unilateral Gauntlet',           1, '1 round', 'Single Arm Plank 20s/side -> Archer Push-Up 5/side -> Single Arm Band Pallof Press 10/side, no rest'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',            1, '45 sec/side'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Stretch',              1, '30 sec/side'],
    ['Wednesday', 'Cool-Down', "Child's Pose",                  1, '30 sec'],

    // Friday
    ['Friday', 'Warm-Up',   'Cat-Camel Stretch',             1, '10 reps'],
    ['Friday', 'Warm-Up',   'Single-Leg Hip Circle',         1, '10/direction each leg'],
    ['Friday', 'Warm-Up',   'Arm Circles',                   1, '10 each direction'],
    ['Friday', 'Block 1',   'Archer Push-Up',                3, '8/side'],
    ['Friday', 'Block 1',   'Single Leg Hollow Hold',        3, '30 sec/side'],
    ['Friday', 'Block 1',   'Single Arm Band Pallof Press',  3, '10/side', 'Anchor the blue band at chest height and kneel or stand sideways to it. Press one hand straight out from your chest and hold 2 seconds - resist any rotation. Switch sides and repeat.'],
    ['Friday', 'Block 2',   'Single Leg Glute Bridge',       3, '15/side'],
    ['Friday', 'Block 2',   'Lateral Plank Walk',            3, '12 steps each direction'],
    ['Friday', 'Block 2',   'Single Arm Tricep Push-Up',     3, '10/side'],
    ['Friday', 'Finisher',  'Max Archer Push-Ups',           1, '1 set to failure/side'],
    ['Friday', 'Cool-Down', "Child's Pose",                  1, '45 sec'],
    ['Friday', 'Cool-Down', 'Supine Hamstring Stretch',      1, '30 sec/side'],
    ['Friday', 'Cool-Down', 'Cobra to Upward Dog',           1, '30 sec']

  ]);


  // == HUSBAND ARMS P4C1 == Tue / Thu / Sat =====================================

  buildSheet('Husband Arms P4C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Arm Circles',                       1, '15 each direction'],
    ['Tuesday', 'Warm-Up',   'Single Arm Band Pull-Apart Tan',    1, '15/side'],
    ['Tuesday', 'Warm-Up',   'Wrist Circles',                     1, '10 each direction'],
    ['Tuesday', 'Block 1',   'Single Arm Table Inverted Row',     3, '8/side'],
    ['Tuesday', 'Block 1',   'Single Arm Band Curl Blue',         3, '12/side'],
    ['Tuesday', 'Block 1',   'One Arm Hang on Table Edge',        3, '20 sec/side'],
    ['Tuesday', 'Block 2',   'Single Arm Push-Up Progression',    3, '5/side'],
    ['Tuesday', 'Block 2',   'Single Arm Tricep Dip',             3, '8/side'],
    ['Tuesday', 'Block 2',   'Archer Row',                        3, '8/side'],
    ['Tuesday', 'Finisher',  'Max Single Arm Push-Up',            1, '1 set to failure/side'],
    ['Tuesday', 'Cool-Down', 'Overhead Tricep Stretch',           1, '30 sec/side'],
    ['Tuesday', 'Cool-Down', 'Cross-Body Bicep Stretch',          1, '30 sec/side'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Arm Circles',                        1, '15 each direction'],
    ['Thursday', 'Warm-Up',   'Single Arm Band Pull-Apart Tan',     1, '15/side'],
    ['Thursday', 'Warm-Up',   'Wrist Circles',                      1, '10 each direction'],
    ['Thursday', 'Block 1',   'Single Arm Pike Push-Up',            3, '6/side'],
    ['Thursday', 'Block 1',   'Single Arm Wall Push',               3, '8/side'],
    ['Thursday', 'Block 1',   'Single Arm Band Overhead Press Blue',3, '12/side'],
    ['Thursday', 'Block 2',   'Archer Row',                         3, '10/side'],
    ['Thursday', 'Block 2',   'Single Arm Band Face Pull Green',    3, '15/side'],
    ['Thursday', 'Block 2',   'Single Arm Table Row',               3, '10/side'],
    ['Thursday', 'Finisher',  'Unilateral Shoulder Gauntlet',       1, '1 round', 'Single Arm Pike Push-Up 5/side -> Single Arm Band Press 10/side -> Archer Row 8/side, no rest'],
    ['Thursday', 'Cool-Down', 'Doorway Chest Stretch',              1, '30 sec'],
    ['Thursday', 'Cool-Down', "Child's Pose Wide Arms",             1, '45 sec'],

    // Saturday
    ['Saturday', 'Warm-Up',   'Single Arm Band Pull-Apart Tan',     1, '15/side'],
    ['Saturday', 'Warm-Up',   'Inchworm Walk-Out',                  1, '6 reps'],
    ['Saturday', 'Warm-Up',   'Wrist Circles',                      1, '10 each direction'],
    ['Saturday', 'Block 1',   'Single Arm Push-Up Progression',     3, '6/side'],
    ['Saturday', 'Block 1',   'Single Arm Table Inverted Row',      3, '10/side'],
    ['Saturday', 'Block 1',   'Archer Push-Up',                     3, '8/side'],
    ['Saturday', 'Block 2',   'Single Arm Tricep Dip',              3, '10/side'],
    ['Saturday', 'Block 2',   'Single Arm Band Curl Blue',          3, '12/side'],
    ['Saturday', 'Block 2',   'Single Arm Pike Push-Up',            3, '6/side'],
    ['Saturday', 'Finisher',  'Unilateral Triset',                  1, '1 round', 'Single Arm Push-Up 5/side -> Archer Row 8/side -> Single Arm Pike Push-Up 5/side, no rest'],
    ['Saturday', 'Cool-Down', 'Overhead Tricep Stretch',            1, '30 sec/side'],
    ['Saturday', 'Cool-Down', 'Doorway Chest Opener',               1, '30 sec'],
    ['Saturday', 'Cool-Down', 'Wrist Flexor Stretch',               1, '20 sec/side']

  ]);


  // == WIFE CORE P4C1 == Mon / Wed / Fri ========================================

  buildSheet('Wife Core P4C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Diaphragmatic Breathing',       1, '8 breaths'],
    ['Monday', 'Warm-Up',   'Single-Leg Glute Bridge',       1, '10/side'],
    ['Monday', 'Warm-Up',   'Pelvic Tilts',                  1, '12 reps'],
    ['Monday', 'Block 1',   'Single Leg Glute Bridge',       3, '12/side'],
    ['Monday', 'Block 1',   'Unilateral Dead Bug',           3, '8/side'],
    ['Monday', 'Block 1',   'Single Arm Bird Dog',           3, '10/side'],
    ['Monday', 'Block 2',   'Single Leg Balance Hold',       3, '30 sec/side'],
    ['Monday', 'Block 2',   'Single Arm Modified Plank',     3, '20 sec/side'],
    ['Monday', 'Block 2',   'Unilateral Heel Slide',         3, '10/side'],
    ['Monday', 'Finisher',  'Max Single Leg Glute Bridge',   1, '1 set to failure/side', 'Breathe throughout'],
    ['Monday', 'Cool-Down', "Child's Pose",                  1, '45 sec'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',           1, '30 sec/side'],
    ['Monday', 'Cool-Down', 'Knees-to-Chest Hug',            1, '30 sec'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Diaphragmatic Breathing',    1, '6 breaths'],
    ['Wednesday', 'Warm-Up',   'Single-Leg Glute Bridge',    1, '10/side'],
    ['Wednesday', 'Warm-Up',   'Cat-Camel',                  1, '10 reps'],
    ['Wednesday', 'Block 1',   'Modified Archer Push-Up',    3, '6/side'],
    ['Wednesday', 'Block 1',   'Single Leg Calf Raise Hold', 3, '30 sec/side'],
    ['Wednesday', 'Block 1',   'Lateral Lunge',              3, '10/side'],
    ['Wednesday', 'Block 2',   'Unilateral Dead Bug',        3, '10/side'],
    ['Wednesday', 'Block 2',   'Single Arm Bird Dog',        3, '12/side'],
    ['Wednesday', 'Block 2',   'Single Leg Balance Hold',    3, '35 sec/side'],
    ['Wednesday', 'Finisher',  'Unilateral Flow',            1, '1 round', 'Single Leg Glute Bridge 12/side -> Single Arm Bird Dog 10/side -> Lateral Lunge 8/side, no rest'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',         1, '40 sec/side'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Glute Stretch',     1, '30 sec/side'],
    ['Wednesday', 'Cool-Down', "Child's Pose",               1, '30 sec'],

    // Friday
    ['Friday', 'Warm-Up',   'Diaphragmatic Breathing',           1, '8 breaths'],
    ['Friday', 'Warm-Up',   'Single-Leg Glute Bridge',           1, '10/side'],
    ['Friday', 'Warm-Up',   'Pelvic Tilts',                      1, '12 reps'],
    ['Friday', 'Block 1',   'Single Leg Glute Bridge',           3, '15/side'],
    ['Friday', 'Block 1',   'Modified Archer Push-Up',           3, '8/side'],
    ['Friday', 'Block 1',   'Single Arm Band Pallof Press Tan',  3, '12/side', 'Anchor the tan band at chest height and stand sideways to it. Press one hand straight out and hold 2 seconds - resist any pull to rotate. Breathe steadily and never hold your breath.'],
    ['Friday', 'Block 2',   'Single Arm Modified Plank',         3, '25 sec/side'],
    ['Friday', 'Block 2',   'Lateral Lunge',                     3, '12/side'],
    ['Friday', 'Block 2',   'Unilateral Heel Slide',             3, '12/side'],
    ['Friday', 'Finisher',  'Max Single Arm Modified Plank',     1, '1 set to failure/side'],
    ['Friday', 'Cool-Down', "Child's Pose",                      1, '1 min'],
    ['Friday', 'Cool-Down', 'Supine Spinal Twist',               1, '40 sec/side'],
    ['Friday', 'Cool-Down', 'Happy Baby Pose',                   1, '30 sec']

  ]);


  // == WIFE BANDS P4C1 == Tue / Thu =============================================

  buildSheet('Wife Bands P4C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Diaphragmatic Breathing',            1, '6 breaths'],
    ['Tuesday', 'Warm-Up',   'Single Arm Band Pull-Apart Tan',     1, '12/side'],
    ['Tuesday', 'Warm-Up',   'Hip Circles Standing',               1, '10/direction'],
    ['Tuesday', 'Block 1',   'Single Arm Band Curl Green',         3, '12/side'],
    ['Tuesday', 'Block 1',   'Single Arm Band Overhead Press Tan', 3, '12/side'],
    ['Tuesday', 'Block 1',   'Single Arm Band Front Raise Tan',    3, '12/side'],
    ['Tuesday', 'Block 2',   'Single Leg Band Glute Bridge Green', 3, '12/side'],
    ['Tuesday', 'Block 2',   'Single Leg Band Kickback Tan',       3, '15/side'],
    ['Tuesday', 'Block 2',   'Single Leg Band Clamshell Tan',      3, '15/side'],
    ['Tuesday', 'Finisher',  'Max Single Leg Band Glute Bridge',   1, '1 set to failure/side'],
    ['Tuesday', 'Cool-Down', "Child's Pose",                       1, '45 sec'],
    ['Tuesday', 'Cool-Down', 'Seated Spinal Twist',                1, '30 sec/side'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Diaphragmatic Breathing',                1, '6 breaths'],
    ['Thursday', 'Warm-Up',   'Single Arm Band Pull-Apart Tan',         1, '12/side'],
    ['Thursday', 'Warm-Up',   'Single Leg Glute Bridge',                1, '10/side'],
    ['Thursday', 'Block 1',   'Single Arm Band Row Green',              3, '12/side'],
    ['Thursday', 'Block 1',   'Single Arm Band Face Pull Tan',          3, '15/side'],
    ['Thursday', 'Block 1',   'Single Arm Band Reverse Fly Tan',        3, '12/side'],
    ['Thursday', 'Block 2',   'Single Leg Band Romanian Deadlift Tan',  3, '10/side'],
    ['Thursday', 'Block 2',   'Single Leg Band Squat Tan',              3, '10/side'],
    ['Thursday', 'Block 2',   'Single Leg Band Side Step Tan',          3, '12/side'],
    ['Thursday', 'Finisher',  'Unilateral Band Circuit',                1, '1 round', 'Single Arm Row 10/side -> Single Leg RDL 10/side -> Single Arm Curl 10/side, no rest'],
    ['Thursday', 'Cool-Down', 'Figure-4 Glute Stretch',                 1, '30 sec/side'],
    ['Thursday', 'Cool-Down', "Child's Pose",                           1, '30 sec']

  ]);


  SpreadsheetApp.flush();
  Logger.log('Done - all 4 Phase 4 Unilateral tabs created successfully.');
}
