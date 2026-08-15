function createPhase5Sheets() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  // Phase 5 Culmination - gold (matches CONJUGATE_PHASE_META color #d4af37 in workout.html)
  var PHASE_BG    = '#d4af37';   // gold title banner
  var DAY_TINT    = '#faf1d3';   // light gold for Day column
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
    'Husband Core P5C1': 'Husband Core Workout - P5C1 (Monday, Wednesday, Friday) - Culmination Phase',
    'Husband Arms P5C1': 'Husband Arms Workout - P5C1 (Tuesday, Thursday, Saturday) - Culmination Phase',
    'Wife Core P5C1':    'Wife Core Workout - P5C1 (Monday, Wednesday, Friday) - Culmination Phase, Breastfeeding-Safe',
    'Wife Bands P5C1':   'Wife Bands Workout - P5C1 (Tuesday, Thursday) - Culmination Phase, Breastfeeding-Safe'
  };

  // Coaching notes keyed by exact exercise name. Used as fallback whenever a
  // row does not carry its own inline note (6th array element).
  var NOTES = {
    'Cat-Camel Stretch': 'Move slowly between arching and rounding your spine, syncing each phase with a full breath in and out.',
    'Cat-Camel': 'Alternate between arching and rounding your spine on all fours, moving slowly and breathing with each phase.',
    'Dead Bug Arms Only': 'Keep your lower back pressed flat into the floor and lower your arms only as far as you can control without your ribs flaring.',
    'Inchworm Walk-Out': 'Walk your hands out to a plank keeping your legs as straight as comfortable, then walk them back and stand tall.',
    'Lying Knee Rocks': 'Hug both knees to your chest and rock gently side to side to loosen your lower back before standing work.',
    'Single-Leg Glute Bridge': 'Bridge your hips up one leg at a time, keeping the resting knee bent and your hips square to the ceiling.',
    'Arm Circles': 'Keep your arms fully extended and make slow, controlled circles in each direction to open up your shoulders.',
    'Arm Circles Fast': 'Keep your arms fully extended and pump quick, controlled circles in each direction to raise your heart rate.',
    'Wrist Circles': 'Rotate slowly through the full range of each wrist in both directions to prep the joints for load-bearing work.',
    'Bear Crawl In Place': 'Stay low on hands and toes with your knees hovering just off the floor, moving opposite hand and foot together.',
    'Diaphragmatic Breathing': 'Place one hand on your belly and breathe deeply into it, keeping your chest and shoulders relaxed on every inhale.',
    'Pelvic Tilts': 'Gently rock your pelvis to flatten and then arch your lower back against the floor, moving slowly with your breath.',
    'Hip Circles Standing': 'Stand tall with hands on hips and trace slow, wide circles with your hips in each direction.',
    'Band Pull-Apart Blue': 'Hold the band at chest height with your arms extended and pull it apart by squeezing your shoulder blades together.',
    'RKC Plank': 'Squeeze every muscle from your fists to your toes as hard as possible while holding a rigid, dead-straight plank.',
    'Copenhagen Plank': 'Prop your top foot on a chair or bench and hold a side plank, squeezing your inner thigh to keep your hips lifted.',
    'L-Sit Hold on Chairs': 'Support yourself between two chairs and lift your legs straight out, keeping your shoulders pressed down away from your ears.',
    'Explosive Push-Up': 'Drive your hands hard into the floor and press so fast that your hands leave the ground, landing with soft elbows to reset.',
    'Single Arm Plank': 'Shift your weight onto one hand and keep your hips completely square - do not let them rotate toward the lifted side.',
    'Archer Push-Up': 'Shift your weight to one side as you lower, keeping the opposite arm straight and barely touching the floor.',
    'Hollow Body Hold': 'Press your lower back into the floor and lift your shoulders and legs slightly, keeping your ribs down for the entire hold.',
    'Dragon Flag Negative': 'Lower your body as a single rigid unit from your shoulders, resisting the urge to let your hips sag or bend.',
    'Stir the Pot': 'Hold a forearm plank on a stability surface and trace slow circles with your torso, keeping your hips as still as possible.',
    'Mountain Climber Tabata': 'Drive each knee toward your chest as fast as possible during work intervals, keeping your hips level throughout.',
    'Unilateral Dead Bug': 'Extend the opposite arm and leg together, keeping your lower back glued to the floor for the entire rep.',
    'Dead Bug Full': 'Lower the opposite arm and leg together while keeping your lower back pinned to the floor, then return and switch sides.',
    'Band Pallof Press Blue': 'Anchor the band at chest height and press straight out, bracing your core to resist any pull to rotate.',
    'Diamond Push-Up': 'Form a diamond shape with your hands under your chest and lower with your elbows tracking back, not flared out.',
    'Single Leg Glute Bridge': 'Drive through the heel of your working leg and squeeze your glute hard at the top, keeping your hips level.',
    'Slow Negative Table Inverted Row': 'Pull your chest to the table edge, then lower yourself back down as slowly as possible over several seconds.',
    'Band Curl Blue': 'Curl the band up with both arms together, keeping your elbows pinned to your sides throughout the movement.',
    'Clap Push-Up': 'Press explosively enough that both hands leave the floor to clap, landing with soft elbows and resetting fully before the next rep.',
    'Single Arm Table Inverted Row': 'Grip the table edge with one hand and pull your chest toward it, keeping your body in a straight line.',
    'Pike Push-Up': 'Keep your hips high in an inverted-V position and lower the crown of your head toward the floor between your hands.',
    'Pseudo Planche Lean': 'Lean your shoulders forward over your turned-in hands, keeping your arms straight and your core braced the whole hold.',
    'Band Face Pull Green': 'Pull the band toward your face with both hands, driving your elbows high and squeezing your shoulder blades together.',
    'Explosive Pike Push-Up': 'From the pike position, press explosively enough to feel light in your hands, landing softly to reset for the next rep.',
    'Single Arm Pike Push-Up': 'Keep your hips high in a pike position and lower toward one hand, driving through your shoulder to press back up.',
    'Archer Row': 'Row to one side at a time, pulling your elbow back hard while the opposite arm stays extended and relaxed.',
    'Typewriter Push-Up': 'Lower into a wide push-up and shift your chest side to side between your hands before pressing back up.',
    'Explosive Table Inverted Row': 'Pull explosively enough that your chest nearly leaves the table edge, then lower back down under control.',
    'Single Arm Push-Up Progression': 'Lower with one hand doing most of the work while the other rests lightly behind your back, going only as deep as you control.',
    'Band Curl to Press Blue': 'Curl the band up to your shoulders, then press straight overhead in one smooth continuous motion.',
    'Glute Bridge': 'Drive through both heels and squeeze your glutes hard at the top, holding briefly before lowering under control.',
    'Bird Dog': 'Extend opposite arm and leg together, keeping your hips and shoulders square and your lower back from arching.',
    'Bear Hold': 'Hold your knees just off the floor in a tabletop position, keeping your back flat and your core braced throughout.',
    'Modified Copenhagen Plank': 'Prop your top knee on a chair or bench and hold a side plank, keeping your hips lifted and square.',
    'Speed Bird Dog': 'Move through the opposite arm and leg extension quickly while keeping your hips level and your lower back from arching.',
    'Low Impact Mountain Climber': 'Step your knees toward your chest one at a time in a plank position, keeping your hips low and steady.',
    'Forearm Plank': 'Stack your elbows under your shoulders and hold a straight line from head to heels, squeezing your glutes and core.',
    'Hollow Body Hold Bent Knees': 'Press your lower back into the floor and lift your shoulders with your knees bent toward your chest, keeping your ribs down.',
    'Band Pallof Press Tan': 'Anchor the band at chest height and press straight out with both hands, resisting any pull to rotate your torso.',
    'Modified Burpee': 'Step back into a plank instead of jumping, then step back in and rise to standing, keeping the pace steady.',
    'Single Arm Bird Dog': 'Extend the opposite arm forward only, keeping your hips and shoulders square to the floor throughout.',
    'Lateral Lunge': "Step wide to one side and sink your hips back and down, keeping the straight leg's foot flat on the floor.",
    'Small Squat Jump': 'Perform a shallow squat and a small hop, landing softly through your feet with your knees slightly bent.',
    'Band Bicep Curl Green': 'Curl the band up with both arms together, keeping your elbows still and close to your sides throughout.',
    'Band Overhead Press Green': 'Press the band straight overhead from shoulder height, bracing your core to avoid arching your lower back.',
    'Band Glute Bridge Blue': 'Bridge your hips up with the band around your thighs, driving your knees out slightly against the band tension.',
    'Band Bicep Curl Fast Green': 'Curl the band up at a quick, controlled tempo, keeping your elbows fixed at your sides throughout.',
    'Single Arm Band Curl Green': 'Curl one arm at a time against the band, keeping your elbow fixed at your side throughout the movement.',
    'Single Leg Band Glute Bridge Green': 'Bridge up on one leg with the band around your hips, driving through your heel and squeezing at the top.',
    'Band Row Green': 'Row the band toward your ribs with both hands, driving your elbows back and squeezing your shoulder blades together.',
    'Band Squat Blue': 'Stand on the band with feet shoulder-width apart and squat down, keeping your knees tracking over your toes.',
    'Band Squat Fast Green': 'Squat down and drive back up at a quick, controlled tempo, keeping your chest up and knees tracking over your toes.',
    'Single Arm Band Row Green': 'Row one arm at a time against the band, driving your elbow back and squeezing your shoulder blade.',
    'Single Leg Band Romanian Deadlift Tan': 'Hinge at your hip over one leg while the band resists, keeping your back flat and reaching toward the floor.',
    "Child's Pose": 'Sink your hips back toward your heels and reach your arms forward, breathing deeply and letting your back release.',
    'Supine Hamstring Stretch': 'Lie on your back and gently pull one straight leg toward you, keeping the other leg relaxed on the floor.',
    'Seated Spinal Twist': 'Sit tall and rotate your torso, using a gentle hand press on your knee to deepen the twist while breathing slowly.',
    'Lying Spinal Twist': 'Lie on your back and let your knees fall gently to one side, breathing deeply and letting gravity do the work.',
    'Figure-4 Stretch': 'Cross one ankle over the opposite knee and gently pull the bottom leg toward you, breathing steadily to deepen the stretch.',
    'Cobra to Upward Dog': 'Flow from a low cobra into a full upward dog, pressing through your hands and lifting your chest with each breath.',
    'Overhead Tricep Stretch': 'Reach one arm overhead and bend your elbow, gently pulling it back with your other hand to deepen the stretch.',
    'Cross-Body Bicep Stretch': 'Extend one arm across your body and gently pull it toward your chest with your other arm, feeling the stretch through your bicep.',
    'Doorway Chest Stretch': 'Place your forearm on the doorframe and gently lean forward until you feel a stretch across your chest.',
    "Child's Pose Wide Arms": 'Sink your hips back toward your heels with your arms spread wide, breathing deeply into your upper back and shoulders.',
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

      // Day cell - gold tint, bold, centered
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


  // == HUSBAND CORE P5C1 == Mon / Wed / Fri ====================================

  buildSheet('Husband Core P5C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Cat-Camel Stretch',          1, '10 reps'],
    ['Monday', 'Warm-Up',   'Dead Bug Arms Only',         1, '10/side'],
    ['Monday', 'Warm-Up',   'Inchworm Walk-Out',          1, '6 reps'],
    ['Monday', 'Block 1',   'RKC Plank',                  3, '45 sec'],
    ['Monday', 'Block 1',   'Copenhagen Plank',           3, '20 sec/side'],
    ['Monday', 'Block 1',   'L-Sit Hold on Chairs',       3, '20 sec'],
    ['Monday', 'Block 2',   'Explosive Push-Up',          3, '8 reps'],
    ['Monday', 'Block 2',   'Single Arm Plank',           3, '20 sec/side'],
    ['Monday', 'Block 2',   'Archer Push-Up',             3, '6/side'],
    ['Monday', 'Finisher',  'Full Integration Circuit',   1, '1 round', 'RKC Plank 30s -> Explosive Push-Up 8 -> Single Arm Plank 20s/side -> L-Sit 20s, no rest'],
    ['Monday', 'Cool-Down', "Child's Pose",               1, '45 sec'],
    ['Monday', 'Cool-Down', 'Supine Hamstring Stretch',   1, '30 sec/side'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',        1, '30 sec/side'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Lying Knee Rocks',           1, '12 reps'],
    ['Wednesday', 'Warm-Up',   'Single-Leg Glute Bridge',    1, '10/side'],
    ['Wednesday', 'Warm-Up',   'Arm Circles',                1, '10 each direction'],
    ['Wednesday', 'Block 1',   'Hollow Body Hold',           3, '35 sec'],
    ['Wednesday', 'Block 1',   'Dragon Flag Negative',       3, '5 reps'],
    ['Wednesday', 'Block 1',   'Stir the Pot',               3, '8/direction'],
    ['Wednesday', 'Block 2',   'Mountain Climber Tabata',    3, '20 sec on/10 sec off'],
    ['Wednesday', 'Block 2',   'Unilateral Dead Bug',        3, '8/side'],
    ['Wednesday', 'Block 2',   'Archer Push-Up',             3, '6/side'],
    ['Wednesday', 'Finisher',  'Culmination Gauntlet',       1, '1 round', 'Hollow Body 30s -> Clap Push-Up 6 -> Unilateral Dead Bug 8/side -> Archer Push-Up 5/side, no rest'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',         1, '45 sec/side'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Stretch',           1, '30 sec/side'],
    ['Wednesday', 'Cool-Down', "Child's Pose",               1, '30 sec'],

    // Friday
    ['Friday', 'Warm-Up',   'Cat-Camel Stretch',       1, '10 reps'],
    ['Friday', 'Warm-Up',   'Bear Crawl In Place',     1, '30 sec'],
    ['Friday', 'Warm-Up',   'Inchworm Walk-Out',       1, '6 reps'],
    ['Friday', 'Block 1',   'Dead Bug Full',           3, '8/side'],
    ['Friday', 'Block 1',   'Band Pallof Press Blue',  3, '12/side'],
    ['Friday', 'Block 1',   'Copenhagen Plank',        3, '20 sec/side'],
    ['Friday', 'Block 2',   'Diamond Push-Up',         3, '10 reps'],
    ['Friday', 'Block 2',   'Archer Push-Up',          3, '6/side'],
    ['Friday', 'Block 2',   'Single Leg Glute Bridge', 3, '12/side'],
    ['Friday', 'Finisher',  'Max Effort Integration',  1, '1 round', '30s Explosive Push-Up -> rest 15s -> 30s Mountain Climbers -> rest 15s -> Max Archer Push-Ups/side'],
    ['Friday', 'Cool-Down', "Child's Pose",            1, '45 sec'],
    ['Friday', 'Cool-Down', 'Cobra to Upward Dog',     1, '30 sec'],
    ['Friday', 'Cool-Down', 'Seated Spinal Twist',     1, '30 sec/side']

  ]);


  // == HUSBAND ARMS P5C1 == Tue / Thu / Sat =====================================

  buildSheet('Husband Arms P5C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Arm Circles',                       1, '15 each direction'],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Blue',              1, '15 reps'],
    ['Tuesday', 'Warm-Up',   'Wrist Circles',                     1, '10 each direction'],
    ['Tuesday', 'Block 1',   'Diamond Push-Up',                   3, '12 reps'],
    ['Tuesday', 'Block 1',   'Slow Negative Table Inverted Row',  3, '8 reps'],
    ['Tuesday', 'Block 1',   'Band Curl Blue',                    3, '12 reps'],
    ['Tuesday', 'Block 2',   'Clap Push-Up',                      3, '8 reps'],
    ['Tuesday', 'Block 2',   'Single Arm Table Inverted Row',     3, '8/side'],
    ['Tuesday', 'Block 2',   'Archer Push-Up',                    3, '6/side'],
    ['Tuesday', 'Finisher',  'Arms Integration Triset',           1, '1 round', 'Typewriter Push-Up 6/side -> Clap Push-Up 8 -> Single Arm Push-Up 4/side, no rest'],
    ['Tuesday', 'Cool-Down', 'Overhead Tricep Stretch',           1, '30 sec/side'],
    ['Tuesday', 'Cool-Down', 'Cross-Body Bicep Stretch',          1, '30 sec/side'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Arm Circles Fast',           1, '15 each direction'],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Blue',       1, '15 reps'],
    ['Thursday', 'Warm-Up',   'Wrist Circles',              1, '10 each direction'],
    ['Thursday', 'Block 1',   'Pike Push-Up',               3, '10 reps'],
    ['Thursday', 'Block 1',   'Pseudo Planche Lean',        3, '20 sec'],
    ['Thursday', 'Block 1',   'Band Face Pull Green',       3, '15 reps'],
    ['Thursday', 'Block 2',   'Explosive Pike Push-Up',     3, '8 reps'],
    ['Thursday', 'Block 2',   'Single Arm Pike Push-Up',    3, '6/side'],
    ['Thursday', 'Block 2',   'Archer Row',                 3, '8/side'],
    ['Thursday', 'Finisher',  'Shoulder Culmination',       1, '1 round', 'Pseudo Planche Lean 20s -> Explosive Pike Push-Up 8 -> Single Arm Pike Push-Up 5/side, no rest'],
    ['Thursday', 'Cool-Down', 'Doorway Chest Stretch',      1, '30 sec'],
    ['Thursday', 'Cool-Down', "Child's Pose Wide Arms",     1, '45 sec'],

    // Saturday
    ['Saturday', 'Warm-Up',   'Band Pull-Apart Blue',               1, '15 reps'],
    ['Saturday', 'Warm-Up',   'Inchworm Walk-Out',                  1, '6 reps'],
    ['Saturday', 'Warm-Up',   'Wrist Circles',                      1, '10 each direction'],
    ['Saturday', 'Block 1',   'Typewriter Push-Up',                 3, '6/side'],
    ['Saturday', 'Block 1',   'Explosive Table Inverted Row',       3, '8 reps'],
    ['Saturday', 'Block 1',   'Single Arm Push-Up Progression',     3, '5/side'],
    ['Saturday', 'Block 2',   'Clap Push-Up',                       3, '8 reps'],
    ['Saturday', 'Block 2',   'Band Curl to Press Blue',            3, '10 reps'],
    ['Saturday', 'Block 2',   'Archer Row',                         3, '10/side'],
    ['Saturday', 'Finisher',  'Culmination Triset',                 1, '1 round max effort', 'Clap Push-Up 8 -> Explosive Row 8 -> Single Arm Push-Up 4/side, no rest'],
    ['Saturday', 'Cool-Down', 'Overhead Tricep Stretch',            1, '30 sec/side'],
    ['Saturday', 'Cool-Down', 'Doorway Chest Opener',               1, '30 sec'],
    ['Saturday', 'Cool-Down', 'Wrist Flexor Stretch',               1, '20 sec/side']

  ]);


  // == WIFE CORE P5C1 == Mon / Wed / Fri ========================================

  buildSheet('Wife Core P5C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Diaphragmatic Breathing',    1, '8 breaths'],
    ['Monday', 'Warm-Up',   'Pelvic Tilts',               1, '12 reps'],
    ['Monday', 'Warm-Up',   'Glute Bridge',               1, '12 reps'],
    ['Monday', 'Block 1',   'Bird Dog',                   3, '10/side'],
    ['Monday', 'Block 1',   'Bear Hold',                  3, '25 sec'],
    ['Monday', 'Block 1',   'Modified Copenhagen Plank',  3, '20 sec/side'],
    ['Monday', 'Block 2',   'Speed Bird Dog',             3, '10/side fast'],
    ['Monday', 'Block 2',   'Single Leg Glute Bridge',    3, '12/side'],
    ['Monday', 'Block 2',   'Low Impact Mountain Climber',3, '10/side'],
    ['Monday', 'Finisher',  'Wife Integration Flow',      1, '1 round', 'Bird Dog 30s -> Bear Hold 20s -> Single Leg Glute Bridge 10/side -> Speed March 20s, no rest'],
    ['Monday', 'Cool-Down', "Child's Pose",               1, '45 sec'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',        1, '30 sec/side'],
    ['Monday', 'Cool-Down', 'Knees-to-Chest Hug',         1, '30 sec'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Diaphragmatic Breathing',        1, '6 breaths'],
    ['Wednesday', 'Warm-Up',   'Lying Knee Rocks',               1, '12 reps'],
    ['Wednesday', 'Warm-Up',   'Glute Bridge',                   1, '12 reps'],
    ['Wednesday', 'Block 1',   'Forearm Plank',                  3, '30 sec'],
    ['Wednesday', 'Block 1',   'Hollow Body Hold Bent Knees',    3, '30 sec'],
    ['Wednesday', 'Block 1',   'Band Pallof Press Tan',          3, '12/side'],
    ['Wednesday', 'Block 2',   'Modified Burpee',                3, '8 reps'],
    ['Wednesday', 'Block 2',   'Single Arm Bird Dog',            3, '10/side'],
    ['Wednesday', 'Block 2',   'Lateral Lunge',                  3, '10/side'],
    ['Wednesday', 'Finisher',  'Culmination Flow',               1, '1 round', 'Forearm Plank 30s -> Modified Burpee 6 -> Single Arm Bird Dog 8/side -> Lateral Lunge 8/side, no rest'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',             1, '40 sec/side'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Glute Stretch',         1, '30 sec/side'],
    ['Wednesday', 'Cool-Down', "Child's Pose",                   1, '30 sec'],

    // Friday
    ['Friday', 'Warm-Up',   'Diaphragmatic Breathing',       1, '8 breaths'],
    ['Friday', 'Warm-Up',   'Pelvic Tilts',                  1, '12 reps'],
    ['Friday', 'Warm-Up',   'Cat-Camel',                     1, '10 reps'],
    ['Friday', 'Block 1',   'Dead Bug Full',                 3, '8/side'],
    ['Friday', 'Block 1',   'Modified Copenhagen Plank',     3, '20 sec/side'],
    ['Friday', 'Block 1',   'Bear Hold',                     3, '25 sec'],
    ['Friday', 'Block 2',   'Small Squat Jump',              3, '10 reps'],
    ['Friday', 'Block 2',   'Single Leg Glute Bridge',       3, '12/side'],
    ['Friday', 'Block 2',   'Low Impact Mountain Climber',   3, '12/side'],
    ['Friday', 'Finisher',  'Wife Culmination Circuit',      1, '1 round', 'Dead Bug 8/side -> Small Squat Jump 10 -> Single Leg Glute Bridge 10/side -> Bear Hold 20s, no rest'],
    ['Friday', 'Cool-Down', "Child's Pose",                  1, '1 min'],
    ['Friday', 'Cool-Down', 'Supine Spinal Twist',           1, '40 sec/side'],
    ['Friday', 'Cool-Down', 'Happy Baby Pose',               1, '30 sec']

  ]);


  // == WIFE BANDS P5C1 == Tue / Thu =============================================

  buildSheet('Wife Bands P5C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Diaphragmatic Breathing',            1, '6 breaths'],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Blue',               1, '12 reps'],
    ['Tuesday', 'Warm-Up',   'Hip Circles Standing',               1, '10/direction'],
    ['Tuesday', 'Block 1',   'Band Bicep Curl Green',              3, '12 reps'],
    ['Tuesday', 'Block 1',   'Band Overhead Press Green',          3, '10 reps'],
    ['Tuesday', 'Block 1',   'Band Glute Bridge Blue',             3, '12 reps'],
    ['Tuesday', 'Block 2',   'Band Bicep Curl Fast Green',         3, '15 reps fast'],
    ['Tuesday', 'Block 2',   'Single Arm Band Curl Green',         3, '12/side'],
    ['Tuesday', 'Block 2',   'Single Leg Band Glute Bridge Green', 3, '12/side'],
    ['Tuesday', 'Finisher',  'Band Culmination Circuit',           1, '1 round', 'Band Row 15 -> Fast Curl 15 -> Single Arm Overhead Press 10/side -> Single Leg Glute Bridge 10/side, no rest'],
    ['Tuesday', 'Cool-Down', "Child's Pose",                       1, '45 sec'],
    ['Tuesday', 'Cool-Down', 'Seated Spinal Twist',                1, '30 sec/side'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Diaphragmatic Breathing',               1, '6 breaths'],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Blue',                  1, '12 reps'],
    ['Thursday', 'Warm-Up',   'Single Leg Glute Bridge',               1, '10/side'],
    ['Thursday', 'Block 1',   'Band Row Green',                        3, '12 reps'],
    ['Thursday', 'Block 1',   'Band Squat Blue',                       3, '10 reps'],
    ['Thursday', 'Block 1',   'Band Face Pull Green',                  3, '15 reps'],
    ['Thursday', 'Block 2',   'Band Squat Fast Green',                 3, '15 reps fast'],
    ['Thursday', 'Block 2',   'Single Arm Band Row Green',             3, '12/side'],
    ['Thursday', 'Block 2',   'Single Leg Band Romanian Deadlift Tan', 3, '10/side'],
    ['Thursday', 'Finisher',  'Band Culmination Burnout',              1, '1 round', 'Fast Squat 15 -> Single Arm Row 10/side -> Single Leg RDL 8/side -> Fast Glute Bridge 15, no rest'],
    ['Thursday', 'Cool-Down', 'Figure-4 Glute Stretch',                1, '30 sec/side'],
    ['Thursday', 'Cool-Down', "Child's Pose",                          1, '30 sec']

  ]);


  SpreadsheetApp.flush();
  Logger.log('Done - all 4 Phase 5 Culmination tabs created successfully.');
}
