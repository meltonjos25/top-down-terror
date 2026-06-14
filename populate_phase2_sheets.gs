function populatePhase2Sheets() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  // Phase 2 Strength orange
  var PHASE_BG    = '#e65c00';   // rich orange title banner
  var DAY_TINT    = '#fff3e0';   // light orange for Day column
  var HEADER_BG   = '#263238';   // dark slate header row

  // Block background colors
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
    'Husband Core P2C1': 'Husband Core Workout — P2C1 (Monday · Wednesday · Friday) | Intermediate–Advanced',
    'Husband Arms P2C1': 'Husband Arms Workout — P2C1 (Tuesday · Thursday · Saturday) | Intermediate–Advanced',
    'Wife Core P2C1':    'Wife Core Workout — P2C1 (Monday · Wednesday · Friday) | Beginner · Breastfeeding-Safe',
    'Wife Bands P2C1':   'Wife Bands Workout — P2C1 (Tuesday · Thursday) | Beginner · Breastfeeding-Safe'
  };

  function buildSheet(name, exercises) {
    // Delete and recreate
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

      // Day cell — orange tint, bold, centered
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


  // ── HUSBAND CORE P2C1 ── Mon / Wed / Fri ──────────────────────────────────

  buildSheet('Husband Core P2C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Cat-Camel Stretch',              1, '10 reps',           'Move slowly between arching and rounding your spine, 2 seconds each position. Keep arms straight and breathe throughout.'],
    ['Monday', 'Warm-Up',   'Dead Bug Arms Only',             1, '10/side',            'Press your lower back firmly into the floor the entire time. Extend one arm overhead at a time while keeping your core braced.'],
    ['Monday', 'Warm-Up',   'Band Pull-Apart Green',          1, '15 reps',            'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Monday', 'Warm-Up',   'Hip Flexor Lunge Stretch',       1, '30 sec/side',        'Step into a lunge and lower your back knee to the floor. Tuck your pelvis slightly and keep your torso upright to feel the stretch in your front hip.'],
    ['Monday', 'Block 1',   'Copenhagen Plank Knee Down',     3, '20 sec/side',        'Lie on your side with top foot on a chair and bottom knee on the floor. Lift your hips to form a straight line and breathe steadily - never hold your breath.'],
    ['Monday', 'Block 1',   'Hollow Body Hold',               3, '40 sec',             'Flatten your lower back into the floor and lift your shoulder blades and legs off the ground. Breathe steadily - if your back rises, raise your legs higher.'],
    ['Monday', 'Block 1',   'L-Sit Hold on Chairs',           3, '20 sec',             'Press down hard through your palms with arms fully locked straight and lift your hips off the chairs. Breathe steadily and never hold your breath.'],
    ['Monday', 'Block 2',   'Dragon Flag Negative',           3, '5 reps slow',        'Grip a sturdy surface behind your head and lift your body up as one rigid plank. Lower yourself down over 4 to 5 seconds - as slowly as possible.'],
    ['Monday', 'Block 2',   'Stir the Pot on Forearms',       3, '8/direction',        'Get into a forearm plank and draw slow circles with your forearms. Keep your hips completely still and level - no rocking side to side.'],
    ['Monday', 'Block 2',   'Band Pallof Press Blue',         3, '12/side',            'Anchor the blue band at chest height and stand sideways to it. Press the band straight out from your chest, hold 2 seconds, and resist any pull to rotate.'],
    ['Monday', 'Finisher',  'Max L-Sit Hold on Chairs',       1, '1 set to failure',   'Press down through both palms with arms fully locked straight. Breathe steadily and hold until your form breaks.'],
    ['Monday', 'Cool-Down', "Child's Pose",                   1, '45 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],
    ['Monday', 'Cool-Down', 'Cobra Stretch',                  1, '30 sec',             'Lie face down and press your hands into the floor under your shoulders. Straighten your arms slowly and keep your hips on the floor throughout.'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',            1, '30 sec/side',        'Sit tall and rotate your torso, using a gentle hand press on your knee to deepen the twist. Keep your spine long and breathe slowly.'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Hip Circles',                    1, '10/direction',       'Stand with feet shoulder-width apart and hands on hips. Draw large slow circles with your hips in each direction to loosen the hip joints.'],
    ['Wednesday', 'Warm-Up',   'Lying Knee Rocks',               1, '12 reps',            'Lie on your back with knees bent and slowly lower both knees to one side. Keep your shoulders flat on the floor, then return to center.'],
    ['Wednesday', 'Warm-Up',   'Bear Crawl In Place',            1, '30 sec',             'Start on all fours with knees hovering 1 inch off the floor. Hold with a flat back and breathe steadily throughout.'],
    ['Wednesday', 'Warm-Up',   'Inchworm Walk-Out',              1, '6 reps',             'Bend forward and walk your hands out to a plank position, keeping legs as straight as possible. Walk hands back and stand tall to complete each rep.'],
    ['Wednesday', 'Block 1',   'Slow Negative Push-Up',          3, '8 reps',             'Lower yourself down for a full 4 seconds with elbows tracking straight back along your ribs. Press back up normally once you reach the floor.'],
    ['Wednesday', 'Block 1',   'Band Resisted Crunch Blue',      3, '15 reps',            'Anchor the blue band overhead and hold both ends at your temples. Curl your upper back off the floor and squeeze your abs hard at the top.'],
    ['Wednesday', 'Block 1',   'Jefferson Curl with Band Tan',   3, '8 reps',             'Stand on the tan band and roll down one vertebra at a time, letting your head drop first. Uncurl slowly back to standing - keep the movement controlled.'],
    ['Wednesday', 'Block 2',   'Side Plank with Band Row Blue',  3, '10/side',            'Anchor the blue band at floor level and get into a side plank facing the anchor. Row the band to your hip while keeping your hips lifted throughout.'],
    ['Wednesday', 'Block 2',   'Hollow Body Rock',               3, '35 sec',             'Press your lower back into the floor in hollow position with arms overhead and rock forward and back. Maintain total body tension throughout.'],
    ['Wednesday', 'Block 2',   'Dead Bug with Band Tan',         3, '8/side',             'Hold the tan band overhead with both hands while on your back. Extend the opposite leg as you reach overhead - keep lower back pressed into the floor.'],
    ['Wednesday', 'Finisher',  'Max Slow Negative Push-Ups',     1, '1 set to failure',   'Lower yourself to the floor as slowly as possible - aim for 5 seconds per rep. Continue until you can no longer control the descent.'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',             1, '45 sec/side',        'Lie on your back and let your knees fall to one side. Breathe deeply and let gravity do the work - never force or bounce the stretch.'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Stretch',               1, '30 sec/side',        'Lie on your back and cross one ankle over the opposite knee. Pull the bottom leg toward you gently and breathe steadily to deepen the stretch.'],
    ['Wednesday', 'Cool-Down', "Child's Pose",                   1, '30 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],

    // Friday
    ['Friday', 'Warm-Up',   'Cat-Camel Stretch',              1, '10 reps',            'Move slowly between arching and rounding your spine, 2 seconds each position. Keep arms straight and breathe throughout.'],
    ['Friday', 'Warm-Up',   'Band Pull-Apart Green',          1, '15 reps',            'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Friday', 'Warm-Up',   'Hip Flexor Lunge Stretch',       1, '30 sec/side',        'Step into a lunge and lower your back knee to the floor. Tuck your pelvis slightly and keep your torso upright to feel the stretch in your front hip.'],
    ['Friday', 'Warm-Up',   'Arm Circles',                    1, '10 each direction',  'Keep arms straight and make large controlled circles forward then backward. Start small and gradually increase the range of motion.'],
    ['Friday', 'Block 1',   'Copenhagen Plank Full',          3, '25 sec/side',        'Lie on your side with top foot on a chair and your bottom leg lifted off the floor. Lift hips to form a straight line and breathe steadily - never hold your breath.'],
    ['Friday', 'Block 1',   'Dragon Flag Negative',           3, '6 reps',             'Grip a sturdy surface behind your head and lift your body up as one rigid plank. Lower yourself down over 4 to 5 seconds - as slowly as possible.'],
    ['Friday', 'Block 1',   'Stir the Pot',                   3, '10/direction',       'Get into a forearm plank and draw slow circles with your forearms in each direction. Keep your hips completely still and level throughout.'],
    ['Friday', 'Block 2',   'Band Resisted Hollow Body Blue', 3, '30 sec',             'Anchor the blue band overhead and hold both ends in your hollow body position. The band pulls you up - resist it to increase core tension and breathe steadily.'],
    ['Friday', 'Block 2',   'L-Sit Hold on Chairs',           3, '25 sec',             'Press down hard through your palms with arms fully locked straight and lift your hips off the chairs. Breathe steadily and never hold your breath.'],
    ['Friday', 'Block 2',   'Band Pallof Press Blue',         3, '12/side',            'Anchor the blue band at chest height and stand sideways to it. Press straight out from your chest, hold 2 seconds, and resist any pull to rotate.'],
    ['Friday', 'Finisher',  'Strength Core Gauntlet',         1, '1 round',            'Copenhagen 20s/side -> Dragon Flag 5 -> L-Sit 20s, zero rest. Move immediately between exercises and maintain quality form throughout.'],
    ['Friday', 'Cool-Down', "Child's Pose",                   1, '45 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],
    ['Friday', 'Cool-Down', 'Supine Hamstring Stretch',       1, '30 sec/side',        'Lie on your back and extend one leg straight up. Use your hands behind the thigh and breathe steadily to deepen the stretch with each exhale.'],
    ['Friday', 'Cool-Down', 'Cobra to Upward Dog',            1, '30 sec',             'From a push-up position, press hips toward the floor to cobra, then push hips up and back to upward dog. Move fluidly and breathe throughout.']

  ]);


  // ── HUSBAND ARMS P2C1 ── Tue / Thu / Sat ─────────────────────────────────

  buildSheet('Husband Arms P2C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Arm Circles',                       1, '15 each direction', 'Keep arms straight and make large controlled circles forward then backward. Start small and gradually increase the range of motion.'],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Green',             1, '15 reps',           'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Tuesday', 'Warm-Up',   'Wrist Circles',                     1, '10 each direction', 'Lace your fingers together and draw slow circles in both directions. Move through the full range of motion before band and push-up work.'],
    ['Tuesday', 'Warm-Up',   'Shoulder Cross-Body Stretch',       1, '20 sec/side',       'Hold one arm straight across your chest and use your other hand to gently press it closer. Keep the shoulder of the working arm pressed down, not hunched up.'],
    ['Tuesday', 'Block 1',   'Slow Negative Table Inverted Row',  3, '8 reps',            'Grip the table edge and hang underneath with your body rigid like a plank. Pull up normally, then lower yourself down for a full 4 seconds.'],
    ['Tuesday', 'Block 1',   'Band Curl Blue',                    3, '12 reps',           'Stand on the blue band and curl both hands up to your shoulders. Lower slowly over 3 seconds - keep elbows pinned at your sides and never swing.'],
    ['Tuesday', 'Block 1',   'Isometric Table Inverted Row Hold', 3, '30 sec',            'Pull yourself up until your chest touches the table edge and hold that position. Breathe steadily and never hold your breath throughout.'],
    ['Tuesday', 'Block 2',   'Slow Negative Close-Grip Push-Up',  3, '8 reps',            'Keep your elbows tracking straight back along your ribs throughout. Lower yourself for a full 4 seconds, then press back up normally.'],
    ['Tuesday', 'Block 2',   'Band Tricep Pushdown Blue',         3, '15 reps',           'Anchor the blue band overhead and press both hands down to full extension. Hold 1 second at the bottom, then let the band rise slowly over 3 seconds.'],
    ['Tuesday', 'Block 2',   'Tricep Dip Negative',               3, '8 reps',            'Place hands on a chair behind you with arms straight. Lower yourself down for a full 4 seconds, then press back up to the start.'],
    ['Tuesday', 'Finisher',  'Max Slow Negative Push-Ups',        1, '1 set to failure',  'Take as long as possible to lower yourself - aim for 5 or more seconds per rep. Continue until you can no longer control the descent.'],
    ['Tuesday', 'Cool-Down', 'Overhead Tricep Stretch',           1, '30 sec/side',       'Reach one arm straight overhead, bend at the elbow, and gently press the elbow with your opposite hand. Keep your neck long and breathe steadily.'],
    ['Tuesday', 'Cool-Down', 'Cross-Body Bicep Stretch',          1, '30 sec/side',       'Extend one arm straight across your chest and use your other hand to gently pull it closer. Keep the shoulder of the working arm pressed down.'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Arm Circles',                      1, '15 each direction', 'Keep arms straight and make large controlled circles forward then backward. Start small and gradually increase the range of motion.'],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Green',            1, '15 reps',           'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Thursday', 'Warm-Up',   'Cat-Camel',                        1, '10 reps',           'Move slowly between arching and rounding your spine, 2 seconds each position. Keep arms straight and breathe throughout.'],
    ['Thursday', 'Warm-Up',   'Wrist Circles',                    1, '10 each direction', 'Lace your fingers together and draw slow circles in both directions. Move through the full range of motion before pressing work.'],
    ['Thursday', 'Block 1',   'Slow Negative Pike Push-Up',       3, '8 reps',            'Start in a pike position with hips high and elbows pointing back. Lower your head toward the floor for a full 4 seconds, then press back up explosively.'],
    ['Thursday', 'Block 1',   'Wall Handstand Hold',              3, '25 sec',            'Place your feet up the wall with arms straight and core tight. Breathe steadily - open your ribs to breathe rather than holding your breath.'],
    ['Thursday', 'Block 1',   'Band Overhead Press Blue',         3, '12 reps',           'Stand on the blue band with feet shoulder-width apart. Press both hands straight overhead to full lockout, then lower slowly with control.'],
    ['Thursday', 'Block 2',   'Typewriter Push-Up',               3, '6/side',            'Lower to the bottom of a push-up, then slide your chest from side to side before pressing back up. Move deliberately and keep your hips level throughout.'],
    ['Thursday', 'Block 2',   'Pseudo Planche Lean',              3, '20 sec',            'In push-up position, shift your body forward until your shoulders are past your hands. Hold perfectly still and breathe steadily - never hold your breath.'],
    ['Thursday', 'Block 2',   'Band Face Pull Green',             3, '15 reps',           'Anchor the green band at face height and pull both ends to your temples with elbows flared wide. Squeeze your rear shoulder blades hard at the end of each rep.'],
    ['Thursday', 'Finisher',  'Max Typewriter Push-Ups',          1, '1 set to failure',  'Perform full typewriter reps - lower, slide left, slide right, press up - as many as you can with good form. Stop when your hips start rocking.'],
    ['Thursday', 'Cool-Down', 'Doorway Chest Stretch',            1, '30 sec',            'Stand in a doorway with both arms at 90 degrees and lean your body forward gently. Keep your chin tucked and breathe steadily.'],
    ['Thursday', 'Cool-Down', "Child's Pose Wide Arms",           1, '45 sec',            'Sink your hips back toward your heels and reach your arms as wide as possible on the floor. Breathe deeply and let your lats release with each exhale.'],

    // Saturday
    ['Saturday', 'Warm-Up',   'Jumping Jacks',                    1, '30 sec',            'Keep arms fully extended overhead at the top and land softly with slightly bent knees. Maintain a steady rhythm throughout.'],
    ['Saturday', 'Warm-Up',   'Band Pull-Apart Green',            1, '15 reps',           'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Saturday', 'Warm-Up',   'Inchworm Walk-Out',                1, '6 reps',            'Bend forward and walk your hands out to a plank, keeping legs as straight as possible. Walk hands back in and stand tall to complete each rep.'],
    ['Saturday', 'Warm-Up',   'Wrist Circles',                    1, '10 each direction', 'Lace your fingers together and draw slow circles in both directions to warm up the wrists before pushing and curling movements.'],
    ['Saturday', 'Block 1',   'Pseudo Planche Push-Up',           3, '8 reps',            'Place your hands by your hips with fingers pointed back. Push up with your body leaning forward - this loads your shoulders much more than a regular push-up.'],
    ['Saturday', 'Block 1',   'Slow Negative Table Inverted Row', 3, '8 reps',            'Grip the table edge and hang underneath with your body rigid like a plank. Pull up normally, then lower yourself down for a full 4 seconds.'],
    ['Saturday', 'Block 1',   'Band Curl to Press Blue',          3, '10 reps',           'Stand on the blue band, curl up to your shoulders, then press overhead in one smooth motion. Lower with control and no jerking.'],
    ['Saturday', 'Block 2',   'Typewriter Push-Up',               3, '6/side',            'Lower to the bottom of a push-up, then slide your chest from side to side before pressing back up. Move deliberately and keep hips level.'],
    ['Saturday', 'Block 2',   'Band Tricep Pushdown Blue',        3, '15 reps',           'Anchor the blue band overhead and press both hands down to full extension. Hold 1 second at the bottom, then let the band rise slowly.'],
    ['Saturday', 'Block 2',   'Pike Push-Up Negative',            3, '8 reps',            'Start in a pike position with hips high and elbows pointing back. Lower your head toward the floor for a full 4 seconds, then press back up explosively.'],
    ['Saturday', 'Finisher',  'Strength Triset',                  1, '1 round',           'Pseudo Planche Push-Up 8 -> Band Curl Blue 12 -> Typewriter Push-Up 5/side, zero rest. Maintain quality form on every rep - stop if form breaks.'],
    ['Saturday', 'Cool-Down', 'Overhead Tricep Stretch',          1, '30 sec/side',       'Reach one arm straight overhead, bend at the elbow, and press the elbow gently with your other hand. Keep your neck long and breathe steadily.'],
    ['Saturday', 'Cool-Down', 'Doorway Chest Opener',             1, '30 sec',            'Stand in a doorway with both arms at 90 degrees and lean your body forward gently. Breathe steadily and let your chest open up.'],
    ['Saturday', 'Cool-Down', 'Wrist Flexor Stretch',             1, '20 sec/side',       'Extend one arm straight with palm facing up and gently pull the fingers back with your other hand. Hold steadily without bouncing.']

  ]);


  // ── WIFE CORE P2C1 ── Mon / Wed / Fri ────────────────────────────────────

  buildSheet('Wife Core P2C1', [

    // Monday
    ['Monday', 'Warm-Up',   'Diaphragmatic Breathing',              1, '8 breaths',          'Breathe in through your nose letting your belly rise, then exhale gently. Keep it slow - never hold your breath.'],
    ['Monday', 'Warm-Up',   'Pelvic Tilts',                         1, '12 reps',            'Lie on your back with knees bent and gently flatten your lower back into the floor. Hold 2 seconds, then release - use your core, not your feet.'],
    ['Monday', 'Warm-Up',   'Cat-Camel',                            1, '10 reps',            'Move slowly between arching and rounding your spine, 2 seconds each position. Keep arms straight and breathe throughout.'],
    ['Monday', 'Warm-Up',   'Band Pull-Apart Tan',                  1, '12 reps',            'Hold the tan band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades together.'],
    ['Monday', 'Block 1',   'Modified Copenhagen Plank Knee Down',  3, '20 sec/side',        'Lie on your side with top foot on a chair and bottom knee resting on the floor. Lift your hips to form a straight line and breathe steadily - never hold your breath.'],
    ['Monday', 'Block 1',   'Hollow Body Hold Bent Knees',          3, '25 sec',             'Press your lower back firmly into the floor and pull your knees toward your chest while lifting your shoulder blades. Breathe steadily - if your back rises, tuck knees in closer.'],
    ['Monday', 'Block 1',   'Bear Hold',                            3, '20 sec',             'Start on all fours and lift your knees just 1 inch off the floor. Keep your back flat like a table and breathe steadily - never hold your breath.'],
    ['Monday', 'Block 2',   'Dead Bug with Tan Band',               3, '8/side',             'Hold the tan band overhead with both hands while on your back. Extend one leg at a time while keeping your lower back pressed firmly into the floor.'],
    ['Monday', 'Block 2',   'Slow Heel Slides',                     3, '10/side',            'Lie on your back with knees bent and slowly slide one heel out until your leg is nearly straight. Keep your lower back pressed to the floor and return with control.'],
    ['Monday', 'Block 2',   'Seated Band Core Row Tan',             3, '12 reps',            'Sit upright with the tan band under your feet and row both hands to your waist. Squeeze your shoulder blades at the top, then resist the band as you return slowly.'],
    ['Monday', 'Finisher',  'Max Bear Hold',                        1, '1 set to failure',   'Lift your knees 1 inch off the floor with a completely flat back. Breathe slowly and steadily - never hold your breath - and hold until your form breaks.'],
    ['Monday', 'Cool-Down', "Child's Pose",                         1, '45 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],
    ['Monday', 'Cool-Down', 'Seated Spinal Twist',                  1, '30 sec/side',        'Sit tall and rotate your torso, using a gentle hand press on your knee to deepen the twist. Keep your spine long and breathe slowly.'],
    ['Monday', 'Cool-Down', 'Knees-to-Chest Hug',                   1, '30 sec',             'Lie on your back and pull both knees to your chest. Gently rock side to side and breathe deeply to release your lower back.'],

    // Wednesday
    ['Wednesday', 'Warm-Up',   'Diaphragmatic Breathing',           1, '6 breaths',          'Breathe in through your nose letting your belly rise, then exhale gently. Keep it slow - never hold your breath.'],
    ['Wednesday', 'Warm-Up',   'Hip Circles Standing',              1, '10/direction',       'Stand with feet shoulder-width apart and hands on hips. Draw large slow circles with your hips to loosen the hip joints.'],
    ['Wednesday', 'Warm-Up',   'Lying Knee Rocks',                  1, '12 reps',            'Lie on your back with knees bent and slowly lower both knees to one side. Keep your shoulders flat on the floor, then return and repeat.'],
    ['Wednesday', 'Warm-Up',   'Glute Bridge',                      1, '12 reps',            'Lie on your back with knees bent and feet flat. Press through your heels and squeeze your glutes to lift your hips until your body forms a straight line.'],
    ['Wednesday', 'Block 1',   'Slow Negative Push-Up Knees Down',  3, '8 reps',             'Keep elbows tracking straight back along your ribs and lower yourself from your knees for a full 4 seconds. Press back up once you reach the floor.'],
    ['Wednesday', 'Block 1',   'Band Resisted Bird Dog Tan',        3, '10/side',            'Loop the tan band around one foot and hold the other end with the opposite hand. Extend arm and leg simultaneously and keep your hips completely level.'],
    ['Wednesday', 'Block 1',   'Side Plank Knees Down Hold',        3, '30 sec/side',        'Lie on your side with bottom knee on the floor and press up onto your forearm. Lift your hips to form a diagonal line and breathe steadily - never hold your breath.'],
    ['Wednesday', 'Block 2',   'Hollow Body Hold Bent Knees',       3, '30 sec',             'Press your lower back into the floor and pull your knees toward your chest while lifting your shoulder blades. Breathe steadily throughout.'],
    ['Wednesday', 'Block 2',   'Glute Bridge with Band Blue',       3, '15 reps',            'Place the blue band just above your knees. Press through your heels to bridge up and push your knees outward against the band at the top - never let knees cave in.'],
    ['Wednesday', 'Block 2',   'Band Pallof Press Tan',             3, '12/side',            'Anchor the tan band at chest height and stand sideways to it. Press straight out from your chest, hold 2 seconds, and resist any pull to rotate your torso.'],
    ['Wednesday', 'Finisher',  'Max Slow Negative Push-Ups Knees Down', 1, '1 set to failure', 'Lower yourself from your knees as slowly as possible - aim for 5 seconds per rep. Continue until you can no longer control the descent.'],
    ['Wednesday', 'Cool-Down', 'Lying Spinal Twist',                1, '40 sec/side',        'Lie on your back and let your knees fall gently to one side. Breathe deeply and let gravity do the work - never force or bounce the stretch.'],
    ['Wednesday', 'Cool-Down', 'Figure-4 Glute Stretch',            1, '30 sec/side',        'Lie on your back and cross one ankle over the opposite knee. Pull the bottom leg toward you gently and breathe steadily to deepen the glute stretch.'],
    ['Wednesday', 'Cool-Down', "Child's Pose",                      1, '30 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],

    // Friday
    ['Friday', 'Warm-Up',   'Diaphragmatic Breathing',              1, '8 breaths',          'Breathe in through your nose letting your belly rise, then exhale gently. Keep it slow - never hold your breath.'],
    ['Friday', 'Warm-Up',   'Pelvic Tilts',                         1, '12 reps',            'Gently flatten your lower back into the floor, hold 2 seconds, then release. Focus on using your core muscles - not pushing with your feet.'],
    ['Friday', 'Warm-Up',   'Cat-Camel',                            1, '10 reps',            'Move slowly between arching and rounding your spine, 2 seconds each position. Keep arms straight and breathe throughout.'],
    ['Friday', 'Warm-Up',   'Standing Hip Flexor Stretch',          1, '30 sec/side',        'Step into a lunge and lower your back knee to the floor. Tuck your pelvis slightly and keep your torso upright to feel the stretch deep in your front hip.'],
    ['Friday', 'Block 1',   'Modified Copenhagen Plank',            3, '25 sec/side',        'Lie on your side with your top foot on a chair. Lift your hips to form a straight line and breathe steadily - never hold your breath.'],
    ['Friday', 'Block 1',   'Bear Hold',                            3, '25 sec',             'Start on all fours and lift your knees just 1 inch off the floor. Keep your back flat like a table and breathe steadily throughout.'],
    ['Friday', 'Block 1',   'Band Pallof Press Tan',                3, '12/side',            'Anchor the tan band at chest height and stand sideways to it. Press straight out from your chest, hold 2 seconds, and resist any pull to rotate.'],
    ['Friday', 'Block 2',   'Dead Bug with Tan Band',               3, '10/side',            'Hold the tan band overhead with both hands. Extend one leg at a time while keeping your lower back pressed firmly into the floor.'],
    ['Friday', 'Block 2',   'Slow Negative Push-Up Knees Down',     3, '8 reps',             'Keep elbows tracking back along your ribs and lower yourself from your knees for a full 4 seconds. Press back up once you reach the floor.'],
    ['Friday', 'Block 2',   'Glute Bridge with Band Green',         3, '15 reps',            'Place the green band just above your knees. Press through your heels to bridge up and push your knees outward against the band at the top.'],
    ['Friday', 'Finisher',  'Strength Flow',                        1, '1 round',            'Bear Hold 20s -> Modified Copenhagen 20s/side -> Hollow Body Bent Knees 25s, zero rest. Focus on steady breathing throughout each hold.'],
    ['Friday', 'Cool-Down', "Child's Pose",                         1, '1 min',              'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back release with each exhale.'],
    ['Friday', 'Cool-Down', 'Supine Spinal Twist',                  1, '40 sec/side',        'Lie on your back and let your knees fall to one side. Breathe deeply and let gravity do the work - never force the stretch.'],
    ['Friday', 'Cool-Down', 'Happy Baby Pose',                      1, '30 sec',             'Lie on your back and hold the outside edges of your feet. Gently pull your knees toward your armpits, rock side to side, and breathe deeply.']

  ]);


  // ── WIFE BANDS P2C1 ── Tue / Thu ── arms and shoulders only ──────────────

  buildSheet('Wife Bands P2C1', [

    // Tuesday
    ['Tuesday', 'Warm-Up',   'Diaphragmatic Breathing',              1, '6 breaths',          'Breathe in through your nose letting your belly rise, then exhale gently. Keep it slow - never hold your breath.'],
    ['Tuesday', 'Warm-Up',   'Band Pull-Apart Green',                1, '12 reps',            'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Tuesday', 'Warm-Up',   'Wrist Circles',                        1, '10/direction',       'Lace your fingers together and draw slow circles in both directions. Move through the full range of motion before band work.'],
    ['Tuesday', 'Warm-Up',   'Shoulder Rolls',                       1, '10 each direction',  'Roll your shoulders in large circles forward then backward. Keep your arms relaxed and your neck long throughout.'],
    ['Tuesday', 'Block 1',   'Band Bicep Curl Slow Negative Green',  3, '10 reps',            'Stand on the green band and curl up in 2 seconds. Lower back down for a full 4 seconds - keep elbows pinned to your sides and never swing.'],
    ['Tuesday', 'Block 1',   'Band Overhead Press Green',            3, '10 reps',            'Stand on the green band with feet shoulder-width apart. Press both hands straight overhead to full lockout, then lower slowly with control.'],
    ['Tuesday', 'Block 1',   'Band Upright Row Green',               3, '10 reps',            'Stand on the green band with feet together. Pull both hands up to chin height with elbows leading upward, then lower with control.'],
    ['Tuesday', 'Block 2',   'Band Bicep Isometric Hold Green',      3, '20 sec',             'Stand on the green band and curl your hands to 90 degrees. Hold perfectly still with maximum tension in your biceps and breathe steadily - never hold your breath.'],
    ['Tuesday', 'Block 2',   'Band Tricep Overhead Extension Green', 3, '12 reps',            'Anchor the green band low and hold it with both hands behind your head. Extend your arms to full lockout overhead, then lower slowly with control.'],
    ['Tuesday', 'Block 2',   'Band Reverse Fly Tan',                 3, '12 reps',            'Hold the tan band with both hands in front of you and hinge slightly forward. Pull both arms wide with straight elbows, squeezing your rear shoulder blades together.'],
    ['Tuesday', 'Finisher',  'Band Curl Burnout Green',              1, 'Max reps',           'Stand on the green band and perform full bicep curls to failure. Keep elbows pinned to your sides - avoid swinging as you fatigue.'],
    ['Tuesday', 'Cool-Down', "Child's Pose",                         1, '45 sec',             'Sink your hips back toward your heels and reach your arms forward on the floor. Breathe deeply and let your back release with each exhale.'],
    ['Tuesday', 'Cool-Down', 'Overhead Tricep Stretch',              1, '30 sec/side',        'Reach one arm straight overhead, bend at the elbow, and gently press the elbow with your other hand. Keep your neck long and breathe steadily.'],
    ['Tuesday', 'Cool-Down', 'Cross-Body Shoulder Stretch',          1, '20 sec/side',        'Hold one arm straight across your chest and use your other hand to gently press it closer. Keep the shoulder of the working arm pressed down, not hunched.'],

    // Thursday
    ['Thursday', 'Warm-Up',   'Diaphragmatic Breathing',             1, '6 breaths',          'Breathe in through your nose letting your belly rise, then exhale gently. Keep it slow - never hold your breath.'],
    ['Thursday', 'Warm-Up',   'Band Pull-Apart Green',               1, '12 reps',            'Hold the green band at shoulder height with arms straight. Pull both ends apart until it touches your chest, squeezing your shoulder blades.'],
    ['Thursday', 'Warm-Up',   'Arm Circles',                         1, '10/direction',       'Keep arms straight and make large controlled circles forward then backward. Start small and gradually increase the range of motion.'],
    ['Thursday', 'Warm-Up',   'Shoulder Rolls',                      1, '10 each direction',  'Roll your shoulders in large circles forward then backward. Keep your arms relaxed and your neck long throughout.'],
    ['Thursday', 'Block 1',   'Band Row Slow Negative Green',        3, '10 reps',            'Anchor the green band at waist height and row both hands in, squeezing your shoulder blades. Lower your arms forward over a full 4 seconds - resist the pull slowly.'],
    ['Thursday', 'Block 1',   'Band Face Pull Green',                3, '12 reps',            'Anchor the green band at face height and pull both ends to your temples with elbows flared wide. Squeeze your rear shoulder blades hard at the end of each rep.'],
    ['Thursday', 'Block 1',   'Band Reverse Fly Tan',                3, '12 reps',            'Hold the tan band with both hands and hinge slightly forward. Pull both arms wide with straight elbows, squeezing your rear shoulder blades together.'],
    ['Thursday', 'Block 2',   'Band Lateral Raise Green',            3, '12/side',            'Stand on the green band and raise both arms out to shoulder height with a slight bend in your elbows. Lower slowly - do not let the band snap your arms down.'],
    ['Thursday', 'Block 2',   'Band Front Raise Tan',                3, '12 reps',            'Hold the tan band under your feet and raise both arms straight in front of you to shoulder height. Lower slowly with full control.'],
    ['Thursday', 'Block 2',   'Band External Rotation Tan',          3, '12/side',            'Anchor the tan band at elbow height, keep your elbow bent at 90 degrees and pinned to your side, and rotate your forearm outward. Return with control.'],
    ['Thursday', 'Finisher',  'Row Burnout',                         1, 'Max reps',           'Use the blue band anchored at waist height and sit upright. Row both hands to your waist repeatedly to failure - keep your back straight and breathe steadily throughout.'],
    ['Thursday', 'Cool-Down', 'Doorway Chest Stretch',               1, '30 sec',             'Stand in a doorway with both arms at 90 degrees and lean your body forward gently. Keep your chin tucked and breathe steadily.'],
    ['Thursday', 'Cool-Down', 'Overhead Tricep Stretch',             1, '30 sec/side',        'Reach one arm straight overhead, bend at the elbow, and gently press the elbow with your other hand. Keep your neck long and breathe steadily.'],
    ['Thursday', 'Cool-Down', 'Wrist Flexor Stretch',                1, '20 sec/side',        'Extend one arm straight with palm facing up and gently pull the fingers back with your other hand. Hold steadily and breathe.']

  ]);

  SpreadsheetApp.flush();
  Logger.log('Done - all 4 Phase 2 Strength tabs reformatted successfully.');
}
