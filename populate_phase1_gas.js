// ============================================================
// TOGETHER APP — Populate Phase 1 Cycle 1 Exercise Data
// Run once from the spreadsheet's Apps Script editor (Extensions > Apps Script)
// Columns: A=Day, B=Block, C=Exercise, D=Sets, E=Target, F=Demo (left blank), G=Note
// ============================================================

function populatePhase1Data() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var data = getPhase1ExerciseData();

  for (var sheetName in data) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log('Sheet not found: ' + sheetName); continue; }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { Logger.log('No data rows in: ' + sheetName); continue; }

    var numRows = lastRow - 1;

    // Single read: columns A and B to locate rows
    var existing = sheet.getRange(2, 1, numRows, 2).getValues();

    // Map "Day|Block" -> array of 0-based row indices within existing[]
    var rowMap = {};
    for (var r = 0; r < existing.length; r++) {
      var day   = String(existing[r][0]).trim();
      var block = String(existing[r][1]).trim();
      if (!day || !block) continue;
      var key = day + '|' + block;
      if (!rowMap[key]) rowMap[key] = [];
      rowMap[key].push(r);
    }

    // Build output array for columns C–G (5 cols), initialized to empty strings
    var output = [];
    for (var i = 0; i < numRows; i++) output.push(['', '', '', '', '']);

    var exercises = data[sheetName];
    var cursor = {};
    for (var i = 0; i < exercises.length; i++) {
      var ex  = exercises[i];
      var key = ex[0] + '|' + ex[1];
      if (!rowMap[key]) { Logger.log('No rows for: ' + key + ' (' + ex[2] + ')'); continue; }
      if (!cursor[key]) cursor[key] = 0;
      var idx = rowMap[key][cursor[key]++];
      if (idx === undefined) { Logger.log('Ran out of rows for: ' + key + ' (' + ex[2] + ')'); continue; }
      output[idx][0] = ex[2]; // C: Exercise
      output[idx][1] = ex[3]; // D: Sets
      output[idx][2] = ex[4]; // E: Target
      // output[idx][3] left blank — F: Demo link (fill manually)
      output[idx][4] = ex[5]; // G: Note
    }

    // Single write — one API call per sheet instead of one per cell
    sheet.getRange(2, 3, numRows, 5).setValues(output);
    Logger.log('Populated: ' + sheetName);
  }

  Logger.log('All done! Phase 1 Cycle 1 data has been populated.');
}

// ============================================================
// EXERCISE DATA
// Format per entry: [day, block, exercise, sets, target, note]
// ============================================================

function getPhase1ExerciseData() {
  return {

    // ─────────────────────────────────────────────
    'Husband Core P1C1': [
      // Monday
      ['Monday','Warm-Up','Thoracic Rotation Stretch',1,'8/side','On all fours, rotate elbow to ceiling. Feel mid-back open.'],
      ['Monday','Warm-Up','90/90 Hip Stretch',1,'45 sec/side','Front and back leg at 90 degrees. Sit tall and breathe.'],
      ['Monday','Warm-Up','Dead Bug Legs Only',1,'10/side','Arms stay vertical. Lower one leg at a time. Low back stays flat.'],
      ['Monday','Warm-Up','Band Pull-Apart',1,'15 reps','Tan band. Pull to chest height. Squeeze shoulder blades together.'],
      ['Monday','Block 1','Stir the Pot',3,'8/direction','Forearms on chair. Small circles. Core locked. Do not let hips move.'],
      ['Monday','Block 1','Copenhagen Plank Knees Down',3,'20 sec/side','Bottom knee on floor. Top leg on chair. Hips level throughout.'],
      ['Monday','Block 1','Pallof Press',3,'12/side','Blue band anchored to door. Press out, hold 2 sec, return. Anti-rotation.'],
      ['Monday','Block 2','Bear Hold',3,'25 sec','Knees hover 1 inch off floor. Back flat. Breathe steadily.'],
      ['Monday','Block 2','Suitcase Hold',3,'30 sec/side','Hold blue band in one hand at side. Resist lateral lean. Stand tall.'],
      ['Monday','Block 2','Reverse Hyper on Table',3,'12 reps','Hips on table edge, legs hang. Lift legs to parallel. Squeeze glutes.'],
      ['Monday','Finisher','Max Bear Hold',1,'To failure','Knees 1 inch off floor. Breathe. Hold until form breaks.'],
      ['Monday','Cool-Down','Pigeon Pose',1,'45 sec/side',''],
      ['Monday','Cool-Down','Thread the Needle',1,'30 sec/side',''],
      ['Monday','Cool-Down','Supine Figure-4 Stretch',1,'30 sec/side',''],

      // Wednesday
      ['Wednesday','Warm-Up','Standing Hip Flexor March',1,'10/side','Standing, drive knee up to hip height. Hold 1 sec each rep.'],
      ['Wednesday','Warm-Up','Quadruped Reach',1,'10/side','On all fours, extend opposite arm and leg. Hold 2 sec.'],
      ['Wednesday','Warm-Up','Scapular Push-Up',1,'12 reps','In plank, pinch and spread shoulder blades. Arms stay straight.'],
      ['Wednesday','Warm-Up','Thoracic Extension on Floor',1,'10 reps','Sit on heels, hands behind head, extend upper back toward ceiling.'],
      ['Wednesday','Block 1','Half-Kneeling Band Chop',3,'10/side','Green band anchored high. Pull diagonally across body. Resist rotation.'],
      ['Wednesday','Block 1','Dead Bug with Band Reach',3,'8/side','Tan band overhead. Lower opposite leg while reaching band up. Core flat.'],
      ['Wednesday','Block 1','Seated Band Anti-Rotation',3,'12/side','Sit upright. Green band at chest height. Press forward, resist twist.'],
      ['Wednesday','Block 2','Side-Lying Hip Abduction',3,'15/side','Lie on side. Lift top leg to 45 degrees. Slow and controlled.'],
      ['Wednesday','Block 2','Prone Hip Extension',3,'12/side','Lying face down. Lift one straight leg off floor. Squeeze glute.'],
      ['Wednesday','Block 2','Standing Band Lift',3,'10/side','Green band anchored low. Pull diagonally up across body. Anti-rotation.'],
      ['Wednesday','Finisher','Plank Hold with Reach',1,'To failure','In forearm plank, alternately reach one arm forward. Hips stay still.'],
      ['Wednesday','Cool-Down','Lying Knee-to-Chest',1,'45 sec/side',''],
      ['Wednesday','Cool-Down','Cat-Cow',1,'10 reps',''],
      ['Wednesday','Cool-Down','Doorway Chest Opener',1,'30 sec',''],

      // Friday
      ['Friday','Warm-Up','Supine Leg Lowering',1,'10 reps','Both legs vertical. Lower slowly to floor. Low back stays down.'],
      ['Friday','Warm-Up','Glute Bridge with Hold',1,'5 reps','Bridge up and hold 5 sec each rep. Feel posterior chain activate.'],
      ['Friday','Warm-Up','Band Chest Pull',1,'12 reps','Tan band. Hold in front at shoulder height, pull wide. Squeeze back.'],
      ['Friday','Warm-Up','Hip 90/90 Switch',1,'8/side','Seated on floor. Rotate between 90/90 positions each side.'],
      ['Friday','Block 1','Tall Kneeling Band Row',3,'12/side','Kneel on both knees. Blue band. Row to hip. No leaning back.'],
      ['Friday','Block 1','Side Plank with Thread',3,'10/side','In side plank, reach top arm under body and back up. Controlled.'],
      ['Friday','Block 1','Glute Bridge with Band',3,'15 reps','Blue band above knees. Bridge up. Push knees out against band.'],
      ['Friday','Block 2','Quadruped Hip Circle',3,'8/direction','On all fours. Draw large circle with one knee. Keep hips level.'],
      ['Friday','Block 2','Prone Cobra Hold',3,'30 sec','Face down, arms at sides. Lift chest and hands off floor. Hold.'],
      ['Friday','Block 2','Band Clamshell',3,'15/side','Green band above knees. Lie on side. Open top knee like a clam.'],
      ['Friday','Finisher','Full Body Tension Hold',1,'To failure','Stand tall. Squeeze every muscle simultaneously. Breathe through it.'],
      ['Friday','Cool-Down','Supine Spinal Twist',1,'40 sec/side',''],
      ['Friday','Cool-Down','Kneeling Hip Flexor Stretch',1,'30 sec/side',''],
      ['Friday','Cool-Down',"Child's Pose",1,'45 sec',''],
    ],

    // ─────────────────────────────────────────────
    'Husband Arms P1C1': [
      // Tuesday
      ['Tuesday','Warm-Up','Band Chest Stretch',1,'30 sec/side','Tan band anchored. Step forward, stretch chest open.'],
      ['Tuesday','Warm-Up','Wall Forearm Stretch',1,'20 sec/side','Forearm against wall. Gently rotate body away. Stretches bicep.'],
      ['Tuesday','Warm-Up','Scapular Wall Slides',1,'10 reps','Back against wall. Slide arms up and down in W shape.'],
      ['Tuesday','Biceps','Band Curl Isometric Hold',3,'30 sec/side','Green band. Curl to 90 degrees, hold still. Maximum tension.'],
      ['Tuesday','Biceps','Inverted Row Wide Grip',3,'10 reps','Table rows with hands wider than shoulder width. Full range.'],
      ['Tuesday','Biceps','Band Hammer Curl',3,'12/side','Green band. Palms face each other. Slow and controlled.'],
      ['Tuesday','Triceps','Band Tricep Overhead Ext',3,'12 reps','Green band anchored low. Both hands overhead. Extend fully.'],
      ['Tuesday','Triceps','Tricep Push-Up Hold',3,'20 sec','Lower halfway down. Hold. Elbows track back along ribs.'],
      ['Tuesday','Triceps','Bench Dip Negative',3,'6 reps','Chair dip. Take 5 seconds to lower. Press back up normally.'],
      ['Tuesday','Finisher','Band Curl Burnout',1,'Max reps','Green band. Full curls to failure. No swinging.'],
      ['Tuesday','Cool-Down','Bicep Wall Stretch',1,'30 sec/side',''],
      ['Tuesday','Cool-Down','Tricep Overhead Stretch',1,'30 sec/side',''],
      ['Tuesday','Cool-Down','Wrist Flexor Stretch',1,'20 sec/side',''],

      // Thursday
      ['Thursday','Warm-Up','Shoulder CARs',1,'5/side','Slow full arm circle. Feel the joint move through full range.'],
      ['Thursday','Warm-Up','Band Face Pull',1,'15 reps','Green band at face height. Pull to nose. Squeeze rear delts.'],
      ['Thursday','Warm-Up','Prone Shoulder Tap',1,'10 reps','In push-up position. Tap opposite shoulder. Hips still.'],
      ['Thursday','Shoulders','Band Overhead Press',3,'12 reps','Blue band under feet. Press overhead. Full lockout at top.'],
      ['Thursday','Shoulders','Pike Push-Up Hold',3,'20 sec','In pike position, lower halfway. Hold. Feel shoulder burn.'],
      ['Thursday','Shoulders','Band Lateral Raise',3,'12/side','Green band. Single arm lateral raise. Lead with elbow.'],
      ['Thursday','Upper Back','Band Reverse Fly',3,'12 reps','Green band. Hinge forward. Pull arms wide. Squeeze rear delts.'],
      ['Thursday','Upper Back','Inverted Row Underhand',3,'10 reps','Table rows palms facing up. Targets biceps and upper back.'],
      ['Thursday','Upper Back','Band High Row',3,'12 reps','Blue band anchored high. Pull to chin. Elbows flare wide.'],
      ['Thursday','Finisher','Shoulder Isometric Burnout',1,'Max hold','Green band. Hold arms at 90 degrees at sides. Hold to failure.'],
      ['Thursday','Cool-Down','Doorway Chest Stretch',1,'30 sec',''],
      ['Thursday','Cool-Down','Cross-Body Shoulder Stretch',1,'30 sec/side',''],
      ['Thursday','Cool-Down','Neck Side Stretch',1,'20 sec/side',''],

      // Saturday
      ['Saturday','Warm-Up','Band Chest Pass',1,'15 reps','Tan band anchored behind. Punch arms forward. Full extension.'],
      ['Saturday','Warm-Up','Elbow Circles',1,'10/direction','Hands on shoulders. Draw large circles with elbows. Both directions.'],
      ['Saturday','Warm-Up','Push-Up Plus',1,'10 reps','At top of push-up, push extra through shoulder blades. Serratus activation.'],
      ['Saturday','Power','Band Curl to Press',3,'10 reps','Green band. Curl up then press overhead. Control down. Full chain.'],
      ['Saturday','Power','Typewriter Push-Up',3,'6/side','Push up, then slide to one side, then the other. Slow and controlled.'],
      ['Saturday','Power','Band Pull-Down',3,'12 reps','Blue band anchored high. Pull straight down to hips. Lats engaged.'],
      ['Saturday','Endurance','Push-Up with Hold',3,'8 reps','Lower down, hold at bottom 2 sec, press up. Full control.'],
      ['Saturday','Endurance','Band Tricep Pushdown',3,'15 reps','Blue band anchored high. Push down to full extension. Slow up.'],
      ['Saturday','Endurance','Inverted Row with Pause',3,'8 reps','Table row. Pause at top 2 sec. Squeeze shoulder blades.'],
      ['Saturday','Finisher','Push-Up Position Hold',1,'To failure','Top of push-up position. Hold perfectly still. Every muscle engaged.'],
      ['Saturday','Cool-Down','Overhead Tricep Stretch',1,'30 sec/side',''],
      ['Saturday','Cool-Down','Chest Opener',1,'30 sec',''],
      ['Saturday','Cool-Down','Wrist Extensor Stretch',1,'20 sec/side',''],
    ],

    // ─────────────────────────────────────────────
    'Wife Core P1C1': [
      // Monday
      ['Monday','Warm-Up','Diaphragmatic Breathing',1,'8 breaths','Inhale through nose, belly rises. Exhale slowly. Never hold breath.'],
      ['Monday','Warm-Up','Supine Knee Rocks',1,'12 reps','Lying down. Both knees bent. Rock gently side to side. Breathe.'],
      ['Monday','Warm-Up','Cat-Camel Stretch',1,'10 reps','Slow and fluid. Feel each vertebra move.'],
      ['Monday','Warm-Up','Glute Bridge March',1,'10/side','Bridge up. Alternate lifting each knee. Keep hips level.'],
      ['Monday','Block 1','Quadruped Breathing',3,'8 breaths','On all fours. Inhale, let belly drop. Exhale, draw navel up gently.'],
      ['Monday','Block 1','Supine Leg Slides',3,'10/side','Lying flat. Slide one foot out straight. Core gently engaged. Breathe.'],
      ['Monday','Block 1','Kneeling Hip Hinge',3,'10 reps','Kneel tall. Hinge forward from hips. Return upright. Back stays flat.'],
      ['Monday','Block 1','Side-Lying Clam',3,'12/side','Green band above knees. Lie on side. Open top knee. Pelvic floor safe.'],
      ['Monday','Block 2','Supported Side Plank',3,'20 sec/side','Both knees down. Hold side plank position. Breathe steadily.'],
      ['Monday','Block 2','Standing Core Squeeze',3,'15 reps','Stand tall. Exhale and draw navel in. Hold 3 sec. Release.'],
      ['Monday','Block 2','Supine Hip Lift',3,'12 reps','Knees bent, feet flat. Tilt pelvis and lift hips slightly. Controlled.'],
      ['Monday','Block 2','Seated Band Row',3,'12 reps','Tan band. Sit tall. Row to waist. Keep shoulders down. Breathe out.'],
      ['Monday','Finisher','Supported Bear Hold',1,'40 sec','Knees hover 1 inch off floor. Back flat. Breathe throughout. No holding breath.'],
      ['Monday','Cool-Down',"Child's Pose",1,'45 sec',''],
      ['Monday','Cool-Down','Supine Figure-4 Stretch',1,'30 sec/side',''],
      ['Monday','Cool-Down','Gentle Spinal Twist',1,'30 sec/side','Knees together, fall to one side. No forcing the rotation.'],

      // Wednesday
      ['Wednesday','Warm-Up','Diaphragmatic Breathing',1,'6 breaths','Belly expands on inhale. Never hold breath.'],
      ['Wednesday','Warm-Up','Standing Shoulder Rolls',1,'10/direction','Roll shoulders forward then backward. Slow and relaxed.'],
      ['Wednesday','Warm-Up','Lying Hip Circles',1,'8/direction','Lying on back, knees bent. Draw circles with knees together.'],
      ['Wednesday','Warm-Up','Standing Glute Squeeze',1,'15 reps','Stand tall. Squeeze both glutes. Hold 2 sec. Release.'],
      ['Wednesday','Block 1','Standing Side Bend',3,'10/side','Stand tall. Reach one arm overhead, lean to opposite side. Breathe.'],
      ['Wednesday','Block 1','Seated Band Twist',3,'10/side','Sit upright. Tan band. Rotate torso gently each side. No jerking.'],
      ['Wednesday','Block 1','Hip Hinge with Band',3,'12 reps','Green band under feet. Hinge forward, band resists. Back flat throughout.'],
      ['Wednesday','Block 1','Standing Hip Abduction',3,'12/side','Tan band above ankles. Stand on one leg. Lift other leg to side slowly.'],
      ['Wednesday','Block 2','Quadruped Leg Slide',3,'10/side','On all fours. Slide one leg back along floor. Keep hips level. Breathe.'],
      ['Wednesday','Block 2','Modified Clamshell',3,'15/side','Tan band. Lie on side. Open knee. Slow controlled movement.'],
      ['Wednesday','Block 2','Wall Push Hold',3,'20 sec','Hands on wall. Push firmly. Engage core. Breathe throughout.'],
      ['Wednesday','Block 2','Seated Leg Press with Band',3,'12/side','Sit on floor. Tan band around foot. Press leg forward. Core engaged.'],
      ['Wednesday','Finisher','Standing Balance Hold',1,'30 sec/side','Stand on one leg. Arms at sides. Breathe. Hold steady each side.'],
      ['Wednesday','Cool-Down','Lying Spinal Twist',1,'40 sec/side','Gentle. Never force. Breathe throughout.'],
      ['Wednesday','Cool-Down','Kneeling Hip Flexor Stretch',1,'30 sec/side',''],
      ['Wednesday','Cool-Down',"Child's Pose",1,'30 sec',''],

      // Friday
      ['Friday','Warm-Up','Diaphragmatic Breathing',1,'8 breaths','End of week reset. Breathe deeply before moving.'],
      ['Friday','Warm-Up','Pelvic Tilts Standing',1,'12 reps','Stand with back against wall. Tilt pelvis forward and back gently.'],
      ['Friday','Warm-Up','Hip Circle Standing',1,'8/direction','Hands on hips. Draw large circles. Loosen hip flexors.'],
      ['Friday','Warm-Up','Shoulder Opener Stretch',1,'30 sec/side','Hold one arm across chest. Other hand gently assists. Breathe.'],
      ['Friday','Block 1','Standing March with Hold',3,'10/side','Drive knee up. Hold 2 sec. Lower slow. Hands on hips. Breathe.'],
      ['Friday','Block 1','Supported Single Leg Stand',3,'20 sec/side','Hand on wall for support. Stand on one leg. Core gently engaged.'],
      ['Friday','Block 1','Band Hip Hinge',3,'12 reps','Green band under feet. Hinge at hips. Feel hamstrings stretch. Return.'],
      ['Friday','Block 1','Standing Cat-Cow',3,'10 reps','Stand with hands on knees. Arch and round the spine. Breathe.'],
      ['Friday','Block 2','Seated Ab Brace',3,'15 reps','Sit upright. Inhale. On exhale, brace core gently. Hold 3 sec.'],
      ['Friday','Block 2','Side-Lying Leg Lift',3,'12/side','Lie on side. Lift top leg to 45 degrees. Slow up and slow down.'],
      ['Friday','Block 2','Quadruped Hold',3,'20 sec','On all fours. Hold perfectly still. Back flat. Breathe throughout.'],
      ['Friday','Block 2','Standing Band Pull-Down',3,'12 reps','Tan band anchored high. Pull down to hips. Engage core. Breathe.'],
      ['Friday','Finisher','Breathing Wall Sit',1,'40 sec','Sit against wall at 90 degrees. Breathe steadily. No holding breath.'],
      ['Friday','Cool-Down','Happy Baby Pose',1,'30 sec',''],
      ['Friday','Cool-Down','Supine Knee Hug',1,'45 sec','Both knees to chest. Rock gently side to side.'],
      ['Friday','Cool-Down','Lying Spinal Twist',1,'40 sec/side',''],
    ],

    // ─────────────────────────────────────────────
    'Wife Bands P1C1': [
      // Tuesday
      ['Tuesday','Warm-Up','Arm Circles',1,'15/direction','Forward then backward. Slow and controlled.'],
      ['Tuesday','Warm-Up','Band Chest Stretch',1,'30 sec','Tan band. Hold wide in front. Pull gently behind back. Open chest.'],
      ['Tuesday','Warm-Up','Wrist Circles',1,'10/direction','Both directions. Protect wrists before band work.'],
      ['Tuesday','Warm-Up','Neck Rolls',1,'5/direction','Slow gentle half circles. Release neck and upper trap tension.'],
      ['Tuesday','Block 1','Band Overhead Press Seated',3,'12 reps','Sit tall. Green band under feet. Press overhead. Full extension.'],
      ['Tuesday','Block 1','Band Front Raise Single Arm',3,'12/side','Tan band. One arm at a time. Raise to shoulder height. Control down.'],
      ['Tuesday','Block 2','Band Upright Row',3,'12 reps','Green band under feet. Pull to chin. Elbows lead upward.'],
      ['Tuesday','Block 2','Band External Rotation',3,'12/side','Tan band. Elbow at side, bent 90 degrees. Rotate forearm outward.'],
      ['Tuesday','Finisher','Shoulder Circuit',1,'1 round','Tan band. 10 Front Raises, 10 Upright Rows, 10 External Rotations. No rest.'],
      ['Tuesday','Cool-Down','Overhead Tricep Stretch',1,'30 sec/side',''],
      ['Tuesday','Cool-Down','Cross-Body Shoulder Stretch',1,'20 sec/side',''],
      ['Tuesday','Cool-Down','Chest Opener',1,'30 sec','Clasp hands behind back. Gently lift and open chest.'],

      // Thursday
      ['Thursday','Warm-Up','Arm Circles',1,'15/direction',''],
      ['Thursday','Warm-Up','Scapular Squeezes',1,'12 reps','Arms at sides. Squeeze shoulder blades together. Hold 2 sec.'],
      ['Thursday','Warm-Up','Wrist Circles',1,'10/direction',''],
      ['Thursday','Warm-Up','Shoulder Shrugs',1,'10 reps','Shrug both shoulders up to ears. Hold 1 sec. Release slowly.'],
      ['Thursday','Block 1','Band Tricep Kickback',3,'12/side','Hinge slightly forward. Green band. Extend arm back fully. Squeeze.'],
      ['Thursday','Block 1','Band Rear Delt Fly',3,'12 reps','Tan band. Hinge forward. Pull arms wide. Squeeze rear shoulder.'],
      ['Thursday','Block 2','Band Bicep Curl Slow',3,'10 reps','Green band. 2 sec up, 3 sec down. Full range. No swinging.'],
      ['Thursday','Block 2','Band Overhead Tricep Ext',3,'12 reps','Green band anchored low. Both hands. Extend overhead. Control down.'],
      ['Thursday','Finisher','Row Burnout',1,'Max reps','Blue band. Seated row. Full range. Go to failure. Breathe throughout.'],
      ['Thursday','Cool-Down','Doorway Chest Stretch',1,'30 sec',''],
      ['Thursday','Cool-Down','Overhead Tricep Stretch',1,'30 sec/side',''],
      ['Thursday','Cool-Down','Wrist Flexor Stretch',1,'20 sec/side',''],
    ],

  };
}
