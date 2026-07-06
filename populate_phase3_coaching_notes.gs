function populatePhase3CoachingNotes() {
  var ss = SpreadsheetApp.openById('1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8');

  function writeNotes(sheetName, notes) {
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { Logger.log('Sheet not found: ' + sheetName); return; }
    var data = notes.map(function(n) { return [n]; });
    sheet.getRange(3, 7, data.length, 1).setValues(data);
    Logger.log('Updated ' + data.length + ' notes in ' + sheetName);
  }


  // ── HUSBAND CORE P3C1 ── 42 rows ──────────────────────────────────────────

  writeNotes('Husband Core P3C1', [

    // Monday — Warm-Up
    'Keep arms fully extended overhead at the top and land softly with slightly bent knees. Breathe steadily and build to your working pace.',
    'Stand with hands on hips and draw large, slow circles in each direction to loosen your hip joints before explosive work.',
    'Press through your heels and squeeze your glutes hard at the top of each rep, holding one second before lowering.',
    'Walk your hands out to a full plank keeping legs as straight as possible, then walk back and stand tall — this primes your core and shoulders.',

    // Monday — Block 1
    'Drive your hands hard into the floor and press so fast that your hands leave the ground. Land with slightly bent elbows and reset fully before the next rep.',
    'Generate maximum force so both hands leave the ground simultaneously. Land with soft elbows and immediately load into the next rep.',
    'Move as fast as possible through each phase and snap your hips through at the top. At the bottom, hold a tight hollow body for one second before driving back up.',

    // Monday — Block 2
    'Drive each knee toward your chest as fast as possible while keeping your hips level and low. Breathe in a steady rhythm — never hold your breath.',
    'Explode up as high as possible from the squat, then shoot your feet back to plank on landing. Land softly through your feet and absorb immediately into the plank.',
    'Drive your torso up fast and reach your hands past your knees, then lower under control. Generate momentum from your core, not from swinging your arms.',

    // Monday — Finisher
    'Move as fast as you can for the full 30 seconds — speed is the goal. Keep your hips level and breathe out on every single push.',

    // Monday — Cool-Down
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back release with each exhale.',
    'Lie face down and press your hands into the floor under your shoulders. Straighten your arms slowly and keep your hips on the floor throughout.',
    'Sit tall and rotate your torso, using a gentle hand press on your knee to deepen the twist. Keep your spine long and breathe slowly.',

    // Wednesday — Warm-Up
    'Drive each knee up to hip height and pump your arms. Breathe steadily and build to your working pace before the main blocks.',
    'Start on all fours with knees hovering one inch off the floor. Keep your back flat like a table and breathe steadily throughout.',
    'Bridge your hips up, then lift one foot at a time without letting your hips drop. Keep your core braced throughout each march.',
    'Keep arms fully extended and make large controlled circles in each direction. This opens up the shoulders before pressing and jumping movements.',

    // Wednesday — Block 1
    'Jump sideways as far as possible, land softly on one foot, then immediately shoot both hands to the floor. Keep the landing quiet and controlled.',
    'Press so hard and fast that you have time to clap before landing. Land with slightly bent elbows to absorb impact and reset before the next rep.',
    'Move your opposite arm and leg simultaneously as fast as possible while keeping your lower back pressed into the floor. Never let speed cause your back to arch.',

    // Wednesday — Block 2
    'Drive both feet off the floor simultaneously and switch legs in the air as fast as possible. Keep your hips level and breathe in a steady rhythm.',
    'Explode up from the squat and land softly, then hold the bottom position for one full second before the next jump. The pause builds power — do not skip it.',
    'Press your lower back into the floor in hollow position and rock forward and back as fast as possible while maintaining total body tension. Breathe steadily throughout.',

    // Wednesday — Finisher
    'Go at maximum speed for each 20-second interval — drop to the floor fast, press up, and jump with arms overhead. The goal is as many reps as possible each round.',

    // Wednesday — Cool-Down
    'Lie on your back and let your knees fall gently to one side. Breathe deeply and let gravity do the work — never force or bounce.',
    'Lie on your back and cross one ankle over the opposite knee. Pull the bottom leg toward you gently and breathe steadily to deepen the stretch.',
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back release fully with each exhale.',

    // Friday — Warm-Up
    'Keep arms fully extended overhead at the top and land softly with slightly bent knees. Breathe steadily and build to your working pace.',
    'Step into a lunge and lower your back knee to the floor. Tuck your pelvis slightly and keep your torso upright to feel the stretch deep in your front hip.',
    'Drive each knee to hip height and pump your arms. This primes your hips and core for the explosive work ahead.',
    'Walk your hands out to a full plank and back with legs as straight as possible. This primes your core and shoulders for the power blocks.',

    // Friday — Block 1
    'Drive your hands into the floor with maximum force and press as fast as possible. Land with slightly bent elbows and reset completely before the next rep.',
    'Move through each phase as fast as possible. At the bottom, snap into a tight hollow body for one second before driving back up.',
    'Drive your torso up explosively and reach past your knees, then lower under full control. Generate the speed from your core, not from momentum.',

    // Friday — Block 2
    'Generate enough force to leave the floor on every rep. Land with soft elbows, absorb immediately, and reload for the next rep.',
    'Drive each knee toward your chest as fast as possible while keeping your hips level and low. Breathe in a steady rhythm throughout.',
    'Explode up as high as possible, shoot your feet back to plank on landing, then jump immediately back up. Keep every landing soft.',

    // Friday — Finisher
    'Move immediately from one exercise to the next with zero rest. Maintain fast, powerful movement throughout and breathe out on every exertion.',

    // Friday — Cool-Down
    'Sink your hips back toward your heels and reach forward. Breathe deeply and let your back fully release with each exhale.',
    'Lie on your back and extend one leg straight up. Hold behind the thigh and breathe steadily, deepening the stretch with each exhale.',
    'From a push-up position press hips toward the floor to cobra, then push hips up and back to upward dog. Move fluidly and breathe throughout.'

  ]);


  // ── HUSBAND ARMS P3C1 ── 40 rows ──────────────────────────────────────────

  writeNotes('Husband Arms P3C1', [

    // Tuesday — Warm-Up
    'Make large circles with straight arms going as fast as you can in each direction. This primes your shoulders for the explosive pushing and pulling ahead.',
    'Keep arms fully extended overhead at the top and land softly with slightly bent knees. Breathe steadily throughout.',
    'Lace your fingers and draw slow circles in both directions to prepare your wrists for explosive pushing and curling movements.',
    'Hold the green band at shoulder height with arms straight and pull both ends apart until it touches your chest. Squeeze your shoulder blades together at the end of each rep.',

    // Tuesday — Block 1
    'Pull yourself up as fast as possible so your chest nearly reaches the table edge. Drive your elbows back hard and squeeze your shoulder blades at the top.',
    'Stand on the blue band and curl both hands to your shoulders as fast as possible. Keep your elbows pinned to your sides — never let them swing forward.',
    'Press so hard and fast that you have time to clap before landing. Land with slightly bent elbows to absorb impact and reset before the next rep.',

    // Tuesday — Block 2
    'Drive your hands into the floor with maximum force and press as fast as possible. Land with slightly bent elbows and reset completely before the next rep.',
    'Push down through your palms so hard that your arms fully extend and your body rises. Land with slightly bent elbows and immediately reload for the next rep.',
    'Lower slowly, then press as fast and hard as possible on the way up — maximum speed in the pressing phase is the goal. Breathe out on every push.',

    // Tuesday — Finisher
    'Go as fast as you can for the full 30 seconds — any push-up variation counts. Drop to your knees only when your form completely breaks, not before.',

    // Tuesday — Cool-Down
    'Reach one arm straight overhead, bend at the elbow, and gently press it with your other hand. Keep your neck long and breathe steadily.',
    'Extend one arm straight across your chest and use your other hand to gently pull it closer. Keep the shoulder of the working arm pressed down, not hunched up.',

    // Thursday — Warm-Up
    'Drive each knee to hip height and pump your arms. Breathe steadily and build to your working pace before the shoulder-intensive work ahead.',
    'Make large circles with straight arms going as fast as you can in each direction to fully warm up the shoulder joints before pressing.',
    'Pull the green band to chest height with arms straight, squeezing your shoulder blades together. This opens your chest before overhead pressing work.',
    'Lace your fingers and draw slow circles in both directions to warm up the wrists before pike push-ups and overhead work.',

    // Thursday — Block 1
    'Start with hips high in a pike position and lower your head toward the floor as fast as possible. Press back up with maximum speed, keeping elbows pointing slightly back.',
    'Feet on a chair, hips high — lower fast then drive up so hard your hands briefly leave the floor. Land with soft elbows and immediately reload.',
    'Stand on the blue band, squat slightly, then jump and press both hands overhead simultaneously in one explosive movement. Land softly and reset before the next rep.',

    // Thursday — Block 2
    'Pull yourself up as fast as possible with maximum force. Drive your elbows back and squeeze your shoulder blades hard at the top of each rep.',
    'Anchor the green band at face height and pull both ends to your temples as fast as possible with elbows flared wide. Squeeze your rear shoulders at the end of every rep.',
    'Lie face down and simultaneously lift your arms and legs as fast as possible. Squeeze your glutes and upper back at the top, then lower with control.',

    // Thursday — Finisher
    'Move immediately from one exercise to the next with zero rest. Maintain maximum speed and breathe out on every exertion.',

    // Thursday — Cool-Down
    'Stand in a doorway with both arms at 90 degrees and lean your body forward gently. Keep your chin tucked and breathe steadily.',
    'Sink your hips back toward your heels and reach your arms as wide as possible on the floor. Breathe deeply and let your lats fully release with each exhale.',

    // Saturday — Warm-Up
    'Keep arms fully extended overhead at the top and land softly with slightly bent knees. Breathe steadily throughout.',
    'Drive each knee to hip height and pump your arms. Build to your working pace before the power work ahead.',
    'Walk your hands out to a full plank and back with legs as straight as possible. This primes your core and shoulders for clapping push-ups and explosive rows.',
    'Pull the green band to chest height with arms straight, squeezing your shoulder blades together at the end of each rep.',

    // Saturday — Block 1
    'Press so hard and fast that you have time to clap before landing. Land with slightly bent elbows to absorb impact and reset completely before the next rep.',
    'Pull yourself up as fast as possible with maximum force. Drive your elbows back hard and squeeze your shoulder blades at the top.',
    'Push down through your palms so hard that your arms fully extend and your body rises. Land with slightly bent elbows and immediately reload.',

    // Saturday — Block 2
    'Move through all 20 reps as fast as possible while keeping your body rigid like a plank. Breathe out on every rep and never let your hips sag.',
    'Stand on the blue band and curl both hands to your shoulders as fast as possible. Keep elbows pinned to your sides throughout all 15 reps.',
    'Move through all 12 reps as fast as possible with hips high throughout. Speed is what makes this a power exercise — commit to it.',

    // Saturday — Finisher
    'Move from exercise to exercise with zero rest. Maintain explosive speed on every rep and breathe out on every push and pull.',

    // Saturday — Cool-Down
    'Reach one arm overhead, bend at the elbow, and gently press it with your other hand. Keep your neck long and breathe steadily.',
    'Stand in a doorway with both arms at 90 degrees and lean forward gently. Breathe steadily and let your chest open up.',
    'Extend one arm straight with palm facing up and gently pull the fingers back with your other hand. Hold steadily without bouncing.'

  ]);


  // ── WIFE CORE P3C1 ── 42 rows ─────────────────────────────────────────────

  writeNotes('Wife Core P3C1', [

    // Monday — Warm-Up
    'Breathe in through your nose letting your belly rise, then exhale fully and gently. Never hold your breath — keep it slow and steady throughout every exercise.',
    'Press through your heels and squeeze your glutes at the top of each rep. Breathe steadily and keep your core gently braced throughout.',
    'Drive each knee to hip height in a controlled march while pumping your arms. Breathe steadily and build to your working pace.',
    'Move slowly between arching and rounding your spine, breathing in as you arch and out as you round. Keep your arms straight throughout.',

    // Monday — Block 1
    'Step back to plank instead of jumping, move as fast as you can, and step back up. Breathe steadily — never hold your breath during any part of the movement.',
    'Step each knee toward your chest as fast as you comfortably can while keeping your hips level and low. Breathe in a steady rhythm throughout.',
    'Extend opposite arm and leg as fast as possible while keeping your hips completely level. Breathe steadily — speed should never cause your lower back to arch.',

    // Monday — Block 2
    'Drive each knee as high and fast as possible while pumping your arms. Breathe steadily and never hold your breath during the march.',
    'Jump only a few inches — focus on a perfectly quiet, controlled landing with soft bent knees. Breathe out as you jump and in as you land.',
    'Slide one heel out quickly until your leg is nearly straight, then return fast. Keep your lower back pressed into the floor and breathe steadily throughout.',

    // Monday — Finisher
    'Move immediately from one exercise to the next. Breathe steadily throughout — if you need to slow down to keep breathing steadily, slow down.',

    // Monday — Cool-Down
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back release with each exhale.',
    'Sit tall and rotate your torso gently, using a hand on your knee to deepen the twist. Keep your spine long and breathe slowly throughout.',
    'Lie on your back and pull both knees to your chest. Gently rock side to side and breathe deeply to release your lower back.',

    // Wednesday — Warm-Up
    'Breathe in through your nose letting your belly rise, then exhale fully. This activates your deep core before the speed work ahead.',
    'Drive each knee to hip height with a steady rhythm. Breathe steadily and build to your working pace.',
    'Stand with hands on hips and draw large, slow circles in each direction to loosen your hip joints before lateral and plank movements.',
    'Press through your heels and squeeze your glutes at the top. Breathe steadily and keep your core gently braced throughout.',

    // Wednesday — Block 1
    'Step sideways and shoot both hands to the floor, holding plank for one second before stepping back. Move as fast as you can and breathe steadily throughout.',
    'From your knees, press up as hard as possible so your hands briefly leave the floor. Land with soft elbows and breathe out on every push.',
    'Move your opposite arm and leg as fast as possible while keeping your lower back pressed firmly into the floor. Never hold your breath — breathe steadily throughout.',

    // Wednesday — Block 2
    'Drive each knee as high and fast as possible while pumping your arms. Breathe steadily — if you need to slow down to breathe, slow down.',
    'Press through your heels and pump your hips up and down as fast as possible. Breathe steadily and never hold your breath throughout all 15 reps.',
    'Step side to side as fast as possible staying on the balls of your feet. Breathe steadily throughout and keep your steps light and quiet.',

    // Wednesday — Finisher
    'Go as hard as you can for each 20-second interval — step back instead of jumping at all times. Breathe steadily throughout and never hold your breath.',

    // Wednesday — Cool-Down
    'Lie on your back and let your knees fall gently to one side. Breathe deeply and let gravity release the stretch — never force it.',
    'Lie on your back and cross one ankle over the opposite knee. Gently pull the bottom leg toward you and breathe steadily to deepen the glute stretch.',
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back release with each exhale.',

    // Friday — Warm-Up
    'Breathe in through your nose letting your belly rise, then exhale fully. Steady breathing is your constant focus — never hold your breath throughout the workout.',
    'Bridge your hips up, then lift one foot off the floor at a time without letting your hips drop. Breathe steadily and keep your core braced throughout.',
    'Drive each knee to hip height with a steady rhythm. Breathe steadily and build to your working pace.',
    'Move slowly between arching and rounding your spine, taking two seconds in each position. Breathe in as you arch and out as you round.',

    // Friday — Block 1
    'Jump only a few inches — the focus is on a perfectly quiet, controlled landing with soft bent knees. Breathe out as you jump and in as you land.',
    'Extend opposite arm and leg as fast as possible while keeping your hips completely level. Breathe steadily — speed should never cause your lower back to arch.',
    'Step back to plank instead of jumping, move as fast as you can, and step back up. Breathe steadily — never hold your breath during any part of this movement.',

    // Friday — Block 2
    'Slide each heel out quickly until your leg is nearly straight, then return fast. Keep your lower back pressed into the floor and breathe steadily throughout.',
    'Step each knee toward your chest as fast as you can while keeping your hips level. Breathe steadily throughout and never hold your breath.',
    'Drive each knee to hip height as fast as you can while pumping your arms. Breathe steadily — if you need to slow down to keep breathing, slow down.',

    // Friday — Finisher
    'Move immediately from one exercise to the next. Breathe steadily throughout — never hold your breath, especially during the jumping and marching movements.',

    // Friday — Cool-Down
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back fully release with each exhale.',
    'Lie on your back and let your knees fall gently to one side. Breathe deeply and let gravity do the work — never force the stretch.',
    'Lie on your back and hold the outside edges of your feet. Gently pull your knees toward your armpits, rock side to side, and breathe deeply.'

  ]);


  // ── WIFE BANDS P3C1 ── 26 rows ────────────────────────────────────────────

  writeNotes('Wife Bands P3C1', [

    // Tuesday — Warm-Up
    'Breathe in through your nose letting your belly rise, then exhale gently. Never hold your breath — steady breathing throughout every exercise today.',
    'Hold the green band at shoulder height with arms straight and pull both ends apart to chest height. Squeeze your shoulder blades together at the end of each rep.',
    'Drive each knee to hip height with a steady rhythm. Breathe steadily and build to your working pace before the band work ahead.',
    'Roll your shoulders in large circles forward then backward. Keep your arms relaxed and your neck long throughout.',

    // Tuesday — Block 1
    'Stand on the green band and curl both hands to your shoulders as fast as possible. Keep your elbows pinned to your sides and breathe steadily — never hold your breath.',
    'Stand on the green band and press both hands overhead as fast as possible to full lockout. Breathe out as you press and in as you lower — never hold your breath.',
    'Hold the tan band under your feet and raise both arms straight in front of you as fast as possible to shoulder height. Breathe steadily and lower with control.',

    // Tuesday — Block 2
    'Use the green band for light resistance and jump only a few inches, landing as quietly as possible with soft bent knees. Breathe out as you jump and in as you land.',
    'Place the blue band just above your knees and push your knees out against it at the top of every rep. Pump your hips up and down as fast as possible and breathe steadily throughout.',
    'Keep the tan band just above your ankles and step side to side as fast as possible in a slight squat. Breathe steadily — never hold your breath during the steps.',

    // Tuesday — Finisher
    'Move immediately from one exercise to the next with zero rest. Breathe steadily throughout — if you need to slow down to keep breathing, slow down.',

    // Tuesday — Cool-Down
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back fully release with each exhale.',
    'Sit tall and rotate your torso gently, using a hand on your knee to deepen the twist. Keep your spine long and breathe slowly throughout.',

    // Thursday — Warm-Up
    'Breathe in through your nose letting your belly rise, then exhale fully and gently. Never hold your breath — this is your foundation for every exercise today.',
    'Hold the green band at shoulder height with arms straight and pull both ends apart to chest height. Squeeze your shoulder blades together at the end of each rep.',
    'Press through your heels and pump your hips up and down quickly. Breathe steadily and squeeze your glutes at the top of each rep.',
    'Roll your shoulders in large circles forward then backward. Keep your arms relaxed and your neck long throughout.',

    // Thursday — Block 1
    'Anchor the green band at waist height and row both hands in as fast as possible, squeezing your shoulder blades together. Breathe out with each pull and in as you extend.',
    'Anchor the green band at face height and pull both ends to your temples as fast as possible with elbows flared wide. Squeeze your rear shoulders on every rep and breathe steadily.',
    'Hold the tan band with both hands, hinge slightly forward, and pull both arms wide as fast as possible. Squeeze your rear shoulder blades together and breathe steadily throughout.',

    // Thursday — Block 2
    'Stand on the green band and perform squats as fast as possible while keeping your chest up. Breathe out as you stand up and in as you lower — never hold your breath.',
    'Loop the tan band around one foot and kick straight back as fast as possible, squeezing your glute at the top. Breathe steadily and keep your lower back from arching.',
    'Loop the tan band around one foot and kick straight back as fast as possible. Breathe steadily throughout and never hold your breath.',

    // Thursday — Finisher
    'Move immediately from one exercise to the next with zero rest. Breathe steadily throughout — if you need to slow down to keep breathing, slow down.',

    // Thursday — Cool-Down
    'Lie on your back and cross one ankle over the opposite knee. Gently pull the bottom leg toward you and breathe steadily to deepen the glute stretch.',
    'Sink your hips back toward your heels and reach your arms forward. Breathe deeply and let your back fully release with each exhale.'

  ]);


  SpreadsheetApp.flush();
  Logger.log('Done — Phase 3 coaching notes written to all 4 tabs. Demo links and all other columns untouched.');
}
