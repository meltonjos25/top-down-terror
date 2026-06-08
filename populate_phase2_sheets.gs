// =====================================================================
// Together App — Phase 2 Strength Sheet Populator
// Paste this entire file into a new Google Apps Script project,
// then run populatePhase2Sheets() once.
// =====================================================================

var SHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';

var BLOCK_COLORS = {
  'Warm-Up':   '#e8f5e9',
  'Block 1':   '#e3f2fd',
  'Block 2':   '#fce4ec',
  'Finisher':  '#fff8e1',
  'Cool-Down': '#e0f7fa'
};

function populatePhase2Sheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  var sheets = [
    { name: 'Husband Core P2C1', data: HUSBAND_CORE_P2 },
    { name: 'Husband Arms P2C1', data: HUSBAND_ARMS_P2 },
    { name: 'Wife Core P2C1',    data: WIFE_CORE_P2    },
    { name: 'Wife Bands P2C1',   data: WIFE_BANDS_P2   }
  ];

  sheets.forEach(function(def) {
    createSheet(ss, def.name, def.data);
  });

  SpreadsheetApp.flush();
  Logger.log('Done — all 4 Phase 2 Strength sheets created.');
}

function createSheet(ss, sheetName, data) {
  var existing = ss.getSheetByName(sheetName);
  if (existing) ss.deleteSheet(existing);

  var sheet = ss.insertSheet(sheetName);

  // Build rows: [Day, Block, Exercise, Sets, Target Reps/Time, Exercise Demo Link, Coaching Note]
  var rows = [['Day', 'Block', 'Exercise', 'Sets', 'Target Reps/Time', 'Exercise Demo Link', 'Coaching Note']];
  data.forEach(function(ex) {
    rows.push([ex.day, ex.block, ex.name, ex.sets, ex.target, '', ex.note]);
  });

  sheet.getRange(1, 1, rows.length, 7).setValues(rows);

  // Header row styling
  var header = sheet.getRange(1, 1, 1, 7);
  header.setBackground('#263238');
  header.setFontColor('#ffffff');
  header.setFontWeight('bold');
  header.setFontSize(11);

  // Data row styling
  for (var i = 1; i < rows.length; i++) {
    var block = rows[i][1];
    var rowNum = i + 1;
    var blockColor = BLOCK_COLORS[block] || '#ffffff';
    // Day column: phase-2 orange tint, bold
    sheet.getRange(rowNum, 1).setBackground('#fff3e0').setFontWeight('bold');
    // Block column: block color, bold
    sheet.getRange(rowNum, 2).setBackground(blockColor).setFontWeight('bold');
    // Remaining columns: same block color
    sheet.getRange(rowNum, 3, 1, 5).setBackground(blockColor);
  }

  // Column widths
  sheet.setColumnWidth(1, 110); // Day
  sheet.setColumnWidth(2, 130); // Block
  sheet.setColumnWidth(3, 270); // Exercise
  sheet.setColumnWidth(4,  60); // Sets
  sheet.setColumnWidth(5, 200); // Target Reps/Time
  sheet.setColumnWidth(6, 100); // Exercise Demo Link
  sheet.setColumnWidth(7, 330); // Coaching Note

  sheet.setFrozenRows(1);
  sheet.getRange(1, 4, rows.length, 1).setHorizontalAlignment('center');

  Logger.log('Created: ' + sheetName + ' (' + (rows.length - 1) + ' rows)');
}

// =====================================================================
// HUSBAND CORE P2C1  —  Mon / Wed / Fri
// =====================================================================
var HUSBAND_CORE_P2 = [
  // ── Monday ──
  {day:'Monday', block:'Warm-Up',   name:'Cat-Camel Stretch',              sets:1, target:'10 reps',           note:''},
  {day:'Monday', block:'Warm-Up',   name:'Dead Bug Arms Only',             sets:1, target:'10/side',            note:''},
  {day:'Monday', block:'Warm-Up',   name:'Band Pull-Apart Green',          sets:1, target:'15 reps',            note:''},
  {day:'Monday', block:'Warm-Up',   name:'Hip Flexor Lunge Stretch',       sets:1, target:'30 sec/side',        note:''},
  {day:'Monday', block:'Block 1',   name:'Copenhagen Plank Knee Down',     sets:3, target:'20 sec/side',        note:''},
  {day:'Monday', block:'Block 1',   name:'Hollow Body Hold',               sets:3, target:'40 sec',             note:''},
  {day:'Monday', block:'Block 1',   name:'L-Sit Hold on Chairs',           sets:3, target:'20 sec',             note:''},
  {day:'Monday', block:'Block 2',   name:'Dragon Flag Negative',           sets:3, target:'5 reps slow',        note:''},
  {day:'Monday', block:'Block 2',   name:'Stir the Pot on Forearms',       sets:3, target:'8/direction',        note:''},
  {day:'Monday', block:'Block 2',   name:'Band Pallof Press Blue',         sets:3, target:'12/side',            note:''},
  {day:'Monday', block:'Finisher',  name:'Max L-Sit Hold on Chairs',       sets:1, target:'1 set to failure',   note:''},
  {day:'Monday', block:'Cool-Down', name:"Child's Pose",                   sets:1, target:'45 sec',             note:''},
  {day:'Monday', block:'Cool-Down', name:'Cobra Stretch',                  sets:1, target:'30 sec',             note:''},
  {day:'Monday', block:'Cool-Down', name:'Seated Spinal Twist',            sets:1, target:'30 sec/side',        note:''},
  // ── Wednesday ──
  {day:'Wednesday', block:'Warm-Up',   name:'Hip Circles',                    sets:1, target:'10/direction',       note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Lying Knee Rocks',               sets:1, target:'12 reps',            note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Bear Crawl In Place',            sets:1, target:'30 sec',             note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Inchworm Walk-Out',              sets:1, target:'6 reps',             note:''},
  {day:'Wednesday', block:'Block 1',   name:'Slow Negative Push-Up',          sets:3, target:'8 reps',             note:'4 sec down'},
  {day:'Wednesday', block:'Block 1',   name:'Band Resisted Crunch Blue',      sets:3, target:'15 reps',            note:''},
  {day:'Wednesday', block:'Block 1',   name:'Jefferson Curl with Band Tan',   sets:3, target:'8 reps',             note:''},
  {day:'Wednesday', block:'Block 2',   name:'Side Plank with Band Row Blue',  sets:3, target:'10/side',            note:''},
  {day:'Wednesday', block:'Block 2',   name:'Hollow Body Rock',               sets:3, target:'35 sec',             note:''},
  {day:'Wednesday', block:'Block 2',   name:'Dead Bug with Band Tan',         sets:3, target:'8/side',             note:''},
  {day:'Wednesday', block:'Finisher',  name:'Max Slow Negative Push-Ups',     sets:1, target:'1 set to failure',   note:''},
  {day:'Wednesday', block:'Cool-Down', name:'Lying Spinal Twist',             sets:1, target:'45 sec/side',        note:''},
  {day:'Wednesday', block:'Cool-Down', name:'Figure-4 Stretch',               sets:1, target:'30 sec/side',        note:''},
  {day:'Wednesday', block:'Cool-Down', name:"Child's Pose",                   sets:1, target:'30 sec',             note:''},
  // ── Friday ──
  {day:'Friday', block:'Warm-Up',   name:'Cat-Camel Stretch',              sets:1, target:'10 reps',           note:''},
  {day:'Friday', block:'Warm-Up',   name:'Band Pull-Apart Green',          sets:1, target:'15 reps',           note:''},
  {day:'Friday', block:'Warm-Up',   name:'Hip Flexor Lunge Stretch',       sets:1, target:'30 sec/side',       note:''},
  {day:'Friday', block:'Warm-Up',   name:'Arm Circles',                    sets:1, target:'10 each direction', note:''},
  {day:'Friday', block:'Block 1',   name:'Copenhagen Plank Full',          sets:3, target:'25 sec/side',       note:''},
  {day:'Friday', block:'Block 1',   name:'Dragon Flag Negative',           sets:3, target:'6 reps',            note:''},
  {day:'Friday', block:'Block 1',   name:'Stir the Pot',                   sets:3, target:'10/direction',      note:''},
  {day:'Friday', block:'Block 2',   name:'Band Resisted Hollow Body Blue', sets:3, target:'30 sec',            note:''},
  {day:'Friday', block:'Block 2',   name:'L-Sit Hold on Chairs',           sets:3, target:'25 sec',            note:''},
  {day:'Friday', block:'Block 2',   name:'Band Pallof Press Blue',         sets:3, target:'12/side',           note:''},
  {day:'Friday', block:'Finisher',  name:'Strength Core Gauntlet',         sets:1, target:'1 round',           note:'Copenhagen 20s/side -> Dragon Flag 5 -> L-Sit 20s, no rest'},
  {day:'Friday', block:'Cool-Down', name:"Child's Pose",                   sets:1, target:'45 sec',            note:''},
  {day:'Friday', block:'Cool-Down', name:'Supine Hamstring Stretch',       sets:1, target:'30 sec/side',       note:''},
  {day:'Friday', block:'Cool-Down', name:'Cobra to Upward Dog',            sets:1, target:'30 sec',            note:''},
];

// =====================================================================
// HUSBAND ARMS P2C1  —  Tue / Thu / Sat
// =====================================================================
var HUSBAND_ARMS_P2 = [
  // ── Tuesday ──
  {day:'Tuesday', block:'Warm-Up',   name:'Arm Circles',                       sets:1, target:'15 each direction', note:''},
  {day:'Tuesday', block:'Warm-Up',   name:'Band Pull-Apart Green',             sets:1, target:'15 reps',           note:''},
  {day:'Tuesday', block:'Warm-Up',   name:'Wrist Circles',                     sets:1, target:'10 each direction', note:''},
  {day:'Tuesday', block:'Warm-Up',   name:'Shoulder Cross-Body Stretch',       sets:1, target:'20 sec/side',       note:''},
  {day:'Tuesday', block:'Block 1',   name:'Slow Negative Table Inverted Row',  sets:3, target:'8 reps',            note:'4 sec down'},
  {day:'Tuesday', block:'Block 1',   name:'Band Curl Blue',                    sets:3, target:'12 reps',           note:''},
  {day:'Tuesday', block:'Block 1',   name:'Isometric Table Inverted Row Hold', sets:3, target:'30 sec',            note:''},
  {day:'Tuesday', block:'Block 2',   name:'Slow Negative Close-Grip Push-Up',  sets:3, target:'8 reps',            note:'4 sec down'},
  {day:'Tuesday', block:'Block 2',   name:'Band Tricep Pushdown Blue',         sets:3, target:'15 reps',           note:''},
  {day:'Tuesday', block:'Block 2',   name:'Tricep Dip Negative',               sets:3, target:'8 reps',            note:'4 sec down'},
  {day:'Tuesday', block:'Finisher',  name:'Max Slow Negative Push-Ups',        sets:1, target:'1 set to failure',  note:''},
  {day:'Tuesday', block:'Cool-Down', name:'Overhead Tricep Stretch',           sets:1, target:'30 sec/side',       note:''},
  {day:'Tuesday', block:'Cool-Down', name:'Cross-Body Bicep Stretch',          sets:1, target:'30 sec/side',       note:''},
  // ── Thursday ──
  {day:'Thursday', block:'Warm-Up',   name:'Arm Circles',                      sets:1, target:'15 each direction', note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Band Pull-Apart Green',            sets:1, target:'15 reps',           note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Cat-Camel',                        sets:1, target:'10 reps',           note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Wrist Circles',                    sets:1, target:'10 each direction', note:''},
  {day:'Thursday', block:'Block 1',   name:'Slow Negative Pike Push-Up',       sets:3, target:'8 reps',            note:'4 sec down'},
  {day:'Thursday', block:'Block 1',   name:'Wall Handstand Hold',              sets:3, target:'25 sec',            note:''},
  {day:'Thursday', block:'Block 1',   name:'Band Overhead Press Blue',         sets:3, target:'12 reps',           note:''},
  {day:'Thursday', block:'Block 2',   name:'Typewriter Push-Up',               sets:3, target:'6/side',            note:''},
  {day:'Thursday', block:'Block 2',   name:'Pseudo Planche Lean',              sets:3, target:'20 sec',            note:''},
  {day:'Thursday', block:'Block 2',   name:'Band Face Pull Green',             sets:3, target:'15 reps',           note:''},
  {day:'Thursday', block:'Finisher',  name:'Max Typewriter Push-Ups',          sets:1, target:'1 set to failure',  note:''},
  {day:'Thursday', block:'Cool-Down', name:'Doorway Chest Stretch',            sets:1, target:'30 sec',            note:''},
  {day:'Thursday', block:'Cool-Down', name:"Child's Pose Wide Arms",           sets:1, target:'45 sec',            note:''},
  // ── Saturday ──
  {day:'Saturday', block:'Warm-Up',   name:'Jumping Jacks',                    sets:1, target:'30 sec',            note:''},
  {day:'Saturday', block:'Warm-Up',   name:'Band Pull-Apart Green',            sets:1, target:'15 reps',           note:''},
  {day:'Saturday', block:'Warm-Up',   name:'Inchworm Walk-Out',                sets:1, target:'6 reps',            note:''},
  {day:'Saturday', block:'Warm-Up',   name:'Wrist Circles',                    sets:1, target:'10 each direction', note:''},
  {day:'Saturday', block:'Block 1',   name:'Pseudo Planche Push-Up',           sets:3, target:'8 reps',            note:''},
  {day:'Saturday', block:'Block 1',   name:'Slow Negative Table Inverted Row', sets:3, target:'8 reps',            note:''},
  {day:'Saturday', block:'Block 1',   name:'Band Curl to Press Blue',          sets:3, target:'10 reps',           note:''},
  {day:'Saturday', block:'Block 2',   name:'Typewriter Push-Up',               sets:3, target:'6/side',            note:''},
  {day:'Saturday', block:'Block 2',   name:'Band Tricep Pushdown Blue',        sets:3, target:'15 reps',           note:''},
  {day:'Saturday', block:'Block 2',   name:'Pike Push-Up Negative',            sets:3, target:'8 reps',            note:''},
  {day:'Saturday', block:'Finisher',  name:'Strength Triset',                  sets:1, target:'1 round',           note:'Pseudo Planche Push-Up 8 -> Band Curl Blue 12 -> Typewriter Push-Up 5/side, no rest'},
  {day:'Saturday', block:'Cool-Down', name:'Overhead Tricep Stretch',          sets:1, target:'30 sec/side',       note:''},
  {day:'Saturday', block:'Cool-Down', name:'Doorway Chest Opener',             sets:1, target:'30 sec',            note:''},
  {day:'Saturday', block:'Cool-Down', name:'Wrist Flexor Stretch',             sets:1, target:'20 sec/side',       note:''},
];

// =====================================================================
// WIFE CORE P2C1  —  Mon / Wed / Fri
// =====================================================================
var WIFE_CORE_P2 = [
  // ── Monday ──
  {day:'Monday', block:'Warm-Up',   name:'Diaphragmatic Breathing',              sets:1, target:'8 breaths',          note:''},
  {day:'Monday', block:'Warm-Up',   name:'Pelvic Tilts',                         sets:1, target:'12 reps',            note:''},
  {day:'Monday', block:'Warm-Up',   name:'Cat-Camel',                            sets:1, target:'10 reps',            note:''},
  {day:'Monday', block:'Warm-Up',   name:'Band Pull-Apart Tan',                  sets:1, target:'12 reps',            note:''},
  {day:'Monday', block:'Block 1',   name:'Modified Copenhagen Plank Knee Down',  sets:3, target:'20 sec/side',        note:''},
  {day:'Monday', block:'Block 1',   name:'Hollow Body Hold Bent Knees',          sets:3, target:'25 sec',             note:''},
  {day:'Monday', block:'Block 1',   name:'Bear Hold',                            sets:3, target:'20 sec',             note:''},
  {day:'Monday', block:'Block 2',   name:'Dead Bug with Tan Band',               sets:3, target:'8/side',             note:''},
  {day:'Monday', block:'Block 2',   name:'Slow Heel Slides',                     sets:3, target:'10/side',            note:''},
  {day:'Monday', block:'Block 2',   name:'Seated Band Core Row Tan',             sets:3, target:'12 reps',            note:''},
  {day:'Monday', block:'Finisher',  name:'Max Bear Hold',                        sets:1, target:'1 set to failure',   note:'Breathe throughout'},
  {day:'Monday', block:'Cool-Down', name:"Child's Pose",                         sets:1, target:'45 sec',             note:''},
  {day:'Monday', block:'Cool-Down', name:'Seated Spinal Twist',                  sets:1, target:'30 sec/side',        note:''},
  {day:'Monday', block:'Cool-Down', name:'Knees-to-Chest Hug',                   sets:1, target:'30 sec',             note:''},
  // ── Wednesday ──
  {day:'Wednesday', block:'Warm-Up',   name:'Diaphragmatic Breathing',               sets:1, target:'6 breaths',          note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Hip Circles Standing',                  sets:1, target:'10/direction',        note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Lying Knee Rocks',                      sets:1, target:'12 reps',             note:''},
  {day:'Wednesday', block:'Warm-Up',   name:'Glute Bridge',                          sets:1, target:'12 reps',             note:''},
  {day:'Wednesday', block:'Block 1',   name:'Slow Negative Push-Up Knees Down',      sets:3, target:'8 reps',              note:'4 sec down'},
  {day:'Wednesday', block:'Block 1',   name:'Band Resisted Bird Dog Tan',            sets:3, target:'10/side',             note:''},
  {day:'Wednesday', block:'Block 1',   name:'Side Plank Knees Down Hold',            sets:3, target:'30 sec/side',         note:''},
  {day:'Wednesday', block:'Block 2',   name:'Hollow Body Hold Bent Knees',           sets:3, target:'30 sec',              note:''},
  {day:'Wednesday', block:'Block 2',   name:'Glute Bridge with Band Blue',           sets:3, target:'15 reps',             note:''},
  {day:'Wednesday', block:'Block 2',   name:'Band Pallof Press Tan',                 sets:3, target:'12/side',             note:''},
  {day:'Wednesday', block:'Finisher',  name:'Max Slow Negative Push-Ups Knees Down', sets:1, target:'1 set to failure',    note:''},
  {day:'Wednesday', block:'Cool-Down', name:'Lying Spinal Twist',                    sets:1, target:'40 sec/side',         note:''},
  {day:'Wednesday', block:'Cool-Down', name:'Figure-4 Glute Stretch',                sets:1, target:'30 sec/side',         note:''},
  {day:'Wednesday', block:'Cool-Down', name:"Child's Pose",                          sets:1, target:'30 sec',              note:''},
  // ── Friday ──
  {day:'Friday', block:'Warm-Up',   name:'Diaphragmatic Breathing',              sets:1, target:'8 breaths',          note:''},
  {day:'Friday', block:'Warm-Up',   name:'Pelvic Tilts',                         sets:1, target:'12 reps',            note:''},
  {day:'Friday', block:'Warm-Up',   name:'Cat-Camel',                            sets:1, target:'10 reps',            note:''},
  {day:'Friday', block:'Warm-Up',   name:'Standing Hip Flexor Stretch',          sets:1, target:'30 sec/side',        note:''},
  {day:'Friday', block:'Block 1',   name:'Modified Copenhagen Plank',            sets:3, target:'25 sec/side',        note:''},
  {day:'Friday', block:'Block 1',   name:'Bear Hold',                            sets:3, target:'25 sec',             note:''},
  {day:'Friday', block:'Block 1',   name:'Band Pallof Press Tan',                sets:3, target:'12/side',            note:''},
  {day:'Friday', block:'Block 2',   name:'Dead Bug with Tan Band',               sets:3, target:'10/side',            note:''},
  {day:'Friday', block:'Block 2',   name:'Slow Negative Push-Up Knees Down',     sets:3, target:'8 reps',             note:''},
  {day:'Friday', block:'Block 2',   name:'Glute Bridge with Band Green',         sets:3, target:'15 reps',            note:''},
  {day:'Friday', block:'Finisher',  name:'Strength Flow',                        sets:1, target:'1 round',            note:'Bear Hold 20s -> Modified Copenhagen 20s/side -> Hollow Body Bent Knees 25s, no rest'},
  {day:'Friday', block:'Cool-Down', name:"Child's Pose",                         sets:1, target:'1 min',              note:''},
  {day:'Friday', block:'Cool-Down', name:'Supine Spinal Twist',                  sets:1, target:'40 sec/side',        note:''},
  {day:'Friday', block:'Cool-Down', name:'Happy Baby Pose',                      sets:1, target:'30 sec',             note:''},
];

// =====================================================================
// WIFE BANDS P2C1  —  Tue / Thu
// =====================================================================
var WIFE_BANDS_P2 = [
  // ── Tuesday ──
  {day:'Tuesday',  block:'Warm-Up',   name:'Diaphragmatic Breathing',            sets:1, target:'6 breaths',             note:''},
  {day:'Tuesday',  block:'Warm-Up',   name:'Band Pull-Apart Green',              sets:1, target:'12 reps',               note:''},
  {day:'Tuesday',  block:'Warm-Up',   name:'Hip Circles Standing',               sets:1, target:'10/direction',          note:''},
  {day:'Tuesday',  block:'Warm-Up',   name:'Shoulder Rolls',                     sets:1, target:'10 each direction',     note:''},
  {day:'Tuesday',  block:'Block 1',   name:'Band Bicep Curl Slow Negative Green', sets:3, target:'10 reps',              note:''},
  {day:'Tuesday',  block:'Block 1',   name:'Band Overhead Press Green',          sets:3, target:'10 reps',               note:''},
  {day:'Tuesday',  block:'Block 1',   name:'Band Upright Row Green',             sets:3, target:'10 reps',               note:''},
  {day:'Tuesday',  block:'Block 2',   name:'Band Glute Bridge Slow Blue',        sets:3, target:'12 reps',               note:''},
  {day:'Tuesday',  block:'Block 2',   name:'Band Sumo Squat Green',              sets:3, target:'12 reps',               note:''},
  {day:'Tuesday',  block:'Block 2',   name:'Band Clamshell Slow Green',          sets:3, target:'12/side',               note:''},
  {day:'Tuesday',  block:'Finisher',  name:'Max Band Glute Bridge Hold Blue',    sets:1, target:'1 set hold to failure', note:''},
  {day:'Tuesday',  block:'Cool-Down', name:"Child's Pose",                       sets:1, target:'45 sec',                note:''},
  {day:'Tuesday',  block:'Cool-Down', name:'Seated Spinal Twist',                sets:1, target:'30 sec/side',           note:''},
  // ── Thursday ──
  {day:'Thursday', block:'Warm-Up',   name:'Diaphragmatic Breathing',            sets:1, target:'6 breaths',             note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Band Pull-Apart Green',              sets:1, target:'12 reps',               note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Glute Bridge',                       sets:1, target:'12 reps',               note:''},
  {day:'Thursday', block:'Warm-Up',   name:'Shoulder Rolls',                     sets:1, target:'10 each direction',     note:''},
  {day:'Thursday', block:'Block 1',   name:'Band Row Slow Negative Green',       sets:3, target:'10 reps',               note:''},
  {day:'Thursday', block:'Block 1',   name:'Band Face Pull Green',               sets:3, target:'12 reps',               note:''},
  {day:'Thursday', block:'Block 1',   name:'Band Reverse Fly Tan',               sets:3, target:'12 reps',               note:''},
  {day:'Thursday', block:'Block 2',   name:'Band Squat Slow Blue',               sets:3, target:'10 reps',               note:''},
  {day:'Thursday', block:'Block 2',   name:'Band Romanian Deadlift Green',       sets:3, target:'12 reps',               note:''},
  {day:'Thursday', block:'Block 2',   name:'Band Lateral Walk Green',            sets:3, target:'15 steps/side',         note:''},
  {day:'Thursday', block:'Finisher',  name:'Max Band Squat Slow Blue',           sets:1, target:'1 set to failure',      note:''},
  {day:'Thursday', block:'Cool-Down', name:'Figure-4 Glute Stretch',             sets:1, target:'30 sec/side',           note:''},
  {day:'Thursday', block:'Cool-Down', name:"Child's Pose",                       sets:1, target:'30 sec',                note:''},
];
