const SHEET_ID = '1zqR0CeERQg_8YtgWnQYxiFKAbGcZb4Zvf-R1Iqn46t8';

function doGet(e) {
  var params = (e && e.parameter) ? e.parameter : {};
  var action = params.action || 'ping';
  var ss = SpreadsheetApp.openById(SHEET_ID);

  if (action === 'get_plan') {
    var sheet = ss.getSheetByName(params.sheet);
    if (!sheet) return res({error: 'Sheet not found: ' + params.sheet});

    // Always a live read of the current cell values -- no caching anywhere
    // in this handler, so every request reflects whatever is in the sheet
    // right now, including any edit made seconds ago.
    var rows = sheet.getDataRange().getValues();
    var day = String(params.day || '').trim();

    var exercises = rows
      .slice(2) // row 1 = title banner, row 2 = column headers -- data starts row 3
      .filter(function(r){
        var d = String(r[0] || '').trim();
        return d && d.toLowerCase() !== 'day' && d === day;
      })
      .map(function(r){
        return {
          block:  String(r[1] || ''),
          name:   String(r[2] || ''),
          sets:   parseInt(r[3], 10) || 1,
          target: String(r[4] || ''),
          demo:   String(r[5] || ''),
          note:   String(r[6] || ''),
        };
      });

    return res({exercises: exercises});
  }

  if (action === 'log') {
    var data = JSON.parse(decodeURIComponent(params.data));
    var logSheet = ss.getSheetByName(data.sheet) || ss.insertSheet(data.sheet);
    data.rows.forEach(function(row){ logSheet.appendRow(row); });
    return res({success: true});
  }

  if (action === 'read') {
    var s = ss.getSheetByName(params.sheet);
    var vals = s ? s.getDataRange().getValues().slice(1) : [];
    return res({rows: vals});
  }

  if (action === 'read_all') {
    var s = ss.getSheetByName(params.sheet);
    var vals = s ? s.getDataRange().getValues() : [];
    return res({rows: vals});
  }

  if (action === 'claude_benefit') {
    var data = JSON.parse(decodeURIComponent(params.data));
    var apiKey = (data.apiKey || '').trim();
    var exName = (data.exName || '').trim();
    if (!apiKey || !exName) return res({error: 'Missing apiKey or exName'});
    Logger.log('claude_benefit: keyPrefix=' + apiKey.slice(0,10) + ' keyLen=' + apiKey.length + ' exName=' + exName);
    var prompt = 'For the exercise called ' + exName + ', give me exactly three sentences and nothing else. Sentence 1 must begin with exactly the words Muscles Targeted: followed by the specific named muscles this exercise works such as transverse abdominis, glutes, triceps - never say core or upper body. Sentence 2 must begin with exactly the words Movement Benefit: followed by the specific functional movement pattern this exercise builds. Sentence 3 must begin with exactly the words Sport Application: followed by one specific sport and one specific movement in that sport that this exercise directly improves. Every response must be completely unique and specific to the exact exercise name provided. Never give the same response for two different exercises.';
    try {
      var claudeResp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
        method: 'post',
        contentType: 'application/json',
        headers: { 'x-api-key': apiKey, 'Authorization': 'Bearer ' + apiKey, 'anthropic-version': '2023-06-01' },
        payload: JSON.stringify({
          model: 'claude-sonnet-4-20250514',
          max_tokens: 300,
          messages: [{ role: 'user', content: prompt }]
        }),
        muteHttpExceptions: true
      });
      var code = claudeResp.getResponseCode();
      if (code !== 200) return res({error: 'Claude error ' + code + ': ' + claudeResp.getContentText(), keyPrefix: apiKey.slice(0,10), keyLen: apiKey.length});
      var result = JSON.parse(claudeResp.getContentText());
      var text = '';
      if (result.content && Array.isArray(result.content)) {
        var tb = result.content.find(function(b){ return b.type === 'text'; });
        text = tb ? (tb.text || '').trim() : '';
      }
      return res({text: text});
    } catch(err) {
      return res({error: 'Proxy exception: ' + err.toString()});
    }
  }

  return res({status: 'ok'});
}

function res(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// Run this function ONCE manually after pasting code to grant all required permissions.
// Select "triggerAuth" in the function dropdown and click Run. Accept all permission prompts.
// After it logs "Auth OK" you can deploy. You never need to run it again unless you
// create a brand-new GAS project.
function triggerAuth() {
  SpreadsheetApp.openById(SHEET_ID);
  UrlFetchApp.fetch('https://www.google.com', {muteHttpExceptions: true});
  Logger.log('Auth OK - permissions granted. Now deploy via Manage deployments > Edit > New version > Deploy.');
}
