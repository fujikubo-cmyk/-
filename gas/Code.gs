// ============================================================
// 資金繰りメール日次レポート
// Google Apps Script + Claude API + LINE Notify
// ============================================================

var PROPS = PropertiesService.getScriptProperties();

// ---- エントリポイント（毎朝8:00 トリガーで実行） ----
function runDailyReport() {
  var results = fetchCashFlowEmails();

  if (results.length === 0) {
    Logger.log('対象メールなし');
    return;
  }

  var reportItems = [];

  for (var i = 0; i < results.length; i++) {
    var item = results[i];
    try {
      var sheetData = convertExcelToSheetData(item.blob, item.fileName);
      var summary   = summarizeWithClaude(sheetData, item.fileName);
      reportItems.push({
        subject:  item.subject,
        from:     item.from,
        date:     item.date,
        fileName: item.fileName,
        summary:  summary,
        kpi:      extractKpi(sheetData)
      });
    } catch (e) {
      Logger.log('処理エラー [' + item.fileName + ']: ' + e.message);
      reportItems.push({
        subject:  item.subject,
        from:     item.from,
        date:     item.date,
        fileName: item.fileName,
        summary:  '処理中にエラーが発生しました: ' + e.message,
        kpi:      null
      });
    }
  }

  sendEmailReport(reportItems);
  sendLineNotify(reportItems);
}

// ---- Gmail検索 ----
function fetchCashFlowEmails() {
  var query = '(subject:資金繰り OR subject:経理) has:attachment (filename:xlsx OR filename:xls) newer_than:1d';
  var threads = GmailApp.search(query, 0, 20);
  var results = [];

  for (var t = 0; t < threads.length; t++) {
    var messages = threads[t].getMessages();
    var msg = messages[messages.length - 1]; // 最新メッセージ

    var attachments = msg.getAttachments();
    for (var a = 0; a < attachments.length; a++) {
      var att = attachments[a];
      var name = att.getName().toLowerCase();
      if (name.endsWith('.xlsx') || name.endsWith('.xls')) {
        results.push({
          subject:  msg.getSubject(),
          from:     msg.getFrom(),
          date:     Utilities.formatDate(msg.getDate(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm'),
          fileName: att.getName(),
          blob:     att.copyBlob()
        });
      }
    }
  }

  return results;
}

// ---- Excel → Googleスプレッドシート変換してデータ取得 ----
function convertExcelToSheetData(blob, fileName) {
  // DriveにExcelをアップロード
  var excelFile = DriveApp.createFile(blob.setName(fileName));

  // スプレッドシートとしてコピー（Drive API v3）
  var resource = {
    name:     fileName + '_tmp',
    mimeType: MimeType.GOOGLE_SHEETS,
    parents:  [excelFile.getParents().next().getId()]
  };
  var copied = Drive.Files.copy({ name: resource.name, mimeType: resource.mimeType }, excelFile.getId());
  var ssId = copied.id;

  var ss = SpreadsheetApp.openById(ssId);
  var allData = [];

  try {
    var sheets = ss.getSheets();
    for (var s = 0; s < sheets.length; s++) {
      var sheet = sheets[s];
      var values = sheet.getDataRange().getValues();
      allData.push('=== シート: ' + sheet.getName() + ' ===');
      for (var r = 0; r < values.length; r++) {
        allData.push(values[r].map(function(c) {
          if (c instanceof Date) return Utilities.formatDate(c, 'Asia/Tokyo', 'yyyy/MM/dd');
          return String(c);
        }).join('\t'));
      }
    }
  } finally {
    // 一時ファイルを削除
    DriveApp.getFileById(excelFile.getId()).setTrashed(true);
    DriveApp.getFileById(ssId).setTrashed(true);
  }

  var text = allData.join('\n');
  // 先頭5000文字に制限してトークン節約
  return text.length > 5000 ? text.substring(0, 5000) + '\n...(省略)' : text;
}

// ---- Gemini APIで要約 ----
function summarizeWithClaude(sheetData, fileName) {
  var apiKey = PROPS.getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('GEMINI_API_KEY が設定されていません');

  var prompt =
    '以下は「' + fileName + '」という資金繰り表のデータです。\n' +
    '日付・入金・出金・残高などの情報から、経営者が知るべき重要ポイントを日本語で箇条書き（3〜5点）にまとめてください。\n' +
    '特に【資金不足リスク】【大きな入出金】【残高の最低値と日付】を含めてください。\n\n' +
    sheetData;

  var payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { maxOutputTokens: 512 }
  };

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
  var response = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) throw new Error('Gemini API エラー: HTTP ' + code + ' ' + response.getContentText());

  var json = JSON.parse(response.getContentText());
  return json.candidates[0].content.parts[0].text.trim();
}

// ---- スプレッドシートからKPI数値を抽出（簡易版） ----
function extractKpi(sheetData) {
  // 数値行から残高・入金・出金の合計を試みる（フォーマット不定のため簡易）
  var lines = sheetData.split('\n');
  var amounts = [];
  for (var i = 0; i < lines.length; i++) {
    var matches = lines[i].match(/[\d,]+/g);
    if (matches) {
      matches.forEach(function(m) {
        var n = parseInt(m.replace(/,/g, ''), 10);
        if (n > 1000) amounts.push(n); // 1000以上の数値のみ
      });
    }
  }
  if (amounts.length === 0) return null;
  amounts.sort(function(a, b) { return b - a; });
  return {
    max: amounts[0].toLocaleString(),
    min: amounts[amounts.length - 1].toLocaleString()
  };
}

// ---- モダンHTMLメール送信 ----
function sendEmailReport(items) {
  var email = PROPS.getProperty('REPORT_EMAIL');
  if (!email) throw new Error('REPORT_EMAIL が設定されていません');

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy年MM月dd日');
  var subject = '【資金繰り日次レポート】' + today;

  var itemsHtml = items.map(function(item) {
    var kpiHtml = item.kpi ? [
      '<div style="display:flex;gap:12px;margin:12px 0;flex-wrap:wrap;">',
      kpiCard('最大値', item.kpi.max + '円', '#0f4c81'),
      kpiCard('最小値', item.kpi.min + '円', '#c0392b'),
      '</div>'
    ].join('') : '';

    var summaryLines = item.summary.split('\n').map(function(line) {
      line = line.trim();
      if (!line) return '';
      return '<li style="margin:6px 0;line-height:1.6;">' + escapeHtml(line.replace(/^[-・•]\s*/, '')) + '</li>';
    }).join('');

    return '<div style="background:#fff;border-radius:12px;padding:20px 24px;margin-bottom:20px;box-shadow:0 2px 8px rgba(0,0,0,0.08);">' +
      '<table style="width:100%;border-collapse:collapse;">' +
        '<tr>' +
          '<td style="padding:4px 0;font-size:13px;color:#888;width:80px;">件名</td>' +
          '<td style="padding:4px 0;font-size:13px;color:#333;">' + escapeHtml(item.subject) + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td style="padding:4px 0;font-size:13px;color:#888;">送信者</td>' +
          '<td style="padding:4px 0;font-size:13px;color:#333;">' + escapeHtml(item.from) + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td style="padding:4px 0;font-size:13px;color:#888;">日時</td>' +
          '<td style="padding:4px 0;font-size:13px;color:#333;">' + item.date + '</td>' +
        '</tr>' +
        '<tr>' +
          '<td style="padding:4px 0;font-size:13px;color:#888;">ファイル</td>' +
          '<td style="padding:4px 0;font-size:13px;color:#0f4c81;">&#128206; ' + escapeHtml(item.fileName) + '</td>' +
        '</tr>' +
      '</table>' +
      kpiHtml +
      '<div style="margin-top:16px;padding:16px;background:#f0f6ff;border-left:4px solid #0f4c81;border-radius:0 8px 8px 0;">' +
        '<div style="font-size:12px;font-weight:700;color:#0f4c81;margin-bottom:8px;text-transform:uppercase;letter-spacing:0.5px;">AI 要約（Claude）</div>' +
        '<ul style="margin:0;padding-left:18px;color:#333;font-size:14px;">' + summaryLines + '</ul>' +
      '</div>' +
    '</div>';
  }).join('');

  var header =
    '<div style="background:linear-gradient(135deg,#0f4c81 0%,#1a73e8 100%);border-radius:16px;padding:32px 28px;margin-bottom:24px;color:#fff;">' +
      '<div style="font-size:11px;font-weight:600;letter-spacing:2px;opacity:0.7;margin-bottom:8px;">DAILY REPORT</div>' +
      '<div style="font-size:24px;font-weight:700;margin-bottom:4px;">💴 資金繰りレポート</div>' +
      '<div style="font-size:14px;opacity:0.85;">' + today + '&nbsp;&nbsp;|&nbsp;&nbsp;' + items.length + '件のファイルを処理</div>' +
    '</div>';

  var footer =
    '<div style="text-align:center;padding:20px 0;font-size:12px;color:#aaa;">' +
      '自動生成 by Google Apps Script + Claude AI<br>' +
      '&copy; ' + new Date().getFullYear() + ' 経理レポートシステム' +
    '</div>';

  var html =
    '<!DOCTYPE html><html><head><meta charset="UTF-8"></head>' +
    '<body style="margin:0;padding:0;font-family:-apple-system,BlinkMacSystemFont,\'Segoe UI\',sans-serif;background:#f5f7fa;">' +
      '<div style="max-width:680px;margin:0 auto;padding:24px 16px;">' +
        header +
        itemsHtml +
        footer +
      '</div>' +
    '</body></html>';

  GmailApp.sendEmail(email, subject, '※このメールはHTML形式です', { htmlBody: html });
  Logger.log('メール送信完了: ' + email);
}

function kpiCard(label, value, color) {
  return '<div style="flex:1;min-width:140px;background:' + color + ';color:#fff;border-radius:10px;padding:14px 16px;">' +
    '<div style="font-size:11px;opacity:0.8;margin-bottom:4px;">' + label + '</div>' +
    '<div style="font-size:18px;font-weight:700;">' + value + '</div>' +
    '</div>';
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

// ---- LINE Notify通知 ----
function sendLineNotify(items) {
  var token = PROPS.getProperty('LINE_NOTIFY_TOKEN');
  if (!token) {
    Logger.log('LINE_NOTIFY_TOKEN 未設定のためスキップ');
    return;
  }

  var today = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'MM/dd');
  var lines = ['📊 資金繰りレポート ' + today];

  items.forEach(function(item, idx) {
    lines.push('\n【' + (idx + 1) + '】' + item.fileName);
    // Claude要約の最初の2〜3行
    var summaryLines = item.summary.split('\n').filter(function(l) { return l.trim(); }).slice(0, 3);
    summaryLines.forEach(function(l) { lines.push('• ' + l.trim().replace(/^[-・•]\s*/, '')); });
  });

  var message = lines.join('\n');
  if (message.length > 1000) message = message.substring(0, 997) + '…';

  UrlFetchApp.fetch('https://notify-api.line.me/api/notify', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + token },
    payload: { message: message },
    muteHttpExceptions: true
  });

  Logger.log('LINE通知送信完了');
}

// ---- トリガー自動設定（初回1回だけ実行） ----
function setupTrigger() {
  // 既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'runDailyReport') {
      ScriptApp.deleteTrigger(t);
    }
  });
  // 毎朝8:00に設定
  ScriptApp.newTrigger('runDailyReport')
    .timeBased()
    .atHour(8)
    .everyDays(1)
    .inTimezone('Asia/Tokyo')
    .create();
  Logger.log('トリガー設定完了：毎日 08:00 JST');
}
