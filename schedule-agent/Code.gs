// ============================================================
// スケジュール管理エージェント
// Google Apps Script + Claude API + Google Calendar
// ============================================================

var PROPS = PropertiesService.getScriptProperties();
var CLAUDE_API_KEY = PROPS.getProperty('CLAUDE_API_KEY');
var CALENDAR_ID    = PROPS.getProperty('CALENDAR_ID') || 'primary';

// ---- Web App エントリポイント ----
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('スケジュール管理エージェント')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---- チャットメッセージ処理（クライアントから呼ばれる） ----
function processMessage(userMessage, history) {
  try {
    var calendarContext = buildCalendarContext();
    var response = callClaude(userMessage, history, calendarContext);
    var action = parseAction(response);

    if (action) {
      var result = executeAction(action);
      return { reply: result.message, events: result.events || null };
    }

    return { reply: response, events: null };
  } catch (e) {
    Logger.log('processMessage error: ' + e.message);
    return { reply: 'エラーが発生しました: ' + e.message, events: null };
  }
}

// ---- 直近のカレンダー情報をコンテキストとして取得 ----
function buildCalendarContext() {
  try {
    var now = new Date();
    var past  = new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);  // 1週間前
    var future = new Date(now.getTime() + 30 * 24 * 60 * 60 * 1000); // 30日後

    var cal    = CalendarApp.getCalendarById(CALENDAR_ID);
    var events = cal.getEvents(past, future);
    var lines  = ['【現在の登録予定（直近30日）】'];

    events.forEach(function(ev) {
      var start = Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      var end   = Utilities.formatDate(ev.getEndTime(),   'Asia/Tokyo', 'HH:mm');
      lines.push('ID:' + ev.getId() + ' | ' + start + '〜' + end + ' | ' + ev.getTitle());
    });

    if (events.length === 0) lines.push('（登録された予定はありません）');
    return lines.join('\n');
  } catch (e) {
    return '（カレンダー情報の取得に失敗しました）';
  }
}

// ---- Claude API 呼び出し ----
function callClaude(userMessage, history, calendarContext) {
  if (!CLAUDE_API_KEY) throw new Error('CLAUDE_API_KEY が設定されていません');

  var systemPrompt =
    'あなたはスケジュール管理アシスタントです。ユーザーの自然な日本語の指示をもとに、カレンダー操作を行います。\n\n' +
    '以下のアクションが必要な場合、必ずJSON形式で応答してください（マークダウンのコードブロック不要）:\n\n' +
    '【予定追加】\n' +
    '{"action":"add","title":"タイトル","start":"2024/01/15 14:00","end":"2024/01/15 15:00","description":"メモ（省略可）"}\n\n' +
    '【予定削除】\n' +
    '{"action":"delete","eventId":"イベントID"}\n\n' +
    '【予定閲覧】\n' +
    '{"action":"list","from":"2024/01/15","to":"2024/01/20"}\n\n' +
    '【前後の予定確認】\n' +
    '{"action":"around","date":"2024/01/15","days":3}\n\n' +
    'アクション不要の会話・質問には通常の日本語で回答してください。\n' +
    '今日の日付: ' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd (E)') + '\n\n' +
    calendarContext;

  var messages = [];
  if (history && history.length > 0) {
    history.slice(-10).forEach(function(h) {
      messages.push({ role: h.role, content: h.content });
    });
  }
  messages.push({ role: 'user', content: userMessage });

  var payload = {
    model: 'claude-haiku-4-5-20251001',
    max_tokens: 1024,
    system: systemPrompt,
    messages: messages
  };

  var response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: {
      'x-api-key': CLAUDE_API_KEY,
      'anthropic-version': '2023-06-01',
      'content-type': 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = response.getResponseCode();
  if (code !== 200) throw new Error('Claude API エラー: HTTP ' + code);

  var json = JSON.parse(response.getContentText());
  return json.content[0].text.trim();
}

// ---- Claudeの応答からアクションJSONをパース ----
function parseAction(text) {
  var match = text.match(/\{[\s\S]*?"action"\s*:\s*"(add|delete|list|around)"[\s\S]*?\}/);
  if (!match) return null;
  try {
    return JSON.parse(match[0]);
  } catch (e) {
    return null;
  }
}

// ---- カレンダー操作実行 ----
function executeAction(action) {
  var cal = CalendarApp.getCalendarById(CALENDAR_ID);

  switch (action.action) {

    case 'add': {
      var start = parseDate(action.start);
      var end   = parseDate(action.end || action.start);
      if (!start || !end) return { message: '日時の形式が正しくありません。例: 2024/01/15 14:00' };

      var opts = {};
      if (action.description) opts.description = action.description;

      var ev = cal.createEvent(action.title, start, end, opts);
      var startStr = Utilities.formatDate(start, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm');
      var endStr   = Utilities.formatDate(end,   'Asia/Tokyo', 'HH:mm');
      return {
        message: '✅ 予定を追加しました！\n\n' +
          '📅 **' + action.title + '**\n' +
          '🕐 ' + startStr + '〜' + endStr + '\n' +
          (action.description ? '📝 ' + action.description : ''),
        events: [{ id: ev.getId(), title: action.title, start: startStr, end: endStr }]
      };
    }

    case 'delete': {
      try {
        var evToDelete = cal.getEventById(action.eventId);
        if (!evToDelete) return { message: '指定されたIDの予定が見つかりませんでした。' };
        var title = evToDelete.getTitle();
        evToDelete.deleteEvent();
        return { message: '🗑️ 予定「' + title + '」を削除しました。' };
      } catch (e) {
        return { message: '削除エラー: ' + e.message };
      }
    }

    case 'list': {
      var from = parseDate(action.from + ' 00:00') || new Date();
      var to   = parseDate(action.to   + ' 23:59') || new Date(from.getTime() + 7 * 24 * 60 * 60 * 1000);
      return formatEventList(cal.getEvents(from, to), action.from + '〜' + action.to);
    }

    case 'around': {
      var center = parseDate(action.date + ' 00:00') || new Date();
      var days   = action.days || 3;
      var rangeFrom = new Date(center.getTime() - days * 24 * 60 * 60 * 1000);
      var rangeTo   = new Date(center.getTime() + days * 24 * 60 * 60 * 1000);
      return formatEventList(cal.getEvents(rangeFrom, rangeTo), action.date + ' 前後' + days + '日間');
    }

    default:
      return { message: '未知のアクションです: ' + action.action };
  }
}

// ---- イベント一覧フォーマット ----
function formatEventList(events, label) {
  if (events.length === 0) {
    return { message: '📭 ' + label + ' の予定はありません。' };
  }

  var lines = ['📅 **' + label + ' の予定（' + events.length + '件）**\n'];
  var formatted = events.map(function(ev) {
    var start = Utilities.formatDate(ev.getStartTime(), 'Asia/Tokyo', 'MM/dd(E) HH:mm');
    var end   = Utilities.formatDate(ev.getEndTime(),   'Asia/Tokyo', 'HH:mm');
    return { id: ev.getId(), title: ev.getTitle(), start: start, end: end };
  });

  formatted.forEach(function(e) {
    lines.push('• ' + e.start + '〜' + e.end + '  **' + e.title + '**');
    lines.push('  `ID: ' + e.id + '`');
  });

  return { message: lines.join('\n'), events: formatted };
}

// ---- 日付文字列パース ----
function parseDate(str) {
  if (!str) return null;
  // yyyy/MM/dd HH:mm または yyyy/MM/dd
  var m = str.match(/(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})(?:\s+(\d{1,2}):(\d{2}))?/);
  if (!m) return null;
  var d = new Date(
    parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]),
    m[4] ? parseInt(m[4]) : 0,
    m[5] ? parseInt(m[5]) : 0
  );
  return isNaN(d.getTime()) ? null : d;
}

// ---- セットアップ確認（デバッグ用） ----
function checkSetup() {
  Logger.log('CLAUDE_API_KEY: ' + (CLAUDE_API_KEY ? '設定済み' : '未設定'));
  Logger.log('CALENDAR_ID: ' + CALENDAR_ID);
  Logger.log(buildCalendarContext());
}
