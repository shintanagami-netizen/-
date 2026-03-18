// ─── 設定 ────────────────────────────────────────────────
var SLACK_BOT_TOKEN = "xoxb-xxxxxxxxxxxx"; // ← あなたのBot Tokenに変更
var SLACK_USER_ID   = "U09FKNYN0LD";
var DAYS_AHEAD      = 4; // 当日 + 4日後まで（金曜実行時に月・火もカバー）

var SEARCH_KEYWORDS = [
  "商談", "打ち合わせ", "アポ", "MTG", "ミーティング",
  "面談", "訪問", "meeting", "appointment"
];
// ─────────────────────────────────────────────────────────

function sendAppointmentReminder() {
  var today   = new Date();
  var endDate = new Date(today);
  endDate.setDate(today.getDate() + DAYS_AHEAD);

  // Gmail検索クエリ組み立て
  var keywordPart = SEARCH_KEYWORDS.map(function(kw) {
    return 'subject:"' + kw + '"';
  }).join(" OR ");

  var fmt = function(d) {
    return Utilities.formatDate(d, "Asia/Tokyo", "yyyy/MM/dd");
  };
  var query = "(" + keywordPart + ") after:" + fmt(today) + " before:" + fmt(new Date(endDate.getTime() + 86400000));

  // Gmailを検索
  var threads = GmailApp.search(query, 0, 50);
  var appointments = [];

  threads.forEach(function(thread) {
    var msg = thread.getMessages()[0];
    appointments.push({
      subject: msg.getSubject(),
      from:    msg.getFrom(),
      date:    Utilities.formatDate(msg.getDate(), "Asia/Tokyo", "MM/dd HH:mm")
    });
  });

  // Slackメッセージ組み立て
  var todayStr  = Utilities.formatDate(today,   "Asia/Tokyo", "yyyy年MM月dd日");
  var endStr    = Utilities.formatDate(endDate,  "Asia/Tokyo", "MM月dd日");
  var lines = [
    "<@" + SLACK_USER_ID + "> *:calendar: アポイントメントリマインダー*",
    todayStr + "（本日）〜 " + endStr + " の商談・アポ一覧",
    "─".repeat(25)
  ];

  if (appointments.length === 0) {
    lines.push("該当するアポイントメントが見つかりませんでした。");
  } else {
    appointments.forEach(function(appo, i) {
      lines.push("*" + (i + 1) + ".* " + appo.subject);
      if (appo.from && appo.from.indexOf("shinta.nagami@salescore.jp") === -1) {
        lines.push("　差出人: " + appo.from);
      }
      lines.push("　日時: " + appo.date);
      lines.push("");
    });
  }

  // Slack DM送信
  var payload = JSON.stringify({
    channel: SLACK_USER_ID,
    text:    lines.join("\n")
  });

  UrlFetchApp.fetch("https://slack.com/api/chat.postMessage", {
    method:  "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + SLACK_BOT_TOKEN },
    payload: payload
  });

  Logger.log("送信完了: " + appointments.length + " 件のアポイントメント");
}
