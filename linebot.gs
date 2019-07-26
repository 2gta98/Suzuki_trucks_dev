function doPost(e) {
  var messageInfo = JSON.parse(e.postData.contents).events[0];
  var replyToken = messageInfo.replyToken;  // WebHookで受信した応答用Token
  var userMessage = messageInfo.message.text;  // ユーザーのメッセージを取得
  var timestamp = messageInfo.timestamp;
  var userId = messageInfo.source.userId;
  var groupId = messageInfo.source.groupId;
  
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh = ss.getSheetByName('LINE botログ');
  var lastRow = sh.getLastRow();
  
  var data = [];
  data.push(timestamp);
  data.push(userId);
  data.push(userMessage);
  data.push(replyToken);
  
  sh.appendRow(data);
}

// プッシュメッセージ
function pushMessage(userId,text) {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU=';
  var postData = {
    "to": userId,
    "messages": [{
      "type": "text",
      "text": text,
    }]
  };

  var url = "https://api.line.me/v2/bot/message/push";
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };

  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  var response = UrlFetchApp.fetch(url, options);
}