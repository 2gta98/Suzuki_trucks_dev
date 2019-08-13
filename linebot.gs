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

function richMenuPic(){
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU='; 
  var richMenuId = 'richmenu-5f9c6ff72a53f4bfa1ff483dbc23ac0d';
  var url = 'https://api.line.me/v2/bot/richmenu/' + richMenuId + '/content';
  var postData = DriveApp.getFileById('10N7bJHNU7uxoKKHDCsJX3s9Zx5hctCWM');
  
  var blob = postData.getBlob().getAs('image/jpeg');
  Logger.log(blob.getName());
  Logger.log(blob.getContentType());

  var headers = {
    "Content-Type": "image/jpeg",
    "Authorization": "Bearer " + CHANNEL_ACCESS_TOKEN
  };
  
  var options = {
    "method": "post",
    "headers": headers,
    "payload": blob
  };
  
  UrlFetchApp.fetch(url, options);
}

function getRichMenuId() {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU='; 
  var postData = {
    "size": {
      "width": 2500,
      "height": 843
    },
    "selected": false,
    "name": "鈴木商会ver.1",
    "chatBarText": "メニュー",
    "areas": [
      {
        "bounds": {
          "x": 0,
          "y": 0,
          "width": 1250,
          "height": 843
        },
        "action": {
          "type":"uri",
          "label":"契約登録",
          "uri":"line://app/1602629741-RAwkOl2w"
        }
      },
      {
        "bounds": {
          "x": 1251,
          "y": 0,
          "width": 1250,
          "height": 843
        },
        "action": {
          "type":"uri",
          "label":"顧客登録",
          "uri":"line://app/1602629741-RVBEoy4B"
        }
      }
   ]
}
  
  var url = 'https://api.line.me/v2/bot/richmenu';
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };
  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  var richMenuId = UrlFetchApp.fetch(url, options);
  Logger.log(richMenuId);
}


function linkRichMenu() {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU=';
  var url = 'https://api.line.me/v2/bot/richmenu/bulk/link'; 
  var richMenuId = 'richmenu-5f9c6ff72a53f4bfa1ff483dbc23ac0d';
  
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh = ss.getSheetByName('営業担当者リスト');
  var lastRow = sh.getLastRow();
  var picData = sh.getRange(2, 2, lastRow-1, 1).getValues();
  
  var userIds = [];
  
  for (var i=0; i<picData.length; i++) {
    if (picData[i][0] != '') {
      userIds.push(picData[i][0]);
    }
  }
  
  Logger.log(userIds);
  
  var headers = {
    "Content-Type": "application/json",
    'Authorization': 'Bearer ' + CHANNEL_ACCESS_TOKEN,
  };
  
  var postData = {
    "richMenuId": richMenuId,
    "userIds": userIds
  }
  
  var options = {
    "method": "post",
    "headers": headers,
    "payload": JSON.stringify(postData)
  };
  
  UrlFetchApp.fetch(url, options);
  
}