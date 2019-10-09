// 受信したメッセージをログとして記録する

function doPost(e) {
  var event = JSON.parse(e.postData.contents).events[0];
  
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh = ss.getSheetByName('LINE botログ');
  var lastRow = sh.getLastRow();
  
  sh.getRange(lastRow+1,1).setValue(event);

  if (event.type === 'message') {
    replyToMessage(event);
  }
}

// プッシュメッセージ
function pushMessage(postData) {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU=';
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
  
  return response.getResponseCode();
}


// リッチメニューの画像の設定
function richMenuPic(){
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU='; 
  var richMenuId = 'richmenu-1c976aa07340c54dae340ea218d51aad';
  var url = 'https://api.line.me/v2/bot/richmenu/' + richMenuId + '/content';
  var postData = DriveApp.getFileById('1DFx-qy-p3XvPOSLGNDuiBIC6kMdGm_bc');
  
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

// リッチメニューIDの取得
function getRichMenuId() {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU='; 
  var postData = {
    "size": {
      "width": 2500,
      "height": 1686
    },
    "selected": false,
    "name": "鈴木商会ver.2",
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
          "label":"配車依頼",
          "uri":"line://app/1602629741-d6lrEGPl"
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
          "label":"契約登録",
          "uri":"line://app/1602629741-RAwkOl2w"
        }
      },
      {
        "bounds": {
          "x": 0,
          "y": 844,
          "width": 1250,
          "height": 843
        },
        "action": {
          "type":"uri",
          "label":"顧客登録",
          "uri":"line://app/1602629741-RVBEoy4B"
        }
      },
      {
        "bounds": {
          "x": 1251,
          "y": 844,
          "width": 1250,
          "height": 843
        },
        "action": {
          "type":"uri",
          "label":"使い方",
          "uri":"line://app/1602629741-8Rr5G0or"
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

// リッチメニューをユーザーに反映、反映の対象は「営業担当者リスト」シート。
function linkRichMenu() {
  var CHANNEL_ACCESS_TOKEN = 'eTEeP70yky7Yq9XebReX6ZlAOnKY2RiaSkM5wc3eFFG46dwOdTZb9E4ySbb4Hh2RHBY4Blfcmhc6Y2Z4wvk3rCxZ4JGvmHZUgrTI1LcrBlrP80xNswUUYSQkgRPiVHBD4Isz/fzsw8hWS43PQeyHpQdB04t89/1O/w1cDnyilFU=';
  var url = 'https://api.line.me/v2/bot/richmenu/bulk/link'; 
  var richMenuId = 'richmenu-1c976aa07340c54dae340ea218d51aad';
  
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