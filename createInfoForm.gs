/* 担当者に送信する留意事項を記入するフォームを作成
* Triggerは配車担当者の送信ボタン
* 
* 
* 
*/


function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet().addMenu('特設メニュー', [
    {name: 'プルダウンを更新する', functionName: 'updateValidationList'},
    {name: '契約情報を取得する', functionName: 'fillPlan'},
    {name: 'カレンダーに登録する', functionName: 'addEvent'},
    {name: 'カレンダーから予定を削除する', functionName: 'delEvent'},
    {name: '営業に配車登録通知を送る', functionName: 'sendForm'},
  ]);
}

// フォームURLを担当者に連絡する
// とりあえずメールで実装

function sendForm() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeLine = sheet.getRange(sheet.getActiveCell().getRowIndex(), 1, 1, sheet.getLastColumn()).getValues();
//  var activeLine = sh_plan.getRange(7, 1, 1, 20).getValues();
  Logger.log(activeLine);
    
  var url = createFormUrl(sheet,activeLine[0]);
  Logger.log(url);

  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh_plan = ss.getSheetByName('配車予定記入表');
  var sh_pic = ss.getSheetByName('営業担当者リスト');
  var lastRow_pic = sh_pic.getLastRow();

  var pic_data = sh_pic.getDataRange().getValues();
  Logger.log(pic_data);
  
  for (var i=1; i<lastRow_pic; i++) {
    if (activeLine[0][6] == pic_data[i][0]) {
      var address = pic_data[i][2];
      var userId = pic_data[i][1];
    }
  }
  Logger.log(address);

  var subject = '【'+ activeLine[0][2] +'向け】 配車登録連絡';

  var body = activeLine[0][6] + '様\n\n以下の通り配車完了致しました。\n以下のURLより内容を確認し、運転手および経理担当者への申し送り事項の記載をお願い致します。';
  body += createBody(activeLine[0],url);

  Logger.log(subject);
  Logger.log(body);

  GmailApp.sendEmail(
    address,
    subject,
    body,
    {
      cc: 'y-yamaki@suzuki-shokai.co.jp',
      bcc: 'tsuji@deal-connect.co.jp',
      name: '配車担当者からの送信',
    }
  );

  if (userId != '') {
    pushMessage(userId,body);
  }
}

// 配車係が手動で実行する際に利用する関数

function addEvent() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeLine = sheet.getRange(sheet.getActiveCell().getRowIndex(), 1, 1, sheet.getLastColumn()).getValues();
  
  createEvent(sheet,activeLine[0]);
}

/* Googleカレンダーに予定を追加・修正する関数
*
* @param 指定するシート
* @param 更新するイベントのデータ
*
*/

function createEvent(sheet, activeLine) {
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh_dri = ss.getSheetByName('車両運転手リスト');
  var lastRow_dri = sh_dri.getLastRow();
  
  // 運転手を特定する
  var dri_data = sh_dri.getDataRange().getValues();
  
  for (var i=1; i<lastRow_dri; i++) {
    if (activeLine[7] == dri_data[i][0]) {
      var calendarId = dri_data[i][3];
    }
  }
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  var time = new Date(activeLine[3]);
  time = Utilities.formatDate(time,'JST','H:mm');
  
  var title = activeLine[2] + '（現着時間： '+ time +'）';
  var startTime = new Date(activeLine[1]);
  startTime.setHours(activeLine[3].getHours());
  startTime.setMinutes(activeLine[3].getMinutes());
  
  var endTime = new Date(activeLine[1]);
  endTime.setHours(activeLine[8].getHours());
  endTime.setMinutes(activeLine[8].getMinutes());
  
  var description = '契約者名： '+activeLine[10]+'（'+activeLine[11]+'）\n';
  description += '排出事業者名： '+activeLine[12]+'（'+activeLine[13]+'）\n';
  description += 'マニフェスト回付先： '+activeLine[15]+'\n';
  description += '起票先： '+activeLine[16]+'\n';
  description += '請求先コード： '+activeLine[17]+'\n';
  description += '\n【運転手向け留意事項】\n'+activeLine[18]+'\n';
  description += '\n【経理担当者向け向け留意事項】\n'+activeLine[19]+'\n';  
  
  var option = {
    description: description,
    location: activeLine[9]
  }
  
  // 未登録の場合は新規イベント作成
  if (activeLine[20] == '') {
    var newEvent = calendar.createEvent(title, startTime, endTime, option);
    
    var eventId = newEvent.getId();
    sheet.getRange(sheet.getActiveCell().getRowIndex(),21).setValue(eventId);
  
  // 登録済みの場合はイベント情報の更新
  } else {
    var eventId = activeLine[20];
    var now = new Date();
    var sixtyDaysFromNow = new Date(now.getTime() + (60 * 24 * 60 * 60 * 1000));
    var events = calendar.getEvents(now, sixtyDaysFromNow);
    
    for each(var evt in events) {
      if (evt.getId() == eventId) {
        evt.setTitle(title);
        evt.setTime(startTime,endTime);
        evt.setDescription(description);
        evt.setLocation(activeLine[9]);
      }
    }
  }
}



function delEvent() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeLine = sheet.getRange(sheet.getActiveCell().getRowIndex(), 1, 1, sheet.getLastColumn()).getValues();
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh_plan = ss.getSheetByName('配車予定記入表');
  var sh_dri = ss.getSheetByName('車両運転手リスト');
  var lastRow_dri = sh_dri.getLastRow();
  
  // 運転手を特定する
  var dri_data = sh_dri.getDataRange().getValues();
  
  for (var i=1; i<lastRow_dri; i++) {
    if (activeLine[0][7] == dri_data[i][0]) {
      var calendarId = dri_data[i][3];
    }
  }
  
  var calendar = CalendarApp.getCalendarById(calendarId);
  
  var eventId = activeLine[0][20];
  var date = new Date(activeLine[0][1]);
  var myEvents = calendar.getEventsForDay(date);
  
  for each(var evt in myEvents) {
      if (evt.getId() == eventId) {
        evt.deleteEvent();
        sheet.getRange(sheet.getActiveCell().getRowIndex(),8).clearContent();
        sheet.getRange(sheet.getActiveCell().getRowIndex(),21).clearContent();        
      }
    }
}
  
// フォーム用URLを作成する
function createFormUrl(sheet, activeLine) {
//  var sheet = SpreadsheetApp.getActiveSheet();
//  var activeLine = sheet.getRange(sheet.getActiveCell().getRowIndex(), 1, 1, sheet.getLastColumn()).getValues();
  
  var index = activeLine[0];
  var genba = activeLine[2];
  var date = new Date(activeLine[1]);
  date = Utilities.formatDate(date,'JST','M/d');
  var time = new Date(activeLine[3]);
  time = Utilities.formatDate(time,'JST','H:mm');
  var commodity = activeLine[4];
  var discharge = activeLine[5];
  var truckNumber = activeLine[7];
  var contractor = activeLine[10];
  var discharger = activeLine[12];
  var collector = activeLine[14];
  var manifesto = activeLine[15];
  var draft = activeLine[16];
  var code = activeLine[17];
  var addInfoDri = activeLine[18];
  var addInfoAcc = activeLine[19];
  
  var googleFormUrl = 'https://docs.google.com/forms/d/e/1FAIpQLSdxnlp_VDqy3oEFmqvHZk_vuwDjvBWyXdsRsQwomA4ankWOng/viewform?usp=pp_url';

  var index_id = '&entry.1993948160=';
  var genba_id = '&entry.1285972937=';
  var date_id = '&entry.732365146=';
  var time_id = '&entry.880017413=';
  var commodity_id = '&entry.2118083035=';
  var discharge_id = '&entry.505498251=';
  var truckNumber_id = '&entry.1001402207=';
  var contractor_id = '&entry.1424716867=';
  var discharger_id = '&entry.173007907=';
  var collector_id = '&entry.587232385=';
  var manifesto_id = '&entry.890536605=';
  var draft_id = '&entry.1806874395=';
  var code_id = '&entry.1917515837=';
  var addInfoDri_id = '&entry.1873014216=';
  var addInfoAcc_id = '&entry.1512727423=';

  var url = googleFormUrl;
  url += index_id + index;
  url += genba_id + genba;
  url += date_id + date;
  url += time_id + time;
  url += commodity_id + commodity;
  url += discharge_id + discharge;
  url += truckNumber_id + truckNumber;
  url += contractor_id + contractor;
  url += discharger_id + discharger;
  url += collector_id + collector;
  url += manifesto_id + manifesto;
  url += draft_id + draft;
  url += code_id + code;
  url += addInfoDri_id + addInfoDri;
  url += addInfoAcc_id + addInfoAcc;
  
var shortUrl = shortenUrl(url);
  
return shortUrl;  
}

// アクティブセルの現場名をキーに「新規契約フォーム回答」シートから現場住所、収集運搬業者名、マニフェスト回付先、起票先、請求先コード、契約者名、排出事業者名を取得し貼り付ける

function fillPlan() {
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh_plan = ss.getSheetByName('配車予定記入表');
  var ss_con = SpreadsheetApp.openById('1ZDSBGJNDqO8yEoh-EgHBrajQ465im4Tjj-Wt1h_ulJ4');
  var sh_con = ss_con.getSheetByName('契約情報リスト');
  
  var sheet = SpreadsheetApp.getActiveSheet();
  var activeLine = sheet.getRange(sheet.getActiveCell().getRowIndex(), 1, 1, sheet.getLastColumn()).getValues();
  
  var lastRow_con = sh_con.getLastRow();
  var data_con = sh_con.getDataRange().getValues();
  var data = [];
  
  for (var i=1; i<lastRow_con; i++) {
    if (activeLine[0][2] == data_con[i][0]) {
      data_con[i].shift();
      data.push(data_con[i]);
      sheet.getRange(sheet.getActiveCell().getRowIndex(), 10, 1, data_con[i].length).setValues(data);
    }
  }
}


// 指示フォームの回答を配車予定記入表に貼り付ける
function updatePlan() {
  var ss = SpreadsheetApp.openById('1LO4sh9eDBk1rdKJLlc-WDPfybo2e3TXcMcfchecVeqg');
  var sh_form = ss.getSheetByName('配車情報Update');
  var sh_plan = ss.getSheetByName('配車予定記入表');
  var lastRow_form = sh_form.getLastRow();
  var lastRow_plan = sh_plan.getLastRow();
  
  var updateData = sh_form.getRange(lastRow_form, 1, 1, 16).getValues();
  var planData = sh_plan.getDataRange().getValues();
  
  // 顧客リストのデータを取得
  var ss_cus = SpreadsheetApp.openById('18GzcJ1gLeA9oefYQ0jUkN3KSRrVGJBHGP4HM5v_6B50');
  var sh_cus = ss_cus.getSheetByName('サマリー');
  var lastRow_cus = sh_cus.getLastRow();
  var data_cus = sh_cus.getDataRange().getValues();
  
  for (var i=6; i<lastRow_plan; i++) { 
    if (updateData[0][15] == planData[i][0]) {
      sh_plan.getRange(i+1,2).setValue(updateData[0][1]);
      sh_plan.getRange(i+1,3).setValue(updateData[0][2]);
      sh_plan.getRange(i+1,4).setValue(updateData[0][3]);
      sh_plan.getRange(i+1,5).setValue(updateData[0][4]);
      sh_plan.getRange(i+1,6).setValue(updateData[0][5]);
      sh_plan.getRange(i+1,8).setValue(updateData[0][6]);
      sh_plan.getRange(i+1,11).setValue(updateData[0][7]);
      sh_plan.getRange(i+1,13).setValue(updateData[0][8]);
      sh_plan.getRange(i+1,15).setValue(updateData[0][9]);
      sh_plan.getRange(i+1,16).setValue(updateData[0][10]);
      sh_plan.getRange(i+1,17).setValue(updateData[0][11]);
      sh_plan.getRange(i+1,18).setValue(updateData[0][12]);
      sh_plan.getRange(i+1,19).setValue(updateData[0][13]);
      sh_plan.getRange(i+1,20).setValue(updateData[0][14]);
      
      for (var j=1; j<lastRow_cus; j++) {
        if (data_cus[j][0] == updateData[0][7]) {
          sh_plan.getRange(i+1,12).setValue(data_cus[j][2]);
        }
      }
      
      for (var j=1; j<lastRow_cus; j++) {
        if (data_cus[j][0] == updateData[0][8]) {
          sh_plan.getRange(i+1,14).setValue(data_cus[j][2]);
        }
      }
      var updatedData = sh_plan.getDataRange().getValues();
      createEvent(sh_plan,updatedData[i]);
    }
  }
}


// URLを短縮する
function shortenUrl(longUrl) {

  const ACCESS_TOKEN = 'e61a6f6f180029dacb75805ba03829adc4cd980c';
  const ACCESS_URL   = 'https://api-ssl.bitly.com/v4/shorten';

  var payload = {
      'long_url': longUrl,
  };

  var headers = {
      'Authorization' : 'Bearer ' + ACCESS_TOKEN,
      'Content-Type': 'application/json',
  }

  var options = {
      "method"      : 'POST',
      'headers'     : headers,
      'payload'     : JSON.stringify(payload),
  }

　 var response = UrlFetchApp.fetch(ACCESS_URL, options);
　 var content = response.getContentText("UTF-8");

　 return JSON.parse(content).link;
}

function createBody(data,url) {
  var date = new Date(data[1]);
  date = Utilities.formatDate(date,'JST','M/d');
  var time = new Date(data[3]);
  time = Utilities.formatDate(time,'JST','H:mm');
  
  var body = '';
  body += '\n\n===================\n'
  body += '日時： '+ date + ' ' + time;
  body += '\n引取現場： '　+ data[2];
  body += '\n品物: '+ data[4];
  body += '\n降ろし場所： '+ data[5];
  body += '\nURL： '+ url + ' 【必ず記入してください】';
  body += '\n=====================\n'
  body += '\n以上\n\n道央配車担当';
  
  return body;
}
