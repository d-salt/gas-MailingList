var debug = true;

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    { name: 'シートの準備', functionName: 'prepareSheet' },
    { name: 'シフト確認メールプレビュー', functionName: 'showPreview' },
    { name: 'シフト確認メールを送信', functionName: 'sendShift' }
  ];
  spreadsheet.addMenu('シフト管理', menuItems);
}

function sendShift () {
  if (!Browser.msgBox('シフト確認メール送信', 'メールを送信します．\\n よろしいですか？', Browser.Buttons.OK_CANCEL)) {
    return false;
  }
  var strDate = Utilities.formatDate(new Date(), "JST", "MM月dd日");
  var shift = getShift(strDate);
  var mail = makeMailObject(shift);

  for (worker in mail) {
    MailApp.sendEmail({
      to: mail[worker].to,
      cc: mail[worker].cc,
      replyTo: mail[worker].replyTo,
      subject: mail[worker].subject,
      body: mail[worker].body
    });
  }
}

function showPreview () {
  var shift = getShift('プレビュー用');
  var mail = makeMailObject(shift);
  
  var prompt = '';
  
  for (worker in mail) {
    prompt += mail[worker].body + "\\n\\n++++++++++++++++\\n\\n";
  }
  
  if (debug) Logger.log(prompt);
  return Browser.msgBox('プレビュー', prompt, Browser.Buttons.OK);
}

function getShift (strDate) {
  var spreadsheet = SpreadsheetApp.getActive();
  var shiftTimeSheet = spreadsheet.getSheetByName('シフト[' + strDate + ']');
  var lastRow = shiftTimeSheet.getLastRow();
  var lastCol = shiftTimeSheet.getLastColumn();
  
  var shiftTimes = shiftTimeSheet.getRange(2, 1, lastRow, lastCol).getValues();
  
  var shift = {}
  shiftTimes.forEach(function (row) {
    var worker = row[3];
    if (!isset(worker)) return;
    if (shift[worker] === undefined) shift[worker] = [];
    shift[worker].push({ user: row[2], begin: datetimeToTime(row[0]), end: datetimeToTime(row[1]), note: row[4]});
  });
  if (debug) Logger.log(shift);
  
  return shift;
}

function makeMailObject (shift) {
  var spreadsheet = SpreadsheetApp.getActive();
  var settingsSheet = spreadsheet.getSheetByName('設定');
  var settings = settingsSheet.getRange(1, 1, 7, 2).getValues();
  
  var emails = getWorkersEmail();
  
  var cc = isset(settings[1][1]) ? settings[1][1] : '';
  var replyTo = isset(settings[0][1]) ? settings[0][1] : '';
  var subject = settings[3][1];
  
  var bodyFMT = settings[6][1];
  var slotFMT = settings[5][1];
  var signeture = settings[4][1];
  
  var mail = {}
  
  for (worker in shift) {
    mail[worker] = { to: emails[worker], cc: cc, replyTo: replyTo, subject: subject, body: '', shiftText: '' };
    shift[worker].forEach(function (slot, j) {
      mail[worker].shiftText += slotFMT.replace('＜訪問先＞', slot.user).replace('＜開始時間＞', slot.begin).replace('＜終了時間＞', slot.end).replace('＜備考＞', slot.note) + "\\n\\n";
    });
    mail[worker].body = bodyFMT.replace('＜担当者＞', worker).replace('＜シフト＞', mail[worker].shiftText);
    mail[worker].body += signeture;
  }
  
  if (debug) Logger.log(mail);
  return mail;
}

function getWorkersEmail () {
  var spreadsheet = SpreadsheetApp.getActive();
  var workersSheet = spreadsheet.getSheetByName('従業員');
  var workers = workersSheet.getRange(2, 1, workersSheet.getLastRow(), workersSheet.getLastColumn()).getValues();
  
  var emails = {};
  workers.forEach(function (row) {
    emails[row[1]] = row[4];
  })
  return emails;
}

function prepareSheet () {
  var spreadsheet = SpreadsheetApp.getActive();
  var today = new Date();
  var strToday = Utilities.formatDate(today, "JST", "MM月dd日");
  var sheetName = 'シフト[' + strToday + ']';
  if (isset(spreadsheet.getSheetByName(sheetName))) {
    Browser.msgBox('本日のシートはすでに作成済みです．', Browser.Buttons.OK);
    return false;  
  }
  var dayOfWeek = ['日', '月', '火', '水', '木', '金', '土'];
  var baseSeet = spreadsheet.getSheetByName('シフト(' + dayOfWeek[today.getDay()] + ')');
  var sheetCopy = baseSeet.copyTo(spreadsheet);
  sheetCopy.setName(sheetName);
  sheetCopy.activate();
  spreadsheet.moveActiveSheet(2);
}

function datetimeToTime (datetime) {
  return Utilities.formatDate(new Date(datetime), "JST", "HH:mm");
}

function isset (val) {
  if (val === null || val === undefined || val === '') {
    return false;
  } else {
    return true;
  }
}