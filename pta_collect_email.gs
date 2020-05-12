/** 名簿が記録されている sheet の名前　*/
var memberSheet='名簿';
var memberSheetNameColumn='名前';
var myEmail=Session.getActiveUser().getEmail();

Logger.log(myEmail);

/**
 * Gmail の　inbox を確認して、sheet のメンバーのリストを更新する
 * - Sheet　について
 *  - *この処理を実行するとsheetの内容はクリアされる*
 *  - 現在 active な sheet を更新
 *  - 「名簿」という sheet がある場合には、その内容をコピーした上で、該当する
 * - Gmailのinbox　について
 *  - 対象となるメールは、 全てのメール
 */
function updateList() {
  var sheet=SpreadsheetApp.getActiveSheet();
  sheet.clearContents();
  sheet.clearNotes();
 
  // タイトルの設定
  var range=sheet.getRange(1, 1, 1, 4)
  /* column について
   * 1=番号 (id)
   * 2=名前 (name)
   * 3=連絡先 (email)
   * 4=メール件名 (subject) <- note にメールの本文
   * 5以降に、過去のメールの内容を記録
   */
  range.setValues([['番号', '名前', '連絡先', 'メール件名']]);
  range.setBackground('yellow');
  // 名前の読み込み
  var members=readMemberSheet();
  Logger.log(members);
  var id=1;
  members.forEach(function(memberName){
    sheet.getRange(id+1, 1, 1, 2).setValues([[id, memberName]]);
    id+=1;
  });
  var nextEmptyRow=id+1;

  // メールの読み込み
  GmailApp.getInboxThreads().forEach(function(thread) {
    var firstMessage=thread.getMessages()[0];
    var subject=firstMessage.getSubject();
    var body=firstMessage.getPlainBody();
    var email=firstMessage.getFrom();
    if (email.search(myEmail) == -1) {
      // メンバーリストの何処にあるのかを探す
      var r=null;
      for (var i=0; i<members.length;i++) {
        memberName=members[i];
        if ((body+subject).replace(/\s/g, '').search(memberName) != -1) {
          Logger.log('name[%s] in [%s]', memberName, body+subject);
          r=i+2; /* i 0 start; row 2 start */
          break;
        }
      }
      // ない場合は次の空いている行に書き出す
      if(r==null){ r=nextEmptyRow++; }
    
      // 記録するカラム
      var c=3;
      while(!sheet.getRange(r, c).isBlank()) {
        c+=2;
      }
      Logger.log("%s: (r=%d, c=%d)", email, r, c);
      sheet.getRange(r, 3, 1, 2).setValues([[email, subject]]);
      sheet.getRange(r, 4, 1, 1).setNote([[body]]);
    }
    else {
      Logger.log('skipping: my email');
    }    
  })
}

/* 名簿の sheet を読み込む。
 */
function readMemberSheet() {
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if (ss==null) {
    Logger.log('No active spreadsheet');
    return [];
  }
  var ms=ss.getSheetByName(memberSheet);
  if (ms==null) {
    Logger.log('Could not find sheet: %s¥nAvailable shets are', memberSheet);
    ss.getSheets().forEach(function(s) {
      Logger.log(' [%s]%s', s.getSheetId(), s.getSheetName());
    });
    return [];
  }
  var maxRow=ms.getLastRow();
  var maxColumn=ms.getLastColumn();
  var headerRange=ms.getRange(1, 1, 1, maxColumn); 
  var nameColumnIndex=headerRange.getValues()[0].indexOf(memberSheetNameColumn)+1; // 0 start and 1 start
  var names=ms.getRange(2, nameColumnIndex, maxRow-1, 1).getValues().map(function(r){
    return r[0].replace(/\s/g, '');
  })
  Logger.log(names);
  return names;
}

/**
 * 現在activeなsheetの、2列目にある目連絡先のリストを取得して Bcc に入れる。
 */
function sendMail() {
  var sheet=SpreadsheetApp.getActiveSheet();

  var maxRow=sheet.getLastRow()  ;
  var range=sheet.getRange(2, 3, maxRow-1, 1);
  var emails=range.getValues().map(function(i) {return i[0]}).filter(function(m) {return m!=''}).join(', ')
  Logger.log(emails);
  var draft=GmailApp.createDraft(Session.getActiveUser().getEmail(), "PTAからの連絡", "", {
    bcc: emails
  })
}
