/* Gmail の　inbox を確認して、メンバーのリストを更新する。
 * - 対象となるメールは、 Subject のないもの or 「登録」　と入っているもの
 * - ドキュメントのカラムは連絡先・メールの本文という形式
 */
var debug=false;

function updateList() {
  var sheet=SpreadsheetApp.getActiveSheet();
  var range=sheet.getRange(1, 1, 1, 2);
  range.setValues([['連絡先', '本文']]);
  var r=1;
  GmailApp.getInboxThreads().forEach(function(thread) {
    subject=thread.getFirstMessageSubject();
    if (subject=='' || subject.match('登録')) {
      r+=1;
      var firstMessage=thread.getMessages()[0];
      var body=firstMessage.getPlainBody();
      var email=firstMessage.getFrom();
      Logger.log("%s: %d", email, r);
      var range=sheet.getRange(r, 1, 1, 2);
      range.setValues([[email, body]]);
    }
  })
}

function sendMail() {
  var sheet=SpreadsheetApp.getActiveSheet();

  var maxRow=sheet.getLastRow()  ;
  var range=sheet.getRange(2, 1, maxRow-1, 1);
  var emails=range.getValues().map(function(i) {return i[0]}).join(', ')
  Logger.log(emails);
  var draft=GmailApp.createDraft(GmailApp.getUserLabels(), "PTAからの連絡", "", {
    bcc: emails
  })

}
