function sendEmails() {
  // アクティブシートを取得して変数「sheet」に代入
  var sheet = SpreadsheetApp.getActiveSheet();
  
  // 文字列の「送信済み」を変数「EMAIL_SENT」に代入
  var markedSent = "送信済み";
  
  // sheetに格納されたアクティブシートのデータが入力されている最終行番号を取得してlastrowに代入
  var lastrow = sheet.getLastRow();
  
  // sheetに格納されたアクティブシートをgetLastColumnメソッドでデータが入力されている最終列番号を取得してlastcolumnに代入
  var lastcolumn = sheet.getLastColumn();
  
  // sheetにgetRangeメソッドで1行目の項目行と1列目の「名前」列を除いたデータ範囲を取得してdataRangeに代入
  var dataRange = sheet.getRange(2, 2,lastrow-1,lastcolumn-1)
  
  // データが格納されているdataRangeの、行と列の2次元配列を返すgetValuesを使用し、データ範囲の値を2次元配列に整理してdataに格納
  var data = dataRange.getValues();
  
  // For文で繰り返し処理を開始
  // 初期化でiに0を代入
  // 条件式はiがlengthで取得した2次元配列変数のdataの配列の長さより小さい間を指定
  // 更新式はiに1を加算、dataに格納されている配列の長さ分だけ処理を繰り返し
  // A列は除いているので、B列を0番目の要素として取り込み開始
   for (var i = 0; i < data.length; ++i) {
    var row        = data[i];
     var to        = row[1];                 // C列
     var subject   = row[4];                 // F列
     var body      = row[5];                 // G列
     var option    = {
       name         :row[0],                 // B列
       cc           :row[2],                 // D列
       bcc          :row[3],                 // E列
       htmlBody     :row[6]                  // H列
     }
     var isSent    = row[7];                 // I列
     
     if (isSent    != markedSent) {
       // MailAppクラスのsendEmailメソッドでメールを送信
       // sendEmailメソッドの引数は、（メールアドレス,件名,内容）で各変数で指定
       MailApp.sendEmail(to, subject, body, option);
       // 送信済の行は、5列目のE列にsetValueメソッドを使用し、
       // 変数「EMAIL_SENT」に格納されている文字列「送信済み」を代入
       // getRange(row, column)
      sheet.getRange(2 + i,lastcolumn).setValue(markedSent);
       // スプレッドシートへのアクセスを表すSpreadsheetAppクラスのflushメソッドを使用してスプレッドシートの変更を確定
      SpreadsheetApp.flush();
      }
   }
 }