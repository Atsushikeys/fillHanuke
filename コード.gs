function hanuke() {
  //Spreadsheetオブジェクトを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //コピペする前にC列を初期化
  var lastRowToCopy = sheet.getLastRow();
  sheet.getRange(3, 3, lastRowToCopy).clear();
  
  
}
