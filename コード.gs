function hanuke() {
  //Spreadsheetオブジェクトを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //lastRowを取得
  var lastRowToCopy = sheet.getLastRow();
  
  //上下を比較して歯抜けを埋めていく
  for(var i=4; i<=lastRowToCopy; i++){
    
    Logger.log("現在は%s行目",i);
    //上のセルを取得
    var upCell = sheet.getRange(i-1, 3).getValue();
    Logger.log("%n上のセルの値は「%s」",upCell);
    
    //当該セルの値を取得
    var thisCell = sheet.getRange(i, 3).getValue();
    Logger.log("%n当該セルの値は「%s」",thisCell);
    
    //下のセルを取得
    var downCell = sheet.getRange(i+1, 3).getValue();
    Logger.log("%n下のセルの値は「%s」",downCell);
    
    //upCellとdownCellが一緒なら当該セルに書き込み
    if(upCell === downCell){
       sheet.getRange(i-1, 3).copyTo(sheet.getRange(i, 3));
      Logger.log("%n空白セルを発見しました。");
    }
  
  }
  
  
}
