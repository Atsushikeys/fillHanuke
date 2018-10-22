function hanuke() {
  //開始時間取得
  var startTime = new Date();
  
  //Spreadsheetオブジェクトを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //lastRowを取得
  var lastRowToCopy = sheet.getLastRow();
  
  //何行上まで見るか
  var lookRows = sheet.getRange(3, 1).getValue();
  
  //上下を比較して歯抜けを埋めていく
  for(var i=4; i<=lastRowToCopy; i++){
    
    Logger.log("現在は%s行目",i);
    //上のセルを取得
    var upCell = sheet.getRange(i-lookRows, 3).getValue();
    Logger.log("%n上のセルの値は「%s」",upCell);
    
    //当該セルの値を取得
    var thisCell = sheet.getRange(i, 3).getValue();
    Logger.log("%n当該セルの値は「%s」",thisCell);
    
    //下のセルを取得
    var downCell = sheet.getRange(i+lookRows, 3).getValue();
    Logger.log("%n下のセルの値は「%s」",downCell);
    
    //upCellとdownCellが一緒なら当該セルに書き込み
    if(upCell === downCell){
       sheet.getRange(i-lookRows, 3).copyTo(sheet.getRange(i, 3));
      Logger.log("%n空白セルを発見しました。");
    }
    
  //forループ終了
  }
  
  //終了時間
  var endTime = new Date();
  Logger.log("実行時間は「%s秒」でした",(endTime-startTime)/1000);
  
}

//セル初期化用関数
function initialize(){
    
  //Spreadsheetオブジェクトを取得
  var sheet = SpreadsheetApp.getActiveSheet();
  
  //lastRowを取得
  var lastRow = sheet.getLastRow();
  
  //シートを初期化
  sheet.getRange(3, 3, lastRow).clear();
  
  
  
}