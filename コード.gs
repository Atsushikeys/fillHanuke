//Spreadsheetオブジェクトを取得
var sheet = SpreadsheetApp.getActiveSheet();

//lastRowを取得
var lastRowToCopy = sheet.getLastRow();

//スクリプトを実行する列を取得
var actColumn = sheet.getRange("B7").getValue();

function hanuke() {
  //開始時間取得
  var startTime = new Date();
  

  
  //歯抜けの行数
  var lookRows = sheet.getRange("B5").getValue();
  
  //上下を比較して歯抜けを埋めていく
  for(var i=6; i<=lastRowToCopy; i++){
    
    Logger.log("現在は%s行目",i);
    //上のセルを取得
    var upCell = sheet.getRange(i-lookRows, actColumn).getValue();
    Logger.log("%n上のセルの値は「%s」",upCell);
    
    //当該セルの値を取得
    var thisCell = sheet.getRange(i, actColumn).getValue();
    Logger.log("%n当該セルの値は「%s」",thisCell);
    
    //下のセルを取得
    var downCell = sheet.getRange(i+lookRows, actColumn).getValue();
    Logger.log("%n下のセルの値は「%s」",downCell);
    
    //upCellとdownCellが一緒なら当該セルに書き込み
    if(upCell === downCell){
       sheet.getRange(i-lookRows, actColumn).copyTo(sheet.getRange(i, actColumn));
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
  
  //シートを初期化
  sheet.getRange(3, actColumn, lastRowToCopy).clear();
  
  
  
}