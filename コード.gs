//↓グローバル変数定義領域

//Spreadsheetオブジェクトを取得
var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("bot");

//lastRowを取得
var lastRowToCopy = sheet.getLastRow();

//スクリプトを実行する列を取得
var actColumn = sheet.getRange("B7").getValue();

//埋めたい歯抜けの行数
var lookRows = sheet.getRange("B5").getValue();

//↑グローバル変数定義領域終了

//新型歯抜け埋めBot
function hanukeNew(){
  //開始時間取得
  var startTime = new Date();
  
  //操作対象セルをオブジェクトで取得
  var changedCells = sheet.getRange(4,actColumn,lastRowToCopy);
  
  //現在のデータを配列に格納
  var hanukeArray = changedCells.getValues();
  Logger.log("配列の長さは「%s」",hanukeArray.length);
  
  //上下を比較して歯抜けを埋めていく
  for(var i=1; i<=hanukeArray.length; i++){
    
    Logger.log("現在は%s行目",i+3);
    //1つ前の要素を取得
    var beforeIndex = hanukeArray[i-1];
    Logger.log("%n上のセルの値は「%s」",beforeIndex);
    
    //当該要素をログ出力
    Logger.log("当該要素の値は「%s」",hanukeArray[i]);
    
    //1つ後の要素を取得
    var afterIndex = hanukeArray[i+1];
    Logger.log("%n下のセルの値は「%s」",afterIndex);
    
    //beforeIndexとafterIndexが一緒なら当該要素に書き込み
    if(beforeIndex == afterIndex){
       hanukeArray[i] = beforeIndex;
       Logger.log("上と下が一致しました");
    }
    
  //forループ終了
  }
  
  //配列hanukeArrayを書き出し
  changedCells.setValues(hanukeArray);
  
  
  //終了時間
  var endTime = new Date();
  Logger.log("実行時間は「%s秒」でした",(endTime-startTime)/1000);
}


//古い歯抜け埋めBot
function hanuke() {
  //開始時間取得
  var startTime = new Date();
  
  //上下を比較して歯抜けを埋めていく
  for(var i=6; i<=lastRowToCopy; i++){
    
    Logger.log("現在は%s行目",i+4);
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
      Logger.log("空白セルを発見しました。");
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
  
  //初期化後に行数を入力
  sheet.getRange(3, actColumn).setValue(actColumn);
  
  
  
}