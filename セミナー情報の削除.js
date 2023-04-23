/////////////////////////////////

//セミナー情報転記シートの一括削除を行う

/////////////////////////////////
function Data_All_Reset() {
  
    //シート名定義
    var SHEET_NAME_GIKEN="技研商事インターナショナル";
    var SHEET_NAME_MACRO="マクロミル";
  
    //行・列番号定義
    var START_ROW=1;
    var START_COL=1;
    var LAST_COL=10;
  
    //削除対象シートの配列作成
    var delete_array=[SHEET_NAME_GIKEN,SHEET_NAME_MACRO];
  
    //スプシの定義
    var ss=SpreadsheetApp.getActiveSpreadsheet();
  
    for(var i=0;i<delete_array.length;i++){
  
      //操作シートの定義
      var sh=ss.getSheetByName(delete_array[i]);
  
      //削除対象範囲の取得
      var LAST_ROW=sh.getLastRow();
  
      if(LAST_ROW==0){
        //削除処理を行わず次のシートへ
        continue;
  
      }else{
        sh.getRange(START_ROW,START_COL,LAST_ROW,LAST_COL).clearContent();
      }
    }
    Browser.msgBox("全データを削除しました")
  
  }
  
  /////////////////////////////////
  
  //マクロミルの削除を行う
  
  /////////////////////////////////
  function Data_Reset_MACRO(){
  
    //シート名定義
    var SHEET_NAME_MACRO="マクロミル";
  
    //行・列番号定義
    var START_ROW=1;
    var START_COL=1;
    var LAST_COL=10;
  
    //シート定義
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh_macro=ss.getSheetByName(SHEET_NAME_MACRO);
  
    //削除範囲の指定
    var LAST_ROW=sh_macro.getLastRow();
  
    if(LAST_ROW!==0){
      sh_macro.getRange(START_ROW,START_COL,LAST_ROW,LAST_COL).clearContent();
    }
    Browser.msgBox("データを削除しました");
  
  }
  
  
  /////////////////////////////////
  
  //マクロミルシートの削除を行う
  
  /////////////////////////////////
  function Data_Reset_GIKEN(){
  
    //シート名定義
    var SHEET_NAME_GIKEN="技研商事インターナショナル";
  
    //行・列番号定義
    var START_ROW=1;
    var START_COL=1;
    var LAST_COL=10;
  
    //シート定義
    var ss=SpreadsheetApp.getActiveSpreadsheet();
    var sh_giken=ss.getSheetByName(SHEET_NAME_GIKEN);
  
    //削除範囲の指定
    var LAST_ROW=sh_giken.getLastRow();
  
    if(LAST_ROW!==0){
      sh_giken.getRange(START_ROW,START_COL,LAST_ROW,LAST_COL).clearContent();
    }
    Browser.msgBox("データを削除しました");
  
  }
  