/**
* アポ獲得したものを自動でアポ一覧に追加
*
* @param {number} adminNo 管理番号
* @param {string} accuracy アポ確度
* @return
*/
function setCoreClientDetail_( adminNo, accuracy ) {
  
  
  // 作業ファイルを取得
  var fileCore = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheetCore = fileCore.getSheetByName( '顧客リスト' );
  
  var lastRow = sheetCore.getLastRow();
  var lastColumn = sheetCore.getLastColumn();
  
  // リードセールス一覧を取得
  var clientInfo = new oldTableApp.GetClientInfo( adminNo, null, null );
  
  if ( clientInfo ) {
    // M_LEAD_FILE_LISTからファイル情報を取得する
    var leadList = IcTableApp.getLeadList( [[ 'region', '=', clientInfo[0]['REGION']]] );
    
    if ( leadList.length > 0 ) {
      
      var leadClientList = IcFileApp.getLeadClientList( leadList[0]['id_file'] );
      
      for ( var i = 0; i < leadClientList.length; i++ ) {
        
        if ( adminNo == leadClientList[i]['no'] ) {
          
          // データの不備を確認し、不備があればエラーを吐き出す。
          if ( leadClientList[i]['current_status'] != 'アポ' ) {
            
            Browser.msgBox('ステータスがアポではありません。ご確認お願いします。');
            Logger.log('ステータスがアポではありません。ご確認お願いします。');
            return;
          }
          
          // 訪問日付書式をチェック
          if (! underscoreGS._isDate( leadClientList[i]['visit_date'] ) ) {
            
            Browser.msgBox('訪問日付が日付ではありません。ご確認お願いします。');
            Logger.log('訪問日付が日付ではありません。ご確認お願いします。');
            return;
          }
          
          
          // 行を追加
          sheetCore.insertRowAfter( lastRow );
          // 行をコピー
          var rangeCopy = sheetCore.getRange( lastRow, 1, 1, lastColumn );
          rangeCopy.copyTo( sheetCore.getRange( lastRow + 1, 1, 1, lastColumn ));
          
          
          // スプレッドシートに吐き出すデータを格納
          var coreClientInfo = [];
          coreClientInfo.push( leadClientList[i]['no'] );
          coreClientInfo.push( leadClientList[i]['company'] );
          coreClientInfo.push( leadClientList[i]['url'] );
          coreClientInfo.push( leadClientList[i]['tel'] );
          coreClientInfo.push( leadClientList[i]['ceo'] );
          coreClientInfo.push( leadClientList[i]['ceo_ruby'] );
          coreClientInfo.push( leadClientList[i]['address'] );
          coreClientInfo.push( leadClientList[i]['licence'] );
          coreClientInfo.push( leadClientList[i]['ad_yahoo'] );
          coreClientInfo.push( leadClientList[i]['ad_athome'] );
          coreClientInfo.push( leadClientList[i]['ad_suumo'] );
          coreClientInfo.push( leadClientList[i]['ad_others'] );
          coreClientInfo.push( accuracy );
          coreClientInfo.push( leadClientList[i]['visit_date'] );
          coreClientInfo.push( leadClientList[i]['visit_time'] );
          coreClientInfo.push( leadClientList[i]['position'] );
          coreClientInfo.push( leadClientList[i]['client_staff'] );
          coreClientInfo.push( leadClientList[i]['visit_member'] );
          coreClientInfo.push( 'コール' );
          coreClientInfo.push( leadClientList[i]['current_date'] );
          coreClientInfo.push( leadClientList[i]['current_member'] );
          
          var coreClientList = [];
          coreClientList.push( coreClientInfo );
          
          sheetCore.getRange( lastRow + 1, 2, 1, coreClientList[0].length ).setValues( coreClientList );
          sheetCore.getRange( lastRow + 1, 23, 1, 1 ).clearContent();
          sheetCore.getRange( lastRow + 1, 30, 1, 38 ).clearContent();
          sheetCore.getRange( lastRow + 1, 69, 1, lastColumn - 68 ).clearContent();
          sheetCore.getRange( lastRow + 1, 1, 1,lastColumn ).setBackground('white');
          
          break;
        }
      }
    } 
  }
  return;
}


function test_setCoreClientDetail() {
  
  setCoreClientDetail_( '16963', 'B' );
}