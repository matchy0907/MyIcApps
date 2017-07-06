/**
* アポイント速報を送信する
*
* @param {String} adminNo 管理番号
* @param {String} station 最寄駅
* @return
*/
function sendMailAppointmentDetail_( adminNo, station ) {
  
  // 定数宣言
  var WEEKDAY = ['日','月','火','水','木','金','土']; // 曜日
  
  // M_CLIENT_LISTから顧客情報を取得する
  var clientInfo = new oldTableApp.GetClientInfo( adminNo, null, null );
  
  var fileCore = SpreadsheetApp.getActiveSpreadsheet();
  var fileId = fileCore.getId();
  
  if ( clientInfo ) {
    // M_LEAD_FILE_LISTからファイル情報を取得する
    var leadList = IcTableApp.getLeadList( [[ 'region', '=', clientInfo[0]['REGION']]] );
    var leadClientList = IcFileApp.getLeadClientList( leadList[0]['id_file'] );
    
    var coreClientList = IcFileApp.getCoreClientList( fileId );
    
    // コアセールスリストを不動産マスタIDで検索する
    for( var i=0; i<coreClientList.length; i++ ) {
      
      if ( coreClientList[i]['no'] == adminNo ) {
        
        
        
        // リードセールスリストを不動産マスタIDで検索する
        for( var j=0; j<leadClientList.length; j++ ) {
          if ( leadClientList[j]['no'] == adminNo ) {
            
            
            // アポ日時を文字列に変換
            var appointDate = Utilities.formatDate( coreClientList[i]['visit_date'], 'JST', 'MM/dd') + '（' + WEEKDAY[coreClientList[i]['visit_date'].getDay()] + '）';
            var appointTime = Utilities.formatDate( coreClientList[i]['visit_time'], 'JST', 'HH:mm');
            
            
            // 送付先アドレスを設定
            var sendTo = [];
            sendTo.push( 'nex001-all@imprexc.com' );
            
            // 題名
            var subject = '【NEXTアポ速報】 ' + appointDate + ' ' + appointTime + '～ ＠' + station + ' ' + coreClientList[i]['company'];
            
            // メール本文
            var body =
                  '各位'                                  + '\r\n'
                + 'お疲れ様です。'                        + '\r\n'
                + 'NEXTのアポイントを取得しましたので'    + '\r\n'
                + '取り急ぎ、日時・社名をご報告します。'  + '\r\n'
                + '\r\n'
                + '■アポ獲得者：' + coreClientList[i]['call_member']                                   + '\r\n'
                + '■アポ確度：' + coreClientList[i]['accuracy']                                        + '\r\n'
                + '■訪問日時：' + appointDate + ' ' + appointTime + '～'                               + '\r\n'
                + '■訪問担当：' + coreClientList[i]['visit_member']                                    + '\r\n'
                + '■社名：' + coreClientList[i]['company']                                             + '\r\n'
                + '■最寄駅：' + station                                                                + '\r\n'
                + '■TEL：' + coreClientList[i]['tel']                                                  + '\r\n'
                + '■住所：' + coreClientList[i]['address']                                             + '\r\n'
                + '■URL：' + coreClientList[i]['url']                                                  + '\r\n'
                + '■先方：' + coreClientList[i]['position'] + ' ' + coreClientList[i]['client_staff']  + '\r\n'
                + '■備考：'                                                                            + '\r\n'
                + leadClientList[j]['current_detail']                                                   + '\r\n'
                + '以上、よろしくお願いします。'                                                        + '\r\n';
            
            
            // メールを送信
            try {
              
              MailApp.sendEmail( sendTo, subject, body );
              
              return;
            } catch(e) {
              
              Browser.msgBox('メールが送れませんでした。システム担当者にご確認ください。');
              return;
            }
          }
        }
      }
    }
  }
}

function test_sendMailAppointmentDetail() {
  
  sendMailAppointmentDetail_( '8522', '押上駅徒歩10分（半蔵門線・浅草線など）' );
}
