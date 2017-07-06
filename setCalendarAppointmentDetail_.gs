/**
* アポイント情報をカレンダーに登録する
* 
* @param {String} adminNo 管理番号
* @param {String} station 最寄駅
* @return
*/
function setCalendarAppointmentDetail_( adminNo, station ) {
  
  // 定数宣言
  var SHEET_LEAD = 'リードセールス';   // シート名：リードセールス
  var SHEET_CORE = 'コアセールス';     // シート名：コアセールス
  var WEEKDAY = ['日','月','火','水','木','金','土']; // 曜日
  var UNDECIDED = '未定';
  
  
  // M_CLIENT_LISTから顧客情報を取得する
  var clientInfo = new oldTableApp.GetClientInfo( adminNo, null, null );
  
  var fileCore = SpreadsheetApp.getActiveSpreadsheet(); //SpreadsheetApp.openById('1MjKBDllrHOL4NolSal9sRNJbGsHVhStHz1K0ZGuHG_c')
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
            
            
            
            // 題名
            var subject = coreClientList[i]['company'] + '＠' + station;
            
            // カレンダー用の日付と時刻
            var calendarDate = Utilities.formatDate( coreClientList[i]['visit_date'], 'JST', 'yyyy/MM/dd') + ' ' + Utilities.formatDate( coreClientList[i]['visit_time'], 'JST', 'HH:mm');
            var start = new Date(calendarDate);
            var end = new Date(calendarDate);
            end.setHours(end.getHours() + 1);
            
            // 共有するメールアドレスを設定する
            var calendarId = 'matsumoto-tad@imprexc.com';
            var memberList = new IcTableApp.getMemberList( [['nam_mem', '=', coreClientList[i]['visit_member']],['status', '<', 9]] );
            if ( memberList.length > 0 ) {
              var calendarId = memberList[0]['id_google'];
            }
            
            // カレンダー用備考
            var body = 
                  '■アポ獲得者：' + coreClientList[i]['call_member']                                   + '\r\n'
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
            
            // カレンダーに共有
            try {
              var calendar = CalendarApp.getCalendarById( calendarId );
              calendar.createEvent( subject , start, end, { description: body, location: coreClientList[i]['address'], sendInvites: true });
            } catch(e) {
              
              Browser.msgBox('カレンダーに登録できませんでした。お手数ですが、手動で予定を登録してください。');
              return;
            }
            
            return;
          }
        }
      }
    }
  }
}

function test_setCalendarAppointmentDetail() {
  
  setCalendarAppointmentDetail_( '114466', 'テスト駅' );
}
