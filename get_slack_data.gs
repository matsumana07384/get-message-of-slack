var slack_api_channels = "https://slack.com/api/channels";
var slack_api_users= "https://slack.com/api/users";

function myFunction() {
  var token = PropertiesService.getScriptProperties().getProperty('OAuth_token');
  
  var cnanel_list = getSlackAPI(slack_api_channels +".list?token="+token+"");
  
  var channel_name = PropertiesService.getScriptProperties().getProperty('Channel_name');
  
  // 対象となるチャンネル名から対応するチャンネルIDを探し出す
  for (var i=0; i<cnanel_list.channels.length; i++) {
    if (cnanel_list.channels[i].name == channel_name) {
      var channel_id = cnanel_list.channels[i].id;
      break;
    }
  }
  
 // 取得したチャンネルIDをもとにチャンネル内のすべてのメッセージを取得
 var message_list = getSlackAPI(slack_api_channels + ".history?token="+token+"&channel="+channel_id);
  
  //リストに入れるデータを入れる配列を設定
  var release_history = [];

  for (var i=0; i<message_list.messages.length; i++) {
  
      var tmp_line_data = [];
    
      var tmp_message = message_list.messages[i]["text"];
      var tmp_user = message_list.messages[i]["user"];
      var tmp_date = message_list.messages[i]["ts"];
      
      var now_date = new Date();
      var update_date = Utilities.formatDate( now_date, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');
      var yesterday = new Date(now_date.getFullYear(), now_date.getMonth(), now_date.getDate() - 1);
      //UNIX形式になっているので、人間のわかる形式に変換
      var posting_date = new Date(tmp_date*1000);

      //一致する文字列を見つける
      var match_word = PropertiesService.getScriptProperties().getProperty('Match_word');
      if(tmp_message.match(match_word) && posting_date > yesterday){
      
         // ユーザー名をIDから名称に変更
         var user_profile = getSlackAPI(slack_api_users +".info?token="+token+"&user=" + tmp_user+"&pretty=1");
         
         // ユーザー情報を取得
         var usr_realname = user_profile.user["profile"]["real_name"];
         var usr_displayname = user_profile.user["profile"]["display_name"];
         
         tmp_line_data.push(posting_date);         
         tmp_line_data.push(tmp_message);
         tmp_line_data.push(usr_realname);
         tmp_line_data.push(usr_displayname); 
         tmp_line_data.push(update_date);
         
         release_history.push(tmp_line_data);  
      }
  }
  
  //最新から取ってきているので、古い順に並び替え
  var sortedMessages = release_history.reverse(); 
  
  var spread_active = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet(); 
  //最終行を取得
  var last_list = spread_active.getLastRow() +1;
  var item = 5;//列の数
  //スプレッドシートの最終行に追加 - 1行ずつ追加する
  spread_active.getRange(last_list,1,sortedMessages.length,item).setValues(sortedMessages);  
}

//slackAPIを取得するための関数
function getSlackAPI(slack_api_url){
  var url = slack_api_url;
  var response = UrlFetchApp.fetch(url);
  var json = response.getContentText();
  var data = JSON.parse(json);
  return data;
}
