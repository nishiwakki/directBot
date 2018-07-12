'use strict';

const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
const TOKEN_PATH = 'credentials.json';
const content = fs.readFileSync('client_secret.json');

module.exports = (robot) => {

  robot.respond(/start$/i, (res) => {
    var howtouse = "";
    howtouse += `このBOTは課題を伝えるためのものです。\n`;
    howtouse += `下記のメッセージを送信することで、それに\n`;
    howtouse += `対応したアクションを取ることができます。\n`;
    howtouse += `"danger" -> 最も提出期限が近い課題表示\n`;
    howtouse += `"must" -> 最も優先度の高い課題表示\n`;
    howtouse += `"add" -> 課題追加\n`;
    howtouse += `"clear" -> 達成した課題削除\n`;
    howtouse += `"all" -> 課題一覧`;
    res.send(howtouse);
  });


  robot.respond(/danger$/i, (res) => {
    var results = "";
    authorize(JSON.parse(content), (auth) => {
      const sheets = google.sheets({version: 'v4', auth});
      sheets.spreadsheets.values.get({
        spreadsheetId: '1_t7FLELQ6T08TYRI9F5cA3YFpbA0baVXOO8plMCr1jo',
        range: 'A3:I',
      }, (err, result) => {
        if (err) {
          return console.log('The API returned an error: ' + err);
        }
        const rows = result.data.values;
        // 多次元配列([提出期限, 課題内容])
        var dates = [];
        // 課題が1つ以上あれば
        if (rows.length > 0) {
          rows.map((row) => {
            // [0]年 [1]月 [2]日 [3]曜日 [4]時 [5]分 [6]年 [7]課題内容 [8]優先度
            dates.push([new Date(row[0], row[1], row[2], row[4], row[5], 0), row[3], row[6], row[7], row[8]]);
          });
        } else {
          console.log('No data found.');
        }
        // datesを昔順に並べ替える
        dates.sort(
          function(a, b){
            return (a[0] < b[0] ? -1 : 1);
          }
        );
        // 今現在の日時を取得
        var nowdate = new Date();
        // 課題が1つしかないとき
        if ((rows.length == 1) && (nowdate < dates[0][0])){
          results += `期限：${formatDate(dates[0][0], dates[0][1])}\n`;
          results += `講義：${dates[0][2]}\n`;
          results += `内容：${dates[0][3]}`;
        }
        // 課題が2つ以上あるとき
        if (rows.length > 1) {
          var flag = 0;
          for (var i = 0 ; i < rows.length ; i++){
            if((nowdate < dates[i][0]) && (flag == 0)){
              results += `期限：${formatDate(dates[i][0], dates[i][1])}\n`;
              results += `講義：${dates[i][2]}\n`;
              results += `内容：${dates[i][3]}`;;
              flag = 1;
            }
          }
        }
        res.send(results);
      });
    });
  });

  robot.respond(/danger$/i, (res) => {
    var results = "";
    authorize(JSON.parse(content), (auth) => {
      const sheets = google.sheets({version: 'v4', auth});
      sheets.spreadsheets.values.get({
        spreadsheetId: '1_t7FLELQ6T08TYRI9F5cA3YFpbA0baVXOO8plMCr1jo',
        range: 'A3:I',
      }, (err, result) => {
        if (err) {
          return console.log('The API returned an error: ' + err);
        }
        const rows = result.data.values;
        // 多次元配列([提出期限, 課題内容])
        var dates = [];
        // 課題が1つ以上あれば
        if (rows.length > 0) {
          rows.map((row) => {
            // [0]年 [1]月 [2]日 [3]曜日 [4]時 [5]分 [6]年 [7]課題内容 [8]優先度
            dates.push([new Date(row[0], row[1], row[2], row[4], row[5], 0), row[3], row[6], row[7], row[8]]);
          });
        } else {
          console.log('No data found.');
        }
        // datesを昔順に並べ替える
        dates.sort(
          function(a, b){
            return (a[0] < b[0] ? -1 : 1);
          }
        );
        // 今現在の日時を取得
        var nowdate = new Date();
        // 課題が1つしかないとき
        if ((rows.length == 1) && (nowdate < dates[0][0])){
          results += `期限：${formatDate(dates[0][0], dates[0][1])}\n`;
          results += `講義：${dates[0][2]}\n`;
          results += `内容：${dates[0][3]}`;
        }
        // 課題が2つ以上あるとき
        if (rows.length > 1) {
          var flag = 0;
          for (var i = 0 ; i < rows.length ; i++){
            if((nowdate < dates[i][0]) && (flag == 0)){
              results += `期限：${formatDate(dates[i][0], dates[i][1])}\n`;
              results += `講義：${dates[i][2]}\n`;
              results += `内容：${dates[i][3]}`;;
              flag = 1;
            }
          }
        }
        res.send(results);
      });
    });
  });
};

// "2018-08-13T04:05:00.000Z"表記を変換する関数（曜日だけ別対応）
function formatDate(date, weekdays) {
  var year = date.getFullYear();
  var month = date.getMonth();
  var day = date.getDate();;
  var weekday = weekdays;
  var hour = date.getHours();
  var minute = ("0" + date.getMinutes()).slice(-2);
  return `${year}年${month}月${day}日(${weekday}) ${hour}:${minute}`;
}


function authorize(credentials, callback) {
  const {client_secret, client_id, redirect_uris} = credentials.installed;
  const oAuth2Client = new google.auth.OAuth2(
      client_id, client_secret, redirect_uris[0]);

  fs.readFile(TOKEN_PATH, (err, token) => {
    if (err) return getNewToken(oAuth2Client, callback);
    oAuth2Client.setCredentials(JSON.parse(token));
    callback(oAuth2Client);
  });
}

function getNewToken(oAuth2Client, callback) {
  const authUrl = oAuth2Client.generateAuthUrl({
    access_type: 'offline',
    scope: SCOPES,
  });
  console.log('Authorize this app by visiting this url:', authUrl);
  const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout,
  });
  rl.question('Enter the code from that page here: ', (code) => {
    rl.close();
    oAuth2Client.getToken(code, (err, token) => {
      if (err) return callback(err);
      oAuth2Client.setCredentials(token);
      fs.writeFile(TOKEN_PATH, JSON.stringify(token), (err) => {
        if (err) console.error(err);
        console.log('Token stored to', TOKEN_PATH);
      });
      callback(oAuth2Client);
    });
  });
}
