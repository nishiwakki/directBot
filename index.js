'use strict';

const fs = require('fs');
const readline = require('readline');
const {google} = require('googleapis');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];
const TOKEN_PATH = 'credentials.json';
const content = fs.readFileSync('client_secret.json');

module.exports = (robot) => {

  robot.join((res) => {
    res.send(`重要課題通知BOTです。startと送信してみてください！`);
  });

  robot.respond(/start$/i, (res) => {
    res.send({
      question: "2つの観点から項目を選んでください。",
      options: ["締切間近", "重要度"],
    });
  });

  robot.respond("select", (res) => {
    switch (res.json.options[res.json.response]) {
      case "締切間近":
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
                dates.push([new Date(row[0], (row[1]- 1), row[2], row[4], row[5], 0), row[3], row[6], row[7], row[8]]);
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
              results += `内容：${dates[0][3]}\n`;
              results += `優先度：${dates[0][4]}`;;
            }
            // 課題が2つ以上あるとき
            if (rows.length > 1) {
              var flag = 0;
              for (var i = 0 ; i < rows.length ; i++){
                if((nowdate < dates[i][0]) && (flag == 0)){
                  results += `期限：${formatDate(dates[i][0], dates[i][1])}\n`;
                  results += `講義：${dates[i][2]}\n`;
                  results += `内容：${dates[i][3]}\n`;;
                  results += `優先度：${dates[i][4]}`;;
                  flag = 1;
                }
              }
            }
            res.send(results);
          });
        });
        break;

      case "重要度":
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
            var dates = [];
            if (rows.length > 0) {
              rows.map((row) => {
                dates.push([new Date(row[0], (row[1]- 1), row[2], row[4], row[5], 0), row[3], row[6], row[7], row[8]]);
              });
            } else {
              console.log('No data found.');
            }
            dates.sort(
              function(a, b){
                return (a[4] < b[4] ? -1 : 1);
              }
            );
            var nowdate = new Date();
            if ((rows.length == 1) && (nowdate < dates[0][0])){
              results += `期限：${formatDate(dates[0][0], dates[0][1])}\n`;
              results += `講義：${dates[0][2]}\n`;
              results += `内容：${dates[0][3]}\n`;
              results += `優先度：${dates[0][4]}`;
            }

            if (rows.length > 1) {
              var flag = 0;
              for (var i = 0 ; i < rows.length ; i++){
                if ((nowdate < dates[i][0]) && (flag == 0)){
                  results += `期限：${formatDate(dates[i][0], dates[i][1])}\n`;
                  results += `講義：${dates[i][2]}\n`;
                  results += `内容：${dates[i][3]}\n`;
                  results += `優先度：${dates[i][4]}`;
                  flag = 1;
                  // 同じ優先度が2つ以上ある時の対応
                  if (dates[i][4] == dates[i+1][4]){
                    results += `\n`;
                    results += `-------------------------------------\n`;
                    flag = 0;
                  }
                }
              }
            }
            res.send(results);
          });
        });
        break;
    }
  });
};

// "2018-08-13T04:05:00.000Z"表記を変換する関数（曜日だけ別対応）
function formatDate(date, weekdays) {
  var year = date.getFullYear();
  var month = date.getMonth() + 1;
  var day = date.getDate();;
  var weekday = weekdays;
  var hour = date.getHours();
  // 00表記を可能にする
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
