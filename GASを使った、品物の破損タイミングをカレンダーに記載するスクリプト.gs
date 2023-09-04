function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  // カスタムメニューアイテムを追加
  ui.createMenu('カレンダー連携')
   //.addItem('test', 'writeCalendar')
    .addItem('時刻入力用', 'insertCurrentDate')
    .addItem('カレンダーに入力', 'add_calendar_event')
    .addToUi();

}
//　注意。　サービスから、GoogleカレンダーAPIを利用してください。　自分でAPI作らずとも既に登録されています。

// 時刻入力用 この型の時刻ならばカレンダーが受け付ける
function insertCurrentDate() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var range = sheet.getActiveCell(); // 現在選択されているセルを取得
  
  var currentDate = new Date(); // 現在の日付・時刻を取得
  range.setValue(currentDate); // 選択されたセルに日付・時刻を設定
}


// 月の1日に定期的にカレンダーの情報を取得して、その月に破損の問題がある日が来た際に、通知を送る。
// 案１。　破損用カレンダーから取得。　この場合、デフォルトのカレンダーを汚染しないし、カレンダー内容を、月で全取得し、予定がある際に通知でOKになる。　カレンダーID変更でいけたはず。　カレンダー追加コードが必要
// 案２。　識別用のIDをカレンダーに埋め込む案。　カレンダー追加しなくていい。　ただし、デフォルトの予定と被った際、ちょっと見ずらい編集しずらい、ユーザーが見たときにナニコレ？　ってなって消されそう。




// とりあえずカレンダーに記載できるかテスト

// function writeCalendar() {
//   var sheet = SpreadsheetApp.getActiveSheet();
//   var range = sheet.getDataRange();
//   var values = range.getValues();

  

//   var calendar = CalendarApp.getDefaultCalendar();//アカウントのカレンダー引っ張ってくる。

//   //別カレンダーで可能か検証 IDを直接入力する事で飛ぶことは確認。 スプレッドシートへの記載での対応もOK
//   // var getCalendarId_array = sheet.getRange('K1').getValues();
//   // var cal_calendar = CalendarApp.getCalendarById(getCalendarId_array);
//   // プルダウンの数値がどう入るのか確認　連想配列で入ることを確認 [['破損日１年後']] という事は、testPullButton[][]で呼べるってことなので、このままC列を全部配列で突っ込んで、一気に呼び出したほうが早い？
//   var testPullButton = sheet.getRange('C2:C20').getValues();
//   // console.log(testPullButton);
//   // console.log(testPullButton[0][0]);
//   // console.log(testPullButton[5][0]);

//   // javascriptベースのため、連想配列は実際にはオブジェクトとして処理されてるため、.lengthではなくObject.keys(obj).lengthで取得
//   // console.log(Object.keys(testPullButton).length);



//   for (var i = 0; i < values.length; i++) {
//     // Aカラムを日付　Bカラムをタイトルとする
//     var date = values[i][0];
//     var title = values[i][1];

//     // 既存の同じタイトルの予定を検索して削除
//     var events = calendar.getEventsForDay(date);
//     for (var j = 0; j < events.length; j++) {
//       if (events[j].getTitle() === title) {
//           events[j].deleteEvent();
//   }
//     }

//     // 新しい予定を追加
//     calendar.createAllDayEvent(title, date);
//   }
// }





function add_calendar_event(){
//　シートの入力取得
  var sheet = SpreadsheetApp.getActiveSheet();
  // var range = sheet.getDataRange();
  // var values = range.getValues();

  // カレンダー取得
  // getDefaultCalendar()はgmailのメールアドレスと一致したものを呼び出しています。デフォルトを使いたくない場合は別カレンダーを作成し、そちらに記載する事で、他の予定を上書きしてしまう事をさけれます。
  var calendar = CalendarApp.getDefaultCalendar();

  //　破損用カレンダーを別に作ってそこで管理する事も両方できるように変更予定
  // var getCalendarId_array = sheet.getRange('K1').getValues();//K1にカレンダーID入れて動かす予定だった
  // var cal_calendar = CalendarApp.getCalendarById(getCalendarId_array);

  // シート取得座標がバラバラの時ように３つとも分離
  var getStartTime_array = sheet.getRange('A2:A20').getValues();//　購入時として取得　一旦２０行で設定。
  //　プルダウン　[['破損日１年後']] の形で格納。
  var getEndDay_array = sheet.getRange('C2:C20').getValues();// 破損予想日として取得 プルダウンなので連想配列かつ、呼び出しは.lengthではなくObject.keys(obj).lengthで取得 Object.values(obj).lengthでも可
  var getTitele_array = sheet.getRange('D2:D20').getValues();// 商品の名前　車や備品の名前を取得

  var setEndTime_array = new Date();// これが算出した破損予定日で、警告期間のスタートとしてgoogleカレンダーに記載する。
  var setBeCareful_array = new Date();// こっちを注意期間の初日
  var setEndRiskTime_array = new Date();// これを破損予定日越えでいつ壊れてもおかしくない期間の通知初日とする。
  
  
  console.log(getEndDay_array);
  


  for(var i = 0; i < Object.keys(getEndDay_array).length; i++){

    // 
    var setValue_array = getStartTime_array[i][0];//購入日時を入力
    var currentYear = setValue_array.getFullYear();//年数だけ取得
    var currentMonth = setValue_array.getMonth();//月だけ取得

    
    switch (getEndDay_array[i][0]){
      case '破損日１年後':
            // 一旦インスタンス化して入力
            var ec_days = new Date(setValue_array);
            ec_days.setFullYear(currentYear + 1);
            
            
            // 年数を追加 参照渡しを避けるために、ec_days.getTime()を使用。　new Date(ec_days)だと参照渡ししちゃうので注意
            setEndTime_array = new Date(ec_days.getTime());
            
            ec_days.setFullYear(currentYear + 2);//　いつ壊れてもおかしくない期間用
            setEndRiskTime_array = new Date(ec_days.getTime());


            // 確認したところ、１月より前や１２月を超えればば、年を変える処理もしてくれていた。　setMonth()便利
            ec_days.setFullYear(currentYear);//購入日に戻し＋６か月後を注意期間に
            console.log(ec_days + '確認前');
            ec_days.setMonth(currentMonth + 6);
            console.log(ec_days + '確認後');
            setBeCareful_array = new Date(ec_days.getTime());

            break;

      case '破損日２年後':
            // 一旦インスタンス化して入力
            var ec_days = new Date(setValue_array);
            ec_days.setFullYear(currentYear + 2);
            setEndTime_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 1);//　注意期間用
            setBeCareful_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 3);//　いつ壊れてもおかしくない期間用
            setEndRiskTime_array = new Date(ec_days.getTime());
            break;

      case '破損日３年後':
            // 一旦インスタンス化して入力
            var ec_days = new Date(setValue_array);
            ec_days.setFullYear(currentYear + 3);
            setEndTime_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 2);//　注意期間用
            setBeCareful_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 4);//　いつ壊れてもおかしくない期間用
            setEndRiskTime_array = new Date(ec_days.getTime());
            break;

      case '破損日４年後':
            // 一旦インスタンス化して入力
            var ec_days = new Date(setValue_array);
            ec_days.setFullYear(currentYear + 4);
            setEndTime_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 3);//　注意期間用
            setBeCareful_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 5);//　いつ壊れてもおかしくない期間用
            setEndRiskTime_array = new Date(ec_days.getTime());
            break;
      case '破損日５年後':
            // 一旦インスタンス化して入力
            var ec_days = new Date(setValue_array);
            ec_days.setFullYear(currentYear + 5);
            setEndTime_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 4);//　注意期間用
            setBeCareful_array = new Date(ec_days.getTime());
            ec_days.setFullYear(currentYear + 6);//　いつ壊れてもおかしくない期間用
            setEndRiskTime_array = new Date(ec_days.getTime());
            break;


    }    

    console.log(setEndTime_array);
    // getRangeで数値を挿入する場合、シートのA1はgetRange(1,1)なことに注意。0スタートじゃない。
    sheet.getRange(i+2,2).setValue(setEndTime_array);//破損予想日　破損期間
    sheet.getRange(i+2,5).setValue(setBeCareful_array);//注意
    sheet.getRange(i+2,6).setValue(setEndRiskTime_array);//警告
  }
  


  // Logger.log("test_A: " + getStartTime_array[0]);
  // Logger.log("test_A: " + getStartTime_array[1]);
  // Logger.log("test_A: " + getStartTime_array[2]);

  
  
  for(var i = 0; i < getStartTime_array.length; i++){
        // 値が存在する時のみ記載　一列入る想定なので、開始時間のみ比較
        if (getStartTime_array[i] && getStartTime_array[i][0]){// 記載があればTrueで動く
          var title = getTitele_array[i];

          // シートの破損日から日時を作成
          var startTime = new Date(sheet.getRange(i+2,2).getValue());
          startTime.setHours(2);//他予定の邪魔にならないように２時に記載
          startTime.setMinutes(0);
          // シートから破損日取得
          var endTime = new Date(sheet.getRange(i+2,2).getValue());
          endTime.setHours(2);
          endTime.setMinutes(15);
          

          console.log(title);
          console.log(startTime);
          console.log(endTime);

          var setumei = title + 'が平均的な故障する期間に入りました。 買い替えが発生するのでお気をつけください';


          //カレンダーに予定を登録する 破損予定日 setumeiは{description:}で入れないと認識しないため注意です。
          calendar.createEvent((title + 'が破損期間に入りました。'), startTime, endTime,{description:setumei});


          //----------------------------------注意期間---------------------------------------------
          // シートの破損日から日時を作成
          startTime = new Date(sheet.getRange(i+2,5).getValue());
          startTime.setHours(2);//他予定の邪魔にならないように２時に記載
          startTime.setMinutes(0);
          // シートから破損日取得
          endTime = new Date(sheet.getRange(i+2,5).getValue());
          endTime.setHours(2);
          endTime.setMinutes(15);

          setumei = title + 'が低確率ながら故障しかねない期間にはいりました。';
          //カレンダーに登録　注意期間
          calendar.createEvent((title + 'が破損注意期間に入りました。'), startTime, endTime,{description:setumei});


          //----------------------------------警告期間---------------------------------------------
          // シートの破損日から日時を作成
          startTime = new Date(sheet.getRange(i+2,6).getValue());
          startTime.setHours(2);//他予定の邪魔にならないように２時に記載
          startTime.setMinutes(0);
          // シートから破損日取得
          endTime = new Date(sheet.getRange(i+2,6).getValue());
          endTime.setHours(2);
          endTime.setMinutes(15);

          setumei = title + 'がいつ壊れてもおかしくない期間にはいりました。 寿命が近いので、予想外の破損などで、お怪我などしないようお気を付けください。';
          //カレンダーに登録　警告期間
          calendar.createEvent((title + 'が警告期間に入りました。'), startTime, endTime,{description:setumei});

        }


  }

  

}
