var ss = SpreadsheetApp.getActive();

// シートの取得（なければ作成）
var schedule = ss.getSheetByName('schedule');
var resource = ss.getSheetByName('resource');
var EVM = ss.getSheetByName('EVM');
var PDM = ss.getSheetByName('PDM');
if(!schedule){schedule = ss.insertSheet('schedule', 1);}
if(!resource){resource = ss.insertSheet('resource', 2);}
if(!EVM){EVM = ss.insertSheet('EVM', 3);}
if(!PDM){PDM = ss.insertSheet('PDM', 4)}

//スケジュールシート情報取得
var firstRow = schedule.getRange('1:1');
var columnNum = schedule.getMaxColumns();
var rowNum = schedule.getMaxRows();



//フォーマット用アセット
var items = [
['No.', 'タスク名', '予定開始', '予定終了', '予定工数', '実際開始', '実際終了', '実際工数', '担当', '進捗'],
['wbs', 'tasks','planedStart', 'planedFinish', 'planedWorkload', 'actualStart', 'actualFinish', 'actualWorkload', 'responsiblity', 'progress']
];
var wbsColumnNameLength = items[0].length;





function onOpen() {
  //WBSの情報がなければガントチャート追加
  var today = Moment.moment();
  var checkedRange = schedule.getRange(3, 2, rowNum-3, 1);
  if(firstRow.isBlank()) {
    //月曜スタートになるよう調整
    var tmp = 0;
    while (Math.abs(today.day()) + tmp <= 7) {
     tmp++;
   };
   var monday = today.add(tmp, 'days');

   init();
   formatGantchart(7, monday.format('YYYY/MM/DD'));
 } else {
   drawTodayLine(today);
 };

  //カスタムメニューをUIに追加
  SpreadsheetApp.getUi()
  .createMenu('カスタムメニュー')
      .addItem('サイドバーの表示', 'showSidebar') //表示名、関数名
      .addToUi();
      showSidebar();
    }

    function onEdit(e) {
  // スケジュールシートの機能
  if (e.source.getActiveSheet().getName() == 'schedule') {
    var range = e.range;
    var editedRow = parseInt(range.getRow());
    var editedColumn = parseInt(range.getColumn());
    var selectedItem = schedule.getRange(2, editedColumn).getValue();
    //該当範囲のセルが編集されたらガントチャートを色ぬり
    if (items[1].indexOf(selectedItem, 2) > 0) { //行の項目を追加するときは、検索開始位置を注意
      var lastRow = parseInt(range.getLastRow());
      var lastColumn = parseInt(range.getLastColumn());
      var baseLine = parseInt(findStartPoint('progress'))+1;
      var itemKeys = schedule.getRange(2, 1, 1, baseLine-1).getValues();
      var indexOfPlanedStart = itemKeys[0].indexOf('planedStart');
      var indexOfPlanedFinish = itemKeys[0].indexOf('planedFinish');
      var indexOfActualStart = itemKeys[0].indexOf('actualStart');
      var indexOfActualFinish = itemKeys[0].indexOf('actualFinish');
      var indexOfProgress = itemKeys[0].indexOf('progress');
      var baseDate = Moment.moment(schedule.getRange(2, baseLine).getValue());

      //フォーマットをリセット
      schedule.getRange(editedRow, baseLine, lastRow-editedRow+1, columnNum-baseLine).setBackground('').clearContent();
      for (var j = baseLine; columnNum >= j; j++) {
        if ((j - baseLine+1) % 7 === 0) {
          schedule.getRange(editedRow, j-1, lastRow-editedRow+1, 2).setBackground('#fcefe3');
        };
      };

      //複数セルを考慮してfor文
      for (var i = 0; i < lastRow; i++) {
        var planedStart = Moment.moment(schedule.getRange(editedRow, indexOfPlanedStart+1).getValue());
        var planedFinish = Moment.moment(schedule.getRange(editedRow, indexOfPlanedFinish+1).getValue());
        var actualStart = Moment.moment(schedule.getRange(editedRow, indexOfActualStart+1).getValue());
        var actualFinish = Moment.moment(schedule.getRange(editedRow, indexOfActualFinish+1).getValue());
        var progress = schedule.getRange(editedRow, indexOfProgress+1).getValue();

        //予定終了でオレンジ色のマイルストーン('#FFBB00')
        if (planedFinish.format('YYYY') != 'Invalid date' && planedStart.format('YYYY') == 'Invalid date') {
          var chartFinish = baseLine + planedFinish.diff(baseDate, 'days');
          if (chartFinish >= baseLine && chartFinish < columnNum){
            schedule.getRange(editedRow, chartFinish).setBackground('#FFBB00');
          };
        };

        //予定開始と予定終了かつ実際がどちらか空白なら青色('#e3f0f9')
        if (planedFinish.format('YYYY') != 'Invalid date' && planedStart.format('YYYY') != 'Invalid date') {
          var chartStart = baseLine + planedStart.diff(baseDate, 'days');
          var duration = planedFinish.diff(planedStart, 'days')+1;
          if (chartStart >= baseLine){
            if (chartStart + duration > columnNum) {
              schedule.getRange(editedRow, chartStart, 1, columnNum-chartStart+1).setBackground('#e3f0f9');
            } else {
              schedule.getRange(editedRow, chartStart, 1, duration).setBackground('#e3f0f9');
            };
          };
        };

        //実際開始と実際終了で緑色('#aadca8')
        if (actualFinish.format('YYYY') != 'Invalid date' && actualStart.format('YYYY') != 'Invalid date'){
          var chartStart = baseLine + actualStart.diff(baseDate, 'days');
          var duration = actualFinish.diff(actualStart, 'days')+1;
          if (chartStart >= baseLine){
            if (chartStart + duration > columnNum) {
              schedule.getRange(editedRow, chartStart, 1, columnNum-chartStart+1).setBackground('#aadca8');
            } else {
              schedule.getRange(editedRow, chartStart, 1, duration).setBackground('#aadca8');
            };
          };
          //予定との重複分
          if (planedFinish.format('YYYY') != 'Invalid date' && planedStart.format('YYYY') != 'Invalid date'){
           var planChartStart = baseLine + planedStart.diff(baseDate, 'days');
           var planDuration = planedFinish.diff(planedStart, 'days')+1;
           //重複をチェック
           if(planChartStart <= chartStart+duration && planChartStart+planDuration >= chartStart){
             var isBigger = chartStart >= planChartStart ? chartStart : planChartStart;
             var isSmaller = chartStart+duration <= planChartStart+planDuration ? chartStart+duration : planChartStart+planDuration;
             schedule.getRange(editedRow, isBigger, 1, isSmaller-isBigger).setBackground('#99c6ca');
           };
         };
       };

        //進捗率でマークつける
        if (progress != 'NaN') {
          var chartStart = baseLine + actualStart.diff(baseDate, 'days');
          var duration = actualFinish.diff(actualStart, 'days')+1;
          var markLength = Math.round(duration * progress);
          var progressLine = [];
          progressLine.push([]);
          for (var i = 0; i < markLength; i++) {
            progressLine[0].push("'=");
          };
          schedule.getRange(editedRow, chartStart, 1, markLength).setValues(progressLine);
        };
        var editedRow = editedRow + 1;
      };
    };
  };
};








//関数〜
//サイドバーの表示
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
  .setTitle('プロジェクトマネジメント')
  .setWidth(300);
  SpreadsheetApp.getUi()
  .showSidebar(html);
};

//初期フォーマット
function init(){
  schedule.setFrozenRows(1);
  schedule.setFrozenColumns(2);
  schedule.hideRows(2);
  schedule.setRowHeight(1, 31.5);
  firstRow.setBackground('#f3f3f3');
  var wbsColumnRange = schedule.getRange(1, 1, 2, wbsColumnNameLength);
  wbsColumnRange.setValues(items);

  // 列の追加
  if (columnNum < 30) {
    schedule.insertColumns(wbsColumnNameLength+1, 180);
    columnNum = schedule.getMaxColumns();
  }
};

//ガントチャートのフォーマット
//日付を再設定した時の色の塗り直しに対応してない
function formatGantchart(span, date) {
  var line_column = findStartPoint('progress')+1;
  var date = Moment.moment(date);

  //日付を記録
  schedule.getRange(2, line_column).setValue(date.format('YYYY/MM/DD'));

  //列幅と土日の色
  for (var i = line_column; columnNum >= i; i++) { //ここおかしい
    schedule.setColumnWidth(i, 25);
    if ((i - line_column+1) % 7 === 0) {
      schedule.getRange(2, i-1, rowNum-2, 2).setBackground('#fcefe3');
    }
  };

  //枠線と日付
  while (columnNum >= line_column){
   var range = schedule.getRange(1, line_column, rowNum, line_column+span);
   range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
   var range = schedule.getRange(1, line_column, 1, 1);
   range.setValue('(' + date.format('YYYY/MM/DD') + ')');
   line_column += span;
   date = date.add(span, 'days');
 };
};

function findStartPoint(text) {
  var ary = schedule.getRange('2:2').getValues();
  if (ary[0].indexOf(text) < 0) {
    Logger.log('2列目が変更されています。初期化してください');
  } else {
   return ary[0].indexOf(text)+1;
 }
}



function drawTodayLine (today) {
 var baseLine = parseInt(findStartPoint('progress'))+1;
 var baseDate = Moment.moment(schedule.getRange(2, baseLine).getValue());
 var lastRowOfContents = schedule.getLastRow();
 var nextBaseLine = baseLine + 1;
 var todayLine = baseLine + today.diff(baseDate, 'days');
 //古い線を削除
 var markInAry = schedule.getRange(2, nextBaseLine, 1, columnNum-nextBaseLine).getValues();
 var markColumn = markInAry[0].indexOf('|') + nextBaseLine;
 var targetColumn = schedule.getRange(2, markColumn, lastRowOfContents-1, 1);
 var savedValues = schedule.getRange(2, markColumn, lastRowOfContents-1, 1).getValues();
 if (markColumn-nextBaseLine > 0) {
   targetColumn.clearContent();
   for (var i = 1; i < lastRowOfContents-1; i++) {
    if (savedValues[i][0] === '='){
      schedule.getRange(i+2, markColumn).setValue("'=");
    };
  };
};
 //新しい線を引く
 var verticalLine = [];
 var ary = [];
 for (var i = 0; i < lastRowOfContents+1; i++) {
  verticalLine.push(ary);
};
verticalLine[0].push('|');
if (nextBaseLine < todayLine) {
 schedule.getRange(2, todayLine, lastRowOfContents-1, 1).setValue(verticalLine);
};
};


