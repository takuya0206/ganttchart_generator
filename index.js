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
['No.', 'タスク名', '予定開始', '予定終了', '実際開始', '実際終了', '工数\n（予｜実）','', '担当', '進捗'],
['wbs', 'tasks','planedStart', 'planedFinish', 'actualStart', 'actualFinish', 'planedWorkload', 'actualWorkload', 'responsiblity', 'progress']
];
var wbsColumnNameLength = items[0].length;





function onOpen() {
  var today = Moment.moment();
  if(firstRow.isBlank()) {
    //開始日が月曜スタートになるよう調整
    var tmp = 0;
    while (Math.abs(today.day()) + tmp <= 7) {
     tmp++;
   };
   var monday = today.add(tmp, 'days');

   init();
   formatGantchart(7, monday.format('YYYY/MM/DD'));
 } else {
   drawTodayLine(today); //線を引くのは24時間のクローンでもいいかも
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
      var baseLine = parseInt(findStartPoint('progress'))+1;
      var baseDate = Moment.moment(schedule.getRange(2, baseLine).getValue());
      var lastColumn = range.getLastColumn();
      var lastRow = range.getLastRow();
      var indexOfPlanedStart = items[1].indexOf('planedStart');
      var indexOfPlanedFinish = items[1].indexOf('planedFinish');
      var indexOfActualStart = items[1].indexOf('actualStart');
      var indexOfActualFinish = items[1].indexOf('actualFinish');
      var indexOfProgress = items[1].indexOf('progress');
      var datas = schedule.getRange(editedRow, 1, lastRow-editedRow+1, baseLine-1).getValues();

      //フォーマットとコンテンツを削除
      clearContents(editedRow, baseLine, lastRow-editedRow+1, columnNum-baseLine);

      //複数セルを考慮してfor文
      for (var i = 0, len = datas.length; i < len; i++){
        var planedStart = Moment.moment(datas[i][indexOfPlanedStart]);
        var planedFinish = Moment.moment(datas[i][indexOfPlanedFinish]);
        var actualStart = Moment.moment(datas[i][indexOfActualStart]);
        var actualFinish = Moment.moment(datas[i][indexOfActualFinish]);
        var progress = datas[i][indexOfProgress];
        Logger.log(planedStart.format('YYYY'));
        //予定終了でオレンジ色のマイルストーン('#FFBB00')
        if (planedFinish !== '' && planedStart.format('YYYY') === 'Invalid date') {
          setMilestone(editedRow, baseLine, baseDate, planedStart, planedFinish, '#FFBB00');
        };
        //予定開始と予定終了で青色('#e3f0f9')
        if (planedFinish !== '' && planedStart !== '') {
          paintChart(editedRow, baseLine, baseDate, planedStart, planedFinish, '#e3f0f9');
        };
        //実際開始と実際終了で緑色('#aadca8')
        if (actualFinish !== '' && actualStart !== ''){
          paintChart(editedRow, baseLine, baseDate, actualStart, actualFinish, '#aadca8');
        };
        //重複分があれば濃い緑（'#99c6ca'）
        if(planedFinish !== '' && planedStart !== '' && actualFinish !== '' && actualStart !== ''){
          var isOverlap = checkOverlap(planedStart, planedFinish, actualStart, actualFinish);
          if (isOverlap !== false) {
            paintChart(editedRow, baseLine, baseDate, isOverlap[0], isOverlap[1], '#99c6ca');
          };
        };
        //進捗率でマークつける
        if (progress > 0 && actualFinish !== '' && actualStart !== '') {
          markProgress(editedRow, baseLine, baseDate, actualStart, actualFinish, progress);
        };
        editedRow = editedRow + 1;
      };
    };
  };
};




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
  var wbsColumnRange = schedule.getRange(1, 1, 2, wbsColumnNameLength);
  wbsColumnRange.setValues(items);
  schedule.setFrozenRows(1);
  schedule.setFrozenColumns(2);
  schedule.hideRows(2);
  schedule.setRowHeight(1, 31.5);
  firstRow.setBackground('#f3f3f3');
  schedule.getRange(1, items[1].indexOf('planedStart')+1, 1, 2).setBackground('#e3f0f9');
  schedule.getRange(1, items[1].indexOf('actualStart')+1, 1, 2).setBackground('#aadca8');
  schedule.setColumnWidth(items[1].indexOf('planedWorkload')+1, 45);
  schedule.setColumnWidth(items[1].indexOf('actualWorkload')+1, 45);
  schedule.getRange(1, items[1].indexOf('planedWorkload')+1, 1, 2).merge();
  schedule.getRange(1, 1,rowNum, columnNum).setHorizontalAlignment('center');
  schedule.getRange(3, items[1].indexOf('progress')+1, rowNum, 1).setNumberFormat('0%');
  schedule.getRange(3, items[1].indexOf('planedStart')+1, rowNum, 2).setNumberFormat('yyyy/mm/dd');
  schedule.getRange(3, items[1].indexOf('actualStart')+1, rowNum, 2).setNumberFormat('yyyy/mm/dd');
  if (columnNum < 30) {
    schedule.insertColumns(wbsColumnNameLength+1, 180);
    columnNum = schedule.getMaxColumns();
  }
};

//ガントチャートのフォーマット
function formatGantchart(span, date) {
  var line_column = findStartPoint('progress')+1;
  var date = Moment.moment(date);
  var baseDateCell = schedule.getRange(2, line_column);
  var baseDate = Moment.moment(baseDateCell.getValue()).format('YYYY');
  baseDateCell.setValue(date.format('YYYY/MM/DD'));
  //初回の場合
  if(baseDate === 'Invalid date') {
    schedule.getRange(1, line_column, 1, columnNum-line_column).setHorizontalAlignment('left');
    //列幅と土日の色
    for (var i = line_column; columnNum >= i; i++) {
      schedule.setColumnWidth(i, 25);
      if ((i - line_column+1) % 7 === 0) {
        schedule.getRange(2, i-1, rowNum-2, 2).setBackground('#fcefe3');
      };
    };
    //枠線
    var j = line_column;
    while (columnNum >= j){
      var range = schedule.getRange(1, j, rowNum, j+span);
      range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      j += span;
    };
  };
  //初回じゃない場合、既存の情報をグラフに反映
  if(baseDate !== 'Invalid date'){
    rowNum = schedule.getMaxRows();
    var range = schedule.getRange(2, 1, rowNum, line_column-1);
    var data = schedule.getRange(2, 1, range.getLastRow(), line_column-1).getValues();
    var indexOfPlanedStart = data[0].indexOf('planedStart');
    var indexOfPlanedFinish = data[0].indexOf('planedFinish');
    var indexOfActualStart = data[0].indexOf('actualStart');
    var indexOfActualFinish = data[0].indexOf('actualFinish');
    var indexOfProgress = data[0].indexOf('progress');
    clearContents(3, line_column, rowNum-2, columnNum-line_column);
    //ガントチャートの色ぬり
    var repeatedNum = data.length;
    for (var i = 1; i < repeatedNum; i++) {
      if(data[i][indexOfPlanedStart] === '' && data[i][indexOfPlanedFinish] !== ''){
        setMilestone(i+2, line_column, date, Moment.moment(data[i][indexOfPlanedStart]), Moment.moment(data[i][indexOfPlanedFinish]), '#FFBB00');
      };
      if(data[i][indexOfPlanedStart] !== '' && data[i][indexOfPlanedFinish] !== ''){
        paintChart(i+2, line_column, date, Moment.moment(data[i][indexOfPlanedStart]), Moment.moment(data[i][indexOfPlanedFinish]), '#e3f0f9');
      }
      if(data[i][indexOfActualStart] !== '' && data[i][indexOfPlanedFinish] !== ''){
        paintChart(i+2, line_column, date, Moment.moment(data[i][indexOfActualStart]), Moment.moment(data[i][indexOfActualFinish]), '#aadca8');
      }
      if(data[i][indexOfPlanedStart] !== '' && data[i][indexOfPlanedFinish] !== '' && data[i][indexOfActualStart] !== '' && data[i][indexOfPlanedFinish] !== ''){
        var isOverlap = checkOverlap(data[i][indexOfPlanedStart], data[i][indexOfPlanedFinish], data[i][indexOfActualStart], data[i][indexOfPlanedFinish]);
        if (isOverlap !== false) {
          paintChart(i+2, line_column, date, Moment.moment(isOverlap[0]), Moment.moment(isOverlap[1]), '#99c6ca');
        };
      };
      if (data[i][indexOfProgress] > 0 && data[i][indexOfActualStart] !== '' && data[i][indexOfPlanedFinish] !== '') {
        markProgress(i+2, line_column, date, Moment.moment(data[i][indexOfActualStart]), Moment.moment(data[i][indexOfPlanedFinish]), data[i][indexOfProgress]);
      };
    };
  };

  //いずれのケースも日付の記入
  var k = line_column;
  while (columnNum >= k){
   var range = schedule.getRange(1, k, 1, 1);
   range.setValue('(' + date.format('YYYY/MM/DD') + ')');
   k += span;
   date = date.add(span, 'days');
 };
};

//開始位置を見つける
function findStartPoint(text) {
  var ary = schedule.getRange('2:2').getValues();
  if (ary[0].indexOf(text) < 0) {
    Logger.log('2列目が変更されています。初期化してください');
  } else {
   return ary[0].indexOf(text)+1;
 };
};

//フォーマットと値をリセット
function clearContents(top, left, height, width) {
  schedule.getRange(top, left, height, width).setBackground('').clearContent();
  for (var j = left; j <= columnNum; j++) {
    if((j-left+1) % 7 === 0) {
      schedule.getRange(top, j-1, height, 2).setBackground('#fcefe3');
    };
  };
};

//予定開始でマイルストーンを置く
function setMilestone(top, left, baseDate, startDate, finishDate, color){
  var chartFinish = left + finishDate.diff(baseDate, 'days');
  if (chartFinish >= left && chartFinish < columnNum){
    schedule.getRange(top, chartFinish).setBackground(color);
  };
};

//色ぬり機能
function paintChart(top, left, baseDate, startDate, finishDate, color){
  var chartStart = left + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;
  if (chartStart >= left){
    if (chartStart+duration > columnNum) {
      schedule.getRange(top, chartStart, 1, columnNum-chartStart+1).setBackground(color);
    } else {
      schedule.getRange(top, chartStart, 1, duration).setBackground(color);
    };
  };
};

//重複をチェック
function checkOverlap(firstStart, firstFinish, secondStart, secondFinish) {
  if(firstStart <= secondFinish && firstFinish >= secondStart) {
   var isBigger = firstStart >= secondStart ? firstStart : secondStart;
   var isSmaller = firstFinish <= secondFinish ? firstFinish : secondFinish;
   return [isBigger, isSmaller];
 } else {
   return false;
 };
};

//進捗率を=でマーク
function markProgress(top, left, baseDate, startDate, finishDate, progress){
  var chartStart = left + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;
  if (chartStart >= left){
    var markLength = Math.round(duration * progress) > duration ? duration : Math.round(duration * progress);
    var progressLine = [];
    progressLine.push([]);
    for (var g = 0; g < markLength; g++) {
      progressLine[0].push("'=");
    };
    schedule.getRange(top, chartStart, 1, markLength).setValues(progressLine);
  };
};



//日付の線を引く
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

