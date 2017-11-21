// Library
// Moment.js  = key : MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48

var ss = SpreadsheetApp.getActive();

// シートの取得（なければ作成）
var schedule = ss.getSheetByName('schedule');
if(!schedule){schedule = ss.insertSheet('schedule', 1);}

//スケジュールシート情報取得
var firstRow = schedule.getRange('1:1');
var columnNum = schedule.getMaxColumns();
var rowNum = schedule.getMaxRows();

//フォーマット用アセット
var items = [
['No.', '階層別 タスク一覧','','','','', '予定開始', '予定終了', '実際開始', '実際終了', '工数\n（予｜実）','', '担当', '進捗'],
['wbs', 'lv1','lv2','lv3','lv4','lv5','plannedStart', 'plannedFinish', 'actualStart', 'actualFinish', 'plannedWorkload', 'actualWorkload', 'responsiblity', 'progress']
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
    var editedRow = parseInt(e.range.getRow());
    var editedColumn = parseInt(e.range.getColumn());
    var lastColumn = e.range.getLastColumn();
    var lastRow = e.range.getLastRow();
    var selectedItem = schedule.getRange(2, editedColumn).getValue();
    var indexOfSelectedItem = items[1].indexOf(selectedItem);
    var baseLine = parseInt(findStartPoint('progress'))+1;
    var baseDate = Moment.moment(schedule.getRange(2, baseLine).getValue());
    var data = schedule.getRange(2, 1, 1, baseLine-1).getValues();
    var indexOfPlannedStart = data[0].indexOf('plannedStart');
    var indexOfPlannedFinish = data[0].indexOf('plannedFinish');
    var indexOfActualStart = data[0].indexOf('actualStart');
    var indexOfActualFinish = data[0].indexOf('actualFinish');
    var indexOfPlannedWorkload = data[0].indexOf('plannedWorkload');
    var indexOfActualWorkload = data[0].indexOf('actualWorkload');
    var indexOfProgress = data[0].indexOf('progress');
    var datas = schedule.getRange(editedRow, 1, lastRow-editedRow+1, baseLine-1).getValues();

    //plannedStart以降のセルが編集されたらガントチャートを色ぬり
    if (selectedItem === 'plannedStart' || selectedItem === 'plannedFinish' || selectedItem === 'actualStart' || selectedItem === 'actualFinish' || selectedItem === 'progress') {
      //コンテンツとフォーマットを削除
      clearFormat(editedRow, baseLine, lastRow-editedRow+1, columnNum-baseLine);
      //複数セルを考慮してfor文
      for (var i = 0, len = datas.length; i < len; i++){
        var plannedStart = Moment.moment(datas[i][indexOfPlannedStart]);
        var plannedFinish = Moment.moment(datas[i][indexOfPlannedFinish]);
        var actualStart = Moment.moment(datas[i][indexOfActualStart]);
        var actualFinish = Moment.moment(datas[i][indexOfActualFinish]);
        var progress = datas[i][indexOfProgress];
        //予定終了でオレンジ色のマイルストーン('#FFBB00')
        if (plannedFinish !== '' && plannedStart.format('YYYY') === 'Invalid date') {
          setMilestone(editedRow, baseLine, baseDate, plannedStart, plannedFinish, '#FFBB00');
        };
        //予定開始と予定終了で青色('#e3f0f9')
        if (plannedFinish !== '' && plannedStart !== '') {
          paintChart(editedRow, baseLine, baseDate, plannedStart, plannedFinish, '#e3f0f9');
        };
        //実際開始と実際終了で緑色('#aadca8')
        if (actualFinish !== '' && actualStart !== ''){
          paintChart(editedRow, baseLine, baseDate, actualStart, actualFinish, '#aadca8');
        };
        //重複分があれば濃い緑（'#99c6ca'）
        if(plannedFinish !== '' && plannedStart !== '' && actualFinish !== '' && actualStart !== ''){
          var isOverlap = checkOverlap(plannedStart, plannedFinish, actualStart, actualFinish);
          if (isOverlap !== false) {
            paintChart(editedRow, baseLine, baseDate, isOverlap[0], isOverlap[1], '#99c6ca');
          };
        };
        //進捗率でマークつける
        if (progress >= 0 && plannedFinish !== '' && plannedStart !== '') {
          markProgress(editedRow, baseLine, baseDate, plannedStart, plannedFinish, progress);
        };
        editedRow = editedRow + 1;
      };
    };

    //工数（予）が編集されたら...
    if(selectedItem === 'plannedWorkload'){
      var val = 0;
      var lastRowOfContents = schedule.getLastRow();
      var workloadRange = schedule.getRange(1, indexOfPlannedWorkload+1, lastRowOfContents+1, 1);
      var workloadData = workloadRange.getValues();
      var indexData = schedule.getRange(1, 1, lastRowOfContents+1, 1).getValues();
      if(editedRow == lastRow){
        sumWorkload(indexData, workloadData, indexOfPlannedWorkload+1,indexData[editedRow-1][0], val);
      } else {
        sumAllWorkload(indexData, workloadData, workloadRange);
      };
/*
      //e.valueの差異を算出する
      if(typeof e.value === 'object'){
        val = e.oldValue * -1;
        //入力が数値でないとき
      } else if (isNaN(parseInt(e.value))) {
        Browser.msgBox('この列には半角数字をご記入ください');
        if (!e.oldValue){
          e.range.setValue(0);
        } else {
          e.range.setValue(e.oldValue);
        };
        return;
      } else {
        if (!e.oldValue){
          val = e.value
        } else {
          val = e.value - e.oldValue
        };
      };
      */

    };
    //タスク部分のセルが編集されたら...
    if (indexOfSelectedItem > 0 && indexOfSelectedItem < 6){
      var taskEndLine = parseInt(findStartPoint('lv5'));
      var lastRowOfContents = schedule.getLastRow();
      //IDをふる
      //タスクを消したときはIDを消す
      if(e.range.isBlank()){
        var taskRange = schedule.getRange(editedRow, 2, lastRow-editedRow+1,taskEndLine-2+1);
        var idRange = schedule.getRange(editedRow, 1, lastRow-editedRow+1, 1);
        if (taskRange.isBlank() === true){
          idRange.clearContent();
        } else {
          var taskData = taskRange.getValues();
          var idData = idRange.getValues();
          for (var i = 0, len = taskData.length; i < len; i++){
            var isBlank = true;
            for(var j = 0, len2 = taskData[0].length; j < len2; j++){
              if(taskData[i][j] !== ''){isBlank = false;};
            };
            if(isBlank){idData[i][0] = '';};
          };
          idRange.setValues(idData);
        };
      };
      //編集したときはIDをふる
      if(!e.range.isBlank() && editedRow === lastRowOfContents) {
        var range = schedule.getRange(3, 1, lastRowOfContents-3+1, taskEndLine);
        var data = range.getValues();
        var value = writeTaskId(editedRow, editedColumn, data);
        schedule.getRange(editedRow, 1).setValue(value);
      };
      //編集列より下に列があれば合わせて再計算
      if(editedRow !== lastRowOfContents){
        var range = schedule.getRange(3, 1, lastRowOfContents-3+1, taskEndLine);
        var data = range.getValues();
        var row = editedRow;
        for(var i = 0, len = lastRowOfContents-editedRow; i <= len; i++){
          var editedData = data.slice(0, row-2);
          var col = 0;
          label_findCol:
          for(var j = 1, len2 = editedData[0].length; j < len2; j++){
            var lastAry = editedData.length-1;
            if(editedData[lastAry][j] !== ''){
              col = j+1;
              var value = writeTaskId(row, col, editedData);
              data[lastAry][0] = value.toString();
              break label_findCol;
            };
          };
          row += 1;
        };
        range.setValues(data);
      };
      //親タスクのbold
      var fontWeightRange = schedule.getRange(1, 1, lastRowOfContents, baseLine-1);
      var fontWeightData = fontWeightRange.getValues();
      makeParentBold(fontWeightData, fontWeightRange);
    };
    //IDのセルが編集されたら修復
    if (indexOfSelectedItem === 0){
      if(e.oldValue){
        e.range.setValue(e.oldValue);
      };
    };
  };
};

//トリガーのセット
function createTimeDrivenTriggers() {
  ScriptApp.newTrigger('drawTodayLine')
  .timeBased()
  .atHour(0)
  .everyDays(1)
  .create();
}
set_drawTodayLine();



/*↓ functions ↓*/
//親タスクに編集された工数を追加
function sumWorkload(indexData, workloadData, targetRow, taskId, val){
  var taskIdAry = taskId.toString().split('_');
  var parentTasks = [];

  //親タスクの抽出してindexを調べる
  for (var i = 0, len = taskIdAry.length-1; i < len; i++){
    taskIdAry.pop();
    var tmp = taskIdAry.join('_');
    for (var j = 0, len2 = indexData.length; j < len2; j++){
      if(tmp == indexData[j][0]){
        parentTasks.push({
          'No.' : tmp,
          'index': j
        });
        workloadData[j][0] = 0
      };
    };
  };
  //親タスクへvalを足していく
  for (var i = 0, len = parentTasks.length; i < len; i++){
/*
   //値のみを親タスクへ追加（再計算なし）
    if(workloadData[parentTasks[i]['index']][0] === ''){
      workloadData[parentTasks[i]['index']][0] = 0;
    };
    var result = [workloadData[parentTasks[i]['index']][0] + parseInt(val)];
    schedule.getRange(parentTasks[i]['index']+1, targetRow).setValue(result);
  };
  */
    //親タスクに紐づく子タスクを全て再計算
    var parAry = parentTasks[i]['No.'].split('_');
    for (var j = parentTasks[i]['index'], len2 = workloadData.length; j < len2; j++){
      var isChild = true;
      var judgedAry = indexData[j][0].toString().split('_');
      //初めの値と長さで一回目の判定
      if(parAry[0] === judgedAry[0] && parAry.length === judgedAry.length-1){
        //全値で判定
        for (var k = 0, len3 = parAry.length; k < len3; k++){
          if(parAry[k] !== judgedAry[k]){
            isChild = false;
            break;
          };
        };
        if(isChild){
          //ターゲットが空白の場合は0を挿入
          if(workloadData[parentTasks[i]['index']][0] === ''){
            workloadData[parentTasks[i]['index']][0] = 0;
          };
          workloadData[parentTasks[i]['index']][0] += workloadData[j][0];
        };
      };
    };
    schedule.getRange(parentTasks[i]['index']+1, targetRow).setValue([workloadData[parentTasks[i]['index']]]);
  };
};


//特定の行の親タスクを全て合計
function sumAllWorkload(indexData, workloadData, targetRange){
  var parentTasks = [];
  //親タスクの抽出
  for (var i = 0, len = indexData.length; i < len; i++){
    var tmp = indexData[i][0].toString() + '_1';
    for (var j = i, len2 = indexData.length; j < len2; j++){
     if(indexData[j][0].toString() === tmp){
       parentTasks.push({
         'No.': indexData[i][0].toString(),
         'index': i
       });
       workloadData[i][0] = 0;
       break;
     };
   };
 };
  //親に紐づく子を判定して合計していく
  for (var i = parentTasks.length-1; 0 <= i; i--){
    var parAry = parentTasks[i]['No.'].split('_');
    for (var j = parentTasks[i]['index'], len = indexData.length; j < len; j++){
      var isChild = true;
      var judgedAry = indexData[j][0].toString().split('_');
        //初めの値と長さで一回目の判定
        if(parAry[0] === judgedAry[0] && parAry.length === judgedAry.length-1){
          //全値で判定
          for (var k = 0, len2 = parAry.length; k < len2; k++){
            if(parAry[k] !== judgedAry[k]){
              isChild = false;
              break;
            };
          };
          if(isChild){
            //ターゲットが空白の場合は0を挿入
            if(workloadData[j][0] === ''){
              workloadData[j][0] = 0;
            };
            workloadData[parentTasks[i]['index']][0] +=  parseInt(workloadData[j][0]);
          };
        };
      };
    };
    targetRange.setValues(workloadData);
  };


  function front_sumAllWorkload(){
    var colOfPlannedWorkload = findStartPoint('plannedWorkload');
    var lastRowOfContents = schedule.getLastRow();
    var workloadRange = schedule.getRange(1, colOfPlannedWorkload, lastRowOfContents+1, 1);
    var workloadData = workloadRange.getValues();
    var indexData = schedule.getRange(1, 1, lastRowOfContents+1, 1).getValues();
    sumAllWorkload(indexData, workloadData, workloadRange);
  }


//親タスクをbold
function makeParentBold(data, range){
  var info = [];
  //fontWeightの二重配列を作る
  for(var i = 0, len = data.length; i < len; i++){
    info.push([]);
    for (j = 0, len2 = data[i].length; j < len2; j++){
      info[i].push('normal');
    };
  };
  //親タスクの抽出。fontWeightへ反映
  for (var i = 0, len = data.length; i < len; i++){
    var tmp = data[i][0].toString() + '_1';
    for (var j = i, len2 = data.length; j < len2; j++){
     if(data[j][0].toString() === tmp){
       for(var k = 0, len3 = info[0].length; k < len3; k++){
         info[i].push('bold');
         info[i].shift();
       };
       break;
     };
   };
 };
 range.setFontWeights(info);
};


//タスクIDを取得する
function writeTaskId(row, col, data){
  var broId = 0;
  var parId = 0;
  var isBro = false;
  var isPar = false;
  var distanceToBro = 0;
  var distanceToPar = 0;
  //兄弟タスクがあるかチェック
  for (var i = data.length-2; i >= 0; i--){
    if (data[i][col-1] != ''){
      broId = data[i][0];
      isBro = true;
      break;
    };
    distanceToBro += 1;
  };
 //親タスクがあるかチェック
 label_innerFor:
 for (var i = data.length-2; i >= 0; i--){
   if(col === 2){
     //左端が編集されたら親タスクへの距離を最大値に
     distanceToPar = Math.pow(2, 53)-1;
     break;
   };
   for(var j = col-2; j > 0; j--){
     if(data[i][j] != ''){
       parId = data[i][0];
       isPar = true;
       break label_innerFor;
     };
   };
   distanceToPar += 1;
 };
 Logger.log('broId:' + broId);
 Logger.log('parId:' + parId);
 Logger.log('distanceToBro:' + distanceToBro);
 Logger.log('distanceToPar:' + distanceToPar);
  //兄弟タスクも親タスクもなければ第一番目のタスク
  if (isBro === false && isPar === false){
    return 1;
  };
  //判定
  if (distanceToBro < distanceToPar){
    if(broId.toString().length === 1){
      return broId + 1;
    } else {
      var ary = broId.toString().split('_');
      ary[ary.length-1] = parseInt(ary[ary.length-1]) + 1;
      var str = ary.join('_');
      return str;
    };
  };
  if (distanceToBro > distanceToPar){
    return parId + '_' + 1;
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
  var indexOfPlannedStart = items[1].indexOf('plannedStart');
  var indexOfPlannedFinish = items[1].indexOf('plannedFinish');
  var indexOfActualStart = items[1].indexOf('actualStart');
  var indexOfActualFinish = items[1].indexOf('actualFinish');
  var indexOfPlannedWorkload = items[1].indexOf('plannedWorkload');
  var indexOfActualWorkload = items[1].indexOf('actualWorkload');
  var indexOfProgress = items[1].indexOf('progress');
  //目次まわり
  schedule.getRange('A:A').setHorizontalAlignment('left').setBackground('#f3f3f3');
  schedule.getRange(1, indexOfPlannedStart+1, rowNum, columnNum-indexOfPlannedStart).setHorizontalAlignment('center');
  wbsColumnRange.setValues(items);
  schedule.setFrozenRows(1);
  schedule.setFrozenColumns(indexOfPlannedStart);
  schedule.hideRows(2);
  schedule.setColumnWidth(1, 70);
  for(var i = 2; i <= 5; i++){
   schedule.setColumnWidth(i, 20);
 }
  //背景色
  firstRow.setBackground('#f3f3f3');
  //schedule.getRange(1, indexOfPlannedStart+1, 1, 2).setBackground('#e3f0f9');
  //schedule.getRange(1, indexOfActualStart+1, 1, 2).setBackground('#aadca8');
  //セルの統合
  schedule.setColumnWidth(indexOfPlannedWorkload+1, 45);
  schedule.setColumnWidth(indexOfActualWorkload+1, 45);
  schedule.getRange(1, indexOfPlannedWorkload+1, 1, 2).merge();
  //文字表示フォーマット
  schedule.getRange(3, indexOfProgress+1, rowNum-3, 1).setNumberFormat('0%');
  schedule.getRange(3, indexOfPlannedStart+1, rowNum-3, 2).setNumberFormat('yyyy/mm/dd');
  schedule.getRange(3, indexOfActualStart+1, rowNum-3, 2).setNumberFormat('yyyy/mm/dd');
  //列の挿入
  if (columnNum < 30) {
    schedule.insertColumns(wbsColumnNameLength+1, 180);
    columnNum = schedule.getMaxColumns();
  }
};

//ガントチャートのフォーマット
function formatGantchart(span, date) {
  var baseLine = findStartPoint('progress')+1;
  var date = Moment.moment(date);
  var baseDateCell = schedule.getRange(2, baseLine);
  var baseDate = Moment.moment(baseDateCell.getValue()).format('YYYY');
  baseDateCell.setValue(date.format('YYYY/MM/DD'));
  //初回の場合
  if(baseDate === 'Invalid date') {
    schedule.getRange(1, baseLine, 1, columnNum-baseLine+1).setHorizontalAlignment('left');
    //列幅と土日の色
    for (var i = baseLine; columnNum >= i; i++) {
      schedule.setColumnWidth(i, 25);
      if ((i - baseLine+1) % 7 === 0) {
        schedule.getRange(2, i-1, rowNum-2+1, 2).setBackground('#fcefe3');
      };
    };
    //枠線
    var j = baseLine;
    while (columnNum >= j){
      var range = schedule.getRange(1, j, rowNum, j+span);
      range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
      j += span;
    };
  };
  //初回じゃない場合、既存の情報をグラフに反映
  if(baseDate !== 'Invalid date'){
    rowNum = schedule.getMaxRows();
    var range = schedule.getRange(2, 1, rowNum-2+1, baseLine-1);
    var data = schedule.getRange(2, 1, range.getLastRow(), baseLine-1).getValues();
    var indexOfPlannedStart = data[0].indexOf('plannedStart');
    var indexOfPlannedFinish = data[0].indexOf('plannedFinish');
    var indexOfActualStart = data[0].indexOf('actualStart');
    var indexOfActualFinish = data[0].indexOf('actualFinish');
    var indexOfProgress = data[0].indexOf('progress');
    clearContents(3, baseLine, rowNum-3+1, columnNum-baseLine+1);
    //ガントチャートの色ぬり
    for (var i = 1, len = data.length; i < len; i++) {
      if(data[i][indexOfPlannedStart] === '' && data[i][indexOfPlannedFinish] !== ''){
        setMilestone(i+2, baseLine, date, Moment.moment(data[i][indexOfPlannedStart]), Moment.moment(data[i][indexOfPlannedFinish]), '#FFBB00');
      };
      if(data[i][indexOfPlannedStart] !== '' && data[i][indexOfPlannedFinish] !== ''){
        paintChart(i+2, baseLine, date, Moment.moment(data[i][indexOfPlannedStart]), Moment.moment(data[i][indexOfPlannedFinish]), '#e3f0f9');
      }
      if(data[i][indexOfActualStart] !== '' && data[i][indexOfPlannedFinish] !== ''){
        paintChart(i+2, baseLine, date, Moment.moment(data[i][indexOfActualStart]), Moment.moment(data[i][indexOfActualFinish]), '#aadca8');
      }
      if(data[i][indexOfPlannedStart] !== '' && data[i][indexOfPlannedFinish] !== '' && data[i][indexOfActualStart] !== '' && data[i][indexOfPlannedFinish] !== ''){
        var isOverlap = checkOverlap(data[i][indexOfPlannedStart], data[i][indexOfPlannedFinish], data[i][indexOfActualStart], data[i][indexOfPlannedFinish]);
        if (isOverlap !== false) {
          paintChart(i+2, baseLine, date, Moment.moment(isOverlap[0]), Moment.moment(isOverlap[1]), '#99c6ca');
        };
      };
      if (data[i][indexOfProgress] > 0 && data[i][indexOfActualStart] !== '' && data[i][indexOfPlannedFinish] !== '') {
        markProgress(i+2, baseLine, date, Moment.moment(data[i][indexOfActualStart]), Moment.moment(data[i][indexOfPlannedFinish]), data[i][indexOfProgress]);
      };
    };
  };
  //いずれのケースも日付の記入
  var k = baseLine;
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

//フォーマットをリセット
//カラーを取得して一週ずつgetRangeを使うのをやめる、|以外の値を削除する
function clearFormat(top, left, height, width) {
  schedule.getRange(top, left, height, width).setBackground('');
  for (var j = left; j < columnNum; j++) {
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
    //clearContent()はclearFomatの調整が終われば削除
    schedule.getRange(top, chartStart, 1, duration).clearContent();
    schedule.getRange(top, chartStart, 1, markLength).setValues(progressLine);
  };
};


//日付の線を引く
function drawTodayLine () {
 var today = Moment.moment();
 var baseLine = parseInt(findStartPoint('progress'))+1;
 var baseDate = Moment.moment(schedule.getRange(2, baseLine).getValue());
 var lastRowOfContents = schedule.getLastRow();
 var nextBaseLine = baseLine + 1;
 var todayLine = baseLine + today.diff(baseDate, 'days');
 //古い線を削除
 var markInAry = schedule.getRange(2, nextBaseLine, 1, columnNum-nextBaseLine+1).getValues();
 var markColumn = markInAry[0].indexOf('|') + nextBaseLine;
 var targetColumn = schedule.getRange(2, markColumn, lastRowOfContents-2+1, 1);
 var savedValues = targetColumn.getValues();
 if (markColumn-nextBaseLine > 0) {
   for (var i = 0, len = savedValues.length; i < len; i++) {
     if (savedValues[i][0] === "="){
       savedValues[i][0] = "'=";
     } else {
       savedValues[i][0] = '';
     };
   };
   targetColumn.setValues(savedValues);
 };
 //新しい線を引く
 if (nextBaseLine < todayLine) {
   var todayColumn = schedule.getRange(2, todayLine, lastRowOfContents+2-1, 1);
   var todayValues = todayColumn.getValues();
   var verticalLine = [];
   var ary = [];
   for (var i = 0, len = todayValues.length; i < len; i++){
     if(todayValues[i][0] === "="){
       todayValues[i][0] = "'=";
     } else {
       todayValues[i][0] = '|';
     };
   };
   todayColumn.setValues(todayValues);
 };
};


//drawTodayLineの重複を許さない形でTriggerをセット
function set_drawTodayLine(){
  var triggers = ScriptApp.getUserTriggers(ss);
  for (var i = 0, len = triggers.length; i < len; i++){
    if(triggers[i].getHandlerFunction() === 'drawTodayLine'){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
  createTimeDrivenTriggers();
};
