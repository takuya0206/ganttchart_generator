// Library
// Moment.js  = key : MHMchiX6c1bwSqGM1PZiW_PxhMjh3Sh48

var ss = SpreadsheetApp.getActive();
var schedule = ss.getSheetByName('schedule');
var holiday = ss.getSheetByName('holiday');
var columnNum = schedule.getMaxColumns();
var rowNum = schedule.getMaxRows();


function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
  var memo = PropertiesService.getDocumentProperties();
  var lang = Session.getActiveUserLocale();
  memo.setProperty('lang', lang);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  var createChart_text = lang === 'ja' ? 'ガントチャートの作成' : 'Create Gantt Chart';
  var sidebar_text = lang === 'ja' ? 'サイドバーの表示' : 'Show Sidebar';

  if (e && e.authMode === ScriptApp.AuthMode.NONE) {
    menu.addItem(sidebar_text, 'showSidebar');
    showSidebar();
  } else {
    menu.addItem(createChart_text, 'createChart');
  }
  menu.addToUi();
};


function onEdit(e) {
  // スケジュールシートの機能
  if (e.source.getActiveSheet().getName() == 'schedule') {
    var memo = PropertiesService.getDocumentProperties();
    var editedRow = parseInt(e.range.getRow());
    var editedColumn = parseInt(e.range.getColumn());
    var lastColumn = e.range.getLastColumn();
    var lastRow = e.range.getLastRow();
    var selectedItem = schedule.getRange(2, editedColumn).getValue();
    var baseLine = parseInt(findStartPoint('progress'))+1;
    var baseDate = Moment.moment(memo.getProperty('baseDate'));
    var data = schedule.getRange(2, 1, 1, baseLine-1).getValues();
    var indexOfSelectedItem = data[0].indexOf(selectedItem);
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
      var row = editedRow;
      copyDefaultRow(editedRow, baseLine, lastRow-editedRow+1, columnNum-baseLine);
      //複数セルを考慮してfor文
      for (var i = 0, len = datas.length; i < len; i++){
        var plannedStart = Moment.moment(datas[i][indexOfPlannedStart]);
        var plannedFinish = Moment.moment(datas[i][indexOfPlannedFinish]);
        var actualStart = Moment.moment(datas[i][indexOfActualStart]);
        var actualFinish = Moment.moment(datas[i][indexOfActualFinish]);
        var progress = datas[i][indexOfProgress];
        //予定終了でオレンジ色のマイルストーン('#FFBB00')
        if (plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') === 'Invalid date') {
          setMilestone(row, baseLine, baseDate, plannedStart, plannedFinish, '#FFBB00');
        };
        //予定開始と予定終了で青色('#e3f0f9')
        if (plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date') {
          //Error message
          if(plannedStart > plannedFinish){
            var lang = memo.getProperty('lang');
            var text = lang === 'ja' ? '予定開始が予定終了よりも大きな値です。' : 'Planned Start is bigger than Planned Finish.';
            Browser.msgBox('System Error (ID ' + schedule.getRange(row, 1).getValue() + ') : ' + text);
            return;
          }
          //進捗に合わせて色を変える
          var color = judgeColor(plannedStart, plannedFinish, progress);
          if(color === ''){
            color = '#e3f0f9';
            //進捗の行はオンスケなら白色
            schedule.getRange(row, indexOfProgress+1).setBackground('');
          } else {
            schedule.getRange(row, indexOfProgress+1).setBackground(color);
          }
          paintChart(row, baseLine, baseDate, plannedStart, plannedFinish, color);
        };
        //実際開始と実際終了で緑色('#aadca8')
        if (actualFinish.format('YYYY') !== 'Invalid date' && actualStart.format('YYYY') !== 'Invalid date'){
          //Error message
          if(actualStart > actualFinish){
            var lang = memo.getProperty('lang');
            var text = lang === 'ja' ? '実際開始が実際終了よりも大きな値です。' : 'Actual Start is bigger than Actual Finish.';
            Browser.msgBox('System Error (ID ' + schedule.getRange(row, 1).getValue() + ') : ' + text);
            return;
          }
          paintChart(row, baseLine, baseDate, actualStart, actualFinish, '#aadca8');
        };
        //重複分があれば濃い緑（'#99c6ca'）
        if(plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date' && actualFinish.format('YYYY') !== 'Invalid date' && actualStart.format('YYYY') !== 'Invalid date'){
          var isOverlap = checkOverlap(plannedStart, plannedFinish, actualStart, actualFinish);
          if (isOverlap !== false) {
            paintChart(row, baseLine, baseDate, isOverlap[0], isOverlap[1], '#99c6ca');
          };
        };
        //進捗率でマークつける
        if (progress >= 0 && plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date') {
          markProgress(row, baseLine, baseDate, plannedStart, plannedFinish, progress);
        };
        row = row + 1;
      };
    };

    //進捗が編集されたら...
    if(selectedItem === 'progress'){
      var lastRowOfContents = schedule.getLastRow();
      var data = schedule.getRange(1, 1, lastRowOfContents, indexOfPlannedWorkload+1).getValues();
      var progressRange = schedule.getRange(1, indexOfProgress+1, lastRowOfContents, 1);
      var progressData = progressRange.getValues();
      if(datas.length === 1){
        sumProgress(data, progressData, indexOfProgress+1, data[editedRow-1][0]);
      } else {
        var progressFormulas = progressRange.getFormulas();
        var result = sumAllProgress(data, progressData, progressFormulas, indexOfProgress+1);
        progressRange.setValues(result.progressData);
      };
    };

    //工数（予）が編集されたら...
    if(selectedItem === 'plannedWorkload'){
      var lastRowOfContents = schedule.getLastRow();
      var data = schedule.getRange(1, 1, lastRowOfContents, indexOfPlannedWorkload+1).getValues();
      var progressRange = schedule.getRange(1, indexOfProgress+1, lastRowOfContents, 1);
      var progressData = progressRange.getValues();
      if(datas.length === 1){
        sumWnP(data, progressData, indexOfProgress+1, data[editedRow-1][0]);
      } else {
        var PWRange = schedule.getRange(1, data[0].length, lastRowOfContents, 1);
        var PWData = PWRange.getValues();
        var formulas = PWRange.getFormulas();
        var progressFormulas = progressRange.getFormulas();
        var result = sumAllWnP(data, PWData, progressData, formulas, progressFormulas, indexOfProgress+1);
        PWRange.setValues(result.PWData);
        progressRange.setValues(result.progressData);
      };
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
        var indexRange = schedule.getRange(3, 1, lastRowOfContents-3+1, 1);
        var data = range.getValues();
        var indexData = indexRange.getValues();
        var row = editedRow;
        //編集した行から最後の行まで
        for(var i = 0, len = lastRowOfContents-editedRow; i <= len; i++){
          var editedData = data.slice(0, row-2); //一番上からIDを割り振りたい行まで抽出
          var lastIndex = editedData.length-1;
          var col = 0;
          label_findCol:
          //行の中で値が入っている列を探す
          for(var j = 1, len2 = editedData[0].length; j < len2; j++){//1からなのはIDをスキップする為
            if(editedData[lastIndex][j] !== ''){
              col = j+1;
              var value = writeTaskId(row, col, editedData);
              data[lastIndex][0] = value.toString(); //次の計算用
              indexData[lastIndex][0] = value.toString();
              break label_findCol;
            };
          };
          row += 1;
        };
        indexRange.setValues(indexData);
      };
      //親タスクのbold
      var fontWeightRange = schedule.getRange(1, 1, lastRowOfContents, baseLine-1);
      var fontWeightData = fontWeightRange.getValues();
      makeParentBold(fontWeightData, fontWeightRange);
    };

    //IDのセルが編集されたら修復
    if (indexOfSelectedItem === 0){
      if(e.oldValue){
        var lang = memo.getProperty('lang');
        var text = lang === 'ja' ? '本当に編集しますか？タスクIDを編集すると「工数（予）」と「進捗」の計算に不具合が発生します。' : 'Will you edit the task ID? If you do so, system error will happen when calculating workload and progress.';
        var isDelete = Browser.msgBox(text, Browser.Buttons.YES_NO);
        if(isDelete === 'no'){
          e.range.setValue(e.oldValue);
        };
      };
    };
  };
  //holidayの機能
  if (e.source.getActiveSheet().getName() === 'holiday'){
    var memo = PropertiesService.getDocumentProperties();
    if(e.range.getColumn() === 1){
      formatGantchart(7, memo.getProperty('baseDate'));
    };
  };
};


/*↓ functions ↓*/
//ガントチャートの作成
function createChart(){
  var memo = PropertiesService.getDocumentProperties();
  var schedule = ss.getSheetByName('schedule');
  var holiday = ss.getSheetByName('holiday');
  if(!schedule){schedule = ss.insertSheet('schedule', 1);}
  if(!holiday){holiday = ss.insertSheet('holiday', 2);}
  var lang = memo.getProperty('lang');
  var text = lang === 'ja' ? 'ガントチャートの作成を行いますか？' : 'Will you create a gantt chart?';
  var data = getHolidays();
  holiday.clear();
  setHolidays(data);
  resetAll(text);
};

//初期フォーマット
function init(){
  //フォーマット用アセット
  var memo = PropertiesService.getDocumentProperties();
  var lang = memo.getProperty('lang');
  var scheduleItems = lang === 'ja' ? [['タスクID', '階層別 タスク一覧','','','','', '予定開始', '予定終了', '実際開始', '実際終了', '工数\n（予｜実）','', '担当', '進捗'],  
  ['id', 'lv1','lv2','lv3','lv4','lv5','plannedStart', 'plannedFinish', 'actualStart', 'actualFinish', 'plannedWorkload', 'actualWorkload', 'responsiblity', 'progress']
  ] :
  [['Task ID', 'Work Breakdown Structure','','','','', 'Planned Start', 'Planned Finish', 'Actual Start', 'Actual Finish', 'Workload\n（Plan｜Actual）','', 'In Charge', 'Progress'],
  ['id', 'lv1','lv2','lv3','lv4','lv5','plannedStart', 'plannedFinish', 'actualStart', 'actualFinish', 'plannedWorkload', 'actualWorkload', 'responsiblity', 'progress']
  ];
  var note = lang === 'ja' ? '手動で祝日を編集するときは、必ず日付をA列に入力してください。': 'The default holidays are based on Japanese calendar. When editing holidays, please set date in the A column.';
  var format = lang === 'ja' ? 'YYYY/MM/DD' : 'MMM Do YY';
  var cellformat = lang === 'ja' ? 'yyyy/mm/dd':  'MMM d yyyy';
  var scheduleItemsLength = scheduleItems[0].length;
  var firstRow = schedule.getRange('1:1');
  var idColumnRange = schedule.getRange(1, 1, 2, scheduleItemsLength);
  var indexOfPlannedStart = scheduleItems[1].indexOf('plannedStart');
  var indexOfPlannedFinish = scheduleItems[1].indexOf('plannedFinish');
  var indexOfActualStart = scheduleItems[1].indexOf('actualStart');
  var indexOfActualFinish = scheduleItems[1].indexOf('actualFinish');
  var indexOfPlannedWorkload = scheduleItems[1].indexOf('plannedWorkload');
  var indexOfActualWorkload = scheduleItems[1].indexOf('actualWorkload');
  var indexOfProgress = scheduleItems[1].indexOf('progress');

  memo.setProperties({
    'format': format,
    'cellformat': cellformat
  })
  //目次まわり
  schedule.getRange('A:A').setHorizontalAlignment('left').setBackground('#f3f3f3');
  schedule.getRange(1, indexOfPlannedStart+1, rowNum, columnNum-indexOfPlannedStart).setHorizontalAlignment('center');
  idColumnRange.setValues(scheduleItems);
  schedule.setFrozenRows(1);
  schedule.setFrozenColumns(indexOfPlannedStart);
  schedule.hideRows(2);
  schedule.hideColumns(indexOfActualStart+1, 2);
  schedule.setColumnWidth(1, 70);
  holiday.getRange(1, 1).setNote(note);
  for(var i = 2; i <= 5; i++){
    schedule.setColumnWidth(i, 20);
  }
  //背景色
  firstRow.setBackground('#f3f3f3');
  //セルの統合
  schedule.setColumnWidth(indexOfPlannedWorkload+1, 45);
  schedule.setColumnWidth(indexOfActualWorkload+1, 45);
  schedule.getRange(1, indexOfPlannedWorkload+1, 1, 2).merge();
  //文字表示フォーマット
  schedule.getRange(3, indexOfProgress+1, rowNum-3, 1).setNumberFormat('0.0%');
  schedule.getRange(3, indexOfPlannedStart+1, rowNum-3, 4).setNumberFormat(cellformat);
  //行の数
  if(rowNum > 200){
    schedule.deleteRows(200, rowNum-200);
  }

  //トリガーのセット
  set_front_repaintChart();
};


//チャート部分の列数と列幅を調整する
function adjustColums(baseLine, num, width){
  var deleteNum = columnNum - baseLine;
  schedule.setColumnWidth(baseLine, width);
  schedule.getRange(1, baseLine, rowNum, 1).clearContent();
  schedule.getRange(2, baseLine, rowNum-2+1, 1).setBackground('');
  schedule.deleteColumns(baseLine+1,deleteNum);
  schedule.insertColumnsAfter(baseLine, num-1);
  columnNum = schedule.getMaxColumns();
}

//土日の色を変える
function paintWeekdays(baseLine, span, color){
  var wkendStart = 5;
  var range = schedule.getRange(1, baseLine, rowNum+1, span);
  range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  schedule.getRange(2, baseLine+wkendStart, rowNum-2+1, 2).setBackground(color);
  range.copyTo(schedule.getRange(1, baseLine+span, rowNum+1, columnNum-baseLine+span+1));
}

//祝日の色を変える
function paintHolidays(baseLine, date, color){
  try{
    var data = holiday.getRange(1, 1, holiday.getLastRow(), 1).getValues();
  }
  catch(e){
    Logger.log(e.message);
    Logger.log('祝日の適用なし');
    return;
  }

  for (var i = 0, len = data.length; i < len; i++){
    var diff = Moment.moment(data[i][0]).diff(date, 'days');
    if (0 <= diff && diff <= columnNum-baseLine){
      schedule.getRange(2, baseLine+diff, rowNum-2+1, 1).setBackground(color);
    };
  };
};


//ガントチャートのフォーマット
function formatGantchart(span, date) {
  var memo = PropertiesService.getDocumentProperties();
  var baseLine = findStartPoint('progress')+1;
  var date = Moment.moment(date);
  var format = memo.getProperty('format');
  var chartWidth = 168;
  memo.setProperty('baseDate', date.format(format));

  //列幅、列数、土日祝の色塗り
  adjustColums(baseLine, chartWidth, 25);
  paintWeekdays(baseLine, span, '#fcefe3');
  paintHolidays(baseLine, date, '#fcefe3');

  //シートに記入があればチャートに反映
  if (schedule.getLastRow() > 2){
    repaintChart(baseLine, date);
  };

  //日付の記入
  var calDate = date;
  var chartRange = schedule.getRange(1, baseLine, 1, columnNum-baseLine+1);
  var chartData = chartRange.getValues();
  chartRange.setHorizontalAlignment('left');
  chartData[0][0] = '(' + calDate.format(format) + ')';
  for (var i = 1, len = chartData[0].length; i < len; i++){
    if(i % 7 === 0){
      calDate = calDate.add(span, 'days');
      chartData[0][i] = '(' + calDate.format(format) + ')';
    };
  };
  chartRange.setValues(chartData);
};

//開始位置を見つける
function findStartPoint(text) {
  var ary = schedule.getRange('2:2').getValues();
  if (ary[0].indexOf(text) < 0) {
    Browser.msgBox('System Error: Please initalize the gantt chart.');
  } else {
   return ary[0].indexOf(text)+1;
 };
};

//フォーマットをリセット
function copyDefaultRow(top, left, height, width, option){
  var range = schedule.getRange(2, left, 1, columnNum-left+1);
  if (!option){
    range.copyTo(schedule.getRange(top, left, height, width));
  } else {
    range.copyTo(schedule.getRange(top, left, height, width), {contentsOnly:true})
  }
}

//予定開始でマイルストーンを置く
function setMilestone(top, baseLine, baseDate, startDate, finishDate, color){
  var chartFinish = baseLine + finishDate.diff(baseDate, 'days');
  if (chartFinish >= baseLine && chartFinish < columnNum){
    schedule.getRange(top, chartFinish).setBackground(color);
  };
};

//色ぬり機能
function paintChart(top, baseLine, baseDate, startDate, finishDate, color){
  var chartStart = baseLine + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;
  if(chartStart < baseLine){
    duration -= baseLine-chartStart;
    if(duration <= 0){
      return;
    } else {
      if(baseLine+duration > columnNum-baseLine+1) {
        //エラー処理入れる
        schedule.getRange(top, baseLine, 1, columnNum-baseLine+1).setBackground(color);
      } else {
        //エラー処理入れる
        schedule.getRange(top, baseLine, 1, duration).setBackground(color);
      }
    }
  };
  if (chartStart >= baseLine){
    if (chartStart+duration > columnNum-baseLine+1) {
      //エラー処理入れる
      schedule.getRange(top, chartStart, 1, columnNum-chartStart+1).setBackground(color);
    } else {
      //エラー処理入れる
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
function markProgress(top, baseLine, baseDate, startDate, finishDate, progress){
  var chartStart = baseLine + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;

  if(chartStart < baseLine){
    duration = columnNum-baseLine+1 > duration-(baseLine-chartStart) ? duration-(baseLine-chartStart) : columnNum-baseLine+1;
    if(duration <= 0){
      return;
    } else {
      if(baseLine+duration > columnNum-baseLine+1){
        var markLength = Math.round(duration * progress) > duration ? duration : Math.round(duration * progress);
        if(markLength === 0){return;}
        var progressLine = [];
        progressLine.push([]);
        for (var g = 0; g < markLength; g++) {
          progressLine[0].push("'=");
        };
        try{
          schedule.getRange(top, baseLine, 1, markLength).setValues(progressLine);
        } catch(e){
          Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
        }

      };
    };
  };

  if (chartStart >= baseLine){
    var markLength = Math.round(duration * progress) > duration ? duration : Math.round(duration * progress);
    if(markLength === 0){return;}
    var progressLine = [];
    progressLine.push([]);
    for (var g = 0; g < markLength; g++) {
      progressLine[0].push("'=");
    };
    try{
      schedule.getRange(top, chartStart, 1, markLength).setValues(progressLine);
    } catch(e){
      Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
    }

  };
};

//色の塗り直し
function repaintChart(baseLine, date){
  var memo = PropertiesService.getDocumentProperties();
  var date = Moment.moment(date);
  var lastRowOfContents = schedule.getLastRow();
  var data = schedule.getRange(2, 1, lastRowOfContents-2+1, baseLine-1).getValues();
  var indexOfPlannedStart = data[0].indexOf('plannedStart');
  var indexOfPlannedFinish = data[0].indexOf('plannedFinish');
  var indexOfActualStart = data[0].indexOf('actualStart');
  var indexOfActualFinish = data[0].indexOf('actualFinish');
  var indexOfProgress = data[0].indexOf('progress');

  copyDefaultRow(3, baseLine, lastRowOfContents-3+1, columnNum-baseLine+1);
  drawTodayLine();
  for (var i = 1, len = data.length; i < len; i++) {
    var plannedStart = Moment.moment(data[i][indexOfPlannedStart]);
    var plannedFinish = Moment.moment(data[i][indexOfPlannedFinish]);
    var actualStart = Moment.moment(data[i][indexOfActualStart]);
    var actualFinish = Moment.moment(data[i][indexOfActualFinish]);
    var progress = data[i][indexOfProgress];

    if(plannedStart.format('YYYY') === 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
      setMilestone(i+2, baseLine, date, plannedStart, plannedFinish, '#FFBB00');
    };
    if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
      if(plannedStart > plannedFinish){
        var lang = memo.getProperty('lang');
        var text = lang === 'ja' ? '予定開始が予定終了よりも大きな値です。' : 'Planned Start is bigger than Planned Finish.';
        Browser.msgBox('System Error (ID ' + schedule.getRange(i+2, 1).getValue() + ') : ' + text);
        return;
      }
      var color = judgeColor(plannedStart, plannedFinish, progress);
      if(color === ''){
        color = '#e3f0f9';
        schedule.getRange(i+2, indexOfProgress+1).setBackground('');
      } else {
        schedule.getRange(i+2, indexOfProgress+1).setBackground(color);
      }
      paintChart(i+2, baseLine, date, plannedStart, plannedFinish, color);
    }
    if(actualStart.format('YYYY') !== 'Invalid date' && actualFinish.format('YYYY') !== 'Invalid date'){
      if(actualStart > actualFinish){
        var lang = memo.getProperty('lang');
        var text = lang === 'ja' ? '実際開始が実際終了よりも大きな値です。' : 'Actual Start is bigger than Actual Finish.';
        Browser.msgBox('System Error (ID ' + schedule.getRange(row, 1).getValue() + ') : ' + text);
        return;
      }
      paintChart(i+2, baseLine, date, actualStart, actualFinish, '#aadca8');
    }
    if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date' && actualStart.format('YYYY') !== 'Invalid date' && actualFinish.format('YYYY') !== 'Invalid date'){
      var isOverlap = checkOverlap(plannedStart, plannedFinish, actualStart, plannedFinish);
      if (isOverlap !== false) {
        paintChart(i+2, baseLine, date, Moment.moment(isOverlap[0]), Moment.moment(isOverlap[1]), '#99c6ca');
      };
    };
    if (progress > 0 && plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date') {
      markProgress(i+2, baseLine, date, plannedStart, plannedFinish, progress);
    };
  };
};

//日付の線を引く
function drawTodayLine () {
 var memo = PropertiesService.getDocumentProperties();
 var today = Moment.moment();
 var baseLine = parseInt(findStartPoint('progress'))+1;
 var baseDate = Moment.moment(memo.getProperty('baseDate'));
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
   var todayColumn = schedule.getRange(2, todayLine, lastRowOfContents-2+1, 1);
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


//plannedWorkloadとprogressを合計する
function sumWnP(data, progressData, targetCol, taskId){
  Logger.log('起動: sumWnP');
  var memo = PropertiesService.getDocumentProperties();
  var taskIdAry = taskId.toString().split('_');
  var parentTasks = [];
  var lastIndex = data[0].length - 1;
  var indexOfPlannedStart = data[1].indexOf('plannedStart');
  var indexOfPlannedFinish = data[1].indexOf('plannedFinish');

  //親タスクを抽出してindexを調べる
  for (var i = 0, len = taskIdAry.length-1; i < len; i++){
    taskIdAry.pop();
    var tmp = taskIdAry.join('_');
    for (var j = 0, len2 = data.length; j < len2; j++){
      if(tmp == data[j][0]){
        parentTasks.push({
          'ID' : tmp,
          'index': j
        });
        progressData[j][0] = 0;

        data[j][lastIndex] = 0;
      };
    };
  };
  Logger.log(parentTasks);
  //工数（予）と進捗を計算していく
  for (var i = 0, len = parentTasks.length; i < len; i++){
    var earnedVal = 0;
  //親タスクに紐づく子タスクを全て再計算
  var parAry = parentTasks[i]['ID'].split('_');
  for (var j = parentTasks[i]['index'], len2 = data.length; j < len2; j++){
    var isChild = true;
    var judgedAry = data[j][0].toString().split('_');
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
          data[j][lastIndex] = '' === data[j][lastIndex] ? 0 : data[j][lastIndex];
          progressData[j][0] = '' === progressData[j][0] ? 0 : progressData[j][0];
          data[parentTasks[i]['index']][lastIndex] += data[j][lastIndex];
          earnedVal += data[j][lastIndex] * progressData[j][0];
        };
      };
    };
    //plannedWorkloadが全て0のとき用
    if(earnedVal === 0 && data[parentTasks[i]['index']][lastIndex] === 0){
      progressData[parentTasks[i]['index']][0] = 0;
    } else {
      progressData[parentTasks[i]['index']][0] = earnedVal / data[parentTasks[i]['index']][lastIndex];
    }
    schedule.getRange(parentTasks[i]['index']+1, lastIndex+1).setValue(data[parentTasks[i]['index']][lastIndex]);
    schedule.getRange(parentTasks[i]['index']+1, targetCol).setValue(progressData[parentTasks[i]['index']][0]);

    //日付が入っていればガントチャートへ進捗を反映
    var plannedStart = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedStart]);
    var plannedFinish = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedFinish]);
    if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
      copyDefaultRow(parentTasks[i]['index']+1, targetCol+1, 1, columnNum-targetCol+1, true);
      //進捗に合わせて色を変える
      var color = judgeColor(plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
      if(color === ''){
        color = '#e3f0f9';
        //進捗の行はオンスケなら白色
        schedule.getRange(parentTasks[i]['index']+1, targetCol).setBackground('');
      } else {
        schedule.getRange(parentTasks[i]['index']+1, targetCol).setBackground(color);
      }
      paintChart(parentTasks[i]['index']+1, targetCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, color);
      markProgress(parentTasks[i]['index']+1, targetCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
    };
  };
};


//全てのplannedWorkloadとprogressを合計する
function sumAllWnP(data, PWData, progressData, PWFormulas, progressFormulas, progressCol){
  Logger.log('起動: sumAllWnP');
  var memo = PropertiesService.getDocumentProperties();
  var parentTasks = [];
  var indexOfPlannedStart = data[1].indexOf('plannedStart');
  var indexOfPlannedFinish = data[1].indexOf('plannedFinish');
  //親タスクを抽出してindexを調べる
  for (var i = 0, len = data.length-1; i < len; i++){
    var tmp = data[i][0].toString() + '_1';
    for (var j = i, len2 = data.length; j < len2; j++){
      if(data[j][0].toString() === tmp){
        parentTasks.push({
          'ID': data[i][0].toString(),
          'index': i
        });
        PWData[i][0] = 0;
        progressData[i][0] = 0;
        PWFormulas[i][0] = '';
        progressFormulas[i][0] = '';
        break;
      };
    };
  };
  //親に紐づく子を判定して合計していく
  for (var i = parentTasks.length-1; 0 <= i; i--){
    var earnedVal = 0;
    var parAry = parentTasks[i]['ID'].split('_');
    for (var j = parentTasks[i]['index'], len = data.length; j < len; j++){
      var isChild = true;
      var judgedAry = data[j][0].toString().split('_');
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
            PWData[j][0] = '' === PWData[j][0] ? 0 : PWData[j][0];
            progressData[j][0] = '' === progressData[j][0] ? 0 : progressData[j][0];
            PWData[parentTasks[i]['index']][0] += PWData[j][0];
            earnedVal += PWData[j][0] * progressData[j][0];
          };
        };
      };
      //plannedWorkloadが全て0のとき用
      if(earnedVal === 0 && PWData[parentTasks[i]['index']][0] === 0){
        progressData[parentTasks[i]['index']][0] = 0;
      } else {
        progressData[parentTasks[i]['index']][0] = earnedVal / PWData[parentTasks[i]['index']][0];
      }

      //日付が入っていればガントチャートへ進捗を反映
      var plannedStart = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedStart]);
      var plannedFinish = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedFinish]);
      if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
        copyDefaultRow(parentTasks[i]['index']+1, progressCol+1, 1, columnNum-progressCol+1, true);
        //進捗に合わせて色を変える
        var color = judgeColor(plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
        if(color === ''){
          color = '#e3f0f9';
          //進捗の行はオンスケなら白色
          schedule.getRange(parentTasks[i]['index']+1, progressCol).setBackground('');
        } else {
          schedule.getRange(parentTasks[i]['index']+1, progressCol).setBackground(color);
        }
        paintChart(parentTasks[i]['index']+1, progressCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, color);
        markProgress(parentTasks[i]['index']+1, progressCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
      };
    };

    for(var i = 0, len = PWFormulas.length; i < len; i++){
      if(PWFormulas[i][0] !== ''){
        PWData[i][0] = PWFormulas[i][0];
      };
    };

    for(var i = 0, len = progressFormulas.length; i < len; i++){
      if(progressFormulas[i][0] !== ''){
        progressData[i][0] = progressFormulas[i][0];
      };
    };
    var result = {
      'PWData': PWData,
      'progressData': progressData
    };
    return result;
  }




//progressを合計する
function sumProgress(data, progressData, targetCol, taskId){
  Logger.log('起動: sumProgress');
  var memo = PropertiesService.getDocumentProperties();
  var taskIdAry = taskId.toString().split('_');
  var parentTasks = [];
  var lastIndex = data[0].length - 1;
  var indexOfPlannedStart = data[1].indexOf('plannedStart');
  var indexOfPlannedFinish = data[1].indexOf('plannedFinish');

  //親タスクを抽出してindexを調べる
  for (var i = 0, len = taskIdAry.length-1; i < len; i++){
    taskIdAry.pop();
    var tmp = taskIdAry.join('_');
    for (var j = 0, len2 = data.length; j < len2; j++){
      if(tmp == data[j][0]){
        parentTasks.push({
          'ID' : tmp,
          'index': j
        });
        progressData[j][0] = 0;
      };
    };
  };
  Logger.log(parentTasks);
  //工数（予）と進捗を計算していく
  for (var i = 0, len = parentTasks.length; i < len; i++){
    var earnedVal = 0;
  //親タスクに紐づく子タスクを全て再計算
  var parAry = parentTasks[i]['ID'].split('_');
  for (var j = parentTasks[i]['index'], len2 = data.length; j < len2; j++){
    var isChild = true;
    var judgedAry = data[j][0].toString().split('_');
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
          data[j][lastIndex] = '' === data[j][lastIndex] ? 0 : data[j][lastIndex];
          progressData[j][0] = '' === progressData[j][0] ? 0 : progressData[j][0];
          earnedVal += data[j][lastIndex] * progressData[j][0];
        };
      };
    };
    //plannedWorkloadが全て0のとき用
    if(earnedVal === 0 && data[parentTasks[i]['index']][lastIndex] === 0){
      progressData[parentTasks[i]['index']][0] = 0;
    } else {
      progressData[parentTasks[i]['index']][0] = earnedVal / data[parentTasks[i]['index']][lastIndex];
    }
    schedule.getRange(parentTasks[i]['index']+1, targetCol).setValue(progressData[parentTasks[i]['index']][0]);

    //日付が入っていればガントチャートへ進捗を反映
    var plannedStart = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedStart]);
    var plannedFinish = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedFinish]);
    if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
      copyDefaultRow(parentTasks[i]['index']+1, targetCol+1, 1, columnNum-targetCol+1, true);
      //進捗に合わせて色を変える
      var color = judgeColor(plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
      if(color === ''){
        color = '#e3f0f9';
        //進捗の行はオンスケなら白色
        schedule.getRange(parentTasks[i]['index']+1, targetCol).setBackground('');
      } else {
        schedule.getRange(parentTasks[i]['index']+1, targetCol).setBackground(color);
      }
      paintChart(parentTasks[i]['index']+1, targetCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, color);
      markProgress(parentTasks[i]['index']+1, targetCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
    };
  };
};


//全てのprogressを合計する
function sumAllProgress(data, progressData, progressFormulas, progressCol){
  Logger.log('起動: sumAllProgress');
  var memo = PropertiesService.getDocumentProperties();
  var parentTasks = [];
  var lastIndex = data[0].length - 1;
  var indexOfPlannedStart = data[1].indexOf('plannedStart');
  var indexOfPlannedFinish = data[1].indexOf('plannedFinish');

  //親タスクを抽出してindexを調べる
  for (var i = 0, len = data.length-1; i < len; i++){
    var tmp = data[i][0].toString() + '_1';
    for (var j = i, len2 = data.length; j < len2; j++){
      if(data[j][0].toString() === tmp){
        parentTasks.push({
          'ID': data[i][0].toString(),
          'index': i
        });
        progressData[i][0] = 0;
        progressFormulas[i][0] = '';
        break;
      };
    };
  };
  Logger.log(parentTasks);
  //親に紐づく子を判定して合計していく
  for (var i = parentTasks.length-1; 0 <= i; i--){
    var earnedVal = 0;
    var parAry = parentTasks[i]['ID'].split('_');
    for (var j = parentTasks[i]['index'], len = data.length; j < len; j++){
      var isChild = true;
      var judgedAry = data[j][0].toString().split('_');
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
            progressData[j][0] = '' === progressData[j][0] ? 0 : progressData[j][0];
            earnedVal += data[j][lastIndex] * progressData[j][0];
          };
        };
      };
      //plannedWorkloadが全て0のとき用
      if(earnedVal === 0 && data[parentTasks[i]['index']][lastIndex] === 0){
        progressData[parentTasks[i]['index']][0] = 0;
      } else {
        progressData[parentTasks[i]['index']][0] = earnedVal / data[parentTasks[i]['index']][lastIndex];
      }

      //日付が入っていればガントチャートへ進捗を反映
      var plannedStart = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedStart]);
      var plannedFinish = Moment.moment(data[parentTasks[i]['index']][indexOfPlannedFinish]);
      if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
        copyDefaultRow(parentTasks[i]['index']+1, progressCol+1, 1, columnNum-progressCol+1, true);
        //進捗に合わせて色を変える
        var color = judgeColor(plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
        if(color === ''){
          color = '#e3f0f9';
          //進捗の行はオンスケなら白色
          schedule.getRange(parentTasks[i]['index']+1, progressCol).setBackground('');
        } else {
          schedule.getRange(parentTasks[i]['index']+1, progressCol).setBackground(color);
        }
        paintChart(parentTasks[i]['index']+1, progressCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, color);
        markProgress(parentTasks[i]['index']+1, progressCol+1, Moment.moment(memo.getProperty('baseDate')), plannedStart, plannedFinish, progressData[parentTasks[i]['index']][0]);
      };
    };

    //数式を反映
    for(var i = 0, len = progressFormulas.length; i < len; i++){
      if(progressFormulas[i][0] !== ''){
        progressData[i][0] = progressFormulas[i][0];
      };
    };
    var result = {
      'progressData': progressData
    };
    return result;
  };

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
      return parseInt(broId) + 1;
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

//進捗に合わせて色を選ぶ
function judgeColor(start, finish, progress){
  var today = Moment.moment();
  var color = '';

  if(start.isSame(today, 'days') && progress < 1){
    color = '#ffff00'; //黄色
  };
  if(start.isBefore(today, 'days') && progress < 1){
    var actualProgress = finish.diff(start, 'days') * progress;
    var idealProgress = today.diff(start, 'days');
    if(actualProgress > idealProgress){
      color = '#ffff00'; //黄色
    } else {
    color = '#ff0000'; //赤色
  };
};
return color;
}

//祝日のセット
function setHolidays(data){
  var range = holiday.getRange(1, 1,data.length, data[0].length);
  range.setValues(data);
};

//日本の祝日を取得
function getHolidays() {
  //今年の1/1〜
  var startDate = new Date();
  startDate.setMonth(0, 1);
  startDate.setHours(0, 0, 0, 0);

  //来年の〜12/31
  var endDate = new Date();
  endDate.setFullYear(endDate.getFullYear()+1, 11, 31);
  endDate.setHours(0, 0, 0, 0);

  var calendar = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  var holidays = calendar.getEvents(startDate, endDate);
  var values = [];
  for(var i = 0; i < holidays.length; i++) {
    values[i] = [holidays[i].getStartTime(), holidays[i].getTitle()];
  }
  return values;
};


//初期値する
function resetAll(msg){
  var isDelete = Browser.msgBox(msg, Browser.Buttons.YES_NO);
  if(isDelete === 'yes'){
    schedule.clear();
    holiday.clear();
    init();
    var data = getHolidays();
    var date = Moment.moment();
    var tmp = 0;
    while (Math.abs(date.day()) + tmp <= 7) {
      tmp++;
    };
    var monday = date.add(tmp, 'days');
    formatGantchart(7, monday);
    setHolidays(data);
  };
};

//サイドバーの表示
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
  .setTitle('Gantt Chart Generator')
  .setWidth(300);
  SpreadsheetApp.getUi()
  .showSidebar(html);
};



/*↓ クライアント用functions ↓*/

function front_repaintChart(){
  var memo = PropertiesService.getDocumentProperties();
  var baseLine = findStartPoint('progress')+1;
  var date = Moment.moment(memo.getProperty('baseDate'));
  repaintChart(baseLine, date);

}

function front_sumAllWnP(){
  var lastRowOfContents = schedule.getLastRow();
  var progressCol = findStartPoint('progress');
  var data = schedule.getRange(1, 1, lastRowOfContents, findStartPoint('plannedWorkload')).getValues();
  var PWRange = schedule.getRange(1, data[0].length, lastRowOfContents, 1);
  var progressRange = schedule.getRange(1, progressCol, lastRowOfContents, 1);
  var PWData = PWRange.getValues();
  var progressData = progressRange.getValues();
  var formulas = PWRange.getFormulas();
  var progressFormulas = progressRange.getFormulas();
  var result = sumAllWnP(data, PWData, progressData, formulas, progressFormulas, progressCol);
  PWRange.setValues(result.PWData);
  progressRange.setValues(result.progressData);
}


function front_sumAllProgress(){
  var lastRowOfContents = schedule.getLastRow();
  var progressCol = findStartPoint('progress');
  var data = schedule.getRange(1, 1, lastRowOfContents, findStartPoint('plannedWorkload')).getValues();
  var progressRange = schedule.getRange(1, progressCol, lastRowOfContents, 1);
  var progressData = progressRange.getValues();
  var progressFormulas = progressRange.getFormulas();
  var result = sumAllProgress(data, progressData, progressFormulas, progressCol);
  progressRange.setValues(result.progressData);
}



//権限でエラーが出るのでdrawTodayLineはonOpenで動かす
//トリガーのセット



//drawTodayLineの重複を許さない形でTriggerをセット
function set_front_repaintChart(){

  function createTimeDrivenTriggers() {
    ScriptApp.newTrigger('front_repaintChart')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
  }

  var triggers = ScriptApp.getUserTriggers(ss);
  for (var i = 0, len = triggers.length; i < len; i++){
    if(triggers[i].getHandlerFunction() === 'front_repaintChart'){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
  createTimeDrivenTriggers();
};
