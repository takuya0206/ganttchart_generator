
function getSpreadSheet(){
  if(getSpreadSheet.ss){return getSpreadSheet.ss; };
  getSpreadSheet.ss = SpreadsheetApp.getActive();
  return getSpreadSheet.ss;
};


function getScheduleSheet(){
  var ss = getSpreadSheet();
  if(getScheduleSheet.s_sheet){return getScheduleSheet.s_sheet; };
  getScheduleSheet.s_sheet = ss.getSheetByName('schedule');
  return getScheduleSheet.s_sheet;
};


function getHolidaySheet(){
  var ss = getSpreadSheet();
  if(getScheduleSheet.h_sheet){return getScheduleSheet.h_sheet;};
  getHolidaySheet.h_sheet = ss.getSheetByName('holiday');
  return getHolidaySheet.h_sheet;
};


function askEnabled(){
  var lang = Session.getActiveUserLocale();
  var title = 'Gantt Chart Generator';
  var msg = lang === 'ja' ? 'Gantt Chart Generatorが有効になりました。もしアドオンのメニューに「ガントチャートの作成」が表示されていない場合は一度リロードをお願いします。' : 'Gantt Chart Generator has been enabled. Just in case that the menu: "Create Gantt Chart" does not appear, please reload this spreadsheet.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
};


function showSidebar() {
  Logger.log('showSidebar start');
  var html = HtmlService.createHtmlOutputFromFile('Page')
  .setTitle('Gantt Chart Generator')
  .setWidth(300);
  SpreadsheetApp.getUi().showSidebar(html);
};


function createChart(){
  Logger.log('createChart start');
  var schedule = getScheduleSheet();
  var holiday = getHolidaySheet();
  var ss = getSpreadSheet();
  var memo = PropertiesService.getDocumentProperties();
  if(!holiday){
    try{
      ss.insertSheet('holiday', 2);
    } catch(e){
      Logger.log(e.message);
    };
  };
  if(!schedule){
    var lang = memo.getProperty('lang');
    var text = lang === 'ja' ? 'ガントチャートの作成を行いますか？' : 'Will you create a gantt chart?';
    try{
      ss.insertSheet('schedule', 1);
    } catch(e){
      Logger.log(e.message);
    };
    resetAll(text);
  } else {
    var lang = memo.getProperty('lang');
    var text = lang === 'ja' ? '既にscheduleシートが存在しています。新たに作成をするとこれまでの内容が削除されますがよろしいですか？' : 'You already have the schedule sheet. Please confirm that the existing contents will be deleted if you create a new gantt chart.';
    resetAll(text);
    };
};


function resetAll(msg){
  Logger.log('resetAll start');
  var ss = getSpreadSheet();
  var schedule = getScheduleSheet();
  var holiday = getHolidaySheet();
  var isComfirmed = Browser.msgBox(msg, Browser.Buttons.YES_NO);
  if(isComfirmed === 'yes'){
    schedule.clear();
    holiday.clear();
    init();
    var data = getHolidays();
    setHolidays(data);
    //start date should be Monday
    var date = Moment.moment();
    var tmp = 0;
    while (Math.abs(date.day()) + tmp <= 7) {
      tmp++;
    };
    var monday = date.add(tmp, 'days');
    formatGantchart(7, monday.format('YYYY/MM/DD'));
    makeSampleTask(monday);
  };
};


function init(){
  Logger.log('init start');
  //template for items
  var schedule = getScheduleSheet();
  var holiday = getHolidaySheet();
  var memo = PropertiesService.getDocumentProperties();
  var lang = memo.getProperty('lang');
  var scheduleItems = lang === 'ja' ? [['タスクID', '階層別 タスク一覧','','','','', '予定開始', '予定終了', '実際開始', '実際終了', '工数\n（予｜実）','', '担当', '進捗'],
  ['id', 'lv1','lv2','lv3','lv4','lv5','plannedStart', 'plannedFinish', 'actualStart', 'actualFinish', 'plannedWorkload', 'actualWorkload', 'responsiblity', 'progress']
  ] :
  [['Task ID', 'Work Breakdown Structure','','','','', 'Planned Start', 'Planned Finish', 'Actual Start', 'Actual Finish', 'Workload\n（Plan｜Actual）','', 'In Charge', 'Progress'],
  ['id', 'lv1','lv2','lv3','lv4','lv5','plannedStart', 'plannedFinish', 'actualStart', 'actualFinish', 'plannedWorkload', 'actualWorkload', 'responsiblity', 'progress']
  ];
  var note = lang === 'ja' ? '手動で祝日を編集するときは、必ず日付をA列に入力してください。': 'When editing holidays, please set date only in the A column.';
  var format = lang === 'ja' ? 'YYYY/MM/DD' : 'MMM-Do-YY';
  var cellformat = lang === 'ja' ? 'yyyy/mm/dd':  'MMM-d-yyyy';
  var scheduleItemsLength = scheduleItems[0].length;
  var firstRow = schedule.getRange('1:1');
  var range = schedule.getRange(1, 1, 2, scheduleItemsLength);
  var indexOfPlannedStart = scheduleItems[1].indexOf('plannedStart');
  var indexOfPlannedFinish = scheduleItems[1].indexOf('plannedFinish');
  var indexOfActualStart = scheduleItems[1].indexOf('actualStart');
  var indexOfActualFinish = scheduleItems[1].indexOf('actualFinish');
  var indexOfPlannedWorkload = scheduleItems[1].indexOf('plannedWorkload');
  var indexOfActualWorkload = scheduleItems[1].indexOf('actualWorkload');
  var indexOfProgress = scheduleItems[1].indexOf('progress');
  var rowNum = schedule.getMaxRows();
  var columnNum = schedule.getMaxColumns();
  memo.setProperties({
    'format': format,
    'cellformat': cellformat
  });
  //dataValidation
  var rule_date = SpreadsheetApp.newDataValidation().requireDate().build();
  var rule_int = SpreadsheetApp.newDataValidation().requireNumberGreaterThan(-1).build();
  //format related to items
  schedule.getRange('A:A').setHorizontalAlignment('left').setBackground('#f3f3f3');
  firstRow.setBackground('#f3f3f3');
  schedule.getRange(1, indexOfPlannedStart+1, rowNum, columnNum-indexOfPlannedStart).setHorizontalAlignment('center');
  range.setValues(scheduleItems);
  schedule.setFrozenRows(1);
  schedule.setFrozenColumns(indexOfPlannedStart);
  schedule.hideRows(2);
  schedule.hideColumns(indexOfActualStart+1, 2);
  schedule.setColumnWidth(1, 70);
  for(var i = 2; i <= 5; i++){
    schedule.setColumnWidth(i, 20);
  };
  schedule.setColumnWidth(indexOfPlannedWorkload+1, 45);
  schedule.setColumnWidth(indexOfActualWorkload+1, 45);
  schedule.getRange(3, indexOfPlannedWorkload+1,rowNum-3+1, 2).setDataValidation(rule_int);
  schedule.getRange(1, indexOfPlannedWorkload+1, 1, 2).merge();
  schedule.getRange(3, indexOfProgress+1, rowNum-3+1, 1).setNumberFormat('0.0%').setDataValidation(rule_int);
  schedule.getRange(3, indexOfPlannedStart+1, rowNum-3+1, 4).setNumberFormat(cellformat).setDataValidation(rule_date);
  holiday.getRange(1, 1).setNote(note);
  holiday.getRange('A:A').setNumberFormat(cellformat).setDataValidation(rule_date);
  //adjust the number of rows
  if(rowNum > 200){
    schedule.deleteRows(200, rowNum-200);
  }
  //related triggers
  setDailyTiggers();
};


function setDailyTiggers(){
  Logger.log('setDailyTiggers start');
  var ss = getSpreadSheet();
  //GAS-based function
  function createTimeDrivenTriggers() {
    ScriptApp.newTrigger('front_updateChart')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();
  };

  //delete exiting triggers
  var triggers = ScriptApp.getUserTriggers(ss);
  for (var i = 0, len = triggers.length; i < len; i++){
    var tmp = triggers[i].getHandlerFunction();
    if(tmp === 'front_updateChart'){
      ScriptApp.deleteTrigger(triggers[i]);
    };
  };
  createTimeDrivenTriggers();
};


function makeSampleTask(date){
  Logger.log('makeSampleTask start')
  date = Moment.moment(date);
  var schedule = getScheduleSheet();
  var memo = PropertiesService.getDocumentProperties();
  var range = schedule.getRange(4, 1, 1, 14);
  range.setValues([['1', 'Sample Task','','','','', date.format(memo.getProperty('format')), date.add(5, 'days').format(memo.getProperty('format')), '','' , 10,'','' , 1]])
  front_updateChart();
};



function updateChart(data, startRow, endRow, baseLine, baseDate){
  Logger.log('updateChart start');
  var schedule = getScheduleSheet();
  var indexOfPlannedStart = data[1].indexOf('plannedStart');
  var indexOfPlannedFinish = data[1].indexOf('plannedFinish');
  var indexOfActualStart = data[1].indexOf('actualStart');
  var indexOfActualFinish = data[1].indexOf('actualFinish');
  var indexOfPlannedWorkload = data[1].indexOf('plannedWorkload');
  var indexOfActualWorkload = data[1].indexOf('actualWorkload');
  var indexOfProgress = data[1].indexOf('progress');
  var columnNum = schedule.getMaxColumns();
  copyDefaultRow(startRow, baseLine, endRow-startRow+1, columnNum-baseLine+1); //initalize the target range
  for (var i = startRow; i <= endRow; i++){
    var index = i-1;
    var plannedStart = Moment.moment(data[index][indexOfPlannedStart]);
    var plannedFinish = Moment.moment(data[index][indexOfPlannedFinish]);
    var actualStart = Moment.moment(data[index][indexOfActualStart]);
    var actualFinish = Moment.moment(data[index][indexOfActualFinish]);
    var progress = data[index][indexOfProgress];
    //if only planned finish is filled, paint a cell in orange(#FFBB00)
    if (plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') === 'Invalid date') {
      setMilestone(i, baseLine, baseDate, plannedStart, plannedFinish, '#FFBB00', columnNum);
    };
    //if planned start and planned finish are filled, paint a range in blue(#e3f0f9)
    if (plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date') {
      if(plannedStart > plannedFinish){
        showDateErrorMsg(i);
        return;
      };
      var color = judgeColor(plannedStart, plannedFinish, progress); //show alert color based on the progress
      //if the progress is on schedule, don't paint the progress column
      if(color === ''){
        color = '#e3f0f9'; //blue
        schedule.getRange(i, indexOfProgress+1).setBackground('');
      } else {
        schedule.getRange(i, indexOfProgress+1).setBackground(color);
      }
      paintChart(i, baseLine, baseDate, plannedStart, plannedFinish, color, columnNum);
    };
    //if actual start and actual finish are filled, paint a range in green(#aadca8)
    if (actualFinish.format('YYYY') !== 'Invalid date' && actualStart.format('YYYY') !== 'Invalid date'){
      if(actualStart > actualFinish){
        showDateErrorMsg(i);
        return;
      };
      paintChart(i, baseLine, baseDate, actualStart, actualFinish, '#aadca8', columnNum);
    };
    //if there is overlap, make it strong green（#99c6ca）
    if (plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date' && actualFinish.format('YYYY') !== 'Invalid date' && actualStart.format('YYYY') !== 'Invalid date'){
      var isOverlap = checkOverlap(plannedStart, plannedFinish, actualStart, actualFinish);
      if (isOverlap !== false) {
        paintChart(i, baseLine, baseDate, isOverlap[0], isOverlap[1], '#99c6ca', columnNum);
      };
    };
    //mark progress by using '='
    if (progress >= 0 && plannedFinish.format('YYYY') !== 'Invalid date' && plannedStart.format('YYYY') !== 'Invalid date') {
      markProgress(i, baseLine, baseDate, plannedStart, plannedFinish, progress);
    };
  };
};


function showDateErrorMsg(row){
  Logger.log('showDateErrorMsg start');
  var memo = PropertiesService.getDocumentProperties();
  var schedule = getScheduleSheet();
  var lang = memo.getProperty('lang');
  var text = lang === 'ja' ? '開始が終了よりも大きな値です。' : 'The start date is bigger than the finish date.';
  Browser.msgBox('System Error (ID ' + schedule.getRange(row, 1).getValue() + ') : ' + text);
  return;
};


function copyDefaultRow(top, left, height, width, option){
  Logger.log('copyDefaultRow start');
  var schedule = getScheduleSheet();
  var range = schedule.getRange(2, left, 1, width);
  if (!option){
    range.copyTo(schedule.getRange(top, left, height, width));
  } else {
    range.copyTo(schedule.getRange(top, left, height, width), {contentsOnly:true})
  };
};


function setMilestone(top, baseLine, baseDate, startDate, finishDate, color, columnNum){
  Logger.log('setMilestone start');
  var schedule = getScheduleSheet();
  var chartFinish = baseLine + finishDate.diff(baseDate, 'days');
  if (chartFinish >= baseLine && chartFinish <= columnNum){
    schedule.getRange(top, chartFinish).setBackground(color);
  };
};


function judgeColor(start, finish, progress){
  Logger.log('judgeColor start');
  var today = Moment.moment();
  var color = '';
  var memo = PropertiesService.getDocumentProperties();
  var isRequired = memo.getProperty('colorIndicator');
  if(isRequired === 'true'){
    if(start.isSame(today, 'days') && progress < 1){
    color = '#ffff00'; //yellow
  };
  if(start.isBefore(today, 'days') && progress < 1){
    var actualProgress = finish.diff(start, 'days') * progress;
    var idealProgress = today.diff(start, 'days');
    if(actualProgress >= idealProgress){
        color = '#ffff00'; //yellow
      } else {
        color = '#ff0000'; //red
      };
    };
  };
  return color;
};


function paintChart(top, baseLine, baseDate, startDate, finishDate, color, columnNum){
  Logger.log('paintChart start');
  var schedule = getScheduleSheet();
  var chartStart = baseLine + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;
  if(chartStart < baseLine){
    duration -= baseLine-chartStart;
    if(duration <= 0){
      return;
    } else {
      if(baseLine+duration > columnNum){
        try{
          schedule.getRange(top, baseLine, 1, columnNum-baseLine+1).setBackground(color);
        } catch(e){
          Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
        };
      } else {
        try{
          schedule.getRange(top, baseLine, 1, duration).setBackground(color);
        } catch(e){
          Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
        };
      };
    };
  };
  if (chartStart >= baseLine){
    if (chartStart+duration > columnNum){
      try{
        schedule.getRange(top, chartStart, 1, columnNum-chartStart+1).setBackground(color);
      } catch(e){
        Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
      };
    } else {
      try{
        schedule.getRange(top, chartStart, 1, duration).setBackground(color);
      } catch(e){
        Browser.msgBox('System Error (ID ' + schedule.getRange(top, 1).getValue() + ') : ' + e.message);
      };
    };
  };
};


function checkOverlap(firstStart, firstFinish, secondStart, secondFinish) {
  Logger.log('checkOverlap start');
  if(firstStart <= secondFinish && firstFinish >= secondStart) {
   var isBigger = firstStart >= secondStart ? firstStart : secondStart;
   var isSmaller = firstFinish <= secondFinish ? firstFinish : secondFinish;
   return [isBigger, isSmaller];
 } else {
   return false;
 };
};

function markProgress(top, baseLine, baseDate, startDate, finishDate, progress){
  Logger.log('markProgress start');
  var schedule = getScheduleSheet();
  var columnNum = schedule.getMaxColumns();
  var chartStart = baseLine + startDate.diff(baseDate, 'days');
  var duration = finishDate.diff(startDate, 'days')+1;
  if(chartStart < baseLine){
    duration = columnNum-baseLine+1 > duration-(baseLine-chartStart) ? duration-(baseLine-chartStart) : columnNum-baseLine+1;
    Logger.log('duration:' + duration);
    if(duration <= 0){
      return;
    } else {
      if(baseLine+duration < columnNum-baseLine+1){
        var markLength = Math.round(duration * progress) > duration ? duration : Math.round(duration * progress);
        Logger.log('markLength:' + markLength);
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
        };
      };
    };
  };
  if (chartStart >= baseLine){
    var markLength = Math.round(duration * progress) > duration ? duration : Math.round(duration * progress);
    Logger.log('duration:' + duration);
    Logger.log('markLength:' + markLength);
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
    };
  };
};


function drawTodayLine() {
 Logger.log('drawTodayLine start');
 var schedule = getScheduleSheet();
 var memo = PropertiesService.getDocumentProperties();
 var today = Moment.moment();
 var baseLine = findStartPoint('progress')+1;
 var baseDate = Moment.moment(memo.getProperty('baseDate'));
 var lastRowOfContents = schedule.getLastRow();
 var nextBaseLine = baseLine + 1;
 var todayLine = baseLine + today.diff(baseDate, 'days');
 var columnNum = schedule.getMaxColumns();
 //delete an old line
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
 //drow a new line
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


function findStartPoint(text) {
  Logger.log('findStartPoint start');
  var schedule = getScheduleSheet();
  var ary = schedule.getRange('2:2').getValues();
  if (ary[0].indexOf(text) < 0) {
    Browser.msgBox('System Error: Please initalize your gantt chart.');
    return false;
  } else {
    return ary[0].indexOf(text)+1;
  };
};


function setHolidays(data){
  Logger.log('setHolidays start');
  var holiday = getHolidaySheet();
  var range = holiday.getRange(1, 1, data.length, data[0].length);
  range.setValues(data);
};


function getHolidays(){
  Logger.log('getHolidays start');
  var values = [];
  var startDate = new Date();//from 1st Jan in the current year
  startDate.setMonth(0, 1);
  startDate.setHours(0, 0, 0, 0);
  var endDate = new Date();//until 31st Dec in the next year
  endDate.setFullYear(endDate.getFullYear()+1, 11, 31);
  endDate.setHours(0, 0, 0, 0);
  //Japanese holidays
  var calendar = CalendarApp.getCalendarById("ja.japanese#holiday@group.v.calendar.google.com");
  if(!calendar){
    values[0] = [startDate, 'Sample Holiday']
  } else {
    var holidays = calendar.getEvents(startDate, endDate);
    for(var i = 0; i < holidays.length; i++) {
      values[i] = [holidays[i].getStartTime(), holidays[i].getTitle()];
    };
  };
  return values;
};


function formatGantchart(span, date) {
  Logger.log('formatGantchart');
  var schedule = getScheduleSheet();
  var memo = PropertiesService.getDocumentProperties();
  var baseLine = findStartPoint('progress')+1;
  var date = Moment.moment(date);
  var format = memo.getProperty('format');
  var chartWidth = 168;
  var rowNum = schedule.getMaxRows();
  var columnNum = schedule.getMaxColumns();
  memo.setProperty('baseDate', date.format('YYYY/MM/DD'));
  //The number and the width of rows
  adjustColums(baseLine, chartWidth, 25, rowNum, columnNum);
  columnNum = schedule.getMaxColumns();
  //Change the color in weekdays and holidays
  paintWeekdays(baseLine, span, '#fcefe3', rowNum, columnNum);
  paintHolidays(baseLine, date, '#fcefe3', rowNum, columnNum);
  //if the schedule sheet has some contents...
  if (schedule.getLastRow() > 2){
    front_updateChart();
  };
  //write date
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


function adjustColums(baseLine, num, width, rowNum, columnNum){
  Logger.log('adjustColums start');
  var schedule = getScheduleSheet();
  var deleteNum = columnNum - baseLine;
  schedule.setColumnWidth(baseLine, width);
  schedule.getRange(1, baseLine, rowNum, 1).clearContent();
  schedule.getRange(2, baseLine, rowNum-2+1, 1).setBackground('');
  schedule.deleteColumns(baseLine+1,deleteNum);
  schedule.insertColumnsAfter(baseLine, num-1);
};


function paintWeekdays(baseLine, span, color, rowNum, columnNum){
  Logger.log('paintWeekdays start');
  var wkendStart = 5;
  var schedule = getScheduleSheet();
  var range = schedule.getRange(1, baseLine, rowNum+1, span);
  range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
  schedule.getRange(2, baseLine+wkendStart, rowNum-2+1, 2).setBackground(color);
  range.copyTo(schedule.getRange(1, baseLine+span, rowNum+1, columnNum-baseLine+span+1));
};


function paintHolidays(baseLine, date, color, rowNum, columnNum){
  Logger.log('paintHolidays start');
  var schedule = getScheduleSheet();
  var holiday = getHolidaySheet();
  try{
    var data = holiday.getRange(1, 1, holiday.getLastRow(), 1).getValues();
  }
  catch(e){
    Logger.log('No holidays: ' + e.message);
    return;
  };
  for (var i = 0, len = data.length; i < len; i++){
    var diff = Moment.moment(data[i][0]).diff(date, 'days');
    if (0 <= diff && diff <= columnNum-baseLine){
      schedule.getRange(2, baseLine+diff, rowNum-2+1, 1).setBackground(color);
    };
  };
};


function findParentTasks(data, baseId){
  var parentTasks = [];
  if(baseId){ //find parent tasks related to baseId
    var taskIdAry = baseId.toString().split('_');
    for (var i = 0, len = taskIdAry.length-1; i < len; i++){
      taskIdAry.pop();
      var tmp = taskIdAry.join('_');
      for (var j = 0, len2 = data.length; j < len2; j++){
        if(tmp == data[j][0]){
          parentTasks.push({
            'ID' : tmp,
            'index': j
          });
        };
      };
    };
 } else { //find all parent tasks
   for (var i = 0, len = data.length-1; i < len; i++){
    var tmp = data[i][0].toString() + '_1';
    for (var j = i, len2 = data.length; j < len2; j++){
      if(data[j][0].toString() === tmp){
        parentTasks.push({
          'ID': data[i][0].toString(),
          'index': i
        });
        break;
      };
    };
  };
};
return parentTasks;
};


function sumTwoColumns(data, formulas, workloadCol, progressCol, parentTasks, baseDate){
  Logger.log('sumTwoColumns start');
  var schedule = getScheduleSheet();
  for (var i = parentTasks.length-1; 0 <= i ; i--){
    var earnedVal = 0;
    var parAry = parentTasks[i]['ID'].toString().split('_');
    var currentIndex = parentTasks[i]['index'];
    var indexOfPlannedStart = data[1].indexOf('plannedStart');
    var indexOfPlannedFinish = data[1].indexOf('plannedFinish');
    data[currentIndex][workloadCol] = 0;
    data[currentIndex][progressCol] = 0;
    //find and calculate child tasks
    for (var j = currentIndex, len2 = data.length; j < len2; j++){
      var isChild = true;
      var judgedAry = data[j][0].toString().split('_');
        //screening based on the first value and the length
        if(parAry[0] === judgedAry[0] && parAry.length === judgedAry.length-1){
          //judge based on the whole value
          for (var k = 0, len3 = parAry.length; k < len3; k++){
            if(parAry[k] !== judgedAry[k]){
              isChild = false;
              break;
            };
          };
          if(isChild){
            //set 0 if the value is NaN
            data[j][workloadCol] = false === isNum(data[j][workloadCol]) ? 0 : data[j][workloadCol];
            data[j][progressCol] = false === isNum(data[j][progressCol]) ? 0 : data[j][progressCol];

            data[currentIndex][workloadCol] += data[j][workloadCol];
            earnedVal += data[j][workloadCol] * data[j][progressCol];
          };
        };
      };
      //handle "0 divided by 0 equal error"
      if(earnedVal == 0 && data[currentIndex][workloadCol] == 0){
        data[currentIndex][progressCol] = 0;
      } else {
        data[currentIndex][progressCol] = earnedVal / data[currentIndex][workloadCol];
      };
      //if data has planned start and finish, update the gannt chart
      var plannedStart = Moment.moment(data[currentIndex][indexOfPlannedStart]);
      var plannedFinish = Moment.moment(data[currentIndex][indexOfPlannedFinish]);
      if(plannedStart.format('YYYY') !== 'Invalid date' && plannedFinish.format('YYYY') !== 'Invalid date'){
        updateChart(data, currentIndex+1, currentIndex+1, progressCol+2, baseDate);
      };
    };
  //reflect the exiting formulas
  for(var i = 0, len = formulas.length; i < len; i++){
    for(var j = 0, len2 = formulas[0].length; j < len2; j++){
      if(formulas[i][j] !== ''){
        data[i][j] = formulas[i][j];
      };
    };
  };
  return data;
};


function isNum(val){
  val = String(val).replace(/^[ 　]+|[ 　]+$/g, '');
  if(val.length == 0){
    return false;
  };
  if(isNaN(val) || !isFinite(val)){
    return false;
  };
  return true;
};

function createTaskId(baseData, taskData, taskEndLine, startRow){
  Logger.log('createTaskId start');
  var taskPos = {row: '', col: ''};
  var index = startRow-1;
  for(var i = index, len = baseData.length; i < len; i++){
    var broId = 0;
    var parId = 0;
    var isBro = false;
    var isPar = false;
    var distanceToBro = 0;
    var distanceToPar = 0;
    var result = 0;
    //if there is a value in the i row...
    for(var j = 1; j < taskEndLine; j++){
      if(baseData[i][j] != ''){
        taskPos.row = i;
        taskPos.col = j;
        break;
      };
    };
    if(taskPos.row && taskPos.col){
      //check there is a brother task
      for(var j = taskPos.row-1; j >= 2; j--){
        if(baseData[j][taskPos.col] !== ''){
          broId = taskData[j][0];
          isBro = true;
          break;
        }
        distanceToBro += 1;
      };
      //check there is a parent task
      label_innerFor:
      for (var j = taskPos.row-1; j >= 2; j--){
        if(taskPos.col === 1){
          distanceToPar = Math.pow(2, 53)-1; //max value
          break;
        };
        for(var k = taskPos.col-1; k > 0; k--){
          if(baseData[j][k] !== ''){
            parId = baseData[j][0];
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

      //judgement
      if (isBro === false && isPar === false){
        result = 1;
      };
      if (distanceToBro < distanceToPar){
        if(broId.toString().length === 1){
          result = parseInt(broId) + 1;
        } else {
          var ary = broId.toString().split('_');
          ary[ary.length-1] = parseInt(ary[ary.length-1]) + 1;
          var str = ary.join('_');
          result = str;
        };
      };
      if (distanceToBro > distanceToPar){
        result = parId + '_' + 1;
      };
      baseData[i][0] = result;
      taskData[i][0] = result;
    };
  };
  var output = {'baseData': baseData, 'taskData': taskData};
  return output;
};


function makeParentBold(data, range){
  var info = [];
  //make a two-dimensional array for fontWeight
  for(var i = 0, len = data.length; i < len; i++){
    info.push([]);
    for (j = 0, len2 = data[0].length; j < len2; j++){
      info[i].push('normal');
    };
  };
  //reflect baseData on the array if it's a parent task
  for (var i = 2, len = data.length; i < len; i++){
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
