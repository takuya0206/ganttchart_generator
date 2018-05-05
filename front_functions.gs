
function front_updateChart(){
  Logger.log('front_updateChart start')
  var schedule = getScheduleSheet();
  var baseLine = findStartPoint('progress')+1;
  var baseDate = Moment.moment(schedule.getRange(1, baseLine).getValue());
  var startRow = 3;
  var endRow = schedule.getLastRow();
  var data = schedule.getRange(1, 1, endRow, baseLine-1).getValues();
  var parentTasks = findParentTasks(data);
  if(startRow > endRow){return;}; //No contents
  updateChart(data, startRow, endRow, baseLine, baseDate, parentTasks);
  drawTodayLine();
};

function front_sumAllWnP(){
  Logger.log('front_sumAllWnP start');
  var schedule = getScheduleSheet();
  var lastRowOfContents = schedule.getLastRow();
  var baseLine = findStartPoint('progress')+1;
  var baseRange = schedule.getRange(1, 1, lastRowOfContents, baseLine-1);
  var baseData = baseRange.getValues();
  var baseDate = Moment.moment(schedule.getRange(1, baseLine).getValue());
  var formulas = baseRange.getFormulas();
  var indexOfPlannedWorkload = baseData[1].indexOf('plannedWorkload');
  var indexOfProgress = baseData[1].indexOf('progress');
  var parentTasks = findParentTasks(baseData);
  var newData = sumTwoColumns(baseData, formulas, indexOfPlannedWorkload, indexOfProgress, parentTasks, baseDate);
  baseRange.setValues(newData);
}

function showPrompt(){
  Logger.log('showPrompt start');
  var memo = PropertiesService.getDocumentProperties();
  var lang = memo.getProperty('lang');
  var htmlString = '';
  var msg = lang === 'ja' ? '投げ銭のお願い' : 'Donation'
  if(lang === 'ja'){
    htmlString =
    '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">' +
    '<div>' +
    '<p>よろしければ投げ銭をお願いします。Gantt Chart Generatorはフリーツールですが、ユーザーの皆様の暖かい支援によって開発が成り立っています。どうぞよろしくお願いします。</p>' +
    '<br />' +
    '<input type="button" class="share" value="Amazon" onclick="window.open(\'http://amzn.asia/bAlH4Wk\')">　 ' +
    '<input type="button" class="action" value="PayPal" onclick="window.open(\'https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=YUYUT2MH5UFA8\')">　 ' +
    '<input type="button" value="Close" onclick="google.script.host.close()">' +
    '</div>';
  } else {
    htmlString =
    '<link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons1.css">' +
    '<div>' +
    '<p>Please pay add-on fee if you like Gantt Chart Generator. We appreciate your warm and kind support.</p>' +
    '<br />' +
    '<input type="button" class="action" value="PayPal" onclick="window.open(\'https://www.paypal.com/cgi-bin/webscr?cmd=_s-xclick&hosted_button_id=YUYUT2MH5UFA8\')">  ' +
    ' <input type="button" value="Close" onclick="google.script.host.close()">' +
    '</div>';
  }

  var htmlOutput = HtmlService
  .createHtmlOutput(htmlString)
  .setSandboxMode(HtmlService.SandboxMode.IFRAME)
  .setHeight(140);

  SpreadsheetApp
  .getUi()
  .showModalDialog(htmlOutput, msg);
};

function colorIndicator(isChecked){
  Logger.log('colorIndicator start');
  Logger.log(isChecked);
  var memo = PropertiesService.getDocumentProperties();
  memo.setProperty('colorIndicator', isChecked);
  front_updateChart();
};

function get_colorIndicator(){
  var memo = PropertiesService.getDocumentProperties();
  var isChecked = memo.getProperty('colorIndicator');
  isChecked = (isChecked == null || isChecked == '') ? false : isChecked;
  return isChecked;
};


function front_showParentChart(isChecked){
  Logger.log('front_showParentChart');
  Logger.log(isChecked);
  var memo = PropertiesService.getDocumentProperties();
  memo.setProperty('ParentChart', isChecked);
  var schedule = getScheduleSheet();
  var lastRowOfContents = schedule.getLastRow();
  var baseLine = findStartPoint('progress')+1;
  var baseRange = schedule.getRange(1, 1, lastRowOfContents, baseLine-1);
  var baseData = baseRange.getValues();
  var baseDate = Moment.moment(schedule.getRange(1, baseLine).getValue());
  var formulas = baseRange.getFormulas();
  var parentTasks = findParentTasks(baseData);
  var indexOfPlannedStart = baseData[1].indexOf('plannedStart');
  var indexOfPlannedFinish = baseData[1].indexOf('plannedFinish');

  if(isChecked){//make all parents' bars
    var newData = makeParentChart(baseData, formulas, indexOfPlannedStart, indexOfPlannedFinish, parentTasks, baseDate, baseLine);
    baseRange.setValues(newData);
  } else {//delete all parents' bars
    //delete date and color in the data
    var indexOfProgress = baseData[1].indexOf('progress');
    for (var i = parentTasks.length-1; 0 <= i ; i--){
      baseData[parentTasks[i]['index']][indexOfPlannedStart] = '';
      baseData[parentTasks[i]['index']][indexOfPlannedFinish] = '';
    };
    schedule.getRange(3, indexOfProgress+1, lastRowOfContents-3+1, 1).setBackground(''); //delete the color in the progress column
    updateChart(baseData, 3, lastRowOfContents, baseLine, baseDate, parentTasks);
    //reflect the exiting formulas
    for(var i = 0, len = formulas.length; i < len; i++){
      for(var j = 0, len2 = formulas[0].length; j < len2; j++){
        if(formulas[i][j] !== ''){
          baseData[i][j] = formulas[i][j];
        };
      };
    };
    baseRange.setValues(baseData);
  };
};

function get_parentChart(){
  Logger.log('get_parentChart start');
  var memo = PropertiesService.getDocumentProperties();
  var isChecked = memo.getProperty('ParentChart');
  isChecked = (isChecked == null || isChecked == '') ? false : isChecked;
  memo.setProperty('ParentChart', isChecked);
  return isChecked;
};


function setTimeDiff(num){
  Logger.log('setTimeDiff start');
  var memo = PropertiesService.getDocumentProperties();
  Logger.log('timeDiff: ' + num);
  memo.setProperty('timeDiff', num);
};


