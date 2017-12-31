
function front_updateChart(){
  Logger.log('front_updateChart start')
  var schedule = getScheduleSheet();
  var memo = PropertiesService.getDocumentProperties();
  var baseLine = findStartPoint('progress')+1;
  var baseDate = Moment.moment(memo.getProperty('baseDate'));
  var startRow = 3;
  var endRow = schedule.getLastRow();
  var data = schedule.getRange(1, 1, endRow, baseLine-1).getValues();
  if(startRow > endRow){return;}; //No contents
  updateChart(data, startRow, endRow, baseLine, baseDate);
  drawTodayLine();
};

function front_sumAllWnP(){
  Logger.log('front_sumAllWnP start');
  var schedule = getScheduleSheet();
  var memo = PropertiesService.getDocumentProperties();
  var lastRowOfContents = schedule.getLastRow();
  var baseLine = findStartPoint('progress')+1;
  var baseRange = schedule.getRange(1, 1, lastRowOfContents, baseLine-1);
  var baseData = baseRange.getValues();
  var baseDate = Moment.moment(memo.getProperty('baseDate'));
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
  isChecked = isChecked === null ? false : isChecked;
  return isChecked;
};
