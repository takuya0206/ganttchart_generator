function front_updateChart(){
  Logger.log('front_updateChart start')
  var memo = PropertiesService.getDocumentProperties();
  var baseLine = findStartPoint('progress')+1;
  var baseDate = Moment.moment(memo.getProperty('baseDate'));
  var startRow = 3;
  var endRow = schedule.getLastRow();
  var data = schedule.getRange(1, 1, endRow, baseLine-1).getValues();
  if(startRow > endRow){return;}; //No contents
  updateChart(data, startRow, endRow, baseLine, baseDate);
};

function front_sumAllWnP(){
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