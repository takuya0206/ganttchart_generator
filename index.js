var ss = SpreadsheetApp.getActive();

// シートの取得（なければ作成）
var masterTable = ss.getSheetByName('masterTable');
var resource = ss.getSheetByName('resource');
var schedule = ss.getSheetByName('schedule');
var EVM = ss.getSheetByName('EVM');
var PDM = ss.getSheetByName('PDM');
if(!masterTable){masterTable = ss.insertSheet('masterTable', 0);}
if(!resource){resource = ss.insertSheet('resource', 1);}
if(!schedule){schedule = ss.insertSheet('schedule', 2);}
if(!EVM){EVM = ss.insertSheet('EVM', 3);}
if(!PDM){PDM = ss.insertSheet('PDM', 4)}

//スケジュールシート情報取得
var firstRow = schedule.getRange('1:1')
var RowNum = firstRow.getNumColumns();


//フォーマット用アセット
var items = [
  ['No.', 'タスク名', '予定開始', '予定終了', '予定工数', '実際開始', '実際終了', '実際工数', '担当', '進捗'],
  ['wbs', 'tasks','planedStart', 'planedFinish', 'planedWorkload', 'actualStart', 'actualFinish', 'actualWorkload', 'responsiblity', 'progress']
];
var wbsColumnNameLength = items[0].length;


function onOpen() {
  //WBSの情報がなければガントチャート追加
  if(!schedule.getRange('B4').getValue()) {
    //月曜スタートになるよう調整
    var today = Moment.moment();
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
  if (RowNum < 30) {
    schedule.insertColumns(wbsColumnNameLength+1, 180);
    RowNum = schedule.getRange('1:1').getNumColumns();
  }
};

//ガントチャートのフォーマット
function formatGantchart(span, date) {
  var line_column = wbsColumnNameLength+1;
  var date = Moment.moment(date);
  //列幅
  for (var i = wbsColumnNameLength+1; RowNum >= i; i++) {
    schedule.setColumnWidth(i, 25);
    if ((i - wbsColumnNameLength) % 7 === 0) {
      schedule.getRange(2, i-1, 1000, 2).setBackground('#fcefe3');
    }
  };
  //枠線と日付
  while (RowNum >= line_column){
   var range = schedule.getRange(1, line_column, 1000, line_column+span);
   range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
   var range = schedule.getRange(1, line_column, 1, 1);
   range.setValue('(' + date.format('YYYY/MM/DD') + ')');
   line_column += span;
   date = date.add(span, 'days');
 };
};








