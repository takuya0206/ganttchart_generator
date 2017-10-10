var ss = SpreadsheetApp.getActive();

// シートの取得（なければ作成）
var masterTable = ss.getSheetByName('masterTable');
var resource = ss.getSheetByName('resource');
var schedule = ss.getSheetByName('schedule');
var EVM = ss.getSheetByName('EVM');
var PND = ss.getSheetByName('PND');
if(!masterTable){masterTable = ss.insertSheet('masterTable', 0);}
if(!resource){resource = ss.insertSheet('resource', 1);}
if(!schedule){schedule = ss.insertSheet('schedule', 2);}
if(!EVM){EVM = ss.insertSheet('EVM', 3);}
if(!PND){PND = ss.insertSheet('PND', 4)}

//マスターテーブル情報取得
var firstRow = masterTable.getRange('1:1');
var firstRowNum = firstRow.getNumColumns();


//フォーマット用アセット
var items = [
  ['No.', 'タスク名', '開始日', '終了日', '工数 （人日）', '進捗'],
  ['wbs', 'tasks', 'start', 'finish', 'workload', 'progress']
];

var wbsColumnName = 0;
var wbsCodeName = 1;
var wbsColumnNameLength = items[0].length;


function onOpen() {
  init();
  //WBSの情報がなければガントチャート追加
  if(!masterTable.getRange('B4').getValue()) {
    var today = Moment.moment().format('YYYY-MM-DD');
    formatGantchart(7, today);
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
  masterTable.setFrozenRows(2)
  var wbsColumnRange = masterTable.getRange(2, 1, 2, wbsColumnNameLength);
  wbsColumnRange.setValues(items);
  masterTable.hideRows(3)

  if (firstRowNum < 30) {
    masterTable.insertColumns(wbsColumnNameLength+1, 200);
  }
};

//ガントチャートのフォーマット
function formatGantchart(span, date) {
  var line_column = wbsColumnNameLength+1;
  var date = Moment.moment(date);
  //列幅
  for (var i = wbsColumnNameLength+1; firstRowNum >= i; i++) {
    masterTable.setColumnWidth(i, 25);
  };
  //枠線と日付
  while (firstRowNum >= line_column){
   var range = masterTable.getRange(1, line_column, 1000, line_column+span);
   range.setBorder(null, true, null, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID);
   var range = masterTable.getRange(2, line_column, 1, 1);
   range.setValue('(' + date.format('YYYY/MM/DD') + ')');
   line_column += span;
   date = date.add(span, 'days');
 };
};








