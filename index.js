var ss = SpreadsheetApp.getActive();

// シートの取得（なければ作成）
var masterTable = ss.getSheetByName('masterTable');
var resource = ss.getSheetByName('resource');
var schedule = ss.getSheetByName('schedule');
var EVM = ss.getSheetByName('EVM');
if(!EVM){EVM = ss.insertSheet('EVM', 3);}
if(!resource){resource = ss.insertSheet('resource', 1);}
if(!schedule){schedule = ss.insertSheet('schedule', 2);}
if(!masterTable){masterTable = ss.insertSheet('masterTable', 0);}

//フォーマット用アセット
var items = [
  ['No.', 'Tasks', 'Start Date', 'Finish Date', 'Workload', 'Progress']
];
var itemsLength = items[0].length;
var gantchartSpan = 7;


function onOpen() {
  var firstRow = masterTable.getRange('1:1');
  var firstRowNum = firstRow.getNumColumns();

  //セルを固定
  masterTable.setFrozenRows(2)

  //目次をセット
  var itemsRange = masterTable.getRange(2, 1, 1, itemsLength);
  itemsRange.setValues(items);

  //列数を追加
  if (firstRowNum < 30) {
    masterTable.insertColumns(itemsLength+1, 200);
  }
  //列幅の調整
  for (var i = itemsLength + 1; firstRowNum >= i; i++) {
    masterTable.setColumnWidth(i, 25);
  };

  //カスタムメニューをUIに追加
  SpreadsheetApp.getUi()
      .createMenu('カスタムメニュー')
      .addItem('サイドバーの表示', 'showSidebar') //表示名、関数名
      .addToUi();

  //サイトバーを自動的に表示
  showSidebar()
}

//サイドバー関連
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('プロジェクトマネジメント')
      .setWidth(300);
  SpreadsheetApp.getUi()
      .showSidebar(html);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}




function formatGantchart(data) {
  Logger.log(data.startDate);
}








