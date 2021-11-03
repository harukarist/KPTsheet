const ss = SpreadsheetApp.getActiveSpreadsheet();
const formSheet = ss.getSheetByName('KPT');
const startCol = 2; //開始列
const numCols = 5; //対象列数
const dateCol = 2; //日付列
let lastRow = formSheet.getLastRow(); // 最終行を取得
let range = formSheet.getRange(2,startCol,lastRow,numCols);

function onOpen() {
  const ui = SpreadsheetApp.getUi(); // UIクラス取得
  const menu = ui.createMenu('GASメニュー'); // スプレッドシートにメニューを追加
  menu.addItem('セルの整形','formatCells'); // 関数をセット
  menu.addToUi(); // スプレッドシートに反映
  formatCells(); //セルの整形
}

// 対象範囲のセルを整形する
function formatCells(){
  range.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP); //文字列を折り返す
  range.setBorder(true,true,true,true,true,true,'#000000',SpreadsheetApp.BorderStyle.DOTTED); //枠線を引く
  formSheet.getRange(2,dateCol,lastRow, 1).setNumberFormat('MM/dd (aaa)');
}