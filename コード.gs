// スプレッドシートが編集されたときに実行される関数
function onEditInstallable(e) {
  console.time('ダイアログ表示処理タイム');
  var sheet = e.source.getActiveSheet();
  var sheetNames = ['食品', '非食品'];

  if (sheetNames.indexOf(sheet.getName()) === -1) {
    return; // 対象外のシートの場合は処理を終了
  }

  var range = e.range;
  var row = range.getRow();
  var rowValues = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
  var isRowEmptyOrBoolean = rowValues.every(function(cell) { return cell === "" || typeof cell === "boolean"; });
  // 選択された行がすべて空白の場合、新しい行が追加されたと判定
  if (isRowEmptyOrBoolean) {
    var ui = SpreadsheetApp.getUi();
    console.timeEnd('ダイアログ表示処理タイム');
    var response = ui.prompt('追加する商品の商品コードを入力してください。\n※商品コードは、EOSの発注入力ページで、商品名の上部に記載されています。');

    // 入力された内容を現在の編集行の2列目に設定
    if (response.getSelectedButton() == ui.Button.OK) {
      var scode = response.getResponseText();
      Logger.log('scode: ' + scode);  // デバッグ用ログ出力
      editNewRow(sheet, row, scode);
      updateONumbers(sheet);
    }
  }
}

// 新しく追加された行に対して行う処理
function editNewRow(sheet, editedRow, scode) {
  var sheetName = sheet.getName();
  if(sheetName == '食品'){
    sheet.getRange(editedRow, 2).setValue(scode);
  } else {
    sheet.getRange(editedRow, 3).setValue(scode);
  }
  SpreadsheetApp.flush();
  
  var itemData = getShohinInfo.getItemData(scode);

  if (itemData) {
    if(sheetName == '食品'){
      sheet.getRange(editedRow, 1).setValue(itemData.productName); // 商品名
      sheet.getRange(editedRow, 3).setValue(itemData.category); // カテゴリ
      sheet.getRange(editedRow, 4).setValue(itemData.setCount); // セット数
      sheet.getRange(editedRow, 6).setValue(itemData.unitOrder); // 単位(発注書)
      sheet.getRange(editedRow, 8).setFormula('=MIN(I' + editedRow + '-(G' + editedRow + '+E' + editedRow + '), K' + editedRow + ')'); // 発注数関数
      updateONumbers(sheet);//「並び順」
      sheet.setActiveSelection(sheet.getRange(editedRow, 12)); //合わせ数のセル選択
    } else {
      sheet.getRange(editedRow, 2).setValue(itemData.productName); // 商品名
      sheet.getRange(editedRow, 5).setValue(itemData.unitOrder); // 単位(発注書)
      sheet.getRange(editedRow, 7).setFormula('=IF(D' + editedRow + '="", 0, H' + editedRow + '-F' + editedRow + '-D' + editedRow + ')'); // 発注数関数
      sheet.setActiveSelection(sheet.getRange(editedRow, 8)); //合わせ数のセル選択
    }
    
    SpreadsheetApp.flush();
    SpreadsheetApp.getUi().alert('商品を追加しました。合わせ数を入力してください。');
  } else {
    sheet.deleteRow(editedRow); // 商品が見つからなかった場合、行を削除する
    SpreadsheetApp.getUi().alert('商品が見つかりませんでした。');
  }
  // 変更を即座に反映
  SpreadsheetApp.flush();
}

// 「並び順」列に順番に自然数を割り当てる関数
function updateONumbers(sheet) {
  var lastRow = sheet.getLastRow();
  for (var i = 2; i <= lastRow; i++) { // 先頭行を除く
    sheet.getRange(i, 15).setValue(i - 1);
  }
}
