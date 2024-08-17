// シート1のIDとシート2のIDが完全一致したら書き込む
function getDataByID() {
  const sheet1 = SpreadsheetApp.getActive().getSheetByName("シート1");
  const sheet2 = SpreadsheetApp.getActive().getSheetByName("シート2");
  const sheet3 = SpreadsheetApp.getActive().getSheetByName("シート3");

  // シート1からタイトル行を取得
  const title = sheet1
    .getRange(1, 1, 1, sheet1.getLastColumn())
    .getValues()[0];

  // 既存データがある場合は削除
  const existingDataRange = sheet3.getRange(2, 1, sheet3.getLastRow(), sheet3.getLastColumn());
  if (existingDataRange.getNumRows() > 0) {
    existingDataRange.clearContent();
  }

  // タイトル行を書き込み
  sheet3
    .getRange(1, 1, 1, title.length)
    .setValues([title]);

  // シート2 検索値タイトルリストから検索対象列を取得
  const searchValues = sheet2
    .getRange(TARGET_RANGE_FOR_TITLE_TO_SEARCH)
    .getValues()
    .flat();
  //console.log("searchValues is", searchValues);

  // シート1（書誌）から検索対象を取得
  let i = 0;
  let matchData;

  console.log("Loop for Write - START");
  while (i < searchValues.length) {
    //console.log("matchData: ID[%s] - Row is[%s]", matchData.getValue(), matchData.getRow());

    // createTextFinder実行時の空文字・nullエラー対策
    if (searchValues[i]) {
      matchData = sheet1
        .getRange(TARGET_RANGE_FOR_DATA_TO_COPY)
        .createTextFinder(searchValues[i].toString())
        .matchCase(true)
        .findNext()
      // 指定行取得
      const rowData = sheet1.getRange(matchData.getRow(), 1, 1, sheet1.getLastColumn()).getValues()[0];

      // データを書き込み
      sheet3
        .getRange(sheet3.getLastRow() + 1, 1, 1, rowData.length)
        .setValues([rowData]);
    }

    i++;
  }
  console.log("Loop for Write - END");
}
