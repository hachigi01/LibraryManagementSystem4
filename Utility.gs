function SSId() {
  return "1SWkFydsnvyyihLJ2ok8uPtFX3YWXAlgdPPLTD_96rI4";
}

function FormFolderId() {
  return "1wh9rk0FAnzoEN54yCaqzeG8q3afVxEG6";
}

function SearchBookRows(bookData, SS){
  // bookData = {bookNumber : 2}
  // SS = SpreadsheetApp.openById("19yUkB2P7c9IM6yv_FMoLu21VUMaC9AxiktGU5gfmu-c");

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = bookData.bookNumber +"-貸出";
  error.where = "SearchBookRows(Utility)";

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  let range = STATUS_SHEET.getRange("A:A");
  let lastRow = STATUS_SHEET.getLastRow();

  let bookRows = [];
  for (let i = 2; i <= lastRow; i++){
    if (range.getCell(i, 1).getValue() == bookData.bookNumber){
      bookRows.push(i);
    }
  }
  if (bookRows == ""){
    error.what = "「貸出状況」シートから書籍番号が見つかりませんでした";
    InsertError(error);
    return;
  }
  Logger.log(bookRows);
  return bookRows;
}
