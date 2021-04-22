function ConstSS(){
  const SSId = "1SWkFydsnvyyihLJ2ok8uPtFX3YWXAlgdPPLTD_96rI4";

  try {
    const SS = SpreadsheetApp.openById(SSId);
  }
  catch (e) {
    let error = {};
    error.book = "";
    // error.employeeName = "";
    // error.employeeNumber = "";
    // error.formAnswer1 = "";
    // error.formAnswer2 = "";
    error.key = "answers取得前";
    error.where = "ConstSS(Utility)";
    error.what = "スプレッドシート「図書貸出管理」のIDが間違っています";
    InsertError(error);
    return;
  }
  const SS = SpreadsheetApp.openById(SSId);

  return SS;
}

function FormFolderId() {
  return "1wh9rk0FAnzoEN54yCaqzeG8q3afVxEG6";
}

function ConstStatusSheet(){
  const SS = ConstSS();

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    let error = {};
    error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
    error.book = "";
    // error.employeeName = "";
    // error.employeeNumber = "";
    // error.formAnswer1 = "";
    // error.formAnswer2 = "";
    error.key = "answers取得前"
    error.where = "ConstStatusSheet(Utility)";
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }
  return STATUS_SHEET;
}

function SearchBookRows(bookData, STATUS_SHEET, ){

  let error = {};
  error.book = bookData.bookNumber;
  error.where = "SearchBookRows(Utility)";

  let range = STATUS_SHEET.getRange("A:A");
  let lastRow = STATUS_SHEET.getLastRow();
  let bookNumbers = STATUS_SHEET.getRange(2, 1, lastRow - 1/*1行目（見出し行）の分を引く*/, 1).getValues();

  let bookRows = [];
  for (let i = 2; i <= lastRow; i++){
    if (bookNumbers[i] == bookData.bookNumber){
      bookRows.push(i + 2);//行＝i+2
    }
  }
  if (bookRows == ""){
    error.what = "「貸出状況」シートから書籍番号が見つかりませんでした";
    InsertError(error);
    return;
  }
  return bookRows;
}

// function Test (){
//   const SS = ConstSS();
//   const STATUS_SHEET = ConstStatusSheet();
//   let range = STATUS_SHEET.getRange(10, 10, 2, 2);
//   let hairetsu = {"apple" : "red", "banana" : "yellow", "pear" : "green"};
//   Logger.log(hairetsu.apple);

//   range.setValues(hairetsu);
// }