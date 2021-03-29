function BorrowBook(bookData, SS, STATUS_SHEET){
  //bookData = {sheetName, bookNumber};

  let answers = GetBorrowData(bookData);
  if (answers == null){
    return;
  }
  Logger.log("本No." + answers.bookNumber + "，"
  　　　　　　 + answers.employeeName + "さん（社員番号" + answers.employeeNumber + "）の貸出"
  　　　　　　 + "（貸出日：" + Utilities.formatDate(answers.borrowDate,"JST", "yyyy/MM/dd")
             + "，返却予定日：" + Utilities.formatDate(answers.backDeadline,"JST", "yyyy/MM/dd") + "）");

  let bookRows = SearchBookRows(bookData, STATUS_SHEET);
  if (bookRows == ""){
    return;
  }
  Logger.log("bookRows:貸出状況シート" + bookRows + "行目");

  InsertBorrowLogData(answers, SS);

  ResisterStatus(answers, bookRows, STATUS_SHEET);

  UpdateFormByBorrow(answers, bookRows, STATUS_SHEET);
}

function GetBorrowData(bookData){

  let error = {};
  error.book = bookData.bookNumber +"-貸出";
  error.where = "GetBorrowData(BorrowManager)";

  const TRIGGER_SS = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = TRIGGER_SS.getSheetByName(bookData.sheetName);

  let lastRow = sheet.getLastRow();
  let range = sheet.getRange(lastRow, 2, 1, sheet.getLastColumn());
  let cells = range.getValues(); //列＝配列番号＋２

  //回答の場所を探す
  let col = 0;
  while (cells[0][col] == ""){
    if (col >= cells[0].length){
      error.key = "answers取得前";
      error.what = "フォームの回答がありません（トリガーシート" + bookData.sheetName + "，"
      　　　　　　　　 + lastRow + "行目のタイムスタンプ）";
      InsertError(error);
      return;
    }
    col++
  }

  let answers = {};
  answers.bookNumber = bookData.bookNumber;
  answers.employeeName = cells[0][col];
  answers.employeeNumber = cells[0][col + 1];
  answers.borrowDate = cells[0][col + 2];
  answers.backDeadline = cells[0][col + 3];

  if (typeof answers.employeeName == "" ||
      typeof answers.employeeNumber != "number" ||
      typeof answers.borrowDate != "object" ||
      typeof answers.backDeadline != "object"){
    error.key = "貸出";
    error.what = "フォームの回答の取得に失敗しました（トリガーシート" + bookData.sheetName + "，"
    　　　　　　　　 + lastRow + "行目のタイムスタンプ，フォームの回答" + (col + 2) + "列目～）";
    InsertError(error, answers);
    return;
  }

  return answers;
}

function InsertBorrowLogData(answers, SS){

  let error = {};
  error.book = answers.bookNumber +"-貸出";
  // error.employeeName = answers.employeeName;
  // error.employeeNumber = answers.employeeNumber;
  // error.formAnswer1 = answers.borrowDate;
  // error.formAnswer2 = answers.backDeadline;
  error.key = "貸出";
  error.where = "InsertBorrowLogData(BorrowManager)";

  let sheet = SS.getSheetByName(answers.bookNumber);
  if (sheet == null || sheet == ""){
    error.what = "貸出履歴シート「" + answers.bookNumber + "」の取得に失敗しました";
    InsertError(error, answers);
    return;
  }

  let cells = sheet.getRange(sheet.getLastRow() + 1, 2, 1, 4);

  let values = [[answers.employeeName, answers.employeeNumber, answers.borrowDate, answers.backDeadline]];
  cells.setValues(values);
}

function ResisterStatus(answers, bookRows, STATUS_SHEET){

  let error = {};
  error.book = answers.bookNumber +"-貸出";
  // error.employeeName = answers.employeeName;
  // error.employeeNumber = answers.employeeNumber;
  // error.formAnswer1 = answers.borrowDate;
  // error.formAnswer2 = answers.backDeadline;
  error.key = "貸出";
  error.where = "ResisterStatus(BorrowManager)";
  
  let lastRow = STATUS_SHEET.getLastRow();
  let borrowersNumbers = STATUS_SHEET.getRange(bookRows[0], 4, bookRows.length, 1).getValues();

  //同じ人によってこの本が借りられているとき
  let tmp = borrowersNumbers.filter(value => value == answers.employeeNumber);
  if (tmp.length > 0){
    error.what = "この本の貸出はもう済んでいます";
    InsertError(error, answers);
    return;
  }

  //貸出状況シートに履歴を書き込む
  let flag = 0;
  for (let i = 0; i < borrowersNumbers.length; i++){
    if (borrowersNumbers[i][0] > 0){
      continue;
    }
    var statusCells = STATUS_SHEET.getRange(bookRows[i], 3, 1, 4);
    flag++;
    break;
  }
  if (flag == 0){
    error.what = "この本はすべて貸し出されており、貸出手続きができません";
    InsertError(error);
    return;
  }

  let values = [[answers.employeeName, answers.employeeNumber, answers.borrowDate, answers.backDeadline]];
  statusCells.setValues(values);
}

function UpdateFormByBorrow(answers, bookRows, STATUS_SHEET){

  let error = {};
  error.book = answers.bookNumber +"-貸出";
  // error.employeeName = answers.employeeName;
  // error.employeeNumber = answers.employeeNumber;
  // error.formAnswer1 = answers.borrowDate;
  // error.formAnswer2 = answers.backDeadline;
  error.key = "貸出";
  error.where = "UpdateFormByBorrow(BorrowManager)";

  //本がすべて借りられていない場合はフォームの書き換えを行わない
  let borrowersNumbers = STATUS_SHEET.getRange(bookRows[0], 4, bookRows.length, 1).getValues();
  borrowersNumbers = borrowersNumbers.filter(value => value > 0)
  if (borrowersNumbers.length < bookRows.length){
    return;
  }
  
  //フォームを取ってくる
  let formId = STATUS_SHEET.getRange(bookRows[0], 7).getValue();
  if (formId == null || formId == ""){
    error.what = "「貸出状況」シートにフォームIDがありません";
    InsertError(error);
    return;
  }
  
  try {
    var form = FormApp.openById(formId);
  }
  catch(e){
    error.what = "「貸出状況」シートのフォームIDが間違っています";
    InsertError(error);
    return;
  }

  //いちばん近い返却予定日を探す
  let backDeadlines = STATUS_SHEET.getRange(bookRows[0], 6, bookRows.length, 1).getValues();
  backDeadlines.sort((a, b) => a - b);
  backDeadlines[0][0] = Utilities.formatDate(backDeadlines[0][0],"JST", "yyyy/MM/dd");

  //フォームの書き換え
  let items = form.getItems();  
  for (let i = 0; i < items.length; i++){
    form.deleteItem(items[i]);
  }
  form.setDescription("貸出中につき現在借りられません。しばらくお待ちください。 \n返却予定日：" + backDeadlines[0][0]);
}
