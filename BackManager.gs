function BackBook(bookData, SS, STATUS_SHEET){
  //bookData = {sheetName};
  
  let answers = GetBackData(bookData);
    if (answers == null){
    return;
  }
  Logger.log("本No." + answers.bookNumber + "，"
  　　　　　　 + answers.employeeName + "さん（社員番号" + answers.employeeNumber + "）の返却"
  　　　　　　 + "（返却日：" + answers.backDate + "）");

  bookData.bookNumber = answers.bookNumber;
  let bookRows = SearchBookRows(bookData, STATUS_SHEET);
  if (bookRows == ""){
    return;
  }
  Logger.log("bookRows:貸出状況シート" + bookRows + "行目");

  InsertBackLogData(answers, SS);

  ResetStatus(answers, bookRows,STATUS_SHEET);

  UpdateFormByBack(answers, bookRows, STATUS_SHEET);
}

function GetBackData(bookData){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = bookData.bookNumber +"-返却";
  error.where = "GetBackData(BorrowManager)";

  const TRIGGER_SS = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = TRIGGER_SS.getSheetByName(bookData.sheetName);

  let lastRow = sheet.getLastRow();

  let answerCells = sheet.getRange(lastRow, 1, 1, 4).getValues();

  let answers = {};
  answers.bookNumber = answerCells[0][1];
  answers.employeeName = answerCells[0][2];
  answers.employeeNumber = answerCells[0][3];
  answers.backDate = answerCells[0][0];

  if (typeof answers.bookNumber != "number" ||
      typeof answers.employeeName == "" ||
      typeof answers.employeeNumber != "number" ||
      typeof answers.backDate != "object"){
    error.employeeName = answers.employeeName;
    error.employeeNumber = answers.employeeNumber;
    error.formAnswer1 = answers.borrowDate;
    error.formAnswer2 = answers.backDeadline;
    error.what = "フォームの回答の取得に失敗しました（トリガーシート" + bookData.sheetName + "，"
    　　　　　　　　 + lastRow + "行目）";
    InsertError(error);
    return;
  }

  return answers;
}

function InsertBackLogData(answers, SS){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "InsertBackLogData(BackManager)";

  let sheet = SS.getSheetByName(answers.bookNumber);
  if (sheet == null || sheet == ""){
    error.what = "貸出履歴シート「" + answers.bookNumber + "」の取得に失敗しました";
    InsertError(error);
    return;
  }

  let lastRow = sheet.getLastRow();
  let cells = sheet.getRange(2, 3, lastRow - 1, 4).getValues();
  // Logger.log(cells);
  // Logger.log(cells.length);
  // Logger.log(cells[0].length)

  let borrowersCells = [];
  // Logger.log(cells[2][0]);
  // if (cells[2][0] == answers.employeeNumber){
  //   Logger.log("in bookNumber");
  // }
  // Logger.log(cells[2][3]);
  // if (cells[2][3] == ""){
  //   Logger.log("in backDate");
  // }
  for (let i = 0; i < cells.length; i++){
    if (cells[i][0] == answers.employeeNumber && cells[i][3] == ""){
      borrowersCells.push(i);
    }
  }
  // Logger.log(borrowersCells);

  if (borrowersCells.length == 0){
    error.what = "こちらの社員番号による，返却のない貸出記録が見つかりませんでした";
    InsertError(error);
    return;
  }
  if (borrowersCells.length > 1){
    error.what = "こちらの社員番号による，返却のない貸出記録が２か所以上見つかりました";
    InsertError(error);
    return;
  }

  sheet.getRange(borrowersCells[0] + 2, 6).getCell(1, 1).setValue(answers.backDate);  //行=配列番号＋２
}

function ResetStatus(answers, bookRows, STATUS_SHEET){
  // answers = {bookNumber : 1,
  //            employeeName : "山田太郎",
  //            employeeNumber : 4444,
  //            borrowDate : new Date,
  //            backDeadline : new Date};
  // SS = SpreadsheetApp.openById("19yUkB2P7c9IM6yv_FMoLu21VUMaC9AxiktGU5gfmu-c");

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "ResetStatus(BackManager)";

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();
  
  let flag = 0;
  for (let i = 0; i < bookRows.length; i++){
    if (range.getCell(bookRows[i], 4).getValue() == answers.employeeNumber){
      var statusCells = STATUS_SHEET.getRange(bookRows[i], 3, 1, 4);
      flag++;
      break;
    }
  }
  if(flag == 0){
    error.what = "この社員番号の貸出がありません";
    InsertError(error);
    return;
  }

  statusCells.clear();
}

function UpdateFormByBack(answers, bookRows, STATUS_SHEET) {
  // answers = {bookNumber : 3};
  // SS = SpreadsheetApp.openById("19yUkB2P7c9IM6yv_FMoLu21VUMaC9AxiktGU5gfmu-c");

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = answers.bookNumber　+ "-返却";
  error.employeeName = answers.employeeName;
  error.employeeNumber = answers.employeeNumber;
  error.formAnswer1 = answers.backDate;
  error.formAnswer2 = "-";
  error.where = "UpdateFormByBack(BackManager)";

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  //フォームを取ってくる
  let formId = range.getCell(bookRows[0], 7).getValue();

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
 
  //フォームの書き換え
  let items = form.getItems();
  for (let i = 0; i < items.length; i++){
    form.deleteItem(items[i]);
  }
  form.setDescription("ご記入いただいた情報は図書管理目的のみに使用します。"
                     + "\n借りた人の名前や社員番号が，他の社員の方々に公開されることはございませんのでご安心ください。"
                     + "\n\n貸し出しは一人一冊までです。");
  form.addTextItem().setTitle("お名前").setRequired(true);
  const validation = FormApp.createTextValidation().requireNumber().build();
  form.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation)
    .setHelpText("半角数字４桁でご入力ください");
  form.addDateItem().setTitle('貸出日').setRequired(true).setHelpText("今日の日付をご記入ください");
  form.addDateItem().setTitle('返却日').setRequired(true).setHelpText("２週間後の日付をご記入ください");
}
