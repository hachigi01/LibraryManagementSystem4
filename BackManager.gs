function BackBook(bookData, SS){
  
  let answers = GetBackData(bookData);
    if (answers == null){
    return;
  }
  Logger.log("answers:" + answers);

  let bookRows = SearchBookRows(answers, SS);
  if (bookRows == ""){
    return;
  }
  Logger.log("bookRows:貸出状況シート" + bookRows + "行目");

  InsertBackLogData(answers, SS);

  ResetStatus(answers, bookRows, SS);

  UpdateFormByBack(answers, bookRows, SS);
}

function GetBackData(bookData){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = bookData.bookNumber +"-返却";
  error.where = "GetBackData(BorrowManager)";

  const TriggerSS = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = TriggerSS.getSheetByName(bookData.sheetName);

  let lastRow = sheet.getLastRow();
  let range = sheet.getRange("A:D");

  let answers = {};
  answers.bookNumber = range.getCell(lastRow, 2).getValue();
  answers.employeeName = range.getCell(lastRow, 3).getValue();
  answers.employeeNumber = range.getCell(lastRow, 4).getValue();
  answers.backDate = range.getCell(lastRow, 1).getValue();

  if (answers.employeeName == null || answers.employeeName == "" ||
      answers.employeeNumber == null || answers.employeeNumber == "" ||
      answers.backDate == null || answers.backDate == ""){
    error.employeeName = answers.employeeName;
    error.employeeNumber = answers.employeeNumber;
    error.formAnswer1 = answers.backDate;
    error.formAnswer2 = "-";
    error.what = "フォームの回答の取得に失敗しました（トリガーシート" + bookData.sheetName + "，"
    　　　　　　　　 + lastRow + "行目のタイムスタンプ）";
    InsertError(error);
    return;
  }
  // Logger.log(answers);
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

  let range = sheet.getRange("B:F");
  let flag = 0;
  for (let row = 2; row <= sheet.getLastRow(); row++){
    if (range.getCell(row, 2).getValue() == answers.employeeNumber && range.getCell(row, 5).isBlank()){
      if (flag > 0){
        error.what = "こちらの社員番号による，返却のない貸出記録が２か所以上見つかりました";
        InsertError(error);
        return;
      }
      range.getCell(row, 5).setValue(answers.backDate);
      flag++;
    }
  }
  if (flag == 0){
    error.what = "こちらの社員番号による，返却のない貸出記録が見つかりませんでした";
    InsertError(error);
    return;
  }

}

function ResetStatus(answers, bookRows, SS){
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

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }
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

function UpdateFormByBack(answers, bookRows, SS) {
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

  const STATUS_SHEET = SS.getSheetByName("貸出状況");
  if (STATUS_SHEET == null || STATUS_SHEET == ""){
    error.what = "スプレッドシート「図書貸出管理」内，「貸出状況」シートの名前が間違っています";
    InsertError(error);
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  if (answers.bookNumber == null || answers.bookNumber == ""){
    error.what = "answersが取得できませんでした";
    InsertError(error);
    return;
  }

  var formId = range.getCell(bookRows[0], 7).getValue();
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
 
  let items = form.getItems();  
  for (let i = 0; i < items.length; i++){
    form.deleteItem(items[i]);
  }
  form.setDescription("ご記入いただいた情報は図書管理目的のみに使用します。"
                     + "\n借りた人の名前や社員番号が，他の社員の方々に公開されることはございませんのでご安心ください。"
                     + "\n\n一人一冊まで借りられます。");
  form.addTextItem().setTitle("お名前").setRequired(true);
  const validation = FormApp.createTextValidation().requireNumber().build();
  form.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation)
    .setHelpText("半角数字４桁でご入力ください");
  form.addDateItem().setTitle('貸出日').setRequired(true).setHelpText("今日の日付をご記入ください");
  form.addDateItem().setTitle('返却日').setRequired(true).setHelpText("２週間後の日付をご記入ください");

}
