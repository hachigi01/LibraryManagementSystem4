function ManageLibrary(){

  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.book = "";
  error.employeeName = "";
  error.employeeNumber = "";
  error.formAnswer1 = "";
  error.formAnswer2 = "";
  error.where = "ManageLibrary(FormManager)";

  //定数の宣言
  const SS = ConstSS();
  if (SS == null){
    return;
  }

  const STATUS_SHEET = ConstStatusSheet();
  if (STATUS_SHEET == null){
    return;
  }
   
  //「図書貸出管理トリガー用」SSを取得
  const TRIGGER_SS = SpreadsheetApp.getActiveSpreadsheet();
  const SHEETS = TRIGGER_SS.getSheets();
  let timestamps = [];
  let sortedTimestamps = [];
  let bookData = {};

  //それぞれのシートの一番新しいタイムスタンプを取得
  for (let i = 0; i < SHEETS.length; i++){
    if (SHEETS[i].getLastRow() == 1){
      timestamps[i] = 0;
      sortedTimestamps = 0;
    } else {
      timestamps[i] = SHEETS[i].getRange(SHEETS[i].getLastRow(), 1).getCell(1,1).getValue();
      sortedTimestamps[i] = timestamps[i];
 
      if (SHEETS[i].getRange(SHEETS[i].getLastRow(), 1).getCell(1,1).getValue() == ""){
        error.what = "シート「" + SHEETS[i].getName() +"」の最終行" + SHEETS[i].getLastRow() +"行目に" 
                    +"タイムスタンプがありません";
        InsertError(error);
        return;
      }
  　}
  }

  //すべてのシートの中で一番新しいタイムスタンプの本を探す
  sortedTimestamp.sort(function(a, b) {return b - a;});

  for (let i = 0; i < SHEETS.length; i++){
    if (sortedTimestamp[0].toString() == timestamp[i].toString()){
      bookData.sheetName = SHEETS[i].getName();
    }
  }

  if (bookData.sheetName.indexOf("貸出")　>= 0){
    var sheetNameSplit = bookData.sheetName.split("-");
    bookData.bookNumber = sheetNameSplit[0];
    BorrowBook(bookData, SS, STATUS_SHEET);  //bookData = {sheetName, bookNumber}

  } else if(bookData.sheetName.indexOf("返却")　>= 0){
    BackBook(bookData, SS, STATUS_SHEET);  //bookData = {sheetName}
  }
}



function CreateNewForm() {
  let error = {};
  error.timestamp = new Date(),"JST", "yyyy/MM/dd HH:mm:ss";
  error.employeeName = "";
  error.employeeNumber = "";
  error.formAnswer1 = "";
  error.formAnswer2 = "";

  const SS = ConstSS();
  if (SS == null){
    return;
  }
  
  const STATUS_SHEET = ConstStatusSheet();
  if (STATUS_SHEET == null){
    return;
  }

  let range = STATUS_SHEET.getRange("A:G");
  let lastRow = STATUS_SHEET.getLastRow();

  //貸出状況シートから、一番下の書籍番号を取得
  let bookNumber = range.getCell(lastRow, 1).getValue();

  if (bookNumber == ""){
    let error = {};
    error.book = bookNumber;
    error.where = "CreateNewForm(FormManager)";
    error.what = "書籍番号がありません";
    InsertError(error);
    return;
  }

  let bookTitle = range.getCell(lastRow, 2).getValue();
  if (bookTitle == ""){
    let error = {};
    error.book = bookNumber;
    error.where = "CreateNewForm(FormManager)";
    error.what = "タイトルがありません";
    InsertError(error);
    return;
  }

  //貸出履歴シートの作成
  SS.insertSheet();
  SS.getActiveSheet().setName(bookNumber);
  SS.moveActiveSheet(SS.getNumSheets()); //新しい貸出履歴シートを最後尾に移動

  let logSheet = SS.getActiveSheet();
  // logSheet.getRange(1, 1).getCell(1, 1).setValue("bookTitle");
  // logSheet.getRange(1, 2).getCell(1, 1).setValue("employeeName");
  // logSheet.getRange(1, 3).getCell(1, 1).setValue("employeeNumber");
  // logSheet.getRange(1, 4).getCell(1, 1).setValue("borrowDate");
  // logSheet.getRange(1, 5).getCell(1, 1).setValue("backDeadline");
  // logSheet.getRange(1, 6).getCell(1, 1).setValue("backDate");

  let hairetsu = [["bookTitle", "employeeName", "employeeNumber", "borrowDate", "backDeadline", "backDate"],
                  [bookTitle, "", "", "", "", ""]];
  SHEET.getRange(1, 1, 2, 6).setValues(hairetsu);
  SS.setFrozenRows(1);


  //貸出フォームの作成
  let borrowFormTitle = bookNumber + "-『" + bookTitle + "』の貸出";

  let borrowForm = FormApp.create(borrowFormTitle);
  let borrowFormId = borrowForm.getId();
  let borrowFormFile = DriveApp.getFileById(borrowFormId);

  borrowForm.setDescription("このフォームを送信することによって，個人情報が特定されることはありませんのでご安心ください。");
  borrowForm.addTextItem().setTitle("お名前").setRequired(true);
  const validation = FormApp.createTextValidation().requireNumber().build();//社員番号を数字のみ入力可に
  borrowForm.addTextItem().setTitle("社員番号").setRequired(true).setValidation(validation);
  borrowForm.addDateItem().setTitle('貸出日').setRequired(true);
  borrowForm.addDateItem().setTitle('返却日').setRequired(true);

  //貸出フォームをフォームフォルダへ移動
  try {
    DriveApp.getFolderById(FormFolderId()).addFile(borrowFormFile);
    // Logger.log(DriveApp.getFolderById(FormFolderId()).getName());
    DriveApp.getRootFolder().removeFile(borrowFormFile);
  }
  catch (e) {
    error.book = bookNumber　+"-貸出";
    error.where = "CreateNewForm(FormManager)";
    error.what = "フォームフォルダのIDが間違っています";
    InsertError(error);
    return;
  }

  //貸出フォームIDを「貸出状況」シートに追加
  range.getCell(lastRow, 7).setValue(borrowFormId);

  //貸出フォームとシートを紐づけ
  const TRIGGER_SS = SpreadsheetApp.getActiveSpreadsheet();

  borrowForm.setDestination(FormApp.DestinationType.SPREADSHEET, TRIGGER_SS.getId());

  //紐づけされたシートの名前変更
  var triggerSheets = TRIGGER_SS.getSheets();
  for (let i = 0; i < triggerSheets.length; i++) {
    if (triggerSheets[i].getName() == bookNumber +"-貸出"){
      // Logger.log("in「5-貸出」は既に存在しています")
      error.book = bookNumber +"-貸出";
      error.where = "CreateNewForm(FormManager)";
      error.what = "フォームと紐づけられた「" + bookNumber + "-貸出」シートは既に存在しています。";
      InsertError(error);
      return;
    }
  }

  let flag = 0;
  for (let i = 0; i < triggerSheets.length; i++) {
    
    if (triggerSheets[i].getName().indexOf("フォームの回答") >= 0) {
      if (flag > 0){
        error.book = bookNumber +"-貸出";
        error.where = "CreateNewForm(FormManager)";
        error.what = "（貸出シートを紐づけ）新しいシートが２枚以上あります";
        InsertError(error);
        break;
      }
      triggerSheets[i].setName(bookNumber + "-貸出");
      flag++;
    }
  }
  if (flag == 0){
    error.book = bookNumber +"-貸出";
    error.where = "CreateNewForm(FormManager)";
    error.what = "（貸出シートを紐づけ）新しいシートがありません";
    InsertError(error);
  }

}
