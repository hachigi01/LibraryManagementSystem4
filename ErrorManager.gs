function InsertError(error, answers) {
  // answers = {employeeName : "山田太郎", employeeNumber : 1,
  //            borrowDate : new Date(), backDeadline: new Date()};
  // error = {"book" : 1, "key" : "貸出", "where" : "function", "what" : "what"};
  Logger.log(error.where + "でエラーが発生しました");

  error = EditErrorContents(error, answers);

  const SS = ConstSS();
  if (SS == null){
    return;
  }

  const ERROR_SHEET = SS.getSheetByName("エラー用");
  let lastRow = ERROR_SHEET.getLastRow();

  let range = ERROR_SHEET.getRange(lastRow + 1, 1, 1, 9);
  range.setValues([["未", Utilities.formatDate(new Date(),"JST", "yyyy/MM/dd HH:mm:ss"), error.book,
                    error.where, error.what,
                    error.employeeName, error.employeeNumber, error.formAnswer1, error.formAnswer2]]);

  MailApp.sendEmail("shiorit.53e@gmail.com", "【図書貸出管理システム】エラーのご報告(step4)" ,
                        "\n図書貸出管理システムにエラーがありました。\nエラー用シート" + lastRow +"行目をご確認ください。"
                        + "\n\n---------------\n"
                        + "エラー内容：" + error.where 
                        + "\n　　　　　　" + error.what);

  return;
}

function EditErrorContents (error, answers){
  switch (error.key){
    case "貸出":
      error.employeeName = answers.employeeName;
      error.employeeNumber = answers.employeeNumber;
      error.formAnswer1 = answers.borrowDate;
      error.formAnswer2 = answers.backDeadline;
      break;

    case "返却":
      error.employeeName = answers.employeeName;
      error.employeeNumber = answers.employeeNumber;
      error.formAnswer1 = answers.backDate;
      error.formAnswer2 = "-";
      break;

    case "answers取得前":
      error.employeeName = "-";
      error.employeeNumber = "-";
      error.formAnswer1 = "-";
      error.formAnswer2 = "-";
      break;

    default:
      error.employeeName = "answersの取得に失敗しました";
      error.employeeNumber = "";
      error.formAnswer1 = "";
      error.formAnswer2 = "";
      break;
  }
  return error;
}

