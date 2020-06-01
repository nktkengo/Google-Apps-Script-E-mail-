//version 30
//  getValuesへの変更におけるエラー：string()を用いて　[Ljava.lang.Object;@5a26d44c　（エラー）の解消。
//　id取得時のエラーがあったが、replaceの候補を増やして対応。


//　大量送信向けである為7通以上送信しない場合には、チェック機構が働かない（cellが空かどうか確認する項目）
//  code 側でタイトルなどを設定すればこの問題は解決出来る。（今回の利用者は、codeを読みたがらない為この形式を採用した。
function mainMail(){
  try{
    /*下準備　　spreadsheet、document,driveの指定を行う*/
    //  const sheet_id = '1PEL6YsE7k1MdqOkyu8x8sVeFF8UD_BtEpODxZm2pQF8' //spreadsheetのid
    //  const sheet1 = SpreadsheetApp.openById(sheet_id);  
    const sheet = SpreadsheetApp.getActiveSheet();　　
    const lastRow = sheet.getLastRow();  //最終行の取得
    const document_id = '14VzUthT4xvFm3MVj401hRGkSqhCKLuL8r4m7D2yH7RI'; //dcumentのid
    const doc = DocumentApp.openById(document_id);
    const docText = doc.getBody().getText();
    const subValues = sheet.getRange(1, 8, 8, 1).getValues();
    const arrayData = ['B','C','D','E'];
    sheet.getRange(1, 1, lastRow, 1).clear();
    sheet.getRange(1, 1, lastRow, 1).setFontColor("red");

    //タイトルが空かどうかの確認を行う。
    if(sheet.getRange('H3').isBlank()){
      popupTitle();//popupTitleは関数
    }
    
    //添付ファイルを付けるのか付けないのかの確認
    if(subValues[7][0]){
      dataURL(sheet);//dataURLは関数
    }
    
    /*メール送信前にspreadsheetに抜けがないかの確認。もし、抜けていたら。処理を終了させる。
    但し、行数が８行までの場合は、確認をしない。
    */
    if(lastRow == 8);
    else{
      for(let j = 0; j < 4; j++){
        for(let s = 2; s <= lastRow; s++){
          let contentArr = arrayData[j] + s
          let blank = sheet.getRange(contentArr).isBlank();
          if(blank){
            let qq =　msgBoxque2(contentArr,'がemptyです。\\nこのままでよろしいですか？');//msgBoxNoque2は関数
            if(qq != "yes"){
              msgBoxNoque2(contentArr, 'がemptyです。\\n入力を行ってください。');
              throw new Error('終了します' + contentArr +'のミス');
            }
          }
        }
      }
    }  
  
    const x = msgBoxque1('メールを送信しますがよろしいですか？');
    if(x != "yes"){
      msgBoxNoque1("確認してから実行してください。");//msgBoxNoque1は関数
      throw new Error('終了します。確認するそうだ');
    }
    
    //メール送信
    mail(sheet, docText,subValues);//mailは関数
    
    msgBoxNoque1('送信が完了しました\\n ');
    sheet.getRange(1, 10, lastRow, 2).clear();
    const ww = msgBoxque1('cellのデータをすべて削除してもよろしいですか?');//msgBoxNoque1は関数
    if(ww == true){
      sheet.getRange(2, 1, lastRow, 5).clear();
    }    
  }catch(e){
    console.error("エラー：", e.message);
  }
}


//添付ファイルのURLを取得。そして、名前の中にソートし、cellに出力。
function dataURL(sheet) {
  const message = msgBoxque1('添付ファイルのURLを設定しましたか？');
  if(message == "yes");
  else if(message == "no"){
    const folderName = Browser.inputBox("Google DriveのフォルダURLを入れてください");
    const folderID=folderName.replace('https://drive.google.com/drive/folders/','');
    const folder = DriveApp.getFolderById(folderID);
    const files = folder.getFiles();
    let list =[];
    rowIndex = 1; 
    colIndex = 10;  
    list.push(["ファイル名","URL"]);//リストの追加
    
    while(files.hasNext()) {
      const buff = files.next();
      list.push([buff.getName(), buff.getUrl()]);//リストの追加
    }
    list.sort(function(a,b){return(a[0] - b[0]);});
    
    // 対象の範囲にまとめて書き出します
    sheet.getRange(rowIndex, colIndex, list.length, list[0].length).setValues(list);
    msgBoxNoque1("確認してから実行してください。"); //msgBoxNoque1は関数
    throw new Error('終了します。初見');
  }else{
    throw new Error('終了します。添付ファイルPOP-UPで×を選択');
  }
}

/*メール送信。　forLoopでdocumentの置換を行い、個々のdocument作成。その後送信。*/  
function mail(sheet, docText,subValues){
  const lastRow = sheet.getLastRow();  
  const values = sheet.getRange(2, 2, lastRow, 4).getValues();
  let options;
  let a = 0;
  
  for(let i = 0; i < lastRow - 1; i++){
    const body = docText       //document 置換
    .replace('{送信日}',subValues[0][0])
    .replace('{姓}',values[i][0])  
    .replace('{名}',values[i][1])
    .replace('{点数}',values[i][2]);
    
    
    const attached = sheet.getRange(i + 2, 11).getValue();
    
    if(attached){
      //添付ファイルがemptyかどうかを調べ、値が入力されていれば添付ファイルを送付する。
      const attachfile = replaceDrive(attached)
      options = {name: String(subValues[1]), noReply: String(subValues[5]), cc: String(subValues[3]), bcc: String(subValues[4]), attachments:attachfile};
    }else{
      options = {name: String(subValues[1]), noReply: String(subValues[5]), cc: String(subValues[3]), bcc: String(subValues[4])};
    }
    //差出人、CC、BCC、replayの有無、replay先の変更を設定
    GmailApp.sendEmail(values[i][3], subValues[2][0], body,options);//メール送信 , options
    sheet.getRange(i + 2, 1).setValue('送信済み');
  }
}


//pdf document spreadsheet のIDを取得する
function replaceDrive(attached){
  const attachedID = attached
  .replace('https://drive.google.com/file/d/','')
  .replace('https://docs.google.com/document/d/','')
  .replace('https://docs.google.com/spreadsheets/d/','')
  .replace('/view?usp=drivesdk','')
  .replace('/edit?usp=drivesdk','');
  const attachfile = DriveApp.getFileById(attachedID)
  return attachfile;
}


//メッセージボックス　引数1つ
function msgBoxNoque1(value) {
 Browser.msgBox(value, Browser.Buttons);
}

//メッセージボックス　引数2つ
function msgBoxNoque2(value1, value2){
  Browser.msgBox(value1 + value2, Browser.Buttons);
}

//メッセージボックス質問(yes no)　引数1つ
function msgBoxque1(value) {
  return message =Browser.msgBox(value, Browser.Buttons.YES_NO);
}

//メッセージボックス質問(yes no)　引数2つ
function msgBoxque2(value1, value2) {
  return message =Browser.msgBox(value1 + value2, Browser.Buttons.YES_NO);
}

//ポップアップ警告処理。タイトル
function popupTitle(){
  const rezult = msgBoxque1("タイトルが入力されていません。\\nタイトルなしで送信されますか？");
  if(rezult == "yes"){
    const flag = msgBoxque1("本当によろしいでしょうか？");
    if(flag == "no"){
      msgBoxNoque1("タイトルを入力してから実行してください。");
      throw new Error('終了します。タイトルのミス');
    }else;//実行される。
  }else{
    msgBoxNoque1("タイトルを入力してから実行してください。");
    throw new Error('終了します。タイトルのミス');
  }
}


//
//// optionsに属する配列がemptyかどうかでoptionの返す値を変化させる。
////配列の長さが2未満であれば、入力していないとみなす。（タイトル、ccなどは最低2文字は必要となる事を利用する）
//function optionIsempty(subValues){
//  let options;
//  if(subValues[3][0].length > 1 && subValues[4][0].length > 1 ){
//    return options = {name: subValues[1][0], noReply: subValues[5][0], replyTo: subValues[6][0]};
//  }else if (subValues[3][0].length > 1 || subValues[4][0].length > 1){
//    if(subValues[3][0].length > 1){
//      return options = {name: subValues[1][0], bcc: subValues[4][0], noReply: subValues[5][0]};
//    }else{
//      return options = {name: subValues[1][0], cc: subValues[3][0], noReply: subValues[5][0]};
//    }
//  }else{ 
//    return options = {name: subValues[1][0], noReply: subValues[5][0]};
//  }
//}