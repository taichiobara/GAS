function myFunction() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('シート1');
  const sheet2 = ss.getSheetByName('シート2');
  const lastRow = sheet.getLastRow();
  
  
  
  for (let i=3;i<=lastRow;i++){
    if(sheet.getRange(i, 3).getValue() === '見出し作製済み' && sheet.getRange(i, 6).getValue() === '')
    { 
       
        var project_name = sheet.getRange(i, 2).getValue();
        var date= sheet.getRange(i, 1).getValue();
        var genre= sheet.getRange(i, 4).getValue();
        var title= sheet.getRange(i, 5).getValue();
        var purpose=sheet.getRange(i, 7).getValue();
        var discription= sheet.getRange(i, 9).getValue();
        var heading1= sheet.getRange(i, 10).getValue();
        var heading2= sheet.getRange(i, 11).getValue();
        var heading3= sheet.getRange(i, 12).getValue();
        var heading4= sheet.getRange(i, 13).getValue();
        var heading5= sheet.getRange(i, 14).getValue();
        
      
        var docName=date+"_"+project_name+"_"+genre+"_"+title;
        var contents="title."+title+"\n\n"+"discription "+discription+"\n\n"+"##"+heading1+"\n\n"+"##"+heading2+"\n\n"+"##"+heading3+"\n\n"+"##"+heading4+"\n\n"+"##"+heading5;
      
        var document = DocumentApp.create(docName);
        document.getBody().setText(contents);
        var docFile = DriveApp.getFileById(document.getId());
        
      
        var destfolder = DriveApp.getFolderById("1FcQiaTOg0hrB069rd76hdmIk7DWseLyv");//ここに大元のファルダのidを入れる
        var destfolder_childs = destfolder.getFolders();//大元のフォルダの中に入っているフォルダの一覧を取得した。
      
        if ( destfolder_childs.hasNext() ){                           //何かしら子フォルダがあるので処理をする

            var flag = 0;                                               //存在チェック用
      
            while ( destfolder_childs.hasNext() ){
              var folder = destfolder_childs.next();
              if ( folder.getName()==genre  ){  //名前とフォルダ名が一致した場合
                flag = 1;                                              //"1"：該当するフォルダが存在した
                
                folder.addFile(docFile);                                 //子フォルダにファイル追加
              }
            }
          if (flag == 0){
              var foldername = genre;
              var newfolder = destfolder.createFolder(foldername);               //任意の名称でフォルダ作成
                                       
              newfolder.addFile(docFile);                                          //ファイルを追加
              
            }
          }else{
              var foldername = genre;
              var newfolder = destfolder.createFolder(foldername);               //任意の名称でフォルダ作成
                                       
              newfolder.addFile(docFile); 
    }
      var documentUrl = document.getUrl();

      
        sheet.getRange(i, 6).setValue(documentUrl);
        sheet.getRange(i, 3).setValue("document作製済み");
        //以下は２つ目のシート
        var lastRow2 = sheet2.getLastRow();
        sheet2.getRange(lastRow2+1, 1).setValue(documentUrl);
        sheet2.getRange(lastRow2+1, 2).setValue("執筆待ち");
        sheet2.getRange(lastRow2+1, 3).setValue(genre);
        sheet2.getRange(lastRow2+1, 4).setValue(title);
        sheet2.getRange(lastRow2+1, 5).setValue(purpose);
        console.log("finish");
      }
  }
  
  
}


function checkstatus(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const mysheet = ss.getSheetByName('シート2');
  const lastRow = mysheet.getLastRow();
  //var myCell = mySheet.getActiveCell(); //アクティブセルを取得
  for (let i=3;i<=lastRow;i++){
    var cell = mysheet.getRange(i, 2).getValue();
    if(cell=="執筆中"&& mysheet.getRange(i, 10).getValue() === ''){
      mysheet.getRange(i, 10).setValue( new Date());
    }else if(cell=="編集待ち"&& mysheet.getRange(i, 11).getValue() === ''){
      mysheet.getRange(i, 11).setValue( new Date());
      
    }else if(cell=="確認待ち"&& mysheet.getRange(i, 12).getValue() === ''){
      mysheet.getRange(i, 12).setValue( new Date());
    }else{}
  }
  
  
}


function onOpen(){
 
  //メニュー配列
  var myMenu=[
    {name: "document製作", functionName: "myFunction"}
    //{name: "配信リスト更新", functionName: "inportContacts2"}
  ];
 
  SpreadsheetApp.getActiveSpreadsheet().addMenu("独自メニュー",myMenu); //メニューを追加
 
}