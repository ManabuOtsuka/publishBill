function getNewBill() {
  var book  = SpreadsheetApp.getActiveSpreadsheet();
  var billSheet = book.getSheetByName("管理画面(請求書)");
  var getBillNo = Browser.inputBox('開始行をA列の請求番号で入力してください',Browser.Buttons.OK_CANCEL);
  var getBillNoEnd = Browser.inputBox('終了行をA列の請求番号で入力してください',Browser.Buttons.OK_CANCEL);
  //getBillNoの変数の型をString型から10進数のInt型へ変換しbillNoと再定義
  var billNo = parseInt(getBillNo, 10);
  var billNoEnd = parseInt(getBillNoEnd, 10);
  //何行分の情報を請求書に入れるか計算
  var dataRange = billNoEnd-billNo;
  //請求書に必要な情報の範囲を指定
  var copyRange = billSheet.getRange(billNo+3,1,billNo+dataRange,11);
  
  var serviceDateArray = [];
  var titleArray = [];
  var priceArray = [];
  var amountArray = [];
  var totalPriceArray = [];
  //以下、必要な情報を変数に格納

  //請求先情報の取得
  var billingTo = copyRange.getCell(1,4).getValue();
  //請求日情報の取得
  var billingDate = copyRange.getCell(1,10).getValue();
  //支払期限情報の取得
  var limit = copyRange.getCell(1,11).getValue();
  
  //以下、dataRangeの行数分の情報を取得
  for(var i=0; i<=dataRange; i++){
  //実施日情報を取得し配列へ格納
  var serviceDate = copyRange.getCell(i+1,2).getValue();
  serviceDateArray.push( serviceDate );
  //品目情報を取得し配列へ格納
  var title = copyRange.getCell(i+1,6).getValue();
  titleArray.push( title );
  //単価情報を取得し配列へ格納
  var price = copyRange.getCell(i+1,7).getValue();
  priceArray.push( price );
  //数量情報を取得し配列へ格納
  var amount = copyRange.getCell(i+1,8).getValue();
  amountArray.push( amount );
  //金額情報を取得し配列へ格納
  var totalPrice = copyRange.getCell(i+1,9).getValue();
  totalPriceArray.push( totalPrice );
  }
//新しい請求書シートを作成
  var today = new Date();
  var sheet = SpreadsheetApp.getActive().getSheetByName('請求書フォーマット');
  SpreadsheetApp.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet().setName(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd')+'_'+billingTo);

//新しい請求書シートへの記入
  var newSheet = SpreadsheetApp.getActive().getSheetByName(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd')+'_'+billingTo);
  SpreadsheetApp.setActiveSheet(newSheet);
  
  var billingNo = newSheet.getRange("B3").setValue(billNo);
  var billingToRange = newSheet.getRange("A4").setValue(billingTo);
  var billingDateRange = newSheet.getRange("F3").setValue(billingDate);
  var limitRange = newSheet.getRange("B9").setValue(limit);
  
  var pasteRange = newSheet.getRange(12,1,12,5);
  for(var i=0; i<=dataRange; i++){
  var serviceDateRange = pasteRange.getCell(i+1,1).setValue(serviceDateArray[i]);
  var titleRange = pasteRange.getCell(i+1,2).setValue(titleArray[i]);
  var priceRange = pasteRange.getCell(i+1,3).setValue(priceArray[i]);
  var amountRange = pasteRange.getCell(i+1,4).setValue(amountArray[i]);
  var totalPriceRange = pasteRange.getCell(i+1,5).setValue(totalPriceArray[i]);
  }
}

function addNewBill(){
  var book  = SpreadsheetApp.getActiveSpreadsheet();
  var billSheet = book.getSheetByName("管理画面(請求書)");
  var getBillNo = Browser.inputBox('A列より請求番号を入力してください',Browser.Buttons.OK_CANCEL);
  var billNo = parseInt(getBillNo, 10);
  var getAddNo = Browser.inputBox('請求書の何行目に追記しますか？',Browser.Buttons.OK_CANCEL);
  var AddNo = parseInt(getAddNo, 10);
  var range = billSheet.getRange(billNo+3,1,billNo+3,11);

  //以下、必要な情報を変数に格納
  //実施日
  var serviceDate = range.getCell(1,2).getValue();
  //品目
  var title = range.getCell(1,6).getValue();
  //単価
  var price = range.getCell(1,7).getValue();
  //数量
  var amount = range.getCell(1,8).getValue();
  //金額
  var totalPrice = range.getCell(1,9).getValue();
  
  var editSheet = SpreadsheetApp.getActive().getSheetByName(Browser.inputBox('追記するシート名を入力してください(該当シート名のコピー)',Browser.Buttons.OK_CANCEL));
  SpreadsheetApp.setActiveSheet(editSheet);
  var serviceDateRange = editSheet.getRange("A12").offset(AddNo-1, 0).setValue(serviceDate);
  var titleRange = editSheet.getRange("B12").offset(AddNo-1, 0).setValue(title);
  var priceRange = editSheet.getRange("C12").offset(AddNo-1, 0).setValue(price);
  var amountRange = editSheet.getRange("D12").offset(AddNo-1, 0).setValue(amount);
  var totalPriceRange = editSheet.getRange("E12").offset(AddNo-1, 0).setValue(totalPrice);
  
  
}  
  
