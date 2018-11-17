function getNewEstimate() {
  var book  = SpreadsheetApp.getActiveSpreadsheet();
  var billSheet = book.getSheetByName("管理画面(見積書)");
  var getBillNo = Browser.inputBox('開始行をA列の見積番号で入力してください',Browser.Buttons.OK_CANCEL);
  var getBillNoEnd = Browser.inputBox('終了行をA列の見積番号で入力してください',Browser.Buttons.OK_CANCEL);
  var billNo = parseInt(getBillNo, 10);
  var billNoEnd = parseInt(getBillNoEnd, 10);
  var dataRange = billNoEnd-billNo;
  var copyRange = billSheet.getRange(billNo+3,1,billNo+dataRange,12);
  
  var serviceDateArray = [];
  var titleArray = [];
  var priceArray = [];
  var amountArray = [];
  var unitArray = [];
  var totalPriceArray = [];
  
  
  
  //以下、必要な情報を変数に格納
  //請求先
  var billingTo = copyRange.getCell(1,4).getValue();
  //発行日
  var billingDate = copyRange.getCell(1,11).getValue();
  //支払条件
  var limit = copyRange.getCell(1,12).getValue();
  
  //以下、dataRangeの行数分の情報を取得
  for(var i=0; i<=dataRange; i++){
  //実施日
  var serviceDate = copyRange.getCell(i+1,2).getValue();
  serviceDateArray.push( serviceDate );
  //品目
  var title = copyRange.getCell(i+1,6).getValue();
  titleArray.push( title );
  //単価
  var price = copyRange.getCell(i+1,7).getValue();
  priceArray.push( price );
  //数量
  var amount = copyRange.getCell(i+1,8).getValue();
  amountArray.push( amount );
  //単位
  var unit = copyRange.getCell(i+1,9).getValue();
  unitArray.push( unit );
  //金額
  var totalPrice = copyRange.getCell(i+1,10).getValue();
  totalPriceArray.push( totalPrice );
  }
  
//新しい見積書シートを作成
  var today = new Date();
  var sheet = SpreadsheetApp.getActive().getSheetByName('見積書フォーマット');
  SpreadsheetApp.setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().duplicateActiveSheet().setName(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd')+'_'+billingTo);
//新しい見積書シートを作成 

//新しい見積書シートへの記入
  var newSheet = SpreadsheetApp.getActive().getSheetByName(Utilities.formatDate(today, 'Asia/Tokyo', 'yyyyMMdd')+'_'+billingTo);
  SpreadsheetApp.setActiveSheet(newSheet);
  
  var billingNo = newSheet.getRange("B3").setValue(billNo);
  var billingToRange = newSheet.getRange("A4").setValue(billingTo);
  var billingDateRange = newSheet.getRange("G3").setValue(billingDate);
  var limitRange = newSheet.getRange("B9").setValue(limit);
  
  var pasteRange = newSheet.getRange(12,1,12,6);
  for(var i=0; i<=dataRange; i++){
  var serviceDateRange = pasteRange.getCell(i+1,1).setValue(serviceDateArray[i]);
  var titleRange = pasteRange.getCell(i+1,2).setValue(titleArray[i]);
  var priceRange = pasteRange.getCell(i+1,3).setValue(priceArray[i]);
  var amountRange = pasteRange.getCell(i+1,4).setValue(amountArray[i]);
  var unitRange = pasteRange.getCell(i+1,5).setValue(unitArray[i]);
  var totalPriceRange = pasteRange.getCell(i+1,6).setValue(totalPriceArray[i]);
  }
}

function addNewEstimate(){
  var book  = SpreadsheetApp.getActiveSpreadsheet();
  var billSheet = book.getSheetByName("管理画面(見積書)");
  var getBillNo = Browser.inputBox('A列より見積番号を入力してください',Browser.Buttons.OK_CANCEL);
  var billNo = parseInt(getBillNo, 10);
  var getAddNo = Browser.inputBox('見積書の何行目に追記しますか？',Browser.Buttons.OK_CANCEL);
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
  //単位
  var unit = range.getCell(1,9).getValue();
  //金額
  var totalPrice = range.getCell(1,10).getValue();
  
  var editSheet = SpreadsheetApp.getActive().getSheetByName(Browser.inputBox('追記するシート名を入力してください(該当シート名のコピー)',Browser.Buttons.OK_CANCEL));
  SpreadsheetApp.setActiveSheet(editSheet);
  var serviceDateRange = editSheet.getRange("A12").offset(AddNo-1, 0).setValue(serviceDate);
  var titleRange = editSheet.getRange("B12").offset(AddNo-1, 0).setValue(title);
  var priceRange = editSheet.getRange("C12").offset(AddNo-1, 0).setValue(price);
  var amountRange = editSheet.getRange("D12").offset(AddNo-1, 0).setValue(amount);
  var unitRange = editSheet.getRange("E12").offset(AddNo-1, 0).setValue(unit);
  var totalPriceRange = editSheet.getRange("F12").offset(AddNo-1, 0).setValue(totalPrice);
  
  
}  
  
