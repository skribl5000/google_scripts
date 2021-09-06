function onOpen(e)
{
  SpreadsheetApp.getUi()
    .createMenu('Действия')
    .addItem('Переместить', 'movements')
    .addToUi();
}

function movements(){
  var sheet = SpreadsheetApp
               .getActive()
               .getSheetByName("Movements")

  var box_numbers_range = sheet
               .getRange(2,1,1000);
  var box_numbers = box_numbers_range.getValues();
  box_numbers = [...new Set(box_numbers.join().split(',').filter(Boolean))];

  if (box_numbers.length > 100){
    SpreadsheetApp.getUi().alert('Нельзя двигать больше 100 коробов за раз - сервак лопнет :)');
    return;
  }

  var store_name_range = sheet
               .getRange(1,5);
  var store_name = store_name_range.getValue()
  var income_id_range = sheet
               .getRange(1,7);
  var income_id = income_id_range.getValue()
  var movement_date_range = sheet
               .getRange(1,9);
  var movement_date = movement_date_range.getValue()

  let moves_url = 'https://potapovka-sport.ru/movements';
  let headers = {"Content-Type": `Application/json`};
  
  var options = {
  'method' : 'put',
  'payload' : JSON.stringify({
    "store_name": store_name,
    "income_id": income_id,
    "movement_date": movement_date,
    "box_numbers": box_numbers
  }),
  'headers':headers,
  'muteHttpExceptions': true,
}
let response = UrlFetchApp.fetch(moves_url, options=options);
if (response.getResponseCode() == 200){
  SpreadsheetApp.getUi().alert('Всё перемещено!');
  box_numbers_range.clearContent();
  income_id_range.clearContent();
  movement_date_range.clearContent();
  store_name_range.clearContent();
}
else{
  data = response.getContentText()
  SpreadsheetApp.getUi().alert(JSON.parse(data)['error']);
}
}
