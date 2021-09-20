const sheetMovements = SpreadsheetApp.getActive().getSheetByName("Movements");
const sheetLog = SpreadsheetApp.getActive().getSheetByName("log");

const boxNumbersRange = sheetMovements.getRange(2, 1, 1000);
const storeNameRange = sheetMovements.getRange(1, 5);
const incomeIdRange = sheetMovements.getRange(1, 7);
const movementDateRange = sheetMovements.getRange(1, 9);

let boxNumbers = boxNumbersRange.getValues();
boxNumbers = [...new Set(boxNumbers.join().split(',').filter(Boolean))];

const parcel = {
  boxNumbers: boxNumbers,
  storeName: storeNameRange.getValue(),
  incomeId: incomeIdRange.getValue(),
  movementDate: movementDateRange.getValue(),
  movesUrl: 'https://potapovka-sport.ru/movement'
}

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu('Действия')
    .addItem('Переместить', 'movementBoxes.sendParcel')
    .addToUi();
}

function clearRanges() {
  boxNumbersRange.clearContent();
  storeNameRange.clearContent();
  incomeIdRange.clearContent();
  movementDateRange.clearContent();
}

class movementData {
  constructor(shipment) {
    this._boxNubmers = shipment.boxNumbers;
    this._storeName = shipment.storeName;
    this._movementDate = shipment.movementDate;
    this._incomeId = shipment.incomeId;
    this._movesUrl = shipment.movesUrl;
    this._message = '';
    this._options = {
      'method': 'put',
      'payload': JSON.stringify({
        "store_name": shipment.storeName,
        "income_id": shipment.incomeId,
        "movement_date": shipment.movementDate,
        "box_numbers": shipment.boxNumbers
      }),
      'headers': { 'Content-Type': 'Application/json' },
      'muteHttpExceptions': true,
    }
  }

  _log(date, message) {
    const lastRowNumber = sheetLog.getLastRow();
    const logId = sheetLog.getRange(lastRowNumber, 1).getValue();
    const shipmentId = !isNaN(logId) ? logId + 1 : 1;

    const boxNubmerArrays = this._boxNubmers.map((boxNubmer) => {
      return [boxNubmer]
    });

    sheetLog.getRange(lastRowNumber + 1, 1, this._boxNubmers.length).setValue(shipmentId)
    sheetLog.getRange(lastRowNumber + 1, 2, this._boxNubmers.length).setValue(date)
    sheetLog.getRange(lastRowNumber + 1, 3, this._boxNubmers.length).setValues(boxNubmerArrays)
    sheetLog.getRange(lastRowNumber + 1, 4, this._boxNubmers.length).setValue(this._storeName)
    sheetLog.getRange(lastRowNumber + 1, 5, this._boxNubmers.length).setValue(this._movementDate)
    sheetLog.getRange(lastRowNumber + 1, 6, this._boxNubmers.length).setValue(this._incomeId)
    sheetLog.getRange(lastRowNumber + 1, 7, this._boxNubmers.length).setValue(message)
  }

  sendParcel() {
    if (this._boxNubmers.length > 100) {
      SpreadsheetApp.getUi().alert('Нельзя двигать больше 100 коробов за раз - сервак лопнет :)');
      return;
    }

    const executionDate = new Date();
    const response = UrlFetchApp.fetch(this._movesUrl, this._options);

    if (response.getResponseCode() === 200) {
      SpreadsheetApp.getUi().alert('Всё перемещено!');
      clearRanges()
      this._message = 'OK'
    }
    else {
      const data = response.getContentText()
      try {
        this._message = JSON.parse(data)['error'];
      } catch {
        this._message = "Неизвестная ошибка"
      }
      SpreadsheetApp.getUi().alert(this._message);
    }

    this._log(executionDate, this._message)
  }
}

const movementBoxes = new movementData(parcel)
