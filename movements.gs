const sheetMovements = SpreadsheetApp.getActive().getSheetByName("Movements");
const sheetLog = SpreadsheetApp.getActive().getSheetByName("log");

const boxNumbersRange = sheetMovements.getRange(2, 1, 150);
const boxBarcodesRange = sheetMovements.getRange(2, 2, 150);
const storeNameRange = sheetMovements.getRange(1, 5);
const incomeIdRange = sheetMovements.getRange(1, 7);
const movementDateRange = sheetMovements.getRange(1, 9);

let boxNumbers = boxNumbersRange.getValues();
boxNumbers = [...boxNumbers.join().split(',')].filter(function (e) { return e != ''; });

let boxBarcodes = boxBarcodesRange.getValues();
boxBarcodes = [...boxBarcodes.join().split(',')].filter(function (e) { return e != ''; });

const parcel = {
  boxNumbers: boxNumbers,
  boxBarcodes: boxBarcodes,
  storeName: storeNameRange.getValue(),
  incomeId: incomeIdRange.getValue(),
  movementDate: movementDateRange.getValue(),
  movesUrl: 'https://potapovka-sport.ru/movements'
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
  boxBarcodesRange.clearContent();
}

class movementData {
  constructor(shipment) {
    this._boxNubmers = shipment.boxNumbers;
    this._boxBarcodes = shipment.boxBarcodes;
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

  _validate() {
    const badArray = [
      this._boxNubmers,
      this._boxBarcodes,
      this._storeName,
      this._movementDate,
      this._incomeId
    ]

    this._message = 'Не все поля заполнены!'

    let error = badArray.every((item) => {
      return item !== '' && item > []
    })

    if (this._boxNubmers.length > this._boxBarcodes.length) {
      for (let i = 0; i < this._boxNubmers.length; i++) {
        const value = sheetMovements.getRange(2 + i, 2).getValue();
        if (value !== '') continue
        const boxNumber = sheetMovements.getRange(2 + i, 1).getValue();
        this._message = `Не указан штрих-код у коробки под номером: ${boxNumber}`
        break
      }
      error = false;
    } else if (this._boxNubmers.length < this._boxBarcodes.length) {
      for (let i = 0; i < this._boxBarcodes.length; i++) {
        const value = sheetMovements.getRange(2 + i, 1).getValue();
        if (value !== '') continue
        const barCodeNumber = sheetMovements.getRange(2 + i, 2).getValue();
        this._message = `Не указан номер коробки у штрих-кода под номером: ${barCodeNumber}`
        break
      }
      error = false;
    }

    return error;
  }

  _log(date, message) {
    const lastRowNumber = sheetLog.getLastRow();
    const logId = sheetLog.getRange(lastRowNumber, 1).getValue();
    const shipmentId = !isNaN(logId) ? logId + 1 : 1;

    const boxNubmerArrays = this._boxNubmers.map((boxNubmer) => {
      return [boxNubmer]
    });
    const boxBarcodesArrays = this._boxBarcodes.map((boxBarcode) => {
      return [boxBarcode]
    })

    sheetLog.getRange(lastRowNumber + 1, 1, this._boxNubmers.length).setValue(shipmentId)
    sheetLog.getRange(lastRowNumber + 1, 2, this._boxNubmers.length).setValue(date)
    sheetLog.getRange(lastRowNumber + 1, 3, this._boxNubmers.length).setValues(boxNubmerArrays)
    sheetLog.getRange(lastRowNumber + 1, 4, this._boxBarcodes.length).setValues(boxBarcodesArrays)
    sheetLog.getRange(lastRowNumber + 1, 5, this._boxNubmers.length).setValue(this._storeName)
    sheetLog.getRange(lastRowNumber + 1, 6, this._boxNubmers.length).setValue(this._movementDate)
    sheetLog.getRange(lastRowNumber + 1, 7, this._boxNubmers.length).setValue(this._incomeId)
    sheetLog.getRange(lastRowNumber + 1, 8, this._boxNubmers.length).setValue(message)
  }

  sendParcel() {

    if (this._validate()) {
      if (this._boxNubmers.length > 150) {
        SpreadsheetApp.getUi().alert('Нельзя двигать больше 150 коробов за раз - сервак лопнет :)');
        return;
      }

      const executionDate = new Date();
      const response = UrlFetchApp.fetch(this._movesUrl, this._options);

      if (response.getResponseCode() === 200) {
        clearRanges()
        this._message = 'Всё перемещено!'
      }
      else {
        const data = response.getContentText()
        try {
          this._message = JSON.parse(data)['error'];

        } catch {
          this._message = "Неизвестная ошибка"
        }
      }
      this._log(executionDate, this._message)
      SpreadsheetApp.getUi().alert(this._message);
    } else {
      SpreadsheetApp.getUi().alert(this._message);
    }

  }
}

const movementBoxes = new movementData(parcel)