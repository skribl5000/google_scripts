const sheetMovements = SpreadsheetApp.getActive().getSheetByName("Movements");
const sheetLog = SpreadsheetApp.getActive().getSheetByName("log");

let boxes;
let message;
let boxesOptions;
const movesUrl = "https://potapovka-sport.ru/movements";
const headers = { "Content-Type": "Application/json" };

function onOpen(e) {
  SpreadsheetApp.getUi()
    .createMenu("Действия")
    .addItem("Переместить", "movements")
    .addToUi();
}

function movements() {
  createBoxes();
  sendParcels(movesUrl, boxesOptions);

  boxes.forEach((box) => {
    log(box, message);
  });
}

function createBoxes() {
  const lastRow = sheetLog.getLastRow();
  const lastID = sheetLog.getRange(lastRow, 1).getValue();
  const boxId = !isNaN(lastID) ? lastID + 1 : 1;

  let boxNumbers = sheetMovements.getRange(2, 1, 1000).getValues();
  boxNumbers = [...new Set(boxNumbers.join().split(",").filter(Boolean))];

  if (boxNumbers.length > 100) {
    SpreadsheetApp.getUi().alert(
      "Нельзя двигать больше 100 коробов за раз - сервак лопнет :)"
    );
    return;
  }

  boxes = boxNumbers.map((boxNumber) => {
    return {
      id: boxId,
      boxNumber: boxNumber,
      storeName: sheetMovements.getRange(1, 5).getValue(),
      incomeId: sheetMovements.getRange(1, 7).getValue(),
      movementDate: sheetMovements.getRange(1, 9).getValue(),
    };
  });

  boxesOptions = {
    method: "put",
    payload: JSON.stringify({
      store_name: boxes[0].storeName,
      income_id: boxes[0].incomeId,
      movement_date: boxes[0].movementDate,
      box_numbers: boxNumbers,
    }),
    headers: headers,
    muteHttpExceptions: true,
  };
}

function sendParcels(url, options) {
  const response = UrlFetchApp.fetch(url, (options = options));

  if (response.getResponseCode() === 200) {
    SpreadsheetApp.getUi().alert("Всё перемещено!");
    /* sheetMovements.getRange(2, 1, 1000).clearContent();
    sheetMovements.getRange(1, 5).clearContent();
    sheetMovements.getRange(1, 7).clearContent();
    sheetMovements.getRange(1, 9).clearContent(); */
  } else {
    data = response.getContentText();
    try {
      message = JSON.parse(data)["error"];
    } catch {
      message = "Неизвестная ошибка";
    }
    SpreadsheetApp.getUi().alert(message);
  }
}

function log(parcel, comment) {
  const lastRow = sheetLog.getLastRow();
  const nowDate = new Date();
  sheetLog.getRange(lastRow + 1, 1).setValue(parcel.id);
  sheetLog.getRange(lastRow + 1, 2).setValue(nowDate);
  sheetLog.getRange(lastRow + 1, 3).setValue(parcel.boxNumber);
  sheetLog.getRange(lastRow + 1, 4).setValue(parcel.storeName);
  sheetLog.getRange(lastRow + 1, 5).setValue(parcel.movementDate);
  sheetLog.getRange(lastRow + 1, 6).setValue(parcel.incomeId);
  sheetLog.getRange(lastRow + 1, 7).setValue(comment);
}
