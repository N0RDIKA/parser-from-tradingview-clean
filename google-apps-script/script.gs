// Прием сигналов TradingView по вебхуку и запись в Google Таблицу.
// Развертывание: Расширения > Apps Script > Развернуть > Веб-приложение.
//
// Секретный токен хранится в свойствах скрипта, а не в коде:
// Настройки проекта > Свойства скрипта > WEBHOOK_TOKEN.
// В алерте TradingView тот же токен передается в поле "token":
// {"token":"...","ticker":"{{ticker}}","signal":"LE","price":{{close}},"time":"{{time}}"}

var SHEET_NAME = "Лист1";

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = JSON.parse(e.postData.contents);

    // Без проверки токена любой, кто узнает адрес вебхука, сможет прислать сигнал
    var expected = PropertiesService.getScriptProperties().getProperty("WEBHOOK_TOKEN");
    if (!expected || data.token !== expected) {
      return ContentService.createTextOutput("FORBIDDEN");
    }

    if (!data.ticker || !data.signal) {
      return ContentService.createTextOutput("BAD REQUEST");
    }

    // Блокировка, чтобы одновременные алерты не перезаписали одну строку
    lock.waitLock(10000);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    sheet.appendRow([
      new Date(),    // время получения
      data.ticker,   // символ или формула корзины
      data.signal,   // LE, LE_ADD, LX
      data.price,    // цена
      data.time      // время сигнала из TradingView
    ]);
    return ContentService.createTextOutput("OK");

  } catch (err) {
    console.error(err);
    return ContentService.createTextOutput("ERROR");
  } finally {
    lock.releaseLock();
  }
}
