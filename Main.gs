/**
 * Функция для получения данных из API и записи их в Google Sheets.
 * 
 * @param {string} API_URL - URL-адрес API .
 * @param {string} API_KEY - Ключ API .
 */
function fetchData() {
  // Укажите ваш API URL и ключ
  const API_URL = 'https://api.site.ru/...'; // Укажите ваш API URL
  const API_KEY = 'YOUR_API_KEY'; // Укажите ваш API ключ

  // Отправляем GET-запрос к API
  const response = UrlFetchApp.fetch(API_URL, {
    method: 'get',
    headers: {
      'Authorization': API_KEY // Используем ключ API для аутентификации
    }
  });

  // Парсим ответ в JSON
  const data = JSON.parse(response.getContentText());

  // Получаем активную таблицу
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  // Очищаем предыдущие данные
  sheet.clear();

  // Укажите заголовки столбцов
  const headers = ['Column1', 'Column2', 'Column3']; // Укажите ваши заголовки
  sheet.appendRow(headers);

  // Добавляем данные в таблицу
  data.forEach(item => {
    // Укажите соответствующие поля для каждого столбца
    const row = [item.field1, item.field2, item.field3]; 
    sheet.appendRow(row);
  });
}

/**
 * Создает триггер для автоматического выполнения функции fetchData.
 */
function createTrigger() {
  ScriptApp.newTrigger('fetchData')
    .timeBased()
    .everyDays(1) // Выполнять каждый день
    .atHour(5) // Время выполнения (5 утра)
    .create();
}

/**
 * Создает пользовательское меню в Google Sheets.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Fetch Data', 'fetchData') // Добавляем пункт меню для вызова функции fetchData
    .addToUi();
}
