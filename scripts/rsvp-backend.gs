/**
 * RSVP backend для весільного сайту.
 *
 * Після змін у коді: Deploy → Manage deployments → Edit → New version → Deploy
 */
const NOTIFY_EMAIL = 'viktoria.suharevskaya@gmail.com';
const SHEET_RESPONSES = 'Відповіді';
const SHEET_TASKS = 'Задачі';

const HEADERS = [
  'Час', 'Ім\'я', 'Email', 'Участь', 'Гостей', 'Приїзд', 'Welcome 11.09',
  'Вінчання', 'Трансфер', 'Nadiya Palace', 'Заїзд', 'Виїзд', 'Допомога з житлом',
  'Деталі житла', 'Безалкогольне', 'Зачіска/макіяж', 'Неділя 13.09', 'Алергії', 'Харчування',
  'Діти', 'Коментар'
];

const DATA_KEYS = [
  'name', 'email', 'attend', 'guests', 'arrival', 'welcome', 'church', 'transfer',
  'hotel', 'hotelCheckIn', 'hotelCheckOut', 'housingHelp', 'housingHelpDetails', 'nonAlcohol', 'beauty',
  'sunday', 'allergies', 'food', 'children', 'comment'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);

    const row = [new Date()].concat(DATA_KEYS.map(function(key) {
      return data[key] || '';
    }));

    ss.getSheetByName(SHEET_RESPONSES).appendRow(row);
    sendNotification(data);
    createTasks(data);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('RSVP endpoint is active.');
}

function ensureSheets(ss) {
  if (!ss.getSheetByName(SHEET_RESPONSES)) {
    const s = ss.insertSheet(SHEET_RESPONSES);
    s.appendRow(HEADERS);
  } else {
    const s = ss.getSheetByName(SHEET_RESPONSES);
    if (s.getLastColumn() < HEADERS.length) {
      s.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }
  if (!ss.getSheetByName(SHEET_TASKS)) {
    const t = ss.insertSheet(SHEET_TASKS);
    t.appendRow(['Час', 'Гість', 'Задача', 'Деталі', 'Статус']);
  }
}

function sendNotification(data) {
  const subject = 'Нова відповідь на весілля: ' + (data.name || 'без імені');
  const body = HEADERS.slice(1).map(function(h, i) {
    return h + ': ' + (data[DATA_KEYS[i]] || '—');
  }).join('\n');

  MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
}

function createTasks(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_TASKS);
  const name = data.name || 'Гість';
  const tasks = [];

  if (data.transfer === 'Так') tasks.push(['Потрібен трансфер', data.comment || '']);
  if (data.housingHelp === 'Так') tasks.push(['Допомога з житлом', data.housingHelpDetails || data.comment || '']);
  if (data.hotel === 'Так, потрібен номер') {
    var hotelDetails = 'Знижка 50% для гостей весілля';
    if (data.hotelCheckIn) hotelDetails += ', заїзд: ' + data.hotelCheckIn;
    if (data.hotelCheckOut) hotelDetails += ', виїзд: ' + data.hotelCheckOut;
    tasks.push(['Бронювання Nadiya Palace', hotelDetails]);
  }
  if (data.beauty === 'Так') tasks.push(['Зачіска/макіяж', data.comment || '']);
  if (data.nonAlcohol === 'Так') tasks.push(['Безалкогольні напої', data.food || data.comment || '']);
  if (data.allergies) tasks.push(['Алергії', data.allergies]);
  if (data.food) tasks.push(['Особливості харчування', data.food]);
  if (data.children) tasks.push(['Діти', data.children]);

  tasks.forEach(function(item) {
    sheet.appendRow([new Date(), name, item[0], item[1], 'Нова']);
  });
}
