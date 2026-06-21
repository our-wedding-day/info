/**
 * RSVP backend для весільного сайту.
 *
 * Перший запуск: Run → setupRsvpSecrets (один раз)
 * Після змін у коді: Deploy → Manage deployments → Edit → New version → Deploy
 *
 * RSVP_TOKEN має збігатися з index.html
 */
const INITIAL_RSVP_TOKEN = 'rsvp_S9xK2mP7vQ4nL8wR3jF6';
const SHEET_RESPONSES = 'Відповіді';
const SHEET_TASKS = 'Задачі';

const MAX_NAME_LENGTH = 200;
const MAX_EMAIL_LENGTH = 254;
const MAX_SHORT_LENGTH = 500;
const MAX_LONG_LENGTH = 2000;
const RATE_LIMIT_MAX = 3;
const RATE_LIMIT_SECONDS = 3600;

const HEADERS = [
  'Час', 'Ім\'я', 'Email', 'Участь', 'Гостей', 'Приїзд', 'Welcome 11.09',
  'Вінчання', 'Трансфер', 'Деталі трансферу', 'Nadiya Palace', 'Заїзд', 'Виїзд', 'Допомога з житлом',
  'Деталі житла', 'Безалкогольне', 'Деталі напоїв', 'Зачіска/макіяж', 'Неділя 13.09', 'Алергії', 'Харчування',
  'Діти', 'Коментар'
];

const DATA_KEYS = [
  'name', 'email', 'attend', 'guests', 'arrival', 'welcome', 'church', 'transfer', 'transferDetails',
  'hotel', 'hotelCheckIn', 'hotelCheckOut', 'housingHelp', 'housingHelpDetails', 'nonAlcohol', 'nonAlcoholDetails', 'beauty',
  'sunday', 'allergies', 'food', 'children', 'comment'
];

const ALLOWED = {
  attend: ['Так, буду', 'Під питанням', 'На жаль, ні'],
  guests: ['1', '2', '3', '4', '5+'],
  welcome: ['Так', 'Ні', 'Можливо'],
  church: ['Так', 'Ні'],
  transfer: ['Так', 'Ні'],
  hotel: ['Так, потрібен номер', 'Ні', 'Ще думаю'],
  housingHelp: ['Так', 'Ні'],
  nonAlcohol: ['Так', 'Ні'],
  beauty: ['Так', 'Ні'],
  sunday: ['Так', 'Ні', 'Можливо']
};

const LONG_FIELDS = {
  transferDetails: MAX_LONG_LENGTH,
  housingHelpDetails: MAX_LONG_LENGTH,
  comment: MAX_LONG_LENGTH
};

const SHORT_FIELDS = {
  nonAlcoholDetails: MAX_SHORT_LENGTH,
  allergies: MAX_SHORT_LENGTH,
  food: MAX_SHORT_LENGTH,
  children: MAX_SHORT_LENGTH
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'invalid_request' });
    }

    var raw = JSON.parse(e.postData.contents);

    if (raw.website) {
      return jsonResponse({ ok: true });
    }

    var expectedToken = PropertiesService.getScriptProperties().getProperty('RSVP_TOKEN');
    if (!expectedToken || raw.token !== expectedToken) {
      return jsonResponse({ ok: false, error: 'unauthorized' });
    }

    var validated = validatePayload(raw);
    if (validated.error) {
      return jsonResponse({ ok: false, error: validated.error });
    }

    var data = validated.data;

    if (!checkRateLimit(data.email)) {
      return jsonResponse({ ok: false, error: 'rate_limit' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);

    var row = [new Date()].concat(DATA_KEYS.map(function(key) {
      return data[key] || '';
    }));

    ss.getSheetByName(SHEET_RESPONSES).appendRow(row);
    sendNotification(data);
    createTasks(data);

    return jsonResponse({ ok: true });
  } catch (err) {
    Logger.log(err);
    return jsonResponse({ ok: false, error: 'server_error' });
  }
}

function doGet() {
  return ContentService.createTextOutput('RSVP endpoint is active.');
}

function setupRsvpSecrets() {
  var props = PropertiesService.getScriptProperties();
  props.setProperties({
    NOTIFY_EMAIL: 'viktoria.suharevskaya@gmail.com',
    RSVP_TOKEN: INITIAL_RSVP_TOKEN
  });
  Logger.log('Готово. RSVP_TOKEN: ' + INITIAL_RSVP_TOKEN);
  Logger.log('NOTIFY_EMAIL: viktoria.suharevskaya@gmail.com');
}

function validatePayload(raw) {
  var data = {};
  var name = trimString(raw.name, MAX_NAME_LENGTH);
  var email = trimString(raw.email, MAX_EMAIL_LENGTH).toLowerCase();

  if (!name) return { error: 'validation' };
  if (!email || !isValidEmail(email)) return { error: 'validation' };

  data.name = name;
  data.email = email;

  if (ALLOWED.attend.indexOf(raw.attend) === -1) {
    return { error: 'validation' };
  }
  data.attend = raw.attend;

  var guests = String(raw.guests || '1');
  if (ALLOWED.guests.indexOf(guests) === -1) guests = '1';
  data.guests = guests;

  var optionalKeys = ['welcome', 'church', 'transfer', 'hotel', 'housingHelp', 'nonAlcohol', 'beauty', 'sunday'];
  for (var i = 0; i < optionalKeys.length; i++) {
    var key = optionalKeys[i];
    var val = raw[key] ? String(raw[key]) : '';
    if (val && ALLOWED[key].indexOf(val) === -1) val = '';
    data[key] = val;
  }

  var dateKeys = ['arrival', 'hotelCheckIn', 'hotelCheckOut'];
  for (var j = 0; j < dateKeys.length; j++) {
    var dateKey = dateKeys[j];
    var dateVal = trimString(raw[dateKey], 10);
    if (dateVal && !isValidIsoDate(dateVal)) return { error: 'validation' };
    data[dateKey] = dateVal;
  }

  var longKey;
  for (longKey in LONG_FIELDS) {
    data[longKey] = sanitizeCell(trimString(raw[longKey], LONG_FIELDS[longKey]));
  }

  var shortKey;
  for (shortKey in SHORT_FIELDS) {
    data[shortKey] = sanitizeCell(trimString(raw[shortKey], SHORT_FIELDS[shortKey]));
  }

  return { data: data };
}

function trimString(value, maxLen) {
  if (value == null) return '';
  return String(value).trim().substring(0, maxLen);
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function isValidIsoDate(value) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(value)) return false;
  var parts = value.split('-').map(Number);
  var d = new Date(parts[0], parts[1] - 1, parts[2]);
  return d.getFullYear() === parts[0] && d.getMonth() === parts[1] - 1 && d.getDate() === parts[2];
}

function sanitizeCell(value) {
  if (!value) return '';
  if (/^[=+\-@]/.test(value)) return "'" + value;
  return value;
}

function sanitizeEmailSubjectPart(value) {
  return String(value || 'без імені').replace(/[\r\n]/g, ' ').substring(0, MAX_NAME_LENGTH);
}

function checkRateLimit(email) {
  var cache = CacheService.getScriptCache();
  var key = 'rsvp_' + email.replace(/[^a-z0-9@._-]/gi, '').substring(0, 120);
  var count = parseInt(cache.get(key) || '0', 10);
  if (count >= RATE_LIMIT_MAX) return false;
  cache.put(key, String(count + 1), RATE_LIMIT_SECONDS);
  return true;
}

function getNotifyEmail() {
  return PropertiesService.getScriptProperties().getProperty('NOTIFY_EMAIL');
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets(ss) {
  if (!ss.getSheetByName(SHEET_RESPONSES)) {
    var s = ss.insertSheet(SHEET_RESPONSES);
    s.appendRow(HEADERS);
  } else {
    var responses = ss.getSheetByName(SHEET_RESPONSES);
    if (responses.getLastColumn() < HEADERS.length) {
      responses.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    }
  }
  if (!ss.getSheetByName(SHEET_TASKS)) {
    var t = ss.insertSheet(SHEET_TASKS);
    t.appendRow(['Час', 'Гість', 'Задача', 'Деталі', 'Статус']);
  }
}

function sendNotification(data) {
  var notifyEmail = getNotifyEmail();
  if (!notifyEmail) {
    Logger.log('NOTIFY_EMAIL не налаштовано. Запустіть setupRsvpSecrets.');
    return;
  }

  var subject = 'Нова відповідь на весілля: ' + sanitizeEmailSubjectPart(data.name);
  var body = HEADERS.slice(1).map(function(h, i) {
    return h + ': ' + (data[DATA_KEYS[i]] || '—');
  }).join('\n');

  MailApp.sendEmail(notifyEmail, subject, body);
}

function createTasks(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_TASKS);
  var name = data.name || 'Гість';
  var tasks = [];

  if (data.transfer === 'Так') tasks.push(['Потрібен трансфер', data.transferDetails || data.comment || '']);
  if (data.housingHelp === 'Так') tasks.push(['Допомога з житлом', data.housingHelpDetails || data.comment || '']);
  if (data.hotel === 'Так, потрібен номер') {
    var hotelDetails = 'Знижка 50% для гостей весілля';
    if (data.hotelCheckIn) hotelDetails += ', заїзд: ' + data.hotelCheckIn;
    if (data.hotelCheckOut) hotelDetails += ', виїзд: ' + data.hotelCheckOut;
    tasks.push(['Бронювання Nadiya Palace', hotelDetails]);
  }
  if (data.beauty === 'Так') tasks.push(['Зачіска/макіяж', data.comment || '']);
  if (data.nonAlcohol === 'Так') tasks.push(['Безалкогольні напої', data.nonAlcoholDetails || data.comment || '']);
  if (data.allergies) tasks.push(['Алергії', data.allergies]);
  if (data.food) tasks.push(['Особливості харчування', data.food]);
  if (data.children) tasks.push(['Діти', data.children]);

  tasks.forEach(function(item) {
    sheet.appendRow([new Date(), sanitizeCell(name), sanitizeCell(item[0]), sanitizeCell(item[1]), 'Нова']);
  });
}
