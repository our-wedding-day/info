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
const SHEET_SUMMARY = 'Підсумок';

const MAX_NAME_LENGTH = 200;
const MAX_CONTACT_LENGTH = 254;
const MAX_LONG_LENGTH = 2000;
const RATE_LIMIT_MAX = 3;
const RATE_LIMIT_SECONDS = 3600;

const HEADERS = [
  'Час', 'Ім\'я', 'Контакт', 'Участь', 'Гостей', 'Приїзд', 'Welcome 11.09',
  'Вінчання', 'Неділя 13.09', 'Логістика та житло', 'Харчування та інше', 'Коментар'
];

const DATA_KEYS = [
  'name', 'contact', 'attend', 'guests', 'arrival', 'welcome', 'church', 'sunday',
  'travelNotes', 'foodNotes', 'comment'
];

const ALLOWED = {
  attend: ['Так, буду', 'Під питанням', 'На жаль, ні'],
  guests: ['1', '2', '3', '4', '5+'],
  welcome: ['Так', 'Ні', 'Можливо'],
  church: ['Так', 'Ні'],
  sunday: ['Так', 'Ні', 'Можливо']
};

const LONG_FIELDS = {
  travelNotes: MAX_LONG_LENGTH,
  foodNotes: MAX_LONG_LENGTH,
  comment: MAX_LONG_LENGTH
};

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ ok: false, error: 'invalid_request' });
    }

    var raw = JSON.parse(e.postData.contents);

    if (raw.website) {
      return jsonResponse({ ok: true, action: 'ignored' });
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

    if (!checkRateLimit(data.contact)) {
      return jsonResponse({ ok: false, error: 'rate_limit' });
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    ensureSheets(ss);

    var row = [new Date()].concat(DATA_KEYS.map(function(key) {
      return data[key] || '';
    }));

    var saveAction = upsertResponse(ss, row);
    sendNotification(data, saveAction === 'updated');
    if (saveAction === 'created') {
      createTasks(data);
    } else {
      appendTaskUpdateNote(ss, data.name);
    }
    sendGuestConfirmation(data, saveAction);
    refreshSummary(ss);

    return jsonResponse({ ok: true, action: saveAction });
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
  var contact = normalizeContact(raw.contact || raw.email);

  if (!name) return { error: 'validation' };
  if (!contact) return { error: 'validation' };

  data.name = name;
  data.contact = contact;

  if (ALLOWED.attend.indexOf(raw.attend) === -1) {
    return { error: 'validation' };
  }
  data.attend = raw.attend;

  var guests = String(raw.guests || '1');
  if (ALLOWED.guests.indexOf(guests) === -1) guests = '1';
  data.guests = guests;

  var optionalKeys = ['welcome', 'church', 'sunday'];
  for (var i = 0; i < optionalKeys.length; i++) {
    var key = optionalKeys[i];
    var val = raw[key] ? String(raw[key]) : '';
    if (val && ALLOWED[key].indexOf(val) === -1) val = '';
    data[key] = val;
  }

  var arrival = trimString(raw.arrival, 10);
  if (arrival && !isValidIsoDate(arrival)) return { error: 'validation' };
  data.arrival = arrival;

  var longKey;
  for (longKey in LONG_FIELDS) {
    data[longKey] = sanitizeCell(trimString(raw[longKey], LONG_FIELDS[longKey]));
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

function isValidPhone(value) {
  var digits = String(value).replace(/[\s\-().]/g, '');
  return /^\+?\d{10,15}$/.test(digits) || /^0\d{9}$/.test(digits);
}

function isValidContact(value) {
  var v = trimString(value, MAX_CONTACT_LENGTH);
  return isValidEmail(v) || isValidPhone(v);
}

function normalizeContact(value) {
  var v = trimString(value, MAX_CONTACT_LENGTH);
  if (!v) return '';
  if (isValidEmail(v)) return v.toLowerCase();
  if (isValidPhone(v)) return formatPhone(v);
  return '';
}

function formatPhone(value) {
  var digits = String(value).replace(/\D/g, '');
  if (digits.length === 10 && digits.charAt(0) === '0') {
    digits = '38' + digits;
  }
  return '+' + digits;
}

function contactMatchKey(value) {
  var v = String(value || '').trim();
  if (!v) return '';
  if (isValidEmail(v)) return v.toLowerCase();
  var digits = v.replace(/\D/g, '');
  if (digits.length === 10 && digits.charAt(0) === '0') {
    digits = '38' + digits;
  }
  return digits;
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

function checkRateLimit(contact) {
  var cache = CacheService.getScriptCache();
  var key = 'rsvp_' + contactMatchKey(contact).replace(/[^a-z0-9@._-]/gi, '').substring(0, 120);
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
    ss.getSheetByName(SHEET_RESPONSES).getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  if (!ss.getSheetByName(SHEET_TASKS)) {
    var t = ss.insertSheet(SHEET_TASKS);
    t.appendRow(['Час', 'Гість', 'Задача', 'Деталі', 'Статус']);
  }
  if (!ss.getSheetByName(SHEET_SUMMARY)) {
    ss.insertSheet(SHEET_SUMMARY);
  }
}

function getContactColumnIndex(headers) {
  for (var c = 0; c < headers.length; c++) {
    var h = String(headers[c] || '').trim().toLowerCase();
    if (h === 'contact' || h === 'email' || h === 'контакт' || h.indexOf('контакт') !== -1) {
      return c;
    }
  }
  return 2;
}

function upsertResponse(ss, row) {
  var sheet = ss.getSheetByName(SHEET_RESPONSES);
  var values = sheet.getDataRange().getValues();
  var contactCol = values.length ? getContactColumnIndex(values[0]) : 2;
  var contactKey = contactMatchKey(row[2]);
  var existingRow = -1;

  if (contactKey) {
    for (var i = values.length - 1; i >= 1; i--) {
      if (contactMatchKey(values[i][contactCol]) === contactKey) {
        existingRow = i + 1;
        break;
      }
    }
  }

  if (existingRow > 0) {
    row[0] = new Date();
    sheet.getRange(existingRow, 1, 1, row.length).setValues([row]);
    return 'updated';
  }

  sheet.appendRow(row);
  return 'created';
}

function appendTaskUpdateNote(ss, name) {
  var sheet = ss.getSheetByName(SHEET_TASKS);
  if (!sheet) return;
  sheet.appendRow([
    new Date(),
    sanitizeCell(name || 'Гість'),
    sanitizeCell('Оновлено RSVP'),
    sanitizeCell('Гість змінив відповідь у формі'),
    'Оновлено'
  ]);
}

function sendGuestConfirmation(data, action) {
  if (!data.contact || !isValidEmail(data.contact)) return;

  var subject = action === 'updated'
    ? 'Вашу відповідь оновлено — Сергій & Вікторія'
    : 'Дякуємо за відповідь — Сергій & Вікторія';

  var lines = [
    'Вітаємо, ' + (data.name || '') + '!',
    '',
    action === 'updated'
      ? 'Ми оновили вашу відповідь на запрошення.'
      : 'Ми отримали вашу відповідь на запрошення.',
    '',
    'Участь: ' + (data.attend || '—'),
  ];

  if (data.attend !== 'На жаль, ні') {
    lines.push('Гостей: ' + (data.guests || '1'));
  }

  lines.push(
    '',
    'Якщо щось зміниться — напишіть організатору Яні: Telegram @dayX_yana або +38 066 779 50 08.',
    '',
    'До зустрічі!',
    'Сергій & Вікторія'
  );

  try {
    MailApp.sendEmail(data.contact, subject, lines.join('\n'));
  } catch (err) {
    Logger.log('Guest confirmation failed: ' + err);
  }
}

function refreshSummary(ss) {
  ensureSheets(ss);
  var sheet = ss.getSheetByName(SHEET_SUMMARY);
  var responses = ss.getSheetByName(SHEET_RESPONSES).getDataRange().getValues();
  var stats = {
    total: 0,
    yes: 0,
    maybe: 0,
    no: 0,
    guests: 0,
    notesCount: 0
  };

  for (var i = 1; i < responses.length; i++) {
    var row = responses[i];
    var attend = String(row[3] || '');
    var guests = String(row[4] || '1');
    stats.total++;
    if (attend === 'Так, буду') stats.yes++;
    else if (attend === 'Під питанням') stats.maybe++;
    else if (attend === 'На жаль, ні') stats.no++;
    if (guests === '5+') stats.guests += 5;
    else stats.guests += parseInt(guests, 10) || 1;
    if (row[9] || row[10]) stats.notesCount++;
  }

  sheet.clear();
  sheet.appendRow(['Показник', 'Значення']);
  sheet.appendRow(['Всього відповідей', stats.total]);
  sheet.appendRow(['Так, будуть', stats.yes]);
  sheet.appendRow(['Під питанням', stats.maybe]);
  sheet.appendRow(['На жаль, ні', stats.no]);
  sheet.appendRow(['Орієнтовно гостей (max)', stats.guests]);
  sheet.appendRow(['Відповідей з нотатками', stats.notesCount]);
  sheet.appendRow(['Оновлено', new Date()]);
}

function sendNotification(data, isUpdate) {
  var notifyEmail = getNotifyEmail();
  if (!notifyEmail) {
    Logger.log('NOTIFY_EMAIL не налаштовано. Запустіть setupRsvpSecrets.');
    return;
  }

  var subject = (isUpdate ? 'Оновлена відповідь' : 'Нова відповідь') + ' на весілля: ' + sanitizeEmailSubjectPart(data.name);
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

  if (data.travelNotes) tasks.push(['Логістика та житло', data.travelNotes]);
  if (data.foodNotes) tasks.push(['Харчування та інше', data.foodNotes]);
  if (data.comment) tasks.push(['Коментар', data.comment]);

  tasks.forEach(function(item) {
    sheet.appendRow([new Date(), sanitizeCell(name), sanitizeCell(item[0]), sanitizeCell(item[1]), 'Нова']);
  });
}
