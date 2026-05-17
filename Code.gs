// ============================================================
// НАЛАШТУВАННЯ
// ============================================================
const DRIVE_FOLDER_ID = '1awpzQqAKZGRa_mVQrFaLtzXzCkjqv8MP';
const SHEET_RESPONSES = 'Відповіді';
const SHEET_ДОВІДНИК  = 'Довідник';
const SHEET_WORKERS   = 'Працівники';
const SHEET_OBJECTS   = 'Об\'єкти';
// ============================================================

/**
 * Структура папок на Google Drive:
 *   Фотозвіт/
 *   └── {рік}/
 *       └── {назва об'єкту}/
 *           └── {населений пункт}/
 *               └── {рахунок}, {Прізвище ІБ}, {вулиця}, {буд}, кв.{кв}/
 *                   └── {ПІБ працівника}_{дата}_N.jpg
 *
 * Структура листа "Довідник" (рядок 1 — заголовки):
 *   A - Особ. рах.
 *   B - Прізвище ІБ
 *   C - Насел. пункт
 *   D - Вулиця
 *   E - Буд/др
 *   F - Кв/др
 *
 * Структура листа "Відповіді" (рядок 1 — заголовки):
 *   A  - Дата/час
 *   B  - ПІБ працівника
 *   C  - Назва об'єкту
 *   D  - Особ. рах.
 *   E  - Прізвище ІБ
 *   F  - Насел. пункт
 *   G  - Вулиця
 *   H  - Буд/др
 *   I  - Кв/др
 *   J  - Примітка
 *   K  - Широта
 *   L  - Довгота
 *   M  - Google Maps
 *   N  - Джерело координат
 *   O  - К-сть фото
 *   P  - Фото (посилання)
 */

// ── Утиліта: знайти або створити підпапку ──
function getOrCreateFolder(parent, name) {
  const safeName = name.replace(/[\/\\:*?"<>|]/g, '_').trim() || '_';
  const existing = parent.getFoldersByName(safeName);
  return existing.hasNext() ? existing.next() : parent.createFolder(safeName);
}

// ── Утиліта: відповідь JsonP або JSON ──
// JsonP використовується з GitHub Pages (обходить CORS)
// JSON використовується з Apps Script (звичайний fetch)
function makeOutput(data, callback) {
  const json = JSON.stringify(data);
  if (callback) {
    // JsonP — повертаємо JavaScript виклик функції
    return ContentService
      .createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  // Звичайний JSON
  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET: роздача форми або довідників ──
function doGet(e) {
  var action   = e.parameter.action;
  var callback = e.parameter.callback; // JsonP callback

  if (action === 'getAbonents') return makeOutput(getAbonentsData(), callback);
  if (action === 'getWorkers')  return makeOutput(getSimpleListData(SHEET_WORKERS), callback);
  if (action === 'getObjects')  return makeOutput(getSimpleListData(SHEET_OBJECTS), callback);

  // Повертає HTML форму
  var url     = ScriptApp.getService().getUrl();
  var html    = HtmlService.createHtmlOutputFromFile('photoform');
  var content = html.getContent().replace('__SCRIPT_URL__', url);
  return HtmlService.createHtmlOutput(content)
    .setTitle('Фото-звіт')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ── Довідник абонентів (дані) ──
function getAbonentsData() {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_ДОВІДНИК);
    const data  = sheet.getDataRange().getValues();

    const abonents = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;
      abonents.push({
        account:  String(row[0]).trim(),
        name:     String(row[1]).trim(),
        locality: String(row[2]).trim(),
        street:   String(row[3]).trim(),
        building: String(row[4]).trim(),
        apt:      String(row[5]).trim()
      });
    }
    return { status: 'ok', data: abonents };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ── Простий довідник (дані) ──
function getSimpleListData(sheetName) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { status: 'ok', data: [] };

    const data = sheet.getDataRange().getValues();

    if (sheetName === SHEET_OBJECTS) {
      const list = [];
      for (let i = 1; i < data.length; i++) {
        const name = String(data[i][0]).trim();
        const type = String(data[i][1]).trim();
        if (name) list.push({ name, type: type || '1' });
      }
      return { status: 'ok', data: list };
    }

    const list = [];
    for (let i = 1; i < data.length; i++) {
      const val = String(data[i][0]).trim();
      if (val) list.push(val);
    }
    return { status: 'ok', data: list };
  } catch (err) {
    return { status: 'error', message: err.message };
  }
}

// ── POST: зберегти дані форми та фото ──
function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const {
      workerName,
      objectName,
      account,
      abonentName,
      locality,
      street,
      building,
      apt,
      note,
      photos,
      coords
    } = payload;

    const now  = new Date();
    const year = Utilities.formatDate(now, 'Europe/Kiev', 'yyyy');
    const date = Utilities.formatDate(now, 'Europe/Kiev', 'dd.MM.yyyy');

    // ── Структура папок ──
    const rootFolder     = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const yearFolder     = getOrCreateFolder(rootFolder, year);
    const objectFolder   = getOrCreateFolder(yearFolder, objectName);
    const localityFolder = getOrCreateFolder(objectFolder, locality || 'Без населеного пункту');

    const parts = [];
    if (account)     parts.push(account);
    if (abonentName) parts.push(abonentName);
    if (street)      parts.push(street);
    if (building)    parts.push(building);
    if (apt)         parts.push('кв.' + apt);
    const abonentFolder = getOrCreateFolder(localityFolder, parts.join(', ') || 'Без адреси');

    // ── Збереження фото ──
    const photoLinks = [];
    if (photos && photos.length > 0) {
      const safeName = (workerName || 'Працівник').replace(/[\/\\:*?"<>|]/g, '_');
      photos.forEach((photo, idx) => {
        const fileName = `${safeName}_${date}_${idx + 1}.jpg`;
        const blob = Utilities.newBlob(
          Utilities.base64Decode(photo.base64),
          photo.mimeType || 'image/jpeg',
          fileName
        );
        const file = abonentFolder.createFile(blob);
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        photoLinks.push(file.getUrl());
      });
    }

    // ── Координати ──
    const lat    = (coords && coords.lat) ? coords.lat : '';
    const lon    = (coords && coords.lon) ? coords.lon : '';
    const source = (coords && coords.source) ? coords.source : 'none';
    const mapsUrl = (lat && lon) ? `https://maps.google.com/?q=${lat},${lon}` : '';
    const sourceLabels = { exif: 'EXIF фото', browser: 'GPS пристрою', none: 'Відсутні' };
    const sourceLabel  = sourceLabels[source] || source;

    // ── Запис у таблицю ──
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_RESPONSES);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_RESPONSES);
      sheet.appendRow([
        'Дата/час', 'ПІБ працівника', 'Назва об\'єкту',
        'Особ. рах.', 'Прізвище ІБ', 'Насел. пункт',
        'Вулиця', 'Буд/др', 'Кв/др', 'Примітка',
        'Широта', 'Довгота', 'Google Maps', 'Джерело координат',
        'К-сть фото', 'Фото (посилання)'
      ]);
      sheet.setFrozenRows(1);
    }

    const datetime = Utilities.formatDate(now, 'Europe/Kiev', 'dd.MM.yyyy HH:mm:ss');
    sheet.appendRow([
      datetime, workerName, objectName,
      account, abonentName, locality,
      street, building, apt, note || '',
      lat, lon, mapsUrl, sourceLabel,
      photoLinks.length, photoLinks.join('\n')
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'ok',
        message: 'Дані збережено',
        photoCount: photoLinks.length,
        coordsSource: sourceLabel
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
