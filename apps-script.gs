// ============================================================
//  Google Apps Script для приёма RSVP с сайта-приглашения
//  Вставьте этот код в Apps Script (script.google.com)
//  внутри Вашей Google-таблицы (см. инструкцию).
// ============================================================

const SHEET_NAME = 'RSVP';

// Принимает POST { name, answer } с сайта
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const name = String(data.name || '').trim();
    const answer = String(data.answer || '').trim();

    if (!name || (answer !== 'yes' && answer !== 'no')) {
      return jsonResponse({ ok: false, error: 'invalid input' });
    }

    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    const nameLower = name.toLowerCase();

    // Если такое имя уже есть — обновляем строку, иначе добавляем новую
    let rowIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0] || '').toLowerCase().trim() === nameLower) {
        rowIndex = i + 1;
        break;
      }
    }

    const answerRu = answer === 'yes' ? 'Придёт' : 'Не сможет';

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, 3).setValues([[name, answerRu, new Date()]]);
    } else {
      sheet.appendRow([name, answerRu, new Date()]);
    }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

// Отдаёт публичный список гостей (без времени) для отображения на сайте
function doGet(e) {
  try {
    const sheet = getSheet();
    const values = sheet.getDataRange().getValues();
    const guests = [];

    for (let i = 1; i < values.length; i++) {
      const name = String(values[i][0] || '').trim();
      const ans = String(values[i][1] || '').trim();
      if (!name) continue;
      guests.push({
        name: name,
        answer: ans === 'Придёт' ? 'yes' : 'no'
      });
    }

    return jsonResponse({ guests: guests });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err) });
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['Имя', 'Ответ', 'Время']);
    sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return sheet;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
