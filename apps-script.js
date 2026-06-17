function normalizeName(raw) {
  return String(raw || '')
    .toLowerCase()
    .replace(/[ё]/g, 'е')
    .replace(/[0]/g, 'о')
    .replace(/[1]/g, 'и')
    .replace(/[3]/g, 'з')
    .replace(/[4]/g, 'ч')
    .replace(/[5]/g, 'с')
    .replace(/[6]/g, 'б')
    .replace(/[7]/g, 'т')
    .replace(/[8]/g, 'в')
    .replace(/[^а-яa-z\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function ensureHeaders_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Дата и время',
      'Событие',
      'ФИО',
      'ФИО нормализованное',
      'Институт',
      'Согласие',
      'Попытка',
      'Очки за попытку',
      'Лучший результат',
      'Финал после 3 попыток',
      'Статус'
    ]);
    sheet.getRange(1, 1, 1, 11)
      .setFontWeight('bold')
      .setBackground('#ff6b35')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
}

function getPlayerStats_(sheet, normalizedName) {
  const values = sheet.getDataRange().getValues();
  let attempts = 0;
  let bestScore = 0;

  for (let i = 1; i < values.length; i++) {
    const rowName = String(values[i][3] || '');
    const event = String(values[i][1] || '');
    const attempt = Number(values[i][6] || 0);
    const best = Number(values[i][8] || 0);

    if (rowName === normalizedName && event === 'attempt') {
      attempts = Math.max(attempts, attempt);
      bestScore = Math.max(bestScore, best);
    }
  }

  return { attempts: attempts, bestScore: bestScore };
}

function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders_(sheet);

  try {
    const data = JSON.parse(e.postData.contents);
    const normalizedName = normalizeName(data.normalizedName || data.fullname);
    const stats = getPlayerStats_(sheet, normalizedName);

    let status = 'ok';
    let attempt = Number(data.attempt || 0);
    let score = Number(data.score || 0);
    let bestScore = Number(data.bestScore || 0);

    if (data.action === 'attempt') {
      if (stats.attempts >= 3) {
        status = 'blocked: 3 attempts already used';
        attempt = stats.attempts;
        bestScore = stats.bestScore;
      } else {
        attempt = Math.min(stats.attempts + 1, 3);
        bestScore = Math.max(stats.bestScore, score, bestScore);
      }
    }

    if (data.action === 'registration' && stats.attempts >= 3) {
      status = 'blocked registration: 3 attempts already used';
    }

    sheet.appendRow([
      data.timestamp || new Date().toLocaleString('ru-RU'),
      data.action || 'registration',
      data.fullname || '',
      normalizedName,
      data.institute || '',
      data.consent ? 'Да' : 'Нет',
      attempt || '',
      score || '',
      bestScore || '',
      attempt >= 3 ? 'Да' : 'Нет',
      status
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: status, attempts: attempt, bestScore: bestScore }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    sheet.appendRow([new Date().toLocaleString('ru-RU'), 'error', '', '', '', '', '', '', '', '', String(err)]);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function testWrite() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  ensureHeaders_(sheet);
  sheet.appendRow([new Date().toLocaleString('ru-RU'), 'test', 'Тест Тестов', 'тест тестов', '1', 'Да', 1, 100, 100, 'Нет', 'ok']);
}
