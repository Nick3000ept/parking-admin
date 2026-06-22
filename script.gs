/**
 * СБ3 Паркинг — backend (Google Apps Script Web App)
 * См. TZ.md.
 *
 * Этот файл — бэкап. Активный код живёт в Apps Script Editor,
 * привязанном к Google Sheets через Extensions → Apps Script.
 */

const CONFIG = {
  SPREADSHEET_ID: '1zB1IQJlla93txI9o-sH-mxD3CeeO0Ugkv8mHPDzcqiI',

  SHEET_GLAVNY: 'Главный',
  SHEET_RABOTY: 'Ведомость_работ',
  SHEET_PODRYADCHIKI: 'Ведомость_подрядчиков',
  SHEET_TIPY: 'Типы_помещений',
  SHEET_LOG: 'Лог_входов',

  // L = 12-я колонка. GAS пишет только в M+ (>= 13). См. §3 TZ.md.
  REGISTRY_LAST_COL: 12,

  // Колонки реестра (1-based)
  COL_KORP: 1,
  COL_FLOOR: 2,
  COL_NUM: 3,
  COL_NAME: 4,
  COL_AREA: 5,
  COL_PERIM: 6,
  COL_WALLS: 7,
  COL_KATEGORIA: 8,
  COL_CHERTEZH: 9,
  COL_ZAHVATKA: 10,
  COL_COMMENT: 11,
  COL_KS: 12,

  TIMEZONE: 'Europe/Moscow',
  LOCK_TIMEOUT_MS: 30000,

  ADMIN_ID: 'admin',
  VIEWER_ID: 'viewer'
};

// =============================================================================
// Entry points
// =============================================================================

function doGet(e) {
  try {
    const params = (e && e.parameter) || {};
    const action = params.action || 'load';

    if (action === 'ping') {
      return jsonResponse(pingDiagnostic());
    }

    const token = params.token || '';
    const user = authenticate(token);
    if (!user) {
      return jsonResponse({ ok: false, error: 'Доступ закрыт' });
    }

    if (action === 'load') {
      logLogin_(user, params.ua || '');
      return jsonResponse(loadSnapshot(user));
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function doPost(e) {
  // Без обёрточного lock'а — read-операции (auth, проверка прав) идут параллельно.
  // Lock берётся ТОЛЬКО на финальный setValue/clearContent внутри setDate/clearDate.
  try {
    const body = JSON.parse(e.postData.contents);
    const user = authenticate(body.token);
    if (!user) {
      return jsonResponse({ ok: false, error: 'Доступ закрыт' });
    }

    if (body.action === 'setDate') {
      return jsonResponse(setDate(user, body.num, body.id_raboty, body.hint));
    }
    if (body.action === 'clearDate') {
      return jsonResponse(clearDate(user, body.num, body.id_raboty, body.hint));
    }
    if (body.action === 'setMarks') {
      return jsonResponse(setMarks(user, body.marks));
    }

    return jsonResponse({ ok: false, error: 'Unknown action: ' + body.action });
  } catch (err) {
    return jsonResponse({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// =============================================================================
// Auth
// =============================================================================

function authenticate(token) {
  if (!token) return null;
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_PODRYADCHIKI);
  if (!sheet) return null;
  const data = sheet.getDataRange().getValues();
  // Header: A:Подрядчик, B:ИД_подрядчик, C:Статус, D:Токен
  for (var i = 1; i < data.length; i++) {
    const rowToken = String(data[i][3] || '');
    const rowStatus = String(data[i][2] || '');
    if (rowToken === String(token) && rowStatus === 'Активен') {
      const id = String(data[i][1] || '').trim();
      return {
        name: String(data[i][0] || '').trim(),
        id: id,
        isAdmin: id === CONFIG.ADMIN_ID,
        isViewer: id === CONFIG.VIEWER_ID
      };
    }
  }
  return null;
}

// =============================================================================
// Load snapshot
// =============================================================================

function loadSnapshot(user) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Работы
  const rabotySheet = ss.getSheetByName(CONFIG.SHEET_RABOTY);
  if (!rabotySheet) throw new Error('Лист "' + CONFIG.SHEET_RABOTY + '" не найден');
  const rabotyData = rabotySheet.getDataRange().getValues();
  const works = [];
  // Header: A:Полное название, B:Поверхность, C:Название для площадки, D:ИД_работы
  for (var i = 1; i < rabotyData.length; i++) {
    if (!rabotyData[i][0]) continue;
    works.push({
      full_name: String(rabotyData[i][0] || '').trim(),
      surface: String(rabotyData[i][1] || '').trim(),
      short_name: String(rabotyData[i][2] || rabotyData[i][0] || '').trim(),
      id_raboty: String(rabotyData[i][3] || '').trim(),
      order: i
    });
  }

  // Типы помещений
  const tipySheet = ss.getSheetByName(CONFIG.SHEET_TIPY);
  if (!tipySheet) throw new Error('Лист "' + CONFIG.SHEET_TIPY + '" не найден');
  const tipyData = tipySheet.getDataRange().getValues();
  const tipyMap = {};
  // Header: A:Тип_помещения_детально, B:Тип_помещения_панель, C:Цвет
  for (var i = 1; i < tipyData.length; i++) {
    if (!tipyData[i][0]) continue;
    tipyMap[String(tipyData[i][0])] = {
      panel: String(tipyData[i][1] || ''),
      color_name: String(tipyData[i][2] || '')
    };
  }

  // Главный
  const glSheet = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
  if (!glSheet) throw new Error('Лист "' + CONFIG.SHEET_GLAVNY + '" не найден');
  const glData = glSheet.getDataRange().getValues();
  if (glData.length < 2) {
    return {
      ok: true,
      user: user,
      rooms: [],
      works: works,
      work_cols: {},
      assignments: {},
      tip_panels: collectUniquePanels(tipyMap),
      server_time: now()
    };
  }
  const headers = glData[0];

  // Маппинг ИД_работы -> { col_date, col_sp } (1-based)
  // Заголовок колонки даты = Работы.Полное название
  // Заголовок колонки СП = Работы.Полное название + " СП"
  const workColMap = {};
  for (var j = CONFIG.REGISTRY_LAST_COL; j < headers.length; j++) {
    const h = String(headers[j] || '').trim();
    for (var w = 0; w < works.length; w++) {
      const fullName = works[w].full_name;
      const idR = works[w].id_raboty;
      if (h === fullName) {
        if (!workColMap[idR]) workColMap[idR] = {};
        workColMap[idR].col_date = j + 1;
      }
      if (h === fullName + ' СП') {
        if (!workColMap[idR]) workColMap[idR] = {};
        workColMap[idR].col_sp = j + 1;
      }
    }
  }

  // Помещения и назначения
  const rooms = [];
  const assignments = {};
  for (var i = 1; i < glData.length; i++) {
    const row = glData[i];
    const num = row[CONFIG.COL_NUM - 1];
    if (!num) continue;

    const numStr = String(num).trim();
    const naimenovanie = String(row[CONFIG.COL_NAME - 1] || '');
    const tipyInfo = tipyMap[naimenovanie] || { panel: '', color_name: '' };

    rooms.push({
      row: i + 1,
      korp: String(row[CONFIG.COL_KORP - 1] || ''),
      floor: String(row[CONFIG.COL_FLOOR - 1] || ''),
      num: numStr,
      name: naimenovanie,
      area: formatNumber(row[CONFIG.COL_AREA - 1]),
      perim: formatNumber(row[CONFIG.COL_PERIM - 1]),
      walls: formatNumber(row[CONFIG.COL_WALLS - 1]),
      kategoria: String(row[CONFIG.COL_KATEGORIA - 1] || ''),
      chertezh: String(row[CONFIG.COL_CHERTEZH - 1] || ''),
      zahvatka: String(row[CONFIG.COL_ZAHVATKA - 1] || ''),
      ks: String(row[CONFIG.COL_KS - 1] || ''),
      tip_panel: tipyInfo.panel,
      color_name: tipyInfo.color_name
    });

    const cellAssignments = {};
    for (var wid in workColMap) {
      if (!workColMap.hasOwnProperty(wid)) continue;
      const cols = workColMap[wid];
      if (!cols.col_date || !cols.col_sp) continue;

      const dateVal = row[cols.col_date - 1];
      const spVal = row[cols.col_sp - 1];
      const sp = spVal ? String(spVal).trim() : '';

      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      } else if (dateVal !== '' && dateVal !== null && dateVal !== undefined) {
        dateStr = String(dateVal);
      }

      if (sp || dateStr) {
        cellAssignments[wid] = { sp: sp, date: dateStr };
      }
    }
    assignments[numStr] = cellAssignments;
  }

  return {
    ok: true,
    user: user,
    rooms: rooms,
    works: works,
    work_cols: workColMap,
    assignments: assignments,
    tip_panels: collectUniquePanels(tipyMap),
    server_time: now(),
    version_hash: computeVersionHash(headers, works.length, Object.keys(tipyMap).length)
  };
}

function collectUniquePanels(tipyMap) {
  const set = {};
  for (var key in tipyMap) {
    if (tipyMap.hasOwnProperty(key)) {
      const p = tipyMap[key].panel;
      if (p) set[p] = true;
    }
  }
  return Object.keys(set);
}

function computeVersionHash(headers, worksCount, tipyCount) {
  // Простой хеш для детекта изменений структуры. Не криптографический.
  const s = headers.join('|') + '#works=' + worksCount + '#tipy=' + tipyCount;
  var hash = 0;
  for (var i = 0; i < s.length; i++) {
    hash = ((hash << 5) - hash) + s.charCodeAt(i);
    hash |= 0;
  }
  return String(hash);
}

// =============================================================================
// Write actions
// =============================================================================

function setDate(user, num, idRaboty, hint) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };

  const cell = resolveCell_(num, idRaboty, hint);
  if (cell.error) return { ok: false, error: cell.error };

  // Проверка прав
  if (!user.isAdmin) {
    if (cell.spValue !== user.name) {
      return { ok: false, error: 'Эта работа не назначена вам' };
    }
  } else {
    if (!cell.spValue) {
      return { ok: false, error: 'Работа не назначена никому, нечего отмечать' };
    }
  }

  // Идемпотентность
  if (cell.currentDate) {
    return { ok: true, date: cell.currentDate, note: 'already_set' };
  }

  assertWritableColumn(cell.colDate);
  const today = new Date();
  withLock_(function () {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_GLAVNY);
    sheet.getRange(cell.row, cell.colDate).setValue(today);
  });

  return {
    ok: true,
    date: Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    num: num,
    id_raboty: idRaboty
  };
}

function clearDate(user, num, idRaboty, hint) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };

  const cell = resolveCell_(num, idRaboty, hint);
  if (cell.error) return { ok: false, error: cell.error };

  if (!user.isAdmin) {
    if (cell.spValue !== user.name) {
      return { ok: false, error: 'Эта работа не назначена вам' };
    }
  }

  if (!cell.currentDate) {
    return { ok: true, date: '', note: 'already_empty', num: num, id_raboty: idRaboty };
  }

  assertWritableColumn(cell.colDate);
  withLock_(function () {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_GLAVNY);
    sheet.getRange(cell.row, cell.colDate).clearContent();
  });

  return { ok: true, date: '', num: num, id_raboty: idRaboty };
}

/**
 * Батчевая запись/очистка отметок. Один HTTP, один Lock, один открытый sheet.
 * Снимает конкуренцию за LockService при множественных кликах в режиме отметки.
 *
 * marks: [{ action: 'set'|'clear', num, id_raboty, hint: {row, colDate, colSp} }]
 * Возвращает: { ok: true, results: [{ ok, num, id_raboty, date?, error?, note? }] }
 */
function setMarks(user, marks) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };
  if (!Array.isArray(marks) || marks.length === 0) {
    return { ok: false, error: 'Пустой батч' };
  }

  const results = new Array(marks.length);
  withLock_(function () {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_GLAVNY);
    if (!sheet) {
      for (var k = 0; k < marks.length; k++) {
        results[k] = { ok: false, num: marks[k].num, id_raboty: marks[k].id_raboty, error: 'Лист "' + CONFIG.SHEET_GLAVNY + '" не найден' };
      }
      return;
    }
    const today = new Date();
    const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd');

    for (var i = 0; i < marks.length; i++) {
      const m = marks[i] || {};
      const num = m.num;
      const idRaboty = m.id_raboty;
      const action = m.action;
      try {
        const cell = resolveCellInSheet_(sheet, num, idRaboty, m.hint);
        if (cell.error) { results[i] = { ok: false, num: num, id_raboty: idRaboty, error: cell.error }; continue; }

        if (!user.isAdmin) {
          if (cell.spValue !== user.name) {
            results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Эта работа не назначена вам' };
            continue;
          }
        } else {
          if (action === 'set' && !cell.spValue) {
            results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Работа не назначена никому, нечего отмечать' };
            continue;
          }
        }

        assertWritableColumn(cell.colDate);

        if (action === 'set') {
          if (cell.currentDate) {
            results[i] = { ok: true, num: num, id_raboty: idRaboty, date: cell.currentDate, note: 'already_set' };
            continue;
          }
          sheet.getRange(cell.row, cell.colDate).setValue(today);
          results[i] = { ok: true, num: num, id_raboty: idRaboty, date: todayStr };
        } else if (action === 'clear') {
          if (!cell.currentDate) {
            results[i] = { ok: true, num: num, id_raboty: idRaboty, date: '', note: 'already_empty' };
            continue;
          }
          sheet.getRange(cell.row, cell.colDate).clearContent();
          results[i] = { ok: true, num: num, id_raboty: idRaboty, date: '' };
        } else {
          results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Unknown action: ' + action };
        }
      } catch (err) {
        results[i] = { ok: false, num: num, id_raboty: idRaboty, error: String(err && err.message ? err.message : err) };
      }
    }
  });

  return { ok: true, results: results };
}

/**
 * Берёт script lock только на короткое время выполнения write-операции.
 * Защищает от одновременной записи в одну и ту же ячейку.
 */
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(CONFIG.LOCK_TIMEOUT_MS);
  try {
    fn();
    SpreadsheetApp.flush();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/**
 * Fast path: если фронт передал hint (row + colDate + colSp), читаем только
 * нужный диапазон. Защита — проверяем что в hint.row реально стоит ожидаемый
 * Номер; если нет — fallback на полный locateCell (Sheets могли поменять).
 */
function resolveCell_(num, idRaboty, hint) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_GLAVNY);
  if (!sheet) return { error: 'Лист "' + CONFIG.SHEET_GLAVNY + '" не найден' };
  return resolveCellInSheet_(sheet, num, idRaboty, hint);
}

/** То же что resolveCell_, но работает с уже открытым sheet — для батча. */
function resolveCellInSheet_(sheet, num, idRaboty, hint) {
  if (hint && hint.row && hint.colDate && hint.colSp &&
      hint.colDate > CONFIG.REGISTRY_LAST_COL && hint.colSp > CONFIG.REGISTRY_LAST_COL) {
    const fromCol = Math.min(CONFIG.COL_NUM, hint.colDate, hint.colSp);
    const toCol = Math.max(CONFIG.COL_NUM, hint.colDate, hint.colSp);
    const values = sheet.getRange(hint.row, fromCol, 1, toCol - fromCol + 1).getValues()[0];
    const actualNum = String(values[CONFIG.COL_NUM - fromCol] || '').trim();

    if (actualNum === String(num).trim()) {
      const dateVal = values[hint.colDate - fromCol];
      const spVal = values[hint.colSp - fromCol];
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'yyyy-MM-dd');
      } else if (dateVal !== '' && dateVal !== null && dateVal !== undefined) {
        dateStr = String(dateVal);
      }
      return {
        row: hint.row,
        colDate: hint.colDate,
        colSp: hint.colSp,
        spValue: spVal ? String(spVal).trim() : '',
        currentDate: dateStr
      };
    }
    // Hint протух — fallback
  }
  return locateCell(num, idRaboty);
}

function locateCell(num, idRaboty) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
  if (!sheet) return { error: 'Лист "' + CONFIG.SHEET_GLAVNY + '" не найден' };

  // Найти работу
  const rabotySheet = ss.getSheetByName(CONFIG.SHEET_RABOTY);
  if (!rabotySheet) return { error: 'Лист "' + CONFIG.SHEET_RABOTY + '" не найден' };
  const rabotyData = rabotySheet.getDataRange().getValues();
  var workName = null;
  for (var i = 1; i < rabotyData.length; i++) {
    if (String(rabotyData[i][3] || '').trim() === String(idRaboty).trim()) {
      workName = String(rabotyData[i][0] || '').trim();
      break;
    }
  }
  if (!workName) return { error: 'Работа не найдена: ' + idRaboty };

  // Найти колонки
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colDate = null;
  var colSp = null;
  for (var j = 0; j < headers.length; j++) {
    const h = String(headers[j] || '').trim();
    if (h === workName) colDate = j + 1;
    if (h === workName + ' СП') colSp = j + 1;
  }
  if (!colDate || !colSp) {
    return { error: 'Колонки работы "' + workName + '" не найдены в листе Главный' };
  }

  // Найти строку
  const lastRow = sheet.getLastRow();
  const numValues = sheet.getRange(2, CONFIG.COL_NUM, lastRow - 1, 1).getValues();
  var row = null;
  for (var k = 0; k < numValues.length; k++) {
    if (String(numValues[k][0] || '').trim() === String(num).trim()) {
      row = k + 2;
      break;
    }
  }
  if (!row) return { error: 'Помещение не найдено: ' + num };

  // Текущие значения
  const dateVal = sheet.getRange(row, colDate).getValue();
  const spVal = sheet.getRange(row, colSp).getValue();
  var dateStr = '';
  if (dateVal instanceof Date) {
    dateStr = Utilities.formatDate(dateVal, CONFIG.TIMEZONE, 'yyyy-MM-dd');
  } else if (dateVal !== '' && dateVal !== null && dateVal !== undefined) {
    dateStr = String(dateVal);
  }

  return {
    row: row,
    colDate: colDate,
    colSp: colSp,
    spValue: spVal ? String(spVal).trim() : '',
    currentDate: dateStr
  };
}

// =============================================================================
// Diagnostics
// =============================================================================

function pingDiagnostic() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const allSheets = ss.getSheets().map(function (s) { return s.getName(); });
    const expected = [CONFIG.SHEET_GLAVNY, CONFIG.SHEET_RABOTY, CONFIG.SHEET_PODRYADCHIKI, CONFIG.SHEET_TIPY];
    const sheetsFound = {};
    expected.forEach(function (name) {
      sheetsFound[name] = allSheets.indexOf(name) !== -1;
    });
    const allOk = expected.every(function (n) { return sheetsFound[n]; });

    var glavnyInfo = null;
    if (sheetsFound[CONFIG.SHEET_GLAVNY]) {
      const g = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
      glavnyInfo = { rows: g.getLastRow(), cols: g.getLastColumn() };
    }

    return {
      ok: allOk,
      version: 'full',
      sheets_found: sheetsFound,
      all_sheets_in_file: allSheets,
      glavny_info: glavnyInfo,
      server_time: now()
    };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

// =============================================================================
// Setup (запускается вручную из Apps Script Editor один раз)
// =============================================================================

/**
 * Создаёт листы Работы, Подрядчики, Типы помещений если их нет,
 * наполняет стартовыми данными. Также добавляет недостающие парные колонки
 * работ в лист Главный. Безопасно: ничего не перезаписывает.
 *
 * Запуск: в Apps Script Editor выбрать setupAll в dropdown функций, нажать Run.
 * Результат смотреть в View → Logs (или Ctrl+Enter).
 */
function setupAll() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const report = [];

  // 1. Работы
  report.push(setupSheet_(ss, CONFIG.SHEET_RABOTY,
    ['Полное название', 'Поверхность', 'Название для площадки', 'ИД_работы'],
    [
      ['Гидроизоляция полов', 'Полы', 'Гидроизоляция', 'down_hydro'],
      ['Заливка стяжки', 'Полы', 'Стяжка', 'down_st'],
      ['Укладка плитки', 'Полы', 'Плитка', 'down_kerama'],
      ['Покраска потолка', 'Потолки', 'Покраска потолок', 'up_color'],
      ['Штукатурка стен', 'Стены', 'Штукатурка', 'wall_stukat'],
      ['Покраска стен', 'Стены', 'Покраска стены', 'wall_paint'],
      ['Монтаж двери', 'Двери', 'Двери', 'doors'],
      ['Заливка топпинга', 'Полы', 'Топпинг', 'down_topping']
    ]
  ));

  // 2. Подрядчики
  report.push(setupSheet_(ss, CONFIG.SHEET_PODRYADCHIKI,
    ['Подрядчик', 'ИД_подрядчик', 'Статус', 'Токен'],
    [
      ['Овчинников',   'ovchi',  'Активен', '81bc3299c7ceb6fb'],
      ['Спектр',       'spectr', 'Активен', 'ce3d0facb5187091'],
      ['Альфа Строй',  'alfa',   'Активен', '8ca318c5521b5b66'],
      ['Админ',        'admin',  'Активен', genToken_()]
    ]
  ));

  // 3. Типы помещений
  report.push(setupSheet_(ss, CONFIG.SHEET_TIPY,
    ['Тип_помещения_детально', 'Тип_помещения_панель', 'Цвет'],
    [
      ['Кладовая', 'Кладовая', 'Желтый'],
      ['Тамбур-шлюз (ЛХ/ПБЗ)', 'ТШ', ''],
      ['Лестничная клетка', 'ЛК', ''],
      ['Архив', 'Прочие помещения', ''],
      ['Душевая охраны', 'Прочие помещения', ''],
      ['Комната отдыха (охрана)', 'Прочие помещения', ''],
      ['Раздевалка мужская', 'Прочие помещения', ''],
      ['Комната приема пищи', 'Прочие помещения', ''],
      ['Душевая мужская', 'Прочие помещения', ''],
      ['Душевая женская', 'Прочие помещения', ''],
      ['Раздевалка женская', 'Прочие помещения', ''],
      ['ПУИ', 'Прочие помещения', ''],
      ['Комната отдыха (техники)', 'Прочие помещения', ''],
      ['Склад материалов и оборудования', 'Прочие помещения', ''],
      ['Мастерская', 'Прочие помещения', ''],
      ['Помещение КНС', 'Тех помещения', 'Серый'],
      ['Проход', 'Прочие помещения', ''],
      ['Коридор', 'Прочие помещения', ''],
      ['Тамбур-шлюз', 'Прочие помещения', ''],
      ['Помещение для уборочной машины', 'Прочие помещения', ''],
      ['Инвентарная', 'Прочие помещения', ''],
      ['Автостоянка', 'Автостоянка', ''],
      ['Лифтовой холл', 'ЛХ', ''],
      ['Тамбур', 'Тамбур', '']
    ]
  ));

  // 4. Колонки работ в Главном
  report.push(addMissingWorkColumns_(ss));

  // 5. Покажем итоговые токены для удобства
  report.push(getTokensReport_(ss));

  const summary = report.join('\n\n');
  Logger.log(summary);
  return summary;
}

/**
 * Показывает текущие токены подрядчиков. Запускать из редактора.
 */
function getTokens() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const report = getTokensReport_(ss);
  Logger.log(report);
  return report;
}

/**
 * Регенерирует токен указанного подрядчика. Запускать из редактора.
 * Поменяйте параметр idPodryadchik на нужный (например 'admin' или 'spectr').
 */
function regenerateToken(idPodryadchik) {
  const id = idPodryadchik || 'admin';
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.SHEET_PODRYADCHIKI);
  if (!sheet) throw new Error('Лист Подрядчики не найден');
  const data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === id) {
      const newToken = genToken_();
      sheet.getRange(i + 1, 4).setValue(newToken);
      const msg = 'Подрядчик "' + data[i][0] + '" (id=' + id + ') — новый токен: ' + newToken;
      Logger.log(msg);
      return msg;
    }
  }
  throw new Error('Подрядчик с ИД_подрядчик=' + id + ' не найден');
}

function setupSheet_(ss, name, headers, data) {
  var sheet = ss.getSheetByName(name);
  if (sheet) {
    return '[пропущено] Лист "' + name + '" уже существует (' + sheet.getLastRow() + ' строк)';
  }
  sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  if (data.length > 0) {
    sheet.getRange(2, 1, data.length, headers[0] ? headers.length : data[0].length).setValues(data);
  }
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return '[создан] Лист "' + name + '" — ' + data.length + ' строк';
}

function addMissingWorkColumns_(ss) {
  const glSheet = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
  const rabotySheet = ss.getSheetByName(CONFIG.SHEET_RABOTY);
  if (!glSheet || !rabotySheet) return '[пропущено] Главный или Работы не найдены, колонки не добавлены';

  const rabotyData = rabotySheet.getDataRange().getValues();
  const works = [];
  for (var i = 1; i < rabotyData.length; i++) {
    if (rabotyData[i][0]) works.push(String(rabotyData[i][0]).trim());
  }

  const headers = glSheet.getRange(1, 1, 1, glSheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim(); });
  const headerSet = {};
  headers.forEach(function (h) { headerSet[h] = true; });

  const toAdd = [];
  works.forEach(function (w) {
    if (!headerSet[w]) toAdd.push(w);
    if (!headerSet[w + ' СП']) toAdd.push(w + ' СП');
  });

  if (toAdd.length === 0) {
    return '[ok] Все колонки работ уже есть в Главном (' + works.length + ' работ)';
  }

  const startCol = glSheet.getLastColumn() + 1;
  // Проверка безопасности: добавляем только после реестровой части
  if (startCol <= CONFIG.REGISTRY_LAST_COL) {
    throw new Error('Безопасность: startCol ' + startCol + ' внутри реестра');
  }
  glSheet.getRange(1, startCol, 1, toAdd.length).setValues([toAdd]).setFontWeight('bold');
  return '[добавлены] Колонки работ в Главном: ' + toAdd.join(', ');
}

function getTokensReport_(ss) {
  const sheet = ss.getSheetByName(CONFIG.SHEET_PODRYADCHIKI);
  if (!sheet) return '[ошибка] Лист Подрядчики не найден';
  const data = sheet.getDataRange().getValues();
  const lines = ['Токены подрядчиков:'];
  for (var i = 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    lines.push('  ' + data[i][0] + ' (' + data[i][1] + ', ' + data[i][2] + '): ' + data[i][3]);
  }
  return lines.join('\n');
}

function genToken_() {
  var hex = '';
  while (hex.length < 16) {
    hex += Math.floor(Math.random() * 4294967295).toString(16);
  }
  return hex.slice(0, 16);
}

// =============================================================================
// Utilities
// =============================================================================

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Записывает строку в лист «Лог_входов». Если листа нет — создаёт с заголовками.
 * Безопасно: любая ошибка проглатывается, чтобы не сломать вход пользователя.
 */
function logLogin_(user, userAgent) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_LOG);
      sheet.appendRow(['Дата', 'Время', 'Подрядчик', 'ИД_подрядчик', 'Роль', 'User-Agent']);
      sheet.setFrozenRows(1);
      sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
      sheet.getRange('B:B').setNumberFormat('HH:mm:ss');
    }
    const now = new Date();
    const role = user.isAdmin ? 'Админ' : (user.isViewer ? 'Наблюдатель' : 'Подрядчик');
    sheet.appendRow([
      Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd'),
      Utilities.formatDate(now, CONFIG.TIMEZONE, 'HH:mm:ss'),
      user.name || '',
      user.id || '',
      role,
      String(userAgent || '').slice(0, 200)
    ]);
  } catch (e) {
    // молчим — лог не должен мешать работе
  }
}

function assertWritableColumn(col) {
  if (col <= CONFIG.REGISTRY_LAST_COL) {
    throw new Error(
      'Запись в реестровую часть листа Главный запрещена. ' +
      'Колонка ' + col + ' <= ' + CONFIG.REGISTRY_LAST_COL + ' (L). ' +
      'Только M+ (>= 13) разрешены для записи.'
    );
  }
}

function formatNumber(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (typeof v === 'number') {
    // округление до 2 знаков чтобы не было артефактов вроде 26.369999999999997
    return String(Math.round(v * 100) / 100).replace('.', ',');
  }
  return String(v);
}

function now() {
  return Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
}
