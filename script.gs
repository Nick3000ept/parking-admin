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
  SHEET_EQUIP: 'Монтаж_оборудования',
  SHEET_COMMENTS: 'Комментарии_помещений',
  SHEET_EQUIP_LOG: 'Журнал_монтажа',

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

  // Колонки листа Монтаж_оборудования (1-based). Одна строка = одна работа
  // в одном помещении. A–F — справочная часть (сайт не пишет: номер, корпус,
  // этаж, наименование, работа, подрядчик), G–I — процент/статус/журнал.
  EQUIP_COL_NUM: 1,
  EQUIP_COL_KORP: 2,
  EQUIP_COL_FLOOR: 3,
  EQUIP_COL_NAME: 4,
  EQUIP_COL_WORK: 5,
  EQUIP_COL_SP: 6,
  EQUIP_COL_PCT: 7,
  EQUIP_COL_STATUS: 8,
  EQUIP_COL_UPDATED: 9,
  EQUIP_FIRST_WRITE_COL: 7,
  EQUIP_LAST_WRITE_COL: 9,
  // Помещения какой группы панели попадают в лист при setupEquip
  EQUIP_PANEL: 'Тех помещения',

  TIMEZONE: 'Europe/Moscow',
  LOCK_TIMEOUT_MS: 30000,

  ADMIN_ID: 'admin',
  VIEWER_ID: 'viewer'
};

// Наборы работ по типам тех-помещений (match — подстрока наименования,
// без учёта регистра). Используются только при первичном наполнении листа
// setupEquip; дальше источник правды — сами строки листа, админ может
// добавлять строки для любых помещений и работ вручную.
// Подстроки сравниваются с наименованием в нижнем регистре и с ё→е.
// Порядок важен: первое совпадение выигрывает.
const EQUIP_WORKSETS = [
  { match: 'электрощитов', works: ['Монтаж кабеля', 'Монтаж оборудования', 'Расключение', 'Монтаж шинопровода', 'ПНР'] },
  { match: 'кроссов',      works: ['Заказ оборудования', 'Монтаж лотка', 'Монтаж кабеля', 'Монтаж оборудования', 'Расключение'] },
  { match: 'венткамера',   works: ['Монтаж оборудования', 'Монтаж воздуховодов'] },
  { match: 'итп',          works: ['Монтаж гребенки', 'Разводка трубопроводов'] },
  { match: 'цтп',          works: ['Монтаж оборудования', 'Разводка трубопроводов', 'Монтаж УКУТ'] },
  { match: 'узел учета',   works: ['Монтаж УУ', 'Монтаж трубопроводов'] },
  { match: 'кнс',          works: ['Монтаж оборудования', 'Монтаж трубопроводов'] },
  { match: 'насосная',     works: ['Монтаж оборудования', 'Монтаж трубопроводов'] },
  { match: 'тп-',          works: ['Монтаж оборудования', 'Расключение оборудования'] }
];

const EQUIP_STATUSES = ['Не начато', 'В работе', 'Готово'];

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

    if (action === 'setupEquip') {
      return jsonResponse(setupEquipSheet(user));
    }

    if (action === 'equipStats') {
      return jsonResponse(equipStats());
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
    if (body.action === 'setEquip') {
      return jsonResponse(setEquip(user, body.num, body.work, body.pct));
    }
    if (body.action === 'setEquipBatch') {
      return jsonResponse(setEquipBatch(user, body.items));
    }
    if (body.action === 'setComment') {
      return jsonResponse(setComment(user, body.num, body.text));
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
      equip: {},
      comments: {},
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
    equip: loadEquip_(ss),
    comments: loadComments_(ss),
    server_time: now(),
    version_hash: computeVersionHash(headers, works.length, Object.keys(tipyMap).length)
  };
}

/**
 * Читает лист Монтаж_оборудования (одна строка = одна работа в помещении)
 * в map: номер помещения -> { works: [{ work, sp, pct, status, upd }] }.
 * Листа нет или он старой структуры (E1 не «Работа») — пустой map:
 * фронт просто не покажет блок.
 */
function loadEquip_(ss) {
  const equip = {};
  const sheet = ss.getSheetByName(CONFIG.SHEET_EQUIP);
  if (!sheet || sheet.getLastRow() < 2) return equip;
  if (String(sheet.getRange(1, CONFIG.EQUIP_COL_WORK).getValue() || '').trim() !== 'Работа') return equip;

  const data = sheet.getRange(1, 1, sheet.getLastRow(), CONFIG.EQUIP_COL_UPDATED).getValues();
  for (var i = 1; i < data.length; i++) {
    const num = String(data[i][CONFIG.EQUIP_COL_NUM - 1] || '').trim();
    const work = String(data[i][CONFIG.EQUIP_COL_WORK - 1] || '').trim();
    if (!num || !work) continue;
    if (!equip[num]) equip[num] = { works: [] };
    equip[num].works.push({
      work: work,
      sp: String(data[i][CONFIG.EQUIP_COL_SP - 1] || '').trim(),
      pct: normalizePct_(data[i][CONFIG.EQUIP_COL_PCT - 1]),
      status: String(data[i][CONFIG.EQUIP_COL_STATUS - 1] || '').trim() || 'Не начато',
      upd: String(data[i][CONFIG.EQUIP_COL_UPDATED - 1] || '')
    });
  }
  return equip;
}

/**
 * Читает лист Комментарии_помещений в map: номер помещения -> { text, upd }.
 * Листа нет — пустой map (создастся при первом сохранении комментария).
 */
function loadComments_(ss) {
  const comments = {};
  const sheet = ss.getSheetByName(CONFIG.SHEET_COMMENTS);
  if (!sheet || sheet.getLastRow() < 2) return comments;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  for (var i = 0; i < data.length; i++) {
    const num = String(data[i][0] || '').trim();
    const text = String(data[i][1] || '');
    if (!num || !text) continue;
    comments[num] = { text: text, upd: String(data[i][2] || '') };
  }
  return comments;
}

/** Процент из ячейки -> целое 0..100. Пустое/мусор -> 0. */
function normalizePct_(v) {
  var n = Math.round(Number(v));
  if (isNaN(n) || n < 0) return 0;
  if (n > 100) return 100;
  return n;
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
 * Масштабируется на сотни отметок: три bulk-чтения (Ведомость_работ, весь лист
 * Главный) + вся валидация в памяти + две групповые записи через RangeList —
 * вместо пары обращений к Sheets API на каждую отметку. Запись точечная по
 * ячейкам (не целыми колонками), поэтому параллельные ручные правки в Sheets
 * не затираются.
 *
 * marks: [{ action: 'set'|'clear', num, id_raboty, hint }] — hint игнорируется
 * (оставлен в протоколе для совместимости с уже открытыми сессиями фронта).
 * Возвращает: { ok: true, results: [{ ok, num, id_raboty, date?, error?, note? }] }
 */
function setMarks(user, marks) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };
  if (!Array.isArray(marks) || marks.length === 0) {
    return { ok: false, error: 'Пустой батч' };
  }

  const results = new Array(marks.length);
  withLock_(function () {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
    if (!sheet) {
      for (var k = 0; k < marks.length; k++) {
        results[k] = { ok: false, num: marks[k].num, id_raboty: marks[k].id_raboty, error: 'Лист "' + CONFIG.SHEET_GLAVNY + '" не найден' };
      }
      return;
    }

    // ИД_работы -> Полное название (одно чтение Ведомость_работ)
    const workNameById = {};
    const rabotySheet = ss.getSheetByName(CONFIG.SHEET_RABOTY);
    if (rabotySheet) {
      const rabotyData = rabotySheet.getDataRange().getValues();
      for (var w = 1; w < rabotyData.length; w++) {
        const id = String(rabotyData[w][3] || '').trim();
        if (id) workNameById[id] = String(rabotyData[w][0] || '').trim();
      }
    }

    // Весь Главный одним чтением: заголовки, строки помещений, текущие значения
    const values = sheet.getRange(1, 1, sheet.getLastRow(), sheet.getLastColumn()).getValues();
    const headers = values[0];

    // Колонки работы ищем лениво и кэшируем — в батче обычно 1-2 разные работы
    const colsByWorkId = {};
    function colsForWork_(idRaboty) {
      const key = String(idRaboty).trim();
      if (colsByWorkId[key]) return colsByWorkId[key];
      const workName = workNameById[key];
      var out;
      if (!workName) {
        out = { error: 'Работа не найдена: ' + idRaboty };
      } else {
        var colDate = null;
        var colSp = null;
        for (var j = 0; j < headers.length; j++) {
          const h = String(headers[j] || '').trim();
          if (h === workName) colDate = j + 1;
          if (h === workName + ' СП') colSp = j + 1;
        }
        out = (colDate && colSp)
          ? { colDate: colDate, colSp: colSp }
          : { error: 'Колонки работы "' + workName + '" не найдены в листе Главный' };
      }
      colsByWorkId[key] = out;
      return out;
    }

    // Номер помещения -> строка листа (первое совпадение, как в locateCell)
    const rowByNum = {};
    for (var r = 1; r < values.length; r++) {
      const numVal = String(values[r][CONFIG.COL_NUM - 1] || '').trim();
      if (numVal && !(numVal in rowByNum)) rowByNum[numVal] = r + 1;
    }

    const today = new Date();
    const todayStr = Utilities.formatDate(today, CONFIG.TIMEZONE, 'yyyy-MM-dd');
    const setCells = [];   // A1-адреса для групповой записи даты
    const clearCells = []; // A1-адреса для групповой очистки

    for (var i = 0; i < marks.length; i++) {
      const m = marks[i] || {};
      const num = m.num;
      const idRaboty = m.id_raboty;
      const action = m.action;
      try {
        const cols = colsForWork_(idRaboty);
        if (cols.error) { results[i] = { ok: false, num: num, id_raboty: idRaboty, error: cols.error }; continue; }

        const row = rowByNum[String(num).trim()];
        if (!row) { results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Помещение не найдено: ' + num }; continue; }

        const rowVals = values[row - 1];
        const spRaw = rowVals[cols.colSp - 1];
        const spValue = spRaw ? String(spRaw).trim() : '';
        const dateRaw = rowVals[cols.colDate - 1];
        var currentDate = '';
        if (dateRaw instanceof Date) {
          currentDate = Utilities.formatDate(dateRaw, CONFIG.TIMEZONE, 'yyyy-MM-dd');
        } else if (dateRaw !== '' && dateRaw !== null && dateRaw !== undefined) {
          currentDate = String(dateRaw);
        }

        if (!user.isAdmin) {
          if (spValue !== user.name) {
            results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Эта работа не назначена вам' };
            continue;
          }
        } else {
          if (action === 'set' && !spValue) {
            results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Работа не назначена никому, нечего отмечать' };
            continue;
          }
        }

        assertWritableColumn(cols.colDate);

        if (action === 'set') {
          if (currentDate) {
            results[i] = { ok: true, num: num, id_raboty: idRaboty, date: currentDate, note: 'already_set' };
            continue;
          }
          setCells.push(a1_(row, cols.colDate));
          rowVals[cols.colDate - 1] = today; // чтобы повтор той же ячейки в батче увидел дату
          results[i] = { ok: true, num: num, id_raboty: idRaboty, date: todayStr };
        } else if (action === 'clear') {
          if (!currentDate) {
            results[i] = { ok: true, num: num, id_raboty: idRaboty, date: '', note: 'already_empty' };
            continue;
          }
          clearCells.push(a1_(row, cols.colDate));
          rowVals[cols.colDate - 1] = '';
          results[i] = { ok: true, num: num, id_raboty: idRaboty, date: '' };
        } else {
          results[i] = { ok: false, num: num, id_raboty: idRaboty, error: 'Unknown action: ' + action };
        }
      } catch (err) {
        results[i] = { ok: false, num: num, id_raboty: idRaboty, error: String(err && err.message ? err.message : err) };
      }
    }

    if (setCells.length > 0) sheet.getRangeList(setCells).setValue(today);
    if (clearCells.length > 0) sheet.getRangeList(clearCells).clearContent();
  });

  return { ok: true, results: results };
}

/**
 * Запись факта по монтажу оборудования: процент одной работы в одном помещении
 * (строка листа Монтаж_оборудования ищется по паре номер+работа при каждом
 * запросе — защита от сдвижки строк). Статус вычисляется из процента
 * (0 — Не начато, 1–99 — В работе, 100 — Готово) и пишется автоматически.
 *
 * Права: Админ — любая строка; подрядчик — только строки, где он вписан
 * в колонку «Подрядчик»; Наблюдатель — нет.
 */
function setEquip(user, num, work, pct) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };

  const workStr = String(work || '').trim();
  if (!workStr) return { ok: false, error: 'Не указана работа' };
  if (pct === null || pct === undefined || pct === '' || isNaN(Number(pct))) {
    return { ok: false, error: 'Процент не распознан: ' + pct };
  }
  const pctNum = normalizePct_(pct);
  const statusStr = pctNum >= 100 ? 'Готово' : (pctNum > 0 ? 'В работе' : 'Не начато');

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_EQUIP);
  if (!sheet) return { ok: false, error: 'Лист "' + CONFIG.SHEET_EQUIP + '" не найден' };
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: false, error: 'Лист "' + CONFIG.SHEET_EQUIP + '" пуст' };

  const refData = sheet.getRange(2, 1, lastRow - 1, CONFIG.EQUIP_COL_PCT).getValues();
  var row = null;
  var sp = '';
  var oldPct = 0;
  for (var k = 0; k < refData.length; k++) {
    if (String(refData[k][CONFIG.EQUIP_COL_NUM - 1] || '').trim() === String(num).trim() &&
        String(refData[k][CONFIG.EQUIP_COL_WORK - 1] || '').trim() === workStr) {
      row = k + 2;
      sp = String(refData[k][CONFIG.EQUIP_COL_SP - 1] || '').trim();
      oldPct = normalizePct_(refData[k][CONFIG.EQUIP_COL_PCT - 1]);
      break;
    }
  }
  if (!row) return { ok: false, error: 'Строка не найдена в листе монтажа: ' + num + ' / ' + workStr };

  if (!user.isAdmin && sp !== user.name) {
    return { ok: false, error: 'Эта работа в этом помещении не назначена вам' };
  }

  const updated = now() + ' ' + user.name;
  withLock_(function () {
    const ssw = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sh = ssw.getSheetByName(CONFIG.SHEET_EQUIP);
    sh.getRange(row, CONFIG.EQUIP_COL_PCT).setValue(pctNum);
    sh.getRange(row, CONFIG.EQUIP_COL_STATUS).setValue(statusStr);
    sh.getRange(row, CONFIG.EQUIP_COL_UPDATED).setValue(updated);
    if (pctNum !== oldPct) {
      logEquip_(ssw, user, [{ num: String(num).trim(), work: workStr, oldPct: oldPct, newPct: pctNum }]);
    }
  });

  return { ok: true, num: num, work: workStr, status: statusStr, pct: pctNum, updated: updated };
}

/**
 * Пакетная запись процентов монтажа: один HTTP, один Lock, одно чтение листа.
 * items: [{ num, work, pct }]. Повтор пары (num, work) в батче — последний
 * выигрывает. Запись группами через RangeList (ячейки с одинаковым значением
 * пишутся одним вызовом), точечно по ячейкам — параллельные ручные правки
 * других строк в Sheets не затираются.
 * Возвращает { ok: true, results: [{ ok, num, work, pct?, status?, error? }], updated }.
 */
function setEquipBatch(user, items) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };
  if (!Array.isArray(items) || items.length === 0) return { ok: false, error: 'Пустой батч' };
  if (items.length > 500) return { ok: false, error: 'Слишком большой батч (>500)' };

  const results = new Array(items.length);
  const updated = now() + ' ' + user.name;

  withLock_(function () {
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.SHEET_EQUIP);
    if (!sheet || sheet.getLastRow() < 2) {
      for (var k = 0; k < items.length; k++) {
        results[k] = { ok: false, num: items[k] && items[k].num, work: items[k] && items[k].work, error: 'Лист "' + CONFIG.SHEET_EQUIP + '" не найден или пуст' };
      }
      return;
    }

    // Одно чтение: пара (номер, работа) -> строка + подрядчик + текущий процент
    const refData = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.EQUIP_COL_PCT).getValues();
    const rowByKey = {};
    for (var r = 0; r < refData.length; r++) {
      const key = String(refData[r][CONFIG.EQUIP_COL_NUM - 1] || '').trim() + '|' +
        String(refData[r][CONFIG.EQUIP_COL_WORK - 1] || '').trim();
      if (!(key in rowByKey)) {
        rowByKey[key] = {
          row: r + 2,
          sp: String(refData[r][CONFIG.EQUIP_COL_SP - 1] || '').trim(),
          oldPct: normalizePct_(refData[r][CONFIG.EQUIP_COL_PCT - 1])
        };
      }
    }

    // Валидация в памяти; повтор пары в батче — последний выигрывает
    const finalByKey = {}; // key -> { row, pctNum, statusStr }
    for (var i = 0; i < items.length; i++) {
      const it = items[i] || {};
      const num = String(it.num || '').trim();
      const workStr = String(it.work || '').trim();
      if (!num || !workStr) {
        results[i] = { ok: false, num: it.num, work: it.work, error: 'Не указано помещение или работа' };
        continue;
      }
      if (it.pct === null || it.pct === undefined || it.pct === '' || isNaN(Number(it.pct))) {
        results[i] = { ok: false, num: num, work: workStr, error: 'Процент не распознан: ' + it.pct };
        continue;
      }
      const rec = rowByKey[num + '|' + workStr];
      if (!rec) {
        results[i] = { ok: false, num: num, work: workStr, error: 'Строка не найдена в листе монтажа' };
        continue;
      }
      if (!user.isAdmin && rec.sp !== user.name) {
        results[i] = { ok: false, num: num, work: workStr, error: 'Эта работа не назначена вам' };
        continue;
      }
      const pctNum = normalizePct_(it.pct);
      const statusStr = pctNum >= 100 ? 'Готово' : (pctNum > 0 ? 'В работе' : 'Не начато');
      finalByKey[num + '|' + workStr] = {
        row: rec.row, pctNum: pctNum, statusStr: statusStr,
        num: num, work: workStr, oldPct: rec.oldPct
      };
      results[i] = { ok: true, num: num, work: workStr, pct: pctNum, status: statusStr };
    }

    // Групповая запись: ячейки с одинаковым значением — одним setValue
    const pctGroups = {};    // pct -> [a1, ...]
    const statusGroups = {}; // статус -> [a1, ...]
    const updCells = [];
    for (var key in finalByKey) {
      if (!finalByKey.hasOwnProperty(key)) continue;
      const f = finalByKey[key];
      (pctGroups[f.pctNum] = pctGroups[f.pctNum] || []).push(a1_(f.row, CONFIG.EQUIP_COL_PCT));
      (statusGroups[f.statusStr] = statusGroups[f.statusStr] || []).push(a1_(f.row, CONFIG.EQUIP_COL_STATUS));
      updCells.push(a1_(f.row, CONFIG.EQUIP_COL_UPDATED));
    }
    for (var p in pctGroups) {
      if (pctGroups.hasOwnProperty(p)) sheet.getRangeList(pctGroups[p]).setValue(Number(p));
    }
    for (var s in statusGroups) {
      if (statusGroups.hasOwnProperty(s)) sheet.getRangeList(statusGroups[s]).setValue(s);
    }
    if (updCells.length > 0) sheet.getRangeList(updCells).setValue(updated);

    // Журнал изменений — для аналитики «за неделю/месяц»
    const logEntries = [];
    for (var lk in finalByKey) {
      if (!finalByKey.hasOwnProperty(lk)) continue;
      const lf = finalByKey[lk];
      if (lf.pctNum !== lf.oldPct) {
        logEntries.push({ num: lf.num, work: lf.work, oldPct: lf.oldPct, newPct: lf.pctNum });
      }
    }
    logEquip_(SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID), user, logEntries);
  });

  return { ok: true, results: results, updated: updated };
}

/**
 * Пишет изменения процентов в лист Журнал_монтажа (создаёт при первом вызове).
 * Колонки: Дата (настоящая дата+время), Номер, Работа, Было %, Стало %, Кто.
 * Любая ошибка проглатывается — журнал не должен ломать сохранение.
 */
function logEquip_(ss, user, entries) {
  if (!entries || entries.length === 0) return;
  try {
    var sheet = ss.getSheetByName(CONFIG.SHEET_EQUIP_LOG);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_EQUIP_LOG);
      sheet.getRange(1, 1, 1, 6).setValues([['Дата', 'Номер помещения', 'Работа', 'Было %', 'Стало %', 'Кто']]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd HH:mm:ss');
    }
    const stamp = new Date();
    const rows = [];
    for (var i = 0; i < entries.length; i++) {
      rows.push([stamp, entries[i].num, entries[i].work, entries[i].oldPct, entries[i].newPct, user.name || '']);
    }
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 6).setValues(rows);
  } catch (e) {
    // молчим — журнал вторичен
  }
}

/**
 * Комментарий к помещению — лист Комментарии_помещений (A: Номер, B: Комментарий,
 * C: Обновлено). Лист создаётся при первом сохранении. Права: любой активный
 * пользователь, кроме Наблюдателя. Пустой текст очищает комментарий.
 * Ввод экранируется от formula injection и ограничен 500 символами.
 */
function setComment(user, num, text) {
  if (user.isViewer) return { ok: false, error: 'У роли «Наблюдатель» нет прав на редактирование' };
  const numStr = String(num || '').trim();
  if (!numStr) return { ok: false, error: 'Не указано помещение' };

  var textStr = String(text === null || text === undefined ? '' : text).slice(0, 500).trim();
  // Защита от formula injection: значения на =, +, -, @ — апостроф спереди
  if (/^[=+\-@]/.test(textStr)) textStr = "'" + textStr;

  const updated = textStr ? (now() + ' ' + user.name) : '';
  withLock_(function () {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_COMMENTS);
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_COMMENTS);
      sheet.getRange(1, 1, 1, 3).setValues([['Номер помещения', 'Комментарий', 'Обновлено']]).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    const lastRow = sheet.getLastRow();
    var row = null;
    if (lastRow > 1) {
      const nums = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      for (var i = 0; i < nums.length; i++) {
        if (String(nums[i][0] || '').trim() === numStr) { row = i + 2; break; }
      }
    }
    if (row) {
      sheet.getRange(row, 2, 1, 2).setValues([[textStr, updated]]);
    } else if (textStr) {
      sheet.appendRow([numStr, textStr, updated]);
    }
  });

  return { ok: true, num: numStr, text: textStr.replace(/^'/, ''), updated: updated };
}

/** A1-адрес ячейки по (row, col) 1-based — для getRangeList. */
function a1_(row, col) {
  var letters = '';
  var c = col;
  while (c > 0) {
    letters = String.fromCharCode(65 + ((c - 1) % 26)) + letters;
    c = Math.floor((c - 1) / 26);
  }
  return letters + row;
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

/**
 * Сводка по монтажу оборудования для окна «Аналитика»:
 * текущий срез листа + прогресс за 7 и 30 дней из Журнала_монтажа.
 * Доступна любой авторизованной роли (только чтение).
 */
function equipStats() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);

  // Текущий срез
  const totals = { positions: 0, done: 0, inProgress: 0, notStarted: 0 };
  const eqSheet = ss.getSheetByName(CONFIG.SHEET_EQUIP);
  if (eqSheet && eqSheet.getLastRow() > 1) {
    const data = eqSheet.getRange(2, 1, eqSheet.getLastRow() - 1, CONFIG.EQUIP_COL_PCT).getValues();
    for (var i = 0; i < data.length; i++) {
      if (!String(data[i][CONFIG.EQUIP_COL_NUM - 1] || '').trim()) continue;
      totals.positions++;
      const pct = normalizePct_(data[i][CONFIG.EQUIP_COL_PCT - 1]);
      if (pct >= 100) totals.done++;
      else if (pct > 0) totals.inProgress++;
      else totals.notStarted++;
    }
  }

  // Прогресс из журнала: окна 7 и 30 дней от текущего момента
  function emptyWindow() { return { positions: 0, completed: 0, delta: 0, byWork: {} }; }
  const week = emptyWindow();
  const month = emptyWindow();
  const logSheet = ss.getSheetByName(CONFIG.SHEET_EQUIP_LOG);
  if (logSheet && logSheet.getLastRow() > 1) {
    const nowMs = Date.now();
    const weekFrom = nowMs - 7 * 24 * 3600 * 1000;
    const monthFrom = nowMs - 30 * 24 * 3600 * 1000;
    const log = logSheet.getRange(2, 1, logSheet.getLastRow() - 1, 6).getValues();
    const weekSeen = {};
    const monthSeen = {};
    for (var r = 0; r < log.length; r++) {
      const d = log[r][0];
      if (!(d instanceof Date)) continue;
      const t = d.getTime();
      if (t < monthFrom) continue;
      const key = String(log[r][1] || '').trim() + '|' + String(log[r][2] || '').trim();
      const work = String(log[r][2] || '').trim();
      const oldPct = normalizePct_(log[r][3]);
      const newPct = normalizePct_(log[r][4]);
      const entryDelta = newPct - oldPct;
      const completed = oldPct < 100 && newPct >= 100 ? 1 : 0;

      [{ win: month, seen: monthSeen, from: monthFrom },
       { win: week,  seen: weekSeen,  from: weekFrom }].forEach(function (w) {
        if (t < w.from) return;
        if (!w.seen[key]) { w.seen[key] = true; w.win.positions++; }
        w.win.delta += entryDelta;
        w.win.completed += completed;
        if (!w.win.byWork[work]) w.win.byWork[work] = { positions: 0, completed: 0, delta: 0, seen: {} };
        const bw = w.win.byWork[work];
        if (!bw.seen[key]) { bw.seen[key] = true; bw.positions++; }
        bw.delta += entryDelta;
        bw.completed += completed;
      });
    }
    // Служебные seen-наборы наружу не отдаём
    [week, month].forEach(function (w) {
      for (var k in w.byWork) {
        if (w.byWork.hasOwnProperty(k)) delete w.byWork[k].seen;
      }
    });
  }

  return { ok: true, totals: totals, week: week, month: month, server_time: now() };
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
      ['Заливка топпинга', 'Полы', 'Топпинг', 'down_topping'],
      ['Укладка линолеума', 'Полы', 'Линолеум', 'down_linoleum']
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
 * Создаёт лист Монтаж_оборудования (одна строка = одна работа в помещении)
 * и наполняет по наборам EQUIP_WORKSETS: электрощитовые, кроссовые, венткамеры
 * из группы «Тех помещения» Главного. Идемпотентно: лист новой структуры уже
 * есть — ничего не меняет. Лист СТАРОЙ структуры (плоские позиции в колонках)
 * не удаляется — переименовывается в архив.
 *
 * Вызывается через ?action=setupEquip&token=... (любой активный пользователь,
 * кроме Наблюдателя) или вручную из редактора: setupEquipSheetManual.
 */
function setupEquipSheet(user) {
  if (user && user.isViewer) {
    return { ok: false, error: 'У роли «Наблюдатель» нет прав на настройку' };
  }
  const result = { ok: true };
  withLock_(function () {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    var sheet = ss.getSheetByName(CONFIG.SHEET_EQUIP);
    var existingSheet = null; // лист новой структуры — дозаполним недостающие строки
    if (sheet) {
      const e1 = String(sheet.getRange(1, CONFIG.EQUIP_COL_WORK).getValue() || '').trim();
      if (e1 === 'Работа') {
        existingSheet = sheet;
      } else {
        // Старая структура — в архив (не удаляем данные)
        var archName = CONFIG.SHEET_EQUIP + '_архив';
        if (ss.getSheetByName(archName)) {
          archName += '_' + Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyyMMdd_HHmmss');
        }
        sheet.setName(archName);
        result.archived = archName;
      }
    }

    // Тип помещения -> группа панели
    const tipySheet = ss.getSheetByName(CONFIG.SHEET_TIPY);
    if (!tipySheet) throw new Error('Лист "' + CONFIG.SHEET_TIPY + '" не найден');
    const tipyData = tipySheet.getDataRange().getValues();
    const panelByName = {};
    for (var i = 1; i < tipyData.length; i++) {
      if (tipyData[i][0]) panelByName[String(tipyData[i][0]).trim()] = String(tipyData[i][1] || '').trim();
    }

    // Тех-помещения из Главного (реестровый порядок), подбор набора работ по имени
    const glSheet = ss.getSheetByName(CONFIG.SHEET_GLAVNY);
    if (!glSheet) throw new Error('Лист "' + CONFIG.SHEET_GLAVNY + '" не найден');
    const glData = glSheet.getDataRange().getValues();
    const rows = [];
    var roomsMatched = 0;
    for (var r = 1; r < glData.length; r++) {
      const num = String(glData[r][CONFIG.COL_NUM - 1] || '').trim();
      if (!num) continue;
      const name = String(glData[r][CONFIG.COL_NAME - 1] || '').trim();
      if (panelByName[name] !== CONFIG.EQUIP_PANEL) continue;

      var workset = null;
      const nameLow = name.toLowerCase().replace(/ё/g, 'е');
      for (var w = 0; w < EQUIP_WORKSETS.length; w++) {
        if (nameLow.indexOf(EQUIP_WORKSETS[w].match) !== -1) { workset = EQUIP_WORKSETS[w]; break; }
      }
      if (!workset) continue; // тип без набора работ — админ добавит строки вручную при необходимости

      roomsMatched++;
      for (var j = 0; j < workset.works.length; j++) {
        rows.push([
          num,
          String(glData[r][CONFIG.COL_KORP - 1] || ''),
          String(glData[r][CONFIG.COL_FLOOR - 1] || ''),
          name,
          workset.works[j],
          '', // Подрядчик — заполняет админ в таблице
          0, 'Не начато', '' // Процент / Статус / Обновлено
        ]);
      }
    }
    if (rows.length === 0) throw new Error('Не найдено помещений под наборы работ');

    const pctRule = SpreadsheetApp.newDataValidation()
      .requireNumberBetween(0, 100).setAllowInvalid(false).build();
    const statusRule = SpreadsheetApp.newDataValidation()
      .requireValueInList(EQUIP_STATUSES, true).setAllowInvalid(false).build();

    if (existingSheet) {
      // Дозаполнение: добавляем в конец только пары (помещение, работа),
      // которых ещё нет. Существующие строки (проценты, подрядчиков) не трогаем.
      sheet = existingSheet;
      const existingPairs = {};
      if (sheet.getLastRow() > 1) {
        const cur = sheet.getRange(2, 1, sheet.getLastRow() - 1, CONFIG.EQUIP_COL_WORK).getValues();
        for (var c = 0; c < cur.length; c++) {
          existingPairs[String(cur[c][CONFIG.EQUIP_COL_NUM - 1] || '').trim() + '|' +
            String(cur[c][CONFIG.EQUIP_COL_WORK - 1] || '').trim()] = true;
        }
      }
      const toAdd = [];
      for (var a = 0; a < rows.length; a++) {
        if (!existingPairs[String(rows[a][0]).trim() + '|' + String(rows[a][4]).trim()]) toAdd.push(rows[a]);
      }
      if (toAdd.length === 0) {
        result.note = 'up_to_date';
        result.rows = Math.max(0, sheet.getLastRow() - 1);
        return;
      }
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, toAdd.length, toAdd[0].length).setValues(toAdd);
      sheet.getRange(startRow, CONFIG.EQUIP_COL_PCT, toAdd.length, 1).setDataValidation(pctRule);
      sheet.getRange(startRow, CONFIG.EQUIP_COL_STATUS, toAdd.length, 1).setDataValidation(statusRule);
      result.added = toAdd.length;
      result.rows = sheet.getLastRow() - 1;
      return;
    }

    const headers = [
      'Номер помещения', 'Корпус', 'Этаж', 'Наименование',
      'Работа', 'Подрядчик', '% готовности', 'Статус', 'Обновлено'
    ];
    sheet = ss.insertSheet(CONFIG.SHEET_EQUIP);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    sheet.setFrozenRows(1);
    sheet.getRange(2, CONFIG.EQUIP_COL_PCT, rows.length, 1).setDataValidation(pctRule);
    sheet.getRange(2, CONFIG.EQUIP_COL_STATUS, rows.length, 1).setDataValidation(statusRule);
    sheet.autoResizeColumns(1, headers.length);

    result.created = true;
    result.rooms = roomsMatched;
    result.rows = rows.length;
  });
  return result;
}

/** Ручной запуск создания листа монтажа из редактора Apps Script. */
function setupEquipSheetManual() {
  const res = setupEquipSheet(null);
  Logger.log(JSON.stringify(res));
  return res;
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
