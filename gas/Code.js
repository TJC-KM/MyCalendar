/**
 * 出大事行事曆 後端（Google Apps Script）
 * 前端 index.html（GitHub Pages）用 POST 呼叫，資料存在試算表的「活動」「待辦」「設定」分頁。
 * 第一次使用：在編輯器執行一次 setup()（會要求授權），再部署成網頁應用程式。
 */

const SPREADSHEET_ID = '1F_Y9K35er3HHPVVjgxmc6KAcQoxz3s1MH9a3EnxquyE';
const EVENT_SHEET = '活動';
const TASK_SHEET = '待辦';
const SETTING_SHEET = '設定';
const TIME_ZONE = 'Asia/Taipei';

const STATUS = {
  PENDING: '待確認',
  CONFIRMED: '已確認',
  DISCUSS: '再議',
  CANCELLED: '取消',
};

/** 「參加」欄除了兩位成員的名字，還可以填這個 */
const BOTH = '兩人';

/** [程式用的名稱, 試算表表頭]；程式依表頭找欄位，欄位順序可以調 */
const COLUMNS = [
  ['id', 'ID'],
  ['startDate', '開始日期'],
  ['endDate', '結束日期'],
  ['startTime', '開始時間'],
  ['endTime', '結束時間'],
  ['period', '時段'],
  ['title', '標題'],
  ['who', '參加'],
  ['note', '備註'],
  ['status', '狀態'],
  ['proposedBy', '提出人'],
  ['repliedBy', '回覆人'],
  ['repliedAt', '回覆時間'],
  ['createdBy', '建立人'],
  ['createdAt', '建立時間'],
  ['updatedAt', '修改時間'],
  ['log', '留言紀錄'],
  ['location', '地點'],
];

/** 改了這些欄位，要另一個人重新確認；只改備註不用 */
const CONFIRM_FIELDS = ['startDate', 'endDate', 'startTime', 'endTime', 'period', 'title', 'location', 'who'];

/** 「時段」可以填的值；空白＝整天。有填開始時間時，時段依開始時間自動決定 */
const PERIODS = ['上午', '下午', '晚上'];

/** 「待辦」分頁的欄位；完成填 Y */
const TASK_COLUMNS = [
  ['id', 'ID'],
  ['text', '內容'],
  ['who', '負責'],
  ['dueDate', '期限'],
  ['done', '完成'],
  ['doneBy', '完成人'],
  ['doneAt', '完成時間'],
  ['createdBy', '建立人'],
  ['createdAt', '建立時間'],
  ['updatedAt', '修改時間'],
];

/** 已完成的待辦，回傳最近幾天內完成的就好 */
const TASK_DONE_KEEP_DAYS = 30;

/** GitHub Pages 的網址，setup 用它組出秘密連結 */
const PAGE_URL = 'https://tjc-km.github.io/MyCalendar/';

/** 「設定」分頁記在快取裡，省掉每次開試算表；在試算表改設定最多 5 分鐘後生效，執行 setup 立即生效 */
const SETTINGS_CACHE_KEY = 'settings';
const SETTINGS_CACHE_SECONDS = 300;

const SETTING_ROWS = [
  ['成員1', '老公', '顯示在網頁上的名字。改名後，舊活動的建立人、提出人不會跟著改'],
  ['成員2', '老婆', '同上'],
  ['鑰匙', '', '執行 setup 時自動產生，沒有它網頁不能用。要換一組：清空這格再執行 setup，舊連結就失效，兩支手機都要點新連結'],
  ['秘密連結', '', '把這個網址傳給要用的人，手機點一次就會記住。執行 setup 會自動更新'],
];

function doGet() {
  return ContentService.createTextOutput('出大事行事曆 API 運作中');
}

function doPost(e) {
  let result;
  try {
    const req = JSON.parse(e.postData.contents);
    result = handleRequest(req);
  } catch (err) {
    result = { ok: false, error: err.message || String(err), code: err.code || '' };
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function handleRequest(req) {
  const settings = readSettings();
  if (!settings.key) {
    fail('還沒有鑰匙：請在 GAS 編輯器執行一次 setup', 'NO_KEY');
  }
  if (String(req.key || '').trim() !== settings.key) {
    fail('連結不對，或鑰匙已經換了：請用試算表「設定」分頁的秘密連結重新打開', 'KEY');
  }
  if (req.action === 'hello') {
    return { ok: true, members: settings.members };
  }

  const who = String(req.who || '');
  if (settings.members.indexOf(who) < 0) {
    fail('找不到「' + who + '」，請重新選擇你是哪一位', 'WHO');
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  if (req.action === 'list') {
    return listEvents(ss, settings, who, req);
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // 拿到鎖之後才讀，才不會蓋掉對方剛存的資料
    if (String(req.action).indexOf('task') === 0) {
      const result = handleTask(ss, settings, who, req);
      SpreadsheetApp.flush();
      return result;
    }
    const table = readEvents(ss);
    let event;
    if (req.action === 'save') {
      event = saveEvent(table, settings, who, req);
    } else if (req.action === 'respond') {
      event = respondEvent(table, who, req);
    } else if (req.action === 'comment') {
      event = commentEvent(table, who, req);
    } else if (req.action === 'cancel') {
      event = cancelEvent(table, who, req);
    } else if (req.action === 'repropose') {
      event = reproposeEvent(table, who, req);
    } else {
      fail('不認得的動作：' + req.action);
    }
    SpreadsheetApp.flush();
    return { ok: true, event: event };
  } finally {
    lock.releaseLock();
  }
}

/**
 * 查詢區間內的活動，另外回傳「等我回覆」與「我提的、等對方確認」兩份清單（不限區間，只看還沒過的）
 */
function listEvents(ss, settings, who, req) {
  const from = normalizeDate(req.from);
  const to = normalizeDate(req.to);
  if (!from || !to || from > to) {
    fail('查詢的日期區間不對');
  }
  const table = readEvents(ss);
  const today = todayString();
  const events = [];
  const todo = [];
  const waiting = [];
  table.items.forEach(function (item) {
    const ev = item.data;
    if (!ev.startDate) {
      return;
    }
    if (ev.startDate <= to && ev.endDate >= from) {
      events.push(ev);
    }
    if (ev.endDate < today) {
      return;
    }
    if (needsReply(ev, who)) {
      todo.push(ev);
    } else if (ev.status === STATUS.PENDING && ev.proposedBy === who) {
      waiting.push(ev);
    }
  });
  return {
    ok: true, members: settings.members, today: today,
    events: events, todo: todo, waiting: waiting, tasks: listTasks(ss),
  };
}

/**
 * 輪到這個人處理：對方提的還沒確認，或自己提的被對方說再議
 */
function needsReply(ev, who) {
  if (ev.status === STATUS.PENDING && ev.proposedBy !== who) {
    return true;
  }
  return ev.status === STATUS.DISCUSS && ev.proposedBy === who;
}

/**
 * 新增或修改活動；改到時間、地點等內容會變回待確認，換對方確認
 */
function saveEvent(table, settings, who, req) {
  const source = req.event || {};
  const input = cleanInput(source, settings);
  const now = nowString();

  if (!source.id) {
    const created = {
      id: Utilities.getUuid().replace(/-/g, '').slice(0, 12),
      startDate: input.startDate,
      endDate: input.endDate,
      startTime: input.startTime,
      endTime: input.endTime,
      period: input.period,
      title: input.title,
      location: input.location,
      who: input.who,
      note: input.note,
      status: STATUS.PENDING,
      proposedBy: who,
      repliedBy: '',
      repliedAt: '',
      createdBy: who,
      createdAt: now,
      updatedAt: now,
      log: '',
    };
    created.log = appendLog('', who, '新增');
    writeRow(table, nextEmptyRow(table), null, created);
    return created;
  }

  const item = findItem(table, source.id);
  checkFresh(item, req.baseUpdatedAt);
  const ev = item.data;
  const changes = describeChanges(ev, input);
  if (changes.length === 0) {
    return ev;
  }

  let needConfirm = false;
  CONFIRM_FIELDS.forEach(function (key) {
    if (ev[key] !== input[key]) {
      needConfirm = true;
    }
  });
  Object.keys(input).forEach(function (key) {
    ev[key] = input[key];
  });
  if (needConfirm) {
    ev.status = STATUS.PENDING;
    ev.proposedBy = who;
    ev.repliedBy = '';
    ev.repliedAt = '';
  }
  ev.updatedAt = now;
  ev.log = appendLog(ev.log, who, '修改 ' + changes.join('、'));
  writeRow(table, item.row, item.raw, ev);
  return ev;
}

/**
 * 另一個人回覆：確認或再議
 */
function respondEvent(table, who, req) {
  const status = req.status;
  if (status !== STATUS.CONFIRMED && status !== STATUS.DISCUSS) {
    fail('回覆只能選確認或再議');
  }
  const item = findItem(table, req.id);
  checkFresh(item, req.baseUpdatedAt);
  const ev = item.data;
  if (ev.status === STATUS.CANCELLED) {
    fail('這個活動已經取消了');
  }
  if (ev.proposedBy === who) {
    fail('這是你提出的，要等對方回覆');
  }
  const message = cleanText(req.message, 500);
  const now = nowString();
  ev.status = status;
  ev.repliedBy = who;
  ev.repliedAt = now;
  ev.updatedAt = now;
  let text = status === STATUS.CONFIRMED ? '確認 ✓' : '要再議';
  if (message) {
    text += '：' + message;
  }
  ev.log = appendLog(ev.log, who, text);
  writeRow(table, item.row, item.raw, ev);
  return ev;
}

/**
 * 只留言，不改狀態，也不算修改（不會讓對方手上的畫面變成過期）
 */
function commentEvent(table, who, req) {
  const message = cleanText(req.message, 500);
  if (!message) {
    fail('留言是空的');
  }
  const item = findItem(table, req.id);
  const ev = item.data;
  ev.log = appendLog(ev.log, who, '留言：' + message);
  writeRow(table, item.row, item.raw, ev);
  return ev;
}

function cancelEvent(table, who, req) {
  const item = findItem(table, req.id);
  checkFresh(item, req.baseUpdatedAt);
  const ev = item.data;
  if (ev.status === STATUS.CANCELLED) {
    return ev;
  }
  const message = cleanText(req.message, 500);
  ev.status = STATUS.CANCELLED;
  ev.proposedBy = who;
  ev.repliedBy = '';
  ev.repliedAt = '';
  ev.updatedAt = nowString();
  ev.log = appendLog(ev.log, who, '取消活動' + (message ? '：' + message : ''));
  writeRow(table, item.row, item.raw, ev);
  return ev;
}

/**
 * 再議後維持原案、或恢復已取消的活動：回到待確認，由另一個人確認
 */
function reproposeEvent(table, who, req) {
  const item = findItem(table, req.id);
  checkFresh(item, req.baseUpdatedAt);
  const ev = item.data;
  if (ev.status !== STATUS.DISCUSS && ev.status !== STATUS.CANCELLED) {
    fail('只有再議或已取消的活動可以重新請對方確認');
  }
  const message = cleanText(req.message, 500);
  let text = ev.status === STATUS.CANCELLED ? '恢復活動' : '維持原案，請再確認';
  if (message) {
    text += '：' + message;
  }
  ev.status = STATUS.PENDING;
  ev.proposedBy = who;
  ev.repliedBy = '';
  ev.repliedAt = '';
  ev.updatedAt = nowString();
  ev.log = appendLog(ev.log, who, text);
  writeRow(table, item.row, item.raw, ev);
  return ev;
}

/**
 * 檢查前端送來的活動內容，回傳整理過的欄位
 */
function cleanInput(source, settings) {
  const input = {
    title: cleanText(source.title, 100),
    startDate: normalizeDate(source.startDate),
    endDate: normalizeDate(source.endDate),
    startTime: normalizeTime(source.startTime),
    endTime: normalizeTime(source.endTime),
    period: cleanText(source.period, 10),
    location: cleanText(source.location, 200),
    who: cleanText(source.who, 50) || BOTH,
    note: cleanText(source.note, 2000),
  };
  if (input.startTime) {
    input.period = periodOf(input.startTime);
  }
  if (input.period && PERIODS.indexOf(input.period) < 0) {
    fail('時段只能選整天、上午、下午、晚上');
  }
  if (!input.title) {
    fail('請填標題');
  }
  if (!input.startDate) {
    fail('請選日期');
  }
  if (!input.endDate) {
    input.endDate = input.startDate;
  }
  if (input.endDate < input.startDate) {
    fail('結束日期不能早於開始日期');
  }
  if (input.endTime && !input.startTime) {
    fail('有結束時間就要有開始時間');
  }
  if (input.startDate === input.endDate && input.endTime && input.endTime < input.startTime) {
    fail('結束時間不能早於開始時間');
  }
  if (input.who !== BOTH && settings.members.indexOf(input.who) < 0) {
    fail('「參加」要選兩人或其中一位');
  }
  return input;
}

/**
 * 列出這次改了什麼，寫進留言紀錄；沒改回傳空陣列
 */
function describeChanges(ev, input) {
  const changes = [];
  if (ev.startDate !== input.startDate || ev.endDate !== input.endDate) {
    changes.push('日期 ' + formatDateRange(ev) + '→' + formatDateRange(input));
  }
  if (ev.startTime !== input.startTime || ev.endTime !== input.endTime || ev.period !== input.period) {
    changes.push('時間 ' + formatTimeRange(ev) + '→' + formatTimeRange(input));
  }
  if (ev.title !== input.title) {
    changes.push('標題「' + input.title + '」');
  }
  if (ev.location !== input.location) {
    changes.push(input.location ? '地點「' + input.location + '」' : '拿掉地點');
  }
  if (ev.who !== input.who) {
    changes.push('參加改成' + input.who);
  }
  if (ev.note !== input.note) {
    changes.push('備註');
  }
  return changes;
}

function formatDateRange(ev) {
  const start = shortDate(ev.startDate);
  if (ev.endDate && ev.endDate !== ev.startDate) {
    return start + '～' + shortDate(ev.endDate);
  }
  return start;
}

function formatTimeRange(ev) {
  if (!ev.startTime) {
    return ev.period || '整天';
  }
  if (ev.endTime) {
    return ev.startTime + '-' + ev.endTime;
  }
  return ev.startTime;
}

/**
 * 依開始時間決定時段：中午 12 點前上午、傍晚 6 點前下午，之後晚上
 */
function periodOf(time) {
  const hour = Number(String(time).slice(0, 2));
  if (hour < 12) {
    return '上午';
  }
  if (hour < 18) {
    return '下午';
  }
  return '晚上';
}

function shortDate(dateString) {
  const parts = dateString.split('-');
  return Number(parts[1]) + '/' + Number(parts[2]);
}

function readSettings() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(SETTINGS_CACHE_KEY);
  if (cached) {
    return JSON.parse(cached);
  }
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(SETTING_SHEET);
  if (!sheet) {
    fail('試算表沒有「設定」分頁：請在 GAS 編輯器執行一次 setup');
  }
  const values = sheet.getDataRange().getDisplayValues();
  const map = {};
  values.forEach(function (row) {
    map[String(row[0]).trim()] = String(row[1] === undefined ? '' : row[1]).trim();
  });
  const members = [map['成員1'] || '', map['成員2'] || ''];
  if (!members[0] || !members[1] || members[0] === members[1] || members.indexOf(BOTH) >= 0) {
    fail('「設定」分頁的成員1、成員2 要填兩個不同的名字，也不能叫「' + BOTH + '」');
  }
  const settings = { members: members, key: map['鑰匙'] || '' };
  cache.put(SETTINGS_CACHE_KEY, JSON.stringify(settings), SETTINGS_CACHE_SECONDS);
  return settings;
}

/**
 * 讀一個資料分頁；回傳每列的列號、原始值（寫回時保留不認得的欄位）、依欄位名稱整理的資料。
 * 分頁或欄位不見了（例如程式加了新欄位），會先自動補上再讀，不用手動執行 setup
 */
function readTable(ss, sheetName, columns, repaired) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return repairTable(ss, sheetName, columns, repaired, '試算表沒有「' + sheetName + '」分頁');
  }
  const values = sheet.getDataRange().getDisplayValues();
  const header = values[0].map(function (h) { return String(h).trim(); });
  const col = {};
  let missing = '';
  columns.forEach(function (c) {
    const index = header.indexOf(c[1]);
    if (index < 0 && !missing) {
      missing = c[1];
    }
    col[c[0]] = index;
  });
  if (missing) {
    return repairTable(ss, sheetName, columns, repaired, '「' + sheetName + '」分頁少了欄位「' + missing + '」');
  }

  const items = [];
  for (let r = 1; r < values.length; r++) {
    const raw = values[r];
    if (!String(raw[col.id]).trim()) {
      continue;
    }
    const data = {};
    columns.forEach(function (c) {
      data[c[0]] = String(raw[col[c[0]]]).trim();
    });
    items.push({ row: r + 1, raw: raw, data: data });
  }
  return { sheet: sheet, header: header, col: col, columns: columns, items: items };
}

/**
 * 用 setup 的同一套做法補分頁或欄位，再讀一次；補過還是不對才報錯
 */
function repairTable(ss, sheetName, columns, repaired, problem) {
  if (repaired) {
    fail(problem + '，自動補不上：請在 GAS 編輯器執行一次 setup');
  }
  if (sheetName === EVENT_SHEET) {
    setupEventSheet(ss);
  } else if (sheetName === TASK_SHEET) {
    setupTaskSheet(ss);
  }
  SpreadsheetApp.flush();
  return readTable(ss, sheetName, columns, true);
}

/**
 * 讀「活動」分頁，日期時間統一格式、狀態空白當待確認
 */
function readEvents(ss) {
  const table = readTable(ss, EVENT_SHEET, COLUMNS);
  table.items.forEach(function (item) {
    const ev = item.data;
    ev.startDate = normalizeDate(ev.startDate);
    ev.endDate = normalizeDate(ev.endDate) || ev.startDate;
    ev.startTime = normalizeTime(ev.startTime);
    ev.endTime = normalizeTime(ev.endTime);
    if (ev.startTime) {
      ev.period = periodOf(ev.startTime);
    } else if (PERIODS.indexOf(ev.period) < 0) {
      ev.period = '';
    }
    if (!ev.status) {
      ev.status = STATUS.PENDING;
    }
  });
  return table;
}

function findItem(table, id) {
  const target = String(id || '');
  for (let i = 0; i < table.items.length; i++) {
    if (table.items[i].data.id === target) {
      return table.items[i];
    }
  }
  const what = table.columns === TASK_COLUMNS ? '待辦' : '活動';
  fail('找不到這個' + what + '，可能被刪掉了，請重新整理', 'STALE');
}

/**
 * 前端拿到的資料如果已經被別人改過，就不讓它存，避免蓋掉
 */
function checkFresh(item, baseUpdatedAt) {
  if (String(baseUpdatedAt || '') !== item.data.updatedAt) {
    fail('這個活動剛剛被改過了，已重新整理，請再看一次', 'STALE');
  }
}

function nextEmptyRow(table) {
  const sheet = table.sheet;
  const row = sheet.getLastRow() + 1;
  if (row > sheet.getMaxRows()) {
    const oldMax = sheet.getMaxRows();
    sheet.insertRowsAfter(oldMax, 100);
    sheet.getRange(oldMax + 1, 1, 100, table.header.length).setNumberFormat('@');
  }
  return row;
}

/**
 * 寫一整列；欄位格式在 setup 與加列時已設成純文字，試算表才不會把日期、時間自動轉型
 */
function writeRow(table, rowNumber, raw, ev) {
  const width = table.header.length;
  const out = [];
  for (let i = 0; i < width; i++) {
    out.push(raw && i < raw.length ? raw[i] : '');
  }
  table.columns.forEach(function (c) {
    out[table.col[c[0]]] = ev[c[0]] === undefined ? '' : String(ev[c[0]]);
  });
  table.sheet.getRange(rowNumber, 1, 1, width).setValues([out]);
}

/**
 * 待辦：新增／修改（taskSave）、打勾或取消打勾（taskDone）、刪除（taskDelete）
 */
function handleTask(ss, settings, who, req) {
  const table = readTable(ss, TASK_SHEET, TASK_COLUMNS);
  if (req.action === 'taskSave') {
    return { ok: true, task: saveTask(table, settings, who, req.task || {}) };
  }
  if (req.action === 'taskDone') {
    return { ok: true, task: doneTask(table, who, req) };
  }
  if (req.action === 'taskDelete') {
    const item = findItem(table, req.id);
    table.sheet.deleteRow(item.row);
    return { ok: true, id: item.data.id };
  }
  fail('不認得的動作：' + req.action);
}

/**
 * 沒完成的全部回傳；完成的只回傳最近 TASK_DONE_KEEP_DAYS 天內的
 */
function listTasks(ss) {
  const table = readTable(ss, TASK_SHEET, TASK_COLUMNS);
  const since = Utilities.formatDate(new Date(Date.now() - TASK_DONE_KEEP_DAYS * 86400000), TIME_ZONE, 'yyyy-MM-dd');
  const tasks = [];
  table.items.forEach(function (item) {
    const task = item.data;
    task.dueDate = normalizeDate(task.dueDate);
    if (task.done !== 'Y' || task.doneAt.slice(0, 10) >= since) {
      tasks.push(task);
    }
  });
  return tasks;
}

function saveTask(table, settings, who, source) {
  const input = {
    text: cleanText(source.text, 500),
    who: cleanText(source.who, 50) || BOTH,
    dueDate: normalizeDate(source.dueDate),
  };
  if (!input.text) {
    fail('待辦內容是空的');
  }
  if (input.who !== BOTH && settings.members.indexOf(input.who) < 0) {
    fail('「負責」要選兩人或其中一位');
  }
  const now = nowString();
  if (!source.id) {
    const created = {
      id: Utilities.getUuid().replace(/-/g, '').slice(0, 12),
      text: input.text,
      who: input.who,
      dueDate: input.dueDate,
      done: '',
      doneBy: '',
      doneAt: '',
      createdBy: who,
      createdAt: now,
      updatedAt: now,
    };
    writeRow(table, nextEmptyRow(table), null, created);
    return created;
  }
  const item = findItem(table, source.id);
  const task = item.data;
  task.text = input.text;
  task.who = input.who;
  task.dueDate = input.dueDate;
  task.updatedAt = now;
  writeRow(table, item.row, item.raw, task);
  return task;
}

function doneTask(table, who, req) {
  const item = findItem(table, req.id);
  const task = item.data;
  const now = nowString();
  if (req.done) {
    task.done = 'Y';
    task.doneBy = who;
    task.doneAt = now;
  } else {
    task.done = '';
    task.doneBy = '';
    task.doneAt = '';
  }
  task.updatedAt = now;
  writeRow(table, item.row, item.raw, task);
  return task;
}

function appendLog(log, who, text) {
  const line = Utilities.formatDate(new Date(), TIME_ZONE, 'M/d HH:mm') + ' ' + who + ' ' + text;
  return log ? log + '\n' + line : line;
}

function cleanText(value, maxLength) {
  return String(value === undefined || value === null ? '' : value).trim().slice(0, maxLength);
}

/**
 * 接受 2026-11-20、2026/11/20，統一成 2026-11-20；認不得回傳空字串
 */
function normalizeDate(value) {
  const match = String(value || '').trim().match(/^(\d{4})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!match) {
    return '';
  }
  return match[1] + '-' + pad2(match[2]) + '-' + pad2(match[3]);
}

/**
 * 接受 18:00、下午 6:00:00、6:00 PM，統一成 18:00；認不得回傳空字串
 */
function normalizeTime(value) {
  const text = String(value || '').trim();
  const match = text.match(/(\d{1,2}):(\d{2})/);
  if (!match) {
    return '';
  }
  let hour = Number(match[1]);
  if (/下午|晚上|PM/i.test(text) && hour < 12) {
    hour += 12;
  }
  if (/上午|凌晨|AM/i.test(text) && hour === 12) {
    hour = 0;
  }
  if (hour > 23 || Number(match[2]) > 59) {
    return '';
  }
  return pad2(hour) + ':' + match[2];
}

function pad2(value) {
  return ('0' + Number(value)).slice(-2);
}

function todayString() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd');
}

function nowString() {
  return Utilities.formatDate(new Date(), TIME_ZONE, 'yyyy-MM-dd HH:mm:ss');
}

function fail(message, code) {
  const err = new Error(message);
  err.code = code || '';
  throw err;
}

/**
 * 建立或補齊「活動」「待辦」「設定」分頁。可以重複執行，不會動到已經有的資料
 */
function setup() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  setupEventSheet(ss);
  setupTaskSheet(ss);
  setupSettingSheet(ss);
  CacheService.getScriptCache().remove(SETTINGS_CACHE_KEY);
  Logger.log('完成。秘密連結在「設定」分頁；要改名字也在那裡。');
}

function setupEventSheet(ss) {
  let sheet = ss.getSheetByName(EVENT_SHEET);
  if (!sheet) {
    const old = ss.getSheetByName('工作表1');
    if (old && old.getLastRow() <= 1) {
      old.setName(EVENT_SHEET);
      sheet = old;
    } else {
      sheet = ss.insertSheet(EVENT_SHEET, 0);
    }
  }
  const wide = { '標題': 160, '地點': 120, '備註': 200, '留言紀錄': 320, '建立時間': 140, '修改時間': 140, '回覆時間': 140 };
  setupHeader(sheet, COLUMNS, wide);
}

function setupTaskSheet(ss) {
  let sheet = ss.getSheetByName(TASK_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(TASK_SHEET, 1);
  }
  const wide = { '內容': 260, '完成時間': 140, '建立時間': 140, '修改時間': 140 };
  setupHeader(sheet, TASK_COLUMNS, wide);
}

/**
 * 空的分頁寫入全部表頭；已有資料的只在後面補缺的欄位。整頁設成純文字
 */
function setupHeader(sheet, columns, wide) {
  const labels = columns.map(function (c) { return c[1]; });
  if (sheet.getLastRow() <= 1) {
    const oldWidth = Math.max(sheet.getLastColumn(), 1);
    sheet.getRange(1, 1, sheet.getMaxRows(), oldWidth).clearContent().clearDataValidations();
    sheet.getRange(1, 1, 1, labels.length).setValues([labels]);
  } else {
    const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
    const missing = labels.filter(function (label) { return header.indexOf(label) < 0; });
    if (missing.length > 0) {
      sheet.getRange(1, header.length + 1, 1, missing.length).setValues([missing]);
    }
  }

  const width = sheet.getLastColumn();
  sheet.getRange(1, 1, sheet.getMaxRows(), width).setNumberFormat('@');
  sheet.getRange(1, 1, 1, width).setFontWeight('bold').setBackground('#eef2ff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidths(1, width, 90);
  const header = sheet.getRange(1, 1, 1, width).getDisplayValues()[0];
  Object.keys(wide).forEach(function (label) {
    const index = header.indexOf(label);
    if (index >= 0) {
      sheet.setColumnWidth(index + 1, wide[label]);
    }
  });
}

function setupSettingSheet(ss) {
  let sheet = ss.getSheetByName(SETTING_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SETTING_SHEET);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['項目', '值', '說明']]);
  }
  sheet.getRange(1, 1, sheet.getMaxRows(), 3).setNumberFormat('@');
  const existing = sheet.getDataRange().getDisplayValues().map(function (row) { return String(row[0]).trim(); });
  SETTING_ROWS.forEach(function (row) {
    if (existing.indexOf(row[0]) < 0) {
      sheet.appendRow(row);
    }
  });

  // 鑰匙是空的才產生新的；秘密連結每次都照目前的鑰匙重寫
  const values = sheet.getDataRange().getDisplayValues();
  let keyRow = 0;
  let linkRow = 0;
  values.forEach(function (row, index) {
    const name = String(row[0]).trim();
    if (name === '鑰匙') {
      keyRow = index + 1;
    }
    if (name === '秘密連結') {
      linkRow = index + 1;
    }
  });
  let key = String(values[keyRow - 1][1]).trim();
  if (!key) {
    key = Utilities.getUuid().replace(/-/g, '');
    sheet.getRange(keyRow, 2, 1, 1).setValues([[key]]);
  }
  sheet.getRange(linkRow, 2, 1, 1).setValues([[PAGE_URL + '#k=' + key]]);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#eef2ff');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 160);
  sheet.setColumnWidth(3, 480);
}
