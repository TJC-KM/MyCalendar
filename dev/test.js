// 測 gas/Code.js 的規則：node dev/test.js
const assert = require('assert');
const { createGas } = require('./fake-gas');

let passed = 0;
function test(name, fn) {
  fn();
  passed++;
  console.log('ok  ' + name);
}

function settingValue(gas, name) {
  const row = gas.spreadsheet.getSheetByName('設定').cells.find((r) => r[0] === name);
  return row ? row[1] : undefined;
}

function freshGas() {
  const gas = createGas();
  gas.context.setup();
  gas.key = settingValue(gas, '鑰匙');
  return gas;
}

function call(gas, who, action, data) {
  return gas.post(Object.assign({ action, who, key: gas.key }, data || {}));
}

const A = '老公';
const B = '老婆';
const base = { title: '吃晚餐', startDate: '2030-01-10', startTime: '18:30', endTime: '20:00', location: '老地方', who: '兩人', note: '' };

test('setup 把舊的工作表1 改成活動、表頭換新', () => {
  const gas = createGas();
  gas.spreadsheet.sheets[0].cells = [['id', 'date', 'time', 'title', 'confirmed']];
  gas.context.setup();
  const sheet = gas.spreadsheet.getSheetByName('活動');
  assert.ok(sheet);
  assert.strictEqual(gas.spreadsheet.getSheetByName('工作表1'), null);
  assert.strictEqual(sheet.cells[0][0], 'ID');
  assert.strictEqual(sheet.cells[0].length, 17);
  assert.ok(!sheet.cells[0].includes('confirmed'));
  const setting = gas.spreadsheet.getSheetByName('設定');
  assert.deepStrictEqual(setting.cells.map((r) => r[0]), ['項目', '成員1', '成員2', '鑰匙', '秘密連結']);
});

test('setup 產生 32 碼鑰匙與秘密連結', () => {
  const gas = freshGas();
  assert.match(gas.key, /^[0-9a-f]{32}$/);
  assert.strictEqual(settingValue(gas, '秘密連結'), 'https://tjc-km.github.io/MyCalendar/#k=' + gas.key);
});

test('setup 重複執行：不重複加設定列、不換鑰匙、不清資料', () => {
  const gas = freshGas();
  gas.spreadsheet.getSheetByName('設定').cells[1][1] = '阿明';
  call(gas, '阿明', 'save', { event: base });
  gas.context.setup();
  assert.strictEqual(gas.spreadsheet.getSheetByName('設定').getLastRow(), 5);
  assert.strictEqual(settingValue(gas, '成員1'), '阿明');
  assert.strictEqual(settingValue(gas, '鑰匙'), gas.key);
  assert.strictEqual(gas.spreadsheet.getSheetByName('活動').getLastRow(), 2);
});

test('鑰匙清空再 setup 會換新的，舊連結失效', () => {
  const gas = freshGas();
  const oldKey = gas.key;
  const setting = gas.spreadsheet.getSheetByName('設定');
  setting.cells.find((r) => r[0] === '鑰匙')[1] = '';
  assert.strictEqual(call(gas, A, 'hello').code, 'NO_KEY');
  gas.context.setup();
  gas.key = settingValue(gas, '鑰匙');
  assert.notStrictEqual(gas.key, oldKey);
  assert.ok(settingValue(gas, '秘密連結').endsWith(gas.key));
  assert.strictEqual(gas.post({ action: 'hello', key: oldKey }).code, 'KEY');
});

test('沒帶鑰匙、鑰匙錯都擋下；hello 回名字；名字不對擋下', () => {
  const gas = freshGas();
  assert.strictEqual(gas.post({ action: 'hello' }).code, 'KEY');
  assert.strictEqual(gas.post({ action: 'list', who: A, key: 'wrong', from: '2030-01-01', to: '2030-01-31' }).code, 'KEY');
  const hello = gas.post({ action: 'hello', key: gas.key });
  assert.ok(hello.ok, hello.error);
  assert.deepStrictEqual(hello.members, [A, B]);
  assert.strictEqual(hello.key, undefined);
  assert.strictEqual(call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' }).ok, true);
  assert.strictEqual(call(gas, '路人', 'list', { from: '2030-01-01', to: '2030-01-31' }).code, 'WHO');
  assert.strictEqual(call(gas, '', 'save', { event: base }).code, 'WHO');
});

test('設定的兩個名字一樣或叫「兩人」時報錯', () => {
  const gas = freshGas();
  const setting = gas.spreadsheet.getSheetByName('設定');
  setting.cells[2][1] = A;
  assert.strictEqual(call(gas, A, 'hello').ok, false);
  setting.cells[2][1] = '兩人';
  assert.strictEqual(call(gas, A, 'hello').ok, false);
});

test('新增：待確認、提出人是自己、留言紀錄有新增', () => {
  const gas = freshGas();
  const res = call(gas, A, 'save', { event: base });
  assert.ok(res.ok, res.error);
  assert.strictEqual(res.event.status, '待確認');
  assert.strictEqual(res.event.proposedBy, A);
  assert.strictEqual(res.event.endDate, '2030-01-10');
  assert.match(res.event.log, /老公 新增$/);
  const row = gas.spreadsheet.getSheetByName('活動').cells[1];
  assert.strictEqual(row[1], '2030-01-10');
  assert.strictEqual(row[3], '18:30');
});

test('欄位檢查：沒標題、結束早於開始、參加亂填', () => {
  const gas = freshGas();
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { title: ' ' }) }).ok, false);
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { endDate: '2030-01-09' }) }).ok, false);
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { endTime: '17:00' }) }).ok, false);
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { startTime: '', endTime: '19:00' }) }).ok, false);
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { who: '隔壁' }) }).ok, false);
  assert.strictEqual(call(gas, A, 'save', { event: Object.assign({}, base, { who: B }) }).ok, true);
});

test('清單：對方提的出現在我的待回覆；我提的在等待中', () => {
  const gas = freshGas();
  call(gas, A, 'save', { event: base });
  const forB = call(gas, B, 'list', { from: '2030-01-01', to: '2030-01-31' });
  assert.strictEqual(forB.events.length, 1);
  assert.strictEqual(forB.todo.length, 1);
  assert.strictEqual(forB.waiting.length, 0);
  const forA = call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' });
  assert.strictEqual(forA.todo.length, 0);
  assert.strictEqual(forA.waiting.length, 1);
  const outside = call(gas, A, 'list', { from: '2030-02-01', to: '2030-02-28' });
  assert.strictEqual(outside.events.length, 0);
  assert.strictEqual(outside.waiting.length, 1);
});

test('多天活動跨進查詢區間也查得到', () => {
  const gas = freshGas();
  call(gas, A, 'save', { event: Object.assign({}, base, { startDate: '2030-01-30', endDate: '2030-02-02', startTime: '', endTime: '' }) });
  const feb = call(gas, A, 'list', { from: '2030-02-01', to: '2030-02-28' });
  assert.strictEqual(feb.events.length, 1);
});

test('自己不能確認自己提的；對方確認後狀態與回覆人正確', () => {
  const gas = freshGas();
  const ev = call(gas, A, 'save', { event: base }).event;
  const self = call(gas, A, 'respond', { id: ev.id, status: '已確認', baseUpdatedAt: ev.updatedAt });
  assert.strictEqual(self.ok, false);
  const res = call(gas, B, 'respond', { id: ev.id, status: '已確認', message: '好', baseUpdatedAt: ev.updatedAt });
  assert.ok(res.ok, res.error);
  assert.strictEqual(res.event.status, '已確認');
  assert.strictEqual(res.event.repliedBy, B);
  assert.match(res.event.log, /老婆 確認 ✓：好$/);
  assert.strictEqual(call(gas, B, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 0);
});

test('再議後輪到提出人處理；維持原案再回到對方', () => {
  const gas = freshGas();
  let ev = call(gas, A, 'save', { event: base }).event;
  ev = call(gas, B, 'respond', { id: ev.id, status: '再議', message: '那天加班', baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '再議');
  assert.strictEqual(call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 1);
  assert.strictEqual(call(gas, B, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 0);
  ev = call(gas, A, 'repropose', { id: ev.id, baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '待確認');
  assert.strictEqual(ev.proposedBy, A);
  assert.strictEqual(call(gas, B, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 1);
});

test('已確認後改時間：變回待確認、換對方確認；只改備註不影響', () => {
  const gas = freshGas();
  let ev = call(gas, A, 'save', { event: base }).event;
  ev = call(gas, B, 'respond', { id: ev.id, status: '已確認', baseUpdatedAt: ev.updatedAt }).event;
  ev = call(gas, B, 'save', { event: Object.assign({}, ev, { note: '記得帶禮物' }), baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '已確認');
  assert.strictEqual(ev.proposedBy, A);
  assert.match(ev.log, /老婆 修改 備註$/);
  ev = call(gas, B, 'save', { event: Object.assign({}, ev, { startTime: '19:00' }), baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '待確認');
  assert.strictEqual(ev.proposedBy, B);
  assert.strictEqual(ev.repliedBy, '');
  assert.match(ev.log, /時間 18:30-20:00→19:00-20:00$/);
  assert.strictEqual(call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 1);
});

test('沒改東西就不寫入', () => {
  const gas = freshGas();
  const ev = call(gas, A, 'save', { event: base }).event;
  const res = call(gas, A, 'save', { event: Object.assign({}, ev), baseUpdatedAt: ev.updatedAt });
  assert.strictEqual(res.event.log.split('\n').length, 1);
});

test('畫面資料過期（對方剛改過）時不讓存', () => {
  const gas = freshGas();
  const ev = call(gas, A, 'save', { event: base }).event;
  const sheet = gas.spreadsheet.getSheetByName('活動');
  const col = sheet.cells[0].indexOf('修改時間');
  sheet.cells[1][col] = '2000-01-01 00:00:00';
  const res = call(gas, B, 'respond', { id: ev.id, status: '已確認', baseUpdatedAt: ev.updatedAt });
  assert.strictEqual(res.code, 'STALE');
});

test('留言不改狀態也不改修改時間', () => {
  const gas = freshGas();
  const ev = call(gas, A, 'save', { event: base }).event;
  const res = call(gas, B, 'comment', { id: ev.id, message: '幾點出門？' }).event;
  assert.strictEqual(res.status, '待確認');
  assert.strictEqual(res.updatedAt, ev.updatedAt);
  assert.match(res.log, /老婆 留言：幾點出門？$/);
  assert.strictEqual(call(gas, B, 'comment', { id: ev.id, message: '  ' }).ok, false);
});

test('取消後不能確認；恢復後回到待確認', () => {
  const gas = freshGas();
  let ev = call(gas, A, 'save', { event: base }).event;
  ev = call(gas, B, 'cancel', { id: ev.id, message: '下雨', baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '取消');
  assert.strictEqual(call(gas, A, 'respond', { id: ev.id, status: '已確認', baseUpdatedAt: ev.updatedAt }).ok, false);
  assert.strictEqual(call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' }).todo.length, 0);
  ev = call(gas, A, 'repropose', { id: ev.id, baseUpdatedAt: ev.updatedAt }).event;
  assert.strictEqual(ev.status, '待確認');
  assert.match(ev.log, /老公 恢復活動$/);
});

test('試算表手動打的日期時間格式也讀得懂', () => {
  const gas = freshGas();
  const sheet = gas.spreadsheet.getSheetByName('活動');
  const header = sheet.cells[0];
  const row = header.map(() => '');
  row[header.indexOf('ID')] = 'manual1';
  row[header.indexOf('開始日期')] = '2030/1/5';
  row[header.indexOf('開始時間')] = '下午 6:30:00';
  row[header.indexOf('標題')] = '手打的';
  sheet.cells[1] = row;
  const res = call(gas, A, 'list', { from: '2030-01-01', to: '2030-01-31' });
  assert.strictEqual(res.events[0].startDate, '2030-01-05');
  assert.strictEqual(res.events[0].endDate, '2030-01-05');
  assert.strictEqual(res.events[0].startTime, '18:30');
  assert.strictEqual(res.events[0].status, '待確認');
});

test('使用者自己加的欄位，存檔時不會被清掉', () => {
  const gas = freshGas();
  const sheet = gas.spreadsheet.getSheetByName('活動');
  sheet.cells[0].push('我的欄位');
  const ev = call(gas, A, 'save', { event: base }).event;
  sheet.cells[1][17] = '保留我';
  call(gas, B, 'respond', { id: ev.id, status: '已確認', baseUpdatedAt: ev.updatedAt });
  assert.strictEqual(sheet.cells[1][17], '保留我');
});

console.log('\n全部通過：' + passed + ' 項');
