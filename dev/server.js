// 本機測試伺服器：node dev/server.js，然後開 http://localhost:8787/?api=/exec
// 後端跑的是真的 gas/Code.js，但試算表是記憶體裡的假資料，關掉就沒了。
const http = require('http');
const fs = require('fs');
const path = require('path');
const { createGas } = require('./fake-gas');

const PORT = Number(process.env.PORT || 8787);
const gas = createGas();
gas.context.setup();
const settingRows = gas.spreadsheet.getSheetByName('設定').cells;
const key = settingRows.find((r) => r[0] === '鑰匙')[1];

function pad(n) { return String(n).padStart(2, '0'); }
function dayFromToday(offset) {
  const d = new Date();
  d.setDate(d.getDate() + offset);
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}
function call(who, action, data) {
  const res = gas.post(Object.assign({ action, who, key }, data));
  if (!res.ok) throw new Error(res.error);
  return res.event;
}

// 範例資料：各種狀態都有一筆
const A = '老公';
const B = '老婆';
call(A, 'save', { event: { title: '一起吃晚餐', startDate: dayFromToday(2), startTime: '18:30', endTime: '20:00', location: '老地方', who: '兩人' } });
let ev = call(B, 'save', { event: { title: '回娘家', startDate: dayFromToday(5), endDate: dayFromToday(6), who: '兩人', note: '要帶水果' } });
call(A, 'respond', { id: ev.id, status: '已確認', message: '好', baseUpdatedAt: ev.updatedAt });
ev = call(A, 'save', { event: { title: '公司聚餐', startDate: dayFromToday(1), startTime: '19:00', who: A } });
call(B, 'respond', { id: ev.id, status: '再議', message: '那天我要用車', baseUpdatedAt: ev.updatedAt });
call(B, 'save', { event: { title: '瑜珈課', startDate: dayFromToday(3), startTime: '10:00', endTime: '11:30', who: B } });
ev = call(A, 'save', { event: { title: '看電影', startDate: dayFromToday(-2), startTime: '20:00', who: '兩人' } });
call(B, 'cancel', { id: ev.id, message: '太累了', baseUpdatedAt: ev.updatedAt });
call(B, 'save', { event: { title: '牙醫', startDate: dayFromToday(12), startTime: '15:00', who: B } });
call(A, 'save', { event: { title: '大掃除', startDate: dayFromToday(4), period: '上午', who: '兩人' } });
call(B, 'save', { event: { title: '出國', startDate: dayFromToday(60), endDate: dayFromToday(64), who: '兩人' } });
function task(who, data) {
  const res = gas.post(Object.assign({ action: 'taskSave', who, key }, { task: data }));
  if (!res.ok) throw new Error(res.error);
  return res.task;
}
task(B, { text: '買衛生紙', who: A, dueDate: dayFromToday(1) });
task(A, { text: '繳電費', who: '兩人', dueDate: dayFromToday(-1) });
task(A, { text: '整理冬天的衣服', who: B });
const doneTask = task(B, { text: '訂餐廳', who: A });
gas.post({ action: 'taskDone', who: A, key, id: doneTask.id, done: true });

http.createServer((req, res) => {
  if (req.method === 'POST' && req.url.startsWith('/exec')) {
    let body = '';
    req.on('data', (chunk) => { body += chunk; });
    req.on('end', () => {
      // 模擬 GAS 的延遲（正式環境一次大約 1.5～3 秒）
      setTimeout(() => {
        const out = gas.context.doPost({ postData: { contents: body } });
        res.writeHead(200, { 'Content-Type': 'application/json; charset=utf-8' });
        res.end(out.text);
      }, Number(process.env.LATENCY || 1500));
    });
    return;
  }
  if (req.url === '/' || req.url.startsWith('/?') || req.url.startsWith('/index.html')) {
    res.writeHead(200, { 'Content-Type': 'text/html; charset=utf-8' });
    res.end(fs.readFileSync(path.join(__dirname, '..', 'index.html')));
    return;
  }
  res.writeHead(404);
  res.end('not found');
}).listen(PORT, () => {
  console.log('http://localhost:' + PORT + '/?api=/exec#k=' + key);
});
