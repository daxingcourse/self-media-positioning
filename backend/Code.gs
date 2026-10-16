/**
 * 奶茶定位所｜後台（Google 試算表 + Apps Script + Resend）
 *
 * 功能：
 *  1. 接收測驗網站送來的 Email 與測驗結果 → 存進「名單」分頁 → 立刻寄出完整報告
 *  2. 每小時自動檢查「名單」，依「信件序列」分頁的設定寄出後續信件
 *  3. 「產品」分頁放你的課程／產品，會自動推薦給適合的奶茶人格（網站結果頁＋信件裡）
 *  4. 每封信都有取消訂閱連結
 *  5. 網頁版後台 admin.html：看數據、管理名單、編輯序列信、產品、設定（用後台密碼登入）
 *
 * 寄信：預設用 Resend（需要自己的網域）；也可以在「設定」把「寄信方式」改成 gmail。
 * 金鑰與後台密碼存在「指令碼屬性」，不會出現在試算表或網站上。
 *
 * 安裝步驟請看 backend/後台設定教學.md
 */

const SHEETS = {
  LEADS: '名單',
  SEQ: '信件序列',
  PRODUCTS: '產品',
  SETTINGS: '設定',
  DASH: '儀表板',
};

const LEAD_HEADERS = [
  '建立時間', 'Email', '暱稱', '奶茶人格', '副口味', '賽道1', '賽道2', '賽道3',
  '受眾', '目標', '露臉', '每週時間',
  '知識力', '表達力', '美感力', '共感力', '行動力',
  '定位句', '推薦產品ID', '已寄序列', '下次寄送', '狀態', '退訂碼', '最後更新', '來源',
];
const SEQ_HEADERS = ['順序', '第幾天寄', '主旨', '內文', '啟用'];
const PRODUCT_HEADERS = ['產品ID', '名稱', '一句話介紹', '價格', '連結', '按鈕文字', '適合人格', '優先序', '啟用', '在結果頁顯示'];
const SETTING_KEYS = ['寄信方式', '寄件人名稱', '寄件信箱', '回覆信箱', '網站網址', '報告信主旨', '每次最多寄送', '寄送時段開始', '寄送時段結束', '時區'];

const PERSONA_TITLES = ['知識導師', '人氣主持人', '美感創作者', '療癒陪伴者', '實驗紀錄家'];

const STATUS = { ACTIVE: '訂閱中', UNSUB: '已退訂', DONE: '序列完成' };

const PROP = { RESEND: 'RESEND_API_KEY', ADMIN: 'ADMIN_KEY' };

/* =========================================================
   選單
   ========================================================= */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🧋 奶茶定位所')
    .addItem('① 初始化後台（第一次使用）', 'setup')
    .addItem('② 設定 Resend API 金鑰', 'promptResendKey')
    .addItem('③ 設定網頁後台密碼', 'promptAdminKey')
    .addSeparator()
    .addItem('寄測試信給自己', 'sendTestToMe')
    .addItem('立即處理寄信佇列', 'processQueue')
    .addToUi();
}

function promptResendKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('設定 Resend API 金鑰', '貼上 Resend 後台建立的 API Key（re_ 開頭）：', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (!/^re_/.test(key)) { ui.alert('格式不對，Resend 的金鑰是 re_ 開頭喔'); return; }
  PropertiesService.getScriptProperties().setProperty(PROP.RESEND, key);
  ui.alert('✅ 已儲存 Resend 金鑰');
}

function promptAdminKey() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('設定網頁後台密碼', '登入 admin.html 用的密碼（至少 8 個字）：', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const key = r.getResponseText().trim();
  if (key.length < 8) { ui.alert('密碼至少要 8 個字'); return; }
  PropertiesService.getScriptProperties().setProperty(PROP.ADMIN, key);
  ui.alert('✅ 已儲存後台密碼');
}

/* =========================================================
   初始化：建立分頁、表頭、範例資料、排程
   ========================================================= */
function setup() {
  const ss = SpreadsheetApp.getActive();

  const leads = ensureSheet_(ss, SHEETS.LEADS, LEAD_HEADERS);
  leads.setFrozenRows(1);
  leads.getRange('A:A').setNumberFormat('yyyy/mm/dd hh:mm');
  leads.getRange('U:U').setNumberFormat('yyyy/mm/dd hh:mm');
  leads.getRange('X:X').setNumberFormat('yyyy/mm/dd hh:mm');

  const seq = ensureSheet_(ss, SHEETS.SEQ, SEQ_HEADERS);
  if (seq.getLastRow() < 2) seq.getRange(2, 1, SAMPLE_SEQUENCE.length, SEQ_HEADERS.length).setValues(SAMPLE_SEQUENCE);
  seq.setFrozenRows(1);
  seq.setColumnWidth(3, 260);
  seq.setColumnWidth(4, 520);
  seq.getRange('D:D').setWrap(true);
  seq.getRange('E2:E200').insertCheckboxes();

  const products = ensureSheet_(ss, SHEETS.PRODUCTS, PRODUCT_HEADERS);
  if (products.getLastRow() < 2) products.getRange(2, 1, SAMPLE_PRODUCTS.length, PRODUCT_HEADERS.length).setValues(SAMPLE_PRODUCTS);
  products.setFrozenRows(1);
  products.getRange('I2:J200').insertCheckboxes();

  const settings = ensureSheet_(ss, SHEETS.SETTINGS, ['項目', '值', '說明']);
  const have = settings.getDataRange().getValues().map(r => r[0]);
  SAMPLE_SETTINGS.filter(r => have.indexOf(r[0]) < 0).forEach(r => settings.appendRow(r));
  settings.setColumnWidth(1, 160); settings.setColumnWidth(2, 320); settings.setColumnWidth(3, 380);

  buildDashboard_(ss);
  installTrigger_();

  ss.toast('後台建立完成！下一步：設定 Resend 金鑰與後台密碼（選單 ②③）', '🧋 奶茶定位所', 8);
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#EBDCC8').setFontColor('#4A3226');
  }
  return sh;
}

function buildDashboard_(ss) {
  let sh = ss.getSheetByName(SHEETS.DASH);
  if (!sh) sh = ss.insertSheet(SHEETS.DASH, 0);
  sh.clear();
  const L = `'${SHEETS.LEADS}'`;
  const rows = [
    ['🧋 奶茶定位所 儀表板', ''],
    ['', ''],
    ['總名單數', `=MAX(0,COUNTA(${L}!B:B)-1)`],
    ['訂閱中', `=COUNTIF(${L}!V:V,"${STATUS.ACTIVE}")`],
    ['已退訂', `=COUNTIF(${L}!V:V,"${STATUS.UNSUB}")`],
    ['序列已寄完', `=COUNTIF(${L}!V:V,"${STATUS.DONE}")`],
    ['近 7 天新增', `=COUNTIF(${L}!A:A,">="&(TODAY()-7))`],
    ['', ''],
    ['奶茶人格分布', '人數'],
  ].concat(PERSONA_TITLES.map(t => [t, `=COUNTIF(${L}!D:D,"*${t}*")`]));
  sh.getRange(1, 1, rows.length, 2).setValues(rows);
  sh.getRange('A1').setFontSize(16).setFontWeight('bold').setFontColor('#8A5E3F');
  sh.getRange('A3:A7').setFontWeight('bold');
  sh.getRange('A9:B9').setFontWeight('bold').setBackground('#EBDCC8');
  sh.setColumnWidth(1, 180);
}

function installTrigger_() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processQueue')
    .forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('processQueue').timeBased().everyHours(1).create();
}

/* =========================================================
   Web App 入口
   ========================================================= */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'products') return json_({ ok: true, products: listProducts_(true) });
  if (p.action === 'unsub') return unsubscribe_(p.c);
  return json_({ ok: true, service: '奶茶定位所' });
}

function doPost(e) {
  let data;
  try { data = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: '資料格式錯誤' }); }
  if (data.admin) {
    try { return json_(handleAdmin_(data)); }
    catch (err) { console.error(err); return json_({ ok: false, error: String(err.message || err) }); }
  }
  return handleSignup_(data);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* =========================================================
   收集名單：測驗者在結果頁送出 Email
   ========================================================= */
function handleSignup_(data) {
  if (data.hp) return json_({ ok: true });                       // 機器人陷阱欄位有填 → 假裝成功
  const email = String(data.email || '').trim().toLowerCase();
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email) || email.length > 120) return json_({ ok: false, error: 'Email 格式不正確' });
  if (!data.consent) return json_({ ok: false, error: '需要勾選同意才能寄送' });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let lead;
  try {
    lead = upsertLead_(email, data);
  } finally {
    lock.releaseLock();
  }
  try {
    sendReport_(lead, data.report || {});
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: '名單已收到，但報告寄送失敗，請稍後再試' });
  }
  return json_({ ok: true });
}

function upsertLead_(email, d) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LEADS);
  const { rows, col } = readTable_(sh);
  const now = new Date();
  const sc = d.scores || {};
  const product = pickProduct_(d.persona || '');
  const fields = {
    'Email': email,
    '暱稱': clip_(d.nick, 30),
    '奶茶人格': clip_(d.personaFull, 40),
    '副口味': clip_(d.second, 40),
    '賽道1': clip_((d.niches || [])[0], 20),
    '賽道2': clip_((d.niches || [])[1], 20),
    '賽道3': clip_((d.niches || [])[2], 20),
    '受眾': clip_((d.audience || []).join('、'), 60),
    '目標': clip_(d.goal, 20),
    '露臉': clip_(d.face, 20),
    '每週時間': clip_(d.time, 20),
    '知識力': Number(sc.know) || '',
    '表達力': Number(sc.talk) || '',
    '美感力': Number(sc.create) || '',
    '共感力': Number(sc.care) || '',
    '行動力': Number(sc.act) || '',
    '定位句': clip_(d.statement, 300),
    '推薦產品ID': product ? product['產品ID'] : '',
    '最後更新': now,
    '來源': clip_(d.source, 200),
  };

  const idx = rows.findIndex(r => String(r['Email']).toLowerCase() === email);
  if (idx >= 0) {
    const rowNum = idx + 2;
    const cur = rows[idx];
    if (cur['狀態'] === STATUS.UNSUB) {        // 退訂後又主動重新勾選同意 → 恢復訂閱，從原進度繼續
      fields['狀態'] = STATUS.ACTIVE;
      fields['下次寄送'] = now;
    }
    Object.keys(fields).forEach(k => sh.getRange(rowNum, col[k] + 1).setValue(fields[k]));
    return Object.assign({}, cur, fields, { _row: rowNum });
  }

  const seq = listSequence_();
  const lead = Object.assign(fields, {
    '建立時間': now,
    '已寄序列': 0,
    '下次寄送': seq.length ? addDays_(now, seq[0]['第幾天寄']) : '',
    '狀態': seq.length ? STATUS.ACTIVE : STATUS.DONE,
    '退訂碼': Utilities.getUuid().replace(/-/g, ''),
  });
  sh.appendRow(LEAD_HEADERS.map(h => lead[h] === undefined ? '' : lead[h]));
  lead._row = sh.getLastRow();
  return lead;
}

function unsubscribe_(code) {
  const page = msg => HtmlService.createHtmlOutput(
    `<div style="font-family:sans-serif;text-align:center;padding:60px 20px;color:#4A3226;background:#FBF5EC;min-height:100vh">
      <div style="font-size:56px">🧋</div><h2>${msg}</h2></div>`).setTitle('奶茶定位所');
  if (!code) return page('連結不完整');
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LEADS);
  const { rows, col } = readTable_(sh);
  const idx = rows.findIndex(r => r['退訂碼'] === code);
  if (idx < 0) return page('找不到這筆訂閱');
  sh.getRange(idx + 2, col['狀態'] + 1).setValue(STATUS.UNSUB);
  return page('已取消訂閱，之後不會再收到信件。謝謝你陪我們喝了這杯奶茶 💛');
}

/* =========================================================
   寄信：Resend（預設）或 Gmail
   ========================================================= */
function sendMail_(opt, settings) {
  settings = settings || getSettings_();
  const key = PropertiesService.getScriptProperties().getProperty(PROP.RESEND);
  if (settings.provider === 'resend') {
    if (!key) throw new Error('還沒設定 Resend API 金鑰');
    if (!settings.fromEmail) throw new Error('還沒在「設定」填寫寄件信箱（要用你在 Resend 驗證過的網域）');
    const payload = {
      from: `${settings.senderName} <${settings.fromEmail}>`,
      to: [opt.to],
      subject: opt.subject,
      html: opt.html,
    };
    if (settings.replyTo) payload.reply_to = settings.replyTo;
    if (opt.unsubUrl) payload.headers = { 'List-Unsubscribe': `<${opt.unsubUrl}>` };
    const res = UrlFetchApp.fetch('https://api.resend.com/emails', {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
    const code = res.getResponseCode();
    if (code >= 300) {
      let msg = res.getContentText();
      try { msg = JSON.parse(msg).message || msg; } catch (e) {}
      const err = new Error(`Resend 寄送失敗（${code}）：${msg}`);
      err.rateLimited = code === 429;
      throw err;
    }
    return;
  }
  MailApp.sendEmail({
    to: opt.to, subject: opt.subject, htmlBody: opt.html,
    name: settings.senderName, replyTo: settings.replyTo || undefined,
  });
}

function unsubUrl_(lead) {
  return ScriptApp.getService().getUrl() + '?action=unsub&c=' + encodeURIComponent(lead['退訂碼'] || '');
}

/* =========================================================
   排程寄信：每小時跑一次
   ========================================================= */
function processQueue() {
  const settings = getSettings_();
  const hour = Number(Utilities.formatDate(new Date(), settings.timezone, 'H'));
  if (hour < settings.sendFrom || hour >= settings.sendTo) return;   // 只在設定的時段寄

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.LEADS);
    const { rows, col } = readTable_(sh);
    const seq = listSequence_();
    const now = new Date();
    let sent = 0;
    for (let i = 0; i < rows.length; i++) {
      if (sent >= settings.maxPerRun) break;
      if (settings.provider !== 'resend' && MailApp.getRemainingDailyQuota() < 1) break;
      const r = rows[i], rowNum = i + 2;
      if (r['狀態'] !== STATUS.ACTIVE || !r['Email']) continue;
      const step = Number(r['已寄序列']) || 0;
      if (step >= seq.length) { sh.getRange(rowNum, col['狀態'] + 1).setValue(STATUS.DONE); continue; }
      const due = r['下次寄送'] ? new Date(r['下次寄送']) : now;
      if (due > now) continue;

      try {
        sendSequenceMail_(r, seq[step], settings);
        sent++;
      } catch (err) {
        console.error('寄送失敗', r['Email'], err);
        if (err.rateLimited) break;            // 今天的額度用完了，留到下次
        continue;
      }
      const nextStep = step + 1;
      sh.getRange(rowNum, col['已寄序列'] + 1).setValue(nextStep);
      if (nextStep >= seq.length) {
        sh.getRange(rowNum, col['狀態'] + 1).setValue(STATUS.DONE);
        sh.getRange(rowNum, col['下次寄送'] + 1).setValue('');
      } else {
        const base = r['建立時間'] ? new Date(r['建立時間']) : now;
        let next = addDays_(base, seq[nextStep]['第幾天寄']);
        if (next < now) next = addDays_(now, 0);
        sh.getRange(rowNum, col['下次寄送'] + 1).setValue(next);
      }
    }
  } finally {
    lock.releaseLock();
  }
}

function sendSequenceMail_(lead, mail, settings) {
  const product = findProduct_(lead['推薦產品ID']) || pickProduct_(lead['奶茶人格']);
  const vars = varsFor_(lead, product, settings);
  sendMail_({
    to: lead['Email'],
    subject: fill_(mail['主旨'], vars),
    html: layout_(renderBody_(mail['內文'], vars, product), lead, settings),
    unsubUrl: unsubUrl_(lead),
  }, settings);
}

/* =========================================================
   報告信（填完 Email 立即寄出）
   ========================================================= */
function sendReport_(lead, rep, settings) {
  settings = settings || getSettings_();
  const product = findProduct_(lead['推薦產品ID']) || pickProduct_(lead['奶茶人格']);
  const e = escapeHtml_;
  const list = arr => (arr || []).map(x => `<li style="margin:4px 0">${e(x)}</li>`).join('');
  const h = t => `<h3 style="color:#8A5E3F;font-size:17px;margin:26px 0 8px">${t}</h3>`;
  const scores = (rep.scores || []).map(s =>
    `<tr><td style="padding:4px 8px 4px 0;white-space:nowrap">${e(s.name)}</td>
      <td style="width:100%"><div style="background:#F5EADB;border-radius:99px;height:10px"><div style="background:#C08B5C;border-radius:99px;height:10px;width:${Math.max(0, Math.min(100, Number(s.value) || 0))}%"></div></div></td>
      <td style="padding-left:8px;color:#8E7766">${Number(s.value) || ''}</td></tr>`).join('');
  const dirs = (rep.directions || []).map((d, i) =>
    `<div style="border:2px solid ${i === 0 ? '#C08B5C' : '#E6D3BD'};border-radius:16px;padding:12px 14px;margin:8px 0;background:#FFFDF9">
      <b>${['🥇', '🥈', '🥉'][i] || ''} ${e(d.name)}</b> <span style="color:#C08B5C;font-weight:bold">契合 ${e(d.pct)}%</span>
      <div style="font-size:13px;color:#8E7766;margin-top:4px">內容支柱：${e((d.pillars || []).join('、'))}<br>推薦形式：${e(d.forms)}<br>變現方式：${e(d.monetize)}</div></div>`).join('');

  const body = `
    <div style="text-align:center">
      <div style="font-size:52px">🧋</div>
      <div style="color:#8E7766;font-size:13px">${e(lead['暱稱'] || '你')} 的奶茶人格</div>
      <h2 style="color:#8A5E3F;margin:6px 0">${e(lead['奶茶人格'])}</h2>
      <div style="color:#8E7766">${e(rep.slogan)}</div>
    </div>
    ${h('📊 五大創作者特質')}<table style="width:100%;border-collapse:collapse;font-size:14px">${scores}</table>
    ${h('💪 你的天賦')}<ul style="padding-left:20px">${list(rep.strengths)}</ul>
    ${h('🧊 小心這些冰塊')}<ul style="padding-left:20px">${list(rep.blindSpots)}</ul>
    ${h('🎯 最適合你的 3 個賽道')}${dirs}
    ${h('📜 你的一句話定位')}
    <div style="background:#4A3226;color:#FBF5EC;border-radius:16px;padding:16px;line-height:1.8">${e(rep.statement)}</div>
    ${h('📝 社群簡介範本')}
    <div style="border:2px dashed #D9BC9A;border-radius:14px;padding:12px 14px;white-space:pre-line">${e(rep.bio)}</div>
    ${h('📱 主力平台與頻率')}<ul style="padding-left:20px">${list(rep.platforms)}</ul>
    ${h('💡 前 6 篇選題靈感')}<ol style="padding-left:20px">${list(rep.ideas)}</ol>
    ${product ? h('🎁 推薦給你') + productCard_(product) : ''}
    <p style="margin-top:26px">接下來幾天，我會再寄幾封信給你，教你把這份定位變成真正的內容。記得留意信箱喔！</p>`;

  sendMail_({
    to: lead['Email'],
    subject: fill_(settings.reportSubject, varsFor_(lead, product, settings)),
    html: layout_(body, lead, settings),
    unsubUrl: unsubUrl_(lead),
  }, settings);
}

/* =========================================================
   產品
   ========================================================= */
function listProducts_(onlyShowOnSite) {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  if (!sh) return [];
  return readTable_(sh).rows
    .filter(p => p['產品ID'] && p['名稱'] && isTrue_(p['啟用']) && (!onlyShowOnSite || isTrue_(p['在結果頁顯示'])))
    .sort((a, b) => (Number(a['優先序']) || 99) - (Number(b['優先序']) || 99))
    .map(p => ({
      id: String(p['產品ID']), name: String(p['名稱']), desc: String(p['一句話介紹'] || ''),
      price: String(p['價格'] || ''), url: String(p['連結'] || ''), cta: String(p['按鈕文字'] || '了解更多'),
      personas: String(p['適合人格'] || '全部'),
    }));
}

function pickProduct_(persona) {
  const all = listProducts_(false);
  const fits = p => p.personas.split(/[,，、\s]+/).some(t => t === '全部' || (t && String(persona).indexOf(t) >= 0));
  const p = all.find(fits);
  return p ? toRow_(p) : null;
}

function findProduct_(id) {
  if (!id) return null;
  const p = listProducts_(false).find(x => x.id === String(id));
  return p ? toRow_(p) : null;
}

function toRow_(p) {
  return { '產品ID': p.id, '名稱': p.name, '一句話介紹': p.desc, '價格': p.price, '連結': p.url, '按鈕文字': p.cta };
}

function productCard_(p) {
  const e = escapeHtml_;
  return `<div style="border:2px solid #C08B5C;border-radius:18px;padding:16px;background:#F5EADB;margin:12px 0">
    <div style="font-size:17px;font-weight:bold;color:#8A5E3F">${e(p['名稱'])}</div>
    <div style="margin:6px 0">${e(p['一句話介紹'])}</div>
    ${p['價格'] ? `<div style="color:#8E7766;font-size:13px">${e(p['價格'])}</div>` : ''}
    ${p['連結'] ? `<a href="${e(p['連結'])}" style="display:inline-block;margin-top:10px;background:#C08B5C;color:#fff;text-decoration:none;padding:10px 22px;border-radius:99px;font-weight:bold">${e(p['按鈕文字'] || '了解更多')} →</a>` : ''}
  </div>`;
}

/* =========================================================
   信件版型與變數
   ========================================================= */
function varsFor_(lead, product, settings) {
  return {
    '暱稱': lead['暱稱'] || '你',
    '人格': lead['奶茶人格'] || '',
    '副口味': lead['副口味'] || '',
    '賽道1': lead['賽道1'] || '',
    '賽道2': lead['賽道2'] || '',
    '受眾': lead['受眾'] || '',
    '定位句': lead['定位句'] || '',
    '產品名稱': product ? product['名稱'] : '',
    '產品介紹': product ? product['一句話介紹'] : '',
    '產品價格': product ? product['價格'] : '',
    '產品連結': product ? product['連結'] : '',
    '網站網址': settings.siteUrl || '',
    '寄件人': settings.senderName || '',
  };
}

function fill_(text, vars) {
  return String(text || '').replace(/\{([^{}]+)\}/g, (m, k) => vars[k] !== undefined ? vars[k] : m);
}

/* 內文是純文字：空一行＝新段落；{課程卡片} 會換成產品卡片 */
function renderBody_(text, vars, product) {
  const parts = String(text || '').split('{課程卡片}');
  return parts.map(part => escapeHtml_(fill_(part, vars))
    .split(/\n{2,}/).map(p => p.trim() ? `<p style="margin:0 0 14px">${p.replace(/\n/g, '<br>')}</p>` : '').join(''))
    .join(product ? productCard_(product) : '');
}

function layout_(inner, lead, settings) {
  return `<div style="background:#FBF5EC;padding:24px 12px;font-family:'Noto Sans TC',-apple-system,'PingFang TC','Microsoft JhengHei',sans-serif;color:#3B2A20;line-height:1.7">
    <div style="max-width:560px;margin:0 auto;background:#FFFDF9;border:2px solid #E6D3BD;border-radius:24px;padding:26px 22px">
      ${inner}
      <p style="margin-top:24px">— ${escapeHtml_(settings.senderName)} 🧋</p>
    </div>
    <div style="max-width:560px;margin:14px auto 0;text-align:center;font-size:12px;color:#8E7766">
      你會收到這封信，是因為你在奶茶定位所留下 Email 並同意接收內容。<br>
      ${settings.siteUrl ? `<a href="${escapeHtml_(settings.siteUrl)}" style="color:#8A5E3F">重新測驗</a>　·　` : ''}<a href="${unsubUrl_(lead)}" style="color:#8A5E3F">取消訂閱</a>
    </div></div>`;
}

/* =========================================================
   網頁版後台 API（admin.html 用）
   ========================================================= */
function handleAdmin_(d) {
  const adminKey = PropertiesService.getScriptProperties().getProperty(PROP.ADMIN);
  if (!adminKey) return { ok: false, error: '還沒設定後台密碼（試算表選單 ③）' };
  if (String(d.key || '') !== adminKey) {
    Utilities.sleep(800);                       // 拖慢猜密碼
    return { ok: false, error: '密碼錯誤', auth: false };
  }
  const ss = SpreadsheetApp.getActive();
  switch (d.action) {
    case 'login':
      return { ok: true };

    case 'stats': {
      const { rows } = readTable_(ss.getSheetByName(SHEETS.LEADS));
      const weekAgo = Date.now() - 7 * 864e5;
      const persona = {}, niche = {}, days = {};
      rows.forEach(r => {
        const t = PERSONA_TITLES.find(p => String(r['奶茶人格']).indexOf(p) >= 0) || '其他';
        persona[t] = (persona[t] || 0) + 1;
        if (r['賽道1']) niche[r['賽道1']] = (niche[r['賽道1']] || 0) + 1;
        const ts = r['建立時間'] ? new Date(r['建立時間']) : null;
        if (ts && ts.getTime() >= Date.now() - 14 * 864e5) {
          const k = Utilities.formatDate(ts, getSettings_().timezone, 'MM/dd');
          days[k] = (days[k] || 0) + 1;
        }
      });
      const count = s => rows.filter(r => r['狀態'] === s).length;
      return {
        ok: true,
        total: rows.length,
        active: count(STATUS.ACTIVE), unsub: count(STATUS.UNSUB), done: count(STATUS.DONE),
        week: rows.filter(r => r['建立時間'] && new Date(r['建立時間']).getTime() >= weekAgo).length,
        persona, niche, days,
        mail: mailStatus_(),
      };
    }

    case 'leads': {
      const { rows } = readTable_(ss.getSheetByName(SHEETS.LEADS));
      return { ok: true, leads: rows.map(r => {
        const o = {};
        LEAD_HEADERS.filter(h => h !== '退訂碼').forEach(h => { o[h] = r[h] instanceof Date ? r[h].toISOString() : r[h]; });
        return o;
      }).reverse() };
    }

    case 'setLeadStatus': {
      if ([STATUS.ACTIVE, STATUS.UNSUB, STATUS.DONE].indexOf(d.status) < 0) return { ok: false, error: '狀態不正確' };
      const sh = ss.getSheetByName(SHEETS.LEADS);
      const { rows, col } = readTable_(sh);
      const idx = rows.findIndex(r => String(r['Email']).toLowerCase() === String(d.email).toLowerCase());
      if (idx < 0) return { ok: false, error: '找不到這筆名單' };
      sh.getRange(idx + 2, col['狀態'] + 1).setValue(d.status);
      if (d.status === STATUS.ACTIVE && d.restart) {
        sh.getRange(idx + 2, col['已寄序列'] + 1).setValue(0);
        sh.getRange(idx + 2, col['下次寄送'] + 1).setValue(new Date());
      }
      return { ok: true };
    }

    case 'getSequence':
      return { ok: true, rows: readTable_(ss.getSheetByName(SHEETS.SEQ)).rows };
    case 'saveSequence':
      writeTable_(ss.getSheetByName(SHEETS.SEQ), SEQ_HEADERS, (d.rows || []).map(r => ({
        '順序': Number(r['順序']) || 0, '第幾天寄': Number(r['第幾天寄']) || 0,
        '主旨': clip_(r['主旨'], 200), '內文': clip_(r['內文'], 20000), '啟用': !!r['啟用'],
      })));
      return { ok: true };

    case 'getProducts':
      return { ok: true, rows: readTable_(ss.getSheetByName(SHEETS.PRODUCTS)).rows };
    case 'saveProducts': {
      const ids = {};
      const rows = (d.rows || []).map(r => {
        const id = clip_(r['產品ID'], 40).trim();
        if (!id) throw new Error('每個產品都要有產品ID');
        if (ids[id]) throw new Error('產品ID 重複：' + id);
        ids[id] = 1;
        return {
          '產品ID': id, '名稱': clip_(r['名稱'], 100), '一句話介紹': clip_(r['一句話介紹'], 300),
          '價格': clip_(r['價格'], 40), '連結': clip_(r['連結'], 500), '按鈕文字': clip_(r['按鈕文字'], 20),
          '適合人格': clip_(r['適合人格'] || '全部', 100), '優先序': Number(r['優先序']) || 99,
          '啟用': !!r['啟用'], '在結果頁顯示': !!r['在結果頁顯示'],
        };
      });
      writeTable_(ss.getSheetByName(SHEETS.PRODUCTS), PRODUCT_HEADERS, rows);
      return { ok: true };
    }

    case 'getSettings': {
      const sh = ss.getSheetByName(SHEETS.SETTINGS);
      const map = {};
      sh.getDataRange().getValues().slice(1).forEach(r => { map[String(r[0]).trim()] = r[1]; });
      return { ok: true, settings: map, mail: mailStatus_() };
    }
    case 'saveSettings': {
      const sh = ss.getSheetByName(SHEETS.SETTINGS);
      const values = sh.getDataRange().getValues();
      Object.keys(d.settings || {}).forEach(k => {
        if (SETTING_KEYS.indexOf(k) < 0) return;
        const i = values.findIndex(r => String(r[0]).trim() === k);
        const v = clip_(d.settings[k], 300);
        if (i > 0) sh.getRange(i + 1, 2).setValue(v);
        else sh.appendRow([k, v, '']);
      });
      return { ok: true };
    }
    case 'setResendKey': {
      const k = String(d.resendKey || '').trim();
      if (!/^re_/.test(k)) return { ok: false, error: 'Resend 金鑰是 re_ 開頭' };
      PropertiesService.getScriptProperties().setProperty(PROP.RESEND, k);
      return { ok: true, mail: mailStatus_() };
    }

    case 'testSend': {
      const to = String(d.to || '').trim();
      if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(to)) return { ok: false, error: 'Email 格式不正確' };
      sendTest_(to, d.only);
      return { ok: true };
    }

    case 'runQueue':
      processQueue();
      return { ok: true };
  }
  return { ok: false, error: '未知的動作' };
}

function mailStatus_() {
  const s = getSettings_();
  const key = PropertiesService.getScriptProperties().getProperty(PROP.RESEND) || '';
  return {
    provider: s.provider,
    resendKey: key ? key.slice(0, 6) + '••••' + key.slice(-4) : '',
    fromEmail: s.fromEmail,
    ready: s.provider === 'resend' ? !!(key && s.fromEmail) : true,
  };
}

function writeTable_(sh, headers, objs) {
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2, 1, last - 1, Math.max(headers.length, sh.getLastColumn())).clearContent();
  if (objs.length) sh.getRange(2, 1, objs.length, headers.length).setValues(objs.map(o => headers.map(h => o[h] === undefined ? '' : o[h])));
}

/* =========================================================
   測試
   ========================================================= */
function sendTestToMe() {
  const me = Session.getActiveUser().getEmail();
  sendTest_(me);
  SpreadsheetApp.getActive().toast('已寄出測試信到 ' + me, '🧋 奶茶定位所', 6);
}

/* only：不填＝報告信＋全部序列信；'report'＝只寄報告信；數字＝只寄第幾封序列信 */
function sendTest_(to, only) {
  const settings = getSettings_();
  const lead = {
    'Email': to, '暱稱': '測試小粉圓', '奶茶人格': '伯爵奶茶・知識導師', '副口味': '黑糖珍珠鮮奶・人氣主持人',
    '賽道1': '職場成長', '賽道2': '科技 AI', '受眾': '上班族', '定位句': '我是測試小粉圓，一位知識導師型的職場成長創作者。',
    '退訂碼': 'test',
  };
  if (!only || only === 'report') {
    sendReport_(lead, {
      slogan: '把複雜的事，泡成一杯好入口的知識茶。',
      scores: [{ name: '知識力', value: 86 }, { name: '表達力', value: 70 }, { name: '美感力', value: 52 }, { name: '共感力', value: 48 }, { name: '行動力', value: 40 }],
      strengths: ['擅長拆解問題、整理成步驟與框架', '內容收藏率高'], blindSpots: ['容易寫得太硬太長'],
      directions: [{ name: '職場成長', pct: 88, pillars: ['職場技能', '溝通與人際', '轉職故事'], forms: '口播短影音', monetize: '線上課' }],
      statement: lead['定位句'], bio: '💼 測試小粉圓｜職場成長 × 知識導師', platforms: ['YouTube 中長片'], ideas: ['上班族最常踩的 3 個地雷'],
    }, settings);
  }
  listSequence_().forEach((m, i) => {
    if (only && only !== 'all' && Number(only) !== i + 1) return;
    sendSequenceMail_(lead, Object.assign({}, m, { '主旨': '[測試 第' + m['第幾天寄'] + '天] ' + m['主旨'] }), settings);
  });
}

/* =========================================================
   工具
   ========================================================= */
function readTable_(sh) {
  const values = sh.getDataRange().getValues();
  const headers = values.shift() || [];
  const col = {};
  headers.forEach((h, i) => { col[h] = i; });
  const rows = values.map(v => { const o = {}; headers.forEach((h, i) => { o[h] = v[i]; }); return o; });
  return { rows, col };
}

function listSequence_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SEQ);
  if (!sh) return [];
  return readTable_(sh).rows
    .filter(r => isTrue_(r['啟用']) && r['主旨'])
    .sort((a, b) => (Number(a['順序']) || 0) - (Number(b['順序']) || 0));
}

function getSettings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEETS.SETTINGS);
  const map = {};
  if (sh) sh.getDataRange().getValues().slice(1).forEach(r => { map[String(r[0]).trim()] = r[1]; });
  return {
    provider: String(map['寄信方式'] || 'resend').trim().toLowerCase() === 'gmail' ? 'gmail' : 'resend',
    senderName: String(map['寄件人名稱'] || '奶茶定位所'),
    fromEmail: String(map['寄件信箱'] || '').trim(),
    replyTo: String(map['回覆信箱'] || '').trim(),
    siteUrl: String(map['網站網址'] || ''),
    reportSubject: String(map['報告信主旨'] || '🧋 {暱稱}，你的自媒體定位報告來囉！'),
    maxPerRun: Number(map['每次最多寄送']) || 40,
    sendFrom: Number(map['寄送時段開始']) || 9,
    sendTo: Number(map['寄送時段結束']) || 21,
    timezone: String(map['時區'] || 'Asia/Taipei'),
  };
}

function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + (Number(days) || 0));
  return d;
}

function isTrue_(v) { return v === true || String(v).toUpperCase() === 'TRUE' || v === '是' || v === 'V'; }
function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function escapeHtml_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* =========================================================
   範例資料（第一次初始化時寫入，之後在網頁後台或試算表修改）
   ========================================================= */
const SAMPLE_SETTINGS = [
  ['寄信方式', 'resend', 'resend 或 gmail（Gmail 每天約 100 封、寄件者是你的 Gmail）'],
  ['寄件人名稱', '奶茶定位所', '收件人看到的寄件者名字，可改成你的名字或品牌'],
  ['寄件信箱', '', '用 Resend 時必填，例如 hello@你的網域.com（網域要先在 Resend 驗證）'],
  ['回覆信箱', '', '讀者按「回覆」時寄到哪裡，例如你的 Gmail'],
  ['網站網址', 'https://daxingcourse.github.io/self-media-positioning/', '信件底部「重新測驗」連結'],
  ['報告信主旨', '🧋 {暱稱}，你的自媒體定位報告來囉！', '可用變數：{暱稱} {人格}'],
  ['每次最多寄送', 40, '每小時排程最多寄幾封（Resend 免費版每天 100 封）'],
  ['寄送時段開始', 9, '幾點之後才寄（24 小時制）'],
  ['寄送時段結束', 21, '幾點之前寄完'],
  ['時區', 'Asia/Taipei', ''],
];

const SAMPLE_SEQUENCE = [
  [1, 1, '{暱稱}，你的「{人格}」天賦，要這樣用',
    '嗨 {暱稱}，\n\n昨天你測出來是「{人格}」，副口味是「{副口味}」。\n\n很多人做自媒體卡住，不是因為不夠努力，而是一直在用「別人的口味」做內容。你的天賦已經寫在報告裡了，今天只做一件事：\n\n👉 打開報告的「一句話定位」，把它改成你自己的話，貼在手機備忘錄最上面。\n\n之後每次不知道要發什麼，就回來看這一句。\n\n明天見！', true],
  [2, 3, '前 3 支「{賽道1}」內容，照這個順序做',
    '嗨 {暱稱}，\n\n新帳號最怕一開始就想做「爆款」。其實前 3 支內容只要完成三件事：\n\n1️⃣ 自我介紹：你是誰、為什麼做 {賽道1}\n2️⃣ 一支乾貨：解決 {受眾} 一個很小但很痛的問題\n3️⃣ 一支故事：你自己踩過的坑\n\n不用完美，先發出去。發完回信告訴我，我會看！', true],
  [3, 5, '為什麼你發了內容卻沒人看？',
    '嗨 {暱稱}，\n\n這是我最常被問的問題。通常原因只有三個：\n\n・開頭 3 秒沒有說「這跟你有什麼關係」\n・一次講太多重點\n・沒有固定發文節奏，演算法認不得你\n\n今天挑一支你發過的內容，只改開頭第一句，再發一次看看。', true],
  [4, 7, '{暱稱}，想更快上手嗎？',
    '嗨 {暱稱}，\n\n這一週你已經有了定位、知道前 3 支內容怎麼做。如果你想要有人帶著你、一步步把帳號做起來，我準備了這個：\n\n{課程卡片}\n\n有任何問題，直接回信給我就好 💛', true],
  [5, 10, '最後一封：給「{人格}」的你',
    '嗨 {暱稱}，\n\n這是這個系列的最後一封信。\n\n做自媒體最重要的不是天分，是「持續」。你的定位是：\n\n「{定位句}」\n\n只要照著這句話，每週穩定產出，三個月後你會感謝現在的自己。\n\n{課程卡片}\n\n我們內容裡見！', true],
];

const SAMPLE_PRODUCTS = [
  ['course-01', '（範例）自媒體 28 天實戰營', '從定位、選題到變現，28 天陪你做出第一個產品', 'NT$ 3,980', 'https://example.com/course', '立即了解', '全部', 1, true, true],
  ['ebook-01', '（範例）知識型創作者懶人包模板', '30 款圖文模板，讓乾貨一眼就想收藏', 'NT$ 490', 'https://example.com/ebook', '下載模板', '知識導師,美感創作者', 2, false, false],
];
