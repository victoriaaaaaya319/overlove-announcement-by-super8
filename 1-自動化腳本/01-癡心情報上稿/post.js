#!/usr/bin/env node
/**
 * 癡心情報上稿腳本
 *
 * 讀取 Excel 的一列 → 上傳圖片 → 在 Super 8 Studio 建立「群發訊息」草稿（排程 21:00）
 *
 * 用法：
 *   node post.js                    最後一列（主題有填的最後一列）
 *   node post.js 2026-09-10         指定發送日期那一列
 *   node post.js --dry-run          只檢查與顯示要送出的內容，不建立草稿
 *   node post.js --login            只做登入（存到 browser-profile），不上稿
 *
 * 登入：先用 browser-profile/ 裡保存的登入狀態；過期就用 .env 的
 *       SUPER8_EMAIL / SUPER8_PASSWORD 在背景自動登入；沒有 .env 才會開視窗請你手動登入。
 */

const fs = require('fs');
const path = require('path');
const XLSX = require('xlsx');
const { chromium } = require('playwright-core');

const ROOT = __dirname;
const config = JSON.parse(fs.readFileSync(path.join(ROOT, 'config.json'), 'utf8'));

// 組織 ID 不放進版本控制，改從 .env 的 SUPER8_ORG_ID 讀（見 .env.example）
config.orgId = (readEnv().SUPER8_ORG_ID || process.env.SUPER8_ORG_ID || '').trim();
if (!config.orgId) {
  console.error('缺少 SUPER8_ORG_ID。請在這個資料夾的 .env 裡加上一行：SUPER8_ORG_ID=你的組織ID');
  process.exit(1);
}

const PARSE_APP_ID = 'number8';
const PARSE_JS_KEY = 'javascriptKey';
const CURRENT_USER_KEY = `Parse/${PARSE_APP_ID}/currentUser`;

// ---------- 小工具 ----------

function die(msg) {
  console.error('\n❌ ' + msg + '\n');
  process.exit(1);
}

function log(msg) {
  console.log('▸ ' + msg);
}

function parseArgs(argv) {
  const out = { date: null, dryRun: false, loginOnly: false };
  for (const a of argv) {
    if (a === '--dry-run' || a === '--dry') out.dryRun = true;
    else if (a === '--login') out.loginOnly = true;
    else if (/^\d{4}[-/]\d{1,2}[-/]\d{1,2}$/.test(a)) out.date = a;
    else die(`看不懂的參數：${a}`);
  }
  return out;
}

/** 把 Excel 讀出來的日期（Date / 序號 / 字串）轉成 {y, m, d} */
function toYMD(v) {
  if (v instanceof Date) {
    // SheetJS cellDates 給的是「本地時間」的 Date，要用本地取值，不能用 UTC
    return { y: v.getFullYear(), m: v.getMonth() + 1, d: v.getDate() };
  }
  if (typeof v === 'number') {
    const p = XLSX.SSF.parse_date_code(v);
    return { y: p.y, m: p.m, d: p.d };
  }
  if (typeof v === 'string') {
    const m = v.trim().match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
    if (m) return { y: +m[1], m: +m[2], d: +m[3] };
  }
  return null;
}

const pad2 = (n) => String(n).padStart(2, '0');
const ymdDash = ({ y, m, d }) => `${y}-${pad2(m)}-${pad2(d)}`;
const ymdShort = ({ y, m, d }) => `${String(y).slice(2)}${pad2(m)}${pad2(d)}`;

function str(v) {
  return v == null ? '' : String(v).replace(/\r\n/g, '\n').trim();
}

function checkLen(label, value, max) {
  const n = [...value].length;
  if (n > max) die(`${label} 超過 ${max} 字（目前 ${n} 字）：${value}`);
}

// ---------- 讀 Excel ----------

function readRow(dateArg) {
  const file = path.join(ROOT, config.excelFile);
  if (!fs.existsSync(file)) die(`找不到 Excel：${file}`);

  const wb = XLSX.readFile(file, { cellDates: true });
  const ws = wb.Sheets[config.excelSheet];
  if (!ws) die(`Excel 裡沒有「${config.excelSheet}」這個工作表`);

  const rows = XLSX.utils.sheet_to_json(ws, { defval: '' });
  const filled = rows
    .map((r, i) => ({ r, excelRow: i + 2 }))
    .filter(({ r }) => str(r['主題']) !== '' || str(r['卡片標題']) !== '');

  if (filled.length === 0) die('Excel 裡沒有任何填了「主題」的資料列');

  const isDone = ({ r }) => str(r['狀態']) !== '';
  const pending = filled.filter((x) => !isDone(x));

  let pick;
  if (dateArg) {
    const want = ymdDash(toYMD(dateArg));
    pick = filled.find(({ r }) => {
      const d = toYMD(r['發送日期']);
      return d && ymdDash(d) === want;
    });
    if (!pick) die(`找不到發送日期 = ${want} 的那一列`);
    if (isDone(pick)) die(`第 ${pick.excelRow} 列（${want}）已經跑過了：「${str(pick.r['狀態'])}」。要重跑請先把「狀態」欄清空。`);
  } else {
    if (pending.length === 0) die('每一列的「狀態」都已填，沒有要跑的。新增一列或清空某列的「狀態」再跑。');
    pick = pending[pending.length - 1];
    if (pending.length > 1) {
      const others = pending.slice(0, -1).map((x) => `第 ${x.excelRow} 列`).join('、');
      console.warn(`⚠️  還有其他列尚未跑過（${others}），這次只跑最後一列（第 ${pick.excelRow} 列）。要跑別列請指定日期。`);
    }
  }

  const r = pick.r;
  const date = toYMD(r['發送日期']);
  if (!date) die(`第 ${pick.excelRow} 列的「發送日期」讀不出來：${r['發送日期']}`);

  const topic = str(r['主題']);
  const title = str(r['卡片標題']) || (config.titlePrefix + topic);
  const notify = str(r['推播通知']) || (config.notifyPrefix + topic);
  const subtitle = str(r['卡片敘述']);
  const imagePath = str(r['圖片路徑']);
  const url = str(r['電子報網址']);
  const text = str(r['文字訊息']);

  const missing = [];
  if (!topic && !title) missing.push('主題');
  if (!subtitle) missing.push('卡片敘述');
  if (!imagePath) missing.push('圖片路徑');
  if (!url) missing.push('電子報網址');
  if (!text) missing.push('文字訊息');
  if (missing.length) die(`第 ${pick.excelRow} 列缺少欄位：${missing.join('、')}`);

  if (!/^https?:\/\//.test(url)) die(`電子報網址要以 http 或 https 開頭：${url}`);
  if (!fs.existsSync(imagePath)) die(`找不到圖片：${imagePath}`);
  if (!/\.(png|jpe?g|gif)$/i.test(imagePath)) die(`圖片只支援 png / jpg / jpeg / gif：${imagePath}`);

  checkLen('卡片標題', title, config.limits.title);
  checkLen('卡片敘述', subtitle, config.limits.subtitle);
  checkLen('推播通知', notify, config.limits.altText);

  return { excelRow: pick.excelRow, date, topic, title, notify, subtitle, imagePath, url, text };
}

// ---------- 組 payload ----------

function buildSchedule(date) {
  const [hh, mm] = config.sendTime.split(':').map(Number);
  // 台灣時間 → UTC ISO 字串
  const iso = `${ymdDash(date)}T${pad2(hh)}:${pad2(mm)}:00+08:00`;
  const dt = new Date(iso);
  if (isNaN(dt)) die(`排程時間組不出來：${iso}`);
  return dt;
}

function buildPayload(row, imageUrl, customerNum) {
  const name = ymdShort(row.date) + config.nameSuffix;
  const scheduleAt = buildSchedule(row.date).toISOString();

  const buttons = config.buttons.map((b) => {
    if (b.type === 'url') return { type: 'url', title: b.title, data: row.url, tags: b.tags || [] };
    if (b.type === 'postback') return { type: 'postback', title: b.title, data: b.title, tags: b.tags || [] };
    die(`config.json 的按鈕類型不支援：${b.type}`);
  });

  return {
    name,
    orgId: config.orgId,
    query: {
      orgId: config.orgId,
      platforms: config.platforms,
      tagDensity: [{ type: 'include', count: config.audienceTagCount, tags: config.audienceTags }],
      partnerTag: [],
    },
    scheduleAt,
    quickReply: config.quickReply.map((q) => ({ text: q.text, label: q.label })),
    messages: [
      { contentType: 'text/plain', data: { content: row.text }, index: 1, version: 2 },
    ],
    templates: [
      {
        className: 'Template',
        templateType: 'card',
        contentType: 'application/x-template',
        data: {
          elements: [
            {
              title: row.title,
              subtitle: row.subtitle,
              imageType: 'upload',
              imageUrl,
              aspectRatio: config.cardImageAspectRatio,
              buttons,
            },
          ],
          templateType: 'card',
        },
        index: 0,
        altText: row.notify,
        version: 2,
      },
    ],
    options: {
      name,
      applyFacebook24Policy: true,
      applyLineUnfollowFilter: true,
      customerNum,
      isDraft: true,
    },
  };
}

// ---------- API ----------

async function parseUsersMe(token) {
  const res = await fetch(`${config.parseUrl}/users/me`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ _method: 'GET', _ApplicationId: PARSE_APP_ID, _JavaScriptKey: PARSE_JS_KEY, _SessionToken: token }),
  });
  if (!res.ok) return null;
  return res.json();
}

async function apiNext(pathname, token, body, method = 'POST') {
  const res = await fetch(`${config.apiNextUrl}${pathname}`, {
    method,
    headers: {
      Accept: 'application/json, text/plain, */*',
      'Content-Type': 'application/json',
      _SessionToken: token,
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  const txt = await res.text();
  let json = null;
  try { json = JSON.parse(txt); } catch (_) { /* 非 JSON */ }
  if (!res.ok) die(`API ${pathname} 失敗（HTTP ${res.status}）：${txt.slice(0, 500)}`);
  return json ?? txt;
}

async function uploadImage(imagePath) {
  const filename = path.basename(imagePath);
  const ext = path.extname(filename).toLowerCase();
  const mime = ext === '.png' ? 'image/png' : ext === '.gif' ? 'image/gif' : 'image/jpeg';

  const signRes = await fetch(`${config.uploaderUrl}/${encodeURIComponent(filename)}`, {
    headers: { Accept: 'application/json, text/plain, */*' },
  });
  if (!signRes.ok) die(`取得圖片上傳網址失敗（HTTP ${signRes.status}）`);
  const { url: putUrl, location } = await signRes.json();
  if (!putUrl || !location) die('圖片上傳網址回應格式不對：' + JSON.stringify(await signRes.text()).slice(0, 300));

  const buf = fs.readFileSync(imagePath);
  const putRes = await fetch(putUrl, { method: 'PUT', headers: { 'Content-Type': mime }, body: buf });
  if (!putRes.ok) die(`圖片上傳到 S3 失敗（HTTP ${putRes.status}）`);

  return location;
}

async function getCustomerCount(token) {
  const body = {
    where: {
      orgId: config.orgId,
      platforms: config.platforms,
      customerObjectIds: [],
      tagDensity: [{ type: 'include', count: config.audienceTagCount, tags: config.audienceTags }],
      partnerTag: [],
      gender: [],
    },
    options: { applyFacebook24Policy: true, applyLineUnfollowFilter: true },
    returnCount: true,
    returnPipeline: false,
    connControl: true,
    sliceTagsCount: 1,
  };
  const res = await apiNext('/broadcast/getCustomers', token, body);
  const count = res && res.result && typeof res.result.count === 'number' ? res.result.count : null;
  if (count == null) die('計算人數的回應看不懂：' + JSON.stringify(res).slice(0, 300));
  return count;
}

// ---------- 登入狀態（用瀏覽器 profile 保存） ----------

async function launch(headless) {
  const profileDir = path.join(ROOT, config.profileDir);
  const opts = { headless, viewport: { width: 1280, height: 800 } };
  // 優先用電腦上已裝好的 Edge / Chrome，不用另外下載瀏覽器
  for (const channel of ['msedge', 'chrome']) {
    try {
      return await chromium.launchPersistentContext(profileDir, { ...opts, channel });
    } catch (e) {
      if (!/Chromium distribution|not found|executable/i.test(String(e))) throw e;
    }
  }
  die('找不到 Edge 或 Chrome。請安裝其中一個，或執行 npx playwright install chromium 後修改 launch()。');
}

async function readTokenFromPage(page) {
  const raw = await page.evaluate((k) => localStorage.getItem(k), CURRENT_USER_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw).sessionToken || null; } catch (_) { return null; }
}

/** 讀 .env（SUPER8_ORG_ID / SUPER8_EMAIL / SUPER8_PASSWORD） */
function readEnv() {
  const file = path.join(ROOT, '.env');
  if (!fs.existsSync(file)) return {};
  const out = {};
  for (const line of fs.readFileSync(file, 'utf8').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!m || line.trim().startsWith('#')) continue;
    let v = m[2];
    if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
    out[m[1]] = v;
  }
  return out;
}

/** 等 localStorage 出現有效 token，最多 waitMs 毫秒 */
async function waitForToken(page, waitMs) {
  const deadline = Date.now() + waitMs;
  while (Date.now() < deadline) {
    await page.waitForTimeout(1500);
    if (page.isClosed()) return null;
    const t = await readTokenFromPage(page).catch(() => null);
    if (t && (await parseUsersMe(t))) return t;
  }
  return null;
}

/** 用 .env 的帳密在登入頁自動填表 */
async function loginWithEnv(page, env) {
  // 「系統已推出新版本」對話框擋在前面的話先按「稍後」
  const later = page.getByRole('button', { name: '稍後' });
  if (await later.isVisible().catch(() => false)) await later.click().catch(() => {});

  const email = page.locator('input[name="email"]');
  const password = page.locator('input[name="password"]');
  await email.waitFor({ state: 'visible', timeout: 20000 });
  await email.fill(env.SUPER8_EMAIL);
  await password.fill(env.SUPER8_PASSWORD);
  await page.locator('button[type="submit"]').click();

  const token = await waitForToken(page, 60 * 1000);
  if (!token) {
    // 抓一下畫面上的錯誤訊息（帳密錯、被鎖等）
    const msg = await page.locator('main, body').innerText().then((t) => t.slice(0, 300)).catch(() => '');
    log('自動登入沒成功，頁面內容：' + msg.replace(/\s+/g, ' '));
  }
  return token;
}

/** 回傳有效的 sessionToken。順序：保存的登入狀態 → .env 自動登入 → 開視窗手動登入 */
async function getSessionToken({ forceWindow = false } = {}) {
  const env = readEnv();
  const hasEnv = !!(env.SUPER8_EMAIL && env.SUPER8_PASSWORD);

  // 1. 先用保存的登入狀態（無頭）
  let ctx = await launch(!forceWindow);
  let page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(config.consoleUrl, { waitUntil: 'domcontentloaded' });

  let token = await readTokenFromPage(page);
  if (token && (await parseUsersMe(token))) {
    await ctx.close();
    return token;
  }
  token = null;

  // 2. 有 .env 就自動填帳密登入（同一個無頭視窗）
  if (hasEnv) {
    log('登入狀態過期或不存在，用 .env 的帳號自動登入…');
    token = await loginWithEnv(page, env).catch((e) => { log('自動登入出錯：' + e.message); return null; });
    if (token) {
      log('自動登入成功，登入狀態已保存到 ' + config.profileDir + '/');
      await ctx.close();
      return token;
    }
  } else {
    log('沒有 .env（或缺 SUPER8_EMAIL / SUPER8_PASSWORD），改為開視窗手動登入。');
  }

  // 3. 開視窗請使用者手動登入
  if (!forceWindow) {
    await ctx.close();
    ctx = await launch(false);
    page = ctx.pages()[0] || (await ctx.newPage());
    await page.goto(config.consoleUrl, { waitUntil: 'domcontentloaded' });
  }
  console.log('\n🔐 請在剛開啟的瀏覽器視窗登入 Super 8 Studio。登入完成後這裡會自動繼續。\n');
  token = await waitForToken(page, 10 * 60 * 1000);
  if (!token) die('等了 10 分鐘還沒登入（或視窗被關掉），先結束。');
  log('登入成功，登入狀態已保存到 ' + config.profileDir + '/');

  await ctx.close();
  return token;
}

// ---------- 主流程 ----------

(async () => {
  const args = parseArgs(process.argv.slice(2));

  if (args.loginOnly) {
    await getSessionToken({ forceWindow: true });
    log('完成。');
    return;
  }

  const row = readRow(args.date);
  const sched = buildSchedule(row.date);

  console.log('');
  console.log('══════════ 這一期要上稿的內容 ══════════');
  console.log(`Excel 第 ${row.excelRow} 列`);
  console.log(`群發名稱：${ymdShort(row.date)}${config.nameSuffix}`);
  console.log(`發送時間：${ymdDash(row.date)} ${config.sendTime}（台灣時間）`);
  console.log(`卡片標題：${row.title}`);
  console.log(`推播通知：${row.notify}`);
  console.log(`卡片敘述：${row.subtitle}`);
  console.log(`圖片：${row.imagePath}`);
  console.log(`電子報網址：${row.url}`);
  console.log('文字訊息：');
  console.log(row.text.split('\n').map((l) => '    ' + l).join('\n'));
  console.log('════════════════════════════════════════');
  console.log('');

  if (sched.getTime() < Date.now() + 5 * 60 * 1000) {
    console.warn('⚠️  排程時間已經過了（或不到 5 分鐘後）。草稿仍會建立，但排程前記得改時間。');
  }

  if (args.dryRun) {
    const preview = buildPayload(row, '<圖片上傳後的網址>', 0);
    console.log('（dry-run）要送出的 payload：');
    console.log(JSON.stringify(preview, null, 2));
    return;
  }

  log('確認登入狀態…');
  const token = await getSessionToken();

  log('上傳圖片…');
  const imageUrl = await uploadImage(row.imagePath);
  log('圖片網址：' + imageUrl);

  log('計算符合條件的客戶數…');
  const customerNum = await getCustomerCount(token);
  log(`符合條件：${customerNum.toLocaleString()} 人`);

  log('建立草稿…');
  const payload = buildPayload(row, imageUrl, customerNum);
  const res = await apiNext('/broadcast/create', token, payload);

  const id = res && (res._id || res.id || res.taskId || res.objectId || (res.result && (res.result._id || res.result.id)));
  const listUrl = `${config.consoleUrl}/broadcast/${config.orgId}`;
  const draftUrl = id ? `${listUrl}/create/${id}` : listUrl;
  console.log('');
  console.log('✅ 草稿已建立：' + payload.name);
  if (id) {
    console.log('   編輯網址：' + draftUrl);
  } else {
    console.log('   （回應裡沒找到 id，請到群發訊息列表的「草稿」分頁確認）');
    console.log('   回應：' + JSON.stringify(res).slice(0, 500));
  }
  console.log('   群發列表：' + listUrl);
  console.log('');

  const now = new Date();
  const stamp = `${now.getFullYear()}-${pad2(now.getMonth() + 1)}-${pad2(now.getDate())} ${pad2(now.getHours())}:${pad2(now.getMinutes())}`;
  writeStatus(row.excelRow, `已建草稿 ${stamp}`, draftUrl);
})().catch((e) => die(String(e && e.stack ? e.stack : e)));

// ---------- 寫回 Excel 的「狀態」「草稿網址」 ----------

function writeStatus(excelRow, status, url) {
  const { spawnSync } = require('child_process');
  const file = path.join(ROOT, config.excelFile);
  const ps1 = path.join(ROOT, 'write-status.ps1');
  const r = spawnSync('powershell', [
    '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', ps1,
    '-File', file, '-Sheet', config.excelSheet, '-Row', String(excelRow),
    '-Status', status, '-Url', url || '',
  ], { encoding: 'utf8', windowsHide: true });

  if (r.status === 0) {
    log(`已在 Excel 第 ${excelRow} 列寫入「${status}」`);
  } else {
    console.warn(`⚠️  草稿建好了，但寫回 Excel 失敗（Excel 是不是還開著？）。請自己在第 ${excelRow} 列的「狀態」填「${status}」，避免下次重複建立。`);
    if (r.stderr) console.warn('   ' + String(r.stderr).trim().split('\n')[0]);
  }
}
