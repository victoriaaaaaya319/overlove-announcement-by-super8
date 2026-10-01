#!/usr/bin/env node
/**
 * 第三階段：把真金、白銀的 uid csv 上傳到 Super 8 Studio，各建立一個「目標客戶群」
 *
 * 從 .env 的 OUTPUT_DIR 找日期最新的「真金 YYMMDD.csv」「白銀 YYMMDD.csv」，
 * 客戶群名稱為「真金_YYYYMMDD」「白銀_YYYYMMDD」，匯入方式是 LINE UID。
 *
 * 用法：
 *   node upload.js              上傳
 *   node upload.js --dry-run    只顯示會上傳哪些檔、建立什麼名稱，不登入也不上傳
 *
 * 登入：先用 browser-profile/ 裡保存的登入狀態；過期就用 .env 的
 *       SUPER8_EMAIL / SUPER8_PASSWORD 在背景自動登入；沒有才會開視窗請你手動登入。
 */

const fs = require('fs');
const path = require('path');

const ROOT = __dirname;
const CONSOLE_URL = 'https://console.no8.io';
const PARSE_URL = 'https://prod-api-lq.no8.io';
const API_NEXT_URL = 'https://api-next.no8.io';
const UPLOADER_URL = 'https://n.no8.io/uploader/files';
const PROFILE_DIR = 'browser-profile';
const LEVELS = ['真金', '白銀'];
// 客戶群建好後，整群要加上的標籤
const LEVEL_TAGS = { 真金: '【會員】真金工友', 白銀: '【會員】白銀工友' };
const WAIT_MS = 5 * 60 * 1000; // 等 Super 8 背景比對名單／上標籤的上限

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

/** 讀 .env（OUTPUT_DIR / SUPER8_ORG_ID / SUPER8_EMAIL / SUPER8_PASSWORD） */
function readEnv() {
  const file = path.join(ROOT, '.env');
  if (!fs.existsSync(file)) return {};
  const out = {};
  for (const line of fs.readFileSync(file, 'utf8').replace(/^﻿/, '').split(/\r?\n/)) {
    const m = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (!m || line.trim().startsWith('#')) continue;
    let v = m[2];
    if ((v.startsWith('"') && v.endsWith('"')) || (v.startsWith("'") && v.endsWith("'"))) v = v.slice(1, -1);
    out[m[1]] = v;
  }
  return out;
}

// ---------- 找要上傳的 csv ----------

/** 每個級別挑檔名日期最新的一份。回傳 [{ level, file, title }] */
function pickCsvFiles(outputDir) {
  if (!fs.existsSync(outputDir)) die(`找不到資料夾：${outputDir}（先跑第一階段）`);
  const names = fs.readdirSync(outputDir);
  const picks = [];
  for (const level of LEVELS) {
    const dates = names
      .map((n) => n.match(new RegExp(`^${level} (\\d{6})\\.csv$`)))
      .filter(Boolean)
      .map((m) => m[1])
      .sort();
    if (dates.length === 0) {
      console.warn(`⚠️  ${outputDir} 裡沒有「${level} YYMMDD.csv」，這個級別略過。`);
      continue;
    }
    const yymmdd = dates[dates.length - 1];
    picks.push({ level, file: path.join(outputDir, `${level} ${yymmdd}.csv`), title: `${level}_20${yymmdd}` });
  }
  return picks;
}

// ---------- API ----------

async function parseUsersMe(token) {
  const res = await fetch(`${PARSE_URL}/users/me`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({ _method: 'GET', _ApplicationId: PARSE_APP_ID, _JavaScriptKey: PARSE_JS_KEY, _SessionToken: token }),
  });
  if (!res.ok) return null;
  return res.json();
}

/** 呼叫 Parse API（prod-api-lq）。method 是 Parse 的 _method，查詢用 GET */
async function parse(pathname, token, body, method) {
  const res = await fetch(`${PARSE_URL}${pathname}`, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain' },
    body: JSON.stringify({
      ...body,
      ...(method ? { _method: method } : {}),
      _ApplicationId: PARSE_APP_ID,
      _JavaScriptKey: PARSE_JS_KEY,
      _SessionToken: token,
    }),
  });
  const txt = await res.text();
  if (!res.ok) die(`API ${pathname} 失敗（HTTP ${res.status}）：${txt.slice(0, 500)}`);
  return JSON.parse(txt);
}

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

/** 等剛匯入的客戶群比對完成，回傳 { objectId, count } */
async function waitForGroup(token, orgId, title, since) {
  const deadline = Date.now() + WAIT_MS;
  while (Date.now() < deadline) {
    const res = await parse('/classes/CustomerGroup', token, {
      where: {
        organization: { __type: 'Pointer', className: 'Organization', objectId: orgId },
        name: title,
        expired: { $ne: true },
      },
      order: '-createdAt',
      limit: 5,
    }, 'GET');
    // 同名的舊客戶群不算，只認這次送出之後才建立的
    const group = (res.results || []).find((g) => new Date(g.createdAt).getTime() >= since);
    if (group && group.status === 'ready') return group;
    await sleep(3000);
  }
  die(`等了 ${WAIT_MS / 60000} 分鐘，「${title}」還沒比對完成。到後台確認後，標籤要自己上。`);
}

/** 對整個客戶群加標籤 */
function addTagToGroup(token, orgId, groupId, tag) {
  return runTagTask(token, orgId, tag, { tags: [tag], query: { orgId }, groupId, method: 'add' });
}

/** 把身上有 hasTag 的客戶的 tag 標籤拿掉（等同後台「客戶資訊」用標籤篩選後按移除標籤） */
function removeTagFromTagged(token, orgId, hasTag, tag) {
  return runTagTask(token, orgId, tag, {
    tags: [tag],
    query: {
      originalDisplayName: null,
      displayName: null,
      tagDensity: [{ type: 'include', count: 1, tags: [hasTag] }],
      partnerTag: [],
      platforms: [],
      cellPhone: null,
      email: null,
      friendship: null,
      inboxes: [],
      orgId,
      gender: [],
    },
    options: { isNoTagCustomer: false },
    method: 'del',
  });
}

/** 送出標籤任務，等背景跑完，回傳處理的人數 */
async function runTagTask(token, orgId, tag, body) {
  const res = await parse('/functions/task.tag', token, body);
  const taskId = res && res.result && res.result.taskId;
  if (!taskId) die('標籤任務的回應看不懂：' + JSON.stringify(res).slice(0, 300));

  const deadline = Date.now() + WAIT_MS;
  while (Date.now() < deadline) {
    await sleep(2000);
    const tasks = await apiNext(`/task?limit=10&orgId=${orgId}&skip=0`, token, null, 'GET');
    const task = ((tasks && tasks.data) || []).find((t) => t._id === taskId);
    if (task && task.status === 'done') return task.total;
    if (task && /fail|error/i.test(task.status)) die(`標籤「${tag}」的任務失敗（狀態：${task.status}）`);
  }
  die(`等了 ${WAIT_MS / 60000} 分鐘，標籤「${tag}」的任務還沒跑完。到後台右上的「系統任務」確認。`);
}

async function apiNext(pathname, token, body, method = 'POST') {
  const res = await fetch(`${API_NEXT_URL}${pathname}`, {
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

/** 把 csv 傳到 Super 8 的檔案空間，回傳檔案網址 */
async function uploadCsv(file) {
  const signRes = await fetch(`${UPLOADER_URL}/${encodeURIComponent(path.basename(file))}`, {
    headers: { Accept: 'application/json, text/plain, */*' },
  });
  if (!signRes.ok) die(`取得上傳網址失敗（HTTP ${signRes.status}）`);
  const { url: putUrl, location } = await signRes.json();
  if (!putUrl || !location) die('上傳網址回應格式不對');

  const putRes = await fetch(putUrl, { method: 'PUT', headers: { 'Content-Type': 'text/csv' }, body: fs.readFileSync(file) });
  if (!putRes.ok) die(`csv 上傳失敗（HTTP ${putRes.status}）`);
  return location;
}

// ---------- 登入狀態（用瀏覽器 profile 保存） ----------

async function launch(headless) {
  const { chromium } = require('playwright-core');
  const profileDir = path.join(ROOT, PROFILE_DIR);
  const opts = { headless, viewport: { width: 1280, height: 800 } };
  // 優先用電腦上已裝好的 Edge / Chrome，不用另外下載瀏覽器
  for (const channel of ['msedge', 'chrome']) {
    try {
      return await chromium.launchPersistentContext(profileDir, { ...opts, channel });
    } catch (e) {
      if (!/Chromium distribution|not found|executable/i.test(String(e))) throw e;
    }
  }
  die('找不到 Edge 或 Chrome。請安裝其中一個。');
}

async function readTokenFromPage(page) {
  const raw = await page.evaluate((k) => localStorage.getItem(k), CURRENT_USER_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw).sessionToken || null; } catch (_) { return null; }
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
    const msg = await page.locator('main, body').innerText().then((t) => t.slice(0, 300)).catch(() => '');
    log('自動登入沒成功，頁面內容：' + msg.replace(/\s+/g, ' '));
  }
  return token;
}

/** 回傳有效的 sessionToken。順序：保存的登入狀態 → .env 自動登入 → 開視窗手動登入 */
async function getSessionToken(env) {
  // 1. 先用保存的登入狀態（無頭）
  let ctx = await launch(true);
  let page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(CONSOLE_URL, { waitUntil: 'domcontentloaded' });

  let token = await readTokenFromPage(page);
  if (token && (await parseUsersMe(token))) {
    await ctx.close();
    return token;
  }

  // 2. 有 .env 帳密就自動登入（同一個無頭視窗）
  if (env.SUPER8_EMAIL && env.SUPER8_PASSWORD) {
    log('登入狀態過期或不存在，用 .env 的帳號自動登入…');
    token = await loginWithEnv(page, env).catch((e) => { log('自動登入出錯：' + e.message); return null; });
    if (token) {
      log('自動登入成功，登入狀態已保存到 ' + PROFILE_DIR + '/');
      await ctx.close();
      return token;
    }
  } else {
    log('.env 沒有 SUPER8_EMAIL / SUPER8_PASSWORD，改為開視窗手動登入。');
  }

  // 3. 開視窗請使用者手動登入
  await ctx.close();
  ctx = await launch(false);
  page = ctx.pages()[0] || (await ctx.newPage());
  await page.goto(CONSOLE_URL, { waitUntil: 'domcontentloaded' });
  console.log('\n🔐 請在剛開啟的瀏覽器視窗登入 Super 8 Studio。登入完成後這裡會自動繼續。\n');
  token = await waitForToken(page, 10 * 60 * 1000);
  if (!token) die('等了 10 分鐘還沒登入（或視窗被關掉），先結束。');
  log('登入成功，登入狀態已保存到 ' + PROFILE_DIR + '/');

  await ctx.close();
  return token;
}

// ---------- 主流程 ----------

(async () => {
  const args = process.argv.slice(2);
  const dryRun = args.includes('--dry-run') || args.includes('--dry');
  const unknown = args.filter((a) => !['--dry-run', '--dry'].includes(a));
  if (unknown.length) die(`看不懂的參數：${unknown.join(' ')}`);

  const env = readEnv();
  if (!env.OUTPUT_DIR) die('還沒設定 csv 在哪裡：複製 .env.example 改名成 .env，填上 OUTPUT_DIR');
  const orgId = (env.SUPER8_ORG_ID || '').trim();
  if (!orgId) die('缺少 SUPER8_ORG_ID。請在這個資料夾的 .env 裡加上一行：SUPER8_ORG_ID=你的組織ID');

  const picks = pickCsvFiles(env.OUTPUT_DIR);
  if (picks.length === 0) die('沒有可以上傳的 csv。');

  console.log('');
  console.log('══════════ 要建立的目標客戶群 ══════════');
  for (const p of picks) console.log(`${p.title}  ←  ${path.basename(p.file)}  →  標籤 ${LEVEL_TAGS[p.level]}`);
  console.log('════════════════════════════════════════');
  console.log('');

  if (dryRun) {
    console.log('（dry-run）沒有登入，也沒有上傳。');
    return;
  }

  log('確認登入狀態…');
  const token = await getSessionToken(env);

  // 先把客戶群都建好，再動標籤：上傳或比對失敗時，現有的標籤不會被清掉
  for (const p of picks) {
    log(`上傳 ${path.basename(p.file)}…`);
    const fileUrl = await uploadCsv(p.file);
    const since = Date.now() - 60 * 1000; // 容許本機與伺服器時間差
    const res = await apiNext('/customer_group/import/csv', token, { title: p.title, type: 'uid', fileUrl, orgId });
    if (!res || res.ok !== true) die(`建立「${p.title}」的回應不是成功：${JSON.stringify(res).slice(0, 300)}`);

    log(`等 Super 8 比對名單…`);
    p.group = await waitForGroup(token, orgId, p.title, since);
    console.log(`✅ 客戶群「${p.title}」建立完成：${p.group.count.toLocaleString()} 人`);
    if (!p.group.count) die(`「${p.title}」比對到 0 人，csv 可能有問題。標籤都沒有動。`);
  }

  // 標籤以最新名單為準：先清掉所有人身上的舊標籤，再加給這次名單裡的人。
  // 到期、降級、退會的人這樣就不會留著舊標籤
  for (const p of picks) {
    const tag = LEVEL_TAGS[p.level];
    log(`清掉舊的 ${tag}…`);
    const cleared = await removeTagFromTagged(token, orgId, tag, tag);
    console.log(`✅ 已從 ${Number(cleared || 0).toLocaleString()} 人身上拿掉舊的「${tag}」`);

    log(`加上標籤 ${tag}…`);
    const total = await addTagToGroup(token, orgId, p.group.objectId, tag);
    console.log(`✅ 已為 ${Number(total).toLocaleString()} 人加上「${tag}」`);
  }

  // 升級成真金的人不該再帶著白銀標籤。真金沒上傳成功就不做，避免用到舊的真金名單
  if (picks.some((p) => p.level === '真金')) {
    log(`拿掉真金工友身上的 ${LEVEL_TAGS.白銀}…`);
    const removed = await removeTagFromTagged(token, orgId, LEVEL_TAGS.真金, LEVEL_TAGS.白銀);
    console.log(`✅ 已為 ${Number(removed || 0).toLocaleString()} 位真金工友拿掉「${LEVEL_TAGS.白銀}」`);
  }

  console.log('');
  console.log('目標客戶群列表：');
  console.log(`   ${CONSOLE_URL}/customer-management/${orgId}/customer-groups`);
  console.log('');
})().catch((e) => die(String(e && e.stack ? e.stack : e)));
