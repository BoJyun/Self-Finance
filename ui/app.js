'use strict';

let STATE = {
  rows: [], summary: null, settings: {},
  sortKey: '市值', sortDir: -1,
  marketState: '', autoSec: 0
};
let EDIT = { rows: [], mtime: null, dirty: false, seq: 0 };
let autoTimer = null;

const el = (id) => document.getElementById(id);
const K = { CODE: '代號', NAME: '名稱', MARKET: '市場', SHARES: '股數', AVG: '均價', NOTE: '備註' };

/* ── 格式化 ─────────────────────────────────────────────── */
function money(v, digits = 0) {
  if (v === null || v === undefined) return '—';
  return v.toLocaleString('zh-TW', { minimumFractionDigits: digits, maximumFractionDigits: digits });
}
function signed(v, digits = 0) {
  if (v === null || v === undefined) return '—';
  return (v > 0 ? '+' : '') + money(v, digits);
}
function pct(v, digits = 2) {
  if (v === null || v === undefined) return '—';
  return (v > 0 ? '+' : '') + v.toFixed(digits) + '%';
}
function price(v) {
  if (v === null || v === undefined) return '—';
  return v.toLocaleString('zh-TW', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}
/* 0 視為平盤，用灰色 —— 跟台股軟體一致 */
function dirClass(v) {
  if (v === null || v === undefined) return 'flat';
  if (v > 0) return 'up';
  if (v < 0) return 'down';
  return 'flat';
}
function esc(s) {
  return String(s ?? '').replace(/[&<>"]/g, (c) =>
    ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function toast(msg, ms = 3600) {
  const t = el('toast');
  t.textContent = msg;
  t.classList.remove('hidden');
  clearTimeout(toast._t);
  toast._t = setTimeout(() => t.classList.add('hidden'), ms);
}

function banner(msg, isError = false) {
  const b = el('banner');
  if (!msg) { b.classList.add('hidden'); return; }
  b.textContent = msg;
  b.classList.toggle('error', isError);
  b.classList.remove('hidden');
}

/* ── 統計卡 ─────────────────────────────────────────────── */
function renderSummary(s) {
  el('s-today').textContent = signed(s.today_pnl);
  el('s-today').className = 'stat-value ' + dirClass(s.today_pnl);
  el('s-today-pct').textContent = pct(s.today_return);
  el('s-today-pct').className = 'stat-sub ' + dirClass(s.today_pnl);

  el('s-total').textContent = signed(s.total_pnl);
  el('s-total').className = 'stat-value ' + dirClass(s.total_pnl);
  el('s-total-pct').textContent = pct(s.total_return);
  el('s-total-pct').className = 'stat-sub ' + dirClass(s.total_pnl);

  el('s-value').textContent = money(s.market_value);
  el('s-cost').textContent = '成本 ' + money(s.cost);

  if (s.has_us) {
    el('split-card').classList.remove('hidden');
    el('s-split').textContent = money(s.tw_value) + ' / ' + money(s.us_value);
    el('s-split-sub').textContent = '美股佔 '
      + (s.us_ratio === null ? '—' : s.us_ratio.toFixed(1) + '%') + '（已換算台幣）';
  } else {
    el('split-card').classList.add('hidden');
  }

  el('updated').textContent = '上次更新 ' + (s.updated_at || '—')
    + (STATE.marketState ? '  ·  ' + STATE.marketState : '');
  const fx = el('fxline');
  if (s.fx_rate) {
    fx.textContent = `USD/TWD ${s.fx_rate.toFixed(3)} · ${s.fx_source}`;
  } else if (s.has_us) {
    fx.textContent = 'USD/TWD 抓取失敗';
  } else {
    fx.textContent = '';
  }
}

/* ── 庫存表 ─────────────────────────────────────────────── */
function renderRows() {
  const tbody = el('tbody');
  const rows = [...STATE.rows];
  const total = STATE.summary ? STATE.summary.market_value : 0;
  const k = STATE.sortKey, dir = STATE.sortDir;

  rows.sort((a, b) => {
    let x = a[k], y = b[k];
    if (typeof x === 'string' || typeof y === 'string') {
      return String(x ?? '').localeCompare(String(y ?? '')) * dir;
    }
    if (x === null || x === undefined) return 1;      // 沒資料的永遠排最後
    if (y === null || y === undefined) return -1;
    return (x - y) * dir;
  });

  if (!rows.length) {
    tbody.innerHTML = '<tr class="empty"><td colspan="6">沒有持股資料，按「編輯持股」新增</td></tr>';
    return;
  }

  tbody.innerHTML = rows.map((r) => {
    const isUS = r[K.MARKET] === 'US';
    const tag = isUS ? '<span class="tag us">美股</span>' : '<span class="tag">現股</span>';
    const flag = r['遞補價'] ? '<span class="badge">參考價</span>'
      : (r['警告'] ? '<span class="badge">' + esc(r['警告']) + '</span>' : '');
    const share = (total && r['市值'] !== null) ? (r['市值'] / total * 100).toFixed(1) + '%' : '—';
    const origValue = (isUS && r['現價'] !== null)
      ? `<div class="orig">${money(r[K.SHARES] * r['現價'], 0)} USD</div>` : '';
    // 同一代號有多筆時標出「第 n / 共 m 筆」，免得看起來像畫面重複
    const dupMark = r['重複總數'] > 1
      ? `<span class="dup-mark" title="這檔有 ${r['重複總數']} 筆，分開計算">`
        + `${r['重複序號']}/${r['重複總數']}</span>` : '';

    return `<tr>
      <td>
        <div class="code-cell">${tag}
          <div>
            <div class="name">${esc(r[K.NAME])}${dupMark}${flag}</div>
            <div class="code">${esc(r[K.CODE])}</div>
          </div>
        </div>
      </td>
      <td class="num">
        <div class="big ${dirClass(r['漲跌幅'])}">${price(r['現價'])}</div>
        <div class="sub-line ${dirClass(r['漲跌幅'])}">${pct(r['漲跌幅'])}</div>
      </td>
      <td class="num"><div class="big ${dirClass(r['今日損益'])}">${signed(r['今日損益'])}</div></td>
      <td class="num">
        <div class="big ${dirClass(r['總損益'])}">${signed(r['總損益'])}</div>
        <div class="sub-line ${dirClass(r['總損益'])}">${pct(r['報酬率'])}</div>
      </td>
      <td class="num">
        <div class="big">${money(r[K.SHARES])}</div>
        <div class="sub-line dim">${price(r[K.AVG])}</div>
      </td>
      <td class="num">
        <div class="big">${money(r['市值'])}</div>
        <div class="sub-line dim">${share}</div>
        ${origValue}
      </td>
    </tr>`;
  }).join('');
}

function renderFootnote(payload) {
  const s = payload.summary;
  const notes = [];

  if (s.has_us) {
    notes.push('美股損益未計匯差 —— 成本與市值都用當下匯率換算，所以損益只反映股價漲跌。'
      + '要用你實際的換匯價，請到「設定」填手動匯率。');
    if (s.fx_note) notes.push(`匯率來源：${s.fx_source}（${s.fx_note}）`);
  }
  if (s.incomplete && s.incomplete.length) {
    notes.push(`以下標的抓不到報價，未計入總計：${s.incomplete.join('、')}`);
  }
  if (s.no_prev_close && s.no_prev_close.length) {
    notes.push(`以下標的缺昨收，今日損益無法計算：${s.no_prev_close.join('、')}`);
  }
  if (payload.warnings && payload.warnings.length) notes.push(...payload.warnings);
  if (payload.from_cache) {
    notes.push('目前顯示的是上次的快取資料，按「更新報價」取得最新價格。');
  }

  const f = el('footnote');
  if (!notes.length) { f.classList.add('hidden'); return; }
  f.textContent = notes.map((n) => '· ' + n).join('\n');
  f.classList.remove('hidden');
}

function applyScheme(settings) {
  document.body.classList.toggle('scheme-us', settings['漲跌顏色'] === '綠漲紅跌');
}

function apply(payload) {
  STATE.rows = payload.rows || [];
  STATE.summary = payload.summary;
  STATE.marketState = payload.market_state || '';
  STATE.autoSec = payload.auto_refresh_sec || 0;
  STATE.intervalSec = payload.auto_interval_sec || 0;
  if (payload.settings) { STATE.settings = payload.settings; applyScheme(payload.settings); }

  renderSummary(payload.summary);
  renderRows();
  renderFootnote(payload);
  updateSortIndicator();
  rescheduleAuto();
}

function rescheduleAuto() {
  if (autoTimer) { clearInterval(autoTimer); autoTimer = null; }
  const sec = STATE.settings['自動更新'] ? (STATE.autoSec || 0) : 0;
  if (sec >= 5) autoTimer = setInterval(() => { if (!isEditing()) refresh(true); }, sec * 1000);
}

/* ── 更新報價 ───────────────────────────────────────────── */
async function refresh(silent = false) {
  const btn = el('btn-refresh');
  if (!silent) { btn.disabled = true; btn.textContent = '更新中…'; }
  try {
    const payload = await window.pywebview.api.refresh();
    if (payload.settings) { STATE.settings = payload.settings; applyScheme(payload.settings); }
    if (payload.error) {
      banner(payload.error, payload.error_kind !== 'empty');
      if (payload.error_kind === 'empty') { STATE.rows = []; renderRows(); }
    } else {
      banner(null);
      apply(payload);
    }
  } catch (e) {
    banner('更新失敗：' + e, true);
  } finally {
    if (!silent) { btn.disabled = false; btn.textContent = '更新報價'; }
  }
}

async function exportOverview() {
  if (!STATE.summary) { toast('請先更新報價'); return; }
  const res = await window.pywebview.api.export_overview(STATE.rows, STATE.summary);
  toast(res.ok ? '已匯出：' + res.path : '匯出失敗：' + res.error, 5000);
}

function updateSortIndicator() {
  document.querySelectorAll('#table th.sortable').forEach((th) => {
    th.classList.remove('sort-asc', 'sort-desc');
    if (th.dataset.key === STATE.sortKey) {
      th.classList.add(STATE.sortDir === 1 ? 'sort-asc' : 'sort-desc');
    }
  });
}

/* ══════════════ 編輯模式 ══════════════ */
function isEditing() { return !el('view-edit').classList.contains('hidden'); }

async function enterEdit() {
  const res = await window.pywebview.api.get_holdings_for_edit();
  if (res.error) { banner(res.error, true); return; }

  EDIT = { rows: res.rows.map((r) => ({ ...r })), mtime: res.mtime, dirty: false, seq: 0 };
  if (!EDIT.rows.length) EDIT.rows.push(blankRow());

  el('edit-hint').innerHTML = res.excel_open
    ? '⚠ 偵測到 <b>持股.xlsx</b> 正被 Excel 開著，儲存會失敗。請先在 Excel 關掉它。'
    : '直接在格子裡打字。輸入代號後會自動帶出名稱與現價，方便你確認沒打錯。';

  el('edit-error').classList.add('hidden');
  renderEdit();
  el('view-main').classList.add('hidden');
  el('view-edit').classList.remove('hidden');
  setTimeout(() => { const f = document.querySelector('#edit-tbody .cell-input'); if (f) f.focus(); }, 30);
}

function blankRow() {
  return { [K.CODE]: '', [K.NAME]: '', [K.MARKET]: '', [K.SHARES]: '', [K.AVG]: '', [K.NOTE]: '' };
}

function renderEdit() {
  const tb = el('edit-tbody');
  tb.innerHTML = EDIT.rows.map((r, i) => `
    <tr data-i="${i}">
      <td><input class="cell-input" data-f="${K.CODE}" value="${esc(r[K.CODE])}" placeholder="2882"></td>
      <td><div class="lookup" data-lookup="${i}">${r[K.NAME] ? esc(r[K.NAME]) : '<span class="dim">輸入代號自動帶出</span>'}</div></td>
      <td>
        <select class="cell-input" data-f="${K.MARKET}">
          <option value=""${r[K.MARKET] ? '' : ' selected'}>自動</option>
          <option value="TW"${r[K.MARKET] === 'TW' ? ' selected' : ''}>台股</option>
          <option value="US"${r[K.MARKET] === 'US' ? ' selected' : ''}>美股</option>
        </select>
      </td>
      <td><input class="cell-input num" data-f="${K.SHARES}" value="${esc(r[K.SHARES])}" placeholder="1000"></td>
      <td><input class="cell-input num" data-f="${K.AVG}" value="${esc(r[K.AVG])}" placeholder="44.66"></td>
      <td><input class="cell-input" data-f="${K.NOTE}" value="${esc(r[K.NOTE])}" placeholder="選填"></td>
      <td><button class="rowdel" title="刪除這一列">✕</button></td>
    </tr>`).join('');

  tb.querySelectorAll('tr').forEach((tr) => {
    const i = +tr.dataset.i;
    tr.querySelectorAll('.cell-input').forEach((inp) => {
      inp.addEventListener('input', () => {
        EDIT.rows[i][inp.dataset.f] = inp.value;
        EDIT.dirty = true;
        inp.classList.remove('bad');
        const msg = inp.parentElement.querySelector('.cell-msg');
        if (msg) msg.remove();
        if (inp.dataset.f === K.CODE) markDuplicates();
      });
      if (inp.dataset.f === K.CODE || inp.dataset.f === K.MARKET) {
        inp.addEventListener('change', () => doLookup(i));
        inp.addEventListener('blur', () => doLookup(i));
      }
    });
    tr.querySelector('.rowdel').addEventListener('click', () => {
      EDIT.rows.splice(i, 1);
      if (!EDIT.rows.length) EDIT.rows.push(blankRow());
      EDIT.dirty = true;
      renderEdit();
    });
  });

  markDuplicates();
  // 已經有名稱的列，開啟時就把現價補上
  EDIT.rows.forEach((r, i) => { if (r[K.CODE]) doLookup(i); });
}

/* 代號重複是允許的，所以用黃色提示（可以存），不是紅框（不修就不能存） */
function findDuplicates() {
  const seen = {};
  EDIT.rows.forEach((r, i) => {
    const c = String(r[K.CODE] || '').trim();
    if (c) (seen[c] = seen[c] || []).push(i);
  });
  return Object.fromEntries(Object.entries(seen).filter(([, v]) => v.length > 1));
}

function markDuplicates() {
  const dups = findDuplicates();
  document.querySelectorAll('#edit-tbody tr').forEach((tr) => {
    const inp = tr.querySelector(`[data-f="${K.CODE}"]`);
    inp.classList.remove('dup');
    const old = tr.querySelector('.cell-note');
    if (old) old.remove();
  });

  Object.entries(dups).forEach(([code, idxs]) => {
    idxs.forEach((i, n) => {
      const tr = document.querySelector(`#edit-tbody tr[data-i="${i}"]`);
      if (!tr) return;
      const inp = tr.querySelector(`[data-f="${K.CODE}"]`);
      inp.classList.add('dup');
      const d = document.createElement('div');
      d.className = 'cell-note';
      d.textContent = `第 ${n + 1}/${idxs.length} 筆，會分開計算`;
      inp.parentElement.appendChild(d);
    });
  });
}

async function doLookup(i) {
  const row = EDIT.rows[i];
  const box = document.querySelector(`[data-lookup="${i}"]`);
  if (!box) return;
  const code = String(row[K.CODE] || '').trim();
  if (!code) { box.innerHTML = '<span class="dim">輸入代號自動帶出</span>'; return; }

  if (box._code === code && box._mkt === row[K.MARKET]) return;   // 沒變就不重打
  box._code = code; box._mkt = row[K.MARKET];

  const seq = ++EDIT.seq;
  box.classList.remove('bad');
  box.innerHTML = '<span class="dim">查詢中…</span>';

  const res = await window.pywebview.api.lookup(code, row[K.MARKET] || '');
  if (seq !== EDIT.seq && box._code !== code) return;             // 有更新的查詢就丟棄

  if (res.ok) {
    row[K.NAME] = res.name || row[K.NAME];
    row[K.MARKET] = row[K.MARKET] || res.market;
    box.innerHTML = `${esc(res.name)} <span class="px">${price(res.price)}`
      + `${res.currency === 'USD' ? ' USD' : ''}</span>`;
  } else {
    box.classList.add('bad');
    box.textContent = res.error || '查無此代號';
  }
}

function exitEdit() {
  el('view-edit').classList.add('hidden');
  el('view-main').classList.remove('hidden');
}

async function cancelEdit() {
  if (EDIT.dirty && !confirm('有還沒儲存的變更，確定要放棄嗎？')) return;
  exitEdit();
}

async function saveEdit(confirmed = false) {
  const btn = el('btn-save');
  btn.disabled = true; btn.textContent = '儲存中…';
  const errBox = el('edit-error');
  errBox.classList.add('hidden');
  document.querySelectorAll('#edit-tbody .cell-input').forEach((i) => i.classList.remove('bad'));
  document.querySelectorAll('.cell-msg').forEach((m) => m.remove());

  try {
    // 完全空白的列直接略過，使用者不必先刪掉
    const rows = EDIT.rows.filter((r) =>
      String(r[K.CODE] || '').trim() || String(r[K.SHARES] || '').trim()
      || String(r[K.AVG] || '').trim());

    const res = await window.pywebview.api.save_holdings(rows, EDIT.mtime, confirmed);

    if (res.ok) {
      EDIT.dirty = false;
      exitEdit();
      const n = Object.keys(res.duplicates || {}).length;
      toast(`已儲存 ${res.count} 檔到 持股.xlsx`
        + (n ? `（其中 ${n} 個代號有重複，已分開列出）` : ''));
      refresh();
      return;
    }

    // 代號重複：先跳提醒，使用者確認後才真的存
    if (res.needs_confirm === 'duplicates') {
      showDupConfirm(res.duplicates);
      return;
    }

    if (res.errors) {
      res.errors.forEach((e) => {
        const inp = document.querySelector(`#edit-tbody tr[data-i="${e.index}"] [data-f="${e.field}"]`);
        if (inp) {
          inp.classList.add('bad');
          const d = document.createElement('div');
          d.className = 'cell-msg';
          d.textContent = e.msg;
          inp.parentElement.appendChild(d);
        }
      });
      errBox.textContent = `有 ${res.errors.length} 個地方要修正，已用紅框標出來。修好再按儲存。`;
      errBox.classList.remove('hidden');
    } else {
      errBox.textContent = res.error || '儲存失敗';
      errBox.classList.remove('hidden');
    }
  } catch (e) {
    errBox.textContent = '儲存失敗：' + e;
    errBox.classList.remove('hidden');
  } finally {
    btn.disabled = false; btn.textContent = '儲存';
  }
}

function showDupConfirm(dups) {
  // 後端回傳的已經是 1 起算的列號，不要再加 1
  const lines = Object.entries(dups).map(([code, rowNos]) =>
    `<div><span class="code">${esc(code)}</span> 出現 ${rowNos.length} 次`
    + `（第 ${rowNos.join('、')} 列）</div>`);
  el('dup-list').innerHTML = lines.join('');
  el('dup-overlay').classList.remove('hidden');
}

function closeDupConfirm() { el('dup-overlay').classList.add('hidden'); }

/* ══════════════ 設定 ══════════════ */
function openSettings() {
  const s = STATE.settings || {};
  const manual = s['手動匯率'];
  document.querySelector(`input[name="fxmode"][value="${manual ? 'manual' : 'auto'}"]`).checked = true;
  el('fx-manual').value = manual || '';
  el('fx-manual').disabled = !manual;
  el('fx-current').textContent = (STATE.summary && STATE.summary.fx_rate)
    ? `目前 ${STATE.summary.fx_rate.toFixed(3)}（${STATE.summary.fx_source}）` : '';

  const scheme = s['漲跌顏色'] === '綠漲紅跌' ? '綠漲紅跌' : '紅漲綠跌';
  document.querySelector(`input[name="scheme"][value="${scheme}"]`).checked = true;
  el('auto-on').checked = !!s['自動更新'];
  renderMarketState();

  el('settings-overlay').classList.remove('hidden');
}

function renderMarketState() {
  const box = el('market-state');
  if (!STATE.marketState) { box.textContent = ''; return; }
  const secs = STATE.intervalSec || 0;
  const every = secs >= 60 ? `${Math.round(secs / 60)} 分鐘` : `${secs} 秒`;
  box.textContent = el('auto-on').checked
    ? `目前狀態：${STATE.marketState}，每 ${every} 更新一次`
    : `目前狀態：${STATE.marketState}（自動更新已關閉，打開的話會是每 ${every} 一次）`;
}

function closeSettings() { el('settings-overlay').classList.add('hidden'); }

async function commitSettings({ refreshAfter = false } = {}) {
  const mode = document.querySelector('input[name="fxmode"]:checked').value;
  const values = {
    '手動匯率': mode === 'manual' ? el('fx-manual').value.trim() : '',
    '漲跌顏色': document.querySelector('input[name="scheme"]:checked').value,
    '自動更新': el('auto-on').checked,
  };
  const res = await window.pywebview.api.save_settings(values);
  if (!res.ok) { toast('設定儲存失敗：' + res.error); return; }

  STATE.settings = res.settings;
  applyScheme(res.settings);
  rescheduleAuto();
  if (refreshAfter) refresh();
}

function initSettings() {
  document.querySelectorAll('input[name="fxmode"]').forEach((r) => {
    r.addEventListener('change', () => {
      const manual = document.querySelector('input[name="fxmode"]:checked').value === 'manual';
      el('fx-manual').disabled = !manual;
      if (manual) el('fx-manual').focus();
      else commitSettings({ refreshAfter: true });
    });
  });
  // 手動匯率打完才送出，不然每打一個字就重算
  el('fx-manual').addEventListener('change', () => commitSettings({ refreshAfter: true }));
  el('fx-manual').addEventListener('blur', () => commitSettings({ refreshAfter: true }));
  el('fx-manual').addEventListener('keydown', (e) => { if (e.key === 'Enter') e.target.blur(); });

  document.querySelectorAll('input[name="scheme"]').forEach((r) =>
    r.addEventListener('change', () => commitSettings()));      // 顏色不用重抓報價
  el('auto-on').addEventListener('change', () => { commitSettings(); renderMarketState(); });

  el('btn-close-settings').addEventListener('click', closeSettings);
  el('settings-overlay').addEventListener('click', (e) => {
    if (e.target === el('settings-overlay')) closeSettings();
  });
}

/* ── 啟動 ───────────────────────────────────────────────── */
window.addEventListener('pywebviewready', async () => {
  document.querySelectorAll('#table th.sortable').forEach((th) => {
    th.addEventListener('click', () => {
      const k = th.dataset.key;
      if (STATE.sortKey === k) STATE.sortDir *= -1;
      else { STATE.sortKey = k; STATE.sortDir = -1; }
      renderRows();
      updateSortIndicator();
    });
  });

  el('btn-refresh').addEventListener('click', () => refresh());
  el('btn-export').addEventListener('click', exportOverview);
  el('btn-edit').addEventListener('click', enterEdit);
  el('btn-settings').addEventListener('click', openSettings);
  el('btn-cancel').addEventListener('click', cancelEdit);
  el('btn-save').addEventListener('click', () => saveEdit());
  el('dup-cancel').addEventListener('click', closeDupConfirm);
  el('dup-ok').addEventListener('click', () => { closeDupConfirm(); saveEdit(true); });
  el('btn-addrow').addEventListener('click', () => {
    EDIT.rows.push(blankRow());
    EDIT.dirty = true;
    renderEdit();
    const inputs = document.querySelectorAll('#edit-tbody tr:last-child .cell-input');
    if (inputs.length) inputs[0].focus();
  });
  el('btn-folder').addEventListener('click', () => window.pywebview.api.open_data_dir());

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (!el('dup-overlay').classList.contains('hidden')) closeDupConfirm();
      else if (!el('settings-overlay').classList.contains('hidden')) closeSettings();
      else if (isEditing()) cancelEdit();
    }
    if (isEditing() && e.ctrlKey && e.key.toLowerCase() === 's') { e.preventDefault(); saveEdit(); }
    if (!isEditing() && (e.key === 'F5' || (e.ctrlKey && e.key.toLowerCase() === 'r'))) {
      e.preventDefault(); refresh();
    }
  });

  initSettings();

  const init = await window.pywebview.api.get_initial();
  if (init.settings) { STATE.settings = init.settings; applyScheme(init.settings); }
  if (init.notice) banner(init.notice);
  if (init.cached) apply(init.cached);
  refresh();
});
