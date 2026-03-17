/**
 * FVdoc Chrome Extension — content.js
 *
 * 役割：
 *  - Google Docs ページ右端にサイドバーを注入する
 *  - ユーザー操作（テキスト入力・設定）を受け付ける
 *  - background.js へメッセージを送って Docs API を呼び出す
 *  - 既存 FVdoc テーブルの一覧取得・均等割付適用に対応
 */

'use strict';

const FVDOC_SIDEBAR_ID = 'fvdoc-sidebar';
const FVDOC_TOGGLE_ID  = 'fvdoc-toggle-btn';
const FVDOC_STYLE_ID   = 'fvdoc-styles';

// 均等割付パネル用の状態（選択中のテーブル情報）
let _justifyTarget = null; // { tableStartIndex, numCols, tablePreview }

// 直近の挿入テーブル情報（自動均等割付ボタン用）
let _lastInserted = null;  // { tableStartIndex, numCols, originalText, charsPerLine }

// 均等割付パネルを閉じた後の戻り先（'insert' or 'list'）
let _justifyBackTarget = 'list';

// テーブル一覧の表示モード（'restore' = テキスト復元+均等割付 / 'justify' = 均等割付のみ）
let _tableListMode = 'restore';

// ─── サイドバー HTML ────────────────────────────────────────────────
function buildSidebarHTML() {
  return `
<div id="fvdoc-header">
  <span id="fvdoc-title">縦書きエディタ <small>FVdoc</small></span>
  <button id="fvdoc-close-btn" title="閉じる">✕</button>
</div>

<div class="fvdoc-section">
  <div class="fvdoc-label">テキスト入力（横書きで入力）</div>
  <textarea id="fvdoc-input"
    placeholder="ここに縦書きにしたい文章を入力してください。&#10;&#10;例：吾輩は猫である。名前はまだ無い。"
    rows="9"></textarea>
  <div id="fvdoc-char-count-row"><span id="fvdoc-char-count">0</span> 文字</div>
  <div class="fvdoc-hint">※ 記号は自動変換：。→︒　、→︑　（）→︵︶　ー→｜　-→︲　―→︱　～→〜　行頭禁則：。、</div>
</div>

<div class="fvdoc-section">
  <div class="fvdoc-label">設定</div>
  <div class="fvdoc-grid">
    <div class="fvdoc-field">
      <label for="fvdoc-chars">1行の文字数</label>
      <input type="number" id="fvdoc-chars" value="20" min="5" max="60">
    </div>
    <div class="fvdoc-field">
      <label for="fvdoc-fontsize">フォントサイズ (pt)</label>
      <input type="number" id="fvdoc-fontsize" value="11" min="6" max="48" step="0.5">
    </div>
    <div class="fvdoc-field">
      <label for="fvdoc-linespacing">行間 (%)</label>
      <input type="number" id="fvdoc-linespacing" value="70" min="20" max="200" step="5">
    </div>
    <div class="fvdoc-field">
      <label for="fvdoc-colgap">列間 (pt)</label>
      <input type="number" id="fvdoc-colgap" value="0" min="0" max="30" step="0.5">
    </div>
    <div class="fvdoc-field fvdoc-full">
      <label for="fvdoc-font">フォント</label>
      <select id="fvdoc-font">
        <option value="Noto Serif JP">Noto Serif JP（明朝体）</option>
        <option value="Noto Sans JP">Noto Sans JP（ゴシック体）</option>
        <option value="BIZ UDMincho">BIZ UDMincho</option>
        <option value="Shippori Mincho">Shippori Mincho</option>
        <option value="default">ドキュメントのデフォルト</option>
      </select>
    </div>
  </div>
</div>

<div class="fvdoc-section" id="fvdoc-preview-section" style="display:none;">
  <div class="fvdoc-label">プレビュー（変換後）</div>
  <div id="fvdoc-preview-box"></div>
</div>

<div class="fvdoc-btns">
  <button id="fvdoc-insert-btn"  class="fvdoc-btn-primary">▶  ドキュメントに挿入</button>
  <button id="fvdoc-load-btn"    class="fvdoc-btn-secondary">↩  既存の縦書き表を読み込む</button>
  <button id="fvdoc-justify-btn" class="fvdoc-btn-justify-menu">⚖  均等割付</button>
</div>

<div id="fvdoc-status" style="display:none;"></div>
<div id="fvdoc-auto-justify-wrap" style="display:none; padding: 0 16px 10px;">
  <button id="fvdoc-auto-justify-btn" class="fvdoc-btn-justify-auto">
    ⚖ 均等割付を設定...
  </button>
</div>

<!-- 既存テーブル一覧パネル -->
<div id="fvdoc-table-list" style="display:none;">
  <div id="fvdoc-table-list-title" class="fvdoc-label" style="margin-bottom:6px;">縦書き表を選択</div>
  <div id="fvdoc-table-items"></div>
  <button id="fvdoc-table-cancel" class="fvdoc-btn-secondary" style="margin-top:8px;width:100%;">キャンセル</button>
</div>

<!-- 均等割付 列選択パネル -->
<div id="fvdoc-justify-panel" style="display:none;">
  <div class="fvdoc-label" style="margin-bottom:4px;">均等割付 — 列を選択</div>
  <div id="fvdoc-justify-desc" class="fvdoc-hint" style="margin-bottom:8px;"></div>
  <div id="fvdoc-justify-cols"></div>
  <div class="fvdoc-justify-panel-btns">
    <button id="fvdoc-justify-select-all"  class="fvdoc-btn-secondary fvdoc-btn-sm">全列選択</button>
    <button id="fvdoc-justify-deselect"    class="fvdoc-btn-secondary fvdoc-btn-sm">全解除</button>
    <button id="fvdoc-justify-reset"       class="fvdoc-btn-secondary fvdoc-btn-sm">均等割付を解除</button>
  </div>
  <div class="fvdoc-justify-panel-btns" style="margin-top:6px;">
    <button id="fvdoc-justify-apply"  class="fvdoc-btn-primary" style="flex:2;">▶ 適用</button>
    <button id="fvdoc-justify-cancel" class="fvdoc-btn-secondary" style="flex:1;">戻る</button>
  </div>
</div>
  `;
}

// ─── スタイル ──────────────────────────────────────────────────────
function buildStyles() {
  return `
/* ──── FVdoc トグルボタン ──── */
#fvdoc-toggle-btn {
  position: fixed;
  right: 0;
  top: 50%;
  transform: translateY(-50%);
  z-index: 99998;
  width: 28px;
  height: 80px;
  background: #1a73e8;
  color: white;
  border: none;
  border-radius: 6px 0 0 6px;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 12px;
  font-weight: 700;
  writing-mode: vertical-rl;
  letter-spacing: 2px;
  box-shadow: -2px 2px 8px rgba(0,0,0,0.25);
  transition: background 0.2s, width 0.2s;
  font-family: 'Noto Serif JP', serif;
}
#fvdoc-toggle-btn:hover { background: #1557b0; width: 32px; }

/* ──── サイドバー本体 ──── */
#fvdoc-sidebar {
  position: fixed;
  top: 0;
  right: -340px;
  width: 320px;
  height: 100vh;
  z-index: 99999;
  background: #fff;
  box-shadow: -4px 0 20px rgba(0,0,0,0.18);
  overflow-y: auto;
  display: flex;
  flex-direction: column;
  transition: right 0.28s cubic-bezier(.4,0,.2,1);
  font-family: 'Google Sans', Roboto, Arial, sans-serif;
  font-size: 13px;
  color: #202124;
}
#fvdoc-sidebar.fvdoc-open { right: 0; }

/* ──── ヘッダー ──── */
#fvdoc-header {
  background: #1a73e8;
  color: white;
  padding: 14px 16px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  flex-shrink: 0;
}
#fvdoc-title { font-size: 15px; font-weight: 600; }
#fvdoc-title small { font-size: 11px; opacity: 0.75; font-weight: 400; margin-left: 4px; }
#fvdoc-close-btn {
  background: none; border: none; color: white; font-size: 16px;
  cursor: pointer; padding: 2px 6px; border-radius: 4px; opacity: 0.85;
}
#fvdoc-close-btn:hover { opacity: 1; background: rgba(255,255,255,0.15); }

/* ──── セクション ──── */
.fvdoc-section { padding: 12px 16px; border-bottom: 1px solid #e8eaed; }
.fvdoc-label {
  font-size: 11px; font-weight: 600; color: #5f6368;
  text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 7px;
}
.fvdoc-hint { font-size: 10px; color: #9aa0a6; margin-top: 5px; line-height: 1.5; }

/* ──── テキストエリア ──── */
#fvdoc-input {
  width: 100%; border: 1px solid #dadce0; border-radius: 4px;
  padding: 8px 10px; font-size: 13px; font-family: inherit; resize: vertical;
  line-height: 1.6; outline: none; box-sizing: border-box; transition: border-color 0.2s;
}
#fvdoc-input:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.12); }
#fvdoc-char-count-row { text-align: right; font-size: 11px; color: #9aa0a6; margin-top: 4px; }

/* ──── 設定グリッド ──── */
.fvdoc-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 10px; }
.fvdoc-field { display: flex; flex-direction: column; gap: 4px; }
.fvdoc-full  { grid-column: 1 / -1; }
.fvdoc-field label { font-size: 11px; color: #5f6368; font-weight: 500; }
.fvdoc-field input,
.fvdoc-field select {
  border: 1px solid #dadce0; border-radius: 4px; padding: 6px 8px;
  font-size: 13px; font-family: inherit; color: #202124; outline: none;
  background: #fff; width: 100%; box-sizing: border-box; transition: border-color 0.2s;
  -webkit-appearance: none; appearance: none;
}
.fvdoc-field input:focus,
.fvdoc-field select:focus { border-color: #1a73e8; box-shadow: 0 0 0 2px rgba(26,115,232,0.12); }

/* ──── プレビュー ──── */
#fvdoc-preview-box {
  background: #f8f9fa; border: 1px solid #e8eaed; border-radius: 4px;
  padding: 8px 10px; font-size: 12px; color: #3c4043;
  word-break: break-all; line-height: 1.7; min-height: 36px;
}

/* ──── ボタン ──── */
.fvdoc-btns { padding: 12px 16px; display: flex; flex-direction: column; gap: 8px; }
.fvdoc-btn-primary {
  width: 100%; padding: 10px; background: #1a73e8; color: white; border: none;
  border-radius: 4px; font-size: 14px; font-weight: 500; cursor: pointer;
  font-family: inherit; transition: background 0.2s, box-shadow 0.2s;
}
.fvdoc-btn-primary:hover    { background: #1557b0; box-shadow: 0 2px 6px rgba(0,0,0,0.2); }
.fvdoc-btn-primary:disabled { background: #dadce0; color: #9aa0a6; cursor: not-allowed; }
.fvdoc-btn-secondary {
  width: 100%; padding: 8px; background: transparent; color: #1a73e8;
  border: 1px solid #1a73e8; border-radius: 4px; font-size: 13px; font-weight: 500;
  cursor: pointer; font-family: inherit; transition: background 0.2s;
}
.fvdoc-btn-secondary:hover  { background: rgba(26,115,232,0.06); }
.fvdoc-btn-sm { padding: 5px 8px; font-size: 11px; width: auto; flex: 1; }
.fvdoc-btn-justify-menu {
  width: 100%; padding: 8px; background: #fff8e1; color: #e37400;
  border: 1px solid #e37400; border-radius: 4px; font-size: 13px; font-weight: 500;
  cursor: pointer; font-family: inherit; transition: background 0.2s;
}
.fvdoc-btn-justify-menu:hover { background: #fef3e2; }
.fvdoc-btn-justify-auto {
  width: 100%; padding: 8px; background: #fff8e1; color: #e37400;
  border: 1px solid #e37400; border-radius: 4px; font-size: 13px; font-weight: 500;
  cursor: pointer; font-family: inherit; transition: background 0.2s;
}
.fvdoc-btn-justify-auto:hover    { background: #fef3e2; }
.fvdoc-btn-justify-auto:disabled { background: #f5f5f5; color: #aaa; border-color: #ddd; cursor: not-allowed; }

/* ──── ステータス ──── */
#fvdoc-status {
  margin: 0 16px 12px; padding: 8px 10px; border-radius: 4px;
  font-size: 12px; line-height: 1.4;
}
#fvdoc-status.ok   { background: #e6f4ea; color: #137333; }
#fvdoc-status.err  { background: #fce8e6; color: #c5221f; }
#fvdoc-status.info { background: #e8f0fe; color: #1a73e8; }

/* ──── テーブル一覧 ──── */
#fvdoc-table-list, #fvdoc-justify-panel {
  padding: 12px 16px;
  border-top: 1px solid #e8eaed;
}
.fvdoc-table-item {
  border: 1px solid #dadce0; border-radius: 4px;
  margin-bottom: 8px; font-size: 12px; color: #202124;
  overflow: hidden;
}
.fvdoc-table-item-top {
  padding: 8px 10px;
}
.fvdoc-table-item-preview { font-weight: 500; }
.fvdoc-table-item-meta    { color: #5f6368; font-size: 11px; margin-top: 2px; }
.fvdoc-table-item-actions {
  display: flex; gap: 0; border-top: 1px solid #e8eaed;
}
.fvdoc-table-action-btn {
  flex: 1; padding: 7px 4px; background: transparent; border: none;
  font-size: 11px; font-weight: 500; cursor: pointer;
  font-family: inherit; transition: background 0.15s; color: #1a73e8;
}
.fvdoc-table-action-btn:hover { background: #e8f0fe; }
.fvdoc-table-action-btn + .fvdoc-table-action-btn { border-left: 1px solid #e8eaed; }
.fvdoc-table-action-btn.justify { color: #e37400; }
.fvdoc-table-action-btn.justify:hover { background: #fef3e2; }

/* ──── 均等割付パネル ──── */
#fvdoc-justify-cols {
  max-height: 240px; overflow-y: auto;
  border: 1px solid #e8eaed; border-radius: 4px;
  margin-bottom: 8px;
}
.fvdoc-justify-col-row {
  display: flex; align-items: center; gap: 8px;
  padding: 6px 10px; border-bottom: 1px solid #f1f3f4;
  cursor: pointer; transition: background 0.1s;
}
.fvdoc-justify-col-row:last-child { border-bottom: none; }
.fvdoc-justify-col-row:hover { background: #f8f9fa; }
.fvdoc-justify-col-row input[type=checkbox] {
  width: 14px; height: 14px; cursor: pointer; accent-color: #e37400;
}
.fvdoc-justify-col-label { flex: 1; font-size: 12px; color: #202124; }
.fvdoc-justify-col-chars { font-size: 11px; color: #9aa0a6; }
.fvdoc-justify-panel-btns {
  display: flex; gap: 6px;
}
  `;
}

// ─── UI 注入 ────────────────────────────────────────────────────────
function inject() {
  if (document.getElementById(FVDOC_SIDEBAR_ID)) return;

  const style = document.createElement('style');
  style.id = FVDOC_STYLE_ID;
  style.textContent = buildStyles();
  document.head.appendChild(style);

  const toggle = document.createElement('button');
  toggle.id = FVDOC_TOGGLE_ID;
  toggle.textContent = '縦書き';
  toggle.title = 'FVdoc 縦書きエディタを開く / 閉じる';
  document.body.appendChild(toggle);

  const sidebar = document.createElement('div');
  sidebar.id = FVDOC_SIDEBAR_ID;
  sidebar.innerHTML = buildSidebarHTML();
  document.body.appendChild(sidebar);

  // イベント設定
  toggle.addEventListener('click', () => sidebar.classList.toggle('fvdoc-open'));
  document.getElementById('fvdoc-close-btn').addEventListener('click', () => sidebar.classList.remove('fvdoc-open'));
  document.getElementById('fvdoc-input').addEventListener('input', onInputChange);
  document.getElementById('fvdoc-chars').addEventListener('change', onInputChange);
  document.getElementById('fvdoc-insert-btn').addEventListener('click', onInsert);
  document.getElementById('fvdoc-load-btn').addEventListener('click', onLoadTables);

  // 自動均等割付ボタン（挿入直後）
  document.getElementById('fvdoc-auto-justify-btn').addEventListener('click', onAutoJustify);

  // 常設「⚖ 均等割付」ボタン
  document.getElementById('fvdoc-justify-btn').addEventListener('click', onOpenJustifyMenu);

  // テーブル一覧キャンセル
  document.getElementById('fvdoc-table-cancel').addEventListener('click', () => {
    document.getElementById('fvdoc-table-list').style.display = 'none';
  });

  // 均等割付パネルのボタン
  document.getElementById('fvdoc-justify-select-all').addEventListener('click', () => {
    document.querySelectorAll('#fvdoc-justify-cols input[type=checkbox]')
      .forEach(cb => cb.checked = true);
  });
  document.getElementById('fvdoc-justify-deselect').addEventListener('click', () => {
    document.querySelectorAll('#fvdoc-justify-cols input[type=checkbox]')
      .forEach(cb => cb.checked = false);
  });
  document.getElementById('fvdoc-justify-reset').addEventListener('click', onResetJustify);
  document.getElementById('fvdoc-justify-apply').addEventListener('click', onApplyJustify);
  document.getElementById('fvdoc-justify-cancel').addEventListener('click', () => {
    document.getElementById('fvdoc-justify-panel').style.display = 'none';
    _justifyTarget = null;
    if (_justifyBackTarget !== 'insert') {
      document.getElementById('fvdoc-table-list').style.display = 'block';
    }
    // insert からの場合は挿入フォームがそのまま表示されるので追加操作不要
  });
}

// ─── 記号変換プレビュー ──────────────────────────────────────────────
function previewReplace(text) {
  return text
    .replace(/。/g,      '︒').replace(/、/g,      '︑')
    .replace(/（/g,      '︵').replace(/）/g,      '︶')
    .replace(/ー/g,      '｜')
    .replace(/「/g,      '﹁').replace(/」/g,      '﹂')
    .replace(/『/g,      '﹃').replace(/』/g,      '﹄')
    .replace(/\[/g,      '﹇').replace(/\]/g,      '﹈')
    .replace(/：/g,      '‥').replace(/:/g,       '‥')
    .replace(/；/g,      '︔').replace(/;/g,       '︔')
    .replace(/！/g,      '︕').replace(/!/g,       '︕')
    .replace(/？/g,      '︖').replace(/\?/g,      '︖')
    .replace(/\u2015/g,  '︱').replace(/\u2014/g,  '︱') // ― — → ︱ 長ダッシュ
    .replace(/\u2013/g,  '︲').replace(/-/g,       '︲') // – - → ︲ 短ダッシュ
    .replace(/\uFF5E/g,  '〜').replace(/\u301C/g,  '〜'); // ～ 〜 → 〜
}

// ─── 禁則処理プレビュー ──────────────────────────────────────────────
// 列頭に来てはいけない句読点（変換後）を前列末尾に移動する
function previewKinsoku(chunks) {
  const forbidden = new Set(['︒', '︑']); // 変換後の句点・読点
  const result = [...chunks];
  let i = 1;
  while (i < result.length) {
    while (result[i].length > 0 && forbidden.has(result[i][0])) {
      result[i - 1] += result[i][0];
      result[i]      = result[i].substring(1);
    }
    if (result[i].length === 0) result.splice(i, 1);
    else i++;
  }
  return result;
}

function onInputChange() {
  const text = document.getElementById('fvdoc-input').value;
  document.getElementById('fvdoc-char-count').textContent =
    text.replace(/[\r\n]/g, '').length;

  const sec = document.getElementById('fvdoc-preview-section');
  if (!text.trim()) { sec.style.display = 'none'; return; }

  const charsPerLine = parseInt(document.getElementById('fvdoc-chars').value) || 20;
  const inputLines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  const rawChunks = [];
  for (const line of inputLines) {
    const converted = previewReplace(line);
    for (let i = 0; i < converted.length; i += charsPerLine)
      rawChunks.push(converted.substring(i, i + charsPerLine));
  }
  const chunks  = previewKinsoku(rawChunks);
  const numCols = chunks.length;
  const previewText = numCols <= 2
    ? chunks.join('  ／  ')
    : chunks.slice(0, 2).join('  ／  ') + `  …（全${numCols}列）`;

  document.getElementById('fvdoc-preview-box').textContent = previewText;
  sec.style.display = 'block';
}

// ─── ドキュメント ID 取得 ────────────────────────────────────────────
function getDocId() {
  const m = location.pathname.match(/\/document\/d\/([a-zA-Z0-9_-]+)/);
  return m ? m[1] : null;
}

// ─── ステータス表示 ──────────────────────────────────────────────────
function showStatus(type, msg) {
  const el = document.getElementById('fvdoc-status');
  el.className = type;
  el.textContent = msg;
  el.style.display = 'block';
  if (type === 'ok') setTimeout(() => { el.style.display = 'none'; }, 4500);
}

function setLoading(on) {
  const btn = document.getElementById('fvdoc-insert-btn');
  btn.disabled = on;
  btn.textContent = on ? '⏳  処理中...' : '▶  ドキュメントに挿入';
}

// ─── 挿入ボタン ──────────────────────────────────────────────────────
async function onInsert() {
  const text = document.getElementById('fvdoc-input').value.trim();
  if (!text) { showStatus('err', 'テキストを入力してください。'); return; }

  const docId = getDocId();
  if (!docId) { showStatus('err', 'ドキュメントIDを取得できませんでした。'); return; }

  const params = {
    text,
    charsPerLine:    parseInt(document.getElementById('fvdoc-chars').value)       || 20,
    fontSize:        parseFloat(document.getElementById('fvdoc-fontsize').value)   || 11,
    fontFamily:      document.getElementById('fvdoc-font').value,
    lineSpacingPct:  parseFloat(document.getElementById('fvdoc-linespacing').value) || 70,
    colGapPt:        parseFloat(document.getElementById('fvdoc-colgap').value)      || 0
  };

  setLoading(true);
  showStatus('info', '処理中です。しばらくお待ちください...');

  try {
    const res = await sendMsg({ action: 'insertVerticalTable', docId, params });
    if (res.ok) {
      showStatus('ok', '✓ 縦書きテーブルを挿入しました！');
      // 均等割付ボタンを表示（挿入したテーブルの情報を保持）
      _lastInserted = {
        tableStartIndex: res.result?.tableStartIndex,
        numCols: res.result?.numCols,
        originalText:  text,
        charsPerLine:  params.charsPerLine
      };
      const wrap = document.getElementById('fvdoc-auto-justify-wrap');
      wrap.style.display = (_lastInserted.tableStartIndex != null) ? 'block' : 'none';
    } else {
      showStatus('err', 'エラー: ' + res.error);
      document.getElementById('fvdoc-auto-justify-wrap').style.display = 'none';
    }
  } catch (e) {
    showStatus('err', 'エラー: ' + e.message);
  } finally {
    setLoading(false);
  }
}

// ─── 均等割付メニューから開く ─────────────────────────────────────────
async function onOpenJustifyMenu() {
  _tableListMode = 'justify';
  await onLoadTablesInternal('fvdoc-justify-btn');
}

// ─── 既存テーブル読み込み（テキスト復元モード） ──────────────────────
async function onLoadTables() {
  _tableListMode = 'restore';
  await onLoadTablesInternal('fvdoc-load-btn');
}

async function onLoadTablesInternal(triggerId) {
  const docId = getDocId();
  if (!docId) { showStatus('err', 'ドキュメントIDを取得できませんでした。'); return; }

  document.getElementById(triggerId).disabled = true;
  showStatus('info', 'ドキュメント内のFVdoc表を検索中...');

  try {
    const res = await sendMsg({ action: 'loadFvdocTables', docId });
    if (!res.ok) throw new Error(res.error);

    const tables = res.tables;
    if (!tables.length) {
      showStatus('err', 'このドキュメントにFVdoc縦書き表は見つかりませんでした。');
      return;
    }

    const listEl  = document.getElementById('fvdoc-table-list');
    const itemsEl = document.getElementById('fvdoc-table-items');
    itemsEl.innerHTML = '';

    // ヘッダーをモードに応じて切り替え
    document.getElementById('fvdoc-table-list-title').textContent =
      _tableListMode === 'justify' ? '均等割付するテーブルを選択' : '縦書き表を選択';

    for (const t of tables) {
      const item = document.createElement('div');
      item.className = 'fvdoc-table-item';

      // チャンク数（≒列数）を計算して表示
      const text     = t.meta.originalText || '';
      const cpLine   = parseInt(t.meta.charsPerLine) || 20;
      const lines    = text.split(/\r?\n/).filter(l => l.trim());
      let   numCols  = 0;
      for (const l of lines) numCols += Math.ceil(l.length / cpLine) || 1;

      if (_tableListMode === 'justify') {
        // 均等割付モード：「均等割付を設定」「均等割付を解除」のみ表示
        item.innerHTML = `
          <div class="fvdoc-table-item-top">
            <div class="fvdoc-table-item-preview">「${t.preview}…」</div>
            <div class="fvdoc-table-item-meta">
              ${cpLine}文字/行 &nbsp;|&nbsp; ${t.meta.fontSize}pt &nbsp;|&nbsp; 全${numCols}列
            </div>
          </div>
          <div class="fvdoc-table-item-actions">
            <button class="fvdoc-table-action-btn justify">設定</button>
            <button class="fvdoc-table-action-btn justify-reset" style="color:#c62828;">解除</button>
          </div>`;

        item.querySelector('.justify').addEventListener('click', () => {
          listEl.style.display = 'none';
          _justifyBackTarget = 'list';
          openJustifyPanel(t, numCols);
        });

        item.querySelector('.justify-reset').addEventListener('click', () => {
          listEl.style.display = 'none';
          _justifyBackTarget = 'list';
          _justifyTarget = { tableStartIndex: t.tableIndex, numCols, meta: t.meta };
          onResetJustify();
        });

      } else {
        // 通常モード：「テキストを復元」「均等割付」
        item.innerHTML = `
          <div class="fvdoc-table-item-top">
            <div class="fvdoc-table-item-preview">「${t.preview}…」</div>
            <div class="fvdoc-table-item-meta">
              ${cpLine}文字/行 &nbsp;|&nbsp;
              ${t.meta.fontSize}pt &nbsp;|&nbsp;
              ${t.meta.fontFamily || 'default'} &nbsp;|&nbsp;
              全${numCols}列
            </div>
          </div>
          <div class="fvdoc-table-item-actions">
            <button class="fvdoc-table-action-btn restore">テキストを復元</button>
            <button class="fvdoc-table-action-btn justify">均等割付</button>
          </div>`;

        item.querySelector('.restore').addEventListener('click', () => {
          applyMeta(t.meta);
          listEl.style.display = 'none';
          document.getElementById('fvdoc-status').style.display = 'none';
          showStatus('ok', '✓ テーブルのテキストを復元しました。');
        });

        item.querySelector('.justify').addEventListener('click', () => {
          listEl.style.display = 'none';
          _justifyBackTarget = 'list';
          openJustifyPanel(t, numCols);
        });
      }

      itemsEl.appendChild(item);
    }

    document.getElementById('fvdoc-status').style.display = 'none';
    listEl.style.display = 'block';
  } catch (e) {
    showStatus('err', 'エラー: ' + e.message);
  } finally {
    document.getElementById(triggerId).disabled = false;
  }
}

// ─── 挿入直後の「均等割付を設定」ボタン → 列選択パネルを開く ─────────
function onAutoJustify() {
  if (!_lastInserted?.tableStartIndex) return;

  // _lastInserted を openJustifyPanel が期待する tableInfo 形式に変換
  const tableInfo = {
    tableIndex: _lastInserted.tableStartIndex,
    meta: {
      originalText: _lastInserted.originalText,
      charsPerLine: _lastInserted.charsPerLine
    }
  };
  // 「戻る」ボタンで挿入画面に戻れるよう、呼び出し元を記録
  _justifyBackTarget = 'insert';
  openJustifyPanel(tableInfo, _lastInserted.numCols);
}

// ─── 均等割付パネルを開く ─────────────────────────────────────────
function openJustifyPanel(tableInfo, numCols) {
  _justifyTarget = {
    tableStartIndex: tableInfo.tableIndex,
    numCols,
    meta: tableInfo.meta
  };

  // チャンクの文字数を計算（プレビュー用）— 禁則処理済みで正確に反映
  const text   = tableInfo.meta.originalText || '';
  const cpLine = parseInt(tableInfo.meta.charsPerLine) || 20;
  const lines  = text.split(/\r?\n/).filter(l => l.trim());
  const rawChunks = [];
  for (const l of lines) {
    const converted = previewReplace(l);
    for (let i = 0; i < converted.length; i += cpLine)
      rawChunks.push(converted.substring(i, i + cpLine));
  }
  const chunks   = previewKinsoku(rawChunks);
  const maxChars = Math.max(...chunks.map(c => c.length), 1);

  // 説明文
  document.getElementById('fvdoc-justify-desc').textContent =
    `最長列 ${maxChars} 文字に合わせて字間を広げます。均等割付を適用する列にチェックを入れてください。`;

  // 列チェックボックスを生成（右端＝第1列から左端まで）
  const colsEl = document.getElementById('fvdoc-justify-cols');
  colsEl.innerHTML = '';
  for (let i = 0; i < chunks.length; i++) {
    const chunkIdx  = i;              // 0=右端列
    const chunkLen  = chunks[i].length;
    const colLabel  = i === 0
      ? `第1列（右端）` : i === chunks.length - 1
      ? `第${i + 1}列（左端）` : `第${i + 1}列`;
    const isMax = chunkLen === maxChars;

    const row = document.createElement('label');
    row.className = 'fvdoc-justify-col-row';
    row.innerHTML = `
      <input type="checkbox" value="${chunkIdx}" ${isMax ? 'disabled' : ''}>
      <span class="fvdoc-justify-col-label">${colLabel}</span>
      <span class="fvdoc-justify-col-chars">${chunkLen}文字${isMax ? '（基準）' : ''}</span>`;
    colsEl.appendChild(row);
  }

  document.getElementById('fvdoc-justify-panel').style.display = 'block';
}

// ─── 均等割付を適用 ───────────────────────────────────────────────
async function onApplyJustify() {
  if (!_justifyTarget) return;

  const docId = getDocId();
  if (!docId) { showStatus('err', 'ドキュメントIDを取得できませんでした。'); return; }

  const checked = Array.from(
    document.querySelectorAll('#fvdoc-justify-cols input[type=checkbox]:checked')
  ).map(cb => parseInt(cb.value));

  if (!checked.length) {
    showStatus('err', '列を1つ以上選択してください。');
    return;
  }

  const applyBtn = document.getElementById('fvdoc-justify-apply');
  applyBtn.disabled = true;
  applyBtn.textContent = '⏳ 処理中...';
  showStatus('info', '均等割付を適用中...');

  try {
    const res = await sendMsg({
      action: 'applyColumnJustify',
      docId,
      tableStartIndex: _justifyTarget.tableStartIndex,
      chunkIndices: checked
    });
    if (res.ok) {
      document.getElementById('fvdoc-justify-panel').style.display = 'none';
      _justifyTarget = null;
      if (_justifyBackTarget === 'insert') {
        document.getElementById('fvdoc-auto-justify-wrap').style.display = 'none';
        _lastInserted = null;
      } else {
        document.getElementById('fvdoc-table-list').style.display = 'block';
      }
      showStatus('ok', `✓ ${checked.length}列に均等割付を適用しました。`);
    } else {
      showStatus('err', 'エラー: ' + res.error);
    }
  } catch (e) {
    showStatus('err', 'エラー: ' + e.message);
  } finally {
    applyBtn.disabled = false;
    applyBtn.textContent = '▶ 適用';
  }
}

// ─── 均等割付を解除（spaceAbove をすべて0に） ────────────────────
async function onResetJustify() {
  if (!_justifyTarget) return;

  const docId = getDocId();
  if (!docId) return;

  // 全列を対象に chunkIndices を構築（全列 spaceAbove=0 は applyColumnJustify で
  // gapBetween=0 と同義なので、「最長列と同じ文字数」として扱う）
  // → 全列のインデックスを送り、background 側で均等割付計算を行うが
  //   すべての列が maxChars と同じ長さなら gap=0 になる
  // → ここでは別途 resetJustify アクションを使う方が確実だが、
  //   現実装では全列を選択して「適用」すると maxChars 基準で計算される。
  // 簡易解除: 全チェックボックスを選択→適用ではなく
  // spaceAbove=0 を直接送る専用アクションとして実装
  const { numCols } = _justifyTarget;
  const allIndices  = Array.from({ length: numCols }, (_, i) => i);

  const resetBtn = document.getElementById('fvdoc-justify-reset');
  resetBtn.disabled = true;

  try {
    const res = await sendMsg({
      action: 'applyColumnJustify',
      docId,
      tableStartIndex: _justifyTarget.tableStartIndex,
      chunkIndices: allIndices,
      resetMode: true  // background.js が gapBetween=0 で処理
    });
    if (res.ok) {
      document.getElementById('fvdoc-justify-panel').style.display = 'none';
      _justifyTarget = null;
      if (_justifyBackTarget !== 'insert') {
        document.getElementById('fvdoc-table-list').style.display = 'block';
      }
      showStatus('ok', '✓ 均等割付を解除しました（全列の字間を均等に戻しました）。');
    } else {
      showStatus('err', 'エラー: ' + res.error);
    }
  } catch (e) {
    showStatus('err', 'エラー: ' + e.message);
  } finally {
    resetBtn.disabled = false;
  }
}

function applyMeta(meta) {
  document.getElementById('fvdoc-input').value       = meta.originalText    || '';
  document.getElementById('fvdoc-chars').value       = meta.charsPerLine    || 20;
  document.getElementById('fvdoc-fontsize').value    = meta.fontSize        || 11;
  document.getElementById('fvdoc-linespacing').value = meta.lineSpacingPct  || 70;
  document.getElementById('fvdoc-colgap').value      = meta.colGapPt        || 0;
  if (meta.fontFamily) document.getElementById('fvdoc-font').value = meta.fontFamily;
  onInputChange();
}

// ─── background.js へメッセージ送信 ─────────────────────────────────
function sendMsg(msg) {
  return new Promise((resolve, reject) => {
    chrome.runtime.sendMessage(msg, res => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve(res);
    });
  });
}

// ─── 注入タイミング ──────────────────────────────────────────────────
function tryInject() {
  if (document.querySelector('.docs-editor-container') ||
      document.querySelector('[data-is-document-content]') ||
      document.querySelector('.kix-appview-editor')) {
    inject();
  } else {
    setTimeout(tryInject, 800);
  }
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => setTimeout(tryInject, 1500));
} else {
  setTimeout(tryInject, 1500);
}
