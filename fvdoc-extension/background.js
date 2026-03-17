/**
 * FVdoc Chrome Extension — background.js (Service Worker)
 *
 * 【アーキテクチャ】
 *  旧: 1文字 = 1セル（行×列のグリッド） → 字間はセルパディングで制御
 *  新: 1列   = 1セル（改行区切りで文字を縦積み）→ 字間は段落スペースで制御
 *
 *  テーブル構造:
 *    行0: コンテンツ（各セルに列の文字を \n 区切りで格納）
 *    行1: メタデータ（FVDOC_META: JSON を不可視テキストで格納）
 *    ※ スペーサー列あり → totalCols = numCols + 1
 *
 *  均等割付:
 *    各列の文字数が最長列より短い場合、spaceAbove で字間を広げて高さを揃える
 */

'use strict';

const DOCS_API    = 'https://docs.googleapis.com/v1/documents';
const FVDOC_MARKER = 'FVDOC_META:';

// ─── メッセージハンドラ ────────────────────────────────────────────
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === 'insertVerticalTable') {
    handleInsert(request.docId, request.params)
      .then(result => sendResponse({ ok: true, result }))
      .catch(err   => sendResponse({ ok: false, error: err.message }));
    return true;
  }
  if (request.action === 'loadFvdocTables') {
    loadFvdocTables(request.docId)
      .then(tables => sendResponse({ ok: true, tables }))
      .catch(err   => sendResponse({ ok: false, error: err.message }));
    return true;
  }
  if (request.action === 'applyColumnJustify') {
    applyColumnJustify(request.docId, request.tableStartIndex, request.chunkIndices, request.resetMode)
      .then(result => sendResponse({ ok: true, result }))
      .catch(err   => sendResponse({ ok: false, error: err.message }));
    return true;
  }
});

// ─── OAuth トークン取得 ────────────────────────────────────────────
function getToken() {
  return new Promise((resolve, reject) => {
    chrome.identity.getAuthToken({ interactive: true }, token => {
      if (chrome.runtime.lastError) reject(new Error(chrome.runtime.lastError.message));
      else resolve(token);
    });
  });
}

// ─── Docs API ユーティリティ ──────────────────────────────────────
async function docsGet(token, docId) {
  const res = await fetch(`${DOCS_API}/${docId}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  if (!res.ok) throw new Error(`GET ${res.status}: ${await res.text()}`);
  return res.json();
}

async function batchUpdate(token, docId, requests) {
  if (!requests.length) return;
  const res = await fetch(`${DOCS_API}/${docId}:batchUpdate`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json'
    },
    body: JSON.stringify({ requests })
  });
  if (!res.ok) throw new Error(`BatchUpdate ${res.status}: ${await res.text()}`);
  return res.json();
}

// ─── 縦書き記号変換 ───────────────────────────────────────────────
function replaceSymbols(text) {
  return text
    .replace(/。/g,      '\uFE12') // ︒ 縦書き句点
    .replace(/、/g,      '\uFE11') // ︑ 縦書き読点
    .replace(/（/g,      '\uFE35') // ︵ 縦書き左括弧
    .replace(/）/g,      '\uFE36') // ︶ 縦書き右括弧
    .replace(/ー/g,      '\uFF5C') // ｜ 縦書き長音符（カタカナ）
    .replace(/「/g,      '\uFE41') // ﹁ 縦書き左鉤括弧
    .replace(/」/g,      '\uFE42') // ﹂ 縦書き右鉤括弧
    .replace(/『/g,      '\uFE43') // ﹃ 縦書き左二重鉤括弧
    .replace(/』/g,      '\uFE44') // ﹄ 縦書き右二重鉤括弧
    .replace(/\[/g,      '\uFE47') // ﹇ 縦書き左角括弧
    .replace(/\]/g,      '\uFE48') // ﹈ 縦書き右角括弧
    .replace(/：/g,      '\u2025') // ‥ 2点リーダー（全角コロン）
    .replace(/:/g,       '\u2025') // ‥ 2点リーダー（半角コロン）
    .replace(/；/g,      '\uFE14') // ︔ 縦書きセミコロン（全角）
    .replace(/;/g,       '\uFE14') // ︔ 縦書きセミコロン（半角）
    .replace(/！/g,      '\uFE15') // ︕ 縦書き感嘆符（全角）
    .replace(/!/g,       '\uFE15') // ︕ 縦書き感嘆符（半角）
    .replace(/？/g,      '\uFE16') // ︖ 縦書き疑問符（全角）
    .replace(/\?/g,      '\uFE16') // ︖ 縦書き疑問符（半角）
    // ── ダッシュ類（短 → ︲ / 長 → ︱ で区別）──────────────────
    .replace(/\u2015/g,  '\uFE31') // ― → ︱ 縦書き長ダッシュ（水平バー）
    .replace(/\u2014/g,  '\uFE31') // — → ︱ 縦書き長ダッシュ（エムダッシュ）
    .replace(/\u2013/g,  '\uFE32') // – → ︲ 縦書き短ダッシュ（エンダッシュ）
    .replace(/-/g,       '\uFE32') // - → ︲ 縦書き短ダッシュ（ハイフン）
    // ── 波ダッシュ ────────────────────────────────────────────────
    .replace(/\uFF5E/g,  '\u301C') // ～ → 〜 波ダッシュ（縦向きは近似）
    .replace(/\u301C/g,  '\u301C'); // 〜 はそのまま保持
}

// ─── 単位変換ヘルパー（PT / MM / INCH → PT） ─────────────────────
function toPt(dim) {
  if (!dim || dim.magnitude == null) return null;
  switch (dim.unit) {
    case 'MM':   return dim.magnitude * (72 / 25.4);
    case 'INCH': return dim.magnitude * 72;
    default:     return dim.magnitude; // 'PT' or unspecified
  }
}

// ─── 禁則処理 ──────────────────────────────────────────────────────
// 列頭に来てはいけない文字（句読点）を前の列の末尾に送る
//
// 変換後の文字コードで判定：
//   \uFE12 = ︒（句点）  \uFE11 = ︑（読点）
//
// 処理後のチャンクは charsPerLine + 1 文字になる場合があるが、
// lineSpacing を比例縮小することで全列の高さを揃える（STEP5で対応）
function applyKinsoku(chunks) {
  // 行頭禁則文字（変換後）
  const forbidden = new Set(['\uFE12', '\uFE11']); // ︒ ︑
  const result = [...chunks];
  let i = 1;
  while (i < result.length) {
    // 前列に吸収できるだけ先頭の禁則文字を移動
    while (result[i].length > 0 && forbidden.has(result[i][0])) {
      result[i - 1] += result[i][0];
      result[i]      = result[i].substring(1);
    }
    if (result[i].length === 0) {
      result.splice(i, 1); // 空になった列を削除
    } else {
      i++;
    }
  }
  return result;
}

// ─── 列チャンク構築 ───────────────────────────────────────────────
// 入力テキスト → 列ごとの文字列配列
//   chunks[0] = 右端の列（最初の入力行の先頭 charsPerLine 文字）
//   改行 = 列区切り、1行が charsPerLine を超える場合はさらに分割
function buildChunks(text, charsPerLine) {
  const inputLines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  const rawChunks = [];
  for (const line of inputLines) {
    const converted = replaceSymbols(line);
    if (!converted) continue;
    for (let i = 0; i < converted.length; i += charsPerLine) {
      rawChunks.push(converted.substring(i, i + charsPerLine));
    }
  }
  return applyKinsoku(rawChunks);
}

// ─── メインロジック：テーブルの挿入 ──────────────────────────────
async function handleInsert(docId, params) {
  const { text, charsPerLine, fontSize, fontFamily, lineSpacingPct, colGapPt } = params;
  const cpLine    = parseInt(charsPerLine);
  const fSize     = parseFloat(fontSize);
  const lSpacing  = parseFloat(lineSpacingPct) || 70;  // 行間（%）
  const colGap    = parseFloat(colGapPt)       || 0;   // 列間（pt）
  const token     = await getToken();

  const chunks = buildChunks(text, cpLine);
  if (!chunks.length) throw new Error('テキストがありません');

  const numCols  = chunks.length;

  // ── STEP 1: ドキュメント情報取得 ──────────────────────────────
  const doc = await docsGet(token, docId);
  const bodyContent  = doc.body.content;
  const insertAt     = bodyContent[bodyContent.length - 1].startIndex;

  // 列幅 = フォントサイズ + 左右パディング各1pt + 列間（半分ずつ左右に配分）
  const cellPadH   = 1 + colGap / 2;              // 左右パディング（pt）
  const colWidthPt = fSize + cellPadH * 2;         // 列幅合計
  const contentWidthPt = colWidthPt * numCols;

  // ページ幅・マージンからスペーサー幅を正確に計算
  // EVENLY_DISTRIBUTED は多列時に右へのはみ出しが発生するため FIXED_WIDTH に変更
  const docStyle      = doc.documentStyle || {};
  const pageWidthPt   = toPt(docStyle.pageSize?.width)   || 595; // A4 デフォルト
  const marginLeftPt  = toPt(docStyle.marginLeft)         || 72;  // 1 inch デフォルト
  const marginRightPt = toPt(docStyle.marginRight)        || 72;
  const usableWidthPt = pageWidthPt - marginLeftPt - marginRightPt;
  // コンテンツがページ幅を超える場合はスペーサー=1pt（左端揃え）、収まる場合は右端揃え
  const spacerWidthPt = Math.max(1, usableWidthPt - contentWidthPt);

  // スペーサー列は常に追加（右揃え用）
  const totalCols = numCols + 1;

  // ── STEP 2: 空テーブルを挿入（スペーサー列込み） ─────────────
  await batchUpdate(token, docId, [{
    insertTable: {
      rows: 2,
      columns: totalCols,
      location: { index: insertAt }
    }
  }]);

  // ── STEP 3: 更新済みドキュメントを取得 ───────────────────────
  const updatedDoc    = await docsGet(token, docId);
  const tableElements = updatedDoc.body.content.filter(el => el.table);
  if (!tableElements.length) throw new Error('テーブルの挿入に失敗しました');
  const newTableEl    = tableElements[tableElements.length - 1];
  const newTable      = newTableEl.table;
  const newTableStart = newTableEl.startIndex;

  // ── STEP 4a: 列幅を設定 ──────────────────────────────────────
  // スペーサー列（col 0）: FIXED_WIDTH = spacerWidthPt（ページ幅から逆算）
  //   ページ幅 - マージン - コンテンツ列幅合計 = スペーサー幅
  //   → コンテンツをページ右端に揃える。超過時は1ptにして左端揃え。
  // コンテンツ列（col 1〜）: FIXED_WIDTH = fSize + cellPadH*2
  const colWidthReqs = [{
    updateTableColumnProperties: {
      tableStartLocation: { index: newTableStart },
      columnIndices: [0],
      tableColumnProperties: {
        widthType: 'FIXED_WIDTH',
        width: { magnitude: spacerWidthPt, unit: 'PT' }
      },
      fields: 'widthType,width'
    }
  }];
  for (let col = 1; col < totalCols; col++) {
    colWidthReqs.push({
      updateTableColumnProperties: {
        tableStartLocation: { index: newTableStart },
        columnIndices: [col],
        tableColumnProperties: {
          widthType: 'FIXED_WIDTH',
          width: { magnitude: colWidthPt, unit: 'PT' }
        },
        fields: 'widthType,width'
      }
    });
  }
  await batchUpdate(token, docId, colWidthReqs);

  // ── STEP 4b: テキスト挿入（インデックス降順） ─────────────────
  const insertions = [];

  for (let col = 1; col < totalCols; col++) {  // col 0 = スペーサー、スキップ
    const contentColIdx = col - 1;              // 0-based content col
    const chunkIdx      = numCols - 1 - contentColIdx; // 右端 = chunks[0]
    const chunk         = chunks[chunkIdx];
    const cellText      = chunk.split('').join('\n');

    const tableCell = newTable.tableRows[0].tableCells[col];
    const firstEl   = tableCell.content?.[0]?.paragraph?.elements?.[0];
    if (!firstEl) continue;
    insertions.push({ index: firstEl.startIndex, text: cellText });
  }

  // 行1: メタデータ（col1=最初のコンテンツ列に格納）
  // ※ col0 は多列時に極端に狭くなる場合があるため
  const metaJson    = JSON.stringify({ originalText: text, charsPerLine: cpLine, fontSize: fSize, fontFamily, lineSpacingPct: lSpacing, colGapPt: colGap });
  const metaCell    = newTable.tableRows[1].tableCells[1];
  const metaFirstEl = metaCell.content?.[0]?.paragraph?.elements?.[0];
  if (metaFirstEl) {
    insertions.push({ index: metaFirstEl.startIndex, text: FVDOC_MARKER + metaJson });
  }

  insertions.sort((a, b) => b.index - a.index);
  await batchUpdate(token, docId, insertions.map(({ index, text: t }) => ({
    insertText: { location: { index }, text: t }
  })));

  // ── STEP 5: スタイル適用（ドキュメントを再取得）──────────────
  const styledDoc    = await docsGet(token, docId);
  const styledTables = styledDoc.body.content.filter(el => el.table);
  const styledTableEl = styledTables[styledTables.length - 1];
  const styledTable   = styledTableEl.table;
  const tableStartIndex = styledTableEl.startIndex;

  // ── セル・テキストスタイル（列幅は STEP4a で設定済み）────────
  const styleReqs = [];
  const white      = { color: { rgbColor: { red: 1, green: 1, blue: 1 } } };
  const whiteBorder = { color: white, width: { magnitude: 1, unit: 'PT' }, dashStyle: 'SOLID' };
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < totalCols; col++) {
      const isSpacerCol = (col === 0); // スペーサー列は常に col 0
      const tableCell   = styledTable.tableRows[row].tableCells[col];

      // セル枠線を白（不可視）・パディングをゼロに
      styleReqs.push({
        updateTableCellStyle: {
          tableRange: {
            tableCellLocation: {
              tableStartLocation: { index: tableStartIndex },
              rowIndex: row,
              columnIndex: col
            },
            rowSpan: 1, columnSpan: 1
          },
          tableCellStyle: {
            borderTop:     whiteBorder,
            borderBottom:  whiteBorder,
            borderLeft:    whiteBorder,
            borderRight:   whiteBorder,
            paddingTop:    { magnitude: 0, unit: 'PT' },
            paddingBottom: { magnitude: 0, unit: 'PT' },
            paddingLeft:   { magnitude: isSpacerCol ? 0 : cellPadH, unit: 'PT' },
            paddingRight:  { magnitude: isSpacerCol ? 0 : cellPadH, unit: 'PT' },
            contentAlignment: 'TOP'
          },
          fields: 'borderTop,borderBottom,borderLeft,borderRight,paddingTop,paddingBottom,paddingLeft,paddingRight,contentAlignment'
        }
      });

      if (row === 0 && !isSpacerCol) {
        // ── コンテンツセル：段落ごとに1文字 ──────────────────
        const contentColIdx = col - 1; // col 0 = スペーサー、content は 1〜
        const chunkIdx      = numCols - 1 - contentColIdx;
        const chunkLen      = chunks[chunkIdx].length;

        // tableCell.content = 段落の配列（各段落が1文字 + 末尾の空段落）
        const paragraphs = tableCell.content || [];
        const lastParaIdx = paragraphs.length - 1; // 末尾の空段落インデックス
        let charIdx = 0;

        for (let pi = 0; pi < paragraphs.length; pi++) {
          const contentEl = paragraphs[pi];
          if (!contentEl.paragraph) continue;

          const paraStart  = contentEl.startIndex;
          const paraEnd    = contentEl.endIndex;
          const isTrailing = pi === lastParaIdx; // 末尾の空段落

          // 禁則処理で文字が増えた列は lineSpacing を比例縮小して高さを揃える
          // 例: charsPerLine=20, lSpacing=70, chunkLen=21 → 70 * 20/21 ≈ 67
          const lineSpacing = chunkLen > cpLine
            ? Math.round(cpLine * lSpacing / chunkLen)
            : lSpacing;

          styleReqs.push({
            updateParagraphStyle: {
              range: { startIndex: paraStart, endIndex: paraEnd },
              paragraphStyle: {
                alignment: 'CENTER',
                lineSpacing,
                spaceAbove: { magnitude: 0, unit: 'PT' },
                spaceBelow: { magnitude: 0, unit: 'PT' }
              },
              fields: 'alignment,lineSpacing,spaceAbove,spaceBelow'
            }
          });

          // テキストスタイル（フォントサイズ・書体）
          for (const pe of contentEl.paragraph.elements || []) {
            if (!pe.textRun) continue;
            const textStyle = { fontSize: { magnitude: fSize, unit: 'PT' } };
            let fields = 'fontSize';
            if (fontFamily && fontFamily !== 'default') {
              textStyle.weightedFontFamily = { fontFamily };
              fields += ',weightedFontFamily';
            }
            styleReqs.push({
              updateTextStyle: {
                range: { startIndex: pe.startIndex, endIndex: pe.endIndex },
                textStyle,
                fields
              }
            });
          }

          if (!isTrailing) charIdx++;
        }

      } else if (row === 1) {
        // ── メタデータ行：フォントサイズ1pt・白文字・行間最小化で不可視化 ──
        // row1 の高さを最小限に抑えてページ増加を防ぐ
        for (const contentEl of tableCell.content || []) {
          if (!contentEl.paragraph) continue;

          // 段落スタイル：行間1%・前後スペースゼロで高さを最小化
          styleReqs.push({
            updateParagraphStyle: {
              range: { startIndex: contentEl.startIndex, endIndex: contentEl.endIndex },
              paragraphStyle: {
                lineSpacing:  1,
                spaceAbove:   { magnitude: 0, unit: 'PT' },
                spaceBelow:   { magnitude: 0, unit: 'PT' }
              },
              fields: 'lineSpacing,spaceAbove,spaceBelow'
            }
          });

          for (const pe of contentEl.paragraph?.elements || []) {
            if (!pe.textRun) continue;
            styleReqs.push({
              updateTextStyle: {
                range: { startIndex: pe.startIndex, endIndex: pe.endIndex },
                textStyle: {
                  fontSize:        { magnitude: 1, unit: 'PT' },
                  foregroundColor: white
                },
                fields: 'fontSize,foregroundColor'
              }
            });
          }
        }
      }
    }
  }

  await batchUpdate(token, docId, styleReqs);

  // 均等割付ボタン用に tableStartIndex と列数を返す
  return { success: true, tableStartIndex, numCols };
}

// ─── 均等割付：指定列の字間を再計算して適用 ──────────────────────────
//
// chunkIndices: 均等割付する列のチャンクインデックス（0=右端列, 1=その左…）
//
// resetMode=true のとき spaceAbove をすべて 0 にリセット（均等割付解除）
async function applyColumnJustify(docId, tableStartIndex, chunkIndices, resetMode = false) {
  const token = await getToken();
  const doc   = await docsGet(token, docId);

  // tableStartIndex が一致するテーブルを検索
  const tableEl = doc.body.content.find(el => el.table && el.startIndex === tableStartIndex);
  if (!tableEl) throw new Error('指定されたテーブルが見つかりませんでした');
  const table = tableEl.table;

  // メタデータを最終行から取得（col1 優先、見つからなければ col0 にフォールバック）
  const lastRow   = table.tableRows[table.tableRows.length - 1];
  let cellText = '';
  const metaCandidates = [lastRow.tableCells[1], lastRow.tableCells[0]].filter(Boolean);
  for (const candidate of metaCandidates) {
    let t = '';
    for (const contentEl of candidate.content || []) {
      for (const pe of contentEl.paragraph?.elements || []) {
        if (pe.textRun) t += pe.textRun.content;
      }
    }
    if (t.includes(FVDOC_MARKER)) { cellText = t; break; }
  }
  if (!cellText.includes(FVDOC_MARKER)) throw new Error('FVdocメタデータが見つかりません（このテーブルはFVdoc形式ではありません）');

  const jsonStr = cellText
    .substring(cellText.indexOf(FVDOC_MARKER) + FVDOC_MARKER.length)
    .replace(/\n/g, '');
  const meta = JSON.parse(jsonStr);
  const { originalText, charsPerLine, fontSize: fSizeRaw } = meta;
  const fSize = parseFloat(fSizeRaw);

  // チャンク再構築
  const chunks  = buildChunks(originalText, parseInt(charsPerLine));
  const numCols = chunks.length;
  const maxChars = Math.max(...chunks.map(c => c.length));

  // スペーサー列の有無を判定
  const totalDisplayCols = table.tableRows[0].tableCells.length;
  const hasSpacerCol     = totalDisplayCols > numCols;

  const styleReqs = [];

  for (const chunkIdx of chunkIndices) {
    if (chunkIdx < 0 || chunkIdx >= numCols) continue;

    // chunkIdx → 表示列インデックス変換
    const contentColIdx = numCols - 1 - chunkIdx;
    const displayCol    = contentColIdx + (hasSpacerCol ? 1 : 0);

    const chunk    = chunks[chunkIdx];
    const chunkLen = chunk.length;

    // 均等割付の字間計算（resetMode=true のときは常に 0）
    const extraSpace = resetMode ? 0 : (maxChars - chunkLen) * fSize;
    const numGaps    = Math.max(chunkLen - 1, 1);
    const gapBetween = (!resetMode && chunkLen > 1) ? extraSpace / numGaps : 0;

    const tableCell   = table.tableRows[0].tableCells[displayCol];
    const paragraphs  = tableCell.content || [];
    const lastParaIdx = paragraphs.length - 1;
    let charIdx = 0;

    for (let pi = 0; pi < paragraphs.length; pi++) {
      const contentEl  = paragraphs[pi];
      if (!contentEl.paragraph) continue;
      const isTrailing = pi === lastParaIdx;

      const spaceAbovePt = (!isTrailing && charIdx > 0) ? gapBetween : 0;

      styleReqs.push({
        updateParagraphStyle: {
          range: { startIndex: contentEl.startIndex, endIndex: contentEl.endIndex },
          paragraphStyle: {
            spaceAbove: { magnitude: spaceAbovePt, unit: 'PT' }
          },
          fields: 'spaceAbove'
        }
      });

      if (!isTrailing) charIdx++;
    }
  }

  if (styleReqs.length) await batchUpdate(token, docId, styleReqs);
  return { success: true, appliedCols: chunkIndices.length };
}

// ─── ドキュメント内の FVdoc テーブルを一覧取得 ──────────────────────
async function loadFvdocTables(docId) {
  const token = await getToken();
  const doc   = await docsGet(token, docId);
  const results = [];

  for (const el of doc.body.content || []) {
    if (!el.table) continue;
    const table   = el.table;
    const lastRow = table.tableRows[table.tableRows.length - 1];
    if (!lastRow) continue;

    // col1（新形式）→ col0（旧形式）の順でメタデータを探す
    let cellText = '';
    const candidates = [lastRow.tableCells[1], lastRow.tableCells[0]].filter(Boolean);
    for (const cell of candidates) {
      let t = '';
      for (const contentEl of cell.content || []) {
        if (contentEl.paragraph) {
          for (const pe of contentEl.paragraph.elements || []) {
            if (pe.textRun) t += pe.textRun.content;
          }
        }
      }
      if (t.includes(FVDOC_MARKER)) { cellText = t; break; }
    }

    if (cellText.includes(FVDOC_MARKER)) {
      const jsonStr = cellText
        .substring(cellText.indexOf(FVDOC_MARKER) + FVDOC_MARKER.length)
        .replace(/\n/g, '');
      try {
        const meta = JSON.parse(jsonStr);
        results.push({
          tableIndex: el.startIndex,
          preview: (meta.originalText || '').substring(0, 20),
          meta
        });
      } catch (_) {}
    }
  }
  return results;
}
