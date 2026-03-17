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
    .replace(/。/g, '\uFE12')  // ︒ 縦書き句点
    .replace(/、/g, '\uFE11')  // ︑ 縦書き読点
    .replace(/（/g, '\uFE35')  // ︵ 縦書き左括弧
    .replace(/）/g, '\uFE36')  // ︶ 縦書き右括弧
    .replace(/ー/g, '\uFF5C')  // ｜ 縦書き長音符
    .replace(/「/g, '\uFE41')  // ﹁ 縦書き左鉤括弧
    .replace(/」/g, '\uFE42')  // ﹂ 縦書き右鉤括弧
    .replace(/『/g, '\uFE43')  // ﹃ 縦書き左二重鉤括弧
    .replace(/』/g, '\uFE44')  // ﹄ 縦書き右二重鉤括弧
    .replace(/\[/g,  '\uFE47') // ﹇ 縦書き左角括弧
    .replace(/\]/g,  '\uFE48') // ﹈ 縦書き右角括弧
    .replace(/：/g,  '\u2025') // ‥ 2点リーダー（全角コロン）
    .replace(/:/g,   '\u2025') // ‥ 2点リーダー（半角コロン）
    .replace(/；/g,  '\uFE14') // ︔ 縦書きセミコロン（全角）
    .replace(/;/g,   '\uFE14') // ︔ 縦書きセミコロン（半角）
    .replace(/！/g,  '\uFE15') // ︕ 縦書き感嘆符（全角）
    .replace(/!/g,   '\uFE15') // ︕ 縦書き感嘆符（半角）
    .replace(/？/g,  '\uFE16') // ︖ 縦書き疑問符（全角）
    .replace(/\?/g,  '\uFE16'); // ︖ 縦書き疑問符（半角）
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

// ─── 列チャンク構築 ───────────────────────────────────────────────
function buildChunks(text, charsPerLine) {
  const inputLines = text.split(/\r?\n/).filter(l => l.trim().length > 0);
  const chunks = [];
  for (const line of inputLines) {
    const converted = replaceSymbols(line);
    if (!converted) continue;
    for (let i = 0; i < converted.length; i += charsPerLine) {
      chunks.push(converted.substring(i, i + charsPerLine));
    }
  }
  return chunks;
}

// ─── メインロジック：テーブルの挿入 ──────────────────────────────
async function handleInsert(docId, params) {
  const { text, charsPerLine, fontSize, fontFamily } = params;
  const cpLine = parseInt(charsPerLine);
  const fSize  = parseFloat(fontSize);
  const token  = await getToken();

  const chunks = buildChunks(text, cpLine);
  if (!chunks.length) throw new Error('テキストがありません');

  const numCols  = chunks.length;

  // ── STEP 1: ドキュメント情報取得 ──────────────────────────────
  const doc = await docsGet(token, docId);
  const bodyContent  = doc.body.content;
  const insertAt     = bodyContent[bodyContent.length - 1].startIndex;

  const colWidthPt   = fSize + 2;
  const totalCols    = numCols + 1; // +1 = スペーサー列

  // ── STEP 2: 空テーブルを挿入 ─────────────────────────────────
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
  const colWidthReqs = [{
    updateTableColumnProperties: {
      tableStartLocation: { index: newTableStart },
      columnIndices: [0],
      tableColumnProperties: { widthType: 'EVENLY_DISTRIBUTED' },
      fields: 'widthType'
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

  for (let col = 1; col < totalCols; col++) {
    const contentColIdx = col - 1;
    const chunkIdx      = numCols - 1 - contentColIdx;
    const chunk         = chunks[chunkIdx];
    const cellText      = chunk.split('').join('\n');

    const tableCell = newTable.tableRows[0].tableCells[col];
    const firstEl   = tableCell.content?.[0]?.paragraph?.elements?.[0];
    if (!firstEl) continue;
    insertions.push({ index: firstEl.startIndex, text: cellText });
  }

  const metaJson    = JSON.stringify({ originalText: text, charsPerLine: cpLine, fontSize: fSize, fontFamily });
  const metaCell    = newTable.tableRows[1].tableCells[0];
  const metaFirstEl = metaCell.content?.[0]?.paragraph?.elements?.[0];
  if (metaFirstEl) {
    insertions.push({ index: metaFirstEl.startIndex, text: FVDOC_MARKER + metaJson });
  }

  insertions.sort((a, b) => b.index - a.index);
  await batchUpdate(token, docId, insertions.map(({ index, text: t }) => ({
    insertText: { location: { index }, text: t }
  })));

  // ── STEP 5: スタイル適用 ──────────────────────────────────────
  const styledDoc    = await docsGet(token, docId);
  const styledTables = styledDoc.body.content.filter(el => el.table);
  const styledTableEl = styledTables[styledTables.length - 1];
  const styledTable   = styledTableEl.table;
  const tableStartIndex = styledTableEl.startIndex;

  const styleReqs = [];
  const white      = { color: { rgbColor: { red: 1, green: 1, blue: 1 } } };
  const whiteBorder = { color: white, width: { magnitude: 1, unit: 'PT' }, dashStyle: 'SOLID' };
  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < totalCols; col++) {
      const isSpacerCol = (col === 0);
      const tableCell   = styledTable.tableRows[row].tableCells[col];

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
            paddingLeft:   { magnitude: isSpacerCol ? 0 : 1, unit: 'PT' },
            paddingRight:  { magnitude: isSpacerCol ? 0 : 1, unit: 'PT' },
            contentAlignment: 'TOP'
          },
          fields: 'borderTop,borderBottom,borderLeft,borderRight,paddingTop,paddingBottom,paddingLeft,paddingRight,contentAlignment'
        }
      });

      if (row === 0 && !isSpacerCol) {
        const contentColIdx = col - 1;
        const chunkIdx      = numCols - 1 - contentColIdx;
        const chunkLen      = chunks[chunkIdx].length;

        const paragraphs = tableCell.content || [];
        const lastParaIdx = paragraphs.length - 1;
        let charIdx = 0;

        for (let pi = 0; pi < paragraphs.length; pi++) {
          const contentEl = paragraphs[pi];
          if (!contentEl.paragraph) continue;

          const paraStart  = contentEl.startIndex;
          const paraEnd    = contentEl.endIndex;
          const isTrailing = pi === lastParaIdx;

          styleReqs.push({
            updateParagraphStyle: {
              range: { startIndex: paraStart, endIndex: paraEnd },
              paragraphStyle: {
                alignment: 'CENTER',
                lineSpacing: 70,
                spaceAbove: { magnitude: 0, unit: 'PT' },
                spaceBelow: { magnitude: 0, unit: 'PT' }
              },
              fields: 'alignment,lineSpacing,spaceAbove,spaceBelow'
            }
          });

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
        for (const contentEl of tableCell.content || []) {
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
  return { success: true, tableStartIndex, numCols };
}

// ─── 均等割付：指定列の字間を再計算して適用 ──────────────────────────
async function applyColumnJustify(docId, tableStartIndex, chunkIndices, resetMode = false) {
  const token = await getToken();
  const doc   = await docsGet(token, docId);

  const tableEl = doc.body.content.find(el => el.table && el.startIndex === tableStartIndex);
  if (!tableEl) throw new Error('指定されたテーブルが見つかりませんでした');
  const table = tableEl.table;

  const lastRow   = table.tableRows[table.tableRows.length - 1];
  let cellText = '';
  for (const contentEl of lastRow.tableCells[0].content || []) {
    for (const pe of contentEl.paragraph?.elements || []) {
      if (pe.textRun) cellText += pe.textRun.content;
    }
  }
  if (!cellText.includes(FVDOC_MARKER)) throw new Error('FVdocメタデータが見つかりません');

  const jsonStr = cellText
    .substring(cellText.indexOf(FVDOC_MARKER) + FVDOC_MARKER.length)
    .replace(/\n/g, '');
  const meta = JSON.parse(jsonStr);
  const { originalText, charsPerLine, fontSize: fSizeRaw } = meta;
  const fSize = parseFloat(fSizeRaw);

  const chunks  = buildChunks(originalText, parseInt(charsPerLine));
  const numCols = chunks.length;
  const maxChars = Math.max(...chunks.map(c => c.length));

  const totalDisplayCols = table.tableRows[0].tableCells.length;
  const hasSpacerCol     = totalDisplayCols > numCols;

  const styleReqs = [];

  for (const chunkIdx of chunkIndices) {
    if (chunkIdx < 0 || chunkIdx >= numCols) continue;

    const contentColIdx = numCols - 1 - chunkIdx;
    const displayCol    = contentColIdx + (hasSpacerCol ? 1 : 0);

    const chunk    = chunks[chunkIdx];
    const chunkLen = chunk.length;

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
    const firstCell = lastRow.tableCells[0];
    if (!firstCell) continue;

    let cellText = '';
    for (const contentEl of firstCell.content || []) {
      if (contentEl.paragraph) {
        for (const pe of contentEl.paragraph.elements || []) {
          if (pe.textRun) cellText += pe.textRun.content;
        }
      }
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
