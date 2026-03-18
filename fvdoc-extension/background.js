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
 *    ※ スペーサー列あり → totalCols = numPageCols + 1
 *
 *  改ページ:
 *    1ページに収まる列数を計算し、超過分は insertSectionBreak(NEXT_PAGE) で
 *    次ページに続くテーブルを挿入する
 *
 *  均等割付:
 *    各列の文字数が最長列より短い場合、spaceAbove で字間を広げて高さを揃える
 */

'use strict';

const DOCS_API     = 'https://docs.googleapis.com/v1/documents';
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
    .replace(/\u2015/g,  '\uFE31') // ― → ︱ 縦書き長ダッシュ（水平バー）
    .replace(/\u2014/g,  '\uFE31') // — → ︱ 縦書き長ダッシュ（エムダッシュ）
    .replace(/\u2013/g,  '\uFE32') // – → ︲ 縦書き短ダッシュ（エンダッシュ）
    .replace(/-/g,       '\uFE32') // - → ︲ 縦書き短ダッシュ（ハイフン）
    .replace(/\uFF5E/g,  '\u301C') // ～ → 〜 波ダッシュ
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
function applyKinsoku(chunks) {
  const forbidden = new Set(['\uFE12', '\uFE11']); // ︒ ︑
  const result = [...chunks];
  let i = 1;
  while (i < result.length) {
    while (result[i].length > 0 && forbidden.has(result[i][0])) {
      result[i - 1] += result[i][0];
      result[i]      = result[i].substring(1);
    }
    if (result[i].length === 0) {
      result.splice(i, 1);
    } else {
      i++;
    }
  }
  return result;
}

// ─── 列チャンク構築 ───────────────────────────────────────────────
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

// ─── 1ページ分のテーブルを挿入・設定する ─────────────────────────
// pageChunks : このページに配置するチャンク配列（chunks[0]=右端＝最初に読む列）
// 列構成（左→右）: col0=スペーサー | col1=左端の列(最後に読む) … col numPageCols=右端の列(最初に読む)
// 縦書きは右から左へ読むので、画面では右端が col numPageCols = chunks[0]
// isFirstPage: true のときメタデータを col1 に挿入する
// useEndOfSegment: true のとき文書末尾（改ページの後）に挿入（2ページ目以降で謎の空白を防ぐ）
async function insertPageTable(token, docId, pageChunks, {
  cpLine, fSize, lSpacing, colGap, cellPadH, colWidthPt, usableWidthPt,
  isFirstPage, metaJson
}) {
  const numPageCols   = pageChunks.length;
  const totalPageCols = numPageCols + 1; // col0 = スペーサー（左端・右寄せ用）

  const colMagnitude = Math.round(colWidthPt);

  // スペーサー幅 = ページ有効幅 - コンテンツ列幅合計
  // EVENLY_DISTRIBUTED は Google Docs が最小幅(~70pt)を強制するため FIXED_WIDTH に変更
  // 最小5ptを確保（0/負になるのを防ぐ）
  const spacerMagnitude = Math.max(5, Math.round(usableWidthPt) - numPageCols * colMagnitude);

  function makeColWidthReqs(tableStartIdx) {
    // col0: FIXED_WIDTH = 計算済みスペーサー幅（右寄せ用）
    const reqs = [{
      updateTableColumnProperties: {
        tableStartLocation: { index: tableStartIdx },
        columnIndices: [0],
        tableColumnProperties: {
          widthType: 'FIXED_WIDTH',
          width: { magnitude: spacerMagnitude, unit: 'PT' }
        },
        fields: 'width,widthType'
      }
    }];
    // col1〜: FIXED_WIDTH = fSize + 左右パディング（列間込み）
    for (let col = 1; col < totalPageCols; col++) {
      reqs.push({
        updateTableColumnProperties: {
          tableStartLocation: { index: tableStartIdx },
          columnIndices: [col],
          tableColumnProperties: {
            widthType: 'FIXED_WIDTH',
            width: { magnitude: colMagnitude, unit: 'PT' }
          },
          fields: 'width,widthType'
        }
      });
    }
    return reqs;
  }

  // ── 挿入位置を取得（endIndex - 1 = ドキュメント末尾の有効な最後のインデックス）
  const currentDoc = await docsGet(token, docId);
  const lastEl     = currentDoc.body.content[currentDoc.body.content.length - 1];
  const insertAt   = lastEl.endIndex - 1;

  // ── テーブル挿入
  await batchUpdate(token, docId, [{
    insertTable: {
      rows: 2,
      columns: totalPageCols,
      location: { index: insertAt }
    }
  }]);

  // ── テーブル情報取得（実際の startIndex を使用）
  const updatedDoc    = await docsGet(token, docId);
  const tableElements = updatedDoc.body.content.filter(el => el.table);
  const newTableEl    = tableElements[tableElements.length - 1];
  const newTable      = newTableEl.table;
  const newTableStart = newTableEl.startIndex;

  // ── 列幅設定
  await batchUpdate(token, docId, makeColWidthReqs(newTableStart));

  // ── テキスト挿入（col1=左端列 chunks[n-1] … col numPageCols=右端列 chunks[0]、右から左へ読む） ──
  const insertions = [];

  for (let col = 1; col < totalPageCols; col++) {
    const contentColIdx = col - 1;
    const pageChunkIdx  = numPageCols - 1 - contentColIdx; // 右端=chunks[0] → 最後の列に
    const chunk         = pageChunks[pageChunkIdx];
    const cellText      = chunk.split('').join('\n');

    const tableCell = newTable.tableRows[0].tableCells[col];
    const firstEl   = tableCell.content?.[0]?.paragraph?.elements?.[0];
    if (!firstEl) continue;
    insertions.push({ index: firstEl.startIndex, text: cellText });
  }

  // メタデータは最初のページの col1（2行目）にのみ格納
  if (isFirstPage && metaJson) {
    const metaCell    = newTable.tableRows[1].tableCells[1];
    const metaFirstEl = metaCell.content?.[0]?.paragraph?.elements?.[0];
    if (metaFirstEl) {
      insertions.push({ index: metaFirstEl.startIndex, text: FVDOC_MARKER + metaJson });
    }
  }

  insertions.sort((a, b) => b.index - a.index);
  await batchUpdate(token, docId, insertions.map(({ index, text: t }) => ({
    insertText: { location: { index }, text: t }
  })));

  // ── スタイル適用 ────────────────────────────────────────────────
  const styledDoc     = await docsGet(token, docId);
  const styledTables  = styledDoc.body.content.filter(el => el.table);
  const styledTableEl  = styledTables[styledTables.length - 1];
  const styledTable    = styledTableEl.table;
  const tableStartIndex = styledTableEl.startIndex;

  const styleReqs  = [];
  const white       = { color: { rgbColor: { red: 1, green: 1, blue: 1 } } };
  const whiteBorder = { color: white, width: { magnitude: 1, unit: 'PT' }, dashStyle: 'SOLID' };

  for (let row = 0; row < 2; row++) {
    for (let col = 0; col < totalPageCols; col++) {
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
            paddingLeft:   { magnitude: isSpacerCol ? 0 : cellPadH, unit: 'PT' },
            paddingRight:  { magnitude: isSpacerCol ? 0 : cellPadH, unit: 'PT' },
            contentAlignment: 'TOP'
          },
          fields: 'borderTop,borderBottom,borderLeft,borderRight,paddingTop,paddingBottom,paddingLeft,paddingRight,contentAlignment'
        }
      });

      if (row === 0 && !isSpacerCol) {
        const contentColIdx = col - 1;
        const pageChunkIdx  = numPageCols - 1 - contentColIdx;
        const chunkLen      = pageChunks[pageChunkIdx].length;

        const paragraphs  = tableCell.content || [];
        const lastParaIdx = paragraphs.length - 1;
        let charIdx = 0;

        for (let pi = 0; pi < paragraphs.length; pi++) {
          const contentEl  = paragraphs[pi];
          if (!contentEl.paragraph) continue;

          const paraStart  = contentEl.startIndex;
          const paraEnd    = contentEl.endIndex;
          const isTrailing = pi === lastParaIdx;

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

          for (const pe of contentEl.paragraph.elements || []) {
            if (!pe.textRun) continue;
            const textStyle = { fontSize: { magnitude: fSize, unit: 'PT' } };
            let fields = 'fontSize';
            if (pageChunks._fontFamily && pageChunks._fontFamily !== 'default') {
              textStyle.weightedFontFamily = { fontFamily: pageChunks._fontFamily };
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
        // メタデータ行：1pt白文字・行間最小化
        for (const contentEl of tableCell.content || []) {
          if (!contentEl.paragraph) continue;

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

  // スタイル適用後に全列の幅を再適用（Docs が列幅を上書きすることがあるため、スペーサー含め確実に効かせる）
  await batchUpdate(token, docId, makeColWidthReqs(tableStartIndex));

  return tableStartIndex;
}

// ─── メインロジック：テーブルの挿入（改ページ対応） ──────────────
async function handleInsert(docId, params) {
  const { text, charsPerLine, fontSize, fontFamily, lineSpacingPct, colGapPt } = params;
  const cpLine   = parseInt(charsPerLine);
  const fSize    = parseFloat(fontSize);
  const lSpacing = parseFloat(lineSpacingPct) || 70;
  const colGap   = parseFloat(colGapPt)       || 0;
  const token    = await getToken();

  const chunks = buildChunks(text, cpLine);
  if (!chunks.length) throw new Error('テキストがありません');

  // ── ページ寸法とスペーサー計算 ────────────────────────────────
  const doc            = await docsGet(token, docId);
  const docStyle       = doc.documentStyle || {};
  const pageHeightPt   = toPt(docStyle.pageSize?.height)  || 842;
  const marginTopPt    = toPt(docStyle.marginTop)          || 72;
  const marginBottomPt = toPt(docStyle.marginBottom)       || 72;
  const usableHeightPt = pageHeightPt - marginTopPt - marginBottomPt;

  // ページ幅（水平分割の判定に使用）
  const pageWidthPt     = toPt(docStyle.pageSize?.width)  || 595;
  const marginLeftPt    = toPt(docStyle.marginLeft)        || 72;
  const marginRightPt   = toPt(docStyle.marginRight)       || 72;
  const rawUsableWidthPt = pageWidthPt - marginLeftPt - marginRightPt;
  // 有効幅が異常に小さい場合（200pt未満 ≈ A5未満）はデフォルトマージン(各72pt)で再計算
  const usableWidthPt   = (rawUsableWidthPt >= 200) ? rawUsableWidthPt : Math.max(200, pageWidthPt - 144);

  // デバッグ: サービスワーカーのコンソールで確認可能
  // chrome://extensions → FVdoc → 「Service Worker」リンク → Console タブ
  console.log('[FVdoc] page:', {
    pageWidthPt, marginLeftPt, marginRightPt,
    rawUsableWidthPt, usableWidthPt
  });

  // 列間(colGap)を反映: 各セル左右パディング = 1 + 列間/2 → 隣接列の間隔 = 列間
  const cellPadH   = 1 + colGap / 2;
  const colWidthPt = fSize + cellPadH * 2; // 1列の幅 = フォント幅 + 左右パディング（列間込み）

  // 1ページに収まる行数（縦書きセル内の「行」＝1文字＝1段落の高さ）
  const lineHeightPt    = fSize * (lSpacing / 100);
  const maxLinesPerPage = Math.max(1, Math.floor(usableHeightPt / lineHeightPt));

  // 水平方向の列数ページ分割
  //
  //   Google Docs はテーブル列幅合計がページ幅を超えると全列を自動スケーリングする。
  //   colGap=0 の場合：圧縮されても列間は元々0なので視覚的に問題なし → ~1.33倍まで許容
  //   colGap>0 の場合：圧縮されると列間(colGap)も潰れて0になってしまう → 許容しない(1.0倍)
  //
  //   1.33 の根拠: colGap=0 で実測した際に Docs が許容できたスケーリング比率の上限
  const SCALE_TOLERANCE   = (colGap > 0) ? 1.0 : 1.33;
  const totalContentWidth = chunks.length * colWidthPt;
  let   columnsPerPage;
  if (totalContentWidth <= usableWidthPt * SCALE_TOLERANCE) {
    // スケーリング許容範囲内 → 全列を1ページに
    columnsPerPage = chunks.length;
  } else {
    // 許容超過 → 1ページ最大列数で切り分け（1ページ目を先に埋める）
    columnsPerPage = Math.max(1, Math.floor(usableWidthPt / colWidthPt));
  }

  // チャンクをページ単位に分割（超過の場合のみ複数グループ）
  const pageChunkGroups = [];
  for (let i = 0; i < chunks.length; i += columnsPerPage) {
    const group = chunks.slice(i, i + columnsPerPage);
    group._fontFamily = fontFamily;
    pageChunkGroups.push(group);
  }

  const metaJson = JSON.stringify({
    originalText: text, charsPerLine: cpLine, fontSize: fSize,
    fontFamily, lineSpacingPct: lSpacing, colGapPt: colGap
  });

  let firstTableStartIndex = null;

  // 高さベースの改ページ管理:
  //   remainingHeightPt = 現在ページの残り有効高さ
  //   テーブルが収まる場合は同じページに続けて配置（セクション区切りなし）
  //   収まらない場合のみ NEXT_PAGE セクション区切りを挿入して新ページへ
  let remainingHeightPt = usableHeightPt;
  let isFirstTable = true;

  for (let pageIdx = 0; pageIdx < pageChunkGroups.length; pageIdx++) {
    const pageChunks  = pageChunkGroups[pageIdx];
    const numPageCols = pageChunks.length;
    const maxChunkLen = Math.max(...pageChunks.map(c => c.length), 1);
    let numHeightSlices = Math.ceil(maxChunkLen / maxLinesPerPage);

    // 最後のスライスが極端に短い場合は前のテーブルに含め、空白ページを防ぐ
    if (numHeightSlices > 1) {
      const lastSliceLines = maxChunkLen - (numHeightSlices - 1) * maxLinesPerPage;
      if (lastSliceLines < maxLinesPerPage * 0.25) {
        numHeightSlices -= 1;
      }
    }

    for (let sliceIdx = 0; sliceIdx < numHeightSlices; sliceIdx++) {
      const lineStart = sliceIdx * maxLinesPerPage;
      const lineEnd   = (sliceIdx === numHeightSlices - 1)
        ? maxChunkLen
        : Math.min(lineStart + maxLinesPerPage, maxChunkLen);
      const slicedChunks = pageChunks.map(chunk =>
        chunk.substring(lineStart, Math.min(lineEnd, chunk.length))
      );
      const hasContent = slicedChunks.some(c => c.length > 0);
      if (!hasContent) continue;

      // このスライスのテーブル高さ（行数 × 行高さ）
      const sliceLines     = lineEnd - lineStart;
      const sliceHeightPt  = sliceLines * lineHeightPt;

      if (!isFirstTable) {
        if (sliceHeightPt > remainingHeightPt) {
          // 現在ページに収まらない → 改ページ（セクション区切り）
          const dBreak  = await docsGet(token, docId);
          const breakAt = dBreak.body.content[dBreak.body.content.length - 1].endIndex - 1;
          await batchUpdate(token, docId, [{
            insertSectionBreak: {
              location:    { index: breakAt },
              sectionType: 'NEXT_PAGE'
            }
          }]);
          remainingHeightPt = usableHeightPt; // 新ページの残り高さをリセット
        }
        // 収まる場合はそのまま続けて配置（セクション区切りなし）
      }

      slicedChunks._fontFamily = pageChunks._fontFamily;

      const tableStartIndex = await insertPageTable(token, docId, slicedChunks, {
        cpLine, fSize, lSpacing, colGap, cellPadH, colWidthPt, usableWidthPt,
        isFirstPage: isFirstTable,
        metaJson: isFirstTable ? metaJson : null
      });

      if (isFirstTable) firstTableStartIndex = tableStartIndex;
      isFirstTable = false;

      // このテーブルを配置した後の残り高さを更新
      remainingHeightPt = Math.max(0, remainingHeightPt - sliceHeightPt);
    }
  }

  console.log('[FVdoc] result:', { columnsPerPage, numPages: pageChunkGroups.length, numCols: chunks.length, usableWidthPt, colWidthPt });
  return { success: true, tableStartIndex: firstTableStartIndex, numCols: chunks.length,
    _debug: `usableW=${Math.round(usableWidthPt)}pt colW=${colWidthPt}pt cols/pg=${columnsPerPage} pages=${pageChunkGroups.length}` };
}

// ─── 均等割付 ─────────────────────────────────────────────────────
async function applyColumnJustify(docId, tableStartIndex, chunkIndices, resetMode = false) {
  const token = await getToken();
  const doc   = await docsGet(token, docId);

  const tableEl = doc.body.content.find(el => el.table && el.startIndex === tableStartIndex);
  if (!tableEl) throw new Error('指定されたテーブルが見つかりませんでした');
  const table = tableEl.table;

  // メタデータ取得（col1 優先、col0 フォールバック）
  const lastRow        = table.tableRows[table.tableRows.length - 1];
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
  if (!cellText.includes(FVDOC_MARKER)) throw new Error('FVdocメタデータが見つかりません');

  const jsonStr = cellText
    .substring(cellText.indexOf(FVDOC_MARKER) + FVDOC_MARKER.length)
    .replace(/\n/g, '');
  const meta = JSON.parse(jsonStr);
  const { originalText, charsPerLine, fontSize: fSizeRaw } = meta;
  const fSize = parseFloat(fSizeRaw);

  const chunks       = buildChunks(originalText, parseInt(charsPerLine));
  const numCols      = chunks.length;
  const maxChars     = Math.max(...chunks.map(c => c.length));

  const totalDisplayCols = table.tableRows[0].tableCells.length;
  const hasSpacerCol     = totalDisplayCols > numCols;

  // このテーブルが持つ列数（ページ1のみ対象）
  const pageNumCols = totalDisplayCols - (hasSpacerCol ? 1 : 0);

  const styleReqs = [];

  for (const chunkIdx of chunkIndices) {
    if (chunkIdx < 0 || chunkIdx >= numCols) continue;

    // このテーブルに含まれる列のみ（chunkIdx=0 が右端列＝col numPageCols）
    const contentColIdx = numCols - 1 - chunkIdx;
    if (contentColIdx >= pageNumCols) continue;

    const displayCol = contentColIdx + (hasSpacerCol ? 1 : 0);

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
