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

// ★ バージョン確認用: chrome://extensions → FVdoc → Service Worker → Console で確認
console.log('[FVdoc v4] ★★★ Service Worker 起動 ★★★', new Date().toISOString());

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
    // 元から空のチャンク（空行由来）はスキップ：禁則処理対象外かつ削除しない
    const originallyEmpty = result[i].length === 0;
    while (result[i].length > 0 && forbidden.has(result[i][0])) {
      result[i - 1] += result[i][0];
      result[i]      = result[i].substring(1);
    }
    // 禁則処理の結果として空になったチャンクだけ削除（空行由来の空チャンクは保持）
    if (result[i].length === 0 && !originallyEmpty) {
      result.splice(i, 1);
    } else {
      i++;
    }
  }
  return result;
}

// ─── 列チャンク構築 ───────────────────────────────────────────────
function buildChunks(text, charsPerLine) {
  const inputLines = text.split(/\r?\n/);
  const rawChunks = [];
  for (const line of inputLines) {
    if (line.trim().length === 0) {
      // 空行 → 空のチャンク（縦書きで1列空ける）
      rawChunks.push('');
      continue;
    }
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
  columnsPerPage, isFirstPage, metaJson
}) {
  const numPageCols   = pageChunks.length;
  // totalPageCols = 常に columnsPerPage（ページ1と同じ列数）
  // 先頭 numSpacerCols 列は空セル（右寄せ用スペーサー）、全列 EVENLY_DISTRIBUTED
  // → Google Docs が usableWidthPt / columnsPerPage を全ページ均等に配分
  const totalPageCols = columnsPerPage;
  const numSpacerCols = totalPageCols - numPageCols;

  console.log('[FVdoc] insertPageTable:', {
    numPageCols, totalPageCols, numSpacerCols,
    colWidthPt, usableWidthPt: Math.round(usableWidthPt),
    estColWidth: Math.round(usableWidthPt / totalPageCols * 10) / 10
  });
  // ─────────────────────────────────────────────────────────────────────────────

  function makeColWidthReqs(tableStartIdx) {
    // 全列 EVENLY_DISTRIBUTED: totalPageCols は常に columnsPerPage で一定
    // → Google Docs が usableWidthPt を totalPageCols 等分
    // → 全ページで列幅が一致し、FIXED_WIDTH によるオーバー/アンダー問題を完全回避
    // 先頭 numSpacerCols 列は空セルで右寄せ効果（FIXED_WIDTH スペーサー不要）
    const estColWidth = Math.round(usableWidthPt / totalPageCols * 10) / 10;
    console.log('[FVdoc] makeColWidthReqs(ALL EVENLY_DISTRIBUTED): ' + JSON.stringify({
      totalPageCols, numSpacerCols, numPageCols,
      estColWidth,
      usableWidthPt: Math.round(usableWidthPt)
    }));

    const allIndices = [];
    for (let col = 0; col < totalPageCols; col++) allIndices.push(col);

    return [{
      updateTableColumnProperties: {
        tableStartLocation: { index: tableStartIdx },
        columnIndices: allIndices,
        tableColumnProperties: {
          widthType: 'EVENLY_DISTRIBUTED'
        },
        fields: 'widthType'
      }
    }];
  }

  // ── 挿入位置を取得（endIndex - 1 = ドキュメント末尾の有効な最後のインデックス）
  const currentDoc = await docsGet(token, docId);
  const lastEl     = currentDoc.body.content[currentDoc.body.content.length - 1];
  const insertAt   = lastEl.endIndex - 1;

  // ── 挿入位置の段落インデントを診断（テーブル有効幅に影響する可能性がある）
  try {
    const ps = lastEl.paragraph?.paragraphStyle || {};
    // namedStyles から NORMAL_TEXT の継承インデントも確認
    const nsNormal = currentDoc.namedStyles?.styles?.find(s => s.namedStyleType === 'NORMAL_TEXT');
    const nsIndent = nsNormal?.paragraphStyle?.indentStart ? toPt(nsNormal.paragraphStyle.indentStart) : null;
    console.log('[FVdoc] 挿入段落スタイル: ' + JSON.stringify({
      namedStyleType:  ps.namedStyleType,
      indentStart:     ps.indentStart    ? toPt(ps.indentStart)    : null,
      indentEnd:       ps.indentEnd      ? toPt(ps.indentEnd)      : null,
      indentFirstLine: ps.indentFirstLine? toPt(ps.indentFirstLine): null,
      namedStyleIndentStart: nsIndent,    // ← 継承インデント（null以外なら原因）
      alignment:       ps.alignment,
      insertAt,
      lastElType: lastEl.paragraph ? 'PARA' : (lastEl.table ? 'TABLE' : 'OTHER')
    }));
  } catch (e) {
    console.error('[FVdoc] 段落スタイル診断エラー:', e.message);
  }

  // ── 挿入前にインデント・段落行間をリセット（namedStyle 継承インデントを上書き）
  // テーブルは段落の有効幅（ページ幅 − マージン − インデント）に制約されるため、
  // インデントをゼロリセット。行間・前後スペースもゼロにして段落の高さを最小化する。
  //
  // ⚠️ updateTextStyle（fontSize/foregroundColor）はここで適用しない。
  //    テーブル挿入前にカーソル位置の textStyle を変更すると、挿入後の新規セルが
  //    その書式を継承してしまい（1pt 白文字）文字がすべて不可視になるため。
  //    → テキストスタイルはテーブル挿入「後」に auto_para に適用する（下記参照）。
  // ── 段落スタイルリセット ＋ テーブル挿入（1回の batchUpdate に統合）
  const preInsertReqs = lastEl.paragraph ? [
    {
      updateParagraphStyle: {
        range: { startIndex: insertAt, endIndex: insertAt + 1 },
        paragraphStyle: {
          indentStart:     { magnitude: 0, unit: 'PT' },
          indentEnd:       { magnitude: 0, unit: 'PT' },
          indentFirstLine: { magnitude: 0, unit: 'PT' },
          lineSpacing:     100,
          spaceAbove:      { magnitude: 0, unit: 'PT' },
          spaceBelow:      { magnitude: 0, unit: 'PT' }
        },
        fields: 'indentStart,indentEnd,indentFirstLine,lineSpacing,spaceAbove,spaceBelow'
      }
    }
  ] : [];
  await batchUpdate(token, docId, [...preInsertReqs, {
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

  // ── insertTable 直後の初期列幅を診断（updateTableColumnProperties 適用前）
  try {
    const initColProps = newTable.tableStyle?.tableColumnProperties || [];
    const initMapped   = initColProps.map(p => ({ wt: p.widthType, mag: p.width?.magnitude, u: p.width?.unit }));
    const initTotal    = initMapped.reduce((s, p) => s + (p.mag || 0), 0);
    console.log('[FVdoc] 初期列幅(insertTable直後): ' + JSON.stringify({
      numCols: initColProps.length,
      totalInitial: Math.round(initTotal),
      col0: initMapped[0] || null,
      col1: initMapped[1] || null,
      usableWidthPt: Math.round(usableWidthPt)
    }));
  } catch (e) {
    console.error('[FVdoc] 初期列幅診断エラー:', e.message);
  }

  // ── 列幅設定 + テーブル前 auto_para の文字スタイルを最小化（1pt 白文字）
  // ⚠️ updateTextStyle はテーブル挿入「後」に適用する。
  //    挿入「前」に設定するとセルが書式を継承して文字がすべて不可視になる。
  // auto_para はテーブル直前の 1文字（newline）なので範囲 = {newTableStart-1, newTableStart}
  const autoParaWhite = { color: { rgbColor: { red: 1, green: 1, blue: 1 } } };
  const autoParaTextStyleReqs = (newTableStart > 1) ? [{
    updateTextStyle: {
      range: { startIndex: newTableStart - 1, endIndex: newTableStart },
      textStyle: {
        fontSize:        { magnitude: 1, unit: 'PT' },
        foregroundColor: autoParaWhite
      },
      fields: 'fontSize,foregroundColor'
    }
  }] : [];
  await batchUpdate(token, docId, [...autoParaTextStyleReqs, ...makeColWidthReqs(newTableStart)]);

  // ── テキスト挿入（col=numSpacerCols が左端コンテンツ列=最終チャンク、col=totalPageCols-1 が右端=最初チャンク）
  // col 0〜numSpacerCols-1 は空スペーサー列（何も挿入しない）
  const insertions = [];

  for (let col = numSpacerCols; col < totalPageCols; col++) {
    const contentColIdx = col - numSpacerCols;
    const pageChunkIdx  = numPageCols - 1 - contentColIdx; // 右端=chunks[0] → 最後の列に
    const chunk         = pageChunks[pageChunkIdx];

    // 空行由来の空チャンクはテキスト挿入をスキップ（insertText は空文字列不可）
    // → セルは空のまま残り、縦書きレイアウトで1列分の空白として機能する
    if (!chunk || chunk.length === 0) continue;

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
      const isSpacerCol = (col < numSpacerCols);
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
        const contentColIdx = col - numSpacerCols;
        const pageChunkIdx  = numPageCols - 1 - contentColIdx;
        const chunkLen      = pageChunks[pageChunkIdx].length;

        const paragraphs = tableCell.content || [];
        if (paragraphs.length === 0) continue;

        // セル全体を1つの範囲として一括更新（1文字ごとの個別更新から変更）
        // 全段落が同じスタイルを持つため、セル先頭〜末尾の範囲1つで適用可能。
        // これにより styleReqs のリクエスト数を O(文字数) → O(1)/列 に削減。
        const cellStart = paragraphs[0].startIndex;
        const cellEnd   = paragraphs[paragraphs.length - 1].endIndex;

        const lineSpacing = chunkLen > cpLine
          ? Math.round(cpLine * lSpacing / chunkLen)
          : lSpacing;

        styleReqs.push({
          updateParagraphStyle: {
            range: { startIndex: cellStart, endIndex: cellEnd },
            paragraphStyle: {
              alignment: 'CENTER',
              lineSpacing,
              spaceAbove: { magnitude: 0, unit: 'PT' },
              spaceBelow: { magnitude: 0, unit: 'PT' }
            },
            fields: 'alignment,lineSpacing,spaceAbove,spaceBelow'
          }
        });

        const textStyle = { fontSize: { magnitude: fSize, unit: 'PT' } };
        let tsFields = 'fontSize';
        if (pageChunks._fontFamily && pageChunks._fontFamily !== 'default') {
          textStyle.weightedFontFamily = { fontFamily: pageChunks._fontFamily };
          tsFields += ',weightedFontFamily';
        }
        styleReqs.push({
          updateTextStyle: {
            range: { startIndex: cellStart, endIndex: cellEnd },
            textStyle,
            fields: tsFields
          }
        });

      } else if (row === 1) {
        // メタデータ行：1pt白文字・行間最小化（セル全体を一括更新）
        const paragraphs = tableCell.content || [];
        if (paragraphs.length === 0) continue;
        const cellStart = paragraphs[0].startIndex;
        const cellEnd   = paragraphs[paragraphs.length - 1].endIndex;

        styleReqs.push({
          updateParagraphStyle: {
            range: { startIndex: cellStart, endIndex: cellEnd },
            paragraphStyle: {
              lineSpacing:  1,
              spaceAbove:   { magnitude: 0, unit: 'PT' },
              spaceBelow:   { magnitude: 0, unit: 'PT' }
            },
            fields: 'lineSpacing,spaceAbove,spaceBelow'
          }
        });
        styleReqs.push({
          updateTextStyle: {
            range: { startIndex: cellStart, endIndex: cellEnd },
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

  // スタイル適用 ＋ 列幅再適用を1回の batchUpdate に統合
  // （テキストスタイル適用後に列幅が変化することがあるため末尾で列幅を確定）
  await batchUpdate(token, docId, [...styleReqs, ...makeColWidthReqs(tableStartIndex)]);
  // ────────────────────────────────────────────────────────────────────────────────────

  return tableStartIndex;
}

// ─── テーブル直前の auto_para に pageBreakBefore=true を設定 ─────
// Google Docs は insertTable 時にテーブル直前に auto_para を自動生成する。
// この auto_para に pbb=true を設定することで、余分な段落を挿入せずに改ページを実現する。
//
// 旧アプローチ（使用禁止）:
//   pbb_para を先に挿入 → insertTable → auto_para が追加される
//   結果: TABLE_N → pbb_para(pbb=true) → auto_para → TABLE_{N+1}
//   問題: pbb_para と auto_para の2段落が空白ページを生成する
//
// 新アプローチ（このメソッド）:
//   insertTable → auto_para を特定 → auto_para に pbb=true を設定
//   結果: TABLE_N → auto_para(pbb=true) → TABLE_{N+1}
//   1段落のみなので余分な空白ページが生成されない
async function setPBBBeforeTable(token, docId, tableStartIndex) {
  const doc     = await docsGet(token, docId);
  const content = doc.body.content;

  const tableBodyIdx = content.findIndex(el => el.table && el.startIndex === tableStartIndex);
  if (tableBodyIdx <= 0) {
    console.error('[FVdoc v4] setPBBBeforeTable: テーブルが見つかりません (startIndex=', tableStartIndex, ')');
    return;
  }

  const prevEl = content[tableBodyIdx - 1];
  if (!prevEl || !prevEl.paragraph) {
    console.error('[FVdoc v4] setPBBBeforeTable: テーブル直前の段落が見つかりません', JSON.stringify({
      tableBodyIdx, prevType: prevEl ? (prevEl.table ? 'TABLE' : 'OTHER') : 'null'
    }));
    return;
  }

  const white = { color: { rgbColor: { red: 1, green: 1, blue: 1 } } };
  await batchUpdate(token, docId, [
    {
      updateParagraphStyle: {
        range: { startIndex: prevEl.startIndex, endIndex: prevEl.endIndex },
        paragraphStyle: {
          pageBreakBefore: true,
          lineSpacing:  100,
          spaceAbove:   { magnitude: 0, unit: 'PT' },
          spaceBelow:   { magnitude: 0, unit: 'PT' }
        },
        fields: 'pageBreakBefore,lineSpacing,spaceAbove,spaceBelow'
      }
    },
    {
      updateTextStyle: {
        range: { startIndex: prevEl.startIndex, endIndex: prevEl.endIndex },
        textStyle: {
          fontSize:        { magnitude: 1, unit: 'PT' },
          foregroundColor: white
        },
        fields: 'fontSize,foregroundColor'
      }
    }
  ]);
}

// ─── メインロジック：テーブルの挿入（改ページ対応） ──────────────
async function handleInsert(docId, params) {
  const { text, charsPerLine, fontSize, fontFamily, lineSpacingPct, colGapPt } = params;
  let   cpLine   = parseInt(charsPerLine);
  const fSize    = parseFloat(fontSize);
  const lSpacing = parseFloat(lineSpacingPct) || 70;
  const colGap   = parseFloat(colGapPt)       || 0;
  const token    = await getToken();

  // ── ページ寸法とスペーサー計算 ────────────────────────────────
  const doc            = await docsGet(token, docId);
  const docStyle       = doc.documentStyle || {};
  const pageHeightPt   = toPt(docStyle.pageSize?.height)  || 842;
  const marginTopPt    = toPt(docStyle.marginTop)          || 72;
  const marginBottomPt = toPt(docStyle.marginBottom)       || 72;
  const usableHeightPt = pageHeightPt - marginTopPt - marginBottomPt;

  // ── cpLine の上限をページ高さから自動計算してキャップ ──────────
  //
  // 【行高さの実測に基づく計算式】
  // Google Docs の実際の1段落レンダリング高さは、フォントのメトリクスにより
  // 「フォントサイズ × 約2倍」になることが実測で確認されている。
  // （Noto Serif JP等のCJKフォントはフォントのascender/descenderが大きく、
  //   lineSpacing=100(single spacing)でも行高≒fSize×2.0になる）
  //
  // lineSpacing < 100 はGoogle Docsでは最小値（100相当）として扱われるため、
  // 実効的な行高 = fSize × 2.0 × max(lSpacing, 100) / 100
  //
  // 各セルは「本文 chunkLen 段落 + 末尾空段落 1 段落」の構成:
  //   セル実高 = (chunkLen + 1) × lineHeightPt ≤ usableHeightPt
  //   → chunkLen の上限 = floor(usableHeightPt / lineHeightPt) - 1
  //
  // フォントサイズを大きくすると1列の最大文字数が減る（物理的制約）。
  // ユーザーの指定値がこの上限を超えている場合は自動的にキャップする。
  {
    const effectiveLSpacing  = Math.max(lSpacing, 100);
    const lineHeightPtEst    = fSize * 2.0 * (effectiveLSpacing / 100);
    const maxLinesPerPageEst = Math.max(1, Math.floor(usableHeightPt / lineHeightPtEst) - 1);
    if (cpLine > maxLinesPerPageEst) {
      console.log(`[FVdoc] cpLine 自動キャップ: ${cpLine}→${maxLinesPerPageEst}` +
        ` (fSize=${fSize}, lSpacing=${lSpacing}→実効${effectiveLSpacing},` +
        ` lineH=${lineHeightPtEst.toFixed(2)}pt, usableH=${usableHeightPt.toFixed(1)}pt)`);
      cpLine = maxLinesPerPageEst;
    }
  }

  const chunks = buildChunks(text, cpLine);
  if (!chunks.length) throw new Error('テキストがありません');

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

  // ── セクションスタイル診断: documentStyle と実際のセクション余白・段組みを比較 ────
  try {
    const sections = doc.body?.content?.filter(el => el.sectionBreak) || [];
    if (sections.length > 0) {
      const secStyle    = sections[sections.length - 1].sectionBreak?.sectionStyle || {};
      const secLeft     = toPt(secStyle.marginLeft);
      const secRight    = toPt(secStyle.marginRight);
      const secWidth    = toPt(secStyle.pageSize?.width);
      const colProps    = secStyle.columnProperties || [];
      // columnProperties が 2件以上 = 多段組みレイアウト（テーブル幅が制限される！）
      const colWidths   = colProps.map(c => ({ w: toPt(c.width), pad: toPt(c.paddingEnd) }));
      const colCount    = colProps.length;
      console.log('[FVdoc] sectionStyle: ' + JSON.stringify({
        secLeft, secRight, secWidth,
        secUsable: (secWidth || pageWidthPt) - (secLeft || marginLeftPt) - (secRight || marginRightPt),
        multiColumnCount: colCount,       // 2以上なら多段組み → テーブル幅が 1/N に制限
        columnWidths:     colWidths,      // 各段の幅（存在する場合）
        equalColumnsBool: secStyle.equalColumns
      }));
    } else {
      console.log('[FVdoc] sectionStyle: セクション区切りなし（documentStyle のみ）');
    }
  } catch (e) {
    console.error('[FVdoc] sectionStyle 診断エラー:', e.message);
  }
  // ─────────────────────────────────────────────────────────────────────────────

  // 列間(colGap)を反映: 各セル左右パディング = 列間/2 → 隣接列間の視覚的隙間 = 列間
  //   colWidthPt = fSize + colGap  （文字幅 + 列間 = 1列が水平に占める幅）
  //   cellPadH   = colGap / 2      （隙間を左右のセルに均等分割）
  //   columnsPerPage は colWidthPt から導出される（列間設定が先、列数は結果）
  const cellPadH   = colGap / 2;
  const colWidthPt = fSize + cellPadH * 2; // = fSize + colGap

  // 1ページに収まる行数（縦書きセル内の「行」＝1文字＝1段落の高さ）
  //
  // 各セルは「本文 chunkLen 段落 + 末尾空段落 1 段落」の構成になっており、
  // セルの実際の高さ = (chunkLen + 1) × lineHeightPt。
  // そのため chunkLen の上限は floor(usableHeightPt / lineHeightPt) - 1 とする必要がある。
  //
  // lineHeightPt: 実測により Google Docs の実際の行高 = fSize × 2.0 × max(lSpacing,100)/100
  // （上部の cpLine キャップと同じ式を使用して一貫性を保つ）
  const lineHeightPt    = fSize * 2.0 * (Math.max(lSpacing, 100) / 100);
  const maxLinesPerPage = Math.max(1, Math.floor(usableHeightPt / lineHeightPt) - 1);

  // 水平方向の列数ページ分割
  //
  //   Google Docs はテーブル列幅合計がページ幅を超えると全列を比例スケーリングする。
  //   RULES.md Section 1-3 の設計: columnsPerPage = floor(usableWidthPt / colWidthPt)
  //
  //   colGap=0 の場合：スケーリング許容（~1.33倍まで） → 全列を1ページに収める
  //   colGap>0 の場合：スケーリング不可（列間が潰れる） → 1ページあたり上記式で分割
  const SCALE_TOLERANCE   = (colGap > 0) ? 1.0 : 1.33;
  const totalContentWidth = chunks.length * colWidthPt;

  const colMagnitude = Math.round(colWidthPt);  // insertPageTable でも同値を使用

  let columnsPerPage;
  if (totalContentWidth <= usableWidthPt * SCALE_TOLERANCE) {
    // スケーリング許容範囲内 → 全列を1ページに
    columnsPerPage = chunks.length;
  } else {
    // RULES.md Section 1-3: floor(usableWidthPt / colWidthPt)
    // → ページ1の合計がわずかに usableWidthPt を超えることで Docs 比例縮小が発動し
    //   テーブルがページ幅いっぱいにレンダリングされる（原則1・2の設計的意図）
    columnsPerPage = Math.max(1, Math.floor(Math.round(usableWidthPt) / colMagnitude));
  }

  console.log('[FVdoc] columnsPerPage:', {
    colWidthPt, colMagnitude, usableWidthPt: Math.round(usableWidthPt),
    formula: `floor(${Math.round(usableWidthPt)}/${colMagnitude}) = ${columnsPerPage}`,
    totalPerPage: columnsPerPage * colMagnitude + Math.max(5, Math.round(usableWidthPt) - columnsPerPage * colMagnitude)
  });

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

  // ────────────────────────────────────────────────────────────────
  // 改ページ規則（RULES.md に準拠）:
  //
  //   【水平ページグループ境界 (pageIdx > 0)】
  //     → needsPageBreak=true を設定し、テーブル挿入後に setPBBBeforeTable で
  //       auto_para(テーブル直前の段落)に pageBreakBefore=true を適用する
  //     → 各ページグループは独立した物理ページに配置
  //     → ページ1を水平方向に埋め切ってから次ページへ（fill-first）
  //
  //   【高さスライス境界 (sliceIdx > 0)】
  //     → 次スライスが残り有効高さを超える場合のみ needsPageBreak=true を設定
  //     → 収まる場合は同じページに続けて配置（高さ残量を消費）
  //
  //   【改ページ実現の仕組み】
  //     insertTable 時に Google Docs が TABLE 直前に auto_para を自動生成する。
  //     この auto_para に pbb=true を設定することで、余分な段落挿入なしに改ページを実現。
  //     （旧: insertPageBreakParagraph で pbb_para を先行挿入 → auto_para と2段落になり空白ページが発生）
  // ────────────────────────────────────────────────────────────────
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

    // needsPageBreak: 次のテーブル挿入後に setPBBBeforeTable を呼び出すフラグ
    // （テーブル挿入前に pbb_para を挿入すると空白ページが発生するため、挿入後に auto_para へ設定する）
    let needsPageBreak = false;

    // 水平ページグループが変わるとき: 常に改ページ
    if (pageIdx > 0) {
      needsPageBreak = true;
      remainingHeightPt = usableHeightPt;
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

      const sliceLines    = lineEnd - lineStart;
      const sliceHeightPt = sliceLines * lineHeightPt;

      // 高さスライスが変わるとき: 残り高さに収まらない場合のみ改ページ
      if (sliceIdx > 0) {
        const willBreak = sliceHeightPt > remainingHeightPt;
        console.log('[FVdoc v4] heightSlice:', {
          pageIdx, sliceIdx, sliceLines,
          sliceHeightPt: Math.round(sliceHeightPt),
          remainingHeightPt: Math.round(remainingHeightPt),
          willBreak
        });
        if (willBreak) {
          needsPageBreak = true;
          remainingHeightPt = usableHeightPt;
        }
      }

      slicedChunks._fontFamily = pageChunks._fontFamily;

      const tableStartIndex = await insertPageTable(token, docId, slicedChunks, {
        cpLine, fSize, lSpacing, colGap, cellPadH, colWidthPt, usableWidthPt,
        columnsPerPage,
        isFirstPage: isFirstTable,
        metaJson: isFirstTable ? metaJson : null
      });

      // テーブル挿入後に auto_para へ pbb=true を設定（余分な段落挿入を回避）
      if (needsPageBreak) {
        await setPBBBeforeTable(token, docId, tableStartIndex);
        needsPageBreak = false;
      }

      if (isFirstTable) firstTableStartIndex = tableStartIndex;
      isFirstTable = false;
      remainingHeightPt = Math.max(0, remainingHeightPt - sliceHeightPt);
    }
  }

  console.log('[FVdoc v4] result:', { columnsPerPage, numPages: pageChunkGroups.length, numCols: chunks.length, usableWidthPt, colWidthPt, usableHeightPt, lineHeightPt });
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
