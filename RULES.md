# FVdoc レイアウト・改ページ 厳守ルール

このファイルは `background.js` の実装・修正を行う際に
**絶対に守らなければならない規則** を定めたものです。
コードを変更する前に必ずこのファイルを読むこと。

---

## 🔴 最重要原則（TOP 3）

これらは設計の根幹であり、どんな変更でも例外なく守ること。

### 原則1: テーブルは必ずページ右端に揃える

- テーブルの**右端**は、ページ余白を除いた有効領域の**右端（= 余白直内側）**に揃えること
- スペーサー列（col0）を左側に配置し、残りの幅をコンテンツ列が占める構造で右寄せを実現する
- 縦書きは右から左へ読むため、最初の文字は常にページの右端から始まる

```
ページ有効幅
|←────────────────────────────────────────────────→|
| スペーサー(col0) | col_N | col_N-1 | ... | col_1 |
|                 |←── コンテンツ（右端が読み始め）──→|
```

### 原則2: テーブルはページの右余白を絶対に越えない

- テーブルの総幅は `usableWidthPt`（ページ幅 - 左余白 - 右余白）以内に収めること
- 列数 × colWidthPt が usableWidthPt を超えた場合は水平ページ分割を行うこと
- colGap > 0 の場合はスケーリング許容なし（SCALE_TOLERANCE = 1.0）

### 原則3: 改ページはそのページが文字で埋まった場合のみ行う

- ページに収まる最大列数（`columnsPerPage`）分の文字が水平方向に配置されたとき、初めてそのページは「満杯」になる
- 満杯になったときのみ次ページへ改ページする（`pageIdx > 0` のとき）
- 列が途中でも収まらない場合（高さオーバー）は例外的に改ページを行う（`sliceIdx > 0 && willBreak` のとき）
- **「満杯でないのに改ページ」は絶対に禁止**（空白ページや無駄な改ページが発生する）

---

## 1. 基本レイアウト概念

### 1-1. 縦書きの方向
- 文字は**上から下**に並ぶ（1列 = 1セル内に `\n` 区切りで縦積み）
- 列は**右から左**へ並ぶ（日本語縦書きの伝統的な組み方）
- 最初の文字 = ページ右端の列の先頭

### 1-2. 列の寸法
```
colWidthPt   = fSize + (1 + colGap/2) × 2   ← 1列の横幅（セルパディング込み）
lineHeightPt = fSize × (lSpacing / 100)      ← 1文字分の高さ
```

### 1-3. 1ページに収まる列数
```
columnsPerPage = floor(usableWidthPt / colWidthPt)   ← colGap > 0 のとき
```
colGap=0 のときは自動スケーリング許容（SCALE_TOLERANCE=1.33）により全列が1ページに収まる可能性あり。

---

## 2. 改ページの2種類と規則

### ★ 規則A：水平ページグループ境界（`pageIdx > 0`）

**定義**: テキスト列を `columnsPerPage` ずつに分割したグループ。
グループ0 = ページ1、グループ1 = ページ2、グループ2 = ページ3 …

**規則: 水平ページグループが変わるときは、残り高さに関わらず常に改ページする。**

```
グループ0（例: 25列）→ ページ1  ← 水平方向が満杯になった
グループ1（例: 21列）→ ページ2  ← 次の水平グループ
```

**理由**:
「1ページ目を全部埋めてから2ページ目へ」とは、**水平方向（列数）をページ幅いっぱい**に詰めることを意味する。
グループ0が `columnsPerPage` 列（ページ幅いっぱい）配置されたなら、それでページ1は「満杯」である。
グループ1は次の物理ページの先頭から始める（原則3）。

**禁止事項**:
- ❌ グループ1をグループ0と**同じページ**に縦積みしてはいけない
- ❌ 残り有効高さがあるからといって、グループ1の改ページを省略してはいけない

---

### ★ 規則B：高さスライス境界（`sliceIdx > 0`、同一ページグループ内）

**定義**: 1ページグループ内の列が `maxLinesPerPage` を超える場合、縦方向に分割したスライス。

```
maxLinesPerPage = floor(usableHeightPt / lineHeightPt)
```

**規則: 次スライスの高さが現在ページの残り有効高さを超える場合のみ改ページする。**

```
if (sliceHeightPt > remainingHeightPt) → 改ページ
else                                    → 同じページに続けて配置
```

**`remainingHeightPt` の管理**:
- 初期値: `usableHeightPt`
- 改ページ後: `usableHeightPt` にリセット
- テーブル挿入後: `remainingHeightPt -= sliceHeightPt`

---

## 3. 改ページの実装方法（`setPBBBeforeTable`）

### ⚠️ 使用禁止の方法

| 方法 | 理由 |
|------|------|
| `insertSectionBreak(NEXT_PAGE)` | 空白ページが発生するため禁止 |
| `insertPageBreakParagraph`（旧関数） | pbb_para + auto_para の2段落になり空白ページが発生するため廃止済み |

### ✅ 正しい方法: `setPBBBeforeTable`

Google Docs は `insertTable` 時にテーブル直前に `auto_para` を自動生成する。
この `auto_para` に `pageBreakBefore=true` を設定することで、余分な段落を増やさずに改ページを実現する。

```
【構造】TABLE_N → auto_para(pbb=true, 1pt白字, 不可視) → TABLE_{N+1}
```

- 段落は1つのみ → 空白ページが発生しない
- `auto_para` はテーブル挿入後に `body.content` からインデックスで特定する

```javascript
// 正しい実装パターン
let needsPageBreak = false;
if (pageIdx > 0) {
  needsPageBreak = true;
  remainingHeightPt = usableHeightPt;
}
// ...
if (sliceIdx > 0 && sliceHeightPt > remainingHeightPt) {
  needsPageBreak = true;
  remainingHeightPt = usableHeightPt;
}
const tableStartIndex = await insertPageTable(...);
if (needsPageBreak) {
  await setPBBBeforeTable(token, docId, tableStartIndex);
  needsPageBreak = false;
}
```

---

## 4. テーブル構造

| 列 | 役割 | widthType | 幅 |
|---|---|---|---|
| col 0 | スペーサー（右寄せ用、左側に配置） | `FIXED_WIDTH` | `max(5, usableWidthPt - numCols × colWidthPt)` |
| col 1〜N | コンテンツ（縦書き文字） | `FIXED_WIDTH` | `colWidthPt` |

| 行 | 役割 |
|---|---|
| row 0 | コンテンツ（各セルに `\n` 区切りの文字列） |
| row 1 | メタデータ（1pt 白文字、不可視） |

**スペーサーの役割**:
- スペーサーを左側に置くことでコンテンツ列が右側に押し出され、**テーブル右端がページ右端に揃う**（原則1）
- ページ2以降は列数が少ないためスペーサーが広くなり、それでも右端は揃ったまま

---

## 5. スペーサー幅の計算

```
spacerMagnitude = max(5, round(usableWidthPt) - numPageCols × round(colWidthPt))
```

- ページ1（25列、colWidthPt=18pt）: `max(5, 451-450) = 5pt`  → テキストはページ右端から左端まで
- ページ2（21列、colWidthPt=18pt）: `max(5, 451-378) = 73pt` → テキストは右端に寄せられる

**重要**: スペーサーは必ず `FIXED_WIDTH` で指定すること。
`EVENLY_DISTRIBUTED` を使うと Google Docs が最小幅（〜70pt）を強制するため使用禁止。

---

## 6. コード修正時のチェックリスト

改ページロジックを変更する際は以下を必ず確認すること:

- [ ] テーブル右端はページ右端の余白内側に揃っているか？（原則1）
- [ ] テーブル総幅が `usableWidthPt` を超えていないか？（原則2）
- [ ] 改ページは「満杯になった場合のみ」か？（原則3）
- [ ] `pageIdx > 0` のとき、常に `needsPageBreak = true` にしているか？
- [ ] `sliceIdx > 0` のときだけ高さベースの判定をしているか？
- [ ] `pageIdx > 0` の条件で高さベース判定（`sliceHeightPt > remainingHeightPt`）を適用していないか？
- [ ] スペーサーは `FIXED_WIDTH` を使用しているか（`EVENLY_DISTRIBUTED` ではないか）？
- [ ] 新しいページグループの処理に入る前に `remainingHeightPt = usableHeightPt` でリセットしているか？
- [ ] 改ページに `insertSectionBreak` を使っていないか？（`setPBBBeforeTable` を使うこと）
- [ ] 改ページに `insertPageBreakParagraph`（旧関数）を使っていないか？（廃止済み）

---

## 7. よくある誤実装パターン（禁止）

### ❌ パターン1: 高さに空きがあればグループを同じページに置く

```javascript
// 誤: pageIdx > 0 でも高さが余っていれば同じページへ
if (sliceHeightPt > remainingHeightPt) { pageBreak(); }
// → TABLE1(ページ1) と TABLE2 が同じページに縦積みされる（原則3違反）
```

### ❌ パターン2: 常に改ページする（高さスライスも含む）

```javascript
// 誤: pageIdx/sliceIdx 問わず常に改ページ
if (pageIdx > 0 || sliceIdx > 0) { pageBreak(); }
// → 同じページに収まるスライスも別ページに送られてしまう
```

### ❌ パターン3: insertSectionBreak / pbb_para の先行挿入

```javascript
// 誤1: セクション区切り → 空白ページが発生
await insertSectionBreak(token, docId, 'NEXT_PAGE');

// 誤2: pbb_para を先行挿入してから insertTable
// → TABLE_N → pbb_para(pbb=true) → auto_para → TABLE_{N+1}
// → 2段落あるため空白ページが発生
await insertPageBreakParagraph(token, docId);
await insertPageTable(...);
```

### ✅ 正しいパターン

```javascript
// 改ページフラグを立ててからテーブル挿入
let needsPageBreak = false;
if (pageIdx > 0) { needsPageBreak = true; remainingHeightPt = usableHeightPt; }
if (sliceIdx > 0 && sliceHeightPt > remainingHeightPt) {
  needsPageBreak = true;
  remainingHeightPt = usableHeightPt;
}

// テーブルを先に挿入
const tableStartIndex = await insertPageTable(...);

// テーブル挿入後に auto_para へ pbb=true を設定
if (needsPageBreak) {
  await setPBBBeforeTable(token, docId, tableStartIndex);
  needsPageBreak = false;
}
remainingHeightPt -= sliceHeightPt;
```
