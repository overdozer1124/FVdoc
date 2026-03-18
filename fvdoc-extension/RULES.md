# FVdoc 改ページ・レイアウト 厳守ルール

このファイルは `background.js` の改ページロジックを実装・修正する際に
**絶対に守らなければならない規則** を定めたものです。
過去の誤実装を防ぐため、変更前に必ずこのファイルを読むこと。

---

## 1. 基本レイアウト概念

### 1-1. 縦書きの方向
- 文字は**上から下**に並ぶ（1列 = 1セル内に `\n` 区切りで縦積み）
- 列は**右から左**へ並ぶ（日本語縦書きの伝統的な組み方）
- 最初の文字 = ページ右端の列の先頭

### 1-2. 列の寸法
```
colWidthPt  = fSize + (1 + colGap/2) × 2   ← 1列の横幅（セルパディング込み）
lineHeightPt = fSize × (lSpacing / 100)     ← 1文字分の高さ
colHeightPt  = cpLine × lineHeightPt        ← 1列の縦の高さ
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

**規則: 水平ページグループが変わるときは、残り高さに関わらず常に `NEXT_PAGE` セクション区切りを挿入する。**

```
グループ0（25列）→ ページ1  [セクション区切り]
グループ1（21列）→ ページ2  [セクション区切り]
グループ2（??列）→ ページ3
```

**理由**:
「1ページ目を全部埋めてから2ページ目へ」とは、**水平方向（列数）をページ幅いっぱい**に詰めることを意味する。
グループ0が25列（ページ幅いっぱい）配置されたなら、それでページ1は「満杯」である。
グループ1は次の物理ページ（ページ2）の先頭から始める。

**禁止事項**:
- ❌ グループ1をグループ0と**同じページ**に縦積みしてはいけない
- ❌ 残り有効高さがあるからといって、グループ1の挿入前のセクション区切りを省略してはいけない

---

### ★ 規則B：高さスライス境界（`sliceIdx > 0`、同一ページグループ内）

**定義**: 1ページグループ内の列が `maxLinesPerPage` を超える場合、縦方向に分割したスライス。

```
maxLinesPerPage = floor(usableHeightPt / lineHeightPt)
```

**規則: 次スライスの高さが現在ページの残り有効高さを超える場合のみ `NEXT_PAGE` を挿入する。**

```
if (sliceHeightPt > remainingHeightPt) → 改ページ
else                                    → 同じページに続けて配置
```

**`remainingHeightPt` の管理**:
- 初期値: `usableHeightPt`
- セクション区切り挿入後: `usableHeightPt` にリセット
- テーブル挿入後: `remainingHeightPt -= sliceHeightPt`

---

## 3. ページ1の「満杯」の正しい定義

**正しい定義**:
ページ1には `columnsPerPage`（例: 25）列が水平方向にぴったり収まっている状態。
スペーサー列を含めたテーブル総幅 ≈ `usableWidthPt`。

**「満杯」ではない状態**:
- 25列未満しか配置されていない（水平方向が空いている）
- 偶数分割（例: 23+23列）になっている ← fill-first ではなく even-distribution になっている

**「満杯」に関する注意**:
- ページ1の**縦方向の充填率**は、テキストの文字数（`cpLine × lineHeightPt`）によって決まる
- `cpLine=30`、`lineHeightPt=7.7pt` の場合、1テーブルの縦高さ = `30 × 7.7 = 231pt`
- これは A4 有効高さ 698pt の 33% にすぎないが、これは**コンテンツ量の問題**であり、コードのバグではない
- コードの役割は「水平方向にページ幅いっぱい列を詰める」こと。縦方向の充填率を操作する機能はない

---

## 4. テーブル構造

| 列 | 役割 | widthType | 幅 |
|---|---|---|---|
| col 0 | スペーサー（右寄せ用） | `FIXED_WIDTH` | `max(5, usableWidthPt - numCols × colWidthPt)` |
| col 1〜N | コンテンツ（縦書き文字） | `FIXED_WIDTH` | `colWidthPt` |

| 行 | 役割 |
|---|---|
| row 0 | コンテンツ（各セルに `\n` 区切りの文字列） |
| row 1 | メタデータ（1pt 白文字、不可視） |

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

- [ ] `pageIdx > 0` のとき、常にセクション区切りを挿入しているか？
- [ ] `sliceIdx > 0` のときだけ高さベースの判定をしているか？
- [ ] `pageIdx > 0` の条件で高さベース判定（`sliceHeightPt > remainingHeightPt`）を適用していないか？
- [ ] スペーサーは `FIXED_WIDTH` を使用しているか（`EVENLY_DISTRIBUTED` ではないか）？
- [ ] 新しいページグループの処理に入る前に `remainingHeightPt = usableHeightPt` でリセットしているか？

---

## 7. よくある誤実装パターン（禁止）

### ❌ パターン1: 高さに空きがあればグループを同じページに置く
```javascript
// 誤: pageIdx > 0 でも高さが余っていれば同じページへ
if (sliceHeightPt > remainingHeightPt) { insertSectionBreak(); }
// → TABLE1(ページ1) と TABLE2 が同じページに縦積みされ、
//   ページ1が中途半端な充填率になる
```

### ❌ パターン2: 常にセクション区切りを挿入する（高さスライスも含む）
```javascript
// 誤: pageIdx/sliceIdx 問わず常にセクション区切り
if (pageIdx > 0 || sliceIdx > 0) { insertSectionBreak(); }
// → 高さスライスが複数あるとき、同じページに収まるスライスも
//   別ページに送られてしまう
```

### ✅ 正しいパターン
```javascript
// 水平グループ境界: 常にセクション区切り
if (pageIdx > 0) {
  insertSectionBreak();
  remainingHeightPt = usableHeightPt;
}
// 高さスライス境界: 収まらない場合のみセクション区切り
if (sliceIdx > 0 && sliceHeightPt > remainingHeightPt) {
  insertSectionBreak();
  remainingHeightPt = usableHeightPt;
}
remainingHeightPt -= sliceHeightPt;
```
