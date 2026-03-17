# FVdoc — Googleドキュメント 縦書き Chrome拡張機能

どのGoogleドキュメントでも、ページ右端の **「縦書き」ボタン** 1クリックで縦書きが使えます。

---

## 使い方

1. Googleドキュメントを開く
2. 画面右端に表示される青い **「縦書き」** ボタンをクリック
3. サイドバーにテキストを入力し、設定を調整して **「ドキュメントに挿入」** を押す
4. 初回のみ Google アカウントの認証ダイアログが表示されます → 「許可」をクリック

---

## セットアップ手順（初回のみ・約10分）

### ① 拡張機能をChromeに読み込む

1. Chrome で `chrome://extensions` を開く
2. 右上の **「デベロッパー モード」** をオンにする
3. **「パッケージ化されていない拡張機能を読み込む」** をクリック
4. この `fvdoc-extension` フォルダを選択
5. 拡張機能IDが表示されます（例：`abcdefghijklmnopabcdefghijklmnop`）→ コピーしておく

### ② Google Cloud で OAuth2 認証情報を作成

1. [Google Cloud Console](https://console.cloud.google.com/) を開く
2. 新しいプロジェクトを作成（例：「FVdoc」）
3. 左メニュー **「APIとサービス」→「ライブラリ」** → `Google Docs API` を検索して **有効化**
4. 左メニュー **「APIとサービス」→「認証情報」** → **「+ 認証情報を作成」→「OAuthクライアントID」**
5. アプリケーションの種類：**「Chrome アプリ」** を選択
6. アプリケーションID欄に、①でコピーした **拡張機能ID** を貼り付け
7. 「作成」→ 表示される **クライアントID** をコピー（`xxxxxxxx.apps.googleusercontent.com` の形式）

### ③ manifest.json を更新

`manifest.json` の以下の行を編集します：

```json
"client_id": "REPLACE_WITH_YOUR_CLIENT_ID.apps.googleusercontent.com",
```

`REPLACE_WITH_YOUR_CLIENT_ID` の部分を、②でコピーしたクライアントIDに置き換えてください。

### ④ 拡張機能を再読み込み

`chrome://extensions` で FVdoc の **「再読み込み」** ボタン（🔄）をクリック

---

## ファイル構成

```
fvdoc-extension/
├── manifest.json      — 拡張機能の設定（MV3）
├── background.js      — OAuth認証 + Google Docs API処理
├── content.js         — サイドバーUI注入 + 縦書きロジック
├── icons/
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
└── README.md
```

---

## 機能

- 右から左への縦書きレイアウト（日本語の流れに対応）
- 句読点・括弧・記号の縦書き用Unicode自動変換（`。→︒` `、→︑` `（）→︵︶` `[]→﹇﹈` `:→‥` など）
- フォント・サイズ設定対応（Noto Serif JP など）
- 均等割付：任意の列を選択して字間を調整
- 均等割付の解除
- 既存FVdoc表のテキスト復元
- 元テキスト・設定をメタデータとして表内に保存（白文字で不可視）
