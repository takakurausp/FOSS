# ダッシュボード読み込み停止の原因調査と修正プラン



## Context



各ロール（Author, Editor, Reviewer, Managing-Editor, EIC）のダッシュボードが「読み込み中...」のまま停止する問題。ブラウザコンソールに以下のエラーが表示されている:



```

Uncaught SyntaxError: Unexpected identifier 'style' (at userCodeAppPanel:1186:35)

```



## 原因



### 根本原因: `buildEditorReportLinksPanel` 関数の `flatMap` がパースエラーを引き起こしている



**エラー位置の特定:**

- `scripts.html:1185-1190` の `buildEditorReportLinksPanel` 関数内

- `reported.flatMap(e => [...])` の箇所



**影響メカニズム:**

1. `scripts.html` 全体が1つの `<script>` ブロック（2194行）

2. GAS の `HtmlService.createHtmlOutputFromFile('scripts').getContent()` がスクリプト内容を処理する際、`flatMap` と HTML を含むテンプレートリテラルの組み合わせで構文エラーが発生

3. **SyntaxError はパース時エラー** → スクリプトブロック全体の実行が阻止される

4. `DOMContentLoaded` ハンドラーが登録されず、`initUI()` も呼ばれない

5. `index.html:17` の初期ローダー「読み込み中...」がそのまま残る



**全ロールに影響する理由:** パースエラーは関数単位ではなく `<script>` ブロック単位で発生するため、`buildEditorReportLinksPanel` が特定のロールでしか呼ばれなくても、同じスクリプトブロック内の全関数（`initUI`, `renderAuthorView`, `renderEditorView` 等）が定義されない。



## 修正内容



### Fix 1: `flatMap` を `reduce` に置き換え (scripts.html:1185-1190) — **最優先**



`flatMap` は ES2019 のメソッドで、GAS の HTML サービスのサンドボックス処理との互換性に問題がある可能性がある。`reduce` に書き換えて回避する。



**対象ファイル:** `scripts.html`

**対象行:** 1185-1190



**変更前:**

```javascript

const items = reported.flatMap(e => [

    e.reportPdfUrl        ? `<div style=...>...</div>` : '',

    e.reportWordUrl       ? `<div style=...>...</div>` : '',

    e.reportGoogleDocId   ? `<div style=...>...</div>` : '',

    e.reportCommentPdfUrl ? `<div style=...>...</div>` : ''

  ]).filter(Boolean);

```



**変更後:**

```javascript

const items = [];

reported.forEach(e => {

    if (e.reportPdfUrl) items.push(`<div style=...>...</div>`);

    if (e.reportWordUrl) items.push(`<div style=...>...</div>`);

    if (e.reportGoogleDocId) items.push(`<div style=...>...</div>`);

    if (e.reportCommentPdfUrl) items.push(`<div style=...>...</div>`);

});

```



### Fix 2: `convertDatesToStrings` のタイムゾーン修正 (ManuscriptDataHandlers.js:18) — 中



`'JST'` は IANA 標準タイムゾーン ID ではない。`'Asia/Tokyo'` に修正。



### Fix 3: `doGet()` のテンプレートデータサニタイズ (Code.js) — 中（予防的）



`<?!= initialMsData ?>` でエスケープなし出力されるため、データに `</script>` が含まれると HTML が壊れる。`JSON.stringify` 後に `</` を `<\/` に置換。



### Fix 4: managing-editor パスのエラーハンドリング (ManuscriptDataHandlers.js:524-536) — 低



try-catch がない managing-editor パスにエラーハンドリングを追加。



## 対象ファイル



| ファイル | 修正箇所 | 優先度 |

|---------|---------|-------|

| `scripts.html` | L1185-1190: `flatMap` → `forEach` | 最優先 |

| `ManuscriptDataHandlers.js` | L18: `'JST'` → `'Asia/Tokyo'` | 中 |

| `Code.js` | L48,62,70,80,90,98: テンプレートデータサニタイズ | 中 |

| `ManuscriptDataHandlers.js` | L524-536: try-catch追加 | 低 |



## 検証方法



1. 修正後、GAS にデプロイし直す

2. 各ロール（Author, Editor, Reviewer, Managing-Editor, EIC）のダッシュボードURLにアクセス

3. ブラウザの DevTools Console で `SyntaxError` が消えていることを確認

4. ダッシュボードが正常に表示されることを確認

