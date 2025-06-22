# GAS開発上の注意点

Google Apps Script (GAS) で Web アプリケーションを開発する際の注意点をまとめます。特に iframe 関連の問題とその解決策について詳しく説明します。

## 1. iframe 埋め込みによる問題

### 1.1 問題の概要

GAS Web アプリは `googleusercontent.com` ドメインの iframe 内で実行されます。これにより、通常の Web アプリケーションとは異なる制約が発生します。

### 1.2 主な問題点

#### URL構造の問題
- **問題**: アプリケーションのURLが複雑で予測困難
  - 本体URL: `https://script.google.com/macros/s/{SCRIPT_ID}/exec`
  - iframe URL: `https://n-{RANDOM_STRING}.googleusercontent.com/...`
- **影響**: 
  - ハードコードされたURLが機能しない
  - 相対URLが正しく解決されない
- **解決策**:
  ```javascript
  // ScriptApp.getService().getUrl()を使用して動的にURLを取得
  function getWebAppUrl() {
    return ScriptApp.getService().getUrl();
  }
  ```

#### ナビゲーション制限
- **問題**: `<a href>` タグや `location.href` での直接的なページ遷移が正しく動作しない
- **影響**: メニューボタンやリンクをクリックしても白紙ページが表示される
- **解決策**:
  ```javascript
  // window.top.location.hrefを使用してiframeから脱出
  function navigateToPage(page) {
    google.script.run
      .withSuccessHandler(function(url) {
        window.top.location.href = url + '?page=' + page;
      })
      .getDeployedUrl();
  }
  ```

#### 親フレームアクセス制限
- **問題**: セキュリティポリシーにより親フレームへの直接アクセスが制限される
- **影響**: `parent.location` や `top.location` への直接アクセスがブロックされる場合がある
- **解決策**: `google.script.run` を経由してサーバーサイドでURLを取得

#### セッション・Cookie問題
- **問題**: サードパーティCookieとして扱われ、ブラウザの設定によってブロックされる
- **影響**: セッション管理が正しく機能しない可能性
- **解決策**: PropertiesServiceを使用したサーバーサイドでの状態管理

#### CORS（クロスオリジンリソース共有）制限
- **問題**: 外部APIへの直接アクセスが制限される
- **影響**: クライアントサイドからの外部API呼び出しが失敗
- **解決策**: サーバーサイド（GAS）経由でのAPI呼び出し

#### ブラウザ履歴の問題
- **問題**: iframe内のナビゲーションがブラウザ履歴に正しく記録されない
- **影響**: ブラウザの戻る/進むボタンが期待通りに動作しない
- **解決策**: URLパラメータを使用した擬似的なルーティング実装

#### JavaScriptコンテキストの分離
- **問題**: iframe内外でJavaScriptの実行コンテキストが分離される
- **影響**: グローバル変数やイベントの共有ができない
- **解決策**: `google.script.run` を使用した通信

#### フルスクリーン制限
- **問題**: iframe内からのフルスクリーンAPIの使用が制限される
- **影響**: 動画やプレゼンテーションのフルスクリーン表示ができない
- **解決策**: 新しいウィンドウでの表示を提案

#### ファイルダウンロード制限
- **問題**: 直接的なファイルダウンロードが制限される
- **影響**: 生成したファイルのダウンロードが複雑
- **解決策**: 
  ```javascript
  // データURIを使用したダウンロード
  function downloadFile(content, filename) {
    const blob = new Blob([content], {type: 'text/plain'});
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    a.click();
  }
  ```

## 2. テンプレート処理の注意点

### 2.1 正しいテンプレート処理方法

```javascript
// 誤った方法
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index');
  // <?!= ?> タグが処理されずそのまま表示される
}

// 正しい方法
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  return template.evaluate();
  // <?!= ?> タグが正しく処理される
}
```

### 2.2 テンプレートへの変数渡し

```javascript
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('index');
  template.baseUrl = getWebAppUrl(); // テンプレート変数を設定
  return template.evaluate();
}
```

## 3. アクセス権限の設定

### 3.1 appsscript.json の設定

```json
{
  "timeZone": "Asia/Tokyo",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "webapp": {
    "access": "ANYONE",  // 誰でもアクセス可能
    "executeAs": "USER_DEPLOYING"  // デプロイしたユーザーとして実行
  }
}
```

### 3.2 アクセス権限の種類
- `MYSELF`: 自分のみアクセス可能
- `DOMAIN`: 同じドメインのユーザーのみ
- `ANYONE_ANONYMOUS`: 誰でも（匿名）
- `ANYONE`: 誰でも（要Googleアカウント）

## 4. 推奨される実装パターン

### 4.1 ナビゲーション関数の統一

```javascript
// common-scripts.html
function navigateTo(page) {
  google.script.run
    .withSuccessHandler(function(url) {
      window.top.location.href = url + '?page=' + page;
    })
    .getDeployedUrl();
}

function backToMenu() {
  navigateTo('menu');
}
```

### 4.2 エラーハンドリングの統一

```javascript
function handleError(error) {
  console.error('Error:', error);
  showMessage('エラーが発生しました: ' + error.toString(), 'error');
}

// 全てのgoogle.script.run呼び出しに適用
google.script.run
  .withSuccessHandler(handleSuccess)
  .withFailureHandler(handleError)
  .serverFunction();
```

### 4.3 メッセージ表示の統一

```javascript
function showMessage(message, type) {
  const messageDiv = document.getElementById('message');
  messageDiv.className = 'message ' + type;
  messageDiv.textContent = message;
  messageDiv.style.display = 'block';
  
  // 3秒後に自動的に非表示
  setTimeout(() => {
    messageDiv.style.display = 'none';
  }, 3000);
}
```

## 5. パフォーマンス最適化

### 5.1 サーバー呼び出しの最小化

```javascript
// 悪い例: 複数回のサーバー呼び出し
for (let id of bookIds) {
  google.script.run.getBookInfo(id);
}

// 良い例: 一括でデータを取得
google.script.run.getMultipleBookInfo(bookIds);
```

### 5.2 ローディング表示の実装

```javascript
function showLoading() {
  document.getElementById('loading').style.display = 'block';
}

function hideLoading() {
  document.getElementById('loading').style.display = 'none';
}

// 使用例
showLoading();
google.script.run
  .withSuccessHandler(function(result) {
    hideLoading();
    // 処理
  })
  .withFailureHandler(function(error) {
    hideLoading();
    handleError(error);
  })
  .serverFunction();
```

## 6. デバッグのヒント

### 6.1 console.log の活用
- クライアントサイド: ブラウザの開発者ツールで確認
- サーバーサイド: Stackdriver Logging または Logger.log() を使用

### 6.2 エラーの詳細表示

```javascript
.withFailureHandler(function(error) {
  console.error('Detailed error:', error);
  console.error('Stack trace:', error.stack);
  showMessage('エラー: ' + error.message, 'error');
})
```

## 7. まとめ

GAS Web アプリケーション開発では、iframe による制約を理解し、適切な回避策を実装することが重要です。特に：

1. **URLは動的に取得**: ハードコードを避ける
2. **ナビゲーションは window.top を使用**: iframe から脱出
3. **サーバー通信は google.script.run**: 非同期処理を適切に扱う
4. **エラーハンドリングを統一**: ユーザーフレンドリーなエラー表示
5. **パフォーマンスを意識**: サーバー呼び出しを最小化

これらの注意点を守ることで、安定した GAS Web アプリケーションの開発が可能になります。