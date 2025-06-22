# CLAUDE.md

このファイルは、Claude Code (claude.ai/code) がこのリポジトリで作業する際のガイダンスを提供します。

## 必須参照ドキュメント

開発時は常に以下のドキュメントを参照してください：

- @要件定義書.md - システムの機能要件と非機能要件の詳細
- @デザイン要件定義書.md - UIデザインとビジュアルデザインの仕様
- @GAS開発上の注意点.md - GAS特有のiframe問題と解決策

## モデル

### データ構造定義

#### 書籍マスタ (books_master)
```javascript
{
  isbn: string,           // 主キー、ISBN番号
  title: string,          // 書籍名
  titleKana: string,      // 書籍名ヨミ（ひらがなまたはカタカナ）
  author: string,         // 著者
  authorKana: string,     // 著者ヨミ（オプション）
  publisher: string,      // 出版社
  publishYear: number,    // 発行年（オプション）
  price: number,          // 価格（オプション）
  category: string,       // カテゴリ（オプション）
  registeredDate: Date,   // 登録日（自動設定）
  isDeleted: boolean      // 削除フラグ（FALSE/TRUE）
}
```

#### 書籍在庫 (books_inventory)
```javascript
{
  bookId: string,         // 主キー、自動採番（B0001形式）
  isbn: string,           // 外部キー（書籍マスタ）
  location: string,       // 所在場所（棚番号など）
  status: string,         // 状態: "利用可能" | "貸出中" | "修理中" | "廃棄"
  registeredDate: Date,   // 登録日（自動設定）
  updatedDate: Date,      // 更新日（自動更新）
  notes: string           // 備考
}
```

#### 利用者マスタ (users_master)
```javascript
{
  userId: string,         // 主キー、自動採番（U0001形式）
  userName: string,       // 利用者名
  userNameKana: string,   // フリガナ
  email: string,          // メールアドレス（オプション）
  phone: string,          // 電話番号
  address: string,        // 住所（オプション）
  birthDate: Date,        // 生年月日（オプション）
  registeredDate: Date,   // 登録日（自動設定）
  updatedDate: Date,      // 更新日（自動更新）
  status: string,         // 状態: "有効" | "無効" | "退会"
  cardNumber: string      // カード番号（バーコード）
}
```

#### 貸出 (loans)
```javascript
{
  loanId: string,         // 主キー、自動採番（L000001形式）
  bookId: string,         // 外部キー（書籍在庫）
  userId: string,         // 外部キー（利用者マスタ）
  loanDate: Date,         // 貸出日時
  dueDate: Date,          // 返却予定日
  status: string,         // 返却状況: "貸出中" | "返却済み"
  overdueDays: number     // 延滞日数（0以上、自動計算）
}
```

#### バックアップ管理 (backup_management)
```javascript
{
  backupId: string,       // 主キー、自動採番（BK000001形式）
  backupDate: Date,       // バックアップ日時
  backupUrl: string,      // 生成されたスプレッドシートのURL
  backupType: string,     // バックアップ種別: "手動" | "定期"
  executor: string,       // 実行者
  fileSize: string,       // ファイルサイズ（概算）
  notes: string           // 備考
}
```

### ID採番ルール
- 書籍ID: B + 4桁数字（B0001〜B9999）
- 利用者ID: U + 4桁数字（U0001〜U9999）
- 貸出ID: L + 6桁数字（L000001〜L999999）
- バックアップID: BK + 6桁数字（BK000001〜BK999999）

## アプリケーション

### 書籍管理機能

#### 書籍登録
- メニュー画面から「書籍登録」ボタンをクリック
- 書籍情報入力フォームが表示される
  - ISBNコードをバーコードスキャン or 手動入力
  - ISBNがあれば国立国会図書館APIから書籍情報を自動取得
  - 必須項目: 書籍名、著者、出版社、ISBN
  - オプション項目: 発行年、価格、カテゴリ、所在場所
- 「登録」ボタンで書籍マスタに登録
- 同一ISBNで複数冊登録する場合は、書籍在庫に複数レコード作成

#### 書籍検索・編集
- 書籍ID、ISBN、書籍名、書籍名ヨミで検索
- よみがな検索は部分一致に対応
- 検索結果から編集対象を選択
- 書籍情報を修正して「更新」ボタン
- 編集履歴はスプレッドシートの変更履歴で管理

#### 在庫管理
- 在庫一覧画面で全書籍の状態を確認
- フィルター機能: カテゴリ、著者、状態（利用可能/貸出中）
- 書籍名ヨミ、著者ヨミでの部分一致検索
- 検索結果は最大50件まで表示（パフォーマンス対策）

### 利用者管理機能

#### 利用者登録
- メニュー画面から「利用者登録」ボタンをクリック
- 利用者情報入力フォームが表示される
  - 必須項目: 氏名、フリガナ、電話番号
  - オプション項目: メールアドレス、住所、生年月日
- 「登録」ボタンで利用者マスタに登録
- 登録完了後、メールアドレスがあれば確認メール送信
- 図書館カード番号（バーコード）を自動生成

#### カード発行
- 利用者IDで検索
- 「カード発行」ボタンでバーコード付きカードを印刷
- カード再発行も同じ手順で実行

### 貸出・返却機能

#### 書籍貸出
- メニュー画面から「書籍貸出」ボタンをクリック
- 書籍IDをバーコードスキャン or 手動入力
- 利用者IDをバーコードスキャン or 手動入力
- 貸出情報確認画面で内容を確認
- 「貸出実行」ボタンで処理完了
  - 貸出テーブルに新規レコード作成
  - 書籍在庫の状態を「貸出中」に更新
  - 返却予定日は14日後に自動設定
- 複数冊の一括貸出も可能

#### 書籍返却
- メニュー画面から「書籍返却」ボタンをクリック
- 書籍IDをバーコードスキャン or 手動入力
- 返却情報確認画面で延滞日数を表示
- 「返却実行」ボタンで処理完了
  - 貸出テーブルの状態を「返却済み」に更新
  - 書籍在庫の状態を「利用可能」に更新
  - 返却日時を記録
  - 月次で貸出履歴テーブルにアーカイブ

#### 利用者別返却
- 利用者ID、利用者名、フリガナで検索
- 貸出中書籍を一覧表示
- 返却する書籍を選択（複数選択可）
- 一括返却処理を実行

### レポート・統計機能

#### 延滞リスト
- 返却予定日を過ぎた書籍を一覧表示
- 延滞日数で降順ソート
- 「通知メール送信」ボタンで延滞通知
- 週次（毎週月曜日）で自動実行

#### 貸出統計（簡素化版）
- 月次貸出数のグラフ表示
- 人気書籍トップ10
- エクスポート機能（CSV形式）

### システム管理機能

#### 図書館設定
- 貸出期間の設定（デフォルト14日）
- 延滞通知メールのテンプレート編集
- 図書館情報（名称、住所、連絡先）の設定

#### データバックアップ
- 「バックアップ実行」ボタンで手動バックアップ
- 新規スプレッドシートを生成（ファイル名: 図書館システムバックアップ_YYYYMMDD）
- 全シートのデータをコピー
- バックアップ管理シートにURL、日時、実行者を記録
- 月次自動バックアップ（毎月1日）

## システム

### 開発環境（絶対条件）
- **プラットフォーム**: Google Apps Script (GAS) のみ使用
- **データベース**: Google スプレッドシート のみ使用
- **エディタ**: GAS エディタ（ブラウザベース）
- **外部データベース禁止**: MySQL、PostgreSQL等は使用不可

### 技術スタック
```
フロントエンド:
- HTML5 (HtmlService経由で提供)
- CSS3 (インラインスタイル推奨)
- JavaScript ES6
- QuaggaJS (バーコードスキャン)
- Google Fonts (Noto Sans JP)

バックエンド:
- Google Apps Script V8ランタイム
- SpreadsheetApp API
- GmailApp API (メール送信)
- HtmlService (Web UI提供)
- PropertiesService (ID採番管理)
```

### ファイル構成
```
/
├── コード.js           # メインのGASコード
├── appsscript.json    # GAS設定ファイル
├── 要件定義書.md       # システム要件定義
└── CLAUDE.md          # このファイル
```

### 開発方針

#### パフォーマンス最適化
- スプレッドシートAPI呼び出しを最小化
  ```javascript
  // 悪い例: 個別にセルを読み取る
  for (let i = 0; i < 100; i++) {
    const value = sheet.getRange(i, 1).getValue();
  }
  
  // 良い例: 範囲を一括で読み取る
  const values = sheet.getRange(1, 1, 100, 1).getValues();
  ```
- 検索は最大50件に制限
- インデックステーブルは使用せず、フィルター関数で対応
- キャッシュは使用しない（10名程度の同時アクセスでは不要）

#### ふりがな検索の実装
```javascript
function searchByKana(data, searchText, kanaColumnIndex) {
  // ひらがな・カタカナの正規化
  const normalizeKana = (text) => {
    return text.replace(/[\u30a1-\u30f6]/g, (match) => {
      return String.fromCharCode(match.charCodeAt(0) - 0x60);
    });
  };
  
  const normalizedSearch = normalizeKana(searchText.toLowerCase());
  
  return data.filter(row => {
    const kana = normalizeKana(row[kanaColumnIndex].toLowerCase());
    return kana.includes(normalizedSearch);
  });
}
```

#### エラーハンドリング
```javascript
try {
  // 処理
} catch (error) {
  console.error('エラーが発生しました:', error);
  return {
    success: false,
    message: 'エラーが発生しました。管理者に連絡してください。'
  };
}
```

#### ID採番の実装
```javascript
function getNextId(prefix, propertyName, digits) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastId = scriptProperties.getProperty(propertyName) || '0';
  const nextNumber = parseInt(lastId) + 1;
  const nextId = prefix + nextNumber.toString().padStart(digits, '0');
  scriptProperties.setProperty(propertyName, nextNumber.toString());
  return nextId;
}
```

### レスポンシブデザイン実装
```css
/* モバイル (〜767px) */
@media (max-width: 767px) {
  body { font-size: 16px; }
  button { height: 44px; font-size: 16px; }
  .container { padding: 10px; }
}

/* タブレット (768px〜1023px) */
@media (min-width: 768px) and (max-width: 1023px) {
  body { font-size: 18px; }
  button { height: 44px; font-size: 18px; }
  .container { max-width: 750px; }
}

/* デスクトップ (1024px〜) */
@media (min-width: 1024px) {
  body { font-size: 14px; }
  button { height: 32px; font-size: 14px; }
  .container { max-width: 1200px; }
}
```

### GAS固有の制約
- 実行時間: 最大6分/実行
- URLフェッチ: 最大20秒/リクエスト
- メール送信: 100通/日（無料版）
- スプレッドシート: 500万セルまで
- 同時実行: 30実行まで

### デプロイ手順
1. GASエディタで「デプロイ」→「新しいデプロイ」
2. 種類: 「ウェブアプリ」を選択
3. 説明: バージョン情報を記入
4. 実行ユーザー: 「自分」
5. アクセス権: 「全員」（匿名アクセス可）
6. 「デプロイ」ボタンをクリック

### 運用スケジュール
- **日次**: 貸出・返却処理（約50件/日）
- **週次**: 延滞チェック・通知（毎週月曜日）
- **月次**: 統計レポート生成、データ整合性チェック
- **年次**: 古いデータのアーカイブ（2年以上前の返却済みデータ）