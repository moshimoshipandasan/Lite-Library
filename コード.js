// ========================================
// 設定・定数
// ========================================
const CONFIG = {
  SPREADSHEET_ID: PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID'),
  LOAN_PERIOD_DAYS: 14,
  MAX_SEARCH_RESULTS: 50,
  ISBN_API_URL: 'https://ndlsearch.ndl.go.jp/api/sru'
};

const SHEET_NAMES = {
  BOOKS_MASTER: '書籍マスタ',
  BOOKS_INVENTORY: '書籍在庫',
  USERS_MASTER: '利用者マスタ',
  LOANS: '貸出',
  BACKUP_MANAGEMENT: 'バックアップ管理'
};

const COLUMN_INDEXES = {
  BOOKS_MASTER: {
    ISBN: 0, TITLE: 1, TITLE_KANA: 2, AUTHOR: 3, AUTHOR_KANA: 4,
    PUBLISHER: 5, PUBLISH_YEAR: 6, PRICE: 7, CATEGORY: 8,
    REGISTERED_DATE: 9, IS_DELETED: 10
  },
  BOOKS_INVENTORY: {
    BOOK_ID: 0, ISBN: 1, LOCATION: 2, STATUS: 3,
    REGISTERED_DATE: 4, UPDATED_DATE: 5, NOTES: 6
  },
  USERS_MASTER: {
    USER_ID: 0, USER_NAME: 1, USER_NAME_KANA: 2, EMAIL: 3, PHONE: 4,
    ADDRESS: 5, BIRTH_DATE: 6, REGISTERED_DATE: 7, UPDATED_DATE: 8,
    STATUS: 9, CARD_NUMBER: 10
  },
  LOANS: {
    LOAN_ID: 0, BOOK_ID: 1, USER_ID: 2, LOAN_DATE: 3,
    DUE_DATE: 4, RETURN_DATE: 5, STATUS: 6, OVERDUE_DAYS: 7
  }
};

const STATUS = {
  BOOK: {
    AVAILABLE: '利用可能',
    LOANED: '貸出中',
    MAINTENANCE: '修理中',
    DISPOSED: '廃棄'
  },
  USER: {
    ACTIVE: '有効',
    INACTIVE: '無効',
    WITHDRAWN: '退会'
  },
  LOAN: {
    LOANED: '貸出中',
    RETURNED: '返却済み'
  }
};

// ========================================
// 初期化・メイン処理
// ========================================
function doGet(e) {
  // 認証チェックを一時的に無効化（誰でもアクセス可能にするため）
  // const user = Session.getActiveUser().getEmail();
  // if (!user) {
  //   const template = HtmlService.createTemplateFromFile('login');
  //   template.baseUrl = getWebAppUrl();
  //   return template.evaluate()
  //     .setTitle('図書館管理システム - ログイン')
  //     .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  // }
  
  const page = e.parameter.page || 'menu';
  
  // HTMLファイル名のマッピング
  const pageMap = {
    'menu': 'menu',
    'bookRegister': 'book-register',
    'bookSearch': 'book-search',
    'userRegister': 'user-register',
    'userSearch': 'user-search',
    'loan': 'loan',
    'return': 'return',
    'reports': 'reports',
    'settings': 'settings',
    'barcodeScanner': 'barcode-scanner',
    'barcodePopup': 'barcode-popup',
    'cardPrint': 'card-print'
  };
  
  const fileName = pageMap[page] || 'menu';
  
  try {
    const template = HtmlService.createTemplateFromFile(fileName);
    template.baseUrl = getWebAppUrl();
    return template.evaluate()
      .setTitle('図書館管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (error) {
    // ファイルが存在しない場合はメニューに戻る
    const template = HtmlService.createTemplateFromFile('menu');
    template.baseUrl = getWebAppUrl();
    return template.evaluate()
      .setTitle('図書館管理システム')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
}

// Web App URLを取得する関数
function getWebAppUrl() {
  // スクリプトプロパティからURLを取得（設定されている場合）
  const scriptProperties = PropertiesService.getScriptProperties();
  const customUrl = scriptProperties.getProperty('WEB_APP_URL');
  
  if (customUrl) {
    return customUrl;
  }
  
  // デフォルトのURL取得を試みる
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    // エラーの場合は空文字を返す（相対URLになる）
    return '';
  }
}

// ========================================
// データベースアクセス層
// ========================================
class SpreadsheetDB {
  constructor() {
    if (!CONFIG.SPREADSHEET_ID) {
      throw new Error('SPREADSHEET_IDが設定されていません。スクリプトプロパティを確認してください。');
    }
    this.spreadsheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  }
  
  getSheet(sheetName) {
    return this.spreadsheet.getSheetByName(sheetName);
  }
  
  getData(sheetName, startRow = 2) {
    const sheet = this.getSheet(sheetName);
    if (!sheet || sheet.getLastRow() < startRow) return [];
    return sheet.getRange(startRow, 1, sheet.getLastRow() - startRow + 1, sheet.getLastColumn()).getValues();
  }
  
  findRow(sheetName, columnIndex, value) {
    const data = this.getData(sheetName);
    const rowIndex = data.findIndex(row => row[columnIndex] === value);
    return rowIndex === -1 ? -1 : rowIndex + 2;
  }
  
  updateRow(sheetName, rowNumber, data) {
    const sheet = this.getSheet(sheetName);
    sheet.getRange(rowNumber, 1, 1, data.length).setValues([data]);
  }
  
  appendRow(sheetName, data) {
    const sheet = this.getSheet(sheetName);
    sheet.appendRow(data);
  }
  
  searchWithCondition(sheetName, conditions = [], includeAll = false) {
    const data = this.getData(sheetName);
    if (includeAll) return data;
    
    return data.filter(row => {
      return conditions.every(condition => {
        const value = row[condition.columnIndex];
        switch (condition.operator) {
          case 'equals':
            return value === condition.value;
          case 'contains':
            return String(value).includes(condition.value);
          case 'notEquals':
            return value !== condition.value;
          default:
            return true;
        }
      });
    });
  }
}

// ========================================
// ID採番
// ========================================
function getNextId(prefix, propertyName, digits) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lastId = scriptProperties.getProperty(propertyName) || '0';
  const nextNumber = parseInt(lastId) + 1;
  const nextId = prefix + nextNumber.toString().padStart(digits, '0');
  scriptProperties.setProperty(propertyName, nextNumber.toString());
  return nextId;
}

// ========================================
// 共通ユーティリティ
// ========================================
function handleError(functionName, error) {
  console.error(`${functionName}でエラー:`, error);
  return {
    success: false,
    message: `処理中にエラーが発生しました: ${error.message || 'Unknown error'}`
  };
}

function normalizeKana(text) {
  if (!text) return '';
  // カタカナをひらがなに変換
  return text.replace(/[\u30a1-\u30f6]/g, match => {
    return String.fromCharCode(match.charCodeAt(0) - 0x60);
  });
}

function getSpreadsheet() {
  if (!CONFIG.SPREADSHEET_ID) {
    throw new Error('SPREADSHEET_IDが設定されていません');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

// ========================================
// 書籍管理機能
// ========================================

// ISBN検索関数
function searchBookByISBN(isbn) {
  try {
    // ISBNの正規化（ハイフンを除去）
    const normalizedISBN = isbn.replace(/-/g, '');
    
    // 国立国会図書館APIを使用
    const url = `https://ndlsearch.ndl.go.jp/api/sru?operation=searchRetrieve&query=isbn=${normalizedISBN}&recordPacking=xml&maximumRecords=1`;
    
    const response = UrlFetchApp.fetch(url);
    const xml = response.getContentText();
    
    // XMLをパース
    const document = XmlService.parse(xml);
    const root = document.getRootElement();
    
    // 名前空間の定義
    const srw = XmlService.getNamespace('http://www.loc.gov/zing/srw/');
    const dc = XmlService.getNamespace('http://purl.org/dc/elements/1.1/');
    
    // レコードを取得
    const records = root.getChild('records', srw);
    if (!records) {
      return { success: false, message: '書籍が見つかりませんでした' };
    }
    
    const record = records.getChild('record', srw);
    if (!record) {
      return { success: false, message: '書籍が見つかりませんでした' };
    }
    
    const recordData = record.getChild('recordData', srw);
    const dcRecord = recordData.getChildren()[0];
    
    // 書籍情報を抽出
    const bookInfo = {
      isbn: normalizedISBN,
      title: '',
      author: '',
      publisher: '',
      publishYear: ''
    };
    
    // タイトルを取得
    const titleElements = dcRecord.getChildren('title', dc);
    if (titleElements.length > 0) {
      bookInfo.title = titleElements[0].getText();
    }
    
    // 著者を取得
    const creatorElements = dcRecord.getChildren('creator', dc);
    if (creatorElements.length > 0) {
      bookInfo.author = creatorElements.map(el => el.getText()).join(', ');
    }
    
    // 出版社を取得
    const publisherElements = dcRecord.getChildren('publisher', dc);
    if (publisherElements.length > 0) {
      bookInfo.publisher = publisherElements[0].getText();
    }
    
    // 出版年を取得
    const dateElements = dcRecord.getChildren('date', dc);
    if (dateElements.length > 0) {
      const dateText = dateElements[0].getText();
      const yearMatch = dateText.match(/\d{4}/);
      if (yearMatch) {
        bookInfo.publishYear = yearMatch[0];
      }
    }
    
    return {
      success: true,
      bookInfo: bookInfo
    };
    
  } catch (error) {
    console.error('ISBN検索エラー:', error);
    
    // 代替APIとしてOpenBDを試す
    try {
      const normalizedISBN = isbn.replace(/-/g, '');
      const openBdUrl = `https://api.openbd.jp/v1/get?isbn=${normalizedISBN}`;
      const response = UrlFetchApp.fetch(openBdUrl);
      const data = JSON.parse(response.getContentText());
      
      if (data && data[0] && data[0].summary) {
        const summary = data[0].summary;
        return {
          success: true,
          bookInfo: {
            isbn: normalizedISBN,
            title: summary.title || '',
            author: summary.author || '',
            publisher: summary.publisher || '',
            publishYear: summary.pubdate ? summary.pubdate.substring(0, 4) : ''
          }
        };
      }
    } catch (openBdError) {
      console.error('OpenBD APIエラー:', openBdError);
    }
    
    return {
      success: false,
      message: 'ISBN検索に失敗しました'
    };
  }
}

function registerBook(bookData) {
  try {
    const db = new SpreadsheetDB();
    
    // 書籍マスタの重複チェック
    const existingBook = db.findRow(
      SHEET_NAMES.BOOKS_MASTER,
      COLUMN_INDEXES.BOOKS_MASTER.ISBN,
      bookData.isbn
    );
    
    // 書籍マスタへの登録（新規の場合のみ）
    if (existingBook === -1) {
      const masterRow = [
        bookData.isbn,
        bookData.title,
        bookData.titleKana || '',
        bookData.author,
        bookData.authorKana || '',
        bookData.publisher,
        bookData.publishYear || '',
        bookData.price || '',
        bookData.category || '',
        new Date(),
        false
      ];
      db.appendRow(SHEET_NAMES.BOOKS_MASTER, masterRow);
    }
    
    // 書籍在庫への登録
    const bookId = getNextId('B', 'LAST_BOOK_ID', 4);
    const inventoryRow = [
      bookId,
      bookData.isbn,
      bookData.location || '',
      STATUS.BOOK.AVAILABLE,
      new Date(),
      new Date(),
      ''
    ];
    db.appendRow(SHEET_NAMES.BOOKS_INVENTORY, inventoryRow);
    
    return {
      success: true,
      bookId: bookId,
      message: `書籍を登録しました。書籍ID: ${bookId}`
    };
  } catch (error) {
    return handleError('書籍登録', error);
  }
}

function searchBooks(searchCriteria) {
  try {
    const db = new SpreadsheetDB();
    const conditions = [];
    
    if (searchCriteria.bookId) {
      conditions.push({
        columnIndex: COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID,
        operator: 'equals',
        value: searchCriteria.bookId
      });
    }
    
    if (searchCriteria.isbn) {
      conditions.push({
        columnIndex: COLUMN_INDEXES.BOOKS_INVENTORY.ISBN,
        operator: 'equals',
        value: searchCriteria.isbn
      });
    }
    
    const inventoryData = db.searchWithCondition(SHEET_NAMES.BOOKS_INVENTORY, conditions);
    const masterData = db.getData(SHEET_NAMES.BOOKS_MASTER);
    
    const results = inventoryData.map(invRow => {
      const masterRow = masterData.find(m => m[COLUMN_INDEXES.BOOKS_MASTER.ISBN] === invRow[COLUMN_INDEXES.BOOKS_INVENTORY.ISBN]);
      if (!masterRow) return null;
      
      // タイトルまたはタイトルヨミで検索
      if (searchCriteria.title) {
        const searchText = searchCriteria.title.toLowerCase();
        const title = masterRow[COLUMN_INDEXES.BOOKS_MASTER.TITLE].toLowerCase();
        if (!title.includes(searchText)) return null;
      }
      
      if (searchCriteria.titleKana) {
        const searchKana = normalizeKana(searchCriteria.titleKana.toLowerCase());
        const titleKana = normalizeKana(masterRow[COLUMN_INDEXES.BOOKS_MASTER.TITLE_KANA].toLowerCase());
        if (!titleKana.includes(searchKana)) return null;
      }
      
      return {
        bookId: invRow[COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID],
        isbn: invRow[COLUMN_INDEXES.BOOKS_INVENTORY.ISBN],
        title: masterRow[COLUMN_INDEXES.BOOKS_MASTER.TITLE],
        titleKana: masterRow[COLUMN_INDEXES.BOOKS_MASTER.TITLE_KANA],
        author: masterRow[COLUMN_INDEXES.BOOKS_MASTER.AUTHOR],
        authorKana: masterRow[COLUMN_INDEXES.BOOKS_MASTER.AUTHOR_KANA],
        publisher: masterRow[COLUMN_INDEXES.BOOKS_MASTER.PUBLISHER],
        location: invRow[COLUMN_INDEXES.BOOKS_INVENTORY.LOCATION],
        status: invRow[COLUMN_INDEXES.BOOKS_INVENTORY.STATUS]
      };
    }).filter(book => book !== null);
    
    return results.slice(0, CONFIG.MAX_SEARCH_RESULTS);
  } catch (error) {
    console.error('書籍検索エラー:', error);
    return [];
  }
}

function getBookDetails(bookId) {
  try {
    const books = searchBooks({ bookId: bookId });
    return books.length > 0 ? books[0] : null;
  } catch (error) {
    console.error('書籍詳細取得エラー:', error);
    return null;
  }
}

function updateBook(updateData) {
  try {
    const db = new SpreadsheetDB();
    
    // 在庫データの更新
    const invRowNumber = db.findRow(
      SHEET_NAMES.BOOKS_INVENTORY,
      COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID,
      updateData.bookId
    );
    
    if (invRowNumber === -1) {
      return { success: false, message: '書籍が見つかりません' };
    }
    
    const invData = db.getData(SHEET_NAMES.BOOKS_INVENTORY)[invRowNumber - 2];
    invData[COLUMN_INDEXES.BOOKS_INVENTORY.LOCATION] = updateData.location;
    invData[COLUMN_INDEXES.BOOKS_INVENTORY.STATUS] = updateData.status;
    invData[COLUMN_INDEXES.BOOKS_INVENTORY.UPDATED_DATE] = new Date();
    
    db.updateRow(SHEET_NAMES.BOOKS_INVENTORY, invRowNumber, invData);
    
    // マスタデータの更新
    const isbn = invData[COLUMN_INDEXES.BOOKS_INVENTORY.ISBN];
    const masterRowNumber = db.findRow(
      SHEET_NAMES.BOOKS_MASTER,
      COLUMN_INDEXES.BOOKS_MASTER.ISBN,
      isbn
    );
    
    if (masterRowNumber !== -1) {
      const masterData = db.getData(SHEET_NAMES.BOOKS_MASTER)[masterRowNumber - 2];
      masterData[COLUMN_INDEXES.BOOKS_MASTER.TITLE] = updateData.title;
      masterData[COLUMN_INDEXES.BOOKS_MASTER.TITLE_KANA] = updateData.titleKana || '';
      masterData[COLUMN_INDEXES.BOOKS_MASTER.AUTHOR] = updateData.author;
      masterData[COLUMN_INDEXES.BOOKS_MASTER.AUTHOR_KANA] = updateData.authorKana || '';
      masterData[COLUMN_INDEXES.BOOKS_MASTER.PUBLISHER] = updateData.publisher;
      
      db.updateRow(SHEET_NAMES.BOOKS_MASTER, masterRowNumber, masterData);
    }
    
    return { success: true, message: '書籍情報を更新しました' };
  } catch (error) {
    return handleError('書籍更新', error);
  }
}

// ========================================
// 利用者管理機能
// ========================================

// カード印刷用のHTMLを生成
function generateCardPrintHtml(userId) {
  try {
    const db = new SpreadsheetDB();
    const userData = db.getData(SHEET_NAMES.USERS_MASTER);
    const userIndex = userData.findIndex(row => row[COLUMN_INDEXES.USERS_MASTER.USER_ID] === userId);
    
    if (userIndex === -1) {
      return { success: false, message: '利用者が見つかりません' };
    }
    
    const user = userData[userIndex];
    const userName = user[COLUMN_INDEXES.USERS_MASTER.USER_NAME];
    const cardNumber = user[COLUMN_INDEXES.USERS_MASTER.CARD_NUMBER];
    
    // カード印刷用のHTML（定期券サイズ: 85.6mm × 54mm）
    const html = `
      <!DOCTYPE html>
      <html>
      <head>
        <meta charset="UTF-8">
        <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
        <title>図書館利用者カード</title>
        <link href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;700&display=swap" rel="stylesheet">
        <style>
          @page {
            size: 85.6mm 54mm;
            margin: 0;
          }
          
          body {
            margin: 0;
            padding: 0;
            font-family: 'Noto Sans JP', 'メイリオ', sans-serif;
          }
          
          .card {
            width: 85.6mm;
            height: 54mm;
            position: relative;
            background: linear-gradient(135deg, #f0f8ff 0%, #e6f3ff 100%);
            border: 1px solid #4a90e2;
            box-sizing: border-box;
            padding: 4mm;
            display: flex;
            flex-direction: column;
            overflow: hidden;
          }
          
          .header {
            text-align: center;
            border-bottom: 1px solid #4a90e2;
            padding-bottom: 2mm;
            margin-bottom: 2mm;
          }
          
          .library-name {
            font-size: 13px;
            font-weight: bold;
            color: #2c5aa0;
            margin: 0;
          }
          
          .card-title {
            font-size: 9px;
            color: #666;
            margin: 0;
          }
          
          .user-info {
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            margin: 2mm 0;
          }
          
          .info-content {
            display: flex;
            align-items: baseline;
            gap: 10px;
          }
          
          .user-id {
            font-size: 12px;
            color: #666;
          }
          
          .user-name {
            font-size: 18px;
            font-weight: bold;
            color: #333;
          }
          
          .barcode-section {
            text-align: center;
            padding: 2mm 0 1mm 0;
            background-color: white;
            border-radius: 3px;
            margin: 0 -2mm;
          }
          
          .barcode {
            width: 100%;
            max-width: 65mm;
            height: 12mm;
          }
          
          .card-number {
            font-size: 8px;
            color: #333;
            margin-top: 0.5mm;
            letter-spacing: 0.5px;
          }
          
          @media print {
            body {
              margin: 0;
            }
            .card {
              page-break-after: always;
              border: none;
            }
          }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/jsbarcode@3.11.5/dist/JsBarcode.all.min.js"></script>
      </head>
      <body>
        <div class="card">
          <div class="header">
            <h1 class="library-name">図書館管理システム</h1>
            <p class="card-title">利用者カード</p>
          </div>
          
          <div class="user-info">
            <div class="info-content">
              <span class="user-id">${userId}</span>
              <span class="user-name">${userName}</span>
            </div>
          </div>
          
          <div class="barcode-section">
            <svg id="barcode" class="barcode"></svg>
            <div class="card-number">${cardNumber}</div>
          </div>
        </div>
        
        <script>
          // バーコードを生成
          JsBarcode("#barcode", "${cardNumber}", {
            format: "CODE128",
            width: 1.5,
            height: 35,
            displayValue: false,
            margin: 0
          });
          
          // 自動印刷
          window.onload = function() {
            setTimeout(function() {
              window.print();
            }, 500);
          };
        </script>
      </body>
      </html>
    `;
    
    return { success: true, html: html };
    
  } catch (error) {
    console.error('カード印刷エラー:', error);
    return { success: false, message: 'カード印刷の生成に失敗しました' };
  }
}

// カード印刷ウィンドウを開く
function openCardPrintWindow(userId) {
  const result = generateCardPrintHtml(userId);
  
  if (!result.success) {
    return result;
  }
  
  // HTMLを直接返す（Base64エンコードを使わない）
  return {
    success: true,
    html: result.html
  };
}
function registerUser(userData) {
  try {
    const db = new SpreadsheetDB();
    const userId = getNextId('U', 'LAST_USER_ID', 4);
    const cardNumber = generateCardNumber(userId);
    
    const userRow = [
      userId,
      userData.userName,
      userData.userNameKana,
      userData.email || '',
      userData.phone,
      userData.address || '',
      userData.birthDate || '',
      new Date(),
      new Date(),
      STATUS.USER.ACTIVE,
      cardNumber
    ];
    
    db.appendRow(SHEET_NAMES.USERS_MASTER, userRow);
    
    return {
      success: true,
      userId: userId,
      cardNumber: cardNumber,
      message: `利用者を登録しました。利用者ID: ${userId}`
    };
  } catch (error) {
    return handleError('利用者登録', error);
  }
}

function generateCardNumber(userId) {
  // 簡易的なカード番号生成（実際はより複雑なロジックを使用）
  const timestamp = Date.now().toString().slice(-6);
  return `LIB${userId}${timestamp}`;
}

function searchUsers(searchCriteria) {
  try {
    const db = new SpreadsheetDB();
    const userData = db.getData(SHEET_NAMES.USERS_MASTER);
    
    return userData.filter(row => {
      if (row[COLUMN_INDEXES.USERS_MASTER.STATUS] !== STATUS.USER.ACTIVE) return false;
      
      if (searchCriteria.userId && row[COLUMN_INDEXES.USERS_MASTER.USER_ID] !== searchCriteria.userId) {
        return false;
      }
      
      if (searchCriteria.userName) {
        const searchText = searchCriteria.userName.toLowerCase();
        const userName = row[COLUMN_INDEXES.USERS_MASTER.USER_NAME].toLowerCase();
        if (!userName.includes(searchText)) return false;
      }
      
      if (searchCriteria.userNameKana) {
        const searchKana = normalizeKana(searchCriteria.userNameKana.toLowerCase());
        const userNameKana = normalizeKana(row[COLUMN_INDEXES.USERS_MASTER.USER_NAME_KANA].toLowerCase());
        if (!userNameKana.includes(searchKana)) return false;
      }
      
      if (searchCriteria.phone && !row[COLUMN_INDEXES.USERS_MASTER.PHONE].includes(searchCriteria.phone)) {
        return false;
      }
      
      return true;
    }).map(row => ({
      userId: row[COLUMN_INDEXES.USERS_MASTER.USER_ID],
      userName: row[COLUMN_INDEXES.USERS_MASTER.USER_NAME],
      userNameKana: row[COLUMN_INDEXES.USERS_MASTER.USER_NAME_KANA],
      email: row[COLUMN_INDEXES.USERS_MASTER.EMAIL],
      phone: row[COLUMN_INDEXES.USERS_MASTER.PHONE],
      address: row[COLUMN_INDEXES.USERS_MASTER.ADDRESS],
      cardNumber: row[COLUMN_INDEXES.USERS_MASTER.CARD_NUMBER]
    })).slice(0, CONFIG.MAX_SEARCH_RESULTS);
  } catch (error) {
    console.error('利用者検索エラー:', error);
    return [];
  }
}

function getUserDetails(userId) {
  try {
    const users = searchUsers({ userId: userId });
    return users.length > 0 ? users[0] : null;
  } catch (error) {
    console.error('利用者詳細取得エラー:', error);
    return null;
  }
}

function updateUser(updateData) {
  try {
    const db = new SpreadsheetDB();
    const rowNumber = db.findRow(
      SHEET_NAMES.USERS_MASTER,
      COLUMN_INDEXES.USERS_MASTER.USER_ID,
      updateData.userId
    );
    
    if (rowNumber === -1) {
      return { success: false, message: '利用者が見つかりません' };
    }
    
    const userData = db.getData(SHEET_NAMES.USERS_MASTER)[rowNumber - 2];
    userData[COLUMN_INDEXES.USERS_MASTER.USER_NAME] = updateData.userName;
    userData[COLUMN_INDEXES.USERS_MASTER.USER_NAME_KANA] = updateData.userNameKana;
    userData[COLUMN_INDEXES.USERS_MASTER.EMAIL] = updateData.email || '';
    userData[COLUMN_INDEXES.USERS_MASTER.PHONE] = updateData.phone;
    userData[COLUMN_INDEXES.USERS_MASTER.ADDRESS] = updateData.address || '';
    userData[COLUMN_INDEXES.USERS_MASTER.UPDATED_DATE] = new Date();
    
    db.updateRow(SHEET_NAMES.USERS_MASTER, rowNumber, userData);
    
    return { success: true, message: '利用者情報を更新しました' };
  } catch (error) {
    return handleError('利用者更新', error);
  }
}

// ========================================
// 貸出・返却機能
// ========================================
function executeLoan(loanData) {
  try {
    const db = new SpreadsheetDB();
    const results = [];
    
    for (const bookId of loanData.bookIds) {
      // 書籍の利用可能チェック
      const invRowNumber = db.findRow(
        SHEET_NAMES.BOOKS_INVENTORY,
        COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID,
        bookId
      );
      
      if (invRowNumber === -1) {
        results.push({ bookId, success: false, message: '書籍が見つかりません' });
        continue;
      }
      
      const invData = db.getData(SHEET_NAMES.BOOKS_INVENTORY)[invRowNumber - 2];
      
      if (invData[COLUMN_INDEXES.BOOKS_INVENTORY.STATUS] !== STATUS.BOOK.AVAILABLE) {
        results.push({ bookId, success: false, message: '貸出できない状態です' });
        continue;
      }
      
      // 貸出レコード作成
      const loanId = getNextId('L', 'LAST_LOAN_ID', 6);
      const loanDate = new Date();
      const dueDate = new Date();
      dueDate.setDate(dueDate.getDate() + CONFIG.LOAN_PERIOD_DAYS);
      
      const loanRow = [
        loanId,
        bookId,
        loanData.userId,
        loanDate,
        dueDate,
        '',
        STATUS.LOAN.LOANED,
        0
      ];
      
      db.appendRow(SHEET_NAMES.LOANS, loanRow);
      
      // 在庫ステータス更新
      invData[COLUMN_INDEXES.BOOKS_INVENTORY.STATUS] = STATUS.BOOK.LOANED;
      invData[COLUMN_INDEXES.BOOKS_INVENTORY.UPDATED_DATE] = new Date();
      db.updateRow(SHEET_NAMES.BOOKS_INVENTORY, invRowNumber, invData);
      
      results.push({ bookId, success: true, loanId });
    }
    
    const successCount = results.filter(r => r.success).length;
    
    return {
      success: successCount > 0,
      results: results,
      message: `${successCount}冊の貸出処理が完了しました`
    };
  } catch (error) {
    return handleError('貸出処理', error);
  }
}

function getLoanInfoByBookId(bookId) {
  try {
    const db = new SpreadsheetDB();
    const loanData = db.getData(SHEET_NAMES.LOANS);
    
    const currentLoan = loanData.find(row => 
      row[COLUMN_INDEXES.LOANS.BOOK_ID] === bookId &&
      row[COLUMN_INDEXES.LOANS.STATUS] === STATUS.LOAN.LOANED
    );
    
    if (!currentLoan) return null;
    
    // 書籍情報取得
    const bookInfo = getBookDetails(bookId);
    if (!bookInfo) return null;
    
    // 利用者情報取得
    const userInfo = getUserDetails(currentLoan[COLUMN_INDEXES.LOANS.USER_ID]);
    if (!userInfo) return null;
    
    // 延滞日数計算
    const dueDate = new Date(currentLoan[COLUMN_INDEXES.LOANS.DUE_DATE]);
    const today = new Date();
    const overdueDays = Math.max(0, Math.floor((today - dueDate) / (1000 * 60 * 60 * 24)));
    
    return {
      loanId: currentLoan[COLUMN_INDEXES.LOANS.LOAN_ID],
      bookId: bookId,
      bookTitle: bookInfo.title,
      userId: userInfo.userId,
      userName: userInfo.userName,
      loanDate: currentLoan[COLUMN_INDEXES.LOANS.LOAN_DATE],
      dueDate: currentLoan[COLUMN_INDEXES.LOANS.DUE_DATE],
      overdueDays: overdueDays
    };
  } catch (error) {
    console.error('貸出情報取得エラー:', error);
    return null;
  }
}

function searchUserWithLoans(searchText) {
  try {
    const db = new SpreadsheetDB();
    
    // 利用者検索
    const users = searchUsers({
      userId: searchText,
      userName: searchText,
      userNameKana: searchText
    });
    
    if (users.length === 0) return { user: null, loans: [] };
    
    const user = users[0];
    
    // 貸出中の書籍を取得
    const loanData = db.getData(SHEET_NAMES.LOANS);
    const activeLoans = loanData.filter(row => 
      row[COLUMN_INDEXES.LOANS.USER_ID] === user.userId &&
      row[COLUMN_INDEXES.LOANS.STATUS] === STATUS.LOAN.LOANED
    );
    
    const loans = activeLoans.map(loan => {
      const bookInfo = getBookDetails(loan[COLUMN_INDEXES.LOANS.BOOK_ID]);
      const dueDate = new Date(loan[COLUMN_INDEXES.LOANS.DUE_DATE]);
      const today = new Date();
      const overdueDays = Math.max(0, Math.floor((today - dueDate) / (1000 * 60 * 60 * 24)));
      
      return {
        loanId: loan[COLUMN_INDEXES.LOANS.LOAN_ID],
        bookId: loan[COLUMN_INDEXES.LOANS.BOOK_ID],
        title: bookInfo ? bookInfo.title : '不明',
        loanDate: loan[COLUMN_INDEXES.LOANS.LOAN_DATE],
        dueDate: loan[COLUMN_INDEXES.LOANS.DUE_DATE],
        overdueDays: overdueDays
      };
    });
    
    return { user, loans };
  } catch (error) {
    console.error('利用者貸出検索エラー:', error);
    return { user: null, loans: [] };
  }
}

function executeReturn(returnData) {
  try {
    const db = new SpreadsheetDB();
    const results = [];
    const returnDate = new Date();
    
    for (const bookId of returnData.bookIds) {
      // 貸出レコードを検索
      const loanData = db.getData(SHEET_NAMES.LOANS);
      const loanIndex = loanData.findIndex(row => 
        row[COLUMN_INDEXES.LOANS.BOOK_ID] === bookId &&
        row[COLUMN_INDEXES.LOANS.STATUS] === STATUS.LOAN.LOANED
      );
      
      if (loanIndex === -1) {
        results.push({ bookId, success: false, message: '貸出記録が見つかりません' });
        continue;
      }
      
      // 貸出レコードを更新
      const loanRowNumber = loanIndex + 2;
      const loanRow = loanData[loanIndex];
      loanRow[COLUMN_INDEXES.LOANS.RETURN_DATE] = returnDate;
      loanRow[COLUMN_INDEXES.LOANS.STATUS] = STATUS.LOAN.RETURNED;
      
      // 延滞日数計算
      const dueDate = new Date(loanRow[COLUMN_INDEXES.LOANS.DUE_DATE]);
      const overdueDays = Math.max(0, Math.floor((returnDate - dueDate) / (1000 * 60 * 60 * 24)));
      loanRow[COLUMN_INDEXES.LOANS.OVERDUE_DAYS] = overdueDays;
      
      db.updateRow(SHEET_NAMES.LOANS, loanRowNumber, loanRow);
      
      // 在庫ステータスを「利用可能」に戻す
      const invRowNumber = db.findRow(
        SHEET_NAMES.BOOKS_INVENTORY,
        COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID,
        bookId
      );
      
      if (invRowNumber !== -1) {
        const invData = db.getData(SHEET_NAMES.BOOKS_INVENTORY)[invRowNumber - 2];
        invData[COLUMN_INDEXES.BOOKS_INVENTORY.STATUS] = STATUS.BOOK.AVAILABLE;
        invData[COLUMN_INDEXES.BOOKS_INVENTORY.UPDATED_DATE] = returnDate;
        db.updateRow(SHEET_NAMES.BOOKS_INVENTORY, invRowNumber, invData);
      }
      
      results.push({ bookId, success: true });
    }
    
    const successCount = results.filter(r => r.success).length;
    
    return {
      success: successCount > 0,
      results: results,
      returnedIds: results.filter(r => r.success).map(r => r.bookId),
      message: `${successCount}冊の返却処理が完了しました`
    };
  } catch (error) {
    return handleError('返却処理', error);
  }
}

// ========================================
// レポート・統計機能
// ========================================
function getOverdueList() {
  try {
    const db = new SpreadsheetDB();
    const loanData = db.getData(SHEET_NAMES.LOANS);
    const today = new Date();
    
    const overdueLoans = loanData.filter(row => {
      if (row[COLUMN_INDEXES.LOANS.STATUS] !== STATUS.LOAN.LOANED) return false;
      const dueDate = new Date(row[COLUMN_INDEXES.LOANS.DUE_DATE]);
      return dueDate < today;
    });
    
    return overdueLoans.map(loan => {
      const dueDate = new Date(loan[COLUMN_INDEXES.LOANS.DUE_DATE]);
      const overdueDays = Math.floor((today - dueDate) / (1000 * 60 * 60 * 24));
      
      const bookInfo = getBookDetails(loan[COLUMN_INDEXES.LOANS.BOOK_ID]);
      const userInfo = getUserDetails(loan[COLUMN_INDEXES.LOANS.USER_ID]);
      
      return {
        loanId: loan[COLUMN_INDEXES.LOANS.LOAN_ID],
        bookId: loan[COLUMN_INDEXES.LOANS.BOOK_ID],
        bookTitle: bookInfo ? bookInfo.title : '不明',
        userId: loan[COLUMN_INDEXES.LOANS.USER_ID],
        userName: userInfo ? userInfo.userName : '不明',
        email: userInfo ? userInfo.email : '',
        dueDate: dueDate,
        overdueDays: overdueDays
      };
    }).sort((a, b) => b.overdueDays - a.overdueDays);
  } catch (error) {
    console.error('延滞リスト取得エラー:', error);
    return [];
  }
}

function getMonthlyStatistics() {
  try {
    const db = new SpreadsheetDB();
    const loanData = db.getData(SHEET_NAMES.LOANS);
    const now = new Date();
    const currentMonth = now.getMonth();
    const currentYear = now.getFullYear();
    
    // 今月の貸出・返却数
    let monthlyLoans = 0;
    let monthlyReturns = 0;
    
    loanData.forEach(row => {
      const loanDate = new Date(row[COLUMN_INDEXES.LOANS.LOAN_DATE]);
      if (loanDate.getMonth() === currentMonth && loanDate.getFullYear() === currentYear) {
        monthlyLoans++;
      }
      
      if (row[COLUMN_INDEXES.LOANS.RETURN_DATE]) {
        const returnDate = new Date(row[COLUMN_INDEXES.LOANS.RETURN_DATE]);
        if (returnDate.getMonth() === currentMonth && returnDate.getFullYear() === currentYear) {
          monthlyReturns++;
        }
      }
    });
    
    // 人気書籍トップ10
    const bookLoanCounts = {};
    loanData.forEach(row => {
      const bookId = row[COLUMN_INDEXES.LOANS.BOOK_ID];
      bookLoanCounts[bookId] = (bookLoanCounts[bookId] || 0) + 1;
    });
    
    const popularBooks = Object.entries(bookLoanCounts)
      .map(([bookId, count]) => {
        const bookInfo = getBookBasicInfo(bookId);
        return bookInfo ? { bookId, title: bookInfo.title, loanCount: count } : null;
      })
      .filter(book => book !== null)
      .sort((a, b) => b.loanCount - a.loanCount)
      .slice(0, 10);
    
    return {
      currentMonth: {
        loans: monthlyLoans,
        returns: monthlyReturns
      },
      popularBooks: popularBooks
    };
  } catch (error) {
    console.error('統計データ取得エラー:', error);
    return {
      currentMonth: { loans: 0, returns: 0 },
      popularBooks: []
    };
  }
}

// 書籍の基本情報取得（統計用）
function getBookBasicInfo(bookId) {
  try {
    const db = new SpreadsheetDB();
    const inventoryData = db.getData(SHEET_NAMES.BOOKS_INVENTORY);
    const masterData = db.getData(SHEET_NAMES.BOOKS_MASTER);
    
    const invRow = inventoryData.find(row => 
      row[COLUMN_INDEXES.BOOKS_INVENTORY.BOOK_ID] === bookId
    );
    
    if (!invRow) return null;
    
    const masterRow = masterData.find(row => 
      row[COLUMN_INDEXES.BOOKS_MASTER.ISBN] === invRow[COLUMN_INDEXES.BOOKS_INVENTORY.ISBN]
    );
    
    if (!masterRow) return null;
    
    return {
      title: masterRow[COLUMN_INDEXES.BOOKS_MASTER.TITLE],
      author: masterRow[COLUMN_INDEXES.BOOKS_MASTER.AUTHOR]
    };
  } catch (error) {
    return null;
  }
}

// ========================================
// システム設定機能
// ========================================
function saveSystemSettings(settings) {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    // 各設定項目を保存
    const settingKeys = [
      'libraryName', 'libraryAddress', 'libraryPhone', 'libraryEmail',
      'loanPeriod', 'maxLoansPerUser', 'overdueNotificationTemplate'
    ];
    
    settingKeys.forEach(key => {
      if (settings[key] !== undefined) {
        scriptProperties.setProperty(key.toUpperCase(), String(settings[key]));
      }
    });
    
    return {
      success: true,
      message: '設定を保存しました'
    };
  } catch (error) {
    return handleError('システム設定保存', error);
  }
}

function getSystemSettings() {
  try {
    const scriptProperties = PropertiesService.getScriptProperties();
    
    return {
      libraryName: scriptProperties.getProperty('LIBRARYNAME') || '',
      libraryAddress: scriptProperties.getProperty('LIBRARYADDRESS') || '',
      libraryPhone: scriptProperties.getProperty('LIBRARYPHONE') || '',
      libraryEmail: scriptProperties.getProperty('LIBRARYEMAIL') || '',
      loanPeriod: parseInt(scriptProperties.getProperty('LOANPERIOD') || '14'),
      maxLoansPerUser: parseInt(scriptProperties.getProperty('MAXLOANSPERUSER') || '10'),
      overdueNotificationTemplate: scriptProperties.getProperty('OVERDUENOTIFICATIONTEMPLATE') || ''
    };
  } catch (error) {
    console.error('システム設定取得エラー:', error);
    return {
      libraryName: '',
      libraryAddress: '',
      libraryPhone: '',
      libraryEmail: '',
      loanPeriod: 14,
      maxLoansPerUser: 10,
      overdueNotificationTemplate: ''
    };
  }
}

// バックアップ実行
function executeBackup() {
  try {
    const sourceSpreadsheet = getSpreadsheet();
    const backupName = `図書館システムバックアップ_${Utilities.formatDate(new Date(), 'JST', 'yyyyMMdd_HHmmss')}`;
    
    // 新しいスプレッドシートを作成
    const backupSpreadsheet = SpreadsheetApp.create(backupName);
    const backupUrl = backupSpreadsheet.getUrl();
    
    // 各シートをコピー
    const sheets = sourceSpreadsheet.getSheets();
    sheets.forEach((sheet, index) => {
      const targetSheet = index === 0 
        ? backupSpreadsheet.getSheets()[0] 
        : backupSpreadsheet.insertSheet();
      
      targetSheet.setName(sheet.getName());
      
      const range = sheet.getDataRange();
      if (range.getNumRows() > 0 && range.getNumColumns() > 0) {
        const values = range.getValues();
        targetSheet.getRange(1, 1, values.length, values[0].length).setValues(values);
      }
    });
    
    // バックアップ管理に記録
    const db = new SpreadsheetDB();
    const backupId = getNextId('BK', 'LAST_BACKUP_ID', 6);
    const backupRow = [
      backupId,
      new Date(),
      backupUrl,
      '手動',
      Session.getActiveUser().getEmail(),
      'N/A',
      '手動バックアップ'
    ];
    
    db.appendRow(SHEET_NAMES.BACKUP_MANAGEMENT, backupRow);
    
    return {
      success: true,
      backup: {
        backupId: backupId,
        backupDate: new Date(),
        backupUrl: backupUrl,
        backupType: '手動',
        executor: Session.getActiveUser().getEmail()
      }
    };
  } catch (error) {
    return handleError('バックアップ実行', error);
  }
}

// バックアップ履歴を取得
function getBackupHistory() {
  try {
    const db = new SpreadsheetDB();
    const backups = db.getData(SHEET_NAMES.BACKUP_MANAGEMENT);
    
    return backups.map(row => ({
      backupId: row[0],
      backupDate: row[1],
      backupUrl: row[2],
      backupType: row[3],
      executor: row[4]
    }))
    .sort((a, b) => new Date(b.backupDate) - new Date(a.backupDate))
    .slice(0, 10);
  } catch (error) {
    console.error('バックアップ履歴の取得に失敗しました:', error);
    return [];
  }
}

// 月次統計をCSVエクスポート
function exportMonthlyStatistics() {
  try {
    const stats = getMonthlyStatistics();
    
    // CSVヘッダー
    let csv = '項目,値\n';
    csv += `今月の貸出数,${stats.currentMonth.loans}\n`;
    csv += `今月の返却数,${stats.currentMonth.returns}\n`;
    csv += '\n人気書籍トップ10\n';
    csv += '順位,書籍名,貸出回数\n';
    
    stats.popularBooks.forEach((book, index) => {
      csv += `${index + 1},${book.title},${book.loanCount}\n`;
    });
    
    return {
      success: true,
      csv: csv
    };
  } catch (error) {
    console.error('統計のエクスポートに失敗しました:', error);
    return {
      success: false,
      message: '統計のエクスポートに失敗しました'
    };
  }
}

// 延滞通知を送信
function sendOverdueNotifications() {
  try {
    const overdueList = getOverdueList();
    const settings = getSystemSettings();
    let sentCount = 0;
    
    overdueList.forEach(item => {
      if (item.email) {
        try {
          // メールテンプレートの変数を置換
          let emailBody = settings.overdueNotificationTemplate || getDefaultOverdueTemplate();
          emailBody = emailBody
            .replace(/{userName}/g, item.userName)
            .replace(/{bookTitle}/g, item.bookTitle)
            .replace(/{dueDate}/g, new Date(item.dueDate).toLocaleDateString('ja-JP'))
            .replace(/{overdueDays}/g, item.overdueDays)
            .replace(/{libraryName}/g, settings.libraryName || '図書館')
            .replace(/{libraryPhone}/g, settings.libraryPhone || '');
          
          // メール送信
          GmailApp.sendEmail(
            item.email,
            '図書返却のお願い',
            emailBody,
            {
              name: settings.libraryName || '図書館管理システム',
              noReply: true
            }
          );
          
          sentCount++;
        } catch (error) {
          console.error(`メール送信エラー (${item.email}):`, error);
        }
      }
    });
    
    return {
      success: true,
      sentCount: sentCount
    };
  } catch (error) {
    console.error('延滞通知の送信に失敗しました:', error);
    return {
      success: false,
      message: '延滞通知の送信に失敗しました'
    };
  }
}

// デフォルトの延滞通知テンプレート
function getDefaultOverdueTemplate() {
  return `{userName}様

いつも{libraryName}をご利用いただきありがとうございます。

現在、以下の書籍の返却期限が過ぎております。
速やかなご返却をお願いいたします。

書籍名: {bookTitle}
返却予定日: {dueDate}
延滞日数: {overdueDays}日

ご不明な点がございましたら、下記までお問い合わせください。
{libraryPhone}

{libraryName}`;
}

// ========================================
// ユーティリティ関数（HTML側から呼び出し用）
// ========================================
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ページコンテンツを取得する関数
function getPageContent(page) {
  const pageMap = {
    'bookRegister': 'book-register',
    'bookSearch': 'book-search',
    'userRegister': 'user-register',
    'userSearch': 'user-search',
    'loan': 'loan',
    'return': 'return',
    'reports': 'reports',
    'settings': 'settings'
  };
  
  const fileName = pageMap[page];
  if (!fileName) {
    throw new Error('ページが見つかりません: ' + page);
  }
  
  try {
    const template = HtmlService.createTemplateFromFile(fileName);
    return template.evaluate().getContent();
  } catch (error) {
    throw new Error('ページの読み込みに失敗しました: ' + error.message);
  }
}

// デプロイされたURLを取得する関数
function getDeployedUrl() {
  // スクリプトプロパティから取得を試みる
  const scriptProperties = PropertiesService.getScriptProperties();
  const customUrl = scriptProperties.getProperty('DEPLOYED_URL');
  
  if (customUrl) {
    return customUrl;
  }
  
  // ScriptApp.getService().getUrl()を試みる
  try {
    const url = ScriptApp.getService().getUrl();
    if (url) {
      return url;
    }
  } catch (e) {
    console.log('ScriptApp.getService().getUrl() failed:', e);
  }
  
  // 最後の手段として、現在知られているデプロイURLを返す
  // 注意: このURLは新しいデプロイごとに更新する必要があります
  return 'https://script.google.com/macros/s/AKfycbwaXDX2EHPULqIkv0J4gdcn5nnhHycoUExRdrn-THsnVOZniSD1mgRTMn3j8ZDjVsukCA/exec';
}

// バーコードスキャナーのサイドバーを表示（HTMLから呼び出される）
function showBarcodeSidebar() {
  try {
    const template = HtmlService.createTemplateFromFile('barcode-sidebar');
    const html = template.evaluate()
        .setTitle('バーコードスキャナー')
        .setWidth(350);
    
    // Web アプリケーションでは UI は使えないので、新しいウィンドウで開く
    return html.getContent();
  } catch (error) {
    console.error('サイドバーの表示に失敗:', error);
    throw error;
  }
}

// バーコード値を一時保存（セッション内で使用）
let tempBarcodeValue = null;

function setBarcodeValue(code) {
  tempBarcodeValue = code;
  // スクリプトプロパティに一時保存（複数セッション対応）
  const userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('tempBarcode', code);
  return true;
}

function getTempBarcodeValue() {
  // まずメモリから取得
  if (tempBarcodeValue) {
    const value = tempBarcodeValue;
    tempBarcodeValue = null;
    return value;
  }
  
  // メモリになければプロパティから取得
  const userProperties = PropertiesService.getUserProperties();
  const value = userProperties.getProperty('tempBarcode');
  if (value) {
    userProperties.deleteProperty('tempBarcode');
  }
  return value;
}