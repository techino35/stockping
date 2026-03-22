/**
 * StockPing - Code.js
 * エントリーポイント: onInstall / onOpen / showSidebar / checkStock
 */

// プロダクト識別子（Settings.js の namespace に使用）
var PRODUCT_NAME = 'STOCKPING';

// ライセンス検証用ハッシュ済みキー（SHA-256）
// 実運用時は安全な管理方法に変更すること
var VALID_KEY_HASHES = [
  'a665a45920422f9d417e4867efdc4fb8a04a1f3fff1fa07e998e86f7f7a27ae3', // example key: 123
];

/**
 * アドオンインストール時の処理
 * @param {GoogleAppsScript.Events.AddonOnInstall} e
 */
function onInstall(e) {
  Settings.init(PRODUCT_NAME);
  onOpen(e);
}

/**
 * スプレッドシートオープン時の処理
 * @param {GoogleAppsScript.Events.SheetsOnOpen} e
 */
function onOpen(e) {
  Settings.init(PRODUCT_NAME);
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Open StockPing', 'showSidebar')
    .addItem('Check Stock Now', 'checkStock')
    .addItem('Update Color Codes', 'updateColorCodes')
    .addToUi();
}

/**
 * アドオンホームページカードを返す（Workspace Add-on用）
 * @returns {GoogleAppsScript.Card_Service.Card}
 */
function onHomepage() {
  Settings.init(PRODUCT_NAME);
  showSidebar();
  return CardService.newCardBuilder().build();
}

/**
 * サイドバーを表示する
 */
function showSidebar() {
  Settings.init(PRODUCT_NAME);
  try {
    var html = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setTitle('StockPing')
      .setWidth(320);
    SpreadsheetApp.getUi().showSidebar(html);
  } catch (e) {
    console.error('[StockPing] showSidebar: ' + e.message);
    SpreadsheetApp.getUi().alert('StockPing でエラーが発生しました: ' + e.message);
  }
}

/**
 * 在庫チェックのメイン処理（タイマートリガー / 手動実行）
 * 現在庫 < 最低在庫 の品目を検出して通知する
 */
function checkStock() {
  Settings.init(PRODUCT_NAME);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = _loadStockConfig();

    if (!config.sheetName) {
      console.log('[StockPing] checkStock: sheet not configured');
      return;
    }

    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      console.error('[StockPing] checkStock: sheet not found: ' + config.sheetName);
      SpreadsheetApp.getUi().alert('StockPing: シート「' + config.sheetName + '」が見つかりません。');
      return;
    }

    var isPro = _isProUser();
    var items = _getStockItems(sheet, config, isPro);

    if (items.length === 0) {
      console.log('[StockPing] checkStock: no stock data found');
      return;
    }

    // カラーコード更新
    ColorCode.updateColors(sheet, config, items);

    // 在庫不足品目を抽出
    var lowItems = items.filter(function(item) {
      return item.current !== '' && item.minimum !== '' &&
             parseFloat(item.current) < parseFloat(item.minimum);
    });

    if (lowItems.length === 0) {
      console.log('[StockPing] checkStock: all stock levels are OK');
      return;
    }

    var message = _buildStockAlert(lowItems);
    Notify.sendNotification(message);

    console.log('[StockPing] checkStock: ' + lowItems.length + ' low stock items found');

  } catch (err) {
    console.error('[StockPing] checkStock: ' + err.message);
    SpreadsheetApp.getUi().alert('StockPing チェックエラー: ' + err.message);
  }
}

/**
 * カラーコードのみ更新する（メニューから手動実行）
 */
function updateColorCodes() {
  Settings.init(PRODUCT_NAME);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var config = _loadStockConfig();

    if (!config.sheetName) {
      SpreadsheetApp.getUi().alert('StockPing: 先にシートを設定してください。');
      return;
    }

    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      SpreadsheetApp.getUi().alert('StockPing: シート「' + config.sheetName + '」が見つかりません。');
      return;
    }

    var isPro = _isProUser();
    var items = _getStockItems(sheet, config, isPro);
    ColorCode.updateColors(sheet, config, items);
    console.log('[StockPing] updateColorCodes: done');

  } catch (err) {
    console.error('[StockPing] updateColorCodes: ' + err.message);
    SpreadsheetApp.getUi().alert('StockPing カラー更新エラー: ' + err.message);
  }
}

// ===== 設定管理（サイドバー呼び出し用） =====

function getStockSettings() {
  Settings.init(PRODUCT_NAME);
  return {
    sheetName:       Settings.get('SHEET_NAME')        || '',
    headerRow:       parseInt(Settings.get('HEADER_ROW') || '1', 10),
    colItem:         Settings.get('COL_ITEM')          || '',
    colCurrent:      Settings.get('COL_CURRENT')       || '',
    colMinimum:      Settings.get('COL_MINIMUM')       || '',
    colOrderEmail:   Settings.get('COL_ORDER_EMAIL')   || '',
    triggerInterval: Settings.get('TRIGGER_INTERVAL')  || 'daily',
    webhookType:     Settings.get('WEBHOOK_TYPE')      || '',
    webhookUrl:      Settings.get('WEBHOOK_URL')       || '',
    emailTo:         Settings.get('EMAIL_TO')          || '',
  };
}

function saveStockSettings(config) {
  Settings.init(PRODUCT_NAME);
  try {
    Settings.setAll({
      SHEET_NAME:       config.sheetName       || '',
      HEADER_ROW:       String(config.headerRow || '1'),
      COL_ITEM:         config.colItem         || '',
      COL_CURRENT:      config.colCurrent      || '',
      COL_MINIMUM:      config.colMinimum      || '',
      COL_ORDER_EMAIL:  config.colOrderEmail   || '',
      TRIGGER_INTERVAL: config.triggerInterval || 'daily',
      WEBHOOK_TYPE:     config.webhookType     || '',
      WEBHOOK_URL:      config.webhookUrl      || '',
      EMAIL_TO:         config.emailTo         || '',
    });
    _resetTrigger(config.triggerInterval || 'daily');
    return { success: true };
  } catch (e) {
    console.error('[StockPing] saveStockSettings: ' + e.message);
    return { success: false, error: e.message };
  }
}

function detectColumns(sheetName) {
  Settings.init(PRODUCT_NAME);
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: false, error: 'シートが見つかりません: ' + sheetName };
    var headerRow = parseInt(Settings.get('HEADER_ROW') || '1', 10);
    var mapping = Detect.detectColumns(sheet, headerRow);
    return { success: true, mapping: mapping };
  } catch (e) {
    console.error('[StockPing] detectColumns: ' + e.message);
    return { success: false, error: e.message };
  }
}

function getSheetNames() {
  try {
    return SpreadsheetApp.getActiveSpreadsheet().getSheets().map(function(s) { return s.getName(); });
  } catch (e) {
    return [];
  }
}

function getLicenseInfo() {
  Settings.init(PRODUCT_NAME);
  var isPro = _isProUser();
  return { isPro: isPro, itemLimit: isPro ? null : 10 };
}

function validateLicense(key) {
  Settings.init(PRODUCT_NAME);
  try {
    if (!key) return false;
    var hash = _sha256(key.trim());
    var valid = VALID_KEY_HASHES.indexOf(hash) !== -1;
    if (valid) {
      Settings.set('LICENSE_KEY', key.trim());
      Settings.set('LICENSE_STATUS', 'pro');
    }
    return valid;
  } catch (e) {
    return false;
  }
}

// ===== 内部ヘルパー =====

function _loadStockConfig() {
  return {
    sheetName:     Settings.get('SHEET_NAME')        || '',
    headerRow:     parseInt(Settings.get('HEADER_ROW') || '1', 10),
    colItem:       Settings.get('COL_ITEM')          || '',
    colCurrent:    Settings.get('COL_CURRENT')       || '',
    colMinimum:    Settings.get('COL_MINIMUM')       || '',
    colOrderEmail: Settings.get('COL_ORDER_EMAIL')   || '',
  };
}

function _getStockItems(sheet, config, isPro) {
  var lastRow = sheet.getLastRow();
  var headerRow = config.headerRow || 1;
  if (lastRow <= headerRow) return [];
  var dataStartRow = headerRow + 1;
  var dataRowCount = lastRow - headerRow;
  var colItem       = _colLetterToIndex(config.colItem);
  var colCurrent    = _colLetterToIndex(config.colCurrent);
  var colMinimum    = _colLetterToIndex(config.colMinimum);
  var colOrderEmail = _colLetterToIndex(config.colOrderEmail);
  if (!colItem || !colCurrent || !colMinimum) {
    var detected = Detect.detectColumns(sheet, headerRow);
    colItem       = colItem       || _colLetterToIndex(detected.item);
    colCurrent    = colCurrent    || _colLetterToIndex(detected.current);
    colMinimum    = colMinimum    || _colLetterToIndex(detected.minimum);
    colOrderEmail = colOrderEmail || _colLetterToIndex(detected.orderEmail);
  }
  if (!colItem || !colCurrent || !colMinimum) return [];
  var maxCol = Math.max(colItem, colCurrent, colMinimum, colOrderEmail || 0);
  var values = sheet.getRange(dataStartRow, 1, dataRowCount, maxCol).getValues();
  var items = [];
  var limit = isPro ? values.length : Math.min(values.length, 10);
  for (var i = 0; i < limit; i++) {
    var row = values[i];
    var name = row[colItem - 1];
    if (!name) continue;
    items.push({
      row: dataStartRow + i,
      name: String(name),
      current: row[colCurrent - 1],
      minimum: row[colMinimum - 1],
      orderEmail: colOrderEmail ? String(row[colOrderEmail - 1] || '') : '',
      colCurrent: colCurrent,
    });
  }
  return items;
}

function _buildStockAlert(lowItems) {
  var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
  var lines = ['[StockPing] 在庫不足アラート (' + ts + ')', ''];
  lowItems.forEach(function(item) {
    var recommended = Math.max(1, Math.ceil(parseFloat(item.minimum) * 2 - parseFloat(item.current)));
    lines.push('品目: ' + item.name);
    lines.push('  現在庫: ' + item.current + ' / 最低在庫: ' + item.minimum);
    lines.push('  推奨発注数: ' + recommended);
    if (item.orderEmail) lines.push('  発注先: ' + item.orderEmail);
    lines.push('');
  });
  lines.push('合計 ' + lowItems.length + ' 品目の在庫補充が必要です。');
  return lines.join('\n').trim();
}

function _resetTrigger(interval) {
  try {
    ScriptApp.getProjectTriggers().forEach(function(t) {
      if (t.getHandlerFunction() === 'checkStock') ScriptApp.deleteTrigger(t);
    });
    if (interval === '6hours') {
      ScriptApp.newTrigger('checkStock').timeBased().everyHours(6).create();
    } else {
      ScriptApp.newTrigger('checkStock').timeBased().everyDays(1).atHour(9).create();
    }
  } catch (e) {
    console.error('[StockPing] _resetTrigger: ' + e.message);
  }
}

function _isProUser() {
  return Settings.get('LICENSE_STATUS') === 'pro';
}

function _colLetterToIndex(letter) {
  if (!letter) return null;
  var upper = String(letter).trim().toUpperCase();
  var index = 0;
  for (var i = 0; i < upper.length; i++) {
    index = index * 26 + (upper.charCodeAt(i) - 64);
  }
  return index || null;
}

function _sha256(input) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, input, Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    var hex = (b < 0 ? b + 256 : b).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}
