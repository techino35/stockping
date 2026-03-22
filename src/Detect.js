/**
 * StockPing - Detect.js
 * 列自動検出ロジック（fuzzyMatch関数含む）
 */

var ITEM_KEYWORDS = ['品目', '商品名', '品名', '商品', 'アイテム', 'item', 'product', 'name', 'goods', '在庫品'];
var CURRENT_KEYWORDS = ['現在庫', '在庫数', '在庫量', '現在の在庫', '現在', 'current', 'stock', 'quantity', 'qty', '数量', '在庫'];
var MINIMUM_KEYWORDS = ['最低在庫', '最小在庫', '安全在庫', '最低', '閾値', 'minimum', 'min', 'min stock', 'safety', 'threshold', 'reorder'];
var ORDER_EMAIL_KEYWORDS = ['発注先', '発注先メール', '仕入先', '仕入れ先', 'メール', 'email', 'supplier', 'vendor', 'order email', 'contact'];

var Detect = (function() {
  function detectColumns(sheet, headerRow) {
    try {
      var lastCol = sheet.getLastColumn();
      if (lastCol === 0) return _emptyMapping();
      var headers = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
      var mapping = { item: '', current: '', minimum: '', orderEmail: '' };
      headers.forEach(function(header, index) {
        if (!header) return;
        var h = String(header).trim().toLowerCase();
        var colLetter = _indexToColLetter(index + 1);
        if (!mapping.item && fuzzyMatch(h, ITEM_KEYWORDS)) mapping.item = colLetter;
        else if (!mapping.current && fuzzyMatch(h, CURRENT_KEYWORDS)) mapping.current = colLetter;
        else if (!mapping.minimum && fuzzyMatch(h, MINIMUM_KEYWORDS)) mapping.minimum = colLetter;
        else if (!mapping.orderEmail && fuzzyMatch(h, ORDER_EMAIL_KEYWORDS)) mapping.orderEmail = colLetter;
      });
      return mapping;
    } catch (e) {
      console.error('[Detect] detectColumns: ' + e.message);
      return _emptyMapping();
    }
  }

  function fuzzyMatch(header, keywords) {
    if (!header || !keywords || keywords.length === 0) return false;
    var h = header.toLowerCase().trim();
    for (var i = 0; i < keywords.length; i++) {
      if (h === keywords[i].toLowerCase()) return true;
    }
    for (var j = 0; j < keywords.length; j++) {
      if (h.indexOf(keywords[j].toLowerCase()) !== -1) return true;
    }
    for (var k = 0; k < keywords.length; k++) {
      if (keywords[k].toLowerCase().indexOf(h) !== -1 && h.length >= 2) return true;
    }
    return false;
  }

  function _indexToColLetter(index) {
    var letter = '';
    while (index > 0) {
      var remainder = (index - 1) % 26;
      letter = String.fromCharCode(65 + remainder) + letter;
      index = Math.floor((index - 1) / 26);
    }
    return letter;
  }

  function _emptyMapping() {
    return { item: '', current: '', minimum: '', orderEmail: '' };
  }

  return { detectColumns: detectColumns, fuzzyMatch: fuzzyMatch };
})();
