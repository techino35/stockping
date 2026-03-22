/**
 * StockPing - ColorCode.js
 * 在庫行の背景色自動更新ロジック
 */

var COLOR_OUT_OF_STOCK = '#fce8e6';
var COLOR_LOW_STOCK    = '#fef9e7';
var COLOR_NORMAL_STOCK = '#e6f4ea';
var COLOR_NO_DATA      = '#ffffff';

var ColorCode = (function() {
  function updateColors(sheet, config, items) {
    if (!items || items.length === 0) return;
    try {
      var lastCol = sheet.getLastColumn();
      if (lastCol === 0) return;
      items.forEach(function(item) {
        var color = _determineColor(item.current, item.minimum);
        try {
          sheet.getRange(item.row, 1, 1, lastCol).setBackground(color);
        } catch (e) {
          console.error('[ColorCode] row ' + item.row + ': ' + e.message);
        }
      });
    } catch (e) {
      console.error('[ColorCode] updateColors: ' + e.message);
    }
  }

  function _determineColor(current, minimum) {
    var cur = parseFloat(current);
    var min = parseFloat(minimum);
    if (isNaN(cur) || isNaN(min)) return COLOR_NO_DATA;
    if (cur <= 0 || cur < min * 0.5) return COLOR_OUT_OF_STOCK;
    if (cur < min) return COLOR_LOW_STOCK;
    return COLOR_NORMAL_STOCK;
  }

  function resetColors(sheet, headerRow) {
    try {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow <= headerRow || lastCol === 0) return;
      sheet.getRange(headerRow + 1, 1, lastRow - headerRow, lastCol).setBackground(COLOR_NO_DATA);
    } catch (e) {
      console.error('[ColorCode] resetColors: ' + e.message);
    }
  }

  return { updateColors: updateColors, resetColors: resetColors };
})();
