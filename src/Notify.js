/**
 * StockPing - Notify.js
 * NotifyCore の呼び出しラッパー
 */

var Notify = (function() {
  function sendNotification(message) {
    Settings.init(PRODUCT_NAME);
    try {
      var config = {
        type:         Settings.get('WEBHOOK_TYPE') || '',
        webhookUrl:   Settings.get('WEBHOOK_URL')  || '',
        email:        Settings.get('EMAIL_TO')     || '',
        emailSubject: '[StockPing] 在庫不足アラート',
      };
      if (!config.webhookUrl && !config.email) return false;
      return notify(config, message);
    } catch (e) {
      console.error('[Notify] sendNotification: ' + e.message);
      return false;
    }
  }

  function sendTestNotification() {
    Settings.init(PRODUCT_NAME);
    try {
      var ts = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
      var msg = '[StockPing] テスト通知\n送信日時: ' + ts + '\n\n設定が正常に動作しています。';
      var result = sendNotification(msg);
      return result ? { success: true } : { success: false, error: '送信先が未設定か、送信に失敗しました。' };
    } catch (e) {
      return { success: false, error: e.message };
    }
  }

  return { sendNotification: sendNotification, sendTestNotification: sendTestNotification };
})();
