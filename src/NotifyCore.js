/**
 * NotifyCore.js
 * 共通通知基盤 - Webhook送信 / MailApp送信モジュール
 */

function _buildSlackPayload(message) {
  return JSON.stringify({ text: message });
}

function _buildDiscordPayload(message) {
  return JSON.stringify({ content: message });
}

function _buildGoogleChatPayload(message) {
  return JSON.stringify({ text: message });
}

function _buildLinePayload(message) {
  return 'message=' + encodeURIComponent(message);
}

function sendWebhook(type, message, webhookUrl) {
  if (!webhookUrl || !message) return false;
  try {
    var options = { method: 'post', muteHttpExceptions: true };
    if (type === 'line') {
      var token = Settings.get('NOTIFY_LINE_TOKEN');
      options.headers = { 'Authorization': 'Bearer ' + token, 'Content-Type': 'application/x-www-form-urlencoded' };
      options.payload = _buildLinePayload(message);
    } else {
      options.contentType = 'application/json';
      switch (type) {
        case 'slack':      options.payload = _buildSlackPayload(message); break;
        case 'discord':    options.payload = _buildDiscordPayload(message); break;
        case 'googlechat': options.payload = _buildGoogleChatPayload(message); break;
        default: return false;
      }
    }
    var response = UrlFetchApp.fetch(webhookUrl, options);
    var code = response.getResponseCode();
    return code >= 200 && code < 300;
  } catch (e) {
    console.error('[NotifyCore] sendWebhook: ' + e.message);
    return false;
  }
}

function sendEmail(to, subject, body) {
  if (!to || !subject || !body) return false;
  try {
    var addresses = to.split(',').map(function(a) { return a.trim(); }).filter(Boolean);
    addresses.forEach(function(addr) { MailApp.sendEmail({ to: addr, subject: subject, body: body }); });
    return true;
  } catch (e) {
    console.error('[NotifyCore] sendEmail: ' + e.message);
    return false;
  }
}

function notify(config, message) {
  if (!config || !message) return false;
  var results = [];
  if (config.webhookUrl && config.type) results.push(sendWebhook(config.type, message, config.webhookUrl));
  if (config.email) results.push(sendEmail(config.email, config.emailSubject || '[通知]', message));
  if (results.length === 0) return false;
  return results.some(function(r) { return r === true; });
}
