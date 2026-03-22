# StockPing

## English

### Overview
StockPing is a Google Sheets add-on that monitors inventory levels and sends instant alerts when stock falls below minimum thresholds — via Slack, Discord, Google Chat, LINE Notify, or email.

### Features
- **Auto-detect columns**: Automatically identifies item, current stock, minimum stock, and order email columns from your header row
- **Smart color coding**: Rows update automatically — red (out of stock), yellow (low stock), green (normal)
- **Order recommendations**: Each alert includes recommended reorder quantity (min × 2 − current)
- **Multi-channel notifications**: Slack / Discord / Google Chat / LINE Notify / Email
- **Scheduled checks**: Daily at 9:00 AM or every 6 hours
- Free plan: up to 10 items
- Pro plan: unlimited items + LINE Notify support

### Quick Start
1. Open Extensions > StockPing > Open StockPing
2. Go to the **Columns** tab and select your sheet
3. Click **Auto-Detect Columns**
4. Configure notification destination in **Settings** tab
5. Click **Save Settings**

### File Structure
```
src/
├── Code.js        # Entry points
├── Detect.js      # Column auto-detection
├── ColorCode.js   # Row background color logic
├── Notify.js      # Notification wrapper
├── NotifyCore.js  # Webhook/email sender
├── Settings.js    # PropertiesService wrapper
└── Sidebar.html   # UI
```

### Color Code Reference
| Color | Meaning |
|---|---|
| Red (#fce8e6) | Out of stock (current ≤ 0 or < min × 50%) |
| Yellow (#fef9e7) | Low stock (current < minimum) |
| Green (#e6f4ea) | Normal (current ≥ minimum) |

### License
- Free: 10 items max
- Pro: Unlimited items + LINE Notify

---

## 日本語

### 概要
StockPingは、Googleスプレッドシートの在庫レベルを監視し、最低在庫を下回った際にSlack・Discord・Google Chat・LINE Notify・メールへ即座にアラートを送るアドオンです。

### MIT License
