# AI担任 LINE Webhook

LINE Messaging API → Vercel → Google Apps Script の中継用です。

## Vercelに設定する環境変数
- `LINE_CHANNEL_SECRET`: LINE DevelopersのChannel secret
- `APPS_SCRIPT_WEBHOOK_URL`: Apps Scriptの `/exec` で終わるURL

## LINE Developersに登録するWebhook URL
VercelのURLが `https://xxxxx.vercel.app` の場合:
`https://xxxxx.vercel.app/api/webhook`

Vercel側でLINE署名を検証し、正常なWebhookだけApps Scriptへ転送します。
