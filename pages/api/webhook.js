from pathlib import Path

code = r'''import crypto from "crypto";

export const config = {
  api: {
    bodyParser: false,
  },
};

async function readRawBody(req) {
  const chunks = [];
  for await (const chunk of req) {
    chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  }
  return Buffer.concat(chunks);
}

function timingSafeEqualText(a, b) {
  const aa = Buffer.from(String(a || ""), "utf8");
  const bb = Buffer.from(String(b || ""), "utf8");
  return aa.length === bb.length && crypto.timingSafeEqual(aa, bb);
}

export default async function handler(req, res) {
  const channelSecret = (process.env.LINE_CHANNEL_SECRET || "").trim();
  const appsScriptUrl = (process.env.APPS_SCRIPT_WEBHOOK_URL || "").trim();

  // ブラウザで開いたときの診断用。secretそのものは表示しません。
  if (req.method === "GET") {
    return res.status(200).json({
      ok: true,
      service: "AI担任 LINE Webhook",
      lineSecretLoaded: channelSecret.length > 0,
      lineSecretLength: channelSecret.length,
      appsScriptUrlLoaded: appsScriptUrl.length > 0,
      appsScriptUrlLooksValid:
        appsScriptUrl.startsWith("https://script.google.com/") &&
        appsScriptUrl.endsWith("/exec"),
    });
  }

  if (req.method !== "POST") {
    res.setHeader("Allow", "GET, POST");
    return res.status(405).json({ ok: false, error: "Method Not Allowed" });
  }

  if (!channelSecret || !appsScriptUrl) {
    return res.status(500).json({
      ok: false,
      error: "Missing environment variables",
      lineSecretLoaded: channelSecret.length > 0,
      appsScriptUrlLoaded: appsScriptUrl.length > 0,
    });
  }

  const rawBody = await readRawBody(req);
  const receivedSignature = String(req.headers["x-line-signature"] || "").trim();

  const expectedSignature = crypto
    .createHmac("sha256", channelSecret)
    .update(rawBody)
    .digest("base64");

  // LINEの「検証」は空のeventsを送ります。
  // 署名不一致時は診断用の情報だけ返し、secret自体は絶対に返しません。
  if (!timingSafeEqualText(receivedSignature, expectedSignature)) {
    return res.status(401).json({
      ok: false,
      error: "Invalid LINE signature",
      signatureHeaderPresent: receivedSignature.length > 0,
      receivedSignatureLength: receivedSignature.length,
      expectedSignatureLength: expectedSignature.length,
      rawBodyLength: rawBody.length,
    });
  }

  try {
    const upstream = await fetch(appsScriptUrl, {
      method: "POST",
      headers: {
        "content-type": "application/json",
      },
      body: rawBody,
      redirect: "follow",
    });

    const upstreamText = await upstream.text();

    if (!upstream.ok) {
      return res.status(502).json({
        ok: false,
        error: "Apps Script upstream error",
        status: upstream.status,
        bodyPreview: upstreamText.slice(0, 200),
      });
    }

    return res.status(200).json({ ok: true });
  } catch (error) {
    return res.status(500).json({
      ok: false,
      error: "Forwarding failed",
      message: String(error),
    });
  }
}
'''

path = Path("/mnt/data/webhook.js")
path.write_text(code, encoding="utf-8")
print(path)
