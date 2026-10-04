from pathlib import Path

code = '''import crypto from "node:crypto";

export const runtime = "nodejs";

function safeEqual(a, b) {
  const aa = Buffer.from(String(a || ""), "utf8");
  const bb = Buffer.from(String(b || ""), "utf8");
  return aa.length === bb.length && crypto.timingSafeEqual(aa, bb);
}

export async function GET() {
  const channelSecret = (process.env.LINE_CHANNEL_SECRET || "").trim();
  const appsScriptUrl = (process.env.APPS_SCRIPT_WEBHOOK_URL || "").trim();

  return Response.json({
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

export async function POST(request) {
  const channelSecret = (process.env.LINE_CHANNEL_SECRET || "").trim();
  const appsScriptUrl = (process.env.APPS_SCRIPT_WEBHOOK_URL || "").trim();

  if (!channelSecret || !appsScriptUrl) {
    return Response.json(
      { ok: false, error: "Missing environment variables" },
      { status: 500 }
    );
  }

  // LINEから届いた本文を文字列のまま取得。
  // JSON.parse等は署名検証が終わるまで行わない。
  const rawBody = await request.text();
  const receivedSignature =
    (request.headers.get("x-line-signature") || "").trim();

  const expectedSignature = crypto
    .createHmac("sha256", channelSecret)
    .update(rawBody, "utf8")
    .digest("base64");

  if (!safeEqual(receivedSignature, expectedSignature)) {
    console.error("LINE signature mismatch", {
      signatureHeaderPresent: receivedSignature.length > 0,
      receivedSignatureLength: receivedSignature.length,
      expectedSignatureLength: expectedSignature.length,
      rawBodyLength: Buffer.byteLength(rawBody, "utf8"),
    });

    return Response.json(
      {
        ok: false,
        error: "Invalid LINE signature",
        signatureHeaderPresent: receivedSignature.length > 0,
        receivedSignatureLength: receivedSignature.length,
        expectedSignatureLength: expectedSignature.length,
        rawBodyLength: Buffer.byteLength(rawBody, "utf8"),
      },
      { status: 401 }
    );
  }

  try {
    const upstream = await fetch(appsScriptUrl, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: rawBody,
      redirect: "follow",
    });

    const upstreamText = await upstream.text();

    if (!upstream.ok) {
      console.error("Apps Script upstream error", {
        status: upstream.status,
        bodyPreview: upstreamText.slice(0, 200),
      });

      return Response.json(
        {
          ok: false,
          error: "Apps Script upstream error",
          status: upstream.status,
        },
        { status: 502 }
      );
    }

    return Response.json({ ok: true }, { status: 200 });
  } catch (error) {
    console.error("Forwarding failed", error);

    return Response.json(
      { ok: false, error: "Forwarding failed" },
      { status: 500 }
    );
  }
}
'''

path = Path("/mnt/data/route.js")
path.write_text(code, encoding="utf-8")
print(path)
