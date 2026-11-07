import crypto from "node:crypto";
import { waitUntil } from "@vercel/functions";

export const runtime = "nodejs";
export const maxDuration = 60;

function safeEqual(a, b) {
  const aa = Buffer.from(String(a || ""), "utf8");
  const bb = Buffer.from(String(b || ""), "utf8");
  return aa.length === bb.length && crypto.timingSafeEqual(aa, bb);
}

function requiredEnv(name) {
  const value = String(process.env[name] || "").trim();
  if (!value) throw new Error(`Missing environment variable: ${name}`);
  return value;
}

async function forwardToAppsScript(rawBody) {
  const appsScriptUrl = requiredEnv("APPS_SCRIPT_WEBHOOK_URL");
  const relaySecret = requiredEnv("VERCEL_RELAY_SECRET");

  const payload = JSON.stringify({
    relay_secret: relaySecret,
    line_raw_body: rawBody,
  });

  let lastError;

  for (let attempt = 0; attempt < 2; attempt++) {
    try {
      const res = await fetch(appsScriptUrl, {
        method: "POST",
        headers: {
          "content-type": "application/json",
        },
        body: payload,
        redirect: "follow",
      });

      const text = await res.text();

      if (res.ok && text.trim() === "OK") return;

      lastError = new Error(
        `Apps Script error: HTTP ${res.status}, body=${text.slice(0, 300)}`
      );
    } catch (err) {
      lastError = err;
    }

    if (attempt === 0) {
      await new Promise((resolve) => setTimeout(resolve, 1000));
    }
  }

  throw lastError || new Error("Apps Script forwarding failed");
}

export async function GET() {
  return Response.json({
    ok: true,
    service: "AI Study Coach LINE relay v2",
    lineSecretLoaded: Boolean(process.env.LINE_CHANNEL_SECRET),
    appsScriptUrlLoaded: Boolean(process.env.APPS_SCRIPT_WEBHOOK_URL),
    relaySecretLoaded: Boolean(process.env.VERCEL_RELAY_SECRET),
  });
}

export async function POST(request) {
  let rawBody;

  try {
    const channelSecret = requiredEnv("LINE_CHANNEL_SECRET");
    rawBody = await request.text();

    const receivedSignature = String(
      request.headers.get("x-line-signature") || ""
    ).trim();

    const expectedSignature = crypto
      .createHmac("sha256", channelSecret)
      .update(rawBody, "utf8")
      .digest("base64");

    if (!safeEqual(receivedSignature, expectedSignature)) {
      return Response.json(
        { ok: false, error: "Invalid LINE signature" },
        { status: 401 }
      );
    }

    const body = JSON.parse(rawBody);

    // LINE Developers の「検証」は events: []。
    if (Array.isArray(body.events) && body.events.length === 0) {
      return Response.json({ ok: true, verification: true }, { status: 200 });
    }

    // LINEにはすぐ200を返し、GASへの転送はレスポンス後も継続。
    waitUntil(
      forwardToAppsScript(rawBody).catch((err) => {
        console.error("Background forwarding failed:", err);
      })
    );

    return Response.json({ ok: true }, { status: 200 });
  } catch (err) {
    console.error("Webhook error:", err);

    return Response.json(
      {
        ok: false,
        error: "Webhook processing failed",
      },
      { status: 500 }
    );
  }
}
