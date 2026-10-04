import crypto from "node:crypto";

export const runtime = "nodejs";

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
  const appsScriptUrl = (process.env.APPS_SCRIPT_WEBHOOK_URL || "").trim();

  if (!appsScriptUrl) {
    return Response.json(
      { ok: false, error: "Missing APPS_SCRIPT_WEBHOOK_URL" },
      { status: 500 }
    );
  }

  const rawBody = await request.text();

  let body;

  try {
    body = JSON.parse(rawBody);
  } catch {
    return Response.json(
      { ok: false, error: "Invalid JSON" },
      { status: 400 }
    );
  }

  // LINE Developersの「検証」は events: [] で送られる。
  // Apps Scriptへ転送せず、即200を返す。
  if (Array.isArray(body.events) && body.events.length === 0) {
    return Response.json(
      {
        ok: true,
        verification: true,
      },
      { status: 200 }
    );
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

    if (!upstream.ok) {
      return Response.json(
        {
          ok: false,
          error: "Apps Script upstream error",
          status: upstream.status,
        },
        { status: 502 }
      );
    }

    return Response.json(
      { ok: true },
      { status: 200 }
    );

  } catch (error) {
    return Response.json(
      {
        ok: false,
        error: "Forwarding failed",
        message: String(error),
      },
      { status: 500 }
    );
  }
}
