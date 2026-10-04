import crypto from "crypto";

export const config = { api: { bodyParser: false } };

async function readRawBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk));
  return Buffer.concat(chunks);
}

function safeEqualBase64(a, b) {
  try {
    const ba = Buffer.from(a || "", "base64");
    const bb = Buffer.from(b || "", "base64");
    return ba.length === bb.length && crypto.timingSafeEqual(ba, bb);
  } catch {
    return false;
  }
}

export default async function handler(req, res) {
  if (req.method === "GET") return res.status(200).json({ ok: true, service: "AI担任 LINE Webhook" });
  if (req.method !== "POST") return res.status(405).json({ ok: false, error: "Method Not Allowed" });

  const channelSecret = process.env.LINE_CHANNEL_SECRET;
  const appsScriptUrl = process.env.APPS_SCRIPT_WEBHOOK_URL;
  if (!channelSecret || !appsScriptUrl) return res.status(500).json({ ok: false, error: "Server configuration error" });

  const rawBody = await readRawBody(req);
  const receivedSignature = req.headers["x-line-signature"] || "";
  const expectedSignature = crypto.createHmac("sha256", channelSecret).update(rawBody).digest("base64");

  if (!safeEqualBase64(receivedSignature, expectedSignature)) {
    return res.status(401).json({ ok: false, error: "Invalid signature" });
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
      console.error("Apps Script error", upstream.status, upstreamText);
      return res.status(502).json({ ok: false, error: "Apps Script upstream error", status: upstream.status });
    }
    return res.status(200).json({ ok: true });
  } catch (error) {
    console.error(error);
    return res.status(500).json({ ok: false, error: "Forwarding failed" });
  }
}
