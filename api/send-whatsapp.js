// Vercel Serverless Function — WhatsApp Cloud API
// Receives a PDF as base64, uploads to Meta, sends as WhatsApp document
// Env vars required: WHATSAPP_PHONE_ID, WHATSAPP_TOKEN

export default async function handler(req, res) {
  // Allow CORS from same origin
  res.setHeader("Access-Control-Allow-Origin", "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST") return res.status(405).json({ error: "Method not allowed" });

  const PHONE_ID = process.env.WHATSAPP_PHONE_ID;
  const TOKEN    = process.env.WHATSAPP_TOKEN;

  if (!PHONE_ID || !TOKEN) {
    // API not configured yet — tell app to fall back to manual share
    return res.status(503).json({ error: "WhatsApp API not configured" });
  }

  const { phone, pdfBase64, filename, message } = req.body || {};

  if (!phone || !pdfBase64) {
    return res.status(400).json({ error: "Missing phone or pdfBase64" });
  }

  // Clean phone number — keep only last 10 digits, prepend country code 91
  const cleanPhone = "91" + String(phone).replace(/\D/g, "").slice(-10);
  const fname = filename || "MR_NKD_Bajaj.pdf";

  try {
    // ── STEP 1: Upload PDF to Meta media servers ──────────────────────────
    const buf  = Buffer.from(pdfBase64, "base64");
    const form = new FormData();
    form.append("messaging_product", "whatsapp");
    form.append("type", "application/pdf");
    form.append("file", new Blob([buf], { type: "application/pdf" }), fname);

    const uploadRes = await fetch(
      `https://graph.facebook.com/v20.0/${PHONE_ID}/media`,
      { method: "POST", headers: { Authorization: `Bearer ${TOKEN}` }, body: form }
    );
    const uploadData = await uploadRes.json();

    if (!uploadData.id) {
      console.error("Meta media upload failed:", JSON.stringify(uploadData));
      return res.status(502).json({ error: "Media upload failed", detail: uploadData });
    }

    // ── STEP 2: Send document message ─────────────────────────────────────
    const msgRes = await fetch(
      `https://graph.facebook.com/v20.0/${PHONE_ID}/messages`,
      {
        method: "POST",
        headers: { Authorization: `Bearer ${TOKEN}`, "Content-Type": "application/json" },
        body: JSON.stringify({
          messaging_product: "whatsapp",
          to: cleanPhone,
          type: "document",
          document: {
            id: uploadData.id,
            caption: message || "Please find your document from NKD Bajaj, Dhanbad.",
            filename: fname,
          },
        }),
      }
    );
    const msgData = await msgRes.json();

    if (msgData.error) {
      console.error("Meta send failed:", JSON.stringify(msgData));
      return res.status(502).json({ error: "WhatsApp send failed", detail: msgData });
    }

    return res.json({ success: true, messageId: msgData.messages?.[0]?.id });
  } catch (e) {
    console.error("send-whatsapp error:", e.message);
    return res.status(500).json({ error: e.message });
  }
}
