const GAS_URL =
  "https://script.google.com/macros/s/AKfycbyO9RspNQ7h7DJRTGwnStojVAguA2KSqjPMjiKxeH9RludbFcXLl0HSPlLbOkJlMO5Zvg/exec";
 
module.exports = async function handler(req, res) {
  /* ── CORS headers ── */
  res.setHeader("Access-Control-Allow-Origin",  "*");
  res.setHeader("Access-Control-Allow-Methods", "POST, OPTIONS");
  res.setHeader("Access-Control-Allow-Headers", "Content-Type");
 
  if (req.method === "OPTIONS") return res.status(200).end();
  if (req.method !== "POST")
    return res.status(405).json({ ok: false, error: "Method not allowed" });
 
  try {
    /* Forward body to GAS exactly as received */
    const body = typeof req.body === "string"
      ? req.body
      : JSON.stringify(req.body);
 
    const gasRes = await fetch(GAS_URL, {
      method:   "POST",
      headers:  { "Content-Type": "application/json" },
      body,
      redirect: "follow"
    });
 
    const text = await gasRes.text();
 
    let data;
    try {
      data = JSON.parse(text);
    } catch (_) {
      data = { ok: false, error: "GAS returned non-JSON: " + text.slice(0, 200) };
    }
 
    return res.status(200).json(data);
  } catch (err) {
    return res.status(500).json({ ok: false, error: err.message });
  }
};
