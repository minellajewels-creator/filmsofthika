// api/gas.js
// Proxies POST requests from the frontend to GAS.
// GAS_URL is stored as a Vercel environment variable — never exposed to the browser.

export default async function handler(req, res) {
  // Allow GET for health check
  if (req.method === "GET") {
    return res.status(200).json({ ok: true });
  }

  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed" });
  }

  const GAS_URL = process.env.GAS_URL;
  if (!GAS_URL) {
    return res.status(500).json({ error: "GAS_URL environment variable not set." });
  }

  try {
    const body = typeof req.body === "string" ? req.body : JSON.stringify(req.body);

    const gasRes = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body,
    });

    const text = await gasRes.text();

    // Try to parse as JSON; fall back to raw text
    try {
      const json = JSON.parse(text);
      return res.status(200).json(json);
    } catch {
      return res.status(200).send(text);
    }
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
}
