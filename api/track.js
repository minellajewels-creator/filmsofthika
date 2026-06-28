// api/track.js
// Called when a visitor clicks a product link on the storefront.
// Records the click in GAS (background) and immediately redirects to the affiliate URL.
// The GAS_URL never leaves the server.

export default async function handler(req, res) {
  const { pid } = req.query; // product_id

  if (!pid) {
    return res.status(400).send("Missing product id.");
  }

  const GAS_URL = process.env.GAS_URL;
  if (!GAS_URL) {
    return res.status(500).send("Server misconfigured.");
  }

  try {
    // Ask GAS for the product's affiliate URL
    const gasRes = await fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ action: "getProductForTrack", product_id: pid }),
    });

    const data = await gasRes.json();

    if (!data || data.error || (!data.affiliate_url && !data.regular_url)) {
      return res.status(404).send("Link not found.");
    }

    // Record the click in the background (fire-and-forget via GAS)
    fetch(GAS_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        action: "recordClick",
        product_id: pid,
        card_id:    data.card_id,
        ua:         req.headers["user-agent"] || "",
      }),
    }).catch(() => {}); // intentionally non-blocking

    const destination = data.affiliate_url || data.regular_url;
    res.setHeader("Cache-Control", "no-store");
    return res.redirect(302, destination);
  } catch (err) {
    return res.status(500).send("Error: " + err.message);
  }
}
