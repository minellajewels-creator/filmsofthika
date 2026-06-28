// ============================================================
// filmsofthika's bag — Google Apps Script Backend
// ============================================================

const SHEET_CARDS    = "Cards";
const SHEET_PRODUCTS = "Products";
const SHEET_CLICKS   = "Clicks";

// ── INIT ──────────────────────────────────────────────────────
function initiateSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  function ensure(name, headers) {
    let sh = ss.getSheetByName(name);
    if (!sh) {
      sh = ss.insertSheet(name);
      sh.appendRow(headers);
      sh.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
    }
    return sh;
  }

  ensure(SHEET_CARDS, [
    "card_id", "reel_url", "created_at", "domain"
  ]);

  ensure(SHEET_PRODUCTS, [
    "product_id", "card_id", "name", "price", "regular_url",
    "affiliate_url", "image_url", "platform", "created_at"
  ]);

  ensure(SHEET_CLICKS, [
    "click_id", "product_id", "card_id", "timestamp", "user_agent"
  ]);

  SpreadsheetApp.getUi().alert("✅ Sheets initialised: Cards, Products, Clicks");
}

// ── DOGET — serve storefront, redirect, or JSONP API ─────────
function doGet(e) {
  const params = e.parameter;

  // JSONP: /exec?callback=fn&action=getCards&...
  if (params.callback) {
    try {
      const result = routeAction(params);
      const json   = JSON.stringify(result);
      return ContentService
        .createTextOutput(params.callback + "(" + json + ")")
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    } catch (err) {
      return ContentService
        .createTextOutput(params.callback + '({"error":"' + err.message + '"})')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
  }

  // /exec?card=CARD_ID  → serve the storefront page for that card
  if (params.card) {
    const tmpl = HtmlService.createTemplateFromFile("Index");
    tmpl.cardId = params.card;
    return tmpl.evaluate()
      .setTitle("filmsofthika's bag")
      .addMetaTag("viewport", "width=device-width, initial-scale=1")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // /exec?redirect=PRODUCT_ID  → track click then redirect
  if (params.redirect) {
    return trackAndRedirect(params.redirect, e);
  }

  // Default
  return HtmlService.createHtmlOutput("<h2>No card specified.</h2>");
}

// Routes actions from both JSONP (GET) and POST
function routeAction(params) {
  const action = params.action;
  if (action === "getCards")           return getCards();
  if (action === "getCard")            return getCard(params.card_id);
  if (action === "getAnalytics")       return getAnalytics();
  if (action === "addCard")            return addCard(params);
  if (action === "updateCard")         return updateCard(params);
  if (action === "deleteCard")         return deleteCard(params);
  if (action === "addProduct")         return addProduct(params);
  if (action === "updateProduct")      return updateProduct(params);
  if (action === "deleteProduct")      return deleteProduct(params);
  if (action === "fetchMeta")          return fetchMeta(params.url);
  if (action === "getProductForTrack") return getProductForTrack(params.product_id);
  if (action === "recordClick")        return recordClick(params);
  return { error: "Unknown action" };
}

// ── DOPOST — JSON API for admin.html ─────────────────────────
function doPost(e) {
  try {
    const body   = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "addCard")       return jsonResponse(addCard(body));
    if (action === "updateCard")    return jsonResponse(updateCard(body));
    if (action === "deleteCard")    return jsonResponse(deleteCard(body));
    if (action === "addProduct")    return jsonResponse(addProduct(body));
    if (action === "updateProduct") return jsonResponse(updateProduct(body));
    if (action === "deleteProduct") return jsonResponse(deleteProduct(body));
    if (action === "getCards")           return jsonResponse(getCards());
    if (action === "getCard")            return jsonResponse(getCard(body.card_id));
    if (action === "getAnalytics")       return jsonResponse(getAnalytics());
    if (action === "generateLink")       return jsonResponse(generateLink(body));
    if (action === "fetchMeta")          return jsonResponse(fetchMeta(body.url));
    if (action === "getProductForTrack") return jsonResponse(getProductForTrack(body.product_id));
    if (action === "recordClick")        return jsonResponse(recordClick(body));

    return jsonResponse({ error: "Unknown action" });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ── HELPERS ───────────────────────────────────────────────────
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function uid() {
  return Utilities.getUuid().split("-")[0];
}

function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function sheetToObjects(sheet) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row[i]; });
    return obj;
  });
}

// ── CARDS ─────────────────────────────────────────────────────
function addCard(body) {
  const sh      = getSheet(SHEET_CARDS);
  const card_id = uid();
  sh.appendRow([card_id, body.reel_url, new Date().toISOString(), body.domain || ""]);
  return { success: true, card_id };
}

function updateCard(body) {
  const sh   = getSheet(SHEET_CARDS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.card_id) {
      if (body.reel_url !== undefined) sh.getRange(i + 1, 2).setValue(body.reel_url);
      if (body.domain   !== undefined) sh.getRange(i + 1, 4).setValue(body.domain);
      return { success: true };
    }
  }
  return { error: "Card not found" };
}

function deleteCard(body) {
  const sh   = getSheet(SHEET_CARDS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.card_id) {
      sh.deleteRow(i + 1);
      // Also delete products
      deleteProductsByCard(body.card_id);
      return { success: true };
    }
  }
  return { error: "Card not found" };
}

function getCards() {
  const cards    = sheetToObjects(getSheet(SHEET_CARDS));
  const products = sheetToObjects(getSheet(SHEET_PRODUCTS));
  const clicks   = sheetToObjects(getSheet(SHEET_CLICKS));

  return cards.map(c => {
    const prods = products.filter(p => p.card_id === c.card_id).map(p => {
      const clickCount = clicks.filter(cl => cl.product_id === p.product_id).length;
      return { ...p, clicks: clickCount };
    });
    return { ...c, products: prods };
  });
}

function getCard(card_id) {
  const cards    = sheetToObjects(getSheet(SHEET_CARDS));
  const products = sheetToObjects(getSheet(SHEET_PRODUCTS));
  const clicks   = sheetToObjects(getSheet(SHEET_CLICKS));

  const card = cards.find(c => c.card_id === card_id);
  if (!card) return { error: "Card not found" };

  card.products = products
    .filter(p => p.card_id === card_id)
    .map(p => ({
      ...p,
      clicks: clicks.filter(cl => cl.product_id === p.product_id).length
    }));

  return card;
}

// ── PRODUCTS ──────────────────────────────────────────────────
function addProduct(body) {
  const sh         = getSheet(SHEET_PRODUCTS);
  const product_id = uid();
  sh.appendRow([
    product_id,
    body.card_id,
    body.name          || "",
    body.price         || "",
    body.regular_url   || "",
    body.affiliate_url || "",
    body.image_url     || "",
    body.platform      || detectPlatform(body.regular_url),
    new Date().toISOString()
  ]);
  return { success: true, product_id };
}

function updateProduct(body) {
  const sh   = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  const keys = ["product_id","card_id","name","price","regular_url","affiliate_url","image_url","platform","created_at"];
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.product_id) {
      if (body.name          !== undefined) sh.getRange(i+1, keys.indexOf("name")+1).setValue(body.name);
      if (body.price         !== undefined) sh.getRange(i+1, keys.indexOf("price")+1).setValue(body.price);
      if (body.regular_url   !== undefined) sh.getRange(i+1, keys.indexOf("regular_url")+1).setValue(body.regular_url);
      if (body.affiliate_url !== undefined) sh.getRange(i+1, keys.indexOf("affiliate_url")+1).setValue(body.affiliate_url);
      if (body.image_url     !== undefined) sh.getRange(i+1, keys.indexOf("image_url")+1).setValue(body.image_url);
      if (body.platform      !== undefined) sh.getRange(i+1, keys.indexOf("platform")+1).setValue(body.platform);
      return { success: true };
    }
  }
  return { error: "Product not found" };
}

function deleteProduct(body) {
  const sh   = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === body.product_id) {
      sh.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { error: "Product not found" };
}

function deleteProductsByCard(card_id) {
  const sh   = getSheet(SHEET_PRODUCTS);
  const rows = sh.getDataRange().getValues();
  for (let i = rows.length - 1; i >= 1; i--) {
    if (rows[i][1] === card_id) sh.deleteRow(i + 1);
  }
}

// ── CLICK TRACKING ────────────────────────────────────────────
function trackAndRedirect(product_id, e) {
  try {
    const products = sheetToObjects(getSheet(SHEET_PRODUCTS));
    const product  = products.find(p => p.product_id === product_id);
    if (!product) return HtmlService.createHtmlOutput("<p>Link not found.</p>");

    const sh = getSheet(SHEET_CLICKS);
    sh.appendRow([
      uid(),
      product_id,
      product.card_id,
      new Date().toISOString(),
      e.parameter.ua || ""
    ]);

    // Redirect
    const url = product.affiliate_url || product.regular_url;
    return HtmlService.createHtmlOutput(
      `<html><head><meta http-equiv="refresh" content="0;url=${url}"></head>
       <body>Redirecting...</body></html>`
    );
  } catch (err) {
    return HtmlService.createHtmlOutput("<p>Error tracking click.</p>");
  }
}

// ── ANALYTICS ─────────────────────────────────────────────────
function getAnalytics() {
  const cards    = sheetToObjects(getSheet(SHEET_CARDS));
  const products = sheetToObjects(getSheet(SHEET_PRODUCTS));
  const clicks   = sheetToObjects(getSheet(SHEET_CLICKS));

  const totalClicks = clicks.length;

  const byProduct = products.map(p => ({
    product_id:    p.product_id,
    card_id:       p.card_id,
    name:          p.name,
    platform:      p.platform,
    affiliate_url: p.affiliate_url,
    clicks:        clicks.filter(c => c.product_id === p.product_id).length
  })).sort((a, b) => b.clicks - a.clicks);

  const byCard = cards.map(c => ({
    card_id:  c.card_id,
    reel_url: c.reel_url,
    clicks:   clicks.filter(cl => cl.card_id === c.card_id).length
  })).sort((a, b) => b.clicks - a.clicks);

  return { totalClicks, byProduct, byCard };
}

// ── LINK GENERATOR ────────────────────────────────────────────
function generateLink(body) {
  const domain  = body.domain || "";
  const card_id = body.card_id;
  if (!domain || !card_id) return { error: "domain and card_id required" };
  const link = `${domain.replace(/\/$/, "")}?card=${card_id}`;
  return { success: true, link };
}

// ── VERCEL TRACKING HELPERS ───────────────────────────────────
// Called by Vercel api/track.js — returns just what's needed for redirect + recording
function getProductForTrack(product_id) {
  const products = sheetToObjects(getSheet(SHEET_PRODUCTS));
  const p = products.find(p => p.product_id === product_id);
  if (!p) return { error: "Product not found" };
  return {
    product_id:    p.product_id,
    card_id:       p.card_id,
    affiliate_url: p.affiliate_url,
    regular_url:   p.regular_url,
  };
}

function recordClick(body) {
  try {
    const sh = getSheet(SHEET_CLICKS);
    sh.appendRow([
      uid(),
      body.product_id || "",
      body.card_id    || "",
      new Date().toISOString(),
      body.ua         || "",
    ]);
    return { success: true };
  } catch (err) {
    return { error: err.message };
  }
}

// ── META FETCHER ──────────────────────────────────────────────
function fetchMeta(url) {
  if (!url) return { error: "No URL" };
  try {
    const res  = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true });
    const html = res.getContentText();

    const ogTitle   = (html.match(/<meta[^>]+property=["']og:title["'][^>]+content=["']([^"']+)["']/i)   || [])[1] || "";
    const ogImage   = (html.match(/<meta[^>]+property=["']og:image["'][^>]+content=["']([^"']+)["']/i)   || [])[1] || "";
    const ogPrice   = (html.match(/<meta[^>]+property=["']product:price:amount["'][^>]+content=["']([^"']+)["']/i) || [])[1] || "";
    const titleTag  = (html.match(/<title[^>]*>([^<]+)<\/title>/i) || [])[1] || "";

    // Fallback image patterns for Indian platforms
    let image = ogImage;
    if (!image) {
      const imgMatch = html.match(/data-src=["'](https:\/\/[^"']+\.(?:jpg|jpeg|png|webp)[^"']*)/i)
                    || html.match(/src=["'](https:\/\/[^"']+\.(?:jpg|jpeg|png|webp)[^"']*)/i);
      if (imgMatch) image = imgMatch[1];
    }

    // Price fallback
    let price = ogPrice;
    if (!price) {
      const priceMatch = html.match(/(?:₹|Rs\.?|INR)\s*([\d,]+)/);
      if (priceMatch) price = "₹" + priceMatch[1];
    }

    return {
      name:     ogTitle || titleTag.split("|")[0].trim(),
      image:    image,
      price:    price,
      platform: detectPlatform(url)
    };
  } catch (err) {
    return { error: err.message };
  }
}

function detectPlatform(url) {
  if (!url) return "other";
  const u = url.toLowerCase();
  if (u.includes("ajio"))   return "ajio";
  if (u.includes("myntra")) return "myntra";
  if (u.includes("nykaa"))  return "nykaa";
  if (u.includes("meesho")) return "meesho";
  return "other";
}
