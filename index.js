import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json({ limit: "2mb" }));

// =========================
// ENV / CONFIG
// =========================
const PORT = process.env.PORT || 3000;
const TOKEN = process.env.LINE_CHANNEL_ACCESS_TOKEN || "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";

if (!TOKEN) {
  console.error("❌ Missing LINE_CHANNEL_ACCESS_TOKEN in environment variables");
}

const PRICE_CSV_URL =
  process.env.PRICE_CSV_URL ||
  "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";

const STOCK_CSV_URL =
  process.env.STOCK_CSV_URL ||
  "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=262793173";

const PAYMENT_IMAGE_URL =
  process.env.PAYMENT_IMAGE_URL ||
  "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg";

const FETCH_TIMEOUT_MS = Number(process.env.FETCH_TIMEOUT_MS || 12000);
const CACHE_TTL_MS = Number(process.env.CACHE_TTL_MS || 60 * 1000);
const LINE_TEXT_LIMIT = 5000;

// =========================
// SIMPLE MEMORY CACHE
// =========================
const cache = {
  prices: { data: null, fetchedAt: 0 },
  stock: { data: null, fetchedAt: 0 }
};

// =========================
// UTILITIES
// =========================
function nowISO() {
  return new Date().toISOString();
}

function logInfo(...args) {
  console.log(`[${nowISO()}]`, ...args);
}

function logError(...args) {
  console.error(`[${nowISO()}]`, ...args);
}

function formatBaht(num) {
  const n = Number(num);
  if (Number.isNaN(n)) return "-";
  return n.toLocaleString("en-US");
}

function normalizeCode(code) {
  return String(code || "")
    .toUpperCase()
    .replace(/[^A-Z0-9]/g, "");
}

function displayCode(code) {
  const clean = normalizeCode(code);
  const m = clean.match(/^([A-Z]+)(\d+)$/);
  return m ? `${m[1]}-${m[2]}` : clean;
}

function displayUnit(unit) {
  return unit === "pack" ? "ซอง" : "กล่อง";
}

function getOrderGuideText() {
  return `หากต้องการสรุปราคา พิมพ์เช่น
OP13 2 ซอง OP15 2 ซอง
หรือ
OP13 1 กล่อง`;
}

function chunkText(text, limit = LINE_TEXT_LIMIT) {
  if (!text || text.length <= limit) return [text];

  const chunks = [];
  let remaining = text;

  while (remaining.length > limit) {
    let cut = remaining.lastIndexOf("\n", limit);
    if (cut <= 0) cut = limit;
    chunks.push(remaining.slice(0, cut));
    remaining = remaining.slice(cut).trimStart();
  }

  if (remaining) chunks.push(remaining);
  return chunks;
}

function withTimeout(promise, ms) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Timeout after ${ms} ms`)), ms)
    )
  ]);
}

// CSV parser รองรับ quote เบื้องต้น
function parseCSVLine(line) {
  const result = [];
  let current = "";
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    const next = line[i + 1];

    if (ch === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i++;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === "," && !inQuotes) {
      result.push(current);
      current = "";
      continue;
    }

    current += ch;
  }

  result.push(current);
  return result.map((x) => x.trim());
}

function parseCSV(csvText) {
  return csvText
    .replace(/\r/g, "")
    .split("\n")
    .filter((line) => line.trim() !== "")
    .map(parseCSVLine);
}

async function fetchText(url) {
  logInfo("Fetching URL:", url);

  const res = await withTimeout(fetch(url), FETCH_TIMEOUT_MS);
  if (!res.ok) {
    throw new Error(`Fetch failed ${res.status} ${res.statusText}`);
  }

  const text = await res.text();
  if (!text || !text.trim()) {
    throw new Error("Fetched empty response");
  }

  return text;
}

// =========================
// LOAD PRICES
// Sheet1:
// A = Code
// B = Pack_price
// C = Box_price
// =========================
async function loadPrices(forceRefresh = false) {
  const age = Date.now() - cache.prices.fetchedAt;
  if (!forceRefresh && cache.prices.data && age < CACHE_TTL_MS) {
    logInfo("Using cached prices");
    return cache.prices.data;
  }

  const csv = await fetchText(PRICE_CSV_URL);
  const rows = parseCSV(csv);

  if (rows.length < 2) {
    throw new Error("Price sheet has no data rows");
  }

  const table = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const code = normalizeCode(row[0] || "");
    const pack = Number(String(row[1] || "").replace(/,/g, ""));
    const box = Number(String(row[2] || "").replace(/,/g, ""));

    if (!code) continue;

    table[code] = {
      pack: Number.isNaN(pack) ? null : pack,
      box: Number.isNaN(box) ? null : box
    };
  }

  if (Object.keys(table).length === 0) {
    throw new Error("Price table parsed but empty");
  }

  cache.prices = {
    data: table,
    fetchedAt: Date.now()
  };

  logInfo("Loaded prices:", Object.keys(table).length, "items");
  return table;
}

// =========================
// LOAD STOCK
// Sheet2:
// A = รุ่น
// B = หมวด
// C = เหลือ
// =========================
async function loadStock(forceRefresh = false) {
  const age = Date.now() - cache.stock.fetchedAt;
  if (!forceRefresh && cache.stock.data && age < CACHE_TTL_MS) {
    logInfo("Using cached stock");
    return cache.stock.data;
  }

  const csv = await fetchText(STOCK_CSV_URL);
  const rows = parseCSV(csv);

  if (rows.length < 2) {
    throw new Error("Stock sheet has no data rows");
  }

  const stock = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const model = normalizeCode(row[0] || "");
    const category = String(row[1] || "").trim().toUpperCase();
    const qty = Number(String(row[2] || "").replace(/,/g, ""));

    if (!model || !category) continue;

    if (!stock[model]) stock[model] = {};
    stock[model][category] = Number.isNaN(qty) ? 0 : qty;
  }

  if (Object.keys(stock).length === 0) {
    throw new Error("Stock table parsed but empty");
  }

  cache.stock = {
    data: stock,
    fetchedAt: Date.now()
  };

  logInfo("Loaded stock:", Object.keys(stock).length, "models");
  return stock;
}

// =========================
// BUILD PRICE LIST
// =========================
function buildPriceList(table) {
  const codes = Object.keys(table).sort();
  const lines = ["📋 ราคาสินค้า"];

  for (const code of codes) {
    const item = table[code];
    const packText = item.pack == null ? "-" : `${formatBaht(item.pack)} บาท`;
    const boxText = item.box == null ? "-" : `${formatBaht(item.box)} บาท`;

    lines.push(
      `${displayCode(code)} | ซอง ${packText} | กล่อง ${boxText}`
    );
  }

  return lines.join("\n");
}

// =========================
// STOCK FORMAT
// =========================
function qtyToBar(qty) {
  const n = Number(qty) || 0;
  return n > 0 ? "I ".repeat(n).trim() : "-";
}

function formatStock(model, data) {
  const lines = [`📦 ${displayCode(model)} เหลือใน carton`];

  for (const key of ["SP", "SEC", "LPA", "DON"]) {
    if (data[key] !== undefined) {
      lines.push(`${key}: ${qtyToBar(data[key])}`);
    }
  }

  return lines.join("\n");
}

function formatAllStock(stock) {
  const models = Object.keys(stock).sort();
  const lines = ["📦 สถานะสินค้าใน carton"];

  for (const model of models) {
    lines.push("");
    lines.push(displayCode(model));

    for (const key of ["SP", "SEC", "LPA", "DON"]) {
      if (stock[model][key] !== undefined) {
        lines.push(`${key}: ${qtyToBar(stock[model][key])}`);
      }
    }
  }

  return lines.join("\n");
}

// =========================
// PARSE ORDER ITEMS
// รองรับ:
// OP13 2 ซอง
// OP-13 2 ซอง
// op13 x2 ซอง
// prb01 1 box
// =========================
function parseItems(text) {
  const regex =
    /([A-Z]+-?\d+)\s*(?:x?\s*)?(\d+)\s*(ซอง|ซ็อง|pack|box|กล่อง|บ็อก)/gi;

  const items = [];
  let m;

  while ((m = regex.exec(text))) {
    const unitWord = String(m[3] || "").toLowerCase();
    items.push({
      code: normalizeCode(m[1]),
      qty: Number(m[2]),
      unit:
        unitWord.includes("ซอง") || unitWord === "pack" || unitWord === "ซ็อง"
          ? "pack"
          : "box"
    });
  }

  return items;
}

// =========================
// CALCULATE ORDER
// =========================
async function calculate(text) {
  let clean = String(text || "").trim();

  if (!clean) return null;

  if (clean === "รวมราคา") {
    return {
      status: "guide",
      message: getOrderGuideText()
    };
  }

  clean = clean.replace(/^รวมราคา/i, "").trim();

  const items = parseItems(clean);
  logInfo("PARSED ITEMS:", items);

  if (items.length === 0) {
    return null;
  }

  const table = await loadPrices();

  let total = 0;
  let valid = 0;
  const lines = ["🧾 สรุปรายการ"];

  for (const item of items) {
    const priceData = table[item.code];

    if (!priceData) {
      lines.push(`${displayCode(item.code)} ❌ ไม่มีสินค้า`);
      continue;
    }

    const unitPrice = priceData[item.unit];

    if (unitPrice == null) {
      lines.push(
        `${displayCode(item.code)} ❌ ไม่มีราคาประเภท${displayUnit(item.unit)}`
      );
      continue;
    }

    const sum = unitPrice * item.qty;
    total += sum;
    valid++;

    lines.push(
      `${displayCode(item.code)} ${displayUnit(item.unit)}ละ ${formatBaht(unitPrice)} x${item.qty} = ${formatBaht(sum)} บาท`
    );
  }

  if (valid === 0) {
    return {
      status: "invalid",
      message: "พิมพ์เช่น OP13 2 ซอง OP15 1 box"
    };
  }

  lines.push("━━━━━━━━━━");
  lines.push(`รวมทั้งหมด = ${formatBaht(total)} บาท`);

  return {
    status: "success",
    summary: lines.join("\n"),
    payment: `📌 กรุณาโอน ${formatBaht(total)} บาท`
  };
}

// =========================
// LINE REPLY
// =========================
async function reply(replyToken, messages) {
  const safeMessages = messages
    .filter(Boolean)
    .flatMap((msg) => {
      if (msg.type === "text") {
        return chunkText(msg.text).map((chunk) => ({
          type: "text",
          text: chunk
        }));
      }
      return [msg];
    })
    .slice(0, 5);

  logInfo("Replying with", safeMessages.length, "message(s)");

  const res = await withTimeout(
    fetch("https://api.line.me/v2/bot/message/reply", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${TOKEN}`
      },
      body: JSON.stringify({
        replyToken,
        messages: safeMessages
      })
    }),
    FETCH_TIMEOUT_MS
  );

  const body = await res.text();

  if (!res.ok) {
    throw new Error(`LINE reply failed ${res.status}: ${body}`);
  }

  logInfo("LINE reply success:", body);
}

// =========================
// MESSAGE HANDLERS
// =========================
async function handlePriceList(replyToken) {
  const table = await loadPrices();
  await reply(replyToken, [
    { type: "text", text: buildPriceList(table) },
    { type: "text", text: getOrderGuideText() }
  ]);
}

async function handleCheckRate(replyToken, text) {
  const stock = await loadStock();
  const parts = text.trim().split(/\s+/);

  // check rate
  // check rate op13
  const modelRaw = parts.slice(2).join("");
  if (!modelRaw) {
    await reply(replyToken, [{ type: "text", text: formatAllStock(stock) }]);
    return;
  }

  const model = normalizeCode(modelRaw);

  if (!stock[model]) {
    await reply(replyToken, [{ type: "text", text: "ไม่พบข้อมูลรุ่นนี้" }]);
    return;
  }

  await reply(replyToken, [
    { type: "text", text: formatStock(model, stock[model]) }
  ]);
}

async function handleOrder(replyToken, text) {
  const result = await calculate(text);

  if (!result) return false;

  if (result.status === "guide") {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return true;
  }

  if (result.status === "invalid") {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return true;
  }

  if (result.status === "success") {
    await reply(replyToken, [
      { type: "text", text: result.summary },
      { type: "text", text: result.payment },
      {
        type: "image",
        originalContentUrl: PAYMENT_IMAGE_URL,
        previewImageUrl: PAYMENT_IMAGE_URL
      },
      {
        type: "text",
        text: `สามารถชำระเงินผ่านช่องทางอื่น ๆ ได้ดังนี้

ชื่อบัญชี ปรัชญา สุดใจดี

K-Bank 0503228092

True Wallet 0982652650

✨ ชำระแล้วโปรดแปะ Pay Slip การโอนทุกครั้ง ✨

ขอบคุณนะครับ`
      }
    ]);
    return true;
  }

  return false;
}

// =========================
// ROUTES
// =========================
app.get("/", (req, res) => {
  res.status(200).json({
    ok: true,
    service: "line-webhook",
    time: nowISO()
  });
});

app.get("/health", async (req, res) => {
  try {
    const [prices, stock] = await Promise.all([
      loadPrices(),
      loadStock()
    ]);

    res.status(200).json({
      ok: true,
      priceItems: Object.keys(prices).length,
      stockModels: Object.keys(stock).length,
      cache: {
        pricesAgeMs: Date.now() - cache.prices.fetchedAt,
        stockAgeMs: Date.now() - cache.stock.fetchedAt
      }
    });
  } catch (err) {
    logError("Health check error:", err);
    res.status(500).json({
      ok: false,
      error: err.message
    });
  }
});

app.post("/webhook", async (req, res) => {
  try {
    const events = req.body?.events || [];
    logInfo("Webhook hit. events =", events.length);

    for (const e of events) {
      try {
        if (e.type !== "message") continue;
        if (!e.message || e.message.type !== "text") continue;

        const text = String(e.message.text || "").trim();
        logInfo("Incoming text:", text);

        if (text === "ราคาสินค้า") {
          await handlePriceList(e.replyToken);
          continue;
        }

        if (/^check\s+rate/i.test(text)) {
          await handleCheckRate(e.replyToken, text);
          continue;
        }

        const handled = await handleOrder(e.replyToken, text);
        if (handled) continue;

        // ไม่ตอบข้อความอื่น เพื่อไม่รบกวนแชต
      } catch (eventErr) {
        logError("Event handling error:", eventErr);

        try {
          await reply(e.replyToken, [
            {
              type: "text",
              text: "ระบบเกิดข้อผิดพลาดชั่วคราว กรุณาลองใหม่อีกครั้ง"
            }
          ]);
        } catch (replyErr) {
          logError("Fallback reply failed:", replyErr);
        }
      }
    }

    res.sendStatus(200);
  } catch (err) {
    logError("Webhook error:", err);
    res.sendStatus(500);
  }
});

// =========================
// START
// =========================
app.listen(PORT, () => {
  logInfo(`✅ Server running on port ${PORT}`);
  logInfo("PRICE_CSV_URL =", PRICE_CSV_URL);
  logInfo("STOCK_CSV_URL =", STOCK_CSV_URL);
});