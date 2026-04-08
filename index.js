import express from "express";
import fetch from "node-fetch";
import { google } from "googleapis";

const app = express();
app.use(express.json({ limit: "2mb" }));

// =========================
// CONFIG
// =========================
const CONFIG = {
  PORT: Number(process.env.PORT || 3000),

  LINE: {
    TOKEN: process.env.LINE_CHANNEL_ACCESS_TOKEN || ""
  },

  SHEETS: {
    SPREADSHEET_ID:
      process.env.SPREADSHEET_ID ||
      "1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc",

    PRICE: {
      NAME: process.env.PRICE_SHEET_NAME || "ชีต1",
      GID: process.env.PRICE_SHEET_GID || "0",
      CSV_URL:
        process.env.PRICE_CSV_URL ||
        "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0"
    },

    STOCK: {
      NAME: process.env.STOCK_SHEET_NAME || "ชีต2",
      GID: process.env.STOCK_SHEET_GID || "262793173",
      CSV_URL:
        process.env.STOCK_CSV_URL ||
        "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=262793173"
    },

    SPOT: {
      NAME: process.env.SPOT_SHEET_NAME || "ชีต3",
      GID: process.env.SPOT_SHEET_GID || "1276217956"
    }
  },

  GOOGLE: {
    SERVICE_ACCOUNT_EMAIL: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "",
    PRIVATE_KEY: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n")
  },

  PAYMENT: {
    IMAGE_URL:
      process.env.PAYMENT_IMAGE_URL ||
      "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg"
  },

  SYSTEM: {
    FETCH_TIMEOUT_MS: Number(process.env.FETCH_TIMEOUT_MS || 12000),
    CACHE_TTL_MS: Number(process.env.CACHE_TTL_MS || 60 * 1000),
    LINE_TEXT_LIMIT: Number(process.env.LINE_TEXT_LIMIT || 5000),
    BOOKING_STATE_TTL_MS: Number(process.env.BOOKING_STATE_TTL_MS || 10 * 60 * 1000)
  }
};

if (!CONFIG.LINE.TOKEN) {
  console.error("❌ Missing LINE_CHANNEL_ACCESS_TOKEN");
}

const hasGoogleSheetsWriteConfig =
  !!CONFIG.GOOGLE.SERVICE_ACCOUNT_EMAIL && !!CONFIG.GOOGLE.PRIVATE_KEY;

// =========================
// GOOGLE SHEETS AUTH
// =========================
const auth = hasGoogleSheetsWriteConfig
  ? new google.auth.JWT({
      email: CONFIG.GOOGLE.SERVICE_ACCOUNT_EMAIL,
      key: CONFIG.GOOGLE.PRIVATE_KEY,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    })
  : null;

const sheets = hasGoogleSheetsWriteConfig
  ? google.sheets({ version: "v4", auth })
  : null;

// =========================
// CACHE / STATE
// =========================
const cache = {
  prices: { data: null, fetchedAt: 0 },
  stock: { data: null, fetchedAt: 0 },
  spots: { data: null, fetchedAt: 0 }
};

const bookingStates = new Map();
let bookingQueue = Promise.resolve();

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

function withTimeout(promise, ms = CONFIG.SYSTEM.FETCH_TIMEOUT_MS) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Timeout after ${ms} ms`)), ms)
    )
  ]);
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

function getSpotGuideText(available) {
  return `🎯 ตอนนี้มีสปอตว่าง ${available} สปอต`;
}

function getSpotBookingPromptText() {
  return "หากต้องการจองสปอต ให้พิมพ์จำนวนสปอตที่ต้องการจอง เช่น 2";
}

function chunkText(text, limit = CONFIG.SYSTEM.LINE_TEXT_LIMIT) {
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

function qtyToBar(qty) {
  const n = Number(qty) || 0;
  return n > 0 ? "I ".repeat(n).trim() : "-";
}

function isPositiveIntegerText(text) {
  return /^[1-9]\d*$/.test(String(text || "").trim());
}

function setBookingState(userId, mode) {
  if (!userId) return;
  bookingStates.set(userId, {
    mode,
    expiresAt: Date.now() + CONFIG.SYSTEM.BOOKING_STATE_TTL_MS
  });
}

function getBookingState(userId) {
  if (!userId) return null;

  const state = bookingStates.get(userId);
  if (!state) return null;

  if (Date.now() > state.expiresAt) {
    bookingStates.delete(userId);
    return null;
  }

  return state;
}

function clearBookingState(userId) {
  if (!userId) return;
  bookingStates.delete(userId);
}

async function enqueueBooking(task) {
  const run = bookingQueue.then(task, task);
  bookingQueue = run.catch(() => {});
  return run;
}

// =========================
// FETCH HELPERS
// =========================
async function fetchText(url) {
  logInfo("Fetching URL:", url);

  const res = await withTimeout(fetch(url));
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
// GOOGLE SHEETS HELPERS
// =========================
function ensureSheetsEnabled() {
  if (!sheets) {
    throw new Error(
      "Google Sheets write system is not configured. Please set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY."
    );
  }
}

async function getSheetValues(range) {
  ensureSheetsEnabled();

  const res = await withTimeout(
    sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEETS.SPREADSHEET_ID,
      range
    })
  );

  return res.data.values || [];
}

async function updateSheetValuesBatch(data) {
  ensureSheetsEnabled();

  return withTimeout(
    sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: CONFIG.SHEETS.SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "RAW",
        data
      }
    })
  );
}

// =========================
// LOAD PRICE CSV
// =========================
async function loadPrices(forceRefresh = false) {
  const age = Date.now() - cache.prices.fetchedAt;
  if (!forceRefresh && cache.prices.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    logInfo("Using cached prices");
    return cache.prices.data;
  }

  const csv = await fetchText(CONFIG.SHEETS.PRICE.CSV_URL);
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
// LOAD STOCK CSV
// =========================
async function loadStock(forceRefresh = false) {
  const age = Date.now() - cache.stock.fetchedAt;
  if (!forceRefresh && cache.stock.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    logInfo("Using cached stock");
    return cache.stock.data;
  }

  const csv = await fetchText(CONFIG.SHEETS.STOCK.CSV_URL);
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
// LOAD SPOT SHEET
// A = number
// B = Name
// =========================
async function loadSpotSheet(forceRefresh = false) {
  const age = Date.now() - cache.spots.fetchedAt;
  if (!forceRefresh && cache.spots.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    logInfo("Using cached spots");
    return cache.spots.data;
  }

  const rows = await getSheetValues(`${CONFIG.SHEETS.SPOT.NAME}!A:B`);

  if (rows.length < 2) {
    throw new Error("Spot sheet has no data rows");
  }

  const spots = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const spotNumber = String(row[0] || "").trim();
    const name = String(row[1] || "").trim();

    if (!spotNumber) continue;

    spots.push({
      rowNumber: i + 1,
      spotNumber,
      sortNumber: Number(spotNumber),
      name
    });
  }

  spots.sort((a, b) => {
    const aNum = Number.isNaN(a.sortNumber) ? Number.MAX_SAFE_INTEGER : a.sortNumber;
    const bNum = Number.isNaN(b.sortNumber) ? Number.MAX_SAFE_INTEGER : b.sortNumber;
    return aNum - bNum;
  });

  cache.spots = {
    data: spots,
    fetchedAt: Date.now()
  };

  logInfo("Loaded spots:", spots.length, "rows");
  return spots;
}

function getAvailableSpots(spots) {
  return spots.filter((spot) => !spot.name);
}

function buildBookedSpotsText(spots) {
  const booked = spots.filter((spot) => spot.name);

  if (booked.length === 0) {
    return "ตอนนี้ยังไม่มีผู้จอง";
  }

  const lines = ["📌 รายชื่อผู้จองก่อนหน้า"];

  for (const spot of booked) {
    lines.push(`สปอต ${spot.spotNumber}: ${spot.name}`);
  }

  return lines.join("\n");
}

async function reserveSpots({ qty, displayName }) {
  return enqueueBooking(async () => {
    const spots = await loadSpotSheet(true);
    const availableSpots = getAvailableSpots(spots);

    if (availableSpots.length === 0) {
      return {
        ok: false,
        message: "สปอตเต็ม"
      };
    }

    if (availableSpots.length < qty) {
      return {
        ok: false,
        message: `สปอตคงเหลือไม่พอ ตอนนี้เหลือ ${availableSpots.length} สปอต`
      };
    }

    const selectedSpots = availableSpots.slice(0, qty);

    const updates = selectedSpots.map((spot) => ({
      range: `${CONFIG.SHEETS.SPOT.NAME}!B${spot.rowNumber}`,
      values: [[displayName]]
    }));

    await updateSheetValuesBatch(updates);
    await loadSpotSheet(true);

    return {
      ok: true,
      spots: selectedSpots.map((spot) => String(spot.spotNumber))
    };
  });
}

// =========================
// DISPLAY HELPERS
// =========================
function buildPriceList(table) {
  const codes = Object.keys(table).sort();
  const lines = ["📋 ราคาสินค้า"];

  for (const code of codes) {
    const item = table[code];
    const packText = item.pack == null ? "-" : `${formatBaht(item.pack)} บาท`;
    const boxText = item.box == null ? "-" : `${formatBaht(item.box)} บาท`;

    lines.push(`${displayCode(code)} | ซอง ${packText} | กล่อง ${boxText}`);
  }

  return lines.join("\n");
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
// =========================
function parseItems(text) {
  const regex =
    /([A-Z]+-?\d+)\s*(?:x?\s*)?(\d+)\s*(ซอง|ซ็อง|pack|box|กล่อง|บ็อก)/gi;

  const items = [];
  let match;

  while ((match = regex.exec(text))) {
    const unitWord = String(match[3] || "").toLowerCase();

    items.push({
      code: normalizeCode(match[1]),
      qty: Number(match[2]),
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

  if (items.length === 0) return null;

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
      lines.push(`${displayCode(item.code)} ❌ ไม่มีราคาประเภท${displayUnit(item.unit)}`);
      continue;
    }

    const sum = unitPrice * item.qty;
    total += sum;
    valid += 1;

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
// LINE HELPERS
// =========================
async function getLineProfile(userId) {
  const res = await withTimeout(
    fetch(`https://api.line.me/v2/bot/profile/${userId}`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${CONFIG.LINE.TOKEN}`
      }
    })
  );

  const body = await res.text();

  if (!res.ok) {
    throw new Error(`Get profile failed ${res.status}: ${body}`);
  }

  return JSON.parse(body);
}

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
        Authorization: `Bearer ${CONFIG.LINE.TOKEN}`
      },
      body: JSON.stringify({
        replyToken,
        messages: safeMessages
      })
    })
  );

  const body = await res.text();

  if (!res.ok) {
    throw new Error(`LINE reply failed ${res.status}: ${body}`);
  }

  logInfo("LINE reply success:", body);
}

// =========================
// HANDLERS
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

  await reply(replyToken, [{ type: "text", text: formatStock(model, stock[model]) }]);
}

async function handleSpotBookingStart(replyToken, userId) {
  const spots = await loadSpotSheet(true);
  const availableCount = getAvailableSpots(spots).length;
  const bookedText = buildBookedSpotsText(spots);

  if (availableCount <= 0) {
    clearBookingState(userId);
    await reply(replyToken, [
      { type: "text", text: bookedText },
      { type: "text", text: "สปอตเต็ม" }
    ]);
    return;
  }

  setBookingState(userId, "spot_booking");

  await reply(replyToken, [
    { type: "text", text: getSpotGuideText(availableCount) },
    { type: "text", text: bookedText },
    { type: "text", text: getSpotBookingPromptText() }
  ]);
}

async function handleSpotBookingQuantity(replyToken, userId, qtyText) {
  const qty = Number(qtyText);

  if (!Number.isInteger(qty) || qty <= 0) {
    await reply(replyToken, [
      { type: "text", text: "กรุณาพิมพ์เป็นตัวเลขจำนวนสปอต เช่น 1 หรือ 2" }
    ]);
    return;
  }

  const profile = await getLineProfile(userId);
  const displayName = profile.displayName || "ไม่ทราบชื่อ";

  const result = await reserveSpots({ qty, displayName });
  clearBookingState(userId);

  if (!result.ok) {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return;
  }

  await reply(replyToken, [
    {
      type: "text",
      text: `✅ จองสำเร็จ ${qty} สปอต
ชื่อ: ${displayName}
หมายเลขสปอต: ${result.spots.join(", ")}`
    }
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
        originalContentUrl: CONFIG.PAYMENT.IMAGE_URL,
        previewImageUrl: CONFIG.PAYMENT.IMAGE_URL
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
    const [prices, stock] = await Promise.all([loadPrices(), loadStock()]);

    const payload = {
      ok: true,
      priceItems: Object.keys(prices).length,
      stockModels: Object.keys(stock).length,
      cache: {
        pricesAgeMs: Date.now() - cache.prices.fetchedAt,
        stockAgeMs: Date.now() - cache.stock.fetchedAt
      },
      bookingEnabled: hasGoogleSheetsWriteConfig,
      sheets: {
        priceGid: CONFIG.SHEETS.PRICE.GID,
        stockGid: CONFIG.SHEETS.STOCK.GID,
        spotGid: CONFIG.SHEETS.SPOT.GID
      }
    };

    if (hasGoogleSheetsWriteConfig) {
      const spots = await loadSpotSheet(true);
      payload.totalSpots = spots.length;
      payload.availableSpots = getAvailableSpots(spots).length;
      payload.cache.spotsAgeMs = Date.now() - cache.spots.fetchedAt;
    }

    res.status(200).json(payload);
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

    for (const event of events) {
      try {
        if (event.type !== "message") continue;
        if (!event.message || event.message.type !== "text") continue;

        const text = String(event.message.text || "").trim();
        const userId = event.source?.userId || "";

        logInfo("Incoming text:", text, "userId:", userId);

        if (text === "ราคาสินค้า") {
          await handlePriceList(event.replyToken);
          continue;
        }

        if (/^check\s+rate/i.test(text)) {
          await handleCheckRate(event.replyToken, text);
          continue;
        }

        if (text === "จองสปอตสุ่ม") {
          await handleSpotBookingStart(event.replyToken, userId);
          continue;
        }

        const bookingState = getBookingState(userId);
        if (bookingState?.mode === "spot_booking" && isPositiveIntegerText(text)) {
          await handleSpotBookingQuantity(event.replyToken, userId, text);
          continue;
        }

        const handled = await handleOrder(event.replyToken, text);
        if (handled) continue;
      } catch (eventErr) {
        logError("Event handling error:", eventErr);

        try {
          await reply(event.replyToken, [
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
app.listen(CONFIG.PORT, () => {
  logInfo(`✅ Server running on port ${CONFIG.PORT}`);
  logInfo("PRICE_CSV_URL =", CONFIG.SHEETS.PRICE.CSV_URL);
  logInfo("STOCK_CSV_URL =", CONFIG.SHEETS.STOCK.CSV_URL);
  logInfo("SPOT_SHEET_NAME =", CONFIG.SHEETS.SPOT.NAME);
  logInfo("SPOT_SHEET_GID =", CONFIG.SHEETS.SPOT.GID);
  logInfo("BOOKING_ENABLED =", hasGoogleSheetsWriteConfig);
});