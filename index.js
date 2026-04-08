import express from "express";
import fetch from "node-fetch";
import { google } from "googleapis";

const app = express();
app.use(express.json({ limit: "10mb" }));

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
    },

    CAMPAIGN: {
      NAME: process.env.CAMPAIGN_SHEET_NAME || "ชีต4",
      GID: process.env.CAMPAIGN_SHEET_GID || "1706280606"
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
  spots: { data: null, fetchedAt: 0 },
  campaign: { data: null, fetchedAt: 0 }
};

const bookingStates = new Map();
// userId => { mode: "spot_booking", expiresAt }
const paymentTracking = new Map();
// userId => { displayName, qty, total, spotNumbers, expiresAt }

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

function setPaymentTracking(userId, payload) {
  if (!userId) return;
  paymentTracking.set(userId, {
    ...payload,
    expiresAt: Date.now() + CONFIG.SYSTEM.BOOKING_STATE_TTL_MS
  });
}

function getPaymentTracking(userId) {
  if (!userId) return null;
  const item = paymentTracking.get(userId);
  if (!item) return null;

  if (Date.now() > item.expiresAt) {
    paymentTracking.delete(userId);
    return null;
  }

  return item;
}

function clearPaymentTracking(userId) {
  if (!userId) return;
  paymentTracking.delete(userId);
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

  cache.prices = { data: table, fetchedAt: Date.now() };
  return table;
}

// =========================
// LOAD STOCK CSV
// =========================
async function loadStock(forceRefresh = false) {
  const age = Date.now() - cache.stock.fetchedAt;
  if (!forceRefresh && cache.stock.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
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

  cache.stock = { data: stock, fetchedAt: Date.now() };
  return stock;
}

// =========================
// LOAD SPOT SHEET (ชีต3)
// A = number
// B = Name
// C = ชำระ
// =========================
async function loadSpotSheet(forceRefresh = false) {
  const age = Date.now() - cache.spots.fetchedAt;
  if (!forceRefresh && cache.spots.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.spots.data;
  }

  const rows = await getSheetValues(`${CONFIG.SHEETS.SPOT.NAME}!A:C`);

  if (rows.length < 2) {
    throw new Error("Spot sheet has no data rows");
  }

  const spots = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const spotNumber = String(row[0] || "").trim();
    const name = String(row[1] || "").trim();
    const paymentStatus = String(row[2] || "").trim();

    if (!spotNumber) continue;

    spots.push({
      rowNumber: i + 1,
      spotNumber,
      sortNumber: Number(spotNumber),
      name,
      paymentStatus
    });
  }

  spots.sort((a, b) => {
    const aNum = Number.isNaN(a.sortNumber) ? Number.MAX_SAFE_INTEGER : a.sortNumber;
    const bNum = Number.isNaN(b.sortNumber) ? Number.MAX_SAFE_INTEGER : b.sortNumber;
    return aNum - bNum;
  });

  cache.spots = { data: spots, fetchedAt: Date.now() };
  return spots;
}

// =========================
// LOAD CAMPAIGN SHEET (ชีต4)
// A2 = หัวเรื่อง
// B2 = ราคา
// C2 = รูป (ควรเป็น public image URL)
// =========================
async function loadCampaign(forceRefresh = false) {
  const age = Date.now() - cache.campaign.fetchedAt;
  if (!forceRefresh && cache.campaign.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.campaign.data;
  }

  const rows = await getSheetValues(`${CONFIG.SHEETS.CAMPAIGN.NAME}!A2:C2`);
  const row = rows[0] || [];

  const campaign = {
    title: String(row[0] || "").trim(),
    price: Number(String(row[1] || "").replace(/,/g, "")) || 0,
    imageUrl: String(row[2] || "").trim()
  };

  cache.campaign = { data: campaign, fetchedAt: Date.now() };
  return campaign;
}

// =========================
// SPOT HELPERS
// =========================
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

function buildAllBookedSpotsText(spots) {
  const booked = spots.filter((spot) => spot.name);

  if (booked.length === 0) {
    return "ตอนนี้ยังไม่มีผู้จอง";
  }

  const lines = ["📌 รายชื่อผู้จองทั้งหมด"];

  for (const spot of booked) {
    const status = spot.paymentStatus ? ` (${spot.paymentStatus})` : "";
    lines.push(`สปอต ${spot.spotNumber}: ${spot.name}${status}`);
  }

  return lines.join("\n");
}

async function reserveSpots({ qty, displayName }) {
  return enqueueBooking(async () => {
    const spots = await loadSpotSheet(true);
    const availableSpots = getAvailableSpots(spots);

    if (availableSpots.length === 0) {
      return { ok: false, message: "สปอตเต็ม" };
    }

    if (availableSpots.length < qty) {
      return {
        ok: false,
        message: `สปอตคงเหลือไม่พอ ตอนนี้เหลือ ${availableSpots.length} สปอต`
      };
    }

    const selectedSpots = availableSpots.slice(0, qty);

    const updates = [];
    for (const spot of selectedSpots) {
      updates.push({
        range: `${CONFIG.SHEETS.SPOT.NAME}!B${spot.rowNumber}`,
        values: [[displayName]]
      });
      updates.push({
        range: `${CONFIG.SHEETS.SPOT.NAME}!C${spot.rowNumber}`,
        values: [["รอชำระ"]]
      });
    }

    await updateSheetValuesBatch(updates);
    const refreshed = await loadSpotSheet(true);

    return {
      ok: true,
      spots: selectedSpots.map((spot) => String(spot.spotNumber)),
      allBookedText: buildAllBookedSpotsText(refreshed)
    };
  });
}

async function markLatestBookingAsSlipSent(displayName) {
  const spots = await loadSpotSheet(true);

  const rowsToUpdate = spots
    .filter(
      (spot) =>
        spot.name === displayName &&
        (!spot.paymentStatus || spot.paymentStatus === "รอชำระ")
    )
    .map((spot) => spot.rowNumber);

  if (rowsToUpdate.length === 0) {
    return { ok: false, message: "ไม่พบรายการรอชำระของลูกค้า" };
  }

  const updates = rowsToUpdate.map((rowNumber) => ({
    range: `${CONFIG.SHEETS.SPOT.NAME}!C${rowNumber}`,
    values: [["ส่งสลิปแล้ว"]]
  }));

  await updateSheetValuesBatch(updates);
  return { ok: true, count: rowsToUpdate.length };
}

// =========================
// DISPLAY HELPERS
// =========================
function buildCampaignIntroText(campaign) {
  const lines = [];

  if (campaign.title) {
    lines.push(`🎯 รายการจองปัจจุบัน`);
    lines.push(campaign.title);
  } else {
    lines.push("🎯 ระบบจองสปอตสุ่ม");
  }

  if (campaign.price > 0) {
    lines.push(`💰 ราคา / สปอต: ${formatBaht(campaign.price)} บาท`);
  }

  return lines.join("\n");
}

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
  if (!userId) {
    throw new Error("Missing userId");
  }

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
  const [spots, campaign] = await Promise.all([
    loadSpotSheet(true),
    loadCampaign(true)
  ]);

  const availableCount = getAvailableSpots(spots).length;
  const bookedText = buildBookedSpotsText(spots);
  const introText = buildCampaignIntroText(campaign);

  const messages = [{ type: "text", text: introText }];

  if (campaign.imageUrl) {
    messages.push({
      type: "image",
      originalContentUrl: campaign.imageUrl,
      previewImageUrl: campaign.imageUrl
    });
  }

  messages.push({
    type: "text",
    text: `🎯 ตอนนี้มีสปอตว่าง ${availableCount} สปอต`
  });

  messages.push({
    type: "text",
    text: bookedText
  });

  if (availableCount <= 0) {
    clearBookingState(userId);
    messages.push({
      type: "text",
      text: "สปอตเต็ม"
    });
    await reply(replyToken, messages);
    return;
  }

  setBookingState(userId, "spot_booking");

  messages.push({
    type: "text",
    text: "หากต้องการจองสปอต ให้พิมพ์จำนวนสปอตที่ต้องการจอง เช่น 2"
  });

  await reply(replyToken, messages);
}

async function handleSpotBookingQuantity(replyToken, userId, qtyText) {
  const qty = Number(qtyText);

  if (!Number.isInteger(qty) || qty <= 0) {
    await reply(replyToken, [
      { type: "text", text: "กรุณาพิมพ์เป็นตัวเลขจำนวนสปอต เช่น 1 หรือ 2" }
    ]);
    return;
  }

  let displayName = "ลูกค้า";

  try {
    if (userId) {
      const profile = await getLineProfile(userId);
      displayName = profile.displayName || "ลูกค้า";
    }
  } catch (err) {
    logError("Get profile error:", err.message);
  }

  const campaign = await loadCampaign(true);
  const result = await reserveSpots({ qty, displayName });
  clearBookingState(userId);

  if (!result.ok) {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return;
  }

  const total = (campaign.price || 0) * qty;

  if (userId) {
    setPaymentTracking(userId, {
      displayName,
      qty,
      total,
      spotNumbers: result.spots
    });
  }

  await reply(replyToken, [
    {
      type: "text",
      text: `✅ จองสำเร็จ ${qty} สปอต
ชื่อ: ${displayName}
หมายเลขสปอต: ${result.spots.join(", ")}`
    },
    {
      type: "text",
      text: result.allBookedText
    },
    {
      type: "text",
      text: `🧾 สรุปยอดชำระ
${campaign.title || "รายการจองสปอต"}
ราคา / สปอต: ${formatBaht(campaign.price)} บาท
จำนวน: ${qty} สปอต
รวมทั้งหมด: ${formatBaht(total)} บาท`
    },
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

✨ ชำระแล้วโปรดส่ง Pay Slip การโอนในแชตนี้ ✨

เมื่อคุณส่งรูปสลิป ระบบจะบันทึกสถานะเป็น "ส่งสลิปแล้ว" เพื่อรอตรวจสอบ`
    }
  ]);
}

async function handleSlipImage(replyToken, userId) {
  const payment = getPaymentTracking(userId);

  if (!payment) {
    await reply(replyToken, [
      {
        type: "text",
        text: "ได้รับรูปแล้ว แต่ไม่พบรายการจองล่าสุดที่รอชำระ กรุณาติดต่อแอดมินเพื่อตรวจสอบ"
      }
    ]);
    return;
  }

  const result = await markLatestBookingAsSlipSent(payment.displayName);
  clearPaymentTracking(userId);

  if (!result.ok) {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return;
  }

  await reply(replyToken, [
    {
      type: "text",
      text: `📩 ได้รับรูปสลิปแล้ว
ชื่อ: ${payment.displayName}
จำนวนสปอต: ${payment.qty}
หมายเลขสปอต: ${payment.spotNumbers.join(", ")}

ระบบได้อัปเดตสถานะเป็น "ส่งสลิปแล้ว" เรียบร้อย กรุณารอแอดมินตรวจสอบ`
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
    const [prices, stock, spots, campaign] = await Promise.all([
      loadPrices(),
      loadStock(),
      loadSpotSheet(true),
      loadCampaign(true)
    ]);

    res.status(200).json({
      ok: true,
      priceItems: Object.keys(prices).length,
      stockModels: Object.keys(stock).length,
      totalSpots: spots.length,
      availableSpots: getAvailableSpots(spots).length,
      bookingEnabled: hasGoogleSheetsWriteConfig,
      campaign
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

    for (const event of events) {
      try {
        if (event.type !== "message") continue;
        if (!event.message) continue;

        const replyToken = event.replyToken;
        const userId = event.source?.userId || "";

        if (event.message.type === "image") {
          await handleSlipImage(replyToken, userId);
          continue;
        }

        if (event.message.type !== "text") continue;

        const text = String(event.message.text || "").trim();
        logInfo("Incoming text:", text, "userId:", userId || "(empty)");

        if (text === "ราคาสินค้า") {
          await handlePriceList(replyToken);
          continue;
        }

        if (/^check\s+rate/i.test(text)) {
          await handleCheckRate(replyToken, text);
          continue;
        }

        if (text === "จองสปอตสุ่ม") {
          await handleSpotBookingStart(replyToken, userId);
          continue;
        }

        const bookingState = getBookingState(userId);
        if (bookingState?.mode === "spot_booking" && isPositiveIntegerText(text)) {
          await handleSpotBookingQuantity(replyToken, userId, text);
          continue;
        }

        const handled = await handleOrder(replyToken, text);
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
});