import express from "express";
import fetch from "node-fetch";
import { google } from "googleapis";
import { Readable } from "stream";

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
    },

    LOG: {
      NAME: process.env.LOG_SHEET_NAME || "ชีต5",
      GID: process.env.LOG_SHEET_GID || "2113781022"
    },

    CUSTOMER: {
      NAME: process.env.CUSTOMER_SHEET_NAME || "ชีต6",
      GID: process.env.CUSTOMER_SHEET_GID || "1948361968"
    }
  },

  GOOGLE: {
    SERVICE_ACCOUNT_EMAIL: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "",
    PRIVATE_KEY: (process.env.GOOGLE_PRIVATE_KEY || "").replace(/\\n/g, "\n"),
    DRIVE_FOLDER_ID: process.env.GOOGLE_DRIVE_FOLDER_ID || ""
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

const hasGoogleConfig =
  !!CONFIG.GOOGLE.SERVICE_ACCOUNT_EMAIL && !!CONFIG.GOOGLE.PRIVATE_KEY;

const auth = hasGoogleConfig
  ? new google.auth.JWT({
      email: CONFIG.GOOGLE.SERVICE_ACCOUNT_EMAIL,
      key: CONFIG.GOOGLE.PRIVATE_KEY,
      scopes: [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
      ]
    })
  : null;

const sheets = hasGoogleConfig
  ? google.sheets({ version: "v4", auth })
  : null;

const drive = hasGoogleConfig
  ? google.drive({ version: "v3", auth })
  : null;

// =========================
// CACHE / STATE
// =========================
const cache = {
  prices: { data: null, fetchedAt: 0 },
  stock: { data: null, fetchedAt: 0 },
  spots: { data: null, fetchedAt: 0 },
  campaign: { data: null, fetchedAt: 0 },
  customers: { data: null, fetchedAt: 0 }
};

const bookingStates = new Map();
const paymentTracking = new Map();
let bookingQueue = Promise.resolve();

// =========================
// UTILITIES
// =========================
function nowISO() {
  return new Date().toISOString();
}

function nowLocalText() {
  return new Date().toLocaleString("sv-SE", { timeZone: "Asia/Bangkok" }).replace(" ", " ");
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

function padCustomerNo(num) {
  return String(num).padStart(4, "0");
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
// GOOGLE HELPERS
// =========================
function ensureGoogleEnabled() {
  if (!sheets || !drive) {
    throw new Error(
      "Google system is not configured. Please set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY."
    );
  }
}

async function getSheetValues(range) {
  ensureGoogleEnabled();

  const res = await withTimeout(
    sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SHEETS.SPREADSHEET_ID,
      range
    })
  );

  return res.data.values || [];
}

async function updateSheetValuesBatch(data) {
  ensureGoogleEnabled();

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

async function appendSheetRow(range, values) {
  ensureGoogleEnabled();

  return withTimeout(
    sheets.spreadsheets.values.append({
      spreadsheetId: CONFIG.SHEETS.SPREADSHEET_ID,
      range,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [values]
      }
    })
  );
}

// =========================
// LINE HELPERS
// =========================
async function getLineProfile(userId) {
  if (!userId) throw new Error("Missing userId");

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

async function downloadLineImageBuffer(messageId) {
  const res = await withTimeout(
    fetch(`https://api-data.line.me/v2/bot/message/${messageId}/content`, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${CONFIG.LINE.TOKEN}`
      }
    })
  );

  if (!res.ok) {
    const body = await res.text();
    throw new Error(`Download LINE image failed ${res.status}: ${body}`);
  }

  const arrayBuffer = await res.arrayBuffer();
  return Buffer.from(arrayBuffer);
}

async function uploadSlipToDrive(buffer, fileName) {
  ensureGoogleEnabled();

  if (!CONFIG.GOOGLE.DRIVE_FOLDER_ID) {
    throw new Error("Missing GOOGLE_DRIVE_FOLDER_ID");
  }

  const stream = Readable.from(buffer);

  const createRes = await withTimeout(
    drive.files.create({
      requestBody: {
        name: fileName,
        parents: [CONFIG.GOOGLE.DRIVE_FOLDER_ID]
      },
      media: {
        mimeType: "image/jpeg",
        body: stream
      },
      fields: "id, webViewLink, webContentLink",
      supportsAllDrives: true
    })
  );

  const fileId = createRes.data.id;

  await withTimeout(
    drive.permissions.create({
      fileId,
      requestBody: {
        role: "reader",
        type: "anyone"
      },
      supportsAllDrives: true
    })
  );

  const file = await withTimeout(
    drive.files.get({
      fileId,
      fields: "id, webViewLink, webContentLink",
      supportsAllDrives: true
    })
  );

  return {
    fileId: file.data.id,
    url: file.data.webViewLink || file.data.webContentLink || ""
  };
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
// LOADERS
// =========================
async function loadPrices(forceRefresh = false) {
  const age = Date.now() - cache.prices.fetchedAt;
  if (!forceRefresh && cache.prices.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.prices.data;
  }

  const csv = await fetchText(CONFIG.SHEETS.PRICE.CSV_URL);
  const rows = parseCSV(csv);
  if (rows.length < 2) throw new Error("Price sheet has no data rows");

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

async function loadStock(forceRefresh = false) {
  const age = Date.now() - cache.stock.fetchedAt;
  if (!forceRefresh && cache.stock.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.stock.data;
  }

  const csv = await fetchText(CONFIG.SHEETS.STOCK.CSV_URL);
  const rows = parseCSV(csv);
  if (rows.length < 2) throw new Error("Stock sheet has no data rows");

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

async function loadSpotSheet(forceRefresh = false) {
  const age = Date.now() - cache.spots.fetchedAt;
  if (!forceRefresh && cache.spots.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.spots.data;
  }

  const rows = await getSheetValues(`${CONFIG.SHEETS.SPOT.NAME}!A:C`);
  if (rows.length < 2) throw new Error("Spot sheet has no data rows");

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

async function loadCustomers(forceRefresh = false) {
  const age = Date.now() - cache.customers.fetchedAt;
  if (!forceRefresh && cache.customers.data && age < CONFIG.SYSTEM.CACHE_TTL_MS) {
    return cache.customers.data;
  }

  const rows = await getSheetValues(`${CONFIG.SHEETS.CUSTOMER.NAME}!A:D`);
  const customers = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const customerNo = String(row[0] || "").trim();
    const userId = String(row[1] || "").trim();
    const latestName = String(row[2] || "").trim();
    const createdAt = String(row[3] || "").trim();

    if (!customerNo && !userId) continue;

    customers.push({
      rowNumber: i + 1,
      customerNo,
      userId,
      latestName,
      createdAt
    });
  }

  cache.customers = { data: customers, fetchedAt: Date.now() };
  return customers;
}

// =========================
// CUSTOMER MASTER
// =========================
async function getOrCreateCustomer(userId, displayName) {
  const customers = await loadCustomers(true);

  if (userId) {
    const found = customers.find((c) => c.userId === userId);
    if (found) {
      if (displayName && found.latestName !== displayName) {
        await updateSheetValuesBatch([
          {
            range: `${CONFIG.SHEETS.CUSTOMER.NAME}!C${found.rowNumber}`,
            values: [[displayName]]
          }
        ]);
        await loadCustomers(true);
      }

      return {
        customerNo: found.customerNo,
        userId: found.userId,
        displayName: displayName || found.latestName || "ลูกค้า"
      };
    }
  }

  const maxNo = customers.reduce((max, c) => {
    const num = Number(c.customerNo);
    return Number.isNaN(num) ? max : Math.max(max, num);
  }, 0);

  const nextNo = padCustomerNo(maxNo + 1);
  const nowText = nowLocalText();

  await appendSheetRow(`${CONFIG.SHEETS.CUSTOMER.NAME}!A:D`, [
    nextNo,
    userId || "",
    displayName || "ลูกค้า",
    nowText
  ]);

  await loadCustomers(true);

  return {
    customerNo: nextNo,
    userId: userId || "",
    displayName: displayName || "ลูกค้า"
  };
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
    return { ok: false, message: "not_found" };
  }

  const updates = rowsToUpdate.map((rowNumber) => ({
    range: `${CONFIG.SHEETS.SPOT.NAME}!C${rowNumber}`,
    values: [["ส่งสลิปแล้ว"]]
  }));

  await updateSheetValuesBatch(updates);
  return { ok: true, count: rowsToUpdate.length };
}

// =========================
// LOG HELPERS
// =========================
async function appendOrderLog({
  type,
  customerNo,
  userId,
  displayName,
  itemName,
  detail,
  qty,
  unitPrice,
  total,
  spotNumbers,
  slipStatus,
  slipUrl,
  note
}) {
  await appendSheetRow(`${CONFIG.SHEETS.LOG.NAME}!A:N`, [
    nowLocalText(),
    type || "",
    customerNo || "",
    userId || "",
    displayName || "",
    itemName || "",
    detail || "",
    qty != null ? String(qty) : "",
    unitPrice != null ? String(unitPrice) : "",
    total != null ? String(total) : "",
    spotNumbers || "",
    slipStatus || "",
    slipUrl || "",
    note || ""
  ]);
}

// =========================
// DISPLAY HELPERS
// =========================
function buildCampaignIntroText(campaign) {
  const lines = [];

  if (campaign.title) {
    lines.push("🎯 รายการจองปัจจุบัน");
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
  const detailLines = [];

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

    const detailText =
      `${displayCode(item.code)} ${displayUnit(item.unit)}ละ ${formatBaht(unitPrice)} x${item.qty} = ${formatBaht(sum)} บาท`;

    lines.push(detailText);
    detailLines.push({
      code: displayCode(item.code),
      qty: item.qty,
      unit: displayUnit(item.unit),
      unitPrice,
      sum,
      text: detailText
    });
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
    payment: `📌 กรุณาโอน ${formatBaht(total)} บาท`,
    total,
    detailLines
  };
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
    messages.push({ type: "text", text: "สปอตเต็ม" });
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

  const customer = await getOrCreateCustomer(userId, displayName);
  displayName = customer.displayName;

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
      type: "spot_booking",
      customerNo: customer.customerNo,
      userId,
      displayName,
      itemName: campaign.title || "รายการจองสปอต",
      detail: "จองสปอต",
      qty,
      unitPrice: campaign.price || 0,
      total,
      spotNumbers: result.spots.join(", ")
    });
  }

  await reply(replyToken, [
    {
      type: "text",
      text: `✅ จองสำเร็จ ${qty} สปอต
ชื่อ: ${displayName}
รหัสลูกค้า: ${customer.customerNo}
หมายเลขสปอต: ${result.spots.join(", ")}`
    },
    {
      type: "text",
      text: result.allBookedText
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
    },
    {
      type: "text",
      text: `🧾 สรุปยอดชำระ
${campaign.title || "รายการจองสปอต"}
ราคา / สปอต: ${formatBaht(campaign.price)} บาท
จำนวน: ${qty} สปอต
รวมทั้งหมด: ${formatBaht(total)} บาท`
    }
  ]);
}

async function handleSlipImage(replyToken, userId, messageId) {
  let displayName = "ลูกค้า";

  try {
    if (userId) {
      const profile = await getLineProfile(userId);
      displayName = profile.displayName || "ลูกค้า";
    }
  } catch (err) {
    logError("Get profile error:", err.message);
  }

  const customer = await getOrCreateCustomer(userId, displayName);
  const payment = getPaymentTracking(userId);

  let slipUrl = "";
  try {
    if (messageId && CONFIG.GOOGLE.DRIVE_FOLDER_ID) {
      const buffer = await downloadLineImageBuffer(messageId);
      const uploaded = await uploadSlipToDrive(
        buffer,
        `slip_${customer.customerNo}_${Date.now()}.jpg`
      );
      slipUrl = uploaded.url || "";
    }
  } catch (err) {
    logError("Upload slip error:", err.message);
  }

  if (!payment) {
    await appendOrderLog({
      type: "normal_order",
      customerNo: customer.customerNo,
      userId: customer.userId,
      displayName: customer.displayName,
      itemName: "",
      detail: "ลูกค้าส่งสลิป แต่ไม่พบรายการ tracking ล่าสุด",
      qty: "",
      unitPrice: "",
      total: "",
      spotNumbers: "",
      slipStatus: "ส่งสลิปแล้ว",
      slipUrl,
      note: "รอแอดมินตรวจสอบ"
    });

    await reply(replyToken, [
      {
        type: "text",
        text: "ได้รับรูปสลิปแล้ว กรุณารอแอดมินตรวจสอบ"
      }
    ]);
    return;
  }

  if (payment.type === "spot_booking") {
    const markResult = await markLatestBookingAsSlipSent(payment.displayName);
    if (!markResult.ok) {
      logError("Mark booking slip status failed:", markResult.message);
    }
  }

  await appendOrderLog({
    type: payment.type,
    customerNo: payment.customerNo || customer.customerNo,
    userId: payment.userId || customer.userId,
    displayName: payment.displayName || customer.displayName,
    itemName: payment.itemName || "",
    detail: payment.detail || "",
    qty: payment.qty != null ? payment.qty : "",
    unitPrice: payment.unitPrice != null ? payment.unitPrice : "",
    total: payment.total != null ? payment.total : "",
    spotNumbers: payment.spotNumbers || "",
    slipStatus: "ส่งสลิปแล้ว",
    slipUrl,
    note: "รอแอดมินตรวจสอบ"
  });

  clearPaymentTracking(userId);

  await reply(replyToken, [
    {
      type: "text",
      text: `📩 ได้รับรูปสลิปแล้ว
ชื่อ: ${payment.displayName || customer.displayName}
รหัสลูกค้า: ${payment.customerNo || customer.customerNo}

ระบบได้บันทึกข้อมูลเรียบร้อย กรุณารอแอดมินตรวจสอบ`
    }
  ]);
}

async function handleOrder(replyToken, userId, text) {
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
    let displayName = "ลูกค้า";

    try {
      if (userId) {
        const profile = await getLineProfile(userId);
        displayName = profile.displayName || "ลูกค้า";
      }
    } catch (err) {
      logError("Get profile error:", err.message);
    }

    const customer = await getOrCreateCustomer(userId, displayName);

    if (userId) {
      setPaymentTracking(userId, {
        type: "normal_order",
        customerNo: customer.customerNo,
        userId,
        displayName: customer.displayName,
        itemName: "คำสั่งซื้อสินค้า",
        detail: result.detailLines.map((x) => x.text).join(" | "),
        qty: result.detailLines.reduce((sum, x) => sum + Number(x.qty || 0), 0),
        unitPrice: "",
        total: result.total,
        spotNumbers: ""
      });
    }

    await reply(replyToken, [
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

รหัสลูกค้า: ${customer.customerNo}`
      },
      { type: "text", text: result.summary }
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
    const [prices, stock, spots, campaign, customers] = await Promise.all([
      loadPrices(),
      loadStock(),
      loadSpotSheet(true),
      loadCampaign(true),
      loadCustomers(true)
    ]);

    res.status(200).json({
      ok: true,
      priceItems: Object.keys(prices).length,
      stockModels: Object.keys(stock).length,
      totalSpots: spots.length,
      availableSpots: getAvailableSpots(spots).length,
      totalCustomers: customers.length,
      bookingEnabled: hasGoogleConfig,
      driveEnabled: !!CONFIG.GOOGLE.DRIVE_FOLDER_ID,
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
          await handleSlipImage(replyToken, userId, event.message.id);
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

        const handled = await handleOrder(replyToken, userId, text);
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
  logInfo("LOG_SHEET_NAME =", CONFIG.SHEETS.LOG.NAME);
  logInfo("CUSTOMER_SHEET_NAME =", CONFIG.SHEETS.CUSTOMER.NAME);
  logInfo("DRIVE_ENABLED =", !!CONFIG.GOOGLE.DRIVE_FOLDER_ID);
});