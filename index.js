import express from "express";
import fetch from "node-fetch";
import { google } from "googleapis";

const app = express();
app.use(express.json({ limit: "10mb" }));

// =========================================================
// CONFIG
// =========================================================
const CONFIG = {
  PORT: Number(process.env.PORT || 3000),

  LINE_TOKEN: process.env.LINE_CHANNEL_ACCESS_TOKEN || "",
  LINE_TEXT_LIMIT: Number(process.env.LINE_TEXT_LIMIT || 5000),

  SPREADSHEET_ID:
    process.env.SPREADSHEET_ID || "1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc",

  PRICE_SHEET_NAME: process.env.PRICE_SHEET_NAME || "ชีต1",
  PRICE_CSV_URL:
    process.env.PRICE_CSV_URL ||
    "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0",

  STOCK_SHEET_NAME: process.env.STOCK_SHEET_NAME || "ชีต2",
  STOCK_CSV_URL:
    process.env.STOCK_CSV_URL ||
    "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=262793173",

  SPOT_SHEET_NAME: process.env.SPOT_SHEET_NAME || "ชีต3",
  CAMPAIGN_SHEET_NAME: process.env.CAMPAIGN_SHEET_NAME || "ชีต4",
  LOG_SHEET_NAME: process.env.LOG_SHEET_NAME || "ชีต5",
  CUSTOMER_SHEET_NAME: process.env.CUSTOMER_SHEET_NAME || "ชีต6",
  TRACKING_SHEET_NAME: process.env.TRACKING_SHEET_NAME || "ชีต7",

  GOOGLE_SERVICE_ACCOUNT_EMAIL: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL || "",
  GOOGLE_PRIVATE_KEY: process.env.GOOGLE_PRIVATE_KEY
    ? process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n")
    : "",

  PAYMENT_IMAGE_URL:
    process.env.PAYMENT_IMAGE_URL ||
    "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg",

  FETCH_TIMEOUT_MS: Number(process.env.FETCH_TIMEOUT_MS || 12000),
  FETCH_RETRY: Number(process.env.FETCH_RETRY || 2),
  CACHE_TTL_MS: Number(process.env.CACHE_TTL_MS || 60 * 1000),
  STATE_TTL_MS: Number(process.env.STATE_TTL_MS || 10 * 60 * 1000),
  MAX_SPOTS_PER_BOOKING: Number(process.env.MAX_SPOTS_PER_BOOKING || 5),

  ADMIN_KEY: process.env.ADMIN_KEY || ""
};

function validateConfig() {
  const missing = [];

  if (!CONFIG.LINE_TOKEN) missing.push("LINE_CHANNEL_ACCESS_TOKEN");
  if (!CONFIG.SPREADSHEET_ID) missing.push("SPREADSHEET_ID");
  if (!CONFIG.GOOGLE_SERVICE_ACCOUNT_EMAIL) {
    missing.push("GOOGLE_SERVICE_ACCOUNT_EMAIL");
  }
  if (!CONFIG.GOOGLE_PRIVATE_KEY) {
    missing.push("GOOGLE_PRIVATE_KEY");
  }

  if (missing.length > 0) {
    console.error("❌ Missing ENV:", missing.join(", "));
  }
}

validateConfig();

// =========================================================
// GOOGLE SHEETS
// =========================================================
const hasGoogleConfig =
  !!CONFIG.GOOGLE_SERVICE_ACCOUNT_EMAIL && !!CONFIG.GOOGLE_PRIVATE_KEY;

const auth = hasGoogleConfig
  ? new google.auth.JWT({
      email: CONFIG.GOOGLE_SERVICE_ACCOUNT_EMAIL,
      key: CONFIG.GOOGLE_PRIVATE_KEY,
      scopes: ["https://www.googleapis.com/auth/spreadsheets"]
    })
  : null;

const sheets = hasGoogleConfig
  ? google.sheets({ version: "v4", auth })
  : null;

function ensureGoogleEnabled() {
  if (!sheets) {
    throw new Error(
      "Google Sheets not configured. Please set GOOGLE_SERVICE_ACCOUNT_EMAIL and GOOGLE_PRIVATE_KEY."
    );
  }
}

// =========================================================
// CACHE / STATE
// =========================================================
const cache = {
  prices: { data: null, fetchedAt: 0 },
  stock: { data: null, fetchedAt: 0 },
  spots: { data: null, fetchedAt: 0 },
  campaign: { data: null, fetchedAt: 0 },
  customers: { data: null, fetchedAt: 0 },
  slipIds: { data: null, fetchedAt: 0 },
  trackingLatest: { data: null, fetchedAt: 0 }
};

const sessionStates = new Map();
const paymentTracking = new Map();
const recentSlipMessageIds = new Map();

let bookingQueue = Promise.resolve();

setInterval(() => {
  const now = Date.now();

  for (const [key, state] of sessionStates.entries()) {
    if (state.expiresAt <= now) sessionStates.delete(key);
  }

  for (const [key, payment] of paymentTracking.entries()) {
    if (payment.expiresAt <= now) paymentTracking.delete(key);
  }

  for (const [key, expiresAt] of recentSlipMessageIds.entries()) {
    if (expiresAt <= now) recentSlipMessageIds.delete(key);
  }
}, 60 * 1000);

// =========================================================
// UTILS
// =========================================================
function nowISO() {
  return new Date().toISOString();
}

function nowLocalDate() {
  return new Date().toLocaleDateString("sv-SE", { timeZone: "Asia/Bangkok" });
}

function nowLocalText() {
  return new Date().toLocaleString("sv-SE", { timeZone: "Asia/Bangkok" });
}

function logInfo(...args) {
  console.log(`[${nowISO()}]`, ...args);
}

function logError(...args) {
  console.error(`[${nowISO()}]`, ...args);
}

function makeRef(prefix = "REF") {
  return `${prefix}-${Date.now()}-${Math.random()
    .toString(36)
    .slice(2, 8)
    .toUpperCase()}`;
}

function withTimeout(promise, ms = CONFIG.FETCH_TIMEOUT_MS) {
  return Promise.race([
    promise,
    new Promise((_, reject) =>
      setTimeout(() => reject(new Error(`Timeout after ${ms} ms`)), ms)
    )
  ]);
}

function sleep(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchWithRetry(url, options = {}, retries = CONFIG.FETCH_RETRY) {
  let lastError;

  for (let i = 0; i <= retries; i++) {
    try {
      const res = await withTimeout(fetch(url, options));
      return res;
    } catch (err) {
      lastError = err;
      if (i < retries) {
        await sleep(400 * (i + 1));
      }
    }
  }

  throw lastError;
}

async function fetchText(url) {
  logInfo("Fetching URL:", url);

  const res = await fetchWithRetry(url);
  if (!res.ok) {
    throw new Error(`Fetch failed ${res.status} ${res.statusText}`);
  }

  const text = await res.text();
  if (!text || !text.trim()) {
    throw new Error("Fetched empty response");
  }

  return text;
}

function formatBaht(num) {
  const n = Number(num);
  if (Number.isNaN(n)) return "-";
  return n.toLocaleString("en-US");
}

function normalizeCode(code) {
  return String(code || "")
    .toUpperCase()
    .replace(/\s+/g, "")
    .replace(/[^A-Z0-9-]/g, "")
    .replace(/-+/g, "-")
    .replace(/^([A-Z]+)-?(\d+)$/, "$1$2");
}

function displayCode(code) {
  const clean = normalizeCode(code);
  const m = clean.match(/^([A-Z]+)(\d+)$/);
  return m ? `${m[1]}-${m[2]}` : clean;
}

function displayUnit(unit) {
  return unit === "pack" ? "ซอง" : "กล่อง";
}

function qtyToBar(qty) {
  const n = Number(qty) || 0;
  return n > 0 ? "I ".repeat(n).trim() : "-";
}

function chunkText(text, limit = CONFIG.LINE_TEXT_LIMIT) {
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

function getOrderGuideText() {
  return `หากต้องการสรุปราคา พิมพ์เช่น
OP13 2 ซอง OP15 2 ซอง
หรือ
OP13 1 กล่อง`;
}

function getAddressTemplateText() {
  return `ชื่อผู้รับ:
เบอร์โทร:
ที่อยู่:`;
}

function getSpotSelectionPromptText() {
  return `หากต้องการจอง กรุณาพิมพ์หมายเลขสปอตที่ต้องการ
ตัวอย่าง:
2,3,7
1,3
5`;
}

function getMenuText() {
  return `คำสั่งที่ใช้ได้:
- เมนู
- ราคาสินค้า
- รวมราคา
- Check Rate
- จองสปอตสุ่ม
- แก้ไขที่อยู่จัดส่ง
- เช็คพัสดุ

หากต้องการสอบถามกับร้าน สามารถพิมพ์ข้อความได้ตามปกติ`;
}

function padCustomerNo(num) {
  return String(num).padStart(4, "0");
}

function verifyAdmin(req) {
  const key = req.query.key || req.headers["x-admin-key"];
  return !!CONFIG.ADMIN_KEY && key === CONFIG.ADMIN_KEY;
}

function buildNote(parts = {}) {
  const lines = [];

  for (const [key, value] of Object.entries(parts)) {
    if (value !== undefined && value !== null && String(value).trim() !== "") {
      lines.push(`${key}=${String(value).trim()}`);
    }
  }

  return lines.join(" | ");
}

function formatThaiDateLong(date = new Date()) {
  const months = [
    "มกราคม",
    "กุมภาพันธ์",
    "มีนาคม",
    "เมษายน",
    "พฤษภาคม",
    "มิถุนายน",
    "กรกฎาคม",
    "สิงหาคม",
    "กันยายน",
    "ตุลาคม",
    "พฤศจิกายน",
    "ธันวาคม"
  ];

  const d = new Date(
    date.toLocaleString("en-US", { timeZone: "Asia/Bangkok" })
  );

  const day = d.getDate();
  const month = months[d.getMonth()];
  const year = d.getFullYear();

  return `${day} ${month} ${year}`;
}

// =========================================================
// CSV PARSER
// =========================================================
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

// =========================================================
// SHEET HELPERS
// =========================================================
async function getSheetValues(range) {
  ensureGoogleEnabled();

  const res = await withTimeout(
    sheets.spreadsheets.values.get({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range
    })
  );

  return res.data.values || [];
}

async function appendSheetRow(range, values) {
  ensureGoogleEnabled();

  return withTimeout(
    sheets.spreadsheets.values.append({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      range,
      valueInputOption: "RAW",
      insertDataOption: "INSERT_ROWS",
      requestBody: {
        values: [values]
      }
    })
  );
}

async function batchUpdateSheetValues(data) {
  ensureGoogleEnabled();

  return withTimeout(
    sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: CONFIG.SPREADSHEET_ID,
      requestBody: {
        valueInputOption: "RAW",
        data
      }
    })
  );
}

// =========================================================
// SESSION / PAYMENT STATE
// =========================================================
function setSessionState(userId, mode, data = {}) {
  if (!userId) return;
  sessionStates.set(userId, {
    mode,
    data,
    expiresAt: Date.now() + CONFIG.STATE_TTL_MS
  });
}

function getSessionState(userId) {
  if (!userId) return null;
  const state = sessionStates.get(userId);
  if (!state) return null;

  if (Date.now() > state.expiresAt) {
    sessionStates.delete(userId);
    return null;
  }

  return state;
}

function clearSessionState(userId) {
  if (!userId) return;
  sessionStates.delete(userId);
}

function setPaymentTracking(userId, payload) {
  if (!userId) return;
  paymentTracking.set(userId, {
    ...payload,
    expiresAt: Date.now() + CONFIG.STATE_TTL_MS
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

// =========================================================
// DATA LOADERS
// =========================================================
async function loadPrices(forceRefresh = false) {
  const age = Date.now() - cache.prices.fetchedAt;
  if (!forceRefresh && cache.prices.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.prices.data;
  }

  const csv = await fetchText(CONFIG.PRICE_CSV_URL);
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
  if (!forceRefresh && cache.stock.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.stock.data;
  }

  const csv = await fetchText(CONFIG.STOCK_CSV_URL);
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
  if (!forceRefresh && cache.spots.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.spots.data;
  }

  const rows = await getSheetValues(`${CONFIG.SPOT_SHEET_NAME}!A:C`);
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
    const aNum = Number.isNaN(a.sortNumber)
      ? Number.MAX_SAFE_INTEGER
      : a.sortNumber;
    const bNum = Number.isNaN(b.sortNumber)
      ? Number.MAX_SAFE_INTEGER
      : b.sortNumber;
    return aNum - bNum;
  });

  cache.spots = { data: spots, fetchedAt: Date.now() };
  return spots;
}

async function loadCampaign(forceRefresh = false) {
  const age = Date.now() - cache.campaign.fetchedAt;
  if (!forceRefresh && cache.campaign.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.campaign.data;
  }

  const rows = await getSheetValues(`${CONFIG.CAMPAIGN_SHEET_NAME}!A2:C2`);
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
  if (!forceRefresh && cache.customers.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.customers.data;
  }

  const rows = await getSheetValues(`${CONFIG.CUSTOMER_SHEET_NAME}!A:E`);
  const customers = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const customerNo = String(row[0] || "").trim();
    const userId = String(row[1] || "").trim();
    const latestName = String(row[2] || "").trim();
    const createdAt = String(row[3] || "").trim();
    const shippingAddress = String(row[4] || "").trim();

    if (!customerNo && !userId) continue;

    customers.push({
      rowNumber: i + 1,
      customerNo,
      userId,
      latestName,
      createdAt,
      shippingAddress
    });
  }

  cache.customers = { data: customers, fetchedAt: Date.now() };
  return customers;
}

async function loadRecentSlipMessageIds(forceRefresh = false) {
  const age = Date.now() - cache.slipIds.fetchedAt;
  if (!forceRefresh && cache.slipIds.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.slipIds.data;
  }

  const rows = await getSheetValues(`${CONFIG.LOG_SHEET_NAME}!N:N`);
  const ids = new Set();

  for (let i = 1; i < rows.length; i++) {
    const note = String(rows[i]?.[0] || "").trim();
    if (!note) continue;

    const match = note.match(/lineMessageId=([A-Za-z0-9_-]+)/);
    if (match?.[1]) ids.add(match[1]);
  }

  cache.slipIds = {
    data: ids,
    fetchedAt: Date.now()
  };

  return ids;
}

async function loadLatestTrackingImage(forceRefresh = false) {
  const age = Date.now() - cache.trackingLatest.fetchedAt;
  if (!forceRefresh && cache.trackingLatest.data && age < CONFIG.CACHE_TTL_MS) {
    return cache.trackingLatest.data;
  }

  const rows = await getSheetValues(`${CONFIG.TRACKING_SHEET_NAME}!A:A`);

  let latest = null;

  for (let i = rows.length - 1; i >= 1; i--) {
    const value = String(rows[i]?.[0] || "").trim();
    if (!value) continue;

    latest = {
      rowNumber: i + 1,
      imageUrl: value
    };
    break;
  }

  cache.trackingLatest = {
    data: latest,
    fetchedAt: Date.now()
  };

  return latest;
}

// =========================================================
// CUSTOMER MASTER
// =========================================================
async function getLineProfile(userId) {
  if (!userId) throw new Error("Missing userId");

  const res = await fetchWithRetry(`https://api.line.me/v2/bot/profile/${userId}`, {
    method: "GET",
    headers: {
      Authorization: `Bearer ${CONFIG.LINE_TOKEN}`
    }
  });

  const body = await res.text();

  if (!res.ok) {
    throw new Error(`Get profile failed ${res.status}: ${body}`);
  }

  return JSON.parse(body);
}

async function getDisplayName(userId) {
  if (!userId) return "ลูกค้า";

  try {
    const profile = await getLineProfile(userId);
    return profile.displayName || "ลูกค้า";
  } catch (err) {
    logError("Get profile error:", err.message);
    return "ลูกค้า";
  }
}

async function getOrCreateCustomer(userId, displayName) {
  const customers = await loadCustomers(true);

  if (userId) {
    const found = customers.find((c) => c.userId === userId);
    if (found) {
      const updates = [];
      if (displayName && found.latestName !== displayName) {
        updates.push({
          range: `${CONFIG.CUSTOMER_SHEET_NAME}!C${found.rowNumber}`,
          values: [[displayName]]
        });
      }

      if (updates.length > 0) {
        await batchUpdateSheetValues(updates);
        await loadCustomers(true);
      }

      return {
        rowNumber: found.rowNumber,
        customerNo: found.customerNo,
        userId: found.userId,
        displayName: displayName || found.latestName || "ลูกค้า",
        shippingAddress: found.shippingAddress || ""
      };
    }
  }

  const maxNo = customers.reduce((max, c) => {
    const num = Number(c.customerNo);
    return Number.isNaN(num) ? max : Math.max(max, num);
  }, 0);

  const nextNo = padCustomerNo(maxNo + 1);

  await appendSheetRow(`${CONFIG.CUSTOMER_SHEET_NAME}!A:E`, [
    nextNo,
    userId || "",
    displayName || "ลูกค้า",
    nowLocalText(),
    ""
  ]);

  const refreshed = await loadCustomers(true);
  const created = refreshed.find((c) => c.customerNo === nextNo);

  return {
    rowNumber: created?.rowNumber || "",
    customerNo: nextNo,
    userId: userId || "",
    displayName: displayName || "ลูกค้า",
    shippingAddress: ""
  };
}

async function saveCustomerShippingAddress(userId, addressBlock, displayName = "ลูกค้า") {
  const customer = await getOrCreateCustomer(userId, displayName);

  if (!customer.rowNumber) {
    throw new Error("Customer row not found");
  }

  await batchUpdateSheetValues([
    {
      range: `${CONFIG.CUSTOMER_SHEET_NAME}!E${customer.rowNumber}`,
      values: [[addressBlock]]
    }
  ]);

  const refreshed = await loadCustomers(true);
  const updated = refreshed.find((c) => c.customerNo === customer.customerNo);

  return {
    ...customer,
    shippingAddress: updated?.shippingAddress || addressBlock
  };
}

// =========================================================
// LOG HELPERS - ชีต5 A:N
// =========================================================
async function appendOrderLog({
  type = "",
  customerNo = "",
  userId = "",
  displayName = "",
  itemName = "",
  detail = "",
  qty = "",
  unitPrice = "",
  total = "",
  spotNumbers = "",
  slipStatus = "",
  slipUrl = "",
  note = ""
}) {
  await appendSheetRow(`${CONFIG.LOG_SHEET_NAME}!A:N`, [
    nowLocalText(),
    type,
    customerNo,
    userId,
    displayName,
    itemName,
    detail,
    qty !== "" && qty != null ? String(qty) : "",
    unitPrice !== "" && unitPrice != null ? String(unitPrice) : "",
    total !== "" && total != null ? String(total) : "",
    spotNumbers,
    slipStatus,
    slipUrl,
    note
  ]);
}

// =========================================================
// DISPLAY HELPERS
// =========================================================
function buildCampaignIntroText(campaign) {
  const lines = [];

  if (campaign.title) {
    lines.push("🎯 รายการจองปัจจุบัน");
    lines.push(campaign.title);
  } else {
    lines.push("🎯 ระบบจองสปอต");
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

function buildBankText(extra = "") {
  const parts = [
    "สามารถชำระเงินผ่านช่องทางอื่น ๆ ได้ดังนี้",
    "",
    "ชื่อบัญชี ปรัชญา สุดใจดี",
    "",
    "K-Bank 0503228092",
    "",
    "True Wallet 0982652650",
    "",
    "✨ ชำระแล้วโปรดแปะ Pay Slip การโอนทุกครั้ง ✨"
  ];

  if (extra) {
    parts.push("", extra);
  }

  return parts.join("\n");
}

// =========================================================
// PARSERS
// =========================================================
function parseItems(text) {
  const regex =
    /([A-Z]+\s*-?\s*\d+)\s*(?:x?\s*)?(\d+)\s*(ซอง|ซ็อง|pack|box|กล่อง|บ็อก)/gi;

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

function parseAddressTemplate(text) {
  const input = String(text || "").trim();

  const recipientMatch = input.match(/ชื่อผู้รับ\s*:\s*(.+)/i);
  const phoneMatch = input.match(/เบอร์โทร\s*:\s*(.+)/i);
  const addressMatch = input.match(/ที่อยู่\s*:\s*([\s\S]+)/i);

  if (!recipientMatch || !phoneMatch || !addressMatch) {
    return null;
  }

  const recipient = recipientMatch[1].trim();
  const phone = phoneMatch[1].trim();
  const address = addressMatch[1].trim();

  if (!recipient || !phone || !address) {
    return null;
  }

  const onlyDigits = phone.replace(/\D/g, "");
  if (onlyDigits.length < 9 || onlyDigits.length > 15) {
    return null;
  }

  return {
    recipient,
    phone,
    address
  };
}

function formatAddressBlock(addressData) {
  return `ชื่อผู้รับ: ${addressData.recipient}
เบอร์โทร: ${addressData.phone}
ที่อยู่: ${addressData.address}`;
}

function parseSpotNumbers(text) {
  const clean = String(text || "").trim();
  if (!clean) return null;

  const parts = clean
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  if (parts.length === 0) return null;

  const numbers = [];

  for (const part of parts) {
    if (!/^\d+$/.test(part)) {
      return null;
    }

    const num = Number(part);

    if (!Number.isInteger(num) || num <= 0) {
      return null;
    }

    numbers.push(String(num));
  }

  const unique = [...new Set(numbers)];
  if (unique.length > CONFIG.MAX_SPOTS_PER_BOOKING) {
    return {
      error: `จองได้สูงสุด ${CONFIG.MAX_SPOTS_PER_BOOKING} สปอตต่อครั้ง`
    };
  }

  return { values: unique };
}

// =========================================================
// LINE HELPERS
// =========================================================
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

  if (!replyToken || safeMessages.length === 0) return false;

  try {
    const res = await fetchWithRetry("https://api.line.me/v2/bot/message/reply", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${CONFIG.LINE_TOKEN}`
      },
      body: JSON.stringify({
        replyToken,
        messages: safeMessages
      })
    });

    const body = await res.text();

    if (!res.ok) {
      throw new Error(`LINE reply failed ${res.status}: ${body}`);
    }

    return true;
  } catch (err) {
    logError("LINE reply error:", err.message);
    return false;
  }
}

// =========================================================
// ORDER CALCULATOR
// =========================================================
async function calculateOrder(text) {
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

    const detailText = `${displayCode(item.code)} ${displayUnit(item.unit)}ละ ${formatBaht(
      unitPrice
    )} x${item.qty} = ${formatBaht(sum)} บาท`;

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
    paymentText: `📌 กรุณาโอน ${formatBaht(total)} บาท`,
    total,
    detailLines
  };
}

async function buildOrderPaymentMessages(userId, originalText) {
  const result = await calculateOrder(originalText);
  if (!result) return null;

  if (result.status === "guide") {
    return {
      handled: true,
      messages: [{ type: "text", text: result.message }]
    };
  }

  if (result.status === "invalid") {
    return {
      handled: true,
      messages: [{ type: "text", text: result.message }]
    };
  }

  const displayName = await getDisplayName(userId);
  const customer = await getOrCreateCustomer(userId, displayName);

  if (!customer.shippingAddress) {
    return {
      handled: true,
      requiresAddress: true,
      messages: [
        {
          type: "text",
          text: `ยังไม่พบข้อมูลที่อยู่จัดส่งของคุณ

กรุณาก็อปปี้เทมเพลตด้านล่าง แล้วเพิ่มข้อมูลจัดส่งของคุณให้ครบถ้วน

${getAddressTemplateText()}`
        }
      ],
      state: {
        mode: "address_create",
        data: {
          origin: "order",
          originalText
        }
      }
    };
  }

  const paymentRef = makeRef("PAY");

  setPaymentTracking(userId, {
    type: "normal_order",
    paymentRef,
    customerNo: customer.customerNo,
    userId,
    displayName: customer.displayName,
    itemName: "คำสั่งซื้อสินค้า",
    detail: result.detailLines.map((x) => x.text).join(" | "),
    qty: result.detailLines.reduce((sum, x) => sum + Number(x.qty || 0), 0),
    unitPrice: "",
    total: result.total,
    spotNumbers: "",
    orderText: originalText
  });

  await appendOrderLog({
    type: "normal_order",
    customerNo: customer.customerNo,
    userId,
    displayName: customer.displayName,
    itemName: "คำสั่งซื้อสินค้า",
    detail: result.detailLines.map((x) => x.text).join(" | "),
    qty: result.detailLines.reduce((sum, x) => sum + Number(x.qty || 0), 0),
    unitPrice: "",
    total: result.total,
    spotNumbers: "",
    slipStatus: "รอส่งสลิป",
    slipUrl: "",
    note: buildNote({
      action: "payment_pending_created",
      paymentRef
    })
  });

  return {
    handled: true,
    messages: [
      { type: "text", text: result.paymentText },
      {
        type: "image",
        originalContentUrl: CONFIG.PAYMENT_IMAGE_URL,
        previewImageUrl: CONFIG.PAYMENT_IMAGE_URL
      },
      {
        type: "text",
        text: buildBankText(`รหัสลูกค้า: ${customer.customerNo}
ที่อยู่จัดส่งปัจจุบัน:
${customer.shippingAddress}`)
      },
      { type: "text", text: result.summary }
    ]
  };
}

// =========================================================
// SPOT HELPERS
// =========================================================
async function validateSpecificSpots(spotNumbers) {
  const spots = await loadSpotSheet(true);
  const spotMap = new Map(spots.map((spot) => [String(spot.spotNumber), spot]));

  const invalidSpots = [];
  const occupiedSpots = [];
  const validSpots = [];

  for (const spotNo of spotNumbers) {
    const spot = spotMap.get(String(spotNo));

    if (!spot) {
      invalidSpots.push(String(spotNo));
      continue;
    }

    if (spot.name) {
      occupiedSpots.push(String(spotNo));
      continue;
    }

    validSpots.push(spot);
  }

  return {
    invalidSpots,
    occupiedSpots,
    validSpots
  };
}

async function enqueueBooking(task) {
  const run = bookingQueue.then(task, task);
  bookingQueue = run.catch(() => {});
  return run;
}

async function reserveSpecificSpots({ spotNumbers, displayName }) {
  return enqueueBooking(async () => {
    const { invalidSpots, occupiedSpots, validSpots } =
      await validateSpecificSpots(spotNumbers);

    if (invalidSpots.length > 0) {
      return {
        ok: false,
        message: `ไม่พบหมายเลขสปอต: ${invalidSpots.join(", ")}`
      };
    }

    if (occupiedSpots.length > 0) {
      return {
        ok: false,
        message: `สปอตต่อไปนี้ถูกจองแล้ว: ${occupiedSpots.join(", ")}`
      };
    }

    if (validSpots.length === 0) {
      return {
        ok: false,
        message: "ไม่พบสปอตที่จองได้"
      };
    }

    const updates = [];

    for (const spot of validSpots) {
      updates.push({
        range: `${CONFIG.SPOT_SHEET_NAME}!B${spot.rowNumber}`,
        values: [[displayName]]
      });
      updates.push({
        range: `${CONFIG.SPOT_SHEET_NAME}!C${spot.rowNumber}`,
        values: [["รอชำระ"]]
      });
    }

    await batchUpdateSheetValues(updates);
    const refreshed = await loadSpotSheet(true);

    return {
      ok: true,
      spots: validSpots.map((spot) => String(spot.spotNumber)),
      allBookedText: buildAllBookedSpotsText(refreshed)
    };
  });
}

async function markSpecificSpotsSlipSent(spotNumbersText) {
  const spots = await loadSpotSheet(true);

  const targetNumbers = String(spotNumbersText || "")
    .split(",")
    .map((x) => x.trim())
    .filter(Boolean);

  const updates = [];

  for (const spot of spots) {
    if (targetNumbers.includes(String(spot.spotNumber))) {
      updates.push({
        range: `${CONFIG.SPOT_SHEET_NAME}!C${spot.rowNumber}`,
        values: [["ส่งสลิปแล้ว"]]
      });
    }
  }

  if (updates.length === 0) {
    return { ok: false };
  }

  await batchUpdateSheetValues(updates);
  await loadSpotSheet(true);

  return { ok: true };
}

async function buildSpotBookingStartMessages(userId) {
  const displayName = await getDisplayName(userId);
  const customer = await getOrCreateCustomer(userId, displayName);

  if (!customer.shippingAddress) {
    return {
      handled: true,
      requiresAddress: true,
      messages: [
        {
          type: "text",
          text: `ยังไม่พบข้อมูลที่อยู่จัดส่งของคุณ

กรุณาก็อปปี้เทมเพลตด้านล่าง แล้วเพิ่มข้อมูลจัดส่งของคุณให้ครบถ้วน

${getAddressTemplateText()}`
        }
      ],
      state: {
        mode: "address_create",
        data: {
          origin: "spot_booking"
        }
      }
    };
  }

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
    return {
      handled: true,
      messages: [...messages, { type: "text", text: "สปอตเต็ม" }]
    };
  }

  return {
    handled: true,
    messages: [...messages, { type: "text", text: getSpotSelectionPromptText() }],
    state: {
      mode: "spot_booking",
      data: {}
    }
  };
}

// =========================================================
// TRACKING
// =========================================================
async function buildTrackingMessages() {
  const latest = await loadLatestTrackingImage(true);

  if (!latest || !latest.imageUrl) {
    return [
      {
        type: "text",
        text: "ยังไม่พบข้อมูลพัสดุในระบบ กรุณาติดต่อร้าน"
      }
    ];
  }

  const dateText = formatThaiDateLong(new Date());

  return [
    {
      type: "text",
      text: `📦 สถานะพัสดุอัปเดตล่าสุด
จัดส่งวันที่ ${dateText}

รายการพัสดุอยู่ในภาพด้านล่าง
กรุณาเลื่อนดูและค้นหาชื่อของคุณได้เลย

หากไม่พบ กรุณาทักแชตร้านได้ทันที`
    },
    {
      type: "image",
      originalContentUrl: latest.imageUrl,
      previewImageUrl: latest.imageUrl
    }
  ];
}

// =========================================================
// DUPLICATE SLIP CHECK
// =========================================================
async function isDuplicateSlipMessageId(messageId) {
  if (!messageId) return false;

  const memExpiresAt = recentSlipMessageIds.get(messageId);
  if (memExpiresAt && memExpiresAt > Date.now()) {
    return true;
  }

  const logged = await loadRecentSlipMessageIds();
  return logged.has(messageId);
}

function markRecentSlipMessageId(messageId) {
  if (!messageId) return;
  recentSlipMessageIds.set(messageId, Date.now() + 24 * 60 * 60 * 1000);
}

// =========================================================
// ADMIN DASHBOARD
// =========================================================
async function getAdminSummary() {
  const [spots, campaign] = await Promise.all([
    loadSpotSheet(true),
    loadCampaign(true)
  ]);

  const sortedSpots = [...spots].sort(
    (a, b) => Number(a.spotNumber) - Number(b.spotNumber)
  );

  return {
    serviceTime: nowISO(),
    campaign,
    totalSpots: sortedSpots.length,
    availableSpots: sortedSpots.filter((x) => !x.name).length,
    bookedSpots: sortedSpots.filter((x) => x.name).length,
    spots: sortedSpots.map((spot) => ({
      number: spot.spotNumber,
      name: spot.name || "",
      paymentStatus: spot.paymentStatus || ""
    }))
  };
}

function renderSpotsGrid(summary) {
  const spots = Array.isArray(summary?.spots) ? summary.spots : [];

  if (spots.length === 0) {
    return `
      <section class="spot-board">
        <div class="section-head">
          <h2>SPOT Status</h2>
          <p>ยังไม่พบข้อมูลสปอต</p>
        </div>
      </section>
    `;
  }

  return `
    <section class="spot-board">
      <div class="section-head">
        <h2>SPOT Status</h2>
        <p>ตรวจสอบรายชื่อผู้จองแต่ละช่องได้แบบเรียลไทม์</p>
      </div>

      <div class="spot-grid">
        ${spots
          .map((spot) => {
            const isFilled = !!spot.name;
            return `
              <div class="spot-card ${isFilled ? "filled" : "empty"}">
                <div class="spot-no">SPOT ${spot.number}</div>
                <div class="spot-name">${spot.name || "ว่าง"}</div>
                ${
                  isFilled && spot.paymentStatus
                    ? `<div class="spot-status">${spot.paymentStatus}</div>`
                    : `<div class="spot-status empty-status">พร้อมจอง</div>`
                }
              </div>
            `;
          })
          .join("")}
      </div>
    </section>
  `;
}

function renderAdminHtml(summary) {
  const campaignTitle = summary?.campaign?.title || "-";
  const campaignPrice = formatBaht(summary?.campaign?.price || 0);
  const campaignImage = summary?.campaign?.imageUrl || "";

  return `
<!doctype html>
<html lang="th">
<head>
<meta charset="utf-8" />
<meta name="viewport" content="width=device-width,initial-scale=1" />
<title>ระบบการจอง SPOT TCG Ordinary</title>
<style>
  * {
    box-sizing: border-box;
  }

  body {
    margin: 0;
    font-family: Arial, sans-serif;
    color: #111827;
    background:
      radial-gradient(circle at top left, #dbeafe 0%, transparent 35%),
      radial-gradient(circle at top right, #ede9fe 0%, transparent 30%),
      linear-gradient(135deg, #f8fafc 0%, #eef2ff 100%);
    min-height: 100vh;
  }

  .container {
    max-width: 1440px;
    margin: 0 auto;
    padding: 28px;
  }

  .hero {
    display: grid;
    grid-template-columns: 1.2fr 0.9fr;
    gap: 24px;
    align-items: stretch;
    margin-bottom: 24px;
  }

  .hero-left {
    background: rgba(255,255,255,0.82);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.7);
    border-radius: 28px;
    padding: 28px;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.10);
  }

  .hero-badge {
    display: inline-block;
    padding: 8px 14px;
    border-radius: 999px;
    background: linear-gradient(135deg, #111827, #374151);
    color: #fff;
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.3px;
    margin-bottom: 14px;
  }

  .hero-title {
    margin: 0;
    font-size: 42px;
    line-height: 1.08;
    font-weight: 800;
    color: #0f172a;
  }

  .hero-subtitle {
    margin: 14px 0 0;
    font-size: 16px;
    line-height: 1.7;
    color: #475569;
  }

  .updated {
    margin-top: 18px;
    font-size: 13px;
    color: #64748b;
  }

  .hero-right {
    background: rgba(255,255,255,0.82);
    backdrop-filter: blur(10px);
    border: 1px solid rgba(255,255,255,0.7);
    border-radius: 28px;
    padding: 18px;
    box-shadow: 0 20px 50px rgba(15, 23, 42, 0.10);
    display: flex;
    align-items: center;
    justify-content: center;
    min-height: 320px;
    overflow: hidden;
  }

  .campaign-image {
    width: 100%;
    height: 100%;
    min-height: 280px;
    object-fit: cover;
    border-radius: 22px;
    display: block;
    box-shadow: 0 10px 30px rgba(15, 23, 42, 0.14);
  }

  .image-placeholder {
    width: 100%;
    min-height: 280px;
    border-radius: 22px;
    display: flex;
    align-items: center;
    justify-content: center;
    background: linear-gradient(135deg, #e2e8f0, #f8fafc);
    color: #64748b;
    font-size: 18px;
    font-weight: 700;
    text-align: center;
    padding: 24px;
  }

  .stats-grid {
    display: grid;
    grid-template-columns: repeat(5, minmax(200px, 1fr));
    gap: 18px;
    margin-bottom: 28px;
  }

  .stat-card {
    position: relative;
    overflow: hidden;
    background: rgba(255,255,255,0.88);
    border: 1px solid rgba(255,255,255,0.7);
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 16px 40px rgba(15, 23, 42, 0.08);
  }

  .stat-card::after {
    content: "";
    position: absolute;
    inset: auto -30px -30px auto;
    width: 110px;
    height: 110px;
    border-radius: 50%;
    background: linear-gradient(135deg, rgba(59,130,246,0.10), rgba(139,92,246,0.10));
  }

  .stat-label {
    position: relative;
    z-index: 1;
    font-size: 14px;
    color: #64748b;
    margin-bottom: 12px;
    font-weight: 600;
  }

  .stat-value {
    position: relative;
    z-index: 1;
    font-size: 38px;
    line-height: 1;
    font-weight: 800;
    color: #0f172a;
    word-break: break-word;
  }

  .campaign-title-card .stat-value {
    font-size: 28px;
    line-height: 1.25;
  }

  .spot-board {
    background: rgba(255,255,255,0.88);
    border: 1px solid rgba(255,255,255,0.7);
    border-radius: 28px;
    padding: 24px;
    box-shadow: 0 16px 40px rgba(15, 23, 42, 0.08);
  }

  .section-head h2 {
    margin: 0 0 8px;
    font-size: 28px;
    color: #0f172a;
  }

  .section-head p {
    margin: 0 0 18px;
    color: #64748b;
    font-size: 14px;
  }

  .spot-grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 14px;
  }

  .spot-card {
    border-radius: 20px;
    padding: 18px 14px;
    text-align: center;
    box-shadow: 0 8px 20px rgba(15, 23, 42, 0.05);
    border: 1px solid transparent;
  }

  .spot-card.empty {
    background: linear-gradient(135deg, #ecfdf5, #d1fae5);
    border-color: #6ee7b7;
  }

  .spot-card.filled {
    background: linear-gradient(135deg, #fee2e2, #fecaca);
    border-color: #fca5a5;
  }

  .spot-no {
    font-size: 13px;
    font-weight: 700;
    letter-spacing: 0.4px;
    color: #475569;
    margin-bottom: 8px;
  }

  .spot-name {
    font-size: 22px;
    font-weight: 800;
    color: #0f172a;
    line-height: 1.25;
    word-break: break-word;
    min-height: 54px;
    display: flex;
    align-items: center;
    justify-content: center;
  }

  .spot-status {
    margin-top: 10px;
    font-size: 12px;
    font-weight: 700;
    color: #334155;
    background: rgba(255,255,255,0.7);
    border-radius: 999px;
    padding: 6px 10px;
    display: inline-block;
  }

  .empty-status {
    color: #047857;
    background: rgba(255,255,255,0.72);
  }

  .footer-note {
    margin-top: 18px;
    font-size: 13px;
    color: #64748b;
  }

  @media (max-width: 1200px) {
    .stats-grid {
      grid-template-columns: repeat(2, minmax(220px, 1fr));
    }
  }

  @media (max-width: 900px) {
    .hero {
      grid-template-columns: 1fr;
    }

    .hero-title {
      font-size: 34px;
    }
  }

  @media (max-width: 640px) {
    .container {
      padding: 16px;
    }

    .stats-grid {
      grid-template-columns: 1fr;
    }

    .hero-title {
      font-size: 28px;
    }

    .stat-value {
      font-size: 32px;
    }

    .campaign-title-card .stat-value {
      font-size: 24px;
    }

    .spot-grid {
      grid-template-columns: repeat(2, minmax(120px, 1fr));
    }

    .spot-name {
      font-size: 18px;
      min-height: 48px;
    }
  }
</style>
</head>
<body>
  <div class="container">
    <section class="hero">
      <div class="hero-left">
        <div class="hero-badge">SPOT CAMPAIGN DASHBOARD</div>
        <h1 class="hero-title">ระบบการจอง SPOT TCG Ordinary</h1>
        <p class="hero-subtitle">
          หน้าสรุปภาพรวมแคมเปญสำหรับติดตามจำนวนสปอต ราคา และสถานะการจองแบบเรียลไทม์
        </p>
        <div class="updated">Updated: ${summary.serviceTime}</div>
      </div>

      <div class="hero-right">
        ${
          campaignImage
            ? `<img class="campaign-image" src="${campaignImage}" alt="Campaign Image" />`
            : `<div class="image-placeholder">ยังไม่มีรูปแคมเปญ</div>`
        }
      </div>
    </section>

    <section class="stats-grid">
      <div class="stat-card campaign-title-card">
        <div class="stat-label">Campaign</div>
        <div class="stat-value">${campaignTitle}</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Campaign Price</div>
        <div class="stat-value">${campaignPrice} ฿</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Total Spots</div>
        <div class="stat-value">${summary.totalSpots}</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Available Spots</div>
        <div class="stat-value">${summary.availableSpots}</div>
      </div>

      <div class="stat-card">
        <div class="stat-label">Booked Spots</div>
        <div class="stat-value">${summary.bookedSpots}</div>
      </div>
    </section>

    ${renderSpotsGrid(summary)}

    <div class="footer-note">
      Ordinary Campaign Dashboard • Focus view for SPOT booking only
    </div>
  </div>
</body>
</html>
  `;
}

app.get("/admin", async (req, res) => {
  if (!verifyAdmin(req)) {
    return res.status(401).send("Unauthorized");
  }

  try {
    const summary = await getAdminSummary();
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.status(200).send(renderAdminHtml(summary));
  } catch (err) {
    logError("Admin page error:", err.message);
    res.status(500).send("Admin dashboard error");
  }
});

app.get("/admin/api/summary", async (req, res) => {
  if (!verifyAdmin(req)) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }

  try {
    const summary = await getAdminSummary();
    res.status(200).json({ ok: true, summary });
  } catch (err) {
    logError("Admin api error:", err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});
// =========================================================
// MESSAGE HANDLERS
// =========================================================
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

async function handleAddressInput(replyToken, userId, text, mode) {
  if (text.trim() === "ยกเลิก") {
    clearSessionState(userId);
    await reply(replyToken, [{ type: "text", text: "ยกเลิกการกรอกข้อมูลที่อยู่แล้ว" }]);
    return;
  }

  const parsed = parseAddressTemplate(text);

  if (!parsed) {
    await reply(replyToken, [
      {
        type: "text",
        text: `ข้อมูลยังไม่ครบ กรุณาก็อปปี้เทมเพลตด้านล่าง แล้วเพิ่มข้อมูลจัดส่งของคุณให้ครบถ้วน

${getAddressTemplateText()}`
      }
    ]);
    return;
  }

  const displayName = await getDisplayName(userId);
  const addressBlock = formatAddressBlock(parsed);
  await saveCustomerShippingAddress(userId, addressBlock, displayName);

  const state = getSessionState(userId);
  clearSessionState(userId);

  await appendOrderLog({
    type: "customer",
    customerNo: "",
    userId,
    displayName,
    itemName: mode === "address_update" ? "แก้ไขที่อยู่" : "บันทึกที่อยู่",
    detail: addressBlock,
    qty: "",
    unitPrice: "",
    total: "",
    spotNumbers: "",
    slipStatus: "",
    slipUrl: "",
    note: buildNote({
      action: mode === "address_update" ? "address_updated" : "address_created"
    })
  });

  if (mode === "address_create" && state?.data?.origin === "order" && state?.data?.originalText) {
    const orderResult = await buildOrderPaymentMessages(userId, state.data.originalText);

    const messages = [
      {
        type: "text",
        text: `บันทึกที่อยู่จัดส่งเรียบร้อยแล้ว

กรุณาตรวจสอบข้อมูลของคุณอีกครั้ง

${addressBlock}`
      }
    ];

    if (orderResult?.messages?.length) {
      messages.push(...orderResult.messages);
    }

    await reply(replyToken, messages.slice(0, 5));
    return;
  }

  if (mode === "address_create" && state?.data?.origin === "spot_booking") {
    const spotStart = await buildSpotBookingStartMessages(userId);

    const messages = [
      {
        type: "text",
        text: `บันทึกที่อยู่จัดส่งเรียบร้อยแล้ว

กรุณาตรวจสอบข้อมูลของคุณอีกครั้ง

${addressBlock}`
      }
    ];

    if (spotStart?.state?.mode) {
      setSessionState(userId, spotStart.state.mode, spotStart.state.data || {});
    }

    if (spotStart?.messages?.length) {
      messages.push(...spotStart.messages);
    }

    await reply(replyToken, messages.slice(0, 5));
    return;
  }

  await reply(replyToken, [
    {
      type: "text",
      text:
        mode === "address_update"
          ? `แก้ไขที่อยู่จัดส่งเรียบร้อยแล้ว

กรุณาตรวจสอบข้อมูลล่าสุดของคุณอีกครั้ง

${addressBlock}`
          : `บันทึกที่อยู่จัดส่งเรียบร้อยแล้ว

กรุณาตรวจสอบข้อมูลของคุณอีกครั้ง

${addressBlock}`
    }
  ]);
}

async function handleSpotBookingStart(replyToken, userId) {
  const result = await buildSpotBookingStartMessages(userId);

  if (result?.state?.mode) {
    setSessionState(userId, result.state.mode, result.state.data || {});
  }

  await reply(replyToken, result.messages || [{ type: "text", text: "เกิดข้อผิดพลาด" }]);
}

async function handleSpotBookingSelection(replyToken, userId, text) {
  if (text.trim() === "ยกเลิก") {
    clearSessionState(userId);
    await reply(replyToken, [{ type: "text", text: "ยกเลิกการจองสปอตแล้ว" }]);
    return;
  }

  const parsed = parseSpotNumbers(text);

  if (!parsed || !parsed.values) {
    await reply(replyToken, [
      {
        type: "text",
        text:
          parsed?.error ||
          `กรุณาพิมพ์หมายเลขสปอตให้ถูกต้อง
ตัวอย่าง:
2,3,7
1,3
5`
      }
    ]);
    return;
  }

  const displayName = await getDisplayName(userId);
  const customer = await getOrCreateCustomer(userId, displayName);
  const campaign = await loadCampaign(true);

  const validation = await validateSpecificSpots(parsed.values);

  if (validation.invalidSpots.length > 0) {
    await reply(replyToken, [
      {
        type: "text",
        text: `ไม่พบหมายเลขสปอต: ${validation.invalidSpots.join(", ")}`
      }
    ]);
    return;
  }

  if (validation.occupiedSpots.length > 0) {
    await reply(replyToken, [
      {
        type: "text",
        text: `สปอตต่อไปนี้ถูกจองแล้ว: ${validation.occupiedSpots.join(", ")}`
      }
    ]);
    return;
  }

  const qty = validation.validSpots.length;
  const total = (campaign.price || 0) * qty;

  setSessionState(userId, "spot_confirm", {
    spotNumbers: parsed.values,
    displayName: customer.displayName,
    customerNo: customer.customerNo,
    campaignTitle: campaign.title || "รายการจองสปอต",
    unitPrice: campaign.price || 0,
    total
  });

  await reply(replyToken, [
    {
      type: "text",
      text: `คุณต้องการจองสปอต: ${parsed.values.join(", ")} ใช่หรือไม่

พิมพ์:
ยืนยัน
หรือ
ยกเลิก`
    },
    {
      type: "text",
      text: `🧾 สรุปก่อนยืนยัน
${campaign.title || "รายการจองสปอต"}
ราคา / สปอต: ${formatBaht(campaign.price)} บาท
จำนวน: ${qty} สปอต
รวมทั้งหมด: ${formatBaht(total)} บาท`
    }
  ]);
}

async function handleSpotBookingConfirm(replyToken, userId, text) {
  const state = getSessionState(userId);

  if (!state || state.mode !== "spot_confirm") {
    await reply(replyToken, [
      { type: "text", text: "ไม่พบรายการจองที่รอยืนยัน กรุณาเริ่มใหม่อีกครั้ง" }
    ]);
    return;
  }

  const input = text.trim();

  if (input === "ยกเลิก") {
    clearSessionState(userId);
    await reply(replyToken, [{ type: "text", text: "ยกเลิกการจองสปอตแล้ว" }]);
    return;
  }

  if (input !== "ยืนยัน") {
    await reply(replyToken, [{ type: "text", text: "กรุณาพิมพ์ 'ยืนยัน' หรือ 'ยกเลิก'" }]);
    return;
  }

  const {
    spotNumbers,
    displayName,
    customerNo,
    campaignTitle,
    unitPrice,
    total
  } = state.data;

  const result = await reserveSpecificSpots({
    spotNumbers,
    displayName
  });

  clearSessionState(userId);

  if (!result.ok) {
    await reply(replyToken, [{ type: "text", text: result.message }]);
    return;
  }

  const paymentRef = makeRef("SPOT");

  if (userId) {
    setPaymentTracking(userId, {
      type: "spot_booking",
      paymentRef,
      customerNo,
      userId,
      displayName,
      itemName: campaignTitle,
      detail: "จองสปอตตามหมายเลข",
      qty: result.spots.length,
      unitPrice,
      total,
      spotNumbers: result.spots.join(", ")
    });
  }

  await appendOrderLog({
    type: "spot_booking",
    customerNo,
    userId,
    displayName,
    itemName: campaignTitle,
    detail: "จองสปอตตามหมายเลข",
    qty: result.spots.length,
    unitPrice,
    total,
    spotNumbers: result.spots.join(", "),
    slipStatus: "รอส่งสลิป",
    slipUrl: "",
    note: buildNote({
      action: "spot_booking_created",
      paymentRef
    })
  });

  await reply(replyToken, [
    {
      type: "text",
      text: `✅ จองสำเร็จ
ชื่อ: ${displayName}
รหัสลูกค้า: ${customerNo}
หมายเลขสปอต: ${result.spots.join(", ")}`
    },
    {
      type: "text",
      text: result.allBookedText
    },
    {
      type: "image",
      originalContentUrl: CONFIG.PAYMENT_IMAGE_URL,
      previewImageUrl: CONFIG.PAYMENT_IMAGE_URL
    },
    {
      type: "text",
      text: buildBankText("")
    },
    {
      type: "text",
      text: `🧾 สรุปยอดชำระ
${campaignTitle}
ราคา / สปอต: ${formatBaht(unitPrice)} บาท
จำนวน: ${result.spots.length} สปอต
รวมทั้งหมด: ${formatBaht(total)} บาท`
    }
  ]);
}

async function handleOrder(replyToken, userId, text) {
  const result = await buildOrderPaymentMessages(userId, text);
  if (!result?.handled) return false;

  if (result.requiresAddress && result.state?.mode) {
    setSessionState(userId, result.state.mode, result.state.data || {});
  }

  await reply(replyToken, result.messages || [{ type: "text", text: "เกิดข้อผิดพลาด" }]);
  return true;
}

async function handleTracking(replyToken) {
  const messages = await buildTrackingMessages();
  await reply(replyToken, messages);
}

async function handleSlipImage(replyToken, userId, messageId) {
  const isDuplicate = await isDuplicateSlipMessageId(messageId);

  if (isDuplicate) {
    await reply(replyToken, [
      {
        type: "text",
        text: "สลิปรายการนี้ถูกส่งเข้าระบบแล้ว กรุณารอแอดมินตรวจสอบ"
      }
    ]);
    return;
  }

  const payment = getPaymentTracking(userId);

  if (!payment) {
    markRecentSlipMessageId(messageId);

    await appendOrderLog({
      type: "payment_slip",
      customerNo: "",
      userId,
      displayName: "",
      itemName: "รับสลิป",
      detail: "ได้รับรูปสลิป แต่ไม่พบ payment context ในระบบ",
      qty: "",
      unitPrice: "",
      total: "",
      spotNumbers: "",
      slipStatus: "ส่งสลิปแล้ว",
      slipUrl: "",
      note: buildNote({
        action: "slip_received_without_context",
        lineMessageId: messageId
      })
    });

    await reply(replyToken, [
      {
        type: "text",
        text: "ได้รับรูปสลิปแล้ว กรุณารอแอดมินตรวจสอบ"
      }
    ]);
    return;
  }

  if (payment.type === "spot_booking" && payment.spotNumbers) {
    try {
      await markSpecificSpotsSlipSent(payment.spotNumbers);
    } catch (err) {
      logError("Mark spot slip status error:", err.message);
    }
  }

  await appendOrderLog({
    type: "payment_slip",
    customerNo: payment.customerNo,
    userId: payment.userId,
    displayName: payment.displayName,
    itemName: payment.itemName,
    detail: payment.detail,
    qty: payment.qty,
    unitPrice: payment.unitPrice,
    total: payment.total,
    spotNumbers: payment.spotNumbers,
    slipStatus: "ส่งสลิปแล้ว",
    slipUrl: "",
    note: buildNote({
      action: "slip_received",
      paymentRef: payment.paymentRef,
      lineMessageId: messageId
    })
  });

  markRecentSlipMessageId(messageId);
  await loadRecentSlipMessageIds(true);
  clearPaymentTracking(userId);

  await reply(replyToken, [
    {
      type: "text",
      text: `📩 ได้รับรูปสลิปแล้ว
ชื่อ: ${payment.displayName}
รหัสลูกค้า: ${payment.customerNo}

เลขอ้างอิง: ${payment.paymentRef}

กรุณารอแอดมินตรวจสอบ`
    }
  ]);
}

// =========================================================
// ADMIN DASHBOARD
// =========================================================
app.get("/", (req, res) => {
  res.status(200).json({
    ok: true,
    service: "line-webhook-final",
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
      bookingEnabled: hasGoogleConfig,
      priceItems: Object.keys(prices).length,
      stockModels: Object.keys(stock).length,
      totalSpots: spots.length,
      availableSpots: getAvailableSpots(spots).length,
      totalCustomers: customers.length,
      sessionStates: sessionStates.size,
      paymentTracking: paymentTracking.size,
      campaign
    });
  } catch (err) {
    logError("Health check error:", err.message);
    res.status(500).json({
      ok: false,
      error: err.message
    });
  }
});

app.get("/admin", async (req, res) => {
  if (!verifyAdmin(req)) {
    return res.status(401).send("Unauthorized");
  }

  try {
    const summary = await getAdminSummary();
    res.setHeader("Content-Type", "text/html; charset=utf-8");
    res.status(200).send(renderAdminHtml(summary));
  } catch (err) {
    logError("Admin page error:", err.message);
    res.status(500).send("Admin dashboard error");
  }
});

app.get("/admin/api/summary", async (req, res) => {
  if (!verifyAdmin(req)) {
    return res.status(401).json({ ok: false, error: "Unauthorized" });
  }

  try {
    const summary = await getAdminSummary();
    res.status(200).json({ ok: true, summary });
  } catch (err) {
    logError("Admin api error:", err.message);
    res.status(500).json({ ok: false, error: err.message });
  }
});

// =========================================================
// WEBHOOK
// =========================================================
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
          await handleSlipImage(replyToken, userId, event.message.id || "");
          continue;
        }

        if (event.message.type !== "text") continue;

        const text = String(event.message.text || "").trim();
        logInfo("Incoming text:", text, "userId:", userId || "(empty)");

        const currentState = getSessionState(userId);

        if (currentState?.mode === "address_create") {
          await handleAddressInput(replyToken, userId, text, "address_create");
          continue;
        }

        if (currentState?.mode === "address_update") {
          await handleAddressInput(replyToken, userId, text, "address_update");
          continue;
        }

        if (currentState?.mode === "spot_booking") {
          await handleSpotBookingSelection(replyToken, userId, text);
          continue;
        }

        if (currentState?.mode === "spot_confirm") {
          await handleSpotBookingConfirm(replyToken, userId, text);
          continue;
        }

        if (text === "เมนู" || /^help$/i.test(text)) {
          await reply(replyToken, [{ type: "text", text: getMenuText() }]);
          continue;
        }

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

        if (text === "แก้ไขที่อยู่จัดส่ง") {
          setSessionState(userId, "address_update");
          await reply(replyToken, [
            {
              type: "text",
              text: `กรุณาก็อปปี้เทมเพลตด้านล่าง แล้วเพิ่มข้อมูลจัดส่งใหม่ของคุณให้ครบถ้วน

${getAddressTemplateText()}`
            }
          ]);
          continue;
        }

        if (text === "เช็คพัสดุ") {
          await handleTracking(replyToken);
          continue;
        }

        const handled = await handleOrder(replyToken, userId, text);
        if (handled) continue;

        // ไม่ตอบ fallback เพื่อให้ลูกค้าคุยกับร้านแบบปกติได้
      } catch (eventErr) {
        logError("Event handling error:", eventErr.message);

        try {
          await appendOrderLog({
            type: "system_error",
            customerNo: "",
            userId: event.source?.userId || "",
            displayName: "",
            itemName: "event_error",
            detail: event.message?.text || event.message?.type || "",
            qty: "",
            unitPrice: "",
            total: "",
            spotNumbers: "",
            slipStatus: "",
            slipUrl: "",
            note: buildNote({
              error: eventErr.message,
              messageType: event.message?.type || ""
            })
          });
        } catch (auditErr) {
          logError("Audit error failed:", auditErr.message);
        }

        try {
          await reply(event.replyToken, [
            {
              type: "text",
              text: "ระบบเกิดข้อผิดพลาดชั่วคราว กรุณาลองใหม่อีกครั้ง"
            }
          ]);
        } catch (replyErr) {
          logError("Fallback reply failed:", replyErr.message);
        }
      }
    }

    res.sendStatus(200);
  } catch (err) {
    logError("Webhook error:", err.message);
    res.sendStatus(500);
  }
});

// =========================================================
// START
// =========================================================
app.listen(CONFIG.PORT, () => {
  logInfo(`✅ Server running on port ${CONFIG.PORT}`);
  logInfo("PRICE_SHEET_NAME =", CONFIG.PRICE_SHEET_NAME);
  logInfo("STOCK_SHEET_NAME =", CONFIG.STOCK_SHEET_NAME);
  logInfo("SPOT_SHEET_NAME =", CONFIG.SPOT_SHEET_NAME);
  logInfo("CAMPAIGN_SHEET_NAME =", CONFIG.CAMPAIGN_SHEET_NAME);
  logInfo("LOG_SHEET_NAME =", CONFIG.LOG_SHEET_NAME);
  logInfo("CUSTOMER_SHEET_NAME =", CONFIG.CUSTOMER_SHEET_NAME);
  logInfo("TRACKING_SHEET_NAME =", CONFIG.TRACKING_SHEET_NAME);
  logInfo("ADMIN enabled =", !!CONFIG.ADMIN_KEY);
});