import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";

// ลิงก์รูป QR / รูปชำระเงิน ต้องเป็น public https
const PAYMENT_IMAGE_URL = "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg";


// =========================
// BASIC ROUTES
// =========================
app.get("/", (req, res) => {
  res.status(200).send("Server is running");
});

app.get("/webhook", (req, res) => {
  res.status(200).send("Webhook endpoint is alive");
});

// =========================
// HELPERS
// =========================
function formatBaht(num) {
  return Number(num).toLocaleString("en-US");
}

function normalizeProductCode(raw) {
  return raw.toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function displayProductCode(code) {
  const match = code.match(/^([A-Z]+)(\d+)$/);
  if (!match) return code;
  return `${match[1]}-${match[2]}`;
}

function displayUnit(unit) {
  if (unit === "pack") return "ซอง";
  if (unit === "box") return "กล่อง";
  return "";
}

// =========================
// LOAD PRICES FROM SHEET CSV
// คอลัมน์:
// A = Code
// B = Pack_price
// C = Box_Price
// =========================
async function loadPrices() {
  const response = await fetch(SHEET_CSV_URL);
  const csv = await response.text();

  console.log("CSV RAW:", csv);

  const lines = csv.trim().split("\n");
  const priceTable = {};

  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(",");
    if (row.length < 3) continue;

    const code = row[0].trim().replace(/"/g, "").toUpperCase();
    const packPrice = Number(row[1].trim().replace(/"/g, ""));
    const boxPrice = Number(row[2].trim().replace(/"/g, ""));

    if (!code) continue;

    priceTable[code] = {
      pack: Number.isNaN(packPrice) ? null : packPrice,
      box: Number.isNaN(boxPrice) ? null : boxPrice
    };
  }

  console.log("PRICE TABLE:", priceTable);
  return priceTable;
}

// =========================
// BUILD PRICE LIST
// สำหรับคำสั่ง "ราคาสินค้า"
// =========================
function buildPriceList(priceTable) {
  const lines = ["📋 ราคาสินค้า"];

  for (const code of Object.keys(priceTable)) {
    const item = priceTable[code];
    const displayCode = displayProductCode(code);

    let row = `${displayCode}`;

    if (item.pack != null) {
      row += ` | ซอง ${formatBaht(item.pack)} บาท`;
    }

    if (item.box != null) {
      row += ` | กล่อง ${formatBaht(item.box)} บาท`;
    }

    lines.push(row);
  }

  return lines.join("\n");
}

// =========================
// PARSE ITEMS
// รองรับ:
// OP13 2 ซอง
// OP-13 2 ซอง
// op13 x2 ซอง
// prb01 1 box
// รวมราคา op13 2 ซอง prb01 1 box ส่งด่วน
// =========================
function parseItems(orderText) {
  const regex = /([A-Z]+-?\d+)\s*(?:x?\s*)?(\d+)\s*(ซอง|ซ็อง|pack|กล่อง|box|บ็อก)/gi;
  const items = [];

  let match;
  while ((match = regex.exec(orderText)) !== null) {
    const rawCode = match[1];
    const qty = parseInt(match[2], 10);
    const rawUnit = match[3].toLowerCase();

    let unit = null;
    if (rawUnit === "ซอง" || rawUnit === "ซ็อง" || rawUnit === "pack") unit = "pack";
    if (rawUnit === "กล่อง" || rawUnit === "box" || rawUnit === "บ็อก") unit = "box";

    if (!unit || !qty || qty <= 0) continue;

    items.push({
      code: normalizeProductCode(rawCode),
      qty,
      unit
    });
  }

  return items;
}

// =========================
// CALCULATE ORDER
// =========================
async function calculateOrder(text) {
  const originalText = text.trim();

  if (!originalText.toLowerCase().startsWith("รวมราคา")) {
    return null;
  }

  const orderText = originalText.replace(/^รวมราคา\s*/i, "").trim();

  if (!orderText) {
    return {
      status: "invalid",
      message: "พิมพ์เช่น รวมราคา OP-13 2 ซอง OP-14 1 box"
    };
  }

  const items = parseItems(orderText);
  console.log("PARSED ITEMS:", items);

  if (items.length === 0) {
    return {
      status: "invalid",
      message: "พิมพ์เช่น รวมราคา OP-13 2 ซอง OP-14 1 box"
    };
  }

  const priceTable = await loadPrices();

  let total = 0;
  let validCount = 0;
  const resultLines = ["🧾 สรุปรายการ"];

  for (const item of items) {
    const product = priceTable[item.code];

    if (!product) {
      resultLines.push(`${displayProductCode(item.code)} ❌ ไม่มีสินค้า`);
      continue;
    }

    const unitPrice = product[item.unit];

    if (unitPrice == null) {
      resultLines.push(`${displayProductCode(item.code)} ❌ ไม่มีราคาประเภท${displayUnit(item.unit)}`);
      continue;
    }

    const subtotal = unitPrice * item.qty;
    total += subtotal;
    validCount++;

    resultLines.push(
      `${displayProductCode(item.code)} ${displayUnit(item.unit)}ละ ${formatBaht(unitPrice)} x${item.qty} = ${formatBaht(subtotal)} บาท`
    );
  }

  if (validCount === 0) {
    return {
      status: "invalid",
      message: "พิมพ์เช่น รวมราคา OP-13 2 ซอง OP-14 1 box"
    };
  }

  resultLines.push("━━━━━━━━━━");
  resultLines.push(`รวมทั้งหมด = ${formatBaht(total)} บาท`);

  return {
    status: "success",
    summaryText: resultLines.join("\n"),
    paymentText: `📌 กรุณาโอนชำระจำนวน ${formatBaht(total)} บาท\nหลังโอนแล้วส่งสลิปในแชตนี้ได้เลย`
  };
}

// =========================
// REPLY TO LINE
// =========================
async function replyLine(replyToken, messages) {
  const response = await fetch("https://api.line.me/v2/bot/message/reply", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + TOKEN
    },
    body: JSON.stringify({
      replyToken,
      messages
    })
  });

  const resultText = await response.text();
  console.log("Reply status:", response.status);
  console.log("Reply body:", resultText);
}

// =========================
// WEBHOOK
// =========================
app.post("/webhook", async (req, res) => {
  try {
    console.log("Webhook hit");
    console.log(JSON.stringify(req.body));

    const events = req.body.events || [];

    for (const e of events) {
      if (e.type !== "message") continue;
      if (!e.message || e.message.type !== "text") continue;

      const userText = e.message.text.trim();

      // คำสั่งดูราคาสินค้า
      if (userText === "ราคาสินค้า") {
        const priceTable = await loadPrices();
        const priceListText = buildPriceList(priceTable);

        await replyLine(e.replyToken, [
          {
            type: "text",
            text: priceListText
          }
        ]);
        continue;
      }

      const orderResult = await calculateOrder(userText);

      // ไม่ใช่คำสั่งรวมราคา ก็ไม่ตอบ
      if (!orderResult) {
        continue;
      }

      // ถ้ารูปแบบไม่ครบ ไม่ส่ง QR
      if (orderResult.status === "invalid") {
        await replyLine(e.replyToken, [
          {
            type: "text",
            text: orderResult.message
          }
        ]);
        continue;
      }

      // ถ้ามีสรุปรายการแล้ว ค่อยส่งข้อความสรุป + ข้อความชำระเงิน + รูป QR
      if (orderResult.status === "success") {
        await replyLine(e.replyToken, [
          {
            type: "text",
            text: orderResult.summaryText
          },
          {
            type: "text",
            text: orderResult.paymentText
          },
          {
            type: "image",
            originalContentUrl: PAYMENT_IMAGE_URL,
            previewImageUrl: PAYMENT_IMAGE_URL
          }
        ]);
      }
    }

    return res.sendStatus(200);
  } catch (err) {
    console.error("Webhook error:", err);
    return res.sendStatus(500);
  }
});

// =========================
// START SERVER
// =========================
const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});