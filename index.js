import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";
// Sheet 1 = Price list
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";
// Sheet 2 = stock carton
const STOCK_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=262793173";


// ลิงก์รูป QR / รูปชำระเงิน ต้องเป็น public https
const PAYMENT_IMAGE_URL = "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg";


// =========================
// ROUTES
// =========================
app.get("/", (req, res) => {
  res.status(200).send("Server is running");
});

app.get("/webhook", (req, res) => {
  res.status(200).send("Webhook OK");
});

// =========================
// HELPERS
// =========================
function formatBaht(num) {
  return Number(num).toLocaleString("en-US");
}

function normalizeCode(code) {
  return code.toUpperCase().replace(/[^A-Z0-9]/g, "");
}

function displayCode(code) {
  const m = code.match(/^([A-Z]+)(\d+)$/);
  return m ? `${m[1]}-${m[2]}` : code;
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

// =========================
// LOAD PRICE CSV (Sheet1)
// A = Code
// B = Pack_price
// C = Box_Price
// =========================
async function loadPrices() {
  const res = await fetch(PRICE_CSV_URL);
  const csv = await res.text();

  const lines = csv.trim().split("\n");
  const table = {};

  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(",");
    if (row.length < 3) continue;

    const code = row[0].trim().replace(/"/g, "").toUpperCase();
    const pack = Number(row[1].trim().replace(/"/g, ""));
    const box = Number(row[2].trim().replace(/"/g, ""));

    if (!code) continue;

    table[code] = {
      pack: Number.isNaN(pack) ? null : pack,
      box: Number.isNaN(box) ? null : box
    };
  }

  console.log("PRICE TABLE:", table);
  return table;
}

// =========================
// LOAD STOCK CSV (Sheet2)
// A = รุ่น
// B = หมวด
// C = เหลือ
// =========================
async function loadStock() {
  const res = await fetch(STOCK_CSV_URL);
  const csv = await res.text();

  const lines = csv.trim().split("\n");
  const stock = {};

  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(",");
    if (row.length < 3) continue;

    const model = row[0].trim().replace(/"/g, "").toUpperCase();
    const category = row[1].trim().replace(/"/g, "").toUpperCase();
    const qty = Number(row[2].trim().replace(/"/g, ""));

    if (!model || !category) continue;

    if (!stock[model]) stock[model] = {};
    stock[model][category] = Number.isNaN(qty) ? 0 : qty;
  }

  console.log("STOCK TABLE:", stock);
  return stock;
}

// =========================
// BUILD PRICE LIST
// =========================
function buildPriceList(table) {
  let txt = "📋 ราคาสินค้า\n";

  for (const code in table) {
    const item = table[code];
    txt += `${displayCode(code)} | ซอง ${formatBaht(item.pack)} บาท | กล่อง ${formatBaht(item.box)} บาท\n`;
  }

  return txt.trim();
}

// =========================
// FORMAT STOCK
// =========================
function formatStock(model, data) {
  const lines = [`📦 ${model} เหลือใน carton`];

  for (const key of ["SP", "SEC", "LPA", "DON"]) {
    if (data[key] !== undefined) {
      const iText = "I ".repeat(data[key]).trim();
      lines.push(`${key}: ${iText || "-"}`);
    }
  }

  return lines.join("\n");
}

function formatAllStock(stock) {
  const messages = ["📦 สถานะสินค้าใน carton"];

  for (const model in stock) {
    const data = stock[model];
    messages.push(`\n${model}`);

    for (const key of ["SP", "SEC", "LPA", "DON"]) {
      if (data[key] !== undefined) {
        const iText = "I ".repeat(data[key]).trim();
        messages.push(`${key}: ${iText || "-"}`);
      }
    }
  }

  return messages.join("\n");
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
  const regex = /([A-Z]+-?\d+)\s*(?:x?\s*)?(\d+)\s*(ซอง|ซ็อง|pack|box|กล่อง|บ็อก)/gi;
  const items = [];

  let m;
  while ((m = regex.exec(text))) {
    items.push({
      code: normalizeCode(m[1]),
      qty: Number(m[2]),
      unit:
        m[3].includes("ซอง") || m[3] === "pack" || m[3] === "ซ็อง"
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
  let clean = text.trim();

  if (clean === "รวมราคา") {
    return {
      status: "guide",
      message: getOrderGuideText()
    };
  }

  clean = clean.replace(/^รวมราคา/i, "").trim();

  const items = parseItems(clean);
  console.log("PARSED ITEMS:", items);

  if (items.length === 0) {
    return null;
  }

  const table = await loadPrices();

  let total = 0;
  let lines = ["🧾 สรุปรายการ"];
  let valid = 0;

  for (const i of items) {
    const p = table[i.code];

    if (!p) {
      lines.push(`${displayCode(i.code)} ❌ ไม่มีสินค้า`);
      continue;
    }

    const price = p[i.unit];

    if (price == null) {
      lines.push(`${displayCode(i.code)} ❌ ไม่มีราคาประเภท${displayUnit(i.unit)}`);
      continue;
    }

    const sum = price * i.qty;
    total += sum;
    valid++;

    lines.push(
      `${displayCode(i.code)} ${displayUnit(i.unit)}ละ ${formatBaht(price)} x${i.qty} = ${formatBaht(sum)} บาท`
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
// REPLY TO LINE
// =========================
async function reply(token, messages) {
  const response = await fetch("https://api.line.me/v2/bot/message/reply", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: "Bearer " + TOKEN
    },
    body: JSON.stringify({
      replyToken: token,
      messages
    })
  });

  const result = await response.text();
  console.log("Reply status:", response.status);
  console.log("Reply body:", result);
}

// =========================
// WEBHOOK
// =========================
app.post("/webhook", async (req, res) => {
  try {
    const events = req.body.events || [];

    for (const e of events) {
      if (e.type !== "message") continue;
      if (!e.message || e.message.type !== "text") continue;

      const text = e.message.text.trim();

      // ราคาสินค้า
      if (text === "ราคาสินค้า") {
        const table = await loadPrices();

        await reply(e.replyToken, [
          { type: "text", text: buildPriceList(table) },
          { type: "text", text: getOrderGuideText() }
        ]);
        continue;
      }

      // Check rate ทั้งหมด หรือรายรุ่น
      if (text.toLowerCase().startsWith("check rate")) {
        const parts = text.split(" ");
        const modelRaw = parts[2];
        const stock = await loadStock();

        if (!modelRaw) {
          await reply(e.replyToken, [
            { type: "text", text: formatAllStock(stock) }
          ]);
          continue;
        }

        const model = modelRaw.toUpperCase().replace("-", "");

        if (!stock[model]) {
          await reply(e.replyToken, [
            { type: "text", text: "ไม่พบข้อมูลรุ่นนี้" }
          ]);
          continue;
        }

        await reply(e.replyToken, [
          { type: "text", text: formatStock(model, stock[model]) }
        ]);
        continue;
      }

      // รวมราคา / หรือพิมพ์รายการสินค้าเลย
      const result = await calculate(text);

      if (!result) continue;

      if (result.status === "guide") {
        await reply(e.replyToken, [
          { type: "text", text: result.message }
        ]);
        continue;
      }

      if (result.status === "invalid") {
        await reply(e.replyToken, [
          { type: "text", text: result.message }
        ]);
        continue;
      }

      // SUCCESS
      await reply(e.replyToken, [
        { type: "text", text: result.summary },
        { type: "text", text: result.payment },
        {
          type: "image",
          originalContentUrl: PAYMENT_IMAGE_URL,
          previewImageUrl: PAYMENT_IMAGE_URL
        },
        {
          type: "text",
          text:
`สามารถชำระเงินผ่านช่องทางอื่น ๆ ได้ดังนี้

ชื่อบัญชี ปรัชญา สุดใจดี

K-Bank 0503228092

True Wallet 0982652650

✨ ชำระแล้วโปรดแปะ Pay Slip การโอนทุกครั้ง ✨

ขอบคุณนะครับ`
        }
      ]);
    }

    res.sendStatus(200);
  } catch (err) {
    console.error("Webhook error:", err);
    res.sendStatus(500);
  }
});

// =========================
// START
// =========================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Running on " + PORT));