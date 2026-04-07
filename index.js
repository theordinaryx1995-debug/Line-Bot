import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";

app.get("/", (req, res) => {
  res.status(200).send("Server is running");
});

app.get("/webhook", (req, res) => {
  res.status(200).send("Webhook endpoint is alive. Use POST from LINE only.");
});

async function loadPrices() {
  const response = await fetch(SHEET_CSV_URL);
  const csv = await response.text();

  const lines = csv.trim().split("\n");
  const priceTable = {};

  for (let i = 1; i < lines.length; i++) {
    const row = lines[i].split(",");
    if (row.length < 2) continue;

    const code = row[0].trim().replace(/"/g, "").toUpperCase();
    const price = Number(row[1].trim().replace(/"/g, ""));

    if (!code || Number.isNaN(price)) continue;
    priceTable[code] = price;
  }

  return priceTable;
}

function normalizeCode(raw) {
  if (!raw) return "";
  return raw.trim().toUpperCase();
}

function parseItems(text) {
  const parts = text.split("/");
  const items = [];

  for (let part of parts) {
    const cleaned = part.trim();

    // รองรับรูปแบบ เช่น:
    // OP-13 2
    // OP-13 x2
    // OP-13 = 2
    // prb-01 1
    const match = cleaned.match(/([A-Z]+-?\d+)[^\d]*(\d+)/i);
    if (!match) continue;

    const code = normalizeCode(match[1]);
    const qty = parseInt(match[2], 10);

    if (!code || Number.isNaN(qty) || qty <= 0) continue;

    items.push({ code, qty });
  }

  return items;
}

async function calculate(text) {
  const priceTable = await loadPrices();
  const items = parseItems(text);

  if (items.length === 0) {
    return "พิมพ์เช่น รวมราคา OP-13 2 / OP-14 1";
  }

  let total = 0;
  let result = "🧾 สรุปรายการ\n";

  for (const item of items) {
    if (!priceTable[item.code]) {
      result += `${item.code} ❌ ไม่มีสินค้า\n`;
      continue;
    }

    const subtotal = priceTable[item.code] * item.qty;
    total += subtotal;

    result += `${item.code} x${item.qty} = ${subtotal} บาท\n`;
  }

  if (total === 0) {
    return "ไม่พบสินค้าที่คำนวณได้";
  }

  result += "-----------------\n";
  result += `รวมทั้งหมด = ${total} บาท`;

  return result;
}

app.post("/webhook", async (req, res) => {
  try {
    console.log("Webhook hit");
    console.log(JSON.stringify(req.body));

    const events = req.body.events || [];

    for (const e of events) {
      if (e.type !== "message") continue;
      if (!e.message || e.message.type !== "text") continue;

      const userText = e.message.text.trim().toUpperCase();

      // ตอบเฉพาะข้อความที่มีคำว่า "รวมราคา"
      if (!userText.includes("รวมราคา")) {
        continue;
      }

      // ลบคำ trigger ออกก่อนคำนวณ
      const cleanText = userText.replace("รวมราคา", "").trim();

      const replyText = await calculate(cleanText);

      const replyPayload = {
        replyToken: e.replyToken,
        messages: [
          {
            type: "text",
            text: replyText
          }
        ]
      };

      const response = await fetch("https://api.line.me/v2/bot/message/reply", {
        method: "POST",
        headers: {
          "Content-Type": "application/json",
          Authorization: "Bearer " + TOKEN
        },
        body: JSON.stringify(replyPayload)
      });

      const resultText = await response.text();
      console.log("Reply status:", response.status);
      console.log("Reply body:", resultText);
    }

    return res.sendStatus(200);
  } catch (err) {
    console.error("Webhook error:", err);
    return res.sendStatus(500);
  }
});

const PORT = process.env.PORT || 3000;

app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});