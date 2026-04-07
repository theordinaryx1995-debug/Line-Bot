import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";

// ลิงก์รูป QR / รูปชำระเงิน ต้องเป็น public https
const PAYMENT_IMAGE_URL = "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg";

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
  let result = "สรุปรายการ\n";

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

  result += "\nรวมทั้งหมด = " + total + " บาท";
  return result;
}

function buildPaymentText() {
  return `สามารถชำระเงินผ่านช่องทางอื่น ๆ ได้ดังนี้

ชื่อบัญชี ปรัชญา สุดใจดี

K-Bank 0503228092

True Wallet 0982652650

✨ ชำระแล้วโปรดแนบ Pay Slip การโอนทุกครั้ง ✨

ขอบคุณนะครับ`;
}

async function buildPriceMenu() {
  const priceTable = await loadPrices();

  const pack = [];
  const box = [];

  const codes = Object.keys(priceTable).sort();

  for (const code of codes) {
    const price = priceTable[code];

    // box code ขึ้นต้นด้วย B เช่น BOP-13, BPRB-01, BEB-04
    if (code.startsWith("B")) {
      const displayCode = code.substring(1); // ตัด B ออกตอนแสดง
      box.push(`${displayCode} ${price}`);
    } else {
      pack.push(`${code} ${price}`);
    }
  }

  let result = "✨ ซอง ✨\n";
  result += pack.join("\n");
  result += "\n\n✨ Box ✨\n";
  result += box.join("\n");

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

      // 1) ถ้าพิมพ์ ราคา / เมนูราคา → ส่งตารางราคา
      if (
        userText === "ราคา" ||
        userText === "เมนูราคา" ||
        userText === "ราคาสินค้า"
      ) {
        const menuText = await buildPriceMenu();

        const replyPayload = {
          replyToken: e.replyToken,
          messages: [
            {
              type: "text",
              text: menuText
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

        console.log("Price menu reply status:", response.status);
        console.log("Price menu reply body:", await response.text());
        continue;
      }

      // 2) ถ้าพิมพ์ รวมราคา ... → คำนวณยอด + ส่ง QR + ข้อความชำระเงิน
      if (userText.includes("รวมราคา")) {
        const cleanText = userText.replace("รวมราคา", "").trim();
        const summaryText = await calculate(cleanText);

        const paymentText = buildPaymentText();

        const replyPayload = {
          replyToken: e.replyToken,
          messages: [
            {
              type: "text",
              text: summaryText
            },
            {
              type: "image",
              originalContentUrl: PAYMENT_IMAGE_URL,
              previewImageUrl: PAYMENT_IMAGE_URL
            },
            {
              type: "text",
              text: paymentText
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

        console.log("Summary reply status:", response.status);
        console.log("Summary reply body:", await response.text());
        continue;
      }

      // ถ้าไม่เข้าเงื่อนไขไหน ไม่ต้องตอบ
      continue;
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