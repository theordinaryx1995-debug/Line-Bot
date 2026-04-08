import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";
const SHEET_CSV_URL = "https://docs.google.com/spreadsheets/d/1BkjMteb8JN1RjOz_CmDoTgNBrA9wCIC1Y3bf4xQWMhc/export?format=csv&gid=0";

// ลิงก์รูป QR / รูปชำระเงิน ต้องเป็น public https
const PAYMENT_IMAGE_URL = "https://raw.githubusercontent.com/theordinaryx1995-debug/Line-Bot/main/image-1824349084924438.jpg";


// =========================
// ROUTE
// =========================
app.get("/", (req, res) => {
  res.send("Server is running");
});

app.get("/webhook", (req, res) => {
  res.send("Webhook OK");
});

// =========================
// UTIL
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
// LOAD CSV
// =========================
async function loadPrices() {
  const res = await fetch(SHEET_CSV_URL);
  const csv = await res.text();

  const lines = csv.trim().split("\n");
  const table = {};

  for (let i = 1; i < lines.length; i++) {
    const [code, pack, box] = lines[i].split(",");

    if (!code) continue;

    table[code.trim().toUpperCase()] = {
      pack: Number(pack),
      box: Number(box)
    };
  }

  console.log("PRICE TABLE:", table);
  return table;
}

// =========================
// SHOW PRICE LIST
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
// PARSE
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
// CALCULATE
// =========================
async function calculate(text) {
  const clean = text.trim();

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

    if (!price) {
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
      message: "พิมพ์เช่น OP13 2 ซอง OP15 2 ซอง หรือ OP13 1 กล่อง"
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
// REPLY
// =========================
async function reply(token, messages) {
  await fetch("https://api.line.me/v2/bot/message/reply", {
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
}

// =========================
// WEBHOOK
// =========================
app.post("/webhook", async (req, res) => {
  try {
    const events = req.body.events || [];

    for (const e of events) {
      if (e.type !== "message") continue;
      if (e.message.type !== "text") continue;

      const text = e.message.text.trim();

      // 1) พิมพ์ "ราคาสินค้า"
      if (text === "ราคาสินค้า") {
        const table = await loadPrices();

        await reply(e.replyToken, [
          { type: "text", text: buildPriceList(table) },
          { type: "text", text: getOrderGuideText() }
        ]);
        continue;
      }

      // 2) พิมพ์ "รวมราคา" อย่างเดียว
      if (text === "รวมราคา") {
        await reply(e.replyToken, [
          { type: "text", text: getOrderGuideText() }
        ]);
        continue;
      }

      // 3) คำนวณอัตโนมัติ ถ้าจับแพทเทิร์นสินค้าได้
      const result = await calculate(text);

      if (!result) continue;

      if (result.status === "invalid") {
        await reply(e.replyToken, [
          { type: "text", text: result.message }
        ]);
        continue;
      }

      // 4) SUCCESS
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
    console.error(err);
    res.sendStatus(500);
  }
});

// =========================
// START
// =========================
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log("Running on " + PORT));