import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "ใส่ token ใหม่";

app.post("/webhook", async (req, res) => {
  const events = req.body.events;

  for (const e of events) {
    if (e.type !== "message" || e.message.type !== "text") continue;

    await fetch("https://api.line.me/v2/bot/message/reply", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: "Bearer " + TOKEN
      },
      body: JSON.stringify({
        replyToken: e.replyToken,
        messages: [
          { type: "text", text: "คุณพิมพ์ว่า: " + e.message.text }
        ]
      })
    });
  }

  res.sendStatus(200);
});

app.listen(3000);