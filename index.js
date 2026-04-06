import express from "express";
import fetch from "node-fetch";

const app = express();
app.use(express.json());

const TOKEN = "9RZmKVgzTnr75by2V6nzHyxxZsaIqt0h1v9FZ4OA8haa6fHrOLpJ/ocPI8PIQb3lxF2yTJo1Z3pWZOLtoX/kfa6c8ce5L/zwddp4420nRe+Al8bsVXFjjm3lkp17IGPIhQ/KRn61rl5bGxiv7pnvRgdB04t89/1O/w1cDnyilFU=";

app.get("/", (req, res) => {
  res.status(200).send("Server is running");
});

app.get("/webhook", (req, res) => {
  res
    .status(200)
    .send("Webhook endpoint is alive. LINE must use POST, not GET.");
});

app.post("/webhook", async (req, res) => {
  try {
    console.log("Webhook hit");
    console.log(JSON.stringify(req.body));

    const events = req.body.events || [];

    for (const e of events) {
      if (e.type !== "message") continue;
      if (!e.message || e.message.type !== "text") continue;

      const replyPayload = {
        replyToken: e.replyToken,
        messages: [
          {
            type: "text",
            text: "คุณพิมพ์ว่า: " + e.message.text
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