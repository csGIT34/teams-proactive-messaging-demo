// Bare-minimum proactive Teams messaging test.
// No Bot Framework SDK — raw REST so every moving part is visible.
//
// Flow:
//  1. Teams POSTs activities to /api/messages when a user installs or messages the bot.
//  2. We store the "conversation reference" (serviceUrl + conversation id) per user.
//  3. POST /api/send uses that reference to push a message to the user at any time.

require("dotenv").config();
const express = require("express");
const fs = require("fs");
const path = require("path");

const { BOT_APP_ID, BOT_APP_SECRET, TENANT_ID, PORT = 3000 } = process.env;
const STORE_FILE = path.join(__dirname, "conversations.json");

const app = express();
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));

// ---- conversation reference store (JSON file so restarts don't lose it) ----

function loadStore() {
  try {
    return JSON.parse(fs.readFileSync(STORE_FILE, "utf8"));
  } catch {
    return {};
  }
}

function saveStore(store) {
  fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2));
}

// ---- bot access token (client credentials) ----

async function getBotToken() {
  // Single-tenant bots authenticate against your tenant; multi-tenant bots use botframework.com.
  const authority = TENANT_ID || "botframework.com";
  const res = await fetch(
    `https://login.microsoftonline.com/${authority}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: new URLSearchParams({
        grant_type: "client_credentials",
        client_id: BOT_APP_ID,
        client_secret: BOT_APP_SECRET,
        scope: "https://api.botframework.com/.default",
      }),
    }
  );
  const body = await res.json();
  if (!res.ok) throw new Error(`Token request failed: ${JSON.stringify(body)}`);
  return body.access_token;
}

// ---- send an activity into an existing conversation ----

async function sendToConversation(ref, text) {
  const token = await getBotToken();
  const url = `${ref.serviceUrl.replace(/\/$/, "")}/v3/conversations/${encodeURIComponent(ref.conversationId)}/activities`;
  const res = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      type: "message",
      from: ref.bot,
      conversation: { id: ref.conversationId },
      text,
    }),
  });
  const body = await res.text();
  if (!res.ok) throw new Error(`Send failed (${res.status}): ${body}`);
  return JSON.parse(body);
}

// ---- bot messaging endpoint (Teams calls this) ----
// NOTE: a real bot must validate the JWT in the Authorization header
// (issuer api.botframework.com). Skipped here to keep the test minimal.

app.post("/api/messages", async (req, res) => {
  const activity = req.body;
  console.log(`Incoming activity: ${activity.type} from ${activity.from?.name || "unknown"}`);

  // Any activity (installation, conversationUpdate, message) carries the
  // conversation reference we need for proactive sends. Capture it.
  if (activity.conversation && activity.from?.aadObjectId) {
    const store = loadStore();
    store[activity.from.aadObjectId] = {
      userName: activity.from.name,
      userId: activity.from.id,
      conversationId: activity.conversation.id,
      serviceUrl: activity.serviceUrl,
      bot: activity.recipient, // the bot's channel account, used as "from" on sends
      capturedAt: new Date().toISOString(),
    };
    saveStore(store);
    console.log(`Stored conversation reference for ${activity.from.name}`);
  }

  res.status(200).end();

  // Echo a confirmation when the user messages the bot, so the round trip is visible.
  if (activity.type === "message") {
    try {
      const store = loadStore();
      const ref = store[activity.from.aadObjectId];
      if (ref) await sendToConversation(ref, `Got it — I can now message you proactively. You said: "${activity.text}"`);
    } catch (err) {
      console.error("Echo reply failed:", err.message);
    }
  }
});

// ---- web UI API ----

app.get("/api/conversations", (req, res) => {
  res.json(loadStore());
});

app.post("/api/send", async (req, res) => {
  const { aadObjectId, text } = req.body;
  const ref = loadStore()[aadObjectId];
  if (!ref) return res.status(404).json({ error: "No conversation reference for that user" });
  if (!text) return res.status(400).json({ error: "text is required" });
  try {
    const result = await sendToConversation(ref, text);
    res.json({ ok: true, activityId: result.id });
  } catch (err) {
    console.error("Proactive send failed:", err.message);
    res.status(500).json({ error: err.message });
  }
});

app.listen(PORT, () => {
  console.log(`Listening on http://localhost:${PORT}`);
  console.log(`Bot endpoint: POST /api/messages (must be publicly reachable for Teams)`);
});
