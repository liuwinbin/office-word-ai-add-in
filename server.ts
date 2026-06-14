import express from "express";
import path from "path";
import fs from "fs";
import https from "https";

const app = express();
const PORT = 3000;
const CERT_DIR = path.join(process.cwd(), '.certs');
const PFX_PATH = path.join(CERT_DIR, 'localhost.pfx');
const PFX_PASSPHRASE = 'officeai123';

console.log("=== OfficeAI Server Starting ===");

app.use(express.json());

// ── Serve Word add-in taskpane static files ──────────────────
app.get('/taskpane.html', (_req, res) => {
  res.sendFile(path.join(process.cwd(), 'taskpane.html'));
});
app.get('/taskpane.js', (_req, res) => {
  res.sendFile(path.join(process.cwd(), 'taskpane.js'));
});
app.use('/assets', express.static(path.join(process.cwd(), 'assets')));

// ── API Health Check ─────────────────────────────────────────
app.get("/api/health", (_req, res) => {
  res.json({
    status: "ok",
    https: true,
    hasDeepseekKey: !!process.env.DEEPSEEK_API_KEY
  });
});

// ── Chat Proxy for DeepSeek API ──────────────────────────────
app.post("/api/chat", async (req, res) => {
  try {
    const { messages, model, temperature, apiKey, customModelName, customBaseUrl } = req.body;
    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: "Messages array is required." });
    }
    const dsKey = apiKey || process.env.DEEPSEEK_API_KEY;
    if (!dsKey || dsKey.trim() === "") {
      return res.status(401).json({ error: "API key not configured.", needKey: true });
    }
    let targetUrl = "https://api.deepseek.com/chat/completions";
    let dsModel = "deepseek-chat";
    if (model === "custom-model") {
      dsModel = customModelName?.trim() || "deepseek-chat";
      if (customBaseUrl?.trim()) {
        let base = customBaseUrl.trim();
        if (base.endsWith('/')) base = base.slice(0, -1);
        targetUrl = base.includes('/chat/completions') ? base : base + '/chat/completions';
      }
    } else {
      dsModel = model === "deepseek-reasoner" ? "deepseek-reasoner" : "deepseek-chat";
    }
    const body: any = { model: dsModel, messages };
    if (!dsModel.includes("reasoner") && !dsModel.includes("r1")) {
      body.temperature = temperature ?? 0.7;
    }
    const dsRes = await fetch(targetUrl, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": "Bearer " + dsKey.trim()
      },
      body: JSON.stringify(body)
    });
    if (!dsRes.ok) {
      const errText = await dsRes.text();
      return res.status(dsRes.status).json({ error: errText.slice(0, 300) });
    }
    return res.json(await dsRes.json());
  } catch (error: any) {
    console.error("API error:", error);
    res.status(500).json({ error: error.message || "Internal server error" });
  }
});

// ── Start HTTPS server ───────────────────────────────────────
if (fs.existsSync(PFX_PATH)) {
  const pfx = fs.readFileSync(PFX_PATH);
  https.createServer({ pfx, passphrase: PFX_PASSPHRASE }, app).listen(PORT, "0.0.0.0", () => {
    console.log("HTTPS server running at https://localhost:" + PORT);
    console.log("Taskpane:  https://localhost:" + PORT + "/taskpane.html");
  });
} else {
  console.error("ERROR: PFX certificate not found at " + PFX_PATH);
  process.exit(1);
}