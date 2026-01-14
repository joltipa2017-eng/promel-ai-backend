// server.js
// ProMEL OpenAI Backend – ES Module version (because "type": "module" in package.json)

import dotenv from "dotenv";
import express from "express";
import cors from "cors";

// ✅ Load .env explicitly
dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

// =====================================================
// STEP 1: Google Sheets CSV URLs (same links as dashboard)
// =====================================================
const MONITORING_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vQKiND_eb3gPR0spQHucadf14xH_rX5gRtpHShan_OWJqSThbHrcx8tqCa4V_3nXjqq3duum0i6XCSE/pub?gid=1298321526&single=true&output=csv";

const EVALUATION_CSV_URL =
  "https://docs.google.com/spreadsheets/d/e/2PACX-1vTd_zFcgy_c_5JDgK9wTVFLi5WeHTPF4uQBq9cOPl8vurc2DVu3DWrqwXiE1FmGcL5f2TtWjWLlGs1L/pub?gid=1608876672&single=true&output=csv";

// =====================================================
// Middleware
// =====================================================

// ✅ More practical CORS for GitHub Pages + local dev + Railway
const allowedOrigins = [
  "https://joltipa2017-eng.github.io",
  "http://localhost:5500",
  "http://127.0.0.1:5500",
  "http://localhost:8080",
  "http://127.0.0.1:8080",
];

// Optional: allow extra origins via env (comma-separated)
const extraOrigins = (process.env.ALLOWED_ORIGINS || "")
  .split(",")
  .map((s) => s.trim())
  .filter(Boolean);

const finalAllowedOrigins = [...new Set([...allowedOrigins, ...extraOrigins])];

app.use(
  cors({
    origin: (origin, cb) => {
      if (!origin) return cb(null, true);
      if (finalAllowedOrigins.includes(origin)) return cb(null, true);

      // keep permissive; change to cb(null, false) for strict lock-down
      return cb(null, true);
    },
    methods: ["GET", "POST", "OPTIONS"],
    allowedHeaders: ["Content-Type", "Authorization"],
  })
);

// Support bigger payloads safely (AI reports can be long)
app.use(express.json({ limit: "2mb" }));

// Simple health check
app.get("/", (req, res) => {
  res.send("ProMEL OpenAI backend is running ✅");
});

// =====================================================
// STEP 2: Helpers — fetch CSV + parse CSV (robust)
// =====================================================
async function fetchCsvText(url) {
  const res = await fetch(url);
  if (!res.ok) throw new Error(`Failed to fetch CSV (${res.status}) ${url}`);
  return await res.text();
}

// Robust CSV parser (handles commas inside quotes)
function parseCSV(text) {
  const rows = [];
  let row = [];
  let current = "";
  let inQuotes = false;

  const s = (text || "").replace(/\r\n/g, "\n").replace(/\r/g, "\n");

  for (let i = 0; i < s.length; i++) {
    const ch = s[i];
    const next = s[i + 1];

    if (ch === '"' && inQuotes && next === '"') {
      current += '"';
      i++;
      continue;
    }

    if (ch === '"') {
      inQuotes = !inQuotes;
      continue;
    }

    if (ch === "," && !inQuotes) {
      row.push(current);
      current = "";
      continue;
    }

    if (ch === "\n" && !inQuotes) {
      row.push(current);
      rows.push(row);
      row = [];
      current = "";
      continue;
    }

    current += ch;
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  return rows.map((r) => r.map((c) => (c === undefined ? "" : String(c))));
}

function safeLower(s) {
  return (s || "").trim().toLowerCase();
}

function findIndex(headersLower, exactName) {
  return headersLower.indexOf(exactName);
}

// =====================================================
// STEP 3: Filter + summarise functions (AI-friendly)
// =====================================================
function filterMonitoringRows(rows, filters) {
  if (!rows || rows.length <= 1) return { header: [], data: [] };

  const header = rows[0];
  const headersLower = header.map(safeLower);

  const idxProject = findIndex(headersLower, "project name");
  const idxPeriod = findIndex(headersLower, "reporting period");
  const idxLocation = findIndex(headersLower, "province / district / location");

  const project = (filters?.project || "").trim();
  const period = (filters?.period || "").trim();
  const location = (filters?.location || "").trim();

  const data = rows
    .slice(1)
    .filter((r) => !r.every((c) => String(c || "").trim() === ""));

  const filtered = data.filter((r) => {
    const p = idxProject !== -1 ? String(r[idxProject] || "").trim() : "";
    const per = idxPeriod !== -1 ? String(r[idxPeriod] || "").trim() : "";
    const loc = idxLocation !== -1 ? String(r[idxLocation] || "").trim() : "";

    const matchProj = project ? p === project : true;
    const matchPer = period ? per === period : true;
    const matchLoc = location ? loc === location : true;

    return matchProj && matchPer && matchLoc;
  });

  return { header, data: filtered };
}

function filterEvaluationRows(rows, filters) {
  if (!rows || rows.length <= 1) return { header: [], data: [] };

  const header = rows[0];
  const headersLower = header.map(safeLower);

  const idxProject = findIndex(headersLower, "project name");
  const idxPeriod = findIndex(headersLower, "reporting period");

  const project = (filters?.project || "").trim();
  const period = (filters?.period || "").trim();

  const data = rows
    .slice(1)
    .filter((r) => !r.every((c) => String(c || "").trim() === ""));

  const filtered = data.filter((r) => {
    const p = idxProject !== -1 ? String(r[idxProject] || "").trim() : "";
    const per = idxPeriod !== -1 ? String(r[idxPeriod] || "").trim() : "";

    const matchProj = project ? p === project : true;
    const matchPer = period ? per === period : true;

    return matchProj && matchPer;
  });

  return { header, data: filtered };
}

function averageNumericColumn(header, data, colNameLower) {
  const headersLower = header.map(safeLower);
  const idx = headersLower.indexOf(colNameLower);
  if (idx === -1) return 0;

  let sum = 0;
  let count = 0;

  data.forEach((r) => {
    const v = parseFloat(r[idx]);
    if (!isNaN(v)) {
      sum += v;
      count++;
    }
  });

  return count ? sum / count : 0;
}

function percentFromRating(avgRating, max = 5) {
  if (!avgRating || isNaN(avgRating)) return 0;
  const pct = (avgRating / max) * 100;
  return Math.max(0, Math.min(100, pct));
}

function summariseMonitoring(header, data) {
  if (!header.length) return "No monitoring sheet loaded.";
  if (!data.length) return "No matching monitoring records for current filters.";

  const headersLower = header.map(safeLower);

  const avgAct = averageNumericColumn(header, data, "activity implementation on schedule");
  const avgBud = averageNumericColumn(header, data, "budget utilization as planned");
  const avgStaff = averageNumericColumn(header, data, "staff availability");
  const avgComm = averageNumericColumn(header, data, "community participation");
  const avgCoord = averageNumericColumn(header, data, "coordination with partners");

  const idxDate = headersLower.indexOf("reporting date");
  const idxBenef = headersLower.indexOf("number of beneficiaries reached this period.");
  const idxAch = headersLower.indexOf("main achievements this period.");
  const idxChal = headersLower.indexOf("key challenges / risks");
  const idxActn = headersLower.indexOf("immediate actions required / recommendations");

  let totalBenef = 0;
  data.forEach((r) => {
    if (idxBenef !== -1) {
      const b = parseFloat(r[idxBenef] || "0");
      if (!isNaN(b)) totalBenef += b;
    }
  });

  const latest = data.slice(-5).reverse();
  const latestLines = latest
    .map((r, i) => {
      const dt = idxDate !== -1 ? r[idxDate] || "" : "";
      const ach = idxAch !== -1 ? r[idxAch] || "" : "";
      const chal = idxChal !== -1 ? r[idxChal] || "" : "";
      const actn = idxActn !== -1 ? r[idxActn] || "" : "";
      return `${i + 1}) Date: ${dt}
   - Achievement: ${ach}
   - Challenge: ${chal}
   - Action: ${actn}`;
    })
    .join("\n");

  return `Records: ${data.length}
Total beneficiaries (filtered): ${totalBenef}

Average ratings (1–5):
- Activity: ${avgAct.toFixed(2)}
- Budget: ${avgBud.toFixed(2)}
- Staff: ${avgStaff.toFixed(2)}
- Community: ${avgComm.toFixed(2)}
- Coordination: ${avgCoord.toFixed(2)}

Most recent entries (up to 5):
${latestLines}`;
}

function summariseEvaluation(header, data) {
  if (!header.length) return "No evaluation sheet loaded.";
  if (!data.length) return "No matching evaluation records for current filters.";

  const headersLower = header.map(safeLower);

  const idxOutcome = headersLower.findIndex((h) => h.includes("outcome") && h.includes("rating"));
  const idxImpact = headersLower.findIndex((h) => h.includes("impact") && h.includes("rating"));

  let idxPerf = headersLower.findIndex(
    (h) => (h.includes("overall") || h.includes("performance")) && h.includes("rating")
  );
  if (idxPerf === -1) idxPerf = headersLower.findIndex((h) => h.includes("rating"));

  const avgOutcome = idxOutcome !== -1 ? averageNumericColumn(header, data, headersLower[idxOutcome]) : 0;
  const avgImpact = idxImpact !== -1 ? averageNumericColumn(header, data, headersLower[idxImpact]) : 0;
  const avgPerf = idxPerf !== -1 ? averageNumericColumn(header, data, headersLower[idxPerf]) : 0;

  const phaseIdx = headersLower.findIndex((h) => h.includes("phase"));

  const latest = data.slice(-5).reverse();
  const latestLines = latest
    .map((r, i) => {
      const phase = phaseIdx !== -1 ? r[phaseIdx] || "" : "";
      const out = idxOutcome !== -1 ? r[idxOutcome] || "" : "N/A";
      const imp = idxImpact !== -1 ? r[idxImpact] || "" : "N/A";
      const perf = idxPerf !== -1 ? r[idxPerf] || "" : "N/A";
      return `${i + 1}) Phase: ${phase} | Outcome: ${out} | Impact: ${imp} | Performance: ${perf}`;
    })
    .join("\n");

  return `Records: ${data.length}

Average ratings (1–5):
- Outcome: ${avgOutcome.toFixed(2)}
- Impact: ${avgImpact.toFixed(2)}
- Performance: ${avgPerf.toFixed(2)}

Most recent entries (up to 5):
${latestLines}`;
}

// =====================================================
// Visuals builder (backend computed, reliable)
// =====================================================
function computeVisualsFromMonitoring(header, data) {
  if (!header?.length || !data?.length) {
    return { kpi_scores: [], distribution: { good: 0, watch: 0, poor: 0 } };
  }

  const headersLower = header.map(safeLower);

  const idxAct = headersLower.indexOf("activity implementation on schedule");
  const idxBud = headersLower.indexOf("budget utilization as planned");
  const idxStaff = headersLower.indexOf("staff availability");
  const idxComm = headersLower.indexOf("community participation");
  const idxCoord = headersLower.indexOf("coordination with partners");

  const avgAct = averageNumericColumn(header, data, "activity implementation on schedule");
  const avgBud = averageNumericColumn(header, data, "budget utilization as planned");
  const avgStaff = averageNumericColumn(header, data, "staff availability");
  const avgComm = averageNumericColumn(header, data, "community participation");
  const avgCoord = averageNumericColumn(header, data, "coordination with partners");

  const pctAct = Math.round(percentFromRating(avgAct));
  const pctBud = Math.round(percentFromRating(avgBud));
  const pctStaff = Math.round(percentFromRating(avgStaff));
  const pctComm = Math.round(percentFromRating(avgComm));
  const pctCoord = Math.round(percentFromRating(avgCoord));

  const available = [pctAct, pctBud, pctStaff, pctComm, pctCoord].filter((p) => p > 0);
  const pctOverall = available.length
    ? Math.round(available.reduce((a, x) => a + x, 0) / available.length)
    : 0;

  let good = 0, watch = 0, poor = 0;

  data.forEach((r) => {
    const vals = [];
    [idxAct, idxBud, idxStaff, idxComm, idxCoord].forEach((idx) => {
      if (idx !== -1) {
        const v = parseFloat(r[idx] || "");
        if (!isNaN(v)) vals.push(v);
      }
    });
    if (!vals.length) return;
    const rowAvg = vals.reduce((a, x) => a + x, 0) / vals.length;

    if (rowAvg >= 4) good++;
    else if (rowAvg <= 2) poor++;
    else watch++;
  });

  return {
    kpi_scores: [
      { label: "Activity", percent: pctAct },
      { label: "Budget", percent: pctBud },
      { label: "Staff", percent: pctStaff },
      { label: "Community", percent: pctComm },
      { label: "Coordination", percent: pctCoord },
      { label: "Overall", percent: pctOverall },
    ],
    distribution: { good, watch, poor },
  };
}

function computeVisualsFromEvaluation(header, data) {
  if (!header?.length || !data?.length) {
    return { kpi_scores: [], distribution: { good: 0, watch: 0, poor: 0 } };
  }

  const headersLower = header.map(safeLower);
  const idxOutcome = headersLower.findIndex((h) => h.includes("outcome") && h.includes("rating"));
  const idxImpact = headersLower.findIndex((h) => h.includes("impact") && h.includes("rating"));

  let idxPerf = headersLower.findIndex(
    (h) => (h.includes("overall") || h.includes("performance")) && h.includes("rating")
  );
  if (idxPerf === -1) idxPerf = headersLower.findIndex((h) => h.includes("rating"));

  const avgOutcome = idxOutcome !== -1 ? averageNumericColumn(header, data, headersLower[idxOutcome]) : 0;
  const avgImpact = idxImpact !== -1 ? averageNumericColumn(header, data, headersLower[idxImpact]) : 0;
  const avgPerf = idxPerf !== -1 ? averageNumericColumn(header, data, headersLower[idxPerf]) : 0;

  const pctOutcome = Math.round(percentFromRating(avgOutcome));
  const pctImpact = Math.round(percentFromRating(avgImpact));
  const pctPerf = Math.round(percentFromRating(avgPerf));

  let good = 0, watch = 0, poor = 0;

  data.forEach((r) => {
    let perfVal = NaN;
    if (idxPerf !== -1) perfVal = parseFloat(r[idxPerf] || "");

    if (isNaN(perfVal)) {
      const vals = [];
      if (idxOutcome !== -1) {
        const v = parseFloat(r[idxOutcome] || "");
        if (!isNaN(v)) vals.push(v);
      }
      if (idxImpact !== -1) {
        const v = parseFloat(r[idxImpact] || "");
        if (!isNaN(v)) vals.push(v);
      }
      if (vals.length) perfVal = vals.reduce((a, x) => a + x, 0) / vals.length;
    }

    if (isNaN(perfVal)) return;

    if (perfVal >= 4) good++;
    else if (perfVal <= 2) poor++;
    else watch++;
  });

  return {
    kpi_scores: [
      { label: "Outcome", percent: pctOutcome },
      { label: "Impact", percent: pctImpact },
      { label: "Performance", percent: pctPerf },
    ],
    distribution: { good, watch, poor },
  };
}

function mergeVisuals(monVisuals, evalVisuals) {
  const combinedDist = {
    good: (monVisuals?.distribution?.good || 0) + (evalVisuals?.distribution?.good || 0),
    watch: (monVisuals?.distribution?.watch || 0) + (evalVisuals?.distribution?.watch || 0),
    poor: (monVisuals?.distribution?.poor || 0) + (evalVisuals?.distribution?.poor || 0),
  };

  const allPercents = []
    .concat((monVisuals?.kpi_scores || []).map((x) => x.percent))
    .concat((evalVisuals?.kpi_scores || []).map((x) => x.percent))
    .filter((p) => typeof p === "number" && !isNaN(p) && p > 0);

  const combinedScore = allPercents.length
    ? Math.round(allPercents.reduce((a, x) => a + x, 0) / allPercents.length)
    : 0;

  return {
    combined_score_percent: combinedScore,
    combined_distribution: combinedDist,
  };
}

// =====================================================
// Safe parse JSON from AI response
// =====================================================
function safeParseAiJson(text) {
  if (!text) return null;

  let t = String(text).trim();
  t = t.replace(/^```json\s*/i, "");
  t = t.replace(/^```\s*/i, "");
  t = t.replace(/```$/i, "").trim();

  try {
    return JSON.parse(t);
  } catch {
    const match = t.match(/\{[\s\S]*\}$/);
    if (match) {
      try {
        return JSON.parse(match[0]);
      } catch {
        return null;
      }
    }
    return null;
  }
}

// =====================================================
// Extract output text robustly (IMPORTANT FIX)
// =====================================================
function extractResponseText(data) {
  if (!data) return "";
  if (typeof data.output_text === "string" && data.output_text.trim()) return data.output_text;

  // Fallback: try to reconstruct from output[].content[]
  const chunks = [];
  const out = Array.isArray(data.output) ? data.output : [];
  for (const item of out) {
    const content = Array.isArray(item?.content) ? item.content : [];
    for (const c of content) {
      if (typeof c?.text === "string") chunks.push(c.text);
      // Some variants use nested objects; keep it tolerant
      if (typeof c?.content === "string") chunks.push(c.content);
    }
  }
  return chunks.join("\n").trim();
}

// =====================================================
// Export helpers (Word .doc download without extra libs)
// =====================================================
function escapeHtml(str) {
  return String(str || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function markdownToHtml(md) {
  const t = String(md || "").replace(/\r\n/g, "\n");
  let html = escapeHtml(t);

  html = html
    .replace(/^######\s?(.*)$/gm, "<h6>$1</h6>")
    .replace(/^#####\s?(.*)$/gm, "<h5>$1</h5>")
    .replace(/^####\s?(.*)$/gm, "<h4>$1</h4>")
    .replace(/^###\s?(.*)$/gm, "<h3>$1</h3>")
    .replace(/^##\s?(.*)$/gm, "<h2>$1</h2>")
    .replace(/^#\s?(.*)$/gm, "<h1>$1</h1>");

  html = html.replace(/\*\*(.+?)\*\*/g, "<strong>$1</strong>");
  html = html.replace(/\*(.+?)\*/g, "<em>$1</em>");

  html = html.replace(/^\s*-\s+(.*)$/gm, "<li>$1</li>");
  html = html.replace(/((?:\s*<li>[\s\S]*?<\/li>\s*)+)/g, "<ul>$1</ul>");

  html = html
    .replace(/\n{2,}/g, "\n\n")
    .split("\n\n")
    .map((chunk) => {
      const c = chunk.trim();
      if (!c) return "";
      if (c.startsWith("<h") || c.startsWith("<ul>")) return c;
      return `<p>${c.replace(/\n/g, "<br/>")}</p>`;
    })
    .filter(Boolean)
    .join("\n");

  return html;
}

// =====================================================
// Export endpoint - Word .doc
// =====================================================
app.post("/api/export/word", (req, res) => {
  try {
    const { title, content, filters } = req.body || {};
    const safeTitle = (title || "ProMEL_AI_Report").toString().trim() || "ProMEL_AI_Report";

    if (!content) {
      return res.status(400).json({ success: false, error: 'Missing "content" to export.' });
    }

    const filterLine = `Project: ${filters?.project || "All"} | Period: ${filters?.period || "All"} | Location: ${
      filters?.location || "All"
    }`;
    const generated = `Generated: ${new Date().toLocaleString()}`;

    const htmlBody = markdownToHtml(content);

    const docHtml = `
      <html>
        <head>
          <meta charset="utf-8" />
          <title>${escapeHtml(safeTitle)}</title>
          <style>
            body { font-family: Arial, sans-serif; font-size: 11pt; }
            h1 { font-size: 16pt; color: #003366; }
            h2 { font-size: 14pt; color: #003366; }
            h3 { font-size: 12pt; color: #003366; }
            .meta { font-size: 10pt; color: #444; margin-bottom: 12px; }
            ul { margin-top: 6px; }
          </style>
        </head>
        <body>
          <h1>${escapeHtml(safeTitle)}</h1>
          <div class="meta">${escapeHtml(filterLine)}<br>${escapeHtml(generated)}</div>
          ${htmlBody}
        </body>
      </html>
    `.trim();

    res.setHeader("Content-Type", "application/msword; charset=utf-8");
    res.setHeader(
      "Content-Disposition",
      `attachment; filename="${safeTitle.replace(/[^a-z0-9_\-]+/gi, "_")}.doc"`
    );
    return res.send(docHtml);
  } catch (err) {
    console.error("Export error /api/export/word:", err);
    return res.status(500).json({ success: false, error: "Failed to export Word document." });
  }
});

// =====================================================
// MEL Intent Router
// =====================================================
function detectMELIntent(message) {
  const m = String(message || "").toLowerCase().trim();

  if (m.startsWith("what is") || m.startsWith("define") || m.includes("meaning of")) return "DEFINITION";

  if (
    m.includes("draft") ||
    m.includes("create") ||
    m.includes("develop") ||
    m.includes("design") ||
    m.includes("template") ||
    m.includes("tool") ||
    m.includes("instrument") ||
    m.includes("survey") ||
    m.includes("questionnaire") ||
    m.includes("interview questions") ||
    m.includes("evaluation questions") ||
    m.includes("indicators") ||
    m.includes("kpi") ||
    m.includes("logframe") ||
    m.includes("logical framework") ||
    m.includes("theory of change") ||
    m.includes("toc") ||
    m.includes("sampling") ||
    m.includes("methodology") ||
    m.includes("tor") ||
    m.includes("terms of reference") ||
    m.includes("data collection")
  ) {
    return "TOOLS";
  }

  if (m.includes("report") || m.includes("evaluation report") || m.includes("mel report")) return "REPORT";

  if (
    m.includes("dashboard") ||
    m.includes("trend") ||
    m.includes("performance") ||
    m.includes("interpret") ||
    m.includes("explain the data") ||
    m.includes("why is") ||
    m.includes("why are") ||
    m.includes("what does the data say")
  ) {
    return "DASHBOARD_ANALYSIS";
  }

  if (m.includes("lessons learned") || m.includes("after action") || m.includes("learning agenda"))
    return "LEARNING";

  return "HOW_TO";
}

function getIntentOutputRules(intent) {
  switch (intent) {
    case "DEFINITION":
      return `
INTENT: DEFINITION
- Give a clear definition (2–4 sentences).
- Why it matters in MEL (2–4 bullets).
- One short PNG/public-service example.
- Do NOT generate a report. Do NOT refer to dashboard visuals/charts.
`.trim();

    case "TOOLS":
      return `
INTENT: TOOLS
- Produce the requested tool/template in a structured way.
- If user asks for questions: group by theme + suggest response scales.
- If user asks for indicators: include indicator, definition, data source, frequency, disaggregation.
- Do NOT produce a full evaluation report unless explicitly requested.
- Do NOT describe charts/dashboards.
`.trim();

    case "DASHBOARD_ANALYSIS":
      return `
INTENT: DASHBOARD_ANALYSIS
- Use the provided LIVE summaries (if provided).
- Output TEXT ONLY (no chart descriptions).
- Structure:
  1) What the data suggests (3–6 bullets)
  2) Likely causes (hypotheses)
  3) Recommended management actions (SMART)
`.trim();

    case "REPORT":
      return `
INTENT: REPORT
- Generate a narrative report (TEXT ONLY) with headings:
  1) Executive Summary
  2) Purpose & Scope
  3) Methods (state assumptions if missing)
  4) Findings
  5) Conclusions
  6) Recommendations
- Do NOT describe charts; do NOT replicate dashboard cards.
`.trim();

    case "LEARNING":
      return `
INTENT: LEARNING
- Provide learning questions, reflection prompts, and practical actions.
- Keep it usable for teams (bullets, short sections).
- Do NOT generate a report unless asked.
`.trim();

    default:
      return `
INTENT: HOW_TO
- Provide practical step-by-step guidance.
- Use short headings and bullets.
- Do NOT generate a report unless asked.
`.trim();
  }
}

// =====================================================
// Main AI endpoint for your dashboard card
// =====================================================
app.post("/api/pas-ai-chat", async (req, res) => {
  try {
    const { message, filters, intent_override } = req.body || {};

    if (!message) {
      return res.status(400).json({ success: false, error: 'Missing "message" in request body' });
    }

    if (!process.env.OPENAI_API_KEY) {
      return res.status(500).json({ success: false, error: "OPENAI_API_KEY is not set on the server." });
    }

    const [monCsvText, evalCsvText] = await Promise.all([
      fetchCsvText(MONITORING_CSV_URL),
      fetchCsvText(EVALUATION_CSV_URL),
    ]);

    const monRows = parseCSV(monCsvText);
    const evalRows = parseCSV(evalCsvText);

    const monFiltered = filterMonitoringRows(monRows, filters);
    const evalFiltered = filterEvaluationRows(evalRows, filters);

    const monitoringSummary = summariseMonitoring(monFiltered.header, monFiltered.data);
    const evaluationSummary = summariseEvaluation(evalFiltered.header, evalFiltered.data);

    const monVisuals = computeVisualsFromMonitoring(monFiltered.header, monFiltered.data);
    const evalVisuals = computeVisualsFromEvaluation(evalFiltered.header, evalFiltered.data);
    const combinedVisuals = mergeVisuals(monVisuals, evalVisuals);

    const filterText = `
Dashboard Context:
- Location: ${filters?.location || "All"}
- Reporting period: ${filters?.period || "All"}
- Project: ${filters?.project || "All"}
`.trim();

    const intent = (intent_override && String(intent_override).trim()) || detectMELIntent(message);
    const intentRules = getIntentOutputRules(intent);

    const includeLiveDataInPrompt = intent === "REPORT" || intent === "DASHBOARD_ANALYSIS";

    const SYSTEM_PROMPT = `
You are ProMEL AI, a Monitoring, Evaluation & Learning (MEL) assistant for Papua New Guinea projects.

CRITICAL BEHAVIOR:
- Answer ONLY what the user asked for. Do NOT generate a report unless the user asks for a report.
- TEXT OUTPUT ONLY: do not describe charts/graphs/cards; do not replicate dashboard visuals.
- If LIVE summaries are provided, use them ONLY when the intent is REPORT or DASHBOARD_ANALYSIS.
- If LIVE summaries are not provided, answer as general MEL guidance.

JSON OUTPUT REQUIREMENTS:
- Always return STRICT JSON following the required schema.
- key_findings must contain 3–7 items (never empty).
- recommendations must contain 3–7 items (never empty).
`.trim();

    const JSON_SCHEMA = {
      name: "promel_ai_response",
      strict: true,
      schema: {
        type: "object",
        additionalProperties: false,
        properties: {
          report_title: { type: "string" },
          report_markdown: { type: "string" },
          key_findings: { type: "array", items: { type: "string" } },
          recommendations: { type: "array", items: { type: "string" } },
          visuals: {
            type: "object",
            additionalProperties: false,
            properties: {
              kpi_scores: {
                type: "array",
                items: {
                  type: "object",
                  additionalProperties: false,
                  properties: {
                    label: { type: "string" },
                    percent: { type: "number" },
                  },
                  required: ["label", "percent"],
                },
              },
              distribution: {
                type: "object",
                additionalProperties: false,
                properties: {
                  good: { type: "number" },
                  watch: { type: "number" },
                  poor: { type: "number" },
                },
                required: ["good", "watch", "poor"],
              },
              combined_score_percent: { type: "number" },
              combined_distribution: {
                type: "object",
                additionalProperties: false,
                properties: {
                  good: { type: "number" },
                  watch: { type: "number" },
                  poor: { type: "number" },
                },
                required: ["good", "watch", "poor"],
              },
            },
            required: ["kpi_scores", "distribution", "combined_score_percent", "combined_distribution"],
          },
        },
        required: ["report_title", "report_markdown", "key_findings", "recommendations", "visuals"],
      },
    };

    const liveDataBlock = includeLiveDataInPrompt
      ? `
${filterText}

LIVE MONITORING SUMMARY:
${monitoringSummary}

LIVE EVALUATION SUMMARY:
${evaluationSummary}
`.trim()
      : `
NOTE: LIVE summaries are intentionally NOT provided for this request.
Answer as general MEL guidance (do not assume project performance).
`.trim();

    const visualsBlock = `
VISUALS DATA (USE EXACTLY AS GIVEN; DO NOT INVENT):
- Monitoring KPI scores: ${JSON.stringify(monVisuals.kpi_scores)}
- Monitoring distribution: ${JSON.stringify(monVisuals.distribution)}
- Evaluation KPI scores: ${JSON.stringify(evalVisuals.kpi_scores)}
- Evaluation distribution: ${JSON.stringify(evalVisuals.distribution)}
- Combined score percent: ${JSON.stringify(combinedVisuals.combined_score_percent)}
- Combined distribution: ${JSON.stringify(combinedVisuals.combined_distribution)}
`.trim();

    const userPrompt = `
${intentRules}

${liveDataBlock}

${visualsBlock}

USER REQUEST:
${message}

Output rules:
- Put the main answer in report_markdown (use clean headings/bullets).
- key_findings: 3–7 items, never empty.
- recommendations: 3–7 items, never empty.
- visuals must reflect the numbers provided above.
`.trim();

    const MODEL = process.env.OPENAI_MODEL || "gpt-4o-mini";

    const openaiResponse = await fetch("https://api.openai.com/v1/responses", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${process.env.OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: MODEL,
        temperature: 0.2,
        input: [
          { role: "system", content: SYSTEM_PROMPT },
          { role: "user", content: userPrompt },
        ],

        // ✅ FIX: Correct Structured Outputs format for Responses API
        text: {
          format: {
            type: "json_schema",
            name: JSON_SCHEMA.name,
            strict: JSON_SCHEMA.strict,
            schema: JSON_SCHEMA.schema,
          },
        },
      }),
    });

    if (!openaiResponse.ok) {
      let errorBody = null;
      try {
        errorBody = await openaiResponse.json();
      } catch {
        const text = await openaiResponse.text();
        errorBody = { raw: text };
      }

      console.error("OpenAI API error:", openaiResponse.status, JSON.stringify(errorBody, null, 2));

      return res.status(openaiResponse.status).json({
        success: false,
        error: errorBody?.error?.message || `OpenAI API error (status ${openaiResponse.status})`,
        status: openaiResponse.status,
      });
    }

    const data = await openaiResponse.json();

    // ✅ FIX: Extract robustly
    const rawText = extractResponseText(data);

    const aiJson = safeParseAiJson(rawText);

    if (!aiJson || typeof aiJson !== "object") {
      return res.json({
        success: true,
        reply: rawText || "AI returned no usable JSON. Showing raw response.",
        visuals: {
          kpi_scores: monVisuals.kpi_scores,
          distribution: monVisuals.distribution,
          combined_score_percent: combinedVisuals.combined_score_percent,
          combined_distribution: combinedVisuals.combined_distribution,
        },
        used_filters: filters || {},
        monitoring_records_used: monFiltered.data?.length || 0,
        evaluation_records_used: evalFiltered.data?.length || 0,
        note: "AI response was not valid JSON; returned raw text + backend visuals.",
      });
    }

    const safeTitle =
      (aiJson.report_title && String(aiJson.report_title).trim()) || "ProMEL MEL Output";

    const safeFindings =
      Array.isArray(aiJson.key_findings) && aiJson.key_findings.length
        ? aiJson.key_findings
        : [
            "Clear MEL concepts and tools improve decision-making and accountability throughout the project cycle.",
            "Using fit-for-purpose indicators and data collection methods increases the reliability of findings and learning.",
            "Regular reflection and adaptive action planning strengthens outcomes and sustainability over time.",
          ];

    const safeRecs =
      Array.isArray(aiJson.recommendations) && aiJson.recommendations.length
        ? aiJson.recommendations
        : [
            "Use an intent-based approach: definitions, tools, analysis, and reports should be produced only when requested.",
            "Standardize templates (indicators, tools, learning notes) so teams apply MEL consistently across projects.",
            "Institutionalize periodic learning reviews and management action tracking to close feedback loops.",
          ];

    const replyTextRaw = (aiJson.report_markdown && String(aiJson.report_markdown).trim()) || "";
    const replyText =
      replyTextRaw ||
      `# ${safeTitle}\n\n` +
        `## Key Findings\n- ${safeFindings.join("\n- ")}\n\n` +
        `## Recommendations\n- ${safeRecs.join("\n- ")}\n`;

    return res.json({
      success: true,
      reply: replyText,
      report_title: safeTitle,
      key_findings: safeFindings,
      recommendations: safeRecs,
      visuals: {
        kpi_scores: monVisuals.kpi_scores,
        distribution: monVisuals.distribution,
        combined_score_percent: combinedVisuals.combined_score_percent,
        combined_distribution: combinedVisuals.combined_distribution,
      },
      used_filters: filters || {},
      monitoring_records_used: monFiltered.data?.length || 0,
      evaluation_records_used: evalFiltered.data?.length || 0,
      detected_intent: intent,
    });
  } catch (err) {
    console.error("Backend error in /api/pas-ai-chat:", err);
    return res.status(500).json({ success: false, error: "Server error in /api/pas-ai-chat" });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`ProMEL OpenAI backend listening on http://localhost:${PORT}`);
});
