// ======================================================
// ProMEL AI Backend - Railway Ready (ES Modules + Secure)
// ======================================================

import express from "express";
import cors from "cors";
import dotenv from "dotenv";
import OpenAI from "openai";

// 1. Load environment variables (from Railway or local .env)
dotenv.config();

console.log("DEBUG: ProMEL AI backend starting…");
console.log("DEBUG: OPENAI_API_KEY present?", !!process.env.OPENAI_API_KEY);

// IMPORTANT:
// - Do NOT hard-code your key here.
// - Set OPENAI_API_KEY in Railway → Variables.
// - If developing locally, create a .env file with:
//   OPENAI_API_KEY=sk-...

const openai = new OpenAI({
  apiKey: process.env.OPENAI_API_KEY,
});

const app = express();
const PORT = process.env.PORT || 5000;

// 2. Middleware
app.use(cors());
app.use(express.json());

// =============================
// Helper: Build system prompt
// =============================
function buildSystemPrompt(mode) {
  switch (mode) {
    case "explain_trends":
      return `
You are ProMEL AI, a Monitoring, Evaluation and Learning (MEL) assistant for development projects in Papua New Guinea.
Explain current trends, risks, and performance in clear language for managers and decision makers.
Focus on: activity implementation, budget utilisation, staff, community participation, coordination, outcomes and impact.
Use the KPI snapshot and filters as your evidence base. Keep the tone practical and concise.`;
    case "draft_report":
      return `
You are a senior MEL specialist writing a short narrative Monitoring & Evaluation report.
Structure your response with clear headings such as:
1. Project Overview
2. Key Achievements this Period
3. Challenges / Risks
4. Quantitative Performance (explain scores and beneficiary numbers in words)
5. Recommendations and Next Steps
Write in a professional style suitable for government, donors and NGOs in PNG.`;
    case "learning_summary":
      return `
You are facilitating a "Lessons Learned" reflection session for a project team.
Summarise:
- Key successes
- What worked well and why
- Key problems / failures
- What should be done differently next time
- Practical recommendations for adaptation and improvement
Use bullet points and short paragraphs.`;
    case "design_questions":
      return `
You design practical Monitoring & Evaluation questions and indicators for Google Forms.
Produce rating-based questions (1–5 scale) and some open questions covering:
- Activity implementation
- Outputs
- Outcomes
- Impact
- Risks / assumptions
- Sustainability and capacity development.
Make questions easy to understand for field staff in PNG.`;
    default:
      return `
You are ProMEL AI, a Monitoring, Evaluation and Learning (MEL) assistant for development projects in Papua New Guinea.
You help users design logframes, M&E questionnaires, progress reports, summaries, and learning notes.
Be clear, practical, and context-aware for PNG public service and NGOs.`;
  }
}

// =============================
// Helper: Build context text
// =============================
function buildContextText({ include_filters, filters, kpi_snapshot, history }) {
  let context = "";

  // Filter context (Project, Period, Location)
  if (include_filters) {
    const project = filters?.project || "All projects";
    const period = filters?.period || "All periods";
    const location = filters?.location || "All locations";

    context += "FILTER CONTEXT:\n";
    context += `- Project: ${project}\n`;
    context += `- Reporting Period: ${period}\n`;
    context += `- Location: ${location}\n\n`;
  }

  // KPI snapshot context (from dashboard)
  if (kpi_snapshot && typeof kpi_snapshot === "object") {
    context += "KPI SNAPSHOT (Dashboard):\n";
    context += `- Total Projects: ${kpi_snapshot.totalProjects || "0"}\n`;
    context += `- Monitoring Records: ${kpi_snapshot.totalMonitoring || "0"}\n`;
    context += `- Evaluation Records: ${kpi_snapshot.totalEvaluation || "0"}\n`;
    context += `- Total Beneficiaries: ${kpi_snapshot.totalBeneficiaries || "0"}\n`;
    context += `- Overall Score (Activity & Budget): ${kpi_snapshot.overallScore || "0"}\n\n`;
  }

  // Optional: short conversation history (if provided)
  if (Array.isArray(history) && history.length > 0) {
    context += "CONVERSATION HISTORY (previous exchanges):\n";
    history.forEach((m, i) => {
      const role = m.role || "user";
      const content = (m.content || "").slice(0, 400);
      context += `${i + 1}. [${role}] ${content}\n`;
    });
    context += "\n";
  }

  return context;
}

// =============================
// Routes
// =============================

// Health check
app.get("/", (req, res) => {
  res.send("ProMEL AI backend is running ✅");
});

// Test route to confirm correct server file is running
app.get("/ping-promel", (req, res) => {
  console.log("DEBUG: /ping-promel route hit");
  res.json({
    ok: true,
    message: "PING from the NEW ProMEL backend 🟢",
    version: "Railway-ESM",
  });
});

// Main AI chat endpoint
app.post("/api/pas-ai-chat", async (req, res) => {
  try {
    console.log("DEBUG: /api/pas-ai-chat called with body:", req.body);

    // Accept both old and new payload styles
    const {
      mode = "custom",
      user_prompt,
      message, // old field
      include_filters = true,
      filters = {},
      kpi_snapshot = {},
      conversation_id = null,
      history = [], // optional [{role, content}, ...]
    } = req.body || {};

    // Check OpenAI key
    if (!process.env.OPENAI_API_KEY) {
      console.error("ERROR: OPENAI_API_KEY is not set.");
      return res.status(500).json({
        success: false,
        error: "Server misconfigured: OPENAI_API_KEY is missing.",
      });
    }

    // Decide which text to use as user main question
    const rawUserText =
      (user_prompt && String(user_prompt).trim()) ||
      (message && String(message).trim()) ||
      "Explain what Monitoring is in Monitoring & Evaluation (M&E).";

    const systemPrompt = buildSystemPrompt(mode);
    const contextText = buildContextText({
      include_filters,
      filters,
      kpi_snapshot,
      history,
    });

    const finalUserContent =
      contextText +
      "USER QUESTION / TASK:\n" +
      rawUserText;

    console.log("DEBUG: Mode =", mode);
    console.log("DEBUG: Final user content length =", finalUserContent.length);

    // =============================
    // Call OpenAI Chat Completions
    // =============================
    const completion = await openai.chat.completions.create({
      model: "gpt-4o-mini",
      temperature: 0.4,
      messages: [
        {
          role: "system",
          content: systemPrompt,
        },
        {
          role: "user",
          content: finalUserContent,
        },
      ],
    });

    const reply =
      completion.choices?.[0]?.message?.content ||
      "No reply generated by OpenAI.";

    console.log("DEBUG: OpenAI reply received, length =", reply.length);

    // Send response to the dashboard
    return res.json({
      success: true,
      reply,
      conversation_id, // echo back if you want to manage it on the frontend
    });
  } catch (err) {
    console.error("ProMEL AI backend error (exception):", err);
    return res.status(500).json({
      success: false,
      error: "ProMEL AI backend error (exception)",
      message: err.message || String(err),
    });
  }
});

// 5. Start server
app.listen(PORT, () => {
  console.log(`ProMEL AI backend listening on port ${PORT}`);
});
