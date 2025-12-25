// server.js
// ProMEL OpenAI Backend – ES Module version (because "type": "module" in package.json)

import dotenv from 'dotenv';
import express from 'express';
import cors from 'cors';

// Load .env explicitly
dotenv.config();

const app = express();
const PORT = process.env.PORT || 5000;

// =====================================================
// STEP 1: Google Sheets CSV URLs (same links as dashboard)
// =====================================================
const MONITORING_CSV_URL =
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vQKiND_eb3gPR0spQHucadf14xH_rX5gRtpHShan_OWJqSThbHrcx8tqCa4V_3nXjqq3duum0i6XCSE/pub?gid=1298321526&single=true&output=csv';

const EVALUATION_CSV_URL =
  'https://docs.google.com/spreadsheets/d/e/2PACX-1vTd_zFcgy_c_5JDgK9wTVFLi5WeHTPF4uQBq9cOPl8vurc2DVu3DWrqwXiE1FmGcL5f2TtWjWLlGs1L/pub?gid=1608876672&single=true&output=csv';

// Middleware
app.use(cors());
app.use(express.json());

// Simple health check
app.get('/', (req, res) => {
  res.send('ProMEL OpenAI backend is running ✅');
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
  let current = '';
  let inQuotes = false;

  const s = (text || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

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

    if (ch === ',' && !inQuotes) {
      row.push(current);
      current = '';
      continue;
    }

    if (ch === '\n' && !inQuotes) {
      row.push(current);
      rows.push(row);
      row = [];
      current = '';
      continue;
    }

    current += ch;
  }

  if (current.length || row.length) {
    row.push(current);
    rows.push(row);
  }

  return rows.map((r) => r.map((c) => (c === undefined ? '' : String(c))));
}

function safeLower(s) {
  return (s || '').trim().toLowerCase();
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

  const idxProject = findIndex(headersLower, 'project name');
  const idxPeriod = findIndex(headersLower, 'reporting period');
  const idxLocation = findIndex(headersLower, 'province / district / location');

  const project = (filters?.project || '').trim();
  const period = (filters?.period || '').trim();
  const location = (filters?.location || '').trim();

  const data = rows
    .slice(1)
    .filter((r) => !r.every((c) => String(c || '').trim() === ''));

  const filtered = data.filter((r) => {
    const p = idxProject !== -1 ? String(r[idxProject] || '').trim() : '';
    const per = idxPeriod !== -1 ? String(r[idxPeriod] || '').trim() : '';
    const loc = idxLocation !== -1 ? String(r[idxLocation] || '').trim() : '';

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

  const idxProject = findIndex(headersLower, 'project name');
  const idxPeriod = findIndex(headersLower, 'reporting period');

  const project = (filters?.project || '').trim();
  const period = (filters?.period || '').trim();

  const data = rows
    .slice(1)
    .filter((r) => !r.every((c) => String(c || '').trim() === ''));

  const filtered = data.filter((r) => {
    const p = idxProject !== -1 ? String(r[idxProject] || '').trim() : '';
    const per = idxPeriod !== -1 ? String(r[idxPeriod] || '').trim() : '';

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

// Keep summaries compact so you don't overload the prompt
function summariseMonitoring(header, data) {
  if (!header.length) return 'No monitoring sheet loaded.';
  if (!data.length) return 'No matching monitoring records for current filters.';

  const headersLower = header.map(safeLower);

  const avgAct = averageNumericColumn(header, data, 'activity implementation on schedule');
  const avgBud = averageNumericColumn(header, data, 'budget utilization as planned');
  const avgStaff = averageNumericColumn(header, data, 'staff availability');
  const avgComm = averageNumericColumn(header, data, 'community participation');
  const avgCoord = averageNumericColumn(header, data, 'coordination with partners');

  const idxDate = headersLower.indexOf('reporting date');
  const idxBenef = headersLower.indexOf('number of beneficiaries reached this period.');
  const idxAch = headersLower.indexOf('main achievements this period.');
  const idxChal = headersLower.indexOf('key challenges / risks');
  const idxActn = headersLower.indexOf('immediate actions required / recommendations');

  let totalBenef = 0;
  data.forEach((r) => {
    if (idxBenef !== -1) {
      const b = parseFloat(r[idxBenef] || '0');
      if (!isNaN(b)) totalBenef += b;
    }
  });

  // last 5 (compact)
  const latest = data.slice(-5).reverse();
  const latestLines = latest
    .map((r, i) => {
      const dt = idxDate !== -1 ? (r[idxDate] || '') : '';
      const ach = idxAch !== -1 ? (r[idxAch] || '') : '';
      const chal = idxChal !== -1 ? (r[idxChal] || '') : '';
      const actn = idxActn !== -1 ? (r[idxActn] || '') : '';
      return `${i + 1}) Date: ${dt}
   - Achievement: ${ach}
   - Challenge: ${chal}
   - Action: ${actn}`;
    })
    .join('\n');

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
  if (!header.length) return 'No evaluation sheet loaded.';
  if (!data.length) return 'No matching evaluation records for current filters.';

  const headersLower = header.map(safeLower);

  const idxOutcome = headersLower.findIndex((h) => h.includes('outcome') && h.includes('rating'));
  const idxImpact = headersLower.findIndex((h) => h.includes('impact') && h.includes('rating'));

  let idxPerf = headersLower.findIndex(
    (h) => (h.includes('overall') || h.includes('performance')) && h.includes('rating')
  );
  if (idxPerf === -1) idxPerf = headersLower.findIndex((h) => h.includes('rating'));

  const avgOutcome = idxOutcome !== -1 ? averageNumericColumn(header, data, headersLower[idxOutcome]) : 0;
  const avgImpact = idxImpact !== -1 ? averageNumericColumn(header, data, headersLower[idxImpact]) : 0;
  const avgPerf = idxPerf !== -1 ? averageNumericColumn(header, data, headersLower[idxPerf]) : 0;

  const phaseIdx = headersLower.findIndex((h) => h.includes('phase'));

  const latest = data.slice(-5).reverse();
  const latestLines = latest
    .map((r, i) => {
      const phase = phaseIdx !== -1 ? (r[phaseIdx] || '') : '';
      const out = idxOutcome !== -1 ? (r[idxOutcome] || '') : 'N/A';
      const imp = idxImpact !== -1 ? (r[idxImpact] || '') : 'N/A';
      const perf = idxPerf !== -1 ? (r[idxPerf] || '') : 'N/A';
      return `${i + 1}) Phase: ${phase} | Outcome: ${out} | Impact: ${imp} | Performance: ${perf}`;
    })
    .join('\n');

  return `Records: ${data.length}

Average ratings (1–5):
- Outcome: ${avgOutcome.toFixed(2)}
- Impact: ${avgImpact.toFixed(2)}
- Performance: ${avgPerf.toFixed(2)}

Most recent entries (up to 5):
${latestLines}`;
}

// =====================================================
// NEW: Visuals builder (backend computed, reliable)
// =====================================================
function computeVisualsFromMonitoring(header, data) {
  if (!header?.length || !data?.length) {
    return {
      kpi_scores: [],
      distribution: { good: 0, watch: 0, poor: 0 },
    };
  }

  const headersLower = header.map(safeLower);

  const idxAct = headersLower.indexOf('activity implementation on schedule');
  const idxBud = headersLower.indexOf('budget utilization as planned');
  const idxStaff = headersLower.indexOf('staff availability');
  const idxComm = headersLower.indexOf('community participation');
  const idxCoord = headersLower.indexOf('coordination with partners');

  const avgAct = averageNumericColumn(header, data, 'activity implementation on schedule');
  const avgBud = averageNumericColumn(header, data, 'budget utilization as planned');
  const avgStaff = averageNumericColumn(header, data, 'staff availability');
  const avgComm = averageNumericColumn(header, data, 'community participation');
  const avgCoord = averageNumericColumn(header, data, 'coordination with partners');

  const pctAct = Math.round(percentFromRating(avgAct));
  const pctBud = Math.round(percentFromRating(avgBud));
  const pctStaff = Math.round(percentFromRating(avgStaff));
  const pctComm = Math.round(percentFromRating(avgComm));
  const pctCoord = Math.round(percentFromRating(avgCoord));

  const available = [pctAct, pctBud, pctStaff, pctComm, pctCoord].filter((p) => p > 0);
  const pctOverall = available.length
    ? Math.round(available.reduce((a, x) => a + x, 0) / available.length)
    : 0;

  // Distribution (row-level average)
  let good = 0, watch = 0, poor = 0;

  data.forEach((r) => {
    const vals = [];
    [idxAct, idxBud, idxStaff, idxComm, idxCoord].forEach((idx) => {
      if (idx !== -1) {
        const v = parseFloat(r[idx] || '');
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
      { label: 'Activity', percent: pctAct },
      { label: 'Budget', percent: pctBud },
      { label: 'Staff', percent: pctStaff },
      { label: 'Community', percent: pctComm },
      { label: 'Coordination', percent: pctCoord },
      { label: 'Overall', percent: pctOverall },
    ],
    distribution: { good, watch, poor },
  };
}

function computeVisualsFromEvaluation(header, data) {
  if (!header?.length || !data?.length) {
    return {
      kpi_scores: [],
      distribution: { good: 0, watch: 0, poor: 0 },
    };
  }

  const headersLower = header.map(safeLower);
  const idxOutcome = headersLower.findIndex((h) => h.includes('outcome') && h.includes('rating'));
  const idxImpact = headersLower.findIndex((h) => h.includes('impact') && h.includes('rating'));

  let idxPerf = headersLower.findIndex(
    (h) => (h.includes('overall') || h.includes('performance')) && h.includes('rating')
  );
  if (idxPerf === -1) idxPerf = headersLower.findIndex((h) => h.includes('rating'));

  const avgOutcome = idxOutcome !== -1 ? averageNumericColumn(header, data, headersLower[idxOutcome]) : 0;
  const avgImpact = idxImpact !== -1 ? averageNumericColumn(header, data, headersLower[idxImpact]) : 0;
  const avgPerf = idxPerf !== -1 ? averageNumericColumn(header, data, headersLower[idxPerf]) : 0;

  const pctOutcome = Math.round(percentFromRating(avgOutcome));
  const pctImpact = Math.round(percentFromRating(avgImpact));
  const pctPerf = Math.round(percentFromRating(avgPerf));

  // Distribution based on performance rating if present
  let good = 0, watch = 0, poor = 0;

  data.forEach((r) => {
    let perfVal = NaN;
    if (idxPerf !== -1) perfVal = parseFloat(r[idxPerf] || '');

    if (isNaN(perfVal)) {
      const vals = [];
      if (idxOutcome !== -1) {
        const v = parseFloat(r[idxOutcome] || '');
        if (!isNaN(v)) vals.push(v);
      }
      if (idxImpact !== -1) {
        const v = parseFloat(r[idxImpact] || '');
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
      { label: 'Outcome', percent: pctOutcome },
      { label: 'Impact', percent: pctImpact },
      { label: 'Performance', percent: pctPerf },
    ],
    distribution: { good, watch, poor },
  };
}

function mergeVisuals(monVisuals, evalVisuals) {
  // Combined distribution and a combined KPI score (simple average of available percents)
  const combinedDist = {
    good: (monVisuals?.distribution?.good || 0) + (evalVisuals?.distribution?.good || 0),
    watch: (monVisuals?.distribution?.watch || 0) + (evalVisuals?.distribution?.watch || 0),
    poor: (monVisuals?.distribution?.poor || 0) + (evalVisuals?.distribution?.poor || 0),
  };

  const allPercents = []
    .concat((monVisuals?.kpi_scores || []).map((x) => x.percent))
    .concat((evalVisuals?.kpi_scores || []).map((x) => x.percent))
    .filter((p) => typeof p === 'number' && !isNaN(p) && p > 0);

  const combinedScore = allPercents.length
    ? Math.round(allPercents.reduce((a, x) => a + x, 0) / allPercents.length)
    : 0;

  return {
    combined_score_percent: combinedScore,
    combined_distribution: combinedDist,
  };
}

// =====================================================
// NEW: Safe parse JSON from AI response (handles ```json)
// =====================================================
function safeParseAiJson(text) {
  if (!text) return null;

  // Remove ```json blocks if present
  let t = String(text).trim();

  // Common cases: ```json\n{...}\n```
  t = t.replace(/^```json\s*/i, '');
  t = t.replace(/^```\s*/i, '');
  t = t.replace(/```$/i, '').trim();

  // Try parse directly
  try {
    return JSON.parse(t);
  } catch {
    // Try to extract first {...} blob
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

// Main AI endpoint for your dashboard card
app.post('/api/pas-ai-chat', async (req, res) => {
  try {
    const { message, filters } = req.body || {};

    if (!message) {
      return res.status(400).json({
        success: false,
        error: 'Missing "message" in request body',
      });
    }

    // =====================================================
    // STEP 4: Fetch live sheet data + filter + summarise
    // =====================================================
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

    // =====================================================
    // NEW: Compute visuals (reliable numeric data)
    // =====================================================
    const monVisuals = computeVisualsFromMonitoring(monFiltered.header, monFiltered.data);
    const evalVisuals = computeVisualsFromEvaluation(evalFiltered.header, evalFiltered.data);
    const combinedVisuals = mergeVisuals(monVisuals, evalVisuals);

    // Dashboard context
    const filterText = `
Dashboard Context:
- Location: ${filters?.location || 'All'}
- Reporting period: ${filters?.period || 'All'}
- Project: ${filters?.project || 'All'}
`.trim();

    // =====================================================
    // NEW: Force JSON-only output from the model
    // =====================================================
    const jsonInstruction = `
Return STRICT JSON only. No markdown fences. No commentary.

The JSON schema must be exactly:
{
  "report_title": "string",
  "report_markdown": "string",
  "key_findings": ["string", "..."],
  "recommendations": ["string", "..."],
  "visuals": {
    "kpi_scores": [{"label":"string","percent":number}, ...],
    "distribution": {"good":number,"watch":number,"poor":number},
    "combined_score_percent": number,
    "combined_distribution": {"good":number,"watch":number,"poor":number}
  }
}

Rules:
- report_markdown must include sections: Overview, Monitoring Summary, Evaluation Summary, Key Findings, Recommendations.
- Base narrative ONLY on the LIVE summaries.
- For visuals.kpi_scores and visuals.distribution, you MUST use the visuals provided below (do not invent).
`.trim();

    const userPrompt = `
You are ProMEL AI, a Monitoring, Evaluation and Learning (MEL) assistant for Papua New Guinea projects.

${filterText}

LIVE MONITORING SUMMARY:
${monitoringSummary}

LIVE EVALUATION SUMMARY:
${evaluationSummary}

VISUALS DATA (USE EXACTLY AS GIVEN):
- Monitoring KPI scores: ${JSON.stringify(monVisuals.kpi_scores)}
- Monitoring distribution: ${JSON.stringify(monVisuals.distribution)}
- Evaluation KPI scores: ${JSON.stringify(evalVisuals.kpi_scores)}
- Evaluation distribution: ${JSON.stringify(evalVisuals.distribution)}
- Combined score percent: ${JSON.stringify(combinedVisuals.combined_score_percent)}
- Combined distribution: ${JSON.stringify(combinedVisuals.combined_distribution)}

USER REQUEST:
${message}

${jsonInstruction}
`.trim();

    // Call OpenAI
    const openaiResponse = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${process.env.OPENAI_API_KEY}`,
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: [
          { role: 'system', content: 'You output strict JSON only when requested.' },
          { role: 'user', content: userPrompt },
        ],
        temperature: 0.3,
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

      console.error(
        'OpenAI API error:',
        openaiResponse.status,
        JSON.stringify(errorBody, null, 2)
      );

      return res.status(openaiResponse.status).json({
        success: false,
        error:
          errorBody?.error?.message ||
          `OpenAI API error (status ${openaiResponse.status})`,
        status: openaiResponse.status,
      });
    }

    const data = await openaiResponse.json();
    const rawText = data.choices?.[0]?.message?.content || '';

    // Parse strict JSON (or fallback)
    const aiJson = safeParseAiJson(rawText);

    if (!aiJson || typeof aiJson !== 'object') {
      // Fallback: still return computed visuals so dashboard can draw charts
      return res.json({
        success: true,
        reply: rawText || 'AI returned no usable JSON. Showing raw response.',
        visuals: {
          kpi_scores: monVisuals.kpi_scores,
          distribution: monVisuals.distribution,
          combined_score_percent: combinedVisuals.combined_score_percent,
          combined_distribution: combinedVisuals.combined_distribution,
        },
        used_filters: filters || {},
        monitoring_records_used: monFiltered.data?.length || 0,
        evaluation_records_used: evalFiltered.data?.length || 0,
        note: 'AI response was not valid JSON; returned raw text + backend visuals.',
      });
    }

    // Return report_markdown + visuals for the dashboard
    return res.json({
      success: true,
      reply: aiJson.report_markdown || '',
      report_title: aiJson.report_title || '',
      key_findings: Array.isArray(aiJson.key_findings) ? aiJson.key_findings : [],
      recommendations: Array.isArray(aiJson.recommendations) ? aiJson.recommendations : [],
      visuals: {
        // Use AI-provided but should match our computed visuals
        kpi_scores: aiJson.visuals?.kpi_scores || monVisuals.kpi_scores,
        distribution: aiJson.visuals?.distribution || monVisuals.distribution,
        combined_score_percent:
          typeof aiJson.visuals?.combined_score_percent === 'number'
            ? aiJson.visuals.combined_score_percent
            : combinedVisuals.combined_score_percent,
        combined_distribution: aiJson.visuals?.combined_distribution || combinedVisuals.combined_distribution,
      },
      used_filters: filters || {},
      monitoring_records_used: monFiltered.data?.length || 0,
      evaluation_records_used: evalFiltered.data?.length || 0,
    });
  } catch (err) {
    console.error('Backend error in /api/pas-ai-chat:', err);
    return res.status(500).json({
      success: false,
      error: 'Server error in /api/pas-ai-chat',
    });
  }
});

// Start server
app.listen(PORT, () => {
  console.log(`ProMEL OpenAI backend listening on http://localhost:${PORT}`);
});
