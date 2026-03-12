require('dotenv').config();
const express = require('express');
const multer = require('multer');
const path = require('path');
const Anthropic = require('@anthropic-ai/sdk');

const app = express();
const PORT = process.env.PORT || 3000;

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    const allowedExts = ['.csv', '.xlsx', '.xls', '.txt', '.json', '.pdf'];
    const ext = path.extname(file.originalname).toLowerCase();
    allowedExts.includes(ext) ? cb(null, true) : cb(new Error('File type not supported. Use CSV, Excel, PDF, TXT, or JSON.'));
  }
});

app.use(express.json({ limit: '15mb' }));
app.use(express.urlencoded({ extended: true, limit: '15mb' }));
app.use(express.static('public'));

async function parseFile(file) {
  const ext = path.extname(file.originalname).toLowerCase();

  if (['.csv', '.txt', '.json'].includes(ext)) {
    return file.buffer.toString('utf-8');
  }

  if (['.xlsx', '.xls'].includes(ext)) {
    const XLSX = require('xlsx');
    const workbook = XLSX.read(file.buffer, { type: 'buffer' });
    let content = '';
    workbook.SheetNames.forEach(name => {
      content += `--- Sheet: ${name} ---\n`;
      content += XLSX.utils.sheet_to_csv(workbook.Sheets[name]) + '\n\n';
    });
    return content;
  }

  if (ext === '.pdf') {
    try {
      const pdfParse = require('pdf-parse');
      const data = await pdfParse(file.buffer);
      return data.text;
    } catch {
      return 'PDF content could not be parsed. Please convert to CSV for best results.';
    }
  }

  return file.buffer.toString('utf-8');
}

function buildSystemPrompt() {
  return `You are a senior sustainability consultant at a world-class advisory firm with deep expertise in ESG reporting. You produce publication-ready, boardroom-quality reports that meet the standards of institutional investors, regulators, and auditors.

Your reports are:
- Precisely structured to each framework's requirements
- Professional in tone — clear, authoritative, and data-driven
- Populated with actual data from the provided input
- Honest about data gaps — noting "Data not available" or making clearly labeled estimates
- Formatted in clean markdown with proper headers, tables, and lists`;
}

function buildPrompt(framework, companyName, industry, year, data, context) {
  const truncatedData = data.substring(0, 18000);

  const frameworkInstructions = {
    TCFD: `Generate a comprehensive TCFD (Task Force on Climate-related Financial Disclosures) report.

# ${companyName} — TCFD Report | FY${year}
**Industry:** ${industry} | **Prepared by:** GrUE | **Report Date:** ${new Date().toLocaleDateString('en-GB', {day:'numeric',month:'long',year:'numeric'})}

---

## Executive Summary
Provide a concise 200-word overview of climate-related activities, key risks, and strategic commitments.

## 1. Governance
### 1.1 Board Oversight of Climate-Related Risks and Opportunities
Describe the board-level committee or individual responsible for climate oversight, frequency of reviews, and mechanisms used.

### 1.2 Management's Role in Assessing and Managing Climate-Related Risks
Describe management-level positions and committees, processes for climate risk escalation to the board.

## 2. Strategy
### 2.1 Climate-Related Risks and Opportunities Identified
| Risk / Opportunity | Category | Time Horizon | Potential Impact | Likelihood |
|---|---|---|---|---|
| [Derive from data] | | Short / Medium / Long | | High / Medium / Low |

### 2.2 Impact on Business, Strategy & Financial Planning
Describe how climate risks affect products, services, supply chains, and capital allocation.

### 2.3 Climate Scenario Analysis
Analyze resilience under three climate pathways:
- **1.5°C pathway** (Paris-aligned, aggressive transition)
- **2°C pathway** (Orderly transition)
- **4°C pathway** (Physical risk, delayed transition)

## 3. Risk Management
### 3.1 Identifying and Assessing Climate-Related Risks
### 3.2 Managing Climate-Related Risks
### 3.3 Integration into Overall Enterprise Risk Management

## 4. Metrics & Targets
### 4.1 Climate-Related Metrics Used

### 4.2 GHG Emissions Inventory
| Scope | Description | FY${year} (tCO2e) | Prior Year (tCO2e) | Methodology |
|---|---|---|---|---|
| Scope 1 | Direct emissions | | | |
| Scope 2 (Market-based) | Indirect energy emissions | | | |
| Scope 2 (Location-based) | Indirect energy emissions | | | |
| Scope 3 | Value chain emissions | | | |

### 4.3 Targets & Performance Against Targets
| Target | Base Year | Target Year | % Reduction | Progress |
|---|---|---|---|---|

## Recommendations
Provide 5–7 specific, actionable recommendations for improving TCFD alignment.

## TCFD Recommendation Index
| TCFD Recommended Disclosure | Section in This Report | Status |
|---|---|---|
| Board oversight | 1.1 | ✓ Addressed |
| Management's role | 1.2 | ✓ Addressed |
| Climate risks/opportunities | 2.1 | ✓ Addressed |
| Impact on strategy | 2.2 | ✓ Addressed |
| Scenario analysis | 2.3 | ✓ Addressed |
| Risk identification process | 3.1 | ✓ Addressed |
| Risk management process | 3.2 | ✓ Addressed |
| Integration of risk management | 3.3 | ✓ Addressed |
| Metrics used | 4.1 | ✓ Addressed |
| Scope 1, 2, 3 emissions | 4.2 | ✓ Addressed |
| Targets | 4.3 | ✓ Addressed |`,

    SASB: `Generate a comprehensive SASB (Sustainability Accounting Standards Board) disclosure report.

# ${companyName} — SASB Disclosure | FY${year}
**Industry:** ${industry} | **Prepared by:** GrUE | **Report Date:** ${new Date().toLocaleDateString('en-GB', {day:'numeric',month:'long',year:'numeric'})}

---

## Executive Summary
Brief overview of material sustainability topics and key disclosures.

## 1. Industry Classification
**SASB Industry:** [Determine the appropriate SASB industry from data/context]
**Applicable SASB Standard:** [Specify the standard]

## 2. Activity Metrics
| Activity Metric | SASB Code | Unit | FY${year} Value |
|---|---|---|---|
| [Relevant activity metrics for the industry] | | | |

## 3. Environmental Topics
### 3.1 Energy Management
| Metric | SASB Code | Unit | FY${year} |
|---|---|---|---|
| Total energy consumed | | GJ | |
| % grid electricity | | % | |
| % renewable energy | | % | |

### 3.2 Water Management
### 3.3 GHG Emissions & Air Quality
| Metric | SASB Code | Unit | FY${year} |
|---|---|---|---|
| Gross Scope 1 emissions | | tCO2e | |
| Gross Scope 2 emissions | | tCO2e | |

### 3.4 Waste & Hazardous Materials

## 4. Social Capital
### 4.1 Data Security & Privacy
### 4.2 Access & Affordability
### 4.3 Customer Welfare & Product Quality

## 5. Human Capital
| Metric | SASB Code | Unit | FY${year} |
|---|---|---|---|
| Total employees | | # | |
| Employee turnover rate | | % | |
| Total recordable incident rate | | per 1M hrs | |
| Fatality rate | | per 1M hrs | |
| % female employees | | % | |
| % employees — minority groups | | % | |

## 6. Business Model & Innovation
### 6.1 Product Innovation & Lifecycle Management
### 6.2 Supply Chain Management

## 7. Leadership & Governance
### 7.1 Business Ethics & Transparency
### 7.2 Competitive Behavior
### 7.3 Systemic Risk Management

## Recommendations
Provide 5–7 specific recommendations for improving SASB disclosure quality.

## SASB Disclosure Index
| Disclosure | SASB Code | Section | Data Status |
|---|---|---|---|`,

    GRI: `Generate a comprehensive GRI Standards sustainability report (GRI 2021 framework).

# ${companyName} — GRI Sustainability Report | FY${year}
**Industry:** ${industry} | **Reporting Period:** 1 January – 31 December ${year}
**Prepared by:** GrUE | **GRI Standards:** GRI 1: Foundation 2021

---

## About This Report
**Reporting Boundary:** [Organizational boundary and consolidation approach]
**Reporting Frequency:** Annual
**External Assurance:** [Assurance status]
**Contact:** [Sustainability reporting contact]

## Executive Summary

## GRI 2: General Disclosures
### 2-1 Organizational Details
### 2-2 Entities Included in Sustainability Reporting
### 2-3 Reporting Period, Frequency and Contact Point
### 2-4 Restatements of Information
### 2-6 Activities, Value Chain and Other Business Relationships
### 2-9 Governance Structure and Composition
### 2-22 Statement on Sustainable Development Strategy
### 2-23 Policy Commitments
### 2-29 Approach to Stakeholder Engagement
### 2-30 Collective Bargaining Agreements

## GRI 3: Material Topics
### 3-1 Process to Determine Material Topics
Describe the materiality assessment methodology, stakeholder groups consulted, and prioritization approach.

### 3-2 List of Material Topics
| Material Topic | GRI Topic Standard | Relevance to Business | Stakeholder Priority |
|---|---|---|---|

### Materiality Matrix
Describe or represent the materiality assessment results across impact significance and business relevance axes.

## Environmental Performance (GRI 300 Series)

### GRI 302: Energy
| Energy Metric | Unit | FY${year} | Prior Year |
|---|---|---|---|
| Total energy consumption within org | GJ | | |
| Energy from renewable sources | GJ | | |
| Energy intensity | GJ / [unit] | | |

### GRI 303: Water and Effluents
| Metric | Unit | FY${year} |
|---|---|---|
| Total water withdrawal | ML | |
| Water recycled and reused | ML | |

### GRI 305: Emissions
| Emission Category | Unit | FY${year} | Prior Year |
|---|---|---|---|
| Scope 1 — Direct GHG emissions | tCO2e | | |
| Scope 2 — Indirect GHG emissions (market-based) | tCO2e | | |
| Scope 2 — Indirect GHG emissions (location-based) | tCO2e | | |
| Scope 3 — Other indirect emissions | tCO2e | | |
| GHG emissions intensity | tCO2e / [unit] | | |

### GRI 306: Waste
| Metric | Unit | FY${year} |
|---|---|---|
| Total waste generated | tonnes | |
| Waste diverted from disposal | tonnes | |
| Waste directed to disposal | tonnes | |

## Social Performance (GRI 400 Series)

### GRI 401: Employment
| Metric | Unit | FY${year} |
|---|---|---|
| New employee hires | # | |
| Employee turnover | # and % | |

### GRI 403: Occupational Health & Safety
### GRI 404: Training and Education
### GRI 405: Diversity and Equal Opportunity
### GRI 413: Local Communities

## Governance Performance
### Anti-Corruption (GRI 205)
### Anti-Competitive Behavior (GRI 206)

## Recommendations
5–7 actionable recommendations for improving GRI disclosure quality and sustainability performance.

## GRI Content Index
| GRI Standard | Disclosure | Report Section | Omission / Reason |
|---|---|---|---|
| GRI 2: General Disclosures 2021 | 2-1 | About This Report | — |
| GRI 2: General Disclosures 2021 | 2-6 | About This Report | — |
| GRI 2: General Disclosures 2021 | 2-22 | Executive Summary | — |
| GRI 3: Material Topics 2021 | 3-1 | Material Topics | — |
| GRI 3: Material Topics 2021 | 3-2 | Material Topics | — |
| GRI 302: Energy 2016 | 302-1 | Energy | — |
| GRI 305: Emissions 2016 | 305-1 | Emissions | — |
| GRI 305: Emissions 2016 | 305-2 | Emissions | — |`,

    CDP: `Generate a comprehensive CDP Climate Change questionnaire response.

# ${companyName} — CDP Climate Change Response | ${year}
**Industry:** ${industry} | **Prepared by:** GrUE | **Disclosure Year:** ${year}

---

## C0. Introduction
**Legal Name:** ${companyName}
**Primary Activity:** ${industry}
**Reporting Year:** ${year}
**CDP Submission Status:** Draft

## C1. Governance
### C1.1 Board-Level Oversight of Climate
| Board-level body | Frequency of review | Governance mechanisms |
|---|---|---|
| [Board committee] | [Frequency] | [Review of strategy, targets, major projects] |

### C1.2 Management-Level Responsibility for Climate
### C1.3 Employee Incentives for Climate Management

## C2. Risks and Opportunities
### C2.1 Risk and Opportunity Assessment Process
Describe scope, time horizons (short: 0–2 yrs, medium: 2–5 yrs, long: 5+ yrs), and assessment methodology.

### C2.2 Identified Climate-Related Risks
| Risk | Category | Time Horizon | Likelihood | Magnitude of Impact | Financial Implications |
|---|---|---|---|---|---|
| [Physical / Transition risk from data] | | | | | |

### C2.3 Identified Climate-Related Opportunities
| Opportunity | Category | Time Horizon | Likelihood | Potential Financial Impact |
|---|---|---|---|---|

## C3. Business Strategy
### C3.1 Climate Integration in Strategy
### C3.2 Climate Scenario Analysis
Describe scenarios used (e.g., IEA NZE 2050, IPCC SSP1-1.9, SSP3-7.0) and key insights.

## C4. Targets and Performance
### C4.1 Absolute Emissions Targets
| Target | Scope | Base Year | Base Year Emissions (tCO2e) | Target Year | % Reduction | Status |
|---|---|---|---|---|---|---|

### C4.2 Intensity Targets
### C4.3 Other Climate Targets (Renewable Energy, etc.)

## C5. Emissions Methodology
### C5.1 Organizational Boundary
### C5.2 GHG Calculation Methodology and Tools

## C6. Emissions Data
### C6.1 Gross Global Scope 1 Emissions
**Total Scope 1 (tCO2e):** [Value from data]

### C6.2 Gross Scope 2 Emissions
| Approach | tCO2e |
|---|---|
| Market-based | |
| Location-based | |

### C6.3 Scope 3 Categories
| Scope 3 Category | tCO2e | Relevance | Data Quality | Verification |
|---|---|---|---|---|
| 1. Purchased goods & services | | | | |
| 3. Fuel & energy related activities | | | | |
| 6. Business travel | | | | |
| 11. Use of sold products | | | | |
| 15. Investments | | | | |

## C7. Emissions Breakdown
### C7.1 Geographical Breakdown
### C7.2 Business Division Breakdown

## C8. Energy
| Energy Source | Consumption (MWh) | % of Total |
|---|---|---|
| Coal | | |
| Natural gas | | |
| Electricity (grid) | | |
| Renewable electricity | | |
| **Total** | | **100%** |

## C9. Additional Metrics
### C9.1 Other Climate-Related Metrics

## C10. Verification
### C10.1 Scope 1 & 2 Verification
| Verification body | Standard | Boundary | Type |
|---|---|---|---|

## C11. Carbon Pricing
### C11.2 Exposure to Carbon Pricing Mechanisms

## C12. Engagement
### C12.1 Supply Chain Engagement
### C12.3 Industry Association Engagement

## Recommendations
5–7 specific recommendations to improve CDP score and climate performance.

## CDP Scoring Reference
Key areas assessed: Governance (C1), Risk & Opportunity (C2), Business Strategy (C3), Targets (C4), Emissions (C6), Verification (C10). Aiming for Leadership band (A/A-) requires robust scenario analysis, verified emissions data, and board-level accountability.`
  };

  return `${frameworkInstructions[framework]}

---

## COMPANY DATA PROVIDED:
\`\`\`
${truncatedData}
\`\`\`

## ADDITIONAL CONTEXT:
${context || 'None provided.'}

---

**Instructions:** Populate ALL tables and sections using the actual data above. Where data is unavailable, write "Data not available — recommend collection for next reporting cycle." Make estimates only when clearly labeled as "[Estimated]". Use precise, professional language throughout. Ensure every section is substantive and adds value.`;
}

app.post('/api/generate-report', upload.single('file'), async (req, res) => {
  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');
  res.flushHeaders();

  const send = (data) => res.write(`data: ${JSON.stringify(data)}\n\n`);

  try {
    const { companyName, industry, year, additionalContext } = req.body;
    const apiKey = req.body.apiKey || process.env.ANTHROPIC_API_KEY;

    if (!apiKey) {
      send({ type: 'error', message: 'No API key found. Please add your Anthropic API key.' });
      return res.end();
    }

    let fileContent = '';
    if (req.file) {
      try {
        fileContent = await parseFile(req.file);
      } catch (err) {
        send({ type: 'error', message: `File parsing error: ${err.message}` });
        return res.end();
      }
    } else if (req.body.pastedData) {
      fileContent = req.body.pastedData;
    }

    if (!fileContent.trim()) {
      send({ type: 'error', message: 'No data provided. Please upload a file or paste your data.' });
      return res.end();
    }

    let frameworkList;
    try {
      frameworkList = JSON.parse(req.body.frameworks || '[]');
    } catch {
      frameworkList = ['TCFD'];
    }

    if (!frameworkList.length) {
      send({ type: 'error', message: 'Please select at least one reporting framework.' });
      return res.end();
    }

    const client = new Anthropic({ apiKey });

    for (const framework of frameworkList) {
      send({ type: 'framework_start', framework });

      const stream = await client.messages.stream({
        model: 'claude-opus-4-6',
        max_tokens: 8000,
        system: buildSystemPrompt(),
        messages: [{ role: 'user', content: buildPrompt(framework, companyName || 'Company', industry || 'General Industry', year || new Date().getFullYear(), fileContent, additionalContext) }]
      });

      for await (const chunk of stream) {
        if (chunk.type === 'content_block_delta' && chunk.delta?.type === 'text_delta') {
          send({ type: 'text', framework, text: chunk.delta.text });
        }
      }

      send({ type: 'framework_done', framework });
    }

    send({ type: 'done' });
    res.end();

  } catch (err) {
    send({ type: 'error', message: err.message || 'An unexpected error occurred.' });
    res.end();
  }
});

app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', apiKeyConfigured: !!process.env.ANTHROPIC_API_KEY });
});

app.listen(PORT, () => {
  console.log(`\n  GrUE Sustainability Report Generator`);
  console.log(`  Running at http://localhost:${PORT}\n`);
  if (!process.env.ANTHROPIC_API_KEY) {
    console.log(`  Note: No ANTHROPIC_API_KEY in .env — users will need to provide their own key.\n`);
  }
});
