import os
import io
import json
import csv
from flask import Flask, request, jsonify, Response, send_from_directory
import anthropic
import google.generativeai as genai
from groq import Groq

app = Flask(__name__, static_folder='public')

# ── File Parsing ──────────────────────────────────────────────────────────────

def parse_file(file):
    filename = file.filename.lower()

    if filename.endswith(('.csv', '.txt')):
        return file.read().decode('utf-8', errors='replace')

    if filename.endswith('.json'):
        raw = file.read().decode('utf-8', errors='replace')
        try:
            parsed = json.loads(raw)
            return json.dumps(parsed, indent=2)
        except Exception:
            return raw

    if filename.endswith(('.xlsx', '.xls')):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(file.read()), read_only=True, data_only=True)
            output = []
            for sheet in wb.worksheets:
                output.append(f'--- Sheet: {sheet.title} ---')
                for row in sheet.iter_rows(values_only=True):
                    output.append(','.join(str(c) if c is not None else '' for c in row))
            return '\n'.join(output)
        except Exception as e:
            return f'Excel parsing error: {e}. Please convert to CSV.'

    if filename.endswith('.pdf'):
        try:
            from pdfminer.high_level import extract_text
            text = extract_text(io.BytesIO(file.read()))
            return text if text.strip() else 'PDF text could not be extracted.'
        except Exception as e:
            return f'PDF parsing error: {e}. Please convert to text/CSV.'

    return file.read().decode('utf-8', errors='replace')


# ── Prompt Construction ───────────────────────────────────────────────────────

SYSTEM_PROMPT = """You are a senior sustainability consultant at a world-class advisory firm with deep expertise in ESG reporting. You produce publication-ready, boardroom-quality reports that meet the standards of institutional investors, regulators, and auditors.

Your reports are:
- Precisely structured to each framework's requirements
- Professional in tone — clear, authoritative, and data-driven
- Populated with actual data from the provided input
- Honest about data gaps — noting "Data not available" or making clearly labeled estimates
- Formatted in clean markdown with proper headers, tables, and lists"""


def build_prompt(framework, company, industry, year, data, context):
    data_truncated = data[:18000]
    from datetime import datetime
    report_date = datetime.now().strftime('%d %B %Y')

    frameworks = {
        'TCFD': f"""Generate a comprehensive TCFD (Task Force on Climate-related Financial Disclosures) report.

# {company} — TCFD Report | FY{year}
**Industry:** {industry} | **Prepared by:** GrUE | **Report Date:** {report_date}

---

## Executive Summary
Provide a concise 200-word overview of climate-related activities, key risks, and strategic commitments.

## 1. Governance
### 1.1 Board Oversight of Climate-Related Risks and Opportunities
Describe the board-level committee responsible for climate oversight, frequency of reviews, and mechanisms used.

### 1.2 Management's Role in Assessing and Managing Climate-Related Risks
Describe management-level positions, committees, and escalation processes.

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
| Scope | Description | FY{year} (tCO2e) | Prior Year (tCO2e) | Methodology |
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
| TCFD Recommended Disclosure | Section | Status |
|---|---|---|
| Board oversight | 1.1 | ✓ Addressed |
| Management's role | 1.2 | ✓ Addressed |
| Climate risks/opportunities | 2.1 | ✓ Addressed |
| Impact on strategy | 2.2 | ✓ Addressed |
| Scenario analysis | 2.3 | ✓ Addressed |
| Risk identification | 3.1 | ✓ Addressed |
| Risk management | 3.2 | ✓ Addressed |
| Integration | 3.3 | ✓ Addressed |
| Metrics | 4.1 | ✓ Addressed |
| Scope 1/2/3 emissions | 4.2 | ✓ Addressed |
| Targets | 4.3 | ✓ Addressed |""",

        'SASB': f"""Generate a comprehensive SASB (Sustainability Accounting Standards Board) disclosure report.

# {company} — SASB Disclosure | FY{year}
**Industry:** {industry} | **Prepared by:** GrUE | **Report Date:** {report_date}

---

## Executive Summary
Brief overview of material sustainability topics and key disclosures.

## 1. Industry Classification
**SASB Industry:** [Determine the appropriate SASB industry from data/context]
**Applicable SASB Standard:** [Specify the standard]

## 2. Activity Metrics
| Activity Metric | SASB Code | Unit | FY{year} Value |
|---|---|---|---|
| [Relevant activity metrics for the industry] | | | |

## 3. Environmental Topics
### 3.1 Energy Management
| Metric | SASB Code | Unit | FY{year} |
|---|---|---|---|
| Total energy consumed | | GJ | |
| % grid electricity | | % | |
| % renewable energy | | % | |

### 3.2 Water Management
| Metric | SASB Code | Unit | FY{year} |
|---|---|---|---|

### 3.3 GHG Emissions & Air Quality
| Metric | SASB Code | Unit | FY{year} |
|---|---|---|---|
| Gross Scope 1 emissions | | tCO2e | |
| Gross Scope 2 emissions | | tCO2e | |

### 3.4 Waste & Hazardous Materials

## 4. Social Capital
### 4.1 Data Security & Privacy
### 4.2 Access & Affordability
### 4.3 Customer Welfare & Product Quality

## 5. Human Capital
| Metric | SASB Code | Unit | FY{year} |
|---|---|---|---|
| Total employees | | # | |
| Employee turnover rate | | % | |
| Total recordable incident rate | | per 1M hrs | |
| Fatality rate | | per 1M hrs | |
| % female employees | | % | |
| % minority employees | | % | |

## 6. Business Model & Innovation
### 6.1 Product Innovation & Lifecycle Management
### 6.2 Supply Chain Management

## 7. Leadership & Governance
### 7.1 Business Ethics & Transparency
### 7.2 Competitive Behavior

## Recommendations
5–7 specific recommendations for improving SASB disclosure quality.

## SASB Disclosure Index
| Disclosure | SASB Code | Section | Data Status |
|---|---|---|---|""",

        'GRI': f"""Generate a comprehensive GRI Standards sustainability report (GRI 2021 framework).

# {company} — GRI Sustainability Report | FY{year}
**Industry:** {industry} | **Reporting Period:** 1 January – 31 December {year}
**Prepared by:** GrUE | **GRI Standards:** GRI 1: Foundation 2021

---

## About This Report
**Reporting Boundary:** [Organizational boundary and consolidation approach]
**Reporting Frequency:** Annual
**GRI Content Index:** Included at end of report

## Executive Summary

## GRI 2: General Disclosures
### 2-1 Organizational Details
### 2-2 Entities Included in Sustainability Reporting
### 2-3 Reporting Period, Frequency and Contact Point
### 2-6 Activities, Value Chain and Other Business Relationships
### 2-9 Governance Structure and Composition
### 2-22 Statement on Sustainable Development Strategy
### 2-29 Approach to Stakeholder Engagement

## GRI 3: Material Topics
### 3-1 Process to Determine Material Topics
Describe the materiality assessment methodology, stakeholder groups consulted, and prioritization approach.

### 3-2 List of Material Topics
| Material Topic | GRI Topic Standard | Relevance | Stakeholder Priority |
|---|---|---|---|

## Environmental Performance (GRI 300 Series)

### GRI 302: Energy
| Energy Metric | Unit | FY{year} | Prior Year |
|---|---|---|---|
| Total energy consumption | GJ | | |
| Renewable energy | GJ | | |
| Energy intensity | GJ/[unit] | | |

### GRI 303: Water and Effluents
| Metric | Unit | FY{year} |
|---|---|---|
| Total water withdrawal | ML | |
| Water recycled/reused | ML | |

### GRI 305: Emissions
| Emission Category | Unit | FY{year} | Prior Year |
|---|---|---|---|
| Scope 1 — Direct GHG | tCO2e | | |
| Scope 2 — Market-based | tCO2e | | |
| Scope 2 — Location-based | tCO2e | | |
| Scope 3 — Other indirect | tCO2e | | |
| GHG intensity | tCO2e/[unit] | | |

### GRI 306: Waste
| Metric | Unit | FY{year} |
|---|---|---|
| Total waste generated | tonnes | |
| Waste diverted from disposal | tonnes | |

## Social Performance (GRI 400 Series)
### GRI 401: Employment
### GRI 403: Occupational Health & Safety
### GRI 404: Training and Education
### GRI 405: Diversity and Equal Opportunity
### GRI 413: Local Communities

## Recommendations
5–7 actionable recommendations for improving GRI disclosure quality.

## GRI Content Index
| GRI Standard | Disclosure | Report Section | Omission |
|---|---|---|---|
| GRI 2: General 2021 | 2-1 | About This Report | — |
| GRI 2: General 2021 | 2-6 | About This Report | — |
| GRI 2: General 2021 | 2-22 | Executive Summary | — |
| GRI 3: Material Topics 2021 | 3-1 | Material Topics | — |
| GRI 302: Energy 2016 | 302-1 | Energy | — |
| GRI 305: Emissions 2016 | 305-1 | Emissions | — |
| GRI 305: Emissions 2016 | 305-2 | Emissions | — |""",

        'CDP': f"""Generate a comprehensive CDP Climate Change questionnaire response.

# {company} — CDP Climate Change Response | {year}
**Industry:** {industry} | **Prepared by:** GrUE | **Report Date:** {report_date}

---

## C0. Introduction
**Legal Name:** {company}
**Primary Activity:** {industry}
**Reporting Year:** {year}

## C1. Governance
### C1.1 Board-Level Oversight
| Board-level body | Frequency | Governance mechanisms |
|---|---|---|
| [Committee] | [Frequency] | [Review of strategy, targets, major projects] |

### C1.2 Management-Level Responsibility
### C1.3 Employee Incentives for Climate Management

## C2. Risks and Opportunities
### C2.1 Risk and Opportunity Assessment Process
Scope, time horizons (short: 0–2 yrs, medium: 2–5 yrs, long: 5+ yrs), and methodology.

### C2.2 Identified Climate-Related Risks
| Risk | Category | Time Horizon | Likelihood | Magnitude | Financial Impact |
|---|---|---|---|---|---|

### C2.3 Identified Climate-Related Opportunities
| Opportunity | Category | Time Horizon | Likelihood | Potential Impact |
|---|---|---|---|---|

## C3. Business Strategy
### C3.1 Climate Integration in Strategy
### C3.2 Climate Scenario Analysis
Describe scenarios used (IEA NZE 2050, IPCC SSP1-1.9, SSP3-7.0) and key insights.

## C4. Targets and Performance
### C4.1 Absolute Emissions Targets
| Target | Scope | Base Year | Base Emissions (tCO2e) | Target Year | % Reduction | Status |
|---|---|---|---|---|---|---|

### C4.2 Intensity Targets
### C4.3 Other Climate Targets

## C5. Emissions Methodology
### C5.1 Organizational Boundary
### C5.2 Calculation Methodology and Tools

## C6. Emissions Data
### C6.1 Gross Scope 1 Emissions
**Total Scope 1 (tCO2e):** [Value from data]

### C6.2 Gross Scope 2 Emissions
| Approach | tCO2e |
|---|---|
| Market-based | |
| Location-based | |

### C6.3 Scope 3 Categories
| Category | tCO2e | Relevance | Data Quality | Verified |
|---|---|---|---|---|
| 1. Purchased goods & services | | | | |
| 3. Fuel & energy activities | | | | |
| 6. Business travel | | | | |
| 11. Use of sold products | | | | |
| 15. Investments | | | | |

## C7. Emissions Breakdown
### C7.1 By Geography
### C7.2 By Business Division

## C8. Energy
| Source | MWh | % of Total |
|---|---|---|
| Coal | | |
| Natural gas | | |
| Grid electricity | | |
| Renewable electricity | | |
| **Total** | | **100%** |

## C10. Verification
| Verification body | Standard | Boundary | Type |
|---|---|---|---|

## C11. Carbon Pricing
### Exposure to Carbon Pricing Mechanisms

## C12. Engagement
### Supply Chain Engagement
### Industry Association Engagement

## Recommendations
5–7 recommendations to improve CDP score and climate performance.

## CDP Scoring Reference
Leadership band (A/A-) requires: board accountability (C1), scenario analysis (C3), verified data (C10), supply chain engagement (C12)."""
    }

    return f"""{frameworks[framework]}

---

## COMPANY DATA PROVIDED:
```
{data_truncated}
```

## ADDITIONAL CONTEXT:
{context or 'None provided.'}

---

**Instructions:** Populate ALL tables and sections using the actual data above. Where data is unavailable, write "Data not available — recommend collection for next reporting cycle." Label estimates clearly as "[Estimated]". Use precise, professional language. Ensure every section is substantive and adds value."""


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route('/')
def index():
    return send_from_directory('public', 'index.html')

@app.route('/assets/<path:filename>')
def assets(filename):
    return send_from_directory('public/assets', filename)

@app.route('/api/health')
def health():
    return jsonify({
        'status': 'ok',
        'apiKeyConfigured': bool(os.environ.get('ANTHROPIC_API_KEY'))
    })

@app.route('/api/generate-report', methods=['POST'])
def generate_report():
    # ── CRITICAL: extract ALL request data here, BEFORE the generator.
    # Flask request context does not exist inside a generator body.
    provider = request.form.get('provider', 'anthropic').strip().lower()
    api_key = request.form.get('apiKey', '').strip()
    if not api_key:
        api_key = os.environ.get('ANTHROPIC_API_KEY', '') if provider == 'anthropic' else os.environ.get('GOOGLE_API_KEY', '')

    file_content = ''
    file_parse_error = ''
    if 'file' in request.files and request.files['file'].filename:
        try:
            file_content = parse_file(request.files['file'])
        except Exception as e:
            file_parse_error = str(e)

    pasted = request.form.get('pastedData', '').strip()
    if not file_content and pasted:
        file_content = pasted

    company  = (request.form.get('companyName') or 'Company').strip()
    industry = (request.form.get('industry') or 'General Industry').strip()
    year     = (request.form.get('year') or '2024').strip()
    context  = (request.form.get('additionalContext') or '').strip()

    try:
        framework_list = json.loads(request.form.get('frameworks', '[]'))
    except Exception:
        framework_list = ['TCFD']

    def stream():
        def send(data):
            yield f"data: {json.dumps(data)}\n\n"

        try:
            if not api_key:
                yield from send({'type': 'error', 'message': 'No API key found. Please enter your Anthropic API key in the form below.'})
                return

            if file_parse_error:
                yield from send({'type': 'error', 'message': f'Could not read your file: {file_parse_error}. Try converting to CSV or paste the data directly.'})
                return

            if not file_content.strip():
                yield from send({'type': 'error', 'message': 'No data provided. Please upload a file or paste your sustainability data.'})
                return

            if not framework_list:
                yield from send({'type': 'error', 'message': 'Please select at least one reporting framework.'})
                return

            for fw in framework_list:
                yield from send({'type': 'framework_start', 'framework': fw})
                prompt = build_prompt(fw, company, industry, year, file_content, context)

                if provider == 'groq':
                    client = Groq(api_key=api_key)
                    stream_resp = client.chat.completions.create(
                        model='llama-3.3-70b-versatile',
                        messages=[
                            {'role': 'system', 'content': SYSTEM_PROMPT},
                            {'role': 'user', 'content': prompt}
                        ],
                        max_tokens=8000,
                        stream=True
                    )
                    for chunk in stream_resp:
                        text = chunk.choices[0].delta.content
                        if text:
                            yield from send({'type': 'text', 'framework': fw, 'text': text})

                elif provider == 'gemini':
                    genai.configure(api_key=api_key)
                    model = genai.GenerativeModel(
                        model_name='gemini-1.5-flash',
                        system_instruction=SYSTEM_PROMPT
                    )
                    response = model.generate_content(prompt, stream=True)
                    for chunk in response:
                        if chunk.text:
                            yield from send({'type': 'text', 'framework': fw, 'text': chunk.text})

                else:
                    client = anthropic.Anthropic(api_key=api_key)
                    with client.messages.stream(
                        model='claude-opus-4-6',
                        max_tokens=8000,
                        system=SYSTEM_PROMPT,
                        messages=[{'role': 'user', 'content': prompt}]
                    ) as stream_resp:
                        for text in stream_resp.text_stream:
                            yield from send({'type': 'text', 'framework': fw, 'text': text})

                yield from send({'type': 'framework_done', 'framework': fw})

            yield from send({'type': 'done'})

        except anthropic.AuthenticationError:
            yield from send({'type': 'error', 'message': 'Invalid Anthropic API key — check it at console.anthropic.com/keys'})
        except anthropic.RateLimitError:
            yield from send({'type': 'error', 'message': 'Anthropic rate limit reached. Please wait a moment and try again.'})
        except anthropic.BadRequestError as e:
            yield from send({'type': 'error', 'message': f'Request error: {e}'})
        except Exception as e:
            msg = str(e)
            if 'api_key' in msg.lower() or 'api key' in msg.lower() or '400' in msg or 'invalid' in msg.lower():
                yield from send({'type': 'error', 'message': f'Invalid API key — please check your {"Google AI Studio" if provider == "gemini" else "Anthropic"} key.'})
            else:
                yield from send({'type': 'error', 'message': msg})

    return Response(
        stream(),
        mimetype='text/event-stream',
        headers={
            'Cache-Control': 'no-cache',
            'X-Accel-Buffering': 'no',
            'Connection': 'keep-alive'
        }
    )


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 3000))
    print(f'\n  GrUE Sustainability Report Generator')
    print(f'  Running at http://localhost:{port}\n')
    if not os.environ.get('ANTHROPIC_API_KEY'):
        print('  Note: No ANTHROPIC_API_KEY in .env — users will need to provide their own key.\n')
    app.run(host='0.0.0.0', port=port, debug=False, threaded=True)
