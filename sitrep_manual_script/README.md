# 🛡️ Daily Cybersecurity SITREP Generator

**Evergreen Healthcare — Security Operations**

A standalone Python script that gathers cybersecurity intelligence from multiple sources and generates a multi-page categorized Daily Situation Report (SITREP) as a Word document and HTML-formatted email.

---

## 📋 Features

- **Multi-source intelligence gathering:**
  - CISA Known Exploited Vulnerabilities (KEV) catalog
  - CISA alerts and advisories (RSS)
  - 30+ vendor security advisory feeds (Cisco, Microsoft, Palo Alto, medical device vendors, etc.)
  - General cybersecurity news feeds
- **Healthcare-focused relevance filtering** against the hospital's technology stack
- **Three-tier threat classification:**
  - 🔴 **IMMEDIATE** (Red) — Active exploitation, critical vulns in our stack
  - 🟠 **PRIORITY** (Amber) — High-severity items requiring attention
  - 🟢 **ROUTINE** (Green) — Informational, lower-severity items
- **Overall threat posture assessment:** ELEVATED / GUARDED / LOW
- **Multi-page Word document** with detailed technical appendices:
  - **Page 1:** Concise executive summary (1-page overview for leadership)
  - **Page 2:** APPENDIX A — IMMEDIATE Threats (detailed technical breakdown)
  - **Page 3:** APPENDIX B — PRIORITY Threats (detailed technical breakdown)
  - **Page 4:** APPENDIX C — ROUTINE Items (detailed technical breakdown)
- **CVSS score lookup** via NVD (National Vulnerability Database) API
- **Hospital tech stack mapping** — matches threats to specific hospital systems
- **Dual-audience approach:**
  - Executives → concise 1-page summary (email body)
  - Technical staff → comprehensive appendices with remediation details (Word document)
- **Email delivery** via SMTP (Office 365, Gmail, or any SMTP server)
- **Parallel fetching** for fast collection from 30+ sources
- **Detailed logging** for troubleshooting

---

## 📄 Appendix Content (Per Threat)

Each threat entry in the appendices includes 11 comprehensive technical detail fields:

| # | Field | Description |
|---|-------|-------------|
| 1 | **CVE/Advisory ID** | Full CVE number or vendor advisory identifier |
| 2 | **Severity Score (CVSS)** | CVSS base score, severity level, and vector string (from NVD API) |
| 3 | **Affected Systems/Vendors** | Specific products, versions, and platforms |
| 4 | **Systems in Our Environment** | Which hospital systems are affected (mapped from tech stack) |
| 5 | **Technical Description** | Detailed technical explanation of the vulnerability |
| 6 | **Attack Vector** | How the threat could be exploited (Network, Local, Physical, etc.) |
| 7 | **Potential Impact** | Healthcare-specific impact analysis |
| 8 | **Remediation Steps** | Specific technical actions — patches, configs, workarounds |
| 9 | **Timeline/Urgency** | When action must be taken |
| 10 | **Source/Reference** | Full URL or source citation |
| 11 | **Date Published** | When the advisory was released |

---

## 🚀 Quick Start

### 1. Prerequisites

- Python 3.8 or later
- Internet access (to reach CISA, vendor sites, news feeds, NVD API)

### 2. Install Dependencies

```bash
# Clone or copy this directory, then:
cd sitrep_manual_script
pip install -r requirements.txt
```

### 3. Run the Script

```bash
# Generate SITREP (DOCX + HTML, attempt email)
python generate_sitrep_manual.py

# Generate without sending email
python generate_sitrep_manual.py --no-email

# Specify output directory
python generate_sitrep_manual.py --output /path/to/reports

# Override email recipient
python generate_sitrep_manual.py --email someone@example.com
```

---

## 📧 Email Configuration

The script sends the SITREP as an HTML-formatted email. The email body contains the **concise executive summary** with a note: *"Detailed technical appendices are included in the full report document."* The complete multi-page Word document with appendices is generated alongside.

Configure SMTP credentials using **environment variables** (recommended) or by editing the `CONFIG` section in the script.

### Option A: Environment Variables (recommended)

```bash
# Linux / macOS
export SITREP_SMTP_USER="your-email@evergreenhealthcare.org"
export SITREP_SMTP_PASS="your-app-password"

# Windows (PowerShell)
$env:SITREP_SMTP_USER = "your-email@evergreenhealthcare.org"
$env:SITREP_SMTP_PASS = "your-app-password"
```

### Option B: Edit the Script

Open `generate_sitrep_manual.py` and edit the `CONFIG` dictionary near the top.

> ⚠️ **Note:** If SMTP credentials are not configured, the script will still generate all report files. The HTML email body will be saved locally as a fallback.

### Default Settings

| Setting | Value |
|---------|-------|
| Recipient | `mdsimbre@evergreenhealthcare.org` |
| SMTP Server | `smtp.office365.com` |
| SMTP Port | `587` (STARTTLS) |
| Subject | `Daily Cybersecurity SITREP - [Date] - Threat Level: [LEVEL]` |

---

## 📂 Output Files

After running, the script creates these files in the output directory:

| File | Description |
|------|-------------|
| `Daily_Cybersecurity_SITREP_YYYYMMDD.docx` | **Multi-page** Word document: executive summary + 3 appendix pages |
| `Daily_Cybersecurity_SITREP_YYYYMMDD.html` | HTML version of the executive summary (email body) |
| `sitrep_generator.log` | Detailed log for troubleshooting |

---

## ⚙️ Customization

All configuration is at the top of `generate_sitrep_manual.py`:

### Technology Stack

Edit the `TECHNOLOGY_STACK` list to match your hospital's systems:

```python
TECHNOLOGY_STACK = [
    "Epic", "Palo Alto", "Microsoft", "Cisco", ...
]
```

### Hospital System Mapping

Edit `TECH_STACK_DETAILS` dictionary to customize the "Systems in Our Environment" descriptions that appear in appendix entries:

```python
TECH_STACK_DETAILS = {
    "Epic": "Epic EHR / EpicCare – Primary electronic health record system",
    "Palo Alto": "Palo Alto Networks – Perimeter firewalls & GlobalProtect VPN",
    ...
}
```

### Vendor Advisory URLs

Add or remove vendor security feeds in `VENDOR_ADVISORY_URLS`.

### Severity Classification

The `_classify_severity()` function uses keyword heuristics. Adjust the keyword lists to tune sensitivity.

### CVSS Lookup

The script queries the NVD API for CVSS scores. The lookup:
- Caches results to avoid duplicate requests
- Respects NVD API rate limits (0.6s between requests)
- Tries CVSS v3.1 → v3.0 → v2.0 in order
- Limits to 30 CVEs per run to avoid excessive API calls
- Gracefully handles failures (shows "Not available" if lookup fails)

---

## 🏗️ Script Architecture

```
generate_sitrep_manual.py
├── CONFIG / TECHNOLOGY_STACK / TECH_STACK_DETAILS   ← Configuration
├── IntelCollector                                    ← Gathers threat data
│   ├── fetch_cisa_kev()                              ← CISA KEV catalog
│   ├── fetch_cisa_news()                             ← CISA alerts RSS
│   ├── fetch_vendor_advisories()                     ← 30+ vendor feeds
│   └── fetch_news_feeds()                            ← News RSS feeds
├── CVSS & Analysis Functions                         ← Enrichment layer
│   ├── _lookup_cvss() / _batch_lookup_cvss()         ← NVD API lookups
│   ├── _get_systems_in_environment()                 ← Hospital mapping
│   ├── _infer_attack_vector()                        ← Attack vector analysis
│   ├── _infer_impact()                               ← Healthcare impact
│   ├── _infer_remediation()                          ← Remediation steps
│   └── _infer_timeline()                             ← Urgency timelines
├── SITREPGenerator                                   ← Report generation
│   ├── prefetch_cvss_scores()                        ← Batch CVSS lookup
│   ├── generate_docx()                               ← Multi-page Word doc
│   │   ├── Page 1: Executive Summary
│   │   ├── Page 2: APPENDIX A (IMMEDIATE)
│   │   ├── Page 3: APPENDIX B (PRIORITY)
│   │   └── Page 4: APPENDIX C (ROUTINE)
│   └── generate_html()                               ← Email body HTML
├── Email Delivery                                    ← SMTP sending
└── main()                                            ← CLI entry point
    Steps: 1) Collect → 2) CVSS Lookup → 3) DOCX → 4) HTML → 5) Email
```

---

## 🔧 Troubleshooting

| Issue | Solution |
|-------|----------|
| `ModuleNotFoundError` | Run `pip install -r requirements.txt` |
| SSL errors | Update Python/certifi: `pip install --upgrade certifi` |
| Email not sending | Check SMTP credentials and firewall rules |
| Empty report | Check `sitrep_generator.log` for fetch errors |
| Slow execution | CVSS lookups add time; reduce CVEs or check network |
| CVSS scores "Not available" | NVD API may be rate-limited or CVE not yet indexed |

---

## 📊 How Threat Classification Works

1. **Intelligence is collected** from all configured sources in parallel
2. **Each item is checked** for relevance to the hospital's technology stack
3. **Severity is assigned** based on:
   - Keywords ("actively exploited", "zero-day", "critical", etc.)
   - Relevance to our specific technology stack
   - Healthcare-specific indicators
4. **CVSS scores are fetched** from NVD for all discovered CVEs
5. **Overall posture** is derived from the count and severity of findings:
   - **ELEVATED**: Any IMMEDIATE items, or 3+ PRIORITY items
   - **GUARDED**: 1-2 PRIORITY items
   - **LOW**: Only ROUTINE items
6. **Appendices are generated** with detailed technical breakdowns mapped to hospital systems

---

## 📜 License

Internal use only — Evergreen Healthcare Security Operations.
