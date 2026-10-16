# NASA Electronic Institutional Review Board (`eirb.jsc.nasa.gov`) — Pre-Rendered Mirror & Sitemap Generator

Automated harvester, static HTML pre-renderer, document mirror, and XML sitemap generator for **NASA's Electronic Institutional Review Board (eIRB)** portal at Lyndon B. Johnson Space Center (`https://eirb.jsc.nasa.gov`).

## Why Pre-Rendering Is Needed for RAG Ingestion

`https://eirb.jsc.nasa.gov` is a Huron Click Portal application hosted behind JSC's firewall (`*.jsc.nasa.gov`):

1. **JSC Firewall Restriction (`*.jsc.nasa.gov`)**: Direct connections from AWS-hosted Onyx indexers (`onyx.nasa.gov`) to `eirb.jsc.nasa.gov:443` time out due to JSC perimeter firewall rules.
2. **Dynamic Huron Click Codex REST APIs**: The **Library** tab (`55` IRB templates, SOPs, conflict of interest forms, and help guides across 7 categories) and **Related Links** tab (`15` regulatory and historical links across 3 categories) are populated asynchronously via JavaScript calls to `/EIRB/app/portal/codex/documentscomponent/...` and `/EIRB/app/portal/codex/webpagecomponent/...`.
3. **Visual Workflow Diagrams Transcribed to Semantic HTML**: Key procedural diagrams (`Slide1.png` — *IRB Submission Roadmap* and `statesofeirb.jpg` — *eIRB States*) are transcribed directly into structured HTML step-by-step lists and tables so every workflow stage and pre-review gate is searchable.
4. **Full-Text Document Mirroring (`54` PDFs & Word `.docx` Templates)**: Every downloadable `.docx` template (including Word `<w:sdt>` Content Controls) and PDF SOP/guide is mirrored both in its native format (`/EIRB/sd/resource/doc/<DOC_ID>/<filename>`) and as a rich semantic HTML document page (`/docs/<DOC_ID>`) for zero-timeout Playwright ingestion.

---

## Published Sitemaps & URL Lists

| File | URL / Path | Description |
| :--- | :--- | :--- |
| **`eirb_sitemap.xml`** | `https://raw.githubusercontent.com/Oht8wooWi8yait9n/eirb/refs/heads/main/eirb_sitemap.xml` | Primary XML sitemap (**77 URLs**: 18 pre-rendered portal routes + 54 full-text IRB document pages + 5 external NASA/NODIS directives) |
| **`sitemap.xml`** | `https://oht8woowi8yait9n.github.io/eirb/sitemap.xml` | Root alias of `eirb_sitemap.xml` |
| **`eirb_direct_sitemap.xml`** | `https://oht8woowi8yait9n.github.io/eirb/eirb_direct_sitemap.xml` | Direct XML sitemap pointing to `https://eirb.jsc.nasa.gov/...`, `https://www.nasa.gov/...`, and `https://nodis3.gsfc.nasa.gov/...` |
| **`eirb_urls.txt`** | `https://oht8woowi8yait9n.github.io/eirb/eirb_urls.txt` | Plain-text URL list (mirror + linked NASA references) |
| **`eirb_direct_urls.txt`** | `https://oht8woowi8yait9n.github.io/eirb/eirb_direct_urls.txt` | Plain-text URL list (direct NASA URLs) |
| **`eirb_catalog.json`** | `https://oht8woowi8yait9n.github.io/eirb/eirb_catalog.json` | Complete metadata catalog of all 55 Library documents, 15 Related Links, and SHA-256 page hashes for zero-churn CI/CD |

---

## Pre-Rendered Portal Container Pages (18 Routes)

Each Huron Click Portal Container page is generated at both `<route>/index.html` and `<route>.html`:

- `/` — **Home** (`4E9439392D8EF04BB83423B71F93D313`): Portal Overview, Announcements, International Partner Contacts, and Complete Index
- `/library` — **Library** (`3274673DA99811EEAD81FF18B24EF50A`): All 55 Official NASA IRB Documents Across 7 Categories (Informed Consent Templates, Research Protocol Templates, Conflict of Interest, Help Guides, Standard Operating Procedures, Clinical Trials, and Miscellaneous)
- `/nasa-irb` — **NASA IRB** (`9AD98B7EE42E28489071D16CE385ACB1`): NASA IRB Overview, Jurisdiction, Regulations & Submission Guidance
- `/irb-submission-roadmap` — **IRB Submission Roadmap** (`7729244CCC2411EEAE81FF18B24EF50A`): 10-Step Protocol Review Process (including full semantic transcription of `Slide1.png`)
- `/types-of-submissions` — **Types of Submissions** (`30F1298FA0F411EEAD81FF18B24EF50A`): Initial Study, Modification (MOD), Continuing Review (CR), Reportable New Information (RNI), and Study Closure
- `/eirb-states` — **eIRB States** (`B6E6C13CA11D11EEAD81FF18B24EF50A`): Submission State Workflow & Transitions (including full semantic transcription of `statesofeirb.jpg`)
- `/cooperative-research` — **Cooperative Research** (`52599855CC1011EEAE81FF18B24EF50A`): Single IRB Reliance (14 CFR 1230.114 / FWA00019876) & Reliance Agreements
- `/exempt-5-activities` — **Exempt #5 Activities** (`A4E8064D4744874F8A2BDD417F6B8807`): Public Service or Benefit Program Exemption Guidance
- `/medical-monitoring-level` — **Medical Monitoring Level** (`5579A4F3DA7411EEB081FF18B24EF50A`): MML Guidelines for JSC & Spaceflight Human Research
- `/additional-information` — **Additional Information** (`401D08D3A98E11EEAD81FF18B24EF50A`): CITI Training, COI, Risk Level, Merit & Safety Review, Blood Volume Limits & Subject Screening
- `/hrmrb` — **HRMRB** (`7C95F4A0A99611EEAD81FF18B24EF50A`): Human Research Multilateral Review Board (ISS International Partners: CSA, ESA, JAXA, NASA, Roscosmos)
- `/ahrcb` — **AHRCB** (`8BEB7DB54C8411F0E581FF18B24EF50A`): Artemis Human Research Coordination Board
- `/contact-us` — **Contact Us** (`A7F2212BCC1C11EEAE81FF18B24EF50A`): NASA IRB & HRMRB Office Support Contacts (`jsc-irb@mail.nasa.gov`)
- `/report-a-concern` — **Report a Concern** (`CCCE787EA9C111EEAD81FF18B24EF50A`): Research Participant Protection & Office of Research Assurance Contacts
- `/calendar-of-events` — **Calendar of Events** (`3491E9BBA9BC11EEAD81FF18B24EF50A`): 2025/2026 NASA IRB & HRMRB/AHRCB Meeting Schedules & Submission Deadlines
- `/related-links` — **Related Links** (`63B0D080C0B8994880C9DD20982017B3`): All 15 NASA Specific, General Regulatory, and Historical Bioethics Links
- `/registration` — **Registration** (`2660DD9607A4FB4ABEBF7928381E5945`): Self-Registration Form & Instructions for Requesting a NASA eIRB Account
- `/login` — **Login** (`0A7646F3B149874E902185897C144551`): Login Portal & Privacy Act of 1974 Notice

---

## Onyx Web Connector & Proxy URL Rewrite Rule

Use the raw GitHub sitemap URL in Onyx's Web Connector (**Sitemap** mode):

```text
https://raw.githubusercontent.com/Oht8wooWi8yait9n/eirb/refs/heads/main/eirb_sitemap.xml
```

To rewrite the GitHub Pages mirror URLs back to canonical `https://eirb.jsc.nasa.gov` URLs in Onyx, add this prefix mapping in `url_rewrites.json`:

```json
"https://oht8woowi8yait9n.github.io/eirb": "https://eirb.jsc.nasa.gov"
```

---

## Local Execution

```bash
pip install -r requirements.txt
python3 generate_eirb_sitemap.py
```
