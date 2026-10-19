#!/usr/bin/env python3
"""
NASA Electronic Institutional Review Board (eirb.jsc.nasa.gov)
Harvester, Static HTML/Document Mirror & Sitemap Generator
==============================================================
Dynamically harvests:
  1. All 18 public Huron Click Portal Container pages from https://eirb.jsc.nasa.gov
     (plus automatic discovery of any newly linked Container OIDs).
  2. All 7 Huron Click `DocumentsComponent` JSON API endpoints on the Library page
     (`/EIRB/app/portal/codex/documentscomponent/<source_entity>`) covering all 55
     official NASA IRB documents (Informed Consent Templates, Research Protocol
     Templates, Conflict of Interest Forms, eIRB Help Guides, Standard Operating
     Procedures [ORA-101 through ORA-990], Clinical Trials Policy/Guidance, and
     Miscellaneous questionnaires/medical monitoring tables).
  3. All 3 Huron Click `WebPageComponent` JSON API endpoints on the Related Links page
     (`/EIRB/app/portal/codex/webpagecomponent/<source_entity>`) covering NASA Specific
     Links, General Regulatory Links, and Historical Links.
  4. All 54 downloadable PDF and Word (.docx) documents and embedded portal diagram
     images (`Slide1.png`, `statesofeirb.jpg`, `NASA-Logo-Large.png`), mirroring the
     raw files under `/EIRB/sd/resource/doc/<DOC_ID>/<filename>` and pre-rendering
     full-text HTML pages at `/docs/<DOC_ID>` so Onyx's Web Connector indexes 100%
     of both PDF SOPs/Guides and Word (.docx) Consent/Protocol Templates.

Generates:
  - Pre-rendered, self-contained semantic HTML pages for all 18 portal routes and all
    54 Library documents, with zero external dependencies on firewall-blocked
    `eirb.jsc.nasa.gov` subresources.
  - Mirrored PDFs, DOCX files, and diagram images at their exact `/EIRB/...` paths.
  - `eirb_sitemap.xml` & `sitemap.xml` (GitHub Pages mirror URLs + verified NASA policy URLs)
  - `eirb_direct_sitemap.xml` (100% direct live `eirb.jsc.nasa.gov` + NASA reference URLs)
  - `eirb_urls.txt` & `eirb_direct_urls.txt`
  - `eirb_catalog.json` (structured metadata & content hashes for zero-churn CI/CD)
"""

import hashlib
import html
import http.cookiejar
import io
import json
import os
import re
import shutil
import ssl
import subprocess
import sys
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from datetime import datetime, timezone

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SITE_ORIGIN = "https://eirb.jsc.nasa.gov"
GH_PAGES_BASE = os.environ.get(
    "GH_PAGES_BASE", "https://oht8woowi8yait9n.github.io/eirb"
).rstrip("/")

CATALOG_FILE = os.path.join(BASE_DIR, "eirb_catalog.json")
SITEMAP_FILE = os.path.join(BASE_DIR, "eirb_sitemap.xml")
SITEMAP_ROOT_FILE = os.path.join(BASE_DIR, "sitemap.xml")
DIRECT_SITEMAP_FILE = os.path.join(BASE_DIR, "eirb_direct_sitemap.xml")
URLS_FILE = os.path.join(BASE_DIR, "eirb_urls.txt")
DIRECT_URLS_FILE = os.path.join(BASE_DIR, "eirb_direct_urls.txt")

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36 (NASA-Onyx-RAG-Indexer/1.0)"
)

# Canonical ordered list of the 18 public eIRB Container pages
PORTAL_PAGES = [
    {
        "oid": "4E9439392D8EF04BB83423B71F93D313",
        "slug": "",
        "route": "/",
        "title": "Home — NASA Electronic Institutional Review Board (eIRB)",
        "short_title": "Home",
        "section": "Portal Home",
        "priority": "1.0",
    },
    {
        "oid": "3274673DA99811EEAD81FF18B24EF50A",
        "slug": "library",
        "route": "/library",
        "title": "Library — Consent & Protocol Templates, COI Forms, Help Guides & SOPs | NASA eIRB",
        "short_title": "Library",
        "section": "IRB Document Library",
        "priority": "0.95",
    },
    {
        "oid": "9AD98B7EE42E28489071D16CE385ACB1",
        "slug": "nasa-irb",
        "route": "/nasa-irb",
        "title": "NASA IRB Overview, Regulations & Submission Guidance | NASA eIRB",
        "short_title": "NASA IRB",
        "section": "NASA IRB",
        "priority": "0.95",
    },
    {
        "oid": "7729244CCC2411EEAE81FF18B24EF50A",
        "slug": "irb-submission-roadmap",
        "route": "/irb-submission-roadmap",
        "title": "IRB Submission Roadmap — 10-Step Protocol Review Process | NASA eIRB",
        "short_title": "IRB Submission Roadmap",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "30F1298FA0F411EEAD81FF18B24EF50A",
        "slug": "types-of-submissions",
        "route": "/types-of-submissions",
        "title": "Types of Submissions — Initial, Modification, Continuing Review, RNI & Closure | NASA eIRB",
        "short_title": "Types of Submissions",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "B6E6C13CA11D11EEAD81FF18B24EF50A",
        "slug": "eirb-states",
        "route": "/eirb-states",
        "title": "eIRB States — Follow Your eIRB Submission Workflow | NASA eIRB",
        "short_title": "eIRB States",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "52599855CC1011EEAE81FF18B24EF50A",
        "slug": "cooperative-research",
        "route": "/cooperative-research",
        "title": "Cooperative Research & Single IRB Reliance (14 CFR 1230.114 / FWA00019876) | NASA eIRB",
        "short_title": "Cooperative Research",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "A4E8064D4744874F8A2BDD417F6B8807",
        "slug": "exempt-5-activities",
        "route": "/exempt-5-activities",
        "title": "Exempt #5 Activities — Public Service or Benefit Programs | NASA eIRB",
        "short_title": "Exempt #5 Activities",
        "section": "NASA IRB",
        "priority": "0.85",
    },
    {
        "oid": "5579A4F3DA7411EEB081FF18B24EF50A",
        "slug": "medical-monitoring-level",
        "route": "/medical-monitoring-level",
        "title": "Medical Monitoring Level (MML) Guidelines — JSC & Spaceflight | NASA eIRB",
        "short_title": "Medical Monitoring Level",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "401D08D3A98E11EEAD81FF18B24EF50A",
        "slug": "additional-information",
        "route": "/additional-information",
        "title": "Additional Information — CITI Training, COI, Risk Level, Merit & Safety Review, Blood Volume & Screening | NASA eIRB",
        "short_title": "Additional Information",
        "section": "NASA IRB",
        "priority": "0.90",
    },
    {
        "oid": "7C95F4A0A99611EEAD81FF18B24EF50A",
        "slug": "hrmrb",
        "route": "/hrmrb",
        "title": "HRMRB — Human Research Multilateral Review Board (ISS International Partners) | NASA eIRB",
        "short_title": "HRMRB",
        "section": "International Review Boards",
        "priority": "0.90",
    },
    {
        "oid": "8BEB7DB54C8411F0E581FF18B24EF50A",
        "slug": "ahrcb",
        "route": "/ahrcb",
        "title": "AHRCB — Artemis Human Research Coordination Board | NASA eIRB",
        "short_title": "AHRCB",
        "section": "International Review Boards",
        "priority": "0.90",
    },
    {
        "oid": "A7F2212BCC1C11EEAE81FF18B24EF50A",
        "slug": "contact-us",
        "route": "/contact-us",
        "title": "Contact Us — NASA IRB & HRMRB Support | NASA eIRB",
        "short_title": "Contact Us",
        "section": "Support & Governance",
        "priority": "0.85",
    },
    {
        "oid": "CCCE787EA9C111EEAD81FF18B24EF50A",
        "slug": "report-a-concern",
        "route": "/report-a-concern",
        "title": "Report a Concern — Research Participant Protection & Office of Research Assurance | NASA eIRB",
        "short_title": "Report a Concern",
        "section": "Support & Governance",
        "priority": "0.85",
    },
    {
        "oid": "3491E9BBA9BC11EEAD81FF18B24EF50A",
        "slug": "calendar-of-events",
        "route": "/calendar-of-events",
        "title": "Calendar of Events — NASA IRB & HRMRB/AHRCB Meeting Schedules & Submission Deadlines | NASA eIRB",
        "short_title": "Calendar of Events",
        "section": "Schedules & Deadlines",
        "priority": "0.90",
    },
    {
        "oid": "63B0D080C0B8994880C9DD20982017B3",
        "slug": "related-links",
        "route": "/related-links",
        "title": "Related Links — NASA Specific, General Regulatory & Historical Links | NASA eIRB",
        "short_title": "Related Links",
        "section": "Regulatory References",
        "priority": "0.85",
    },
    {
        "oid": "2660DD9607A4FB4ABEBF7928381E5945",
        "slug": "registration",
        "route": "/registration",
        "title": "Self Registration — Request a New NASA eIRB User Account | NASA eIRB",
        "short_title": "Registration",
        "section": "Account Access",
        "priority": "0.80",
    },
    {
        "oid": "0A7646F3B149874E902185897C144551",
        "slug": "login",
        "route": "/login",
        "title": "Login & Privacy Act of 1974 Notice | NASA eIRB",
        "short_title": "Login",
        "section": "Account Access",
        "priority": "0.75",
    },
]

OID_TO_PAGE = {p["oid"].upper(): p for p in PORTAL_PAGES}

EXTERNAL_REFERENCE_URLS = [
    {
        "url": "https://www.nasa.gov/institutional-review-board/",
        "title": "NASA Institutional Review Board (IRB) Official Website",
        "priority": "0.85",
    },
    {
        "url": "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPD&c=7100&s=8F",
        "title": "NASA Policy Directive NPD 7100.8F: Protection of Human Research Subjects",
        "priority": "0.85",
    },
    {
        "url": "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPR&c=7100&s=1B",
        "title": "NASA Procedural Requirements NPR 7100.1B: Protection of Human Research Subjects",
        "priority": "0.85",
    },
    {
        "url": "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPD&c=1382&s=17J",
        "title": "NASA Policy Directive NPD 1382.17J: NASA Privacy Policy",
        "priority": "0.80",
    },
    {
        "url": "https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPR&c=1382&s=1B",
        "title": "NASA Procedural Requirements NPR 1382.1B: NASA Privacy Procedural Requirements",
        "priority": "0.80",
    },
]

PORTAL_IMAGES = [
    {
        "remote_url": "https://eirb.jsc.nasa.gov/EIRB/sd/Doc/0/KNSB3KEC4G8UTBK1VSCB4JNL0A/Slide1.png",
        "rel_path": "EIRB/sd/Doc/0/KNSB3KEC4G8UTBK1VSCB4JNL0A/Slide1.png",
    },
    {
        "remote_url": "https://eirb.jsc.nasa.gov/EIRB/sd/Doc/0/JR4I09L9HK8UTBC1VSCB4JNL0A/statesofeirb.jpg",
        "rel_path": "EIRB/sd/Doc/0/JR4I09L9HK8UTBC1VSCB4JNL0A/statesofeirb.jpg",
    },
    {
        "remote_url": "https://eirb.jsc.nasa.gov/EIRB/sd/Doc/0/7E127INJPG8UVNC1VSCB4JNL0A/NASA-Logo-Large.png",
        "rel_path": "EIRB/sd/Doc/0/7E127INJPG8UVNC1VSCB4JNL0A/NASA-Logo-Large.png",
    },
]


def container_direct_url(oid: str) -> str:
    return (
        f"{SITE_ORIGIN}/EIRB/sd/Rooms/DisplayPages/LayoutInitial"
        f"?Container=com.webridge.entity.Entity%5BOID%5B{oid}%5D%5D"
    )


def write_if_changed(path: str, content: str) -> bool:
    """Write text file only if content actually changed to prevent unnecessary Git churn."""
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                if f.read() == content:
                    return False
        except Exception:
            pass
    with open(path, "w", encoding="utf-8") as f:
        f.write(content)
    return True


def write_bytes_if_changed(path: str, data: bytes) -> bool:
    """Write binary file only if bytes actually changed."""
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    if os.path.exists(path):
        try:
            with open(path, "rb") as f:
                if f.read() == data:
                    return False
        except Exception:
            pass
    with open(path, "wb") as f:
        f.write(data)
    return True


def load_existing_catalog() -> dict:
    if os.path.exists(CATALOG_FILE):
        try:
            with open(CATALOG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def load_existing_sitemap_dates(path: str) -> dict:
    dates = {}
    if not os.path.exists(path):
        return dates
    try:
        tree = ET.parse(path)
        root = tree.getroot()
        for elem in root.iter():
            if elem.tag.endswith("url"):
                loc = None
                lastmod = None
                for child in elem:
                    if child.tag.endswith("loc") and child.text:
                        loc = child.text.strip()
                    elif child.tag.endswith("lastmod") and child.text:
                        lastmod = child.text.strip()
                if loc and lastmod:
                    dates[loc] = lastmod
    except Exception:
        pass
    return dates


def build_opener():
    ctx = ssl._create_unverified_context()
    cj = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(cj),
        urllib.request.HTTPSHandler(context=ctx),
    )
    opener.addheaders = [("User-Agent", USER_AGENT)]
    return opener


def fetch_url_bytes(opener, url: str, headers: dict | None = None, timeout: int = 30) -> bytes:
    req = urllib.request.Request(url, headers=headers or {"User-Agent": USER_AGENT})
    with opener.open(req, timeout=timeout) as resp:
        return resp.read()


def extract_docx_html(docx_bytes: bytes) -> tuple[str, str]:
    """
    Parse a Word .docx file (ZIP archive containing word/document.xml) using Python's
    standard library and return (html_content, plain_text).
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    blocks_html = []
    plain_lines = []

    try:
        with zipfile.ZipFile(io.BytesIO(docx_bytes)) as zf:
            xml_names = []
            for candidate in ["word/header1.xml", "word/header2.xml", "word/document.xml", "word/footer1.xml"]:
                if candidate in zf.namelist():
                    xml_names.append(candidate)
            for name in xml_names:
                raw_xml = zf.read(name)
                root = ET.fromstring(raw_xml)
                body = root.find("w:body", ns) if name == "word/document.xml" else root
                if body is None:
                    body = root

                def walk_node(node):
                    for child in node:
                        tag = child.tag.split("}")[-1]
                        if tag == "p":
                            texts = [t.text for t in child.findall(".//w:t", ns) if t.text]
                            line = "".join(texts).strip()
                            if not line:
                                continue
                            plain_lines.append(line)
                            pstyle = child.find(".//w:pStyle", ns)
                            style_val = (
                                pstyle.attrib.get(f"{{{ns['w']}}}val", "")
                                if pstyle is not None
                                else ""
                            ).lower()
                            if "heading1" in style_val or style_val == "title":
                                blocks_html.append(f"<h3>{html.escape(line)}</h3>")
                            elif "heading2" in style_val or "heading3" in style_val:
                                blocks_html.append(f"<h4>{html.escape(line)}</h4>")
                            elif "list" in style_val:
                                blocks_html.append(f"<ul><li>{html.escape(line)}</li></ul>")
                            else:
                                blocks_html.append(f"<p>{html.escape(line)}</p>")
                        elif tag == "tbl":
                            rows_html = []
                            for tr in child.findall(".//w:tr", ns):
                                cells_html = []
                                row_texts = []
                                for tc in tr.findall(".//w:tc", ns):
                                    cell_pars = []
                                    for p in tc.findall(".//w:p", ns):
                                        pts = [t.text for t in p.findall(".//w:t", ns) if t.text]
                                        s = "".join(pts).strip()
                                        if s:
                                            cell_pars.append(s)
                                    cell_str = " ".join(cell_pars).strip()
                                    if cell_str:
                                        row_texts.append(cell_str)
                                    cells_html.append(f"<td>{html.escape(cell_str)}</td>")
                                if row_texts:
                                    plain_lines.append(" | ".join(row_texts))
                                    rows_html.append("<tr>" + "".join(cells_html) + "</tr>")
                            if rows_html:
                                blocks_html.append(
                                    '<div class="table-wrap"><table class="eirb-table"><tbody>'
                                    + "".join(rows_html)
                                    + "</tbody></table></div>"
                                )
                        elif tag not in ("sdtPr", "sdtEndPr", "pPr", "tblPr", "sectPr"):
                            walk_node(child)

                walk_node(body)
    except Exception as e:
        return f"<p><em>Unable to parse DOCX XML: {html.escape(str(e))}</em></p>", ""

    # Merge adjacent <ul><li>...</li></ul> into a single <ul>
    merged_html = "\n".join(blocks_html)
    merged_html = re.sub(r"</ul>\s*<ul>", "\n", merged_html)
    return merged_html, "\n".join(plain_lines)


def extract_pdf_html(pdf_path: str) -> tuple[str, str]:
    """
    Extract text from a PDF file using `/usr/bin/pdftotext -layout` (preferred for deterministic
    layout preservation) or `pypdf` as fallback, and format into clean HTML paragraphs.
    """
    text = ""
    if shutil.which("pdftotext"):
        try:
            res = subprocess.run(
                ["pdftotext", "-layout", pdf_path, "-"],
                capture_output=True,
                text=True,
                timeout=30,
                check=False,
            )
            if res.returncode == 0 and res.stdout:
                text = res.stdout
        except Exception:
            pass

    if not text.strip():
        try:
            import pypdf  # type: ignore

            reader = pypdf.PdfReader(pdf_path)
            pages_text = []
            for idx, page in enumerate(reader.pages, 1):
                pt = (page.extract_text() or "").strip()
                if pt:
                    pages_text.append(pt)
            text = "\n\n".join(pages_text)
        except Exception:
            pass

    text = text.replace("\x0c", "\n\n--- Page Break ---\n\n").strip()
    if not text:
        return "<p><em>No extractable text layer found in PDF. Use the download link above to view the original PDF document.</em></p>", ""

    paragraphs = []
    for block in re.split(r"\n\s*\n", text):
        block = re.sub(r"^---\s*Page\s+\d+\s*---\s*", "", block.strip())
        cleaned = "\n".join(line.rstrip() for line in block.splitlines() if line.strip()).strip()
        if not cleaned:
            continue
        if cleaned == "--- Page Break ---" or re.match(r"^--- Page \d+ ---$", cleaned):
            continue
        # If block looks like a short uppercase/numbered section heading
        if len(cleaned) < 100 and "\n" not in cleaned and (
            re.match(r"^(\d+\.0|\d+\.|PURPOSE|SCOPE|POLICY|DEFINITIONS|RESPONSIBILITIES|PROCEDURES?|REFERENCES|MATERIALS|OVERVIEW)\b", cleaned, re.I)
        ):
            paragraphs.append(f"<h3>{html.escape(cleaned)}</h3>")
        else:
            paragraphs.append(f"<p>{html.escape(cleaned)}</p>")

    return "\n".join(paragraphs), text


def harvest_library_documents(opener, lib_html: str, lib_oid: str) -> tuple[list[dict], dict[str, dict]]:
    """
    Query all 7 DocumentsComponent API endpoints on the Library page and return:
      - sections: list of {section_id, source_entity, title, documents: [...]}
      - doc_by_id: dict mapping DOC_ID -> document metadata dict
    """
    sections = []
    doc_by_id = {}

    for m in re.finditer(r'<div[^>]*id="(DocumentsComponent_[^"]+)_data"[^>]*>', lib_html):
        comp_id = m.group(1)
        tag = m.group(0)
        se_m = re.search(r'data-source-entity="([^"]+)"', tag)
        if not se_m:
            continue
        se = se_m.group(1)

        start = max(0, m.start() - 1200)
        snippet = lib_html[start : m.start()]
        sec_names = re.findall(r'<span class="textControl"[^>]*>(.*?)</span>', snippet, re.S)
        sec_title = html.unescape(sec_names[-1].strip()) if sec_names else comp_id

        api_url = (
            f"{SITE_ORIGIN}/EIRB/app/portal/codex/documentscomponent/{urllib.parse.quote(se)}"
            f"?currentContainer=com.webridge.entity.Entity[OID[{lib_oid}]]"
            f"&skip=0&top=100&forceFullCount=true&currentContainerOid={lib_oid}"
        )
        raw = fetch_url_bytes(opener, api_url, headers={"Accept": "application/json", "User-Agent": USER_AGENT})
        payload = json.loads(raw.decode("utf-8", errors="ignore"))
        rows = payload.get("data", [])

        docs = []
        for r in rows:
            doc_id = (r.get("ID") or "").strip()
            custom_url = (r.get("EncodeCustomUrl()") or r.get("$url$") or "").strip()
            parsed_u = urllib.parse.urlparse(custom_url)
            rel_decoded_path = urllib.parse.unquote(parsed_u.path).lstrip("/")
            ext = os.path.splitext(rel_decoded_path)[1].lower()
            is_video = ext in (".mp4", ".mov", ".avi", ".mkv")

            mirror_file_url = (
                custom_url.replace(SITE_ORIGIN, GH_PAGES_BASE)
                if custom_url.startswith(SITE_ORIGIN) and not is_video
                else custom_url
            )
            doc_info = {
                "id": doc_id,
                "name": (r.get("name") or doc_id).strip(),
                "section": sec_title,
                "version": str(r.get("version") or ""),
                "dateCreated": (r.get("dateCreated") or "").strip(),
                "dateModified": (r.get("dateModified") or "").strip(),
                "author": (r.get("author") or "").strip(),
                "fileSize": (r.get("getFileSize()") or "").strip(),
                "description": (r.get("description") or "").strip(),
                "directUrl": custom_url,
                "mirrorFileUrl": mirror_file_url,
                "relFilePath": rel_decoded_path,
                "ext": ext,
                "isVideo": is_video,
                "docPageRoute": f"/docs/{doc_id}",
                "docPageMirrorUrl": f"{GH_PAGES_BASE}/docs/{doc_id}",
            }
            docs.append(doc_info)
            if doc_id:
                doc_by_id[doc_id] = doc_info

        sections.append(
            {
                "componentId": comp_id,
                "sourceEntity": se,
                "title": sec_title,
                "documents": docs,
            }
        )

    return sections, doc_by_id


def harvest_related_links(opener, rel_html: str, rel_oid: str) -> list[dict]:
    """
    Query all 3 WebPageComponent API endpoints on the Related Links page.
    """
    sections = []
    for m in re.finditer(r'<div[^>]*id="([0-9A-F]{32})_data"[^>]*>', rel_html):
        se = m.group(1)
        start = max(0, m.start() - 1200)
        snippet = rel_html[start : m.start()]
        sec_names = re.findall(r'<span class="textControl"[^>]*>(.*?)</span>', snippet, re.S)
        sec_title = html.unescape(sec_names[-1].strip()) if sec_names else se

        api_url = (
            f"{SITE_ORIGIN}/EIRB/app/portal/codex/webpagecomponent/{urllib.parse.quote(se)}"
            f"?currentContainer=com.webridge.entity.Entity[OID[{rel_oid}]]"
            f"&skip=0&top=100&forceFullCount=true&currentContainerOid={rel_oid}"
        )
        raw = fetch_url_bytes(opener, api_url, headers={"Accept": "application/json", "User-Agent": USER_AGENT})
        payload = json.loads(raw.decode("utf-8", errors="ignore"))
        rows = payload.get("data", [])

        links = []
        for r in rows:
            links.append(
                {
                    "id": (r.get("ID") or "").strip(),
                    "name": (r.get("name") or "").strip(),
                    "url": (r.get("$url$") or "").strip(),
                    "description": (r.get("description") or "").strip(),
                    "dateModified": (r.get("dateModified") or "").strip(),
                }
            )
        sections.append(
            {
                "sourceEntity": se,
                "title": sec_title,
                "links": links,
            }
        )
    return sections


def render_library_sections_html(sections: list[dict]) -> str:
    """Render all 7 Library document sections as clean semantic HTML tables."""
    parts = [
        '<div class="library-overview-banner">',
        "  <p>The <strong>NASA eIRB Library</strong> contains official Informed Consent Templates, "
        "Research Protocol Templates, Conflict of Interest (COI) Disclosure Forms, eIRB Help Guides, "
        "Standard Operating Procedures (SOPs), ClinicalTrials.gov Guidance, and Questionnaires. "
        "Click any document title below to read its full indexed text, or use the direct file link to "
        "download the original Word (<code>.docx</code>) or PDF file.</p>",
        "</div>",
    ]
    for sec in sections:
        parts.append(f'<section class="card library-section" id="sec-{html.escape(sec["sourceEntity"])}">')
        parts.append(f'  <h2>{html.escape(sec["title"])} ({len(sec["documents"])} Documents)</h2>')
        parts.append('  <div class="table-wrap"><table class="eirb-table">')
        parts.append(
            "    <thead><tr>"
            "<th>Document Title</th>"
            "<th>Description &amp; Notes</th>"
            "<th>Doc ID</th>"
            "<th>Version</th>"
            "<th>Last Modified</th>"
            "<th>Size</th>"
            "<th>Download</th>"
            "</tr></thead><tbody>"
        )
        for d in sec["documents"]:
            desc_html = html.escape(d["description"]).replace("\r\n", "<br>").replace("\n", "<br>") if d["description"] else "—"
            if d["isVideo"]:
                title_cell = f'<strong>{html.escape(d["name"])}</strong>'
                dl_cell = f'<a href="{html.escape(d["directUrl"])}" target="_blank" rel="noopener">NASA Video ({html.escape(d["ext"].upper().lstrip("."))})</a>'
            else:
                title_cell = (
                    f'<a href="{GH_PAGES_BASE}{html.escape(d["docPageRoute"])}">'
                    f'<strong>{html.escape(d["name"])}</strong></a>'
                )
                dl_cell = (
                    f'<a href="{html.escape(d["mirrorFileUrl"])}">Mirror ({html.escape(d["ext"].upper().lstrip("."))})</a>'
                    f' | <a href="{html.escape(d["directUrl"])}" target="_blank" rel="noopener">NASA Live</a>'
                )
            parts.append(
                "    <tr>"
                f"<td>{title_cell}</td>"
                f"<td>{desc_html}</td>"
                f"<td><code>{html.escape(d['id'])}</code></td>"
                f"<td>{html.escape(d['version'])}</td>"
                f"<td>{html.escape(d['dateModified'])}</td>"
                f"<td>{html.escape(d['fileSize'])}</td>"
                f"<td>{dl_cell}</td>"
                "</tr>"
            )
        parts.append("    </tbody></table></div>")
        parts.append("</section>")
    return "\n".join(parts)


def render_related_links_html(sections: list[dict]) -> str:
    """Render all 3 Related Links sections as clean semantic HTML tables."""
    parts = []
    for sec in sections:
        parts.append(f'<section class="card related-links-section" id="links-{html.escape(sec["sourceEntity"])}">')
        parts.append(f'  <h2>{html.escape(sec["title"])}</h2>')
        parts.append('  <div class="table-wrap"><table class="eirb-table">')
        parts.append("    <thead><tr><th>Resource / Policy Name</th><th>URL</th></tr></thead><tbody>")
        for lnk in sec["links"]:
            u = lnk["url"]
            # Also note if NPR 1382.1A moved to 1382.1B in NODIS
            extra = ""
            if "N_PR_1382_001A_" in u:
                extra = (
                    ' <br><small>(Current NODIS version: '
                    '<a href="https://nodis3.gsfc.nasa.gov/displayDir.cfm?t=NPR&amp;c=1382&amp;s=1B">'
                    "NPR 1382.1B</a>)</small>"
                )
            parts.append(
                "    <tr>"
                f'<td><strong><a href="{html.escape(u)}" target="_blank" rel="noopener">{html.escape(lnk["name"])}</a></strong>{extra}</td>'
                f'<td><a href="{html.escape(u)}" target="_blank" rel="noopener"><code>{html.escape(u)}</code></a></td>'
                "</tr>"
            )
        parts.append("    </tbody></table></div>")
        parts.append("</section>")
    return "\n".join(parts)


ROADMAP_TRANSCRIPTION_HTML = """
<section class="card roadmap-transcription">
  <h2>NASA IRB Submission Roadmap — Step-by-Step Workflow Transcription</h2>
  <p>The NASA IRB Submission Roadmap guides Principal Investigators (PIs) and research teams through the 10 sequential milestones required for human subjects research review and implementation:</p>
  <ol>
    <li><strong>CITI Training</strong>: Human subjects training is required for all engaged personnel. Training is valid for 3 years.</li>
    <li><strong>Develop Research Protocol</strong>: Prepare the study protocol using the appropriate NASA IRB, Center (Ames/Langley), or Exempt protocol template from the <a href="{GH_PAGES_BASE}/library">eIRB Library</a>.</li>
    <li><strong>Local Ethics</strong>: International PIs submit local ethics board approval with their NASA submission.</li>
    <li><strong>Finalize Materials</strong>: The IRB must review complete and final study materials, informed consent forms, recruitment materials, questionnaires, etc.</li>
    <li><strong>Timely Submission</strong>: Budget 6–8 weeks for review. The full NASA IRB board meets once per month (see <a href="{GH_PAGES_BASE}/calendar-of-events">Calendar of Events</a>). Be mindful of submission deadlines.</li>
    <li><strong>Reliance Process</strong>: IRB Reliance Acknowledgments must be established for all non-NASA personnel engaged in cooperative research (see <a href="{GH_PAGES_BASE}/cooperative-research">Cooperative Research</a>).</li>
    <li><strong>Clarify</strong>: Promptly respond to reviewer or coordinator clarification requests in eIRB to keep the review on track.</li>
    <li><strong>Ancillary Reviews</strong>: Complete required ancillary reviews including Safety Board, Crew Office, Flight Certification, Test Readiness Review (TRR), Radiation, and Medical Monitoring Level assignment.</li>
    <li><strong>HRMRB Review</strong>: Studies occurring on the International Space Station (ISS) or involving international space agency crew/subjects require review by the <a href="{GH_PAGES_BASE}/hrmrb">Human Research Multilateral Review Board (HRMRB)</a> (meets quarterly) or <a href="{GH_PAGES_BASE}/ahrcb">Artemis Human Research Coordination Board (AHRCB)</a>.</li>
    <li><strong>Implementation</strong>: <em>Note:</em> IRB approval does not guarantee approval from other review boards or that a study is selected for flight.</li>
  </ol>
</section>
"""

EIRB_STATES_TRANSCRIPTION_HTML = """
<section class="card states-workflow-summary">
  <h2>eIRB Workflow Diagram Progression</h2>
  <p><strong>Workflow Sequence:</strong> <code>Pre-Submission</code> &rarr; <code>Pre-Review</code> (with loop to/from <code>Clarification Requested</code>) &rarr; <code>IRB Review</code> (with loop to/from <code>Clarification Requested</code>) &rarr; <code>Post-Review</code> (with loop to/from <code>Modifications Required</code>) &rarr; <code>Review Complete</code>.</p>
</section>
"""


def clean_portal_main_html(
    raw_html: str,
    page_spec: dict,
    doc_by_id: dict[str, dict],
    library_sections: list[dict],
    related_links_sections: list[dict],
) -> str:
    """
    Extract the `<main>` content from a Huron Click Portal page, remove UI scripts,
    decorative layout GIFs, and duplicate sidebars, rewrite internal Container OID links
    and document links to local mirror URLs, and inject dynamic component tables.
    """
    m = re.search(r"<main\b[^>]*>(.*?)</main>", raw_html, re.S | re.I)
    body = m.group(1) if m else raw_html

    # Remove <script> and <style> blocks
    body = re.sub(r"<script\b.*?</script>", "", body, flags=re.S | re.I)
    body = re.sub(r"<style\b.*?</style>", "", body, flags=re.S | re.I)

    # Remove decorative layout images (spacer.gif, roundedCorner_*.gif)
    body = re.sub(
        r'<img[^>]+(?:spacer\.gif|roundedCorner_left\.gif|roundedCorner_right\.gif)[^>]*>',
        "",
        body,
        flags=re.I,
    )

    # Replace external logo URLs on HRMRB/AHRCB with clean text/alt badges or safe inline references
    # so Playwright never waits on third-party space agency servers (esa.int, jaxa.jp)
    def replace_img(match: re.Match) -> str:
        tag = match.group(0)
        src_m = re.search(r'src=["\']([^"\']+)["\']', tag, re.I)
        if not src_m:
            return tag
        src = src_m.group(1)
        if src.startswith("data:"):
            return tag
        for img_info in PORTAL_IMAGES:
            if img_info["rel_path"] in src:
                new_src = f"{GH_PAGES_BASE}/{img_info['rel_path']}"
                return tag.replace(src, new_src)
        if "eirb.jsc.nasa.gov" in src or "esamultimedia.esa.int" in src or "global.jaxa.jp" in src or "www.nasa.gov" in src:
            alt_m = re.search(r'alt=["\']([^"\']*)["\']', tag, re.I)
            alt_txt = alt_m.group(1).strip() if alt_m else ""
            return f'<span class="agency-logo-badge">{html.escape(alt_txt)}</span>' if alt_txt else ""
        return tag

    body = re.sub(r"<img\b[^>]*>", replace_img, body, flags=re.I)

    # Rewrite internal Container OID links and /resource/doc/<DOC_ID> links
    def replace_href(match: re.Match) -> str:
        prefix, url, suffix = match.group(1), match.group(2), match.group(3)
        oid_m = re.search(r"OID(?:\[|%5B)([0-9A-F]{32})(?:\]|%5D)", url, re.I)
        if oid_m:
            oid = oid_m.group(1).upper()
            if oid in OID_TO_PAGE:
                target_route = OID_TO_PAGE[oid]["route"]
                new_url = f"{GH_PAGES_BASE}/" if target_route == "/" else f"{GH_PAGES_BASE}{target_route}"
                return f"{prefix}{new_url}{suffix}"
        doc_m = re.search(r"/resource/doc/(DOC[0-9A-F]+)/", url, re.I)
        if doc_m:
            doc_id = doc_m.group(1).upper()
            if doc_id in doc_by_id:
                return f"{prefix}{GH_PAGES_BASE}/docs/{doc_id}{suffix}"
        if url.startswith("../../#"):
            anchor = url.split("#", 1)[1]
            return f"{prefix}#{anchor}{suffix}"
        return match.group(0)

    body = re.sub(r'(href=["\'])([^"\']+)(["\'])', replace_href, body, flags=re.I)

    # Remove Huron Click screen-reader / drag-and-drop boilerplate
    body = re.sub(
        r"<span class=['\"]dragAndDropHint['\"][^>]*>.*?</span>\s*</span>",
        "",
        body,
        flags=re.S | re.I,
    )
    body = re.sub(
        r'<div class="RoomComponentTitleArea">.*?</div>',
        "",
        body,
        flags=re.S | re.I,
    )
    body = re.sub(
        r'(name=["\']_webrLocale["\']\s+value=["\'])en-us(["\'])',
        r"\1en-US\2",
        body,
        flags=re.I,
    )
    body = re.sub(
        r'(action=["\'][^"\']*?)Entity\[OID\[([^\]]+)\]\]',
        r"\1Entity%5BOID%5B\2%5D%5D",
        body,
    )

    # Append pre-rendered dynamic components for Library, Related Links, Roadmap, and eIRB States
    slug = page_spec["slug"]
    if slug == "library":
        body += "\n" + render_library_sections_html(library_sections)
    elif slug == "related-links":
        body += "\n" + render_related_links_html(related_links_sections)
    elif slug == "irb-submission-roadmap":
        body += "\n" + ROADMAP_TRANSCRIPTION_HTML.format(GH_PAGES_BASE=GH_PAGES_BASE)
    elif slug == "eirb-states":
        body += "\n" + EIRB_STATES_TRANSCRIPTION_HTML

    return body.strip()


def render_page_shell(
    title: str,
    route_path: str,
    direct_url: str,
    section_label: str,
    body_html: str,
    extra_meta_html: str = "",
) -> str:
    """Wrap page content in a clean, accessible NASA eIRB semantic HTML5 document."""
    nav_links = []
    for p in PORTAL_PAGES:
        href = f"{GH_PAGES_BASE}/" if p["route"] == "/" else f"{GH_PAGES_BASE}{p['route']}"
        active_cls = ' class="active"' if p["route"] == route_path else ""
        nav_links.append(f'<a href="{href}"{active_cls}>{html.escape(p["short_title"])}</a>')
    nav_bar = "\n      ".join(nav_links)

    mirror_canonical = f"{GH_PAGES_BASE}/" if route_path == "/" else f"{GH_PAGES_BASE}{route_path}"

    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html.escape(title)}</title>
  <link rel="canonical" href="{html.escape(direct_url)}">
  <meta name="description" content="{html.escape(title)} — NASA Electronic Institutional Review Board (eIRB) at Johnson Space Center.">
  <style>
    :root {{
      --nasa-blue: #0b3d91;
      --nasa-dark: #061f4a;
      --nasa-red: #fc3d21;
      --bg-light: #f5f7fa;
      --card-bg: #ffffff;
      --text-main: #1b1b1b;
      --text-muted: #4b5563;
      --border-color: #d1d5db;
    }}
    * {{ box-sizing: border-box; }}
    body {{
      margin: 0;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
      line-height: 1.6;
      color: var(--text-main);
      background: var(--bg-light);
    }}
    header.site-header {{
      background: linear-gradient(135deg, var(--nasa-dark), var(--nasa-blue));
      color: #fff;
      padding: 1.25rem 1.5rem;
      border-bottom: 4px solid var(--nasa-red);
    }}
    header.site-header .header-inner {{
      max-width: 1200px;
      margin: 0 auto;
    }}
    header.site-header .eyebrow {{
      font-size: 0.8rem;
      text-transform: uppercase;
      letter-spacing: 0.08em;
      opacity: 0.9;
      margin-bottom: 0.25rem;
    }}
    header.site-header h1 {{
      margin: 0;
      font-size: 1.5rem;
      font-weight: 700;
    }}
    nav.site-nav {{
      background: #fff;
      border-bottom: 1px solid var(--border-color);
      padding: 0.5rem 1.5rem;
    }}
    nav.site-nav .nav-inner {{
      max-width: 1200px;
      margin: 0 auto;
      display: flex;
      flex-wrap: wrap;
      gap: 0.35rem 0.75rem;
    }}
    nav.site-nav a {{
      color: var(--nasa-blue);
      text-decoration: none;
      font-size: 0.88rem;
      font-weight: 500;
      padding: 0.25rem 0.5rem;
      border-radius: 4px;
    }}
    nav.site-nav a:hover, nav.site-nav a.active {{
      background: var(--nasa-blue);
      color: #fff;
    }}
    .container {{
      max-width: 1200px;
      margin: 1.5rem auto;
      padding: 0 1.5rem;
    }}
    .source-banner {{
      background: #eef2ff;
      border-left: 4px solid var(--nasa-blue);
      padding: 0.75rem 1rem;
      margin-bottom: 1.25rem;
      border-radius: 0 6px 6px 0;
      font-size: 0.9rem;
    }}
    .card {{
      background: var(--card-bg);
      border: 1px solid var(--border-color);
      border-radius: 8px;
      padding: 1.25rem 1.5rem;
      margin-bottom: 1.25rem;
      box-shadow: 0 1px 2px rgba(0,0,0,0.04);
    }}
    .table-wrap {{
      overflow-x: auto;
    }}
    table.eirb-table, .container table {{
      width: 100%;
      border-collapse: collapse;
      margin: 0.75rem 0;
    }}
    table.eirb-table th, table.eirb-table td {{
      border: 1px solid var(--border-color);
      padding: 0.55rem 0.75rem;
      text-align: left;
      vertical-align: top;
      font-size: 0.92rem;
    }}
    table.eirb-table th {{
      background: #f1f5f9;
      font-weight: 600;
    }}
    img {{
      max-width: 100%;
      height: auto;
    }}
    code {{
      background: #f1f5f9;
      padding: 0.1rem 0.35rem;
      border-radius: 3px;
      font-size: 0.88em;
    }}
    footer.site-footer {{
      max-width: 1200px;
      margin: 2rem auto;
      padding: 1rem 1.5rem;
      border-top: 1px solid var(--border-color);
      font-size: 0.85rem;
      color: var(--text-muted);
    }}
  </style>
</head>
<body>
  <header class="site-header">
    <div class="header-inner">
      <div class="eyebrow">National Aeronautics and Space Administration &bull; Office of Research Assurance (ORA) &bull; {html.escape(section_label)}</div>
      <h1>{html.escape(title)}</h1>
    </div>
  </header>
  <nav class="site-nav" aria-label="eIRB Portal Navigation">
    <div class="nav-inner">
      {nav_bar}
    </div>
  </nav>
  <div class="container">
    <div class="source-banner">
      <strong>Official NASA eIRB Portal Resource:</strong>
      <a href="{html.escape(direct_url)}" target="_blank" rel="noopener">{html.escape(direct_url)}</a>
      &bull; <strong>Mirror URL:</strong> <a href="{html.escape(mirror_canonical)}">{html.escape(mirror_canonical)}</a>
      &bull; <strong>Contact:</strong> <a href="mailto:NASA-IRB@nasa.gov">NASA-IRB@nasa.gov</a> | <a href="mailto:hq-hrmrb@mail.nasa.gov">hq-hrmrb@mail.nasa.gov</a>
      {extra_meta_html}
    </div>
    <main id="main-content" class="card">
{body_html}
    </main>
  </div>
  <footer class="site-footer">
    <p>NASA Electronic Institutional Review Board (eIRB) &bull; Johnson Space Center (JSC) / Office of the Chief Health and Medical Officer (OCHMO) &bull; Federalwide Assurance FWA00019876 &bull; 14 CFR Part 1230 &bull; NPD 7100.8 &bull; NPR 7100.1</p>
  </footer>
</body>
</html>
"""


def build_sitemap_xml(entries: list[dict]) -> str:
    lines = [
        '<?xml version="1.0" encoding="UTF-8"?>',
        '<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">',
    ]
    for e in entries:
        lines.append("  <url>")
        lines.append(f"    <loc>{html.escape(e['loc'])}</loc>")
        lines.append(f"    <lastmod>{html.escape(e['lastmod'])}</lastmod>")
        lines.append("    <changefreq>weekly</changefreq>")
        lines.append(f"    <priority>{html.escape(e['priority'])}</priority>")
        lines.append("  </url>")
    lines.append("</urlset>")
    return "\n".join(lines) + "\n"


def main():
    print("=" * 74)
    print("NASA Electronic Institutional Review Board (eirb.jsc.nasa.gov)")
    print("Harvester, Static HTML/Document Mirror & Sitemap Generator")
    print("=" * 74)

    existing_catalog = load_existing_catalog()
    existing_mirror_dates = load_existing_sitemap_dates(SITEMAP_FILE)
    existing_direct_dates = load_existing_sitemap_dates(DIRECT_SITEMAP_FILE)
    prev_hashes = existing_catalog.get("pageHashes", {})
    today = datetime.now(timezone.utc).strftime("%Y-%m-%d")

    opener = build_opener()

    print("[1/5] Fetching all 18 Huron Click Portal Container pages from eirb.jsc.nasa.gov...")
    raw_pages: dict[str, str] = {}
    discovered_oids: set[str] = set()

    for p in PORTAL_PAGES:
        oid = p["oid"]
        url = container_direct_url(oid)
        raw_bytes = fetch_url_bytes(opener, url)
        raw_html = raw_bytes.decode("utf-8", errors="ignore")
        raw_pages[oid] = raw_html
        for m_oid in re.findall(r"OID(?:\[|%5B)([0-9A-F]{32})(?:\]|%5D)", raw_html, re.I):
            discovered_oids.add(m_oid.upper())
        print(f"      Fetched {p['short_title']:26s} ({oid}) -> {len(raw_html):,} bytes")

    extra_oids = sorted(discovered_oids - set(OID_TO_PAGE.keys()))
    if extra_oids:
        print(f"      Discovered {len(extra_oids)} additional Container OIDs: {extra_oids}")

    print("[2/5] Querying Huron Click Codex JSON APIs for Library Documents & Related Links...")
    lib_oid = "3274673DA99811EEAD81FF18B24EF50A"
    rel_oid = "63B0D080C0B8994880C9DD20982017B3"
    library_sections, doc_by_id = harvest_library_documents(opener, raw_pages[lib_oid], lib_oid)
    related_links_sections = harvest_related_links(opener, raw_pages[rel_oid], rel_oid)
    print(
        f"      Harvested {len(doc_by_id)} Library documents across {len(library_sections)} sections "
        f"and {sum(len(s['links']) for s in related_links_sections)} Related Links across {len(related_links_sections)} sections."
    )

    print("[3/5] Mirroring portal diagram images and 54 downloadable IRB documents (PDF/DOCX)...")
    for img in PORTAL_IMAGES:
        dst_path = os.path.join(BASE_DIR, img["rel_path"])
        try:
            img_bytes = fetch_url_bytes(opener, img["remote_url"])
            write_bytes_if_changed(dst_path, img_bytes)
            print(f"      Mirrored image {img['rel_path']} ({len(img_bytes):,} bytes)")
        except Exception as e:
            print(f"      [WARN] Could not mirror image {img['remote_url']}: {e}")

    rendered_doc_pages = []
    for doc_id, doc in doc_by_id.items():
        if doc["isVideo"]:
            print(f"      Skipping binary mirror of large video {doc['name']} ({doc['fileSize']})")
            continue

        local_file_path = os.path.join(BASE_DIR, doc["relFilePath"])
        doc_bytes = b""
        # If file already exists and dateModified + version match catalog, reuse local bytes to save bandwidth;
        # otherwise download fresh from eirb.jsc.nasa.gov
        prev_doc_meta = existing_catalog.get("documents", {}).get(doc_id, {})
        if (
            os.path.exists(local_file_path)
            and prev_doc_meta.get("dateModified") == doc["dateModified"]
            and prev_doc_meta.get("version") == doc["version"]
        ):
            with open(local_file_path, "rb") as f:
                doc_bytes = f.read()
        else:
            doc_bytes = fetch_url_bytes(opener, doc["directUrl"], timeout=60)
            write_bytes_if_changed(local_file_path, doc_bytes)

        if doc["ext"] == ".docx":
            extracted_html, _ = extract_docx_html(doc_bytes)
        elif doc["ext"] == ".pdf":
            extracted_html, _ = extract_pdf_html(local_file_path)
        else:
            extracted_html = "<p><em>Binary resource file.</em></p>"

        desc_text = (
            html.escape(doc["description"]).replace("\r\n", "<br>").replace("\n", "<br>")
            if doc["description"]
            else "Official NASA eIRB Library document."
        )
        doc_body_html = f"""
      <section class="doc-metadata">
        <h2>{html.escape(doc['name'])}</h2>
        <p>{desc_text}</p>
        <div class="table-wrap">
          <table class="eirb-table">
            <tbody>
              <tr><th>Document ID</th><td><code>{html.escape(doc['id'])}</code></td><th>Library Section</th><td><a href="{GH_PAGES_BASE}/library">{html.escape(doc['section'])}</a></td></tr>
              <tr><th>Version</th><td>{html.escape(doc['version'])}</td><th>File Format / Size</th><td><code>{html.escape(doc['ext'].upper().lstrip('.'))}</code> ({html.escape(doc['fileSize'])})</td></tr>
              <tr><th>Date Created</th><td>{html.escape(doc['dateCreated'])}</td><th>Last Modified</th><td>{html.escape(doc['dateModified'])}</td></tr>
              <tr><th>Author / Owner</th><td>{html.escape(doc['author'])}</td><th>Download Original File</th><td><a href="{html.escape(doc['mirrorFileUrl'])}"><strong>Download Mirrored {html.escape(doc['ext'].upper().lstrip('.'))}</strong></a> | <a href="{html.escape(doc['directUrl'])}" target="_blank" rel="noopener">NASA eIRB Live File</a></td></tr>
            </tbody>
          </table>
        </div>
      </section>
      <hr style="border:0;border-top:1px solid #d1d5db;margin:1.25rem 0;">
      <section class="doc-fulltext">
        <h2>Full Document Content</h2>
        {extracted_html}
      </section>
"""
        doc_title = f"{doc['name']} ({doc['id']}) — {doc['section']} | NASA eIRB Library"
        doc_html_page = render_page_shell(
            title=doc_title,
            route_path=doc["docPageRoute"],
            direct_url=doc["directUrl"],
            section_label=f"eIRB Library — {doc['section']}",
            body_html=doc_body_html,
        )
        doc_hash = hashlib.sha256(doc_html_page.encode("utf-8")).hexdigest()[:16]

        for rel_out in [f"docs/{doc_id}/index.html", f"docs/{doc_id}.html"]:
            write_if_changed(os.path.join(BASE_DIR, rel_out), doc_html_page)

        lastmod_date = (
            doc["dateModified"].split(" ")[0]
            if re.match(r"^\d{4}-\d{2}-\d{2}", doc["dateModified"])
            else today
        )
        rendered_doc_pages.append(
            {
                "route": doc["docPageRoute"],
                "title": doc_title,
                "mirror_url": doc["docPageMirrorUrl"],
                "direct_url": doc["directUrl"],
                "priority": "0.85" if "Standard Operating Procedures" in doc["section"] else "0.80",
                "content_hash": doc_hash,
                "doc_lastmod": lastmod_date,
            }
        )

    print(f"      Pre-rendered {len(rendered_doc_pages)} full-text document HTML pages under /docs/<DOC_ID>")

    print("[4/5] Pre-rendering 18 static HTML portal pages...")
    rendered_portal_pages = []
    for p in PORTAL_PAGES:
        oid = p["oid"]
        cleaned_body = clean_portal_main_html(
            raw_pages[oid],
            p,
            doc_by_id,
            library_sections,
            related_links_sections,
        )
        direct_url = container_direct_url(oid)
        page_html = render_page_shell(
            title=p["title"],
            route_path=p["route"],
            direct_url=direct_url,
            section_label=p["section"],
            body_html=cleaned_body,
        )
        content_hash = hashlib.sha256(page_html.encode("utf-8")).hexdigest()[:16]

        if p["route"] == "/":
            rel_files = ["index.html"]
            mirror_url = f"{GH_PAGES_BASE}/"
        else:
            slug = p["slug"]
            rel_files = [f"{slug}/index.html", f"{slug}.html"]
            mirror_url = f"{GH_PAGES_BASE}/{slug}"

        for rel_path in rel_files:
            write_if_changed(os.path.join(BASE_DIR, rel_path), page_html)

        rendered_portal_pages.append(
            {
                "oid": oid,
                "route": p["route"],
                "title": p["title"],
                "mirror_url": mirror_url,
                "direct_url": direct_url,
                "priority": p["priority"],
                "content_hash": content_hash,
            }
        )

    # Also write /EIRB/sd/Rooms/DisplayPages/LayoutInitial/index.html & LayoutInitial.html
    # so any direct query-string navigation to LayoutInitial?Container=... on GitHub Pages redirects cleanly
    oid_map_js = json.dumps(
        {
            p["oid"]: (f"{GH_PAGES_BASE}/" if p["route"] == "/" else f"{GH_PAGES_BASE}{p['route']}")
            for p in PORTAL_PAGES
        }
    )
    layout_initial_html = render_page_shell(
        title="NASA eIRB Portal Directory — LayoutInitial Container Router",
        route_path="/",
        direct_url=f"{SITE_ORIGIN}/",
        section_label="Portal Directory",
        body_html=(
            "<h2>NASA eIRB Portal Pages</h2><ul>"
            + "".join(
                f'<li><a href="{GH_PAGES_BASE if p["route"] != "/" else GH_PAGES_BASE + "/"}{p["route"] if p["route"] != "/" else ""}">'
                f'{html.escape(p["title"])}</a> (<code>{p["oid"]}</code>)</li>'
                for p in PORTAL_PAGES
            )
            + "</ul>"
            + f"<script>(function(){{var m={oid_map_js};var q=window.location.search||'';var match=q.match(/OID(?:\\[|%5B)([0-9A-F]{{32}})(?:\\]|%5D)/i);if(match&&m[match[1].toUpperCase()]){{window.location.replace(m[match[1].toUpperCase()]);}}}})();</script>"
        ),
    )
    write_if_changed(
        os.path.join(BASE_DIR, "EIRB/sd/Rooms/DisplayPages/LayoutInitial/index.html"),
        layout_initial_html,
    )
    write_if_changed(
        os.path.join(BASE_DIR, "EIRB/sd/Rooms/DisplayPages/LayoutInitial.html"),
        layout_initial_html,
    )

    print("[5/5] Generating XML sitemaps and URL lists (preserving unchanged <lastmod>)...")
    mirror_entries = []
    direct_entries = []
    new_hashes = {}

    for p in rendered_portal_pages:
        route = p["route"]
        chash = p["content_hash"]
        new_hashes[route] = chash

        if prev_hashes.get(route) == chash and p["mirror_url"] in existing_mirror_dates:
            m_lastmod = existing_mirror_dates[p["mirror_url"]]
        else:
            m_lastmod = today

        if prev_hashes.get(route) == chash and p["direct_url"] in existing_direct_dates:
            d_lastmod = existing_direct_dates[p["direct_url"]]
        else:
            d_lastmod = today

        mirror_entries.append(
            {"loc": p["mirror_url"], "lastmod": m_lastmod, "priority": p["priority"]}
        )
        direct_entries.append(
            {"loc": p["direct_url"], "lastmod": d_lastmod, "priority": p["priority"]}
        )

    for dp in rendered_doc_pages:
        route = dp["route"]
        chash = dp["content_hash"]
        new_hashes[route] = chash

        if prev_hashes.get(route) == chash and dp["mirror_url"] in existing_mirror_dates:
            m_lastmod = existing_mirror_dates[dp["mirror_url"]]
        else:
            m_lastmod = dp.get("doc_lastmod") or today

        if prev_hashes.get(route) == chash and dp["direct_url"] in existing_direct_dates:
            d_lastmod = existing_direct_dates[dp["direct_url"]]
        else:
            d_lastmod = dp.get("doc_lastmod") or today

        mirror_entries.append(
            {"loc": dp["mirror_url"], "lastmod": m_lastmod, "priority": dp["priority"]}
        )
        direct_entries.append(
            {"loc": dp["direct_url"], "lastmod": d_lastmod, "priority": dp["priority"]}
        )

    for ext in EXTERNAL_REFERENCE_URLS:
        u = ext["url"]
        m_lastmod = existing_mirror_dates.get(u, today)
        d_lastmod = existing_direct_dates.get(u, today)
        mirror_entries.append(
            {"loc": u, "lastmod": m_lastmod, "priority": ext["priority"]}
        )
        direct_entries.append(
            {"loc": u, "lastmod": d_lastmod, "priority": ext["priority"]}
        )

    mirror_xml = build_sitemap_xml(mirror_entries)
    direct_xml = build_sitemap_xml(direct_entries)

    write_if_changed(SITEMAP_FILE, mirror_xml)
    write_if_changed(SITEMAP_ROOT_FILE, mirror_xml)
    write_if_changed(DIRECT_SITEMAP_FILE, direct_xml)

    mirror_urls_txt = "\n".join(e["loc"] for e in mirror_entries) + "\n"
    direct_urls_txt = "\n".join(e["loc"] for e in direct_entries) + "\n"
    write_if_changed(URLS_FILE, mirror_urls_txt)
    write_if_changed(DIRECT_URLS_FILE, direct_urls_txt)

    catalog_payload = {
        "siteOrigin": SITE_ORIGIN,
        "ghPagesBase": GH_PAGES_BASE,
        "pageHashes": new_hashes,
        "portalPages": rendered_portal_pages,
        "documentPages": rendered_doc_pages,
        "documents": doc_by_id,
        "relatedLinks": related_links_sections,
        "externalReferences": EXTERNAL_REFERENCE_URLS,
    }
    write_if_changed(
        CATALOG_FILE, json.dumps(catalog_payload, indent=2, sort_keys=True) + "\n"
    )

    print(
        f"Done! Total URLs in sitemap: {len(mirror_entries)} "
        f"({len(rendered_portal_pages)} portal pages + {len(rendered_doc_pages)} full-text IRB documents + {len(EXTERNAL_REFERENCE_URLS)} NASA policy URLs)."
    )


if __name__ == "__main__":
    main()
