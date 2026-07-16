"""
Baker Tilly AI Audit Reviewer — Stage 4
Adds: post-login chooser page (Financial Statements review vs Working-paper
review, each with its own tailored AI review logic), an optional instruction
box at upload (tell the AI what to focus on), and a dark navy header band so
the white firm logo is visible. Retains from earlier stages: login with roles,
multi-file drag-and-drop, knowledge-library citations, Excel/PDF downloads,
gunicorn 600s timeout, tolerant JSON parser, lightweight Excel reader with
shared-text cap, per-file memory release, friendly 413/500 pages.

USER ACCOUNTS (managed by the administrator, never stored in this public code):
Set an environment variable on Render called USERS in this format:
    username:password:role;username2:password2:role2
Roles:  full    = can review and download reports (Partner / Manager)
        limited = can review only (no downloads)
Example:
    partner1:Str0ngPass!:full;manager1:An0therPass!:full;staff1:StaffPass1:limited
If USERS is not set, a single default login exists:
    admin / bakertilly2025  (full)  — CHANGE THIS by setting USERS.
Also set SECRET_KEY to any long random text (keeps logins secure).
"""

import os
import io
import json
import uuid
import tempfile
from functools import wraps
from flask import (Flask, request, render_template_string, session,
                   redirect, url_for, send_file)
from openai import OpenAI

from openpyxl import Workbook
from docx import Document as DocxDocument
from pypdf import PdfReader

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib import colors

app = Flask(__name__)
UPLOAD_LIMIT_MB = 50
app.config["MAX_CONTENT_LENGTH"] = UPLOAD_LIMIT_MB * 1024 * 1024
app.secret_key = os.environ.get("SECRET_KEY", "change-me-set-SECRET-KEY-env-var")


@app.errorhandler(413)
def too_large(e):
    """Friendly message instead of a crash page when the upload is too big."""
    msg = (f"Your upload is too large. The limit is {UPLOAD_LIMIT_MB} MB per batch "
           f"on this hosting. Please upload fewer or smaller files, or split the "
           f"batch. (Unlimited sizes become possible once the tool moves to the "
           f"firm's own server.)")
    return render_template_string(MAIN_PAGE, user=session.get("user", ""),
                                  role=session.get("role", "limited"), error=msg,
                                  batch=None, batch_id=None,
                                  maxfiles=MAX_FILES_PER_BATCH,
                                  mode=session.get("mode", "wp"),
                                  disclaimer=DISCLAIMER), 413


@app.errorhandler(500)
def server_error(e):
    """Friendly message instead of the bare 'Internal Server Error' page."""
    msg = ("Something went wrong while processing your request. Please try again "
           "with fewer or smaller files. If it keeps happening, note what you "
           "uploaded and report it.")
    try:
        return render_template_string(MAIN_PAGE, user=session.get("user", ""),
                                      role=session.get("role", "limited"), error=msg,
                                      batch=None, batch_id=None,
                                      maxfiles=MAX_FILES_PER_BATCH,
                                      mode=session.get("mode", "wp"),
                                      disclaimer=DISCLAIMER), 500
    except Exception:
        return msg, 500

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")

MAX_FILES_PER_BATCH = 8
MAX_EXTRACT_CHARS = 45000
RESULTS_DIR = os.path.join(tempfile.gettempdir(), "audit_results")
os.makedirs(RESULTS_DIR, exist_ok=True)


def load_users():
    """Users come from the USERS environment variable (set on Render).
    Format: username:password:role;username2:password2:role2"""
    raw = os.environ.get("USERS", "").strip()
    users = {}
    if raw:
        for entry in raw.split(";"):
            parts = entry.strip().split(":")
            if len(parts) == 3:
                name, pw, role = parts[0].strip(), parts[1], parts[2].strip().lower()
                if name and pw and role in ("full", "limited"):
                    users[name] = {"password": pw, "role": role}
    if not users:
        users["admin"] = {"password": "bakertilly2025", "role": "full"}
    return users


def login_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return wrapper


def full_access_required(f):
    @wraps(f)
    def wrapper(*args, **kwargs):
        if not session.get("user"):
            return redirect(url_for("login"))
        if session.get("role") != "full":
            return "Downloads are available to full-access users only.", 403
        return f(*args, **kwargs)
    return wrapper


def load_knowledge_base():
    kb_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge")
    sections = []
    if os.path.isdir(kb_dir):
        for fname in sorted(os.listdir(kb_dir)):
            if fname.endswith(".txt"):
                try:
                    with open(os.path.join(kb_dir, fname), "r", encoding="utf-8") as f:
                        sections.append(f.read().strip())
                except Exception:
                    pass
    return "\n\n==========\n\n".join(sections)


KNOWLEDGE_BASE = load_knowledge_base()


def _extract_xlsx_lightweight(file_bytes):
    """Read sheet text straight from the xlsx internals (an xlsx is a zip of
    XML files). This avoids openpyxl building the full workbook object —
    external links, styles and structures are skipped entirely, keeping
    memory tiny even for complex, heavily-linked audit workbooks."""
    import zipfile
    import re as _re
    from xml.etree.ElementTree import iterparse

    NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
    text_parts = []
    total = 0

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        names = z.namelist()

        # shared strings (xlsx stores text centrally)
        shared = []
        if "xl/sharedStrings.xml" in names:
            SHARED_CAP = 2_000_000  # cap total shared-text characters (memory guard)
            shared_total = 0
            with z.open("xl/sharedStrings.xml") as f:
                for ev, el in iterparse(f, events=("end",)):
                    if el.tag == NS + "si":
                        if shared_total < SHARED_CAP:
                            texts = [t.text or "" for t in el.iter(NS + "t")]
                            s = "".join(texts)
                            shared.append(s)
                            shared_total += len(s)
                        else:
                            shared.append("")  # beyond cap: placeholder
                        el.clear()  # clear only completed string items

        # sheet name map (falls back to file order if unavailable)
        sheet_titles = {}
        try:
            if "xl/workbook.xml" in names:
                with z.open("xl/workbook.xml") as f:
                    idx = 0
                    for ev, el in iterparse(f, events=("end",)):
                        if el.tag == NS + "sheet":
                            idx += 1
                            sheet_titles[idx] = el.get("name", f"Sheet{idx}")
                        el.clear()
        except Exception:
            pass

        sheet_files = sorted(
            n for n in names
            if _re.match(r"xl/worksheets/sheet\d+\.xml$", n)
        )
        for snum, sname in enumerate(sheet_files, start=1):
            if total >= MAX_EXTRACT_CHARS:
                text_parts.append("\n[... file is large; remaining sheets not "
                                  "included in this review pass ...]")
                break
            title = sheet_titles.get(snum, f"Sheet{snum}")
            header = "\n===== SHEET: " + title + " ====="
            text_parts.append(header)
            total += len(header)

            row_cells = []
            with z.open(sname) as f:
                for ev, el in iterparse(f, events=("end",)):
                    tag = el.tag
                    if tag == NS + "c":  # a cell
                        ctype = el.get("t")
                        v = el.find(NS + "v")
                        val = None
                        if ctype == "s" and v is not None:
                            try:
                                val = shared[int(v.text)]
                            except Exception:
                                val = v.text
                        elif ctype == "inlineStr":
                            is_el = el.find(NS + "is")
                            if is_el is not None:
                                t = is_el.find(NS + "t")
                                val = t.text if t is not None else None
                        elif ctype == "e" and v is not None:
                            val = v.text  # keep #REF!, #VALUE! etc — we WANT these
                        elif v is not None:
                            val = v.text
                        if val is not None and str(val).strip() != "":
                            row_cells.append(str(val))
                    elif tag == NS + "row":
                        if row_cells:
                            line = " | ".join(row_cells)
                            text_parts.append(line)
                            total += len(line)
                        row_cells = []
                        el.clear()  # safe to clear once the whole row is done
                        if total >= MAX_EXTRACT_CHARS:
                            break

    return "\n".join(text_parts)


def extract_text_from_file(filename, file_bytes):
    name = filename.lower()

    if name.endswith((".xlsx", ".xlsm")):
        return _extract_xlsx_lightweight(file_bytes)

    elif name.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(file_bytes))
        parts, total = [], 0
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text)
                total += len(p.text)
                if total >= MAX_EXTRACT_CHARS:
                    parts.append("[... document is large; remainder not included ...]")
                    break
        return "\n".join(parts)

    elif name.endswith(".pdf"):
        reader = PdfReader(io.BytesIO(file_bytes))
        parts, total = [], 0
        for page in reader.pages:
            t = page.extract_text() or ""
            parts.append(t)
            total += len(t)
            if total >= MAX_EXTRACT_CHARS:
                parts.append("[... document is large; remaining pages not included ...]")
                break
        return "\n".join(parts)

    elif name.endswith((".csv", ".txt")):
        return file_bytes.decode("utf-8", errors="ignore")[:MAX_EXTRACT_CHARS]

    return None


REVIEWER_INSTRUCTIONS = """You are an experienced audit reviewer at an accounting firm, reviewing audit working papers to the standard expected in an ICAP Quality Control Review or an Audit Oversight Board inspection.

You will be given the text extracted from an audit working paper (often a revenue or other head, sometimes with supporting figures and calculations).

Review it carefully and identify EVERY discrepancy, error, omission, weakness, or matter needing attention. Look specifically for:
- Figures or totals that do not add up, or that do not agree between different parts of the document
- Broken spreadsheet values such as #REF!, #DIV/0!, #VALUE! - these are hard errors
- Content that appears to belong to a DIFFERENT client or engagement (wrong client name, another file reference left in from a reused template) - copy-paste contamination
- Conclusions that are pre-filled or boilerplate ("satisfactory", "fairly stated") without evidence that actual work supports them
- Missing sign-offs, missing dates, or dates out of logical sequence
- Vague or unquantified work (e.g. a "sample" with no number of items tested)
- Calculations that look wrong or unsupported
- Anything a working paper needs but is missing (evidence, cross-references, explanations)
- Non-compliance with the applicable accounting or auditing standards

For EACH issue you find, give:
1. A short, clear title of the issue
2. A plain-English explanation (simple language a junior staff member can understand - avoid unnecessary jargon)
3. The applicable standard or rule reference
4. A severity: High, Medium, or Low (or "Factual" for arithmetic/broken-value errors that are simply right or wrong)
5. A suggested fix - what the team should do to resolve it

IMPORTANT RULES:
- You are given the FIRM'S STANDARDS LIBRARY below. Base every standard reference on that library. When your finding is supported by the library, cite it (e.g. "IFRS 15 - control transfer (per firm standards library)").
- If an issue is real but the library does not cover it, still raise it, but mark the reference as "outside loaded library - reference to be confirmed".
- Never invent a standard, paragraph number, or fact. If unsure, say so.
- Write everything in easy-to-understand English.
- Base your findings on what is actually in the document provided, not assumptions.

Return your answer as a JSON object with this exact structure:
{
  "findings": [
    {
      "title": "...",
      "explanation": "...",
      "reference": "...",
      "severity": "High | Medium | Low | Factual",
      "fix": "..."
    }
  ],
  "summary": "A one or two sentence overall summary of the file's condition.",
  "conclusion": "A 2-4 sentence head-wise conclusion in plain English: the overall condition of this working paper, whether its documented conclusions can currently be relied on, and what must be fixed first."
}
Return ONLY the JSON, no other text."""


FS_REVIEWER_INSTRUCTIONS = """You are an experienced audit reviewer at an accounting firm, reviewing a set of FINANCIAL STATEMENTS (or extracts from them) to the standard expected in an ICAP Quality Control Review or an Audit Oversight Board inspection.

You will be given text extracted from draft or final financial statements (statement of financial position, profit or loss, changes in equity, cash flows, and/or the notes).

Review carefully and identify EVERY discrepancy, error, omission, weakness, or matter needing attention. Look specifically for:
- Figures that do not agree between the face of the statements and the supporting notes (tie-out failures)
- Totals or subtotals that do not add up; casting errors
- Broken spreadsheet values such as #REF!, #DIV/0!, #VALUE! - these are hard errors
- Missing or incomplete disclosures required by the applicable standards (e.g. related party disclosures per IAS 24, revenue disaggregation per IFRS 15)
- IAS 1 presentation problems: material classes not presented separately, missing comparative figures, missing cross-references between the face and the notes
- Accounting policies that are missing, boilerplate, or inconsistent with the figures presented
- Inconsistencies between different statements (e.g. profit per P&L not agreeing with the movement in retained earnings)
- Companies Act 2017 concerns: anything preventing a true and fair view
- Content that appears to belong to a DIFFERENT company (wrong name, copy-paste contamination from a template)

For EACH issue you find, give:
1. A short, clear title of the issue
2. A plain-English explanation (simple language a junior staff member can understand - avoid unnecessary jargon)
3. The applicable standard or rule reference
4. A severity: High, Medium, or Low (or "Factual" for arithmetic/broken-value errors that are simply right or wrong)
5. A suggested fix - what the team should do to resolve it

IMPORTANT RULES:
- You are given the FIRM'S STANDARDS LIBRARY below. Base every standard reference on that library. When your finding is supported by the library, cite it (e.g. "IAS 1 para 29 (per firm standards library)").
- If an issue is real but the library does not cover it, still raise it, but mark the reference as "outside loaded library - reference to be confirmed".
- Never invent a standard, paragraph number, or fact. If unsure, say so.
- Write everything in easy-to-understand English.
- Base your findings on what is actually in the document provided, not assumptions.

Return your answer as a JSON object with this exact structure:
{
  "findings": [
    {
      "title": "...",
      "explanation": "...",
      "reference": "...",
      "severity": "High | Medium | Low | Factual",
      "fix": "..."
    }
  ],
  "summary": "A one or two sentence overall summary of the statements' condition.",
  "conclusion": "A 2-4 sentence conclusion in plain English: the overall condition of these financial statements, whether they currently appear ready for sign-off, and what must be fixed first."
}
Return ONLY the JSON, no other text."""


def review_with_ai(document_text, mode="wp", user_instructions=""):
    trimmed = document_text[:MAX_EXTRACT_CHARS]
    instructions = FS_REVIEWER_INSTRUCTIONS if mode == "fs" else REVIEWER_INSTRUCTIONS
    doc_label = ("financial statements" if mode == "fs" else "working paper")
    messages = [
        {"role": "system", "content": instructions},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY (check against these texts):\n\n" + KNOWLEDGE_BASE},
    ]
    if user_instructions.strip():
        messages.append({"role": "system", "content":
            "SPECIFIC INSTRUCTIONS FROM THE REVIEWER FOR THIS BATCH (follow these, "
            "give the requested areas extra attention, and answer any questions asked "
            "within your findings or summary — but still report any other significant "
            "issues you notice):\n\n" + user_instructions.strip()[:2000]})
    messages.append({"role": "user", "content":
        "Here is the " + doc_label + " to review:\n\n" + trimmed})
    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=messages,
        max_tokens=6000,
        temperature=0.2,
    )
    raw = response.choices[0].message.content.strip()
    return parse_ai_json(raw)




BATCH_INSTRUCTIONS = """You are an experienced audit reviewer. You are given the review results for a BATCH of related audit files (working papers and possibly their supporting evidence such as confirmations, invoices, schedules).

Produce:
1. "overall_conclusion": a plain-English batch conclusion (3-5 sentences): the overall condition across the files, the weakest areas, and what the team should fix first.
2. "common_themes": a list of short strings - recurring problems appearing across multiple files (e.g. "Sign-offs missing in 4 of 6 files").
3. "cross_file_observations": a list of short strings - inconsistencies or corroboration issues BETWEEN the files (e.g. a figure in one file not agreeing with the supporting document in another, or a working paper claiming evidence that the attached evidence does not show). If none can be determined, return an empty list.

Base everything only on the material provided. Never invent facts or references. Plain English.
Return ONLY a JSON object: {"overall_conclusion": "...", "common_themes": [...], "cross_file_observations": [...]}"""


def batch_conclusion_with_ai(batch):
    """One extra AI pass across the whole batch: overall conclusion, themes,
    and cross-file (evidence corroboration) observations."""
    parts = []
    for item in batch["files"]:
        parts.append("FILE: " + item["filename"])
        if item.get("error"):
            parts.append("  (could not be reviewed: " + item["error"][:200] + ")")
            continue
        res = item.get("result", {})
        if res.get("summary"):
            parts.append("  Summary: " + res["summary"])
        for f in res.get("findings", [])[:12]:
            parts.append("  - [" + f.get("severity", "") + "] " + f.get("title", "")
                         + ": " + f.get("explanation", "")[:200])
        excerpt = (item.get("excerpt") or "")[:3000]
        if excerpt:
            parts.append("  EXCERPT OF FILE CONTENT:\n" + excerpt)
    material = "\n".join(parts)[:30000]

    try:
        response = client.chat.completions.create(
            model="deepseek-chat",
            messages=[
                {"role": "system", "content": BATCH_INSTRUCTIONS},
                {"role": "user", "content": material},
            ],
            max_tokens=1500,
            temperature=0.2,
        )
        raw = response.choices[0].message.content.strip()
        parsed, err = parse_ai_json(raw)
        if parsed and "overall_conclusion" in str(parsed):
            return parsed
    except Exception:
        pass
    return None


def parse_ai_json(raw):
    """Read the AI's JSON response, tolerating common formatting quirks."""
    import re as _re
    text = raw.strip()

    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
        if text.lstrip().lower().startswith("json"):
            text = text.lstrip()[4:]
        text = text.strip()

    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    try:
        return json.loads(text), None
    except Exception:
        pass

    cleaned = (text
               .replace("\u201c", '"').replace("\u201d", '"')
               .replace("\u2018", "'").replace("\u2019", "'"))
    cleaned = _re.sub(r",\s*([}\]])", r"\1", cleaned)
    try:
        return json.loads(cleaned), None
    except Exception:
        pass

    try:
        objs = []
        depth = 0
        buf = ""
        in_list = False
        i = cleaned.find('"findings"')
        if i != -1:
            rest = cleaned[i:]
            for ch in rest:
                if not in_list:
                    if ch == "[":
                        in_list = True
                    continue
                if ch == "{":
                    depth += 1
                if depth > 0:
                    buf += ch
                if ch == "}":
                    depth -= 1
                    if depth == 0 and buf.strip():
                        try:
                            objs.append(json.loads(buf))
                        except Exception:
                            pass
                        buf = ""
                if ch == "]" and depth == 0:
                    break
        if objs:
            return {"findings": objs,
                    "summary": "Note: the AI's response was cut off, so the "
                               "findings below may be incomplete."}, None
    except Exception:
        pass

    return None, ("The AI's response could not be read as structured findings. "
                  "Raw response:\n\n" + raw)


def save_results(batch):
    rid = uuid.uuid4().hex[:12]
    with open(os.path.join(RESULTS_DIR, rid + ".json"), "w", encoding="utf-8") as f:
        json.dump(batch, f)
    return rid


def load_results(rid):
    safe = "".join(c for c in rid if c.isalnum())
    path = os.path.join(RESULTS_DIR, safe + ".json")
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    return None


DISCLAIMER = ("This review has been prepared by an AI-assisted tool to support the audit "
              "review process by identifying possible discrepancies, errors, omissions, and "
              "matters requiring attention. It does not replace the judgement of the engagement "
              "team. All findings are observations for consideration, not conclusions. Every "
              "point should be reviewed, verified, and decided upon by a qualified member of "
              "the audit team. Final responsibility for the audit - including all professional "
              "judgements, the sufficiency of audit evidence, and the audit opinion - rests "
              "entirely with the Engagement Partner and the audit team, not with this tool. "
              "The AI does not sign off, approve, or conclude on any matter.")


def build_excel(batch):
    from openpyxl.styles import Font
    wb = Workbook()
    ws = wb.active
    ws.title = "Review Points"
    ws.append(["File", "No.", "Title", "Severity", "Explanation",
               "Reference", "Suggested fix"])
    for c in ws[1]:
        c.font = Font(bold=True)
    if batch.get("overall"):
        ws.append(["BATCH", "-", "OVERALL CONCLUSION", "-",
                   batch["overall"].get("overall_conclusion", ""), "", ""])
        for t in batch["overall"].get("common_themes", []):
            ws.append(["BATCH", "-", "Common theme", "-", t, "", ""])
        for t in batch["overall"].get("cross_file_observations", []):
            ws.append(["BATCH", "-", "Cross-file observation", "-", t, "", ""])
        ws.append([])
    for item in batch["files"]:
        fname = item["filename"]
        if item.get("error"):
            ws.append([fname, "-", "REVIEW ERROR", "-", item["error"], "-", "-"])
            continue
        if item.get("result", {}).get("conclusion"):
            ws.append([fname, "-", "HEAD-WISE CONCLUSION", "-",
                       item["result"]["conclusion"], "", ""])
        for i, f in enumerate(item["result"].get("findings", []), start=1):
            ws.append([fname, i, f.get("title", ""), f.get("severity", ""),
                       f.get("explanation", ""), f.get("reference", ""),
                       f.get("fix", "")])
    ws.append([])
    ws.append(["Professional judgement statement:"])
    ws.append([DISCLAIMER])
    widths = [28, 5, 34, 10, 60, 40, 50]
    for idx, w in enumerate(widths, start=1):
        ws.column_dimensions[chr(64 + idx)].width = w
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


def build_pdf(batch):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4,
                            leftMargin=1.6 * cm, rightMargin=1.6 * cm,
                            topMargin=1.6 * cm, bottomMargin=1.6 * cm)
    styles = getSampleStyleSheet()
    h1 = ParagraphStyle("h1x", parent=styles["Heading1"], fontSize=15)
    h2 = ParagraphStyle("h2x", parent=styles["Heading2"], fontSize=12,
                        textColor=colors.HexColor("#00A09B"))
    body = ParagraphStyle("bodyx", parent=styles["BodyText"], fontSize=9.5, leading=13)
    small = ParagraphStyle("smallx", parent=styles["BodyText"], fontSize=8,
                           leading=11, textColor=colors.HexColor("#5A4A28"))

    sev_color = {"High": "#B23A2E", "Medium": "#B0791C",
                 "Low": "#5B7083", "Factual": "#002B49"}

    story = [Paragraph("Baker Tilly - AI Audit Reviewer: Review Points", h1),
             Spacer(1, 8)]
    if batch.get("overall"):
        story.append(Paragraph("Overall batch conclusion", h2))
        story.append(Paragraph(batch["overall"].get("overall_conclusion", ""), body))
        for t in batch["overall"].get("common_themes", []):
            story.append(Paragraph("- " + t, body))
        for t in batch["overall"].get("cross_file_observations", []):
            story.append(Paragraph("- (cross-file) " + t, body))
        story.append(Spacer(1, 10))
    for item in batch["files"]:
        story.append(Paragraph("File: " + item["filename"], h2))
        if item.get("error"):
            story.append(Paragraph("Review error: " + item["error"], body))
            story.append(Spacer(1, 8))
            continue
        result = item["result"]
        if result.get("summary"):
            story.append(Paragraph("<b>Overall:</b> " + result["summary"], body))
            story.append(Spacer(1, 6))
        if result.get("conclusion"):
            story.append(Paragraph("<b>Head-wise conclusion:</b> " + result["conclusion"], body))
            story.append(Spacer(1, 6))
        for i, f in enumerate(result.get("findings", []), start=1):
            colr = sev_color.get(f.get("severity", ""), "#002B49")
            story.append(Paragraph(
                "<b>" + str(i) + ". " + f.get("title", "") + "</b> "
                "<font color='" + colr + "'>[" + f.get("severity", "") + "]</font>", body))
            story.append(Paragraph(f.get("explanation", ""), body))
            if f.get("reference"):
                story.append(Paragraph("<i>Reference: " + f["reference"] + "</i>", body))
            story.append(Paragraph("<b>Suggested fix:</b> " + f.get("fix", ""), body))
            story.append(Spacer(1, 7))
        story.append(Spacer(1, 10))
    story.append(Spacer(1, 8))
    story.append(Paragraph("<b>Professional judgement statement</b>", body))
    story.append(Paragraph(DISCLAIMER, small))
    doc.build(story)
    buf.seek(0)
    return buf


CHOOSE_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Baker Tilly - AI Audit Reviewer : Choose review type</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#0c1b34;margin:0;
      min-height:100vh;display:grid;place-items:center;color:#fff;overflow-x:hidden;}
 .stage{position:relative;width:100%;max-width:720px;padding:40px 20px;text-align:center;}
 .dot{position:absolute;border-radius:50%;pointer-events:none;}
 .d1{top:10%;left:6%;width:6px;height:6px;background:#2dd4bf;opacity:.35;animation:drift1 9s ease-in-out infinite;}
 .d2{top:72%;left:14%;width:9px;height:9px;background:#5eead4;opacity:.25;animation:drift2 11s ease-in-out infinite;}
 .d3{top:26%;left:86%;width:7px;height:7px;background:#2dd4bf;opacity:.3;animation:drift1 13s ease-in-out infinite;}
 .d4{top:84%;left:78%;width:5px;height:5px;background:#7c6cf0;opacity:.35;animation:drift2 8s ease-in-out infinite;}
 .d5{top:52%;left:47%;width:4px;height:4px;background:#5eead4;opacity:.2;animation:drift1 10s ease-in-out infinite;}
 .brand{display:inline-flex;align-items:center;gap:12px;margin-bottom:4px;
        animation:float 4.5s ease-in-out infinite;}
 .brand img{height:46px;}
 .logofb{width:46px;height:46px;border-radius:50%;background:radial-gradient(circle at 32% 30%,#2FD6D0,#00A09B);}
 h1{font-size:23px;margin:10px 0 2px;font-weight:600;animation:fadeUp .8s ease both;}
 .sub{color:#9fb3cc;font-size:14px;margin-bottom:30px;animation:fadeUp .8s .05s ease both;}
 .who{position:absolute;top:14px;right:18px;font-size:12px;color:#9fb3cc;}
 .who b{color:#fff;} .who a{color:#5eead4;margin-left:8px;text-decoration:none;}
 .choices{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;
          max-width:540px;margin:0 auto;}
 .choice{display:block;text-decoration:none;background:rgba(255,255,255,.05);
         border:1px solid rgba(94,234,212,.18);border-radius:12px;padding:26px 20px;
         transition:transform .25s ease,border-color .25s ease,background .25s ease;}
 .choice:hover{transform:translateY(-6px);border-color:#2dd4bf;background:rgba(45,212,191,.12);}
 .c1{animation:fadeUp .8s .15s ease both;} .c2{animation:fadeUp .8s .3s ease both;}
 .cico{font-size:30px;margin-bottom:10px;}
 .ctitle{color:#fff;font-size:16px;font-weight:600;margin-bottom:6px;}
 .cdesc{color:#9fb3cc;font-size:12.5px;line-height:1.6;}
 .foot{color:#6c8099;font-size:11px;max-width:440px;margin:30px auto 0;line-height:1.55;
       animation:fadeUp .8s .45s ease both;}
 @keyframes fadeUp{from{opacity:0;transform:translateY(14px);}to{opacity:1;transform:translateY(0);}}
 @keyframes float{0%,100%{transform:translateY(0);}50%{transform:translateY(-8px);}}
 @keyframes drift1{0%,100%{transform:translate(0,0);}50%{transform:translate(10px,-16px);}}
 @keyframes drift2{0%,100%{transform:translate(0,0);}50%{transform:translate(-12px,14px);}}
</style></head><body>
<div class="stage">
 <span class="dot d1"></span><span class="dot d2"></span><span class="dot d3"></span>
 <span class="dot d4"></span><span class="dot d5"></span>
 <div class="who">Signed in as <b>{{ user }}</b>
  <a href="{{ url_for('logout') }}">Log out</a></div>
 <div class="brand">
  <img src="https://www.bakertilly.pk/assets/images/logo.svg" alt="Baker Tilly"
       onerror="this.outerHTML=&quot;<div class=logofb></div>&quot;">
 </div>
 <h1>AI Audit Reviewer</h1>
 <div class="sub">Choose a review type to begin</div>
 <div class="choices">
  <a class="choice c1" href="{{ url_for('select_mode', mode='fs') }}">
   <div class="cico">&#128202;</div>
   <div class="ctitle">Financial Statements review</div>
   <div class="cdesc">Disclosures, IAS 1 presentation, note tie-outs, true and fair view</div>
  </a>
  <a class="choice c2" href="{{ url_for('select_mode', mode='wp') }}">
   <div class="cico">&#128203;</div>
   <div class="ctitle">Working-paper review</div>
   <div class="cdesc">Evidence, sign-offs, ISA 230 documentation, ISA 500 sufficiency</div>
  </a>
 </div>
 <div class="foot">Every AI output is a draft — final professional judgement rests with the audit team.</div>
</div></body></html>
"""

LOGIN_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Baker Tilly - AI Audit Reviewer : Sign in</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#ECEEF0;margin:0;
      min-height:100vh;display:grid;place-items:center;color:#002B49;}
 .card{background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:36px 32px;
       width:100%;max-width:380px;box-shadow:0 4px 24px rgba(20,35,59,.08);text-align:center;}
 .logo{height:44px;margin:0 auto 14px;display:flex;align-items:center;justify-content:center;}
 .logo img{height:44px;}
 .logofb{width:46px;height:46px;border-radius:50%;background:radial-gradient(circle at 32% 30%,#2FD6D0,#00A09B);}
 h1{font-size:20px;margin:0 0 4px;} .sub{font-size:13px;color:#5B7083;margin-bottom:24px;}
 label{display:block;text-align:left;font-size:12px;font-weight:600;color:#3A4A64;margin:10px 0 5px;}
 input{width:100%;box-sizing:border-box;padding:11px 13px;border:1px solid #B7BFC6;
       border-radius:6px;font-size:14px;}
 button{width:100%;padding:12px;background:#00A09B;color:#fff;border:none;border-radius:6px;
        font-size:14px;font-weight:600;cursor:pointer;margin-top:16px;}
 .err{color:#B23A2E;font-size:12.5px;min-height:16px;text-align:left;margin-top:8px;}
</style></head><body>
<div class="card">
 <div class="logo"><img src="https://www.bakertilly.pk/assets/images/logo.svg" alt="Baker Tilly" onerror="this.outerHTML=&quot;<div class=logofb></div>&quot;"></div>
 <h1>AI Audit Reviewer</h1>
 <div class="sub">Baker Tilly - Authorised users only</div>
 <form method="POST">
  <label>Username</label><input name="username" autocomplete="username" required>
  <label>Password</label><input type="password" name="password" autocomplete="current-password" required>
  <div class="err">{{ error or "" }}</div>
  <button type="submit">Sign in</button>
 </form>
</div></body></html>
"""

MAIN_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Baker Tilly - AI Audit Reviewer</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#ECEEF0;margin:0;
      padding:0 0 28px;color:#002B49;}
 .band{background:#0c1b34;padding:16px 28px;margin-bottom:20px;}
 .wrap{max-width:920px;margin:0 auto;padding:0 28px;}
 .top{max-width:920px;margin:0 auto;display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;}
 .brand{display:flex;align-items:center;gap:12px;}
 .logo{height:36px;display:flex;align-items:center;}
 .logo img{height:36px;}
 .logofb{width:38px;height:38px;border-radius:50%;background:radial-gradient(circle at 32% 30%,#2FD6D0,#00A09B);}
 h1{font-size:20px;margin:0;color:#fff;} .sub{color:#5B7083;font-size:13px;margin-bottom:18px;}
 .sub a{color:#00A09B;text-decoration:none;font-weight:600;}
 .who{font-size:12.5px;color:#9fb3cc;} .who b{color:#fff;}
 .who a{color:#5eead4;margin-left:10px;}
 .instr{width:100%;box-sizing:border-box;margin-top:14px;padding:11px 13px;
        border:1px solid #B7BFC6;border-radius:8px;font-size:13px;font-family:inherit;
        min-height:64px;resize:vertical;color:#002B49;}
 .instr-label{font-size:12.5px;font-weight:600;color:#3A4A64;margin:16px 0 5px;text-align:left;}
 .instr-hint{font-size:11.5px;color:#5B7083;margin-top:4px;text-align:left;}
 .card{background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:24px;
       box-shadow:0 2px 12px rgba(20,35,59,.06);margin-bottom:18px;}
 .notice{background:#F3ECDB;color:#5A4A28;font-size:12.5px;padding:9px 14px;border-radius:8px;margin-bottom:16px;}
 .drop{border:2px dashed #B7BFC6;border-radius:10px;padding:28px;text-align:center;transition:.15s;}
 .drop.over{border-color:#00A09B;background:#F1FAFA;}
 .drop .big{font-weight:600;margin-bottom:4px;}
 .drop .small{font-size:12.5px;color:#5B7083;margin-bottom:10px;}
 .filelist{font-size:12.5px;color:#3A4A64;margin-top:10px;text-align:left;display:inline-block;}
 button.go{background:#00A09B;color:#fff;border:none;padding:12px 24px;border-radius:8px;
        font-size:14px;font-weight:600;cursor:pointer;margin-top:12px;}
 .browse{display:inline-block;background:#EAF6F6;color:#00A09B;padding:9px 16px;border-radius:6px;
        font-weight:600;font-size:13px;cursor:pointer;}
 input[type=file]{display:none;}
 .filehead{font-family:ui-monospace,Menlo,monospace;font-size:12px;color:#5B7083;margin:20px 0 8px;}
 .summary{background:#EFF5F5;border:1px solid #D9DDE1;border-radius:8px;padding:12px 14px;margin-bottom:14px;font-size:13.5px;}
 .finding{border:1px solid #D9DDE1;border-radius:8px;margin-bottom:12px;overflow:hidden;}
 .bar{height:4px;} .bar.High{background:#B23A2E;} .bar.Medium{background:#B0791C;}
 .bar.Low{background:#5B7083;} .bar.Factual{background:#002B49;}
 .fbody{padding:13px 15px;}
 .ftop{display:flex;align-items:center;gap:8px;margin-bottom:7px;flex-wrap:wrap;}
 .sev{font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:10px;}
 .sev.High{background:#F5E1DE;color:#B23A2E;} .sev.Medium{background:#F3EAD3;color:#B0791C;}
 .sev.Low{background:#EAECEE;color:#5B7083;} .sev.Factual{background:#E4E7EB;color:#002B49;}
 .ftitle{font-weight:600;font-size:14.5px;}
 .fexpl{font-size:13px;color:#3A4A64;margin-bottom:8px;}
 .ref{font-family:ui-monospace,Menlo,monospace;font-size:11px;background:#F0F5F5;color:#00A09B;
      padding:5px 9px;border-radius:4px;border-left:3px solid #00A09B;margin-bottom:8px;display:inline-block;}
 .fix{font-size:12.5px;color:#3A4A64;background:#F7F9F9;border:1px solid #eee;border-radius:6px;padding:8px 10px;}
 .fix b{color:#002B49;}
 .dl{display:flex;gap:10px;margin:6px 0 14px;flex-wrap:wrap;}
 .dl a{background:#002B49;color:#fff;text-decoration:none;padding:9px 16px;border-radius:6px;
       font-size:13px;font-weight:600;}
 .dl a.x{background:#1F6B4F;}
 .disclaimer{margin-top:18px;padding:14px;background:#FBF6EE;border:1px solid #E8D9BE;border-radius:8px;
       font-size:12px;color:#5A4A28;line-height:1.55;}
 .err{background:#FBEAE8;border:1px solid #E4B4AD;color:#B23A2E;padding:13px;border-radius:8px;
       font-size:13px;white-space:pre-wrap;margin-bottom:12px;}
 .wait{font-size:12.5px;color:#5B7083;margin-top:8px;}
 .overall{background:#EAF5F5;border:1px solid #BFE0DE;border-radius:10px;padding:16px 18px;margin-bottom:18px;}
 .ov-title{font-weight:700;font-size:14px;color:#00706C;margin-bottom:6px;}
 .ov-body{font-size:13.5px;}
 .ov-sub{font-weight:600;font-size:12.5px;margin-top:10px;}
 .ov-list{margin:4px 0 0 18px;font-size:12.5px;color:#3A4A64;}
 .cnt{font-size:10px;font-weight:700;padding:2px 7px;border-radius:9px;margin-left:6px;}
 .cnt.h{background:#F5E1DE;color:#B23A2E;} .cnt.m{background:#F3EAD3;color:#B0791C;}
 .cnt.l{background:#EAECEE;color:#5B7083;} .cnt.f{background:#E4E7EB;color:#002B49;}
 .conclusion{background:#FDF9F0;border:1px solid #EADFC6;border-radius:8px;padding:11px 13px;margin-bottom:14px;font-size:13px;}
</style></head><body>
<div class="band">
 <div class="top">
  <div class="brand"><div class="logo"><img src="https://www.bakertilly.pk/assets/images/logo.svg" alt="Baker Tilly" onerror="this.outerHTML=&quot;<div class=logofb></div>&quot;"></div><h1>AI Audit Reviewer</h1></div>
  <div class="who">Signed in as <b>{{ user }}</b> ({{ 'Full access' if role=='full' else 'Limited access' }})
   <a href="{{ url_for('logout') }}">Log out</a></div>
 </div>
</div>
<div class="wrap">
 <div class="sub">Baker Tilly - {{ 'Financial Statements review' if mode=='fs' else 'Working-paper review' }} - Stage 4
   &nbsp;|&nbsp; <a href="{{ url_for('choose') }}">Change review type</a></div>

 <div class="notice"><b>Note:</b> Reviews are checked against the firm's loaded standards library. Use sample / public data until the tool moves to the firm's own server. Up to {{ maxfiles }} files per batch (each file takes 1-3 minutes; for fastest results review 3-4 at a time).</div>

 <div class="card">
  <form method="POST" enctype="multipart/form-data" id="upform">
   <div class="drop" id="drop">
    <div class="big">Drag &amp; drop {{ 'financial statements' if mode=='fs' else 'working papers' }} here</div>
    <div class="small">Excel (.xlsx), Word (.docx), PDF, or CSV - up to {{ maxfiles }} files</div>
    <label class="browse">Browse files<input type="file" id="fileinput" name="files" multiple
      accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt"></label>
    <div class="filelist" id="filelist"></div>
   </div>
   <div class="instr-label">Instructions for the AI (optional)</div>
   <textarea class="instr" name="instructions" maxlength="2000"
     placeholder="e.g. Focus on cut-off testing near year end, or: Explain the related-party issue in the revenue file"></textarea>
   <div class="instr-hint">Tell the reviewer what to focus on or ask a question about the files. Leave blank for a full standard review.</div>
   <div style="text-align:center;">
     <button class="go" type="submit">Review selected files</button>
     <div class="wait">Reviews take 1-3 minutes per file. Please leave the page open and wait.</div>
   </div>
  </form>
 </div>

 {% if error %}<div class="err">{{ error }}</div>{% endif %}

 {% if batch %}
  <div class="card">
   <h2 style="font-size:17px;margin:0 0 10px;">Review Points</h2>
   {% if role == 'full' %}
   <div class="dl">
     <a class="x" href="{{ url_for('download_excel', rid=batch_id) }}">Download Excel</a>
     <a href="{{ url_for('download_pdf', rid=batch_id) }}">Download PDF</a>
   </div>
   {% endif %}
   {% if batch.get('overall') %}
     <div class="overall">
       <div class="ov-title">Overall batch conclusion</div>
       <div class="ov-body">{{ batch['overall'].get('overall_conclusion','') }}</div>
       {% if batch['overall'].get('common_themes') %}
         <div class="ov-sub">Common themes across files:</div>
         <ul class="ov-list">{% for t in batch['overall']['common_themes'] %}<li>{{ t }}</li>{% endfor %}</ul>
       {% endif %}
       {% if batch['overall'].get('cross_file_observations') %}
         <div class="ov-sub">Cross-file observations (corroboration):</div>
         <ul class="ov-list">{% for t in batch['overall']['cross_file_observations'] %}<li>{{ t }}</li>{% endfor %}</ul>
       {% endif %}
     </div>
   {% endif %}
   {% for item in batch['files'] %}
     <div class="filehead">FILE: {{ item['filename'] }}
       {% if item.get('counts') %}
         <span class="cnt h">{{ item['counts']['High'] }} High</span>
         <span class="cnt m">{{ item['counts']['Medium'] }} Med</span>
         <span class="cnt l">{{ item['counts']['Low'] }} Low</span>
         <span class="cnt f">{{ item['counts']['Factual'] }} Factual</span>
       {% endif %}
     </div>
     {% if item.get('error') %}
       <div class="err">{{ item['error'] }}</div>
     {% else %}
       {% if item['result'].get('summary') %}
         <div class="summary"><b>Overall:</b> {{ item['result']['summary'] }}</div>
       {% endif %}
       {% if item['result'].get('conclusion') %}
         <div class="conclusion"><b>Head-wise conclusion:</b> {{ item['result']['conclusion'] }}</div>
       {% endif %}
       {% for f in item['result'].get('findings', []) %}
        <div class="finding">
         <div class="bar {{ f.get('severity','Low') }}"></div>
         <div class="fbody">
          <div class="ftop"><span class="ftitle">{{ f.get('title','') }}</span>
            <span class="sev {{ f.get('severity','Low') }}">{{ f.get('severity','') }}</span></div>
          <div class="fexpl">{{ f.get('explanation','') }}</div>
          {% if f.get('reference') %}<div class="ref">{{ f['reference'] }}</div>{% endif %}
          <div class="fix"><b>Suggested fix:</b> {{ f.get('fix','') }}</div>
         </div>
        </div>
       {% endfor %}
     {% endif %}
   {% endfor %}
   <div class="disclaimer"><b>Professional judgement statement:</b> {{ disclaimer }}</div>
  </div>
 {% endif %}
</div>

<script>
const drop = document.getElementById('drop');
const input = document.getElementById('fileinput');
const list = document.getElementById('filelist');
const MAXF = {{ maxfiles }};

function showFiles(files){
  if(!files || files.length===0){ list.innerHTML=''; return; }
  let html = '';
  const n = Math.min(files.length, MAXF);
  for(let i=0;i<n;i++){ html += '&#128196; ' + files[i].name + '<br>'; }
  if(files.length > MAXF){ html += '<i>(only the first ' + MAXF + ' will be reviewed)</i>'; }
  list.innerHTML = html;
}
input.addEventListener('change', () => showFiles(input.files));
['dragover','dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => {
  if(e.dataTransfer.files.length){ input.files = e.dataTransfer.files; showFiles(input.files); }
});
</script>
</body></html>
"""


@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        users = load_users()
        name = request.form.get("username", "").strip()
        pw = request.form.get("password", "")
        u = users.get(name)
        if u and u["password"] == pw:
            session["user"] = name
            session["role"] = u["role"]
            return redirect(url_for("choose"))
        error = "Incorrect username or password."
    return render_template_string(LOGIN_PAGE, error=error)


@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for("login"))


@app.route("/choose")
@login_required
def choose():
    return render_template_string(CHOOSE_PAGE, user=session.get("user"))


@app.route("/select/<mode>")
@login_required
def select_mode(mode):
    if mode not in ("fs", "wp"):
        return redirect(url_for("choose"))
    session["mode"] = mode
    return redirect(url_for("home"))


@app.route("/", methods=["GET", "POST"])
@login_required
def home():
    if session.get("mode") not in ("fs", "wp"):
        return redirect(url_for("choose"))
    error = None
    batch = None
    batch_id = None

    if request.method == "POST":
        if not DEEPSEEK_API_KEY:
            error = "The DeepSeek API key is not set. Add it in Render's Environment Variables."
        else:
            uploads = [f for f in request.files.getlist("files") if f and f.filename]
            if not uploads:
                error = "Please choose at least one file."
            else:
                uploads = uploads[:MAX_FILES_PER_BATCH]
                batch = {"files": []}
                import gc
                for up in uploads:
                    entry = {"filename": up.filename}
                    try:
                        data = up.read()
                        text = extract_text_from_file(up.filename, data)
                        del data  # release the raw file bytes immediately
                        if text is None:
                            entry["error"] = "Unsupported file type."
                        elif not text.strip():
                            entry["error"] = ("The file appears to be empty, or its text "
                                              "could not be read (a scanned PDF with no "
                                              "text layer, perhaps).")
                        else:
                            entry["excerpt"] = text[:3000]
                            result, ai_err = review_with_ai(
                                text, mode=session.get("mode", "wp"),
                                user_instructions=request.form.get("instructions", ""))
                            del text  # release the extracted text
                            if ai_err:
                                entry["error"] = ai_err
                            else:
                                entry["result"] = result
                    except Exception as e:
                        entry["error"] = "Could not process this file. Details: " + str(e)
                    batch["files"].append(entry)
                    gc.collect()  # reclaim memory before the next file

                # severity counts per file (computed here, not by the AI)
                for item in batch["files"]:
                    counts = {"High": 0, "Medium": 0, "Low": 0, "Factual": 0}
                    for f in item.get("result", {}).get("findings", []):
                        sev = f.get("severity", "")
                        if sev in counts:
                            counts[sev] += 1
                    item["counts"] = counts

                # batch-level conclusion + cross-file corroboration (2+ files)
                if len(batch["files"]) > 1:
                    batch["overall"] = batch_conclusion_with_ai(batch)
                for item in batch["files"]:
                    item.pop("excerpt", None)
                batch_id = save_results(batch)

    return render_template_string(MAIN_PAGE, user=session.get("user"),
                                  role=session.get("role"), error=error,
                                  batch=batch, batch_id=batch_id,
                                  maxfiles=MAX_FILES_PER_BATCH,
                                  mode=session.get("mode", "wp"),
                                  disclaimer=DISCLAIMER)


@app.route("/download/excel/<rid>")
@full_access_required
def download_excel(rid):
    batch = load_results(rid)
    if not batch:
        return "This report has expired. Please run the review again.", 404
    buf = build_excel(batch)
    return send_file(buf, as_attachment=True,
                     download_name="review_points.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument."
                              "spreadsheetml.sheet")


@app.route("/download/pdf/<rid>")
@full_access_required
def download_pdf(rid):
    batch = load_results(rid)
    if not batch:
        return "This report has expired. Please run the review again.", 404
    buf = build_pdf(batch)
    return send_file(buf, as_attachment=True,
                     download_name="review_points.pdf",
                     mimetype="application/pdf")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
