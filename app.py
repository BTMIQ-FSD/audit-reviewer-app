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
import time
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
@app.context_processor
def _engine_ctx():
    cur = current_engine()
    other = next((k for k in ENGINES if k != cur), None)
    return {"engine_current": ENGINES.get(cur, {}).get("label", ""),
            "engine_other_key": other,
            "engine_other_label": ENGINES.get(other, {}).get("label", "") if other else "",
            "engine_multi": len(ENGINES) > 1}


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
                                  history=[],
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
                                      history=[],
                                      disclaimer=DISCLAIMER), 500
    except Exception:
        return msg, 500

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
ANTHROPIC_API_KEY = os.environ.get("ANTHROPIC_API_KEY", "")
# Dual-engine setup: both providers are available side by side when their keys
# are set. The person switches engines in the app (stored per login session);
# AI_PROVIDER only sets the default. Model names overridable via env:
#   DEEPSEEK_MODEL (default deepseek-v4-flash — the old deepseek-chat alias
#   was retired by DeepSeek on 24 Jul 2026), CLAUDE_MODEL (default sonnet).
# timeout: never wait more than 150s per attempt (1 retry) so a hung response
# fails with a friendly message instead of a killed worker.
ENGINES = {}
if DEEPSEEK_API_KEY:
    ENGINES["deepseek"] = {
        "client": OpenAI(api_key=DEEPSEEK_API_KEY,
                         base_url="https://api.deepseek.com",
                         timeout=150.0, max_retries=1),
        "model": os.environ.get("DEEPSEEK_MODEL",
                 os.environ.get("AI_MODEL", "deepseek-v4-flash")),
        "label": "DeepSeek (fast & cheap)"}
if ANTHROPIC_API_KEY:
    ENGINES["claude"] = {
        "client": OpenAI(api_key=ANTHROPIC_API_KEY,
                         base_url="https://api.anthropic.com/v1/",
                         timeout=150.0, max_retries=1),
        "model": os.environ.get("CLAUDE_MODEL", "claude-sonnet-5"),
        "label": "Claude (best quality)"}
AI_PROVIDER = os.environ.get("AI_PROVIDER", "deepseek").strip().lower()
DEFAULT_ENGINE = AI_PROVIDER if AI_PROVIDER in ENGINES else (
    "deepseek" if "deepseek" in ENGINES else
    ("claude" if "claude" in ENGINES else ""))
AI_KEY_SET = bool(ENGINES)


def current_engine():
    e = session.get("engine") if session else None
    return e if e in ENGINES else DEFAULT_ENGINE


class EmptyAIResponse(Exception):
    pass


def ai_chat(messages, max_tokens=6000, temperature=0.2, cheap=False):
    """One door to both engines. cheap=True routes tiny classification jobs
    to DeepSeek even when Claude is selected (no point paying Claude prices
    to answer yes/no questions)."""
    name = current_engine()
    if cheap and "deepseek" in ENGINES:
        name = "deepseek"
    eng = ENGINES[name]
    kwargs = {"model": eng["model"], "messages": messages,
              "max_tokens": max_tokens, "temperature": temperature}
    if name == "deepseek":
        # V4 models default to "thinking" mode, which can consume the whole
        # response budget internally and return EMPTY content. The retired
        # deepseek-chat name was the NON-thinking mode, so we ask for that
        # explicitly (per DeepSeek's docs: set thinking mode explicitly).
        kwargs["extra_body"] = {"thinking": {"type": "disabled"}}
    response = eng["client"].chat.completions.create(**kwargs)
    content = ""
    try:
        content = (response.choices[0].message.content or "").strip()
    except Exception:
        content = ""
    if not content:
        # known DeepSeek V4 issue: intermittent completely empty responses.
        # One retry usually clears it; if not, fail with a clear name.
        response = eng["client"].chat.completions.create(**kwargs)
        try:
            content = (response.choices[0].message.content or "").strip()
        except Exception:
            content = ""
        if not content:
            raise EmptyAIResponse("the AI returned an empty response twice")
    return response


# kept for backward references
AI_MODEL = ENGINES.get(DEFAULT_ENGINE, {}).get("model", "")
client = ENGINES.get(DEFAULT_ENGINE, {}).get("client")

MAX_FILES_PER_BATCH = 8
MAX_EXTRACT_CHARS = 120000

# Head-wise segregation for Complete Client Review
HEADS = [
    ("nca", "Non-current assets",
     "property plant and equipment, operating fixed assets, depreciation, CWIP/capital work in progress, intangibles, long-term investments, long-term deposits and advances, right-of-use assets"),
    ("ca", "Current assets",
     "inventories/stock in trade, stores and spares, trade receivables/debtors, advances and prepayments, other receivables, tax refunds due, short-term investments, cash and bank balances"),
    ("eq", "Equity",
     "share capital, reserves, unappropriated profit/retained earnings, revaluation surplus, statement of changes in equity"),
    ("ncl", "Non-current liabilities",
     "long-term financing/loans, lease liabilities, deferred taxation liability, long-term provisions, staff retirement benefits/gratuity"),
    ("cl", "Current liabilities",
     "trade and other payables/creditors, accrued liabilities and mark-up, short-term borrowings, current portion of long-term debt, taxes payable"),
    ("rev", "Revenue",
     "sales, local and export revenue, revenue recognition and cut-off, rebates and discounts, contract assets and liabilities"),
    ("exp", "Expenses and other income",
     "cost of sales, administrative expenses, distribution/selling expenses, finance cost, other operating expenses, other income, payroll expense testing"),
]
# File sections that sit alongside the balance-sheet/P&L heads
SECTIONS = [
    ("plan", "Planning",
     "engagement acceptance and continuance, engagement letter, independence confirmations, planning memorandum, materiality and performance materiality computation, risk assessment, fraud risk, planned audit responses, team allocation and timetable"),
    ("fin", "Finalisation",
     "summary of uncorrected misstatements vs materiality, subsequent events review, going concern assessment, management representation letter, final analytical review, partner completion checklist, archiving checklist"),
    ("perm", "Permanent file",
     "certificate of incorporation, memorandum and articles of association, statutory forms and registers, loan and lease agreements, title documents, related party register, prior year financial statements, continuing-relevance updates"),
]
ALL_AREAS = HEADS + SECTIONS
HEAD_NAMES = {k: n for k, n, _ in ALL_AREAS}

# extra review focus for the file sections (appended to the standard WP brain)
SECTION_FOCUS = {
    "plan": ("SECTION FOCUS - PLANNING: check that materiality and performance "
             "materiality are computed, the benchmark and percentages are stated "
             "and consistent with the client financial statements; identified "
             "risks are each linked to a planned response; the engagement letter "
             "and independence confirmations exist and are for the CURRENT year; "
             "planning sign-offs are dated BEFORE the fieldwork/review dates; and "
             "the planning memorandum covers the entity's business, applicable "
             "framework, and team allocation."),
    "fin": ("SECTION FOCUS - FINALISATION: check that the summary of uncorrected "
            "misstatements is totalled and compared with materiality; subsequent "
            "events procedures extend to the audit report date; the going concern "
            "assessment is documented and consistent with the financial "
            "statements; the management representation letter is dated the same "
            "date as (or immediately before) the audit report; all review notes "
            "and open points are cleared before archiving; and the completion "
            "checklist is signed."),
    "perm": ("SECTION FOCUS - PERMANENT FILE: check that documents are current, "
             "legible and complete; loan and lease agreement terms (amounts, "
             "rates, security) agree with the financial statement disclosures; "
             "statutory records agree with issued share capital per the FS; "
             "related party register is consistent with FS related party "
             "disclosures; and outdated items are flagged for update."),
}

# which heads inform each other's review (evidence library cross-feeding)
RELATED_HEADS = {
    "nca": ["exp", "eq", "ncl"],   # depreciation, revaluation surplus, borrowing costs
    "ca":  ["rev", "cl", "exp"],   # debtors<->revenue, advances from customers, ECL charge
    "eq":  ["nca", "exp"],         # revaluation surplus, profit movement
    "ncl": ["exp", "nca"],         # finance cost, financed assets
    "cl":  ["ca", "rev", "exp"],   # advances<->debtors/revenue, accruals
    "rev": ["ca", "cl"],           # revenue<->debtors, advances from customers
    "exp": ["nca", "ncl", "rev"],  # depreciation, finance cost, cost vs revenue
}
ANCHOR_CHARS = 20000  # how much of the anchor (e.g. signed FS) each review sees
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
    """Load every .txt in knowledge/ twice over: as one full text (small-library
    mode) and as scored chunks (retrieval mode for the full Stage-5 library)."""
    kb_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge")
    sections = []
    chunks = []          # [(source_filename, chunk_text)]
    if os.path.isdir(kb_dir):
        for fname in sorted(os.listdir(kb_dir)):
            if fname.endswith(".txt"):
                try:
                    with open(os.path.join(kb_dir, fname), "r", encoding="utf-8") as f:
                        text = f.read().strip()
                    sections.append(text)
                    # chunk on blank-line paragraph groups, ~1500 chars each,
                    # so retrieval can serve just the relevant sections later
                    paras = [p.strip() for p in text.split("\n\n") if p.strip()]
                    buf = ""
                    for p in paras:
                        if buf and len(buf) + len(p) > 1500:
                            chunks.append((fname, buf))
                            buf = p
                        else:
                            buf = (buf + "\n\n" + p) if buf else p
                    if buf:
                        chunks.append((fname, buf))
                except Exception:
                    pass
    return "\n\n==========\n\n".join(sections), chunks


KNOWLEDGE_BASE, KNOWLEDGE_CHUNKS = load_knowledge_base()

# When the whole library fits comfortably in one request, send it all (today's
# behaviour, zero quality change). Beyond that — the full Stage-5 library —
# retrieval kicks in automatically and sends only the most relevant sections.
KNOWLEDGE_SEND_ALL_LIMIT = 30000   # chars
KNOWLEDGE_BUDGET = 18000           # chars of selected sections per request

_STOPWORDS = set(("the and for are with that this from have has been must should "
                  "shall was were will would can could may might not its all any "
                  "each per was into over under between more than when where which "
                  "who whom these those such only also other than then them they "
                  "there here what your our his her out but use used using does do "
                  "did is in on of to a an as at by or if be it we you no yes").split())


def _kb_tokens(text):
    import re as _r
    return [w for w in _r.findall(r"[a-z0-9]+", text.lower())
            if len(w) > 2 and w not in _STOPWORDS]


def _build_kb_index():
    """Per-chunk token counts + document frequencies for simple tf-idf scoring."""
    import math
    from collections import Counter
    counts = []
    df = Counter()
    for _src, chunk in KNOWLEDGE_CHUNKS:
        c = Counter(_kb_tokens(chunk))
        counts.append(c)
        for t in c:
            df[t] += 1
    n = max(len(counts), 1)
    idf = {t: math.log(1 + n / (1 + d)) for t, d in df.items()}
    return counts, idf


_KB_COUNTS, _KB_IDF = _build_kb_index()


def select_knowledge(query_text, budget=KNOWLEDGE_BUDGET):
    """Return the library text to send with a request: everything while the
    library is small; the most relevant sections once it is large."""
    if len(KNOWLEDGE_BASE) <= KNOWLEDGE_SEND_ALL_LIMIT or not KNOWLEDGE_CHUNKS:
        return KNOWLEDGE_BASE
    from collections import Counter
    q = Counter(_kb_tokens(str(query_text)[:40000]))
    scored = []
    for i, c in enumerate(_KB_COUNTS):
        s = 0.0
        for t, qn in q.items():
            if t in c:
                s += min(qn, 5) * min(c[t], 5) * _KB_IDF.get(t, 0.0)
        if s > 0:
            scored.append((s, i))
    scored.sort(reverse=True)
    picked = []
    used = 0
    SEP = 14  # separator between selected sections
    for s, i in scored:
        srcname, chunk = KNOWLEDGE_CHUNKS[i]
        piece = "[" + srcname + "]\n" + chunk
        if used + len(piece) + SEP > budget:
            continue
        picked.append(piece)
        used += len(piece) + SEP
        if used >= budget * 0.95:
            break
    if not picked:
        # nothing matched: fall back to the first sections up to budget
        for srcname, chunk in KNOWLEDGE_CHUNKS:
            piece = "[" + srcname + "]\n" + chunk
            if used + len(piece) + 14 > budget:
                break
            picked.append(piece)
            used += len(piece) + 14
    return ("RELEVANT SECTIONS SELECTED FROM THE FIRM'S FULL STANDARDS LIBRARY "
            "(sources in brackets):\n\n" + "\n\n----------\n\n".join(picked))


def _extract_xlsx_lightweight(file_bytes, include_hidden=True):
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

        # sheet registry: names, order, hidden state, and file targets
        # (workbook.xml is authoritative; rels map sheet ids to xml files)
        REL_NS = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
        rels = {}
        try:
            if "xl/_rels/workbook.xml.rels" in names:
                with z.open("xl/_rels/workbook.xml.rels") as f:
                    for ev, el in iterparse(f, events=("end",)):
                        if el.tag.endswith("Relationship"):
                            tgt = el.get("Target", "")
                            if tgt.startswith("/"):
                                tgt = tgt.lstrip("/")
                            elif not tgt.startswith("xl/"):
                                tgt = "xl/" + tgt
                            rels[el.get("Id", "")] = tgt
                        el.clear()
        except Exception:
            rels = {}

        sheets = []  # (title, path, hidden)
        try:
            if "xl/workbook.xml" in names:
                with z.open("xl/workbook.xml") as f:
                    idx = 0
                    for ev, el in iterparse(f, events=("end",)):
                        if el.tag == NS + "sheet":
                            idx += 1
                            title = el.get("name", f"Sheet{idx}")
                            hidden = el.get("state", "visible") in ("hidden", "veryHidden")
                            path = rels.get(el.get(REL_NS + "id", ""), "")
                            sheets.append((title, path, hidden))
                        el.clear()
        except Exception:
            sheets = []
        if not sheets or not any(p in names for _t, p, _h in sheets):
            sheets = [(f"Sheet{i}", n, False) for i, n in enumerate(sorted(
                n for n in names
                if _re.match(r"xl/worksheets/sheet\d+\.xml$", n)), start=1)]

        skipped_hidden = []
        for title, sname, hidden in sheets:
            if hidden and not include_hidden:
                skipped_hidden.append(title)
                continue
            if sname not in names:
                continue
            if total >= MAX_EXTRACT_CHARS:
                text_parts.append("\n[... file is large; remaining sheets not "
                                  "included in this review pass ...]")
                break
            header = ("\n===== SHEET: " + title
                      + (" (hidden)" if hidden else "") + " =====")
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

        if skipped_hidden:
            text_parts.append("\n[Hidden sheets skipped by user setting: "
                              + ", ".join(skipped_hidden[:20]) + "]")

    return "\n".join(text_parts)


def extract_text_from_file(filename, file_bytes, include_hidden=True):
    name = filename.lower()

    if name.endswith((".xlsx", ".xlsm")):
        return _extract_xlsx_lightweight(file_bytes, include_hidden=include_hidden)

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

QUALITY RULES (these determine whether the review is professional or noise):
- Report each underlying issue exactly ONCE. If the same problem appears in several sheets or places, write ONE finding and list the affected locations inside it. Never raise the same conclusion, comment, or error twice under different titles.
- NEVER include an item that turns out to be fine. If you check something and it is correct, silently omit it — a finding must always identify a real problem needing action.
- Cite a standard only when it genuinely governs that specific issue. If no loaded standard directly applies, use "outside loaded library - reference to be confirmed" rather than stretching an unrelated standard to fit.
- Order findings from most important to least important (highest risk to the audit first).
- A maximum of 20 findings. If more issues exist, keep the 19 most important and combine the remaining minor points into one final finding titled "Other minor matters" that lists them briefly.
- Keep each explanation specific and tight: name the sheet/cell/figure and state the problem in at most ~60 words. Do not think out loud, do not narrate calculations that turned out correct.
- Severity discipline: High = could indicate material misstatement or makes the work unreliable; Medium = documentation/consistency weakness; Low = minor improvement; Factual = broken values or arithmetic that is simply wrong.

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

QUALITY RULES (these determine whether the review is professional or noise):
- Report each underlying issue exactly ONCE. If the same problem appears in several statements or notes, write ONE finding and list the affected locations inside it. Never raise the same issue twice under different titles.
- NEVER include an item that turns out to be fine. If you check something and it is correct, silently omit it — a finding must always identify a real problem needing action.
- Cite a standard only when it genuinely governs that specific issue. If no loaded standard directly applies, use "outside loaded library - reference to be confirmed" rather than stretching an unrelated standard to fit.
- Order findings from most important to least important (highest risk first).
- A maximum of 20 findings. If more issues exist, keep the 19 most important and combine the remaining minor points into one final finding titled "Other minor matters" that lists them briefly.
- Keep each explanation specific and tight: name the statement/note/figure and state the problem in at most ~60 words. Do not think out loud, do not narrate checks that turned out correct.
- Severity discipline: High = could indicate material misstatement or prevents a true and fair view; Medium = disclosure/presentation weakness; Low = minor improvement; Factual = broken values or arithmetic that is simply wrong.

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


def review_with_ai(document_text, mode="wp", user_instructions="",
                   anchor_name="", anchor_text=""):
    trimmed = document_text[:MAX_EXTRACT_CHARS]
    instructions = FS_REVIEWER_INSTRUCTIONS if mode == "fs" else REVIEWER_INSTRUCTIONS
    doc_label = ("financial statements" if mode == "fs" else "working paper")
    messages = [
        {"role": "system", "content": instructions},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY (check against these texts):\n\n"
            + select_knowledge(trimmed[:30000])},
    ]
    if anchor_text:
        messages.append({"role": "system", "content":
            "REFERENCE / ANCHOR DOCUMENT (\"" + anchor_name + "\") — extract:\n\n"
            + anchor_text +
            "\n\nCROSS-CHECK REQUIREMENT: review the document below AGAINST this "
            "reference document, in addition to the normal review. Specifically:\n"
            "(1) TIE-OUTS: amounts and totals that should agree between the two "
            "(e.g. lead schedules vs the face of the statements and notes) — report "
            "any that do not agree, quoting BOTH figures;\n"
            "(2) CONTRADICTIONS: matters disclosed or stated in one document but "
            "denied, ignored, or treated inconsistently in the other;\n"
            "(3) IMPOSSIBLE DATES: dates in the document under review that are "
            "impossible or illogical relative to the reference (e.g. audit evidence "
            "dated after the audit report date, transactions after year end included "
            "in the year);\n"
            "(4) OMISSIONS: items appearing in one document that are unexplainably "
            "missing from the other.\n"
            "Tie-out failures and direct contradictions are High severity. In such "
            "findings, name both documents so the team can locate the difference."})
    if user_instructions.strip():
        messages.append({"role": "system", "content":
            "SPECIFIC INSTRUCTIONS FROM THE REVIEWER FOR THIS BATCH (follow these, "
            "give the requested areas extra attention, and answer any questions asked "
            "within your findings or summary — but still report any other significant "
            "issues you notice):\n\n" + user_instructions.strip()[:2000]})
    messages.append({"role": "user", "content":
        "Here is the " + doc_label + " to review:\n\n" + trimmed})
    try:
        response = ai_chat(messages, max_tokens=6000, temperature=0.2)
    except Exception as e:
        err_name = type(e).__name__
        if err_name == "EmptyAIResponse":
            return None, ("The AI returned an empty answer for this request "
                          "(a known DeepSeek issue). Please press the review "
                          "button again — or switch to the other engine.")
        if "Timeout" in err_name or "timeout" in str(e).lower():
            return None, ("The AI service took too long to respond for this file "
                          "(over 5 minutes including a retry). This is usually "
                          "temporary — please try this file again in a few minutes.")
        return None, ("The AI service could not be reached for this file. "
                      "Please try again shortly. Details: " + err_name)
    raw = response.choices[0].message.content.strip()
    return parse_ai_json(raw)


def head_review_with_ai(document_text, head_key, prior_points,
                        user_instructions="", related_docs=None,
                        fs_name="", fs_text=""):
    """Review one file inside a client head: guard the head, re-check open
    points against new evidence, and raise new findings."""
    head_name = HEAD_NAMES.get(head_key, "")
    examples = next((e for k, n, e in ALL_AREAS if k == head_key), "")
    all_heads = "; ".join(n + " (" + e + ")" for k, n, e in ALL_AREAS)
    trimmed = document_text[:MAX_EXTRACT_CHARS]
    messages = [
        {"role": "system", "content": REVIEWER_INSTRUCTIONS},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY (check against these texts):\n\n"
            + select_knowledge(head_name + " " + examples + " " + trimmed[:30000])},
        {"role": "system", "content":
            "HEAD CONTEXT: this review belongs to the head \"" + head_name + "\" "
            "(covers: " + examples + ").\n\n"
            "STEP 0 - HEAD CHECK (do this first): decide which head the document "
            "belongs to, from this list: " + all_heads + ". If it clearly belongs "
            "to a DIFFERENT head than \"" + head_name + "\", return ONLY this JSON "
            "and nothing else: {\"wrong_head\": \"<name of the head it belongs to>\"}. "
            "If it belongs here (or is genuinely ambiguous), proceed with the review.\n\n"
            "OUTPUT FORMAT OVERRIDE: return a JSON object with keys: "
            "\"findings\" (as instructed above, NEW issues only), "
            "\"point_updates\" (see below; [] if none), "
            "\"summary\", \"conclusion\"."},
    ]
    if head_key in SECTION_FOCUS:
        messages.append({"role": "system", "content": SECTION_FOCUS[head_key]})
    if fs_text:
        messages.append({"role": "system", "content":
            "CLIENT FINANCIAL STATEMENTS (\"" + fs_name + "\") — the master anchor "
            "for this whole engagement — extract:\n\n" + fs_text +
            "\n\nFS CROSS-CHECK REQUIREMENT: in addition to the normal review, "
            "check the document below AGAINST these financial statements:\n"
            "(1) TIE-OUTS: amounts that should agree with the face of the "
            "statements or the notes — report disagreements quoting BOTH figures;\n"
            "(2) CONTRADICTIONS: matters disclosed in the FS but ignored or "
            "treated inconsistently in the working paper, and vice versa;\n"
            "(3) IMPOSSIBLE DATES: evidence dated after the audit report date, or "
            "outside the financial year;\n"
            "(4) OMISSIONS: items in the FS this working paper should cover but "
            "does not, or vice versa.\n"
            "Tie-out failures and direct contradictions with the FS are High "
            "severity; name the FS note/statement and the working paper location "
            "in each such finding."})
    if related_docs:
        parts = []
        for d in related_docs:
            parts.append("--- FROM HEAD \"" + d["head_name"] + "\", FILE \""
                         + d["name"] + "\" (excerpt) ---\n" + d["excerpt"])
        messages.append({"role": "system", "content":
            "RELATED-HEAD DOCUMENTS ALREADY ON RECORD for this client (uploaded "
            "earlier in other heads). Use them to CROSS-CHECK the document under "
            "review: tie the interlinked figures (e.g. trade debtors vs revenue, "
            "advances from customers vs sales/receivables, depreciation vs asset "
            "schedules, finance cost vs borrowings), and raise a finding for any "
            "figure, date, party name, or treatment that is INCONSISTENT between "
            "heads — naming both documents and quoting both figures. Do not "
            "re-review these related documents themselves; they are context.\n\n"
            + "\n\n".join(parts)})
    open_points = [p for p in prior_points
                   if p.get("status", "pending") == "pending"][:30]
    if open_points:
        listing = json.dumps([{"id": p["id"], "title": p.get("title", ""),
                               "explanation": (p.get("explanation", "") or "")[:200]}
                              for p in open_points])
        messages.append({"role": "system", "content":
            "PRIOR OPEN REVIEW POINTS for this head:\n" + listing + "\n\n"
            "The new document may contain evidence or corrections for these. For "
            "each prior point this document speaks to, add an entry to "
            "\"point_updates\": {\"id\": <id>, \"resolution\": \"resolved\" or "
            "\"still_open\", \"comment\": \"short plain-English reason, citing what "
            "the new document shows or still lacks\"}. Judge honestly: say resolved "
            "only when the evidence genuinely settles the point. Do NOT repeat these "
            "prior points inside \"findings\" - findings are for NEW issues only."})
    if user_instructions.strip():
        messages.append({"role": "system", "content":
            "SPECIFIC INSTRUCTIONS FROM THE REVIEWER (follow these):\n\n"
            + user_instructions.strip()[:2000]})
    messages.append({"role": "user", "content":
        "Here is the working paper to review:\n\n" + trimmed})
    try:
        response = ai_chat(messages, max_tokens=6000, temperature=0.2)
    except Exception as e:
        err_name = type(e).__name__
        if err_name == "EmptyAIResponse":
            return None, ("The AI returned an empty answer for this request "
                          "(a known DeepSeek issue). Please press the review "
                          "button again — or switch to the other engine.")
        if "Timeout" in err_name or "timeout" in str(e).lower():
            return None, ("The AI service took too long to respond for this file. "
                          "Please try again in a few minutes.")
        return None, ("The AI service could not be reached. Details: " + err_name)
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
        response = ai_chat(
            [
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


HISTORY_FILE = os.path.join(RESULTS_DIR, "history.json")


def load_history():
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def _append_history(rid, user, mode, batch):
    hist = load_history()
    entry = {
        "rid": rid,
        "time": time.strftime("%d %b %Y, %H:%M"),
        "user": user,
        "mode": "FS" if mode == "fs" else "WP",
        "files": [it.get("filename", "") for it in batch.get("files", [])],
        "findings": sum(len(it.get("result", {}).get("findings", []))
                        for it in batch.get("files", [])),
    }
    hist.insert(0, entry)
    hist = hist[:30]  # keep the most recent 30 reviews
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(hist, f)
    except Exception:
        pass


def save_results(batch, user="", mode=""):
    rid = uuid.uuid4().hex[:12]
    with open(os.path.join(RESULTS_DIR, rid + ".json"), "w", encoding="utf-8") as f:
        json.dump(batch, f)
    _append_history(rid, user, mode, batch)
    return rid


def update_results(rid, batch):
    safe = "".join(c for c in rid if c.isalnum())
    with open(os.path.join(RESULTS_DIR, safe + ".json"), "w", encoding="utf-8") as f:
        json.dump(batch, f)


CLIENTS_FILE = os.path.join(RESULTS_DIR, "clients.json")
CLIENT_FILES_DIR = os.path.join(RESULTS_DIR, "clientfiles")

# ---- permanent client storage (automatic, survives restarts/deploys) ----
# Set MONGODB_URI in Render's Environment to a free MongoDB Atlas connection
# string and all client workspaces + stored files save there automatically.
# Without it, the app falls back to the free server's temporary disk.
MONGODB_URI = os.environ.get("MONGODB_URI", "")
_mongo = None
_gridfs = None
if MONGODB_URI:
    try:
        from pymongo import MongoClient
        import gridfs as _gridfs_mod
        _mongo = MongoClient(MONGODB_URI, serverSelectionTimeoutMS=8000)["audit_reviewer"]
        _gridfs = _gridfs_mod.GridFS(_mongo)
        _mongo["meta"].find_one({"_id": "ping"})  # fail fast if unreachable
    except Exception:
        _mongo = None
        _gridfs = None
PERMANENT_STORE = _mongo is not None
os.makedirs(CLIENT_FILES_DIR, exist_ok=True)
MAX_STORED_FILE = 15 * 1024 * 1024  # keep originals up to 15 MB each


def _safe_ext(filename):
    ext = os.path.splitext(filename)[1].lower()
    return ext if _re_ext.match(ext or "") else ""


import re as _re_mod
_re_ext = _re_mod.compile(r"^\.[a-z0-9]{1,6}$")


def store_client_file(cid, filename, raw):
    """Keep the original upload so it can be opened later. Returns (fid, ext)."""
    try:
        if not raw or len(raw) > MAX_STORED_FILE:
            return None, ""
        fid = uuid.uuid4().hex[:12]
        ext = _safe_ext(filename)
        if PERMANENT_STORE:
            _gridfs.put(raw, filename=fid + ext,
                        metadata={"cid": cid, "fid": fid, "name": filename})
            return fid, ext
        safe_cid = "".join(c for c in cid if c.isalnum())
        with open(os.path.join(CLIENT_FILES_DIR, safe_cid + "_" + fid + ext), "wb") as f:
            f.write(raw)
        return fid, ext
    except Exception:
        return None, ""


def discard_client_file(cid, fid, ext):
    if not fid:
        return
    if PERMANENT_STORE:
        try:
            for gf in _gridfs.find({"metadata.cid": cid, "metadata.fid": fid}):
                _gridfs.delete(gf._id)
            return
        except Exception:
            pass
    try:
        safe_cid = "".join(c for c in cid if c.isalnum())
        os.remove(os.path.join(CLIENT_FILES_DIR, safe_cid + "_" + fid + ext))
    except Exception:
        pass


def get_client_file(cid, fid, ext):
    """Return the stored original's bytes, or None."""
    if PERMANENT_STORE:
        try:
            gf = _gridfs.find_one({"metadata.cid": cid, "metadata.fid": fid})
            if gf:
                return gf.read()
            return None
        except Exception:
            pass
    try:
        safe_cid = "".join(c for c in cid if c.isalnum())
        path = os.path.join(CLIENT_FILES_DIR, safe_cid + "_" + fid + ext)
        with open(path, "rb") as f:
            return f.read()
    except Exception:
        return None


def load_clients():
    if PERMANENT_STORE:
        try:
            doc = _mongo["meta"].find_one({"_id": "clients"})
            return doc["list"] if doc else []
        except Exception:
            pass
    try:
        with open(CLIENTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return []


def save_clients(clients):
    if PERMANENT_STORE:
        try:
            _mongo["meta"].replace_one({"_id": "clients"},
                                       {"_id": "clients", "list": clients},
                                       upsert=True)
            return
        except Exception:
            pass
    with open(CLIENTS_FILE, "w", encoding="utf-8") as f:
        json.dump(clients, f)


def client_path(cid):
    safe = "".join(c for c in cid if c.isalnum())
    return os.path.join(RESULTS_DIR, "client_" + safe + ".json")


def load_client(cid):
    if PERMANENT_STORE:
        try:
            doc = _mongo["clients"].find_one({"_id": cid})
            if doc:
                doc.pop("_id", None)
                return doc
            return None
        except Exception:
            pass
    try:
        with open(client_path(cid), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def save_client(data):
    if PERMANENT_STORE:
        try:
            doc = dict(data)
            doc["_id"] = data["cid"]
            _mongo["clients"].replace_one({"_id": data["cid"]}, doc, upsert=True)
            return
        except Exception:
            pass
    with open(client_path(data["cid"]), "w", encoding="utf-8") as f:
        json.dump(data, f)


def delete_client_data(cid):
    if PERMANENT_STORE:
        try:
            _mongo["clients"].delete_one({"_id": cid})
            for gf in _gridfs.find({"metadata.cid": cid}):
                _gridfs.delete(gf._id)
        except Exception:
            pass
    try:
        os.remove(client_path(cid))
    except Exception:
        pass
    try:
        safe_cid = "".join(c for c in cid if c.isalnum())
        for fn in os.listdir(CLIENT_FILES_DIR):
            if fn.startswith(safe_cid + "_"):
                os.remove(os.path.join(CLIENT_FILES_DIR, fn))
    except Exception:
        pass


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
    ws.append(["File", "No.", "Title", "Severity", "Status", "Flagged",
               "Explanation", "Reference", "Suggested fix"])
    for c in ws[1]:
        c.font = Font(bold=True)
    if batch.get("overall"):
        ws.append(["BATCH", "-", "OVERALL CONCLUSION", "-", "", "",
                   batch["overall"].get("overall_conclusion", ""), "", ""])
        for t in batch["overall"].get("common_themes", []):
            ws.append(["BATCH", "-", "Common theme", "-", "", "", t, "", ""])
        for t in batch["overall"].get("cross_file_observations", []):
            ws.append(["BATCH", "-", "Cross-file observation", "-", "", "", t, "", ""])
        ws.append([])
    for item in batch["files"]:
        fname = item["filename"]
        if item.get("error"):
            ws.append([fname, "-", "REVIEW ERROR", "-", "", "", item["error"], "-", "-"])
            continue
        if item.get("result", {}).get("conclusion"):
            ws.append([fname, "-", "HEAD-WISE CONCLUSION", "-", "", "",
                       item["result"]["conclusion"], "", ""])
        for i, f in enumerate(item["result"].get("findings", []), start=1):
            ws.append([fname, i, f.get("title", ""), f.get("severity", ""),
                       f.get("status", "pending").capitalize(),
                       "Yes" if f.get("flagged") else "",
                       f.get("explanation", ""), f.get("reference", ""),
                       f.get("fix", "")])
    ws.append([])
    ws.append(["Professional judgement statement:"])
    ws.append([DISCLAIMER])
    widths = [26, 5, 30, 10, 10, 8, 52, 34, 44]
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
            stat = f.get("status", "pending")
            if stat != "pending" or f.get("flagged"):
                story.append(Paragraph(
                    "<i>Status: " + stat.capitalize()
                    + (" | FLAGGED" if f.get("flagged") else "") + "</i>", body))
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



COMMON_UI = """
<style>
 :root{--bt-accent:#00A09B;--bt-navy:#0c1b34;}
 @media(min-width:1400px){
   html{font-size:17px;}
   .wrap,.layout,.cols{max-width:1360px !important;}
   .stage{max-width:940px !important;}
   .choice{padding:34px 26px !important;}
   .choices{max-width:720px !important;gap:22px !important;}
   .brand img{height:56px !important;}
   h1{font-size:26px !important;}
 }
 @media(max-width:700px){
   .top,.band{flex-direction:column;align-items:flex-start !important;gap:6px !important;}
   .who{font-size:11.5px !important;}
   .stat-row .stbtn,.dsend,.go{padding:9px 14px !important;font-size:12.5px !important;}
   .sbar{gap:6px !important;} .sb{font-size:11px !important;padding:5px 9px !important;}
   .grid{grid-template-columns:1fr !important;}
   .choices{grid-template-columns:1fr !important;}
   .drow{flex-direction:column;align-items:stretch !important;}
   .pdrow{flex-direction:column;align-items:stretch !important;}
   h1{font-size:17px !important;}
 }
 .rise{opacity:0;animation:btRise .55s cubic-bezier(.2,.7,.3,1) forwards;}
 @keyframes btRise{from{opacity:0;transform:translateY(16px) scale(.985);}to{opacity:1;transform:translateY(0) scale(1);}}
 .card,.hd,.cl,.choice{transition:transform .22s ease, box-shadow .22s ease, border-color .22s ease;}
 .card:hover{box-shadow:0 6px 22px rgba(12,27,52,.07);}
 .hd:hover,.cl:hover{transform:translateY(-3px);}
 .go:hover,.dsend:hover{filter:brightness(1.07);}
 .bt-dot{position:fixed;border-radius:50%;pointer-events:none;z-index:0;}
 @keyframes btDriftA{0%,100%{transform:translate(0,0) scale(1);}50%{transform:translate(26px,-34px) scale(1.25);}}
 @keyframes btDriftB{0%,100%{transform:translate(0,0) scale(1);}50%{transform:translate(-30px,24px) scale(.8);}}
 @keyframes btDriftC{0%{transform:translate(0,0);opacity:.12;}50%{opacity:.4;}100%{transform:translate(14px,-60px);opacity:.12;}}



 .fontctl{position:fixed;bottom:14px;right:14px;z-index:50;display:flex;gap:6px;
   background:rgba(12,27,52,.85);border-radius:20px;padding:6px 10px;}
 .fontctl button{background:none;border:none;color:#9FE1CB;font-weight:700;font-size:14px;
   cursor:pointer;padding:2px 7px;line-height:1;}
 .fontctl button:hover{color:#fff;}
</style>
<script>
document.addEventListener('DOMContentLoaded', function(){
  var fs = localStorage.getItem('bt_fontsize');
  if(fs){ document.documentElement.style.fontSize = fs + '%'; }
  var ctl = document.createElement('div');
  ctl.className = 'fontctl';
  ctl.innerHTML = '<button type="button" title="Smaller text" aria-label="Smaller text">A\u2212</button>'
                + '<button type="button" title="Larger text" aria-label="Larger text">A+</button>';
  document.body.appendChild(ctl);
  var bs = ctl.querySelectorAll('button');
  function setFs(d){
    var cur = parseInt(localStorage.getItem('bt_fontsize') || '100', 10);
    cur = Math.min(140, Math.max(80, cur + d));
    localStorage.setItem('bt_fontsize', cur);
    document.documentElement.style.fontSize = cur + '%';
  }
  bs[0].onclick = function(){ setFs(-10); };
  bs[1].onclick = function(){ setFs(10); };

  if(document.body.classList.contains('darkbg')){
    var colors = ['#2dd4bf','#5eead4','#7c6cf0','#9FE1CB'];
    var anims = ['btDriftA','btDriftB','btDriftC'];
    for(var i=0;i<16;i++){
      var d = document.createElement('span');
      d.className = 'bt-dot';
      var s = 3 + Math.random()*7;
      d.style.width = s+'px'; d.style.height = s+'px';
      d.style.left = (Math.random()*97)+'vw';
      d.style.top = (Math.random()*94)+'vh';
      d.style.background = colors[i % colors.length];
      d.style.opacity = (0.12 + Math.random()*0.28).toFixed(2);
      d.style.animation = anims[i % anims.length] + ' ' + (7+Math.random()*9).toFixed(1)
                        + 's ease-in-out ' + (Math.random()*4).toFixed(1) + 's infinite';
      document.body.appendChild(d);
    }
  } else {
    var els = document.querySelectorAll('.card,.hd,.cl,.finding,.pt,.side,.fside,.sub,.notice,.sbar');
    var n = 0;
    els.forEach(function(el){
      if(n > 24) return;
      el.classList.add('rise');
      el.style.animationDelay = (n * 0.06).toFixed(2) + 's';
      n++;
    });
  }
});
</script>
"""

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


WP_CHOICE_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Working-paper review : choose scope</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#0c1b34;margin:0;
      min-height:100vh;display:grid;place-items:center;color:#fff;}
 .stage{width:100%;max-width:660px;padding:40px 20px;text-align:center;}
 h1{font-size:22px;font-weight:600;margin:0 0 4px;}
 .sub{color:#9fb3cc;font-size:14px;margin-bottom:28px;}
 .choices{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px;max-width:540px;margin:0 auto;}
 .choice{display:block;text-decoration:none;background:rgba(255,255,255,.05);
         border:1px solid rgba(94,234,212,.18);border-radius:12px;padding:26px 20px;
         transition:transform .25s,border-color .25s,background .25s;}
 .choice:hover{transform:translateY(-6px);border-color:#2dd4bf;background:rgba(45,212,191,.12);}
 .cico{font-size:28px;margin-bottom:10px;}
 .ctitle{color:#fff;font-size:16px;font-weight:600;margin-bottom:6px;}
 .cdesc{color:#9fb3cc;font-size:12.5px;line-height:1.6;}
 .back{display:inline-block;margin-top:24px;color:#5eead4;font-size:13px;text-decoration:none;}
</style></head><body>
<div class="stage">
 <h1>Working-paper review</h1>
 <div class="sub">Choose the scope</div>
 <div class="choices">
  <a class="choice" href="{{ url_for('clients_page') }}">
   <div class="cico">&#128193;</div>
   <div class="ctitle">Complete client review</div>
   <div class="cdesc">A workspace per client with head-wise tabs (assets, liabilities, equity, revenue, expenses) and running review points</div>
  </a>
  <a class="choice" href="{{ url_for('home') }}">
   <div class="cico">&#9889;</div>
   <div class="ctitle">General review</div>
   <div class="cdesc">Quick one-off review of any files, with optional cross-check anchor</div>
  </a>
 </div>
 <a class="back" href="{{ url_for('choose') }}">&larr; Back</a>
</div></body></html>
"""

CLIENTS_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Clients : Complete client review</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#ECEEF0;margin:0;color:#002B49;}
 .band{background:#0c1b34;padding:16px 28px;margin-bottom:20px;color:#fff;
       display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}
 .band h1{font-size:18px;margin:0;}
 .band a{color:#5eead4;font-size:12.5px;text-decoration:none;margin-left:12px;}
 .wrap{max-width:720px;margin:0 auto;padding:0 24px 40px;}
 .card{background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:20px;margin-bottom:16px;}
 .addrow{display:flex;gap:10px;}
 .addrow input{flex:1;padding:10px 12px;border:1px solid #B7BFC6;border-radius:8px;font-size:14px;}
 .addrow button{background:#00A09B;color:#fff;border:none;border-radius:8px;padding:10px 20px;
       font-weight:600;font-size:14px;cursor:pointer;}
 .cl{display:block;text-decoration:none;color:#002B49;border:1px solid #E4E7EB;border-radius:10px;
     padding:14px 16px;margin-bottom:10px;font-weight:600;}
 .cl:hover{border-color:#00A09B;background:#F4FAFA;}
 .cl small{display:block;color:#5B7083;font-weight:400;font-size:11.5px;margin-top:3px;}
 .mini{display:inline-block;border:1px solid #D9DDE1;background:#fff;color:#3A4A64;font-size:11.5px;
       font-weight:600;padding:5px 12px;border-radius:8px;cursor:pointer;text-decoration:none;}
 .mini:hover{border-color:#00A09B;}
 .note{font-size:11px;color:#8595A5;line-height:1.5;}
 .err{background:#FBE9E7;border:1px solid #E5B5AC;color:#8C2F22;border-radius:8px;
      padding:10px 12px;font-size:13px;margin-bottom:12px;}
</style></head><body>
<div class="band"><h1>Complete client review &mdash; Clients</h1>
 <div><span style="font-size:12.5px;color:#9fb3cc;">Signed in as <b style="color:#fff;">{{ user }}</b></span>
  <a href="{{ url_for('wp_choice') }}">&larr; Back</a>
  <a href="{{ url_for('logout') }}">Log out</a></div></div>
<div class="wrap">
 {% if error %}<div class="err">{{ error }}</div>{% endif %}
 <div class="card">
  <form method="post" class="addrow">
   <input name="client_name" maxlength="80" placeholder="New client name, e.g. Gohar Textile Mills (Pvt) Ltd - FY2025" required>
   <button type="submit">Add client</button>
  </form>
 </div>
 <div class="card">
  {% if clients %}
    {% for c in clients %}
      <div class="cl" style="{{ 'opacity:.55;background:#F4F5F6;' if c.get('status')=='closed' else '' }}">
        <a href="{{ url_for('client_heads', cid=c['cid']) }}" style="text-decoration:none;color:#002B49;">
          &#128193; {{ c['name'] }}
          {% if c.get('status')=='closed' %}<span style="background:#EFF2F4;color:#5B7083;font-size:10px;font-weight:700;padding:2px 8px;border-radius:10px;margin-left:6px;vertical-align:middle;">CLOSED</span>{% endif %}
          <small>Created {{ c['created'] }}</small>
        </a>
        <div style="margin-top:8px;display:flex;gap:6px;flex-wrap:wrap;">
          <a href="{{ url_for('client_heads', cid=c['cid']) }}" class="mini">Open</a>
          <button type="button" class="mini" onclick="renameClient('{{ c['cid'] }}', {{ c['name'] | tojson }})">Rename</button>
          <form method="post" action="{{ url_for('client_close', cid=c['cid']) }}" style="display:inline;">
            <button type="submit" class="mini">{{ 'Reopen' if c.get('status')=='closed' else 'Close' }}</button></form>
          <a href="{{ url_for('client_export', cid=c['cid']) }}" class="mini">Backup</a>
          <form method="post" action="{{ url_for('client_delete', cid=c['cid']) }}" style="display:inline;"
                onsubmit="return confirm('Delete {{ c['name'] }} permanently, including all its review points, files and history? This cannot be undone.');">
            <button type="submit" class="mini" style="color:#B23A2E;border-color:#E5B5AC;">Delete</button></form>
        </div>
      </div>
    {% endfor %}
  {% else %}
    <div style="color:#5B7083;font-size:13.5px;">No clients yet &mdash; add your first one above.</div>
  {% endif %}
 </div>

 <div style="text-align:right;margin-top:2px;">
  <a href="#" onclick="document.getElementById('restorebox').style.display='block';this.style.display='none';return false;"
     style="font-size:11px;color:#8595A5;text-decoration:none;">Restore a client from a backup file&hellip;</a>
  <div id="restorebox" style="display:none;text-align:left;background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:14px 16px;margin-top:8px;">
   <form method="post" action="{{ url_for('client_import') }}" enctype="multipart/form-data"
         style="display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
     <input type="file" name="backup" accept=".json" style="font-size:12px;">
     <button type="submit" class="mini" style="background:#00A09B;color:#fff;border-color:#00A09B;">Import backup</button>
   </form>
   <div style="font-size:11px;color:#8595A5;margin-top:6px;line-height:1.5;">Only for _backup.json files made with a client's Backup button &mdash; restores that client's whole workspace after a server restart.</div>
  </div>
 </div>

 <form id="renameForm" method="post" style="display:none;"><input name="new_name" id="renameInput"></form>
 <script>
 function renameClient(cid, current){
   const n = prompt('New client name:', current);
   if(!n || !n.trim()) return;
   const f = document.getElementById('renameForm');
   document.getElementById('renameInput').value = n.trim();
   f.action = '/clients/' + cid + '/rename';
   f.submit();
 }
 </script>
 <div class="note">{% if permanent %}&#9989; Permanent storage is ON &mdash; clients, review points and files save automatically to the firm database and survive server restarts and updates.{% else %}&#9888;&#65039; Automatic permanent storage is OFF &mdash; workspaces live on the free server and are cleared if it restarts or redeploys. Ask your admin to set MONGODB_URI to switch on automatic saving.{% endif %} Use sample / training data only on this free hosting. Client workspaces are kept on the free server and are cleared if it restarts or redeploys &mdash; download reports for permanent records. Permanent storage arrives with the firm's own server (Stage 5).</div>
</div></body></html>
"""

CLIENT_HEADS_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ client['name'] }} : heads</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#ECEEF0;margin:0;color:#002B49;}
 .band{background:#0c1b34;padding:16px 28px;margin-bottom:20px;color:#fff;
       display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}
 .band h1{font-size:17px;margin:0;}
 .band a{color:#5eead4;font-size:12.5px;text-decoration:none;margin-left:12px;}
 .wrap{max-width:820px;margin:0 auto;padding:0 24px 40px;}
 .grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(230px,1fr));gap:14px;}
 .hd{display:block;text-decoration:none;color:#002B49;background:#fff;border:1px solid #D9DDE1;
     border-radius:12px;padding:18px 16px;}
 .hd:hover{border-color:#00A09B;background:#F4FAFA;}
 .hd b{display:block;font-size:14.5px;margin-bottom:6px;}
 .hd small{color:#5B7083;font-size:11px;line-height:1.5;display:block;}
 .badges{margin-top:9px;font-size:10.5px;font-weight:700;}
 .badges span{padding:2px 8px;border-radius:10px;margin-right:5px;display:inline-block;}
 .bp{background:#F3EAD3;color:#B0791C;} .br{background:#E2F2E9;color:#1F6B4F;}
</style></head><body>
<div class="band"><h1>&#128193; {{ client['name'] }}</h1>
 <div><a href="{{ url_for('clients_page') }}">&larr; All clients</a>
  <a href="{{ url_for('logout') }}">Log out</a></div></div>
<div class="wrap">
 {% if error %}<div style="background:#FBE9E7;border:1px solid #E5B5AC;color:#8C2F22;border-radius:8px;padding:11px 13px;font-size:13px;margin-bottom:14px;">{{ error }}</div>{% endif %}
 {% if okmsg %}<div style="background:#E2F2E9;border:1px solid #B5D8C4;color:#1F6B4F;border-radius:8px;padding:11px 13px;font-size:13px;margin-bottom:14px;">{{ okmsg }}</div>{% endif %}
 {% if engine_multi %}<div style="font-size:12px;color:#5B7083;margin-bottom:12px;"><b>AI engine: {{ engine_current }}</b> &mdash; <a style="color:#00A09B;" href="{{ url_for('set_engine', name=engine_other_key) }}">switch to {{ engine_other_label }}</a>. Tip: DeepSeek for first passes and questions; Claude for final reviews, FS review and cross-checks.</div>{% endif %}

 <div style="background:#fff;border:2px solid {{ '#1F6B4F' if client.get('fs') else '#E8B84B' }};border-radius:12px;padding:16px 18px;margin-bottom:18px;">
  <b style="font-size:14px;">&#128209; Financial statements — engagement anchor</b>
  {% if client.get('fs') %}
    <div style="font-size:12.5px;color:#1F6B4F;margin-top:7px;">&#10003; On record: <b>{{ client['fs']['name'] }}</b>
      <span style="color:#8595A5;">(saved {{ client['fs']['time'] }})</span></div>
    <div style="font-size:11.5px;color:#5B7083;margin-top:4px;">Every working paper reviewed in any head is automatically checked against these FS (tie-outs, contradictions, impossible dates, omissions).</div>
    <form method="post" action="{{ url_for('client_fs', cid=client['cid']) }}" enctype="multipart/form-data" style="margin-top:9px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <input type="file" name="fsfile" accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt" style="font-size:12px;">
      <button type="submit" style="background:#EFF2F4;color:#3A4A64;border:1px solid #D9DDE1;border-radius:8px;padding:7px 14px;font-size:12px;font-weight:600;cursor:pointer;">Replace FS</button>
    </form>
    <form method="post" action="{{ url_for('client_fsreview', cid=client['cid']) }}" style="margin-top:8px;"
          onsubmit="this.querySelector('button').textContent='Reviewing FS... please wait';">
      <button type="submit" style="background:#00A09B;color:#fff;border:none;border-radius:8px;padding:8px 16px;font-size:12.5px;font-weight:600;cursor:pointer;">&#128209; Run financial statements review</button>
    </form>
  {% else %}
    <div style="font-size:12.5px;color:#8A5E12;margin-top:7px;">Not uploaded yet — upload the client's financial statements once, and every working paper in every head will automatically be reviewed against them.</div>
    <form method="post" action="{{ url_for('client_fs', cid=client['cid']) }}" enctype="multipart/form-data" style="margin-top:9px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;">
      <input type="file" name="fsfile" accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt" style="font-size:12px;">
      <button type="submit" style="background:#00A09B;color:#fff;border:none;border-radius:8px;padding:8px 16px;font-size:12.5px;font-weight:600;cursor:pointer;">Save FS as anchor</button>
    </form>
  {% endif %}
 </div>

 <p style="font-size:13px;color:#5B7083;">Choose a head to review its working papers:</p>
 <div class="grid">
  {% for k, n, e in heads %}
   {% set pts = client.get('heads', {}).get(k, {}).get('points', []) %}
   {% set open = pts | selectattr('status', 'equalto', 'pending') | list | length %}
   {% set res = pts | selectattr('status', 'equalto', 'resolved') | list | length %}
   <a class="hd" href="{{ url_for('head_page', cid=client['cid'], head=k) }}">
    <b>{{ loop.index }}. {{ n }}</b>
    <small>{{ e[:90] }}...</small>
    {% if pts %}<div class="badges"><span class="bp">{{ open }} open</span><span class="br">{{ res }} resolved</span></div>{% endif %}
   </a>
  {% endfor %}
 </div>

 <p style="font-size:13px;color:#5B7083;margin-top:20px;">File sections:</p>
 <div class="grid">
  {% for k, n, e in sections %}
   {% set pts = client.get('heads', {}).get(k, {}).get('points', []) %}
   {% set open = pts | selectattr('status', 'equalto', 'pending') | list | length %}
   {% set res = pts | selectattr('status', 'equalto', 'resolved') | list | length %}
   <a class="hd" href="{{ url_for('head_page', cid=client['cid'], head=k) }}">
    <b>{{ n }}</b>
    <small>{{ e[:90] }}...</small>
    {% if pts %}<div class="badges"><span class="bp">{{ open }} open</span><span class="br">{{ res }} resolved</span></div>{% endif %}
   </a>
  {% endfor %}
  {% set fpts = client.get('heads', {}).get('fsr', {}).get('points', []) %}
  {% if fpts %}
   <a class="hd" style="border-color:#00A09B;" href="{{ url_for('head_page', cid=client['cid'], head='fsr') }}">
    <b>&#128209; Financial statements review</b>
    <small>The anchored FS reviewed on their own</small>
    <div class="badges"><span class="bp">{{ fpts | selectattr('status','equalto','pending') | list | length }} open</span></div>
   </a>
  {% endif %}
  {% set cpts = client.get('heads', {}).get('cross', {}).get('points', []) %}
  {% if cpts %}
   <a class="hd" style="border-color:#5B4FC0;" href="{{ url_for('head_page', cid=client['cid'], head='cross') }}">
    <b>&#128279; Cross-head checks</b>
    <small>Inconsistencies found between heads</small>
    <div class="badges"><span class="bp">{{ cpts | selectattr('status','equalto','pending') | list | length }} open</span></div>
   </a>
  {% endif %}
 </div>

 <div style="background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:16px 18px;margin-top:18px;">
  <div style="display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:10px;">
   <b style="font-size:14px;">&#128194; Files on record ({{ client.get('library', []) | length }})</b>
   <form method="post" action="{{ url_for('client_crosscheck', cid=client['cid']) }}"
         onsubmit="this.querySelector('button').textContent='Cross-checking... please wait';">
    <button type="submit" style="background:#5B4FC0;color:#fff;border:none;border-radius:8px;
      padding:9px 18px;font-weight:600;font-size:13px;cursor:pointer;">&#128279; Run client-wide cross-check</button>
   </form>
  </div>
  {% if client.get('library') %}
   <div style="margin-top:10px;font-size:12px;color:#3A4A64;line-height:1.8;">
    {% for d in client['library'] | reverse %}
      &#128196; {{ d['name'] }} <span style="color:#8595A5;">({{ head_names.get(d['head'], d['head']) }}, {{ d.get('time','') }})</span>{% if d.get('fid') %} &nbsp;<a href="{{ url_for('open_client_file', cid=client['cid'], fid=d['fid']) }}" style="color:#00706C;font-weight:700;text-decoration:none;">Open</a>{% endif %}<br>
    {% endfor %}
   </div>
   <div style="font-size:11px;color:#8595A5;margin-top:6px;">Every reviewed file is remembered here and automatically feeds related heads &mdash; no need to upload the same file in multiple tabs. Cleared if the free server restarts.</div>
  {% else %}
   <div style="margin-top:8px;font-size:12px;color:#5B7083;">No files on record yet &mdash; review files in any head and they appear here.</div>
  {% endif %}
 </div>
</div></body></html>
"""

HEAD_PAGE = """
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{{ head_name }} : {{ client['name'] }}</title>
<style>
 body{font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:#ECEEF0;margin:0;color:#002B49;}
 .band{background:#0c1b34;padding:14px 28px;margin-bottom:18px;color:#fff;
       display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px;}
 .band h1{font-size:16px;margin:0;}
 .band a{color:#5eead4;font-size:12.5px;text-decoration:none;margin-left:12px;}
 .wrap{max-width:900px;margin:0 auto;padding:0 24px 40px;}
 .card{background:#fff;border:1px solid #D9DDE1;border-radius:12px;padding:18px 20px;margin-bottom:16px;}
 .err{background:#FBE9E7;border:1px solid #E5B5AC;color:#8C2F22;border-radius:8px;
      padding:12px 14px;font-size:13.5px;margin-bottom:14px;}
 .err a{color:#8C2F22;font-weight:700;}
 .okmsg{background:#E2F2E9;border:1px solid #B5D8C4;color:#1F6B4F;border-radius:8px;
      padding:12px 14px;font-size:13.5px;margin-bottom:14px;}
 input[type=file]{font-size:13px;}
 .instr{width:100%;box-sizing:border-box;margin-top:10px;padding:10px 12px;border:1px solid #B7BFC6;
        border-radius:8px;font-size:13px;font-family:inherit;min-height:52px;resize:vertical;}
 .go{background:#00A09B;color:#fff;border:none;border-radius:8px;padding:11px 22px;
     font-weight:600;font-size:14px;cursor:pointer;margin-top:12px;}
 .hint{font-size:11.5px;color:#5B7083;margin-top:8px;line-height:1.5;}
 .sbar{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px;}
 .sb{font-size:12px;font-weight:700;padding:6px 12px;border-radius:16px;background:#EFF2F4;color:#3A4A64;}
 .sb.p{background:#F3EAD3;color:#B0791C;} .sb.r{background:#E2F2E9;color:#1F6B4F;}
 .sb.x{background:#F5E1DE;color:#B23A2E;} .sb.fl{background:#E7E4F7;color:#5B4FC0;}
 .pt{border:1px solid #E4E7EB;border-radius:10px;padding:14px 16px;margin-bottom:12px;}
 .ptop{display:flex;justify-content:space-between;gap:10px;align-items:flex-start;flex-wrap:wrap;}
 .ptitle{font-weight:700;font-size:14px;}
 .sev{font-size:10.5px;font-weight:700;padding:2px 9px;border-radius:10px;}
 .sev.High{background:#F5E1DE;color:#B23A2E;} .sev.Medium{background:#F3EAD3;color:#B0791C;}
 .sev.Low{background:#EFF2F4;color:#3A4A64;} .sev.Factual{background:#E4EFF9;color:#0A3556;}
 .pexpl{font-size:13px;color:#3A4A64;margin:7px 0;line-height:1.6;}
 .pref{font-family:ui-monospace,monospace;font-size:11.5px;background:#EAF6F6;color:#00706C;
       border-left:3px solid #00A09B;padding:6px 9px;border-radius:4px;margin:7px 0;}
 .pfix{font-size:12.5px;color:#3A4A64;background:#F7F9F9;border:1px solid #eee;border-radius:6px;padding:7px 10px;}
 .aiu{font-size:12px;border-radius:6px;padding:7px 10px;margin-top:8px;}
 .aiu.res{background:#E2F2E9;color:#1F6B4F;border:1px solid #B5D8C4;}
 .aiu.open{background:#FBF3E3;color:#8A5E12;border:1px solid #E8D3A3;}
 .meta{font-size:10.5px;color:#8595A5;margin-top:7px;}
 .stat-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px;}
 .stbtn{border:1px solid #D9DDE1;background:#fff;color:#5B7083;font-size:11.5px;
        font-weight:600;padding:5px 11px;border-radius:14px;cursor:pointer;}
 .pt[data-status="pending"] .stbtn.pen{background:#B0791C;color:#fff;border-color:#B0791C;}
 .pt[data-status="resolved"] .stbtn.res{background:#1F6B4F;color:#fff;border-color:#1F6B4F;}
 .pt[data-status="rejected"] .stbtn.rej{background:#B23A2E;color:#fff;border-color:#B23A2E;}
 .pt[data-flag="1"] .stbtn.flg{background:#5B4FC0;color:#fff;border-color:#5B4FC0;}
 .pt[data-status="rejected"] .ptitle,.pt[data-status="rejected"] .pexpl{opacity:.5;text-decoration:line-through;}
 .round{font-size:12px;color:#5B7083;border-left:3px solid #D9DDE1;padding:4px 10px;margin-bottom:8px;}
 .disc-btn{margin-top:10px;background:#EAF6F6;color:#00706C;border:1px solid #BFE0DE;
        padding:6px 13px;border-radius:6px;font-size:12px;font-weight:600;cursor:pointer;}
 .disc-btn:hover{background:#DDF1F0;}
 .pdisc{margin-top:10px;border:1px solid #BFE0DE;border-radius:8px;background:#F7FBFB;padding:10px;}
 .pdlog{max-height:320px;overflow-y:auto;margin-bottom:8px;}
 .dmsg{padding:8px 11px;border-radius:8px;margin-bottom:7px;font-size:12.5px;line-height:1.55;white-space:pre-wrap;}
 .du{background:#E4EFF9;color:#0A3556;margin-left:12%;}
 .da{background:#fff;border:1px solid #D9DDE1;color:#3A4A64;margin-right:12%;}
 .dwait{color:#5B7083;font-size:12px;font-style:italic;margin-bottom:7px;}
 .pdin{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #B7BFC6;border-radius:6px;
      font-size:12.5px;font-family:inherit;resize:vertical;min-height:38px;margin-bottom:7px;}
 .pdrow{display:flex;gap:8px;align-items:center;flex-wrap:wrap;}
 .pdfile{font-size:11px;flex:1;min-width:150px;}
 .dsend{background:#00A09B;color:#fff;border:none;padding:8px 16px;border-radius:6px;
        font-size:12.5px;font-weight:600;cursor:pointer;}
 .dnote2{font-size:10.5px;color:#8595A5;margin-top:5px;}
 h2{font-size:16px;margin:0 0 12px;}
 .spin{display:none;color:#00706C;font-size:13px;margin-top:10px;font-weight:600;}
</style></head><body>
<style>
 .cols{max-width:1160px;margin:0 auto;padding:0 24px 40px;display:flex;gap:18px;align-items:flex-start;}
 .mainc{flex:1;min-width:0;}
 .fside{width:230px;flex-shrink:0;background:#fff;border:1px solid #D9DDE1;border-radius:12px;
        padding:13px;position:sticky;top:14px;max-height:calc(100vh - 28px);overflow-y:auto;}
 @media(max-width:900px){.cols{flex-direction:column;}.fside{width:auto;position:static;max-height:none;}}
 .fside b{font-size:12.5px;} .fitem{font-size:11.5px;margin:8px 0;line-height:1.5;word-break:break-word;}
 .fitem a{color:#00706C;font-weight:700;text-decoration:none;}
 .fitem small{color:#8595A5;display:block;}
 .fnote{font-size:10px;color:#8595A5;margin-top:8px;line-height:1.5;}
</style>
<div class="band"><h1>{{ client['name'] }} &rsaquo; {{ head_name }}</h1>
 <div><a href="{{ url_for('client_heads', cid=client['cid']) }}">&larr; Back to heads</a>
  <a href="{{ url_for('clients_page') }}">All clients</a>
  <a href="{{ url_for('logout') }}">Log out</a></div></div>
<div class="cols">
<div class="mainc">
 {% if error %}<div class="err">{{ error|safe }}</div>{% endif %}
 {% if okmsg %}<div class="okmsg">{{ okmsg }}</div>{% endif %}
 {% if engine_multi %}<div style="font-size:12px;color:#5B7083;margin-bottom:10px;"><b>AI engine: {{ engine_current }}</b> &mdash; <a style="color:#00A09B;" href="{{ url_for('set_engine', name=engine_other_key) }}">switch to {{ engine_other_label }}</a></div>{% endif %}

 {% if not is_cross %}
 <div class="card">
  <h2>Upload working papers or additional evidence &mdash; {{ head_name }}</h2>
  <form method="post" enctype="multipart/form-data" onsubmit="document.getElementById('spin').style.display='block'">
   <input type="file" name="files" multiple accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt">
   <textarea class="instr" name="instructions" maxlength="2000"
     placeholder="Optional instructions, e.g. This file answers point 3 - the missing conclusion has been added, please re-check"></textarea>
   <label style="display:block;margin-top:9px;font-size:12.5px;color:#3A4A64;cursor:pointer;">
     <input type="checkbox" name="hidden" checked> Review hidden Excel sheets (untick to skip them)</label>
   <button class="go" type="submit">Review in this head</button>
   <div class="spin" id="spin">Reviewing... this can take 1-3 minutes per file. Please keep the page open.</div>
  </form>
  <div class="hint">This head covers: {{ head_examples }}. Files belonging to a different head will be redirected to the correct tab. Upload follow-up evidence any time &mdash; the AI re-checks the open points below against it.</div>
 </div>
 {% else %}
 <div class="card"><div class="hint">These points were found by the client-wide cross-check across all files on record. Re-run it from the client page after new uploads.</div></div>
 {% endif %}

 {% if points %}
 <div class="card">
  <h2>Review points &mdash; {{ head_name }}</h2>
  <div class="sbar">
    <span class="sb" id="sb-t"></span><span class="sb p" id="sb-p"></span>
    <span class="sb r" id="sb-r"></span><span class="sb x" id="sb-x"></span>
    <span class="sb fl" id="sb-f"></span>
  </div>
  {% for p in points %}
   <div class="pt" data-pid="{{ p['id'] }}" data-status="{{ p.get('status','pending') }}"
        data-flag="{{ '1' if p.get('flagged') else '0' }}">
    <div class="ptop"><span class="ptitle">{{ p['id'] }}. {{ p.get('title','') }}</span>
      <span class="sev {{ p.get('severity','Low') }}">{{ p.get('severity','') }}</span></div>
    <div class="pexpl">{{ p.get('explanation','') }}</div>
    {% if p.get('reference') %}<div class="pref">{{ p['reference'] }}</div>{% endif %}
    {% if p.get('fix') %}<div class="pfix"><b>Suggested fix:</b> {{ p['fix'] }}</div>{% endif %}
    {% if p.get('ai_update') %}
      <div class="aiu {{ 'res' if p['ai_update'].get('resolution')=='resolved' else 'open' }}">
        <b>AI re-check ({{ p['ai_update'].get('time','') }}):</b>
        {{ 'Appears RESOLVED by the new evidence' if p['ai_update'].get('resolution')=='resolved' else 'Still OPEN' }}
        &mdash; {{ p['ai_update'].get('comment','') }}
        {% if p['ai_update'].get('resolution')=='resolved' %} (Confirm by clicking Resolved.){% endif %}
      </div>
    {% endif %}
    <div class="meta">Raised {{ p.get('time','') }} from {{ p.get('source','') }}</div>
    <div class="stat-row">
      <button type="button" class="stbtn pen" onclick="hstat(this,'pending')">Pending</button>
      <button type="button" class="stbtn res" onclick="hstat(this,'resolved')">&#10003; Resolved</button>
      <button type="button" class="stbtn rej" onclick="hstat(this,'rejected')">&#10007; Rejected</button>
      <button type="button" class="stbtn flg" onclick="hflag(this)">&#9873; Flag</button>
    </div>
    <button type="button" class="disc-btn" onclick="pdToggle(this)">&#128172; Discuss / add evidence</button>
    <div class="pdisc" hidden>
      <div class="pdlog">
        {% for m in p.get('discussion', []) %}<div class="dmsg {{ 'du' if m['role']=='user' else 'da' }}">{{ m['content'] }}</div>{% endfor %}
      </div>
      <textarea class="pdin" placeholder="Ask a question, object to this point, or explain what you're attaching..."></textarea>
      <div class="pdrow">
        <input type="file" class="pdfile" accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt">
        <button type="button" class="dsend" onclick="pdSend(this)">Send</button>
      </div>
      <div class="dnote2">Attach supporting documents any time &mdash; the AI answers knowing this point, its source file, and the FS anchor, and says whether the evidence resolves it. The discussion is saved with the client. Final judgement stays with the audit team.</div>
    </div>
   </div>
  {% endfor %}
 </div>
 {% endif %}

 {% if rounds %}
 <div class="card">
  <h2>Review rounds in this head</h2>
  {% for r in rounds|reverse %}
    <div class="round"><b>{{ r.get('time','') }}</b> &mdash; {{ r.get('files',[])|join(', ') }}
      {% if r.get('conclusion') %}<br>{{ r['conclusion'] }}{% endif %}</div>
  {% endfor %}
 </div>
 {% endif %}
</div>

<aside class="fside">
 <b>&#128194; Client files &mdash; open any time</b>
 {% if client.get('fs') %}
  <div class="fitem">&#128209; {{ client['fs']['name'] }}
    <small>FS anchor &middot; {{ client['fs'].get('time','') }}</small>
    {% if client['fs'].get('fid') %}<a href="{{ url_for('open_client_file', cid=client['cid'], fid=client['fs']['fid']) }}">Open</a>{% endif %}</div>
 {% endif %}
 {% for d in client.get('library', []) | reverse %}
  <div class="fitem">&#128196; {{ d['name'] }}
    <small>{{ head_names.get(d['head'], d['head']) if head_names else d['head'] }} &middot; {{ d.get('time','') }}</small>
    {% if d.get('fid') %}<a href="{{ url_for('open_client_file', cid=client['cid'], fid=d['fid']) }}">Open</a>{% else %}<small>(text on record; original not stored)</small>{% endif %}</div>
 {% endfor %}
 {% if not client.get('library') and not client.get('fs') %}
  <div class="fitem" style="color:#5B7083;">No files on record yet.</div>
 {% endif %}
 <div class="fnote">Originals up to 15 MB are saved with the client and open from here at any time. Cleared if the free server restarts &mdash; permanent storage arrives at Stage 5.</div>
</aside>
</div>
<script>
const CID = {{ client['cid'] | tojson }};
const HEAD = {{ head_key | tojson }};
function hcount(){
  const ps = document.querySelectorAll('.pt');
  let p=0,r=0,x=0,fl=0;
  ps.forEach(el=>{const s=el.dataset.status||'pending';
    if(s==='resolved')r++;else if(s==='rejected')x++;else p++;
    if(el.dataset.flag==='1')fl++;});
  const set=(id,t)=>{const e=document.getElementById(id);if(e)e.textContent=t;};
  set('sb-t','Total: '+ps.length); set('sb-p','Open/Pending: '+p);
  set('sb-r','Resolved/Closed: '+r); set('sb-x','Rejected: '+x); set('sb-f','Flagged: '+fl);
}
function hpost(el, action, ok){
  fetch('/hstatus',{method:'POST',headers:{'Content-Type':'application/json'},
    body:JSON.stringify({cid:CID,head:HEAD,pid:parseInt(el.dataset.pid),action:action})})
   .then(r=>r.json()).then(d=>{ if(d.ok){ok();hcount();} else alert(d.error||'Could not save.'); })
   .catch(()=>alert('Could not reach the server.'));
}
function hstat(btn,s){ const el=btn.closest('.pt'); hpost(el,s,()=>{el.dataset.status=s;}); }
function pdToggle(btn){ const d = btn.nextElementSibling; d.hidden = !d.hidden;
  if(!d.hidden){ const l=d.querySelector('.pdlog'); l.scrollTop=l.scrollHeight; d.querySelector('.pdin').focus(); } }
function pdMsg(log, cls, text){ const m=document.createElement('div'); m.className='dmsg '+cls;
  m.textContent=text; log.appendChild(m); log.scrollTop=log.scrollHeight; return m; }
function pdSend(btn){
  const box = btn.closest('.pdisc');
  const el = btn.closest('.pt');
  const log = box.querySelector('.pdlog');
  const inp = box.querySelector('.pdin');
  const fin = box.querySelector('.pdfile');
  const q = inp.value.trim();
  if(!q && !fin.files.length) return;
  btn.disabled = true;
  const fd = new FormData();
  fd.append('cid', CID); fd.append('head', HEAD);
  fd.append('pid', el.dataset.pid); fd.append('question', q);
  if(fin.files.length) fd.append('doc', fin.files[0]);
  const shown = q + (fin.files.length ? '  [attached: ' + fin.files[0].name + ']' : '');
  pdMsg(log, 'du', shown || '(attachment)');
  const wait = document.createElement('div'); wait.className='dwait';
  wait.textContent = 'The reviewer is examining...'; log.appendChild(wait); log.scrollTop=log.scrollHeight;
  inp.value = ''; fin.value = '';
  fetch('/hdiscuss', {method:'POST', body: fd})
   .then(r=>r.json()).then(d=>{
     wait.remove();
     pdMsg(log, 'da', d.answer || d.error || 'Something went wrong. Please try again.');
   }).catch(()=>{ wait.remove(); pdMsg(log,'da','Could not reach the server. Please try again.'); })
   .finally(()=>{ btn.disabled=false; });
}
function hflag(btn){ const el=btn.closest('.pt');
  const a=el.dataset.flag==='1'?'unflag':'flag';
  hpost(el,a,()=>{el.dataset.flag=el.dataset.flag==='1'?'0':'1';}); }
hcount();
</script>
</body></html>
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
 .anchorsel{width:100%;box-sizing:border-box;padding:10px 12px;border:1px solid #B7BFC6;border-radius:8px;font-size:13px;font-family:inherit;color:#002B49;background:#fff;}
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
 .disc-btn{margin-top:10px;background:#EAF6F6;color:#00706C;border:1px solid #BFE0DE;
        padding:7px 14px;border-radius:6px;font-size:12.5px;font-weight:600;cursor:pointer;}
 .disc-btn:hover{background:#DDF1F0;}
 .disc{margin-top:10px;border:1px solid #BFE0DE;border-radius:8px;background:#F7FBFB;padding:10px;}
 .dlog{max-height:340px;overflow-y:auto;margin-bottom:8px;}
 .dmsg{padding:8px 11px;border-radius:8px;margin-bottom:7px;font-size:12.5px;line-height:1.55;white-space:pre-wrap;}
 .du{background:#E4EFF9;color:#0A3556;margin-left:12%;}
 .da{background:#fff;border:1px solid #D9DDE1;color:#3A4A64;margin-right:12%;}
 .dwait{color:#5B7083;font-size:12px;font-style:italic;margin-bottom:7px;}
 .drow{display:flex;gap:8px;align-items:flex-end;}
 .din{flex:1;box-sizing:border-box;padding:8px 10px;border:1px solid #B7BFC6;border-radius:6px;
      font-size:12.5px;font-family:inherit;resize:vertical;min-height:38px;}
 .dsend{background:#00A09B;color:#fff;border:none;padding:9px 16px;border-radius:6px;
        font-size:12.5px;font-weight:600;cursor:pointer;}
 .dnote{font-size:10.5px;color:#8595A5;margin-top:5px;}
 .layout{max-width:1200px;margin:0 auto;padding:0 24px;display:flex;gap:20px;align-items:flex-start;}
 .maincol{flex:1;min-width:0;}
 .side{width:242px;flex-shrink:0;background:#fff;border:1px solid #D9DDE1;border-radius:12px;
       padding:14px;position:sticky;top:14px;max-height:calc(100vh - 28px);overflow-y:auto;}
 @media(max-width:900px){.layout{flex-direction:column;}.side{width:auto;position:static;max-height:none;}}
 .shead{font-weight:700;font-size:13px;margin-bottom:9px;color:#002B49;}
 .hitem{display:block;text-decoration:none;border:1px solid #E4E7EB;border-radius:8px;
        padding:8px 10px;margin-bottom:8px;color:#3A4A64;}
 .hitem:hover{border-color:#00A09B;background:#F4FAFA;}
 .hitem.cur{border-color:#00A09B;background:#EAF6F6;}
 .htime{font-size:10.5px;color:#5B7083;margin-bottom:3px;}
 .hfiles{font-size:11.5px;color:#0A3556;word-break:break-word;}
 .hcount{font-size:10.5px;color:#00706C;margin-top:3px;font-weight:600;}
 .hempty{font-size:12px;color:#5B7083;}
 .hnote{font-size:10px;color:#8595A5;margin-top:8px;line-height:1.5;}
 .sbar{display:flex;gap:8px;flex-wrap:wrap;margin:0 0 14px;}
 .sb{font-size:12px;font-weight:700;padding:6px 12px;border-radius:16px;background:#EFF2F4;color:#3A4A64;}
 .sb.p{background:#F3EAD3;color:#B0791C;} .sb.r{background:#E2F2E9;color:#1F6B4F;}
 .sb.x{background:#F5E1DE;color:#B23A2E;} .sb.fl{background:#E7E4F7;color:#5B4FC0;}
 .stat-row{display:flex;gap:6px;flex-wrap:wrap;margin-top:10px;align-items:center;}
 .stbtn{border:1px solid #D9DDE1;background:#fff;color:#5B7083;font-size:11.5px;
        font-weight:600;padding:5px 11px;border-radius:14px;cursor:pointer;}
 .stbtn:hover{border-color:#00A09B;}
 .finding[data-status="pending"] .stbtn.pen{background:#B0791C;color:#fff;border-color:#B0791C;}
 .finding[data-status="resolved"] .stbtn.res{background:#1F6B4F;color:#fff;border-color:#1F6B4F;}
 .finding[data-status="rejected"] .stbtn.rej{background:#B23A2E;color:#fff;border-color:#B23A2E;}
 .finding[data-flag="1"] .stbtn.flg{background:#5B4FC0;color:#fff;border-color:#5B4FC0;}
 .finding[data-status="resolved"] .ftitle{color:#1F6B4F;}
 .finding[data-status="rejected"] .fexpl,.finding[data-status="rejected"] .ftitle{opacity:.55;text-decoration:line-through;}
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
<div class="layout">
 <div class="maincol">
 <div class="sub">Baker Tilly - {{ 'Financial Statements review' if mode=='fs' else 'Working-paper review' }} - Stage 4
   &nbsp;|&nbsp; <a href="{{ url_for('choose') }}">Change review type</a>{% if engine_multi %} &nbsp;|&nbsp; <b>AI: {{ engine_current }}</b> &mdash; <a href="{{ url_for('set_engine', name=engine_other_key) }}">switch to {{ engine_other_label }}</a>{% endif %}</div>

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
   <div class="instr-label">Cross-check anchor (optional — needs 2+ files)</div>
   <select class="anchorsel" name="anchor" id="anchorsel">
     <option value="">No anchor — review each file on its own</option>
   </select>
   <div class="instr-hint">If you upload the signed financial statements together with working papers, choose the FS here — every other file is then reviewed AGAINST it: tie-outs, contradictions, impossible dates, omissions.</div>
   <label style="display:block;margin-top:10px;font-size:12.5px;color:#3A4A64;cursor:pointer;text-align:left;">
     <input type="checkbox" name="hidden" checked> Review hidden Excel sheets (untick to skip them)</label>
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
   <div class="sbar">
     <span class="sb" id="sb-t">Total: 0</span>
     <span class="sb p" id="sb-p">Pending: 0</span>
     <span class="sb r" id="sb-r">Resolved: 0</span>
     <span class="sb x" id="sb-x">Rejected: 0</span>
     <span class="sb fl" id="sb-f">Flagged: 0</span>
   </div>
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
     {% set fidx = loop.index0 %}
     <div class="filehead">FILE: {{ item['filename'] }}{% if item.get('is_anchor') %} <span style="background:#E7E4F7;color:#5B4FC0;font-size:10.5px;font-weight:700;padding:2px 8px;border-radius:10px;vertical-align:middle;">ANCHOR</span>{% endif %}
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
        <div class="finding" data-f="{{ fidx }}" data-g="{{ loop.index0 }}"
             data-status="{{ f.get('status','pending') }}" data-flag="{{ '1' if f.get('flagged') else '0' }}">
         <div class="bar {{ f.get('severity','Low') }}"></div>
         <div class="fbody">
          <div class="ftop"><span class="ftitle">{{ f.get('title','') }}</span>
            <span class="sev {{ f.get('severity','Low') }}">{{ f.get('severity','') }}</span></div>
          <div class="fexpl">{{ f.get('explanation','') }}</div>
          {% if f.get('reference') %}<div class="ref">{{ f['reference'] }}</div>{% endif %}
          <div class="fix"><b>Suggested fix:</b> {{ f.get('fix','') }}</div>
          <div class="stat-row">
            <button type="button" class="stbtn pen" onclick="setStat(this,'pending')">Pending</button>
            <button type="button" class="stbtn res" onclick="setStat(this,'resolved')">&#10003; Resolved</button>
            <button type="button" class="stbtn rej" onclick="setStat(this,'rejected')">&#10007; Rejected</button>
            <button type="button" class="stbtn flg" onclick="toggleFlag(this)">&#9873; Flag</button>
          </div>
          <button class="disc-btn" type="button" onclick="discToggle(this)">&#128172; Discuss with AI</button>
          <div class="disc" hidden data-f="{{ fidx }}" data-g="{{ loop.index0 }}">
            <div class="dlog"></div>
            <div class="drow">
              <textarea class="din" rows="2" placeholder="Ask a question, object, or challenge this point..."></textarea>
              <button type="button" class="dsend" onclick="discSend(this)">Send</button>
            </div>
            <div class="dnote">AI discussion is a draft aid — final judgement rests with the audit team. Discussions are not saved and end when the report expires.</div>
          </div>
         </div>
        </div>
       {% endfor %}
     {% endif %}
   {% endfor %}
   <div class="disclaimer"><b>Professional judgement statement:</b> {{ disclaimer }}</div>
  </div>
 {% endif %}
 </div>

 <aside class="side">
  <div class="shead">&#128337; Review history</div>
  {% if history %}
    {% for h in history %}
      <a class="hitem{{ ' cur' if batch_id and h['rid'] == batch_id else '' }}"
         href="{{ url_for('view_report', rid=h['rid']) }}">
        <div class="htime">{{ h['time'] }} &middot; {{ h['mode'] }} &middot; {{ h['user'] }}</div>
        <div class="hfiles">{{ h['files']|join(', ') }}</div>
        <div class="hcount">{{ h['findings'] }} finding{{ '' if h['findings']==1 else 's' }}</div>
      </a>
    {% endfor %}
  {% else %}
    <div class="hempty">No reviews yet in this server session.</div>
  {% endif %}
  <div class="hnote">History and statuses last until the free server restarts or redeploys — download Excel/PDF for a permanent record. Permanent storage arrives with the firm's own server (Stage 5).</div>
 </aside>
</div>

<script>
const drop = document.getElementById('drop');
const input = document.getElementById('fileinput');
const list = document.getElementById('filelist');
const MAXF = {{ maxfiles }};

// the basket: files accumulate across any mix of drags and browses
let picked = [];

function fillAnchor(){
  const sel = document.getElementById('anchorsel');
  if(!sel) return;
  const keep = sel.value;
  sel.innerHTML = '<option value="">No anchor — review each file on its own</option>';
  if(picked.length > 1){
    picked.forEach(f => { const o = document.createElement('option');
      o.value = f.name; o.textContent = 'Anchor: ' + f.name; sel.appendChild(o); });
    if([...sel.options].some(o => o.value === keep)) sel.value = keep;
  }
}

function syncInput(){
  const dt = new DataTransfer();
  picked.forEach(f => dt.items.add(f));
  input.files = dt.files;   // what actually gets submitted with the form
  renderList();
  fillAnchor();
}

function renderList(){
  if(picked.length === 0){ list.innerHTML = ''; return; }
  let html = '';
  picked.forEach((f, i) => {
    html += '<div>&#128196; ' + f.name.replace(/</g,'&lt;')
          + ' <a href="#" onclick="removeFile(' + i + ');return false;"'
          + ' style="color:#B23A2E;font-weight:700;text-decoration:none;margin-left:6px;"'
          + ' title="Remove this file">&#10005;</a></div>';
  });
  html += '<div style="margin-top:6px;"><b>' + picked.length + ' of ' + MAXF + ' files</b>'
        + (picked.length > 1
           ? ' &nbsp;<a href="#" onclick="clearFiles();return false;" style="color:#5B7083;">clear all</a>'
           : '')
        + '</div>';
  list.innerHTML = html;
}

function removeFile(i){ picked.splice(i, 1); syncInput(); }
function clearFiles(){ picked = []; syncInput(); }

function addFiles(files){
  let skippedDup = 0, hitCap = false;
  for(const f of files){
    if(picked.length >= MAXF){ hitCap = true; break; }
    if(picked.some(p => p.name === f.name && p.size === f.size)){ skippedDup++; continue; }
    picked.push(f);
  }
  syncInput();
  if(hitCap) alert('Maximum ' + MAXF + ' files per batch — extra files were not added.');
  else if(skippedDup) alert(skippedDup + ' file(s) skipped: already in the list.');
}

input.addEventListener('change', () => { addFiles(input.files); });
['dragover','dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
['dragleave','drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
drop.addEventListener('drop', e => {
  if(e.dataTransfer.files.length){ addFiles(e.dataTransfer.files); }
});

const RID = {{ (batch_id or "") | tojson }};

function recount(){
  const fs = document.querySelectorAll('.finding');
  let p=0, r=0, x=0, fl=0;
  fs.forEach(el => {
    const s = el.dataset.status || 'pending';
    if(s === 'resolved') r++; else if(s === 'rejected') x++; else p++;
    if(el.dataset.flag === '1') fl++;
  });
  const set = (id, txt) => { const e = document.getElementById(id); if(e) e.textContent = txt; };
  set('sb-t', 'Total: ' + fs.length);
  set('sb-p', 'Pending: ' + p);
  set('sb-r', 'Resolved: ' + r);
  set('sb-x', 'Rejected: ' + x);
  set('sb-f', 'Flagged: ' + fl);
}

function statPost(el, action, onOk){
  fetch('/status', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({rid: RID, file_index: parseInt(el.dataset.f),
                          finding_index: parseInt(el.dataset.g), action: action})
  }).then(r => r.json()).then(d => {
    if(d.ok){ onOk(); recount(); }
    else alert(d.error || 'Could not save the change.');
  }).catch(() => alert('Could not reach the server. Please try again.'));
}

function setStat(btn, s){
  const el = btn.closest('.finding');
  statPost(el, s, () => { el.dataset.status = s; });
}

function toggleFlag(btn){
  const el = btn.closest('.finding');
  const action = el.dataset.flag === '1' ? 'unflag' : 'flag';
  statPost(el, action, () => { el.dataset.flag = (el.dataset.flag === '1') ? '0' : '1'; });
}

recount();

function discToggle(btn){
  const d = btn.nextElementSibling;
  d.hidden = !d.hidden;
  if(!d.hidden){ d.querySelector('.din').focus(); }
}

function addMsg(log, cls, text){
  const m = document.createElement('div');
  m.className = 'dmsg ' + cls;
  m.textContent = text;
  log.appendChild(m);
  log.scrollTop = log.scrollHeight;
  return m;
}

function discSend(btn){
  const box = btn.closest('.disc');
  const log = box.querySelector('.dlog');
  const inp = box.querySelector('.din');
  const q = inp.value.trim();
  if(!q) return;
  if(!box._hist) box._hist = [];
  inp.value = '';
  btn.disabled = true;
  addMsg(log, 'du', q);
  const wait = document.createElement('div');
  wait.className = 'dwait';
  wait.textContent = 'The reviewer is thinking...';
  log.appendChild(wait); log.scrollTop = log.scrollHeight;
  fetch('/discuss', {
    method: 'POST',
    headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({
      rid: RID,
      file_index: parseInt(box.dataset.f),
      finding_index: parseInt(box.dataset.g),
      question: q,
      history: box._hist
    })
  }).then(r => r.json()).then(data => {
    wait.remove();
    if(data.answer){
      addMsg(log, 'da', data.answer);
      box._hist.push({role:'user', content:q});
      box._hist.push({role:'assistant', content:data.answer});
      if(box._hist.length > 16){ box._hist = box._hist.slice(-16); }
    } else {
      addMsg(log, 'da', data.error || 'Something went wrong. Please try again.');
    }
  }).catch(() => {
    wait.remove();
    addMsg(log, 'da', 'Could not reach the server. Please check your connection and try again.');
  }).finally(() => { btn.disabled = false; });
}
</script>
</body></html>
"""




for _tpl in ("CHOOSE_PAGE", "WP_CHOICE_PAGE", "CLIENTS_PAGE", "CLIENT_HEADS_PAGE",
             "HEAD_PAGE", "LOGIN_PAGE", "MAIN_PAGE"):
    if _tpl in globals():
        globals()[_tpl] = globals()[_tpl].replace("</head>", COMMON_UI + "</head>")
# dark ambient pages get the full-screen drifting dots
for _tpl in ("CHOOSE_PAGE", "WP_CHOICE_PAGE"):
    if _tpl in globals():
        globals()[_tpl] = globals()[_tpl].replace("<body>", "<body class=\"darkbg\">", 1)


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


@app.route("/engine/<name>")
@login_required
def set_engine(name):
    if name in ENGINES:
        session["engine"] = name
    return redirect(request.referrer or url_for("home"))


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
    if mode == "wp":
        return redirect(url_for("wp_choice"))
    return redirect(url_for("home"))


@app.route("/wp-choice")
@login_required
def wp_choice():
    session["mode"] = "wp"
    return render_template_string(WP_CHOICE_PAGE)


@app.route("/clients", methods=["GET", "POST"])
@login_required
def clients_page():
    error = None
    if request.method == "POST":
        name = (request.form.get("client_name") or "").strip()[:80]
        if not name:
            error = "Please enter a client name."
        else:
            clients = load_clients()
            if any(c["name"].lower() == name.lower() for c in clients):
                error = "A client with this name already exists."
            else:
                cid = uuid.uuid4().hex[:10]
                clients.insert(0, {"cid": cid, "name": name,
                                   "created": time.strftime("%d %b %Y")})
                save_clients(clients)
                save_client({"cid": cid, "name": name, "heads": {}})
                return redirect(url_for("client_heads", cid=cid))
    return render_template_string(CLIENTS_PAGE, clients=load_clients(),
                                  user=session.get("user"), error=error,
                                  permanent=PERMANENT_STORE)


@app.route("/clients/<cid>/rename", methods=["POST"])
@login_required
def client_rename(cid):
    new = (request.form.get("new_name") or "").strip()[:80]
    clients = load_clients()
    if new:
        for cl in clients:
            if cl["cid"] == cid:
                cl["name"] = new
                save_clients(clients)
                data = load_client(cid)
                if data:
                    data["name"] = new
                    save_client(data)
                break
    return redirect(url_for("clients_page"))


@app.route("/clients/<cid>/close", methods=["POST"])
@login_required
def client_close(cid):
    clients = load_clients()
    for cl in clients:
        if cl["cid"] == cid:
            cl["status"] = "active" if cl.get("status") == "closed" else "closed"
            save_clients(clients)
            break
    return redirect(url_for("clients_page"))


@app.route("/clients/<cid>/delete", methods=["POST"])
@login_required
def client_delete(cid):
    clients = [cl for cl in load_clients() if cl["cid"] != cid]
    save_clients(clients)
    delete_client_data(cid)
    return redirect(url_for("clients_page"))


@app.route("/clients/<cid>/export")
@login_required
def client_export(cid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    meta = next((cl for cl in load_clients() if cl["cid"] == cid), {})
    payload = {"kind": "bt-audit-client-backup", "version": 1,
               "meta": meta, "client": data,
               "note": ("Review points, statuses, rounds, extracted file text and "
                        "the FS anchor text are included. Original binary files "
                        "are not - keep those separately.")}
    buf = io.BytesIO(json.dumps(payload, ensure_ascii=False).encode("utf-8"))
    safe = "".join(ch if ch.isalnum() or ch in " -_" else "_" for ch in data.get("name", "client"))
    return send_file(buf, as_attachment=True, mimetype="application/json",
                     download_name=safe.strip().replace(" ", "_") + "_backup.json")


@app.route("/clients/import", methods=["POST"])
@login_required
def client_import():
    up = request.files.get("backup")
    error = None
    if not up or not up.filename:
        error = "Please choose a backup file."
    else:
        try:
            payload = json.loads(up.read().decode("utf-8"))
            assert payload.get("kind") == "bt-audit-client-backup"
            data = payload["client"]
            assert isinstance(data.get("name"), str) and isinstance(data.get("heads", {}), dict)
            clients = load_clients()
            cid = data.get("cid", "")
            if not cid or any(cl["cid"] == cid for cl in clients):
                cid = uuid.uuid4().hex[:10]
            data["cid"] = cid
            meta = payload.get("meta") or {}
            clients.insert(0, {"cid": cid, "name": data["name"],
                               "created": meta.get("created", time.strftime("%d %b %Y")),
                               "status": meta.get("status", "active")})
            save_clients(clients)
            save_client(data)
            return redirect(url_for("client_heads", cid=cid))
        except Exception:
            error = ("That file is not a valid client backup. Use a file downloaded "
                     "with a client's Backup button.")
    return render_template_string(CLIENTS_PAGE, clients=load_clients(),
                                  user=session.get("user"), error=error,
                                  permanent=PERMANENT_STORE)


@app.route("/client/<cid>")
@login_required
def client_heads(cid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                  sections=SECTIONS, error=None, okmsg=None,
                                  head_names=HEAD_NAMES)


@app.route("/client/<cid>/<head>", methods=["GET", "POST"])
@login_required
def head_page(cid, head):
    data = load_client(cid)
    is_cross = head in ("cross", "fsr")   # button-driven sections: no upload form
    if not data or (head not in HEAD_NAMES and not is_cross):
        return redirect(url_for("clients_page"))
    if head == "cross":
        head_name = "Cross-head checks"
        head_examples = "inconsistencies between heads: debtors vs revenue, advances vs sales, depreciation vs assets, finance cost vs borrowings"
    elif head == "fsr":
        head_name = "Financial statements review"
        head_examples = "the anchored FS reviewed on their own: presentation, tie-outs between face and notes, disclosures, arithmetic"
    else:
        head_name = HEAD_NAMES[head]
        head_examples = next(e for k, n, e in ALL_AREAS if k == head)
    hd = data.setdefault("heads", {}).setdefault(head, {"points": [], "rounds": []})
    error = None
    okmsg = None

    if request.method == "POST" and is_cross:
        return redirect(url_for("client_heads", cid=cid))
    if request.method == "POST":
        if not AI_KEY_SET:
            error = "The AI API key is not set in Render's Environment Variables."
        else:
            uploads = [f for f in request.files.getlist("files") if f and f.filename]
            if not uploads:
                error = "Please choose at least one file."
            else:
                uploads = uploads[:MAX_FILES_PER_BATCH]
                instructions = request.form.get("instructions", "")
                include_hidden = request.form.get("hidden") == "on"
                wrongs, done, new_pts, resolved_n = [], [], 0, 0
                import gc
                for up in uploads:
                    try:
                        raw = up.read()
                        stored_fid, stored_ext = store_client_file(cid, up.filename, raw)
                        text = extract_text_from_file(up.filename, raw,
                                include_hidden=include_hidden)
                        del raw
                        if text is None:
                            discard_client_file(cid, stored_fid, stored_ext)
                            wrongs.append(up.filename + ": unsupported file type.")
                            continue
                        if not text.strip():
                            discard_client_file(cid, stored_fid, stored_ext)
                            wrongs.append(up.filename + ": no readable text found.")
                            continue
                        related = []
                        budget = 3
                        for d in reversed(data.get("library", [])):
                            if budget == 0:
                                break
                            if d["head"] in RELATED_HEADS.get(head, []):
                                related.append({"name": d["name"],
                                                "head_name": HEAD_NAMES.get(d["head"], d["head"]),
                                                "excerpt": d["excerpt"][:8000]})
                                budget -= 1
                        fs = data.get("fs") or {}
                        result, ai_err = head_review_with_ai(
                            text, head, hd["points"], instructions,
                            related_docs=related,
                            fs_name=fs.get("name", ""),
                            fs_text=fs.get("excerpt", ""))
                        keep_excerpt = text[:12000]
                        del text
                        if ai_err:
                            discard_client_file(cid, stored_fid, stored_ext)
                            wrongs.append(up.filename + ": " + ai_err)
                            continue
                        if result.get("wrong_head"):
                            other = result["wrong_head"]
                            link = ""
                            for k2, n2, _e2 in ALL_AREAS:
                                if n2.lower() == str(other).strip().lower():
                                    link = url_for("head_page", cid=cid, head=k2)
                                    break
                            msg = ("<b>" + up.filename + "</b>: this tab is for <b>"
                                   + head_name + "</b> review. The file appears to "
                                   "belong to <b>" + str(other) + "</b>")
                            msg += (" &mdash; <a href='" + link + "'>go to that tab</a>."
                                    if link else ". Please use the relevant tab.")
                            discard_client_file(cid, stored_fid, stored_ext)
                            wrongs.append(msg)
                            continue
                        now = time.strftime("%d %b %Y, %H:%M")
                        for upd in result.get("point_updates", []):
                            try:
                                pid = int(upd.get("id"))
                            except Exception:
                                continue
                            for p in hd["points"]:
                                if p["id"] == pid:
                                    p["ai_update"] = {
                                        "resolution": ("resolved" if
                                            str(upd.get("resolution", "")).startswith("resolv")
                                            else "still_open"),
                                        "comment": str(upd.get("comment", ""))[:500],
                                        "time": now}
                                    if p["ai_update"]["resolution"] == "resolved":
                                        resolved_n += 1
                        next_id = max([p["id"] for p in hd["points"]], default=0) + 1
                        for f in result.get("findings", []):
                            hd["points"].append({
                                "id": next_id, "title": str(f.get("title", ""))[:200],
                                "explanation": str(f.get("explanation", ""))[:1500],
                                "reference": str(f.get("reference", ""))[:300],
                                "severity": f.get("severity", "Low"),
                                "fix": str(f.get("fix", ""))[:800],
                                "status": "pending", "flagged": False,
                                "time": now, "source": up.filename})
                            next_id += 1
                            new_pts += 1
                        hd["rounds"].append({
                            "time": now, "files": [up.filename],
                            "conclusion": str(result.get("conclusion", ""))[:600]})
                        lib = data.setdefault("library", [])
                        lib[:] = [d for d in lib
                                  if not (d["name"] == up.filename and d["head"] == head)]
                        lib.append({"name": up.filename, "head": head,
                                    "excerpt": keep_excerpt, "time": now,
                                    "fid": stored_fid, "ext": stored_ext})
                        del lib[:-24]
                        done.append(up.filename)
                    except Exception as e:
                        wrongs.append(up.filename + ": could not process ("
                                      + str(e)[:120] + ")")
                    gc.collect()
                save_client(data)
                if done:
                    okmsg = ("Reviewed: " + ", ".join(done) + " — "
                             + str(new_pts) + " new point(s)"
                             + (", " + str(resolved_n) +
                                " prior point(s) appear resolved (see AI re-checks below)"
                                if resolved_n else "") + ".")
                if wrongs:
                    error = "<br>".join(wrongs)

    return render_template_string(HEAD_PAGE, client=data, head_key=head,
                                  head_name=head_name, head_examples=head_examples,
                                  points=hd["points"], rounds=hd["rounds"],
                                  error=error, okmsg=okmsg, is_cross=is_cross,
                                  head_names=HEAD_NAMES)


@app.route("/client/<cid>/fs", methods=["POST"])
@login_required
def client_fs(cid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    up = request.files.get("fsfile")
    error = None
    okmsg = None
    if not up or not up.filename:
        error = "Please choose the financial statements file."
    else:
        try:
            raw = up.read()
            text = extract_text_from_file(up.filename, raw)
            del raw
            if text is None:
                error = "Unsupported file type for the financial statements."
            elif not text.strip():
                error = ("No readable text found in that file (a scanned PDF "
                         "without a text layer, perhaps).")
            else:
                is_fs, looks_like, gerr = fs_gate_with_ai(text)
                if gerr:
                    error = gerr
                elif not is_fs:
                    error = ("This upload must be the client's FINANCIAL "
                             "STATEMENTS (statement of financial position, "
                             "profit or loss, cash flows, equity, notes). "
                             "\"" + up.filename + "\" appears to be: "
                             + (looks_like or "a different kind of document")
                             + ". Working papers belong in their head tabs - "
                             "please upload the financial statements here.")
                else:
                    up.seek(0)
                    fs_fid, fs_ext = store_client_file(cid, up.filename, up.read())
                    data["fs"] = {"name": up.filename,
                                  "excerpt": text[:ANCHOR_CHARS],
                                  "time": time.strftime("%d %b %Y, %H:%M"),
                                  "fid": fs_fid, "ext": fs_ext}
                    save_client(data)
                    okmsg = ("Financial statements saved as the engagement anchor. "
                             "Every working paper reviewed in any head will now also "
                             "be checked against them automatically.")
                del text
        except Exception as e:
            error = "Could not read the file: " + str(e)[:120]
    return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                  sections=SECTIONS, error=error, okmsg=okmsg,
                                  head_names=HEAD_NAMES)


@app.route("/client/<cid>/fsreview", methods=["POST"])
@login_required
def client_fsreview(cid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    if not AI_KEY_SET:
        return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                      sections=SECTIONS, error="The AI API key is not set.",
                                      okmsg=None, head_names=HEAD_NAMES)
    fs = data.get("fs") or {}
    if not fs.get("excerpt"):
        return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                      sections=SECTIONS, okmsg=None, head_names=HEAD_NAMES,
                                      error="Upload the financial statements (anchor) first — the FS review reviews that file.")
    result, err = review_with_ai(fs["excerpt"], mode="fs")
    if err:
        return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                      sections=SECTIONS, error=err, okmsg=None,
                                      head_names=HEAD_NAMES)
    hd = data.setdefault("heads", {}).setdefault("fsr", {"points": [], "rounds": []})
    now = time.strftime("%d %b %Y, %H:%M")
    next_id = max([p["id"] for p in hd["points"]], default=0) + 1
    for f in result.get("findings", []):
        hd["points"].append({
            "id": next_id, "title": str(f.get("title", ""))[:200],
            "explanation": str(f.get("explanation", ""))[:1500],
            "reference": str(f.get("reference", ""))[:300],
            "severity": f.get("severity", "Medium"),
            "fix": str(f.get("fix", ""))[:800],
            "status": "pending", "flagged": False,
            "time": now, "source": fs.get("name", "financial statements")})
        next_id += 1
    hd["rounds"].append({"time": now, "files": [fs.get("name", "")],
                         "conclusion": str(result.get("conclusion", ""))[:600]})
    save_client(data)
    return redirect(url_for("head_page", cid=cid, head="fsr"))


@app.route("/client/<cid>/crosscheck", methods=["POST"])
@login_required
def client_crosscheck(cid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    if not AI_KEY_SET:
        return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                      sections=SECTIONS,
                                      error="The AI API key is not set.",
                                      okmsg=None, head_names=HEAD_NAMES)
    result, err = client_cross_check_with_ai(data)
    if err:
        return render_template_string(CLIENT_HEADS_PAGE, client=data, heads=HEADS,
                                      sections=SECTIONS, error=err, okmsg=None,
                                      head_names=HEAD_NAMES)
    hd = data.setdefault("heads", {}).setdefault("cross", {"points": [], "rounds": []})
    now = time.strftime("%d %b %Y, %H:%M")
    next_id = max([p["id"] for p in hd["points"]], default=0) + 1
    for f in result.get("findings", []):
        hd["points"].append({
            "id": next_id, "title": str(f.get("title", ""))[:200],
            "explanation": str(f.get("explanation", ""))[:1500],
            "reference": str(f.get("reference", ""))[:300],
            "severity": f.get("severity", "Medium"),
            "fix": str(f.get("fix", ""))[:800],
            "status": "pending", "flagged": False,
            "time": now, "source": "client-wide cross-check"})
        next_id += 1
    hd["rounds"].append({"time": now,
                         "files": [d["name"] for d in data.get("library", [])],
                         "conclusion": str(result.get("conclusion", ""))[:600]})
    save_client(data)
    return redirect(url_for("head_page", cid=cid, head="cross"))


@app.route("/client/<cid>/file/<fid>")
@login_required
def open_client_file(cid, fid):
    data = load_client(cid)
    if not data:
        return redirect(url_for("clients_page"))
    fid = "".join(c for c in fid if c.isalnum())
    entry = None
    fs = data.get("fs") or {}
    if fs.get("fid") == fid:
        entry = {"name": fs.get("name", "file"), "ext": fs.get("ext", "")}
    else:
        for d in data.get("library", []):
            if d.get("fid") == fid:
                entry = {"name": d.get("name", "file"), "ext": d.get("ext", "")}
                break
    if not entry:
        return "File not on record (it may predate file-saving, or the free server restarted).", 404
    blob = get_client_file(cid, fid, entry["ext"])
    if blob is None:
        return "The stored copy is no longer available.", 404
    return send_file(io.BytesIO(blob), as_attachment=True,
                     download_name=entry["name"])


@app.route("/hdiscuss", methods=["POST"])
@login_required
def hdiscuss():
    if not AI_KEY_SET:
        return {"error": "The AI API key is not set."}, 500
    cid = str(request.form.get("cid", ""))
    head = str(request.form.get("head", ""))
    data = load_client(cid)
    if not data:
        return {"error": "Client workspace not found."}, 404
    try:
        pid = int(request.form.get("pid", -1))
    except Exception:
        return {"error": "Bad point id."}, 400
    pts = data.get("heads", {}).get(head, {}).get("points", [])
    point = next((p for p in pts if p["id"] == pid), None)
    if point is None:
        return {"error": "That review point could not be found."}, 404

    question = (request.form.get("question") or "").strip()[:2000]
    doc = request.files.get("doc")
    doc_text, doc_name = "", ""
    if doc and doc.filename:
        try:
            raw = doc.read()
            t = extract_text_from_file(doc.filename, raw)
            del raw
            if t is None:
                return {"error": "Unsupported attachment type."}, 400
            if not t.strip():
                return {"error": "No readable text found in the attachment "
                                 "(a scanned PDF without a text layer, perhaps)."}, 400
            doc_text = t[:15000]
            doc_name = doc.filename
        except Exception as e:
            return {"error": "Could not read the attachment: " + str(e)[:100]}, 400
    if not question and not doc_text:
        return {"error": "Type a question or attach a document (or both)."}, 400
    if not question:
        question = "Please assess the attached document against this review point."

    context = ("CLIENT: " + data.get("name", "") + "\n"
               "AREA: " + HEAD_NAMES.get(head, head) + "\n\n"
               "THE REVIEW POINT UNDER DISCUSSION (id " + str(pid) + ", current "
               "status: " + point.get("status", "pending") + "):\n"
               "Title: " + str(point.get("title", "")) + "\n"
               "Severity: " + str(point.get("severity", "")) + "\n"
               "Explanation: " + str(point.get("explanation", "")) + "\n"
               "Reference: " + str(point.get("reference", "")) + "\n"
               "Suggested fix: " + str(point.get("fix", "")) + "\n"
               "Raised from file: " + str(point.get("source", "")))
    src_excerpt = ""
    for d in data.get("library", []):
        if d.get("name") == point.get("source"):
            src_excerpt = d.get("excerpt", "")[:6000]
            break
    fs_excerpt = (data.get("fs") or {}).get("excerpt", "")[:6000]

    messages = [
        {"role": "system", "content": DISCUSS_INSTRUCTIONS},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY:\n\n"
            + select_knowledge(context + " " + question + " " + doc_text[:8000])},
        {"role": "system", "content": context},
    ]
    if src_excerpt:
        messages.append({"role": "system", "content":
            "EXCERPT OF THE SOURCE FILE the point was raised from:\n" + src_excerpt})
    if fs_excerpt:
        messages.append({"role": "system", "content":
            "EXCERPT OF THE CLIENT FINANCIAL STATEMENTS (engagement anchor) for "
            "tie-outs and dates:\n" + fs_excerpt})
    for m in point.get("discussion", [])[-8:]:
        if m.get("role") in ("user", "assistant") and m.get("content"):
            messages.append({"role": m["role"], "content": str(m["content"])[:2000]})
    user_msg = question
    if doc_text:
        user_msg += ("\n\n[ATTACHED DOCUMENT \"" + doc_name + "\" — extracted "
                     "content:]\n" + doc_text +
                     "\n\nAssess honestly whether this attachment resolves the "
                     "point, partly resolves it, or leaves it open — and say "
                     "exactly what (if anything) is still missing.")
    messages.append({"role": "user", "content": user_msg})

    try:
        response = ai_chat(messages, max_tokens=1200, temperature=0.2)
        answer = response.choices[0].message.content.strip()
    except Exception as e:
        err_name = type(e).__name__
        if err_name == "EmptyAIResponse":
            return None, ("The AI returned an empty answer for this request "
                          "(a known DeepSeek issue). Please press the review "
                          "button again — or switch to the other engine.")
        if "Timeout" in err_name or "timeout" in str(e).lower():
            return {"error": "The AI took too long. Please try again."}, 504
        return {"error": "The AI could not be reached (" + err_name + ")."}, 502

    disc = point.setdefault("discussion", [])
    shown_q = question + (("  [attached: " + doc_name + "]") if doc_name else "")
    disc.append({"role": "user", "content": shown_q})
    disc.append({"role": "assistant", "content": answer})
    del disc[:-20]
    save_client(data)
    return {"answer": answer, "shown_q": shown_q}


@app.route("/hstatus", methods=["POST"])
@login_required
def hstatus():
    j = request.get_json(silent=True) or {}
    data = load_client(str(j.get("cid", "")))
    if not data:
        return {"error": "Client workspace not found (the free server may have "
                         "restarted)."}, 404
    head = str(j.get("head", ""))
    pts = data.get("heads", {}).get(head, {}).get("points", [])
    try:
        pid = int(j.get("pid", -1))
    except Exception:
        return {"error": "Bad point id."}, 400
    action = str(j.get("action", ""))
    for p in pts:
        if p["id"] == pid:
            if action in ("pending", "resolved", "rejected"):
                p["status"] = action
            elif action == "flag":
                p["flagged"] = True
            elif action == "unflag":
                p["flagged"] = False
            else:
                return {"error": "Unknown action."}, 400
            save_client(data)
            return {"ok": True}
    return {"error": "Point not found."}, 404


@app.route("/", methods=["GET", "POST"])
@login_required
def home():
    if session.get("mode") not in ("fs", "wp"):
        return redirect(url_for("choose"))
    error = None
    batch = None
    batch_id = None

    if request.method == "POST":
        if not AI_KEY_SET:
            error = "The AI API key is not set. Add it in Render's Environment Variables."
        else:
            uploads = [f for f in request.files.getlist("files") if f and f.filename]
            if not uploads:
                error = "Please choose at least one file."
            else:
                uploads = uploads[:MAX_FILES_PER_BATCH]
                batch = {"files": []}
                import gc
                # anchor pre-pass: if the user marked one file (e.g. the signed FS)
                # as the anchor, extract it first so every other file is reviewed
                # against it (tie-outs, contradictions, impossible dates, omissions)
                include_hidden = request.form.get("hidden") == "on"
                anchor_name = (request.form.get("anchor") or "").strip()
                anchor_text = ""
                if anchor_name and len(uploads) > 1:
                    for up in uploads:
                        if up.filename == anchor_name:
                            try:
                                d = up.read()
                                t = extract_text_from_file(up.filename, d,
                                        include_hidden=include_hidden) or ""
                                anchor_text = t[:ANCHOR_CHARS]
                                del d, t
                            except Exception:
                                anchor_text = ""
                            up.seek(0)
                            break
                for up in uploads:
                    entry = {"filename": up.filename}
                    if anchor_text and up.filename == anchor_name:
                        entry["is_anchor"] = True
                    try:
                        data = up.read()
                        text = extract_text_from_file(up.filename, data,
                                include_hidden=include_hidden)
                        del data  # release the raw file bytes immediately
                        if text is None:
                            entry["error"] = "Unsupported file type."
                        elif not text.strip():
                            entry["error"] = ("The file appears to be empty, or its text "
                                              "could not be read (a scanned PDF with no "
                                              "text layer, perhaps).")
                        else:
                            entry["excerpt"] = text[:3000]
                            use_anchor = bool(anchor_text) and up.filename != anchor_name
                            result, ai_err = review_with_ai(
                                text, mode=session.get("mode", "wp"),
                                user_instructions=request.form.get("instructions", ""),
                                anchor_name=(anchor_name if use_anchor else ""),
                                anchor_text=(anchor_text if use_anchor else ""))
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
                # keep a trimmed excerpt in the saved results so the per-finding
                # "Discuss with AI" feature has the file content as context
                for item in batch["files"]:
                    if item.get("excerpt"):
                        item["excerpt"] = item["excerpt"][:3000]
                batch_id = save_results(batch, user=session.get("user", ""),
                                        mode=session.get("mode", "wp"))

    return render_template_string(MAIN_PAGE, user=session.get("user"),
                                  role=session.get("role"), error=error,
                                  batch=batch, batch_id=batch_id,
                                  maxfiles=MAX_FILES_PER_BATCH,
                                  mode=session.get("mode", "wp"),
                                  history=load_history(),
                                  disclaimer=DISCLAIMER)


@app.route("/report/<rid>")
@login_required
def view_report(rid):
    batch = load_results(rid)
    error = None
    if not batch:
        error = ("This saved report is no longer available. On the free hosting, "
                 "history is cleared whenever the server restarts or redeploys — "
                 "the Excel/PDF downloads are the permanent record.")
    return render_template_string(MAIN_PAGE, user=session.get("user"),
                                  role=session.get("role"), error=error,
                                  batch=batch, batch_id=(rid if batch else None),
                                  maxfiles=MAX_FILES_PER_BATCH,
                                  mode=session.get("mode", "wp"),
                                  history=load_history(),
                                  disclaimer=DISCLAIMER)


@app.route("/status", methods=["POST"])
@login_required
def set_status():
    data = request.get_json(silent=True) or {}
    rid = str(data.get("rid", ""))
    action = str(data.get("action", ""))
    batch = load_results(rid)
    if not batch:
        return {"error": "This report has expired — statuses can no longer be "
                         "saved for it. Please run the review again."}, 404
    try:
        finding = (batch["files"][int(data.get("file_index", -1))]
                   ["result"]["findings"][int(data.get("finding_index", -1))])
    except Exception:
        return {"error": "That finding could not be found."}, 404
    if action in ("pending", "resolved", "rejected"):
        finding["status"] = action
    elif action == "flag":
        finding["flagged"] = True
    elif action == "unflag":
        finding["flagged"] = False
    else:
        return {"error": "Unknown action."}, 400
    update_results(rid, batch)
    return {"ok": True}


def fs_gate_with_ai(text):
    """Check that an upload offered as the FS anchor really is financial
    statements. Returns (is_fs, looks_like, err)."""
    sample = text[:6000]
    messages = [
        {"role": "system", "content":
            "You are a strict document classifier at an audit firm. Decide whether "
            "the document is FINANCIAL STATEMENTS: a statement of financial "
            "position / balance sheet, statement of profit or loss / income "
            "statement, statement of cash flows, statement of changes in equity, "
            "and/or the notes to the financial statements - complete or draft, "
            "full set or a substantial extract. The following are NOT financial "
            "statements: audit working papers, lead schedules, vouching or "
            "verification sheets, trial balances, ledgers, planning memoranda, "
            "checklists, engagement letters, bank statements, invoices, "
            "correspondence, or any other document. "
            "Return ONLY JSON: {\"is_fs\": true or false, "
            "\"looks_like\": \"2-6 word description of what the document "
            "actually appears to be\"}."},
        {"role": "user", "content": "Document extract:\n\n" + sample},
    ]
    try:
        response = ai_chat(messages, max_tokens=200, temperature=0, cheap=True)
        result, perr = parse_ai_json(response.choices[0].message.content.strip())
        if perr or not isinstance(result, dict) or "is_fs" not in result:
            return None, "", "The check could not be completed. Please try again."
        return bool(result.get("is_fs")), str(result.get("looks_like", ""))[:120], None
    except Exception as e:
        err_name = type(e).__name__
        return None, "", ("The AI check could not run (" + err_name
                          + "). Please try again in a moment.")


def client_cross_check_with_ai(data):
    """One pass across everything on record for a client: hunt inter-head
    inconsistencies (figures, parties, dates, treatments)."""
    lib = list(data.get("library", []))[:12]
    fs = data.get("fs") or {}
    if fs.get("excerpt"):
        lib = [{"name": fs.get("name", "financial statements"),
                "head": "fs", "excerpt": fs["excerpt"]}] + lib
    if len(lib) < 2:
        return None, ("At least two files (in different heads) must be on record "
                      "before a client-wide cross-check can run.")
    parts = []
    for d in lib:
        label = ("FINANCIAL STATEMENTS (MASTER ANCHOR)" if d["head"] == "fs"
                 else "HEAD \"" + HEAD_NAMES.get(d["head"], d["head"]) + "\"")
        parts.append("--- " + label + ", FILE \"" + d["name"]
                     + "\" (excerpt) ---\n" + d["excerpt"][:6000])
    open_pts = []
    for k, hd in data.get("heads", {}).items():
        for p in hd.get("points", []):
            if p.get("status", "pending") == "pending":
                open_pts.append(HEAD_NAMES.get(k, k) + " #" + str(p["id"]) + ": "
                                + p.get("title", ""))
    messages = [
        {"role": "system", "content":
            "You are an experienced audit reviewer performing a CLIENT-WIDE "
            "CROSS-CHECK across the working papers of different audit heads of one "
            "client. You are given excerpts from documents filed under different "
            "heads. Your ONLY job is to find issues BETWEEN heads: interlinked "
            "figures that do not tie (debtors vs revenue, advances vs sales, "
            "depreciation vs asset schedules, finance cost vs borrowings, profit "
            "vs equity movement), the same party or transaction treated "
            "inconsistently in different heads, dates that conflict, and items "
            "present in one head but unexplainably missing where they should also "
            "appear. Do NOT repeat single-document issues. Quote both figures and "
            "name both files in every finding. Report each issue once, most "
            "important first, maximum 15 findings, plain English. "
            "Return JSON: {\"findings\": [{\"title\", \"explanation\", "
            "\"reference\", \"severity\": \"High|Medium|Low|Factual\", "
            "\"fix\"}], \"summary\", \"conclusion\"}. Return ONLY the JSON."},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY:\n\n"
            + select_knowledge(" ".join(p[:4000] for p in parts))},
        {"role": "user", "content":
            "DOCUMENTS ON RECORD:\n\n" + "\n\n".join(parts)
            + ("\n\nOPEN REVIEW POINTS ACROSS HEADS (context, do not repeat):\n"
               + "\n".join(open_pts[:40]) if open_pts else "")},
    ]
    try:
        response = ai_chat(messages, max_tokens=5000, temperature=0.2)
    except Exception as e:
        err_name = type(e).__name__
        if err_name == "EmptyAIResponse":
            return None, ("The AI returned an empty answer for this request "
                          "(a known DeepSeek issue). Please press the review "
                          "button again — or switch to the other engine.")
        if "Timeout" in err_name or "timeout" in str(e).lower():
            return None, "The AI took too long. Please try again."
        return None, "The AI could not be reached. Details: " + err_name
    return parse_ai_json(response.choices[0].message.content.strip())


DISCUSS_INSTRUCTIONS = """You are an experienced audit reviewer at an accounting firm, in a follow-up discussion about ONE specific review finding that was raised on a file.

The user may: question whether the finding is correct, object to it, ask you to explain it more deeply or more simply, ask what evidence or fix is needed, or ask how the standards apply.

RULES:
- Be honest and objective. If the user's objection is valid or the original finding looks wrong or doubtful given the file content, SAY SO plainly and explain why — never defend a finding just because it was raised. It is normal for some findings to be revised or withdrawn on discussion.
- If the finding still stands, explain clearly why, using the file content and the firm's standards library.
- Cite standards only from the provided firm standards library; if the point falls outside it, say "outside loaded library — reference to be confirmed".
- Never invent facts, figures, or paragraph numbers. If the provided file excerpt does not show enough to be sure, say what additional evidence would settle it.
- Plain, easy English. Be concise: a few short paragraphs at most.
- You are a draft reviewer only — final professional judgement rests with the audit team. Do not claim authority to conclude.
Respond with plain text only (no JSON, no markdown headings)."""


@app.route("/discuss", methods=["POST"])
@login_required
def discuss():
    if not AI_KEY_SET:
        return {"error": "The AI API key is not set."}, 500
    data = request.get_json(silent=True) or {}
    rid = str(data.get("rid", ""))
    question = str(data.get("question", "")).strip()[:2000]
    history = data.get("history", [])
    if not question:
        return {"error": "Please type a question."}, 400
    batch = load_results(rid)
    if not batch:
        return {"error": "This report has expired (reports are kept temporarily). "
                         "Please run the review again, then discuss."}, 404
    try:
        fidx = int(data.get("file_index", -1))
        gidx = int(data.get("finding_index", -1))
        item = batch["files"][fidx]
        finding = item["result"]["findings"][gidx]
    except Exception:
        return {"error": "That finding could not be found in the saved report."}, 404

    context = ("FILE NAME: " + item.get("filename", "") + "\n\n"
               "THE FINDING UNDER DISCUSSION:\n"
               "Title: " + str(finding.get("title", "")) + "\n"
               "Severity: " + str(finding.get("severity", "")) + "\n"
               "Explanation: " + str(finding.get("explanation", "")) + "\n"
               "Reference: " + str(finding.get("reference", "")) + "\n"
               "Suggested fix: " + str(finding.get("fix", "")) + "\n\n"
               "EXCERPT OF THE FILE CONTENT (may be partial):\n"
               + (item.get("excerpt") or "(no excerpt retained)"))

    messages = [
        {"role": "system", "content": DISCUSS_INSTRUCTIONS},
        {"role": "system", "content": "FIRM'S STANDARDS LIBRARY:\n\n"
            + select_knowledge(context + " " + question)},
        {"role": "system", "content": context},
    ]
    # replay up to the last 8 turns of this discussion so the AI has the thread
    if isinstance(history, list):
        for h in history[-8:]:
            r = h.get("role")
            c = str(h.get("content", ""))[:2000]
            if r in ("user", "assistant") and c:
                messages.append({"role": r, "content": c})
    messages.append({"role": "user", "content": question})

    try:
        response = ai_chat(messages, max_tokens=1200, temperature=0.2)
        answer = response.choices[0].message.content.strip()
    except Exception as e:
        err_name = type(e).__name__
        if err_name == "EmptyAIResponse":
            return None, ("The AI returned an empty answer for this request "
                          "(a known DeepSeek issue). Please press the review "
                          "button again — or switch to the other engine.")
        if "Timeout" in err_name or "timeout" in str(e).lower():
            return {"error": "The AI took too long to answer. Please try again."}, 504
        return {"error": "The AI could not be reached. Please try again. "
                         "(" + err_name + ")"}, 502
    return {"answer": answer}


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
