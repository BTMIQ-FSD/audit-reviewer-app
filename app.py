"""
Baker Tilly AI Audit Reviewer — Stage 1 (Real Reviewer)
Upload a working paper (Excel / PDF / Word) -> the tool reads it ->
sends it to DeepSeek with proper audit-reviewer instructions ->
returns review points in plain English, each with a reference and a suggested fix,
ending with a professional-judgement statement.

NOTE: For now the AI reviews using its own knowledge of the standards.
Stage 2 will add the firm's real knowledge base so every reference is exact.
"""

import os
import io
import json
from flask import Flask, request, render_template_string
from openai import OpenAI

# Libraries to read the different file types
from openpyxl import load_workbook          # Excel
from docx import Document as DocxDocument    # Word
from pypdf import PdfReader                  # PDF

app = Flask(__name__)
# Cap uploads at 25 MB — larger files get a clear message instead of
# exhausting the server's memory.
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")
client = OpenAI(api_key=DEEPSEEK_API_KEY, base_url="https://api.deepseek.com")


# ---------------------------------------------------------------------------
# KNOWLEDGE BASE: load the firm's standards library (files in /knowledge)
# The AI must check against THESE texts, not its own memory.
# To update a standard later: replace the file, redeploy. No code change.
# ---------------------------------------------------------------------------
def load_knowledge_base():
    kb_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "knowledge")
    sections = []
    if os.path.isdir(kb_dir):
        for fname in sorted(os.listdir(kb_dir)):
            if fname.endswith(".txt"):
                path = os.path.join(kb_dir, fname)
                try:
                    with open(path, "r", encoding="utf-8") as f:
                        sections.append(f.read().strip())
                except Exception:
                    pass
    return "\n\n==========\n\n".join(sections)


KNOWLEDGE_BASE = load_knowledge_base()


# ---------------------------------------------------------------------------
# STEP 1: Read the uploaded file and turn it into plain text the AI can review
# ---------------------------------------------------------------------------
# The AI is only ever sent this many characters, so never hold more than
# slightly above it in memory (prevents out-of-memory on huge workbooks).
MAX_EXTRACT_CHARS = 45000


def extract_text_from_file(filename, file_bytes):
    """Return the text content of an uploaded file, based on its type.
    Memory-frugal: stops reading once MAX_EXTRACT_CHARS is reached."""
    name = filename.lower()

    if name.endswith((".xlsx", ".xlsm")):
        text_parts = []
        total = 0
        wb = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        try:
            for sheet in wb.worksheets:
                if total >= MAX_EXTRACT_CHARS:
                    text_parts.append("\n[... file is large; remaining sheets not "
                                      "included in this review pass ...]")
                    break
                header = f"\n===== SHEET: {sheet.title} ====="
                text_parts.append(header)
                total += len(header)
                for row in sheet.iter_rows(values_only=True):
                    cells = [str(c) for c in row
                             if c is not None and str(c).strip() != ""]
                    if cells:
                        line = " | ".join(cells)
                        text_parts.append(line)
                        total += len(line)
                        if total >= MAX_EXTRACT_CHARS:
                            break
        finally:
            wb.close()
        return "\n".join(text_parts)

    elif name.endswith(".docx"):
        doc = DocxDocument(io.BytesIO(file_bytes))
        parts = []
        total = 0
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
        parts = []
        total = 0
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

    else:
        return None  # unsupported type


# ---------------------------------------------------------------------------
# STEP 2: The instructions that turn raw DeepSeek into an audit reviewer
# ---------------------------------------------------------------------------
REVIEWER_INSTRUCTIONS = """You are an experienced audit reviewer at an accounting firm, reviewing audit working papers to the standard expected in an ICAP Quality Control Review or an Audit Oversight Board inspection.

You will be given the text extracted from an audit working paper (often a revenue or other head, sometimes with supporting figures and calculations).

Review it carefully and identify EVERY discrepancy, error, omission, weakness, or matter needing attention. Look specifically for:
- Figures or totals that do not add up, or that do not agree between different parts of the document
- Broken spreadsheet values such as #REF!, #DIV/0!, #VALUE! — these are hard errors
- Content that appears to belong to a DIFFERENT client or engagement (wrong client name, another file reference left in from a reused template) — copy-paste contamination
- Conclusions that are pre-filled or boilerplate ("satisfactory", "fairly stated") without evidence that actual work supports them
- Missing sign-offs, missing dates, or dates out of logical sequence
- Vague or unquantified work (e.g. a "sample" with no number of items tested)
- Calculations that look wrong or unsupported
- Anything a working paper needs but is missing (evidence, cross-references, explanations)
- Non-compliance with the applicable accounting or auditing standards

For EACH issue you find, give:
1. A short, clear title of the issue
2. A plain-English explanation (simple language a junior staff member can understand — avoid unnecessary jargon)
3. The applicable standard or rule reference where you are reasonably confident (e.g. IFRS 15, IAS 24, IAS 1, ISA 500, ISA 230, Companies Act 2017). If you are NOT sure of the exact reference, say "reference to be confirmed" rather than inventing one.
4. A severity: High, Medium, or Low (or "Factual" for arithmetic/broken-value errors that are simply right or wrong)
5. A suggested fix — what the team should do to resolve it

IMPORTANT RULES:
- You are given the FIRM'S STANDARDS LIBRARY below. Base every standard reference on that library. When your finding is supported by the library, cite it (e.g. "IFRS 15 — control transfer (per firm standards library)").
- If an issue is real but the library does not cover it, still raise it, but mark the reference as "outside loaded library — reference to be confirmed".
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
  "summary": "A one or two sentence overall summary of the file's condition."
}
Return ONLY the JSON, no other text."""


def review_with_ai(document_text):
    """Send the document to DeepSeek with the reviewer instructions, get findings back."""
    # Keep the input to a sensible size (very large files get trimmed for this stage)
    trimmed = document_text[:40000]

    response = client.chat.completions.create(
        model="deepseek-chat",
        messages=[
            {"role": "system", "content": REVIEWER_INSTRUCTIONS},
            {"role": "system", "content": f"FIRM'S STANDARDS LIBRARY (check against these texts):\n\n{KNOWLEDGE_BASE}"},
            {"role": "user", "content": f"Here is the working paper to review:\n\n{trimmed}"},
        ],
        max_tokens=3000,
        temperature=0.2,  # low = more consistent, less "creative"
    )
    raw = response.choices[0].message.content.strip()
    return parse_ai_json(raw)


def parse_ai_json(raw):
    """Read the AI's JSON response, tolerating common formatting quirks."""
    import re as _re
    text = raw.strip()

    # 1) Strip ```json fences if present
    if text.startswith("```"):
        parts = text.split("```")
        if len(parts) >= 2:
            text = parts[1]
        if text.lstrip().lower().startswith("json"):
            text = text.lstrip()[4:]
        text = text.strip()

    # 2) If there's chatter before/after, cut to the outermost { ... }
    start = text.find("{")
    end = text.rfind("}")
    if start != -1 and end != -1 and end > start:
        text = text[start:end + 1]

    # 3) First attempt: parse as-is
    try:
        return json.loads(text), None
    except Exception:
        pass

    # 4) Second attempt: fix the most common quirks
    #    smart quotes -> normal quotes; remove trailing commas before } or ]
    cleaned = (text
               .replace("\u201c", '"').replace("\u201d", '"')
               .replace("\u2018", "'").replace("\u2019", "'"))
    cleaned = _re.sub(r",\s*([}\]])", r"\1", cleaned)
    try:
        return json.loads(cleaned), None
    except Exception:
        pass

    # 5) Last resort: the response was cut off mid-way. Salvage the complete
    #    findings by extracting every complete {...} object inside "findings".
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
                            pass  # skip the incomplete/broken one
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
                  f"Raw response:\n\n{raw}")


# ---------------------------------------------------------------------------
# STEP 3: The web page
# ---------------------------------------------------------------------------
PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Baker Tilly — AI Audit Reviewer</title>
<style>
  body { font-family:-apple-system,Segoe UI,Roboto,sans-serif; background:#ECEEF0; margin:0; padding:32px; color:#14233B; }
  .wrap { max-width:900px; margin:0 auto; }
  .head { display:flex; align-items:center; gap:14px; margin-bottom:8px; }
  .logo { width:42px; height:42px; border-radius:50%; background:radial-gradient(circle at 32% 30%, #2FD6D0, #0B7C7C); }
  h1 { font-size:21px; margin:0; }
  .sub { color:#5B7083; font-size:13px; margin-bottom:24px; }
  .card { background:#fff; border:1px solid #D9DDE1; border-radius:12px; padding:26px; box-shadow:0 2px 12px rgba(20,35,59,.06); margin-bottom:20px; }
  .drop { border:2px dashed #B7BFC6; border-radius:10px; padding:30px; text-align:center; }
  input[type=file] { margin:12px 0; font-size:14px; }
  button { background:#0B7C7C; color:#fff; border:none; padding:12px 24px; border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; }
  button:hover { background:#0A6E6E; }
  .notice { background:#F3ECDB; color:#5A4A28; font-size:12.5px; padding:9px 14px; border-radius:8px; margin-bottom:20px; }
  .summary { background:#EFF5F5; border:1px solid #D9DDE1; border-radius:8px; padding:14px 16px; margin-bottom:18px; font-size:14px; }
  .finding { border:1px solid #D9DDE1; border-radius:8px; margin-bottom:14px; overflow:hidden; }
  .bar { height:4px; }
  .bar.High{background:#B23A2E;} .bar.Medium{background:#B0791C;} .bar.Low{background:#5B7083;} .bar.Factual{background:#14233B;}
  .fbody { padding:14px 16px; }
  .ftop { display:flex; align-items:center; gap:8px; margin-bottom:8px; flex-wrap:wrap; }
  .sev { font-size:10.5px; font-weight:700; padding:2px 8px; border-radius:10px; }
  .sev.High{background:#F5E1DE;color:#B23A2E;} .sev.Medium{background:#F3EAD3;color:#B0791C;}
  .sev.Low{background:#EAECEE;color:#5B7083;} .sev.Factual{background:#E4E7EB;color:#14233B;}
  .ftitle { font-weight:600; font-size:15px; }
  .fexpl { font-size:13.5px; color:#3A4A64; margin-bottom:8px; }
  .ref { font-family:ui-monospace,Menlo,monospace; font-size:11.5px; background:#F0F5F5; color:#0B7C7C; padding:5px 9px; border-radius:4px; border-left:3px solid #0EA5A5; margin-bottom:8px; display:inline-block; }
  .fix { font-size:12.5px; color:#3A4A64; background:#F7F9F9; border:1px solid #eee; border-radius:6px; padding:8px 10px; }
  .fix b{color:#14233B;}
  .disclaimer { margin-top:22px; padding:16px; background:#FBF6EE; border:1px solid #E8D9BE; border-radius:8px; font-size:12.5px; color:#5A4A28; line-height:1.6; }
  .err { background:#FBEAE8; border:1px solid #E4B4AD; color:#B23A2E; padding:14px; border-radius:8px; font-size:13px; white-space:pre-wrap; }
  .spin { display:inline-block; }
</style>
</head>
<body>
<div class="wrap">
  <div class="head">
    <div class="logo"></div>
    <div><h1>AI Audit Reviewer</h1></div>
  </div>
  <div class="sub">Baker Tilly · Revenue &amp; working-paper review (Stage 1)</div>

  <div class="notice"><b>Note:</b> This build reviews against the firm's loaded standards library (summaries of IFRS 15, IAS 24, IAS 1, ISA 230, ISA 500 and related requirements). Official full texts can replace the summaries at any time without code changes. Use sample / public data until the tool moves to the firm's own server.</div>

  <div class="card">
    <form method="POST" enctype="multipart/form-data">
      <div class="drop">
        <div style="font-weight:600; margin-bottom:6px;">Upload a working paper to review</div>
        <div style="font-size:12.5px; color:#5B7083;">Excel (.xlsx), Word (.docx), PDF, or CSV</div>
        <input type="file" name="file" accept=".xlsx,.xlsm,.docx,.pdf,.csv,.txt" required>
        <br>
        <button type="submit">Review this file</button>
      </div>
    </form>
  </div>

  {% if error %}
    <div class="card"><div class="err">{{ error }}</div></div>
  {% endif %}

  {% if result %}
    <div class="card">
      <h2 style="font-size:17px; margin:0 0 12px;">Review Points — {{ filename }}</h2>
      {% if result.summary %}
        <div class="summary"><b>Overall:</b> {{ result.summary }}</div>
      {% endif %}

      {% for f in result.findings %}
        <div class="finding">
          <div class="bar {{ f.severity }}"></div>
          <div class="fbody">
            <div class="ftop">
              <span class="ftitle">{{ f.title }}</span>
              <span class="sev {{ f.severity }}">{{ f.severity }}</span>
            </div>
            <div class="fexpl">{{ f.explanation }}</div>
            {% if f.reference %}<div class="ref">{{ f.reference }}</div>{% endif %}
            <div class="fix"><b>Suggested fix:</b> {{ f.fix }}</div>
          </div>
        </div>
      {% endfor %}

      <div class="disclaimer">
        <b>Professional judgement statement:</b> This review has been prepared by an AI-assisted tool to support the audit review process by identifying possible discrepancies, errors, omissions, and matters requiring attention. It does not replace the judgement of the engagement team. All findings above are observations for your consideration, not conclusions. Every point should be reviewed, verified, and decided upon by a qualified member of the audit team. Final responsibility for the audit — including all professional judgements, the sufficiency of audit evidence, and the audit opinion — rests entirely with the Engagement Partner and the audit team, not with this tool. The AI does not sign off, approve, or conclude on any matter.
      </div>
    </div>
  {% endif %}
</div>
</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def home():
    result = None
    error = None
    filename = None

    if request.method == "POST":
        if not DEEPSEEK_API_KEY:
            error = "The DeepSeek API key is not set. Add it in Render's Environment Variables."
            return render_template_string(PAGE, result=None, error=error, filename=None)

        uploaded = request.files.get("file")
        if not uploaded or uploaded.filename == "":
            error = "Please choose a file first."
            return render_template_string(PAGE, result=None, error=error, filename=None)

        filename = uploaded.filename
        file_bytes = uploaded.read()

        try:
            text = extract_text_from_file(filename, file_bytes)
        except Exception as e:
            error = f"Could not read the file. Details: {e}"
            return render_template_string(PAGE, result=None, error=error, filename=filename)

        if text is None:
            error = "Unsupported file type. Please upload Excel, Word, PDF, or CSV."
        elif text.strip() == "":
            error = "The file appears to be empty, or its text could not be read (a scanned PDF with no text layer, perhaps)."
        else:
            result, ai_error = review_with_ai(text)
            if ai_error:
                error = ai_error

    return render_template_string(PAGE, result=result, error=error, filename=filename)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
