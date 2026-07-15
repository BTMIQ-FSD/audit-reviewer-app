"""
Baker Tilly AI Audit Reviewer — Piece 1 (Foundation Test)
This tiny app proves the whole chain works: browser -> our app -> DeepSeek -> back.
Once this runs live, we build the real review features on top of it.
"""

import os
from flask import Flask, request, render_template_string
from openai import OpenAI

app = Flask(__name__)

# The DeepSeek API key is read from a secure "environment variable" (set on Render),
# so the secret key is NEVER written inside this code file.
DEEPSEEK_API_KEY = os.environ.get("DEEPSEEK_API_KEY", "")

# DeepSeek uses the same connection style as OpenAI, just pointed at DeepSeek's address.
client = OpenAI(
    api_key=DEEPSEEK_API_KEY,
    base_url="https://api.deepseek.com",
)

# The web page (kept simple for this foundation test).
PAGE = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Baker Tilly AI Audit Reviewer — Foundation Test</title>
  <style>
    body { font-family: -apple-system, Segoe UI, Roboto, sans-serif; background:#ECEEF0;
           margin:0; padding:40px; color:#14233B; }
    .card { max-width:640px; margin:0 auto; background:#fff; border:1px solid #D9DDE1;
            border-radius:12px; padding:32px; box-shadow:0 2px 12px rgba(20,35,59,.06); }
    .logo { width:44px; height:44px; border-radius:50%;
            background:radial-gradient(circle at 32% 30%, #2FD6D0, #0B7C7C); margin-bottom:14px; }
    h1 { font-size:20px; margin:0 0 4px; }
    .sub { color:#5B7083; font-size:13px; margin-bottom:24px; }
    textarea { width:100%; box-sizing:border-box; padding:12px; border:1px solid #B7BFC6;
               border-radius:8px; font-size:14px; font-family:inherit; resize:vertical; min-height:80px; }
    button { margin-top:12px; background:#0B7C7C; color:#fff; border:none; padding:12px 22px;
             border-radius:8px; font-size:14px; font-weight:600; cursor:pointer; }
    button:hover { background:#0A6E6E; }
    .answer { margin-top:24px; padding:18px; background:#F6F8F8; border:1px solid #D9DDE1;
              border-radius:8px; font-size:14px; line-height:1.6; white-space:pre-wrap; }
    .status-ok { color:#1F6B4F; font-weight:600; }
    .status-bad { color:#B23A2E; font-weight:600; }
    .note { font-size:12px; color:#5B7083; margin-top:20px; border-top:1px solid #eee; padding-top:14px; }
  </style>
</head>
<body>
  <div class="card">
    <div class="logo"></div>
    <h1>AI Audit Reviewer — Foundation Test</h1>
    <div class="sub">Baker Tilly · Piece 1 · Confirms the app is live and connected to the AI</div>

    <form method="POST">
      <textarea name="question" placeholder="Type anything to test — e.g. 'In one line, what does IFRS 15 cover?'">{{ question or "" }}</textarea>
      <br>
      <button type="submit">Ask the AI</button>
    </form>

    {% if answer %}
      <div class="answer"><span class="status-ok">AI responded ✓</span>

{{ answer }}</div>
    {% endif %}

    {% if error %}
      <div class="answer"><span class="status-bad">Something went wrong ✗</span>

{{ error }}</div>
    {% endif %}

    <div class="note">
      This is a foundation test. If you see an AI answer above, the whole chain works
      (browser → app → DeepSeek → back) and we can start building the real review features.
    </div>
  </div>
</body>
</html>
"""


@app.route("/", methods=["GET", "POST"])
def home():
    answer = None
    error = None
    question = None

    if request.method == "POST":
        question = request.form.get("question", "").strip()
        if not question:
            error = "Please type a question first."
        elif not DEEPSEEK_API_KEY:
            error = "The DeepSeek API key is not set yet. (We set this on Render in a later step.)"
        else:
            try:
                response = client.chat.completions.create(
                    model="deepseek-chat",
                    messages=[{"role": "user", "content": question}],
                    max_tokens=300,
                )
                answer = response.choices[0].message.content
            except Exception as e:
                error = f"Could not reach DeepSeek. Details: {e}"

    return render_template_string(PAGE, answer=answer, error=error, question=question)


# This lets Render run the app. It reads the port Render gives us.
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
