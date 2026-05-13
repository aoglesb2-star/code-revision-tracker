# Municipal Code Revision Tracker

Compares two versions of a municipal code (old vs. new) and produces a
formatted Excel redline with green additions, red strikethrough deletions,
amber amendments, and a plain-English council memo summary.

---

## What you need before starting

- A free [GitHub](https://github.com) account
- A free [Render](https://render.com) account
- An [Anthropic API key](https://console.anthropic.com)

---

## Step 1 — Put the code on GitHub

1. Go to [github.com](https://github.com) → **+** → **New repository**
2. Name it `code-revision-tracker`, click **Create repository**
3. Click **uploading an existing file**
4. Upload all files from this folder:
   - `app.py`
   - `requirements.txt`
   - `Procfile`
   - `render.yaml`
   - `.gitignore`
   - `templates/index.html`
5. Click **Commit changes**

---

## Step 2 — Get your Anthropic API key

1. Go to [console.anthropic.com](https://console.anthropic.com)
2. Click **API Keys** → **Create Key**
3. Copy and save the key securely

---

## Step 3 — Deploy on Render

1. Go to [render.com](https://render.com) → **New +** → **Web Service**
2. Connect your GitHub account and select `code-revision-tracker`
3. Render auto-detects settings from `render.yaml`. Confirm:
   - Environment: Python
   - Build command: `pip install -r requirements.txt`
   - Start command: `gunicorn app:app --timeout 300 --workers 1`
4. Under **Environment Variables**, add:
   - Key: `ANTHROPIC_API_KEY`
   - Value: *(your key)*
   - Key: `APP_PASSWORD`
   - Value: *(any password you choose — users will need this to run comparisons)*
5. Click **Create Web Service** — deploys in 2–4 minutes

---

## Step 4 — Use it

1. Open the Render URL (e.g. `https://code-revision-tracker.onrender.com`)
2. Enter the **access password** you set in step 3
3. Upload or paste the **old version** on the left
4. Upload or paste the **new version** on the right
5. Click **Generate Change Log**
6. Download the Excel file — it has three tabs:
   - **Summary** — stats, council memo text, and warnings (if any)
   - **Change Log** — every change, color-coded
   - **Deleted Sections** — all deletions in one place

---

## Powered by

This app uses **Claude Opus 4.7** (`claude-opus-4-7`) for the comparison — Anthropic's most capable model, chosen for legal-accuracy work where verbatim quoting matters. Switch to `claude-sonnet-4-6` in `app.py` if you'd rather trade some accuracy for ~5× cheaper runs.

---

## Accepts

| Format | Notes |
|--------|-------|
| `.docx` | Word documents (15 MB max) |
| `.pdf`  | Text-based PDFs, 15 MB max (not scanned images) |
| `.txt`  | Plain text (15 MB max) |
| Paste   | Copy/paste directly into the text box — best option for very large docs |

---

## Costs

- GitHub: Free
- Render free tier: Free (sleeps after 15 min idle; ~30s wake time)
- Anthropic API (Opus 4.7): ~$1–$5 per comparison depending on document length

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| "ANTHROPIC_API_KEY not configured" | Check Environment Variables in Render dashboard |
| "APP_PASSWORD not configured" | Set `APP_PASSWORD` in Render Environment Variables |
| "Invalid or missing password" | Make sure the password field on the page matches `APP_PASSWORD` exactly |
| PDF shows no text | PDF may be a scanned image — use OCR first or paste text manually |
| "Document was truncated" warning in Excel | Compare the document in smaller sections |
| "AI response was cut off" | Too many changes for one run — split the comparison |
| Times out | Very large documents — upgrade to Render Starter ($7/mo) for longer timeout |
| File too large (15 MB) | Paste the text directly instead of uploading |
| App won't start | Check Render Logs tab |
