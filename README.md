# DepEd Grade 8 ILAW Lesson Planning Agent

AI-powered lesson planning tool for Grade 8 teachers aligned with:
- **ILAW Framework** (DepEd Order No. 09, s. 2026)
- **SY 2026–2027 Three-Term Calendar**
- **MATATAG Curriculum**

## Features

- Generates complete **ILAW lesson plans** (Intentions, Learning Experiences, Assessment, Ways Forward)
- Generates **student-facing classroom PowerPoint presentations** (.pptx)
- Supports **TLE 8, Values Education 8, Filipino 8**
- Includes **Declaration of AI Use** per DepEd Order No. 03, s. 2026
- Downloads: `.pptx` (student slides), `.docx` and `.pdf` (lesson plan)

---

## Deploy to Vercel

### Step 1 — Push to GitHub

```bash
git init
git add .
git commit -m "Initial commit — DepEd ILAW Agent"
git remote add origin https://github.com/YOUR_USERNAME/deped-ilaw-agent.git
git push -u origin main
```

### Step 2 — Deploy on Vercel

1. Go to [vercel.com](https://vercel.com) and log in
2. Click **Add New Project**
3. Import your GitHub repository
4. Leave all settings as default — Vercel auto-detects the config
5. Click **Deploy**

Your app will be live at `https://deped-ilaw-agent.vercel.app` (or your custom URL).

### Step 3 — Set Environment Variable (Important!)

The app uses the Anthropic API for AI generation. In your Vercel project:

1. Go to **Settings → Environment Variables**
2. The API key is currently called from the browser directly (no server-side key needed for the AI part)
3. No additional setup required — the app works out of the box

---

## Project Structure

```
deped-ilaw-agent/
├── api/
│   └── generate-pptx.js    # Vercel serverless function — builds .pptx
├── public/
│   └── index.html           # Frontend web app
├── package.json
└── vercel.json              # Routing configuration
```

---

## Local Development

```bash
npm install
npx vercel dev
```

Open `http://localhost:3000`

---

## How It Works

1. Teacher fills in subject, term, week, and topic
2. App calls **Claude Sonnet** (Anthropic API) to generate ILAW lesson plan + slide content
3. Teacher previews the lesson plan and slide outline
4. Clicking **Download Student .pptx** calls the `/api/generate-pptx` serverless function
5. Server builds a full colorful 10-slide classroom presentation and returns the file

---

## Built With

- [PptxGenJS](https://gitbrent.github.io/PptxGenJS/) — PowerPoint generation
- [Claude API](https://anthropic.com) — AI lesson plan generation
- [Vercel](https://vercel.com) — Serverless hosting

---

*Built for Filipino Grade 8 teachers · Department of Education · SY 2026–2027*
