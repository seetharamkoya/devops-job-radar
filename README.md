# 🚀 DevOps Job Radar — Germany

A self-hosted job intelligence dashboard that:
- **Scrapes** Adzuna daily for DevOps / SRE / Platform Engineer roles in Germany
- **Scores** each job against your resume (Kubernetes, Terraform, ArgoCD, Azure, CI/CD…)
- **Publishes** a beautiful dashboard to GitHub Pages — live, searchable, filterable
- **Emails** you an Excel file with the day's best matches

**Live site:** `https://YOUR_USERNAME.github.io/YOUR_REPO/`

---

## Repo Structure

```
├── docs/
│   ├── index.html       ← GitHub Pages UI (do not edit)
│   └── jobs.json        ← Auto-updated daily by the scanner
├── .github/
│   └── workflows/
│       └── job_scan.yml ← Daily cron + auto-deploy
├── job_tracker.py       ← Core scanner script
├── requirements.txt
└── .env.example
```

---

## Setup (15 minutes)

### 1. Fork / Clone this repo

```bash
git clone https://github.com/YOUR_USERNAME/devops-job-radar.git
cd devops-job-radar
```

### 2. Enable GitHub Pages

1. Go to your repo → **Settings** → **Pages**
2. Source: **Deploy from a branch**
3. Branch: `main` / folder: `/docs`
4. Save — your site will be live at `https://YOUR_USERNAME.github.io/YOUR_REPO/`

### 3. Get free Adzuna API credentials

Sign up at [developer.adzuna.com](https://developer.adzuna.com/signup) — free, 250 req/day.

### 4. Add GitHub Secrets

Go to **Settings → Secrets → Actions → New repository secret** and add:

| Secret | Value |
|--------|-------|
| `ADZUNA_APP_ID` | Your Adzuna App ID |
| `ADZUNA_API_KEY` | Your Adzuna API Key |
| `SMTP_USER` | your.email@gmail.com |
| `SMTP_PASS` | Gmail App Password ([get one](https://myaccount.google.com/apppasswords)) |
| `TO_EMAIL` | seetharamaiah.koya@gmail.com |

### 5. Run manually to test

In your repo → **Actions** → **Daily DevOps Job Scan + Deploy** → **Run workflow**

This will:
- Fetch today's jobs from Adzuna
- Score them against your resume
- Commit `docs/jobs.json` to your repo
- GitHub Pages auto-deploys — your dashboard updates within ~60 seconds

---

## How It Works

```
GitHub Actions (daily 6 AM UTC)
  │
  ├─► job_tracker.py
  │     ├─ Searches Adzuna API (5 queries × 20 results)
  │     ├─ Deduplicates by job ID
  │     ├─ Scores each job against resume skills
  │     ├─ Filters jobs ≥ 55% match
  │     ├─ Writes output/devops_jobs_YYYYMMDD.xlsx
  │     └─ Writes docs/jobs.json  ◄────────────┐
  │                                             │
  ├─► git commit docs/jobs.json                 │
  │     └─ GitHub Pages redeploys               │
  │           └─ index.html fetches jobs.json ──┘
  │
  └─► Email with Excel attachment
```

---

## Scoring Weights

| Skill | Weight |
|-------|--------|
| Kubernetes / K8s | 20% |
| CI/CD Pipelines | 18% |
| Cloud (Azure/AWS/GCP) | 17% |
| Terraform / IaC | 15% |
| Docker / Containers | 12% |
| GitOps / ArgoCD | 10% |
| Monitoring (Prometheus/Grafana) | 8% |
| Security (Trivy/SonarQube) | 7% |
| Platform Eng / SRE | 7% |
| Python / Bash | 6% |

Only jobs scoring **≥ 55%** are shown. Edit `MATCH_THRESHOLD` in `job_tracker.py` to adjust.

---

## Local Development

```bash
pip install -r requirements.txt
cp .env.example .env
# Fill in .env
python job_tracker.py

# Preview the UI locally
cd docs && python -m http.server 8080
# Open http://localhost:8080
```