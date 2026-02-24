"""
DevOps Job Tracker — Germany Remote
Searches Adzuna API for DevOps/SRE/Platform Engineer roles,
scores them against Seetharam's resume, exports to Excel + jobs.json.
"""

import os, re, json, smtplib, logging, requests
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email import encoders
from pathlib import Path
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from dotenv import load_dotenv

load_dotenv()
logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

# ─── Resume Profile ────────────────────────────────────────────────────────────
RESUME_SKILLS = {
    "cloud":      ["azure","aws","gcp","cloud","eks","aks"],
    "kubernetes": ["kubernetes","k8s","eks","aks","helm","kustomize","kubectl"],
    "cicd":       ["ci/cd","cicd","github actions","gitlab ci","azure devops","jenkins","argocd","gitops"],
    "iac":        ["terraform","infrastructure as code","iac","ansible","pulumi"],
    "containers": ["docker","containers","containerization","helm","karpenter"],
    "monitoring": ["prometheus","grafana","loki","observability","monitoring","fluent-bit"],
    "security":   ["sonarqube","trivy","zero trust","waf","cert-manager","owasp","sast"],
    "scripting":  ["python","bash","shell","scripting"],
    "gitops":     ["argocd","gitops","argo rollouts","flux","blue-green","canary"],
    "platform":   ["platform engineering","sre","site reliability","devops"],
}

SKILL_WEIGHTS = {
    "kubernetes":20,"cicd":18,"cloud":17,"iac":15,
    "containers":12,"gitops":10,"monitoring":8,
    "security":7,"scripting":6,"platform":7,
}

MATCH_THRESHOLD = 55

# ─── Adzuna Search ─────────────────────────────────────────────────────────────
ADZUNA_APP_ID  = os.getenv("ADZUNA_APP_ID")
ADZUNA_API_KEY = os.getenv("ADZUNA_API_KEY")
SEARCH_QUERIES = [
    "DevOps Engineer remote",
    "Site Reliability Engineer remote",
    "Platform Engineer remote",
    "Cloud Infrastructure Engineer remote",
    "Kubernetes Engineer remote",
]

def fetch_adzuna_jobs(query, max_results=20):
    if not ADZUNA_APP_ID or not ADZUNA_API_KEY:
        log.warning("Adzuna credentials missing — set ADZUNA_APP_ID and ADZUNA_API_KEY in .env")
        return []
    url = "https://api.adzuna.com/v1/api/jobs/de/search/1"
    params = {
        "app_id": ADZUNA_APP_ID,
        "app_key": ADZUNA_API_KEY,
        "results_per_page": max_results,
        "what": query,
        "max_days_old": 1,
        "content-type": "application/json",
    }
    try:
        r = requests.get(url, params=params, timeout=15)
        r.raise_for_status()
        return r.json().get("results", [])
    except Exception as e:
        log.error(f"Adzuna fetch error for '{query}': {e}")
        return []

def collect_all_jobs():
    seen, jobs = set(), []
    for q in SEARCH_QUERIES:
        log.info(f"Searching: {q}")
        for job in fetch_adzuna_jobs(q):
            jid = job.get("id")
            if jid and jid not in seen:
                seen.add(jid)
                jobs.append(job)
    log.info(f"Total unique jobs fetched: {len(jobs)}")
    return jobs

# ─── Scoring ───────────────────────────────────────────────────────────────────
def score_job(job):
    text = " ".join([
        job.get("title",""),
        job.get("description",""),
        job.get("category",{}).get("label",""),
    ]).lower()

    total_weight = sum(SKILL_WEIGHTS.values())
    earned = 0
    matched = []
    for category, keywords in RESUME_SKILLS.items():
        for kw in keywords:
            if kw in text:
                earned += SKILL_WEIGHTS.get(category, 5)
                matched.append(kw.title())
                break

    for kw in ["remote","home office","homeoffice","distributed"]:
        if kw in text:
            earned += 5
            break

    score = min(100, int((earned / total_weight) * 100))
    return score, list(dict.fromkeys(matched))

def classify_remote(job):
    text = (job.get("description","") + job.get("title","")).lower()
    if "fully remote" in text or "100% remote" in text:
        return "Fully Remote"
    if "remote" in text or "home office" in text or "homeoffice" in text:
        return "Hybrid / Remote"
    return "On-site"

def summarize(job):
    desc = job.get("description","")
    clean = re.sub(r"<[^>]+>"," ", desc)
    clean = re.sub(r"\s+"," ", clean).strip()
    return (clean[:220] + "…") if len(clean) > 220 else clean

def get_career_url(job):
    redirect = job.get("redirect_url","")
    company  = job.get("company",{}).get("display_name","")
    if company:
        slug = re.sub(r"[^a-z0-9]","", company.lower())
        return f"https://www.{slug}.com/careers"
    return redirect

# ─── JSON Export (for GitHub Pages UI) ────────────────────────────────────────
def export_json(jobs_data, docs_dir="docs"):
    """Write jobs.json into the GitHub Pages docs/ folder."""
    Path(docs_dir).mkdir(parents=True, exist_ok=True)
    payload = {
        "updated": datetime.utcnow().isoformat() + "Z",
        "jobs": jobs_data,
    }
    out = Path(docs_dir) / "jobs.json"
    out.write_text(json.dumps(payload, indent=2, ensure_ascii=False))
    log.info(f"jobs.json written → {out}")

# ─── Excel Export ──────────────────────────────────────────────────────────────
HEADERS = ["Job Title","Company","Location","Remote Type",
           "Date Posted","Match Score","Key Matching Skills","Role Summary","Career Page URL"]
HEADER_COLOR = "1F3864"

def _border():
    s = Side(style="thin", color="B8CCE4")
    return Border(left=s, right=s, top=s, bottom=s)

def score_fill(score):
    color = "C6EFCE" if score >= 75 else "FFEB9C" if score >= 60 else "FCE4D6"
    return PatternFill("solid", fgColor=color)

def build_excel(jobs_data, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Matching Jobs"
    ws.freeze_panes = "A3"

    ws.merge_cells("A1:I1")
    tc = ws["A1"]
    tc.value = f"DevOps Job Matches — Germany Remote  |  {datetime.now().strftime('%d %b %Y %H:%M')}"
    tc.font  = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    tc.fill  = PatternFill("solid", fgColor=HEADER_COLOR)
    tc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    for col, h in enumerate(HEADERS, 1):
        cell = ws.cell(row=2, column=col, value=h)
        cell.font      = Font(name="Arial", bold=True, color="FFFFFF", size=11)
        cell.fill      = PatternFill("solid", fgColor=HEADER_COLOR)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _border()
    ws.row_dimensions[2].height = 22

    for row_idx, job in enumerate(jobs_data, start=3):
        fill   = score_fill(job["score"])
        values = [job["title"],job["company"],job["location"],job["remote_type"],
                  job["date_posted"],job["score"],", ".join(job["matched_skills"]),
                  job["summary"],job["career_url"]]
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font      = Font(name="Arial", size=10)
            cell.fill      = fill
            cell.border    = _border()
            cell.alignment = Alignment(wrap_text=True, vertical="top")

        ws.cell(row=row_idx, column=6).font      = Font(name="Arial", bold=True, size=11)
        ws.cell(row=row_idx, column=6).alignment = Alignment(horizontal="center", vertical="top")
        url_cell = ws.cell(row=row_idx, column=9)
        url_cell.hyperlink = job["career_url"]
        url_cell.font = Font(name="Arial", size=10, color="0563C1", underline="single")
        ws.row_dimensions[row_idx].height = 60

    for i, w in enumerate([36,22,20,18,14,12,36,60,40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.auto_filter.ref = f"A2:I{len(jobs_data)+2}"

    ws2 = wb.create_sheet("Summary Stats")
    ws2.merge_cells("A1:B1")
    ws2["A1"].value = "Summary"
    ws2["A1"].font  = Font(name="Arial", bold=True, size=13, color="FFFFFF")
    ws2["A1"].fill  = PatternFill("solid", fgColor=HEADER_COLOR)
    stats = [
        ("Total Jobs",len(jobs_data)),
        ("High Match (≥75)",sum(1 for j in jobs_data if j["score"]>=75)),
        ("Good Match (60–74)",sum(1 for j in jobs_data if 60<=j["score"]<75)),
        ("Fully Remote",sum(1 for j in jobs_data if "Fully" in j["remote_type"])),
        ("Run Date",datetime.now().strftime("%Y-%m-%d %H:%M")),
    ]
    for r,(label,val) in enumerate(stats,2):
        ws2.cell(row=r,column=1,value=label).font = Font(name="Arial",bold=True,size=10)
        ws2.cell(row=r,column=2,value=val).font   = Font(name="Arial",size=10)
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 20

    wb.save(output_path)
    log.info(f"Excel saved: {output_path}")

# ─── Email ─────────────────────────────────────────────────────────────────────
def send_email(filepath, job_count):
    smtp_host = os.getenv("SMTP_HOST","smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT",587))
    smtp_user = os.getenv("SMTP_USER")
    smtp_pass = os.getenv("SMTP_PASS")
    to_email  = os.getenv("TO_EMAIL")
    if not all([smtp_user,smtp_pass,to_email]):
        log.warning("Email credentials missing — skipping.")
        return
    msg = MIMEMultipart()
    msg["From"]    = smtp_user
    msg["To"]      = to_email
    msg["Subject"] = f"DevOps Job Digest — {job_count} matches — {datetime.now().strftime('%d %b %Y')}"
    msg.attach(MIMEText(
        f"Hi Seetharam,\n\n{job_count} matching DevOps roles found in Germany today.\n"
        f"Excel attached. Also check your GitHub Pages dashboard for the live view.\n\n— Job Tracker Bot",
        "plain"))
    with open(filepath,"rb") as f:
        part = MIMEBase("application","octet-stream")
        part.set_payload(f.read())
    encoders.encode_base64(part)
    part.add_header("Content-Disposition",f"attachment; filename={Path(filepath).name}")
    msg.attach(part)
    try:
        with smtplib.SMTP(smtp_host,smtp_port) as srv:
            srv.starttls()
            srv.login(smtp_user,smtp_pass)
            srv.send_message(msg)
        log.info(f"Email sent to {to_email}")
    except Exception as e:
        log.error(f"Email failed: {e}")

# ─── Main ──────────────────────────────────────────────────────────────────────
def run(docs_dir="docs"):
    log.info("=" * 60)
    log.info("DevOps Job Tracker — Starting")
    log.info("=" * 60)

    raw_jobs  = collect_all_jobs()
    processed = []
    for job in raw_jobs:
        score, matched = score_job(job)
        if score < MATCH_THRESHOLD:
            continue
        created = job.get("created","")
        try:
            from datetime import datetime, timezone
            dt = datetime.fromisoformat(created.replace("Z","+00:00"))
            date_str = dt.strftime("%d %b %Y %H:%M")
        except Exception:
            date_str = created[:10] if created else "Unknown"

        processed.append({
            "title":         job.get("title","N/A"),
            "company":       job.get("company",{}).get("display_name","N/A"),
            "location":      job.get("location",{}).get("display_name","Germany"),
            "remote_type":   classify_remote(job),
            "date_posted":   date_str,
            "score":         score,
            "matched_skills":matched,
            "summary":       summarize(job),
            "career_url":    get_career_url(job),
        })

    processed.sort(key=lambda x: x["score"], reverse=True)

    output_dir = Path("output")
    output_dir.mkdir(exist_ok=True)
    filename = f"devops_jobs_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    filepath = str(output_dir / filename)

    if processed:
        build_excel(processed, filepath)
        export_json(processed, docs_dir=docs_dir)   # ← writes docs/jobs.json
        send_email(filepath, len(processed))
        log.info(f"Done. {len(processed)} jobs written.")
    else:
        log.info("No matching jobs found.")
        export_json([], docs_dir=docs_dir)

    return filepath, len(processed)

if __name__ == "__main__":
    run()