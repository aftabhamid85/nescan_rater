#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
import io
import math
from datetime import datetime

app = Flask(__name__)
app.config["SECRET_KEY"] = "change-me"

# ---------- Load organisations ----------
def load_organisations(path="data/Organisations Name.xlsx"):
    df = pd.read_excel(path)

    # ✅ First try to find your actual Excel column
    if "Organisation Names" in df.columns:
        col = "Organisation Names"
    elif "Organisation" in df.columns:
        col = "Organisation"
    else:
        # fallback: just use the first column if names differ
        col = df.columns[0]

    orgs = df[col].dropna().astype(str).tolist()
    return orgs


ORGS = load_organisations()


# ---------- Options ----------
SECTOR_OPTIONS = [
    "Public Sector", "Third Sector", "Corporate", "Education/Research",
    "Media", "Parliament/Political"
]
SUBJECT_OPTIONS = [
    "Lobbying & Activism", "Building & Energy", "Action & Justice",
    "Education & Awareness", "Land & Nature", "Organisation Development",
    "Climate Know-How", "Food", "Circular Economy", "Transport",
    "Education", "Health", "Culture", "Biodiversity", "Multi sectoral Approach"
]
TYPE_OPTIONS = [
    "Local Authority", "National Authority", "Regulatory Authority",
    "Institutional Donor", "Charity/Not for Profit", "Trusts and Foundations",
    "Higher Education Institution", "Research institute", "Schools/Colleges",
    "Individual", "Media", "Network/Forum etc", "Schools"
]
GEO_OPTIONS = ["Local", "Regional", "National", "International"]

# ---------- Questions & labels (5 = midpoint phrasing) ----------
POWER_QUESTIONS = [
    dict(
        id="p_influence",
        text="Influence on Policy & Decisions",
        long="What is the stakeholder's capacity to directly influence or shape policy, project approvals, or key decisions within their sector or geographic area?",
        label1="No influence",
        label5="Moderate, occasional influence",
        label10="Directly influences major decisions",
    ),
    dict(
        id="p_resources",
        text="Control of Critical Resources",
        long="To what extent does this stakeholder control financial resources (e.g., grant funding, sponsorship), skills, or in-kind assets that are essential for NESCAN's work?",
        label1="No control",
        label5="Controls some relevant resources",
        label10="Controls significant resources",
    ),
    dict(
        id="p_reputation",
        text="Public Standing & Reputation",
        long="What is the stakeholder's ability to affect NESCAN's reputation, either positively or negatively, through their public influence and standing as a trusted partner?",
        label1="No effect",
        label5="Moderate effect on public trust",
        label10="Strong positive/negative influence",
    ),
    dict(
        id="p_audience",
        text="Impact on Target Audience",
        long="How significant is this stakeholder's reach and influence over the communities, organizations, or individuals that NESCAN aims to serve or engage?",
        label1="Minimal reach",
        label5="Moderate reach within target audience",
        label10="Influences a large portion of the audience",
    ),
]

INTEREST_QUESTIONS = [
    dict(
        id="i_alignment",
        text="Alignment with Mission & Vision",
        long="How closely do the stakeholder's core mission, values, and strategic objectives align with NESCAN's vision for community climate action and a just transition?",
        label1="Completely misaligned",
        label5="Partially aligned with some shared goals",
        label10="Perfectly aligned",
    ),
    dict(
        id="i_engagement",
        text="Current Level of Engagement",
        long="How actively engaged is the stakeholder with NESCAN's activities, communications, or network? Is there an existing relationship that is active and mutually beneficial?",
        label1="No prior contact",
        label5="Occasional/limited engagement",
        label10="Regular, high-level interaction",
    ),
    dict(
        id="i_partnership",
        text="Potential for Partnership & Collaboration",
        long="How strong is the opportunity for a high-impact collaboration or a new project with this stakeholder?",
        label1="No potential",
        label5="Moderate potential (needs development)",
        label10="Immediate, high-potential opportunities",
    ),
    dict(
        id="i_overlap",
        text="Overlap of Services & Competitiveness",
        long="To what degree do this stakeholder's services or objectives complement NESCAN's, rather than compete for the same funding, members, or projects? ",
        label1="Direct competitor",
        label5="Some overlap; generally complementary",
        label10="Highly complementary, non-competitive",
    ),
    dict(
        id="i_value",
        text="Strategic Value to Future Goals",
        long="How critical is this stakeholder to achieving one or more of NESCAN's long-term strategic objectives, such as scaling a program, influencing policy, or securing major funding?",
        label1="Not critical",
        label5="Helpful but not pivotal",
        label10="Essential for long-term success",
    ),
    dict(
        id="i_champions",
        text="Internal Champions & Relationships",
        long="Is there a specific individual or department within the stakeholder organization that is a known champion or ally for NESCAN's work?",
        label1="No known contact",
        label5="At least one contact; limited championing",
        label10="Multiple high-level champions",
    ),
]

# ---------- In-memory results (single-rater per run) ----------
RESULT_ROWS = []
CURRENT_RATER = None

def to_float(v):
    try:
        return float(v)
    except:
        return math.nan

def compute_scores(row):
    # Averages (1–10)
    power_vals = [to_float(row.get(q["id"])) for q in POWER_QUESTIONS]
    interest_vals = [to_float(row.get(q["id"])) for q in INTEREST_QUESTIONS]
    power_vals = [v for v in power_vals if not math.isnan(v)]
    interest_vals = [v for v in interest_vals if not math.isnan(v)]

    power_score = sum(power_vals) / len(power_vals) if power_vals else math.nan
    interest_score = sum(interest_vals) / len(interest_vals) if interest_vals else math.nan

    combined_avg = (power_score + interest_score) / 2 if (not math.isnan(power_score) and not math.isnan(interest_score)) else math.nan
    combined_total = (power_score + interest_score) if (not math.isnan(power_score) and not math.isnan(interest_score)) else math.nan

    return power_score, interest_score, combined_avg, combined_total

def quadrant_from_avg(avg):
    if math.isnan(avg):
        return ""
    if avg >= 8.0:
        return "Manage Closely"
    if avg >= 6.0:
        return "Keep Satisfied"
    if avg >= 3.0:
        return "Keep Informed"
    return "Monitor"

@app.route("/", methods=["GET", "POST"])
def index():
    global CURRENT_RATER, RESULT_ROWS
    if request.method == "POST":
        CURRENT_RATER = request.form.get("rater_name", "").strip()
        RESULT_ROWS = []  # reset per rater session
        return redirect(url_for("rate", idx=0))
    return render_template("index.html")

@app.route("/rate/<int:idx>", methods=["GET", "POST"])
def rate(idx):
    global RESULT_ROWS, CURRENT_RATER

    if idx >= len(ORGS):
        return redirect(url_for("done"))

    org = ORGS[idx]

    # Preserve values if validation fails (optional/UX)
    form_values = {k: request.form.get(k) for k in request.form} if request.method == "POST" else {}

    if request.method == "POST":
        action = request.form.get("action")

        if action == "skip":
            # Only store org + rater
            RESULT_ROWS.append({
                "Organisation": org,
                "Rater_Name": CURRENT_RATER,
                # Other fields intentionally blank
            })
            return redirect(url_for("rate", idx=idx + 1))

        # Validate required dropdowns (server-side)
        required_fields = ["sector", "subject_area", "org_type", "geo"]
        missing = [f for f in required_fields if not request.form.get(f)]
        # Likert fields required too
        missing += [q["id"] for q in POWER_QUESTIONS if not request.form.get(q["id"])]
        missing += [q["id"] for q in INTEREST_QUESTIONS if not request.form.get(q["id"])]

        if missing:
            # Re-render with previously entered values
            return render_template(
                "rate.html",
                rater=CURRENT_RATER,
                idx=idx, total=len(ORGS), org=org,
                sector_options=SECTOR_OPTIONS,
                subject_options=SUBJECT_OPTIONS,
                type_options=TYPE_OPTIONS,
                geo_options=GEO_OPTIONS,
                power_questions=POWER_QUESTIONS,
                interest_questions=INTEREST_QUESTIONS,
                form_values=form_values,
            )

        # Build row
        row = {
            "Organisation": org,
            "Rater_Name": CURRENT_RATER,
            "Sector": request.form.get("sector"),
            "Subject_Area": request.form.get("subject_area"),
            "Type_of_Organisation": request.form.get("org_type"),
            "Geographical_Scope": request.form.get("geo"),
            "Description": request.form.get("description", "").strip(),
        }
        # Raw scores
        for q in POWER_QUESTIONS + INTEREST_QUESTIONS:
            row[q["id"]] = request.form.get(q["id"])

        # Derived scores
        p, i, avg, tot = compute_scores(row)
        row["Power_Score"] = round(p, 2)
        row["Interest_Score"] = round(i, 2)
        row["Combined_Average_Score"] = round(avg, 2)
        row["Combined_Total_Score"] = round(tot, 2)
        row["Strategic_Engagement_Quadrant"] = quadrant_from_avg(avg)

        RESULT_ROWS.append(row)

        # Next org
        return redirect(url_for("rate", idx=idx + 1))

    # GET
    return render_template(
        "rate.html",
        rater=CURRENT_RATER,
        idx=idx, total=len(ORGS), org=org,
        sector_options=SECTOR_OPTIONS,
        subject_options=SUBJECT_OPTIONS,
        type_options=TYPE_OPTIONS,
        geo_options=GEO_OPTIONS,
        power_questions=POWER_QUESTIONS,
        interest_questions=INTEREST_QUESTIONS,
        form_values={},
    )

@app.route("/done")
def done():
    return render_template("done.html", rater=CURRENT_RATER)

@app.route("/download")
def download():
    # Build DataFrame and stream as Excel
    if not RESULT_ROWS:
        return redirect(url_for("index"))

    # Order columns nicely
    base_cols = [
        "Organisation","Rater_Name","Sector","Subject_Area","Type_of_Organisation",
        "Geographical_Scope","Description"
    ]
    power_cols = [q["id"] for q in POWER_QUESTIONS]
    interest_cols = [q["id"] for q in INTEREST_QUESTIONS]
    derived_cols = [
        "Power_Score","Interest_Score","Combined_Average_Score",
        "Combined_Total_Score","Strategic_Engagement_Quadrant"
    ]

    # Some rows (skips) may miss columns → normalize
    df = pd.DataFrame(RESULT_ROWS)
    for c in base_cols + power_cols + interest_cols + derived_cols:
        if c not in df.columns:
            df[c] = ""

    df = df[base_cols + power_cols + interest_cols + derived_cols]

    # Stream as Excel
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Ratings")
        # Add an info sheet with the mapping
        mapping = pd.DataFrame({
            "Strategic Quadrant":[
                "Manage Closely","Keep Satisfied","Keep Informed","Monitor"
            ],
            "Score (Combined Avg)": ["8–10","6–<8","3–<6","<3"]
        })
        mapping.to_excel(writer, index=False, sheet_name="Quadrant Mapping")
    output.seek(0)

    ts = datetime.now().strftime("%Y%m%d_%H%M")
    return send_file(
        output,
        as_attachment=True,
        download_name=f"nescan_ratings_{ts}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

if __name__ == "__main__":
    app.run(debug=True)

