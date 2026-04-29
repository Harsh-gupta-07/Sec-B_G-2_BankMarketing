from __future__ import annotations

from collections import Counter, defaultdict
from pathlib import Path
import subprocess

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
ASSETS = REPORTS / "_project_report_assets"
DATA = pd.read_csv(ROOT / "data" / "processed" / "cleaned_dataset.csv")
SCREENSHOT_DIR = ROOT / "tableau" / "screenshots"
OUT = REPORTS / "project_report.pdf"

TEAM_MEMBERS = [
    "Harshvardhan Gupta",
    "Kartik Madaan",
    "Dev Tyagi",
    "Shitanshu Tiwari",
    "Satwik Mani Tripathi",
    "Siddharth Dangi",
]


def fmt_int(v: float) -> str:
    return f"{int(round(v)):,}"


def fmt_pct(v: float) -> str:
    return f"{v * 100:.2f}%"


def fmt_currency(v: float) -> str:
    return f"€{v:,.0f}"


def crosstab_table(df, col):
    rate = (df.groupby(col)["y"].mean() * 100).round(2)
    count = df.groupby(col).size()
    out = pd.DataFrame({"Category": rate.index, "Count": count.values, "Conversion Rate %": rate.values})
    return out.sort_values("Conversion Rate %", ascending=False)


def cramers_v(x, y):
    ct = pd.crosstab(x, y)
    obs = ct.to_numpy(dtype=float)
    n = obs.sum()
    row_sums = obs.sum(axis=1, keepdims=True)
    col_sums = obs.sum(axis=0, keepdims=True)
    expected = row_sums @ col_sums / n
    chi2 = ((obs - expected) ** 2 / expected).sum()
    r, c = obs.shape
    v = (chi2 / (n * (min(r - 1, c - 1)))) ** 0.5
    return chi2, v


def repo_commit_summary():
    out = subprocess.check_output(["git", "log", "--format=%an|%ad|%s", "--date=short"], cwd=ROOT).decode().splitlines()
    by_author = defaultdict(list)
    for line in out:
        author, date, msg = line.split("|", 2)
        by_author[author].append((date, msg))
    return by_author


def heading_style(styles, name, size, color):
    s = ParagraphStyle(
        name=name,
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=size,
        leading=size + 2,
        textColor=color,
        spaceAfter=8,
    )
    styles.add(s)


def body_style(styles, name, size=10):
    s = ParagraphStyle(
        name=name,
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=size,
        leading=size + 4,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=6,
    )
    styles.add(s)


def build_styles():
    styles = getSampleStyleSheet()
    heading_style(styles, "ReportTitle", 24, colors.HexColor("#0F172A"))
    heading_style(styles, "SectionHeading", 16, colors.HexColor("#0F172A"))
    heading_style(styles, "SubHeading", 11, colors.HexColor("#0E7490"))
    body_style(styles, "ReportBody", 10)
    body_style(styles, "SmallBody", 8.6)
    styles.add(
        ParagraphStyle(
            name="CoverTag",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=10,
            leading=12,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#0E7490"),
            spaceAfter=10,
        )
    )
    return styles


def p(text, style):
    return Paragraph(text, style)


def table_style(tbl, header_fill="0F172A", alt_fill="F8FAFC", font_size=8.4, header_font_size=8.2):
    style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(f"#{header_fill}")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, 0), header_font_size),
            ("FONTSIZE", (0, 1), (-1, -1), font_size),
            ("LEADING", (0, 0), (-1, -1), 10),
            ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9E2EC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 6),
            ("RIGHTPADDING", (0, 0), (-1, -1), 6),
            ("TOPPADDING", (0, 0), (-1, -1), 5),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ]
    )
    for row in range(1, len(tbl._cellvalues)):
        if row % 2 == 0:
            style.add("BACKGROUND", (0, row), (-1, row), colors.HexColor(f"#{alt_fill}"))
    tbl.setStyle(style)
    return tbl


def df_table(df, col_widths, font_size=8.4):
    data = [list(df.columns)] + df.astype(str).values.tolist()
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    table_style(tbl, font_size=font_size)
    return tbl


def cover_table():
    wrap_style = ParagraphStyle(
        name="CoverCell",
        parent=getSampleStyleSheet()["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#1E293B"),
    )
    facts = [
        ["Team ID", "Sec-B G-2"],
        ["Institute", "Newton School of Technology, Rishihood University"],
        ["Team members", Paragraph("Harshvardhan Gupta, Kartik Madaan, Dev Tyagi<br/>Shitanshu Tiwari, Satwik Mani Tripathi, Siddharth Dangi", wrap_style)],
        ["Sector", "Finance / Bank Marketing"],
        ["Dataset", "UCI Bank Marketing (bank-full.csv)"],
        ["Prepared on", "29 April 2026"],
    ]
    tbl = Table(facts, colWidths=[1.35 * inch, 5.9 * inch])
    tbl_style = TableStyle(
        [
            ("BACKGROUND", (0, 0), (-1, -1), colors.white),
            ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1E293B")),
            ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
            ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
            ("FONTSIZE", (0, 0), (-1, -1), 9.5),
            ("LEADING", (0, 0), (-1, -1), 11),
            ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D9E2EC")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ]
    )
    tbl.setStyle(tbl_style)
    return tbl


def metric_table():
    rows = [
        ["Total contacts", fmt_int(len(DATA))],
        ["Subscriptions", fmt_int(int(DATA["y"].sum()))],
        ["Subscription rate", fmt_pct(DATA["y"].mean())],
        ["Avg. campaign touches", f"{DATA['campaign'].mean():.2f}"],
    ]
    tbl = Table(rows, colWidths=[2.2 * inch, 1.6 * inch])
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.white),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#1E293B")),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 9.2),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9E2EC")),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    return tbl


def cover_story(styles):
    story = [
        Spacer(1, 0.3 * inch),
        p("DATA VISUALIZATION & ANALYTICS", styles["CoverTag"]),
        p("Bank Term Deposit Campaign Intelligence", styles["ReportTitle"]),
        p("Finance / Bank Marketing | Project Report", styles["CoverTag"]),
        Spacer(1, 0.12 * inch),
        p(
            "This report analyses the UCI Bank Marketing dataset to identify which clients are most likely to subscribe to a term deposit, how campaign design affects conversion, and how Tableau dashboards can support operational targeting decisions.",
            styles["ReportBody"],
        ),
        Spacer(1, 0.08 * inch),
        cover_table(),
        Spacer(1, 0.22 * inch),
        p("At a glance", styles["SubHeading"]),
        metric_table(),
        Spacer(1, 0.1 * inch),
        p("The report uses only repository-visible data and dashboard assets stored in this project workspace.", styles["SmallBody"]),
    ]
    return story


def section_intro(title, subtitle=None, styles=None):
    story = [p(title, styles["SectionHeading"])]
    if subtitle:
        story.append(p(subtitle, styles["SmallBody"]))
    story.append(Spacer(1, 0.06 * inch))
    return story


def img(path, width=6.7 * inch):
    im = Image(str(path), width=width, height=width * 0.64)
    im.hAlign = "CENTER"
    return im


def build_pdf():
    styles = build_styles()
    by_author = repo_commit_summary()
    story = []

    # Cover
    story.extend(cover_story(styles))
    story.append(PageBreak())

    story += section_intro("1. Executive Summary", "One-page summary for decision makers", styles)
    story += [
        p(f"The dataset contains {fmt_int(len(DATA))} client contacts and {fmt_int(int(DATA['y'].sum()))} term-deposit subscriptions, producing an overall subscription rate of {fmt_pct(DATA['y'].mean())}.", styles["ReportBody"]),
        p(f"The strongest business signal is prior campaign history: previously contacted clients convert at {fmt_pct(DATA.groupby('previously_contacted')['y'].mean().loc[1])}, compared with {fmt_pct(DATA.groupby('previously_contacted')['y'].mean().loc[0])} for clients without prior contact.", styles["ReportBody"]),
        p(f"Students ({DATA.groupby('job')['y'].mean().sort_values(ascending=False).loc['student']*100:.2f}%), retirees ({DATA.groupby('job')['y'].mean().sort_values(ascending=False).loc['retired']*100:.2f}%) and older clients aged 65+ ({DATA.groupby('age_group')['y'].mean().sort_values(ascending=False).loc['65+']*100:.2f}%) respond well.", styles["ReportBody"]),
        p(f"Repeated calling is inefficient: conversion falls from {DATA.groupby('campaign_band')['y'].mean().loc['1']*100:.2f}% on first touch to {DATA.groupby('campaign_band')['y'].mean().loc['6+']*100:.2f}% for six or more contacts.", styles["ReportBody"]),
        p("The report therefore recommends a warm-lead-first call queue, cellular-first outreach, a cap on repeated touches after three attempts, and a stronger focus on high-response months such as March, September, October, and December.", styles["ReportBody"]),
    ]
    for txt in [
        "Decision rule: prioritize warm leads before cold prospects.",
        "Campaign design rule: keep contact attempts short and targeted.",
        "Channel rule: use cellular-first routing for higher conversion.",
        "Timing rule: schedule the heaviest pushes in the best-performing months.",
    ]:
        story.append(Paragraph(f"• {txt}", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("2. Sector Context and Problem Statement", "Finance-sector lead generation is a prioritization problem, not a volume problem.", styles)
    for para in [
        "Retail banks run large outbound campaigns to sell term deposits. The challenge is that calling everyone wastes capacity, depresses conversion efficiency, and creates noise for low-propensity clients.",
        "The decision-maker in this project is a campaign or CRM manager who needs a simple rule set: who to call first, how often to call, and which channels should be used for the best return on effort.",
    ]:
        story.append(p(para, styles["ReportBody"]))
    for item in [
        "Business question: which clients should be contacted first to maximize subscription probability?",
        "Operational question: how can the bank reduce wasted calls while protecting customer experience?",
        "Outcome: a Tableau dashboard and Python analysis that support segmentation-driven outreach.",
    ]:
        story.append(Paragraph(f"• {item}", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("3. Data Description", "Source, structure, size, and known limitations", styles)
    story.append(p("The analysis uses the UCI Machine Learning Repository's Bank Marketing dataset (`bank-full.csv`). The raw file is semicolon-delimited and has one row per client contact record.", styles["ReportBody"]))
    data_desc = pd.DataFrame(
        [
            ["Source", "UCI Machine Learning Repository"],
            ["Raw file", "bank-full.csv"],
            ["Rows", fmt_int(len(DATA))],
            ["Raw columns", "17"],
            ["Cleaned columns", str(len(DATA.columns))],
            ["Unit of analysis", "One client contact record"],
            ["Target variable", "y (term-deposit subscription: 0/1)"],
        ],
        columns=["Attribute", "Details"],
    )
    story.append(df_table(data_desc, [1.8 * inch, 4.8 * inch], font_size=8.8))
    story.append(Spacer(1, 0.08 * inch))
    story.append(p("Key limitations: the file uses the label `unknown` in several categorical fields rather than missing values; `duration` is informative for descriptive analysis but leaks information in a predictive setting; and the data represents one historical campaign period, so seasonality should be validated before operational use.", styles["ReportBody"]))
    story.append(p("Important derived fields introduced for Tableau and analysis include `age_group`, `balance_band`, `campaign_band`, `previously_contacted`, and `duration_minutes`.", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("4. Data Cleaning and ETL Methodology", "Python pipeline from raw file to Tableau-ready dataset", styles)
    story.append(p("The ETL pipeline in `scripts/etl_pipeline.py` standardizes the raw file, validates required columns, preserves semantic labels, and adds business-friendly derived columns.", styles["ReportBody"]))
    for item in [
        "Normalize headers to lowercase snake_case and standardize string values.",
        "Validate expected raw columns before further processing.",
        "Drop duplicates and preserve explicit `unknown` labels rather than silently imputing them.",
        "Create `previously_contacted` from `pdays != -1`.",
        "Bucket `age` into life-stage bands and `balance` into financial bands.",
        "Create `campaign_band` to distinguish first-touch, low-touch, and high-touch campaigns.",
        "Export the cleaned dataset and Tableau-ready subset with 22 columns.",
    ]:
        story.append(Paragraph(f"• {item}", styles["ReportBody"]))
    etl_steps = pd.DataFrame(
        [
            ["Input", "data/raw/bank-full.csv", "Raw UCI source"],
            ["Standardize", "Lowercase / snake_case / strip spaces", "Consistent joins and filtering"],
            ["Derive", "age_group, balance_band, campaign_band, etc.", "Business segmentation"],
            ["Validate", "Required columns and yes/no target values", "Data quality control"],
            ["Output", "cleaned_dataset.csv / Tableau-ready CSV", "Analysis and dashboard consumption"],
        ],
        columns=["Stage", "What happens", "Why it matters"],
    )
    story.append(df_table(etl_steps, [1.0 * inch, 3.0 * inch, 2.0 * inch], font_size=8.2))
    story.append(PageBreak())

    story += section_intro("5. KPI and Metric Framework", "Metrics used to judge conversion, targeting quality, and financial context", styles)
    kpis = pd.DataFrame(
        [
            ["Subscription rate", fmt_pct(DATA["y"].mean()), "Subscribers / total contacts", "North-star conversion metric"],
            ["Total subscriptions", fmt_int(int(DATA["y"].sum())), "Count of y = 1 records", "Absolute campaign success"],
            ["Avg. contacts / client", f"{DATA['campaign'].mean():.2f}", "Mean campaign touch count", "Effort intensity"],
            ["Previously contacted share", fmt_pct(DATA["previously_contacted"].mean()), "previously_contacted = 1", "Warm-lead prevalence"],
            ["Avg. yearly balance", fmt_currency(DATA["balance"].mean()), "Mean balance across all clients", "Financial capability"],
            ["Default rate", fmt_pct((DATA["default"] == 'yes').mean()), "Clients with default = yes", "Credit risk context"],
            ["Loan rate", fmt_pct((DATA["loan"] == 'yes').mean()), "Clients with loan = yes", "Debt burden context"],
        ],
        columns=["KPI", "Current value", "Definition", "Business meaning"],
    )
    story.append(df_table(kpis, [1.45 * inch, 1.05 * inch, 2.05 * inch, 1.55 * inch], font_size=8.1))
    story.append(p("For campaign management, the most useful pairings are subscription rate versus contact intensity, and subscription rate versus prior contact history. These pairs reveal where the bank is gaining conversion and where it is wasting touches.", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("6. Exploratory Data Analysis", "Visual patterns that shape segmentation and targeting", styles)
    story.append(p("The EDA shows that subscription likelihood is not uniform. Certain customer groups respond materially better, and some outreach patterns are clearly inefficient.", styles["ReportBody"]))
    story.append(img(ASSETS / "job_conversion.png"))
    story.append(p("Students and retirees lead the conversion table, while blue-collar, entrepreneur, and housemaid segments sit near the bottom. That makes occupation a practical prioritization filter rather than a descriptive label.", styles["ReportBody"]))
    story.append(img(ASSETS / "month_conversion.png"))
    story.append(p("Campaign response spikes in March, December, September, and October. Timing matters, so calendar windows should be treated as a campaign lever rather than an afterthought.", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("7. Statistical Analysis Results", "Association strength and decision relevance", styles)
    stats_rows = []
    for col in ["poutcome", "previously_contacted", "contact", "age_group", "job", "balance_band", "campaign_band", "education", "marital"]:
        chi2, v = cramers_v(DATA[col], DATA["y"])
        interp = ""
        if col == "poutcome":
            interp = "Strongest relationship; prior success is highly informative."
        elif col == "previously_contacted":
            interp = "Warm leads materially outperform cold leads."
        elif col == "contact":
            interp = "Cellular contact clearly outperforms telephone."
        elif col == "age_group":
            interp = "Older bands convert materially better."
        elif col == "job":
            interp = "Occupation is a useful segmentation feature."
        elif col == "balance_band":
            interp = "Higher balances align with better conversion."
        elif col == "campaign_band":
            interp = "Conversion decays as touches increase."
        elif col == "education":
            interp = "Education matters, but less than contact history."
        elif col == "marital":
            interp = "Useful as a secondary segmentation feature."
        stats_rows.append([col, f"{chi2:,.1f}", f"{v:.3f}", interp])
    story.append(df_table(pd.DataFrame(stats_rows, columns=["Driver", "Chi-square", "Cramér's V", "Interpretation"]), [1.6 * inch, 1.0 * inch, 1.0 * inch, 2.7 * inch], font_size=7.9))
    story.append(p(f"Subscribers have a higher average balance ({fmt_currency(DATA.loc[DATA['y'] == 1, 'balance'].mean())}) than non-subscribers ({fmt_currency(DATA.loc[DATA['y'] == 0, 'balance'].mean())}), and they receive fewer campaign touches on average ({DATA.loc[DATA['y'] == 1, 'campaign'].mean():.2f} versus {DATA.loc[DATA['y'] == 0, 'campaign'].mean():.2f}).", styles["ReportBody"]))
    story.append(p(f"Previously contacted clients convert at {DATA.groupby('previously_contacted')['y'].mean().loc[1]*100:.2f}% versus {DATA.groupby('previously_contacted')['y'].mean().loc[0]*100:.2f}% for clients with no prior contact, a lift of roughly {((DATA.groupby('previously_contacted')['y'].mean().loc[1] / DATA.groupby('previously_contacted')['y'].mean().loc[0]) - 1) * 100:.1f}%.", styles["ReportBody"]))
    story.append(PageBreak())

    for title, screenshot, bullets, subtitle in [
        ("8. Dashboard Design", SCREENSHOT_DIR / "executive_view.png", ["Shows headline KPIs, high-level segments, and balance distribution.", "Best for leadership reporting and fast portfolio review.", "Main filters: job, age group, education, marital status, balance band."], "Screenshots and what each view is meant to answer"),
        ("8. Dashboard Design (continued)", SCREENSHOT_DIR / "campaign strategy view.png", ["Focuses on contact volume, channel performance, prior-contact lift, and campaign touch effects.", "Best for CRM managers deciding how to route the next batch of leads.", "Main filters: number of contacts, prior contact, channel, campaign history."], "Operational lens one"),
        ("8. Dashboard Design (continued)", SCREENSHOT_DIR / "risk_balance_view.png", ["Brings credit risk context into the story through default, loan, and balance bands.", "Best for balancing conversion uplift against customer quality and risk exposure.", "Main filters: age group, balance band, loan, default, and outcome."], "Operational lens two"),
    ]:
        story += section_intro(title, subtitle, styles)
        story.append(img(screenshot))
        for item in bullets:
            story.append(Paragraph(f"• {item}", styles["ReportBody"]))
        story.append(PageBreak())

    story += section_intro("9. Key Insights and Recommendations", "Decision language, not chart captions", styles)
    insights = pd.DataFrame(
        [
            ["1", "Prioritize students and retirees first.", "They deliver the highest conversion and deserve the first call queue."],
            ["2", "Treat 65+ and 18-24 as high-response age bands.", "Age is a practical targeting dimension, not just a descriptive one."],
            ["3", "Use prior contact history as a primary routing rule.", "Warm leads convert far better than cold leads."],
            ["4", "Cap repeated outreach after three touches.", "Conversion falls sharply after touch count increases."],
            ["5", "Prefer cellular-first contact plans.", "Cellular outperforms telephone and unknown channels are weak."],
            ["6", "Push campaigns in March, September, October, and December.", "Response is materially stronger in these windows."],
            ["7", "Give medium and high balance clients higher priority.", "Higher balance bands outperform low and negative bands."],
            ["8", "Use prior success as a premium routing signal.", "Successful previous outcomes dominate all other signals."],
            ["9", "De-emphasize blue-collar and services segments for premium outreach.", "These groups sit near the bottom of the response table."],
            ["10", "Keep unknown contact records out of prime call queues.", "Missing channel context lowers conversion and should be cleaned."],
        ],
        columns=["#", "Insight", "Business implication"],
    )
    story.append(df_table(insights, [0.35 * inch, 2.75 * inch, 3.2 * inch], font_size=7.9))
    story.append(Spacer(1, 0.06 * inch))
    story.append(p("Recommended actions:", styles["SubHeading"]))
    recs = pd.DataFrame(
        [
            ["Prioritize warm leads", "Route previously contacted leads and prior-success records first.", "Highest immediate conversion gain"],
            ["Apply touch caps", "Stop repeating low-yield calls after three touches.", "Lower wasted call volume"],
            ["Use cellular-first routing", "Make cellular the default channel where possible.", "Stronger response rate"],
            ["Seasonal campaign bursts", "Concentrate spend in peak months.", "Better calendar efficiency"],
        ],
        columns=["Recommendation", "What to do", "Expected impact"],
    )
    story.append(df_table(recs, [1.35 * inch, 3.1 * inch, 1.45 * inch], font_size=7.9))
    story.append(PageBreak())

    story += section_intro("10. Impact Estimation, Limitations, and Future Scope", "What this work can realistically change", styles)
    story.append(p(f"If 10,000 contacts are routed from the cold baseline to previously contacted leads, the expected incremental lift is about {int(round(10_000 * (DATA.groupby('previously_contacted')['y'].mean().loc[1] - DATA.groupby('previously_contacted')['y'].mean().loc[0]))):,} extra subscriptions.", styles["ReportBody"]))
    story.append(p(f"Using a conservative illustrative contribution of INR 50 per incremental subscription, the same shift implies roughly INR {round(10_000 * (DATA.groupby('previously_contacted')['y'].mean().loc[1] - DATA.groupby('previously_contacted')['y'].mean().loc[0]) * 50):,.0f} in value per 10,000 contacted leads.", styles["ReportBody"]))
    for item in [
        "Limitation 1: no call-cost or offer-cost data, so ROI is estimated rather than exact.",
        "Limitation 2: `duration` is useful for analysis but leaks outcome information and should not be used as a pre-call predictor.",
        "Limitation 3: the dataset captures one campaign period, so seasonality should be revalidated before deployment.",
        "Limitation 4: unknown categories are explicit labels, not absent rows, which is useful for analysis but imperfect for modeling.",
    ]:
        story.append(Paragraph(f"• {item}", styles["ReportBody"]))
    story.append(p("Future scope should include a cost-per-touch model, a propensity scoring layer, and A/B-tested routing rules that compare warm-lead targeting against broader outreach mixes.", styles["ReportBody"]))
    story.append(PageBreak())

    story += section_intro("11. Contribution Matrix", None, styles)
    commit_counts = Counter({author: len(commits) for author, commits in by_author.items()})
    contrib = pd.DataFrame(
        [
            ["Harshvardhan Gupta", f"{commit_counts.get('Harshvardhan Gupta', 0)} commits", "Raw data, cleaning, EDA, statistical analysis, workbook updates"],
            ["Kartik Madaan", f"{commit_counts.get('madaankartik', 0)} commits", "Statistical analysis, Tableau dashboard, report writing, PPT"],
            ["Dev Tyagi", "4 commits", "Tableau dashboard"],
            ["Shitanshu Tiwari", "2 commits", "EDA and analysis"],
            ["Satwik Mani Tripathi", "No commit", "N/A"],
            ["Siddharth Dangi", "No commit", "N/A"],
        ],
        columns=["Team member", "Git history status", "Repository-visible contribution"],
    )
    story.append(df_table(contrib, [1.55 * inch, 1.35 * inch, 3.0 * inch], font_size=7.4))

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D9E2EC"))
        canvas.setLineWidth(0.5)
        canvas.line(doc.leftMargin, 0.55 * inch, A4[0] - doc.rightMargin, 0.55 * inch)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#64748B"))
        canvas.drawString(doc.leftMargin, 0.35 * inch, "Bank Marketing DVA Capstone")
        canvas.drawRightString(A4[0] - doc.rightMargin, 0.35 * inch, f"Page {canvas.getPageNumber()}")
        canvas.restoreState()

    doc = SimpleDocTemplate(
        str(OUT),
        pagesize=A4,
        leftMargin=0.7 * inch,
        rightMargin=0.7 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.75 * inch,
    )
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    return OUT


if __name__ == "__main__":
    print(build_pdf())
