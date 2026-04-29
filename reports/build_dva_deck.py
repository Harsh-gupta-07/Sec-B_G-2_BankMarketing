from __future__ import annotations

import math
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Paragraph, Table, TableStyle
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader


ROOT = Path(__file__).resolve().parents[1]
REPORTS = ROOT / "reports"
ASSETS = REPORTS / "_deck_assets"
ASSETS.mkdir(parents=True, exist_ok=True)

DATA = pd.read_csv(ROOT / "data" / "processed" / "cleaned_dataset.csv")

SLIDE_W, SLIDE_H = landscape((13.333 * inch, 7.5 * inch))

BG = colors.HexColor("#F7FAFC")
INK = colors.HexColor("#0F172A")
MUTED = colors.HexColor("#64748B")
NAVY = colors.HexColor("#08111F")
BLUE = colors.HexColor("#2F6DAE")
TEAL = colors.HexColor("#14B8A6")
AMBER = colors.HexColor("#E7A33E")
RED = colors.HexColor("#D95C5C")
CARD = colors.white
LINE = colors.HexColor("#D7E2EC")
SOFT = colors.HexColor("#E8EFF6")

styles = getSampleStyleSheet()
styles.add(
    ParagraphStyle(
        name="SlideTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=24,
        leading=28,
        textColor=INK,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="SlideSubtitle",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        textColor=MUTED,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="Body",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=14,
        textColor=INK,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="Small",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        leading=10,
        textColor=MUTED,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="CardTitle",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=12,
        leading=14,
        textColor=INK,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="MetricValue",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=24,
        leading=26,
        textColor=BLUE,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="CoverCardNote",
        parent=styles["Normal"],
        fontName="Helvetica",
        fontSize=10.5,
        leading=13,
        textColor=MUTED,
        alignment=TA_LEFT,
    )
)
styles.add(
    ParagraphStyle(
        name="CoverCardFooter",
        parent=styles["Normal"],
        fontName="Helvetica-Bold",
        fontSize=7.6,
        leading=9,
        textColor=colors.HexColor("#B6C8DA"),
        alignment=TA_LEFT,
    )
)


def fmt_pct(v: float) -> str:
    return f"{v * 100:.1f}%"


def fmt_eur(v: float) -> str:
    return f"₹{v:,.0f}"


def wrap_paragraph(text: str, style_name: str, width: float):
    para = Paragraph(text, styles[style_name])
    _, h = para.wrap(width, 1000)
    return para, h


def draw_paragraph(c: canvas.Canvas, text: str, style_name: str, x: float, top: float, width: float):
    para, h = wrap_paragraph(text, style_name, width)
    para.drawOn(c, x, top - h)
    return h


def draw_card(c, x, y, w, h, radius=18, fill=CARD, stroke=LINE, stroke_width=1):
    c.setFillColor(fill)
    c.setStrokeColor(stroke)
    c.setLineWidth(stroke_width)
    c.roundRect(x, y, w, h, radius, stroke=1, fill=1)


def draw_chip(c, x, y, text, fill, ink=colors.white, pad_x=10, pad_y=6, r=10, font_size=8.5):
    c.setFont("Helvetica-Bold", font_size)
    text_w = stringWidth(text, "Helvetica-Bold", font_size)
    w = text_w + pad_x * 2
    h = font_size + pad_y * 2
    c.setFillColor(fill)
    c.setStrokeColor(fill)
    c.roundRect(x, y, w, h, r, stroke=0, fill=1)
    c.setFillColor(ink)
    c.drawString(x + pad_x, y + pad_y - 0.5, text)
    return w, h


def draw_header(c, slide_no: int, title: str, subtitle: str):
    c.setFillColor(BG)
    c.rect(0, 0, SLIDE_W, SLIDE_H, fill=1, stroke=0)
    draw_chip(c, 34, SLIDE_H - 38, "BANK MARKETING DVA", NAVY, font_size=7.5)
    draw_chip(c, SLIDE_W - 78, SLIDE_H - 38, f"{slide_no:02d}", BLUE, font_size=8.5, pad_x=7)
    draw_paragraph(c, title, "SlideTitle", 34, SLIDE_H - 68, SLIDE_W - 170)
    draw_paragraph(c, subtitle, "SlideSubtitle", 34, SLIDE_H - 100, SLIDE_W - 170)
    c.setStrokeColor(TEAL)
    c.setLineWidth(2)
    c.line(34, SLIDE_H - 116, SLIDE_W - 34, SLIDE_H - 116)


def draw_footer(c, source: str, extra: str = ""):
    c.setStrokeColor(LINE)
    c.setLineWidth(1)
    c.line(34, 26, SLIDE_W - 34, 26)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 7.5)
    c.drawString(34, 14, source)
    if extra:
        c.drawRightString(SLIDE_W - 34, 14, extra)


def draw_bullet_list(c, bullets, x, top, width, bullet_color=BLUE, gap=10, font="Body"):
    y = top
    for item in bullets:
        c.setFillColor(bullet_color)
        c.circle(x + 5, y - 5, 2.4, stroke=0, fill=1)
        h = draw_paragraph(c, item, font, x + 14, y, width - 14)
        y -= h + gap
    return top - y


def draw_metric_box(c, x, y, w, h, title, value, body, accent=BLUE, value_size=23):
    draw_card(c, x, y, w, h, radius=16)
    c.setFillColor(accent)
    c.rect(x, y + h - 7, w, 7, stroke=0, fill=1)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(x + 16, y + h - 24, title)
    c.setFillColor(accent)
    c.setFont("Helvetica-Bold", value_size)
    c.drawString(x + 16, y + h - 60, value)
    draw_paragraph(c, body, "Small", x + 16, y + 26, w - 32)


def draw_section_card(c, x, y, w, h, title, body, accent=BLUE):
    draw_card(c, x, y, w, h, radius=16)
    c.setFillColor(accent)
    c.rect(x, y + h - 7, w, 7, stroke=0, fill=1)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(x + 16, y + h - 24, title)
    draw_paragraph(c, body, "Body", x + 16, y + h - 46, w - 32)


def _mix_color(a, b, t):
    return colors.Color(
        a.red + (b.red - a.red) * t,
        a.green + (b.green - a.green) * t,
        a.blue + (b.blue - a.blue) * t,
    )


def _draw_chart_frame(c, x, y, w, h, title, subtitle=None):
    draw_card(c, x, y, w, h, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12.5)
    c.drawString(x + 16, y + h - 24, title)
    if subtitle:
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8.8)
        c.drawString(x + 16, y + h - 38, subtitle)
    c.setStrokeColor(SOFT)
    c.setLineWidth(1)
    c.line(x + 16, y + h - 44, x + w - 16, y + h - 44)


def draw_bar_chart(
    c,
    x,
    y,
    w,
    h,
    title,
    labels,
    values,
    colors_list,
    subtitle=None,
    y_max=None,
    percent=True,
):
    _draw_chart_frame(c, x, y, w, h, title, subtitle)
    plot_left, plot_right = 42, 16
    plot_bottom, plot_top = 34, 58
    pw = w - plot_left - plot_right
    ph = h - plot_bottom - plot_top
    px = x + plot_left
    py = y + plot_bottom
    max_v = y_max or max(values) * 1.18 if values else 1
    max_v = max(max_v, 0.01)
    c.setStrokeColor(LINE)
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 7.5)
    for t in range(5):
        frac = t / 4
        yy = py + ph * frac
        c.setStrokeColor(colors.HexColor("#E6EDF5"))
        c.line(px, yy, px + pw, yy)
        tick = max_v * frac
        c.drawRightString(px - 6, yy - 2, f"{tick*100:.0f}%" if percent else f"{tick:,.0f}")
    n = len(labels)
    bar_w = pw / max(n, 1) * 0.62
    slot = pw / max(n, 1)
    for i, (lab, val, col) in enumerate(zip(labels, values, colors_list)):
        bar_h = ph * (val / max_v)
        bx = px + slot * i + (slot - bar_w) / 2
        by = py
        c.setFillColor(col)
        c.roundRect(bx, by, bar_w, max(bar_h, 1.5), 6, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 8.5)
        c.drawCentredString(bx + bar_w / 2, by + bar_h + 7, f"{val*100:.1f}%" if percent else f"{val:,.0f}")
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8.2)
        if len(lab) > 10:
            c.drawCentredString(bx + bar_w / 2, y + 12, lab[:10] + "...")
        else:
            c.drawCentredString(bx + bar_w / 2, y + 12, lab)


def draw_heatmap_chart(c, x, y, w, h, title, row_labels, col_labels, matrix, subtitle=None):
    _draw_chart_frame(c, x, y, w, h, title, subtitle)
    plot_left, plot_right = 56, 18
    plot_bottom, plot_top = 32, 54
    pw = w - plot_left - plot_right
    ph = h - plot_bottom - plot_top
    px = x + plot_left
    py = y + plot_bottom
    rows, cols = len(row_labels), len(col_labels)
    cell_w = pw / cols
    cell_h = ph / rows
    flat = [v for row in matrix for v in row if v is not None]
    mn, mx = min(flat), max(flat)
    for j, lab in enumerate(col_labels):
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8.2)
        c.drawCentredString(px + cell_w * (j + 0.5), py + ph + 8, lab)
    for i, rlab in enumerate(row_labels):
        c.setFillColor(MUTED)
        c.setFont("Helvetica", 8.2)
        c.drawRightString(px - 6, py + ph - cell_h * (i + 0.62), rlab)
    for i, row in enumerate(matrix):
        for j, val in enumerate(row):
            frac = 0 if mx == mn else (val - mn) / (mx - mn)
            fill = _mix_color(colors.HexColor("#EAF3FB"), TEAL, frac)
            cell_x = px + cell_w * j
            cell_y = py + ph - cell_h * (i + 1)
            c.setFillColor(fill)
            c.setStrokeColor(colors.white)
            c.roundRect(cell_x + 1, cell_y + 1, cell_w - 2, cell_h - 2, 5, stroke=1, fill=1)
            c.setFillColor(colors.white if frac > 0.55 else INK)
            c.setFont("Helvetica-Bold", 8.4)
            c.drawCentredString(cell_x + cell_w / 2, cell_y + cell_h / 2 - 3, f"{val*100:.1f}%")
    # mini legend
    lx, ly = x + w - 132, y + 4
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 7.5)
    c.drawString(lx, ly + 16, "Low")
    c.drawString(lx + 84, ly + 16, "High")
    for k in range(60):
        frac = k / 59
        fill = _mix_color(colors.HexColor("#EAF3FB"), TEAL, frac)
        c.setFillColor(fill)
        c.rect(lx + 18 + k * 1.1, ly + 8, 1.1, 8, stroke=0, fill=1)


def create_assets():
    return {
        "dash1": ROOT / "tableau" / "screenshots" / "Dashboard 1.png",
        "dash2": ROOT / "tableau" / "screenshots" / "Dashboard 2.png",
        "dash3": ROOT / "tableau" / "screenshots" / "Dashboard 3.png",
        "resume": ROOT / "DVA-oriented-Resume" / "DVA_Oriented_Resume.pdf",
    }


def draw_cover(c, assets):
    c.setFillColor(NAVY)
    c.rect(0, 0, SLIDE_W, SLIDE_H, fill=1, stroke=0)
    # abstract geometry
    c.setStrokeColor(colors.HexColor("#17304F"))
    c.setLineWidth(2)
    for i in range(6):
        c.circle(770 + i * 12, 320 + i * 8, 125 + i * 18, stroke=1, fill=0)
    c.setFillColor(colors.HexColor("#10253F"))
    c.roundRect(610, 108, 300, 318, 30, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#0E1B31"))
    c.roundRect(626, 124, 268, 286, 24, stroke=0, fill=1)
    c.setFillColor(TEAL)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(52, 466, "FINANCE / BANK MARKETING")
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 28)
    c.drawString(52, 408, "Bank Term Deposit")
    c.drawString(52, 372, "Campaign Intelligence")
    c.setFillColor(colors.HexColor("#B6C8DA"))
    c.setFont("Helvetica", 12.5)
    c.drawString(52, 338, "A Tableau and Python capstone on how customer segments,")
    c.drawString(52, 318, "contact history, and campaign design shape subscription lift.")
    c.setStrokeColor(colors.HexColor("#21405F"))
    c.line(52, 292, 390, 292)
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 44)
    c.drawString(52, 228, "11.70%")
    c.setFillColor(colors.HexColor("#B6C8DA"))
    c.setFont("Helvetica", 12)
    c.drawString(52, 206, "Baseline term-deposit subscription rate")
    c.setFillColor(colors.white)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(52, 156, "Team ID")
    c.setFont("Helvetica", 14)
    c.drawString(140, 156, "Sec-B G-2")
    c.setFont("Helvetica-Bold", 16)
    c.drawString(52, 132, "Members")
    c.setFont("Helvetica", 14)
    c.drawString(140, 132, "Harshvardhan Gupta, Kartik Madaan, Dev Tyagi")
    c.drawString(140, 116, "Shitanshu Tiwari, Satwik Mani Tripathi, Siddharth Dangi")
    # stat card
    c.setFillColor(colors.white)
    c.roundRect(632, 150, 268, 246, 22, stroke=0, fill=1)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(654, 360, "Project signal")
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 54)
    c.drawString(654, 284, "5,289")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 12)
    c.drawString(654, 262, "subscriptions out of 45,211 contacts")
    c.setFillColor(TEAL)
    c.setFont("Helvetica-Bold", 20)
    c.drawString(654, 220, "23.07%")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 8)
    c.drawString(654, 202, "previously contacted clients convert nearly 2x better")
    c.setFillColor(colors.HexColor("#B6C8DA"))
    c.setFont("Helvetica-Bold", 6.5)
    c.drawString(654, 176, "UCI Bank Marketing dataset | Finance sector | Executive analytics deck")


def slide_2(c):
    draw_header(
        c,
        2,
        "Context and problem statement",
        "The bank needs to sell term deposits, but the campaign budget is wasted if every lead is treated the same way.",
    )
    left_x, top_y = 34, 392
    panel_w, panel_h = 430, 190
    draw_section_card(
        c,
        left_x,
        220,
        panel_w,
        panel_h,
        "Business context",
        "Retail banking campaigns are high-volume and low-conversion. In this dataset, the bank contacted 45,211 clients and only 11.7% subscribed. That means the core challenge is not reaching more people, but identifying which people deserve the next call.",
        accent=BLUE,
    )
    draw_section_card(
        c,
        496,
        220,
        430,
        190,
        "Decision to unlock",
        "Which customers should be contacted first, which channels should be prioritized, and when should the bank stop repeating low-yield calls?",
        accent=TEAL,
    )
    draw_section_card(
        c,
        34,
        74,
        892,
        120,
        "What the analysis must answer",
        "1) Which demographic and financial segments respond best? 2) Does prior contact history matter more than raw volume? 3) How can the bank lift subscriptions without inflating call cost or annoying low-propensity clients?",
        accent=AMBER,
    )
    draw_footer(c, "Source: UCI Bank Marketing dataset cleaned for Tableau and statistical analysis.", "Business question: improve term-deposit conversion")


def slide_3(c):
    draw_header(
        c,
        3,
        "Data engineering",
        "The raw UCI file was cleaned into a Tableau-ready dataset with derived segments for targeting and reporting.",
    )
    # pipeline row
    stages = [
        ("Source", "UCI Bank Marketing CSV<br/>45,211 rows | 17 raw fields"),
        ("Clean", "Standardized labels, preserved unknowns, converted yes/no response to binary y"),
        ("Derive", "age_group, balance_band, campaign_band, previously_contacted, duration_minutes"),
        ("Publish", "Processed dataset with 22 columns for analysis and Tableau"),
    ]
    sx = 34
    sy = 314
    sw = 206
    sh = 108
    for i, (t, b) in enumerate(stages):
        x = sx + i * (sw + 14)
        draw_card(c, x, sy, sw, sh, radius=16)
        c.setFillColor(BLUE if i in (0, 3) else TEAL)
        c.circle(x + 22, sy + sh - 22, 8, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x + 38, sy + sh - 26, t)
        draw_paragraph(c, b, "Body", x + 18, sy + 46, sw - 36)
    # left cards
    draw_section_card(
        c,
        34,
        140,
        290,
        150,
        "Source and size",
        f"Raw source: UCI Machine Learning Repository.<br/>Rows: {len(DATA):,}.<br/>Columns after cleaning: {len(DATA.columns):,}.<br/>Granularity: one row per client contact record.",
        accent=BLUE,
    )
    draw_section_card(
        c,
        340,
        140,
        290,
        150,
        "Cleaning steps",
        "Stripped whitespace, kept unknown categories explicit, re-coded response into binary y, derived age and balance bands, and created a prior-contact flag so Tableau filters could stay business-readable.",
        accent=TEAL,
    )
    # data dictionary table
    draw_card(c, 646, 96, 280, 194, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(662, 272, "Data dictionary snapshot")
    table_data = [
        ["Field", "Meaning", "Role"],
        ["job", "Occupation category", "Segmentation"],
        ["balance", "Avg. yearly balance", "Financial context"],
        ["previous", "Prior contacts count", "Campaign history"],
        ["poutcome", "Previous outcome", "Root cause"],
        ["y", "Subscribed? (0/1)", "North-star KPI"],
    ]
    tbl = Table(table_data, colWidths=[60, 118, 58], rowHeights=[20] * len(table_data))
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF2FB")),
                ("TEXTCOLOR", (0, 0), (-1, 0), INK),
                ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#1F2937")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("GRID", (0, 0), (-1, -1), 0.4, LINE),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FBFD")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 6),
                ("RIGHTPADDING", (0, 0), (-1, -1), 4),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    tbl.wrapOn(c, 280, 180)
    tbl.drawOn(c, 662, 126)
    draw_footer(c, "Source: `data/raw/bank-full.csv` -> `data/processed/cleaned_dataset.csv`.", "Derived for Tableau and EDA")


def slide_4(c):
    draw_header(
        c,
        4,
        "KPI and metrics framework",
        "The deck tracks outcome, targeting, and financial context metrics so the bank can judge both performance and risk.",
    )
    # three framework cards
    cards = [
        (
            34,
            204,
            280,
            200,
            "Outcome KPIs",
            [
                ("Subscription rate", "11.70% of clients subscribed."),
                ("Total subscriptions", "5,289 wins from 45,211 contacts."),
            ],
            BLUE,
        ),
        (
            340,
            204,
            280,
            200,
            "Targeting KPIs",
            [
                ("Avg contacts / client", f"{DATA['campaign'].mean():.2f} touches per client."),
                ("Previously contacted rate", "18.3% had prior campaign exposure."),
                ("Best channel conversion", "Cellular outperforms telephone."),
            ],
            TEAL,
        ),
        (
            646,
            204,
            280,
            200,
            "Financial context KPIs",
            [
                ("Avg yearly balance", fmt_eur(DATA['balance'].mean()) + " overall."),
                ("Default rate", f"{DATA['default'].eq('yes').mean()*100:.2f}% of clients."),
                ("Loan rate", f"{DATA['loan'].eq('yes').mean()*100:.2f}% of clients."),
            ],
            AMBER,
        ),
    ]
    for x, y, w, h, title, rows, accent in cards:
        draw_card(c, x, y, w, h, radius=16)
        c.setFillColor(accent)
        c.rect(x, y + h - 7, w, 7, stroke=0, fill=1)
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 13)
        c.drawString(x + 16, y + h - 26, title)
        yy = y + h - 54
        for label, desc in rows:
            c.setFillColor(accent)
            c.circle(x + 18, yy - 4, 2.8, stroke=0, fill=1)
            c.setFillColor(INK)
            c.setFont("Helvetica-Bold", 9.2)
            c.drawString(x + 28, yy - 8, label)
            c.setFillColor(MUTED)
            c.setFont("Helvetica", 8.2)
            c.drawString(x + 28, yy - 20, desc)
            yy -= 34
    draw_section_card(
        c,
        34,
        62,
        892,
        104,
        "Framework logic",
        "Outcome KPIs tell us whether the campaign worked. Targeting KPIs show how the bank is spending attention. Financial context KPIs explain which customers are safe, risky, or simply more likely to convert.",
        accent=RED,
    )
    draw_footer(c, "KPI logic documented from the cleaned dataset and Tableau workbook.", "North-star metric: subscription rate")


def slide_5(c):
    draw_header(
        c,
        5,
        "Key insights from EDA",
        "The analysis reveals who responds, when they respond, and which segment characteristics deserve a call first.",
    )
    # left insight stack
    insights = [
        "<b>Students and retirees convert best.</b> Their subscription rates are materially above the project average, which makes them high-priority prospect pools.",
        "<b>Single clients respond better than married clients.</b> Marital status is not the target on its own, but it helps refine message framing and timing.",
        "<b>Higher balances correlate with stronger conversion.</b> Medium and high balance bands outperform low and zero/negative balances.",
        "<b>Cellular contact beats telephone.</b> Unknown or untracked channels underperform sharply, suggesting poor routing quality.",
        "<b>Seasonality matters.</b> March, September, October, and December have the strongest response rates, so campaign timing is not random.",
    ]
    y = 388
    for i, item in enumerate(insights, start=1):
        draw_card(c, 34, y - 62, 394, 56, radius=14)
        c.setFillColor(BLUE if i < 3 else TEAL if i < 5 else AMBER)
        c.circle(53, y - 34, 11, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 10)
        c.drawCentredString(53, y - 37, str(i))
        draw_paragraph(c, item, "Small", 74, y - 14, 342)
        y -= 66
    matrix_df = DATA.pivot_table(index="age_group", columns="balance_band", values="y", aggfunc="mean")
    row_labels = ["18-24", "25-34", "35-44", "45-54", "55-64", "65+"]
    col_labels = ["negative_or_zero", "low", "medium", "high"]
    matrix = matrix_df.reindex(row_labels).reindex(columns=col_labels).values.tolist()
    draw_heatmap_chart(
        c,
        450,
        118,
        476,
        292,
        "Segment heatmap",
        row_labels,
        col_labels,
        matrix,
        subtitle="Conversion rate by age group and balance band",
    )
    draw_section_card(
        c,
        450,
        40,
        476,
        74,
        "Business readout",
        "The strongest segments are old, warm, and financially healthier. The weakest are low-balance clients and low-context channel records.",
        accent=TEAL,
    )
    draw_footer(c, "EDA based on `cleaned_dataset.csv` and Tableau dashboard 1.", "Strongest single segment: 65+")


def slide_6(c):
    draw_header(
        c,
        6,
        "Advanced analysis",
        "A simple root-cause view shows that warm leads and low-touch campaigns outperform broad repeated calling.",
    )
    draw_bar_chart(
        c,
        34,
        176,
        430,
        224,
        "More contacts, lower conversion",
        ["1", "2-3", "4-5", "6+"],
        list(DATA.groupby("campaign_band")["y"].mean().reindex(["1", "2-3", "4-5", "6+"])),
        [BLUE, TEAL, AMBER, RED],
        subtitle="Subscription rate by campaign band",
        y_max=0.16,
    )
    draw_bar_chart(
        c,
        488,
        176,
        438,
        224,
        "Warm leads convert much better",
        ["No prior", "Warm"],
        [DATA.groupby("previously_contacted")["y"].mean().loc[0], DATA.groupby("previously_contacted")["y"].mean().loc[1]],
        [BLUE, TEAL],
        subtitle="Previously contacted vs first-time contact",
        y_max=0.26,
    )
    c.setStrokeColor(LINE)
    c.setLineWidth(1)
    c.line(34, 146, 926, 146)
    draw_card(c, 34, 56, 892, 48, radius=14, fill=colors.HexColor("#EEF7F5"))
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 11)
    c.drawString(52, 86, "Interpretation")
    c.setFont("Helvetica", 10)
    c.drawString(132, 86, "The problem is not only volume. It is misallocated volume: too many low-propensity clients receive repeated outreach.")
    c.drawString(132, 71, "That makes segmentation, contact caps, and warm-lead routing the main optimization levers.")
    draw_footer(c, "Advanced analysis uses segment comparisons and nonparametric tests on the cleaned dataset.", "Root cause: over-contacting weak leads")


def slide_7(c, assets):
    draw_header(
        c,
        7,
        "Dashboard walkthrough",
        "The executive view summarizes what happened; the operational views explain why and where the lift lives.",
    )
    # left executive
    draw_card(c, 34, 126, 426, 290, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(52, 394, "Executive view")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawString(52, 379, "Dashboard 1: customer segments, balance bands, and overall conversion")
    c.drawImage(str(assets["dash1"]), 50, 144, width=410, height=232, preserveAspectRatio=True, anchor='c')
    # right operational stack
    draw_card(c, 478, 126, 448, 138, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(496, 246, "Operational view 1")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawString(496, 231, "Dashboard 2: contact lens")
    c.drawImage(str(assets["dash2"]), 494, 142, width=412, height=112, preserveAspectRatio=True, anchor='c')
    draw_card(c, 478, 278, 448, 138, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(496, 398, "Operational view 2")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawString(496, 383, "Dashboard 3: risk lens")
    c.drawImage(str(assets["dash3"]), 494, 294, width=412, height=112, preserveAspectRatio=True, anchor='c')
    draw_card(c, 34, 38, 892, 72, radius=16)
    c.setFillColor(TEAL)
    c.rect(34, 38 + 65, 892, 7, stroke=0, fill=1)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(50, 38 + 48, "How to read the dashboards")
    draw_paragraph(
        c,
        "Use the executive view for leadership reporting. Use the operational views to filter by job, age group, education, contact type, balance band, and prior campaign history before deciding who should be called next.",
        "Body",
        50,
        38 + 28,
        842,
    )
    draw_footer(c, "Tableau screenshots from the project workbook.", "Executive and operational views")


def slide_8(c):
    draw_header(
        c,
        8,
        "Recommendations",
        "Each recommendation is directly tied to one of the strongest observed signals in the analysis.",
    )
    recs = [
        ("Prioritize students, retirees, and 65+ clients", "These groups consistently show the highest conversion, so they should enter the call queue first.", BLUE),
        ("Use prior-contact flags and previous outcomes", "Warm leads and prior success history deliver the biggest lift, so the CRM should route them differently.", TEAL),
        ("Cap repeated outreach after 3 touches", "Conversion drops sharply after that point, so over-contacting should be treated as waste.", AMBER),
        ("Prefer cellular-first contact plans", "The cellular channel beats telephone, while unknown channels are weak and should be cleaned up.", RED),
    ]
    positions = [(34, 272), (478, 272), (34, 116), (478, 116)]
    for (title, body, accent), (x, y) in zip(recs, positions):
        draw_card(c, x, y, 448, 132, radius=16)
        c.setFillColor(accent)
        c.rect(x, y + 125, 448, 7, stroke=0, fill=1)
        c.setFillColor(accent)
        c.circle(x + 26, y + 82, 16, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 18)
        c.drawCentredString(x + 26, y + 75, "+")
        c.setFillColor(INK)
        c.setFont("Helvetica-Bold", 12.5)
        c.drawString(x + 56, y + 92, title)
        draw_paragraph(c, body, "Body", x + 56, y + 75, 364)
    draw_footer(c, "Recommendations linked to EDA, KPI framework, and advanced analysis.", "Actionable and measurable")


def slide_9(c):
    draw_header(
        c,
        9,
        "Impact and value",
        "The opportunity is to convert the same outreach volume into a better subscription yield and fewer wasted calls.",
    )
    current = DATA["y"].mean()
    warm = DATA.groupby("previously_contacted")["y"].mean().loc[1]
    success = DATA.groupby("poutcome")["y"].mean().loc["success"]
    one_touch = DATA.groupby("campaign_band")["y"].mean().loc["1"]
    cards = [
        ("Baseline", fmt_pct(current), "1,170 subscriptions per 10,000 contacts", MUTED),
        ("Previously contacted", fmt_pct(warm), "2,307 subscriptions per 10,000 contacts", TEAL),
        ("Prior success", fmt_pct(success), "6,473 subscriptions per 10,000 contacts", AMBER),
    ]
    xs = [34, 340, 646]
    for (title, value, body, accent), x in zip(cards, xs):
        draw_metric_box(c, x, 272, 280, 138, title, value, body, accent=accent, value_size=28)
    draw_card(c, 34, 96, 510, 170, radius=16)
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(52, 236, "Illustrative value model")
    c.setFillColor(MUTED)
    c.setFont("Helvetica", 9.5)
    c.drawString(52, 221, "Assumption used only for presentation: each incremental subscription contributes ₹50 of net value.")
    c.setFillColor(BLUE)
    c.setFont("Helvetica-Bold", 16)
    c.drawString(52, 186, "If 10,000 contacts are routed to previously contacted leads:")
    c.setFillColor(INK)
    c.setFont("Helvetica", 12)
    c.drawString(52, 164, f"Baseline = {current*100:.1f}% -> {int(round(10_000*current)):,} subscriptions")
    c.drawString(52, 144, f"Targeted = {warm*100:.1f}% -> {int(round(10_000*warm)):,} subscriptions")
    c.drawString(52, 124, f"Incremental lift = {int(round(10_000*(warm-current))):,} extra subscriptions")
    c.drawString(52, 104, f"Illustrative value = ~₹{round(10_000*(warm-current)*50):,.0f} per 10,000 contacts")
    draw_bar_chart(
        c,
        560,
        96,
        366,
        150,
        "Scenario lift",
        ["Base", "1-touch", "Warm", "Success"],
        [current, one_touch, warm, success],
        [MUTED, BLUE, TEAL, AMBER],
        subtitle="Estimated response scenarios",
        y_max=0.7,
    )
    draw_footer(c, "Value model is conservative and assumption-based; replace ₹50 with the bank's approved contribution margin.", "Efficiency and revenue proxy")


def slide_10(c):
    draw_header(
        c,
        10,
        "Limitations and next steps",
        "The dashboard is decision-ready, but the next layer should optimize what to do with each lead, not just describe who converted.",
    )
    draw_section_card(
        c,
        34,
        188,
        418,
        206,
        "Limitations",
        "1) No channel-cost data, so ROI is estimated rather than exact.<br/>2) `duration` is useful for analysis but leaks outcome information if used in prediction.<br/>3) The dataset does not contain product price, interest-rate, or offer-level detail.<br/>4) The campaign is a single historical window, so seasonality should be validated in new data.",
        accent=AMBER,
    )
    draw_section_card(
        c,
        478,
        188,
        448,
        206,
        "Next steps",
        "1) Build an uplift model or scorecard to rank leads before dialing.<br/>2) Test contact caps with an A/B experiment.<br/>3) Combine this bank dataset with call-cost and margin data for a true ROI model.<br/>4) Expand the dashboard with monthly trend overlays and segment-level drilldowns.",
        accent=TEAL,
    )
    draw_card(c, 34, 72, 892, 88, radius=16, fill=colors.HexColor("#F0F5FB"))
    c.setFillColor(INK)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(56, 122, "Closing note")
    c.setFont("Helvetica", 13)
    c.drawString(56, 98, "The project already tells the bank where the lift is. The next step is to turn those signals into a routing engine.")
    c.drawString(56, 78, "In other words: better targeting, fewer wasted calls, and a cleaner path from Tableau insight to campaign action.")
    draw_footer(c, "Project summary based on the cleaned UCI Bank Marketing dataset and Tableau dashboards.", "Ready for mentor review")


def build_pdf():
    assets = create_assets()
    out = REPORTS / "bank_marketing_dva_deck.pdf"
    c = canvas.Canvas(str(out), pagesize=(SLIDE_W, SLIDE_H))
    draw_cover(c, assets)
    c.showPage()
    slide_2(c)
    c.showPage()
    slide_3(c)
    c.showPage()
    slide_4(c)
    c.showPage()
    slide_5(c)
    c.showPage()
    slide_6(c)
    c.showPage()
    slide_7(c, assets)
    c.showPage()
    slide_8(c)
    c.showPage()
    slide_9(c)
    c.showPage()
    slide_10(c)
    c.showPage()
    c.save()
    return out


if __name__ == "__main__":
    pdf = build_pdf()
    print(pdf)
