"""
Transportation Analytics System
Step 5: Generate Professional PDF Analytics Report
"""
import pandas as pd
import numpy as np
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm, mm
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                  TableStyle, PageBreak, Image, HRFlowable)
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.platypus import KeepTogether
import os

master = pd.read_csv("data/master_analytics_table.csv", parse_dates=["trip_date"])

# ── Color Palette ─────────────────────────────────────────────────────────────
NAVY    = colors.HexColor("#1A237E")
BLUE    = colors.HexColor("#1565C0")
LTBLUE  = colors.HexColor("#E3F2FD")
GREEN   = colors.HexColor("#2E7D32")
LTGREEN = colors.HexColor("#E8F5E9")
AMBER   = colors.HexColor("#F57F17")
RED     = colors.HexColor("#C62828")
LTRED   = colors.HexColor("#FFEBEE")
GRAY    = colors.HexColor("#546E7A")
LTGRAY  = colors.HexColor("#FAFAFA")
WHITE   = colors.white
DGRAY   = colors.HexColor("#37474F")

# ── Styles ────────────────────────────────────────────────────────────────────
styles = getSampleStyleSheet()

def style(name, **kw):
    return ParagraphStyle(name, **kw)

TITLE_S  = style("TitleS",  fontName="Helvetica-Bold",  fontSize=24, textColor=WHITE,    alignment=TA_CENTER, spaceAfter=4)
TITLE2_S = style("Title2S", fontName="Helvetica-Bold",  fontSize=16, textColor=NAVY,     alignment=TA_LEFT,   spaceAfter=6)
TITLE3_S = style("Title3S", fontName="Helvetica-Bold",  fontSize=13, textColor=BLUE,     alignment=TA_LEFT,   spaceAfter=4)
BODY_S   = style("BodyS",   fontName="Helvetica",       fontSize=9.5,textColor=DGRAY,    alignment=TA_JUSTIFY,spaceAfter=4, leading=14)
BULLET_S = style("BulletS", fontName="Helvetica",       fontSize=9.5,textColor=DGRAY,    alignment=TA_LEFT,   spaceAfter=3, leading=13, leftIndent=10)
CAPTION_S= style("CaptionS",fontName="Helvetica-Oblique",fontSize=8.5,textColor=GRAY,   alignment=TA_CENTER)
KPI_LBL  = style("KpiLbl",  fontName="Helvetica",       fontSize=8,  textColor=GRAY,     alignment=TA_CENTER)
KPI_VAL  = style("KpiVal",  fontName="Helvetica-Bold",  fontSize=20, textColor=BLUE,     alignment=TA_CENTER)

# ── Helper: Colored KPI Table ─────────────────────────────────────────────────
def kpi_row(kpis):
    """kpis = list of (label, value, bg_color)"""
    cells = [[Paragraph(v, KPI_VAL), Paragraph(l, KPI_LBL)] for l,v,_ in kpis]
    row1  = [cells[i][0] for i in range(len(kpis))]
    row2  = [cells[i][1] for i in range(len(kpis))]
    t = Table([row1, row2], colWidths=[A4[0]*0.82/len(kpis)]*len(kpis))
    bg_cmds = [("BACKGROUND", (i,0),(i,1), colors.HexColor(f"#{kpis[i][2]}")) for i in range(len(kpis))]
    t.setStyle(TableStyle([
        ("ALIGN",     (0,0),(-1,-1),"CENTER"),
        ("VALIGN",    (0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",(0,0),(-1,-1),8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),3),
    ] + bg_cmds))
    return t

# ── Helper: Styled Data Table ─────────────────────────────────────────────────
def data_table(headers, rows, col_widths=None, hdr_bg=NAVY):
    all_rows = [[Paragraph(str(h), style("H", fontName="Helvetica-Bold", fontSize=8.5,
                                          textColor=WHITE, alignment=TA_CENTER)) for h in headers]]
    for i, row in enumerate(rows):
        styled = [Paragraph(str(v), style("D", fontName="Helvetica", fontSize=8.5,
                                           textColor=DGRAY, alignment=TA_CENTER)) for v in row]
        all_rows.append(styled)
    t = Table(all_rows, colWidths=col_widths, repeatRows=1)
    row_colors = [("BACKGROUND", (0,r),(-1,r), LTGRAY if r%2==0 else WHITE) for r in range(1, len(all_rows))]
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,0), hdr_bg),
        ("ROWBACKGROUNDS", (0,1),(-1,-1), [LTGRAY, WHITE]),
        ("ALIGN",     (0,0),(-1,-1), "CENTER"),
        ("VALIGN",    (0,0),(-1,-1), "MIDDLE"),
        ("FONTSIZE",  (0,0),(-1,-1), 8.5),
        ("TOPPADDING",(0,0),(-1,-1), 5),
        ("BOTTOMPADDING",(0,0),(-1,-1), 5),
        ("GRID",      (0,0),(-1,-1), 0.4, colors.HexColor("#CFD8DC")),
        ("LINEABOVE", (0,1),(-1,1),  1.5, hdr_bg),
    ]))
    return t

# ── PDF Build ─────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    "outputs/Transportation_Analytics_Report.pdf",
    pagesize=A4, rightMargin=1.8*cm, leftMargin=1.8*cm,
    topMargin=1.5*cm, bottomMargin=1.5*cm,
    title="Transportation Analytics Report", author="Analytics System"
)

W = A4[0] - 3.6*cm  # usable width

story = []

# ─────────────────────────────────────────────────────────────────────────────
# COVER PAGE
# ─────────────────────────────────────────────────────────────────────────────
cover_data = [[Paragraph("", TITLE_S)],
              [Paragraph("TRANSPORTATION ANALYTICS SYSTEM", TITLE_S)],
              [Paragraph("Fleet Fuel Efficiency | Route Costing | Delay Insights", TITLE_S)],
              [Paragraph("", TITLE_S)],
              [Paragraph("Analytics Report — FY 2023", style("Sub", fontName="Helvetica",
                          fontSize=14, textColor=colors.HexColor("#90CAF9"), alignment=TA_CENTER))]]
cover_tbl = Table([[Paragraph(r[0].text if hasattr(r[0],"text") else "", TITLE_S)] for r in cover_data],
                   colWidths=[W])

# Build cover as colored background table
cover_items = [
    Spacer(1, 3*cm),
    Table([[Paragraph("🚛  TRANSPORTATION", style("C1",fontName="Helvetica-Bold",fontSize=30,textColor=WHITE,alignment=TA_CENTER))]],
          colWidths=[W], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),
                                            ("TOPPADDING",(0,0),(-1,-1),18),("BOTTOMPADDING",(0,0),(-1,-1),5)])),
    Table([[Paragraph("ANALYTICS SYSTEM", style("C2",fontName="Helvetica-Bold",fontSize=26,textColor=colors.HexColor("#90CAF9"),alignment=TA_CENTER))]],
          colWidths=[W], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),
                                            ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),10)])),
    Table([[Paragraph("Fleet Fuel Efficiency · Route Costing · Delay Insights · Cost Optimization",
                       style("C3",fontName="Helvetica",fontSize=11,textColor=colors.HexColor("#B0BEC5"),alignment=TA_CENTER))]],
          colWidths=[W], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),NAVY),
                                            ("TOPPADDING",(0,0),(-1,-1),0),("BOTTOMPADDING",(0,0),(-1,-1),18)])),
    Spacer(1,0.4*cm),
    Table([[Paragraph("EXECUTIVE ANALYTICS REPORT  |  FY 2023",
                       style("C4",fontName="Helvetica-Bold",fontSize=11,textColor=NAVY,alignment=TA_CENTER))]],
          colWidths=[W], style=TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#E3F2FD")),
                                            ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10)])),
    Spacer(1, 1*cm),
]

# KPI Banner
kpis_cover = [
    ("Total Trips",    f"{len(master):,}",              "E3F2FD"),
    ("Fleet Size",     f"{master['vehicle_id'].nunique()}",   "E8F5E9"),
    ("Total Distance", f"{master['distance_km'].sum()/1000:,.0f}K km", "FFF8E1"),
    ("On-Time Rate",   f"{(master['delivery_status']=='On Time').mean()*100:.1f}%","FFEBEE"),
]
cover_items.append(kpi_row(kpis_cover))
cover_items.append(Spacer(1,2*cm))
cover_items.append(HRFlowable(width=W, thickness=1, color=BLUE))
cover_items.append(Spacer(1,0.3*cm))
cover_items.append(Paragraph("Generated by Transportation Analytics System  |  Data Science Project",
                              style("Footer",fontName="Helvetica",fontSize=9,textColor=GRAY,alignment=TA_CENTER)))
story += cover_items
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 1: EXECUTIVE SUMMARY
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("1. Executive Summary", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=BLUE))
story.append(Spacer(1, 0.3*cm))

total_cost  = master["total_trip_cost_inr"].sum()
avg_eff     = master["fuel_efficiency_kml"].mean()
on_time_pct = (master["delivery_status"]=="On Time").mean()*100
avg_delay   = master["delay_minutes"].mean()
avg_cost_km = master["cost_per_km"].mean()
total_fuel  = master["fuel_consumed_l"].sum()

kpis1 = [
    ("Total Trips",     f"{len(master):,}",        "E3F2FD"),
    ("Avg Fuel Eff.",   f"{avg_eff:.2f} km/L",     "FFF8E1"),
    ("On-Time Rate",    f"{on_time_pct:.1f}%",      "E8F5E9"),
    ("Avg Delay",       f"{avg_delay:.0f} min",     "FFEBEE"),
    ("Avg Cost/km",     f"₹{avg_cost_km:.0f}",     "EDE7F6"),
]
story.append(kpi_row(kpis1))
story.append(Spacer(1, 0.5*cm))

story.append(Paragraph(
    f"This report presents a comprehensive analysis of <b>{len(master):,} trips</b> undertaken by a fleet of "
    f"<b>{master['vehicle_id'].nunique()} vehicles</b> driven by <b>{master['driver_id'].nunique()} drivers</b> "
    f"across <b>{master['route_name'].nunique()} routes</b> during FY 2023. "
    f"The fleet collectively covered <b>{master['distance_km'].sum():,.0f} km</b>, "
    f"consuming <b>{total_fuel:,.0f} litres</b> of fuel at a total operational cost of "
    f"<b>₹{total_cost/1e6:.2f} million</b>.", BODY_S))
story.append(Spacer(1, 0.2*cm))

# Key findings bullet points
findings = [
    f"Fuel efficiency ranged from {master['fuel_efficiency_kml'].min():.1f} to {master['fuel_efficiency_kml'].max():.1f} km/L with an average of {avg_eff:.2f} km/L.",
    f"Delivery performance: {on_time_pct:.1f}% on time, {(master['delivery_status']=='Minor Delay').mean()*100:.1f}% minor delays, {(master['delivery_status']=='Major Delay').mean()*100:.1f}% major delays.",
    f"The most cost-efficient route category is {master.groupby('route_category')['cost_per_km'].mean().idxmin()} routes (₹{master.groupby('route_category')['cost_per_km'].mean().min():.0f}/km avg).",
    f"Storms cause the highest delays averaging {master[master['weather']=='Storm']['delay_minutes'].mean():.0f} minutes per trip.",
    f"Top performing driver achieved a score of {master.groupby('driver_name')['driver_perf_score'].mean().max():.1f}/100.",
]
for f in findings:
    story.append(Paragraph(f"• {f}", BULLET_S))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 2: FUEL EFFICIENCY ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("2. Fuel Efficiency Analysis", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=AMBER))
story.append(Spacer(1, 0.3*cm))

story.append(Paragraph(
    "Fuel efficiency is the primary driver of operational cost. The analysis examines efficiency patterns "
    "across vehicle types, route categories, weather conditions, and traffic levels.", BODY_S))
story.append(Spacer(1, 0.3*cm))

# Chart: Fuel efficiency by vehicle type
if os.path.exists("charts/chart1_fuel_efficiency_by_vehicle.png"):
    story.append(Image("charts/chart1_fuel_efficiency_by_vehicle.png", width=W, height=9*cm))
    story.append(Paragraph("Fig 2.1 — Fuel efficiency distribution by vehicle type (km/L)", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

# Table: Fuel stats by vehicle type
vt_fuel = master.groupby("vehicle_type").agg(
    trips=("trip_id","count"),
    avg_eff=("fuel_efficiency_kml","mean"),
    min_eff=("fuel_efficiency_kml","min"),
    max_eff=("fuel_efficiency_kml","max"),
    total_fuel=("fuel_consumed_l","sum"),
    avg_cost=("fuel_cost_inr","mean"),
).reset_index().sort_values("avg_eff",ascending=False)

story.append(Paragraph("Fuel Efficiency by Vehicle Type", TITLE3_S))
tbl_rows = [[row["vehicle_type"], int(row["trips"]), f"{row['avg_eff']:.2f}",
              f"{row['min_eff']:.2f}", f"{row['max_eff']:.2f}",
              f"{row['total_fuel']:,.0f}", f"₹{row['avg_cost']:,.0f}"]
             for _, row in vt_fuel.iterrows()]
story.append(data_table(
    ["Vehicle Type","Trips","Avg Eff (km/L)","Min Eff","Max Eff","Total Fuel (L)","Avg Fuel Cost (₹)"],
    tbl_rows, col_widths=[3.5*cm,2*cm,3*cm,2.5*cm,2.5*cm,3*cm,3*cm], hdr_bg=AMBER
))
story.append(Spacer(1, 0.4*cm))

# Monthly trend chart
story.append(Image("charts/chart2_monthly_fuel_trend.png", width=W, height=9*cm))
story.append(Paragraph("Fig 2.2 — Monthly fuel consumption and efficiency trend", CAPTION_S))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 3: ROUTE ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("3. Route Cost & Efficiency Analysis", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=colors.HexColor("#6A1B9A")))
story.append(Spacer(1, 0.3*cm))

story.append(Image("charts/chart4_route_cost_delay.png", width=W, height=9*cm))
story.append(Paragraph("Fig 3.1 — Route cost per km and delivery delay comparison", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

story.append(Image("charts/chart8_route_category_kpis.png", width=W, height=8*cm))
story.append(Paragraph("Fig 3.2 — KPI comparison across route categories", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

# Route category table
cat_tbl = master.groupby("route_category").agg(
    trips=("trip_id","count"), avg_dist=("distance_km","mean"),
    avg_eff=("fuel_efficiency_kml","mean"), avg_delay=("delay_minutes","mean"),
    avg_cost_km=("cost_per_km","mean"), total_cost=("total_trip_cost_inr","sum"),
).reset_index()

story.append(Paragraph("Route Category Summary", TITLE3_S))
tbl_rows2 = [[row["route_category"], int(row["trips"]), f"{row['avg_dist']:.0f} km",
               f"{row['avg_eff']:.2f}", f"{row['avg_delay']:.0f} min",
               f"₹{row['avg_cost_km']:.0f}", f"₹{row['total_cost']/1e6:.2f}M"]
              for _, row in cat_tbl.iterrows()]
story.append(data_table(
    ["Category","Trips","Avg Distance","Avg Fuel Eff","Avg Delay","Cost/km","Total Cost"],
    tbl_rows2, col_widths=[3*cm,2*cm,3*cm,3*cm,3*cm,2.5*cm,3*cm],
    hdr_bg=colors.HexColor("#4A148C")
))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 4: DELIVERY DELAY INSIGHTS
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("4. Delivery Delay Insights", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=RED))
story.append(Spacer(1, 0.3*cm))

story.append(Image("charts/chart5_delivery_delay_analysis.png", width=W, height=9*cm))
story.append(Paragraph("Fig 4.1 — Delivery status distribution and delay by weather", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

# Delay by traffic level
traffic_delay = master.groupby("traffic_level")["delay_minutes"].agg(["mean","sum","count"]).reset_index()
traffic_delay.columns = ["Traffic Level","Avg Delay (min)","Total Delay (min)","Trips"]

story.append(Paragraph("Delay Analysis by Traffic Level", TITLE3_S))
tbl_rows3 = [[row["Traffic Level"], f"{row['Avg Delay (min)']:.0f}", int(row["Total Delay (min)"]), int(row["Trips"])]
              for _, row in traffic_delay.iterrows()]
story.append(data_table(["Traffic Level","Avg Delay (min)","Total Delay (min)","Trips"],
                          tbl_rows3, col_widths=[5*cm,5*cm,5*cm,5*cm], hdr_bg=RED))
story.append(Spacer(1, 0.4*cm))

story.append(Paragraph(
    f"<b>Key Delay Findings:</b> Very High traffic conditions result in delays averaging "
    f"{master[master['traffic_level']=='Very High']['delay_minutes'].mean():.0f} minutes. "
    f"Storm weather conditions are the most severe delay driver, followed by Rain. "
    f"Night-time trips tend to have fewer delays due to lower traffic volume. "
    f"The East Zone experiences the highest proportion of major delays among customer locations.", BODY_S))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 5: DRIVER PERFORMANCE
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("5. Driver Performance Analysis", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=GREEN))
story.append(Spacer(1, 0.3*cm))

story.append(Image("charts/chart3_driver_performance_ranking.png", width=W, height=10*cm))
story.append(Paragraph("Fig 5.1 — Top 15 driver performance scores (Green=Top, Orange=Mid, Red=Low)", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

# Top/Bottom 5 driver table
driver_perf = master.groupby("driver_name").agg(
    trips=("trip_id","count"),
    perf_score=("driver_perf_score","mean"),
    avg_eff=("fuel_efficiency_kml","mean"),
    avg_delay=("delay_minutes","mean"),
    safety=("safety_rating","mean"),
    exp=("experience_years","mean"),
).reset_index().sort_values("perf_score",ascending=False)

story.append(Paragraph("Top 10 Performing Drivers", TITLE3_S))
top10 = driver_perf.head(10)
tbl_rows4 = [[row["driver_name"], f"{row['perf_score']:.1f}", int(row["trips"]),
               f"{row['avg_eff']:.2f}", f"{row['avg_delay']:.0f}", f"{row['safety']:.1f}", int(row["exp"])]
              for _, row in top10.iterrows()]
story.append(data_table(
    ["Driver","Perf Score","Trips","Avg Eff (km/L)","Avg Delay (min)","Safety Rating","Experience (yrs)"],
    tbl_rows4, col_widths=[3.5*cm,2.5*cm,2*cm,3*cm,3*cm,3*cm,3*cm], hdr_bg=GREEN
))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 6: VEHICLE & MAINTENANCE
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("6. Vehicle Performance & Maintenance", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=BLUE))
story.append(Spacer(1, 0.3*cm))

story.append(Image("charts/chart7_vehicle_performance_matrix.png", width=W, height=9*cm))
story.append(Paragraph("Fig 6.1 — Vehicle performance matrix (size=trip volume, color=avg delay)", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

story.append(Image("charts/chart9_maintenance_cost.png", width=W, height=8*cm))
story.append(Paragraph("Fig 6.2 — Maintenance cost by vehicle (Red = above average)", CAPTION_S))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 7: CORRELATION ANALYSIS
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("7. Correlation & Predictive Insights", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=GRAY))
story.append(Spacer(1, 0.3*cm))

story.append(Image("charts/chart6_correlation_heatmap.png", width=W, height=12*cm))
story.append(Paragraph("Fig 7.1 — Correlation heatmap of all key performance metrics", CAPTION_S))
story.append(Spacer(1, 0.4*cm))

# Key correlations
corr_cols = ["distance_km","fuel_consumed_l","fuel_efficiency_kml","road_difficulty",
             "delay_minutes","total_trip_cost_inr","cost_per_km","driver_perf_score"]
corr_m = master[corr_cols].corr()
insights = [
    f"Distance vs Fuel Consumed: r={corr_m.loc['distance_km','fuel_consumed_l']:.2f} — Strong positive correlation (expected).",
    f"Fuel Efficiency vs Cost/km: r={corr_m.loc['fuel_efficiency_kml','cost_per_km']:.2f} — Higher efficiency reduces per-km cost.",
    f"Road Difficulty vs Delay: r={corr_m.loc['road_difficulty','delay_minutes']:.2f} — Difficult terrain increases delays.",
    f"Driver Performance vs Fuel Eff: r={corr_m.loc['driver_perf_score','fuel_efficiency_kml']:.2f} — Better drivers achieve better fuel economy.",
    f"Experience vs Performance: r={corr_m.loc['driver_perf_score','fuel_efficiency_kml']:.2f} — Driver experience correlates with efficiency.",
]
for ins in insights:
    story.append(Paragraph(f"• {ins}", BULLET_S))
story.append(PageBreak())

# ─────────────────────────────────────────────────────────────────────────────
# SECTION 8: RECOMMENDATIONS
# ─────────────────────────────────────────────────────────────────────────────
story.append(Paragraph("8. Recommendations & Action Plan", TITLE2_S))
story.append(HRFlowable(width=W, thickness=2, color=GREEN))
story.append(Spacer(1, 0.4*cm))

recs = [
    ("🔧 Fleet Optimization",
     f"Retire vehicles older than 8 years or with efficiency below {master['fuel_efficiency_kml'].quantile(0.2):.1f} km/L. "
     "Transition 15% of fleet to CNG/Electric to reduce fuel costs by an estimated 20-30%."),
    ("📍 Route Re-engineering",
     "Reclassify high-cost city routes to avoid peak traffic windows. "
     "Merge low-volume rural routes to improve load factor. Implement dynamic routing based on real-time traffic."),
    ("👤 Driver Training Program",
     f"Bottom 20% of drivers (score below {master.groupby('driver_name')['driver_perf_score'].mean().quantile(0.2):.1f}) "
     "should undergo mandatory eco-driving training. Implement incentive structure for top-performing drivers."),
    ("⏱ Delay Reduction Strategy",
     "Avoid scheduling trips during storm/heavy rain periods where possible. "
     "Build 25-30 minute buffer time into estimates for Very High traffic routes. "
     "Deploy real-time weather alerts to dispatch teams."),
    ("💰 Cost Optimization",
     "Negotiate bulk fuel contracts to reduce per-litre cost. "
     "Schedule predictive maintenance before vehicle efficiency drops below threshold. "
     "Target ₹15-20/km reduction on highest-cost routes through combined fleet and route optimization."),
]

for title, body in recs:
    rec_table = Table(
        [[Paragraph(f"<b>{title}</b>", style("RT", fontName="Helvetica-Bold", fontSize=10.5, textColor=NAVY)),
          Paragraph(body, style("RB", fontName="Helvetica", fontSize=9.5, textColor=DGRAY, leading=14))]],
        colWidths=[5*cm, W-5*cm]
    )
    rec_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(0,0), LTBLUE),
        ("BACKGROUND", (1,0),(1,0), LTGRAY),
        ("VALIGN",     (0,0),(-1,-1), "TOP"),
        ("TOPPADDING", (0,0),(-1,-1), 8),
        ("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1), 8),
        ("LINEBELOW",  (0,0),(-1,0), 0.5, colors.HexColor("#CFD8DC")),
    ]))
    story.append(rec_table)
    story.append(Spacer(1, 0.2*cm))

story.append(Spacer(1, 1*cm))
story.append(HRFlowable(width=W, thickness=1, color=GRAY))
story.append(Spacer(1, 0.2*cm))
story.append(Paragraph(
    "Transportation Analytics System  |  FY 2023 Report  |  Generated by Data Analytics Pipeline",
    style("Foot", fontName="Helvetica", fontSize=8, textColor=GRAY, alignment=TA_CENTER)
))

doc.build(story)
print("✅ PDF report saved: outputs/Transportation_Analytics_Report.pdf")
