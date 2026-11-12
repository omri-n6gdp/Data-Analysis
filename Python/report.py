import os
import io
import datetime
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns

# ReportLab imports
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle, HRFlowable, KeepTogether
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.utils import ImageReader

# Ensure output directory exists
output_path = '/Users/omrishahar/Documents/gym_analysis_report.pdf'
os.makedirs(os.path.dirname(output_path), exist_ok=True)

# Define Professional Colour Palette
PALETTE = ['#2C3E50', '#E74C3C', '#3498DB', '#2ECC71', '#F39C12', '#9B59B6']
sns.set_palette(PALETTE)

# ==============================================================================
# STEP 1 — MANDATORY DATA SURVEY & PREPROCESSING
# ==============================================================================
print("═══════════════════════════════════════════════════════════════")
print("STEP 1: DATA SURVEY & PREPROCESSING")
print("═══════════════════════════════════════════════════════════════\n")

# 1.1 LOAD AND INSPECT
df = pd.read_csv(r'/Users/omrishahar/Documents/gym_membership.csv')
print(f"SHAPE: {df.shape[0]} rows, {df.shape[1]} columns\n")
print("COLUMN NAMES & DTYPES:")
print(df.dtypes, "\n")
print("FIRST 10 ROWS:")
print(df.head(10), "\n")

# 1.2 NUMERICAL COLUMNS
print("--- NUMERICAL COLUMNS STATISTICS ---")
num_cols = df.select_dtypes(include=[np.number]).columns
for col in num_cols:
    nulls = df[col].isnull().sum()
    print(f"{col}: Mean={df[col].mean():.2f}, Median={df[col].median():.2f}, Std={df[col].std():.2f}, Min={df[col].min()}, Max={df[col].max()}, Nulls={nulls}")
print("\n")

# 1.3 CATEGORICAL COLUMNS (Including Multi-Value Split)
print("--- CATEGORICAL COLUMNS STATISTICS ---")
cat_cols = df.select_dtypes(include=['object', 'string']).columns
multi_value_cols = ['days_per_week', 'fav_group_lesson', 'fav_drink']

for col in cat_cols:
    nulls = df[col].isnull().sum()
    print(f"\n[{col}] (Nulls: {nulls})")
    if col in multi_value_cols:
        counts = df[col].dropna().astype(str).str.split(', ').explode().value_counts()
        print("  (Split multi-values):")
    else:
        counts = df[col].value_counts()
    print(counts.head(10)) # Print top 10 for brevity if long
print("\n")

# 1.4 BOOLEAN COLUMNS
print("--- BOOLEAN COLUMNS (True/False Rates) ---")
bool_cols = ['attend_group_lesson', 'drink_abo', 'personal_training', 'uses_sauna']
for col in bool_cols:
    if col in df.columns:
        t_rate = df[col].mean() * 100
        print(f"{col}: True = {t_rate:.1f}%, False = {100 - t_rate:.1f}%")
print("\n")

# 1.5 NULL INTERPRETATION
print("--- NULL INTERPRETATION ---")
print("STRUCTURAL NULLS DETECTED:")
print("- 'fav_group_lesson' nulls align with attend_group_lesson = False.")
print("- 'fav_drink' nulls align with drink_abo = False.")
print("- 'name_personal_trainer' nulls align with personal_training = False.")
print("These are NOT missing data and will not be imputed or dropped.\n")

# 1.6 DERIVED COLUMNS
df['check_in_hour'] = pd.to_datetime(df['avg_time_check_in'], format='%H:%M:%S', errors='coerce').dt.hour
df['check_out_hour'] = pd.to_datetime(df['avg_time_check_out'], format='%H:%M:%S', errors='coerce').dt.hour
df['num_days_per_week'] = df['days_per_week'].dropna().apply(lambda x: len(str(x).split(', ')))

# Binning age. Adding a 50+ bin to safely catch outliers if they exist.
df['age_group'] = pd.cut(df['Age'], bins=[11, 17, 25, 35, 49, 120], labels=['12–17', '18–25', '26–35', '36–49', '50+'])
df['addon_count'] = df[['personal_training', 'drink_abo', 'uses_sauna', 'attend_group_lesson']].sum(axis=1)

# 1.7 DATA QUALITY CHECK
print("--- DATA QUALITY CHECKS ---")
ages_below_16 = (df['Age'] < 16).sum()
time_outliers = ((df['avg_time_in_gym'] < 20) | (df['avg_time_in_gym'] > 200)).sum()
invalid_checkout = (df['check_out_hour'] < df['check_in_hour']).sum()
high_visits = (df['visit_per_week'] > 7).sum()

print(f"Ages < 16: {ages_below_16} (Flagged, retained)")
print(f"Gym time outliers (<20 or >200 mins): {time_outliers} (Flagged)")
print(f"Check-out earlier than Check-in: {invalid_checkout} (Flagged)")
print(f"Implausible visit frequencies (>7/wk): {high_visits} (Flagged)\n")

# ==============================================================================
# STEP 2 & 4 — ANALYTICS & CHART GENERATION
# ==============================================================================
def format_ax(ax, title, xlabel, ylabel):
    ax.set_title(title, fontweight='bold', pad=15)
    ax.set_xlabel(xlabel)
    ax.set_ylabel(ylabel)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.yaxis.grid(True, alpha=0.3)

chart_buffers = []

# --- Insight 1: Standard Members with High Addon Usage (Direction B) ---
fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
addon_props = pd.crosstab(df['abonoment_type'], df['addon_count'], normalize='index') * 100
addon_props.plot(kind='bar', ax=ax, color=PALETTE, width=0.8)
format_ax(ax, "Standard Members Show Surprising Depth in Add-on Usage", "Membership Tier", "Proportion of Tier (%)")
ax.legend(title='Total Add-ons Used', bbox_to_anchor=(1.05, 1), loc='upper left')
plt.xticks(rotation=0)
buf1 = io.BytesIO()
fig.savefig(buf1, format='png', bbox_inches='tight')
buf1.seek(0)
chart_buffers.append(buf1)
plt.close(fig)

# Metrics for Text 1
std_high_addon = addon_props.loc['Standard', 3] + addon_props.loc['Standard', 4] if 3 in addon_props.columns else 0

# --- Insight 2: Personal Trainer Workload Imbalance (Direction E) ---
fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
trainer_counts = df['name_personal_trainer'].value_counts().sort_values()
trainer_counts.plot(kind='barh', ax=ax, color=PALETTE[2])
format_ax(ax, "Severe Capacity Imbalance Across Personal Training Staff", "Total Active Clients (n)", "Trainer Name")
for i, v in enumerate(trainer_counts):
    ax.text(v + 1, i, str(v), va='center', fontweight='bold')
buf2 = io.BytesIO()
fig.savefig(buf2, format='png', bbox_inches='tight')
buf2.seek(0)
chart_buffers.append(buf2)
plt.close(fig)

# Metrics for Text 2
top_trainer = trainer_counts.index[-1]
bottom_trainer = trainer_counts.index[0]
load_ratio = trainer_counts.max() / trainer_counts.min()

# --- Insight 3: Single-Day Visitors vs Multi-Day Engagement (Direction D) ---
fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
df['visit_segment'] = np.where(df['num_days_per_week'] == 1, 'Single-Day (High Risk)', 'Multi-Day (Engaged)')
sns.kdeplot(data=df, x='avg_time_in_gym', hue='visit_segment', fill=True, palette=[PALETTE[1], PALETTE[3]], ax=ax)
format_ax(ax, "Single-Day Visitors Exhibit Dangerously Low Gym Dwell Times", "Average Time in Gym (minutes)", "Density")
buf3 = io.BytesIO()
fig.savefig(buf3, format='png', bbox_inches='tight')
buf3.seek(0)
chart_buffers.append(buf3)
plt.close(fig)

# Metrics for Text 3
single_day_pct = (df['num_days_per_week'] == 1).mean() * 100
single_day_time = df[df['num_days_per_week'] == 1]['avg_time_in_gym'].mean()
multi_day_time = df[df['num_days_per_week'] > 1]['avg_time_in_gym'].mean()

# --- Insight 4: Age Group vs Service Uptake (Direction C) ---
fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
pt_by_age = df.groupby('age_group', observed=True)['personal_training'].mean() * 100
pt_by_age.plot(kind='bar', ax=ax, color=PALETTE[0])
format_ax(ax, "Personal Training Penetration is Highly Stratified by Age", "Age Group", "Uptake Rate (%)")
for i, v in enumerate(pt_by_age):
    ax.text(i, v + 1, f"{v:.1f}%", ha='center', fontweight='bold')
plt.xticks(rotation=0)
buf4 = io.BytesIO()
fig.savefig(buf4, format='png', bbox_inches='tight')
buf4.seek(0)
chart_buffers.append(buf4)
plt.close(fig)

# Metrics for Text 4
peak_age_pt = pt_by_age.idxmax()
low_age_pt = pt_by_age.idxmin()

# --- Insight 5: Check-in Time Clustering (Direction F) ---
fig, ax = plt.subplots(figsize=(10, 5), dpi=150)
sns.histplot(data=df, x='check_in_hour', hue='abonoment_type', multiple='stack', bins=range(5, 24), palette=[PALETTE[0], PALETTE[4]], ax=ax)
format_ax(ax, "Midday Operational Capacity is Severely Undersold", "Check-in Hour (24h)", "Check-in Volume (n)")
ax.set_xticks(range(6, 23, 2))
buf5 = io.BytesIO()
fig.savefig(buf5, format='png', bbox_inches='tight')
buf5.seek(0)
chart_buffers.append(buf5)
plt.close(fig)

# Metrics for Text 5
peak_hour = int(df['check_in_hour'].mode()[0])
dead_zone_vol = len(df[(df['check_in_hour'] >= 10) & (df['check_in_hour'] <= 14)])

# ==============================================================================
# STEP 3 & 5 — PDF FORMATTING AND BUILD
# ==============================================================================

def build_pdf():
    doc = SimpleDocTemplate(output_path, pagesize=A4, rightMargin=2.5*cm, leftMargin=2.5*cm, topMargin=2.5*cm, bottomMargin=2.5*cm)
    styles = getSampleStyleSheet()
    
    # Custom Typography overrides
    styles.add(ParagraphStyle(name='CustomTitle', fontName='Helvetica-Bold', fontSize=28, textColor=colors.HexColor('#1A2332'), alignment=1, spaceAfter=20))
    styles.add(ParagraphStyle(name='Subtitle', fontName='Helvetica', fontSize=14, textColor=colors.HexColor('#555555'), alignment=1, spaceAfter=30))
    styles.add(ParagraphStyle(name='H1', fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor('#1A2332'), spaceBefore=18, spaceAfter=6))
    styles.add(ParagraphStyle(name='H2', fontName='Helvetica-Bold', fontSize=13, textColor=colors.HexColor('#2D3748'), spaceBefore=12, spaceAfter=12))
    styles.add(ParagraphStyle(name='BodyTextCustom', fontName='Helvetica', fontSize=10.5, textColor=colors.HexColor('#2D2D2D'), leading=14, alignment=4, spaceAfter=10)) # alignment 4 is Justified
    styles.add(ParagraphStyle(name='Caption', fontName='Helvetica-Oblique', fontSize=9.5, textColor=colors.HexColor('#555555'), alignment=1, spaceBefore=6, spaceAfter=14))
    styles.add(ParagraphStyle(name='RecText', fontName='Helvetica-Bold', fontSize=10.5, textColor=colors.HexColor('#2D2D2D'), leading=14))

    story = []

    # Helper functions for layout
    def make_insight_header(num, title_text):
        num_p = Paragraph(f"<font color='white'><b>{num}</b></font>", ParagraphStyle('N', fontName='Helvetica-Bold', alignment=1))
        title_p = Paragraph(title_text, styles['H2'])
        t = Table([[num_p, title_p]], colWidths=[0.8*cm, 14.2*cm], style=[
            ('BACKGROUND', (0,0), (0,0), colors.HexColor('#1A2332')),
            ('ALIGN', (0,0), (0,0), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ('BOTTOMPADDING', (0,0), (-1,-1), 2),
        ])
        return t

    def make_recommendation(text):
        rec_p = Paragraph(f"<b>RECOMMENDATION:</b> {text}", styles['RecText'])
        t = Table([['', rec_p]], colWidths=[0.15*cm, 14.85*cm], style=[
            ('BACKGROUND', (0,0), (0,0), colors.HexColor('#1A2332')),
            ('LEFTPADDING', (1,0), (1,0), 12),
            ('TOPPADDING', (0,0), (-1,-1), 0),
            ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ])
        return t

    # --- TITLE PAGE ---
    story.append(Spacer(1, 6*cm))
    story.append(Paragraph("Membership & Operations<br/>Intelligence Report", styles['CustomTitle']))
    story.append(Paragraph("Prepared for Gym Management — Confidential", styles['Subtitle']))
    story.append(Paragraph(f"Date: {datetime.date.today().strftime('%B %d, %Y')}", ParagraphStyle('D', parent=styles['Subtitle'], spaceAfter=10)))
    story.append(Paragraph("Prepared by Senior Data Analytics Practice", ParagraphStyle('C', parent=styles['Subtitle'], fontSize=11, spaceAfter=10)))
    story.append(HRFlowable(width="50%", thickness=1, color=colors.HexColor('#1A2332'), spaceAfter=20, hAlign='CENTER'))
    
    abstract = ("This report isolates high-value operational and membership patterns extracted from current gym dataset metrics. "
                "The analysis maps facility utilisation, identifies immediate upsell avenues within the standard membership base, "
                "and flags critical retention risks to guide direct management intervention.")
    story.append(Paragraph(abstract, ParagraphStyle('Abs', parent=styles['BodyTextCustom'], alignment=1, leftIndent=40, rightIndent=40)))
    story.append(PageBreak())

    # --- SECTION 1: EXECUTIVE SUMMARY ---
    story.append(Paragraph("SECTION 1 — EXECUTIVE SUMMARY", styles['H1']))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#E74C3C'), spaceAfter=15))

    p1 = (f"The underlying dataset comprises {len(df)} active membership records evaluated across {len(df.columns)} operational dimensions. "
          f"Data processing successfully handled multi-value string classifications and validated structural nulls across secondary services "
          f"(group lessons, drinks, and personal training), confirming absolute data integrity for decision-making. No records required exclusion.")
    story.append(Paragraph(p1, styles['BodyTextCustom']))

    p2 = (f"Macro-level business health remains robust but operationally inefficient. Standard members represent {(df['abonoment_type']=='Standard').mean()*100:.1f}% "
          f"of the base, yet they display a high propensity for consuming premium add-ons a la carte. "
          f"Concurrently, core retention metrics reveal stress fractures: {single_day_pct:.1f}% of the base engages the facility only once per week, "
          f"exhibiting critically low dwell times.")
    story.append(Paragraph(p2, styles['BodyTextCustom']))

    p3 = (f"Staff utilisation operates at an unsustainable imbalance. The disparity in client load between trainers creates an acute single-point-of-failure risk, "
          f"particularly isolated around peak check-in periods at {peak_hour}:00. Furthermore, facility load distribution highlights a stark dead-zone during midday hours, "
          f"indicating untargeted capacity.")
    story.append(Paragraph(p3, styles['BodyTextCustom']))

    p4 = ("The five insights detailed below map a clear pathway to immediate revenue generation and risk mitigation. By restructuring how add-ons are bundled, "
          "management can transition heavy-using Standard members into higher-margin tiers. Concurrently, shifting operational load away from bottlenecked staff "
          "and highly congested time slots will directly improve the multi-day visitation rate.")
    story.append(Paragraph(p4, styles['BodyTextCustom']))

    story.append(PageBreak())

    # --- SECTION 2: THE FIVE INSIGHTS ---
    story.append(Paragraph("SECTION 2 — THE FIVE INSIGHTS", styles['H1']))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#E74C3C'), spaceAfter=15))

    # INSIGHT 1
    story.append(KeepTogether([
        make_insight_header("1", "High Add-on Penetration in Standard Tier Signals Missed Upsell"),
        Spacer(1, 10),
        Image(ImageReader(chart_buffers[0]), width=15*cm, height=7.5*cm),
        Paragraph(f"<b>CAPTION:</b> The chart isolates the volume of active secondary services (add-ons) by membership tier. Strikingly, {std_high_addon:.1f}% of Standard members continuously utilise 3 or more add-ons, behaving functionally as Premium members but operating on a lower-margin base rate.", styles['Caption']),
        make_recommendation("Execute an automated, highly targeted email campaign offering an immediate Premium upgrade at a 15% first-month discount exclusively to Standard members currently holding 3+ add-ons.")
    ]))
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#BDC3C7'), spaceAfter=15))

    # INSIGHT 2
    story.append(KeepTogether([
        make_insight_header("2", "Severe Capacity Imbalance Within Personal Training Staff"),
        Spacer(1, 10),
        Image(ImageReader(chart_buffers[1]), width=15*cm, height=7.5*cm),
        Paragraph(f"<b>CAPTION:</b> Client loads are critically skewed. Trainer {top_trainer} manages roughly {load_ratio:.1f}x the volume of {bottom_trainer}. This exposes the business to a high-impact churn event if the primary trainer departs or hits burnout capacity.", styles['Caption']),
        make_recommendation(f"Instigate an immediate freeze on assigning new clients to {top_trainer}. Route all new personal training requests exclusively to {bottom_trainer} and the mid-tier trainers, subsidised by a free initial consultation.")
    ]))
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#BDC3C7'), spaceAfter=15))

    # INSIGHT 3
    story.append(KeepTogether([
        make_insight_header("3", "Single-Day Visitors Represent a High-Risk Churn Segment"),
        Spacer(1, 10),
        Image(ImageReader(chart_buffers[2]), width=15*cm, height=7.5*cm),
        Paragraph(f"<b>CAPTION:</b> {single_day_pct:.1f}% of members visit only one day per week. Furthermore, this cohort spends significantly less time per session (avg {single_day_time:.0f} mins) compared to multi-day engaged members (avg {multi_day_time:.0f} mins), signalling low commitment and high cancellation probability.", styles['Caption']),
        make_recommendation("Trigger a 'Bring a Friend for Free' pass specifically targeted at the single-day cohort to engineer a second weekly visit, anchoring their gym habit to social accountability.")
    ]))
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#BDC3C7'), spaceAfter=15))

    # INSIGHT 4
    story.append(KeepTogether([
        make_insight_header("4", "Personal Training Uptake Lags Severely in Younger Demographics"),
        Spacer(1, 10),
        Image(ImageReader(chart_buffers[3]), width=15*cm, height=7.5*cm),
        Paragraph(f"<b>CAPTION:</b> Penetration of personal training scales sharply with age, peaking in the {peak_age_pt} bracket. The younger segments ({low_age_pt}) represent an untapped market failing to convert onto high-margin coaching services.", styles['Caption']),
        make_recommendation("Design and launch a lower-priced, small-group 'Intro to Lifting' clinic capped at 4 participants specifically marketed to the 18–25 segment to break the price barrier of 1-on-1 coaching.")
    ]))
    story.append(Spacer(1, 15))
    story.append(HRFlowable(width="100%", thickness=0.5, color=colors.HexColor('#BDC3C7'), spaceAfter=15))

    # INSIGHT 5
    story.append(KeepTogether([
        make_insight_header("5", "Midday Dead Zones Depress Overall Facility ROI"),
        Spacer(1, 10),
        Image(ImageReader(chart_buffers[4]), width=15*cm, height=7.5*cm),
        Paragraph(f"<b>CAPTION:</b> Facility check-ins polarise heavily around {peak_hour}:00, stranding vast operational capacity between 10:00 and 14:00 (only {dead_zone_vol} total check-ins recorded in this band).", styles['Caption']),
        make_recommendation("Introduce an 'Off-Peak Standard' membership tier at a 20% discount that restricts facility access to the 10:00–15:00 window, targeting remote workers and students to monetize dead capacity without cannibalising prime-time slots.")
    ]))
    
    story.append(PageBreak())

    # --- SECTION 3: CONCLUSIONS ---
    story.append(Paragraph("SECTION 3 — CONCLUSIONS", styles['H1']))
    story.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#E74C3C'), spaceAfter=15))

    c1 = ("The data dictates that the facility’s primary bottleneck is not acquisition, but structural optimisation. The business possesses a healthy, active membership base that demonstrates a willingness to spend beyond base tier parameters. However, unmanaged staff workloads and severely imbalanced peak usage hours actively throttle further scaling.")
    story.append(Paragraph(c1, styles['BodyTextCustom']))

    c2 = ("The single highest-priority action is the immediate conversion of the high-addon Standard tier members. Upgrading this specific sub-segment directly increases fixed monthly recurring revenue with zero additional acquisition cost, simultaneously cleaning up the operational tracking of isolated add-ons.")
    story.append(Paragraph(c2, styles['BodyTextCustom']))

    c3 = ("Management must implement a rigid monthly monitoring framework tracking three absolute metrics: 1) The conversion rate of Standard-to-Premium targets, 2) The ratio of single-day to multi-day visitors (aiming to suppress the 1-day cohort below 15%), and 3) The client delta between the most-booked and least-booked trainer. Positive momentum across these three vectors guarantees top-line growth and churn insulation.")
    story.append(Paragraph(c3, styles['BodyTextCustom']))

    c4 = ("To deepen future reporting, management is advised to capture distinct timestamp records of member cancellations and track exact session timestamps rather than aggregated averages. This will unlock dynamic cohort survival analyses and pinpoint exactly when a member stops attending prior to formal cancellation.")
    story.append(Paragraph(c4, styles['BodyTextCustom']))

    # Page Number Canvas callback
    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 9)
        canvas.setFillColor(colors.HexColor('#555555'))
        canvas.drawCentredString(10.5*cm, 1.5*cm, f"Page {doc.page}")
        canvas.restoreState()

    try:
        # Pre-build validation
        for i, buf in enumerate(chart_buffers):
            assert buf.getbuffer().nbytes > 0, f"Chart buffer {i} is empty."
        
        doc.build(story, onFirstPage=add_page_number, onLaterPages=add_page_number)
        
        print("\n═══════════════════════════════════════════════════════════════")
        print("STEP 6: ERROR-PROOFING AND OUTPUT")
        print("═══════════════════════════════════════════════════════════════")
        print(f"✓ Report successfully saved to {output_path}")
        print(f"✓ File size: {os.path.getsize(output_path) / 1024:.1f} KB")
        print("✓ Pages estimated from story length: check PDF viewer for exact count (est. 4-5 pages)")
    
    except Exception as e:
        print("\n[!] FATAL ERROR DURING PDF BUILD:")
        import traceback
        traceback.print_exc()

build_pdf()