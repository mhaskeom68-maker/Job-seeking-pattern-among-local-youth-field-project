import os
import matplotlib.pyplot as plt
import numpy as np
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

BASE_DIR = r"C:\Users\mhask\.gemini\antigravity\scratch\field_project_report"
CHARTS_DIR = os.path.join(BASE_DIR, "charts")
OUTPUT_DOCX = os.path.join(BASE_DIR, "Field_Project_Survey_Report.docx")

os.makedirs(CHARTS_DIR, exist_ok=True)

# -------------------------------------------------------------
# 1. GENERATE VISUALIZATIONS
# -------------------------------------------------------------
plt.rcParams['font.family'] = 'sans-serif'
plt.rcParams['font.sans-serif'] = ['Segoe UI', 'DejaVu Sans', 'Arial', 'Helvetica']
plt.rcParams['axes.edgecolor'] = '#D1D5DB'
plt.rcParams['axes.linewidth'] = 0.8

PRIMARY = '#1F4E79'
SECONDARY = '#2E75B6'
ACCENT_ORANGE = '#D97706'
ACCENT_RED = '#DC2626'
ACCENT_GREEN = '#16A34A'
ACCENT_TEAL = '#0D9488'
TEXT_DARK = '#1F2937'

def create_charts():
    # Q1: Age Distribution
    fig, ax = plt.subplots(figsize=(6.8, 3.8), dpi=300)
    ages = ['18–20', '21–23', '24–26', '27–29']
    counts_q1 = [46, 11, 11, 7]
    bars = ax.bar(ages, counts_q1, color=['#1F4E79', '#2E75B6', '#3B82F6', '#60A5FA'], width=0.52, zorder=3)
    ax.grid(axis='y', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 1: Age Distribution of Survey Respondents (N = 75)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Age Group', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_ylabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_ylim(0, 52)
    for bar, count in zip(bars, counts_q1):
        pct = (count / 75) * 100
        ax.annotate(f'{count} ({pct:.1f}%)',
                    xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    xytext=(0, 4), textcoords='offset points',
                    ha='center', va='bottom', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q1_age.png'), dpi=300)
    plt.close()

    # Q2: Education
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=300)
    edu = ["Bachelor's", 'High School', 'Vocational', "Master's", 'PhD', 'Other']
    counts_q2 = [40, 12, 6, 6, 5, 3]
    y_pos = np.arange(len(edu))
    bars = ax.barh(y_pos, counts_q2, color=['#1F4E79', '#2E75B6', '#3B82F6', '#60A5FA', '#93C5FD', '#CBD5E1'], height=0.58, zorder=3)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(edu, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.invert_yaxis()
    ax.grid(axis='x', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title("Figure 2: Highest Level of Education Attained (N = 72)", fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_xlim(0, 46)
    for bar, count in zip(bars, counts_q2):
        pct = (count / 72) * 100
        ax.annotate(f' {count} ({pct:.1f}%)',
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q2_education.png'), dpi=300)
    plt.close()

    # Q3: Employment Status
    fig, ax = plt.subplots(figsize=(6.2, 4.0), dpi=300)
    emp_labels = ['Student', 'Employed', 'Looking for Work', 'Unemployed']
    emp_counts = [45, 18, 10, 2]
    colors_q3 = ['#1F4E79', '#2E75B6', '#F59E0B', '#EF4444']
    wedges, texts, autotexts = ax.pie(
        emp_counts, labels=emp_labels, autopct='%1.1f%%', startangle=140,
        pctdistance=0.76, colors=colors_q3,
        wedgeprops=dict(width=0.40, edgecolor='white', linewidth=2.5)
    )
    for t in texts:
        t.set_fontsize(9)
        t.set_fontweight('bold')
        t.set_color(TEXT_DARK)
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_fontweight('bold')
        at.set_color('white')
    ax.set_title('Figure 3: Current Employment Status (N = 75)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q3_employment.png'), dpi=300)
    plt.close()

    # Q4: Methods
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=300)
    methods = ['Online Portals', 'Social Media', 'Company Website', 'Networking', 'Agencies', 'Newspaper']
    counts_q4 = [56, 41, 37, 28, 21, 18]
    y_pos = np.arange(len(methods))
    bars = ax.barh(y_pos, counts_q4, color=['#1F4E79', '#2563EB', '#3B82F6', '#60A5FA', '#93C5FD', '#BAE6FD'], height=0.58, zorder=3)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(methods, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.invert_yaxis()
    ax.grid(axis='x', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 4: Job Search Channels & Methods Used', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Total Selections (Multiple Responses Allowed)', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_xlim(0, 65)
    for bar, count in zip(bars, counts_q4):
        pct = (count / 75) * 100
        ax.annotate(f' {count} ({pct:.1f}% resp.)',
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q4_methods.png'), dpi=300)
    plt.close()

    # Q5: Factors
    fig, ax = plt.subplots(figsize=(8.0, 4.2), dpi=300)
    factors = ['Salary', 'Work-Life Balance', 'Growth Opp.', 'Company Rep.', 'Location']
    very_imp = [53, 18, 2, 1, 1]
    imp = [32, 28, 10, 2, 3]
    neutral = [41, 23, 10, 2, 1]
    unimp = [32, 20, 12, 3, 8]
    very_unimp = [30, 19, 20, 3, 3]
    x = np.arange(len(factors))
    w = 0.16
    ax.bar(x - 2*w, very_imp, width=w, label='Very Important', color='#1E3A8A', zorder=3)
    ax.bar(x - w, imp, width=w, label='Important', color='#2563EB', zorder=3)
    ax.bar(x, neutral, width=w, label='Neutral', color='#64748B', zorder=3)
    ax.bar(x + w, unimp, width=w, label='Unimportant', color='#F59E0B', zorder=3)
    ax.bar(x + 2*w, very_unimp, width=w, label='Very Unimportant', color='#EF4444', zorder=3)
    ax.set_xticks(x)
    ax.set_xticklabels(factors, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.grid(axis='y', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 5: Evaluation of Career & Job Decision Factors', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_ylabel('Response Count', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_ylim(0, 60)
    ax.legend(frameon=True, facecolor='white', edgecolor='#E2E8F0', fontsize=8, loc='upper right')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q5_factors.png'), dpi=300)
    plt.close()

    # Q6: Challenges
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=300)
    challenges = ['Lack of Experience', 'High Competition', 'Salary Expectancy', 'Few Job Openings', 'Inadequate Networking', 'Complex App Process']
    counts_q6 = [19, 18, 14, 11, 5, 5]
    y_pos = np.arange(len(challenges))
    colors_c = ['#DC2626', '#EA580C', '#D97706', '#2563EB', '#475569', '#64748B']
    bars = ax.barh(y_pos, counts_q6, color=colors_c, height=0.58, zorder=3)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(challenges, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.invert_yaxis()
    ax.grid(axis='x', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 6: Primary Challenges Encountered in Job Search (N = 72)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_xlim(0, 24)
    for bar, count in zip(bars, counts_q6):
        pct = (count / 72) * 100
        ax.annotate(f' {count} ({pct:.1f}%)',
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q6_challenges.png'), dpi=300)
    plt.close()

    # Q7: Working Mode
    fig, ax = plt.subplots(figsize=(6.2, 4.0), dpi=300)
    modes = ['Hybrid', 'Fully Remote', 'Fully On-site']
    counts_q7 = [28, 27, 10]
    colors_m = ['#1F4E79', '#0D9488', '#F59E0B']
    wedges, texts, autotexts = ax.pie(
        counts_q7, labels=modes, autopct='%1.1f%%', startangle=140,
        pctdistance=0.76, colors=colors_m,
        wedgeprops=dict(width=0.40, edgecolor='white', linewidth=2.5)
    )
    for t in texts:
        t.set_fontsize(9)
        t.set_fontweight('bold')
        t.set_color(TEXT_DARK)
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_fontweight('bold')
        at.set_color('white')
    ax.set_title('Figure 7: Preference for Working Arrangement (N = 65)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q7_working_mode.png'), dpi=300)
    plt.close()

    # Q8: Job Satisfaction
    fig, ax = plt.subplots(figsize=(6.8, 3.8), dpi=300)
    sat_labels = ['1 - Very Dissatisfied', '2 - Dissatisfied', '3 - Neutral', '4 - Satisfied', '5 - Very Satisfied']
    sat_counts = [10, 19, 20, 17, 9]
    x_pos = np.arange(len(sat_labels))
    bars = ax.bar(x_pos, sat_counts, color=['#EF4444', '#F97316', '#64748B', '#0D9488', '#16A34A'], width=0.52, zorder=3)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(sat_labels, rotation=12, ha='right', fontsize=8.5, fontweight='bold', color=TEXT_DARK)
    ax.grid(axis='y', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 8: Respondent Job Satisfaction Rating (N = 75)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_ylabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_ylim(0, 25)
    for bar, count in zip(bars, sat_counts):
        pct = (count / 75) * 100
        ax.annotate(f'{count}\n({pct:.1f}%)',
                    xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    xytext=(0, 4), textcoords='offset points',
                    ha='center', va='bottom', fontsize=8, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q8_satisfaction.png'), dpi=300)
    plt.close()

    # Q9: Job Portals
    fig, ax = plt.subplots(figsize=(7.2, 4.0), dpi=300)
    portals = ['LinkedIn', 'Naukri', 'Indeed', 'Foundit', 'Apna', 'WorkIndia', 'TimesJobs', 'Freshersworld', 'Shine']
    counts_q9 = [33, 9, 7, 6, 6, 5, 4, 4, 1]
    y_pos = np.arange(len(portals))
    bars = ax.barh(y_pos, counts_q9, color='#2563EB', height=0.58, zorder=3)
    bars[0].set_color('#1E3A8A')
    ax.set_yticks(y_pos)
    ax.set_yticklabels(portals, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.invert_yaxis()
    ax.grid(axis='x', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 9: Preferred Job Search Platforms & Portals (N = 75)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_xlim(0, 38)
    for bar, count in zip(bars, counts_q9):
        pct = (count / 75) * 100
        ax.annotate(f' {count} ({pct:.1f}%)',
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q9_portals.png'), dpi=300)
    plt.close()

    # Q10: Job Sector
    fig, ax = plt.subplots(figsize=(7.2, 3.8), dpi=300)
    sectors = ['IT / Technology', 'Education', 'Government', 'Healthcare', 'Retail', 'Manufacturing', 'Finance', 'Other']
    counts_q10 = [44, 7, 7, 6, 5, 3, 1, 1]
    y_pos = np.arange(len(sectors))
    bars = ax.barh(y_pos, counts_q10, color='#2563EB', height=0.58, zorder=3)
    bars[0].set_color('#1E3A8A')
    ax.set_yticks(y_pos)
    ax.set_yticklabels(sectors, fontsize=9, fontweight='bold', color=TEXT_DARK)
    ax.invert_yaxis()
    ax.grid(axis='x', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 10: Target Employment Sector (N = 74)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_xlabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_xlim(0, 50)
    for bar, count in zip(bars, counts_q10):
        pct = (count / 74) * 100
        ax.annotate(f' {count} ({pct:.1f}%)',
                    xy=(bar.get_width(), bar.get_y() + bar.get_height() / 2),
                    xytext=(3, 0), textcoords='offset points',
                    ha='left', va='center', fontsize=8.5, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q10_sectors.png'), dpi=300)
    plt.close()

    # Q11: Clarity of Process
    fig, ax = plt.subplots(figsize=(6.8, 3.8), dpi=300)
    clarity_levels = ['1 - Very Unclear', '2 - Unclear', '3 - Neutral', '4 - Clear', '5 - Very Clear']
    counts_q11 = [12, 10, 21, 19, 10]
    x_pos = np.arange(len(clarity_levels))
    bars = ax.bar(x_pos, counts_q11, color=['#EF4444', '#F97316', '#64748B', '#2563EB', '#16A34A'], width=0.52, zorder=3)
    ax.set_xticks(x_pos)
    ax.set_xticklabels(clarity_levels, rotation=12, ha='right', fontsize=8.5, fontweight='bold', color=TEXT_DARK)
    ax.grid(axis='y', linestyle='--', alpha=0.5, zorder=0)
    ax.set_title('Figure 11: Perceived Clarity of Application Process (N = 72)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    ax.set_ylabel('Number of Respondents', fontsize=9.5, fontweight='bold', labelpad=6, color=TEXT_DARK)
    ax.set_ylim(0, 26)
    for bar, count in zip(bars, counts_q11):
        pct = (count / 72) * 100
        ax.annotate(f'{count}\n({pct:.1f}%)',
                    xy=(bar.get_x() + bar.get_width() / 2, bar.get_height()),
                    xytext=(0, 4), textcoords='offset points',
                    ha='center', va='bottom', fontsize=8, fontweight='bold', color=PRIMARY)
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q11_clarity.png'), dpi=300)
    plt.close()

    # Q12: Interview Format
    fig, ax = plt.subplots(figsize=(6.2, 4.0), dpi=300)
    formats = ['In-Person', 'Zoom Meeting', 'Phone Interview', 'Practical Assessment']
    counts_q12 = [30, 26, 13, 6]
    colors_f = ['#1E3A8A', '#2563EB', '#0D9488', '#F59E0B']
    wedges, texts, autotexts = ax.pie(
        counts_q12, labels=formats, autopct='%1.1f%%', startangle=140,
        pctdistance=0.76, colors=colors_f,
        wedgeprops=dict(width=0.40, edgecolor='white', linewidth=2.5)
    )
    for t in texts:
        t.set_fontsize(9)
        t.set_fontweight('bold')
        t.set_color(TEXT_DARK)
    for at in autotexts:
        at.set_fontsize(8.5)
        at.set_fontweight('bold')
        at.set_color('white')
    ax.set_title('Figure 12: Preferred Interview Assessment Formats (N = 75)', fontsize=11, fontweight='bold', pad=12, color=PRIMARY)
    plt.tight_layout()
    plt.savefig(os.path.join(CHARTS_DIR, 'q12_interview.png'), dpi=300)
    plt.close()
    print("Charts generated successfully.")

create_charts()
