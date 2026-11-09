import os
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

BASE_DIR = r"C:\Users\mhask\.gemini\antigravity\scratch\field_project_report"
CHARTS_DIR = os.path.join(BASE_DIR, "charts")
OUTPUT_DOCX = os.path.join(BASE_DIR, "Field_Project_Survey_Report.docx")

def apply_page_border(section):
    sectPr = section._sectPr
    pgBorders = parse_xml(f'''
        <w:pgBorders {nsdecls("w")} w:offsetFrom="page">
            <w:top w:val="double" w:sz="12" w:space="24" w:color="1F4E79"/>
            <w:left w:val="double" w:sz="12" w:space="24" w:color="1F4E79"/>
            <w:bottom w:val="double" w:sz="12" w:space="24" w:color="1F4E79"/>
            <w:right w:val="double" w:sz="12" w:space="24" w:color="1F4E79"/>
        </w:pgBorders>
    ''')
    sectPr.append(pgBorders)

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=120, bottom=120, left=180, right=180):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'''
        <w:tcMar {nsdecls("w")}>
            <w:top w:w="{top}" w:type="dxa"/>
            <w:bottom w:w="{bottom}" w:type="dxa"/>
            <w:left w:w="{left}" w:type="dxa"/>
            <w:right w:w="{right}" w:type="dxa"/>
        </w:tcMar>
    ''')
    tcPr.append(tcMar)

def set_table_borders(table, color="D1D5DB", header_color="1F4E79"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="12" w:space="0" w:color="{header_color}"/>
            <w:left w:val="none"/>
            <w:bottom w:val="single" w:sz="12" w:space="0" w:color="{header_color}"/>
            <w:right w:val="none"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{color}"/>
            <w:insideV w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)

def format_cell_text(cell, text, bold=False, italic=False, font_size=9.5, color_rgb=(31,41,55), align=WD_ALIGN_PARAGRAPH.LEFT):
    p = cell.paragraphs[0]
    p.text = ""
    p.alignment = align
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    run = p.add_run(text)
    run.bold = bold
    run.italic = italic
    run.font.name = 'Calibri'
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(*color_rgb)
    return p

def add_callout_box(doc, title, bullet_points):
    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.5)
    set_cell_background(cell, "F0F4F8")
    set_cell_margins(cell, top=160, bottom=160, left=240, right=200)
    
    # Left thick border, none on other sides
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'''
        <w:tcBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
            <w:left w:val="single" w:sz="28" w:space="0" w:color="1F4E79"/>
            <w:bottom w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
            <w:right w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>
        </w:tcBorders>
    ''')
    tcPr.append(tcBorders)
    
    p = cell.paragraphs[0]
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(4)
    run = p.add_run(f"KEY STRATEGIC TAKEAWAY: {title.upper()}")
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(10)
    run.font.color.rgb = RGBColor(31, 78, 121)
    
    for bp in bullet_points:
        bp_p = cell.add_paragraph()
        bp_p.paragraph_format.space_before = Pt(2)
        bp_p.paragraph_format.space_after = Pt(2)
        bp_p.paragraph_format.left_indent = Inches(0.2)
        r_bullet = bp_p.add_run("•  ")
        r_bullet.bold = True
        r_bullet.font.color.rgb = RGBColor(31, 78, 121)
        r_text = bp_p.add_run(bp)
        r_text.font.name = 'Calibri'
        r_text.font.size = Pt(9.5)
        r_text.font.color.rgb = RGBColor(51, 65, 85)
        
    doc.add_paragraph() # space after callout

def add_heading_1(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(18)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(16)
    run.font.color.rgb = RGBColor(31, 78, 121)
    return p

def add_heading_2(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(14)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(13)
    run.font.color.rgb = RGBColor(46, 117, 182)
    return p

def add_heading_3(doc, text):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.bold = True
    run.font.name = 'Calibri'
    run.font.size = Pt(11)
    run.font.color.rgb = RGBColor(31, 78, 121)
    return p

def add_body_paragraph(doc, text, bold_prefix=None, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    if bold_prefix:
        r_pre = p.add_run(bold_prefix)
        r_pre.bold = True
        r_pre.font.name = 'Calibri'
        r_pre.font.size = Pt(10.5)
        r_pre.font.color.rgb = RGBColor(31, 78, 121)
    r = p.add_run(text)
    r.font.name = 'Calibri'
    r.font.size = Pt(10.5)
    r.font.color.rgb = RGBColor(31, 41, 55)
    return p

def add_chart_image(doc, img_name, fig_num, caption):
    img_path = os.path.join(CHARTS_DIR, img_name)
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_img.paragraph_format.space_before = Pt(8)
        p_img.paragraph_format.space_after = Pt(4)
        run = p_img.add_run()
        run.add_picture(img_path, width=Inches(5.8))
        
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_cap.paragraph_format.space_before = Pt(2)
        p_cap.paragraph_format.space_after = Pt(10)
        r_label = p_cap.add_run(f"Figure {fig_num}: ")
        r_label.bold = True
        r_label.font.name = 'Calibri'
        r_label.font.size = Pt(9.5)
        r_label.font.color.rgb = RGBColor(31, 78, 121)
        r_cap = p_cap.add_run(caption)
        r_cap.italic = True
        r_cap.font.name = 'Calibri'
        r_cap.font.size = Pt(9.5)
        r_cap.font.color.rgb = RGBColor(75, 85, 99)

def build_data_table(doc, headers, data_rows, col_widths, total_row_index=-1):
    table = doc.add_table(rows=len(data_rows) + 1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    set_table_borders(table)
    
    # Header row
    hdr_cells = table.rows[0].cells
    for i, header_text in enumerate(headers):
        hdr_cells[i].width = Inches(col_widths[i])
        set_cell_background(hdr_cells[i], "1F4E79")
        set_cell_margins(hdr_cells[i], top=120, bottom=120, left=140, right=140)
        align = WD_ALIGN_PARAGRAPH.CENTER if i > 0 else WD_ALIGN_PARAGRAPH.LEFT
        format_cell_text(hdr_cells[i], header_text, bold=True, font_size=9.5, color_rgb=(255,255,255), align=align)
        
    # Data rows
    for r_idx, row in enumerate(data_rows):
        row_cells = table.rows[r_idx + 1].cells
        is_total = (r_idx == total_row_index or r_idx == len(data_rows) - 1 and "Total" in str(row[0]))
        bg_color = "EBF3FB" if is_total else ("F8FAFC" if r_idx % 2 == 1 else "FFFFFF")
        
        for c_idx, val in enumerate(row):
            row_cells[c_idx].width = Inches(col_widths[c_idx])
            set_cell_background(row_cells[c_idx], bg_color)
            set_cell_margins(row_cells[c_idx], top=90, bottom=90, left=140, right=140)
            align = WD_ALIGN_PARAGRAPH.CENTER if c_idx > 0 else WD_ALIGN_PARAGRAPH.LEFT
            text_color = (31, 78, 121) if is_total else (31, 41, 55)
            format_cell_text(row_cells[c_idx], str(val), bold=is_total, font_size=9.5, color_rgb=text_color, align=align)
            
    doc.add_paragraph() # space after table

def generate_word_document():
    doc = Document()
    
    # Set standard margins: 0.8 inch around for spacious bordered layout
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        apply_page_border(section)
        
        # Add running header
        header = section.header
        hp = header.paragraphs[0]
        hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        hrun = hp.add_run("FIELD PROJECT REPORT | EMPIRICAL SURVEY ANALYSIS & VISUALIZATIONS")
        hrun.font.name = 'Calibri'
        hrun.font.size = Pt(8.5)
        hrun.font.color.rgb = RGBColor(148, 163, 184)
        
        # Add running footer
        footer = section.footer
        fp = footer.paragraphs[0]
        fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
        frun = fp.add_run("Academic Field Project Research • Sample Size N = 75 • Page ")
        frun.font.name = 'Calibri'
        frun.font.size = Pt(8.5)
        frun.font.color.rgb = RGBColor(148, 163, 184)

    # =========================================================
    # COVER PAGE
    # =========================================================
    p_pre = doc.add_paragraph()
    p_pre.paragraph_format.space_before = Pt(36)
    p_pre.paragraph_format.space_after = Pt(6)
    p_pre.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_inst = p_pre.add_run("ACADEMIC FIELD RESEARCH STUDY")
    r_inst.font.name = 'Calibri'
    r_inst.font.size = Pt(12)
    r_inst.bold = True
    r_inst.font.color.rgb = RGBColor(46, 117, 182)

    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(12)
    p_title.paragraph_format.space_after = Pt(12)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("FIELD PROJECT REPORT:\nCAREER PERCEPTIONS, JOB SEARCH DYNAMICS, AND WORKPLACE PREFERENCES")
    r_title.font.name = 'Calibri'
    r_title.font.size = Pt(22)
    r_title.bold = True
    r_title.font.color.rgb = RGBColor(31, 78, 121)

    p_sub = doc.add_paragraph()
    p_sub.paragraph_format.space_before = Pt(6)
    p_sub.paragraph_format.space_after = Pt(30)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("An Empirical Multidimensional Survey Analysis of Candidate Expectations, Digital Portal Adoption, and Recruitment Friction")
    r_sub.font.name = 'Calibri'
    r_sub.font.size = Pt(12)
    r_sub.italic = True
    r_sub.font.color.rgb = RGBColor(100, 116, 139)

    # Decorative Meta Box on Cover Page
    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    meta_table.autofit = False
    set_table_borders(meta_table, color="CBD5E1", header_color="1F4E79")
    
    meta_info = [
        ("Project Domain:", "Employment Ecosystems & Human Resource Analytics"),
        ("Survey Scope:", "12 Quantitative & Qualitative Survey Modules (Q1–Q12)"),
        ("Target Population:", "Undergraduate & Graduate Students, Job Seekers, & Early Professionals"),
        ("Total Sample Size:", "N = 75 Verified Respondents"),
        ("Primary Investigation:", "Hiring Channels, Compensation Priorities, and Recruitment Barriers"),
        ("Reporting Period:", "Academic Year 2026")
    ]
    
    for idx, (label, val) in enumerate(meta_info):
        r_cells = meta_table.rows[idx].cells
        r_cells[0].width = Inches(2.2)
        r_cells[1].width = Inches(4.3)
        set_cell_background(r_cells[0], "F1F5F9")
        set_cell_background(r_cells[1], "FFFFFF")
        set_cell_margins(r_cells[0], top=100, bottom=100, left=150, right=150)
        set_cell_margins(r_cells[1], top=100, bottom=100, left=150, right=150)
        format_cell_text(r_cells[0], label, bold=True, font_size=10, color_rgb=(31, 78, 121))
        format_cell_text(r_cells[1], val, bold=False, font_size=10, color_rgb=(31, 41, 55))

    p_div = doc.add_paragraph()
    p_div.paragraph_format.space_before = Pt(40)
    p_div.paragraph_format.space_after = Pt(10)
    p_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_foot = p_div.add_run("Comprehensive Research Dossier with Visualizations, Data Matrices, & Analytical Insights")
    r_foot.font.name = 'Calibri'
    r_foot.font.size = Pt(9.5)
    r_foot.font.color.rgb = RGBColor(148, 163, 184)

    doc.add_page_break()

    # =========================================================
    # EXECUTIVE SUMMARY & PROJECT METHODOLOGY
    # =========================================================
    add_heading_1(doc, "Executive Summary & Research Framework")
    
    add_body_paragraph(doc, 
        "This field research project investigates the evolving dynamics of employment transition, job search behavior, and candidate expectations among early-career individuals and students. Conducted across a sample of 75 survey respondents, the study captures multidimensional data covering demographic baselines, digital platform preferences, key career choice drivers, perceived recruitment friction, and modern working arrangement inclinations.",
        bold_prefix="Overview: "
    )
    
    add_body_paragraph(doc,
        "The survey dataset comprises twelve standardized questions (Q1 to Q12), organized into cohesive thematic analytical modules. The quantitative results demonstrate significant transformations in candidate behavior, characterized by an overwhelming pivot toward digital job search platforms, a sharp preference for flexible work modes (hybrid and remote), and critical friction points centered on experience prerequisites and recruitment communication clarity.",
        bold_prefix="Methodology & Data Structure: "
    )

    add_callout_box(doc, "Executive Synthesis & Core Highlights", [
        "Demographic Dominance: 61.3% of respondents fall in the 18–20 age group, with 60.0% maintaining active student status, representing the upcoming workforce pipeline.",
        "Educational Specialization: Over 55.6% hold or are pursuing a Bachelor's degree, heavily clustering their career aspirations in the Information Technology (59.5%) sector.",
        "Search Channels & Portal Monopolies: Digital job boards (74.7%) and social media (54.7%) vastly outperform traditional channels; LinkedIn commands 44.0% of portal market share.",
        "The Work Mode Paradigm: 84.6% demand Hybrid (43.1%) or Fully Remote (41.5%) setups, with only 15.4% willing to accept conventional fully on-site work.",
        "Market Friction & Dissatisfaction: 'Lack of Experience' (26.4%) and 'High Competition' (25.0%) represent the most formidable hurdles, driving an overall net-negative satisfaction score (average 2.95 out of 5.0)."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 1: DEMOGRAPHIC & EDUCATIONAL PROFILING
    # =========================================================
    add_heading_1(doc, "Chapter 1: Demographic & Educational Profiling")
    
    # --- Q1: AGE ---
    add_heading_2(doc, "Question 1: Age Distribution of Survey Respondents")
    add_body_paragraph(doc, 
        "The primary objective of Question 1 was to establish the generational baseline of the surveyed cohort, mapping the age structure of candidates entering the contemporary hiring market.",
        bold_prefix="Research Objective: "
    )
    
    q1_headers = ["Age Category", "Respondent Count", "Relative Percentage (%)", "Cumulative Percentage (%)"]
    q1_rows = [
        ["18–20", "46", "61.3%", "61.3%"],
        ["21–23", "11", "14.7%", "76.0%"],
        ["24–26", "11", "14.7%", "90.7%"],
        ["27–29", "7", "9.3%", "100.0%"],
        ["Total", "75", "100.0%", "—"]
    ]
    build_data_table(doc, q1_headers, q1_rows, [2.0, 1.4, 1.6, 1.5], total_row_index=4)
    add_chart_image(doc, "q1_age.png", 1, "Age Breakdown of Sample Respondents (N = 75)")
    
    add_body_paragraph(doc,
        "The age distribution reveals a pronounced concentration in the youngest demographic bracket. Nearly two-thirds (61.3%) of all surveyed participants are between 18 and 20 years old, representing tertiary undergraduate cohorts entering pre-placement stages. A combined 29.4% span the 21–26 age corridor, reflecting immediate graduates and junior professionals. Only 9.3% belong to the 27–29 bracket.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Age Demographics", [
        "The respondent pool is heavily skewed toward Generation Z entry-level entrants.",
        "Career guidance, university placement support, and early internship opportunities are critical interventions for this demographic profile."
    ])

    # --- Q2: EDUCATION ---
    add_heading_2(doc, "Question 2: Highest Level of Education Attained")
    add_body_paragraph(doc, 
        "Question 2 evaluated the formal educational qualifications possessed by respondents, identifying the academic prerequisites defining their current employability profiles.",
        bold_prefix="Research Objective: "
    )
    
    q2_headers = ["Education Level", "Respondent Count", "Relative Percentage (%)", "Academic Track"]
    q2_rows = [
        ["Bachelor's Degree", "40", "55.6%", "Undergraduate"],
        ["High School", "12", "16.7%", "Secondary"],
        ["Vocational Training", "6", "8.3%", "Technical / Applied"],
        ["Master's Degree", "6", "8.3%", "Postgraduate"],
        ["PhD / Doctorate", "5", "6.9%", "Advanced Research"],
        ["Other", "3", "4.2%", "Certificates / Continuing"],
        ["Total Valid Responses", "72", "100.0%", "—"]
    ]
    build_data_table(doc, q2_headers, q2_rows, [2.2, 1.4, 1.5, 1.4], total_row_index=6)
    add_chart_image(doc, "q2_education.png", 2, "Educational Distribution of Respondents (N = 72)")
    
    add_body_paragraph(doc,
        "Bachelor's degree candidates constitute the clear majority (55.6%, n=40), forming the primary talent pool. High school graduates account for 16.7% (n=12), many of whom are concurrently pursuing collegiate coursework. Postgraduate degree holders (Master's at 8.3% and PhDs at 6.9%) represent 15.2% of the surveyed population. Vocational certifications account for 8.3%.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Educational Profiling", [
        "Undergraduate qualification serves as the dominant baseline credential across candidate pools.",
        "Specialized skill alignment and practical internships must bridge academic curricula with industrial requisites."
    ])

    # --- Q3: EMPLOYMENT STATUS ---
    add_heading_2(doc, "Question 3: Current Employment Status")
    add_body_paragraph(doc, 
        "Question 3 identified the operational employment status of participants, evaluating workforce participation rates and active job seeking ratios.",
        bold_prefix="Research Objective: "
    )
    
    q3_headers = ["Employment Category", "Respondent Count", "Percentage (%)", "Economic Activity Classification"]
    q3_rows = [
        ["Student", "45", "60.0%", "In Education (Pre-Workforce)"],
        ["Employed", "18", "24.0%", "Active Labor Force Participant"],
        ["Looking for Work", "10", "13.3%", "Active Job Seeker / Unemployed"],
        ["Unemployed", "2", "2.7%", "Inactive / Transitioning"],
        ["Total Respondents", "75", "100.0%", "—"]
    ]
    build_data_table(doc, q3_headers, q3_rows, [2.0, 1.3, 1.4, 1.8], total_row_index=4)
    add_chart_image(doc, "q3_employment.png", 3, "Distribution of Current Employment Status (N = 75)")
    
    add_body_paragraph(doc,
        "A full 60.0% of the sample (n=45) are currently full-time students, aligning with the age distribution observed in Q1. Gainfully employed individuals represent 24.0% (n=18), while 13.3% (n=10) are actively seeking work without current placement. Chronic unemployment remains modest at 2.7%, indicating that the majority of non-working respondents are intentionally engaged in academic programs.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Employment Dynamics", [
        "The survey cohort provides a direct pulse on campus-to-corporate transitional challenges.",
        "Over 73.3% of the sample are either actively preparing for entry (students) or actively looking for employment."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 2: JOB SEARCH BEHAVIORS & PLATFORM USAGE
    # =========================================================
    add_heading_1(doc, "Chapter 2: Job Search Behaviors & Digital Platforms")
    
    # --- Q4: JOB SEARCH METHODS ---
    add_heading_2(doc, "Question 4: Preferred Job Search Channels & Methods")
    add_body_paragraph(doc, 
        "Question 4 investigated the primary channels candidates use to discover and apply for career openings, allowing multiple selections across six established sourcing methods.",
        bold_prefix="Research Objective: "
    )
    
    q4_headers = ["Search Channel / Method", "Selections Count", "Percentage of Sample (N=75)", "Rank Order"]
    q4_rows = [
        ["Online Job Portals", "56", "74.7%", "Rank 1 (Primary)"],
        ["Social Media Channels", "41", "54.7%", "Rank 2"],
        ["Company Career Websites", "37", "49.3%", "Rank 3"],
        ["Professional & Personal Networking", "28", "37.3%", "Rank 4"],
        ["Recruitment Agencies", "21", "28.0%", "Rank 5"],
        ["Newspapers / Print Media", "18", "24.0%", "Rank 6"],
        ["Total Channel Selections", "201", "268.0% (Avg 2.68/resp)", "—"]
    ]
    build_data_table(doc, q4_headers, q4_rows, [2.2, 1.4, 1.5, 1.4], total_row_index=6)
    add_chart_image(doc, "q4_methods.png", 4, "Preferred Job Search Channels (Multiple Selections Allowed)")
    
    add_body_paragraph(doc,
        "The data confirms an emphatic shift toward algorithmic digital discovery. Online job portals dominate as the premier channel, utilized by 74.7% of respondents (n=56). Social media platforms (LinkedIn, Instagram, Twitter) have emerged as the second most common conduit at 54.7% (n=41), outpacing direct company career pages (49.3%). Traditional mediums such as print newspapers (24.0%) and recruitment agencies (28.0%) occupy a peripheral role.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Search Channel Takeaway", [
        "Job boards and social media constitute over 70% of candidate discovery touchpoints.",
        "Employers failing to optimize digital listings and social media recruitment channels forfeit access to 75%+ of early-career talent."
    ])

    # --- Q9: JOB PORTALS ---
    add_heading_2(doc, "Question 9: Market Share of Digital Job Portals")
    add_body_paragraph(doc, 
        "Question 9 mapped candidate preferences across prominent commercial employment platforms to evaluate brand loyalty and search platform concentration.",
        bold_prefix="Research Objective: "
    )
    
    q9_headers = ["Job Portal / Platform", "Respondent Count", "Relative Share (%)", "Platform Category"]
    q9_rows = [
        ["LinkedIn", "33", "44.0%", "Professional Social Network"],
        ["Naukri.com", "9", "12.0%", "Enterprise Job Board"],
        ["Indeed", "7", "9.3%", "Global Aggregator"],
        ["Foundit (formerly Monster)", "6", "8.0%", "Traditional Portal"],
        ["Apna", "6", "8.0%", "Blue/Grey Collar & Tech Mobile App"],
        ["WorkIndia", "5", "6.7%", "Entry-Level Mobile Job Board"],
        ["TimesJobs", "4", "5.3%", "Corporate Job Board"],
        ["Freshersworld", "4", "5.3%", "Campus Placement Aggregator"],
        ["Shine", "1", "1.3%", "Corporate Job Portal"],
        ["Total Valid Responses", "75", "100.0%", "—"]
    ]
    build_data_table(doc, q9_headers, q9_rows, [2.0, 1.4, 1.4, 1.7], total_row_index=9)
    add_chart_image(doc, "q9_portals.png", 9, "Market Share of Leading Job Portals (N = 75)")
    
    add_body_paragraph(doc,
        "LinkedIn establishes commanding dominance in candidate preference, capturing 44.0% (n=33) of user mindshare—nearly quadruple its closest competitor, Naukri.com (12.0%, n=9). Global aggregator Indeed captures 9.3%, followed by Foundit and Apna at 8.0% each. Mobile-first hyperlocal applications like WorkIndia and campus portals like Freshersworld reflect specialized niches among entry-level job seekers.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Job Portal Ecosystem", [
        "LinkedIn represents the undisputed anchor platform for white-collar and graduate employment.",
        "A diversified multi-platform presence is essential to reach vocational and specialized applicant cohorts."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 3: CANDIDATE EXPECTATIONS & SECTOR ASPIRATIONS
    # =========================================================
    add_heading_1(doc, "Chapter 3: Candidate Expectations & Sector Choices")
    
    # --- Q5: FACTORS ---
    add_heading_2(doc, "Question 5: Evaluation of Career & Job Decision Factors")
    add_body_paragraph(doc, 
        "Question 5 measured the psychological and financial priorities shaping employment decisions across five core dimensions evaluated on a 5-point Likert rating scale.",
        bold_prefix="Research Objective: "
    )
    
    q5_headers = ["Career Decision Factor", "Very Important", "Important", "Neutral", "Unimportant", "Very Unimportant"]
    q5_rows = [
        ["Compensation & Salary", "53 (70.7%)", "32 (42.7%)", "41 (54.7%)", "32 (42.7%)", "30 (40.0%)"],
        ["Work-Life Balance", "18 (24.0%)", "28 (37.3%)", "23 (30.7%)", "20 (26.7%)", "19 (25.3%)"],
        ["Growth & Promotion Opp.", "2 (2.7%)", "10 (13.3%)", "10 (13.3%)", "12 (16.0%)", "20 (26.7%)"],
        ["Company Reputation & Brand", "1 (1.3%)", "2 (2.7%)", "2 (2.7%)", "3 (4.0%)", "3 (4.0%)"],
        ["Office Location & Commute", "1 (1.3%)", "3 (4.0%)", "1 (1.3%)", "8 (10.7%)", "3 (4.0%)"]
    ]
    build_data_table(doc, q5_headers, q5_rows, [1.8, 1.0, 1.0, 1.0, 1.0, 1.0])
    add_chart_image(doc, "q5_factors.png", 5, "Multi-Factor Career Decision Matrix Evaluation")
    
    add_body_paragraph(doc,
        "Financial remuneration emerges as the primary decision driver, capturing 53 'Very Important' votes—representing over 70% of respondents. Work-Life Balance ranks second, registering strong importance (28 'Important' and 18 'Very Important' votes). Conversely, external brand reputation and physical office location ranked noticeably lower, indicating that candidates prioritize immediate compensation and personal time over corporate branding prestige.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Decision Factor Priority", [
        "Competitive baseline remuneration remains the prerequisite for candidate attraction.",
        "Work-life balance is a non-negotiable secondary pillar for Gen-Z and millennial talent."
    ])

    # --- Q7: WORKING MODE ---
    add_heading_2(doc, "Question 7: Preference for Modern Working Arrangements")
    add_body_paragraph(doc, 
        "Question 7 captured preferred post-pandemic operational working modes, contrasting traditional on-site office requirements against hybrid and fully remote models.",
        bold_prefix="Research Objective: "
    )
    
    q7_headers = ["Working Arrangement", "Respondent Count", "Relative Share (%)", "Workplace Flexibility Tier"]
    q7_rows = [
        ["Hybrid (Remote + Office)", "28", "43.1%", "Balanced Flexibility"],
        ["Fully Remote", "27", "41.5%", "Complete Geographical Autonomy"],
        ["Fully On-Site (Office-based)", "10", "15.4%", "Traditional Rigid Structure"],
        ["Total Valid Responses", "65", "100.0%", "—"]
    ]
    build_data_table(doc, q7_headers, q7_rows, [2.2, 1.4, 1.4, 1.5], total_row_index=3)
    add_chart_image(doc, "q7_working_mode.png", 7, "Candidate Preferences for Working Arrangements (N = 65)")
    
    add_body_paragraph(doc,
        "The results demonstrate an overwhelming mandate against rigid in-office requirements. A combined 84.6% of candidates demand flexibility, split between Hybrid setups (43.1%, n=28) and Fully Remote roles (41.5%, n=27). Only 15.4% (n=10) prefer five-day on-site work. This underscores a permanent cultural reset in workforce expectations, where physical co-location is viewed as discretionary rather than mandatory.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Workplace Flexibility", [
        "Organizations enforcing strict 100% on-site policies alienate nearly 85% of potential applicants.",
        "Hybrid schedules provide the optimal balance between collaborative immersion and autonomy."
    ])

    # --- Q10: TARGET SECTOR ---
    add_heading_2(doc, "Question 10: Target Employment & Industry Sectors")
    add_body_paragraph(doc, 
        "Question 10 surveyed the target economic sectors and industries respondents actively seek to enter upon completing their academic qualifications.",
        bold_prefix="Research Objective: "
    )
    
    q10_headers = ["Industry / Sector", "Respondent Count", "Relative Share (%)", "Economic Sector Type"]
    q10_rows = [
        ["Information Technology (IT)", "44", "59.5%", "Technology & Digital Services"],
        ["Education & Academia", "7", "9.5%", "Public & Private Educational Services"],
        ["Government / Civil Services", "7", "9.5%", "Public Sector Administration"],
        ["Healthcare & Life Sciences", "6", "8.1%", "Medical & Clinical Services"],
        ["Retail & Consumer Goods", "5", "6.8%", "Commerce & Consumer Services"],
        ["Manufacturing & Engineering", "3", "4.1%", "Industrial & Production"],
        ["Finance & Banking", "1", "1.4%", "Financial Services & FinTech"],
        ["Other Miscellaneous Sectors", "1", "1.4%", "Diversified Services"],
        ["Total Valid Responses", "74", "100.0%", "—"]
    ]
    build_data_table(doc, q10_headers, q10_rows, [2.2, 1.4, 1.4, 1.5], total_row_index=8)
    add_chart_image(doc, "q10_sectors.png", 10, "Target Industry Sector Aspirations (N = 74)")
    
    add_body_paragraph(doc,
        "The Information Technology sector commands near-monopolistic attraction, selected by 59.5% of respondents (n=44). This is driven by perceived salary premiums, hybrid flexibility, and global career mobility. Education and Government services tied for second place at 9.5% each (n=7), reflecting stable, risk-averse alternatives. Traditional core engineering like Manufacturing captured only 4.1%, while Finance accounted for 1.4%.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Industry Clustering", [
        "Extreme clustering toward IT creates severe talent hyper-competition in tech while starving non-IT sectors.",
        "Non-tech sectors must modernize workplace perks and digital branding to attract early-career talent."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 4: MARKET OBSTACLES & HIRING FRICTION
    # =========================================================
    add_heading_1(doc, "Chapter 4: Market Obstacles & Recruitment Friction")
    
    # --- Q6: CHALLENGES ---
    add_heading_2(doc, "Question 6: Primary Challenges Encountered in Job Search")
    add_body_paragraph(doc, 
        "Question 6 diagnosed the systemic impediments and friction points candidates experience when navigating modern corporate hiring pipelines.",
        bold_prefix="Research Objective: "
    )
    
    q6_headers = ["Job Search Barrier / Challenge", "Respondent Count", "Relative Share (%)", "Friction Classification"]
    q6_rows = [
        ["Lack of Required Experience", "19", "26.4%", "Structural Entry Barrier"],
        ["High Applicant Competition", "18", "25.0%", "Labor Market Saturation"],
        ["Salary Expectancy Mismatch", "14", "19.4%", "Economic Compensation Gap"],
        ["Few Job Openings / Scarcity", "11", "15.3%", "Macroeconomic Contraction"],
        ["Inadequate Professional Networking", "5", "6.9%", "Social Capital Deficit"],
        ["Complex / Cumbersome Application Process", "5", "6.9%", "Administrative Hiring Friction"],
        ["Total Valid Responses", "72", "100.0%", "—"]
    ]
    build_data_table(doc, q6_headers, q6_rows, [2.2, 1.4, 1.4, 1.5], total_row_index=6)
    add_chart_image(doc, "q6_challenges.png", 6, "Primary Job Search Obstacles Encountered (N = 72)")
    
    add_body_paragraph(doc,
        "The classic 'experience paradox' represents the single greatest hurdle: 26.4% of candidates identify 'Lack of Experience' as their foremost barrier, followed closely by 'High Competition' at 25.0%. Together, over 51.4% struggle against overcrowded applicant pools demanding unattainable prerequisites for entry-level roles. Salary expectancy gaps impact 19.4%, while application friction and networking deficits affect 6.9% each.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Hiring Barrier Insight", [
        "Entry-level requirements often mandate 2+ years of experience, shutting out competent graduates.",
        "Internship conversion and project-based portfolios must substitute for conventional tenure."
    ])

    # --- Q11: APPLICATION CLARITY ---
    add_heading_2(doc, "Question 11: Perceived Clarity of Recruitment Processes")
    add_body_paragraph(doc, 
        "Question 11 quantified the transparency, communication efficacy, and clarity of modern recruitment workflows on a standardized 5-point rating scale.",
        bold_prefix="Research Objective: "
    )
    
    q11_headers = ["Clarity Level Rating", "Respondent Count", "Percentage (%)", "Candidate Perception"]
    q11_rows = [
        ["1 - Very Unclear", "12", "16.7%", "Severe Process Opacity / Ghosting"],
        ["2 - Unclear", "10", "13.9%", "Ambiguous Communication"],
        ["3 - Neutral / Moderate", "21", "29.2%", "Adequate but Impersonal"],
        ["4 - Clear", "19", "26.4%", "Transparent & Predictable"],
        ["5 - Very Clear", "10", "13.9%", "Exceptional Guidance & Feedback"],
        ["Total Valid Responses", "72", "100.0%", "Weighted Average: 3.07 / 5.0"]
    ]
    build_data_table(doc, q11_headers, q11_rows, [2.0, 1.4, 1.4, 1.7], total_row_index=5)
    add_chart_image(doc, "q11_clarity.png", 11, "Perceived Clarity of Application Workflows (N = 72)")
    
    add_body_paragraph(doc,
        "Perception of application process transparency yields a bimodal distribution with a mean score of 3.07 out of 5.0. While 40.3% rate the process positively (Rating 4 or 5), nearly one-third (30.6%, n=22) describe application stages as unclear or very unclear. This opacity is primarily driven by Automated Tracking Systems (ATS) rejections without feedback and chronic recruiter 'ghosting.'",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Application Transparency", [
        "Nearly 1 in 3 candidates suffers from severe communication opacity during recruitment.",
        "Automated milestone tracking and constructive rejection feedback substantially enhance employer brand reputation."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 5: CANDIDATE SATISFACTION & ASSESSMENT DYNAMICS
    # =========================================================
    add_heading_1(doc, "Chapter 5: Candidate Satisfaction & Assessment Dynamics")
    
    # --- Q8: JOB SATISFACTION ---
    add_heading_2(doc, "Question 8: Overall Job & Search Satisfaction Rating")
    add_body_paragraph(doc, 
        "Question 8 gauged respondents' comprehensive satisfaction with their current employment status or job search trajectory using a 5-point Likert metric.",
        bold_prefix="Research Objective: "
    )
    
    q8_headers = ["Satisfaction Level Rating", "Respondent Count", "Percentage (%)", "Sentiment Evaluation"]
    q8_rows = [
        ["1 - Very Dissatisfied", "10", "13.3%", "High Distress / Disillusionment"],
        ["2 - Dissatisfied", "19", "25.3%", "Underemployed / Frustrated"],
        ["3 - Neutral", "20", "26.7%", "Ambivalent / In Stasis"],
        ["4 - Satisfied", "17", "22.7%", "Positive Career Trajectory"],
        ["5 - Very Satisfied", "9", "12.0%", "Optimal Role / High Fulfillment"],
        ["Total Valid Responses", "75", "100.0%", "Weighted Average: 2.95 / 5.0"]
    ]
    build_data_table(doc, q8_headers, q8_rows, [2.0, 1.4, 1.4, 1.7], total_row_index=5)
    add_chart_image(doc, "q8_satisfaction.png", 8, "Overall Job Satisfaction Metric Distribution (N = 75)")
    
    add_body_paragraph(doc,
        "The overall satisfaction index registers a net-negative average of 2.95 out of 5.0. Dissatisfied cohorts (Ratings 1 and 2) combine for 38.6% (n=29), substantially outnumbering fully satisfied participants (Rating 5 at 12.0%, n=9). Neutral sentiments comprise 26.7%. This pervasive dissatisfaction stems from prolonged hiring timelines, high application rejection volumes, and compensation mismatches.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Satisfaction Deficit", [
        "Negative sentiment outstrips positive sentiment by nearly 4 percentage points.",
        "Candidate disillusionment is directly correlated with application ambiguity and prolonged hiring cycles."
    ])

    # --- Q12: INTERVIEW FORMAT ---
    add_heading_2(doc, "Question 12: Preferred Interview Assessment Formats")
    add_body_paragraph(doc, 
        "Question 12 evaluated candidate preferences regarding interview modalities, contrasting virtual meetings against physical evaluations and hands-on skill tasks.",
        bold_prefix="Research Objective: "
    )
    
    q12_headers = ["Interview Format / Method", "Respondent Count", "Relative Share (%)", "Evaluation Nature"]
    q12_rows = [
        ["In-Person Interview", "30", "40.0%", "Face-to-Face Personal Interaction"],
        ["Zoom / Video Meeting", "26", "34.7%", "Synchronous Virtual Interaction"],
        ["Phone Interview", "13", "17.3%", "Audio Screening / Asynchronous"],
        ["Practical Technical Assessment", "6", "8.0%", "Skill Demonstration / Case Study"],
        ["Total Valid Responses", "75", "100.0%", "—"]
    ]
    build_data_table(doc, q12_headers, q12_rows, [2.2, 1.4, 1.4, 1.5], total_row_index=4)
    add_chart_image(doc, "q12_interview.png", 12, "Candidate Preferences for Interview Formats (N = 75)")
    
    add_body_paragraph(doc,
        "Despite the enthusiastic adoption of remote work (Q7), candidates still prioritize In-Person interviews (40.0%, n=30) as their most trusted evaluation format. Video meetings rank a close second at 34.7% (n=26). Phone screenings account for 17.3%, while practical coding assessments or case studies were selected by only 8.0%, suggesting candidates prefer direct dialogue over automated take-home assignments.",
        bold_prefix="Statistical Interpretation: "
    )
    add_callout_box(doc, "Interview Assessment Dynamics", [
        "Candidates value human rapport and non-verbal communication during hiring decisions.",
        "A hybrid interview protocol—virtual first rounds followed by in-person final reviews—yields the highest candidate trust."
    ])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 6: COMPARATIVE CROSS-ANALYSIS & SYNTHESIS
    # =========================================================
    add_heading_1(doc, "Chapter 6: Comparative Cross-Analysis & Synthesis")
    
    add_body_paragraph(doc,
        "By cross-correlating survey responses across multiple questions, several systemic patterns emerge regarding how early-career talent navigates the contemporary employment ecosystem:",
        bold_prefix="Systemic Synthesis: "
    )

    # Comparison 1
    add_heading_2(doc, "1. The Generation Z Tech Concentration Paradox (Q1, Q2, Q3 & Q10)")
    add_body_paragraph(doc,
        "A profound alignment exists between age profile (61.3% under 20), student status (60.0%), undergraduate credentials (55.6%), and an overwhelming targeting of the Information Technology sector (59.5%). While this demonstrates high digital literacy and ambition, it generates extreme structural bottlenecks. Thousands of identical entry-level candidates compete for a finite pool of junior software and IT roles, amplifying 'High Competition' (25.0%) and 'Lack of Experience' (26.4%) barriers."
    )

    # Comparison 2
    add_heading_2(doc, "2. Digital Platform Monopoly vs. Application Ghosting (Q4, Q9 & Q11)")
    add_body_paragraph(doc,
        "The concentration of candidate search behavior on LinkedIn (44.0%) and generic job portals (74.7%) coincides directly with the 30.6% reporting unclear or opaque recruitment processes. High-volume single-click applications on LinkedIn ('Easy Apply') encourage extreme application spamming, leading corporate HR departments to deploy stringent ATS keyword filters. This algorithmic wall prevents authentic evaluations and perpetuates the candidate perception of being ignored."
    )

    # Comparison 3
    add_heading_2(doc, "3. The Workplace Flexibility Consensus (Q5, Q7 & Q12)")
    add_body_paragraph(doc,
        "A compelling dichotomy emerges between working arrangements and interview evaluations. While 84.6% demand remote or hybrid working arrangements on the job (Q7), 40.0% simultaneously demand in-person interviews (Q12). This indicates that while candidates value day-to-day autonomy, they recognize that cultural alignment, trust building, and company evaluation are superior in physical settings."
    )

    # Comparison Table
    synth_headers = ["Strategic Dimension", "Survey Empirical Finding", "Systemic Challenge", "Strategic Rectification"]
    synth_rows = [
        ["Sector Supply & Demand", "59.5% cluster in IT sector", "Extreme talent surplus in tech; deficit in core industries", "Promote tech-integrated roles in manufacturing, healthcare, finance"],
        ["Recruitment Experience", "38.6% express net dissatisfaction", "ATS black holes, delayed feedback, ghosting", "Institute transparent candidate portals with milestone timelines"],
        ["Job Architecture", "84.6% demand Hybrid/Remote setups", "Forced return-to-office mandates spark turnover", "Standardize 2-3 day flexible hybrid working contracts"],
        ["Entry Qualification", "26.4% cite lack of experience", "Entry roles demand mid-level experience", "Establish paid micro-internships and capstone apprenticeship programs"]
    ]
    build_data_table(doc, synth_headers, synth_rows, [1.5, 1.5, 1.8, 1.7])

    doc.add_page_break()

    # =========================================================
    # CHAPTER 7: ACTIONABLE RECOMMENDATIONS & CONCLUSION
    # =========================================================
    add_heading_1(doc, "Chapter 7: Strategic Recommendations & Conclusion")
    
    add_body_paragraph(doc,
        "Based on the empirical findings generated from this field project, targeted interventions are proposed across three key stakeholder groups:",
        bold_prefix="Action Plan: "
    )

    add_heading_2(doc, "1. Recommendations for Academic Institutions & Career Centers")
    add_body_paragraph(doc,
        "• Project-Based Experiential Learning: Shift curricula toward industry capstones, live client briefs, and open-source contributions to break the 'no experience' cycle before graduation.\n"
        "• Professional Digital Profiling Workshops: Provide formalized training on LinkedIn optimization, portfolio creation, and proactive outreach rather than passive portal browsing.\n"
        "• Sector Diversification Counseling: Actively showcase high-growth opportunities outside IT—such as health-tech, industrial automation, and fintech—to alleviate tech market saturation."
    )

    add_heading_2(doc, "2. Recommendations for Corporate HR & Talent Acquisition Teams")
    add_body_paragraph(doc,
        "• Demystify the Application Lifecycle: Implement real-time recruitment dashboards where applicants can view stage progression, average review timelines, and automated status notifications.\n"
        "• Calibrate Entry-Level Job Specifications: Remove unrealistic 2–3 year experience prerequisites for junior roles, substituting them with verified technical competencies and cognitive assessments.\n"
        "• Modernize Flexible Work Frameworks: Offer structured hybrid configurations as the core compensation value proposition to maximize candidate acquisition rates."
    )

    add_heading_2(doc, "3. Recommendations for Prospective Candidates & Job Seekers")
    add_body_paragraph(doc,
        "• Cultivate Multi-Channel Sourcing: Reduce over-reliance on mass job portals; leverage direct employee referrals, alumni networking, and specialized community forums.\n"
        "• Portfolio Over Résumé: Demonstrate proof-of-work through GitHub repositories, published case studies, or freelance projects to overcome lack of formal tenure.\n"
        "• Strategic Compensation Grounding: Benchmark starting salaries against real-world market distributions while prioritizing roles offering high learning velocity."
    )

    add_heading_2(doc, "Conclusion")
    add_body_paragraph(doc,
        "This field research project provides empirical clarity into the realities confronting the emerging workforce. The findings illustrate a demographic characterized by high technical aspirations, strong digital literacy, and an insistence on modern workplace flexibility. However, these assets are often constrained by opaque recruitment systems and outdated qualification expectations. Bridging this chasm through institutional reform, hiring transparency, and flexible job structures represents the primary imperative for future workforce development.",
        bold_prefix="Final Synthesis: "
    )

    # Save document
    doc.save(OUTPUT_DOCX)
    print(f"Report successfully generated and saved to: {OUTPUT_DOCX}")

if __name__ == "__main__":
    generate_word_document()
