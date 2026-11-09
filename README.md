# Job-seeking-pattern-among-local-youth-field-project

I created a project of the students or the people of age between 18 and 29 overall, collected about 75 responses, studied them, and derived key findings from them. This repository contains the complete dataset analysis, data cleanings, 12 high-resolution visualizations, generation scripts, and a professionally formatted Microsoft Word research report.

---

## 📁 Repository Contents

- **`Field_Project_Survey_Report.docx`**: The complete formatted Microsoft Word research report featuring:
  - Custom navy page borders
  - Project executive summary and methodology
  - 12 comprehensive analytical chapters with structured tables
  - Embedded high-resolution survey visualizations
  - Strategic takeaways, cross-analysis, and actionable recommendations
- **`charts/`**: High-resolution (300 DPI) publication-ready charts (Figures 1 to 12):
  - `q1_age.png` - Age Distribution ($N = 75$)
  - `q2_education.png` - Educational Qualifications ($N = 72$)
  - `q3_employment.png` - Current Employment Status ($N = 75$)
  - `q4_methods.png` - Job Search Channels & Sourcing Methods
  - `q5_factors.png` - Career & Job Decision Factors Matrix
  - `q6_challenges.png` - Major Job Search Obstacles ($N = 72$)
  - `q7_working_mode.png` - Workplace Arrangement Preferences ($N = 65$)
  - `q8_satisfaction.png` - Job Satisfaction Ratings ($N = 75$)
  - `q9_portals.png` - Job Search Platform & Portal Market Share ($N = 75$)
  - `q10_sectors.png` - Target Industry / Employment Sectors ($N = 74$)
  - `q11_clarity.png` - Perceived Clarity of Application Workflows ($N = 72$)
  - `q12_interview.png` - Preferred Interview Assessment Formats ($N = 75$)
- **`build_full_report.py`**: Python script using `python-docx` and XML manipulation to generate the Word report with custom borders, callout boxes, and formatted data tables.
- **`generate_report.py`**: Python script generating all 12 survey visualizations using `matplotlib`.

---

## 🛠️ Requirements & Installation

- Python 3.8+
- `matplotlib`
- `python-docx`
- `numpy`

```bash
pip install matplotlib python-docx numpy
```

---

## 🚀 Regenerating the Project

To regenerate all charts and build the Word report from scratch:
```bash
python generate_report.py
python build_full_report.py
```
