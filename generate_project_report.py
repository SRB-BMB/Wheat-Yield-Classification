"""
Full Report Generation Script for INT234 CA2 Project.
Student: Parth Ghuge (Reg No: omitted from public copy, Roll No: omitted from public copy, Group: omitted from public copy, Section: omitted from public copy)
Course: INT234 - Predictive Analytics
Supervisor: Course Supervisor
Semester: Project Semester 2026-27

Generates Wheat_Yield_Classification_Project_Report.docx adhering strictly to
report_format.doc and INT234 Predictive Analytics syllabus.
Compiled to PDF using LibreOffice.
Meets 30-80 pages requirement with rigorous academic, agronomic, and mathematical depth.
"""

import os
import json
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

def create_full_report(output_docx_path, results_json_path, figures_dir):
    doc = docx.Document()
    
    # Page setup - Standard A4 dimensions with 1.0 inch margins
    for section in doc.sections:
        section.top_margin = Inches(1.0)
        section.bottom_margin = Inches(1.0)
        section.left_margin = Inches(1.0)
        section.right_margin = Inches(1.0)
        section.page_width = Inches(8.27)   # A4 Width
        section.page_height = Inches(11.69) # A4 Height

    with open(results_json_path, 'r') as f:
        results = json.load(f)

    # Typography helper functions
    def add_styled_p(text="", space_after=6, space_before=0, line_spacing=1.5, align=WD_ALIGN_PARAGRAPH.LEFT, bold=False, italic=False, font_size=12, font_color=RGBColor(0,0,0)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.space_after = Pt(space_after)
        p.paragraph_format.space_before = Pt(space_before)
        p.paragraph_format.line_spacing = line_spacing
        if text:
            run = p.add_run(text)
            run.bold = bold
            run.italic = italic
            run.font.name = 'Times New Roman'
            run.font.size = Pt(font_size)
            run.font.color.rgb = font_color
        return p

    def add_h1(text):
        p = add_styled_p(text, space_after=12, space_before=20, bold=True, font_size=16, font_color=RGBColor(26, 82, 118))
        p.paragraph_format.keep_with_next = True
        return p

    def add_h2(text):
        p = add_styled_p(text, space_after=8, space_before=14, bold=True, font_size=13.5, font_color=RGBColor(44, 62, 80))
        p.paragraph_format.keep_with_next = True
        return p

    def add_h3(text):
        p = add_styled_p(text, space_after=6, space_before=10, bold=True, italic=True, font_size=12, font_color=RGBColor(52, 73, 94))
        p.paragraph_format.keep_with_next = True
        return p

    def add_body(text):
        return add_styled_p(text, space_after=6, space_before=0, line_spacing=1.5, font_size=12)

    def add_bullet(bold_prefix, text):
        p = doc.add_paragraph(style='List Bullet')
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.line_spacing = 1.5
        run_b = p.add_run(bold_prefix)
        run_b.bold = True
        run_b.font.name = 'Times New Roman'
        run_b.font.size = Pt(12)
        run_t = p.add_run(" " + text)
        run_t.font.name = 'Times New Roman'
        run_t.font.size = Pt(12)
        return p

    def add_fig(img_filename, caption, width=6.0):
        img_path = os.path.join(figures_dir, img_filename)
        if os.path.exists(img_path):
            p_img = doc.add_paragraph()
            p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_img.paragraph_format.space_before = Pt(12)
            p_img.paragraph_format.space_after = Pt(4)
            run = p_img.add_run()
            run.add_picture(img_path, width=Inches(width))
            
            p_cap = doc.add_paragraph()
            p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_cap.paragraph_format.space_after = Pt(12)
            run_cap = p_cap.add_run(caption)
            run_cap.italic = True
            run_cap.bold = True
            run_cap.font.name = 'Times New Roman'
            run_cap.font.size = Pt(10.5)
            run_cap.font.color.rgb = RGBColor(60, 60, 60)

    def set_cell_bg(cell, fill_hex):
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>'))

    def set_cell_pad(cell, top=100, bottom=100, left=140, right=140):
        tcPr = cell._tc.get_or_add_tcPr()
        tcPr.append(parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>'))

    def style_table(table, header_bg="2C3E50", alt_bg="F8F9FA"):
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        for cell in table.rows[0].cells:
            set_cell_bg(cell, header_bg)
            set_cell_pad(cell, 120, 120, 140, 140)
            for p in cell.paragraphs:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
                for r in p.runs:
                    r.bold = True
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(10.5)
                    r.font.color.rgb = RGBColor(255, 255, 255)
        for row_idx, row in enumerate(table.rows[1:], start=1):
            bg = alt_bg if row_idx % 2 == 1 else "FFFFFF"
            for cell in row.cells:
                set_cell_bg(cell, bg)
                set_cell_pad(cell, 85, 85, 120, 120)
                for p in cell.paragraphs:
                    p.paragraph_format.space_before = Pt(0)
                    p.paragraph_format.space_after = Pt(0)
                    p.paragraph_format.line_spacing = 1.15
                    for r in p.runs:
                        r.font.name = 'Times New Roman'
                        r.font.size = Pt(10.5)

    # =========================================================
    # 1. COVER PAGE
    # =========================================================
    add_styled_p("PREDICTIVE ANALYTICS", space_before=36, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=14)
    add_styled_p("PROJECT REPORT", space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=14)
    add_styled_p("(Project Semester 2026-27)", space_after=36, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)

    # 3rd line: 16 pt bold
    add_styled_p("WHEAT YIELD CATEGORY CLASSIFICATION USING FIELD, SOIL, AND PRE-SOWING INFORMATION",
                 space_before=24, space_after=36, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=16)

    add_styled_p("Submitted by", space_before=24, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("PARTH GHUGE", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=12)
    add_styled_p("Roll Number: omitted from public copy  |  Group: omitted from public copy", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("Registration No: omitted from public copy", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("Programme: B.Tech Computer Science and Engineering", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("Section: omitted from public copy", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("Course Code: INT234 (Predictive Analytics)", space_after=36, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)

    add_styled_p("Under the Guidance of", space_before=18, space_after=12, align=WD_ALIGN_PARAGRAPH.CENTER, font_size=12)
    add_styled_p("Course Supervisor", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=12)
    add_styled_p("Discipline of CSE/IT", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=12)
    add_styled_p("Lovely School of Computer Science and Engineering", space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=12)
    add_styled_p("Lovely Professional University, Phagwara", space_after=24, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, font_size=12)

    doc.add_page_break()

    # =========================================================
    # 2. DECLARATION
    # =========================================================
    add_h1("DECLARATION")
    add_body("I, Parth Ghuge, bearing an omitted registration number, student of Bachelor of Technology in Computer Science and Engineering (Section omitted from public copy, Roll No. omitted from public copy, Group omitted from public copy) under CSE/IT Discipline at Lovely Professional University, Punjab, hereby declare that all the information furnished in this project report is based on my own intensive work and is genuine. The project titled 'Wheat Yield Category Classification Using Field, Soil, and Pre-sowing Information' has been completed under the course INT234: Predictive Analytics.")
    add_body("I further declare that this report has not been submitted previously to any other university, institute, or examining body for the award of any degree, diploma, fellowship, or other similar academic title.")

    add_styled_p("", space_before=40)
    p_dec = doc.add_paragraph()
    p_dec.paragraph_format.line_spacing = 1.5
    run1 = p_dec.add_run("Date: 31st October 2026                                              Signature: ______________________\n")
    run1.font.name = 'Times New Roman'
    run1.font.size = Pt(12)
    run2 = p_dec.add_run("Registration No: omitted from public copy                                            Name of Student: Parth Ghuge")
    run2.font.name = 'Times New Roman'
    run2.font.size = Pt(12)

    doc.add_page_break()

    # =========================================================
    # 3. CERTIFICATE
    # =========================================================
    add_h1("CERTIFICATE")
    add_body("This is to certify that Parth Ghuge bearing Registration no. omitted from public copy has completed INT234 (Predictive Analytics) project titled, “Wheat Yield Category Classification Using Field, Soil, and Pre-sowing Information” under my guidance and supervision. To the best of my knowledge, the present work is the result of his original development, effort, and study.")
    add_body("The project adheres strictly to the academic requirements and syllabus boundaries of the course INT234, demonstrating the practical application of supervised classification models, exploratory data analysis, and empirical performance evaluation.")

    add_styled_p("", space_before=48)
    p_cert = doc.add_paragraph()
    p_cert.paragraph_format.line_spacing = 1.5
    cert_text = (
        "Signature and Name of the Supervisor: _________________________________\n"
        "Supervisor Name: Course Supervisor\n"
        "Discipline of Computer Science and Engineering\n"
        "Lovely School of Computer Science and Engineering\n"
        "Lovely Professional University\n"
        "Phagwara, Punjab.\n\n"
        "Date: ________________________"
    )
    r_cert = p_cert.add_run(cert_text)
    r_cert.font.name = 'Times New Roman'
    r_cert.font.size = Pt(12)

    doc.add_page_break()

    # =========================================================
    # 4. ACKNOWLEDGEMENT
    # =========================================================
    add_h1("ACKNOWLEDGEMENT")
    add_body("I would like to express my deepest gratitude to Lovely Professional University and the School of Computer Science and Engineering for providing the academic infrastructure, computing facilities, and learning environment necessary to complete this project.")
    add_body("I am profoundly indebted to my course coordinator and project supervisor, Course Supervisor, for her guidance, constant encouragement, and insightful feedback throughout the duration of this predictive analytics investigation. Her emphasis on methodological rigor, elimination of data leakage, and principled model evaluation greatly enriched this study.")
    add_body("I also acknowledge the International Maize and Wheat Improvement Center (CIMMYT) and author Anurag Ajay for curating and making publicly accessible the rich, multi-year agronomic dataset of wheat cultivation across eastern India under CC0 1.0 license. Open data initiatives such as this are vital for empowering data analytics researchers to address regional agricultural and food security challenges.")
    add_body("Finally, I extend my sincere thanks to my family, peers, and faculty members for their enduring patience, moral support, and enthusiasm during the course of my engineering studies.")

    doc.add_page_break()

    # =========================================================
    # 5. TABLE OF CONTENTS, LIST OF TABLES, LIST OF FIGURES
    # =========================================================
    add_h1("TABLE OF CONTENTS")
    
    # We populate TOC, LOT, and LOF placeholders which will be synchronized with exact PDF pages
    toc_entries = [
        ("Cover Page", "i"),
        ("Declaration", "ii"),
        ("Certificate", "iii"),
        ("Acknowledgement", "iv"),
        ("Table of Contents", "v"),
        ("List of Tables", "vii"),
        ("List of Figures", "vii"),
        ("1. Introduction", "8"),
        ("    1.1 Context and Agronomic Significance of Wheat in Eastern India", "8"),
        ("    1.2 The Agricultural Imperative for Pre-Sowing Predictive Analytics", "8"),
        ("    1.3 Problem Formulation: Yield Categorization vs. Continuous Regression", "9"),
        ("    1.4 Scoping and Framing: Wheat Classification vs. Multi-Crop Recommendation", "10"),
        ("    1.5 Alignment with INT234 Course Objectives and Syllabus Units", "10"),
        ("    1.6 Organization of the Project Report", "11"),
        ("2. Source of Dataset", "12"),
        ("    2.1 Provenance and Official CIMMYT Repository Citation", "12"),
        ("    2.2 Survey Methodology and Regional Sampling Design", "12"),
        ("    2.3 Temporal Horizon (Crop Years 2016-17 to 2021-22)", "13"),
        ("    2.4 Data Attributes and Companion Documentation", "13"),
        ("    2.5 Data Quality, Governance, and Open Access Licensing", "15"),
        ("3. EDA Process (Exploratory Data Analysis)", "16"),
        ("    3.1 Structural Overview and Initial Taxonomy of Variables", "16"),
        ("    3.2 Identification of Temporal Information Boundaries & Leakage Prevention Audit", "16"),
        ("    3.3 Systematic Missing Value Analysis and Domain-Driven Handling", "18"),
        ("    3.4 Text Standardization and Categorical Normalization", "20"),
        ("    3.5 Target Distribution Analysis and Empirical Tertile Cutoff Protocol", "21"),
        ("    3.6 Stratified 80/20 Train-Test Splitting Design", "22"),
        ("    3.7 Univariate, Bivariate, and Multivariate Relationships", "22"),
        ("    3.8 Pearson and Spearman Correlation Analysis", "23"),
        ("    3.9 Outlier Diagnosis and Handling", "25"),
        ("    3.10 Design of Scikit-Learn Preprocessing Pipelines", "25"),
        ("4. Analysis on Dataset", "26"),
        ("    4.1 Logistic Regression (Multinomial)", "26"),
        ("    4.2 Decision Tree Classifier", "28"),
        ("    4.3 Support Vector Machine (SVM)", "29"),
        ("    4.4 K-Nearest Neighbors (KNN)", "31"),
        ("    4.5 Gaussian Naive Bayes", "32"),
        ("    4.6 Comprehensive Comparative Performance Synthesis", "33"),
        ("5. Conclusion", "36"),
        ("    5.1 Synthesis of Empirical Findings", "36"),
        ("    5.2 Observational Insights for Field Management and Sowing Scheduling", "36"),
        ("    5.3 Critical Methodological Reflection: Why Accuracy Alone is Insufficient", "36"),
        ("6. Future Scope", "37"),
        ("    6.1 Temporal Validation via Future-Year Holdouts", "37"),
        ("    6.2 Integration of High-Resolution Satellite Remote Sensing", "37"),
        ("    6.3 Coupling with Short-Range Numerical Weather Prediction", "37"),
        ("    6.4 Deployment of Lightweight Decision Support Tools for Smallholders", "37"),
        ("7. References", "38")
    ]

    for title, pg in toc_entries:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.line_spacing = 1.15
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(11)
        dots_len = max(2, 85 - len(title) - len(pg))
        r_dots = p_t.add_run(" " + "." * dots_len + " ")
        r_dots.font.name = 'Times New Roman'
        r_dots.font.size = Pt(11)
        r_dots.font.color.rgb = RGBColor(150, 150, 150)
        r_p = p_t.add_run(pg)
        r_p.bold = True
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(11)

    doc.add_page_break()

    # List of Tables
    add_h1("LIST OF TABLES")
    lot_entries = [
        ("Table 2.1: Variable Definitions, Types, and Agronomic Significance", "13"),
        ("Table 3.1: Temporal Information Boundary Audit and Feature Inclusion Rationale", "16"),
        ("Table 3.2: Missing Value Audit Across Pre-Sowing Candidate Variables", "18"),
        ("Table 3.3: Cross-Tabulation of BasalNPK Application Rate vs GradeNPK Reporting", "20"),
        ("Table 3.4: Class Boundaries, Cutoff Values, and Stratified Partition Distribution", "21"),
        ("Table 3.5: Correlation Coefficients of Numeric Features with Grain Yield", "23"),
        ("Table 4.1: Logistic Regression Confusion Matrix", "27"),
        ("Table 4.2: Logistic Regression Class-Wise Performance Metrics", "27"),
        ("Table 4.3: Decision Tree Confusion Matrix", "28"),
        ("Table 4.4: Decision Tree Class-Wise Performance Metrics", "29"),
        ("Table 4.5: Support Vector Machine Confusion Matrix", "30"),
        ("Table 4.6: Support Vector Machine Class-Wise Performance Metrics", "30"),
        ("Table 4.7: K-Nearest Neighbors Confusion Matrix", "31"),
        ("Table 4.8: K-Nearest Neighbors Class-Wise Performance Metrics", "31"),
        ("Table 4.9: Gaussian Naive Bayes Confusion Matrix", "33"),
        ("Table 4.10: Gaussian Naive Bayes Class-Wise Performance Metrics", "33"),
        ("Table 4.11: Master Model Performance Comparison Across All Syllabus Metrics", "33")
    ]
    for title, pg in lot_entries:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.line_spacing = 1.15
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(11)
        dots_len = max(2, 85 - len(title) - len(pg))
        r_dots = p_t.add_run(" " + "." * dots_len + " ")
        r_dots.font.name = 'Times New Roman'
        r_dots.font.size = Pt(11)
        r_dots.font.color.rgb = RGBColor(150, 150, 150)
        r_p = p_t.add_run(pg)
        r_p.bold = True
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(11)

    # List of Figures
    add_styled_p("", space_before=18)
    add_h1("LIST OF FIGURES")
    lof_entries = [
        ("Figure 3.1: Continuous Grain Yield Distribution & Data-Derived Tertile Boundaries", "21"),
        ("Figure 3.2: Terminal Heat Penalty - Grain Yield vs Sowing Day (Julian Day)", "22"),
        ("Figure 3.3: District-Wise Grain Yield Distribution Across 10 Surveyed Districts", "23"),
        ("Figure 3.4: Pearson Correlation Matrix of Pre-Sowing Features & Target", "24"),
        ("Figure 4.1: Normalized Confusion Matrices Across All 5 Syllabus Classifiers", "26"),
        ("Figure 4.2: Top 15 Feature Importances from Decision Tree Model", "29"),
        ("Figure 4.3: Leakage-Free Pipeline 5-Fold Cross-Validation Curves for KNN and Decision Tree", "32"),
        ("Figure 4.4: Cross-Model Comparative Performance on Syllabus Unit III Metrics", "34"),
        ("Figure 4.5: One-vs-Rest (OvR) Macro-Averaged ROC Curves Across Classifiers", "35")
    ]
    for title, pg in lof_entries:
        p_t = doc.add_paragraph()
        p_t.paragraph_format.line_spacing = 1.15
        p_t.paragraph_format.space_after = Pt(2)
        r_t = p_t.add_run(title)
        r_t.font.name = 'Times New Roman'
        r_t.font.size = Pt(11)
        dots_len = max(2, 85 - len(title) - len(pg))
        r_dots = p_t.add_run(" " + "." * dots_len + " ")
        r_dots.font.name = 'Times New Roman'
        r_dots.font.size = Pt(11)
        r_dots.font.color.rgb = RGBColor(150, 150, 150)
        r_p = p_t.add_run(pg)
        r_p.bold = True
        r_p.font.name = 'Times New Roman'
        r_p.font.size = Pt(11)

    doc.add_page_break()

    # =========================================================
    # CHAPTER 1: INTRODUCTION
    # =========================================================
    add_h1("1. INTRODUCTION")
    
    add_h2("1.1 Context and Agronomic Significance of Wheat in Eastern India")
    add_body("Wheat (Triticum aestivum L.) represents one of the foundational pillars of national food security in India. It accounts for approximately 35% of total food grain production and supplies over 20% of total dietary energy consumed nationwide. While the high-input agricultural regimes of north-western India (such as Punjab and Haryana) have reached productivity plateaus, the eastern Indo-Gangetic Plains (IGP)—predominantly comprising the states of Bihar and eastern Uttar Pradesh—represent the primary geographic frontier for sustainable crop yield improvement.")
    add_body("The agricultural landscape of the eastern IGP is characterized by small and marginal landholdings, high population density, fragmented land parcels, and diverse socio-economic conditions among farming communities. Despite endowed groundwater resources and fertile alluvial soils, crop yields in eastern India have historically exhibited substantial variation. Farmers in this region face complex biophysical constraints, most notably terminal heat stress, delayed planting following late-harvested rice, variable fertilizer applications, and erratic monsoon precipitation patterns.")
    add_body("Understanding the patterns that govern wheat productivity in this agro-ecological zone requires moving beyond uniform blanket recommendations toward data-driven predictive insights. The availability of large-scale, field-level agronomic monitoring datasets offers an empirical foundation to analyze interactions between soil texture, management practices, weather conditions, and crop performance.")

    add_h2("1.2 The Agricultural Imperative for Pre-Sowing Predictive Analytics")
    add_body("In predictive analytics applied to agriculture, predictive tasks can be categorized into two distinct paradigms: retrospective post-harvest estimation and prospective pre-sowing categorization. Retrospective estimation utilizes post-harvest measurements, remote sensing vegetation indices throughout the reproductive cycle, and seasonal weather accumulations. While valuable for macro-economic grain accounting, retrospective estimation provides zero operational benefit to an individual grower who must make critical decisions months before harvest.")
    add_body("In contrast, pre-sowing predictive analytics operates under strict temporal information boundaries: all explanatory inputs must be reasonably known and measurable by the date of sowing. The pre-sowing window is a decisive period in the cropping calendar. It is during this narrow interval that farmers determine the planting date, select the wheat cultivar (short duration vs. long duration), choose the crop establishment method (conventional tillage vs. zero tillage), apply basal fertilizers (such as DAP, NPK, and Potash), and manage preceding crop residue.")
    add_body("Providing agricultural stakeholders with pre-sowing yield category classifications enables informed planning. For example, if a field is classified as facing high risk of a 'Low' yield under current timing, proactive adjustments can be explored—such as avoiding delayed planting, selecting early-maturing cultivars, or adjusting basal nutrient applications.")

    add_h2("1.3 Problem Formulation: Yield Categorization vs. Continuous Regression")
    add_body("Wheat grain yield in agronomic surveys is recorded as a continuous numerical variable (measured in tonnes per hectare, t/ha). A conventional modeling approach might formulate this challenge as continuous regression. However, from decision-analytic and statistical standpoints, categorizing continuous yields into discrete, risk-adjusted tertiles ('Low', 'Medium', 'High') offers key advantages:")
    add_bullet("Decision-Theoretic Practicality:", "Agricultural stakeholders often make discrete decisions rather than continuous adjustments. Decisions such as crop insurance enrollment, institutional credit underwriting, and input procurement are fundamentally categorical. A probabilistic classification indicating class membership provides practical guidance for risk assessment.")
    add_bullet("Robustness to Measurement Noise:", "Field-level agricultural survey data inevitably contains spatial and sampling variability from quadrant crop-cuts and self-reported estimations. Binning continuous yields into tertile categories dampens micro-level noise, allowing classification algorithms to focus on broader structural patterns.")
    add_bullet("Balanced Class Representation:", "Partitioning the yield distribution into exact tertiles (lower one-third, middle one-third, and upper one-third) ensures balanced target classes. This eliminates class imbalance artifacts and allows standard classification loss functions (such as cross-entropy and multinomial log-loss) to evaluate discrimination power fairly across the entire distribution.")

    add_h2("1.4 Scoping and Framing: Wheat Classification vs. Multi-Crop Recommendation")
    add_body("It is critical to establish the precise methodological framing of this project. Academic integrity and rigorous data science necessitate that predictive models reflect the empirical domain of the training dataset. The supplied dataset comprises exclusively wheat cultivation records collected across 5,689 fields over six consecutive crop years. No observations exist for alternative crops such as mustard, chickpea, lentil, or maize in the Rabi season.")
    add_body("Therefore, this project is framed strictly and accurately as:")
    add_styled_p("“Wheat Yield Category Classification Using Field, Soil, and Pre-sowing Information”", space_before=6, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True)
    add_body("The model developed herein does not recommend which crop to sow, nor does it classify between different crop types. Instead, it addresses an intra-crop classification task: given a field managed for wheat, how will the field perform (Low, Medium, or High) based on pre-sowing soil, management, and meteorological predictors? If general templates describe multi-crop recommendation, we explicitly clarify that the available real dataset supports this narrower, focused classification task.")

    add_h2("1.5 Alignment with INT234 Course Objectives and Syllabus Units")
    add_body("This project is engineered in strict compliance with the curriculum of INT234: Predictive Analytics. Specifically, it directly operationalizes the following course outcomes (COs) and curriculum units:")
    add_bullet("Unit I & Practical 1, 2 (Data Preparation & Preprocessing):", "Exploration of variable types, handling missing data with documented domain reasoning, exploring numeric and categorical distributions, feature encoding, and normalization/standardization.")
    add_bullet("Unit II & Practical 3 (Supervised Learning - Foundations):", "Evaluating model performance, correlations, and formulating supervised predictive objectives. Implementation of Logistic Regression.")
    add_bullet("Unit III & Practical 4, 5 (Supervised Learning - Classification):", "Comprehensive implementation and empirical comparison of the core classification algorithms specified in the syllabus: Lazy learning (K-Nearest Neighbors), Probabilistic learning (Naive Bayes), Divide and Conquer (Decision Trees), and Kernel methods (Support Vector Machines).")
    add_bullet("Unit III (Evaluation Metrics):", "Rigorous evaluation using Accuracy, Logarithmic Loss (Log-Loss), Area Under the ROC Curve (multiclass One-vs-Rest ROC AUC), Precision, Recall, F1-Score (macro and weighted), and full Confusion Matrix analysis.")
    add_body("In accordance with course guidelines, no black-box or ensemble algorithms outside the syllabus (such as XGBoost, LightGBM, or deep neural networks) have been introduced. The investigation remains rigorously focused on comparing the five core syllabus classification paradigms under identical, leakage-free conditions.")

    add_h2("1.6 Organization of the Project Report")
    add_body("The remainder of this report is structured as follows:")
    add_body("Chapter 2 details the source of the dataset, providing the official CIMMYT citation, multi-year survey sampling design across 10 districts in Bihar and eastern UP, variable descriptions, and open access licensing.")
    add_body("Chapter 3 presents the Exploratory Data Analysis (EDA) process, establishing the temporal boundary audit to prevent data leakage, detailing the missing value handling for BasalZn and GradeNPK, illustrating target and feature distributions, and defining the scikit-learn preprocessing pipeline.")
    add_body("Chapter 4 forms the analytical core, presenting an exhaustive, modular examination of all five syllabus classification models. Each algorithm is evaluated through its mathematical foundation, hyperparameter cross-validation on training data, empirical results, confusion matrices, and ROC curves, concluding with a comparative synthesis.")
    add_body("Chapter 5 summarizes the conclusions, highlighting key observational associations and explaining why classification accuracy must be interpreted alongside class-wise recall.")
    add_body("Chapter 6 outlines the future scope, discussing future-year holdout validation and satellite remote sensing integration.")
    add_body("Chapter 7 lists verifiable academic, institutional, and textbook references.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 2: SOURCE OF DATASET
    # =========================================================
    add_h1("2. SOURCE OF DATASET")

    add_h2("2.1 Provenance and Official CIMMYT Repository Citation")
    add_body("The empirical dataset analyzed in this project originates from the International Maize and Wheat Improvement Center (CIMMYT), a non-profit international agricultural research organization belonging to the CGIAR consortium. The dataset is officially cited as follows:")
    add_styled_p("Ajay, Anurag (2026), “Field-level crop management, weather, and yield data of wheat from eastern India,” CIMMYT Research Data & Software Repository Network, Version 1. DOI: https://doi.org/10.71682/10549424. License: CC0 1.0.",
                 space_before=6, space_after=6, align=WD_ALIGN_PARAGRAPH.CENTER, bold=True, italic=True)
    add_body("The dataset was published as open-access research data on the CIMMYT Research Data & Software Repository Network and is accompanied by variable documentation. It represents an extensive field-level agronomic survey collected across smallholder farming systems in South Asia.")

    add_h2("2.2 Survey Methodology and Regional Sampling Design")
    add_body("The dataset encompasses a total of 5,689 unique wheat field records collected across two major states in eastern India: Bihar (4,464 fields, 78.47%) and eastern Uttar Pradesh (1,225 fields, 21.53%). The survey spanned ten target districts representing diverse agro-climatic sub-zones of the Middle and Lower Gangetic Plains:")
    add_bullet("Bihar Districts (8):", "Begusarai (380 fields), Bhojpur (632 fields), Buxar (709 fields), East Champaran (825 fields), Lakhisarai (347 fields), Madhepura (558 fields), Muzaffarpur (427 fields), and Rohtas (586 fields).")
    add_bullet("Uttar Pradesh Districts (2):", "Deoria (669 fields) and Kushinagar (556 fields).")
    add_body("Data collection was conducted using structured face-to-face farmer interviews coupled with in-field measurements by trained agronomic enumerators. Survey protocols captured the entire agronomic cycle of each field, including parcel history, tillage practices, cultivar selection, seed rate, sowing timing, basal and top-dressed fertilization, irrigation events, weed management, lodging incidence, harvesting dates, and final grain yield determined through standard crop-cutting quadrant sampling.")

    add_h2("2.3 Temporal Horizon (Crop Years 2016-17 to 2021-22)")
    add_body("The survey was conducted across six consecutive wheat cropping seasons (Rabi seasons), spanning 2016–17 through 2021–22. The multi-year temporal coverage is valuable for predictive modeling as it captures inter-annual weather variation, including monsoon patterns, pre-sowing soil moisture reserves, and winter temperature profiles.")
    add_body("The distribution of surveyed fields across the six crop years is as follows:")
    add_bullet("2016-17:", "777 fields (13.66%)")
    add_bullet("2017-18:", "1,313 fields (23.08%)")
    add_bullet("2018-19:", "1,403 fields (24.66%)")
    add_bullet("2019-20:", "1,318 fields (23.17%)")
    add_bullet("2020-21:", "725 fields (12.74%)")
    add_bullet("2021-22:", "153 fields (2.69%)")
    add_body("Total observations across all years equal 5,689 fields. The concentration of records between 2017 and 2020 provides a substantial empirical base of smallholder management.")

    add_h2("2.4 Data Attributes and Companion Documentation")
    add_body("The raw dataset is provided in two companion Excel workbooks: 'Wheat_Yld_Eastern_IND_Data.xlsx' containing the 5,689 observations across 49 variables, and 'Wheat_Yld_Eastern_IND_Variables.xlsx' providing the formal variable metadata, variable types, descriptions, and measurement units.")
    
    # Table 2.1
    add_styled_p("Table 2.1: Variable Definitions, Types, and Agronomic Significance", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_vars = doc.add_table(rows=1, cols=4)
    hdr_cells = table_vars.rows[0].cells
    hdr_cells[0].text = "Variable Name"
    hdr_cells[1].text = "Data Type"
    hdr_cells[2].text = "Measurement Unit"
    hdr_cells[3].text = "Agronomic Description & Role"
    
    var_rows = [
        ("Field_ID", "Nominal", "Dimensionless", "Unique identifier assigned to each monitored field parcel."),
        ("Year", "Ordinal", "Crop Year", "Cropping season (2016-17 to 2021-22) capturing macro-climate."),
        ("District", "Nominal", "Administrative", "District name capturing spatial geography and regional soil."),
        ("State", "Nominal", "State name", "Bihar or Uttar Pradesh (hierarchically subsumed by District)."),
        ("LandType", "Ordinal", "Topography", "Land position: Low (lowland), Medium (midland), Up (upland)."),
        ("SoilType", "Ordinal", "Texture", "Farmer-reported soil texture: Medium, Heavy, Light."),
        ("PreviousCrop", "Nominal", "Crop name", "Crop preceding wheat (Rice, Maize, Soybean, Fallow, Sugarcane)."),
        ("PrevCropResidue", "Discrete", "Percentage (%)", "Proportion of preceding crop residue retained in field."),
        ("CropEstablishment", "Binary", "Categorical", "Planting method: Zero Tillage (ZT) vs. Conventional Tillage (CT)."),
        ("Variety", "Nominal", "Cultivar code", "Specific wheat variety sown (e.g. HD2967, PBW373, HI1563)."),
        ("VarietyClass", "Ordinal", "Duration class", "Cultivar maturity duration category: Long duration vs. Short duration."),
        ("SeedRate", "Continuous", "kg per hectare", "Seeding density used per unit area (ranges from 80 to 225 kg/ha)."),
        ("SowingDay", "Continuous", "Julian Day", "Day of calendar year of sowing (ranges from day 295 to 355)."),
        ("BasalDAP", "Continuous", "kg per hectare", "Di-ammonium Phosphate fertilizer applied at sowing time."),
        ("BasalNPK", "Continuous", "kg per hectare", "Composite NPK complex fertilizer mixture applied at sowing."),
        ("GradeNPK", "Nominal", "Ratio string", "Grade formulation of composite NPK mixture (e.g. 12:32:16)."),
        ("BasalMOP", "Continuous", "kg per hectare", "Muriate of Potash (potassium chloride) applied at sowing."),
        ("BasalZn", "Continuous", "kg per hectare", "Zinc fertilizer applied at sowing (excluded due to 91% missingness)."),
        ("Prcp_Monsoon", "Continuous", "Millimeters (mm)", "Cumulative precipitation during preceding monsoon season."),
        ("Prcp_Oct", "Continuous", "Millimeters (mm)", "Cumulative precipitation during October prior to wheat sowing."),
        ("GrainYield", "Continuous", "Tonnes / hectare", "Final grain yield measured at harvest (Continuous Target Variable).")
    ]
    for row_data in var_rows:
        row = table_vars.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_vars)

    add_h2("2.5 Data Quality, Governance, and Open Access Licensing")
    add_body("The dataset is released under the Creative Commons Zero (CC0 1.0) Public Domain Dedication, permitting public academic reuse, modification, and reproduction. The data was collected under institutional ethical protocols and anonymized to protect individual privacy. In this project, raw dataset files are preserved in their original condition inside the 'Dataset/dataverse_files/' directory, with all cleaning and transformation executed via reproducible Python scripts.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 3: EDA PROCESS
    # =========================================================
    add_h1("3. EDA PROCESS (EXPLORATORY DATA ANALYSIS)")

    add_h2("3.1 Structural Overview and Initial Taxonomy of Variables")
    add_body("Exploratory Data Analysis (EDA) forms the foundational diagnostic phase of predictive analytics. It bridges raw data ingestion with machine learning modeling by uncovering distribution shapes, identifying data quality anomalies, auditing information timing, and guiding feature engineering. The raw dataset contains 5,689 rows and 49 columns, spanning diverse data structures including identifiers, administrative factors, continuous management quantities, calendar dates, climatic metrics, and geographic coordinates.")

    add_h2("3.2 Identification of Temporal Information Boundaries & Leakage Prevention Audit")
    add_body("A critical methodological requirement in predictive analytics is temporal data leakage prevention. Data leakage occurs when variables that would only be observed after the prediction event are inadvertently included as input features during model training. Models trained on leaked variables exhibit high offline test scores but fail when evaluated in prospective deployment.")
    add_body("In this wheat yield classification task, the temporal boundary is rigidly fixed at the moment of sowing. Any phenomenon that occurs subsequent to sowing belongs to the in-season growth and harvest phases and is excluded from the input feature set.")
    
    # Table 3.1
    add_styled_p("Table 3.1: Temporal Information Boundary Audit and Feature Inclusion Rationale", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_leakage = doc.add_table(rows=1, cols=4)
    hdr = table_leakage.rows[0].cells
    hdr[0].text = "Variable Name"
    hdr[1].text = "Temporal Stage"
    hdr[2].text = "Status"
    hdr[3].text = "Detailed Rationale & Exclusion Mechanism"

    leakage_rows = [
        ("GrainYield", "Post-Harvest", "EXCLUDED (Target)", "Continuous prediction target. Partitioned into tertiles; excluded from inputs."),
        ("Field_ID", "Administrative", "EXCLUDED", "Arbitrary record identifier without generalizable predictive value."),
        ("Latitude, Longitude", "Spatial centroid", "EXCLUDED", "Excluded to prevent spatial overfitting and GPS cluster memorization; spatial variance is captured by District."),
        ("HarvestDate, HarvestDay", "Post-Harvest", "EXCLUDED (Leakage)", "Recorded 110-140 days after sowing. Unavailable at sowing time."),
        ("CropLodging", "Late Grain-Filling", "EXCLUDED (Leakage)", "Proportion of canopy lodged due to wind/rain during maturity. Strictly post-sowing occurrence."),
        ("IrrigationNumber", "In-Season Management", "EXCLUDED (Leakage)", "Total number of irrigation events applied throughout the season; unknown at planting."),
        ("FirstIrrigationDay - FifthIrrigationDay", "In-Season Timing", "EXCLUDED (Leakage)", "Dates of subsequent irrigations applied across the vegetative cycle."),
        ("Split1Urea, Split2Urea, Split3Urea", "In-Season Management", "EXCLUDED (Leakage)", "Top-dressed nitrogen applications executed weeks after crop emergence."),
        ("HerbicideName, HerbicideType, WeedingNumber", "In-Season Management", "EXCLUDED (Leakage)", "Post-emergence chemical weed control and manual weeding events occurring after sowing."),
        ("Nov_Tmax, EDD_Nov through Apr_Tmax, EDD_Apr", "In-Season Meteorology", "EXCLUDED (Leakage)", "In-season maximum temperatures and extreme degree days across crop months."),
        ("Year, District", "Pre-Sowing Context", "INCLUDED", "Known prior to planting. Captures broad macro-environmental, administrative, and soil conditions."),
        ("LandType, SoilType", "Inherent Soil Attributes", "INCLUDED", "Physical parcel characteristics reported by farmer based on topography and texture."),
        ("PreviousCrop, PrevCropResidue", "Rotational History", "INCLUDED", "Historical management of preceding Kharif crop and surface residue."),
        ("CropEstablishment", "Sowing Operation", "INCLUDED", "Tillage decision (Zero Tillage vs Conventional Tillage) made at sowing."),
        ("Variety, VarietyClass", "Cultivar Choice", "INCLUDED", "Seed variety and maturity duration classification selected by farmer at planting."),
        ("SeedRate, SowingDay", "Sowing Parameters", "INCLUDED", "Seeding density (kg/ha) and calendar day of sowing (Julian day)."),
        ("BasalDAP, BasalNPK, BasalMOP", "Basal Fertilization", "INCLUDED", "Soil nutrients applied concurrently with seed placement during field sowing."),
        ("BasalZn", "Basal Fertilization", "EXCLUDED", "Excluded due to 91.14% missingness and absence of verification that blank implies zero."),
        ("Prcp_Monsoon, Prcp_Oct", "Pre-Sowing Meteorology", "INCLUDED", "Preceding rainfall accumulation capturing subsoil moisture prior to planting.")
    ]
    for row_data in leakage_rows:
        row = table_leakage.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_leakage)

    add_h2("3.3 Systematic Missing Value Analysis and Domain-Driven Handling")
    add_body("A systematic audit of missing values was conducted across all candidate variables before preprocessing. In tabular survey datasets, handling missing entries requires documented justification to prevent introducing arbitrary artifacts.")
    
    # Table 3.2
    add_styled_p("Table 3.2: Missing Value Audit Across Pre-Sowing Candidate Variables", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_na = doc.add_table(rows=1, cols=5)
    hdr = table_na.rows[0].cells
    hdr[0].text = "Variable Name"
    hdr[1].text = "Total Records"
    hdr[2].text = "Missing Count"
    hdr[3].text = "Missingness (%)"
    hdr[4].text = "Handling Strategy & Methodological Rationale"

    na_rows = [
        ("BasalZn", "5,689", "5,185", "91.14%", "Excluded from inputs. The dataset documentation does not confirm blank implies zero application."),
        ("GradeNPK", "5,689", "4,659", "81.89%", "Excluded from inputs. Missing entries correspond perfectly to BasalNPK == 0 (no NPK applied)."),
        ("Prcp_Monsoon", "5,689", "19", "0.33%", "Imputed using the general training-set median (SimpleImputer fit on training data only)."),
        ("Prcp_Oct", "5,689", "0", "0.00%", "Complete. Zero missing entries."),
        ("SoilType, LandType", "5,689", "0", "0.00%", "Complete. Casing standardized."),
        ("Variety, VarietyClass", "5,689", "0", "0.00%", "Complete. 25 distinct varieties."),
        ("SowingDay, SeedRate", "5,689", "0", "0.00%", "Complete continuous variables.")
    ]
    for row_data in na_rows:
        row = table_na.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_na)

    add_h3("Methodological Rationale for Excluding BasalZn")
    add_body("The variable BasalZn records the quantity of zinc micronutrient fertilizer applied as basal dressing during sowing. In the supplied dataset, 5,185 out of 5,689 records (91.14%) are empty (NaN). The 504 reported entries range from 1.5 kg/ha to 45 kg/ha. Importantly, the companion variable documentation ('Wheat_Yld_Eastern_IND_Variables.xlsx') does not state or verify that an unrecorded entry represents a zero application. In survey research, unrecorded entries may reflect unasked questions, farmer non-response, or unmonitored practices. Imputing zero across over 91% of records would impose an unverified assumption. Therefore, BasalZn is excluded from the model feature set to ensure all retained inputs rest on verified data.")

    add_h3("Structural Collinearity of GradeNPK with BasalNPK")
    add_body("The variable GradeNPK records the formulation grade of composite fertilizer. In the raw dataset, 4,659 fields have missing GradeNPK. Cross-tabulating GradeNPK against continuous BasalNPK reveals complete structural alignment:")
    
    # Table 3.3
    add_styled_p("Table 3.3: Cross-Tabulation of BasalNPK Application Rate vs GradeNPK Reporting", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_npk = doc.add_table(rows=1, cols=6)
    hdr = table_npk.rows[0].cells
    hdr[0].text = "GradeNPK Status"
    hdr[1].text = "Record Count"
    hdr[2].text = "BasalNPK Mean"
    hdr[3].text = "BasalNPK Min"
    hdr[4].text = "BasalNPK Max"
    hdr[5].text = "Interpretation"

    npk_rows = [
        ("Missing (NaN)", "4,659", "0.00 kg/ha", "0.0 kg/ha", "0.0 kg/ha", "Field received 0 kg/ha composite NPK mixture."),
        ("Reported (Non-null)", "1,030", "84.38 kg/ha", "5.0 kg/ha", "185.0 kg/ha", "Field received composite NPK mixture.")
    ]
    for row_data in npk_rows:
        row = table_npk.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_npk)

    add_body("As shown in Table 3.3, whenever GradeNPK is missing, BasalNPK is 0.0 kg/ha. Whenever GradeNPK is reported, BasalNPK is positive (ranging from 5 to 185 kg/ha). Because the quantitative fertilizer dosage is already represented by continuous BasalNPK, GradeNPK is omitted to avoid redundant high-cardinality nominal parameters.")

    add_h3("Consistent Training-Set Median Imputation for Prcp_Monsoon")
    add_body("The variable Prcp_Monsoon contains 19 missing values (0.33% of records). To prevent data snooping across partitions, these missing values are imputed using a general median imputer (SimpleImputer with strategy='median') fit strictly on the training partition. The identical median value computed from the training partition is then used to impute any missing entries in the testing set.")

    add_h2("3.4 Text Standardization and Categorical Normalization")
    add_body("Prior to one-hot encoding, string representations across categorical attributes were harmonized. The attribute SoilType contained inconsistent capitalization: 'Medium' (4,392), 'Heavy' (1,248), 'medium' (34), 'light' (8), and 'Light' (7). These were standardized to title case, collapsing categories into three strata: Medium (4,426), Heavy (1,248), and Light (15). Similarly, PreviousCrop contained 'Rice' (5,439), 'Maize' (99), 'Soybean' (90), 'Fallow' (32), 'maize' (18), 'soybean' (10), and 'Sugarcane' (1), which were mapped to title case: Rice (5,439), Maize (117), Soybean (100), Fallow (32), and Sugarcane (1).")

    add_h2("3.5 Target Distribution Analysis and Empirical Tertile Cutoff Protocol")
    add_body("The continuous target variable GrainYield spans from 1.390 to 6.955 t/ha, with a mean of 4.145 t/ha, median of 4.190 t/ha, and standard deviation of 1.037 t/ha. To establish balanced classification targets, the yield distribution is partitioned into tertiles.")
    add_body("To prevent data leakage, the lower cutoff (c1) and upper cutoff (c2) are computed strictly from the training partition (80% subset) after stratified splitting:")
    add_bullet("Lower Tertile Cutoff (c1):", "3.6000 t/ha (Yields <= 3.6000 t/ha are categorized as 'Low')")
    add_bullet("Upper Tertile Cutoff (c2):", "4.6700 t/ha (Yields > 4.6700 t/ha are categorized as 'High')")
    add_bullet("Intermediate Range:", "Yields strictly between 3.6000 and 4.6700 t/ha are categorized as 'Medium'")
    add_body("It is important to emphasize that these cutoff thresholds represent empirical data-derived categories based on the sample distribution of the training records, rather than official agronomic standards.")
    
    # Table 3.4
    add_styled_p("Table 3.4: Class Boundaries, Cutoff Values, and Stratified Partition Distribution", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_classes = doc.add_table(rows=1, cols=6)
    hdr = table_classes.rows[0].cells
    hdr[0].text = "Yield Class"
    hdr[1].text = "Yield Range (t/ha)"
    hdr[2].text = "Train Count"
    hdr[3].text = "Train Proportion"
    hdr[4].text = "Test Count"
    hdr[5].text = "Test Proportion"

    class_rows = [
        ("Low", "<= 3.6000 t/ha", "1,527", "33.55%", "382", "33.57%"),
        ("Medium", "3.6000 - 4.6700 t/ha", "1,518", "33.36%", "380", "33.39%"),
        ("High", "> 4.6700 t/ha", "1,506", "33.09%", "376", "33.04%"),
        ("Total", "1.390 - 6.955 t/ha", "4,551", "100.00%", "1,138", "100.00%")
    ]
    for row_data in class_rows:
        row = table_classes.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_classes)

    add_fig("01_target_distribution.png", "Figure 3.1: Continuous Grain Yield Distribution & Data-Derived Tertile Boundaries across Partitions")

    add_h2("3.6 Stratified 80/20 Train-Test Splitting Design")
    add_body("To ensure reproducible, unbiased evaluation, the 5,689 records were partitioned into an 80% training set (4,551 observations) and a 20% holdout test set (1,138 observations) using stratified sampling on initial tertile boundaries with a fixed random seed (random_state=42). Stratification guarantees that the target class proportions in the test set mirror the training distribution identically (approximately 33.33% each).")

    add_h2("3.7 Univariate, Bivariate, and Multivariate Relationships")
    add_body("Exploratory analysis reveals observed statistical associations between pre-sowing predictors and final yield:")
    add_h3("Observed Sowing Day Association and Terminal Heat")
    add_body("In the eastern Gangetic Plains, wheat is planted between late October and late December. The Julian sowing day (SowingDay) exhibits a negative linear correlation with grain yield. Fields sown earlier (prior to Julian day 320) tend to record higher yields, whereas fields planted later (after Julian day 335) are observed to record lower yields, consistent with agronomic literature describing heat exposure during late grain filling.")
    
    add_fig("02_yield_vs_sowing_day.png", "Figure 3.2: Terminal Heat Penalty - Grain Yield vs Sowing Day (Julian Day)")

    add_h3("District-Wise Yield Differences")
    add_body("Analysis of yield distributions across districts shows regional variation. South Bihar districts such as Buxar, Rohtas, and Bhojpur exhibit higher median yields in the sample, whereas northern districts such as East Champaran and Deoria record lower median yields.")

    add_fig("03_district_yield_distribution.png", "Figure 3.3: District-Wise Grain Yield Distribution Across 10 Surveyed Districts")

    add_h2("3.8 Pearson and Spearman Correlation Analysis")
    add_body("Correlations among pre-sowing numeric features and continuous yield were assessed using Pearson correlation coefficients (excluding BasalZn):")
    
    # Table 3.5
    add_styled_p("Table 3.5: Correlation Coefficients of Numeric Features with Grain Yield", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    table_corr = doc.add_table(rows=1, cols=4)
    hdr = table_corr.rows[0].cells
    hdr[0].text = "Feature Name"
    hdr[1].text = "Pearson r"
    hdr[2].text = "p-value"
    hdr[3].text = "Observed Statistical Relationship"

    corr_rows = [
        ("SowingDay", "-0.248", "< 0.001", "Negative linear association; later planting dates correlate with lower yields."),
        ("Prcp_Monsoon", "+0.182", "< 0.001", "Positive association; higher monsoon rainfall correlates with higher yields."),
        ("BasalDAP", "+0.145", "< 0.001", "Positive correlation with basal phosphate application."),
        ("BasalNPK", "+0.098", "< 0.001", "Moderate positive correlation with basal composite NPK application."),
        ("SeedRate", "+0.076", "< 0.001", "Weak positive correlation across seeding rates."),
        ("BasalMOP", "+0.062", "< 0.001", "Weak positive correlation with basal potash application."),
        ("PrevCropResidue", "-0.041", "0.002", "Weak negative correlation with retained surface residue."),
        ("Prcp_Oct", "+0.012", "0.365", "Insignificant direct linear correlation.")
    ]
    for row_data in corr_rows:
        row = table_corr.add_row()
        for idx, val in enumerate(row_data):
            row.cells[idx].text = val
    style_table(table_corr)

    add_fig("04_correlation_heatmap.png", "Figure 3.4: Pearson Correlation Matrix of Pre-Sowing Numeric Features & Target (BasalZn Excluded)")

    add_h2("3.9 Outlier Diagnosis and Handling")
    add_body("Inspection of continuous variables confirmed that extreme values in features (such as SeedRate up to 225 kg/ha or SowingDay up to day 355) represent documented farmer practices rather than data entry errors. Because distance-based and gradient algorithms are sensitive to feature scales, standardization was incorporated into the pipeline.")

    add_h2("3.10 Design of Scikit-Learn Preprocessing Pipelines")
    add_body("All feature transformations were encapsulated inside a scikit-learn ColumnTransformer. The numeric pipeline applies general median imputation followed by StandardScaler. The categorical pipeline applies most-frequent imputation followed by OneHotEncoder, expanding the 16 retained raw predictors (8 numeric + 8 categorical) into 62 transformed features. Crucially, to eliminate data leakage during cross-validation, preprocessing was embedded alongside each classifier inside a unified scikit-learn Pipeline. When executing cross-validation on raw training features (X_train_raw), imputation statistics, scaling parameters, and one-hot encodings were fitted strictly within each 4/5 training fold and applied to the 1/5 validation fold, ensuring zero leakage of distribution properties across validation folds.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 4: ANALYSIS ON DATASET
    # =========================================================
    add_h1("4. ANALYSIS ON DATASET")
    add_body("In accordance with the INT234 syllabus, this chapter evaluates the five required classification algorithms: Logistic Regression, Decision Tree Classifier, Support Vector Machine, K-Nearest Neighbors, and Gaussian Naive Bayes. Model and hyperparameter choices were validated strictly using leakage-free 5-fold cross-validation on raw training data with preprocessing embedded inside each fold's pipeline. Final models were evaluated on the holdout test set only after all architectural and hyperparameter decisions were fixed.")

    add_fig("05_confusion_matrices.png", "Figure 4.1: Normalized Confusion Matrices Across All 5 Syllabus Classification Models on Holdout Test Set")

    # ---------------------------------------------------------
    # 4.1 Logistic Regression
    # ---------------------------------------------------------
    add_h2("4.1 Logistic Regression (Multinomial)")
    add_h3("i. Introduction")
    add_body("Logistic Regression is a foundational linear classification model covered in Unit II and Unit III of the syllabus. For multi-class tasks, it models probabilities via the multinomial softmax formulation, serving as a baseline to evaluate linear class separability.")

    add_h3("ii. General Description")
    add_body("Multinomial Logistic Regression models class probabilities using linear combinations of input features mapped through the softmax function. Model parameters are estimated by minimizing the L2-regularized cross-entropy loss via the L-BFGS optimizer.")

    add_h3("iii. Specific Requirements, Functions, and Formulas")
    add_body("For a K-class problem (where K=3: Low, Medium, High) and feature vector x, class probability is given by:")
    add_styled_p("P(Y = k | x) = exp(w_k^T x + b_k) / sum_{j=1}^K exp(w_j^T x + b_j)", space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
    add_body("The objective function minimizes the penalized cross-entropy:")
    add_styled_p("J(W) = - (1/N) * sum_{i=1}^N sum_{k=1}^K y_{i,k} * log(P(Y = k | x_i)) + (1 / 2C) * sum_{k=1}^K ||w_k||_2^2", space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)

    add_h3("iv. Analysis Results")
    add_body("Trained with C=1.0 and solver='lbfgs', Logistic Regression achieved 63.94% training accuracy and 62.04% test accuracy. Macro precision was 0.6108, macro recall was 0.6207, macro F1-score was 0.6106, log loss was 0.8130, and ROC AUC was 0.8004.")
    
    # Table 4.1 & 4.2
    add_styled_p("Table 4.1: Logistic Regression Confusion Matrix", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cm_lr = doc.add_table(rows=1, cols=4)
    for idx, txt in enumerate(["Actual Class", "Pred: Low", "Pred: Medium", "Pred: High"]):
        t_cm_lr.rows[0].cells[idx].text = txt
    for r_data in [("Actual Low (382)", "269 (70.4%)", "78 (20.4%)", "35 (9.2%)"),
                  ("Actual Medium (380)", "104 (27.4%)", "151 (39.7%)", "125 (32.9%)"),
                  ("Actual High (376)", "35 (9.3%)", "55 (14.6%)", "286 (76.1%)")]:
        row = t_cm_lr.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_cm_lr)

    add_styled_p("Table 4.2: Logistic Regression Class-Wise Performance Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_m_lr = doc.add_table(rows=1, cols=5)
    for idx, txt in enumerate(["Class", "Precision", "Recall", "F1-Score", "Support"]):
        t_m_lr.rows[0].cells[idx].text = txt
    for r_data in [("Low", "0.6593", "0.7042", "0.6810", "382"),
                  ("Medium", "0.5317", "0.3974", "0.4548", "380"),
                  ("High", "0.6413", "0.7606", "0.6959", "376")]:
        row = t_m_lr.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_m_lr)

    add_h3("v. Visualization")
    add_body("Logistic Regression shows reasonable identification of extreme classes (70.4% recall on Low, 76.1% recall on High), while exhibiting lower recall (39.7%) on the intermediate Medium class due to linear decision boundaries.")

    # ---------------------------------------------------------
    # 4.2 Decision Tree Classifier
    # ---------------------------------------------------------
    add_h2("4.2 Decision Tree Classifier")
    add_h3("i. Introduction")
    add_body("Decision Trees represent non-parametric divide-and-conquer classifiers covered in Unit III. They construct recursive orthogonal partitions that are easy to inspect.")

    add_h3("ii. General Description")
    add_body("The algorithm evaluates candidate binary splits across features, choosing splits that maximize purity. Hyperparameters were tuned strictly using 5-fold pipeline cross-validation on raw training data: while tree depth 9 reached the higher cross-validation mean of 65.02% (+/- 0.90%) compared to depth 7 at 63.94% (+/- 1.82%), representing a 1.08-percentage-point difference, depth 7 was retained as a deliberate simplicity trade-off. Choosing depth 7 yields a shallower, simpler tree structure with fewer decision rules, avoiding excessive leaf fragmentation and over-branching on noisy field survey records, thereby preserving transparent, actionable agronomic interpretability for agricultural extension workers.")

    add_h3("iii. Specific Requirements, Functions, and Formulas")
    add_body("Purity was evaluated using Gini Impurity:")
    add_styled_p("Gini(m) = 1 - sum_{k=1}^K (p_{m,k})^2", space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
    add_body("The selected split maximizes impurity reduction Delta I(m, theta).")

    add_h3("iv. Analysis Results")
    add_body("Configured with max_depth=7 and min_samples_leaf=5, the Decision Tree achieved 69.17% training accuracy and 64.32% test accuracy. Macro precision was 0.6500, macro recall was 0.6436, macro F1-score was 0.6447, and ROC AUC was 0.8128.")
    
    # Table 4.3 & 4.4
    add_styled_p("Table 4.3: Decision Tree Confusion Matrix", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cm_dt = doc.add_table(rows=1, cols=4)
    for idx, txt in enumerate(["Actual Class", "Pred: Low", "Pred: Medium", "Pred: High"]):
        t_cm_dt.rows[0].cells[idx].text = txt
    for r_data in [("Actual Low (382)", "244 (63.9%)", "108 (28.3%)", "30 (7.9%)"),
                  ("Actual Medium (380)", "64 (16.8%)", "212 (55.8%)", "104 (27.4%)"),
                  ("Actual High (376)", "17 (4.5%)", "83 (22.1%)", "276 (73.4%)")]:
        row = t_cm_dt.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_cm_dt)

    doc.add_page_break()
    add_styled_p("Table 4.4: Decision Tree Class-Wise Performance Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_m_dt = doc.add_table(rows=1, cols=5)
    for idx, txt in enumerate(["Class", "Precision", "Recall", "F1-Score", "Support"]):
        t_m_dt.rows[0].cells[idx].text = txt
    for r_data in [("Low", "0.7508", "0.6387", "0.6902", "382"),
                  ("Medium", "0.5261", "0.5579", "0.5415", "380"),
                  ("High", "0.6732", "0.7340", "0.7023", "376")]:
        row = t_m_dt.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_m_dt)

    add_h3("v. Visualization")
    add_body("Feature importances (Figure 4.2) identify SowingDay, District indicators, Prcp_Monsoon, and Variety as primary splitting features.")
    add_fig("08_feature_importance_tree.png", "Figure 4.2: Top 15 Feature Importances from Decision Tree Classifier (max_depth=7)")

    # ---------------------------------------------------------
    # 4.3 Support Vector Machine (SVM)
    # ---------------------------------------------------------
    add_h2("4.3 Support Vector Machine (SVM)")
    add_h3("i. Introduction")
    add_body("Support Vector Machines are maximum-margin classifiers covered in Unit III. By mapping features into reproducing kernel Hilbert spaces, SVMs construct optimal hyperplanes separating class clusters.")

    add_h3("ii. General Description")
    add_body("Using an RBF kernel, SVM measures non-linear similarity across standardized continuous inputs and binary indicators. On 5-fold leakage-free pipeline cross-validation across C in [0.5, 4.0], C=3.0 scored 69.96% (+/- 1.10%) while C=2.0 scored 69.83% (+/- 1.06%). The difference of 0.13% is very small. C=2.0 was selected over C=3.0 because a slightly lower penalty parameter enforces a wider soft margin, providing stronger regularization against smallholder field survey measurement errors and boundary noise, mitigating overfitting risks while matching the performance.")

    add_h3("iii. Specific Requirements, Functions, and Formulas")
    add_body("The dual optimization maximizes W(alpha) subject to box constraints 0 <= alpha_i <= C. The RBF kernel is defined as:")
    add_styled_p("K(x_i, x_j) = exp(- gamma * ||x_i - x_j||^2)", space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)

    add_h3("iv. Analysis Results")
    add_body("SVM achieved 77.41% training accuracy and 69.42% test accuracy. It recorded macro precision of 0.6886, macro recall of 0.6943, macro F1-score of 0.6893, log loss of 0.7216, and ROC AUC of 0.8458, performing best overall among the five tested models.")
    
    # Table 4.5 & 4.6
    add_styled_p("Table 4.5: Support Vector Machine Confusion Matrix", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cm_svm = doc.add_table(rows=1, cols=4)
    for idx, txt in enumerate(["Actual Class", "Pred: Low", "Pred: Medium", "Pred: High"]):
        t_cm_svm.rows[0].cells[idx].text = txt
    for r_data in [("Actual Low (382)", "304 (79.6%)", "56 (14.7%)", "22 (5.8%)"),
                  ("Actual Medium (380)", "89 (23.4%)", "198 (52.1%)", "93 (24.5%)"),
                  ("Actual High (376)", "22 (5.9%)", "66 (17.6%)", "288 (76.6%)")]:
        row = t_cm_svm.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_cm_svm)

    add_styled_p("Table 4.6: Support Vector Machine Class-Wise Performance Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_m_svm = doc.add_table(rows=1, cols=5)
    for idx, txt in enumerate(["Class", "Precision", "Recall", "F1-Score", "Support"]):
        t_m_svm.rows[0].cells[idx].text = txt
    for r_data in [("Low", "0.7325", "0.7958", "0.7629", "382"),
                  ("Medium", "0.6188", "0.5211", "0.5657", "380"),
                  ("High", "0.7146", "0.7660", "0.7394", "376")]:
        row = t_m_svm.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_m_svm)

    add_h3("v. Visualization")
    add_body("The SVM confusion matrix shows 79.6% recall on Low and 76.6% recall on High, with extreme cross-class errors held to 5.8% and 5.9%.")

    # ---------------------------------------------------------
    # 4.4 K-Nearest Neighbors (KNN)
    # ---------------------------------------------------------
    add_h2("4.4 K-Nearest Neighbors (KNN)")
    add_h3("i. Introduction")
    add_body("K-Nearest Neighbors is an instance-based lazy learning classifier covered in Unit III. It classifies query points using local neighborhoods without constructing explicit parametric models.")

    add_h3("ii. General Description")
    add_body("KNN measures distances across standardized predictors. Using 5-fold leakage-free pipeline cross-validation on raw training data across k in [5, 21], k=13 with distance-weighted Manhattan distance emerged as the empirical CV-winning setting, achieving 69.22% (+/- 1.11%) mean accuracy and outperforming both lower k (k=5 at 68.64%) and higher k (k=21 at 68.78%).")

    add_h3("iii. Specific Requirements, Functions, and Formulas")
    add_body("The Manhattan distance metric is defined as:")
    add_styled_p("d_1(x, z) = sum_{j=1}^p |x_j - z_j|", space_before=4, space_after=4, align=WD_ALIGN_PARAGRAPH.CENTER, italic=True)
    add_body("Neighbor votes are weighted inversely by distance: w_i = 1 / (d(x, x_i) + epsilon).")

    add_h3("iv. Analysis Results")
    add_body("On the test set, KNN achieved 68.63% accuracy, validating the reference baseline (~69%). Macro precision was 0.6813, macro recall was 0.6862, macro F1-score was 0.6823, log loss was 2.1825, and ROC AUC was 0.8381.")
    
    # Table 4.7 & 4.8
    add_styled_p("Table 4.7: K-Nearest Neighbors Confusion Matrix", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cm_knn = doc.add_table(rows=1, cols=4)
    for idx, txt in enumerate(["Actual Class", "Pred: Low", "Pred: Medium", "Pred: High"]):
        t_cm_knn.rows[0].cells[idx].text = txt
    for r_data in [("Actual Low (382)", "306 (80.1%)", "52 (13.6%)", "24 (6.3%)"),
                  ("Actual Medium (380)", "83 (21.8%)", "202 (53.2%)", "95 (25.0%)"),
                  ("Actual High (376)", "26 (6.9%)", "77 (20.5%)", "273 (72.6%)")]:
        row = t_cm_knn.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_cm_knn)

    add_styled_p("Table 4.8: K-Nearest Neighbors Class-Wise Performance Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_m_knn = doc.add_table(rows=1, cols=5)
    for idx, txt in enumerate(["Class", "Precision", "Recall", "F1-Score", "Support"]):
        t_m_knn.rows[0].cells[idx].text = txt
    for r_data in [("Low", "0.7373", "0.8010", "0.7679", "382"),
                  ("Medium", "0.6103", "0.5316", "0.5682", "380"),
                  ("High", "0.6964", "0.7261", "0.7109", "376")]:
        row = t_m_knn.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_m_knn)

    add_h3("v. Visualization")
    add_body("Leakage-free pipeline cross-validation sweeps on raw training data (Figure 4.3) confirm that k=13 is the empirical CV winner for KNN (69.22%), while for Decision Tree, depth 9 reached the highest CV mean (65.02%) and depth 7 (63.94%, a 1.08-point difference) was retained as a deliberate simplicity trade-off to prioritize a shallower, more interpretable tree over minor empirical gains.")
    add_fig("09_hyperparameter_tuning.png", "Figure 4.3: Leakage-Free Pipeline 5-Fold Cross-Validation Curves for KNN and Decision Tree on Raw Training Data")

    # ---------------------------------------------------------
    # 4.5 Gaussian Naive Bayes
    # ---------------------------------------------------------
    add_h2("4.5 Gaussian Naive Bayes")
    add_h3("i. Introduction")
    add_body("Naive Bayes is a probabilistic classifier based on Bayes' Theorem with conditional feature independence assumptions, covered in Unit III.")

    add_h3("ii. General Description")
    add_body("Gaussian Naive Bayes models feature likelihoods using normal distributions. However, because agricultural predictors (district, rainfall, soil) co-vary, the independence assumption is violated.")

    add_h3("iii. Specific Requirements, Functions, and Formulas")
    add_body("Posterior probabilities are computed as P(Y = k | x) proportional to P(Y = k) * prod P(x_j | Y = k), where continuous likelihoods follow the univariate Gaussian density.")

    add_h3("iv. Analysis Results")
    add_body("Gaussian Naive Bayes recorded 50.70% test accuracy. It exhibited an acute failure mode on the intermediate Medium class, correctly predicting only 5 out of 380 fields (1.32% recall, F1 of 0.0255), as correlated features pushed posterior probabilities into extreme classes.")
    
    # Table 4.9 & 4.10
    add_styled_p("Table 4.9: Gaussian Naive Bayes Confusion Matrix", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_cm_nb = doc.add_table(rows=1, cols=4)
    for idx, txt in enumerate(["Actual Class", "Pred: Low", "Pred: Medium", "Pred: High"]):
        t_cm_nb.rows[0].cells[idx].text = txt
    for r_data in [("Actual Low (382)", "221 (57.9%)", "4 (1.0%)", "157 (41.1%)"),
                  ("Actual Medium (380)", "116 (30.5%)", "5 (1.3%)", "259 (68.2%)"),
                  ("Actual High (376)", "22 (5.9%)", "3 (0.8%)", "351 (93.4%)")]:
        row = t_cm_nb.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_cm_nb)

    add_styled_p("Table 4.10: Gaussian Naive Bayes Class-Wise Performance Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_m_nb = doc.add_table(rows=1, cols=5)
    for idx, txt in enumerate(["Class", "Precision", "Recall", "F1-Score", "Support"]):
        t_m_nb.rows[0].cells[idx].text = txt
    for r_data in [("Low", "0.6156", "0.5785", "0.5965", "382"),
                  ("Medium", "0.4167", "0.0132", "0.0255", "380"),
                  ("High", "0.4576", "0.9335", "0.6142", "376")]:
        row = t_m_nb.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_m_nb)

    add_h3("v. Visualization")
    add_body("The confusion matrix in Figure 4.1 illustrates Naive Bayes' bimodal predictions, routing nearly all cases into either Low or High.")

    # ---------------------------------------------------------
    # 4.6 Comparative Synthesis
    # ---------------------------------------------------------
    add_h2("4.6 Comprehensive Comparative Performance Synthesis")
    add_body("Table 4.11 summarizes the final test set performance across all five algorithms using syllabus Unit III metrics:")
    
    # Table 4.11
    add_styled_p("Table 4.11: Master Model Performance Comparison Across All Syllabus Metrics", bold=True, font_size=11, align=WD_ALIGN_PARAGRAPH.CENTER)
    t_master = doc.add_table(rows=1, cols=8)
    for idx, txt in enumerate(["Model", "Train Acc (%)", "Test Acc (%)", "Precision", "Recall", "F1-Score", "Log Loss", "ROC AUC"]):
        t_master.rows[0].cells[idx].text = txt
    for r_data in [
        ("Support Vector Machine", "77.41", "69.42", "0.6886", "0.6943", "0.6893", "0.7216", "0.8458"),
        ("K-Nearest Neighbors", "98.44", "68.63", "0.6813", "0.6862", "0.6823", "2.1825", "0.8381"),
        ("Decision Tree", "69.17", "64.32", "0.6500", "0.6436", "0.6447", "1.1241", "0.8128"),
        ("Logistic Regression", "63.94", "62.04", "0.6108", "0.6207", "0.6106", "0.8130", "0.8004"),
        ("Gaussian Naive Bayes", "51.15", "50.70", "0.4966", "0.5084", "0.4121", "12.9987", "0.7091")
    ]:
        row = t_master.add_row()
        for idx, val in enumerate(r_data): row.cells[idx].text = val
    style_table(t_master)

    add_fig("06_model_comparison_metrics.png", "Figure 4.4: Cross-Model Comparative Performance on Syllabus Unit III Metrics")
    add_fig("07_roc_curves.png", "Figure 4.5: One-vs-Rest (OvR) Macro-Averaged ROC Curves Across All Syllabus Classifiers")

    add_h3("Synthesis of Model Performance")
    add_body("Comparison reveals clear empirical trends:")
    add_bullet("Top Classification Performance:", "SVM (69.42% accuracy, 0.7216 log loss, 0.8458 ROC AUC) and KNN (68.63% accuracy, 0.8381 ROC AUC) perform best, handling multi-dimensional management interactions effectively.")
    add_bullet("Validation of Reference Baseline:", "KNN closely corroborates the ~69% reference baseline reported in preliminary analyses.")
    add_bullet("Boundary Overlap:", "Errors occur primarily between adjacent classes (Low vs Medium, Medium vs High), reflecting continuous yield variation across empirical tertile thresholds.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 5: CONCLUSION
    # =========================================================
    add_h1("5. CONCLUSION")

    add_h2("5.1 Synthesis of Empirical Findings")
    add_body("This project evaluated five supervised classification algorithms from the INT234 syllabus to predict wheat yield categories (Low, Medium, High) using exclusively pre-sowing information. By excluding post-sowing variables, removing BasalZn due to unverified missingness, performing parameter cross-validation strictly on training data, and evaluating on an isolated test set, the analysis maintains strict methodological validity.")
    add_body("SVM (69.42% test accuracy, 0.8458 ROC AUC) and KNN (68.63% test accuracy, 0.8381 ROC AUC) demonstrated the strongest performance, showing that pre-sowing indicators contain meaningful predictive signal.")

    add_h2("5.2 Observational Insights for Field Management and Sowing Scheduling")
    add_body("The empirical analysis highlights notable observational patterns in the dataset:")
    add_bullet("Sowing Date Association:", "Julian sowing day exhibits a clear negative association with yield in this observational dataset, consistent with regional agronomic studies regarding terminal heat avoidance.")
    add_bullet("Pre-Sowing Rainfall:", "Cumulative monsoon precipitation prior to sowing shows a positive correlation with final yield categories, reflecting the importance of residual subsoil moisture.")
    add_bullet("Observational Nature:", "These findings represent statistical associations from survey data and should not be interpreted as proven causal effects without controlled field trials.")

    add_h2("5.3 Critical Methodological Reflection: Why Accuracy Alone is Insufficient")
    add_body("Aggregate accuracy does not capture the full performance picture. Class-wise recall and confusion matrices are essential to identify failure modes—such as Naive Bayes' collapse on the Medium class despite 50.70% overall accuracy, and the differing risk profiles of misclassifying extreme vs. adjacent classes.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 6: FUTURE SCOPE
    # =========================================================
    add_h1("6. FUTURE SCOPE")

    add_h2("6.1 Temporal Validation via Future-Year Holdouts")
    add_body("While the stratified random split provides a standard evaluation, real-world deployment requires predicting entirely unseen future crop seasons. A logical next validation step is conducting Leave-One-Year-Out (LOYO) cross-validation to assess generalization across shifting seasonal weather regimes.")

    add_h2("6.2 Integration of High-Resolution Satellite Remote Sensing")
    add_body("Satellite reflectance data prior to sowing (such as Sentinel-2 soil moisture and tillage indices) could offer objective radiometric measurements to complement farmer-reported variables.")

    add_h2("6.3 Coupling with Short-Range Numerical Weather Prediction")
    add_body("Coupling pre-sowing predictors with short-term (14-day) numerical weather forecasts around the sowing and emergence window may provide additional early-season signal.")

    add_h2("6.4 Deployment of Lightweight Decision Support Tools for Smallholders")
    add_body("Deploying trained models in mobile-compatible formats could assist agricultural extension personnel in providing localized risk assessments prior to planting.")

    doc.add_page_break()

    # =========================================================
    # CHAPTER 7: REFERENCES
    # =========================================================
    add_h1("7. REFERENCES")
    
    references_list = [
        ("Ajay", "Ajay, Anurag (2026). Field-level crop management, weather, and yield data of wheat from eastern India. CIMMYT Research Data & Software Repository Network, Version 1. https://doi.org/10.71682/10549424. License: CC0 1.0."),
        ("Erenstein et al.", "Erenstein, O., Thorpe, W., Singh, J., & Varma, A. (2007). Crop-livestock interactions and livelihoods in the Indo-Gangetic Plains of India: A regional synthesis. CIMMYT / ILRI Research Report, New Delhi, India."),
        ("Geron", "Geron, A. (2022). Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow: Concepts, Tools, and Techniques to Build Intelligent Systems (3rd ed.). O'Reilly Media."),
        ("Hastie et al.", "Hastie, T., Tibshirani, R., & Friedman, J. (2009). The Elements of Statistical Learning: Data Mining, Inference, and Prediction (2nd ed.). Springer Science & Business Media."),
        ("Jat et al.", "Jat, M. L., Chakraborty, D., Ladha, J. K., et al. (2020). Conservation agriculture for sustainable intensification in South Asia. Nature Food, 1(6), 336-343."),
        ("Muller & Guido", "Muller, A. C., & Guido, S. (2016). Introduction to Machine Learning with Python: A Guide for Data Scientists. O'Reilly Media."),
        ("Pedregosa et al.", "Pedregosa, F., Varoquaux, G., Gramfort, A., Michel, V., Thirion, B., et al. (2011). Scikit-learn: Machine learning in Python. Journal of Machine Learning Research, 12, 2825-2830."),
        ("Reynolds et al.", "Reynolds, M. P., Ortiz-Monasterio, J. I., & McNab, A. (Eds.). (2001). Application of Physiology in Wheat Breeding. CIMMYT, Mexico, D.F."),
        ("Weidman", "Weidman, S. (2019). Deep Learning from Scratch: Building with Python from First Principles. O'Reilly Media.")
    ]

    for author, ref in references_list:
        p_ref = doc.add_paragraph()
        p_ref.paragraph_format.line_spacing = 1.15
        p_ref.paragraph_format.space_after = Pt(8)
        p_ref.paragraph_format.left_indent = Inches(0.5)
        p_ref.paragraph_format.first_line_indent = Inches(-0.5)
        r_ref = p_ref.add_run(ref)
        r_ref.font.name = 'Times New Roman'
        r_ref.font.size = Pt(11)

    os.makedirs(os.path.dirname(output_docx_path), exist_ok=True)
    doc.save(output_docx_path)
    print(f"Project report successfully written to: {output_docx_path}")

if __name__ == '__main__':
    script_dir = os.path.dirname(__file__)
    docx_out = os.path.join(script_dir, '..', 'Report and Presentation', 'report', 'Wheat_Yield_Classification_Project_Report.docx')
    res_json = os.path.join(script_dir, '..', 'figures', 'model_evaluation_results.json')
    fig_dir = os.path.join(script_dir, '..', 'figures')
    create_full_report(docx_out, res_json, fig_dir)
