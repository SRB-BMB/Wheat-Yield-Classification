"""
Presentation Generation Script for Data Analytics CA2 Project.
Generates Wheat_Yield_Classification_Presentation.pptx in
'Report and Presentation/presentation/' using python-pptx.
"""

import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_pptx_path, figures_dir):
    prs = Presentation()
    # 16:9 widescreen
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    PRIMARY = RGBColor(26, 82, 118)      # Deep Navy Blue
    SECONDARY = RGBColor(44, 62, 80)     # Dark Slate
    ACCENT = RGBColor(230, 126, 34)      # Warm Orange
    BG_LIGHT = RGBColor(248, 249, 250)   # Light gray
    TEXT_DARK = RGBColor(40, 40, 40)
    WHITE = RGBColor(255, 255, 255)

    blank_layout = prs.slide_layouts[6]

    def add_header(slide, title_text, category_text="INT234 PREDICTIVE ANALYTICS | CA2 PROJECT"):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_top = Inches(0)
        tf.margin_bottom = Inches(0)
        tf.margin_left = Inches(0)
        tf.margin_right = Inches(0)
        
        p_cat = tf.paragraphs[0]
        p_cat.text = category_text.upper()
        p_cat.font.size = Pt(10)
        p_cat.font.bold = True
        p_cat.font.color.rgb = ACCENT
        
        p_title = tf.add_paragraph()
        p_title.text = title_text
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = PRIMARY
        p_title.space_before = Pt(4)

    # -------------------------------------------------------------
    # Slide 1: Title Slide
    # -------------------------------------------------------------
    slide1 = prs.slides.add_slide(blank_layout)
    bg1 = slide1.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
    bg1.fill.solid()
    bg1.fill.fore_color.rgb = PRIMARY
    bg1.line.fill.background()

    tb = slide1.shapes.add_textbox(Inches(1.0), Inches(1.2), Inches(11.333), Inches(5.2))
    tf = tb.text_frame
    tf.word_wrap = True

    p_sub = tf.paragraphs[0]
    p_sub.text = "INT234: PREDICTIVE ANALYTICS — CONTINUOUS ASSESSMENT 2 (CA2)"
    p_sub.font.size = Pt(14)
    p_sub.font.bold = True
    p_sub.font.color.rgb = ACCENT

    p_main = tf.add_paragraph()
    p_main.text = "Wheat Yield Category Classification Using Field, Soil, and Pre-Sowing Information"
    p_main.font.size = Pt(28)
    p_main.font.bold = True
    p_main.font.color.rgb = WHITE
    p_main.space_before = Pt(14)
    p_main.space_after = Pt(20)

    p_info = tf.add_paragraph()
    p_info.text = (
        "Student: Parth Ghuge  |  Roll Number: omitted from public copy  |  Group: omitted from public copy  |  Registration No: omitted from public copy  |  Section: omitted from public copy\n"
        "Faculty Supervisor: Course Supervisor  |  Course: INT234 (Predictive Analytics)\n"
        "Programme: B.Tech Computer Science and Engineering  |  Project Semester 2026-27\n"
        "Discipline of CSE/IT, School of Computer Science and Engineering, Lovely Professional University, Punjab"
    )
    p_info.font.size = Pt(13)
    p_info.font.color.rgb = RGBColor(210, 225, 240)
    p_info.space_before = Pt(14)

    # -------------------------------------------------------------
    # Slide 2: Project Scope & Framing
    # -------------------------------------------------------------
    slide2 = prs.slides.add_slide(blank_layout)
    add_header(slide2, "Project Context, Motivation & Methodological Scope")

    tb = slide2.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(6.0), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Agricultural Imperative:"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = SECONDARY

    bullets = [
        ("Eastern Gangetic Plains:", "Bihar and eastern UP form India's core food security frontier, but suffer from smallholder fragmentation and persistent yield gaps."),
        ("Proactive vs. Retrospective:", "Traditional post-harvest yield models provide zero decision support to farmers. Pre-sowing predictive analytics empowers proactive input optimization before planting."),
        ("Continuous vs. Categorical:", "Categorizing yields into risk tertiles (Low, Medium, High) matches real-world decision-making (insurance, credit, input adjustments) and dampens measurement noise.")
    ]
    for b_title, b_text in bullets:
        p = tf.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(13)
        p.font.color.rgb = TEXT_DARK
        p.space_before = Pt(10)

    tb2 = slide2.shapes.add_textbox(Inches(7.2), Inches(1.8), Inches(5.3), Inches(5.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "Accurate Project Scope & Framing:"
    p2.font.bold = True
    p2.font.size = Pt(16)
    p2.font.color.rgb = ACCENT

    framing_bullets = [
        ("Dataset Boundary:", "The dataset contains exclusively wheat records (5,689 fields). No competing crops (mustard, chickpea) exist in the data."),
        ("No Crop Recommendation Claim:", "The project is framed honestly and strictly as intra-crop wheat yield classification, NOT crop recommendation."),
        ("Syllabus Fidelity:", "Restricted exclusively to INT234 Unit III classification algorithms and metrics. No out-of-syllabus methods were introduced.")
    ]
    for b_title, b_text in framing_bullets:
        p = tf2.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(13)
        p.font.color.rgb = TEXT_DARK
        p.space_before = Pt(10)

    # -------------------------------------------------------------
    # Slide 3: Dataset & Temporal Leakage Prevention
    # -------------------------------------------------------------
    slide3 = prs.slides.add_slide(blank_layout)
    add_header(slide3, "Dataset Provenance & Rigorous Temporal Leakage Audit")

    tb = slide3.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.8), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "CIMMYT Dataset Profile:"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = SECONDARY

    ds_info = [
        ("Official Citation:", "Ajay, Anurag (2026). Field-level crop management, weather, and yield data of wheat from eastern India. CIMMYT Research Data & Software Repository Network, V1. DOI: https://doi.org/10.71682/10549424 (CC0 1.0 Universal)."),
        ("Sample Size:", "5,689 field-level monitoring records across 6 wheat crop years (2016-17 to 2021-22)."),
        ("Spatial Coverage:", "10 districts across Bihar (Begusarai, Bhojpur, Buxar, East Champaran, Lakhisarai, Madhepura, Muzaffarpur, Rohtas) and Eastern UP (Deoria, Kushinagar)."),
        ("Data Handling Strategy:", "BasalZn (~91.1% missing) excluded from input features as survey guide does not confirm blank implies zero application. GradeNPK excluded due to structural collinearity with BasalNPK. Prcp_Monsoon (19 nulls) imputed using training-set median.")
    ]
    for b_title, b_text in ds_info:
        p = tf.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(12)
        p.space_before = Pt(7)

    tb2 = slide3.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(5.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "Temporal Information Boundary (At Sowing):"
    p2.font.bold = True
    p2.font.size = Pt(16)
    p2.font.color.rgb = PRIMARY

    leak_info = [
        ("INCLUDED Pre-Sowing Predictors (16 raw):", "Year, District, LandType, SoilType, PreviousCrop, PrevCropResidue, CropEstablishment (ZT vs CT), Variety, VarietyClass, SeedRate, SowingDay, BasalDAP, BasalNPK, BasalMOP, Prcp_Monsoon, Prcp_Oct."),
        ("PURGED Post-Sowing Leakage:", "GrainYield (target), Field_ID, Latitude, Longitude, HarvestDate, HarvestDay, CropLodging, IrrigationNumber, FirstIrrigationDay-FifthIrrigationDay, Split1Urea-Split3Urea, HerbicideName, HerbicideType, WeedingNumber, Nov_Tmax-EDD_Apr.")
    ]
    for b_title, b_text in leak_info:
        p = tf2.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(12)
        p.space_before = Pt(10)

    # -------------------------------------------------------------
    # Slide 4: Target Tertiles & Splitting Protocol
    # -------------------------------------------------------------
    slide4 = prs.slides.add_slide(blank_layout)
    add_header(slide4, "Target Tertile Formulation & Stratified Splitting Protocol")

    # Add Figure 1
    fig1_path = os.path.join(figures_dir, '01_target_distribution.png')
    if os.path.exists(fig1_path):
        slide4.shapes.add_picture(fig1_path, Inches(0.8), Inches(1.8), width=Inches(6.8))

    tb = slide4.shapes.add_textbox(Inches(7.8), Inches(1.8), Inches(4.8), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Target Partitioning Design:"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = SECONDARY

    split_pts = [
        ("Empirical Tertile Protocol:", "Target tertiles are data-derived categories reflecting empirical yield terciles in eastern India, not official agronomic standards."),
        ("Zero Data Snooping:", "Cutoffs derived strictly on the 80% training set (4,551 fields) after stratified split (`random_state=42`)."),
        ("Lower Cutoff (c1):", "3.6000 t/ha (Yields <= 3.60 t/ha are classified as 'Low')."),
        ("Upper Cutoff (c2):", "4.6700 t/ha (Yields > 4.67 t/ha are classified as 'High')."),
        ("Intermediate Zone:", "3.60 to 4.67 t/ha classified as 'Medium'."),
        ("Balanced Classes:", "Train: Low=1,527, Med=1,518, High=1,506 (33.3% each). Test: Low=382, Med=380, High=376 (33.3% each).")
    ]
    for b_title, b_text in split_pts:
        p = tf.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(12)
        p.space_before = Pt(7)

    # -------------------------------------------------------------
    # Slide 5: Agronomic Exploratory Insights
    # -------------------------------------------------------------
    slide5 = prs.slides.add_slide(blank_layout)
    add_header(slide5, "Key Agronomic Findings: Sowing Timing & Spatial Patterns")

    fig2_path = os.path.join(figures_dir, '02_yield_vs_sowing_day.png')
    fig3_path = os.path.join(figures_dir, '03_district_yield_distribution.png')

    if os.path.exists(fig2_path):
        slide5.shapes.add_picture(fig2_path, Inches(0.8), Inches(1.8), width=Inches(5.6))
    if os.path.exists(fig3_path):
        slide5.shapes.add_picture(fig3_path, Inches(6.8), Inches(1.8), width=Inches(5.7))

    tb = slide5.shapes.add_textbox(Inches(0.8), Inches(5.8), Inches(11.7), Inches(1.3))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Key EDA Insights: (1) Sowing past Julian Day 330 shows an observational decline in yield, consistent with terminal heat stress literature. (2) South Bihar (Buxar, Rohtas) enjoys superior irrigation infrastructure and clay-loam soils compared to flood-prone northern districts."
    p.font.size = Pt(12.5)
    p.font.bold = True
    p.font.color.rgb = SECONDARY

    # -------------------------------------------------------------
    # Slide 6: Model Comparison Master Table
    # -------------------------------------------------------------
    slide6 = prs.slides.add_slide(blank_layout)
    add_header(slide6, "Syllabus Classification Models: Master Performance Comparison")

    rows = 6
    cols = 8
    table_shape = slide6.shapes.add_table(rows, cols, Inches(0.8), Inches(1.8), Inches(11.733), Inches(3.2))
    table = table_shape.table

    headers = ["Model Name", "Train Acc", "Test Acc", "Precision", "Recall", "F1-Score", "Log Loss", "ROC AUC"]
    for i, h in enumerate(headers):
        cell = table.cell(0, i)
        cell.text = h
        cell.fill.solid()
        cell.fill.fore_color.rgb = PRIMARY
        for p in cell.text_frame.paragraphs:
            p.alignment = PP_ALIGN.CENTER
            for r in p.runs:
                r.font.bold = True
                r.font.size = Pt(12)
                r.font.color.rgb = WHITE

    model_data = [
        ("Support Vector Machine (RBF, C=2.0)", "77.41%", "69.42%", "0.6886", "0.6943", "0.6893", "0.7216", "0.8458"),
        ("K-Nearest Neighbors (k=13, Manhattan)", "98.44%", "68.63%", "0.6813", "0.6862", "0.6823", "2.1825", "0.8381"),
        ("Decision Tree (Depth=7, Gini)", "69.17%", "64.32%", "0.6500", "0.6436", "0.6447", "1.1241", "0.8128"),
        ("Logistic Regression (Multinomial, L2)", "63.94%", "62.04%", "0.6108", "0.6207", "0.6106", "0.8130", "0.8004"),
        ("Gaussian Naive Bayes", "51.15%", "50.70%", "0.4966", "0.5084", "0.4121", "12.9987", "0.7091")
    ]
    for r_idx, r_vals in enumerate(model_data, start=1):
        bg = RGBColor(240, 244, 248) if r_idx % 2 == 1 else WHITE
        for c_idx, val in enumerate(r_vals):
            cell = table.cell(r_idx, c_idx)
            cell.text = val
            cell.fill.solid()
            cell.fill.fore_color.rgb = bg
            for p in cell.text_frame.paragraphs:
                p.alignment = PP_ALIGN.CENTER
                for r in p.runs:
                    r.font.size = Pt(11)
                    if r_idx == 1:
                        r.font.bold = True

    tb = slide6.shapes.add_textbox(Inches(0.8), Inches(5.2), Inches(11.733), Inches(2.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Leakage-Free Cross-Validation & Parameter Justification:"
    p.font.bold = True
    p.font.size = Pt(13.5)
    p.font.color.rgb = SECONDARY

    pts = [
        "Leakage-Free Pipeline CV: Imputer, scaler, and encoder are encapsulated in scikit-learn Pipelines, fitted strictly per fold on raw training features.",
        "KNN Parameter Choice: k=13 (Manhattan, distance) selected as the unambiguous empirical CV-winning setting (69.22% mean accuracy).",
        "Decision Tree Choice: Depth 7 retained over depth 9 (63.94% vs 65.02%, Delta=1.08%) as a deliberate simplicity trade-off to prevent excessive branching on survey noise and preserve interpretability.",
        "SVM Parameter Choice: C=2.0 selected over C=3.0 (69.83% vs 69.96%, Delta=0.13%) to enforce wider soft margins against survey noise. SVM achieved top test metrics (69.42% acc, 0.7216 log-loss, 0.8458 AUC)."
    ]
    for pt in pts:
        p = tf.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(11)
        p.space_before = Pt(2.5)

    # -------------------------------------------------------------
    # Slide 7: Confusion Matrices Analysis
    # -------------------------------------------------------------
    slide7 = prs.slides.add_slide(blank_layout)
    add_header(slide7, "Confusion Matrix Diagnostics & Error Pattern Analysis")

    fig5_path = os.path.join(figures_dir, '05_confusion_matrices.png')
    if os.path.exists(fig5_path):
        slide7.shapes.add_picture(fig5_path, Inches(0.8), Inches(1.7), width=Inches(7.2))

    tb = slide7.shapes.add_textbox(Inches(8.2), Inches(1.8), Inches(4.5), Inches(5.2))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Diagnostic Patterns:"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = SECONDARY

    diag_pts = [
        ("High Extreme Discrimination:", "SVM correctly identifies 79.6% of Low yield fields (304/382) and 76.6% of High yield fields (288/376)."),
        ("Low Catastrophic Errors:", "Misclassifying actual High as Low (5.9%), or actual Low as High (5.8%), is restricted to under 6% in SVM and KNN."),
        ("The 'Medium' Class Challenge:", "All models experience highest ambiguity on the boundary-adjacent 'Medium' category (3.60-4.67 t/ha) due to transitional overlap."),
        ("Naive Bayes Collapse:", "Naive Bayes completely fails on Medium (recall 1.3%), demonstrating severe bimodal polarization.")
    ]
    for b_title, b_text in diag_pts:
        p = tf.add_paragraph()
        p.text = f"• {b_title} {b_text}"
        p.font.size = Pt(12)
        p.space_before = Pt(8)

    # -------------------------------------------------------------
    # Slide 8: ROC Curves & Feature Importance
    # -------------------------------------------------------------
    slide8 = prs.slides.add_slide(blank_layout)
    add_header(slide8, "Multiclass ROC Curves & Feature Importance Ranking")

    fig7_path = os.path.join(figures_dir, '07_roc_curves.png')
    fig8_path = os.path.join(figures_dir, '08_feature_importance_tree.png')

    if os.path.exists(fig7_path):
        slide8.shapes.add_picture(fig7_path, Inches(0.8), Inches(1.8), width=Inches(5.6))
    if os.path.exists(fig8_path):
        slide8.shapes.add_picture(fig8_path, Inches(6.8), Inches(1.8), width=Inches(5.7))

    tb = slide8.shapes.add_textbox(Inches(0.8), Inches(6.0), Inches(11.7), Inches(1.1))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Insights: (1) SVM (0.846) and KNN (0.838) exhibit dominant true positive rates across false positive thresholds. (2) SowingDay, District indicators, Variety, and Monsoon Rainfall are the top pre-sowing predictors."
    p.font.size = Pt(12)
    p.font.bold = True
    p.font.color.rgb = SECONDARY

    # -------------------------------------------------------------
    # Slide 9: Conclusion & Future Scope
    # -------------------------------------------------------------
    slide9 = prs.slides.add_slide(blank_layout)
    add_header(slide9, "Conclusion, Practical Implications & Future Directions")

    tb = slide9.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(5.8), Inches(5.0))
    tf = tb.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = "Conclusions & Practical Utility:"
    p.font.bold = True
    p.font.size = Pt(16)
    p.font.color.rgb = PRIMARY

    c_pts = [
        "Pre-Sowing Predictability: Achieving 69.42% 3-class yield categorization accuracy before sowing indicates that pre-sowing management and historical environmental conditions provide meaningful early signals of yield potential.",
        "Sowing Date Sensitivity: Sowing date is an influential pre-sowing predictor, consistent with established literature on terminal heat avoidance in eastern India.",
        "Empirical vs. Causal Framing: Results describe observational patterns and empirical risk categorization in the surveyed sample, rather than causal agronomic guarantees."
    ]
    for pt in c_pts:
        p = tf.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.space_before = Pt(8)

    tb2 = slide9.shapes.add_textbox(Inches(7.0), Inches(1.8), Inches(5.5), Inches(5.0))
    tf2 = tb2.text_frame
    tf2.word_wrap = True
    p2 = tf2.paragraphs[0]
    p2.text = "Future Scope & Validation:"
    p2.font.bold = True
    p2.font.size = Pt(16)
    p2.font.color.rgb = ACCENT

    f_pts = [
        "Temporal Validation via Future-Year Holdouts: Evaluating performance on future crop seasons to test temporal generalization under shifting climate regimes.",
        "Satellite Remote Sensing: Integrating Copernicus Sentinel-2 bare-soil reflectance indices (NDTI, SAVI) before sowing.",
        "Short-Range Weather Forecasts: Incorporating numerical weather predictions around the seed emergence window.",
        "Lightweight Decision Support Tools: Packaging pre-sowing classification models for localized risk awareness and agricultural extension."
    ]
    for pt in f_pts:
        p = tf2.add_paragraph()
        p.text = f"• {pt}"
        p.font.size = Pt(12)
        p.space_before = Pt(8)

    os.makedirs(os.path.dirname(output_pptx_path), exist_ok=True)
    prs.save(output_pptx_path)
    print(f"Presentation saved successfully to: {output_pptx_path}")

if __name__ == '__main__':
    script_dir = os.path.dirname(__file__)
    out_pptx = os.path.join(script_dir, '..', 'Report and Presentation', 'presentation', 'Wheat_Yield_Classification_Presentation.pptx')
    fig_dir = os.path.join(script_dir, '..', 'figures')
    create_deck(out_pptx, fig_dir)
