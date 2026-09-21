from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from pathlib import Path


# ============================================================
# PROJECT REPORT GENERATOR
# Student Learning Pattern Analysis Using Clustering
# Group - 11
# ============================================================

OUTPUT_FILE = "Student_Learning_Pattern_Analysis_Project_Report.docx"


# ------------------------------------------------------------
# Create Document
# ------------------------------------------------------------

doc = Document()

# Page settings
for section in doc.sections:
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)


# ------------------------------------------------------------
# Styles
# ------------------------------------------------------------

styles = doc.styles

styles["Normal"].font.name = "Times New Roman"
styles["Normal"].font.size = Pt(12)

styles["Normal"].paragraph_format.space_after = Pt(6)
styles["Normal"].paragraph_format.line_spacing = 1.15

for style_name in ["Title", "Heading 1", "Heading 2"]:
    styles[style_name].font.name = "Times New Roman"


# ------------------------------------------------------------
# Helper Functions
# ------------------------------------------------------------

def centered_text(text, size=12, bold=False, space_after=8):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(space_after)

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(size)
    run.bold = bold

    return p


def normal_paragraph(text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.space_after = Pt(7)
    p.paragraph_format.line_spacing = 1.15

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

    return p


def bullet(text):
    p = doc.add_paragraph(style="List Bullet")
    p.paragraph_format.space_after = Pt(4)

    run = p.add_run(text)
    run.font.name = "Times New Roman"
    run.font.size = Pt(12)

    return p


def add_heading(text, level=1):
    p = doc.add_heading(text, level=level)

    for run in p.runs:
        run.font.name = "Times New Roman"

        if level == 1:
            run.font.size = Pt(15)
        else:
            run.font.size = Pt(13)

    return p


def screenshot_placeholder(title, height_lines=8):
    """
    Creates a simple placeholder where the user can insert
    the original project screenshot later.
    """

    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("\n")
    run.font.size = Pt(10)

    run = p.add_run(
        "[ INSERT YOUR ORIGINAL PROJECT SCREENSHOT HERE ]"
    )
    run.bold = True
    run.font.name = "Times New Roman"
    run.font.size = Pt(13)

    for _ in range(height_lines):
        p.add_run("\n")

    p.add_run("\n")

    caption = doc.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = caption.add_run(title)
    run.font.name = "Times New Roman"
    run.font.size = Pt(10)
    run.italic = True

    doc.add_paragraph()


# ============================================================
# COVER PAGE
# ============================================================

centered_text("GROUP – 11", 22, True, 25)

centered_text("PROJECT REPORT", 26, True, 25)

centered_text(
    "STUDENT LEARNING PATTERN ANALYSIS",
    19,
    True,
    5
)

centered_text(
    "USING CLUSTERING",
    19,
    True,
    30
)

centered_text(
    "DAY 2 – IMPLEMENTATION & FINAL REPORT",
    15,
    True,
    40
)

centered_text("TEAM MEMBERS", 14, True, 12)

centered_text(
    "C Srikishore – 610524091079",
    13,
    True,
    5
)

centered_text(
    "Rithik Vishnu – 610524091059",
    13,
    True,
    5
)

centered_text(
    "D Sabarivel – 610524091060",
    13,
    True,
    30
)

centered_text(
    "Department of Artificial Intelligence and Data Science",
    13,
    False,
    8
)

centered_text(
    "2026",
    13,
    False,
    0
)

doc.add_page_break()


# ============================================================
# CERTIFICATE
# ============================================================

add_heading("CERTIFICATE")

normal_paragraph(
    "This is to certify that the project report entitled "
    "“STUDENT LEARNING PATTERN ANALYSIS USING CLUSTERING” "
    "is a bonafide record of the project work carried out by "
    "C Srikishore (610524091079), Rithik Vishnu (610524091059), "
    "and D Sabarivel (610524091060) as part of their academic "
    "project work."
)

normal_paragraph(
    "The project focuses on analysing student learning patterns "
    "and academic performance using clustering techniques."
)

doc.add_paragraph("\n\n\n")

table = doc.add_table(rows=2, cols=2)
table.alignment = WD_TABLE_ALIGNMENT.CENTER
table.style = "Table Grid"

table.cell(0, 0).text = "Project Guide"
table.cell(0, 1).text = "Head of the Department"

table.cell(1, 0).text = "\n\nSignature"
table.cell(1, 1).text = "\n\nSignature"

doc.add_page_break()


# ============================================================
# DECLARATION
# ============================================================

add_heading("DECLARATION")

normal_paragraph(
    "We hereby declare that the project entitled "
    "“STUDENT LEARNING PATTERN ANALYSIS USING CLUSTERING” "
    "is our original academic project work. The project was "
    "developed to study student learning patterns using data "
    "analysis and the K-Means clustering algorithm."
)

doc.add_paragraph()

for member in [
    "C Srikishore – 610524091079",
    "Rithik Vishnu – 610524091059",
    "D Sabarivel – 610524091060"
]:
    normal_paragraph(member)
    normal_paragraph("Signature: ______________________________")

doc.add_page_break()


# ============================================================
# ACKNOWLEDGEMENT
# ============================================================

add_heading("ACKNOWLEDGEMENT")

normal_paragraph(
    "We sincerely thank our institution and the Department of "
    "Artificial Intelligence and Data Science for providing us "
    "with the opportunity to complete this project."
)

normal_paragraph(
    "We express our gratitude to our project guide and faculty "
    "members for their valuable guidance, support, and suggestions "
    "during the development of the project."
)

normal_paragraph(
    "We also thank our friends and classmates for their support "
    "and encouragement in completing this project successfully."
)


# ============================================================
# ABSTRACT
# ============================================================

add_heading("ABSTRACT")

normal_paragraph(
    "Student learning and academic performance depend on several "
    "factors such as study hours, attendance, assignment scores, "
    "internal marks, and previous academic performance. Analysing "
    "these factors manually can be difficult when there are many "
    "students."
)

normal_paragraph(
    "This project uses K-Means clustering to group students with "
    "similar learning and academic characteristics. The data is "
    "cleaned, prepared, and analysed before clustering. The results "
    "are presented using simple tables and graphs."
)

normal_paragraph(
    "The system helps teachers understand different student "
    "performance groups and helps students understand their own "
    "learning pattern."
)

doc.add_page_break()


# ============================================================
# TABLE OF CONTENTS
# ============================================================

add_heading("TABLE OF CONTENTS")

contents = [
    "1. Introduction",
    "2. Problem Statement",
    "3. Objectives",
    "4. Scope of the Project",
    "5. Existing System",
    "6. Proposed System",
    "7. Technologies Used",
    "8. Dataset Description",
    "9. Methodology",
    "10. Data Preprocessing",
    "11. K-Means Clustering",
    "12. System Modules",
    "13. System Architecture",
    "14. Implementation",
    "15. Results",
    "16. Testing",
    "17. Advantages",
    "18. Limitations",
    "19. Future Enhancement",
    "20. Conclusion",
    "21. References",
    "22. Appendix – Screenshots"
]

for item in contents:
    bullet(item)

doc.add_page_break()


# ============================================================
# 1. INTRODUCTION
# ============================================================

add_heading("1. INTRODUCTION")

normal_paragraph(
    "Student learning patterns can be understood by studying "
    "academic and study-related information. Different students "
    "may have different combinations of study hours, attendance, "
    "assignment performance, internal marks, and previous GPA."
)

normal_paragraph(
    "Clustering is an unsupervised machine learning technique "
    "that groups similar data points. In this project, K-Means "
    "clustering is used to group students based on their learning "
    "and academic characteristics."
)


# ============================================================
# 2. PROBLEM STATEMENT
# ============================================================

add_heading("2. PROBLEM STATEMENT")

normal_paragraph(
    "Teachers may have to analyse a large amount of student "
    "information manually. It can be difficult to identify "
    "common learning patterns and performance groups by looking "
    "at individual records."
)

normal_paragraph(
    "The project aims to provide a simple clustering-based system "
    "that groups students with similar academic and learning "
    "characteristics."
)


# ============================================================
# 3. OBJECTIVES
# ============================================================

add_heading("3. OBJECTIVES")

objectives = [
    "To analyse student academic and learning data.",
    "To clean and prepare the student dataset.",
    "To select useful features for analysis.",
    "To apply K-Means clustering.",
    "To group students with similar learning patterns.",
    "To visualize the clustering results.",
    "To support teachers in understanding student performance."
]

for item in objectives:
    bullet(item)


# ============================================================
# 4. SCOPE
# ============================================================

add_heading("4. SCOPE OF THE PROJECT")

normal_paragraph(
    "The project can be used as a basic academic analysis system "
    "for educational institutions. It can be extended later with "
    "real-time academic data and additional machine learning "
    "methods."
)

scope_items = [
    "Student performance analysis",
    "Learning pattern identification",
    "Attendance analysis",
    "Assignment and internal mark analysis",
    "Student grouping",
    "Graphical visualization",
    "Academic report generation"
]

for item in scope_items:
    bullet(item)


# ============================================================
# 5. EXISTING SYSTEM
# ============================================================

add_heading("5. EXISTING SYSTEM")

normal_paragraph(
    "In a traditional system, student marks, attendance, and "
    "other academic records are commonly reviewed manually or "
    "through spreadsheets. This can become time-consuming when "
    "the number of students increases."
)

existing = [
    "Manual comparison requires more time.",
    "Similar students are not automatically grouped.",
    "Large datasets are difficult to analyse quickly.",
    "Patterns between different academic factors may be difficult to identify."
]

for item in existing:
    bullet(item)


# ============================================================
# 6. PROPOSED SYSTEM
# ============================================================

add_heading("6. PROPOSED SYSTEM")

normal_paragraph(
    "The proposed system uses K-Means clustering to analyse "
    "student learning patterns. Student records are loaded, "
    "cleaned, standardized, and processed by the clustering "
    "algorithm. The resulting groups are displayed using tables "
    "and graphs."
)

proposed = [
    "Automatic grouping of similar students",
    "Simple preprocessing",
    "K-Means clustering",
    "Performance comparison",
    "Interactive visualization",
    "Easy interpretation"
]

for item in proposed:
    bullet(item)


# ============================================================
# 7. TECHNOLOGIES USED
# ============================================================

add_heading("7. TECHNOLOGIES USED")

table = doc.add_table(rows=1, cols=2)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER

table.cell(0, 0).text = "Technology"
table.cell(0, 1).text = "Purpose"

technologies = [
    ("Python", "Main programming language"),
    ("Pandas", "Data handling and analysis"),
    ("NumPy", "Numerical calculations"),
    ("Scikit-learn", "K-Means clustering and preprocessing"),
    ("Streamlit", "Web application dashboard"),
    ("Plotly", "Interactive graphs"),
    ("SQLite", "Data storage"),
    ("OpenPyXL", "Excel report generation")
]

for technology, purpose in technologies:
    cells = table.add_row().cells
    cells[0].text = technology
    cells[1].text = purpose


# ============================================================
# 8. DATASET
# ============================================================

add_heading("8. DATASET DESCRIPTION")

normal_paragraph(
    "The dataset contains student academic and learning "
    "information. The main features used in the project are "
    "listed below."
)

table = doc.add_table(rows=1, cols=3)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER

table.cell(0, 0).text = "Feature"
table.cell(0, 1).text = "Description"
table.cell(0, 2).text = "Type"

dataset_features = [
    ("Student_ID", "Unique student identification", "Text"),
    ("Study_Hours", "Study hours", "Numerical"),
    ("Attendance", "Attendance percentage", "Numerical"),
    ("Assignment_Score", "Assignment performance", "Numerical"),
    ("Internal_Marks", "Internal examination marks", "Numerical"),
    ("Previous_GPA", "Previous academic GPA", "Numerical"),
    ("Extracurricular_Hours", "Extracurricular activity hours", "Numerical")
]

for feature, description, data_type in dataset_features:
    cells = table.add_row().cells
    cells[0].text = feature
    cells[1].text = description
    cells[2].text = data_type


# ============================================================
# 9. METHODOLOGY
# ============================================================

add_heading("9. METHODOLOGY")

methodology = [
    "Collect student data",
    "Load the dataset",
    "Clean the data",
    "Handle missing values",
    "Select relevant features",
    "Standardize numerical features",
    "Apply K-Means clustering",
    "Assign cluster labels",
    "Analyse clusters",
    "Display results"
]

for item in methodology:
    bullet(item)


# ============================================================
# 10. DATA PREPROCESSING
# ============================================================

add_heading("10. DATA PREPROCESSING")

normal_paragraph(
    "Before clustering, the dataset is checked for missing and "
    "invalid values. Numerical features are converted into "
    "suitable numeric formats and missing values are handled."
)

normal_paragraph(
    "The selected features are standardized so that features with "
    "different numerical scales can contribute fairly to the "
    "clustering process."
)


# ============================================================
# 11. K-MEANS
# ============================================================

add_heading("11. K-MEANS CLUSTERING")

normal_paragraph(
    "K-Means is an unsupervised machine learning algorithm used "
    "to divide data into a selected number of groups called "
    "clusters. Each student is assigned to the cluster whose "
    "centre is closest to the student's feature values."
)

add_heading("11.1 Working Steps", 2)

kmeans_steps = [
    "Choose the number of clusters K.",
    "Initialize cluster centres.",
    "Calculate distances between students and centres.",
    "Assign students to the nearest centre.",
    "Recalculate cluster centres.",
    "Repeat until the clusters become stable."
]

for item in kmeans_steps:
    bullet(item)


# ============================================================
# 12. SYSTEM MODULES
# ============================================================

add_heading("12. SYSTEM MODULES")

modules = [
    ("Data Module", "Loads and manages student data."),
    ("Preprocessing Module", "Cleans and prepares data."),
    ("Clustering Module", "Applies K-Means clustering."),
    ("Analysis Module", "Analyses student performance."),
    ("Visualization Module", "Displays charts and graphs."),
    ("Student Module", "Shows individual student information."),
    ("Teacher Module", "Shows class and student information."),
    ("Report Module", "Generates academic reports.")
]

for module, description in modules:
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(5)

    r = p.add_run(module + ": ")
    r.bold = True
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)

    r = p.add_run(description)
    r.font.name = "Times New Roman"
    r.font.size = Pt(12)


# ============================================================
# 13. SYSTEM ARCHITECTURE
# ============================================================

add_heading("13. SYSTEM ARCHITECTURE")

centered_text("STUDENT DATASET", 14, True, 4)
centered_text("↓", 16, True, 4)
centered_text("DATA PREPROCESSING", 14, True, 4)
centered_text("↓", 16, True, 4)
centered_text("FEATURE SCALING", 14, True, 4)
centered_text("↓", 16, True, 4)
centered_text("K-MEANS CLUSTERING", 14, True, 4)
centered_text("↓", 16, True, 4)
centered_text("CLUSTER ANALYSIS", 14, True, 4)
centered_text("↓", 16, True, 4)
centered_text("DASHBOARD / REPORT", 14, True, 10)


# ============================================================
# 14. IMPLEMENTATION
# ============================================================

add_heading("14. IMPLEMENTATION")

normal_paragraph(
    "The application is implemented using Python and Streamlit. "
    "Pandas is used for dataset handling, Scikit-learn is used "
    "for preprocessing and K-Means clustering, and Plotly is "
    "used for interactive visualization."
)

normal_paragraph(
    "The application provides separate student and teacher views. "
    "Students can view their performance, while teachers can "
    "analyse multiple student records."
)

add_heading("14.1 Student View", 2)

student_features = [
    "Student login",
    "Academic performance summary",
    "Attendance and GPA information",
    "Performance group",
    "Graphs",
    "Personal report"
]

for item in student_features:
    bullet(item)

add_heading("14.2 Teacher View", 2)

teacher_features = [
    "Teacher login",
    "Total student count",
    "Performance group distribution",
    "Student performance table",
    "Individual student analysis",
    "Interactive graphs",
    "Excel report"
]

for item in teacher_features:
    bullet(item)


# ============================================================
# 15. RESULTS
# ============================================================

add_heading("15. RESULTS")

normal_paragraph(
    "The system groups students according to similarities in "
    "their academic and learning-related features. The resulting "
    "clusters can be compared using attendance, study hours, "
    "assignment scores, internal marks, and previous GPA."
)

normal_paragraph(
    "The graphs make it easier to understand student distribution "
    "and compare different performance groups."
)

add_heading("15.1 K-Means Cluster Visualization", 2)

screenshot_placeholder(
    "Figure 1: K-Means student cluster visualization."
)

add_heading("15.2 Project Dashboard", 2)

screenshot_placeholder(
    "Figure 2: EduPulse AI dashboard."
)

add_heading("15.3 Performance Analysis Graph", 2)

screenshot_placeholder(
    "Figure 3: Student performance analysis."
)

add_heading("15.4 Student Dashboard", 2)

screenshot_placeholder(
    "Figure 4: Student performance dashboard."
)

add_heading("15.5 Teacher Dashboard", 2)

screenshot_placeholder(
    "Figure 5: Teacher intelligence dashboard."
)


# ============================================================
# 16. TESTING
# ============================================================

add_heading("16. TESTING")

table = doc.add_table(rows=1, cols=4)
table.style = "Table Grid"
table.alignment = WD_TABLE_ALIGNMENT.CENTER

headers = [
    "Test Case",
    "Input",
    "Expected Result",
    "Status"
]

for i, header in enumerate(headers):
    table.cell(0, i).text = header

test_cases = [
    ("Dataset Loading", "Valid CSV", "Dataset loads successfully", "Pass"),
    ("Data Cleaning", "Missing values", "Values are handled", "Pass"),
    ("K-Means", "Valid features", "Clusters are generated", "Pass"),
    ("Student Login", "Valid credentials", "Student dashboard opens", "Pass"),
    ("Teacher Login", "Valid credentials", "Teacher dashboard opens", "Pass"),
    ("Graph Display", "Valid data", "Graphs are displayed", "Pass"),
    ("Excel Report", "Report request", "Excel file is generated", "Pass")
]

for test_case, input_data, expected, status in test_cases:
    cells = table.add_row().cells

    cells[0].text = test_case
    cells[1].text = input_data
    cells[2].text = expected
    cells[3].text = status


# ============================================================
# 17. ADVANTAGES
# ============================================================

add_heading("17. ADVANTAGES")

advantages = [
    "Simple and easy to use.",
    "Reduces manual grouping work.",
    "Uses machine learning for student grouping.",
    "Analyses multiple academic features.",
    "Provides graphical visualization.",
    "Helps identify learning patterns."
]

for item in advantages:
    bullet(item)


# ============================================================
# 18. LIMITATIONS
# ============================================================

add_heading("18. LIMITATIONS")

limitations = [
    "Results depend on the quality of the dataset.",
    "K-Means requires a suitable number of clusters.",
    "Clusters show patterns but do not explain every reason for student performance.",
    "The project is an academic prototype."
]

for item in limitations:
    bullet(item)


# ============================================================
# 19. FUTURE ENHANCEMENT
# ============================================================

add_heading("19. FUTURE ENHANCEMENT")

future = [
    "Use real-time college data.",
    "Add subject-wise analysis.",
    "Add semester-wise records.",
    "Add automated teacher notifications.",
    "Add mobile application support.",
    "Add more machine learning algorithms."
]

for item in future:
    bullet(item)


# ============================================================
# 20. CONCLUSION
# ============================================================

add_heading("20. CONCLUSION")

normal_paragraph(
    "The Student Learning Pattern Analysis using Clustering "
    "project demonstrates how machine learning can be used to "
    "group students according to similar learning and academic "
    "characteristics."
)

normal_paragraph(
    "Using data preprocessing, K-Means clustering, and "
    "visualization, the system provides a simple method for "
    "understanding student performance patterns."
)

normal_paragraph(
    "The project can be extended in the future with real-time "
    "data and additional features to provide better academic "
    "support."
)


# ============================================================
# 21. REFERENCES
# ============================================================

add_heading("21. REFERENCES")

references = [
    "Python Documentation",
    "Pandas Documentation",
    "NumPy Documentation",
    "Scikit-learn Documentation",
    "Streamlit Documentation",
    "Plotly Documentation"
]

for item in references:
    bullet(item)


# ============================================================
# 22. APPENDIX
# ============================================================

doc.add_page_break()

add_heading("22. APPENDIX – PROJECT SCREENSHOTS")

normal_paragraph(
    "The following section can be used to add the original "
    "screenshots of the implemented project. Replace each "
    "placeholder with the corresponding screenshot before "
    "final submission."
)

screenshots = [
    "Figure A1 – Login Page",
    "Figure A2 – Student Dashboard",
    "Figure A3 – Teacher Dashboard",
    "Figure A4 – K-Means Cluster Graph",
    "Figure A5 – Performance Distribution Graph",
    "Figure A6 – Individual Student Analysis",
    "Figure A7 – Excel Report"
]

for title in screenshots:
    add_heading(title, 2)
    screenshot_placeholder(
        f"{title} – Original Project Screenshot"
    )


# ============================================================
# Footer with page numbers
# ============================================================

# Simple footer

for section in doc.sections:
    footer = section.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    run = p.add_run("Student Learning Pattern Analysis Using Clustering")
    run.font.name = "Times New Roman"
    run.font.size = Pt(9)

# ============================================================
# Save
# ============================================================

doc.save(OUTPUT_FILE)

print()
print("=" * 60)
print("WORD REPORT CREATED SUCCESSFULLY")
print("=" * 60)
print(f"File: {Path(OUTPUT_FILE).resolve()}")
print()
print("Open the .docx file in Microsoft Word.")
print("Replace the screenshot placeholders with your original")
print("project screenshots before final submission.")
print("=" * 60)