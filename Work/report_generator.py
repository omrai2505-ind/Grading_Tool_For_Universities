from fpdf import FPDF
import datetime

def generate_pdf_report(df, stats, file_path):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", 'B', 16)
    pdf.cell(0, 10, "Relative Grading System - Final Report", ln=True, align="C")
    
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 10, f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M')}", ln=True)
    pdf.ln(10)
    
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "Class Statistics", ln=True)
    pdf.set_font("Arial", size=12)
    pdf.cell(0, 8, f"Total Students: {stats.get('count', 0)}", ln=True)
    pdf.cell(0, 8, f"Mean Score: {stats.get('mean', 0):.2f}", ln=True)
    pdf.cell(0, 8, f"Std Deviation: {stats.get('std', 0):.2f}", ln=True)
    pdf.cell(0, 8, f"Max Score: {stats.get('max', 0):.2f}", ln=True)
    pdf.cell(0, 8, f"Min Score: {stats.get('min', 0):.2f}", ln=True)
    
    pdf.ln(10)
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "Grade Tally", ln=True)
    pdf.set_font("Arial", size=12)
    
    if "Grade" in df.columns:
        counts = df["Grade"].value_counts().to_dict()
        for g in ["A", "A-", "B", "B-", "C", "C-", "D", "F"]:
            pdf.cell(0, 8, f"{g}: {counts.get(g, 0)} student(s)", ln=True)
            
    pdf.output(file_path)
