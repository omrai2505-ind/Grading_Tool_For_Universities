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
            
    # Student Breakdowns
    pdf.add_page()
    pdf.set_font("Arial", 'B', 14)
    pdf.cell(0, 10, "Student Assessment Breakdowns (Scaled)", ln=True)
    pdf.set_font("Arial", size=10)
    
    scaled_cols = [c for c in df.columns if c.startswith("Final ")]
    if scaled_cols and "Roll_No" in df.columns:
        cols_to_print = ["Roll_No"]
        if "Name" in df.columns:
            cols_to_print.append("Name")
        cols_to_print += scaled_cols + ["Total_Score"]
        if "Grade" in df.columns:
            cols_to_print.append("Grade")
        if "Marks_to_Next_Grade" in df.columns:
            cols_to_print.append("Marks_to_Next_Grade")
            
        max_width = 190.0
        col_w = max_width / max(1, len(cols_to_print))
        
        pdf.set_font("Arial", 'B', 8)
        for c in cols_to_print:
            disp_name = c.replace("Final ", "")
            pdf.cell(col_w, 8, disp_name[:12], border=1, align="C")
        pdf.ln()
        
        pdf.set_font("Arial", size=8)
        for _, row in df.iterrows():
            for c in cols_to_print:
                val = row.get(c, "")
                if isinstance(val, (float, int)) and c != "Roll_No":
                    val_str = f"{val:.1f}"
                else:
                    val_str = str(val)[:15]
                pdf.cell(col_w, 6, val_str, border=1, align="C")
            pdf.ln()

    pdf.output(file_path)
