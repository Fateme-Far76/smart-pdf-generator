import sys
import os
import tempfile
import platform
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from common.pdf_builder import generate_pdf as _shared_generate_pdf
from math import ceil
from reportlab.lib import colors
from pathlib import Path
from reportlab.platypus import Image

# Helper to get the path of the bundled font folder
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # temp folder used by PyInstaller
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def generate_page1(self, block, village, fid, farmer, contact):
    elements = []
    
    englishnormal = ParagraphStyle(
        name="eng",
        fontName="Helvetica",
        fontSize=10,
        alignment=0,
    )
    
    elements.append(Image("images/l2_title.PNG", width=20*cm, height=30))
    elements.append(Spacer(1, 10))


    # fetch values from the DataFrame
    survey_number = self.df_final_filtered["Survey Number"].astype(str).iloc[0] if not self.df_final_filtered.empty else ""
    event_date = self.df_final_filtered["Event occurred Date"].astype(str).iloc[0] if not self.df_final_filtered.empty else ""
    intimation_date = self.df_final_filtered["Date of Intimation"].astype(str).iloc[0] if not self.df_final_filtered.empty else ""
    crop = self.df_final_filtered["Crop"].astype(str).iloc[0] if not self.df_final_filtered.empty else ""
    
    data = [
        ["1.", Image(resource_path("images/1.PNG"), width=100, height=15), Paragraph(farmer, englishnormal), 
        "2.", Image("images/2.PNG", width=100, height=15), ""],

        ["3.", Image(resource_path("images/3.PNG"), width=100, height=15), Paragraph(village, englishnormal), 
        "4.", Image("images/4.PNG", width=100, height=15), Paragraph(block, englishnormal)],

        ["5.", Image(resource_path("images/5.PNG"), width=100, height=15), Paragraph("Hisar", englishnormal), 
        "6.", Image(resource_path("images/6.PNG"), width=100, height=15), Paragraph("HDFC Ergo", englishnormal)],

        ["7.", Image(resource_path("images/7.PNG"), width=100, height=20), Paragraph(crop, englishnormal), 
        "8.", Image(resource_path("images/2.PNG"), width=100, height=20), Paragraph(survey_number, englishnormal)],

        ["9.", Image(resource_path("images/9.PNG"), width=100, height=20), "", 
        "10.", Image(resource_path("images/10.PNG"), width=100, height=20), ""],

        ["11.", Image(resource_path("images/11.PNG"), width=100, height=20), "", 
        "12.", Image(resource_path("images/12.PNG"), width=100, height=20), Paragraph(event_date, englishnormal)],

        ["13.", Image(resource_path("images/13.PNG"), width=100, height=30), Paragraph(intimation_date, englishnormal), 
        "14.", Image(resource_path("images/14.PNG"), width=100, height=30), ""],

        ["15.", Image(resource_path("images/l2_15.PNG"), width=100, height=15), ""]
    ]


    # Set outer table column widths
    table = Table(data, colWidths=[
        1*cm, 4.2*cm, 4.74*cm, 1*cm, 4.2*cm, 4.74*cm
    ])



    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),

        # Merge mini-table area (columns 2-5 in row 8)
        ("SPAN", (2, 7), (5, 7)),

        # Merge notes area (columns 2–5 in row 9)
        ("SPAN", (2, 8), (5, 8)),

        # Remove padding from the mini-table container cell
        ("LEFTPADDING", (2, 7), (5, 7), 0),
        ("RIGHTPADDING", (2, 7), (5, 7), 0),
        ("TOPPADDING", (2, 7), (5, 7), 0),
        ("BOTTOMPADDING", (2, 7), (5, 7), 0),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 10))
    elements.append(Image("images/l2_middle.PNG", width=20*cm, height=100))
    
    info_style = ParagraphStyle(
        name="FarmerInfo",
        fontName="Helvetica-Bold",
        fontSize=8,
        alignment=0,  # Left-aligned
        spaceAfter=4
    )
    style = TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("BACKGROUND", (0, 0), (-1, 0), colors.white),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
    ])
    
    # Remove existing index column (we'll insert it manually)
    df_transformed = df_transformed.drop(columns=["क्र.सं."])
    data_rows = df_transformed.astype(str).values.tolist()

    # Pad first page if needed
    rows_per_page = 22
    if len(data_rows) < rows_per_page:
        blanks_needed = rows_per_page - len(data_rows)
        blank_row = [""] * (df_transformed.shape[1])
        data_rows += [blank_row.copy() for _ in range(blanks_needed)]

    # Add row numbers (index)
    for i, row in enumerate(data_rows):
        row.insert(0, str(i + 1))  # insert row number at the start

    # Header row
    header_row = [
        Image("images/col1.PNG", width=1.4*cm, height=20),
        Image("images/col2.PNG", width=2*cm, height=20),
        Image("images/col3.PNG", width=4.3*cm, height=20),
        Image("images/col4.PNG", width=3*cm, height=20),
        Image("images/col5.PNG", width=2.6*cm, height=20),
        Image("images/col6.PNG", width=2.8*cm, height=20),
        Image("images/col7.PNG", width=2.4*cm, height=20),
        Image("images/col8.PNG", width=2.4*cm, height=20)
    ]


    elements.append(PageBreak())
    return elements
            
            
            
                     
            
def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact):
    return _shared_generate_pdf(
        df_transformed, save_path, block, village, fid, farmer, contact,
        generate_page1, generate_page2
    )