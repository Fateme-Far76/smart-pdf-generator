import sys
import os
import tempfile
import platform
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from common.pdf_builder import generate_pdf2
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

def generate_page2(df_transformed, block, village, fid, farmer, contact, df_final_filtered):
    pass

def generate_page1(df_transformed, block, village, fid, farmer, contact, df_final_filtered):
    elements = []
    
    englishnormal = ParagraphStyle(
        name="eng",
        fontName="Helvetica",
        fontSize=10,
        alignment=0,
    )
    
    elements.append(Image(resource_path("images/l2_title.PNG"), width=20*cm, height=30))
    elements.append(Spacer(1, 10))


    # fetch values from the DataFrame
    survey_number = df_final_filtered["Survey Number"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    event_date = df_final_filtered["Event occurred Date"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    intimation_date = df_final_filtered["Date of Intimation"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    crop = df_final_filtered["Crop"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    Intimation_Application_id = df_final_filtered["Intimation_Application id"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    
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

        ["15.", Image(resource_path("images/l2_15.PNG"), width=100, height=15), Paragraph(Intimation_Application_id, englishnormal)]
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

    cell_style = ParagraphStyle(
        name="cell",
        fontName="Helvetica",
        fontSize=10,
        alignment=1,
        leading=12,
    )
    
    ROWS_TARGET = 10
    df = df_transformed.copy()
    
    # Compute the sum for crop area
    area_series= (
        pd.to_numeric(df["बीमित क्षेत्र"].replace("", 0), errors="coerce")
        .fillna(0.0)
    )
    total_area = float(area_series.sum())
    
    page_df= (
        df.reset_index(drop=True)
               .where(df.notna(), "")
               .astype(str)
    )
    
    # Split into first 10 and leftover
    page_display = page_df.iloc[:ROWS_TARGET].copy()
    leftover_df = page_df.iloc[ROWS_TARGET:].copy()
    
    # Pad to exactly 10 rows
    if len(page_display) < 10:
        blanks_needed = 10 - len(page_display)
        pad = pd.DataFrame([[""] * page_display.shape[1]] * blanks_needed, columns=page_display.columns)
        page_display = pd.concat([page_display, pad], ignore_index=True)
    
    # Header row
    header_row = [
        Image(resource_path("images/l2_table1.PNG"), width=1.4*cm, height=20),
        Image(resource_path("images/l2_table2.PNG"), width=2*cm, height=20),
        Image(resource_path("images/l2_table3.PNG"), width=4.3*cm, height=20),
        Image(resource_path("images/l2_table4.PNG"), width=3*cm, height=20),
        Image(resource_path("images/l2_table5.PNG"), width=2.6*cm, height=20),
        Image(resource_path("images/l2_table6.PNG"), width=2.8*cm, height=20),
        Image(resource_path("images/l2_table7.PNG"), width=2.4*cm, height=20),
        Image(resource_path("images/l2_table8.PNG"), width=2.4*cm, height=20)
    ]

    
    table_rows = []
    for i, row_vals in enumerate(page_display.values.tolist(), start=1):
        row_vals[0] = str(i)  # force first col to 1..10 shown
        # Wrap into Paragraphs for consistent spacing (optional)
        row_vals = [Paragraph(str(v), cell_style) for v in row_vals]
        table_rows.append(row_vals)

    data = [header_row] + table_rows
    total_label = Image(resource_path("images/l2_sum.PNG"), width=4*cm, height=20)
    total_value = Paragraph(f"<b>{total_area}</b>", cell_style)
    
    total_row = [""] * 8
    total_row[3] = total_label
    total_row[4] = total_value
    data.append(total_row)   
    
    table = Table(data, colWidths=[
        1.5*cm, 4.32*cm, 2.02*cm, 2.42*cm, 2.62*cm, 2*cm, 2*cm, 3*cm
    ], rowHeights=20)
    
    style = TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("GRID", (0, 0), (-1, -2), 0.5, colors.black),  # all grid except total row bottom
        ("GRID", (0, -1), (-1, -1), 0.5, colors.black), # grid for total row too
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # Make the 4th column (index 3) slightly smaller font (optional, like your earlier rule)
        ("FONTSIZE", (3, 0), (3, -1), 9),

        # TOTAL ROW MERGES:
        # Merge first four cells of the total row (0..3) to make whitespace block on left
        ("SPAN", (0, len(data)-1), (3, len(data)-1)),

        # You can also merge the last two cells if you want a wide remarks cell on total row:
        ("SPAN", (4, len(data)-1), (7, len(data)-1)),

        # Alignments for total row label/value
        ("ALIGN", (4, len(data)-1), (4, len(data)-1), "RIGHT"),
        ("ALIGN", (5, len(data)-1), (5, len(data)-1), "LEFT"),
        ("FONTNAME", (5, len(data)-1), (5, len(data)-1), "Helvetica-Bold"),
        ("BOTTOMPADDING", (0, len(data)-1), (-1, len(data)-1), 6),
        ("TOPPADDING", (0, len(data)-1), (-1, len(data)-1), 6),
    ])
    style.add("SPAN", (7, 1), (7, 10))
    table.setStyle(style)
    elements.append(table)
    elements.append(Spacer(1, 10))
    elements.append(Image(resource_path("images/l2_p1_below_table.PNG"), width=12*cm, height=20))
    elements.append(Spacer(1, 20))    
    # Example data (replace with your content)
    data_sign = [[
        Image(resource_path("images/l2_sign1.PNG"), width=4.9*cm, height=40),
        Image(resource_path("images/l2_sign2.PNG"), width=4.9*cm, height=40),
        Image(resource_path("images/l2_sign3.PNG"), width=4.9*cm, height=40),
        Image(resource_path("images/l2_sign4.PNG"), width=4.9*cm, height=40)  
    ]]

    # Create table
    table_sign = Table(data_sign, colWidths=[4.97*cm, 4.97*cm, 4.97*cm, 4.97*cm], rowHeights=40)
    elements.append(table_sign)
    elements.append(Image(resource_path("images/l2_bottom.PNG"), width=12*cm, height=30))

    if len(leftover_df) > 0:
        # e.g. start a new page, then:
        # elements.append(PageBreak())
        # generate_page2(leftover, block, village, fid, farmer, contact)
        pass

    return elements
                             
            
def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact, df_final_filtered):
    return generate_pdf2(
        df_transformed, save_path, block, village, fid, farmer, contact,
        df_final_filtered, generate_page1
    )