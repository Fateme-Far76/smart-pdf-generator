import sys
import pandas as pd
import os
import tempfile
import platform
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from math import ceil
from reportlab.lib import colors
from pathlib import Path
from common.pdf_builder import generate_pdf as _shared_generate_pdf
from reportlab.platypus import Image, Paragraph, Table, TableStyle, PageBreak, Spacer

def generate_page1(df_final_filtered, block, village, fid, farmer, contact):
    elements = []
    
    englishnormal = ParagraphStyle(
        name="eng",
        fontName="Helvetica",
        fontSize=10,
        alignment=0,
    )
    
    elements.append(Image("images/title.PNG", width=20*cm, height=30))
    elements.append(Spacer(1, 10))
    
    styles = getSampleStyleSheet()
    # Mini-table for row 15
    mini_table_data = [
        [Paragraph("1.", styles["Normal"]), Paragraph("2.", styles["Normal"])],
        [Paragraph("3.", styles["Normal"]), Paragraph("4.", styles["Normal"])],
        [Paragraph("5.", styles["Normal"]), Paragraph("6.", styles["Normal"])]
    ]

    # Mini-table for row 15, col 3 (spans cols 2–5: total width = 270)
    mini_table = Table(mini_table_data, colWidths=[7.34*cm, 7.34*cm])
    mini_table._argW = [7.34*cm, 7.34*cm]  # force the width

    mini_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.25, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))


    # fetch values from the DataFrame
    survey_number = df_final_filtered["Survey Number"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    event_date = df_final_filtered["Event occurred Date"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    intimation_date = df_final_filtered["Date of Intimation"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    crop = df_final_filtered["Crop"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    
    data = [
        ["1.", Image("images/1.PNG", width=100, height=15), Paragraph(farmer, englishnormal), 
        "2.", Image("images/2.PNG", width=100, height=15), ""],

        ["3.", Image("images/3.PNG", width=100, height=15), Paragraph(village, englishnormal), 
        "4.", Image("images/4.PNG", width=100, height=15), Paragraph(block, englishnormal)],

        ["5.", Image("images/5.PNG", width=100, height=15), Paragraph("Hisar", englishnormal), 
        "6.", Image("images/6.PNG", width=100, height=15), Paragraph("HDFC Ergo", englishnormal)],

        ["7.", Image("images/7.PNG", width=100, height=20), Paragraph(crop, englishnormal), 
        "8.", Image("images/2.PNG", width=100, height=20), Paragraph(survey_number, englishnormal)],

        ["9.", Image("images/9.PNG", width=100, height=20), "", 
        "10.", Image("images/10.PNG", width=100, height=20), ""],

        ["11.", Image("images/11.PNG", width=100, height=20), "", 
        "12.", Image("images/12.PNG", width=100, height=20), Paragraph(event_date, englishnormal)],

        ["13.", Image("images/13.PNG", width=100, height=30), Paragraph(intimation_date, englishnormal), 
        "14.", Image("images/14.PNG", width=100, height=30), ""],

        ["15.", Image("images/15.PNG", width=100, height=40), mini_table],

        ["16.", Image("images/16.PNG", width=100, height=15), ""]
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
    elements.append(Image("images/rest_page1.PNG", width=20*cm, height=410))
    elements.append(PageBreak())
    return elements
            
        
def generate_page2(df_transformed, block, village, fid, farmer, contact):

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

    elements = []

    # Add 3 centered heading lines
    elements.append(Image("images/title_page2.PNG", width=20*cm, height=50))
    elements.append(Spacer(1, 12))  # Add space before table

    # 1. Two-column layout for Name and Village
    info_table = Table([
        [Paragraph(f"Name = {farmer}", info_style), Paragraph(f"Village = {village}", info_style)]
    ], colWidths=[9*cm, 9*cm])

    info_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("ALIGN", (0, 0), (0, 0), "LEFT"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        # no grid or background, just layout
    ]))

    elements.append(info_table)

    # 2. Add Farmer ID (left-aligned)
    elements.append(Paragraph(f"Farmer ID = {fid}", info_style))

    df_transformed["प्रभावित क्षेत्र (हेक्टेयर में)"] = df_transformed["प्रभावित क्षेत्र (हेक्टेयर में)"].replace("", 0).astype(float)
    total_area = df_transformed["प्रभावित क्षेत्र (हेक्टेयर में)"].sum()
    total_pages = ceil(len(data_rows) / rows_per_page)
    for page_index in range(total_pages):
        start = page_index * rows_per_page
        end = start + rows_per_page
        page_rows = data_rows[start:end]

        table_data = [header_row] + page_rows
        table = Table(table_data, colWidths=[
            1.5*cm, 2.0*cm, 4.3*cm, 2.1*cm, 2.4*cm, 2.6*cm, 2.5*cm, 2.5*cm
        ])
        table.setStyle(style)
        elements.append(table)

        # Only append total row after last page
        if page_index == total_pages - 1:
            total_row_data = [[""] * 8]
            total_row_data[0][4] = "Total Area ="
            total_row_data[0][5] = f"{total_area:.5f} Hect"

            total_row = Table(total_row_data, colWidths=[
                1.5*cm, 2.0*cm, 4.2*cm, 2.2*cm, 2.4*cm, 2.6*cm, 2.5*cm, 2.5*cm
            ])
            total_row.setStyle(TableStyle([
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("ALIGN", (4, 0), (4, 0), "RIGHT"),
                ("ALIGN", (5, 0), (5, 0), "CENTER"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))

            elements.append(total_row)
            
        # Add vertical space before signature area
        elements.append(Spacer(1, 30))

        contact_style = ParagraphStyle(
            name="Contact",
            fontName="Helvetica",
            fontSize=10,
            alignment=0,
            spaceBefore=4
        )

        signature_row = Table([
            [
                Image("images/sign_1.PNG", width=7*cm, height=30),
                Image("images/sign_2.PNG", width=7*cm, height=30),
                Image("images/sign_3.PNG", width=7*cm, height=30)
            ],
            [
                Paragraph(f"{contact}", contact_style),
                "",
                ""
            ]
        ])

        elements.append(signature_row)
        
        # Page break unless it's the last page
        if page_index < total_pages - 1:
            elements.append(PageBreak())

    return elements

def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact, df_final_filtered):
    return _shared_generate_pdf(
        df_transformed, save_path, block, village, fid, farmer, contact,
        df_final_filtered, generate_page1, generate_page2
    )
