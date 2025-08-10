import os
import sys
import pandas as pd
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import ParagraphStyle
from math import ceil
from reportlab.lib import colors
from common.pdf_builder import generate_pdf as _shared_generate_pdf
from reportlab.platypus import Image, Paragraph, Table, TableStyle, PageBreak, Spacer

# Helper to get the path of the bundled font folder
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS  # temp folder used by PyInstaller
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def generate_page1(df_final_filtered, block, village, fid, farmer, contact):
    elements = []
    
    englishnormal = ParagraphStyle(
        name="eng",
        fontName="Helvetica",
        fontSize=10,
        alignment=0,
    )
    
    elements.append(Image(resource_path("images/title.PNG"), width=20*cm, height=30))
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
        ["1.", Image(resource_path("images/1.PNG"), width=100, height=15), Paragraph(farmer, englishnormal), 
        "2.", Image(resource_path("images/2.PNG"), width=100, height=15), ""],

        ["3.", Image(resource_path("images/3.PNG"), width=100, height=15), Paragraph(village, englishnormal), 
        "4.", Image(resource_path("images/4.PNG"), width=100, height=15), Paragraph(block, englishnormal)],

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

        ["15.", Image(resource_path("images/15.PNG"), width=100, height=40), mini_table],

        ["16.", Image(resource_path("images/16.PNG"), width=100, height=15), ""]
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
    elements.append(Image(resource_path("images/rest_page1.PNG"), width=20*cm, height=410))
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
        ("FONTSIZE", (0, 0), (-1, -1), 10),          # default size for all cells
        ("FONTSIZE", (3, 0), (3, -1), 9)
    ])


    # Remove existing index column (we'll insert it manually)
    df_transformed = df_transformed.drop(columns=["क्र.सं."])

    # Header row
    header_row = [
        Image(resource_path("images/col1.PNG"), width=1.4*cm, height=20),
        Image(resource_path("images/col2.PNG"), width=1.8*cm, height=20),
        Image(resource_path("images/col3.PNG"), width=4.3*cm, height=20),
        Image(resource_path("images/col4.PNG"), width=3*cm, height=20),
        Image(resource_path("images/col5.PNG"), width=2.6*cm, height=20),
        Image(resource_path("images/col6.PNG"), width=2.8*cm, height=20),
        Image(resource_path("images/col7.PNG"), width=2.4*cm, height=20),
        Image(resource_path("images/col8.PNG"), width=2.4*cm, height=20)
    ]


    elements = []

    # Add 3 centered heading lines
    elements.append(Image(resource_path("images/title_page2.PNG"), width=20*cm, height=50))
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

    # --- Group by crop, then paginate 22 rows per page ---
    rows_per_page = 22

    # Build once (used below)
    contact_style = ParagraphStyle(
        name="Contact",
        fontName="Helvetica",
        fontSize=10,
        alignment=0,
        spaceBefore=4
    )

    # Get groups in a stable order so we can know the last crop
    crop_groups = list(df_transformed.groupby("फसल का नाम", dropna=False, sort=False))

    for crop_idx, (crop_value, grp) in enumerate(crop_groups):
        is_last_crop = (crop_idx == len(crop_groups) - 1)

        # Display copy (strings; NaN -> "")
        grp_display = (
            grp.reset_index(drop=True)
            .where(grp.notna(), "")
            .astype(str)
        )

        # Per-crop total (area column), robust to blanks
        total_area = None
        if "प्रभावित क्षेत्र (हेक्टेयर में)" in grp.columns:
            area_series = pd.to_numeric(
                grp["प्रभावित क्षेत्र (हेक्टेयर में)"].replace("", 0),
                errors="coerce"
            ).fillna(0.0)
            total_area = float(area_series.sum())

        total_pages = ceil(len(grp_display) / rows_per_page) or 1

        for page_index in range(total_pages):
            start = page_index * rows_per_page
            end = start + rows_per_page
            page_df = grp_display.iloc[start:end].copy()

            # Pad this page to exactly 22 rows for layout consistency
            if len(page_df) < rows_per_page:
                blanks_needed = rows_per_page - len(page_df)
                page_df = pd.concat(
                    [page_df, pd.DataFrame([[""] * page_df.shape[1]] * blanks_needed, columns=page_df.columns)],
                    ignore_index=True
                )

            # Page numbering 1..22 (per page)
            numbered_rows = []
            for i, row in enumerate(page_df.values.tolist(), start=1):
                numbered_rows.append([str(i)] + row)

            # Build table (reuse your header_row/style/colWidths)
            table_data = [header_row] + numbered_rows
            table = Table(table_data, colWidths=[
                1.5*cm, 2.0*cm, 4.3*cm, 2.9*cm, 2.4*cm, 2.6*cm, 2.1*cm, 2.1*cm
            ])
            table.setStyle(style)
            elements.append(table)

            # Show total row only on the LAST page of this crop
            if (total_area is not None) and (page_index == total_pages - 1):
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

            # Spacing + signatures (same as your current code)
            elements.append(Spacer(1, 120))
            signature_row = Table([
                [
                    Image(resource_path("images/sign_1.PNG"), width=7*cm, height=30),
                    Image(resource_path("images/sign_2.PNG"), width=7*cm, height=30),
                    Image(resource_path("images/sign_3.PNG"), width=7*cm, height=30)
                ],
                [
                    Paragraph(f"{contact}", contact_style),
                    "",
                    ""
                ]
            ])
            elements.append(signature_row)

            # New page unless this is the very last page of the very last crop
            if not (is_last_crop and page_index == total_pages - 1):
                elements.append(PageBreak())

    return elements

def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact, df_final_filtered):
    return _shared_generate_pdf(
        df_transformed, save_path, block, village, fid, farmer, contact,
        df_final_filtered, generate_page1, generate_page2
    )
