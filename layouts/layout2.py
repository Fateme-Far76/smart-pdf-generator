import sys
import os
import pandas as pd
from PIL import Image as PILimage
from PIL import ImageDraw, ImageFont
import io
from reportlab.lib.pagesizes import A4
from reportlab.platypus import Table, TableStyle, Paragraph, PageBreak, Spacer
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

def save_image_with_texts(
    image_path, 
    text1, 
    text2, 
    font_path,
    out_w,
    out_h, 
    font_size=28,   # smaller font
    margin=25, 
    offset_down=30, # push both lines down more
    spacing=15      # more vertical space between lines
):
    # Open image
    img = PILimage.open(image_path).convert("RGBA")
    width, height = img.size

    # Draw on it
    draw = ImageDraw.Draw(img)
    font = ImageFont.truetype(font_path, font_size)

    # Measure first text
    bbox1 = draw.textbbox((0, 0), text1, font=font)
    text1_width = bbox1[2] - bbox1[0]
    text1_height = bbox1[3] - bbox1[1]

    # Measure second text
    bbox2 = draw.textbbox((0, 0), text2, font=font)
    text2_width = bbox2[2] - bbox2[0]
    text2_height = bbox2[3] - bbox2[1]

    # Position first text
    y1 = ((height - (text1_height + spacing + text2_height)) // 2) + offset_down
    x1 = width - text1_width - margin

    # Position second text
    y2 = y1 + text1_height + spacing
    x2 = width - text2_width - margin

    # Draw both in black
    draw.text((x1, y1), text1, font=font, fill=(0, 0, 0, 255))
    draw.text((x2, y2), text2, font=font, fill=(0, 0, 0, 255))

    buf = io.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return Image(buf, width=out_w, height=out_h)


def generate_page2(leftover_df, header_row, data_tbl_width, tbl_style, table_sign, total_value, fid, englishnormal):
    
    elements = []
    elements.append(Paragraph(f"<b>{fid}</b>", englishnormal))
    elements.append(Spacer(1, 3)) 
    
    # --- constants ---
    DATA_PER_PAGE = 40
    START_NUM = 11                 # numbering starts at 11 on every page

    # Make a display-safe copy (no NaN strings)
    df = leftover_df.reset_index(drop=True).copy()

    total_pages = ceil(len(df) / DATA_PER_PAGE) or 1

    for page_index in range(total_pages):
        start = page_index * DATA_PER_PAGE
        end   = start + DATA_PER_PAGE
        chunk = df.iloc[start:end].copy()

        # Build exactly 50 visible rows (numbers 11..50), first fill with data, then blanks
        numbered_rows = []
        for i in range(DATA_PER_PAGE):
            row_num = START_NUM + i
            if i < len(chunk):
                row_vals = chunk.iloc[i].tolist()
                # ensure correct width and set serial number in col 0
                if len(row_vals) >= 1:
                    row_vals[0] = str(row_num)
                else:
                    row_vals = [str(row_num)] + [""] * 7
            else:
                row_vals = [str(row_num)] + [""] * 7
            numbered_rows.append(row_vals)
        START_NUM += 40
        # Assemble table: header + numbered rows
        table_data = [header_row] + numbered_rows

        # Build a row with the same number of columns as the main table (+1 for the serial column)
        total_row = [""] * 8

        total_label = Image(resource_path("images/l2_sum.PNG"), width=3*cm, height=15)
        total_row[0] = total_label
        total_row[3] = total_value
        table_data.append(total_row)
        
        tbl = Table(table_data, colWidths=data_tbl_width)
        
        tbl.setStyle(tbl_style)
        elements.append(tbl)
        elements.append(Spacer(1, 10))   
        elements.append(table_sign)
        elements.append(Image(resource_path("images/l2_bottom.PNG"), width=20*cm, height=40))
        # Page break between pages (not after the last page)
        if page_index < total_pages - 1:
            elements.append(PageBreak())

    return elements

def generate_page1(df_transformed, block, village, fid, farmer, contact, df_final_filtered):
    elements = []
    
    englishnormal = ParagraphStyle(
        name="eng",
        fontName="Helvetica",
        fontSize=9,
        alignment=0,
    )
    elements.append(Paragraph(f"<b>{fid}</b>", englishnormal))
    elements.append(Image(resource_path("images/l2_title.PNG"), width=20*cm, height=45))
    elements.append(Spacer(1, 10))


    # fetch values from the DataFrame
    survey_number = df_final_filtered["Survey Number"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    event_date = df_final_filtered["Event occurred Date"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    intimation_date = df_final_filtered["Date of Intimation"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    crop = df_final_filtered["Crop"].astype(str).iloc[0] if not df_final_filtered.empty else ""
    Intimation_Application_id = df_final_filtered["Intimation_Application id"].astype(str).iloc[0][:19] if not df_final_filtered.empty else ""
    
    data = [
        ["1.", Image(resource_path("images/l2_1.PNG"), width=100, height=13), Paragraph(farmer, englishnormal), 
        "2.", Image(resource_path("images/l2_2.PNG"), width=100, height=13), ""],

        ["3.", Image(resource_path("images/l2_3.PNG"), width=100, height=17), Paragraph(village, englishnormal), 
        "4.", Image(resource_path("images/l2_4.PNG"), width=100, height=17), Paragraph(block, englishnormal)],

        ["5.", Image(resource_path("images/l2_5.PNG"), width=100, height=15), Paragraph("Hisar", englishnormal), 
        "6.", Image(resource_path("images/l2_6.PNG"), width=100, height=15), Paragraph("HDFC Ergo", englishnormal)],

        ["7.", Image(resource_path("images/l2_7.PNG"), width=100, height=22), Paragraph(crop, englishnormal), 
        "8.", Image(resource_path("images/l2_8.PNG"), width=100, height=23), Paragraph(survey_number, englishnormal)],

        ["9.", Image(resource_path("images/l2_9.PNG"), width=100, height=18), "", 
        "10.", Image(resource_path("images/l2_10.PNG"), width=100, height=15), ""],

        ["11.", Image(resource_path("images/l2_11.PNG"), width=100, height=18), "", 
        "12.", Image(resource_path("images/l2_12.PNG"), width=100, height=18), Paragraph(event_date, englishnormal)],

        ["13.", Image(resource_path("images/l2_13.PNG"), width=100, height=25), Paragraph(intimation_date, englishnormal), 
        "14.", Image(resource_path("images/l2_14.PNG"), width=100, height=22), ""],

        ["15.", Image(resource_path("images/l2_15.PNG"), width=100, height=14), Paragraph(Intimation_Application_id, englishnormal)]
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
    elements.append(Spacer(1, 6))
    elements.append(Image(resource_path("images/l2_middle.PNG"), width=20*cm, height=140))

    ROWS_TARGET = 10
    df = df_transformed.copy()
    
    # Compute the sum for crop area
    area_series= (
        pd.to_numeric(df["बीमित क्षेत्र"].replace("", 0), errors="coerce")
        .fillna(0.0)
    )
    sum_area = float(area_series.sum())
    total_area = f"{sum_area:.5f}"
    page_df = df.reset_index(drop=True).copy()
    page_df = page_df.map(lambda x: "" if pd.isna(x) else str(x))

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
        Image(resource_path("images/l2_table1.PNG"), width=0.6*cm, height=35),
        Image(resource_path("images/l2_table2.PNG"), width=4.2*cm, height=42),
        Image(resource_path("images/l2_table3.PNG"), width=2.7*cm, height=41),
        Image(resource_path("images/l2_table4.PNG"), width=2.4*cm, height=39),
        Image(resource_path("images/l2_table5.PNG"), width=2.2*cm, height=39),
        Image(resource_path("images/l2_table6.PNG"), width=1.1*cm, height=35),
        Image(resource_path("images/l2_table7.PNG"), width=1.4*cm, height=39),
        Image(resource_path("images/l2_table8.PNG"), width=1.7*cm, height=38)
    ]

    table_rows = []
    for i, row_vals in enumerate(page_display.values.tolist(), start=1):
        row_vals[0] = str(i)  # force first col to 1..10 shown
        table_rows.append(row_vals)

    data = [header_row] + table_rows
    total_label = Image(resource_path("images/l2_sum.PNG"), width=3*cm, height=15)
    left_style = ParagraphStyle("Left", fontName="Helvetica", fontSize=9, alignment=0)
    total_value = Paragraph(f"<b>{total_area}</b>", left_style)
    
    total_row = [""] * 8
    total_row[0] = total_label
    total_row[3] = total_value
    data.append(total_row)   
    data_tbl_width = [0.9*cm, 4.72*cm, 2.92*cm, 2.42*cm, 2.02*cm, 2*cm, 2*cm, 2.9*cm]
    table = Table(data, colWidths=data_tbl_width)
    
    tbl_style = TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("LEADING",  (0, 1), (-1, -1), 9),
        ("GRID", (0, 0), (-1, -2), 0.5, colors.black),  # all grid except total row bottom
        ("GRID", (0, -1), (-1, -1), 0.5, colors.black), # grid for total row too
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),

        # TOTAL ROW MERGES:
        # Merge first four cells of the total row (0..3) to make whitespace block on left
        ("SPAN", (0, -1), (2, -1)),
        ("SPAN", (3, -1), (7, -1)),
        ("SPAN", (7, 1), (7, -2)),
        # --- TOTAL ROW ALIGNMENTS ---
        ("ALIGN", (0, -1), (0, -1), "RIGHT"),  # label right
        ("ALIGN", (3, -1), (3, -1), "LEFT"),   # value left

        # --- TOTAL ROW HEIGHT (smaller) ---
        ("TOPPADDING", (0, -1), (-1, -1), 0),
        ("BOTTOMPADDING", (0, -1), (-1, -1), 0),
    ])
    
    table.setStyle(tbl_style)
    elements.append(table)
    elements.append(Spacer(1, 2))   
    
    blow_tbl_img = Image(resource_path("images/l2_p1_below_table.PNG"), width=6*cm, height=0.8*cm)
    blow_tbl_img_tbl = Table([[blow_tbl_img]], colWidths=[19.88*cm])  # colWidths same as the image width
    blow_tbl_img_tbl.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    elements.append(blow_tbl_img_tbl)
    
    elements.append(Spacer(1, 50))    
    sign1_img = save_image_with_texts(resource_path("images/l2_sign1.PNG"), farmer, contact, font_path=resource_path("fonts/Helvetica.ttf"), out_w=4.9*cm, out_h=40)
    
    # Example data (replace with your content)
    data_sign = [[
        sign1_img,
        Image(resource_path("images/l2_sign2.PNG"), width=4.9*cm, height=40),
        Image(resource_path("images/l2_sign3.PNG"), width=4.9*cm, height=40),
        Image(resource_path("images/l2_sign4.PNG"), width=4.9*cm, height=40)  
    ]]

    # Create table
    table_sign = Table(data_sign, colWidths=[4.97*cm, 4.97*cm, 4.97*cm, 4.97*cm], rowHeights=40)
    elements.append(table_sign)
    elements.append(Image(resource_path("images/l2_bottom.PNG"), width=20*cm, height=40))

    if len(leftover_df) > 0:
        elements.append(PageBreak())
        elements += generate_page2(leftover_df, header_row, data_tbl_width, tbl_style, table_sign, total_value, fid, englishnormal)
    return elements
                             
            
def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact, df_final_filtered):
    return generate_pdf2(
        df_transformed, save_path, block, village, fid, farmer, contact,
        df_final_filtered, generate_page1
    )