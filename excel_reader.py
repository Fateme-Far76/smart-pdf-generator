import sys
import pandas as pd
import os
import tempfile
import platform
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, PageBreak, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.styles import ParagraphStyle
from math import ceil
from reportlab.lib import colors
import subprocess
from pathlib import Path

# Register the bold Hindi font
pdfmetrics.registerFont(TTFont("HindiFont-Bold", "fonts/NotoSansDevanagari-Bold.ttf"))

# Register Hindi font
pdfmetrics.registerFont(TTFont("HindiFont", "fonts/NotoSansDevanagari-VariableFont_wdth,wght.ttf"))


from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QComboBox, QMessageBox, QListView, QLabel, QSizePolicy
)
from PyQt5.QtCore import Qt

class HoverComboBox(QComboBox):
    def __init__(self):
        super().__init__()
        view = QListView()
        view.setStyleSheet("""
            QListView::item:hover {
                background-color: #cce5ff;
                font-weight: bold;
            }
        """)
        self.setView(view)
        self.setMaximumWidth(200)

class ExcelFilterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to PDF Tool")
        self.setGeometry(400, 300, 320, 240)

        self.df_full = None
        self.df_block_filtered = None
        self.df_final_filtered = None

        self.init_ui()

    def init_ui(self):
        # Outer layout to center everything in the window
        outer_layout = QVBoxLayout(self)
        outer_layout.setAlignment(Qt.AlignCenter)

        # Inner layout for content
        main_layout = QVBoxLayout()
        main_layout.setSpacing(15)
        main_layout.setAlignment(Qt.AlignCenter)

        # After adding all widgets to main_layout:
        outer_layout.addLayout(main_layout)
        self.setLayout(outer_layout)

        # Title label
        title_label = QLabel("Select file:")
        title_label.setStyleSheet("font-weight: bold; font-size: 12pt;")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)

        # Browse button (centered)
        self.load_button = QPushButton("Browse")
        self.load_button.setFixedSize(120, 36)
        self.load_button.setStyleSheet("font-size: 10pt;")
        self.load_button.clicked.connect(self.load_excel)
        main_layout.addWidget(self.load_button, alignment=Qt.AlignCenter)

        # ---- Helper to add label + dropdown centered ----
        def add_centered_dropdown(label_text, dropdown_widget):
            container = QVBoxLayout()
            label = QLabel(label_text)
            label.setAlignment(Qt.AlignCenter)
            label.setStyleSheet("font-size: 11pt;")
            dropdown_widget.setFixedSize(220, 32)  # Wider & taller
            dropdown_widget.setStyleSheet("font-size: 10pt;")
            container.addWidget(label, alignment=Qt.AlignCenter)
            container.addWidget(dropdown_widget, alignment=Qt.AlignCenter)
            main_layout.addLayout(container)

        # Block dropdown
        self.block_dropdown = HoverComboBox()
        self.block_dropdown.currentIndexChanged.connect(self.populate_villages)
        add_centered_dropdown("Select Block:", self.block_dropdown)

        # Village dropdown
        self.village_dropdown = HoverComboBox()
        self.village_dropdown.currentIndexChanged.connect(self.populate_fids)
        add_centered_dropdown("Select Village:", self.village_dropdown)

        # Farmer ID dropdown
        self.fid_dropdown = HoverComboBox()
        self.fid_dropdown.currentIndexChanged.connect(self.populate_farmers)
        add_centered_dropdown("Select Farmer ID:", self.fid_dropdown)

        # Farmer Name dropdown
        self.farmer_dropdown = HoverComboBox()
        add_centered_dropdown("Select Farmer Name:", self.farmer_dropdown)

        # Save + Print button row
        button_row = QHBoxLayout()
        self.pdf_button = QPushButton("Save as PDF")
        self.pdf_button.setFixedSize(130, 36)
        self.pdf_button.setStyleSheet("font-size: 10pt;")
        self.pdf_button.clicked.connect(self.save_pdf)
        button_row.addWidget(self.pdf_button)

        self.print_button = QPushButton("Print")
        self.print_button.setFixedSize(100, 36)
        self.print_button.setStyleSheet("font-size: 10pt;")
        self.print_button.clicked.connect(self.print_pdf)
        button_row.addWidget(self.print_button)

        main_layout.addLayout(button_row)


    def load_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Open Excel File", "", "Excel Files (*.xlsx *.xls)"
        )

        if not file_path:
            return

        try:
            all_sheets = pd.read_excel(file_path, sheet_name=None)
            sheet_name = list(all_sheets.keys())[0]
            df = all_sheets[sheet_name]

            if not {"Block", "Village"}.issubset(df.columns):
                QMessageBox.critical(self, "Missing Columns", "Excel file must contain 'Block' and 'Village' columns.")
                return

            self.df_full = df
            self.df_block_filtered = None
            self.df_final_filtered = None

            blocks = sorted(df["Block"].dropna().unique().astype(str))
            self.block_dropdown.clear()
            self.block_dropdown.addItem("Select Block")
            self.block_dropdown.addItems(blocks)

            self.village_dropdown.clear()
            self.village_dropdown.addItem("Select Village")


        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error reading Excel:\n{str(e)}")

    def populate_villages(self):
        block = self.block_dropdown.currentText()
        if block == "Select Block" or self.df_full is None:
            return

        self.df_block_filtered = self.df_full[self.df_full["Block"].astype(str) == block]

        villages = sorted(self.df_block_filtered["Village"].dropna().unique().astype(str))
        self.village_dropdown.clear()
        self.village_dropdown.addItem("Select Village")
        self.village_dropdown.addItems(villages)

        self.fid_dropdown.clear()
        self.fid_dropdown.addItem("Select Farmer ID")

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")

    def filter_data(self):
        block = self.block_dropdown.currentText()
        village = self.village_dropdown.currentText()
        fid = self.fid_dropdown.currentText()
        farmer = self.farmer_dropdown.currentText()

        if (
            block == "Select Block" or
            village == "Select Village" or
            fid == "Select Farmer ID" or
            farmer == "Select Farmer Name"
        ):
            QMessageBox.warning(self, "Missing Selection", "Please select all filters.")
            return

        self.df_final_filtered = self.df_full[
            (self.df_full["Block"].astype(str) == block) &
            (self.df_full["Village"].astype(str) == village) &
            (self.df_full["FID"].astype(str) == fid) &
            (self.df_full["Farmer Name"].astype(str) == farmer)
        ]
        transformed = self.transform_filtered_data()
        print(transformed)

    def populate_farmers(self):
        fid = self.fid_dropdown.currentText()
        if fid == "Select Farmer ID" or self.df_village_filtered is None:
            return

        df_fid_filtered = self.df_village_filtered[self.df_village_filtered["FID"].astype(str) == fid]
        farmers = sorted(df_fid_filtered["Farmer Name"].dropna().unique().astype(str))

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")
        self.farmer_dropdown.addItems(farmers)

    def populate_fids(self):
        village = self.village_dropdown.currentText()
        if village == "Select Village" or self.df_block_filtered is None:
            return

        df_village_filtered = self.df_block_filtered[self.df_block_filtered["Village"].astype(str) == village]
        self.df_village_filtered = df_village_filtered  # Save for next step

        fids = sorted(df_village_filtered["FID"].dropna().unique().astype(str))

        self.fid_dropdown.clear()
        self.fid_dropdown.addItem("Select Farmer ID")
        self.fid_dropdown.addItems(fids)

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")

    def transform_filtered_data(self):
        if self.df_final_filtered is None or self.df_final_filtered.empty:
            QMessageBox.warning(self, "No Data", "No filtered data to transform.")
            return None

        df = self.df_final_filtered.copy()

        # Build the output DataFrame
        transformed = pd.DataFrame({
            "क्र.सं.": range(1, len(df) + 1),
            "सीएलएस नंबर": df["Saksham ID"].astype(str).apply(lambda x: x.split("-")[-1] if "-" in x else x),
            "आवेदन संख्या": df["Intimation_Application id"],
            "फसल का नाम": df["Crop"],
            "सर्वे/उप सर्वे नंबर": df["Survey Number"],
            "प्रभावित क्षेत्र (हेक्टेयर में)": df["Crop Area"],
            "हानि प्रतिशत": "",  # Blank
            "रिमार्क": ""         # Blank
        })

        return transformed

    def transform_filtered_data_direct(self, df_input):
        df = df_input.copy()

        transformed = pd.DataFrame({
            "क्र.सं.": range(1, len(df) + 1),
            "सीएलएस नंबर": df["Saksham ID"].astype(str).apply(lambda x: x.split("-")[-1] if "-" in x else x),
            "आवेदन संख्या": df["Intimation_Application id"],
            "फसल का नाम": df["Crop"],
            "सर्वे/उप सर्वे नंबर": df["Survey Number"],
            "प्रभावित क्षेत्र (हेक्टेयर में)": df["Crop Area"],
            "हानि प्रतिशत": "",
            "रिमार्क": ""
        })

        return transformed

    def save_pdf(self):
        block = self.block_dropdown.currentText()
        village = self.village_dropdown.currentText()
        fid = self.fid_dropdown.currentText()
        farmer = self.farmer_dropdown.currentText()

        if block == "Select Block":
            QMessageBox.warning(self, "Missing Selection", "Please select at least Block.")
            return


        desktop = Path.home() / "Desktop"

        # New Option 0: Block-only mode — single combined PDF for whole block
        if village == "Select Village":
            df_block = self.df_full[self.df_full["Block"].astype(str) == block]

            if df_block.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this block.")
                return

            df_transformed = self.transform_filtered_data_direct(df_block)

            desktop = Path.home() / "Desktop"
            save_path = desktop / f"{block}.pdf"

            try:
                self.generate_pdf(df_transformed, save_path, block, "", "", "", "")  # empty village/fid/farmer/contact
                QMessageBox.information(self, "Saved", f"PDF saved to:\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save PDF:\n{e}")
            return

        # Option 1: All four selected — Single farmer mode
        if fid != "Select Farmer ID" and farmer != "Select Farmer Name":
            self.df_final_filtered = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village) &
                (self.df_full["FID"].astype(str) == fid) &
                (self.df_full["Farmer Name"].astype(str) == farmer)
            ]
            contact = self.df_final_filtered["Contact Number"].dropna().astype(str).iloc[0]  # safe
            if self.df_final_filtered.empty:
                QMessageBox.warning(self, "No Data", "No data found for selected farmer.")
                return

            df_transformed = self.transform_filtered_data()
            save_path = desktop / f"{block}_{village}_{fid}.pdf"

            try:
                self.generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact)
                QMessageBox.information(self, "Saved", f"PDF saved to:\n{save_path}")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to save PDF:\n{e}")
            return

        # Option 2: Only Block + Village — Batch mode
        elif fid == "Select Farmer ID" and farmer == "Select Farmer Name":
            df_village = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village)
            ]
            farmers = df_village[["FID", "Farmer Name"]].dropna().drop_duplicates()

            if farmers.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this village.")
                return

            errors = []

            for _, row in farmers.iterrows():
                current_fid = str(row["FID"])
                current_name = str(row["Farmer Name"])
                filtered = df_village[
                    (df_village["FID"].astype(str) == current_fid) &
                    (df_village["Farmer Name"].astype(str) == current_name)
                ]
                contact = filtered["Contact Number"].dropna().astype(str).iloc[0]
                if filtered.empty:
                    continue

                df_transformed = self.transform_filtered_data_direct(filtered)
                save_path = desktop / f"{block}_{village}_{current_fid}.pdf"

                try:
                    self.generate_pdf(df_transformed, save_path, block, village, current_fid, current_name, contact)
                except Exception as e:
                    errors.append(f"{current_name}: {e}")

            QMessageBox.information(self, "Done", f"Saved {len(farmers)} PDFs to Desktop.\n" +
                                    ("Some errors occurred." if errors else ""))
            return

        # Invalid combination
        else:
            QMessageBox.warning(self, "Incomplete Selection", "Please either select:\n• Block + Village only\nOR\n• Block + Village + Farmer ID + Name.")

    def generate_page1(self, block, village, fid, farmer, contact):
        elements = []

        return elements
                
            
    def generate_page2(self, df_transformed, block, village, fid, farmer, contact):

        header_style = ParagraphStyle(
            name="Header",
            fontName="HindiFont-Bold",
            fontSize=10,
            alignment=1
        )
        
        info_style = ParagraphStyle(
            name="FarmerInfo",
            fontName="Helvetica-Bold",
            fontSize=8,
            alignment=0,  # Left-aligned
            spaceAfter=4
        )
        
        style = TableStyle([
            ("FONTNAME", (0, 0), (-1, -1), "HindiFont"),
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
            Paragraph("क्र.सं.", header_style),
            Paragraph("सीएलएस<br/>नंबर", header_style),
            Paragraph("आवेदन संख्या", header_style),
            Paragraph("फसल का<br/>नाम", header_style),
            Paragraph("सर्वे/उप सर्वे<br/>नंबर", header_style),
            Paragraph("प्रभावित क्षेत्र<br/>(हेक्टेयर में)", header_style),
            Paragraph("हानि प्रतिशत", header_style),
            Paragraph("रिमार्क", header_style)
        ]

        elements = []

        # Hindi heading style
        heading_style = ParagraphStyle(
            name="Heading",
            fontName="HindiFont-Bold",
            fontSize=12,
            alignment=1,  # Center
            spaceAfter=6
        )

        # Add 3 centered heading lines
        elements.append(Paragraph("प्रधान मंत्री फसल बीमा योजना", heading_style))
        elements.append(Paragraph("प्रधान मंत्री फसल बीमा योजना के तहत फसल मे हुई हानी के आँकलन की रिपोर्ट", heading_style))
        elements.append(Paragraph("स्थानीय आपदाओं / फसल कटाई उपरान्त बीमा दावो के लिए", heading_style))
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

        total_area = df_transformed["प्रभावित क्षेत्र (हेक्टेयर में)"].astype(float).sum()
        total_pages = ceil(len(data_rows) / rows_per_page)
        for page_index in range(total_pages):
            start = page_index * rows_per_page
            end = start + rows_per_page
            page_rows = data_rows[start:end]

            table_data = [header_row] + page_rows
            table = Table(table_data, colWidths=[
                1.5*cm, 2.0*cm, 4.2*cm, 2.2*cm, 2.4*cm, 2.6*cm, 2.5*cm, 2.5*cm
            ])
            table.setStyle(style)
            elements.append(table)

            # ✅ Only append total row after last page
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

            # Signature labels
            signature_style = ParagraphStyle(
                name="Signature",
                fontName="HindiFont",
                fontSize=10,
                alignment=0
            )

            contact_style = ParagraphStyle(
                name="Contact",
                fontName="Helvetica",
                fontSize=10,
                alignment=0,
                spaceBefore=4
            )

            signature_row = Table([
                [
                    Paragraph("कृषक के हस्त हस्ताक्षर", signature_style),
                    Paragraph("खंड कृषि अधिकारी के हस्ताक्षर", signature_style),
                    Paragraph("बीमा कंपनी के अधिकारी का हस्ताक्षर", signature_style)
                ],
                [
                    Paragraph(f"{contact}", contact_style),
                    "",
                    ""
                ]
            ], colWidths=[6*cm, 6*cm, 6*cm])

            signature_row.setStyle(TableStyle([
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("TOPPADDING", (0, 0), (-1, -1), 10),
            ]))

            elements.append(signature_row)
            
            # Page break unless it's the last page
            if page_index < total_pages - 1:
                elements.append(PageBreak())

        return elements
        
    def generate_pdf(self, df_transformed, save_path, block, village, fid, farmer, contact):
        elements = []

        # Page 1
        elements += self.generate_page1(block, village, fid, farmer, contact)

        # Page 2
        elements += self.generate_page2(df_transformed, block, village, fid, farmer, contact)

        # Build the PDF
        doc = SimpleDocTemplate(str(save_path), pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=2*cm)
        doc.build(elements)

    def print_pdf(self):
        block = self.block_dropdown.currentText()
        village = self.village_dropdown.currentText()
        fid = self.fid_dropdown.currentText()
        farmer = self.farmer_dropdown.currentText()

        if block == "Select Block":
            QMessageBox.warning(self, "Missing Selection", "Please select at least Block.")
            return


        system = platform.system()
        if system != "Windows":
            QMessageBox.warning(self, "Unsupported", "Printing is only supported on Windows.")
            return

        temp_dir = tempfile.gettempdir()

        # ✅ New Option 0: Block-only mode
        if village == "Select Village":
            df_block = self.df_full[self.df_full["Block"].astype(str) == block]

            if df_block.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this block.")
                return

            df_transformed = self.transform_filtered_data_direct(df_block)

            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"{block}_print.pdf")

            try:
                self.generate_pdf(df_transformed, temp_path, block, "", "", "", "")
                os.startfile(temp_path, "print")
                QMessageBox.information(self, "Printed", "PDF sent to printer.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to print PDF:\n{e}")
            return

        # Option 1: Single farmer
        if fid != "Select Farmer ID" and farmer != "Select Farmer Name":
            self.df_final_filtered = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village) &
                (self.df_full["FID"].astype(str) == fid) &
                (self.df_full["Farmer Name"].astype(str) == farmer)
            ]
            contact = self.df_final_filtered["Contact Number"].dropna().astype(str).iloc[0]
            if self.df_final_filtered.empty:
                QMessageBox.warning(self, "No Data", "No data found for selected farmer.")
                return

            df_transformed = self.transform_filtered_data()
            temp_path = os.path.join(temp_dir, f"{block}_{village}_{fid}_print.pdf")

            try:
                self.generate_pdf(df_transformed, temp_path, block, village, fid, farmer, contact)
                os.startfile(temp_path, "print")
                QMessageBox.information(self, "Printed", "PDF sent to printer.")
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Failed to print PDF:\n{e}")

        # Option 2: All farmers in village
        elif fid == "Select Farmer ID" and farmer == "Select Farmer Name":
            df_village = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village)
            ]

            farmers = df_village[["FID", "Farmer Name"]].dropna().drop_duplicates()

            if farmers.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this village.")
                return

            errors = []

            for _, row in farmers.iterrows():
                current_fid = str(row["FID"])
                current_name = str(row["Farmer Name"])
                filtered = df_village[
                    (df_village["FID"].astype(str) == current_fid) &
                    (df_village["Farmer Name"].astype(str) == current_name)
                ]
                contact = filtered["Contact Number"].dropna().astype(str).iloc[0]  # safe
                if filtered.empty:
                    continue

                df_transformed = self.transform_filtered_data_direct(filtered)
                temp_path = os.path.join(temp_dir, f"{block}_{village}_{current_fid}_print.pdf")

                try:
                    self.generate_pdf(df_transformed, temp_path, block, village, current_fid, current_name, contact)
                    os.startfile(temp_path, "print")
                except Exception as e:
                    errors.append(f"{current_name}: {e}")

            QMessageBox.information(self, "Printed", f"Sent {len(farmers)} files to printer.\n" +
                                    ("Some errors occurred." if errors else ""))

        # Invalid selection
        else:
            QMessageBox.warning(self, "Incomplete Selection",
                                "Please either select:\n• Block + Village only\nOR\n• Block + Village + Farmer ID + Name.")


def main():
    app = QApplication(sys.argv)
    window = ExcelFilterApp()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()
