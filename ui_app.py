import sys
from ui_widgets import HoverComboBox
from layouts import layout1, layout2
from pathlib import Path
import tempfile
import platform
import os
import pandas as pd
from common.data_transform import transform_filtered_data, transform_filtered_data2
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton, QFileDialog,
    QComboBox, QMessageBox, QListView, QLabel, QSizePolicy
)
from PyQt5.QtCore import Qt

class ExcelFilterApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel to PDF Tool")
        self.setGeometry(400, 300, 320, 240)

        self.df_full = None
        self.df_block_filtered = None
        self.df_final_filtered = None
        self.df_village_filtered = None

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

        # inside init_ui()
        self.layout_dropdown = HoverComboBox()
        self.layout_dropdown.addItems(["Layout1", "Layout2"])   # names shown to user
        add_centered_dropdown("Select Layout:", self.layout_dropdown)

        # keep a map from dropdown text -> module
        self.layouts = {
            "Layout1": layout1,
            "Layout2": layout2,
        }

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
            df = pd.read_excel(file_path, sheet_name=0).fillna("")

            if not {"Block", "Village"}.issubset(df.columns):
                QMessageBox.critical(self, "Missing Columns", "Excel file must contain 'Block' and 'Village' columns.")
                return

            self.df_full = df
            self.df_block_filtered = None
            self.df_final_filtered = None

            blocks = sorted(df["Block"].unique().astype(str))
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

        villages = sorted(self.df_block_filtered["Village"].unique().astype(str))
        self.village_dropdown.clear()
        self.village_dropdown.addItem("Select Village")
        self.village_dropdown.addItems(villages)

        self.fid_dropdown.clear()
        self.fid_dropdown.addItem("Select Farmer ID")

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")

    def populate_farmers(self):
        fid = self.fid_dropdown.currentText()
        if fid == "Select Farmer ID" or self.df_village_filtered is None:
            return

        df_fid_filtered = self.df_village_filtered[self.df_village_filtered["FID"].astype(str) == fid]
        farmers = sorted(df_fid_filtered["Farmer Name"].unique().astype(str))

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")
        self.farmer_dropdown.addItems(farmers)

    def populate_fids(self):
        village = self.village_dropdown.currentText()
        if village == "Select Village" or self.df_block_filtered is None:
            return

        df_village_filtered = self.df_block_filtered[self.df_block_filtered["Village"].astype(str) == village]
        self.df_village_filtered = df_village_filtered  # Save for next step

        fids = sorted(df_village_filtered["FID"].unique().astype(str))

        self.fid_dropdown.clear()
        self.fid_dropdown.addItem("Select Farmer ID")
        self.fid_dropdown.addItems(fids)

        self.farmer_dropdown.clear()
        self.farmer_dropdown.addItem("Select Farmer Name")

    def save_pdf(self):
        block = self.block_dropdown.currentText()
        village = self.village_dropdown.currentText()
        fid = self.fid_dropdown.currentText()
        farmer = self.farmer_dropdown.currentText()
        selected_layout_name = self.layout_dropdown.currentText()
        layout = self.layouts.get(selected_layout_name, layout1)
        
        if block == "Select Block" or village == "Select Village":
            QMessageBox.warning(self, "Missing Selection", "Please select at least Block and Village.")
            return

        if getattr(sys, 'frozen', False):
            # App is running from a PyInstaller .exe
            app_dir = Path(sys.executable).parent
        else:
            # App is running from a .py file
            app_dir = Path(__file__).resolve().parent

        # Option 1: All four selected — Single farmer mode
        if fid != "Select Farmer ID" and farmer != "Select Farmer Name":
            self.df_final_filtered = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village) &
                (self.df_full["FID"].astype(str) == fid) &
                (self.df_full["Farmer Name"].astype(str) == farmer)
            ]
            contact = self.df_final_filtered["Contact Number"].astype(str).iloc[0]  # safe
            if self.df_final_filtered.empty:
                QMessageBox.warning(self, "No Data", "No data found for selected farmer.")
                return

            if layout ==layout1:
                df_transformed = transform_filtered_data(self.df_final_filtered)
            else:
                df_transformed = transform_filtered_data2(self.df_final_filtered)
                
            save_path = app_dir / f"{block} {village} {fid}.pdf"

            try:
                layout.generate_pdf(
                    df_transformed, save_path, block, village, fid, farmer, contact, self.df_final_filtered
                )
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
            farmers = df_village[["FID", "Farmer Name"]].drop_duplicates()

            if farmers.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this village.")
                return

            errors = []

            for _, row in farmers.iterrows():
                current_fid = str(row["FID"])
                current_name = str(row["Farmer Name"])
                self.df_final_filtered = df_village[
                    (df_village["FID"].astype(str) == current_fid) &
                    (df_village["Farmer Name"].astype(str) == current_name)
                ]
                if self.df_final_filtered.empty:
                    continue
                contact = self.df_final_filtered["Contact Number"].astype(str).iloc[0]
                
                if layout ==layout1:            
                    df_transformed = transform_filtered_data(self.df_final_filtered)
                else:
                    df_transformed = transform_filtered_data2(self.df_final_filtered)
                    
                save_path = app_dir / f"{block} {village} {current_fid}.pdf"

                try:
                    layout.generate_pdf(
                        df_transformed, save_path, block, village, current_fid, current_name, contact, self.df_final_filtered
                    )
                except Exception as e:
                    errors.append(f"{current_name}: {e}")

            QMessageBox.information(self, "Done", f"Saved {len(farmers)} PDFs.\n" +
                                    ("Some errors occurred." if errors else ""))
            return

        # Invalid combination
        else:
            QMessageBox.warning(self, "Incomplete Selection", "Please either select:\n• Block + Village only\nOR\n• Block + Village + Farmer ID + Name.")

    def print_pdf(self):
        block = self.block_dropdown.currentText()
        village = self.village_dropdown.currentText()
        fid = self.fid_dropdown.currentText()
        farmer = self.farmer_dropdown.currentText()
        selected_layout_name = self.layout_dropdown.currentText()
        layout = self.layouts.get(selected_layout_name, layout1)
        
        if block == "Select Block" or village == "Select Village":
            QMessageBox.warning(self, "Missing Selection", "Please select at least Block and Village.")
            return

        system = platform.system()
        if system != "Windows":
            QMessageBox.warning(self, "Unsupported", "Printing is only supported on Windows.")
            return

        temp_dir = tempfile.gettempdir()
        
        # Option 1: Single farmer
        if fid != "Select Farmer ID" and farmer != "Select Farmer Name":
            self.df_final_filtered = self.df_full[
                (self.df_full["Block"].astype(str) == block) &
                (self.df_full["Village"].astype(str) == village) &
                (self.df_full["FID"].astype(str) == fid) &
                (self.df_full["Farmer Name"].astype(str) == farmer)
            ]
            contact = self.df_final_filtered["Contact Number"].astype(str).iloc[0]
            if self.df_final_filtered.empty:
                QMessageBox.warning(self, "No Data", "No data found for selected farmer.")
                return

            if layout ==layout1:  
                df_transformed = transform_filtered_data(self.df_final_filtered)
            else:
                df_transformed = transform_filtered_data2(self.df_final_filtered)
                
            temp_path = os.path.join(temp_dir, f"{block}_{village}_{fid}_print.pdf")

            try:
                layout.generate_pdf(
                    df_transformed, temp_path, block, village, fid, farmer, contact, self.df_final_filtered
                )
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

            farmers = df_village[["FID", "Farmer Name"]].drop_duplicates()

            if farmers.empty:
                QMessageBox.warning(self, "No Data", "No farmer records found in this village.")
                return

            errors = []

            for _, row in farmers.iterrows():
                current_fid = str(row["FID"])
                current_name = str(row["Farmer Name"])
                self.df_final_filtered = df_village[
                    (df_village["FID"].astype(str) == current_fid) &
                    (df_village["Farmer Name"].astype(str) == current_name)
                ]
                contact = self.df_final_filtered["Contact Number"].astype(str).iloc[0]  # safe
                if self.df_final_filtered.empty:
                    continue

                if layout ==layout1:
                    df_transformed = transform_filtered_data(self.df_final_filtered)
                else:
                    df_transformed = transform_filtered_data2(self.df_final_filtered)
                    
                temp_path = os.path.join(temp_dir, f"{block}_{village}_{current_fid}_print.pdf")

                try:
                    layout.generate_pdf(
                    df_transformed, temp_path, block, village, current_fid, current_name, contact, self.df_final_filtered
                )
                    os.startfile(temp_path, "print")
                except Exception as e:
                    errors.append(f"{current_name}: {e}")

            QMessageBox.information(self, "Printed", f"Sent {len(farmers)} files to printer.\n" +
                                    ("Some errors occurred." if errors else ""))

        # Invalid selection
        else:
            QMessageBox.warning(self, "Incomplete Selection",
                                "Please either select:\n• Block + Village only\nOR\n• Block + Village + Farmer ID + Name.")


