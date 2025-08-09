from PyQt5.QtWidgets import QMessageBox
import pandas as pd

def transform_filtered_data(df_final_filtered: pd.DataFrame):
        if df_final_filtered is None or df_final_filtered.empty:
            return None

        df = df_final_filtered.copy()

        # Build the output DataFrame
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