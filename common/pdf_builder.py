from reportlab.platypus import SimpleDocTemplate
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

def generate_pdf(df_transformed, save_path, block, village, fid, farmer, contact,
                 df_final_filtered, generate_page1, generate_page2):
    elements = []
    elements += generate_page1(df_final_filtered, block, village, fid, farmer, contact)
    elements += generate_page2(df_transformed, block, village, fid, farmer, contact)

    doc = SimpleDocTemplate(str(save_path), pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=2*cm)
    doc.build(elements)
    
def generate_pdf2(df_transformed, save_path, block, village, fid, farmer, contact,
                 df_final_filtered, generate_page1):
    elements = []
    elements += generate_page1(df_transformed, block, village, fid, farmer, contact, df_final_filtered)

    doc = SimpleDocTemplate(str(save_path), pagesize=A4, leftMargin=1.5*cm, rightMargin=1.5*cm, topMargin=2*cm, bottomMargin=2*cm)
    doc.build(elements)    
