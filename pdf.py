from RPA.PDF import PDF
from robot.libraries.String import String
import re

pdf = PDF()
string = String()

def extract_data_from_first_page():
    text = pdf.get_text_from_pdf("PDF/13-76244139-k-Inmobiliaria Los Rosales/Cupon de pago 6037-5  .pdf")
    if str(text).__contains__("3-23")  :
        print("true")
    else:
        print("false")  
    

   
    


extract_data_from_first_page()