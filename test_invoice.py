from fpdf import FPDF
from fpdf.enums import XPos, YPos

pdf = FPDF()
pdf.add_page()
pdf.set_font("Helvetica", size=12)

pdf.cell(200, 10, text="Hello PDF world!", new_x=XPos.LEFT, new_y=YPos.NEXT)
pdf.output("test_invoice.pdf")
