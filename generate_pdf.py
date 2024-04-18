import openpyxl
import reportlab
import reportlab.pdfgen.canvas

excel_file = r"C:\Users\olvm\Downloads\ndc-techtown-2024 flattened sessions - exported 2024-04-17(1).xlsx"
wb = openpyxl.load_workbook(excel_file)
sheet = wb.active

pdf_filename = "output.pdf"
c = reportlab.pdfgen.canvas.Canvas(pdf_filename, pagesize=reportlab.lib.pagesizes.landscape(reportlab.lib.pagesizes.A4))

title_font_size = 16
name_font_size = 16
details_font_size = 12
description_font_size = 10
margin = 50
line_height = 12
(page_height, page_width) = reportlab.lib.pagesizes.A4

for row in sheet.iter_rows(min_row=2, values_only=True):
    pid = row[0]
    title = row[1]
    name = row[3]
    form = row[5]
    level = row[6]
    tags = row[8]
    loc = row[11]
    status = row[13]
    if status != "Nominated":
        continue
    if not "day" in form:
        continue
    
    if row[2]:
        description = row[2].replace('_x000D_', '')
    else:
        description = ""

    c.setFont("Helvetica-Bold", title_font_size)
    c.drawString(margin, page_height - margin, title)
    c.setFont("Helvetica", name_font_size)
    c.drawString(margin, page_height - margin - 2 * line_height, f"{name}, {form}, {level}")
    c.setFont("Helvetica", details_font_size)
    c.drawString(margin, page_height - margin - 4 * line_height, f"{pid}, [{loc}], Tags: {tags}")
    c.setFont("Helvetica", description_font_size)
    description_lines = reportlab.lib.utils.simpleSplit(description, "Helvetica", description_font_size, page_width - 2*margin)
    y = page_height - 3*(page_height/10)
    for line in description_lines:
        if y < 2*line_height:
            c.drawString(margin, y, "... cut ...")
            break
        c.drawString(margin, y, line)
        y -= line_height
    c.showPage()

c.save()

print(f"Done! {pdf_filename} created from {excel_file}")
