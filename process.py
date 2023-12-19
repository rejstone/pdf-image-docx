import os
import pdf2image
import docx
from docx.shared import Mm
from PIL import Image
from sys import platform

a4_width = 210
a4_height = 297
margins = 10
slash = "/"
poppler = ""

if platform == "win32":
    slash = "\\"
    poppler = "poppler{0}Library{0}bin".format(slash)

if os.path.isdir("image_here") == False:
    os.mkdir("image_here")
if os.path.isdir("doc_here") == False:
    os.mkdir("doc_here")

print("===========================================")
print("+                                         +")
print("+           PDFs to images ...            +")
print("+                                         +")
print("===========================================\n\n")
for input_file in os.listdir("pdf_here"):
    if input_file.endswith(".pdf"):
        pdf = None
        if poppler != "":
            pdf = pdf2image.convert_from_path("pdf_here{1}{0}".format(input_file, slash), 500, poppler_path=poppler)
        else:
            pdf = pdf2image.convert_from_path("pdf_here{1}{0}".format(input_file, slash))
        for index, page in enumerate(pdf):
            page.save("image_here{2}{0}_{1}.png".format(input_file[:-4], index, slash))

print("===========================================")
print("+                                         +")
print("+    Converting to images is done !!!     +")
print("+                                         +")
print("===========================================\n\n")

print("===========================================")
print("+                                         +")
print("+            Images to docx ...           +")
print("+                                         +")
print("===========================================\n\n")
doc = docx.Document()
section = doc.sections[0]
section.page_height = Mm(a4_height)
section.page_width = Mm(a4_width)
section.top_margin = Mm(margins/2)
section.right_margin = Mm(margins/2+1)
section.bottom_margin = section.top_margin
section.left_margin = Mm(margins/2-1)

for input_file in os.listdir("image_here"):
    if input_file.endswith(".png"):
        path = "image_here{1}{0}".format(input_file, slash)
        img = Image.open(path)
        img.resize((a4_width, a4_height))
        img.save(path)
        doc.add_picture(path, width=Mm(a4_width-margins), height=Mm(a4_height-margins))

doc.save("doc_here{0}print.docx".format(slash))
