from comtypes.client import CreateObject
import os

folder = '...'  #input address in your folder
wdToPDF = CreateObject('Word.Application')
wdFormatPDF = 17
files = os.listdir(folder)
word_files = [f for f in files if f.endswith(('.doc', '.docx'))]
for word_file in word_files:
    word_path = os.path.join(folder, word_file)
    pdf_path = word_path
    if pdf_path[-3:] != 'pdf':
        pdf_path = pdf_path + '.pdf'

    if os.path.exists(pdf_path):
        os.remove(pdf_path)

    pdfCreate = wdToPDF.Documents.Open(word_path)
    pdfCreate.SaveAs(pdf_path, wdFormatPDF)

