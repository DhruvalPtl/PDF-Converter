import PyPDF2
import pdfkit
from docx2pdf import convert
from pdf2docx import Converter
from win32com import client
import os
import shutil
import tabula
import pandas as pd
from pptx import Presentation
from PyPDF2 import PdfReader
import img2pdf
import fitz
class PDFConverter():
    def __init__(self):
        self.input_path_dir = "D:/Python/project 1/Pdfconverter/TEMP"
        os.chdir(self.input_path_dir)
        self.input_files = [i for i in os.listdir() if  i.endswith(".pdf")]
        self.output_path = "D:/Python/project 1/Pdfconverter/TEMP1"
        
    def pdfmerger(self):
        os.chdir(self.input_path_dir)
        pdf_files = [i for i in os.listdir() if  i.endswith(".pdf")]
        merger = PyPDF2.PdfMerger()

        for pdf in pdf_files:
            with open(pdf, 'rb') as b:
                merger.append(b)
        merger.write(self.output_path + f"/{pdf}")

    def pdfcompress(self):
        os.chdir(self.input_path_dir)
        pdf_files = [i for i in os.listdir() if  i.endswith(".pdf")]
        for pdf in pdf_files:
            with open(pdf, 'rb') as b:
                pdf_reader = PyPDF2.PdfReader(b)
                pdf_writer = PyPDF2.PdfWriter()

            for page in pdf_reader.pages:
                page.compress_content_streams()
                pdf_writer.add_page(page)
            pdf_writer.write(self.output_path + f"/{pdf}")

    def split_range(self, start_page, end_page):
        try:
            for files in self.input_files:
                with open(files, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    range_writer = PyPDF2.PdfWriter()
                    for page_num in range(start_page - 1, min(end_page, len(reader.pages))):
                        range_writer.add_page(reader.pages[page_num])
                    range_writer_output = os.path.join(self.output_path, f'{files}')
                    range_writer.write(range_writer_output)
        except Exception as e:
            print(f"Error splitting PDF range: {e}")
            
    def extract_pages(self, page_numbers):
        try:
            for input_file in self.input_files:
                with open(input_file, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    extract_writer = PyPDF2.PdfWriter()
                    for page_num in page_numbers:
                        page_index = page_num - 1  
                        if page_index < len(reader.pages):
                            extract_writer.add_page(reader.pages[page_index])
                    extract_writer_output = f"{self.output_path}/split_extracted.pdf"
                    extract_writer.write(extract_writer_output)
        except Exception as e:
            print(f"Error extracting PDF pages: {e}")

    def single_pages(self):
        try:
            for input_path in self.input_files:
                with open(input_path, 'rb') as file:
                    reader = PyPDF2.PdfReader(file)
                    for page_num, page in enumerate(reader.pages):
                        single_page_writer = PyPDF2.PdfWriter()
                        single_page_writer.add_page(page)
                        single_page_output = f"{self.output_path}/split_page_{page_num + 1}.pdf"
                        single_page_writer.write(single_page_output)
        except Exception as e:
            print(f"Error splitting PDF into single pages: {e}")

    def emptydir(self):
        try:
            shutil.rmtree(self.input_path_dir,  ignore_errors=True)
        except Exception as e:
            print(f"Error removing directory: {e}")

    def html2pdf(self):
        config = pdfkit.configuration(wkhtmltopdf="D:/Python/project 1/Pdfconverter/wkhtmltopdf/bin/wkhtmltopdf.exe")

        os.chdir(self.input_path_dir)
        html_files = [i for i in os.listdir() if  i.endswith(".html")]
        for i in html_files:
            output_pdf = os.path.join(self.output_path,f"{i[:-5]}.pdf")
            pdfkit.from_file(i, output_pdf, configuration=config)
            
    def pdf2word(self):
        try:
            for pdf_file in self.input_files:
                output_word = os.path.join(self.output_path, pdf_file[:-4] + '.docx')
                cv = Converter(pdf_file)
                cv.convert(output_word, start=0, end=None)
                cv.close()
        except Exception as e:
            print(f"Error during conversion: {e}")

    def word2pdf(self):
        try:
            word_files = [f for f in os.listdir() if f.endswith('.docx')]
            for word_file in word_files:
                input_word = os.path.join(self.input_path_dir, word_file)
                output_pdf = os.path.join(self.output_path, word_file[:-5] + '.pdf')
                convert(input_word, output_pdf)
            print("Conversion successful!")
        except Exception as e:
            print(f"Error during conversion: {e}")

    def excel2pdf(self):
        excel_files = [i for i in os.listdir() if  i.endswith(".xlsx")]
        for i in excel_files:
            df = pd.read_excel(i)
            df.to_html(os.path.join(self.input_path_dir,"Temp.html"))
            output_pdf = os.path.join(self.output_path,f"{i[:-5]}.pdf")
            pdfkit.from_file(os.path.join(self.input_path_dir,"Temp.html"),output_pdf)

    def pdf2excel(self):
        for pdf in self.input_files:
            try:
                tables = tabula.read_pdf(pdf, pages='all', multiple_tables=True)
                combined_table = pd.concat(tables)
                for c in combined_table.columns:
                    try:
                        combined_table[c] = pd.to_numeric(combined_table[c])
                    except ValueError:
                        pass
                output_pdf = os.path.join(self.output_path,f"{pdf[:-4]}.xlsx")
                combined_table.to_excel(output_pdf, index=False)
            except Exception as e:
                print(f"Error during conversion: {e}")

    def pdf2pptx(self):
        for pdf_file in self.input_files:
            pptx_file = os.path.join(self.output_path, os.path.splitext(pdf_file)[0] + ".pptx")
            prs = Presentation()
            pdf_reader = PdfReader(pdf_file)
            for page_num in range(len(pdf_reader.pages)):
                slide_layout = prs.slide_layouts[5]
                slide = prs.slides.add_slide(slide_layout)

                text = pdf_reader.pages[page_num].extract_text()
                text_box = slide.shapes.add_textbox(left=0, top=0, width=prs.slide_width, height=prs.slide_height)
                text_frame = text_box.text_frame
                text_frame.text = text

            prs.save(pptx_file)

    def pptx2pdf(self): 
        pptx_files = [i for i in os.listdir() if i.endswith(".pptx")]
        for pptx_file in pptx_files:
            output_pdf = os.path.join(self.output_path, f"{pptx_file[:-5]}.pdf")

            powerpoint = client.Dispatch("Powerpoint.Application")
            powerpoint.Visible = True
            presentation = powerpoint.Presentations.Open(os.path.abspath(pptx_file))
            presentation.SaveAs(os.path.abspath(output_pdf), 32)
            
            presentation.Close()
            powerpoint.Quit()

    def image2pdf(self):
        img_files = [i for i in os.listdir() if  i.endswith(".jpg") or i.endswith(".png")]
        output_pdf = os.path.join(self.output_path,f'{os.path.splitext(img_files[0])[0]}.pdf')
        with open(output_pdf, "wb") as f:
            f.write(img2pdf.convert(img_files))

    def pdf2images(self):

        for pdf_file in self.input_files:
            document = fitz.open(pdf_file)

            for page_number in range(len(document)):
                page = document.load_page(page_number)
                pixmap = page.get_pixmap()
                image_path = os.path.join(self.output_path, f"page_{page_number + 1}.png")
                pixmap.save(image_path)

        document.close()

#Using Microsoft Excel
# def excel_to_pdf(input_excel, output_pdf):
#     excel = client.CreateObject("Excel.Application")
#     excel.Visible = False
    
#     workbook = excel.Workbooks.Open(os.path.abspath(input_excel))
    
#     # Set the zoom factor
#     for sheet in workbook.Sheets:
#         sheet.PageSetup.Zoom = 85
    
#     # Export as PDF
#     workbook.ExportAsFixedFormat(0, os.path.abspath(output_pdf))  # 0 is the value for PDF format
    
#     # Close Excel and release resources
#     workbook.Close(False)
#     excel.Quit()

#     print("Conversion successful!")

# # Example usage
# input_excel = "D:/Python/project 1/Pdfconverter/TEMP/1.xlsx"
# output_pdf = "D:/Python/project 1/Pdfconverter/TEMP1/1.pdf"
# excel_to_pdf(input_excel, output_pdf)