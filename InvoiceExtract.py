#!/bin/env python

import pdfplumber
from openpyxl import Workbook
import re
import sys
import os

def extract_text_from_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = ""
        for page in pdf.pages:
            text += page.extract_text()
    return text

def extract_info_from_text(pdf_text):
    billing_date = re.search(r"开票日期\s*[:：]\s*(.*)", pdf_text).group(1).replace(" ", "")
    invoice_code = re.search(r"发票代码\s*[:：]\s*(\d*)", pdf_text).group(1)
    invoice_number = re.search(r"发票号码\s*[:：]\s*(\d*)", pdf_text).group(1)
    invoice_issuer = re.findall(r"名 称\s*[:：]\s*(\w*)", pdf_text)[1]
    total_amount = re.search(r"小写(.*)", pdf_text).group(1).replace(" ", "")[1:]

    return billing_date, invoice_code, invoice_number, invoice_issuer, total_amount

def process_pdf_path(pdf_path, out_path):
    pdf_files = []
    
    for file in os.listdir(pdf_path):
        if file.endswith('.pdf'):
            pdf_files.append(os.path.join(pdf_path, file))

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['开票日期', '发票代码', '发票号码', '开票方', '票面金额'])

    for pdf_file in pdf_files:
        pdf_text = extract_text_from_pdf(pdf_file)
        try:
            sheet.append(extract_info_from_text(pdf_text))
        except AttributeError:
            print("Reading error file: " + pdf_file)

        workbook.save(out_path)


if __name__ == '__main__':
    args = sys.argv

    try:
        pdf_path, out_path = args[1], args[2]
    except IndexError:
        print("Please enter the invoice path or output file path.")    
        exit()

    process_pdf_path(pdf_path, out_path)