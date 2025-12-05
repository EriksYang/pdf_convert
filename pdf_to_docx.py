"""
pip install pdf2docx
"""
import uuid

import fitz
from pdf2docx import Converter


def split_pdf_docx(pdf, docx):
    cv = Converter(pdf)
    pages = []
    for page in range(get_pdf_page_count(pdf)):
        _uuid = uuid.uuid4()
        file_name = f"./file/temp/{page}_{_uuid}.docx"
        # cv.convert(file_name, start=page, end=page + 1)
        cv.convert(file_name, start=page, end=page + 1)
        pages.append(file_name)
    cv.close()
    print(pages)
    print(len(pages))


def get_pdf_page_count(pdf_path):
    doc = fitz.open(pdf_path)
    return doc.page_count


def singlong_pdf_docx(pdf, docx):
    cv = Converter(pdf)
    cv.convert(docx, start=0, end=None)
    cv.close()


if __name__ == '__main__':
    pdf = "./file/传统医学师承关系合同书.pdf"
    docx = "./file/传统医学师承关系合同书.docx"
    singlong_pdf_docx(pdf, docx)
    # print(get_pdf_page_count(pdf))
    # split_pdf_docx(pdf, docx)
