"""
pip install pdf2docx docxcompose
"""
import uuid
import fitz
from pdf2docx import Converter
from docx import Document
from docxcompose.composer import Composer


def merge_docx_with_images(docx_list, output):
    master = Document(docx_list[0])
    composer = Composer(master)

    for docx in docx_list[1:]:
        sub_doc = Document(docx)
        composer.append(sub_doc)

    composer.save(output)


def split_pdf_docx(pdf, output_docx):
    cv = Converter(pdf)
    pages = []
    page_count = get_pdf_page_count(pdf)

    for page in range(page_count):
        _uuid = uuid.uuid4()
        file_name = f"./file/temp/{page}_{_uuid}.docx"
        cv.convert(file_name, start=page, end=page + 1)
        pages.append(file_name)

    cv.close()

    # 使用 docxcompose 合并（支持图片）
    merge_docx_with_images(pages, output_docx)

    print(f"共生成 {len(pages)} 页 docx，已合并：{output_docx}")


def get_pdf_page_count(pdf_path):
    doc = fitz.open(pdf_path)
    return doc.page_count


if __name__ == '__main__':
    pdf = "./file/传统医学师承关系合同书.pdf"
    docx = "./file/传统医学师承关系合同书.docx"
    split_pdf_docx(pdf, docx)
