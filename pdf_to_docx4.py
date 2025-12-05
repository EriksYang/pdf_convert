"""
pip install pdf2docx docxcompose
"""
import os
import uuid
import fitz
from pdf2docx import Converter
from docx import Document
from docxcompose.composer import Composer


TEMP_DIR = "./file/temp"


def ensure_temp_dir():
    if not os.path.exists(TEMP_DIR):
        os.makedirs(TEMP_DIR)


def split_pdf_to_single_page_docx(pdf_path):
    """将 PDF 每页转换为独立的 docx 文件，并返回文件列表"""
    ensure_temp_dir()

    cv = Converter(pdf_path)
    page_count = get_pdf_page_count(pdf_path)

    docx_pages = []

    for page in range(page_count):
        temp_name = f"{TEMP_DIR}/{page}_{uuid.uuid4()}.docx"
        cv.convert(temp_name, start=page, end=page + 1)
        docx_pages.append(temp_name)

    cv.close()
    return docx_pages


def merge_docx_with_section_break(docx_list, output_docx):
    """使用 docxcompose 合并 docx，不改变版式，每个子文档强制从新页开始"""
    master = Document(docx_list[0])
    composer = Composer(master)

    for sub in docx_list[1:]:
        sub_doc = Document(sub)
        composer.append(sub_doc)   # docxcompose 自动使用 section break

    composer.save(output_docx)


def cleanup_temp_files(docx_list):
    """删除临时 docx 文件"""
    for f in docx_list:
        if os.path.exists(f):
            os.remove(f)


def split_pdf_docx(pdf, output_docx):
    pages = split_pdf_to_single_page_docx(pdf)

    merge_docx_with_section_break(pages, output_docx)

    cleanup_temp_files(pages)

    print(f"已成功生成：{output_docx}，临时文件已删除。共 {len(pages)} 页。")


def get_pdf_page_count(pdf_path):
    doc = fitz.open(pdf_path)
    return doc.page_count


if __name__ == "__main__":
    pdf = "./file/传统医学师承关系合同书.pdf"
    docx = "./file/传统医学师承关系合同书.docx"
    split_pdf_docx(pdf, docx)
