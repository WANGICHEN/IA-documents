import pdfplumber
from docx import Document
import re

def find_info(text, keyword):
    if not isinstance(text, str):
        text = ''
        return text
    lines = text.splitlines()
    for line in range(len(lines)):
        if keyword.lower() in lines[line].lower():
            if keyword.lower() == 'ratings':
                after_keyword = lines[line][lines[line].find(':') + 1:].strip() + '\n'
                after_keyword += lines[line+1]
            else:
                # 取關鍵字後面部分（原文對應大小寫）
                after_keyword = lines[line][lines[line].find(':') + 1:].strip()
            

            return after_keyword.replace(keyword + ' ', '')
    return None

def extract_texts_from_pdf(pdf):
    if isinstance(pdf, pdfplumber.page.Page):
        pages = [pdf]  # 包成 list 統一處理
    else:
        pages = pdf  # 假設是整本 PDF 文件
    text = ""

    for page in pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text


def extract_field_from_table(pdf, field_name):

    if isinstance(pdf, pdfplumber.page.Page):
        pages = [pdf]  # 包成 list 統一處理
    else:
        pages = pdf  # 假設是整本 PDF 文件
    for page in pages:
        tables = page.extract_tables()
        for table in tables:
            for row in table:
                # 找欄位標題的儲存格位置
                for i, cell in enumerate(row):
                    if cell and field_name.lower() in cell.lower():
                        return cell
    return None

def ectract_factory_name_address(info):
    factory_info = []
    factory_name = []
    pattern = r'(.+?(?:Ltd\.|LTD|Inc\.|INC|LLC|Co\.|CO\.|Company|Corporation|Corp\.|CORP\.))\s*(.*)'
    matches = re.findall(r'\d+\.\s+(.+?)(?=\d+\.|$)', info.replace('\n', ''), re.DOTALL)
    for match in matches:
        # 假設公司名稱在前面，用 "Ltd." 或 "CO.,LTD" 作為切割點
        name_match = re.match(pattern, match.strip(), re.IGNORECASE)
        if name_match:
            factory_name.append(name_match.group(1).strip())
            factory_info.append(name_match.group(2).strip())

    return factory_name, factory_info

def run(pdf_path, word_path):
    with pdfplumber.open(pdf_path) as pdf:
        report_number = find_info(extract_field_from_table(pdf.pages[0], "Report Number"), "Report Number")
        applicant_name = find_info(extract_field_from_table(pdf.pages[0], "Applicant’s name"), "Applicant’s name")
        address = find_info(extract_field_from_table(pdf.pages[0], "Address"), "Address")
        date = find_info(extract_field_from_table(pdf.pages[0], "Date of issue"), "Date of issue")
        standard = find_info(extract_field_from_table(pdf.pages[0], "Standard"), "Standard")
        issuing_lab = find_info(extract_field_from_table(pdf.pages[0], "Name of Testing Laboratory"), "Name of Testing Laboratory")


        model_number = find_info(extract_texts_from_pdf(pdf.pages[:2]), "Model/Type reference")
        description_of_product = find_info(extract_texts_from_pdf(pdf.pages[:2]), "Test item description")
        rating = find_info(extract_texts_from_pdf(pdf.pages[:2]), "Ratings")

        factory_name, factory_info = ectract_factory_name_address(extract_field_from_table(pdf.pages, "Name and address of factory"))

    doc = Document(word_path)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if '{applicant_name}' in cell.text:
                    cell.text = cell.text.replace('{applicant_name}', applicant_name)
                if '{applicant_address}' in cell.text:
                    cell.text = cell.text.replace('{applicant_address}', address)
                if '{model_number}' in cell.text:
                    cell.text = cell.text.replace('{model_number}', model_number)
                if '{description_of_product}' in cell.text:
                    cell.text = cell.text.replace('{description_of_product}', description_of_product)
                if '{rating}' in cell.text:
                    cell.text = cell.text.replace('{rating}', rating)
                if '{report_number}' in cell.text:
                    cell.text = cell.text.replace('{report_number}', report_number)
                if '{date}' in cell.text:
                    cell.text = cell.text.replace('{date}', date)
                if '{standard}' in cell.text:
                    cell.text = cell.text.replace('{standard}', standard)
                if '{issuing_lab}' in cell.text:
                    cell.text = cell.text.replace('{issuing_lab}', issuing_lab)
                if '{factory_name}' in cell.text:
                    cell.text = cell.text.replace('{factory_name}', '\n'.join([f"{i+1}. {item}" for i, item in enumerate(factory_name)]))
                if '{factory_info}' in cell.text:
                    cell.text = cell.text.replace('{factory_info}', '\n'.join([f"{i+1}. {item}" for i, item in enumerate(factory_info)]))

    return doc


