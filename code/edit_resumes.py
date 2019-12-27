from docx import Document
from docx.shared import Pt
import win32com.client as client
import pathlib
import csv
import argparse
from progress.bar import ChargingBar

def convert_to_pdf(filepath:str, target_path:str):
    """Save a pdf of a docx file."""    
    try:
        word = client.DispatchEx("Word.Application")
        word_doc = word.Documents.Open(filepath)
        word_doc.SaveAs(target_path, FileFormat=17)
        word_doc.Close()
    except Exception as e:
            raise e
    finally:
            word.Quit()

def replace_position_company(doc_name, company, position):
    filepath = pathlib.Path.cwd().joinpath("to_update", doc_name)
    document = Document(filepath)
    style = document.styles['Normal']
    font = style.font
    font.name = 'Arial'
    font.size = Pt(10)
    for par in document.paragraphs:
        current_p = par.text
        if( "<company>" in current_p) | ("<position>" in current_p):
            par.text = current_p.replace("<company>", company).replace("<position>", position)
    new_document_name = f"{company}_{position}_{doc_name}"
    filepath = pathlib.Path.cwd().joinpath("docs", new_document_name)
    
    document.save(filepath)

    target_path = pathlib.Path.cwd().joinpath("pdfs", new_document_name.replace(".docx", ".pdf"))

    convert_to_pdf(str(filepath), str(target_path))

def update_all_files(doc_list, csv_path):
    with open(csv_path) as f:
        comp_pos_tups = [tuple(line) for line in csv.reader(f)]
    bar = ChargingBar('Processing', max=len(comp_pos_tups[1:]))
    bar.start()
    for company, position in comp_pos_tups[1:]:
        for doc in doc_list:
            replace_position_company(doc, company, position)
        bar.next()
    bar.finish()

if __name__ == "__main__":
    parser = argparse.ArgumentParser()

    parser.add_argument("--csv_path", help="CSV file containing job posting companies and positions.")

    arguments = parser.parse_args()
    documents = [doc.name for doc in pathlib.Path("to_update").glob("*.docx")]
    update_all_files(documents, **vars(arguments))
