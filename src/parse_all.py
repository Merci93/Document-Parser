"""Function responsible for parsing all available documents."""
import csv
import os

from tqdm import tqdm

from configuration import settings
from logger import log
from pdf_parser import PdfParser
from word_parser import WordParser


def parse_all_documents(
    input_directory: str = settings.file_location,
    output_directory: str = settings.parsed_data_directory,
    ) -> None:
    """
    Parse all documents that exist in input_directory and save results to output_directory.

    All extracted information is saved as CSV files (see relevant methods definitions for more details). Additionaly,
    extracted images are stored as .PNG files.

    The following elements are extracted from each document:
    - table of content
    - images
    - tables
    - text from each paragraph.

    Additionally, two summary csv files ("parse_report.csv" and "images.csv") are created. The first summarizes
    extracted elements counters, and the second contains image titles extracted from each document.
    """
    os.makedirs(output_directory, exist_ok=True)
    
    counters = []
    images = []
    for document_name in tqdm(os.listdir(input_directory)):
        document_path = os.path.join(input_directory, document_name)
        if document_name.endswith("pdf"):
            parser = PdfParser(document_path, output_directory)
        elif document_name.endswith("docx"):
            parser = WordParser(document_path, output_directory)
        else:
            continue
        parser.run()
        parser.counters.update({"document": document_name})
        counters.append(parser.counters)
        images.append(parser.image_data)
    log.info(f"Storing counters to {settings.parse_report} file")
    with open(os.path.join(output_directory, "parse_report.csv"), "w", newline="") as csvfile:
        w = csv.DictWriter(csvfile, counters[0].keys())
        w.writeheader()
        w.writerows(counters)
    log.info(f"Storing image data to {settings.extracted_images_file_name} file")
    columns = ["document_name", "document_type", "figure_number", "figure_title", "page_number",
               "image_filename", "extracted_image"]
    with open(os.path.join(output_directory, settings.extracted_images_file_name), "w", newline="") as csvfile:
        w = csv.DictWriter(csvfile, columns)
        w.writeheader()
        for sublist in images:
            w.writerows(sublist)


if __name__ == "__main__":
    parse_all_documents()
