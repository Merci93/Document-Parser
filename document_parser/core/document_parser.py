"""Extract the content of given document, including table of content, images, tables and text."""

import os

from logger import log


class DocumentParser:
    """
    An abstract class responsible for reading and parsing document.

    Subclassed by PDFParser and Word parser.
    """

    def __init__(self, document_path: str, output_directory: str) -> None:
        """
        Open document.

        :param document_path: Full path to the PDF document.
        :param output_directory: Full path to directory, where parsing results are stored.
        """
        log.info(f"Opening document {document_path}.")
        self.document_name = os.path.splitext(os.path.basename(document_path))[0]
        self.output_directory = output_directory
        self.counters = {
            'toc': 0,
            'images': 0,
            'tables': 0,
            'paragraphs': 0,
        }

    def extract_table_of_content(self) -> None:
        """
        Extract table of content from document.

        Content is saved in a CSV file inside a {self.output_directory}/table_of_content directory.
        with same name as the input document name.

        Output CSV file contains two columns in the order:
        - title
        - page number.
        """
        pass

    def extract_images(self) -> None:
        """
        Extract images with figure labels from document and save them.

        All images are saved to `images` subdirectory in self.output_directory tree as PNG images in RGB color mode.
        A CSV file with the name format: <PDF_document_name>_image_data is generated, and contains image_titles and
        page_number in which the image is located in the document.
        Image file name contains:
        1) document name
        2) page number on which the image was found (in PDF documents only).
        3) consecutive image number on that page.
        """
        pass

    def extract_tables(self) -> None:
        """
        Extract tables from document and save them to csv files in {self.output_directory}/tables directory.

        Each extracted table is stored to a separate .CSV file with name:
        1) document name
        2) page number in which the table is located.
        3) consecutive table number on the page.
        """
        pass

    def extract_texts(self) -> None:
        """
        Extract all paragraphs from document and save as a .CSV file.

        Output CSV file is saved in a {self.output_directory}/texts directory with the same name as persed document.

        Relusting file contains two or three columns:
        1) heading
        2) text (text associated with the heading)
        3) page number (in PDF documents only).
        """
        pass

    def run(self) -> None:
        """Parse given document"""
        self.extract_table_of_content()
        self.extract_images()
        self.extract_tables()
        self.extract_texts()
