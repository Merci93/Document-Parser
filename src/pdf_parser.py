"""
Extract the content of a PDF document, including table of content, images, tables and text.
"""
import os
import re

import fitz
import pandas
import pdfplumber
from unidecode import unidecode

from configuration import settings
from document_parser import DocumentParser
from logger import log


class PdfParser(DocumentParser):
    """A Class responsible for reading and parsing PDF document."""

    def __init__(self, document_path: str, output_directory: str) -> None:
        """
        Open document.

        As for different purposes two independent parsing libraries are used: fitz (PyMuPDF) and pdfplumber,
        both parsers are loaded here.

        :param document_path: Full path to the PDF document.
        :param output_directory: Full path to directory, where parsing results are stored.
        """
        super().__init__(document_path, output_directory)
        self.document_fitz = fitz.open(document_path)
        self.document_plumber = pdfplumber.open(document_path)

    def extract_table_of_content(self) -> None:
        """
        Extract table of content from PDF document.

        TOC is extracted using PyMuPDF library. Unfortunately it cannot recognize TOC in each document.
         In case this method doesn't return anything, the second try uses regular expression. In this approach
         first 21 pages are checked (we assume that TOC cannot be later in the document).
        """
        final_directory = os.path.join(self.output_directory, settings.extracted_table_of_content)
        os.makedirs(final_directory, exist_ok=True)
        table_of_content = []
        raw_table_of_content = self.document_fitz.get_toc(simple=False)
        for content in raw_table_of_content:
            title = content[1].strip()
            page_number = content[2]
            table_of_content.append((title, page_number))
        if len(table_of_content) == 0:
            try:
                for i in range(21):
                    page = self.document_fitz[i]
                    text = page.get_text("text")
                    lines = text.split("\n")
                    for line in lines:
                        line = line.strip()
                        toc_pattern = r"^\d+(\.\d+)*\s+[A-Za-z\s]+\.+\s+\d+$"
                        title_match = re.match(toc_pattern, line)
                        if line != "" and title_match:
                            remove_trailing_dots = re.sub(r"\.+(\s+\d+)?$", "", line)
                            page_number = re.search(r"(\d+)\s*$", line)
                            table_of_content.append((remove_trailing_dots, page_number.group(1)))
            except IndexError:
                log.warning("No table of content found.")
        log.info(f"Extracted TOC with {len(table_of_content)} titles.")
        toc_df = pandas.DataFrame(table_of_content, columns=["title", "page_number"])
        toc_df.to_csv(os.path.join(final_directory, f"{self.document_name}.csv"), index=False)
        self.counters['toc'] = len(table_of_content)

    def extract_images(self) -> None:
        """Extract images with figure labels from document and save them."""
        log.info("Extracting images...")
        final_directory = os.path.join(self.output_directory, settings.extracted_images)
        os.makedirs(final_directory, exist_ok=True)

        def save_image(save_as: str, final_directory: str, img_bbox: fitz.Rect) -> None:
            """Save extracted image.

            :param save_as: image name
            :param final_directory: directory where image will be saved.
            :param img_bbox: image bounding box.
            """
            image = page_content.get_pixmap(matrix=fitz.Matrix(1, 1).prescale(2, 2), clip=img_bbox)
            image.save(os.path.join(final_directory, save_as))

        counter = 0
        image_data = []
        for page_number, page_content in enumerate(self.document_fitz, start=1):
            text_blocks = page_content.get_text("dict")["blocks"]
            for idx, block in enumerate(text_blocks):
                try:
                    for i in range(len(block["lines"]) + 1):
                        figure_block = block["lines"][i]["spans"][0]
                        figure_match = re.match(r"^Figure\s+\d+(?:[\s:-].*)?", figure_block["text"].strip())
                        font_type = figure_block["font"]
                        if "lines" in block and figure_match and "Bold" in font_type:
                            title_block = block["lines"][i]["spans"][0]
                            title_bbox = fitz.Rect(title_block["bbox"])
                            title_text = title_block["text"].strip()
                            title_match = re.match(r"^(Figure\s+\d+)(?:[\s:-])-?(.*)", title_text)
                            figure_number = title_match.group(1)
                            save_as = f"{self.document_name}_{page_number}_{figure_number}.png"
                            figure_data = {
                                "document_name": self.document_name,
                                "document_type": "pdf",
                                "figure_number": figure_number,
                                "figure_title": title_match.group(2).strip(),
                                "page_number": page_number,
                                "image_filename": save_as,
                                "extracted_image": "No. Image on a different page"
                            }
                            text_match = r"^[A-Za-z0-9].*"
                            prev_block = text_blocks[idx - 1]["lines"][0]["spans"][0]
                            is_prev_text = re.match(text_match, prev_block["text"])
                            prev_bbox = fitz.Rect(prev_block["bbox"])
                            try:
                                next_block = text_blocks[idx + 1]["lines"][0]["spans"][0]
                                next_bbox = fitz.Rect(next_block["bbox"])
                            except (KeyError, IndexError):
                                image_data.append(figure_data)
                                break

                            image_data.append(figure_data)
                            if prev_bbox.x1 - title_bbox.x0 > 80 and not is_prev_text:
                                prev_block = text_blocks[idx - 1]["lines"][0]["spans"][0]
                                prev_bbox = fitz.Rect(prev_block["bbox"])
                                img_bbox = fitz.Rect(
                                    -title_bbox.x1,
                                    -title_bbox.y1,
                                    prev_bbox.x1,
                                    prev_bbox.y1,
                                )
                                save_image(save_as, final_directory, img_bbox)
                                image_data[-1]["extracted_image"] = "Yes"
                                counter += 1

                            elif next_bbox.x1 - title_bbox.x0 > 80:
                                img_bbox = fitz.Rect(
                                    title_bbox.x0 - 5,
                                    title_bbox.y0 + 12,
                                    next_bbox.x1,
                                    next_bbox.y1,
                                )
                                save_image(save_as, final_directory, img_bbox)
                                image_data[-1]["extracted_image"] = "Yes"
                                counter += 1

                            elif next_bbox.x1 - title_bbox.x0 < 80:
                                for j in range(2, 6):
                                    next_block = text_blocks[idx + j]["lines"][0]["spans"][0]
                                    next_bbox = fitz.Rect(next_block["bbox"])
                                    if next_bbox.x1 - title_bbox.x0 > 80 and next_block["text"] == " ":
                                        img_bbox = fitz.Rect(
                                            title_bbox.x0 - 5,
                                            title_bbox.y0 + 15,
                                            next_bbox.x1,
                                            next_bbox.y1,
                                        )
                                        save_image(save_as, final_directory, img_bbox)
                                        image_data[-1]["extracted_image"] = "Yes"
                                        counter += 1
                except (KeyError, IndexError, AttributeError):
                    pass
        log.info(f"...done. Successfully extracted {counter} images.")
        self.counters["images"] = counter
        self.image_data = image_data

    def extract_tables(self) -> None:
        """Extract tables from PDF document and save them."""
        final_directory = os.path.join(self.output_directory, settings.extracted_tables)
        os.makedirs(final_directory, exist_ok=True)
        counter = 0
        log.info("Extracting tables...")
        for page in self.document_plumber.pages:
            table_objects = page.find_tables(
                table_settings={
                    "vertical_strategy": "lines",
                    "horizontal_strategy": "lines",
                    "intersection_x_tolerance": 2,
                    "snap_tolerance": 8,
                    "join_tolerance": 40,
                }
            )
            for table_counter, table in enumerate(table_objects, start=1):
                extract_table = table.extract()
                header = extract_table[0]
                data = extract_table[1:]
                save_as = f"{self.document_name}_{page.page_number}_{table_counter}.csv"
                if len(header) >= 2:
                    table_dataframe = pandas.DataFrame(data, columns=header)
                    table_dataframe.dropna(axis=0, how="all")
                    table_dataframe.to_csv(os.path.join(final_directory, save_as), index=False)
                    counter += 1
                log.debug(f"Extracted table {header}.")
        log.info(f"...done. Successfully extracted {counter} tables.")
        self.counters['tables'] = counter

    def extract_texts(self) -> None:
        """Extract text from PDF document and save as CSV file."""
        log.info("Extracting text...")
        final_directory = os.path.join(self.output_directory, settings.extracted_texts)
        os.makedirs(final_directory, exist_ok=True)

        def not_within_bboxes(obj) -> bool:
            """
            Check if the object is in any of the table's bounding box.

            :param obj: text object from PDF document
            :return: boolean indicating text object is outside a table bounding box
            """

            def obj_in_bbox(_bbox) -> bool:
                """
                Get the bounding box dimensions of a text object

                :param _bbox: table bounding box
                :return: boolean indicating if text object is in a table bounding box
                """
                v_mid = (obj["top"] + obj["bottom"]) / 2
                h_mid = (obj["x0"] + obj["x1"]) / 2
                x0, top, x1, bottom = _bbox
                return (h_mid >= x0) and (h_mid < x1) and (v_mid >= top) and (v_mid < bottom)

            return not any(obj_in_bbox(__bbox) for __bbox in bounding_boxes)

        def get_table_settings(strategy: str) -> dict:
            """Return table settings dictionary. Strategy should be one of `explicit`, `lines`."""
            assert strategy in ("explicit", "lines"), "Table settings strategy should be 'explicit' or 'lines'."
            return {
                "vertical_strategy": strategy,
				"horizontal_strategy": strategy,
				"explicit_vertical_lines": page.curves + page.edges,
				"explicit_horizontal_lines": page.curves + page.edges,
				"snap_tolerance": 8,
				"join_tolerance": 50,
    		}

        extracted_sentences = []
        for page in self.document_plumber.pages:
            try:
                bounding_boxes = [
                    table.bbox for table in page.find_tables(table_settings=get_table_settings("explicit"))
                ]
            except ValueError:
                bounding_boxes = [
                    table.bbox for table in page.find_tables(table_settings=get_table_settings("lines"))
                ]
            sentence_lines = page.filter(not_within_bboxes).extract_words(
                keep_blank_chars=True,
                use_text_flow=True,
                extra_attrs=["fontname", "size"],
            )
            for sentence in sentence_lines:
                text = unidecode(sentence["text"])
                font_type = sentence["fontname"]
                font_size_ = sentence["size"]
                is_bold = False
                if "bold" in font_type.lower():
                    is_bold = True
                table_title = re.match(r"^Table\s+\d+", text)
                figure_title = re.match(r"^Figure\s+\d+", text)
                page_marker = re.match(r"^Page\s+\d+", text)
                table_of_content_1 = re.match(r"^[\w\s]+\.+\s+\d+$", text)
                table_of_content_2 = re.match(r"^(\d+(?:\.\d+)*)\.?\s+(.+)\d$", text)

                if ((text.replace(" ", "") != "")
                    and (not table_title)
                    and (not figure_title)
                    and (not page_marker)
                    and (not table_of_content_1)
                    and (not table_of_content_2)
                    ):
                    extracted_sentences.append((text, page.page_number, font_type, is_bold, font_size_))
        text_data_frame = pandas.DataFrame(
            extracted_sentences,
            columns=["text", "page_number", "font_type", "is_bold", "font_size"],
        )
        headers = []
        body_text = []
        page_num = []
        tmp = []
        for _, row_value in text_data_frame.iterrows():
            page_num_ = row_value.page_number
            text = row_value.text
            if row_value.is_bold:
                headers.append(text)
                body_text.append("\n".join(tmp))
                page_num.append(page_num_)
                tmp = []
            else:
                tmp.append(text)
        body_text.append("\n".join(tmp))
        body_text = body_text[1:]
        text_df = pandas.DataFrame(zip(headers, body_text, page_num), columns=["heading", "text", "page_number"])
        text_df.to_csv(os.path.join(final_directory, f"{self.document_name}.csv"), index=False)
        log.info(f"...done. Successfully extracted {len(headers)} paragraphs.")
        self.counters['paragraphs'] = len(headers)


def parse_all_pdf_documents(
    input_directory: str = settings.file_location,
    output_directory: str = settings.parsed_data_directory,
    ) -> None:
    """Parse all pdf documents that exist in input_directory and save results to output_directory."""
    for document in os.listdir(input_directory):
        if document.endswith("pdf"):
            try:
                parser = PdfParser(os.path.join(input_directory, document), output_directory)
            except (fitz.EmptyFileError, fitz.FileDataError) as e:
                log.error(f"{document} cannot be parsed. Error: {e}")
                break
            parser.extract_table_of_content()
            parser.extract_images()
            parser.extract_tables()
            parser.extract_texts()
            log.info(f"{document} successfully processed.")
        else:
            log.info(f"Skipping not a PDF document {document}.")


if __name__ == "__main__":
    parse_all_pdf_documents()
