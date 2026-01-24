"""Tests for the parsers module."""
import os
from document_parser.parser.pdf_parser import PdfParser


def test_pdf_parser_can_be_initialized():
    parser = PdfParser(f"{str(os.path.abspath(os.path.dirname(__file__)))}/test_data/jcm.pdf", "tests/output")
    assert parser is not None
