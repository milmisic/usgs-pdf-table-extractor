"""
PDF/Word Document Table Extractor

Automated extraction of tables from PDF and Word documents into structured datasets.
"""

__version__ = "0.1.0"
__author__ = "Mile Mišić"

from .extractor import DocxTableExtractor
from .pdf_converter import PDFConverter

__all__ = ["DocxTableExtractor", "PDFConverter"]

