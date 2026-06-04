"""docling_dev — модульный конвертер отсканированных PDF → DOCX."""
from .converter import convert_pdf, DoclingBatchConverter
from .pipeline import build_converter

__all__ = ["convert_pdf", "DoclingBatchConverter", "build_converter"]
