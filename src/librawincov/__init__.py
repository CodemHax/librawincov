"""
librawincov - Convert DOCX files to PDF using LibreOffice with WSL fallback support.
"""

from librawincov.converter import convert_docx_to_pdf, find_libreoffice

__version__ = "0.1.0"
__all__ = ["convert_docx_to_pdf", "find_libreoffice"]
