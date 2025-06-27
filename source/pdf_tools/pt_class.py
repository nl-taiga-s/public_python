from pathlib import Path

from pypdf import PdfReader


class PdfTools:
    """PDFを扱うツール"""

    def __init__(self):
        """初期化します"""
        print(self.__class__.__doc__)

    def read_metadata(self):
        pdf_path = Path("~/pocket/sample_pptx.pdf").expanduser()
        reader = PdfReader(pdf_path)
        self.meta = reader.metadata

    def print_metadata(self):
        print(f"{self.meta.title}")
        print(f"{self.meta.author}")
        print(f"{self.meta.subject}")
        print(f"{self.meta.creator}")
        print(f"{self.meta.producer}")
        print(f"{self.meta.creation_date}")
        print(f"{self.meta.modification_date}")
