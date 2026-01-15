import unittest
from unittest.mock import MagicMock, patch
import pandas as pd
from src.extractor import InvoiceExtractor

class TestInvoiceExtractor(unittest.TestCase):
    @patch('src.extractor.pdfplumber.open')
    def test_extract_no_pages(self, mock_open):
        # Setup mock
        mock_pdf = MagicMock()
        mock_pdf.pages = []
        mock_open.return_value.__enter__.return_value = mock_pdf

        extractor = InvoiceExtractor("dummy.pdf")
        df = extractor.extract()
        
        self.assertTrue(df.empty)

    @patch('src.extractor.pdfplumber.open')
    def test_extract_page_structure(self, mock_open):
        mock_pdf = MagicMock()
        mock_page = MagicMock()
        mock_pdf.pages = [mock_page]
        
        # text with VKNs
        mock_page.extract_text.return_value = "VKN: 12345\nVKN: 67890"
        
        # tables
        # Table 0: Top table. Needs at least 2 columns.
        table0 = [['c1r1', 'c2r1'], ['c1r2', 'c2r2']]
        
        # Table 1: Middle/Bottom table.
        # Rows with > 2 items are middle.
        # Rows with <= 2 items are bottom.
        table1 = [
            ['m1', 'm2', 'm3'], # Middle
            ['b1', 'b2']        # Bottom
        ]
        
        mock_page.extract_tables.return_value = [table0, table1]
        
        mock_open.return_value.__enter__.return_value = mock_pdf
        
        extractor = InvoiceExtractor("dummy.pdf")
        df = extractor.extract()
        
        self.assertFalse(df.empty)
        # Check that we got a dataframe back
        self.assertIsInstance(df, pd.DataFrame)
        
        # Verify columns are string indices
        self.assertTrue(all(isinstance(c, str) for c in df.columns))

if __name__ == '__main__':
    unittest.main()
