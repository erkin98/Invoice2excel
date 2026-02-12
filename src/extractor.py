import re
import logging
from typing import List, Optional, Tuple, Any
import pandas as pd
import pdfplumber

logger = logging.getLogger(__name__)

class InvoiceExtractor:
    def __init__(self, file_path: str):
        self.file_path = file_path

    def extract(self) -> pd.DataFrame:
        """
        Extracts invoice data from the PDF file.
        Returns a DataFrame containing the consolidated data from all valid pages.
        """
        all_data = []

        try:
            with pdfplumber.open(self.file_path) as pdf:
                for page in pdf.pages:
                    page_data = self._process_page(page)
                    if page_data is not None:
                        all_data.append(page_data)
        except Exception as e:
            logger.error(f"Error opening or processing PDF {self.file_path}: {e}")
            raise

        if not all_data:
            return pd.DataFrame()

        # Combine all page dataframes
        consolidated_df = pd.concat(all_data, ignore_index=True)
        
        # Rename columns to string indices as per original logic
        # Although keeping semantic names would be better, we'll stick to 0, 1, 2... for now 
        # to match original output structure if that's critical, or perhaps just let pandas handle it.
        # The original code did: fat.columns.values[i] = (str(i))
        consolidated_df.columns = [str(i) for i in range(len(consolidated_df.columns))]
        
        return consolidated_df

    def _process_page(self, page: pdfplumber.page.Page) -> Optional[pd.DataFrame]:
        text = page.extract_text()
        if not text or len(text) == 0:
            logger.warning(f"Page {page.page_number} is empty or skipped.")
            return None

        tables = page.extract_tables()
        if not tables or len(tables) < 2:
            logger.warning(f"Page {page.page_number} does not contain enough tables.")
            return None

        # Extract VKN
        vkn_match = re.findall(r"VKN:\s+(\d+)", text)
        if len(vkn_match) < 2:
            logger.warning(f"Could not find two VKNs on page {page.page_number}. Found: {vkn_match}")
            # Continue with empty VKNs or skip? Original code would fail index access or produce partial data.
            # Looking at original: vkn[0], vkn[1]. It would raise IndexError if < 2.
            # We'll try to handle it gracefully or raise.
            # Let's assume strict format as per 'Invoice2excel' description.
            if len(vkn_match) == 0:
                 vkn_list = ['Unknown', 'Unknown']
            elif len(vkn_match) == 1:
                 vkn_list = [vkn_match[0], 'Unknown']
            else:
                 vkn_list = vkn_match[:2]
        else:
            vkn_list = vkn_match[:2]

        # VKN DataFrames
        vkn1_df = pd.DataFrame(['VKN-1', vkn_list[0]])
        vkn2_df = pd.DataFrame(['VKN-2', vkn_list[1]])

        # First Table Processing (Top Table)
        # Original logic: 
        # tb_u1 = [row[1] for row in table0]
        # tb_u2 = [row[0] for row in table0]
        # tb_u = [tb_u2, tb_u1] -> Transposed?
        # u = pd.DataFrame(tb_u)
        
        table0 = tables[0]
        # Original code assumes table0 has at least 2 columns (index 0 and 1)
        t0_col0 = [row[0] for row in table0 if len(row) > 0]
        t0_col1 = [row[1] for row in table0 if len(row) > 1]
        
        # Pad if lengths differ (unlikely if structured, but good for safety)
        u_df = pd.DataFrame([t0_col0, t0_col1]) # This creates a dataframe where rows are the columns of the table? 
        # Yes, original: tb_u.append(tb_u2); tb_u.append(tb_u1); u = pd.DataFrame(tb_u)
        # So u is 2 rows, N columns.

        # Second Table Processing (Middle/Bottom)
        # Original logic: iterates table1, checks if non-null cells > 2 -> middle, else bottom.
        table1 = tables[1]
        middle_rows = []
        bottom_rows = []

        for row in table1:
            non_null_row = [cell for cell in row if cell is not None]
            if len(non_null_row) > 2:
                middle_rows.append(non_null_row)
            else:
                bottom_rows.append(non_null_row)

        o_df = pd.DataFrame(middle_rows)

        # Bottom Table Processing
        # Original: b_u1 = col1, b_u2 = col0. b_u = [col0, col1].
        b_col0 = []
        b_col1 = []
        for row in bottom_rows:
            if len(row) >= 2:
                b_col0.append(row[0])
                b_col1.append(row[1])
            elif len(row) == 1:
                b_col0.append(row[0])
                b_col1.append(None)
        
        a_df = pd.DataFrame([b_col0, b_col1])

        # Concatenate horizontally
        # pd.concat([VKN1, VKN2, u, o, a], axis=1, ignore_index=True)
        # Note: 'o' (middle table) is a standard DF (rows=rows), but 'u' and 'a' seem transposed (rows=columns).
        # This structure seems weird: 
        # VKN1: 2 rows (header, value) - actually original was col vector?
        # vkn1 = ['VKN-1']; vkn1.append(val); VKN1 = pd.DataFrame(vkn1) -> This is a column vector (2 rows, 1 col).
        
        # Let's verify dimensions.
        # VKN1: 2x1
        # VKN2: 2x1
        # u: 2 x N (where N is rows in table0)
        # o: M x K (where M is rows in middle table)
        # a: 2 x P (where P is rows in bottom table)
        
        # Concatenating axis=1 means we align by index (row number).
        # If these have different number of rows, pandas fills with NaN.
        # This seems to be the intended behavior of the original script.
        
        combined = pd.concat([vkn1_df, vkn2_df, u_df, o_df, a_df], axis=1, ignore_index=True)
        return combined
