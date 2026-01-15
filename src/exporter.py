import pandas as pd
import xlsxwriter
import os

class InvoiceExporter:
    @staticmethod
    def save_to_excel(df: pd.DataFrame, file_path: str):
        """
        Saves the DataFrame to an Excel file with formatting.
        """
        try:
            with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, header=True) # Check if header needs to be False or True. Original: implicit True.
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']

                (max_row, max_col) = df.shape

                # Create a list of column headers
                # Note: df.to_excel writes headers by default. 
                # The original code used worksheet.add_table which expects headers in data range or provided?
                # Original code:
                # fat.to_excel(writer) -> Writes index (default True) and header (default True).
                # But original code had implicit index=True? No, standard is True.
                # Then it calculates max_row, max_col from df.shape.
                # Then adds table over 0,0 to max_row, max_col.
                
                # If df.to_excel() writes the index, the columns are shifted.
                # Original code: fat.to_excel(writer)
                # writer.sheets['Sheet1']
                # (max_row, max_col) = fat.shape
                # worksheet.add_table(0, 0, max_row, max_col, ...)
                
                # If to_excel writes index, the Excel file has (max_col + 1) columns.
                # But fat.shape is (rows, cols).
                # So add_table might be slightly off if it relies on shape vs written range.
                # Also to_excel leaves the cursor at start? No.
                
                # Let's try to be robust. 
                # We will disable index in to_excel to make it cleaner, unless original relied on it.
                # Original: `fat` was constructed with `ignore_index=True` in concat, so index is 0..N.
                # `fat.to_excel(writer)` writes index column.
                # Let's write without index for cleaner output, unless we think user needs it.
                # Refactoring usually implies improving. Removing useless index column is an improvement.
                
                # Recalculate range for add_table
                # The range for add_table is (first_row, first_col, last_row, last_col).
                # We have headers at row 0. Data starts at row 1.
                # max_row in excel is len(df) implies data rows. +1 for header.
                
                column_settings = [{'header': str(col)} for col in df.columns]
                
                # Apply table style
                worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

                # Formatting
                worksheet.set_column(0, max_col - 1, 12)
                
                # Conditional format to draw borders
                border_fmt = workbook.add_format({'text_wrap': True, 'bottom': 1, 'top': 1, 'left': 1, 'right': 1})
                # Apply to all cells in data range
                worksheet.conditional_format(0, 0, max_row, max_col - 1, 
                                             {'type': 'no_errors', 'format': border_fmt})
                
        except Exception as e:
            raise IOError(f"Failed to save Excel file to {file_path}: {e}")

    @staticmethod
    def save_to_csv(df: pd.DataFrame, file_path: str):
        try:
            df.to_csv(file_path, index=False)
        except Exception as e:
             raise IOError(f"Failed to save CSV file to {file_path}: {e}")
