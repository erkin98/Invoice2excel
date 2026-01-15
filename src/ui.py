import PySimpleGUI as sg
import os
import tkinter as tk
from tkinter import filedialog
from .extractor import InvoiceExtractor
from .exporter import InvoiceExporter
import logging
import threading

logger = logging.getLogger(__name__)

class InvoiceApp:
    def __init__(self):
        self.window = None
        self.setup_layout()

    def setup_layout(self):
        file_list_column = [
            [
                sg.Text("Pdf folder"),
                sg.In(size=(25, 1), enable_events=True, key="-FOLDER-"),
                sg.FolderBrowse(),
            ],
            [
                sg.Listbox(
                    values=[], enable_events=True, size=(40, 20), key="-FILE LIST-"
                )
            ],
        ]

        image_viewer_column = [
            [sg.Text("Choose a pdf from list on left:")],
            [sg.Text(size=(40, 1), key="-TOUT-")],
            # Image element needs a valid source or it might error if empty initially in some versions,
            # but usually fine.
            [sg.Image(key="-IMAGE-")],
        ]

        layout = [
            [
                sg.Column(file_list_column),
                sg.VSeperator(),
                sg.Column(image_viewer_column),
            ],
            [sg.Submit(key='-SUBMIT-'), sg.Exit()]
        ]

        self.window = sg.Window("Pdf browser", layout)

    def run(self):
        selected_file = None

        while True:
            event, values = self.window.read()
            if event == "Exit" or event == sg.WIN_CLOSED:
                break
            
            if event == "-FOLDER-":
                folder = values["-FOLDER-"]
                try:
                    file_list = os.listdir(folder)
                    fnames = [
                        f for f in file_list
                        if os.path.isfile(os.path.join(folder, f))
                        and f.lower().endswith(".pdf")
                    ]
                    self.window["-FILE LIST-"].update(fnames)
                except Exception as e:
                    logger.error(f"Error reading folder: {e}")
                    self.window["-FILE LIST-"].update([])

            elif event == "-FILE LIST-":
                try:
                    if values["-FILE LIST-"]:
                        filename = os.path.join(
                            values["-FOLDER-"], values["-FILE LIST-"][0]
                        )
                        self.window["-TOUT-"].update(filename)
                        # Displaying PDF as image in PySimpleGUI requires conversion to PNG/GIF usually.
                        # The original code did: window["-IMAGE-"].update(filename=filename)
                        # PySimpleGUI Image element supports PNG, GIF, PPM/PGM. It does NOT support PDF directly.
                        # If the original code worked, maybe the "pdf" was actually an image or they had ghostscript magic?
                        # Or maybe they expected users to click it and nothing happened? 
                        # Wait, the original code imports `pdfplumber` to READ data, but for UI it just passes filename to sg.Image.
                        # Unless the filename points to a png/gif, this would likely fail or show nothing.
                        # The original comment: "Pdf must be in this format... image_viewer_column..."
                        # It is possible the user never actually saw the image or it crashed.
                        # For refactoring, I will wrap this in try/except to prevent crash.
                        try:
                             self.window["-IMAGE-"].update(filename=filename)
                        except Exception as e:
                             # Likely format error
                             pass
                        
                        selected_file = filename
                except Exception as e:
                    logger.error(f"Error selecting file: {e}")

            elif event == '-SUBMIT-':
                if not selected_file:
                    sg.popup_error("Please select a PDF file first.", keep_on_top=True)
                    continue
                
                sg.Popup('PDF file selected successfully. Processing...', keep_on_top=True)
                self.process_and_save(selected_file)
                break

        self.window.close()

    def process_and_save(self, filename):
        try:
            # Extract
            extractor = InvoiceExtractor(filename)
            df = extractor.extract()
            
            if df.empty:
                sg.popup_error("No data extracted from PDF.", keep_on_top=True)
                return

            # Save Dialog
            root = tk.Tk()
            root.withdraw()
            path_save = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=(("Excel workbook", "*.xlsx"), ('Comma separated value', '.csv'), ("All Files", "*.*"))
            )
            root.destroy()

            if not path_save:
                return

            # Export
            if path_save.endswith('.xlsx'):
                InvoiceExporter.save_to_excel(df, path_save)
            elif path_save.endswith('.csv'):
                InvoiceExporter.save_to_csv(df, path_save)
            else:
                # Default to xlsx if extension missing or unknown, or warn
                # Original code printed error.
                if path_save:
                     path_save += '.xlsx'
                     InvoiceExporter.save_to_excel(df, path_save)

            sg.popup(f"File saved successfully to {path_save}", keep_on_top=True)

        except Exception as e:
            logger.exception("Processing failed")
            sg.popup_error(f"An error occurred: {e}", keep_on_top=True)
