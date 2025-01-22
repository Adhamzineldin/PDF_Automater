import xlwings as xw
import os

class ExcelModifier:
    def __init__(self, template_filename, modified_folder):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, template_filename)
        self.modified_folder = os.path.join(self.script_dir, modified_folder)


        if not os.path.exists(self.modified_folder):
            os.makedirs(self.modified_folder)

        self.app = None
        self.workbook = None
        self.sheet = None

    def open_workbook(self):
        """Opens the Excel workbook and initializes the sheet."""
        self.app = xw.App(visible=False)
        self.workbook = self.app.books.open(self.excel_path)
        self.sheet = self.workbook.sheets[0]

    def modify_cell(self, cell_range, value):
        """Modifies a specific cell range with a new value."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
        self.sheet.range(cell_range).value = value
        print(f"Cell {cell_range} updated to {value}")

    def save_workbook(self, filename='modified.xlsx'):
        """Saves the workbook with a new name."""
        if self.workbook is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
        save_path = os.path.join(self.modified_folder, filename)
        self.workbook.save(save_path)
        print(f"Workbook saved at {save_path}")
        return save_path

    def export_to_pdf(self, filename='modified.pdf'):
        """Exports the sheet to a PDF."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
        pdf_path = os.path.join(self.modified_folder, filename)
        self.sheet.api.ExportAsFixedFormat(0, pdf_path)  # 0 refers to xlTypePDF
        print(f"PDF exported at {pdf_path}")
        return pdf_path

    def close_workbook(self):
        """Closes the workbook and Excel application."""
        if self.workbook:
            self.workbook.close()
        if self.app:
            self.app.quit()
