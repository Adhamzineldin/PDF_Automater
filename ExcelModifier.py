import xlwings as xw
import os
from svgpathtools import svg2paths
from PIL import Image, ImageDraw


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
        self.app = xw.App(visible=False, add_book=False)
        self.workbook = self.app.books.open(self.excel_path)
        self.sheet = self.workbook.sheets[0]

    def modify_cell(self, cell_range, value):
        """Modifies a specific cell range with a new value and optional formatting."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
    
        cell = self.sheet.range(cell_range)
        cell.value = value
    
        print(f"Cell {cell_range} updated to {value}.")

    def auto_fit_columns(self):
        """Automatically adjusts all columns to fit content."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
    
        self.sheet.api.Columns.AutoFit()
        print("Auto-fit applied to all columns.")
    
    def add_gridlines(self):
        """Adds gridlines to the sheet."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
    
        # Use borders to simulate gridlines
        for border_id in range(7, 13):  # Borders IDs for Excel
            self.sheet.api.Cells.Borders(border_id).LineStyle = 1  # xlContinuous
    
        print("Gridlines added to the sheet.")

    def save_workbook(self, filename='modified.xlsx'):
        """Saves the workbook with a new name."""
        if self.workbook is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
        save_path = os.path.join(self.modified_folder, filename)
        self.workbook.save(save_path)
        print(f"Workbook saved at {save_path}")
        return save_path

    def export_to_pdf(self, filename='modified.pdf'):
        """Exports the sheet to a PDF, fitting it to a single page."""
        if self.sheet is None:
            raise Exception("Workbook is not opened. Call open_workbook() first.")
    
        # Get the ActiveSheet object from the sheet's API
        sheet_api = self.sheet.api
    
        # Set the page setup to fit the sheet to one page
        sheet_api.PageSetup.FitToPagesWide = 1  # Fit the sheet to one page wide
        sheet_api.PageSetup.FitToPagesTall = 1  # Fit the sheet to one page tall
    
        # Ensure that the sheet does not use zoom and scales automatically
        sheet_api.PageSetup.Zoom = False  # Disable zoom, let FitToPages take control
    
        # Optionally set the print area if necessary (uncomment and adjust if needed)
        # sheet_api.PageSetup.PrintArea = "A1:Z100"  # Adjust the print area if required
    
        # Define the path for saving the PDF
        pdf_path = os.path.join(self.modified_folder, filename)
    
        # Export as PDF
        try:
            sheet_api.ExportAsFixedFormat(0, pdf_path)  # 0 refers to xlTypePDF
            print(f"PDF exported at {pdf_path}")
        except Exception as e:
            print(f"Error exporting to PDF: {e}")
            return None
    
        return pdf_path

    def close_workbook(self):
        """Closes the workbook and Excel application."""
        if self.workbook:
            self.workbook.close()
        if self.app:
            self.app.quit()





    def insert_svg_as_image(self, svg_code, cell_range):
        """
        Converts SVG code to PNG using svgpathtools and Pillow, and inserts it into the Excel sheet.
        
        Parameters:
        - svg_code: str, SVG code as a string.
        - cell_range: str, Excel cell range where the image should be inserted.
        """
        try:
            # Step 1: Save the SVG code to a temporary file
            temp_svg_path = os.path.join(self.modified_folder, "temp_image.svg")
            with open(temp_svg_path, "w", encoding="utf-8") as svg_file:
                svg_file.write(svg_code)
    
            # Step 2: Parse the SVG to extract paths
            paths, attributes = svg2paths(temp_svg_path)
    
            # Step 3: Create a new blank image (white background)
            width, height = 600, 300  # You can adjust the size as needed
            img = Image.new('RGBA', (width, height), (0, 0, 0, 0))  # Transparent background
            draw = ImageDraw.Draw(img)
    
            # Step 4: Draw the paths onto the image
            for path in paths:
                for segment in path:
                    start = segment.start
                    end = segment.end
                    draw.line((start.real, start.imag, end.real, end.imag), fill='black', width=2)
    
            # Step 5: Save the image as PNG
            temp_png_path = os.path.join(self.modified_folder, "temp_image.png")
            img.save(temp_png_path, "PNG")
    
            # Step 6: Insert the PNG into the Excel sheet
            self.sheet.pictures.add(temp_png_path,
                                    left=self.sheet.range(cell_range).left,
                                    top=self.sheet.range(cell_range).top)
    
            print(f"SVG inserted as image at {cell_range}")
        except Exception as e:
            print(f"An error occurred while processing the SVG: {e}")


