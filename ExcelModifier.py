import os
import subprocess
from tempfile import NamedTemporaryFile
from svgpathtools import svg2paths
from PIL import Image, ImageDraw
from openpyxl import load_workbook
from openpyxl.styles import Border, Side
from openpyxl.drawing.image import Image as ExcelImage
from openpyxl.utils import get_column_letter

class ExcelModifier:
    def __init__(self, template_filename, modified_folder):
        self.script_dir = os.path.dirname(os.path.abspath(__file__))
        self.excel_path = os.path.join(self.script_dir, template_filename)
        self.modified_folder = os.path.join(self.script_dir, modified_folder)

        if not os.path.exists(self.modified_folder):
            os.makedirs(self.modified_folder)

        self.wb = load_workbook(self.excel_path)
        self.sheet = self.wb.active

    def modify_cell(self, cell_range, value):
        """Modifies a specific cell with a new value"""
        self.sheet[cell_range] = value
        print(f"Cell {cell_range} updated to {value}.")

    def auto_fit_columns(self):
        """Approximates auto-fit for columns"""
        for column in self.sheet.columns:
            max_length = 0
            column_letter = get_column_letter(column[0].column)

            for cell in column:
                try:
                    value_length = len(str(cell.value))
                    if value_length > max_length:
                        max_length = value_length
                except:
                    pass

            adjusted_width = (max_length + 2) * 1.2
            self.sheet.column_dimensions[column_letter].width = adjusted_width

        print("Columns adjusted for content.")

    def add_gridlines(self):
        """Adds gridlines using cell borders"""
        thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
        )

        for row in self.sheet.iter_rows():
            for cell in row:
                cell.border = thin_border
        print("Gridlines added using borders.")

    def insert_svg_as_image(self, svg_code, cell_range):
        """Inserts SVG image converted to PNG into the worksheet"""
        try:
            # Create temporary files
            with NamedTemporaryFile(delete=False, suffix='.svg', dir=self.modified_folder) as temp_svg:
                temp_svg.write(svg_code.encode('utf-8'))
                temp_svg_path = temp_svg.name

            # Convert SVG to PNG
            paths, _ = svg2paths(temp_svg_path)
            img_size = (600, 300)  # Default image size
            img = Image.new('RGBA', img_size, (255, 255, 255, 0))
            draw = ImageDraw.Draw(img)

            # Scale and draw paths
            for path in paths:
                for segment in path:
                    start = segment.start
                    end = segment.end
                    draw.line(
                            (start.real, start.imag, end.real, end.imag),
                            fill='black',
                            width=2
                    )

            # Save PNG
            with NamedTemporaryFile(delete=False, suffix='.png', dir=self.modified_folder) as temp_png:
                img.save(temp_png.name, "PNG")
                temp_png_path = temp_png.name

            # Insert into Excel
            img = ExcelImage(temp_png_path)
            self.sheet.add_image(img, cell_range)
            print(f"Image inserted at {cell_range}")

            # Clean up temporary files
            os.unlink(temp_svg_path)
            os.unlink(temp_png_path)

        except Exception as e:
            print(f"Error processing SVG: {str(e)}")

    def save_workbook(self, filename='modified.xlsx'):
        """Saves the modified workbook"""
        save_path = os.path.join(self.modified_folder, filename)
        self.wb.save(save_path)
        print(f"Workbook saved to {save_path}")
        return save_path

    def export_to_pdf(self, filename='modified.pdf'):
        """Converts the Excel file to PDF using LibreOffice"""
        xlsx_path = self.save_workbook()
        pdf_path = os.path.join(self.modified_folder, filename)

        try:
            # Convert using LibreOffice
            command = [
                    'libreoffice',
                    '--headless',
                    '--convert-to', 'pdf',
                    '--outdir', self.modified_folder,
                    xlsx_path
            ]

            result = subprocess.run(
                    command,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    check=True
            )

            print(f"PDF successfully created at {pdf_path}")
            return pdf_path

        except subprocess.CalledProcessError as e:
            print(f"PDF conversion failed: {e.stderr.decode()}")
            return None

    def close_workbook(self):
        """Closes the workbook (not necessary with openpyxl, but included for compatibility)"""
        pass
    
    def open_workbook(self):
        # """Opens the workbook"""
        # self.wb = load_workbook(self.excel_path)
        # self.sheet = self.wb.active
        pass
    
# Example usage
if __name__ == "__main__":
    svg_example = """<svg viewBox="0 0 100 100" xmlns="http://www.w3.org/2000/svg">
        <circle cx="50" cy="50" r="40" stroke="black" fill="none"/>
    </svg>"""

    modifier = ExcelModifier('template.xlsx', 'modified')

    # Basic operations
    modifier.modify_cell('A1', 'Linux-Compatible Report')
    modifier.modify_cell('B2', 12345)
    modifier.auto_fit_columns()
    modifier.add_gridlines()

    # Image insertion
    modifier.insert_svg_as_image(svg_example, 'C5')

    # Save and export
    modifier.save_workbook()
    modifier.export_to_pdf()
    modifier.close_workbook()