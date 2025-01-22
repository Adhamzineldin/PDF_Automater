from ExcelModifier import ExcelModifier

def main():
    # Specify the template file and output folder
    template_file = '../PDF Automater/template.xlsx'
    output_folder = 'Modified_Files'

    # Create an instance of the ExcelModifier class
    modifier = ExcelModifier(template_file, output_folder)

    try:
        # Open the workbook
        modifier.open_workbook()

        # Modify multiple cells
        modifier.modify_cell('E12', 1000000)
        modifier.modify_cell('E13', 2501.2501)
        modifier.modify_cell('F22', 'Sample Text')
        modifier.modify_cell('G31', 420000)

        # Save the modified workbook
        modifier.save_workbook()

        # Export to PDF
        modifier.export_to_pdf()

    finally:
        # Ensure the workbook is closed
        modifier.close_workbook()

if __name__ == '__main__':
    main()
