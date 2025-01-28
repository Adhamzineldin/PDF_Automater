from flask import Flask, request, send_file
import re
from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier
from flask_cors import CORS
app = Flask(__name__)
CORS(app)

@app.route('/generate-pdf', methods=['POST', 'OPTIONS'])
def generate_pdf():

    if request.method == 'OPTIONS':
        return '', 200  # Respond to preflight request
    
    # Retrieve the URL from the POST request
    data = request.get_json()
    url = data.get('url')

    if not url:
        return {"error": "URL not provided"}, 400

    project_id = None
    pattern = r"projects/([a-f0-9-]{36})"

    # Search for the project ID using regex
    match = re.search(pattern, url)

    if match:
        project_id = match.group(1)
        print("Project ID:", project_id)
    else:
        return {"error": "Project ID not found in the URL"}, 400
    
    # Call API to fetch budgets
    acc_api = ACCAPI()
    budgets = acc_api.call_api(f"cost/v1/containers/{project_id}/budgets")["results"]


    # Create an ExcelModifier instance
    excel_modifier = ExcelModifier(template_filename="templates/template.xlsx", modified_folder="modified_files")
    
    try:
        # Open the Excel workbook
        excel_modifier.open_workbook()
    
        # Modify the Excel workbook with the fetched budgets
        # excel_modifier.modify_cell('A1', "Budget Code")
        # excel_modifier.modify_cell('B1', "Budget Name")
        # excel_modifier.modify_cell('C1', "Original Budget")
    
        for i, budget in enumerate(budgets):
            excel_modifier.modify_cell(f'A{i+2}', budget['formattedCode'])
            excel_modifier.modify_cell(f'B{i+2}', budget['name'])
            excel_modifier.modify_cell(f'C{i+2}', budget['originalAmount'])
            
        excel_modifier.save_workbook(filename='output.xlsx')
        # Export the modified Excel sheet to PDF
        pdf_path = excel_modifier.export_to_pdf(filename='output.pdf')
    finally:
        excel_modifier.close_workbook()

    

    # Return the PDF as a downloadable response
    return send_file(pdf_path, as_attachment=True, download_name="output.pdf", mimetype="application/pdf")

if __name__ == '__main__':
    app.run(debug=True)
