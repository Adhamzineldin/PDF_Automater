import re

from flask import Flask, request, send_file
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

    # Determine the section based on URL content
    section = None
    if "/budget" in url:
        section = "Budgets"
    elif "/cost/cost" in url:
        section = "Costs"
    elif "/forms" in url:
        section = "Forms"
    else:
        return {"error": "Unrecognized section in URL"}, 400

    print(f"Detected Section: {section}, URL: {url}")

    # Fetch data based on the detected section
    acc_api = ACCAPI()
    try:
        if section == "Budgets":
            response = acc_api.call_api(f"cost/v1/containers/{project_id}/budgets")["results"]
            print(response)
        elif section == "Costs":
            response = acc_api.call_api(f"cost/v1/containers/{project_id}/contracts")["results"]
            print(response)
        elif section == "Forms":
            response = acc_api.call_api(f"construction/forms/v1/projects/{project_id}/forms")["data"]
            print(response)
        else:
            return {"error": "Invalid section logic"}, 500
    except Exception as e:
        return {"error": f"Failed to fetch data: {str(e)}"}, 500

    # Create and modify the Excel file
    excel_modifier = ExcelModifier(template_filename="../templates/template.xlsx", modified_folder="modified_files")


    try:
        excel_modifier.open_workbook()

        # Add headers for each section
        if section == "Budgets":
            headers = ["Formatted Code", "Unit Price", "Original Amount"]
            for col, header in enumerate(headers, start=1):
                excel_modifier.modify_cell(f"{chr(64 + col)}1", header)
    
            # Add data rows
            for i, budget in enumerate(response, start=2):
                excel_modifier.modify_cell(f'A{i}', budget['formattedCode'])
                excel_modifier.modify_cell(f'B{i}', budget['unitPrice'])
                excel_modifier.modify_cell(f'C{i}', budget['originalAmount'])
    
        elif section == "Costs":
            headers = ["Cost Code", "Type", "Amount"]
            for col, header in enumerate(headers, start=1):
                excel_modifier.modify_cell(f"{chr(64 + col)}1", header)
    
            # Add data rows
            for i, cost in enumerate(response, start=2):
                excel_modifier.modify_cell(f'A{i}', cost['code'])
                excel_modifier.modify_cell(f'B{i}', cost['type'])
                excel_modifier.modify_cell(f'C{i}', cost['allocatedAmount'])
    
        elif section == "Forms":
            headers = ["Form Id", "Form Name", "Status"]
            for col, header in enumerate(headers, start=1):
                excel_modifier.modify_cell(f"{chr(64 + col)}1", header)
    
            # Add data rows
            for i, form in enumerate(response, start=2):
                excel_modifier.modify_cell(f'A{i}', form['id'])
                excel_modifier.modify_cell(f'B{i}', form['name'])
                excel_modifier.modify_cell(f'C{i}', form['status'])
                
    
        # Apply formatting to fit content and look professional
        excel_modifier.auto_fit_columns()  # Automatically adjust column widths
        excel_modifier.add_gridlines()  # Add gridlines for better visibility
    
        excel_modifier.save_workbook(filename='output.xlsx')
        pdf_path = excel_modifier.export_to_pdf(filename='output.pdf')
    finally:
        excel_modifier.close_workbook()

# Return the PDF as a downloadable file
    return send_file(pdf_path, as_attachment=True, download_name="output.pdf", mimetype="application/pdf")

if __name__ == '__main__':
    app.run(debug=True)
