import json
import re
import os
import threading
import queue
from flask import Flask, request, send_file, jsonify
from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier
from flask_cors import CORS



def pretty_print_json(data):
    print(json.dumps(data, indent=4, ensure_ascii=False))



app = Flask(__name__)
CORS(app)

# Thread-safe queue to store incoming requests
request_queue = queue.Queue()
lock = threading.Lock()


def process_request(data):
    # Retrieve URL from the data
    url = data.get('url')
    if not url:
        return {"error": "URL not provided", "status_code": 400}

    # Extract Project ID using regex
    project_id = None
    pattern = r"projects/([a-f0-9-]{36})"
    match = re.search(pattern, url)
    if match:
        project_id = match.group(1)
    else:
        return {"error": "Project ID not found in the URL", "status_code": 400}

    # Detect the section in the URL
    section = None
    if "/budget" in url:
        section = "Budgets"
    elif "/cost/cost" in url:
        section = "Costs"
    elif "/forms" in url:
        section = "Forms"
    else:
        return {"error": "Unrecognized section in URL", "status_code": 400}

    # Fetch data based on the section
    acc_api = ACCAPI()
    try:
        if section == "Budgets":
            response = acc_api.call_api(f"cost/v1/containers/{project_id}/budgets")["results"]
        elif section == "Costs":
            response = acc_api.call_api(f"cost/v1/containers/{project_id}/contracts")["results"]
        elif section == "Forms":
            response = acc_api.call_api(f"construction/forms/v1/projects/{project_id}/forms")["data"]
    except Exception as e:
        return {"error": f"Failed to fetch data: {str(e)}", "status_code": 500}

    # Create and modify Excel file based on the section data
    excel_modifier = ExcelModifier(template_filename="templates/template.xlsx", modified_folder="modified_files")
    try:
        excel_modifier.open_workbook()

        # Add headers and data to the Excel file based on the section
        if section == "Budgets":
            headers = ["Formatted Code", "Unit Price", "Original Amount"]
            # for col, header in enumerate(headers, start=1):
            #     excel_modifier.modify_cell(f"{chr(64 + col)}1", header)
            
            pretty_print_json(response)

            for i, budget in enumerate(response, start=11):
                excel_modifier.modify_cell(f'D{i}', budget['unitPrice'])
                

        elif section == "Costs":
            from TEST import print_cost_cover
            pdf_path = print_cost_cover()
            return {"pdf_path": pdf_path, "status_code": 200}

        elif section == "Forms":
            headers = ["Form Id", "Form Name", "Status"]
            for col, header in enumerate(headers, start=1):
                excel_modifier.modify_cell(f"{chr(64 + col)}1", header)

            for i, form in enumerate(response, start=2):
                excel_modifier.modify_cell(f'A{i}', form['id'])
                excel_modifier.modify_cell(f'B{i}', form['name'])
                excel_modifier.modify_cell(f'C{i}', form['status'])

        # excel_modifier.auto_fit_columns()
        # excel_modifier.add_gridlines()

        # Save Excel file and export to PDF
        excel_modifier.save_workbook(filename='output.xlsx')
        pdf_path = excel_modifier.export_to_pdf(filename='output.pdf')

        # Return the generated PDF path
        return {"pdf_path": pdf_path, "status_code": 200}

    finally:
        excel_modifier.close_workbook()


def worker():
    """Background thread that processes requests from the queue."""
    while True:
        request_data, response_queue = request_queue.get()  # Get a request from the queue
        try:
            response = process_request(request_data)
            response_queue.put(response)  # Send the response back to the main thread
        except Exception as e:
            response_queue.put({"error": str(e), "status_code": 500})
        finally:
            request_queue.task_done()


@app.route('/generate-pdf', methods=['POST'])
def generate_pdf():
    data = request.get_json()

    # Response queue to get the result from the background thread
    response_queue = queue.Queue()

    # Add the request to the queue
    request_queue.put((data, response_queue))

    # Wait for the response from the worker
    response = response_queue.get()

    # Send the PDF file if processing is successful
    if "pdf_path" in response:
        pdf_path = response["pdf_path"]
        if os.path.exists(pdf_path):
            return send_file(pdf_path, as_attachment=True, download_name="output.pdf", mimetype="application/pdf")
        else:
            return jsonify({"error": "PDF generation failed."}), 500
    else:
        return jsonify({"error": response.get("error", "Unknown error")}), response.get("status_code", 500)


# Start the worker thread
threading.Thread(target=worker, daemon=True).start()

if __name__ == '__main__':
    app.run(debug=True)
