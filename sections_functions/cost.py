import json
import os
import re
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs

from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier


def pretty_print_json(data):
    print(json.dumps(data, indent=4, ensure_ascii=False))

def extract_cost_id(url):
    # Parse the URL
    parsed_url = urlparse(url)

    # Check if 'preview' exists in the query parameters
    query_params = parse_qs(parsed_url.query)
    if 'preview' in query_params:
        return query_params['preview'][0]
    elif 'selectId' in query_params:
        return query_params['selectId'][0]

    # Extract the last segment of the path
    last_segment = parsed_url.path.rstrip('/').split('/')[-1]

    # Validate if it's a UUID
    if re.fullmatch(r'[a-f0-9-]{36}', last_segment):
        return last_segment

    return None


def print_cost_cover(project_id, url):
    acc_api = ACCAPI()

    cost_payment_response = acc_api.call_api(f"cost/v1/containers/{project_id}/payments")["results"]
    change_order_response = acc_api.call_api(f"cost/v1/containers/{project_id}/cost-items")["results"]

    # Initialize variables
    current_date = datetime.now()
    
    # Keep checking previous months if no payments found
    cost_payments = []
    
    
    cost_id = extract_cost_id(url)
    
    
    if cost_id:
        cost_payments = [
                cost_payment for cost_payment in cost_payment_response
                if cost_payment["associationType"] == "Contract"
                   and cost_payment["id"] == cost_id
        ]
    else:
        while not cost_payments:
            cost_payments = [
                    cost_payment for cost_payment in cost_payment_response
                    if cost_payment["associationType"] == "Contract"
                       and datetime.strptime(cost_payment["endDate"], "%Y-%m-%d").strftime("%Y-%m") == current_date.strftime("%Y-%m")
            ]
            # Move to the previous month if no results
            if not cost_payments:
                current_date = current_date.replace(day=1) - timedelta(days=1)
    
    print(f"cost id is {cost_id}")
    print(len(cost_payments))
    change_orders = [change_order for change_order in change_order_response if change_order["contractId"] in [cost_payment["associationId"] for cost_payment in cost_payments]]
    print(change_orders)
    
    print("Change payment response:")
    pretty_print_json(change_order_response)
    
    for payment in cost_payments:
        association_Id = payment["associationId"]
        payment_number = payment["id"]
        items = [item for item in change_orders if item["contractId"] == association_Id]
        similar_item = 0
        new_item = 0
        for item in items:
            if item["name"] == "Block Work":
                if item["estimated"]:
                    similar_item += float(item["estimated"])
            else:
                if item["estimated"]:
                    new_item += float(item["estimated"])

        # Determine the template path based on association ID
        template_filename = f"{payment_number}.xlsx"
        template_path = os.path.join("modified_files", template_filename)
        new = True
        if os.path.exists(template_path):
            new = False
            selected_template = template_path
        else:
            selected_template = "templates/cost_cover_template.xlsx"

        excel_modifier = ExcelModifier(template_filename=selected_template, modified_folder="modified_files")
        try:
            excel_modifier.open_workbook()
            print("Payment:")
            pretty_print_json(payment)
            
            if new:
                excel_modifier.modify_cell("D10", float(payment["originalAmount"]))
                excel_modifier.modify_cell("D13", new_item)
                excel_modifier.modify_cell("D14", similar_item)
                excel_modifier.modify_cell("D15", float(payment["amount"]))
                payment["status"] = "Main-Contractor"
            elif payment["status"] == "revise" or payment["status"] == "inReview":
                excel_modifier.modify_cell("E10", float(payment["originalAmount"]))
                excel_modifier.modify_cell("E13", new_item)
                excel_modifier.modify_cell("E14", similar_item)
                excel_modifier.modify_cell("E15", float(payment["amount"]))
                payment["status"] = "Consultant"
            elif payment["status"] == "accepted" or payment["status"] == "approved":
                excel_modifier.modify_cell("F10", float(payment["originalAmount"]))
                excel_modifier.modify_cell("F13", new_item)
                excel_modifier.modify_cell("F14", similar_item)
                excel_modifier.modify_cell("F15", float(payment["amount"]))
                payment["status"] = "Owner"
            else:
                excel_modifier.modify_cell("D10", float(payment["originalAmount"]))
                excel_modifier.modify_cell("D13", new_item)
                excel_modifier.modify_cell("D14", similar_item)
                excel_modifier.modify_cell("D15", float(payment["amount"]))
                payment["status"] = "Main-Contractor"
            
            

            excel_modifier.save_workbook(filename=f'{payment_number}.xlsx')
            project = acc_api.call_api(f"construction/admin/v1/projects/{project_id}")
            pdf_path = excel_modifier.export_to_pdf(payment, filename='output.pdf', excel_filename=payment_number, project_name=project["name"])

            # Return the generated PDF path
            return pdf_path
        except Exception as e:
            print(f"Failed to modify Excel file: {str(e)}")
        finally:
            excel_modifier.close_workbook()