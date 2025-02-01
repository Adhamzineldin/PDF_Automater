import json
import os
from datetime import datetime, timedelta

from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier


def pretty_print_json(data):
    print(json.dumps(data, indent=4, ensure_ascii=False))

def print_cost_cover(project_id):
    acc_api = ACCAPI()

    cost_payment_response = acc_api.call_api(f"cost/v1/containers/{project_id}/payments")["results"]
    change_order_response = acc_api.call_api(f"cost/v1/containers/{project_id}/cost-items")["results"]

    # Initialize variables
    current_date = datetime.now()
    
    # Keep checking previous months if no payments found
    cost_payments = []
    while not cost_payments:
        cost_payments = [
                cost_payment for cost_payment in cost_payment_response
                if cost_payment["associationType"] == "Contract"
                   and datetime.strptime(cost_payment["endDate"], "%Y-%m-%d").strftime("%Y-%m") == current_date.strftime("%Y-%m")
        ]
        # Move to the previous month if no results
        if not cost_payments:
            current_date = current_date.replace(day=1) - timedelta(days=1)
    print(len(cost_payments))
    change_orders = [change_order for change_order in change_order_response if change_order["contractId"] in [cost_payment["associationId"] for cost_payment in cost_payments]]
    print(change_orders)
    
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
            pretty_print_json(payment)
            if new:
                excel_modifier.modify_cell("D10", payment["originalAmount"])
                excel_modifier.modify_cell("D13", new_item)
                excel_modifier.modify_cell("D14", similar_item)
                excel_modifier.modify_cell("D15", payment["amount"])
            elif payment["status"] == "revise" or payment["status"] == "inReview":
                excel_modifier.modify_cell("E10", payment["originalAmount"])
                excel_modifier.modify_cell("E13", new_item)
                excel_modifier.modify_cell("E14", similar_item)
                excel_modifier.modify_cell("E15", payment["amount"])
            elif payment["status"] == "accepted" or payment["status"] == "approved":
                excel_modifier.modify_cell("F10", payment["originalAmount"])
                excel_modifier.modify_cell("F13", new_item)
                excel_modifier.modify_cell("F14", similar_item)
                excel_modifier.modify_cell("F15", payment["amount"])
            else:
                excel_modifier.modify_cell("D10", payment["originalAmount"])
                excel_modifier.modify_cell("D13", new_item)
                excel_modifier.modify_cell("D14", similar_item)
                excel_modifier.modify_cell("D15", payment["amount"])
                
                
                
            

            excel_modifier.save_workbook(filename=f'{payment_number}.xlsx')
            pdf_path = excel_modifier.export_to_pdf(filename='output.pdf', excel_filename=payment_number)

            # Return the generated PDF path
            return pdf_path
        except Exception as e:
            print(f"Failed to modify Excel file: {str(e)}")
        finally:
            excel_modifier.close_workbook()