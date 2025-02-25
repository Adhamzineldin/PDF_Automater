import calendar
import json
import os
import re
from datetime import datetime, timedelta
from urllib.parse import urlparse, parse_qs

from xlwings.mac_dict import properties

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



def modify_cell_with_null_check(excel_modifier, letter, cell, value):
    print(f"typeof value is {type(value)}")
    if value:
        excel_modifier.modify_cell(f"{letter}{cell}", float(value))
    else:
        excel_modifier.modify_cell(f"{letter}{cell}", 0)

    



def print_cost_cover(project_id, url):
    acc_api = ACCAPI()

    cost_payment_response = acc_api.call_api(f"cost/v1/containers/{project_id}/payments")["results"]
    change_order_response = acc_api.call_api(f"cost/v1/containers/{project_id}/cost-items")["results"]
    sov_response = acc_api.call_api(f"cost/v1/containers/{project_id}/schedule-of-values")["results"]
    # pretty_print_json(sov_response)

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
                    
        inflation_rate = sum([item for item in items if item["type"] == "INF"])
        remeasured = sum([item for item in items if item["type"] == "REM"])
        

        # Determine the template path based on association ID
        template_filename = f"{payment_number}.xlsx"
        template_path = os.path.join("modified_files", template_filename)
        new = True
        if os.path.exists(template_path):
            new = False
            selected_template = template_path
        else:
            selected_template = "templates/cost_cover_template.xlsx"


        sovs = [sov for sov in sov_response if sov["contractId"] == association_Id]
        
        project_mobilization = [sov for sov in sovs if sov["name"] == "Project Mobilization"]
        if project_mobilization:
            project_mobilization = project_mobilization[0]
        else:
            project_mobilization = {"amount": 0}
        excel_modifier = ExcelModifier(template_filename=selected_template, modified_folder="modified_files")
        try:
            excel_modifier.open_workbook()
            print("Payment:")
            pretty_print_json(payment)
            letter = "D"
            
            if new:
                letter = "D"
                payment["status"] = "Main-Contractor"
            elif payment["status"] == "revise" or payment["status"] == "inReview":
                letter = "E"
                payment["status"] = "Consultant"
            elif payment["status"] == "accepted" or payment["status"] == "approved":
                letter = "F"
                payment["status"] = "Owner"
            else:
                letter = "D"
                payment["status"] = "Main-Contractor"


            print("--------------------------------TEST----------------------------------------------")
            payment_items = acc_api.call_api(f"cost/v1/containers/{project_id}/payment-items?filter[paymentId]={payment_number}")["results"]
            project_mobilization = [item for item in payment_items if item["number"] in ["01-71", "01-72"]]
            project_mobilization = sum([float(item["amount"]) for item in project_mobilization])
            
            properties = payment["properties"]

            property_000 = next(iter([p for p in properties if "000" in p["name"]]), None)
            property_001 = next(iter([p for p in properties if "001" in p["name"]]), None)
            property_002 = next(iter([p for p in properties if "002" in p["name"]]), None)
            property_003 = next(iter([p for p in properties if "003" in p["name"]]), None)
            property_004 = next(iter([p for p in properties if "004" in p["name"]]), None)
            property_005 = next(iter([p for p in properties if "005" in p["name"]]), None)



            current_date = datetime.now().strftime("%m/%d/%y")
            today = datetime.now()
            last_day = calendar.monthrange(today.year, today.month)[1]
            last_date = f"{today.month:02}/{last_day:02}/{str(today.year)[-2:]}" 
            

            excel_modifier.modify_cell("C7", current_date)
            excel_modifier.modify_cell("F6", last_date)
            modify_cell_with_null_check(excel_modifier, letter, "10", payment["originalAmount"])
            modify_cell_with_null_check(excel_modifier, letter, "13", new_item)
            modify_cell_with_null_check(excel_modifier, letter, "14", similar_item)
            modify_cell_with_null_check(excel_modifier, letter, "15", remeasured)
            modify_cell_with_null_check(excel_modifier, letter, "16", inflation_rate)
            modify_cell_with_null_check(excel_modifier, letter, "20", payment["amount"])
            modify_cell_with_null_check(excel_modifier, letter, "23", property_004["value"])
            modify_cell_with_null_check(excel_modifier, letter, "24", property_005["value"])
            modify_cell_with_null_check(excel_modifier, letter, "25", project_mobilization)
            modify_cell_with_null_check(excel_modifier, letter, "28", payment["materials"])
            modify_cell_with_null_check(excel_modifier, letter, "37", property_000["value"])
            modify_cell_with_null_check(excel_modifier, letter, "38", property_001["value"])
            modify_cell_with_null_check(excel_modifier, letter, "39", property_002["value"])
            modify_cell_with_null_check(excel_modifier, letter, "40", property_003["value"])
           
            
            


            
            print(f"Payment Number: {payment_number}")
            excel_modifier.save_workbook(filename=f'{payment_number}.xlsx')
            try:
                project = acc_api.call_api(f"construction/admin/v1/projects/{project_id}")
            except Exception:
                project =  None
                print("Failed to fetch project name PROP PERMISSION ISSUE")
            pdf_path = excel_modifier.export_to_pdf(payment, filename='output.pdf', excel_filename=payment_number)
            if project:
                acc_api.upload_pdf_to_acc(pdf_path=pdf_path, filename='output.pdf', project_name=project["name"], folder_name="Cost Cover Sheets")
            else:
                print("Failed to upload PDF to ACC Because no project name")
            
            print(f"FINAL PDF PATH: {pdf_path}")
            return pdf_path
        except Exception as e:
            print(f"Failed to modify Excel file: {str(e)}")
        finally:
            excel_modifier.close_workbook()