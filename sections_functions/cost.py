import json
from datetime import datetime

from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier


def pretty_print_json(data):
    print(json.dumps(data, indent=4, ensure_ascii=False))

def print_cost_cover(project_id):
    acc_api = ACCAPI()

    cost_payment_response = acc_api.call_api(f"cost/v1/containers/{project_id}/payments")["results"]
    change_order_response = acc_api.call_api(f"cost/v1/containers/{project_id}/cost-items")["results"]


    cost_payments = [
            cost_payment for cost_payment in cost_payment_response
            if cost_payment["associationType"] == "Contract"
               and datetime.strptime(cost_payment["endDate"], "%Y-%m-%d").strftime("%Y-%m") == datetime.now().strftime("%Y-%m")
    ]

    change_order = [ change_order for change_order in change_order_response if change_order["contractId"] in [ cost_payment["associationId"] for cost_payment in cost_payments ] ]

    for payment in cost_payments:
        association_Id = payment["associationId"]
        items = [ item for item in change_order if item["contractId"] == association_Id ]
        similar_item = 0
        new_item = 0
        for item in items:
            if item["name"] == "Block Work":
                similar_item += float(item["estimated"])
            else:
                new_item += float(item["estimated"])


        excel_modifier = ExcelModifier(template_filename="templates/cost_cover_template.xlsx", modified_folder="modified_files")
        try:
            excel_modifier.open_workbook()
            pretty_print_json(payment)
            excel_modifier.modify_cell("D10", float(payment["originalAmount"]))
            excel_modifier.modify_cell("D13", new_item)
            excel_modifier.modify_cell("D14", similar_item)
            # excel_modifier.modify_cell("D14", similar_item)

            excel_modifier.save_workbook(filename='output.xlsx')
            pdf_path = excel_modifier.export_to_pdf(filename='output.pdf')


            # Return the generated PDF path
            return pdf_path

        finally:
            excel_modifier.close_workbook()




