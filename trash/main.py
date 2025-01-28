import sys

import requests

from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier
import json




def pretty_print_json(data):
    print(json.dumps(data, indent=4, ensure_ascii=False))

def main():

    sys.stdout.reconfigure(encoding='utf-8')
    
    # Specify the template file and output folder
    template_file = 'templates/template.xlsx'
    output_folder = 'Modified_Files'

    # Create an instance of the ExcelModifier class
    modifier = ExcelModifier(template_file, output_folder)

    
    try:
        # Instantiate the ACCAPI class
        acc_api = ACCAPI()

        # Dynamic API call example
        endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms"
        result = acc_api.call_api(endpoint)
        data = result['data']
        
        
        for form in data:
            if form["formNum"] == 30:
                pretty_print_json(form)
        
      
        divider = "=" * 60
        
        
        
        
        
 
        
        
        


    except EnvironmentError as env_err:
        print(f"Environment error: {env_err}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure the workbook is closed
        modifier.close_workbook()
    

    

       

    

if __name__ == '__main__':
    main()
