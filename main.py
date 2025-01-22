from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier

def main():
    # Specify the template file and output folder
    template_file = './template.xlsx'
    output_folder = 'Modified_Files'

    # Create an instance of the ExcelModifier class
    modifier = ExcelModifier(template_file, output_folder)

    
    try:
        # Instantiate the ACCAPI class
        acc_api = ACCAPI()

        # Dynamic API call example
        endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms"
        result = acc_api.call_api(endpoint)
        print("API Response:", result)
        
        print("-----------------------------------------------------------------------------------------------")
        
        endpoint = f"/cost/v1/containers/{acc_api.CONTAINER_ID}/schedule-of-values"
        result = acc_api.call_api(endpoint)
        print("API Response:", result)
        


      
        # modifier.open_workbook()
        # 
        # modifier.modify_cell('E12', 1000000)
        # modifier.modify_cell('E13', 2501.2501)
        # modifier.modify_cell('F22', 'Sample Text')
        # modifier.modify_cell('G31', 420000)
        # 
        # modifier.save_workbook()
        # 
        # modifier.export_to_pdf()
        # 

    except EnvironmentError as env_err:
        print(f"Environment error: {env_err}")
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Ensure the workbook is closed
        modifier.close_workbook()
    

    

       

    

if __name__ == '__main__':
    main()
