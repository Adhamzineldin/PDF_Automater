from datetime import datetime
from ACCAPI import ACCAPI
from ExcelModifier import ExcelModifier
import pandas as pd


def generate_smart_form():

    pd.set_option('display.max_rows', None)  # Show all rows
    pd.set_option('display.max_columns', None)  # Show all columns
    pd.set_option('display.width', None)  # Auto-detect width
    pd.set_option('display.max_colwidth', None)  # Show full column contents
    
    try:
        #     # Open the workbook
        # Instantiate the ACCAPI class
        acc_api = ACCAPI()

        # Get the current month and year
        today = datetime.today()
        year = today.year
        month = today.month

        # Calculate the previous month and adjust year if needed
        previous_month = month - 1 if month > 1 else 12
        previous_year = year if month > 1 else year - 1

        # Convert to string format
        previous_month_abbr = datetime(previous_year, previous_month, 1).strftime("%b")  # Abbreviated month name

        # Dynamic API call example
        endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms?formDateMin={datetime.now().strftime("%b")} 6, {datetime.now().year}&templateId=3c852249-47f5-4234-987a-06b2717ba550"
        result = acc_api.call_api(endpoint)
        if not result["data"]:
            endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms?formDateMin={previous_month_abbr} 1, {datetime.now().year}"
            result = acc_api.call_api(endpoint)

        elif not result["data"] and datetime.now().strftime("%b") == "Jan":
            endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms?formDateMin={previous_month_abbr} 1, {today.year-1}"
            result = acc_api.call_api(endpoint)



        def filter_by_form_num(data, form_num):
            """
            Filters the response data to include only entries with the specified formNum.

            Args:
                data (list): The list of data dictionaries to filter.
                form_num (int): The formNum to filter by.

            Returns:
                list: A filtered list of dictionaries with the specified formNum.
            """
            return [item for item in data if item.get("formNum") == form_num]



        # Specify the template file and output folder
        template_file = 'templates/template.xlsx'
        output_folder = 'modified_files'

        # Create an instance of the ExcelModifier class
        modifier = ExcelModifier(template_file, output_folder)
        new_df2 = []
        for data in result["data"]:
            # pretty_print_json(data["customValues"])
            # print(data["pdfValues"])


            # Convert data to dataframe (table in excel)
            df = pd.DataFrame(data["pdfValues"])

            categories = {'المعدات': [], 'عدد ساعات معدة': [], 'اسم المقاول': []}
            for item in data["pdfValues"]:
                for key in categories:
                    if item['name'].startswith(key):
                        categories[key].append(item['value'])

            df2 = pd.DataFrame({k: pd.Series(v) for k, v in categories.items()})
            df2 = df2[~df2["اسم المقاول"].isna()]

            df2["new2"] = df2["اسم المقاول"].astype(str) + ' ' + df2["المعدات"].astype(str)

            categories = {'الخصم': [],'قيمة الساعة': [], 'المعدات بيان ': [], 'بيان': []}
            for item in data["pdfValues"]:
                for key in categories:
                    if item["name"].startswith(key):
                        categories[key].append(item["value"])

            df3 = pd.DataFrame({k: pd.Series(v) for k, v in categories.items()})
            df3 = df3[~df3["بيان"].isna()]
            df3["new1"] = df3["بيان"].astype(str) + ' ' + df3["المعدات بيان "].astype(str)



            new_df = pd.merge(df3, df2, left_on='new1', right_on='new2', how='outer')

            new_df["عدد ساعات معدة"]=new_df["عدد ساعات معدة"].astype(float)
            new_df["الخصم"]=new_df["الخصم"].astype(float)

            project_name = next((item["value"] for item in data["pdfValues"] if item["name"] == "اسم المشروع"), None)
            form_Num= data["formNum"]
            form_date= data["formDate"]
            form_desc=data["description"]
            provider_type = next((item["value"] for item in data["pdfValues"] if item["name"] == "نوع مقدم الخدمة"), None)
            WBS_code = next((item["value"] for item in data["pdfValues"] if item["name"] == "كود البند"), None)
            change_on = next((item["value"] for item in data["pdfValues"] if item["name"] == "بالخصم على"), None)
            order_name = next((item["value"] for item in data["pdfValues"] if item["name"] == "اسم الطلب"), None)
            new_df["project_name"]=project_name
            new_df["form_Num"] = form_Num
            new_df["form_date"] = form_date
            new_df["form_desc"] = form_desc
            new_df["provider_type"] = provider_type
            new_df["WBS_code"] = WBS_code
            new_df["change_on"] = change_on
            new_df["order_name"] = order_name


            new_df2.append(new_df)

        final_df = pd.concat(new_df2, ignore_index=True)


        df_filtered = final_df[final_df["اسم المقاول"].str.strip() != ""]

        final_df["اسم المقاول"] = final_df["اسم المقاول"].apply(lambda x: x.strip() if isinstance(x, str) else x)
        final_df["اسم المقاول"].replace("", pd.NA, inplace=True)

        df_filtered = final_df.dropna(subset=["اسم المقاول"])
        print(df_filtered)

        # df_filtered.to_excel("D:\\OneDrive - Square Engineering Firm\\Users\\ABDALLAH.MAMDOUH\\Desktop\\New Microsoft Excel Worksheet9.xlsx", index=False)
        date_now = (datetime.now().strftime("%A, %B %d, %Y"))

        for proj in df_filtered["project_name"].unique():
            df_project = df_filtered[df_filtered["project_name"] == proj]
            print(f"Processing project: {proj}")
            print(f"Data for project {proj}:")
            print(df_project)  # Check if the data is correct

            # Open workbook for current project
            modifier.open_workbook()
            m = 7  # Starting from row 7

            # Modify each row for the project
            for _, row in df_project.iterrows():
                print(f"Modifying row {m} with data: {row}")  # Check the row data before modifying

                # Ensure all required columns have data
                if not pd.isnull(row["form_Num"]) and not pd.isnull(row["project_name"]) and not pd.isnull(row["form_date"]):
                    # Modify cells only if data is available
                    modifier.modify_cell(f'A{m}', row.get("form_Num", ""))
                    modifier.modify_cell(f'B{m}', row.get("project_name", ""))
                    modifier.modify_cell(f'D{m}', row.get("form_date", ""))
                    modifier.modify_cell(f'H{m}', row.get("form_desc", ""))
                    modifier.modify_cell(f'P{m}', row.get("قيمة الساعة", ""))
                    modifier.modify_cell(f'R{m}', row.get("الخصم", ""))
                    modifier.modify_cell(f'S{m}', f"=P{m}*Q{m}-R{m}")
                    modifier.modify_cell(f'L{m}', row.get("اسم المقاول", ""))
                    modifier.modify_cell(f'M{m}', row.get("المعدات", ""))
                    modifier.modify_cell(f'Q{m}', row.get("عدد ساعات معدة", ""))
                    modifier.modify_cell(f'O{m}', 1)
                    modifier.modify_cell(f'N{m}', provider_type)
                    modifier.modify_cell(f'I{m}', WBS_code)
                    modifier.modify_cell(f'J{m}', change_on)
                    modifier.modify_cell(f'K{m}', order_name)

                    # Only increment row number if data is populated
                    m += 1
                else:
                    print(f"Skipping row {m} due to missing data")

                # Insert a new row if necessary
                if m >= 8:  # Ensuring it doesn't insert too early
                    modifier.insert_row(m)

            modifier.modify_cell(f'S{m + 2}', f"=SUM(S7:S{m})")


            try:
                # modifier.save_workbook()  # Save workbook after all rows are processed
                modifier.export_to_pdf(f"{proj} - {date_now}.pdf")

            except Exception as e:
                print(f"Error saving workbook: {e}")
            finally:
                modifier.close_workbook()

        df_2 = df_filtered.copy()
        df_2["الكمية"]=1
        final_df = df_2.groupby(["اسم المقاول", "المعدات", "قيمة الساعة","project_name"], as_index=False)[["الخصم", "عدد ساعات معدة", "الكمية"]].sum()
        # final_df.to_excel("D:\\OneDrive - Square Engineering Firm\\Users\\ABDALLAH.MAMDOUH\\Desktop\\New Microsoft Excel Worksheet22.xlsx", index=False)

        for proj in final_df["project_name"].unique():
            df_project = final_df[final_df["project_name"] == proj]
            print(f"Processing project: {proj}")
            print(f"Data for project {proj}:")
            print(df_project)  # Check if the data is correct


            # Specify the template file and output folder
            template_file = 'templates/summaryTemplate.xlsx'
            output_folder = 'modified_files'

            # Create an instance of the ExcelModifier class
            modifier = ExcelModifier(template_file, output_folder)

            modifier.open_workbook()
            m = 7  # Starting from row 7

            # Modify each row for the project
            for _, row in df_project.iterrows():
                print(f"Modifying row {m} with data: {row}")  # Check the row data before modifying

                # Ensure all required columns have data
                if not pd.isnull(row["project_name"]):
                    # Modify cells only if data is available
                    modifier.modify_cell(f'A{m}', row.get("project_name", ""))
                    modifier.modify_cell(f'I{m}', row.get("قيمة الساعة", ""))
                    modifier.modify_cell(f'K{m}', row.get("الخصم", ""))
                    modifier.modify_cell(f'L{m}', f"=I{m}*J{m}-K{m}")
                    modifier.modify_cell(f'E{m}', row.get("اسم المقاول", ""))
                    modifier.modify_cell(f'F{m}', row.get("المعدات", ""))
                    modifier.modify_cell(f'J{m}', row.get("عدد ساعات معدة", ""))
                    modifier.modify_cell(f'H{m}', row.get("الكمية", ""))
                    modifier.modify_cell(f'G{m}', provider_type)
                    modifier.modify_cell(f'B{m}', WBS_code)
                    modifier.modify_cell(f'C{m}', change_on)
                    modifier.modify_cell(f'D{m}', order_name)

                    # Only increment row number if data is populated
                    m += 1
                else:
                    print(f"Skipping row {m} due to missing data")

                # Insert a new row if necessary
                if m >= 8:  # Ensuring it doesn't insert too early
                    modifier.insert_row(m)

            modifier.modify_cell(f'L{m + 2}', f"=SUM(L7:L{m})")


            try:
                # modifier.save_workbook()  # Save workbook after all rows are processed
                file_name = f"{proj} - {date_now} Summary"
                pdf_path = f"modified_files/{file_name}.pdf"    
                modifier.export_to_pdf_no_upload(excel_filename=file_name)
                acc_api.upload_pdf_to_acc(pdf_path=pdf_path, filename=file_name, folder_name=f"Equipment/{proj}")

            except Exception as e:
                print(f"Error saving workbook: {e}")
            finally:
                modifier.close_workbook()




    except EnvironmentError as env_err:
        print(f"Environment error: {env_err}")
    # except Exception as e:
    #     print(f"An error occurred: {e}")

