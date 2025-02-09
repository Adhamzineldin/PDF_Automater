import requests
import ACCAPI


accapi = ACCAPI.ACCAPI()


# Get list of files in the folder
url = f"data/v1/projects/32d52d7d-b9fd-473d-8963-21ba503f854b/folders/topFolders"
response = accapi.call_api(url)



print(response)



# for file in files:
#     if file["attributes"]["name"].endswith(".zip"):
#         download_url = file["links"]["download"]
#         zip_name = file["attributes"]["name"]
#         zip_response = requests.get(download_url, headers=headers)
# 
#         with open(zip_name, "wb") as f:
#             f.write(zip_response.content)
#         print(f"Downloaded: {zip_name}")
