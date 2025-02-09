import requests
import ACCAPI


accapi = ACCAPI.ACCAPI()


# Get list of files in the folder
url = f"cost/v1/containers/{accapi.CONTAINER_ID}/documents"
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
