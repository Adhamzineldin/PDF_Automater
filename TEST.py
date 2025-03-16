import requests
import ACCAPI

accapi = ACCAPI.ACCAPI()

# Step 1: Get the first Hub ID
def get_hub_id():
    url = "project/v1/hubs"
    response = accapi.call_api(url)

    if response and "data" in response and len(response["data"]) > 0:
        hub_id = response["data"][0].get("id")  # Use .get() to avoid KeyError
        if hub_id:
            print(f"Hub ID: {hub_id}")
            return hub_id
    print("No hubs found.")
    return None

# Step 2: Get all projects in the hub
def get_projects(hub_id):
    url = f"project/v1/hubs/{hub_id}/projects"
    response = accapi.call_api(url)

    if response and "data" in response:
        projects = response["data"]
        print(f"Found {len(projects)} projects.")
        return projects
    else:
        print("No projects found.")
        return []

# Step 3: Get top-level folders in the project (using the correct endpoint with hub_id)
def get_top_folders(hub_id, project_id):
    url = f"project/v1/hubs/{hub_id}/projects/{project_id}/topFolders"  # Correct endpoint with hub_id
    response = accapi.call_api(url)

    if response and "data" in response:
        folders = response["data"]
        if len(folders) == 0:
            print(f"No top-level folders found for project {project_id}.")
        for folder in folders:
            folder_name = folder["attributes"].get("name", "Unknown Folder")  # Use .get() to avoid KeyError
            folder_id = folder.get("id", "Unknown ID")  # Default to Unknown if ID is not found
            print(f"📂 Folder: {folder_name} ({folder_id})")
            list_folder_contents(project_id, folder_id)
    else:
        print(f"No top-level folders found for project {project_id}.")

# Step 4: List contents of each folder (files and subfolders)
def list_folder_contents(project_id, folder_id):
    url = f"data/v1/projects/{project_id}/folders/{folder_id}/contents"
    response = accapi.call_api(url)

    if response and "data" in response:
        for item in response["data"]:
            # print(f"📄 {item}")
            # Try to get displayName first, then fall back to name
            item_name = item["attributes"].get("displayName", item["attributes"].get("name", "Unnamed Item"))
            item_type = item.get("type", "Unknown Type")  # Default to Unknown Type if not found
            print(f"   - {item_type}: {item_name}")
    
            if item_type == "folders":
                list_folder_contents(project_id, item.get("id", "Unknown ID"))  # Recursively handle subfolders
    else:
        print(f"No contents found in folder {folder_id}.")


# Main Execution
hub_id = get_hub_id()
if hub_id:
    projects = get_projects(hub_id)
    for project in projects:
        project_id = project.get("id", "Unknown ID")  # Handle missing project ID
        project_name = project["attributes"].get("name", "Unnamed Project")  # Handle missing project name
        print(f"\n🛠️ Project: {project_name} ({project_id})")
        if project_name == "Sample Project - Seaport Civic Center":
            
            get_top_folders(hub_id, project_id)  # Pass both hub_id and project_id
