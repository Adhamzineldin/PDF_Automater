import subprocess

from dotenv import load_dotenv
from urllib.parse import urlencode
import requests
import base64
import re
import os
from svgpathtools import svg2paths
from PIL import Image, ImageDraw

# Load environment variables
load_dotenv()

class ACCAPI:
    def __init__(self):
        self.modified_folder = "./Modified_Files"
        self.CLIENT_ID = os.getenv("AUTODESK_CLIENT_ID")
        self.CLIENT_SECRET = os.getenv("AUTODESK_CLIENT_SECRET")
        self.BASE_URL = os.getenv("AUTODESK_API_URL", "https://developer.api.autodesk.com")
        self.REDIRECT_URI = os.getenv("AUTODESK_REDIRECT_URI")
        self.CONTAINER_ID = os.getenv("AUTODESK_CONTAINER_ID")

        self.validate_env_vars()

    # Function to validate environment variables
    def validate_env_vars(self):
        missing_vars = [var for var in ["AUTODESK_CLIENT_ID", "AUTODESK_CLIENT_SECRET", "AUTODESK_REDIRECT_URI", "AUTODESK_CONTAINER_ID"] if not os.getenv(var)]
        if missing_vars:
            raise EnvironmentError(f"Missing required environment variables: {', '.join(missing_vars)}")

    # Function to get the authorization URL
    def get_authorization_url(self):
        auth_url = f"{self.BASE_URL}/authentication/v2/authorize"
        params = {
                "client_id": self.CLIENT_ID,
                "response_type": "code",
                "redirect_uri": self.REDIRECT_URI,
                "scope": "data:read data:write account:read account:write"
        }
        return f"{auth_url}?{urlencode(params)}"

    # Function to get the access token and refresh token using the authorization code
    def get_access_token(self, auth_code):
        token_url = f"{self.BASE_URL}/authentication/v2/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        payload = {
                "client_id": self.CLIENT_ID,
                "client_secret": self.CLIENT_SECRET,
                "grant_type": "authorization_code",
                "code": auth_code,
                "redirect_uri": self.REDIRECT_URI
        }

        try:
            response = requests.post(token_url, headers=headers, data=payload)
            response.raise_for_status()  # Raise an exception for HTTP errors
            data = response.json()
            access_token = data.get("access_token")
            refresh_token = data.get("refresh_token")

            if not access_token or not refresh_token:
                raise ValueError("Access token or refresh token not found in response.")

            # Save the refresh token securely
            self.save_refresh_token(refresh_token)

            return access_token, refresh_token

        except requests.exceptions.HTTPError as http_err:
            print(f"HTTP error occurred: {http_err}")
            print("Response content:", response.content.decode())
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            raise

    # Function to save the refresh token in a local file
    def save_refresh_token(self, refresh_token):
        with open("refresh_token.txt", "w") as file:
            file.write(refresh_token)

    # Function to load the refresh token from the local file
    def load_refresh_token(self):
        if os.path.exists("refresh_token.txt"):
            with open("refresh_token.txt", "r") as file:
                return file.read().strip()
        return None

    # Function to refresh the access token using the refresh token
    def refresh_access_token(self, refresh_token):
        token_url = f"{self.BASE_URL}/authentication/v2/token"
        headers = {"Content-Type": "application/x-www-form-urlencoded"}
        payload = {
                "client_id": self.CLIENT_ID,
                "client_secret": self.CLIENT_SECRET,
                "grant_type": "refresh_token",
                "refresh_token": refresh_token
        }

        try:
            response = requests.post(token_url, headers=headers, data=payload)
            response.raise_for_status()  # Raise an exception for HTTP errors
            data = response.json()
            new_access_token = data.get("access_token")
            new_refresh_token = data.get("refresh_token")  # Get the new refresh token

            if not new_access_token:
                raise ValueError("New access token not found in response.")

            # Save the new refresh token securely
            self.save_refresh_token(new_refresh_token)

            return new_access_token, new_refresh_token  # Return both access and refresh tokens

        except requests.exceptions.HTTPError as http_err:
            if response.status_code == 400 and "invalid_grant" in response.json().get("error", ""):
                print("Refresh token invalid or expired. Triggering re-authentication.")
                return None, None  # Indicate failure for token refresh, re-authentication needed
            print(f"HTTP error occurred: {http_err}")
            print("Response content:", response.content.decode())
            raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            raise
    
    def decode_svg(self, coded_svg_code):
      
                
        # Check if the string looks like base64
        if re.match(r'^[A-Za-z0-9+/=]+$', coded_svg_code ):
            print("It looks like a Base64 string. Let's try decoding it.")
        
            # Decode the base64 string
            decoded_bytes = base64.b64decode(coded_svg_code)
        
            # Convert the decoded bytes back to string (UTF-8)
            decoded_svg = decoded_bytes.decode('utf-8')
        
            
            return decoded_svg
        else:
            print("This does not appear to be a Base64-encoded string.")

    def convert_svg_to_png(self, svg_code, output_path):
        """
        Converts SVG code to PNG using svgpathtools and Pillow and returns the path to the generated PNG.
        
        Parameters:
        - svg_code: str, SVG code as a string.
        
        Returns:
        - str: Path to the saved PNG file.
        """
        try:
            # Step 1: Save the SVG code to a temporary file
            temp_svg_path = os.path.join(self.modified_folder, f"{output_path}.svg")
            with open(temp_svg_path, "w", encoding="utf-8") as svg_file:
                svg_file.write(svg_code)
    
            # Step 2: Parse the SVG to extract paths
            paths, attributes = svg2paths(temp_svg_path)
    
            # Step 3: Create a new blank image (white background)
            width, height = 600, 300  # You can adjust the size as needed
            img = Image.new('RGBA', (width, height), (255, 255, 255, 255))  # White background
            draw = ImageDraw.Draw(img)
    
            # Step 4: Draw the paths onto the image
            for path in paths:
                for segment in path:
                    start = segment.start
                    end = segment.end
                    draw.line((start.real, start.imag, end.real, end.imag), fill='black', width=2)
    
            # Step 5: Save the image as PNG
            temp_png_path = os.path.join(self.modified_folder, f"{output_path}.png")
            img.save(temp_png_path, "PNG")
    
            # Return the path to the PNG image
            return temp_png_path
    
        except Exception as e:
            print(f"An error occurred while processing the SVG: {e}")
            return None

    def upload_pdf_to_acc(self, pdf_path, excel_filename, project_name="Information Systems Workspace", folder_name="Adhams_Server"):
        """
        Function to export the PDF to a specified location on the Autodesk Odrive and refresh the directory.
    
        :param project_name: 
        :param folder_name: 
        :param pdf_path: The path to the generated PDF.
        :param excel_filename: The filename to save the PDF as (without extension).
        """
        # Get the user's home directory path
        home_dir = os.path.expanduser("~")
    
        # Define the new PDF path
        new_pdf_path = os.path.join(home_dir, f'server/odrive/Autodesk/Square Engineering Firm/{project_name}/Project Files/{folder_name}/{excel_filename}.pdf')
    
        # Save the current working directory to return to it later
        original_dir = os.getcwd()
    
        # Ensure the target directory exists, if not, create it
        project_files_dir = os.path.join(home_dir, f'server/odrive/Autodesk/Square Engineering Firm/{project_name}/Project Files')
        adham_server_dir = os.path.dirname(new_pdf_path)
        if not os.path.exists(adham_server_dir):
            os.makedirs(adham_server_dir)
            print(f"Directory {adham_server_dir} created.")
    
        # If the output file already exists, delete it to avoid conflicts
        if os.path.exists(new_pdf_path):
            os.remove(new_pdf_path)
    
        # Use the 'cp' command to copy the generated PDF to the new location
        subprocess.run(['cp', pdf_path, new_pdf_path], check=True)
    
        # Change the current working directory to the folder containing the PDF
        os.chdir(project_files_dir)
    
        # Run the 'odrive refresh' command in the current directory (which is now pdf_dir)
        subprocess.run([os.path.expanduser("~/.odrive-agent/bin/odrive"), 'refresh', '.'], check=True)
    
        print(f"PDF also exported at {new_pdf_path}")
    
        # Change back to the original working directory
        os.chdir(original_dir)
        print(f"Changed back to the original working directory: {original_dir}")
        
        

    # Dynamic function to call any Autodesk API endpoint and return the unfiltered response
    def call_api(self, endpoint, params=None):
        # Load the refresh token
        refresh_token = self.load_refresh_token()

        # If no refresh token is found, prompt for authorization code
        if not refresh_token:
            print("No refresh token found. Please authenticate first.")
            auth_url = self.get_authorization_url()
            print(f"Visit this URL to authenticate and get the code: {auth_url}")
            auth_code = input("Enter the authorization code: ")
            access_token, refresh_token = self.get_access_token(auth_code)

        # Attempt to refresh the token and get a valid access token
        access_token, _ = self.refresh_access_token(refresh_token)  # Refresh token to get the access token

        # If the refresh token failed, prompt for the initial authorization flow
        if not access_token:
            print("Refresh token expired or invalid. Please authenticate again.")
            auth_url = self.get_authorization_url()
            print(f"Visit this URL to authenticate and get the code: {auth_url}")
            auth_code = input("Enter the authorization code: ")
            access_token, refresh_token = self.get_access_token(auth_code)  # Re-authenticate
            self.save_refresh_token(refresh_token)  # Save the new refresh token

        # API call to the specified endpoint
        url = f"{self.BASE_URL}/{endpoint}"
        headers = {
                "Authorization": f"Bearer {access_token}",
                "Content-Type": "application/json",
        }

        try:
            # Send the GET request to the API endpoint
            response = requests.get(url, headers=headers, params=params)
            response.raise_for_status()  # Raise an exception for HTTP errors
            return response.json()  # Return the raw JSON response from the API
        except requests.exceptions.HTTPError as http_err:
            if response.status_code == 401:  # Unauthorized, typically means access token expired
                print("Access token expired. Refreshing token and retrying...")
                # Refresh the token and retry the request
                access_token, refresh_token = self.refresh_access_token(refresh_token)
                self.save_refresh_token(refresh_token)  # Save the new refresh token
                return self.call_api(endpoint, params)  # Retry the API call with the new token
            else:
                print(f"HTTP error occurred: {http_err}")
                print("Response content:", response.content.decode())
                raise
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            raise

# Main workflow
def main():
    try:
        # Instantiate the ACCAPI class
        acc_api = ACCAPI()

        # Dynamic API call example
        endpoint = f"construction/forms/v1/projects/{acc_api.CONTAINER_ID}/forms"
        result = acc_api.call_api(endpoint)

        # Print the raw result (unfiltered)
        print("API Response:", result)

    except EnvironmentError as env_err:
        print(f"Environment error: {env_err}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Run the main function
if __name__ == "__main__":
    main()
