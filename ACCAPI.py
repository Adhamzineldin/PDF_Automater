import os
from dotenv import load_dotenv
from urllib.parse import urlencode
import requests

# Load environment variables
load_dotenv()

class ACCAPI:
    def __init__(self):
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
                "scope": "data:read"
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
