from flask import Flask, request
import os
import requests
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Retrieve credentials from .env file
CLIENT_ID = os.getenv("AUTODESK_CLIENT_ID")
CLIENT_SECRET = os.getenv("AUTODESK_CLIENT_SECRET")
BASE_URL = os.getenv("AUTODESK_API_URL", "https://developer.api.autodesk.com")
REDIRECT_URI = os.getenv("AUTODESK_REDIRECT_URI")  # The redirect URI you set in Autodesk Developer Console

# Flask setup
app = Flask(__name__)

@app.route('/')
def home():
    return "Welcome to the Autodesk OAuth callback server!"

# Callback endpoint to capture the authorization code
@app.route('/callback')
def callback():
    # Get the authorization code from the URL
    auth_code = request.args.get('code')

    if not auth_code:
        return "Error: No authorization code found in the URL.", 400

    # Display the authorization code for debugging
    print(f"Authorization Code received: {auth_code}")


    return f"AUTH CODE: {auth_code}"


if __name__ == '__main__':
    # Run the server on localhost:5000
    app.run(host='localhost', port=5000)
