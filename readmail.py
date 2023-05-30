import json
import logging
import requests
import msal

# Define the Office 365 API endpoint for your application
endpoint = "https://graph.microsoft.com/v1.0"

# Define your Office 365 application's client ID and secret
client_id = "<CLientID>"
client_secret = "<ClientSecret>"

# Define the Office 365 application's authority URL and scope
authority_url = "https://login.microsoftonline.com/<TenantID>"
scope = ["https://graph.microsoft.com/.default"]

app = msal.ConfidentialClientApplication(
    client_id, authority=authority_url,
    client_credential=client_secret,
)

result = None

result = app.acquire_token_silent(scope, account=None)

if not result:
    logging.info("No suitable token exists in cache. Let's get a new one from AAD.")
    result = app.acquire_token_for_client(scopes=scope)

if "access_token" in result:
    # Set up the request headers
    headers = {
        "Authorization": "Bearer " + result["access_token"],
        "Accept": "application/json",
        "Content-Type": "application/json"
    }
    # Define the API endpoint for reading messages from a mailbox
    mailbox = "example@email.com" # replace with the mailbox you want to read messages from
    endpoint = endpoint + f"/users/{mailbox}/messages"
    # Make the API request
    response = requests.get(endpoint, headers=headers)
    # Print the response to the console
    print(response.json())
else:
    print(result.get("error"))
    print(result.get("error_description"))
    print(result.get("correlation_id"))
