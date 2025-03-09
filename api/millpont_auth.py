import time
import requests

# Token management
token = None
token_expiry = 0

def get_token():
    global token, token_expiry
    url = "https://api.meti.millpont.com/token"
    payload = {
        "email": "email",
        "password": "password"
    }
    response = requests.post(url, json=payload)
    
    # Print response for debugging
    print(f"Status code: {response.status_code}")
    print(f"Response content: {response.text}")
    
    response_data = response.json()
    
    # Check if the expected keys are in the response
    if "access_token" in response_data:
        token = response_data["access_token"]
        token_expiry = time.time() + response_data.get("expires_in", 3600)  # Default to 1 hour if not specified
    elif "token" in response_data:
        token = response_data["token"]
        token_expiry = time.time() + response_data.get("expires_in", 3600)  # Default to 1 hour if not specified
    else:
        raise ValueError(f"Authentication failed: {response_data.get('error', 'Unknown error')}")

def make_authenticated_request(endpoint, method="GET", data=None):
    global token, token_expiry
    if token is None or time.time() >= token_expiry:
        get_token()  # Refresh the token
    
    headers = {"Authorization": f"Bearer {token}"}
    url = f"https://api.meti.millpont.com/{endpoint}"
    
    try:
        if method == "GET":
            response = requests.get(url, headers=headers)
        elif method == "POST":
            response = requests.post(url, json=data, headers=headers)
        else:
            raise ValueError(f"Unsupported HTTP method: {method}")
        
        response.raise_for_status()  # Raise an exception for 4XX/5XX responses
        return response.json()
    except requests.exceptions.RequestException as e:
        print(f"Request error: {e}")
        if hasattr(e.response, 'text'):
            print(f"Response content: {e.response.text}")
        raise

# Example usage
try:
    get_token()  # Get the token initially
    response = make_authenticated_request("sources", method="GET")
    print(response)
except Exception as e:
    print(f"Error: {e}")
