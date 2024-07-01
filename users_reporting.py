import os
import requests
import logging
import time
from openpyxl import Workbook
from openpyxl.styles import NamedStyle
from dotenv import load_dotenv
from requests.exceptions import HTTPError, RequestException
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def clean_value(value):
    """Extract the value after the last '/' if it exists."""
    if value and isinstance(value, str) and '/' in value:
        return value.rsplit('/', 1)[-1]
    return value

def handle_non_convertible_value(value):
    """Convert non-convertible values to a suitable format for Excel."""
    if isinstance(value, list):
        return ', '.join(map(str, value))  # Convert list to comma-separated string
    if value is None:
        return ''
    return value

def flatten_dict(d, parent_key='', sep='.'):
    """Flatten nested dictionaries."""
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        else:
            items.append((new_key, v))
    return dict(items)

def main():
    # Load environment variables from .env file
    load_dotenv()

    # Get environment variables
    username = os.getenv("USERNAME")
    password = os.getenv("PSSWD")
    special_switch_admin = os.getenv("SPECIAL_SWITCH_ADMIN")
    base_url = os.getenv("BASE_URL")
    auth_url = f"{base_url}/auth/json"
    api_base_url = f"{base_url}/histories"

    # Function to authenticate and get auth token
    @retry(
        stop=stop_after_attempt(3),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((HTTPError, RequestException))
    )
    def authenticate(auth_url, username, password, switch_user):
        auth_data = {
            "username": username,
            "password": password
        }
        headers = {
            "X-Switch-User": switch_user,
            "Content-Type": "application/json"
        }
        response = requests.post(auth_url, json=auth_data, headers=headers, timeout=10)
        response.raise_for_status()
        return response.json().get("token")

    try:
        auth_token = authenticate(auth_url, username, password, special_switch_admin)
    except Exception as e:
        logging.error(f"Failed to obtain authentication token: {e}")
        return

    logging.info("Successfully authenticated")

    # Set headers for API requests
    api_headers = {
        "X-Switch-User": special_switch_admin,
        "Accept": "application/ld+json",
        "Authorization": f"Bearer {auth_token}"
    }

    # Function to fetch data from the API with pagination
    @retry(
        stop=stop_after_attempt(5),
        wait=wait_exponential(multiplier=1, min=2, max=10),
        retry=retry_if_exception_type((HTTPError, RequestException))
    )
    def fetch_page(api_url, headers, page):
        response = requests.get(f"{api_url}?page={page}", headers=headers, timeout=10)
        response.raise_for_status()
        return response.json()

    def fetch_data(api_url, headers):
        data = []
        page = 1
        while True:
            try:
                logging.info(f"Fetching page {page}")
                json_data = fetch_page(api_url, headers, page)
                members = json_data.get("hydra:member", [])
                if not members:
                    logging.info("No more data to fetch")
                    break
                data.extend(members)
                page += 1
                time.sleep(1)  # Delay between page fetches
            except HTTPError as http_err:
                if http_err.response.status_code == 429:  # Too Many Requests
                    logging.warning("Throttling error: Too Many Requests. Retrying...")
                else:
                    logging.error(f"HTTP error occurred while fetching page {page}: {http_err}")
                    break
            except RequestException as req_err:
                logging.error(f"Request error occurred while fetching page {page}: {req_err}")
                break
            except Exception as err:
                logging.error(f"Error occurred while fetching page {page}: {err}")
                break
        return data

    # Fetch data
    data = fetch_data(api_base_url, api_headers)

    if not data:
        logging.info("No data fetched")
        return

    logging.info("Data fetched successfully, saving to Excel")

    # Ensure results directory exists
    results_dir = "results"
    os.makedirs(results_dir, exist_ok=True)

    # Save data to Excel file
    excel_file = os.path.join(results_dir, "history_data.xlsx")
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "History Data"

    # Flatten and clean the data
    flat_data = [flatten_dict(item) for item in data]

    # Extract headers dynamically from the first item in the data list
    if flat_data:
        headers = list(flat_data[0].keys())
        sheet.append(headers)

        # Define text style for email
        text_style = NamedStyle(name="text_style", number_format="@")

        for item in flat_data:
            row = []
            for header in headers:
                value = clean_value(item.get(header))
                value = handle_non_convertible_value(value)
                row.append(value)

            sheet.append(row)

        # Apply text style to the email column
        for cell in sheet["G"]:
            cell.style = text_style

        workbook.save(excel_file)
        logging.info(f"Data saved to {excel_file}")

if __name__ == "__main__":
    main()