# reporting
This repository contains two Python scripts to process and filter historical data. The scripts fetch users authentication details, save it to an Excel file (.xlsx format), and filter the data based on specified criteria.

## Prerequisites

- Python 3.6 or higher
- `pip` for package management

## Setup

1. **Clone the repository**

    ```
    git clone https://github.com/houdabouta/reporting.git
    cd reporting
    ```

2. **Create a virtual environment and activate it**

    ```
    python -m venv venv
    source venv/bin/activate  # On Windows use `venv\Scripts\activate`
    ```

3. **Install the required packages**

    ```
    pip install -r requirements.txt
    ```

## Scripts

### 1. `users_reporting.py`

This script fetches historical data from a given API endpoint, saves it to an Excel file, and ensures proper formatting for specific fields.

#### Usage

1. **Set up the environment variables in a `.env` file**

    ```
    USERNAME=your_username
    PSSWD=your_password
    SPECIAL_SWITCH_ADMIN=your_special_switch_admin
    BASE_URL=https://admin-api-v1.cloudi-fi.net
    ```

2. **Run the script to fetch the history from Cloudi-Fi**

    ```
    python users_reporting.py
    ```

### 2. `filter_data.py`

This script filters the data in the generated Excel file based on `authDate` and `location`.

#### Usage

    ```
    python filter_data.py --input_file results/history_data.xlsx --output_file results/filtered_data.xlsx --start_date YYYY-MM-DD --end_date YYYY-MM-DD --location your_location
    ```
