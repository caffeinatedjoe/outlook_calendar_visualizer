# Calendar Visualizer

The Calendar Visualizer is a Python script that generates an Excel-based calendar report of employee time off (PTO), travel, and holidays. It fetches event data from iCalendar (.ics) URLs, intelligently maps events to employees using a generative AI model, and presents the information in a hierarchically organized and color-coded Excel spreadsheet.

## Features

- **Hierarchical Employee Display**: Reflects the company's organizational structure in the Excel output.
- **Multiple Calendar Integration**: Fetches and combines data from separate iCalendar URLs for PTO and travel.
- **AI-Powered Event Mapping**: Utilizes Google's Gemini model to accurately associate calendar events with the correct employees, even with ambiguous event titles.
- **Holiday Recognition**: Automatically identifies and applies US, French, and company-wide holidays to the appropriate employees based on their location.
- **Customizable Date Range**: Allows users to specify the number of months for which to generate the report.
- **Formatted Excel Output**: Creates a clean, color-coded, and easy-to-read Excel file with events organized by month, week, and day.

## Prerequisites

- Python 3.x
- Pip (Python package installer)

## Setup

1.  **Clone the repository:**
    ```bash
    git clone <repository-url>
    cd <repository-directory>
    ```

2.  **Install the required Python libraries:**
    ```bash
    pip install -r requirements.txt
    ```

3.  **Configure API Keys and URLs:**
    - Create a `config.py` file in the root directory.
    - Add your Google Generative AI API key and iCalendar URLs to this file:
      ```python
      # config.py
      API_KEY = "YOUR_GOOGLE_AI_API_KEY"
      PTO_CALENDAR_URL = "URL_for_PTO_iCal_feed"
      TRAVEL_CALENDAR_URL = "URL_for_Travel_iCal_feed"
      ```

4.  **Create the Employee Data File:**
    - Create an `employees.json` file in the root directory.
    - This file should contain a hierarchical list of all employees, their locations, and their reporting structure.
    - **Example `employees.json` structure:**
      ```json
      [
        {
          "name": "CEO Name",
          "location": "US",
          "reports": [
            {
              "name": "Manager Name",
              "location": "France",
              "reports": [
                {
                  "name": "Direct Report Name",
                  "location": "France",
                  "reports": []
                }
              ]
            }
          ]
        }
      ]
      ```

5.  **Zscaler Certificate (If applicable):**
    - If your network uses Zscaler for SSL inspection, you will need to provide the Zscaler root certificate to avoid SSL errors.

    - **How to export the certificate:**
      1.  **Open Windows Certificate Manager**:
          - Press `Win + R`, type `certmgr.msc`, and press Enter.
      2.  **Locate the Zscaler root certificate**:
          - In the left panel, navigate to `Trusted Root Certification Authorities` > `Certificates`.
          - Find the certificate named `Zscaler Root CA` (or similar).
      3.  **Export the certificate**:
          - Double-click the certificate to open its details.
          - Go to the `Details` tab and click `Copy to File...`.
          - In the Certificate Export Wizard, select `Base-64 encoded X.509 (.CER)`.
          - Save the file as `zscaler_root.crt`.

    - **Important**: Place the exported `zscaler_root.crt` file in the root directory of this project. The script is pre-configured to use this path for verification.

## Usage

Run the script from your terminal. You can optionally specify the number of months you want the report to cover.

-   **To generate a report for the default 6 months:**
    ```bash
    python main.py
    ```

-   **To generate a report for a specific number of months (e.g., 3):**
    ```bash
    python main.py 3
    ```

The script will generate an Excel file in the root directory named `calendar_view_MMDDYY.xlsx`, where `MMDDYY` is the current date.

## How It Works

1.  **Load Data**: The script begins by loading the employee hierarchy from `employees.json`.
2.  **Fetch Events**: It then fetches calendar data from the iCalendar URLs specified in `config.py`.
3.  **Filter by Date**: Events are filtered to include only those within the specified date range.
4.  **AI Mapping**: The unique names of the filtered events are sent to the Google Gemini model. The model returns a JSON object that maps each event to the corresponding employee(s) or a special holiday tag.
5.  **Process Data**: The script processes this mapping to create a data structure that associates each employee with their PTO, travel, and holiday dates.
6.  **Generate Excel**: Finally, it generates a formatted Excel spreadsheet. Employees are listed hierarchically, and their corresponding dates are color-coded for easy visualization.