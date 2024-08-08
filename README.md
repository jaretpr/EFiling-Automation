# Automated E-Filing Script

This repository contains a Python script that automates the e-filing process for garnishment service returns using Selenium and Tkinter. The script logs into a court portal, uploads documents, and updates a status spreadsheet.

## Features

- Automated login to the court portal
- File uploading and form submission
- Status updates in an Excel workbook
- Tkinter-based GUI for easy interaction

## Prerequisites

- Python 3.x
- Google Chrome browser
- ChromeDriver
- Required Python packages (specified in `requirements.txt`)

## Setup

1. **Clone the repository:**

    ```sh
    git clone https://github.com/yourusername/your-repo-name.git
    cd your-repo-name
    ```

2. **Install dependencies:**

    ```sh
    pip install -r requirements.txt
    ```

3. **Download and setup ChromeDriver:**

    - Download ChromeDriver from [here](https://sites.google.com/chromium.org/driver/downloads).
    - Place the `chromedriver` executable in a known directory and update the `executable_path` in the script.

4. **Update credentials:**

    - In the script, replace the placeholder credentials with actual usernames and passwords for the attorneys.

5. **Set file paths:**

    - Update the file paths in the script (`wb_path`, `file_path`) to match your directory structure.

## Usage

1. **Run the script:**

    ```sh
    python E-filing_Automation.py
    ```

2. **Interact with the GUI:**

    - Click the "Start Process" button to begin the e-filing process.
    - The log output will be displayed in the GUI.
  
## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.

