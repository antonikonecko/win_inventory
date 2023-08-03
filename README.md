# Inventory Script and Report Parser

This repository contains two files: `inventory_script.ps1` and `parser.py`.

## `inventory_script.ps1`

This PowerShell script collects hardware and software information from a remote computer and saves it in a CSV file. The script performs the following tasks:

1. Retrieves Microsoft Office version and license keys.
2. Collects GPU information.
3. Gathers monitor details.
4. Collects computer system information such as name, IP, manufacturer, model, BIOS version, OS details, RAM, CPU, GPU, disks, system drive details, user, last reboot time, and more.
5. Exports the collected information to a CSV file on a network share.

Please make sure to adjust the `$NETWORK_SHARE_PATH` variable to the correct network share path before running the script.

## `parser.py`

This Python script processes the CSV files generated by the PowerShell script and creates an Excel workbook with formatted data. The script performs the following tasks:

1. Reads and concatenates all CSV files in the specified directory.
2. Extracts relevant columns for Microsoft Office information.
3. Creates an Excel workbook and a worksheet.
4. Populates the worksheet with data from the CSV files.
5. Applies formatting to the data based on Office version, presence of Office keys, and other criteria.
6. Saves the workbook with formatted data to the network share.

Please make sure to set the correct `NETWORK_SHARE_DIRECTORY`, `OFFICE_FILENAME`, and `REPORT_FILENAME` variables in the script before running it.
### Requirements

Before running `parser.py`, make sure you have the following Python libraries installed:

- pandas
- numpy
- openpyxl

You can install these libraries using `pip`, the Python package manager. Open a terminal or command prompt and run the following commands:

```bash
pip install pandas
pip install numpy
pip install openpyxl
```
Feel free to use and modify these scripts according to your needs. If you encounter any issues or have any suggestions, please create an issue in this repository. Happy inventorying!
