# Overview
This tool automates the conversion of CSV files into Excel workbooks (.xlsx), organizing temperature measurements into experimental days based on a specified starting time and sampling rate. It is particularly useful for processing iButton data or similar loggers that record periodic measurements.

### Key features:
- Reads multiple CSV files from a folder automatically.
- Extracts the sample rate interval and temperature units (e.g., °C, °F) directly from the CSV content.
- Validates that temperature units and sample rates are the same from all CSV files.
- Easy configuration with a JSON file.
- Groups measurements by day based on start time and sample rate
  - Configurable option to repeat the last value of a day as the first day od the next one.
- Creaters .xlsx files with columns labeled ad Day 1, Day 2, ..., Day N, with cleand and ordered data.

This script is ideal for researchers and lab technicians who routinely work with temperature loggers and need to reformat data into clean, structured Excel files for analysis.

# How to use it

### First setup
1. Install Python 3.11.x on the device
2. Download all content from the dist folder in this reposiroty. Files can be placed anywhere on the device.
3. Install the libraries specified in requeriments.txt.
   - Open a terminal on the device
   - Use the `cd` command to navigate to the folder where the files are located.
   - Install the required packages with the following command:

     `pip install -r .\requirements.txt`

### Guide
1. Place the CSV files to work with in the same folder where the ibuttons.exe and ibuttons.config are located.
   - CSV files must share the same sample rate and units. If these conditions are not meet, an error will be displayed in the console.
2. Modify ibuttons.config  set the appropriate values:
   - `StartRow` specifies the row number in the CSV file where valid temperature measurements begin. This corresponds to the time   that marks the start of the first experimental day. Any rows before this are ignored.
   - `EndRow` indicates the row number where temperature data ends. Rows beyond this point are excluded from processing.
   - `RepeatLastValues` is a boolean flag (true or false) that determines whether the last value of each experimental day should be repeated as the first value of the next day. Useful for maintaining continuity in day-to-day measurement comparisons.
3. Double-click on ibuttons.exe to start the tool automatically.
If no problems are found, the window will be closed and the excel files will be ready.

