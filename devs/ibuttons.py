'''
Script Name:   ibuttons.py
Description:   Builds a macro from csv data
Author:        Itzel Torres (https://github.com/itzelbtl/Ibuttons_funciones)
Updated by:    Erick Gómez
Date Created: August 2023
Last Modified: June 2025
Version:       2.0.0

Overview: 

This program processes temperature data captured by iButtons and stored in CSV files, generating a consolidated .xlsx file. Each temperature measurement is timestamped with the corresponding date and time it was recorded.
The generated Excel file includes:
    The raw data (timestamp and temperature) exactly as found in the original CSV files.
    The processed data, where only the temperature values are grouped by experimental day.

Key features:
    -The starting point of the first experimental day is defined by the startRow parameter in the JSON configuration file. This indicates the visual row in the CSV where the first experimental day begins.
    -The endRow parameter specifies the last row of interest in the dataset.
    -If the RepeatLastValue flag is set to true, each experimental day will end with the same value that starts the following day. If false, each day's data ends normally without repetition.

Data validation:
    -All CSV files must have the same sample rate.
    -All CSV files must use the same temperature units (e.g., °C or °F).

The script is designed to be portable and can be placed in any folder alongside the corresponding CSV files to process them automatically.
'''
#-------------------- libraries used --------------------
import os 
import csv
import re
import json
from openpyxl import Workbook, load_workbook
from tqdm import tqdm
from time import sleep
import sys

class MacroHandler():
    def __init__(self):
        self.startRow = None
        self.endRow = None
        self.repeatLastValue = None
        self.validateSampleRate = []
        self.validateUnits = []
        self.sampleRate = None
        self.units = None
        self.banner = ''' 
██╗██████╗ ██╗   ██╗████████╗████████╗ ██████╗ ███╗   ██╗███████╗
██║██╔══██╗██║   ██║╚══██╔══╝╚══██╔══╝██╔═══██╗████╗  ██║██╔════╝
██║██████╔╝██║   ██║   ██║      ██║   ██║   ██║██╔██╗ ██║███████╗
██║██╔══██╗██║   ██║   ██║      ██║   ██║   ██║██║╚██╗██║╚════██║
██║██████╔╝╚██████╔╝   ██║      ██║   ╚██████╔╝██║ ╚████║███████║
╚═╝╚═════╝  ╚═════╝    ╚═╝      ╚═╝    ╚═════╝ ╚═╝  ╚═══╝╚══════
        '''
        #self.path = os.path.dirname(os.path.abspath(__file__))
        if getattr(sys, 'frozen', False):
            self.path = os.path.dirname(sys.executable)
        else:
            self.path = os.path.dirname(os.path.abspath(__file__))
        print(self.banner)
        self.get_config()
        
    #-------------------- retrieves all parameters from config file --------------------
    def get_config(self):
        config_path = os.path.join(self.path, 'ibuttons.config') #use the path where the script is located and join it with the config file
        print(f'path {config_path}')
        try:
            with open(config_path, 'r') as f:
                config = json.load(f) #save all parameters in config
            required_keys = ["StartRow", "EndRow", "RepeatLastValue"] #all necessary information
            for key in required_keys:
                if key not in config:
                    raise KeyError(f"Missing required key '{key}' in config file.")
                
            # validates data types
            if not isinstance(config["StartRow"], int):
                raise TypeError("StartRow must be an integer.")
            if not isinstance(config["EndRow"], int):
                raise TypeError("EndRow must be an integer.")
            if not isinstance(config["RepeatLastValue"], bool):
                raise TypeError("RepeatLastValue must be a boolean.")
            
            #save values into attributes
            self.startRow = config.get("StartRow", 0) - 1 #retreive startRow value, is necesarry to substrac 1 to match the csv files row
            self.endRow = config.get("EndRow", 0) - 1 #retreive endRow value, is necesarry to substrac 1 to match the csv files row
            self.repeatLastValue = config["RepeatLastValue"] #retreive RepeatLastValue value
        except (FileNotFoundError, json.JSONDecodeError):
            raise FileExistsError("Couldn't open or parse config file.")
        except Exception as e:
            raise RuntimeError(f"Config validation error: {e}")
    
    #-------------------- retrieves sample rate and units from csv file --------------------
    def get_sample_rate(self, reader):
        for row in reader: #search in every full-row
            for cell in row: #search in every individual element inside the row
                if "Sample Rate" in cell:
                    match = re.search(r'(\d+)', cell)
                    if match:
                        return int(match.group(1)) #return sample rate as integer
        raise ValueError("Sample rate not found")
    
    def get_units(self, reader):
        pattern = re.compile(r'[-+]?\d+(?:\.\d+)?\s*(°[CF])') 
        for row in reader: #search in every full-row
            for cell in row: #search in every individual element inside the row
                if "High Temperature Alarm:" in cell or "Low Temperature Alarm:" in cell:
                    match = pattern.search(cell)
                    if match:
                        return match.group(1) #return units 
        raise ValueError("Units not found")

    # -------------------- retreives all the information needed (and specified ) from csv file --------------------
    def get_data_from_csv(self, full_path: str):
        clean_data = []
        #opens csv file located in the full_path, newline assures a propper reading with EOL, latin1 is used for a special characters
        with open(full_path, newline='', encoding='latin1') as csvfile: 
            lines = list(csv.reader(csvfile, delimiter=',')) #reads all the content using coma as delimiter, convert object into a list
        sample_rate = self.get_sample_rate(lines) #retreive the sample rate per csv file
        self.validateSampleRate.append(sample_rate) #save every sample rate value
        units = self.get_units(lines) #retreive the sample rate per csv file
        self.validateUnits.append(units) #save every sample rate value
        for idx, row in enumerate(lines): 
            #only rows within the margin [startRow, endRow] are processed 
            if self.startRow <= idx <= self.endRow:
                try:
                    timestamp = row[0] #retreive date/time value 
                    value = float(row[2]) #retreive numeric value
                    clean_data.append([timestamp, value]) #save information into a clean list
                except Exception as e:
                    print(f'An error happened with row {row}: {e}')
        return clean_data #return the list with processed data
    
    #-------------------- from information retreive, build the excel file --------------------   
    def build_xlsx_file(self, dicti: dict):
        #iterate through the dictionary, the key is used to create the file, the value is the information to be saved
        for key, value in dicti.items():
            excel_filename = os.path.join(self.path, f'{key}_macro.xlsx') #adds _macro to the name
            self.save_raw_data(filename=excel_filename, data_list=value) #save the raw data
            self.save_by_experimental_days(filename=excel_filename, data_list= value) #save the processed values
    
    #-------------------- takes raw data and saves it --------------------
    def save_raw_data(self, filename, data_list):
        #checks if the file exist, then it gets activate
        if os.path.exists(filename): 
            wb = load_workbook(filename)
            ws = wb.active
        else:
            wb = Workbook()
            ws = wb.active
        ws.title = "Data" #the tab is called Data
        ws.append(["Original CSV values"]) #first row with this string
        ws.append([f'Units: {self.units}']) #second row with the units
        ws.append(['Date/Time', 'Value']) #third row with the title of every column
        for line in data_list:
            ws.append(line) #append every value in the list to the xsxl file
        wb.save(filename) #save all information into the file
        
     #-------------------- takes raw data, separe it into experimental days and then saves only the measurements in a specific order--------------------
    def save_by_experimental_days(self, filename, data_list):       
        records_per_day = int((24*60)/self.sampleRate) #calculate how much measurements are done per day considering the sample rate
        chunks = []
        for i in range(0, len(data_list), records_per_day):
            chunks.append(data_list[i:i+records_per_day]) #creates a list with data devided by blocks with the respective amount of records per day 
            
        if self.repeatLastValue: #if repeatLastValue is true within the config file, the last value in every block is the first of the next one
            for chunk in range(1, len(chunks)):
                extra_data = chunks[chunk][0]
                chunks[chunk-1].append(extra_data)
        
        wb = load_workbook(filename)
        ws = wb.active
        start_col = ws.max_column + 2 #ws.max_column represents the first available column to use. It leaves a blank column between raw and processed data
        ws.cell(row=1, column=start_col, value = 'Processed values') #header for this new section
        #iterate over every block
        for day_index, chunk in enumerate(chunks, start=1):
            col = start_col + (day_index-1)
            ws.cell(row = 3, column=col, value = f'Day {day_index}') #writes Day N as header
            #every measurement is written from row 4 onwards
            for row_offset, value in enumerate(chunk):
                ws.cell(row = 4+row_offset, column=col, value = value[1]) #value[1] is the temperature
        wb.save(filename) #saves excel file
    
    def main(self):
        list_with_all_data = {}
        files_in_path = os.listdir(self.path) #list all files inside the path
        csv_files = [file for file in files_in_path if file.lower().endswith('.csv')]
        for file in tqdm(csv_files, desc="Processing CSV files"):
            full_path = os.path.join(self.path, file)  # create the full path
            data = self.get_data_from_csv(full_path=full_path)  # retrieve data
            file_key = os.path.splitext(file)[0]  # remove .csv extension
            list_with_all_data[file_key] = data  # store in dict
            sleep(0.25)
        if len(set(self.validateSampleRate)) == 1: #evaluate if all the sample rates are the same
            self.sampleRate = self.validateSampleRate[0]
        else:
            raise ValueError(f'Different sample rates found in csv files')
        if len(set(self.validateUnits)) == 1: #evaluates all the units are the same
            self.units = self.validateUnits[0]
        else:
            raise ValueError(f'Different units found in csv files: {self.validateUnits}')
        self.build_xlsx_file(dicti=list_with_all_data) #create the xsxl files
        
    
    
def main():
    macro = MacroHandler()
    macro.main()

if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        print('--- Caught Exception ---')
        print(e)
        print('----------------------------')
        input("Press any key to close the program...")
    finally:
        sys.exit