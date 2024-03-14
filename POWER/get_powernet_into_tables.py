import xlrd
from openpyxl import Workbook
from datetime import datetime, timedelta
import os

class Meter:
    '''Defines a class for each meter which contains all the information/data relating to a single meter'''
    def __init__(self, name, dates, off_peaks, on_peaks, weekends, totals):
        self.name = name
        self.dates = [dates]
        self.off_peaks = [off_peaks]
        self.on_peaks = [on_peaks]
        self.weekends = [weekends]
        self.totals = [totals]

    def __str__(self):
        # Return a string representation of the class instance
        return f"MyClass: Name={self.name}, Date={self.dates}, Off Peak={self.off_peaks}, On Peak={self.on_peaks}, Weekend={self.weekends}, Total={self.totals}"


def add_to_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename, meters_list):
    """Adds the information on meter {name} to that meters class"""
    for meter in meters_list:
        if meter.name == name:
            meter.dates.append(date)
            meter.off_peaks.append(off_p_usage)
            meter.on_peaks.append(on_p_usage)
            meter.weekends.append(weeknd_usage)
            meter.totals.append(total_usage)


def extract_date_times_from_string(input_string):
    '''Extracts the date from the cell associated with it as a string'''
    # Define the format of the date and time in the string
    date_format = "%B %d, %Y %I:%M %p"

    # Extract the dates and times from the input string
    date_time_str = input_string.split('From:', 1)[1]

    # Split the extracted string to separate the start and end date_time_str
    start_date_time_str, end_date_time_str = date_time_str.split(' to:', 1)

    # Parse the date_time strings into datetime objects
    start_date_time = datetime.strptime(start_date_time_str, date_format)
    end_date_time = datetime.strptime(end_date_time_str.strip(), date_format)

    # Check if the period is exactly one day
    one_day_difference = (end_date_time - start_date_time) == timedelta(days=1)

    if one_day_difference:
        return start_date_time
    else:
        return None


def extract_PN_meters_to_list(excel_filename, meters_class_list):
    """Main function for processing a single xls file where each file contains every meters data over 1 day:
    Function will create a new meter class for each new meter, and add the data to it if a meter already exists """
    name_criteria_string = 'Energy User:'  # Replace with the criteria you are looking for
    date_criteria_string = 'For Electric Usage From:'
    off_p_string = 'Off-Peak'
    on_p_string = 'On-Peak'
    weeknd_string = 'Weekend'
    '''Looks in a single file (corresponding to a sinlge date) and creates or adds to existing meter class for each meter found'''
    try:
        # Need to use xlrd for xls (old excel) files
        workbook = xlrd.open_workbook(excel_filename)
        sheet = workbook.sheet_by_index(0)  # Assuming you want to work with the first sheet
    except xlrd.biffh.XLRDError:
        print(f"{excel_filename} is not a xls file, SKIPPING... ")
        return
    # Iterate through all cells to find relevant information
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell_value = str(sheet.cell_value(row, col))
            # Get the name of meter
            if cell_value.startswith(name_criteria_string):
                name = cell_value.replace(name_criteria_string, '')
            # Get the start of the period as the date for usage
            elif cell_value.startswith(date_criteria_string):
                date = extract_date_times_from_string(cell_value)
                # target_datetime = datetime(2023, 9, 26, 0, 0, 0)
                if not date:
                    print(f"WARNING!! {name} period is not 1 day, SKIPPING file...... ({excel_filename})")
                    return
                # if date == target_datetime:
                #     date = datetime(2023, 8, 26, 0, 0, 0)
                # print(date)
            # Get off peak, on peak, and weekend breakdown
            elif cell_value == off_p_string:
                off_p_usage = sheet.cell_value(row, col+2)
            elif cell_value == on_p_string:
                on_p_usage = sheet.cell_value(row, col+2)
            elif cell_value == weeknd_string:
                # Section should only occur once the weekend value (last bit of info) has been found
                weeknd_usage = sheet.cell_value(row, col+2)
                total_usage = off_p_usage+on_p_usage+weeknd_usage
                if not any(meter.name == name for meter in meters_class_list):
                    meters_class_list.append(Meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage))
                else:
                    add_to_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename, meters_class_list)


def create_xlsx_from_M_class(obj, output_path):
    '''Creates the plots from data associated with a single class'''
    # Get the dates and totals from the object
    dates = obj.dates
    off_ps = obj.off_peaks
    on_ps = obj.on_peaks
    wknds = obj.weekends
    totals = obj.totals
    
    # # # Convert the dates to datetime objects
    dates = [datetime.strftime(date, "%d-%b-%y") for date in dates]

    # Save the data used in the plots
    wb = Workbook()
    sheet = wb.active
    sheet.append([f'{obj.name}'])
    dates.insert(0, "Dates")
    off_ps.insert(0, "Off Peak")
    on_ps.insert(0, "On Peak")
    wknds.insert(0, "Weekend")
    totals.insert(0, "Total")
    sheet.append(dates)
    sheet.append(off_ps)
    sheet.append(on_ps)
    sheet.append(wknds)
    sheet.append(totals)
    # for i in range(0, len(dates)):
    #     sheet.append([dates_strings[i], off_ps[i], on_ps[i], wknds[i], totals[i]])

    if not os.path.exists(output_path):
        os.makedirs(output_path)

    wb.save(f"{output_path}/{obj.name}.xlsx")

# month_folder = 'POWER/2023 12 December'  

# # month_folder = 'POWER/2024 01 January'  
    
month_folder_input = input("Enter the name of the month folder eg. '2023 11 November': ")

month_folder = f'POWER/{month_folder_input}'

raw_data_folder = input("Enter the name of the raw data folder which contains all the powernet raw data eg 'Raw_PowerNET_Data' (NOTE: If any Encoding error occurs, you must open each xls file and 'save as' to a new folder, then rerun this code using the new folder: ")

if raw_data_folder == '':
    raw_data_folder = 'Raw_PowerNET_Data'

input_data_path = f"{month_folder}/{raw_data_folder}"

input_files = os.listdir(input_data_path)

meter_plot_data_folder = 'each_meter_data'
meter_output_data_path = f"{month_folder}/{meter_plot_data_folder}"

all_METERS_list = []
all_METERS_dict = {}
### Get a class for each meter to store the data
for filename in input_files:
    if filename.endswith(".xls"):
        file_path = f"{input_data_path}/{filename}"
        print(f"Extracting Data From: {filename}")
        # extract_meters_into_class(file_path, all_METERS_list)
        extract_PN_meters_to_list(file_path, all_METERS_list)

for meter in all_METERS_list:
    create_xlsx_from_M_class(meter, meter_output_data_path)


