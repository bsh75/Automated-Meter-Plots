import xlrd
import openpyxl
from openpyxl import Workbook
from datetime import datetime, timedelta
import os
import matplotlib.pyplot as plt
import struct

METERS = []
NONXLSFILES = 0
ERRORS = 0

class Meter:
    def __init__(self, name, dates, off_peaks, on_peaks, weekends, totals, in_files):
        self.name = name
        self.dates = [dates]
        self.off_peaks = [off_peaks]
        self.on_peaks = [on_peaks]
        self.weekends = [weekends]
        self.totals = [totals]
        self.in_files = [in_files]

    def __str__(self):
        # Return a string representation of the class instance
        return f"MyClass: Name={self.name}, Date={self.dates}, Off Peak={self.off_peaks}, On Peak={self.on_peaks}, Weekend={self.weekends}, Total={self.totals}"
    

def extract_date_times_from_string(input_string):
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
        return start_date_time.strftime("%d-%b-%y")
    else:
        return None

# def get_date(string, start):
#     """Gets the date from a substring"""
#     date_line = string.replace(start, '')
#     monthDay_rest = date_line.split(',', 1)
#     month_day =  monthDay_rest[0]
#     year = monthDay_rest[1][:4]
#     return month_date

def find_matching_values(excel_filename, name_criteria_string, date_criteria_string, off_p_string, on_p_string, weeknd_string):
    try:
        workbook = xlrd.open_workbook(excel_filename)
        sheet = workbook.sheet_by_index(0)  # Assuming you want to work with the first sheet
    except xlrd.biffh.XLRDError:
        print(f"{excel_filename} is not a xls file, SKIPPING... ")
        return
    # except struct.error as err:
    #     print(f"{excel_filename} ERROR, SKIPPING... ")
    #     return
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            cell_value = str(sheet.cell_value(row, col))
            if cell_value.startswith(name_criteria_string):
                name = cell_value.replace(name_criteria_string, '')
            elif cell_value.startswith(date_criteria_string):
                date = extract_date_times_from_string(cell_value)
                if not date:
                    print(f"{excel_filename} period is not 1 day, SKIPPING file......")
                    return
            elif cell_value == off_p_string:
                off_p_usage = sheet.cell_value(row, col+2)
            elif cell_value == on_p_string:
                on_p_usage = sheet.cell_value(row, col+2)
            elif cell_value == weeknd_string:
                # Section should only occur once the weekend value has been found
                weeknd_usage = sheet.cell_value(row, col+2)
                total_usage = off_p_usage+on_p_usage+weeknd_usage
                if not any(meter.name == name for meter in METERS):
                    create_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename)
                else:
                    add_to_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename)
                # print(f"{name}\t{date}\t{off_p_usage}\t{on_p_usage}\t{weeknd_usage}\t{total_usage}\t")

def create_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename):
    """Initialises a new meter and adds it to list"""
    METERS.append(Meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename.replace(folder, '')))

def add_to_meter(name, date, off_p_usage, on_p_usage, weeknd_usage, total_usage, excel_filename):
    """Adds the information on meter {name} to that meters class"""
    for meter in METERS:
        if meter.name == name:
            meter.dates.append(date)
            meter.off_peaks.append(off_p_usage)
            meter.on_peaks.append(on_p_usage)
            meter.weekends.append(weeknd_usage)
            meter.totals.append(total_usage)
            meter.in_files.append(excel_filename.replace(folder, ''))


def plot_dates_vs_totals(obj, output_filename):
    # Get the dates and totals from the object
    dates = obj.dates
    off_ps = obj.off_peaks
    on_ps = obj.on_peaks
    wknds = obj.weekends
    totals = obj.totals

    # Convert the dates to datetime objects
    dates = [datetime.strptime(date, "%d-%b-%y") for date in dates]

    # Plot the data
    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Plot the usage data on the left y-axis
    p3 = ax1.bar(dates, wknds, color='#6794a7', width=0.6, label='Weekend')
    p1 = ax1.bar(dates, off_ps, bottom=wknds, color='#7ad2f6', width=0.6, label='Off Peak')
    p2 = ax1.bar(dates, on_ps, bottom=[sum(x) for x in zip(wknds, off_ps)], color='#014d64', width=0.6, label='On Peak')

    ax1.set_xlabel('Dates')
    ax1.set_ylabel('Usage')
    ax1.set_title(f'Dates vs. Usage for {obj.name}')
    ax1.grid(axis='y', linestyle='--', alpha=0.7)

    # Create a new y-axis for plotting the accumulation data
    ax2 = ax1.twinx()

    # Calculate the accumulation data (cumulative sum of totals)
    accumulations = [sum(totals[:i + 1]) for i in range(len(totals))]

    # Plot the accumulation data on the right y-axis
    ax2.plot(dates, accumulations, color='#90353B', marker='o', label='Accumulation')
    ax2.set_ylabel('Accumulation')

    # Combine the handles for bars and lines for a single legend
    handles, labels = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(handles + handles2, labels + labels2, loc='upper left')

    # Rotate x-axis labels for better visibility
    ax1.set_xticks(dates)
    ax1.set_xticklabels([date.strftime("%d-%b-%y") for date in dates], rotation=45, ha='right')

    # Save the plot
    plt.tight_layout()
    output_filename = output_filename.replace('.', '_')
    print(f"Saving image file with name: {output_filename}")
    plt.savefig(output_filename + '.png')
    plt.close()

    # Save the data used in the plots
    wb = Workbook()
    sheet = wb.active
    sheet.append(["Date", "Off Peak", "On Peak", "Weekend", "Total"])
    for i in range(0, len(dates)):
        sheet.append([dates[i], off_ps[i], on_ps[i], wknds[i], totals[i]])
    wb.save(f"{output_filename}.xlsx")

# Example usage:
folder = '80Q - Power Meters - Powernet Data/'  # Replace with the path to your Excel file
name_start = 'Energy User:'  # Replace with the criteria you are looking for
date_start = 'For Electric Usage From:'
off_peak = 'Off-Peak'
on_peak = 'On-Peak'
weekend = 'Weekend'


file_list = os.listdir('80Q - Power Meters - Powernet Data')
bad_files = []

for filename in file_list[0:-1]:
    filepath = f"{folder}{filename}"
    print(f"Processing file {filename}")
    if (filename == 'Bills_DeloitteFloorDBs_2023_07_01.xls') or (filename == 'Bills_TotalMechanical_2023_07_01.xls'): 
        bad_files.append(filename)
        print("SHIIIIIIIIT")
    else:
        all_meters = find_matching_values(filepath, name_start, date_start, off_peak, on_peak, weekend)

print(f"Total of {len(METERS)}")
print(f"Bad files: {bad_files}")
single_date_files = []
for meter in METERS:
    # print(meter.dates)
    if len(meter.dates) == 1:
        for file in meter.in_files:
            if file not in single_date_files:
                single_date_files.append(file)            
    else:
        plot_dates_vs_totals(meter, f"{folder}Plot Data/{meter.name}")

print(f"{len(single_date_files)} files with only one date: {single_date_files}")
# plot_dates_vs_totals(METERS[0])