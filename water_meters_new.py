import pandas as pd
from datetime import datetime
import os

def process_csv_files(folder_path, desired_month):
    file_list = os.listdir(folder_path)
    meter_data_dict = {}
    start_index = 4
    for file_name in file_list:
        # Read the CSV file into a pandas DataFrame
        file_path = f'{folder_path}/{file_name}'
        df = pd.read_csv(file_path, names=['A', 'B', 'C', 'D'])

        # Get the date and time from A3
        date_time_ampm_zone = df.iloc[2, 0].split()
        date_str = date_time_ampm_zone[0]
        time_str = date_time_ampm_zone[1] + date_time_ampm_zone[2]

        if desired_month not in date_str:
            continue

        # Convert date and time strings to datetime objects
        date_time_str = f"{date_str} {time_str}"
        date_time = datetime.strptime(date_time_str, '%d-%b-%y %I:%M%p')

        # Get the total usage from B3, C3, and D3
        start_values = df.loc[start_index:, 'B']
        end_values = df.loc[start_index:, 'C']
        total_usages = df.loc[start_index:, 'D'].tolist()
        # Get the meters from col A starting from row 5
        meters = df.loc[start_index:, 'A'].tolist()

        # Add the date/total usage pair to each meter in the dictionary
        for i in range(0, len(meters)):
            meter = meters[i]
            total_usage = float(total_usages[i])
            if meter in meter_data_dict:
                meter_data_dict[meter].append((date_time, total_usage))
            else:
                meter_data_dict[meter] = [(date_time, total_usage)]

    return meter_data_dict

# Folder pointing to all the water meter data from Niagara
folder_path = '80Q - Water Meters/Water_PreviousDayUsage_files'

# Gets a dictionary containing each meter, and a list of (datetime, usage) pairs for each
result_dict = process_csv_files(folder_path, desired_month='Jul')

# Result DI
print(result_dict)
