from functions import *

month_folder = 'POWER/2023 09 September'  
currentMonth = 'Sep'

raw_data_folder = 'Raw_Elec_Data'
input_data_path = f"{month_folder}/{raw_data_folder}"

group_plot_data_folder = 'Grouped_Elec_Plots'
output_data_path = f"{month_folder}/{group_plot_data_folder}"

input_files = os.listdir(input_data_path)

all_METERS_list = []
### Get a class for each meter to store the data
for filename in input_files:
    if filename.endswith(".xls"):
        file_path = f"{input_data_path}/{filename}"
        print(f"Extracting Data From: {filename}")
        extract_meters_into_class(file_path, all_METERS_list)

for meter in all_METERS_list:
    print(meter.dates)

# for each_meter in all_METERS_list:
#     plot_dates_vs_total_from_CLASS(each_meter, output_filename=f"{month_folder}/each_meter_data/{each_meter.name}")


# def combine_meters_in_list_to_class(grouped_METER_list, grouped_meter_name):
#     meter1 = grouped_METER_list[0]
#     grouped_off_peaks = meter1.off_peaks
#     grouped_on_peaks = meter1.on_peaks
#     grouped_weekends = meter1.weekends
#     grouped_totals = meter1.totals
#     grouped_dates = meter1.dates

#     for meter in grouped_METER_list[1:]:
#         for i in range(0, len(meter.dates)):
#             grouped_off_peaks[i] += meter.off_peaks[i]
#             grouped_on_peaks[i] += meter1.on_peaks[i]
#             grouped_weekends[i] += meter1.weekends[i]
#             grouped_totals[i] += meter1.totals[i]

#     return Meter(grouped_meter_name, grouped_dates, grouped_off_peaks, grouped_on_peaks, grouped_weekends, grouped_totals)

def combine_meters_in_list_to_class(grouped_METER_list, grouped_meter_name):
    base_meter = grouped_METER_list[0]
    dates = base_meter.dates
    lol_off_peaks = [base_meter.off_peaks]
    lol_on_peaks = [base_meter.on_peaks]
    lol_weekends = [base_meter.weekends]
    lol_totals = [base_meter.totals]

    for meter in grouped_METER_list[1:]:
        lol_off_peaks.append(meter.off_peaks)
        lol_on_peaks.append(meter.on_peaks)
        lol_weekends.append(meter.weekends)
        lol_totals.append(meter.totals)

    # print(lol_off_peaks)
    total_off_peaks = [sum(x) for x in zip(*lol_off_peaks)]
    total_on_peaks = [sum(x) for x in zip(*lol_on_peaks)]
    total_weekends = [sum(x) for x in zip(*lol_weekends)]
    total_totals = [sum(x) for x in zip(*lol_totals)]
        
    return groupedMeters(grouped_meter_name, dates, total_off_peaks, total_on_peaks, total_weekends, total_totals)

meter_groups_dict = {
    'Basement': ['D44', 'D43', 'D55'],
    'Common Areas': ['D4', 'D5', 'D6', 'D27', 'D31', 'D11', 'D20', 'D15', 'D10', 'D14', 'D17', 'D18', 'D29', 'D41', 'D42', 'D30', 'D9', 'D13', 'D28', 'D22', 'D26', 'D16', 'D19', 'D12'],            
    'Level 01-08': ['D40', 'D59', 'D39', 'D38', 'D37', 'D36', 'D63', 'D35', 'D60', 'D57', 'D62', 'D68', 'D65', 'D34', 'D33', 'D61', 'D58', 'D66', 'D67', 'D64', 'D32', 'D7', 'D8'],
    'Level 09-18': [
        'D54','D53','D52','D51','D50','D23','D24','D49','D48','D47','D46','D45','D25']
}
for key, value in meter_groups_dict.items():
    print(key, len(value))


group_classes_to_plot = []
grouped_meterslist_dict = split_meter_list_into_groups(meter_groups_dict, all_METERS_list)
for key, value in grouped_meterslist_dict.items():
    print(key, len(value))
    group_classes_to_plot.append(combine_meters_in_list_to_class(grouped_METER_list=value, grouped_meter_name=key))

# Plot each of the groups and save data to excel
# for each_group in group_classes_to_plot:
#     print(each_group.off_peaks)
#     plot_dates_vs_total_from_CLASS(each_group, output_filename=f"{output_data_path}/{each_group.name}")

# Fill out sheets showing all the meter data
# write_all_data_to_excel(all_METERS_list, 'POWER/all_power_meters_table.xlsx', month='Aug', output_sheet_name='Aug')

print(grouped_meterslist_dict)
print('\n\n')
print(group_classes_to_plot)
# for meter in group_classes_to_plot:
#     print(meter.)

write_groups_to_excel(group_classes_to_plot, 'POWER/80 Queen St - Analytics Report - Calendar Data.xlsx', currentMonth, currentMonth+'-grouped', dates_row=3)

write_all_data_grouped_to_excel(grouped_meterslist_dict, 'POWER/80 Queen St - Analytics Report - Calendar Data.xlsx', currentMonth, currentMonth, dates_row=4)
