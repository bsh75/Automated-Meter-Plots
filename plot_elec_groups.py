import pandas as pd
import matplotlib.pyplot as plt
from datetime import datetime

# Load the Excel file
excel_file = 'Electrical Meters - Grouped.xlsx'

# Read the Excel file, skipping rows with header information
df = pd.read_excel(excel_file, skiprows=3)
print(df)
# Reverse the order of rows to have the latest date at the top
df = df[::-1].reset_index(drop=True)
print(df)
# Define the categories and subcategories (excluding "Total")
categories = ['Basement', 'Common Areas', 'Levels 01 - 08', 'Levels 09 - 18']
subcategories = ['Off Peak', 'On Peak', 'Weekend']

# Create empty DataFrames to store subcategory data
subcategory_data = {}

# Convert the 'Date' column to datetime objects
df['Date'] = pd.to_datetime(df['Date'])

dates = df['Date']
# print(dates)
basement_data = df[df.columns[1:4]]
basement_total = df[df.columns[4]]
common_data = df[df.columns[5:8]]  # Columns F to H
common_total = df[df.columns[8]]
low_level_data = df[df.columns[9:12]]  # Columns K to M
low_level_total = df[df.columns[12]]
high_level_data = df[df.columns[13:16]]  # Columns P to R
high_level_total = df[df.columns[16]]

# Function to create stacked bar plots with accumulating total in red
def create_stacked_bar_plot(data, total_data, category_name):
    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Create a stacked bar plot for the subcategories with specified colors
    colors = ['#6794a7', '#7ad2f6', '#014d64']
    data.plot(kind='bar', stacked=True, width=0.8, ax=ax1, color=colors)

    # Set labels and title
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Usage  (kwH)')
    ax1.set_title(f'{category_name} Energy Usage')

    # Format the x-axis labels to show only the date
    date_labels = [d.strftime('%Y-%m-%d') for d in dates]
    ax1.set_xticks(range(len(dates)))
    ax1.set_xticklabels(date_labels, rotation=45, ha='right')

    # Create a secondary axis for total accumulation
    ax2 = ax1.twinx()
    accumulating_total = total_data.cumsum()  # Calculate accumulating total
    ax2.plot(range(len(dates)), accumulating_total, color='#90353B',  marker='o', label='Accumulating Total')
    ax2.set_ylabel('Total Accumulation (kwH)')

    # Add legend for total accumulation
    # Combine the handles for bars and lines for a single legend
    handles, labels = ax1.get_legend_handles_labels()
    handles2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(handles + handles2, labels + labels2, loc='upper left')

    # Save the figure as an image file with the elec prefix
    plt.savefig(f'elec_{category_name}_usage.png', bbox_inches='tight')

    # Don't show the plot
    plt.close()

# Create stacked bar plots for each category with accumulating total in red
create_stacked_bar_plot(basement_data, basement_total, 'Basement')
create_stacked_bar_plot(common_data, common_total, 'Common Areas')
create_stacked_bar_plot(low_level_data, low_level_total, 'Levels 01 - 08')
create_stacked_bar_plot(high_level_data, high_level_total, 'Levels 09 - 18')
