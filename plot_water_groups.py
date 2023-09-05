import pandas as pd
import matplotlib.pyplot as plt

# Load the data from the Excel file
excel_file = 'Aug2023/80 Queen St - Water Usage Grouped - August.xlsx'
df = pd.read_excel(excel_file)

# Convert the 'Date' column to datetime
df['Date'] = pd.to_datetime(df['Date'], format='%d/%m/%Y')

# Reverse the order of the DataFrame
df = df[::-1].reset_index(drop=True)

# Define the columns to plot
columns_to_plot = ['Basement', 'Common Areas', 'Level 01 - 08', 'Level 09 - 18 ']

# Iterate through the columns and create plots
for column in columns_to_plot:
    fig, ax1 = plt.subplots(figsize=(10, 6))

    # Calculate daily changes
    daily_changes = df[column].diff()

    # Calculate cumulative sums
    cumulative_sums = daily_changes.cumsum()
    dates = df['Date']
    # Create bar plot for daily changes
    ax1.bar(dates, daily_changes, label='Daily Change', alpha=0.7)
    ax1.set_xlabel('Date')
    ax1.set_ylabel('Water Usage ($m^3$)')
    ax1.tick_params(axis='y')

    # Create a secondary axis for cumulative sums
    ax2 = ax1.twinx()
    ax2.plot(dates, cumulative_sums, label='Cumulative Sum', color='red')
    ax2.set_ylabel(f'Cumulative Sum ($m^3$)')
    ax2.tick_params(axis='y')

    ax1.set_xticks(dates)
    ax1.set_xticklabels([date.strftime("%d-%b-%y") for date in dates], rotation=45, ha='right')
    # Rotate the x-axis labels to show each date
    plt.xticks(rotation=45, ha="right")

    plt.title(f'{column} Daily Changes and Cumulative Sum')
    
    # Combine legends in upper left
    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    lines = lines1 + lines2
    labels = labels1 + labels2
    ax1.legend(lines, labels, loc='upper left')

    # Save the figure as an image file
    plt.savefig(f'{column}_plot.png', bbox_inches='tight')

    # Show the plot
    plt.show()
