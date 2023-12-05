import matplotlib.pyplot as plt
from datetime import datetime
from matplotlib.dates import DateFormatter

data = [['6/09/2023 4:00:00 am', 109333.47],
        ['6/09/2023 3:00:00 am', 109333.23],
        ['6/09/2023 2:00:00 am', 109333.00],
        ['6/09/2023 1:00:00 am', 109327.25],
        ['6/09/2023 12:00:00 am', 109306.28],
        ['5/09/2023 11:00:00 pm', 109283.73],
        ['5/09/2023 10:00:00 pm', 109264.42],
        ['5/09/2023 9:00:00 pm', 109261.54],
        ['5/09/2023 8:00:00 pm', 109261.12],
        ['5/09/2023 7:00:00 pm', 109260.62]]

# Convert date strings to datetime objects
dates = [datetime.strptime(row[0], '%d/%m/%Y %I:%M:%S %p') for row in data]

# Extract values for plotting
values = [row[1] for row in data]

# Plotting
plt.figure(figsize=(10, 6))
plt.plot(dates, values, marker='o')
plt.title('M1: 6$^{th}$ Sep at 7pm - 7$^{th}$ Sep at 4am')
plt.xlabel('Date and Time')
plt.ylabel('Meter Reading ($m^3$)')

# Customize x-axis ticks format
date_format = DateFormatter('%d$^{th}$ %I %p')
plt.gca().xaxis.set_major_formatter(date_format)

plt.xticks(rotation=45)
plt.tight_layout()

# Show the plot
plt.savefig('M1 Wack readings')
