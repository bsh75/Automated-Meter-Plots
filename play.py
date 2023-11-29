from datetime import datetime

date_string = '26-Nov-23 11:00:00 p.m.'
date_format = '%d-%b-%y %I:%M:%S %p.'

try:
    parsed_date = datetime.strptime(date_string, date_format)
    print(parsed_date)
except ValueError as e:
    print(f"Error: {e}")