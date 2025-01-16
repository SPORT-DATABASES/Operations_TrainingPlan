import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta, timezone

# Connect to a report generated via Smartabase
url = "https://aspire.smartabase.com/aspireacademy/live?report=TRAINING_OPERATIONS_REPORT&updategroup=true"
username = "kenneth.mcmillan"
password = "Quango76"

# Create a session
session = requests.Session()
session.auth = (username, password)

# Fetch the page
response = session.get(url)
response.raise_for_status()  # Ensure the request was successful

# Parse the HTML content
soup = BeautifulSoup(response.text, 'html.parser')

# Identify the table
tables = soup.find_all('table')
first_table = tables[0]
data = []
headers = [header.text.strip() for header in first_table.find_all('th')]
    
for row in first_table.find_all('tr')[1:]:  # Skip the header row
    cells = row.find_all(['td', 'th'])
    data.append([cell.text.strip() for cell in cells])

TrainingPlan_data = pd.DataFrame(data, columns=headers)
TrainingPlan_data = TrainingPlan_data.drop(columns=['About'])
TrainingPlan_data = TrainingPlan_data.drop_duplicates()

# Get the current date
current_date = datetime.now()
print(f"Current date: {current_date.strftime('%Y-%m-%d %H:%M:%S')}")

# Find the upcoming Sunday
days_until_sunday = (6 - current_date.weekday()) % 7  # Adjust to include Sunday if today is Thursday
next_sunday = current_date + timedelta(days=days_until_sunday)
print(f"Closest next Sunday: {next_sunday.date()}")

# Calculate the following Saturday (6 days after Sunday)
next_saturday = next_sunday + timedelta(days=6)
print(f"Upcoming week starts on: {next_sunday.date()} and ends on: {next_saturday.date()}")

next_sunday = next_sunday.date()
next_saturday = next_saturday.date()

date_column_name = 'Date'  # Replace this with the actual date column name

TrainingPlan_data[date_column_name] = pd.to_datetime(TrainingPlan_data[date_column_name], errors='coerce').dt.date

TrainingPlan_data = TrainingPlan_data[
    (TrainingPlan_data[date_column_name] >= next_sunday) &
    (TrainingPlan_data[date_column_name] <= next_saturday)
]

TrainingPlan_data = TrainingPlan_data[TrainingPlan_data['Sport'].notna()]  # Remove NaN values
TrainingPlan_data = TrainingPlan_data[TrainingPlan_data['Sport'].str.strip() != ""]  # Remove empty strings

TrainingPlan_data[date_column_name] = pd.to_datetime(TrainingPlan_data[date_column_name], errors='coerce')

# Remove timestamp from the date
TrainingPlan_data[date_column_name] = TrainingPlan_data[date_column_name].dt.date

# Create a new column 'Date_long' with the format 'ddd dd mmm yyyy'
TrainingPlan_data['Date_long'] = TrainingPlan_data[date_column_name].apply(
    lambda x: x.strftime('%a %d %b %Y') if pd.notnull(x) else None
)

TrainingPlan_data = TrainingPlan_data.sort_values(by=[date_column_name, 'Sport']).reset_index(drop=True)

# Convert columns to numeric, coercing errors to NaN
TrainingPlan_data['Start Time'] = pd.to_numeric(TrainingPlan_data['Start Time'], errors='coerce')
TrainingPlan_data['End Time'] = pd.to_numeric(TrainingPlan_data['End Time'], errors='coerce')

# Function to convert timestamps to time
def convert_to_time(timestamp_ms):
    if pd.notnull(timestamp_ms):  # Ensure the value is not NaN
        timestamp_s = timestamp_ms / 1000  # Convert milliseconds to seconds
        return datetime.fromtimestamp(timestamp_s, tz=timezone.utc).strftime('%H:%M')
    return None  # Return None for invalid or NaN values

# Apply the conversion function to the columns
TrainingPlan_data['Start Time'] = TrainingPlan_data['Start Time'].apply(convert_to_time)
TrainingPlan_data['End Time'] = TrainingPlan_data['End Time'].apply(convert_to_time)


