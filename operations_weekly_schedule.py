import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta, timezone

# Function to adjust timestamps and convert to local time
def convert_to_time(timestamp_ms, offset_hours=11):
    try:
        if pd.notnull(timestamp_ms):
            timestamp_s = float(timestamp_ms) / 1000
            return (datetime.fromtimestamp(timestamp_s, tz=timezone.utc) - timedelta(hours=offset_hours)).strftime('%H:%M')
    except (ValueError, TypeError):
        return None

# Fetch and parse the report
session = requests.Session()
session.auth = ("kenneth.mcmillan", "Quango76")
response = session.get("https://aspire.smartabase.com/aspireacademy/live?report=TRAINING_OPERATIONS_REPORT&updategroup=true")
response.raise_for_status()

# Parse HTML and extract table
soup = BeautifulSoup(response.text, 'html.parser')
table = soup.find('table')
headers = [th.text.strip() for th in table.find_all('th')]
data = [[td.text.strip() for td in row.find_all('td')] for row in table.find_all('tr')[1:]]

# Create DataFrame and clean data
df = pd.DataFrame(data, columns=headers).drop(columns=['About'], errors='ignore').drop_duplicates()

# Filter for upcoming week
today = datetime.now()
next_sunday = today + timedelta(days=(6 - today.weekday()) % 7)
next_saturday = next_sunday + timedelta(days=6)
df['Date'] = pd.to_datetime(df['Date'], errors='coerce').dt.date
df = df[(df['Date'] >= next_sunday.date()) & (df['Date'] <= next_saturday.date())]

# Clean and sort data
df = (df.dropna(subset=['Sport'])
        .query("Sport.str.strip() != ''", engine='python')
        .assign(
            Date_long=lambda x: x['Date'].apply(lambda d: d.strftime('%a %d %b %Y') if pd.notnull(d) else None),
            Start_Time=lambda x: pd.to_numeric(x['Start Time'], errors='coerce').apply(lambda t: convert_to_time(t)),
            Finish_Time=lambda x: pd.to_numeric(x['Finish Time'], errors='coerce').apply(lambda t: convert_to_time(t))
        )
        .drop(columns=['Date Reverse'], errors='ignore')
        .sort_values(by=['Date', 'Sport', 'Coach', 'AM/PM'])
        .reset_index(drop=True)
)

# Ensure 'AM/PM' column sorts correctly and filter rows
df['AM/PM'] = pd.Categorical(df['AM/PM'], categories=['AM', 'PM'], ordered=True)
df = df[df['by'] != 'Fusion Support']

# Reorder columns to move 'Date_long' to the start
df = df[['Date_long'] + [col for col in df.columns if col != 'Date_long']]

# Drop 'Start Time' and 'Finish Time' columns
df = df.drop(columns=['Start Time', 'Finish Time'], errors='ignore')

# Replace spaces in column headers with underscores
df.columns = df.columns.str.replace(' ', '_')




