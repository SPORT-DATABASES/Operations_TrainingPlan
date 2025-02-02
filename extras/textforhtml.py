import pandas as pd
from datetime import datetime, timedelta
from IPython.display import HTML

import requests
from bs4 import BeautifulSoup
import pandas as pd
from datetime import datetime, timedelta, timezone
import numpy as np

import os
import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import NamedStyle
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

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

# Replace spaces in column headers with underscores
df.columns = df.columns.str.replace(' ', '_')

# Add the "Group" column
df['Group'] = df.apply(
    lambda row: f"{row['Sport']}-{row['Coach']}" 
    if row['Sport'] == row['Training_Group'] 
    else f"{row['Sport']}-{row['Training_Group']}-{row['Coach']}", 
    axis=1
)

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
            Start_Time=lambda x: pd.to_numeric(x['Start_Time'], errors='coerce').apply(lambda t: convert_to_time(t)),
            Finish_Time=lambda x: pd.to_numeric(x['Finish_Time'], errors='coerce').apply(lambda t: convert_to_time(t))
        )
        .drop(columns=['Date_Reverse'], errors='ignore')
        .sort_values(by=['Date', 'Sport', 'Coach', 'AM/PM'])
        .reset_index(drop=True)
)

# Ensure 'AM/PM' column sorts correctly and filter rows
df['AM/PM'] = pd.Categorical(df['AM/PM'], categories=['AM', 'PM'], ordered=True)
df = df[df['by'] != 'Fusion Support']

# Reorder columns to move 'Date_long' to the start
df = df[['Date_long'] + [col for col in df.columns if col != 'Date_long']]


# Remove Friday data
df = df[~df['Day_AM/PM'].str.contains('Friday')]

# Create a mapping of Training_Group to Coach
group_coach_map = df.groupby('Training_Group')['Coach'].first()

def format_session(group):
    venues = []
    times = set()
    for _, row in group.iterrows():
        venues.append(row['Venue'])
        times.add(f"{row['Start_Time']}-{row['Finish_Time']}")
    venue_str = ' + '.join(venues)
    time_str = list(times)[0]
    return f"{venue_str}\
{time_str}"

# Group and pivot data
grouped = df.groupby(['Sport', 'Training_Group', 'Day_AM/PM']).apply(format_session)
grouped = grouped.reset_index()
grouped.columns = ['Sport', 'Training_Group', 'Day_AM/PM', 'Session']

pivot_df = pd.pivot_table(
    grouped,
    values='Session',
    index=['Sport', 'Training_Group'],
    columns=['Day_AM/PM'],
    aggfunc='first',
    fill_value=' '
)

# Get current week dates (excluding Friday)
today = datetime.now()
start_of_week = today - timedelta(days=today.weekday() + 1)
day_names = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Saturday']
dates = [(start_of_week + timedelta(days=i if i < 5 else i+1)).strftime('%Y-%m-%d') for i in range(6)]

# Create HTML with embedded CSS - added cell divider
html = """
<style>
    table { 
        border-collapse: collapse; 
        width: 100%;
        table-layout: fixed;
        font-family: Arial, sans-serif;
        font-size: 12px;
        margin: 20px 0;
    }
    th, td { 
        border: 1px solid black; 
        padding: 0; 
        text-align: center;
        overflow: hidden;
    }
    th.date { 
        background-color: #f2f2f2;
        font-size: 13px;
        width: 15%;
        padding: 10px;
    }
    th.sport-header {
        background-color: #f2f2f2;
        width: 7%;
        padding: 10px;
    }
    th.group-header {
        background-color: #f2f2f2;
        width: 10%;
        padding: 10px;
    }
    th.ampm { 
        background-color: #e6e6e6;
        font-size: 12px;
        width: 7.5%;
        padding: 10px;
    }
    td { 
        white-space: pre-line; 
        vertical-align: top; 
        text-align: left;
        line-height: 1.3;
        height: 50px;
    }
    .venue {
        color: #444;
        font-weight: bold;
        font-size: 11px;
        padding: 8px;
        display: block;
        border-bottom: 1px solid #ccc;
        background-color: #f8f8f8;
    }
    .time {
        color: #666;
        font-size: 11px;
        padding: 8px;
        display: block;
    }
    .sport-cell {
        background-color: #f2f2f2;
        font-weight: bold;
        font-size: 11px;
        padding: 10px;
    }
    .group-cell {
        background-color: #f8f8f8;
        padding: 10px;
    }
    .coach-name {
        color: #666;
        font-style: italic;
        font-size: 11px;
        margin-top: 4px;
        display: block;
    }
    tr:nth-child(even) td {
        background-color: #f9f9f9;
    }
    .session-cell {
        padding: 0;
        margin: 0;
    }
</style>
<table>
    <thead>
    <tr>
        <th class="sport-header" rowspan="2">Sport</th>
        <th class="group-header" rowspan="2">Group<br>&<br>Coach</th>
"""

# Add date headers spanning AM/PM
for date, day in zip(dates, day_names):
    html += f"""
        <th class="date" colspan="2">{day}<br>{date}</th>
    """

html += """
    </tr>
    <tr>
"""

# Add AM/PM headers
for _ in range(6):
    html += """
        <th class="ampm">AM</th>
        <th class="ampm">PM</th>
    """

html += """
    </tr>
    </thead>
    <tbody>
"""

# Add data rows
current_sport = None
for (sport, group) in pivot_df.index:
    html += "<tr>"
    
    # Add sport column (only if it changed)
    if sport != current_sport:
        sport_rowspan = len(pivot_df.loc[sport])
        html += f'<td class="sport-cell" rowspan="{sport_rowspan}">{sport}</td>'
        current_sport = sport
    
    # Add group column with coach name
    coach = group_coach_map.get(group, "")
    html += f'<td class="group-cell">{group}<span class="coach-name">{coach}</span></td>'
    
    # Add data cells
    for day in day_names:
        am_key = f"{day} AM"
        pm_key = f"{day} PM"
        am_val = pivot_df.get(am_key, {}).get((sport, group), " ")
        pm_val = pivot_df.get(pm_key, {}).get((sport, group), " ")
        
        def format_cell(content):
            if content.strip() == " ":
                return " "
            parts = content.split('\n')  # Corrected the separator to '\n' for line breaks
            if len(parts) >= 2:
                return f"""<div class="session-cell"><span class="venue">{parts[0]}</span><span class="time">{parts[1]}</span></div>"""
            elif len(parts) == 1:
                return f"""<div class="session-cell"><span class="venue">{parts[0]}</span></div>"""
            return " "
            
        html += f'<td class="session-cell">{format_cell(am_val)}</td><td class="session-cell">{format_cell(pm_val)}</td>'
    html += "</tr>"

html += """
    </tbody>
</table>
"""

# Save HTML file
with open('schedule_divided_cells.html', 'w') as f:
    f.write(html)


from IPython.display import HTML, display
# Display the HTML
display(HTML(html))

print("Reverted to previous version with:")
print("- Simple dividing line between venue and time")
print("- Light gray background for venue section")
print("- Original cell spacing and padding")
print("\
Saved as 'schedule_divided_cells.html'")


from weasyprint import HTML, CSS
from weasyprint.text.fonts import FontConfiguration

# Configure fonts
font_config = FontConfiguration()

# Create custom CSS with even narrower sport/group columns
css = CSS(string='''
    @page {
        size: 1500px 1240px;
        margin: 0;
    }
    
    table { 
        width: 1500px !important;
        font-size: 11px !important;
        margin: 0 !important;
    }
    
    th.date { 
        width: 13% !important;
        font-size: 13px !important;
        padding: 8px !important;
    }
    
    th.sport-header {
        width: 5.4% !important;  /* Reduced by 10% from 6% */
        padding: 8px !important;
    }
    
    th.group-header {
        width: 7.2% !important;  /* Reduced by 10% from 8% */
        padding: 8px !important;
    }
    
    th.ampm { 
        width: 6.5% !important;
        padding: 8px !important;
    }
    
    td { 
        padding: 8px !important;
    }
    
    .venue {
        padding: 6px !important;
    }
    
    .time {
        padding: 6px !important;
    }
''', font_config=font_config)

# Convert HTML to PDF
HTML(filename='schedule_divided_cells.html').write_pdf(
    'schedule_final.pdf',
    stylesheets=[css],
    font_config=font_config
)

# Create matching PNG
import imgkit

options = {
    'format': 'png',
    'width': 1500,
    'height': 1240,
    'quality': 100
}

imgkit.from_file('schedule_divided_cells.html', 'schedule_final.png', options=options)

print("Created final versions with adjusted column widths:")
print("- Sport column reduced to 5.4% (10% smaller)")
print("- Training group column reduced to 7.2% (10% smaller)")
print("- All other dimensions maintained")
print("- Created both formats:")
print("  * schedule_final.pdf")
print("  * schedule_final.png")