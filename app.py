import streamlit as st
from io import BytesIO
from io import StringIO
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import shutil
import os

# ---- Page Configuration ----
st.set_page_config(
    page_title="Operations - Weekly Training Plan",
    layout="wide",  # Use wide layout
    initial_sidebar_state="expanded"
)

# ---- Custom CSS to hide default Streamlit elements and reduce top spacing ----
hide_streamlit_style = """
<style>
/* Hide the default hamburger menu and footer */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
"""
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

# Function to adjust timestamps and convert to local time
def convert_to_time(timestamp_ms, offset_hours=11):
    try:
        if pd.notnull(timestamp_ms):
            timestamp_s = float(timestamp_ms) / 1000
            return (datetime.fromtimestamp(timestamp_s, tz=timezone.utc) - timedelta(hours=offset_hours)).strftime('%H:%M')
    except (ValueError, TypeError):
        return None

# Function to ensure all expected columns are present in the pivot DataFrame
def ensure_all_columns(pivot_df, day_order):
    return pivot_df.reindex(columns=['Sport', 'Training_Group'] + day_order, fill_value=' ')

# Function to format session information for grouping
def format_session(group):
    venue_time_pairs = []
    for _, row in group.iterrows():
        venue = str(row['Venue']) if pd.notnull(row['Venue']) else ''  # Ensure Venue is a string
        start_time = str(row['Start_Time']) if pd.notnull(row['Start_Time']) else ''
        finish_time = str(row['Finish_Time']) if pd.notnull(row['Finish_Time']) else ''
        time = f"{start_time}-{finish_time}" if start_time or finish_time else ''

        if venue or time:  # Include only non-empty venue or time
            venue_time_pairs.append((start_time, f"{venue} {time}".strip()))

    # Sort the venue-time pairs by the start time
    sorted_venue_time_pairs = sorted(
        venue_time_pairs, 
        key=lambda x: datetime.strptime(x[0], '%H:%M') if x[0] else datetime.min
    )

    # Extract only the formatted strings
    sorted_sessions = [pair[1] for pair in sorted_venue_time_pairs]

    return ' + '.join(filter(None, sorted_sessions))


# Function to paste filtered data into the Template sheet
def paste_filtered_data_to_template(pivot_df, workbook, sport, training_group, start_cell):
    template_sheet = workbook["Template"]
    filtered_row = pivot_df[(pivot_df['Sport'] == sport) & (pivot_df['Training_Group'] == training_group)]

    if not filtered_row.empty:
        values_to_paste = filtered_row.iloc[0, 2:].tolist()  # Exclude Sport and Training_Group columns
        col_letter, row_num = start_cell[0], int(start_cell[1:])
        start_col_idx = ord(col_letter.upper()) - ord("A") + 1

        for col_idx, value in enumerate(values_to_paste, start=start_col_idx):
            cell = template_sheet.cell(row=row_num, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Function to paste concatenated data for a sport
def paste_concatenated_data(pivot_df, workbook, sport, start_cell):
    template_sheet = workbook["Template"]
    filtered_df = pivot_df[pivot_df['Sport'] == sport]

    if not filtered_df.empty:
        concatenated_values = filtered_df.iloc[:, 2:].apply(lambda col: "\n".join(col.dropna()), axis=0).tolist()
        col_letter, row_num = start_cell[0], int(start_cell[1:])
        start_col_idx = ord(col_letter.upper()) - ord("A") + 1

        for col_idx, value in enumerate(concatenated_values, start=start_col_idx):
            cell = template_sheet.cell(row=row_num, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Function to generate the Excel report
def generate_excel(selected_date):
    template_path = "Excel_template.xlsx"  # Template path
    output_filename = f"Training_Report_{selected_date.strftime('%d%b%Y')}.xlsx"

    # Copy the template
    shutil.copy(template_path, output_filename)

    # Load the copied template
    workbook = load_workbook(output_filename)
    template_sheet = workbook["Template"]

    # Use selected date directly as the start date
    start_date = selected_date
    end_date = start_date + timedelta(days=6)

    # Show the selected date range in the Streamlit app
    st.write(f"**Selected Date Range:** {start_date.strftime('%a %d %b %Y')} to {end_date.strftime('%a %d %b %Y')}")

    # Fetch data
    session = requests.Session()
    session.auth = ("kenneth.mcmillan", "Quango76")
    response = session.get("https://aspire.smartabase.com/aspireacademy/live?report=PYTHON3_TRAINING_PLAN&updategroup=true")
    response.raise_for_status()
    data = pd.read_html(StringIO(response.text))[0]
    df = data.drop(columns=['About'], errors='ignore').drop_duplicates()

    df.columns = df.columns.str.replace(' ', '_')
    df['Start_Time'] = pd.to_numeric(df['Start_Time'], errors='coerce').apply(lambda x: convert_to_time(x))
    df['Finish_Time'] = pd.to_numeric(df['Finish_Time'], errors='coerce').apply(lambda x: convert_to_time(x))
    # Drop rows where 'Sport' is blank (NaN or empty string)
    df = df[df['Sport'].notna() & (df['Sport'].str.strip() != '')]

    df = df[df['Venue'] != 'AASMC']
    df = df[df['Sport'] != 'Generic Athlete']

    # Filter and clean data
    df['Date'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True).dt.date
    filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    filtered_df.loc[:, 'AM/PM'] = pd.Categorical(filtered_df['AM/PM'], categories=['AM', 'PM'], ordered=True)
    filtered_df = filtered_df.dropna(subset=['Sport']).sort_values(by=['Date', 'Sport', 'Coach', 'AM/PM'])

    # Group and pivot data
    grouped = (
        filtered_df.groupby(['Sport', 'Training_Group', 'Day_AM/PM'])
        .apply(format_session)
        .reset_index()
    )
    grouped.columns = ['Sport', 'Training_Group', 'Day_AM/PM', 'Session']

    pivot_df = pd.pivot_table(
        grouped,
        values='Session',
        index=['Sport', 'Training_Group'],
        columns=['Day_AM/PM'],
        aggfunc='first',
        fill_value=' '
    ).reset_index()

    # Ensure all day/time columns are present
    day_order = [
        f"{day} {time}" for day in ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
        for time in ['AM', 'PM']
    ]
    pivot_df = ensure_all_columns(pivot_df, day_order)

    # Insert data into the template
    rows_to_paste = [
        
        # first page
        
        {"sport": "Development", "training_group": "Development 1", "start_cell": "C6"},
        {"sport": "Development", "training_group": "Development 2", "start_cell": "C8"},
        {"sport": "Development", "training_group": "Development 3", "start_cell": "C10"},
        {"sport": "Endurance", "training_group": "Endurance_Senior", "start_cell": "C12"},
        {"sport": "Jumps", "training_group": "Jumps_PV", "start_cell": "C14"},
        {"sport": "Jumps", "training_group": "Jumps_Martin Bercel", "start_cell": "C16"},
        {"sport": "Jumps", "training_group": "Jumps_Ross Jeffs", "start_cell": "C18"},
        {"sport": "Jumps", "training_group": "Jumps_ElWalid", "start_cell": "C20"},
        {"sport": "Sprints", "training_group": "Sprints_Lee", "start_cell": "C22"},
        {"sport": "Sprints", "training_group": "Sprints_Hamdi", "start_cell": "C24"},
        {"sport": "Throws", "training_group": "Performance Throws", "start_cell": "C26"},
        
        # second page
        
        {"sport": "Squash", "training_group": "Squash", "start_cell": "C37"},
        {"sport": "Table Tennis", "training_group": "Table Tennis", "start_cell": "C39"},
        {"sport": "Fencing", "training_group": "Fencing", "start_cell": "C41"},
        {"sport": "Swimming", "training_group": "Swimming", "start_cell": "C43"},
        {"sport": "Padel", "training_group": "Padel", "start_cell": "C45"},
        
        {"sport": "Pre Academy Padel", "training_group": "Explorers", "start_cell": "C48"},
        {"sport": "Pre Academy Padel", "training_group": "Explorers+", "start_cell": "C49"},
        {"sport": "Pre Academy Padel", "training_group": "Starters", "start_cell": "C50"},
        {"sport": "Pre Academy", "training_group": "Pre Academy Fencing", "start_cell": "C51"},
        {"sport": "Pre Academy", "training_group": "Pre Academy Squash Girls", "start_cell": "C53"},
        {"sport": "Pre Academy", "training_group": "Pre Academy Athletics", "start_cell": "C55"},
        {"sport": "Girls Programe", "training_group": "Kids", "start_cell": "C58"},
        {"sport": "Girls Programe", "training_group": "Mini Cadet_U14", "start_cell": "C59"},
        {"sport": "Girls Programe", "training_group": "Cadet_U16", "start_cell": "C60"},
        {"sport": "Girls Programe", "training_group": "Youth_U18", "start_cell": "C61"},
        
        # third page
        
        {"sport": "Sprints", "training_group": "Sprints_Short", "start_cell": "C69"},
        {"sport": "Sprints", "training_group": "Sprints_Long", "start_cell": "C71"},
    ]

    for row in rows_to_paste:
        paste_filtered_data_to_template(
            pivot_df=pivot_df,
            workbook=workbook,
            sport=row["sport"],
            training_group=row["training_group"],
            start_cell=row["start_cell"],
        )

    #paste_concatenated_data(pivot_df, workbook, sport="Pre Academy Padel", start_cell="C47")
    #paste_concatenated_data(pivot_df, workbook, sport="Girls Programe", start_cell="C55")

    # Add dates to the template
    date_cells = ['C4', 'E4', 'G4', 'I4', 'K4', 'M4', 'O4',
                  'C35', 'E35', 'G35', 'I35', 'K35', 'M35', 'O35',
                  'C67', 'E67', 'G67', 'I67', 'K67', 'M67', 'O67']
    for idx, cell in enumerate(date_cells):
        day_offset = idx
        template_sheet[cell].value = (start_date + timedelta(days=day_offset)).strftime('%a %d %b %Y')
        template_sheet[cell].alignment = Alignment(horizontal="center", vertical="center")

    # Add "Week Beginning" text to specific cells
    week_number = start_date.isocalendar()[1]  # Get the ISO week number
    week_beginning_text = f"Week beginning {start_date.strftime('%d %b')}\nWeek {week_number}"


    # Populate cells O2, O33, and O65 with the week information
    template_sheet["O2"].value = week_beginning_text
    template_sheet["O33"].value = week_beginning_text
    template_sheet["O65"].value = week_beginning_text

    # Center align the text in these cells
    for cell in ["O2", "O33", "O65"]:
        template_sheet[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


    # Save to a binary stream
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return output, pivot_df

################################################################################################

# Streamlit App
st.title("Operations - Weekly Training Plan App")
st.markdown("Generate an Excel report for any week (past or future).")

# Date input
selected_date = st.date_input("Select a starting Sunday (make sure to choose a Sunday)", value=datetime.now().date())

# Button to generate the report
if st.button("Generate Report"):
    try:
        # Generate the Excel report and show the pivot DataFrame
        excel_file, pivot_df = generate_excel(selected_date)

        # Display the pivot DataFrame for debugging
        st.markdown("### DataFrame for checking data")
        st.dataframe(pivot_df)

        # Provide download button for the generated Excel report
        st.download_button(
            label="Download Excel Report",
            data=excel_file,
            file_name=f"Training_Report_{selected_date.strftime('%d%b%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    except Exception as e:
        st.error(f"An error occurred: {e}")
