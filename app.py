import streamlit as st
from io import BytesIO, StringIO
import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.styles import Alignment
import shutil
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn  # Needed for setting cell shading

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

# ----------------------------------------
# Helper function to set cell background color in Word tables
def set_cell_background(cell, color):
    """
    Set the background shading color for a cell.
    :param cell: a docx.table._Cell object
    :param color: Hex color string (e.g., "ADD8E6" for light blue)
    """
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), color)
    tcPr.append(shd)

# ----------------------------------------
# Function to adjust timestamps and convert to local time
def convert_to_time(timestamp_ms, offset_hours=12):
    if isinstance(timestamp_ms, (pd.Series, list, tuple)):
        # Handle iterable case
        return [convert_to_time(x, offset_hours) for x in timestamp_ms]
    
    try:
        if pd.notnull(timestamp_ms):
            timestamp_s = float(timestamp_ms) / 1000
            return (datetime.fromtimestamp(timestamp_s, tz=timezone.utc) - timedelta(hours=offset_hours)).strftime('%H:%M')
        return None
    except (ValueError, TypeError) as e:
        print(f"Error converting timestamp {timestamp_ms}: {str(e)}")
        return None

# ----------------------------------------
# Helper function to safely parse a time value (ensuring it is a string in '%H:%M' format)
def parse_time(time_val):
    try:
        # Always convert time_val to string before parsing.
        return datetime.strptime(str(time_val), '%H:%M')
    except Exception:
        return datetime.min

# Helper function to be used as key in sorting.
def safe_parse(x):
    try:
        return parse_time(x[0])
    except Exception:
        return datetime.min

# Function to ensure all expected columns are present in the pivot DataFrame
def ensure_all_columns(pivot_df, day_order):
    return pivot_df.reindex(columns=['Sport', 'Training_Group'] + day_order, fill_value=' ')

# Function to format session information for grouping using itertuples()
def format_session(group):
    if not isinstance(group, pd.DataFrame):
        return ""
    
    venue_time_pairs = []
    for row in group.itertuples(index=False):
        type_value = str(row.Session_Type) if pd.notnull(row.Session_Type) else ''
        venue = str(row.Venue) if pd.notnull(row.Venue) else ''
        start_time = str(row.Start_Time) if pd.notnull(row.Start_Time) else ''
        finish_time = str(row.Finish_Time) if pd.notnull(row.Finish_Time) else ''
        time_str = f"{start_time}-{finish_time}" if start_time or finish_time else ''

        if type_value == "Training Camp":
            return "TRAINING CAMP"

        if type_value == "Competition":
            formatted_entry = f"Competition\n{venue}\n{time_str}".strip()
        else:
            formatted_entry = f"{venue}\n{time_str}".strip()

        if venue or time_str or type_value:
            venue_time_pairs.append((start_time, formatted_entry))

    sorted_venue_time_pairs = sorted(venue_time_pairs, key=safe_parse)
    sorted_sessions = [pair[1] for pair in sorted_venue_time_pairs]
    return '\n'.join(filter(None, sorted_sessions))

# Global dictionary with athlete count placeholders (each group has 10 athletes)
rows_to_paste = [
    {"sport": "Development", "training_group": "Development 1", "start_cell": "C6", "athlete_count": 14},
    {"sport": "Development", "training_group": "Development 2", "start_cell": "C8", "athlete_count": 11},
    {"sport": "Development", "training_group": "Development 3", "start_cell": "C10", "athlete_count": 9},
    
    {"sport": "Endurance", "training_group": "Endurance_Senior", "start_cell": "C12", "athlete_count": 18},
    {"sport": "Jumps", "training_group": "Jumps_Jaco", "start_cell": "C14", "athlete_count": 5},
    {"sport": "Jumps", "training_group": "Jumps Martin", "start_cell": "C16", "athlete_count": 5},
    {"sport": "Jumps", "training_group": "Jumps_Ross Jeffs", "start_cell": "C18", "athlete_count": 6},
    {"sport": "Jumps", "training_group": "Jumps_ElWalid", "start_cell": "C20", "athlete_count": 9},
    {"sport": "Sprints", "training_group": "Sprints_Lee", "start_cell": "C22", "athlete_count": 8},
    {"sport": "Sprints", "training_group": "Sprints_Hamdi", "start_cell": "C24", "athlete_count": 9},
    {"sport": "Throws", "training_group": "Senior Performance Throws", "start_cell": "C26", "athlete_count": 12},
    {"sport": "Squash", "training_group": "Squash", "start_cell": "C37", "athlete_count": 13},
    {"sport": "Table Tennis", "training_group": "Table Tennis", "start_cell": "C39", "athlete_count": 5},
    {"sport": "Fencing", "training_group": "Fencing", "start_cell": "C41", "athlete_count": 16},
    {"sport": "Swimming", "training_group": "Swimming", "start_cell": "C43", "athlete_count": 16},
    {"sport": "Padel", "training_group": "Padel", "start_cell": "C45", "athlete_count": 9},
    {"sport": "Pre Academy Padel", "training_group": "Explorers", "start_cell": "C48", "athlete_count": 10},
    {"sport": "Pre Academy Padel", "training_group": "Explorers+", "start_cell": "C49", "athlete_count": 10},
    {"sport": "Pre Academy Padel", "training_group": "Starters", "start_cell": "C50", "athlete_count": 10},
    {"sport": "Pre Academy", "training_group": "Pre Academy Fencing", "start_cell": "C51", "athlete_count": 10},
    {"sport": "Pre Academy", "training_group": "Pre Academy Squash Girls", "start_cell": "C53", "athlete_count": 10},
    {"sport": "Pre Academy", "training_group": "Pre Academy Athletics", "start_cell": "C55", "athlete_count": 10},
    {"sport": "Girls Programe", "training_group": "Kids", "start_cell": "C58", "athlete_count": 16},
    {"sport": "Girls Programe", "training_group": "Mini Cadet_U14", "start_cell": "C59", "athlete_count": 8},
    {"sport": "Girls Programe", "training_group": "Cadet_U16", "start_cell": "C60", "athlete_count": 6},
    {"sport": "Girls Programe", "training_group": "Youth_U18", "start_cell": "C61", "athlete_count": 3},

    {"sport": "Sprints", "training_group": "Sprints_Steve", "start_cell": "C69", "athlete_count": 11},
    {"sport": "Sprints", "training_group": "Sprints_Kurt", "start_cell": "C71", "athlete_count": 14},
    {"sport": "Sprints", "training_group": "Sprints_Rafal", "start_cell": "C73", "athlete_count": 10},
    {"sport": "Sprints", "training_group": "Sprints_Francis", "start_cell": "C75", "athlete_count": 3},
    {"sport": "Sprints", "training_group": "Sprints_Yasmani", "start_cell": "C77", "athlete_count": 8},

    {"sport": "Endurance", "training_group": "Endurance_Driss", "start_cell": "C81", "athlete_count": 10},
    {"sport": "Endurance", "training_group": "Endurance_Kada", "start_cell": "C83", "athlete_count": 5},
    {"sport": "Endurance", "training_group": "Endurance_Khamis", "start_cell": "C85", "athlete_count": 11},
    {"sport": "Decathlon", "training_group": "Decathlon_Willem", "start_cell": "C87", "athlete_count": 7},
    
    {"sport": "Jumps", "training_group": "Jumps_Linus", "start_cell": "C96", "athlete_count": 4},
    {"sport": "Jumps", "training_group": "Jumps_Pawel", "start_cell": "C98", "athlete_count": 4},

    {"sport": "Throws", "training_group": "Throws_Kemal", "start_cell": "C102", "athlete_count": 8},
    {"sport": "Throws", "training_group": "Throws_Krzysztof", "start_cell": "C104", "athlete_count": 4},
    {"sport": "Throws", "training_group": "Throws_Keida", "start_cell": "C106", "athlete_count": 3},
]

# Function to paste filtered data into the Excel template sheet    !
def paste_filtered_data_to_template(pivot_df, workbook, sport, training_group, start_cell):
    template_sheet = workbook["Template"]
    filtered_row = pivot_df[(pivot_df['Sport'] == sport) & (pivot_df['Training_Group'] == training_group)]
    if not filtered_row.empty:
        values_to_paste = filtered_row.iloc[0, 2:].tolist()
        col_letter, row_num = start_cell[0], int(start_cell[1:])
        start_col_idx = ord(col_letter.upper()) - ord("A") + 1
        for col_idx, value in enumerate(values_to_paste, start=start_col_idx):
            cell = template_sheet.cell(row=row_num, column=col_idx, value=value)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

# Function to paste concatenated data for a sport (if needed)
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
    template_path = "Excel_template.xlsx"
    output_filename = f"Training_Report_{selected_date.strftime('%d%b%Y')}.xlsx"
    shutil.copy(template_path, output_filename)
    workbook = load_workbook(output_filename)
    template_sheet = workbook["Template"]

    start_date = selected_date
    end_date = start_date + timedelta(days=6)
    st.write(f"**Selected Date Range:** {start_date.strftime('%a %d %b %Y')} to {end_date.strftime('%a %d %b %Y')}")
    
    session = requests.Session()
    session.auth = ("sb_sap.etl", "A1s2p3!re")
    response = session.get("https://aspire.smartabase.com/aspireacademy/live?report=PYTHON6_TRAINING_PLAN&updategroup=true")
    response.raise_for_status()
    data = pd.read_html(StringIO(response.text))[0]
    df = data.drop(columns=['About'], errors='ignore').drop_duplicates()
    df.columns = df.columns.astype(str).str.replace(' ', '_')
    
    df['Start_Time'] = pd.to_numeric(df['Start_Time'], errors='coerce').apply(lambda x: convert_to_time(x))
    df['Finish_Time'] = pd.to_numeric(df['Finish_Time'], errors='coerce').apply(lambda x: convert_to_time(x))
    df = df[df['Sport'].notna() & (df['Sport'].astype(str).str.strip() != '')]
    df = df[df['Venue'] != 'AASMC']
    df = df[df['Sport'] != 'Generic_Athlete']
    df = df[df['Training_Group'] != 'Practice']

    df['Date'] = pd.to_datetime(df['Date'], errors='coerce', dayfirst=True).dt.date
    filtered_df = df[(df['Date'] >= start_date) & (df['Date'] <= end_date)]
    
    filtered_df['AM/PM'] = filtered_df['AM/PM'].fillna('').astype(str)
    filtered_df['Session_Type'] = filtered_df['Session_Type'].fillna('').astype(str)
    
    filtered_df['Day_AM/PM'] = filtered_df.apply(
        lambda row: row['Date'].strftime('%A') + " " + str(row['AM/PM'])
        if pd.notnull(row['Date']) and pd.notnull(row['AM/PM']) else '',
        axis=1
    )
    
    filtered_df = filtered_df.dropna(subset=['Sport']).sort_values(by=['Date', 'Sport', 'Coach', 'AM/PM'])

    grouped = (
        filtered_df.groupby(['Sport', 'Training_Group', 'Day_AM/PM', 'Session_Type'])
        .apply(format_session)
        .reset_index()
    )
    grouped.columns = ['Sport', 'Training_Group', 'Day_AM/PM', 'Session_Type', 'Session']
    pivot_df = pd.pivot_table(
        grouped,
        values='Session',
        index=['Sport', 'Training_Group'],
        columns=['Day_AM/PM'],
        aggfunc=lambda x: "\n".join(x),  # Concatenate all session strings
        fill_value=' '
    ).reset_index()

    day_order = [f"{day} {time}" for day in 
                 ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']
                 for time in ['AM', 'PM']]
    pivot_df = ensure_all_columns(pivot_df, day_order)

    for row in rows_to_paste:
        paste_filtered_data_to_template(
            pivot_df=pivot_df,
            workbook=workbook,
            sport=row["sport"],
            training_group=row["training_group"],
            start_cell=row["start_cell"],
        )

    date_cells_groups = [
        ['C4', 'E4', 'G4', 'I4', 'K4', 'M4', 'O4'],
        ['C35', 'E35', 'G35', 'I35', 'K35', 'M35', 'O35'],
        ['C67', 'E67', 'G67', 'I67', 'K67', 'M67', 'O67'],
    ]
    for day_offset, cell_group in enumerate(zip(*date_cells_groups)):
        date_value = (start_date + timedelta(days=day_offset)).strftime('%a %d %b %Y')
        for cell in cell_group:
            template_sheet[cell].value = date_value
            template_sheet[cell].alignment = Alignment(horizontal="center", vertical="center")

    week_number = start_date.isocalendar()[1]
    week_beginning_text = f"Week beginning {start_date.strftime('%d %b')}\nWeek {week_number}"
    for cell in ["O2", "O33", "O65"]:
        template_sheet[cell].value = week_beginning_text
        template_sheet[cell].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    return output, pivot_df, filtered_df

# Function to generate a nicely formatted Word document for venue usage,
# including an "Athletes" column.
def generate_venue_usage_report(filtered_df, start_date):
    doc = Document()
    section = doc.sections[0]
    section.orientation = 1  # Landscape
    new_width, new_height = section.page_height, section.page_width
    section.page_width = new_width
    section.page_height = new_height

    title = doc.add_heading('Venue Usage Report', level=1)
    title.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    doc.add_paragraph(f'Week Beginning: {start_date.strftime("%d %b %Y")}', style='Normal')
    
    day_colors = {
        "Sunday": "D3D3D3",
        "Monday": "FFFFFF",
        "Tuesday": "D3D3D3",
        "Wednesday": "FFFFFF",
        "Thursday": "D3D3D3",
        "Friday": "FFFFFF",
        "Saturday": "D3D3D3"
    }
    
    athlete_count_map = {
        (row["sport"], row["training_group"]): row.get("athlete_count", 10)
        for row in rows_to_paste
    }
    
    venues = sorted([str(v) for v in filtered_df['Venue'].dropna().unique()])
    page_capacity = 5

    for i in range(0, len(venues), page_capacity):
        if i > 0:
            doc.add_page_break()
        venue_subset = venues[i:i+page_capacity]
        for venue in venue_subset:
            venue_data = filtered_df[filtered_df['Venue'].apply(lambda x: str(x)) == venue].sort_values(by=['Date', 'Start_Time'])
            venue_heading = doc.add_heading(f'📍 {venue}', level=2)
            venue_heading.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            table = doc.add_table(rows=1, cols=5)
            table.style = 'Table Grid'
            hdr_cells = table.rows[0].cells
            hdr_cells[0].text = 'Date'
            hdr_cells[1].text = 'Time'
            hdr_cells[2].text = 'Training Group'
            hdr_cells[3].text = 'Sport'
            hdr_cells[4].text = 'Athletes'
            for cell in hdr_cells:
                cell.paragraphs[0].runs[0].bold = True
                cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                set_cell_background(cell, "ADD8E6")
            
            for _, row in venue_data.iterrows():
                row_cells = table.add_row().cells
                date_str = row['Date'].strftime('%A %d %b %Y')
                time_str = f"{row['Start_Time']} - {row['Finish_Time']}"
                training_group = str(row['Training_Group'])
                sport = str(row['Sport'])
                athlete_count = athlete_count_map.get((sport, training_group), 10)
                
                row_cells[0].text = date_str
                row_cells[1].text = time_str
                row_cells[2].text = training_group
                row_cells[3].text = sport
                row_cells[4].text = str(athlete_count)
                
                day_name = row['Date'].strftime('%A')
                color = day_colors.get(day_name, "FFFFFF")
                for cell in row_cells:
                    set_cell_background(cell, color)
                    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# NEW FUNCTION:
# Generate a report showing, for each venue and date, the maximum number of people in any 30-minute interval,
# along with the groups present during that interval.
def generate_max_occupancy_report(filtered_df, start_date):
    from collections import defaultdict

    # Build athlete count lookup from global rows_to_paste.
    athlete_count_map = {
        (row["sport"], row["training_group"]): row.get("athlete_count", 10)
        for row in rows_to_paste
    }
    
    report_rows = []
    
    # Group by Date and Venue.
    for (date, venue), group in filtered_df.groupby(['Date', 'Venue']):
        sessions = []
        for idx, row in group.iterrows():
            try:
                start_dt = datetime.combine(date, datetime.strptime(row['Start_Time'], "%H:%M").time())
                finish_dt = datetime.combine(date, datetime.strptime(row['Finish_Time'], "%H:%M").time())
            except Exception:
                continue
            athlete_count = athlete_count_map.get((str(row['Sport']), str(row['Training_Group'])), 10)
            # Also record the group identifier.
            group_id = f"{str(row['Sport'])} - {str(row['Training_Group'])}"
            sessions.append({
                'start': start_dt,
                'finish': finish_dt,
                'count': athlete_count,
                'group': group_id
            })
        
        if not sessions:
            continue

        candidate_times = sorted({s['start'] for s in sessions} | {s['finish'] for s in sessions})
        max_occupancy = 0
        best_interval = None
        best_groups = set()
        
        for cand in candidate_times:
            window_start = cand
            window_end = cand + timedelta(minutes=30)
            overlapping = [s for s in sessions if s['start'] < window_end and s['finish'] > window_start]
            occupancy = sum(s['count'] for s in overlapping)
            if occupancy > max_occupancy:
                max_occupancy = occupancy
                best_interval = (window_start.strftime("%H:%M"), window_end.strftime("%H:%M"))
                best_groups = {s['group'] for s in overlapping}
        
        report_rows.append({
            'Date': date.strftime("%A %d %b %Y"),
            'Venue': venue,
            'Interval': f"{best_interval[0]} - {best_interval[1]}" if best_interval else "N/A",
            'Max Occupancy': max_occupancy,
            'Groups': ", ".join(sorted(best_groups)) if best_groups else "N/A"
        })
    
    # Generate Word document report.
    doc = Document()
    doc.add_heading("Maximum Occupancy Report", level=1)
    doc.add_paragraph(f"Week Beginning: {start_date.strftime('%d %b %Y')}")
    
    table = doc.add_table(rows=1, cols=5)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Date"
    hdr_cells[1].text = "Venue"
    hdr_cells[2].text = "30-min Interval"
    hdr_cells[3].text = "Max Occupancy"
    hdr_cells[4].text = "Groups"
    for cell in hdr_cells:
        cell.paragraphs[0].runs[0].bold = True
        cell.paragraphs[0].alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    
    for row in report_rows:
        row_cells = table.add_row().cells
        row_cells[0].text = row['Date']
        row_cells[1].text = row['Venue']
        row_cells[2].text = row['Interval']
        row_cells[3].text = str(row['Max Occupancy'])
        row_cells[4].text = row['Groups']
    
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output

# --- Streamlit App with Session State for Preserving Generated Data ---
if "generated" not in st.session_state:
    st.session_state.generated = False
if "excel_file" not in st.session_state:
    st.session_state.excel_file = None
if "venue_file" not in st.session_state:
    st.session_state.venue_file = None
if "max_occ_file" not in st.session_state:
    st.session_state.max_occ_file = None
if "pivot_df" not in st.session_state:
    st.session_state.pivot_df = None
if "filtered_data" not in st.session_state:
    st.session_state.filtered_data = None

st.title("Operations - Weekly Training Plan App")
st.markdown("Generate Training Calendar and Venue Usage reports for any week from 1st January 2025.")

selected_date = st.date_input("Select a starting date (make sure to choose a SUNDAY!)", value=datetime.now().date())

if st.button("Generate Reports"):
    try:
        excel_file, pivot_df, filtered_data = generate_excel(selected_date)
        venue_file = generate_venue_usage_report(filtered_data, selected_date)
        max_occ_file = generate_max_occupancy_report(filtered_data, selected_date)
        st.session_state.excel_file = excel_file
        st.session_state.pivot_df = pivot_df
        st.session_state.filtered_data = filtered_data
        st.session_state.venue_file = venue_file
        st.session_state.max_occ_file = max_occ_file
        st.session_state.generated = True
    except Exception as e:
        st.error(f"An error occurred: {e}")

if st.session_state.generated:
    st.markdown("### Pivot DataFrame for checking data")
    st.dataframe(st.session_state.pivot_df)
    st.download_button(
        label="📅 Download Training Calendar Excel Report",
        data=st.session_state.excel_file,
        file_name=f"Training_Report_{selected_date.strftime('%d%b%Y')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.download_button(
        label="📄 Download Venue Usage Report",
        data=st.session_state.venue_file,
        file_name=f"Venue_Usage_Report_{selected_date.strftime('%d%b%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )
    
    st.download_button(
        label="📈 Download Maximum Occupancy Report",
        data=st.session_state.max_occ_file,
        file_name=f"Max_Occupancy_Report_{selected_date.strftime('%d%b%Y')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )