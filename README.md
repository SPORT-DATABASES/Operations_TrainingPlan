# Operations Training Plan Generator

A Streamlit web application that generates comprehensive training schedules and venue usage reports for sports academy operations. The application connects to Smartabase data and produces Excel training calendars and Word venue reports.

## 🎯 Features

- **Training Calendar Generation**: Creates detailed Excel reports showing weekly training schedules
- **Venue Usage Reports**: Generates Word documents with venue occupancy details
- **Maximum Occupancy Analysis**: Calculates peak venue usage with 30-minute interval analysis
- **Multi-Sport Support**: Handles Athletics, Swimming, Squash, Fencing, Padel, and more
- **Time Zone Conversion**: Converts UTC timestamps to local time (Qatar timezone)
- **Automated Data Processing**: Fetches live data from Smartabase API

## 📋 Requirements

```
requests
pandas
openpyxl
streamlit
lxml
python-docx
streamlit-aggrid
```

## 🚀 Installation

1. Clone the repository:
```bash
git clone https://github.com/SPORT-DATABASES/Operations_TrainingPlan.git
cd Operations_TrainingPlan
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Ensure `Excel_template.xlsx` is in the project directory

## 🖥️ Usage

### Running the Application

**Primary Application:**
```bash
streamlit run app2.py
```

**Alternative (Legacy):**
```bash
streamlit run app.py
```

### Using the Interface

1. **Select Date**: Choose a Sunday as the starting date for your weekly report
2. **Generate Reports**: Click the button to create all three reports
3. **Download**: Use the download buttons for:
   - 📅 Training Calendar Excel Report
   - 📄 Venue Usage Report (Word)
   - 📈 Maximum Occupancy Report (Word)

## 📊 Generated Reports

### 1. Training Calendar Excel Report
- **Format**: Excel (.xlsx)
- **Content**: Weekly training schedule organized by sport and training group
- **Features**:
  - Separate sections for different programs
  - AM/PM session breakdown
  - Athlete count integration
  - Date headers and week numbering

### 2. Venue Usage Report
- **Format**: Word Document (.docx)
- **Content**: Detailed venue utilization by date and time
- **Features**:
  - Organized by venue
  - Training group and sport details
  - Athlete count per session
  - Color-coded by day of week

### 3. Maximum Occupancy Report
- **Format**: Word Document (.docx)
- **Content**: Peak occupancy analysis for each venue
- **Features**:
  - 30-minute interval analysis
  - Maximum occupancy calculations
  - Groups present during peak times

## 🏗️ Project Structure

```
├── app.py                      # Legacy main application
├── app2.py                     # Primary application (recommended)
├── weekly_training_plan_email.py  # Email functionality
├── Excel_template.xlsx         # Excel template file
├── requirements.txt            # Python dependencies
├── extras/                     # Additional utilities
│   ├── app_backup.py
│   ├── app_venue_doc.py
│   ├── debug.py
│   └── ...
└── old_app/                    # Archived versions
```

## ⚙️ Configuration

### Time Zone Settings
The application converts UTC timestamps to local time using an 11-hour offset (Qatar Standard Time). This is configured in the `convert_to_time()` function.

### Athlete Count Mapping
Training group athlete counts are defined in the `rows_to_paste` array in both `app.py` and `app2.py`. Modify these values to update athlete numbers.

### Sports Categories
The application supports multiple sports categories:
- **Athletics**: Development, Endurance, Sprints, Jumps, Throws, Decathlon
- **Racket Sports**: Squash, Padel, Table Tennis
- **Combat Sports**: Fencing
- **Aquatic Sports**: Swimming
- **Youth Programs**: Pre Academy, Girls Programme

## 🔧 Data Source

The application connects to Smartabase using authenticated API calls:
- **Endpoint**: `https://aspire.smartabase.com/aspireacademy/live`
- **Report**: `PYTHON6_TRAINING_PLAN`
- **Authentication**: Basic authentication with credentials

## 📝 Technical Details

### Data Processing Pipeline
1. **Data Retrieval**: Fetch training data from Smartabase API
2. **Data Cleaning**: Remove invalid entries and duplicates
3. **Time Conversion**: Convert UTC timestamps to local time
4. **Data Grouping**: Group sessions by sport, training group, and time
5. **Pivot Processing**: Create pivot tables for report generation
6. **Template Population**: Fill Excel templates with processed data
7. **Document Generation**: Create Word documents with venue information

### Key Functions
- `convert_to_time()`: Handles timestamp conversion with timezone offset
- `format_session()`: Formats training session information for display
- `generate_excel()`: Creates the main Excel training calendar
- `generate_venue_usage_report()`: Produces venue utilization Word document
- `generate_max_occupancy_report()`: Analyzes peak venue occupancy

## 🐛 Troubleshooting

### Common Issues

1. **Times showing 1 hour early**: Check the `offset_hours` parameter in `convert_to_time()` function
2. **Template not found**: Ensure `Excel_template.xlsx` exists in the project directory
3. **API connection issues**: Verify Smartabase credentials and network connectivity
4. **Date selection errors**: Make sure to select a Sunday as the starting date

### Debug Mode
Use the debug utilities in the `extras/` folder for troubleshooting data processing issues.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/new-feature`)
3. Commit your changes (`git commit -m 'Add new feature'`)
4. Push to the branch (`git push origin feature/new-feature`)
5. Open a Pull Request

## 📄 License

This project is developed for internal operations use at sports academies.

## 📞 Support

For technical support or feature requests, please open an issue in the GitHub repository.

---

**Note**: This application is specifically designed for sports academy operations and requires access to Smartabase data systems.