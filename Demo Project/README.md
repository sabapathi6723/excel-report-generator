# Excel Report Generator

A Flask web application that allows users to upload Excel files and generate either Participation Reports or Performance Reports with pivot tables and bar charts.

## Features

- 📊 **Participation Report**: Generates a report with pivot tables showing participation by department and test status
- 📈 **Performance Report**: Categorizes performance scores and generates detailed performance analysis
- 📅 **Parul Weekly Report**: Automates the Parul Overall Data sheet with portal/attempt/category logic, divisional and overall performance summaries, participation summary, and professional formatting
- 🎨 **Modern UI**: Beautiful glassmorphism design with blue gradient background
- 📁 **Excel Processing**: Uses pandas and openpyxl for advanced Excel manipulation
- 📉 **Charts**: Automatically generates bar charts in the output Excel files

## Project Structure

```
project/
├── app.py                 # Flask application
├── templates/
│   └── index.html        # Main UI template
├── reports/
│   ├── participation.py  # Participation report generator
│   ├── performance.py    # Performance report generator
│   └── parul_weekly.py   # Parul Weekly Overall Data generator
├── uploads/              # Temporary storage for uploaded files
├── requirements.txt      # Python dependencies
└── README.md            # This file
```

## Installation

1. **Clone or navigate to the project directory:**
   ```bash
   cd "Demo Project"
   ```

2. **Create a virtual environment (recommended):**
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. **Run the Flask application:**
   ```bash
   python app.py
   ```

2. **Open your web browser and navigate to:**
   ```
   http://localhost:5000
   ```

3. **Upload a CSV or Excel file:**
   - Click on the upload box or drag and drop your `.csv`, `.xlsx`, or `.xls` file
   - Select "Generate Participation Report", "Generate Performance Report", or "Generate Parul Weekly Report"
   - Wait for processing (the file will automatically download when ready)

## Excel File Requirements

### For Participation Report:
- Must contain a column named "Department" (case-insensitive)
- Must contain a column named "Test Status" or "Status" (case-insensitive)

### For Performance Report:
- Must meet all Participation Report requirements
- Must contain a column with "total percentage" or "percentage" in the name (case-insensitive)
- This percentage column should appear after the "Test Status" column

## Report Details

### Participation Report
- Creates a "Data" sheet with styled raw data
- Generates a "Participation Summary" sheet with pivot table
- Includes a bar chart showing participation by department and status
- Applies professional styling with dark blue headers and alternating row colors

### Performance Report
- Includes all features from Participation Report
- Adds a "Category" column to the data based on performance percentages:
  - **Good**: > 75%
  - **Satisfactory**: > 50%
  - **Need Attention**: > 25%
  - **Intervention**: ≤ 25%
  - **Not Attended**: NA/empty values
- Creates a "Performance Summary" sheet with pivot table by category
- Includes a bar chart showing performance distribution

### Parul Weekly Report
- Processes the Overall Data sheet for the Parul weekly assessment
- Adds **Portal Status**, **Attempt Status**, and **Category** columns with business rules
- Reorders and formats the key score columns (max, student, percentage)
- Applies professional formatting: blue headers, center alignment, borders, and auto column widths
- Outputs a polished workbook named `Parul_Weekly_Report_Processed.xlsx`

## Technical Details

- **Framework**: Flask 3.1.2
- **Data Processing**: pandas 2.3.3
- **Excel Manipulation**: openpyxl 3.1.5
- **Styling**: Custom CSS with glassmorphism effects
- **File Upload**: Secure file handling with validation

## Error Handling

The application includes comprehensive error handling for:
- Missing or invalid file uploads
- Missing required columns in Excel files
- File processing errors
- Invalid file formats

All errors are displayed as user-friendly flash messages in the UI.

## Development

To run in development mode:
```bash
python app.py
```

The app runs with debug mode enabled by default on `http://0.0.0.0:5000`.

## Production Deployment

For production deployment, consider:
- Using a production WSGI server like gunicorn (included in requirements.txt)
- Setting a secure `SECRET_KEY` in `app.py`
- Configuring proper file upload limits
- Setting up proper logging
- Using environment variables for configuration

Example with gunicorn:
```bash
gunicorn -w 4 -b 0.0.0.0:5000 app:app
```

## License

This project is designed and developed by Sabapathi.

## Support

For issues or questions, please check the error messages displayed in the application UI or review the console output for detailed error information.

