"""
Flask Application for Excel Report Generation

This application allows users to upload Excel files and generate
either Participation Reports or Performance Reports.
"""

import os
import pandas as pd
from flask import Flask, render_template, request, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from reports.participation import generate_participation_report
from reports.performance import generate_performance_report

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
app.config['ALLOWED_EXTENSIONS'] = {'csv', 'xlsx', 'xls'}

# Create uploads directory if it doesn't exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)


def detect_file_type(file_path):
    """
    Detect the actual file type by reading magic bytes and content.
    
    Args:
        file_path (str): Path to the file
    
    Returns:
        str: 'excel', 'csv', or 'unknown'
    """
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
            # Excel files start with PK (ZIP signature)
            if header[:2] == b'PK':
                # Verify it's actually a valid Excel file by checking for Excel structure
                try:
                    f.seek(0)
                    # Read more to check for Excel structure
                    full_header = f.read(50)
                    # Excel files have specific ZIP structure
                    if b'xl/' in full_header or b'[Content_Types].xml' in full_header:
                        return 'excel'
                except:
                    pass
                # If PK but not Excel structure, might be other ZIP file
                return 'unknown'
            
            # CSV files are typically text - check for BOM or text content
            if header[:3] == b'\xef\xbb\xbf':  # UTF-8 BOM
                return 'csv'
            elif header[:2] == b'\xff\xfe' or header[:2] == b'\xfe\xff':  # UTF-16 BOM
                return 'csv'
            
            # Try to read as text to see if it's CSV
            try:
                # Try multiple encodings
                for encoding in ['utf-8', 'latin-1', 'cp1252']:
                    try:
                        with open(file_path, 'r', encoding=encoding, errors='ignore') as f:
                            first_line = f.readline()
                            # Check if it looks like CSV (has commas, semicolons, or tabs)
                            if first_line and (',' in first_line or ';' in first_line or '\t' in first_line):
                                return 'csv'
                            # If it's readable text but no delimiters, might still be CSV
                            if first_line and len(first_line.strip()) > 0:
                                # Check if it's not binary
                                try:
                                    first_line.encode('ascii')
                                    return 'csv'  # Likely CSV if it's readable text
                                except:
                                    pass
                    except:
                        continue
            except:
                pass
            return 'unknown'
    except Exception:
        return 'unknown'


def read_csv_with_encoding(file_path, nrows=None):
    """
    Read CSV file with automatic encoding detection and robust parsing.
    Tries multiple encodings and delimiters to handle various CSV formats.
    
    Args:
        file_path (str): Path to the CSV file
        nrows (int, optional): Number of rows to read (for testing)
    
    Returns:
        pandas.DataFrame: DataFrame with CSV data
    
    Raises:
        ValueError: If file cannot be read with any method
    """
    encodings = ['utf-8', 'latin-1', 'iso-8859-1', 'cp1252', 'windows-1252']
    delimiters = [',', ';', '\t', '|']
    
    # Try different combinations of encoding and delimiter
    for encoding in encodings:
        for delimiter in delimiters:
            try:
                if nrows:
                    df = pd.read_csv(
                        file_path,
                        encoding=encoding,
                        delimiter=delimiter,
                        nrows=nrows,
                        quotechar='"',
                        on_bad_lines='skip',  # Skip malformed lines
                        engine='python',  # Use Python engine for better error handling
                        skipinitialspace=True
                    )
                else:
                    df = pd.read_csv(
                        file_path,
                        encoding=encoding,
                        delimiter=delimiter,
                        quotechar='"',
                        on_bad_lines='skip',  # Skip malformed lines
                        engine='python',  # Use Python engine for better error handling
                        skipinitialspace=True
                    )
                # If we got here and have data, return it
                if len(df) > 0 or nrows:
                    return df
            except (UnicodeDecodeError, UnicodeError):
                continue
            except Exception as e:
                # If it's not an encoding error, try next delimiter
                if 'codec' in str(e).lower() or 'decode' in str(e).lower():
                    continue
                # For other errors, try next delimiter
                continue
    
    # If all combinations failed, try with most lenient settings (auto-detect separator)
    for encoding in encodings:
        try:
            if nrows:
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    nrows=nrows,
                    on_bad_lines='skip',
                    engine='python',
                    sep=None,  # Auto-detect separator
                    skipinitialspace=True
                )
            else:
                df = pd.read_csv(
                    file_path,
                    encoding=encoding,
                    on_bad_lines='skip',
                    engine='python',
                    sep=None,  # Auto-detect separator
                    skipinitialspace=True
                )
            if len(df) > 0 or nrows:
                return df
        except Exception:
            continue
    
    # Last resort: try with errors='replace'
    try:
        if nrows:
            return pd.read_csv(
                file_path,
                encoding='utf-8',
                errors='replace',
                nrows=nrows,
                on_bad_lines='skip',
                engine='python',
                sep=None,
                skipinitialspace=True
            )
        else:
            return pd.read_csv(
                file_path,
                encoding='utf-8',
                errors='replace',
                on_bad_lines='skip',
                engine='python',
                sep=None,
                skipinitialspace=True
            )
    except Exception as e:
        raise ValueError(f"Unable to read CSV file. Error: {str(e)}")


def allowed_file(filename):
    """
    Check if the uploaded file has an allowed extension.
    
    Args:
        filename (str): Name of the file
    
    Returns:
        bool: True if file extension is allowed
    """
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']


@app.route('/')
def index():
    """
    Render the main page with file upload form.
    
    Returns:
        str: Rendered HTML template
    """
    return render_template('index.html')


@app.route('/', methods=['POST'])
def upload_file():
    """
    Handle file upload and report generation.
    
    Returns:
        Response: File download or error message
    """
    try:
        # Check if file was uploaded
        if 'file' not in request.files:
            flash('No file uploaded. Please select a CSV or Excel file.', 'error')
            return redirect(url_for('index'))
        
        file = request.files['file']
        action = request.form.get('action')
        
        # Check if file is selected
        if file.filename == '':
            flash('No file selected. Please choose a CSV or Excel file.', 'error')
            return redirect(url_for('index'))
        
        # Check if action is selected
        if not action:
            flash('Please select a report type (Participation or Performance).', 'error')
            return redirect(url_for('index'))
        
        # Check if file extension is allowed
        if not allowed_file(file.filename):
            flash('Unsupported file format. Please upload CSV or Excel files.', 'error')
            return redirect(url_for('index'))
        
        # Get file extension
        file_ext = file.filename.lower().split('.')[-1]
        
        # Validate file extension
        if file_ext not in ['csv', 'xlsx', 'xls']:
            flash('Unsupported file format. Please upload CSV or Excel files.', 'error')
            return redirect(url_for('index'))
        
        # Save uploaded file
        filename = secure_filename(file.filename)
        upload_folder = os.path.abspath(app.config['UPLOAD_FOLDER'])
        input_path = os.path.join(upload_folder, filename)
        file.save(input_path)
        
        # Detect actual file type and verify file can be read
        actual_file_type = detect_file_type(input_path)
        
        # Determine how to read the file
        read_as_csv = False
        
        try:
            if file_ext == "csv":
                # Definitely CSV - read as CSV
                read_csv_with_encoding(input_path, nrows=1)
                read_as_csv = True
            elif file_ext in ["xlsx", "xls"]:
                # Check if it's actually CSV
                if actual_file_type == "csv":
                    # File is actually CSV, read it as CSV
                    read_csv_with_encoding(input_path, nrows=1)
                    read_as_csv = True
                elif actual_file_type == "excel":
                    # It's Excel, try reading as Excel
                    try:
                        pd.read_excel(input_path, nrows=1)
                        read_as_csv = False
                    except Exception as excel_error:
                        error_msg = str(excel_error).lower()
                        if 'cannot be used in worksheets' in error_msg or 'badzipfile' in error_msg or 'corrupt' in error_msg:
                            # Try CSV as fallback
                            try:
                                read_csv_with_encoding(input_path, nrows=1)
                                read_as_csv = True
                            except:
                                flash(f'Invalid or corrupted file. The file cannot be read as Excel or CSV.', 'error')
                                return redirect(url_for('index'))
                        else:
                            flash(f'Error reading Excel file: {str(excel_error)}. Please ensure the file is valid.', 'error')
                            return redirect(url_for('index'))
                else:
                    # Unknown type - try Excel first, then CSV
                    try:
                        pd.read_excel(input_path, nrows=1)
                        read_as_csv = False
                    except Exception as excel_error:
                        error_msg = str(excel_error).lower()
                        if 'cannot be used in worksheets' in error_msg or 'badzipfile' in error_msg or 'corrupt' in error_msg:
                            # Try CSV as fallback
                            try:
                                read_csv_with_encoding(input_path, nrows=1)
                                read_as_csv = True
                            except Exception as csv_error:
                                flash(f'Invalid file. Cannot read as Excel ({str(excel_error)}) or CSV ({str(csv_error)}).', 'error')
                                return redirect(url_for('index'))
                        else:
                            flash(f'Error reading Excel file: {str(excel_error)}. Please ensure the file is valid.', 'error')
                            return redirect(url_for('index'))
        except Exception as e:
            flash(f'Error reading file: {str(e)}. Please ensure the file is valid.', 'error')
            return redirect(url_for('index'))
        
        # Store the actual file type for report generation
        # We'll pass this information to the report generators
        
        # Generate output filename
        base_name = os.path.splitext(filename)[0]
        
        # If we determined it should be read as CSV but has Excel extension,
        # create a temporary CSV file so report generators treat it correctly
        temp_input_path = input_path
        temp_csv_path = None
        if read_as_csv and file_ext in ["xlsx", "xls"]:
            import shutil
            temp_csv_path = os.path.join(upload_folder, f"{base_name}_temp.csv")
            shutil.copy2(input_path, temp_csv_path)
            temp_input_path = temp_csv_path
            print(f"File detected as CSV but has {file_ext} extension. Using temporary CSV file: {temp_csv_path}")
        
        try:
            if action == 'participation':
                output_filename = f"{base_name}_participation_report.xlsx"
                output_path = os.path.join(upload_folder, output_filename)
                
                # Generate participation report
                print(f"Generating participation report: {temp_input_path} -> {output_path} (read_as_csv={read_as_csv})")
                generate_participation_report(temp_input_path, output_path)
                
            elif action == 'performance':
                output_filename = f"{base_name}_performance_report.xlsx"
                output_path = os.path.join(upload_folder, output_filename)
                
                # Generate performance report
                print(f"Generating performance report: {temp_input_path} -> {output_path} (read_as_csv={read_as_csv})")
                generate_performance_report(temp_input_path, output_path)
        finally:
            # Clean up temporary CSV file if created
            if temp_csv_path and os.path.exists(temp_csv_path):
                try:
                    os.remove(temp_csv_path)
                except:
                    pass
        
        if action not in ['participation', 'performance']:
            flash('Invalid action selected.', 'error')
            return redirect(url_for('index'))
        
        # Verify file was created
        if not os.path.exists(output_path):
            raise FileNotFoundError(f"Report file was not created: {output_path}")
        
        print(f"Sending file for download: {output_path}")
        
        # Send file for download using absolute path
        from flask import Response
        response = send_file(
            os.path.abspath(output_path),
            as_attachment=True,
            download_name=output_filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
        
        # Add headers to ensure download works
        response.headers['Content-Disposition'] = f'attachment; filename="{output_filename}"'
        response.headers['X-Content-Type-Options'] = 'nosniff'
        
        return response
    
    except ValueError as e:
        import traceback
        print(f"ValueError: {str(e)}")
        print(traceback.format_exc())
        flash(f'Error processing file: {str(e)}', 'error')
        return redirect(url_for('index'))
    
    except Exception as e:
        import traceback
        print(f"Exception: {str(e)}")
        print(traceback.format_exc())
        flash(f'An error occurred: {str(e)}. Please check the console for details.', 'error')
        return redirect(url_for('index'))


if __name__ == '__main__':
    app.run(debug=True, host='127.0.0.1', port=5001)

