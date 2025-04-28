from flask import Flask, request, jsonify, send_from_directory
import os
import pandas as pd
from datetime import datetime
import pyodbc
import logging

app = Flask(__name__)

# Database Configuration
SERVER = 'dsb.intelehealth.org,4747'
DATABASE = 'eSanjeevani'
USERNAME = 'sa'
PASSWORD = 'pJoshi@123'
TABLE_NAME = 'es_hwc_aggregate_jj'

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filename='upload_log.txt'
)

def get_db_connection():
    """Create database connection"""
    conn = pyodbc.connect(
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={SERVER};'
        f'DATABASE={DATABASE};'
        f'UID={USERNAME};'
        f'PWD={PASSWORD}'
    )
    return conn

def check_duplicates(df):
    """Check for duplicate records using pyodbc"""
    duplicates = []
    
    try:
        conn = get_db_connection()
        cursor = conn.cursor()
        
        for idx, row in df.iterrows():
            query = """
                SELECT COUNT(*) FROM es_hwc_aggregate_jj
                WHERE project = ? 
                AND from_district = ?
                AND from_health_facility = ?
                AND from_user = ?
                AND to_district = ?
                AND to_health_facility = ?
                AND to_user = ?
                AND month_year = ?
                AND speciality = ?
            """
            
            params = (
                str(row['project']),
                str(row['from_district']),
                str(row['from_health_facility']),
                str(row['from_user']),
                str(row['to_district']),
                str(row['to_health_facility']),
                str(row['to_user']),
                str(row['month_year']),
                str(row['speciality'])
            )
            
            cursor.execute(query, params)
            result = cursor.fetchval()
            
            if result > 0:
                duplicates.append(idx)
                
    finally:
        cursor.close()
        conn.close()
    
    return duplicates

def insert_data(df):
    """Insert data using pyodbc"""
    conn = get_db_connection()
    cursor = conn.cursor()
    
    try:
        for _, row in df.iterrows():
            query = """
                INSERT INTO es_hwc_aggregate_jj (
                    project, year, quarter, month, week, month_year,
                    from_district, from_health_facility, from_user,
                    to_district, to_health_facility, to_user,
                    speciality, hub_type, total_consultations,
                    month_year_sort_order
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            
            params = (
                str(row['project']),
                int(row['year']),
                str(row['quarter']),
                str(row['month']),
                str(row['week']),
                str(row['month_year']),
                str(row['from_district']),
                str(row['from_health_facility']),
                str(row['from_user']),
                str(row['to_district']),
                str(row['to_health_facility']),
                str(row['to_user']),
                str(row['speciality']),
                str(row['hub_type']),
                int(row['total_consultations']),
                int(row['month_year_sort_order'])
            )
            
            cursor.execute(query, params)
        
        conn.commit()
        
    finally:
        cursor.close()
        conn.close()

def process_excel_files(files):
    """Process uploaded Excel files"""
    try:
        # Create output directory if it doesn't exist
        output_dir = os.path.join(os.path.dirname(__file__), "output_files")
        os.makedirs(output_dir, exist_ok=True)
        
        # Process all uploaded files
        df_all = pd.DataFrame()
        processed_files = []
        
        for file in files:
            try:
                df = pd.read_excel(file, sheet_name="RawData")
                df_all = pd.concat([df_all, df], ignore_index=True)
                processed_files.append(file.filename)
            except Exception as e:
                logging.error(f"Error processing {file.filename}: {str(e)}")
                return False, f"Error processing {file.filename}: Make sure it has a 'RawData' sheet"

        if df_all.empty:
            return False, "No data found in the uploaded files"

        # Process dates
        df_all["Start Date"] = pd.to_datetime(df_all["Start Date"], errors="coerce")
        df_all["End Date"] = pd.to_datetime(df_all["End Date"], errors="coerce")

        # Create aggregated columns
        df_all["project"] = df_all["To State"]
        df_all["year"] = df_all["Start Date"].dt.year
        df_all["quarter"] = df_all["Start Date"].dt.quarter.astype(str)
        df_all["month"] = df_all["Start Date"].dt.strftime("%B")
        df_all["week"] = df_all["Start Date"].dt.strftime("%U")
        df_all["month_year"] = df_all["End Date"].dt.strftime("%B-%Y")
        df_all["month_year_sort_order"] = df_all["Start Date"].dt.year * 100 + df_all["Start Date"].dt.month

        # Rename columns
        df_all = df_all.rename(columns={
            "From District": "from_district",
            "From Health Facility": "from_health_facility",
            "From User": "from_user",
            "To District": "to_district",
            "To Health Facility": "to_health_facility",
            "To User": "to_user",
            "Speciality": "speciality",
            "Categorization of Hub": "hub_type"
        })

        # Aggregate data
        df_agg = df_all.groupby([
            "project", "year", "quarter", "month", "week", "month_year", 
            "from_district", "from_health_facility", "from_user", 
            "to_district", "to_health_facility", "to_user", 
            "speciality", "hub_type"
        ]).agg(
            total_consultations=("Consultation Id", "count"),
            month_year_sort_order=("month_year_sort_order", "max")
        ).reset_index()

        # Save backup CSV
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = os.path.join(output_dir, f"aggregated_data_{timestamp}.csv")
        df_agg.to_csv(output_file, index=False)

        # Check for duplicates
        duplicate_indices = check_duplicates(df_agg)
        
        if duplicate_indices:
            # Remove duplicate rows
            df_agg = df_agg.drop(duplicate_indices)
            
            if df_agg.empty:
                return True, "All records already exist in the database. No new data to add."

        # Insert non-duplicate records
        if not df_agg.empty:
            insert_data(df_agg)
            
        return True, f"""Successfully processed {len(processed_files)} files.
Found {len(duplicate_indices)} duplicate records that were skipped.
Added {len(df_agg)} new records to the database."""

    except Exception as e:
        logging.error(f"Error: {str(e)}")
        return False, f"An error occurred: {str(e)}"

@app.route('/')
def index():
    return send_from_directory('.', 'index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'files' not in request.files:
        return jsonify({'error': 'No files uploaded'}), 400
    
    files = request.files.getlist('files')
    if not files:
        return jsonify({'error': 'No files selected'}), 400

    success, message = process_excel_files(files)
    if success:
        return jsonify({'message': message}), 200
    else:
        return jsonify({'error': message}), 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)