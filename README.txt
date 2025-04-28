HOW TO USE THE EXCEL UPLOAD SYSTEM
================================

First Time Setup:
---------------
1. Make sure Python is installed on your computer (Python 3.7 or higher)
2. Open Command Prompt (cmd) as administrator
3. Navigate to the tool folder:
   cd path\to\Excel_Upload_Tool
4. Install required packages:
   pip install -r requirements.txt

Using the Upload Tool:
--------------------
1. Double-click "start_server.bat"
2. Open your web browser and go to: http://localhost:5000
3. Click "Select Excel Files" to choose your files
4. Click "Upload and Process Files"
5. Wait for the success message

Requirements for Excel Files:
---------------------------
- Each Excel file must have a sheet named "RawData"
- Required columns in the RawData sheet:
  * Start Date
  * End Date
  * To State
  * Consultation Id
  * From District
  * From Health Facility
  * From User
  * To District
  * To Health Facility
  * To User
  * Speciality
  * Categorization of Hub

What Happens After Upload:
------------------------
1. Your files will be processed automatically
2. The data will be uploaded to the database
3. A backup CSV file will be saved in the output_files folder
4. You'll see a success message showing:
   - Number of files processed
   - Number of duplicate records skipped
   - Number of new records added

Folder Structure:
---------------
- Make sure these folders exist in the same directory:
  * input_files/  (optional - for storing input Excel files)
  * output_files/ (will be created automatically for backups)

Common Issues and Solutions:
-------------------------
1. "No module named 'flask'" error:
   - Open Command Prompt as administrator
   - Run: pip install -r requirements.txt

2. Excel file errors:
   - Make sure your Excel files have a sheet named "RawData"
   - Close all Excel files before uploading
   - Check that all required columns are present

3. Connection errors:
   - Make sure you're connected to the office network
   - Verify VPN is connected (if working remotely)

If You See Any Errors:
--------------------
1. Take a screenshot of the error message
2. Note what you were trying to do
3. Contact IT support with:
   - The screenshot
   - Name of the Excel file(s) you were trying to upload
   - Time when the error occurred

Need Help?
---------
Contact IT Support if you have any questions or issues.