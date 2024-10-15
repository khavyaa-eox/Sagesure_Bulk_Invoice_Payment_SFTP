import os
from datetime import datetime
import pandas as pd
import pysftp
import io
import paths
import credfile

# Disable host key checking
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

SFTP_HOST = credfile.host
SFTP_USERNAME = credfile.username
SFTP_PASSWORD = credfile.password
ERROR_LOG = paths.error_file_path
SFTP_ERROR_LOG = paths.sftp_error_log

# Function to log the error
def log_error(file_name, row_number, exc_type, exc_obj, exc_tb, eox_comments):
    # Get exception time
    exception_time = datetime.now().strftime('%m_%d_%Y %H:%M:%S')
    
    # Load the existing error log or create a new DataFrame if it doesn't exist
    if os.path.exists(ERROR_LOG):
        error_data = pd.read_excel(ERROR_LOG)
    else:
        # Create a DataFrame with the required headers if file does not exist
        error_data = pd.DataFrame(columns=['Filename', 'Row Number', 'Exception Type', 'Exception Object', 'Exception Traceback', 'Log Time', 'EOX Comments'])
    
    # Create a new entry
    new_entry = {
        'Filename': file_name,
        'Row Number': row_number,
        'Exception Type': exc_type,
        'Exception Object': exc_obj,
        'Exception Traceback': exc_tb,
        'Log Time': exception_time,
        'EOX Comments': eox_comments
    }

    # Append the new entry to the DataFrame
    error_data = error_data.append(new_entry, ignore_index=True)

    # Save the updated DataFrame back to the local Excel file
    error_data.to_excel(ERROR_LOG, index=False)

    # Update the same in the error_log in SFTP
    try:
        with pysftp.Connection(SFTP_HOST, username=SFTP_USERNAME, password=SFTP_PASSWORD, cnopts=cnopts) as sftp:
            # Step 1: Download the error log file from SFTP
            with sftp.open(SFTP_ERROR_LOG, mode='r') as remote_file:
                error_log_data = io.BytesIO(remote_file.read())  # Load file content into memory
                sftp_error_data = pd.read_excel(error_log_data)

            # Step 2: Append the new entry to the DataFrame from SFTP
            sftp_error_data = sftp_error_data.append(new_entry, ignore_index=True)

            # Step 3: Save the updated DataFrame into memory
            output_data = io.BytesIO()
            sftp_error_data.to_excel(output_data, index=False, engine='openpyxl')
            output_data.seek(0)  # Rewind the buffer

            # Step 4: Upload the updated file back to SFTP
            with sftp.open(SFTP_ERROR_LOG, mode='w') as remote_file:
                remote_file.write(output_data.read())  # Write updated content back to SFTP

            print(f"Error logged successfully in {SFTP_ERROR_LOG}")

    except Exception as e:
        print(f"Unable to log error in {SFTP_ERROR_LOG}. Check error in RDC error_log.xlsx\nError details: {e}")
