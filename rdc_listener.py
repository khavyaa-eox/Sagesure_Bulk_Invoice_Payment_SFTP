# Necessary Libraries
import os
import time
import shutil
from datetime import datetime
import signal
import pysftp
from multiprocessing import Lock, Manager
import warnings
from requests.exceptions import RequestsDependencyWarning
warnings.simplefilter('ignore', RequestsDependencyWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)
warnings.simplefilter(action="ignore", category=UserWarning)

# Necessary Files
import credfile
import paths
from Bulk_Invoice_Payment_SFTP import call_process

# Disable host key checking
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

# Users credentials for Snapsheet
USERS = credfile.cred_details

# SFTP and RDC Folder details
SFTP_HOST = credfile.host
SFTP_USERNAME = credfile.username
SFTP_PASSWORD = credfile.password

SFTP_ASSIGNMET_FOLDER = paths.sftp_assignment_folder
SFTP_PROCESSING_FOLDER = paths.sftp_processing_folder
SFTP_ERROR_FOLDER = paths.sftp_error_folder
SFTP_COMPLETED_FOLDER = paths.sftp_completed_folder

LOCAL_DOWNLOADS = paths.local_download_path
LOCAL_OUTPUT = paths.local_completed_path
LOCAL_ERROR = paths.local_error_path


# RDC Monitoring Function
def monitor_assignment(lock, credentials):
    print("Connecting to SFTP...")
    try:
        with pysftp.Connection(SFTP_HOST, username=SFTP_USERNAME, password=SFTP_PASSWORD, cnopts=cnopts) as sftp:
            print("Connected!")
            rdc_assignment_folder = os.path.join(SFTP_ASSIGNMET_FOLDER, 'rdc1')
            files = sftp.listdir(rdc_assignment_folder)

            for file in files:
                # Move file from assignment folder to the respective RDC processing folder
                src_path = os.path.join(rdc_assignment_folder, file)  # Correct file path in the source folder               
                dest_file_path = os.path.join(SFTP_PROCESSING_FOLDER, 'rdc1', file)  # Append the file name

                print(f"Processing file '{file}' in rdc1")
                sftp.rename(src_path, dest_file_path)
                
                # Download the file to the local path
                local_file_path = os.path.join(LOCAL_DOWNLOADS, file)
                print(f"Local file path: {local_file_path}")
                os.makedirs(os.path.dirname(local_file_path), exist_ok=True)  # Ensure the directory exists
                print(f"Created directory in local path!")
                
                sftp.get(dest_file_path, local_file_path)
                print(f"\nDownloaded: {file}")

                set_credentials(lock, credentials, file)

    except Exception as e:
        print(f"\nError with RDC connection or file handling: {e}")

# Set credentials and start processing
def set_credentials(lock, credentials, filename):
    try:
        user = None
        # Acquire the lock to pop a user from the credentials pool
        with lock:
            if credentials:
                user = credentials.pop(0)  # Fetch the user
                email, password = user
                print(f"\nUser acquired for processing: {email}")
            else:
                print(f"\nNo available credentials to process {filename}")

        if user:

            # Hard coded for trial
            base_name, ext = os.path.splitext(filename)
            processed_file = str(f"{base_name}_completed{ext}")

            # Actual code
            # processed_file = call_process(filename, {'username': email, 'password': password})

            ###---------------------------------------------------------------###
            ###                      Completed Processing                     ###
            ###---------------------------------------------------------------###
            # After processing, return the user to the pool
            with lock:
                credentials.append(user)

            # Move the file to the respective rdc folder in sftp's completed folder
            try:
                with pysftp.Connection(SFTP_HOST, username=SFTP_USERNAME, password=SFTP_PASSWORD, cnopts=cnopts) as sftp:
                    base_name, ext = os.path.splitext(processed_file)
                    if "_completed" in base_name:
                        # Move file from the processing to completed folder in SFTP server
                        source = os.path.join(SFTP_PROCESSING_FOLDER, 'rdc1', filename)
                        destination = os.path.join(SFTP_COMPLETED_FOLDER, 'rdc1', filename)
                        sftp.rename(source, destination)
                        print(f"\nCompleted processing\nFile {filename} moved from source: {source} to destination: {destination}")

                        # Move file from the local processing to completed folder
                        local_source = os.path.join(LOCAL_DOWNLOADS, filename)
                        local_destination = os.path.join(LOCAL_OUTPUT, filename)
                        os.makedirs(os.path.dirname(local_destination), exist_ok=True)
                        shutil.move(local_source, local_destination)

                    else:
                        # Move file from the processing to error folder in SFTP server
                        source = os.path.join(SFTP_PROCESSING_FOLDER, 'rdc1', filename)
                        destination = os.path.join(SFTP_ERROR_FOLDER, 'rdc1', filename)
                        sftp.rename(source, destination)
                        print(f"\nFile {filename} moved from source: {source} to destination: {destination}")

                        # Move file from the local processing to error folder
                        local_source = os.path.join(LOCAL_DOWNLOADS, filename)
                        local_destination = os.path.join(LOCAL_ERROR, filename)
                        os.makedirs(os.path.dirname(local_destination), exist_ok=True)
                        shutil.move(local_source, local_destination)

            except Exception as e:
                print(f"\nError processing {source}: {e}")

        else:
            print(f"\nNo user was available to process {filename}")
            print("\nWaiting for 5 minutes...")
            time.sleep(300)
    except Exception as e:
        print(f"\nError during processing for {filename}: {e}")

# Graceful shutdown handler
def handle_exit(signum, frame):
    print("Received exit signal, terminating...")
    exit(0)

# Main function with multiprocessing
def main():
    signal.signal(signal.SIGINT, handle_exit)  # To Handle Ctrl+C
    lock = Lock()
    manager = Manager()
    credentials = manager.list(USERS)  # Shared list for user credentials

    while True:
        monitor_assignment(lock, credentials)
        print("Reconnecting to SFTP...")
        time.sleep(30)  # Check SFTP every 30 seconds
if __name__ == '__main__':
    main()