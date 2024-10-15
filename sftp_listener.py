# Necessary Libraries
import os
import time
import signal
import pandas as pd
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


# Disable host key checking
cnopts = pysftp.CnOpts()
cnopts.hostkeys = None

# SFTP and RDC Folder details
SFTP_HOST = credfile.host
SFTP_USERNAME = credfile.username
SFTP_PASSWORD = credfile.password

SFTP_FOLDER = paths.sftp_upload_folder
SFTP_ASSIGNMET_FOLDER = paths.sftp_assignment_folder
SFTP_PROCESSING_FOLDER = paths.sftp_processing_folder
RDC_FOLDERS = ['rdc1', 'rdc2', 'rdc3']
MAX_FILES_PER_RDC = 2

# Initialize a global variable to track the last assigned RDC index
last_assigned_rdc_index = 0

# SFTP Monitoring Function
def monitor_sftp():
    print("Connecting to SFTP...")
    global last_assigned_rdc_index
    try:
        with pysftp.Connection(SFTP_HOST, username=SFTP_USERNAME, password=SFTP_PASSWORD, cnopts=cnopts) as sftp:
            print("Connected!")
            files = sftp.listdir(SFTP_FOLDER)
            print(f"\nFiles available in uploads folder:\n{files}")

            for file in files:
                # Check file counts for each RDC in the assignment folder
                rdc_file_counts_assignment = get_rdc_file_counts(sftp, SFTP_ASSIGNMET_FOLDER)

                # Attempt to assign the file to the next RDC in round-robin fashion
                for _ in range(len(RDC_FOLDERS)):
                    rdc = RDC_FOLDERS[last_assigned_rdc_index]
                    # Check if the current RDC can accept more files
                    if rdc_file_counts_assignment[rdc] < MAX_FILES_PER_RDC:
                        assign_file_to_rdc(sftp, file, rdc)
                        # Update the index for the next round
                        last_assigned_rdc_index = (last_assigned_rdc_index + 1) % len(RDC_FOLDERS)
                        break
                else:
                    print(f"All RDCs are full. Cannot assign file '{file}' at this moment.")
                    print("Waiting for 10 minutes...")
                    time.sleep(600)
    except Exception as e:
        print(f"\nError with SFTP connection or file handling: {e}")

# Function to get file counts in RDC folders
def get_rdc_file_counts(sftp, base_folder):
    rdc_file_counts = {}
    for rdc in RDC_FOLDERS:
        folder_path = os.path.join(base_folder, rdc)
        try:
            files_in_rdc = sftp.listdir(folder_path)
            rdc_file_counts[rdc] = len(files_in_rdc)
        except FileNotFoundError:
            # If folder doesn't exist, it's empty
            rdc_file_counts[rdc] = 0
    return rdc_file_counts

# Function to assign file to available RDC
def assign_file_to_rdc(sftp, file, rdc):
    # Move file from uploads to the selected RDC in the assignment folder
    src_path = os.path.join(SFTP_FOLDER, file)
    dest_path = os.path.join(SFTP_ASSIGNMET_FOLDER, rdc, file)
    
    print(f"\nAssigning file '{file}' to {rdc}")
    sftp.rename(src_path, dest_path)

# Graceful shutdown handler
def handle_exit(signum, frame):
    print("Received exit signal, terminating...")
    exit(0)

# Main function with multiprocessing
def main():
    signal.signal(signal.SIGINT, handle_exit)  # To Handle Ctrl+C 
    while True:
        monitor_sftp()
        print("Reconnecting to SFTP...")
        time.sleep(30)  # Check SFTP every 30 seconds
if __name__ == '__main__':
    main()
