import os
from datetime import datetime, timedelta
import pytz
import configparser
import logging
import time
import pywintypes
# Import win32 modules
import win32serviceutil
import win32service
import win32event
import win32evtlogutil
import win32wnet
import winerror

current_directory = (os.path.dirname(os.path.abspath(__file__)))
local_file_path = os.path.join(current_directory, "service_log.txt")


log_file = local_file_path
logging.basicConfig(filename=log_file, level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

CONNECT_INTERACTIVE = 0x00000008

HOST_NAME = "xx"
# Path for DFS
SHARE_FULL_NAME = r"\\DFS\xx\x\x"

# Function to connect to DFS


def connect_to_dfs(SHARE_USER, SHARE_PWD):
    net_resource = win32wnet.NETRESOURCE()
    net_resource.lpRemoteName = SHARE_FULL_NAME
    flags = 0
    flags |= CONNECT_INTERACTIVE
    logging.info("Trying to create connection to: {:s}".format(SHARE_FULL_NAME))
    try:
        win32wnet.WNetAddConnection2(net_resource, SHARE_PWD, SHARE_USER, flags)
    except pywintypes.error as e:
        logging.error(f"Failed to connect to DFS: {e}")
        return False
    else:
        logging.info("Connected to DFS successfully.")
        return True

# Function to clean files


def clean_files(dfs_path, days_threshold, SHARE_USER, SHARE_PWD):
    # Set timezone
    tz = pytz.timezone('Europe/Warsaw')

    try:
        if not connect_to_dfs(SHARE_USER, SHARE_PWD):
            logging.error("Failed to connect to DFS. Skipping file cleaning.")
            return False

        remote_path = dfs_path

        # List files in DFS directory
        files = os.listdir(remote_path)
        logging.info(f"Listing: {len(files)} files")

        current_date = datetime.now(tz).date()
        logging.info(f"Current date: {current_date}")
        threshold_date = current_date - timedelta(days=days_threshold)
        logging.info(f"Timedelta: {timedelta(days=days_threshold)}")
        logging.info(f"Threshold date: {threshold_date}")

        # Delete files older than the specified date
        for file_name in files:
            try:
                file_path = os.path.join(remote_path, file_name)
                logging.info(f"Processing file: {file_path}")
            except Exception as e:
                logging.error(f"Error listing files for deletion: {(e)}")
                continue

            try:
                modified_time = os.path.getmtime(file_path)
                modified_datetime = datetime.fromtimestamp(modified_time)
                formatted_date = modified_datetime.strftime('%Y-%m-%d')
                file_modified_time = datetime.strptime(formatted_date, '%Y-%m-%d').date()
                logging.info(f"File modified time: {file_modified_time}")

            except Exception as e:
                logging.error(f"Error calculating file modified time: {str(e)}")
                return False



            if file_modified_time < threshold_date:
                os.remove(file_path)
                logging.info(f"Deleted {file_name} from {dfs_path}")

        return True

    except FileNotFoundError as e:
        logging.error(f"Directory not found: {dfs_path}")
    except PermissionError as e:
        logging.error(f"Permission denied for directory: {dfs_path}")
    except Exception as e:
        logging.error(f"An error occurred while cleaning files: {(e)}")


# Function to read configuration
def read_config():

    current_directory = (os.path.dirname(os.path.abspath(__file__)))
    local_file_path = os.path.join(current_directory, "config.ini")

    config = configparser.ConfigParser()
    config_path = local_file_path
    config.read(config_path)
    return config


def main():
    config = read_config()

    # Read configuration
    dfs_path = config.get('Settings', 'path')
    days_threshold = int(config.get('Settings', 'days_threshold'))
    SHARE_USER = config.get('Settings', 'SHARE_USER', fallback=None)
    SHARE_PWD = config.get('Settings', 'SHARE_PWD', fallback=None)

    try:
        if not clean_files(dfs_path, days_threshold, SHARE_USER, SHARE_PWD):
            logging.error("File cleaning operation completed unsuccessfully.")
            return
        logging.info("File cleaning operation completed successfully.")
    except Exception as e:
        logging.error(f"An error occurred while trying to clean files: {str(e)}")


class RemoteFileCleanerService(win32serviceutil.ServiceFramework):
    _svc_name_ = "RemoteFileCleaner"
    _svc_display_name_ = "Remote File Cleaner Service"
    _svc_description_ = "Service to clean remote files older than specified date"

    def __init__(self, args=None):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        self.is_running = True

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # Set self.is_running to False after the cleanup is done
        try:

            self.is_running = False
            win32event.SetEvent(self.hWaitStop)
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
        except Exception as e:
            # Log the error
            win32evtlogutil.ReportEvent(
                'RemoteFileCleaner',
                1,  # EVENTLOG_ERROR_TYPE (numeric constant for error type)
                0,
                1,  # Event category
                [f"An error occurred while stopping service: {str(e)}"],
                None
            )

    def SvcDoRun(self):
        while self.is_running:
            try:
                # Call main in loop to execute code
                main()
            except Exception as e:
                # Log any errors that occur during file cleaning
                logging.error(f"An error occurred while cleaning files: {(e)}")

            # Time between cycles *now 24h*
            time.sleep(60 * 60 * 24)

    def SvcStart(self):

        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        self.is_running = True
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)

    # For timeout
    def StartService(self):

        # Timeout setting
        dwTimeout = 60000

        try:
            hscm = win32service.OpenSCManager(None, None, win32service.SC_MANAGER_ALL_ACCESS)
            hs = win32serviceutil.SmartOpenService(hscm, self._svc_name_, win32service.SERVICE_ALL_ACCESS)
            win32service.StartService(hs, None, dwTimeout)
        except win32service.error as details:
            if details[0] == winerror.ERROR_SERVICE_ALREADY_RUNNING:
                pass
            else:
                raise


if __name__ == '__main__':
    try:
        # Install the service
        win32serviceutil.HandleCommandLine(RemoteFileCleanerService)
    except Exception as e:
        logging.error(f"An error occurred while installing the service: {str(e)}")
