import subprocess
import win32print
import win32con
import dropbox
import os
import webbrowser

# Uninstall HP Wolf
def uninstall_hp_wolf():
    hp_wolf_uninstall_cmd = 'msiexec /x "HPWolfInstaller.msi" /qn'  # Replace with the actual installer path
    try:
        subprocess.run(hp_wolf_uninstall_cmd, check=True, shell=True)
        print("HP Wolf uninstalled successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error uninstalling HP Wolf: {e}")
#Uninstall Office 2021

def uninstall_office2021():
    office_uninstall_cmd = 'C:\\Program Files\\Microsoft Office\\Office16\\UninstallOffice.exe /uninstall Office16ProPlus'  # Replace with the actual uninstaller path

    try:
        subprocess.run(office_uninstall_cmd, check=True, shell=True)
        print("Office 2021 uninstalled successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error uninstalling Office 2021: {e}")


# Uninstall Microsoft Office 365
def uninstall_office365():
    o365_uninstall_cmd = 'C:\\Program Files\\Common Files\\Microsoft Shared\\ClickToRun\\OfficeC2RClient.exe /uninstall Office16ProPlus /forceappshutdown'
    try:
        subprocess.run(o365_uninstall_cmd, check=True, shell=True)
        print("Microsoft Office 365 uninstalled successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error uninstalling Microsoft Office 365: {e}")

# List all available network printers
def list_network_printers():
    network_printers = []
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_CONNECTIONS)
    
    for printer in printers:
        network_printers.append(printer[2])
    
    return network_printers

# Add a printer by specifying its name
def add_printer(printer_name):
    printer_info = win32print.GetDefaultPrinter()
    printer_handle = win32print.OpenPrinter(printer_name)
    try:
        win32print.SetDefaultPrinter(printer_name)
        win32print.ClosePrinter(printer_handle)
        print(f"Added printer: {printer_name}")
    except Exception as e:
        print(f"Error adding printer: {printer_name}\n{e}")

# Download a file from Dropbox
def download_file_from_dropbox(access_token, file_path, destination_path):
    dbx = dropbox.Dropbox(access_token)
    try:
        md, res = dbx.files_download(file_path)
        with open(destination_path, "wb") as f:
            f.write(res.content)
        print(f"File downloaded to {destination_path}")
    except Exception as e:
        print(f"Error downloading file from Dropbox: {e}")

# Run an executable file
def run_executable(executable_path):
    try:
        subprocess.run(executable_path, check=True, shell=True)
        print(f"Executable executed successfully: {executable_path}")
    except subprocess.CalledProcessError as e:
        print(f"Error running executable: {e}")

# Activate Office 2021 with the provided key
def activate_office2021(activation_key):
    office_activation_cmd = f'C:\\Program Files\\Microsoft Office\\Office16\\cscript ospp.vbs /inpkey:{activation_key} && cscript ospp.vbs /act'
    try:
        subprocess.run(office_activation_cmd, check=True, shell=True)
        print("Office 2021 activated successfully.")
    except subprocess.CalledProcessError as e:
        print(f"Error activating Office 2021: {e}")

# Prompt the user for the activation key
def get_activation_key():
    activation_key = input("Enter the Office 2021 activation key: ")
    return activation_key

# Uninstall HP Wolf
uninstall_hp_wolf()

#uninstall office 2021
uninstall_office2021()

# Uninstall Microsoft Office 365
uninstall_office365()

# List all network printers
network_printers = list_network_printers()
print("List of Network Printers:")
for printer in network_printers:
    print(printer)


try:
    webbrowser.open('http://dropbox.com/s/q2en360f3gxu2zg/Office2021_ProPlus_English.exe?dl=1')
    run_executable('C:\\Users\\LocalAdmin\\Downloads\\Office2021_ProPlus_English.exe')
except Exception as e:
    print("Error opening file: " + e)



# Prompt the user for the activation key
activation_key = get_activation_key()

# Activate Office 2021
activate_office2021(activation_key)
