import os
import shutil
import win32com.client as win32
import argparse
import sys

TEMPLATE_PATH = r"D:\workspace\Import_Images\template\pdc_template.xlsx"
TEMPLATE_NAME = r"pdc_template.xlsx"

DEFAULT_FOLDER_PATH = r"D:\workspace\Import_Images\default_folder"

DEFAULT_PROJECT = r"<Project name>"
DEFAULT_DELIVERY = r"<Delivery ID>"
DEFAULT_SN = r"<Snapshot name>"
DEFAULT_PCM = r"<PCM name>"


PROJECT_ADDRESS                         = "A1"
DELIVERY_ADDRESS                        = "A2"
SN_ADDRESS                              = "A3"
PCM_ADDRESS                             = "A4"
PLANCHECKER_ADDRESS                     = "A5"
PARAVALX_ADDRESS                        = "A8"
BAT_ADDRESS                             = "A11"
BUILDWARNING_ADDRESS                    = "A14"
MEMORY_REPORT_ADDRESS                   = "A17"
CODE_GEN_AND_CHECK_MISMATCH_ADDRESS     = "A20"
EMAIL_ADDRESS                           = "A23"

excel = None
workbook = None
sheet = None

# Define a file object
class File:
    def __init__(self, description: str = "", names: list = None, keyword: str = "", extension: str = "", address: str = ""):
        self.description = description
        self.names = names if names is not None else []
        self.keyword = keyword
        self.extension = extension
        self.address = address


def main() -> int:
    global excel
    global workbook
    global sheet
    
    ret_val = 1 # Default return value is 1, which means the script has failed. It will be set to 0 at the end of the function if everything goes well.
    
    args = parse_args()
    folder_path = args.folder_path
    project_name = args.project_name
    delivery_id = args.delivery_id
    snapshot_name = args.snapshot_name
    pcm_name = args.pcm_name
    
    if (folder_path != DEFAULT_FOLDER_PATH):
        print(f"Using the provided folder path: {folder_path}")
    else:
        print(f"No folder path provided, using the default folder path: {folder_path}")

    try:
        # Check the validity of the folder path
        check_folder(folder_path)
        
        # Convert the path to absolute path
        folder_path = os.path.abspath(folder_path)
        
        # Check if the file already exists in the target folder
        if ( os.path.exists(os.path.join(folder_path, TEMPLATE_NAME)) ):
            print(f"Warning: The file '{TEMPLATE_NAME}' already exists in the target folder '{folder_path}'")
            answer = input("Do you want to overwrite it or not? (y/n): ")
            if answer.lower() != "y" and answer.lower() != "yes":
                ret_val = 0
                print("No overwrite, exiting the script...")
                return ret_val

        # Copy the template to the target folder
        template_file_path = copy_template_to_folder(folder_path)

        # Open the template
        excel, workbook, sheet = open_workbook(template_file_path)
        
        # Get the needed files in the folder
        needed_files = get_files(folder_path)
        
        # Import files to the workbook
        for file in needed_files:
            import_files(folder_path=folder_path, file=file)
            
        
        # Write the project name, delivery ID, snapshot name and PCM name to the corresponding cells
        sheet.Range(PROJECT_ADDRESS).Value = project_name
        sheet.Range(DELIVERY_ADDRESS).Value = delivery_id
        sheet.Range(SN_ADDRESS).Value = snapshot_name
        sheet.Range(PCM_ADDRESS).Value = pcm_name
        print("Project name, delivery ID, snapshot name and PCM name have been written to the workbook.")

    except Exception as e:
        print(f"An error occurred: {e}")
        ret_val = 1
    finally:
        close_excel()
    
    return ret_val

    
# Function to find all the files in the folder based on the keyword and extension
def get_files(folder_path: str) -> list[File]:
    planchecker = File(description = "Plan Checker", keyword="planchecker", extension=".xlsx", address=PLANCHECKER_ADDRESS)
    bat = File(description = "BAT", keyword="bat", extension=".xlsx", address=BAT_ADDRESS)
    paravalx = File(description = "_paravalX", keyword="paraval", extension=".zip", address=PARAVALX_ADDRESS)
    builwarning = File(description = "LatestBuildWarning", keyword="warning", extension=".log", address=BUILDWARNING_ADDRESS)
    memory_report = File(description = "Memory Report", keyword="memory", extension=".html", address=MEMORY_REPORT_ADDRESS)
    code_gen_images = File(description = "Code Gen Images", extension=".png", address=CODE_GEN_AND_CHECK_MISMATCH_ADDRESS)
    email = File(description = "Release Email", extension=".msg", address=EMAIL_ADDRESS)
    needed_files = [planchecker, paravalx, bat, builwarning, memory_report, code_gen_images, email]
    
    for index, file in enumerate(needed_files):
        if( file == code_gen_images or file == email):
            for filename in os.listdir(folder_path):
                if filename.lower().endswith(file.extension):
                    needed_files[index].names.append(filename)
                    print(f"Found file '{filename}' for '{file.description}'")
        else:
            for filename in os.listdir(folder_path):
                if (file.keyword in filename.lower()):
                    needed_files[index].names.append(filename)
                    print(f"Found file '{filename}' for '{file.description}'")
                    
    return needed_files


# Parse the arguments passed into this script
def parse_args():
    parser = argparse.ArgumentParser(description="Create a Checkpoint PDC")
    # Mandatory argument i.e. folder path
    parser.add_argument("folder_path", nargs='?', default=DEFAULT_FOLDER_PATH, help="path of folder containing all the input")
    # Optional arguments
    # Project name e.g. BYD C18 CP23 SA5EB 6D
    parser.add_argument("--project_name", "-p", required=False, default=DEFAULT_PROJECT, help="The project name e.g. BYD C18 CP23 SA5EB 6D")
    # Delivery ID
    parser.add_argument("--delivery_id", "-d", required=False, default=DEFAULT_DELIVERY, help="The delivery ID e.g. 1234567")
    # Snapshot name
    parser.add_argument("--snapshot_name", "-sn", required=False, default=DEFAULT_SN, help="The snapshot name e.g. Sn_abcdefg")
    # PCM name
    parser.add_argument("--pcm_name", "-pcm", required=False, default=DEFAULT_PCM, help="The PCM name e.g. Nguyen Van A")
    
    return parser.parse_args()

# Check the folder path and file existence
def check_folder(folder_path: str)-> bool:
    # Check if the folder path exists
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"The folder path '{folder_path}' does not exist.")

# Copy the template file to the target folder
def copy_template_to_folder(folder_path: str)-> str:
    # Copy the template file to the target folder
    source_file = TEMPLATE_PATH
    target_folder = folder_path
    # Copy the file and handle any exceptions
    try:
        file_path = shutil.copy(source_file, target_folder)
        print(f"Template file '{source_file}' has been copied to '{target_folder}'.")
        # Update the global FILE_PATH variable to point to the new location of the copied file
        return file_path
    except Exception as e:
        raise RuntimeError("Failed to copy the template file") from e

# Open excel and workbook
def open_workbook(name: str):
    print("Opening Excel application...")
    excel = win32.Dispatch("Excel.Application")
    # Make Excel visible
    excel.Visible = True
    print("Open Workbook...")
    try:
        workbook = excel.Workbooks.Open(name)
        # Access a sheet (the first one)
        sheet = workbook.Sheets(1)  # or Sheets("Sheet1")
        return excel, workbook, sheet
    except Exception as e:
        excel.Quit()
        raise RuntimeError("Failed to open the workbook") from e

# Import files to the workbook
def import_files(folder_path: str, file: File):
    global sheet
    if (len(file.names) == 0):
        print(f"Warning: No file '{file.description}' to import")
    else:
        cell = sheet.Range(file.address)
        file_names = file.names
        left_position = cell.Left
        for name in file_names:
            file_path = os.path.join(folder_path, name)
            sheet.Shapes.AddOLEObject(
                Filename = file_path,
                ClassType = None,
                Link = False,
                DisplayAsIcon = True,
                IconFileName = r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE" if (".xlsx" in file.extension) else None,
                IconIndex = 0,
                IconLabel = os.path.basename(file_path),
                Left = left_position,
                Top = cell.Top,
                Width = None,
                Height = None
            )          
            
            left_position += 100 # Move the position for the next image

# Close workbook and excel application
def close_excel():
    global workbook
    global excel
    
    # Closing the workbook if it exists
    if (workbook):
        is_saved = True
        while (True):
            answer = input("The workbook has unsaved changes, do you want to save? (y/n): ")
            if (answer.lower() == "y" or answer.lower() == "yes"):
                print("Saving the workbook..")
                is_saved = True
                break
            elif (answer.lower() == "n" or answer.lower() == "no"):
                print("No save, closing the workbook...")
                is_saved = False
                break
            else:
                print("Invalid option, please choose Y-yes or N-no")
        
        workbook.Close(SaveChanges=is_saved)
    
    #Then closing excel application
    excel.Quit() 
    

if __name__ == '__main__':
    sys.exit(main())