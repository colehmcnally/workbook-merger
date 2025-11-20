# Goal is to combine all excel files in a single root folder
# If there are subfolders within the root folder, it pulls all excel files from that subfolder as well
# All sheets from each file should be added to the master file
# Order of the files added is based on alphabetical naming order, and sheets within each file are added in the order they appear in the file
# User should be able to specify the root folder and the output file name

import os

# Function to get user input for root folder
def input_root_folder():
    root_folder = input("Please enter the root folder path: ")
    return root_folder


# Once the root folder is specified, this function finds all files within the root folder and its subfolders that are excel files
def find_excel_files(root_folder):
    excel_files = []
    for dirpath, dirnames, filenames in os.walk(root_folder): # for each folder, subfolder, and file within the root folder
        for filename in filenames:
            if filename.endswith('.xlsx') or filename.endswith('.xls'): # check if the file is an excel file
                excel_files.append(os.path.join(dirpath, filename)) # add the excel file to the list by joining the directory path and filename, making it a callable path
    print(excel_files) # print the list of excel files found
    return list(excel_files)


# Now that we have the list of excel files, we can combine them into a single master file
import pandas as pd
def combine_excel_files(excel_files, root_folder):
    output_path = os.path.join(root_folder, 'combined_excel.xlsx')
    with pd.ExcelWriter(output_path) as writer: # Create a new excel file to write the combined data to
        for file in excel_files: # For each file in the list of excel files
            
            filename = os.path.basename(file).replace('.xlsx', '').replace('.xls', '') # Get the filename without the path and extension
            
            sheets_dict = pd.read_excel(file, sheet_name=None) # Read all of the sheets in each excel file
            
            
            for sheet_name, df in sheets_dict.items(): # For each sheet in the dictionary
                unique_sheet_name = f"{filename}_{sheet_name}" # Create a unique sheet name by combining the filename and sheet name
                df.to_excel(writer, sheet_name=unique_sheet_name, index=False) # Write the sheet to the new excel file


# Main execution
folder = input_root_folder() # Get the root folder from user input
excel_files = find_excel_files(folder) # Get the list of excel files
excel_files.sort() # Sort the list of excel files in alphabetical order
combine_excel_files(excel_files, folder) # Combine the excel files into a single master file
print("Excel files combined successfully into 'combined_excel.xlsx'")