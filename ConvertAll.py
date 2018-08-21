#%% Convert a Folder of PowerPoint PPTs to PDFs

# Purpose: Converts all PowerPoint PPTs in a folder to Adobe PDF

# Author:  Matthew Renze

# Usage:   python.exe ConvertAll.py input-folder output-folder
#   - input-folder = the folder containing the PowerPoint files to be converted
#   - output-folder = the folder where the Adobe PDFs will be created

# Example: python.exe ConvertAll.py C:\InputFolder C:\OutputFolder

# Note: Also works with PPTX file format

#%% Import libraries
import sys
import os
import comtypes.client

#%% Get console arguments
input_folder_path = sys.argv[1]
output_folder_path = sys.argv[2]

#%% Convert folder paths to Windows format
input_folder_path = os.path.abspath(input_folder_path)
output_folder_path = os.path.abspath(output_folder_path)

#%% Get files in input folder
input_file_paths = os.listdir(input_folder_path)

#%% Convert each file
for input_file_name in input_file_paths:

    # Skip if file does not contain a power point extension
    if not input_file_name.lower().endswith((".ppt", ".pptx")):
        continue
    
    # Create input file path
    input_file_path = os.path.join(input_folder_path, input_file_name)
        
    # Create powerpoint application object
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    
    # Set visibility to minimize
    powerpoint.Visible = 1
    
    # Open the powerpoint slides
    slides = powerpoint.Presentations.Open(input_file_path)
    
    # Get base file name
    file_name = os.path.splitext(input_file_name)[0]
    
    # Create output file path
    output_file_path = os.path.join(output_folder_path, file_name + ".pdf")
    
    # Save as PDF (formatType = 32)
    slides.SaveAs(output_file_path, 32)
    
    # Close the slide deck
    slides.Close()
    
    