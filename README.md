# Automating_TIBA_Corporation_Workflow
 A Python software to automate the workflow of TIBA corporation which works with architects and engineers, upon each one of their project they receive a BOQ sheet (Bill of Quantities) which have all the materials information that they need to make a building, the company has a spreadsheet 'Reference File' for the most common materials and their prices and other insider information. the software takes these two files and scans them then collects needed data from each one of them and provides the company with three sheets with the extension of xls, 1st is a Material Cost Calculation sheet with provides the known materials, their cost, total cost , and the final cost of all the total costs, 2nd is a Requesting Quotes sheet which collects all the data that needs to request quotes equipped with and their volume and units. 3rd is a Construction Schedule sheet which provides them with the schedule of the materials that are valid in this project and guide their employees upon each week of the project. the software also has a GUI to facilitate the user experience. 

 This code provides a graphical user interface (GUI) for performing material cost calculations based on input files. The GUI is implemented using the PyQt5 library in Python.

 Dependencies
The code requires the following libraries to be installed:

re
sys
pandas
PyQt5
openpyxl
You can install these dependencies using the pip package manager:
pip install pandas PyQt5 openpyxl

How to Use
Run the code by executing the script. Make sure you have the required dependencies installed.
python theproject.py

The GUI window will appear with the title "Material Cost Calculation" and a company logo displayed at the top.

BOQ Sheet:

Click the "Browse" button next to the "BOQ Sheet" label.
Select an Excel file (.xlsx) containing the Bill of Quantities (BOQ) data.
The file path will be displayed in the corresponding text entry.
Reference Sheet:

Click the "Browse" button next to the "Reference Sheet" label.
Select an Excel file (.xlsx) containing the reference sheet data.
The file path will be displayed in the corresponding text entry.
Process Files:

Click the "Process Files" button to start processing the selected files.
The code will read the BOQ and reference sheet data, perform calculations, and generate output.
Processed Data:

The processed data will be displayed in the "Processed Data" section of the GUI.
It includes the following information:
BOQ Data: Original material data from the BOQ file.
Matched Data: Material data matched with reference sheet data, including calculated values.
Non-Matched Materials: Material items from the BOQ that could not be matched with the reference sheet.
Construction Schedule: Schedule data for the materials, including start week, end week, and number of weeks.
Output Files:

The code will generate three output Excel files in the current directory:
"Material Cost Calculation.xlsx": Contains the matched data with calculated totals.
"Requesting Quotes.xlsx": Contains the non-matched material items for requesting quotes.
"Construction Schedule.xlsx": Contains the construction schedule data.
Success/Error Messages:

If the processing is successful, a message box will appear with the text "Files processed successfully. Output files saved."
If an error occurs during processing, an error message box will appear with details of the error.

Note:
The code assumes that the BOQ and reference sheet files are in the .xlsx format.
Ensure that the BOQ file contains columns with material, volume, and unit information.
The code uses regular expressions to remove bold text formatting from the material column in the BOQ file.
The output files are saved in the current directory with the specified file names.


