# Excel Variance Finder
This program automates the process of finding and saving filtered rows based on variance criteria from Excel files in a specified folder.

## Usage
### Prerequisites
* Python 3.x
* pandas library
* openpyxl library (required by xlsxwriter for Excel writing)
### Setup
1. Clone this repository to your local machine.
2. Install the required dependencies using pip install pandas openpyxl.
3. Ensure your Excel files to be checked are located in the To-Check folder and the output folder Checked-Reports exists.
## Running the Program
1. Navigate to the directory containing the script ExcelVarianceFinder.py.
2. Run the program using python ExcelVarianceFinder.py.
## Output
1. Filtered rows based on the variance criteria (variable_one and variable_two) will be saved in the Checked-Reports folder.
2. Each output file will be named after the corresponding input file with ' Variance Report.xlsx' appended.
## Notes
* Ensure all input Excel files (*.xls or *.xlsx) have a sheet named 'Sheet1' where the variance filtering will be applied.
* Adjust the values of variable_one and variable_two as needed to customize the variance criteria.
* Modify file paths (folder_path and output_path) in the script if your folder structure differs.
