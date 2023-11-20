# ExcelVarianceFinder
An excel variance finder to parse through an excel stocktake report file created by H&L and pull out rows that meet variance parameters.

The script has been designeed to read Excel files from a specified folder, filter rows based on certain conditions, and write the filtered rows to a new Excel file 
The conditions for filtering are related to the 'Quantity Difference' column in the Excel files, and the filtered rows are saved in a sheet named 'Filtered Rows' in the new Excel file.

The script was written in order to filter out rows in a stocktake report that meet a certain parameter in order to reduce a stocktake variance report from approximately 20-30 pages to 1-2 pages.

## Usage

1. Set the input parameters in the script:
   - `folder_path`: Specify the folder containing the input Excel files.
   - `output_path`: Specify the folder where the filtered reports will be saved.
   - `variable_one` and `variable_two`: Define the range of acceptable variance in the 'Quantity Difference' column.

2. Run the script using a Python interpreter:
   ```bash
   python excel_variance_finder.py
