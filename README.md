# Excel Sheet Password Remover

A comprehensive Python utility to remove passwords and protection from Excel files (.xlsx format) using advanced XML manipulation techniques.

## ⚠️ Important Warning

**This script may cause minor layout changes to your Excel files.** When Excel files are processed and rebuilt from their internal XML structure, some formatting or layout adjustments might occur. Please ensure you have backups of your original files before running this tool.

## Features

- **Complete Protection Removal**: Handles both worksheet protection and workbook-level protection
- **Advanced XML Manipulation**: Directly modifies Excel's internal XML structure for thorough protection removal
- **Batch Processing**: Process multiple Excel files at once
- **Flexible Output Options**: Specify custom output paths or use automatic naming
- **Safe Processing**: Uses temporary directories to avoid corrupting original files
- **Class-Based Architecture**: Clean, object-oriented design for easy integration

## Prerequisites

Make sure you have the following Python packages installed:

```bash
pip install pandas openpyxl
```

## Installation

1. Clone this repository:
```bash
git clone https://github.com/SakibAhmedShuva/Excel-Sheet-Password-Remover.git
cd Excel-Sheet-Password-Remover
```

2. Install the required dependencies:
```bash
pip install pandas openpyxl
```

## Usage

### Method 1: Using Jupyter Notebook (Primary Method)

1. Open the notebook:
```bash
jupyter notebook password_remover.ipynb
```

2. Run all cells to load the classes and functions

3. Modify the file path in the last cell's example usage section:
```python
# Change this line to your file path
result = remove_excel_passwords(r'C:\path\to\your\protected_file.xlsx')
if result:
    print(f"Success! Unprotected file saved as: {result}")
else:
    print("Failed to remove protection")
```

4. Run the cell to remove protection from your file

### Method 2: Extract and Use as Python Script

If you want to use this as a standalone Python script:

1. Copy the code from the notebook cell into a new file called `password_remover.py`
2. Then you can use it as:

```python
# Simple function usage
result = remove_excel_passwords(r'path/to/your/protected_file.xlsx')

# Batch processing
unprotected_files = batch_remove_passwords('/path/to/directory')

# Using the class directly
remover = ExcelPasswordRemover()
result = remover.remove_all_protection('input.xlsx', 'output.xlsx')
```

### Method 3: Direct Usage in Jupyter

You can also create new cells in the notebook and use the functions directly:

```python
# Remove protection from a single file
result = remove_excel_passwords(r'your_file_path_here.xlsx')

# Batch process multiple files
results = batch_remove_passwords(r'C:\your\directory\path')

# Use the class for more control
remover = ExcelPasswordRemover()
result = remover.remove_all_protection('input.xlsx', 'custom_output.xlsx')
```

## How It Works

This tool uses a sophisticated approach to remove Excel protection:

1. **XML Extraction**: Extracts the Excel file's internal XML structure (Excel files are essentially ZIP archives)
2. **Protection Removal**: 
   - Removes `<workbookProtection>` elements from `workbook.xml`
   - Removes `<sheetProtection>` elements from individual worksheet XML files
3. **File Reconstruction**: Rebuilds the Excel file from the modified XML structure
4. **Final Cleanup**: Uses openpyxl to ensure all protection is completely removed

## Output File Naming

- `remove_worksheet_protection()`: `filename_unprotected.xlsx`
- `remove_workbook_protection()`: `filename_unprotected.xlsx`
- `remove_all_protection()`: `filename_fully_unprotected.xlsx`

## Supported File Types

- `.xlsx` - Excel Workbook format

**Note**: This tool specifically targets .xlsx files. For .xls files, convert them to .xlsx format first.

## Class Structure

### ExcelPasswordRemover Class

**Methods:**
- `remove_worksheet_protection(file_path, output_path=None)` - Removes sheet-level protection
- `remove_workbook_protection(file_path, output_path=None)` - Removes workbook-level protection  
- `remove_all_protection(file_path, output_path=None)` - Removes all types of protection
- `_remove_sheet_protection_from_xml(xml_path)` - Internal method for XML manipulation

## Example Output

```
Success! Unprotected file saved as: C:\Users\Sakib\Downloads\fm200-calculator_fully_unprotected.xlsx
```

## Error Handling

The tool includes comprehensive error handling:
- File access errors
- XML parsing errors
- ZIP file corruption issues
- Missing file/directory handling

## Safety Recommendations

1. **Always create backups** of your original files before processing
2. Test the script on a copy of your file first
3. Verify that the unprotected file contains all expected data
4. The tool uses temporary directories to avoid corrupting original files during processing

## Limitations

- Designed specifically for .xlsx format
- May alter some formatting due to XML reconstruction process
- Requires write permissions in the output directory

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License.

## Disclaimer

This tool is intended for legitimate use cases such as recovering access to your own files when passwords are forgotten. Users are responsible for ensuring they have the right to remove protection from any files they process.

## Author

**SakibAhmedShuva**
- GitHub: [@SakibAhmedShuva](https://github.com/SakibAhmedShuva)

---

⭐ If you find this tool useful, please consider giving it a star!

## Troubleshooting

### Common Issues

1. **"Permission denied" error**: Ensure the Excel file is not open in another program
2. **Module not found**: Install required dependencies: `pip install pandas openpyxl`
3. **File corruption**: The tool uses temporary directories to prevent corruption of original files
4. **XML parsing errors**: Ensure the input file is a valid .xlsx file

### Need Help?

If you encounter any issues, please [open an issue](https://github.com/SakibAhmedShuva/Excel-Sheet-Password-Remover/issues) on GitHub.
