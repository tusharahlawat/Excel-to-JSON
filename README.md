# Excel-to-JSON

A Python script that converts Excel files (including multiple sheets) into structured JSON format. It reads the Excel file, processes each sheet, and exports the data to a JSON file.

## Features
- Reads all sheets from an Excel file.
- Converts sheet data into JSON format.
- Saves the output as a well-structured JSON file.
- Handles exceptions for better error management.

## Prerequisites
Ensure you have Python installed along with the required dependencies:
```sh
pip install pandas openpyxl
```

## Usage
1. Place your Excel file (e.g., `input.xlsx`) in the same directory as the script.
2. Run the script:
```sh
python script.py
```
3. The JSON output file (`output.json`) will be created in the same directory.

## Code Explanation
- **read_excel_file(file_path)**: Reads an Excel file and loads its sheets into a dictionary.
- **convert_to_json(data)**: Converts the loaded data into JSON format.
- **save_json_file(json_data, output_path)**: Saves the JSON data to a file.
- **excel_to_json(excel_file, json_file)**: Orchestrates the entire conversion process.

## Example
```python
excel_to_json('input.xlsx', 'output.json')
```

## Error Handling
- If the Excel file is missing or corrupt, an error message is displayed.
- If the JSON conversion fails, an appropriate error message is logged.

## License
This project is licensed under the MIT License.

## Contributions
Feel free to contribute by submitting pull requests or reporting issues!

## Author
Tushar Ahlawat
