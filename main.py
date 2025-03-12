import pandas as pd
import json

def read_excel_file(file_path):
    """Reads an Excel file and returns a dictionary of dataframes for each sheet."""
    try:
        data = pd.read_excel(file_path, sheet_name=None)
        return data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def convert_to_json(data):
    """Converts a dictionary of dataframes into a JSON-compatible format."""
    if data is None:
        return None
    try:
        json_data = {sheet: df.to_dict(orient='records') for sheet, df in data.items()}
        return json_data
    except Exception as e:
        print(f"Error converting data to JSON: {e}")
        return None

def save_json_file(json_data, output_path):
    """Saves JSON data to a file."""
    if json_data is None:
        print("No JSON data to save.")
        return
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(json_data, f, indent=4, ensure_ascii=False)
        print(f"JSON file '{output_path}' created successfully!")
    except Exception as e:
        print(f"Error saving JSON file: {e}")

def excel_to_json(excel_file, json_file):
    """Main function to convert Excel to JSON."""
    data = read_excel_file(excel_file)
    json_data = convert_to_json(data)
    save_json_file(json_data, json_file)

# Example usage
if __name__ == "__main__":
    input_excel = 'input.xlsx'
    output_json = 'output.json'
    excel_to_json(input_excel, output_json)
