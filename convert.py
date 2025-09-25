import json
import pandas as pd

def convert_json_to_excel(json_file_path, excel_file_path):
    """
    Converts a JSON file with multiple data sources into a single Excel file
    with separate sheets for each data source.

    Args:
        json_file_path (str): The path to the input JSON file.
        excel_file_path (str): The path where the output Excel file will be saved.
    """
    try:
        # Load the JSON data from the file
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)

        # Create a Pandas Excel writer object
        with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
            # Iterate through each key (data source) in the JSON data
            for source_name, source_data in data.items():
                # Convert the data for the current source into a DataFrame
                df = pd.DataFrame(source_data)

                # Write the DataFrame to a new sheet in the Excel file
                # The sheet name will be the same as the data source name
                df.to_excel(writer, sheet_name=source_name, index=False)

        print(f"✅ Success! Data has been written to '{excel_file_path}' with separate sheets for each source.")

    except FileNotFoundError:
        print(f"❌ Error: The file '{json_file_path}' was not found.")
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")

# --- Configuration ---
# Set the paths for your input and output files
# Make sure 'SatrioDwiSetiawan_V3925035.json' is in the same directory as this script, or provide the full path
json_file = 'SatrioDwiSetiawan_V3925035.json'
excel_file = 'Data_Hoax_Analisis.xlsx'

# Run the conversion function
convert_json_to_excel(json_file, excel_file)