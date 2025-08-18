# Unqork JSON Generator

## Requirements
- Python 3.9+
- pip install pandas openpyxl

## How to Run
1. Open terminal in this folder:
   cd "C:\Users\<YourUser>\Downloads\unqork_json_bundle"

2. Run the script:
   python excel_to_unqork.py --excel aggregated_components_uniqueId.xlsx --out unqork_output

3. The script will create one JSON file per `Module_path` in the `unqork_output` folder.
