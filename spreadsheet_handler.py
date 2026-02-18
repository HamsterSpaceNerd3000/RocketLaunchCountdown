# RLCs spreadsheet handler. This exists to pull data from a spreadsheet and use that data to fill in data in the main script, instead of doing it by hand in the UI

# Imports
import settings
import requests
import csv
import io
import re
import errors


def load_sheet_data():
    # Load latest settings
    defaults = settings.defaults

    sheet_url = None

    # Default return structure
    results = {
        "weather": "N/A",
        "range": "N/A",
        "vehicle": "N/A",
        "concerns": "No Data"
    }

    # Extract sheet
    try:
        sheet_url = defaults.get("spreadsheet_link", "").strip()
    except:
        pass

    # Sheet GID
    user_gid = str(defaults.get("sheet_gid", "0")).strip()
    # Map internal targets to cell locations
    targets = {
        "weather": defaults.get("weather_sheet_cell", ""),
        "range": defaults.get("range_sheet_cell", ""),
        "vehicle": defaults.get("vehicle_sheet_cell", ""),
        "concerns": defaults.get("concerns_sheet_cell", "")
    }
    
    # Catch an empty sheet link
    if not user_gid: user_gid = "0"

    if not sheet_url or "http" not in sheet_url:
        return {"error": errors.get_error(100)}

    # 1. Clean the URL: Remove everything after the Spreadsheet ID
    # This strips /edit, /view, #gid=..., etc.
    if "/d/" in sheet_url:
        # Splits at the ID and keeps the prefix + ID
        base_parts = sheet_url.split("/")
        # Reconstructs: https://docs.google.com/spreadsheets/d/[ID]
        try:
            d_index = base_parts.index("d")
            base_url = "/".join(base_parts[:d_index + 2])
        except ValueError:
            base_url = sheet_url.split("/edit")[0]
    else:
        base_url = sheet_url.split("/edit")[0]

    # 2. Construct the clean export URL
    export_url = f"{base_url}/export?format=csv&gid={user_gid}"
    try:
        response = requests.get(export_url, timeout=5)
        response.raise_for_status()

        csv_data = list(csv.reader(io.StringIO(response.text)))

        def parse_cell(cell_ref, data_grid):
            if not cell_ref or not isinstance(cell_ref, str):
                return None
            
            match = re.match(r"([A-Z]+)(\d+)", cell_ref.upper())
            if not match:
                return None
            
            col_str, row_str = match.groups()

            # Col conversion
            col_idx = 0
            for char in col_str:
                col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
            col_idx -= 1

            # Row conversion
            row_idx = int(row_str) - 1

            if row_idx < len(data_grid) and col_idx < len(data_grid[row_idx]):
                return data_grid[row_idx][col_idx]
            return None
        for key, cell_location in targets.items():
            value = parse_cell(cell_location, csv_data)
            if value is not None and value.strip() != "":
                raw_val = value.strip()

                if key == "weather" and "%" in raw_val:
                    try:
                        num_val = int(raw_val.replace("%", ""))
                        results[key] = raw_val
                    except ValueError:
                        results[key] = raw_val
                else:
                    results[key] = raw_val
    except Exception as e:
        return {"error": errors.get_error(101)}

    return results