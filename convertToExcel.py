import os
import shutil
import pandas as pd
from config import DATA_HTML_PATH

def convertToExcel():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    CONFIG_FILE_PATH = os.path.join(script_dir, 'config.py')
    results_dir = os.path.join(script_dir, 'stock_score_results')
    os.makedirs(results_dir, exist_ok=True)

    new_html_path = os.path.join(results_dir, 'data.html')
    try:
        shutil.copy(DATA_HTML_PATH, new_html_path)
    except:
        raise ValueError(f"Invalid HTML path {DATA_HTML_PATH}. Make sure the html path is correct.")

    tables = pd.read_html(new_html_path)

    unit_multipliers = {
        "Lac": 1e5,
        "Lakh": 1e5,
        "Cr": 1e7,
        "Crore": 1e7,
        "Arab": 1e9
    }

    def convert_value(cell):
        if isinstance(cell, str):
            parts = cell.replace(",", "").split()
            if len(parts) == 2 and parts[1] in unit_multipliers:
                try:
                    return float(parts[0]) * unit_multipliers[parts[1]]
                except ValueError:
                    return cell
            else:
                try:
                    return float(cell.replace(",", ""))
                except ValueError:
                    return cell
        return cell

    if tables:
        df = tables[0].applymap(convert_value)
        output_excel_path = os.path.join(results_dir, 'processed_data.xlsx')
        df.to_excel(output_excel_path, index=False)
        print(f"✅ Processed data saved to: {output_excel_path}")
        update_config(output_excel_path, CONFIG_FILE_PATH, script_dir)
    else:
        print("⚠️ No tables found in the HTML file.")

def update_config(excel_path, CONFIG_FILE_PATH, script_dir):
    try:
        with open(CONFIG_FILE_PATH, 'r') as file:
            config_content = file.readlines()

        for i, line in enumerate(config_content):
            if line.startswith('EXCEL_FILE_PATH'):
                config_content[i] = f'EXCEL_FILE_PATH = r"{excel_path}"\n'
                break
        else:
            config_content.append(f'EXCEL_FILE_PATH = r"{excel_path}"\n')

        for i, line in enumerate(config_content):
            if line.startswith('SCRIPT_DIRECTORY'):
                config_content[i] = f'SCRIPT_DIRECTORY = r"{script_dir}"\n'
                break
        else:
            config_content.append(f'SCRIPT_DIRECTORY = r"{script_dir}"\n')

        with open(CONFIG_FILE_PATH, 'w') as file:
            file.writelines(config_content)

        print(f"✅ Updated config.py with new Excel file path and script directory.")
    except Exception as e:
        print(f"⚠️ Failed to update config.py: {e}")

if __name__ == "__main__":
    convertToExcel()
