Below is the updated `README.md` file tailored to match the provided project details, code, and description. It incorporates the provided information, ensures clarity, and follows the structure outlined in the example while addressing the specific files and functionality.

---

# Quantitative Stock Ranking System

A Python-based tool to convert financial data from HTML tables to Excel and rank companies based on key financial metrics such as EPS, P/E Ratio, Net Profit, and Net Worth. The system normalizes these metrics, applies user-defined weights, and generates a composite score to rank companies, outputting the results to Excel files.

- **Sorted results are saved in the `stock_score_results` folder as `Sorted_Companies.xlsx`.**

## Table of Contents
- [Project Overview](#project-overview)
- [Features](#features)
- [Requirements](#requirements)
- [Setup Instructions](#setup-instructions)
- [Usage](#usage)
- [File Structure](#file-structure)
- [How It Works](#how-it-works)
- [License](#license)
- [Example Output](#example-output)
- [Screenshot of the Excel File](#screenshot-of-the-excel-file)

## Project Overview
This project provides a quantitative approach to rank companies based on financial metrics extracted from an HTML table. It processes the data, converts it to an Excel file, and performs a scoring analysis using normalized values of Earnings Per Share (EPS), Price-to-Earnings (P/E) Ratio, Net Profit, and Net Worth. User-defined weights are applied to compute a composite score for ranking companies. The results are saved as Excel files in a specified output directory and a dedicated results folder (`stock_score_results`). The system is configurable through a `config.py` file, allowing users to adjust file paths and weights.

## Features
- **HTML to Excel Conversion**: Converts financial data from an HTML table to an Excel file, handling units like "Lac" and "Crore."
- **Data Normalization**: Normalizes EPS, P/E Ratio, Net Profit, and Net Worth for consistent scoring.
- **Weighted Scoring**: Computes a composite score using user-defined weights for each metric.
- **Error Handling**: Validates file paths and required columns, providing clear error messages.
- **Configurable**: Allows customization of file paths and weights via `config.py`.
- **Output Management**: Saves sorted results to both the user-specified output directory and the `stock_score_results` folder.

## Requirements
- Python 3.6 or higher
- Required Python packages:
  - `pandas`
  - `openpyxl`
  - `shutil` (standard library)
- An HTML file (`data.html`) containing financial data in a tabular format (sample data for Q4 2082 provided).
- A valid output directory for saving results.

## Setup Instructions
1. **Clone the Repository**:
   ```bash
   git clone https://github.com/gyawaliaadim/Quantitative-Stock-Ranking-System-in-Python.git
   cd Quantitative-Stock-Ranking-System-in-Python
   ```

2. **Install Dependencies**:
   Install the required Python packages using pip:
   ```bash
   pip install pandas openpyxl
   ```

3. **Prepare the HTML File**:
   Ensure the `data.html` file contains financial data in a table format with the following required columns:
   - `Close`
   - `Listed Share`
   - `EPS`
   - `P/E Ratio`
   - `Net Profit`
   - `Net Worth`
   A sample `data.html` file for Q4 2082 is provided in the repository.

4. **Configure the System**:
   Edit `config.py` to set the correct paths and weights:
   - `DATA_HTML_PATH`: Path to the input HTML file.
   - `OUTPUT_PATH`: Directory where sorted results will be saved.
   - `WEIGHTS`: Dictionary of weights for EPS, P/E Ratio, and Net Profit (must sum to 1).

   Example `config.py`:
   ```python
   DATA_HTML_PATH = r"C:\Users\YourName\Stock market project\data.html"
   OUTPUT_PATH = r"C:\Users\YourName\Downloads"
   WEIGHTS = {
       'EPS': 0.35,
       'P/E Ratio': 0.30,
       'Net Profit': 0.35
   }
   ```

5. **Verify Paths**:
   Ensure the HTML file and output directory exist and are accessible.

## Usage
1. **Run the Main Script**:
   Execute `main.py` to convert the HTML data to Excel and rank companies:
   ```bash
   python main.py
   ```
   This will:
   - Copy the HTML file to the `stock_score_results` folder.
   - Convert the HTML table to `processed_data.xlsx`.
   - Update `config.py` with the Excel file path and script directory.
   - Analyze the data, normalize metrics, calculate composite scores, and save sorted results to `Sorted_Companies.xlsx`.

2. **Check Outputs**:
   - `stock_score_results/processed_data.xlsx`: Processed financial data from the HTML file.
   - `stock_score_results/Sorted_Companies.xlsx`: Ranked companies with composite scores.
   - `OUTPUT_PATH/Sorted_Companies.xlsx`: Copy of the sorted results in the user-specified directory.

## File Structure
```
Quantitative-Stock-Ranking-System-in-Python/
│
├── config.py                    # Configuration file for paths and weights
├── main.py                      # Main script to run conversion and analysis
├── analysis.py                  # Script for data normalization and ranking
├── convertToExcel.py            # Script for HTML to Excel conversion
├── stock_score_results/         # Folder for storing processed data and results
│   ├── data.html                # Copied input HTML file
│   ├── processed_data.xlsx      # Converted Excel file
│   ├── Sorted_Companies.xlsx    # Sorted results with composite scores
├── data.html                    # Input HTML file with financial data
├── README.md                    # Project documentation
```

## How It Works
1. **HTML to Excel Conversion** (`convertToExcel.py`):
   - Reads the HTML file using `pandas.read_html()`.
   - Converts financial values with units (e.g., "Lac" = 100,000, "Crore" = 10,000,000) to numeric format.
   - Saves the processed data to `processed_data.xlsx` in the `stock_score_results` folder.
   - Updates `config.py` with the Excel file path and script directory.

2. **Data Analysis and Ranking** (`analysis.py`):
   - Reads the Excel file (`processed_data.xlsx`).
   - Validates required columns (`Close`, `Listed Share`, `EPS`, `P/E Ratio`, `Net Profit`, `Net Worth`).
   - Normalizes financial metrics:
     - **EPS**: Divided by stock price and scaled by the maximum value (higher is better).
     - **P/E Ratio**: Inverted for positive values, scaled for negative values (lower is better).
     - **Net Profit**: Divided by the product of Net Worth and Listed Share, then scaled.
   - Computes a composite score using weights from `config.py`.
   - Sorts companies by composite score and saves results to `Sorted_Companies.xlsx`.

3. **Error Handling**:
   - Checks for the existence of the HTML file and output directory.
   - Validates required columns in the Excel data.
   - Provides clear error messages for missing files or columns.

## License
This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.

## Example Output
```bash
> python main.py
Converting the html to excel
✅ Processed data saved to: c:\Users\Aadim Gyawali\Stock market project\stock_score_results\processed_data.xlsx
✅ Updated config.py with new Excel file path and script directory.
Analyzing the excel.
✅ Scoring complete!
   Sorted companies saved to: C:\Users\Aadim Gyawali\Downloads\Sorted_Companies.xlsx
   Copy also saved to: c:\Users\Aadim Gyawali\Stock market project\stock_score_results\Sorted_Companies.xlsx
```

## Screenshot of the Excel File
![Sorted Companies Excel](https://github.com/user-attachments/assets/e2397d77-b4a7-4f5f-be88-3efc422843e6)

---

This `README.md` is structured to be clear, concise, and consistent with the provided project details. It includes all necessary sections, reflects the provided code and output, and incorporates the screenshot link. Let me know if you need further refinements or additional details!
