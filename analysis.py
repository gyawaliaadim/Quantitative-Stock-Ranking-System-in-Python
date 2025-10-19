import pandas as pd
import os
import time
import config

time.sleep(2)

def analysis():
    try:
        df = pd.read_excel(config.EXCEL_FILE_PATH)
    except:
        print(f'''
Error: Excel file not found.
It may take time for the Excel path to be ready.
Try running main.py again!''')
        exit()

    required_cols = ['Close','EPS', 'P/E Ratio', 'Net Profit', 'Net Worth']
    for col in required_cols:
        if col not in df.columns:
            print(f"Error: Required column '{col}' is missing from the Excel data.")
            exit()
        df[col] = pd.to_numeric(df[col], errors='coerce')

    # Normalize EPS (higher is better)
    df['Normalized EPS'] = df['EPS'] / df['Close']

    # Normalize P/E Ratio (lower is better)
    df['Normalized P/E'] = 0.0
    positive_pe = df[df['P/E Ratio'] > 0]['P/E Ratio']
    negative_pe = df[df['P/E Ratio'] < 0]['P/E Ratio']

    if not positive_pe.empty:
        min_positive_pe = positive_pe.min()
        df.loc[df['P/E Ratio'] > 0, 'Normalized P/E'] = min_positive_pe / df.loc[df['P/E Ratio'] > 0, 'P/E Ratio']

    if not negative_pe.empty:
        min_negative_pe = negative_pe.min()
        df.loc[df['P/E Ratio'] < 0, 'Normalized P/E'] = -1 * (df.loc[df['P/E Ratio'] < 0, 'P/E Ratio'] / min_negative_pe)

    # Normalize Net Profit and Net Worth
    max_net_profit = df['Net Profit'].max()
    df['Normalized Net Profit'] = df['Net Profit'] / max_net_profit if max_net_profit != 0 else 0

    max_net_worth = df['Net Worth'].max()
    df['Normalized Net Worth'] = df['Net Worth'] / max_net_worth if max_net_worth != 0 else 0

    # Calculate Composite Score
    df['Composite Score'] = (
        df['Normalized EPS'] * config.WEIGHTS['EPS'] +
        df['Normalized P/E'] * config.WEIGHTS['P/E Ratio'] +
        df['Normalized Net Profit'] * config.WEIGHTS['Net Profit'] +
        df['Normalized Net Worth'] * config.WEIGHTS['Net Worth']
    )

    df_sorted = df.sort_values(by='Composite Score', ascending=False)

    # Save results
    main_output_path = os.path.join(config.OUTPUT_PATH, "Sorted_Companies.xlsx")
    results_folder = os.path.join(config.SCRIPT_DIRECTORY, "stock_score_results")
    os.makedirs(results_folder, exist_ok=True)
    copy_output_path = os.path.join(results_folder, "Sorted_Companies.xlsx")

    try:
        df_sorted.to_excel(main_output_path, index=False)
    except:
        raise ValueError(f"Invalid output directory {config.OUTPUT_PATH}, make sure it exists")

    df_sorted.to_excel(copy_output_path, index=False)

    print(f"✅ Scoring complete!")
    print(f"   Sorted companies saved to: {main_output_path}")
    print(f"   Copy also saved to: {copy_output_path}")

if __name__ == "__main__":
    analysis()
