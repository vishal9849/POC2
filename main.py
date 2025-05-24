import pandas as pd

def calculate_amount_difference(file_path_april, file_path_may, output_file_path="difference_report.xlsx"):
    """
    Compares two Excel files, calculates the difference in 'Sum of Amount in local cur.'
    for matching 'DocumentNo', and saves the result to a new Excel file.

    Args:
        file_path_april (str): Path to the Excel file for April.
        file_path_may (str): Path to the Excel file for May.
        output_file_path (str): Path where the output Excel file will be saved.
    """
    try:
        # Read the Excel files
        df_april = pd.read_excel(file_path_april)
        df_may = pd.read_excel(file_path_may)

        # Ensure 'DocumentNo' is treated as a string to avoid merging issues with different data types
        df_april['DocumentNo'] = df_april['DocumentNo'].astype(str)
        df_may['DocumentNo'] = df_may['DocumentNo'].astype(str)

        # Merge the two dataframes based on 'DocumentNo'
        # We use an outer merge to keep all records from both files.
        # This will create NaN values for DocumentNo not present in both files.
        merged_df = pd.merge(df_april, df_may, on='DocumentNo', how='outer', suffixes=('_April', '_May'))

        # Rename the 'Sum of Amount in local cur.' columns for clarity after merge
        # This step is crucial if the column names were not uniquely suffixed during merge.
        # However, if 'suffixes' parameter is used, this step is automatically handled.
        # Let's explicitly ensure we are working with the correct suffixed columns
        amount_col_april = 'Sum of Amount in local cur._April'
        amount_col_may = 'Sum of Amount in local cur._May'

        # Fill NaN values with 0 before calculating the difference,
        # so that missing entries are treated as 0 for difference calculation.
        merged_df[amount_col_april] = merged_df[amount_col_april].fillna(0)
        merged_df[amount_col_may] = merged_df[amount_col_may].fillna(0)

        # Calculate the difference
        # The difference is May's amount minus April's amount
        merged_df['Difference in Amount'] = merged_df[amount_col_may] - merged_df[amount_col_april]

        # Select and reorder columns for the output file
        # We'll include columns from April and May and the calculated difference
        output_columns = [
            'Date_April', 'G/L_April', 'DocumentNo', 'Text_April', 'LCurr_April', 'DocType_April', amount_col_april,
            'Date_May', 'G/L_May', 'Text_May', 'LCurr_May', 'DocType_May', amount_col_may,
            'Difference in Amount'
        ]

        # Handle cases where 'DocumentNo' might not be present in both files.
        # The merged_df will have columns from both original dataframes, but some might be NaN.
        # We need to select the most complete set of 'Date', 'G/L', 'Text', 'LCurr', 'DocType' for the output.
        # For simplicity, we can prioritize the April's data for non-amount columns if available,
        # otherwise use May's data.

        final_output_df = pd.DataFrame()
        final_output_df['Date'] = merged_df['Date_April'].fillna(merged_df['Date_May'])
        final_output_df['G/L'] = merged_df['G/L_April'].fillna(merged_df['G/L_May'])
        final_output_df['DocumentNo'] = merged_df['DocumentNo']
        final_output_df['Text'] = merged_df['Text_April'].fillna(merged_df['Text_May'])
        final_output_df['LCurr'] = merged_df['LCurr_April'].fillna(merged_df['LCurr_May'])
        final_output_df['DocType'] = merged_df['DocType_April'].fillna(merged_df['DocType_May'])
        final_output_df['Sum of Amount in local cur. (April)'] = merged_df[amount_col_april]
        final_output_df['Sum of Amount in local cur. (May)'] = merged_df[amount_col_may]
        final_output_df['Difference in Amount'] = merged_df['Difference in Amount']


        # Save the result to a new Excel file
        final_output_df.to_excel(output_file_path, index=False)
        print(f"Difference report saved to '{output_file_path}'")

    except FileNotFoundError:
        print("Error: One or both of the input files were not found. Please check the file paths.")
    except KeyError as e:
        print(f"Error: Missing expected column in one of the Excel files: {e}. Please ensure the column names are exactly as specified.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

# --- How to use the script ---
if __name__ == "__main__":
    # Replace these with the actual paths to your Excel files
    april_file = "poc2_april.xlsx"
    may_file = "poc2_may.xlsx"
    output_file = "amount_difference_report.xlsx"

    calculate_amount_difference(april_file, may_file, output_file)