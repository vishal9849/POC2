import pandas as pd
import os
from openai import AzureOpenAI # Use AzureOpenAI for Azure-specific features

# --- Azure OpenAI Configuration ---
# Replace with your actual Azure OpenAI details
# AZURE_OPENAI_API_KEY = "YOUR_AZURE_OPENAI_API_KEY"
# AZURE_OPENAI_ENDPOINT = "YOUR_AZURE_OPENAI_ENDPOINT"
# AZURE_OPENAI_API_VERSION = "2024-02-15"  # Check your deployment for the correct API version
# AZURE_OPENAI_MODEL_NAME = "your-deployment-name" # This is your deployment name, not the model name like 'gpt-4'

AZURE_OPENAI_ENDPOINT="https://anilk-maqg6o7m-eastus2.cognitiveservices.azure.com/"
AZURE_OPENAI_MODEL_NAME="gpt-4"
AZURE_OPENAI_API_KEY="7AOJWoLBYXQjchlS1pmG3EzPjJASQF0xCaBjXthpTSXMBW4GQXPLJQQJ99BEACHYHv6XJ3w3AAAAACOGuXgs"
AZURE_OPENAI_API_VERSION="2024-12-01-preview"  # Use 2025-01-01-preview if errors occurs
AZURE_OPENAI_API_TYPE="azure"

# Initialize the Azure OpenAI client
client = AzureOpenAI(
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_API_VERSION
)

def get_ai_summary(text_context, amount_difference):
    """
    Sends text context and amount difference to Azure OpenAI for summarization.

    Args:
        text_context (str): The concatenated 'Text' and 'DocType' for context.
        amount_difference (float): The difference in amount for the current record.

    Returns:
        str: The summarized text from Azure OpenAI, or an error message if something goes wrong.
    """
    try:
        prompt = f"""
        Summarize the following financial transaction context and its change in amount.
        Context: "{text_context}"
        Amount Change: {amount_difference:.2f}

        Provide a concise summary, highlighting the nature of the transaction and the significance of the amount change (e.g., "significant increase", "minor decrease", "new transaction").
        """

        response = client.chat.completions.create(
            model=AZURE_OPENAI_MODEL_NAME,
            messages=[
                {"role": "system", "content": "You are a financial assistant that summarizes transaction details."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7, # Adjust creativity (0.0 for factual, 1.0 for more creative)
            max_tokens=150 # Limit the length of the summary
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"Error calling Azure OpenAI: {e}")
        return f"Error: Could not get summary - {e}"

def calculate_amount_difference_and_summarize(file_path_april, file_path_may, output_file_path="difference_and_summary_report.xlsx"):
    """
    Compares two Excel files, calculates the difference in 'Sum of Amount in local cur.',
    concatenates 'Text' and 'DocType', uses Azure OpenAI to summarize,
    and saves the result to a new Excel file.

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
        merged_df = pd.merge(df_april, df_may, on='DocumentNo', how='outer', suffixes=('_April', '_May'))

        amount_col_april = 'Sum of Amount in local cur._April'
        amount_col_may = 'Sum of Amount in local cur._May'

        # Fill NaN values with 0 before calculating the difference
        merged_df[amount_col_april] = merged_df[amount_col_april].fillna(0)
        merged_df[amount_col_may] = merged_df[amount_col_may].fillna(0)

        # Calculate the difference
        merged_df['Difference in Amount'] = merged_df[amount_col_may] - merged_df[amount_col_april]

        # Prepare the final output DataFrame with consolidated non-amount columns
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

        # Concatenate 'Text' and 'DocType' for AI context
        # Ensure they are strings before concatenation to avoid errors with NaN or other types
        final_output_df['Concatenated Context'] = final_output_df['Text'].fillna('') + ' | DocType: ' + final_output_df['DocType'].fillna('')

        # Add a column for AI Summary
        final_output_df['AI Summary'] = ''

        print("Generating AI summaries. This might take some time depending on the number of rows and API rate limits...")
        # Iterate through rows to get AI summaries
        for index, row in final_output_df.iterrows():
            text_context = row['Concatenated Context']
            amount_diff = row['Difference in Amount']
            final_output_df.at[index, 'AI Summary'] = get_ai_summary(text_context, amount_diff)
            # Optional: Add a small delay to respect API rate limits
            # import time
            # time.sleep(0.1)

        # Save the result to a new Excel file
        final_output_df.to_excel(output_file_path, index=False)
        print(f"Difference and summarized report saved to '{output_file_path}'")

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
    output_file = "amount_difference_and_summary_report.xlsx"

    # Make sure to set your Azure OpenAI API_KEY, ENDPOINT, and MODEL_NAME above!
    if AZURE_OPENAI_API_KEY == "YOUR_AZURE_OPENAI_API_KEY" or AZURE_OPENAI_ENDPOINT == "YOUR_AZURE_OPENAI_ENDPOINT" or AZURE_OPENAI_MODEL_NAME == "your-deployment-name":
        print("WARNING: Please update AZURE_OPENAI_API_KEY, AZURE_OPENAI_ENDPOINT, and AZURE_OPENAI_MODEL_NAME in the script before running.")
    else:
        calculate_amount_difference_and_summarize(april_file, may_file, output_file)