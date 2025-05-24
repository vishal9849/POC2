import streamlit as st
import pandas as pd
import io
from openai import AzureOpenAI

# --- Azure OpenAI Configuration ---
# AZURE_OPENAI_API_KEY = "YOUR_AZURE_OPENAI_API_KEY"
# AZURE_OPENAI_ENDPOINT = "YOUR_AZURE_OPENAI_ENDPOINT"
# AZURE_OPENAI_API_VERSION = "2024-02-15"
# AZURE_OPENAI_MODEL_NAME = "your-deployment-name"


client = AzureOpenAI(
    azure_endpoint=AZURE_OPENAI_ENDPOINT,
    api_key=AZURE_OPENAI_API_KEY,
    api_version=AZURE_OPENAI_API_VERSION
)

def get_ai_summary(text_context, amount_difference):
    try:
        prompt = f"""
        Summarize the following financial transaction context and its change in amount.
        Context: "{text_context}"
        Amount Change: {amount_difference:.2f}
        Provide a concise summary, highlighting the nature of the transaction and the significance of the amount change.
        """
        response = client.chat.completions.create(
            model=AZURE_OPENAI_MODEL_NAME,
            messages=[
                {"role": "system", "content": "You are a financial assistant that summarizes transaction details."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.7,
            max_tokens=150
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return f"Error: {e}"

def process_files(april_file, may_file):
    df_april = pd.read_excel(april_file)
    df_may = pd.read_excel(may_file)

    df_april['DocumentNo'] = df_april['DocumentNo'].astype(str)
    df_may['DocumentNo'] = df_may['DocumentNo'].astype(str)

    merged_df = pd.merge(df_april, df_may, on='DocumentNo', how='outer', suffixes=('_April', '_May'))

    amount_col_april = 'Sum of Amount in local cur._April'
    amount_col_may = 'Sum of Amount in local cur._May'

    merged_df[amount_col_april] = merged_df[amount_col_april].fillna(0)
    merged_df[amount_col_may] = merged_df[amount_col_may].fillna(0)
    merged_df['Difference in Amount'] = merged_df[amount_col_may] - merged_df[amount_col_april]

    final_df = pd.DataFrame()
    final_df['Date'] = merged_df['Date_April'].fillna(merged_df['Date_May'])
    final_df['G/L'] = merged_df['G/L_April'].fillna(merged_df['G/L_May'])
    final_df['DocumentNo'] = merged_df['DocumentNo']
    final_df['Text'] = merged_df['Text_April'].fillna(merged_df['Text_May'])
    final_df['LCurr'] = merged_df['LCurr_April'].fillna(merged_df['LCurr_May'])
    final_df['DocType'] = merged_df['DocType_April'].fillna(merged_df['DocType_May'])
    final_df['Sum of Amount in local cur. (April)'] = merged_df[amount_col_april]
    final_df['Sum of Amount in local cur. (May)'] = merged_df[amount_col_may]
    final_df['Difference in Amount'] = merged_df['Difference in Amount']
    final_df['Concatenated Context'] = final_df['Text'].fillna('') + ' | DocType: ' + final_df['DocType'].fillna('')

    summaries = []
    with st.spinner("Generating AI summaries..."):
        for _, row in final_df.iterrows():
            summary = get_ai_summary(row['Concatenated Context'], row['Difference in Amount'])
            summaries.append(summary)

    final_df['AI Summary'] = summaries
    output = io.BytesIO()
    final_df.to_excel(output, index=False)
    output.seek(0)
    return output

# --- Streamlit UI ---
st.set_page_config(page_title="Financial Difference Analyzer", layout="centered")
st.title("📊 Financial Transaction Difference Analyzer with AI Summary")

with st.form("upload_form"):
    st.subheader("Upload April and May Excel Files")
    april_file = st.file_uploader("Upload April file", type=["xlsx"], key="april")
    may_file = st.file_uploader("Upload May file", type=["xlsx"], key="may")
    submitted = st.form_submit_button("Compare and Summarize")

if submitted:
    if not (april_file and may_file):
        st.warning("Please upload both files to continue.")
    else:
        output_excel = process_files(april_file, may_file)
        st.success("✅ Processing complete. Download the result below:")
        st.download_button(
            label="📥 Download Summary Excel",
            data=output_excel,
            file_name="amount_difference_and_summary_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
