import pandas as pd

# Read the existing Excel file
df = pd.read_excel('output2.xlsx')

# Loop through each row and update the summary_text and summary_gpt columns
for index, row in df.iterrows():
    text_summary = row['Text Summary:']
    chatgpt_summary = row['ChatGPT Summary:']
    
    df.at[index, 'summary_text'] = text_summary
    df.at[index, 'summary_gpt'] = chatgpt_summary

# Write the updated DataFrame to a new Excel file
df.to_excel('updated_convert_file.xlsx', index=False)

