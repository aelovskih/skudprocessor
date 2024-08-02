import streamlit as st
import pandas as pd
import io

def process_hr_report(file):
    # Load the provided Excel file
    df = pd.read_excel(file)

    # Drop any summary rows if present
    df = df[~df['Фамилия'].str.contains('Итого:', na=False)]

    # Initialize the output dataframe
    days = list(range(1, 32))
    output_columns = ['Фамилия', 'Имя'] + [str(day) for day in days]
    output_df = pd.DataFrame(columns=output_columns)

    # Group by employee
    grouped = df.groupby(['Фамилия', 'Имя'])

    rows = []  # List to collect rows

    for (last_name, first_name), group in grouped:
        # Create a row for the current employee
        row = {'Фамилия': last_name, 'Имя': first_name}

        for _, entry in group.iterrows():
            day = str(pd.to_datetime(entry['Дата']).day)
            if pd.notna(entry['Вход']):
                time_in = pd.to_datetime(entry['Вход']).strftime('%H:%M')
            else:
                time_in = 'x'
            if pd.notna(entry['Выход']):
                time_out = pd.to_datetime(entry['Выход']).strftime('%H:%M')
            else:
                time_out = '!'
            row[day] = f"{time_in}-{time_out}" if time_out != '!' else f"{time_in}-!"

        rows.append(row)  # Add the row to the list

    # Convert list of rows to DataFrame
    output_df = pd.DataFrame(rows, columns=output_columns)

    return output_df

st.title("HR Report Processor")

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    processed_report = process_hr_report(uploaded_file)
    st.write("Processed Report")
    st.dataframe(processed_report)

    # Provide download link
    def to_excel(df):
        output = io.BytesIO()
        writer = pd.ExcelWriter(output, engine='openpyxl')
        df.to_excel(writer, index=False, sheet_name='Sheet1')
        writer.close()
        processed_data = output.getvalue()
        return processed_data

    df_xlsx = to_excel(processed_report)
    st.download_button(label='📥 Download Processed Report',
                       data=df_xlsx,
                       file_name='processed_hr_report.xlsx')

