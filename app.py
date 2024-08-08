import streamlit as st
import pandas as pd
import io
from datetime import datetime, timedelta

def calculate_time_difference(time_in, time_out):
    time_in_dt = datetime.strptime(time_in, '%H:%M')
    time_out_dt = datetime.strptime(time_out, '%H:%M')
    return time_out_dt - time_in_dt

def format_timedelta(td):
    total_minutes = int(td.total_seconds() // 60)
    hours, minutes = divmod(total_minutes, 60)
    return f'{hours:02}:{minutes:02}'

def process_hr_report(file):
    # Load the provided Excel file, skipping the first three rows
    df = pd.read_excel(file, skiprows=3)

    # Drop any summary rows if present
    df = df[~df['Фамилия'].str.contains('Итого:', na=False)]

    # Extract unique dates from the 'Дата' column and format them as needed
    unique_dates = pd.to_datetime(df['Дата']).dt.date.unique()
    unique_dates.sort()

    # Format dates as 'dd.mm' after sorting
    formatted_dates = [date.strftime('%d.%m') for date in unique_dates]

    # Initialize the output dataframe
    output_columns = ['Фамилия', 'Имя', 'Должность'] + formatted_dates + ['Среднее время присутствия', 'Общее время присутствия']
    output_df = pd.DataFrame(columns=output_columns)

    # Group by employee
    grouped = df.groupby(['Фамилия', 'Имя', 'Должность'])

    rows = []  # List to collect rows

    for (last_name, first_name, position), group in grouped:
        # Create a row for the current employee
        row = {'Фамилия': last_name, 'Имя': first_name, 'Должность': position}

        total_time = timedelta()
        valid_entries = 0

        for _, entry in group.iterrows():
            date = pd.to_datetime(entry['Дата']).strftime('%d.%m')
            if pd.notna(entry['Вход']):
                time_in = pd.to_datetime(entry['Вход']).strftime('%H:%M')
            else:
                time_in = 'x'
            if pd.notna(entry['Выход']):
                time_out = pd.to_datetime(entry['Выход']).strftime('%H:%M')
            else:
                time_out = '!'

            if time_in != 'x' and time_out != '!':
                time_diff = calculate_time_difference(time_in, time_out)
                total_time += time_diff
                valid_entries += 1
                row[date] = f"{time_in}-{time_out}"
            else:
                row[date] = f"{time_in}-{time_out}"

        if valid_entries > 0:
            average_time = total_time / valid_entries
            if valid_entries == len(group):
                row['Среднее время присутствия'] = format_timedelta(average_time)
                row['Общее время присутствия'] = format_timedelta(total_time)
            else:
                row['Среднее время присутствия'] = f"~{format_timedelta(average_time)}"
                row['Общее время присутствия'] = f"~{format_timedelta(total_time)}"
        else:
            row['Среднее время присутствия'] = "Невозможно определить"
            row['Общее время присутствия'] = "Невозможно определить"

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
