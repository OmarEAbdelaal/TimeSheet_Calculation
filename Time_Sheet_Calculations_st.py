import streamlit as st
import pandas as pd
import numpy as np
import os
from openpyxl import load_workbook
from io import BytesIO

# Streamlit UI
st.title("Clockify Attendance Calculator ")

# File uploaders
csv_file = st.file_uploader("Upload CSV File", type=["csv"])
current_dir = os.path.dirname(os.path.abspath(__file__))
attendance_file = os.path.join(current_dir, "Attendance.xlsx")

if csv_file:
    try:
        # Load CSV data
        time_sheet = pd.read_csv(csv_file)
        
        # Process columns
        columns = ['User', 'Start Date', 'Start Time', 'End Time', 'Duration (h)', 'Task']
        time_data = time_sheet[columns]
        time_data['Start Date'] = pd.to_datetime(time_data['Start Date'])
        time_data['Duration (h)'] = pd.to_timedelta(time_data['Duration (h)'])
        time_data['Start Time'] = pd.to_datetime(time_data['Start Time']).dt.time
        time_data['End Time'] = pd.to_datetime(time_data['End Time']).dt.time

        # Process 'Task' column
        time_data['Task'] = time_data['Task'].fillna('')
        filtered_data = time_data[time_data['Task'].str.contains('LEAVE')]
        time_data.loc[~time_data['Task'].str.contains('LEAVE'), 'Task'] = np.nan
        final_data = pd.concat([filtered_data, time_data[time_data['Task'].notna()]]).reset_index(drop=True)

        # Group data
        grouped_time_data = time_data.groupby(['Start Date', 'User']).agg({
            'Start Time': 'min',
            'End Time': 'max',
            'Duration (h)': 'sum',
            'Task': 'sum'
        }).reset_index()

        unique_users = grouped_time_data['User'].unique()

        # Read the uploaded Excel file
        attendance_bytes = BytesIO(attendance_file)
        destination_workbook = load_workbook(attendance_bytes)

        for user in unique_users:
            user_time_data = grouped_time_data[grouped_time_data['User'] == user]
            date_range = pd.date_range(start=user_time_data['Start Date'].min(), end=user_time_data['Start Date'].max(), freq='D')
            user_date_range_df = pd.DataFrame({'Start Date': date_range})

            merged_df = pd.merge(user_date_range_df, user_time_data, on='Start Date', how='left')
            merged_df['User'].fillna(user, inplace=True)
            merged_df.drop(columns=['User'], inplace=True)
            merged_df.replace(0, '', inplace=True)

            # If sheet exists in destination workbook, update it
            if user in destination_workbook.sheetnames:
                destination_sheet = destination_workbook[user]

                for i, row in enumerate(merged_df.itertuples(index=False), start=2):
                    for j, value in enumerate(row, start=1):
                        destination_sheet.cell(row=i, column=j, value=value)

        # Save the updated Excel file
        output_stream = BytesIO()
        destination_workbook.save(output_stream)
        output_stream.seek(0)

        st.success("Processing complete! Click below to download the updated Excel file.")

        # Provide download button
        st.download_button(label="Download Updated Excel File",
                           data=output_stream,
                           file_name="Updated_Attendance.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"An error occurred: {str(e)}")
