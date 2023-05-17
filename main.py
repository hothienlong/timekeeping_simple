import pandas as pd
from datetime import datetime, timedelta

def export_report(file_input, file_output):
  df = pd.read_excel(file_input)

  # Convert 'Date/Time' column to datetime
  df['Date/Time'] = pd.to_datetime(df['Date/Time'])

  # Extract 'No.', 'Name', and 'Date' columns
  df['Date'] = df['Date/Time'].dt.date
  df['Time'] = df['Date/Time'].dt.time

  # Calculate the actual hours worked
  df['Clock-in'] = df.groupby(['No.', 'Date'])['Time'].transform('min')
  df['Clock-out'] = df.groupby(['No.', 'Date'])['Time'].transform('max')

  # Convert Clock-in and Clock-out to datetime objects
  df['Clock-in'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Clock-in'].astype(str))
  df['Clock-out'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Clock-out'].astype(str))

  # Calculate the hours worked
  df['Hours'] = (df['Clock-out'] - df['Clock-in']).dt.total_seconds() / 3600

  # Round the Hours column to 2 decimal places
  df['Hours'] = df['Hours'].round(2)

  # Drop duplicate rows based on 'No.' and 'Date' columns
  df = df.drop_duplicates(subset=['No.', 'Date'])
  
  # Pivot the dataframe to have one row for each name and one column for each date
  pivot_df = df.pivot_table(index=['No.', 'Name'], columns='Date', values='Hours', aggfunc='sum').reset_index()

  # Rename the columns to match the desired format
  pivot_df.columns.name = None

  # Save the pivot dataframe to a new Excel file
  pivot_df.to_excel(file_output, index=False)

  # df = pd.read_excel(file_input)

  # # Convert 'Date/Time' column to datetime
  # df['Date/Time'] = pd.to_datetime(df['Date/Time'])

  # # Extract 'No.' and 'Date' columns
  # df['Date'] = df['Date/Time'].dt.date
  # df['Time'] = df['Date/Time'].dt.time

  # # Calculate the actual hours worked
  # df['Clock-in'] = df.groupby(['No.', 'Date'])['Time'].transform('min')
  # df['Clock-out'] = df.groupby(['No.', 'Date'])['Time'].transform('max')

  # print(df['Clock-in'])
  # print("========")
  # print(df['Clock-out'])
  
  # # Convert Clock-in and Clock-out to datetime objects
  # df['Clock-in'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Clock-in'].astype(str))
  # df['Clock-out'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Clock-out'].astype(str))

  # # Calculate the hours worked
  # df['Hours'] = (df['Clock-out'] - df['Clock-in']).dt.total_seconds() / 3600

  # df = df.drop_duplicates(subset=['No.', 'Date'])

  # # Group by 'Name', 'No.', and 'Date' and sum the hours worked
  # summary = df.groupby(['Name', 'No.', 'Date'])['Hours'].sum().reset_index()

  # # Round the 'Hours' column to 2 decimal places
  # summary['Hours'] = summary['Hours'].round(2)

  # # Sort by 'No.' column
  # summary = summary.sort_values('No.')
  
  # print(summary)
  # # Xuất kết quả ra tệp Excel mới
  # summary.to_excel(file_output, index=False)

if __name__ == "__main__":
  export_report("input.xlsx", "report.xlsx")
