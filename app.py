import streamlit as st
import pandas as pd
import re

# Function to process the Excel file
def process_excel(file):
    xls = pd.ExcelFile(file)
    all_rows = []

    # Read the first sheet by name
    sheet_name = xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=None)

    department_name = None
    employee_id = None
    employee_name = None
    status_columns = []
    status_indices = []

    for index, row in df.iterrows():
        if row[0] == 'Department:':
            department_name = row[4]
            continue

        if row[0] == 'Employee:':
            ot_hours = None
            employee_info = row[4]
            if ':' in employee_info:
                employee_id, employee_name = employee_info.split(':', 1)
                employee_id = employee_id.strip()
                employee_name = employee_name.strip()
            ot_cell_value = row[8]
            ot_match = re.search(r'Total OT: (\d{2}:\d{2}) Hrs', str(ot_cell_value))
            ot_hours = ot_match.group(1) if ot_match else '00:00'
            continue

        if row[0] == 'Days':
            status_columns = [col for col in row[1:] if pd.notna(col)]
            status_indices = [i for i, col in enumerate(row[1:], start=1) if pd.notna(col)]
            continue

        if department_name and employee_id and employee_name and status_columns:
            statuses = [row[i] for i in status_indices]
            present_days = sum(status == 'P' for status in statuses)
            absent_days = sum(status == 'A' for status in statuses)
            total_days = present_days + absent_days

            all_rows.append([
                department_name, employee_id, employee_name, *statuses, present_days, absent_days, ot_hours, total_days
            ])

            employee_id = None
            employee_name = None

    columns = ['Department', 'Employee ID', 'Employee Name'] + status_columns + ['Present Days', 'Absent Days', 'OT Hours', 'Total Days']
    structured_df = pd.DataFrame(all_rows, columns=columns)

    return structured_df

# Streamlit app
def main():
    st.title("Excel Transformer")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        df = process_excel(uploaded_file)
        st.write("Transformed DataFrame:")
        st.dataframe(df)
        
        # Allow users to download the DataFrame
        csv = df.to_csv(index=False)
        st.download_button("Download CSV", csv, "output_data.csv", "text/csv")

if __name__ == "__main__":
    main()
