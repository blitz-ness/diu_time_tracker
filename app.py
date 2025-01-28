import cv2
import pytesseract
import pandas as pd
import numpy as np
import re
from datetime import datetime
import streamlit as st

# Configure Tesseract executable path (if necessary)
# Uncomment and specify path if not in default PATH
#pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' #for local pc
pytesseract.pytesseract.tesseract_cmd = "/usr/bin/tesseract" #for streamlit cloud

# Streamlit app
st.title("DIU Time Tracking 📊")
st.write("Upload your attendance screenshots to process and generate an Excel summary.")

# File uploader
uploaded_files = st.file_uploader("Upload Attendance Screenshots", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    all_punch_times = []

    for uploaded_file in uploaded_files:
        # Load and preprocess image
        file_bytes = uploaded_file.read()
        image = cv2.imdecode(np.frombuffer(file_bytes, np.uint8), cv2.IMREAD_COLOR)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)

        # Extract text using pytesseract
        extracted_text = pytesseract.image_to_string(thresh)

        # Extract punch times
        lines = extracted_text.split("\n")
        punch_times = [
            line.strip()
            for line in lines
            if re.match(r"[A-Za-z]{3} \d{1,2}, \d{4}, \d{1,2}:\d{2}:\d{2} [APap][Mm]", line.strip())
        ]
        all_punch_times.extend(punch_times)

    # Convert to DataFrame
    df = pd.DataFrame(all_punch_times, columns=["Punch Time"])

    # Parse datetime
    df["Punch Time"] = pd.to_datetime(df["Punch Time"], format="%b %d, %Y, %I:%M:%S %p", errors="coerce")

    # Group by date and pair times
    df['Date'] = df['Punch Time'].dt.date
    df = df.sort_values(by=['Date', 'Punch Time'])

    pair_start, pair_end = [], []
    for date, group in df.groupby('Date'):
        group = group.reset_index(drop=True)
        for i in range(0, len(group), 2):
            pair_start.append(group.loc[i, 'Punch Time'])
            pair_end.append(group.loc[i + 1, 'Punch Time'] if i + 1 < len(group) else None)

    df_pairs = pd.DataFrame({
        'Punch In': pair_start,
        'Punch Out': pair_end
    })

    df_pairs['Duration (hrs)'] = (df_pairs['Punch Out'] - df_pairs['Punch In']).dt.total_seconds() / 3600
    df_pairs['Duration (hrs)'].fillna(0, inplace=True)

    # Daily time summary
    daily_time = df_pairs.groupby(df_pairs['Punch In'].dt.date)['Duration (hrs)'].sum().reset_index()
    daily_time.columns = ['Date', 'Total Hours']

    # Monthly average summary
    daily_time['Month'] = pd.to_datetime(daily_time['Date']).dt.to_period('M')
    monthly_avg = daily_time.groupby('Month')['Total Hours'].mean().reset_index()

    # Show results
    st.write("### Paired Punch Times and Durations")
    st.dataframe(df_pairs)

    st.write("### Daily Time Summary")
    st.dataframe(daily_time)

    st.write("### Monthly Average Time")
    st.dataframe(monthly_avg)

    # Download Excel
    output_filename = "attendance_summary.xlsx"
    with pd.ExcelWriter(output_filename) as writer:
        daily_time.to_excel(writer, sheet_name="Daily Time", index=False)
        monthly_avg.to_excel(writer, sheet_name="Monthly Average", index=False)

    with open(output_filename, "rb") as file:
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
