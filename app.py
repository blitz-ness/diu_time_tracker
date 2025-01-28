import streamlit as st
import cv2
import pytesseract
import pandas as pd
import re
from datetime import datetime
from PIL import Image
import io

# Streamlit App Title
st.title("📊 DIU Time Tracker")

# Step 1: Upload Multiple Images
uploaded_files = st.file_uploader("📂 Upload attendance screenshots", type=["png", "jpg", "jpeg"], accept_multiple_files=True)

if uploaded_files:
    # Initialize an empty list to store punch times from all images
    all_punch_times = []

    # Process each uploaded image
    for uploaded_file in uploaded_files:
        st.write(f"✅ Processing file: {uploaded_file.name}")

        # Step 2: Load and Preprocess Image
        image = Image.open(uploaded_file)
        image = cv2.cvtColor(np.array(image), cv2.COLOR_RGB2BGR)  # Convert PIL image to OpenCV format
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)  # Convert to grayscale
        _, thresh = cv2.threshold(gray, 150, 255, cv2.THRESH_BINARY)  # Apply thresholding

        # Step 3: Extract Text using OCR
        extracted_text = pytesseract.image_to_string(thresh)

        # Step 4: Extract Punch Times manually
        lines = extracted_text.split("\n")
        punch_times = []

        for line in lines:
            line = line.strip()
            # Match date-time format (e.g., "Jan 28, 2025, 11:32:48 AM")
            if re.match(r"[A-Za-z]{3} \d{1,2}, \d{4}, \d{1,2}:\d{2}:\d{2} [APap][Mm]", line):
                punch_times.append(line)

        # Add punch times from current image to the main list
        all_punch_times.extend(punch_times)

    # Step 5: Convert to Pandas DataFrame
    df = pd.DataFrame(all_punch_times, columns=["Punch Time"])

    # Step 6: Convert "Punch Time" to datetime format
    try:
        df["Punch Time"] = pd.to_datetime(df["Punch Time"], format="%b %d, %Y, %I:%M:%S %p")
    except Exception as e:
        st.error(f"⛔ Date Parsing Error! Check extracted text format: {e}")

    # Step 7: Pair Punch-In and Punch-Out Times based on the Same Date
    df['Date'] = df['Punch Time'].dt.date  # Extract the date
    df = df.sort_values(by=['Date', 'Punch Time'])  # Sort by Date and Time

    # Initialize lists to store paired times
    pair_start = []
    pair_end = []

    # Loop to pair punch-in and punch-out based on the same date
    for date, group in df.groupby('Date'):
        group = group.reset_index(drop=True)  # Reset index for the group
        for i in range(0, len(group), 2):  # Loop in steps of 2 to pair them
            if i + 1 < len(group):  # If there is a pair
                pair_start.append(group.loc[i, 'Punch Time'])
                pair_end.append(group.loc[i + 1, 'Punch Time'])
            else:
                # If unpaired punch, store only the start time
                pair_start.append(group.loc[i, 'Punch Time'])
                pair_end.append(None)  # No pair for the last punch-in

    # Add the paired times to the DataFrame
    df_pairs = pd.DataFrame({
        'Punch In': pair_start,
        'Punch Out': pair_end
    })

    # Step 8: Calculate Duration (Time Difference) for Each Pair
    df_pairs['Duration (hrs)'] = (df_pairs['Punch Out'] - df_pairs['Punch In']).dt.total_seconds() / 3600
    df_pairs['Duration (hrs)'].fillna(0, inplace=True)  # For unpaired punches, set duration to 0

    # Step 9: Calculate Daily Time (Total time spent per day)
    daily_time = df_pairs.groupby(df_pairs['Punch In'].dt.date)['Duration (hrs)'].sum().reset_index()

    # Step 10: Calculate Monthly Average Punch-in Time
    daily_time['Month'] = pd.to_datetime(daily_time['Punch In']).dt.to_period('M')  # Extract month
    monthly_avg = daily_time.groupby('Month')['Duration (hrs)'].mean().reset_index()

    # Step 11: Save the results to Excel
    output_filename = "attendance_summary.xlsx"
    with pd.ExcelWriter(output_filename) as writer:
        daily_time.to_excel(writer, sheet_name='Daily Time', index=False)
        monthly_avg.to_excel(writer, sheet_name='Monthly Average', index=False)

    # Step 12: Provide Download Link
    st.success("✅ Processing complete!")
    with open(output_filename, "rb") as file:
        st.download_button(
            label="📥 Download Excel File",
            data=file,
            file_name=output_filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # Debugging: Display DataFrames
    st.write("📊 Paired Punch Times and Durations:")
    st.dataframe(df_pairs)

    st.write("📅 Daily Time (Total time per day in hours):")
    st.dataframe(daily_time)

    st.write("📈 Monthly Average Punch-in Time (in hours):")
    st.dataframe(monthly_avg)