from io import BytesIO
import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta

st.set_page_config(page_title="IoT Summary Report", layout="centered")
st.title("📊 IoT Excel Summary Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=1)

        required_columns = ['Machine Name', 'Product', 'Time']
        for col in ['Machine Name', 'Time']:
            if col not in df.columns:
                st.error(f"❌ Your Excel must contain the column: '{col}'")
                st.stop()

        df['Product'] = df['Product'].fillna("(Unnamed)")
        df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
        df = df.dropna(subset=['Time'])

        if df.empty:
            st.warning("The uploaded file contains no valid time data after processing. Please check your 'Time' column format.")
            st.stop()

        machine_groups = df.groupby('Machine Name')
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Summary Report")
            writer.sheets["Summary Report"] = worksheet

            bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
            wrap_center = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
            time_format = workbook.add_format({'num_format': 'hh:mm:ss', 'align': 'center'})
            datetime_format = workbook.add_format({'num_format': 'yyyy-mm-dd hh:mm:ss', 'align': 'center'})

            row = 0
            for machine, machine_df in machine_groups:
                machine_df = machine_df.sort_values('Time')

                product_groups = machine_df.groupby('Product')
                worksheet.write_row(row, 0, ["Machine Name", "Product", "Count", "Start Time", "End Time"], bold_center)
                row += 1

                for product, product_df in product_groups:
                    count = len(product_df)
                    start_time = product_df['Time'].min()
                    end_time = product_df['Time'].max()
                    worksheet.write(row, 0, machine)
                    worksheet.write(row, 1, product)
                    worksheet.write(row, 2, count)
                    worksheet.write(row, 3, start_time, datetime_format)
                    worksheet.write(row, 4, end_time, datetime_format)
                    row += 1

                row += 1

                worksheet.merge_range(row, 2, row, 4, f"{machine} – Lag Periods", wrap_center)
                row += 1
                worksheet.write_row(row, 2, ["Lag Start", "Lag End", "Production Loss"], bold_center)
                row += 1

                times = machine_df['Time'].sort_values().tolist()
                for i in range(1, len(times)):
                    gap_minutes = (times[i] - times[i - 1]).total_seconds() / 60
                    if gap_minutes > 5:
                        lag_start = times[i - 1]
                        lag_end = times[i]
                        duration = times[i] - times[i - 1]

                        total_seconds = int(duration.total_seconds())
                        hours = total_seconds // 3600
                        minutes = (total_seconds % 3600) // 60
                        seconds = total_seconds % 60
                        duration_str = f"{hours:02}:{minutes:02}:{seconds:02}"

                        worksheet.write(row, 2, lag_start, datetime_format)
                        worksheet.write(row, 3, lag_end, datetime_format)
                        worksheet.write(row, 4, duration_str)
                        row += 1

                row += 2

            for col_num in range(5):
                worksheet.set_column(col_num, col_num, 22)

            # Create Gantt chart
            # Increase the width of the figure significantly
            # Original: figsize=(12, len(machine_groups) * 0.6)
            # New width (e.g., 20 or 25, adjust as needed)
            fig, ax = plt.subplots(figsize=(60, len(machine_groups) * 0.6)) # <-- Increased width here

            cmap = plt.get_cmap("tab10")
            machines_list = sorted(machine_groups.groups.keys()) # Ensure consistent order
            machine_colors = {name: cmap(i % 10) for i, name in enumerate(machines_list)}

            y_labels = []
            for idx, machine in enumerate(machines_list): # Iterate through sorted machines
                machine_df = df[(df['Machine Name'] == machine)].sort_values("Time")
                
                # Skip if no data for the machine after filtering
                if machine_df.empty:
                    continue

                blocks = []
                start = machine_df.iloc[0]['Time']
                prev = start

                for current in machine_df['Time'].iloc[1:]:
                    if (current - prev).total_seconds() / 60 > 5:
                        blocks.append((start, prev))
                        start = current
                    prev = current
                blocks.append((start, prev))

                y_labels.append(machine) # Add label only if machine has data
                for block_start, block_end in blocks:
                    ax.barh(
                        y=idx,
                        width=block_end - block_start,
                        left=block_start,
                        height=0.4,
                        color=machine_colors[machine],
                        edgecolor='black'
                    )

            ax.set_yticks(range(len(y_labels)))
            ax.set_yticklabels(y_labels)

            # Ensure x-axis ticks are formatted properly and are not too dense
            # Consider adjusting interval if 1 hour is still too crowded
            ax.xaxis.set_major_locator(mdates.HourLocator(interval=1)) # Keep 1-hour interval for now
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            
            # Dynamic X-axis limits based on overall data range, potentially padding
            if not df.empty:
                min_time_data = df['Time'].min()
                max_time_data = df['Time'].max()
                
                # To ensure the axis covers the full relevant day(s) and avoid cutting off data
                # Let's target the actual first and last data points for the limits
                # If you want a fixed 6 AM to 11:59 PM range for the *day of the first entry*, use the commented out line below
                # ax.set_xlim(datetime.combine(min_time_data.date(), datetime.strptime("06:00", "%H:%M").time()),
                #             datetime.combine(min_time_data.date(), datetime.strptime("23:59", "%H:%M").time()))

                # More dynamic: Use actual min/max time with padding
                x_start_limit = min_time_data.replace(hour=6, minute=0, second=0, microsecond=0) # Start at 6 AM of the first day
                x_end_limit = max_time_data.replace(hour=23, minute=59, second=0, microsecond=0) # End at 11:59 PM of the last day

                # If the data spans multiple days, this will show the full range.
                # If only one day, it will show from 6AM to 11:59 PM for that day, encompassing all data.
                ax.set_xlim(x_start_limit, x_end_limit)


            ax.set_xlabel("Time")
            ax.set_title("Machine Usage Timeline")
            ax.grid(True, axis='x', linestyle='--', color='gray')
            ax.grid(False, axis='y') # Keep y-grid off

            # Adjust x-axis tick label rotation for readability if they overlap
            fig.autofmt_xdate()

            chart_image = BytesIO()
            # Increase DPI for sharper image, especially for wider charts
            plt.tight_layout()
            plt.savefig(chart_image, format='png', dpi=300) # <-- Increased DPI here
            plt.close()

            # The 'x_scale' and 'y_scale' in insert_image control how large the image appears in Excel.
            # If the image is too big, you might need to adjust these or Excel column/row sizes.
            worksheet.insert_image("F1", "timeline.png", {"image_data": chart_image, "x_scale": 0.8, "y_scale": 0.8}) # Adjusted scale for Excel

        st.success("✅ Summary report generated!")
        st.download_button("📥 Download Summary Report", data=output.getvalue(),
                           file_name="Summary_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
