import streamlit as st
import pandas as pd
from io import BytesIO
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import datetime

st.set_page_config(page_title="IoT Summary Report", layout="centered")
st.title("📊 IoT Excel Summary Report Generator")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=1)

        for col in ['Machine Name', 'Time']:
            if col not in df.columns:
                st.error(f"❌ Your Excel must contain the column: '{col}'")
                st.stop()

        df['Product'] = df.get('Product', '(Unnamed)').fillna("(Unnamed)")
        df['Time'] = pd.to_datetime(df['Time'], errors='coerce')
        df = df.dropna(subset=['Time'])

        machine_groups = df.groupby('Machine Name')
        output = BytesIO()

        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            worksheet = workbook.add_worksheet("Summary Report")
            writer.sheets["Summary Report"] = worksheet

            bold_center = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter'})
            wrap_center = workbook.add_format({'text_wrap': True, 'align': 'center', 'valign': 'vcenter'})
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

                times = machine_df['Time'].tolist()
                for i in range(1, len(times)):
                    gap = (times[i] - times[i-1]).total_seconds() / 60
                    if gap > 5:
                        lag_start = times[i-1]
                        lag_end = times[i]
                        duration = lag_end - lag_start
                        h = int(duration.total_seconds()) // 3600
                        m = (int(duration.total_seconds()) % 3600) // 60
                        s = int(duration.total_seconds()) % 60
                        duration_str = f"{h:02}:{m:02}:{s:02}"
                        worksheet.write(row, 2, lag_start, datetime_format)
                        worksheet.write(row, 3, lag_end, datetime_format)
                        worksheet.write(row, 4, duration_str)
                        row += 1
                row += 2

            for col in range(5):
                worksheet.set_column(col, col, 22)

            # Gantt Chart
            usage_blocks = []
            for machine, group in df.groupby('Machine Name'):
                times = group['Time'].sort_values().tolist()
                if not times:
                    continue
                start_time = times[0]
                for i in range(1, len(times)):
                    if (times[i] - times[i-1]).total_seconds() > 300:
                        usage_blocks.append((machine, start_time, times[i-1]))
                        start_time = times[i]
                usage_blocks.append((machine, start_time, times[-1]))

            # Plot Gantt
            fig, ax = plt.subplots(figsize=(11.7, 8.3))
            machines = sorted(df['Machine Name'].unique())
            machine_index = {m: i for i, m in enumerate(machines)}
            colors = plt.cm.get_cmap("tab10", len(machines))

            for idx, (machine, start, end) in enumerate(usage_blocks):
                y = machine_index[machine]
                ax.barh(y, (end - start).total_seconds()/3600, left=start, height=0.4, color=colors(idx % 10))

            ax.set_yticks(list(machine_index.values()))
            ax.set_yticklabels(machine_index.keys())
            ax.set_xlim(datetime.datetime.combine(df['Time'].min().date(), datetime.time(6, 0)),
                        datetime.datetime.combine(df['Time'].min().date(), datetime.time(23, 59)))
            ax.xaxis.set_major_locator(mdates.HourLocator())
            ax.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
            ax.set_xlabel("Time")
            ax.set_title("Machine Usage Timeline")
            fig.autofmt_xdate()

            # Save and insert chart
            img_data = BytesIO()
            fig.tight_layout()
            plt.savefig(img_data, format='png')
            plt.close(fig)
            img_data.seek(0)

            chart_sheet = workbook.add_worksheet("Machine Usage Timeline")
            chart_sheet.insert_image("B2", "usage_timeline.png", {"image_data": img_data})

        st.success("✅ Report created successfully!")
        st.download_button("📥 Download Summary Report", data=output.getvalue(),
                           file_name="Summary_Report_with_Timeline.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")
