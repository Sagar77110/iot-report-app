import streamlit as st
import pandas as pd
from io import BytesIO

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

                # Merged header row for Lag Periods
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

                        # Convert duration to HH:MM:SS string without days
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

            # Auto-fit columns A–E
            for col_num in range(5):
                worksheet.set_column(col_num, col_num, 22)

        st.success("✅ Summary report generated!")
        st.download_button("📥 Download Summary Report", data=output.getvalue(),
                           file_name="Summary_Report.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    except Exception as e:
        st.error(f"⚠️ Error processing file: {e}")

