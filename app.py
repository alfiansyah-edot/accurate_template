import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
import io
from datetime import datetime

st.title("Accurate SO Template Generator")

uploaded_file = st.file_uploader("Upload your Sales Order CSV file", type=["csv", "xlsx"])

if uploaded_file is not None:
    with st.spinner("Processing file, please wait..."):
        # Read file based on extension
        if uploaded_file.name.endswith(".csv"):
            data = pd.read_csv(uploaded_file)
        else:
            data = pd.read_excel(uploaded_file)

        # Filter and convert date
        data = data[data["type"] == "item"]
        data["date"] = pd.to_datetime(data["date"])

        # Ranking logic
        data = data.sort_values(by=["date", "number"])
        data["rank"] = data.drop_duplicates(subset=["date", "number"]).reset_index(drop=True).reset_index().set_index(["date", "number"]).reindex(data.set_index(["date", "number"]).index)["index"].values + 1
        data["product_rank"] = data.groupby("number")["product_id"].rank(method="dense").astype(int)

        data["date"] = data["date"].dt.strftime('%m/%d/%Y')
        data = data.sort_values(by=["rank", "product_rank"])

        # Prepare header, item, expense rows
        header_rows = data[["rank", "date", "customer_id"]].drop_duplicates().copy()
        header_rows["row_or_header"] = 2
        header_rows["rank_part"] = 1
        header_rows["header"] = "HEADER"
        header_rows["no_form"] = ""
        header_rows["tgl_pesanan"] = header_rows["date"]
        header_rows["no_pelanggan"] = header_rows["customer_id"]
        header_rows = header_rows[["row_or_header", "rank", "rank_part", "header", "no_form", "tgl_pesanan", "no_pelanggan"]]

        item_rows = data[["rank", "product_id", "product", "qty"]].copy()
        item_rows["row_or_header"] = 2
        item_rows["rank_part"] = 2
        item_rows["header"] = "ITEM"
        item_rows["no_form"] = item_rows["product_id"]
        item_rows["tgl_pesanan"] = item_rows["product"]
        item_rows["no_pelanggan"] = pd.to_numeric(item_rows["qty"], errors='coerce')
        item_rows = item_rows[["row_or_header", "rank", "rank_part", "header", "no_form", "tgl_pesanan", "no_pelanggan"]]

        expense_rows = data[["rank"]].drop_duplicates().copy()
        expense_rows["row_or_header"] = 2
        expense_rows["rank_part"] = 3
        expense_rows["header"] = "EXPENSE"
        expense_rows["no_form"] = ""
        expense_rows["tgl_pesanan"] = ""
        expense_rows["no_pelanggan"] = ""
        expense_rows = expense_rows[["row_or_header", "rank", "rank_part", "header", "no_form", "tgl_pesanan", "no_pelanggan"]]

        transform_data = pd.concat([header_rows, item_rows, expense_rows], ignore_index=True)
        transform_data = transform_data.sort_values(by=["rank", "rank_part"]).reset_index(drop=True)

        # Load Excel template
        excel_path = "header.xlsx"  # Adjust this path to your template
        wb = load_workbook(excel_path)
        ws = wb.active

        start_row = 4
        for r_idx, row in enumerate(dataframe_to_rows(transform_data.iloc[:, 3:], index=False, header=False), start=start_row):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"final_sales_order_{timestamp}.xlsx"

    st.success("Processing complete!")

    st.download_button(
        label="Download",
        data=output,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
