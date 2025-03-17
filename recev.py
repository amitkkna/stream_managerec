import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
from datetime import date, datetime

# --------------------------------------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------------------------------------
st.set_page_config(page_title="Receivables Dashboard", layout="wide")

# --------------------------------------------------------------------------------
# LOAD DATA (Two Sheets: Invoices + Payments)
# --------------------------------------------------------------------------------
@st.cache_data
def load_data(excel_file: str):
    """
    Reads 'Invoices' & 'Payments' sheets from Excel.
    """
    df_invoices = pd.read_excel(
        excel_file,
        sheet_name="Invoices",
        parse_dates=["Invoice Date", "Due Date"]
    )
    df_payments = pd.read_excel(
        excel_file,
        sheet_name="Payments",
        parse_dates=["Payment Date"]
    )
    return df_invoices, df_payments

# Adjust path to your local file
EXCEL_FILE_PATH = r"exceldata/SVP Sample data.xlsx"
df_invoices, df_payments = load_data(EXCEL_FILE_PATH)

# For global date filters
min_date = df_invoices["Invoice Date"].min().date()
max_date = df_invoices["Invoice Date"].max().date()

# --------------------------------------------------------------------------------
# HELPER FUNCTIONS FOR RECEIVABLES, BANKER, LEDGER
# --------------------------------------------------------------------------------

def create_receivables_report(df_invoices, df_payments, from_date, to_date, group_by):
    """
    Generates Receivables Report with partial payments properly distributed
    across Machine, Parts, and Service lines, ensuring Total OS matches
    the sum of Machine OS + Parts OS + Service OS.
    """
    today_date = date.today()

    # 1) Sum partial payments up to 'today'
    df_pay_lim = df_payments[df_payments["Payment Date"].dt.date <= today_date].copy()
    paid_agg = df_pay_lim.groupby("Invoice ID")["Payment Amount"].sum().rename("PaidToDate")

    merged = df_invoices.merge(paid_agg, on="Invoice ID", how="left")
    merged["PaidToDate"] = merged["PaidToDate"].fillna(0.0)
    merged["Outstanding"] = merged["Total Amount"] - merged["PaidToDate"]

    # 2) Filter by invoice date
    filtered_df = merged[
        (merged["Invoice Date"].dt.date >= from_date)
        & (merged["Invoice Date"].dt.date <= to_date)
    ].copy()

    # 3) Compute Days Past Due & Aging
    filtered_df["Days Past Due"] = (
        pd.to_datetime(today_date) - filtered_df["Due Date"]
    ).dt.days.fillna(0)

    def aging_bucket(days):
        if days <= 0:
            return "Current"
        elif days <= 30:
            return "1-30 Days"
        elif days <= 60:
            return "31-60 Days"
        elif days <= 90:
            return "61-90 Days"
        else:
            return "90+ Days"

    filtered_df["Aging Bucket"] = filtered_df["Days Past Due"].apply(aging_bucket)

    # 4) Distribute partial payments proportionally to Machine, Parts, Service
    def distribute_partial_payments(row):
        total_inv = row["Total Amount"]
        paid = row["PaidToDate"]
        if total_inv <= 0:
            # Avoid divide-by-zero or negative totals
            return row["Machine Revenue"], row["Parts Revenue"], row["Service Revenue"]
        
        # Proportions of each line item
        mach_share = row["Machine Revenue"] / total_inv
        parts_share = row["Parts Revenue"] / total_inv
        serv_share = row["Service Revenue"] / total_inv
        
        # Payment allocated to each line
        mach_paid = paid * mach_share
        parts_paid = paid * parts_share
        serv_paid = paid * serv_share
        
        # Outstanding for each line
        mach_os = row["Machine Revenue"] - mach_paid
        parts_os = row["Parts Revenue"] - parts_paid
        serv_os = row["Service Revenue"] - serv_paid
        return mach_os, parts_os, serv_os

    filtered_df["Machine OS"], filtered_df["Parts OS"], filtered_df["Service OS"] = zip(
        *filtered_df.apply(distribute_partial_payments, axis=1)
    )

    # 5) Determine grouping column
    if group_by == "Grand Total":
        group_col = None
    elif group_by == "Branch Wise Details":
        group_col = "Branch"  # Ensure your Invoices data has a "Branch" column
    else:
        group_col = group_by

    # 6) Summation for the 'Total OS'
    if group_col:
        df_total = filtered_df.groupby(group_col)["Outstanding"].sum().rename("Total OS")
    else:
        df_total = pd.Series(
            filtered_df["Outstanding"].sum(),
            index=["Grand Total"],
            name="Total OS"
        )

    # 7) Aging pivot
    if group_col:
        aging_pivot = filtered_df.pivot_table(
            index=group_col,
            columns="Aging Bucket",
            values="Outstanding",
            aggfunc="sum",
            fill_value=0
        )
    else:
        pivot_series = filtered_df.groupby("Aging Bucket")["Outstanding"].sum()
        aging_pivot = pd.DataFrame([pivot_series], index=["Grand Total"])
        aging_pivot.fillna(0, inplace=True)

    for bucket in ["Current", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]:
        if bucket not in aging_pivot.columns:
            aging_pivot[bucket] = 0

    # 8) Summation for the line-item OS
    if group_col:
        df_machine = filtered_df.groupby(group_col)["Machine OS"].sum().rename("Machine OS")
        df_parts   = filtered_df.groupby(group_col)["Parts OS"].sum().rename("Parts OS")
        df_service = filtered_df.groupby(group_col)["Service OS"].sum().rename("Service OS")
    else:
        df_machine = pd.Series(
            filtered_df["Machine OS"].sum(),
            index=["Grand Total"],
            name="Machine OS"
        )
        df_parts = pd.Series(
            filtered_df["Parts OS"].sum(),
            index=["Grand Total"],
            name="Parts OS"
        )
        df_service = pd.Series(
            filtered_df["Service OS"].sum(),
            index=["Grand Total"],
            name="Service OS"
        )

    # 9) Combine everything into final_df
    final_df = pd.DataFrame(df_total).join(aging_pivot, how="left")
    final_df = final_df.join(df_machine, how="left")
    final_df = final_df.join(df_parts, how="left")
    final_df = final_df.join(df_service, how="left")
    final_df.reset_index(inplace=True)

    # Rename grouping column to 'Group'
    if group_col is None:
        final_df.rename(columns={"index": "Group"}, inplace=True)
    else:
        # For Branch grouping, rename the 'Branch' column to 'Group'
        col_name = "Branch" if group_by == "Branch Wise Details" else group_by
        final_df.rename(columns={col_name: "Group"}, inplace=True)

    # 10) Final column order
    col_order = [
        "Group",
        "Total OS",
        "Current",
        "1-30 Days",
        "31-60 Days",
        "61-90 Days",
        "90+ Days",
        "Machine OS",
        "Parts OS",
        "Service OS",
    ]
    final_df = final_df[col_order]

    return final_df


def create_banker_report(df_invoices, df_payments, from_date, to_date):
    """
    Banker report using partial payments as credits.
    """
    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Customer Name"]],
        on="Invoice ID",
        how="left"
    )

    df_inv_before = df_invoices[df_invoices["Invoice Date"].dt.date < from_date]
    df_pay_before = dfp_merged[dfp_merged["Payment Date"].dt.date < from_date]

    df_inv_range = df_invoices[
        (df_invoices["Invoice Date"].dt.date >= from_date)
        & (df_invoices["Invoice Date"].dt.date <= to_date)
    ]
    df_pay_range = dfp_merged[
        (dfp_merged["Payment Date"].dt.date >= from_date)
        & (dfp_merged["Payment Date"].dt.date <= to_date)
    ]

    inv_before_agg = df_inv_before.groupby("Customer Name")["Total Amount"].sum().rename("InvBefore")
    pay_before_agg = df_pay_before.groupby("Customer Name")["Payment Amount"].sum().rename("PayBefore")

    inv_range_agg = df_inv_range.groupby("Customer Name")["Total Amount"].sum().rename("InvRange")
    pay_range_agg = df_pay_range.groupby("Customer Name")["Payment Amount"].sum().rename("PayRange")

    all_cust = set(inv_before_agg.index).union(
        pay_before_agg.index, inv_range_agg.index, pay_range_agg.index
    )
    all_cust = sorted(all_cust)

    banker_df = pd.DataFrame(index=all_cust)
    banker_df["Opening (Invoices)"] = inv_before_agg
    banker_df["Opening (Payments)"] = pay_before_agg
    banker_df["Debits (Invoices)"]  = inv_range_agg
    banker_df["Credits (Payments)"] = pay_range_agg
    banker_df.fillna(0.0, inplace=True)

    banker_df["Opening Balance"] = banker_df["Opening (Invoices)"] - banker_df["Opening (Payments)"]
    banker_df["Balance"] = banker_df["Opening Balance"] + banker_df["Debits (Invoices)"] - banker_df["Credits (Payments)"]

    banker_df.reset_index(inplace=True)
    banker_df.rename(columns={"index": "Customer Name"}, inplace=True)

    col_order = [
        "Customer Name",
        "Opening Balance",
        "Debits (Invoices)",
        "Credits (Payments)",
        "Balance"
    ]
    banker_df = banker_df[col_order]
    return banker_df


def create_customer_ledger(df_invoices, df_payments, from_date, to_date, customer_name):
    """
    True ledger with invoice lines as Debits & payment lines as Credits.
    """
    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Customer Name"]],
        on="Invoice ID",
        how="left"
    )

    df_inv = df_invoices[
        (df_invoices["Customer Name"] == customer_name)
        & (df_invoices["Invoice Date"].dt.date >= from_date)
        & (df_invoices["Invoice Date"].dt.date <= to_date)
    ]

    df_pay_cust = dfp_merged[
        (dfp_merged["Customer Name"] == customer_name)
        & (dfp_merged["Payment Date"].dt.date >= from_date)
        & (dfp_merged["Payment Date"].dt.date <= to_date)
    ]

    ledger_rows = []

    # Invoices => Debits
    for _, row in df_inv.iterrows():
        ledger_rows.append({
            "Date": row["Invoice Date"],
            "Txn Type": f"Invoice {row['Invoice ID']}",
            "Debits": row["Total Amount"],
            "Credits": 0.0
        })

    # Payments => Credits
    for _, row in df_pay_cust.iterrows():
        ledger_rows.append({
            "Date": row["Payment Date"],
            "Txn Type": f"Payment {row['Payment ID']}",
            "Debits": 0.0,
            "Credits": row["Payment Amount"]
        })

    ledger_df = pd.DataFrame(ledger_rows)
    ledger_df.sort_values(by="Date", inplace=True)

    df_inv_before = df_invoices[
        (df_invoices["Customer Name"] == customer_name)
        & (df_invoices["Invoice Date"].dt.date < from_date)
    ]["Total Amount"].sum()

    df_pay_before = dfp_merged[
        (dfp_merged["Customer Name"] == customer_name)
        & (dfp_merged["Payment Date"].dt.date < from_date)
    ]["Payment Amount"].sum()

    opening_balance = df_inv_before - df_pay_before
    running_balance = []
    current = opening_balance
    for i, row in ledger_df.iterrows():
        current = current + row["Debits"] - row["Credits"]
        running_balance.append(current)

    ledger_df["Running Balance"] = running_balance
    ledger_df["Date"] = ledger_df["Date"].dt.strftime("%d/%m/%Y")

    ledger_df = ledger_df[["Date", "Txn Type", "Debits", "Credits", "Running Balance"]]
    return ledger_df


def plot_aging_distribution(final_df):
    """
    Optional bar chart of aging buckets.
    """
    aging_cols = ["Current", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days"]
    if not all(col in final_df.columns for col in aging_cols):
        return
    sums = final_df[aging_cols].sum()
    fig, ax = plt.subplots()
    ax.bar(sums.index, sums.values)
    ax.set_title("Aging Distribution")
    ax.set_ylabel("Amount")
    st.pyplot(fig)

# --------------------------------------------------------------------------------
# SEGMENT WISE FUNCTION
# --------------------------------------------------------------------------------
def create_segment_wise_report(df_invoices, df_payments, company="All Companies"):
    """
    Partial payment logic with multi-bucket approach for segment wise.
    """
    if company != "All Companies":
        df_invoices = df_invoices[df_invoices["Company Name"] == company].copy()

    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Invoice Date", "Machine Revenue", "Parts Revenue", "Service Revenue"]],
        on="Invoice ID",
        how="left"
    )

    time_buckets = [
        ("Older Years",   date(1900,1,1),  date(2023,3,31)),
        ("FY 23-24 H1",   date(2023,4,1),  date(2023,9,30)),
        ("FY 23-24 H2",   date(2023,10,1), date(2024,3,31)),
        ("FY 24-25 Q1",   date(2024,4,1),  date(2024,6,30)),
        ("FY 24-25 Q2",   date(2024,7,1),  date(2024,9,30)),
        ("FY 24-25 Q3",   date(2024,10,1), date(2024,12,31)),
        ("FY 24-25 Q4",   date(2025,1,1),  date(2025,3,31)),
    ]

    segments = ["Machine", "Parts", "Service"]
    row_tuples = []
    for seg in segments:
        row_tuples.append((seg, "Outstanding as on Date"))
        row_tuples.append((seg, "Less: Payment Received"))
        row_tuples.append((seg, "Balance OS"))
    col_labels = [tb[0] for tb in time_buckets]
    index = pd.MultiIndex.from_tuples(row_tuples, names=["Segment", "Line"])
    seg_df = pd.DataFrame(index=index, columns=col_labels, data=0.0)

    for (col_label, start_d, end_d) in time_buckets:
        df_inv_range = df_invoices[
            (df_invoices["Invoice Date"].dt.date >= start_d)
            & (df_invoices["Invoice Date"].dt.date <= end_d)
        ]

        machine_os = df_inv_range["Machine Revenue"].sum()
        parts_os   = df_inv_range["Parts Revenue"].sum()
        service_os = df_inv_range["Service Revenue"].sum()

        df_pay_range = dfp_merged[
            (dfp_merged["Invoice Date"].dt.date >= start_d)
            & (dfp_merged["Invoice Date"].dt.date <= end_d)
            & (dfp_merged["Payment Date"].dt.date >= start_d)
            & (dfp_merged["Payment Date"].dt.date <= end_d)
        ]

        machine_pay_total = 0.0
        parts_pay_total   = 0.0
        service_pay_total = 0.0

        for _, row2 in df_pay_range.iterrows():
            inv_mach  = row2.get("Machine Revenue", 0.0)
            inv_parts = row2.get("Parts Revenue", 0.0)
            inv_serv  = row2.get("Service Revenue", 0.0)
            inv_sum   = inv_mach + inv_parts + inv_serv
            pay_amt   = row2["Payment Amount"]

            if inv_sum > 0:
                mach_share   = pay_amt * (inv_mach / inv_sum)
                parts_share  = pay_amt * (inv_parts / inv_sum)
                serv_share   = pay_amt * (inv_serv / inv_sum)
            else:
                mach_share = parts_share = serv_share = 0.0

            machine_pay_total += mach_share
            parts_pay_total   += parts_share
            service_pay_total += serv_share

        seg_df.loc[("Machine", "Outstanding as on Date"), col_label] = machine_os
        seg_df.loc[("Machine", "Less: Payment Received"), col_label] = machine_pay_total
        seg_df.loc[("Machine", "Balance OS"), col_label] = machine_os - machine_pay_total

        seg_df.loc[("Parts", "Outstanding as on Date"), col_label] = parts_os
        seg_df.loc[("Parts", "Less: Payment Received"), col_label] = parts_pay_total
        seg_df.loc[("Parts", "Balance OS"), col_label] = parts_os - parts_pay_total

        seg_df.loc[("Service", "Outstanding as on Date"), col_label] = service_os
        seg_df.loc[("Service", "Less: Payment Received"), col_label] = service_pay_total
        seg_df.loc[("Service", "Balance OS"), col_label] = service_os - service_pay_total

    return seg_df

# --------------------------------------------------------------------------------
# PAGE FUNCTIONS
# --------------------------------------------------------------------------------
def show_receivables_report():
    st.title("Receivables Report")

    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]
    group_opt = st.session_state["group_choice"]

    if st.button("Generate Receivables Report"):
        final_df = create_receivables_report(df_invoices, df_payments, from_dt, to_dt, group_opt)
        total_os = final_df["Total OS"].sum()
        col1, col2 = st.columns(2)
        col1.metric("Total Outstanding", f"{total_os:,.2f}")

        st.dataframe(final_df.style.format(precision=2))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            final_df.to_excel(writer, sheet_name="Receivables", index=False)
        st.download_button(
            "Download Excel",
            data=output.getvalue(),
            file_name="Receivables_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def show_banker_report():
    st.title("Banker Report")
    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]

    if st.button("Generate Banker Report"):
        banker_df = create_banker_report(df_invoices, df_payments, from_dt, to_dt)
        st.dataframe(banker_df.style.format(precision=2))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            banker_df.to_excel(writer, sheet_name="Banker Report", index=False)
        st.download_button(
            "Download Excel",
            data=output.getvalue(),
            file_name="Banker_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def show_customer_ledger():
    st.title("Customer Ledger")
    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]

    all_cust = sorted(df_invoices["Customer Name"].dropna().unique())
    chosen_cust = st.selectbox("Select Customer:", all_cust, index=0)

    if st.button("Generate Customer Ledger"):
        ledger_df = create_customer_ledger(df_invoices, df_payments, from_dt, to_dt, chosen_cust)
        st.dataframe(ledger_df.style.format(precision=2))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            ledger_df.to_excel(writer, sheet_name="Customer Ledger", index=False)
        st.download_button(
            "Download Excel",
            data=output.getvalue(),
            file_name="Customer_Ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def show_segment_wise():
    st.title("Segment Wise Outstanding & Payment")

    # Let user pick company
    companies = ["All Companies"] + sorted(df_invoices["Company Name"].dropna().unique())
    chosen_company = st.selectbox("Select Company:", companies, index=0)

    if st.button("Generate Segment-Wise Report"):
        seg_df = create_segment_wise_report(df_invoices, df_payments, chosen_company)

        # -- STYLING CHANGES HERE --
        def highlight_balance(row):
            """
            If row is for 'Balance OS', highlight that row and make it bold.
            """
            if row.name[1] == "Balance OS":
                return ["font-weight: bold; background-color: #FFFACD;" for _ in row]
            else:
                return ["" for _ in row]

        styled_df = seg_df.style \
            .format(precision=2) \
            .apply(highlight_balance, axis=1) \
            .set_table_styles([
                {
                    "selector": "th",
                    "props": [
                        ("font-weight","bold"),
                        ("background-color","#EEE"),
                        ("color","#333")
                    ]
                }
            ], overwrite=False)

        st.subheader("Segment Wise Styled Table")
        st.dataframe(styled_df, use_container_width=True)

        # Provide Excel download
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            seg_df.to_excel(writer, sheet_name="SegmentWise", index=True)
        st.download_button(
            label="Download Excel",
            data=output.getvalue(),
            file_name="SegmentWise_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


def show_management_dashboard():
    st.title("Management Dashboard - 360° View")

    # Quick KPI row
    total_inv = df_invoices["Total Amount"].sum()
    total_pay = df_payments["Payment Amount"].sum()
    total_os  = total_inv - total_pay

    c1, c2, c3 = st.columns(3)
    c1.metric("Total Invoiced", f"{total_inv:,.2f}")
    c2.metric("Total Paid",     f"{total_pay:,.2f}")
    c3.metric("Outstanding",    f"{total_os:,.2f}")

    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Company Name", "Customer Name"]],
        on="Invoice ID",
        how="left"
    )

    inv_company = df_invoices.groupby("Company Name")["Total Amount"].sum().rename("InvTotal")
    pay_company = dfp_merged.groupby("Company Name")["Payment Amount"].sum().rename("PayTotal")
    mg = pd.DataFrame(inv_company).join(pay_company, how="outer").fillna(0)
    mg["Outstanding"] = mg["InvTotal"] - mg["PayTotal"]
    st.subheader("Total Outstanding by Company")
    st.bar_chart(data=mg.reset_index(), x="Company Name", y="Outstanding")

    inv_cust = df_invoices.groupby("Customer Name")["Total Amount"].sum().rename("InvTotal")
    pay_cust = dfp_merged.groupby("Customer Name")["Payment Amount"].sum().rename("PayTotal")
    mg2 = pd.DataFrame(inv_cust).join(pay_cust, how="outer").fillna(0)
    mg2["Outstanding"] = mg2["InvTotal"] - mg2["PayTotal"]
    mg2_sorted = mg2.sort_values("Outstanding", ascending=False).head(5)

    st.subheader("Top 5 Customers by Outstanding")
    st.table(mg2_sorted[["Outstanding"]])

    st.subheader("Monthly Invoice Trend")
    temp_df = df_invoices.copy()
    temp_df["Invoice Month"] = temp_df["Invoice Date"].dt.to_period("M")
    monthly_sum = temp_df.groupby("Invoice Month")["Total Amount"].sum().reset_index()
    monthly_sum["Invoice Month"] = monthly_sum["Invoice Month"].astype(str)
    st.line_chart(data=monthly_sum, x="Invoice Month", y="Total Amount")

# --------------------------------------------------------------------------------
# SIDEBAR & NAV
# --------------------------------------------------------------------------------
st.sidebar.header("Global Filters")

if "from_date" not in st.session_state:
    st.session_state["from_date"] = min_date
if "to_date" not in st.session_state:
    st.session_state["to_date"] = max_date

st.session_state["from_date"] = st.sidebar.date_input(
    "From Date (Invoice):",
    value=st.session_state["from_date"],
    min_value=min_date,
    max_value=max_date
)
st.session_state["to_date"] = st.sidebar.date_input(
    "To Date (Invoice):",
    value=st.session_state["to_date"],
    min_value=min_date,
    max_value=max_date
)

# Updated grouping options to include Branch Wise Details
group_opts = ["Grand Total", "Customer ID", "Company Name", "Customer Name", "Branch Wise Details"]
if "group_choice" not in st.session_state:
    st.session_state["group_choice"] = group_opts[0]
st.session_state["group_choice"] = st.sidebar.selectbox(
    "Group By (Receivables):",
    group_opts,
    index=0
)

st.sidebar.write("---")

page = st.sidebar.radio(
    "Go to Page:",
    [
        "Receivables Report",
        "Banker Report",
        "Customer Ledger",
        "Segment Wise",
        "Management Dashboard"
    ]
)

# --------------------------------------------------------------------------------
# MAIN LOGIC
# --------------------------------------------------------------------------------
if page == "Receivables Report":
    show_receivables_report()
elif page == "Banker Report":
    show_banker_report()
elif page == "Customer Ledger":
    show_customer_ledger()
elif page == "Segment Wise":
    show_segment_wise()
elif page == "Management Dashboard":
    show_management_dashboard()
