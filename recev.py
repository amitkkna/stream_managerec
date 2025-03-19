import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
from datetime import date, datetime

# --------------------------------------------------------------------------------
# PAGE CONFIGURATION
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

EXCEL_FILE_PATH = r"exceldata/SVP Sample data.xlsx"
df_invoices, df_payments = load_data(EXCEL_FILE_PATH)

min_date = df_invoices["Invoice Date"].min().date()
max_date = df_invoices["Invoice Date"].max().date()

# --------------------------------------------------------------------------------
# HELPER FUNCTION: Append Grand Total Row
# --------------------------------------------------------------------------------
def append_grand_total_row(df, label_col, label="Grand Total"):
    """
    Appends a grand total row to the given DataFrame.
    For numeric columns, sums up their values.
    For non-numeric columns, sets an empty string except for the label_col,
    which gets the specified label.
    """
    df_numeric = df.select_dtypes(include=[np.number])
    total_row = df_numeric.sum()
    row = {}
    for col in df.columns:
        if col in df_numeric.columns:
            row[col] = total_row[col]
        else:
            if col == label_col:
                row[col] = label
            else:
                row[col] = ""
    return pd.concat([df, pd.DataFrame([row])], ignore_index=True)

# --------------------------------------------------------------------------------
# HELPER FUNCTION: Calculate Color Code (CC)
# --------------------------------------------------------------------------------
def calc_cc(row):
    """
    If invoice is overdue (today > Due Date):
      - no payment => red (🟥)
      - partial payment => blue (🟦)
      - full payment => green (🟩)
    Otherwise, blank.
    """
    due = row.get("Due Date")
    if pd.isnull(due):
        return ""
    due_date = due.date() if isinstance(due, pd.Timestamp) else due
    today = date.today()
    if today > due_date:
        if row["PaidToDate"] == 0:
            return "🟥"
        elif row["PaidToDate"] < row["Total Amount"]:
            return "🟦"
        else:
            return "🟩"
    else:
        return ""

# --------------------------------------------------------------------------------
# HELPER FUNCTIONS: Aging Calculation for Pending Invoices
# --------------------------------------------------------------------------------
def calc_days_overdue(row):
    """
    Returns how many days overdue the invoice is, or 0 if not overdue or no due date.
    """
    if pd.isnull(row.get("Due Date")):
        return 0
    d = row["Due Date"]
    due_date = d.date() if isinstance(d, pd.Timestamp) else d
    delta = (date.today() - due_date).days
    return delta if delta > 0 else 0

def aging_bucket(days):
    """
    Converts a numeric days-overdue into an aging bucket.
    """
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

# --------------------------------------------------------------------------------
# HELPER FUNCTIONS FOR REPORTS
# --------------------------------------------------------------------------------

def create_receivables_report(df_invoices, df_payments, from_date, to_date, group_by):
    """
    Returns (final_df, overall_total_os).
    final_df is the pivoted table (with an appended row if needed).
    overall_total_os is the sum of all outstanding (no double count).
    """
    today_date = date.today()
    df_pay_lim = df_payments[df_payments["Payment Date"].dt.date <= today_date].copy()
    paid_agg = df_pay_lim.groupby("Invoice ID")["Payment Amount"].sum().rename("PaidToDate")

    merged = df_invoices.merge(paid_agg, on="Invoice ID", how="left")
    merged["PaidToDate"] = merged["PaidToDate"].fillna(0.0)
    merged["Outstanding"] = merged["Total Amount"] - merged["PaidToDate"]

    filtered_df = merged[
        (merged["Invoice Date"].dt.date >= from_date) &
        (merged["Invoice Date"].dt.date <= to_date)
    ].copy()

    overall_total_os = filtered_df["Outstanding"].sum()

    # Days Past Due & Aging
    filtered_df["Days Past Due"] = (pd.to_datetime(today_date) - filtered_df["Due Date"]).dt.days.fillna(0)
    def local_aging_bucket(days):
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
    filtered_df["Aging Bucket"] = filtered_df["Days Past Due"].apply(local_aging_bucket)

    # Distribute partial payments
    def distribute_partial_payments(row):
        total_inv = row["Total Amount"]
        paid = row["PaidToDate"]
        if total_inv <= 0:
            return row["Machine Revenue"], row["Parts Revenue"], row["Service Revenue"]
        mach_share = row["Machine Revenue"] / total_inv
        parts_share = row["Parts Revenue"] / total_inv
        serv_share = row["Service Revenue"] / total_inv
        mach_paid = paid * mach_share
        parts_paid = paid * parts_share
        serv_paid = paid * serv_share
        return (
            row["Machine Revenue"] - mach_paid,
            row["Parts Revenue"]   - parts_paid,
            row["Service Revenue"] - serv_paid
        )

    filtered_df["Machine OS"], filtered_df["Parts OS"], filtered_df["Service OS"] = zip(
        *filtered_df.apply(distribute_partial_payments, axis=1)
    )

    # Determine grouping
    if group_by == "Grand Total":
        group_col = None
    elif group_by == "Branch Wise Details":
        group_col = "Branch"
    else:
        group_col = group_by

    # Summation
    if group_col:
        df_total = filtered_df.groupby(group_col)["Outstanding"].sum().rename("Total OS")
    else:
        df_total = pd.Series(filtered_df["Outstanding"].sum(), index=["Grand Total"], name="Total OS")

    # Aging pivot
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

    # Summation for line items
    if group_col:
        df_machine = filtered_df.groupby(group_col)["Machine OS"].sum().rename("Machine OS")
        df_parts   = filtered_df.groupby(group_col)["Parts OS"].sum().rename("Parts OS")
        df_service = filtered_df.groupby(group_col)["Service OS"].sum().rename("Service OS")
    else:
        df_machine = pd.Series(filtered_df["Machine OS"].sum(), index=["Grand Total"], name="Machine OS")
        df_parts   = pd.Series(filtered_df["Parts OS"].sum(), index=["Grand Total"], name="Parts OS")
        df_service = pd.Series(filtered_df["Service OS"].sum(), index=["Grand Total"], name="Service OS")

    final_df = pd.DataFrame(df_total).join(aging_pivot, how="left")
    final_df = final_df.join(df_machine, how="left").join(df_parts, how="left").join(df_service, how="left")
    final_df.reset_index(inplace=True)

    if group_col is None:
        final_df.rename(columns={"index": "Group"}, inplace=True)
    else:
        col_name = "Branch" if group_by == "Branch Wise Details" else group_by
        final_df.rename(columns={col_name: "Group"}, inplace=True)

    col_order = [
        "Group", "Total OS",
        "Current", "1-30 Days", "31-60 Days", "61-90 Days", "90+ Days",
        "Machine OS", "Parts OS", "Service OS"
    ]
    final_df = final_df[col_order]

    # Append grand total row only if group_by != "Grand Total"
    if group_by != "Grand Total":
        final_df = append_grand_total_row(final_df, label_col="Group", label="Grand Total")

    return final_df, overall_total_os

def create_banker_report(df_invoices, df_payments, from_date, to_date, company="All"):
    """
    Banker Report
    """
    if company != "All":
        df_invoices = df_invoices[df_invoices["Company Name"] == company]
    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Customer Name", "Company Name"]],
        on="Invoice ID",
        how="left"
    )
    df_inv_before = df_invoices[df_invoices["Invoice Date"].dt.date < from_date]
    df_pay_before = dfp_merged[dfp_merged["Payment Date"].dt.date < from_date]
    df_inv_range = df_invoices[
        (df_invoices["Invoice Date"].dt.date >= from_date) &
        (df_invoices["Invoice Date"].dt.date <= to_date)
    ]
    df_pay_range = dfp_merged[
        (dfp_merged["Payment Date"].dt.date >= from_date) &
        (dfp_merged["Payment Date"].dt.date <= to_date)
    ]
    inv_before_agg = df_inv_before.groupby("Customer Name")["Total Amount"].sum().rename("InvBefore")
    pay_before_agg = df_pay_before.groupby("Customer Name")["Payment Amount"].sum().rename("PayBefore")
    inv_range_agg  = df_inv_range.groupby("Customer Name")["Total Amount"].sum().rename("InvRange")
    pay_range_agg  = df_pay_range.groupby("Customer Name")["Payment Amount"].sum().rename("PayRange")

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

    col_order = ["Customer Name", "Opening Balance", "Debits (Invoices)", "Credits (Payments)", "Balance"]
    banker_df = banker_df[col_order]
    return banker_df

def create_customer_ledger(df_invoices, df_payments, from_date, to_date, customer_name, company="All", branch="All"):
    """
    Builds a ledger with color-coded invoice rows if overdue, plus opening/closing balances.
    """
    dfp_merged = df_payments.merge(
        df_invoices[["Invoice ID", "Customer Name", "Company Name", "Branch", "Due Date", "Total Amount"]],
        on="Invoice ID",
        how="left"
    )
    df_inv = df_invoices[
        (df_invoices["Customer Name"] == customer_name) &
        (df_invoices["Invoice Date"].dt.date >= from_date) &
        (df_invoices["Invoice Date"].dt.date <= to_date)
    ]
    if company != "All":
        df_inv = df_inv[df_inv["Company Name"] == company]
    if branch != "All":
        df_inv = df_inv[df_inv["Branch"] == branch]

    df_pay_cust = dfp_merged[
        (dfp_merged["Customer Name"] == customer_name) &
        (dfp_merged["Payment Date"].dt.date >= from_date) &
        (dfp_merged["Payment Date"].dt.date <= to_date)
    ]
    if company != "All":
        df_pay_cust = df_pay_cust[df_pay_cust["Company Name"] == company]
    if branch != "All":
        df_pay_cust = df_pay_cust[df_pay_cust["Branch"] == branch]

    paid_dict = df_pay_cust.groupby("Invoice ID")["Payment Amount"].sum().to_dict()

    ledger_rows = []
    for _, row in df_inv.iterrows():
        invoice_id = row["Invoice ID"]
        total = row["Total Amount"]
        paid  = paid_dict.get(invoice_id, 0.0)
        due_date = row["Due Date"]
        if pd.notnull(due_date):
            d = due_date.date()
            overdue = (date.today() > d)
        else:
            overdue = False

        if overdue:
            if paid == 0:
                cc = "🟥"
            elif paid < total:
                cc = "🟦"
            else:
                cc = "🟩"
        else:
            cc = ""

        ledger_rows.append({
            "Date": row["Invoice Date"],
            "Txn Type": f"Invoice {invoice_id}",
            "CC": cc,
            "Debits": total,
            "Credits": 0.0
        })

    for _, row in df_pay_cust.iterrows():
        ledger_rows.append({
            "Date": row["Payment Date"],
            "Txn Type": f"Payment {row['Payment ID']}",
            "CC": "",
            "Debits": 0.0,
            "Credits": row["Payment Amount"]
        })

    ledger_df = pd.DataFrame(ledger_rows)
    ledger_df.sort_values(by="Date", inplace=True)

    # Opening balance from before 'from_date'
    df_inv_before = df_invoices[
        (df_invoices["Customer Name"] == customer_name) &
        (df_invoices["Invoice Date"].dt.date < from_date)
    ]
    if company != "All":
        df_inv_before = df_inv_before[df_inv_before["Company Name"] == company]
    if branch != "All":
        df_inv_before = df_inv_before[df_inv_before["Branch"] == branch]

    df_pay_before = dfp_merged[
        (dfp_merged["Customer Name"] == customer_name) &
        (dfp_merged["Payment Date"].dt.date < from_date)
    ]
    if company != "All":
        df_pay_before = df_pay_before[df_pay_before["Company Name"] == company]
    if branch != "All":
        df_pay_before = df_pay_before[df_pay_before["Branch"] == branch]

    opening_balance = df_inv_before["Total Amount"].sum() - df_pay_before["Payment Amount"].sum()

    # Compute running balance
    running_balance = []
    current = opening_balance
    for _, row in ledger_df.iterrows():
        current = current + row["Debits"] - row["Credits"]
        running_balance.append(current)

    ledger_df["Running Balance"] = running_balance
    ledger_df["Date"] = ledger_df["Date"].dt.strftime("%d/%m/%Y")
    ledger_df = ledger_df[["Date", "Txn Type", "CC", "Debits", "Credits", "Running Balance"]]

    # Insert opening/closing
    opening_row = pd.DataFrame({
        "Date": ["Opening Balance"],
        "Txn Type": ["Opening Balance"],
        "CC": [""],
        "Debits": [0.0],
        "Credits": [0.0],
        "Running Balance": [opening_balance]
    })
    closing_balance = ledger_df["Running Balance"].iloc[-1] if not ledger_df.empty else opening_balance
    closing_row = pd.DataFrame({
        "Date": ["Closing Balance"],
        "Txn Type": ["Closing Balance"],
        "CC": [""],
        "Debits": [0.0],
        "Credits": [0.0],
        "Running Balance": [closing_balance]
    })

    final_ledger = pd.concat([opening_row, ledger_df, closing_row], ignore_index=True)
    return final_ledger

def create_segment_wise_report(df_invoices, df_payments, company="All Companies"):
    """
    Segment-wise partial payments approach.
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
            (df_invoices["Invoice Date"].dt.date >= start_d) &
            (df_invoices["Invoice Date"].dt.date <= end_d)
        ]
        machine_os = df_inv_range["Machine Revenue"].sum()
        parts_os   = df_inv_range["Parts Revenue"].sum()
        service_os = df_inv_range["Service Revenue"].sum()

        df_pay_range = dfp_merged[
            (dfp_merged["Invoice Date"].dt.date >= start_d) &
            (dfp_merged["Invoice Date"].dt.date <= end_d) &
            (dfp_merged["Payment Date"].dt.date >= start_d) &
            (dfp_merged["Payment Date"].dt.date <= end_d)
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
                mach_share  = pay_amt * (inv_mach / inv_sum)
                parts_share = pay_amt * (inv_parts / inv_sum)
                serv_share  = pay_amt * (inv_serv / inv_sum)
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

def show_invoice_summary():
    today_date = date.today()
    df_pay_lim = df_payments[df_payments["Payment Date"].dt.date <= today_date].copy()
    paid_agg = df_pay_lim.groupby("Invoice ID")["Payment Amount"].sum().rename("PaidToDate")
    merged = df_invoices.merge(paid_agg, on="Invoice ID", how="left")
    merged["PaidToDate"] = merged["PaidToDate"].fillna(0.0)
    merged["Outstanding"] = merged["Total Amount"] - merged["PaidToDate"]
    summary = merged.groupby(["Company Name", "Branch"]).agg(
        Total_Invoices=("Invoice ID", "count"),
        Pending_Invoices=("Outstanding", lambda x: (x > 0).sum())
    ).reset_index()
    st.dataframe(summary.style.format(precision=2))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="Invoice Summary", index=False)
    st.download_button(
        "Download Excel",
        data=output.getvalue(),
        file_name="Invoice_Summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def plot_aging_distribution(final_df):
    """
    Example bar chart for aging columns
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
# PAGE FUNCTIONS
# --------------------------------------------------------------------------------

def show_receivables_report():
    """
    Displays the Receivables Report with a single top-level metric for total outstanding
    that does not double count the appended row.
    """
    st.title("Receivables Report")
    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]
    group_opt = st.session_state["group_choice"]

    if st.button("Generate Receivables Report"):
        final_df, overall_total_os = create_receivables_report(
            df_invoices, df_payments,
            from_dt, to_dt,
            group_opt
        )

        col1, col2 = st.columns(2)
        col1.metric("Total Outstanding", f"{overall_total_os:,.2f}")

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
    """
    Banker Report
    """
    st.title("Banker Report")
    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]

    companies = ["All"] + sorted(df_invoices["Company Name"].dropna().unique())
    chosen_company = st.selectbox("Select Company:", companies, index=0)

    if st.button("Generate Banker Report"):
        banker_df = create_banker_report(df_invoices, df_payments, from_dt, to_dt, company=chosen_company)
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
    """
    Customer Ledger page, with Pending & Fully Collected Invoices sections,
    including color-coded columns and grand total rows.
    """
    st.title("Customer Ledger")
    from_dt = st.session_state["from_date"]
    to_dt   = st.session_state["to_date"]

    all_cust = sorted(df_invoices["Customer Name"].dropna().unique())
    chosen_cust = st.selectbox("Select Customer:", all_cust, index=0)

    companies = ["All"] + sorted(df_invoices["Company Name"].dropna().unique())
    chosen_company = st.selectbox("Select Company:", companies, index=0)

    branches = ["All"] + sorted(df_invoices["Branch"].dropna().unique())
    chosen_branch = st.selectbox("Select Branch:", branches, index=0)

    if st.button("Generate Customer Ledger"):
        # 1) Transaction Ledger
        ledger_df = create_customer_ledger(
            df_invoices, df_payments,
            from_dt, to_dt,
            customer_name=chosen_cust,
            company=chosen_company,
            branch=chosen_branch
        )
        st.subheader("Transaction Ledger")
        st.dataframe(ledger_df.style.format(precision=2))

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            ledger_df.to_excel(writer, sheet_name="Customer Ledger", index=False)
        st.download_button(
            "Download Ledger as Excel",
            data=output.getvalue(),
            file_name="Customer_Ledger.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 2) Pending Invoices
        df_inv_pending = df_invoices[
            (df_invoices["Customer Name"] == chosen_cust) &
            (df_invoices["Invoice Date"].dt.date >= from_dt) &
            (df_invoices["Invoice Date"].dt.date <= to_dt)
        ].copy()
        if chosen_company != "All":
            df_inv_pending = df_inv_pending[df_inv_pending["Company Name"] == chosen_company]
        if chosen_branch != "All":
            df_inv_pending = df_inv_pending[df_inv_pending["Branch"] == chosen_branch]

        if "Due Date" in df_inv_pending.columns:
            df_inv_pending = df_inv_pending[
                [
                    "Invoice ID",
                    "Customer ID",
                    "Customer Name",
                    "Invoice Date",
                    "Due Date",
                    "Company Name",
                    "Branch",
                    "Machine Revenue",
                    "Parts Revenue",
                    "Service Revenue",
                    "Total Amount"
                ]
            ]
        else:
            df_inv_pending = df_inv_pending[
                [
                    "Invoice ID",
                    "Customer ID",
                    "Customer Name",
                    "Invoice Date",
                    "Company Name",
                    "Branch",
                    "Machine Revenue",
                    "Parts Revenue",
                    "Service Revenue",
                    "Total Amount"
                ]
            ]

        today_date = date.today()
        df_pay_lim = df_payments[df_payments["Payment Date"].dt.date <= today_date].copy()
        paid_agg = df_pay_lim.groupby("Invoice ID")["Payment Amount"].sum().rename("PaidToDate")

        merged_inv = df_inv_pending.merge(paid_agg, on="Invoice ID", how="left")
        merged_inv["PaidToDate"] = merged_inv["PaidToDate"].fillna(0.0)
        merged_inv["Outstanding"] = merged_inv["Total Amount"] - merged_inv["PaidToDate"]

        # Color code
        merged_inv["CC"] = merged_inv.apply(calc_cc, axis=1)

        # Insert 'Days Overdue' and 'Aging Bucket'
        merged_inv["Days Overdue"] = merged_inv.apply(calc_days_overdue, axis=1).astype(int)  # whole number
        merged_inv["Aging Bucket"] = merged_inv["Days Overdue"].apply(aging_bucket)

        # We want them after "Service OS" but we haven't computed line OS yet.
        # We'll do that after we compute line item OS. So let's do line OS first.
        pending_invoices_df = merged_inv[merged_inv["Outstanding"] > 0]

        # Grand total row
        pending_invoices_df = append_grand_total_row(
            pending_invoices_df,
            label_col="Invoice ID",
            label="Grand Total"
        )

        # Calculate line item OS
        def calc_line_os(row):
            total = row["Total Amount"]
            paid  = row["PaidToDate"]
            if total > 0:
                mach_os = row["Machine Revenue"] - (paid * row["Machine Revenue"] / total)
                parts_os= row["Parts Revenue"]   - (paid * row["Parts Revenue"] / total)
                serv_os = row["Service Revenue"] - (paid * row["Service Revenue"] / total)
            else:
                mach_os = row["Machine Revenue"]
                parts_os= row["Parts Revenue"]
                serv_os = row["Service Revenue"]
            return pd.Series({
                "Machine OS": mach_os,
                "Parts OS":   parts_os,
                "Service OS": serv_os
            })

        os_df = pending_invoices_df.apply(calc_line_os, axis=1)
        pending_invoices_df = pd.concat([pending_invoices_df, os_df], axis=1)

        # Now reorder columns so that "Days Overdue" and "Aging Bucket" appear AFTER "Service OS".
        cols = list(pending_invoices_df.columns)
        for c in ["Days Overdue", "Aging Bucket"]:
            if c in cols:
                cols.remove(c)
                if "Service OS" in cols:
                    idx = cols.index("Service OS") + 1
                    cols.insert(idx, c)
                else:
                    cols.append(c)
        pending_invoices_df = pending_invoices_df[cols]

        st.subheader("Pending Invoices (Payment yet to be collected)")
        st.dataframe(
            pending_invoices_df.style.format({
                "Total Amount": "{:,.2f}",
                "Machine Revenue": "{:,.2f}",
                "Parts Revenue": "{:,.2f}",
                "Service Revenue": "{:,.2f}",
                "PaidToDate": "{:,.2f}",
                "Outstanding": "{:,.2f}",
                "Machine OS": "{:,.2f}",
                "Parts OS": "{:,.2f}",
                "Service OS": "{:,.2f}"
            })
        )

        output2 = io.BytesIO()
        with pd.ExcelWriter(output2, engine="openpyxl") as writer:
            pending_invoices_df.to_excel(writer, sheet_name="Pending Invoices", index=False)
        st.download_button(
            "Download Pending Invoices as Excel",
            data=output2.getvalue(),
            file_name="Pending_Invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # 3) Fully Collected Invoices
        fully_paid_invoices_df = merged_inv[merged_inv["Outstanding"] == 0]
        if not fully_paid_invoices_df.empty:
            # Recompute line item OS
            os_df_paid = fully_paid_invoices_df.apply(calc_line_os, axis=1)
            fully_paid_invoices_df = pd.concat([fully_paid_invoices_df, os_df_paid], axis=1)

            # Also reorder "Days Overdue" / "Aging Bucket" if you want them after "Service OS" here too
            fully_paid_invoices_df["Days Overdue"] = fully_paid_invoices_df.apply(calc_days_overdue, axis=1).astype(int)
            fully_paid_invoices_df["Aging Bucket"] = fully_paid_invoices_df["Days Overdue"].apply(aging_bucket)

            # Insert them after "Service OS" as well
            cols2 = list(fully_paid_invoices_df.columns)
            for c in ["Days Overdue", "Aging Bucket"]:
                if c in cols2:
                    cols2.remove(c)
                    if "Service OS" in cols2:
                        idx = cols2.index("Service OS") + 1
                        cols2.insert(idx, c)
                    else:
                        cols2.append(c)
            fully_paid_invoices_df = fully_paid_invoices_df[cols2]

            # Append grand total
            fully_paid_invoices_df = append_grand_total_row(
                fully_paid_invoices_df,
                label_col="Invoice ID",
                label="Grand Total"
            )

            st.subheader("Fully Collected Invoices (No Outstanding)")
            st.dataframe(
                fully_paid_invoices_df.style.format({
                    "Total Amount": "{:,.2f}",
                    "PaidToDate": "{:,.2f}",
                    "Outstanding": "{:,.2f}",
                    "Machine Revenue": "{:,.2f}",
                    "Parts Revenue": "{:,.2f}",
                    "Service Revenue": "{:,.2f}",
                    "Machine OS": "{:,.2f}",
                    "Parts OS": "{:,.2f}",
                    "Service OS": "{:,.2f}"
                })
            )

            output3 = io.BytesIO()
            with pd.ExcelWriter(output3, engine="openpyxl") as writer:
                fully_paid_invoices_df.to_excel(writer, sheet_name="Fully Paid Invoices", index=False)
            st.download_button(
                "Download Fully Collected Invoices as Excel",
                data=output3.getvalue(),
                file_name="Fully_Paid_Invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.subheader("No Fully Collected Invoices found (all have outstanding amounts).")

def show_segment_wise():
    """
    Segment Wise page
    """
    st.title("Segment Wise Outstanding & Payment")
    companies = ["All Companies"] + sorted(df_invoices["Company Name"].dropna().unique())
    chosen_company = st.selectbox("Select Company:", companies, index=0)
    if st.button("Generate Segment-Wise Report"):
        seg_df = create_segment_wise_report(df_invoices, df_payments, chosen_company)

        def highlight_balance(row):
            if row.name[1] == "Balance OS":
                return ["font-weight: bold; background-color: #FFFACD;" for _ in row]
            else:
                return ["" for _ in row]

        styled_df = seg_df.style.format(precision=2)\
            .apply(highlight_balance, axis=1)\
            .set_table_styles([
                {
                    "selector": "th",
                    "props": [
                        ("font-weight", "bold"),
                        ("background-color", "#EEE"),
                        ("color", "#333")
                    ]
                }
            ], overwrite=False)

        st.subheader("Segment Wise Styled Table")
        st.dataframe(styled_df, use_container_width=True)

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
    """
    Management Dashboard
    """
    st.title("Management Dashboard - 360° View")
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

def show_invoice_summary_page():
    """
    Invoice Summary by Company and Branch
    """
    st.title("Invoice Summary by Company and Branch")
    show_invoice_summary()

# --------------------------------------------------------------------------------
# SIDEBAR & NAVIGATION
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
        "Management Dashboard",
        "Invoice Summary"
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
elif page == "Invoice Summary":
    show_invoice_summary_page()
