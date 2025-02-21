import streamlit as st
import pandas as pd


def process_excel(file):
    """
    Process the Excel file to:
      - Skip the first 9 rows and use only columns B:F.
      - Remove the final row that only contains a total in the Amount Due column.
      - Rename/convert the columns as required.
    """
    # Read the Excel file:
    # - header=9 means that row 10 (0-indexed row 9) will be used as the header.
    # - usecols="B:G" means that we ignore column A (which is blank).
    df = pd.read_excel(file, header=9, usecols="B:G")

    # The Excel file has these columns:
    #   'Date', 'Invoice No.', 'Customer Name', 'Amount Due', 'Card ID'
    #
    # Remove the ending total row: that row has value only in the "Amount Due" column;
    # the other columns will be NaN. We drop any row where *any* of the columns except
    # "Amount Due" are missing.
    df = df.dropna(
        subset=["Date", "Invoice No.", "Customer Name", "Card ID"], how="any"
    )

    # --- Process and convert each column as needed ---

    # 1. Debtor Reference: derived from "Card ID"
    # If "Card ID" equals "*None", use "Customer Name" instead.
    df["Debtor Reference"] = df.apply(
        lambda row: row["Customer Name"]
        if str(row["Card ID"]).strip() == "*None"
        else row["Card ID"],
        axis=1,
    )

    # 2. Transaction Type: if Amount Due is negative -> "CRD", otherwise "INV"
    df["Transaction Type"] = df["Amount Due"].apply(lambda x: "CRD" if x < 0 else "INV")

    # 3. Document Number: convert "Invoice No." by stripping any leading zeros if the
    #    value is entirely numeric. Otherwise, leave it unchanged.
    def convert_invoice(invoice):
        if pd.isna(invoice):
            return ""
        # Convert to string first.
        invoice_str = str(invoice).strip()
        # If the string consists only of digits, convert to int to drop leading zeros.
        if invoice_str.isdigit():
            return str(int(invoice_str))
        else:
            return invoice_str

    df["Document Number"] = df["Invoice No."].apply(convert_invoice)

    # 4. Document Date: use the value in "Date". If it is a Timestamp, format as YYYY-MM-DD.
    def convert_date(val):
        if pd.isna(val):
            return ""
        if isinstance(val, pd.Timestamp):
            return val.strftime("%Y-%m-%d")
        return str(val)

    df["Document Date"] = df["Date"].apply(convert_date)

    # 5. Document Balance: use the value in "Amount Due" formatted to 2 decimal places.
    df["Document Balance"] = df["Amount Due"].apply(lambda x: f"{x:.2f}")

    # Finally, select the columns in the required order:
    final_df = df[
        [
            "Debtor Reference",
            "Transaction Type",
            "Document Number",
            "Document Date",
            "Document Balance",
        ]
    ]

    return final_df


def main():
    st.title("GSD-84 Reformatting Tool")
    st.write(
        """
        Upload an XLSX file with the following characteristics:
          - The first 9 rows are unneeded.
          - The header row (row 10) contains:
              - Column B: Date
              - Column C: Invoice No.
              - Column D: Customer Name
              - Column E: Amount (skipped)
              - Column F: Amount Due
              - Column G: Card ID
          - There is a final row containing a total in Amount Due only.
        
        The app will convert the data into a CSV file with these columns:
          Debtor Reference, Transaction Type, Document Number, Document Date, Document Balance

        The conversion rules are:
          - **Debtor Reference:** taken from "Card ID", unless the value is "*None", in which case "Customer Name" is used.
          - **Transaction Type:** "CRD" if Amount Due is negative; otherwise "INV"
          - **Document Number:** if "Invoice No." is entirely numeric (e.g., "00048071"), the leading zeros are removed (resulting in "48071"); otherwise, the value is left unchanged.
          - **Document Date:** taken from "Date" (formatted as YYYY-MM-DD if a date object)
          - **Document Balance:** taken from "Amount Due" and formatted with 2 decimals (e.g., "174.95")
        """
    )

    # File uploader accepts only XLSX files.
    uploaded_file = st.file_uploader("Choose an XLSX file", type="xlsx")
    if uploaded_file is not None:
        try:
            # Process the Excel file.
            df_final = process_excel(uploaded_file)

            st.subheader("Preview of converted data")
            st.dataframe(df_final)

            # Convert the final DataFrame to CSV format.
            csv_data = df_final.to_csv(index=False).encode("utf-8")
            st.download_button(
                label="Download CSV",
                data=csv_data,
                file_name="converted_data.csv",
                mime="text/csv",
            )
        except Exception as e:
            st.error(f"Error processing file: {e}")


if __name__ == "__main__":
    main()
