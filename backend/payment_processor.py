"""
payment_processor.py

Purpose
-------
Reads the Myntra Revamped Payment Report and generates a Payment Register.

Output Columns

Settlement Date
UTR Number
Payment Amount
"""

from pathlib import Path
from openpyxl import load_workbook
from utils import format_payment_register


import pandas as pd


class PaymentProcessor:

    def __init__(self, excel_file):

        self.excel_file = excel_file

        self.df = None

    # -----------------------------------------------------

    def load_excel(self):

        self.df = pd.read_excel(
            self.excel_file,
            engine="openpyxl"
        )

        self.df.columns = (
            self.df.columns
            .str.strip()
            .str.replace("\n", "")
            .str.replace("\r", "")
        )

    # -----------------------------------------------------

    def extract_postpaid(self):

        required = [
            "settlement_date_postpaid_payment",
            "UTR_Number_Postpaid",
            "Settled_Amount_Postpaid"
        ]

        df = self.df.copy()

        df = df[df["UTR_Number_Postpaid"].notna()]

        df = df[df["UTR_Number_Postpaid"].astype(str).str.strip() != ""]

        df = df[required]

        df.columns = [
            "Settlement Date",
            "UTR Number",
            "Payment Amount"
        ]

        return df

    # -----------------------------------------------------

    def extract_prepaid(self):

        required = [
            "settlement_date_prepaid_payment",
            "UTR_Number_Prepaid",
            "Settled_Amount_Prepaid"
        ]

        df = self.df.copy()

        df = df[df["UTR_Number_Prepaid"].notna()]

        df = df[df["UTR_Number_Prepaid"].astype(str).str.strip() != ""]

        df = df[required]

        df.columns = [
            "Settlement Date",
            "UTR Number",
            "Payment Amount"
        ]

        return df

    # -----------------------------------------------------

    def create_payment_register(self):

        self.load_excel()

        postpaid = self.extract_postpaid()

        prepaid = self.extract_prepaid()

        payment_df = pd.concat(
            [postpaid, prepaid],
            ignore_index=True
        )

        payment_df["Payment Amount"] = pd.to_numeric(
            payment_df["Payment Amount"],
            errors="coerce"
        ).fillna(0)

        payment_df["Settlement Date"] = pd.to_datetime(
            payment_df["Settlement Date"],
            errors="coerce"
        )

        payment_register = (

            payment_df

            .groupby(
                [
                    "Settlement Date",
                    "UTR Number"
                ],
                as_index=False
            )["Payment Amount"]

            .sum()

            .sort_values(
                [
                    "Settlement Date",
                    "UTR Number"
                ]
            )

        )

        return payment_register

    # -----------------------------------------------------

    def save_excel(self, output_file):

    register = self.create_payment_register()

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl"
    ) as writer:

        register.to_excel(
            writer,
            sheet_name="Payment Register",
            index=False
        )

    wb = load_workbook(output_file)

    ws = wb["Payment Register"]

    format_payment_register(ws)

    wb.save(output_file)

    return output_file


# ---------------------------------------------------------

if __name__ == "__main__":

    processor = PaymentProcessor(
        "sample.xlsx"
    )

    processor.save_excel(
        "Payment_Register.xlsx"
    )

    print("Completed")
