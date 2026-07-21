"""
payment_processor.py

Purpose
-------
Reads the Myntra Revamped Payment Report and generates a Payment Register.

Output Columns
--------------
Settlement Date
UTR Number
Payment Amount
"""

from openpyxl import load_workbook
from utils import format_payment_register

import pandas as pd
import logging
logger = logging.getLogger(__name__)

class PaymentProcessor:

    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None

    # -----------------------------------------------------
    # Load Excel
    # -----------------------------------------------------

    def load_excel(self):

        print("Step A - Loading Excel")

        self.df = pd.read_excel(
            self.excel_file,
            engine="openpyxl"
        )

        self.df.columns = (
            self.df.columns.astype(str)
            .str.strip()
            .str.replace("\n", "", regex=False)
            .str.replace("\r", "", regex=False)
        )

        print(f"Loaded {len(self.df)} rows")
        print("Columns Found:")
        print(list(self.df.columns))

    # -----------------------------------------------------
    # Validate Columns
    # -----------------------------------------------------

    def validate_columns(self, required):

        missing = [c for c in required if c not in self.df.columns]

        if missing:
            raise Exception(
                f"Missing required columns: {', '.join(missing)}"
            )

    # -----------------------------------------------------
    # Extract Postpaid
    # -----------------------------------------------------

    def extract_postpaid(self):

        print("Step B - Extracting Postpaid")

        required = [
            "settlement_date_postpaid_payment",
            "UTR_Number_Postpaid",
            "Settled_Amount_Postpaid"
        ]

        self.validate_columns(required)

        df = self.df.copy()

        df = df[df["UTR_Number_Postpaid"].notna()]
        df = df[df["UTR_Number_Postpaid"].astype(str).str.strip() != ""]

        df = df[required].copy()

        df.columns = [
            "Settlement Date",
            "UTR Number",
            "Payment Amount"
        ]

        print(f"Postpaid Records : {len(df)}")

        return df

    # -----------------------------------------------------
    # Extract Prepaid
    # -----------------------------------------------------

    def extract_prepaid(self):

        print("Step C - Extracting Prepaid")

        required = [
            "settlement_date_prepaid_payment",
            "UTR_Number_Prepaid",
            "Settled_Amount_Prepaid"
        ]

        self.validate_columns(required)

        df = self.df.copy()

        df = df[df["UTR_Number_Prepaid"].notna()]
        df = df[df["UTR_Number_Prepaid"].astype(str).str.strip() != ""]

        df = df[required].copy()

        df.columns = [
            "Settlement Date",
            "UTR Number",
            "Payment Amount"
        ]

        print(f"Prepaid Records : {len(df)}")

        return df

    # -----------------------------------------------------
    # Create Register
    # -----------------------------------------------------

    def create_payment_register(self):

        print("Step D - Creating Payment Register")

        self.load_excel()

        postpaid = self.extract_postpaid()

        prepaid = self.extract_prepaid()

        print("Step E - Concatenating")

        payment_df = pd.concat(
            [postpaid, prepaid],
            ignore_index=True
        ).copy()

        print("Step F - Converting Amount")

        payment_df.loc[:, "Payment Amount"] = pd.to_numeric(
            payment_df["Payment Amount"],
            errors="coerce"
        ).fillna(0)

        print("Step G - Converting Date")

        payment_df.loc[:, "Settlement Date"] = pd.to_datetime(
            payment_df["Settlement Date"],
            errors="coerce"
        )

        print("Step H - Grouping")

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

        print(f"Payment Register Rows : {len(payment_register)}")

        return payment_register

    # -----------------------------------------------------
    # Save Excel
    # -----------------------------------------------------

    def save_excel(self, output_file):

        logger.info("STEP I")

        register = self.create_payment_register()

        logger.info("STEP I.1 - Register created")

        with pd.ExcelWriter(
            output_file,
            engine="openpyxl"
        ) as writer:

            logger.info("STEP I.2 - ExcelWriter opened")

            register.to_excel(
                writer,
                sheet_name="Payment Register",
                index=False
            )

            logger.info("STEP I.3 - DataFrame written")

        logger.info("STEP J - Workbook written to disk")

        wb = load_workbook(output_file)

        logger.info("STEP K - Workbook loaded")

        ws = wb["Payment Register"]

        logger.info("STEP L - Formatting")

        format_payment_register(ws)

        logger.info("STEP M - Saving workbook")

        wb.save(output_file)

        logger.info("STEP N - Closing workbook")

        wb.close()

        logger.info("STEP O - Completed")

        return output_file

# ---------------------------------------------------------
# Local Testing
# ---------------------------------------------------------

if __name__ == "__main__":

    processor = PaymentProcessor("sample.xlsx")

    processor.save_excel("Payment_Register.xlsx")

    print("Completed")