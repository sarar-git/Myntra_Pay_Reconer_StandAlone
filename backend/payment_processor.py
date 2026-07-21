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

        logger.info("CPR-1")
        self.load_excel()

        logger.info("CPR-2")
        postpaid = self.extract_postpaid()

        logger.info("CPR-3")
        prepaid = self.extract_prepaid()

        logger.info("CPR-4")
        payment_df = pd.concat(
            [postpaid, prepaid],
            ignore_index=True
        ).copy()

        # -----------------------------------------------------
        # Coerce Payment Amount — log how many rows are affected
        # instead of silently zeroing bad values
        # -----------------------------------------------------
        logger.info("CPR-5")
        payment_df["Payment Amount"] = pd.to_numeric(
            payment_df["Payment Amount"],
            errors="coerce"
        )

        bad_amounts = int(payment_df["Payment Amount"].isna().sum())
        if bad_amounts:
            logger.warning(
                f"{bad_amounts} row(s) had a non-numeric Payment Amount "
                f"and were coerced to 0 — check source file for bad values."
            )

        payment_df["Payment Amount"] = payment_df["Payment Amount"].fillna(0)

        # -----------------------------------------------------
        # Coerce Settlement Date — normalize to strip time-of-day
        # so same-day settlements always group together, and log
        # any rows that fail to parse instead of silently NaT-ing
        # -----------------------------------------------------
        logger.info("CPR-6")
        payment_df["Settlement Date"] = pd.to_datetime(
            payment_df["Settlement Date"],
            errors="coerce"
        ).dt.normalize()

        bad_dates = int(payment_df["Settlement Date"].isna().sum())
        if bad_dates:
            logger.warning(
                f"{bad_dates} row(s) had an unparseable Settlement Date "
                f"and were set to NaT — check source file for bad values."
            )

        logger.info("CPR-7")
        payment_register = (
            payment_df.groupby(
                ["Settlement Date", "UTR Number"],
                as_index=False
            )["Payment Amount"]
            .sum()
        )

        logger.info("CPR-8")

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