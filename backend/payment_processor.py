"""
payment_processor.py

Purpose
-------
Reads the Myntra Revamped Payment Report and generates a Payment Register.

The real report ships as a workbook with three sheets:
    - Glossary          (column definitions only, not data)
    - forward_settled   (order settlements)
    - reverse_settled   (return/refund settlements, same columns,
                          typically negative amounts)

Both settlement sheets share the same column layout, with the real
header row sitting on the 3rd row (index 2) — the two rows above it
are merged section labels (e.g. "Payment Details", "Postpaid"/"Prepaid")
and are not real column names.

Myntra also uses a literal two-character placeholder string of two
double-quotes for "not yet settled" cells (UTR/date columns), not a
truly blank cell — this is filtered out explicitly below.

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

REQUIRED_SHEETS = ["forward_settled", "reverse_settled"]
HEADER_ROW = 2  # 0-indexed — real column names sit on the 3rd row


class PaymentProcessor:

    # Myntra's "not yet settled" placeholder — a literal string of
    # two double-quotes, not an actually-empty cell.
    PLACEHOLDER_VALUES = {"", '""'}

    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None

    # -----------------------------------------------------
    # Load Excel
    # -----------------------------------------------------

    def load_excel(self):

        print("Step A - Loading Excel")

        xl = pd.ExcelFile(self.excel_file, engine="openpyxl")

        missing_sheets = [s for s in REQUIRED_SHEETS if s not in xl.sheet_names]
        if missing_sheets:
            raise Exception(
                f"Missing required sheet(s): {', '.join(missing_sheets)}. "
                f"Sheets found in file: {', '.join(xl.sheet_names)}"
            )

        sheet_frames = []
        for sheet_name in REQUIRED_SHEETS:
            sheet_df = xl.parse(sheet_name, header=HEADER_ROW)
            sheet_df.columns = (
                sheet_df.columns.astype(str)
                .str.strip()
                .str.replace("\n", "", regex=False)
                .str.replace("\r", "", regex=False)
            )
            print(f"  Loaded '{sheet_name}': {len(sheet_df)} rows")
            sheet_frames.append(sheet_df)

        self.df = pd.concat(sheet_frames, ignore_index=True)

        print(f"Loaded {len(self.df)} total rows across {REQUIRED_SHEETS}")
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
    # Placeholder-aware UTR filter
    # -----------------------------------------------------

    def _is_real_utr(self, series):
        cleaned = series.astype(str).str.strip()
        return series.notna() & (~cleaned.isin(self.PLACEHOLDER_VALUES))

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

        df = df[self._is_real_utr(df["UTR_Number_Postpaid"])]

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

        df = df[self._is_real_utr(df["UTR_Number_Prepaid"])]

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
        # Coerce Payment Amount — rebuild via .assign() rather than
        # in-place .loc mutation, which raises a dtype error when the
        # source column infers as pandas' newer "str" dtype. Log row
        # counts instead of silently zeroing bad values.
        # -----------------------------------------------------
        logger.info("CPR-5")
        numeric_amount = pd.to_numeric(
            payment_df["Payment Amount"], errors="coerce"
        )

        bad_amounts = int(numeric_amount.isna().sum())
        if bad_amounts:
            logger.warning(
                f"{bad_amounts} row(s) had a non-numeric Payment Amount "
                f"and were coerced to 0 — check source file for bad values."
            )

        payment_df = payment_df.assign(
            **{"Payment Amount": numeric_amount.fillna(0)}
        )

        # -----------------------------------------------------
        # Coerce Settlement Date — same .assign() approach, plus
        # normalize to strip time-of-day so same-day settlements
        # always group together.
        # -----------------------------------------------------
        logger.info("CPR-6")
        parsed_date = pd.to_datetime(
            payment_df["Settlement Date"], errors="coerce"
        ).dt.normalize()

        bad_dates = int(parsed_date.isna().sum())
        if bad_dates:
            logger.warning(
                f"{bad_dates} row(s) had an unparseable Settlement Date "
                f"and were set to NaT — check source file for bad values."
            )

        payment_df = payment_df.assign(
            **{"Settlement Date": parsed_date}
        )

        # -----------------------------------------------------
        # groupby() silently drops rows whose group key is NaT
        # (pandas default dropna=True). Fine for genuine placeholder
        # rows (amount 0), but a nonzero amount with no valid date
        # would otherwise vanish from the Grand Total with zero trace.
        # -----------------------------------------------------
        orphaned = payment_df[
            payment_df["Settlement Date"].isna()
            & (payment_df["Payment Amount"] != 0)
        ]
        if len(orphaned):
            logger.warning(
                f"{len(orphaned)} row(s) have a nonzero Payment Amount but "
                f"no valid Settlement Date — these would be SILENTLY "
                f"DROPPED from the register. Amount at risk: "
                f"{orphaned['Payment Amount'].sum()}"
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