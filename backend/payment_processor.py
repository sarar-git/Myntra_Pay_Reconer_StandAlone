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
double-quotes for "not yet settled" / "not applicable" cells (UTR,
date, and some fee columns), not a truly blank cell — this is
filtered/coerced around explicitly below.

Reading uses the "calamine" engine (python-calamine) instead of
openpyxl — measured 6-31x faster on real reports. Writing still uses
openpyxl (calamine is read-only).

Output sheets
--------------
- Payment Register: Settlement Date / UTR Number / Payment Amount
- Combined Data: every row from both source sheets, with every
  "<X>_Postpaid" / "<X>_Prepaid" column pair combined into a single
  suffix-free "<X>" column, arranged in a logical reading order
  (identifiers -> pricing -> settlement -> fees -> tax -> logistics
  -> lifecycle dates). Numeric pairs are summed (placeholder treated
  as 0); identifier pairs (UTR_Number) are coalesced without ever
  silently dropping a value.
- Summary: Forward vs Reverse breakdown (Qty split into Forward_Qty /
  Reverse_Qty [negated] / net Qty, plus Total_Settled_Amount), Fee &
  Commission breakdown, Tax breakdown, and an Unlinked Settlements
  flag (real settled money with no UTR on that side — money the
  Register can't include since it groups by UTR).
- SKU Summary: Forward_Qty / Reverse_Qty [negated] / Qty (net),
  Total_Settled_Amount, and Average_Settled_Amount per SKU.

Note: there is no dedicated quantity/units column in the source
report, so "Qty" here means count of settlement rows (order line
items), not units sold. Reverse (return) rows are negated so net Qty
reflects returns reducing the effective settled quantity.
"""

from openpyxl import load_workbook
from utils import format_payment_register

import pandas as pd
import logging
logger = logging.getLogger(__name__)

REQUIRED_SHEETS = ["forward_settled", "reverse_settled"]
HEADER_ROW = 2  # 0-indexed — real column names sit on the 3rd row

# Logical column grouping for the Combined Data sheet. Any column not
# listed here (e.g. a new field Myntra adds in a future report) is
# appended at the end automatically, in its original order, rather
# than silently dropped — see get_combined_data().
COLUMN_ORDER_GROUPS = [
    # Identifiers
    ["Source Sheet", "order_date", "order_release_id", "seller_order_id",
     "sku_id", "style_id", "hsn_code", "brand", "gender",
     "article_type", "article_level"],
    # Order & pricing
    ["MRP", "Seller_Discount", "Platform_Discount", "Customer_Paid_Amount",
     "Discount_ID", "Seller_Tier"],
    # Settlement
    ["UTR_Number", "settlement_date_postpaid_payment",
     "settlement_date_prepaid_payment", "Expected_Settled_Amount",
     "Settled_Amount", "Amount_pending_settlement", "Seller_Product_Amount"],
    # Fees & commission
    ["Commission_percent_incl_GST", "Commission_Amount_incl_GST",
     "Fixed_Fee", "Logistics_Cost_Forward_incl_Tax", "Shipped_Fee",
     "Pick_and_Pack_Fees", "Tech_Enablement_Charges", "Royalty_percent",
     "Royalty_charges", "Forward_additional_charges", "SJIT_Incentive",
     "Commission_Discount", "Logistics_Cost_Reverse_incl_Tax",
     "Reverse_additional_charges"],
    # Tax
    ["product_tax_category", "total_tax_rate", "taxable_amount",
     "igst_amount", "cgst_amount", "sgst_amount", "tcs_amount",
     "tds_amount", "seller_GST"],
    # Logistics & shipment
    ["Warehouse_ID", "po_type", "invoice_number", "packet_id",
     "tracking_no_fwd", "tracking_no_rvs", "shipment_zone_classification",
     "delivery_pin_code", "shipping_city", "shipping_state",
     "shipping_pin_code"],
    # Lifecycle dates
    ["packing_date", "delivery_date", "return_date", "return_type"],
]

FEE_COLUMNS = [
    "Commission_Amount_incl_GST", "Fixed_Fee",
    "Logistics_Cost_Forward_incl_Tax", "Shipped_Fee", "Pick_and_Pack_Fees",
    "Tech_Enablement_Charges", "Royalty_charges",
    "Forward_additional_charges", "SJIT_Incentive", "Commission_Discount",
    "Logistics_Cost_Reverse_incl_Tax", "Reverse_additional_charges",
]

TAX_COLUMNS = [
    "tcs_amount", "tds_amount", "taxable_amount",
    "igst_amount", "cgst_amount", "sgst_amount",
]


class PaymentProcessor:

    # Myntra's "not yet settled" / "not applicable" placeholder — a
    # literal string of two double-quotes, not an actually-empty cell.
    PLACEHOLDER_VALUES = {"", '""'}

    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.df = None

    # -----------------------------------------------------
    # Load Excel
    # -----------------------------------------------------

    def load_excel(self):

        print("Step A - Loading Excel")

        xl = pd.ExcelFile(self.excel_file, engine="calamine")

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
            # Tag each row with its origin sheet before concatenation,
            # so the Combined Data sheet can show which settlement
            # type (order vs. return/refund) each row came from.
            sheet_df.insert(0, "Source Sheet", sheet_name)
            print(f"  Loaded '{sheet_name}': {len(sheet_df)} rows")
            sheet_frames.append(sheet_df)

        self.df = pd.concat(sheet_frames, ignore_index=True)

        print(f"Loaded {len(self.df)} total rows across {REQUIRED_SHEETS}")

        self._combine_prepaid_postpaid()

        print("Columns Found:")
        print(list(self.df.columns))

    # -----------------------------------------------------
    # Combine every <X>_Postpaid / <X>_Prepaid pair into <X>
    # -----------------------------------------------------

    def _combine_prepaid_postpaid(self):

        print("Step A.1 - Combining Postpaid/Prepaid column pairs")

        postpaid_cols = [c for c in self.df.columns if c.endswith("_Postpaid")]

        combined_numeric = []
        combined_identifier = []
        skipped = []

        for pcol in postpaid_cols:
            base = pcol[:-len("_Postpaid")]
            prepaid_col = f"{base}_Prepaid"

            if prepaid_col not in self.df.columns:
                skipped.append(f"{base} (no matching Prepaid column)")
                continue

            if base in self.df.columns:
                skipped.append(f"{base} (a column with this exact name already exists)")
                continue

            # Coerce first, then judge numeric-ness from the coercion
            # result — not the raw dtype. A column can be dtype=object
            # because it mixes real 0s with Myntra's '""' placeholder
            # string, and still be a genuinely numeric column once the
            # placeholder is treated as 0.
            post_numeric = pd.to_numeric(self.df[pcol], errors="coerce")
            pre_numeric = pd.to_numeric(self.df[prepaid_col], errors="coerce")

            # A true identifier column (e.g. UTR_Number) coerces to ALL
            # NaN — no row in it was ever a real number. Combine those
            # as text identifiers instead of dropping them, so a UTR
            # number is never silently lost from the sheet.
            if post_numeric.isna().all() and pre_numeric.isna().all():
                self.df[base] = self._coalesce_identifier_pair(
                    self.df[pcol], self.df[prepaid_col]
                )
                combined_identifier.append(base)
                continue

            self.df[base] = post_numeric.fillna(0) + pre_numeric.fillna(0)
            combined_numeric.append(base)

        print(f"  Combined {len(combined_numeric)} numeric pair(s): {combined_numeric}")
        print(f"  Combined {len(combined_identifier)} identifier pair(s): {combined_identifier}")
        if skipped:
            print(f"  Skipped {len(skipped)}: {skipped}")

    def _coalesce_identifier_pair(self, postpaid_series, prepaid_series):
        """
        Combines a Postpaid/Prepaid identifier pair (e.g. UTR_Number) into
        one column without ever silently dropping a value. If only one
        side is a real (non-placeholder) value, use it. If both sides are
        real and equal, use the single value. If both sides are real and
        DIFFERENT (an order can be split-settled across both payment
        arms), join both values with "; " so neither is lost.
        """
        post_mask = self._is_real_utr(postpaid_series)
        pre_mask = self._is_real_utr(prepaid_series)

        result = []
        for p, q, pm, qm in zip(postpaid_series, prepaid_series, post_mask, pre_mask):
            if pm and qm:
                result.append(p if p == q else f"{p}; {q}")
            elif pm:
                result.append(p)
            elif qm:
                result.append(q)
            else:
                result.append(None)

        return result

    # -----------------------------------------------------
    # Combined Data sheet (combined columns + non-suffixed columns,
    # original _Postpaid/_Prepaid columns dropped, logically ordered)
    # -----------------------------------------------------

    def get_combined_data(self):
        drop_cols = [
            c for c in self.df.columns
            if c.endswith("_Postpaid") or c.endswith("_Prepaid")
        ]
        df = self.df.drop(columns=drop_cols)

        # Apply the logical column order; anything not in a known
        # group (e.g. a future new field) is appended at the end in
        # its original order, so nothing is ever silently dropped.
        ordered = [c for group in COLUMN_ORDER_GROUPS for c in group if c in df.columns]
        remaining = [c for c in df.columns if c not in ordered]
        return df[ordered + remaining]

    # -----------------------------------------------------
    # Summary sheet data (Forward vs Reverse, Fees, Tax,
    # and an Unlinked Settlements flag)
    # -----------------------------------------------------

    def get_summary_data(self):

        sections = {}

        # Forward vs Reverse breakdown — Qty is the count of settlement
        # rows (no dedicated quantity column exists in the source
        # report). Reverse rows are negated so the TOTAL row shows the
        # net Qty (forward minus returns), consistent with how negative
        # amounts already work for reverse_settled.
        fwd_rev = self.df.groupby("Source Sheet").agg(
            Qty=("sku_id", "count"),
            Total_Settled_Amount=("Settled_Amount", "sum"),
        ).reset_index()
        fwd_rev.loc[fwd_rev["Source Sheet"] == "reverse_settled", "Qty"] *= -1
        total_row = pd.DataFrame([{
            "Source Sheet": "TOTAL",
            "Qty": fwd_rev["Qty"].sum(),
            "Total_Settled_Amount": fwd_rev["Total_Settled_Amount"].sum(),
        }])
        sections["Forward vs Reverse Breakdown"] = pd.concat(
            [fwd_rev, total_row], ignore_index=True
        )

        # Fee / commission breakdown
        fee_cols_present = [c for c in FEE_COLUMNS if c in self.df.columns]
        fee_totals = self.df[fee_cols_present].sum().reset_index()
        fee_totals.columns = ["Fee Type", "Total Amount"]
        sections["Fee & Commission Breakdown"] = fee_totals

        # Tax breakdown
        tax_cols_present = [c for c in TAX_COLUMNS if c in self.df.columns]
        tax_totals = self.df[tax_cols_present].sum().reset_index()
        tax_totals.columns = ["Tax Type", "Total Amount"]
        sections["Tax Breakdown"] = tax_totals

        # Unlinked Settlements flag — real settled money with no UTR
        # to attach it to on that side (Register can't include it
        # since it groups by UTR). Surfaced here so it's visible
        # instead of showing up as an unexplained total mismatch.
        sections["Unlinked Settlements (no UTR on that side)"] = (
            self._get_unlinked_settlements_summary()
        )

        return sections

    def _get_unlinked_settlements_summary(self):
        rows = []
        for arm in ["Postpaid", "Prepaid"]:
            utr_col = f"UTR_Number_{arm}"
            amt_col = f"Settled_Amount_{arm}"
            if utr_col not in self.df.columns or amt_col not in self.df.columns:
                continue
            real_mask = self._is_real_utr(self.df[utr_col])
            amt = pd.to_numeric(self.df[amt_col], errors="coerce").fillna(0)
            unlinked_mask = (~real_mask) & (amt != 0)
            rows.append({
                "Settlement Type": arm,
                "Row Count": int(unlinked_mask.sum()),
                "Total Amount": amt[unlinked_mask].sum(),
            })
        return pd.DataFrame(rows)

    # -----------------------------------------------------
    # SKU-level summary (average payout per SKU)
    # -----------------------------------------------------

    def get_sku_summary(self):
        # Qty = count of settlement rows (no dedicated quantity column
        # exists in the source report). Reverse row counts are negated
        # so Qty (net) = Forward_Qty + Reverse_Qty reflects returns
        # reducing the effective settled quantity for that SKU.
        counts = self.df.groupby(["sku_id", "Source Sheet"]).size().unstack(fill_value=0)
        forward_qty = counts.get("forward_settled", 0)
        reverse_qty = -counts.get("reverse_settled", 0)

        amounts = self.df.groupby("sku_id")["Settled_Amount"].agg(["sum", "mean"])

        summary = pd.DataFrame({
            "sku_id": counts.index,
            "Forward_Qty": forward_qty.values if hasattr(forward_qty, "values") else forward_qty,
            "Reverse_Qty": reverse_qty.values if hasattr(reverse_qty, "values") else reverse_qty,
        })
        summary["Qty"] = summary["Forward_Qty"] + summary["Reverse_Qty"]
        summary["Total_Settled_Amount"] = amounts["sum"].values
        summary["Average_Settled_Amount"] = amounts["mean"].values

        return summary.sort_values("sku_id").reset_index(drop=True)

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

        mask = self._is_real_utr(self.df["UTR_Number_Postpaid"])

        df = self.df.loc[mask, required].copy()

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

        mask = self._is_real_utr(self.df["UTR_Number_Prepaid"])

        df = self.df.loc[mask, required].copy()

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
        )

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

        logger.info(
            f"CPR-8 - REGISTER COMPLETE - {len(payment_register)} row(s), "
            f"Grand Total = {payment_register['Payment Amount'].sum():.2f}"
        )

        return payment_register

    # -----------------------------------------------------
    # Save Excel
    # -----------------------------------------------------

    def save_excel(self, output_file):

        logger.info("STEP I")

        register = self.create_payment_register()

        logger.info("STEP I.1 - Register created")

        combined_data = self.get_combined_data()

        logger.info(f"STEP I.1b - Combined Data sheet built ({len(combined_data)} rows, {len(combined_data.columns)} cols)")

        summary_sections = self.get_summary_data()
        sku_summary = self.get_sku_summary()

        logger.info(f"STEP I.1c - Summary ({len(summary_sections)} sections) and SKU Summary ({len(sku_summary)} SKUs) built")

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

            combined_data.to_excel(
                writer,
                sheet_name="Combined Data",
                index=False
            )

            # Summary sheet: each section is a small table stacked
            # vertically, with its title as a row above it and one
            # blank row of separation.
            current_row = 0
            for title, section_df in summary_sections.items():
                title_df = pd.DataFrame([[title]])
                title_df.to_excel(
                    writer, sheet_name="Summary", index=False, header=False,
                    startrow=current_row
                )
                current_row += 1
                section_df.to_excel(
                    writer, sheet_name="Summary", index=False,
                    startrow=current_row
                )
                current_row += len(section_df) + 2  # +1 header row, +1 blank row

            sku_summary.to_excel(
                writer,
                sheet_name="SKU Summary",
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