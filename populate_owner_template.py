from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd
import win32com.client


OUTPUT_FILE = "owner_template_export.xlsx"
SOURCE_FILE = "6-28-26 Inventory Report - Environment vSphere cleaned_20260630_111059.xlsx"

TEMPLATE_COLUMNS = [
    "COMPUTER_NAME",
    "DOMAIN_NAME",
    "CUSTOMER_NAME",
    "OWNER",
    "LOCATION",
    "SEARCH_TAG",
    "NOTES",
    "PRODUCT_NUMBER",
    "SHIPPING_DATE",
    "WARRANTY_EXPIRY_DATE",
    "OWNER_EMAIL_ID",
]

SOURCE_HEADER_CANDIDATES = {
    "computer": ["Name", "Computer Name", "Server Name"],
    "owner": ["Asset custodian", "Asset Custodian", "Owner"],
    "criticality": ["Criticality of the VM"],
    "department": ["Department / Sections", "Department / sections"],
    "asset_owner": ["Asset Owner"],
    "combined_notes": ["Department / Sections + Asset Owner", "Department / sections + Asset Owner"],
    "email": ["Asset custodian email", "Asset Custodian Email", "Owner Email", "OWNER_EMAIL_ID", "Email"],
}


def normalize_header(value: object) -> str:
    return str(value).replace("\ufeff", "").strip().strip('"').lower()


def clean_header(value: object) -> str:
    return str(value).replace("\ufeff", "").strip().strip('"')


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\ufeff", "").strip().strip("'\"")


def find_column(dataframe: pd.DataFrame, candidates: list[str]) -> str:
    exact_map = {clean_header(column): column for column in dataframe.columns}
    normalized_map = {normalize_header(column): column for column in dataframe.columns}
    for candidate in candidates:
        exact_column = exact_map.get(clean_header(candidate))
        if exact_column:
            return exact_column

        column = normalized_map.get(normalize_header(candidate))
        if column:
            return column
    raise KeyError(f"Could not find any of these columns: {', '.join(candidates)}")


def resolve_email(owner_name: str, namespace, cache: dict[str, str]) -> str:
    normalized_name = clean_text(owner_name)
    if not normalized_name:
        return ""

    if normalized_name in cache:
        return cache[normalized_name]

    recipient = namespace.CreateRecipient(normalized_name)
    recipient.Resolve()

    email = "Not Found"
    if recipient.Resolved:
        try:
            exch_user = recipient.AddressEntry.GetExchangeUser()
            if exch_user:
                email = exch_user.PrimarySmtpAddress
            else:
                email = recipient.Address
        except Exception:
            email = recipient.Address

    cache[normalized_name] = email
    return email


def score_sheet_headers(headers: set[str]) -> int:
    score = 0
    required_keys = ["name", "asset custodian", "criticality of the vm"]
    for key in required_keys:
        if key in headers:
            score += 2

    if any(header in headers for header in ["department / sections", "department / sections + asset Owner"]):
        score += 2

    if any(header in headers for header in ["asset Owner", "owner email", "asset custodian email"]):
        score += 1

    return score


def build_notes(row: pd.Series, combined_column: str | None, department_column: str | None, asset_owner_column: str | None) -> str:
    department_value = clean_text(row.get(department_column)) if department_column else ""
    asset_owner_value = clean_text(row.get(asset_owner_column)) if asset_owner_column else ""

    if department_value or asset_owner_value:
        return f"Department / Sections: {department_value}. Manager: {asset_owner_value}"

    if combined_column and clean_text(row.get(combined_column)):
        return clean_text(row.get(combined_column))

    return ""


def main() -> None:
    source_file = Path(SOURCE_FILE)
    source_frame = pd.read_excel(source_file)

    computer_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["computer"])
    owner_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["owner"])
    criticality_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["criticality"])

    combined_notes_column = None
    department_column = None
    asset_owner_column = None

    try:
        combined_notes_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["combined_notes"])
    except KeyError:
        combined_notes_column = None

    try:
        department_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["department"])
    except KeyError:
        department_column = None

    try:
        asset_owner_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["asset_owner"])
    except KeyError:
        asset_owner_column = None

    try:
        email_column = find_column(source_frame, SOURCE_HEADER_CANDIDATES["email"])
    except KeyError:
        email_column = None

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    email_cache: dict[str, str] = {}

    output_rows: list[dict[str, object]] = []
    for _, row in source_frame.iterrows():
        owner_name = clean_text(row.get(owner_column))
        email_value = clean_text(row.get(email_column)) if email_column else ""
        if not email_value:
            email_value = resolve_email(owner_name, namespace, email_cache)

        output_rows.append(
            {
                "COMPUTER_NAME": clean_text(row.get(computer_column)),
                "DOMAIN_NAME": "",
                "CUSTOMER_NAME": "",
                "OWNER": owner_name,
                "LOCATION": "",
                "SEARCH_TAG": clean_text(row.get(criticality_column)),
                "NOTES": build_notes(row, None, department_column, asset_owner_column),
                "PRODUCT_NUMBER": "",
                "SHIPPING_DATE": "",
                "WARRANTY_EXPIRY_DATE": "",
                "OWNER_EMAIL_ID": email_value,
            }
        )

    export_frame = pd.DataFrame(output_rows, columns=TEMPLATE_COLUMNS)
    export_frame.to_excel(OUTPUT_FILE, index=False)
    print(f"Source workbook: {source_file}")
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()