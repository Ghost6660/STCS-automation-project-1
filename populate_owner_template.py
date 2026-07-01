from __future__ import annotations

import pandas as pd


SOURCE_FILE = "rebootPendingList (3) 1 with owners.xlsx"
OUTPUT_FILE = "owner_template_export.xlsx"

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


def normalize_column_name(value: object) -> str:
    return str(value).replace("\ufeff", "").strip().strip('"').lower()


def find_column(dataframe: pd.DataFrame, candidates: list[str]) -> str:
    normalized_map = {normalize_column_name(column): column for column in dataframe.columns}
    for candidate in candidates:
        column = normalized_map.get(normalize_column_name(candidate))
        if column:
            return column
    raise KeyError(f"Could not find any of these columns: {', '.join(candidates)}")


def clean_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\ufeff", "").strip().strip("'\"")


def main() -> None:
    source_frame = pd.read_excel(SOURCE_FILE)

    computer_column = find_column(source_frame, ["Computer Name", "Name", "Server Name"])
    owner_column = find_column(source_frame, ["Owner"])
    email_column = find_column(source_frame, ["Owner Email", "OWNER_EMAIL_ID", "Email"])

    output_rows = []
    for _, row in source_frame.iterrows():
        output_rows.append(
            {
                "COMPUTER_NAME": clean_text(row[computer_column]),
                "DOMAIN_NAME": "",
                "CUSTOMER_NAME": "",
                "OWNER": clean_text(row[owner_column]),
                "LOCATION": "",
                "SEARCH_TAG": "",
                "NOTES": "",
                "PRODUCT_NUMBER": "",
                "SHIPPING_DATE": "",
                "WARRANTY_EXPIRY_DATE": "",
                "OWNER_EMAIL_ID": clean_text(row[email_column]),
            }
        )

    export_frame = pd.DataFrame(output_rows, columns=TEMPLATE_COLUMNS)
    export_frame.to_excel(OUTPUT_FILE, index=False)
    print(f"Wrote {OUTPUT_FILE}")


if __name__ == "__main__":
    main()