from __future__ import annotations

from datetime import datetime
from pathlib import Path
import csv
import re

import pandas as pd


WORKSPACE_DIR = Path(r"C:\Users\User\Downloads\UNI\summer training\automation project 1")
PENDING_REBOOT_FILE = WORKSPACE_DIR / "rebootPendingList (3) 1.csv"
INVENTORY_FILE = WORKSPACE_DIR / "6-28-26 Inventory Report - Environment vSphere cleaned_20260630_111059.xlsx"
OUTPUT_FILE = WORKSPACE_DIR / "rebootPendingList (3) 1 with owners.xlsx"


EMPTY_VALUES = {"", "--", "none", "nan", "null"}


def normalize_text(value: object) -> str:
    if pd.isna(value):
        return ""
    return str(value).replace("\ufeff", "").strip().strip("'\"")


def normalize_key(value: object) -> str:
    text = normalize_text(value)
    text = text.replace("\ufeff", "")
    return text.lower()


def normalize_identifier(value: object) -> str:
    return re.sub(r"[^a-z0-9]+", "", normalize_key(value))


def is_empty_owner(value: object) -> bool:
    return normalize_key(value) in EMPTY_VALUES


def split_ip_addresses(value: object) -> list[str]:
    text = normalize_text(value)
    if not text:
        return []

    parts = [part.strip() for part in re.split(r"[;,]\s*", text) if part.strip()]
    return parts or [text]


def clean_column_names(dataframe: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in dataframe.columns:
        normalized = normalize_text(column).strip('"')
        if normalized != column:
            rename_map[column] = normalized

    if rename_map:
        dataframe = dataframe.rename(columns=rename_map)

    return dataframe


def load_inventory_lookup() -> dict[tuple[str, str], str]:
    inventory = pd.read_excel(INVENTORY_FILE)
    inventory = clean_column_names(inventory)

    name_column = find_column(inventory, ["Name", '"Name"'])
    ip_column = find_column(inventory, ["IP Address", "IPAddress", "IP"])
    custodian_column = find_column(inventory, ["Asset Custodian", "Custodian"])

    exact_lookup: dict[tuple[str, str], str] = {}
    name_lookup: dict[str, str] = {}
    normalized_name_lookup: dict[str, str] = {}
    ip_lookup: dict[str, str] = {}

    for _, row in inventory.iterrows():
        custodian = normalize_text(row[custodian_column])
        if not custodian:
            continue

        name = normalize_key(row[name_column])
        if name and name not in name_lookup:
            name_lookup[name] = custodian

        normalized_name = normalize_identifier(row[name_column])
        if normalized_name and normalized_name not in normalized_name_lookup:
            normalized_name_lookup[normalized_name] = custodian

        for ip_address in split_ip_addresses(row[ip_column]):
            normalized_ip = normalize_key(ip_address)
            if normalized_ip and normalized_ip not in ip_lookup:
                ip_lookup[normalized_ip] = custodian

            key = (name, normalized_ip)
            if key not in exact_lookup:
                exact_lookup[key] = custodian

    return {
        "exact": exact_lookup,
        "name": name_lookup,
        "normalized_name": normalized_name_lookup,
        "ip": ip_lookup,
    }


def find_column(dataframe: pd.DataFrame, candidates: list[str]) -> str:
    normalized_map = {normalize_key(column): column for column in dataframe.columns}
    for candidate in candidates:
        column = normalized_map.get(normalize_key(candidate))
        if column:
            return column
    raise KeyError(f"Could not find any of these columns: {', '.join(candidates)}")


def fill_owners(dataframe: pd.DataFrame, lookup: dict[str, dict]) -> tuple[pd.DataFrame, int]:
    dataframe = clean_column_names(dataframe)

    name_column = find_column(dataframe, ["Computer Name", "Name", "Server Name"])
    ip_column = find_column(dataframe, ["IP Address", "IPAddress", "IP"])
    owner_column = find_column(dataframe, ["Owner"])

    filled_count = 0

    for index, row in dataframe.iterrows():
        current_owner = row[owner_column]
        if not is_empty_owner(current_owner):
            continue

        name = normalize_key(row[name_column])

        matched_owner = None
        if name:
            for ip_address in split_ip_addresses(row[ip_column]):
                normalized_ip = normalize_key(ip_address)
                matched_owner = lookup["exact"].get((name, normalized_ip))
                if matched_owner:
                    break

        if not matched_owner:
            for ip_address in split_ip_addresses(row[ip_column]):
                matched_owner = lookup["ip"].get(normalize_key(ip_address))
                if matched_owner:
                    break

        if not matched_owner and name:
            matched_owner = lookup["name"].get(name)

        if not matched_owner and name:
            matched_owner = lookup["normalized_name"].get(normalize_identifier(row[name_column]))

        if matched_owner:
            dataframe.at[index, owner_column] = matched_owner
            filled_count += 1

    return dataframe, filled_count


def write_output(dataframe: pd.DataFrame, output_file: Path) -> Path:
    try:
        dataframe.to_excel(output_file, index=False)
        return output_file
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = output_file.with_name(f"{output_file.stem}_{timestamp}{output_file.suffix}")
        dataframe.to_excel(fallback_file, index=False)
        return fallback_file


def main() -> None:
    lookup = load_inventory_lookup()
    pending = pd.read_csv(PENDING_REBOOT_FILE, dtype=str, keep_default_na=False, quoting=csv.QUOTE_MINIMAL)
    updated, filled_count = fill_owners(pending, lookup)
    written_file = write_output(updated, OUTPUT_FILE)

    if written_file == OUTPUT_FILE:
        print(f"Wrote {written_file}")
    else:
        print(f"Target file was locked, wrote fallback file: {written_file}")
    print(f"Filled {filled_count} missing Owner values.")


if __name__ == "__main__":
    main()