from __future__ import annotations

import re
from datetime import datetime
from pathlib import Path

import pandas as pd


SOURCE_FILE = Path(r"C:\Users\User\Downloads\UNI\summer training\automation project 1\6-28-26 Inventory Report - Environment vSphere  - Copy.csv")
OUTPUT_FILE = SOURCE_FILE.with_name("6-28-26 Inventory Report - Environment vSphere cleaned.xlsx")


TAG_PATTERN = re.compile(r"<([^<>]+)>")


def write_output(dataframe: pd.DataFrame, output_file: Path) -> Path:
    try:
        dataframe.to_excel(output_file, index=False)
        return output_file
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = output_file.with_name(f"{output_file.stem}_{timestamp}{output_file.suffix}")
        dataframe.to_excel(fallback_file, index=False)
        return fallback_file


def parse_tags(raw_value: object) -> dict[str, str]:
    if pd.isna(raw_value):
        return {}

    text = str(raw_value).strip()
    if not text or text.lower() == "none":
        return {}

    tags: dict[str, str] = {}
    for entry in TAG_PATTERN.findall(text):
        if "-" not in entry:
            continue

        key, value = entry.split("-", 1)
        key = key.strip()
        value = value.strip()
        if key:
            tags[key] = value

    return tags


def main() -> None:
    dataframe = pd.read_csv(SOURCE_FILE)

    tag_rows = dataframe["vSphere Tag"].apply(parse_tags)
    non_empty_mask = tag_rows.map(bool)
    dataframe = dataframe.loc[non_empty_mask].copy().reset_index(drop=True)
    tag_rows = tag_rows.loc[non_empty_mask].reset_index(drop=True)

    tag_columns = pd.DataFrame(list(tag_rows), index=dataframe.index)
    combined = pd.concat([dataframe.drop(columns=["vSphere Tag"]), tag_columns], axis=1)

    preferred_columns = [
        "Name",
        "Power State",
        "vCPU",
        "Memory (GB)",
        "Disk Space (GB)",
        "vNic",
        "IP Address",
        "Hardware Version",
        "OS",
        "VM Tools Status",
        "Asset Custodian",
        "Asset ID",
        "Asset Owner",
        "Criticality of the VM",
        "Environment",
        "Location",
        "Parent Cluster",
        "Parent vCenter",
    ]

    ordered_columns = [column for column in preferred_columns if column in combined.columns]
    ordered_columns.extend(column for column in combined.columns if column not in ordered_columns)
    combined = combined.loc[:, ordered_columns]

    written_file = write_output(combined, OUTPUT_FILE)
    if written_file == OUTPUT_FILE:
        print(f"Wrote {written_file}")
    else:
        print(f"Target file was locked, wrote fallback file: {written_file}")


if __name__ == "__main__":
    main()