from __future__ import annotations

from datetime import datetime
from pathlib import Path
import re

import pandas as pd


WORKSPACE_DIR = Path(r"C:\Users\User\Downloads\UNI\summer training\automation project 1")
SOURCE_FILE = WORKSPACE_DIR / "6-28-26 Inventory Report - Environment vSphere cleaned20260630_111059.xlsx"
OUTPUT_FILE = WORKSPACE_DIR / "6-28-26 Inventory Report - Environment vSphere grouped by custodian.xlsx"


def write_output(dataframe: pd.DataFrame, output_file: Path) -> Path:
    try:
        dataframe.to_excel(output_file, index=False)
        return output_file
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = output_file.with_name(f"{output_file.stem}_{timestamp}{output_file.suffix}")
        dataframe.to_excel(fallback_file, index=False)
        return fallback_file


def resolve_source_file() -> Path:
    if SOURCE_FILE.exists():
        return SOURCE_FILE

    candidates = [
        path
        for path in WORKSPACE_DIR.glob("6-28-26 Inventory Report - Environment vSphere cleaned*.xlsx")
        if not path.name.startswith("~$")
    ]
    if not candidates:
        raise FileNotFoundError("Could not find a cleaned vSphere workbook to group.")

    return max(candidates, key=lambda path: path.stat().st_mtime)


def normalize_columns(dataframe: pd.DataFrame) -> pd.DataFrame:
    rename_map: dict[str, str] = {}
    for column in dataframe.columns:
        normalized = str(column).strip().replace("\ufeff", "")
        normalized = normalized.strip('"')
        if normalized != column:
            rename_map[column] = normalized

    if rename_map:
        dataframe = dataframe.rename(columns=rename_map)

    return dataframe


def sanitize_sheet_name(name: str) -> str:
    cleaned = re.sub(r"[\\/?*\[\]:]", "_", name).strip()
    if not cleaned:
        cleaned = "Unassigned"
    return cleaned[:31]


def build_custodian_frames(dataframe: pd.DataFrame) -> tuple[pd.DataFrame, dict[str, pd.DataFrame]]:
    dataframe = normalize_columns(dataframe)

    if "Asset Custodian" not in dataframe.columns:
        raise KeyError("The cleaned workbook does not contain an 'Asset Custodian' column.")

    preferred_order = [
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
        "Asset ID",
        "Asset Owner",
        "Criticality of the VM",
        "Environment",
        "Location",
        "Parent Cluster",
        "Parent vCenter",
    ]
    server_columns = [column for column in preferred_order if column in dataframe.columns]
    server_columns.extend(column for column in dataframe.columns if column not in server_columns and column != "Asset Custodian")

    summary_rows: list[dict[str, object]] = []
    custodian_frames: dict[str, pd.DataFrame] = {}

    for custodian, custodian_frame in dataframe.groupby("Asset Custodian", sort=True, dropna=False):
        custodian_name = "Unassigned" if pd.isna(custodian) or str(custodian).strip() == "" else str(custodian).strip()
        custodian_frame = custodian_frame.loc[:, server_columns].copy().reset_index(drop=True)
        custodian_frames[custodian_name] = custodian_frame
        summary_rows.append({"Asset Custodian": custodian_name, "Server Count": int(len(custodian_frame))})

    summary = pd.DataFrame(summary_rows)
    return summary, custodian_frames


def main() -> None:
    source_file = resolve_source_file()
    dataframe = pd.read_excel(source_file)
    summary, custodian_frames = build_custodian_frames(dataframe)

    try:
        with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            for custodian_name, custodian_frame in custodian_frames.items():
                sheet_name = sanitize_sheet_name(custodian_name)
                custodian_frame.to_excel(writer, index=False, sheet_name=sheet_name)
        written_file = OUTPUT_FILE
    except PermissionError:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fallback_file = OUTPUT_FILE.with_name(f"{OUTPUT_FILE.stem}_{timestamp}{OUTPUT_FILE.suffix}")
        with pd.ExcelWriter(fallback_file, engine="openpyxl") as writer:
            summary.to_excel(writer, index=False, sheet_name="Summary")
            for custodian_name, custodian_frame in custodian_frames.items():
                sheet_name = sanitize_sheet_name(custodian_name)
                custodian_frame.to_excel(writer, index=False, sheet_name=sheet_name)
        written_file = fallback_file

    if written_file == OUTPUT_FILE:
        print(f"Wrote {written_file}")
    else:
        print(f"Target file was locked, wrote fallback file: {written_file}")


if __name__ == "__main__":
    main()